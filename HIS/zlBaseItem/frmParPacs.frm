VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmParPacs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ӱ���������"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11250
   Icon            =   "frmParPacs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11250
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picFunc 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   7995
      Left            =   0
      ScaleHeight     =   7995
      ScaleWidth      =   2415
      TabIndex        =   183
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox picTPL 
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   0
         ScaleHeight     =   6135
         ScaleWidth      =   2250
         TabIndex        =   185
         Top             =   0
         Width           =   2250
         Begin XtremeSuiteControls.TaskPanel tplFunc 
            Height          =   5250
            Left            =   0
            TabIndex        =   186
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
            Left            =   1800
            Top             =   360
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
            Icons           =   "frmParPacs.frx":6852
         End
         Begin XtremeSuiteControls.ShortcutCaption sccFunc 
            Height          =   300
            Left            =   0
            TabIndex        =   187
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
      Begin VB.PictureBox picVbar 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         FillColor       =   &H8000000A&
         Height          =   5820
         Left            =   2280
         MousePointer    =   9  'Size W E
         ScaleHeight     =   5820
         ScaleWidth      =   45
         TabIndex        =   184
         Top             =   120
         Width           =   45
      End
      Begin XtremeSuiteControls.ShortcutBar scbFunc 
         Height          =   6765
         Left            =   0
         TabIndex        =   0
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
         Icons           =   "frmParPacs.frx":ADF2
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
      ScaleWidth      =   11250
      TabIndex        =   166
      Top             =   7995
      Width           =   11250
      Begin VB.CommandButton cmdApply 
         Caption         =   "Ӧ��(&A)"
         Height          =   350
         Left            =   10200
         TabIndex        =   309
         Top             =   120
         Width           =   1100
      End
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   1
         Left            =   4700
         TabIndex        =   189
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   0
         Left            =   2400
         TabIndex        =   182
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   60
         TabIndex        =   180
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   9000
         TabIndex        =   179
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   7845
         TabIndex        =   178
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   6000
         TabIndex        =   190
         Top             =   165
         Width           =   2055
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
         TabIndex        =   188
         Top             =   165
         Visible         =   0   'False
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
         TabIndex        =   181
         Top             =   168
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab stabDesign 
      Height          =   7995
      Left            =   2400
      TabIndex        =   191
      TabStop         =   0   'False
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   14102
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   12
      TabHeight       =   520
      TabCaption(0)   =   "Ӱ����������"
      TabPicture(0)   =   "frmParPacs.frx":14466
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "picPar(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Ӱ��ҽ������"
      TabPicture(1)   =   "frmParPacs.frx":14482
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "picPar(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Ӱ��ɼ�����"
      TabPicture(2)   =   "frmParPacs.frx":1449E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picPar(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Ӱ��������"
      TabPicture(3)   =   "frmParPacs.frx":144BA
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picPar(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "����鵵����"
      TabPicture(4)   =   "frmParPacs.frx":144D6
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picPar(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "����軹����"
      TabPicture(5)   =   "frmParPacs.frx":144F2
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "picPar(5)"
      Tab(5).ControlCount=   1
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7575
         Index           =   0
         Left            =   -75000
         ScaleHeight     =   7575
         ScaleWidth      =   8895
         TabIndex        =   197
         Top             =   360
         Width           =   8895
         Begin VB.ComboBox cmbDept 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   120
            Width           =   2055
         End
         Begin VB.TextBox txtLab 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   210
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   198
            Text            =   "Ӱ�����"
            Top             =   165
            Width           =   735
         End
         Begin TabDlg.SSTab stabWorkFlow 
            Height          =   7095
            Left            =   120
            TabIndex        =   199
            Top             =   480
            Width           =   8715
            _ExtentX        =   15372
            _ExtentY        =   12515
            _Version        =   393216
            Style           =   1
            Tabs            =   7
            TabsPerRow      =   7
            TabHeight       =   520
            TabCaption(0)   =   "����������"
            TabPicture(0)   =   "frmParPacs.frx":1450E
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fra(2)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "fra(3)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "fra(4)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "fra(0)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "fra(27)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "chkPreView"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).ControlCount=   6
            TabCaption(1)   =   "ִ�м�����"
            TabPicture(1)   =   "frmParPacs.frx":1452A
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fra(13)"
            Tab(1).Control(1)=   "cmdAdd"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "cmdDel"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "cmdSave"
            Tab(1).Control(4)=   "cmdRestore"
            Tab(1).ControlCount=   5
            TabCaption(2)   =   "�Ǽ�¼������"
            TabPicture(2)   =   "frmParPacs.frx":14546
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "fra(14)"
            Tab(2).Control(1)=   "fra(15)"
            Tab(2).ControlCount=   2
            TabCaption(3)   =   "�����Ŷ�����"
            TabPicture(3)   =   "frmParPacs.frx":14562
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "fra(17)"
            Tab(3).Control(1)=   "fra(16)"
            Tab(3).Control(2)=   "chkUseQueue"
            Tab(3).ControlCount=   3
            TabCaption(4)   =   "����༭������"
            TabPicture(4)   =   "frmParPacs.frx":1457E
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "fra(23)"
            Tab(4).Control(1)=   "fra(22)"
            Tab(4).Control(2)=   "fra(21)"
            Tab(4).Control(3)=   "fra(20)"
            Tab(4).Control(4)=   "fra(19)"
            Tab(4).Control(5)=   "fra(24)"
            Tab(4).Control(6)=   "fra(5)"
            Tab(4).ControlCount=   7
            TabCaption(5)   =   "����б�����"
            TabPicture(5)   =   "frmParPacs.frx":1459A
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "cmdDefault"
            Tab(5).Control(1)=   "fra(28)"
            Tab(5).ControlCount=   2
            TabCaption(6)   =   "��������"
            TabPicture(6)   =   "frmParPacs.frx":145B6
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "fra(1)"
            Tab(6).ControlCount=   1
            Begin VB.CheckBox chkPreView 
               Caption         =   "��������ͼԤ��"
               Height          =   375
               Left            =   840
               TabIndex        =   380
               Top             =   4960
               Width           =   1575
            End
            Begin VB.Frame fra 
               Height          =   1485
               Index           =   5
               Left            =   -74280
               TabIndex        =   355
               Top             =   1080
               Width           =   7215
               Begin VB.Frame fra 
                  Caption         =   "¼��ʱ��"
                  Height          =   1150
                  Index           =   6
                  Left            =   4920
                  TabIndex        =   372
                  Top             =   240
                  Width           =   2055
                  Begin VB.OptionButton optResultInput 
                     Caption         =   "�����ӡǰ"
                     Height          =   240
                     Index           =   2
                     Left            =   210
                     TabIndex        =   375
                     Top             =   810
                     Width           =   1290
                  End
                  Begin VB.OptionButton optResultInput 
                     Caption         =   "���ǩ����"
                     Height          =   240
                     Index           =   1
                     Left            =   210
                     TabIndex        =   374
                     Top             =   525
                     Width           =   1230
                  End
                  Begin VB.OptionButton optResultInput 
                     Caption         =   "���ǩ����"
                     Height          =   240
                     Index           =   0
                     Left            =   210
                     TabIndex        =   373
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   1215
                  End
               End
               Begin VB.TextBox txtImageLevel 
                  Height          =   270
                  Left            =   3690
                  TabIndex        =   364
                  Text            =   "��,��"
                  ToolTipText     =   "��������Ӱ�������ĵǼǣ�����ĸ��ȼ�"
                  Top             =   990
                  Width           =   1035
               End
               Begin VB.TextBox txtReportLevel 
                  Height          =   270
                  Left            =   3690
                  TabIndex        =   363
                  Text            =   "��,��"
                  Top             =   600
                  Width           =   1035
               End
               Begin VB.CheckBox chkImageLevel 
                  Caption         =   "Ӱ�������ȼ�"
                  Height          =   180
                  Left            =   2280
                  TabIndex        =   362
                  Top             =   1035
                  Width           =   1410
               End
               Begin VB.CheckBox chkReportLevel 
                  Caption         =   "���������ȼ�"
                  Height          =   180
                  Left            =   2280
                  TabIndex        =   361
                  Top             =   657
                  Width           =   1410
               End
               Begin VB.CheckBox chkConformDetermine 
                  Caption         =   "��������ж�"
                  Height          =   180
                  Left            =   2280
                  TabIndex        =   360
                  ToolTipText     =   "�������������ܺͲ˵�"
                  Top             =   280
                  Width           =   1455
               End
               Begin VB.Frame fra 
                  Height          =   1125
                  Index           =   9
                  Left            =   120
                  TabIndex        =   356
                  Top             =   270
                  Width           =   2055
                  Begin VB.CheckBox chkDefaultPosi 
                     Caption         =   "��Ͻ��Ĭ������"
                     Height          =   300
                     Left            =   120
                     TabIndex        =   359
                     ToolTipText     =   "����������ѡ�񴰿ڣ�Ĭ��ѡ�����ԡ�"
                     Top             =   300
                     Width           =   1815
                  End
                  Begin VB.CheckBox chkReportAfterResult 
                     Caption         =   "���������Ϊ����"
                     Height          =   180
                     Left            =   120
                     TabIndex        =   358
                     ToolTipText     =   "��д����ʱ��û��¼����ϣ���Ĭ�ϼ�¼Ϊ���ԡ�"
                     Top             =   720
                     Width           =   1740
                  End
                  Begin VB.CheckBox chkIgnorePosi 
                     Caption         =   "���Խ����������"
                     Height          =   180
                     Left            =   120
                     TabIndex        =   357
                     ToolTipText     =   "����¼�ʹ��������ԡ�"
                     Top             =   0
                     Width           =   1800
                  End
               End
            End
            Begin VB.Frame fra 
               Height          =   1695
               Index           =   27
               Left            =   720
               TabIndex        =   354
               Top             =   5040
               Width           =   4215
               Begin VB.OptionButton optClickPreview 
                  Caption         =   "��굥��ʱԤ��ͼ��"
                  Height          =   375
                  Left            =   240
                  TabIndex        =   379
                  Top             =   1200
                  Width           =   1935
               End
               Begin VB.OptionButton optMovePreview 
                  Caption         =   "����ƶ�ʱԤ��ͼ��"
                  Height          =   375
                  Left            =   240
                  TabIndex        =   378
                  Top             =   360
                  Width           =   2055
               End
               Begin VB.TextBox txtDelayTime 
                  Height          =   270
                  Left            =   2880
                  MaxLength       =   2
                  TabIndex        =   377
                  ToolTipText     =   "0��ʾ���Զ��ر�"
                  Top             =   795
                  Width           =   495
               End
               Begin VB.Label lblDelayTime 
                  Caption         =   "�ƶ�Ԥ��ʱ�Զ��ر���ʱʱ��       ��"
                  Height          =   180
                  Left            =   480
                  TabIndex        =   376
                  Top             =   840
                  Width           =   3240
               End
            End
            Begin VB.Frame fra 
               Height          =   6375
               Index           =   1
               Left            =   -74880
               TabIndex        =   314
               Top             =   480
               Width           =   8385
               Begin VB.CheckBox chkCheckMaxNo 
                  Caption         =   "��ȡʵ��������"
                  Height          =   300
                  Left            =   240
                  TabIndex        =   348
                  ToolTipText     =   "��ʵ��������Ϊ����˳���š����ѡ��ǰ׺�������ָ��������������ա����������ݿ������ַ��ͼ��ţ����ֹ����ȡʵ�������롱��"
                  Top             =   5880
                  Width           =   1935
               End
               Begin VB.CheckBox chkChangeNO 
                  Caption         =   "�����ֹ���������"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   347
                  ToolTipText     =   "�������ʵ����Ҫ���ֶ��޸ļ��š�"
                  Top             =   5040
                  Width           =   1935
               End
               Begin VB.CheckBox chkCanOverWrite 
                  Caption         =   "��������ظ�"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   346
                  ToolTipText     =   "����Ǽǲ��˵ļ��ų����ظ�����ѡ�񡰻��߼��ű��ֲ��䡱ʱ����Ҫ��������ظ���"
                  Top             =   5460
                  Width           =   1935
               End
               Begin VB.Frame fra 
                  Caption         =   "����һ����"
                  Height          =   4290
                  Index           =   12
                  Left            =   120
                  TabIndex        =   335
                  Top             =   360
                  Width           =   4000
                  Begin VB.OptionButton OptCode 
                     Caption         =   "ÿ�μ�����¼���"
                     Height          =   180
                     Index           =   0
                     Left            =   120
                     TabIndex        =   345
                     ToolTipText     =   "����ʱ�����µļ��š�"
                     Top             =   360
                     Value           =   -1  'True
                     Width           =   1920
                  End
                  Begin VB.OptionButton OptCode 
                     Caption         =   "���߼��ű��ֲ���"
                     Height          =   180
                     Index           =   1
                     Left            =   120
                     TabIndex        =   344
                     ToolTipText     =   "ͬһ�����ߣ�����ʱ���ּ��Ų��䡣"
                     Top             =   2520
                     Width           =   1935
                  End
                  Begin VB.Frame fra 
                     Height          =   1400
                     Index           =   7
                     Left            =   480
                     TabIndex        =   340
                     Top             =   2760
                     Width           =   3300
                     Begin VB.OptionButton OptUnicode 
                        Caption         =   "��������ͳһ"
                        Height          =   240
                        Index           =   0
                        Left            =   240
                        TabIndex        =   343
                        ToolTipText     =   "��������ͬ�����ּ��Ų��䡣"
                        Top             =   300
                        Value           =   -1  'True
                        Width           =   1590
                     End
                     Begin VB.OptionButton OptUnicode 
                        Caption         =   "������ͳһ"
                        Height          =   210
                        Index           =   1
                        Left            =   240
                        TabIndex        =   342
                        ToolTipText     =   "������ͬ�����ּ��Ų��䡣"
                        Top             =   705
                        Width           =   1290
                     End
                     Begin VB.OptionButton optUsePatientID 
                        Caption         =   "ȫԺͳһ��ʹ�ò���ID��"
                        Height          =   210
                        Left            =   240
                        TabIndex        =   341
                        ToolTipText     =   "ʹ�ò���ID��Ϊ���š�"
                        Top             =   1080
                        Width           =   2370
                     End
                  End
                  Begin VB.Frame Frame2 
                     Height          =   1400
                     Left            =   480
                     TabIndex        =   336
                     Top             =   720
                     Width           =   3300
                     Begin VB.OptionButton OptBuildcode 
                        Caption         =   "��ͬ�������Զ�����"
                        Height          =   210
                        Index           =   0
                        Left            =   240
                        TabIndex        =   339
                        ToolTipText     =   "�����Լ�����Ϊ�������Զ�������"
                        Top             =   300
                        Value           =   -1  'True
                        Width           =   2130
                     End
                     Begin VB.OptionButton OptBuildcode 
                        Caption         =   "���������Զ�����"
                        Height          =   210
                        Index           =   1
                        Left            =   240
                        TabIndex        =   338
                        ToolTipText     =   "�����Կ���Ϊ�������Զ�������"
                        Top             =   690
                        Width           =   1740
                     End
                     Begin VB.OptionButton optUseAdviceID 
                        Caption         =   "ʹ��ҽ��ID"
                        Height          =   210
                        Left            =   240
                        TabIndex        =   337
                        ToolTipText     =   "ʹ��ҽ��ID��Ϊ���š�"
                        Top             =   1080
                        Width           =   1740
                     End
                  End
               End
               Begin VB.Frame Frame3 
                  Caption         =   "���������"
                  Height          =   5895
                  Left            =   4200
                  TabIndex        =   315
                  Top             =   360
                  Width           =   4000
                  Begin VB.ComboBox cboDelimeter 
                     BeginProperty Font 
                        Name            =   "����"
                        Size            =   10.5
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   2
                     Left            =   1280
                     Style           =   2  'Dropdown List
                     TabIndex        =   350
                     Top             =   3694
                     Width           =   2500
                  End
                  Begin VB.ComboBox cboDelimeter 
                     BeginProperty Font 
                        Name            =   "����"
                        Size            =   10.5
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Index           =   1
                     Left            =   1280
                     Style           =   2  'Dropdown List
                     TabIndex        =   349
                     Top             =   1942
                     Width           =   2500
                  End
                  Begin VB.CheckBox chkPreText 
                     Caption         =   "ǰ׺"
                     Height          =   375
                     Left            =   240
                     TabIndex        =   334
                     ToolTipText     =   "�������ӹ̶�ǰ׺��"
                     Top             =   360
                     Width           =   1215
                  End
                  Begin VB.Frame frmPreText 
                     Height          =   1100
                     Left            =   480
                     TabIndex        =   330
                     Top             =   720
                     Width           =   3300
                     Begin VB.OptionButton optPreText 
                        Caption         =   "Ӱ�����"
                        Height          =   255
                        Index           =   0
                        Left            =   240
                        TabIndex        =   333
                        ToolTipText     =   "ʹ�ü���Ӱ�������Ϊǰ׺��"
                        Top             =   240
                        Width           =   1455
                     End
                     Begin VB.OptionButton optPreText 
                        Caption         =   "�����ı�"
                        Height          =   255
                        Index           =   1
                        Left            =   240
                        TabIndex        =   332
                        ToolTipText     =   "ʹ�������ı���Ϊǰ׺��"
                        Top             =   600
                        Value           =   -1  'True
                        Width           =   1215
                     End
                     Begin VB.TextBox txtPreText 
                        Height          =   375
                        Left            =   1440
                        MaxLength       =   10
                        TabIndex        =   331
                        ToolTipText     =   "ǰ׺��������10���ַ�"
                        Top             =   540
                        Width           =   1600
                     End
                  End
                  Begin VB.CheckBox chkDelimiter 
                     Caption         =   "�ָ���1"
                     Height          =   375
                     Index           =   1
                     Left            =   240
                     TabIndex        =   329
                     ToolTipText     =   "ǰ׺֮��ķָ�����"
                     Top             =   1920
                     Width           =   975
                  End
                  Begin VB.CheckBox chkDelimiter 
                     Caption         =   "�ָ���2"
                     Height          =   375
                     Index           =   2
                     Left            =   240
                     TabIndex        =   328
                     ToolTipText     =   "������֮��ķָ�����"
                     Top             =   3672
                     Width           =   975
                  End
                  Begin VB.CheckBox chkYear 
                     Caption         =   "��"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   327
                     ToolTipText     =   "�ڼ���֮ǰ���ӵ�ǰ�ꡣ"
                     Top             =   2448
                     Width           =   735
                  End
                  Begin VB.Frame frmYear 
                     Height          =   500
                     Left            =   1280
                     TabIndex        =   324
                     Top             =   2325
                     Width           =   2500
                     Begin VB.OptionButton optYear 
                        Caption         =   "��λ"
                        Height          =   350
                        Index           =   0
                        Left            =   240
                        TabIndex        =   326
                        ToolTipText     =   "��λ��ݣ����硰2008����"
                        Top             =   120
                        Value           =   -1  'True
                        Width           =   735
                     End
                     Begin VB.OptionButton optYear 
                        Caption         =   "��λ"
                        Height          =   350
                        Index           =   1
                        Left            =   1320
                        TabIndex        =   325
                        ToolTipText     =   "��λ��ݣ����硰08����"
                        Top             =   120
                        Width           =   735
                     End
                  End
                  Begin VB.CheckBox chkMonth 
                     Caption         =   "��"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   323
                     ToolTipText     =   "�ڼ���֮ǰ���ӵ�ǰ�¡�"
                     Top             =   2856
                     Width           =   735
                  End
                  Begin VB.CheckBox chkDay 
                     Caption         =   "��"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   322
                     ToolTipText     =   "�ڼ���֮ǰ���ӵ�ǰ�ա�"
                     Top             =   3264
                     Width           =   615
                  End
                  Begin VB.CheckBox chkNumber 
                     Caption         =   "˳���"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   321
                     ToolTipText     =   "˳�����Ĭ�ϱ���Ҫѡ���"
                     Top             =   4200
                     Value           =   1  'Checked
                     Width           =   975
                  End
                  Begin VB.Frame Frame6 
                     Height          =   1335
                     Left            =   480
                     TabIndex        =   316
                     Top             =   4440
                     Width           =   3300
                     Begin VB.TextBox txtStartNum 
                        Height          =   375
                        Left            =   1440
                        MaxLength       =   4
                        TabIndex        =   319
                        ToolTipText     =   "���ű�ŵ���ʼ���룬С��4λ��"
                        Top             =   300
                        Width           =   1600
                     End
                     Begin VB.CheckBox chkFixedLen 
                        Caption         =   "�̶�λ��"
                        Height          =   255
                        Left            =   240
                        TabIndex        =   318
                        ToolTipText     =   "���Ű��չ̶�λ����ţ�ǰ�油�㡣"
                        Top             =   840
                        Width           =   1095
                     End
                     Begin VB.TextBox txtFixedLen 
                        Height          =   375
                        Left            =   1440
                        MaxLength       =   2
                        TabIndex        =   317
                        ToolTipText     =   "�̶�λ��С��18λ"
                        Top             =   780
                        Width           =   1600
                     End
                     Begin VB.Label lblStartNum 
                        Caption         =   "��ʼ����"
                        Height          =   255
                        Left            =   240
                        TabIndex        =   320
                        ToolTipText     =   "���ű�ŵ���ʼ���롣"
                        Top             =   360
                        Width           =   975
                     End
                  End
               End
            End
            Begin VB.Frame fra 
               Height          =   2715
               Index           =   0
               Left            =   720
               TabIndex        =   200
               Top             =   480
               Width           =   7335
               Begin VB.CheckBox chkSetFocusWithReport 
                  Caption         =   "����л�ʱ��λ����༭"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   301
                  ToolTipText     =   "�л�������ҳ��ʱ�Ƿ�λ����༭"
                  Top             =   2400
                  Width           =   2415
               End
               Begin VB.CheckBox chkCompletePrint 
                  Caption         =   "�����ֱ�Ӵ�ӡ"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   297
                  ToolTipText     =   "����ǩ����ֱ�Ӵ�ӡ���档"
                  Top             =   2100
                  Width           =   1680
               End
               Begin VB.CheckBox chkFinallyCompleteCommit 
                  Caption         =   "�����ֱ�����"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   296
                  ToolTipText     =   "������˺󣬸ü���Զ���ɣ��������ڱ����ĵ��༭����"
                  Top             =   1173
                  Width           =   1815
               End
               Begin VB.Frame Frame11 
                  Caption         =   "ҽ��վ�鿴����"
                  Height          =   615
                  Left            =   4680
                  TabIndex        =   294
                  ToolTipText     =   "�������ڱ����ĵ��༭����"
                  Top             =   1800
                  Width           =   2415
                  Begin VB.ComboBox cboViewReport 
                     Height          =   300
                     ItemData        =   "frmParPacs.frx":145D2
                     Left            =   240
                     List            =   "frmParPacs.frx":145DC
                     Style           =   2  'Dropdown List
                     TabIndex        =   295
                     ToolTipText     =   "�������ڱ����ĵ��༭����"
                     Top             =   240
                     Width           =   1935
                  End
               End
               Begin VB.CheckBox chkAddons 
                  Caption         =   "��ʾ��������"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   289
                  ToolTipText     =   "�ڵǼǱ���������ʾ��������һ��"
                  Top             =   2360
                  Width           =   1935
               End
               Begin VB.CheckBox chkReagent 
                  Caption         =   "��ʾ��Ӱ��"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   288
                  ToolTipText     =   "�ڵǼǱ���������ʾ��Ӱ��һ�������վ����ʾ"
                  Top             =   2040
                  Width           =   1935
               End
               Begin VB.TextBox txtRefreshInterval 
                  Enabled         =   0   'False
                  Height          =   270
                  Left            =   6480
                  TabIndex        =   12
                  Text            =   "0"
                  Top             =   778
                  Width           =   390
               End
               Begin VB.TextBox TxtLike 
                  Enabled         =   0   'False
                  Height          =   270
                  Left            =   6600
                  MaxLength       =   2
                  TabIndex        =   8
                  ToolTipText     =   "0������ʱ������,ģ���������в���"
                  Top             =   470
                  Width           =   270
               End
               Begin VB.CheckBox chkAutoSendWorkList 
                  Caption         =   "����ʱ�Զ�����WorkList"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   19
                  Top             =   1740
                  Value           =   1  'Checked
                  Width           =   2532
               End
               Begin VB.TextBox txtViewHistoryImageDays 
                  Height          =   270
                  Left            =   6600
                  MaxLength       =   2
                  TabIndex        =   20
                  Text            =   "1"
                  Top             =   1395
                  Width           =   465
               End
               Begin VB.CheckBox chkCanViewImage 
                  Caption         =   "��ͼ��ҽ��վ���ɹ�Ƭ"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   16
                  ToolTipText     =   "�ɼ�ͼ�����û�м����ɵ�����£�ҽ��վҲ�ɽ��й�Ƭ��"
                  Top             =   1443
                  Width           =   2160
               End
               Begin VB.CheckBox ChkFinishCommit 
                  Caption         =   "�ޱ�����ɺ�ֱ�����"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   14
                  ToolTipText     =   "����ޱ�����ɺ󣬸ü���Զ���ɡ�"
                  Top             =   1146
                  Width           =   2160
               End
               Begin VB.CheckBox chkPrintCommit 
                  Caption         =   "��ӡ��ֱ�����"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   5
                  ToolTipText     =   "��ӡ����󣬸ü���Զ���ɡ�"
                  Top             =   561
                  Width           =   1815
               End
               Begin VB.CheckBox ChkCompleteCommit 
                  Caption         =   "��˺�ֱ�����"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   9
                  ToolTipText     =   "������˺󣬸ü���Զ���ɡ�"
                  Top             =   867
                  Width           =   1935
               End
               Begin VB.CheckBox chkSample 
                  Caption         =   "����ǼǺ�ֱ�ӱ���"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   15
                  ToolTipText     =   "�Ǽ��뱨��ͬʱ���С�"
                  Top             =   1785
                  Width           =   1935
               End
               Begin VB.TextBox TxtĬ������ 
                  Height          =   270
                  Left            =   6120
                  MaxLength       =   2
                  TabIndex        =   18
                  Text            =   "2"
                  Top             =   1086
                  Width           =   945
               End
               Begin VB.CheckBox chkReportAfterImging 
                  Caption         =   "��ͼ�����д����"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   2
                  ToolTipText     =   "����ɼ�ͼ�����ܱ�дӰ�񱨸档"
                  Top             =   255
                  Width           =   2040
               End
               Begin VB.CheckBox chkPrintNeedComplete 
                  Caption         =   "ƽ��������˲��ܴ򱨸�"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   10
                  ToolTipText     =   "ƽ������뾭����˺���ܴ�ӡ���档"
                  Top             =   849
                  Width           =   2505
               End
               Begin VB.CheckBox chkTechReportSame 
                  Caption         =   "ֻ����д�Լ����ı���"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   6
                  ToolTipText     =   "ֻ���Լ��ɼ�ͼ��ļ�飬������д���档"
                  Top             =   552
                  Width           =   2295
               End
               Begin VB.CheckBox chkWriteCapDoctor 
                  Caption         =   "�ɼ�ͼ����Ϊ��鼼ʦ"
                  Height          =   180
                  Left            =   4680
                  TabIndex        =   4
                  ToolTipText     =   "�ɼ�ͼ��֮���Զ�����ǰ�û���¼�ɼ�鼼ʦ��"
                  Top             =   240
                  Width           =   2400
               End
               Begin VB.CheckBox chkLocalizerBackward 
                  Caption         =   "��λƬ����"
                  Height          =   180
                  Left            =   120
                  TabIndex        =   13
                  ToolTipText     =   "����λƬ�ŵ����һ��������ʾ��"
                  Top             =   1479
                  Width           =   1320
               End
               Begin VB.CheckBox chkRefreshInterval 
                  Caption         =   "�����Զ�ˢ�¼��      ��"
                  Height          =   180
                  Left            =   4680
                  TabIndex        =   11
                  ToolTipText     =   "���˼���б����N���Զ�ˢ�¡�"
                  Top             =   847
                  Width           =   2500
               End
               Begin VB.CheckBox chkAllPatientIsOutside 
                  Caption         =   "���еǼǲ��˱��Ϊ����"
                  Height          =   180
                  Left            =   2160
                  TabIndex        =   3
                  ToolTipText     =   "���ڸù���վ�еǼǵĲ��˾����Ϊ�������ˡ�"
                  Top             =   255
                  Width           =   2295
               End
               Begin VB.CheckBox ChkLike 
                  Caption         =   "�Ǽ�ʱ����ģ������    ��"
                  Height          =   195
                  Left            =   4680
                  TabIndex        =   7
                  ToolTipText     =   "�Ǽ�ʱ֧�ֶ���������ģ�����ң����Բ��ҵ�N���ڵ���Ϣ��"
                  Top             =   536
                  Width           =   2500
               End
               Begin VB.Label lab 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Զ�����ʷͼ������"
                  Height          =   180
                  Index           =   1
                  Left            =   4680
                  TabIndex        =   201
                  ToolTipText     =   "�����ǰ���û��ͼ�����Զ���ָ��ʱ����ڵ���ʷͼ��"
                  Top             =   1440
                  Width           =   1800
               End
               Begin VB.Label lab 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ĭ�ϼ�¼��ѯ����"
                  Height          =   180
                  Index           =   0
                  Left            =   4680
                  TabIndex        =   17
                  ToolTipText     =   "����б���Ĭ����ʾ��Ӧ�����ڵļ���¼��"
                  Top             =   1143
                  Width           =   1440
               End
            End
            Begin VB.Frame fra 
               Caption         =   "�����ĵ��༭������"
               Height          =   4335
               Index           =   24
               Left            =   -74280
               TabIndex        =   213
               Top             =   2640
               Width           =   7245
               Begin VB.Frame fra 
                  Caption         =   "��ʷ����鿴�༭��"
                  Height          =   615
                  Index           =   25
                  Left            =   240
                  TabIndex        =   214
                  Top             =   360
                  Width           =   6855
                  Begin VB.OptionButton optHistoryReportEditor 
                     Caption         =   "PACS����༭��"
                     Height          =   255
                     Index           =   1
                     Left            =   4080
                     TabIndex        =   216
                     Top             =   240
                     Width           =   1695
                  End
                  Begin VB.OptionButton optHistoryReportEditor 
                     Caption         =   "���Ӳ����༭��"
                     Height          =   255
                     Index           =   0
                     Left            =   360
                     TabIndex        =   215
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   1695
                  End
               End
            End
            Begin VB.CheckBox chkUseQueue 
               Caption         =   "�����Ŷӽк�"
               Height          =   180
               Left            =   -74160
               TabIndex        =   286
               ToolTipText     =   "�����ŶӽкŹ��ܣ�������Ӱ��ɼ�վ��Ӱ��ҽ��վ��"
               Top             =   5400
               Width           =   1455
            End
            Begin VB.Frame fra 
               Height          =   5805
               Index           =   13
               Left            =   -74280
               TabIndex        =   243
               Top             =   480
               Width           =   7305
               Begin VB.TextBox txtNoPrefix 
                  Height          =   300
                  Left            =   6075
                  MaxLength       =   20
                  TabIndex        =   37
                  Top             =   5340
                  Width           =   1050
               End
               Begin VB.ComboBox cboDevice 
                  Height          =   300
                  Left            =   3240
                  Style           =   2  'Dropdown List
                  TabIndex        =   36
                  Top             =   5340
                  Width           =   1830
               End
               Begin VB.TextBox txtName 
                  Height          =   300
                  Left            =   840
                  MaxLength       =   20
                  TabIndex        =   35
                  Top             =   5340
                  Width           =   1635
               End
               Begin MSComctlLib.ImageList img16 
                  Left            =   4320
                  Top             =   600
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
                        Picture         =   "frmParPacs.frx":14600
                        Key             =   "Room"
                     EndProperty
                  EndProperty
               End
               Begin MSComctlLib.ListView lvwRoom 
                  Height          =   4695
                  Left            =   120
                  TabIndex        =   34
                  Top             =   480
                  Width           =   7065
                  _ExtentX        =   12462
                  _ExtentY        =   8281
                  View            =   3
                  Arrange         =   1
                  LabelEdit       =   1
                  Sorted          =   -1  'True
                  LabelWrap       =   -1  'True
                  HideSelection   =   0   'False
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  Icons           =   "img16"
                  SmallIcons      =   "img16"
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  Appearance      =   1
                  NumItems        =   0
               End
               Begin VB.Label lab 
                  AutoSize        =   -1  'True
                  Caption         =   "����ǰ׺"
                  Height          =   180
                  Index           =   6
                  Left            =   5250
                  TabIndex        =   247
                  Top             =   5400
                  Width           =   720
               End
               Begin VB.Label lab 
                  Caption         =   "�豸(&D)"
                  Height          =   180
                  Index           =   5
                  Left            =   2565
                  TabIndex        =   246
                  Top             =   5400
                  Width           =   630
               End
               Begin VB.Label lab 
                  AutoSize        =   -1  'True
                  Caption         =   "����(&N)"
                  Height          =   180
                  Index           =   4
                  Left            =   150
                  TabIndex        =   245
                  Top             =   5400
                  Width           =   630
               End
               Begin VB.Label lab 
                  Caption         =   "���ñ����ҵ�ִ�м�󣬲�����Ч����ִ�еİ��š�"
                  Height          =   210
                  Index           =   3
                  Left            =   150
                  TabIndex        =   244
                  Top             =   210
                  Width           =   4140
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "����(&A)"
               Height          =   345
               Left            =   -71760
               Picture         =   "frmParPacs.frx":14B9A
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   6510
               Width           =   1100
            End
            Begin VB.CommandButton cmdDel 
               Caption         =   "ɾ��(&D)"
               Height          =   345
               Left            =   -70560
               Picture         =   "frmParPacs.frx":14CE4
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   6510
               Width           =   1100
            End
            Begin VB.CommandButton cmdSave 
               Caption         =   "����(&S)"
               Height          =   345
               Left            =   -68160
               TabIndex        =   41
               Top             =   6510
               Width           =   1100
            End
            Begin VB.CommandButton cmdRestore 
               Caption         =   "�ָ�(&R)"
               Height          =   345
               Left            =   -69360
               TabIndex        =   40
               Top             =   6510
               Width           =   1100
            End
            Begin VB.Frame fra 
               Caption         =   "���������Ŀѡ��"
               Height          =   2010
               Index           =   14
               Left            =   -74280
               TabIndex        =   242
               Top             =   480
               Width           =   7300
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "��鼼ʦ��"
                  Height          =   180
                  Index           =   24
                  Left            =   6060
                  TabIndex        =   65
                  Top             =   1320
                  Width           =   1220
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "��������"
                  Height          =   180
                  Index           =   23
                  Left            =   180
                  TabIndex        =   66
                  Top             =   1590
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "��鼼ʦ"
                  Height          =   180
                  Index           =   22
                  Left            =   4800
                  TabIndex        =   64
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "��Ӱ��"
                  Height          =   180
                  Index           =   21
                  Left            =   3525
                  TabIndex        =   63
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "��������"
                  Height          =   180
                  Index           =   20
                  Left            =   3525
                  TabIndex        =   45
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "���ʱ��"
                  Height          =   180
                  Index           =   19
                  Left            =   2415
                  TabIndex        =   62
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "����ʱ��"
                  Height          =   180
                  Index           =   18
                  Left            =   1305
                  TabIndex        =   61
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   15
                  Left            =   4800
                  TabIndex        =   58
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   $"frmParPacs.frx":14E2E
                  Height          =   180
                  Index           =   17
                  Left            =   180
                  TabIndex        =   60
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "����豸"
                  Height          =   180
                  Index           =   16
                  Left            =   6060
                  TabIndex        =   59
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "ִ�м�"
                  Height          =   180
                  Index           =   14
                  Left            =   3525
                  TabIndex        =   57
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "��ַ"
                  Height          =   180
                  Index           =   13
                  Left            =   2415
                  TabIndex        =   56
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "�ʱ�"
                  Height          =   180
                  Index           =   12
                  Left            =   1305
                  TabIndex        =   55
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "�绰"
                  Height          =   180
                  Index           =   11
                  Left            =   180
                  TabIndex        =   54
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   10
                  Left            =   3525
                  TabIndex        =   51
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "ְҵ"
                  Height          =   180
                  Index           =   9
                  Left            =   2415
                  TabIndex        =   50
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   8
                  Left            =   6060
                  TabIndex        =   53
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "���֤��"
                  Height          =   180
                  Index           =   7
                  Left            =   4800
                  TabIndex        =   52
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "���ʽ"
                  Height          =   180
                  Index           =   6
                  Left            =   4800
                  TabIndex        =   46
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "�ѱ�"
                  Height          =   180
                  Index           =   5
                  Left            =   6060
                  TabIndex        =   47
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   4
                  Left            =   1305
                  TabIndex        =   49
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "���"
                  Height          =   180
                  Index           =   3
                  Left            =   180
                  TabIndex        =   48
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   2
                  Left            =   2415
                  TabIndex        =   44
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "�Ա�"
                  Height          =   180
                  Index           =   1
                  Left            =   1305
                  TabIndex        =   43
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkMouseMove 
                  Caption         =   "Ӣ����"
                  Height          =   180
                  Index           =   0
                  Left            =   180
                  TabIndex        =   42
                  Top             =   360
                  Width           =   1020
               End
            End
            Begin VB.Frame fra 
               Caption         =   "�ǼǱ�¼��Ŀѡ��"
               Height          =   2010
               Index           =   15
               Left            =   -74280
               TabIndex        =   241
               Top             =   2760
               Width           =   7300
               Begin VB.CheckBox ChkInput 
                  Caption         =   "Ӣ����"
                  Height          =   180
                  Index           =   0
                  Left            =   180
                  TabIndex        =   67
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "�Ա�"
                  Height          =   180
                  Index           =   1
                  Left            =   1305
                  TabIndex        =   68
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   2
                  Left            =   2415
                  TabIndex        =   69
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "���"
                  Height          =   180
                  Index           =   3
                  Left            =   180
                  TabIndex        =   73
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   4
                  Left            =   1305
                  TabIndex        =   74
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "�ѱ�"
                  Height          =   180
                  Index           =   5
                  Left            =   6060
                  TabIndex        =   72
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "���ʽ"
                  Height          =   180
                  Index           =   6
                  Left            =   4800
                  TabIndex        =   71
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "���֤��"
                  Height          =   180
                  Index           =   7
                  Left            =   4800
                  TabIndex        =   77
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   8
                  Left            =   6060
                  TabIndex        =   78
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "ְҵ"
                  Height          =   180
                  Index           =   9
                  Left            =   2415
                  TabIndex        =   75
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   10
                  Left            =   3525
                  TabIndex        =   76
                  Top             =   690
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "�绰"
                  Height          =   180
                  Index           =   11
                  Left            =   180
                  TabIndex        =   79
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "�ʱ�"
                  Height          =   180
                  Index           =   12
                  Left            =   1305
                  TabIndex        =   80
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "��ַ"
                  Height          =   180
                  Index           =   13
                  Left            =   2415
                  TabIndex        =   81
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "ִ�м�"
                  Height          =   180
                  Index           =   14
                  Left            =   3525
                  TabIndex        =   82
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "����豸"
                  Height          =   180
                  Index           =   16
                  Left            =   6060
                  TabIndex        =   84
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   $"frmParPacs.frx":14E3A
                  Height          =   180
                  Index           =   17
                  Left            =   180
                  TabIndex        =   85
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   15
                  Left            =   4800
                  TabIndex        =   83
                  Top             =   1005
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "����ʱ��"
                  Height          =   180
                  Index           =   18
                  Left            =   1305
                  TabIndex        =   86
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "���ʱ��"
                  Height          =   180
                  Index           =   19
                  Left            =   2415
                  TabIndex        =   87
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "��������"
                  Height          =   180
                  Index           =   20
                  Left            =   3525
                  TabIndex        =   70
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "��Ӱ��"
                  Height          =   180
                  Index           =   21
                  Left            =   3525
                  TabIndex        =   88
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "��鼼ʦ"
                  Height          =   180
                  Index           =   22
                  Left            =   4800
                  TabIndex        =   89
                  Top             =   1320
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "��������"
                  Height          =   180
                  Index           =   23
                  Left            =   180
                  TabIndex        =   91
                  Top             =   1590
                  Width           =   1020
               End
               Begin VB.CheckBox ChkInput 
                  Caption         =   "��鼼ʦ��"
                  Height          =   180
                  Index           =   24
                  Left            =   6060
                  TabIndex        =   90
                  Top             =   1320
                  Width           =   1220
               End
            End
            Begin VB.Frame fra 
               Caption         =   "�б���ɫ����"
               Height          =   5415
               Index           =   28
               Left            =   -74280
               TabIndex        =   222
               Top             =   480
               Width           =   7305
               Begin VB.Frame fra 
                  Caption         =   "��ɫ��ʾ����"
                  Height          =   615
                  Index           =   30
                  Left            =   3960
                  TabIndex        =   224
                  ToolTipText     =   "����б���������ɫ���ͣ�Ϊǰ��ɫʱ�����б��ǰ��ɫ����֮������ɫ��"
                  Top             =   4680
                  Width           =   2055
                  Begin VB.OptionButton optListColorMark 
                     Caption         =   "ǰ��ɫ"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   141
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   855
                  End
                  Begin VB.OptionButton optListColorMark 
                     Caption         =   "����ɫ"
                     Height          =   255
                     Index           =   1
                     Left            =   1080
                     TabIndex        =   142
                     Top             =   240
                     Width           =   855
                  End
               End
               Begin VB.Frame fra 
                  Height          =   615
                  Index           =   29
                  Left            =   720
                  TabIndex        =   223
                  Top             =   4680
                  Width           =   2895
                  Begin VB.CheckBox chkNameColColorCfg 
                     Caption         =   "��������ɫ��ʾ"
                     Height          =   180
                     Left            =   120
                     TabIndex        =   139
                     ToolTipText     =   "������ɫ���ݲ���������ʾ��"
                     Top             =   0
                     Width           =   1800
                  End
                  Begin VB.CheckBox chkOrdinaryNameColColorCfg 
                     Caption         =   "����ȱʡ����������ɫ"
                     Height          =   255
                     Left            =   600
                     TabIndex        =   140
                     Top             =   240
                     Width           =   2175
                  End
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "��"
                  Height          =   255
                  Index           =   10
                  Left            =   2655
                  TabIndex        =   135
                  Top             =   3600
                  Width           =   255
               End
               Begin VB.TextBox txtAudit 
                  Height          =   270
                  Left            =   4560
                  MaxLength       =   4
                  TabIndex        =   132
                  Text            =   "0"
                  Top             =   2400
                  Width           =   495
               End
               Begin VB.TextBox txtStudy 
                  Height          =   270
                  Left            =   4560
                  MaxLength       =   4
                  TabIndex        =   128
                  Text            =   "0"
                  Top             =   1440
                  Width           =   495
               End
               Begin VB.TextBox txtReport 
                  Height          =   270
                  Left            =   4560
                  MaxLength       =   4
                  TabIndex        =   130
                  Text            =   "0"
                  Top             =   1920
                  Width           =   495
               End
               Begin VB.TextBox txtCheckIn 
                  Height          =   270
                  Left            =   4560
                  MaxLength       =   4
                  TabIndex        =   126
                  Text            =   "0"
                  Top             =   960
                  Width           =   495
               End
               Begin VB.TextBox txtEnreg 
                  Height          =   270
                  Left            =   4560
                  MaxLength       =   4
                  TabIndex        =   124
                  Text            =   "0"
                  Top             =   480
                  Width           =   495
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "��"
                  Height          =   255
                  Index           =   9
                  Left            =   5295
                  TabIndex        =   134
                  Top             =   3120
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "��"
                  Height          =   255
                  Index           =   8
                  Left            =   2650
                  TabIndex        =   123
                  Top             =   480
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "��"
                  Height          =   255
                  Index           =   7
                  Left            =   2650
                  TabIndex        =   133
                  Top             =   3120
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "��"
                  Height          =   255
                  Index           =   6
                  Left            =   2650
                  TabIndex        =   131
                  Top             =   2400
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "��"
                  Height          =   255
                  Index           =   5
                  Left            =   5310
                  TabIndex        =   138
                  Top             =   4080
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "��"
                  Height          =   255
                  Index           =   4
                  Left            =   2650
                  TabIndex        =   129
                  Top             =   1920
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "��"
                  Height          =   255
                  Index           =   3
                  Left            =   2655
                  TabIndex        =   137
                  Top             =   4080
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "��"
                  Height          =   255
                  Index           =   2
                  Left            =   5295
                  TabIndex        =   136
                  Top             =   3600
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "��"
                  Height          =   255
                  Index           =   0
                  Left            =   2650
                  TabIndex        =   127
                  Top             =   1440
                  Width           =   255
               End
               Begin VB.CommandButton cmdColor 
                  Caption         =   "��"
                  Height          =   255
                  Index           =   1
                  Left            =   2650
                  TabIndex        =   125
                  Top             =   960
                  Width           =   255
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   10
                  Left            =   1560
                  Top             =   3600
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "�Ѳ��أ�"
                  Height          =   255
                  Index           =   19
                  Left            =   720
                  TabIndex        =   240
                  Top             =   3600
                  Width           =   735
               End
               Begin VB.Label lab 
                  BackColor       =   &H00000000&
                  BackStyle       =   0  'Transparent
                  Caption         =   "״̬��������        ������"
                  Height          =   255
                  Index           =   25
                  Left            =   3360
                  TabIndex        =   239
                  Top             =   2430
                  Width           =   2415
               End
               Begin VB.Label lab 
                  BackColor       =   &H00000000&
                  BackStyle       =   0  'Transparent
                  Caption         =   "״̬��������        ������"
                  Height          =   255
                  Index           =   23
                  Left            =   3360
                  TabIndex        =   238
                  Top             =   1470
                  Width           =   2415
               End
               Begin VB.Label lab 
                  BackColor       =   &H00000000&
                  BackStyle       =   0  'Transparent
                  Caption         =   "״̬��������        ������"
                  Height          =   255
                  Index           =   24
                  Left            =   3360
                  TabIndex        =   237
                  Top             =   1950
                  Width           =   2415
               End
               Begin VB.Label lab 
                  BackColor       =   &H00000000&
                  BackStyle       =   0  'Transparent
                  Caption         =   "״̬��������        ������"
                  Height          =   255
                  Index           =   22
                  Left            =   3360
                  TabIndex        =   236
                  Top             =   990
                  Width           =   2415
               End
               Begin VB.Label lab 
                  BackColor       =   &H00000000&
                  BackStyle       =   0  'Transparent
                  Caption         =   "״̬��������        ������"
                  Height          =   255
                  Index           =   21
                  Left            =   3360
                  TabIndex        =   235
                  Top             =   510
                  Width           =   2415
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   9
                  Left            =   4200
                  Top             =   3120
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "�Ѿܾ���"
                  Height          =   255
                  Index           =   26
                  Left            =   3360
                  TabIndex        =   234
                  Top             =   3120
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  BackColor       =   &H00FFFFFF&
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   8
                  Left            =   1560
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "�ѵǼǣ�"
                  Height          =   255
                  Index           =   13
                  Left            =   720
                  TabIndex        =   233
                  Top             =   480
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   7
                  Left            =   1560
                  Top             =   3120
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "����ɣ�"
                  Height          =   255
                  Index           =   18
                  Left            =   720
                  TabIndex        =   232
                  Top             =   3120
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   6
                  Left            =   1560
                  Top             =   2400
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "����ˣ�"
                  Height          =   255
                  Index           =   17
                  Left            =   720
                  TabIndex        =   231
                  Top             =   2400
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   5
                  Left            =   4215
                  Top             =   4080
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "����У�"
                  Height          =   255
                  Index           =   28
                  Left            =   3375
                  TabIndex        =   230
                  Top             =   4080
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   4
                  Left            =   1560
                  Top             =   1920
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "�ѱ��棺"
                  Height          =   255
                  Index           =   16
                  Left            =   720
                  TabIndex        =   229
                  Top             =   1920
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   3
                  Left            =   1560
                  Top             =   4080
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "�����У�"
                  Height          =   255
                  Index           =   20
                  Left            =   720
                  TabIndex        =   228
                  Top             =   4080
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   2
                  Left            =   4200
                  Top             =   3600
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "�����У�"
                  Height          =   255
                  Index           =   27
                  Left            =   3360
                  TabIndex        =   227
                  Top             =   3600
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   0
                  Left            =   1560
                  Top             =   1440
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "�Ѽ�飺"
                  Height          =   255
                  Index           =   15
                  Left            =   720
                  TabIndex        =   226
                  Top             =   1440
                  Width           =   735
               End
               Begin VB.Shape shpColor 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Index           =   1
                  Left            =   1560
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.Label lab 
                  Caption         =   "�ѱ�����"
                  Height          =   255
                  Index           =   14
                  Left            =   720
                  TabIndex        =   225
                  Top             =   960
                  Width           =   735
               End
            End
            Begin VB.CommandButton cmdDefault 
               Caption         =   "�ָ�Ĭ��(&D)"
               Height          =   375
               Left            =   -69000
               TabIndex        =   143
               Top             =   6240
               Width           =   1335
            End
            Begin VB.Frame fra 
               Caption         =   "����༭��"
               Height          =   615
               Index           =   19
               Left            =   -74280
               TabIndex        =   221
               Top             =   480
               Width           =   7245
               Begin VB.OptionButton optReportEditor 
                  Caption         =   "PACS���ܱ���༭��"
                  Height          =   255
                  Index           =   2
                  Left            =   4560
                  TabIndex        =   110
                  Top             =   240
                  Width           =   2052
               End
               Begin VB.OptionButton optReportEditor 
                  Caption         =   "���Ӳ����༭��"
                  Height          =   255
                  Index           =   0
                  Left            =   360
                  TabIndex        =   108
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.OptionButton optReportEditor 
                  Caption         =   "PACS����༭��"
                  Height          =   255
                  Index           =   1
                  Left            =   2400
                  TabIndex        =   109
                  Top             =   240
                  Width           =   1575
               End
            End
            Begin VB.Frame fra 
               Caption         =   "��������"
               Height          =   3255
               Index           =   20
               Left            =   -74280
               TabIndex        =   220
               Top             =   2640
               Width           =   7245
               Begin VB.Frame fra 
                  Caption         =   "�����ı�������"
                  Height          =   1335
                  Index           =   26
                  Left            =   240
                  TabIndex        =   365
                  Top             =   1800
                  Width           =   3615
                  Begin VB.TextBox txtCheckView 
                     Height          =   270
                     Left            =   1560
                     TabIndex        =   368
                     Top             =   225
                     Width           =   1335
                  End
                  Begin VB.TextBox txtResult 
                     Height          =   270
                     Left            =   1560
                     TabIndex        =   367
                     Top             =   600
                     Width           =   1335
                  End
                  Begin VB.TextBox txtAdvice 
                     Height          =   270
                     Left            =   1560
                     TabIndex        =   366
                     Top             =   960
                     Width           =   1335
                  End
                  Begin VB.Label lab 
                     Caption         =   "���������"
                     Height          =   255
                     Index           =   10
                     Left            =   360
                     TabIndex        =   371
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.Label lab 
                     Caption         =   "��������"
                     Height          =   255
                     Index           =   11
                     Left            =   360
                     TabIndex        =   370
                     Top             =   608
                     Width           =   975
                  End
                  Begin VB.Label lab 
                     Caption         =   "��    �飺"
                     Height          =   255
                     Index           =   12
                     Left            =   360
                     TabIndex        =   369
                     Top             =   975
                     Width           =   975
                  End
               End
               Begin VB.Frame Frame7 
                  Caption         =   "��ӡ��ʽѡ��ʽ"
                  Height          =   1275
                  Left            =   4200
                  TabIndex        =   298
                  Top             =   1800
                  Width           =   2745
                  Begin VB.CheckBox chkPrintFormat 
                     Caption         =   "��ѡ�����ʽ"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   310
                     Top             =   960
                     Width           =   1455
                  End
                  Begin VB.OptionButton optPrintFormat 
                     Caption         =   "��¼���һ�δ�ӡ��ʽ"
                     Height          =   255
                     Index           =   0
                     Left            =   240
                     TabIndex        =   300
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   2175
                  End
                  Begin VB.OptionButton optPrintFormat 
                     Caption         =   "ʼ�ձ���Ĭ�ϸ�ʽ"
                     Height          =   255
                     Index           =   1
                     Left            =   240
                     TabIndex        =   299
                     Top             =   600
                     Width           =   1815
                  End
               End
               Begin VB.CheckBox chkUntreadPrinted 
                  Caption         =   "��˴�ӡ���������"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   115
                  Top             =   720
                  Width           =   2055
               End
               Begin VB.CheckBox chkSpecialContent 
                  Caption         =   "��ʾר�Ʊ������ݣ�"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   116
                  Top             =   1080
                  Width           =   2055
               End
               Begin VB.ComboBox cboSpecialContent 
                  Enabled         =   0   'False
                  Height          =   300
                  Left            =   240
                  TabIndex        =   117
                  Text            =   "Combo1"
                  Top             =   1360
                  Width           =   6735
               End
               Begin VB.CheckBox chkExitAfterPrint 
                  Caption         =   "��ӡ���˳�"
                  Height          =   180
                  Left            =   2280
                  TabIndex        =   114
                  Top             =   720
                  Width           =   1335
               End
               Begin VB.CheckBox chkShowVideoCapture 
                  Caption         =   "��ʾ��Ƶ�ɼ�����"
                  Height          =   180
                  Left            =   2280
                  TabIndex        =   113
                  Top             =   360
                  Width           =   2055
               End
               Begin VB.TextBox txtMinImageCount 
                  Height          =   270
                  Left            =   6240
                  MaxLength       =   2
                  TabIndex        =   112
                  Text            =   "8"
                  Top             =   315
                  Width           =   615
               End
               Begin VB.CheckBox chkShowImage 
                  Caption         =   "��ʾ����ͼ������                            ��������ͼ��ʾ������"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   111
                  Top             =   360
                  Width           =   6135
               End
            End
            Begin VB.Frame fra 
               Caption         =   "����ʾ�˫����"
               Height          =   855
               Index           =   21
               Left            =   -74280
               TabIndex        =   219
               Top             =   6060
               Width           =   2415
               Begin VB.OptionButton optWordDblClick 
                  Caption         =   "ֱ��д�뱨��"
                  Height          =   255
                  Index           =   0
                  Left            =   360
                  TabIndex        =   118
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.OptionButton optWordDblClick 
                  Caption         =   "�򿪴ʾ�༭����"
                  Height          =   255
                  Index           =   1
                  Left            =   360
                  TabIndex        =   119
                  Top             =   480
                  Width           =   1750
               End
            End
            Begin VB.Frame fra 
               Caption         =   "����ͼ˫����"
               Height          =   855
               Index           =   22
               Left            =   -71880
               TabIndex        =   218
               Top             =   6060
               Width           =   2415
               Begin VB.OptionButton optImageDblClick 
                  Caption         =   "��ͼƬ�༭����"
                  Height          =   255
                  Index           =   1
                  Left            =   360
                  TabIndex        =   121
                  Top             =   480
                  Width           =   1750
               End
               Begin VB.OptionButton optImageDblClick 
                  Caption         =   "ֱ��д�뱨��"
                  Height          =   255
                  Index           =   0
                  Left            =   360
                  TabIndex        =   120
                  Top             =   240
                  Width           =   1575
               End
            End
            Begin VB.Frame fra 
               Caption         =   "�ʾ�ģ����ʾ"
               Height          =   855
               Index           =   23
               Left            =   -69480
               TabIndex        =   217
               Top             =   6060
               Width           =   2450
               Begin VB.OptionButton optShowWord 
                  Caption         =   "˫������"
                  Height          =   180
                  Index           =   1
                  Left            =   360
                  TabIndex        =   150
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.OptionButton optShowWord 
                  Caption         =   "ֱ����ʾ"
                  Height          =   180
                  Index           =   0
                  Left            =   360
                  TabIndex        =   122
                  Top             =   240
                  Width           =   1095
               End
            End
            Begin VB.Frame fra 
               Caption         =   "��������"
               Height          =   4815
               Index           =   16
               Left            =   -74280
               TabIndex        =   212
               Top             =   480
               Width           =   7335
               Begin VB.CheckBox chkSelectRoom 
                  Caption         =   "����ʱ����Ĭ��ִ�м�"
                  Height          =   210
                  Left            =   3840
                  TabIndex        =   98
                  Top             =   4485
                  Width           =   2220
               End
               Begin VB.CommandButton cmdAddGroup 
                  Caption         =   "��������(&A)"
                  Height          =   375
                  Left            =   120
                  Picture         =   "frmParPacs.frx":14E46
                  TabIndex        =   95
                  TabStop         =   0   'False
                  Top             =   4380
                  Width           =   1170
               End
               Begin VB.CommandButton cmdDelGroup 
                  Caption         =   "ɾ������(&D)"
                  Height          =   375
                  Left            =   1320
                  Picture         =   "frmParPacs.frx":14F90
                  TabIndex        =   96
                  TabStop         =   0   'False
                  Top             =   4380
                  Width           =   1170
               End
               Begin VB.CommandButton cmdStudyAcc 
                  Caption         =   "������Ŀ(&R)"
                  Height          =   375
                  Left            =   6000
                  Picture         =   "frmParPacs.frx":150DA
                  TabIndex        =   99
                  TabStop         =   0   'False
                  Top             =   4380
                  Width           =   1155
               End
               Begin VB.CommandButton cmdModify 
                  Caption         =   "�޸ķ���(&M)"
                  Height          =   375
                  Left            =   2520
                  Picture         =   "frmParPacs.frx":15224
                  TabIndex        =   97
                  TabStop         =   0   'False
                  Top             =   4380
                  Width           =   1170
               End
               Begin zl9BaseItem.ucFlexGrid ufgStudyProCfg 
                  Height          =   1905
                  Left            =   3405
                  TabIndex        =   94
                  Top             =   2430
                  Width           =   3735
                  _ExtentX        =   6588
                  _ExtentY        =   3360
                  DefaultCols     =   ""
                  ColNames        =   "|���������Ŀ>����,w2100,read|��Ŀ����>����,w1100,read|"
                  KeyName         =   "��"
                  HeadCheckValue  =   1
                  IsCopyAdoMode   =   0   'False
                  IsEjectConfig   =   -1  'True
                  HeadFontCharset =   134
                  HeadFontWeight  =   400
                  HeadColor       =   0
                  DataFontCharset =   134
                  DataFontWeight  =   400
                  DataColor       =   -2147483640
                  GridLineColor   =   14737632
               End
               Begin zl9BaseItem.ucFlexGrid ufgRoomCfg 
                  Height          =   2175
                  Left            =   3405
                  TabIndex        =   93
                  Top             =   240
                  Width           =   3735
                  _ExtentX        =   6588
                  _ExtentY        =   3836
                  DefaultCols     =   ""
                  ColNames        =   "|ID,hide|ִ�м�,w1400,read|����ǰ׺,w1400,read|"
                  KeyName         =   "��"
                  HeadCheckValue  =   1
                  IsCopyAdoMode   =   0   'False
                  IsEjectConfig   =   -1  'True
                  HeadFontCharset =   134
                  HeadFontWeight  =   400
                  HeadColor       =   0
                  DataFontCharset =   134
                  DataFontWeight  =   400
                  DataColor       =   -2147483640
                  GridLineColor   =   14737632
               End
               Begin zl9BaseItem.ucFlexGrid ufgGroupCfg 
                  Height          =   4095
                  Left            =   120
                  TabIndex        =   92
                  Top             =   240
                  Width           =   3255
                  _ExtentX        =   5741
                  _ExtentY        =   7223
                  DefaultCols     =   ""
                  ColNames        =   "|ID,hide,key|����,w1400,read|����ǰ׺,w1500,read|"
                  KeyName         =   "ID"
                  IsCopyAdoMode   =   0   'False
                  IsEjectConfig   =   -1  'True
                  HeadFontCharset =   134
                  HeadFontWeight  =   400
                  DataFontCharset =   134
                  DataFontWeight  =   400
               End
            End
            Begin VB.Frame fra 
               Height          =   1515
               Index           =   17
               Left            =   -74280
               TabIndex        =   207
               Top             =   5400
               Width           =   7335
               Begin VB.CheckBox chkAutoInQueue 
                  Caption         =   "�������Զ��Ŷ�"
                  Height          =   180
                  Left            =   3480
                  TabIndex        =   106
                  Top             =   1125
                  Value           =   1  'Checked
                  Width           =   1575
               End
               Begin VB.CheckBox chkUseQueueMsg 
                  Caption         =   "�����Ŷ���Ϣ����"
                  Height          =   180
                  Left            =   5160
                  TabIndex        =   107
                  Top             =   1125
                  Value           =   1  'Checked
                  Width           =   1815
               End
               Begin VB.ComboBox cbxPrintQueueNoWay 
                  Height          =   300
                  ItemData        =   "frmParPacs.frx":1536E
                  Left            =   1635
                  List            =   "frmParPacs.frx":1537B
                  Style           =   2  'Dropdown List
                  TabIndex        =   105
                  Top             =   1080
                  Width           =   1740
               End
               Begin VB.Frame fra 
                  Caption         =   "δָ��ִ�м���Ŷӷ�ʽ"
                  Height          =   810
                  Index           =   18
                  Left            =   4680
                  TabIndex        =   208
                  Top             =   240
                  Width           =   2265
                  Begin VB.OptionButton optNumberRule 
                     Caption         =   "���������Ŷ�"
                     Height          =   180
                     Index           =   0
                     Left            =   105
                     TabIndex        =   103
                     ToolTipText     =   "���ڷ�����ִ�м�ļ�飬�ŶӺ��뽫��ִ�м��������ɣ���δ����ִ�еļ�飬�ŶӺ��뽫�������������ɡ�"
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   1755
                  End
                  Begin VB.OptionButton optNumberRule 
                     Caption         =   "���������Ŷ�"
                     Height          =   180
                     Index           =   1
                     Left            =   105
                     TabIndex        =   104
                     ToolTipText     =   "���ڷ�����ִ�м�ļ�飬�ŶӺ��뽫��ִ�м��������ɣ���δ����ִ�еļ�飬�ŶӺ��뽫���ݼ�����������������ɡ�"
                     Top             =   480
                     Width           =   1665
                  End
               End
               Begin VB.CheckBox chkSynStudyList 
                  Caption         =   "ͬ����λ����б�"
                  Height          =   180
                  Left            =   2760
                  TabIndex        =   101
                  ToolTipText     =   "����Ŷ��б������б����ݺ�ͬ����λ������б�"
                  Top             =   330
                  Width           =   1815
               End
               Begin VB.TextBox txtQueueReport 
                  Height          =   315
                  Left            =   1635
                  TabIndex        =   102
                  Top             =   690
                  Width           =   2820
               End
               Begin VB.TextBox txtValidDays 
                  Height          =   315
                  Left            =   1635
                  MaxLength       =   2
                  TabIndex        =   100
                  Text            =   "1"
                  Top             =   285
                  Width           =   555
               End
               Begin VB.Label lab 
                  Caption         =   "�źŵ���ӡ��ʽ��"
                  Height          =   255
                  Index           =   9
                  Left            =   240
                  TabIndex        =   211
                  Top             =   1110
                  Width           =   1455
               End
               Begin VB.Label lab 
                  Caption         =   "�źŵ������ţ�"
                  Height          =   225
                  Index           =   8
                  Left            =   240
                  TabIndex        =   210
                  ToolTipText     =   "�ŶӴ��ʱ��Ӧ���Զ��屨���š�"
                  Top             =   735
                  Width           =   1455
               End
               Begin VB.Label lab 
                  Caption         =   "������Ч������       ��"
                  Height          =   210
                  Index           =   7
                  Left            =   420
                  TabIndex        =   209
                  Top             =   360
                  Width           =   2235
               End
            End
            Begin VB.Frame fra 
               Caption         =   "ƴ����"
               Height          =   1695
               Index           =   4
               Left            =   5280
               TabIndex        =   205
               Top             =   5040
               Width           =   2775
               Begin VB.OptionButton optCapital 
                  Caption         =   "��д"
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   29
                  ToolTipText     =   "ѡ���ƴ������ʾȫΪ��д��ĸ��"
                  Top             =   260
                  Width           =   735
               End
               Begin VB.OptionButton optCapital 
                  Caption         =   "Сд"
                  Height          =   255
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   30
                  ToolTipText     =   "ѡ���ƴ������ʾȫΪСд��ĸ��"
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton optCapital 
                  Caption         =   "����ĸ��д"
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   31
                  ToolTipText     =   "ѡ���ƴ��������ĸ��д��"
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Frame fra 
                  Caption         =   "���"
                  Height          =   540
                  Index           =   11
                  Left            =   240
                  TabIndex        =   206
                  Top             =   960
                  Width           =   1695
                  Begin VB.OptionButton optSplitter 
                     Caption         =   "��"
                     Height          =   255
                     Index           =   1
                     Left            =   960
                     TabIndex        =   33
                     ToolTipText     =   "ƴ����֮���޼����"
                     Top             =   200
                     Width           =   495
                  End
                  Begin VB.OptionButton optSplitter 
                     Caption         =   "�ո�"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   32
                     ToolTipText     =   "ƴ����֮��ʹ�ÿո�Ϊ�������"
                     Top             =   200
                     Width           =   735
                  End
               End
            End
            Begin VB.Frame fra 
               Caption         =   "�ȼ��󱨵���ͼ��ƥ��"
               Height          =   1545
               Index           =   3
               Left            =   5280
               TabIndex        =   204
               Top             =   3360
               Width           =   2775
               Begin VB.OptionButton optMatch 
                  Caption         =   "����/סԺ��"
                  Height          =   195
                  Index           =   1
                  Left            =   240
                  TabIndex        =   28
                  ToolTipText     =   "����ʱͨ������/סԺ�ź�ͼ����Ϣ����ƥ�䣬������Ӱ��ҽ��վ��"
                  Top             =   1000
                  Width           =   1335
               End
               Begin VB.OptionButton optMatch 
                  Caption         =   "����"
                  Height          =   195
                  Index           =   0
                  Left            =   240
                  TabIndex        =   26
                  ToolTipText     =   "����ʱͨ�����ź�ͼ����Ϣ����ƥ�䣬������Ӱ��ҽ��վ��"
                  Top             =   360
                  Width           =   855
               End
               Begin VB.OptionButton optMatch 
                  Caption         =   "ҽ��ID"
                  Height          =   195
                  Index           =   2
                  Left            =   240
                  TabIndex        =   27
                  ToolTipText     =   "����ʱͨ��ҽ��ID��ͼ����Ϣ����ƥ�䣬������Ӱ��ҽ��վ��"
                  Top             =   680
                  Width           =   855
               End
            End
            Begin VB.Frame fra 
               Caption         =   "��������"
               Height          =   1665
               Index           =   2
               Left            =   720
               TabIndex        =   202
               Top             =   3240
               Width           =   4215
               Begin VB.Frame Frame1 
                  Caption         =   "���ٹ�������"
                  Height          =   780
                  Left            =   1920
                  TabIndex        =   302
                  Top             =   840
                  Width           =   2055
                  Begin VB.CheckBox chkNameQueryTimeLimit 
                     Caption         =   "������ѯʱ������"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   304
                     ToolTipText     =   "��������ѯʱ���Ƿ��в�ѯʱ������"
                     Top             =   480
                     Width           =   1850
                  End
                  Begin VB.CheckBox chkNameFuzzySearch 
                     Caption         =   "����Ĭ��ģ����ѯ"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   303
                     ToolTipText     =   "��������ѯʱʹ��ģ����ѯ��û�й�ѡʱ��ֻ������*��Ž���ģ����ѯ"
                     Top             =   240
                     Width           =   1850
                  End
               End
               Begin VB.CheckBox chkSwitchUser 
                  Caption         =   "�����л��û�"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   22
                  ToolTipText     =   "�����л��û����ܣ����Խ����û��л���"
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.Frame fra 
                  Height          =   660
                  Index           =   10
                  Left            =   1920
                  TabIndex        =   203
                  ToolTipText     =   "ѡ��ɼ�ͼ���ɨ�����뵥��ʹ�õĴ洢�豸��"
                  Top             =   180
                  Width           =   2055
                  Begin VB.ComboBox cboSaveDevice 
                     Height          =   300
                     Left            =   120
                     Style           =   2  'Dropdown List
                     TabIndex        =   25
                     Top             =   240
                     Width           =   1605
                  End
                  Begin VB.CheckBox chkPetitionCapture 
                     Caption         =   "�������뵥ɨ��"
                     Height          =   180
                     Left            =   120
                     TabIndex        =   24
                     ToolTipText     =   "�������뵥ɨ�蹦��"
                     Top             =   0
                     Value           =   1  'Checked
                     Width           =   1575
                  End
               End
               Begin VB.CheckBox chkUseReferencePatient 
                  Caption         =   "���ù�������"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   23
                  ToolTipText     =   "֧�ֶ����������ͬһ��������Ϣ��"
                  Top             =   1080
                  Width           =   1455
               End
               Begin VB.CheckBox chkChangeUser 
                  Caption         =   "���ý����û�"
                  Height          =   180
                  Left            =   240
                  TabIndex        =   21
                  ToolTipText     =   "������û����ܣ����Խ������ҽ���ͱ���ҽ����������Ӱ��ɼ�վ��"
                  Top             =   360
                  Width           =   1455
               End
            End
         End
         Begin MSComDlg.CommonDialog dlgColor 
            Left            =   4920
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6975
         Index           =   5
         Left            =   -75000
         ScaleHeight     =   6975
         ScaleWidth      =   7815
         TabIndex        =   196
         Top             =   360
         Width           =   7815
         Begin VB.Frame fra 
            Height          =   1695
            Index           =   40
            Left            =   120
            TabIndex        =   282
            Top             =   120
            Width           =   4455
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   3
               Left            =   2160
               TabIndex        =   176
               Top             =   720
               Width           =   2175
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   21
               Left            =   2160
               TabIndex        =   175
               Text            =   "100"
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "����ȷ�Ϻ��Զ���ӡ���Ļ�ִ��"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   177
               Top             =   1200
               Width           =   2895
            End
            Begin VB.Label lab 
               Caption         =   "��"
               Height          =   255
               Index           =   56
               Left            =   3000
               TabIndex        =   285
               Top             =   280
               Width           =   255
            End
            Begin VB.Label lab 
               Caption         =   "���Ļ�ִ��Ӧ�������ƣ�"
               Height          =   255
               Index           =   57
               Left            =   120
               TabIndex        =   284
               Top             =   760
               Width           =   2055
            End
            Begin VB.Label lab 
               Caption         =   "���ļ�¼Ĭ�ϲ�ѯ������"
               Height          =   255
               Index           =   55
               Left            =   120
               TabIndex        =   283
               Top             =   285
               Width           =   2055
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6975
         Index           =   4
         Left            =   -75000
         ScaleHeight     =   6975
         ScaleWidth      =   7815
         TabIndex        =   195
         Top             =   360
         Width           =   7815
         Begin VB.Frame fra 
            Height          =   1215
            Index           =   39
            Left            =   120
            TabIndex        =   278
            Top             =   120
            Width           =   4455
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   20
               Left            =   2160
               TabIndex        =   173
               Text            =   "30"
               Top             =   240
               Width           =   735
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   2
               Left            =   2160
               TabIndex        =   174
               Top             =   720
               Width           =   2175
            End
            Begin VB.Label lab 
               Caption         =   "������¼Ĭ�ϲ�ѯ������"
               Height          =   255
               Index           =   52
               Left            =   120
               TabIndex        =   281
               Top             =   285
               Width           =   2055
            End
            Begin VB.Label lab 
               Caption         =   "������ǩ��Ӧ�������ƣ�"
               Height          =   255
               Index           =   54
               Left            =   120
               TabIndex        =   280
               Top             =   760
               Width           =   2055
            End
            Begin VB.Label lab 
               Caption         =   "��"
               Height          =   255
               Index           =   53
               Left            =   3000
               TabIndex        =   279
               Top             =   280
               Width           =   255
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6975
         Index           =   3
         Left            =   -75000
         ScaleHeight     =   6975
         ScaleWidth      =   7815
         TabIndex        =   194
         Top             =   360
         Width           =   7815
         Begin VB.CheckBox chk 
            Caption         =   "¼����Ժ��Ϣ"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   351
            Top             =   4440
            Width           =   1455
         End
         Begin VB.Frame Frame4 
            Caption         =   "Frame4"
            Height          =   735
            Left            =   120
            TabIndex        =   313
            Top             =   4440
            Width           =   4815
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   22
               Left            =   1680
               TabIndex        =   353
               Top             =   300
               Width           =   2895
            End
            Begin VB.Label lab 
               Caption         =   "�ͼ쵥λ���ã�"
               Height          =   255
               Index           =   61
               Left            =   240
               TabIndex        =   352
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.ListBox lst 
            Height          =   2790
            Index           =   0
            Left            =   5280
            Style           =   1  'Checkbox
            TabIndex        =   311
            Top             =   600
            Width           =   2055
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   5
            ItemData        =   "frmParPacs.frx":153AB
            Left            =   1410
            List            =   "frmParPacs.frx":153B8
            Style           =   2  'Dropdown List
            TabIndex        =   307
            Top             =   3960
            Width           =   2205
         End
         Begin VB.Frame fra 
            Caption         =   "����ѡ���ͼ�����ʱ�Զ�������������"
            Height          =   1095
            Index           =   38
            Left            =   120
            TabIndex        =   277
            Top             =   2640
            Width           =   4815
            Begin VB.CheckBox chk 
               Caption         =   "����"
               Height          =   375
               Index           =   2
               Left            =   480
               TabIndex        =   167
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "����"
               Height          =   375
               Index           =   3
               Left            =   1800
               TabIndex        =   168
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "ϸ��"
               Height          =   375
               Index           =   4
               Left            =   3240
               TabIndex        =   169
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "����"
               Height          =   375
               Index           =   5
               Left            =   480
               TabIndex        =   170
               Top             =   600
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "ʬ��"
               Height          =   375
               Index           =   6
               Left            =   1800
               TabIndex        =   171
               Top             =   600
               Width           =   735
            End
            Begin VB.CheckBox chk 
               Caption         =   "��Ƭ"
               Height          =   375
               Index           =   7
               Left            =   3240
               TabIndex        =   172
               Top             =   600
               Width           =   735
            End
         End
         Begin VB.Frame fra 
            Caption         =   "�ʾ�ģ������"
            Height          =   2415
            Index           =   37
            Left            =   120
            TabIndex        =   269
            Top             =   120
            Width           =   4815
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   15
               Left            =   1560
               TabIndex        =   162
               Top             =   495
               Width           =   3015
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   16
               Left            =   1560
               TabIndex        =   163
               Top             =   855
               Width           =   3015
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   17
               Left            =   1560
               TabIndex        =   164
               Top             =   1215
               Width           =   3015
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   18
               Left            =   1560
               TabIndex        =   165
               Top             =   1575
               Width           =   3015
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   19
               Left            =   1560
               TabIndex        =   270
               Top             =   1935
               Width           =   3015
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ӧ�ʾ����"
               Height          =   180
               Index           =   51
               Left            =   2400
               TabIndex        =   276
               Top             =   240
               Width           =   1080
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�޼�����ģ�壺"
               Height          =   180
               Index           =   46
               Left            =   240
               TabIndex        =   275
               Top             =   540
               Width           =   1260
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���汨��ģ�壺"
               Height          =   180
               Index           =   47
               Left            =   240
               TabIndex        =   274
               Top             =   900
               Width           =   1260
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���߱���ģ�壺"
               Height          =   180
               Index           =   48
               Left            =   240
               TabIndex        =   273
               Top             =   1260
               Width           =   1260
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ӱ���ģ�壺"
               Height          =   180
               Index           =   50
               Left            =   240
               TabIndex        =   272
               Top             =   1980
               Width           =   1260
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ⱦ����ģ�壺"
               Height          =   180
               Index           =   49
               Left            =   240
               TabIndex        =   271
               Top             =   1620
               Width           =   1260
            End
         End
         Begin VB.Label lab 
            AutoSize        =   -1  'True
            Caption         =   "����ȡ����ʾ��Ϣ"
            Height          =   180
            Index           =   60
            Left            =   5280
            TabIndex        =   312
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lab 
            Caption         =   "����ִ��ģʽ��"
            Height          =   270
            Index           =   59
            Left            =   120
            TabIndex        =   308
            Top             =   4005
            Width           =   1305
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6975
         Index           =   2
         Left            =   -75000
         ScaleHeight     =   6975
         ScaleWidth      =   7815
         TabIndex        =   193
         Top             =   360
         Width           =   7815
         Begin VB.Frame fra 
            Height          =   1095
            Index           =   36
            Left            =   120
            TabIndex        =   290
            Top             =   120
            Width           =   4095
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   1
               ItemData        =   "frmParPacs.frx":153E0
               Left            =   1410
               List            =   "frmParPacs.frx":153ED
               Style           =   2  'Dropdown List
               TabIndex        =   292
               Top             =   600
               Width           =   2205
            End
            Begin VB.CheckBox chk 
               Caption         =   "¼����Ժ��Ϣ"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   291
               Top             =   240
               Width           =   1590
            End
            Begin VB.Label lab 
               Caption         =   "����ִ��ģʽ��"
               Height          =   270
               Index           =   45
               Left            =   120
               TabIndex        =   293
               Top             =   645
               Width           =   1305
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6975
         Index           =   1
         Left            =   0
         ScaleHeight     =   6975
         ScaleWidth      =   7815
         TabIndex        =   192
         Top             =   360
         Width           =   7815
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   4
            ItemData        =   "frmParPacs.frx":15415
            Left            =   1530
            List            =   "frmParPacs.frx":1541F
            Style           =   2  'Dropdown List
            TabIndex        =   305
            Top             =   6000
            Width           =   2205
         End
         Begin VB.CheckBox chk 
            Caption         =   "¼����Ժ��Ϣ"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   144
            Top             =   120
            Width           =   1590
         End
         Begin VB.Frame fra 
            Caption         =   "XWPACS��Ƭ"
            Height          =   5295
            Index           =   31
            Left            =   240
            TabIndex        =   248
            Top             =   600
            Width           =   7335
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   23
               Left            =   1590
               TabIndex        =   381
               Top             =   3840
               Width           =   5640
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   11
               Left            =   1590
               TabIndex        =   158
               Text            =   "zlhis"
               Top             =   4320
               Width           =   5640
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   12
               Left            =   5430
               TabIndex        =   287
               Text            =   "DCMSHARE"
               Top             =   3840
               Visible         =   0   'False
               Width           =   1785
            End
            Begin VB.Frame fra 
               Caption         =   "ɾ��ͼ���û�"
               Height          =   1215
               Index           =   32
               Left            =   120
               TabIndex        =   266
               Top             =   480
               Width           =   2295
               Begin VB.TextBox txt 
                  Height          =   270
                  IMEMode         =   3  'DISABLE
                  Index           =   1
                  Left            =   720
                  PasswordChar    =   "*"
                  TabIndex        =   146
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.TextBox txt 
                  Height          =   270
                  Index           =   0
                  Left            =   720
                  TabIndex        =   145
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.Label lab 
                  Caption         =   "����"
                  Height          =   255
                  Index           =   33
                  Left            =   120
                  TabIndex        =   268
                  Top             =   780
                  Width           =   615
               End
               Begin VB.Label lab 
                  Caption         =   "�û���"
                  Height          =   255
                  Index           =   32
                  Left            =   120
                  TabIndex        =   267
                  Top             =   420
                  Width           =   615
               End
            End
            Begin VB.Frame fra 
               Caption         =   "����ͼ���û�"
               Height          =   1215
               Index           =   33
               Left            =   2520
               TabIndex        =   263
               Top             =   480
               Width           =   2295
               Begin VB.TextBox txt 
                  Height          =   270
                  Index           =   2
                  Left            =   720
                  TabIndex        =   147
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.TextBox txt 
                  Height          =   270
                  IMEMode         =   3  'DISABLE
                  Index           =   3
                  Left            =   720
                  PasswordChar    =   "*"
                  TabIndex        =   148
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.Label lab 
                  Caption         =   "�û���"
                  Height          =   255
                  Index           =   34
                  Left            =   120
                  TabIndex        =   265
                  Top             =   420
                  Width           =   615
               End
               Begin VB.Label lab 
                  Caption         =   "����"
                  Height          =   255
                  Index           =   35
                  Left            =   120
                  TabIndex        =   264
                  Top             =   780
                  Width           =   615
               End
            End
            Begin VB.Frame fra 
               Caption         =   "���̿�¼�û�"
               Height          =   1215
               Index           =   34
               Left            =   4920
               TabIndex        =   260
               Top             =   480
               Width           =   2295
               Begin VB.TextBox txt 
                  Height          =   270
                  IMEMode         =   3  'DISABLE
                  Index           =   5
                  Left            =   720
                  PasswordChar    =   "*"
                  TabIndex        =   151
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.TextBox txt 
                  Height          =   270
                  Index           =   4
                  Left            =   720
                  TabIndex        =   149
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.Label lab 
                  Caption         =   "����"
                  Height          =   255
                  Index           =   37
                  Left            =   120
                  TabIndex        =   262
                  Top             =   780
                  Width           =   615
               End
               Begin VB.Label lab 
                  Caption         =   "�û���"
                  Height          =   255
                  Index           =   36
                  Left            =   120
                  TabIndex        =   261
                  Top             =   420
                  Width           =   615
               End
            End
            Begin VB.Frame fra 
               Caption         =   "XWPACS ���ݿ������"
               Height          =   855
               Index           =   35
               Left            =   120
               TabIndex        =   249
               Top             =   1800
               Width           =   7095
               Begin VB.CommandButton cmdTestCon 
                  Caption         =   "��"
                  Height          =   375
                  Left            =   6600
                  TabIndex        =   155
                  Top             =   300
                  Width           =   375
               End
               Begin VB.TextBox txt 
                  Height          =   270
                  Index           =   6
                  Left            =   840
                  TabIndex        =   152
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.TextBox txt 
                  Height          =   270
                  Index           =   7
                  Left            =   3000
                  TabIndex        =   153
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.TextBox txt 
                  Height          =   270
                  IMEMode         =   3  'DISABLE
                  Index           =   8
                  Left            =   5040
                  PasswordChar    =   "*"
                  TabIndex        =   154
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.Label lab 
                  Caption         =   "������"
                  Height          =   255
                  Index           =   29
                  Left            =   240
                  TabIndex        =   252
                  Top             =   420
                  Width           =   855
               End
               Begin VB.Label lab 
                  Caption         =   "�û���"
                  Height          =   255
                  Index           =   30
                  Left            =   2400
                  TabIndex        =   251
                  Top             =   420
                  Width           =   855
               End
               Begin VB.Label lab 
                  Caption         =   "����"
                  Height          =   255
                  Index           =   31
                  Left            =   4560
                  TabIndex        =   250
                  Top             =   420
                  Width           =   495
               End
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   13
               Left            =   1320
               TabIndex        =   159
               Text            =   "1"
               Top             =   4800
               Width           =   372
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   14
               Left            =   3135
               TabIndex        =   160
               Text            =   "2"
               Top             =   4800
               Width           =   372
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   9
               Left            =   1590
               TabIndex        =   156
               Text            =   "http://127.0.0.1:8080/TakeImage.aspx?colid0=22&colvalue0=[@STU_NO]"
               Top             =   2880
               Width           =   5640
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   0
               ItemData        =   "frmParPacs.frx":1543B
               Left            =   5400
               List            =   "frmParPacs.frx":15445
               Style           =   2  'Dropdown List
               TabIndex        =   161
               Top             =   4800
               Width           =   1812
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   10
               Left            =   1590
               TabIndex        =   157
               Text            =   "http://127.0.0.1:8080/KeyImage.aspx?colid0=22&colvalue0=[@STU_NO]"
               Top             =   3360
               Width           =   5640
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               Caption         =   "����б��Ƭ��ַ"
               Height          =   180
               Index           =   2
               Left            =   120
               TabIndex        =   382
               Top             =   3885
               Width           =   1560
            End
            Begin VB.Label lab 
               Caption         =   "��ʷͼ����Ŀ¼"
               Height          =   255
               Index           =   41
               Left            =   3855
               TabIndex        =   259
               Top             =   3900
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               Caption         =   "�ӿڰ�ӵ����"
               Height          =   240
               Index           =   40
               Left            =   120
               TabIndex        =   258
               Top             =   4380
               Width           =   1440
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               Caption         =   "��鷽����"
               Height          =   180
               Index           =   42
               Left            =   240
               TabIndex        =   257
               Top             =   4845
               Width           =   900
            End
            Begin VB.Label lab 
               Caption         =   "���з�����"
               Height          =   180
               Index           =   43
               Left            =   2160
               TabIndex        =   256
               Top             =   4845
               Width           =   975
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               Caption         =   "WEB��Ƭ��ַ"
               Height          =   240
               Index           =   38
               Left            =   255
               TabIndex        =   255
               Top             =   2925
               Width           =   1320
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               Caption         =   "3D��Ƭ����"
               Height          =   240
               Index           =   44
               Left            =   4380
               TabIndex        =   254
               Top             =   4845
               Width           =   960
            End
            Begin VB.Label lab 
               AutoSize        =   -1  'True
               Caption         =   "�ؼ�ͼ���ַ"
               Height          =   240
               Index           =   39
               Left            =   135
               TabIndex        =   253
               Top             =   3405
               Width           =   1440
            End
         End
         Begin VB.Label lab 
            Caption         =   "����ִ��ģʽ��"
            Height          =   270
            Index           =   58
            Left            =   240
            TabIndex        =   306
            Top             =   6045
            Width           =   1305
         End
      End
   End
End
Attribute VB_Name = "frmParPacs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type

Private Enum TNeedType
    tNeedName = 0
    tNeedNo = 1
    tNeedAll = 2
End Enum


Private Const Report_Form_frmReportES  As String = "�ھ�������Ϣ"
Private Const Report_Form_frmReportPathology As String = "������Һ��������Ϣ"
Private Const Report_Form_frmReportUS As String = "B�����������Ϣ"
Private Const Report_Form_frmReportCustom As String = "�Զ���ר�Ʊ���"


Private mrsPar As ADODB.Recordset '������ؼ���Ӧ��¼����ͬһ���������ܶ�Ӧһ�����ؼ���
Private marrFunc(1) As String
Private mlngPreFind As Long

Private mrsDeptParas As ADODB.Recordset '���Ʋ�������

Private mstrPrivs As String         '��ģ���Ȩ��
Private mlng����ID As Long          'IN:��ǰִ�п���ID
Private mlngCur����ID As Long       '��ǰ����ID
Private mstrCur���� As String       '��ǰ���� ����-����
Private mstrCanUse���� As String    '��ǰ���ÿ���  ID_����-����
Private mblnOk As Boolean

Private UserInfo As TYPE_USER_INFO

Private Enum constLst
    lst_PatholInfo = 0
End Enum

Private Enum constTxtLocate
    txt_Par = 0
    txt_Dept = 1
End Enum

Private Enum constChk
    chk_¼����Ժ��������Ϣ = 0
    chk_¼����Ժ���������Ϣ = 1
    chk_���没������������ = 2
    chk_������������������ = 3
    chk_ϸ���������������� = 4
    chk_���ﲡ������������ = 5
    chk_ʬ�첡������������ = 6
    chk_��Ƭ�������������� = 7
    chk_���ĺ��Զ���ӡ��ִ�� = 8
    chk_¼����Ժ��Ϣ = 9
End Enum

Private Enum constCbo
    cbo_3D��Ƭ���� = 0
    cbo_�ɼ�����ִ��ģʽ = 1    '�ɼ�վ����ִ��ģʽ0-����ʱִ�У�1-���ʱִ�У�2-����ʱִ��
    cbo_������ǩ�������� = 2
    cbo_���Ļ�ִ�������� = 3
    cbo_ҽ������ִ��ģʽ = 4    '�ɼ�վ����ִ��ģʽ0-����ʱִ�У�1-����ʱִ��
    cbo_�������ִ��ģʽ = 5    '����վ����ִ��ģʽ0-����ʱִ�У�1-���ʱִ�У�2-����ʱִ��
End Enum


Private Enum constTxt
    txt_ͼ��ɾ���û����� = 0
    txt_ͼ��ɾ���û����� = 1
    txt_ͼ�����û����� = 2
    txt_ͼ�����û����� = 3
    txt_ͼ���¼�û����� = 4
    txt_ͼ���¼�û����� = 5
    txt_������IP = 6
    txt_�������û����� = 7
    txt_�������û����� = 8
    txt_WEB��Ƭ��ַ = 9
    txt_�ؼ�ͼ���ַ = 10
    txt_��ӵ���� = 11
    'txt_����Ŀ¼ = 12
    txt_��鷽���� = 13
    txt_���з����� = 14
    txt_�޼�����ģ�� = 15
    txt_��������ģ�� = 16
    txt_��������ģ�� = 17
    txt_��Ⱦ����ģ�� = 18
    txt_��������ģ�� = 19
    txt_����Ĭ�ϲ�ѯ���� = 20
    txt_�軹Ĭ�ϲ�ѯ���� = 21
    txt_������Ժ���� = 22
    txt_����б��Ƭ��ַ = 23
End Enum
'
'Private Enum constListBox
'    lst_סԺ�����Ժ��� = 0
'End Enum
Private Const SERVICE_START = &H10
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const SC_MANAGER_CONNECT As Long = &H1
Private Const SC_MANAGER_CREATE_SERVICE As Long = &H2
Private Const SC_MANAGER_ENUMERATE_SERVICE As Long = &H4
Private Const SC_MANAGER_LOCK As Long = &H8
Private Const SC_MANAGER_QUERY_LOCK_STATUS As Long = &H10
Private Const SC_MANAGER_MODIFY_BOOT_CONFIG As Long = &H20
Private Const SC_MANAGER_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)

Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long




Private Sub cbo_Change(Index As Integer)
    If Not Me.Visible Then Exit Sub
    
    If Index = 1 Then
        Call SetParChange(cbo, Index, mrsPar)
    Else
        
        Call SetParChange(cbo, Index, mrsPar, True, cbo(Index).Text)
    End If
End Sub

Private Sub ConfigAppNoState()
'------------------------------------------------
'���ܣ����ü�����ؿؼ��Ŀ�����
'��������
'���أ�
'------------------------------------------------
    Dim blnUseHisNo As Boolean  'ʹ��HIS�Ĳ���ID��ҽ��ID
    Dim blnCanOverWrite As Boolean  '����������ظ���
    Dim blnCheckMaxNo As Boolean    '����ȡʵ�������롱
    Dim blnChangeNo As Boolean      '�������ֹ��������š�
    
    blnCanOverWrite = True
    blnCheckMaxNo = True
    blnChangeNo = True
    
    '���ü���ѡ���һЩ�߼���ϵ
    '��1�����߼��ű��ֲ��䣬ͬʱ��Ҫ��ѡ�һҵ�����������ظ���
    '��2��ѡ���ˡ�ǰ׺�������ָ��������������ա��󣬽�ֹ�һҵ�����ȡʵ�������롱
    '��3�����ű��ֲ��䣬�����������ͳһ��������ʱ�Զ����桰���տ��ҵ���������Ӱ�����ͳһ��������ʱ�Զ����桰����Ӱ����������
    '��4��ѡ��ҽ��ID����������ID������ֹ�һҵ��������ֹ��������š�����ֹ�һҵ�����ȡʵ�������롱�������һҵ��������ظ�������ֹ���ü��ű���
    '��5��˳�����Զ��ѡ��
    '��6��ѡ���ˡ����������Զ��������򣬽�ֹѡ��Ӱ�����ǰ׺��
    
    If OptCode(1).value = True Then
        chkCanOverWrite.value = 1
        blnCanOverWrite = False
    End If
    
    If chkPreText.value = 1 Or chkDelimiter(1).value = 1 Or chkDelimiter(2).value = 1 Or chkYear.value = 1 _
        Or chkMonth.value = 1 Or chkDay.value = 1 Then
        chkCheckMaxNo.value = 0
        blnCheckMaxNo = False
    End If
    
    If (optUseAdviceID.value = True And OptCode(0).value = True) Or (optUsePatientID.value = True And OptCode(1).value = True) Then
        chkChangeNO.value = 0
        blnChangeNo = False
        chkCanOverWrite.value = 1
        blnCanOverWrite = False
        chkCheckMaxNo.value = 0
        blnCheckMaxNo = False
        blnUseHisNo = True
    Else
        blnUseHisNo = False
    End If
    
    chkNumber.value = 1
    
    If OptBuildcode(1).value = True And optPreText(0).value = True Then
        optPreText(0).value = False
        Call MsgBox("���Ű��ձ��������Զ���������ֹʹ��Ӱ�����ǰ׺��", vbOKOnly)
    End If
    
    chkPreText.Enabled = Not blnUseHisNo
    chkDelimiter(1).Enabled = Not blnUseHisNo
    chkDelimiter(2).Enabled = Not blnUseHisNo
    chkYear.Enabled = Not blnUseHisNo
    chkMonth.Enabled = Not blnUseHisNo
    chkDay.Enabled = Not blnUseHisNo
    chkChangeNO.Enabled = blnChangeNo
    chkCanOverWrite.Enabled = blnCanOverWrite
    chkCheckMaxNo.Enabled = blnCheckMaxNo
    chkNumber.Enabled = Not blnUseHisNo
    
    '���ð�ť�Ŀ�����
    '����һ����
    OptBuildcode(0).Enabled = OptCode(0).value
    OptBuildcode(1).Enabled = OptCode(0).value
    optUseAdviceID.Enabled = OptCode(0).value
    OptUnicode(0).Enabled = OptCode(1).value
    OptUnicode(1).Enabled = OptCode(1).value
    optUsePatientID.Enabled = OptCode(1).value
    
    'ǰ׺
    optPreText(0).Enabled = chkPreText.value And chkPreText.Enabled And OptBuildcode(1).value = False
    optPreText(1).Enabled = chkPreText.value And chkPreText.Enabled
    txtPreText.Enabled = (optPreText(1).value And optPreText(1).Enabled)
    
    '�ָ���
    cboDelimeter(1).Enabled = chkDelimiter(1).value And chkDelimiter(1).Enabled
    cboDelimeter(2).Enabled = chkDelimiter(2).value And chkDelimiter(2).Enabled
    
    '��
    optYear(0).Enabled = chkYear.value And chkYear.Enabled
    optYear(1).Enabled = chkYear.value And chkYear.Enabled
    
    '˳���--�̶�λ��
    chkFixedLen.Enabled = chkNumber.value And chkNumber.Enabled
    txtFixedLen.Enabled = chkFixedLen.value And chkFixedLen.Enabled
    txtStartNum.Enabled = chkNumber.value And chkNumber.Enabled
    lblStartNum.Enabled = chkNumber.value And chkNumber.Enabled
End Sub

Private Sub chkCanOverWrite_Click()
    Call ConfigAppNoState
End Sub

Private Sub chkCheckMaxNo_Click()
    Call ConfigAppNoState
End Sub



Private Sub chkDay_Click()
    Call ConfigAppNoState
End Sub

Private Sub chkDelimiter_Click(Index As Integer)
    Call ConfigAppNoState
End Sub

Private Sub chkFixedLen_Click()
    Call ConfigAppNoState
End Sub

Private Sub chkMonth_Click()
    Call ConfigAppNoState
End Sub

Private Sub chkNumber_Click()
    Call ConfigAppNoState
End Sub

Private Sub chkPreText_Click()
    Call ConfigAppNoState
End Sub

Private Sub chkPreView_Click()
    If chkPreView.value = 1 Then
        optMovePreview.Enabled = True
        lblDelayTime.Enabled = True
        txtDelayTime.Enabled = True
        optClickPreview.Enabled = True
    Else
        optMovePreview.Enabled = False
        lblDelayTime.Enabled = False
        txtDelayTime.Enabled = False
        optClickPreview.Enabled = False
    End If
End Sub


Private Sub chkYear_Click()
    Call ConfigAppNoState
End Sub

Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
    Dim blnValue As Boolean
    Dim strValue As String
    Dim i As Long
    Dim blNoPick As Boolean '����ѡ�û�й�ѡ

    If Not Me.Visible Then Exit Sub

    Select Case Index
    Case lst_PatholInfo
        blNoPick = True
        strValue = ""
        With lst(lst_PatholInfo)
            For i = 0 To .ListCount - 1
                If strValue <> "" Then strValue = strValue & ","
                
                If .Selected(i) Then
                    blNoPick = False
                    strValue = strValue & "1"
                Else
                    strValue = strValue & "0"
                End If
            Next
            
            '����ѡ�û�й�ѡ��ͬ�ڶ���ѡ
            If blNoPick Then
                strValue = "1,1,1,1,1,1,1,1,1,1"
            End If
        End With
        blnValue = True
    End Select
    Call SetParChange(lst, lst_PatholInfo, mrsPar, blnValue, strValue)
End Sub

Private Sub chkImageLevel_Click()
    txtImageLevel.Enabled = chkImageLevel.value = 1
End Sub

Private Sub ChkCompleteCommit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ChkCompleteCommit.value = 1 Then chkFinallyCompleteCommit.value = 0
End Sub

Private Sub chkFinallyCompleteCommit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkFinallyCompleteCommit.value = 1 Then ChkCompleteCommit.value = 0
End Sub

Private Sub ChkLike_Click()
    TxtLike.Enabled = IIF(ChkLike.value, True, False)
End Sub

Private Sub chkNameColColorCfg_Click()
    If chkNameColColorCfg.value = 1 Then
        chkOrdinaryNameColColorCfg.Enabled = True
    Else
        chkOrdinaryNameColColorCfg.value = 0
        chkOrdinaryNameColColorCfg.Enabled = False
    End If
End Sub

Private Sub chkPetitionCapture_Click()
    cboSaveDevice.Enabled = IIF(chkPetitionCapture.value, True, False)
End Sub

Private Sub chkRefreshInterval_Click()
    txtRefreshInterval.Enabled = IIF(chkRefreshInterval.value, True, False)
End Sub

Private Sub chkReportAfterResult_Click()
    If chkReportAfterResult.value = vbChecked Then
        chkIgnorePosi.Enabled = False
        chkIgnorePosi.value = vbUnchecked
    Else
        chkIgnorePosi.Enabled = True
    End If
End Sub

Private Sub chkReportLevel_Click()
    txtReportLevel.Enabled = chkReportLevel.value = 1
End Sub


Private Sub chkSpecialContent_Click()
    If chkSpecialContent.value = 1 Then
        cboSpecialContent.Enabled = True
    Else
        cboSpecialContent.Enabled = False
    End If
End Sub

Private Sub chkUseQueue_Click()
On Error GoTo ErrHandle
    optNumberRule(0).Enabled = chkUseQueue.value
    optNumberRule(1).Enabled = chkUseQueue.value
        
    'ufgGroupCfg.Enabled = chkUseQueue.value
    'ufgRoomCfg.Enabled = chkUseQueue.value
    'ufgStudyProCfg.Enabled = chkUseQueue.value
    
    'cmdAdd.Enabled = chkUseQueue.value
    'cmdDel.Enabled = chkUseQueue.value
    'cmdModify.Enabled = chkUseQueue.value
    'cmdStudyAcc.Enabled = chkUseQueue.value
    chkSynStudyList.Enabled = chkUseQueue.value
    
    txtValidDays.Enabled = chkUseQueue.value
    txtQueueReport.Enabled = chkUseQueue.value
    cbxPrintQueueNoWay.Enabled = chkUseQueue.value
    chkAutoInQueue.Enabled = chkUseQueue.value
    chkUseQueueMsg.Enabled = chkUseQueue.value
    
    lab(7).Enabled = chkUseQueue.value
    lab(8).Enabled = chkUseQueue.value
    lab(9).Enabled = chkUseQueue.value
    fra(18).Enabled = chkUseQueue.value
    
    'framGroup.Enabled = chkUseQueue.value
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmbDept_Click()
    mlng����ID = cmbDept.ItemData(cmbDept.ListIndex)
    
    'If stabWorkFlow.Tabs = IIF(InStr(GetPrivFunc(glngSys, 1160), "����") > 0, 6, 5) Then '�ж�tab������Ŀ����Ϊ��ȷ����װ����tab֮��Ŵ������е����
        
        Call Load���Ҳ���
        
    'End If
End Sub

Private Sub cmdAdd_Click()
    Me.lab(4).Tag = "": Me.txtName.Text = "": Me.txtName.Enabled = True
    Me.cmdDel.Enabled = True: Me.cmdSave.Enabled = True: Me.cmdRestore.Enabled = True: cboDevice.Enabled = True: If cboDevice.ListCount > 0 Then cboDevice.ListIndex = 0
    Me.txtName.SetFocus
End Sub

Private Sub cmdAddGroup_Click()
'����������Ϣ
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    Dim strGroupName As String
    Dim strPrefix As String
    Dim objFrmAdd As frmTechnicGroup
    Dim lngRow As Long
    
    '���÷�����Ӵ���
    Set objFrmAdd = New frmTechnicGroup
    If objFrmAdd.ShowGroupCfg(Me, mlng����ID, lngGroupId, strGroupName, strPrefix) Then
        lngRow = ufgGroupCfg.NewRow
    
        ufgGroupCfg.Text(lngRow, "ID") = lngGroupId
        ufgGroupCfg.Text(lngRow, "����") = strGroupName
        ufgGroupCfg.Text(lngRow, "����ǰ׺") = strPrefix
        
        '�������ִ�м�
        Call subLoadTechniRoom(lngGroupId)
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdApply_Click()
    '�������ܴ���
    If ValidateData() = False Then Exit Sub
    
    Call Save���Ҳ���
    
    If SavePar(mrsPar, Me) = False Then Exit Sub
End Sub

Private Sub cmdColor_Click(Index As Integer)
    dlgColor.Color = shpColor(Index).FillColor
    dlgColor.ShowColor
    shpColor(Index).FillColor = dlgColor.Color
End Sub

Private Sub cmdDefault_Click()
    Call subLoadListDefColorConfig
End Sub

Private Sub cmdDel_Click()
    Dim strSQL As String
    
    If MsgBox("���ɾ��ִ�м䡰" & Trim(Me.txtName.Text) & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
    Err = 0: On Error GoTo ErrHand
    
        strSQL = "zl_ҽ��ִ�з���_Delete(" & Val(mlng����ID) & ",'" & Trim(Me.txtName.Text) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Call subLoadRoomConfig
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdDelGroup_Click()
On Error GoTo ErrHandle
    Dim strSQL As String
    Dim lngGroupId As Long
    Dim lngMsgResult As Long
    
    If Not ufgGroupCfg.IsSelectionRow Then
        MsgBox "��ѡ����Ҫɾ���ķ������ݡ�", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    lngMsgResult = MsgBox("�Ƿ�ȷ��ɾ���÷�������? ɾ������齫���ɻָ���", vbYesNo, "��ʾ")
    If lngMsgResult = vbNo Then Exit Sub
    
    
    lngGroupId = ufgGroupCfg.KeyValue(ufgGroupCfg.SelectionRow)
    
    strSQL = "zl_Ӱ��ִ�з���_Del(" & lngGroupId & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "ɾ��ִ�з���")
    
    Call ufgRoomCfg.ClearListData
    Call ufgGroupCfg.DelCurRow(False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdModify_Click()
'�޸ķ�����Ϣ
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    Dim strGroupName As String
    Dim strPrefix As String
    Dim objFrmUpdate As frmTechnicGroup
    
    If Not ufgGroupCfg.IsSelectionRow Then
        MsgBox "��ѡ����Ҫ�޸ĵķ������ݡ�", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    lngGroupId = ufgGroupCfg.KeyValue(ufgGroupCfg.SelectionRow)
    strGroupName = ufgGroupCfg.Text(ufgGroupCfg.SelectionRow, "����")
    strPrefix = ufgGroupCfg.Text(ufgGroupCfg.SelectionRow, "����ǰ׺")
    
    '���÷�����´���
    Set objFrmUpdate = New frmTechnicGroup
    If objFrmUpdate.ShowGroupCfg(Me, mlng����ID, lngGroupId, strGroupName, strPrefix) Then
        ufgGroupCfg.CurText("����") = strGroupName
        ufgGroupCfg.CurText("����ǰ׺") = strPrefix
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdRestore_Click()
    Call subLoadRoomConfig
End Sub

Private Sub cmdSave_Click()
    Dim blnExist As Boolean, i As Integer
    Dim strSQL As String

    If Trim(Me.txtName.Text) = "" Then
        MsgBox "���Ʊ�������", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > Me.txtName.MaxLength Then
        MsgBox "���Ƴ���" & Me.txtName.MaxLength & "�ĳ�������", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    
    For i = 1 To lvwRoom.ListItems.Count
        If txtName.Text = lvwRoom.ListItems(i).Text Then blnExist = True: Exit For '�Ѿ�����
    Next
    '-----------------------------------------
    Err = 0: On Error GoTo ErrHand
    If Me.lab(4).Tag = "" And Not blnExist Then
        strSQL = "zl_ҽ��ִ�з���_Insert(" & Val(mlng����ID) & ",'" & Trim(Me.txtName.Text) & "','" & zlStr.NeedCode(cboDevice.Text) & "','" & txtNoPrefix.Text & "')"
    Else
        strSQL = "zl_ҽ��ִ�з���_Update(" & Val(mlng����ID) & ",'" & Trim(Me.lab(4).Tag) & "','" & Trim(Me.txtName.Text) & "','" & zlStr.NeedCode(cboDevice.Text) & "','" & txtNoPrefix.Text & "')"
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    MsgBox "ִ�м䱣��ɹ���", vbInformation, gstrSysName
    
    Call subLoadRoomConfig
    
    txtName.SetFocus
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdStudyAcc_Click()
'Ӱ������Ŀ��������
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    Dim objStudyAssocia As frmTechnicStudy
    
    If Not ufgGroupCfg.IsSelectionRow Then
        MsgBox "��ѡ����Ҫ���й����ķ������ݡ�", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    lngGroupId = ufgGroupCfg.KeyValue(ufgGroupCfg.SelectionRow)
    
    Set objStudyAssocia = New frmTechnicStudy
    If objStudyAssocia.ShowStudyAssociation(mlng����ID, lngGroupId, Me) Then
        Call ufgStudyProCfg.ClearListData
        Call subLoadStudyProAssociation(lngGroupId)
    End If
    
Exit Sub
ErrHandle:
If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdTestCon_Click()
On Error GoTo ErrHandle
    Call XWTestDBConnection(txt(6).Text, txt(7).Text, txt(8).Text)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub XWTestDBConnection(ByVal strServerName As String, ByVal strUser As String, ByVal strPwd As String)
'���ܣ� ��������SQLServer���ݿ�����
'������
'���أ��ɹ����ؿ��ַ�
'--------------------------------------------
    Dim cnTest As New ADODB.Connection

    If strServerName = "" Then
        MsgBox "δ�ҵ����ݿ������������Ϣ�������á�"
        Exit Sub
    End If
    
    On Error Resume Next
    Err = 0
    
    If cnTest.State = adStateOpen Then cnTest.Close
    
    Set cnTest = OraDataOpen(strServerName, strUser, strPwd)
    
    If Err <> 0 Or cnTest Is Nothing Then
        '���ݿ����Ӵ���
        MsgBox "���ݿ�����ʧ�ܡ�" & vbCrLf & vbCrLf & "��������ǣ�" & Err.Number & "�����������ǣ� " & Err.Description
        Exit Sub
    End If
    
    MsgBox "���ݿ����ӳɹ���"
    Exit Sub
Err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As ADODB.Connection
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    Dim cnOra As New ADODB.Connection
    
    On Error Resume Next
    Err = 0
    
    DoEvents
    
    With cnOra
        If .State = adStateOpen Then .Close
        
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            'OraDataOpen = Nothing
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo ErrHand
    
    Set OraDataOpen = cnOra
    
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function

Private Sub Form_Activate()
    If Me.Tag = "��ʼ�ɹ�" Then
        Call scbFunc_SelectedChanged(scbFunc.Selected)
        Me.Tag = ""
    End If
End Sub

Private Sub Form_Load()
    Dim strCategory As String
    
    mblnOk = False
    mlng����ID = 0
    mlngCur����ID = 0
    mstrCur���� = ""
    mstrCanUse���� = ""

    Call GetUserInfo
    
    mstrPrivs = gstrPrivs
    
    strCategory = "��������" ',��������
    
    'ͼ����,TaskPanelItem��ID(ͬʱҲ�ǲ�������Picture�ؼ������),TaskPanelItem�ı���;......
    marrFunc(0) = ""
    marrFunc(0) = marrFunc(0) & "102,0,Ӱ����������"
    marrFunc(0) = marrFunc(0) & ";101,1,Ӱ��ҽ������"
    marrFunc(0) = marrFunc(0) & ";103,2,Ӱ��ɼ�����"
    marrFunc(0) = marrFunc(0) & ";100,3,Ӱ��������"
    marrFunc(0) = marrFunc(0) & ";105,4,����鵵����"
    marrFunc(0) = marrFunc(0) & ";106,5,����軹����"
    
    'marrFunc(1) = "102,2,����ҩ������"

    '1.��ʼ���������һ�������б�,ȱʡѡ�е�һ��
    Call InitSCBItem(scbFunc, strCategory, picTPL.hwnd)
    Call scbFunc.Icons.AddIcons(imgType.Icons)
      
    '2.��ʼ���������Ķ��������б�,ȱʡѡ�е�һ��
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    
    '�ж��Ƿ�߱��Ŷӽк�Ȩ��
    If (InStr(GetPrivFunc(glngSys, 1160), "����") <= 0) Then
        stabWorkFlow.TabVisible(3) = False
    End If
    
    Call InitData
    Call ShowErrParasMsg(Me, mrsPar)
    Me.Tag = "��ʼ�ɹ�"
End Sub


Private Sub lvwRoom_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.txtName.Text = Item.Text
    Me.lab(4).Tag = Me.txtName.Text
    Me.txtNoPrefix.Text = Item.SubItems(2)
        
    SeekIndex cboDevice, Item.SubItems(1), True, , tNeedNo
End Sub

Private Sub OptBuildcode_Click(Index As Integer)
    Call ConfigAppNoState
End Sub

Private Sub OptCode_Click(Index As Integer)
    Call ConfigAppNoState
End Sub

Private Sub optMovePreview_Click()
    If optMovePreview.value = True Then
        txtDelayTime.Enabled = True
        lblDelayTime.Enabled = True
    End If
End Sub

Private Sub optPreText_Click(Index As Integer)
    Call ConfigAppNoState
End Sub

Private Sub optReportEditor_Click(Index As Integer)
    Dim hService As Long
    Dim hSCManager As Long

On Error GoTo ErrHandle

    fra(24).Visible = Index = 2
    
    Exit Sub
ErrHandle:
    
End Sub

Private Sub OptUnicode_Click(Index As Integer)
    Call ConfigAppNoState
End Sub

Private Sub optUseAdviceID_Click()
    Call ConfigAppNoState
End Sub

Private Sub optUsePatientID_Click()
    Call ConfigAppNoState
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Dim i As Long
    
    For i = 0 To picPar.UBound
        picPar(i).Visible = (i = Item.ID)
    Next
    
    'PACS���������в����ڰ����Ҳ�ѯ
    'lblLocate(txt_Dept).Visible = (Item.ID = GetFuncID("ҵ�����̿���", marrFunc))
    'txtLocate(txt_Dept).Visible = lblLocate(txt_Dept).Visible
    
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
    Dim objPic As PictureBox
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If picVbar.Left < 1500 Then picVbar.Left = 1500
    If picVbar.Left > Me.ScaleWidth - 3000 Then picVbar.Left = Me.ScaleWidth - 3000
    picVbar.Top = 0
    
    picFunc.Width = picVbar.Left + picVbar.Width
    
    For Each objPic In picPar
        objPic.Top = 0
        objPic.Left = 0
        objPic.Width = Me.ScaleWidth - objPic.Left
        objPic.Height = Me.ScaleHeight - PicBottom.ScaleHeight
    Next
End Sub


Private Sub scbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub picBottom_Resize()
    cmdApply.Left = PicBottom.ScaleWidth - cmdApply.Width - 120
    cmdCancel.Left = cmdApply.Left - cmdCancel.Width - 120
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


Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOk Then
        mrsPar.Filter = "(�޸�״̬=1 ANd ErrType =Null) OR  (�޸�״̬=1 And ErrType=" & PET_ֵ���� & ")"
        If mrsPar.RecordCount > 0 Then
            If MsgBox("�����޸Ĳ��ֲ����������������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    Set mrsPar = Nothing
End Sub

Private Sub InitData()
'���ܣ���ʼ������ؼ�,��ȡ����������
    
    '1.��ʼ������
    mlngPreFind = 1
    
    Call InitSystemPara
    
    
    '2.��ʼ������ؼ�
    Call InitEnv
    
    
    '3.����ϵͳ����
    Call LoadPar
    
    
    '���빤�����̲���
'    Call Load���Ҳ���
End Sub

Private Sub LoadPar()
'���ܣ���ȡ�����ز���������ؼ�
    Dim strValue As String, strTmp As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim arrObj As Variant  '�������ģ��1,������1,�ؼ�����1,ģ��2,������2,�ؼ�����2,......
    

    Set rsTmp = GetPar(mrsPar, pӰ���Ƭ���� & _
                            "," & pӰ��ҽ������ & _
                            "," & pӰ��ɼ����� & _
                            "," & pӰ�������� & _
                            "," & p����鵵���� & _
                            "," & p����軹����)

     '1.����CheckBox�����
    strTmp = pӰ��ҽ������ & ":¼����Ժ��Ϣ:" & chk_¼����Ժ��������Ϣ & _
            "," & pӰ��ɼ����� & ":¼����Ժ��Ϣ:" & chk_¼����Ժ���������Ϣ & _
            "," & pӰ�������� & ":������������:" & chk_���没������������ & _
            "," & pӰ�������� & ":������������:" & chk_������������������ & _
            "," & pӰ�������� & ":ϸ����������:" & chk_ϸ���������������� & _
            "," & pӰ�������� & ":������������:" & chk_���ﲡ������������ & _
            "," & pӰ�������� & ":ʬ����������:" & chk_ʬ�첡������������ & _
            "," & pӰ�������� & ":����ʯ����������:" & chk_��Ƭ�������������� & _
            "," & pӰ�������� & ":¼����Ժ��Ϣ:" & chk_¼����Ժ��Ϣ & _
            "," & p����軹���� & ":����ȷ�Ϻ��Զ���ӡ��ִ:" & chk_���ĺ��Զ���ӡ��ִ��

    Call SetParToControl(strTmp, mrsPar, chk)
    
    
    rsTmp.Filter = "ģ��=" & pӰ�������� & " and ������='¼����Ժ��Ϣ'"
    If rsTmp.RecordCount > 0 Then
        If Val(NVL(rsTmp!����ֵ)) = 1 Then
            txt(txt_������Ժ����).Enabled = True
        Else
            txt(txt_������Ժ����).Enabled = False
        End If
    End If
    

    '2.����ComboBox�����
    strTmp = pӰ���Ƭ���� & ":XW3D��Ƭ����:" & cbo_3D��Ƭ���� & _
            "," & p����鵵���� & ":������ǩ��������:" & cbo_������ǩ�������� & _
            "," & p����軹���� & ":���Ļ�ִ��������:" & cbo_���Ļ�ִ��������
    Call SetParToControl(strTmp, mrsPar, cbo, 3)
    
    strTmp = pӰ��ɼ����� & ":�ɼ�����ִ��ģʽ:" & cbo_�ɼ�����ִ��ģʽ
    Call SetParToControl(strTmp, mrsPar, cbo)
    
    strTmp = pӰ��ҽ������ & ":ҽ������ִ��ģʽ:" & cbo_ҽ������ִ��ģʽ
    Call SetParToControl(strTmp, mrsPar, cbo)
    
    strTmp = pӰ�������� & ":�������ִ��ģʽ:" & cbo_�������ִ��ģʽ
    Call SetParToControl(strTmp, mrsPar, cbo)
    
    
    
'
'    '3.����UpDown�����
'    strTmp = "0:5:" & ud_��¼ҽ��ʶ����
'    Call SetParToControl(strTmp, mrsPar, ud)     'mrsPar�洢�Ŀؼ�����txtUD
'
    '4.����TextBox�����
'
    '"," & pӰ���Ƭ���� & ":13:" & txt_����Ŀ¼ & _'
    
    strTmp = pӰ���Ƭ���� & ":XWɾ��ͼ���û���:" & txt_ͼ��ɾ���û����� & _
            "," & pӰ���Ƭ���� & ":XWɾ��ͼ������:" & txt_ͼ��ɾ���û����� & _
            "," & pӰ���Ƭ���� & ":XW����ͼ���û���:" & txt_ͼ�����û����� & _
            "," & pӰ���Ƭ���� & ":XW����ͼ������:" & txt_ͼ�����û����� & _
            "," & pӰ���Ƭ���� & ":XW���̿�¼�û���:" & txt_ͼ���¼�û����� & _
            "," & pӰ���Ƭ���� & ":XW���̿�¼����:" & txt_ͼ���¼�û����� & _
            "," & pӰ���Ƭ���� & ":XW���ݿ������IP:" & txt_������IP & _
            "," & pӰ���Ƭ���� & ":XW���ݿ�������û���:" & txt_�������û����� & _
            "," & pӰ���Ƭ���� & ":XW���ݿ����������:" & txt_�������û����� & _
            "," & pӰ���Ƭ���� & ":XWWEB��Ƭ��ַ:" & txt_WEB��Ƭ��ַ & _
            "," & pӰ���Ƭ���� & ":XW�ؼ�ͼ���ַ:" & txt_�ؼ�ͼ���ַ & _
            "," & pӰ���Ƭ���� & ":XWOracleӵ����:" & txt_��ӵ���� & _
            "," & pӰ���Ƭ���� & ":XW��鷽����:" & txt_��鷽���� & _
            "," & pӰ���Ƭ���� & ":XW���з�����:" & txt_���з����� & _
            "," & pӰ�������� & ":�޼�����ģ��:" & txt_�޼�����ģ�� & _
            "," & pӰ�������� & ":���汨��ģ��:" & txt_��������ģ�� & _
            "," & pӰ�������� & ":���߱���ģ��:" & txt_��������ģ�� & _
            "," & pӰ�������� & ":��Ⱦ����ģ��:" & txt_��Ⱦ����ģ�� & _
            "," & pӰ�������� & ":���ӱ���ģ��:" & txt_��������ģ�� & _
            "," & pӰ�������� & ":��Ժ��λ�ṹ����:" & txt_������Ժ���� & _
            "," & p����鵵���� & ":����Ĭ�ϲ�ѯ����:" & txt_����Ĭ�ϲ�ѯ���� & _
            "," & p����軹���� & ":����Ĭ�ϲ�ѯ����:" & txt_�軹Ĭ�ϲ�ѯ���� & _
            "," & pӰ���Ƭ���� & ":XWWeb����б��Ƭ��ַ:" & txt_����б��Ƭ��ַ


            
            
    Call SetParToControl(strTmp, mrsPar, txt)
    
    '����ʵ��
    rsTmp.Filter = "������='������ǩ��������'"
    If rsTmp.RecordCount > 0 Then
        cbo(2).Text = "" & NVL(rsTmp!����ֵ)
        Call SetParRelation(cbo, cbo_������ǩ��������, mrsPar, CStr(NVL(rsTmp!������)), p����鵵����)
    End If
    
    rsTmp.Filter = "������='���Ļ�ִ��������'"
    If rsTmp.RecordCount > 0 Then
        cbo(3).Text = "" & NVL(rsTmp!����ֵ)
        Call SetParRelation(cbo, cbo_���Ļ�ִ��������, mrsPar, CStr(NVL(rsTmp!������)), p����軹����)
    End If
    
    rsTmp.Filter = "������='��Ժ��λ�ṹ����'"
    If rsTmp.RecordCount > 0 Then
        txt(txt_������Ժ����).Text = "" & NVL(rsTmp!����ֵ)
        Call SetParRelation(txt, txt_������Ժ����, mrsPar, CStr(NVL(rsTmp!������)), pӰ��������)
    End If
    
    
    rsTmp.Filter = "������='ȡ����������'"
    If rsTmp.RecordCount > 0 Then
        strValue = "" & rsTmp!����ֵ
        With lst(lst_PatholInfo)
            For i = 0 To .ListCount - 1
                If Val(Split(strValue, ",")(i)) = 1 Then
                    .Selected(i) = True
                End If
            Next
        End With
        Call SetParRelation(lst, lst_PatholInfo, mrsPar, CStr(NVL(rsTmp!������)), pӰ��������)
    End If
    
    
            
            
        
    '�������ܴ���
    txt(1).Text = GetDecryptionPassW(txt(1).Text)
    txt(3).Text = GetDecryptionPassW(txt(3).Text)
    txt(5).Text = GetDecryptionPassW(txt(5).Text)
    txt(8).Text = GetDecryptionPassW(txt(8).Text)
'    rsTmp.Filter = "����=2"
'    While Not rsTmp.EOF
'
'        strValue = "" & rsTmp!����ֵ
'        Select Case rsTmp!������
'            Case "������ǩ��������"
'                cbo(2).Text = strValue
'                'Call SetParRelation(cbo, cbo_������ǩ��������, mrsPar)
'            Case "���Ļ�ִ��������"
'                cbo(3).Text = strValue
'                'Call SetParRelation(cbo, cbo_���Ļ�ִ��������, mrsPar)
'        End Select
'
'        rsTmp.MoveNext
'    Wend
'
'
    '5.����ListBox�����
'    strTmp = pסԺҽ���´� & ":4:" & lst_סԺ�����Ժ���
'    Call SetParToControl(strTmp, mrsPar, lst)
'
'    '6.����OptionButton�����
'    arrObj = Array(p����ҽ���´�, 45, opt����Ŀ������, _
'                    pסԺҽ���´�, 51, opt����Ŀ��סԺ)
'    Call SetParToControl("", mrsPar, arrObj)
'
'
'    '7.����ϵͳ����
'    rsTmp.Filter = "ģ��=0"
'    Do Until rsTmp.EOF
'        strValue = "" & rsTmp!����ֵ
'        Select Case rsTmp!������
'        Case 70
'            ud(ud_�����Ǽ���Ч����).value = IIF(Val(strValue) = 0, 1, Val(strValue))
'
'            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "") '����CheckBox�ؼ���������Ҫ�ٲ���һ����¼
'            Call SetParRelation(txtUD, ud_�����Ǽ���Ч����, mrsPar)
'
'        Case 233
'            Call Load��д��������(strValue)
'            Call SetParRelation(vsUnWriteDept, 0, mrsPar, rsTmp!������)
'        End Select
'
'        rsTmp.MoveNext
'    Loop
'
'    '8.����ģ���������
'    rsTmp.Filter = "ģ��=" & p����ҽ���´�
'    Do Until rsTmp.EOF
'        strValue = "" & rsTmp!����ֵ
'        Select Case rsTmp!������
'
'        End Select
'        rsTmp.MoveNext
'    Loop
'
End Sub

Private Sub InitEnv()
''���ܣ���ʼ������ؼ������ػ�������
On Error GoTo ErrHandle
    
    Call subInitTechincRoom
    
    Call LoadCheckNoDelimeter   '����Ҫ����subInitDepartInfoǰ�棬ȷ���ȳ�ʼ���ָ������ݣ��ٴ����ݿ��ȡ����
    
    Call subInitDepartInfo
    Call LoadPathol
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    '�������ܴ���
    If ValidateData() = False Then Exit Sub
    
    Call Save���Ҳ���
    
    If SavePar(mrsPar, Me) = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub

Private Function ValidateData() As Boolean
'���ܣ���֤���ݵ���Ч��
    Dim intTxtLen As Integer
    
    If txtImageLevel.Enabled Then
        '������״̬�µ� �����滻��Ӣ��״̬
        txtImageLevel.Text = Replace(txtImageLevel.Text, "��", ",")
        
        intTxtLen = Len(txtImageLevel.Text) - Len(Replace(txtImageLevel.Text, ",", ""))
        
        If intTxtLen > 3 Or intTxtLen < 1 Then
            MsgBox "Ӱ��ȼ�����Ϊ2�֣����Ϊ4�֣���������д��", vbOKOnly, "��ʾ��Ϣ"
            txtImageLevel.Text = NVL(GetDeptPara(mlng����ID, "Ӱ�������ȼ�", "��,��"))
            txtImageLevel.SetFocus
            Exit Function
        End If
    End If
    
    
    If txtReportLevel.Enabled Then
        '������״̬�µ� �����滻��Ӣ��״̬
        txtReportLevel.Text = Replace(txtReportLevel.Text, "��", ",")
        
        intTxtLen = Len(txtReportLevel.Text) - Len(Replace(txtReportLevel.Text, ",", ""))
        
        If intTxtLen > 3 Or intTxtLen < 1 Then
            MsgBox "����ȼ�����Ϊ2�֣����Ϊ4�֣���������д��", vbOKOnly, "��ʾ��Ϣ"
            txtReportLevel.Text = NVL(GetDeptPara(mlng����ID, "���������ȼ�", "��,��"))
            txtReportLevel.SetFocus
            Exit Function
        End If
    End If
    
    ValidateData = True
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub txtAudit_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtCheckIn_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtDelayTime_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtEnreg_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtFixedLen_Change()
    If Val(txtFixedLen.Text) > 18 Then
        MsgBox "�̶�λ�����Ϊ18λ����������д��", vbOKOnly, "��ʾ��Ϣ"
        txtFixedLen.Text = ""
    End If
End Sub

Private Sub txtFixedLen_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub


Private Sub TxtLike_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

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
        End Select
    End If
End Sub

Private Sub lst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(lst, Index, mrsPar)
End Sub

Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(chk, Index, mrsPar)
End Sub

Private Sub txt_Change(Index As Integer)
On Error GoTo ErrHandle
    If Me.Visible = False Then Exit Sub
    
    If Index = 1 Or Index = 3 Or Index = 5 Or Index = 8 Then
        Call SetParChange(txt, Index, mrsPar, True, GetEncryptionPassW(txt(Index).Text))
    Else
        Call SetParChange(txt, Index, mrsPar)
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = Asc(gstrParSplit1) Or KeyAscii = Asc(gstrParSplit2) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, Index, mrsPar)
End Sub


Private Sub cbo_GotFocus(Index As Integer)
    Call SetParTip(cbo, Index, mrsPar)
End Sub



Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub




Private Sub cbo_Click(Index As Integer)
    Dim blnValue As Boolean, strValue As String

    If Not Me.Visible Then Exit Sub
    
    Call SetParChange(cbo, Index, mrsPar)
    
'    Select Case Index
'    Case cbo_�ɼ�����ִ��ģʽ   '���������б���
'        Call SetParChange(cbo, Index, mrsPar)
'    Case Else       '���ı����ݽ��б���
'        Call SetParChange(cbo, Index, mrsPar)
'    End Select
'
'    If Me.Visible Then
'        Call SetParChange(cbo, Index, mrsPar, blnValue, strValue)
'    End If
    
End Sub

Private Sub chk_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(chk, Index, mrsPar)   'Call SetParChange(chk, Index, mrsPar, blnValue, strValue)
    End If
    
    If Index = chk_¼����Ժ��Ϣ Then
        txt(txt_������Ժ����).Enabled = chk(chk_¼����Ժ��Ϣ).value
    End If
End Sub


Private Sub txtName_GotFocus()
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Or KeyAscii = 95 Then KeyAscii = 0
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtPreText_Change()
    Dim lngSet As Long
    
    '�ж��Ƿ����Сд�ַ����������������ʾ�û������Զ��ĳɴ�д
    If txtPreText.Text <> UCase(txtPreText.Text) Then
        lngSet = txtPreText.SelStart
        txtPreText.Text = UCase(txtPreText.Text)
        txtPreText.SelStart = lngSet
    End If
End Sub

Private Sub txtRefreshInterval_Change()
    If Val(txtRefreshInterval.Text) > 600 Then
        txtRefreshInterval.Text = 600
    End If
End Sub

Private Sub txtRefreshInterval_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub


Private Sub txtRefreshInterval_LostFocus()
    If Val(txtRefreshInterval.Text) < 10 Then
        txtRefreshInterval.Text = 10
    End If
End Sub

Private Sub txtReport_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtStartNum_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtStudy_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtValidDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtViewHistoryImageDays_Change()
    If Val(txtViewHistoryImageDays.Text) > 15 Then
        txtViewHistoryImageDays.Text = 15
    End If
End Sub

Private Sub txtViewHistoryImageDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtViewHistoryImageDays_LostFocus()
    If Val(txtViewHistoryImageDays.Text) <= 0 Then
        txtViewHistoryImageDays.Text = 1
    End If
End Sub

Private Sub TxtĬ������_Change()
    If Val(TxtĬ������.Text) > 15 Then
        TxtĬ������.Text = 15
    End If
End Sub

Private Sub TxtĬ������_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub TxtĬ������_LostFocus()
    If Val(TxtĬ������.Text) <= 0 Then
        TxtĬ������.Text = 1
    End If
End Sub


Private Sub ufgGroupCfg_OnSelChange()
On Error GoTo ErrHandle
    Dim lngGroupId As Long
    lngGroupId = Val(ufgGroupCfg.CurKeyValue)
    
    '����ҽ��ִ�з���
    Call subLoadTechniRoom(lngGroupId)
    
    '�����������Ŀ����
    Call subLoadStudyProAssociation(lngGroupId)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgRoomCfg_OnDblClick()
'˫��ִ�м�ʱ�����з����޸Ĵ���
On Error GoTo ErrHandle
    Call cmdModify_Click
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgStudyProCfg_OnDblClick()
'˫��Ӱ������Ŀʱ�����й������ô���
On Error GoTo ErrHandle
    Call cmdStudyAcc_Click
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub Load���Ҳ���()

    Call subLoadWorkFlowConfig      '��ȡ�������̲���
    Call subLoadRoomConfig          '��ȡִ�м����
    Call subLoadInputConfig         '��ȡ¼�������ò���
    
    If stabWorkFlow.TabVisible(3) = True Then
        Call subLoadQueueGroupConfig    '��ȡ���з������ò���
    End If
    
    Call subLoadSpecifyReportItemName
    Call subLoadReportConfig        '��ȡ����༭�����ò���
    Call subLoadListColorConfig     '��ȡ�б���ɫ���ò���
End Sub


Private Sub Save���Ҳ���()
'����ҽ�����ݶ���
    On Error GoTo ErrHandle
        Call subSaveWorkFlowConfig
        Call subSaveInputConfig
        
        If stabWorkFlow.TabVisible(3) = True Then
            Call subSaveQueueGroupConfig
        End If
        
        Call subSaveReportConfig
        Call subSaveListColorConfig
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


'************************************************************************************************************************************
'************************************************************************************************************************************
    Private Sub subSaveListColorConfig()
        Dim i As Integer, strInput As String
        Dim strSQL As String
        
        If mlng����ID < 0 Then Exit Sub
        
          
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '�ѵǼ�','" & shpColor(8).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '�ѱ���','" & shpColor(1).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '������','" & shpColor(2).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '�Ѽ��','" & shpColor(0).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '������','" & shpColor(3).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '�ѱ���','" & shpColor(4).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '�����','" & shpColor(6).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '�����','" & shpColor(7).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '�����','" & shpColor(5).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '�Ѿܾ�','" & shpColor(9).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '�Ѳ���','" & shpColor(10).FillColor & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '�ǼǺ�����','" & Val(txtEnreg.Text) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '����������','" & Val(txtCheckIn.Text) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '��������','" & Val(txtStudy.Text) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '���������','" & Val(txtReport.Text) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '��˺�����','" & Val(txtAudit.Text) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '������ɫ����','" & chkNameColColorCfg.value & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", 'ȱʡ���Ͳ���������ɫ����','" & chkOrdinaryNameColColorCfg.value & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '��ɫ��ʾ����','" & IIF(optListColorMark(0).value = True, 0, 1) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End Sub
    
    Private Sub subSaveReportConfig()
        Dim intMatch As Integer
        Dim strSQL As String
        
        On Error GoTo ErrHand
        
        If mlng����ID < 0 Then Exit Sub
        
        
        If optReportEditor(0).value = True Then         '���Ӳ����༭��
            intMatch = 0
        ElseIf optReportEditor(1).value = True Then     'PACS����༭��
            intMatch = 1
        ElseIf optReportEditor(2).value = True Then     '�����ĵ��༭��
            intMatch = 2
        End If
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '����༭��','" & intMatch & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '��ʾ����ͼ��','" & chkShowImage.value & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '��������ͼ����','" & txtMinImageCount.Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '��ʾ��Ƶ�ɼ�','" & chkShowVideoCapture.value & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '��ӡ���˳�','" & chkExitAfterPrint.value & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '��ʾר�Ʊ���','" & chkSpecialContent.value & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", 'ר�Ʊ���ҳ','" & cboSpecialContent.Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        If optWordDblClick(0).value = True Then         '����ʾ�˫����ֱ��д�뱨��
            intMatch = 0
        ElseIf optWordDblClick(1).value = True Then     '����ʾ�˫����򿪱༭����
            intMatch = 1
        End If
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '����ʾ�˫������','" & intMatch & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        If optImageDblClick(0).value = True Then         '����ͼ˫����ֱ��д�뱨��
            intMatch = 0
        ElseIf optImageDblClick(1).value = True Then     '����ͼ˫�����ͼ��༭����
            intMatch = 1
        End If
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '����ͼ˫������','" & intMatch & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '�����������','" & txtCheckView.Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '����������','" & txtResult.Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '��������','" & txtAdvice.Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        If optShowWord(0).value = True Then         'ֱ����ʾ�ʾ�ʾ��
            intMatch = 0
        ElseIf optShowWord(1).value = True Then     '˫���������ʾ�ʾ�ʾ��
            intMatch = 1
        End If
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '��ʾ�ʾ�ʾ��','" & intMatch & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '��˴�ӡ���������','" & chkUntreadPrinted.value & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        If optReportEditor(2) Then
            strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '�鿴��ʷ����','" & IIF(optHistoryReportEditor(0).value, 0, 1) & "')"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '��ӡ��ʽѡ��ʽ','" & IIF(optPrintFormat(0).value, 0, 1) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '��ѡ�����ʽ','" & IIF(chkPrintFormat.value, 1, 0) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        Exit Sub
ErrHand:
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End Sub

    Public Sub subSaveQueueGroupConfig()
    '�������ò���
        If mlng����ID < 0 Then Exit Sub
    
        SetDeptPara mlng����ID, "�����Ŷӽк�", chkUseQueue.value
        SetDeptPara mlng����ID, "�Ŷӽкű������", IIF(optNumberRule(0).value, 0, 1)
        SetDeptPara mlng����ID, "�Ŷ����ݱ�������", Val(txtValidDays.Text)
        SetDeptPara mlng����ID, "�Ŷӵ�������", txtQueueReport.Text
        SetDeptPara mlng����ID, "ͬ����λ����б�", chkSynStudyList.value
        SetDeptPara mlng����ID, "����ʱ����Ĭ��ִ�м�", chkSelectRoom.value
        SetDeptPara mlng����ID, "�Ŷӵ���ӡ��ʽ", cbxPrintQueueNoWay.ListIndex
        SetDeptPara mlng����ID, "�����Ŷ���Ϣ����", chkUseQueueMsg.value
        SetDeptPara mlng����ID, "�������Զ��Ŷ�", chkAutoInQueue.value
    End Sub
    
    
    Private Function SetDeptPara(ByVal lngDeptID As Long, ByVal varPara As String, ByVal strValue As String) As Boolean
    '���ܣ�����ָ���Ĳ���ֵ
    '������lngDept=����ID
    '      varPara=������
    '      strValue=������ֵ
    '���أ������Ƿ�ɹ�
        Dim strSQL As String
        
        On Error GoTo errH
            
        strSQL = "ZL_Ӱ�����̲���_UPDATE(" & lngDeptID & ",'" & varPara & "','" & strValue & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "SetPara")
        
        '���óɹ����������
        Set mrsDeptParas = Nothing
        
        SetDeptPara = True
        Exit Function
errH:
        If ErrCenter() = 1 Then Resume
    End Function


    Private Sub subSaveInputConfig()
        Call subSaveInputItem(0)
        Call subSaveInputItem(1)
    End Sub
    
    
    Private Sub subSaveInputItem(intType As Integer)
        Dim i As Integer, strInput As String
        Dim strSQL As String
        
        strInput = ""
        If intType = 0 Then
            For i = 0 To ChkMouseMove.UBound
                If ChkMouseMove(i).value = 1 Then strInput = strInput & "|" & ChkMouseMove(i).Caption
            Next
        Else
            For i = 0 To ChkInput.UBound
                If ChkInput(i).value = 1 Then strInput = strInput & "|" & ChkInput(i).Caption
            Next
        End If
        
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlng����ID & ", '" & IIF(intType = 0, "�������", "��¼����") & "','" & strInput & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End Sub
    
    
    Private Sub subSaveWorkFlowConfig()
        Dim strTemp As String
        Dim lngTemp As Long
        
        On Error GoTo ErrHand
    
        SetDeptPara mlng����ID, "�������뵥ɨ��", chkPetitionCapture.value        '�������뵥ɨ�� ��������
        
        SetDeptPara mlng����ID, "��������ж�", chkConformDetermine.value         '��������ж� ��������
'        SetDeptPara mlng����ID, "Σ������ж�", chkCriticalValues.value           'Σ������ж� ��������
        
        SetDeptPara mlng����ID, "���Խ��������", chkIgnorePosi.value
        SetDeptPara mlng����ID, "��Ӱ�����Ϊ����", chkReportAfterResult.value
        SetDeptPara mlng����ID, "��Ͻ��Ĭ������", chkDefaultPosi.value   '��Ͻ��Ĭ������ ��������
        
        SetDeptPara mlng����ID, "Ӱ�������ж�", chkImageLevel.value           'Ӱ�������ж� ��������
        SetDeptPara mlng����ID, "Ӱ�������ȼ�", txtImageLevel.Text            'ͼ�������ȼ� ��������
        SetDeptPara mlng����ID, "���������ж�", chkReportLevel.value           '���������ж� ��������
        SetDeptPara mlng����ID, "���������ȼ�", txtReportLevel.Text           '���������ȼ� ��������
        
        SetDeptPara mlng����ID, "��Ͻ����ʾ����", IIF(optResultInput(0).value = True, 0, IIF(optResultInput(1).value = True, 1, 2))
        
        SetDeptPara mlng����ID, "�ޱ�����ɺ�ֱ�����", ChkFinishCommit.value
        SetDeptPara mlng����ID, "��ͼ��ҽ��վ���ɹ�Ƭ", chkCanViewImage.value     '��ͼ��ҽ��վ���ɹ�Ƭ
        SetDeptPara mlng����ID, "��ͼ�����д����", chkReportAfterImging.value
        
        '��������
        SetDeptPara mlng����ID, "���߼��ű��ֲ���", IIF(OptCode(1).value, 1, 0)
        SetDeptPara mlng����ID, "���ű��ֲ������", IIF(OptUnicode(1).value, 1, 0)
        SetDeptPara mlng����ID, "�ֹ���������", chkChangeNO.value
        SetDeptPara mlng����ID, "��������ظ�", chkCanOverWrite.value
        SetDeptPara mlng����ID, "��ȡʵ��������", chkCheckMaxNo.value
        
        SetDeptPara mlng����ID, "ʹ�û��ߺ�", IIF(optUsePatientID.value And optUsePatientID.Enabled, 1, 0)
        SetDeptPara mlng����ID, "ʹ��ҽ����", IIF(optUseAdviceID.value And optUseAdviceID.Enabled, 1, 0)
        
        If OptCode(0).value = True Then
            SetDeptPara mlng����ID, "�������ɷ�ʽ", IIF(OptBuildcode(1).value, 1, 0)
        Else
            SetDeptPara mlng����ID, "�������ɷ�ʽ", IIF(OptUnicode(1).value, 1, 0)
        End If
        
        If chkPreText.value = 1 Then
            If optPreText(0).value = True Then
                strTemp = 1
            Else
                strTemp = txtPreText.Text
            End If
        Else
            strTemp = ""
        End If
        SetDeptPara mlng����ID, "����ǰ׺", strTemp
        SetDeptPara mlng����ID, "���ŷָ���1", IIF(chkDelimiter(1).value = 1, Left(cboDelimeter(1).Text, 1), "")
        SetDeptPara mlng����ID, "���ŷָ���2", IIF(chkDelimiter(2).value = 1, Left(cboDelimeter(2).Text, 1), "")
        SetDeptPara mlng����ID, "������", IIF(chkYear.value = 1, IIF(optYear(0).value = True, 1, 2), 0)
        SetDeptPara mlng����ID, "������", chkMonth.value
        SetDeptPara mlng����ID, "������", chkDay.value
        SetDeptPara mlng����ID, "������ʼ��", IIF(Val(txtStartNum.Text) = 0, 1, Val(txtStartNum.Text))
        SetDeptPara mlng����ID, "���Ź̶�λ��", IIF(chkFixedLen.value = 1, Val(txtFixedLen.Text), 0)
        
        SetDeptPara mlng����ID, "��λƬ����", chkLocalizerBackward.value
        SetDeptPara mlng����ID, "�������û�", chkChangeUser.value
        SetDeptPara mlng����ID, "�����л��û�", chkSwitchUser.value
        SetDeptPara mlng����ID, "ֻ����д�Լ����ı���", chkTechReportSame.value
        SetDeptPara mlng����ID, "�ɼ�ͼ����Ϊ��鼼ʦ", chkWriteCapDoctor.value
        SetDeptPara mlng����ID, "��˺�ֱ�����", ChkCompleteCommit.value
        SetDeptPara mlng����ID, "�����ֱ�����", chkFinallyCompleteCommit.value
        SetDeptPara mlng����ID, "��ӡ��ֱ�����", chkPrintCommit.value
        SetDeptPara mlng����ID, "�����ֱ�Ӵ�ӡ", chkCompletePrint.value
        SetDeptPara mlng����ID, "�ǼǺ�ֱ�Ӽ��", chkSample.value
        SetDeptPara mlng����ID, "ƥ�����ݿ���Ŀ", IIF(optMatch(0).value, 0, IIF(optMatch(1), 1, 2))
        
        SetDeptPara mlng����ID, "�Ǽ�ʱ����ģ����������", IIF(ChkLike.value = 1, Abs(Val(TxtLike.Text)), 0)
        SetDeptPara mlng����ID, "���еǼǲ��˱��Ϊ����", chkAllPatientIsOutside
        
        If Val(TxtĬ������.Text) > 15 Or Val(TxtĬ������.Text) <= 0 Then
            TxtĬ������.Text = 2
        End If
        SetDeptPara mlng����ID, "Ĭ�Ϲ�������", Val(TxtĬ������.Text)
        
        If Val(txtViewHistoryImageDays.Text) > 15 Or Val(txtViewHistoryImageDays.Text) <= 0 Then
            txtViewHistoryImageDays.Text = 1
        End If
        SetDeptPara mlng����ID, "�Զ�����ʷͼ������", Val(txtViewHistoryImageDays.Text)
        
        
        SetDeptPara mlng����ID, "������������", chkUseReferencePatient.value
        SetDeptPara mlng����ID, "ƽ������˲��ܴ򱨸�", chkPrintNeedComplete.value
        
        SetDeptPara mlng����ID, "ƴ������Сд", IIF(optCapital(0).value, 0, IIF(optCapital(1), 1, 2))
        SetDeptPara mlng����ID, "ƴ�����ָ���", IIF(optSplitter(0).value, 0, 1)
        
        If cboSaveDevice.Text <> "" Then
            SetDeptPara mlng����ID, "���뵥�洢�豸��", Split(cboSaveDevice.Text, "-")(0)
        Else
            SetDeptPara mlng����ID, "���뵥�洢�豸��", ""
        End If
        
        If Abs(Val(txtRefreshInterval.Text)) = 0 Or Abs(Val(txtRefreshInterval.Text)) > 65 Then
            txtRefreshInterval.Text = 10
        End If
        SetDeptPara mlng����ID, "�Զ�ˢ�¼��", IIF(chkRefreshInterval.value = 1, Abs(Val(txtRefreshInterval.Text)), 0)
        SetDeptPara mlng����ID, "����ʱ�Զ�����WorkList", chkAutoSendWorkList.value
        SetDeptPara mlng����ID, "��ʾ��������", chkAddons.value
        SetDeptPara mlng����ID, "��ʾ��Ӱ��", chkReagent.value
        SetDeptPara mlng����ID, "ҽ��վ�鿴����", cboViewReport.ListIndex
        SetDeptPara mlng����ID, "����л�ʱ��λ����༭", chkSetFocusWithReport.value
        SetDeptPara mlng����ID, "����Ĭ��ģ����ѯ", chkNameFuzzySearch.value
        SetDeptPara mlng����ID, "������ѯʱ������", chkNameQueryTimeLimit.value
        
        If chkPreView.value = 1 Then
            If optMovePreview.value Then
                lngTemp = 1
            ElseIf optClickPreview.value Then
                lngTemp = 2
            End If
        Else
            lngTemp = 0
        End If
        
        SetDeptPara mlng����ID, "����ͼԤ����ʽ", lngTemp
        SetDeptPara mlng����ID, "�ƶ�Ԥ����ʱ", Val(txtDelayTime.Text)
         
        Exit Sub
ErrHand:
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End Sub
'************************************************************************************************************************************
'************************************************************************************************************************************
    
    Private Sub subLoadSpecifyReportItemName()
        'װ��ר�Ʊ�������
        Call cboSpecialContent.Clear
        Call cboSpecialContent.AddItem(Report_Form_frmReportES)
        Call cboSpecialContent.AddItem(Report_Form_frmReportPathology)
        Call cboSpecialContent.AddItem(Report_Form_frmReportUS)
        Call cboSpecialContent.AddItem(Report_Form_frmReportCustom)
    End Sub
    
    
    Private Sub subLoadListDefColorConfig()
    '�����б�Ĭ����ɫ����
        shpColor(10).FillColor = ColorConstants.vbYellow
        shpColor(9).FillColor = ColorConstants.vbRed
        shpColor(7).FillColor = ColorConstants.vbGreen
        
        shpColor(0).FillColor = ColorConstants.vbWhite
        shpColor(1).FillColor = ColorConstants.vbWhite
        shpColor(2).FillColor = ColorConstants.vbWhite
        shpColor(3).FillColor = ColorConstants.vbWhite
        shpColor(4).FillColor = ColorConstants.vbWhite
        shpColor(5).FillColor = ColorConstants.vbWhite
        shpColor(6).FillColor = ColorConstants.vbWhite
        shpColor(8).FillColor = ColorConstants.vbWhite
        
        txtEnreg.Text = "0"
        txtCheckIn.Text = "0"
        txtStudy.Text = "0"
        txtReport.Text = "0"
        txtAudit.Text = "0"
    End Sub
    
    Private Sub subLoadListColorConfig()
        Dim strSQL As String
        Dim rsTemp As ADODB.Recordset
        Dim lngTemp As Long
                 
        On Error GoTo Err
        
        
        Call subLoadListDefColorConfig
        
        strSQL = "select ID ,����ID,������,����ֵ from Ӱ�����̲��� where ����ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        
        While Not rsTemp.EOF
            Select Case rsTemp!������
                Case "�ѵǼ�"
                    shpColor(8).FillColor = Val(NVL(rsTemp!����ֵ))
                Case "�ѱ���"
                    shpColor(1).FillColor = Val(NVL(rsTemp!����ֵ))
                Case "������"
                    shpColor(2).FillColor = Val(NVL(rsTemp!����ֵ))
                Case "�Ѽ��"
                    shpColor(0).FillColor = Val(NVL(rsTemp!����ֵ))
                Case "������"
                    shpColor(3).FillColor = Val(NVL(rsTemp!����ֵ))
                Case "�ѱ���"
                    shpColor(4).FillColor = Val(NVL(rsTemp!����ֵ))
                Case "�����"
                    shpColor(6).FillColor = Val(NVL(rsTemp!����ֵ))
                Case "�����"
                    shpColor(7).FillColor = Val(NVL(rsTemp!����ֵ))
                Case "�����"
                    shpColor(5).FillColor = Val(NVL(rsTemp!����ֵ))
                Case "�Ѿܾ�"
                    shpColor(9).FillColor = Val(NVL(rsTemp!����ֵ))
                Case "�Ѳ���"
                    shpColor(10).FillColor = Val(NVL(rsTemp!����ֵ))
                Case "�ǼǺ�����"
                    txtEnreg.Text = Val(NVL(rsTemp!����ֵ))
                Case "����������"
                    txtCheckIn.Text = Val(NVL(rsTemp!����ֵ))
                Case "��������"
                    txtStudy.Text = Val(NVL(rsTemp!����ֵ))
                Case "���������"
                    txtReport.Text = Val(NVL(rsTemp!����ֵ))
                Case "��˺�����"
                    txtAudit.Text = Val(NVL(rsTemp!����ֵ))
                Case "��ɫ��ʾ����"
                    If Val(NVL(rsTemp!����ֵ)) = 0 Then
                        optListColorMark(0).value = True
                    Else
                        optListColorMark(1).value = True
                    End If
            End Select
            rsTemp.MoveNext
        Wend
        
        chkNameColColorCfg.value = Val(GetDeptPara(mlng����ID, "������ɫ����", 0))
        If chkNameColColorCfg.value = 0 Then
            chkOrdinaryNameColColorCfg.value = 0
            chkOrdinaryNameColColorCfg.Enabled = False
        Else
            chkOrdinaryNameColColorCfg.Enabled = True
            chkOrdinaryNameColColorCfg.value = Val(GetDeptPara(mlng����ID, "ȱʡ���Ͳ���������ɫ����", 0))
        End If
    
        
        Exit Sub
Err:
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End Sub
    

    Public Sub subLoadReportConfig()
        Dim strSQL As String
        Dim rsTemp As ADODB.Recordset
        Dim lngTemp As Long
        
        optReportEditor(0).value = True 'Ĭ��ʹ�õ��Ӳ����༭���༭����
        chkShowImage.value = 0          'Ĭ�ϲ���ʾͼ������
        chkShowVideoCapture.value = 0   'Ĭ�ϲ���ʾ��Ƶ�ɼ�����
        
        chkSpecialContent.value = 0     'Ĭ�ϲ���ʾר�Ʊ���
        cboSpecialContent.Enabled = False
        chkExitAfterPrint.value = 0     'Ĭ�ϴ�ӡ���˳�
        optWordDblClick(0).value = True 'Ĭ��˫���ʾ��ֱ��д�뱨��
        optImageDblClick(0).value = True 'Ĭ�ϱ�������ͼ˫����ֱ��д�뱨��
        txtCheckView.Text = "�������"  'Ĭ��Ϊ�������
        txtResult.Text = "������"     'Ĭ��Ϊ������
        txtAdvice.Text = "����"         'Ĭ��Ϊ����
        optShowWord(0).value = True     'Ĭ��Ϊֱ����ʾ�ʾ�ģ��
        chkUntreadPrinted.value = 0     'Ĭ��Ϊ��˴�ӡ���������
         
        On Error GoTo Err
        strSQL = "select ID ,����ID,������,����ֵ from Ӱ�����̲��� where ����ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        
        While Not rsTemp.EOF
            Select Case rsTemp!������
                Case "����༭��"
                    If NVL(rsTemp!����ֵ, 0) = 0 Then
                        optReportEditor(0).value = True
                    ElseIf NVL(rsTemp!����ֵ, 0) = 1 Then
                        optReportEditor(1).value = True
                    Else
                        optReportEditor(2).value = True
                    End If
                Case "�鿴��ʷ����"
                    If NVL(rsTemp!����ֵ, 0) = 0 Then
                        optHistoryReportEditor(0).value = True
                    Else
                        optHistoryReportEditor(1).value = True
                    End If
                Case "��ʾ����ͼ��"
                    chkShowImage.value = NVL(rsTemp!����ֵ, 0)
                Case "��������ͼ����"
                    txtMinImageCount.Text = NVL(rsTemp!����ֵ, "8")
                Case "��ʾ��Ƶ�ɼ�"
                    chkShowVideoCapture.value = NVL(rsTemp!����ֵ, 0)
                Case "��ӡ���˳�"
                    chkExitAfterPrint.value = NVL(rsTemp!����ֵ, 0)
                Case "��ʾר�Ʊ���"
                    chkSpecialContent.value = NVL(rsTemp!����ֵ, 0)
                    cboSpecialContent.Enabled = IIF(chkSpecialContent.value = 1, True, False)
                Case "ר�Ʊ���ҳ"
                    cboSpecialContent.Text = NVL(rsTemp!����ֵ)
                Case "����ʾ�˫������"
                    If NVL(rsTemp!����ֵ, 0) = 0 Then
                        optWordDblClick(0).value = True
                    Else
                        optWordDblClick(1).value = True
                    End If
                Case "����ͼ˫������"
                    If NVL(rsTemp!����ֵ, 0) = 0 Then
                        optImageDblClick(0).value = True
                    Else
                        optImageDblClick(1).value = True
                    End If
                Case "�����������"
                    txtCheckView.Text = NVL(rsTemp!����ֵ, "�������")
                Case "����������"
                    txtResult.Text = NVL(rsTemp!����ֵ, "������")
                Case "��������"
                    txtAdvice.Text = NVL(rsTemp!����ֵ, "����")
                Case "��ʾ�ʾ�ʾ��"
                    If NVL(rsTemp!����ֵ, 0) = 0 Then
                        optShowWord(0).value = True
                    Else
                        optShowWord(1).value = True
                    End If
                Case "��˴�ӡ���������"
                    chkUntreadPrinted.value = NVL(rsTemp!����ֵ, 0)
                Case "��ӡ��ʽѡ��ʽ"
                If NVL(rsTemp!����ֵ, 0) = 0 Then
                    optPrintFormat(0).value = True
                Else
                    optPrintFormat(1).value = True
                End If
                Case "��ѡ�����ʽ"
                    chkPrintFormat.value = IIF(NVL(rsTemp!����ֵ, 0), 1, 0)
            End Select
            rsTemp.MoveNext
        Wend
        
        If optReportEditor(2).value Then
            fra(24).Visible = True
        Else
            fra(24).Visible = False
        End If
        
        Exit Sub
Err:
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End Sub

    Public Sub subLoadQueueGroupConfig()
    'ˢ�����ò���
        Dim strSQL As String
        Dim rsTemp As ADODB.Recordset
        Dim lngIndex As Long
    
        On Error GoTo Err
    
        lngIndex = Val(GetDeptPara(mlng����ID, "�Ŷӽкű������", 0))
        txtValidDays.Text = GetDeptPara(mlng����ID, "�Ŷ����ݱ�������", 1)
        txtQueueReport.Text = GetDeptPara(mlng����ID, "�Ŷӵ�������", "")
        chkSynStudyList.value = Val(GetDeptPara(mlng����ID, "ͬ����λ����б�", 0))
        chkSelectRoom.value = Val(GetDeptPara(mlng����ID, "����ʱ����Ĭ��ִ�м�", 0))
        chkUseQueueMsg.value = Val(GetDeptPara(mlng����ID, "�����Ŷ���Ϣ����", 1))
        chkAutoInQueue.value = Val(GetDeptPara(mlng����ID, "�������Զ��Ŷ�", 1))
        
        '0-����ӡ��1-�Զ���ӡ��2-��ʾ��ӡ
        cbxPrintQueueNoWay.ListIndex = Val(GetDeptPara(mlng����ID, "�Ŷӵ���ӡ��ʽ", 0))
        
        chkUseQueue.value = Val(GetDeptPara(mlng����ID, "�����Ŷӽк�", 0))
        
        Call subLoadGroupInf
    
        optNumberRule(lngIndex).value = True
    
        Call chkUseQueue_Click
    
        Exit Sub
Err:
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End Sub

    Private Sub subLoadGroupInf()
    '����ҽ��������Ϣ
        Dim strSQL As String
        Dim rsData As ADODB.Recordset
        
        strSQL = "select Id, ����,����ǰ׺ from Ӱ��ִ�з��� where ����ID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ������Ϣ", mlng����ID)
        
        Call ufgGroupCfg.ClearListData
        If rsData.RecordCount <= 0 Then Exit Sub
        
        rsData.Sort = "���� asc"
        
        Set ufgGroupCfg.AdoData = rsData
        Call ufgGroupCfg.BindData
    End Sub
    
    Private Sub subLoadTechniRoom(ByVal lngGroupId As Long)
    '�������������ҽ��ִ�з���
        Dim strSQL As String
        Dim rsData As ADODB.Recordset
        
        strSQL = "select ִ�м�, ����ǰ׺ from ҽ��ִ�з��� where ����Id=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯҽ��ִ�з���", lngGroupId)
        
        Call ufgRoomCfg.ClearListData
        If rsData.RecordCount <= 0 Then Exit Sub
        
        rsData.Sort = "ִ�м� asc"
        
        Set ufgRoomCfg.AdoData = rsData
        Call ufgRoomCfg.BindData
    End Sub
    
    Private Sub subLoadStudyProAssociation(ByVal lngGroupId As Long)
    '��������Ŀ����
        Dim strSQL As String
        Dim rsData As ADODB.Recordset
        
        strSQL = "select ����,���� from ������ĿĿ¼ a, Ӱ�������� b where a.id=b.������ĿId and b.����Id=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯӰ�������������Ŀ", lngGroupId)
        
        Call ufgStudyProCfg.ClearListData
        If rsData.RecordCount <= 0 Then Exit Sub
        
        rsData.Sort = "����"
        
        Set ufgStudyProCfg.AdoData = rsData
        Call ufgStudyProCfg.BindData
    End Sub

    Private Sub subLoadInputConfig()
        Call subLoadInputItem(0)
        Call subLoadInputItem(1)
    End Sub
    
    Private Sub subLoadInputItem(intType As Integer)
    '����¼������
    'intType 0-������ƣ�1-��¼����
        Dim i As Integer, strInput As String, j As Integer
        Dim strSQL As String
        Dim rsTemp As ADODB.Recordset
        
        
        If intType = 0 Then
            '��ʼ���ر��ƶ�ѡ���
            For i = 0 To ChkMouseMove.UBound
                ChkMouseMove(i).value = 0
            Next
        Else
            '��ʼ��¼ѡ���
            For i = 0 To ChkInput.UBound
                ChkInput(i).value = 0
            Next
        End If
        
        strSQL = "select ID ,����ID,����ֵ from Ӱ�����̲��� where ����ID = [1] and ������ = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, CStr(IIF(intType = 0, "�������", "��¼����")))
        
        If Not rsTemp.EOF Then
            strInput = NVL(rsTemp!����ֵ)
            For i = 0 To UBound(Split(strInput, "|"))
                If intType = 0 Then
                    For j = 0 To ChkMouseMove.UBound
                        If ChkMouseMove(j).Caption = Split(strInput, "|")(i) Then ChkMouseMove(j).value = 1: Exit For
                    Next
                
                Else
                    For j = 0 To ChkInput.UBound
                        If ChkInput(j).Caption = Split(strInput, "|")(i) Then ChkInput(j).value = 1: Exit For
                    Next
                End If
            Next
        End If
    End Sub


    Private Sub SeekIndex(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean, Optional blnPreserve As Boolean = False, Optional intIsSearchNo As TNeedType = tNeedName)
    '���ܣ���ComboBox�в��Ҳ���λ
    '������blnEvent=��λʱ�Ƿ񴥷�Click�¼�
          'blnPreserve--����Ҳ���ƥ����Ŀ���򱣳�ԭ����Ŀ
          'intIsSearchNo -- 0:ͨ�����붨λ,1:ͨ�����ֶ�λ,2:�ù���������ֶ�λ
    '˵����δ�ܶ�λʱ,����ListIndex=-1
    '       Cbo.SeekIndex���ܱȽϼ򵥣�����index��ᴥ���¼������ʺ�ʹ��
        Dim i As Long
    
        For i = 0 To objCbo.ListCount - 1
            If IIF(Abs(intIsSearchNo) = tNeedAll, objCbo.List(i), IIF(Abs(intIsSearchNo) = tNeedNo, zlStr.NeedCode(objCbo.List(i)), zlStr.NeedName(objCbo.List(i)))) = strText Then
                If blnEvent Then
                    objCbo.ListIndex = i
                Else
                    Call zlControl.CboSetIndex(objCbo.hwnd, i)
                End If
                Exit Sub
            End If
        Next
        
        If blnPreserve = True Then
            If blnEvent = False Then
                Call zlControl.CboSetIndex(objCbo.hwnd, objCbo.ListIndex)
            End If
        Else
            If blnEvent Then
                objCbo.ListIndex = -1
            Else
                Call zlControl.CboSetIndex(objCbo.hwnd, -1)
            End If
        End If
        
    End Sub


    Private Sub subLoadRoomConfig()
        Dim ObjItem As ListItem
        Dim rsTemp As New ADODB.Recordset
        Dim strSQL As String
        
        On Error GoTo ErrHand
        
        strSQL = "Select ִ�м�,����豸,����ǰ׺ From ҽ��ִ�з��� where ����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(Val(mlng����ID)))
        Me.lvwRoom.ListItems.Clear
        With rsTemp
            Do While Not .EOF
                Set ObjItem = Me.lvwRoom.ListItems.Add(, , !ִ�м�, 1, 1)
                
                ObjItem.SubItems(1) = NVL(!����豸)
                ObjItem.SubItems(2) = NVL(!����ǰ׺)
                .MoveNext
            Loop
        End With
        
        Err = 0: On Error Resume Next
        If Me.lvwRoom.ListItems.Count > 0 Then
            Me.lvwRoom.ListItems(1).Selected = True
            Me.lvwRoom.SelectedItem.EnsureVisible
        End If
        
        Err = 0: On Error GoTo 0
        If Me.lvwRoom.ListItems.Count > 0 Then
            Call lvwRoom_ItemClick(Me.lvwRoom.SelectedItem)
            Me.txtName.Enabled = True: cboDevice.Enabled = True
            Me.cmdDel.Enabled = True: Me.cmdSave.Enabled = True: Me.cmdRestore.Enabled = True
        Else
            Me.lab(4).Tag = "": Me.txtName.Text = "": If cboDevice.ListCount > 0 Then cboDevice.ListIndex = 0
            Me.txtName.Enabled = False: cboDevice.Enabled = False
            Me.cmdDel.Enabled = False: Me.cmdSave.Enabled = False: Me.cmdRestore.Enabled = False
        End If
        Exit Sub
ErrHand:
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End Sub


    Private Sub subInitTechincRoom()
        Dim rsTemp As New ADODB.Recordset
        Dim strSQL As String
        
        Me.lvwRoom.ListItems.Clear
        With Me.lvwRoom.ColumnHeaders
            .Clear
            .Add , "����", "����", 3000
            .Add , "����豸", "����豸", 3000
            .Add , "����ǰ׺", "����ǰ׺", 2000
        End With
        
        strSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ״̬=1 and ����=4"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
        cboDevice.Clear
        Do Until rsTemp.EOF
            cboDevice.AddItem rsTemp!�豸�� & "-" & rsTemp!�豸��
            rsTemp.MoveNext
        Loop
    End Sub
    
    
    Private Sub subLoadWorkFlowConfig()
        Dim rsTemp As ADODB.Recordset
        Dim lngTemp As Long
        Dim strTemp As String
            
        '��ʼ��Ĭ��ֵ,Ӧ����һ��ͳһ�ĵط�����Ĭ��ֵ������������ʾ�����ն�ȡ
        chkIgnorePosi.value = 0     '���Խ��������
        chkReportAfterResult.value = 0 '��Ӱ�����Ϊ����
        ChkFinishCommit.value = 0   '�ޱ�����ɺ�ֱ�����
        chkReportAfterImging.value = 0  '��ͼ�񲻿ɱ༭����
        chkLocalizerBackward.value = 0  '��λƬ����
        chkChangeUser.value = 0         '�������û�
        chkSwitchUser.value = 0         '�����л��û�
        chkTechReportSame.value = 0     'ֻ����д�Լ����ı���
        chkWriteCapDoctor.value = 0     '�ɼ�ͼ����Ϊ��鼼ʦ
        ChkCompleteCommit.value = 0     '��˺�ֱ�����
        chkFinallyCompleteCommit.value = 0  '�����ֱ�����
        optMatch(0).value = True        'ƥ�����ݿ���Ŀ
        
        ChkLike.value = 0               '���õǼ�ʱ����ģ������
        TxtLike.Text = 0                '�Ǽ�ʱ����ģ����������
        TxtĬ������.Text = 2            'Ĭ�Ϲ�������
        txtViewHistoryImageDays.Text = 1 'Ĭ���Զ�����ʷͼ������
        chkRefreshInterval.value = 0    '���ò����б��Զ�ˢ��
        txtRefreshInterval.Text = 0     'Ĭ�ϲ����б��Զ�ˢ�¼��Ϊ0�룬��ˢ��
        cboSaveDevice.Clear                 '�洢�豸
        chkPrintCommit.value = 0        '��ӡ��ֱ�����
        chkCompletePrint.value = 0      '�����ֱ�Ӵ�ӡ
        chkUseReferencePatient.value = 0  'Ĭ�ϲ����ù�������
        optCapital(0).value = True      'Ĭ��ƴ��ʹ�ô�д
        optCapital(1).value = True      'Ĭ��ƴ������ÿո�
        chkCheckMaxNo.value = 1         'Ĭ����ȡʵ��������
        chkDefaultPosi.value = 0        '��Ͻ��Ĭ������Ϊδ��ѡ
        chkConformDetermine.value = 1       '��������ж�Ĭ��Ϊѡ��
        txtImageLevel.Text = "��,��"     'Ĭ��Ӱ�������ȼ�
        txtReportLevel.Text = "��,��"    'Ĭ�ϱ��������ȼ�
        chkPetitionCapture.value = 1     'Ĭ�Ϲ�ѡ�������뵥ɨ��
        chkAddons.value = 1              '�ڵǼǴ�����ʾ��������
        chkReagent.value = 1             '�ڵǼǴ�����ʾ��Ӱ��

        If cboViewReport.ListCount > 0 Then cboViewReport.ListIndex = 0
        
        On Error GoTo Err
        
        lngTemp = Val(GetDeptPara(mlng����ID, "��Ͻ����ʾ����", 0))
        optResultInput(lngTemp).value = True
        
        chkIgnorePosi.value = Val(GetDeptPara(mlng����ID, "���Խ��������", 0)) '��һ��ʹ��ʱ��Ҫ���¶�ȡ
        chkDefaultPosi.value = Val(GetDeptPara(mlng����ID, "��Ͻ��Ĭ������", 0))  '��ȡĬ�����Բ���
        chkReportAfterResult.value = Val(GetDeptPara(mlng����ID, "��Ӱ�����Ϊ����", 0))
        
        chkConformDetermine.value = Val(GetDeptPara(mlng����ID, "��������ж�", 0))    '��ȡ��������ж�
        
        chkImageLevel.value = Val(GetDeptPara(mlng����ID, "Ӱ�������ж�", 0))   '��ȡӰ�������ж�
        txtImageLevel.Text = NVL(GetDeptPara(mlng����ID, "Ӱ�������ȼ�", "��,��"))  '��ȡӰ�������ȼ�
        txtImageLevel.Enabled = chkImageLevel.value = 1
        
        chkReportLevel.value = Val(GetDeptPara(mlng����ID, "���������ж�", 0)) '��ȡ���������ж�
        txtReportLevel.Text = NVL(GetDeptPara(mlng����ID, "���������ȼ�", "��,��"))  '��ȡ���������ȼ�
        txtReportLevel.Enabled = chkReportLevel.value = 1
        
        chkPetitionCapture.value = Val(GetDeptPara(mlng����ID, "�������뵥ɨ��", 1))    '��ȡ�������뵥ɨ�����
    
        ChkFinishCommit.value = Val(GetDeptPara(mlng����ID, "�ޱ�����ɺ�ֱ�����", 0))
        chkCanViewImage.value = Val(GetDeptPara(mlng����ID, "��ͼ��ҽ��վ���ɹ�Ƭ", 0))
        chkReportAfterImging.value = Val(GetDeptPara(mlng����ID, "��ͼ�����д����", 0))
        chkCanOverWrite.value = Val(GetDeptPara(mlng����ID, "��������ظ�", 0))
        chkCheckMaxNo.value = Val(GetDeptPara(mlng����ID, "��ȡʵ��������", 1))
        chkChangeNO.value = Val(GetDeptPara(mlng����ID, "�ֹ���������", 0))
        chkLocalizerBackward.value = Val(GetDeptPara(mlng����ID, "��λƬ����", 0))
        chkChangeUser.value = Val(GetDeptPara(mlng����ID, "�������û�", 0))
        chkSwitchUser.value = Val(GetDeptPara(mlng����ID, "�����л��û�", 0))
        chkTechReportSame.value = Val(GetDeptPara(mlng����ID, "ֻ����д�Լ����ı���", 0))
        chkWriteCapDoctor.value = Val(GetDeptPara(mlng����ID, "�ɼ�ͼ����Ϊ��鼼ʦ", 0))
        ChkCompleteCommit.value = Val(GetDeptPara(mlng����ID, "��˺�ֱ�����", 0))
        chkFinallyCompleteCommit.value = Val(GetDeptPara(mlng����ID, "�����ֱ�����", 0))
        chkPrintCommit.value = Val(GetDeptPara(mlng����ID, "��ӡ��ֱ�����", 0))
        chkCompletePrint.value = Val(GetDeptPara(mlng����ID, "�����ֱ�Ӵ�ӡ", 0))
        
        TxtLike.Text = Val(GetDeptPara(mlng����ID, "�Ǽ�ʱ����ģ����������", 0))
        chkSample.value = Val(GetDeptPara(mlng����ID, "�ǼǺ�ֱ�Ӽ��", 0))
        ChkLike.value = IIF(Val(TxtLike.Text) <> 0, 1, 0)
        chkAllPatientIsOutside.value = Val(GetDeptPara(mlng����ID, "���еǼǲ��˱��Ϊ����", 0))
        
        TxtĬ������.Text = Val(GetDeptPara(mlng����ID, "Ĭ�Ϲ�������", 2))
        
        If Val(TxtĬ������.Text) > 15 Or Val(TxtĬ������.Text) <= 0 Then
            TxtĬ������.Text = 2
        End If
        
        txtViewHistoryImageDays.Text = Val(GetDeptPara(mlng����ID, "�Զ�����ʷͼ������", 1))
        If Val(txtViewHistoryImageDays.Text) > 15 Or Val(txtViewHistoryImageDays.Text) <= 0 Then
            txtViewHistoryImageDays.Text = 1
        End If
        
        txtRefreshInterval.Text = Val(GetDeptPara(mlng����ID, "�Զ�ˢ�¼��", 0))
        chkRefreshInterval.value = IIF(Val(txtRefreshInterval.Text) <> 0, 1, 0)
        optMatch(Val(GetDeptPara(mlng����ID, "ƥ�����ݿ���Ŀ", 0))).value = True
        
        chkAutoSendWorkList.value = Val(GetDeptPara(mlng����ID, "����ʱ�Զ�����WorkList", "1"))
        chkAddons.value = Val(GetDeptPara(mlng����ID, "��ʾ��������", "1"))
        chkReagent.value = Val(GetDeptPara(mlng����ID, "��ʾ��Ӱ��", "1"))
        chkSetFocusWithReport.value = Val(GetDeptPara(mlng����ID, "����л�ʱ��λ����༭", "1"))
        chkNameFuzzySearch.value = Val(GetDeptPara(mlng����ID, "����Ĭ��ģ����ѯ", "1"))
        chkNameQueryTimeLimit.value = Val(GetDeptPara(mlng����ID, "������ѯʱ������", "1"))
        
        chkPreView.value = IIF(Val(GetDeptPara(mlng����ID, "����ͼԤ����ʽ", "0")) > 0, 1, 0)
        
        If chkPreView.value = 1 Then
            optMovePreview.Enabled = True
            lblDelayTime.Enabled = True
            txtDelayTime.Enabled = True
            optClickPreview.Enabled = True
        Else
            optMovePreview.Enabled = False
            lblDelayTime.Enabled = False
            txtDelayTime.Enabled = False
            optClickPreview.Enabled = False
        End If
        
        optMovePreview.value = Val(GetDeptPara(mlng����ID, "����ͼԤ����ʽ", "0")) = 1
        optClickPreview.value = Val(GetDeptPara(mlng����ID, "����ͼԤ����ʽ", "0")) = 2
        txtDelayTime.Text = Val(GetDeptPara(mlng����ID, "�ƶ�Ԥ����ʱ", "2"))
        
        If Val(GetDeptPara(mlng����ID, "ҽ��վ�鿴����", "1")) = 0 Then
            cboViewReport.ListIndex = 0
        Else
            cboViewReport.ListIndex = 1
        End If
        
        OptCode(Val(GetDeptPara(mlng����ID, "���߼��ű��ֲ���", 0))).value = True
        OptUnicode(Val(GetDeptPara(mlng����ID, "���ű��ֲ������", 0))).value = True
        optUsePatientID.value = Val(GetDeptPara(mlng����ID, "ʹ�û��ߺ�", 0))
        OptBuildcode(Val(GetDeptPara(mlng����ID, "�������ɷ�ʽ", 0))).value = True
        optUseAdviceID.value = Val(GetDeptPara(mlng����ID, "ʹ��ҽ����", 0))
        
        '���ű������
        strTemp = GetDeptPara(mlng����ID, "����ǰ׺", "")
        If strTemp = "" Then
            '��ʹ��ǰ׺
            chkPreText.value = 0
        Else
            'ʹ��ǰ׺
            chkPreText.value = 1
            If strTemp = "1" Then
                optPreText(0).value = 1
                txtPreText.Text = ""
            Else
                optPreText(1).value = 1
                txtPreText.Text = strTemp
            End If
        End If
        
        strTemp = GetDeptPara(mlng����ID, "���ŷָ���1", "")
        strTemp = Left(strTemp, 1) 'ֻȡһ���ַ�
        Call setCheckNoDelimeter(1, strTemp)
        
        strTemp = GetDeptPara(mlng����ID, "���ŷָ���2", "")
        strTemp = Left(strTemp, 1) 'ֻȡһ���ַ�
        Call setCheckNoDelimeter(2, strTemp)
        
        lngTemp = Val(GetDeptPara(mlng����ID, "������", 0))
        chkYear.value = IIF(lngTemp = 0, 0, 1)
        optYear(0).value = IIF((lngTemp = 1 Or lngTemp = 0), 1, 0)
        optYear(1).value = IIF(lngTemp = 2, 1, 0)
        
        chkMonth.value = IIF(Val(GetDeptPara(mlng����ID, "������", 0)) = 1, 1, 0)
        chkDay.value = IIF(Val(GetDeptPara(mlng����ID, "������", 0)) = 1, 1, 0)
        
        txtStartNum.Text = Val(GetDeptPara(mlng����ID, "������ʼ��", 1))
        lngTemp = Val(GetDeptPara(mlng����ID, "���Ź̶�λ��", 0))
        chkFixedLen.value = IIF(lngTemp = 0, 0, 1)
        txtFixedLen.Text = IIF(lngTemp = 0, "", lngTemp)
        
        '���ü������õĲ�����ؿ�����
        Call ConfigAppNoState
        
        chkUseReferencePatient.value = Val(GetDeptPara(mlng����ID, "������������", 0))
        chkPrintNeedComplete.value = Val(GetDeptPara(mlng����ID, "ƽ������˲��ܴ򱨸�", 0))
        
        'ƴ��������
        optCapital(Val(GetDeptPara(mlng����ID, "ƴ������Сд", 0))).value = True
        optSplitter(Val(GetDeptPara(mlng����ID, "ƴ�����ָ���", 0))).value = True
        
        Call LoadScanDevice
        
        Exit Sub
Err:
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End Sub

    
    Private Function GetDeptPara(ByVal lngDeptID As Long, ByVal varPara As String, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
    '���ܣ���ȡָ���Ĳ���ֵ
    '������lngDept=����ID
    '      varPara=������
    '      strDefault=�����ݿ���û�иò���ʱʹ�õ�ȱʡֵ(ע�ⲻ��Ϊ��ʱ)
    '      blnNotCache=�Ƿ񲻴ӻ����ж�ȡ
    '���أ�����ֵ���ַ�����ʽ
        Dim rsTmp As ADODB.Recordset
        Dim strSQL As String, blnNew As Boolean
        
        On Error GoTo errH
        
        If blnNotCache Then
            Set rsTmp = New ADODB.Recordset
            strSQL = "Select ����ֵ from Ӱ�����̲��� where ����ID = [1] and ������=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����", lngDeptID, varPara)
            
            If Not rsTmp.EOF Then
                GetDeptPara = NVL(rsTmp!����ֵ)
            Else
                GetDeptPara = strDefault
            End If
        Else
            '��һ�μ��ز�������
            If mrsDeptParas Is Nothing Then
                blnNew = True
            ElseIf mrsDeptParas.State = 0 Then
                blnNew = True
            End If
            If blnNew Then
                strSQL = "Select ����ֵ,������,����ID from Ӱ�����̲���"
                Set mrsDeptParas = New ADODB.Recordset
                Set mrsDeptParas = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����")
            End If
            
            '���ݻ����ȡ����ֵ
            mrsDeptParas.Filter = "������='" & CStr(varPara) & "' AND ����ID=" & lngDeptID
            If Not mrsDeptParas.EOF Then
                GetDeptPara = NVL(mrsDeptParas!����ֵ)
            Else
                GetDeptPara = strDefault
            End If
        End If
        Exit Function
errH:
        If ErrCenter() = 1 Then Resume
    End Function

    Private Sub LoadScanDevice()
        Dim strSQL As String
        Dim rsTemp As ADODB.Recordset
        
        strSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=1 and NVL(״̬,0)=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTemp.EOF Then
            MsgBox "δ�������뵥�洢�豸���뵽Ӱ���豸Ŀ¼�����ã�", vbInformation, gstrSysName
            Exit Sub
        Else
            cboSaveDevice.AddItem ""
            
            Do While Not rsTemp.EOF
                cboSaveDevice.AddItem rsTemp!�豸�� & "-" & NVL(rsTemp!�豸��)
                
                If GetDeptPara(mlng����ID, "���뵥�洢�豸��", "") = rsTemp!�豸�� Then
                    cboSaveDevice.ListIndex = cboSaveDevice.NewIndex
                End If
                
                rsTemp.MoveNext
            Loop
        End If
    End Sub

    Private Function GetUserInfo() As Boolean
    '���ܣ���ȡ��½�û���Ϣ
        Dim rsTmp As New ADODB.Recordset
        Dim strSQL As String
        
        Set rsTmp = zlDatabase.GetUserInfo
        
        UserInfo.�û��� = gstrDbUser
        UserInfo.���� = gstrDbUser
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.��� = rsTmp!���
            UserInfo.����ID = IIF(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
            UserInfo.���� = IIF(IsNull(rsTmp!����), "", rsTmp!����)
            UserInfo.���� = IIF(IsNull(rsTmp!����), "", rsTmp!����)
            UserInfo.�û��� = IIF(IsNull(rsTmp!�û���), "", rsTmp!�û���)
            GetUserInfo = True
        End If
        Exit Function
errH:
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End Function
    

    Private Function GetUser����IDs(Optional ByVal bln���� As Boolean) As String
    '���ܣ���ȡ����Ա�����Ŀ���(�������ڿ���+�������������Ŀ���),�����ж��
    '�������Ƿ�ȡ���������µĿ���
        Dim rsTmp As New ADODB.Recordset
        Dim strSQL As String, i As Long
        
        strSQL = "Select ����ID From ������Ա Where ��ԱID=[1]"
        If bln���� Then
            strSQL = strSQL & " Union" & _
                " Select Distinct B.����ID From ������Ա A,��λ״����¼ B" & _
                " Where A.����ID=B.����ID And A.��ԱID=[1]"
        End If
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
        For i = 1 To rsTmp.RecordCount
            GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
            rsTmp.MoveNext
        Next
        GetUser����IDs = Mid(GetUser����IDs, 2)
        Exit Function
errH:
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End Function


    Private Function subInitDepartInfo()
    '���������
        Dim rsTmp As New ADODB.Recordset
        Dim strSQL As String, i As Long
        Dim str����IDs As String, str��Դ As String
        Dim strDepartment() As String
        Dim intCurDept As Integer
        
        On Error GoTo errH
        
        If InStr(mstrPrivs, "���п���") > 0 Then
            strSQL = _
                " Select Distinct A.ID,A.����,A.����" & _
                " From ���ű� A,��������˵�� B " & _
                " Where B.����ID = A.ID " & _
                " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
                " And B.�������� IN('���')  Order by A.����"
        Else
            strSQL = _
                " Select Distinct A.ID,A.����,A.����" & _
                " From ���ű� A,��������˵�� B,������Ա C " & _
                " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=" & UserInfo.ID & _
                " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
                " And B.�������� IN('���')  Order by A.����"
        End If
         
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
        If rsTmp.EOF Then
            MsgBox "û�з���ҽ��������Ϣ,���ȵ����Ź��������á�", vbInformation, gstrSysName
            Exit Function
        Else
            str����IDs = GetUser����IDs
            Do Until rsTmp.EOF
                mstrCanUse���� = mstrCanUse���� & "|" & rsTmp!ID & "_" & rsTmp!���� & "-" & rsTmp!����
                If rsTmp!ID = UserInfo.����ID Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� '��ȡĬ�Ͽ���
                If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur����ID = 0 Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� 'û��Ĭ�Ͽ���,ȡ���������ҵ�һ��
                rsTmp.MoveNext
            Loop
            
            str����IDs = GetUser����IDs
            Do Until rsTmp.EOF
                mstrCanUse���� = mstrCanUse���� & "|" & rsTmp!ID & "_" & rsTmp!���� & "-" & rsTmp!����
                If rsTmp!ID = UserInfo.����ID Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� '��ȡĬ�Ͽ���
                If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur����ID = 0 Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� 'û��Ĭ�Ͽ���,ȡ���������ҵ�һ��
                rsTmp.MoveNext
            Loop
            mstrCanUse���� = Mid(mstrCanUse����, 2)
            If InStr(mstrPrivs, "���п���") > 0 And mlngCur����ID = 0 Then
                mlngCur����ID = Split(Split(mstrCanUse����, "|")(0), "_")(0)
                mstrCur���� = Split(Split(mstrCanUse����, "|")(0), "_")(1)
            End If
            
            If mlngCur����ID = 0 And InStr(mstrPrivs, "���п���") <= 0 Then 'û�����п��Ҳ���Ȩ��,���Ҳ����߿��Ҳ����ڼ�������
                MsgBox "û�з�������������,����ʹ��ҽ������վ��", vbInformation, gstrSysName
                Exit Function
            End If
            
            '���cmbDept
            cmbDept.Clear
            intCurDept = -1
            strDepartment = Split(mstrCanUse����, "|")
            For i = 0 To UBound(strDepartment)
                cmbDept.AddItem Split(strDepartment(i), "_")(1)
                cmbDept.ItemData(cmbDept.ListCount - 1) = Split(strDepartment(i), "_")(0)
                If Split(strDepartment(i), "_")(0) = mlngCur����ID Then
                    intCurDept = i
                End If
            Next i
            If intCurDept <> -1 Then
                cmbDept.ListIndex = intCurDept
            Else
                cmbDept.ListIndex = 0
            End If
            mlng����ID = cmbDept.ItemData(cmbDept.ListIndex)
            
            subInitDepartInfo = True
        End If
        
        
        Exit Function
errH:
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End Function

Private Function GetEncryptionPassW(ByVal strPswd As String) As String
'��ȡ��������
    Dim strEncryptionPassW As String
    
    If strPswd = "" Then Exit Function
    
    strEncryptionPassW = EncryptionPassW(Trim(strPswd))
    strEncryptionPassW = Mid(strEncryptionPassW, 1, 1) & "��" & Mid(strEncryptionPassW, 2)
    strEncryptionPassW = "��" & strEncryptionPassW & "��"
    strEncryptionPassW = Replace(strEncryptionPassW, "'", "''")
    
    GetEncryptionPassW = strEncryptionPassW
End Function

Private Function GetDecryptionPassW(ByVal strPswd As String) As String
'��ȡ��������
    Dim strDecryptionPassW As String
    
    GetDecryptionPassW = strPswd
    
    If Len(strPswd) >= 3 Then
        If Mid(strPswd, 1, 1) & Mid(strPswd, 3, 1) & Mid(strPswd, Len(strPswd), 1) = "�����" Then
            strDecryptionPassW = Mid(strPswd, 2)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, Len(strDecryptionPassW) - 1)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, 1) & Mid(strDecryptionPassW, 3)
            strDecryptionPassW = DecryptionPassW(strDecryptionPassW)
            
            GetDecryptionPassW = strDecryptionPassW
        End If
    End If
End Function

Private Function GetRandom(ByVal lngBase As Long) As String
    Dim lngNum As Long
    
    Randomize 99
    
    lngNum = Fix(Rnd * lngBase)
    
    If lngNum <= 0 Then lngNum = 1
    
    GetRandom = Chr(lngNum)
End Function

'��ȡ��������
Private Function EncryptionPassW(ByVal strPassW As String) As String
    Dim i As Integer
    Dim lngAsc  As Long
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim strRandom As String
    Dim strBase As String
        
    i = 0
    
    lngPassWLength = Len(strPassW)
    
    strBase = GetRandom(30)
    strRandom = GetRandom(30)
    
    ReDim intASC(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
     
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassW, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strBase) Xor Asc(strRandom)
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop
    
    EncryptionPassW = strBase & Join(strTemp, "") & strRandom '���ܺ���ִ�
End Function

'��ȡ��������
Private Function DecryptionPassW(ByVal strPassW As String) As String
    Dim i As Integer
    Dim lngAsc  As Integer
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim lngBase As Long
    Dim strRandom As String
    Dim strPassSouce As String

    i = 0
    
    strPassSouce = Mid(strPassW, 2, Len(strPassW) - 2)
    lngPassWLength = Len(strPassSouce)
    lngBase = Asc(Mid(strPassW, 1, 1))
    
    strRandom = Right(strPassW, 1)
    
    ReDim intASC(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
    
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassSouce, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strRandom) Xor lngBase
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop

    DecryptionPassW = Join(strTemp, "") '���ܺ���ִ�
End Function

Private Sub optImageDblClick_Click(Index As Integer)
    If Index = 1 Then
        If chkPreText.value = 1 And optClickPreview.value Then
            MsgBox "���Ѿ����á���굥��ʱԤ��ͼ�񡯣�������ͼ˫�����ͼ��༭�Ĺ����غϣ���������ͼ˫����Ĺ���ѡ��ֱ��д�뱨��", vbOKOnly, "��ʾ��Ϣ"
        End If
    End If
End Sub

'Private Sub optBigImgAction_Click(Index As Integer)
'    If Index = 3 Then optImageDblClick(0).value = True
'End Sub

Private Sub optClickPreview_Click()
    If optClickPreview.value Then optImageDblClick(0).value = True
    
    If optMovePreview.value = False Then
        txtDelayTime.Enabled = False
        lblDelayTime.Enabled = False
    End If
End Sub

Private Sub LoadPathol()
    With lst(lst_PatholInfo)
        .AddItem "�걾����", 0
        .AddItem "ȡ��λ��", 1
        .AddItem "��״", 2
        .AddItem "������", 3
        .AddItem "��Ƭ��", 4
        .AddItem "��ȡҽʦ", 5
        .AddItem "ȡ��ʱ��", 6
        .AddItem "����", 7
        .AddItem "��ɫ", 8
        .AddItem "�걾��", 9
    End With
End Sub

'��ʼ�����ŷָ�������������ַ���֧�ּ����е����е��ֽڷ��ţ�����˫���ź͵�����֮�⡣
Private Sub LoadCheckNoDelimeter()
    Dim i As Integer
    
    For i = 1 To 2
        cboDelimeter(i).Clear
        cboDelimeter(i).AddItem "~"
        cboDelimeter(i).AddItem "`"
        cboDelimeter(i).AddItem "!"
        cboDelimeter(i).AddItem "@"
        cboDelimeter(i).AddItem "#"
        cboDelimeter(i).AddItem "$"
        cboDelimeter(i).AddItem "%"
        cboDelimeter(i).AddItem "^"
        cboDelimeter(i).AddItem "&"
        cboDelimeter(i).AddItem "*"
        cboDelimeter(i).AddItem "("
        cboDelimeter(i).AddItem ")"
        cboDelimeter(i).AddItem "_"
        cboDelimeter(i).AddItem "-"
        cboDelimeter(i).AddItem "+"
        cboDelimeter(i).AddItem "="
        cboDelimeter(i).AddItem "{"
        cboDelimeter(i).AddItem "}"
        cboDelimeter(i).AddItem "["
        cboDelimeter(i).AddItem "]"
        cboDelimeter(i).AddItem "\"
        cboDelimeter(i).AddItem "|"
        cboDelimeter(i).AddItem ";"
        cboDelimeter(i).AddItem ":"
        cboDelimeter(i).AddItem ","
        cboDelimeter(i).AddItem "<"
        cboDelimeter(i).AddItem ">"
        cboDelimeter(i).AddItem "."
        cboDelimeter(i).AddItem "/"
        cboDelimeter(i).AddItem "?"
    Next i
    
End Sub

'���ü��ŷָ���������
Private Sub setCheckNoDelimeter(lngIndex As Long, strText As String)
    On Error GoTo Err
    
    cboDelimeter(lngIndex).Text = strText
    chkDelimiter(lngIndex).value = 1
    
    Exit Sub
Err:
    '�����ֵʧ�ܣ���ȡ���ָ�����ѡ��
    cboDelimeter(lngIndex).ListIndex = -1
    chkDelimiter(lngIndex).value = 0
End Sub
