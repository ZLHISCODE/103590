VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSysConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ϵͳ����"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9690
   Icon            =   "frmSysConfig.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   1800
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdDefaultVal 
      Caption         =   "Ĭ������"
      Height          =   350
      Left            =   360
      TabIndex        =   272
      Top             =   6240
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   2460
      Top             =   6150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab sstabConfiguration 
      Height          =   5925
      Left            =   90
      TabIndex        =   28
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   10451
      _Version        =   393216
      Style           =   1
      MousePointer    =   99
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Ӱ����������"
      TabPicture(0)   =   "frmSysConfig.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "sstabModality"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "������Ϣ����"
      TabPicture(1)   =   "frmSysConfig.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label13"
      Tab(1).Control(1)=   "lstInfoLabelAll"
      Tab(1).Control(2)=   "Frame6(1)"
      Tab(1).Control(3)=   "Frame6(2)"
      Tab(1).Control(4)=   "Frame6(3)"
      Tab(1).Control(5)=   "Frame6(4)"
      Tab(1).Control(6)=   "cmdSelInfoLabel(1)"
      Tab(1).Control(7)=   "cmdSelInfoLabel(2)"
      Tab(1).Control(8)=   "cmdSelInfoLabel(4)"
      Tab(1).Control(9)=   "cmdSelInfoLabel(3)"
      Tab(1).Control(10)=   "cmdDeSelInfoLabel"
      Tab(1).Control(11)=   "cmdInfoLabelUpDown(0)"
      Tab(1).Control(12)=   "cmdInfoLabelUpDown(2)"
      Tab(1).Control(13)=   "Frame26"
      Tab(1).Control(14)=   "Frame29"
      Tab(1).Control(15)=   "cmdExportInf"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "����÷�����"
      TabPicture(2)   =   "frmSysConfig.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame30"
      Tab(2).Control(1)=   "Frame27"
      Tab(2).Control(2)=   "Frame24"
      Tab(2).Control(3)=   "cmdLeftRight(2)"
      Tab(2).Control(4)=   "cmdLeftRight(1)"
      Tab(2).Control(5)=   "Frame5(1)"
      Tab(2).Control(6)=   "Frame5(0)"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "��ע��Ϣ����"
      TabPicture(3)   =   "frmSysConfig.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame25"
      Tab(3).Control(1)=   "Frame22"
      Tab(3).Control(2)=   "Frame10"
      Tab(3).Control(3)=   "Frame12"
      Tab(3).Control(4)=   "Frame11"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "DICOM��ӡ����"
      TabPicture(4)   =   "frmSysConfig.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdDICOMPrintAdd"
      Tab(4).Control(1)=   "cmdDICOMPrintUpdate"
      Tab(4).Control(2)=   "cmdDICOMPrintDelete"
      Tab(4).Control(3)=   "Command6"
      Tab(4).Control(4)=   "frmPrintSetup(1)"
      Tab(4).Control(5)=   "MSFPrinter"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "��������"
      TabPicture(5)   =   "frmSysConfig.frx":0D56
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame19"
      Tab(5).Control(1)=   "Frame21(0)"
      Tab(5).Control(2)=   "fram22"
      Tab(5).Control(3)=   "Frame21(1)"
      Tab(5).Control(4)=   "Frame23"
      Tab(5).Control(5)=   "Frame33"
      Tab(5).ControlCount=   6
      TabCaption(6)   =   "ͼ�����"
      TabPicture(6)   =   "frmSysConfig.frx":0D72
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame31"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "�����û�"
      TabPicture(7)   =   "frmSysConfig.frx":0D8E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "CmdGetUserInfo"
      Tab(7).Control(1)=   "livGetUserSetup"
      Tab(7).ControlCount=   2
      Begin VB.Frame Frame25 
         Caption         =   "������Ϣ��ʾ����"
         Height          =   1750
         Left            =   -70080
         TabIndex        =   321
         Top             =   1680
         Width           =   4335
         Begin VB.CheckBox chkImgContainInfo 
            Caption         =   "����ͼ����������Ϣ"
            Height          =   255
            Left            =   120
            TabIndex        =   342
            Top             =   1370
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox chkInfoLabelScale 
            Caption         =   "��ͼ��ͬʱ����"
            Height          =   255
            Left            =   2400
            TabIndex        =   340
            Top             =   1050
            Width           =   1695
         End
         Begin VB.CheckBox chkPatientiInfoFontBold 
            Caption         =   "����"
            Height          =   255
            Left            =   2400
            TabIndex        =   339
            Top             =   310
            Width           =   735
         End
         Begin VB.CheckBox chkPatientInfoFontItalic 
            Caption         =   "б��"
            Height          =   255
            Left            =   3360
            TabIndex        =   338
            Top             =   310
            Width           =   735
         End
         Begin VB.ListBox lstPatientInfoFontSize 
            Height          =   240
            ItemData        =   "frmSysConfig.frx":0DAA
            Left            =   2880
            List            =   "frmSysConfig.frx":0DF9
            TabIndex        =   337
            Top             =   680
            Width           =   1095
         End
         Begin VB.TextBox txtPatientInfoFontName 
            Height          =   270
            Left            =   600
            TabIndex        =   336
            Top             =   640
            Width           =   1095
         End
         Begin VB.CommandButton cmdPatientInfoFont 
            Caption         =   "��"
            Height          =   255
            Left            =   1680
            TabIndex        =   335
            Top             =   648
            Width           =   255
         End
         Begin VB.CommandButton cmdInfoLabelColor 
            Caption         =   "��"
            Height          =   255
            Left            =   1680
            TabIndex        =   323
            Top             =   290
            Width           =   255
         End
         Begin VB.TextBox txtPatientInfoInVisibleSize 
            Height          =   285
            Left            =   840
            TabIndex        =   322
            Top             =   1000
            Width           =   375
         End
         Begin VB.Label Label14 
            Caption         =   "�ֺ�"
            Height          =   255
            Left            =   2400
            TabIndex        =   341
            Top             =   680
            Width           =   375
         End
         Begin VB.Shape shpInfoLabel 
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Left            =   600
            Top             =   290
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "��ɫ"
            Height          =   255
            Left            =   120
            TabIndex        =   326
            Top             =   310
            Width           =   375
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "ͼ��С��     ʱ����ʾ"
            Height          =   180
            Left            =   120
            TabIndex        =   325
            Top             =   1050
            Width           =   1890
         End
         Begin VB.Label lblPatientInfoFont 
            Caption         =   "����"
            Height          =   255
            Left            =   120
            TabIndex        =   324
            Top             =   680
            Width           =   495
         End
      End
      Begin VB.Frame Frame33 
         Caption         =   "���ò���"
         Height          =   2175
         Left            =   -71400
         TabIndex        =   316
         Top             =   3480
         Width           =   2535
         Begin VB.CheckBox chkShowMiniImageInfo 
            Caption         =   "����ͼ��ʾͼ����Ϣ"
            Height          =   180
            Left            =   120
            TabIndex        =   346
            Top             =   1800
            Width           =   2055
         End
         Begin VB.CheckBox chkPrintFilmBeep 
            Caption         =   "��Ƭ��ӡ��ʾ����"
            Height          =   180
            Left            =   120
            TabIndex        =   331
            Top             =   1080
            Width           =   2055
         End
         Begin VB.CheckBox chkShowMPRLine 
            Caption         =   "MPR��ʾ������"
            Height          =   180
            Left            =   120
            TabIndex        =   319
            Top             =   720
            Width           =   1815
         End
         Begin VB.CheckBox chkDockMiniImage 
            Caption         =   "����ͼͣ���ڲ˵���"
            Height          =   180
            Left            =   120
            TabIndex        =   318
            Top             =   1440
            Width           =   2055
         End
         Begin VB.CheckBox chkSquareFrame 
            Caption         =   "�����ο�ѡ"
            Height          =   180
            Left            =   120
            TabIndex        =   317
            ToolTipText     =   "��ѡ����ͼ��ʱ��ǿ��ʹ��������"
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame31 
         Caption         =   "�˾�ģ���������"
         Height          =   5055
         Left            =   -74760
         TabIndex        =   294
         Top             =   600
         Width           =   9015
         Begin VB.Frame Frame32 
            Height          =   975
            Left            =   240
            TabIndex        =   309
            Top             =   3480
            Width           =   8535
            Begin VB.TextBox txtFilterPara 
               Height          =   300
               Index           =   6
               Left            =   7440
               TabIndex        =   303
               Top             =   600
               Width           =   600
            End
            Begin VB.TextBox txtFilterPara 
               Height          =   300
               Index           =   5
               Left            =   7440
               TabIndex        =   302
               Top             =   240
               Width           =   600
            End
            Begin VB.TextBox txtFilterPara 
               Height          =   300
               Index           =   4
               Left            =   5040
               TabIndex        =   301
               Top             =   600
               Width           =   600
            End
            Begin VB.TextBox txtFilterPara 
               Height          =   300
               Index           =   3
               Left            =   5040
               TabIndex        =   300
               Top             =   240
               Width           =   600
            End
            Begin VB.TextBox txtFilterPara 
               Height          =   300
               Index           =   2
               Left            =   2040
               TabIndex        =   299
               Top             =   600
               Width           =   600
            End
            Begin VB.TextBox txtFilterPara 
               Height          =   300
               Index           =   1
               Left            =   2040
               TabIndex        =   298
               Top             =   240
               Width           =   600
            End
            Begin VB.Label Label63 
               Caption         =   "����ƽ��"
               Height          =   255
               Left            =   6480
               TabIndex        =   315
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label62 
               Caption         =   "����ƽ��"
               Height          =   255
               Left            =   6480
               TabIndex        =   314
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label61 
               Caption         =   "����ͼ����ǿ����"
               Height          =   255
               Left            =   3360
               TabIndex        =   313
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label Label60 
               Caption         =   "����ͼ����ǿ����"
               Height          =   255
               Left            =   3360
               TabIndex        =   312
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label59 
               Caption         =   "����ͼ����ǿǿ��"
               Height          =   255
               Left            =   360
               TabIndex        =   311
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label Label58 
               Caption         =   "����ͼ����ǿǿ��"
               Height          =   255
               Left            =   360
               TabIndex        =   310
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.TextBox txtFilterModality 
            Height          =   300
            Left            =   6720
            TabIndex        =   297
            Top             =   3097
            Width           =   2000
         End
         Begin VB.TextBox txtFilterName 
            Height          =   300
            Left            =   1560
            TabIndex        =   296
            Top             =   3097
            Width           =   2000
         End
         Begin VB.CommandButton cmdFilterAdd 
            Caption         =   "����"
            Height          =   345
            Left            =   5160
            TabIndex        =   304
            Top             =   4560
            Width           =   1100
         End
         Begin VB.CommandButton cmdFilterUpdate 
            Caption         =   "�޸�"
            Height          =   345
            Left            =   6360
            TabIndex        =   305
            Top             =   4560
            Width           =   1100
         End
         Begin VB.CommandButton cmdFilterDel 
            Caption         =   "ɾ��"
            Height          =   345
            Left            =   7530
            TabIndex        =   306
            Top             =   4560
            Width           =   1100
         End
         Begin MSFlexGridLib.MSFlexGrid MSFFilter 
            Height          =   2535
            Left            =   240
            TabIndex        =   295
            TabStop         =   0   'False
            Top             =   240
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   4471
            _Version        =   393216
            FixedCols       =   0
            WordWrap        =   -1  'True
            SelectionMode   =   1
            AllowUserResizing=   1
            MousePointer    =   1
         End
         Begin VB.Label Label57 
            Caption         =   "Ӱ�����"
            Height          =   255
            Left            =   5280
            TabIndex        =   308
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label55 
            Caption         =   "�˾�ģ������"
            Height          =   255
            Left            =   240
            TabIndex        =   307
            Top             =   3120
            Width           =   1095
         End
      End
      Begin VB.CommandButton CmdGetUserInfo 
         Caption         =   "��ȡ(&G)"
         Height          =   350
         Left            =   -66960
         Style           =   1  'Graphical
         TabIndex        =   292
         Top             =   5400
         Width           =   1100
      End
      Begin VB.CommandButton cmdExportInf 
         Caption         =   "��Ϣ����"
         Height          =   350
         Left            =   -71640
         TabIndex        =   290
         Top             =   3600
         Width           =   1100
      End
      Begin VB.Frame Frame30 
         Caption         =   "������"
         Height          =   1335
         Left            =   -67440
         TabIndex        =   288
         Top             =   4200
         Width           =   1815
         Begin VB.ComboBox cboMouseWheelDrag 
            Height          =   300
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   344
            Top             =   840
            Width           =   1095
         End
         Begin VB.ComboBox cboMouseWheelRoll 
            Height          =   300
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   289
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label67 
            Caption         =   "��ק"
            Height          =   255
            Left            =   120
            TabIndex        =   345
            Top             =   870
            Width           =   375
         End
         Begin VB.Label Label51 
            Caption         =   "����"
            Height          =   255
            Left            =   120
            TabIndex        =   343
            Top             =   390
            Width           =   495
         End
      End
      Begin VB.Frame Frame29 
         Caption         =   "��ϸ��Ϣ"
         Height          =   1695
         Left            =   -74880
         TabIndex        =   276
         Top             =   4155
         Width           =   3015
         Begin VB.CommandButton cmdInfoDelete 
            Caption         =   "ɾ��"
            Height          =   350
            Left            =   2040
            TabIndex        =   282
            Top             =   1200
            Width           =   900
         End
         Begin VB.CommandButton cmdInfoUpdate 
            Caption         =   "�޸�"
            Height          =   350
            Left            =   1080
            TabIndex        =   281
            Top             =   1200
            Width           =   900
         End
         Begin VB.CommandButton cmdInfoAdd 
            Caption         =   "����"
            Height          =   350
            Left            =   120
            TabIndex        =   280
            Top             =   1200
            Width           =   900
         End
         Begin VB.TextBox txtUserLabelValue 
            Height          =   300
            Left            =   360
            TabIndex        =   278
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label56 
            Caption         =   "ֵ"
            Height          =   300
            Left            =   120
            TabIndex        =   279
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lblInfoType 
            Caption         =   "�û���Ϣ"
            Height          =   255
            Left            =   120
            TabIndex        =   277
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "����λ��ʾλ��"
         Height          =   735
         Left            =   -74760
         TabIndex        =   198
         Top             =   5040
         Width           =   4575
         Begin VB.OptionButton opWinWLLocation 
            Caption         =   "��"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   202
            Top             =   300
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton opWinWLLocation 
            Caption         =   "��"
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   201
            Top             =   300
            Width           =   495
         End
         Begin VB.OptionButton opWinWLLocation 
            Caption         =   "��"
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   200
            Top             =   300
            Width           =   495
         End
         Begin VB.OptionButton opWinWLLocation 
            Caption         =   "��"
            Height          =   255
            Index           =   4
            Left            =   3360
            TabIndex        =   199
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.Frame Frame27 
         Caption         =   "��긽�Ӽ�"
         Height          =   1455
         Left            =   -67440
         TabIndex        =   193
         Top             =   720
         Width           =   1815
         Begin VB.CheckBox chkShiftState 
            Caption         =   "ʹ�� Shift ��"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   196
            Tag             =   "0"
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkShiftState 
            Caption         =   "ʹ�� Ctrl ��"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   195
            Tag             =   "0"
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkShiftState 
            Caption         =   "ʹ�� Alt ��"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   194
            Tag             =   "0"
            Top             =   1080
            Width           =   1575
         End
      End
      Begin VB.Frame Frame26 
         Caption         =   "��ͷ"
         Height          =   1695
         Left            =   -71760
         TabIndex        =   189
         Top             =   4140
         Width           =   1455
         Begin VB.OptionButton optPatientInfoTitle 
            Caption         =   "Ӣ����ͷ"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   192
            Top             =   1320
            Width           =   1215
         End
         Begin VB.OptionButton optPatientInfoTitle 
            Caption         =   "������ͷ"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   191
            Top             =   900
            Width           =   1215
         End
         Begin VB.OptionButton optPatientInfoTitle 
            Caption         =   "����ʾ��ͷ"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   190
            Top             =   480
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "�������"
         Height          =   2280
         Left            =   -70080
         TabIndex        =   162
         Top             =   3495
         Width           =   4335
         Begin VB.Frame Frame13 
            Caption         =   "�����ʾλ��"
            ForeColor       =   &H80000007&
            Height          =   1815
            Left            =   120
            TabIndex        =   170
            Top             =   320
            Width           =   2055
            Begin VB.CheckBox chkRulerDsip 
               Caption         =   "��"
               Height          =   255
               Index           =   1
               Left            =   300
               TabIndex        =   176
               Top             =   240
               Width           =   615
            End
            Begin VB.CheckBox chkRulerDsip 
               Caption         =   "��"
               Height          =   255
               Index           =   2
               Left            =   300
               TabIndex        =   175
               Top             =   600
               Width           =   615
            End
            Begin VB.CheckBox chkRulerDsip 
               Caption         =   "��"
               Height          =   255
               Index           =   3
               Left            =   1350
               TabIndex        =   174
               Top             =   240
               Width           =   615
            End
            Begin VB.CheckBox chkRulerDsip 
               Caption         =   "��"
               Height          =   255
               Index           =   4
               Left            =   1350
               TabIndex        =   173
               Top             =   600
               Width           =   615
            End
            Begin VB.ListBox lstRulerSize 
               Height          =   240
               Index           =   1
               ItemData        =   "frmSysConfig.frx":0E58
               Left            =   960
               List            =   "frmSysConfig.frx":0E5A
               TabIndex        =   172
               Top             =   1080
               Width           =   855
            End
            Begin VB.ListBox lstRulerSize 
               Height          =   240
               Index           =   2
               ItemData        =   "frmSysConfig.frx":0E5C
               Left            =   960
               List            =   "frmSysConfig.frx":0E5E
               TabIndex        =   171
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label39 
               Caption         =   "���ұ߾�"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   178
               Top             =   1110
               Width           =   735
            End
            Begin VB.Label Label39 
               Caption         =   "���±߾�"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   177
               Top             =   1470
               Width           =   735
            End
         End
         Begin VB.ListBox lstRulerLineWidth 
            Height          =   240
            ItemData        =   "frmSysConfig.frx":0E60
            Left            =   3000
            List            =   "frmSysConfig.frx":0E62
            TabIndex        =   169
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Frame Frame14 
            Caption         =   "��ߴ�С"
            Height          =   1215
            Left            =   2280
            TabIndex        =   164
            Top             =   300
            Width           =   1935
            Begin VB.ListBox lstRulerSize 
               Height          =   240
               Index           =   3
               ItemData        =   "frmSysConfig.frx":0E64
               Left            =   720
               List            =   "frmSysConfig.frx":0E66
               TabIndex        =   166
               Top             =   300
               Width           =   1065
            End
            Begin VB.ListBox lstRulerSize 
               Height          =   240
               Index           =   4
               ItemData        =   "frmSysConfig.frx":0E68
               Left            =   720
               List            =   "frmSysConfig.frx":0E6A
               TabIndex        =   165
               Top             =   780
               Width           =   1065
            End
            Begin VB.Label Label39 
               Caption         =   "���"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   168
               Top             =   330
               Width           =   615
            End
            Begin VB.Label Label39 
               Caption         =   "�߶�"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   167
               Top             =   810
               Width           =   615
            End
         End
         Begin VB.CommandButton cmdLabelConfig 
            Caption         =   "��"
            Height          =   255
            Index           =   3
            Left            =   3840
            TabIndex        =   163
            Top             =   1920
            Width           =   255
         End
         Begin VB.Label Label37 
            Caption         =   "�߿�"
            Height          =   225
            Left            =   2400
            TabIndex        =   180
            Top             =   1590
            Width           =   615
         End
         Begin VB.Label Label38 
            Caption         =   "��ɫ"
            Height          =   285
            Left            =   2400
            TabIndex        =   179
            Top             =   1950
            Width           =   495
         End
         Begin VB.Shape shpLabelConfig 
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   3
            Left            =   3000
            Top             =   1920
            Width           =   855
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "��λ�������"
         Height          =   1095
         Left            =   -70080
         TabIndex        =   155
         Top             =   480
         Width           =   4335
         Begin VB.Frame Frame15 
            Caption         =   "��λ�����ʾλ��"
            ForeColor       =   &H80000007&
            Height          =   735
            Left            =   1440
            TabIndex        =   157
            Top             =   240
            Width           =   2775
            Begin VB.CheckBox chkAnatomicMarkers 
               Caption         =   "��"
               Height          =   255
               Index           =   4
               Left            =   2040
               TabIndex        =   161
               Top             =   360
               Width           =   615
            End
            Begin VB.CheckBox chkAnatomicMarkers 
               Caption         =   "��"
               Height          =   255
               Index           =   3
               Left            =   1440
               TabIndex        =   160
               Top             =   360
               Width           =   615
            End
            Begin VB.CheckBox chkAnatomicMarkers 
               Caption         =   "��"
               Height          =   255
               Index           =   2
               Left            =   840
               TabIndex        =   159
               Top             =   360
               Width           =   615
            End
            Begin VB.CheckBox chkAnatomicMarkers 
               Caption         =   "��"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   158
               Top             =   360
               Width           =   615
            End
         End
         Begin VB.CheckBox chkChinaMark 
            Caption         =   "���ı��"
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "��ע����"
         Height          =   4455
         Left            =   -74760
         TabIndex        =   134
         Top             =   480
         Width           =   4575
         Begin VB.Frame Frame16 
            Caption         =   "�������ֵ���ʾ"
            Height          =   2055
            Left            =   120
            TabIndex        =   144
            Top             =   2280
            Width           =   4335
            Begin VB.Frame Frame17 
               Caption         =   "��ʾ����"
               Height          =   975
               Left            =   120
               TabIndex        =   149
               Top             =   240
               Width           =   4095
               Begin VB.CheckBox chkMeasureResult 
                  Caption         =   "��Сֵ"
                  Height          =   255
                  Index           =   6
                  Left            =   3000
                  TabIndex        =   286
                  Top             =   600
                  Width           =   855
               End
               Begin VB.CheckBox chkMeasureResult 
                  Caption         =   "���ֵ"
                  Height          =   255
                  Index           =   5
                  Left            =   3000
                  TabIndex        =   285
                  Top             =   240
                  Width           =   855
               End
               Begin VB.CheckBox chkMeasureResult 
                  Caption         =   "�ܳ�"
                  Height          =   255
                  Index           =   4
                  Left            =   240
                  TabIndex        =   273
                  Top             =   600
                  Width           =   735
               End
               Begin VB.CheckBox chkMeasureResult 
                  Caption         =   "������"
                  Height          =   255
                  Index           =   3
                  Left            =   1560
                  TabIndex        =   152
                  Top             =   240
                  Width           =   855
               End
               Begin VB.CheckBox chkMeasureResult 
                  Caption         =   "ƽ��ֵ"
                  Height          =   255
                  Index           =   2
                  Left            =   1560
                  TabIndex        =   151
                  Top             =   600
                  Width           =   855
               End
               Begin VB.CheckBox chkMeasureResult 
                  Caption         =   "���"
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   150
                  Top             =   240
                  Width           =   735
               End
            End
            Begin VB.CheckBox chkLabelText 
               Caption         =   "������ͼ��ͬʱ����"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   148
               Top             =   1680
               Width           =   1935
            End
            Begin VB.ListBox lstTextoOff 
               Height          =   240
               Index           =   1
               ItemData        =   "frmSysConfig.frx":0E6C
               Left            =   1320
               List            =   "frmSysConfig.frx":0E6E
               TabIndex        =   147
               Top             =   1320
               Width           =   495
            End
            Begin VB.ListBox lstTextoOff 
               Height          =   240
               Index           =   2
               ItemData        =   "frmSysConfig.frx":0E70
               Left            =   3600
               List            =   "frmSysConfig.frx":0E72
               TabIndex        =   146
               Top             =   1320
               Width           =   495
            End
            Begin VB.CheckBox chkLabelText 
               Caption         =   "���Ĳ�����Ϣ"
               Height          =   255
               Index           =   2
               Left            =   2520
               TabIndex        =   145
               Top             =   1680
               Width           =   1455
            End
            Begin VB.Label Label40 
               Caption         =   "Y����ƫ����"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   154
               Top             =   1350
               Width           =   1095
            End
            Begin VB.Label Label40 
               Caption         =   "X����ƫ����"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   153
               Top             =   1350
               Width           =   1095
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "����"
            Height          =   1065
            Index           =   0
            Left            =   120
            TabIndex        =   138
            Top             =   240
            Width           =   4335
            Begin VB.TextBox txtLabelLineWidth 
               Enabled         =   0   'False
               Height          =   285
               Left            =   3120
               TabIndex        =   208
               Text            =   "1"
               Top             =   240
               Width           =   855
            End
            Begin MSComCtl2.UpDown udLabelFontSize 
               Height          =   285
               Left            =   3960
               TabIndex        =   206
               Top             =   600
               Width           =   255
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtLabelFontSize 
               Enabled         =   0   'False
               Height          =   285
               Left            =   3120
               TabIndex        =   205
               Text            =   "1"
               Top             =   600
               Width           =   855
            End
            Begin VB.ComboBox cboLabelLineStyle 
               Height          =   315
               ItemData        =   "frmSysConfig.frx":0E74
               Left            =   600
               List            =   "frmSysConfig.frx":0E87
               TabIndex        =   140
               Top             =   600
               Width           =   1335
            End
            Begin VB.CommandButton cmdLabelConfig 
               Caption         =   "��"
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   139
               Top             =   240
               Width           =   255
            End
            Begin MSComCtl2.UpDown udLabelLineWidth 
               Height          =   285
               Left            =   3960
               TabIndex        =   207
               Top             =   240
               Width           =   255
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label42 
               Caption         =   "���ִ�С"
               Height          =   255
               Index           =   2
               Left            =   2280
               TabIndex        =   204
               Top             =   630
               Width           =   735
            End
            Begin VB.Label Label41 
               Caption         =   "��ɫ"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   143
               Top             =   270
               Width           =   375
            End
            Begin VB.Shape shpLabelConfig 
               FillColor       =   &H008080FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   1
               Left            =   600
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label42 
               Caption         =   "�߿�"
               Height          =   255
               Index           =   0
               Left            =   2280
               TabIndex        =   142
               Top             =   270
               Width           =   495
            End
            Begin VB.Label Label43 
               Caption         =   "����"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   141
               Top             =   630
               Width           =   495
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "ѡ��"
            Height          =   765
            Index           =   1
            Left            =   120
            TabIndex        =   135
            Top             =   1410
            Width           =   4335
            Begin VB.CommandButton cmdLabelConfig 
               Caption         =   "��"
               Height          =   255
               Index           =   2
               Left            =   1680
               TabIndex        =   136
               Top             =   300
               Width           =   255
            End
            Begin VB.Shape shpLabelConfig 
               FillColor       =   &H008080FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   2
               Left            =   600
               Top             =   300
               Width           =   1095
            End
            Begin VB.Label Label41 
               Caption         =   "��ɫ"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   137
               Top             =   330
               Width           =   375
            End
         End
      End
      Begin VB.CommandButton cmdDICOMPrintAdd 
         Caption         =   "����"
         Height          =   350
         Left            =   -74400
         TabIndex        =   21
         Top             =   5400
         Width           =   1100
      End
      Begin VB.CommandButton cmdDICOMPrintUpdate 
         Caption         =   "�޸�"
         Height          =   350
         Left            =   -72960
         TabIndex        =   22
         Top             =   5400
         Width           =   1100
      End
      Begin VB.CommandButton cmdDICOMPrintDelete 
         Caption         =   "ɾ��"
         Height          =   350
         Left            =   -71520
         TabIndex        =   23
         Top             =   5400
         Width           =   1100
      End
      Begin VB.CommandButton Command6 
         Caption         =   "��֤"
         Height          =   350
         Left            =   -66840
         TabIndex        =   24
         Top             =   5400
         Width           =   1100
      End
      Begin VB.Frame Frame24 
         Caption         =   "��경��"
         Height          =   1815
         Left            =   -67440
         TabIndex        =   116
         Top             =   2280
         Width           =   1815
         Begin VB.ListBox lstMouseStep 
            Height          =   240
            Index           =   4
            ItemData        =   "frmSysConfig.frx":0EAF
            Left            =   600
            List            =   "frmSysConfig.frx":0FDC
            TabIndex        =   188
            Top             =   1440
            Width           =   975
         End
         Begin VB.ListBox lstMouseStep 
            Height          =   240
            Index           =   3
            ItemData        =   "frmSysConfig.frx":1163
            Left            =   600
            List            =   "frmSysConfig.frx":1290
            TabIndex        =   187
            Top             =   1080
            Width           =   975
         End
         Begin VB.ListBox lstMouseStep 
            Height          =   240
            Index           =   2
            ItemData        =   "frmSysConfig.frx":1417
            Left            =   600
            List            =   "frmSysConfig.frx":1544
            TabIndex        =   186
            Top             =   720
            Width           =   975
         End
         Begin VB.ListBox lstMouseStep 
            Height          =   240
            Index           =   1
            ItemData        =   "frmSysConfig.frx":16CB
            Left            =   600
            List            =   "frmSysConfig.frx":17F8
            TabIndex        =   185
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label50 
            Caption         =   "����"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   120
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label50 
            Caption         =   "����"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   119
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label50 
            Caption         =   "����"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   118
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label50 
            Caption         =   "����"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   117
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "��λ��"
         Height          =   1335
         Left            =   -71400
         TabIndex        =   113
         Top             =   1800
         Width           =   2535
         Begin VB.ListBox lstReferenceLineSpacing 
            Height          =   240
            ItemData        =   "frmSysConfig.frx":197F
            Left            =   960
            List            =   "frmSysConfig.frx":19A4
            TabIndex        =   181
            Top             =   1020
            Width           =   1095
         End
         Begin VB.CommandButton cmdUserInterfaceColor 
            Caption         =   "��"
            Height          =   255
            Index           =   6
            Left            =   1800
            TabIndex        =   128
            Top             =   300
            Width           =   255
         End
         Begin VB.ComboBox cboReferenceLineStyle 
            Height          =   300
            ItemData        =   "frmSysConfig.frx":19CA
            Left            =   960
            List            =   "frmSysConfig.frx":19DD
            TabIndex        =   122
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label Label47 
            Caption         =   "���"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   182
            Top             =   1050
            Width           =   375
         End
         Begin VB.Label Label46 
            Caption         =   "����"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   115
            Top             =   690
            Width           =   615
         End
         Begin VB.Shape shpUserInterface 
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   6
            Left            =   960
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label44 
            Caption         =   "��ɫ"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   114
            Top             =   330
            Width           =   615
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "��עѡ����"
         Height          =   975
         Index           =   1
         Left            =   -71400
         TabIndex        =   109
         Top             =   600
         Width           =   2535
         Begin VB.CommandButton cmdUserInterfaceColor 
            Caption         =   "��"
            Height          =   255
            Index           =   5
            Left            =   1800
            TabIndex        =   127
            Top             =   300
            Width           =   255
         End
         Begin VB.ListBox lstPeriodSize 
            Height          =   240
            ItemData        =   "frmSysConfig.frx":1A05
            Left            =   960
            List            =   "frmSysConfig.frx":1A07
            TabIndex        =   110
            Top             =   660
            Width           =   1095
         End
         Begin VB.Shape shpUserInterface 
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   5
            Left            =   960
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label44 
            Caption         =   "��ɫ"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   112
            Top             =   330
            Width           =   615
         End
         Begin VB.Label Label48 
            Caption         =   "��С"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   111
            Top             =   690
            Width           =   615
         End
      End
      Begin VB.Frame fram22 
         Caption         =   "��������"
         Height          =   5055
         Left            =   -68640
         TabIndex        =   102
         Top             =   600
         Width           =   3015
         Begin VB.CheckBox chkShowPrintTag 
            Caption         =   "��ʾ��ӡ���"
            Height          =   180
            Left            =   240
            TabIndex        =   291
            ToolTipText     =   "��ѡ����ͼ��ʱ��ǿ��ʹ��������"
            Top             =   3720
            Width           =   2295
         End
         Begin VB.ListBox lstStatusBarFontSize 
            Height          =   240
            ItemData        =   "frmSysConfig.frx":1A09
            Left            =   1560
            List            =   "frmSysConfig.frx":1A0B
            TabIndex        =   210
            Top             =   1980
            Width           =   1215
         End
         Begin VB.CheckBox chkShowFilmConfig 
            Caption         =   "����ʱ������Ƭ���ô���"
            Height          =   180
            Left            =   240
            TabIndex        =   197
            Top             =   4650
            Width           =   2295
         End
         Begin VB.CommandButton cmdUserInterfaceColor 
            Caption         =   "��"
            Height          =   255
            Index           =   8
            Left            =   2400
            TabIndex        =   183
            Top             =   3120
            Width           =   255
         End
         Begin VB.ListBox lstCellSpacing 
            Height          =   240
            ItemData        =   "frmSysConfig.frx":1A0D
            Left            =   1560
            List            =   "frmSysConfig.frx":1A0F
            TabIndex        =   133
            Top             =   1590
            Width           =   1215
         End
         Begin VB.ListBox lstMaxAreaY 
            Height          =   240
            ItemData        =   "frmSysConfig.frx":1A11
            Left            =   1560
            List            =   "frmSysConfig.frx":1A13
            TabIndex        =   132
            Top             =   1170
            Width           =   1215
         End
         Begin VB.ListBox lstMaxAreaX 
            Height          =   240
            ItemData        =   "frmSysConfig.frx":1A15
            Left            =   1560
            List            =   "frmSysConfig.frx":1A17
            TabIndex        =   131
            Top             =   765
            Width           =   1215
         End
         Begin VB.CommandButton cmdUserInterfaceColor 
            Caption         =   "��"
            Height          =   255
            Index           =   7
            Left            =   2400
            TabIndex        =   129
            Top             =   2595
            Width           =   255
         End
         Begin VB.CheckBox chkDsipSpilthBorder 
            Caption         =   "����߿��Ƿ���ʾ"
            Height          =   180
            Left            =   240
            TabIndex        =   107
            Top             =   4215
            Width           =   2295
         End
         Begin VB.ListBox lstSpaceSize 
            Height          =   240
            ItemData        =   "frmSysConfig.frx":1A19
            Left            =   1560
            List            =   "frmSysConfig.frx":1A1B
            TabIndex        =   130
            Top             =   360
            Width           =   1215
         End
         Begin VB.Shape shpUserInterface 
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   8
            Left            =   1560
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label44 
            Caption         =   "״̬�������С"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   209
            Top             =   2010
            Width           =   1305
         End
         Begin VB.Label Label44 
            Caption         =   "���򱳾���ɫ"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   184
            Top             =   3150
            Width           =   1095
         End
         Begin VB.Label Label44 
            Caption         =   "ͼ�񱳾���ɫ"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   108
            Top             =   2625
            Width           =   1095
         End
         Begin VB.Shape shpUserInterface 
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   7
            Left            =   1560
            Top             =   2595
            Width           =   855
         End
         Begin VB.Label Label49 
            Caption         =   "ͼ����"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   106
            Top             =   1620
            Width           =   1095
         End
         Begin VB.Label Label49 
            Caption         =   "�������������"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   105
            Top             =   1215
            Width           =   1455
         End
         Begin VB.Label Label49 
            Caption         =   "�������������"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   104
            Top             =   795
            Width           =   1455
         End
         Begin VB.Label Label49 
            Caption         =   "���м���"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   103
            Top             =   390
            Width           =   1095
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "ͼ��ѡ����"
         Height          =   975
         Index           =   0
         Left            =   -74640
         TabIndex        =   98
         Top             =   4560
         Width           =   2895
         Begin VB.CommandButton cmdUserInterfaceColor 
            Caption         =   "��"
            Height          =   255
            Index           =   4
            Left            =   2400
            TabIndex        =   126
            Top             =   300
            Width           =   285
         End
         Begin VB.ListBox lstImageIdentifierSize 
            Height          =   240
            ItemData        =   "frmSysConfig.frx":1A1D
            Left            =   1560
            List            =   "frmSysConfig.frx":1A1F
            TabIndex        =   101
            Top             =   660
            Width           =   1150
         End
         Begin VB.Label Label48 
            Caption         =   "��С"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   100
            Top             =   690
            Width           =   615
         End
         Begin VB.Label Label44 
            Caption         =   "��ɫ"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   99
            Top             =   330
            Width           =   615
         End
         Begin VB.Shape shpUserInterface 
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   4
            Left            =   1560
            Top             =   300
            Width           =   855
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "ͼ��ѡ���"
         Height          =   5055
         Left            =   -74760
         TabIndex        =   85
         Top             =   600
         Width           =   3255
         Begin VB.Frame Frame20 
            Caption         =   "ѡ��"
            Height          =   1455
            Index           =   1
            Left            =   120
            TabIndex        =   92
            Top             =   360
            Width           =   2895
            Begin VB.CommandButton cmdUserInterfaceColor 
               Caption         =   "��"
               Height          =   255
               Index           =   1
               Left            =   2400
               TabIndex        =   123
               Top             =   300
               Width           =   285
            End
            Begin VB.ListBox lstNoSelectLineWidth 
               Height          =   240
               ItemData        =   "frmSysConfig.frx":1A21
               Left            =   1560
               List            =   "frmSysConfig.frx":1A43
               TabIndex        =   93
               Top             =   1020
               Width           =   1150
            End
            Begin VB.ComboBox cboNoSelectLineStyle 
               Height          =   300
               ItemData        =   "frmSysConfig.frx":1A66
               Left            =   1560
               List            =   "frmSysConfig.frx":1A79
               TabIndex        =   94
               Top             =   660
               Width           =   1150
            End
            Begin VB.Shape shpUserInterface 
               FillColor       =   &H008080FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   1
               Left            =   1560
               Top             =   300
               Width           =   855
            End
            Begin VB.Label Label44 
               Caption         =   "�߿���ɫ"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   97
               Top             =   330
               Width           =   1095
            End
            Begin VB.Label Label46 
               Caption         =   "�߿�����"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   96
               Top             =   690
               Width           =   855
            End
            Begin VB.Label Label47 
               Caption         =   "�߿��߿�"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   95
               Top             =   1050
               Width           =   855
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "��ǰ"
            Height          =   1815
            Index           =   0
            Left            =   120
            TabIndex        =   86
            Top             =   1920
            Width           =   2895
            Begin VB.CommandButton cmdUserInterfaceColor 
               Caption         =   "��"
               Height          =   255
               Index           =   3
               Left            =   2400
               TabIndex        =   125
               Top             =   660
               Width           =   285
            End
            Begin VB.CommandButton cmdUserInterfaceColor 
               Caption         =   "��"
               Height          =   255
               Index           =   2
               Left            =   2400
               TabIndex        =   124
               Top             =   300
               Width           =   285
            End
            Begin VB.ComboBox cboSelectLineStyle 
               Height          =   300
               ItemData        =   "frmSysConfig.frx":1AA1
               Left            =   1560
               List            =   "frmSysConfig.frx":1AB4
               TabIndex        =   121
               Top             =   1020
               Width           =   1150
            End
            Begin VB.ListBox lstSelectLineWidth 
               Height          =   240
               ItemData        =   "frmSysConfig.frx":1ADC
               Left            =   1560
               List            =   "frmSysConfig.frx":1ADE
               TabIndex        =   91
               Top             =   1400
               Width           =   1150
            End
            Begin VB.Label Label47 
               Caption         =   "ͼ��߿��߿�"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   90
               Top             =   1393
               Width           =   1215
            End
            Begin VB.Label Label46 
               Caption         =   "ͼ��߿�����"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   89
               Top             =   1050
               Width           =   1095
            End
            Begin VB.Shape shpUserInterface 
               FillColor       =   &H008080FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   3
               Left            =   1560
               Top             =   660
               Width           =   855
            End
            Begin VB.Label Label45 
               Caption         =   "���б߿���ɫ"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   88
               Top             =   690
               Width           =   1095
            End
            Begin VB.Shape shpUserInterface 
               FillColor       =   &H008080FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   2
               Left            =   1560
               Top             =   300
               Width           =   855
            End
            Begin VB.Label Label44 
               Caption         =   "ͼ��߿���ɫ"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   87
               Top             =   330
               Width           =   1095
            End
         End
      End
      Begin VB.Frame frmPrintSetup 
         BorderStyle     =   0  'None
         Height          =   3135
         Index           =   1
         Left            =   -74880
         TabIndex        =   62
         Top             =   2160
         Width           =   9255
         Begin VB.CheckBox chkPrintOkEcho 
            Caption         =   "��ӡ�ɹ�����ʾ"
            Height          =   255
            Left            =   6360
            TabIndex        =   284
            Top             =   648
            Width           =   1575
         End
         Begin VB.CommandButton CmdFilmFontSizeSetup 
            Caption         =   "��Ƭ��������"
            Height          =   350
            Left            =   7930
            TabIndex        =   283
            Top             =   600
            Width           =   1290
         End
         Begin VB.TextBox txtLocalAE 
            Height          =   300
            Left            =   7200
            TabIndex        =   5
            Top             =   180
            Width           =   1995
         End
         Begin VB.Frame Frame9 
            Caption         =   "�ܶ�"
            Height          =   2055
            Left            =   7800
            TabIndex        =   69
            Top             =   1080
            Width           =   1455
            Begin VB.ComboBox cboEmptyDensity 
               Height          =   300
               ItemData        =   "frmSysConfig.frx":1AE0
               Left            =   120
               List            =   "frmSysConfig.frx":1AEA
               TabIndex        =   18
               Top             =   1080
               Width           =   1215
            End
            Begin VB.ComboBox cboBorderDensity 
               Height          =   300
               ItemData        =   "frmSysConfig.frx":1AFC
               Left            =   120
               List            =   "frmSysConfig.frx":1B06
               TabIndex        =   17
               Top             =   480
               Width           =   1215
            End
            Begin VB.ListBox lstDensity 
               Height          =   240
               Index           =   2
               ItemData        =   "frmSysConfig.frx":1B18
               Left            =   720
               List            =   "frmSysConfig.frx":1B28
               TabIndex        =   20
               Top             =   1680
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.ListBox lstDensity 
               Height          =   240
               Index           =   1
               ItemData        =   "frmSysConfig.frx":1B38
               Left            =   120
               List            =   "frmSysConfig.frx":1B48
               TabIndex        =   19
               Top             =   1680
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label Label34 
               Caption         =   "�߿��ܶ�"
               Height          =   255
               Left            =   120
               TabIndex        =   84
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label33 
               Caption         =   "�հ��ܶ�"
               Height          =   255
               Left            =   120
               TabIndex        =   83
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label32 
               Caption         =   "��Сֵ"
               Height          =   255
               Left            =   720
               TabIndex        =   82
               Top             =   1440
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label31 
               Caption         =   "���ֵ"
               Height          =   255
               Left            =   120
               TabIndex        =   81
               Top             =   1440
               Visible         =   0   'False
               Width           =   615
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "����"
            Height          =   2055
            Left            =   120
            TabIndex        =   68
            Top             =   1080
            Width           =   7455
            Begin VB.TextBox txtImageResolution 
               Height          =   300
               Left            =   6120
               MaxLength       =   3
               TabIndex        =   333
               ToolTipText     =   "ͼ��������ȣ�һ��Ϊ300"
               Top             =   1080
               Width           =   735
            End
            Begin VB.TextBox txtImageBorderWidth 
               Height          =   300
               Left            =   4680
               MaxLength       =   2
               TabIndex        =   329
               ToolTipText     =   "ͼ��ı߿��ȣ�����ֵ��1-30֮��"
               Top             =   1080
               Width           =   1215
            End
            Begin VB.ComboBox cboPolarity 
               Height          =   300
               ItemData        =   "frmSysConfig.frx":1B58
               Left            =   6000
               List            =   "frmSysConfig.frx":1B62
               TabIndex        =   327
               Top             =   1680
               Width           =   1215
            End
            Begin VB.ComboBox cboBitDepth 
               Height          =   315
               ItemData        =   "frmSysConfig.frx":1B77
               Left            =   4680
               List            =   "frmSysConfig.frx":1B81
               TabIndex        =   274
               Top             =   1680
               Width           =   1215
            End
            Begin VB.ComboBox cboTrim 
               Height          =   315
               ItemData        =   "frmSysConfig.frx":1B8C
               Left            =   3360
               List            =   "frmSysConfig.frx":1B96
               TabIndex        =   16
               Top             =   1680
               Width           =   1215
            End
            Begin VB.ComboBox cboSmooth 
               Height          =   315
               ItemData        =   "frmSysConfig.frx":1BA3
               Left            =   1800
               List            =   "frmSysConfig.frx":1BB0
               TabIndex        =   15
               Top             =   1680
               Width           =   1455
            End
            Begin VB.ComboBox cboMagnification 
               Height          =   315
               ItemData        =   "frmSysConfig.frx":1BCB
               Left            =   120
               List            =   "frmSysConfig.frx":1BDB
               TabIndex        =   14
               Top             =   1680
               Width           =   1575
            End
            Begin VB.ComboBox cboResolution 
               Height          =   300
               ItemData        =   "frmSysConfig.frx":1C01
               Left            =   6120
               List            =   "frmSysConfig.frx":1C0B
               TabIndex        =   13
               Top             =   480
               Width           =   1215
            End
            Begin VB.ComboBox cboFilmBox 
               Height          =   315
               ItemData        =   "frmSysConfig.frx":1C1F
               Left            =   3360
               List            =   "frmSysConfig.frx":1C47
               TabIndex        =   12
               Top             =   1080
               Width           =   1215
            End
            Begin VB.ComboBox cboFilmSize 
               Height          =   315
               Left            =   1800
               TabIndex        =   11
               Top             =   1080
               Width           =   1455
            End
            Begin VB.ComboBox cboOrientation 
               Height          =   315
               ItemData        =   "frmSysConfig.frx":1CA7
               Left            =   120
               List            =   "frmSysConfig.frx":1CB1
               TabIndex        =   10
               Top             =   1080
               Width           =   1575
            End
            Begin VB.ListBox lstCopies 
               Height          =   240
               ItemData        =   "frmSysConfig.frx":1CCA
               Left            =   4680
               List            =   "frmSysConfig.frx":1CEC
               TabIndex        =   9
               Top             =   480
               Width           =   1215
            End
            Begin VB.ComboBox cboMedium 
               Height          =   315
               ItemData        =   "frmSysConfig.frx":1D0F
               Left            =   3360
               List            =   "frmSysConfig.frx":1D19
               TabIndex        =   8
               Top             =   450
               Width           =   1215
            End
            Begin VB.ComboBox cboPriority 
               Height          =   315
               ItemData        =   "frmSysConfig.frx":1D34
               Left            =   1800
               List            =   "frmSysConfig.frx":1D41
               TabIndex        =   7
               Top             =   450
               Width           =   1455
            End
            Begin VB.ComboBox cboFormat 
               Height          =   315
               Left            =   120
               TabIndex        =   6
               Top             =   450
               Width           =   1575
            End
            Begin VB.Label Label66 
               Caption         =   "PPI"
               Height          =   255
               Left            =   6960
               TabIndex        =   334
               Top             =   1125
               Width           =   375
            End
            Begin VB.Label Label65 
               Caption         =   "ͼƬ�ֱ���"
               Height          =   255
               Left            =   6120
               TabIndex        =   332
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label64 
               Caption         =   "�߿���"
               Height          =   255
               Left            =   4680
               TabIndex        =   330
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label35 
               Caption         =   "����"
               Height          =   255
               Left            =   6000
               TabIndex        =   328
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label54 
               Caption         =   "ͼ��λ��"
               Height          =   255
               Left            =   4680
               TabIndex        =   275
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label30 
               Caption         =   "����"
               Height          =   255
               Left            =   3360
               TabIndex        =   80
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label29 
               Caption         =   "ƽ��ģʽ"
               Height          =   255
               Left            =   1800
               TabIndex        =   79
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label28 
               Caption         =   "�Ŵ�ģʽ"
               Height          =   255
               Left            =   120
               TabIndex        =   78
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label27 
               Caption         =   "��Ƭ�ֱ���"
               Height          =   255
               Left            =   6120
               TabIndex        =   77
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label26 
               Caption         =   "Ƭ��"
               Height          =   255
               Left            =   3360
               TabIndex        =   76
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label25 
               Caption         =   "��Ƭ���"
               Height          =   255
               Left            =   1800
               TabIndex        =   75
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label24 
               Caption         =   "��Ƭ����"
               Height          =   255
               Left            =   120
               TabIndex        =   74
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label23 
               Caption         =   "��ӡ����"
               Height          =   255
               Left            =   4680
               TabIndex        =   73
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label22 
               Caption         =   "����"
               Height          =   255
               Left            =   3360
               TabIndex        =   72
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label21 
               Caption         =   "���ȼ�"
               Height          =   255
               Left            =   1800
               TabIndex        =   71
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label20 
               Caption         =   "��ʽ"
               Height          =   255
               Left            =   120
               TabIndex        =   70
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "�豸"
            Height          =   975
            Left            =   120
            TabIndex        =   63
            Top             =   0
            Width           =   6135
            Begin VB.TextBox txtPrinterName 
               Height          =   300
               Left            =   240
               TabIndex        =   1
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox txtAETitle 
               Height          =   300
               Left            =   1680
               TabIndex        =   2
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox txtIPAddress 
               Height          =   300
               Left            =   3120
               TabIndex        =   3
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox txtPort 
               Height          =   300
               Left            =   4560
               TabIndex        =   4
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label16 
               Caption         =   "��ӡ������"
               Height          =   255
               Left            =   240
               TabIndex        =   67
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label17 
               Caption         =   "AE����"
               Height          =   255
               Left            =   1680
               TabIndex        =   66
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label18 
               Caption         =   "IP��ַ"
               Height          =   255
               Left            =   3120
               TabIndex        =   65
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label19 
               Caption         =   "�˿ں�"
               Height          =   255
               Left            =   4560
               TabIndex        =   64
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Label Label12 
            Caption         =   "����AE"
            Height          =   255
            Left            =   6360
            TabIndex        =   203
            Top             =   203
            Width           =   615
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFPrinter 
         Height          =   1575
         Left            =   -74760
         TabIndex        =   61
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2778
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin VB.CommandButton cmdInfoLabelUpDown 
         Caption         =   "��"
         Height          =   855
         Index           =   2
         Left            =   -66240
         TabIndex        =   60
         Top             =   2880
         Width           =   350
      End
      Begin VB.CommandButton cmdInfoLabelUpDown 
         Caption         =   "��"
         Height          =   855
         Index           =   0
         Left            =   -66240
         Picture         =   "frmSysConfig.frx":1D55
         TabIndex        =   59
         Top             =   1560
         Width           =   350
      End
      Begin VB.CommandButton cmdDeSelInfoLabel 
         Caption         =   "<<ɾ��"
         Height          =   350
         Left            =   -71640
         TabIndex        =   58
         Top             =   3120
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelInfoLabel 
         Caption         =   ">>����"
         Height          =   350
         Index           =   3
         Left            =   -71640
         TabIndex        =   57
         Top             =   2160
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelInfoLabel 
         Caption         =   ">>����"
         Height          =   350
         Index           =   4
         Left            =   -71640
         TabIndex        =   56
         Top             =   1680
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelInfoLabel 
         Caption         =   ">>����"
         Height          =   350
         Index           =   2
         Left            =   -71640
         TabIndex        =   55
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelInfoLabel 
         Caption         =   ">>����"
         Height          =   350
         Index           =   1
         Left            =   -71640
         TabIndex        =   54
         Top             =   720
         Width           =   1100
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         Caption         =   "���Ͻ�"
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   4
         Left            =   -68280
         TabIndex        =   48
         Top             =   600
         Width           =   1935
         Begin VB.ListBox lstInfoLabelSel 
            Height          =   2040
            Index           =   4
            ItemData        =   "frmSysConfig.frx":1E8F
            Left            =   120
            List            =   "frmSysConfig.frx":1E91
            TabIndex        =   53
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         Caption         =   "���½�"
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   3
         Left            =   -68280
         TabIndex        =   47
         Top             =   3360
         Width           =   1935
         Begin VB.ListBox lstInfoLabelSel 
            Height          =   2040
            Index           =   3
            ItemData        =   "frmSysConfig.frx":1E93
            Left            =   120
            List            =   "frmSysConfig.frx":1E95
            TabIndex        =   52
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         Caption         =   "���½�"
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   2
         Left            =   -70200
         TabIndex        =   46
         Top             =   3360
         Width           =   1935
         Begin VB.ListBox lstInfoLabelSel 
            Height          =   2040
            Index           =   2
            ItemData        =   "frmSysConfig.frx":1E97
            Left            =   120
            List            =   "frmSysConfig.frx":1E99
            TabIndex        =   51
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         Caption         =   "���Ͻ�"
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   1
         Left            =   -70200
         TabIndex        =   45
         Top             =   600
         Width           =   1935
         Begin VB.ListBox lstInfoLabelSel 
            Height          =   2040
            Index           =   1
            ItemData        =   "frmSysConfig.frx":1E9B
            Left            =   120
            List            =   "frmSysConfig.frx":1E9D
            TabIndex        =   50
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.ListBox lstInfoLabelAll 
         Appearance      =   0  'Flat
         Height          =   3270
         ItemData        =   "frmSysConfig.frx":1E9F
         Left            =   -74880
         List            =   "frmSysConfig.frx":1EA1
         TabIndex        =   44
         Top             =   720
         Width           =   3015
      End
      Begin VB.CommandButton cmdLeftRight 
         Height          =   350
         Index           =   2
         Left            =   -71760
         Picture         =   "frmSysConfig.frx":1EA3
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "�ŵ�������"
         Top             =   3480
         Width           =   1100
      End
      Begin VB.CommandButton cmdLeftRight 
         Height          =   350
         Index           =   1
         Left            =   -71760
         Picture         =   "frmSysConfig.frx":22E5
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "�ŵ�����Ҽ�"
         Top             =   2160
         Width           =   1100
      End
      Begin VB.Frame Frame5 
         Caption         =   "����Ҽ�"
         Height          =   4815
         Index           =   1
         Left            =   -70560
         TabIndex        =   39
         Top             =   720
         Width           =   2895
         Begin VB.ListBox lstMouseKey 
            Height          =   4260
            Index           =   2
            ItemData        =   "frmSysConfig.frx":2727
            Left            =   120
            List            =   "frmSysConfig.frx":272E
            Style           =   1  'Checkbox
            TabIndex        =   43
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "������"
         Height          =   4815
         Index           =   0
         Left            =   -74760
         TabIndex        =   38
         Top             =   720
         Width           =   2895
         Begin VB.ListBox lstMouseKey 
            Height          =   4260
            Index           =   1
            ItemData        =   "frmSysConfig.frx":273F
            Left            =   120
            List            =   "frmSysConfig.frx":2746
            Style           =   1  'Checkbox
            TabIndex        =   42
            Top             =   240
            Width           =   2655
         End
      End
      Begin TabDlg.SSTab sstabModality 
         Height          =   5295
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   9340
         _Version        =   393216
         Style           =   1
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Ԥ�贰��λ"
         TabPicture(0)   =   "frmSysConfig.frx":2757
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(1)=   "cboWWModality"
         Tab(0).Control(2)=   "Frame28"
         Tab(0).Control(3)=   "cmdAddWWModality"
         Tab(0).Control(4)=   "cmdModifyWWModality"
         Tab(0).Control(5)=   "cmdWWWLApplyAll"
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Ԥ����Ļ����"
         TabPicture(1)   =   "frmSysConfig.frx":2773
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label9"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "frmSeriesLayout"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Frame2"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "cboLayoutModality"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "cmdAddLayoutModality"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "cmdModifyLayoutModality"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "cmdDelLayoutModality"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "cmdLayoutApplyAll"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "chkAutoSeriesLayout"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "chkAutoImageLayout"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "cboImageSort"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).ControlCount=   12
         TabCaption(2)   =   "ͼ������"
         TabPicture(2)   =   "frmSysConfig.frx":278F
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frmShutter"
         Tab(2).Control(1)=   "optShutter(1)"
         Tab(2).Control(2)=   "optShutter(0)"
         Tab(2).Control(3)=   "cmdShutterImgType(2)"
         Tab(2).Control(4)=   "cmdShutterImgType(0)"
         Tab(2).Control(5)=   "cmdShutterImgType(1)"
         Tab(2).Control(6)=   "cboImageShutter"
         Tab(2).Control(7)=   "Label52"
         Tab(2).ControlCount=   8
         Begin VB.ComboBox cboImageSort 
            Height          =   300
            ItemData        =   "frmSysConfig.frx":27AB
            Left            =   5880
            List            =   "frmSysConfig.frx":27C1
            Style           =   2  'Dropdown List
            TabIndex        =   349
            Top             =   1635
            Width           =   2205
         End
         Begin VB.CheckBox chkAutoImageLayout 
            Caption         =   "�Զ�ͼ�񲼾�"
            Height          =   255
            Left            =   360
            TabIndex        =   348
            Top             =   3405
            Width           =   1935
         End
         Begin VB.CheckBox chkAutoSeriesLayout 
            Caption         =   "�Զ����в���"
            Height          =   255
            Left            =   360
            TabIndex        =   347
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CommandButton cmdLayoutApplyAll 
            Caption         =   "ȫ��Ӧ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   7920
            TabIndex        =   320
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton cmdWWWLApplyAll 
            Caption         =   "ȫ��Ӧ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   -66960
            TabIndex        =   287
            Top             =   570
            Width           =   1095
         End
         Begin VB.Frame frmShutter 
            Caption         =   "ͼ������"
            Height          =   3615
            Left            =   -74760
            TabIndex        =   244
            Top             =   1560
            Width           =   8895
            Begin VB.CommandButton cmdShutterColor 
               Caption         =   "����"
               Height          =   350
               Index           =   1
               Left            =   7440
               TabIndex        =   271
               Top             =   2520
               Width           =   1100
            End
            Begin VB.CommandButton cmdShutterColor 
               Caption         =   "��ǳ"
               Height          =   350
               Index           =   0
               Left            =   7440
               TabIndex        =   270
               Top             =   2160
               Width           =   1100
            End
            Begin VB.CheckBox chkShutterType 
               Caption         =   "���������"
               Height          =   195
               Index           =   2
               Left            =   3960
               TabIndex        =   250
               Top             =   240
               Width           =   1335
            End
            Begin VB.CheckBox chkShutterType 
               Caption         =   "��������"
               Height          =   255
               Index           =   1
               Left            =   2160
               TabIndex        =   249
               Top             =   240
               Width           =   1095
            End
            Begin VB.CheckBox chkShutterType 
               Caption         =   "Բ������"
               Height          =   195
               Index           =   0
               Left            =   360
               TabIndex        =   248
               Top             =   240
               Width           =   1095
            End
            Begin VB.Frame frmShutterType 
               Enabled         =   0   'False
               Height          =   3255
               Index           =   2
               Left            =   3840
               TabIndex        =   247
               Top             =   240
               Width           =   3375
               Begin VB.CommandButton cmdVertices 
                  Caption         =   "ɾ������"
                  Height          =   350
                  Index           =   2
                  Left            =   2040
                  TabIndex        =   268
                  Top             =   2280
                  Width           =   1100
               End
               Begin VB.CommandButton cmdVertices 
                  Caption         =   "�޸Ķ���"
                  Height          =   350
                  Index           =   1
                  Left            =   2040
                  TabIndex        =   267
                  Top             =   1560
                  Width           =   1100
               End
               Begin VB.CommandButton cmdVertices 
                  Caption         =   "���Ӷ���"
                  Height          =   350
                  Index           =   0
                  Left            =   2040
                  TabIndex        =   266
                  Top             =   840
                  Width           =   1100
               End
               Begin VB.ListBox lstVertices 
                  Height          =   2400
                  Left            =   240
                  TabIndex        =   265
                  Top             =   360
                  Width           =   1575
               End
            End
            Begin VB.Frame frmShutterType 
               Enabled         =   0   'False
               Height          =   3255
               Index           =   1
               Left            =   2040
               TabIndex        =   246
               Top             =   240
               Width           =   1575
               Begin VB.TextBox txtRect 
                  Height          =   300
                  Index           =   3
                  Left            =   240
                  TabIndex        =   264
                  Text            =   "0"
                  Top             =   2760
                  Width           =   1095
               End
               Begin VB.TextBox txtRect 
                  Height          =   300
                  Index           =   2
                  Left            =   240
                  TabIndex        =   262
                  Text            =   "0"
                  Top             =   2040
                  Width           =   1095
               End
               Begin VB.TextBox txtRect 
                  Height          =   300
                  Index           =   1
                  Left            =   240
                  TabIndex        =   260
                  Text            =   "0"
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.TextBox txtRect 
                  Height          =   300
                  Index           =   0
                  Left            =   240
                  TabIndex        =   258
                  Text            =   "0"
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.Label Label53 
                  Caption         =   "�����±߽磺"
                  Height          =   255
                  Index           =   6
                  Left            =   240
                  TabIndex        =   263
                  Top             =   2520
                  Width           =   1215
               End
               Begin VB.Label Label53 
                  Caption         =   "�����ϱ߽磺"
                  Height          =   255
                  Index           =   5
                  Left            =   240
                  TabIndex        =   261
                  Top             =   1800
                  Width           =   1095
               End
               Begin VB.Label Label53 
                  Caption         =   "�����ұ߽磺"
                  Height          =   255
                  Index           =   4
                  Left            =   240
                  TabIndex        =   259
                  Top             =   1080
                  Width           =   1095
               End
               Begin VB.Label Label53 
                  Caption         =   "������߽磺"
                  Height          =   255
                  Index           =   3
                  Left            =   240
                  TabIndex        =   257
                  Top             =   360
                  Width           =   1095
               End
            End
            Begin VB.Frame frmShutterType 
               Enabled         =   0   'False
               Height          =   3255
               Index           =   0
               Left            =   240
               TabIndex        =   245
               Top             =   240
               Width           =   1575
               Begin VB.TextBox txtCircle 
                  Height          =   300
                  Index           =   2
                  Left            =   240
                  TabIndex        =   256
                  Text            =   "0"
                  Top             =   2640
                  Width           =   1095
               End
               Begin VB.TextBox txtCircle 
                  Height          =   300
                  Index           =   1
                  Left            =   240
                  TabIndex        =   254
                  Text            =   "0"
                  Top             =   1680
                  Width           =   1095
               End
               Begin VB.TextBox txtCircle 
                  Height          =   300
                  Index           =   0
                  Left            =   240
                  TabIndex        =   252
                  Text            =   "0"
                  Top             =   720
                  Width           =   1095
               End
               Begin VB.Label Label53 
                  Caption         =   "Բ�ΰ뾶��"
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   255
                  Top             =   2280
                  Width           =   1095
               End
               Begin VB.Label Label53 
                  Caption         =   "Բ��Y���꣺"
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   253
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.Label Label53 
                  Caption         =   "Բ��X���꣺"
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   251
                  Top             =   360
                  Width           =   1095
               End
            End
            Begin VB.Label Label53 
               Caption         =   "������ɫ"
               Height          =   255
               Index           =   7
               Left            =   7560
               TabIndex        =   269
               Top             =   480
               Width           =   855
            End
            Begin VB.Shape shpShutterColor 
               FillColor       =   &H8000000F&
               FillStyle       =   0  'Solid
               Height          =   900
               Left            =   7560
               Top             =   960
               Width           =   900
            End
         End
         Begin VB.OptionButton optShutter 
            Caption         =   "ʹ��ͼ������"
            Height          =   255
            Index           =   1
            Left            =   -72120
            TabIndex        =   243
            Top             =   1200
            Width           =   1935
         End
         Begin VB.OptionButton optShutter 
            Caption         =   "��ͼ������"
            Height          =   255
            Index           =   0
            Left            =   -74760
            TabIndex        =   242
            Top             =   1200
            Width           =   2535
         End
         Begin VB.CommandButton cmdShutterImgType 
            Caption         =   "ɾ������"
            Height          =   350
            Index           =   2
            Left            =   -69120
            TabIndex        =   240
            Top             =   570
            Width           =   1100
         End
         Begin VB.CommandButton cmdShutterImgType 
            Caption         =   "��������"
            Height          =   350
            Index           =   0
            Left            =   -71760
            TabIndex        =   238
            Top             =   570
            Width           =   1100
         End
         Begin VB.CommandButton cmdShutterImgType 
            Caption         =   "�޸�����"
            Height          =   350
            Index           =   1
            Left            =   -70440
            TabIndex        =   237
            Top             =   570
            Width           =   1100
         End
         Begin VB.CommandButton cmdDelLayoutModality 
            Caption         =   "ɾ������"
            Height          =   350
            Left            =   5880
            TabIndex        =   236
            Top             =   570
            Width           =   1100
         End
         Begin VB.CommandButton cmdModifyLayoutModality 
            Caption         =   "�޸�����"
            Height          =   350
            Left            =   4560
            TabIndex        =   235
            Top             =   570
            Width           =   1100
         End
         Begin VB.CommandButton cmdAddLayoutModality 
            Caption         =   "��������"
            Height          =   350
            Left            =   3240
            TabIndex        =   234
            Top             =   570
            Width           =   1100
         End
         Begin VB.CommandButton cmdModifyWWModality 
            Caption         =   "�޸�����"
            Height          =   350
            Left            =   -70560
            TabIndex        =   233
            Top             =   570
            Width           =   1100
         End
         Begin VB.CommandButton cmdAddWWModality 
            Caption         =   "��������"
            Height          =   350
            Left            =   -71760
            TabIndex        =   232
            Top             =   570
            Width           =   1100
         End
         Begin VB.Frame Frame28 
            Caption         =   "Ԥ�贰��λ��"
            Height          =   4095
            Left            =   -74880
            TabIndex        =   215
            Top             =   1080
            Width           =   9015
            Begin VB.Frame Frame1 
               Height          =   1575
               Left            =   120
               TabIndex        =   217
               Top             =   2400
               Width           =   8775
               Begin VB.CommandButton cmdWinWLDelete 
                  Caption         =   "ɾ��"
                  Height          =   345
                  Left            =   7170
                  TabIndex        =   226
                  Top             =   1080
                  Width           =   1100
               End
               Begin VB.CommandButton cmdWinWLUpdate 
                  Caption         =   "�޸�"
                  Height          =   345
                  Left            =   6000
                  TabIndex        =   225
                  Top             =   1080
                  Width           =   1100
               End
               Begin VB.CommandButton cmdWinWLAdd 
                  Caption         =   "����"
                  Height          =   345
                  Left            =   4800
                  TabIndex        =   224
                  Top             =   1080
                  Width           =   1100
               End
               Begin VB.TextBox txtWinLevel 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   1
                  EndProperty
                  Height          =   300
                  Left            =   6000
                  TabIndex        =   223
                  Text            =   "0"
                  Top             =   607
                  Width           =   1000
               End
               Begin VB.TextBox txtWinWidth 
                  BeginProperty DataFormat 
                     Type            =   0
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
                  Height          =   300
                  Left            =   4800
                  TabIndex        =   222
                  Text            =   "0"
                  Top             =   607
                  Width           =   1000
               End
               Begin VB.TextBox txtWinWLEName 
                  Height          =   300
                  Left            =   3000
                  TabIndex        =   221
                  Top             =   607
                  Width           =   1600
               End
               Begin VB.TextBox txtWinWLCName 
                  Height          =   300
                  Left            =   1200
                  TabIndex        =   220
                  Top             =   607
                  Width           =   1600
               End
               Begin VB.ComboBox cboFuncKey 
                  Height          =   315
                  ItemData        =   "frmSysConfig.frx":27FB
                  Left            =   120
                  List            =   "frmSysConfig.frx":281D
                  TabIndex        =   219
                  Top             =   600
                  Width           =   975
               End
               Begin VB.CheckBox chkDefaultWWWL 
                  Caption         =   "Ĭ�ϴ���λ"
                  Height          =   375
                  Left            =   7140
                  TabIndex        =   218
                  Top             =   540
                  Width           =   1425
               End
               Begin VB.Label Label7 
                  Caption         =   "��λ"
                  Height          =   255
                  Left            =   6000
                  TabIndex        =   231
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label Label6 
                  Caption         =   "����"
                  Height          =   255
                  Left            =   4800
                  TabIndex        =   230
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label Label5 
                  Caption         =   "Ӣ����"
                  Height          =   255
                  Left            =   3000
                  TabIndex        =   229
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label Label4 
                  Caption         =   "����λ����"
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   228
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label Label2 
                  Caption         =   "��ݼ�"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   227
                  Top             =   240
                  Width           =   735
               End
            End
            Begin MSFlexGridLib.MSFlexGrid MSFModality 
               Height          =   2295
               Left            =   120
               TabIndex        =   216
               TabStop         =   0   'False
               Top             =   240
               Width           =   8775
               _ExtentX        =   15478
               _ExtentY        =   4048
               _Version        =   393216
               FixedCols       =   0
               WordWrap        =   -1  'True
               SelectionMode   =   1
               AllowUserResizing=   1
               MousePointer    =   1
            End
         End
         Begin VB.ComboBox cboLayoutModality 
            Height          =   300
            ItemData        =   "frmSysConfig.frx":284C
            Left            =   1080
            List            =   "frmSysConfig.frx":285C
            Style           =   2  'Dropdown List
            TabIndex        =   213
            Top             =   600
            Width           =   1960
         End
         Begin VB.ComboBox cboWWModality 
            Height          =   300
            ItemData        =   "frmSysConfig.frx":2870
            Left            =   -73920
            List            =   "frmSysConfig.frx":2880
            Style           =   2  'Dropdown List
            TabIndex        =   211
            Top             =   600
            Width           =   1960
         End
         Begin VB.Frame Frame2 
            Height          =   1215
            Left            =   360
            TabIndex        =   34
            Top             =   3840
            Width           =   3135
            Begin VB.ListBox lstImageRows 
               Height          =   240
               ItemData        =   "frmSysConfig.frx":2894
               Left            =   360
               List            =   "frmSysConfig.frx":28B0
               TabIndex        =   36
               Top             =   720
               Width           =   855
            End
            Begin VB.ListBox lstImageCols 
               Height          =   240
               ItemData        =   "frmSysConfig.frx":28CC
               Left            =   1800
               List            =   "frmSysConfig.frx":28E8
               TabIndex        =   35
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label11 
               Caption         =   "����            ����"
               Height          =   255
               Left            =   360
               TabIndex        =   37
               Top             =   360
               Width           =   2535
            End
         End
         Begin VB.Frame frmSeriesLayout 
            Height          =   1215
            Left            =   360
            TabIndex        =   30
            Top             =   1635
            Width           =   3135
            Begin VB.ListBox lstSeriesCols 
               Height          =   240
               ItemData        =   "frmSysConfig.frx":2904
               Left            =   1800
               List            =   "frmSysConfig.frx":2920
               TabIndex        =   32
               Top             =   720
               Width           =   855
            End
            Begin VB.ListBox lstSeriesRows 
               Height          =   240
               ItemData        =   "frmSysConfig.frx":293C
               Left            =   360
               List            =   "frmSysConfig.frx":2958
               TabIndex        =   31
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label8 
               Caption         =   "����            ����"
               Height          =   255
               Left            =   360
               TabIndex        =   33
               Top             =   360
               Width           =   2535
            End
         End
         Begin VB.ComboBox cboImageShutter 
            Height          =   300
            ItemData        =   "frmSysConfig.frx":2974
            Left            =   -73920
            List            =   "frmSysConfig.frx":2984
            Style           =   2  'Dropdown List
            TabIndex        =   239
            Top             =   600
            Width           =   1960
         End
         Begin VB.Label Label9 
            Caption         =   "ͼ������"
            Height          =   210
            Left            =   4920
            TabIndex        =   350
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label52 
            Caption         =   "Ӱ������"
            Height          =   255
            Left            =   -74760
            TabIndex        =   241
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Ӱ������"
            Height          =   255
            Left            =   240
            TabIndex        =   214
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Ӱ������"
            Height          =   255
            Left            =   -74760
            TabIndex        =   212
            Top             =   660
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView livGetUserSetup 
         Height          =   4665
         Left            =   -74880
         TabIndex        =   293
         Top             =   480
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   8229
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "�û���"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "רҵ����ְ��"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "����ʾ�Ĳ�����Ϣ"
         Height          =   195
         Left            =   -74880
         TabIndex        =   49
         Top             =   480
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   8160
      TabIndex        =   27
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ӧ��(&A)"
      Height          =   350
      Left            =   6480
      TabIndex        =   26
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancle 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4800
      TabIndex        =   25
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3120
      TabIndex        =   0
      Top             =   6240
      Width           =   1100
   End
End
Attribute VB_Name = "frmSysConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------
'��  �ܣ���ϵͳ�������ô��������
'�����ˣ��ƽ�
'�������ڣ�2004.06.12
'���̺����嵥��
'    subChkWinWL������          ������������Ƿ�Ϸ����Ϸ����޸�ϵͳ����
'    subFillMSFModality������   �����ʾ����λ���õ��б�ؼ�
'    subInitModifiedLayout��������Ӱ�����в��ֵ��޸ļ�¼��ԭ
'    subKeepScreenLayout������  ��ʱ���汻�޸Ĺ�������û��Ӧ�õı�����Ļ����
'    subInitMouseUsage������    ������÷����õ��޸ĸ�ԭ
'    subFillMouseUsage������    �������÷����ý���Ŀؼ�
'    subSaveMouseUsage������    ��������÷��������õĽ�����浽ϵͳ���������ݿ�
'    subSetchkShiftState������  ���������ֵ�����ý��������shift��״̬����ʾ
'    subKeepMouseUsage������    ���汻�޸ģ�����û�б�Ӧ�õ�����÷��޸ġ�
'    subMoveLeftRight������     ��������ilst1���Ƶ�ilst1ָ���listbox�������ƶ������¼����
'    subFillInfoLabe������      ������ؼ��������Ľ���Ϣ��עλ�ú���ʾ����
'    subSaveInfoLabelLocate�������������Ľ���Ϣ��ʾ���õĽ�����浽��ͼ����Ϣ����ϵͳ�����С�
'    subFillMSFPrinter������    ��ϵͳ��Ϣ��д��MSF�����
'    funSavePrinterToPara������ ������ؼ�������ֵ���浽ָ����clsOnePrinterϵͳ������
'    subFillCboPrintFormat����������ӡ��ʽ�ؼ�
'    subFillCboFilmSize������   ��佺Ƭ���ؼ�
'    subFillUserInterface������ ���������ý���Ŀؼ�����
'    subSaveInterfacePara������ ���ݽ�����޸����������ϵͳ������ֵ������ϵͳ�������浽��Ӱ����������
'�޸ļ�¼��
'    2004.06.29     �ƽ�
'-------------------------------------------------------

Private ilstActive As Integer                   ''��¼��ǰ��ѡ�е��������listbox
Private ilstInfoLabelActvate As Integer         ''��¼��ǰ��ѡ�е��Ľ���Ϣlistbox
Private bMouseKeyShiftClick As Boolean           ''��¼��ǰ�����������shift����checkbox�ؼ���click�¼��Ƿ�Ӧ�÷���
Public f As frmViewer

Private Sub cboBorderDensity_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboEmptyDensity_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboFilmBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboFilmSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboFormat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboFuncKey_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboImageShutter_Click()
    Dim intModality As Integer
    Dim intCircle As Integer
    Dim intRect As Integer
    Dim intPolygon As Integer
    Dim intShutterType As Integer
    Dim aVertices() As String
    Dim strVertices As String
    Dim lngShutterColor As Long
    Dim i As Integer
    intModality = Me.cboImageShutter.ListIndex + 1
        
    '��������������Ѿ����޸ĵ�δ���棬����ʾ�Ѿ��޸ĵ����á�
    If aModifiedImageShutter(intModality).bModified Then
        intShutterType = aModifiedImageShutter(intModality).intShutterType
        strVertices = aModifiedImageShutter(intModality).strVertices
        lngShutterColor = aModifiedImageShutter(intModality).lngColor
    Else
        intShutterType = aImageShutter(intModality).intShutterType
        strVertices = aImageShutter(intModality).strVertices
        lngShutterColor = aImageShutter(intModality).lngColor
    End If
    '������������
    If intShutterType = 0 Or intShutterType > 7 Then
            Me.optShutter(0).Value = True
            Me.frmShutter.Enabled = False
        Else
            Me.optShutter(1).Value = True
            Me.frmShutter.Enabled = True
            
            If intShutterType >= 4 Then
                intShutterType = intShutterType - 4
                intPolygon = 1
            End If
            If intShutterType >= 2 Then
                intShutterType = intShutterType - 2
                intRect = 1
            End If
            If intShutterType >= 1 Then
                intCircle = 1
            End If
            
            Me.lstVertices.Clear
            If strVertices <> "" Then
                aVertices = Split(strVertices, ":")
                If UBound(aVertices) >= 5 And UBound(aVertices) Mod 2 = 1 Then
                    For i = 0 To UBound(aVertices) \ 2
                        lstVertices.AddItem "(" & Val(aVertices(i * 2)) & "," & Val(aVertices(i * 2 + 1)) & ")"
                    Next i
                Else
                    intPolygon = 0
                End If
            End If
            
            Me.chkShutterType(2).Value = intPolygon
            Me.frmShutterType(2).Enabled = intPolygon
            Me.chkShutterType(1).Value = intRect
            Me.frmShutterType(1).Enabled = intRect
            Me.chkShutterType(0).Value = intCircle
            Me.frmShutterType(0).Enabled = intCircle
            
            '�����ɫ
            lngShutterColor = lngShutterColor Mod 65536
            lngShutterColor = lngShutterColor \ 256
            Me.shpShutterColor.FillColor = RGB(Abs(lngShutterColor), Abs(lngShutterColor), Abs(lngShutterColor))
            
            If aModifiedImageShutter(intModality).bModified Then
                Me.txtCircle(0).Text = aModifiedImageShutter(intModality).intCenterX
                Me.txtCircle(1).Text = aModifiedImageShutter(intModality).intCenterY
                Me.txtCircle(2).Text = aModifiedImageShutter(intModality).intRadius
                Me.txtRect(0).Text = aModifiedImageShutter(intModality).intRectLeft
                Me.txtRect(1).Text = aModifiedImageShutter(intModality).intRectRight
                Me.txtRect(2).Text = aModifiedImageShutter(intModality).intRectUpper
                Me.txtRect(3).Text = aModifiedImageShutter(intModality).intRectLower
            Else
                Me.txtCircle(0).Text = aImageShutter(intModality).intCenterX
                Me.txtCircle(1).Text = aImageShutter(intModality).intCenterY
                Me.txtCircle(2).Text = aImageShutter(intModality).intRadius
                Me.txtRect(0).Text = aImageShutter(intModality).intRectLeft
                Me.txtRect(1).Text = aImageShutter(intModality).intRectRight
                Me.txtRect(2).Text = aImageShutter(intModality).intRectUpper
                Me.txtRect(3).Text = aImageShutter(intModality).intRectLower
            End If
        End If
End Sub

Private Sub cboImageSort_LostFocus()
    subKeepScreenLayout
End Sub

Private Sub cboMagnification_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboMedium_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboOrientation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboPolarity_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboPriority_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboResolution_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboSmooth_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboTrim_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub chkDefaultWWWL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub chkDockMiniImage_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub chkPatientiInfoFontBold_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub chkPatientInfoFontItalic_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub chkPrintFilmBeep_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub chkPrintOkEcho_LostFocus()
    blnPrintOkEcho = IIf(chkPrintOkEcho.Value = 1, True, False)
End Sub

Private Sub chkShowMiniImageInfo_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub chkShowMPRLine_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub chkShowPrintTag_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub chkShutterType_Click(Index As Integer)
    Me.frmShutterType(Index).Enabled = Me.chkShutterType(Index).Value
End Sub

Private Sub chkShutterType_LostFocus(Index As Integer)
    subKeepImageShutter
End Sub

Private Sub chkSquareFrame_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub cmdAddLayoutModality_Click()
    Dim strModality As String
    Dim i As Integer
    Dim strSQL As String
    Dim intModality As Integer
    
    '��ȡ��Ӱ����������
    strModality = funcGetNewLayoutModality
    If Len(Trim(strModality)) < 1 Then Exit Sub
    If zl9ComLib.zlCommFun.StrIsValid(strModality, 20, Me.hwnd, "Ԥ����Ļ����") = False Then
        Exit Sub
    End If
    '�����ڴ����
    intModality = Me.cboLayoutModality.ListCount
    intModality = intModality + 1
    ReDim Preserve aPresetLayout(intModality) As TModifiedPresetLayout
    ReDim Preserve aModifiedPresetLayout(intModality) As TModifiedPresetLayout
    aModifiedPresetLayout(intModality).strModality = strModality
    aPresetLayout(intModality).strModality = strModality
    aPresetLayout(intModality).bSeriesAutoFormat = IIf(Me.chkAutoSeriesLayout = 1, True, False)
    aPresetLayout(intModality).bImageAutoFormat = IIf(Me.chkAutoImageLayout = 1, True, False)
       
    '����cboLayoutModality������
    Me.cboLayoutModality.AddItem strModality
    Me.cboLayoutModality.ListIndex = Me.cboLayoutModality.ListCount - 1
    
    '���á��޸���𡱣���ɾ��Ӱ����𡱰�ť�Ŀ�����
    If cboLayoutModality.ListCount = 0 Then
        cmdModifyLayoutModality.Enabled = False
        cmdDelLayoutModality.Enabled = False
    Else
        cmdModifyLayoutModality.Enabled = True
        cmdDelLayoutModality.Enabled = True
    End If
    
    On Error GoTo errh
        
    '�������ݿ��¼
    If blLocalRun = True Then
        strSQL = "insert into Ӱ����Ļ����(Ӱ������) values('" & strModality & "')"
        cnAccess.Execute strSQL, , adCmdText
    Else
        strSQL = "ZL_Ӱ����Ļ����_����_INSERT(" & glngUserID & ",'" & strModality & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    subKeepScreenLayout
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Private Sub cmdAddWWModality_Click()
    Dim intModality As Integer
    Dim i As Integer
    Dim iFuncKey As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strMaxID As String
    Dim blnCreateData As Boolean
    
    '��������Ƿ�����
    
    iFuncKey = Me.cboFuncKey.ListIndex + 3
    If subChkWinWL(iFuncKey, False, True) = False Then
        MsgBox "������д������Ҫ�����ġ�Ԥ�贰��λ����Ϣ��" & vbCrLf & "Ȼ���ٵ�����������͡�������Ӱ�����͵�������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '���ݿ������Ӽ�¼
    intModality = UBound(aPresetWinWL, 2)
    Dim strSQL As String
    
    On Error GoTo errh
    
    If blLocalRun = True Then
        strSQL = "INSERT INTO Ӱ��Ԥ�贰��λ (Ӱ������,��ݼ�,��������,����Ӣ����,����,��λ,�Ƿ�Ĭ��) VALUES ('" & _
                 aPresetWinWL(iFuncKey, intModality).strModality & "'," & iFuncKey & ",'" & aPresetWinWL(iFuncKey, intModality).strWinWLCName & _
                 "','" & aPresetWinWL(iFuncKey, intModality).strWinWLEName & "'," & aPresetWinWL(iFuncKey, intModality).lngWinWidth & _
                 "," & aPresetWinWL(iFuncKey, intModality).lngWinLevel & "," & aPresetWinWL(iFuncKey, intModality).intDefault & ")"
        cnAccess.Execute strSQL, , adCmdText
        
        '��ȡ�µļ�¼id��
        strSQL = "select id from Ӱ��Ԥ�贰��λ WHERE Ӱ������ = '" & aPresetWinWL(iFuncKey, intModality).strModality & _
                 "' AND ��ݼ� = " & iFuncKey
        Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
        aPresetWinWL(iFuncKey, intModality).lngID = rsTmp!Id
    Else
        '�����û�����
        blnCreateData = CreateUserWWWL(glngUserID)
        
        strSQL = "ZL_Ӱ��Ԥ�贰��λ_INSERT(" & glngUserID & ",'" & aPresetWinWL(iFuncKey, intModality).strModality & "'," & _
                iFuncKey & ",'" & aPresetWinWL(iFuncKey, intModality).strWinWLCName & "','" & aPresetWinWL(iFuncKey, intModality).strWinWLEName & _
                "'," & aPresetWinWL(iFuncKey, intModality).lngWinWidth & "," & aPresetWinWL(iFuncKey, intModality).lngWinLevel & "," & _
                aPresetWinWL(iFuncKey, intModality).intDefault & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        '��ȡ�µļ�¼id��
        strSQL = "select id from Ӱ��Ԥ�贰��λ WHERE Ӱ������ = '" & aPresetWinWL(iFuncKey, intModality).strModality & _
                 "' AND ��ݼ� = " & iFuncKey
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
        aPresetWinWL(iFuncKey, intModality).lngID = Val(NVL(rsTmp!Id)) 'strMaxID
    End If
    
    '������������ݣ���ˢ���ڴ����
    If blnCreateData = True Then
        Call subGetWWWLToVal
        '������������б�
        Call subFillWWModality
    Else
        Me.cboWWModality.ListIndex = Me.cboWWModality.ListCount - 1
    End If

    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Private Sub CmdDefaultVal_Click()
    subGetInterfaceParaToVar 0
    subFillUserInterface
End Sub

Private Sub cmdDelLayoutModality_Click()
    Dim strModality As String
    Dim intModality As Integer
    Dim i As Integer
    Dim strSQL As String
    
    '�޸������ڴ����
    If Me.cboLayoutModality.ListIndex = -1 Then Exit Sub
    intModality = Me.cboLayoutModality.ListIndex + 1
    strModality = Me.cboLayoutModality.list(Me.cboLayoutModality.ListIndex)
    For i = intModality To UBound(aPresetLayout) - 1
        aPresetLayout(i) = aPresetLayout(i + 1)
        aModifiedPresetLayout(i) = aModifiedPresetLayout(i + 1)
    Next i
    ReDim Preserve aPresetLayout(UBound(aPresetLayout) - 1)
    ReDim Preserve aModifiedPresetLayout(UBound(aPresetLayout))
    '�޸������б�
    Me.cboLayoutModality.RemoveItem Me.cboLayoutModality.ListIndex
    
    '���á��޸���𡱣���ɾ��Ӱ����𡱰�ť�Ŀ�����
    If cboLayoutModality.ListCount = 0 Then
        cmdModifyLayoutModality.Enabled = False
        cmdDelLayoutModality.Enabled = False
    Else
        cmdModifyLayoutModality.Enabled = True
        cmdDelLayoutModality.Enabled = True
    End If
    
    On Error GoTo errh
    
    '�޸����ݿ�
    If blLocalRun = True Then
        strSQL = "delete from Ӱ����Ļ���� where Ӱ������='" & strModality & "'"
        cnAccess.Execute strSQL, , adCmdText
    Else
        strSQL = "ZL_Ӱ����Ļ����_DELETE(" & glngUserID & ",'" & strModality & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    Me.cboLayoutModality.ListIndex = Me.cboLayoutModality.ListCount - 1
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub

Private Sub CmdFilmFontSizeSetup_Click()
    frmFilmFontSize.Show vbModal, Me
End Sub

Private Sub cmdFilterAdd_Click()
'------------------------------------------------
'���ܣ������µĳ����˾�
'��������
'���أ�
'------------------------------------------------
    Dim strSQL As String
    Dim intCount As Integer
    
    On Error GoTo err
    '������������Ƿ�Ϸ������Ϸ�����ӵ�ϵͳ������
    If ValidateFilter = False Then Exit Sub
    
    '�����µ��˾�
    strSQL = "Zl_Ӱ���˾�ģ��_����(null,'" & txtFilterModality.Text & "','" & txtFilterName.Text & "'," _
             & Val(txtFilterPara(1).Text) & "," & Val(txtFilterPara(2).Text) & "," & Val(txtFilterPara(3).Text) _
             & "," & Val(txtFilterPara(4).Text) & "," & Val(txtFilterPara(5).Text) & "," & Val(txtFilterPara(6).Text) & ")"
    
    zlDatabase.ExecuteProcedure strSQL, "�������˾�"
    
    '���±���ϵͳ����
    Call subGetFilterToVal
    
    '�޸Ľ�����ʾ
    Call subFillMSFFilter
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdFilterDel_Click()
'------------------------------------------------
'���ܣ�ɾ���˾�����
'��������
'���أ�
'------------------------------------------------
    Dim iRow As Integer
    Dim strSQL As String
    
    On Error GoTo err
    
    If MSFFilter.Rows <= 1 Then Exit Sub
    
    '��ȡ��ǰ�˾���ID
    iRow = MSFFilter.Row - 1
    
     'ɾ���˾�
    strSQL = "Zl_Ӱ���˾�ģ��_ɾ��(" & aPresetFilter(iRow).lngID & ")"
    zlDatabase.ExecuteProcedure strSQL, "�������˾�"
    
    '���±���ϵͳ����
    Call subGetFilterToVal
    
    '�޸Ľ�����ʾ
    Call subFillMSFFilter
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ValidateFilter() As Boolean
'------------------------------------------------
'���ܣ�����˾�ģ���������Ч��
'��������
'���أ�True -- ������ȷ , False -- �������
'------------------------------------------------
    Dim i As Integer
    
    '������������Ƿ�Ϸ�
    
    ValidateFilter = False
    
    If txtFilterName.Text = "" Then
        MsgBox "�������˾����ơ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If txtFilterModality.Text = "" Then
        MsgBox "������Ӱ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    For i = 1 To 6
        If Val(txtFilterPara(i).Text) > 999 Then
            MsgBox "�˾�ģ���������999���������������0-999֮������֡� ", vbInformation, gstrSysName
            Exit Function
        End If
    Next i
    
    ValidateFilter = True
End Function
Private Sub cmdFilterUpdate_Click()
'------------------------------------------------
'���ܣ��޸��˾�����
'��������
'���أ�
'------------------------------------------------
    Dim iRow As Integer
    Dim strSQL As String
    
    On Error GoTo err
    
    If MSFFilter.Rows <= 1 Then Exit Sub
    
    '������������Ƿ�Ϸ������Ϸ�����ӵ�ϵͳ������
    If ValidateFilter = False Then Exit Sub
    
    '��ȡ��ǰ�˾���ID
    iRow = MSFFilter.Row - 1
    
    
     '�����µ��˾�
    strSQL = "Zl_Ӱ���˾�ģ��_����(" & aPresetFilter(iRow).lngID & ",'" & txtFilterModality.Text & "','" & txtFilterName.Text & "'," _
             & Val(txtFilterPara(1).Text) & "," & Val(txtFilterPara(2).Text) & "," & Val(txtFilterPara(3).Text) _
             & "," & Val(txtFilterPara(4).Text) & "," & Val(txtFilterPara(5).Text) & "," & Val(txtFilterPara(6).Text) & ")"
    
    zlDatabase.ExecuteProcedure strSQL, "�������˾�"
    
    '���±���ϵͳ����
    Call subGetFilterToVal
    
    '�޸Ľ�����ʾ
    Call subFillMSFFilter
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CmdGetUserInfo_Click()
    Dim lngGetUserID As Long
    If Me.livGetUserSetup.ListItems.Count <= 0 Then Exit Sub
    lngGetUserID = Val(Mid(Me.livGetUserSetup.ListItems(Me.livGetUserSetup.SelectedItem.Index).Key, 2))
    
    '���������
    Call subGetInterfaceParaToVar(lngGetUserID)
    Call subFillUserInterface
    '����÷�
    Call subGetMouseUsageToVar(lngGetUserID)
    Call subFillMouseUsage
    'ͼ��������
    Call subGetImageShutterToVar(lngGetUserID)
    Call subFillShutter
    '���к�ͼ�񲼾�
    Call subGetLayoutToVar(lngGetUserID)
    Call subFillLayoutModality
End Sub

Private Sub cmdInfoAdd_Click()
    Dim strSQL As String
    On Error GoTo errh
    If Me.txtUserLabelValue.Text <> "" Then
        If blLocalRun = True Then
            strSQL = "insert into Ӱ��ͼ����Ϣ��(��ʼ��ַ,������ַ,Ӣ������,��������,���ļ��,Ӣ�ļ��,����) values('2','2','cal','" _
                      & Me.txtUserLabelValue.Text & "','" & Me.txtUserLabelValue.Text & "','USER',True)"
            cnAccess.Execute strSQL, , adCmdText
        Else
            
            strSQL = "ZL_Ӱ��ͼ����Ϣ��_INSERT('2','2','cal','" _
                      & Me.txtUserLabelValue.Text & "','" & Me.txtUserLabelValue.Text & "','USER',-1)"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
        subGetInfoLabelToVar
        subFillInfoLabe
    End If
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Private Sub cmdInfoDelete_Click()
    Dim strSQL As String
    Dim intSel As Integer
    Dim intIndex As Integer
    
    '�ȱ��浱ǰ���ú�ɾ��
    cmdApply_Click
    
    intSel = Me.lstInfoLabelAll.ListIndex
    '��������Ƿ���Ч
    If intSel = -1 Then
        Me.cmdInfoAdd.Enabled = True
        Me.cmdInfoUpdate.Enabled = False
        Me.cmdInfoDelete.Enabled = False
        Exit Sub
    End If
    intIndex = Me.lstInfoLabelAll.ItemData(intSel)
    
    On Error GoTo errh
    
    If aInfoLabelLocate(intIndex).strElement = "2" And aInfoLabelLocate(intIndex).strGroup = "2" Then
        If blLocalRun = True Then
            strSQL = "delete from Ӱ��ͼ����Ϣ�� where id = " & aInfoLabelLocate(intIndex).lngID
            cnAccess.Execute strSQL, , adCmdText
        Else
            strSQL = "ZL_Ӱ��ͼ����Ϣ��_DELETE(" & aInfoLabelLocate(intIndex).lngID & ")"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
        subGetInfoLabelToVar
        subFillInfoLabe
    End If
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Private Sub cmdInfoUpdate_Click()
    Dim strSQL As String
    Dim intSel As Integer
    Dim intIndex As Integer
    
    intSel = Me.lstInfoLabelAll.ListIndex
    intIndex = Me.lstInfoLabelAll.ItemData(intSel)
    '��������Ƿ���Ч
    If intSel = -1 Then
        Me.cmdInfoAdd.Enabled = True
        Me.cmdInfoUpdate.Enabled = False
        Me.cmdInfoDelete.Enabled = False
        Exit Sub
    End If
    
    On Error GoTo errh
    
    If Me.txtUserLabelValue.Text <> "" And aInfoLabelLocate(intIndex).strElement = "2" _
       And aInfoLabelLocate(intIndex).strGroup = "2" Then
        If blLocalRun = True Then
            strSQL = "update ͼ����Ϣ�� set ���ļ�� = '" & Me.txtUserLabelValue.Text & "' where id = " _
                     & aInfoLabelLocate(intIndex).lngID
            cnAccess.Execute strSQL, , adCmdText
        Else
            strSQL = "ZL_Ӱ��ͼ����Ϣ��_UPDATE(" & aInfoLabelLocate(intIndex).lngID & ",'" & Me.txtUserLabelValue.Text & "')"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
        subGetInfoLabelToVar
        subFillInfoLabe
    End If
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Private Sub cmdLayoutApplyAll_Click()
'�ѵ�ǰ����Ļ��������ȫ��Ӧ�õ�Ĭ��������
    Dim strSQL As String
    
    On Error GoTo err
    
    '�ȱ����û��ĵ�ǰ���ã�Ȼ����ȫ��Ӧ��
    
    Call subSaveScreenLayout
    
    strSQL = "Zl_Ӱ����Ļ����_ApplyAll(" & glngUserID & ",'" & cboLayoutModality.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Exit Sub
err:
End Sub

Private Sub cmdModifyLayoutModality_Click()
    Dim strModality As String
    Dim strOldModality As String
    Dim intModality As Integer
    Dim strSQL As String
    
    If Me.cboLayoutModality.ListIndex = -1 Then Exit Sub
    '��ȡ��Ӱ����������
    strModality = funcGetNewLayoutModality
    If strModality = "" Then Exit Sub
    If zl9ComLib.zlCommFun.StrIsValid(strModality, 20, Me.hwnd, "Ԥ����Ļ����") = False Then Exit Sub
    '�޸������ڴ����
    intModality = Me.cboLayoutModality.ListIndex + 1
    strOldModality = aPresetLayout(intModality).strModality
    aPresetLayout(intModality).strModality = strModality
    aModifiedPresetLayout(intModality).strModality = strModality
    '�޸������б�
    Me.cboLayoutModality.list(Me.cboLayoutModality.ListIndex) = strModality
    
    On Error GoTo errh
    
    '�޸����ݿ�
    If blLocalRun = True Then
        strSQL = "UPDATE Ӱ����Ļ���� SET Ӱ������='" & strModality & "' where Ӱ������='" & strOldModality & "'"
        cnAccess.Execute strSQL, , adCmdText
    Else
        strSQL = "ZL_Ӱ����Ļ����_����_UPDATE(" & glngUserID & ",'" & strOldModality & "','" & strModality & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
        
End Sub

Private Function funGetNewShutterModality() As String
    Dim intModality As Integer
    Dim i As Integer
    '��ʾ������Ӱ����������
    funGetNewShutterModality = InputBox("��������Ӱ�����͵����ƣ�", gstrSysName)
    If funGetNewShutterModality = "" Then Exit Function
    '�ж�Ӱ�������Ƿ��ظ�
    intModality = UBound(aImageShutter)
    For i = 1 To intModality
        If UCase(funGetNewShutterModality) = UCase(aImageShutter(i).strModality) Then
            MsgBox "������Ӱ������Ѿ����ڣ�����������Ӱ�����͡�", vbInformation, gstrSysName
            funGetNewShutterModality = ""
            Exit Function
        End If
    Next i
End Function

Private Function funcGetNewLayoutModality() As String
    Dim intModality As Integer
    Dim i As Integer
    '��ʾ������Ӱ����������
    funcGetNewLayoutModality = InputBox("��������Ӱ�����͵����ƣ�", "����Ӱ������")
    If funcGetNewLayoutModality = "" Then Exit Function
    '�ж�Ӱ�������Ƿ��ظ�
    intModality = UBound(aPresetLayout)
    For i = 1 To intModality
        If UCase(funcGetNewLayoutModality) = UCase(aPresetLayout(i).strModality) Then
            MsgBox "������Ӱ������Ѿ����ڣ�����������Ӱ�����͡�", vbInformation, gstrSysName
            funcGetNewLayoutModality = ""
            Exit Function
        End If
    Next i
End Function

Private Sub cmdModifyWWModality_Click()
    Dim strModality As String
    Dim strOldModality As String
    Dim i As Integer
    Dim intModality As Integer
    Dim blnCreateData As Boolean
    
    '������Ӱ����������
    strModality = InputBox("��������Ӱ���������ƣ�", "�޸�Ӱ������")
    If strModality = "" Then Exit Sub
    '���Ӱ�����������Ƿ��ظ�
    For i = 1 To UBound(aPresetWinWL, 2)
        If UCase(strModality) = UCase(aPresetWinWL(3, i).strModality) Then
            MsgBox "������Ӱ������Ѿ����ڣ�����������Ӱ�����͡�", vbInformation, gstrSysName
            Exit Sub
        End If
    Next i
    
    If zl9ComLib.zlCommFun.StrIsValid(strModality, 20, Me.hwnd, "Ԥ�贰��λ") = False Then Exit Sub
    '�޸��ڴ����
    intModality = Me.cboWWModality.ListIndex + 1
    strOldModality = aPresetWinWL(3, intModality).strModality
    'F3������˼�¼Ӱ�����Ĺ��ܣ����Բ���������Ƿ�ʹ�ã���Ҫ�޸�Ӱ�����
    aPresetWinWL(3, intModality).strModality = strModality
    For i = 4 To 12
        If aPresetWinWL(i, intModality).bInUse Then
            aPresetWinWL(i, intModality).strModality = strModality
        End If
    Next i
    '�޸����ݿ�
    Dim strSQL As String
    
    On Error GoTo errh
    
    If blLocalRun = True Then
        strSQL = "UPDATE Ӱ��Ԥ�贰��λ SET Ӱ������='" & strModality & "' where Ӱ������='" & strOldModality & "'"
        cnAccess.Execute strSQL, , adCmdText
    Else
        '�����û�����
        blnCreateData = CreateUserWWWL(glngUserID)
        
        strSQL = "ZL_Ӱ�񴰿�λ_����_UPDATE(" & glngUserID & ",'" & strOldModality & "','" & strModality & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
    End If
    
    If blnCreateData = True Then
        '������������ݣ���ˢ���ڴ����
        Call subGetWWWLToVal
        '������������б�
        Call subFillWWModality
    Else
        '�޸�cboWWModality����ʾ
        Me.cboWWModality.list(Me.cboWWModality.ListIndex) = strModality
    End If
    
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Private Sub cmdPatientInfoFont_Click()
    Call SetPatientInfoFont
    blnInterfaceParaModified = True
End Sub

Private Sub cmdShutterColor_Click(Index As Integer)
    Dim lngColor As Long
    lngColor = Me.shpShutterColor.FillColor
    lngColor = lngColor Mod 256
    If Index = 1 Then       '����
        If lngColor > 6 And lngColor <= 255 Then
            lngColor = lngColor - 5
        Else
            lngColor = 1
        End If
    Else                    '��ǳ
        If lngColor < 250 And lngColor >= 1 Then
            lngColor = lngColor + 5
        Else
            lngColor = 255
        End If
    End If
    Me.shpShutterColor.FillColor = RGB(lngColor, lngColor, lngColor)
    subKeepImageShutter
End Sub

Private Sub cmdShutterImgType_Click(Index As Integer)
    Dim strModality As String
    Dim strOldModality As String
    Dim strSQL As String
    Dim i As Integer
    Dim intModality As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim MaxID As String
    
    On Error GoTo errh
    
    Select Case Index
        Case 0  '����Ӱ�����
            '��ȡ��Ӱ����������
            strModality = funGetNewShutterModality
            If Len(Trim(strModality)) < 1 Then Exit Sub
            If zl9ComLib.zlCommFun.StrIsValid(strModality, 20, Me.hwnd, "ͼ������") = False Then
                Exit Sub
            End If
            '�����ڴ����
            intModality = Me.cboImageShutter.ListCount
            intModality = intModality + 1
            ReDim Preserve aImageShutter(intModality) As TImageShutter
            ReDim Preserve aModifiedImageShutter(intModality) As TImageShutter
            aModifiedImageShutter(intModality).strModality = strModality
            aImageShutter(intModality).strModality = strModality
               
            '����cboLayoutModality������
            Me.cboImageShutter.AddItem strModality
            Me.cboImageShutter.ListIndex = Me.cboImageShutter.ListCount - 1
            
            '�������ݿ��¼
            If blLocalRun = True Then
                strSQL = "insert into Ӱ��ͼ��������(Ӱ������) values('" & strModality & "')"
                cnAccess.Execute strSQL, , adCmdText
            Else
                strSQL = "ZL_Ӱ��ͼ��������_����_INSERT(" & glngUserID & ",'" & strModality & "')"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
            subKeepScreenLayout
        Case 1  '�޸�Ӱ�����
            If Me.cboImageShutter.ListIndex = -1 Then Exit Sub
            '��ȡ��Ӱ����������
            strModality = funGetNewShutterModality
            If strModality = "" Then Exit Sub
            If zl9ComLib.zlCommFun.StrIsValid(strModality, 20, Me.hwnd, "Ԥ����Ļ����") = False Then Exit Sub
            '�޸������ڴ����
            intModality = Me.cboImageShutter.ListIndex + 1
            strOldModality = aImageShutter(intModality).strModality
            aImageShutter(intModality).strModality = strModality
            aModifiedImageShutter(intModality).strModality = strModality
            '�޸������б�
            Me.cboImageShutter.list(Me.cboImageShutter.ListIndex) = strModality
            
            If blLocalRun = True Then
                '�޸����ݿ�
                strSQL = "UPDATE Ӱ��ͼ�������� SET Ӱ������='" & strModality & "' where Ӱ������='" & strOldModality & "'"
                cnAccess.Execute strSQL, , adCmdText
            Else
                strSQL = "ZL_Ӱ��ͼ��������_����_Update(" & glngUserID & ",'" & strOldModality & "')"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        Case 2  'ɾ��Ӱ�����
            '�޸������ڴ����
            If Me.cboImageShutter.ListIndex = -1 Then Exit Sub
            intModality = Me.cboImageShutter.ListIndex + 1
            strModality = Me.cboImageShutter.list(Me.cboImageShutter.ListIndex)
            For i = intModality To UBound(aImageShutter) - 1
                aImageShutter(i) = aImageShutter(i + 1)
                aModifiedImageShutter(i) = aModifiedImageShutter(i + 1)
            Next i
            ReDim Preserve aImageShutter(UBound(aImageShutter) - 1)
            ReDim Preserve aModifiedImageShutter(UBound(aImageShutter))
            '�޸������б�
            Me.cboImageShutter.RemoveItem Me.cboImageShutter.ListIndex
            
            If blLocalRun = True Then
                '�޸����ݿ�
                strSQL = "delete from Ӱ��ͼ�������� where Ӱ������='" & strModality & "'"
                cnAccess.Execute strSQL, , adCmdText
            Else
                strSQL = "ZL_Ӱ��ͼ��������_DELETE(" & glngUserID & ",'" & strModality & "')"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
            Me.cboImageShutter.ListIndex = Me.cboImageShutter.ListCount - 1
    End Select
    
    '���á��޸���𡱺͡�ɾ����𡱵Ŀ�����
    If cboImageShutter.ListCount = 0 Then
        cmdShutterImgType(1).Enabled = False
        cmdShutterImgType(2).Enabled = False
    Else
        cmdShutterImgType(1).Enabled = True
        cmdShutterImgType(2).Enabled = True
    End If
    
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Private Sub cmdVertices_Click(Index As Integer)
    Dim strVertex As String
    Dim strX As String
    Dim strY As String
    Dim strOldX As String
    Dim strOldY As String
    Dim i As Integer
    
    Select Case Index
        Case 0  '���Ӷ���
            If Me.lstVertices.ListCount < 3 Then
                For i = 1 To 3 - Me.lstVertices.ListCount
                    strX = InputBox("���������ε�X�㣺", "ͼ������", "0")
                    strY = InputBox("���������ε�Y�㣺", "ͼ������", "0")
                    If strX <> "" And strY <> "" Then
                        strVertex = "(" & Val(strX) & "," & Val(strY) & ")"
                        Me.lstVertices.AddItem strVertex
                    Else
                        Exit Sub
                    End If
                Next i
            Else
                strX = InputBox("���������ε�X�㣺", "ͼ������", "0")
                strY = InputBox("���������ε�Y�㣺", "ͼ������", "0")
                If strX <> "" And strY <> "" Then
                    strVertex = "(" & Val(strX) & "," & Val(strY) & ")"
                    Me.lstVertices.AddItem strVertex
                Else
                    Exit Sub
                End If
            End If
        Case 1  '�޸Ķ���
            If Me.lstVertices.ListIndex = -1 Then Exit Sub
            strVertex = Me.lstVertices.list(Me.lstVertices.ListIndex)
            strOldX = Mid(strVertex, 2, InStr(strVertex, ",") - 2)
            strOldY = Mid(strVertex, InStr(strVertex, ",") + 1, Len(strVertex) - InStr(strVertex, ",") - 1)
            strX = InputBox("���������ε�X�㣺", "ͼ������", strOldX)
            strY = InputBox("���������ε�Y�㣺", "ͼ������", strOldY)
            If strX <> "" And strY <> "" Then
                strVertex = "(" & Val(strX) & "," & Val(strY) & ")"
                Me.lstVertices.list(Me.lstVertices.ListIndex) = strVertex
            Else
                Exit Sub
            End If
        Case 2  'ɾ������
            If Me.lstVertices.ListIndex = -1 Then Exit Sub
            If Me.lstVertices.ListCount <= 3 Then
                MsgBox "����ɾ�����㣬�����ͼ������������Ҫ�������㡣", vbInformation, gstrSysName
                Exit Sub
            End If
            Me.lstVertices.RemoveItem (Me.lstVertices.ListIndex)
            If Me.lstVertices.ListCount > 0 Then
                Me.lstVertices.ListIndex = 0
            End If
    End Select
    subKeepImageShutter
End Sub

Private Sub cmdExportInf_Click()
    Dim i As Integer
    Dim s As String
    
    On Error GoTo errHandle
    '����Ϣ��ע����Ϊδ��ѡ���״̬
    If ilstInfoLabelActvate = 0 Then
        Exit Sub
    End If
    Dim iSel As Integer
    iSel = Me.lstInfoLabelSel(ilstInfoLabelActvate).ListIndex
    If iSel <> -1 Then
        For i = 0 To Me.lstInfoLabelSel(ilstInfoLabelActvate).ListCount - 1
            If Me.lstInfoLabelSel(ilstInfoLabelActvate).Selected(i) Then
                s = Me.lstInfoLabelSel(ilstInfoLabelActvate).list(i)
                
                If Trim(s) <> "" Then
                    If InStr(1, s, "-�ɵ���") Then
                        Me.lstInfoLabelSel(ilstInfoLabelActvate).list(i) = Split(s, "-�ɵ���")(0)
                    Else
                        Me.lstInfoLabelSel(ilstInfoLabelActvate).list(i) = s & "-�ɵ���"
                    End If
                End If
            End If
        Next i
    End If
    
    bInfoLabelModified = True
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdWWWLApplyAll_Click()
'�ѵ�ǰ�Ĵ���λ����ȫ��Ӧ�õ�Ĭ��������
    Dim strSQL As String
    
    On Error GoTo err
    strSQL = "Zl_Ӱ��Ԥ�贰��λ_ApplyAll(" & glngUserID & ",'" & cboWWModality.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Exit Sub
err:
End Sub

Private Sub Command6_Click()
    Dim g As New DicomGlobal
    Dim result As Integer
    On Error GoTo err
    result = g.Echo(Me.txtIPAddress, Me.txtPort, Me.txtLocalAE, Me.txtAETitle)
    On Error GoTo 0
    If result = 0 Then
        MsgBox " ��֤���� " & result & " ���ӳɹ���", vbInformation, gstrSysName
    Else
        MsgBox " ��֤���� " & result & " ����ʧ�ܡ�", vbExclamation, gstrSysName
    End If
    Exit Sub
err:
    MsgBox "����ʧ�ܣ������������ӡ�", vbExclamation, gstrSysName
End Sub

Private Sub Form_Load()
    Dim i As Integer
    bMouseKeyShiftClick = False
    Me.lstTextoOff(1).Clear
    Me.lstTextoOff(2).Clear
    Me.lstRulerSize(1).Clear
    Me.lstRulerSize(2).Clear
    Me.lstRulerSize(3).Clear
    Me.lstRulerSize(4).Clear
    Me.lstRulerLineWidth.Clear
    For i = 1 To 100 'Step -1
        Me.lstTextoOff(1).AddItem i
        Me.lstTextoOff(2).AddItem i
        Me.lstRulerLineWidth.AddItem i
        Me.lstRulerSize(1).AddItem i
        Me.lstRulerSize(3).AddItem i
    Next
    
    For i = 1 To 700 'Step -1
        Me.lstRulerSize(2).AddItem i
        Me.lstRulerSize(4).AddItem i
    Next
    
    '�����ֹ���������Ĭ�Ϸ�ҳ
    cboMouseWheelRoll.Clear
    cboMouseWheelRoll.AddItem "ͼ��ҳ"
    cboMouseWheelRoll.AddItem "ͼ������"
    cboMouseWheelRoll.ListIndex = 0
    
    '��������ק������Ĭ������
    cboMouseWheelDrag.Clear
    cboMouseWheelDrag.AddItem "ͼ������"
    cboMouseWheelDrag.AddItem "ͼ������"
    cboMouseWheelDrag.AddItem "ͼ�����"
    cboMouseWheelDrag.ListIndex = 0
    
    lstNoSelectLineWidth.Clear
    lstSelectLineWidth.Clear
    lstImageIdentifierSize.Clear
    lstPeriodSize.Clear
    lstSpaceSize.Clear
    lstMaxAreaX.Clear
    lstMaxAreaY.Clear
    lstCellSpacing.Clear
    lstStatusBarFontSize.Clear
    For i = 1 To 100 ' Step -1
        lstNoSelectLineWidth.AddItem i
        lstSelectLineWidth.AddItem i
        lstImageIdentifierSize.AddItem i
        lstPeriodSize.AddItem i
        lstSpaceSize.AddItem i
        lstCellSpacing.AddItem i
        If i <= 40 Then
            lstStatusBarFontSize.AddItem i
        End If
    Next
    For i = 1 To 8
        lstMaxAreaX.AddItem i
        lstMaxAreaY.AddItem i
    Next
    
    '��ʼ��DICOM��ӡ�Ŀؼ����������COMBOBOX
    subFillCboPrintFormat
    subFillCboFilmSize
    Me.chkPrintOkEcho.Value = IIf(blnPrintOkEcho = True, 1, 0)     '��ӡ�ɹ�����ʾ
    sstabConfiguration.Tab = 0
'    Me.txtFilmFontSize = intFilmFontSize
    If blLocalRun = True Then
        CmdGetUserInfo.Enabled = False
    End If
    'ʵʼ��
    lstCopies.ListIndex = 9
    lstSeriesRows.ListIndex = 6
    lstSeriesCols.ListIndex = 6
    lstImageRows.ListIndex = 6
    lstImageCols.ListIndex = 6
    
    'Ȩ��
    If InStr(mstrPrivs, "��Ƭ����") <> 0 Then
        sstabConfiguration.TabEnabled(1) = True
        sstabConfiguration.TabEnabled(4) = True
        sstabConfiguration.TabEnabled(6) = True
        cmdWWWLApplyAll.Enabled = True
        cmdLayoutApplyAll.Enabled = True
    Else
        sstabConfiguration.TabEnabled(1) = False
        sstabConfiguration.TabEnabled(4) = False
        sstabConfiguration.TabEnabled(6) = False
        cmdWWWLApplyAll.Enabled = False
        cmdLayoutApplyAll.Enabled = False
    End If
End Sub

Private Sub Form_Resize()
    '�ڴ�����ʾʱ�������ݽ���ˢ�º���䣬����������Ϊ FixedDialog��ֻ������ʾʱ�Ŵ���resize�¼�
    
    Call subInitModifiedPara     '��ʼ��ϵͳ�������޸�����
    
    Call subFillLayoutModality   '����ϵͳ��������������к�ͼ�񲼾�
    Call subFillWWModality       '����ϵͳ����������䴰��λ
    Call subFillMSFFilter        '����ϵͳ������������˾�����
    Call subFillShutter          '����ϵͳ�������������������
    Call subFillMouseUsage       '����ϵͳ���������������÷����ÿؼ���ʾ
    Call subFillInfoLabe         'ʹ��ϵͳ������䲡�˱�ע��Ϣ�Ľ����ý���
    Call subFillMSFPrinter       '���DicomPrint����Ŀؼ�����
    Call subFillUserInterface    '���������ý���Ŀؼ�����
    Call subLoadUserInfo         '��������û���Ϣ
End Sub

Private Sub cmdApply_Click()
    Dim strRegPath As String
    
    '����ϵͳ���õ��޸�
    subSaveScreenLayout
    subSaveMouseUsage
    subSaveInfoLabelLocate
    subSaveInterfacePara                    '����Ӱ�������������ݿ�
    subSaveImgShutter                       '����ͼ���������õ����ݿ�
    subInitModifiedPara                     '��ʼ��ϵͳ�������޸ļ�¼
    subSaveParameters                       '����ϵͳ������Ĳ���
    
    '�ڽ�����Ӧ��ϵͳ���õ��޸�
    subInitSerial f             '���´����ָ���
    Call subResizeSeries(f)     '���µ�������Viewer����ʾ
    subUpdateIcon f             '������ʾ������
    
    f.sbStatusBar.Font.Size = IIf(intStatusBarFontSize < 1, 10, intStatusBarFontSize)
    strRegPath = "����ģ��\zlPacsCore"
    SaveSetting "ZLSOFT", strRegPath, "����AE", cstrPrintAE
    SaveSetting "ZLSOFT", strRegPath, "��Ƭ����", intFilmFontSize
    SaveSetting "ZLSOFT", strRegPath, "��ӡ�ɹ�����ʾ", blnPrintOkEcho
End Sub

Private Sub cmdCancle_Click()
    subInitModifiedPara
    Unload Me
End Sub

Private Sub subInitModifiedPara()
    subInitModifiedLayout
    subInitMouseUsage
    subInitModifiedImgShutter           '��ʼ��ͼ����������
    bInfoLabelModified = False
    blnInterfaceParaModified = False    '��ʼ��Ӱ�����������޸ı��
End Sub
Private Sub cmdOK_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub cmdWinWLAdd_Click()
'------------------------------------------------
'���ܣ�'�����µĴ���λ
'��������
'���أ�
'�ϼ���������̣��¼�
'�¼���������̣�
'���õ��ⲿ������
'�����ˣ��ƽ�
'------------------------------------------------
    '������������Ƿ�Ϸ������Ϸ�����ӵ�ϵͳ������
    '�жϿ�ݼ��Ƿ��Ѿ���ʹ��
    Dim iFuncKey As Integer
    Dim intModality As Integer
    Dim i As Integer
    Dim strMaxID As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnCreateData As Boolean
    
    iFuncKey = Me.cboFuncKey.ListIndex + 3
    If subChkWinWL(iFuncKey, True) = False Then Exit Sub
    
    intModality = Me.cboWWModality.ListIndex + 1
    
    If Len(Trim(cboWWModality.Text)) <= 0 Then
        MsgBox "����ѡ��Ӱ�����ͺ�������Ĭ�ϴ���λ!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Dim strSQL As String
    
    On Error GoTo errh
    
    If blLocalRun = True Then
        If aPresetWinWL(iFuncKey, intModality).intDefault = 1 Then '��ԭ�е�Ĭ�ϴ���λֵ���
            strSQL = "UPDATE Ӱ��Ԥ�贰��λ SET �Ƿ�Ĭ��=0 where Ӱ������='" & aPresetWinWL(iFuncKey, intModality).strModality & "'"
            cnAccess.Execute strSQL, , adCmdText
        End If
        
        strSQL = "INSERT INTO Ӱ��Ԥ�贰��λ (Ӱ������,��ݼ�,��������,����Ӣ����,����,��λ,�Ƿ�Ĭ��) VALUES ('" & _
                 cboWWModality & "'," & iFuncKey & ",'" & aPresetWinWL(iFuncKey, intModality).strWinWLCName & _
                 "','" & aPresetWinWL(iFuncKey, intModality).strWinWLEName & "'," & aPresetWinWL(iFuncKey, intModality).lngWinWidth & _
                 "," & aPresetWinWL(iFuncKey, intModality).lngWinLevel & "," & aPresetWinWL(iFuncKey, intModality).intDefault & ")"
        cnAccess.Execute strSQL, , adCmdText
        
        For i = 3 To 12
            If i <> iFuncKey Then aPresetWinWL(i, intModality).intDefault = 0
        Next i
        
        '��ȡ�µļ�¼id��
        strSQL = "select id from Ӱ��Ԥ�贰��λ WHERE Ӱ������ = '" & cboWWModality & _
                 "' AND ��ݼ� = " & iFuncKey
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
        aPresetWinWL(iFuncKey, intModality).lngID = rsTemp!Id
    Else
        '�����û�����
        blnCreateData = CreateUserWWWL(glngUserID)
        
        '�����û�����
        strSQL = "ZL_Ӱ��Ԥ�贰��λ_INSERT(" & glngUserID & ",'" & cboWWModality & _
                 "'," & iFuncKey & ",'" & aPresetWinWL(iFuncKey, intModality).strWinWLCName & "','" & aPresetWinWL(iFuncKey, intModality).strWinWLEName & _
                 "'," & aPresetWinWL(iFuncKey, intModality).lngWinWidth & "," & aPresetWinWL(iFuncKey, intModality).lngWinLevel & _
                 "," & aPresetWinWL(iFuncKey, intModality).intDefault & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        '������������ݣ���ˢ���ڴ����
        If blnCreateData = True Then
            Call subGetWWWLToVal
        Else
            For i = 3 To 12
                If i <> iFuncKey Then aPresetWinWL(i, intModality).intDefault = 0
            Next i
            
            '��ȡ�µļ�¼id��
            strSQL = "select id from Ӱ��Ԥ�贰��λ WHERE Ӱ������ = '" & cboWWModality & _
                     "' AND ��ݼ� = " & iFuncKey
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
            aPresetWinWL(iFuncKey, intModality).lngID = Val(NVL(rsTemp!Id)) 'strMaxID
        End If
    End If
    '�޸Ľ���Ŀؼ���ʾ
    subFillMSFModality intModality
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Private Function subChkWinWL(iFuncKey As Integer, bChkFuncKey As Boolean, Optional blnAddModality As Boolean = False) As Boolean
'------------------------------------------------
'���ܣ�������������Ƿ�Ϸ����Ϸ����޸�ϵͳ����
'������iFuncKey--��ݼ����룻bChkFuncKey--�Ƿ����ݼ�
'���أ�����Ƿ�Ϸ���True-�Ϸ���Fasle-���Ϸ���
'�ϼ���������̣�frmSysConfig.cmdWinWLUpdate_Click��frmSysConfig.cmdWinWLAdd_Click
'�¼���������̣���
'���õ��ⲿ������aPresetWinWL
'�����ˣ��ƽ�
'------------------------------------------------
    If iFuncKey < 3 Then
'        MsgBox "��ѡ��һ����ݷ�ʽ��", vbQuestion, "ZLPACS"
        cboFuncKey.SetFocus
        Exit Function
    End If
    Dim intModality As Integer
    Dim strModality As String
    Dim i As Integer
    
    intModality = Me.cboWWModality.ListIndex + 1
    If bChkFuncKey = True Then
        If aPresetWinWL(iFuncKey, intModality).bInUse Then
            MsgBox "����Ԥ�贰���Ѿ�ʹ���˿�ݼ� F" & CStr(iFuncKey), vbExclamation, gstrSysName
            subChkWinWL = False
            Exit Function
        End If
    End If
       
    If Len(Trim(txtWinWLCName)) < 1 Then
        MsgBox "������������!", vbExclamation, gstrSysName
        Me.txtWinWLCName.SelStart = 0
        Me.txtWinWLCName.SelLength = Len(Me.txtWinWLCName.Text)
        Me.txtWinWLCName.SetFocus
        Exit Function
    End If
    
    If Len(Trim(Me.txtWinWLEName)) < 1 Then
        MsgBox "������Ӣ����!", vbExclamation, gstrSysName
        Me.txtWinWLEName.SelStart = 0
        Me.txtWinWLEName.SelLength = Len(Me.txtWinWLEName.Text)
        Me.txtWinWLEName.SetFocus
        Exit Function
    End If
    
    
    If Val(Me.txtWinWidth.Text) = 0 Or Len(Trim(Me.txtWinWidth)) < 1 Then
        MsgBox "�����봰��!", vbExclamation, gstrSysName
        Me.txtWinWidth.SelStart = 0
        Me.txtWinWidth.SelLength = Len(Me.txtWinWidth.Text)
        Me.txtWinWidth.SetFocus
        Exit Function
    End If
    
    If Val(Me.txtWinLevel.Text) = 0 Or Len(Trim(Me.txtWinLevel)) < 1 Then
        MsgBox "�����봰λ!", vbExclamation, gstrSysName
        Me.txtWinLevel.SelStart = 0
        Me.txtWinLevel.SelLength = Len(Me.txtWinLevel.Text)
        Me.txtWinLevel.SetFocus
        Exit Function
    End If
        
    If zl9ComLib.zlCommFun.StrIsValid(txtWinWLCName.Text, 50, Me.hwnd, "��λ������") = False Then
        txtWinWLCName.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(txtWinWLEName.Text, 50, Me.hwnd, "��λӢ����") = False Then
        txtWinWLEName.SetFocus
        Exit Function
    End If
    
'    If Len(Me.txtWinWLCName.Text) > 50 Or Len(Me.txtWinWLEName.Text) > 50 Then GoTo err1
    
    If blnAddModality Then  '��������Ӱ������
        '��ʾ�û������µ�Ӱ�����
        strModality = InputBox("�������µ�Ӱ�����͡�", "����Ӱ������")
        If strModality = "" Then subChkWinWL = False:     Exit Function
        If zl9ComLib.zlCommFun.StrIsValid(strModality, 20, Me.hwnd, "Ӱ������") = False Then Exit Function
        '�ж���Ӱ������Ƿ��ظ�
        For i = 0 To Me.cboWWModality.ListCount - 1
            If UCase(strModality) = UCase(Me.cboWWModality.list(i)) Then
                MsgBox "������Ӱ������Ѿ����ڣ�����������Ӱ�����͡�", vbExclamation, gstrSysName
                subChkWinWL = False
                Exit Function
            End If
        Next i
        ReDim Preserve aPresetWinWL(3 To 12, UBound(aPresetWinWL, 2) + 1) As TPresetWinWL
        aPresetWinWL(3, UBound(aPresetWinWL, 2)).strModality = strModality
        intModality = UBound(aPresetWinWL, 2)
        Me.cboWWModality.AddItem strModality
    End If
    With aPresetWinWL(iFuncKey, intModality)
        .bInUse = True
        .strModality = strModality
        .lngWinLevel = Me.txtWinLevel
        .lngWinWidth = Me.txtWinWidth
        .strWinWLCName = Me.txtWinWLCName
        .strWinWLEName = Me.txtWinWLEName
        .intDefault = Me.chkDefaultWWWL.Value
    End With
    subChkWinWL = True
    Exit Function
err1:
    MsgBox "�����������ȷ�����顣", vbExclamation, gstrSysName
    subChkWinWL = False
    Exit Function
End Function

Private Sub cmdWinWLDelete_Click()
    'ɾ������λ
    Dim iFuncKey As Integer
    Dim intModality As Integer
    Dim lngTableID As Long
    Dim blnCreateData As Boolean
    Dim i As Integer
    
    With MSFModality
        If .Rows <= 1 Then Exit Sub
        iFuncKey = Mid(.TextMatrix(.RowSel, 0), 2)
    End With
    
    On Error GoTo errh
    
    intModality = Me.cboWWModality.ListIndex + 1
    Dim strSQL As String
    If blLocalRun = True Then
        strSQL = "DELETE FROM Ӱ��Ԥ�贰��λ WHERE ID=" & aPresetWinWL(iFuncKey, intModality).lngID
        cnAccess.Execute strSQL, , adCmdText
    Else
        '�����û�����
        blnCreateData = CreateUserWWWL(glngUserID)
        
        '������´������û����ݣ��ǾͶ�ȡ������Ҫ�޸ĵ�ID
        If blnCreateData = True Then
            strSQL = "select id from Ӱ��Ԥ�贰��λ where " & _
                " (Ӱ������,��ݼ�) In (Select Ӱ������,��ݼ� From Ӱ��Ԥ�贰��λ Where Id =[1]) " & _
                " And ��ԱID =[2] "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����λID", aPresetWinWL(iFuncKey, intModality).lngID, glngUserID)
            If rsTemp.EOF = False Then
                lngTableID = rsTemp!Id
            End If
        Else
            lngTableID = aPresetWinWL(iFuncKey, intModality).lngID
        End If
        
        strSQL = "ZL_Ӱ��Ԥ�贰��λ_DELETE(" & lngTableID & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    aPresetWinWL(iFuncKey, intModality).bInUse = False
    
    If blnCreateData = True Then
        '������������ݣ���ˢ���ڴ����
        Call subGetWWWLToVal
        '�޸Ľ���Ŀؼ���ʾ
        subFillMSFModality intModality
    Else
        '���ɾ���������һ�����ڣ���ô��Ҫˢ��aPresetWinWL
        For i = 3 To 12
            If aPresetWinWL(i, intModality).bInUse = True Then Exit For
        Next i
        If i = 13 Then
            Call subGetWWWLToVal
            '������������б�
            Call subFillWWModality
        Else
            '�޸Ľ���Ŀؼ���ʾ
            subFillMSFModality intModality
        End If
    End If
    
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Private Sub cmdWinWLUpdate_Click()
'------------------------------------------------
'���ܣ��޸Ĵ���λ
'��������
'���أ�
'�ϼ���������̣��¼�
'�¼���������̣�
'���õ��ⲿ������
'�����ˣ��ƽ�
'------------------------------------------------
    Dim iFuncKey As Integer, iOldFuncKey As Integer
    Dim iRow As Integer
    Dim intModality As Integer
    Dim i As Integer
    Dim blnCreateData As Boolean
    Dim lngTableID As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    iFuncKey = Me.cboFuncKey.ListIndex + 3
    With MSFModality
        If .Rows <= 1 Then Exit Sub
        iRow = .RowSel
        iOldFuncKey = Mid(.TextMatrix(iRow, 0), 2)
    End With
    
    If subChkWinWL(iFuncKey, Not iOldFuncKey = iFuncKey) = False Then   '��������Ƿ�Ϸ�������Ϸ���ֱ���޸�ϵͳ����
        Exit Sub
    End If
    
    intModality = Me.cboWWModality.ListIndex + 1
    
    On Error GoTo errh
    
    If blLocalRun = True Then
        If aPresetWinWL(iFuncKey, intModality).intDefault = 1 Then '��ԭ�е�Ĭ�ϴ���λֵ���
            strSQL = "UPDATE Ԥ�贰��λ SET �Ƿ�Ĭ��=0 where Ӱ������='" & aPresetWinWL(iFuncKey, intModality).strModality & "'"
            cnAccess.Execute strSQL, , adCmdText
        End If
        
        For i = 3 To 12
            If i <> iFuncKey Then aPresetWinWL(i, intModality).intDefault = 0
        Next i
        
        strSQL = "UPDATE Ӱ��Ԥ�贰��λ SET Ӱ������ = '" & cboWWModality & "', ��ݼ�=" & _
                 iFuncKey & ",�������� = '" & aPresetWinWL(iFuncKey, intModality).strWinWLCName & "',����Ӣ���� = '" & _
                 aPresetWinWL(iFuncKey, intModality).strWinWLEName & "', ����=" & aPresetWinWL(iFuncKey, intModality).lngWinWidth & _
                 ",��λ=" & aPresetWinWL(iFuncKey, intModality).lngWinLevel & ",�Ƿ�Ĭ��=" & aPresetWinWL(iFuncKey, intModality).intDefault & _
                 " WHERE ID=" & aPresetWinWL(iOldFuncKey, intModality).lngID
        cnAccess.Execute strSQL, , adCmdText
    Else
        '�����û�����
        blnCreateData = CreateUserWWWL(glngUserID)
        
        '�޸��û�����
        '������´������û����ݣ��ǾͶ�ȡ������Ҫ�޸ĵ�ID
        If blnCreateData = True Then
            strSQL = "select id from Ӱ��Ԥ�贰��λ where " & _
                " (Ӱ������,��ݼ�) In (Select Ӱ������,��ݼ� From Ӱ��Ԥ�贰��λ Where Id =[1]) " & _
                " And ��ԱID =[2] "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����λID", aPresetWinWL(iOldFuncKey, intModality).lngID, glngUserID)
            If rsTemp.EOF = False Then
                lngTableID = rsTemp!Id
            End If
        Else
            lngTableID = aPresetWinWL(iOldFuncKey, intModality).lngID
        End If
        
        For i = 3 To 12
            If i <> iFuncKey Then aPresetWinWL(i, intModality).intDefault = 0
        Next i
        
        strSQL = "ZL_Ӱ��Ԥ�贰��λ_UPDATE(" & lngTableID & "," & glngUserID & _
                 ",'" & cboWWModality.Text & "'," & iFuncKey & ",'" & aPresetWinWL(iFuncKey, intModality).strWinWLCName & _
                 "','" & aPresetWinWL(iFuncKey, intModality).strWinWLEName & "'," & aPresetWinWL(iFuncKey, intModality).lngWinWidth & _
                 "," & aPresetWinWL(iFuncKey, intModality).lngWinLevel & "," & aPresetWinWL(iFuncKey, intModality).intDefault & ")"
                         
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
    End If
    If iOldFuncKey <> iFuncKey Then
        aPresetWinWL(iOldFuncKey, intModality).bInUse = False
        aPresetWinWL(iFuncKey, intModality).bInUse = True
    End If
    
    If blnCreateData = True Then
        Call subGetWWWLToVal    '���¶�ȡ�������õ��ڴ����
    End If
    '�޸Ľ���Ŀؼ���ʾ
    subFillMSFModality intModality
    
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Private Sub cboLayoutModality_Click()
    Dim intModality As Integer
    intModality = Me.cboLayoutModality.ListIndex + 1
        
    '�������޸ĵ���δ�����Ĭ����Ļ����
    If aModifiedPresetLayout(intModality).bModified Then
        Me.chkAutoSeriesLayout = IIf(aModifiedPresetLayout(intModality).bSeriesAutoFormat, 1, 0)
        Me.lstSeriesCols = aModifiedPresetLayout(intModality).lngSeriesColumns
        Me.lstSeriesRows = aModifiedPresetLayout(intModality).lngSeriesRows
        Me.chkAutoImageLayout = IIf(aModifiedPresetLayout(intModality).bImageAutoFormat, 1, 0)
        Me.lstImageCols = aModifiedPresetLayout(intModality).lngImageColumns
        Me.lstImageRows = aModifiedPresetLayout(intModality).lngImageRows
        Me.cboImageSort.ListIndex = aModifiedPresetLayout(intModality).lngImageSort
    Else
        '��ϵͳ������ֵ��д������ؼ�
        Me.chkAutoSeriesLayout = IIf(aPresetLayout(intModality).bSeriesAutoFormat, 1, 0)
        Me.lstSeriesCols = aPresetLayout(intModality).lngSeriesColumns
        Me.lstSeriesRows = aPresetLayout(intModality).lngSeriesRows
        Me.chkAutoImageLayout = IIf(aPresetLayout(intModality).bImageAutoFormat, 1, 0)
        Me.lstImageCols = aPresetLayout(intModality).lngImageColumns
        Me.lstImageRows = aPresetLayout(intModality).lngImageRows
        Me.cboImageSort.ListIndex = aPresetLayout(intModality).lngImageSort
    End If
End Sub

Private Sub cboWWModality_Click()
    subFillMSFModality Me.cboWWModality.ListIndex + 1             ''���MSF���ݱ��
End Sub

Private Sub subFillMSFFilter()
'------------------------------------------------
'���ܣ������ʾ�˾����õ��б�ؼ�
'��������
'���أ��ޣ�ֱ�������ʾ�ؼ�
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
        
    '��ʼ���б�
    Me.MSFFilter.Rows = 1
    Me.MSFFilter.Cols = 8
    
    Me.MSFFilter.TextMatrix(0, 0) = "Ӱ�����"
    Me.MSFFilter.TextMatrix(0, 1) = "�˾�����"
    Me.MSFFilter.TextMatrix(0, 2) = "��ǿǿ������"
    Me.MSFFilter.TextMatrix(0, 3) = "��ǿǿ�ȼ���"
    Me.MSFFilter.TextMatrix(0, 4) = "��ǿ��������"
    Me.MSFFilter.TextMatrix(0, 5) = "��ǿ���ȼ���"
    Me.MSFFilter.TextMatrix(0, 6) = "ƽ������"
    Me.MSFFilter.TextMatrix(0, 7) = "ƽ������"
    
    For i = 2 To 5
        Me.MSFFilter.ColWidth(i) = 1200
    Next i
    
    '��ʼ������ؼ�
    Me.txtFilterName = ""
    Me.txtFilterModality = ""
    For i = 1 To 6
        Me.txtFilterPara(i) = 0
    Next i
    
    Me.MSFFilter.Rows = UBound(aPresetFilter) + 1
    For i = 1 To UBound(aPresetFilter)
        Me.MSFFilter.TextMatrix(i, 0) = aPresetFilter(i - 1).strModality
        Me.MSFFilter.TextMatrix(i, 1) = aPresetFilter(i - 1).strname
        Me.MSFFilter.TextMatrix(i, 2) = aPresetFilter(i - 1).intUnSharpEnhancementUp
        Me.MSFFilter.TextMatrix(i, 3) = aPresetFilter(i - 1).intUnSharpEnhancementDown
        Me.MSFFilter.TextMatrix(i, 4) = aPresetFilter(i - 1).intUnSharpLengthUp
        Me.MSFFilter.TextMatrix(i, 5) = aPresetFilter(i - 1).intUnSharpLengthDown
        Me.MSFFilter.TextMatrix(i, 6) = aPresetFilter(i - 1).intFilterLengthUp
        Me.MSFFilter.TextMatrix(i, 7) = aPresetFilter(i - 1).intFilterLengthDown
    Next i
    
    If Me.MSFFilter.Rows > 1 Then
        Me.MSFFilter.Row = 1
        Call MSFFilter_Click
    End If
 
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subFillMSFModality(iModalityNo As Integer)
'------------------------------------------------
'���ܣ������ʾ����λ���õ��б�ؼ�
'������iModalityNo--aPresetWinWL������Ӱ�����͵ı��
'���أ��ޣ�ֱ�������ʾ�ؼ�
'------------------------------------------------
    Dim i As Integer
    Me.MSFModality.Rows = 1
    Me.MSFModality.Cols = 6
    Me.MSFModality.TextMatrix(0, 0) = "��ݼ�"
    Me.MSFModality.TextMatrix(0, 1) = "��������"
    Me.MSFModality.TextMatrix(0, 2) = "����Ӣ����"
    Me.MSFModality.TextMatrix(0, 3) = "����"
    Me.MSFModality.TextMatrix(0, 4) = "��λ"
    Me.MSFModality.TextMatrix(0, 5) = "�Ƿ�Ĭ��"
    Dim lngRowPos As Long               '��¼��ǰ������
    
    Me.cboFuncKey.Text = ""
    Me.txtWinLevel = 0
    Me.txtWinWidth = 0
    Me.txtWinWLCName = ""
    Me.txtWinWLEName = ""
    Me.chkDefaultWWWL.Value = 0
    If UBound(aPresetWinWL, 2) < iModalityNo Then Exit Sub
    lngRowPos = 1
    For i = 3 To 12
        If aPresetWinWL(i, iModalityNo).bInUse Then
            '��д���ݱ��
            Me.MSFModality.Rows = Me.MSFModality.Rows + 1
            Me.MSFModality.TextMatrix(lngRowPos, 0) = "F" & CStr(i)
            Me.MSFModality.TextMatrix(lngRowPos, 1) = aPresetWinWL(i, iModalityNo).strWinWLCName
            Me.MSFModality.TextMatrix(lngRowPos, 2) = aPresetWinWL(i, iModalityNo).strWinWLEName
            Me.MSFModality.TextMatrix(lngRowPos, 3) = aPresetWinWL(i, iModalityNo).lngWinWidth
            Me.MSFModality.TextMatrix(lngRowPos, 4) = aPresetWinWL(i, iModalityNo).lngWinLevel
            Me.MSFModality.TextMatrix(lngRowPos, 5) = IIf(aPresetWinWL(i, iModalityNo).intDefault = 1, "��", "��")
            lngRowPos = lngRowPos + 1
        End If
    Next
End Sub

Private Sub lstCellSpacing_KeyPress(KeyAscii As Integer)
    'ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub lstCopies_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub lstDensity_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub lstImageIdentifierSize_KeyPress(KeyAscii As Integer)
    'ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub lstInfoLabelAll_Click()
    Dim intSel As Integer
    Dim intIndex As Integer
    intSel = Me.lstInfoLabelAll.ListIndex
    If intSel = -1 Then
        Me.cmdInfoAdd.Enabled = True
        Me.cmdInfoUpdate.Enabled = False
        Me.cmdInfoDelete.Enabled = False
        Exit Sub
    End If
    intIndex = Me.lstInfoLabelAll.ItemData(intSel)
    Me.txtUserLabelValue.Text = Me.lstInfoLabelAll.list(intSel)
    If aInfoLabelLocate(intIndex).strGroup = "2" And aInfoLabelLocate(intIndex).strElement = "2" Then
        '�û���Ϣ
        Me.lblInfoType = "�û���Ϣ"
        Me.txtUserLabelValue = aInfoLabelLocate(intIndex).strCName
        Me.cmdInfoAdd.Enabled = True
        Me.cmdInfoUpdate.Enabled = True
        Me.cmdInfoDelete.Enabled = True
    ElseIf aInfoLabelLocate(intIndex).strGroup = "3" And aInfoLabelLocate(intIndex).strElement = "3" Then
        '���ݿ���Ϣ
        Me.lblInfoType = "���ݿ���Ϣ"
        Me.txtUserLabelValue = ""
        Me.cmdInfoAdd.Enabled = True
        Me.cmdInfoUpdate.Enabled = False
        Me.cmdInfoDelete.Enabled = False
    Else
        'ϵͳ��Ϣ
        Me.lblInfoType = "ϵͳ��Ϣ"
        Me.txtUserLabelValue = ""
        Me.cmdInfoAdd.Enabled = True
        Me.cmdInfoUpdate.Enabled = False
        Me.cmdInfoDelete.Enabled = False
    End If
End Sub

Private Sub lstInfoLabelSel_Click(Index As Integer)
    Dim i As Integer
    '��¼��ǰ�Ļlistbox
    If ilstInfoLabelActvate <> Index Then
        For i = 1 To 4
            If i <> Index Then
                Me.lstInfoLabelSel(i).ListIndex = -1
            End If
        Next
        ilstInfoLabelActvate = Index
    End If
End Sub

Private Sub lstMaxAreaX_KeyPress(KeyAscii As Integer)
    'ֻ����������
    If InStr("012345678" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub lstMaxAreaY_KeyPress(KeyAscii As Integer)
    'ֻ����������
    If InStr("012345678" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub lstPeriodSize_KeyPress(KeyAscii As Integer)
    'ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub lstRulerSize_KeyPress(Index As Integer, KeyAscii As Integer)
    'ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub lstSelectLineWidth_KeyPress(KeyAscii As Integer)
    'ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub lstSpaceSize_KeyPress(KeyAscii As Integer)
    'ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub lstStatusBarFontSize_KeyPress(KeyAscii As Integer)
    'ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub lstStatusBarFontSize_Scroll()
    blnInterfaceParaModified = True
End Sub

Private Sub lstTextoOff_KeyPress(Index As Integer, KeyAscii As Integer)
    'ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub


Private Sub MSFFilter_Click()
'------------------------------------------------
'���ܣ��˾������б�ĵ����¼�
'������
'���أ�
'------------------------------------------------
    Dim iRow As Integer
    
    If MSFFilter.Rows <= 1 Then Exit Sub
    
    With MSFFilter
        iRow = .Row - 1
        Me.txtFilterModality = aPresetFilter(iRow).strModality
        Me.txtFilterName = aPresetFilter(iRow).strname
        Me.txtFilterPara(1) = aPresetFilter(iRow).intUnSharpEnhancementUp
        Me.txtFilterPara(2) = aPresetFilter(iRow).intUnSharpEnhancementDown
        Me.txtFilterPara(3) = aPresetFilter(iRow).intUnSharpLengthUp
        Me.txtFilterPara(4) = aPresetFilter(iRow).intUnSharpLengthDown
        Me.txtFilterPara(5) = aPresetFilter(iRow).intFilterLengthUp
        Me.txtFilterPara(6) = aPresetFilter(iRow).intFilterLengthDown
    End With
    
End Sub

Private Sub MSFModality_Click()
'------------------------------------------------
'���ܣ�����λ�б�ĵ����¼�
'������
'���أ�
'�ϼ���������̣�
'�¼���������̣�
'���õ��ⲿ������
'�����ˣ��ƽ�
'------------------------------------------------
    Dim iRow As Integer
    Dim iFuncKey As Integer
    Dim intModality As Long
    If MSFModality.Rows <= 1 Then Exit Sub
    intModality = Me.cboWWModality.ListIndex + 1
    With MSFModality
        iRow = .RowSel
        iFuncKey = Mid(.TextMatrix(iRow, 0), 2)
    End With
    Me.cboFuncKey.ListIndex = iFuncKey - 3
    Me.txtWinWLCName.Text = aPresetWinWL(iFuncKey, intModality).strWinWLCName
    Me.txtWinWLEName.Text = aPresetWinWL(iFuncKey, intModality).strWinWLEName
    Me.txtWinWidth.Text = aPresetWinWL(iFuncKey, intModality).lngWinWidth
    Me.txtWinLevel.Text = aPresetWinWL(iFuncKey, intModality).lngWinLevel
    Me.chkDefaultWWWL.Value = aPresetWinWL(iFuncKey, intModality).intDefault
End Sub

Private Sub subInitModifiedImgShutter()
'------------------------------------------------
'���ܣ���ͼ���������޸ļ�¼��ԭ
'��������
'���أ�
'�ϼ���������̣�
'�¼���������̣�
'���õ��ⲿ������
'�����ˣ��ƽ�
'------------------------------------------------
    Dim i As Integer
    For i = 1 To UBound(aModifiedImageShutter)
        aModifiedImageShutter(i).bModified = False
        aModifiedImageShutter(i).strModality = aImageShutter(i).strModality
        aModifiedImageShutter(i).intShutterType = aImageShutter(i).intShutterType
        aModifiedImageShutter(i).intCenterX = aImageShutter(i).intCenterX
        aModifiedImageShutter(i).intCenterY = aImageShutter(i).intCenterY
        aModifiedImageShutter(i).intRadius = aImageShutter(i).intRadius
        aModifiedImageShutter(i).intRectLeft = aImageShutter(i).intRectLeft
        aModifiedImageShutter(i).intRectRight = aImageShutter(i).intRectRight
        aModifiedImageShutter(i).intRectUpper = aImageShutter(i).intRectUpper
        aModifiedImageShutter(i).intRectLower = aImageShutter(i).intRectLower
        aModifiedImageShutter(i).strVertices = aImageShutter(i).strVertices
        aModifiedImageShutter(i).lngColor = aImageShutter(i).lngColor
    Next i
End Sub

Private Sub subInitModifiedLayout()
'------------------------------------------------
'���ܣ���Ӱ�����в��ֵ��޸ļ�¼��ԭ
'������
'���أ�
'�ϼ���������̣�
'�¼���������̣�
'���õ��ⲿ������
'�����ˣ��ƽ�
'------------------------------------------------
    Dim i As Integer
    For i = 1 To UBound(aModifiedPresetLayout)
        aModifiedPresetLayout(i).bModified = False
        aModifiedPresetLayout(i).strModality = aPresetLayout(i).strModality
        aModifiedPresetLayout(i).bImageAutoFormat = aPresetLayout(i).bImageAutoFormat
        aModifiedPresetLayout(i).bSeriesAutoFormat = aPresetLayout(i).bSeriesAutoFormat
        aModifiedPresetLayout(i).lngImageColumns = aPresetLayout(i).lngImageColumns
        aModifiedPresetLayout(i).lngImageRows = aPresetLayout(i).lngImageRows
        aModifiedPresetLayout(i).lngSeriesColumns = aPresetLayout(i).lngSeriesColumns
        aModifiedPresetLayout(i).lngSeriesRows = aPresetLayout(i).lngSeriesRows
        aModifiedPresetLayout(i).bInvert = aPresetLayout(i).bInvert
        aModifiedPresetLayout(i).bShowPatientInfo = aPresetLayout(i).bShowPatientInfo
        aModifiedPresetLayout(i).bAutoSelectReferenceLine = aPresetLayout(i).bAutoSelectReferenceLine
        aModifiedPresetLayout(i).bAutoSelectSeriesSyn = aPresetLayout(i).bAutoSelectSeriesSyn
        aModifiedPresetLayout(i).lngInterpolationMode = aPresetLayout(i).lngInterpolationMode
        aModifiedPresetLayout(i).lngImageSort = aPresetLayout(i).lngImageSort
    Next
End Sub

Private Sub chkAutoSeriesLayout_Click()
    Me.lstSeriesCols.Enabled = IIf(Me.chkAutoSeriesLayout = 1, False, True)
    Me.lstSeriesRows.Enabled = Me.lstSeriesCols.Enabled
End Sub

Private Sub chkAutoImageLayout_Click()
    Me.lstImageCols.Enabled = IIf(Me.chkAutoImageLayout = 1, False, True)
    Me.lstImageRows.Enabled = Me.lstImageCols.Enabled
End Sub

Private Sub subKeepImageShutter()
    '------------------------------------------------
'���ܣ���ʱ���汻�޸Ĺ�������û��Ӧ�õ�ͼ����������
'��������
'���أ���
'�ϼ���������̣�
'�¼���������̣�
'���õ��ⲿ������
'�����ˣ��ƽ�
'------------------------------------------------
    '��ͼ�������������޸�
    Dim intModality As Integer
    Dim strVertices As String
    Dim lngColor As Long
    Dim intShutterType As Integer
    Dim strTemp As String
    Dim i As Integer
    
    intModality = Me.cboImageShutter.ListIndex + 1
    '����ͼ����������
    If Me.optShutter(0).Value = True Then
        intShutterType = 0
    Else
        If Me.chkShutterType(0).Value = 1 Then intShutterType = intShutterType + 1
        If Me.chkShutterType(1).Value = 1 Then intShutterType = intShutterType + 2
        If Me.chkShutterType(2).Value = 1 Then intShutterType = intShutterType + 4
    End If
    '����ͼ��������ɫ
    lngColor = (Me.shpShutterColor.FillColor Mod 256) * 256
    '�������εĶ����ַ���
    If Me.lstVertices.ListCount >= 3 Then
        strTemp = Me.lstVertices.list(0)
            strVertices = Mid(strTemp, 2, InStr(strTemp, ",") - 2) & ":" _
                          & Mid(strTemp, InStr(strTemp, ",") + 1, Len(strTemp) - InStr(strTemp, ",") - 1)
        For i = 1 To Me.lstVertices.ListCount - 1
            strTemp = Me.lstVertices.list(i)
            strVertices = strVertices & ":" & Mid(strTemp, 2, InStr(strTemp, ",") - 2) & ":" _
                          & Mid(strTemp, InStr(strTemp, ",") + 1, Len(strTemp) - InStr(strTemp, ",") - 1)
        Next i
    End If
    aModifiedImageShutter(intModality).bModified = True
    aModifiedImageShutter(intModality).intShutterType = intShutterType
    aModifiedImageShutter(intModality).lngColor = lngColor
    aModifiedImageShutter(intModality).strVertices = strVertices
    aModifiedImageShutter(intModality).intCenterX = Val(Me.txtCircle(0).Text)
    aModifiedImageShutter(intModality).intCenterY = Val(Me.txtCircle(1).Text)
    aModifiedImageShutter(intModality).intRadius = Val(Me.txtCircle(2).Text)
    aModifiedImageShutter(intModality).intRectLeft = Val(Me.txtRect(0).Text)
    aModifiedImageShutter(intModality).intRectRight = Val(Me.txtRect(1).Text)
    aModifiedImageShutter(intModality).intRectUpper = Val(Me.txtRect(2).Text)
    aModifiedImageShutter(intModality).intRectLower = Val(Me.txtRect(3).Text)
End Sub

Private Sub subKeepScreenLayout()
'------------------------------------------------
'���ܣ���ʱ���汻�޸Ĺ�������û��Ӧ�õı�����Ļ����
'��������
'���أ���
'�ϼ���������̣�
'�¼���������̣�
'���õ��ⲿ������
'�����ˣ��ƽ�
'------------------------------------------------
    '��¼��Ļ���ֵ��޸�
    Dim intModality As Integer
    
    intModality = Me.cboLayoutModality.ListIndex + 1
    If (Me.chkAutoSeriesLayout <> IIf(aPresetLayout(intModality).bSeriesAutoFormat = True, 1, 0)) Or _
       (Me.lstSeriesCols.list(Me.lstSeriesCols.TopIndex) <> aPresetLayout(intModality).lngSeriesColumns) Or _
       (Me.lstSeriesRows.list(Me.lstSeriesRows.TopIndex) <> aPresetLayout(intModality).lngSeriesRows) Or _
       (Me.chkAutoImageLayout <> IIf(aPresetLayout(intModality).bImageAutoFormat = True, 1, 0)) Or _
       (Me.lstImageCols.list(Me.lstImageCols.TopIndex) <> aPresetLayout(intModality).lngImageColumns) Or _
       (Me.lstImageRows.list(Me.lstImageRows.TopIndex) <> aPresetLayout(intModality).lngImageRows) Or _
       (Me.cboImageSort.ListIndex <> aPresetLayout(intModality).lngImageSort) Then
       
       aModifiedPresetLayout(intModality).bModified = True
       aModifiedPresetLayout(intModality).bSeriesAutoFormat = IIf(Me.chkAutoSeriesLayout = 1, True, False)
       aModifiedPresetLayout(intModality).lngSeriesColumns = Me.lstSeriesCols.list(Me.lstSeriesCols.TopIndex)
       aModifiedPresetLayout(intModality).lngSeriesRows = Me.lstSeriesRows.list(Me.lstSeriesRows.TopIndex)
       aModifiedPresetLayout(intModality).bImageAutoFormat = IIf(Me.chkAutoImageLayout = 1, True, False)
       aModifiedPresetLayout(intModality).lngImageColumns = Me.lstImageCols.list(Me.lstImageCols.TopIndex)
       aModifiedPresetLayout(intModality).lngImageRows = Me.lstImageRows.list(Me.lstImageRows.TopIndex)
       aModifiedPresetLayout(intModality).lngImageSort = Me.cboImageSort.ListIndex
    End If
End Sub

Private Sub chkAutoImageLayout_LostFocus()
    subKeepScreenLayout
End Sub

Private Sub chkAutoSeriesLayout_LostFocus()
    subKeepScreenLayout
End Sub

Private Sub lstImageCols_Scroll()
    subKeepScreenLayout
End Sub

Private Sub lstImageRows_Scroll()
    subKeepScreenLayout
End Sub

Private Sub lstSeriesCols_Scroll()
    subKeepScreenLayout
End Sub

Private Sub lstSeriesRows_Scroll()
    subKeepScreenLayout
End Sub

Private Sub optInterpolationMode_LostFocus(Index As Integer)
    subKeepScreenLayout
End Sub

Private Sub subInitMouseUsage()
'------------------------------------------------
'���ܣ�������÷����õ��޸ĸ�ԭ
'��������
'���أ���
'�ϼ���������̣�frmSysConfig.cmdApply_Click��frmSysConfig.cmdCancle_Click
'�¼���������̣���
'���õ��ⲿ��������
'�����ˣ��ƽ�
'------------------------------------------------
    Dim i  As Integer
    Dim clsOneMouseUsage As clsMouseUsage
    
    For i = 1 To cModifiedMouseUsage.Count
        cModifiedMouseUsage.Remove 1
    Next
    For i = 1 To cMouseUsage.Count
        Set clsOneMouseUsage = New clsMouseUsage
        clsOneMouseUsage.bModified = False
        clsOneMouseUsage.bSelected = cMouseUsage(i).bSelected
        clsOneMouseUsage.lngFuncNo = cMouseUsage(i).lngFuncNo
        clsOneMouseUsage.lngMouseKey = cMouseUsage(i).lngMouseKey
        clsOneMouseUsage.lngShift = cMouseUsage(i).lngShift
        clsOneMouseUsage.strProgramName = cMouseUsage(i).strProgramName
        clsOneMouseUsage.strShowName = cMouseUsage(i).strShowName
        cModifiedMouseUsage.Add clsOneMouseUsage, CStr(clsOneMouseUsage.lngFuncNo)
    Next
    bMouseUsageModified = False
End Sub

Private Sub subFillMouseUsage()
'------------------------------------------------
'���ܣ��������÷����ý���Ŀؼ���
'��������
'���أ���
'�ϼ���������̣�frmSysConfig.Form_Resize
'�¼���������̣�frmSysConfig.subSetchkShiftState
'���õ��ⲿ������cMouseUsage
'�����ˣ��ƽ�
'------------------------------------------------
    '���listbox�е�ԭ����Ϣ
    Me.lstMouseKey(1).Clear
    Me.lstMouseKey(2).Clear
    Me.lstMouseKey(1).ListIndex = -1
    Me.lstMouseKey(2).ListIndex = -1
    Dim i As Integer
    Dim iMouseKey As Integer
    '��������ϸ��ؼ�����ʾ
    For i = 1 To cMouseUsage.Count
        If cMouseUsage(i).lngFuncNo > lngDrawLabelFuncNo - 1 Then
            iMouseKey = cMouseUsage(i).lngMouseKey
            Me.lstMouseKey(iMouseKey).AddItem cMouseUsage(i).strShowName
            '��䵱ǰ��ѡ���������Ҽ�����
            Me.lstMouseKey(iMouseKey).Selected(Me.lstMouseKey(iMouseKey).NewIndex) = cMouseUsage(i).bSelected
            Me.lstMouseKey(iMouseKey).ItemData(Me.lstMouseKey(iMouseKey).NewIndex) = i
        End If
    Next
    Me.lstMouseKey(1).ListIndex = 0
    Me.lstMouseKey(2).ListIndex = -1
    ilstActive = 1
    
    If Me.lstMouseKey(1).ListCount > 0 Then
        Me.chkShiftState(1).Tag = Me.lstMouseKey(1).ItemData(0)
    End If
    
    subSetchkShiftState cMouseUsage(1).lngShift
End Sub

Public Sub subSaveMouseUsage()
'------------------------------------------------
'���ܣ���������÷��������õĽ�����浽ϵͳ���������ݿ⣬�ж��Ƿ���ֵ�����ı䣬���иı��򱣴�
'��������
'���أ���
'�ϼ���������̣�frmSysConfig.cmdApply_Click
'�¼���������̣���
'���õ��ⲿ������cMouseUsage
'�����ˣ��ƽ�
'------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    If bMouseUsageModified = False Then
        Exit Sub
    End If
    Dim iDrawLabelIndex As Integer
    iDrawLabelIndex = -1
    
    On Error GoTo errh
    
    If blLocalRun = True Then
        strSQL = "select ID,ֱ��,����,��Բ,��ͷ,�����,�����,�Ƕ�,����,����λ,����λ,����,����,�ü�_��ע����,����Ӧ����," & _
                 "��ά���,����ע from Ӱ����갴ť���� "
        Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
        If rsTmp.EOF = True Then
            '--Ӱ����갴ť����
            strSQL = "INSERT INTO Ӱ����갴ť���� (ID,ֱ��,����,��Բ,��ͷ,�����,�����,�Ƕ�,����,����λ,����λ,����,����,�ü�_��ע����,����Ӧ����,��ά���,����ע)" & _
                "VALUES (0,'1,1,0,0,miLabelline,343','2,1,0,0,miLabelRectangle,344','3,1,0,0,miLabelEllipse,339','4,1,0,0,miLabelArrowhead,338','5,1,0,0,miLabelPolygon,342','6,1,0,0,miLabelPolyLine,341','7,1,0,0,miLabelAngle,340','8,1,0,0,miLabeltext,337','101,1,0,0,miStack,308','102,1,0,0,miWidthLevel,314','103,1,0,0,miCruise,309','104,1,0,0,miZoom,311','201,1,0,0,No,0','105,1,0,0,miAutoWidthLevel,315','106,1,0,0,mi3dCursor,321','20,1,0,0,no,0');"
            cnAccess.Execute strSQL
            strSQL = "select ID,ֱ��,����,��Բ,��ͷ,�����,�����,�Ƕ�,����,����λ,����λ,����,����,�ü�_��ע����,����Ӧ����," & _
                 "��ά���,����ע from Ӱ����갴ť���� "
            Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
        End If
        strSQL = "update Ӱ����갴ť���� Set "
                 
        For i = 1 To cModifiedMouseUsage.Count
            
            If i <= 8 Then
                cMouseUsage(i).bSelected = cModifiedMouseUsage(cModifiedMouseUsage.Count).bSelected
                cMouseUsage(i).lngFuncNo = cModifiedMouseUsage(i).lngFuncNo
                cMouseUsage(i).lngMouseKey = cModifiedMouseUsage(cModifiedMouseUsage.Count).lngMouseKey
                cMouseUsage(i).lngShift = cModifiedMouseUsage(cModifiedMouseUsage.Count).lngShift
                cMouseUsage(i).strProgramName = cModifiedMouseUsage(i).strProgramName
                cMouseUsage(i).strShowName = cModifiedMouseUsage(i).strShowName
            Else
                '���浽ϵͳ����
                cMouseUsage(i).bSelected = cModifiedMouseUsage(i).bSelected
                cMouseUsage(i).lngFuncNo = cModifiedMouseUsage(i).lngFuncNo
                cMouseUsage(i).lngMouseKey = cModifiedMouseUsage(i).lngMouseKey
                cMouseUsage(i).lngShift = cModifiedMouseUsage(i).lngShift
                cMouseUsage(i).strProgramName = cModifiedMouseUsage(i).strProgramName
                cMouseUsage(i).strShowName = cModifiedMouseUsage(i).strShowName
            End If
            If i <> 1 Then
                strSQL = strSQL & ","
            End If
            strSQL = strSQL & " " & cMouseUsage(i).strShowName & " = '" & cMouseUsage(i).lngFuncNo & "," & cMouseUsage(i).lngMouseKey & "," & _
            cMouseUsage(i).lngShift & "," & cMouseUsage(i).bSelected & "," & cMouseUsage(i).strProgramName & "," & cMouseUsage(i).ButtomID & "'"
        Next
        
        strSQL = strSQL & " where id = 0 "
        cnAccess.Execute strSQL
    Else
        strSQL = "select ��ԱID,ֱ��,����,��Բ,��ͷ,�����,�����,�Ƕ�,����,����λ,����λ,����,����,�ü�_��ע����,����Ӧ����," & _
                 "��ά���,����ע from Ӱ����갴ť���� where ��Աid = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngUserID)
        If rsTmp.EOF = True Then
            strSQL = "select ��ԱID,ֱ��,����,��Բ,��ͷ,�����,�����,�Ƕ�,����,����λ,����λ,����,����,�ü�_��ע����,����Ӧ����," & _
                 "��ά���,����ע from Ӱ����갴ť���� where ��Աid = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(0))
            If rsTmp.EOF <> True Then
                strSQL = "ZL_Ӱ����갴ť����_INSERT('" & glngUserID
                For i = 1 To rsTmp.Fields.Count - 1
                    strSQL = strSQL & "','" & rsTmp(i).Value
                Next
                strSQL = strSQL & "')"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        End If
        
        strSQL = "ZL_Ӱ����갴ť����_UPDATE('" & glngUserID
        For i = 1 To cModifiedMouseUsage.Count
            
            If i <= 8 Then
                cMouseUsage(i).bSelected = cModifiedMouseUsage(cModifiedMouseUsage.Count).bSelected
                cMouseUsage(i).lngFuncNo = cModifiedMouseUsage(i).lngFuncNo
                cMouseUsage(i).lngMouseKey = cModifiedMouseUsage(cModifiedMouseUsage.Count).lngMouseKey
                cMouseUsage(i).lngShift = cModifiedMouseUsage(cModifiedMouseUsage.Count).lngShift
                cMouseUsage(i).strProgramName = cModifiedMouseUsage(i).strProgramName
                cMouseUsage(i).strShowName = cModifiedMouseUsage(i).strShowName
            Else
                '���浽ϵͳ����
                cMouseUsage(i).bSelected = cModifiedMouseUsage(i).bSelected
                cMouseUsage(i).lngFuncNo = cModifiedMouseUsage(i).lngFuncNo
                cMouseUsage(i).lngMouseKey = cModifiedMouseUsage(i).lngMouseKey
                cMouseUsage(i).lngShift = cModifiedMouseUsage(i).lngShift
                cMouseUsage(i).strProgramName = cModifiedMouseUsage(i).strProgramName
                cMouseUsage(i).strShowName = cModifiedMouseUsage(i).strShowName
            End If
            strSQL = strSQL & "','" & cMouseUsage(i).lngFuncNo & "," & cMouseUsage(i).lngMouseKey & "," & cMouseUsage(i).lngShift & _
                     "," & cMouseUsage(i).bSelected & "," & cMouseUsage(i).strProgramName & "," & cMouseUsage(i).ButtomID
        Next
        
        strSQL = strSQL & "')"
        
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    Exit Sub
    
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
    
'    For i = 1 To cModifiedMouseUsage.Count
'        If cModifiedMouseUsage(i).bModified Then
'            '���浽ϵͳ����
'            cMouseUsage(i).bSelected = cModifiedMouseUsage(i).bSelected
'            cMouseUsage(i).lngFuncNo = cModifiedMouseUsage(i).lngFuncNo
'            cMouseUsage(i).lngMouseKey = cModifiedMouseUsage(i).lngMouseKey
'            cMouseUsage(i).lngShift = cModifiedMouseUsage(i).lngShift
'            cMouseUsage(i).strProgramName = cModifiedMouseUsage(i).strProgramName
'            cMouseUsage(i).strShowName = cModifiedMouseUsage(i).strShowName
'
'
'            '���浽���ݿ�
'            '����һ����ѡ�е���굱ǰ������Ҫ����
'            If cModifiedMouseUsage(i).bSelected Then
'                strSQL = "UPDATE ��갴ť���� SET �Ƿ�ѡ��= false WHERE �Ƿ�ѡ�� = true AND ����λ=" & _
'                         cModifiedMouseUsage(i).lngMouseKey & " AND SHIFT��λ = " & cModifiedMouseUsage(i).lngShift
'                cnAccess.Execute strSQL, , adCmdText
'            End If
'
'            '���ڻ���ע��״̬�ı���Ҫ��������
'            If cMouseUsage(i).lngFuncNo <= lngDrawLabelFuncNo Then
'                If cMouseUsage(i).lngFuncNo = lngDrawLabelFuncNo Then
'                    iDrawLabelIndex = i
'                    strSQL = "UPDATE ��갴ť���� SET ����λ =" & cModifiedMouseUsage(i).lngMouseKey & _
'                     ",SHIFT��λ=" & cModifiedMouseUsage(i).lngShift & _
'                     " WHERE �������=" & cModifiedMouseUsage(i).lngFuncNo
'                cnAccess.Execute strSQL, , adCmdText
'                End If
'            Else
'                strSQL = "UPDATE ��갴ť���� SET ����λ =" & cModifiedMouseUsage(i).lngMouseKey & _
'                     ",SHIFT��λ=" & cModifiedMouseUsage(i).lngShift & _
'                     ",�Ƿ�ѡ��=" & IIf(cModifiedMouseUsage(i).bSelected, True, False) & _
'                     " WHERE �������=" & cModifiedMouseUsage(i).lngFuncNo
'                cnAccess.Execute strSQL, , adCmdText
'            End If
'        End If
'    Next
'    '�ж��Ƿ񻭱�ע��״̬��ѡ��ɵ�ǰ���״̬�ˣ�����ǣ��򱣴湦�����ΪlngDrawLabelFuncNo��
'    '��Ϊ��ǰ���
'    If iDrawLabelIndex <> -1 Then
'        For i = 1 To cMouseUsage.Count
'            If cMouseUsage(i).lngFuncNo < lngDrawLabelFuncNo Then
'
'                cMouseUsage(i).lngMouseKey = cMouseUsage(iDrawLabelIndex).lngMouseKey
'                cMouseUsage(i).lngShift = cMouseUsage(iDrawLabelIndex).lngShift
'
'                If cMouseUsage(i).lngFuncNo = lngDrawLabelCurrent Then
'                    cMouseUsage(i).bSelected = cMouseUsage(iDrawLabelIndex).bSelected
'                Else
'                    cMouseUsage(i).bSelected = False
'                End If
'                strSQL = "UPDATE ��갴ť���� SET ����λ =" & cMouseUsage(i).lngMouseKey & _
'                     ",SHIFT��λ=" & cMouseUsage(i).lngShift & _
'                     ",�Ƿ�ѡ��=" & IIf(cMouseUsage(i).bSelected, True, False) & _
'                     " WHERE �������=" & cMouseUsage(i).lngFuncNo
'                cnAccess.Execute strSQL, , adCmdText
'            End If
'        Next
'    End If
End Sub

Private Sub subSetchkShiftState(ByVal iShift As Integer)
'------------------------------------------------
'���ܣ����������ֵ�����ý��������shift��״̬����ʾ
'������iShift--����Shift��״̬
'���أ���
'�ϼ���������̣�frmSysConfig.lstMouseKey_Click��frmSysConfig.subFillMouseUsage
'�¼���������̣���
'���õ��ⲿ������
'�����ˣ��ƽ�
'------------------------------------------------
    Dim i As Integer
    For i = 1 To 3
        Me.chkShiftState(i) = 0
    Next
    i = iShift
    'shift ���÷���shift,ctrl,alt �ֱ���1��2��4��ʾ
    If i - 4 >= 0 Then
        Me.chkShiftState(3) = 1
        i = i - 4
    End If
    If i - 2 >= 0 Then
        Me.chkShiftState(2) = 1
        i = i - 2
    End If
    If i = 1 Then
        Me.chkShiftState(1) = 1
    End If
End Sub

Private Sub subKeepMouseUsage(bMouseKey As Boolean, Optional iMouseFuncNo As Integer = 0, Optional bSelected As Boolean = False)
'------------------------------------------------
'���ܣ����汻�޸ģ�����û�б�Ӧ�õ�����÷��޸ġ�
'������bMouseKey--True����������Ҽ����޸ģ�False�������Shift�����޸ģ�
'      iMouseFuncNo--��ǰ��ʹ�õ�����÷���ţ��ڽ����ϱ�ʶΪ�򹴣�bSelected--��ǰ����Ƿ�ѡ��
'���أ���
'�ϼ���������̣�frmSysConfig.lstMouseKey_Click��frmSysConfig.subMoveLeftRight
'               frmSysConfig.chkShiftState_LostFocus
'�¼���������̣���
'���õ��ⲿ������cModifiedMouseUsage
'�����ˣ��ƽ�
'------------------------------------------------
    Dim i As Integer, j As Integer
    Dim iCount As Integer
    Dim lngCurrentShift As Long
    lngCurrentShift = Me.chkShiftState(1) + Me.chkShiftState(2) * 2 + Me.chkShiftState(3) * 4
    '����Ǳ��浱ǰʹ�õ��������ã��򱣴���Ͼ��˳�����
    If Not iMouseFuncNo = 0 Then
        For i = 1 To cModifiedMouseUsage.Count
            If (cModifiedMouseUsage(i).lngMouseKey = ilstActive) And _
               (cModifiedMouseUsage(i).lngShift = lngCurrentShift) Then
                cModifiedMouseUsage(i).bSelected = False
            End If
        Next
        cModifiedMouseUsage(iMouseFuncNo).bModified = True
        cModifiedMouseUsage(iMouseFuncNo).bSelected = bSelected
        Exit Sub
    End If
    
    If bMouseKey Then   ''����������Ҽ����޸ļ�¼
        '��������listbox
        For i = 1 To Me.lstMouseKey(1).ListCount
            iCount = Me.lstMouseKey(1).ItemData(i - 1)
            If cMouseUsage(iCount).lngMouseKey <> 1 Then
                cModifiedMouseUsage(iCount).bModified = True
                cModifiedMouseUsage(iCount).lngMouseKey = 1
            End If
        Next
        '�����ұ�listbox
        For i = 1 To Me.lstMouseKey(2).ListCount
            iCount = Me.lstMouseKey(2).ItemData(i - 1)
            If cMouseUsage(iCount).lngMouseKey <> 2 Then
                cModifiedMouseUsage(iCount).bModified = True
                cModifiedMouseUsage(iCount).lngMouseKey = 2
            End If
        Next
    Else    '�������shift��
        i = Me.chkShiftState(1).Tag
        cModifiedMouseUsage(i).bModified = True
        cModifiedMouseUsage(i).lngShift = lngCurrentShift
    End If
End Sub

Private Sub lstMouseKey_Click(Index As Integer)
    'ˢ��shift״̬��checkbox
    Dim iNo As Integer
    Dim i As Integer
    If (Me.lstMouseKey(Index).ListIndex = -1) Then
        Exit Sub
    End If
    
    If (Me.lstMouseKey(Index).ItemData(Me.lstMouseKey(Index).ListIndex) = 0) Then
        Exit Sub
    End If
    
    iNo = Me.lstMouseKey(Index).ItemData(Me.lstMouseKey(Index).ListIndex)
    
    '�����һ��ѡ��ʱ������ǰ��ѡ��������
    If Not ilstActive = Index Then
        '������һ��listbox��ѡ����ȥ��
        If Index = 1 Then
            Me.lstMouseKey(2).ListIndex = -1
        Else
            Me.lstMouseKey(1).ListIndex = -1
        End If
    End If
    
    '����shift״̬�ؼ�����ʾ����
    Me.chkShiftState(1).Tag = iNo
    If cModifiedMouseUsage(iNo).bModified Then
        subSetchkShiftState cModifiedMouseUsage(iNo).lngShift
    Else
        subSetchkShiftState cMouseUsage(iNo).lngShift
    End If
    
    '���ڽ�ѡ��ѡ�е��������Ҫ���䴦��ɵ�ǰ���״̬
    If Me.lstMouseKey(Index).Selected(Me.lstMouseKey(Index).ListIndex) <> cModifiedMouseUsage(iNo).bSelected Then
        Dim iTmpNo As Integer
        '��ԭ��shift״̬��ͬ����굱ǰ��ɾ��
        If Me.lstMouseKey(Index).Selected(Me.lstMouseKey(Index).ListIndex) = True Then
            For i = 0 To Me.lstMouseKey(Index).ListCount - 1
                iTmpNo = Me.lstMouseKey(Index).ItemData(i)
                If (Me.lstMouseKey(Index).Selected(i) = True) And _
                   (cModifiedMouseUsage(iTmpNo).lngShift = cModifiedMouseUsage(iNo).lngShift) And _
                   (iTmpNo <> iNo) Then
                    Me.lstMouseKey(Index).Selected(i) = False
                End If
            Next
        End If
        subKeepMouseUsage True, iNo, Me.lstMouseKey(Index).Selected(Me.lstMouseKey(Index).ListIndex)
        bMouseUsageModified = True
    End If

    ilstActive = Index
End Sub

Private Sub subMoveLeftRight(ilst1 As Integer, ilst2 As Integer)
'------------------------------------------------
'���ܣ���������ilst1���Ƶ�ilst1ָ���listbox�������ƶ������¼������
'������ilst1--Դlistbox�ı��  ��ilst1--Ŀ��listbox�ı��
'���أ���
'�ϼ���������̣�frmSysConfig.cmdLeftRight_Click
'�¼���������̣�frmSysConfig.subKeepMouseUsage
'���õ��ⲿ��������
'�����ˣ��ƽ�
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    i = 0
    If Me.lstMouseKey(ilst1).ListIndex = -1 Then Exit Sub
    
    '�жϵ�ǰ��ѡ�е����Ƿ��Ѿ�������Ϊ��굱ǰ��
    If Me.lstMouseKey(ilst1).Selected(Me.lstMouseKey(ilst1).ListIndex) = True Then
        MsgBox "���ƶ��Ĳ������趨Ϊ���ĵ�ǰ�������޷��ƶ���", vbInformation, gstrSysName
        Exit Sub
    Else
        Me.lstMouseKey(ilst2).AddItem Me.lstMouseKey(ilst1).list(Me.lstMouseKey(ilst1).ListIndex)
        Me.lstMouseKey(ilst2).ItemData(Me.lstMouseKey(ilst2).NewIndex) = Me.lstMouseKey(ilst1).ItemData(Me.lstMouseKey(ilst1).ListIndex)
        '��������ilst1ָ���listbox��ɾ��
        Me.lstMouseKey(ilst1).RemoveItem Me.lstMouseKey(ilst1).ListIndex
        '���޸�������浽�޸ļ�¼��,����������Ҽ����޸Ŀ��Բ���¼���޸ļ�¼��
        subKeepMouseUsage True
    End If
End Sub

Private Sub cmdLeftRight_Click(Index As Integer)
    If Index = 1 Then
        subMoveLeftRight 1, 2
    Else
        subMoveLeftRight 2, 1
    End If
    bMouseUsageModified = True
End Sub

Private Sub lstMouseStep_Scroll(Index As Integer)
    blnInterfaceParaModified = True
End Sub

Private Sub chkShiftState_LostFocus(Index As Integer)
    bMouseKeyShiftClick = False
    subKeepMouseUsage False
    bMouseUsageModified = True
End Sub

Private Sub chkShiftState_Click(Index As Integer)
    If ilstActive = 0 Then Exit Sub
    If (bMouseKeyShiftClick = True) And Me.lstMouseKey(ilstActive).ListIndex <> -1 Then
        If (cModifiedMouseUsage(Me.lstMouseKey(ilstActive).ItemData(Me.lstMouseKey(ilstActive).ListIndex)).bSelected = True) Then
            bMouseKeyShiftClick = False
            '��鵱ǰlist���Ƿ�������λ����ǰ����ͬ�������ͬ���򵯳���ʾ
            Dim i As Integer
            Dim lngCurrentShift As Long
            lngCurrentShift = Me.chkShiftState(1) + Me.chkShiftState(2) * 2 + Me.chkShiftState(3) * 4
            For i = 1 To cModifiedMouseUsage.Count
                If (cModifiedMouseUsage(i).lngMouseKey = ilstActive) And _
                   (cModifiedMouseUsage(i).bSelected = True) And _
                   (cModifiedMouseUsage(i).lngShift = lngCurrentShift) And _
                   (cModifiedMouseUsage(i).lngFuncNo >= lngDrawLabelFuncNo) Then
                        MsgBox "��깦�ܣ�" & cModifiedMouseUsage(i).strShowName & " ������λ��Shift��Ctrl��Alt�����뵱ǰ���������ͬ�����������á�", vbInformation, gstrSysName
                        Me.chkShiftState(Index) = IIf(Me.chkShiftState(Index) = 1, 0, 1)
                        bMouseKeyShiftClick = True
                        Exit Sub
                End If
            Next
        End If
    End If
End Sub

Private Sub chkShiftState_GotFocus(Index As Integer)
    bMouseKeyShiftClick = True
End Sub

Private Sub subFillInfoLabe()
'------------------------------------------------
'���ܣ�������ؼ��������Ľ���Ϣ��עλ�ú���ʾ����
'��������
'���أ���
'�ϼ���������̣�frmSysConfig.Form_Resize
'�¼���������̣���
'���õ��ⲿ������aInfoLabelLocate
'�����ˣ��ƽ�
'------------------------------------------------
    Dim i As Integer, j As Integer
    Dim iLoc As Integer         ''��ʱ���ʹ�����Ǹ���
    Dim iCount(4) As Integer
    Dim iTemp() As Integer
    Dim iMax As Integer
    Dim s As String
    
    On Error GoTo errHandle
    '��ʼ����Ļ�ؼ�
    Me.lstInfoLabelAll.Clear
    For i = 1 To 4
        Me.lstInfoLabelSel(i).Clear
        iCount(i) = 0
    Next
    '�����Ļ�ؼ�
    For i = 1 To lngInfoLabelCount
        If aInfoLabelLocate(i).bUsed Then   ''�ŵ��ĸ�����
            iLoc = aInfoLabelLocate(i).lngLocation
            iCount(iLoc) = iCount(iLoc) + 1
        Else                                ''�ŵ���ѡlst������
            Me.lstInfoLabelAll.AddItem aInfoLabelLocate(i).strCName
            Me.lstInfoLabelAll.ItemData(Me.lstInfoLabelAll.NewIndex) = i            ' aInfoLabelLocate(i).lngID
        End If
    Next
    
    iMax = iCount(1)
    If iMax < iCount(2) Then iMax = iCount(2)
    If iMax < iCount(3) Then iMax = iCount(3)
    If iMax < iCount(4) Then iMax = iCount(4)
    
    ReDim iTemp(4, iMax) As Integer
    For j = 1 To 4
        For i = 1 To lngInfoLabelCount
            If aInfoLabelLocate(i).bUsed And j = aInfoLabelLocate(i).lngLocation Then    ''�ŵ��ĸ�����
                iTemp(j, aInfoLabelLocate(i).lngOrder) = i
            End If
        Next
    Next
    
    For j = 1 To 4
        For i = 0 To iCount(j) - 1
            s = aInfoLabelLocate(iTemp(j, i)).strCName
            If aInfoLabelLocate(iTemp(j, i)).blnIsExport Then
                s = s & "-�ɵ���"
            End If
            
            Me.lstInfoLabelSel(j).AddItem s
            Me.lstInfoLabelSel(j).ItemData(Me.lstInfoLabelSel(j).NewIndex) = iTemp(j, i)   ' aInfoLabelLocate(i).lngID
        Next
    Next
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdSelInfoLabel_Click(Index As Integer)
    Dim iSel As Integer
    iSel = Me.lstInfoLabelAll.ListIndex
    If iSel = -1 Then
        Exit Sub
    End If
    Me.lstInfoLabelSel(Index).AddItem Me.lstInfoLabelAll.list(iSel)
    Me.lstInfoLabelSel(Index).ItemData(Me.lstInfoLabelSel(Index).NewIndex) = Me.lstInfoLabelAll.ItemData(iSel)
    Me.lstInfoLabelAll.RemoveItem iSel
    Me.lstInfoLabelAll.ListIndex = iSel - 1
    bInfoLabelModified = True
End Sub

Private Sub cmdInfoLabelUpDown_Click(Index As Integer)
    '���б������½������ֵ���ʾλ��
    Dim iSel As Integer
    Dim iOSel As Integer
    Dim strTempName As String
    Dim iTempID As Integer
    If ilstInfoLabelActvate = 0 Then
        Exit Sub
    End If
    iSel = Me.lstInfoLabelSel(ilstInfoLabelActvate).ListIndex
    iOSel = iSel + Index - 1
    If iOSel < Me.lstInfoLabelSel(ilstInfoLabelActvate).ListCount And iOSel > -1 Then
        strTempName = Me.lstInfoLabelSel(ilstInfoLabelActvate).list(iOSel)
        iTempID = Me.lstInfoLabelSel(ilstInfoLabelActvate).ItemData(iOSel)
        Me.lstInfoLabelSel(ilstInfoLabelActvate).list(iOSel) = Me.lstInfoLabelSel(ilstInfoLabelActvate).list(iSel)
        Me.lstInfoLabelSel(ilstInfoLabelActvate).ItemData(iOSel) = Me.lstInfoLabelSel(ilstInfoLabelActvate).ItemData(iSel)
        Me.lstInfoLabelSel(ilstInfoLabelActvate).list(iSel) = strTempName
        Me.lstInfoLabelSel(ilstInfoLabelActvate).ItemData(iSel) = iTempID
        Me.lstInfoLabelSel(ilstInfoLabelActvate).ListIndex = iOSel
        bInfoLabelModified = True
    End If
End Sub

Private Sub cmdDeSelInfoLabel_Click()
    Dim s As String
    
    On Error GoTo errHandle
    '����Ϣ��ע����Ϊδ��ѡ���״̬
    If ilstInfoLabelActvate = 0 Then
        Exit Sub
    End If
    Dim iSel As Integer
    iSel = Me.lstInfoLabelSel(ilstInfoLabelActvate).ListIndex
    If iSel <> -1 Then
        s = Me.lstInfoLabelSel(ilstInfoLabelActvate).list(iSel)
        s = Replace(s, "-�ɵ���", "")
                
        Me.lstInfoLabelAll.AddItem s
        Me.lstInfoLabelAll.ItemData(Me.lstInfoLabelAll.NewIndex) = Me.lstInfoLabelSel(ilstInfoLabelActvate).ItemData(iSel)
        Me.lstInfoLabelSel(ilstInfoLabelActvate).RemoveItem iSel
        Me.lstInfoLabelSel(ilstInfoLabelActvate).ListIndex = iSel - 1
        bInfoLabelModified = True
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub subSaveInfoLabelLocate()
'------------------------------------------------
'���ܣ��������Ľ���Ϣ��ʾ���õĽ�����浽��ͼ����Ϣ����ϵͳ�����С�
'��������
'���أ���
'�ϼ���������̣�frmSysConfig.cmdApply_Click
'�¼���������̣���
'���õ��ⲿ������frmSysConfig.aInfoLabelLocate
'�����ˣ��ƽ�
'------------------------------------------------
    If bInfoLabelModified = False Then
        Exit Sub
    End If
    Dim i As Integer
    Dim ilst As Integer
    Dim strSQL As String
    Dim iArray As Integer
    Dim s As String
    
    On Error GoTo errh
    
    '����δ��ѡ�����Ϣ��ע
    For i = 0 To Me.lstInfoLabelAll.ListCount - 1
        '����Ϣ��ע�ı�ʹ��״̬�����ı䣬����м�¼
        If aInfoLabelLocate(Me.lstInfoLabelAll.ItemData(i)).bUsed Then
            'ˢ��ϵͳ����
            aInfoLabelLocate(Me.lstInfoLabelAll.ItemData(i)).bUsed = False
            'д�����ݿ�
            If blLocalRun = True Then
                strSQL = "UPDATE Ӱ��ͼ����Ϣ�� SET ��ѡ��=FALSE WHERE ID=" & _
                         aInfoLabelLocate(Me.lstInfoLabelAll.ItemData(i)).lngID
                cnAccess.Execute strSQL, , adCmdText
            Else
                strSQL = "ZL_Ӱ��ͼ����Ϣ��_UPDATE(" & aInfoLabelLocate(Me.lstInfoLabelAll.ItemData(i)).lngID & ",'',0,Null,Null,Null)"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        End If
    Next
    
    '�������Ͻǵ���Ϣ��ע
    For ilst = 1 To 4
        For i = 0 To Me.lstInfoLabelSel(ilst).ListCount - 1
            iArray = Me.lstInfoLabelSel(ilst).ItemData(i)
            s = Me.lstInfoLabelSel(ilst).list(i)
            
            If (Not aInfoLabelLocate(iArray).bUsed) Or (aInfoLabelLocate(iArray).lngLocation <> ilst) Or _
                (aInfoLabelLocate(iArray).lngOrder <> i) Or (aInfoLabelLocate(iArray).strCName & IIf(aInfoLabelLocate(iArray).blnIsExport, "-�ɵ���", "") <> s) Then
                'ˢ��ϵͳ����
                aInfoLabelLocate(iArray).bUsed = True
                aInfoLabelLocate(iArray).lngLocation = ilst
                aInfoLabelLocate(iArray).lngOrder = i
                aInfoLabelLocate(iArray).blnIsExport = IIf(InStr(1, s, "-�ɵ���") > 0, -1, 0)
                'д�����ݿ�
                If blLocalRun = True Then
                    strSQL = "UPDATE Ӱ��ͼ����Ϣ�� SET ��ѡ��=TRUE, λ��=" & CStr(ilst) & _
                        ",������� = " & CStr(i) & " WHERE ID=" & aInfoLabelLocate(iArray).lngID
                    cnAccess.Execute strSQL, , adCmdText
                Else
                    
                    strSQL = "ZL_Ӱ��ͼ����Ϣ��_UPDATE(" & aInfoLabelLocate(iArray).lngID & ",'',-1," & CStr(ilst) & "," & CStr(i) & "," & IIf(InStr(1, s, "-�ɵ���") > 0, -1, 0) & ")"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                End If
            End If
        Next
    Next
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Private Sub lstInfoLabel_Scroll(Index As Integer)
    blnInterfaceParaModified = True
End Sub

Private Sub cmdInfoLabelColor_Click()
    Me.dlgColor.Color = Me.shpInfoLabel.FillColor
    Me.dlgColor.ShowColor
    Me.shpInfoLabel.FillColor = Me.dlgColor.Color
    blnInterfaceParaModified = True
End Sub

Private Sub optPatientInfoTitle_LostFocus(Index As Integer)
    blnInterfaceParaModified = True
End Sub

Private Sub optShutter_Click(Index As Integer)
    subEnableShutterControl IIf(Index = 1, True, False)
End Sub

Private Sub optShutter_LostFocus(Index As Integer)
    subKeepImageShutter
End Sub



Private Sub txtAETitle_GotFocus()
    txtAETitle.SelStart = 0
    txtAETitle.SelLength = Len(txtAETitle.Text)
End Sub

Private Sub txtAETitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtCircle_GotFocus(Index As Integer)
    txtCircle(Index).SelStart = 0
    txtCircle(Index).SelLength = Len(txtCircle(Index).Text)
End Sub

Private Sub txtCircle_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    'ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtCircle_LostFocus(Index As Integer)
    subKeepImageShutter
End Sub






Private Sub txtFilmFontSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    'ֻ����������
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub


Private Sub txtFilterModality_GotFocus()
    txtFilterModality.SelStart = 0
    txtFilterModality.SelLength = Len(txtFilterModality.Text)
End Sub

Private Sub txtFilterModality_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFilterName_GotFocus()
    txtFilterName.SelStart = 0
    txtFilterName.SelLength = Len(txtFilterName.Text)
End Sub

Private Sub txtFilterName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFilterPara_GotFocus(Index As Integer)
    txtFilterPara(Index).SelStart = 0
    txtFilterPara(Index).SelLength = Len(txtFilterPara(Index).Text)
End Sub

Private Sub txtFilterPara_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    'ֻ����������
    If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtImageBorderWidth_GotFocus()
    txtImageBorderWidth.SelStart = 0
    txtImageBorderWidth.SelLength = Len(txtImageBorderWidth.Text)
End Sub

Private Sub txtImageBorderWidth_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    'ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtImageResolution_GotFocus()
    txtImageResolution.SelStart = 0
    txtImageResolution.SelLength = Len(txtImageResolution.Text)
End Sub

Private Sub txtImageResolution_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    'ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtIPAddress_GotFocus()
    txtIPAddress.SelStart = 0
    txtIPAddress.SelLength = Len(txtIPAddress.Text)
End Sub

Private Sub txtIPAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    'ֻ����������
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtLabelFontSize_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub txtLabelLineWidth_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub txtLocalAE_GotFocus()
    txtLocalAE.SelStart = 0
    txtLocalAE.SelLength = Len(txtLocalAE.Text)
End Sub

Private Sub txtLocalAE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtLocalAE_LostFocus()
    cstrPrintAE = Me.txtLocalAE
End Sub

Private Sub txtPatientInfoFontName_GotFocus()
    txtPatientInfoFontName.SelStart = 0
    txtPatientInfoFontName.SelLength = Len(txtPatientInfoFontName.Text)
End Sub

Private Sub txtPatientInfoFontName_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub txtPatientInfoInVisibleSize_Change()
    blnInterfaceParaModified = True
End Sub

Private Sub chkInfoLabelScale_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub lstPatientInfoFontSize_Scroll()
    blnInterfaceParaModified = True
End Sub
 
Private Sub subFillMSFPrinter()
'------------------------------------------------
'���ܣ���ϵͳ��Ϣ��д��MSF�����
'��������
'���أ���
'�ϼ���������̣�frmSysConfig.cmdDICOMPrintAdd_Click��frmSysConfig.cmdDICOMPrintDelete_Click
'               frmSysConfig.cmdDICOMPrintUpdate_Click��frmSysConfig.Form_Resize
'�¼���������̣���
'���õ��ⲿ������cDICOMPrinter
'�����ˣ��ƽ�
'------------------------------------------------
    Dim i As Integer
    '��ʼ���ؼ�����ʾ
    Me.txtPrinterName = ""
    Me.txtAETitle = ""
    Me.txtIPAddress = ""
    Me.txtPort = ""
    Me.cboFormat.ListIndex = -1
    Me.cboPriority.ListIndex = -1
    Me.cboMedium.ListIndex = -1
    Me.lstCopies.ListIndex = -1
    Me.cboOrientation.ListIndex = -1
    Me.cboFilmSize.ListIndex = -1
    Me.cboFilmBox.ListIndex = -1
    Me.cboResolution.ListIndex = -1
    Me.cboMagnification.ListIndex = -1
    Me.cboSmooth.ListIndex = -1
    Me.cboTrim.ListIndex = -1
    Me.cboBorderDensity.ListIndex = -1
    Me.cboEmptyDensity.ListIndex = -1
    Me.cboPolarity.ListIndex = -1
    Me.lstDensity(1).ListIndex = -1
    Me.lstDensity(2).ListIndex = -1
    Me.cboBitDepth.ListIndex = 0
    Me.txtImageBorderWidth = 1
    Me.txtImageResolution = "300"
    
    Me.MSFPrinter.Clear
    Me.MSFPrinter.Rows = 1
    Me.MSFPrinter.Cols = 24
    Me.MSFPrinter.TextMatrix(0, 0) = "��ӡ������"
    Me.MSFPrinter.TextMatrix(0, 1) = "AE����"
    Me.MSFPrinter.TextMatrix(0, 2) = "IP��ַ"
    Me.MSFPrinter.TextMatrix(0, 3) = "�˿ں�"
    Me.MSFPrinter.TextMatrix(0, 4) = "��ӡ��ʽ"
    Me.MSFPrinter.TextMatrix(0, 5) = "���ȼ�"
    Me.MSFPrinter.TextMatrix(0, 6) = "����"
    Me.MSFPrinter.TextMatrix(0, 7) = "��ӡ����"
    Me.MSFPrinter.TextMatrix(0, 8) = "����"
    Me.MSFPrinter.TextMatrix(0, 9) = "��Ƭ���"
    Me.MSFPrinter.TextMatrix(0, 10) = "ѡ��Ƭ��"
    Me.MSFPrinter.TextMatrix(0, 11) = "�Ŵ�ģʽ"
    Me.MSFPrinter.TextMatrix(0, 12) = "��Ƭ�ֱ���"
    Me.MSFPrinter.TextMatrix(0, 13) = "ƽ��ģʽ"
    Me.MSFPrinter.TextMatrix(0, 14) = "����"
    Me.MSFPrinter.TextMatrix(0, 15) = "����ܶ�"
    Me.MSFPrinter.TextMatrix(0, 16) = "��С�ܶ�"
    Me.MSFPrinter.TextMatrix(0, 17) = "�߿��ܶ�"
    Me.MSFPrinter.TextMatrix(0, 18) = "�հ��ܶ�"
    Me.MSFPrinter.TextMatrix(0, 19) = "����"
    Me.MSFPrinter.TextMatrix(0, 20) = "ͼ��λ��"
    Me.MSFPrinter.TextMatrix(0, 21) = "�û�AE����"
    Me.MSFPrinter.TextMatrix(0, 22) = "ͼ��߿���"
    Me.MSFPrinter.TextMatrix(0, 23) = "ͼƬ�ֱ���"
    Dim lngRowPos As Long               '��¼��ǰ������
    lngRowPos = 1
    For i = 1 To cDICOMPrinter.Count            '��д���ݱ��
        Dim a As clsDicomPrint
        Me.MSFPrinter.Rows = Me.MSFPrinter.Rows + 1
        Me.MSFPrinter.TextMatrix(lngRowPos, 0) = cDICOMPrinter(i).strname
        Me.MSFPrinter.TextMatrix(lngRowPos, 1) = cDICOMPrinter(i).strAETitle
        Me.MSFPrinter.TextMatrix(lngRowPos, 2) = cDICOMPrinter(i).strIPAddress
        Me.MSFPrinter.TextMatrix(lngRowPos, 3) = cDICOMPrinter(i).lngPort
        Me.MSFPrinter.TextMatrix(lngRowPos, 4) = cDICOMPrinter(i).strFormat
        Me.MSFPrinter.TextMatrix(lngRowPos, 5) = cDICOMPrinter(i).strPriority
        Me.MSFPrinter.TextMatrix(lngRowPos, 6) = cDICOMPrinter(i).strMedium
        Me.MSFPrinter.TextMatrix(lngRowPos, 7) = cDICOMPrinter(i).lngCopies
        Me.MSFPrinter.TextMatrix(lngRowPos, 8) = cDICOMPrinter(i).strOrientation
        Me.MSFPrinter.TextMatrix(lngRowPos, 9) = cDICOMPrinter(i).strFilmSize
        Me.MSFPrinter.TextMatrix(lngRowPos, 10) = cDICOMPrinter(i).strFilmBox
        Me.MSFPrinter.TextMatrix(lngRowPos, 11) = cDICOMPrinter(i).strMagnification
        Me.MSFPrinter.TextMatrix(lngRowPos, 12) = cDICOMPrinter(i).strResolution
        Me.MSFPrinter.TextMatrix(lngRowPos, 13) = cDICOMPrinter(i).strSmooth
        Me.MSFPrinter.TextMatrix(lngRowPos, 14) = cDICOMPrinter(i).strTrim
        Me.MSFPrinter.TextMatrix(lngRowPos, 15) = cDICOMPrinter(i).strMaxDensity
        Me.MSFPrinter.TextMatrix(lngRowPos, 16) = cDICOMPrinter(i).strMinDensity
        Me.MSFPrinter.TextMatrix(lngRowPos, 17) = cDICOMPrinter(i).strBorderDensity
        Me.MSFPrinter.TextMatrix(lngRowPos, 18) = cDICOMPrinter(i).strEmptyDensity
        Me.MSFPrinter.TextMatrix(lngRowPos, 19) = cDICOMPrinter(i).strPolarity
        Me.MSFPrinter.TextMatrix(lngRowPos, 20) = cDICOMPrinter(i).lngBitDepth
        Me.MSFPrinter.TextMatrix(lngRowPos, 21) = cDICOMPrinter(i).strSCUAETitle
        Me.MSFPrinter.TextMatrix(lngRowPos, 22) = cDICOMPrinter(i).lngImageBorderWidth
        Me.MSFPrinter.TextMatrix(lngRowPos, 23) = cDICOMPrinter(i).intImageResolution
        lngRowPos = lngRowPos + 1
    Next
End Sub
 
Private Function funSavePrinterToPara(clsOnePrinter As clsDicomPrint) As Boolean
'------------------------------------------------
'���ܣ�������ؼ�������ֵ���浽ָ����clsOnePrinterϵͳ������
'������clsOnePrinter--����ֵ�����浽��������С�
'���أ���
'�ϼ���������̣�frmSysConfig.cmdDICOMPrintAdd_Click��frmSysConfig.cmdDICOMPrintUpdate_Click
'�¼���������̣���
'���õ��ⲿ������
'�����ˣ��ƽ�
'------------------------------------------------
    '���ӱ���
    funSavePrinterToPara = False
    If Me.txtPrinterName.Text = "" Or Me.txtIPAddress.Text = "" Or Me.txtAETitle = "" _
        Or Val(Me.txtPort.Text) = 0 Or Me.txtLocalAE = "" Then
        GoTo err1
    End If
    clsOnePrinter.lngCopies = Me.lstCopies.list(Me.lstCopies.TopIndex)
    clsOnePrinter.lngPort = Val(Me.txtPort.Text)
    clsOnePrinter.strAETitle = Me.txtAETitle.Text
    clsOnePrinter.strBorderDensity = Me.cboBorderDensity.Text
    clsOnePrinter.strEmptyDensity = Me.cboEmptyDensity.Text
    clsOnePrinter.strFilmBox = Me.cboFilmBox.Text
    clsOnePrinter.strFilmSize = Me.cboFilmSize.Text
    clsOnePrinter.strFormat = Me.cboFormat.Text
    clsOnePrinter.strIPAddress = Me.txtIPAddress.Text
    clsOnePrinter.strMagnification = Me.cboMagnification.Text
    clsOnePrinter.strMaxDensity = Me.lstDensity(1).TopIndex
    clsOnePrinter.strMedium = Me.cboMedium.Text
    clsOnePrinter.strMinDensity = Me.lstDensity(2).TopIndex
    clsOnePrinter.strname = Me.txtPrinterName.Text
    clsOnePrinter.strOrientation = Me.cboOrientation.Text
    clsOnePrinter.strPolarity = Me.cboPolarity.Text
    clsOnePrinter.strPriority = Me.cboPriority.Text
    clsOnePrinter.strResolution = Me.cboResolution.Text
    clsOnePrinter.strSmooth = Me.cboSmooth.Text
    clsOnePrinter.strTrim = Me.cboTrim.Text
    clsOnePrinter.lngBitDepth = Me.cboBitDepth.Text
    clsOnePrinter.strSCUAETitle = Me.txtLocalAE.Text
    clsOnePrinter.lngImageBorderWidth = Val(Me.txtImageBorderWidth)
    clsOnePrinter.intImageResolution = Val(Me.txtImageResolution)
    funSavePrinterToPara = True
    Exit Function
err1:
    MsgBox "������Ŀ����ȷ�����顣", vbExclamation, gstrSysName
End Function

Private Sub subFillCboPrintFormat()
'����ӡ��ʽ�ؼ�
    Me.cboFormat.Clear
    Dim strSQL As String
    If blLocalRun = True Then
        strSQL = "SELECT ��ʽ��ʶ as ���� FROM Ӱ���ӡ��ʽ"
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT ���� FROM Ӱ���ӡ��ʽ"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If
    With rsTemp
        While Not .EOF
            Me.cboFormat.AddItem !����
            .MoveNext
        Wend
    End With
End Sub

Private Sub subFillCboFilmSize()
'��佺Ƭ���ؼ�
    Me.cboFilmSize.Clear
    Dim strSQL As String
    If blLocalRun = True Then
        strSQL = "SELECT ����ʶ as ���� FROM Ӱ��Ƭ���"
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT ���� FROM Ӱ��Ƭ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If
    With rsTemp
        While Not .EOF
            Me.cboFilmSize.AddItem !����
            .MoveNext
        Wend
    End With
End Sub

Private Sub MSFPrinter_Click()
    '����ǰѡ�еĴ�ӡ����Ϣ��䵽����Ŀؼ��н�����ʾ��
    Dim iRow As Integer
    Dim PrinterName As String
    Dim clsOnePrinter As clsDicomPrint
    With MSFPrinter
        If .Rows <= 1 Then Exit Sub
        iRow = .RowSel
        PrinterName = .TextMatrix(iRow, 0)
    End With
    Set clsOnePrinter = cDICOMPrinter(PrinterName)
    
    Me.lstCopies = clsOnePrinter.lngCopies
    Me.txtPort.Text = clsOnePrinter.lngPort
    Me.txtAETitle.Text = clsOnePrinter.strAETitle
    Me.cboBorderDensity.Text = clsOnePrinter.strBorderDensity
    Me.cboEmptyDensity.Text = clsOnePrinter.strEmptyDensity
    Me.cboFilmBox.Text = clsOnePrinter.strFilmBox
    Me.cboFilmSize.Text = clsOnePrinter.strFilmSize
    Me.cboFormat.Text = clsOnePrinter.strFormat
    Me.txtIPAddress.Text = clsOnePrinter.strIPAddress
    Me.cboMagnification.Text = clsOnePrinter.strMagnification
    Me.lstDensity(1).TopIndex = clsOnePrinter.strMaxDensity
    Me.cboMedium.Text = clsOnePrinter.strMedium
    Me.lstDensity(2).TopIndex = clsOnePrinter.strMinDensity
    Me.txtPrinterName.Text = clsOnePrinter.strname
    Me.cboOrientation.Text = clsOnePrinter.strOrientation
    Me.cboPolarity.Text = clsOnePrinter.strPolarity
    Me.cboPriority.Text = clsOnePrinter.strPriority
    Me.cboResolution.Text = clsOnePrinter.strResolution
    Me.cboSmooth.Text = clsOnePrinter.strSmooth
    Me.cboTrim.Text = clsOnePrinter.strTrim
    Me.cboBitDepth.Text = clsOnePrinter.lngBitDepth
    Me.txtLocalAE.Text = clsOnePrinter.strSCUAETitle
    Me.txtImageBorderWidth.Text = clsOnePrinter.lngImageBorderWidth
    Me.txtImageResolution.Text = clsOnePrinter.intImageResolution
End Sub

Private Sub cmdDICOMPrintAdd_Click()
    Dim clsOnePrinter As New clsDicomPrint
    Dim i As Integer
    
    '���жϽ�Ƭ��ӡ���������Ƿ񳬹���ɵ�����
    If cDICOMPrinter.Count >= gint��Ƭ��ӡ�� And gint��Ƭ��ӡ�� <> -1 Then
        Call MsgBox(LOGIN_TYPE_��Ƭ��ӡ�� & "�������������������" & gint��Ƭ��ӡ�� & "�����޷�����´�ӡ�������������Ӧ����ϵ��", vbOKOnly, gstrSysName)
        Exit Sub
    End If
    
    If subChkDicomPrintSetup = False Then
        Exit Sub                 '������Ƿ�������ȷ�����������ַ��Ƿ񳬳���)
    End If
    If funSavePrinterToPara(clsOnePrinter) = False Then Exit Sub   '�����������浽ϵͳ������
    
    For i = 1 To cDICOMPrinter.Count
        If cDICOMPrinter(i).strname = clsOnePrinter.strname Then
            MsgBox "����������ͬ���ƵĴ�ӡ�������޸Ĵ�ӡ�����ƺ���������", vbInformation, gstrSysName
            Exit Sub
        End If
    Next i
    cDICOMPrinter.Add clsOnePrinter, clsOnePrinter.strname
    '��������¼���ӵ����ݿ�
    
    On Error GoTo errh
    
    Dim strSQL As String
    If blLocalRun = True Then
        strSQL = "INSERT INTO Ӱ���ӡ������ ( ��ӡ����,IP��ַ,�˿ں�,AE����,��ӡ��ʽ,���ȼ�," & _
                "��ӡ����,����,����,��Ƭ���,ѡ��Ƭ��,�ֱ���,�Ŵ�ģʽ,ƽ��ģʽ,����,����ܶ�," & _
                "��С�ܶ�,�հ��ܶ�,�߿��ܶ�,����,ͼ��λ��) VALUES ('" & clsOnePrinter.strname & "','" & _
                clsOnePrinter.strIPAddress & "','" & clsOnePrinter.lngPort & "','" & _
                clsOnePrinter.strAETitle & "','" & clsOnePrinter.strFormat & "','" & _
                clsOnePrinter.strPriority & "'," & clsOnePrinter.lngCopies & ",'" & _
                clsOnePrinter.strMedium & "','" & clsOnePrinter.strOrientation & "','" & _
                clsOnePrinter.strFilmSize & "','" & clsOnePrinter.strFilmBox & "','" & _
                clsOnePrinter.strResolution & "','" & clsOnePrinter.strMagnification & "','" & _
                clsOnePrinter.strSmooth & "','" & clsOnePrinter.strTrim & "','" & _
                clsOnePrinter.strMaxDensity & "','" & clsOnePrinter.strMinDensity & "','" & _
                clsOnePrinter.strEmptyDensity & "','" & clsOnePrinter.strBorderDensity & "','" & _
                clsOnePrinter.strPolarity & "'," & clsOnePrinter.lngBitDepth & "')"
        cnAccess.Execute strSQL, , adCmdText
    Else
        strSQL = "ZL_Ӱ���ӡ������_INSERT('" & clsOnePrinter.strname & _
        "','" & clsOnePrinter.strIPAddress & "','" & clsOnePrinter.lngPort & "','" & clsOnePrinter.strAETitle & _
        "','" & clsOnePrinter.strFormat & "','" & clsOnePrinter.strPriority & "'," & clsOnePrinter.lngCopies & _
        ",'" & clsOnePrinter.strMedium & "','" & clsOnePrinter.strOrientation & "','" & clsOnePrinter.strFilmSize & _
        "','" & clsOnePrinter.strFilmBox & "','" & clsOnePrinter.strResolution & "','" & clsOnePrinter.strMagnification & _
        "','" & clsOnePrinter.strSmooth & "','" & clsOnePrinter.strTrim & "','" & clsOnePrinter.strMaxDensity & _
        "','" & clsOnePrinter.strMinDensity & "','" & clsOnePrinter.strEmptyDensity & _
        "','" & clsOnePrinter.strBorderDensity & "','" & clsOnePrinter.strPolarity & "'," & clsOnePrinter.lngBitDepth & _
        ",'" & clsOnePrinter.strSCUAETitle & "'," & clsOnePrinter.lngImageBorderWidth & "," & clsOnePrinter.intImageResolution & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    subFillMSFPrinter       'ˢ�¿ؼ���ʾ
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub

Private Sub cmdDICOMPrintDelete_Click()
    'ɾ��һ��DICOM��ӡ��
    Dim PrinterName As String
    Dim lngPrinterID As Long
    If MSFPrinter.Rows = 1 Then Exit Sub
    PrinterName = Me.MSFPrinter.TextMatrix(Me.MSFPrinter.RowSel, 0)
    lngPrinterID = cDICOMPrinter(PrinterName).lngID
    cDICOMPrinter.Remove (PrinterName)
    
    On Error GoTo errh
    
    Dim strSQL As String
    If blLocalRun = True Then
        strSQL = "DELETE FROM Ӱ���ӡ������ WHERE id=" & lngPrinterID
        cnAccess.Execute strSQL, , adCmdText
    Else
        strSQL = "ZL_Ӱ���ӡ������_DELETE(" & lngPrinterID & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    subFillMSFPrinter       'ˢ�¿ؼ���ʾ
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Private Sub cmdDICOMPrintUpdate_Click()
    '�޸�һ��DICOM��ӡ��������
    Dim clsOnePrinter As clsDicomPrint
    Dim PrinterName As String
    Dim lngCollectionIndex As Long
    
    PrinterName = Me.MSFPrinter.TextMatrix(Me.MSFPrinter.RowSel, 0)
    If cDICOMPrinter.Count = 0 Then Exit Sub
    Set clsOnePrinter = cDICOMPrinter(PrinterName)
    If subChkDicomPrintSetup = False Then
        Exit Sub                 '������Ƿ�������ȷ�����������ַ��Ƿ񳬳���)
    End If
    If funSavePrinterToPara(clsOnePrinter) = False Then Exit Sub
    
    '�����������д�����ݿ�
    Dim strSQL As String
    
    On Error GoTo errh
    
    If blLocalRun = True Then
        strSQL = "UPDATE Ӱ���ӡ������ SET ��ӡ����='" & clsOnePrinter.strname & "',IP��ַ='" & _
                clsOnePrinter.strIPAddress & "',�˿ں�=" & clsOnePrinter.lngPort & ",AE����='" & _
                clsOnePrinter.strAETitle & "',��ӡ��ʽ='" & clsOnePrinter.strFormat & "',���ȼ�='" & _
                clsOnePrinter.strPriority & "',��ӡ����= " & clsOnePrinter.lngCopies & ",����='" & _
                clsOnePrinter.strMedium & "',����='" & clsOnePrinter.strOrientation & "',��Ƭ���='" & _
                clsOnePrinter.strFilmSize & "',ѡ��Ƭ��='" & clsOnePrinter.strFilmBox & "',�ֱ���='" & _
                clsOnePrinter.strResolution & "',�Ŵ�ģʽ='" & clsOnePrinter.strMagnification & _
                "',ƽ��ģʽ='" & clsOnePrinter.strSmooth & "',����='" & clsOnePrinter.strTrim & _
                "',����ܶ�='" & clsOnePrinter.strMaxDensity & "',��С�ܶ�='" & clsOnePrinter.strMinDensity & _
                "',�հ��ܶ�='" & clsOnePrinter.strEmptyDensity & "',�߿��ܶ� = '" & clsOnePrinter.strBorderDensity & _
                "',����='" & clsOnePrinter.strPolarity & "',ͼ��λ��=" & clsOnePrinter.lngBitDepth & _
                ",�û�AE���� ='" & clsOnePrinter.strSCUAETitle & "' " & _
                " WHERE ID=" & clsOnePrinter.lngID
        cnAccess.Execute strSQL, , adCmdText
    Else
        strSQL = "ZL_Ӱ���ӡ������_UPDATE(" & clsOnePrinter.lngID & ",'" & clsOnePrinter.strname & _
        "','" & clsOnePrinter.strIPAddress & "','" & clsOnePrinter.lngPort & "','" & clsOnePrinter.strAETitle & _
        "','" & clsOnePrinter.strFormat & "','" & clsOnePrinter.strPriority & "'," & clsOnePrinter.lngCopies & _
        ",'" & clsOnePrinter.strMedium & "','" & clsOnePrinter.strOrientation & "','" & clsOnePrinter.strFilmSize & _
        "','" & clsOnePrinter.strFilmBox & "','" & clsOnePrinter.strResolution & "','" & clsOnePrinter.strMagnification & _
        "','" & clsOnePrinter.strSmooth & "','" & clsOnePrinter.strTrim & "','" & clsOnePrinter.strMaxDensity & _
        "','" & clsOnePrinter.strMinDensity & "','" & clsOnePrinter.strEmptyDensity & _
        "','" & clsOnePrinter.strBorderDensity & "','" & clsOnePrinter.strPolarity & "'," & clsOnePrinter.lngBitDepth & _
        ",'" & clsOnePrinter.strSCUAETitle & "'," & clsOnePrinter.lngImageBorderWidth & "," & clsOnePrinter.intImageResolution & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    'ˢ�´�ӡ���ļ���
    cDICOMPrinter.Remove (PrinterName)
    cDICOMPrinter.Add clsOnePrinter, clsOnePrinter.strname
    subFillMSFPrinter       'ˢ�¿ؼ���ʾ
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub
   
Private Sub cmdLabelConfig_Click(Index As Integer)
    Me.dlgColor.Color = Me.shpLabelConfig(Index).FillColor
    Me.dlgColor.ShowColor
    Me.shpLabelConfig(Index).FillColor = Me.dlgColor.Color
    blnInterfaceParaModified = True
End Sub

Private Sub cboLabelLineStyle_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub chkMeasureResult_LostFocus(Index As Integer)
    blnInterfaceParaModified = True
End Sub

Private Sub chkAnatomicMarkers_LostFocus(Index As Integer)
    blnInterfaceParaModified = True
End Sub

Private Sub chkChinaMark_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub chkRulerDsip_LostFocus(Index As Integer)
    blnInterfaceParaModified = True
End Sub

Private Sub lstTextoOff_Scroll(Index As Integer)
    blnInterfaceParaModified = True
End Sub

Private Sub lstRulerSize_Scroll(Index As Integer)
    blnInterfaceParaModified = True
End Sub

Private Sub lstRulerLineWidth_Scroll()
    blnInterfaceParaModified = True
End Sub

Private Sub chkLabelText_Click(Index As Integer)
    blnInterfaceParaModified = True
End Sub

Private Sub subFillUserInterface()
'------------------------------------------------
'���ܣ����������ý���Ŀؼ�����
'��������
'���أ���
'�ϼ���������̣�frmSysConfig.Form_Resize
'�¼���������̣���
'���õ��ⲿ������
'�����ˣ��ƽ�
'------------------------------------------------
    '��������ɫ
    Me.shpUserInterface(1).FillColor = lngSelectedImageBorderColor   ''ѡ��ͼ��߿���ɫ
    Me.shpUserInterface(2).FillColor = lngCurrentImageBorderColor    ''��ǰͼ��߿���ɫ
    Me.shpUserInterface(3).FillColor = lngCurrentSeriesBorderColor   ''��ǰ��δѡ�У����б߿���ɫ
    Me.shpUserInterface(4).FillColor = lngSelectImageForeColour       ''ѡ��ͼ���ʶ���ɫ
    Me.shpUserInterface(5).FillColor = lngPeriodColor                ''ѡ������ɫ
    Me.shpUserInterface(6).FillColor = lngReferenceLineColor         ''��λ����ɫ
    Me.shpUserInterface(7).FillColor = lngViewerBackColor            ''Viewer������ɫ
    Me.shpUserInterface(8).FillColor = lngProgramBackColor           ''���򱳾���ɫ
    '���ͼ��ѡ�������
    cboNoSelectLineStyle.ListIndex = lngSelectedImageBorderLineStyle ''ѡ��ͼ��߿�����
    lstNoSelectLineWidth = lngSelectedImageBorderLineWidth          ''ѡ��ͼ��߿��߿��
    cboSelectLineStyle.ListIndex = lngCurrentImageBorderLineStyle   ''��ǰͼ��߿�����
    lstSelectLineWidth = lngCurrentImageBorderLineWidth             ''��ǰͼ��߿��߿��
    lstImageIdentifierSize = lngImageIdentifierSize                 ''ͼ��ѡ���Ǵ�С
    lstPeriodSize = intPeriodSize                                   ''ѡ������С
    cboReferenceLineStyle.ListIndex = lngReferenceLineStyle         ''��λ������
    Me.lstReferenceLineSpacing = lngReferenceLineSpacing            ''��λ�߼��
    
    lstSpaceSize = intSpaceSize                            ''����֮��ļ����ȡ��߶�
    lstMaxAreaX = intMaxAreaX                              ''�������ɻ��ֵ�����
    lstMaxAreaY = intMaxAreaY                              ''�������ɻ��ֵ�����
    lstCellSpacing = lngCellSpacing                        ''ͼ����
    chkDsipSpilthBorder = IIf(blnDsipSpilthBorder = True, 1, 0)      ''����߿��Ƿ���ʾ
    chkDockMiniImage = IIf(blnDockMiniImage = True, 1, 0)           ''����ͼ�Ƿ�ͣ���ڲ˵���
    chkShowMiniImageInfo = IIf(blnShowMiniImageInfo = True, 1, 0)   ''����ͼ���Ƿ���ʾϵͳ��Ϣ
    chkSquareFrame = IIf(blnSquareFrame = True, 1, 0)               ''�����ο�ѡ
    chkShowMPRLine = IIf(blnShowMPRLine = True, 1, 0)               ''MPR����ʱ����ʾλ�ø�����
    
    chkShowFilmConfig = IIf(bShowFilmConfig = True, 1, 0)           ''�Ƿ�ֱ�����࣬����ʾ��Ƭ���ô���
    lstStatusBarFontSize.Text = intStatusBarFontSize                ''״̬�������С
    chkShowPrintTag = IIf(blnShowPrintTag = True, 1, 0)             ''�Ƿ���ʾ��Ƭ��ӡ���
    chkPrintFilmBeep = IIf(blnPrintFilmBeep = True, 1, 0)           ''��Ƭ��ӡʱ�Ƿ���ʾ������������ӽ�Ƭ����ӡ
    
    '��������ɫ
    Me.shpLabelConfig(1).FillColor = lngLabelColor               ''��ע��ʾɫ����ɫ
    Me.shpLabelConfig(2).FillColor = lngLabelSelectedColor       ''��עѡ��ɫ����ɫ
    Me.shpLabelConfig(3).FillColor = lngRulerLeftColor           ''�����ɫ
    '��ע����
    Me.opWinWLLocation(lngWinWidthLevelLocation) = True
    Me.cboLabelLineStyle.ListIndex = lngLabelLineStyleNorm
    Me.txtLabelLineWidth = lngLabelLineWidthNorm
    Me.txtLabelFontSize = lngLabelFontSize
    '�������ֵ���ʾ����
    Me.chkMeasureResult(1) = IIf(bROIArea = True, 1, 0)                 ''��ʾ���
    Me.chkMeasureResult(2) = IIf(bROIMean = True, 1, 0)                 ''��ʾƽ��ֵ
    Me.chkMeasureResult(3) = IIf(bROIStandardDeviation = True, 1, 0)    ''��ʾ������
    Me.chkMeasureResult(4) = IIf(bROILength = True, 1, 0)               ''��ʾ�ܳ�
    Me.chkMeasureResult(5) = IIf(bROIMax = True, 1, 0)                  ''��ʾ���ֵ
    Me.chkMeasureResult(6) = IIf(bROIMin = True, 1, 0)                  ''��ʾ��Сֵ
    Me.lstTextoOff(1) = intTextoOffX                             ''��ע���ֵ�ƫ����
    Me.lstTextoOff(2) = intTextoOffY                             ''��ע���ֵ�ƫ����
    Me.chkLabelText(1) = IIf(blnLabelTextScaleFontSize = True, 1, 0)      ''��ע���ִ�С�Ƿ�����ͼ��һ������
    Me.chkLabelText(2) = IIf(bROITextChinese = True, 1, 0)                ''���������Ϣ�Ƿ�ʹ������
    
    '��λ�������
    Me.chkAnatomicMarkers(1) = IIf(blnAnatomicMarkersTop = True, 1, 0)       ''�Ƿ���ʾ�����λ���
    Me.chkAnatomicMarkers(2) = IIf(blnAnatomicMarkersBottom = True, 1, 0)       ''�Ƿ���ʾ�ұ���λ���
    Me.chkAnatomicMarkers(3) = IIf(blnAnatomicMarkersLeft = True, 1, 0)      ''�Ƿ���ʾ�ϱ���λ���
    Me.chkAnatomicMarkers(4) = IIf(blnAnatomicMarkersRight = True, 1, 0)      ''�Ƿ���ʾ�±���λ���
    Me.chkChinaMark = IIf(blnChinaMark = True, 1, 0)                    ''�Ƿ���ú�����ʾ��λ���
    
    '�������
    Me.chkRulerDsip(1) = IIf(blnRulerDsipTop = True, 1, 0)              ''�Ƿ���ʾ�ϱ߱��
    Me.chkRulerDsip(2) = IIf(blnRulerDsipBottom = True, 1, 0)           ''�Ƿ���ʾ�±߱��
    Me.chkRulerDsip(3) = IIf(blnRulerDsipLeft = True, 1, 0)             ''�Ƿ���ʾ��߱��
    Me.chkRulerDsip(4) = IIf(blnRulerDsipRight = True, 1, 0)            ''�Ƿ���ʾ�ұ߱��
    Me.lstRulerSize(1) = intRulerLeft                          ''�����߾�
    Me.lstRulerSize(2) = intRulerTop                          ''����ϱ߾�
    Me.lstRulerSize(3) = intRulerWidth                        ''��߿��
    Me.lstRulerSize(4) = intRulerHeight                       ''��߸߶�
    Me.lstRulerLineWidth = intRulerLineWidth                  ''����߿�
    
    '��������Ϣ
    Me.lstMouseStep(1) = lngStackStep
    Me.lstMouseStep(2) = lngCruiseStep
    Me.lstMouseStep(3) = lngWidthLevelStep
    Me.lstMouseStep(4) = lngZoomStep
    Me.cboMouseWheelRoll.ListIndex = intMouseWheelRoll
    Me.cboMouseWheelDrag.ListIndex = intMouseWheelDrag
    
    '��䲡����Ϣ����ʾ����
    Me.shpInfoLabel.FillColor = lngpatientInfoColor     '������Ϣ��ɫ
    Me.txtPatientInfoInVisibleSize.Text = lngPatientInfoInvisibleSize   '������Ϣ��ʾ��Сֵ
    Me.chkInfoLabelScale = IIf(blnpatientInfoScaleFontSize = True, 1, 0)   '������Ϣ��ͼ������
    Me.lstPatientInfoFontSize = lngPatientInfoFontSize                      '������Ϣ�����С
    Me.chkPatientiInfoFontBold.Value = IIf(blnPatientInfoFontBold, 1, 0)    '������Ϣ�������
    Me.chkPatientInfoFontItalic.Value = IIf(blnPatientInfoFontItalic, 1, 0) '������Ϣ����б��
    Me.txtPatientInfoFontName = strPatientInfoFontName                      '������Ϣ��������
    Me.optPatientInfoTitle(lngPatientInfoTitle + 1) = True      '������Ϣ��ͷ
    Me.chkImgContainInfo.Value = Val(zlDatabase.GetPara("����ͼ����������Ϣ", glngSys, 1289, 1))
    
End Sub

Private Sub cmdUserInterfaceColor_Click(Index As Integer)
    Me.dlgColor.Color = Me.shpUserInterface(Index).FillColor
    Me.dlgColor.ShowColor
    Me.shpUserInterface(Index).FillColor = Me.dlgColor.Color
    blnInterfaceParaModified = True
End Sub

Private Sub cboNoSelectLineStyle_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub cboSelectLineStyle_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub cboReferenceLineStyle_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub chkDsipSpilthBorder_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub lstPeriodSize_Scroll()
    blnInterfaceParaModified = True
End Sub
Private Sub lstNoSelectLineWidth_Scroll()
    blnInterfaceParaModified = True
End Sub

Private Sub lstImageIdentifierSize_Scroll()
    blnInterfaceParaModified = True
End Sub

Private Sub lstSelectLineWidth_Scroll()
     blnInterfaceParaModified = True
End Sub

Private Sub lstSpaceSize_Scroll()
    blnInterfaceParaModified = True
End Sub

Private Sub lstMaxAreaX_Scroll()
    blnInterfaceParaModified = True
End Sub

Private Sub lstMaxAreaY_Scroll()
    blnInterfaceParaModified = True
End Sub

Private Sub lstCellSpacing_Scroll()
     blnInterfaceParaModified = True
End Sub

Private Sub opWinWLLocation_Click(Index As Integer)
    blnInterfaceParaModified = True
End Sub

Private Sub lstReferenceLineSpacing_Scroll()
    blnInterfaceParaModified = True
End Sub

Private Sub chkShowFilmConfig_LostFocus()
    blnInterfaceParaModified = True
End Sub

Private Sub txtPatientInfoInVisibleSize_GotFocus()
    txtPatientInfoInVisibleSize.SelStart = 0
    txtPatientInfoInVisibleSize.SelLength = Len(txtPatientInfoInVisibleSize.Text)
End Sub

Private Sub txtPatientInfoInVisibleSize_KeyPress(KeyAscii As Integer)
    'ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtPort_GotFocus()
    txtPort.SelStart = 0
    txtPort.SelLength = Len(txtPort.Text)
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    'ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtPrinterName_GotFocus()
    txtPrinterName.SelStart = 0
    txtPrinterName.SelLength = Len(txtPrinterName.Text)
End Sub

Private Sub txtPrinterName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtRect_GotFocus(Index As Integer)
    txtRect(Index).SelStart = 0
    txtRect(Index).SelLength = Len(txtRect(Index).Text)
End Sub

Private Sub txtRect_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    'ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtRect_LostFocus(Index As Integer)
    subKeepImageShutter
End Sub

Private Sub txtWinLevel_GotFocus()
    txtWinLevel.SelStart = 0
    txtWinLevel.SelLength = Len(txtWinLevel.Text)
End Sub

Private Sub txtWinLevel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    'ֻ����������
    If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtWinWidth_GotFocus()
    txtWinWidth.SelStart = 0
    txtWinWidth.SelLength = Len(txtWinWidth.Text)
End Sub

Private Sub txtWinWidth_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    'ֻ����������
    If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    
End Sub

Private Sub txtWinWLCName_GotFocus()
    txtWinWLCName.SelStart = 0
    txtWinWLCName.SelLength = Len(txtWinWLCName.Text)
End Sub

Private Sub txtWinWLCName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtWinWLEName_GotFocus()
    txtWinWLEName.SelStart = 0
    txtWinWLEName.SelLength = Len(txtWinWLEName.Text)
End Sub

Private Sub txtWinWLEName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub udLabelFontSize_DownClick()
    If Me.txtLabelFontSize > 2 Then
        Me.txtLabelFontSize = Me.txtLabelFontSize - 1
    Else
        Me.txtLabelFontSize = 1
    End If
    blnInterfaceParaModified = True
End Sub

Private Sub udLabelFontSize_UpClick()
    If Me.txtLabelFontSize < 39 Then
        Me.txtLabelFontSize = Me.txtLabelFontSize + 1
    Else
        Me.txtLabelFontSize = 40
    End If
    blnInterfaceParaModified = True
End Sub

Private Sub udLabelLineWidth_DownClick()
    If Me.txtLabelLineWidth > 2 Then
        Me.txtLabelLineWidth = Me.txtLabelLineWidth - 1
    Else
        Me.txtLabelLineWidth = 1
    End If
    blnInterfaceParaModified = True
End Sub

Private Sub udLabelLineWidth_UpClick()
    If Me.txtLabelLineWidth < 9 Then
        Me.txtLabelLineWidth = Me.txtLabelLineWidth + 1
    Else
        Me.txtLabelLineWidth = 10
    End If
    blnInterfaceParaModified = True
End Sub

Private Sub subUpdateInterfacePara(strPara As String, vValue As Variant)
    Dim strSQL As String
    If blLocalRun = True Then
        strSQL = "update Ӱ���������� SET " & strPara & " = " & "'" & vValue & "'"
        cnAccess.Execute strSQL, , adCmdText
    Else
        strSQL = "ZL_Ӱ����������_UPDATE(" & glngUserID & ",'" & strPara & "','" & IIf(TypeName(vValue) = "Boolean", CLng(vValue), vValue) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
End Sub

Public Sub subSaveInterfacePara(Optional blnReadNewPara As Boolean = True)
'------------------------------------------------
'���ܣ����ݽ�����޸����������ϵͳ������ֵ������ϵͳ�������浽��Ӱ������������
'������blnReadNewPara --- �Ƿ�ӹ�Ƭѡ����ж�ȡ�µĲ������ã�����Ǵӹ�Ƭѡ���ⲿ���ã��������ù�����λ�ã��Ͳ���Ҫ��ȡ��Ƭѡ��Ĳ���
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim strTemp As String                   '��ʱ����
    
    On Error GoTo errh
    
    '���ϵͳ����û�б��޸ģ���ֱ���˳�
    If Not blnInterfaceParaModified Then Exit Sub
    
    '��ȡ��Ƭѡ���������õĲ���
    If blnReadNewPara Then
    
        ''ѡ��ͼ��߿���ɫ
        lngSelectedImageBorderColor = Me.shpUserInterface(1).FillColor
        ''ѡ��ͼ��߿�����
        lngSelectedImageBorderLineStyle = cboNoSelectLineStyle.ListIndex
        ''ѡ��ͼ��߿��߿��
        lngSelectedImageBorderLineWidth = lstNoSelectLineWidth.list(Me.lstNoSelectLineWidth.TopIndex)
        ''ѡ��ͼ��߿���ɫ�����ǵ�ǰ��ɫ
        lngCurrentImageBorderColor = Me.shpUserInterface(2).FillColor
        ''��ǰ��δѡ�У����б߿���ɫ
        lngCurrentSeriesBorderColor = Me.shpUserInterface(3).FillColor
        ''��ǰͼ��߿�����
        lngCurrentImageBorderLineStyle = cboSelectLineStyle.ListIndex
        ''��ǰͼ��߿��߿��
        lngCurrentImageBorderLineWidth = lstSelectLineWidth.list(Me.lstSelectLineWidth.TopIndex)
        ''ѡ��ͼ���ʶ���ɫ
        lngSelectImageForeColour = Me.shpUserInterface(4).FillColor
        ''ͼ��ѡ���Ǵ�С
        lngImageIdentifierSize = lstImageIdentifierSize.list(Me.lstImageIdentifierSize.TopIndex)
        ''ѡ������ɫ
        lngPeriodColor = Me.shpUserInterface(5).FillColor
        ''ѡ������С
        intPeriodSize = lstPeriodSize.list(Me.lstPeriodSize.TopIndex)
        ''��λ����ɫ
        lngReferenceLineColor = Me.shpUserInterface(6).FillColor
        ''��λ������
        lngReferenceLineStyle = cboReferenceLineStyle.ListIndex
        ''��λ�߼��
        lngReferenceLineSpacing = Me.lstReferenceLineSpacing.list(Me.lstReferenceLineSpacing.TopIndex)
        ''����֮��ļ����ȡ��߶�
        intSpaceSize = lstSpaceSize.list(Me.lstSpaceSize.TopIndex)
        ''�������ɻ��ֵ�����
        intMaxAreaX = lstMaxAreaX.list(Me.lstMaxAreaX.TopIndex)
        ''�������ɻ��ֵ�����
        intMaxAreaY = lstMaxAreaY.list(Me.lstMaxAreaY.TopIndex)
        ''ͼ����
        lngCellSpacing = lstCellSpacing.list(Me.lstCellSpacing.TopIndex)
        ''����߿��Ƿ���ʾ
        blnDsipSpilthBorder = IIf(chkDsipSpilthBorder = 1, True, False)
        ''Viewer������ɫ
        lngViewerBackColor = Me.shpUserInterface(7).FillColor
        ''���򱳾���ɫ
        lngProgramBackColor = Me.shpUserInterface(8).FillColor
        ''��ע��ʾɫ����ɫ
        lngLabelColor = Me.shpLabelConfig(1).FillColor
        ''��ע��������
        lngLabelLineStyleNorm = Me.cboLabelLineStyle.ListIndex
        ''��ע�����߿�
        lngLabelLineWidthNorm = Me.txtLabelLineWidth
        ''��עѡ��ɫ����ɫ
        lngLabelSelectedColor = Me.shpLabelConfig(2).FillColor
        ''��ע���ִ�С
        lngLabelFontSize = Me.txtLabelFontSize
        ''��ʾ���
        bROIArea = IIf(Me.chkMeasureResult(1) = 1, True, False)
        ''��ʾƽ��ֵ
        bROIMean = IIf(Me.chkMeasureResult(2) = 1, True, False)
        ''��ʾ������
        bROIStandardDeviation = IIf(Me.chkMeasureResult(3) = 1, True, False)
        ''���������Ϣ�Ƿ�ʹ������
        bROITextChinese = IIf(Me.chkLabelText(2) = 1, True, False)
        ''��ע���ֵ�ƫ����X
        intTextoOffX = Me.lstTextoOff(1).list(Me.lstTextoOff(1).TopIndex)
        ''��ע���ֵ�ƫ����Y
        intTextoOffY = Me.lstTextoOff(2).list(Me.lstTextoOff(2).TopIndex)
        ''��ע���ִ�С�Ƿ�����ͼ��һ������
        blnLabelTextScaleFontSize = IIf(Me.chkLabelText(1) = 1, True, False)
        ''��λ��ע
        blnAnatomicMarkersTop = IIf(Me.chkAnatomicMarkers(1) = 1, True, False)
        blnAnatomicMarkersBottom = IIf(Me.chkAnatomicMarkers(2) = 1, True, False)
        blnAnatomicMarkersLeft = IIf(Me.chkAnatomicMarkers(3) = 1, True, False)
        blnAnatomicMarkersRight = IIf(Me.chkAnatomicMarkers(4) = 1, True, False)
        ''�Ƿ���ú�����ʾ��λ���
        blnChinaMark = IIf(Me.chkChinaMark = 1, True, False)
        ''��ʾ���
        blnRulerDsipTop = IIf(Me.chkRulerDsip(1) = 1, True, False)
        blnRulerDsipBottom = IIf(Me.chkRulerDsip(2) = 1, True, False)         ''�Ƿ���ʾ�±߱��
        blnRulerDsipLeft = IIf(Me.chkRulerDsip(3) = 1, True, False)            ''�Ƿ���ʾ��߱��
        blnRulerDsipRight = IIf(Me.chkRulerDsip(4) = 1, True, False)          ''�Ƿ���ʾ�ұ߱��
        ''�����߾�
        intRulerLeft = Me.lstRulerSize(1).list(Me.lstRulerSize(1).TopIndex)
        ''����ϱ߾�
        intRulerTop = Me.lstRulerSize(2).list(Me.lstRulerSize(2).TopIndex)
        ''��߿��
        intRulerWidth = Me.lstRulerSize(3).list(Me.lstRulerSize(3).TopIndex)
        ''��߸߶�
        intRulerHeight = Me.lstRulerSize(4).list(Me.lstRulerSize(4).TopIndex)
        
        ''����߿�
        intRulerLineWidth = Me.lstRulerLineWidth.list(Me.lstRulerLineWidth.TopIndex)
        ''�����ɫ
        lngRulerLeftColor = Me.shpLabelConfig(3).FillColor
        ''����λλ��
        For i = 1 To 4
            If Me.opWinWLLocation(i) = True Then
                strTemp = i
                Exit For
            End If
        Next
        lngWinWidthLevelLocation = strTemp
        ''��괩�󲽳�
        lngStackStep = Me.lstMouseStep(1).list(Me.lstMouseStep(1).TopIndex)
        ''������β���
        lngCruiseStep = Me.lstMouseStep(2).list(Me.lstMouseStep(2).TopIndex)
        ''����������
        lngWidthLevelStep = Me.lstMouseStep(3).list(Me.lstMouseStep(3).TopIndex)
        ''������Ų���
        lngZoomStep = Me.lstMouseStep(4).list(Me.lstMouseStep(4).TopIndex)
        ''�����ֲ���
        intMouseWheelRoll = cboMouseWheelRoll.ListIndex
        ''������Ϣ��ɫ
        lngpatientInfoColor = Me.shpInfoLabel.FillColor
        ''������Ϣ��ʾ��Сֵ
        lngPatientInfoInvisibleSize = Me.txtPatientInfoInVisibleSize.Text
        ''������Ϣ��ͼ������
        blnpatientInfoScaleFontSize = IIf(Me.chkInfoLabelScale = 1, True, False)
        ''������Ϣ�����С
        lngPatientInfoFontSize = Me.lstPatientInfoFontSize.list(Me.lstPatientInfoFontSize.TopIndex)
        ''������Ϣ��������
        strPatientInfoFontName = Me.txtPatientInfoFontName
        ''������Ϣ�������
        blnPatientInfoFontBold = IIf(Me.chkPatientiInfoFontBold.Value = 1, True, False)
        ''������Ϣ����б��
        blnPatientInfoFontItalic = IIf(Me.chkPatientInfoFontItalic.Value = 1, True, False)
        ''������Ϣ��ͷ
        For i = 1 To 3
            If Me.optPatientInfoTitle(i) = True Then
                strTemp = i - 1
                Exit For
            End If
        Next
        lngPatientInfoTitle = strTemp
        ''�Ƿ�ֱ�����࣬����ʾ��Ƭ���ô���
        bShowFilmConfig = IIf(chkShowFilmConfig = 1, True, False)
        ''״̬�������С
        intStatusBarFontSize = lstStatusBarFontSize.list(Me.lstStatusBarFontSize.TopIndex)
        ''��ʾ�ܳ�
        bROILength = CInt(IIf(Me.chkMeasureResult(4) = 1, True, False))
        ''��ʾ���ֵ
        bROIMax = CInt(IIf(Me.chkMeasureResult(5) = 1, True, False))
        ''��ʾ��Сֵ
        bROIMin = CInt(IIf(Me.chkMeasureResult(6) = 1, True, False))
        ''�����ֲ���
        intMouseWheelRoll = cboMouseWheelRoll.ListIndex
        ''�Ƿ���ʾ��Ƭ��ӡ���
        blnShowPrintTag = IIf(chkShowPrintTag = 1, True, False)
    End If
    
    zlDatabase.SetPara "����ͼ����������Ϣ", chkImgContainInfo.Value, glngSys, 1289
    
    '������������浽���ݿ���
    Call subSaveInterfaceParaIntoDB
    
    '��������������������Щ����ֻ�����ڱ���ע�����
    blnDockMiniImage = IIf(chkDockMiniImage.Value = 1, True, False)
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\" & App.EXEName & "\frmSysConfig", "����ͼͣ���ڲ˵���", blnDockMiniImage
    blnShowMiniImageInfo = IIf(chkShowMiniImageInfo.Value = 1, True, False)
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\" & App.EXEName & "\frmSysConfig", "����ͼ����ʾͼ����Ϣ", blnShowMiniImageInfo
    blnSquareFrame = IIf(chkSquareFrame.Value = 1, True, False)
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\" & App.EXEName & "\frmSysConfig", "��ѡ����ͼ", blnSquareFrame
    blnShowMPRLine = IIf(chkShowMPRLine.Value = 1, True, False)
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\" & App.EXEName & "\frmSysConfig", "MPR��ʾ������", blnShowMPRLine
    blnPrintFilmBeep = IIf(chkPrintFilmBeep.Value = 1, True, False)
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\" & App.EXEName & "\frmSysConfig", "��Ƭ��ӡ��ʾ����", blnPrintFilmBeep
    
    intMouseWheelDrag = cboMouseWheelDrag.ListIndex
    
    blnInterfaceParaModified = True
    
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub
Private Function subChkDicomPrintSetup() As Boolean
    '------------------------------------------------
    '���ܣ�������������Ƿ�Ϸ�
    '������
    '���أ�����Ƿ�Ϸ���True-�Ϸ���Fasle-���Ϸ���
    '�ϼ���������̣�cmdDICOMPrintAdd_Click
    '�¼���������̣���
    '���õ��ⲿ������aPresetWinWL
    '�����ˣ�����
    '------------------------------------------------
    '����Ƿ�Ϊ��
    If Len(Trim(txtPrinterName)) < 1 Then
        MsgBox "�������ӡ�����ơ�"
        txtPrinterName.SelStart = 0
        txtPrinterName.SelLength = Len(txtPrinterName.Text)
        txtPrinterName.SetFocus
        Exit Function
    End If
    If Len(Trim(txtAETitle)) < 1 Then
        MsgBox "������AE���ơ�"
        txtAETitle.SelStart = 0
        txtAETitle.SelLength = Len(txtAETitle.Text)
        txtAETitle.SetFocus
        Exit Function
    End If
    If Len(Trim(txtIPAddress)) < 1 Then
        MsgBox "������IP��ַ��"
        txtIPAddress.SelStart = 0
        txtIPAddress.SelLength = Len(txtIPAddress.Text)
        txtIPAddress.SetFocus
        Exit Function
    End If
    If Len(Trim(txtPort)) < 1 Then
        MsgBox "������οںš�"
        txtPort.SelStart = 0
        txtPort.SelLength = Len(txtPort.Text)
        txtPort.SetFocus
        Exit Function
    End If
    If Len(Trim(cboFilmSize.Text)) < 1 Then
        MsgBox "�����뽺Ƭ���"
        cboFilmSize.SelStart = 0
        cboFilmSize.SetFocus
        Exit Function
    End If

    '����Ƿ��������ַ����ں��Ƿ񳬳�
    If zl9ComLib.zlCommFun.StrIsValid(txtPrinterName.Text, 30, Me.hwnd, "��ӡ����") = False Then            '��ӡ����
        txtPrinterName.SelStart = 0
        txtPrinterName.SelLength = Len(txtPrinterName.Text)
        txtPrinterName.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(txtAETitle.Text, 30, Me.hwnd, "AE") = False Then                      'AE
        txtAETitle.SelStart = 0
        txtAETitle.SelLength = Len(txtAETitle.Text)
        txtAETitle.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(txtIPAddress.Text, 15, Me.hwnd, "IP") = False Then                    'IP
        txtIPAddress.SelStart = 0
        txtIPAddress.SelLength = Len(txtIPAddress.Text)
        txtIPAddress.SetFocus
        Exit Function
    End If
    
    If Len(txtPort) > 10 Then
        MsgBox "����˿ںų���,����������!", vbInformation, gstrSysName                                 '�˿ں�
        txtPort.SelStart = 0
        txtPort.SelLength = Len(txtPort)
        txtPort.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(cboFormat.Text, 50, Me.hwnd, "��ʽ") = False Then                     '��ʽ
        cboFormat.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(cboPriority.Text, 4, Me.hwnd, "���ȼ�") = False Then                  '���ȼ�
        cboPriority.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(cboMedium.Text, 20, Me.hwnd, "����") = False Then                      '����
        cboMedium.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(cboOrientation.Text, 20, Me.hwnd, "����") = False Then                '����
        cboOrientation.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(cboFilmSize.Text, 20, Me.hwnd, "���") = False Then                   '���
        cboFilmSize.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(cboFilmBox.Text, 20, Me.hwnd, "Ƭ��") = False Then                    'Ƭ��
        cboFilmBox.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(cboResolution.Text, 20, Me.hwnd, "�ֱ���") = False Then               '�ֱ���
        cboResolution.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(cboMagnification.Text, 20, Me.hwnd, "�Ŵ�ģʽ") = False Then          '�Ŵ�ģʽ
        cboMagnification.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(cboSmooth.Text, 20, Me.hwnd, "ƽ��ģʽ") = False Then                 'ƽ��ģʽ
        cboSmooth.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(cboTrim.Text, 20, Me.hwnd, "����") = False Then                       '����
        cboTrim.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(cboPolarity.Text, 20, Me.hwnd, "����") = False Then                   '����
        cboPolarity.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(cboBorderDensity.Text, 20, Me.hwnd, "�߿��ܶ�") = False Then          '�߿��ܶ�
        cboBorderDensity.SetFocus
        Exit Function
    End If
    
    If zl9ComLib.zlCommFun.StrIsValid(cboEmptyDensity.Text, 20, Me.hwnd, "�հ��ܶ�") = False Then           '�հ��ܶ�
        cboEmptyDensity.SetFocus
        Exit Function
    End If
    
    If Val(cboBitDepth.Text) <> 8 And Val(cboBitDepth.Text) <> 12 Then
        MsgBox "ͼ��λ�����ԣ�֧�ֵ�ͼ��λ��Ϊ8��12�����������롣", vbInformation, gstrSysName
        cboBitDepth.SetFocus
        Exit Function
    End If
    
    If Val(txtImageBorderWidth) < 1 Or Val(txtImageBorderWidth) > 99 Then
        MsgBox "ͼ�����������õķ�Χ�� 1-99�����������롣", vbInformation, gstrSysName
        txtImageBorderWidth.SetFocus
        Exit Function
    End If
    
    If Val(txtImageResolution) < 10 Or Val(txtImageResolution) > 999 Then
        MsgBox "ͼƬ�ֱ����������õķ�Χ�� 10-999�����������롣", vbInformation, gstrSysName
        txtImageResolution.SetFocus
        Exit Function
    End If
    
    subChkDicomPrintSetup = True
End Function

Private Sub subEnableShutterControl(blnEnable As Boolean)
    Dim i As Integer
    Me.frmShutter.Enabled = blnEnable
    For i = 0 To 2
        Me.chkShutterType(i).Enabled = blnEnable
        Me.txtCircle(i).Enabled = blnEnable
        Me.txtRect(i).Enabled = blnEnable
        Me.cmdVertices(i).Enabled = blnEnable
    Next i
    Me.txtRect(3).Enabled = blnEnable
    For i = 0 To 7
        Me.Label53(i).Enabled = blnEnable
    Next i
    If blnEnable Then
        Me.lstVertices.ForeColor = vbWindowText
    Else
        Me.lstVertices.ForeColor = vbGrayText
    End If
    Me.cmdShutterColor(0).Enabled = blnEnable
    Me.cmdShutterColor(1).Enabled = blnEnable
End Sub
Sub subUpdateIcon(f As frmViewer)
    '------------------------------------------------
    '���ܣ�                                  ���µ�ǰͼ������Ҽ�
    '������
    '           f                            ������
    '���أ�                                  ��
    '2009��
    '------------------------------------------------
    Select Case intToolBarIconSize
        Case 16
            BarterIco f.ImgList16
            f.ComToolBar.AddImageList f.ImgList16
        Case 24
            BarterIco f.ImgList24
            f.ComToolBar.AddImageList f.ImgList24
        Case 32
            BarterIco f.ImgList32
            f.ComToolBar.AddImageList f.ImgList32
    End Select
    f.ComToolBar.RecalcLayout
End Sub

Private Sub subLoadUserInfo()
    '�õ������û���Ϣ,���ڶ��ű�����д�û�ֻʹ������һ�ű���ȷ���û���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim objItem As ListItem
    Dim i As Integer
    
    strSQL = "Select b.Id, b.����, b.רҵ����ְ�� From Ӱ���������� a, ��Ա�� b, " & _
             " (Select Distinct ��Աid  From (Select ����ID From ������Ա Where ��Աid = [1]) a,������Ա b Where a.����ID = b.����ID) c " & _
             " Where a.��Աid = b.Id And b.Id = c.��Աid And b.Id <> [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngUserID)
    i = 1
    Do Until rsTmp.EOF
        Set objItem = Me.livGetUserSetup.ListItems.Add(, "A" & rsTmp!Id, i)
        objItem.SubItems(1) = rsTmp!����
        objItem.SubItems(2) = NVL(rsTmp!רҵ����ְ��)
        i = i + 1
        rsTmp.MoveNext
    Loop
    If Me.livGetUserSetup.ListItems.Count <= 0 Then
        Me.CmdGetUserInfo.Enabled = False
    End If
End Sub

Sub subFillShutter()
    Dim i As Integer
    '��ʼ'''''''''''''''''''���ͼ����������''''''''''''''
    '���Ӱ����������б�
    Me.cboImageShutter.Clear
    For i = 1 To UBound(aImageShutter)
        Me.cboImageShutter.AddItem aImageShutter(i).strModality
    Next i
    If Me.cboImageShutter.ListCount > 0 Then
        Me.cboImageShutter.ListIndex = 0
    Else
        Me.cboImageShutter.ListIndex = -1
    End If
    
    '���á��޸���𡱺͡�ɾ����𡱵Ŀ�����
    If cboImageShutter.ListCount = 0 Then
        cmdShutterImgType(1).Enabled = False
        cmdShutterImgType(2).Enabled = False
    Else
        cmdShutterImgType(1).Enabled = True
        cmdShutterImgType(2).Enabled = True
    End If
    
    cboImageShutter_Click
    '����'''''''''''''''''''���ͼ����������''''''''''''''
End Sub
Sub subFillWWModality()
    '��ʼ''''''''''''��䴰��λ����''''''''''''''''''''
    Dim i As Integer
    Dim intModality As Integer
    '���Ӱ�������б�
    Me.cboWWModality.Clear
    For i = 1 To UBound(aPresetWinWL, 2)
        Me.cboWWModality.AddItem aPresetWinWL(3, i).strModality
    Next i
    If Me.cboWWModality.ListCount > 0 Then
        Me.cboWWModality.ListIndex = 0
        intModality = 1
    Else
        Me.cboWWModality.ListIndex = -1
        intModality = 0
    End If
    
    '���á��޸�Ӱ����𡱰�ť�Ƿ����
    If Me.cboWWModality.ListCount = 0 Then
        cmdModifyWWModality.Enabled = False
    Else
        cmdModifyWWModality.Enabled = True
    End If
    
    subFillMSFModality intModality               ''���MSF���ݱ��
    '����''''''''''''��䴰��λ����''''''''''''''''''''
End Sub
Sub subFillLayoutModality()
    Dim i As Integer
    '��ʼ'''''''''''''''''''''''''���Ԥ��ͼ�񲼾ֽ���'''''''''''''''''''
    '���Ӱ����������б�
    Me.cboLayoutModality.Clear
    For i = 1 To UBound(aPresetLayout)
        Me.cboLayoutModality.AddItem aPresetLayout(i).strModality
    Next
    If Me.cboLayoutModality.ListCount > 0 Then
        Me.cboLayoutModality.ListIndex = 0
    Else
        Me.cboLayoutModality.ListIndex = -1
    End If
    
    '���á��޸���𡱣���ɾ��Ӱ����𡱰�ť�Ŀ�����
    If cboLayoutModality.ListCount = 0 Then
        cmdModifyLayoutModality.Enabled = False
        cmdDelLayoutModality.Enabled = False
    Else
        cmdModifyLayoutModality.Enabled = True
        cmdDelLayoutModality.Enabled = True
    End If
    
    cboLayoutModality_Click
    '����'''''''''''''''''''''''''���Ԥ��ͼ�񲼾ֽ���'''''''''''''''''''
End Sub


Public Sub SetPatientInfoFont()
'------------------------------------------------
'���ܣ���������Ի������ò����Ľ���Ϣʹ�õ����壬����ѡ���������ö�Ӧ�ؼ���ֵ
'��������
'���أ���
'------------------------------------------------

    On Error GoTo err
    dlgFont.CancelError = True '�ѵ�ȡ������������
    dlgFont.Flags = cdlCFBoth

    dlgFont.FontName = txtPatientInfoFontName.Text
    dlgFont.FontBold = IIf(chkPatientiInfoFontBold.Value = 1, True, False)
    dlgFont.FontItalic = IIf(chkPatientInfoFontItalic.Value = 1, True, False)
    dlgFont.FontSize = Val(lstPatientInfoFontSize.list(lstPatientInfoFontSize.TopIndex))
    dlgFont.ShowFont
    
    '��������
    txtPatientInfoFontName.Text = dlgFont.FontName
    chkPatientiInfoFontBold = IIf(dlgFont.FontBold, 1, 0)
    chkPatientInfoFontItalic = IIf(dlgFont.FontItalic, 1, 0)
    lstPatientInfoFontSize = dlgFont.FontSize
    
    Exit Sub
err:
    'ȡ�����ɳ���������
End Sub
