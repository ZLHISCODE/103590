VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVideoSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ɼ���������"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   Icon            =   "frmVideoSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin ScanLibCtl.ImgScan imageScannerConfig 
      Left            =   3015
      Top             =   5325
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�� ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4770
      TabIndex        =   37
      Top             =   5385
      Width           =   1100
   End
   Begin TabDlg.SSTab stbConfig 
      Height          =   5040
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   8890
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "��������"
      TabPicture(0)   =   "frmVideoSetup.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cboBakDevice"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkAllowChangeSize"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkUseCaptureLock"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboSaveDevice"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkUseGameModeCap"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "�ɼ�����"
      TabPicture(1)   =   "frmVideoSetup.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "optDriver(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdSelectDriver"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtDriverPath"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "chkCaptureWindow"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkCaptureSound"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdParameterCfg"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "optDriver(2)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "optDriver(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "optDriver(0)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame1"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "��̤����"
      TabPicture(2)   =   "frmVideoSetup.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblItem(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "labComInterval"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblItem(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "labCaptureWay"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lblItem(3)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lblItem(2)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cbxHotKey"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtComInterval"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cboPort"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cboCommCapType"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cboAfterTagHotKey"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cboAfterHotKey"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      Begin VB.CheckBox chkUseGameModeCap 
         Caption         =   "ʹ����Ϸ�ֱ��ɼ�ģʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74805
         TabIndex        =   42
         Top             =   3800
         Width           =   2400
      End
      Begin VB.ComboBox cboAfterHotKey 
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
         ItemData        =   "frmVideoSetup.frx":0060
         Left            =   -73485
         List            =   "frmVideoSetup.frx":008B
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   2280
         Width           =   4110
      End
      Begin VB.ComboBox cboAfterTagHotKey 
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
         ItemData        =   "frmVideoSetup.frx":00C4
         Left            =   -73485
         List            =   "frmVideoSetup.frx":00EF
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   2730
         Width           =   4110
      End
      Begin VB.Frame Frame2 
         Height          =   780
         Left            =   -74805
         TabIndex        =   24
         Top             =   1770
         Width           =   5430
         Begin VB.ComboBox cboImageType 
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
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   315
            Width           =   3405
         End
         Begin VB.CheckBox chkBackstageCollect 
            Caption         =   "���ú�̨�ɼ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   210
            TabIndex        =   25
            Top             =   -15
            Width           =   1635
         End
         Begin VB.Label labCapModality 
            Caption         =   "�ɼ�Ӱ�����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   495
            TabIndex        =   27
            Top             =   345
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ɨ���������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1240
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   5415
         Begin VB.CommandButton cmdImageCompressConfig 
            Caption         =   "ѹ������(&P)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4005
            TabIndex        =   22
            Top             =   720
            Width           =   1305
         End
         Begin VB.CommandButton cmdSelectScanDevice 
            Caption         =   "�豸ѡ��(&D)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2445
            TabIndex        =   21
            Top             =   720
            Width           =   1305
         End
         Begin VB.TextBox tbxTempDir 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1800
            TabIndex        =   20
            Text            =   "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA"
            Top             =   255
            Width           =   3150
         End
         Begin VB.CommandButton cmdDirSelect 
            Caption         =   "��"
            Height          =   375
            Left            =   4920
            TabIndex        =   19
            Top             =   255
            Width           =   375
         End
         Begin VB.Label labTempDir 
            Caption         =   "ɨ���豸��ʱĿ¼"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   23
            Top             =   330
            Width           =   1695
         End
      End
      Begin VB.ComboBox cboCommCapType 
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
         ItemData        =   "frmVideoSetup.frx":0128
         Left            =   -73485
         List            =   "frmVideoSetup.frx":0135
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   930
         Width           =   4110
      End
      Begin VB.ComboBox cboPort 
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
         ItemData        =   "frmVideoSetup.frx":0157
         Left            =   -73485
         List            =   "frmVideoSetup.frx":0176
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   465
         Width           =   4110
      End
      Begin VB.TextBox txtComInterval 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73485
         TabIndex        =   15
         Text            =   "1"
         Top             =   1395
         Width           =   3810
      End
      Begin VB.OptionButton optDriver 
         Caption         =   "WDM ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   810
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optDriver 
         Caption         =   "VFW ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   13
         Top             =   810
         Width           =   1200
      End
      Begin VB.OptionButton optDriver 
         Caption         =   "TWAIN ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   12
         Top             =   840
         Width           =   1350
      End
      Begin VB.CommandButton cmdParameterCfg 
         Caption         =   "��Ƶ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Top             =   1150
         Width           =   1050
      End
      Begin VB.CheckBox chkCaptureSound 
         Caption         =   "�ɼ�������ʾ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2205
         TabIndex        =   10
         Top             =   2205
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox chkCaptureWindow 
         Caption         =   "�ɼ�������ʾ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   2205
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin VB.ComboBox cboSaveDevice 
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
         Left            =   -73380
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   4020
      End
      Begin VB.CheckBox chkUseCaptureLock 
         Caption         =   "���òɼ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74805
         TabIndex        =   7
         Top             =   3060
         Width           =   1695
      End
      Begin VB.CheckBox chkAllowChangeSize 
         Caption         =   "����ı�ɼ������С"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74805
         TabIndex        =   6
         Top             =   3420
         Width           =   2400
      End
      Begin VB.ComboBox cboBakDevice 
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
         Left            =   -73380
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   945
         Width           =   4020
      End
      Begin VB.ComboBox cbxHotKey 
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
         ItemData        =   "frmVideoSetup.frx":01AE
         Left            =   -73485
         List            =   "frmVideoSetup.frx":01D9
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1845
         Width           =   4110
      End
      Begin VB.TextBox txtDriverPath 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1440
         TabIndex        =   3
         Top             =   1160
         Width           =   2655
      End
      Begin VB.CommandButton cmdSelectDriver 
         Caption         =   "��"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   1150
         Width           =   375
      End
      Begin VB.OptionButton optDriver 
         Caption         =   "ר����Ƶ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "��̨�ɼ��ȼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   -74865
         TabIndex        =   41
         Top             =   2325
         Width           =   1260
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "��Ǹ����ȼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   -74865
         TabIndex        =   40
         Top             =   2805
         Width           =   1260
      End
      Begin VB.Label labCaptureWay 
         Caption         =   "��̤�ɼ���ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74865
         TabIndex        =   35
         Top             =   975
         Width           =   1305
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "��̤�˿�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   -74430
         TabIndex        =   34
         Top             =   480
         Width           =   840
      End
      Begin VB.Label labComInterval 
         Caption         =   "��̤ʱ����                                      ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74865
         TabIndex        =   33
         Top             =   1410
         Width           =   5460
      End
      Begin VB.Label Label2 
         Caption         =   "��Ƶ�����������ã�                  "
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   195
         TabIndex        =   32
         Top             =   510
         Width           =   1920
      End
      Begin VB.Label Label3 
         Caption         =   "�ɼ���ʾ��ʽ���ã�                  "
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   31
         Top             =   1920
         Width           =   1965
      End
      Begin VB.Label Label4 
         Caption         =   "�ɼ��洢�豸"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74805
         TabIndex        =   30
         Top             =   540
         Width           =   1380
      End
      Begin VB.Label Label7 
         Caption         =   "���ݴ洢�豸"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74805
         TabIndex        =   29
         Top             =   990
         Width           =   1305
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "�ɼ��ȼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   -74430
         TabIndex        =   28
         Top             =   1920
         Width           =   840
      End
   End
   Begin MSComDlg.CommonDialog dlgOpenDir 
      Left            =   2445
      Top             =   5265
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3570
      TabIndex        =   36
      Top             =   5385
      Width           =   1100
   End
End
Attribute VB_Name = "frmVideoSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsOK As Boolean

Private DX7 As New DirectX7
Private DxInput As DirectInput
Private DiDevEnum As DirectInputEnumDevices

Private mVideoCapture As clsVideoCapture

Public Event OnVideoDirverChange(ByVal vdtDirverType As TVideoDriverType)
Public Event OnIsUseAfterImgChange(ByVal blUse As Boolean)


'modify by tjh at 2010-01-21
Public Function ShowParameterConfig(ByRef videoCapture As clsVideoCapture, ByRef owner As Object) As Boolean
BUGEX "ShowParameterConfig 1"
    ShowParameterConfig = False
    Set mVideoCapture = videoCapture
  
    IsOK = False
BUGEX "ShowParameterConfig 2"
    Call LoadDriverType
    Call ChangeControlState(glngModule = 1294)
    Call Me.Show(1, owner)
  
    ShowParameterConfig = IsOK
  
BUGEX "ShowParameterConfig 3"
End Function

'modify by tjh at 2010-01-21
'��ȡ��ǰʹ�õ���������
Private Sub LoadDriverType()
    If mVideoCapture Is Nothing Then Exit Sub
  
BUGEX "LoadDriverType 1"
    Select Case mVideoCapture.VideoDriverType
        Case vdtTWAIN
BUGEX "LoadDriverType 2"
            optDriver(2).value = True
            Call ConfigScan(True, False)
      
        Case vdtVFW
BUGEX "LoadDriverType 3"
            optDriver(1).value = True
            Call ConfigScan(False, False)
      
        Case vdtWDM
BUGEX "LoadDriverType 4"
            optDriver(0).value = True
            Call ConfigScan(False, False)
    
        Case vdtCustom
BUGEX "LoadDriverType 5"
            optDriver(3).value = True
            Call ConfigScan(False, True)
    End Select
  
BUGEX "LoadDriverType 5"
End Sub

Private Sub cboCommCapType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub ConfigComFace(ByVal blnIsCom As Boolean)
'����com�˿����ý���
    cboCommCapType.Enabled = blnIsCom
    txtComInterval.Enabled = blnIsCom
    labCaptureWay.Enabled = blnIsCom
    labComInterval.Enabled = blnIsCom
End Sub

Private Sub cboPort_Click()
    Dim blnIsCom As Boolean
     
    blnIsCom = IIf(InStr(UCase(cboPort.Text), "COM") > 0, True, False)
    
    Call ConfigComFace(blnIsCom)
End Sub

Private Sub cboPort_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkBackstageCollect_Click()
    cboImageType.Enabled = chkBackstageCollect.value
    labCapModality.Enabled = chkBackstageCollect.value
End Sub

Private Sub cmdCancel_Click()
    IsOK = False
    
    Unload Me
End Sub

''''''''''''''''''''''''''''''''''
'ѡ��ɨ���豸����ʱͼ��洢Ŀ¼
''''''''''''''''''''''''''''''''''
Private Sub cmdDirSelect_Click()
  Dim shl As Object
  Set shl = CreateObject("Shell.application")
  
  On Error GoTo final
  
    Dim fd As Object
    Set fd = shl.BrowseForFolder(0, "ɨ���豸��ʱĿ¼ѡ��", 0, "\")
  
    If Not fd Is Nothing Then
        tbxTempDir.Text = fd.Self.Path
    End If
final:
  Set shl = Nothing
  Set fd = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
'��ʾѹ������
''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdImageCompressConfig_Click()
On Error GoTo errHandle
    Call imageScannerConfig.ShowScanPreferences
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub LoadStorageDevice()
'����洢�豸
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=1 and NVL(״̬,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.EOF Then Exit Sub

    cboSaveDevice.AddItem ""
    cboBakDevice.AddItem ""
    
    Do While Not rsTemp.EOF
        cboSaveDevice.AddItem rsTemp!�豸�� & "-" & Nvl(rsTemp!�豸��)
        cboBakDevice.AddItem rsTemp!�豸�� & "-" & Nvl(rsTemp!�豸��)
        
        If GetDeptPara(glngDepartId, "�洢�豸��", "") = rsTemp!�豸�� Then
            cboSaveDevice.ListIndex = cboSaveDevice.NewIndex
        End If
        
        If GetDeptPara(glngDepartId, "�����豸��", "") = rsTemp!�豸�� Then
            cboBakDevice.ListIndex = cboBakDevice.NewIndex
        End If
        
        rsTemp.MoveNext
    Loop
End Sub

Private Sub LoadImageDeviceType()
'����ͼ�����
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "select ����,���� from Ӱ�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.EOF Then Exit Sub

    '�����ComboBox�����ݣ��ټ���
    cboImageType.Clear
    
    Do While Not rsTemp.EOF
        cboImageType.AddItem rsTemp!���� & "-" & Nvl(rsTemp!����)
        If GetDeptPara(glngDepartId, "��̨Ӱ�����", "") = rsTemp!���� Then
            cboImageType.ListIndex = cboImageType.NewIndex
        End If
        
        rsTemp.MoveNext
    Loop

End Sub

Private Sub LoadComPort()
'����com�˿ڼ��ֱ��豸
    Dim i As Long
    
    With cboPort
        .Clear
        .AddItem "��"
        .AddItem "COM1"
        .AddItem "COM2"
        .AddItem "COM3"
        .AddItem "COM4"
        .AddItem "COM5"
        .AddItem "COM6"
        .AddItem "COM7"
        .AddItem "COM8"
    End With
    
    Set DxInput = DX7.DirectInputCreate()
    Set DiDevEnum = DxInput.GetDIEnumDevices(DIDEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)
    For i = 1 To DiDevEnum.GetCount
        cboPort.AddItem DiDevEnum.GetItem(i).GetInstanceName
    Next
End Sub

Private Sub ReadDepartmentParameter()
'��ȡ����ͨ�ò�������
    '���ú�̨�ɼ�
    chkBackstageCollect.value = Val(GetDeptPara(glngDepartId, "���ú�̨�ɼ�", 1))
    
    '����ı�ɼ������С
    chkAllowChangeSize.value = Val(GetDeptPara(glngDepartId, "����ı�ɼ������С", 1))
    
    '�ɼ�����
    chkUseCaptureLock.value = Val(GetDeptPara(glngDepartId, "���òɼ�����", 1))
End Sub

Private Sub ReadLocateParameter()
'��ȡ���ز�������(�ͻ�����صĲ�������)
On Error GoTo ErrorHand

    Dim strExeRoom As String
    Dim strDeviceNO As String, iPortNumber As Integer
    Dim iCapType As Integer
    Dim strTmp() As String
    Dim strHotKey As String
    Dim strAfterHotKey As String
    Dim strAfterTagHotKey As String
    Dim strImgMaxSize As String
    Dim intShowBigImg As Integer
    
    If IsNumeric(zlDatabase.GetPara("��̤�˿�", glngSys, glngModule, "1")) Then
        iPortNumber = Val(zlDatabase.GetPara("��̤�˿�", glngSys, glngModule, "1"))
        cboPort.ListIndex = iPortNumber
    Else
        SeekIndex cboPort, zlDatabase.GetPara("��̤�˿�", glngSys, glngModule, "")
    End If
    
    iCapType = Val(zlDatabase.GetPara("��̤�ɼ���ʽ", glngSys, glngModule, "1"))
    
    If iCapType = 0 Then
        cboCommCapType.ListIndex = 0
    ElseIf iCapType = 1 Then
        cboCommCapType.ListIndex = 1
    Else
        cboCommCapType.ListIndex = 2
    End If
    
    strHotKey = GetSetting("ZLSOFT", "����ģ��", "�ɼ��ȼ�", "F8")
    If Trim(strHotKey) = "" Then
        cbxHotKey.ListIndex = 0
    Else
        cbxHotKey.Text = strHotKey
    End If
    
    strAfterHotKey = GetSetting("ZLSOFT", "����ģ��", "��̨�ɼ��ȼ�", "F7")
    If Trim(strAfterHotKey) = "" Then
        cboAfterHotKey.ListIndex = 0
    Else
        cboAfterHotKey.Text = strAfterHotKey
    End If
    
    strAfterTagHotKey = GetSetting("ZLSOFT", "����ģ��", "��Ǹ����ȼ�", "F6")
    If Trim(strAfterTagHotKey) = "" Then
        cboAfterTagHotKey.ListIndex = 0
    Else
        cboAfterTagHotKey.Text = strAfterTagHotKey
    End If
    
    tbxTempDir.Text = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "ɨ���豸��ʱĿ¼", "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA")
    txtDriverPath.Text = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "ר����Ƶ�ɼ�")
    
    txtComInterval.Text = zlDatabase.GetPara("��̤ʱ����", glngSys, glngModule, "1")
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    chkCaptureWindow.value = zlDatabase.GetPara("�ɼ��󵯴���ʾ", glngSys, glngModule, "0")
    chkCaptureSound.value = zlDatabase.GetPara("�ɼ���������ʾ", glngSys, glngModule, "0")
    
    cmdOK.Enabled = InStr(gstrPrivs, "�ɼ���������") > 0
    cmdSelectScanDevice.Enabled = InStr(gstrPrivs, "�ɼ���������") > 0
    cmdImageCompressConfig.Enabled = InStr(gstrPrivs, "�ɼ���������") > 0
    chkUseGameModeCap.value = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "ʹ����Ϸ�ֱ��ɼ�ģʽ"))
    
    Exit Sub
ErrorHand:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub SaveDepartmentParameter()
'�������ͨ�ò�������

    '����洢�豸
    If cboSaveDevice.Text <> "" Then
        SetDeptPara glngDepartId, "�洢�豸��", Split(cboSaveDevice.Text, "-")(0)
    Else
        SetDeptPara glngDepartId, "�洢�豸��", ""
    End If
    
    '���汸���豸
    If cboBakDevice.Text <> "" Then
        SetDeptPara glngDepartId, "�����豸��", Split(cboBakDevice.Text, "-")(0)
    Else
        SetDeptPara glngDepartId, "�����豸��", ""
    End If

    '�����̨�ɼ�����
    SetDeptPara glngDepartId, "���ú�̨�ɼ�", chkBackstageCollect.value     '��̨�ɼ�
    If chkBackstageCollect.value = 1 Then
        If cboImageType.Text <> "" Then
             SetDeptPara glngDepartId, "��̨Ӱ�����", Split(cboImageType.Text, "-")(0)   '��̨Ӱ�����
        End If
    End If
        
    If glngModule = 1291 Then RaiseEvent OnIsUseAfterImgChange(chkBackstageCollect.value)
    
    '��Ƶ��С��������
    Call SetDeptPara(glngDepartId, "����ı�ɼ������С", chkAllowChangeSize.value)
    
    '�ɼ�������������
    Call SetDeptPara(glngDepartId, "���òɼ�����", chkUseCaptureLock.value)
End Sub

Private Sub SaveLocateParameter()
'���汾�ز�������(�ͻ�����صĲ�������)
On Error GoTo errhand
    If optDriver(3).value Then
        If Trim(txtDriverPath.Text) = "" Then
            MsgboxEx hwnd, "ר����Ƶ�ɼ�ʱ����ָ����Ӧ�ɼ��ӿڲ���·����", vbOKOnly, G_STR_HINT_TITLE
            Exit Sub
        End If
    End If
    
    '9������COM��,0��ʾ��ʹ���ⲿ�豸
    If cboPort.ListIndex = 0 Then
        Call zlDatabase.SetPara("��̤�˿�", "��", glngSys, glngModule)
    ElseIf cboPort.ListIndex < 9 Then
        Call zlDatabase.SetPara("��̤�˿�", cboPort.ListIndex, glngSys, glngModule)
    Else
        Call zlDatabase.SetPara("��̤�˿�", cboPort.Text, glngSys, glngModule)
    End If
    
    '���òɼ��ȼ�
    Call SaveSetting("ZLSOFT", "����ģ��", "�ɼ��ȼ�", cbxHotKey.Text)
    Call SaveSetting("ZLSOFT", "����ģ��", "��̨�ɼ��ȼ�", cboAfterHotKey.Text)
    Call SaveSetting("ZLSOFT", "����ģ��", "��Ǹ����ȼ�", cboAfterTagHotKey.Text)

    '������Ƶ�������ͣ�Ŀǰֻ��������������
    If optDriver(0).value Then Call zlDatabase.SetPara("��Ƶ��������", 0, glngSys, glngModule)
    If optDriver(1).value Then Call zlDatabase.SetPara("��Ƶ��������", 1, glngSys, glngModule)
    If optDriver(2).value Then Call zlDatabase.SetPara("��Ƶ��������", 2, glngSys, glngModule)
    If optDriver(3).value Then Call zlDatabase.SetPara("��Ƶ��������", 3, glngSys, glngModule)
    
    Call zlDatabase.SetPara("�ɼ��󵯴���ʾ", chkCaptureWindow.value, glngSys, glngModule)
    Call zlDatabase.SetPara("�ɼ���������ʾ", chkCaptureSound.value, glngSys, glngModule)
    Call zlDatabase.SetPara("��̤�ɼ���ʽ", cboCommCapType.ListIndex, glngSys, glngModule)
    Call zlDatabase.SetPara("��̤ʱ����", IIf(Val(txtComInterval.Text) = 0, 1, Val(txtComInterval.Text)), glngSys, glngModule)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "ɨ���豸��ʱĿ¼", tbxTempDir.Text)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "ר����Ƶ�ɼ�", txtDriverPath.Text)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "ʹ����Ϸ�ֱ��ɼ�ģʽ", chkUseGameModeCap.value)

    Exit Sub
errhand:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub cmdOk_Click()
  On Error GoTo errHandle
    
    '���沿�Ų�������
    Call SaveDepartmentParameter
    
    Call SaveLocateParameter
    
    IsOK = True
    
    Unload Me
    
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub cmdParameterCfg_Click()
On Error GoTo errHandle
    Call mVideoCapture.ShowCaptureParameterCfgDialog(Me)
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub cmdSelectDriver_Click()
    Dim strCustomDeviceDllName As String    'ר����Ƶ�ɼ���������
    Dim objCustomDevice As Object           'ר����Ƶ�ɼ���������
    Dim objFile As New FileSystemObject
    
    On Error GoTo errHandle
    
    dlgOpenDir.ShowOpen
    
    If dlgOpenDir.FileName <> "" Then txtDriverPath.Text = dlgOpenDir.FileName
    
    strCustomDeviceDllName = Trim(Replace(objFile.GetFileName(txtDriverPath.Text), ".dll", ""))
    
    Set objCustomDevice = CreateObject(strCustomDeviceDllName & ".cls" & strCustomDeviceDllName)
    
    If Not objCustomDevice Is Nothing Then Set objCustomDevice = Nothing
    
    Exit Sub
errHandle:
    MsgboxEx hwnd, "ָ����ר����Ƶ�ɼ��ӿڲ�����Ч�����������ã�", vbOKOnly, G_STR_HINT_TITLE
    txtDriverPath.Text = ""
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
'ɨ���豸ѡ��
''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSelectScanDevice_Click()
On Error GoTo errHandle
    Call imageScannerConfig.ShowSelectScanner
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hwnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '�������ö�
    
    '�����̤�˿�����
    Call LoadComPort
    '����洢�豸
    Call LoadStorageDevice
    '�����豸����
    Call LoadImageDeviceType
    
    '��ȡ���Ź�������
    Call ReadDepartmentParameter
    '��ȡ������������
    Call ReadLocateParameter
End Sub

Private Sub optDriver_Click(Index As Integer)
On Error GoTo errHandle
BUGEX "optDriver_Click 1"
    Select Case Index
        Case 0
BUGEX "optDriver_Click 2"
            Call ConfigScan(False, False)
      
            RaiseEvent OnVideoDirverChange(vdtWDM)
        Case 1
BUGEX "optDriver_Click 3"
            Call ConfigScan(False, False)
          
            RaiseEvent OnVideoDirverChange(vdtVFW)
        Case 2
BUGEX "optDriver_Click 4"
            Call ConfigScan(True, False)
      
            RaiseEvent OnVideoDirverChange(vdtTWAIN)
BUGEX "optDriver_Click 5"
        Case 3
            Call ConfigScan(False, True)
            RaiseEvent OnVideoDirverChange(vdtCustom)
    End Select
BUGEX "optDriver_Click 6"
    Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub ConfigScan(ByVal blnIsScan As Boolean, ByVal blnIsCustom As Boolean)
BUGEX "ConfigScan 1"
    labTempDir.Enabled = blnIsScan
    tbxTempDir.Enabled = blnIsScan
    cmdDirSelect.Enabled = blnIsScan
BUGEX "ConfigScan 2"
    cmdSelectScanDevice.Enabled = blnIsScan
    cmdImageCompressConfig.Enabled = blnIsScan
BUGEX "ConfigScan 3"
    Frame1.Enabled = blnIsScan
    cmdParameterCfg.Enabled = Not blnIsScan
BUGEX "ConfigScan 4"
    txtDriverPath.Enabled = blnIsCustom
    cmdSelectDriver.Enabled = blnIsCustom
End Sub

Private Sub txtComInterval_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtComInterval_KeyPress(KeyAscii As Integer)
    If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtComInterval_LostFocus()
    If InStr(txtComInterval.Text, "-") = 0 Then
        If Val(txtComInterval.Text) < 1 Or Val(txtComInterval.Text) > 60 Then
            MsgBox "��̤ʱ����Ӧ��1-60�뷶Χ�ڣ����������롣", vbOKOnly, G_STR_HINT_TITLE
            txtComInterval.SetFocus
        End If
    End If
End Sub

Private Function GetComboxIndex(aSource() As Variant, ByVal SeekString As String) As Long
    Dim i As Long
    
    For i = 0 To UBound(aSource, 2)
        If aSource(0, i) = SeekString Then Exit For
    Next
    If i > UBound(aSource, 2) Then i = 0
    GetComboxIndex = i
End Function

'�����ı�������ֵ
Public Sub TxtInputControl(ByRef TxtBox As TextBox, ByRef KeyAscii As Integer, ByVal intDecimalPointNum As Integer)
'txtBox���ı���ؼ�
'intDecimalPointNum��С����λ��
'KeyAscii:�����ASC

    If Chr(KeyAscii) = "." Then
        If InStr(TxtBox.Text, ".") > 0 Then KeyAscii = 0
    End If
    
    If InStr(TxtBox.Text, ".") > 0 And KeyAscii <> 8 Then
        If Len(Mid(TxtBox.Text, InStr(TxtBox.Text, ".") + 1)) >= intDecimalPointNum Then KeyAscii = 0
    End If
    
End Sub

Private Sub ChangeControlState(ByVal blnIsPathol As Boolean)
'���ܣ������Ƿ��ǲ���վ�ı伸���ؼ���״̬�����ǲ���վ����̨�ɼ���ز��ɼ������ҽ�����3����ѡ������
    If blnIsPathol Then
        Frame2.Visible = False
        chkBackstageCollect.Visible = False
        chkUseCaptureLock.Top = 1560
        chkAllowChangeSize.Top = 1920
        chkUseGameModeCap.Top = 2280
    Else
        Frame2.Visible = True
        chkBackstageCollect.Visible = True
        chkUseCaptureLock.Top = 3060
        chkAllowChangeSize.Top = 3420
        chkUseGameModeCap.Top = 3800
    End If
End Sub
