VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAppUpgradeExecute 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ϵͳ��Ǩ"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   Icon            =   "frmAppUpgradeExecute.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Timer tmrRefresh 
      Interval        =   2000
      Left            =   720
      Top             =   6600
   End
   Begin MSComDlg.CommonDialog cdgPub 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picStepInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11400
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      Begin MSComctlLib.ImageList imgStep 
         Left            =   555
         Top             =   135
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradeExecute.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradeExecute.frx":20DC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������������"
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
         Left            =   1365
         TabIndex        =   1
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "������������������������������������������������������������������������������������������������������������"
         Height          =   360
         Left            =   1365
         TabIndex        =   2
         Top             =   390
         Width           =   8790
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   480
         Top             =   60
         Width           =   720
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   13000
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   13000
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "��ʼ��Ǩ(&N)"
      Height          =   350
      Left            =   8652
      TabIndex        =   3
      Top             =   6456
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&C)"
      Height          =   350
      Left            =   10176
      TabIndex        =   4
      Top             =   6456
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   7
      Top             =   6900
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAppUpgradeExecute.frx":3C2E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16536
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "15:56"
            Key             =   "STANUM"
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
   Begin MSComctlLib.ProgressBar prgThis 
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame fraStep 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   5412
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   11412
      Begin VB.Frame frmOther 
         Caption         =   "����"
         Height          =   975
         Left            =   240
         TabIndex        =   58
         Top             =   4320
         Width           =   10935
         Begin VB.Frame fraErrOption 
            BorderStyle     =   0  'None
            Height          =   252
            Left            =   1320
            TabIndex        =   59
            Top             =   210
            Width           =   4455
            Begin VB.OptionButton optErrOption 
               Caption         =   "���Դ�Ҫ����"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   61
               Top             =   0
               Value           =   -1  'True
               Width           =   1452
            End
            Begin VB.OptionButton optErrOption 
               Caption         =   "�������д���"
               Height          =   255
               Index           =   1
               Left            =   1920
               TabIndex        =   60
               Top             =   0
               Width           =   1452
            End
         End
         Begin VB.Label lblRegist 
            AutoSize        =   -1  'True
            Caption         =   "�޸ġ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   6600
            TabIndex        =   64
            Top             =   600
            Width           =   540
         End
         Begin VB.Label lblRegFile 
            AutoSize        =   -1  'True
            Caption         =   "��Ҫ����ע������ָ��ע���룺*.zcr"
            Height          =   180
            Left            =   240
            TabIndex        =   63
            ToolTipText     =   "�ʺϵ�ǰ�汾�Ĳ���ע�������Ƹ�ʽΪ��"
            Top             =   600
            Width           =   2970
         End
         Begin VB.Label lblErrOption 
            AutoSize        =   -1  'True
            Caption         =   "������ʽ"
            Height          =   180
            Left            =   240
            TabIndex        =   62
            Top             =   240
            Width           =   1080
         End
      End
      Begin VB.Frame frmLog 
         Caption         =   "��־"
         Height          =   975
         Left            =   240
         TabIndex        =   49
         Top             =   3240
         Width           =   10935
         Begin VB.TextBox txtLogLong 
            Alignment       =   2  'Center
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   7035
            MaxLength       =   3
            TabIndex        =   53
            Text            =   "1"
            Top             =   547
            Width           =   405
         End
         Begin VB.Frame fraLogType 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1320
            TabIndex        =   50
            Top             =   503
            Width           =   4215
            Begin VB.OptionButton optLogType 
               Caption         =   "ֻ��¼δ���ԵĴ�����־"
               Height          =   255
               Index           =   1
               Left            =   1920
               TabIndex        =   52
               Top             =   60
               Value           =   -1  'True
               Width           =   2295
            End
            Begin VB.OptionButton optLogType 
               Caption         =   "��¼���д�����־"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   51
               Top             =   60
               Width           =   1935
            End
         End
         Begin VB.CheckBox chkLogLong 
            Caption         =   "��¼ִ�г���     ���ӵ�SQL���"
            Height          =   255
            Left            =   5640
            TabIndex        =   54
            Top             =   570
            Value           =   1  'Checked
            Width           =   3060
         End
         Begin VB.Label lblLog 
            AutoSize        =   -1  'True
            Caption         =   "��Ǩ��־�ļ���C:\APPSOFT\Log\��װ��Ǩ\150930_00010304062124_1645.log"
            Height          =   180
            Left            =   240
            TabIndex        =   57
            Top             =   240
            Width           =   6120
         End
         Begin VB.Label lblLogModi 
            AutoSize        =   -1  'True
            Caption         =   "�޸ġ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   6600
            TabIndex        =   56
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblLogType 
            AutoSize        =   -1  'True
            Caption         =   "��־��¼��ʽ"
            Height          =   180
            Left            =   240
            TabIndex        =   55
            Top             =   600
            Width           =   1080
         End
      End
      Begin VB.Frame frmUpOption 
         Caption         =   "����ѡ��"
         Height          =   1455
         Left            =   240
         TabIndex        =   32
         Top             =   1680
         Width           =   10935
         Begin VB.CheckBox chkParallel 
            Caption         =   "����"
            Height          =   180
            Left            =   240
            TabIndex        =   40
            Top             =   1080
            Value           =   1  'Checked
            Width           =   660
         End
         Begin VB.CheckBox chkOpt 
            Caption         =   "ִ�п�ѡ����"
            Height          =   180
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkRpt 
            Caption         =   "���뱨��"
            Height          =   180
            Left            =   240
            TabIndex        =   38
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox ckhIdxOnLine 
            Caption         =   "����������������ģʽ"
            Height          =   180
            Left            =   4710
            TabIndex        =   37
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.Frame fraImpRpt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1320
            TabIndex        =   34
            Top             =   623
            Width           =   3855
            Begin VB.OptionButton optRpt 
               Caption         =   "���嵼��"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   36
               Top             =   60
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optRpt 
               Caption         =   "ֻ��������Դ"
               Height          =   255
               Index           =   1
               Left            =   1920
               TabIndex        =   35
               Top             =   60
               Width           =   1455
            End
         End
         Begin VB.TextBox txtCpu 
            Alignment       =   2  'Center
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   33
            Text            =   "4"
            Top             =   1020
            Width           =   300
         End
         Begin MSComCtl2.UpDown udCpu 
            Height          =   300
            Left            =   3000
            TabIndex        =   48
            Top             =   1020
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   4
            BuddyControl    =   "txtCpu"
            BuddyDispid     =   196639
            OrigLeft        =   3420
            OrigTop         =   3030
            OrigRight       =   3675
            OrigBottom      =   3330
            Max             =   6
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lblRptTotal 
            AutoSize        =   -1  'True
            Caption         =   "������8�����嵼�룺4��ֻ��������Դ��2"
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   7200
            TabIndex        =   47
            Top             =   720
            Width           =   3330
         End
         Begin VB.Label lblOptTotal 
            AutoSize        =   -1  'True
            Caption         =   "������8��ִ�У�4"
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   7200
            TabIndex        =   46
            Top             =   360
            Width           =   1440
         End
         Begin VB.Label lblOptSel 
            AutoSize        =   -1  'True
            Caption         =   "�޸ġ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   6600
            TabIndex        =   45
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblRptSel 
            AutoSize        =   -1  'True
            Caption         =   "�޸ġ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   6600
            TabIndex        =   44
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblParallel 
            AutoSize        =   -1  'True
            Caption         =   "���ж�="
            Height          =   180
            Left            =   1920
            TabIndex        =   43
            Top             =   1080
            Width           =   630
         End
         Begin VB.Label lblParallelNote 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����DDL"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   960
            TabIndex        =   42
            ToolTipText     =   "����DDLֻ��������Լ���Ĵ�����Ч�����Դ������ִ��ʱ�䡣"
            Top             =   1080
            Width           =   630
         End
         Begin VB.Label lblCpuWarn 
            AutoSize        =   -1  'True
            Caption         =   "δ����4��CPU�����ܲ��У�"
            ForeColor       =   &H002222B2&
            Height          =   180
            Left            =   3240
            TabIndex        =   41
            Top             =   1080
            Visible         =   0   'False
            Width           =   2160
         End
      End
      Begin VB.Frame frmUser 
         Caption         =   $"frmAppUpgradeExecute.frx":44C0
         Height          =   1455
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   10935
         Begin VB.TextBox txtToolsPwd 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4590
            PasswordChar    =   "*"
            TabIndex        =   20
            Top             =   300
            Width           =   1725
         End
         Begin VB.CheckBox chkHisAll 
            Caption         =   "ȫ������"
            Height          =   255
            Left            =   1200
            TabIndex        =   19
            Top             =   1043
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.TextBox txtHisPwd 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4590
            PasswordChar    =   "*"
            TabIndex        =   18
            Top             =   1020
            Width           =   1725
         End
         Begin VB.TextBox txtDBAUser 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1800
            TabIndex        =   16
            Text            =   "System"
            Top             =   660
            Width           =   1725
         End
         Begin VB.TextBox txtDBAPwd 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4590
            PasswordChar    =   "*"
            TabIndex        =   17
            Top             =   660
            Width           =   1725
         End
         Begin VB.TextBox txtToolsUser 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "ZLTOOLS"
            Top             =   300
            Width           =   1725
         End
         Begin VB.Label lblHisPwd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��  ��"
            Height          =   180
            Left            =   3960
            TabIndex        =   31
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label lblHisWarn 
            AutoSize        =   -1  'True
            Caption         =   "3����ʷ��δͨ����֤��"
            ForeColor       =   &H002222B2&
            Height          =   180
            Left            =   7200
            TabIndex        =   30
            Top             =   1080
            Width           =   1890
         End
         Begin VB.Label lblHisTotal 
            AutoSize        =   -1  'True
            Caption         =   "������8��ѡ��2"
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   2280
            TabIndex        =   29
            Top             =   1080
            Width           =   1440
         End
         Begin VB.Label lblHisSel 
            AutoSize        =   -1  'True
            Caption         =   "�޸ġ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   6600
            TabIndex        =   28
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label lblToolsUser 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�û���"
            Height          =   180
            Left            =   1200
            TabIndex        =   27
            Top             =   360
            Width           =   570
         End
         Begin VB.Label lblToolsPwd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��  ��"
            Height          =   180
            Left            =   3960
            TabIndex        =   26
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblDBAUser 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�û���"
            Height          =   180
            Left            =   1200
            TabIndex        =   25
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblDBAPwd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��  ��"
            Height          =   180
            Left            =   3960
            TabIndex        =   24
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblDBA 
            AutoSize        =   -1  'True
            Caption         =   "DBA�û�"
            Height          =   180
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   630
         End
         Begin VB.Label lblTools 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblHis 
            AutoSize        =   -1  'True
            Caption         =   "��ʷ��"
            Height          =   180
            Left            =   240
            TabIndex        =   21
            Top             =   1080
            Width           =   540
         End
      End
   End
   Begin VB.Frame fraStep 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5412
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   11412
      Begin VB.TextBox txtSQL 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   5016
         Left            =   3120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   8172
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPlan 
         Height          =   5412
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   3060
         _cx             =   5397
         _cy             =   9546
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   16777215
         GridColorFixed  =   16777215
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAppUpgradeExecute.frx":44D2
         ScrollTrack     =   -1  'True
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
         OutlineBar      =   5
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
         Begin MSComctlLib.ImageList imgPlan 
            Left            =   2160
            Top             =   0
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
                  Picture         =   "frmAppUpgradeExecute.frx":44FC
                  Key             =   "Finish"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAppUpgradeExecute.frx":4A96
                  Key             =   "Doing"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAppUpgradeExecute.frx":5030
                  Key             =   "Wait"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ļ�:"
         Height          =   180
         Left            =   3120
         TabIndex        =   12
         Top             =   60
         Width           =   450
      End
   End
   Begin VB.Label lblPerCap 
      AutoSize        =   -1  'True
      Caption         =   "��ǰ����"
      Height          =   180
      Left            =   3000
      TabIndex        =   13
      Top             =   6547
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblPer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "##%"
      Height          =   180
      Left            =   8280
      TabIndex        =   5
      Top             =   6540
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   13000
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   13000
      Y1              =   6360
      Y2              =   6360
   End
End
Attribute VB_Name = "frmAppUpgradeExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'====================================================================
'==����
'====================================================================
Private mintStep As Integer '��ǰҳ��
Private Const STEP_INFO = _
    "ϵͳ��Ǩ����|��Ǩʱ�����û���֤����Ǩ����ѡ����Ǩʹ�ò������ã��Լ���־��¼�ȡ�" & _
    "||ϵͳ��Ǩ����|���ڽ�����Ǩ����ע�⵱ǰ��ʾ�Ľ�����Ϣ��������ִ�������ϸ�鿴������Ϣ�����Դ�����з���֮���ٲ�ȡ��Ӧ�Ĵ�ʩ��"
Private Enum IDX_STEP
    SI_��Ǩ���� = 0
    SI_ϵͳ��Ǩ = 1
End Enum

Private Enum ErrType
    ET_���Դ�Ҫ���� = 0
    ET_�������д��� = 1
End Enum

'���̲���
Private Const FS_��Ǩ��� = "UPCHCEK"
Private Const FS_������Ǩ = "TOOLSUP"
Private Const FS_Ӧ��ϵͳ��Ǩ = "APPUP"
Private Const FS_��ʷ����Ǩ = "HISTORYUP"
Private Const FS_����ͬ��� = "PUBSYNONYM"
Private Const FS_��̨�Զ�ҵ���� = "ZLAUTORUN"
Private Const FS_������Ч���� = "COMPILE"
Private Const FS_�������� = "ADJUSTSEQ"
Private Const FS_�������� = "REPORTUP"
Private Const FS_����ֵ�� = "HELPERMAIN"
Private Const FS_��ɫ��Ȩ = "ROLEGRANT"
Private Const FS_�ӳٽű� = "RUNAFTER"
'--��ڲ���
'Ӧ��ϵͳ��Ǩ����
Private mrsSysInfo As ADODB.Recordset '����ϵͳ״̬
Private mrsSysFiles As ADODB.Recordset '����ϵͳ����Ǩ�ļ�
Private mblnExecBef As Boolean '�Ƿ���ǰ����
'--���ز���
Private mblnOK As Boolean '�Ƿ�������ɺ��˳�
Private mstrRunModule As String '��������ת��ģ��
'--����
Private mrsHistorySpace     As ADODB.Recordset '����ϵͳ��ʷ����Ϣ
Private mrsOptionalProc     As ADODB.Recordset '����ϵͳ�Լ���ʷ��Ŀ�ѡ����
Private mrsReport           As ADODB.Recordset '����ϵͳ�ı���
Private mblnFinal           As Boolean '�Ƿ���ϵͳ��Ǩ�����հ汾
Private mblnHaveST          As Boolean '��׼���Ƿ��ڱ���������
Private mstrSysCodes        As String '����������ϵͳ��ŵ��ַ������Զ��ŷָ�
Private mstrChangeTables    As String '�������������нṹ�����ı仯�ı��Զ��ŷָ�
Private mclsRunScript       As New clsRunScript '�ű����ж���
Private mintDDLParallel     As Integer '���ж�
Private mblnInstallPLJson   As Boolean    '���ڰ�װPLJSON������
Private mblnJSONRemain      As Boolean   '����JSOn��װ����
Private mstrToolsFloder     As String  'TOOLSĿ¼
Private mdatStart           As Date '������ʼʱ��

Private mrsHisAfterSPace    As ADODB.Recordset  '���Ӻ�ִ�нű�����ʷ��
Private mrsHisAfter         As ADODB.Recordset  '��ʷ���������Ӻ���ű�
Private mrsSatistics        As ADODB.Recordset  'ͳ����Ϣ�ռ����Ӻ���ű�
Private mblnStUp35          As Boolean  '��׼���Ƿ�35֮��
Private mintToolLob         As Integer      '��λ���壬����ο� LobConst
Private Enum LobConst
    LC_DEFAULT = 0
    LC_ZLTOOLS_IS3590_OR_GREATER = 1        '�������Ƿ�35.90֮��
    LC_ZLHIS_IS3590_OR_GREATER = 2          '��׼���Ƿ�35.90֮��
    LC_ISLONGRAW = 4                        'zlRPTGraphs.ͼƬ�Ƿ���Long Raw����
    LC_ZLTOOLS_CURIS3590_OR_GREATER = 8     '��ǰ�����߰汾�Ƿ���35.90֮��
End Enum
'====================================================================
'==�����ӿ�
'====================================================================
Public Function ShowMe(frmParent As Object, ByVal rsSysInfo As ADODB.Recordset, ByVal rsSysFiles As ADODB.Recordset, Optional ByVal blnExecBef As Boolean, Optional ByRef strRunModule As String) As Boolean
 '���ܣ��������
 '    :strRunModule=�����������ת��ģ��
 '���أ��Ƿ�������ɺ��˳�
    mintToolLob = LC_DEFAULT
    Set mrsSysInfo = rsSysInfo
    Set mrsSysFiles = rsSysFiles
    mblnExecBef = blnExecBef
    mintStep = -1
    mstrRunModule = ""
    Me.Show 1, frmParent
    strRunModule = mstrRunModule
    ShowMe = mblnOK
End Function

Public Function HistoryUp(frmParent As Object, objStep As Object, ByVal lngSys As Long, ByVal strBakDB As String, ByVal strIntFile As String, ByVal strUsername As String, ByVal strPassword As String, ByVal strServer As String, ByVal strMaxVer As String, ByVal strDbLink As String) As Boolean
 '���ܣ���ʷ�ⵥ�������ӿ�
 '������objStep=��ʾ����Ķ���
 '          lngSys=ϵͳ���
 '          strIntFile=��ϵͳ�İ�װ�����ļ�
 '          strBAKDB=��ʷ����
 '          strUserName=��ʷ���û�����
 '          strPassWord=��ʷ���û�����
 '          strServer=��ʷ�������
 '          strMaxVer=��ʷ��Ŀ��汾
 '          strBakSpaceName=��ʷ��ռ���
 '          strDBLInk=DBLink����
 '���أ��Ƿ������ɹ�
 '�ù������̽�ʹ�õ�ǰ�������������,mrsSysFiles,��mclsRunScript
    Dim rsTmp As ADODB.Recordset
    Dim cnHistory As ADODB.Connection
    Dim rsUpFiles As ADODB.Recordset
    Dim rsInitFile  As ADODB.Recordset
    Dim strSteps  As String, arrStep As Variant, i As Long
    Dim strCurMax As String
    Dim strSQL As String
    
    On Error GoTo errH
    mdatStart = Now
    If strIntFile = "" Then
        MsgBox "��Ч�İ�װ�����ļ�!", vbInformation, App.Title
        Exit Function
    End If

    
    '����ʵ���������ʹ�úۼ�
    Set mclsRunScript = New clsRunScript
    If strServer = "" Then strServer = gstrServer
    If strDbLink <> "" Then
        strSQL = "Select Owner, Db_Link, Username, Host" & vbNewLine & _
                    "From All_Db_Links" & vbNewLine & _
                    "Where Owner =[1] And Username =[2] And Db_Link||'.' Like [3]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡDBLink������", gstrUserName, UCase(strUsername), UCase(strDbLink) & ".%")
        If Not rsTmp.EOF Then strServer = rsTmp!HOST & ""
    End If
    
    '���ò��������
    Call mclsRunScript.InitGlobalPara(frmParent, lngSys, False, GetLogPath(LT_��ʷ����Ǩ, strUsername), , , , , True)
    mclsRunScript.Server = strServer
    mclsRunScript.HistoryDB = strBakDB & IIf(strDbLink <> "", "(DBLINK:" & strDbLink & ")", "")
    mclsRunScript.WriteSection "��ʷ�ⵥ��������Ҫ��Ϣ"
    Set rsInitFile = ReadINIToRec(strIntFile)
    rsInitFile.Filter = "��Ŀ='ϵͳ��'"
    
    mclsRunScript.WriteLog "������ʱ�䣺" & Format(CurrentDate, "yyyy-MM-dd HH:mm:ss") & String(4, " ") & "������ʱ�䣺" & Format(Now, "yyyy-MM-dd HH:mm:ss")
    mclsRunScript.WriteLog "˵����Ϊ�˼��������ݿ�������Ľ��������½�ʹ�ñ���ʱ����Ϊ��¼��־��ʱ��"
    mclsRunScript.WriteLog "��  λ  ��  �ƣ�" & gobjRegister.zlRegInfo("��λ����", False, 0)
    mclsRunScript.WriteLog "��    ��    ����" & gstrServer
    If Not rsInitFile.EOF Then
        mclsRunScript.WriteLog "ϵ          ͳ��" & lngSys & "-" & rsInitFile!����
    End If
    mclsRunScript.WriteLog "��    ʷ    �⣺" & strBakDB
    mclsRunScript.WriteLog "Ŀ  ��  ��  ����" & strMaxVer
    
    Set cnHistory = gobjRegister.GetConnection(strServer, strUsername, strPassword, False, MSODBC, "", False)
    If cnHistory.State = adStateOpen Then
        Set rsTmp = ReadHisUpgrade(cnHistory, strUsername, False, lngSys, strDbLink <> "")
        If rsTmp Is Nothing Then
            MsgBox "��ȡ����ʷ��汾��Ϣʧ�ܣ�����ʷ���޷�������", vbInformation, App.Title
            Exit Function
        End If
        If rsTmp.RecordCount = 0 Then
            MsgBox "��ȡ����ʷ��汾��Ϣʧ�ܣ�����ʷ���޷�������", vbInformation, App.Title
            Exit Function
        End If
    Else
        Exit Function
    End If
    
    Call SetSQLTrace(strServer, strUsername, cnHistory)
    
    '���һ������Ӧ�ô�strBakDB�����Ǵ���strBakUser���������������𣬵��ǲ�Ӱ��ű���ȡ
    Set mrsSysFiles = GetUpgradeFiles(rsUpFiles, rsTmp!ϵͳ���, rsTmp!��ǰ�汾, strIntFile, rsTmp!��ֹ��Ϣ, rsTmp!��ǰ��ֹ��Ϣ, strMaxVer, , strBakDB)
    mrsSysFiles.Filter = "": mrsSysFiles.Sort = "FullSPVer"
    Do While Not mrsSysFiles.EOF
        If InStr(strSteps & ",", "," & mrsSysFiles!SPVer & ",") = 0 Then
            strSteps = strSteps & "," & mrsSysFiles!SPVer
            strCurMax = mrsSysFiles!SPVer
        End If
        mrsSysFiles.MoveNext
    Loop
    If strCurMax <> strMaxVer Then 'û�нű�����Ŀ��汾û�нű��������һ���汾������
        strSteps = strSteps & "," & strMaxVer
    End If
    
    strSteps = strSteps & "," & "��ʷ��ṹ����"
    strSteps = Mid(strSteps, 2)
    arrStep = Split(strSteps, ",")
    For i = LBound(arrStep) To UBound(arrStep)
        objStep.Text = IIf(i = UBound(arrStep), "", "��Ǩ��") & arrStep(i)
        objStep.ToolTipText = IIf(i = UBound(arrStep), "", "��Ǩ��") & arrStep(i)
        If i = UBound(arrStep) Then '��ʷ��ṹ����
            Call RepairHisDB(cnHistory, lngSys, strUsername, strServer, strBakDB, strDbLink, , True)
        Else '��Ǩ
            Call RunScriptByVersion(lngSys, arrStep(i), i = LBound(arrStep), , , True, cnHistory, strBakDB, True)
        End If
    Next
    Call mclsRunScript.WriteCSVRow(lngSys, "", mclsRunScript.HistoryDB, "", Round((DateDiff("s", mdatStart, Now)) / 60))
    Call mclsRunScript.WriteCSVRow("", "", "", "", Round((DateDiff("s", mdatStart, Now)) / 60))
    mclsRunScript.HistoryDB = ""
    Call mclsRunScript.CloseLog
    Set mclsRunScript = Nothing
    Set mrsSysFiles = Nothing
    HistoryUp = True
    Exit Function
errH:
    Call mclsRunScript.WriteCSVRow(lngSys, "", mclsRunScript.HistoryDB, "", Round((DateDiff("s", mdatStart, Now)) / 60))
    Call mclsRunScript.WriteCSVRow("", "", "", "", Round((DateDiff("s", mdatStart, Now)) / 60))
    mclsRunScript.HistoryDB = ""
    Call mclsRunScript.CloseLog
    Set mclsRunScript = Nothing
    Set mrsSysFiles = Nothing
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Function

Public Function ToolsInstallUp(frmParent As Object, objStep As Object, ByVal lngSys As Long, ByVal strInstallFile As String, ByVal strLogFile As String) As Boolean
'���ܣ�ϵͳ��װ�й����߰汾�ϵ�ʱ����Ǩ�ӿ�
'������
'       frmParent=������
'       objStep=��ʾ����Ķ���
'       lngSys=��Ҫ��װ��Ӧ��ϵͳ�����
'       strInstallFile   Ӧ��ϵͳ��װ�ű�������λ��
'       strLogFile=ϵͳ��װ��־
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim strToolsVer As String, strMaxToolsVer As String, strCurMax As String
    Dim rsIni As ADODB.Recordset
    Dim strPath As String
    Dim objSys As New Scripting.FileSystemObject
    Dim strBeforeInfo As String, strNormalInfo As String
    Dim strSteps As String, arrStep As Variant, i As Long

    On Error GoTo errH
    mintToolLob = LC_DEFAULT
    mdatStart = Now
    '1����鰲װ�����ļ�
    If Not CheckInitFile(lngSys, strInstallFile, , rsIni) Then Exit Function
    rsIni.Filter = "��Ŀ='�����߰汾��'"
    If Not rsIni.EOF Then strMaxToolsVer = rsIni!���� & ""
    '2���жϹ����ߵİ汾
    Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Ver")
    If rsTmp.EOF Then
        '���û�У��ͽ��а汾��飬��Ҫ����ǰû�а汾����
        strToolsVer = JudgeOldToolsVer
        '���Ҹ������ݿ�
        Call OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Update_Ver", strToolsVer)
    Else
        '����һ��12λ������
        strToolsVer = rsTmp("����") & ""
    End If
    '3���Ƚϰ汾���Ƿ���Ҫ����
    If VerFull(strToolsVer) >= VerFull(strMaxToolsVer) Then
        '����Ҫ�󣬲���Ҫ����
        ToolsInstallUp = True
        Exit Function
    End If
    '4����ȡ�����ű�Ŀ¼
    On Error Resume Next
    strPath = objSys.GetParentFolderName(objSys.GetParentFolderName(objSys.GetParentFolderName(strInstallFile))) & "\Tools\ZLSERVER.SQL"
    If err.Number <> 0 Then err.Clear
    If gobjFSO.FileExists(strPath) Then
        mstrToolsFloder = gobjFSO.GetParentFolderName(strPath)
    End If
    On Error GoTo errH
    If Not objSys.FileExists(strPath) Then
        MsgBox "�򿪹���ű����Ŀ¼��[��װĿ¼]\Tools������", vbInformation, gstrSysName
        Exit Function
    End If
    '��ȡ�������ϴ���Ǩ����ǰ��Ǩ����ֹ��Ϣ
    '���ZLUPGRADE�����ֶΡ���ǰִ�С�
    If CheckAndAdjustMustTable("ZLUPGRADE", "��ǰִ��", False) Then
        '��ȡ����ϵͳ�ϴ���Ǩ�Լ��ϴ���ǰ��Ǩ��Ϣ
        strSQL = "Select  ��ǰִ��, ��ֹ���, ��Ǩ���, ����汾" & vbNewLine & _
                        "From (Select ��ǰִ��, ��Ǩʱ��, ��ֹ���, ��Ǩ���, ����汾, Max(��Ǩʱ��) Over(Partition By Decode(��ǰִ��, Null, -1, 0)) ��ǰʱ��" & vbNewLine & _
                        "       From Zlupgrade Where ϵͳ is Null) a" & vbNewLine & _
                        "Where A.��Ǩʱ�� = A.��ǰʱ�� "
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ�ϴ���Ǩ��Ϣ")
        'ϵͳ�ϴ�ִ����Ǩ��Ϣ
        rsTmp.Filter = "��ǰִ��=Null"
        If Not rsTmp.EOF Then
            strNormalInfo = FormatUpgradeBreak(0, rsTmp!����汾 & "", rsTmp!��ֹ��� & "")
        Else
            strNormalInfo = FormatUpgradeBreak(0, strToolsVer)
        End If
        'ϵͳ�ϴ���ǰִ����Ǩ��Ϣ
        rsTmp.Filter = "��ǰִ��<>Null"
        If Not rsTmp.EOF Then
            strBeforeInfo = FormatUpgradeBreak(0, rsTmp!����汾 & "", rsTmp!��ֹ��� & "")
        Else
            strBeforeInfo = FormatUpgradeBreak(0, strToolsVer)
        End If
    Else
        strBeforeInfo = FormatUpgradeBreak(0, strToolsVer)
        strNormalInfo = FormatUpgradeBreak(0, strToolsVer)
    End If
    '��ȡ��Ǩ�ű�
    Set mrsSysFiles = GetUpgradeFiles(Nothing, 0, strToolsVer, strPath, strNormalInfo, strBeforeInfo, strMaxToolsVer, strCurMax, , True)
    
    
    If VerFull(strCurMax) < VerFull(strMaxToolsVer) Then
        '�ű�֧�ֵ��İ汾С��Ҫ���������İ汾����������
        MsgBox "ȱ�ٹ�����" & strMaxToolsVer & "�汾����Ǩ�ű���", vbInformation, gstrSysName
        Exit Function
    Else
        If VerFull(GetPrimaryVer(strToolsVer, True)) <= VerFull(GetPrimaryVer(strCurMax)) Then
            mrsSysFiles.Filter = "SysType=" & ST_Tools & " And FullSPVer=" & VerFull(GetPrimaryVer(strCurMax))
            If mrsSysFiles.EOF Then
                MsgBox GetLackPrimaryInfo(strCurMax), vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    '6������zltools
    Set gcnTools = GetConnection("ZLTOOLS")
    If gcnTools Is Nothing Then
        MsgBox "�޷���ZLTOOLS�û�����!", vbInformation, gstrSysName
        Exit Function
    End If
    Call CheckToolsLob(True, strToolsVer, strMaxToolsVer)
    '7�������ű�����ִ����
    '����ʵ���������ʹ�úۼ�
    Set mclsRunScript = New clsRunScript
    '���ò��������
    Call mclsRunScript.InitGlobalPara(frmParent, 0, False, strLogFile, , , , , True)
    mclsRunScript.Server = gstrServer
    mclsRunScript.WriteLog "�����߰汾�ϵͣ��޷�֧�ָð汾Ӧ��ϵͳ��װ��"
    mclsRunScript.WriteLog "�������Զ�������" & strToolsVer & "->" & strMaxToolsVer
    Set gcnSystem = gcnOracle 'ϵͳ��װ�ŵ��ù����ߵ�����������ʱgcnOracleΪDBA����
    'PLJSON��װ
    If IsCanInstallPLJson(mstrToolsFloder, mblnJSONRemain) Then
        Call InstallPLJSON(gcnSystem, mstrToolsFloder, mclsRunScript, mblnJSONRemain)
    End If
    mrsSysFiles.Filter = "": mrsSysFiles.Sort = "FullSPVer"
    Do While Not mrsSysFiles.EOF
        If InStr(strSteps & ",", "," & mrsSysFiles!SPVer & ",") = 0 Then
            strSteps = strSteps & "," & mrsSysFiles!SPVer
            strCurMax = mrsSysFiles!SPVer
        End If
        mrsSysFiles.MoveNext
    Loop
    strSteps = strSteps & "," & "������Ȩ����"
    strSteps = Mid(strSteps, 2)
    arrStep = Split(strSteps, ",")
    mclsRunScript.SysNo = 0
    For i = LBound(arrStep) To UBound(arrStep)
        objStep.Text = IIf(i = UBound(arrStep), "", "��������Ǩ��") & arrStep(i)
        objStep.ToolTipText = IIf(i = UBound(arrStep), "", "��������Ǩ��") & arrStep(i)
        If i = UBound(arrStep) Then '������Ȩ����
            gcnOracle.Execute "Update zlUpGrade Set ��ǰִ��=0 Where ��ǰִ�� = 1 And ϵͳ is Null "
            Call ReGrantForTools(gcnTools, , True)
        Else '��Ǩ
            If Not RunScriptByVersion(0, arrStep(i), i = LBound(arrStep), strToolsVer, strMaxToolsVer, , , , True) Then
                MsgBox "�������Զ�����ʧ�ܣ���鿴��־������Ӧ����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Next
    '������LOB����
    If (mintToolLob And LC_ISLONGRAW) = LC_ISLONGRAW Then        '��ȻΪLong Raw
        If (mintToolLob And (LC_ZLTOOLS_IS3590_OR_GREATER Or LC_ZLHIS_IS3590_OR_GREATER)) = (LC_ZLTOOLS_IS3590_OR_GREATER Or LC_ZLHIS_IS3590_OR_GREATER) Then     '���������׼�涼����Ҫ��
            If (mintToolLob And LC_ZLTOOLS_CURIS3590_OR_GREATER) <> LC_ZLTOOLS_CURIS3590_OR_GREATER Then
                Call AdjustToolLob
            End If
        End If
    End If
    mclsRunScript.WriteLog "�������Զ������ɹ���"
    Call mclsRunScript.WriteCSVRow(0, "", "", "", Round((DateDiff("s", mdatStart, Now)) / 60))
    Call mclsRunScript.WriteCSVRow("", "", "", "", Round((DateDiff("s", mdatStart, Now)) / 60))
    Call mclsRunScript.CloseLog
    Set mclsRunScript = Nothing
    Set mrsSysFiles = Nothing
    ToolsInstallUp = True
    Exit Function
errH:
    mclsRunScript.WriteLog "�������Զ�����ʧ�ܣ�"
    Call mclsRunScript.WriteCSVRow(0, "", "", "", Round((DateDiff("s", mdatStart, Now)) / 60))
    Call mclsRunScript.WriteCSVRow("", "", "", "", Round((DateDiff("s", mdatStart, Now)) / 60))
    Call mclsRunScript.CloseLog
    Set mclsRunScript = Nothing
    Set mrsSysFiles = Nothing
    If MsgBox("�������д����Ƿ������" & vbCrLf & "    " & err.Description, vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        Resume
    End If
End Function

'====================================================================
'==�ؼ��¼�
'====================================================================
Private Sub chkHisAll_Click()
    Call RecUpdate(mrsHistorySpace, "", "����", IIf(chkHisAll.value = 0, 0, 1))
    Call RecUpdate(mrsHistorySpace, "����=0 And ��ǰ=1", "����", 1) '��ǰ��ʷ���������
    '���¶�ȡ��ѡ�ű�
    Call ReadOptionalProc(True)
    'ˢ����ʷ�������Ϣ
    Call RefreshTotalInfo(0)
End Sub

Private Sub chkLogLong_Click()
    Call SetCtrlEnabled(chkLogLong.value = 1, txtLogLong)
End Sub

Private Sub chkOpt_Click()
    Call SetCtrlEnabled(chkOpt.value = 1, lblOptSel, lblOptTotal)
    Call RecUpdate(mrsOptionalProc, "", "ִ��", IIf(chkOpt.value = 1, 1, 0))
    Call RefreshTotalInfo(2)
    lblOptSel.Visible = (chkOpt.value = 1): lblOptTotal.Visible = (chkOpt.value = 1)
End Sub

Private Sub chkParallel_Click()
    Call SetCtrlEnabled(chkParallel.Enabled And chkParallel.value = 1, lblParallel, txtCpu, udCpu)
    lblCpuWarn.Visible = chkParallel.value = 1 And lblCpuWarn.Tag <> ""
End Sub

Private Sub chkRpt_Click()
    Call SetCtrlEnabled(chkRpt.value = 1, optRpt(0), optRpt(1), lblRptSel, lblRptTotal)
    Call RecUpdate(mrsReport, "", "��������", IIf(chkRpt.value = 1, IIf(optRpt(0).value, "!Ĭ�ϸ�������", 2), 0))
    Call RefreshTotalInfo(1)
    lblRptSel.Visible = (chkRpt.value = 1): lblRptTotal.Visible = (chkRpt.value = 1)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    Call StepSwitch(1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Dim strSysCodes As String, i As Long
    Dim blnHaveApp As Boolean '�Ƿ���Ӧ��ϵͳ��Ҫ����
    Dim strRgeErr   As String
    
    '��ֹGetText
    HookDefend txtDBAPwd.hwnd
    HookDefend txtHisPwd.hwnd
    HookDefend txtToolsPwd.hwnd
    
    Call ApplyOEM(stbThis)
    If Not mblnExecBef Then ShowFlash ("�����ռ���������Ҫ������Դ�����Ժ�")
    mrsSysInfo.Filter = "ϵͳ���<>0 And ����=1"
    blnHaveApp = mrsSysInfo.RecordCount <> 0
    '//////////////////////////////////////////////////////////////////////
    '///////////////           �������ݳ�ʼ��////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    '����ZLupgrade��Ŀ��汾�ֶΣ���ֹĿ��汾������SP���µ����ݸ��³���
    Call AdjustZLupgrade
    '��ȡ��ʷ��
    Call ReadHistorySpace
    '��ȡ����
    Call ReadImpReports
    '��ȡ��ѡ����
    Call ReadOptionalProc
    '������Ϣˢ��
    Call RefreshTotalInfo
    '�Ƿ����PLJSON��װ����
    If Not mblnExecBef Then
        mrsSysInfo.Filter = "ϵͳ���=0"
        On Error Resume Next
        mstrToolsFloder = gobjFSO.GetParentFolderName(mrsSysInfo!�����ļ� & "")
        If err.Number <> 0 Then err.Clear
        If mstrToolsFloder <> "" Then
            mblnInstallPLJson = IsCanInstallPLJson(mstrToolsFloder, mblnJSONRemain)
        End If
    End If
    '��ǰִ�в������ߴ�������
    ckhIdxOnLine.Visible = mblnExecBef: ckhIdxOnLine.value = IIf(mblnExecBef And blnHaveApp Or Not blnHaveApp, 1, 0)
    ckhIdxOnLine.Enabled = blnHaveApp
    '���ò��ж�
    Call SetCpuCount
    chkParallel.value = IIf(blnHaveApp, chkParallel.value, 0)
    chkParallel.Enabled = chkParallel.Enabled And blnHaveApp
    '��־·����ȡ
    mrsSysInfo.Filter = "����=1": mrsSysInfo.Sort = "Sort"
    For i = 0 To mrsSysInfo.RecordCount - 1
        strSysCodes = strSysCodes & "," & mrsSysInfo!ϵͳ���
        mrsSysInfo.MoveNext
    Next
    lblLog.Tag = GetLogPath(IIf(mblnExecBef, LT_��ǰ��Ǩ, LT_������Ǩ))  '����Ĭ��·��
    '��ǰע����д�����־·�����򽫸�·����Ϊ��ʼ·��,��ǰUpgradeLogDir+��ŵľͲ���ʹ��
    If gobjFile.FolderExists(GetSetting("ZLSOFT", "����ģ��", "UpgradeLogDir", "")) Then
        '�����ļ������ڣ���Ȼ������gobjFile.GetFileName����ȡ�ļ�����ֻҪ���Ǵ�
        lblLogModi.Tag = GetSetting("ZLSOFT", "����ģ��", "UpgradeLogDir", "") & "\" & gobjFile.GetFileName(lblLog.Tag)
    Else
        lblLogModi.Tag = lblLog.Tag
    End If
    lblLog.Caption = "��Ǩ��־�ļ���" & lblLogModi.Tag
    lblLog.ToolTipText = lblLogModi.Tag
    If lblLog.Width >= 8000 Then
        lblLog.Width = 8000 '��ֹ��ʧ�޸ı�ǩ
    End If
    If lblLog.Width + lblLog.Left >= lblLogModi.Left + 30 Then
        Call SetCtrlPosOnLine(False, 0, lblLog, 60, lblLogModi)
    End If
    '//////////////////////////////////////////////////////////////////////
    '/////////////// �û���֤��ؿؼ�Ĭ��ֵ////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    'ZLTOOLS
    Call CheckToolsLob
    mblnStUp35 = False
    If Not mblnExecBef Then
        mrsSysInfo.Filter = "ϵͳ���=100"
        If Not mrsSysInfo.EOF Then
            If VerFull(IIf(mrsSysInfo!Ŀ��汾 & "" = "", mrsSysInfo!ϵͳ�汾�� & "", mrsSysInfo!Ŀ��汾 & "")) >= VerFull("10.35.10") Then
                mblnStUp35 = True
            End If
        End If
    End If
'    mrsSysInfo.Filter = "����=1 And ϵͳ���=0"
    '������֤�������ڼ��ܺ���У��
    Call SetCtrlEnabled(True, lblToolsUser, lblToolsPwd, txtToolsPwd)
    txtToolsPwd.BackColor = IIf(txtToolsPwd.Enabled, &H80000005, &H8000000F)
    'ע�������
    If mblnExecBef Then
        lblRegFile.ForeColor = &H808080
        lblRegist.ForeColor = &H808080
        lblRegist.Enabled = False
    Else
        lblRegFile.ForeColor = &H80000012
        lblRegist.ForeColor = &H8000000D
        lblRegist.Enabled = True
    End If
    If mblnStUp35 Then
        lblRegFile.ToolTipText = "�ʺϵ�ǰ�汾�Ĳ���ע�������Ƹ�ʽΪ��*��10.35.99��*.zcr"
    Else
        lblRegFile.ToolTipText = "�ʺϵ�ǰ�汾�Ĳ���ע�������Ƹ�ʽΪ��*��10.34.99��*.zcr"
    End If
    If Not GetConnection("ZLTOOLS", False) Is Nothing Then
        txtToolsPwd.Text = gstrToolsPwd
    End If
    'DBA�û�
    mrsSysFiles.Filter = " FileType=" & FT_DBA
    If Not mrsSysFiles.EOF Then lblDBA.Tag = 1 '��Ǵ���DBA�ű�
    'lblDba.Tag <> "" Or mblnInstallPLJson,����Ҫ��̨�ռ�ͳ����Ϣ���ٺ�̨��������в�����֤����˴˴���Ҫ��֤����
    Call SetCtrlEnabled(True, lblDBAUser, txtDBAUser, lblDBAPwd, txtDBAPwd)
    txtDBAUser.Text = IIf(gstrSysUser = "", "System", gstrSysUser)
    If Not GetConnection("DBA", False) Is Nothing Then
        txtDBAPwd.Text = gstrSysPwd
    End If
    txtDBAUser.BackColor = IIf(txtDBAUser.Enabled, &H80000005, &H8000000F)
    txtDBAPwd.BackColor = IIf(txtDBAPwd.Enabled, &H80000005, &H8000000F)
    '//////////////////////////////////////////////////////////////////////
    '///////////////ֱ�ӵ��ÿؼ��¼���ˢ�½���////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    '����DDL��ؿؼ�����������
    Call chkParallel_Click
    '��������չʾ
    Call cmdNext_Click
    '�鿴�Ƿ�������հ汾
    If Not mblnExecBef Then
        mblnFinal = True
        mrsSysInfo.Filter = "����=1 And ϵͳ���<>0 And Ŀ��汾<>Null"
        mrsSysInfo.Sort = "ϵͳ���"
        Do While Not mrsSysInfo.EOF
            '����һ��ϵͳ������Ǩ�����հ汾���������н�ɫ��Ȩ
            If mrsSysInfo!Ŀ��汾 & "" <> mrsSysInfo!���հ汾 & "" Then
                mblnFinal = False: Exit Do
            End If
            mrsSysInfo.MoveNext
        Loop
    Else
        mblnFinal = False
    End If
    If Not mblnExecBef Then ShowFlash ("")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOK Then
        If Not cmdCancel.Enabled Then
            Cancel = 1: Exit Sub
        ElseIf mintStep < SI_ϵͳ��Ǩ Then
            If MsgBox("Ҫ�˳�ϵͳ��Ǩ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    Set mrsSysInfo = Nothing
    Set mrsSysFiles = Nothing
    Set mrsHistorySpace = Nothing
    Set mrsOptionalProc = Nothing
    Set mrsReport = Nothing
    Set mclsRunScript = Nothing
    
    Set mrsHisAfterSPace = Nothing
    Set mrsHisAfter = Nothing
    Set mrsSatistics = Nothing
End Sub


Private Sub lblRegist_Click()
    With cdgPub
        If mblnStUp35 Then
            .DialogTitle = "ѡ��ע����Ȩ�ļ�����ʽ��*��10.35.99��*.zcr"
        Else
            .DialogTitle = "ѡ��ע����Ȩ�ļ�����ʽ��*��10.34.99��*.zcr"
        End If
        .Filter = "(ע����Ȩ�ļ�)|*.zcr"
        .InitDir = gobjFile.GetParentFolderName(lblRegist.Tag)
        .Filename = gobjFile.GetFileName(lblRegist.Tag)
        .CancelError = True
        On Error GoTo errH
        .ShowOpen
        If .Filename = "" Then Exit Sub
        On Error GoTo 0
        lblRegist.Tag = .Filename
        SaveSetting "ZLSOFT", "����ģ��", "UpgradeLogDir", gobjFile.GetParentFolderName(.Filename)
        lblRegFile.Caption = "��Ǩ��־�ļ���" & lblRegist.Tag
        lblRegFile.ToolTipText = lblRegFile.Tag
        lblRegFile.Refresh
        If lblRegFile.Width >= 8000 Then
            lblRegFile.Width = 8000
        End If
        If lblRegFile.Width + lblRegFile.Left >= lblRegist.Left + 30 Then
            Call SetCtrlPosOnLine(False, 0, lblRegFile, 60, lblRegist)
        End If
    End With
    Exit Sub
errH:
End Sub

Private Sub lblHisSel_Click()
    '���¶�ȡ��ʷ����Ǩ�ļ�
    If frmAppUpgradeSel.ShowMe(Me, AST_His, mrsHistorySpace, mrsSysFiles, mblnExecBef) Then
    End If
    '���¶�ȡ��ѡ����,��ʷ�����Ҳ�д洢����
    Call ReadOptionalProc(True)
    'ˢ����ʷ�������Ϣ
    Call RefreshTotalInfo(0)
End Sub

Private Sub lblLogModi_Click()
    With cdgPub
        .DialogTitle = "ȷ����Ǩ��־�ļ�"
        .Filter = "��Ǩ��־�ļ�(*.log)|*.log"
        .Flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
        .InitDir = gobjFile.GetParentFolderName(lblLogModi.Tag)
        .Filename = gobjFile.GetFileName(lblLogModi.Tag)
        .CancelError = True
        On Error GoTo errH
        .ShowSave
        On Error GoTo 0
        lblLogModi.Tag = .Filename
        SaveSetting "ZLSOFT", "����ģ��", "UpgradeLogDir", gobjFile.GetParentFolderName(.Filename)
        lblLog.Caption = "��Ǩ��־�ļ���" & lblLogModi.Tag
        lblLog.ToolTipText = lblLogModi.Tag
        lblLog.Refresh
        If lblLog.Width >= 8000 Then
            lblLog.Width = 8000
        End If
        If lblLog.Width + lblLog.Left >= lblLogModi.Left + 30 Then
            Call SetCtrlPosOnLine(False, 0, lblLog, 60, lblLogModi)
        End If
    End With
errH:
End Sub

Private Sub lblOptSel_Click()
    If frmAppUpgradeSel.ShowMe(Me, AST_OptProc, mrsOptionalProc) Then
    End If
    Call RefreshTotalInfo(2)
End Sub

Private Sub lblRptSel_Click()
    If frmAppUpgradeSel.ShowMe(Me, AST_Report, mrsReport) Then
    End If
    Call RefreshTotalInfo(1)
End Sub

Private Sub optErrOption_Click(Index As Integer)
    If Index = ET_���Դ�Ҫ���� Then
        optErrOption(ET_�������д���).ForeColor = &H80000012
    Else
        optErrOption(ET_�������д���).ForeColor = &H80000012
        MsgBox "�������д�����ܻ����һЩ�����ܵõ���Ч����", vbInformation, gstrSysName
    End If
End Sub

Private Sub optRpt_Click(Index As Integer)
    Call RecUpdate(mrsReport, "", "��������", Index + 1)
    Call RefreshTotalInfo(1)
End Sub

Private Sub tmrRefresh_Timer()
    Me.Refresh
End Sub

Private Sub txtCpu_GotFocus()
    Call SelAll(txtCpu)
End Sub

Private Sub txtCpu_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtCpu_Validate(Cancel As Boolean)
    If val(txtCpu.Text) < udCpu.Min Then
        udCpu.value = udCpu.Min
    ElseIf val(txtCpu.Text) > udCpu.Max Then
        udCpu.value = udCpu.Max
    End If
End Sub

Private Sub txtDBAPwd_GotFocus()
    Call SelAll(txtDBAPwd)
End Sub

Private Sub txtDBAPwd_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim cnTmp As New ADODB.Connection
    Dim strErr As String
    
    On Error Resume Next
    If txtDBAPwd.Text <> "" And txtDBAUser.Text <> "" Then
        If UCase(txtDBAUser.Text) = UCase(gstrSysUser) And gstrSysPwd <> txtDBAPwd.Text Then
            MsgBox "DBA�û��������", vbInformation, gstrSysName
            txtDBAPwd.Text = ""
            Cancel = True: Exit Sub
        End If
        If UCase(txtDBAUser.Text) = UCase(gstrSysUser) And gstrSysPwd = txtDBAPwd.Text And Not gcnSystem Is Nothing Then
        
        Else
            Set cnTmp = gobjRegister.GetConnection(gstrServer, txtDBAUser.Text, txtDBAPwd.Text, False, MSODBC, strErr, False)
            If cnTmp.State = adStateClosed Then
                MsgBox strErr, vbInformation, "��֤ʧ��"
                txtDBAPwd.Text = ""
                Cancel = True: Exit Sub
            End If
            
            '����Ƿ�DBA
            If CheckIsDBA(cnTmp) = False Then
                MsgBox "���û�������DBA��ݣ�", vbExclamation, gstrSysName
                txtDBAPwd.Text = ""
                txtDBAUser.Text = ""
                txtDBAUser.SetFocus: Exit Sub
            End If
            '��ʱ������SetSQLTrace��ִ��ǰ������
            Set gcnSystem = cnTmp
            gstrSysUser = txtDBAUser.Text
            gstrSysPwd = txtDBAPwd.Text
        End If
    End If
End Sub

Private Sub txtDBAUser_GotFocus()
    Call SelAll(txtDBAUser)
End Sub

Private Sub txtDBAUser_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim cnTmp As New ADODB.Connection
    
    If UCase(txtDBAUser.Text) = UCase(gstrSysUser) And gstrSysUser <> "" Then
        txtDBAPwd.Text = gstrSysPwd
    Else
        txtDBAPwd.Text = ""
    End If
    If txtDBAPwd.Text <> "" And txtDBAUser.Text <> "" Then
        '��Ϊ���ܴ�Сд���У����ȥ����дת��
        If UCase(txtDBAUser.Text) = UCase(gstrSysUser) And gstrSysPwd <> txtDBAPwd.Text Then
            MsgBox "DBA�û��������", vbInformation, gstrSysName
             Cancel = True: Exit Sub
        End If
        If UCase(txtDBAUser.Text) = UCase(gstrSysUser) And gstrSysPwd = txtDBAPwd.Text And Not gcnSystem Is Nothing Then
            '�û�û�з����仯
        Else
            Set cnTmp = gobjRegister.GetConnection(gstrServer, txtDBAUser.Text, txtDBAPwd.Text, False, MSODBC, "", False)
            If cnTmp.State = adStateClosed Then
                Cancel = True: Exit Sub
            End If
            On Error GoTo 0
            '����Ƿ�DBA
            If CheckIsDBA(cnTmp) = False Then
                MsgBox "���û�������DBA��ݣ�", vbExclamation, gstrSysName
                txtDBAUser.SetFocus: Exit Sub
            End If
            
            '��ʱ������SetSQLTrace��ִ��ǰ������
            Set gcnSystem = cnTmp
            gstrSysUser = txtDBAUser.Text
            gstrSysPwd = txtDBAPwd.Text
        End If
    End If
End Sub

Private Sub txtHisPwd_GotFocus()
    Call SelAll(txtHisPwd)
End Sub

Private Sub txtHisPwd_Validate(Cancel As Boolean)
    Dim cnTmp As ADODB.Connection
    Dim rsTmp As ADODB.Recordset
    Dim cllBakDB As New Collection, Item As Variant, arrTmp As Variant
    Dim strMaxVer As String, strFilter As String, strTmp As String
    Dim strBakName As String
    
    If txtHisPwd.Text <> "" And txtHisPwd.Tag <> Trim(txtHisPwd.Text) Then
        mrsHistorySpace.Filter = "��֤=0"
        mrsHistorySpace.Sort = "����,������,������"
        ShowFlash ("������֤��ʷ�⣬����ȡ��ʷ����Ǩ�ű������Ժ�")
        DoEvents
        On Error Resume Next
        Do While Not mrsHistorySpace.EOF
            strTmp = mrsHistorySpace!������ & ";" & mrsHistorySpace!������ & ";" & mrsHistorySpace!DB����
            cllBakDB.Add strTmp, strTmp
            If err.Number <> 0 Then err.Clear
            mrsHistorySpace.MoveNext
        Loop
        On Error GoTo errH
        For Each Item In cllBakDB
            arrTmp = Split(Item, ";")
            
            Set cnTmp = gobjRegister.GetConnection(arrTmp(1), arrTmp(0), txtHisPwd.Text, False, MSODBC, "", False)
            If cnTmp.State = adStateOpen Then
                 '��ʱ������SetSQLTrace��ִ��ǰ������
                
                Set rsTmp = ReadHisUpgrade(cnTmp, arrTmp(0), , , arrTmp(2) <> "")
                Call RecUpdate(mrsHistorySpace, "������='" & arrTmp(0) & "' And ������='" & arrTmp(1) & "' And ��֤=0", "��֤", 1)
                rsTmp.Sort = ""
                If rsTmp.EOF Then
                    Call RecUpdate(mrsHistorySpace, "������='" & arrTmp(0) & "' And ������='" & arrTmp(1) & "'", "����", txtHisPwd.Text, "������", 0, "����ǰ����", 0, "�����", "��ʷ��ռ����ݽṹȱʧ�����޷�������")
                Else
                    Do While Not rsTmp.EOF
                        mrsHistorySpace.Filter = "ϵͳ���=" & rsTmp!ϵͳ��� & " And ������='" & arrTmp(0) & "' And ������='" & arrTmp(1) & "'"
                        Do While Not mrsHistorySpace.EOF
                            If mrsHistorySpace!��֤ = 1 Then mrsHistorySpace.Update "��֤", 2
                            strBakName = UCase(mrsHistorySpace!���� & "")
                            mrsHistorySpace.Update Array("����", "��ǰ�汾", "��ֹ��Ϣ", "��ǰ��ֹ��Ϣ"), Array(txtHisPwd.Text, rsTmp!��ǰ�汾, rsTmp!��ֹ��Ϣ, rsTmp!��ǰ��ֹ��Ϣ)
                            '�ж��ܷ���Ǩ
                            If Not IsVerSion(rsTmp!��ǰ�汾 & "") Then
                                mrsHistorySpace.Update Array("������", "�����", "����ǰ����"), Array(0, "��ʷ���ݿռ�İ汾����ʶ�����飡", 0)
                            ElseIf VerFull(rsTmp!��ǰ�汾 & "") >= VerFull(mrsHistorySpace!Ŀ��汾 & "") Then '��ʶΪ��������
                                mrsHistorySpace.Update Array("������", "�����", "����ǰ����"), Array(0, "��ʷ���ݿռ�İ汾���ڱ�����ǨĿ��汾��������Ǩ��", 0)
                            Else
                                Set mrsSysFiles = GetUpgradeFiles(mrsSysFiles, rsTmp!ϵͳ���, rsTmp!��ǰ�汾, mrsHistorySpace!�����ļ�, rsTmp!��ֹ��Ϣ, rsTmp!��ǰ��ֹ��Ϣ, mrsHistorySpace!Ŀ��汾, , strBakName)
                                '��ȡ��ǰִ�е�Ŀ��汾
                                If mblnExecBef Then
                                    strFilter = "������='" & strBakName & "' And FileType=" & FT_Before
                                    mrsSysFiles.Filter = strFilter: mrsSysFiles.Sort = "FullSPVer Desc": strMaxVer = ""
                                    If Not mrsSysFiles.EOF Then
                                        strMaxVer = mrsSysFiles!SPVer
                                        mrsSysFiles.Filter = strFilter & " And ���ð汾>'" & VerFull(rsTmp!��ǰ�汾 & "") & "'": mrsSysFiles.Sort = "FullSPVer"
                                        If Not mrsSysFiles.EOF Then
                                            mrsSysFiles.Filter = strFilter & " And FullSPVer<'" & mrsSysFiles!FullSPVer & "'": mrsSysFiles.Sort = "FullSPVer Desc"
                                            If Not mrsSysFiles.EOF Then
                                                strMaxVer = mrsSysFiles!SPVer
                                            Else
                                                strMaxVer = ""
                                                mrsHistorySpace.Update Array("����ǰ����", "��ǰ�����"), Array(0, "û�п�ִ�е���ǰ�����ű���������ǰ��Ǩ��")
                                            End If
                                        End If
                                    Else
                                        mrsHistorySpace.Update Array("����ǰ����", "��ǰ�����"), Array(0, "û����ǰ�����ű���������ǰ��Ǩ��")
                                    End If
                                    mrsHistorySpace.Update "��ǰĿ��汾", strMaxVer
                                    'ɾ������ǰִ�нű�
                                    Call RecDelete(mrsSysFiles, "������='" & strBakName & "' And FileType<>" & FT_Before)
                                    'ɾ��������ǰĿ��汾����ǰ�����ű�
                                    Call RecDelete(mrsSysFiles, strFilter & " And FullSPVer>'" & VerFull(strMaxVer) & "'")
                                End If
                            End If
                            mrsHistorySpace.MoveNext
                        Loop
                        rsTmp.MoveNext
                    Loop
                End If
                '���δ����ʷ�ռ���ע��
                Call RecUpdate(mrsHistorySpace, "��֤=1", "������", 0, "����ǰ����", 0, "�����", "��ϵͳ����ʷ�ռ�δ��ZLBakInfo��ע�ᣡ")
            End If
        Next
        txtHisPwd.Tag = Trim(txtHisPwd.Text)
        '���¶�ȡ��ѡ�ű�
        Call ReadOptionalProc(True)
        'ˢ����ʷ�������Ϣ
        Call RefreshTotalInfo(0)
        ShowFlash ("")
        Me.Refresh
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Private Sub txtLogLong_GotFocus()
    Call SelAll(txtLogLong)
End Sub

Private Sub txtLogLong_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) <= 0 Then KeyAscii = 0
End Sub

Private Sub txtLogLong_Validate(Cancel As Boolean)
    If val(txtLogLong.Text) < 1 Then txtLogLong.Text = 1
End Sub

Private Sub txtToolsPwd_GotFocus()
    Call SelAll(txtToolsPwd)
End Sub

Private Sub txtToolsPwd_Validate(Cancel As Boolean)
    Dim strErr As String
    
    If txtToolsPwd.Text <> "" Then
        If gstrToolsPwd <> "" And UCase(gstrToolsPwd) <> UCase(Trim(txtToolsPwd.Text)) Then
             MsgBox "�������������", vbInformation, gstrSysName
             txtToolsPwd.Text = ""
             Cancel = True: Exit Sub
        End If
        err.Clear: On Error Resume Next
        If gcnTools Is Nothing Then
            Set gcnTools = New ADODB.Connection
        ElseIf gcnTools.State = 1 Then
            gcnTools.Close
        End If
                
        Set gcnTools = gobjRegister.GetConnection(gstrServer, "zltools", txtToolsPwd.Text, False, MSODBC, "", False)
        If gcnTools.State = adStateClosed Then
            MsgBox "���ӹ������û�ʱ���ִ���" & vbCrLf & vbCrLf & strErr, vbCritical, gstrSysName
            txtToolsPwd.Text = ""
            Cancel = True: Exit Sub
        End If
        Call SetSQLTrace(gstrServer, "zltools", gcnTools)
        gstrToolsPwd = txtToolsPwd.Text '��ֵ
    End If
End Sub

Private Sub udCpu_Change()
    Call SelAll(txtCpu)
End Sub


'====================================================================
'==����
'====================================================================
Private Sub ReadImpReports()
'��ȡѡ������ϵͳ�Ŀɵ��뱨��
    Dim strIniPath As String
    Dim blnDo As Boolean, blnAdd As Boolean
    Dim rsIni As ADODB.Recordset
    Dim arrTmp As Variant
    Dim lngID As Long
    Dim strVer As String
    
    On Error GoTo errH
    Set mrsReport = CopyNewRec(Nothing, True, , Array("ID", adInteger, Empty, Empty, "ϵͳ���", adInteger, Empty, Empty, "ϵͳ����", adVarChar, 50, Empty, "SPVer", adVarChar, 30, Empty, "FULLSPVer", adVarChar, 30, Empty, "���", adVarChar, 20, Empty, "����", adVarChar, 30, Empty, _
                                                                                        "FilePath", adVarChar, 1000, Empty, "FileName", adVarChar, 200, Empty, "��������", adInteger, Empty, Empty, "Ĭ�ϸ�������", adInteger, Empty, Empty))
    If mblnExecBef Then Exit Sub '��ǰ��Ǩ��ֻ��ʼ����¼������
    mrsSysInfo.Filter = "����=1"
    mrsSysInfo.Sort = "ϵͳ���"
    Do While Not mrsSysInfo.EOF
        strIniPath = gobjFile.GetParentFolderName(gobjFile.GetParentFolderName(mrsSysInfo!�����ļ�)) & "\��������"
        blnDo = gobjFile.FileExists(strIniPath & "\zlReport.ini")
        If blnDo Then
            Set rsIni = ReadINIToRec(strIniPath & "\zlReport.ini")
            blnDo = rsIni.RecordCount > 0
        End If
        If blnDo Then
            Do While Not rsIni.EOF
                blnAdd = IsVerSion(rsIni!��Ŀ & "")
                If blnAdd Then
                    strVer = rsIni!��Ŀ & ""
                    blnAdd = VerFull(rsIni!��Ŀ & "") > VerFull(mrsSysInfo!ϵͳ�汾��)
                    If blnAdd Then
                        blnAdd = VerFull(rsIni!��Ŀ & "") <= VerFull(mrsSysInfo!Ŀ��汾)
                    End If
                    If blnAdd Then
                        arrTmp = Split(rsIni!����, "|")
                        blnAdd = gobjFile.FileExists(strIniPath & "\" & arrTmp(2))
                    End If
                End If
                If blnAdd Then
                    mrsReport.Filter = "���='" & UCase(arrTmp(0)) & "'"
                    blnAdd = mrsReport.EOF
                    If blnAdd Then
                        mrsReport.AddNew Array("ID", "ϵͳ���", "ϵͳ����", "SPVer", "���", "����", "FilePath", "FileName", "��������", "Ĭ�ϸ�������"), _
                                                        Array(Identity(lngID), mrsSysInfo!ϵͳ���, mrsSysInfo!ϵͳ����, strVer, UCase(Trim(arrTmp(0))), UCase(Trim(arrTmp(1))), strIniPath & "\" & arrTmp(2), arrTmp(2), IIf(val(arrTmp(3)) = 0, 1, 2), IIf(val(arrTmp(3)) = 0, 1, 2))
                    Else
                        mrsReport.Update Array("��������", "Ĭ�ϸ�������", "SPVer"), Array(IIf(val(arrTmp(3)) = 0, 1, 2), IIf(val(arrTmp(3)) = 0, 1, 2), strVer)
                    End If
                End If
                rsIni.MoveNext
            Loop
        End If
        mrsSysInfo.MoveNext
    Loop
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub ReadHistorySpace()
    Dim rsSpaces As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strServer As String
    Dim lngID As Long
    
    On Error GoTo errH
    '��Ҫ�ṹ���
    If Not CheckAndAdjustMustTable("Zlbakspaces", , True) Then
        Exit Sub
    End If
    If Not CheckAndAdjustMustTable("ZLBAKTABLES", , True) Then
        Exit Sub
    End If
    strSQL = "Select ϵͳ, ���, ����, ������, Db����, ��ǰ From Zltools.Zlbakspaces"
    Set rsSpaces = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    '������=1��ѡ��������=0����ѡ��������-1��ѡ���������Ǹı��˷�������,��״̬���м�״̬
    '��������=1,���Գ���������=0,���ܽ��г�������
    '����ǰ������=1,������ǰ������=0,���ܽ�����ǰ����
    '��֤��=0,����ʷ��δͨ����֤��=1������ʷ���û�ͨ����֤��������ʷ�ռ�δע�����ʷ�⣬=2����֤�ɹ�
    'ע����ʷ�������Ϊ��ϵͳ���,����
    Set mrsHistorySpace = CopyNewRec(Nothing, True, , Array("ID", adInteger, Empty, Empty, "ϵͳ���", adInteger, Empty, Empty, "ϵͳ����", adVarChar, 50, Empty, "ϵͳ�汾", adVarChar, 20, Empty, "�����ļ�", adVarChar, 2000, Empty, _
                                                                                                "���", adInteger, Empty, Empty, "����", adVarChar, 30, Empty, "������", adVarChar, 50, Empty, _
                                                                                                "��ǰ", adInteger, Empty, Empty, "DB����", adVarChar, 200, Empty, "����", adVarChar, 100, Empty, _
                                                                                                "������", adVarChar, 500, Empty, "����", adInteger, Empty, Empty, "��ǰ�汾", adVarChar, 20, Empty, _
                                                                                                "Ŀ��汾", adVarChar, 20, Empty, "��ֹ��Ϣ", adVarChar, 2000, Empty, "������", adInteger, 1, 0, "�����", adVarChar, 2000, Empty, _
                                                                                                "��ǰĿ��汾", adVarChar, 20, Empty, "��ǰ��ֹ��Ϣ", adVarChar, 2000, Empty, "����ǰ����", adInteger, 1, 0, "��ǰ�����", adVarChar, 2000, Empty, _
                                                                                                "��֤", adInteger, Empty, Empty))
    mrsSysInfo.Filter = "����=1"
    mrsSysInfo.Sort = "ϵͳ���"
    Do While Not mrsSysInfo.EOF
        rsSpaces.Filter = "ϵͳ=" & mrsSysInfo!ϵͳ���
        rsSpaces.Sort = "��ǰ,���"
        Do While Not rsSpaces.EOF
            strServer = gstrServer
            If rsSpaces!DB���� & "" <> "" Then
                strSQL = "Select Owner, Db_Link, Username, Host" & vbNewLine & _
                            "From All_Db_Links" & vbNewLine & _
                            "Where Owner =[1] And Username =[2] And Db_Link||'.' Like [3]"
                Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, gstrUserName, UCase(rsSpaces!������ & ""), UCase(rsSpaces!DB���� & "") & ".%")
                If Not rsTmp.EOF Then strServer = rsTmp!HOST & ""
            End If
            mrsHistorySpace.AddNew Array("ID", "ϵͳ���", "ϵͳ����", "ϵͳ�汾", "Ŀ��汾", "�����ļ�", "���", "����", "��ǰ", "������", "DB����", "����", "������", "����", "������", "����ǰ����", "��֤"), _
                                                Array(Identity(lngID), mrsSysInfo!ϵͳ���, mrsSysInfo!ϵͳ����, mrsSysInfo!ϵͳ�汾��, mrsSysInfo!Ŀ��汾, mrsSysInfo!�����ļ�, rsSpaces!���, rsSpaces!����, val(rsSpaces!��ǰ & ""), Trim(UCase(rsSpaces!������ & "")), rsSpaces!DB����, Null, UCase(strServer), 1, 1, 1, 0)
            rsSpaces.MoveNext
        Loop
        mrsSysInfo.MoveNext
    Loop
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub ReadOptionalProc(Optional ByVal blnReadHis As Boolean)
'���ܣ���ȡ��ѡ����
'������blnReadHis=�Ƕ�ȡ��ʷ��Ŀ�ѡ�洢����
    Dim arrTmp As Variant, strTmp As String
    Dim strName As String, strTip As String
    Dim lngID As Long, i As Long
    
    On Error GoTo errH
    If mrsOptionalProc Is Nothing Or Not blnReadHis Then
        Set mrsOptionalProc = CopyNewRec(Nothing, True, , Array("ID", adInteger, Empty, Empty, "ϵͳ���", adInteger, Empty, Empty, "ϵͳ����", adVarChar, 50, Empty, "ִ����", adVarChar, 100, Empty, "��ʷ��", adInteger, Empty, Empty, "SPVer", adVarChar, 30, Empty, _
                                                                                                    "����", adVarChar, 100, Empty, "FilePath", adVarChar, 2000, Empty, "ע��", adLongVarChar, 2000, Empty, "ִ��", adInteger, Empty, Empty))
        If mblnExecBef Then Exit Sub '��ǰ��Ǩ��ֻ��ʼ����¼������
        mrsSysInfo.Filter = "����=1"
        mrsSysInfo.Sort = "ϵͳ���"
        Do While Not mrsSysInfo.EOF
            '��ǰϵͳ�ķ���ʷ��Ŀ�ѡ�ű��Ĺ���
            mrsSysFiles.Filter = "SysType<>" & ST_History & " And ϵͳ���=" & mrsSysInfo!ϵͳ��� & " And FullSPVer<='" & VerFull(mrsSysInfo!Ŀ��汾) & "' And FileType=" & FT_Optional
            mrsSysFiles.Sort = "FullSPVer"
            Do While Not mrsSysFiles.EOF
                strTmp = mclsRunScript.CollectProcs(mrsSysFiles!FilePath)
                arrTmp = Split(strTmp, "?")
                For i = LBound(arrTmp) To UBound(arrTmp)
                    strName = Left(arrTmp(i), InStr(arrTmp(i), "|") - 1)
                    strTip = Mid(arrTmp(i), InStr(arrTmp(i), "|") + 1)
                    mrsOptionalProc.AddNew Array("ID", "ϵͳ���", "ϵͳ����", "ִ����", "��ʷ��", "SPVer", "����", "FilePath", "ע��", "ִ��"), _
                                                            Array(Identity(lngID), mrsSysInfo!ϵͳ���, mrsSysInfo!ϵͳ����, IIf(mrsSysInfo!ϵͳ��� = 0, "ZLTOOLS", gstrUserName), 0, mrsSysFiles!SPVer, strName, mrsSysFiles!FilePath, RemoveMark(strTip), 1)
                Next
                mrsSysFiles.MoveNext
            Loop
            mrsSysInfo.MoveNext
        Loop
    ElseIf blnReadHis Then
        If mblnExecBef Then
             '��շ������ı��־
            Call RecUpdate(mrsHistorySpace, "����=-1", "����", 1)
            Exit Sub '��ǰ��Ǩ��ֻ��ʼ����¼������
        End If
        'ɾ��������Ǩ����ʷ�⡢��ѡ����Ǩ���Լ��ı������������֤����ʷ�����Ǩ�ű�
        mrsHistorySpace.Filter = "����=0  OR ������=0 OR ��֤<>2 OR ����=-1 "
        Do While Not mrsHistorySpace.EOF 'ɾ��ȡ����ѡ����ʷ��Ŀ�ѡ����
            Call RecDelete(mrsOptionalProc, "ϵͳ���=" & mrsHistorySpace!ϵͳ��� & " And ִ����='" & UCase(mrsHistorySpace!���� & "") & "'") '��ɾ����ʷ��Ŀ�ѡ�洢����
            mrsHistorySpace.MoveNext
        Loop
        '��շ������ı��־
        Call RecUpdate(mrsHistorySpace, "����=-1", "����", 1)
        mrsOptionalProc.Filter = ""
        lngID = mrsOptionalProc.RecordCount
        mrsHistorySpace.Filter = "����=1 And ������=1 And ��֤=2" '���ӹ�ѡ��������ʷ��Ŀ�ѡ����
        Do While Not mrsHistorySpace.EOF
            mrsOptionalProc.Filter = "ϵͳ���=" & mrsHistorySpace!ϵͳ��� & " And ��ʷ��=1 And ִ����='" & mrsHistorySpace!���� & "'"
            If mrsOptionalProc.EOF Then '����ʷ��û�п�ѡ�洢���̼�¼���������ռ�
                mrsSysFiles.Filter = "ϵͳ���=" & mrsHistorySpace!ϵͳ��� & " And SysType=" & ST_History & " And FileType=" & FT_Optional
                mrsSysFiles.Sort = "FullSPVer"
                Do While Not mrsSysFiles.EOF
                    strTmp = mclsRunScript.CollectProcs(mrsSysFiles!FilePath)
                    arrTmp = Split(strTmp, "?")
                    For i = LBound(arrTmp) To UBound(arrTmp)
                        strName = Left(arrTmp(i), InStr(arrTmp(i), "|") - 1)
                        strTip = Mid(arrTmp(i), InStr(arrTmp(i), "|") + 1)
                        mrsOptionalProc.AddNew Array("ID", "ϵͳ���", "ϵͳ����", "ִ����", "��ʷ��", "SPVer", "����", "FilePath", "ע��", "ִ��"), _
                                                                Array(Identity(lngID), mrsHistorySpace!ϵͳ���, mrsHistorySpace!ϵͳ����, mrsSysFiles!������, 1, mrsSysFiles!SPVer, strName, mrsSysFiles!FilePath, RemoveMark(strTip))
                    Next
                    mrsSysFiles.MoveNext
                Loop
            End If
            mrsHistorySpace.MoveNext
        Loop
        Call RefreshTotalInfo(2) 'ˢ�¿�ѡ���̻�����Ϣ
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshTotalInfo(Optional ByVal intRefreshType As Integer = -1)
'���ܣ�ˢ�»�����Ϣ
'������intRefreshType=ˢ�����ͣ�-1�����еĻ�����Ϣˢ��, 0:ˢ����ʷ��, 1:ˢ�µ��뱨��2��ˢ�¿�ѡ����
    '��ʷ�������Ϣˢ��
    If intRefreshType = -1 Or intRefreshType = 0 Then
        mrsHistorySpace.Filter = ""
        If intRefreshType = -1 Then
            If mrsHistorySpace.RecordCount = 0 Then
                lblHisWarn.Visible = False: lblHisTotal.Visible = False: lblHisSel.Visible = False
                chkHisAll.value = 0
            End If
            Call SetCtrlEnabled(mrsHistorySpace.RecordCount <> 0, chkHisAll, lblHisPwd, txtHisPwd)
        End If
        lblHisTotal.Caption = "������" & mrsHistorySpace.RecordCount & "��ѡ��"
        mrsHistorySpace.Filter = "����=1"
        lblHisTotal.Caption = lblHisTotal.Caption & mrsHistorySpace.RecordCount
        mrsHistorySpace.Filter = "����=1 And ��֤<>2"
        lblHisWarn.Caption = mrsHistorySpace.RecordCount & "����ʷ��δͨ����֤��"
        lblHisWarn.Visible = mrsHistorySpace.RecordCount <> 0
'        If lblHisWarn.Visible Then
'            Call SetCtrlPosOnLine(False, 0, txtHisPwd, 60, lblHisWarn, 60, lblHisSel)
'        Else
'            Call SetCtrlPosOnLine(False, 0, txtHisPwd, 60, lblHisSel)
'        End If
        Call RecToLog(mrsHistorySpace, "ϵͳ���,���", IIf(intRefreshType = -1, "ԭʼ��ʷ���¼��", "��ʷ���¼��ˢ��"))
    End If
    '���뱨�������Ϣˢ��
    If intRefreshType = -1 Or intRefreshType = 1 Then
        mrsReport.Filter = ""
        If intRefreshType = -1 Then
            If mrsReport.RecordCount = 0 Then
                lblRptSel.Visible = False: lblRptTotal.Visible = False
                chkRpt.value = 0: chkRpt.Enabled = False
            End If
            '���뱨����ؿؼ�����������
            Call chkRpt_Click
        End If
        
        lblRptTotal.Caption = "������" & mrsReport.RecordCount & "�����嵼�룺"
        mrsReport.Filter = "��������=1"
        lblRptTotal.Caption = lblRptTotal.Caption & mrsReport.RecordCount & "��ֻ��������Դ��"
        mrsReport.Filter = "��������=2"
        lblRptTotal.Caption = lblRptTotal.Caption & mrsReport.RecordCount
        Call RecToLog(mrsReport, "ϵͳ���,���", IIf(intRefreshType = -1, "ԭʼ���뱨���¼��", "���뱨���¼��ˢ��"))
    End If
    '��ѡ���̻�����Ϣˢ��
    If intRefreshType = -1 Or intRefreshType = 2 Then
        mrsOptionalProc.Filter = ""
        If intRefreshType = -1 Then
            If mrsOptionalProc.RecordCount = 0 Then
                lblOptSel.Visible = False: lblOptTotal.Visible = False
                chkOpt.value = 0: chkOpt.Enabled = False
            End If
            Call chkOpt_Click
        End If
        lblOptTotal.Caption = "������" & mrsOptionalProc.RecordCount & "��ִ�У�"
        mrsOptionalProc.Filter = "ִ��=1"
        lblOptTotal.Caption = lblOptTotal.Caption & mrsOptionalProc.RecordCount
        Call RecToLog(mrsOptionalProc, "ϵͳ���,ID", IIf(intRefreshType = -1, "ԭʼ��ѡ����¼��", "��ѡ����¼��ˢ��"))
    End If
End Sub

Private Sub StepSwitch(ByVal intWay As Integer)
    Dim strPre As String, arrTmp As Variant
    Dim strOptProcs As String
    
    On Error GoTo errH
    If intWay = 1 Then
        If Not StepValidate(mintStep) Then Exit Sub
    End If
    If mintStep = SI_��Ǩ���� Then
        If MsgBox("ϵͳ��Ǩ�����¹��ش���ȷ���Ѿ������˸���׼��������" & vbCrLf & vbCrLf & "Ҫ��ʼ����ϵͳ��Ǩ��", _
                vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    mintStep = mintStep + intWay
    If mintStep = SI_ϵͳ��Ǩ Then
        'ɾ������Ҫ��������ʷ��ű�
        mrsHistorySpace.Filter = "����=0"
        Do While Not mrsHistorySpace.EOF
            Call RecDelete(mrsSysFiles, "ϵͳ���=" & mrsHistorySpace!ϵͳ��� & " And ������='" & UCase(mrsHistorySpace!���� & "") & "' And SysType=" & ST_History)
            mrsHistorySpace.MoveNext
        Loop
        'ɾ����׼��ʷ��ű���¼��
        Call RecDelete(mrsSysFiles, "������=Null And SysType=" & ST_History)
       '�����Ҫִ�еĿ�ѡ����
        If Not mblnExecBef Then
            mrsOptionalProc.Filter = "ִ��=1"
            mrsOptionalProc.Sort = "ϵͳ���,SPVer,ִ����,��ʷ��"
            Do While Not mrsOptionalProc.EOF
                If strPre <> mrsOptionalProc!ִ���� & "|" & mrsOptionalProc!SPVer & "|" & mrsOptionalProc!ϵͳ��� & "|" & mrsOptionalProc!��ʷ�� Then
                    If strPre <> "" Then
                        arrTmp = Split(strPre, "|")
                        Call RecUpdate(mrsSysFiles, "ϵͳ���=" & arrTmp(2) & " And SPVer='" & arrTmp(1) & "' And FileType=" & FT_Optional & IIf(arrTmp(3) = 1, " And SysType=" & ST_History & " And ������='" & arrTmp(0) & "'", " And SysType<>" & ST_History), "Optional", IIf(strOptProcs = "", Null, Mid(strOptProcs, 2)))
                    End If
                    strPre = mrsOptionalProc!ִ���� & "|" & mrsOptionalProc!SPVer & "|" & mrsOptionalProc!ϵͳ��� & "|" & mrsOptionalProc!��ʷ��
                    strOptProcs = ""
                End If
                strOptProcs = strOptProcs & "," & mrsOptionalProc!����
                mrsOptionalProc.MoveNext
            Loop
            If strPre <> "" Then
                arrTmp = Split(strPre, "|")
                Call RecUpdate(mrsSysFiles, "ϵͳ���=" & arrTmp(2) & " And SPVer='" & arrTmp(1) & "' And FileType=" & FT_Optional & IIf(arrTmp(3) = 1, " And SysType=" & ST_History & " And ������='" & arrTmp(0) & "'", " And SysType<>" & ST_History), "Optional", IIf(strOptProcs = "", Null, Mid(strOptProcs, 2)))
            End If
            'ɾ��û��ִ�еĿ�ѡ�ű�
            Call RecDelete(mrsSysFiles, "FileType=" & FT_Optional & " And Optional=Null")
        End If
    End If
    Call StepDisplay(mintStep)
    If mintStep = SI_ϵͳ��Ǩ Then
        '����ʵ���������ʹ�úۼ�
        Set mclsRunScript = New clsRunScript
        
        '���ò��������
        Call mclsRunScript.InitGlobalPara(Me, 0, optErrOption(ET_�������д���).value, _
                                                            lblLogModi.Tag, IIf(chkLogLong.value = 0, 0, val(txtLogLong.Text)), True, mblnExecBef And ckhIdxOnLine.value = 1, optLogType(1).value, True)
        '��ʼ���û�������Ϣ�����ܿ�����õ�
        Call mclsRunScript.InitUserList(gstrUserName, gstrPassword, txtToolsPwd.Text, txtDBAUser.Text, txtDBAPwd.Text)
        mclsRunScript.Server = gstrServer
        '��Ǩ��־��¼��Ǩ���ã��Լ���Ǩ����
        Call LogSetInfo
        Call UpgradeExecute
        On Error Resume Next
        Unload Me
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Private Sub LogSetInfo()
'���ܣ���¼��־��Ϣ
    Dim strLog As String, strTmp As String
    Dim lngLen As Long
    Dim vsTmp As VSFlexNode
    Dim i As Long
    
    On Error GoTo errH
    '��Ǩ��־��¼��Ǩ���ã��Լ���Ǩ����
    lngLen = 16
    mclsRunScript.WriteSection "��Ǩ��Ҫ��Ϣ"
    mclsRunScript.WriteLog "������ʱ�䣺" & Format(CurrentDate, "yyyy-MM-dd HH:mm:ss") & String(4, " ") & "������ʱ�䣺" & Format(Now, "yyyy-MM-dd HH:mm:ss")
    mclsRunScript.WriteLog "˵����Ϊ�˼��������ݿ�������Ľ��������½�ʹ�ñ���ʱ����Ϊ��¼��־��ʱ��"
    mrsSysInfo.Filter = "ϵͳ���=0" '������
    Call LogOracleSet
    mclsRunScript.WriteLog "��  λ  ��  �ƣ�" & gobjRegister.zlRegInfo("��λ����", False, 0)
    mclsRunScript.WriteLog "��    ��    ����" & gstrServer
    strTmp = IIf(mblnExecBef, mrsSysInfo!��ǰĿ��汾, mrsSysInfo!Ŀ��汾) & ""
    mclsRunScript.WriteLog "��  ��  ��  �ߣ�" & mrsSysInfo!ϵͳ�汾�� & IIf(strTmp <> "", "-->" & IIf(mblnExecBef, mrsSysInfo!��ǰĿ��汾, mrsSysInfo!Ŀ��汾), "")
    mrsSysInfo.Filter = "ϵͳ���<>0 and ����=1"
    mrsSysInfo.Sort = "Sort,ϵͳ���"
    Do While Not mrsSysInfo.EOF
        strTmp = IIf(mblnExecBef, mrsSysInfo!��ǰĿ��汾, mrsSysInfo!Ŀ��汾)
        mclsRunScript.WriteLog mrsSysInfo!ϵͳ��� & "-" & mrsSysInfo!ϵͳ���� & "��" & mrsSysInfo!ϵͳ�汾�� & IIf(strTmp <> "", "-->" & IIf(mblnExecBef, mrsSysInfo!��ǰĿ��汾, mrsSysInfo!Ŀ��汾), "")
        mrsSysInfo.MoveNext
    Loop
    mclsRunScript.WriteSection "��Ǩ����"
    '����������־
    mclsRunScript.WriteLog "��Ǩ����"
    If chkParallel.value = 0 Or chkParallel.Enabled = False Then
        mintDDLParallel = 0
        mclsRunScript.WriteLog "  �����ò���DDL"
    Else
        mintDDLParallel = val(txtCpu.Text)
        mclsRunScript.WriteLog "  ���ò���DDL ���ж�=" & val(txtCpu.Text)
    End If
    If Not ckhIdxOnLine.Visible Or ckhIdxOnLine.value = 0 Then
        mclsRunScript.WriteLog "  ����������ģʽ��������"
    Else
        mclsRunScript.WriteLog "  ��������ģʽ��������"
    End If
    mclsRunScript.WriteLog "  ��־��¼��ʽ��ȡ" & IIf(optLogType(1).value, "ֻ��¼δ���ԵĴ�����־", "��¼���д�����־")
    If chkLogLong.value = 0 Then
        mclsRunScript.WriteLog "  ��־����¼��ʱִ��SQL"
    Else
        mclsRunScript.WriteLog "  ��־��¼ִ�г���" & val(txtLogLong.Text) & "���ӵ�SQL���"
    End If
    mclsRunScript.WriteLog "  ������ʽ��ȡ" & IIf(optErrOption(ET_���Դ�Ҫ����).value, "���Դ�Ҫ����", "�������д���")
    '��ʷ��ѡ����־
    mrsHistorySpace.Filter = ""
    mrsHistorySpace.Sort = "ϵͳ���,��ǰ,���"
    If mrsHistorySpace.RecordCount <> 0 Then
        mclsRunScript.WriteLog String(80, "-")
        mclsRunScript.WriteLog "��ʷ�ռ���Ǩ"
        Do While Not mrsHistorySpace.EOF
            strLog = "    " & Lpad(mrsHistorySpace!ϵͳ���, 4) & "-" & RPAD(mrsHistorySpace!ϵͳ����, 16)
            strLog = strLog & "  " & RPAD(mrsHistorySpace!����, 14) & "  " & RPAD(IIf(mrsHistorySpace!��ǰ = 1, "��ǰ", "�ǵ�ǰ"), 5)
            strLog = strLog & "  " & IIf(mrsHistorySpace!���� = 1, "����", "������")
            mclsRunScript.WriteLog strLog
            mrsHistorySpace.MoveNext
        Loop
    End If
    '��ѡ������־
    mrsOptionalProc.Filter = ""
    mrsOptionalProc.Sort = "ϵͳ���,��ʷ��,ID"
    If mrsOptionalProc.RecordCount <> 0 Then
        mclsRunScript.WriteLog String(80, "-")
        mclsRunScript.WriteLog "ִ�п�ѡ����"
        Do While Not mrsOptionalProc.EOF
            strLog = "    " & Lpad(mrsOptionalProc!ϵͳ���, 4) & "-" & RPAD(mrsOptionalProc!ϵͳ����, 16)
            strLog = strLog & "  " & RPAD(mrsOptionalProc!����, 32) & "  " & RPAD(mrsOptionalProc!ִ����, lngLen - 2)
            strLog = strLog & "  " & RPAD(IIf(mrsOptionalProc!��ʷ�� = 1, "��ʷ��", "����ʷ��"), 6) & "  " & RPAD(IIf(mrsOptionalProc!ִ�� = 1, "ִ��", "��ִ��"), 6)
            strLog = strLog & "  " & mrsOptionalProc!FilePath
            mclsRunScript.WriteLog strLog
            mrsOptionalProc.MoveNext
        Loop
    End If
    '���뱨����־
    mrsReport.Filter = ""
    mrsReport.Sort = "ϵͳ���,ID"
    If mrsReport.RecordCount <> 0 Then
        mclsRunScript.WriteLog String(80, "-")
        mclsRunScript.WriteLog "���뱨��"
        Do While Not mrsReport.EOF
            strLog = "    " & Lpad(mrsReport!ϵͳ���, 4) & "-" & RPAD(mrsReport!ϵͳ����, lngLen)
            strLog = strLog & "  " & RPAD(mrsReport!���, 20) & "  " & RPAD(mrsReport!����, 30)
            strLog = strLog & "  " & RPAD(Decode(mrsReport!��������, 0, "������", 1, "���嵼��", 2, "����Դ����"), 10)
            strLog = strLog & "  " & mrsReport!FilePath
            mclsRunScript.WriteLog strLog
            mrsReport.MoveNext
        Loop
    End If
    mclsRunScript.WriteSection "��Ǩ����"
    For i = vsPlan.FixedRows + 1 To vsPlan.Rows - IIf(mblnExecBef, 1, 2)
        Set vsTmp = vsPlan.GetNode(i)
        mclsRunScript.WriteLog vsTmp.Text
        vsTmp.Expanded = False
    Next
    
    mclsRunScript.WriteLog
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Private Function StepValidate(ByVal intStep As IDX_STEP) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim cnTmp As New ADODB.Connection
    Dim strMsg As String
    Dim strErr As String
    Dim blnDo As Boolean
    
    On Error GoTo errH
    If intStep = SI_��Ǩ���� Then
        If txtToolsPwd.Enabled And txtToolsPwd.Text = "" Then
            MsgBox "������������û������롣", vbInformation, gstrSysName
            txtToolsPwd.SetFocus: Exit Function
        End If
        If txtDBAUser.Enabled And txtDBAUser.Text = "" Then
            MsgBox "���������DBA��ݵ��û�����", vbInformation, gstrSysName
            txtDBAUser.SetFocus: Exit Function
        End If
        If txtDBAPwd.Enabled And txtDBAPwd.Text = "" Then
            MsgBox "������DBA�û������롣", vbInformation, gstrSysName
            txtDBAPwd.SetFocus: Exit Function
        End If
        If txtToolsPwd.Enabled Then
            '������������֤
            If gstrToolsPwd <> "" And UCase(gstrToolsPwd) <> UCase(Trim(txtToolsPwd.Text)) Then
                 MsgBox "�������������", vbInformation, gstrSysName
                 Exit Function
            End If
            err.Clear
            
            If gcnTools Is Nothing Then
                blnDo = True
            ElseIf gcnTools.State = adStateClosed Then
                blnDo = True
            End If
            
            If blnDo Then
                Set gcnTools = gobjRegister.GetConnection(gstrServer, "zltools", txtToolsPwd.Text, False, MSODBC, "", False)
                If gcnTools.State = adStateClosed Then
                    MsgBox "���ӹ������û�ʱ���ִ���" & vbCrLf & vbCrLf & strErr, vbInformation, gstrSysName
                    Exit Function
                End If
                Call SetSQLTrace(gstrServer, "zltools", gcnTools)
                gstrToolsPwd = txtToolsPwd.Text '��ֵ
            End If
        End If
        If txtDBAPwd.Enabled Then
            'DBA�û�������֤
            If UCase(txtDBAUser.Text) = UCase(gstrSysUser) And UCase(gstrSysPwd) <> UCase(txtDBAPwd.Text) And gstrSysPwd <> "" Then
                MsgBox "DBA�û��������", vbInformation, gstrSysName
                Exit Function
            End If
            If gcnSystem Is Nothing Then
                blnDo = True
            ElseIf gcnSystem.State = adStateClosed Then
                blnDo = True
            End If
            
            If blnDo Then
                Set cnTmp = gobjRegister.GetConnection(gstrServer, txtDBAUser.Text, txtDBAPwd.Text, False, MSODBC, "", False)
                If cnTmp.State = adStateClosed Then
                    MsgBox "����DBA�û�ʱ���ִ���.", vbInformation, gstrSysName
                    Exit Function
                End If
                On Error GoTo 0
                '����Ƿ�DBA
                If CheckIsDBA(cnTmp) = False Then
                    MsgBox "���û�������DBA��ݣ�", vbExclamation, gstrSysName
                    txtDBAUser.SetFocus: Exit Function
                End If
                
                Call SetSQLTrace(gstrServer, txtDBAUser.Text, cnTmp)
                Set gcnSystem = cnTmp
                gstrSysUser = txtDBAUser.Text
                gstrSysPwd = txtDBAPwd.Text
            Else
                Call SetSQLTrace(gstrServer, gstrSysUser, gcnSystem)
            End If
        End If
        '����������Ǩ��־
        If lblLog.Caption = "��Ǩ��־�ļ���" Then
            MsgBox "��ȷ����Ǩ��־�ļ��Ĵ��λ�ú����֡�", vbInformation, gstrSysName
            Exit Function
        End If
        '��ǰ��ʷ�����������û��ע�����ʷ���򲻼�飬���û����֤���������֤���룬û��ѡ����������ʷ��
        Call RecUpdate(mrsHistorySpace, "��ǰ=1 And ����=0  And ��֤<>1", "����", 1)
        Call RecUpdate(mrsHistorySpace, "��֤=1", "����", 0)
        mrsHistorySpace.Filter = "��ǰ=1 And ��֤=0 And ����=1"
        mrsHistorySpace.Sort = "ϵͳ���,ID": strMsg = ""
        Do While Not mrsHistorySpace.EOF
            strMsg = strMsg & vbNewLine & "��" & mrsHistorySpace!ϵͳ���� & "���ı�ռ�-" & mrsHistorySpace!����
            mrsHistorySpace.MoveNext
        Loop
        If strMsg <> "" Then
            MsgBox "����ϵͳ�ĵ�ǰ��ʷ��ռ����������" & strMsg & "���������֤��", vbInformation, gstrSysName
            '���¶�ȡ��ѡ�ű�
            Call ReadOptionalProc(True)
            'ˢ����ʷ�������Ϣ
            Call RefreshTotalInfo(0)
            Exit Function
        End If
        mrsHistorySpace.Filter = "����=1 And ��֤=0"
        mrsHistorySpace.Sort = "ϵͳ���,ID": strMsg = ""
        Do While Not mrsHistorySpace.EOF
            strMsg = strMsg & vbNewLine & "��" & mrsHistorySpace!ϵͳ���� & "���ı�ռ�-" & mrsHistorySpace!���� & "��"
            mrsHistorySpace.MoveNext
        Loop
        If strMsg <> "" Then
            If MsgBox("������ʷ��ռ�δͨ����֤����������" & strMsg & vbNewLine & "�Ƿ������", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                '���¶�ȡ��ѡ�ű�
                Call ReadOptionalProc(True)
                'ˢ����ʷ�������Ϣ
                Call RefreshTotalInfo(0)
                Exit Function
            End If
            '��û��ͨ����֤����ʷ��ȡ������
            Call RecUpdate(mrsHistorySpace, "����=1 And ��֤<>2 ", "����", 0)
        End If
        '��ͨ����֤�Ҳ�����������ʷ��ȡ������
        Call RecUpdate(mrsHistorySpace, "����=1 And ��֤=2 " & IIf(mblnExecBef, "  And ����ǰ����=0", " And ������=0"), "����", 0)
    End If
    StepValidate = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub StepDisplay(ByVal intStep As IDX_STEP)
    Dim i As Integer
    Dim arrTmp As Variant
    Dim strTmp As String, strMaxVer As String
    Dim vsnRoot As VSFlexNode, vsnTop As VSFlexNode, vsnSecd As VSFlexNode
    Dim vsnAPP As VSFlexNode, vsnHis As VSFlexNode, vsnRpt As VSFlexNode, vsnCompile As VSFlexNode
    Dim vsnCHCEK As VSFlexNode, vsnTools As VSFlexNode, vsnAdjustSeq As VSFlexNode
    
    mblnHaveST = False
    arrTmp = Split(Split(STEP_INFO, "||")(intStep), "|")
    For i = 0 To fraStep.UBound
        fraStep(i).Visible = i = intStep
    Next
    cmdCancel.Enabled = intStep < SI_ϵͳ��Ǩ
    If intStep = SI_ϵͳ��Ǩ Then
        Call SetSQLState(True, True)
        With vsPlan
            'ע�⣺�ؼ��ָ���������^�ָ���»��ߣ���Ҫ�����ڣ���ʷ���Լ��û��ȣ����������»���
            .Rows = .FixedRows: .Rows = .FixedRows + 1: .IsSubtotal(.Rows - 1) = True
            '���һ�����ڵ㣬��������ӽڵ�
            Set vsnRoot = .GetNode(.Rows - 1): vsnRoot.Text = "ϵͳ��Ǩ": vsnRoot.key = "Main": Set vsnRoot.Image = imgPlan.ListImages("Doing").Picture: vsnRoot.Expanded = True
             .Rows = .Rows + 1: .IsSubtotal(.Rows - 1) = True
            If Not mblnExecBef Then
                Set vsnTop = .GetNode(.Rows - 1): vsnTop.Text = "�ͻ���վ�㲿������": vsnTop.key = "Client": Set vsnTop.Image = imgPlan.ListImages("Wait").Picture: vsnTop.Expanded = True
                Set vsnCHCEK = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "��Ǩ���", FS_��Ǩ���, imgPlan.ListImages("Wait").Picture)
            End If
            If txtToolsPwd.Enabled Then

                mrsSysFiles.Filter = "ϵͳ���=0": mrsSysFiles.Sort = "FullSPVer"
                If Not mrsSysFiles.EOF Then
                    Set vsnTools = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "������" & IIf(mblnExecBef, "��ǰ", "") & "��Ǩ", FS_������Ǩ, imgPlan.ListImages("Wait").Picture)
                    If Not mblnExecBef Then Call vsnCHCEK.AddNode(flexNTLastChild, GetCode(vsnCHCEK.Text) & "." & (vsnCHCEK.Children + 1) & "������", vsnCHCEK.key & "^TOOLS", imgPlan.ListImages("Wait").Picture)
                    'PLJSON��װ����,��ǰ����û�и�����
                    If mblnInstallPLJson And Not mblnExecBef Then
                        Call vsnTools.AddNode(flexNTLastChild, "PLJSON��װ", vsnTools.key & "^PLJSON", imgPlan.ListImages("Wait").Picture)
                    End If
                ElseIf Not mblnExecBef Then
                    Set vsnTools = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "��������Ǩ", FS_������Ǩ, imgPlan.ListImages("Wait").Picture)
                    'PLJSON��װ����,��ǰ����û�и�����
                    If mblnInstallPLJson Then
                        Call vsnTools.AddNode(flexNTLastChild, "PLJSON��װ", vsnTools.key & "^PLJSON", imgPlan.ListImages("Wait").Picture)
                    End If
                End If

                strTmp = ""
                Do While Not mrsSysFiles.EOF
                    If InStr(strTmp & ",", "," & mrsSysFiles!SPVer & ",") = 0 Then
                        strTmp = strTmp & "," & mrsSysFiles!SPVer
                        '��ӹ�������Ǩ��ĳһ���汾
                        Call vsnTools.AddNode(flexNTLastChild, mrsSysFiles!SPVer, vsnTools.key & "^" & mrsSysFiles!SPVer, imgPlan.ListImages("Wait").Picture)
                    End If
                    mrsSysFiles.MoveNext
                Loop
               
                If Not mblnExecBef Then
                    Call vsnTools.AddNode(flexNTLastChild, "�޸�ͨ�������û�", vsnTools.key & "^ZLUA", imgPlan.ListImages("Wait").Picture)
                    Call vsnTools.AddNode(flexNTLastChild, "������Ȩ����", vsnTools.key & "^PUBGRANT", imgPlan.ListImages("Wait").Picture)
                End If
            End If
            'ϵͳ��Ǩ����
            mrsSysInfo.Filter = "ϵͳ���<>0 And ����=1": mrsSysInfo.Sort = "Sort"
            If Not mrsSysInfo.EOF Then
                Set vsnAPP = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "Ӧ��ϵͳ" & IIf(mblnExecBef, "��ǰ", "") & "��Ǩ", FS_Ӧ��ϵͳ��Ǩ, imgPlan.ListImages("Wait").Picture)
                mrsHistorySpace.Filter = IIf(mblnExecBef, "����=1", "")
                If mblnExecBef And Not mrsHistorySpace.EOF Or Not mblnExecBef Then
                    Set vsnHis = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "��ʷ��ռ�" & IIf(mblnExecBef, "��ǰ", "") & "��Ǩ", FS_��ʷ����Ǩ, imgPlan.ListImages("Wait").Picture)
                End If
                Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & ".����ͬ��ʴ���", FS_����ͬ���, imgPlan.ListImages("Wait").Picture)
                If Not mblnExecBef Then
                    Set vsnCompile = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "������Ч����", FS_������Ч����, imgPlan.ListImages("Wait").Picture)
                    Set vsnAdjustSeq = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "���µ�������", FS_��������, imgPlan.ListImages("Wait").Picture)
                    mrsReport.Filter = "��������<>0"
                    If Not mrsReport.EOF Then
                        Set vsnRpt = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "����������", FS_��������, imgPlan.ListImages("Wait").Picture)
                    End If
                    If mblnFinal Then
                        Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "����ֵ�ظ�������", FS_����ֵ��, imgPlan.ListImages("Wait").Picture)
                        Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "��ɫ������Ȩ", FS_��ɫ��Ȩ, imgPlan.ListImages("Wait").Picture)
                        Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "�Զ���ҵ���������", FS_��̨�Զ�ҵ����, imgPlan.ListImages("Wait").Picture)
                        Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "ִ���ӳٽű�(��̨)", FS_�ӳٽű�, imgPlan.ListImages("Wait").Picture)
                    End If
                End If
            ElseIf Not mblnExecBef Then
                Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & ".����ͬ��ʴ���", FS_����ͬ���, imgPlan.ListImages("Wait").Picture)
                Set vsnCompile = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "������Ч����", FS_������Ч����, imgPlan.ListImages("Wait").Picture)
                Set vsnAdjustSeq = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "���µ�������", FS_��������, imgPlan.ListImages("Wait").Picture)
                If mblnFinal Then
                    Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "����ֵ�ظ�������", FS_����ֵ��, imgPlan.ListImages("Wait").Picture)
                    Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "��ɫ������Ȩ", FS_��ɫ��Ȩ, imgPlan.ListImages("Wait").Picture)
                    Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "�Զ���ҵ���������", FS_��̨�Զ�ҵ����, imgPlan.ListImages("Wait").Picture)
                    Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "ִ���ӳٽű�(��̨)", FS_�ӳٽű�, imgPlan.ListImages("Wait").Picture)
                End If
            Else
                Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & ".����ͬ��ʴ���", FS_����ͬ���, imgPlan.ListImages("Wait").Picture)
                Set vsnCompile = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "������Ч����", FS_������Ч����, imgPlan.ListImages("Wait").Picture)
            End If
            
            'û��������׼�浫����Ҫ�ؽ����ܺ���
            If Not mblnExecBef And Not vsnAPP Is Nothing Then
                mrsSysInfo.Filter = "ϵͳ���=100 And ����=0"
                If Not mrsSysInfo.EOF Then
                    Set vsnTop = vsnAPP.AddNode(flexNTLastChild, GetCode(vsnAPP.Text) & "." & (vsnAPP.Children + 1) & "." & mrsSysInfo!ϵͳ����, vsnAPP.key & "^" & mrsSysInfo!ϵͳ���, imgPlan.ListImages("Wait").Picture)
                    Call vsnTop.AddNode(flexNTLastChild, "ע����Ȩ�������������", vsnTop.key & "^ZLREGISTER", imgPlan.ListImages("Wait").Picture)
                    Call vsnTop.AddNode(flexNTLastChild, "H�����Ȩ������", vsnTop.key & "^HTABLEREPAIR", imgPlan.ListImages("Wait").Picture)
                End If
                mrsSysInfo.Filter = "ϵͳ���<>0 And ����=1": mrsSysInfo.Sort = "Sort"
            Else
                mrsSysInfo.Filter = "ϵͳ���<>0 And ����=1": mrsSysInfo.Sort = "Sort"
            End If
            
            
            mstrSysCodes = ""
            Do While Not mrsSysInfo.EOF
                If mrsSysInfo!ϵͳ��� \ 100 = 1 Then mblnHaveST = True
                mstrSysCodes = mstrSysCodes & IIf(mstrSysCodes = "", "", ",") & mrsSysInfo!ϵͳ���
                
                '��Ǩ�����������
                 If Not mblnExecBef Then Call vsnCHCEK.AddNode(flexNTLastChild, GetCode(vsnCHCEK.Text) & "." & (vsnCHCEK.Children + 1) & mrsSysInfo!ϵͳ����, vsnCHCEK.key & "^" & mrsSysInfo!ϵͳ���, imgPlan.ListImages("Wait").Picture)
                
                'Ӧ��ϵͳ��Ǩ��������
                mrsSysFiles.Filter = "ϵͳ���=" & mrsSysInfo!ϵͳ��� & " And SysType<>" & ST_History: mrsSysFiles.Sort = "FullSPVer"
                Set vsnTop = vsnAPP.AddNode(flexNTLastChild, GetCode(vsnAPP.Text) & "." & (vsnAPP.Children + 1) & "." & mrsSysInfo!ϵͳ����, vsnAPP.key & "^" & mrsSysInfo!ϵͳ���, imgPlan.ListImages("Wait").Picture)
                strTmp = ""
                Do While Not mrsSysFiles.EOF
                    If InStr(strTmp & ",", "," & mrsSysFiles!SPVer & ",") = 0 Then
                        strTmp = strTmp & "," & mrsSysFiles!SPVer
                        Call vsnTop.AddNode(flexNTLastChild, mrsSysFiles!SPVer, vsnTop.key & "^" & mrsSysFiles!SPVer, imgPlan.ListImages("Wait").Picture)
                    End If
                    mrsSysFiles.MoveNext
                Loop
                
                '���ܺ����ؽ���H�����Ȩ����������
                If mrsSysInfo!ϵͳ��� \ 100 = 1 And Not mblnExecBef Then
                    Call vsnTop.AddNode(flexNTLastChild, "ע����Ȩ�������������", vsnTop.key & "^ZLREGISTER", imgPlan.ListImages("Wait").Picture)
                End If
                If Not mblnExecBef Then
                    Call vsnTop.AddNode(flexNTLastChild, "H�����Ȩ������", vsnTop.key & "^HTABLEREPAIR", imgPlan.ListImages("Wait").Picture)
                End If
                 '��ʷ����Ǩ�������ӣ����Ӳ���������ʷ��չʾ
                If Not vsnHis Is Nothing Then
                    mrsHistorySpace.Filter = "(����=1 And ϵͳ���=" & mrsSysInfo!ϵͳ��� & ") OR (����=0 And ��ǰ=1 And ϵͳ���=" & mrsSysInfo!ϵͳ��� & ")": mrsHistorySpace.Sort = "��ǰ Desc,���"
                    If Not mrsHistorySpace.EOF Then
                        '�����ʷ������ϵͳ
                        Set vsnTop = vsnHis.AddNode(flexNTLastChild, GetCode(vsnHis.Text) & "." & (vsnHis.Children + 1) & "." & mrsSysInfo!ϵͳ����, vsnHis.key & "^" & mrsSysInfo!ϵͳ���, imgPlan.ListImages("Wait").Picture)
                        Do While Not mrsHistorySpace.EOF
                            '���ĳ��ϵͳ��ʷ��
                            Set vsnSecd = vsnTop.AddNode(flexNTLastChild, mrsHistorySpace!����, vsnTop.key & "^" & mrsHistorySpace!����, imgPlan.ListImages("Wait").Picture)
                            mrsSysFiles.Filter = "������='" & UCase(mrsHistorySpace!���� & "") & "' And ϵͳ���=" & mrsSysInfo!ϵͳ��� & " And SysType=" & ST_History: mrsSysFiles.Sort = "FullSPVer"
                            strTmp = "": strMaxVer = ""
                            '���ĳ��ϵͳ��ʷ����Ǩ����
                            Do While Not mrsSysFiles.EOF
                                If InStr(strTmp & ",", "," & mrsSysFiles!SPVer & ",") = 0 Then
                                    strTmp = strTmp & "," & mrsSysFiles!SPVer
                                    Call vsnSecd.AddNode(flexNTLastChild, mrsSysFiles!SPVer, vsnSecd.key & "^" & mrsSysFiles!SPVer, imgPlan.ListImages("Wait").Picture)
                                    strMaxVer = mrsSysFiles!SPVer & ""
                                End If
                                mrsSysFiles.MoveNext
                            Loop
                            If strMaxVer = "" Then strMaxVer = mrsHistorySpace!��ǰ�汾
                            '����ǰִ�У�����ű���֧�ֵ�Ŀ��汾�������Զ�����Ŀ��汾
                            If VerFull(strMaxVer) < VerFull(mrsHistorySpace!Ŀ��汾) And Not mblnExecBef Then
                                Call vsnSecd.AddNode(flexNTLastChild, mrsHistorySpace!Ŀ��汾, vsnSecd.key & "^" & mrsHistorySpace!Ŀ��汾, imgPlan.ListImages("Wait").Picture)
                            End If
                            
                            If Not mblnExecBef Then
                                If VerFull(strMaxVer) > VerFull(mrsHistorySpace!Ŀ��汾) Then
                                    '����ǰ��ʷ��汾����Ŀ��汾����ʲôҲ����
                                    Call vsnSecd.AddNode(flexNTLastChild, "�߰汾��ʷ����", vsnSecd.key & "^DONOTHING", imgPlan.ListImages("Wait").Picture)
                                Else
                                    Call vsnSecd.AddNode(flexNTLastChild, "��ʷ��ṹ����", vsnSecd.key & "^HISREPAIR", imgPlan.ListImages("Wait").Picture)
                                End If
                            End If
                            mrsHistorySpace.MoveNext
                        Loop
                    ElseIf Not mblnExecBef Then   'û����ʷ�⣬����Ҫ��֤
                        Set vsnTop = vsnHis.AddNode(flexNTLastChild, GetCode(vsnHis.Text) & "." & (vsnHis.Children + 1) & "." & mrsSysInfo!ϵͳ���� & "��ʷ����", vsnHis.key & "^" & mrsSysInfo!ϵͳ���, imgPlan.ListImages("Wait").Picture)
                    End If
                End If
                '���������
                If Not vsnRpt Is Nothing Then
                    mrsReport.Filter = "��������<>0 And ϵͳ���=" & mrsSysInfo!ϵͳ���
                    If Not mrsReport.EOF Then
                        Call vsnRpt.AddNode(flexNTLastChild, GetCode(vsnRpt.Text) & "." & (vsnRpt.Children + 1) & "." & mrsSysInfo!ϵͳ����, vsnRpt.key & "^" & mrsSysInfo!ϵͳ���, imgPlan.ListImages("Wait").Picture)
                    End If
                End If
                mrsSysInfo.MoveNext
            Loop
            
            
            '������Ч������������
            If Not vsnCompile Is Nothing Then
                If Not vsnTools Is Nothing Then
                    Call vsnCompile.AddNode(flexNTLastChild, GetCode(vsnCompile.Text) & "." & (vsnCompile.Children + 1) & ".������", vsnCompile.key & "^TOOLS", imgPlan.ListImages("Wait").Picture)
                End If
                If Not vsnAPP Is Nothing Then
                    Call vsnCompile.AddNode(flexNTLastChild, GetCode(vsnCompile.Text) & "." & (vsnCompile.Children + 1) & ".Ӧ��ϵͳ", vsnCompile.key & "^APP", imgPlan.ListImages("Wait").Picture)
                End If
            End If
            '����������������
            If Not vsnAdjustSeq Is Nothing Then
                If Not vsnTools Is Nothing Then
                    Call vsnAdjustSeq.AddNode(flexNTLastChild, GetCode(vsnAdjustSeq.Text) & "." & (vsnAdjustSeq.Children + 1) & ".������", vsnAdjustSeq.key & "^TOOLS", imgPlan.ListImages("Wait").Picture)
                End If
                If Not vsnAPP Is Nothing Then
                    Call vsnAdjustSeq.AddNode(flexNTLastChild, GetCode(vsnAdjustSeq.Text) & "." & (vsnAdjustSeq.Children + 1) & ".Ӧ��ϵͳ", vsnAdjustSeq.key & "^APP", imgPlan.ListImages("Wait").Picture)
                End If
            End If
        End With
        txtSQL.SetFocus: Me.Refresh
    End If
    Set imgInfo.Picture = imgStep.ListImages(intStep + 1).Picture
    lblStep.Caption = arrTmp(0)
    lblInfo.Caption = arrTmp(1)
    cmdNext.Enabled = intStep + 1 <= fraStep.UBound
    cmdNext.Visible = cmdNext.Enabled
End Sub

Private Sub UpgradeExecute()
'���ܣ������򵼵����ã�����ϵͳ��Ǩ
    Dim vsnStep As VSFlexNode
    Dim cnTmp As ADODB.Connection, cnCurrent As ADODB.Connection
    Dim arrTmp As Variant
    Dim strMsg As String, strPreVer As String, strError As String
    Dim i As Long, lngSec As Long, lngCount As Long
    Dim blnFirstUp As Boolean
    Dim datStart As Date, datSysStart As Date
    
    tmrRefresh.Enabled = True
    On Error GoTo errH
    mstrChangeTables = ""
    Call UpdateSysFiles '��¼������Ǩϵͳ�������ļ�
    mdatStart = Now
    blnFirstUp = True
    For i = vsPlan.FixedRows To vsPlan.Rows - 2
        Set cnCurrent = Nothing
        Call vsPlan.ShowCell(i, 0)
        Set vsnStep = vsPlan.GetNode(i)
        If vsnStep.Children = 0 Then  '����ִ�еĲ���
            arrTmp = Split(vsnStep.key, "^")
            If UBound(arrTmp) = 0 Then
                Call SetSQLState(False) '�ر�SQL
                mclsRunScript.WriteSection vsnStep.Text, IIf(i = vsPlan.FixedRows, "=", "-")
            Else
                mclsRunScript.WriteLog "[" & vsnStep.Text & "]"
            End If
            datStart = Now
            Call SetStepStateImg(vsnStep)  '��ʼִ��
            Select Case arrTmp(0)
                Case FS_��Ǩ���
                    If Not UpgradeCheck(val(arrTmp(1))) Then GoTo AbortLine
                Case FS_������Ǩ
                    If arrTmp(1) = "PUBGRANT" Then
                        If Not mblnExecBef Then '����ǰִ���޸�Ϊ0 ,������ǰִ���Ѿ��������м�״̬1
                            Set cnCurrent = gcnOracle
                            gcnOracle.Execute "Update zlUpGrade Set ��ǰִ��=0 Where ��ǰִ�� = 1 And ϵͳ is Null "
                            Set cnCurrent = Nothing
                        End If
                        mclsRunScript.SysNo = 0
                        Call ReGrantForTools(gcnTools, , True)
                        Call mclsRunScript.WriteCSVRow(0, "", "", "", Round((DateDiff("s", datSysStart, Now)) / 60))
                    ElseIf arrTmp(1) = "PLJSON" Then
                        Call InstallPLJSON(gcnSystem, mstrToolsFloder, mclsRunScript, mblnJSONRemain)
                        If blnFirstUp Then
                            mrsSysInfo.Filter = "ϵͳ���=0": mclsRunScript.SysNo = 0: strPreVer = ""
                            datSysStart = Now
                        End If
                    ElseIf arrTmp(1) = "ZLUA" Then
                        If Not RepairGeneralAccount(gcnOracle, "ZLUA", , strError) Then
                            mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "��ʧ��:" & strError
                        Else
                            mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "���ɹ�"
                        End If
                        '������LOB����
                        If (mintToolLob And LC_ISLONGRAW) = LC_ISLONGRAW Then        '��ȻΪLong Raw
                            If (mintToolLob And (LC_ZLTOOLS_IS3590_OR_GREATER Or LC_ZLHIS_IS3590_OR_GREATER)) = (LC_ZLTOOLS_IS3590_OR_GREATER Or LC_ZLHIS_IS3590_OR_GREATER) Then     '���������׼�涼����Ҫ��
                                If (mintToolLob And LC_ZLTOOLS_CURIS3590_OR_GREATER) <> LC_ZLTOOLS_CURIS3590_OR_GREATER Then
                                    Call AdjustToolLob
                                End If
                            End If
                        End If
                        If blnFirstUp Then
                            mrsSysInfo.Filter = "ϵͳ���=0": mclsRunScript.SysNo = 0: strPreVer = ""
                            datSysStart = Now
                        End If
                    Else
                        If blnFirstUp Then
                            mrsSysInfo.Filter = "ϵͳ���=0": mclsRunScript.SysNo = 0: strPreVer = ""
                            datSysStart = Now
                        End If
                        Call SetSQLState(True, True)
                        If Not RunScriptByVersion(0, arrTmp(1), blnFirstUp, IIf(strPreVer = "", mrsSysInfo!ϵͳ�汾��, strPreVer), IIf(mblnExecBef, mrsSysInfo!��ǰĿ��汾, mrsSysInfo!Ŀ��汾)) Then Exit Sub
                        strPreVer = arrTmp(1)
                    End If
                Case FS_Ӧ��ϵͳ��Ǩ
                    If arrTmp(2) = "HTABLEREPAIR" Then
                        If Not mblnExecBef Then '����ǰִ���޸�Ϊ0 ,������ǰִ���Ѿ��������м�״̬1
                            Set cnCurrent = gcnOracle
                            gcnOracle.Execute "Update zlUpGrade Set ��ǰִ��=0 Where ��ǰִ�� = 1 And ϵͳ =" & val(arrTmp(1))
                            Set cnCurrent = Nothing
                        End If
                        mclsRunScript.SysNo = val(arrTmp(1))
                        Call HTablePrivsRepair(val(arrTmp(1)))
                        Call mclsRunScript.WriteCSVRow(val(arrTmp(1)), "", "", "", Round((DateDiff("s", datSysStart, Now)) / 60))
                    ElseIf arrTmp(2) = "ZLREGISTER" Then
                        Call RebuildRegistFile(gcnTools, mstrToolsFloder)
                    Else
                        If blnFirstUp Then
                            mrsSysInfo.Filter = "ϵͳ���=" & arrTmp(1): mclsRunScript.SysNo = val(arrTmp(1)): strPreVer = ""
                            datSysStart = Now
                        End If
                        Call SetSQLState(True, True)
                        If Not RunScriptByVersion(val(arrTmp(1)), arrTmp(2), blnFirstUp, IIf(strPreVer = "", mrsSysInfo!ϵͳ�汾��, strPreVer), IIf(mblnExecBef, mrsSysInfo!��ǰĿ��汾, mrsSysInfo!Ŀ��汾)) Then Exit Sub
                        strPreVer = arrTmp(1)
                    End If
                Case FS_��ʷ����Ǩ
                    If UBound(arrTmp) = 3 Then '��ʷ����Ǩ����
                        If arrTmp(3) = "DONOTHING" Then
                            'Do Nothing
                        Else
                            If blnFirstUp Then
                                mrsHistorySpace.Filter = "ϵͳ���=" & arrTmp(1) & " And ����='" & arrTmp(2) & "'"
                                mclsRunScript.SysNo = val(arrTmp(1))
                                mclsRunScript.HistoryDB = mrsHistorySpace!���� & IIf(mrsHistorySpace!DB���� & "" = "", "", "(DBLINK:" & mrsHistorySpace!DB���� & ")")
                                Set cnTmp = gobjRegister.GetConnection(mrsHistorySpace!������, mrsHistorySpace!������, mrsHistorySpace!����, False, MSODBC, "", False)
                                If Not cnTmp Is Nothing Then
                                    If cnTmp.State = adStateClosed Then
                                       Set cnTmp = Nothing
                                    Else
                                       Call SetSQLTrace(mrsHistorySpace!������, mrsHistorySpace!������, cnTmp)
                                    End If
                                End If
                                strPreVer = ""
                                datSysStart = Now
                                If mrsHisAfterSPace Is Nothing Then Set mrsHisAfterSPace = CopyNewRec(mrsHistorySpace, True)
                                Call RecDataAppend(mrsHisAfterSPace, mrsHistorySpace, 1, , , True)
                            End If
                            If Not cnTmp Is Nothing Then
                                If arrTmp(3) = "HISREPAIR" Then
                                    If Not mblnExecBef Then '����ǰִ���޸�Ϊ0 ,������ǰִ���Ѿ��������м�״̬1
                                        Set cnCurrent = cnTmp
                                        cnTmp.Execute "Update zlbakinfo Set ��ֹ���=NULL,��ǰ��ֹ���=NULL,��ǰִ��=0  Where ϵͳ=" & val(arrTmp(1))
                                        Set cnCurrent = Nothing
                                    End If
                                    Call RepairHisDB(cnTmp, val(arrTmp(1)), mrsHistorySpace!������, mrsHistorySpace!������, mrsHistorySpace!����, mrsHistorySpace!DB���� & "", mrsHistorySpace!��ǰ = 1)
                                    Call mclsRunScript.WriteCSVRow(val(arrTmp(1)), "", mclsRunScript.HistoryDB, "", Round((DateDiff("s", datSysStart, Now)) / 60))
                                    mclsRunScript.HistoryDB = ""
                                Else
                                    Call RunScriptByVersion(val(arrTmp(1)), arrTmp(3), blnFirstUp, IIf(strPreVer = "", mrsHistorySpace!��ǰ�汾, strPreVer), IIf(mblnExecBef, mrsHistorySpace!��ǰĿ��汾, mrsHistorySpace!Ŀ��汾), True, cnTmp, arrTmp(2))
                                    strPreVer = arrTmp(1)
                                End If
                            End If
                        End If
                    ElseIf UBound(arrTmp) = 1 Then 'û����ʷ��
                        lngCount = 0
                        If CheckHavHistory(val(arrTmp(1))) Then
ReDo:
                            lngCount = lngCount + 1
                            MsgBox "���ڸ�ϵͳ������ʷ���ݿռ����δ������Ӧ����ʷ���ݿռ䣬����贴���ÿռ�!", vbInformation + vbDefaultButton1, gstrSysName
                            If frmHistorySpaceSet.ShowInstall(Me, gcnOracle, gstrUserName, gstrPassword, val(arrTmp(1)), 0, 0, , True) = False Then
                                If lngCount < 2 Then
                                    GoTo ReDo
                                Else
                                    MsgBox "������δ����ʷ���ݿռ�,���,����ϵͳ���в�����,�������[���ݹ���-->����ת��]�д���!", vbInformation + vbDefaultButton1, gstrSysName
                                End If
                            End If
                        End If
                    End If
                Case FS_����ͬ���
                    'Ϊ���������Ķ��󴴽�����ͬ���('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION')
                    Set cnCurrent = gcnOracle
                    gcnOracle.Execute "Zl_Createpubsynonyms", , adCmdStoredProc
                    Set cnCurrent = Nothing
                Case FS_������Ч����
                    Call ReCompileObjects(IIf(arrTmp(1) = "TOOLS", gcnTools, gcnOracle))
                Case FS_��������
                    Call ReAdjustSequence(IIf(arrTmp(1) = "TOOLS", gcnTools, gcnOracle))
                Case FS_��������
                    Call ImportReports(val(arrTmp(1)))
                Case FS_����ֵ��
                    Call DoHelperMain
                Case FS_��ɫ��Ȩ
                    Call GrantToRole
                Case FS_��̨�Զ�ҵ����
                    Call StartAutoRun(gcnOracle, mclsRunScript)
                Case FS_�ӳٽű�
                    Call GatherStatistics
                    Call SaveRunAfterInfo(gstrServer, mintDDLParallel, mrsHisAfterSPace, mrsHisAfter, mrsSatistics)
                    If gobjFile.FileExists(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & gstrServer & ".SQL") Then
                        Call MsgBox("����������������ɣ����������ں�̨�����ӳٽű���" _
                                & IIf(gblnInIDE, "C:\APPSOFT", App.Path) _
                                & "\RuntimeFile\RunAfter_" & gstrServer & ".SQL���������ڿ��Խ��пͻ����������õȹ�����" _
                            , vbInformation, gstrSysName)
                        Call StartRunAter(gstrServer)
                    End If
            End Select
            
            mclsRunScript.WriteLog
            lngSec = DateDiff("s", datStart, Now)
            mclsRunScript.WriteLog "[" & vsnStep.Text & "]����" _
                & Format(datStart, "HH:mm:ss") & "��" & Format(Now, "HH:mm:ss") _
                & "������ʱ" & IIf(lngSec > 60, (lngSec \ 60) & "����" & (lngSec Mod 60) & "��", lngSec & "��")
            mclsRunScript.WriteLog
            
            If blnFirstUp Then blnFirstUp = False
            Call SetStepStateImg(vsnStep, True)  '��ʼִ��
        Else
            Call SetSQLState(False)
            blnFirstUp = True
            mclsRunScript.WriteSection vsnStep.Text, IIf(i = vsPlan.FixedRows, "=", "-")
            vsnStep.Expanded = True
        End If
        Me.Refresh
    Next
    
    Call UpgradeFinish(True)
    mblnOK = True
    If Not vsnStep Is Nothing Then Call SetStepStateImg(vsnStep, True)  '��ʼִ��
    '��������
    If Not mblnExecBef Then
        Set vsnStep = vsPlan.GetNode(vsPlan.Rows - 1)
        Call SetStepStateImg(vsnStep)  '��ʼִ��
        Call SetStepStateImg(vsnStep, True)  '��ʼִ��
        For i = 1 To 50
            DoEvents
            Call Sleep(100)
        Next
        If MsgBox("������Ǩ��ɺ���Ҫ�Կͻ���վ����в�������," & vbCrLf & "��Ҫ��վ�㲿����������������?", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
            mstrRunModule = "0109"
            Unload Me
        End If
        
'        '�ӳ�ִ�еĽű�
'        blnNormal = True
'        mrsSQLSys.Filter = "SysType=" & ST_App & " And FileType=" & FT_DefUp
'        mrsSQLSys.Sort = "FullSPVer"
'        mclsRunScript.ConnectType = 0: mclsRunScript.IsGather = False
'        If mrsSQLSys.RecordCount > 0 Then
'            blnNormal = True
'            If Mid(mrsSQLSys!SPVer, 1, 5) = "10.25" Then
'                MsgBox "������Ǩ��ɺ󣬽������������ӳٽű����ڴ��ڼ�ϵͳ������ʹ�ã����������б����������(ZLRPTSQLAdjust)��������Դ���漰[���˷��ü�¼]���SQL��䡣", vbInformation, gstrSysName
'            Else
'                MsgBox "������Ǩ��ɺ󣬽������������ӳٽű����ڴ��ڼ�ϵͳ������ʹ�á�", vbInformation, gstrSysName
'            End If
'            Set mclsRunScript.Connection = gcnOracle
'            Do While Not mrsSQLSys.EOF
'                Call RunSQLScript(mrsSQLSys!FilePath, , , False)
'                mrsSQLSys.MoveNext
'            Loop
'        End If
'
'        mrsSQLSys.Filter = "SysType=" & ST_AppHis & " And FileType=" & FT_DefUp
'        mrsSQLSys.Sort = "UserServer,UserName,FullSPVer"
'        If mrsSQLSys.RecordCount > 0 Then
'            Do While Not mrsSQLSys.EOF
'                If strPreBakUserName <> mrsSQLSys!UserName Or Not blnConn Then
'                    strPreBakUserName = mrsSQLSys!UserName
'                    blnConn = True '�Ƿ�����ӳɹ�
'                    If OpenHistoryConnect(Nvl(mrsSQLSys!UserName), Nvl(mrsSQLSys!UserPass), Nvl(mrsSQLSys!UserServer), True) = False Then
'                        'һ��������������.��Ϊ��֮���Ѿ���飬���ﱣ֤�����ǵ�ǰ��ʷ�������
'                        blnConn = False
'                    End If
'                End If
'                If blnConn Then
'                    Set mclsRunScript.Connection = mcnHistory
'                    Call RunSQLScript(mrsSQLSys!FilePath, , , False)
'                End If
'                mrsSQLSys.MoveNext
'            Loop
'        End If
'        blnNormal = False
'        On Error GoTo 0
    End If
    Exit Sub
    
errH:
    tmrRefresh.Enabled = False
    If 0 = 1 Then
        Resume
    End If
    If cnCurrent Is Nothing Then
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, App.Title
        Else
            MsgBox err.Description, vbInformation, App.Title
        End If
        GoTo AbortLine
    ElseIf ADOConnectionError(err, cnCurrent) Then
        If CheckAdoConnection(cnCurrent) Then
            Resume
        Else
            GoTo AbortLine
        End If
    End If
    Exit Sub
    
AbortLine:
    tmrRefresh.Enabled = False
    cmdCancel.Enabled = True
    Call UpgradeFinish(False)
End Sub

Private Sub SetStepStateImg(ByVal vsnCurrent As VSFlexNode, Optional ByVal blnDone As Boolean)
'���ܣ����ò����״̬ͼƬ
'������vsnCurrent=��ǰ�ڵ�
'          blnDone=�Ƿ�ò����Ѿ����
    Dim vsnTmp As VSFlexNode, vsnParent As VSFlexNode
    Dim strImg As String
    strImg = IIf(blnDone, "Finish", "Doing")
    DoEvents
    If Not blnDone Then
        Set vsnTmp = vsnCurrent
        Do While Not vsnTmp Is Nothing
            Set vsnTmp.Image = imgPlan.ListImages(strImg).Picture
            vsnTmp.Expanded = True
            Set vsnTmp = vsnTmp.GetNode(flexNTParent)
        Loop
    Else
        Set vsnTmp = vsnCurrent.GetNode(flexNTNextSibling)
        Set vsnCurrent.Image = imgPlan.ListImages(strImg).Picture
        vsnCurrent.Expanded = False
        Set vsnParent = vsnCurrent
        Do While vsnParent.GetNode(flexNTNextSibling) Is Nothing '�������һ���ڵ����
            Set vsnParent = vsnParent.GetNode(flexNTParent)
            If vsnParent Is Nothing Then Exit Do
            Set vsnParent.Image = imgPlan.ListImages(strImg).Picture
            vsnParent.Expanded = False
        Loop
    End If
    vsPlan.Refresh
End Sub

Private Function UpgradeCheck(ByVal lngSys As Long) As Boolean
'���ܣ���ϵͳ���ж�����
'������lngSys=ϵͳ��
'          strMsg=������Ϣ
    Dim cnTmp As ADODB.Connection
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strCheckFile As String, strName As String
    Dim strResult As String
    
    On Error GoTo errH
    mrsSysInfo.Filter = "ϵͳ���=" & lngSys
    Call SetSQLState(False)
    If lngSys = 0 Then
        Set cnTmp = GetConnection("ZLTOOLS")
        strName = "zlUpgradeCheck"
        strCheckFile = gobjFile.GetParentFolderName(mrsSysInfo!�����ļ�) & "\" & strName & ".sql"
    Else
        Set cnTmp = gcnOldOra
        strName = "zl" & lngSys \ 100 & "_UpgradeCheck"
        strCheckFile = gobjFile.GetParentFolderName(gobjFile.GetParentFolderName(mrsSysInfo!�����ļ�)) & "\�����ű�\" & strName & ".sql"
    End If
    '������麯��
    mclsRunScript.IsUseLog = False
    lblFile.Caption = strCheckFile
    If Not mclsRunScript.ExecuteFile(strCheckFile, , , IIf(lngSys = 0, 1, 0), cnTmp) Then
        mclsRunScript.IsUseLog = True
        GoTo AbortLine
    End If
    mclsRunScript.IsUseLog = True
makSQL:
    err.Clear: On Error Resume Next
    strSQL = "Select " & strName & "('" & VerSpecialNormal(mrsSysInfo!ϵͳ�汾�� & "") & "', '" & VerSpecialNormal(mrsSysInfo!Ŀ��汾 & "") & "') As Info From Dual"
    Set rsTmp = gclsBase.OpenSQLRecord(IIf(lngSys = 0, cnTmp, gcnOracle), strSQL, App.Title)
    If err.Number <> 0 Then '������
        strResult = err.Description
        If ADOConnectionError(err, gcnOracle) Then
            If CheckAdoConnection(gcnOracle) Then
                GoTo makSQL
            Else
                mclsRunScript.WriteLog "�������" & err.Description
            End If
        Else
            mclsRunScript.WriteLog "�������" & strResult
            MsgBox strResult, vbExclamation, gstrSysName: GoTo AbortLine
        End If
    Else
        strResult = rsTmp!Info & ""
        If strResult <> "" Then
            mclsRunScript.WriteLog "�������" & strResult
            MsgBox strResult, vbExclamation, gstrSysName: GoTo AbortLine
        Else
            mclsRunScript.WriteLog "�������ͨ��"
        End If
    End If
    UpgradeCheck = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
    Exit Function
AbortLine:
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function RunScriptByVersion(ByVal lngSys As Long, ByVal strVersion As String, Optional ByVal blnFirstUpdate As Boolean, _
    Optional ByVal strOldVer As String, Optional ByVal strAimVer As String, Optional blnHistory As Boolean, _
    Optional ByVal cnTmp As ADODB.Connection, Optional ByVal strBakDB As String, Optional ByVal blnUpInterface As Boolean) As Boolean
'���ܣ�ִ�нű��ļ�������ϵͳ�汾
'������lngSys=ϵͳ��
'         strVersion=��ǰ����İ汾
'         blnFirstUpdate=�Ƿ��һ����Ǩ�汾����
'         strOldVer=ԭʼ�汾��blnFirstUpdate=True�贫
'         strAimVer=Ŀ��汾��blnFirstUpdate=True�贫
'         blnHistory=�Ƿ���ʷ��汾����
'         cnTmp=���ӣ���ʷ��汾������Ҫ
'         blnUpInterface=�Ƿ���Ǩ�ӿڵ��ã���Ǩ�ӿڵ��ò��ܷ��ʵ�ǰ����ռ�����Լ����ԣ�
'                                   ��ǰ����ʷ�ⵥ������������ߵ��������ӿ�
    Dim strLogSQL As String, strVerSQL As String
    Dim datNow As Date
    Dim cnCurrent As ADODB.Connection
    Dim blnAbort As Boolean
    
    On Error GoTo errH
    With mrsSysFiles
        datNow = Now
        .Filter = "ϵͳ���=" & lngSys & " And SPVer='" & strVersion & "'" & IIf(blnHistory, " And  SysType=" & ST_History & " And ������='" & UCase(strBakDB) & "'", " And SysType<>" & ST_History)
        .Sort = "FileType"
        If .EOF And Not blnUpInterface Then Call SetSQLState(False)
        mclsRunScript.FileVersion = strVersion
        Do While Not .EOF
            If !FileType = FT_DBA Then
                Set mclsRunScript.Connection = gcnSystem: mclsRunScript.ConnectType = 2
            Else
                If lngSys = 0 Then
                    Set mclsRunScript.Connection = gcnTools: mclsRunScript.ConnectType = 1
                ElseIf Not blnHistory Then
                    Set mclsRunScript.Connection = gcnOldOra: mclsRunScript.ConnectType = 0
                Else
                    Set mclsRunScript.Connection = cnTmp: mclsRunScript.ConnectType = 0
                End If
            End If
            Set cnCurrent = mclsRunScript.Connection
            
            If Not RunSQLScript(!FilePath, val(!AbortLine & ""), !Optional & "", blnHistory Or lngSys = 0, blnUpInterface) Then
                blnAbort = True
                If Not blnHistory Then
                    If blnFirstUpdate Then '��һ�θ��°汾,����Zlupgrade������һ���¼�¼
                        strLogSQL = "Insert Into Zlupgrade" & vbNewLine & _
                                    "  (ϵͳ, ԭʼ�汾, Ŀ��汾, ��Ǩʱ��, ��Ǩ���, ����汾, ��ֹ���, ��ǰִ��)" & vbNewLine & _
                                    "  Select " & IIf(lngSys = 0, "Null", lngSys) & ", '" & strOldVer & "', '" & strAimVer & "', Sysdate, 1, '" & IIf(!FileType <= FT_Standard, strOldVer, strVersion) & "','" & Replace(mclsRunScript.AbortInfo, "'", "''") & "', " & IIf(mblnExecBef, 1, "Null") & " From Dual"
                    Else
                        strLogSQL = "Update Zlupgrade a" & vbNewLine & _
                                        "Set ����汾 =" & IIf(!FileType <= FT_Standard, "����汾", "'" & strVersion & "'") & " , ��Ǩ���=1 ,��ֹ���='" & Replace(mclsRunScript.AbortInfo, "'", "''") & "'" & vbNewLine & _
                                        "Where ϵͳ " & IIf(lngSys = 0, " Is Null", "=" & lngSys) & " And ��Ǩʱ�� = (Select Max(��Ǩʱ��) From Zlupgrade Where ϵͳ " & IIf(lngSys = 0, " Is Null", "=" & lngSys) & " And " & IIf(mblnExecBef, " Not ", "") & "  ��ǰִ�� Is Null)"
                    End If
                    '��־����
                    Set cnCurrent = gcnOracle
                    gcnOracle.Execute strLogSQL
                Else
                    Set cnCurrent = cnTmp
                    If Not mblnExecBef Then
                        '��ʽ��������������ǰִ����Ϣ
                        cnTmp.Execute "Update zlbakinfo Set �汾��=" & IIf(!FileType <= FT_Standard, "�汾��", "'" & strVersion & "'") & " ,��ֹ���='" & Replace(mclsRunScript.AbortInfo, "'", "''") & "' Where ϵͳ=" & lngSys
                    Else
                        '��ǰִ�У�������ǰִ�а汾����¼��ǰִ����Ϣ
                        cnTmp.Execute "Update zlbakinfo Set ��ǰ��ֹ���='" & Replace(mclsRunScript.AbortInfo, "'", "''") & "' ,��ǰִ��=1 Where ϵͳ=" & lngSys
                    End If
                End If
                GoTo AbortLine
            End If
            Set cnCurrent = Nothing
            
            .MoveNext
        Loop
    End With
    
    Set cnCurrent = gcnOracle
    If Not blnHistory Then
        If blnFirstUpdate Then '��һ�θ��°汾,����Zlupgrade������һ���¼�¼
            strLogSQL = "Insert Into Zlupgrade" & vbNewLine & _
                        "  (ϵͳ, ԭʼ�汾, Ŀ��汾, ��Ǩʱ��, ��Ǩ���, ����汾, ��ֹ���, ��ǰִ��)" & vbNewLine & _
                        "  Select " & IIf(lngSys = 0, "Null", lngSys) & ", '" & strOldVer & "', '" & strAimVer & "', Sysdate, 0, '" & strVersion & "', Null, " & IIf(mblnExecBef, 1, "Null") & " From Dual"
        Else
            strLogSQL = "Update Zlupgrade a" & vbNewLine & _
                            "Set ����汾 = '" & strVersion & "'" & vbNewLine & _
                            "Where ϵͳ " & IIf(lngSys = 0, " Is Null", "=" & lngSys) & " And ��Ǩʱ�� = (Select Max(��Ǩʱ��) From Zlupgrade Where ϵͳ " & IIf(lngSys = 0, " Is Null", "=" & lngSys) & " And " & IIf(mblnExecBef, " Not ", "") & "  ��ǰִ�� Is Null)"
        End If
        If Not mblnExecBef Then '��ǰִ�в�����汾
            'ϵͳ�汾����
            If lngSys = 0 Then
                strVerSQL = "zlTools.B_Public.Update_Ver"
                '���¹����߰汾��:zlRegInfo
                '������ZLHIS�������Ӵ���,��ΪgcnTools���õ�����������ִ�нű�
                Call OpenCursor(gcnOracle, strVerSQL, strVersion)
            Else
                strVerSQL = "Update Zlsystems Set �汾�� = '" & strVersion & "' Where ��� = " & lngSys
                gcnOracle.Execute strVerSQL
            End If
        End If
        '��־����
        gcnOracle.Execute strLogSQL
    Else
        If Not mblnExecBef Then
            '��ʽ��������������ǰִ����Ϣ
            cnTmp.Execute "Update zlbakinfo Set �汾��='" & strVersion & "' ,��ֹ���=Null,��ǰ��ֹ���=NULL,��ǰִ��=0 Where ϵͳ=" & lngSys
        Else
            '��ǰִ�У�������ǰִ�а汾����¼��ǰִ����Ϣ
            cnTmp.Execute "Update zlbakinfo Set ��ǰ��ֹ���='" & strVersion & "' ,��ǰִ��=1 Where ϵͳ=" & lngSys
        End If
    End If
    Call mclsRunScript.WriteCSVRow(lngSys, strVersion, mclsRunScript.HistoryDB, "", Round((DateDiff("s", datNow, Now)) / 60))
    mclsRunScript.FileVersion = ""
    RunScriptByVersion = True
    '��Ǹð汾���ӳٽű���ִ��
    Call RecUpdate(mrsSysFiles _
        , "ϵͳ���=" & lngSys & " And SPVer='" & strVersion & "'" & _
          IIf(blnHistory _
                , " And  SysType=" & ST_History & " And ������='" & UCase(strBakDB) & "'" _
                , " And SysType<>" & ST_History) & " And FileType=" & FT_Deferred _
        , "�ӳٿ�ִ��" _
        , 1)
    If Not blnUpInterface Then Call SetSQLState(False)
    Exit Function
    
AbortLine: '�������񵽵���ֹ��ת
    mclsRunScript.FileVersion = ""
    If mclsRunScript.Connection.State = adStateClosed Then
        If MsgBox("�������������������жϣ��Ƿ����ԣ�", vbRetryCancel + vbInformation, App.Title) = vbRetry Then
            Resume
        End If
    End If
    If blnUpInterface Then Exit Function
    Call SetSessionParallel(mclsRunScript.Connection)
    Call SetSessionParallel(gcnOldOra)
    If Not blnAbort Then Call UpgradeFinish(False)
    cmdCancel.Enabled = True '��Ȼ����Form_Unload
    MsgBox "ϵͳ��Ǩʧ�ܣ�������Ǩ��־�ļ���������Ӧ����֮�����½�����Ǩ��", vbExclamation, gstrSysName
    Exit Function
    
errH:
'    If mclsRunScript.Connection.State = adStateClosed Then
'        If MsgBox("�������������������жϣ��Ƿ����ԣ�", vbRetryCancel + vbInformation, App.Title) = vbRetry Then
'            Resume
'        End If
'    End If
    If blnAbort Then
        GoTo AbortLine
    Else
        If ADOConnectionError(err, cnCurrent) Then
            If CheckAdoConnection(cnCurrent) Then Resume
        End If
    End If
'    If MsgBox("���������з����������" & vbNewLine & err.Description & vbNewLine & "�Ƿ����ԣ�", vbRetryCancel + vbInformation, App.Title) = vbRetry Then
'        Resume
'    End If
End Function

Private Sub HTablePrivsRepair(ByVal lngSys As Long)
'���ܣ�H��Ȩ������
    Dim objSQL As New clsSQLInfo
    Dim datStart As Date
    
    datStart = Now
    Call SetSQLState(False)
mak01:
    On Error Resume Next
    objSQL.SQL = "Insert Into zlProgPrivs" & vbNewLine & _
            "  (ϵͳ, ���, ����, ����, ������, Ȩ��)" & vbNewLine & _
            "  Select ϵͳ, ���, ����, 'H' || ����, User, 'SELECT'" & vbNewLine & _
            "  From zlProgPrivs" & vbNewLine & _
            "  Where (Upper(������), Upper(����)) In (Select User, ���� From zlBakTables Where ϵͳ = " & lngSys & ") And Upper(Ȩ��) = 'SELECT' And" & vbNewLine & _
            "        ϵͳ = " & lngSys & vbNewLine & _
            "  Minus" & vbNewLine & _
            "  Select ϵͳ, ���, ����, ����, User, Ȩ��" & vbNewLine & _
            "  From zlProgPrivs" & vbNewLine & _
            "  Where ϵͳ = " & lngSys & "  And Upper(Ȩ��) = 'SELECT' And ���� Like 'H%'"
    gcnOracle.Execute objSQL.SQL
    If err.Number <> 0 Then
        If ADOConnectionError(err, gcnOracle) Then
            If CheckAdoConnection(gcnOracle) Then GoTo mak01
        End If
        mclsRunScript.ErrCount = mclsRunScript.ErrCount + 1
        mclsRunScript.WriteLog "�� �� �� SQL��" & GetLogSQL(objSQL)
        mclsRunScript.WriteLog "����(�Ѻ���)��" & err.Description
        err.Clear
    End If
End Sub

Private Sub UpgradeFinish(ByVal blnSuccess As Boolean)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If Not mblnFinal Or Not blnSuccess Then
        Call GatherStatistics
        Call SaveRunAfterInfo(gstrServer, mintDDLParallel, mrsHisAfterSPace, mrsHisAfter, mrsSatistics)
    End If
    Call SetSQLState(False)
    strSQL = "Select ���, �汾��" & vbNewLine & _
                    "From Zlsystems " & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select 0, ���� From Zlreginfo Where ��Ŀ = '�汾��'"
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    mclsRunScript.WriteSection "��Ǩϵͳ" & IIf(blnSuccess, "�ɹ���", "ʧ�ܣ�")
    mclsRunScript.WriteLog "������ʱ�䣺" & Format(CurrentDate, "yyyy-MM-dd HH:mm:ss") & String(4, " ") & "������ʱ�䣺" & Format(Now, "yyyy-MM-dd HH:mm:ss")
    mrsSysInfo.Filter = "����=1"
    mrsSysInfo.Sort = "Sort,ϵͳ���"
    Do While Not mrsSysInfo.EOF
        rsTmp.Filter = "���=" & mrsSysInfo!ϵͳ���
        mclsRunScript.WriteLog IIf(mrsSysInfo!ϵͳ��� <> 0, mrsSysInfo!ϵͳ��� & "-", "") & mrsSysInfo!ϵͳ���� & "��" & mrsSysInfo!ϵͳ�汾�� & "-->" & rsTmp!�汾��
        mrsSysInfo.MoveNext
    Loop
    mclsRunScript.WriteLog
    mclsRunScript.WriteLog "�ܹ������Ĵ��������" & mclsRunScript.ErrCount
    If mclsRunScript.AbortInfo <> "" Then
        mclsRunScript.WriteLog "��ֹ�ļ����ƣ�" & Split(mclsRunScript.AbortInfo, "|")(0)
        mclsRunScript.WriteLog "��ֹ�ļ��кţ�" & Split(mclsRunScript.AbortInfo, "|")(1)
    End If
    Call mclsRunScript.WriteCSVRow("", "", "", "", Round((DateDiff("s", mdatStart, Now)) / 60))
    Call mclsRunScript.CloseLog
    If lblLog.Tag <> lblLogModi.Tag Then
        Call mclsRunScript.LogSave(lblLog.Tag)
    End If
    Exit Sub
    
errH:
    If 0 = 1 Then
        Resume
    End If
    If ErrCenter(err, gcnOracle, False) = 1 Then Resume
End Sub

Private Function RunSQLScript(ByVal strFile As String, Optional ByVal lngAbort As Long, Optional strExecProcs As String, Optional ByVal blnHistory As Boolean, Optional ByVal blnUpInterface As Boolean) As Boolean
'���ܣ�ִ��SQL�ű�
'      strFile=SQL�ű���
'      lngAbort=�жϺ�
'      strExecProcs=ִ���ļ�ʱ��Ϊѡ��Ŀ�ѡ���̡�
'      blnHistory=�Ƿ�����ʷ��ű�
'      blnUpInterface=�Ƿ���Ǩ�ӿڵ��ã���Ǩ�ӿڵ��ò��ܷ��ʵ�ǰ����ռ�����Լ����ԣ�
'                                   ��ǰ����ʷ�ⵥ������������ߵ��������ӿ�
'���أ�RunSQLScript=�ļ��Ƿ�ִ�гɹ�
    Dim strTmp As String, strTmpPath As String, strCaption As String
    Dim blnToolLobLater As Boolean, blnDo As Boolean, blnCLose As Boolean
    
    With mclsRunScript
        .Procedures = strExecProcs
        .ProcMode = 0
        .GatherTables = ""
        If Not blnUpInterface Then
            Call SetSQLState(True, True)
            If ActualLen(strFile) <= 50 Then
                strCaption = "�ļ�:" & strFile
            Else
                strTmpPath = gobjFile.GetParentFolderName(strFile)
                strTmp = gobjFile.GetFileName(strFile)
                strTmpPath = ActualStr(strTmpPath, 50 - ActualLen(strTmp) - 3) & "..."
                strCaption = "�ļ�:" & strTmpPath & "\" & strTmp
            End If
        End If
        'ִ�д洢���̣�˵���ű��ǿ�ѡ�ű�����ѡ�ű����Ǵ洢���̣�ִ��ʱ���ܴ��ж��к�ִ�С�
        If strExecProcs <> "" Then .Abort = 0: .ProcMode = 2
        If .OpenFile(strFile, lngAbort) Then
            If (mintToolLob And (LC_ZLTOOLS_IS3590_OR_GREATER Or LC_ZLHIS_IS3590_OR_GREATER)) <> (LC_ZLTOOLS_IS3590_OR_GREATER Or LC_ZLHIS_IS3590_OR_GREATER) Then
                '��ǰ�����Ϲ�����Lob��ִ������
                If UCase(gobjFile.GetFileName(strFile)) = "ZLUPGRADE10.35.90.SQL" Then
                    blnToolLobLater = True
                End If
            End If
            blnCLose = False
            Call SetSessionParallel(.Connection, True)
            Do While Not .EOF
                blnDo = True
                If blnToolLobLater And .SQLInfo.LobDDL Then
                    'Alter Table zltools.zlRPTGraphs Modify ͼƬ Blob;
                    If .SQLInfo.BlockName = "ZLTOOLS.ZLRPTGRAPHS" Then
                        mclsRunScript.WriteLog String(17, " ") & "��Ϊ�����������ű�" & .SQLInfo.SQL
                        blnDo = False
                    End If
                End If
                '���ݽṹ����������������DLL����Ҫ�رղ���
                If blnDo Then
                    If Not blnUpInterface Then
                        lblFile.Caption = strCaption & "," & .Line
                        prgThis.value = .Line / .LinesCount * 100
                        lblPer.Caption = Format(prgThis.value / 100, "0%")
                        Me.txtSQL.Text = IIf(.SQLInfo.Tip <> "", .SQLInfo.Tip & vbCrLf, "") & .SQLInfo.SQL
                    End If
                
                    If .SQLInfo.LobDDL And .SectionNumber < 2 Then
                        Call SetSessionParallel(.Connection, False)
                        blnCLose = True
                    ElseIf .SectionNumber > 1 And Not blnCLose Then
                        Call SetSessionParallel(.Connection, False)
                        blnCLose = True
                    End If
                    If .ExecuteSQL = False Then
                        Call SetSessionParallel(.Connection, False)
                        blnCLose = True
                        Exit Function
                    End If
                    If .SQLInfo.LobDDL And .SectionNumber < 2 Then
                        Call SetSessionParallel(.Connection, True)
                        blnCLose = False
                    End If
                    If Not blnUpInterface Then Call .CollectTables
                End If
                Call .ReadNextSQL
            Loop
            '����û��SQL���²���û�йرգ��˴��ر�
            If Not blnCLose Then
                Call SetSessionParallel(.Connection, False)
            End If
            RunSQLScript = True
        Else
            RunSQLScript = False
        End If
        If Not blnHistory And Not blnUpInterface Then
            mstrChangeTables = mstrChangeTables & IIf(mstrChangeTables = "", "", ",") & .GatherTables
        End If
    End With
End Function

Private Sub UpdateSysFiles()
'���ܣ�����ZLSysFiles��
    On Error GoTo errH
    If mstrSysCodes = "" Then Exit Sub
    gcnOracle.Execute "Delete From zlSysFiles Where ϵͳ IN (" & mstrSysCodes & ")  And ���� In(1,2)"
    mrsSysInfo.Filter = "ϵͳ���<>0 And ����=1"
    Do While Not mrsSysInfo.EOF
        gcnOracle.Execute "Insert Into zlSysFiles(ϵͳ,����,�ļ���,����,������) Values(" & _
                mrsSysInfo!ϵͳ��� & ",1,'" & Replace(ActualStr(mrsSysInfo!�����ļ�, 100), "'", "''") & "',Sysdate,User)"
        mrsSysInfo.MoveNext
    Loop
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    If ErrCenter(err, gcnOracle, False) = 1 Then Resume
End Sub

Private Sub ReCompileObjects(cnThis As ADODB.Connection)
'���ܣ�����ָ�����������ߵ���Ч����
'������cnThis=����������,����������Բ�ͬ�����ߵ���
    Dim rsObjects As New ADODB.Recordset
    Dim rsDepends As New ADODB.Recordset
    Dim arrObjects As Variant, strCompile As String
    Dim strSQL As String, i As Long
    Dim strUser As String
    Dim arrTmp As Variant
    
    lblFile.Caption = "���ڱ�����Ч���� ...": txtSQL.Text = ""
    prgThis.value = 0: lblPer.Caption = "0%"
    
    On Error GoTo errHandle
    strSQL = _
        "Select User, Object_Name, Object_Type" & vbNewLine & _
        "From User_Objects" & vbNewLine & _
        "Where Object_Type In" & vbNewLine & _
        "      ('PROCEDURE', 'FUNCTION', 'VIEW', 'MATERIALIZED VIEW', 'TRIGGER', 'PACKAGE', 'PACKAGE BODY', 'TYPE', 'TYPE BODY') And" & vbNewLine & _
        "      Object_Name Not Like 'BIN$%' And Status = 'INVALID'" & vbNewLine & _
        "Order By Object_Type, Object_Name"
    rsObjects.CursorLocation = adUseClient
    rsObjects.Open strSQL, cnThis, adOpenKeyset
    If Not rsObjects.EOF Then
        strUser = rsObjects!User
        strSQL = _
            "Select Name, Type, Referenced_Name, Referenced_Type" & vbNewLine & _
            "From User_Dependencies" & vbNewLine & _
            "Where Referenced_Owner = User And Type In ('PROCEDURE', 'FUNCTION', 'VIEW', 'MATERIALIZED VIEW', 'TRIGGER', 'PACKAGE'," & vbNewLine & _
            "       'PACKAGE BODY', 'TYPE', 'TYPE BODY') And" & vbNewLine & _
            "      Referenced_Type In" & vbNewLine & _
            "      ('PROCEDURE', 'FUNCTION', 'VIEW', 'MATERIALIZED VIEW', 'TRIGGER', 'PACKAGE', 'PACKAGE BODY', 'TYPE', 'TYPE BODY') And" & vbNewLine & _
            "      Not(Name=Referenced_Name And Type=Referenced_Type) And" & vbNewLine & _
            "      Name Not Like 'BIN$%' And Referenced_Name Not Like 'BIN$%'"
        rsDepends.CursorLocation = adUseClient
        rsDepends.Open strSQL, cnThis, adOpenKeyset

        ReDim arrObjects(rsObjects.RecordCount - 1) As String
        For i = 1 To rsObjects.RecordCount
            arrObjects(i - 1) = rsObjects!Object_Name & "," & rsObjects!Object_Type
            rsObjects.MoveNext
        Next

        '������Ч����
        For i = 0 To UBound(arrObjects)
            arrTmp = Split(arrObjects(i), ",")
            lblFile.Caption = "���ڱ�����Ч���� " & i + 1 & "/" & (UBound(arrObjects) + 1) & " ..."
            prgThis.value = (i + 1) / (UBound(arrObjects) + 1) * 100
            lblPer.Caption = Format(prgThis.value / 100, "0%")
            Call ComplieObject(cnThis, arrTmp(0), arrTmp(1), rsObjects, rsDepends, strCompile)
        Next
        mclsRunScript.WriteLog RPAD("�������� " & strUser & " �� " & UBound(arrObjects) + 1 & " ����Ч����", 33)
    End If
    Exit Sub
    
errHandle: '�����ڲ�������δ֪�쳣
    If ADOConnectionError(err, cnThis) Then
        If CheckAdoConnection(cnThis) Then Resume
    End If
    'If MsgBox(err.Description, vbRetryCancel + vbCritical, gstrSysName) = vbRetry Then Resume
End Sub

Private Sub ComplieObject(cnThis As ADODB.Connection, ByVal strName As String, ByVal strType As String, _
    rsObjects As ADODB.Recordset, rsDepends As ADODB.Recordset, strCompile As String)
'���ܣ�����ָ������Ч����
'������strCompile=�Ѿ�����Ķ�������
'˵����ReCompileObjects���Ӻ���
    Dim arrObjRef As Variant, strErrInfor As String
    Dim strSQL As String, i As Long

    If InStr(strCompile & ",", "," & strName & ",") > 0 Then Exit Sub

    '�ݹ���뵱ǰ���������õĶ���
    rsDepends.Filter = "Name='" & strName & "' And Type='" & strType & "'" '�������Ϳ�������ݹ����(ͬ��BODY)
    If Not rsDepends.EOF Then
        ReDim arrObjRef(rsDepends.RecordCount - 1) As String
        For i = 1 To rsDepends.RecordCount
            arrObjRef(i - 1) = rsDepends!Referenced_Name & "," & rsDepends!Referenced_Type
            rsDepends.MoveNext
        Next
        For i = 0 To UBound(arrObjRef)
            rsObjects.Filter = "Object_Name='" & Split(arrObjRef(i), ",")(0) & "' And Object_Type='" & Split(arrObjRef(i), ",")(1) & "'"
            If Not rsObjects.EOF Then '���ö���Ҳ����Ч����ʱ
                Call ComplieObject(cnThis, Split(arrObjRef(i), ",")(0), Split(arrObjRef(i), ",")(1), rsObjects, rsDepends, strCompile)
            End If
        Next
    End If

    '���뵱ǰ����
    Select Case strType
    Case "PROCEDURE"
        strSQL = "ALTER PROCEDURE " & strName & " COMPILE"
    Case "FUNCTION"
        strSQL = "ALTER FUNCTION " & strName & " COMPILE"
    Case "VIEW"
        strSQL = "ALTER VIEW " & strName & " COMPILE"
    Case "MATERIALIZED VIEW"
        strSQL = "ALTER MATERIALIZED VIEW " & strName & " COMPILE"
    Case "TRIGGER"
        strSQL = "ALTER TRIGGER " & strName & " COMPILE"
    Case "PACKAGE"
        strSQL = "ALTER PACKAGE " & strName & " COMPILE"
    Case "PACKAGE BODY"
        strSQL = "ALTER PACKAGE " & strName & " COMPILE BODY"
    Case "TYPE"
        strSQL = "ALTER TYPE " & strName & " COMPILE"
    Case "TYPE BODY"
        strSQL = "ALTER TYPE " & strName & " COMPILE BODY"
    End Select
    If strSQL <> "" Then
        txtSQL.Text = txtSQL.Text & strSQL & vbCrLf
        txtSQL.SelStart = Len(txtSQL.Text): DoEvents
    
        strErrInfor = ""
        err.Clear: On Error Resume Next
        cnThis.Execute strSQL
        If cnThis.Errors.Count > 0 Then
            '�������(δ����):Err.Number=0,NativeError=0
            '[Microsoft][ODBC driver for Oracle]�����Ĺ��̻���������б������
            'û�и���Ľ����
            If Not (cnThis.Errors(0).NativeError = 0 And cnThis.Errors.Count = 1) Then
                If cnThis.Errors(0).NativeError <> 0 Then
                    strErrInfor = strName & ":" & cnThis.Errors(0).Description
                Else
                    strErrInfor = strName & ":�����������"
                End If
            End If
        End If
        If strErrInfor <> "" Then
            mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "��" & strSQL & "������" & strErrInfor
        End If
        err.Clear: On Error GoTo 0
        strCompile = strCompile & "," & strName
    End If
End Sub

Private Sub ReAdjustSequence(ByVal cnThis As ADODB.Connection, Optional ByVal blnBaseTable As Boolean)
'���ܣ����µ�������
'������cnThis=����������,����������Բ�ͬ�����ߵ���
'blnBaseTable=�Ƿ�ֻ��Ӧ��ϵͳ�Ļ����������������
    Dim rsTemp As ADODB.Recordset
    Dim i As Long, K As Long, lngCount As Long, lngAdjustCount As Long
    Dim strUser As String, strError As String
    
    On Error GoTo errHandle

    txtSQL.Text = "": txtSQL.Enabled = False: txtSQL.BackColor = Me.BackColor
    prgThis.value = 0: lblPer.Caption = "0%"
    If Not AdjustAllSequence(Me, cnThis, , True, , , True, lblFile, prgThis, lblPer, lngCount, lngAdjustCount, strError) Then
        mclsRunScript.WriteLog RPAD("�ܹ�����" & lngCount & "������", "ʵ�ʶ�" & K & " �����н�������������������Ϣ��" & strError, 33)
    Else
        mclsRunScript.WriteLog RPAD("�ܹ�����" & lngCount & "������", 33)
    End If
    txtSQL.Enabled = True: txtSQL.BackColor = &H80000005
    Exit Sub
    
errHandle: '�����ڲ�������δ֪�쳣
    'If MsgBox(err.Description, vbRetryCancel + vbCritical, gstrSysName) = vbRetry Then Resume
End Sub

Private Sub ImportReports(ByVal lngSys As Long)
'���ܣ����뱨��
'˵����������ֹ��Ǩ
    Dim i As Long, lngCount As Long, lngAll As Long
    Dim datStart As Date, lngSec As Long
    
    datStart = Now
    mrsReport.Filter = "ϵͳ���=" & lngSys & " And ��������<>0"
    lngAll = mrsReport.RecordCount
    mrsReport.Sort = "ID"
    lblFile.Caption = "���ڵ��뱨�� ...": txtSQL.Text = ""
    prgThis.value = 0: lblPer.Caption = "0%"
    If gobjReport Is Nothing Then
        Set gobjReport = GetZL9Report
    End If
    If gobjReport Is Nothing Then
        txtSQL.Text = "����������ʧ��,���ܶԱ�����е���!"
        mclsRunScript.ErrCount = mclsRunScript.ErrCount + 1
        mclsRunScript.WriteLog String(4, " ") & txtSQL.Text: Sleep 2000: Exit Sub
    End If
    lngCount = 0
    
    For i = 1 To mrsReport.RecordCount
        prgThis.value = i / (mrsReport.RecordCount) * 100
        lblPer.Caption = Format(prgThis.value / 100, "0%")
        lblFile.Caption = "���ڵ��뱨�� " & i & "/" & mrsReport.RecordCount & " ..."
        txtSQL.Text = txtSQL.Text & "����:" & mrsReport!��� & "/" & mrsReport!����
        txtSQL.SelStart = Len(txtSQL.Text): DoEvents
        If gobjFile.FileExists(mrsReport!FilePath) Then
            '###
            If gobjReport.ReportImport(mrsReport!FilePath, gcnOracle, mrsReport!���, mrsReport!�������� = 2) Then
                txtSQL.Text = txtSQL.Text & ",�ɹ�!"
                mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "��" & RPAD(mrsReport!FilePath, 70) & String(4, " ") & IIf(mrsReport!�������� = 2, ",��������Դ�ɹ�", "���嵼��ɹ�")
            Else
                lngCount = lngCount + 1
                txtSQL.Text = txtSQL.Text & ",ʧ��!"
                mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "��" & RPAD(mrsReport!FilePath, 70) & String(4, " ") & IIf(mrsReport!�������� = 2, ",��������Դʧ��", "���嵼��ʧ��")
            End If
        Else
            lngCount = lngCount + 1
            txtSQL.Text = txtSQL.Text & ",�ļ�������!"
            mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "��" & "��ʧ�ļ�:" & RPAD(mrsReport!FilePath, 65) & String(4, " ") & IIf(mrsReport!�������� = 2, ",��������Դ", "���嵼��")
        End If
        txtSQL.Text = txtSQL.Text & vbCrLf
        txtSQL.SelStart = Len(txtSQL.Text): DoEvents
        mrsReport.MoveNext
    Next
    lngSec = DateDiff("s", datStart, Now)
    mclsRunScript.WriteLog RPAD("��" & (lngAll - lngCount) & " �ű�����ɹ�," & lngCount & "�ű�����ʧ��", 33)
    mclsRunScript.ErrCount = mclsRunScript.ErrCount + lngCount
End Sub

Private Sub GrantToRole()
    Dim lngCount As Long

    On Error Resume Next
    '����Ȩ�ޱ�����д��Ȩ��
    Call SetSQLState(True)
    lblFile.Caption = "���ڶԽ�ɫ������Ȩ ..."
    Call ReGrantToRole(gcnOracle, "", True, gstrUserName, prgThis, lblPer, lngCount)
    mclsRunScript.WriteLog RPAD("���� " & lngCount & " ����ɫ������������Ȩ", 33)
    txtSQL.Enabled = True: txtSQL.BackColor = &H80000005
End Sub

Private Sub GatherStatistics()
'���ܣ��Ѽ�ͳ����Ϣ������ʷ������ʱ��ֻ�Ѽ���ʷ�⣬������ʷ�������߿���Ѽ���
    Dim strSQL      As String, rsTmp As ADODB.Recordset
    Dim rsGraTable  As ADODB.Recordset, rsBakTable  As ADODB.Recordset
    Dim lngCount    As Long, i As Long, lngCur As Long
    Dim strUser     As String, strOtherPara As String
    Dim datStart    As Date, datStartTmp As Date, lngSec As Long
    Dim lngErr      As Long
    Dim lngID       As Long
    Dim blnDo       As Boolean
    
    SetSQLState (True)
    Set mrsSatistics = CopyNewRec(Nothing, , , Array("ID", adInteger, Empty, Empty, "Owner", adVarChar, 100, Empty, "TableName", adVarChar, 100, Empty, _
                                                "SQL", adVarChar, 500, Empty))
    lblFile.Caption = "���ڶԴ�����ͳ����Ϣ�ռ� ..."
    datStart = Now
    On Error Resume Next
    strSQL = "Select Distinct A.����" & vbNewLine & _
                    "From (Select ����" & vbNewLine & _
                    "       From Zlbigtables" & vbNewLine & _
                    "       Where ϵͳ in(" & mstrSysCodes & ")" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select ���� From zlBakTables Where ϵͳ in(" & mstrSysCodes & ")) A"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    If err.Number <> 0 Then
        err.Clear
        strSQL = "Select Distinct ����" & vbNewLine & _
                "From zlBakTables" & vbNewLine & _
                "Where ϵͳ in(" & mstrSysCodes & ")" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select Column_Value From Table(F_Str2list('������Ϣ,������ҳ,������Ϣ�ӱ�,������ҳ�ӱ�,����ǼǼ�¼,ҽ�����˵���,ҽ�����˹�����'))"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
        If err.Number <> 0 Then err.Clear
    End If
    strSQL = "Select ����,ϵͳ" & vbNewLine & _
            "From zlBakTables" & vbNewLine & _
            "Where ϵͳ in(" & mstrSysCodes & ")"
    Set rsBakTable = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    If err.Number = 0 Then
        On Error GoTo errH
        Set rsGraTable = CopyNewRec(rsTmp, , , Array("�ռ�", adInteger, Empty, Empty))
        mstrChangeTables = "," & UCase(mstrChangeTables) & ","
        mstrChangeTables = Replace(Replace(Replace(mstrChangeTables, vbNewLine, ""), ",,", ","), ",,", ",")
        
        '�����Ҫ�ռ��ı�
        rsGraTable.Filter = ""
        For i = 1 To rsGraTable.RecordCount
            If mstrChangeTables = "," Then Exit For
            If mstrChangeTables Like "*," & UCase(rsGraTable!����) & ",*" Then
                If ",������Ϣ,������ҳ,������Ϣ�ӱ�,������ҳ�ӱ�,����ǼǼ�¼,ҽ�����˵���,ҽ�����˹�����," Like "*," & rsGraTable!���� & ",*" Then
                    rsGraTable.Update "�ռ�", 2
                Else
                    rsGraTable.Update "�ռ�", 1
                End If
            Else
                rsGraTable.Update "�ռ�", 0
            End If
            mstrChangeTables = Replace(Replace(mstrChangeTables, "," & UCase(rsGraTable!����) & ",", ","), ",,", ",")
            rsGraTable.MoveNext
        Next
        
        mrsHistorySpace.Filter = "����=1 And ��ǰ=1 And Db����=Null"
        rsGraTable.Filter = "�ռ�=1"
        lngCount = rsGraTable.RecordCount * mrsHistorySpace.RecordCount
        rsGraTable.Filter = "�ռ�<>0"
        lngCount = lngCount + rsGraTable.RecordCount
        
        'i=0 ��ʶ���߿�ͳ����Ϣ�ռ�����ʷ���ռ��������߿���ͬ
        strOtherPara = ",cascade => True" & _
                        ",method_opt => 'for all columns size skewonly'" & _
                        IIf(mintDDLParallel = 0, "", ",degree => " & mintDDLParallel) & ",no_invalidate => false)"
'        Set cnDBA = GetConnection("DBA")
        
        For i = 0 To mrsHistorySpace.RecordCount
            If i = 0 Then
                mclsRunScript.WriteLog "�ռ�ͳ����Ϣ�Ĳ�����" & Mid(strOtherPara, 2), , 1
                strUser = gstrUserName
                rsGraTable.Filter = "�ռ�<>0"
            Else
                strUser = mrsHistorySpace!������
                If i = 1 Then rsGraTable.Filter = "�ռ�=1"
            End If
            If rsGraTable.RecordCount <> 0 Then rsGraTable.MoveFirst
            DoEvents
            Do While Not rsGraTable.EOF
                lngCur = lngCur + 1
                prgThis.value = lngCur / lngCount * 100
                lblPer.Caption = Format(prgThis.value / 100, "0%")
                lblFile.Caption = "���ڶԱ�:" & strUser & "." & rsGraTable!���� & "����ͳ����Ϣ�Ѽ� ..."
                datStartTmp = Now
                Me.Refresh
                If i > 0 Then
                    rsBakTable.Filter = "����='" & rsGraTable!���� & "' And ϵͳ=" & mrsHistorySpace!ϵͳ���
                    If Not rsBakTable.EOF Then
                        blnDo = True
                    End If
                Else
                    blnDo = True
                End If
                If blnDo Then
                    strSQL = "dbms_stats.gather_table_stats(ownname => '" & strUser & "',tabname =>'" & rsGraTable!���� & "'" & strOtherPara
                    mrsSatistics.AddNew Array("ID", "Owner", "TableName", "SQL"), Array(Identity(lngID), UCase(strUser), rsGraTable!����, strSQL)
                
    '                If optStatType(0).value Then 'ֱ�������������ռ�
    '                    '���ð�ʱָ������������ODBC���ӷ�ʽ֧��
    '                    '��connection����excute������Options����ֵΪ�⼸�������ԣ�adCmdUnknown 'adCmdStoredProc 'adExecuteNoRecords
    '                    '��Command���󣬱���ָ��CommandType = adCmdStoredProc
    '                    On Error Resume Next
    '                    cnDBA.Execute strSQL, , adExecuteNoRecords
    '                    If err.Number = 0 Then
    '                        lngSecTmp = DateDiff("s", datStartTmp, Now)
    '                        mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "��" & RPAD(strUser & "." & rsGraTable!����, 50) & "��ʱ��" & IIf(lngSecTmp > 60, (lngSecTmp \ 60) & "����" & (lngSecTmp Mod 60) & "��", lngSecTmp & "��")
    '                    Else
    '                        mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "��" & RPAD(strUser & "." & rsGraTable!���� & String(8, " ") & "�ռ�ʧ��", 50) & "����" & err.Description & String(8, " ") & "SQL:" & strSQL
    '                        err.Clear: lngErr = lngErr + 1
    '                    End If
    '                Else '����¼�ռ���
                    mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "��" & strSQL
    '                End If
                End If

                rsGraTable.MoveNext
            Loop
            If i <> 0 Then mrsHistorySpace.MoveNext
        Next
        lngSec = DateDiff("s", datStart, Now)
        mclsRunScript.WriteLog "���� " & lngCount & " ������Ҫ������ͳ����Ϣ�ռ�", , 1
    Else
        mclsRunScript.WriteLog "����δ��ѯ������Ĵ�����û�ж��κα����ͳ����Ϣ�ռ�"
    End If
    mclsRunScript.ErrCount = mclsRunScript.ErrCount + lngErr
    SetSQLState
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Private Sub RepairHisDB(ByVal cnHistory As ADODB.Connection, ByVal lngSys As Long, ByVal strBakUser As String, ByVal strBakServer As String, _
    ByVal strBakSpaceName As String, ByVal strDbLink As String, Optional ByVal blnCurDB As Boolean, Optional ByVal blnAloneUpHistory As Boolean)
'���ܣ�������ʷ��ṹ����
'������blnAloneUpHistory-True:����������ʷ��,false:ϵͳ����������������ʷ��
    Dim datStartTmp As Date, lngSecTmp As Long
    Dim rsRepairSQL As ADODB.Recordset, lngCount As Long, i As Long
    Dim comTmp As New ADODB.Command
    
    On Error GoTo errH
    If Not blnAloneUpHistory Then
        Call SetSQLState(True, True)
        lblFile.Caption = "���ڼ����ʷ��ṹ���� ..."
    End If
    datStartTmp = Now
    
    '�Ѽ���ʷ������SQL
    Call frmHistorySpaceRepair.ShowRepair(Me, lngSys, True, strBakUser, strBakSpaceName, blnCurDB, rsRepairSQL, cnHistory, strDbLink)
    lngSecTmp = DateDiff("s", datStartTmp, Now)
    If Not rsRepairSQL Is Nothing Then
        mclsRunScript.WriteLog RPAD("��ʷ��ṹ��鷢��" & rsRepairSQL.RecordCount & "������", 30) & ",��ʱ" & IIf(lngSecTmp > 60, (lngSecTmp \ 60) & "����" & (lngSecTmp Mod 60) & "��", lngSecTmp & "��")
        rsRepairSQL.Filter = "ExecLater=0"
        rsRepairSQL.Sort = "ExecOrder,FixType,ID"
        lngCount = rsRepairSQL.RecordCount: datStartTmp = Now
        If lngCount <> 0 And Not blnAloneUpHistory Then lblFile.Caption = "��������" & strBakUser & "�Ľṹ���� ..."
        Call SetSessionParallel(cnHistory, True)
        Call SetSessionParallel(gcnOracle, True)
        On Error Resume Next
        For i = 1 To rsRepairSQL.RecordCount
            '��ʽ��������ʷ��������ִ�п����Ӻ�ִ��SQL,����ЩSQL���ں������ִ��
            If rsRepairSQL!ExecLater = 0 Or blnAloneUpHistory Then
                If Not blnAloneUpHistory Then
                    prgThis.value = i / lngCount * 100
                    lblPer.Caption = Format(prgThis.value / 100, "0%")
                    Me.Refresh
                End If
                If rsRepairSQL!ExecDB = 1 Then
                    Set comTmp.ActiveConnection = gcnOracle
                Else
                    Set comTmp.ActiveConnection = cnHistory
                End If
                comTmp.CommandText = rsRepairSQL!SQL
mak01:
                DoEvents
                comTmp.Execute
                If err.Number <> 0 Then
                    If ADOConnectionError(err, comTmp.ActiveConnection) Then
                        If CheckAdoConnection(comTmp.ActiveConnection) Then GoTo mak01
                    End If
                    mclsRunScript.ErrCount = mclsRunScript.ErrCount + 1
                    mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "��" & IIf(rsRepairSQL!ExecDB = 0, "��ʷ�⣺" & strBakUser & "��", "���߿⣬") & rsRepairSQL!SQL
                    mclsRunScript.WriteLog "�����Ѻ��ԣ���" & err.Description
                    err.Clear
                End If
            End If
            rsRepairSQL.MoveNext
        Next
        
        Call SetSessionParallel(cnHistory, False)
        Call SetSessionParallel(gcnOracle, False)
    End If
    '�������Ӻ�ִ�е����ݿⱣ��
    If Not blnAloneUpHistory Then
        If Not rsRepairSQL Is Nothing Then
            rsRepairSQL.Filter = "ExecLater=1"
            rsRepairSQL.Sort = "ID"
            If mrsHisAfter Is Nothing Then
                Set mrsHisAfter = CopyNewRec(rsRepairSQL, True, , Array("DB_ID", adInteger, Empty, Empty))
            End If
            mrsHisAfterSPace.Filter = ""
            If rsRepairSQL.RecordCount = 0 Then '����ʷ��û�п����Ӻ���Ľű����Զ�ɾ������ʷ��
                Call RecDelete(mrsHisAfterSPace, "ϵͳ���=" & lngSys & " And ����='" & strBakSpaceName & "'")
            Else '���Ӻ�ű���������
                Call RecDataAppend(mrsHisAfter, rsRepairSQL, , "-DB_ID", , , Array("DB_ID", mrsHisAfterSPace.RecordCount))
            End If
        Else '����ʷ��û�п����Ӻ���Ľű����Զ�ɾ������ʷ��
            Call RecDelete(mrsHisAfterSPace, "ϵͳ���=" & lngSys & " And ����='" & strBakSpaceName & "'")
        End If
    End If
    If strDbLink = "" Then
         If Not blnAloneUpHistory Then lblFile.Caption = "��������" & strBakUser & "�ķ���Ȩ������ ..."
        '��Ҫ������Ȩ,��������:���˺�20071202
        Call GrantBakToUser(cnHistory, gstrUserName)
    End If
    If blnCurDB Then
         If Not blnAloneUpHistory Then
            lblFile.Caption = "��������" & strBakUser & "����ʷ���ݿռ���ͼ ..."
            lblPer.Caption = ""
        End If
        Call CreateAppView(gstrUserName, strBakUser, lngSys, IIf(strDbLink = "", "", "@" & strDbLink), IIf(blnAloneUpHistory, Nothing, prgThis), mclsRunScript)
    End If
     If Not blnAloneUpHistory Then Me.Refresh
    Exit Sub
    
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Private Sub SetSessionParallel(ByRef cnInput As ADODB.Connection, Optional ByVal blnEnabled As Boolean)
'���û����DDL
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If mintDDLParallel <= 1 Then Exit Sub
    If blnEnabled Then
        strSQL = "Alter Session FORCE PARALLEL DDL PARALLEL " & mintDDLParallel
        cnInput.Execute strSQL
    Else
        strSQL = "ALTER Session DISABLE PARALLEL DDL "
        cnInput.Execute strSQL
        strSQL = "Select 'alter index ' || Index_Name || ' noparallel' SQL" & vbNewLine & _
                    "From User_Indexes" & vbNewLine & _
                    "Where Degree Not In ('0', '1') and index_type='NORMAL' And temporary='N' " & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select 'alter table ' || Table_Name || ' noparallel' SQL" & vbNewLine & _
                    "From User_Tables" & vbNewLine & _
                    "Where Degree != ('         1')"
        Set rsTmp = gclsBase.OpenSQLRecord(cnInput, strSQL, App.Title)
        On Error Resume Next
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                cnInput.Execute rsTmp!SQL, , adCmdText
                If err.Number <> 0 Then
                    mclsRunScript.WriteLog "ȡ�����г���" & rsTmp!SQL
                    If cnInput.Errors.Count > 0 Then
                        mclsRunScript.WriteLog "�����Ѻ��ԣ���" & cnInput.Errors(0).Description
                    Else
                        mclsRunScript.WriteLog "�����Ѻ��ԣ���" & err.Description
                    End If
                    err.Clear
                End If
                rsTmp.MoveNext
            Loop
        End If
    End If
    Exit Sub
    
errH:
    If 0 = 1 Then
        Resume
    End If
    If ErrCenter(err, cnInput, False) = 1 Then Resume
End Sub

Private Function GetCode(ByVal strCaption As String) As String
'���ܣ���ȡ���̵ı���
    Dim arrTmp As Variant, i As Long
    Dim strCode As String
    
    arrTmp = Split(strCaption, ".")
    For i = LBound(arrTmp) To UBound(arrTmp)
        If i <> UBound(arrTmp) Then
            strCode = strCode & "." & arrTmp(i)
        End If
    Next
    GetCode = Mid(strCode, 2)
End Function

Private Sub SetCpuCount()
'���ܣ�����ͳ����Ϣ�ռ��Լ�����DDL�Ĳ��ж�
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim intDefault As Integer, intMax As Integer, intMin As Integer
    Dim blnCanDDL   As Boolean
    On Error Resume Next
    strSQL = "Select Nvl(Max(Value),0) DDLSize From V$parameter Where Name = 'parallel_execution_message_size'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ����parallel_execution_message_size")
    blnCanDDL = rsTmp!DDLSize >= 8192
    
     '�����ΪCPU������ֹ���ߣ�ʵ��ΪCPU����*����CPU�ϲ��н���
'    Dim intPerParallel As Integer
'    strSQL = "Select Nvl(Max(Value),0) Parallel From V$parameter Where Name = 'parallel_threads_per_cpu'"
'    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ����CPU������")
'    intPerParallel = Val(rsTmp!Parallel, "")
'    intPerParallel = IIf(intPerParallel < 1, 1, intPerParallel) '�����Ա�̣����˽�ʵ��ORacle����������
    strSQL = "Select Nvl(Max(Value),0) CPU From V$parameter Where Name = 'cpu_count'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ����CUP��")
    
    intMin = 1
    If rsTmp!cpu <= 4 Or Not blnCanDDL Then
        chkParallel.value = 0: chkParallel.Enabled = False ': lblStaCpuName.Tag = "Cpu<=4"
        intDefault = 1
        intMax = IIf(rsTmp!cpu = 0, 1, rsTmp!cpu)
        If rsTmp!cpu <= 4 Then
            lblCpuWarn.Caption = "δ����4��CPU�����ܲ��У�"
        Else
            lblCpuWarn.Caption = "parallel_execution_message_size<8192�����ܲ��У�"
        End If
        lblCpuWarn.Visible = True: lblCpuWarn.Tag = "��ʾ����"
        Call SetCtrlPosOnLine(False, 0, lblCpuWarn, 60, ckhIdxOnLine)
    ElseIf rsTmp!cpu <= 8 Then
        intDefault = 4
        intMax = rsTmp!cpu
    ElseIf rsTmp!cpu <= 12 Then
        intDefault = 8
        intMax = rsTmp!cpu
    Else
        intDefault = 12
        intMax = rsTmp!cpu
    End If
    txtCpu.Text = intDefault
    udCpu.Max = intMax '�����ֻΪCPU������ֹ���ߣ�ʵ��ΪCPU����*����CPU�ϲ��н���
    udCpu.Min = intMin
End Sub

Private Sub SetSQLState(Optional ByVal blnStart As Boolean, Optional ByVal blnSQLEnable As Boolean)
    lblFile.Caption = "": txtSQL.Text = ""
    prgThis.value = 0: lblPer.Caption = "0%"
    lblPer.Visible = blnStart
    lblFile.Visible = blnStart
    prgThis.Visible = blnStart
    lblPer.Visible = blnStart
    lblPerCap.Visible = blnStart
    txtSQL.Enabled = blnSQLEnable
    txtSQL.BackColor = IIf(blnSQLEnable, &H80000005, Me.BackColor)
End Sub

Private Sub vsPlan_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Function JudgeOldToolsVer() As String
'���ܣ��жϹ����ߵİ汾
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ��� from zlSvrTools where ���='0502'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡzlSvrTools")
    If rsTmp.EOF = True Then
        '��������ģ��汾Ϊ9.0.0
        JudgeOldToolsVer = "9.0.0"
        Exit Function
    End If
    
    strSQL = _
        "SELECT CONSTRAINT_NAME FROM All_Constraints C WHERE C.CONSTRAINT_NAME='ZLOPTIONS_PK' AND C.OWNER='ZLTOOLS' AND C.TABLE_NAME='ZLOPTIONS'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "�б�ZLOPTIONS_PK")
    If rsTmp.EOF = True Then
        '���������ZLOPTIONS_PKԼ����˵��û��ִ�еڶ��������ű����汾Ϊ9.1.0
        JudgeOldToolsVer = "9.1.0"
        Exit Function
    End If
    strSQL = _
        "SELECT CONSTRAINT_NAME FROM All_Constraints C WHERE C.CONSTRAINT_NAME='ZLXLSVERIFY_FK_�����' AND C.OWNER='ZLTOOLS' AND C.TABLE_NAME='ZLXLSVERIFY'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, _
        "�б�ZLXLSVERIFY_FK_�����")
    If Not rsTmp.EOF Then
        '�������ZLXLSVERIFY_FK_�����Լ����˵��û��ִ�е����������ű����汾Ϊ9.2.0
        JudgeOldToolsVer = "9.2.0"
        Exit Function
    End If
    JudgeOldToolsVer = "9.3.0"
    Exit Function
errH:
    MsgBox err.Description, vbCritical, gstrSysName
    err.Clear
End Function

Private Sub AdjustZLupgrade()
'����ZLupgrade��Ŀ��汾
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error Resume Next
    strSQL = "Select a.Owner" & vbNewLine & _
        "From All_Tab_Columns a" & vbNewLine & _
        "Where a.Table_Name = 'ZLUPGRADE' And a.Column_Name = 'Ŀ��汾' And a.Data_Length < 20"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "�б�ZLUPGRADEĿ��汾����")
    If Not rsTmp.EOF Then
        gcnOracle.Execute "alter table " & rsTmp!Owner & ".ZLUPGRADE modify Ŀ��汾 varchar2(20)", , adCmdText
    End If
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub LogOracleSet()
'���ܣ���¼Oracle������
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    mclsRunScript.WriteLog "Oracle Version    :" & GetOracleVersion(False)
    mclsRunScript.WriteLog "Oracle Parameter"
    On Error GoTo errH
    strSQL = "Select a.Name, a.Display_Value" & vbNewLine & _
            "From V$parameter A" & vbNewLine & _
            "Where a.Name In" & vbNewLine & _
            "      ('O7_DICTIONARY_ACCESSIBILITY', 'audit_trail', 'cluster_database', 'compatible', 'cpu_count'," & vbNewLine & _
            "       'db_file_multiblock_read_count', 'log_buffer', 'memory_max_target', 'memory_target', 'optimizer_features_enable'," & vbNewLine & _
            "       'optimizer_index_caching', 'optimizer_index_cost_adj', 'optimizer_mode', 'optimizer_use_sql_plan_baselines'," & vbNewLine & _
            "       'parallel_execution_message_size', 'pga_aggregate_target', 'sga_max_size', 'sga_target')" & vbNewLine & _
            "Order By a.Name"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "LogOracleSet")
    
    Do While Not rsTmp.EOF
        mclsRunScript.WriteLog "  " & RPAD(rsTmp!name, 35) & "=" & rsTmp!Display_Value
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    mclsRunScript.WriteLog "Oracle Sets(Error):" & err.Description
    err.Clear
End Sub


'--------------------------------------------------------------------------------------------------
'�ӿ�           RunAfterInfo
'����           -�����Ӻ�ִ����Ϣ�����ܴ����ж�
'����ֵ         Boolean
'����б�:
'������         ����                        ˵��
'strPath       String                      �Ӻ�ִ�еĽű��ļ���
'strServer     String                      ������
'intDDLParallel Integer                    ���в���
'rsHisDBInfo   ADODB.Recordset             ��ʷ����Ϣ
'rsHisRunafter ADODB.Recordset             ��ʷ���ִ�нű�
'rsStatistics  ADODB.Recordset             ͳ����Ϣ�Ľű�
'-------------------------------------------------------------------------------------------------
Private Function SaveRunAfterInfo(ByVal strServer As String, ByVal intDDLParallel As Integer, ByVal rsHisDBInfo As ADODB.Recordset, ByVal rsHisRunafter As ADODB.Recordset, ByVal rsStatistics As ADODB.Recordset) As Boolean
    Dim objTxt              As TextStream
    Dim strLine             As String, strCurServer     As String
    Dim strNoDDLParallelSQL As String
    Dim strCurCon           As String
    Dim i                   As Long
    Dim rsFiles             As ADODB.Recordset
    
    On Error GoTo errH
    '--[SERVER]:Oracle
    '--[SCRIPT]:SerializeMulti("V1" & intDDLParallel, strServer, gstrUserName, Sm4EncryptEcb(gstrPassword), Sm4EncryptEcb(txtToolsPwd.Text), txtDBAUser.Text, Sm4EncryptEcb(txtDBAPwd.Text), Sm4EncryptEcb(gclsBase.Serialize(rsHisDBInfo), G_APP_KEY), rsHisRunafter, rsStatistics, rsFiles)
    '--[�ű�����]:
    'SQL
    '--[SCRIPT]:Serialize(Array("DDLPARALLEL",Serialize("HISTORY"),Serialize("HISTORYSCRIPT"),Serialize("STATICTICSSCRIPT")))
    '--[�ű�����]:
    'SQL
    
    If gobjFSO.FileExists(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & strServer & ".SQL") Then
        Set objTxt = gobjFSO.OpenTextFile(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & strServer & ".SQL", ForAppending)
    Else
        If Not gobjFSO.FolderExists(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile") Then
            Call gobjFSO.CreateFolder(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile")
        End If
        Set objTxt = gobjFSO.OpenTextFile(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & strServer & ".SQL", ForWriting, True)
        objTxt.WriteLine "--[ִ��˵��]:�ýű���������ʷ��Ĳ��ֽṹ������Ӱ�����ܵĽṹ����ͳ����Ϣ�ռ�����ʷ��ṹ��������ͨ��������ת�ƹ����е���ʷ��ṹ��������ֹ�ִ�иýű���"
        objTxt.WriteLine "--[SERVER]:" & strServer
    End If
    mrsSysFiles.Filter = "�ӳٿ�ִ��=1"
    Set rsFiles = CopyNewRec(mrsSysFiles)
    '������ʷ��ű���ID���У���֤IDΨһ��
    If Not rsHisRunafter Is Nothing Then
        rsHisRunafter.Filter = ""
'        rsHisRunafter.UpdateBatch adAffectAllChapters
        '��ִ��˳��������
        rsHisRunafter.Sort = "BAKDBName,BAKUser,������,DBLINK,DB_ID,ExecOrder,FixType,ID"
        rsHisRunafter.Sort = "" '��ֹ���º�MoveNext������������
        i = 0
        Do While Not rsHisRunafter.EOF
            rsHisRunafter.Update "ID", Identity(i)
            rsHisRunafter.MoveNext
        Loop
    End If
    objTxt.WriteLine "--[SCRIPTVERSION]:V1"
    objTxt.WriteLine "--[SCRIPT]:" & gclsBase.SerializeMulti("V1", intDDLParallel, strServer, gstrUserName, Sm4EncryptEcb(gstrPassword, G_APP_KEY), Sm4EncryptEcb(txtToolsPwd.Text, G_APP_KEY), txtDBAUser.Text, Sm4EncryptEcb(txtDBAPwd.Text, G_APP_KEY), Sm4EncryptEcb(gclsBase.Serialize(rsHisDBInfo), G_APP_KEY), rsHisRunafter, rsStatistics, rsFiles)
    '̫������������������
'    strLine = gclsBase.SerializeMulti(intDDLParallel, Sm4EncryptEcb(gclsBase.Serialize(rsHisDBInfo), G_APP_KEY), rsHisRunafter, rsStatistics)
    objTxt.WriteLine "--[����ʱ��]:" & Format(Now, "YYYY-MM-DD HH:mm:ss")
    objTxt.WriteLine "--[1.��ʷ������]���ֹ�ִ����ע�ⰴ˳���л���Ӧ����ִ��,Ҳ����������ת�ƹ����е���ṹ�����������ṹ��"
    strNoDDLParallelSQL = "ALTER Session DISABLE PARALLEL DDL;" & vbNewLine & _
                         "Declare" & vbNewLine & _
                         "Begin" & vbNewLine & _
                         "  For Rs In (Select Sql" & vbNewLine & _
                         "             From (Select 'alter index ' || Index_Name || ' noparallel' Sql" & vbNewLine & _
                         "                    From User_Indexes" & vbNewLine & _
                         "                    Where Degree Not In ('0', '1') And Index_Type = 'NORMAL' And temporary='N'" & vbNewLine & _
                         "                    Union All" & vbNewLine & _
                         "                    Select 'alter table ' || Table_Name || ' noparallel' Sql" & vbNewLine & _
                         "                    From User_Tables" & vbNewLine & _
                         "                    Where Degree != ('         1'))) Loop" & vbNewLine & _
                         "    Begin" & vbNewLine & _
                         "      Execute Immediate Rs.Sql;" & vbNewLine & _
                         "    Exception" & vbNewLine & _
                         "      When Others Then" & vbNewLine & _
                         "        Null;" & vbNewLine & _
                         "    End;" & vbNewLine & _
                         "  End Loop;" & vbNewLine & _
                         "End;" & vbNewLine & _
                         "/"
    If Not rsHisRunafter Is Nothing Then
        '�����ʷ�������ű����ü�¼�ṹΪ
        rsHisRunafter.Sort = "ID"
        Do While Not rsHisRunafter.EOF
            If strCurCon <> rsHisRunafter!BAKUser & "/&" & rsHisRunafter!BAKUser & "_PWD@" & rsHisRunafter!������ Then
                If strCurCon <> "" And intDDLParallel <> 0 Then
                    objTxt.WriteLine strNoDDLParallelSQL
                End If
                If strCurCon <> "" Then objTxt.WriteLine
                strCurCon = rsHisRunafter!BAKUser & "/&" & rsHisRunafter!BAKUser & "_PWD@" & rsHisRunafter!������
                objTxt.WriteLine "Connect " & strCurCon
                If intDDLParallel <> 0 Then
                    objTxt.WriteLine "Alter Session FORCE PARALLEL DDL PARALLEL " & intDDLParallel & ";"
                End If
            End If
            objTxt.WriteLine rsHisRunafter!SQL & ";"
            rsHisRunafter.MoveNext
        Loop
        If strCurCon <> "" And intDDLParallel <> 0 Then
            objTxt.WriteLine strNoDDLParallelSQL
        End If
    End If
    strCurCon = ""
    If Not rsStatistics Is Nothing Then
        rsStatistics.Sort = "ID"
        strCurCon = ""
        objTxt.WriteLine
         '���ͳ����Ϣ�ռ��ű����ü�¼�ṹΪ
        objTxt.WriteLine "--[2.ͳ����Ϣ�ռ�]���ֹ�ִ����ʹ��DBA����SYSTEM���û�ִ�У�"
        objTxt.WriteLine "Connect SYSTEM/&SYSTEM_PASSWORD@" & strServer
        Do While Not rsStatistics.EOF
            If strCurCon <> rsStatistics!Owner Then
                If strCurCon <> "" Then objTxt.WriteLine
                strCurCon = rsStatistics!Owner
            End If
            objTxt.WriteLine rsStatistics!SQL
            rsStatistics.MoveNext
        Loop
    End If
    If Not rsFiles Is Nothing Then
        rsFiles.Sort = "SysType,������,ϵͳ���,FullSPVer"
        If rsFiles.RecordCount <> 0 Then
            strCurCon = ""
            objTxt.WriteLine
             '����ӳ������ű�
            objTxt.WriteLine "--[3.ϵͳ�ӳ������ű�]���ֹ�ִ����ʹ�ð�ִ���û�ִ�У�"
            Do While Not rsFiles.EOF
                If strCurCon <> rsFiles!SysType & "," & rsFiles!������ & "," & rsFiles!ϵͳ��� Then
                    If strCurCon <> "" Then objTxt.WriteLine
                    If rsFiles!������ & "" <> "" Then
                        rsHisDBInfo.Filter = "ϵͳ���=" & rsFiles!ϵͳ��� & " And ����='" & rsFiles!������ & "'"
                        objTxt.WriteLine "Connect " & mrsHistorySpace!������ & "/&" & mrsHistorySpace!������ & "_PASSWORD@" & mrsHistorySpace!������
                    Else
                        If rsFiles!SysType = ST_Tools Then
                            objTxt.WriteLine "Connect ZLTOOLS/&ZLTOOLS_PASSWORD@" & strServer
                        Else
                            objTxt.WriteLine "Connect " & gstrUserName & "/&" & gstrUserName & "_PASSWORD@" & strServer
                        End If
                    End If
                    strCurCon = rsFiles!SysType & "," & rsFiles!������ & "," & rsFiles!ϵͳ���
                End If
                objTxt.WriteLine "@" & rsFiles!FilePath
                rsFiles.MoveNext
            Loop
        End If
    End If
    objTxt.Close
    Set objTxt = Nothing
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Private Sub RebuildRegistFile(ByVal cnTools As ADODB.Connection, ByVal strToolsFloder As String)
    Dim strRegFunFile As String, strSQL As String, strError As String, strRegCheck As String
    Dim rsTmp As ADODB.Recordset
    Dim cnOralce As ADODB.Connection, cnOralceOld As ADODB.Connection, cnCurrent As ADODB.Connection
    
    On Error GoTo ErrHCheck
    strRegCheck = gobjRegister.zlRegCheck(False, True)
    '��׼��������Ϻ���м��ܺ���У�飬�Ƿ���Ҫ�ؽ����ܺ���
    If strRegCheck Like "*�ָ���ȷ��ע�ắ����*" Then
        On Error Resume Next
        Set cnOralceOld = gobjRegister.GetConnection(gstrServer, gstrUserName, gstrPassword, False, MSODBC, strError, False)
        If strError <> "" Then
            mclsRunScript.WriteLog String(17, " ") & "����MSODBC�������ӳ���" & strError
            strError = ""
        End If
        Set cnOralce = gobjRegister.GetConnection(gstrServer, gstrUserName, gstrPassword, False, OraOLEDB, strError, True)
        If strError <> "" Then
            mclsRunScript.WriteLog String(17, " ") & "����OracleOLEDB�������ӳ���" & strError
            strError = ""
        End If
        If cnOralceOld.State = adStateOpen And cnOralce.State = adStateOpen Then
            gcnOracle.Close
            gcnOldOra.Close
            Set gcnOracle = cnOralce
            Set gcnOldOra = cnOralceOld
            Set cnOralceOld = Nothing
            Set cnOralce = Nothing
        End If
        On Error GoTo ErrHCheck
        mclsRunScript.WriteLog String(17, " ") & "���ܺ���У�飺��Ҫ�ؽ�ע����ܺ���"
    Else
        mclsRunScript.WriteLog String(17, " ") & "���ܺ���У�飺����Ҫ�ؽ�ע����ܺ���"
        Exit Sub
    End If
    strRegCheck = ""
    
    On Error GoTo errH
    Set cnCurrent = cnTools
    strRegFunFile = strToolsFloder & "\" & GetRegistFile
    '1.���ע�ắ������ı�ṹ�Ƿ���Ҫ����
    strSQL = "Select Table_Name" & vbNewLine & _
            "From User_Tab_Columns" & vbNewLine & _
            "Where Table_Name In ('ZLREGFILE', 'ZLREGAUDIT') And Column_Name = '��Ŀ' And Data_Length <> 20"
    Set rsTmp = gclsBase.OpenSQLRecord(cnTools, strSQL, "������ݽṹ")
    If rsTmp.RecordCount > 0 Then
        rsTmp.Filter = "Table_Name='ZLREGFILE'"
        If rsTmp.RecordCount > 0 Then
            strSQL = "Alter Table zlRegFile Modify ��Ŀ Varchar2(20)"
            cnTools.Execute strSQL
        End If
        
        rsTmp.Filter = "Table_Name='ZLREGAUDIT'"
        If rsTmp.RecordCount > 0 Then
            strSQL = "Alter Table ZLREGAUDIT Modify ��Ŀ Varchar2(20)"
            cnTools.Execute strSQL
        End If
        
        strSQL = "Drop Type t_Reg_Rowset Force"
        cnTools.Execute strSQL
        strSQL = "Drop Type t_Reg_Record Force"
        cnTools.Execute strSQL
        strSQL = "Create Or Replace Type t_Reg_Record  As Object(Item Varchar2(20), Prog number(18), Text Varchar2(1000))"
        cnTools.Execute strSQL
        strSQL = "Create Or Replace Type t_Reg_Rowset As Table Of t_Reg_Record"
        cnTools.Execute strSQL
                        
        On Error Resume Next
        strSQL = "Grant Execute on t_Reg_Record to Public"
        cnTools.Execute strSQL
        If err.Number <> 0 Then
            mclsRunScript.WriteLog String(17, " ") & "ִ�У�" & strSQL
            mclsRunScript.WriteLog String(17, " ") & "�����" & "ִ�а���Ȩʱʧ�ܣ�����������" & err.Description
        End If
                        
        '����ҽԺ����ZLHIS�İ�T_DB_ROLEUSER��BH�ܺ���صģ������˸ö��󣬵�����Ȩʧ��
        'ORA-04045: �����±���/������֤ ZLHIS.T_DB_ROLEUSER ʱ����
        'ORA -1031: Ȩ�޲���
        strSQL = "Grant Execute on t_Reg_Rowset to Public"
        cnTools.Execute strSQL
        If err.Number <> 0 Then
            mclsRunScript.WriteLog String(17, " ") & "ִ�У�" & strSQL
            mclsRunScript.WriteLog String(17, " ") & "�����" & "ִ�а���Ȩʱʧ�ܣ�����������" & err.Description
            err.Clear
        End If
    End If
    On Error GoTo errH
                              
    If gobjFile.FileExists(strRegFunFile) = False Then
        mclsRunScript.WriteLog String(17, " ") & "δ�ҵ�ע����ܺ����ļ���" & strRegFunFile
        Exit Sub
    End If
    mclsRunScript.WriteLog String(17, " ") & "ִ�У�" & strRegFunFile
    If Not RunRegistFile(Me, cnTools, gstrToolsPwd, gstrServer, strRegFunFile) Then
        mclsRunScript.WriteLog String(17, " ") & "�����ִ��ʧ��"
    Else
        mclsRunScript.WriteLog String(17, " ") & "�����ִ�гɹ�"
    End If
    If gobjFile.FileExists(lblRegist.Tag) = False Then
        Exit Sub
    End If
    
    '�л����ܺ�����֤��ʽ
    Set cnCurrent = gcnOracle
    Call gobjRegister.zlRegInit(gcnOracle)
    If gobjRegister.zlRegBuild(lblRegist.Tag, prgThis) = False Then
        Exit Sub
    End If
    strRegCheck = gobjRegister.zlRegCheck(True)
    If strRegCheck = "" Then
        gcnOracle.Execute "call zltools.p_Reg_Apply()", , adCmdText
        strRegCheck = gobjRegister.zlRegCheck(False) '�ٴε�����֤
        If strRegCheck = "" Then
            mclsRunScript.WriteLog String(17, " ") & "ע����Ȩ��Ϣ�Ѿ�Ӧ��"
        Else
            mclsRunScript.WriteLog String(17, " ") & strRegCheck & ",����zlRegAudit��zlRegFile���[��Ŀ]�ֶγ��ȣ�����ϵ����пͻ��˲��������ṹ��"
        End If
        Set cnCurrent = gcnOldOra
        Call gobjRegister.zlRegInit(gcnOldOra)
        Call gobjRegister.zlRegCheck(False)
    Else
        mclsRunScript.WriteLog String(17, " ") & "ע����Ϣ����ȷ��������ע��:" & strRegCheck
    End If
    
    Exit Sub
    
errH:
    If ADOConnectionError(err, cnCurrent) Then
        If CheckAdoConnection(cnTools) Then Resume
    End If
    mclsRunScript.WriteLog String(17, " ") & "����" & err.Description
    mclsRunScript.WriteLog String(17, " ") & "�����ִ��ʧ��"
    Exit Sub
    
ErrHCheck:
    mclsRunScript.WriteLog String(17, " ") & "���ܺ���У�����" & err.Description
End Sub

'--------------------------------------------------------------------------------------------------
'����           DoHelperMain
'����           �������ָ���������Ҫ������Ȩ��
'����ֵ
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Private Sub DoHelperMain()
    Dim cnTools     As ADODB.Connection
    Dim rsTmp       As ADODB.Recordset
    Dim strError    As String
    Dim strSQL      As String
    Dim strTmp      As String
    Dim lngJobNum   As Long
        
    '������ʱ��system�����ӣ����������߼�������ֵ����������봰��
    On Error GoTo errH
    Set cnTools = GetConnection("ZLTOOLS")
    If Not cnTools Is Nothing Then
        '��ɫ�жϴ���
        strSQL = "Select 1 From Dba_Roles Where Role =[1]"
        Set rsTmp = gclsBase.OpenSQLRecord(cnTools, strSQL, Me.Caption, "ZL_������֤")
        If rsTmp.RecordCount > 0 Then
            strError = gclsBase.ExecuteCmdText("Drop Role ZL_������֤", Me.Caption, cnTools, True)
            If strError <> "" Then
                mclsRunScript.WriteLog String(17, " ") & "Drop Role ZL_������֤,����" & strError
            End If
        End If
        strError = gclsBase.ExecuteCmdText("Delete ZLRoleGrant Where ��ɫ='ZL_������֤'", Me.Caption, cnTools, True)
        If strError <> "" Then
            mclsRunScript.WriteLog String(17, " ") & "Delete ZLRoleGrant Where ��ɫ='ZL_������֤',����" & strError
        End If
        strError = gclsBase.ExecuteCmdText("Delete ZLRoles Where ����='ZL_������֤'", Me.Caption, cnTools, True)
        If strError <> "" Then
            mclsRunScript.WriteLog String(17, " ") & "Delete ZLRoles Where ����='ZL_������֤',����" & strError
        End If
        strError = gclsBase.ExecuteCmdText("Delete zluserroles Where ��ɫ='ZL_������֤'", Me.Caption, cnTools, True)
        If strError <> "" Then
            mclsRunScript.WriteLog String(17, " ") & "Delete zluserroles Where ��ɫ='ZL_������֤',����" & strError
        End If
        
        strError = gclsBase.ExecuteCmdText("Create Role ZL_������֤ Not Identified", Me.Caption, cnTools, True)
        If strError <> "" Then
            mclsRunScript.WriteLog String(17, " ") & "Create Role ZL_������֤ Not Identified,����" & strError
        End If
        strError = gclsBase.ExecuteCmdText("Insert Into Zlroles(����, ϵͳ) values( 'ZL_������֤',NULL)", Me.Caption, cnTools, True)
        If strError <> "" Then
            mclsRunScript.WriteLog String(17, " ") & "Insert Into Zlroles(����, ϵͳ) values( 'ZL_������֤',NULL),����" & strError
        End If
        
        '������ʷ��֤��Ϣ
        strError = gclsBase.ExecuteCmdText("Delete From Zlclientvertify", Me.Caption, cnTools, True)
        If strError <> "" Then
            mclsRunScript.WriteLog String(17, " ") & "Delete From Zlclientvertify,����" & strError
        End If
        strError = gclsBase.ExecuteCmdText("Grant ZL_������֤ To ZLUA", Me.Caption, cnTools, True)
        If strError <> "" Then
            mclsRunScript.WriteLog String(17, " ") & "Grant ZL_������֤ To ZLUA,����" & strError
        End If
        strError = gclsBase.ExecuteCmdText("insert into  zluserroles(�û�, ��ɫ, ����) values('ZLUA','ZL_������֤',0)", Me.Caption, cnTools, True)
        If strError <> "" Then
            mclsRunScript.WriteLog String(17, " ") & "insert into  zluserroles(�û�, ��ɫ, ����) values('ZLUA','ZL_������֤',0),����" & strError
        End If
        '��Ա��Ӧ�û�����
        strSQL = "Select Distinct ������ From Zlsystems"
        Set rsTmp = gclsBase.OpenSQLRecord(cnTools, strSQL, "��ȡ������")
        Do While Not rsTmp.EOF
            strError = gclsBase.ExecuteCmdText("Delete " & rsTmp!������ & ".�ϻ���Ա�� Where �û���='ZLUA'", Me.Caption, gcnOracle, True)
            If strError <> "" Then
                mclsRunScript.WriteLog String(17, " ") & "Delete " & rsTmp!������ & ".�ϻ���Ա�� Where �û���='ZLUA',����" & strError
            End If
            strError = gclsBase.ExecuteCmdText("Insert Into " & rsTmp!������ & ".�ϻ���Ա�� (�û���, ��Աid)Select 'ZLUA', ��Աid From " & rsTmp!������ & ".�ϻ���Ա�� b Where b.�û���='" & rsTmp!������ & "'", Me.Caption, gcnOracle, True)
            If strError <> "" Then
                mclsRunScript.WriteLog String(17, " ") & "Insert Into " & rsTmp!������ & ".�ϻ���Ա�� (�û���, ��Աid)Select 'ZLUA', ��Աid From " & rsTmp!������ & ".�ϻ���Ա�� b Where b.�û���='" & rsTmp!������ & "',����" & strError
            End If
            rsTmp.MoveNext
        Loop
        
        '����ɫ[ZL_������֤]��������ģ���Ȩ��
        strSQL = "Select Nvl(c.ϵͳ,0) ϵͳ, c.���, c.����" & vbNewLine & _
                "From (Select f.ϵͳ, f.���, f.����" & vbNewLine & _
                "       From zlProgFuncs F, zlRegFunc R" & vbNewLine & _
                "       Where Trunc(f.ϵͳ / 100) = r.ϵͳ And f.��� = r.��� And f.���� = r.����" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select ϵͳ, ���, ����" & vbNewLine & _
                "       From zlProgFuncs" & vbNewLine & _
                "       Where ϵͳ Is Null Or (��� Between 10000 And 19999)" & vbNewLine & _
                "       Union" & vbNewLine & _
                "       Select a.ϵͳ, a.����id As ���, a.����" & vbNewLine & _
                "       From zlReports B, zlRPTPuts A" & vbNewLine & _
                "       Where a.����id = b.Id) C, (Select b.ϵͳ, b.��� From zlPrograms B Where Nvl(b.����, 'NONEDATA') <> 'ZLREPORT') D" & vbNewLine & _
                "Where c.ϵͳ = d.ϵͳ And c.��� = d.���" & vbNewLine & _
                "Order By c.ϵͳ, c.���, c.����"
        Set rsTmp = gclsBase.OpenSQLRecord(cnTools, strSQL, "��ȡ����ģ��Ȩ��")
        Do While Not rsTmp.EOF
            If strTmp <> "" Then strTmp = strTmp & "''"
            strTmp = strTmp & IIf(rsTmp!ϵͳ = 0, "null", rsTmp!ϵͳ) & "''" & rsTmp!���.value & "''" & rsTmp!����.value
            If ActualLen(strTmp) > 2000 Then
                Call ExecuteProcedure("zl_zlRoleGrant_BatchInsert('ZL_������֤','" & strTmp & "')", "��Ȩ", cnTools)
                strTmp = ""
            End If
            rsTmp.MoveNext
        Loop
        If strTmp <> "" Then
            Call ExecuteProcedure("zl_zlRoleGrant_BatchInsert('ZL_������֤','" & strTmp & "')", "��Ȩ", cnTools)
        End If
        '����һ�����Զ���ҵ
        On Error Resume Next
        strError = ""
        Call ExecuteProcedure("zl_JobRemove(Null,1,2)", Me.Caption, cnTools)
        If err.Number <> 0 Then
            mclsRunScript.WriteLog String(17, " ") & "Call zl_JobRemove(Null,1,1),����" & err.Description
            err.Clear
        End If
        '����ִ��ʱ��
        strError = gclsBase.ExecuteCmdText("update zltools.zlautojobs a set a.ִ��ʱ��=sysdate where a.ϵͳ is null and a.���� = 1 And ���=2", Me.Caption, cnTools, True)
        If strError <> "" Then
            mclsRunScript.WriteLog String(17, " ") & "update zltools.zlautojobs a set a.ִ��ʱ��=sysdate where a.ϵͳ is null and a.���� = 1 And ���=2,����" & strError
        End If
        Call ExecuteProcedure("zl_JobSubmit(Null,1,2)", Me.Caption, cnTools)
        If err.Number <> 0 Then
            mclsRunScript.WriteLog String(17, " ") & "Call zl_JobSubmit(Null,1,1),����" & err.Description
            err.Clear
        End If
    End If
    Exit Sub
    
errH:
    If ADOConnectionError(err, cnTools) Then
        If CheckAdoConnection(cnTools) Then Resume
    End If
    mclsRunScript.WriteLog String(17, " ") & "����ֵ����������ʧ��,����" & err.Description
End Sub

Public Sub CheckToolsLob(Optional ByVal blnOnlyToolsUp As Boolean, Optional ByVal strCurToolsVersion As String, Optional ByVal strToolsVersion As String, Optional ByVal strHisVersion As String)
'���ܣ����zlRPTGraphs.ͼƬ�����Լ����������׼��İ汾��
'blnOnlyToolsUp:�Ƿ�ֻ�й�������������ʱֻ���ݹ����߰汾�š�ZLHIS�汾�Ŵ����ݿ��ȡ
'strCurToolsVersion:��ǰ�����߰汾
'strToolsVersion�������߰汾���ǹ���������ʱ�������ݡ�
'strHisVersion :ZLHIS�汾�ţ��ǹ���������ʱ�������ݡ�
    Dim rsTmp       As ADODB.Recordset
    Dim strSQL      As String
    
    On Error Resume Next
    If blnOnlyToolsUp Then
        Set rsTmp = gclsBase.GetSystems(100)
        If Not rsTmp.EOF Then
            strHisVersion = rsTmp!�汾�� & ""
        End If
    Else
        mrsSysInfo.Filter = "ϵͳ���=100"
        If Not mrsSysInfo.EOF Then
            strHisVersion = mrsSysInfo!Ŀ��汾 & ""
            If strHisVersion = "" Then
                strHisVersion = mrsSysInfo!ϵͳ�汾�� & ""
            End If
        End If
        mrsSysInfo.Filter = "ϵͳ���=0"
        If Not mrsSysInfo.EOF Then
            strToolsVersion = mrsSysInfo!Ŀ��汾 & ""
            strCurToolsVersion = strToolsVersion
            If strToolsVersion = "" Then
                strToolsVersion = mrsSysInfo!ϵͳ�汾�� & ""
            End If
        End If
    End If
    If VerFull(strToolsVersion, True) >= VerFull("10.35.90") Then
        mintToolLob = mintToolLob Or LC_ZLTOOLS_IS3590_OR_GREATER
    End If
    
    If VerFull(strHisVersion, True) >= VerFull("10.35.90") Then
        mintToolLob = mintToolLob Or LC_ZLHIS_IS3590_OR_GREATER
    End If
    
    If VerFull(strCurToolsVersion, True) >= VerFull("10.35.90") Then
        mintToolLob = mintToolLob Or LC_ZLTOOLS_CURIS3590_OR_GREATER
    End If
    
    '��ȡ�ֶ�����
    strSQL = "Select 1" & vbNewLine & _
            "From All_Tab_Columns" & vbNewLine & _
            "Where Owner = 'ZLTOOLS' And Table_Name = 'ZLRPTGRAPHS' And Column_Name = 'ͼƬ' And Data_Type = 'LONG RAW'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "CheckToolsLob")
    If Not rsTmp.EOF Then
        mintToolLob = mintToolLob Or LC_ISLONGRAW
    End If
End Sub

Private Sub AdjustToolLob()
'���ܣ�����ZLTOOLS.ZLRPTGRAPHS.ͼƬ�ֶ�ΪLob��ֻ���ڱ�׼��10.35.90���������������
    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset
    Dim strError    As String
    Dim cnTools     As ADODB.Connection
    
    On Error GoTo errH
    mclsRunScript.WriteLog String(17, " ") & "ִ�м�������Alter Table zltools.zlRPTGraphs Modify ͼƬ Blob"
    Set cnTools = GetConnection("ZLTOOLS")
    If Not cnTools Is Nothing Then
        Call SetSessionParallel(cnTools, False)
        strError = gclsBase.ExecuteCmdText("Alter Table zltools.zlRPTGraphs Modify ͼƬ Blob", Me.Caption, cnTools, True)
        If strError <> "" Then
            mclsRunScript.WriteLog String(17, " ") & "Alter Table zltools.zlRPTGraphs Modify ͼƬ Blob,����" & strError
        End If
        strSQL = "Select 'alter index ' || Index_Name || ' rebuild' As ����" & vbNewLine & _
                "From User_Indexes" & vbNewLine & _
                "Where Table_Name = 'ZLRPTGRAPHS' And Status = 'UNUSABLE'"
        Set rsTmp = gclsBase.OpenSQLRecord(cnTools, strSQL, "AdjustToolLob")
        Do While Not rsTmp.EOF
            strError = gclsBase.ExecuteCmdText(rsTmp!���� & "", Me.Caption, cnTools, True)
            If strError <> "" Then
                mclsRunScript.WriteLog String(17, " ") & rsTmp!���� & "����" & strError
            End If
            rsTmp.MoveNext
        Loop
    End If
    Exit Sub
    
errH:
    If ADOConnectionError(err, cnTools) Then
        If CheckAdoConnection(cnTools) Then Resume
    End If
    mclsRunScript.WriteLog String(17, " ") & "������Alter Table zltools.zlRPTGraphs Modify ͼƬ Blob����ʧ��,����" & err.Description
End Sub
