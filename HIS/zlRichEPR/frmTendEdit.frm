VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ŀ�༭"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   Icon            =   "frmTendEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picTemperaPart 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   675
      ScaleHeight     =   1950
      ScaleWidth      =   2595
      TabIndex        =   69
      Top             =   3090
      Width           =   2630
      Begin VB.ComboBox Cbo 
         Height          =   300
         Index           =   12
         Left            =   810
         TabIndex        =   32
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox Cbo 
         Height          =   300
         Index           =   3
         Left            =   810
         TabIndex        =   29
         Top             =   90
         Width           =   1815
      End
      Begin VB.ComboBox Cbo 
         Height          =   300
         Index           =   6
         Left            =   810
         TabIndex        =   30
         Top             =   465
         Width           =   1815
      End
      Begin VB.ComboBox Cbo 
         Height          =   300
         Index           =   7
         Left            =   810
         TabIndex        =   31
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox Cbo 
         Height          =   300
         Index           =   13
         Left            =   810
         TabIndex        =   33
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�� ��"
         Height          =   180
         Index           =   26
         Left            =   150
         TabIndex        =   74
         Top             =   1665
         Width           =   450
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�� ��"
         Height          =   180
         Index           =   25
         Left            =   150
         TabIndex        =   73
         Top             =   1305
         Width           =   450
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�� ��"
         Height          =   180
         Index           =   24
         Left            =   150
         TabIndex        =   72
         Top             =   930
         Width           =   450
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Ҹ ��"
         Height          =   180
         Index           =   23
         Left            =   150
         TabIndex        =   71
         Top             =   570
         Width           =   450
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�� ��"
         Height          =   180
         Index           =   22
         Left            =   135
         TabIndex        =   70
         Top             =   180
         Width           =   450
      End
   End
   Begin VB.Frame fra 
      Caption         =   "���¼�¼ɫ"
      Height          =   3520
      Index           =   1
      Left            =   3375
      TabIndex        =   58
      Top             =   4845
      Width           =   3975
      Begin VB.PictureBox picBack 
         BorderStyle     =   0  'None
         Height          =   2160
         Index           =   0
         Left            =   1665
         ScaleHeight     =   2160
         ScaleWidth      =   2280
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   150
         Width           =   2280
         Begin VB.CommandButton cmd 
            Caption         =   "�Զ�����ɫ(&M)��"
            Height          =   350
            Index           =   1
            Left            =   15
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   1770
            Width           =   2205
         End
         Begin zlRichEPR.ColorPicker usrColor 
            Height          =   2190
            Left            =   30
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   -450
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   3863
         End
      End
      Begin VB.PictureBox picDemo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2010
         Left            =   75
         ScaleHeight     =   1980
         ScaleWidth      =   1530
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   255
         Width           =   1560
      End
   End
   Begin VB.Frame fra 
      Caption         =   "��������"
      Height          =   4785
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   15
      Width           =   7260
      Begin VB.ComboBox Cbo 
         Height          =   300
         Index           =   11
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   4140
         Width           =   1815
      End
      Begin VB.TextBox txtInfo 
         Height          =   1950
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   2745
         Width           =   3810
      End
      Begin VB.ComboBox Cbo 
         Height          =   300
         Index           =   10
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3795
         Width           =   1815
      End
      Begin VB.ComboBox Cbo 
         Height          =   300
         Index           =   9
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3465
         Width           =   1815
      End
      Begin VB.ComboBox Cbo 
         Height          =   300
         Index           =   8
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CheckBox chk 
         Caption         =   "������Ŀ(&A)"
         Height          =   240
         Index           =   1
         Left            =   1380
         TabIndex        =   15
         Top             =   4485
         Width           =   1290
      End
      Begin VB.ComboBox Cbo 
         Height          =   300
         Index           =   5
         Left            =   1380
         TabIndex        =   1
         Top             =   285
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   4
         Top             =   975
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   1380
         TabIndex        =   2
         Top             =   630
         Width           =   1530
      End
      Begin VB.CommandButton cmd 
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2925
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   630
         Width           =   270
      End
      Begin VB.CheckBox chk 
         Caption         =   "����Ҫ��"
         Height          =   270
         Index           =   0
         Left            =   270
         TabIndex        =   37
         Top             =   660
         Width           =   1020
      End
      Begin VB.ComboBox Cbo 
         Height          =   300
         Index           =   1
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   7
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   8
         Top             =   2055
         Width           =   1815
      End
      Begin VB.ComboBox Cbo 
         Height          =   300
         Index           =   2
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2400
         Width           =   1815
      End
      Begin VB.ComboBox Cbo 
         Height          =   300
         Index           =   4
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2760
         Width           =   1815
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Index           =   0
         Left            =   1770
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txt(2)"
         BuddyDispid     =   196619
         BuddyIndex      =   2
         OrigLeft        =   3555
         OrigTop         =   1860
         OrigRight       =   3795
         OrigBottom      =   2160
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Index           =   1
         Left            =   2940
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txt(5)"
         BuddyDispid     =   196619
         BuddyIndex      =   5
         OrigLeft        =   3555
         OrigTop         =   2250
         OrigRight       =   3795
         OrigBottom      =   2550
         Max             =   2
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin zlRichEPR.VsfGrid vsf 
         Height          =   1815
         Left            =   3360
         TabIndex        =   16
         Top             =   570
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   6615
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   5
         Left            =   2670
         MaxLength       =   1
         TabIndex        =   7
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   6
         Top             =   1680
         Width           =   375
      End
      Begin VB.PictureBox picBack 
         BorderStyle     =   0  'None
         Height          =   1860
         Index           =   1
         Left            =   3360
         ScaleHeight     =   1860
         ScaleWidth      =   3810
         TabIndex        =   63
         Top             =   525
         Width           =   3810
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ӧ�ó���"
         Height          =   180
         Index           =   21
         Left            =   570
         TabIndex        =   68
         Top             =   4200
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "˵��(&I)"
         Height          =   180
         Index           =   19
         Left            =   3360
         TabIndex        =   17
         Top             =   2460
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ����"
         Height          =   180
         Index           =   6
         Left            =   570
         TabIndex        =   49
         Top             =   3855
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���ò���"
         Height          =   180
         Index           =   5
         Left            =   570
         TabIndex        =   48
         Top             =   3525
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ӧ�÷�ʽ"
         Height          =   180
         Index           =   4
         Left            =   570
         TabIndex        =   47
         Top             =   3180
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   15
         Left            =   570
         TabIndex        =   36
         Top             =   345
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ֵ��(&Z)"
         Height          =   180
         Index           =   20
         Left            =   3360
         TabIndex        =   50
         Top             =   315
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��ͻ���"
         Height          =   180
         Index           =   14
         Left            =   570
         TabIndex        =   46
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ��ʾ"
         Height          =   180
         Index           =   12
         Left            =   570
         TabIndex        =   45
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ��λ"
         Height          =   180
         Index           =   11
         Left            =   570
         TabIndex        =   44
         Top             =   2115
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "С��"
         Height          =   180
         Index           =   10
         Left            =   2310
         TabIndex        =   42
         Top             =   1740
         Width           =   360
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ����"
         Height          =   180
         Index           =   9
         Left            =   570
         TabIndex        =   40
         Top             =   1740
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ����"
         Height          =   180
         Index           =   7
         Left            =   570
         TabIndex        =   39
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ����"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   38
         Top             =   1035
         Width           =   720
      End
   End
   Begin VB.Frame fra 
      Caption         =   "��������"
      Height          =   3520
      Index           =   2
      Left            =   90
      TabIndex        =   19
      Top             =   4845
      Width           =   3285
      Begin VB.CommandButton cmdTemperature 
         Height          =   260
         Left            =   2895
         Picture         =   "frmTendEdit.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   660
         Width           =   270
      End
      Begin VB.TextBox txtTemper 
         Height          =   300
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   645
         Width           =   1815
      End
      Begin VB.CheckBox chkFirst 
         Caption         =   "��Ժ�ײ�"
         Height          =   240
         Left            =   1080
         TabIndex        =   28
         Top             =   3225
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   8
         Left            =   1380
         MaxLength       =   5
         TabIndex        =   27
         Text            =   "10"
         Top             =   2850
         Width           =   1800
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   25
         Text            =   "2"
         Top             =   2115
         Width           =   1500
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   1380
         MaxLength       =   5
         TabIndex        =   26
         Text            =   "10"
         Top             =   2490
         Width           =   1800
      End
      Begin VB.ComboBox Cbo 
         Height          =   300
         Index           =   0
         ItemData        =   "frmTendEdit.frx":0596
         Left            =   1380
         List            =   "frmTendEdit.frx":0598
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   255
         Width           =   1815
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Index           =   5
         Left            =   2910
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1755
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txt(12)"
         BuddyDispid     =   196619
         BuddyIndex      =   12
         OrigLeft        =   3345
         OrigTop         =   1650
         OrigRight       =   3585
         OrigBottom      =   1950
         Max             =   300
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Index           =   6
         Left            =   2895
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1020
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txt(13)"
         BuddyDispid     =   196619
         BuddyIndex      =   13
         OrigLeft        =   3345
         OrigTop         =   1650
         OrigRight       =   3585
         OrigBottom      =   1950
         Max             =   300
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   13
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   22
         Text            =   "1"
         Top             =   1020
         Width           =   1500
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Index           =   2
         Left            =   2910
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   2115
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txt(3)"
         BuddyDispid     =   196619
         BuddyIndex      =   3
         OrigLeft        =   3345
         OrigTop         =   1650
         OrigRight       =   3585
         OrigBottom      =   1950
         Max             =   300
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1380
         MaxLength       =   5
         TabIndex        =   23
         Text            =   "10"
         Top             =   1395
         Width           =   1800
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   12
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "5"
         Top             =   1755
         Width           =   1500
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��ʾ��"
         Height          =   180
         Index           =   13
         Left            =   735
         TabIndex        =   67
         Top             =   2910
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�̶ȼ��"
         Height          =   180
         Index           =   8
         Left            =   555
         TabIndex        =   66
         Top             =   2550
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��¼Ƶ��"
         Height          =   180
         Index           =   0
         Left            =   555
         TabIndex        =   64
         Top             =   2175
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��λֵ"
         Height          =   180
         Index           =   2
         Left            =   735
         TabIndex        =   55
         Top             =   1455
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�����"
         Height          =   180
         Index           =   17
         Left            =   735
         TabIndex        =   56
         Top             =   1815
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���к�"
         Height          =   180
         Index           =   18
         Left            =   735
         TabIndex        =   53
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��¼��"
         Height          =   180
         Index           =   16
         Left            =   735
         TabIndex        =   52
         Top             =   705
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��¼��"
         Height          =   180
         Index           =   3
         Left            =   735
         TabIndex        =   51
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FF0000&
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7485
      TabIndex        =   34
      Top             =   450
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7485
      TabIndex        =   35
      Top             =   870
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   7860
      Top             =   3450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   7995
      Top             =   2115
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendEdit.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendEdit.frx":06F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendEdit.frx":084E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendEdit.frx":09A8
            Key             =   "User"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTendEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'######################################################################################################################
'�ֲ�������������

Private Type Items
    Ҫ����Ŀ As String
End Type

Private usrSaveItem As Items

Private mblnStartUp As Boolean
Private mblnOk As Boolean
Private mblnDataChanged As Boolean
Private mlngKey As Long
Private mlngPKey As Long
Private mbln������Ŀ As Boolean
Private mintType As Integer  '������Ŀ 0-�ǲ�����Ŀ 1-�������� 2-���̱��
Private mfrmMain As Form
Private mstrSQL As String


'######################################################################################################################
'�Զ��庯������������
Private Property Let DataChanged(ByVal vData As Boolean)
    mblnDataChanged = vData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Sub SetGridFormat()
    Dim strTmp As String
    Dim strTmp2 As String
    Dim lngCol As Long
    
    If Left(cbo(2).Text, 1) = 0 Then
        lngCol = 2
    Else
        lngCol = 1
    End If

    vsf.Body.ColFormat(1) = ""
    vsf.Body.ColEditMask(1) = ""

    vsf.Body.ColFormat(2) = ""
    vsf.Body.ColEditMask(2) = ""
        
    If cbo(1).ListIndex = 0 And (Val(txt(2).Text) - Val(txt(5).Text) - 1) > 0 Then

        strTmp = String(Val(txt(2).Text) - Val(txt(5).Text) - 1, "#")
        strTmp2 = strTmp & "0"
        strTmp = strTmp & "#"

        If Val(txt(5).Text) > 0 Then
            strTmp = strTmp & "." & String(Val(txt(5).Text), "0")
            strTmp2 = strTmp2 & "." & String(Val(txt(5).Text), "0")
            vsf.Body.ColFormat(lngCol) = strTmp2
'            vsf.Body.ColEditMask(lngCol) = strTmp
        Else
            vsf.Body.ColFormat(lngCol) = strTmp2
        End If
    End If
End Sub

Public Function ShowEdit(ByVal frmMain As Form, ByVal lngKey As Long, Optional ByVal lngPKey As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��/��ʾ�༭����,���������������(��ں���)
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOk = False
    
    Set mfrmMain = frmMain
    mlngKey = lngKey
    mlngPKey = lngPKey
    
    If InitData = False Then GoTo errHand
    If ReadData = False Then GoTo errHand
    
    usrSaveItem.Ҫ����Ŀ = txt(6).Text
    txt(6).Tag = ""
    mblnStartUp = False
    
    Call cbo_Click(0)
    Call cbo_Click(1)
    Call chk_Click(0)
    Call chk_Click(1)
    Call SetLabelEnable
    
    
    Call cbo_Click(1)
    'If lngKey = 1 Or lngKey = 2 Then vsf.Body.Editable = flexEDNone
    If cbo(0).ListIndex > 0 Then
        vsf.Body.Editable = flexEDKbdMouse
    End If
'    If lngKey = 7 Or lngKey = 9 Then txt(0).Enabled = True
    DataChanged = False
    Me.Show 1, frmMain
    
    ShowEdit = mblnOk
    
    Exit Function
    
errHand:
    On Error Resume Next
    DataChanged = False
    Unload Me
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������ϣ��Թ���ʾ
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim varTmp As Variant
    Dim lngLoop As Long
    Dim lngKey As Long
    Dim arrTmp() As String
    Dim i As Integer
    
    On Error GoTo errHand
    
    If mlngKey <> 0 Then
        lngKey = mlngKey
    Else
        lngKey = mlngPKey
    End If
    
    If lngKey <> 0 Then
        
        mstrSQL = "Select A.*,B.����,B.������,C.* From �����¼��Ŀ A,����������Ŀ B,���¼�¼��Ŀ C Where C.��Ŀ���(+)=A.��Ŀ��� And A.��Ŀid=B.ID(+) AND A.��Ŀ���=[1]"
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
            '�������ǿ��������Ϊ�̶���Ŀ
            If mlngKey <> 0 Then mbln������Ŀ = (zlCommFun.NVL(rs("������Ŀ"), 0) = 1 Or mlngKey = 1)
            
            If Val(zlCommFun.NVL(rs("��Ŀid"))) > 0 Then
                chk(0).Value = 1
                txt(6).Text = zlCommFun.NVL(rs("������"))
                cmd(0).Tag = zlCommFun.NVL(rs("��Ŀid"))
            End If
            
            txt(0).Text = zlCommFun.NVL(rs("��Ŀ����"))
            txt(2).Text = zlCommFun.NVL(rs("��Ŀ����"))
            txt(5).Text = zlCommFun.NVL(rs("��ĿС��"), "")
            txt(7).Text = zlCommFun.NVL(rs("��Ŀ��λ"))
            
            cbo(5).Text = zlCommFun.NVL(rs("������"))
            
            On Error Resume Next
            cbo(4).ListIndex = zlCommFun.NVL(rs("����ȼ�"), 0)
            cbo(8).ListIndex = zlCommFun.NVL(rs("Ӧ�÷�ʽ"), 1)
            cbo(9).ListIndex = zlCommFun.NVL(rs("���ò���"), 1)
            
            Select Case zlCommFun.NVL(rs("��Ŀ��ʾ"), 0)
            Case 0, 1
                cbo(2).ListIndex = 0
            Case 2
                cbo(2).ListIndex = 1
            Case 3
                cbo(2).ListIndex = 2
            Case 4
                cbo(2).ListIndex = 3
            Case 5
                cbo(2).ListIndex = 4
            End Select
            On Error GoTo errHand
            
            cbo(1).ListIndex = zlCommFun.NVL(rs("��Ŀ����"), 0)
            Call zlControl.CboLocate(cbo(10), zlCommFun.NVL(rs("��Ŀ����"), 1), True)
            If zlCommFun.NVL(rs("��Ŀ����"), "1") = "2" Then
                Call zlControl.CboLocate(cbo(11), zlCommFun.NVL(rs("Ӧ�ó���"), 0), True)
                mblnStartUp = False
                Call cbo_Click(11)
                mblnStartUp = True
            End If
            strTmp = zlCommFun.NVL(rs("��Ŀֵ��"))
            
            Call InitGrid
            
            If strTmp <> "" Then
                varTmp = Split(strTmp, ";")
                
                If Val(vsf.Tag) = 1 Then
                    vsf.TextMatrix(1, 2) = varTmp(0)
                    If UBound(varTmp) >= 1 Then vsf.TextMatrix(2, 2) = varTmp(1)
                    If UBound(varTmp) >= 2 Then vsf.TextMatrix(3, 2) = varTmp(2)
                    vsf.TextMatrix(4, 2) = Val(NVL(rs("��Сֵ"))) & ";" & Val(NVL(rs("���ֵ")))
                    
                    If InStr(1, zlCommFun.NVL(rs("�ٽ�ֵ")), ";") = 0 Then
                        If IsNumeric(zlCommFun.NVL(rs("�ٽ�ֵ"))) = True Then
                            vsf.TextMatrix(5, 2) = Val(NVL(rs("��Сֵ"))) & ";" & Val(zlCommFun.NVL(rs("�ٽ�ֵ")))
                        Else
                            vsf.TextMatrix(5, 2) = ""
                        End If
                    Else
                        varTmp = Split(zlCommFun.NVL(rs("�ٽ�ֵ")), ";")
                        vsf.TextMatrix(5, 2) = Val(varTmp(0)) & ";" & Val(varTmp(1))
                    End If
                Else
                    For lngLoop = 0 To UBound(varTmp)
                        If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then vsf.Rows = vsf.Rows + 1
                        vsf.RowData(vsf.Rows - 1) = 1
                        If NVL(rs("ȱʡֵ")) <> "" Then
                            arrTmp = Split(NVL(rs("ȱʡֵ")), ";")
                            For i = 0 To UBound(arrTmp)
                                If varTmp(lngLoop) = arrTmp(i) Then
                                    vsf.TextMatrix(vsf.Rows - 1, 1) = varTmp(lngLoop)
                                    vsf.TextMatrix(vsf.Rows - 1, 2) = "1"
                                Else
                                    vsf.TextMatrix(vsf.Rows - 1, 1) = varTmp(lngLoop)
                                End If
                            Next i
                        Else
                            vsf.TextMatrix(vsf.Rows - 1, 1) = varTmp(lngLoop)
                        End If
                        
                    Next
                End If
                
            End If
            txtInfo.Text = Trim(NVL(rs("˵��")))
            If zlCommFun.NVL(rs("��¼��")) <> "" Then
                chk(1).Value = 1
                cbo(0).ListIndex = (zlCommFun.NVL(rs("��¼��"), 1) - 1)
                
'                If cbo(0).ListIndex = 0 Then udn(2).Max = 6
'                If cbo(0).ListIndex = 1 Then udn(2).Max = 2
                
                txt(1).Text = Format(zlCommFun.NVL(rs("��λֵ")), "0.0")
                txt(12).Text = zlCommFun.NVL(rs("�����"))
                txt(13).Text = zlCommFun.NVL(rs("�������"))
                txt(3).Text = zlCommFun.NVL(rs("��¼Ƶ��"), 2)
                txt(4).Text = zlCommFun.NVL(rs!�̶ȼ��)
                txt(8).Text = zlCommFun.NVL(rs!��ʾ��)
                chkFirst.Value = Val(zlCommFun.NVL(rs!��Ժ�ײ�, 0))
                
                strTmp = ""
                If mlngKey = 1 Then
                    strTmp = zlCommFun.NVL(rs("��¼��").Value, "��,��,��,��,��")
                    cbo(3).Text = Split(strTmp, ",")(0)
                    cbo(6).Text = Split(strTmp, ",")(1)
                    cbo(7).Text = Split(strTmp, ",")(2)
                    If UBound(Split(strTmp, ",")) > 2 Then cbo(12).Text = Split(strTmp, ",")(3)
                    If UBound(Split(strTmp, ",")) > 3 Then cbo(13).Text = Split(strTmp, ",")(4)
                    txtTemper.Text = strTmp
                Else
                    cbo(3).Text = zlCommFun.NVL(rs("��¼��").Value)
                End If
                
                picDemo.Tag = zlCommFun.NVL(rs("��¼ɫ"), 0)
                Call DrawDemo(picDemo, cbo(0).ListIndex, Val(picDemo.Tag))
            End If
            
            If mbln������Ŀ = True And InStr(1, ",��������,��¶�ߵ�,����,����,", "," & txt(0).Text & ",") <> 0 Then
                chk(1).Caption = "������Ŀ"
                If InStr(1, ",��������,��¶�ߵ�,", "," & txt(0).Text & ",") <> 0 Then
                    mintType = 1
                Else
                    mintType = 2
                End If
            End If
        End If
        
        If mlngKey = 0 Then
            
            txt(13).Text = GetMaxNo(2)
            txt(0).Text = ""
            cmd(0).Tag = ""
            txt(6).Text = ""
            txt(6).Tag = ""
            txtInfo.Text = ""
        End If
    End If
    
    
    ReadData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SetLabelEnable() As Boolean
    
'    chk(1).Enabled = Not mbln������Ŀ And (mlngKey = 6 Or mlngKey = 8)
    
    If mbln������Ŀ Then
        If mlngKey <> 6 And mlngKey <> 8 Then chk(1).Enabled = False
        txt(0).Enabled = False
        cbo(1).Enabled = False
        cbo(2).Enabled = False
        txt(7).Enabled = False
        txt(2).Enabled = False
        txt(5).Enabled = False
        txt(1).Enabled = False
'       txt(12).Enabled = (mlngKey <> 1 And mlngKey <> -1 And cbo(0).ListIndex = 0)
        cbo(10).Enabled = False
        cbo(11).Enabled = False
        cbo(0).Enabled = False
        If mintType = 1 Or mintType = 2 Then
            cbo(4).Enabled = False
            cbo(8).Enabled = False
            cbo(9).Enabled = False
            vsf.Body.Enabled = False
        End If
    End If
    
    cbo(11).Enabled = (cbo(10).ItemData(cbo(10).ListIndex) = 2)
    lbl(14).Enabled = cbo(4).Enabled
    lbl(4).Enabled = cbo(8).Enabled
    lbl(5).Enabled = cbo(9).Enabled
    lbl(6).Enabled = cbo(10).Enabled
    lbl(21).Enabled = cbo(11).Enabled
    
    cbo(0).Enabled = (mintType = 0) And IIf(InStr(1, ",-1,1,2,", "," & mlngKey & ",") <> 0, False, True) And (cbo(10).ItemData(cbo(10).ListIndex) = 1)
    txt(12).Enabled = (cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2) And (mintType = 0)
    txt(1).Enabled = txt(12).Enabled
'    cbo(0).Enabled = False
    
    cmd(0).Enabled = txt(6).Enabled
    lbl(1).Enabled = txt(0).Enabled
    
    lbl(7).Enabled = cbo(1).Enabled
    
    lbl(9).Enabled = txt(2).Enabled
    udn(0).Enabled = txt(2).Enabled
    
    lbl(11).Enabled = txt(7).Enabled
    lbl(10).Enabled = txt(5).Enabled
    udn(1).Enabled = txt(5).Enabled
    lbl(12).Enabled = cbo(2).Enabled
    
    lbl(2).Enabled = txt(1).Enabled
    
    lbl(17).Enabled = txt(12).Enabled
    udn(5).Enabled = txt(12).Enabled
    
    lbl(16).Enabled = cbo(3).Enabled
    lbl(3).Enabled = cbo(0).Enabled
    cbo(6).Enabled = cbo(3).Enabled
    cbo(7).Enabled = cbo(3).Enabled
    cbo(12).Enabled = cbo(3).Enabled
    cbo(13).Enabled = cbo(3).Enabled
    
    lbl(0).Enabled = txt(3).Enabled
    udn(2).Enabled = txt(3).Enabled
    
    txt(4).Enabled = (cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2) And (mintType = 0)
    txt(8).Enabled = (cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2) And (mintType = 0)
    lbl(8).Enabled = txt(4).Enabled
    lbl(13).Enabled = txt(4).Enabled
    chkFirst.Enabled = ((cbo(0).ListIndex = 1) And (chk(1).Value = 1) And (Left(cbo(2).Text, 1) <> 4))
    
End Function

Private Function ClearData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:����ؼ��е�����
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    On Error Resume Next
    
    For lngLoop = 0 To txt.UBound
        If lngLoop <> 9 Then txt(lngLoop).Text = ""
    Next
    
    cmd(0).Tag = ""
    cmd(1).Tag = ""
    
    ClearData = True
    
End Function

Private Function CheckData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:У��༭���ݵ���Ч��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim varTmp, VarTmp1
    Dim int��Ŀ��ʾ As Integer
    Dim lngLoop As Long, strValueRange As String
    On Error GoTo errHand
    
    If Trim(cbo(5).Text) = "" Then
        ShowSimpleMsg "�������Ʋ���Ϊ��ֵ���������룡"
        LocationObj cbo(5)
        Exit Function
    End If
    
    If Trim(txt(0).Text) = "" Then
        ShowSimpleMsg "��Ŀ���Ʋ���Ϊ��ֵ���������룡"
        LocationObj txt(0)
        Exit Function
    End If
    
    If InStr(1, "'����'ʱ��'��ʿ'ǩ����'ǩ��ʱ��'", "'" & txt(0).Text & "'") <> 0 Then
        ShowSimpleMsg "[����,ʱ��,��ʿ,ǩ����,ǩ��ʱ��]Ϊ��¼���̶���Ŀ����������ӣ����������룡"
        LocationObj txt(0)
        Exit Function
    End If
    
    If Val(txt(2).Text) <= Val(txt(5).Text) Then
        ShowSimpleMsg "С��λ�����ܴ��ڵ������ݳ��ȣ�"
        LocationObj txt(5)
        Exit Function
    End If
    
	If mlngKey = 0 Then
    	gstrSQL = "Select Upper(��Ŀ����) As ��Ŀ���� From �����¼��Ŀ"
    	Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���еĻ�����Ŀ")
    	rs.Filter = "��Ŀ���� ='" & UCase(Trim(txt(0).Text)) & "'"
    	If rs.RecordCount > 0 Then
        	ShowSimpleMsg "��ǰ��Ŀ�����Ѵ��ڣ����������룡ע�⣺�����ִ�Сд��"
        	LocationObj txt(0)
        	Exit Function
    	End If
	end if
    
    If mlngKey <> 0 Then
        gstrSQL = "select ��Ŀ����,��Ŀ��ʾ From �����¼��Ŀ where ��Ŀ���=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��������Ŀ", mlngKey)
        If rs.RecordCount > 0 Then
            int��Ŀ��ʾ = NVL(rs!��Ŀ��ʾ)
            If NVL(rs!��Ŀ����) <> Trim(txt(0).Text) Then
                If mfrmMain.CheckItemExistData(2, NVL(rs!��Ŀ����)) = True Then
                    txt(0).Text = NVL(rs!��Ŀ����)
                    Exit Function
                End If
            End If
        End If
    End If
    '96901,����,2016-6-7
    If mlngKey <> 0 And ((cbo(1).ListIndex = 0 And cbo(2).ListIndex = 1) Or int��Ŀ��ʾ = 4) Then
        If Not ((cbo(1).ListIndex = 0 And cbo(2).ListIndex = 1) And int��Ŀ��ʾ = 4) Then
            If mfrmMain.CheckItemExistData(1, mlngKey) = True Then
                Call cbo(2).SetFocus
                Exit Function
            End If
        End If
    End If
    
    
    If cbo(1).ListIndex = 0 And (cbo(2).ListIndex = 0 Or cbo(2).ListIndex = 0) Then
        
        If CheckNumber(Trim(vsf.TextMatrix(1, 2)), Val(txt(2).Text), Val(txt(5).Text)) = False Then
            ShowSimpleMsg "��" & Trim(vsf.TextMatrix(1, 2)) & "�����������ݳ��ȶ��壡"
            Call LocationGrid(vsf, 1, 2)
            Exit Function
        End If
        
        If CheckNumber(Trim(vsf.TextMatrix(2, 2)), Val(txt(2).Text), Val(txt(5).Text)) = False Then
            ShowSimpleMsg "��" & Trim(vsf.TextMatrix(2, 2)) & "�����������ݳ��ȶ��壡"
            Call LocationGrid(vsf, 2, 2)
            Exit Function
        End If
        
        If CheckNumber(Trim(vsf.TextMatrix(3, 2)), Val(txt(2).Text), Val(txt(5).Text)) = False Then
            ShowSimpleMsg "��" & Trim(vsf.TextMatrix(3, 2)) & "�����������ݳ��ȶ��壡"
            Call LocationGrid(vsf, 3, 2)
            Exit Function
        End If
        
        If Trim(vsf.TextMatrix(1, 2)) = "" Then
            ShowSimpleMsg "����������Сֵ��"
            Call LocationGrid(vsf, 1, 2)
            Exit Function
        End If
        
        If Trim(vsf.TextMatrix(2, 2)) = "" Then
            ShowSimpleMsg "�����������ֵ��"
            Call LocationGrid(vsf, 2, 2)
            Exit Function
        End If
        
        If Val(vsf.TextMatrix(1, 2)) >= Val(vsf.TextMatrix(2, 2)) Then
            ShowSimpleMsg "���ֵ����С����Сֵ��"
            Call LocationGrid(vsf, 2, 2)
            Exit Function
        End If

        
        If vsf.TextMatrix(3, 2) <> "" And (Val(vsf.TextMatrix(3, 2)) > Val(vsf.TextMatrix(2, 2)) Or Val(vsf.TextMatrix(3, 2)) < Val(vsf.TextMatrix(1, 2))) Then
            ShowSimpleMsg "ȱʡֵ���ܴ������ֵ����С����Сֵ��"
            Call LocationGrid(vsf, 3, 2)
            Exit Function
        End If
        
        If vsf.Body.RowHidden(4) = False Then
            If Trim(vsf.TextMatrix(4, 2)) = "" Then
                ShowSimpleMsg "�������ÿ̶�ֵ��"
                Call LocationGrid(vsf, 4, 2)
                Exit Function
            End If
            
            If InStr(1, vsf.TextMatrix(4, 2), ";") = 0 Then
                ShowSimpleMsg "�̶�ֵ�ĸ�ʽ����ȷ,��ʽ:��Сֵ;���ֵ��"
                Call LocationGrid(vsf, 4, 2)
                Exit Function
            End If
            varTmp = Split(vsf.TextMatrix(4, 2), ";")
            If Trim(varTmp(0)) = "" Then
                ShowSimpleMsg "�����ÿ̶�ֵ����Сֵ��"
                Call LocationGrid(vsf, 4, 2)
                Exit Function
            End If
            
            If Trim(varTmp(1)) = "" Then
                ShowSimpleMsg "�����ÿ̶�ֵ�����ֵ��"
                Call LocationGrid(vsf, 4, 2)
                Exit Function
            End If
            
            If Val(varTmp(0)) >= Val(varTmp(1)) Then
                ShowSimpleMsg "�̶�ֵ�����ֵ����С����Сֵ��"
                Call LocationGrid(vsf, 4, 2)
                Exit Function
            End If
            If Val(varTmp(0)) < Val(vsf.TextMatrix(1, 2)) Or Val(varTmp(0)) > Val(vsf.TextMatrix(2, 2)) Then
                ShowSimpleMsg "�̶�ֵ����Сֵ����ֵ����Сֵ�����ֵ֮�䣡"
                Call LocationGrid(vsf, 4, 2)
                Exit Function
            End If
            
            If Val(varTmp(1)) < Val(vsf.TextMatrix(1, 2)) Or Val(varTmp(1)) > Val(vsf.TextMatrix(2, 2)) Then
                ShowSimpleMsg "�̶�ֵ�����ֵ����ֵ����Сֵ�����ֵ֮�䣡"
                Call LocationGrid(vsf, 4, 2)
                Exit Function
            End If
        End If
        
        If vsf.Body.RowHidden(5) = False And vsf.TextMatrix(5, 2) <> "" Then
            If InStr(1, vsf.TextMatrix(5, 2), ";") = 0 Then
                ShowSimpleMsg "�ٽ�ֵ�ĸ�ʽ����ȷ,��ʽ:��Сֵ;���ֵ��"
                Call LocationGrid(vsf, 5, 2)
                Exit Function
            End If
            VarTmp1 = Split(vsf.TextMatrix(5, 2), ";")
            
            If Trim(VarTmp1(0)) = "" Then
                ShowSimpleMsg "�������ٽ�ֵ����Сֵ��"
                Call LocationGrid(vsf, 5, 2)
                Exit Function
            End If
            
            If Trim(VarTmp1(1)) = "" Then
                ShowSimpleMsg "�������ٽ�ֵ�����ֵ��"
                Call LocationGrid(vsf, 5, 2)
                Exit Function
            End If
            
            If Val(VarTmp1(0)) >= Val(VarTmp1(1)) Then
                ShowSimpleMsg "�ٽ�ֵ�����ֵ����С����Сֵ��"
                Call LocationGrid(vsf, 5, 2)
                Exit Function
            End If
            varTmp = Split(vsf.TextMatrix(4, 2), ";")
            If Val(VarTmp1(0)) < Val(varTmp(0)) Or Val(VarTmp1(0)) > Val(varTmp(1)) Then
                ShowSimpleMsg "�ٽ�ֵ����Сֵ���ڿ̶�ֵ��Сֵ�����ֵ֮�䣡"
                Call LocationGrid(vsf, 5, 2)
                Exit Function
            End If
            
            If Val(VarTmp1(1)) < Val(varTmp(0)) Or Val(VarTmp1(1)) > Val(varTmp(1)) Then
                ShowSimpleMsg "�ٽ�ֵ�����ֵ���ڿ̶�ֵ��Сֵ�����ֵ֮�䣡"
                Call LocationGrid(vsf, 5, 2)
                Exit Function
            End If
        End If
        
    End If
        
    If chk(1).Value = 1 Then
        If Val(txt(3).Text) = 5 Or Val(txt(3).Text) = 0 Then
            MsgBox "��¼Ƶ��ֻ���ǣ�1,2,3,4,6", vbInformation, gstrSysName
            Exit Function
        End If
        
        If cbo(1).ListIndex <> 0 Then
            If cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2 Then
                ShowSimpleMsg "��������Ŀ�����±��в�����������Ŀ��"
                LocationObj cbo(0)
                Exit Function
            End If
                    
        End If
        
        If cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2 Then
            If Trim(cbo(3).Text) = "" Then
                ShowSimpleMsg IIf(mintType = 0, "����", "����") & "������Ŀ��������һ����¼���ţ�"
                LocationObj cbo(3)
                Exit Function
            End If
            If Trim(txt(4).Text) = "" Then
                ShowSimpleMsg IIf(mintType = 0, "����", "����") & "������Ŀ�������ÿ̶ȼ����"
                txt(4).SetFocus
                Exit Function
            End If
            If Val(txt(4).Text) > (Val(vsf.TextMatrix(2, 2)) - Val(vsf.TextMatrix(1, 2))) Then
                ShowSimpleMsg "�̶ȼ����ֵ���ܳ�����Ч��Χ��"
                txt(4).SetFocus
                Exit Function
            End If
            If Trim(txt(8).Text) <> "" Then
                If Not (Val(txt(8).Text) >= Val(vsf.TextMatrix(1, 2)) And Val(txt(8).Text) <= Val(vsf.TextMatrix(2, 2))) And mintType = 0 Then
                    ShowSimpleMsg "��ʾ�ߵ�ֵֻ�������ֵ����Сֵ�ķ�Χ�ڣ�"
                    txt(8).SetFocus
                    Exit Function
                End If
            End If
        End If
    
        If Val(txt(1).Text) = 0 And (cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2) Then
            ShowSimpleMsg IIf(mintType = 0, "����", "����") & "������Ŀ����ָ����λֵ��"
            LocationObj txt(1)
            Exit Function
        End If
        
        If cbo(1).ListIndex = 0 And cbo(2).ListIndex = 0 Then
            If txt(1).Text <> "" And (Val(txt(1).Text) > (Val(vsf.TextMatrix(2, 2)) - Val(vsf.TextMatrix(1, 2)))) Then
                ShowSimpleMsg "��λֵ���ܴ������ֵ����Сֵ֮��Ĳ�ֵ��"
                LocationObj txt(1)
                Exit Function
            End If
        End If
        
        If CheckStrType(Trim(txt(7).Text), 2, "0123456789") = False And txt(7).Enabled And txt(7).Text <> "" Then
            ShowSimpleMsg "��λ�в��ܺ��������ַ���"
            LocationObj txt(7)
            Exit Function
        End If

        If Val(txt(12).Text) = 0 And (cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2) And mintType = 0 Then
            ShowSimpleMsg "����������Ŀ����ָ������У�"
            LocationObj txt(12)
            Exit Function
        End If
        
        If Trim(cbo(3).Text) = "" And (cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2) Then
            ShowSimpleMsg "��¼������Ϊ��ֵ���������룡"
            LocationObj cbo(3)
            Exit Function
        End If
        
        If Trim(cbo(6).Text) = "" And (cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2) And mlngKey = 1 Then
            ShowSimpleMsg "��¼������Ϊ��ֵ���������룡"
            LocationObj cbo(6)
            Exit Function
        End If
        
        If Trim(cbo(7).Text) = "" And (cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2) And mlngKey = 1 Then
            ShowSimpleMsg "��¼������Ϊ��ֵ���������룡"
            LocationObj cbo(7)
            Exit Function
        End If
        
        If Trim(cbo(12).Text) = "" And (cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2) And mlngKey = 1 Then
            ShowSimpleMsg "��¼������Ϊ��ֵ���������룡"
            LocationObj cbo(12)
            Exit Function
        End If
        
        If Trim(cbo(13).Text) = "" And (cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2) And mlngKey = 1 Then
            ShowSimpleMsg "��¼������Ϊ��ֵ���������룡"
            LocationObj cbo(13)
            Exit Function
        End If
        
        If Trim(cbo(3).Text) = "'" Then
            ShowSimpleMsg "��¼������Ϊ������(')��"
            LocationObj cbo(3)
            Exit Function
        End If

        If Trim(cbo(6).Text) = "'" And mlngKey = 1 Then
            ShowSimpleMsg "��¼������Ϊ������(')��"
            LocationObj cbo(6)
            Exit Function
        End If
        
        If Trim(cbo(7).Text) = "'" And mlngKey = 1 Then
            ShowSimpleMsg "��¼������Ϊ������(')��"
            LocationObj cbo(7)
            Exit Function
        End If
        
        If Trim(cbo(12).Text) = "'" And mlngKey = 1 Then
            ShowSimpleMsg "��¼������Ϊ������(')��"
            LocationObj cbo(12)
            Exit Function
        End If
        
        If Trim(cbo(13).Text) = "'" And mlngKey = 1 Then
            ShowSimpleMsg "��¼������Ϊ������(')��"
            LocationObj cbo(13)
            Exit Function
        End If
        
        '����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
        If cbo(3).Enabled Then

            If InStr(cbo(3).Text, "'") > 0 Or InStr(cbo(3).Text, ";") > 0 Or InStr(cbo(3).Text, ",") > 0 Or InStr(cbo(3).Text, "`") > 0 Or InStr(cbo(3).Text, """") > 0 Then
                ShowSimpleMsg "��¼�������зǷ��ַ���"
                LocationObj cbo(3)
                Exit Function
            End If

            If Len(cbo(3).Text) > 1 Then
                ShowSimpleMsg "��¼���Ų��ܳ��� 1 ���ַ���"
                LocationObj cbo(3)
                Exit Function
            End If
        End If
        
        If mlngKey = 1 Then
        
            If cbo(6).Enabled Then
    
                If InStr(cbo(6).Text, "'") > 0 Or InStr(cbo(6).Text, ";") > 0 Or InStr(cbo(6).Text, ",") > 0 Or InStr(cbo(6).Text, "`") > 0 Or InStr(cbo(6).Text, """") > 0 Then
                    ShowSimpleMsg "��¼�������зǷ��ַ���"
                    LocationObj cbo(6)
                    Exit Function
                End If
    
                If Len(cbo(6).Text) > 1 Then
                    ShowSimpleMsg "��¼���Ų��ܳ��� 1 ���ַ���"
                    LocationObj cbo(6)
                    Exit Function
                End If

                
            End If
        
            If cbo(7).Enabled Then
    
                If InStr(cbo(7).Text, "'") > 0 Or InStr(cbo(7).Text, ";") > 0 Or InStr(cbo(7).Text, ",") > 0 Or InStr(cbo(7).Text, "`") > 0 Or InStr(cbo(7).Text, """") > 0 Then
                    ShowSimpleMsg "��¼�������зǷ��ַ���"
                    LocationObj cbo(7)
                    Exit Function
                End If
    
                If Len(cbo(7).Text) > 1 Then
                    ShowSimpleMsg "��¼���Ų��ܳ��� 1 ���ַ���"
                    LocationObj cbo(7)
                    Exit Function
                End If
            End If
            
            If cbo(12).Enabled Then
    
                If InStr(cbo(12).Text, "'") > 0 Or InStr(cbo(12).Text, ";") > 0 Or InStr(cbo(12).Text, ",") > 0 Or InStr(cbo(12).Text, "`") > 0 Or InStr(cbo(12).Text, """") > 0 Then
                    ShowSimpleMsg "��¼�������зǷ��ַ���"
                    LocationObj cbo(12)
                    Exit Function
                End If
    
                If Len(cbo(12).Text) > 1 Then
                    ShowSimpleMsg "��¼���Ų��ܳ��� 1 ���ַ���"
                    LocationObj cbo(12)
                    Exit Function
                End If
            End If
            
            If cbo(13).Enabled Then
    
                If InStr(cbo(13).Text, "'") > 0 Or InStr(cbo(13).Text, ";") > 0 Or InStr(cbo(13).Text, ",") > 0 Or InStr(cbo(13).Text, "`") > 0 Or InStr(cbo(13).Text, """") > 0 Then
                    ShowSimpleMsg "��¼�������зǷ��ַ���"
                    LocationObj cbo(13)
                    Exit Function
                End If
    
                If Len(cbo(13).Text) > 1 Then
                    ShowSimpleMsg "��¼���Ų��ܳ��� 1 ���ַ���"
                    LocationObj cbo(13)
                    Exit Function
                End If
            End If
        End If
    End If
    
    
    
    '������Ŀ�ļ�¼Ƶ�������2��
    If chk(1).Value = 1 Then
        If Mid(cbo(2).Text, 1, 1) = "4" And Val(txt(3).Text) > 2 Then
            txt(3).Text = 2
            MsgBox "������Ŀ�ļ�¼Ƶ�����Ϊ2�Σ�", vbInformation, gstrSysName
            Exit Function
        End If
        '����Ƿ��ǲ�����Ŀ
        If mlngKey <> 0 Then
            gstrSQL = "select 1 From ��������Ŀ where ��Ŀ���=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��������Ŀ", mlngKey)
            If rs.RecordCount > 0 And Val(txt(3).Text) > 2 Then
                txt(3).Text = 2
                MsgBox "������Ŀ�ļ�¼Ƶ�����Ϊ2�Σ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '90060,��ѡ��ѡ��Ŀ������������
    If Val(vsf.Tag) = 2 Or Val(vsf.Tag) = 3 Then
        For lngLoop = 1 To vsf.Rows - 1
            If vsf.TextMatrix(lngLoop, 1) <> "" Then
                If Abs(Val(vsf.TextMatrix(lngLoop, 2))) = 1 Then
                    strValueRange = strValueRange & ";��" & vsf.TextMatrix(lngLoop, 1)
                Else
                    strValueRange = strValueRange & ";" & vsf.TextMatrix(lngLoop, 1)
                End If
            End If
        Next
        If strValueRange = "" Then
            ShowSimpleMsg "��Ŀ��ֵ��Χ����Ϊ�գ�"
            Call LocationGrid(vsf, 1, 1)
            Exit Function
        End If
    End If
    
    CheckData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SaveData(ByRef lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ������޸Ļ�����������
    '���أ��ɹ����淵��True�����򷵻�False
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim strValueRange As String
    Dim strMax As String
    Dim strMin As String
    Dim strCritical As String
    Dim strTmp As String
    Dim intApplications As Integer
    Dim strȱʡ As String
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    strMin = "NULL"
    strMax = "NULL"
    strCritical = ""
    Select Case Val(vsf.Tag)
    Case 1
        strValueRange = Trim(vsf.TextMatrix(1, 2)) & ";" & Trim(vsf.TextMatrix(2, 2)) & ";" & Trim(vsf.TextMatrix(3, 2))
        If strValueRange = ";;" Then strValueRange = ""
        
        If vsf.Body.RowHidden(5) = False Then
            strMin = Split(vsf.TextMatrix(4, 2), ";")(0)
            strMax = Split(vsf.TextMatrix(4, 2), ";")(1)
        Else
            strMin = Trim(vsf.TextMatrix(1, 2))
            strMax = Trim(vsf.TextMatrix(2, 2))
        End If
        '�ٽ�ֵ
        If vsf.Body.RowHidden(5) = False Then strCritical = Trim(vsf.TextMatrix(5, 2))
        
        strMin = IIf(strMin = "", "NULL", Val(strMin))
        strMax = IIf(strMax = "", "NULL", Val(strMax))
        strCritical = strCritical
        
    Case 2, 3, 4
        For lngLoop = 1 To vsf.Rows - 1
            If vsf.TextMatrix(lngLoop, 1) <> "" Then
                
                If Abs(Val(vsf.TextMatrix(lngLoop, 2))) = 1 Then
                    strValueRange = strValueRange & ";" & vsf.TextMatrix(lngLoop, 1)
                    strȱʡ = strȱʡ & ";" & vsf.TextMatrix(lngLoop, 1)
                Else
                    strValueRange = strValueRange & ";" & vsf.TextMatrix(lngLoop, 1)
                End If
            End If
        Next
        If Trim(strValueRange) <> "" Then strValueRange = Mid(strValueRange, 2)
        If strȱʡ <> "" Then strȱʡ = Mid(strȱʡ, 2)
    End Select
    
    intApplications = 0
    If cbo(10).ItemData(cbo(10).ListIndex) = 2 Then '���Ŀ
        intApplications = IIf(cbo(11).ListIndex < 0, 0, cbo(11).ListIndex)
    End If
    '--48659:������,2012-09-14,����ֶ�'˵��'
    If mlngKey = 0 Then
        '����
    '    ��Ŀ���_IN IN  �����¼��Ŀ.��Ŀ���%TYPE,
    '    ��Ŀ����_IN IN  �����¼��Ŀ.��Ŀ����%TYPE,
    '    ��Ŀ����_IN IN  �����¼��Ŀ.��Ŀ����%TYPE,
    '    ��Ŀ����_IN IN  �����¼��Ŀ.��Ŀ����%TYPE,
    '    ��ĿС��_IN IN  �����¼��Ŀ.��ĿС��%TYPE,
    '    ��Ŀ��λ_IN IN  �����¼��Ŀ.��Ŀ��λ%TYPE,
    '    ��Ŀ��ʾ_IN IN  �����¼��Ŀ.��Ŀ��ʾ%TYPE,
    '    ��Ŀֵ��_IN IN  �����¼��Ŀ.��Ŀֵ��%TYPE,
    '    ����ȼ�_IN   IN  �����¼��Ŀ.����ȼ�%TYPE,
    '    ������_IN   IN  �����¼��Ŀ.������%TYPE,
    '    ��ĿID_IN   IN  �����¼��Ŀ.��ĿID%TYPE
    
        lngKey = GetMaxNo
        strSQL(ReDimArray(strSQL)) = "ZL_�����¼��Ŀ_INSERT(" & lngKey & ",'" & _
                                                            Trim(txt(0).Text) & "'," & _
                                                            cbo(1).ListIndex & "," & _
                                                            Val(txt(2).Text) & "," & _
                                                            IIf(cbo(1).ListIndex = 0, Val(txt(5).Text), "NULL") & ",'" & _
                                                            IIf(cbo(1).ListIndex = 0, Trim(txt(7).Text), "") & "'," & _
                                                            Left(cbo(2).Text, 1) & ",'" & _
                                                            strValueRange & "'," & _
                                                            cbo(4).ListIndex & ",'" & _
                                                            Trim(cbo(5).Text) & "'," & _
                                                            IIf(Val(cmd(0).Tag) = 0, "NULL", Val(cmd(0).Tag)) & "," & Left(cbo(8).Text, 1) & "," & Left(cbo(9).Text, 1) & "," & cbo(10).ItemData(cbo(10).ListIndex) & "," & intApplications & ",'" & Replace(Trim(txtInfo.Text), "'", "") & "','" & strȱʡ & "')"
    Else
        '�޸�
        lngKey = mlngKey
        
        strSQL(ReDimArray(strSQL)) = "ZL_�����¼��Ŀ_UPDATE(" & lngKey & ",'" & _
                                                            Trim(txt(0).Text) & "'," & _
                                                            cbo(1).ListIndex & "," & _
                                                            Val(txt(2).Text) & "," & _
                                                            IIf(cbo(1).ListIndex = 0, Val(txt(5).Text), "NULL") & ",'" & _
                                                            IIf(cbo(1).ListIndex = 0, Trim(txt(7).Text), "") & "'," & _
                                                            Left(cbo(2).Text, 1) & ",'" & _
                                                            strValueRange & "'," & _
                                                            cbo(4).ListIndex & ",'" & _
                                                            Trim(cbo(5).Text) & "'," & _
                                                            IIf(Val(cmd(0).Tag) = 0, "NULL", Val(cmd(0).Tag)) & "," & Left(cbo(8).Text, 1) & "," & Left(cbo(9).Text, 1) & "," & cbo(10).ItemData(cbo(10).ListIndex) & "," & intApplications & ",'" & Replace(Trim(txtInfo.Text), "'", "") & "','" & strȱʡ & "')"
                                                            
    End If
    
    
    
    If chk(1).Value = 1 Then
        
        If mlngKey <> 1 Then
            strTmp = IIf((cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2), UCase(Trim(cbo(3).Text)), "")
        Else
            strTmp = IIf(cbo(0).ListIndex = 0, UCase(Trim(cbo(3).Text)) & "," & UCase(Trim(cbo(6).Text)) & "," & UCase(Trim(cbo(7).Text)) & "," & UCase(Trim(cbo(12).Text)) & "," & UCase(Trim(cbo(13).Text)), "")
        End If
        
        strSQL(ReDimArray(strSQL)) = "ZL_���¼�¼��Ŀ_INSERT(" & lngKey & "," & _
                                                            Val(txt(13).Text) & ",'" & _
                                                            Trim(txt(0).Text) & "'," & _
                                                            cbo(0).ListIndex + 1 & ",'" & _
                                                            strTmp & "'," & _
                                                            Val(picDemo.Tag) & "," & _
                                                            IIf(Val(vsf.Tag) = 1, strMin, "NULL") & "," & _
                                                            IIf(Val(vsf.Tag) = 1, strMax, "NULL") & "," & _
                                                            IIf((cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2), Val(txt(1).Text), "NULL") & ",'" & _
                                                            Trim(txt(7).Text) & "'," & _
                                                            IIf((cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2), Val(txt(12).Text), "NULL") & "," & _
                                                            Val(txt(3).Text) & "," & IIf(txt(4).Text = "", "NULL", Val(txt(4).Text)) & "," & _
                                                            IIf(txt(8).Text = "", "NULL", Val(txt(8).Text)) & ",'" & _
                                                            IIf(Val(vsf.Tag) = 1, strCritical, "") & "'," & _
                                                            IIf(chkFirst.Enabled = True, chkFirst.Value, 0) & ")"
    Else
        strSQL(ReDimArray(strSQL)) = "ZL_���¼�¼��Ŀ_DELETE(" & lngKey & ")"
    End If
    
    'ִ��
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveData = True
    
    Exit Function
    
errHand:
    '������
    
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    
    
End Function

Private Function GetMaxNo(Optional ByVal bytMode As Byte = 1) As Long
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ��һ�����
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    If bytMode = 1 Then
        mstrSQL = "SELECT NVL(MAX(��Ŀ���),0)+1 AS ��� FROM �����¼��Ŀ"
    Else
        mstrSQL = "SELECT NVL(MAX(�������),0)+1 AS ��� FROM ���¼�¼��Ŀ"
    End If
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    If rs.BOF = False Then GetMaxNo = rs("���").Value
        
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ�����ݣ�һ��ָ�ؼ������ݳ�ʼ��
    '------------------------------------------------------------------------------------------------------------------
    Dim obj As ComboItem
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    mbln������Ŀ = False
    mintType = 0
    '3.װ������A-Z���ַ�
    cbo(3).Clear
    cbo(6).Clear
    cbo(7).Clear
    cbo(12).Clear
    cbo(13).Clear
    For intLoop = 65 To 90
        cbo(3).AddItem Chr(intLoop)
        cbo(6).AddItem Chr(intLoop)
        cbo(7).AddItem Chr(intLoop)
        cbo(12).AddItem Chr(intLoop)
        cbo(13).AddItem Chr(intLoop)
    Next
    cbo(3).AddItem "��"
    cbo(3).AddItem "��"
    cbo(3).AddItem "��"
    cbo(3).AddItem "+"
    cbo(3).AddItem "*"
    cbo(3).AddItem "��"

    cbo(6).AddItem "��"
    cbo(6).AddItem "��"
    cbo(6).AddItem "��"
    cbo(6).AddItem "+"
    cbo(6).AddItem "*"
    cbo(6).AddItem "��"
    
    cbo(7).AddItem "��"
    cbo(7).AddItem "��"
    cbo(7).AddItem "��"
    cbo(7).AddItem "+"
    cbo(7).AddItem "*"
    cbo(7).AddItem "��"
    
    cbo(12).AddItem "��"
    cbo(12).AddItem "��"
    cbo(12).AddItem "��"
    cbo(12).AddItem "��"
    cbo(12).AddItem "��"
    cbo(12).AddItem "+"
    cbo(12).AddItem "*"
    cbo(12).AddItem "��"
    
    cbo(13).AddItem "��"
    cbo(13).AddItem "��"
    cbo(13).AddItem "��"
    cbo(13).AddItem "��"
    cbo(13).AddItem "��"
    cbo(13).AddItem "+"
    cbo(13).AddItem "*"
    cbo(13).AddItem "��"
    
    With cbo(10)
        .Clear
        .AddItem "1-�̶���Ŀ"
        .ItemData(.NewIndex) = 1
        .AddItem "2-���Ŀ"
        .ItemData(.NewIndex) = 2
        .ListIndex = 0
    End With
    
    With cbo(11)
        .Clear
        .AddItem "0-ͨ��": .ItemData(.NewIndex) = 0
        .AddItem "1-���µ�": .ItemData(.NewIndex) = 1
        .AddItem "2-��¼��": .ItemData(.NewIndex) = 2
        .ListIndex = 0
    End With
    udn(0).Min = 1
    udn(0).Max = 999
    
    udn(1).Min = 0
    udn(1).Max = 4
    
    udn(5).Min = 1
    udn(5).Max = 60
    
    udn(6).Min = 1
    udn(6).Max = 100

    udn(2).Min = 1
    udn(2).Max = 6
    
    With cbo(8)
        .AddItem "0-����ʹ��"
        .AddItem "1-����ʹ��"
        If mlngKey = -1 Then .AddItem "2-����������"
        .ListIndex = 1
    End With

    With cbo(9)
        .AddItem "0-����"
        .AddItem "1-����"
        .AddItem "2-Ӥ��"
        .ListIndex = 1
    End With
    
    
    '����������󳤶�
    txt(0).MaxLength = GetMaxLength("�����¼��Ŀ", "��Ŀ����")
    txt(7).MaxLength = GetMaxLength("�����¼��Ŀ", "��Ŀ��λ")
    
    cbo(0).Clear
    cbo(0).AddItem "1-����"
    cbo(0).AddItem "2-���"
    cbo(0).AddItem "3-��������"
    cbo(0).ListIndex = 1
    
    cbo(1).Clear
    cbo(1).AddItem "0-��ֵ"
    cbo(1).AddItem "1-����"
    cbo(1).ListIndex = 0
    
    cbo(2).Clear
    cbo(2).AddItem "0-�ı�"
    cbo(2).AddItem "2-��ѡ"
    cbo(2).AddItem "3-��ѡ"
    cbo(2).AddItem "4-����"
    cbo(2).AddItem "5-ѡ��"
    
    cbo(2).ListIndex = 0
    
    cbo(4).Clear
    cbo(4).AddItem "0-�ؼ�����"
    cbo(4).AddItem "1-һ������"
    cbo(4).AddItem "2-��������"
    cbo(4).AddItem "3-��������"
    cbo(4).ListIndex = 3
    
    With vsf
        .Cols = 0
        .NewColumn "", 255
        .NewColumn "", 2700
        .NewColumn "", 450
    End With
    
    cbo(5).Clear
    mstrSQL = "SELECT ������ FROM �����¼��Ŀ GROUP BY ������ Order By ������"
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            cbo(5).AddItem zlCommFun.NVL(rs("������"))
            rs.MoveNext
        Loop
    End If
    
    Call InitGrid
    
    If mlngKey = -1 Then
        cbo(4).Enabled = False
    End If
    
    picTemperaPart.Height = 1980
    picTemperaPart.Visible = False
    cbo(3).Top = 90
    cbo(3).Left = 810
    picTemperaPart.Top = 3480
    If mlngKey <> 1 Then
        picTemperaPart.Height = cbo(1).Height
        picTemperaPart.Width = cbo(1).Width
        picTemperaPart.Visible = True
        picTemperaPart.BorderStyle = 0
        picTemperaPart.Top = fra(2).Top + txtTemper.Top
        picTemperaPart.Left = fra(2).Left + txtTemper.Left
        cbo(3).Top = 0
        cbo(3).Left = 0
        cbo(6).Enabled = False
        cbo(7).Enabled = False
        cbo(12).Enabled = False
        
        cbo(6).Visible = False
        cbo(7).Visible = False
        cbo(12).Visible = False

    End If
    
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function InitGrid() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ�����ؼ�
    '------------------------------------------------------------------------------------------------------------------
    Dim bytMode As Byte

    'bytMode:1��ʾ;2��ʾ;3��ʾ
    
    If cbo(2).ListIndex < 0 Then cbo(2).ListIndex = 0
    Select Case cbo(1).ListIndex
    Case 0  '��ֵ��
        Select Case Left(cbo(2).Text, 1)
        Case 0, 1
            bytMode = 1
        Case Else
            bytMode = Left(cbo(2).Text, 1)
        End Select
    Case 1  '������
        Select Case Left(cbo(2).Text, 1)
        Case 2, 3, 4, 5
            bytMode = Left(cbo(2).Text, 1)
        Case Else
            bytMode = 0
        End Select
    End Select
    
    
    With vsf
        If Val(.Tag) <> bytMode Then
            .Tag = bytMode
            
            Select Case bytMode
            Case 1          '(��ֵ��)��(�ı������¿�)������С���Χ
'                fra(1).Enabled = True
                vsf.Visible = True
                .Cols = 3
                .Rows = 6
                .FixedCols = 2
                
'                .Body.RowHidden(0) = True
                .Body.ColWidth(0) = 255
                .Body.ColWidth(1) = 900
                .Body.ColWidth(2) = 2700 - 450
                .Body.RowHidden(4) = True
                .Body.RowHidden(5) = True
                
                .EditMode(1) = 0
                .EditMode(2) = 1
                .ColDataType(2) = flexDTString
                
                .TextMatrix(0, 1) = "��Ŀ"
                .TextMatrix(0, 2) = "���"
                                
                .TextMatrix(1, 1) = "��Сֵ"
                .TextMatrix(2, 1) = "���ֵ"
                .TextMatrix(3, 1) = "ȱʡֵ"
                .TextMatrix(4, 1) = "�̶�ֵ"
                .TextMatrix(5, 1) = "�ٽ�ֵ"
                .Cell(flexcpText, 1, 2, .Rows - 1, 2) = ""
                
            Case 2, 3   '(��ֵ�͡�������)��(�����򡢵�ѡ�򡢸�ѡ��)�����ѡ����
'                fra(1).Enabled = True
                vsf.Visible = True
                .Cols = 3
                .Rows = 2
                .FixedCols = 1
                .Cell(flexcpText, 1, 1, 1, .Cols - 1) = ""
                
'                .Body.RowHidden(0) = False
                .Body.ColWidth(0) = 255
                .Body.ColWidth(1) = 2700
                .Body.ColWidth(2) = 450
                
                .EditMode(1) = 1
                .EditMode(2) = 1

                .TextMatrix(0, 1) = "��ѡ"
                .TextMatrix(0, 2) = "ȱʡ"
                .Body.ColDataType(2) = flexDTBoolean
                
                .Cell(flexcpText, 1, 2, .Rows - 1, 2) = 0
                
            Case Else
'                fra(1).Enabled = False
                vsf.Visible = False
            End Select
        End If
    End With
    
    InitGrid = True
    
End Function

'######################################################################################################################
'�ؼ�������ȶ�������ԡ����̡��¼�����������

Private Sub cbo_Change(Index As Integer)
    If mblnStartUp Then Exit Sub
    
    DataChanged = True
End Sub

Private Sub cbo_Click(Index As Integer)
    On Error GoTo errHand
    If mblnStartUp Then Exit Sub
    
    DataChanged = True
    
    Select Case Index
    Case 0
        txt(1).Enabled = (cbo(Index).ListIndex = 0 Or cbo(Index).ListIndex = 2)
        txt(12).Enabled = (cbo(Index).ListIndex = 0 Or cbo(Index).ListIndex = 2)
        txt(3).Enabled = (cbo(Index).ListIndex = 1)
        cbo(3).Enabled = (cbo(Index).ListIndex = 0 Or cbo(Index).ListIndex = 2)
        
        Call DrawDemo(picDemo, cbo(0).ListIndex, Val(picDemo.Tag))
        
        Call SetLabelEnable
        
        If cbo(0).ListIndex = 1 Then
            cbo(3).Text = ""
            cbo(6).Text = ""
            cbo(7).Text = ""
            txt(1).Text = ""
            txt(12).Text = ""
        End If
        If Val(vsf.Tag) = 1 Then
            '����������Ŀ�����ٽ�ֵ
            If chk(1).Value = 1 And InStr(1, ",1,3,", "," & Left(cbo(0).Text, 1) & ",") > 0 And chk(1).Caption <> "������Ŀ" Then
                vsf.Body.RowHidden(4) = False
                vsf.Body.RowHidden(5) = False
            Else
                vsf.Body.RowHidden(4) = True
                vsf.Body.RowHidden(5) = True
            End If
        End If
        
        fra(1).Enabled = (cbo(0).ListIndex = 0 Or cbo(0).ListIndex = 2)
    Case 1
        
        Call InitGrid
        
        txt(5).Enabled = (cbo(Index).ListIndex = 0)
        txt(7).Enabled = (cbo(Index).ListIndex = 0)
        
        Call SetLabelEnable
        Dim intSave As Integer
        
        intSave = Left(cbo(2).Text, 1)
        
        If cbo(Index).ListIndex = 1 Then
            txt(5).Text = ""
            txt(7).Text = ""
            
            cbo(2).Clear
            cbo(2).AddItem "0-�ı�"
            cbo(2).AddItem "2-��ѡ"
            cbo(2).AddItem "3-��ѡ"

            Select Case intSave
            Case 2
                cbo(2).ListIndex = 1
            Case 3
                cbo(2).ListIndex = 2
            Case Else
                cbo(2).ListIndex = 0
            End Select
        
        Else
            '��ֵ��
            cbo(2).Clear
            cbo(2).AddItem "0-�ı�"
            cbo(2).AddItem "4-����"
            cbo(2).AddItem "5-ѡ��"
            
            Select Case intSave
            Case 4
                cbo(2).ListIndex = 1
            Case 5
                cbo(2).ListIndex = 2
            Case Else
                cbo(2).ListIndex = 0
            End Select
        End If
        
        Call SetGridFormat
    Case 2
    
        Call InitGrid
        Call SetGridFormat
        chkFirst.Enabled = ((cbo(0).ListIndex = 1) And (chk(1).Value = 1) And (Left(cbo(2).Text, 1) <> 4))
    Case 10
        If cbo(10).ItemData(cbo(10).ListIndex) = 2 Then
            cbo(0).ListIndex = 1
        Else
            cbo(11).ListIndex = 0
        End If
        SetLabelEnable
        Call cbo_Click(11)
    Case 11
        If cbo(11).ListIndex = 2 Then
            chk(1).Value = 0
            chk(1).Enabled = False
        Else
            chk(1).Enabled = True
            If cbo(10).ItemData(cbo(10).ListIndex) = 2 Then
                chk(1).Value = 1: chk(1).Enabled = False
            Else
                If mbln������Ŀ And mlngKey <> 6 And mlngKey <> 8 Then chk(1).Enabled = False
            End If
        End If
    End Select
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbo_GotFocus(Index As Integer)
    If Index = 5 Then
        zlCommFun.OpenIme True
    End If
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 44 Or KeyAscii = 59 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cbo_LostFocus(Index As Integer)
    If Index = 5 Then
        zlCommFun.OpenIme False
    End If
    Call picTemperaPart_LostFocus
End Sub

Private Sub cbo_Validate(Index As Integer, Cancel As Boolean)
    
    If Index = 5 Then
        Cancel = Not StrIsValid(cbo(Index).Text, 20)
    End If
    
End Sub

Private Sub chk_Click(Index As Integer)
    
    Select Case Index
    Case 0
        
        txt(6).Enabled = (chk(Index).Value = 1)
        cbo(1).Enabled = Not txt(6).Enabled
        txt(2).Enabled = Not txt(6).Enabled
        txt(5).Enabled = Not txt(6).Enabled
        txt(7).Enabled = Not txt(6).Enabled
        
        txt(0).Enabled = (chk(Index).Value <> 1)
        txt(6).BackColor = IIf(chk(Index).Value = 1, &H80000005, &H8000000F)
        If txt(6).Enabled = False Then
            
            txt(6).Text = ""
            txt(6).Tag = ""
            cmd(0).Tag = ""
            
        End If
        Call SetLabelEnable
        
    Case 1
        If chk(Index).Value = 1 Then
            Me.Height = 8850
        Else
            Me.Height = 5250
        End If
        
        DataChanged = True
    End Select
    
    '����������Ŀ�����ٽ�ֵ
    If Val(vsf.Tag) = 1 Then
        If chk(1).Value = 1 And InStr(1, ",1,3,", "," & Left(cbo(0).Text, 1) & ",") > 0 And chk(1).Caption <> "������Ŀ" Then
            vsf.Body.RowHidden(4) = False
            vsf.Body.RowHidden(5) = False
        Else
            vsf.Body.RowHidden(4) = True
            vsf.Body.RowHidden(5) = True
        End If
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
            zlCommFun.PressKey vbKeyTab

    End If
End Sub

Private Sub chkFirst_Click()
    DataChanged = True
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
    
    Select Case Index
    Case 0
        mstrSQL = "select * from ( " & _
                    "select A.����,A.ID,A.�ϼ�id,A.����,A.����,'' As Ӣ����,'' As ����,'' As ��λ,0 As ĩ��,0 As ����,0 As С��,0 As ��ʾ��,'' As ��ֵ��,0 As H���� from ( " & _
                        "select '1' As ����,-1 AS ID,Null+0 As �ϼ�id,'0' As ����,'������Ŀ' As ���� From dual " & _
                        "Union All " & _
                        "select '2' As ����,-2 AS ID,Null+0 As �ϼ�id,'0' As ����,'��ʷ����' As ���� From dual " & _
                        "Union All " & _
                        "select '4' As ����,-4 AS ID,Null+0 As �ϼ�id,'0' As ����,'�������' As ���� From dual " & _
                        "Union All " & _
                        "Select ����,ID,Nvl(�ϼ�id,-1) As �ϼ�id,���� ,����  from ������������ where ����=1 Start With �ϼ�id is null connect by prior id =�ϼ�id " & _
                        "Union All " & _
                        "Select ����,ID,Nvl(�ϼ�id,-2) As �ϼ�id,���� ,����  from ������������ where ����=2 Start With �ϼ�id is null connect by prior id =�ϼ�id " & _
                        "Union All " & _
                        "Select ����,ID,Nvl(�ϼ�id,-4) As �ϼ�id,���� ,����  from ������������ where ����=4 Start With �ϼ�id is null connect by prior id =�ϼ�id " & _
                    ") A " & _
                    "Union All " & _
                          "Select '9' As ����,A.ID,A.����id As �ϼ�id,A.����,A.������ As ����,A.Ӣ����,Decode(A.����,1,'����',2,'����',3,'�߼�','��ֵ') As ����,A.��λ,1 As ĩ��,����,С��,��ʾ��,��ֵ��,���� As H���� From ����������Ŀ A Where " & IIf(mbln������Ŀ, " A.����=" & cbo(1).ListIndex, " A.���� Not In (2,3)") & " " & _
                    ") order by ����,����"
        
        If ShowTxtSelectDialog(Me, txt(6), "����,1200,0,1;����,2100,0,0;Ӣ����,900,0,0;����,900,0,0", Me.Name & "\����Ҫ��ѡ��", "�������ѡ��һ������Ҫ�ء�", mstrSQL, rs, 8790, 5100, , Val(cmd(Index).Tag)) Then
            
            txt(6).Text = zlCommFun.NVL(rs("����").Value)
            
            If mbln������Ŀ = False Then
                txt(0).Text = zlCommFun.NVL(rs("����").Value)
                txt(7).Text = zlCommFun.NVL(rs("��λ").Value)
                                            
                On Error Resume Next
                cbo(1).ListIndex = zlCommFun.NVL(rs("H����"), 0)
                If cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
                If cbo(1).ListIndex = 0 Then
                    '�ı���
            
                    '0-�ı�,1-����,2-����,3-��ѡ,4-��ѡ;5-ָ��(����Ŀ������ֵ��ĳ�����ݱ����ͼ���������������ݲ��ṩ)
                    Select Case zlCommFun.NVL(rs("��ʾ��"), 0)
                    Case 3
                        cbo(2).ListIndex = 2
                    Case 4
                        cbo(2).ListIndex = 1
                    Case Else
                        cbo(2).ListIndex = 0
                    End Select
                Else
                    Select Case zlCommFun.NVL(rs("��ʾ��"), 0)
                    Case 1
                        cbo(2).ListIndex = 1
                    Case Else
                        cbo(2).ListIndex = 0
                    End Select
                End If
                On Error GoTo 0
                
                
                txt(2).Text = zlCommFun.NVL(rs("����"))
                txt(5).Text = zlCommFun.NVL(rs("С��"))
                
                
            End If
            
            cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value, 0)
            txt(6).Tag = ""
                       
            usrSaveItem.Ҫ����Ŀ = txt(6).Text
            txt(6).SetFocus
            
            DataChanged = True
        End If
    Case 1
    
        dlg.COLOR = Val(picDemo.Tag)
        dlg.ShowColor
        
        If dlg.COLOR <> Val(picDemo.Tag) Then
            
            picDemo.Tag = dlg.COLOR
            Call DrawDemo(picDemo, cbo(0).ListIndex, dlg.COLOR)
            
            DataChanged = True
            
        End If
        
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    Dim lngKey As Long
    
    If DataChanged Then
        If CheckData = False Then Exit Sub
        If SaveData(lngKey) = False Then Exit Sub
        mblnOk = True
        
        On Error Resume Next
        Call mfrmMain.EditRefresh(lngKey)
        On Error GoTo 0
        
        If mlngKey = 0 Then
            Call ClearData
            cbo(5).SetFocus
            mlngPKey = lngKey
            vsf.Tag = -1
            Call ReadData
            DataChanged = False
            Exit Sub
        End If
        
        DataChanged = False
    End If
    Unload Me
End Sub


Private Sub cmdTemperature_Click()
    If picTemperaPart.Visible = False Then
        picTemperaPart.Visible = True
        cbo(3).SetFocus
    Else
        picTemperaPart.Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DataChanged Then
        Cancel = (MsgBox("����/�޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
End Sub

Private Sub picBack_Paint(Index As Integer)
    If Index = 1 Then
        zlControl.PicShowFlat picBack(1), -1, "����Ŀ��ֵ����", taCenterAlign
    End If
End Sub

Private Sub picTemperaPart_LostFocus()
    If mlngKey <> 1 Then
        picTemperaPart.Visible = True
    Else
        If Me.ActiveControl.Name = "Cbo" Then
            If Me.ActiveControl.Index = 3 Or Me.ActiveControl.Index = 6 Or Me.ActiveControl.Index = 7 Or Me.ActiveControl.Index = 12 Or Me.ActiveControl.Index = 13 Then
                txtTemper.Text = cbo(3).Text & "," & cbo(6).Text & "," & cbo(7).Text & "," & cbo(12).Text & "," & cbo(13).Text

            End If
        ElseIf Me.ActiveControl.Name = "cmdTemperature" Then
            txtTemper.Text = cbo(3).Text & "," & cbo(6).Text & "," & cbo(7).Text & "," & cbo(12).Text & "," & cbo(13).Text
        Else
            picTemperaPart.Visible = False
            txtTemper.Text = cbo(3).Text & "," & cbo(6).Text & "," & cbo(7).Text & "," & cbo(12).Text & "," & cbo(13).Text
        End If
    End If
End Sub

Private Sub txt_Change(Index As Integer)

    DataChanged = True
    
    Select Case Index
    Case 5
        Call SetGridFormat
    Case 6
        txt(Index).Tag = "Changed"
    End Select
        
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    Call zlControl.TxtSelAll(txt(Index))
    
    Select Case Index
    Case 0, 6, 7, 9, 3, 4
        zlCommFun.OpenIme True
    End Select
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rs As New ADODB.Recordset
    Dim strInput As String
    
    If KeyAscii = vbKeyReturn Then
        
        If txt(Index).Tag <> "" Then
            
            strInput = "'%" & UCase(txt(Index).Text) & "%'"
            
            Select Case Index
            Case 6
            
                mstrSQL = "Select ID,����,������ As ����,Ӣ����,Decode(����,1,'����',2,'����',3,'�߼�','��ֵ') As ����,��λ,1 As ĩ��,����,С��,��ʾ��,��ֵ��,���� As H���� From ����������Ŀ Where (���� Like " & strInput & " Or ������ Like " & strInput & " Or Ӣ���� Like " & strInput & ") And ���� Not In (2,3)"
                
                If ShowTxtFilterDialog(Me, txt(Index), "����,1200,0,1;����,2100,0,0;Ӣ����,900,0,0;����,900,0,0", Me.Name & "\����Ҫ�ع���", "����±���ѡ��һ������Ҫ��", mstrSQL, rs, , 4500) Then
                
                    txt(Index).Text = zlCommFun.NVL(rs("����").Value)
                    
                    If mbln������Ŀ = False Then
                        txt(0).Text = zlCommFun.NVL(rs("����").Value)
                        txt(7).Text = zlCommFun.NVL(rs("��λ").Value)
                        
                        On Error Resume Next
                        cbo(1).ListIndex = zlCommFun.NVL(rs("H����"), 0)
                        If cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
                        If cbo(1).ListIndex = 0 Then
                            '�ı���
                    
                            '0-�ı�,1-����,2-����,3-��ѡ,4-��ѡ;5-ָ��(����Ŀ������ֵ��ĳ�����ݱ����ͼ���������������ݲ��ṩ)
                            Select Case zlCommFun.NVL(rs("��ʾ��"), 0)
                            Case 3
                                cbo(2).ListIndex = 2
                            Case 4
                                cbo(2).ListIndex = 1
                            Case Else
                                cbo(2).ListIndex = 0
                            End Select
                        Else
                            Select Case zlCommFun.NVL(rs("��ʾ��"), 0)
                            Case 1
                                cbo(2).ListIndex = 1
                            Case Else
                                cbo(2).ListIndex = 0
                            End Select
                        End If
                        
                        On Error GoTo 0

                        txt(2).Text = zlCommFun.NVL(rs("����"))
                        txt(5).Text = zlCommFun.NVL(rs("С��"))

                    End If
                    
                    cmd(0).Tag = zlCommFun.NVL(rs("ID").Value)
                    txt(Index).Tag = ""
                    
                    DataChanged = True
                    usrSaveItem.Ҫ����Ŀ = txt(Index).Text
                    
                Else
                    txt(Index).Text = usrSaveItem.Ҫ����Ŀ
                    txt(Index).Tag = ""
                    Exit Sub
                End If
            
            End Select
        End If
        zlCommFun.PressKey vbKeyTab
    Else
    
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
        If Chr(KeyAscii) = "*" Then
            
            KeyAscii = 0
            
            Select Case Index
            Case 6
                Call cmd_Click(0)
            End Select
        End If
        
        Select Case Index
        Case 1, 2, 4, 5, 10, 11, 12, 13
            If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
        Case 7
            'ֻ�������������
            If FilterKeyAscii(KeyAscii, 3) = 0 Then KeyAscii = 0
        Case 0
            If Chr(KeyAscii) = ";" Then KeyAscii = 0
        End Select
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0, 6, 7, 9, 3, 4
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Dim intIndex As Integer
    
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    If Cancel Then Exit Sub
    intIndex = -1
    
    Select Case Index
    Case 1
        intIndex = -1
    Case 2
        intIndex = 0
    Case 3
        intIndex = 2
    Case 5
        intIndex = 1
    Case 10
        intIndex = 2
    Case 11
        intIndex = 3
    Case 12
        intIndex = 5
    Case 13
        intIndex = 6
    End Select
    
    If intIndex >= 0 Then
    
        Select Case Index
        Case 1, 12
            If Trim(txt(Index).Text) = "" Then Exit Sub
        End Select
        
        Cancel = (Val(txt(Index).Text) < udn(intIndex).Min Or Val(txt(Index).Text) > udn(intIndex).Max)
        If Cancel Then
            ShowSimpleMsg "��" & Val(txt(Index).Text) & "�������ˡ�" & udn(intIndex).Min & "��" & udn(intIndex).Max & "����Χ��"
            Exit Sub
        End If
    End If
    
    If (txt(Index).Tag = "Changed") Then
        Select Case Index
        Case 6
            txt(Index).Text = usrSaveItem.Ҫ����Ŀ
        End Select
    End If
    
End Sub

Private Sub txtInfo_Change()
    DataChanged = True
End Sub

Private Sub txtInfo_GotFocus()
    Call zlControl.TxtSelAll(txtInfo)
End Sub

Private Sub txtInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub usrColor_pOK()
    If Val(picDemo.Tag) = usrColor.COLOR Then Exit Sub
    picDemo.Tag = usrColor.COLOR
    Call DrawDemo(picDemo, cbo(0).ListIndex, usrColor.COLOR)
    DataChanged = True
End Sub

Private Function DrawDemo(pic As PictureBox, ByVal bytIndex As Byte, lngColor As Long) As Boolean
    Dim lngStartX As Long
    Dim lngStartY As Long
    
    pic.Cls
    
    lngStartX = (pic.Width - pic.TextWidth("������Ŀ")) / 2
    lngStartY = (pic.Height - pic.TextHeight("������Ŀ") * 3) / 2
    '���ߺͶ�������
    If bytIndex = 0 Or bytIndex = 2 Then
        
        Call DrawDemoChart(pic, lngStartX, lngStartY, lngStartX + pic.TextWidth("������Ŀ"), lngStartY + pic.TextHeight("������Ŀ") * 3, lngColor)
    Else
        Call DrawText(pic, (pic.Width - pic.TextWidth("������Ŀ")) / 2, (pic.Height - pic.TextHeight("������Ŀ")) / 2, "������Ŀ", lngColor)
    End If
    usrColor.COLOR = lngColor
End Function

Private Function DrawDemoChart(pic As PictureBox, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal lngColor As Long) As Boolean
    pic.Cls
    
    DrawLine pic, X1, Y1 + (Y2 - Y1) * 3 / 4, X1 + (X2 - X1) / 6, Y1 + (Y2 - Y1) / 2, lngColor, , 2
    DrawLine pic, X1 + (X2 - X1) / 6, Y1 + (Y2 - Y1) / 2, X1 + (X2 - X1) / 3, Y1 + (Y2 - Y1) * 3 / 4, lngColor, , 2
    DrawLine pic, X1 + (X2 - X1) / 3, Y1 + (Y2 - Y1) * 3 / 4, X1 + (X2 - X1) / 2, Y1 + (Y2 - Y1) / 4, lngColor, , 2
    DrawLine pic, X1 + (X2 - X1) / 2, Y1 + (Y2 - Y1) / 4, X1 + (X2 - X1) * 2 / 3, Y1 + (Y2 - Y1) / 2, lngColor, , 2
    DrawLine pic, X1 + (X2 - X1) * 2 / 3, Y1 + (Y2 - Y1) / 2, X1 + (X2 - X1) * 5 / 6, Y1 + (Y2 - Y1) / 4, lngColor, , 2
    DrawLine pic, X1 + (X2 - X1) * 5 / 6, Y1 + (Y2 - Y1) / 4, X2, Y1 + (Y2 - Y1) * 3 / 4, lngColor, , 2

End Function

Private Sub vsf_AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
    
    If vsf.Body.ColDataType(Col) <> flexDTBoolean And Left(cbo(1).Text, 1) = 0 Then
        If Trim(vsf.TextMatrix(Row, Col)) = "" Then Exit Sub
        If Row <> 4 And Row <> 5 Then
            vsf.TextMatrix(Row, Col) = Format(Val(vsf.TextMatrix(Row, Col)), vsf.Body.ColFormat(Col))
        End If
    End If
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'    If cbo(0).ListIndex = 0 And Val(Vsf.Tag) = 1 And chk(1).Value = 1 Then
'
'        Vsf.EditMode(2) = IIf(NewRow = 3, 1, 0)
'
'        If mlngKey <> -1 Then
'            Vsf.EditMode(2) = IIf(NewRow = 3, 1, 0)
'        Else
'            Vsf.EditMode(2) = 0
'        End If
'    End If
    If cbo(0).ListIndex > 0 Then
        vsf.Body.Editable = flexEDKbdMouse
    End If
'    If (mlngKey = 1 Or mlngKey = 2) Then
'        vsf.Body.Editable = IIf(NewRow = 4, flexEDKbdMouse, flexEDNone)
'    Else
'        vsf.Body.Editable = flexEDKbdMouse
'    End If
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    
    Cancel = (Val(vsf.Tag) = 1)
    
    If mbln������Ŀ And Val(vsf.Tag) = 1 And chk(1).Value = 1 Then
        If mlngKey <> -1 Then
        
            If Row = 3 Or Row = 5 Then
                If Cancel Then vsf.TextMatrix(Row, Col) = "": DataChanged = True
            End If
        End If
    Else
        If Cancel Then
            vsf.TextMatrix(Row, Col) = "": DataChanged = True
        Else
            DataChanged = True
        End If
    End If
    
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf.Tag) = 1)
    
    If Cancel = True Then zlCommFun.PressKey vbKeyTab
        
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
'    If cbo(0).ListIndex = 0 And Val(Vsf.Tag) = 1 And chk(1).Value = 1 Then
'
'        Vsf.EditMode(2) = IIf(NewRow = 3, 1, 0)
'
'        If mlngKey <> -1 Then
'            Vsf.EditMode(2) = IIf(NewRow = 3, 1, 0)
'        Else
'            Vsf.EditMode(2) = 0
'        End If
'
'    End If
End Sub

Private Sub vsf_GotFocus()
'    If cbo(0).ListIndex = 0 And Val(Vsf.Tag) = 1 And chk(1).Value = 1 Then
'
'        Vsf.EditMode(2) = IIf(Vsf.Row = 3, 1, 0)
'
'        If mlngKey <> -1 Then
'            Vsf.EditMode(2) = IIf(Vsf.Row = 3, 1, 0)
'        Else
'            Vsf.EditMode(2) = 0
'        End If
'
'    End If
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If KeyAscii = vbKeyReturn And Col = 1 Then
        If Row = vsf.Rows - 1 Then
            If vsf.TextMatrix(Row, 1) = "" Then
                Cancel = True
                zlCommFun.PressKey vbKeyTab
            End If
        End If
    End If
    
    If cbo(1).ListIndex = 0 And (cbo(2).ListIndex = 0 Or cbo(2).ListIndex = 1) Then
        If KeyAscii <> vbKeyReturn Then
            If (Row = 4 Or Row = 5) And InStr(1, vsf.TextMatrix(Row, 2), ";") = 0 Then
                KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789;.")
            Else
                KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789.")
            End If
        End If
    End If
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    If cbo(1).ListIndex = 0 And (cbo(2).ListIndex = 0 Or cbo(2).ListIndex = 1) Then
        If KeyAscii <> vbKeyReturn Then
            If (Row = 4 Or Row = 5) And InStr(1, vsf.EditText, ";") = 0 Then
                KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789;.")
            Else
                KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789.")
            End If
        End If
    ElseIf cbo(1).ListIndex = 1 And (cbo(2).ListIndex = 1 Or cbo(2).ListIndex = 2) Then
        If InStr("';", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
    
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Dim lngLoop As Long
    
    Select Case Col
    Case 1
        For lngLoop = 1 To vsf.Rows - 1
            If lngLoop <> Row Then
                If Trim(vsf.TextMatrix(lngLoop, 1)) = Trim(vsf.EditText) And Trim(vsf.EditText) <> "" Then
                    ShowSimpleMsg "��" & Trim(vsf.EditText) & " ���Ѿ����ڣ�"
                    Cancel = True
                End If
            End If
        Next
    Case 2
        If Val(Trim(vsf.EditText)) = 1 And vsf.Body.ColDataType(Col) = flexDTBoolean Then
       
            For lngLoop = 1 To vsf.Rows - 1
                If lngLoop <> Row Then
                    If Abs(Val(Trim(vsf.TextMatrix(lngLoop, 2)))) = 1 Then
                        vsf.TextMatrix(lngLoop, 2) = "0"
                    End If
                End If
            Next
        End If
    End Select
    
End Sub
