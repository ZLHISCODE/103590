VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCaseTendBodyPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   7125
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8205
   Icon            =   "frmCaseTendBodyPara.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chk 
      Caption         =   "Ӥ�����µ�����������0��ʼ"
      Height          =   180
      Index           =   16
      Left            =   4995
      TabIndex        =   55
      Top             =   5805
      Width           =   2880
   End
   Begin VB.CheckBox chk 
      Caption         =   "�೦�����Է��ӷ�ĸ��ʾ"
      Height          =   180
      Index           =   15
      Left            =   4995
      TabIndex        =   54
      Top             =   5505
      Width           =   2880
   End
   Begin VB.CheckBox chk 
      Caption         =   "ֻ�ڵ�ǰҳ����ʾ��ҳ���ݣ�������ҳ����ʾ��"
      Height          =   180
      Index           =   13
      Left            =   120
      TabIndex        =   50
      Top             =   6120
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "���µ��ļ��Ŀ�ʼʱ��"
      Height          =   645
      Left            =   5385
      TabIndex        =   56
      Top             =   3825
      Width           =   2790
      Begin VB.OptionButton opt���µ���ʼʱ�� 
         Caption         =   "���ʱ��"
         Height          =   195
         Index           =   1
         Left            =   1605
         TabIndex        =   58
         Top             =   300
         Width           =   1125
      End
      Begin VB.OptionButton opt���µ���ʼʱ�� 
         Caption         =   "��Ժʱ��"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   57
         Top             =   300
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "�����Զ���־��˳���ڵ�������"
      Height          =   180
      Index           =   12
      Left            =   5000
      TabIndex        =   53
      Top             =   5205
      Width           =   2895
   End
   Begin VB.CheckBox chk 
      Caption         =   "���ܡ�������Ŀ��ʾ�������ݣ�������ʾ���죩"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   49
      Top             =   5805
      Width           =   4215
   End
   Begin VB.CheckBox chk 
      Caption         =   "���µ����ʱ��ӡҽԺ����"
      Height          =   180
      Index           =   1
      Left            =   5000
      TabIndex        =   52
      Top             =   4905
      Width           =   2895
   End
   Begin VB.Frame FraSplit 
      Height          =   45
      Left            =   120
      TabIndex        =   62
      Top             =   6540
      Width           =   8055
   End
   Begin VB.Frame fra 
      Caption         =   "����С���־"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox cboNodule 
         Height          =   300
         ItemData        =   "frmCaseTendBodyPara.frx":000C
         Left            =   1320
         List            =   "frmCaseTendBodyPara.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2100
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "С��ȱʡ��ʽ"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "�����ļ�ҳ�밴�ļ���ʽ˳����"
      Height          =   180
      Index           =   11
      Left            =   120
      TabIndex        =   46
      Top             =   4905
      Width           =   3135
   End
   Begin VB.CheckBox chk 
      Caption         =   "�����ļ��Ŵ�ģʽ������������ʾ��׼��С��"
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   47
      Top             =   5205
      Width           =   4095
   End
   Begin VB.CheckBox chk 
      Caption         =   "סԺ����ͬһʱ����Ҫ��¼��ݻ����ļ�"
      Height          =   180
      Index           =   9
      Left            =   120
      TabIndex        =   45
      Top             =   4605
      Width           =   4095
   End
   Begin VB.CheckBox chk 
      Caption         =   "Ӥ�����µ���ʾ��Ժ��Ϣ"
      Height          =   180
      Index           =   5
      Left            =   5000
      TabIndex        =   51
      Top             =   4605
      Width           =   2535
   End
   Begin VB.CheckBox chk 
      Caption         =   "�����������ݴ�ӡ���ʱ������ʾ�������ݼ̳�)"
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   48
      Top             =   5505
      Width           =   4400
   End
   Begin VB.Frame fra 
      Caption         =   "�����Զ���־"
      Height          =   3615
      Index           =   15
      Left            =   5400
      TabIndex        =   61
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   7
         ItemData        =   "frmCaseTendBodyPara.frx":0010
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   2835
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   6
         ItemData        =   "frmCaseTendBodyPara.frx":0014
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   2466
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   0
         ItemData        =   "frmCaseTendBodyPara.frx":0018
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   270
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   1
         ItemData        =   "frmCaseTendBodyPara.frx":001C
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   636
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   2
         ItemData        =   "frmCaseTendBodyPara.frx":0020
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1002
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   3
         ItemData        =   "frmCaseTendBodyPara.frx":0024
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1368
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   4
         ItemData        =   "frmCaseTendBodyPara.frx":0028
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1734
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   5
         ItemData        =   "frmCaseTendBodyPara.frx":002C
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   2100
         Width           =   2100
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   2
         Left            =   135
         TabIndex        =   42
         Top             =   2895
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   40
         Top             =   2526
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ"
         Height          =   180
         Index           =   44
         Left            =   135
         TabIndex        =   28
         Top             =   315
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Index           =   45
         Left            =   135
         TabIndex        =   30
         Top             =   675
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ת��"
         Height          =   180
         Index           =   46
         Left            =   135
         TabIndex        =   32
         Top             =   1062
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   48
         Left            =   135
         TabIndex        =   34
         Top             =   1428
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   49
         Left            =   135
         TabIndex        =   36
         Top             =   1794
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ"
         Height          =   180
         Index           =   50
         Left            =   135
         TabIndex        =   38
         Top             =   2160
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5760
      TabIndex        =   59
      Top             =   6690
      Width           =   1100
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6960
      TabIndex        =   60
      Top             =   6690
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   3750
      Index           =   0
      Left            =   120
      TabIndex        =   44
      Top             =   720
      Width           =   5175
      Begin VB.CheckBox chk 
         Caption         =   "���µ���ӡʱ�������������(�������ʵ���ʹ����Ч)"
         Height          =   315
         Index           =   14
         Left            =   210
         TabIndex        =   27
         Top             =   3345
         Width           =   4770
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   3
         Left            =   2900
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "18"
         Top             =   885
         Width           =   350
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   4
         Left            =   4250
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "6"
         Top             =   885
         Width           =   350
      End
      Begin VB.ComboBox cboSplit 
         Height          =   300
         ItemData        =   "frmCaseTendBodyPara.frx":0030
         Left            =   2400
         List            =   "frmCaseTendBodyPara.frx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1860
         Width           =   900
      End
      Begin VB.CheckBox chk 
         Caption         =   "���µ����ʱ����ʾƤ�Խ��"
         Height          =   315
         Index           =   8
         Left            =   210
         TabIndex        =   26
         Top             =   3060
         Width           =   2790
      End
      Begin VB.CheckBox chk 
         Caption         =   "���µ��Ե�����ʾ(����������˫����ʾ)"
         Height          =   315
         Index           =   7
         Left            =   210
         TabIndex        =   25
         Top             =   2760
         Width           =   3630
      End
      Begin VB.CheckBox chk 
         Caption         =   "���µ�����ʾ���˵������Ϣ"
         Height          =   315
         Index           =   3
         Left            =   210
         TabIndex        =   24
         Top             =   2475
         Width           =   2790
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1665
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "1"
         Top             =   1515
         Width           =   420
      End
      Begin VB.CheckBox chk 
         Caption         =   "δ��˵����ʾ�����µ������棨��������ʱ��ʾ�����棩"
         Height          =   315
         Index           =   2
         Left            =   210
         TabIndex        =   23
         Top             =   2175
         Width           =   4800
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2580
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "0"
         Top             =   1185
         Width           =   375
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "0"
         Top             =   885
         Width           =   370
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   1155
         TabIndex        =   4
         Text            =   "14"
         Top             =   270
         Width           =   255
      End
      Begin VB.CheckBox chk 
         Caption         =   "�������ע�������ٴ�����ʱ,ֹͣǰһ��������ע"
         Height          =   375
         Index           =   0
         Left            =   195
         TabIndex        =   5
         Top             =   525
         Width           =   4500
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   6
         Left            =   2050
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   885
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txt(6)"
         BuddyDispid     =   196622
         BuddyIndex      =   6
         OrigLeft        =   2190
         OrigTop         =   870
         OrigRight       =   2430
         OrigBottom      =   1170
         Max             =   4
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   1
         Left            =   2970
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1185
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txt(1)"
         BuddyDispid     =   196622
         BuddyIndex      =   1
         OrigLeft        =   2190
         OrigTop         =   870
         OrigRight       =   2430
         OrigBottom      =   1170
         Max             =   30
         Min             =   2
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   0
         Left            =   2100
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1515
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txt(2)"
         BuddyDispid     =   196622
         BuddyIndex      =   2
         OrigLeft        =   2190
         OrigTop         =   870
         OrigRight       =   2430
         OrigBottom      =   1170
         Max             =   30
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   3
         Left            =   4580
         TabIndex        =   14
         Top             =   885
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   8
         BuddyControl    =   "txt(4)"
         BuddyDispid     =   196622
         BuddyIndex      =   4
         OrigLeft        =   4580
         OrigTop         =   885
         OrigRight       =   4835
         OrigBottom      =   1170
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   2
         Left            =   3230
         TabIndex        =   11
         Top             =   885
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   18
         BuddyControl    =   "txt(3)"
         BuddyDispid     =   196622
         BuddyIndex      =   3
         OrigLeft        =   3230
         OrigTop         =   885
         OrigRight       =   3485
         OrigBottom      =   1170
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ҹ���       ����"
         Height          =   180
         Index           =   7
         Left            =   2350
         TabIndex        =   9
         Top             =   945
         Width           =   1530
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����       ��"
         Height          =   180
         Index           =   8
         Left            =   3885
         TabIndex        =   12
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����Զ���־��ʱ��֮����           ����"
         Height          =   180
         Index           =   6
         Left            =   210
         TabIndex        =   21
         Top             =   1920
         Width           =   3510
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����¼�볬����ǰ        ��Ļ����¼����"
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   18
         Top             =   1560
         Width           =   3600
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���±����ʱ��������ݹ̶�        ��"
         Height          =   180
         Index           =   3
         Left            =   210
         TabIndex        =   15
         Top             =   1230
         Width           =   3240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���¿�ʼ��¼ʱ��"
         Height          =   180
         Index           =   31
         Left            =   210
         TabIndex        =   6
         Top             =   945
         Width           =   1440
      End
      Begin VB.Line Line1 
         X1              =   1125
         X2              =   1410
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�������ע    ��"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   3
         Top             =   270
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmCaseTendBodyPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mfrmMain As Object
Private mblnOK As Boolean
Private mstrPrivs As String

Public Function ShowPara(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strSQL As String, strPar As String
    Dim curDate As Date, intDay As Integer
    Dim intStart As Integer
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    
    '��ʼ���µ����
    '------------------------------------------------------------------------------------------------------------------
    cboBody(0).Clear
    cboBody(0).AddItem "0-����ʾ"
    cboBody(0).AddItem "1-��ʾ˵��"
    cboBody(0).AddItem "2-��ʾ˵����ʱ��"
    
    cboBody(1).Clear
    cboBody(1).AddItem "0-����ʾ"
    cboBody(1).AddItem "1-��ʾ˵��"
    cboBody(1).AddItem "2-��ʾ˵����ʱ��"
    
    cboBody(2).Clear
    cboBody(2).AddItem "0-����ʾ"
    cboBody(2).AddItem "1-��ʾ˵��"
    cboBody(2).AddItem "2-��ʾ˵����ʱ��"
    cboBody(2).AddItem "3-��ʾ˵���Ϳ���"
    cboBody(2).AddItem "4-��ʾ˵��,����,ʱ��"
    
    cboBody(3).Clear
    cboBody(3).AddItem "0-����ʾ"
    cboBody(3).AddItem "1-��ʾ˵��"
    cboBody(3).AddItem "2-��ʾ˵����ʱ��"
    
    cboBody(4).Clear
    cboBody(4).AddItem "0-����ʾ"
    cboBody(4).AddItem "1-��ʾ˵��"
    cboBody(4).AddItem "2-��ʾ˵����ʱ��"
    
    cboBody(5).Clear
    cboBody(5).AddItem "0-����ʾ"
    cboBody(5).AddItem "1-��ʾ˵��"
    cboBody(5).AddItem "2-��ʾ˵����ʱ��"
    
    cboBody(6).Clear
    cboBody(6).AddItem "0-����ʾ"
    cboBody(6).AddItem "1-��ʾ˵��"
    cboBody(6).AddItem "2-��ʾ˵����ʱ��"
    
    cboBody(7).Clear
    cboBody(7).AddItem "0-����ʾ"
    cboBody(7).AddItem "1-��ʾ˵��"
    cboBody(7).AddItem "2-��ʾ˵����ʱ��"
    
    cboNodule.Clear
    cboNodule.AddItem "0-������"
    cboNodule.AddItem "1-���º���"
    cboNodule.AddItem "2-��˫����"
    cboNodule.AddItem "3-�Ϻ���"
    
    cboSplit.Clear
    cboSplit.AddItem "����"
    cboSplit.AddItem "��"
    
    intStart = zldatabase.GetPara("���µ��ļ���ʼʱ��", glngSys, 1255, 1, Array(opt���µ���ʼʱ��(0), opt���µ���ʼʱ��(1)), InStr(mstrPrivs, "����ѡ������") > 0)
    opt���µ���ʼʱ��(intStart).Value = True
    txt(6).Text = zldatabase.GetPara("���¿�ʼʱ��", glngSys, 1255, 4, Array(txt(6), ud(6), lbl(31)), InStr(mstrPrivs, "����ѡ������") > 0)
    txt(1).Text = zldatabase.GetPara("���±������", glngSys, 1255, 8, Array(txt(1), ud(1), lbl(3)), InStr(mstrPrivs, "����ѡ������") > 0)
    strTmp = zldatabase.GetPara("���µ����", glngSys, 1255, "1;1;1;1;1;1;1:1", Array(cboBody(0), cboBody(1), cboBody(2), cboBody(3), cboBody(4), cboBody(5), cboBody(6), cboBody(7)), InStr(mstrPrivs, "����ѡ������") > 0)
    
    For intLoop = 0 To 7
        If UBound(Split(strTmp, ";")) >= intLoop Then
            cboBody(intLoop).ListIndex = Val(Split(strTmp, ";")(intLoop))
        Else
            cboBody(intLoop).ListIndex = 0
        End If
    Next
    strTmp = zldatabase.GetPara("С��ȱʡ��ʽ", glngSys, 1255, "0", Array(cboNodule, lbl(5)), InStr(mstrPrivs, "����ѡ������") > 0)
    
    If Val(strTmp) >= 0 And Val(strTmp) <= 2 Then
        cboNodule.ListIndex = Val(strTmp)
    Else
        cboNodule.ListIndex = 0
    End If
    
    strTmp = zldatabase.GetPara("���±�־�ָ���", glngSys, 1255, "0", Array(cboSplit, lbl(6)), InStr(mstrPrivs, "����ѡ������") > 0)
    
    If Val(strTmp) >= 0 And Val(strTmp) <= 1 Then
        cboSplit.ListIndex = Val(strTmp)
    Else
        cboSplit.ListIndex = 0
    End If
    
    '����ҹ���־
    strTmp = zldatabase.GetPara("����ʱ��ҹ���־", glngSys, 1255, "18;6", Array(lbl(7), txt(3), ud(2), lbl(8), txt(4), ud(3)), InStr(mstrPrivs, "����ѡ������") > 0)
    If UBound(Split(strTmp, ";")) >= 1 Then
        txt(3).Text = Abs(Val(Split(strTmp, ";")(0)))
        txt(4).Text = Abs(Val(Split(strTmp, ";")(1)))
    Else
         txt(3).Text = Abs(Val(strTmp))
    End If
    
    txt(0).Text = Val(zldatabase.GetPara("�������ע����", glngSys, 1255, "10", Array(txt(0), lbl(0)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(0).Value = Val(zldatabase.GetPara("�ٴ�����ֹͣǰ�α�ע", glngSys, 1255, "0", Array(chk(0)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(1).Value = Val(zldatabase.GetPara("��ӡҽԺ����", glngSys, 1255, "1", Array(chk(1)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(2).Value = Val(zldatabase.GetPara("δ��˵����ʾλ��", glngSys, 1255, "0", Array(chk(2)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(3).Value = Val(zldatabase.GetPara("���µ���ʾ���", glngSys, 1255, "1", Array(chk(3)), InStr(mstrPrivs, "����ѡ������") > 0))
    txt(2).Text = Val(zldatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1", Array(txt(2), lbl(4)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(4).Value = Val(zldatabase.GetPara("����������", glngSys, 1255, "0", Array(chk(4)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(5).Value = Val(zldatabase.GetPara("Ӥ�����µ���ʾ��Ժ��Ϣ", glngSys, 1255, "1", Array(chk(5)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(6).Value = Val(zldatabase.GetPara("���ܲ�����ʾ��������", glngSys, 1255, "1", Array(chk(6)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(7).Value = Val(zldatabase.GetPara("���µ���ʾ��ʽ", glngSys, 1255, "0", Array(chk(7)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(8).Value = Val(zldatabase.GetPara("���µ���ʾƤ�Խ��", glngSys, 1255, "0", Array(chk(8))))
    chk(9).Value = Val(zldatabase.GetPara("��Ӧ��ݻ����ļ�", glngSys, 1255, "0", Array(chk(9)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(10).Value = Val(zldatabase.GetPara("�����ļ���ʾģʽ", glngSys, 1255, "0", Array(chk(10)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(11).Value = Val(zldatabase.GetPara("�����ļ�ҳ�����", glngSys, 1255, "0", Array(chk(11)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(12).Value = Val(zldatabase.GetPara("���±�־��˳��������", glngSys, 1255, "0", Array(chk(12)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(13).Value = Val(zldatabase.GetPara("��ҳ����ֻ��ʾ�ڵ�һҳ", glngSys, 1255, "0", Array(chk(13)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(14).Value = Val(zldatabase.GetPara("���µ�����ӡ������", glngSys, 1255, "0", Array(chk(14)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(15).Value = Val(zldatabase.GetPara("�೦������ʾ��ʽ", glngSys, 1255, "0", Array(chk(15)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(16).Value = Val(zldatabase.GetPara("Ӥ�����µ�����������ʾ0", glngSys, 1255, "0", Array(chk(16)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(11).Enabled = Not CheckPrintDate
    
    Me.Show 1, mfrmMain
    ShowPara = mblnOK
    
End Function

Private Function CheckPrintDate() As Boolean
'---------------------------------------------------------
'����:'��鲡���Ƿ���ڴ�ӡ����,������ھͲ��������û����ļ�ҳ�����
'---------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    CheckPrintDate = False
    
    strSQL = "Select 1 From  ���˻����ӡ  Where ��ӡҳ�� is not null and   Rownum<2"
    Call zldatabase.OpenRecordset(rsTemp, strSQL, "����Ƿ���ڴ�ӡ����")
    If rsTemp.RecordCount > 0 Then
        CheckPrintDate = True
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cboBody_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intStart As Integer
    Dim strTmp As String
    
    If opt���µ���ʼʱ��(0).Value Then
        intStart = 0
    Else
        intStart = 1
    End If
    
    strTmp = cboBody(0).ListIndex & ";" & cboBody(1).ListIndex & ";" & cboBody(2).ListIndex & ";" & cboBody(3).ListIndex & ";" & cboBody(4).ListIndex & ";" & cboBody(5).ListIndex & ";" & cboBody(6).ListIndex & ";" & cboBody(7).ListIndex
    Call zldatabase.SetPara("���¿�ʼʱ��", Val(txt(6).Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("���±������", Val(txt(1).Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("���µ����", strTmp, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("�������ע����", Val(txt(0).Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("����¼�뻤����������", Val(txt(2).Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("�ٴ�����ֹͣǰ�α�ע", chk(0).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("δ��˵����ʾλ��", chk(2).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("���µ���ʾ���", chk(3).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("��ӡҽԺ����", chk(1).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("����������", chk(4).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("Ӥ�����µ���ʾ��Ժ��Ϣ", chk(5).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("���ܲ�����ʾ��������", chk(6).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("���µ���ʾ��ʽ", chk(7).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("���µ���ʾƤ�Խ��", chk(8).Value, glngSys, 1255)
    Call zldatabase.SetPara("��Ӧ��ݻ����ļ�", chk(9).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("�����ļ���ʾģʽ", chk(10).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("�����ļ�ҳ�����", chk(11).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("С��ȱʡ��ʽ", Val(cboNodule.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("���±�־�ָ���", Val(cboSplit.ListIndex), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("���±�־��˳��������", chk(12).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("���µ��ļ���ʼʱ��", intStart, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("����ʱ��ҹ���־", txt(3).Text & ";" & txt(4).Text, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("��ҳ����ֻ��ʾ�ڵ�һҳ", chk(13).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("���µ�����ӡ������", chk(14).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("�೦������ʾ��ʽ", chk(15).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zldatabase.SetPara("Ӥ�����µ�����������ʾ0", chk(16).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    
    mblnOK = True
    
    Unload Me
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt(Index))
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

