VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInNurseRoutine 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "����������"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15270
   Icon            =   "frmInNurseRoutine.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleMode       =   0  'User
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   960
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   31
      Top             =   3000
      Width           =   855
   End
   Begin VB.PictureBox picPrompt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1410
      ScaleHeight     =   285
      ScaleWidth      =   11865
      TabIndex        =   29
      Top             =   7530
      Width           =   11865
      Begin VB.Label lblPrompt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   30
         TabIndex        =   30
         Top             =   60
         Width           =   10500
      End
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5955
      Left            =   2115
      ScaleHeight     =   5925
      ScaleWidth      =   5145
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   585
      Visible         =   0   'False
      Width           =   5175
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   5475
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   5160
         _Version        =   589884
         _ExtentX        =   9102
         _ExtentY        =   9657
         _StockProps     =   0
         BorderStyle     =   1
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   3990
         Picture         =   "frmInNurseRoutine.frx":18F2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "ȷ��"
         Top             =   5550
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterCancel 
         Height          =   315
         Left            =   4530
         Picture         =   "frmInNurseRoutine.frx":1E7C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "ȡ��"
         Top             =   5550
         Width           =   450
      End
   End
   Begin VB.PictureBox picCondition 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1320
      ScaleHeight     =   345
      ScaleWidth      =   9990
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   225
      Width           =   9990
      Begin VB.CommandButton cmdWarrant 
         Caption         =   "����"
         Height          =   270
         Left            =   9105
         TabIndex        =   9
         ToolTipText     =   "������Ϣ����"
         Top             =   45
         Width           =   500
      End
      Begin VB.PictureBox picסԺ���� 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7665
         ScaleHeight     =   225
         ScaleWidth      =   1335
         TabIndex        =   7
         Top             =   60
         Width           =   1365
         Begin VB.ComboBox cboPages 
            BackColor       =   &H00EAFFFF&
            Height          =   300
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   -45
            Width           =   1425
         End
      End
      Begin VB.PictureBox pic���� 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   810
         ScaleHeight     =   315
         ScaleWidth      =   1725
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   1755
         Begin VB.TextBox txt���� 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAFFFF&
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   0
            MaxLength       =   100
            TabIndex        =   3
            Top             =   70
            Width           =   1335
         End
         Begin VB.Image img�����б� 
            Height          =   360
            Left            =   1350
            Picture         =   "frmInNurseRoutine.frx":2406
            Tag             =   "�������������в����б�"
            Top             =   -30
            Width           =   360
         End
      End
      Begin VB.PictureBox pic��ʶ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00EAFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3510
         ScaleHeight     =   345
         ScaleWidth      =   3990
         TabIndex        =   4
         Top             =   0
         Width           =   3990
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����������������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1950
            TabIndex        =   6
            Top             =   60
            Width           =   2040
         End
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��:��һ��-173"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   1755
         End
      End
      Begin VB.Image img��һ�� 
         Height          =   360
         Left            =   2940
         Picture         =   "frmInNurseRoutine.frx":2B08
         Tag             =   "��һ������"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image img��һ�� 
         Height          =   360
         Left            =   2580
         Picture         =   "frmInNurseRoutine.frx":320A
         Tag             =   "��һ������"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image img��ϸ��Ϣ 
         Height          =   360
         Left            =   9630
         Picture         =   "frmInNurseRoutine.frx":390C
         Tag             =   "�鿴������ϸ��Ϣ"
         Top             =   0
         Width           =   360
      End
      Begin VB.Label lbl��λ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   75
         TabIndex        =   1
         Top             =   90
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   28
      Top             =   7455
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmInNurseRoutine.frx":400E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "������ɫ"
            TextSave        =   "������ɫ"
            Key             =   "������ɫ"
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
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   5265
      Left            =   255
      TabIndex        =   27
      Top             =   2085
      Width           =   7770
      _Version        =   589884
      _ExtentX        =   13705
      _ExtentY        =   9287
      _StockProps     =   64
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1215
      ScaleWidth      =   20010
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   750
      Width           =   20010
      Begin VB.Frame fraInfo 
         BackColor       =   &H00EAFFFF&
         Caption         =   "������ϸ��Ϣ"
         Height          =   975
         Left            =   150
         TabIndex        =   15
         Top             =   90
         Width           =   16965
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   9120
            Style           =   2  'Dropdown List
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   570
            Width           =   4845
         End
         Begin VB.Label lblסԺ�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ��:12345678"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   16
            Top             =   300
            Width           =   1350
         End
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "һ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2220
            TabIndex        =   17
            Top             =   300
            Width           =   390
         End
         Begin VB.Label lbl�Ա� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   2910
            TabIndex        =   18
            Top             =   300
            Width           =   195
         End
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "32��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   3330
            TabIndex        =   19
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lbl�������� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������ҽ��[YBZH0001]"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   12600
            TabIndex        =   23
            Top             =   300
            Width           =   1800
         End
         Begin VB.Label lbl����ȼ� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "һ������"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   4590
            TabIndex        =   20
            Top             =   300
            Width           =   720
         End
         Begin VB.Label lbl��Ժʱ�� 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "yyyy-MM-dd HH:mm��yyyy-MM-dd HH:mm"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   6030
            TabIndex        =   21
            Top             =   300
            Width           =   3060
         End
         Begin VB.Label lblҽ�Ƹ��ʽ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������ҽ�Ʊ���"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   9945
            TabIndex        =   22
            Top             =   300
            Width           =   1440
         End
         Begin VB.Label lbl��� 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���:����֧�����ס�����֧�����ס�����֧�����ס�����֧��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   24
            Top             =   630
            Width           =   6060
         End
         Begin VB.Label lblҩ����� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ҩ�����:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   8100
            TabIndex        =   25
            Top             =   630
            Width           =   810
         End
      End
   End
   Begin MSComctlLib.ImageList imgRPT 
      Left            =   -165
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":48A0
            Key             =   "Pati"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":4E3A
            Key             =   "Notify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":53D4
            Key             =   "�ȴ����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":596E
            Key             =   "�ܾ����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":5F08
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":64A2
            Key             =   "���ڳ��"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":6EB4
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":78C6
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":7E60
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":8872
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":9284
            Key             =   "���鵵"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":FAE6
            Key             =   "δ����"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":10080
            Key             =   "ִ����"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":1061A
            Key             =   "������"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":1102C
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":115C6
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":11B60
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":120FA
            Key             =   "������"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":1895C
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":18EF6
            Key             =   "����"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":19490
            Key             =   "����"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":1FCF2
            Key             =   "Ů��"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   270
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmInNurseRoutine.frx":26554
      Left            =   705
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmInNurseRoutine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum PATI_TYPE
    pt��Ժ����ס = 0
    ptת�ƴ���ס = 1
    ptת��������ס = 2
    pt��Ժ = 3
'    pt��ͥ���� = 3.1
'    ptԤת�� = 3.2
'    ptת���� = 3.3
    ptԤ�� = 4
    pt��Ժ = 5
    pt���� = 6
    pt���ת�� = 7
End Enum
Private Enum PATI_COLUMN
    c_ͼ�� = 0
    C_״̬ = 1
    c_���� = 2
    C_����ID = 3
    C_��ҳID = 4
    c_���� = 5
    c_סԺ�� = 6
    c_��Ժ���� = 7
    c_��Ժ���� = 8
    c_�������� = 9
End Enum

Private Enum PATI_COLWIDTH
    cw_ͼ�� = 18
    cw_״̬ = 0
    cw_���� = 40
    Cw_����ID = 0
    cw_��ҳID = 0
    cw_���� = 60
    cw_סԺ�� = 60
    cw_��Ժ���� = 70
    cw_��Ժ���� = 70
    cw_�������� = 100
End Enum

Private mblnShow As Boolean
Private mblnAdd As Boolean
Private mobjBar As CommandBar

'�Ӵ��������
Private mclsEMR As Object  '�°没��zlRichEMR.clsDockEMR
Private WithEvents mclsAdvices As zlPublicAdvice.clsDockInAdvices
Attribute mclsAdvices.VB_VarHelpID = -1
Private WithEvents mclsEPRs As zlRichEPR.cDockInEPRs
Attribute mclsEPRs.VB_VarHelpID = -1
Private WithEvents mclsTends As zl9TendFile.clsTendFile
Attribute mclsTends.VB_VarHelpID = -1
Private mclsTendEPRs As zlRichEPR.cDockInTendEPRs
Attribute mclsTendEPRs.VB_VarHelpID = -1
Private WithEvents mclsFeeQuery As zl9InExse.clsFeeQuery
Attribute mclsFeeQuery.VB_VarHelpID = -1
'Private WithEvents mfrmResponse As frmAuditResponse '��鷴������
Private WithEvents mclsPath As zlPublicPath.clsDockPath
Attribute mclsPath.VB_VarHelpID = -1
Private mclsWardMonitor As clsWardMonitor     '�໤�ǽӿ�

Private mcolSubForm As Collection
Private mcolSubFormOperation As Collection
Private mfrmActive As Form
Private mobjMipModule As Object
'�����������
Private mobjParent As Object
Private mstrPrivs As String
Private mstrPage As String
Private mPatiInfo As PatiInfo '��ʷסԺ��¼�е�,��һ��Ϊ��ǰ��
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng����ID As Long
Private mstrScope As String
Private mintChange As Integer
Private mdtOutBegin As Date, mdtOutEnd As Date
Private mintPrePage As Integer
Private mblnUnRefresh As Boolean
Private mblnRefreshBar As Boolean
Private mlngRowIndex As Long

Private mblnMonitor As Boolean '�໤�ǳ����Ƿ����
Private mstrMonitor As String '�໤�ǳ���·��

Private mbytSize As Byte

Private mrsPati As New ADODB.Recordset
Private mblnTabTmp As Boolean
Private mlngӤ������ID As Long
Private mlngӤ������ID As Long

'���廤����ر���
Private mstrNurseParentID As String '���廤���еĲ���ID
Private mstrRelatedUnitID As String '���廤����ID
Private mstrRelatedUserID As String '���廤����ԱID
Private marrTabAttribute '�洢ÿһ��tab������ֵ(0-��ͨҳ��;1-���廤��ҳ��)
Private mColNurseFormUrl As Collection  '���廤����url��Ϣ
Private mobjNurseForm As Object '���廤���壨���ҳ����һ�����壬ÿ���л�����ж�غʹ���,��Ҫ��Ϊ���ͷ��ڴ棩

Public Sub zlInitMip(ByVal objMipModule As Object)
    '��Ϣ����
    Set mobjMipModule = objMipModule
End Sub

Public Sub NurseRoutine(ByVal frmParent As Object, ByVal strPrivs As String, ByVal lng����ID As Long, ByVal lng����ID As Long, _
    ByVal dtOutBegin As Date, ByVal dtOutEnd As Date, ByVal intChange As Integer, ByVal strScope As String, tPati As PatiInfo, _
    Optional ByVal strPage As String = "ҽ��", Optional ByVal rsThis As ADODB.Recordset, Optional ByVal bytSize As Byte = 0)
    
    If lng����ID = 0 Then Exit Sub
    
    Set mobjParent = frmParent
    mstrPrivs = strPrivs
    mstrPage = strPage
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mdtOutBegin = dtOutBegin
    mdtOutEnd = dtOutEnd
    mintChange = intChange
    mstrScope = strScope
    mPatiInfo = tPati
    mintPrePage = -1            'ÿ���л�����ʱ���
    mblnAdd = Not mblnShow
    mbytSize = bytSize
    
    Call RefreshPatiList(rsThis)
    
    If mblnShow Then
        mintPrePage = -1
        Call AddPages
        Exit Sub
    End If
    Call ReSetFontSize
    mblnShow = True
    mobjParent.mblnRoutine = mblnShow
    Me.Show , frmParent
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '����:������
    '����:2012-06-20 15:15:00
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    Dim lngCol As Long
    Dim PATI_COLWIDTH As Variant
    bytFontSize = IIf(mbytSize = 0, 9, IIf(mbytSize = 1, 12, mbytSize))
    
    Me.FontSize = bytFontSize
    Me.FontName = "����"
    
    'CommandBars
    Set CtlFont = cbsMain.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set cbsMain.Options.Font = CtlFont
    'DockingPane
    Set CtlFont = DkpMain.PaintManager.CaptionFont
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set DkpMain.PaintManager.CaptionFont = CtlFont
    'TabControl
    Set CtlFont = tbcSub.PaintManager.Font
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set tbcSub.PaintManager.Font = CtlFont
            
    lbl��λ.FontSize = Me.FontSize
    lbl��λ.Top = pic����.Top + (pic����.Height - lbl��λ.Height) \ 2
    lbl��λ.Left = 60
    txt����.FontSize = Me.FontSize
    txt����.Top = (pic����.Height - txt����.Height)
    pic����.Left = lbl��λ.Left + lbl��λ.Width + 20
    img��һ��.Left = pic����.Left + pic����.Width
    img��һ��.Left = img��һ��.Left + img��һ��.Width
    pic��ʶ.Left = img��һ��.Left + img��һ��.Width + TextWidth("��")
    Me.pic��ʶ.Width = lbl����.Width + lbl����.Left
    picסԺ����.Left = pic��ʶ.Left + pic��ʶ.Width + 50
    cboPages.FontSize = Me.FontSize
    cboPages.Left = -30
    cboPages.Top = -30
    picסԺ����.Height = cboPages.Height - 20
    picסԺ����.Top = (picCondition.Height - picסԺ����.Height) \ 2
    If picסԺ����.Top < 0 Then picסԺ����.Top = 0
    Me.picסԺ����.Width = Me.cboPages.Width - 50
    
    cmdWarrant.FontSize = Me.FontSize
    cmdWarrant.Width = TextWidth(cmdWarrant.Caption & "��")
    cmdWarrant.Height = TextWidth("��") + TextWidth("��") \ 2
    cmdWarrant.Left = picסԺ����.Left + picסԺ����.Width + 50
    cmdWarrant.Top = (picCondition.Height - cmdWarrant.Height) \ 2
    img��ϸ��Ϣ.Left = cmdWarrant.Left + cmdWarrant.Width + 100
    picCondition.Width = img��ϸ��Ϣ.Left + img��ϸ��Ϣ.Width + 60
    
    '����ѡ��
    Set CtlFont = rptPati.PaintManager.CaptionFont
    CtlFont.Size = bytFontSize
    Set rptPati.PaintManager.CaptionFont = CtlFont
    
    Set CtlFont = rptPati.PaintManager.TextFont
    CtlFont.Size = bytFontSize
    Set rptPati.PaintManager.TextFont = CtlFont
    
    PATI_COLWIDTH = Array(cw_ͼ��, cw_״̬, cw_����, Cw_����ID, cw_��ҳID, cw_����, cw_סԺ��, cw_��Ժ����, cw_��Ժ����, cw_��������)
    For lngCol = C_״̬ To rptPati.Columns.Count - 1
        rptPati.Columns.Column(lngCol).Width = PATI_COLWIDTH(lngCol) + (PATI_COLWIDTH(lngCol) * IIf(mbytSize = 0, 0, 1)) / 3
    Next lngCol
    
    rptPati.Redraw
    
    '������Ϣ��
    fraInfo.FontSize = bytFontSize
    lblסԺ��.FontSize = bytFontSize
    lblסԺ��.Height = TextHeight("��")
    lbl����.FontSize = bytFontSize
    lbl����.Height = TextHeight("��")
    lbl�Ա�.FontSize = bytFontSize
    lbl�Ա�.Height = TextHeight("��")
    lbl����.FontSize = bytFontSize
    lbl����.Height = TextHeight("��")
    lbl����ȼ�.FontSize = bytFontSize
    lbl����ȼ�.Height = TextHeight("��")
    lbl��Ժʱ��.FontSize = bytFontSize
    lbl��Ժʱ��.Height = TextHeight("��")
    lblҽ�Ƹ��ʽ.FontSize = bytFontSize
    lblҽ�Ƹ��ʽ.Height = TextHeight("��")
    lbl��������.FontSize = bytFontSize
    lbl��������.Height = TextHeight("��")
    lbl���.FontSize = bytFontSize
    lbl���.Height = TextHeight("��")
    lblҩ�����.FontSize = bytFontSize
    lblҩ�����.Height = TextHeight("��")
    cbo����.FontSize = bytFontSize
    cbo����.Left = lblҩ�����.Left + lblҩ�����.Width + TextHeight("��")
    
    lblPrompt.FontSize = bytFontSize
    Call Form_Resize
End Sub

'55430:������,2013-02-27,˫������ҽ����λ�����������ҽ��ҳ��
Public Sub OrientTabPage(ByVal strTab As String, Optional ByVal strID As String = "")
'-------------------------------------------------------------
'����:��λ������������ָ����ҳ��,�Լ���Ӧҳ��ָ�����ļ���ҽ����
'-------------------------------------------------------------
    Dim intIdx As Integer
    Dim blnSeek As Boolean
    
    blnSeek = False
    If strTab = tbcSub.Tag Then blnSeek = True
    If blnSeek = False Then
        For intIdx = 0 To tbcSub.ItemCount - 1
            If tbcSub(intIdx).Visible And tbcSub(intIdx).Tag = strTab Then Exit For
        Next
        If intIdx <= tbcSub.ItemCount - 1 Then
            blnSeek = True
            mblnAdd = False
            tbcSub.Item(intIdx).Tag = strTab
            tbcSub.Item(intIdx).Selected = True
        End If
    End If
    '��λҳ��ɹ�,��λ�������λ��
    If blnSeek = True Then
        Select Case strTab
            Case "ҽ��"
                If Val(strID) = 0 Then Exit Sub
                Call mclsAdvices.zlSeekAndViewEPRReport(Val(strID))
            Case "����"
        End Select
    End If
End Sub

Public Sub RefreshPatiList(Optional ByVal rsThis As ADODB.Recordset)
    On Error GoTo ErrHand
    
    'ˢ�²����嵥,�Զ�λ����ǰ�����Ĳ�����
    Call LoadPatient(rsThis)
    mrsPati.Filter = "����ID=" & mlng����ID
    '54408:������,2012-10-10,ͬ�������Ҳ������˾Ͷ�λ����һ������
    '�磺��Ժ���˽��벡�����Ȼ���������潫�˲��˳�Ժ��������˳�Ժʱ�䲻�ڲ�ѯ��Ժ��Χ�ڿ��ܾͻ���ִ����
    If mrsPati.RecordCount = 0 Then
        mrsPati.Filter = "": mrsPati.MoveFirst
        mlng����ID = Val(mrsPati!����ID)
    End If
    mlng��ҳID = Val(mrsPati!��ҳID)
    mlngӤ������ID = Val(mrsPati!Ӥ������ID & "")
    mlngӤ������ID = Val(mrsPati!Ӥ������ID & "")
    mrsPati.Filter = ""
    mrsPati.MoveFirst
    '90592:ͬһ���˿��ܴ��ڶ�����¼����״̬��ͬ�����ղ���ID����ҳID����
    mrsPati.Find ("Key='" & mlng����ID & ":" & mlng��ҳID & "'")
    rptPati.Records.DeleteAll
    picPati.Visible = False
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LocatePati(ByVal intType As Integer)
    '����˵��:intType:1-��һ������;2-��һ������
    '���˷�Χ:�ڴ�����ѭ��,���ϰ汣��һ��
    Dim blnExit As Boolean  'ǿ���˳�
    On Error Resume Next
    
redo:
    If intType = 1 Then
        mrsPati.MovePrevious
        If mrsPati.BOF Then mrsPati.MoveLast
    Else
        mrsPati.MoveNext
        If mrsPati.EOF Then mrsPati.MoveFirst
    End If
    If mrsPati!����ID <> 0 Then
        If mrsPati!����ID <> mlng����ID Then
            mlng����ID = mrsPati!����ID
            mlng��ҳID = mrsPati!��ҳID
            mintPrePage = -1
            Call AddPages
        Else
            If blnExit Then Exit Sub
            blnExit = True
            GoTo redo
        End If
    Else
        GoTo redo
    End If
    
    picPati.Visible = False
End Sub

Private Sub cmdFilterCancel_Click()
    picPati.Visible = False
End Sub

Private Sub cmdFilterOK_Click()
    Call rptPati_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub cmdWarrant_Click()
    Call frmPatiSurety.ShowMe(Me, mlng����ID, mlng��ҳID)
End Sub

Private Sub Form_Activate()
    picPrompt.Visible = Me.stbThis.Visible
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("[']", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Function GetVersion() As String
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    strSQL = " select �汾�� from zlsystems where ���=100"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��׼�汾��")
    GetVersion = rsTemp!�汾��
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub LoadPatient(Optional ByVal rsThis As ADODB.Recordset)
    'U10.32��ʼ֧�ֶ�����
    Dim strSQL As String
    Dim strBriefCode As String
    Dim blnSupport As Boolean
    Dim strField As String, strValue As String
    Dim rsPati As New ADODB.Recordset
    On Error GoTo ErrHand
    
    blnSupport = (Val(Split(GetVersion, ".")(1)) >= 32)
    If blnSupport Then
        strBriefCode = ",zlpinyincode(NVL(B.����,A.����),0,0,',',1) AS ���� "
    Else
        strBriefCode = ",zlspellcode(NVL(B.����,A.����)) AS ����"
    End If
    
    '54408:������,2012-10-10,��������ҳ�Ժ���˿�������Чʱ�䷶Χ֮��
    strField = "Key," & adLongVarChar & ",50|����," & adDouble & ",2|����2," & adDouble & ",2|����," & adLongVarChar & ",50|����ID," & adDouble & ",18|��ҳID," & adDouble & ",18|" & _
           "סԺ��," & adDouble & ",18|����," & adLongVarChar & ",20|����," & adLongVarChar & ",200|�Ա�," & adLongVarChar & ",10|����," & adLongVarChar & ",20|����," & adLongVarChar & ",50|" & _
           "����ID," & adDouble & ",18|סԺҽʦ," & adLongVarChar & ",20|���λ�ʿ," & adLongVarChar & ",20|����״̬," & adLongVarChar & ",20|" & _
           "����," & adLongVarChar & ",20|����ȼ�," & adLongVarChar & ",50|�ѱ�," & adLongVarChar & ",50|��ǰ����," & adLongVarChar & ",50|" & _
           "��Ժ����," & adLongVarChar & ",20|��Ժ����," & adLongVarChar & ",20|סԺ����," & adLongVarChar & ",20|��Ժ��ʽ," & adLongVarChar & ",20|" & _
           "��������," & adLongVarChar & ",50|״̬," & adLongVarChar & ",10|����," & adDouble & ",18|���￨��," & adLongVarChar & ",20|·��״̬," & adLongVarChar & ",20|" & _
           "��ɫ," & adDouble & ",18|������," & adLongVarChar & ",10|Ӥ������ID," & adDouble & ",18|Ӥ������ID," & adDouble & ",18"
    Call Record_Init(mrsPati, strField)
    
'    If rsThis Is Nothing Then
        '��Ժ����ƺ�ת�ƴ���Ʋ���(���˿��������Ĳ������ɽ���)
        'c.����id + 0,˵����ͨ��H����������ӹ��˺󣬼�¼�������٣�������B�������
        If Val(Mid(mstrScope, 5, 1)) <> 0 Then
            '84938:�����ɣ������Ż�(�������:A.��ҳID=B.��ҳID)
            strSQL = _
                "Select /*+ RULE */Distinct" & vbNewLine & _
                " Decode(B.״̬,1,0,Decode(c.��ʼԭ��,3,1,2)) As ����, Decode(Nvl(b.����״̬, 0), 0, 999, b.����״̬) As ����2," & _
                " Decode(B.״̬,1,'��Ժ����ס����',Decode(c.��ʼԭ��,3,'ת�ƴ���ס����','ת��������ס����')) As ����," & _
                " a.����id, b.��ҳid, A.�����,B.סԺ��, NVL(b.����,a.����) ����" & strBriefCode & ", NVL(b.�Ա�,a.�Ա�) �Ա�, NVL(b.����,a.����) ����," & vbNewLine & _
                " d.���� As ����, c.����id, c.����ҽʦ As סԺҽʦ,b.���λ�ʿ, b.����״̬, lpad(c.����,10,' ') AS ����," & _
                " e.���� As ����ȼ�, b.�ѱ�,b.��ǰ����, Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) AS ��Ժ���� , b.��Ժ����,B.��Ժ��ʽ, b.��������, b.״̬, b.����, a.���￨��," & vbNewLine & _
                " -1 As ·��״̬,trunc(sysdate)-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��))+1 as סԺ����,Z.��ɫ,B.������,B.Ӥ������ID,B.Ӥ������ID" & vbNewLine & _
                "From ������Ϣ A, ������ҳ B, ���˱䶯��¼ C, ���ű� D, �շ���ĿĿ¼ E, �������� Z" & vbNewLine & _
                "Where B.��������=Z.����(+) And a.��Ժ = 1 And a.����id = b.����id And A.��ҳID=B.��ҳID And Nvl(b.��ҳid, 0) <> 0 And b.����id = c.����id And b.��ҳid = c.��ҳid " & vbNewLine & _
                "      And (C.����ID=[1] or C.����ID is null) And c.����id = d.Id" & vbNewLine & _
                "      And (d.վ��='" & gstrNodeNo & "' Or d.վ�� is Null)" & vbNewLine & _
                "      And b.����ȼ�id = e.Id(+) And Nvl(c.���Ӵ�λ, 0) = 0 And c.��ֹʱ�� Is Null" & vbNewLine & _
                "      And (c.��ʼԭ�� in(1,3) And Exists(Select 1 From �������Ҷ�Ӧ H Where c.����id = h.����id And h.����id = [1]) or c.��ʼԭ��=15 And c.����id = [1])" & vbNewLine & _
                "      And ((c.��ʼԭ�� = 1 And b.״̬ = 1) Or (c.��ʼԭ�� in (3,15) And c.��ʼʱ�� Is Null And b.״̬ = 2)) "
        End If
        '��Ժ����(�̶�ǿ����ʾ)
        strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
            "Select /*+ RULE */ Decode(B.״̬,3,4,DECODE(B.��Ժ����, NULL, 3.1,DECODE(B.״̬,2,3.2,3))) as ����," & _
            " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2," & _
            " Decode(B.״̬,3,'Ԥ��Ժ����',DECODE(B.��Ժ����, NULL, '��ͥ����',DECODE(B.״̬,2,'Ԥת�Ʋ���', '��Ժ����'))) as ����," & _
            " A.����ID,B.��ҳID,A.�����,B.סԺ��,NVL(B.����,A.����) ����" & strBriefCode & ",NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,C.���� as ����,B.��Ժ����ID ����ID,B.סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
            " lpad(B.��Ժ����,10,' ') AS ����,E.���� as ����ȼ�,B.�ѱ�,B.��ǰ����,Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) AS ��Ժ���� ,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
            " B.״̬,B.����,A.���￨��,Nvl(b.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��))+1 as סԺ����,z.��ɫ,B.������,B.Ӥ������ID,B.Ӥ������ID" & _
            " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ E,�������� Z,��Ժ���� R" & _
            " Where B.��������=Z.����(+) And A.����ID=B.����ID And A.��ҳID=B.��ҳID  And Nvl(B.״̬,0)<>1" & _
            " And B.��Ժ����ID=C.ID And B.����ȼ�ID=E.ID(+) And (R.����ID=[1] Or b.Ӥ������ID=[1]) And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
            " And a.����ID=R.����ID And A.��ǰ����ID=R.����ID And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL"
        '��Ժ����:��Ժ���˿������ж��סԺ
        If Val(Mid(mstrScope, 2, 1)) <> 0 Then
            strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
                "Select /*+ RULE */ Decode(B.��Ժ��ʽ,'����',6,5) as ����," & _
                " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2," & _
                " Decode(B.��Ժ��ʽ,'����','��������','��Ժ����') as ����," & _
                " A.����ID,B.��ҳID,A.�����,B.סԺ��,NVL(B.����,A.����) ����" & strBriefCode & ",NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,C.���� as ����,B.��Ժ����ID ����ID,B.סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
                " lpad(B.��Ժ����,10,' ') AS ����,E.���� as ����ȼ�,B.�ѱ�,B.��ǰ����,Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) AS ��Ժ���� ,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
                " B.״̬,B.����,A.���￨��,Nvl(b.·��״̬,-1) ·��״̬,trunc(b.��Ժ����)-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��))+1 as סԺ����,z.��ɫ,B.������,B.Ӥ������ID,B.Ӥ������ID" & _
                " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ E,�������� Z" & _
                " Where B.��������=Z.����(+) And A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.״̬=0" & _
                " And B.��Ժ����ID=C.ID And B.����ȼ�ID=E.ID(+) And B.��ǰ����ID+0=[1] And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
                " And B.��Ժ���� Between [2] And [3] And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL"
        End If
        'ת������:��Ժ,ҽ���ʹ�����ʾ����ת��ǰ��
        If Val(Mid(mstrScope, 4, 1)) <> 0 Then
            strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
                "Select /*+ RULE */ Distinct 7 as ����,Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,'ת������' as ����," & _
                " A.����ID,B.��ҳID,A.�����,B.סԺ��,NVL(B.����,A.����) ����" & strBriefCode & ",NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,D.���� as ����,C.����ID,C.����ҽʦ as סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
                " lpad(c.����,10,' ') AS ����,E.���� as ����ȼ�,B.�ѱ�,B.��ǰ����,Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) AS ��Ժ���� ,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
                " B.״̬,B.����,A.���￨��,Nvl(b.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��))+1 as סԺ����,z.��ɫ,B.������,B.Ӥ������ID,B.Ӥ������ID" & _
                " From ������Ϣ A,������ҳ B,���˱䶯��¼ C,���ű� D,�շ���ĿĿ¼ E,�������� Z" & _
                " Where B.��������=Z.����(+) And A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.����ȼ�ID=E.ID(+)" & _
                " And B.����ID=C.����ID And B.��ҳID=C.��ҳID" & _
                " And B.��ǰ����ID<>[1] And C.����ID+0=[1] And C.����ID=D.ID" & _
                " And Nvl(C.���Ӵ�λ,0)=0 And C.��ֹԭ�� In(3,15) And C.��ֹʱ�� Between Sysdate-[7] And Sysdate" & _
                " And Nvl(B.״̬,0)<>2 And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL"
        End If
        strSQL = strSQL & " Order by ����,����,��ҳID Desc"
        
        Screen.MousePointer = 11
        On Error GoTo ErrHand
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����б�", mlng����ID, _
            CDate(Format(mdtOutBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(mdtOutEnd, "yyyy-MM-dd 23:59:59")), _
            Val(Mid(mstrScope, 1, 1)), Val(Mid(mstrScope, 2, 1)), Val(Mid(mstrScope, 5, 1)), mintChange)
        
        '��ʼװ�ز�����Ϣ
        rsPati.Filter = 0
        Call CopyReocrd(rsPati)
        
        'ͨ����������ֱ�Ӳ��ҵĳ�Ժ���˿��ܲ��ڳ�Ժ��Χ�ڣ��˴���Ҫ���¼���
        If rsThis Is Nothing Then Exit Sub
        If rsThis.State = adStateClosed Then Exit Sub
        rsThis.Filter = "����=5 or ����=6 or ����=7"
        Call CopyReocrd(rsThis)
        '���½��в�������
        mrsPati.Sort = "����,����,��ҳID Desc"
'    Else
'        Set mrsPati = rsThis.Clone
'    End If
    Screen.MousePointer = 0
    Exit Sub
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CopyReocrd(ByVal rsPati As ADODB.Recordset)
    '54408:������,2012-10-10
    Dim strField As String, strValue As String
    
    If rsPati.RecordCount <> 0 Then rsPati.MoveFirst
    strField = "Key|����|����2|����|����ID|��ҳID|סԺ��|����|����|�Ա�|����|����|����ID|סԺҽʦ|���λ�ʿ|����״̬|����|����ȼ�|�ѱ�|��ǰ����|��Ժ����|��Ժ����|סԺ����|��Ժ��ʽ|��������|״̬|����|���￨��|·��״̬|��ɫ|������|Ӥ������ID|Ӥ������ID"
    Do While Not rsPati.EOF
        mrsPati.Filter = "Key='" & Val(rsPati!����ID) & ":" & Val(rsPati!��ҳID) & "'"
        If mrsPati.RecordCount = 0 Then
            strValue = Val(rsPati!����ID) & ":" & Val(rsPati!��ҳID) & "|" & rsPati!���� & "|" & rsPati!����2 & "|" & rsPati!���� & "|" & rsPati!����ID & "|" & rsPati!��ҳID & "|" & NVL(rsPati!סԺ��, 0) & "|" & rsPati!���� & "|" & rsPati!���� & "|" & rsPati!�Ա� & "|" & _
                      rsPati!���� & "|" & NVL(rsPati!����) & "|" & NVL(rsPati!����ID, 0) & "|" & NVL(rsPati!סԺҽʦ) & "|" & NVL(rsPati!���λ�ʿ) & "|" & NVL(rsPati!����״̬, 0) & "|" & NVL(rsPati!����) & "|" & _
                      NVL(rsPati!����ȼ�, "����") & "|" & rsPati!�ѱ� & "|" & NVL(rsPati!��ǰ����, "һ��") & "|" & Format(rsPati!��Ժ����, "yyyy-MM-dd") & "|" & Format(rsPati!��Ժ����, "yyyy-MM-dd") & "|" & rsPati!סԺ���� & "|" & rsPati!��Ժ��ʽ & "|" & _
                      NVL(rsPati!��������, "��ͨ����") & "|" & rsPati!״̬ & "|" & NVL(rsPati!����, 0) & "|" & NVL(rsPati!���￨��) & "|" & NVL(rsPati!·��״̬, 0) & "|" & NVL(rsPati!��ɫ, 0) & "|" & NVL(rsPati!������) & "|" & NVL(rsPati!Ӥ������ID, 0) & "|" & NVL(rsPati!Ӥ������ID, 0)
            
            Call Rec.AddNew(mrsPati, strField, strValue)
        End If
        rsPati.MoveNext
    Loop
End Sub

Private Sub Form_Load()
    Dim objPane As Pane, strTab As String, intIdx As Integer
    Dim blnCol As Boolean, strTmp As String, i As Long, bln·��״̬ As Boolean
    Dim intType As Integer, arrTmp As Variant
    
    '���廤����ҵ��ǩ�����
    Dim strTabs As String, strErrMsg As String
    Dim strName As String, strUrl As String, strParam As String
    Dim j As Integer
    Dim objXML As New DOMDocument
    Dim objNodeList As IXMLDOMNodeList
    
    On Error GoTo ErrHand
    
    picPrompt.Visible = False
    
    mblnRefreshBar = False
    marrTabAttribute = Array()
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons

    'TabControl
    '-----------------------------------------------------
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, True)
    If GetInsidePrivs(p�°�סԺ����, True) <> "" Then
        Set mclsEMR = DynamicCreate("zlRichEMR.clsDockEMR", "���Ӳ���")
        If Not mclsEMR Is Nothing Then
            If Not mclsEMR.Init(gobjEmr, gcnOracle, glngSys) Then
                Set mclsEMR = Nothing
            End If
        End If
    End If
    Set mclsAdvices = New zlPublicAdvice.clsDockInAdvices
    Call mclsAdvices.zlInitMip(mobjMipModule)
    Set mclsEPRs = New zlRichEPR.cDockInEPRs
    Set mclsTends = New zl9TendFile.clsTendFile
    Call mclsTends.InitTendFile(gcnOracle, glngSys)
    Set mclsTendEPRs = New zlRichEPR.cDockInTendEPRs
    Set mclsFeeQuery = New zl9InExse.clsFeeQuery
    Call mclsFeeQuery.InitCallByNurse(gfrmMain, gcnOracle, gstrDBUser, glngSys)

    Set mclsPath = New zlPublicPath.clsDockPath
    Call mclsAdvices.zlInitPath(mclsPath)
    Set mclsWardMonitor = New clsWardMonitor
    
    Set mcolSubFormOperation = New Collection
    Set mcolSubForm = New Collection
    If Not mclsEMR Is Nothing Then
        mcolSubForm.Add mclsEMR.zlGetForm, "_�²���"
    End If
    mcolSubForm.Add mclsPath.zlGetForm, "_·��"
    mcolSubForm.Add mclsAdvices.zlGetForm, "_ҽ��"
    mcolSubForm.Add mclsFeeQuery.zlGetForm, "_����"
    mcolSubForm.Add mclsEPRs.zlGetForm, "_����"
    mcolSubForm.Add mclsTends.zlGetForm, "_����"
    mcolSubForm.Add mclsTendEPRs.zlGetForm, "_������"
    If mclsWardMonitor.Enabled Then
        mcolSubForm.Add mclsWardMonitor.zlGetForm, "_�໤"
    End If
    
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .OneNoteColors = True
            .Position = xtpTabPositionTop
            .ShowIcons = True
        End With
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        If GetInsidePrivs(p�ٴ�·��Ӧ��, True) <> "" Then
            .InsertItem(intIdx, "�ٴ�·��", picTmp.hwnd, 0).Tag = "·��": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(pסԺҽ���´�, True) <> "" Or GetInsidePrivs(pסԺҽ������, True) <> "" Then
            .InsertItem(intIdx, "ҽ����¼", picTmp.hwnd, 0).Tag = "ҽ��": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p���ò�ѯ, True) <> "" Then
            .InsertItem(intIdx, "���ü�¼", picTmp.hwnd, 0).Tag = "����": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(pסԺ��������, True) <> "" Then
            .InsertItem(intIdx, "סԺ����", picTmp.hwnd, 0).Tag = "����": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p�����¼����, True) <> "" Then
            .InsertItem(intIdx, "�����¼", picTmp.hwnd, 0).Tag = "����": intIdx = intIdx + 1
            .InsertItem(intIdx, "������", picTmp.hwnd, 0).Tag = "������": intIdx = intIdx + 1
        End If
        If mclsWardMonitor.Enabled Then
            If InStr(GetInsidePrivs(pסԺ��ʿվ), "����໤") > 0 Then
                .InsertItem(intIdx, "����໤", picTmp.hwnd, 0).Tag = "�໤": intIdx = intIdx + 1
            End If
        End If
        If GetInsidePrivs(p�°�סԺ����, True) <> "" And Not mclsEMR Is Nothing Then
            .InsertItem(intIdx, "���Ӳ���", picTmp.hwnd, 0).Tag = "�²���": intIdx = intIdx + 1
        End If
        
        For i = 0 To tbcSub.ItemCount - 1
            ReDim Preserve marrTabAttribute(UBound(marrTabAttribute) + 1)
            marrTabAttribute(UBound(marrTabAttribute)) = 0
        Next
        '���廤����ҵ��Ƕ��
        If gbln�������廤��ӿ� = True Then
            If InitNurseIntegrate = True Then
                If gobjNurseIntegrate.GetPatientMethod(strTabs, strErrMsg) = False Then
                    MsgBox "��ȡ���廤����ҵ���ǩʧ�ܣ�" & vbCrLf & "��ϸ��Ϣ��" & strErrMsg, vbInformation, gstrSysName
                Else
                    If objXML.loadXML(strTabs) = False Then Exit Sub
                    Set mColNurseFormUrl = New Collection
                    Set objNodeList = objXML.selectNodes(".//Tab//Item")
                    For i = 0 To objNodeList.length - 1
                        strName = objNodeList.Item(i).childNodes(0).Text
                        strUrl = objNodeList.Item(i).childNodes(1).Text
                        '��ȡ�ڵ�����ֵ
                        strParam = ""
                        For j = 0 To objNodeList.Item(i).childNodes(1).Attributes.length - 1
                             strParam = strParam & "&" & objNodeList.Item(i).childNodes(1).Attributes(j).nodeName & "=" & objNodeList.Item(i).childNodes(1).Attributes(j).nodeValue
                        Next j
                        If Left(strParam, 1) = "&" Then strParam = Mid(strParam, 2)
                        strUrl = strUrl & IIf(strParam = "", "", "?" & strParam)
                        .InsertItem(intIdx, strName, picTmp.hwnd, 0).Tag = strName: intIdx = intIdx + 1
                        mColNurseFormUrl.Add strUrl, "_" & strName
                        ReDim Preserve marrTabAttribute(UBound(marrTabAttribute) + 1)
                        marrTabAttribute(UBound(marrTabAttribute)) = 1
                    Next i
                End If
            End If
        End If
        
        Call CreatePlugInOK(pסԺ��ʿվ)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strTmp = gobjPlugIn.GetFormCaption(glngSys, pסԺ��ʿվ)
            Call zlPlugInErrH(err, "GetFormCaption")
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ",")
                For i = 0 To UBound(arrTmp)
                    strTmp = arrTmp(i)
                    
                    mcolSubForm.Add gobjPlugIn.GetForm(glngSys, pסԺ��ʿվ, strTmp), "_" & strTmp
                    .InsertItem(intIdx, strTmp, mcolSubForm("_" & strTmp).hwnd, 0).Tag = strTmp: intIdx = intIdx + 1
                    ReDim Preserve marrTabAttribute(UBound(marrTabAttribute) + 1)
                    marrTabAttribute(UBound(marrTabAttribute)) = 0
                    Call zlPlugInErrH(err, "GetForm")
                Next
            End If
            err.Clear: On Error GoTo 0
        End If
        
        If .ItemCount = 0 Then
            MsgBox "��û��ʹ�ò����������Ȩ�ޡ�", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        '�ָ��ϴ�ѡ��Ŀ�Ƭ
        strTab = zlDatabase.GetPara("ҽ������", glngSys, pסԺ��ʿվ)
        strTab = mstrPage
        For intIdx = 0 To tbcSub.ItemCount - 1
            If tbcSub(intIdx).Visible And tbcSub(intIdx).Tag = strTab Then Exit For
        Next
        If intIdx <= tbcSub.ItemCount - 1 Then
            strTab = .Item(intIdx).Tag
            .Item(intIdx).Tag = "" '���⼤���¼�
            .Item(intIdx).Selected = True
            .Item(intIdx).Tag = strTab
        Else
            .Item(0).Selected = True '�½�ʱ���Զ�ѡ�������,�����ټ����¼�
        End If
        'ֻ����ѡ����Ӵ���
        Call tbcSub_SelectedChanged(.Selected)
    End With
    
    '��ʼ������ѡ����
    Dim objCol As ReportColumn
    With rptPati
        .Columns.DeleteAll
        Set objCol = .Columns.Add(c_ͼ��, "", cw_ͼ��, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_״̬, "״̬", cw_״̬, True)
        Set objCol = .Columns.Add(c_����, "����", cw_����, True)
        Set objCol = .Columns.Add(C_����ID, "����ID", Cw_����ID, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_��ҳID, "��ҳID", cw_��ҳID, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_����, "����", cw_����, True)
        Set objCol = .Columns.Add(c_סԺ��, "סԺ��", cw_סԺ��, True)
        Set objCol = .Columns.Add(c_��Ժ����, "��Ժ����", cw_��Ժ����, True)
        Set objCol = .Columns.Add(c_��Ժ����, "��Ժ����", cw_��Ժ����, True)
        Set objCol = .Columns.Add(c_��������, "��������", cw_��������, True)
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = (objCol.Index = C_״̬)
            objCol.Sortable = True
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�в���..."
        End With
        .PreviewMode = False
        .AllowColumnSort = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgRPT
        
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(C_״̬)
        .GroupsOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(c_����)
    End With

    '��ȡ��������
    '-----------------------------------------------------
    mstrMonitor = ""
    mblnMonitor = Dir(App.Path & "\..\gdhs\AC2005.exe") <> ""
    If mblnMonitor Then mstrMonitor = App.Path & "\..\gdhs\AC2005.exe"
    
    '����ָ�:�������ִ��
    '-----------------------------------------------------
    Call RestoreWinState(Me, App.ProductName)
    
    '�ָ��ϴβ�����Ϣ��������״̬
    If (zlDatabase.GetPara("������Ϣ������", glngSys, pסԺ��ʿվ, 1) = 0) Then
        mobjBar.Visible = False
'        picInfo.Visible = False
    End If
    Me.WindowState = vbMaximized
    
    Call AddPages
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long, strTmp As String
    Dim lng����ID As Long, lng��ҳID As Long
    
    On Error GoTo ErrHand

    If Control.ID <> 0 Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If

    Select Case Control.ID
    Case conMenu_Manage_Monitor '�໤��
        Call ExecuteMonitor
    Case conMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.picPrompt.Visible = Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_Jump '��ת
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
    Case conMenu_View_Refresh 'ˢ��
        '68116:������,2014-01-06,ˢ����Ӳ����б�ˢ��
        'Call tbcSub_SelectedChanged(tbcSub.Item(tbcSub.Selected.Index))
        lng����ID = mlng����ID: lng��ҳID = mlng��ҳID
        Call RefreshPatiList(mrsPati)
        If lng����ID = mlng����ID Then mlng��ҳID = lng��ҳID
        mintPrePage = -1
        Call AddPages
        Call ReSetFontSize
    Case conMenu_Tool_Reference_1 '������ϲο�
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
    Case conMenu_Tool_Reference_2 '���ƴ�ʩ�ο�
        Call gobjKernel.ShowClincHelp(vbModeless, Me)
    
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Help_Help '����
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '�˳�
        Unload Me
    Case Else
        mblnUnRefresh = True
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            'ִ�з�������ǰģ��ı���
            With mPatiInfo
                strTmp = Split(Control.Parameter, ",")(1)
                If strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1132" Then 'סԺ�����ձ�
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strTmp, Me, _
                             "����=" & mlng����ID, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID)
                ElseIf strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1139_2" Or strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1139_1" Then    '������ҳ�ʹ߿��
                    Call mclsFeeQuery.zlExecuteCommandBars(Control)
                Else
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strTmp, Me, _
                        "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID, "סԺ��=" & .סԺ��, "���˲���=" & .����ID, _
                        "���˿���=" & .����ID, "����=" & .����)
                End If
            End With
        Else
            Select Case Me.tbcSub.Selected.Tag
            Case "·��"
                If mlngӤ������ID <> 0 Then
                    If mlngӤ������ID = mlng����ID Then
                        MsgBox "�ò����Ѿ�ת���������ˣ�ֻ��Ӥ�����ڱ����ң����������·����", vbInformation, Me.Caption
                        Exit Sub
                    End If
                End If
                Call mclsPath.zlExecuteCommandBars(Control)
            Case "ҽ��"
                Call mclsAdvices.zlExecuteCommandBars(Control)
            Case "����"
                Call mclsFeeQuery.zlExecuteCommandBars(Control)
            Case "����"
                Call mclsEPRs.zlExecuteCommandBars(Control)
            Case "����"
                Call mclsTends.zlExecuteCommandBars(Control)
            Case "������"
                Call mclsTendEPRs.zlExecuteCommandBars(Control)
            Case "�²���"
                Call mclsEMR.zlExecuteCommandBars(Control)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.ExeButtomClick(glngSys, pסԺ��ʿվ, mcolSubForm("_" & tbcSub.Selected.Tag), tbcSub.Selected.Tag, Control.Caption, mlng����ID, mlng��ҳID, "")
                    Call zlPlugInErrH(err, "ExeButtomClick")
                    err.Clear: On Error GoTo 0
                End If
            End Select
        End If
        mblnUnRefresh = False
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim rsPatiLog As ADODB.Recordset
    Dim i As Long, j As Long, strPrivs As String
    Dim objControl As CommandBarControl
    On Error GoTo ErrHand

    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case tbcSub.Selected.Tag
    Case "·��"
         Call mclsPath.zlPopupCommandBars(CommandBar)
    Case "ҽ��"
        Call mclsAdvices.zlPopupCommandBars(CommandBar)
    Case "����"
        Call mclsFeeQuery.zlPopupCommandBars(CommandBar)
    Case "����"

    Case "����"

    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo ErrHand
    
    Select Case Control.ID
    Case conMenu_Manage_Monitor '�໤��
        Control.Visible = mblnMonitor
    Case conMenu_View_ToolBar_Button '������
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_Tool_Reference_1 '������ϲο�
        Control.Visible = GetInsidePrivs(p������ϲο�) <> ""
    Case conMenu_Tool_Reference_2 'ҩƷ�����Ʋο�
        Control.Visible = GetInsidePrivs(pҩƷ���Ʋο�) <> ""
    Case conMenu_View_ToolBar_Text 'ͼ������
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
'    Case conMenu_Tool_MedRecAuditResponse '��鷴��
'        '�����Ե��ã����ٿ��Բ鿴(��ǰ����ʷ)
'        Control.Enabled = rptPati.Rows.Count > 0
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1139_1" Then Control.Visible = tbcSub.Selected.Tag = "����"  '�߿��
            If Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1139_2" Then Control.Visible = tbcSub.Selected.Tag = "����"  '������ҳ
        End If
        If Not mblnRefreshBar Then Exit Sub
        Select Case tbcSub.Selected.Tag
        Case "·��"
            Call mclsPath.zlUpdateCommandBars(Control)
        Case "ҽ��"
            Call mclsAdvices.zlUpdateCommandBars(Control)
        Case "����"
            Call mclsFeeQuery.zlUpdateCommandBars(Control)
        Case "����"
            Call mclsEPRs.zlUpdateCommandBars(Control)
        Case "����"
            Call mclsTends.zlUpdateCommandBars(Control)
        Case "������"
            Call mclsTendEPRs.zlUpdateCommandBars(Control)
        Case "�²���"
            Call mclsEMR.zlUpdateCommandBars(Control)
        End Select
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Private Sub SubWinDefCommandBar(ByVal objItem As TabControlItem)
'���ܣ�ˢ���Ӵ���˵���������
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long, idx As Long
    Dim strName As String
    Dim blnNurseIntegrate As Boolean
    
    On Error GoTo ErrHand
    If gbln�������廤��ӿ� = True Then
        blnNurseIntegrate = Val(marrTabAttribute(objItem.Index)) = 1
    End If
    '��¼���в˵���ʽ
    mblnRefreshBar = False
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsMain.Count >= 2 Then
        idx = GetFirstCommandBar(cbsMain(2).Controls)
        If idx > 0 Then
            blnShowBar = cbsMain(2).Visible
            bytStyle = cbsMain(2).Controls(idx).Style
        End If
    End If

    'ˢ���Ӵ��ڲ˵�
    Call LockWindowUpdate(Me.hwnd)

    Me.Caption = "���������� - " & objItem.Caption & "(��ǰ�û���" & UserInfo.���� & ")"

    'ɾ�����ڵ����в˵���
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    'ɾ��������
'    For lngCount = cbsMain.Count To 2 Step -1
'        cbsMain(lngCount).Delete
'    Next

    '���������¼���
    Call MainDefCommandBar
    
    '�Ӵ������¼���
    Select Case objItem.Tag
    Case "·��"
        Call mclsPath.zlDefCommandBars(Me, Me.cbsMain, 1, True)
    Case "ҽ��"
        Call mclsAdvices.zlDefCommandBars(Me, Me.cbsMain, 1, True)
    Case "����"
        Call mclsFeeQuery.zlDefCommandBars(Me, Me.cbsMain, 1, True)
    Case "����"
        Call mclsEPRs.zlDefCommandBars(Me.cbsMain, True)
    Case "����"
        Call mclsTends.zlDefCommandBars(Me.cbsMain, True)
    Case "������"
        Call mclsTendEPRs.zlDefCommandBars(Me.cbsMain, True)
    Case "�²���"
        Call mclsEMR.zlDefCommandBars(Me.cbsMain, True)
    Case Else
        If blnNurseIntegrate = False Then
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                strName = gobjPlugIn.GetButtomName(glngSys, pסԺ��ʿվ, mcolSubForm("_" & objItem.Tag), objItem.Tag)
                Call zlPlugInErrH(err, "GetButtomName")
                '�����˵�
                If strName <> "" Then Call PlugInInSideBar(cbsMain, strName, 1)
                err.Clear: On Error GoTo 0
            End If
        End If
    End Select
    mblnRefreshBar = True
    '�ָ����̶���һЩ�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '�ָ���������ť����
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
'        cbsMain(lngCount).EnableDocking xtpFlagHideWrap
'        For Each objControl In cbsMain(lngCount).Controls
'            If objControl.ID = conMenu_Help_Help Or objControl.ID = conMenu_File_Exit Or objControl.ID = conMenu_File_Print Or objControl.ID = conMenu_File_Preview Then
'                objControl.Style = xtpButtonIcon
'            Else
'                objControl.Style = bytStyle
'            End If
'        Next
'        cbsMain(lngCount).Visible = blnShowBar
    Next

    '�������RecalcLayout����������
    Call LockWindowUpdate(0)

    If blnNurseIntegrate = False Then
        Set mfrmActive = mcolSubForm("_" & objItem.Tag)
    Else
        Set mfrmActive = mobjNurseForm
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
 '���ܣ�ˢ���Ӵ������ݼ�״̬
    Dim blnEdit As Boolean, strInPatiNO As String, lng·��״̬ As Long
    Dim lngType As PATI_TYPE, lng����ID As Long, lng����ID As Long
    Dim lngState As TYPE_PATI_State
    Dim blnNurseIntegrate As Boolean
    
    On Error GoTo ErrHand
    If gbln�������廤��ӿ� = True Then
        blnNurseIntegrate = Val(marrTabAttribute(objItem.Index)) = 1
    End If
    
    Call SetOrGetSubFromOperation(objItem.Tag, True)
    If mlng����ID = 0 Then
        'Ҫ���Ӵ��尴�����ݴ������
        Select Case objItem.Tag
        Case "·��"
            Call mclsPath.zlRefresh(0, 0, 0, 0, 0, False)
        Case "ҽ��"
            Call mclsAdvices.zlRefresh(0, 0, 0, 0, CDate(0), 0)
        Case "����"
            Call mclsFeeQuery.zlRefresh(0, 0, 0, 0, 0, False, False, False)
        Case "����"
            Call mclsEPRs.zlRefresh(0, 0, 0, False, False)
        Case "����"
            Call mclsTends.zlRefresh(0, 0, 0, False, False)
        Case "������"
            Call mclsTendEPRs.zlRefresh(0, 0, 0, False, False, False)
        Case "�໤"
            Call mclsWardMonitor.HideWindow
        Case "�²���"
            Call mclsEMR.zlRefresh(0, 0, 0, 0, 3)
        Case Else
            If blnNurseIntegrate = False Then
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.RefreshForm(glngSys, pסԺ��ʿվ, mcolSubForm("_" & objItem.Tag), objItem.Tag, 0, "", 0, False)
                    Call zlPlugInErrH(err, "RefreshForm")
                    err.Clear: On Error GoTo 0
                End If
            Else
                If InitNurseIntegrate = True Then
                    Call gobjNurseIntegrate.RefreshPatientMethod(mobjNurseForm, mobjNurseForm.Tag, mstrNurseParentID, mstrRelatedUnitID, mstrRelatedUserID)
                End If
            End If
        End Select
    Else
        With mPatiInfo
            lngType = .����
            
            '67485:������,2013-11-13,�鿴��ת������Ӧ����ת��֮ǰ�Ŀ���ID
            If pt���ת�� = lngType And mrsPati.RecordCount > 0 Then
                lng����ID = NVL(mrsPati!����ID, 0) '���ת������Ϊԭ����ID
            Else
                lng����ID = Val("" & .����ID)
            End If
            
            If InStr("," & pt��Ժ����ס & "," & pt���ת�� & "," & ptת�ƴ���ס & "," & ptת��������ס & ",", "," & lngType & ",") > 0 Then
                '����ס���ˣ�ת�����ˣ�����ǰ����Ĳ���
                lng����ID = mlng����ID
            Else
                lng����ID = .����ID
            End If
            If lngType = pt���ת�� Then
                lngState = ps���ת��
            ElseIf lngType = ptת�ƴ���ס Or lngType = ptת��������ס Then
                lngState = ps��ת��
            Else
                lngState = IIf(.��Ժ���� = CDate(0), IIf(.״̬ = 3, psԤ��, ps��Ժ), ps��Ժ)
            End If
            
            Select Case objItem.Tag
            Case "·��"
                Call mclsPath.zlRefresh(mlng����ID, .��ҳID, lng����ID, lng����ID, .״̬, .����ת��, True, , mlng����ID)
            Case "ҽ��"
                lng·��״̬ = .·��״̬
                '50906:������,2012-09-18,��Ժ����ס���ˣ����ݲ���"���������ס�����´�ҽ��"�����Ƿ�����´�ҽ��
                If lngType = pt��Ժ����ס And Val(zlDatabase.GetPara("���������ס�����´�ҽ��", glngSys, pסԺҽ���´�, 1)) = 0 Then
                    lngState = ps��ת�� 'lngState=ps��ת��ʱ�¿�ҽ���ȹ��ܲ�����
                End If
                Call mclsAdvices.zlRefresh(mlng����ID, .��ҳID, lng����ID, lng����ID, lngState, .����ת��, , , , lng·��״̬, mlng����ID)
            Case "����"
                Call mclsFeeQuery.zlRefresh(mlng����ID, mlng��ҳID, Val(.סԺ��), lng����ID, .����, .����ת��, .��Ժ���� <> CDate("0:00:00"), .����, False, _
                    lngType = pt���ת�� Or lngType = ptԤ�� Or lngType = pt��Ժ, lng����ID)
            Case "����"
                Call mclsEPRs.zlRefresh(mlng����ID, .��ҳID, mlng����ID, False, .����ת��, 0, True, lng����ID, lngState)
            Case "����"
                blnEdit = True
                If lngType = pt��Ժ Or lngType = pt���� Then
                    If Not (Val(.����״̬) = 0 Or Val(.����״̬) = 2 Or Val(.����״̬) = 999) Then
                        '��������Ժ��鷴��״̬����Ժ��δ�ύ���
                        If Val(.����״̬) = 1 Or Val(.����״̬) = 2 Then blnEdit = False
                    End If
                ElseIf lngType = ptת�ƴ���ס Or lngType = ptת��������ס Then
                    blnEdit = False
                End If
                blnEdit = blnEdit And (mlng����ID = .����ID Or lngType = pt���ת��)
                Call mclsTends.zlRefresh(mlng����ID, .��ҳID, mlng����ID, blnEdit, False, lng����ID, lngState)
            Case "������"
                Call mclsTendEPRs.zlRefresh(mlng����ID, .��ҳID, mlng����ID, True, True, .����ת��)
            Case "�໤"
                strInPatiNO = Trim(.סԺ��)
                If strInPatiNO = "" Then
                    Call mclsWardMonitor.HideWindow
                Else
                    Call mclsWardMonitor.ShowInfor(strInPatiNO)
                End If
            Case "�²���"
                Call mclsEMR.zlRefresh(mlng����ID, .��ҳID, mlng����ID, lngState, 3)
            Case Else
                If blnNurseIntegrate = False Then
                    If Not gobjPlugIn Is Nothing Then
                        On Error Resume Next
                        Call gobjPlugIn.RefreshForm(glngSys, pסԺ��ʿվ, mcolSubForm("_" & objItem.Tag), objItem.Tag, mlng����ID, "", .��ҳID, .����ת��, , , _
                                        lng����ID, lng����ID, , lngState, , lng·��״̬)
                        Call zlPlugInErrH(err, "RefreshForm")
                        err.Clear: On Error GoTo 0
                    End If
                Else
                    If InitNurseIntegrate = True Then
                        Call gobjNurseIntegrate.RefreshPatientMethod(mobjNurseForm, mobjNurseForm.Tag, mstrNurseParentID, mstrRelatedUnitID, mstrRelatedUserID)
                    End If
                End If
            End Select
        End With
    End If
    
    '��������
    Select Case objItem.Tag
        Case "·��"
            Call mclsPath.SetFontSize(mbytSize)
        Case "ҽ��"
            Call mclsAdvices.SetFontSize(mbytSize)
        Case "����"
            Call mclsFeeQuery.SetFontSize(mbytSize)
        Case "����"
            Call mclsEPRs.SetFontSize(mbytSize)
        Case "����"
            Call mclsTends.SetFontSize(mbytSize)
        Case "������"
            Call mclsTendEPRs.SetFontSize(mbytSize)
        Case "�໤"
            'Call mclsWardMonitor.SetFontSize(mbytSize)
        Case "�²���"
            On Error Resume Next
            Call mclsEMR.SetFontSize(mbytSize)
            err.Clear: On Error GoTo 0
        End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl

    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False) '����
    objMenu.ID = conMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��") '����
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True '����
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False) '����
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)") '����
'        With objPopup.CommandBar.Controls
'            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False '����
'            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False '����
'            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False '����
'        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)") '����

        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Jump, "������ת(&J)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", -1, False)
    objMenu.ID = conMenu_ToolPopup
    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������(&I)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "���ϲο�(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "������ϲο�(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "���ƴ�ʩ�ο�(&C)", -1, False
        End With
'        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditResponse, "��鷴��(&S)")
'            objControl.BeginGroup = True
'            objControl.ToolTipText = "�����鿴������鷴��"
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False) '����
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)") '����
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName) '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False '����
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False '����
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True '����
    End With
    cbsMain(1).EnableDocking xtpFlagHideWrap
    
    If mblnAdd Then
        '����������
        '-----------------------------------------------------
        Set objBar = cbsMain.Add("����������", xtpBarTop) '����
        objBar.EnableDocking xtpFlagStretched
        With objBar.Controls
            Set objCustom = .Add(xtpControlCustom, 1, "")
            objCustom.Handle = picCondition.hwnd
        End With
        Set mobjBar = cbsMain.Add("������Ϣ������", xtpBarTop) '����
        mobjBar.EnableDocking xtpFlagStretched
        mobjBar.Closeable = True
        With mobjBar.Controls
            Set objCustom = .Add(xtpControlCustom, 1, "")
            objCustom.Handle = picInfo.hwnd
        End With
        mblnAdd = False
    End If

    '��ȡ��������ģ��ı���(��������ģ���,��:סԺ�����ձ����߿���߿������ʾ,�����ֹ��ӵ��ļ��˵���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, 1265, mstrPrivs, "ZL1_INSIDE_1261_1", "ZL1_INSIDE_1261_5", "ZL1_INSIDE_1261_4", "ZL1_INSIDE_1261_6", "ZL1_INSIDE_1132", "ZL1_INSIDE_1139_1", "ZL1_INSIDE_1139_3", "ZL1_INSIDE_1261_7", "ZL1_INSIDE_1261_8")

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF6, conMenu_View_Jump '��ת
    End With
    
    Call cbsMain.RecalcLayout
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    picInfo.Width = Me.ScaleWidth - 200
    
    Call cbsMain.RecalcLayout
End Sub

Private Sub img�����б�_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngColor As Long, j As Long
    Dim lngloop As Long
    Dim objRow As ReportRow, blnSelect As Boolean
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngLeft As Long, lngTop  As Long, lngRight As Long, lngBottom As Long
    If Button <> 1 Then Exit Sub
    On Error GoTo ErrHand
    
    If rptPati.Records.Count = 0 Then
        '��ʾ�����б�ѡ��
        With mrsPati
            .MoveFirst
            
            Do While Not .EOF
                Set objRecord = Me.rptPati.Records.Add()
                objRecord.Tag = CStr(!����ID & "," & !��ҳID)
                
                Set objItem = objRecord.AddItem("")
                
                '61824:������,2013-05-23,��ʾ�����ֱ�־
                If NVL(!������) <> "" Then
                    objItem.Icon = imgRPT.ListImages("������").Index - 1
                Else
                    objItem.Icon = Val(IIf(!�Ա� = "Ů", imgRPT.ListImages("Ů��").Index, imgRPT.ListImages("����").Index)) - 1
                End If
                Set objItem = objRecord.AddItem(CStr(!���� & !����))
                objItem.Caption = CStr(!���� & !����)
                
                Set objItem = objRecord.AddItem(zlStr.Lpad(NVL(!����), 10))
                objItem.Caption = Trim(NVL(!����, " "))
                objRecord.AddItem Val(!����ID)
                objRecord.AddItem Val(!��ҳID)
                objRecord.AddItem CStr(NVL(!����))
                Set objItem = objRecord.AddItem(CStr(NVL(!סԺ��)))
                objItem.Caption = NVL(!סԺ��, " ")
                
                Set objItem = objRecord.AddItem(Format(!��Ժ����, "yyyy-MM-dd"))
                objItem.Caption = Format(!��Ժ����, "yyyy-MM-dd")
                Set objItem = objRecord.AddItem(Format(!��Ժ����, "yyyy-MM-dd"))
                objItem.Caption = Format(!��Ժ����, "yyyy-MM-dd")
                
                Set objItem = objRecord.AddItem(NVL(!��������))
                objItem.Caption = NVL(!��������)
                
                '��ȡ�������͵���ɫ
                lngColor = NVL(!��ɫ, 0)
                If lngColor <> 0 Then
                    For j = 1 To rptPati.Columns.Count - 1
                        objRecord.Item(j).ForeColor = lngColor
                    Next
                End If
                .MoveNext
            Loop
            
            .MoveFirst
            .Find ("Key='" & mlng����ID & ":" & mlng��ҳID & "'")
            If .EOF Then .MoveFirst: .Find "����ID=" & mlng����ID
        End With
    End If
    '��������
    Call mobjBar.GetWindowRect(lngLeft, lngTop, lngRight, lngBottom)
    rptPati.Populate 'ȱʡ��ѡ���κ���
    picPati.Left = picCondition.Left + Me.pic����.Left
    picPati.Top = lngTop - Me.Top - 480
    picPati.Visible = True
    mlngRowIndex = -1
    'ѡ�е�ǰ����(���۵���Ļ�,Rows.Countֻ����ĸ�����,�����ȶ�λ,���۵�)
    blnSelect = False
    For lngloop = 0 To rptPati.Rows.Count - 1
        If Not (rptPati.Rows(lngloop).Record Is Nothing) Then
            If Val(rptPati.Rows(lngloop).Record.Item(C_����ID).Value) = mlng����ID Then
                Set objRow = rptPati.Rows(lngloop)
            End If
            If Val(rptPati.Rows(lngloop).Record.Item(C_����ID).Value) = mlng����ID And Val(rptPati.Rows(lngloop).Record.Item(C_��ҳID).Value) = mlng��ҳID Then
                Set rptPati.FocusedRow = rptPati.Rows(lngloop)
                blnSelect = True
                Exit For
            End If
        End If
    Next
  
    If blnSelect = False And Not objRow Is Nothing Then
        Set rptPati.FocusedRow = objRow
    End If
    
    '�۵�������(ѡ�в�����һ�鲻�۵�)
    For Each objRow In rptPati.Rows
        If objRow.GroupRow And objRow.Index <> rptPati.FocusedRow.ParentRow.Index Then
            objRow.Expanded = False
        End If
    Next
    rptPati.FocusedRow.EnsureVisible
    If rptPati.Visible Then rptPati.SetFocus
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub img�����б�_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo Me.pic����.hwnd, img�����б�.Tag
End Sub

Private Sub img��һ��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LocatePati(1)
End Sub

Private Sub img��һ��_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCondition.hwnd, img��һ��.Tag
End Sub

Private Sub img��һ��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LocatePati(2)
End Sub

Private Sub img��һ��_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCondition.hwnd, img��һ��.Tag
End Sub

Private Sub img��ϸ��Ϣ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
'    picInfo.Visible = picInfo.Visible Xor True
    mobjBar.Visible = mobjBar.Visible Xor True
End Sub

Private Sub img��ϸ��Ϣ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCondition.hwnd, img��ϸ��Ϣ.Tag
End Sub


Private Sub lbl����_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo pic��ʶ.hwnd, lbl����.Caption
End Sub

Private Sub mclsAdvices_ExecLogModi(ByVal ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal ����ID As Long, ByVal ִ��ʱ�� As String, ��� As Boolean)
    On Error Resume Next
    mblnUnRefresh = True
    ��� = frmTechnicLog.ShowMe(Me, pסԺҽ������, ����ID, ҽ��ID, ���ͺ�, False, ִ��ʱ��)
    mblnUnRefresh = False
    On Error GoTo 0
End Sub

Private Sub mclsAdvices_ExecLogNew(ByVal ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal ����ID As Long, ��� As Boolean)
    On Error Resume Next
    mblnUnRefresh = True
    ��� = frmTechnicLog.ShowMe(Me, pסԺҽ������, ����ID, ҽ��ID, ���ͺ�, False)
    mblnUnRefresh = False
    On Error GoTo 0
End Sub

Private Sub mclsAdvices_RequestRefresh(ByVal RefreshNotify As Boolean)
'���ܣ�ҽ���Ӵ���Ҫ��ˢ��
    If RefreshNotify = True Then
        '��ˢ��ҽ����������(�Զ�ˢ��ʱ)
        frmNotify.mblnFirst = True
    Else
        '55982:������,2012-11-20,�޸ķ��ͳ�Ժҽ��������ҽ����ˢ������
        '����ˢ��ҽ����Ϣ
        Call tbcSub_SelectedChanged(tbcSub.Item(tbcSub.Selected.Index))
    End If
End Sub

Private Sub mclspath_RequestRefresh(ByVal lngPathState As Long)
'���ܣ��ٴ�·����ˢ�²�����Ϣ�б��е�״̬,-1��ʾδ����״̬
    'todo:��Ҫ����
'    With rptPati.SelectedRows(0)
'        .Record(col_·��״̬).Value = lngPathState
'        .Record(col_·��״̬).Caption = " "
'        .Record(col_·��״̬).Icon = -1 + Choose(lngPathState + 2, imgPati.ListImages("δ����").Index, imgPati.ListImages("������").Index, _
'                imgPati.ListImages("ִ����").Index, imgPati.ListImages("��������").Index, imgPati.ListImages("�������").Index)
'    End With
'
'    If rptPati.Columns(col_·��״̬).Visible = False Then
'        rptPati.Columns(col_·��״̬).Visible = True
'    End If
'    rptPati.Populate
End Sub

Private Sub mclsAdvices_StatusTextUpdate(ByVal Text As String)
'���ܣ�ҽ���Ӵ���Ҫ�����״̬��
    'todo:��Ҫ����
    If Text = "" Then
        If mlng����ID > 0 And mlng��ҳID > 0 Then
            lblPrompt.Caption = IIf(stbThis.Panels(2).Tag = "", "", stbThis.Panels(2).Tag & "��") & _
                mrsPati!���� & "��" & GetPati������Ϣ(mlng����ID, mlng��ҳID)
        Else
            lblPrompt.Caption = stbThis.Panels(2).Tag
        End If
    Else
        lblPrompt.Caption = Text
    End If
    lblPrompt.ForeColor = &H80000008
End Sub

Private Sub cboPages_Click()
'���ܣ�ѡ��ĳ��סԺ��¼ʱ����ȡ��صĲ�����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lng��ҳID As Long
    
    If cboPages.ListIndex = -1 Then Exit Sub
    If cboPages.ListIndex = mintPrePage Then Exit Sub
    mintPrePage = cboPages.ListIndex
    mlng��ҳID = cboPages.ItemData(cboPages.ListIndex)

    On Error GoTo errH
    '90592:����б�����ͬ�����ж�������ѡ��סԺ������Ĭ�϶�λ
    lng��ҳID = Val(mrsPati!��ҳID)
    If Not Val(mrsPati!��ҳID) = mlng��ҳID Then
        mrsPati.MoveFirst: mrsPati.Find "Key='" & mlng����ID & ":" & mlng��ҳID & "'"
        If mrsPati.EOF = True Then mrsPati.MoveFirst: mrsPati.Find "Key='" & mlng����ID & ":" & lng��ҳID & "'"
    End If
    strSQL = "Select NVL(b.����,a.����) ����, NVL(b.�Ա�,a.�Ա�) �Ա�, NVL(b.����,a.����) ����, b.סԺ��, b.��Ժ����, b.ҽ�Ƹ��ʽ, d.��Ϣֵ As ҽ����, b.����, b.��ǰ����, c.���� As ����ȼ�, Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) AS ��Ժ���� , b.��Ժ����, b.��Ŀ����," & vbNewLine & _
            "       b.��������, b.״̬, b.����ת��, b.��Ժ����id, b.��ǰ����id,b.����״̬,B.Ӥ������ID,B.Ӥ������ID, a.סԺ����, e.�����" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, �շ���ĿĿ¼ C, ������ҳ�ӱ� D, ��λ״����¼ E" & vbNewLine & _
            "Where a.����id = b.����id And a.����id = [1] And b.��ҳid = [2] And b.����ȼ�id = c.Id(+) And b.����id = d.����id(+) And" & vbNewLine & _
            "      b.��ҳid = d.��ҳid(+) And d.��Ϣ��(+) = 'ҽ����' And b.��Ժ����id = e.����id(+) And b.����id = e.����id(+) And b.��Ժ���� = e.����(+)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    With rsTmp
        '���ղ���������ɫ��ʾ
        lbl����.Caption = "��:" & NVL(!��Ժ����)
        lbl����.Caption = NVL(!����)
        lbl����.ForeColor = NVL(mrsPati!��ɫ, 0)
        lbl�Ա�.Caption = NVL(!�Ա�)
        lbl����.Caption = NVL(!����)
        
        lblסԺ��.Caption = "סԺ��:" & NVL(!סԺ��)
        lbl����ȼ�.Caption = NVL(!����ȼ�)
        lblҽ�Ƹ��ʽ.Caption = NVL(!ҽ�Ƹ��ʽ)

        'Σ�ز��˲�����ɫ��ʾ
        lbl����.Caption = NVL(!��ǰ����)
        If NVL(!��ǰ����) = "Σ" Or NVL(!��ǰ����) = "��" Or NVL(!��ǰ����) = "��" Then
            lbl����.ForeColor = &HC0&
        Else
            lbl����.ForeColor = lblסԺ��.ForeColor
        End If

        lbl��Ժʱ��.Caption = Format(!��Ժ����, "yyyy-MM-dd HH:mm")
        If Not IsNull(!��Ժ����) Then
            lbl��Ժʱ��.Caption = lbl��Ժʱ��.Caption & "��" & Format(!��Ժ����, "yyyy-MM-dd HH:mm")
        End If

        lbl��������.Caption = NVL(!��������, "��ͨ����")
        If NVL(!ҽ����) <> "" Then lbl��������.Caption = lbl��������.Caption & "[" & NVL(!ҽ����) & "]"

        '���
        lbl���.Caption = "���:" & GetPatiDiagnose(mlng����ID, mlng��ҳID, 2)

        '������Ϣ
        mPatiInfo.���� = mrsPati!����
        mPatiInfo.״̬ = NVL(!״̬, 0)
        mPatiInfo.סԺ�� = NVL(!סԺ��)
        mPatiInfo.���� = NVL(!��Ժ����)
        mPatiInfo.��ҳID = mlng��ҳID
        mPatiInfo.����ID = NVL(!��ǰ����ID, 0)
        mPatiInfo.����ID = NVL(!��Ժ����ID, 0)
        mPatiInfo.��Ժ���� = !��Ժ����
        If Not IsNull(!��Ժ����) Then
            mPatiInfo.��Ժ���� = !��Ժ����
        Else
            mPatiInfo.��Ժ���� = CDate(0)
        End If
        mPatiInfo.����ת�� = NVL(!����ת��, 0) <> 0
        mPatiInfo.����״̬ = Val(NVL(!����״̬, 0))
        
        mlngӤ������ID = Val(!Ӥ������ID & "")
        mlngӤ������ID = Val(!Ӥ������ID & "")
    End With


    '������Ϣȡ��ǰסԺ������
    strSQL = "Select B.״̬,Decode(b.���ʱ��,NULL,b.��Ժ����,b.���ʱ��) AS ��Ժ���� , b.��Ժ����,B.סԺ��,b.��Ժ����,B.��������,B.����ת��,B.����,b.��ǰ����id,B.��Ժ����ID,B.��ǰ����ID,Decode(Nvl(X.�������, 0), 0, '��', '') As ����" & _
        " From ������ҳ B,������� X" & _
        " Where B.����ID=[1] And B.��ҳID=[2] And B.����ID = X.����ID(+) And X.����(+) = 1 And X.����(+)=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    With rsTmp
        mPatiInfo.���� = Val("" & !����)
        mPatiInfo.���� = Not IsNull(!����)
        mPatiInfo.���� = NVL(!��������, 0)
        mPatiInfo.���� = Sys.DeptHaveProperty(Val(!��Ժ����ID & ""), "����")
    End With
    
    '�����������ȵ������ؼ�λ�ü���С
    Me.pic��ʶ.Width = lbl����.Width + lbl����.Left
    Me.picסԺ����.Width = Me.cboPages.Width - 50
    Me.picסԺ����.Left = pic��ʶ.Left + pic��ʶ.Width + 50
    Me.cmdWarrant.Left = Me.picסԺ����.Left + Me.picסԺ����.Width + 50
    Me.img��ϸ��Ϣ.Left = Me.cmdWarrant.Left + Me.cmdWarrant.Width + 100
    picCondition.Width = Me.img��ϸ��Ϣ.Left + Me.img��ϸ��Ϣ.Width + 100
    
    '��ȡ���廤�����Ͳ���ID
    Call GeNurseRelatedUnitID
    '��ȡ���˷�����Ϣ
    Call mclsAdvices_StatusTextUpdate("")
    
    'ˢ���Ӵ�������
    Call SubWinRefreshData(tbcSub.Selected)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mclsAdvices_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
'���ܣ��鿴���Ӳ�������
    Call gobjRichEPR.ViewDocument(Me, ����ID, CanPrint)
End Sub

Private Sub mclspath_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
'���ܣ��ٴ�·���в鿴���Ӳ�������
    Call gobjRichEPR.ViewDocument(Me, ����ID, CanPrint)
End Sub

Private Sub mclsTends_RefreshPrompt(ByVal strInfo As String, ByVal blnImportant As Boolean)
'    lblPrompt.Caption = strInfo
'    lblPrompt.ForeColor = IIf(blnImportant, &HFF&, &H80000008)
End Sub

Private Sub picCondition_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCondition.hwnd, ""
End Sub

Private Sub pic��ʶ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo pic��ʶ.hwnd, ""
End Sub

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call cmdFilterCancel_Click
    End If
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If rptPati.Records.Count = 0 Then Exit Sub
    If rptPati.FocusedRow.Record Is Nothing Then Exit Sub
    
    mlng����ID = Split(rptPati.FocusedRow.Record.Tag, ",")(0)
    mlng��ҳID = Split(rptPati.FocusedRow.Record.Tag, ",")(1)
    '�����Ҫ���˶�λ����һ��,��һ��ʱ����λǰ��˳��,�ɰѸ�������ε�
    mrsPati.MoveFirst
    mrsPati.Find "Key='" & mlng����ID & ":" & mlng��ҳID & "'"
    
    picPati.Visible = False
    txt����.Text = ""
    mintPrePage = -1
    Call AddPages
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptPati_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub rptPati_SelectionChanged()
    '59268:������,2013-04-23,Ĭ��������չ�����з��飬���ڲ��Ҳ����㡣����ʽΪ�������������չ����һ��
    Dim objRow As ReportRow
    If rptPati.FocusedRow Is Nothing Then Exit Sub
    If rptPati.FocusedRow.GroupRow = True Then
        mlngRowIndex = rptPati.FocusedRow.Index
        For Each objRow In rptPati.Rows
            If objRow.GroupRow = True Then
                If objRow.Index = rptPati.FocusedRow.Index Then
                    Exit For
                ElseIf objRow.Expanded = True Then
                    mlngRowIndex = mlngRowIndex - objRow.Childs.Count
                End If
            End If
        Next
    End If
End Sub

Private Sub rptPati_SortOrderChanged()
    '59268:������,2013-04-23,Ĭ��������չ�����з��飬���ڲ��Ҳ����㡣����ʽΪ�������������չ����һ��
    Dim lngloop As Long
    Dim objRow As ReportRow
    Dim lng����ID As Long
    If rptPati.FocusedRow Is Nothing Then
        '�۵�������(ѡ�в�����һ�鲻�۵�)
        For Each objRow In rptPati.Rows
            If mlngRowIndex >= 0 And mlngRowIndex <= rptPati.Rows.Count Then
                If objRow.GroupRow And objRow.Index <> mlngRowIndex Then
                    objRow.Expanded = False
                End If
            End If
        Next
    Else
        If rptPati.FocusedRow Is Nothing Then Exit Sub
        lng����ID = rptPati.FocusedRow.Record.Item(C_����ID).Value
        'ѡ�е�ǰ����(���۵���Ļ�,Rows.Countֻ����ĸ�����,�����ȶ�λ,���۵�)
        For lngloop = 0 To rptPati.Rows.Count - 1
            If Not (rptPati.Rows(lngloop).Record Is Nothing) Then
                If Val(rptPati.Rows(lngloop).Record.Item(C_����ID).Value) = lng����ID Then
                    Set rptPati.FocusedRow = rptPati.Rows(lngloop)
                    Exit For
                End If
            End If
        Next
        
        '�۵�������(ѡ�в�����һ�鲻�۵�)
        For Each objRow In rptPati.Rows
            If objRow.GroupRow And objRow.Index <> rptPati.FocusedRow.ParentRow.Index Then
                objRow.Expanded = False
            End If
        Next
    End If
    If Not rptPati.FocusedRow Is Nothing Then rptPati.FocusedRow.EnsureVisible
    If rptPati.Visible Then rptPati.SetFocus
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "������ɫ" Then
        Call zlDatabase.ShowPatiColorTip(Me)
    End If
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'���ܣ�ˢ���Ӵ�����漰����
'˵����������Ϊ�л����濨Ƭ����
    Dim Index As Long, objItem As TabControlItem
    If mblnTabTmp Then Exit Sub
    If Item.Tag = "" Then Exit Sub '��ʼ��ʱ,��û��ֵ
    If Item.Tag <> tbcSub.Tag Then Call UnLoadPageForm '����һ��ҳ��GDI�ͻ�������Ϊ�˿���GDI�������л�ҳ��ʱж����һ��ҳ�洰��
    If Item.Handle = picTmp.hwnd Then
        Index = Item.Index
        mblnTabTmp = True
        Screen.MousePointer = 11
        On Error GoTo errH
        Select Case Item.Tag
            Case "·��"
                Set objItem = tbcSub.InsertItem(Index, "�ٴ�·��", mcolSubForm("_·��").hwnd, 0)
                objItem.Tag = "·��"
            Case "ҽ��"
                Set objItem = tbcSub.InsertItem(Index, "ҽ����¼", mcolSubForm("_ҽ��").hwnd, 0)
                objItem.Tag = "ҽ��"
            Case "����"
                Set objItem = tbcSub.InsertItem(Index, "���ü�¼", mcolSubForm("_����").hwnd, 0)
                objItem.Tag = "����"
            Case "����"
                Set objItem = tbcSub.InsertItem(Index, "סԺ����", mcolSubForm("_����").hwnd, 0)
                objItem.Tag = "����"
            Case "����"
                Set objItem = tbcSub.InsertItem(Index, "�����¼", mcolSubForm("_����").hwnd, 0)
                objItem.Tag = "����"
            Case "������"
                Set objItem = tbcSub.InsertItem(Index, "������", mcolSubForm("_������").hwnd, 0)
                objItem.Tag = "������"
            Case "�໤"
                Set objItem = tbcSub.InsertItem(Index, "����໤", mcolSubForm("_�໤").hwnd, 0)
                objItem.Tag = "�໤"
            Case "�²���"
                Set objItem = tbcSub.InsertItem(Index, "���Ӳ���", mcolSubForm("_�²���").hwnd, 0)
                objItem.Tag = "�²���"
            Case Else '���廤��ҳ��
                Set mobjNurseForm = gobjNurseIntegrate.GetForm(Item.Tag, CStr(mColNurseFormUrl("_" & Item.Tag)))
                Set objItem = tbcSub.InsertItem(Index, Item.Tag, mobjNurseForm.hwnd, 0)
                objItem.Tag = Item.Tag
        End Select
        Call tbcSub.RemoveItem(Index + 1)
        objItem.Selected = True
        Screen.MousePointer = 0
        mblnTabTmp = False
    Else
        Set objItem = Item
    End If
    
    'ˢ���Ӵ����Ӧ��CommandBar
    Call SubWinDefCommandBar(objItem)

    'ˢ���Ӵ�������
    If Visible Then Call SubWinRefreshData(objItem)

    If Visible And mfrmActive.Visible And mfrmActive.Enabled Then mfrmActive.SetFocus
    tbcSub.Tag = Item.Tag   '��¼��һ��ѡ��Ŀ�Ƭ
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub UnLoadPageForm()
'����һ��ҳ��GDI�ͻ�������Ϊ�˿���GDI�������л�ҳ��ʱж����һ��ҳ�洰��(�°���Ӳ���������)
'��ҽ���еĴ�����ֱ�Ӱ󶨵�Ҳ���ô���
    Dim i As Integer, blnFind As Boolean
    Dim Index As Long, objItem As TabControlItem
    Dim blnNurseIntegrate As Boolean
    '�ҵ���һ��ѡ��ҳ�������
    For i = 0 To tbcSub.ItemCount - 1
        If tbcSub.Item(i).Tag = tbcSub.Tag Then
            Index = tbcSub.Item(i).Index
            blnFind = True
            Exit For
        End If
    Next i
    If blnFind = False Then Exit Sub
    blnNurseIntegrate = Val(marrTabAttribute(Index)) = 1
    '�ſ��°没������ҽӿڴ���(Val(marrTabAttribute(Index)) = 1,Ϊ���廤����)
    If InStr(1, "'·��'ҽ��'����'����'����'������'�໤'", "'" & tbcSub.Tag & "'") = 0 And blnNurseIntegrate = False Then Exit Sub
    '128211��1:��ҽ���л�������ҳ�棬�ڶ�λ�����ı����ûس�����ͻ�ʧȥ���㣬����1����ԭ������ҽ������ռ��GDIҲ���㣬��ʱ������ж��
    '              2:��������ֱ���ſ�һ������Ϊ��tabҳ����ɫ����һ����Ϊ�˱�֤����ҳ���л���ҽ������ı�����ȫѡ���ݣ�ҽ�����治֪����ô����
    If tbcSub.Tag <> "ҽ��" Then
        If UnloadSubForm(tbcSub.Tag, blnNurseIntegrate) = False Then Exit Sub
    End If
    
    Screen.MousePointer = 11
    mblnTabTmp = True
    On Error GoTo ErrHand
    Select Case tbcSub.Tag
        Case "·��"
            Set objItem = tbcSub.InsertItem(Index, "�ٴ�·��", picTmp.hwnd, 0)
            objItem.Tag = "·��"
        Case "ҽ��"
            Set objItem = tbcSub.InsertItem(Index, "ҽ����¼", picTmp.hwnd, 0)
            objItem.Tag = "ҽ��"
        Case "����"
            Set objItem = tbcSub.InsertItem(Index, "���ü�¼", picTmp.hwnd, 0)
            objItem.Tag = "����"
        Case "����"
            Set objItem = tbcSub.InsertItem(Index, "סԺ����", picTmp.hwnd, 0)
            objItem.Tag = "����"
        Case "����"
            Set objItem = tbcSub.InsertItem(Index, "�����¼", picTmp.hwnd, 0)
            objItem.Tag = "����"
        Case "������"
            Set objItem = tbcSub.InsertItem(Index, "������", picTmp.hwnd, 0)
            objItem.Tag = "������"
        Case "�໤"
            Set objItem = tbcSub.InsertItem(Index, "����໤", picTmp.hwnd, 0)
            objItem.Tag = "�໤"
        Case Else '���廤��
            Set objItem = tbcSub.InsertItem(Index, tbcSub.Tag, picTmp.hwnd, 0)
            objItem.Tag = tbcSub.Tag
    End Select
    Call tbcSub.RemoveItem(Index + 1)
    Screen.MousePointer = 0
    mblnTabTmp = False
    Exit Sub
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function UnloadSubForm(ByVal strTag As String, Optional blnNurseIntegrate As Boolean = False) As Boolean
'���ܣ�ж��������ⴰ��
'������strTag�������廤����ҳǩ��; blnNurseIntegrate �Ƿ������廤��ҳǩ
    Dim objForm As Object
    On Error Resume Next
    err.Clear
    If blnNurseIntegrate = False Then
        If Not mcolSubForm("_" & strTag) Is Nothing Then
            Call SetOrGetSubFromOperation(strTag, False)  '����ж��֮ǰ��¼������������
            Unload mcolSubForm("_" & strTag)
        End If
    Else
        If Not mobjNurseForm Is Nothing Then Unload mobjNurseForm: Set mobjNurseForm = Nothing
    End If
    If err <> 0 Then err.Clear
    UnloadSubForm = True
    On Error GoTo 0
End Function

Private Sub SetOrGetSubFromOperation(ByVal strTag As String, ByVal blnSet As Boolean)
'���û��ȡ�Ӵ�������,�������ģ���ṩͳһ�ӿ�
'       GetFormOperation() as string --��ȡ�������ѡ�񣬸ýӿڻ��ڴ���ж��ǰ����
'       RestoreFormOperation(byval strValue as string)-�ָ��������ѡ�񣬸ýӿڻ������ⴰ��ˢ��ǰ����
'blnSet =TRUE �ָ��Ӵ�����������(ˢ��ǰ����),=FALSE ��ȡ�Ӵ�����������(����ж��ǰ����)
    Dim strValue As String
    On Error Resume Next
    If blnSet = False Then mcolSubFormOperation.Remove "_" & strTag
    Select Case strTag
        Case "·��"
            If blnSet = False Then
                strValue = mclsPath.GetFormOperation()
                mcolSubFormOperation.Add strValue, "_" & strTag
            Else
                strValue = CStr(mcolSubFormOperation("_" & strTag))
                If strValue <> "" Then Call mclsPath.RestoreFormOperation(strValue)
            End If
        Case "ҽ��"
            If blnSet = False Then
                strValue = mclsAdvices.GetFormOperation()
                mcolSubFormOperation.Add strValue, "_" & strTag
            Else
                strValue = CStr(mcolSubFormOperation("_" & strTag))
                If strValue <> "" Then Call mclsAdvices.RestoreFormOperation(strValue)
            End If
        Case "����"
            If blnSet = False Then
                strValue = mclsFeeQuery.GetFormOperation()
                mcolSubFormOperation.Add strValue, "_" & strTag
            Else
                strValue = CStr(mcolSubFormOperation("_" & strTag))
                If strValue <> "" Then Call mclsFeeQuery.RestoreFormOperation(strValue)
            End If
        Case "����"   'סԺ����
            If blnSet = False Then
                strValue = mclsEPRs.GetFormOperation()
                mcolSubFormOperation.Add strValue, "_" & strTag
            Else
                strValue = CStr(mcolSubFormOperation("_" & strTag))
                If strValue <> "" Then Call mclsEPRs.RestoreFormOperation(strValue)
            End If
        Case "����"
            If blnSet = False Then
                strValue = mclsTends.GetFormOperation()
                mcolSubFormOperation.Add strValue, "_" & strTag
            Else
                strValue = CStr(mcolSubFormOperation("_" & strTag))
                If strValue <> "" Then Call mclsTends.RestoreFormOperation(strValue)
            End If
        Case "������"
            If blnSet = False Then
                strValue = mclsTendEPRs.GetFormOperation()
                mcolSubFormOperation.Add strValue, "_" & strTag
            Else
                strValue = CStr(mcolSubFormOperation("_" & strTag))
                If strValue <> "" Then Call mclsTendEPRs.RestoreFormOperation(strValue)
            End If
    End Select
    If blnSet = True Then mcolSubFormOperation.Remove "_" & strTag
    
    If err <> 0 Then err.Clear
    On Error GoTo 0
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With Me.tbcSub
        .Left = lngLeft: .Width = lngRight - lngLeft
        .Top = lngTop: .Height = lngBottom - lngTop
    End With
    
    With picPrompt
        .Top = stbThis.Top + 50
        .Height = stbThis.Height - 100
        .Left = stbThis.Panels(2).Left + 50
        .Width = stbThis.Panels(2).Width - 100
    End With
    With lblPrompt
        .Width = picPrompt.Width
        .Height = TextHeight("��")
        .Top = (picPrompt.Height - .Height) \ 2
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, strTmp As String
    Dim blnSetup As Boolean
    
    mlng����ID = 0
    mlng��ҳID = 0
    mlng����ID = 0
    mblnShow = False
    
    blnSetup = InStr(";" & mstrPrivs & ";", ";��������;") > 0
    If Not tbcSub.Selected Is Nothing Then
        Call zlDatabase.SetPara("ҽ������", tbcSub.Selected.Tag, glngSys, pסԺ��ʿվ, blnSetup)
    End If
    Call zlDatabase.SetPara("������Ϣ������", IIf(mobjBar.Visible, 1, 0), glngSys, pסԺ��ʿվ, blnSetup)
    Call SaveWinState(Me, App.ProductName)

    'ǿ��Unload,��Ȼ���ἤ���Ӵ�����¼�
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    Set mcolSubForm = Nothing
    Set mcolSubFormOperation = Nothing
    If Not mobjNurseForm Is Nothing Then
        Unload mobjNurseForm
        Set mobjNurseForm = Nothing
    End If
    Set mColNurseFormUrl = Nothing
    Set mclsAdvices = Nothing
    Set mclsEMR = Nothing
    Set mclsEPRs = Nothing
    Set mclsTends = Nothing
    Set mclsTendEPRs = Nothing
    Set mclsFeeQuery = Nothing
    Set mclsWardMonitor = Nothing
    Set mclsPath = Nothing
    
    Set mfrmActive = Nothing
    Set mobjMipModule = Nothing
    
    mobjParent.mblnRoutine = mblnShow
    If Not mobjParent Is Nothing Then Set mobjParent = Nothing
    If Not mrsPati Is Nothing Then
        If mrsPati.State = adStateClosed Then mrsPati.Close
        Set mrsPati = Nothing
    End If
End Sub

Private Sub picInfo_GotFocus()
    If cboPages.Enabled And cboPages.Visible Then cboPages.SetFocus
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraInfo.Width = picInfo.ScaleWidth - fraInfo.Left * 2

    cbo����.Width = fraInfo.Width - cbo����.Left - 100
End Sub

Private Sub tbcSub_GotFocus()
    On Error Resume Next
    If Not mfrmActive Is Nothing Then mfrmActive.SetFocus
End Sub

Private Sub ClearPatiInfo()
'���ܣ��������������ص���ʾ��Ϣ
    mlng����ID = 0
    mlng��ҳID = 0
    mlngӤ������ID = 0
    mlngӤ������ID = 0
    
    mPatiInfo.״̬ = 0
    mPatiInfo.סԺ�� = ""
    mPatiInfo.���� = ""
    mPatiInfo.��ҳID = 0
    mPatiInfo.����ID = 0
    mPatiInfo.����ID = 0
    mPatiInfo.��Ժ���� = CDate(0)
    mPatiInfo.��Ժ���� = CDate(0)
    mPatiInfo.����ת�� = False
    mPatiInfo.���� = False
    mPatiInfo.���� = False
    mPatiInfo.���� = 0
    mPatiInfo.���� = 0

    cboPages.Clear
    cbo����.Clear

    lbl����.Caption = ""
    lbl����.Caption = ""
    lbl�Ա�.Caption = ""
    lbl����.Caption = ""
    lblסԺ��.Caption = ""
    lblҽ�Ƹ��ʽ.Caption = ""
    lbl����ȼ�.Caption = ""
    lbl����.Caption = ""
    lbl��Ժʱ��.Caption = ""
    lbl���.Caption = ""
End Sub

Function ExecuteMonitor() As Boolean
'���ܣ����ü໤��
    Dim strUser As String, strPass As String, strServer As String
    Dim arrInfo As Variant, i As Long

    'Provider=MSDataShape.1;Extended Properties="Driver={Microsoft ODBC for Oracle};Server=ORCL";Persist Security Info=True;User ID=zlhis;Password=HIS;Data Provider=MSDASQL
    'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source=ORCL;Extended Properties="PLSQLRSet=1;DistribTx=0"
    arrInfo = Split(gcnOracle.ConnectionString, ";")
    For i = 0 To UBound(arrInfo)
        If UCase(arrInfo(i)) Like UCase("User ID=*") Then
            strUser = Mid(arrInfo(i), InStr(arrInfo(i), "=") + 1)
        ElseIf UCase(arrInfo(i)) Like UCase("Password=*") Then
            strPass = Mid(arrInfo(i), InStr(arrInfo(i), "=") + 1)
        ElseIf UCase(arrInfo(i)) Like UCase("Data Source=*") Then
            strServer = Mid(arrInfo(i), InStr(arrInfo(i), "=") + 1)
        ElseIf UCase(arrInfo(i)) Like UCase("Server=*") Then
            strServer = Mid(arrInfo(i), InStr(arrInfo(i), "=") + 1)
            strServer = Replace(strServer, """", "")
        End If
    Next

    On Error GoTo errH

    Shell mstrMonitor & " " & strUser & " " & strPass & " " & strServer, vbNormalFocus

    ExecuteMonitor = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AddPages()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long, lng��ҳID As Long
    Dim bln���� As Boolean
    On Error GoTo ErrHand
    
    '������������Ϣ
    lng����ID = mlng����ID: lng��ҳID = mlng��ҳID
    Call ClearPatiInfo
    mlng����ID = lng����ID: mlng��ҳID = lng��ҳID
    '���ݲ���ID��ȡ�ò��˵�סԺ����
    '52004,������,2012-08-10,סԺ����Ӧ��Ĭ�϶�λ����ǰ���˵�ǰסԺ����
    strSQL = " Select ��ҳID,�������� From ������ҳ Where ��ҳID<>0 And ����ID=[1] Order by ��ҳID Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡסԺ����", mlng����ID)
    
    cboPages.Clear
    Do While Not rsTemp.EOF
        cboPages.AddItem "�� " & rsTemp!��ҳID & " ��" & IIf(Val("" & rsTemp!��������) = 1, "(��������)", IIf(Val("" & rsTemp!��������) = 2, "(סԺ����)", ""))
        cboPages.ItemData(cboPages.NewIndex) = rsTemp!��ҳID
        If rsTemp!��ҳID = mlng��ҳID Then
            Call Cbo.SetIndex(cboPages.hwnd, cboPages.NewIndex)
        End If
        If bln���� = False And Val("" & rsTemp!��������) <> 0 Then bln���� = True
        rsTemp.MoveNext
    Loop
    If cboPages.ListIndex = -1 Then
        Call Cbo.SetIndex(cboPages.hwnd, 0)
    End If
    If bln���� = True Then
        Call Cbo.SetListWidth(cboPages.hwnd, 2000)
    End If
    Call cboPages_Click
    '52638,������,2012-08-13,���ز��˹���ҩ����Ϣ
    Call LoadPatiAllergy(mlng����ID, cbo����)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInput As String
    Dim strOrder As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    strInput = Trim(txt����.Text)
    If strInput = "" Then Exit Sub
    
    strOrder = mrsPati.Sort
    strInput = " ����='" & zlStr.Lpad(strInput, 10) & "'"
    mrsPati.Filter = strInput
    If mrsPati.RecordCount = 0 Then
        If Not IsNumeric(Trim(txt����.Text)) Then
            strInput = " ����='" & Trim(txt����.Text) & "'"
        Else
            strInput = " סԺ��=" & Trim(txt����.Text)
        End If
        mrsPati.Filter = strInput
        
        If mrsPati.RecordCount = 0 Then
            '�ٰ������������һ��,���ṩ����ѡ��Ĺ���,Ҫ�󾡿���������ϸ
            mrsPati.Sort = "����"
            mrsPati.Filter = "���� LIKE '*" & UCase(Trim(txt����.Text)) & "*'"
            If mrsPati.RecordCount = 0 Then
                mrsPati.Filter = 0
                mrsPati.Sort = strOrder
                MsgBox "δ�ҵ��ò��˵���Ч���ݣ����������룡", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    mlng����ID = mrsPati!����ID
    mlng��ҳID = mrsPati!��ҳID
    mrsPati.Filter = 0
    mrsPati.Sort = strOrder
    mrsPati.MoveFirst
    mrsPati.Find "Key='" & mlng����ID & ":" & mlng��ҳID & "'"
    
    mintPrePage = -1
    Call AddPages
    
    picPati.Visible = False
End Sub

Private Sub mclsAdvices_DoByAdvice(ByVal lngҽ��ID As Long, ByVal lng���ID As Long, ByVal lngWayID As Long, ByVal strTag As String)
'���ܣ���ҽ������  lngWayID��conMenu_Edit_AdvicePrice
    Dim lngTmp As Long
    lngTmp = IIf(lng���ID = 0, lngҽ��ID, lng���ID)
    Call mclsFeeQuery.zlPatiBilling(Me, mlng����ID, mlng����ID, mlng��ҳID, Val("" & mPatiInfo.����ID), False, lngTmp)
End Sub

Private Sub GeNurseRelatedUnitID()
    Dim strErrMsg As String
    '�����л��ǻ�ȡ
    If gbln�������廤��ӿ� = True Then
        If InitNurseIntegrate = True Then
            If gobjNurseIntegrate.GetRelatedIDToGUID(mlng����ID, strErrMsg, mlng����ID & "|" & mlng��ҳID) = False Then
                MsgBox "��ȡ���廤����ID�ӿڵ���ʧ�ܣ�" & vbCrLf & "��ϸ��Ϣ��" & strErrMsg, vbInformation, gstrSysName
            Else
                mstrRelatedUnitID = gobjNurseIntegrate.RelatedUnitID
                mstrRelatedUserID = gobjNurseIntegrate.RelatedUserID
                mstrNurseParentID = gobjNurseIntegrate.RelatedPatientID
            End If
        End If
    End If
End Sub

