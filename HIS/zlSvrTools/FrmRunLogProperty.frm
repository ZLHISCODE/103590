VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRunLogProperty 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "������־����"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab SSTab1 
      Height          =   3165
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   5583
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "����(&P)"
      TabPicture(0)   =   "FrmRunLogProperty.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Txtģ����"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lblģ����"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Txt�û���"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lbl�û���"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Txt�˳�ԭ��"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lbl�˳�ԭ��"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Txt������"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Lbl������"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Txt�˳�ʱ��"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Lbl�˳�ʱ��"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Txt����վ"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Lbl����վ"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Txt����ʱ��"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Lbl����ʱ��"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Txt�Ự��"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Lbl�Ự��"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.Label Lbl�Ự�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ự��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   17
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Txt�Ự�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   16
         Top             =   450
         Width           =   3150
      End
      Begin VB.Label Lbl����ʱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   15
         Top             =   2340
         Width           =   720
      End
      Begin VB.Label Txt����ʱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   14
         Top             =   2340
         Width           =   3150
      End
      Begin VB.Label Lbl����վ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����վ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   13
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Txt����վ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   12
         Top             =   1080
         Width           =   3150
      End
      Begin VB.Label Lbl�˳�ʱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�˳�ʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   11
         Top             =   2670
         Width           =   720
      End
      Begin VB.Label Txt�˳�ʱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   10
         Top             =   2670
         Width           =   3150
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   9
         Top             =   1380
         Width           =   540
      End
      Begin VB.Label Txt������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   8
         Top             =   1380
         Width           =   3150
      End
      Begin VB.Label Lbl�˳�ԭ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�˳�ԭ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   7
         Top             =   780
         Width           =   720
      End
      Begin VB.Label Txt�˳�ԭ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   6
         Top             =   780
         Width           =   3150
      End
      Begin VB.Label Lbl�û��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�û���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   5
         Top             =   1710
         Width           =   540
      End
      Begin VB.Label Txt�û��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   4
         Top             =   1710
         Width           =   3150
      End
      Begin VB.Label Lblģ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ģ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   3
         Top             =   2010
         Width           =   540
      End
      Begin VB.Label Txtģ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   2
         Top             =   2010
         Width           =   3150
      End
   End
   Begin VB.CommandButton Cmd�˳� 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2970
      TabIndex        =   0
      Top             =   3330
      Width           =   1100
   End
End
Attribute VB_Name = "FrmRunLogProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd�˳�_Click()
    Unload Me
End Sub
