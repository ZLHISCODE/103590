VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm���ñ��� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ñ���"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "frm���ñ���.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6390
      TabIndex        =   63
      Top             =   1410
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6390
      TabIndex        =   62
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmd�ֵ���ϸ 
      Caption         =   "�ֵ���ϸ"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6390
      TabIndex        =   64
      Top             =   4920
      Width           =   1100
   End
   Begin VB.CommandButton cmd������� 
      Caption         =   "Ԥ����(&Y)"
      Height          =   350
      Left            =   6390
      TabIndex        =   61
      Top             =   480
      Width           =   1100
   End
   Begin VB.Frame fra 
      Caption         =   "����"
      Height          =   5295
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   5955
      Begin VB.ComboBox cbo��ҽ 
         Height          =   300
         Index           =   0
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   330
         Width           =   1635
      End
      Begin VB.TextBox txt�ʻ�֧�� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   4110
         TabIndex        =   29
         Top             =   4800
         Width           =   1365
      End
      Begin VB.TextBox txtͳ��֧�� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1470
         TabIndex        =   27
         Top             =   4800
         Width           =   1365
      End
      Begin VB.TextBox txt�����Ը� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   4110
         TabIndex        =   25
         Top             =   4410
         Width           =   1365
      End
      Begin VB.TextBox txtȫ�Ը� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1470
         TabIndex        =   23
         Top             =   4410
         Width           =   1365
      End
      Begin VB.TextBox txt����ͳ�� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   4110
         TabIndex        =   21
         Top             =   4020
         Width           =   1365
      End
      Begin VB.TextBox txt�����ܶ� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1470
         TabIndex        =   19
         Top             =   4020
         Width           =   1365
      End
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   1995
         Index           =   0
         Left            =   450
         TabIndex        =   17
         Top             =   1920
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   3519
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.TextBox txt��� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   16
         Top             =   1500
         Width           =   1635
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   240
         Index           =   0
         Left            =   5190
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   750
         Width           =   255
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Index           =   0
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   12
         Top             =   1110
         Width           =   1605
      End
      Begin VB.TextBox txt״̬ 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   3840
         TabIndex        =   14
         Top             =   1110
         Width           =   1635
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   1605
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   240
         Index           =   0
         Left            =   2520
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Top             =   330
         Width           =   1605
      End
      Begin VB.Label lbl��ҽ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҽ(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3120
         TabIndex        =   4
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl�ʻ�֧�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ʻ�֧��(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3060
         TabIndex        =   28
         Top             =   4860
         Width           =   990
      End
      Begin VB.Label lblͳ��֧�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ��֧��(&Z)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   26
         Top             =   4860
         Width           =   990
      End
      Begin VB.Label lbl�����Ը� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ը�(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3060
         TabIndex        =   24
         Top             =   4470
         Width           =   990
      End
      Begin VB.Label lblȫ�Ը� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ȫ�Ը�(&Q)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   22
         Top             =   4470
         Width           =   810
      End
      Begin VB.Label lbl����ͳ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ͳ��(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3060
         TabIndex        =   20
         Top             =   4080
         Width           =   990
      End
      Begin VB.Label lbl�����ܶ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ܶ�(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   18
         Top             =   4080
         Width           =   990
      End
      Begin VB.Label lbl��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   15
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3120
         TabIndex        =   8
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   11
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label lbl״̬ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "״̬(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3120
         TabIndex        =   13
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   6
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
   End
   Begin VB.Frame fra 
      Caption         =   "סԺ"
      Height          =   5295
      Index           =   1
      Left            =   180
      TabIndex        =   31
      Top             =   120
      Width           =   5955
      Begin VB.ComboBox cbo��ҽ 
         Height          =   300
         Index           =   1
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   330
         Width           =   1635
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   300
         Index           =   1
         Left            =   3840
         TabIndex        =   43
         Top             =   1110
         Width           =   1635
      End
      Begin VB.TextBox txtסԺ���� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   3840
         TabIndex        =   39
         Top             =   720
         Width           =   1635
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   240
         Index           =   1
         Left            =   2520
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   37
         Top             =   720
         Width           =   1605
      End
      Begin VB.TextBox txt״̬ 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   45
         Top             =   1500
         Width           =   1605
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   41
         Top             =   1110
         Width           =   1605
      End
      Begin VB.TextBox txt��� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   3840
         TabIndex        =   47
         Top             =   1500
         Width           =   1635
      End
      Begin VB.TextBox txt�����ܶ� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1470
         TabIndex        =   50
         Top             =   4020
         Width           =   1365
      End
      Begin VB.TextBox txt����ͳ�� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4110
         TabIndex        =   52
         Top             =   4020
         Width           =   1365
      End
      Begin VB.TextBox txtȫ�Ը� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1470
         TabIndex        =   54
         Top             =   4410
         Width           =   1365
      End
      Begin VB.TextBox txt�����Ը� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4110
         TabIndex        =   56
         Top             =   4410
         Width           =   1365
      End
      Begin VB.TextBox txtͳ��֧�� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1470
         TabIndex        =   58
         Top             =   4800
         Width           =   1365
      End
      Begin VB.TextBox txtͳ���Ը� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4110
         TabIndex        =   60
         Top             =   4800
         Width           =   1365
      End
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   1995
         Index           =   1
         Left            =   450
         TabIndex        =   48
         Top             =   1920
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   3519
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   32
         Top             =   330
         Width           =   1605
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   2940
         TabIndex        =   42
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lbl��ҽ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҽ(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3120
         TabIndex        =   34
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lblסԺ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3120
         TabIndex        =   38
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   30
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   36
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl״̬ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "״̬(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   44
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   40
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label lbl��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3120
         TabIndex        =   46
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lbl�����ܶ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ܶ�(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   420
         TabIndex        =   49
         Top             =   4080
         Width           =   990
      End
      Begin VB.Label lbl����ͳ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ͳ��(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3060
         TabIndex        =   51
         Top             =   4080
         Width           =   990
      End
      Begin VB.Label lblȫ�Ը� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ȫ�Ը�(&Q)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   53
         Top             =   4470
         Width           =   810
      End
      Begin VB.Label lbl�����Ը� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ը�(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3060
         TabIndex        =   55
         Top             =   4470
         Width           =   990
      End
      Begin VB.Label lblͳ��֧�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ��֧��(&Z)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   420
         TabIndex        =   57
         Top             =   4860
         Width           =   990
      End
      Begin VB.Label lblͳ���Ը� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ���Ը�(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3060
         TabIndex        =   59
         Top             =   4860
         Width           =   990
      End
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��˾ְ��"
      BeginProperty Font 
         Name            =   "�����п�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   390
      Left            =   6180
      TabIndex        =   65
      Top             =   3150
      Width           =   1440
   End
End
Attribute VB_Name = "frm���ñ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum FrameIndex
    int���� = 0
    intסԺ = 1
End Enum
Private blnModify As Boolean                '�Ƿ��޸ģ��޸����ֹȷ����ť
Private mint�༭ģʽ As Integer             '1-����;2-����
Private mlng��¼ID As Long
Private mblnסԺ As Boolean                 '����/סԺ
Private mblnOK As Boolean
Private Const strFormat_��� As String = "#####0.00;-#####0.00; ;"
Private Const strFormat_���� As String = "#####0.000;-#####0.000; ;"

Public Function ShowME(int�༭ģʽ As Integer, Optional blnסԺ As Boolean = False, Optional ByVal lng��¼ID As Long = 0) As Boolean
    On Error Resume Next
    mint�༭ģʽ = int�༭ģʽ
    mlng��¼ID = lng��¼ID
    mblnסԺ = blnסԺ
    mblnOK = False
    blnModify = False
    
    Me.Show 1
    ShowME = mblnOK
End Function

Private Sub Bill_BeforeDeleteRow(Index As Integer, Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Bill_EnterCell(Index As Integer, Row As Long, Col As Long)
    With Bill(Index)
        If Col = 0 Then Exit Sub
        
        If Row = .Rows - 1 Then
            .ColData(Col) = 0
        Else
            .ColData(Col) = IIf(Col = .Cols - 1, 5, 4)
        End If
    End With
End Sub

Private Sub Bill_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim StrInput As String
    Dim lngRow As Long
    Dim curMoney As Currency, curCount As Currency
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mint�༭ģʽ = 1 Then cmdȷ��.Enabled = False
    
    With Bill(Index)
        If .TxtVisible = False Then
            .Text = .TextMatrix(.Row, .Col)
            If .Text = "" Then .Text = " "
        Else
            StrInput = Format(.Text, IIf(.Col = 2 And mblnסԺ, strFormat_����, strFormat_���))
            
            If Trim(StrInput) = "" Then
                StrInput = " "
            Else
                If Not IsNumeric(StrInput) Then
                    MsgBox "�����к��зǷ��ַ���", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If .RowData(.Row) = 1 And .Col = 2 Then
                    If Not (Val(StrInput) >= 0 And Val(StrInput) < 100) Then
                        MsgBox "ʵ�ʱ�����������С�������ڵ���100%��", vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
            End If
            .Text = StrInput
        End If
        .TextMatrix(.Row, .Col) = .Text
        
        curMoney = 0: curCount = 0
        For lngRow = 1 To .Rows - 2
            curMoney = curMoney + Val(.TextMatrix(lngRow, 1))
            If mblnסԺ Then
                If .RowData(lngRow) = 0 Then
                    curCount = curCount + Val(.TextMatrix(lngRow, 2))
                Else
                    '��ʾ������Ǳ���
                End If
            Else
                curCount = curCount + Val(.TextMatrix(lngRow, 2))
            End If
        Next
        .TextMatrix(.Rows - 1, 1) = Format(curMoney, strFormat_���)
        .TextMatrix(.Rows - 1, 2) = Format(curCount, IIf(mblnסԺ, strFormat_����, strFormat_���))
    End With
End Sub

Private Sub cbo��ҽ_Click(Index As Integer)
    Dim intMax As Integer   '�������һ��סԺ����
    Dim rsTemp As New ADODB.Recordset
    '��ȡָ����ҽ��ʽ�����߱�׼
    If mint�༭ģʽ = 1 Then cmdȷ��.Enabled = False
    
    If Index <> intסԺ Then Exit Sub
    If Val(txt����(intסԺ).Tag) = 0 Then Exit Sub

    '�ٶ����ʻ������Ϣ
    If Val(txt����(Index).Tag) = 0 Then Exit Sub
    gstrSQL = "select * from �ʻ������Ϣ where ����=" & TYPE_�Ĵ�üɽ & _
        " and ����ID=" & Val(txt����(Index).Tag) & " and ���=" & Format(zlDatabase.Currentdate, "yyyy")
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.EOF = False Then
        '�����ʻ����
        If cbo��ҽ(Index).ListIndex = 0 Then
            txtסԺ����(Index).Text = Nvl(rsTemp("סԺ�����ۼ�"), 0) + IIf(mint�༭ģʽ = 1, 1, 0)
        Else
            txtסԺ����(Index).Text = Nvl(rsTemp("��ԺסԺ����"), 0) + IIf(mint�༭ģʽ = 1, 1, 0)
        End If
    Else
        txtסԺ����(Index).Text = 1
    End If
    
    gstrSQL = "select ��� from ���ձ������� " & _
             " Where ����=2 And ����=" & Val(txt����(intסԺ).Tag) & " And ��Ժ=" & cbo��ҽ(intסԺ).ListIndex + 1 & _
             " And ��Ⱥ=" & Val(txt״̬(intסԺ).Tag) & " And ���='" & Val(txtסԺ����(intסԺ)) & "'" & _
             " And ����=" & TYPE_�Ĵ�üɽ & " And ���=" & Format(zlDatabase.Currentdate, "yyyy")
    Call OpenRecordset(rsTemp, Me.Caption)
    If Not rsTemp.EOF Then
        txt����(intסԺ) = Format(rsTemp!���, strFormat_���)
    Else
        'ȡ���סԺ����ʱ���𸶽����Ϊ����סԺ����
        gstrSQL = "select Max(���) סԺ���� from ���ձ������� " & _
                 " Where ����=2 And ����=" & Val(txt����(intסԺ).Tag) & " And ��Ժ=" & cbo��ҽ(intסԺ).ListIndex + 1 & _
                 " And ��Ⱥ=" & Val(txt״̬(intסԺ).Tag) & " And ���<>'A'" & _
                 " And ����=" & TYPE_�Ĵ�üɽ & " And ���=" & Format(zlDatabase.Currentdate, "yyyy")
        Call OpenRecordset(rsTemp, Me.Caption)
        intMax = rsTemp!סԺ����
        
        gstrSQL = "select ��� from ���ձ������� " & _
                 " Where ����=2 And ����=" & Val(txt����(intסԺ).Tag) & " And ��Ժ=" & cbo��ҽ(intסԺ).ListIndex + 1 & _
                 " And ��Ⱥ=" & Val(txt״̬(intסԺ).Tag) & " And ���='" & intMax & "'" & _
                 " And ����=" & TYPE_�Ĵ�üɽ & " And ���=" & Format(zlDatabase.Currentdate, "yyyy")
        Call OpenRecordset(rsTemp, Me.Caption)
        txt����(intסԺ) = Format(rsTemp!���, strFormat_���)
    End If
    
    gComInfo_üɽ.���� = Val(txt����(intסԺ).Text)
End Sub

Private Sub cmd����_Click(Index As Integer)
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & TYPE_�Ĵ�üɽ
    
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "ҽ������", , txt����(Index).Text)
    If rsTemp.State <> 1 Then Exit Sub
    If Not rsTemp Is Nothing Then
        txt����(Index).Text = rsTemp("����")
        txt����(Index).Tag = rsTemp("ID")
        zlControl.TxtSelAll txt����(Index)
    End If
    txt����(Index).SetFocus
    
    If mint�༭ģʽ = 1 Then cmdȷ��.Enabled = False
End Sub

Private Sub cmd�ֵ���ϸ_Click()
    Dim rsTemp As New ADODB.Recordset
    
    If mint�༭ģʽ = 1 Then
        Call frm���ñ���_�ֵ���ϸ.ShowME
        txtͳ��֧��(intסԺ).Text = Format(gComInfo_üɽ.ͳ��֧��, strFormat_���)
        txtͳ���Ը�(intסԺ).Text = Format(gComInfo_üɽ.ͳ���Ը�, strFormat_���)
    Else
        '��ȡ�ֵ���������
        gstrSQL = " Select B.����,A.����,A.����,A.����ͳ���� ����ͳ��,A.ͳ�ﱨ����� ͳ�ﱨ�� " & _
                  " From ���ս������ A,(Select * From ���շ��õ� Where ����=" & TYPE_�Ĵ�üɽ & " And ����=" & Val(txt����(intסԺ).Tag) & " And ����<>0) B" & _
                  " Where A.����=B.���� And A.����ID=" & mlng��¼ID
        Call OpenRecordset(rsTemp, Me.Caption)
        Call frm���ñ���_�ֵ���ϸ.ShowME(True, rsTemp)
    End If
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    If gComInfo_üɽ.�����ܶ� = 0 Then
        MsgBox "δ�����κ����ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If mblnסԺ Then
        If Not SaveData(cbo��ҽ(intסԺ).ListIndex = 0) Then Exit Sub
    Else
        If Not SaveData(cbo��ҽ(int����).ListIndex = 0) Then Exit Sub
    End If
    
    '��ӡƱ��
    Call zl9Report.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1605_" & IIf(mblnסԺ, 2, 1), Me, "����=" & 25, "��¼ID=" & mlng��¼ID, 2)
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd�������_Click()
    Dim cur��� As Currency, curͳ�� As Currency, cur����ͳ�� As Currency
    Dim lngRow As Long, intIndex As Integer, intCol As Integer 'intCol��ʾ��ǰ���������
    Dim msfObj As BillEdit, sin���� As Single, intBound As Integer
    
    '��ʼ��
    Call Init_����_����
    Call Init_�ṹ��_����
    intIndex = IIf(mblnסԺ, intסԺ, int����)
    gComInfo_üɽ.����ID = Val(txt����(intIndex).Tag)
    gComInfo_üɽ.��Ⱥ = txt״̬(intIndex).Text
    If Not mblnסԺ Then
        gComInfo_üɽ.����ID = Val(txt����(intIndex).Tag)
        gComInfo_üɽ.�������� = txt����(intIndex).Text
    End If
    gComInfo_üɽ.���� = Val(txt����(intIndex).Tag)
    If mblnסԺ Then
        gComInfo_üɽ.�ʻ���� = Val(Split(txt���(intIndex).Text, "/")(0))
    Else
        gComInfo_üɽ.�ʻ���� = Val(txt���(intIndex).Text)
    End If
    If gComInfo_üɽ.����ID = 0 Then
        MsgBox "��ѡ��ҽ�����ˣ������������Ļ��ܽ��󣬲ſ�����Ԥ���㣡", vbInformation, gstrSysName
        txt����(intIndex).SetFocus
        Exit Sub
    End If
    Call ʵ�ʱ�������
    
    '���������ܼ�¼��
    cur��� = 0: curͳ�� = 0: cur����ͳ�� = 0
    Set msfObj = Bill(intIndex)
    For lngRow = 1 To msfObj.Rows - 2
        cur��� = cur��� + Val(msfObj.TextMatrix(lngRow, 1))
        rs����_����.MoveFirst
        rs����_����.Find "��������='" & msfObj.TextMatrix(lngRow, 0) & "'"
        If mblnסԺ Then rs����_����!���� = Val(msfObj.TextMatrix(lngRow, 2))
        rs����_����!�����ܶ� = Val(msfObj.TextMatrix(lngRow, 1))
        If Nvl(rs����_����!ͳ��ȶ�, 0) <> 0 Then
            rs����_����!�����ܶ� = Val(msfObj.TextMatrix(lngRow, 1))
            curͳ�� = curͳ�� + Val(msfObj.TextMatrix(lngRow, 1))
        End If
        rs����_����.Update
    Next
    gComInfo_üɽ.�����ܶ� = cur���
    gComInfo_üɽ.ȫ�Ը� = cur��� - curͳ��
    
    '�������ͳ����
    With rs����_����
        .MoveFirst
        Do While Not .EOF
            If !�������� = "��������" Then
                '����û������˱��������û������Ϊ׼
                sin���� = 100
                If InStr(1, gstrʵ�ʱ�������_����, "|" & !�������� & ";") <> 0 Then
                    intBound = UBound(Split(Mid(gstrʵ�ʱ�������_����, 2), "|"))
                    For lngRow = 0 To intBound
                        If Split(Split(Mid(gstrʵ�ʱ�������_����, 2), "|")(lngRow), ";")(0) = !�������� Then
                            sin���� = Val(Split(Split(Mid(gstrʵ�ʱ�������_����, 2), "|")(lngRow), ";")(1))
                            Exit For
                        End If
                    Next
                End If
                !�����ܶ� = !�����ܶ� * sin���� / 100
            Else
                If !��׼���� = 0 And !��׼���� = 0 Then
                    !�����ܶ� = !�����ܶ� * Nvl(!ͳ��ȶ�, 0) / 100
                Else
                    If !���� > !��׼���� Then
                        '���סԺ�ճ�����׼��������ô������ ��׼����*��׼���� +  (����-��׼����)*ͳ��ȶ�
                        '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                        !�����ܶ� = !��׼���� * !��׼���� + _
                            (!���� - IIf(!��׼���� = 0 Or !��׼���� = 0, 0, !��׼����)) * !ͳ��ȶ�
                    Else
                        If !��׼���� = 0 Or !��׼���� = 0 Then
                            '���סԺ�յ�����׼��������ô�������� ����*��׼���� ���� ����*ͳ��ȶ�
                            '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                            !�����ܶ� = !�����ܶ� * !ͳ��ȶ� / 100
                        Else
                            !�����ܶ� = !���� * !��׼����
                        End If
                    End If
                End If
            End If
            cur����ͳ�� = cur����ͳ�� + !�����ܶ�
            .Update
            .MoveNext
        Loop
    End With
    gComInfo_üɽ.�����Ը� = curͳ�� - cur����ͳ��
    gComInfo_üɽ.����ͳ�� = cur����ͳ��
    
    Call Calc_ʵ�ʱ���_����
    
    '����ͳ�Ｐ�ʻ�֧��
    If Not mblnסԺ Then
        Call Calc_���ﱨ������_����(cbo��ҽ(int����).ListIndex = 0)
    Else
        Call Calc_סԺ��������_����(cbo��ҽ(intסԺ).ListIndex = 0)
    End If
    
    Call ��ʾ������
    
    cmdȷ��.Enabled = True
    If mblnסԺ Then cmd�ֵ���ϸ.Enabled = True
End Sub

Private Sub ��ʾ������()
    Dim intCol As Integer, intIndex As Integer, lngRow As Long
    Dim cur��� As Currency
    Dim msfObj As BillEdit
    
    '��������Ŀɱ������д�ؽ��棬������ԱУ��
    intCol = IIf(mblnסԺ, 3, 2)
    intIndex = IIf(mblnסԺ, intסԺ, int����)
    cur��� = 0
    Set msfObj = Bill(intIndex)
    
    For lngRow = 1 To msfObj.Rows - 2
        rs����_����.MoveFirst
        rs����_����.Find "��������='" & msfObj.TextMatrix(lngRow, 0) & "'"
        msfObj.TextMatrix(lngRow, intCol) = Format(Nvl(rs����_����!�����ܶ�, 0), strFormat_���)
        cur��� = cur��� + Nvl(rs����_����!�����ܶ�, 0)
    Next
    msfObj.TextMatrix(msfObj.Rows - 1, intCol) = Format(cur���, strFormat_���)
    
    '��������д�ؽ��棬������ԱУ��
    txt�����ܶ�(intIndex).Text = Format(gComInfo_üɽ.�����ܶ�, strFormat_���)
    txt����ͳ��(intIndex).Text = Format(gComInfo_üɽ.����ͳ��, strFormat_���)
    txtȫ�Ը�(intIndex).Text = Format(gComInfo_üɽ.ȫ�Ը�, strFormat_���)
    txt�����Ը�(intIndex).Text = Format(gComInfo_üɽ.�����Ը�, strFormat_���)
    txtͳ��֧��(intIndex).Text = Format(gComInfo_üɽ.ͳ��֧��, strFormat_���)
    If Not mblnסԺ Then
        txt�ʻ�֧��(intIndex).Text = Format(gComInfo_üɽ.�ʻ�֧��, strFormat_���)
    Else
        txtͳ���Ը�(intIndex).Text = Format(gComInfo_üɽ.ͳ���Ը�, strFormat_���)
    End If
End Sub

Private Sub txt����_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt����(Index)
End Sub

Private Sub txt����_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txt����(Index).Text = ""
        txt����(Index).Tag = ""
    End If
    
    If mint�༭ģʽ = 1 Then cmdȷ��.Enabled = False
End Sub

Private Sub cmd����_Click(Index As Integer)
    gstrSQL = "Select A.����ID as ID,A.����,A.ҽ����,'******' ����,B.����,B.�Ա�,B.��������,B.���֤��,A.���� ����,C.���� ��������  " & _
             "  ,A.��Ա���,A.��λ����,A.����ID,D.���� as ����,E.��� ��Ⱥ,E.���� ״̬,A.����֤��,A.�ʻ���� " & _
             "  From �����ʻ� A,������Ϣ B,��������Ŀ¼ C,���ղ��� D,������Ⱥ E " & _
             "   where A.����ID=B.����ID And Nvl(A.�Ҷȼ�,0)<>9 And A.����=" & TYPE_�Ĵ�üɽ & _
             "   and A.����=C.���� and A.����=C.��� And A.��ְ=E.��� and A.����=E.���� and A.����ID=D.ID(+)"
    Call Get�ʻ����(Index)
    
    If mint�༭ģʽ = 1 Then cmdȷ��.Enabled = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    Call InitFace
    
    gstrSQL = " Select * From ����֧������ " & _
              " Where ����=" & TYPE_�Ĵ�üɽ & _
              " Order by ����"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    Call WriteBill(rsTemp, Bill(int����))
    Call WriteBill(rsTemp, Bill(intסԺ))
    Bill(int����).AllowAddRow = False
    Bill(intסԺ).AllowAddRow = False
    
    fra(int����).Visible = (Not mblnסԺ)
    fra(intסԺ).Visible = (mblnסԺ)
    If mblnסԺ Then
        Call ��ʾʵ�ʱ�������
        fra(intסԺ).ZOrder
    End If
    
    If mint�༭ģʽ = 2 Then
        '��ȡ����
        If Not ReadData Then
            MsgBox "δ�ҵ��������ݣ�", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
        '���ø��ؼ���״̬
        Dim objCon As Object
        For Each objCon In Me.Controls
            If InStr(1, "TEXTBOX;COMBOBOX", UCase(TypeName(objCon))) <> 0 Then
                objCon.Enabled = False
            ElseIf InStr(1, "CMD����,CMD����", UCase(objCon.Name)) <> 0 Then
                objCon.Enabled = False
            End If
        Next
        Bill(IIf(mblnסԺ, intסԺ, int����)).Active = False
        cmd�ֵ���ϸ.Enabled = mblnסԺ
        cmd�������.Enabled = False
        cmdȷ��.Enabled = False
        cmdȡ��.Caption = "ȷ��(&O)"
        cmdȡ��.Top = cmdȷ��.Top
    End If
End Sub

Private Function ReadData() As Boolean
    Dim intIndex As Integer
    Dim intCol As Integer, lngRow As Long
    Dim cur���ý�� As Currency, cur�����ܶ� As Currency
    Dim msfObj As BillEdit
    Dim rsTemp As New ADODB.Recordset
    '�ȶ�ȡҽ�����˵�����
    intIndex = IIf(mblnסԺ, intסԺ, int����)
    gstrSQL = " Select ҽ���� From �����ʻ� Where ����=" & TYPE_�Ĵ�üɽ & _
              " And ����ID = (" & _
              "        Select ����ID From ���ս����¼ Where ��¼ID=" & mlng��¼ID & " And ����=" & IIf(mblnסԺ, "2", "1") & " And ����=" & TYPE_�Ĵ�üɽ & ")"
    Call OpenRecordset(rsTemp, Me.Caption)
    If rsTemp.EOF Then Exit Function
    txt����(intIndex) = rsTemp!ҽ����
    Call txt����_KeyPress(intIndex, vbKeyReturn)
    
    '��������д�ؽ��棬������ԱУ��
    '��ȡ�����¼��Ϣ
    gstrSQL = " Select A.*,B.���� �������� From ���ս����¼ A,(Select ID ����ID,���� From ���ղ��� Where ����=" & TYPE_�Ĵ�üɽ & ") B" & _
              " Where A.��¼ID=" & mlng��¼ID & " And A.����=" & IIf(mblnסԺ, "2", "1") & _
              " And A.����=" & TYPE_�Ĵ�üɽ & " And A.����ID=B.����ID(+)"
    Call OpenRecordset(rsTemp, Me.Caption)
    If Not mblnסԺ Then
        txt����(intIndex).Text = Nvl(rsTemp!��������, "")
        txt����(intIndex).Tag = Nvl(rsTemp!����ID, 0)
    Else
        txt����(intסԺ).Text = Nvl(rsTemp!����, 0)
        txtסԺ����(intסԺ).Text = Nvl(rsTemp!סԺ����, 1)
    End If
    txt�����ܶ�(intIndex).Text = Format(rsTemp!�������ý��, strFormat_���)
    txt����ͳ��(intIndex).Text = Format(rsTemp!����ͳ����, strFormat_���)
    txtȫ�Ը�(intIndex).Text = Format(rsTemp!ȫ�Ը����, strFormat_���)
    txt�����Ը�(intIndex).Text = Format(rsTemp!�����Ը����, strFormat_���)
    txtͳ��֧��(intIndex).Text = Format(rsTemp!ͳ�ﱨ�����, strFormat_���)
    If Not mblnסԺ Then
        txt�ʻ�֧��(intIndex).Text = Format(rsTemp!�����ʻ�֧��, strFormat_���)
    Else
        txtͳ���Ը�(intIndex).Text = Format(Nvl(rsTemp!����ͳ����, 0) - Nvl(rsTemp!ͳ�ﱨ�����, 0), strFormat_���)
    End If
    
    '��ȡ�����������
    gstrSQL = "Select ����,��������,�����ܶ�,�����ܶ� From ���ձ�����¼ Where ��¼ID=" & mlng��¼ID & " Order by �������"
    Call OpenRecordset(rsTemp, Me.Caption)
    cbo��ҽ(intIndex).ListIndex = rsTemp!���� - 1
    intCol = IIf(mblnסԺ, 3, 2)
    cur���ý�� = 0: cur�����ܶ� = 0
    Set msfObj = Bill(intIndex)
    msfObj.Rows = 2 + rsTemp.RecordCount
    Do While Not rsTemp.EOF
        msfObj.TextMatrix(rsTemp.AbsolutePosition, 0) = rsTemp!��������
        msfObj.TextMatrix(rsTemp.AbsolutePosition, 1) = Format(rsTemp!�����ܶ�, strFormat_���)
        msfObj.TextMatrix(rsTemp.AbsolutePosition, intCol) = Format(rsTemp!�����ܶ�, strFormat_���)
        cur���ý�� = cur���ý�� + Nvl(rsTemp!�����ܶ�, 0)
        cur�����ܶ� = cur�����ܶ� + Nvl(rsTemp!�����ܶ�, 0)
        rsTemp.MoveNext
    Loop
    msfObj.TextMatrix(msfObj.Rows - 1, 0) = "�ϼ�"
    msfObj.TextMatrix(msfObj.Rows - 1, 1) = Format(cur���ý��, strFormat_���)
    msfObj.TextMatrix(msfObj.Rows - 1, msfObj.Cols - 1) = Format(cur�����ܶ�, strFormat_���)
    
    ReadData = True
End Function

Private Sub InitFace()
    With cbo��ҽ(int����)
        .Clear
        .AddItem "��Ժ"
        .AddItem "��Ժ"
        .ListIndex = 0
    End With
    With cbo��ҽ(intסԺ)
        .Clear
        .AddItem "��Ժ"
        .AddItem "��Ժ"
        .ListIndex = 0
    End With

    Call InitBill(Bill(int����))
    Call InitBill(Bill(intסԺ))
End Sub

Private Sub InitBill(ByVal msfObj As BillEdit)
    With msfObj
        .ClearBill
        .Active = True
        .Rows = 2
        .Cols = IIf(mblnסԺ, 4, 3)
        .msfObj.FixedCols = 1
        
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "�����ܶ�"
        If mblnסԺ Then
            .TextMatrix(0, 2) = "����"
            .TextMatrix(0, 3) = "������"
        Else
            .TextMatrix(0, 2) = "������"
        End If
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        If mblnסԺ Then .ColWidth(3) = 1200: .ColWidth(0) = 1200: .ColWidth(2) = 800
        
        .ColData(0) = 5
        .ColData(1) = 4
        If mblnסԺ Then
            .ColData(2) = 4
            .ColData(3) = 5
        Else
            .ColData(2) = 5
        End If
        .PrimaryCol = 1
        .LocateCol = 1
    End With
End Sub

Private Sub WriteBill(ByVal rsTemp As ADODB.Recordset, ByVal msfObj As BillEdit)
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If !���� = "��������" And msfObj.Index = 0 Then
                '���ﲻ��ʾ�ô���
                .MoveNext
                If .EOF Then Exit Do
            End If
            msfObj.TextMatrix(.AbsolutePosition, 0) = !����
            msfObj.RowData(.AbsolutePosition) = 0
            msfObj.Rows = msfObj.Rows + 1
            .MoveNext
        Loop
        
        msfObj.TextMatrix(msfObj.Rows - 1, 0) = "�ϼ�"
    End With
End Sub

Private Sub txt����_GotFocus(Index As Integer)
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt����(Index)
End Sub

Private Sub txt����_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strCode As String
    Dim str���� As String
    
    If Len(txt����(Index).Text) = txt����(Index).MaxLength Or KeyAscii = vbKeyReturn Then
        strCode = UCase(Replace(Trim(txt����(Index).Text), "'", ""))
        If Len(strCode) = 0 Then Exit Sub
        
        If IsNumeric(Mid(strCode, 1, Len(strCode) - 1)) Then 'ˢ��
            str���� = " and A.����='" & strCode & "'"
        ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '����ID
            str���� = " and A.����ID=" & Mid(strCode, 2)
        ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then 'סԺ��(��ס(��)Ժ�Ĳ���)
            str���� = " and B.סԺ��=" & Mid(strCode, 2)
        ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '�����(�������ﲡ��)
            str���� = " and B.�����=" & Mid(strCode, 2)
        Else '��������
            str���� = " and A.����='" & strCode & "'"
        End If
    
        gstrSQL = "Select A.����ID as ID,A.����,A.ҽ����,'******' ����,B.����,B.�Ա�,B.��������,B.���֤��,A.���� ����,C.���� ��������  " & _
                 "  ,A.��Ա���,A.��λ����,A.����ID,D.���� as ����,E.��� ��Ⱥ,E.���� ״̬,A.����֤��,A.�ʻ���� " & _
                 "  From �����ʻ� A,������Ϣ B,��������Ŀ¼ C,���ղ��� D,������Ⱥ E " & _
                 "   where A.����ID=B.����ID And Nvl(A.�Ҷȼ�,0)<>9 And A.����=" & TYPE_�Ĵ�üɽ & _
                 "   and A.����=C.���� and A.����=C.��� And A.��ְ=E.��� and A.����=E.���� and A.����ID=D.ID(+)" & str����
        Call Get�ʻ����(Index)
    End If
    
    If mint�༭ģʽ = 1 Then cmdȷ��.Enabled = False
End Sub

Private Sub txt����_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Private Sub Get�ʻ����(Index As Integer)
'���Ѿ����ڵļ�¼�ж����ʻ���Ϣ
    Dim rsTemp As New ADODB.Recordset
    Dim rs�ʻ� As New ADODB.Recordset
    Dim lngIndex As Long
    
    Set rs�ʻ� = frmPubSel.ShowSelect(Me, gstrSQL, 0, "�����ʻ�", , txt����(Index).Text, "", False, True)
    If rs�ʻ�.State <> 1 Then Exit Sub
    If Not rs�ʻ� Is Nothing Then
        '��鲡��״̬
        gstrSQL = "select nvl(��ǰ״̬,0) as ״̬,�Ҷȼ�,��ע from �����ʻ� where ����=25 and ҽ����='" & Trim(txt����(Index).Text) & "'"
        Call OpenRecordset(rsTemp, Me.Caption)
        
        If rsTemp.RecordCount > 0 Then
            Select Case Nvl(rsTemp!�Ҷȼ�, 0)
            Case 1
                MsgBox "��ҽ�����Ѿ�����������ʹ�ã�" & IIf(Nvl(rsTemp!��ע) <> "", "��" & rsTemp!��ע & "��", ""), vbInformation, gstrSysName
                Exit Sub
            Case 9
                MsgBox "��ҽ�����Ѿ�����������ʹ�ã�", vbInformation, gstrSysName
                Exit Sub
            End Select
        End If
        
        '������ݿ��е������Ƿ���ȷ
        If Not ����ʻ���Ϣ_����(txt����(Index).Text) Then Exit Sub
        
        txt����(Index).Text = rs�ʻ�("����")
        txt����(Index).Tag = rs�ʻ�("ID")
        txt����(Index).Text = IIf(IsNull(rs�ʻ�("����")), "", rs�ʻ�("����"))
        txt״̬(Index).Text = rs�ʻ�("״̬")
        txt״̬(Index).Tag = rs�ʻ�("��Ⱥ")
        If mblnסԺ = False Then
            txt����(Index).Text = Nvl(rs�ʻ�("����"))
            txt����(Index).Tag = Nvl(rs�ʻ�("����ID"), 0)
        End If
        txt����(Index).Text = rs�ʻ�("��������")
        txt����(Index).Tag = rs�ʻ�("����")
        txt���(Index).Text = Format(rs�ʻ�!�ʻ����, strFormat_���)
        lblNote.Caption = Nvl(rs�ʻ�!����֤��)
        If Index = intסԺ Then
            '�ٶ����ʻ������Ϣ
            gstrSQL = "select * from �ʻ������Ϣ where ����=" & TYPE_�Ĵ�üɽ & _
                " and ����ID=" & rs�ʻ�("ID") & " and ���=" & Format(zlDatabase.Currentdate, "yyyy")
            Call OpenRecordset(rsTemp, Me.Caption)
            
            If rsTemp.EOF = False Then
                '�����ʻ����
                If cbo��ҽ(Index).ListIndex = 0 Then
                    txtסԺ����(Index).Text = Nvl(rsTemp("סԺ�����ۼ�"), 0) + IIf(mint�༭ģʽ = 1, 1, 0)
                Else
                    txtסԺ����(Index).Text = Nvl(rsTemp("��ԺסԺ����"), 0) + IIf(mint�༭ģʽ = 1, 1, 0)
                End If
                txt���(Index).Text = txt���(Index).Text & "/" & Format(rsTemp!����ͳ���ۼ�, strFormat_���)
            Else
                txtסԺ����(Index).Text = "1"
            End If
            
            Call cbo��ҽ_Click(Index)
        End If
    End If
End Sub

Private Function SaveData(Optional ByVal bln��Ժ As Boolean = True) As Boolean
    '������������סԺ��������
    Dim lng����ID As Long, strҽ���� As String, intIndex As Integer
    Dim blnExecute As Boolean
    
    'ȡ������ˮ�ż�����ҽ������
    intIndex = IIf(mblnסԺ, intסԺ, int����)
    lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    strҽ���� = txt����(intIndex)
    
    '����
    gcnOracle.BeginTrans
    If mblnסԺ Then
        blnExecute = סԺ����(lng����ID, bln��Ժ)
    Else
        blnExecute = �������_üɽ(lng����ID, Val(txt�ʻ�֧��(int����)), strҽ����, bln��Ժ)
    End If
    If Not blnExecute Then
        gcnOracle.RollbackTrans
    Else
        gcnOracle.CommitTrans
        mlng��¼ID = lng����ID
    End If
    
    SaveData = blnExecute
End Function

Private Function סԺ����(lng����ID As Long, Optional ByVal bln��Ժ���� As Boolean = True) As Boolean
    Dim int���� As Integer
    Dim lng��� As Long, int��Ժ As Integer, int��Ժ As Integer
    Dim cur�ʻ���� As Currency, curͳ���ۼ� As Currency
    Dim rsTemp As New ADODB.Recordset
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    On Error GoTo ErrHand
    int���� = IIf(bln��Ժ����, 1, 2)
    
    '��������Ϣ���浽���ս����¼��
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�Ĵ�üɽ & "," & gComInfo_üɽ.����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & ",0,0,0,0," & Val(txtסԺ����(intסԺ).Text) & "," & gComInfo_üɽ.���� & ",0," & gComInfo_üɽ.�������� & "," & _
        gComInfo_üɽ.�����ܶ� & "," & gComInfo_üɽ.ȫ�Ը� & "," & gComInfo_üɽ.�����Ը� & "," & gComInfo_üɽ.����ͳ�� & "," & gComInfo_üɽ.ͳ��֧�� & ",0," & _
        0 & "," & 0 & ",null,null,null,null,null,'" & gstrUserName & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ��������")
    
    '���ո�����ı�����ϸ
    With rs����_����
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!�����ܶ�, 0) <> 0 Then
                gstrSQL = "ZL_���ձ�����¼_INSERT(" & int���� & "," & lng����ID & "," & _
                "'" & !������� & "','" & !�������� & "'," & !ͳ��ȶ� & "," & _
                "" & !��׼���� & "," & !��׼���� & "," & !�����ܶ� & "," & !�����ܶ� & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "������౨������")
            End If
            .MoveNext
        Loop
    End With
    
    '���շֵ�������ϸ
    With rs�ֵ�֧��_����
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If !����ͳ�� <> 0 Then
                gstrSQL = "ZL_���ս������_INSERT(" & lng����ID & "," & !���� & "," & !����ͳ�� & "," & !ͳ�ﱨ�� & "," & !���� & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "����ֵ�������ϸ")
            End If
            .MoveNext
        Loop
    End With
    
    '����סԺ����
    lng��� = Format(zlDatabase.Currentdate, "yyyy")
    cur�ʻ���� = 0: curͳ���ۼ� = gComInfo_üɽ.����ͳ��
    gstrSQL = " Select Nvl(�ʻ������ۼ�,0) �ʻ����,nvl(����ͳ���ۼ�,0) ͳ���ۼ�" & _
              " ,Nvl(סԺ�����ۼ�,0) ��Ժ,Nvl(��ԺסԺ����,0) ��Ժ" & _
              " From �ʻ������Ϣ" & _
              " Where ���=" & lng��� & " And ����ID=" & gComInfo_üɽ.����ID
    Call OpenRecordset(rsTemp, Me.Caption)
    If Not rsTemp.EOF Then
        cur�ʻ���� = rsTemp!�ʻ����: curͳ���ۼ� = curͳ���ۼ� + rsTemp!ͳ���ۼ�
        If bln��Ժ���� Then
            int��Ժ = Val(txtסԺ����(intסԺ).Text)
            int��Ժ = rsTemp!��Ժ
        Else
            int��Ժ = rsTemp!��Ժ
            int��Ժ = Val(txtסԺ����(intסԺ).Text)
        End If
    Else
        If bln��Ժ���� Then
            int��Ժ = 1
            int��Ժ = 0
        Else
            int��Ժ = 0
            int��Ժ = 1
        End If
    End If
    
    gstrSQL = "zl_�ʻ������Ϣ_Insert(" & gComInfo_üɽ.����ID & ",25," & lng��� & _
              "," & cur�ʻ���� & ",0," & curͳ���ۼ� & ",0," & int��Ժ & "," & int��Ժ & "," & Val(txt����(intסԺ).Text) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ����")
    
    סԺ���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ��ʾʵ�ʱ�������()
    Dim lngRow As Long
    Dim sin���� As Single
    Dim rsTemp As New ADODB.Recordset
    
    '�����סԺԤ���㣬������ʵ�ʱ��������Ĵ��಻����100%�����䱨��������ʾ��������
    If mblnסԺ Then
        gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=1 And ���>=10"
        Call OpenRecordset(rsTemp, Me.Caption)
        With Bill(intסԺ)
            For lngRow = 1 To .Rows - 2
                sin���� = 100
                If rsTemp.RecordCount <> 0 Then
                    rsTemp.MoveFirst
                    rsTemp.Find "������='" & .TextMatrix(lngRow, 0) & "'"
                    If Not rsTemp.EOF Then sin���� = Nvl(rsTemp!����ֵ, 100)
                End If
                If sin���� <> 100 Or .TextMatrix(lngRow, 0) = "��������" Then
                    '����������100%���������
                    .TextMatrix(lngRow, 2) = Format(sin����, strFormat_���)
                    .RowData(lngRow) = 1
                End If
            Next
        End With
    End If
End Sub

Private Sub ʵ�ʱ�������()
    Dim lngRow As Long
    gstrʵ�ʱ�������_���� = ""
    If mblnסԺ = False Then Exit Sub
    
    With Bill(intסԺ)
        For lngRow = 1 To .Rows - 2
            If .RowData(lngRow) = 1 Then
                gstrʵ�ʱ�������_���� = gstrʵ�ʱ�������_���� & "|" & .TextMatrix(lngRow, 0) & ";" & Val(.TextMatrix(lngRow, 2))
            End If
        Next
    End With
End Sub
