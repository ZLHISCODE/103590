VERSION 5.00
Begin VB.Form frmMergePatient 
   Caption         =   "���˺ϲ�"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   Icon            =   "frmMergePatient.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdTurn 
      Caption         =   "����(&T)"
      Height          =   350
      Left            =   6675
      TabIndex        =   1
      Top             =   270
      Width           =   1100
   End
   Begin VB.CommandButton cmdPati 
      Height          =   330
      Left            =   4950
      Picture         =   "frmMergePatient.frx":06EA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "ѡ����(F2)"
      Top             =   15
      Width           =   420
   End
   Begin VB.Frame fra 
      Caption         =   "Ҫ�����Ĳ�����Ϣ      "
      Height          =   4725
      Index           =   1
      Left            =   3330
      TabIndex        =   6
      Top             =   90
      Width           =   3135
      Begin VB.TextBox txt״̬ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   270
         Width           =   2025
      End
      Begin VB.TextBox txtסԺ�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   525
         Width           =   2025
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   780
         Width           =   2025
      End
      Begin VB.TextBox txt�Ա� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1035
         Width           =   2025
      End
      Begin VB.TextBox txt�������� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1290
         Width           =   2025
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1545
         Width           =   2025
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2025
      End
      Begin VB.TextBox txtѧ�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2025
      End
      Begin VB.TextBox txt����״�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2325
         Width           =   2025
      End
      Begin VB.TextBox txtְҵ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2580
         Width           =   2025
      End
      Begin VB.TextBox txt��� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2835
         Width           =   2025
      End
      Begin VB.TextBox txt���֤�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3090
         Width           =   2025
      End
      Begin VB.TextBox txt�����ص� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   3345
         Width           =   2025
      End
      Begin VB.TextBox txt��ͥ��ַ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2025
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3855
         Width           =   2025
      End
      Begin VB.TextBox txt��λ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   4110
         Width           =   2025
      End
      Begin VB.TextBox txtסԺ���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   4380
         Width           =   2025
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   74
         Top             =   780
         Width           =   450
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   73
         Top             =   1035
         Width           =   450
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   72
         Top             =   1290
         Width           =   810
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   71
         Top             =   1545
         Width           =   450
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   70
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label lblѧ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѧ��:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   69
         Top             =   2055
         Width           =   450
      End
      Begin VB.Label lbl����״�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   68
         Top             =   2325
         Width           =   810
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   67
         Top             =   2835
         Width           =   450
      End
      Begin VB.Label lblְҵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   66
         Top             =   2580
         Width           =   450
      End
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   65
         Top             =   3090
         Width           =   810
      End
      Begin VB.Label lbl�����ص� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ص�:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   64
         Top             =   3345
         Width           =   810
      End
      Begin VB.Label lbl��ͥ��ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   63
         Top             =   3600
         Width           =   810
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   62
         Top             =   3855
         Width           =   450
      End
      Begin VB.Label lbl��λ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   61
         Top             =   4110
         Width           =   450
      End
      Begin VB.Label lblסԺ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   60
         Top             =   4380
         Width           =   810
      End
      Begin VB.Label lbl״̬ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "״̬:"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   59
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   315
         TabIndex        =   58
         Top             =   525
         Width           =   630
      End
   End
   Begin VB.Frame fra 
      Caption         =   "���ϲ��Ĳ�����Ϣ"
      Height          =   4725
      Index           =   0
      Left            =   135
      TabIndex        =   5
      Top             =   90
      Width           =   3135
      Begin VB.TextBox txtסԺ���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   4380
         Width           =   2025
      End
      Begin VB.TextBox txt��λ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   4110
         Width           =   2025
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3855
         Width           =   2025
      End
      Begin VB.TextBox txt��ͥ��ַ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2025
      End
      Begin VB.TextBox txt�����ص� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3345
         Width           =   2025
      End
      Begin VB.TextBox txt���֤�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   3090
         Width           =   2025
      End
      Begin VB.TextBox txt��� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2835
         Width           =   2025
      End
      Begin VB.TextBox txtְҵ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2580
         Width           =   2025
      End
      Begin VB.TextBox txt����״�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2325
         Width           =   2025
      End
      Begin VB.TextBox txtѧ�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2025
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2025
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1545
         Width           =   2025
      End
      Begin VB.TextBox txt�������� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1290
         Width           =   2025
      End
      Begin VB.TextBox txt�Ա� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1035
         Width           =   2025
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   780
         Width           =   2025
      End
      Begin VB.TextBox txtסԺ�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   525
         Width           =   2025
      End
      Begin VB.TextBox txt״̬ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   270
         Width           =   2025
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   330
         TabIndex        =   23
         Top             =   525
         Width           =   630
      End
      Begin VB.Label lbl״̬ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "״̬:"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   22
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lblסԺ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   21
         Top             =   4380
         Width           =   810
      End
      Begin VB.Label lbl��λ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   20
         Top             =   4110
         Width           =   450
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   19
         Top             =   3855
         Width           =   450
      End
      Begin VB.Label lbl��ͥ��ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   18
         Top             =   3600
         Width           =   810
      End
      Begin VB.Label lbl�����ص� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ص�:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   17
         Top             =   3345
         Width           =   810
      End
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   16
         Top             =   3090
         Width           =   810
      End
      Begin VB.Label lblְҵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   15
         Top             =   2580
         Width           =   450
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   14
         Top             =   2835
         Width           =   450
      End
      Begin VB.Label lbl����״�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   2325
         Width           =   810
      End
      Begin VB.Label lblѧ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѧ��:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   12
         Top             =   2055
         Width           =   450
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   11
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   10
         Top             =   1545
         Width           =   450
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   1290
         Width           =   810
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   8
         Top             =   1035
         Width           =   450
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   7
         Top             =   780
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdMerge 
      Caption         =   "�ϲ�(&M)"
      Height          =   350
      Left            =   6675
      TabIndex        =   2
      Top             =   795
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   6675
      TabIndex        =   3
      Top             =   3570
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6675
      TabIndex        =   4
      Top             =   4095
      Width           =   1100
   End
End
Attribute VB_Name = "frmMergePatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mlng����ID As Long
Private mlng���ϲ�����ID As Long
Private mblnOk As Boolean
Private mcnOracle As ADODB.Connection
Private mobjDataBase As clsDataBase
Private mobjOneDataObject As clsOneCardDataObject

Public Function zlShowPatiMerge(ByVal cnOracle As ADODB.Connection, ByVal frmMain As Object, _
    ByVal lng���ϲ�����ID As Long, ByRef lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���˺ϲ�
    '���:   lng����ID-�ϲ��Ĳ���ID
    '           lng���ϲ���ID-���ϲ��Ĳ���ID
    '����:
    '           lng����ID-�ϲ���Ĳ���ID
    '����:strOutput-Ӧ������
    '����: �ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-08-21 15:43:37
    '˵��:
    '����:52913
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng���ϲ�����ID = lng���ϲ�����ID
    mlng����ID = lng����ID
    mblnOk = False
    If zlGetOneDataBase(cnOracle, mobjDataBase) = False Then Exit Function
    If zlGetOneCardDataObject(cnOracle, mobjOneDataObject) = False Then Exit Function
    
    Me.Show 1, frmMain
    lng����ID = mlng����ID
    zlShowPatiMerge = mblnOk
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    If gobjComLib Is Nothing Then Call zlInitCommLib
    If gobjComLib Is Nothing Then Exit Sub
    gobjComLib.ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdPati_Click()
    Dim lng����ID As Long
    If frmPatiSel.zlShowCard(mcnOracle, Me, "", lng����ID) = False Then Exit Sub
    If mobjOneDataObject.zlIsExistFeeInsurePatient(lng����ID) Then
        MsgBox "��ҽ�����˴���δ�����,���Ƚ�����ٺϲ���", vbExclamation, gstrSysName: Exit Sub
    End If
    Call ShowPatiInfo(lng����ID, 1)
End Sub


Private Sub cmdTurn_Click()
    Dim lngTmp As Long
    
    If Val(fra(1).Tag) = 0 Then
        If glngSys Like "8??" Then
            MsgBox "û������Ҫ�����Ŀͻ�,����ѡ��һ���ͻ���", vbInformation, gstrSysName
        Else
            MsgBox "û������Ҫ�����Ĳ���,����ѡ��һ�����ˣ�", vbInformation, gstrSysName
        End If
        cmdPati.SetFocus: Exit Sub
    End If
    If Val(fra(1).Tag) = Val(fra(0).Tag) Then
        If glngSys Like "8??" Then
            MsgBox "��ѡ����ͬһ���ͻ�,��ѡ�������ͻ���", vbInformation, gstrSysName
        Else
            MsgBox "��ѡ����ͬһ������,��ѡ���������ˣ�", vbInformation, gstrSysName
        End If
        cmdPati.SetFocus: Exit Sub
    End If
    
    lngTmp = fra(0).Tag
    Call ShowPatiInfo(CLng(fra(1).Tag), 0)
    Call ShowPatiInfo(lngTmp, 1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            cmdPati_Click
    End Select
End Sub

Private Sub Form_Load()
    fra(1).Tag = ""
    If Not ShowPatiInfo(mlng����ID, 0) Then Unload Me: Exit Sub
    If Not ShowPatiInfo(mlng���ϲ�����ID, 1) Then Unload Me: Exit Sub
End Sub

Private Sub ClearPatiInfo(X As Integer)
'���ܣ����һ��������Ϣ
'������x=�ؼ�����,0=Դ����,1=Ŀ�겡��
    txt����(X).Text = ""
    txt�Ա�(X).Text = ""
    txt��������(X).Text = ""
    txt����(X).Text = ""
    txt����(X).Text = ""
    txtѧ��(X).Text = ""
    txt���(X).Text = ""
    txtְҵ(X).Text = ""
    txt���֤��(X).Text = ""
    txt�����ص�(X).Text = ""
    txt��ͥ��ַ(X).Text = ""
    txt����״��(X).Text = ""
    txt״̬(X).Text = ""
    lblסԺ��(X).Caption = "סԺ��:"
    txtסԺ��(X).Text = ""
    txt����(X).Text = ""
    txt��λ(X).Text = ""
    txtסԺ����(X).Text = ""
    fra(X).Tag = ""
End Sub

Private Function ShowPatiInfo(lngID As Long, X As Integer) As Boolean
    '���ܣ���ʾһ��������Ϣ
    '������lngID=����ID,x=�ؼ�����,0=Դ����,1=Ŀ�겡��
    Dim cllData As Collection, cllTemp As Collection
    Dim strסԺ�� As String, str����� As String
    Dim lng��ҳID As Long
    
    On Error GoTo errH
    
    If zl_PatiSvr_GetPatiInfo(lngID, Nothing, cllData, 2) = False Then Exit Function
    If cllData.count = 0 Then
        MsgBox "δ���ָò��˵������Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    Set cllTemp = cllData(1)
 
    txt����(X).Text = cllTemp("_pati_name")
    txt�Ա�(X).Text = NVL(cllTemp("_pati_sex"), "")
    If cllTemp("_pati_birthdate") <> "" Then
         txt��������(X).Text = Format(CDate(cllTemp("_pati_birthdate")), "yyyy��MM��dd��")
    End If
    
    txt����(X).Text = cllTemp("_country_name")
    txt����(X).Text = cllTemp("_pati_nation")
    txtѧ��(X).Text = cllTemp("_pati_education")
    txt���(X).Text = cllTemp("_pati_identity")
    txtְҵ(X).Text = cllTemp("_ocpt_name")
    txt���֤��(X).Text = cllTemp("_pati_idcard")
    txt�����ص�(X).Text = cllTemp("_pati_birthplace")
    txt��ͥ��ַ(X).Text = cllTemp("_pat_home_addr")
    txt����״��(X).Text = cllTemp("_pati_marital_cstatus")
    
    str����� = NVL(cllTemp("_outpatient_num"), 0)
    strסԺ�� = NVL(cllTemp("_inpatient_num"), 0)
    lng��ҳID = Val(NVL(cllTemp("_pati_pageid")))
    
    lblסԺ��(X).Caption = "סԺ��:"
    txtסԺ��(X).Text = IIf(strסԺ�� = 0, "", strסԺ��)
    txtסԺ����(X).Text = lng��ҳID
    
    If zl_CisSvr_GetPatPageInfByRange(1, Nothing, lngID & ":" & lng��ҳID, , cllData) Then
        '��ȡ��ҳ��Ϣ
        If cllData.count <> 0 Then
            Set cllTemp = cllData(1)
            If NVL(cllTemp("_adtd_time")) = "" Then
                txt״̬(X).Text = "��Ժ"
            Else
                txt״̬(X).Text = "��Ժ"
            End If
            
            txt����(X).Text = NVL(cllTemp("_pati_dept_name"))
            If NVL(cllTemp("_pati_bed")) = "" Then
                 txt��λ(X).Text = NVL(cllTemp("_pati_bed"))
            Else
                txt��λ(X).Text = "��ͥ"
            End If
        End If
    End If
    
    fra(X).Tag = lngID
    ShowPatiInfo = True
    
    Exit Function
errH:
    If mobjDataBase.ErrCenter() = 1 Then Resume
    Call mobjDataBase.SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    Unload frmPatiSel
    If mcnOracle Is Nothing Then Set mcnOracle = Nothing
    If mobjDataBase Is Nothing Then Set mobjDataBase = Nothing
    If mobjOneDataObject Is Nothing Then Set mobjOneDataObject = Nothing
End Sub

Private Sub cmdMerge_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim rsPatiS As ADODB.Recordset
    Dim rsPatiO As ADODB.Recordset
    Dim strSQL As String, curDate As Date
    Dim I As Integer, J As Integer
    Dim str�ϲ�ԭ�� As String
    
    If Val(fra(1).Tag) = 0 Then
        MsgBox "û������Ҫ�����Ĳ���,����ѡ��һ�����ˣ�", vbInformation, gstrSysName
        cmdPati.SetFocus: Exit Sub
    End If
    If Val(fra(1).Tag) = Val(fra(0).Tag) Then
        MsgBox "��ѡ����ͬһ������,��ѡ���������ˣ�", vbInformation, gstrSysName
        cmdPati.SetFocus: Exit Sub
    End If
        
    Set rsPatiS = GetPatiInfo(CLng(fra(0).Tag))
    Set rsPatiO = GetPatiInfo(CLng(fra(1).Tag))
    
    'A��B��һ��������ԤԼ��Ժ
    If Not IsNull(rsPatiS!��ҳID) And NVL(rsPatiS!��ҳID, 0) = 0 Then
        MsgBox "����:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]������ԤԼ��Ժ�Ǽǣ�����ȡ���õǼǡ�", vbInformation, gstrSysName
    End If
    If Not IsNull(rsPatiO!��ҳID) And NVL(rsPatiO!��ҳID, 0) = 0 Then
        MsgBox "����:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]������ԤԼ��Ժ�Ǽǣ�����ȡ���õǼǡ�", vbInformation, gstrSysName
    End If
    
    'AB��ס��Ժ
    If Not IsNull(rsPatiS!��ҳID) And Not IsNull(rsPatiO!��ҳID) Then
        '1.��סԺ����Ժ,������(�Ⱥ�סԺ����Ϊ����Ժ-��Ժ,��Ժ-��Ժ����������Ժ-��Ժ,��Ժ-��Ժ)
        '��Ϊ�����˺ϲ���,���򲻶��⴦���Զ���Ժ������Ժ
        rsPatiS.MoveLast
        rsPatiO.MoveLast
        If rsPatiS!��Ժ���� <= rsPatiO!��Ժ���� Then
            If IsNull(rsPatiS!��Ժ����) Then
                MsgBox "����:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]���һ��סԺ����Ժ,����ǰδ��Ժ,����ִ�кϲ�������", vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            If IsNull(rsPatiO!��Ժ����) Then
                MsgBox "����:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]���һ��סԺ����Ժ,����ǰδ��Ժ,����ִ�кϲ�������", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '2.ʱ�佻����ʾ�Ƿ����
        curDate = mobjDataBase.Currentdate
        rsPatiS.MoveFirst
        For I = 1 To rsPatiS.RecordCount
            rsPatiO.MoveFirst
            For J = 1 To rsPatiO.RecordCount
                If Not (rsPatiO!��Ժ���� >= IIf(IsNull(rsPatiS!��Ժ����), curDate, rsPatiS!��Ժ����) Or _
                    IIf(IsNull(rsPatiO!��Ժ����), curDate, rsPatiO!��Ժ����) <= rsPatiS!��Ժ����) Then
                    MsgBox "���ֲ���:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]�� " & rsPatiS!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiS!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiS!��Ժ����), curDate, rsPatiS!��Ժ����), "yyyy-MM-dd") & vbCrLf & _
                        "�벡��:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]�ĵ� " & rsPatiO!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiO!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiO!��Ժ����), curDate, rsPatiO!��Ժ����), "yyyy-MM-dd") & _
                        vbCrLf & "���ཻ�棬���ܽ��кϲ���", _
                        vbInformation, gstrSysName
                        Exit Sub
                End If
                rsPatiO.MoveNext
            Next
            rsPatiS.MoveNext
        Next
    End If
    
    '�ϲ�ԭ��
    str�ϲ�ԭ�� = InputBox("�ϲ��������ܳ���,������!" & vbCrLf & vbCrLf & "������ϲ�ԭ��:" & vbCrLf & vbCrLf, gstrSysName, "")
    If ActualLen(str�ϲ�ԭ��) > 250 Then
        MsgBox "�ϲ�ԭ���ܶ���250���ַ�,�밴Ctrl+C�������������,����ִ��ʱ������:" & _
            vbCrLf & vbCrLf & str�ϲ�ԭ��, vbInformation, gstrSysName
        Exit Sub
    ElseIf Trim(str�ϲ�ԭ��) = "" Then
        MsgBox "��������ϲ�ԭ����ܽ��кϲ�!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    DoEvents
    On Error GoTo errH
    
    strSQL = "zl_������Ϣ_MERGE(" & Val(fra(0).Tag) & "," & Val(fra(1).Tag) & ",'" & str�ϲ�ԭ�� & "','" & UserInfo.���� & "')"
    Call mobjDataBase.ExecuteProcedure(strSQL, Me.Caption)
    
    
    
    On Error GoTo 0
    Screen.MousePointer = 0
    
    Dim cllFilter As Collection, cllData As Collection
    
    Set cllFilter = New Collection
    cllFilter.Add Array("����IDS", Val(fra(0).Tag) & "," & Val(fra(1).Tag))
    If zl_PatiSvr_GetPatiInfo(0, cllFilter, cllData) Then
        '�ϲ���Ӧֻʣһ������
        If cllData.count <> 0 Then
            mlng����ID = Val(NVL(cllData(1)("_pati_id")))
            MsgBox "���˺ϲ��ɹ�,�ϲ���Ĳ���IDΪ " & mlng����ID & "��", vbInformation, gstrSysName
        End If
    End If
    
    Call ClearPatiInfo(1)
    Call ShowPatiInfo(mlng����ID, 0)
    Unload frmPatiSel
    mblnOk = True
    cmdPati.SetFocus
    Exit Sub
errH:
    Screen.MousePointer = 0
    If mobjDataBase.ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call mobjDataBase.SaveErrLog
End Sub

Private Function GetPatiInfo(lng����ID As Long) As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim cllData As Collection, cllTemp As Collection, cllPati As Collection
    Dim I As Long
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        If .State = 1 Then .Close
        .fields.Append "��ҳID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "����ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "סԺ��", adLongVarChar, 18, adFldIsNullable
        .fields.Append "��Ժ����", adDate, , adFldIsNullable
        .fields.Append "��Ժ����", adDate, , adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    If Not zl_PatiSvr_GetPatiInfo(lng����ID, Nothing, cllPati) = False Then Exit Function
    If cllPati.count = 0 Then Exit Function
    If Not zl_CisSvr_GetPatPageInfByRange(0, Nothing, lng����ID, , cllData) Then Exit Function
    
    If cllData.count = 0 Then
        Set cllTemp = cllData(1)
        With rsTemp
            .AddNew
            !����ID = cllTemp("_pati_id")
            !��ҳID = Null
            !���� = cllTemp("_pati_name")
            !סԺ�� = CStr(NVL(cllTemp("_inpatient_num")))
            If cllTemp("_adta_time") = "" Then
                !��Ժ���� = Null
            Else
                !��Ժ���� = CDate(cllTemp("_adta_time"))
            End If
            If cllTemp("_adtd_time") = "" Then
                !��Ժ���� = Null
            Else
                !��Ժ���� = CDate(cllTemp("_adtd_time"))
            End If
            .Update
        End With
        If Not rsTemp.EOF Then Set GetPatiInfo = rsTemp
        Exit Function
    End If
    For I = 1 To cllData.count
        Set cllTemp = cllData(I)
        
        With rsTemp
            .AddNew
            !����ID = cllTemp("_pati_id")
            !��ҳID = cllTemp("_pati_pageid")
            !���� = cllTemp("_pati_name")
            !סԺ�� = CStr(NVL(cllTemp("_inpatient_num")))
            If cllTemp("_adta_time") = "" Then
                !��Ժ���� = Null
            Else
                !��Ժ���� = CDate(cllTemp("_adta_time"))
            End If
            If cllTemp("_adtd_time") = "" Then
                !��Ժ���� = Null
            Else
                !��Ժ���� = CDate(cllTemp("_adtd_time"))
            End If
            .Update
        End With
    Next
    If Not rsTemp.EOF Then Set GetPatiInfo = rsTemp
    Exit Function
errH:
    If mobjDataBase.ErrCenter() = 1 Then Resume
    Call mobjDataBase.SaveErrLog
End Function
