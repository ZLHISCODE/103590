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
Private mblnOK As Boolean
Public Function zlShowPatiMerge(ByVal frmMain As Object, _
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
    mblnOK = False
    Me.Show 1, frmMain
    lng����ID = mlng����ID
    zlShowPatiMerge = mblnOK
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdPati_Click()
    frmPatiSel.mstrPrivs = ""
    frmPatiSel.Show 1, Me
    
    If frmPatiSel.mlng����ID <> 0 Then
        If ExistFeeInsurePatient(frmPatiSel.mlng����ID) Then
            MsgBox "��ҽ�����˴���δ�����,���Ƚ�����ٺϲ���", vbExclamation, gstrSysName: Exit Sub
        End If
    
        Call ShowPatiInfo(frmPatiSel.mlng����ID, 1)
    End If
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
    If glngSys Like "8??" Then
        Caption = "�ͻ��ϲ�"
        lblסԺ��(0).Visible = False
        lblסԺ��(1).Visible = False
        txtסԺ��(0).Visible = False
        txtסԺ��(1).Visible = False
        
        lbl����(0).Visible = False
        lbl����(1).Visible = False
        lbl��λ(0).Visible = False
        lbl��λ(1).Visible = False
        lblסԺ����(0).Visible = False
        lblסԺ����(1).Visible = False
    
        txt����(0).Visible = False
        txt����(1).Visible = False
        txt��λ(0).Visible = False
        txt��λ(1).Visible = False
        txtסԺ����(0).Visible = False
        txtסԺ����(1).Visible = False
        
        fra(0).Caption = "���ϲ��Ŀͻ���Ϣ"
        fra(1).Caption = "Ҫ�����Ŀͻ���Ϣ"
        
        fra(0).Height = fra(0).Height - 750
        fra(1).Height = fra(1).Height - 750
        Me.Height = Me.Height - 750
        
        cmdExit.Top = cmdExit.Top - 750
        cmdHelp.Top = cmdHelp.Top - 750
    End If
    
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
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strסԺ�� As String, str����� As String
    
    On Error GoTo errH
    'by lesfeng 2010-03-08 �����Ż� select *
    strSQL = "Select ����ID,�����,סԺ��,���￨��,����,�Ա�,����,��������,�����ص�,���֤��,���,ְҵ,����,����,����,ѧ��,����״��,��ͥ��ַ,��ͥ�绰" & _
             "  From ������Ϣ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
        
    If rsTmp.EOF Then
        MsgBox "δ���ָò��˵������Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If

    txt����(X).Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    txt�Ա�(X).Text = IIf(IsNull(rsTmp!�Ա�), "", rsTmp!�Ա�)
    txt��������(X).Text = Format(IIf(IsNull(rsTmp!��������), "", rsTmp!��������), "yyyy��MM��dd��")
    txt����(X).Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    txt����(X).Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    txtѧ��(X).Text = IIf(IsNull(rsTmp!ѧ��), "", rsTmp!ѧ��)
    txt���(X).Text = IIf(IsNull(rsTmp!���), "", rsTmp!���)
    txtְҵ(X).Text = IIf(IsNull(rsTmp!ְҵ), "", rsTmp!ְҵ)
    txt���֤��(X).Text = IIf(IsNull(rsTmp!���֤��), "", rsTmp!���֤��)
    txt�����ص�(X).Text = IIf(IsNull(rsTmp!�����ص�), "", rsTmp!�����ص�)
    txt��ͥ��ַ(X).Text = IIf(IsNull(rsTmp!��ͥ��ַ), "", rsTmp!��ͥ��ַ)
    txt����״��(X).Text = IIf(IsNull(rsTmp!����״��), "", rsTmp!����״��)
    
    str����� = IIf(IsNull(rsTmp!�����), 0, rsTmp!�����)
    strסԺ�� = IIf(IsNull(rsTmp!סԺ��), 0, rsTmp!סԺ��)
    'by lesfeng 2010-03-08 �����Ż� select A.*
    strSQL = "Select A.��Ժ����,A.סԺ��,A.��Ժ����,A.����ID,A.��ҳID,B.���� as ���� From ������ҳ A,���ű� B" & _
        " Where A.��ҳID=(Select Max(��ҳID) From ������ҳ Where ����ID=[1])" & _
        " And A.��Ժ����ID=B.ID And A.����ID=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    
    If rsTmp.EOF Then
        If glngSys Like "8??" Then
            txt״̬(X).Text = "����"
        Else
            txt״̬(X).Text = "����"
        End If
        lblסԺ��(X).Caption = "�����:"
        txtסԺ��(X).Text = IIf(str����� = "0", "", str�����)
        txt����(X).Text = ""
        txt��λ(X).Text = ""
        txtסԺ����(X).Text = ""
    Else
        txt״̬(X).Text = IIf(IsNull(rsTmp!��Ժ����), "��Ժ", "��Ժ")
        lblסԺ��(X).Caption = "סԺ��:"
        txtסԺ��(X).Text = IIf(strסԺ�� = 0, "", strסԺ��)
        txt����(X).Text = rsTmp!����
        txt��λ(X).Text = IIf(IsNull(rsTmp!��Ժ����), "��ͥ", rsTmp!��Ժ����)
        txtסԺ����(X).Text = rsTmp!��ҳID
    End If
            
    fra(X).Tag = lngID
    
    ShowPatiInfo = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload frmPatiSel
End Sub

Private Sub cmdMerge_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim rsPatiS As ADODB.Recordset
    Dim rsPatiO As ADODB.Recordset
    Dim strSQL As String, Curdate As Date
    Dim i As Integer, j As Integer
    Dim str�ϲ�ԭ�� As String
    
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
        
    Set rsPatiS = GetPatiInfo(CLng(fra(0).Tag))
    Set rsPatiO = GetPatiInfo(CLng(fra(1).Tag))
    
    'A��B��һ��������ԤԼ��Ժ
    If Not IsNull(rsPatiS!��ҳID) And Nvl(rsPatiS!��ҳID, 0) = 0 Then
        MsgBox "����:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]������ԤԼ��Ժ�Ǽǣ�����ȡ���õǼǡ�", vbInformation, gstrSysName
    End If
    If Not IsNull(rsPatiO!��ҳID) And Nvl(rsPatiO!��ҳID, 0) = 0 Then
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
        Curdate = zlDatabase.Currentdate
        rsPatiS.MoveFirst
        For i = 1 To rsPatiS.RecordCount
            rsPatiO.MoveFirst
            For j = 1 To rsPatiO.RecordCount
                If Not (rsPatiO!��Ժ���� >= IIf(IsNull(rsPatiS!��Ժ����), Curdate, rsPatiS!��Ժ����) Or _
                    IIf(IsNull(rsPatiO!��Ժ����), Curdate, rsPatiO!��Ժ����) <= rsPatiS!��Ժ����) Then
                    MsgBox "���ֲ���:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]�� " & rsPatiS!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiS!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiS!��Ժ����), Curdate, rsPatiS!��Ժ����), "yyyy-MM-dd") & vbCrLf & _
                        "�벡��:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]�ĵ� " & rsPatiO!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiO!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiO!��Ժ����), Curdate, rsPatiO!��Ժ����), "yyyy-MM-dd") & _
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
    If zlCommFun.ActualLen(str�ϲ�ԭ��) > 250 Then
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
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    Screen.MousePointer = 0
        
    '�ϲ���Ӧֻʣһ������
    'by lesfeng 2010-03-08 �����Ż�
    strSQL = "Select ����ID From ������Ϣ Where ����ID IN([1],[2])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(fra(0).Tag), Val(fra(1).Tag))
    
    mlng����ID = rsTmp!����ID

    If glngSys Like "8??" Then
        MsgBox "�ͻ��ϲ��ɹ�,�ϲ���Ŀͻ�IDΪ " & mlng����ID & "��", vbInformation, gstrSysName
    Else
        MsgBox "���˺ϲ��ɹ�,�ϲ���Ĳ���IDΪ " & mlng����ID & "��", vbInformation, gstrSysName
    End If
    Call ClearPatiInfo(1)
    Call ShowPatiInfo(mlng����ID, 0)
    Unload frmPatiSel
    mblnOK = True
    cmdPati.SetFocus
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPatiInfo(lng����ID As Long) As ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    '��ҳID=0ʱ(����NULL)����ʾԤԼ��Ժ
    strSQL = _
        " Select A.����ID,Decode(B.����ID,NULL,NULL,Nvl(B.��ҳID,0)) as ��ҳID," & _
        "            decode(B.����,NULL,A.����,B.����) as ����,A.סԺ��,B.��Ժ����,B.��Ժ����" & _
        " From ������Ϣ A,������ҳ B" & _
        " Where A.����ID=B.����ID(+) And A.����ID=[1]" & _
        " Order by Nvl(B.��ҳID,0)"
    On Error GoTo errH

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    If Not rsTmp.EOF Then Set GetPatiInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
