VERSION 5.00
Begin VB.Form frmMergePatient 
   Caption         =   "���˺ϲ�"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   Icon            =   "frmMergePatient.frx":0000
   LinkTopic       =   "Form1"
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
      TabIndex        =   0
      Top             =   630
      Width           =   1100
   End
   Begin VB.Frame fra 
      Caption         =   "Ҫ�����Ĳ�����Ϣ"
      Height          =   4725
      Index           =   1
      Left            =   3330
      TabIndex        =   5
      Top             =   90
      Width           =   3135
      Begin VB.TextBox txt״̬ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   270
         Width           =   2000
      End
      Begin VB.TextBox txtסԺ�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   525
         Width           =   2000
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   780
         Width           =   2000
      End
      Begin VB.TextBox txt�Ա� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1035
         Width           =   2000
      End
      Begin VB.TextBox txt�������� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1290
         Width           =   2000
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1545
         Width           =   2000
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2000
      End
      Begin VB.TextBox txtѧ�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2000
      End
      Begin VB.TextBox txt����״�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   2325
         Width           =   2000
      End
      Begin VB.TextBox txtְҵ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2580
         Width           =   2000
      End
      Begin VB.TextBox txt��� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   2835
         Width           =   2000
      End
      Begin VB.TextBox txt���֤�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   3090
         Width           =   2000
      End
      Begin VB.TextBox txt�����ص� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3345
         Width           =   2000
      End
      Begin VB.TextBox txt��ͥ��ַ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2000
      End
      Begin VB.TextBox txt������ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3870
         Width           =   2000
      End
      Begin VB.TextBox txtסԺ���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4140
         Width           =   2000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "������:"
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   69
         Top             =   3840
         Width           =   630
      End
      Begin VB.Label lbl״̬ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "״̬:"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   1
         Left            =   675
         TabIndex        =   67
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   675
         TabIndex        =   65
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
         Left            =   675
         TabIndex        =   64
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
         Left            =   315
         TabIndex        =   63
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
         Left            =   675
         TabIndex        =   62
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
         Left            =   675
         TabIndex        =   61
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
         Left            =   675
         TabIndex        =   60
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
         Left            =   315
         TabIndex        =   59
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
         Left            =   675
         TabIndex        =   58
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
         Left            =   675
         TabIndex        =   57
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
         Left            =   315
         TabIndex        =   56
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
         Left            =   315
         TabIndex        =   55
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
         Left            =   315
         TabIndex        =   54
         Top             =   3600
         Width           =   810
      End
      Begin VB.Label lblסԺ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ҳID:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   53
         Top             =   4140
         Width           =   990
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   52
         Top             =   525
         Width           =   630
      End
   End
   Begin VB.Frame fra 
      Caption         =   "���ϲ��Ĳ�����Ϣ"
      Height          =   4725
      Index           =   0
      Left            =   135
      TabIndex        =   4
      Top             =   90
      Width           =   3135
      Begin VB.TextBox txtסԺ���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   4140
         Width           =   2000
      End
      Begin VB.TextBox txt������ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3870
         Width           =   2000
      End
      Begin VB.TextBox txt��ͥ��ַ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2000
      End
      Begin VB.TextBox txt�����ص� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3345
         Width           =   2000
      End
      Begin VB.TextBox txt���֤�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3090
         Width           =   2000
      End
      Begin VB.TextBox txt��� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2835
         Width           =   2000
      End
      Begin VB.TextBox txtְҵ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2580
         Width           =   2000
      End
      Begin VB.TextBox txt����״�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2325
         Width           =   2000
      End
      Begin VB.TextBox txtѧ�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2000
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2000
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1545
         Width           =   2000
      End
      Begin VB.TextBox txt�������� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1290
         Width           =   2000
      End
      Begin VB.TextBox txt�Ա� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1035
         Width           =   2000
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   780
         Width           =   2000
      End
      Begin VB.TextBox txtסԺ�� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   525
         Width           =   2000
      End
      Begin VB.TextBox txt״̬ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   270
         Width           =   2000
      End
      Begin VB.Label lbl״̬ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "״̬:"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   0
         Left            =   660
         TabIndex        =   68
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "������:"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   66
         Top             =   3840
         Width           =   630
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   19
         Top             =   525
         Width           =   630
      End
      Begin VB.Label lblסԺ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ҳID:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   4140
         Width           =   990
      End
      Begin VB.Label lbl��ͥ��ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   17
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
         Left            =   300
         TabIndex        =   16
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
         Left            =   300
         TabIndex        =   15
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
         Left            =   660
         TabIndex        =   14
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
         Left            =   660
         TabIndex        =   13
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
         Left            =   300
         TabIndex        =   12
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
         Left            =   660
         TabIndex        =   11
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
         Left            =   660
         TabIndex        =   10
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
         Left            =   660
         TabIndex        =   9
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
         Left            =   300
         TabIndex        =   8
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
         Left            =   660
         TabIndex        =   7
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
         Left            =   660
         TabIndex        =   6
         Top             =   780
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdMerge 
      Caption         =   "�ϲ�(&M)"
      Height          =   350
      Left            =   6675
      TabIndex        =   1
      Top             =   1155
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   6675
      TabIndex        =   2
      Top             =   3570
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6675
      TabIndex        =   3
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
Private mlng����ID As Long '�룺��ʼҪ�ϲ��Ĳ���ID
Private mlng�������� As Long '�룺�������˵Ĳ���ID
Private mstrPrivs As String
Private mstrסԺ�� As String '��:��������סԺ��

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdTurn_Click()
    Dim lngTmp As Long
    
    lngTmp = fra(0).Tag
    Call ShowPatiInfo(CLng(fra(1).Tag), "", 0)
    Call ShowPatiInfo(lngTmp, "", 1)
End Sub



Private Sub ClearPatiInfo(x As Integer)
'���ܣ����һ��������Ϣ
'������x=�ؼ�����,0=Դ����,1=Ŀ�겡��
    txt����(x).Text = ""
    txt�Ա�(x).Text = ""
    txt��������(x).Text = ""
    txt����(x).Text = ""
    txt����(x).Text = ""
    txtѧ��(x).Text = ""
    txt���(x).Text = ""
    txtְҵ(x).Text = ""
    txt���֤��(x).Text = ""
    txt�����ص�(x).Text = ""
    txt��ͥ��ַ(x).Text = ""
    txt����״��(x).Text = ""
    txt״̬(x).Text = ""
    lblסԺ��(x).Caption = "סԺ��:"
    txtסԺ��(x).Text = ""
    txt������(x).Text = ""
    txtסԺ����(x).Text = ""
    fra(x).Tag = ""
End Sub

Private Sub cmdMerge_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim rsPatiS As ADODB.Recordset
    Dim rsPatiO As ADODB.Recordset
    Dim strSQL As String, curDate As Date
    Dim i As Integer, j As Integer
    Dim str�ϲ�ԭ�� As String
        
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
        curDate = zlDatabase.Currentdate
        rsPatiS.MoveFirst
        For i = 1 To rsPatiS.RecordCount
            rsPatiO.MoveFirst
            For j = 1 To rsPatiO.RecordCount
'                If Not (rsPatiO!��Ժ���� >= IIf(IsNull(rsPatiS!��Ժ����), curDate, rsPatiS!��Ժ����) Or _
'                    IIf(IsNull(rsPatiO!��Ժ����), curDate, rsPatiO!��Ժ����) <= rsPatiS!��Ժ����) Then
'                    If MsgBox("���ֲ���:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]�� " & rsPatiS!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiS!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiS!��Ժ����), curDate, rsPatiS!��Ժ����), "yyyy-MM-dd") & vbCrLf & _
'                        "�벡��:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]�ĵ� " & rsPatiO!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiO!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiO!��Ժ����), curDate, rsPatiO!��Ժ����), "yyyy-MM-dd") & _
'                        vbCrLf & "���ཻ�棬Ӧ�ò���ͬһ�����ˣ�ȷʵҪ�ϲ���", _
'                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'                End If
                 If Not (rsPatiO!��Ժ���� >= IIf(IsNull(rsPatiS!��Ժ����), curDate, rsPatiS!��Ժ����) Or _
                    IIf(IsNull(rsPatiO!��Ժ����), curDate, rsPatiO!��Ժ����) <= rsPatiS!��Ժ����) Then
                    MsgBox "���ֲ���:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]�� " & rsPatiS!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiS!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiS!��Ժ����), curDate, rsPatiS!��Ժ����), "yyyy-MM-dd") & vbCrLf & _
                        "�벡��:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]�ĵ� " & rsPatiO!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiO!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiO!��Ժ����), curDate, rsPatiO!��Ժ����), "yyyy-MM-dd") & _
                        vbCrLf & "���ཻ�棬���ܺϲ���", vbInformation, gstrSysName
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
    strSQL = "Select ����ID From ������Ϣ Where ����ID IN(" & Val(fra(0).Tag) & "," & Val(fra(1).Tag) & ")"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    mlng����ID = rsTmp!����ID

    If gclsPros.SysNo Like "8??" Then
        MsgBox "�ͻ��ϲ��ɹ�,�ϲ���Ŀͻ�IDΪ " & mlng����ID & "��", vbInformation, gstrSysName
    Else
        MsgBox "���˺ϲ��ɹ�,�ϲ���Ĳ���IDΪ " & mlng����ID & "��", vbInformation, gstrSysName
    End If
    
    
'    Call ClearPatiInfo(1)
    '56792:������,2012-12-12,�ϲ�֮�󲡰���Ӧ�ô�""������0
    Call ShowPatiInfo(mlng����ID, "", 0)
    mstrסԺ�� = txtסԺ��(0).Text
    Unload Me
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
        " A.����,A.סԺ��,B.��Ժ����,B.��Ժ����" & _
        " From ������Ϣ A,������ҳ B" & _
        " Where A.����ID=B.����ID(+) And A.����ID=" & lng����ID & _
        " Order by Nvl(B.��ҳID,0)"
    On Error GoTo errH
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If Not rsTmp.EOF Then Set GetPatiInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function ShowPatiInfo(lngID As Long, str������ As String, x As Integer) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    Err = 0: On Error GoTo errHandle
    
    '56792:������,2012-12-12
    If str������ = "" Then
        '��Ҫ���Ҫ�ϲ��Ĳ����Ƿ��Ѿ���Ŀ
        strSQL = "Select A.����id, A.סԺ��, A.����, A.�Ա�, A.��������, A.����, A.����, A.ѧ��, A.����״��, A.ְҵ, A.���, A.���֤��, " & _
                 "       A.�����ص� , A.��ͥ��ַ, A.��Ժʱ��, B.������, C.�����ҳID ��ҳID " & _
                 "From ������Ϣ A, סԺ������¼ B, (Select Max(B.��ҳid) ��ҳid,A.����ID,Max(A.��ҳID) �����ҳID From ������ҳ A,סԺ������¼ B Where A.����ID=B.����ID(+) ANd A.����id =[1] Group by a.����ID) C " & _
                 "Where A.����ID = C.����ID And C.����ID=B.����ID(+) And C.��ҳid=B.��ҳID(+) And A.����ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
        
    Else
        '��ȡ�Ѿ���Ŀ�ı�������
        strSQL = "Select A.����id, A.סԺ��, A.����, A.�Ա�, A.��������, A.����, A.����, A.ѧ��, A.����״��, A.ְҵ, A.���, A.���֤��, " & _
                 "       A.�����ص� , A.��ͥ��ַ, A.��Ժʱ��,B.������, C.�����ҳID ��ҳID " & _
                 "From ������Ϣ A, סԺ������¼ B, (Select A.����id,Max(B.��ҳid) ��ҳid,Max(A.��ҳID) �����ҳID From ������ҳ A,סԺ������¼ B Where A.����ID=B.����id  and B.������= [1] group by A.����ID) C " & _
                 "Where A.����ID = B.����ID and a.����id=C.����id and b.��ҳID=c.��ҳID  "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str������)
    End If
    If Not rsTemp.EOF Then
        
        txt����(x).Text = IIf(IsNull(rsTemp!����), "", rsTemp!����)
        txt�Ա�(x).Text = IIf(IsNull(rsTemp!�Ա�), "", rsTemp!�Ա�)
        txt��������(x).Text = Format(IIf(IsNull(rsTemp!��������), "", rsTemp!��������), "yyyy��MM��dd��")
        txt����(x).Text = IIf(IsNull(rsTemp!����), "", rsTemp!����)
        txt����(x).Text = IIf(IsNull(rsTemp!����), "", rsTemp!����)
        txtѧ��(x).Text = IIf(IsNull(rsTemp!ѧ��), "", rsTemp!ѧ��)
        txt���(x).Text = IIf(IsNull(rsTemp!���), "", rsTemp!���)
        txtְҵ(x).Text = IIf(IsNull(rsTemp!ְҵ), "", rsTemp!ְҵ)
        txt���֤��(x).Text = IIf(IsNull(rsTemp!���֤��), "", rsTemp!���֤��)
        txt�����ص�(x).Text = IIf(IsNull(rsTemp!�����ص�), "", rsTemp!�����ص�)
        txt��ͥ��ַ(x).Text = IIf(IsNull(rsTemp!��ͥ��ַ), "", rsTemp!��ͥ��ַ)
        txt����״��(x).Text = IIf(IsNull(rsTemp!����״��), "", rsTemp!����״��)
        txtסԺ��(x).Text = IIf(IsNull(rsTemp!סԺ��), "", rsTemp!סԺ��)
        txt״̬(x).Text = IIf(IsNull(rsTemp!��Ժʱ��), "��Ժ", "��Ժ")
        txt������(x).Text = IIf(IsNull(rsTemp!������), "", rsTemp!������)
        txtסԺ����(x).Text = IIf(IsNull(rsTemp!��ҳID), "", rsTemp!��ҳID)
        fra(x).Tag = rsTemp!����ID
        ShowPatiInfo = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    ShowPatiInfo = False
End Function

Public Function MergePatient(lngID As Long, str������ As String, frmMain As Object) As String
    '����סԺ��
    mstrסԺ�� = ""
    If ShowPatiInfo(lngID, "", 0) = False Then MergePatient = False: Exit Function
    If ShowPatiInfo(0, str������, 1) = False Then MergePatient = False: Exit Function
    
    '56792:������,2012-12-12
    '����������˶��Ѿ���Ŀ����ʾ����Ա���ܺϲ�
    cmdTurn.Enabled = True
    If Trim(txt������(0).Text) <> "" And Trim(txt������(1).Text) <> "" Then
        ShowMsgbox "�ϲ��ͱ������˵Ĳ������Ѿ���Ŀ����������в��˺ϲ�������"
        MergePatient = txtסԺ��(0).Text
        Exit Function
    ElseIf Trim(txt������(0).Text) = "" And Trim(txt������(1).Text) <> "" Then
        cmdTurn.Enabled = False
    Else
        ShowMsgbox "û����ȡ���������˵Ĳ����ţ����飡"
        MergePatient = txtסԺ��(0).Text
        Exit Function
    End If
    
    frmMergePatient.Show 1, frmMain
    MergePatient = mstrסԺ��
End Function

