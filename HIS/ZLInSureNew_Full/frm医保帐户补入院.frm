VERSION 5.00
Begin VB.Form frmҽ���ʻ�����Ժ 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���˲���ҽ����Ժ�Ǽ�"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmҽ���ʻ�����Ժ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "�Ǽ�(&X)"
      Height          =   350
      Left            =   6090
      TabIndex        =   9
      Top             =   6015
      Width           =   1100
   End
   Begin VB.Frame fra������Ϣ 
      Caption         =   "��������Ϣ��"
      ForeColor       =   &H00C00000&
      Height          =   705
      Left            =   75
      TabIndex        =   28
      Top             =   1815
      Width           =   8745
      Begin VB.TextBox txt������� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtԤ����� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txt������ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7380
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox txt������ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ�����"
         Height          =   180
         Left            =   2370
         TabIndex        =   31
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ�����"
         Height          =   180
         Left            =   375
         TabIndex        =   29
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   6765
         TabIndex        =   35
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   4695
         TabIndex        =   33
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.Frame fra������Ϣ 
      Caption         =   "��������Ϣ��"
      ForeColor       =   &H00C00000&
      Height          =   3345
      Left            =   75
      TabIndex        =   59
      Top             =   2580
      Width           =   8745
      Begin VB.TextBox txtҽ�Ƹ��� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt�������� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   570
         Width           =   1140
      End
      Begin VB.TextBox txt��ϵ�˹�ϵ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1890
         Width           =   2000
      End
      Begin VB.TextBox txt��� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txtְҵ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt����״�� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txtѧ�� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt�����ص� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   900
         Width           =   3225
      End
      Begin VB.TextBox txt��ͥ��ַ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1230
         Width           =   3150
      End
      Begin VB.TextBox txt�����ʱ� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1170
      End
      Begin VB.TextBox txt��ϵ������ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1170
      End
      Begin VB.TextBox txt��ϵ�˵�ַ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1890
         Width           =   3225
      End
      Begin VB.TextBox txt��ϵ�˵绰 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2220
         Width           =   2000
      End
      Begin VB.TextBox txt������λ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   2220
         Width           =   3225
      End
      Begin VB.TextBox txt��λ�绰 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   2550
         Width           =   2000
      End
      Begin VB.TextBox txt��λ�ʱ� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2550
         Width           =   1170
      End
      Begin VB.TextBox txt��λ������ 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox txt��λ�ʺ� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   2880
         Width           =   3225
      End
      Begin VB.TextBox txt��ͥ�绰 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2000
      End
      Begin VB.TextBox txt���֤�� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   900
         Width           =   3150
      End
      Begin VB.Label lblҽ�Ƹ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ƹ���"
         Height          =   180
         Left            =   345
         TabIndex        =   81
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   6570
         TabIndex        =   80
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl�����ص� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ص�"
         Height          =   180
         Left            =   4470
         TabIndex        =   79
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   180
         Left            =   345
         TabIndex        =   78
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Left            =   4830
         TabIndex        =   77
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lblְҵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ"
         Height          =   180
         Left            =   2685
         TabIndex        =   76
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4830
         TabIndex        =   75
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2685
         TabIndex        =   74
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblѧ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѧ��"
         Height          =   180
         Left            =   6930
         TabIndex        =   73
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lvl����״�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��"
         Height          =   180
         Left            =   345
         TabIndex        =   72
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl��ͥ��ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ"
         Height          =   180
         Left            =   345
         TabIndex        =   71
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lbl��ͥ�绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ�绰"
         Height          =   180
         Left            =   345
         TabIndex        =   70
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label lbl�����ʱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʱ�"
         Height          =   180
         Left            =   4470
         TabIndex        =   69
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lbl��ϵ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ������"
         Height          =   180
         Left            =   4290
         TabIndex        =   68
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˹�ϵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˹�ϵ"
         Height          =   180
         Left            =   165
         TabIndex        =   67
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˵�ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵�ַ"
         Height          =   180
         Left            =   4290
         TabIndex        =   66
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˵绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵绰"
         Height          =   180
         Left            =   165
         TabIndex        =   65
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label lbl������λ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������λ"
         Height          =   180
         Left            =   4470
         TabIndex        =   64
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label lbl��λ�绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�绰"
         Height          =   180
         Left            =   345
         TabIndex        =   63
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label lbl��λ�ʱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʱ�"
         Height          =   180
         Left            =   4470
         TabIndex        =   62
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label lbl��λ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ������"
         Height          =   180
         Left            =   165
         TabIndex        =   61
         Top             =   2940
         Width           =   900
      End
      Begin VB.Label lbl��λ�ʺ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʺ�"
         Height          =   180
         Left            =   4470
         TabIndex        =   60
         Top             =   2940
         Width           =   720
      End
   End
   Begin VB.Frame fra��Ժ��Ϣ 
      Caption         =   "��סԺ��Ϣ��"
      ForeColor       =   &H00C00000&
      Height          =   1695
      Left            =   75
      TabIndex        =   5
      Top             =   30
      Width           =   8730
      Begin VB.CommandButton cmdTurn 
         Caption         =   "�������תסԺ(&T)"
         Height          =   300
         Left            =   5280
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F12(ҽ��������֤)"
         Top             =   225
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.ComboBox cobסԺ���� 
         Height          =   300
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txt��� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3180
      End
      Begin VB.CommandButton cmdYB 
         Caption         =   "��֤(&V)"
         Height          =   300
         Left            =   4440
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F12(ҽ��������֤)"
         Top             =   225
         Width           =   825
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   885
         Width           =   1065
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   885
         Width           =   1170
      End
      Begin VB.TextBox txt��Ժʱ�� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   885
         Width           =   1110
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   885
         Width           =   1110
      End
      Begin VB.TextBox txtҽ���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3225
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   1170
      End
      Begin VB.TextBox txt�ѱ� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   555
         Width           =   1110
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txt����ID 
         Height          =   300
         Left            =   1110
         TabIndex        =   0
         Top             =   225
         Width           =   1110
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   555
         Width           =   1110
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   555
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   330
         TabIndex        =   83
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4800
         TabIndex        =   24
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2715
         TabIndex        =   22
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   6540
         TabIndex        =   26
         Top             =   945
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   690
         TabIndex        =   20
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����"
         Height          =   180
         Left            =   4620
         TabIndex        =   8
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   2535
         TabIndex        =   7
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   690
         TabIndex        =   12
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2715
         TabIndex        =   14
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4800
         TabIndex        =   16
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
         Height          =   180
         Left            =   6900
         TabIndex        =   18
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl����ID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   510
         TabIndex        =   6
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   780
      TabIndex        =   11
      Top             =   6015
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7380
      TabIndex        =   10
      Top             =   6015
      Width           =   1100
   End
End
Attribute VB_Name = "frmҽ���ʻ�����Ժ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;99-���н������Ӹ��Ӳ���(���°�)

Private mlng����ID As Long 'Ҫ�޸Ļ�鿴�Ĳ���ID
Private mlng��ҳID As Long 'Ҫ�޸Ļ�鿴����ҳID
Private mstrҽ���� As String
Public mint���� As Integer
Private mstrNOS As String   'ѡ��ת��ĵ���,Ʊ��,����ID,����(��ҽ��Ϊ��):H0000001,F000023,81235,901;H0000002,F000045,81263,901;...

Private Function ReadCard() As Boolean
'���ܣ���ȡָ��������Ϣ,����ʾ�ڽ�����
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrH
    #If gverControl < 6 Then
        gstrSQL = "Select * From ������Ϣ Where ����ID=" & mlng����ID
    #Else
        gstrSQL = "Select A.����id, A.�����, A.סԺ��, A.���￨��, A.����֤��, A.�ѱ�, A.ҽ�Ƹ��ʽ, A.����, A.�Ա�, A.����, A.��������, A.�����ص�, A.���֤��, A.����֤��, A.���, A.ְҵ, A.����, A.����, A.����, A.ѧ��, A.����״��, A.��ͥ��ַ," & vbNewLine & _
            "      A.��ͥ�绰, A.��ͥ��ַ�ʱ� As �����ʱ�, A.�໤��, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ, A.��ϵ�˵绰, A.��ͬ��λid, A.������λ, A.��λ�绰, A.��λ�ʱ�, A.��λ������, A.��λ�ʺ�, A.������, A.������, A.��������, A.����ʱ��, A.����״̬," & vbNewLine & _
            "      A.��������, A.סԺ����, A.��ǰ����id, A.��ǰ����id, A.��ǰ����, A.��Ժʱ��, A.��Ժʱ��, A.��Ժ, A.Ic����, A.������, A.ҽ����, A.����, A.��ѯ����, A.�Ǽ�ʱ��, A.ͣ��ʱ��, A.����" & vbNewLine & _
            "From ������Ϣ A Where A.����ID =" & mlng����ID
    #End If
    rsTmp.CursorLocation = adUseClient
    
    Call OpenRecordset(rsTmp, Me.Caption)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp.RecordCount <> 1 Then Exit Function
    
    'סԺ��Ϣ
    txt����ID.Locked = True
    txt����ID.Text = mlng����ID
    txt����ID.Locked = False
    
    txt����.Text = rsTmp!����
    txtסԺ��.Text = IIf(IsNull(rsTmp!סԺ��), "", rsTmp!סԺ��)
    
    '������Ϣ
    txt�Ա�.Text = IIf(IsNull(rsTmp!�Ա�), "", rsTmp!�Ա�)
    txt����.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    txt�ѱ�.Text = IIf(IsNull(rsTmp!�ѱ�), "", rsTmp!�ѱ�)
    txtҽ�Ƹ���.Text = IIf(IsNull(rsTmp!ҽ�Ƹ��ʽ), "", rsTmp!ҽ�Ƹ��ʽ)
    txt����.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    txt����.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    txtѧ��.Text = IIf(IsNull(rsTmp!ѧ��), "", rsTmp!ѧ��)
    txt����״��.Text = IIf(IsNull(rsTmp!����״��), "", rsTmp!����״��)
    txtְҵ.Text = IIf(IsNull(rsTmp!ְҵ), "", rsTmp!ְҵ)
    txt���.Text = IIf(IsNull(rsTmp!���), "", rsTmp!���)
    txt��������.Text = Format(IIf(IsNull(rsTmp!��������), "", rsTmp!��������), "yyyy-MM-dd")
    txt���֤��.Text = IIf(IsNull(rsTmp!���֤��), "", rsTmp!���֤��)
    txt�����ص�.Text = IIf(IsNull(rsTmp!�����ص�), "", rsTmp!�����ص�)
    txt��ͥ��ַ.Text = IIf(IsNull(rsTmp!��ͥ��ַ), "", rsTmp!��ͥ��ַ)
    txt��ͥ�绰.Text = IIf(IsNull(rsTmp!��ͥ�绰), "", rsTmp!��ͥ�绰)
    txt�����ʱ�.Text = IIf(IsNull(rsTmp!�����ʱ�), "", rsTmp!�����ʱ�)
    txt��ϵ������.Text = IIf(IsNull(rsTmp!��ϵ������), "", rsTmp!��ϵ������)
    txt��ϵ�˹�ϵ.Text = IIf(IsNull(rsTmp!��ϵ�˹�ϵ), "", rsTmp!��ϵ�˹�ϵ)
    txt��ϵ�˵�ַ.Text = IIf(IsNull(rsTmp!��ϵ�˵�ַ), "", rsTmp!��ϵ�˵�ַ)
    txt��ϵ�˵绰.Text = IIf(IsNull(rsTmp!��ϵ�˵绰), "", rsTmp!��ϵ�˵绰)
    txt������λ.Text = IIf(IsNull(rsTmp!������λ), "", rsTmp!������λ)
    txt��λ�绰.Text = IIf(IsNull(rsTmp!��λ�绰), "", rsTmp!��λ�绰)
    txt��λ�ʱ�.Text = IIf(IsNull(rsTmp!��λ�ʱ�), "", rsTmp!��λ�ʱ�)
    txt��λ������.Text = IIf(IsNull(rsTmp!��λ������), "", rsTmp!��λ������)
    txt��λ�ʺ�.Text = IIf(IsNull(rsTmp!��λ�ʺ�), "", rsTmp!��λ�ʺ�)
        
    '������Ϣ
    txt������.Text = IIf(IsNull(rsTmp!������), "", rsTmp!������)
    txt������.Text = Format(IIf(IsNull(rsTmp!������), "", rsTmp!������), "0.00")
    
    #If gverControl >= 5 Then
        gstrSQL = "Select * From ������� Where ����=1 And ����=2 And ����ID=" & mlng����ID
    #Else
        gstrSQL = "Select * From ������� Where ����=1 And ����ID=" & mlng����ID
    #End If
    Call OpenRecordset(rsTmp, Me.Caption)
    
    If Not rsTmp.EOF Then
        txt�������.Text = Format(IIf(IsNull(rsTmp!�������), 0, rsTmp!�������), "0.00")
        txtԤ�����.Text = Format(IIf(IsNull(rsTmp!Ԥ�����), 0, rsTmp!Ԥ�����), "0.00")
    End If
    
    
    '����ҽ����Ϣ
    txtҽ����.Text = ""
    mstrҽ���� = ""
    
    
    '������ҳ��Ϣ
    gstrSQL = "Select A.��Ժ����,A.��Ժ����,b.���� as ��Ժ����,C.���� as ����ȼ�,A.��Ժ����ID" & _
              " From ������ҳ A,���ű� B,����ȼ� C" & _
              " Where A.����ID=" & mlng����ID & " And A.��ҳID=" & mlng��ҳID & _
              "       and A.��Ժ����ID=B.ID and A.����ȼ�ID=C.���(+) "
    Call OpenRecordset(rsTmp, Me.Caption)
    '2006-06-13 δ��Ʋ����޿�����Ϣ
    txt����.Text = IIf(IsNull(rsTmp!��Ժ����), "��", rsTmp!��Ժ����)
    txt����.Tag = Val("" & rsTmp!��Ժ����ID)
    txt����.Text = IIf(IsNull(rsTmp!����ȼ�), "��", rsTmp!����ȼ�)
    txt����.Text = IIf(IsNull(rsTmp!��Ժ����), "", rsTmp!��Ժ����)
    txt��Ժʱ��.Text = Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm")
    
    '������(2006-2-17):��Ժ���,HIS+������д��Ժ���ʱ�������Ϊ2
    gstrSQL = "Select ������Ϣ" & _
              " From ������" & _
              " Where ����ID=" & mlng����ID & " And ��ҳID=" & mlng��ҳID & " and ������� in (1,2) "
    Call OpenRecordset(rsTmp, Me.Caption)
    If rsTmp.EOF = False Then
        txt���.Text = Nvl(rsTmp("������Ϣ"))
    End If
    
    If gclsInsure.GetCapability(support����¼��������, 0, mint����) = True Then
        txt���.Locked = False
        txt���.BackColor = txt����ID.BackColor
    Else
        txt���.Locked = True
        txt���.BackColor = txtסԺ��.BackColor
    End If
    
    ReadCard = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
'������Ժ�Ǽ�
    
    If mlng����ID = 0 Then
        MsgBox "����ȷ���Ȳ�����Ժ�ǼǵĲ��ˡ�", vbInformation, gstrSysName
        txt����ID.SetFocus
        Exit Sub
    End If
    If mstrҽ���� = "" Then
        MsgBox "������֤�ò����Ƿ���Խ���ҽ����Ժ��", vbInformation, gstrSysName
        cmdYB.SetFocus
        Exit Sub
    End If
    If txt���.Locked = False And txt���.Text = "" Then
        MsgBox "����д��Ժ��ϡ�", vbInformation, gstrSysName
        txt���.SetFocus
        Exit Sub
    End If
    If zlCommFun.StrIsValid(txt���.Text, txt���.MaxLength, txt���.hwnd, "��Ժ���") = False Then
        Exit Sub
    End If
    
    If mint���� = 106 Then '�ڽ�ҽ����Ҫ�жϲ�������
       If Not �жϲ���Ч��_�ɶ��ڽ�(mlng����ID, mlng��ҳID) Then
          Exit Sub
       End If
    End If
    
        '�������תסԺ
    If mstrNOS <> "" Then
        If Not frmChargeTurn.ExecuteTurn(mstrNOS, txtסԺ��.Text, mlng��ҳID, CDate(txt��Ժʱ��.Text), Val(txt����.Tag)) Then
            Exit Sub
        End If
    End If
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    gstrSQL = "zl_������ҳ_����ҽ����Ժ(" & mlng����ID & "," & mlng��ҳID & "," & mint���� & ",'" & txt��� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'$IF HIS9.19
#If gverControl = 0 Then
    If gclsInsure.ComeInSwap(mlng����ID, mlng��ҳID, mstrҽ����) = False Then
        '�Ǽ�ʧ��
        gcnOracle.RollbackTrans
        Exit Sub
    End If
#Else
'$ELSE  HIS+
    If gclsInsure.ComeInSwap(mlng����ID, mlng��ҳID, mstrҽ����, mint����) = False Then
        '�Ǽ�ʧ��
        gcnOracle.RollbackTrans
        Exit Sub
    End If
#End If
'$END IF
gcnOracle.CommitTrans
    MsgBox "����" & txt����.Text & "����ҽ����Ժ�ɹ���" & IIf(mint���� > 900, vbCrLf & "���˷�����ϸ��ҽ�������Ѿ���ҽ���������㡣", "") _
        , vbInformation, gstrSysName
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdTurn_Click()
    gstrDec = "0.00"
    Call frmChargeTurn.ShowME(Me, Val(txt����ID.Text), mstrNOS)
End Sub

Private Sub cmdYB_Click()
'��֤ҽ���������
    Dim lng����ID As Long, int���� As Integer
    Dim strYBPati As String, arr��Ϣ As Variant
    
    If mlng����ID = 0 Then
        MsgBox "����ȷ���Ȳ�����Ժ�ǼǵĲ��ˡ�", vbInformation, gstrSysName
        txt����ID.SetFocus: Exit Sub
    End If
    lng����ID = mlng����ID
    int���� = mint����
'$IF HIS9.19
#If gverControl = 0 Then
    strYBPati = gclsInsure.Identify(1, lng����ID)
'$ELSE
#Else
    strYBPati = gclsInsure.Identify(1, lng����ID, int����)
#End If
'$END IF
    If strYBPati = "" Then
        MsgBox "�ò��������֤ʧ�ܡ�", vbInformation, gstrSysName
        cmdYB.SetFocus: Exit Sub
    End If
    
    arr��Ϣ = Split(strYBPati, ";")
    If lng����ID <> 0 Then mlng����ID = lng����ID
    
    '�ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID,...
    If UBound(arr��Ϣ) >= 8 Then
        '������ݺϲ�����ID�����仯
        If Val(arr��Ϣ(8)) <> Val(txt����ID.Text) Then
            txt����ID.Text = "-" & Val(arr��Ϣ(8))
            Call GetPatient
            Call ReadCard
        End If
        
        txtҽ����.Text = arr��Ϣ(1)
        mstrҽ���� = txtҽ����.Text
        
        txt����.Text = arr��Ϣ(3)
        txt�Ա�.Text = arr��Ϣ(4)
        txt��������.Text = arr��Ϣ(5)
        txt���֤��.Text = arr��Ϣ(6)
        
        If IsZLHIS10 Then cmdTurn.Visible = True
        
        cmdOK.SetFocus
    End If
End Sub

Private Sub cobסԺ����_Click()
 If mint���� <> TYPE_���������� And mint���� <> TYPE_������ Then Exit Sub
    mlng��ҳID = cobסԺ����.ItemData(cobסԺ����.ListIndex)
  Call ReadCard
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Call cmdYB_Click
    End If
End Sub

Private Sub Form_Load()
    mstrNOS = ""
    mlng����ID = 0
    mlng��ҳID = 0
End Sub

Private Sub txt����ID_Change()
    If txt����ID.Locked = False Then
        mlng����ID = 0
        mlng��ҳID = 0
    End If
End Sub

Private Sub txt����ID_GotFocus()
    zlControl.TxtSelAll txt����ID
End Sub

Private Sub txt����ID_KeyPress(KeyAscii As Integer)
    Dim lng����ID  As Long
    
    'ת���ɴ�д(���ֲ��ɴ���)
    If KeyAscii > 0 Then KeyAscii = asc(UCase(Chr(KeyAscii)))
    
    If InStr("0123456789", Chr(KeyAscii)) = 0 And (txt����ID.Text = "" Or txt����ID.SelLength = Len(txt����ID.Text)) Then
        txt����ID.MaxLength = 15
    End If
    
    If Len(Trim(Me.txt����ID.Text)) = 0 And KeyAscii = 13 Then
        If frmҽ������ѡ��.Get����(lng����ID) = True Then
            txt����ID.Text = "A" & lng����ID
        End If
    End If
    Me.Refresh
    
    'ˢ����ϻ���������س�
    If (KeyAscii = 13 And Trim(txt����ID.Text) <> "") Then
        If Val(txt����ID.Text) = mlng����ID And mlng����ID > 0 Then
            If mstrҽ���� = "" Then
                cmdYB.SetFocus
            Else
                cmdOK.SetFocus
            End If
            Exit Sub
        End If
        
        If KeyAscii <> 13 Then
            txt����ID.Text = txt����ID.Text & Chr(KeyAscii)
            txt����ID.SelStart = Len(txt����ID.Text)
        End If
        KeyAscii = 0
        
        '20040923:���˺����
        Call LoadסԺ����
        
        If Not GetPatient() Then
            MsgBox "û�з��ָò��˵�סԺ��Ϣ,���������룡", vbInformation, gstrSysName
            txt����ID.Text = ""
            txt����ID.SetFocus
            Exit Sub
        Else
            Call ReadCard
            cmdYB.SetFocus
        End If
    End If

End Sub

Private Function GetPatient() As Boolean
''���ܣ���ȡ������Ϣ
'����:�Ƿ��ȡ�ɹ�,�ɹ�ʱrsInfo�а���������Ϣ,ʧ��ʱrsInfo=Close
    Dim rsInfo As New ADODB.Recordset
    Dim strCode As String
    Dim lngסԺ���� As Long
    '���˺�:2004/09/23:ȡ���˳�Ժ���˺����������
    Dim bln���� As Boolean
    bln���� = (mint���� = TYPE_���������� Or mint���� = TYPE_������)
    
    strCode = Trim(txt����ID.Text)
    On Error GoTo ErrH
    If bln���� Then
        lngסԺ���� = cobסԺ����.ItemData(cobסԺ����.ListIndex)
    Else
        lngסԺ���� = 1
    End If
    If (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '����ID
        gstrSQL = _
            "Select C.����ID,C.��ҳID" & _
            " From ������Ϣ A,������ҳ C" & _
            " Where A.����ID=C.����ID " & IIf(bln����, " And C.��ҳid=" & lngסԺ����, " And Nvl(A.סԺ����,0)=C.��ҳID") & _
            "   And A.����ID=" & Val(Mid(strCode, 2)) & _
            "     " & IIf(bln����, "", "  and C.���� is null and C.��Ժ���� is null")
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then 'סԺ��
        gstrSQL = _
            "Select C.����ID,C.��ҳID" & _
            " From ������Ϣ A,������ҳ C" & _
            " Where A.����ID=C.����ID " & IIf(bln����, " And C.��ҳid=" & lngסԺ����, " And Nvl(A.סԺ����,0)=C.��ҳID") & _
            "       And A.סԺ��=" & Mid(strCode, 2) & _
            "      " & IIf(bln����, "", " and C.���� is null and C.��Ժ���� is null")
    ElseIf (Left(strCode, 1) = "C" Or Left(strCode, 1) = ";") And IsNumeric(Split(Mid(strCode, 2), "?")(0)) Then 'סԺ��
        gstrSQL = _
            "Select C.����ID,C.��ҳID" & _
            " From ������Ϣ A,������ҳ C" & _
            " Where A.����ID=C.����ID " & IIf(bln����, " And C.��ҳid=" & lngסԺ����, " And Nvl(A.סԺ����,0)=C.��ҳID") & _
            "       And A.���￨��='" & Split(Mid(strCode, 2), "?")(0) & _
            "'      " & IIf(bln����, "", " and C.���� is null and C.��Ժ���� is null")
    Else '��������
        gstrSQL = _
            "Select C.����ID,C.��ҳID" & _
            " From ������Ϣ A,������ҳ C" & _
            " Where A.����ID=C.����ID " & IIf(bln����, " And C.��ҳid=" & lngסԺ����, " And Nvl(A.סԺ����,0)=C.��ҳID") & _
            "       And A.����='" & strCode & _
            "'    " & IIf(bln����, "", "   and C.���� is null and C.��Ժ���� is null")
    End If
    
    rsInfo.CursorLocation = adUseClient
    Call OpenRecordset(rsInfo, Me.Caption)
    '���Ը�ʱ��
    txt��Ժʱ��.Locked = Not bln����
    '��ȡʧ��
    If rsInfo.EOF Then
        Exit Function
    End If
        
    mlng����ID = rsInfo("����ID")
    mlng��ҳID = rsInfo("��ҳID")
    
    GetPatient = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function LoadסԺ����() As Boolean
     Dim rsInfo  As New ADODB.Recordset
     Dim strCode  As String
     Dim bln���� As Boolean
    '����סԺ����
    bln���� = (mint���� = TYPE_���������� Or mint���� = TYPE_������)
    
    strCode = Trim(txt����ID.Text)
    If (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '����ID
        gstrSQL = _
            "Select C.����ID,C.��ҳID" & _
            " From ������ҳ C" & _
            " Where ����ID=" & Val(Mid(strCode, 2)) & _
            "   order by C.��ҳid Desc"
            
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then 'סԺ��
        gstrSQL = _
            "Select C.����ID,C.��ҳID" & _
            " From ������Ϣ A,������ҳ C" & _
            " Where A.����ID=C.����ID  And A.סԺ��=" & Mid(strCode, 2) & _
            "      " & IIf(bln����, "", " and C.���� is null and C.��Ժ���� is null") & _
            "   order by C.��ҳid Desc"
    Else '��������
        gstrSQL = _
            "Select C.����ID,C.��ҳID" & _
            " From ������Ϣ A,������ҳ C" & _
            " Where A.����ID=C.����ID  And A.����='" & strCode & _
            "'    " & IIf(bln����, "", "   and C.���� is null and C.��Ժ���� is null") & _
            "   order by C.��ҳid Desc"

    End If
    zlDatabase.OpenRecordset rsInfo, gstrSQL, "��ȡסԺ����"
    
     With rsInfo
        cobסԺ����.Clear
        Do While Not rsInfo.EOF
            cobסԺ����.AddItem Nvl(!��ҳID, 0) & "��"
            cobסԺ����.ItemData(cobסԺ����.NewIndex) = Nvl(!��ҳID, 0)
            .MoveNext
        Loop
        If cobסԺ����.ListCount <> 0 Then cobסԺ����.ListIndex = 0
        cobסԺ����.Enabled = bln����
     End With
     
    
End Function

