VERSION 5.00
Begin VB.Form frmIdentify���� 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   Icon            =   "frmIdentify����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox Txt������Ϣ 
      BackColor       =   &H8000000E&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1860
      TabIndex        =   75
      Top             =   7860
      Width           =   7515
   End
   Begin VB.TextBox txt��ע 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1860
      TabIndex        =   69
      Top             =   7440
      Width           =   7515
   End
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "������(&M)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9600
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2130
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9600
      TabIndex        =   14
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9600
      TabIndex        =   13
      Top             =   510
      Width           =   1335
   End
   Begin VB.TextBox txt������Ϣ 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1860
      TabIndex        =   67
      Top             =   7050
      Width           =   7515
   End
   Begin VB.Frame Frame2 
      Caption         =   "�ۼ���Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   180
      TabIndex        =   41
      Top             =   4050
      Width           =   9195
      Begin VB.TextBox txt��ͨ����ҽ�Ʋ�����ת��ʹ�� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   65
         Top             =   2280
         Width           =   1965
      End
      Begin VB.TextBox txt��ͨ����ҽ�Ʋ������� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   63
         Top             =   1890
         Width           =   1965
      End
      Begin VB.TextBox txt��ͨ����ҽ�Ʋ����𸶱�׼ 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   61
         Top             =   1500
         Width           =   1965
      End
      Begin VB.TextBox txt��ͨ����ҽ�Ʋ����ۼ� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   59
         Top             =   1110
         Width           =   1965
      End
      Begin VB.TextBox txt��ͨ����ҽ�Ʋ����޶� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   57
         Top             =   720
         Width           =   1965
      End
      Begin VB.TextBox txt���֧���ۼ� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6990
         TabIndex        =   55
         Top             =   330
         Width           =   1965
      End
      Begin VB.TextBox txt���ͳ���޶� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   53
         Top             =   2280
         Width           =   1965
      End
      Begin VB.TextBox txtͳ��֧���ۼ� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   51
         Top             =   1890
         Width           =   1965
      End
      Begin VB.TextBox txt����ͳ���޶� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   49
         Top             =   1500
         Width           =   1965
      End
      Begin VB.TextBox txt��֧������ 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   47
         Top             =   1110
         Width           =   1965
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   45
         Top             =   720
         Width           =   1965
      End
      Begin VB.TextBox txtסԺ���� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   43
         Top             =   330
         Width           =   1965
      End
      Begin VB.Label lbl��ͨ����ҽ�Ʋ�����ת��ʹ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ͨ����ҽ�Ʋ�����ת��ʹ��"
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
         Height          =   210
         Left            =   4200
         TabIndex        =   64
         Top             =   2340
         Width           =   2730
      End
      Begin VB.Label lbl����Ա���ﲹ������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ͨ����ҽ�Ʋ�������"
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
         Height          =   210
         Left            =   4620
         TabIndex        =   62
         Top             =   1950
         Width           =   2310
      End
      Begin VB.Label lbl����Ա���ﲹ���𸶱�׼ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ͨ����ҽ�Ʋ����𸶱�׼"
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
         Height          =   210
         Left            =   4410
         TabIndex        =   60
         Top             =   1560
         Width           =   2520
      End
      Begin VB.Label lbl����Ա���ﲹ���ۼ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ͨ����ҽ�Ʋ����ۼ�"
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
         Height          =   210
         Left            =   4830
         TabIndex        =   58
         Top             =   1170
         Width           =   2100
      End
      Begin VB.Label lbl����Ա���ﲹ���޶� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ͨ����ҽ�Ʋ����޶�"
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
         Height          =   210
         Left            =   4830
         TabIndex        =   56
         Top             =   780
         Width           =   2100
      End
      Begin VB.Label lbl���֧���ۼ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���֧���ۼ�"
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
         Height          =   210
         Left            =   5670
         TabIndex        =   54
         Top             =   390
         Width           =   1260
      End
      Begin VB.Label lbl���ͳ���޶� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ͳ���޶�"
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
         Height          =   210
         Left            =   390
         TabIndex        =   52
         Top             =   2340
         Width           =   1260
      End
      Begin VB.Label lblͳ��֧���ۼ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ��֧���ۼ�"
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
         Height          =   210
         Left            =   390
         TabIndex        =   50
         Top             =   1950
         Width           =   1260
      End
      Begin VB.Label lbl����ͳ���޶� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ͳ���޶�"
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
         Height          =   210
         Left            =   390
         TabIndex        =   48
         Top             =   1560
         Width           =   1260
      End
      Begin VB.Label lbl��֧������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��֧������"
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
         Height          =   210
         Left            =   390
         TabIndex        =   46
         Top             =   1170
         Width           =   1260
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
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
         Height          =   210
         Left            =   1020
         TabIndex        =   44
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblסԺ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����"
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
         Height          =   210
         Left            =   810
         TabIndex        =   42
         Top             =   390
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3945
      Left            =   180
      TabIndex        =   28
      Top             =   30
      Width           =   9195
      Begin VB.CommandButton cmd���� 
         Caption         =   "����"
         Height          =   350
         Left            =   2940
         TabIndex        =   7
         ToolTipText     =   "�����������"
         Top             =   1875
         Width           =   675
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "����"
         Height          =   350
         Left            =   3600
         TabIndex        =   76
         Top             =   1470
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.OptionButton opt����� 
         Caption         =   "���֤��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   2
         Left            =   3060
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1365
      End
      Begin VB.OptionButton opt����� 
         Caption         =   "IC��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   1
         Left            =   1950
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1170
         Width           =   945
      End
      Begin VB.OptionButton opt����� 
         Caption         =   "�ſ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   0
         Left            =   840
         TabIndex        =   3
         Top             =   1170
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CheckBox chkתԺ���� 
         Caption         =   "תԺ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3210
         TabIndex        =   12
         Top             =   2730
         Width           =   1185
      End
      Begin VB.TextBox txt��Ա��� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   72
         Top             =   338
         Width           =   2595
      End
      Begin VB.TextBox txt�Ա� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7980
         TabIndex        =   70
         Top             =   1118
         Width           =   975
      End
      Begin VB.CheckBox chk���˿���סԺ 
         Caption         =   "���˿���סԺ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   10
         Top             =   2730
         Width           =   1635
      End
      Begin VB.TextBox txt��������� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         TabIndex        =   9
         Top             =   2280
         Width           =   2595
      End
      Begin VB.ComboBox cbo������� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   2595
      End
      Begin VB.CheckBox chk�ƻ����� 
         Caption         =   "�ƻ�����"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1980
         TabIndex        =   11
         Top             =   2730
         Width           =   1185
      End
      Begin VB.TextBox txt�ɷ���� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   40
         Top             =   3450
         Width           =   2595
      End
      Begin VB.TextBox txt�ʻ���� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   38
         Top             =   3060
         Width           =   2595
      End
      Begin VB.TextBox txt��λ���� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   36
         Top             =   2670
         Width           =   2595
      End
      Begin VB.TextBox txt��λ���� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   34
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txt�������� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   27
         Top             =   1890
         Width           =   1335
      End
      Begin VB.TextBox txt���֤�� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   24
         Top             =   1500
         Width           =   2595
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   20
         Top             =   1118
         Width           =   1065
      End
      Begin VB.TextBox txtҽ���չ���Ⱥ 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         TabIndex        =   17
         Top             =   728
         Width           =   2595
      End
      Begin VB.TextBox txt�����ı�� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   32
         Top             =   3450
         Width           =   2595
      End
      Begin VB.TextBox txtҽ���� 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   30
         Top             =   3060
         Width           =   2595
      End
      Begin VB.TextBox txt���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1890
         Width           =   1245
      End
      Begin VB.ComboBox cbo֧����� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2595
      End
      Begin VB.TextBox txt���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1710
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1500
         Width           =   2565
      End
      Begin VB.Label lbl��Ա��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ա���"
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
         Height          =   210
         Left            =   5460
         TabIndex        =   73
         Top             =   390
         Width           =   840
      End
      Begin VB.Label lbl�Ա� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
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
         Height          =   210
         Left            =   7530
         TabIndex        =   71
         Top             =   1170
         Width           =   420
      End
      Begin VB.Label lbl��������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���������"
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
         Height          =   210
         Left            =   600
         TabIndex        =   25
         Top             =   2340
         Width           =   1050
      End
      Begin VB.Label lbl������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
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
         Height          =   210
         Left            =   810
         TabIndex        =   0
         Top             =   390
         Width           =   840
      End
      Begin VB.Label lbl�ɷ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ɷ����"
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
         Height          =   210
         Left            =   5460
         TabIndex        =   39
         Top             =   3502
         Width           =   840
      End
      Begin VB.Label lbl�ʻ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ʻ����"
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
         Height          =   210
         Left            =   5460
         TabIndex        =   37
         Top             =   3112
         Width           =   840
      End
      Begin VB.Label lbl��λ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
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
         Height          =   210
         Left            =   5460
         TabIndex        =   35
         Top             =   2730
         Width           =   840
      End
      Begin VB.Label lbl��λ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
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
         Height          =   210
         Left            =   5460
         TabIndex        =   33
         Top             =   2340
         Width           =   840
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
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
         Height          =   210
         Left            =   5460
         TabIndex        =   26
         Top             =   1950
         Width           =   840
      End
      Begin VB.Label lbl���֤�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
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
         Height          =   210
         Left            =   5460
         TabIndex        =   23
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
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
         Height          =   210
         Left            =   5880
         TabIndex        =   19
         Top             =   1170
         Width           =   420
      End
      Begin VB.Label lblҽ���չ���Ⱥ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ���չ���Ⱥ"
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
         Height          =   210
         Left            =   5040
         TabIndex        =   16
         Top             =   780
         Width           =   1260
      End
      Begin VB.Label lbl�����ı��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ı���"
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
         Height          =   210
         Left            =   600
         TabIndex        =   31
         Top             =   3502
         Width           =   1050
      End
      Begin VB.Label lblҽ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���˱��"
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
         Height          =   210
         Left            =   810
         TabIndex        =   29
         Top             =   3112
         Width           =   840
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
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
         Height          =   210
         Left            =   1230
         TabIndex        =   22
         Top             =   1950
         Width           =   420
      End
      Begin VB.Label lbl֧����� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "֧�����"
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
         Height          =   210
         Left            =   810
         TabIndex        =   18
         Top             =   780
         Width           =   840
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
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
         Height          =   210
         Left            =   1230
         TabIndex        =   21
         Top             =   1560
         Width           =   420
      End
   End
   Begin VB.Label Lab������Ϣ 
      BackStyle       =   0  'Transparent
      Caption         =   "������Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   900
      TabIndex        =   74
      Top             =   7920
      Width           =   840
   End
   Begin VB.Label lbl��ע 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ע"
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
      Height          =   210
      Left            =   1350
      TabIndex        =   68
      Top             =   7500
      Width           =   420
   End
   Begin VB.Label lbl������Ϣ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������Ϣ"
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
      Height          =   210
      Left            =   930
      TabIndex        =   66
      Top             =   7110
      Width           =   840
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr����ʱ�� As String
Private mbytType As Byte
Private mstr�����ı�� As String
Private mstr������� As String
Private mstr������ As String
Private mbln������־ As Boolean
Private mblnOK As Boolean
Private int����סԺ��־ As Integer   '����-0,סԺ-1
Private mstr�����϶���� As String
Private mlng����ID As Long
Private mstr���� As String
Private mstr֧������ As String  '֧����� 31��סԺ��32:�ƻ�������37��תԺ

Private Sub cbo�������_Click()
    chk�ƻ�����.Enabled = (cbo�������.Text Like "*����*")
    chk�ƻ�����.Value = 0
    chkתԺ����.Value = 0
    txt���������.Enabled = False
    chk���˿���סԺ.Enabled = (cbo�������.Text = "���˱���" And mbytType = 0)
    If cbo�������.Text = "������" And (mbytType = 0 Or mbytType = 3) Then
        Me.cbo֧�����.ListIndex = 1
    End If
    'XieRong 2010.10.12 �������ﲡ�˱���ѡ�񴦷��汾��
    txt���������.Enabled = IIf(cbo�������.Text = "������" And (mbytType = 0 Or mbytType = 3) Or cbo֧�����.Text = "��������", True, False)
    lbl���������.Enabled = IIf(cbo�������.Text = "������" And (mbytType = 0 Or mbytType = 3) Or cbo֧�����.Text = "��������", True, False)
   
End Sub

Private Sub cbo֧�����_Click()
    'XieRong 2010.10.12 �������ﲡ�˱���ѡ�񴦷��汾��
    txt���������.Enabled = IIf(cbo�������.Text = "������" And (mbytType = 0 Or mbytType = 3) Or cbo֧�����.Text = "��������", True, False)
    lbl���������.Enabled = IIf(cbo�������.Text = "������" And (mbytType = 0 Or mbytType = 3) Or cbo֧�����.Text = "��������", True, False)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChangePassWord_Click()
    Dim strNewPass As String
    strNewPass = frm�޸�����.ChangePassword("", Me.txt����.Text, 40)
    If strNewPass <> "" Then mstr������ = strNewPass
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    Dim rsTemp As New ADODB.Recordset
    
    If cmdOK.Enabled = False Then Exit Sub
    If Trim(txt����.Text) = "" And opt�����(2).Value = False Then
        MsgBox "δ��ȷ��ˢ��,����ͨ����֤��", vbInformation, gstrSysName
        Exit Sub
    End If
    If Trim(txtҽ����.Text) = "" Then
        MsgBox "δ��ȷ��ˢ��,����ͨ����֤��", vbInformation, gstrSysName
        Exit Sub
    End If
    If Trim(mstr������) <> "" Then
        If ��������_������(txt����.Tag, txt����.Text, mstr������) = False Then Exit Sub
        mstr���� = mstr������
        mstr������ = ""
        txt����.Text = mstr����
    End If
    If chk�ƻ�����.Value = 1 And chkתԺ����.Value = 1 Then MsgBox "���ƻ��������͡�תԺ���ơ�ֻ��ѡ������֮һ���Ҳ���ѡ�����", vbInformation, gstrSysName: Exit Sub
    '2005.11.22,int����סԺ��־,סԺǿ��ѡ�������
    If (int����סԺ��־ = 1 And cbo�������.ListIndex = 0) Then
       MsgBox "��ѡ�������", vbInformation, gstrSysName
       cbo�������.SetFocus
       Exit Sub
    End If
        
    'XieRong 2010.10.12 �������ﲡ�˱���ѡ�񴦷��汾��
    If cbo�������.Text = "������" And (mbytType = 0 Or mbytType = 3) Or cbo֧�����.Text = "��������" Then
        If Trim(txt���������.Text) = "" Then
            MsgBox "��������������¼�봦�������!", vbInformation, gstrSysName
            txt���������.SetFocus
            Exit Sub
        End If
    End If
    If Me.cbo�������.Text = "����ҽ��" And Trim(txt���������.Text) = "" Then
        MsgBox "����ҽ���������¼�봦�������!", vbInformation, gstrSysName
        txt���������.SetFocus
        Exit Sub
    End If
    
    '20111116����ǿ����,�Ƿ�Ϊ����������˲���
     gstr�����־ = cbo֧�����.Text
     gstr���˱�־ = cbo�������.Text
    
    '�п����޸������룬��ɲ��������֤�󷵻ص�XML���ƻ����ٴε��ö���
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' �����
    Call InsertChild(mdomInput.documentElement, "CARDDATA", txt����.Tag)            ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "PASSWORD", txt����.Text)            ' ����
    Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                      ' ��ᱣ�Ϻ�
    Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP��ַ
    Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM��оƬ
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", Me.cbo֧�����.ItemData(Me.cbo֧�����.ListIndex))            ' ֧�����
    If Me.cbo֧�����.Text = "��������" Then Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", mstr����)            ' ���ⲡ
    
    '2005.11.22,int����סԺ��־,ҽ������
    If int����סԺ��־ = 0 Then
        Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo�������.ListIndex + 1)
    Else
        Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo�������.ListIndex)
    End If
    
    Call InsertChild(mdomInput.documentElement, "GSRDBH", mstr�����϶����)
    Call InsertChild(mdomInput.documentElement, "STARTDATE", mstr����ʱ��)           ' ��ʼʱ��
    '���ýӿ�
    If CommServer("GETPSNINFO") = False Then Exit Sub
    
    'mstr���� = Trim(txt����.Text)
    'ҽ���ӿ�δ���ؿ��ţ���ǰ�Ŀ����ֶθ�Ϊ����ſ����ݣ������������Ҫ��
    mstr���� = Me.txt����.Tag
    mstrҽ���� = Trim(txtҽ����.Text)
    mstr�����ı�� = Trim(txt�����ı��.Text)
    mstr���� = Trim(txt����.Text)
    mstr������� = cbo�������.ListIndex + 1
    mbln������־ = (chk�ƻ�����.Value = 1)
    gint���˿���סԺ = Me.chk���˿���סԺ.Value
    
    gstr�������� = txt���������.Text
    'Add By �̳ظ� 2010-01-16 ����ҽѧԺҪ��
    If chk�ƻ�����.Value = 1 Then
        mstr֧������ = "32" '�ƻ�����
    ElseIf chkתԺ����.Value = 1 Then
        mstr֧������ = "37" 'ת��סԺ
    Else
        mstr֧������ = "31" '����סԺ
    End If
    '����˲��˵�ҽ������
'    ҽ����_IN IN ҽ�����˵���.ҽ����%TYPE,
'    סԺ����_IN IN ҽ�����˵���.סԺ����%TYPE,
'    ����_IN IN ҽ�����˵���.����%TYPE,
'    ��֧������_IN IN ҽ�����˵���.��֧������%TYPE,
'    ����ͳ���޶�_IN IN ҽ�����˵���.����ͳ���޶�%TYPE,
'    ͳ��֧���ۼ�_IN IN ҽ�����˵���.ͳ��֧���ۼ�%TYPE,
'    ���ͳ���޶�_IN IN ҽ�����˵���.���ͳ���޶�%TYPE,
'    ���֧���ۼ�_IN IN ҽ�����˵���.���֧���ۼ�%TYPE,
'    ����Ա�����޶�_IN IN ҽ�����˵���.����Ա�����޶�%TYPE,
'    ����Ա�����ۼ�_IN IN ҽ�����˵���.����Ա�����ۼ�%TYPE,
'    ����Ա�𸶱�׼_IN IN ҽ�����˵���.����Ա�𸶱�׼%TYPE,
'    ����Ա��������_IN IN ҽ�����˵���.����Ա��������%TYPE,
'    �μ�75����Ա����_IN IN ҽ�����˵���.�μ�75����Ա����%TYPE)
    On Error GoTo errHand
     '20110812����ǿ����txt��ע.Text��Ϊtxt��ע+Txt������Ϣ,��ҪΪ�˱��湤����Ϣ
    gstrSQL = "zl_ҽ�����˵���_INSERT(" & _
        "'" & mstrҽ���� & "'," & Val(txtסԺ����.Text) & "," & Val(txt����.Text) & "," & Val(txt��֧������.Text) & "," & _
        "" & Val(txt����ͳ���޶�.Text) & "," & Val(txtͳ��֧���ۼ�.Text) & "," & Val(txt���ͳ���޶�.Text) & "," & Val(txt���֧���ۼ�.Text) & "," & _
        "" & Val(txt��ͨ����ҽ�Ʋ����޶�.Text) & "," & Val(txt��ͨ����ҽ�Ʋ����ۼ�.Text) & "," & Val(txt��ͨ����ҽ�Ʋ����𸶱�׼.Text) & "," & _
        "" & Val(txt��ͨ����ҽ�Ʋ�������.Text) & ",'" & txt��ͨ����ҽ�Ʋ�����ת��ʹ��.Text & "','" & txt��ע.Text & "|" & Txt������Ϣ.Text & "','" & gstr�������� & "')"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetIdentify(ByVal bytType As Byte, str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, _
    Optional ByRef bln������־ As Boolean = False, Optional lng����ID As Long, Optional str֧������ As String = "31") As Boolean
    mblnOK = False
    mstr���� = str����
    mstrҽ���� = strҽ����
    mstr���� = ""
    mstr������ = ""
    mstr�����϶���� = ""
    mbytType = bytType
    mlng����ID = 0
    mstr֧������ = ""
    
    gstrSNO = ""
    gstrIDNO = ""
    gstrPSAMNO = ""
    gintType = 1
    frmIdentify����.Show vbModal
    
    GetIdentify = mblnOK
    If mblnOK = True Then
        str���� = mstr���� & "^" & mstr�������
        strҽ���� = mstrҽ����
        str�����ı�� = mstr�����ı��
        str���� = mstr����
        bln������־ = mbln������־
        gstr�����϶���� = mstr�����϶����
        lng����ID = mlng����ID
        str֧������ = mstr֧������
    End If
End Function

Private Sub cmd����_Click()
    Dim IntPort As Integer, intCardType As Integer
    Dim strPort As String
    Dim STRERR As String, strPin As String, strPass As String, strAddr As String
    Dim lngHandle As Long
    On Error GoTo errHand
    
    strPort = GetSetting("ZLSOFT", "����ģ��\������ҽ��", "�˿�", "COM1")
    If strPort = "USB" Then
        IntPort = 100
    Else
        IntPort = Right(strPort, 1)
    End If
    
    '�򿪶�����
    STRERR = Space(2000)
    If SGZ_IFD_Open(IntPort, lngHandle, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��ȡPSAMоƬ����
    STRERR = Space(2000)
    gstrPSAMNO = Space(2000)
    If SGZ_SAM_ReadNmuber(lngHandle, gstrPSAMNO, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    gstrPSAMNO = TruncZero(gstrPSAMNO)
    
    '��ȡ�籣����
    STRERR = Space(2000)
    gstrSNO = Space(2000)
    strPin = "000000"
    strAddr = "MF|EF05|07|$MF|EF06|01|$"
    If SGZ_ICC_ReadCardInfo(lngHandle, intCardType, strPin, strAddr, gstrSNO, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    gstrSNO = TruncZero(gstrSNO)
    gstrIDNO = Split(gstrSNO, "|")(5)
    gstrSNO = Split(gstrSNO, "|")(2)
    txt����.Text = gstrSNO
    
    STRERR = Space(2000)
    strPass = Space(2000)
    If SGZ_IFD_GetPIN(lngHandle, IntPort, strPass, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    strPass = TruncZero(strPass)
    txt����.Text = strPass
    
Exith:
    STRERR = Space(2000)
    If lngHandle > 0 Then Call SGZ_IFD_Close(lngHandle, STRERR)
    
    Call txt����_KeyDown(vbKeyReturn, 0)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    GoTo Exith
End Sub

Private Sub cmd����_Click()
        Dim IntPort As Integer, intCardType As Integer
    Dim strPort As String
    Dim STRERR As String, strPin As String, strPass As String, strAddr As String
    Dim lngHandle As Long
    On Error GoTo errHand
    
    strPort = GetSetting("ZLSOFT", "����ģ��\������ҽ��", "�˿�", "COM1")
    If strPort = "USB" Then
        IntPort = 100
    Else
        IntPort = Right(strPort, 1)
    End If
    
    '�򿪶�����
    STRERR = Space(2000)
    If SGZ_IFD_Open(IntPort, lngHandle, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��ȡPSAMоƬ����
    STRERR = Space(2000)
    gstrPSAMNO = Space(2000)
    If SGZ_SAM_ReadNmuber(lngHandle, gstrPSAMNO, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    gstrPSAMNO = TruncZero(gstrPSAMNO)
    
    STRERR = Space(2000)
    strPass = Space(2000)
    If SGZ_IFD_GetPIN(lngHandle, IntPort, strPass, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    strPass = TruncZero(strPass)
    txt����.Text = strPass
    
Exith:
    STRERR = Space(2000)
    If lngHandle > 0 Then Call SGZ_IFD_Close(lngHandle, STRERR)
    
    Call txt����_KeyDown(vbKeyReturn, 0)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    GoTo Exith
End Sub


Private Sub Form_Activate()
    Dim lng����ID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    mstr����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If mbytType = 1 And mstrҽ���� <> "" Then
        gstrSQL = " Select ����ID,������� From �����ʻ� Where ����=[1] And ҽ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�������", TYPE_������, mstrҽ����)
        If rsTemp.RecordCount <> 0 Then
            lng����ID = rsTemp!����ID
            Me.cbo�������.ListIndex = Nvl(rsTemp!�������, 0)
            
        End If
        'ȡ��Ժ����
        gstrSQL = " Select A.��Ժ���� From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.��ҳID=B.סԺ���� And A.����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ժ����", lng����ID)
        mstr����ʱ�� = Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:mm:ss")
        
        If txt����.Enabled Then Me.txt����.SetFocus
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    '2005.11.22,int����סԺ��־,����cbo�������itemdata
    With cbo֧�����
        .Clear
        If mbytType = 0 Or mbytType = 3 Then
            int����סԺ��־ = 0
            .AddItem "��ͨ����"
            .ItemData(.NewIndex) = 11
            .AddItem "��������"
            .ItemData(.NewIndex) = 18
            With cbo�������
                .Clear
                .AddItem "��ҵְ������ҽ�Ʊ���"
                .AddItem "��ҵ����ҽ�Ʊ���"
                .AddItem "������ҵ��λҽ�Ʊ���"
                .AddItem "��������"
                .AddItem "������ҵ��λ��������"
                .AddItem "������"
                .AddItem "���˱���"
                .ListIndex = 0
            End With
            .ListIndex = 0
         Else
            int����סԺ��־ = 1
            .AddItem "��ͨסԺ"
            .ItemData(.NewIndex) = 31
            With cbo�������
                .Clear
                .AddItem ""
                .AddItem "��ҵְ������ҽ�Ʊ���"
                .AddItem "��ҵ����ҽ�Ʊ���"
                .AddItem "������ҵ��λҽ�Ʊ���"
                .AddItem "��������"
                .AddItem "������ҵ��λ��������"
                .AddItem "������"
                .AddItem "���˱���"
                .ListIndex = 0
            End With
         End If
        .ListIndex = 0
    End With
    chkתԺ����.Visible = False '�𱣶��ں�δ�ã����� 2010-01-18
    
End Sub

Private Sub opt�����_Click(Index As Integer)
    gintType = Index + 1
    txt����.Enabled = (Index <> 1)
    cmdChangePassword.Enabled = (Index <> 2)
    cmd����.Visible = (Index = 1)
    cmd����.Visible = (Index <> 1)
    Select Case Index
    Case 0
        lbl����.Caption = "����"
    Case 1
        lbl����.Caption = "IC����"
    Case 2
        lbl����.Caption = "���֤��"
    End Select
    If Index <> 1 Then txt����.SetFocus
End Sub

Private Sub txt����_GotFocus()
    If gblnLED Then
        zl9LedVoice.Speak "#5"
    End If
End Sub

Private Sub txt����_GotFocus()
    If gblnLED Then
        zl9LedVoice.Speak "#0"
    End If
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str������Ϣ As String
    Dim str���� As String
    Dim rs���� As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Trim(txt����.Text) = "" Then
        MsgBox IIf(opt�����(2).Value = False, "��ˢ����", "���������֤�ţ�"), vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    '2005.11.22,int����סԺ��־,סԺǿ��ѡ�������
    If (int����סԺ��־ = 1 And cbo�������.ListIndex = 0) Then
       MsgBox "��ѡ�������", vbInformation, gstrSysName
       cbo�������.SetFocus
       Exit Sub
    End If
    If opt�����(2).Value Then
        gstrIDNO = txt����.Text     '���֤��
        txt����.Text = ""
    End If
    If Me.cbo�������.Text = "���˱���" Then
        '����ǹ��ˣ����Ȼ�ȡ�����϶���Ϣ�����ٶ���
        If InitXML = False Then Exit Sub
        Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' �����
        Call InsertChild(mdomInput.documentElement, "CARDDATA", txt����.Text)            ' �ſ�����
        Call InsertChild(mdomInput.documentElement, "PASSWORD", txt����.Text)            ' ����
        Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                     ' ��ᱣ�Ϻ�
        Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP��ַ
        Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM��оƬ
        Call InsertChild(mdomInput.documentElement, "PAYTYPE", Me.cbo֧�����.ItemData(Me.cbo֧�����.ListIndex))            ' ֧�����
        If CommServer("GETGSINFO") = False Then Exit Sub
        mstr�����϶���� = frm�����϶����ѡ��.ShowME
    End If
    
    '���������סԺ
    'Modified By ���� 2003-12-03 ������ ԭ����Ժʱȡ������ѡ�񣬸�Ϊ���������ʱ�����û�в��֣�����ѡ��
    mstr���� = ""
    mlng����ID = 0
    If (Me.cbo֧�����.Text = "��������") Then
        gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                " From ���ղ��� A where A.����=" & TYPE_������
        
        Set rs���� = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ������")
        If Not rs���� Is Nothing Then
            mlng����ID = rs����("ID")
            mstr���� = rs����!����
        End If
        If mlng����ID = 0 Then
            MsgBox "����ѡ�����ⲡ�������������ʶ��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If InitXML = False Then Exit Sub
    '�������޸�����
    If Trim(mstr������) <> "" Then
        If ��������_������(txt����.Text, mstr����, mstr������) = False Then Exit Sub
        mstr���� = mstr������
        mstr������ = ""
        txt����.Text = mstr����
    End If

    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' �����
    Call InsertChild(mdomInput.documentElement, "CARDDATA", txt����.Text)            ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "PASSWORD", txt����.Text)            ' ����
    Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                      ' ��ᱣ�Ϻ�
    Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP��ַ
    Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM��оƬ
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", Me.cbo֧�����.ItemData(Me.cbo֧�����.ListIndex))            ' ֧�����
    If Me.cbo֧�����.Text = "��������" Then Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", mstr����)            ' ���ⲡ
    
    '2005.11.22,int����סԺ��־,ҽ������
    If int����סԺ��־ = 0 Then
      Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo�������.ListIndex + 1)
    Else
      Call InsertChild(mdomInput.documentElement, "INSURETYPE", Me.cbo�������.ListIndex)
    End If
    
     '20110812����ǿ�����ж��Ƿ�������������Ϣ
    gstr������Ϣ = mstr�����϶����
    If InStr(mstr�����϶����, "|") > 1 Then
     mstr�����϶���� = Mid(mstr�����϶����, 1, InStr(mstr�����϶����, "|") - 1)
    End If
    'end
    
    Call InsertChild(mdomInput.documentElement, "GSRDBH", mstr�����϶����)
    Call InsertChild(mdomInput.documentElement, "STARTDATE", mstr����ʱ��)           ' ��ʼʱ��
    '���ýӿ�
    If CommServer("GETPSNINFO") = False Then Exit Sub
    
    'ȡ�÷���ֵ
    '������Ϣ
    txt����.Tag = txt����.Text                    '���濨�����ݣ��Ա��������ʱʹ��
    'txt����.Text = GetElemnetValue("CARDID")
    txtҽ����.Text = GetElemnetValue("PERSONCODE")
    txt�����ı��.Text = GetElemnetValue("CENTERCODE")
    txtҽ���չ���Ⱥ.Text = IIf(Val(GetElemnetValue("CAREPSNFLAG")) = 0, "��", "��")
    
    '2005.11.22,int����סԺ��־,סԺ��������ѡ�������,Ĭ���ÿ�
    If int����סԺ��־ = 0 Then
        cbo�������.ListIndex = GetElemnetValue("INSURETYPE") - 1
    Else
'        cbo�������.ListIndex = 0
    End If
    
    '��Ա���    11����ְ��21�����ݣ�32��ʡ�����ݣ�34���������ݣ�41����ͨ����42���ͱ�����43��������Ա��44���������ͥ��45���ضȲм���
    txt��Ա���.Text = GetElemnetValue("PERSONTYPE")
    txt��Ա���.Text = Switch(txt��Ա���.Text = "11", "��ְ", txt��Ա���.Text = "21", "����", _
                      txt��Ա���.Text = "32", "ʡ������", txt��Ա���.Text = "34", "��������", _
                      txt��Ա���.Text = "41", "��ͨ����", txt��Ա���.Text = "42", "�ͱ�����", _
                      txt��Ա���.Text = "43", "������Ա", txt��Ա���.Text = "44", "�������ͥ", _
                      txt��Ա���.Text = "45", "�ضȲм�", True, "����")
    txt����.Text = GetElemnetValue("PERSONNAME")
    txt�Ա�.Text = GetElemnetValue("SEX")
    txt�Ա�.Text = Switch(txt�Ա�.Text = "1", "��", txt�Ա�.Text = "2", "Ů", txt�Ա�.Text = "9", "����", True, txt�Ա�.Text)
    txt���֤��.Text = GetElemnetValue("PID")
    txt��������.Text = GetElemnetValue("BIRTHDAY")
    txt��λ����.Text = GetElemnetValue("DEPTCODE")
    txt��λ����.Text = GetElemnetValue("DEPTNAME")
    txt�ʻ����.Text = GetElemnetValue("ACCTBALANCE")
    '�ۼ���Ϣ
    txtסԺ����.Text = GetElemnetValue("HOSPTIMES")
    txt����.Text = GetElemnetValue("STARTFEE")
    txt��֧������.Text = GetElemnetValue("STARTFEEPAID")
    txt����ͳ���޶�.Text = GetElemnetValue("FUND1LMT")
    txtͳ��֧���ۼ�.Text = GetElemnetValue("FUND1PAID")
    txt���ͳ���޶�.Text = GetElemnetValue("FUND2LMT")
    txt���֧���ۼ�.Text = GetElemnetValue("FUND2PAID")
    txt��ͨ����ҽ�Ʋ����޶�.Text = GetElemnetValue("FUND3LMT")
    txt��ͨ����ҽ�Ʋ����ۼ�.Text = GetElemnetValue("FUND3PAID")
    txt��ͨ����ҽ�Ʋ����𸶱�׼.Text = GetElemnetValue("STARTFEE2STD")
    txt��ͨ����ҽ�Ʋ�������.Text = GetElemnetValue("STARTFEE2")
    txt��ͨ����ҽ�Ʋ�����ת��ʹ��.Text = GetElemnetValue("FUND75BALANCE")
    txt��ע.Text = GetElemnetValue("NOTE")
    txt������Ϣ.Text = GetElemnetValue("LOCKINFO")
    
    '20110812���ӱ���
    Txt������Ϣ.Text = gstr������Ϣ
     gstrSQL = " Select a.������� From zlgyyb.ҽ�����˵��� a,ҽ�����˹����� b" & _
   " Where a.ҽ����=b.ҽ���� and  b.��־=1 and b.����=[1] And a.ҽ����=[2]"
       Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�������", TYPE_������, CStr(Trim(txtҽ����.Text)))
        If rsTemp.RecordCount = 1 Then
        If rsTemp!������� <> "" And txt���������.Enabled = True Then
        txt���������.Text = rsTemp!�������
        End If
       End If
    'end
     cmdOK.Enabled = True
    If txt���������.Enabled = True Then txt���������.SetFocus
  
    If gblnLED Then
        zl9LedVoice.Speak "#26 " & Val(txt�ʻ����.Text)
    End If
End Sub
