VERSION 5.00
Begin VB.Form frmIdentify��������������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����¼��ҽ�����Ľ�����Ϣ"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   Icon            =   "frmIdentify���������������.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame frmDetail 
      Height          =   7110
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   9870
      Begin VB.TextBox txt����Ա�������� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   44
         Top             =   4215
         Width           =   1860
      End
      Begin VB.TextBox txt��ͨ���﹫��Ա�����ۼ� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   53
         Top             =   4680
         Width           =   1860
      End
      Begin VB.TextBox txt����Ա�����𸶱�׼ 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   47
         Top             =   4215
         Width           =   1860
      End
      Begin VB.TextBox txt������޶��Ա���� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   50
         Top             =   4680
         Width           =   1860
      End
      Begin VB.CommandButton cmd����Ա�������� 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify���������������.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   4215
         Width           =   300
      End
      Begin VB.CommandButton cmd��ͨ���﹫��Ա�����ۼ� 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify���������������.frx":006C
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   4680
         Width           =   300
      End
      Begin VB.CommandButton cmd����Ա�����𸶱�׼ 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify���������������.frx":00CC
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   4215
         Width           =   300
      End
      Begin VB.CommandButton cmd������޶��Ա���� 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify���������������.frx":012C
         Style           =   1  'Graphical
         TabIndex        =   51
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   4680
         Width           =   300
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1170
         Width           =   2220
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1215
         Width           =   2220
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   255
         Width           =   2220
      End
      Begin VB.TextBox txt����ID 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   2220
      End
      Begin VB.TextBox txt��ҳID 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   780
         Width           =   2220
      End
      Begin VB.TextBox txtҽ���� 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   315
         Width           =   2220
      End
      Begin VB.TextBox txtͳ��֧�� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   26
         Top             =   2820
         Width           =   1860
      End
      Begin VB.TextBox txtͳ���Ը� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   29
         Top             =   2820
         Width           =   1860
      End
      Begin VB.TextBox txtȫ�Ը� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   14
         Top             =   1905
         Width           =   1860
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   20
         Top             =   2370
         Width           =   1860
      End
      Begin VB.TextBox txt�����Ը� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   23
         Top             =   2370
         Width           =   1860
      End
      Begin VB.TextBox txt����Ա���� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   41
         Top             =   3750
         Width           =   1860
      End
      Begin VB.TextBox txt�����ܷ��� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   56
         Top             =   5160
         Width           =   1860
      End
      Begin VB.TextBox txt��ͳ�� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   32
         Top             =   3285
         Width           =   1860
      End
      Begin VB.TextBox txt�����Ը� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   38
         Top             =   3750
         Width           =   1860
      End
      Begin VB.TextBox txtҽ���ܷ��� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   59
         Top             =   5160
         Width           =   1860
      End
      Begin VB.TextBox txt���Ը� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   35
         Top             =   3285
         Width           =   1860
      End
      Begin VB.TextBox txt����˳��� 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   68
         Top             =   6135
         Width           =   1860
      End
      Begin VB.TextBox txt�������� 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   71
         Top             =   6135
         Width           =   1860
      End
      Begin VB.TextBox txtHIS�ܷ��� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   62
         Top             =   5640
         Width           =   1860
      End
      Begin VB.TextBox txt������ 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   65
         Top             =   5640
         Width           =   1860
      End
      Begin VB.CommandButton cmd���� 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify���������������.frx":018C
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   2370
         Width           =   300
      End
      Begin VB.CommandButton cmdȫ�Ը� 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify���������������.frx":01EC
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   1905
         Width           =   300
      End
      Begin VB.CommandButton cmd�����Ը� 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify���������������.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   3750
         Width           =   300
      End
      Begin VB.CommandButton cmd��ͳ�� 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify���������������.frx":02AC
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   3285
         Width           =   300
      End
      Begin VB.CommandButton cmd����Ա���� 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify���������������.frx":030C
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   3750
         Width           =   300
      End
      Begin VB.CommandButton cmd�������� 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify���������������.frx":036C
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   6135
         Width           =   300
      End
      Begin VB.CommandButton cmd������ 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify���������������.frx":03CC
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   5640
         Width           =   300
      End
      Begin VB.CommandButton cmd�����ܷ��� 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify���������������.frx":042C
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   5160
         Width           =   300
      End
      Begin VB.CommandButton cmdҽ���ܷ��� 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify���������������.frx":048C
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   5160
         Width           =   300
      End
      Begin VB.CommandButton cmd���Ը� 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify���������������.frx":04EC
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   3285
         Width           =   300
      End
      Begin VB.CommandButton cmdͳ���Ը� 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify���������������.frx":054C
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   2820
         Width           =   300
      End
      Begin VB.CommandButton cmd�����Ը� 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify���������������.frx":05AC
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   2370
         Width           =   300
      End
      Begin VB.CommandButton cmd�ҹ��Ը� 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify���������������.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   1905
         Width           =   300
      End
      Begin VB.CommandButton cmd����˳��� 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify���������������.frx":066C
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   6135
         Width           =   300
      End
      Begin VB.CommandButton cmdHIS�ܷ��� 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify���������������.frx":06CC
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   5640
         Width           =   300
      End
      Begin VB.CommandButton cmdͳ��֧�� 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify���������������.frx":072C
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   2820
         Width           =   300
      End
      Begin VB.TextBox txt�ҹ��Ը� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   17
         Top             =   1905
         Width           =   1860
      End
      Begin VB.TextBox txt�������˵�� 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7230
         TabIndex        =   77
         Top             =   6600
         Width           =   1860
      End
      Begin VB.TextBox txt������㷽ʽ 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2445
         TabIndex        =   74
         Top             =   6600
         Width           =   1860
      End
      Begin VB.CommandButton cmd�������˵�� 
         Height          =   300
         Left            =   9150
         Picture         =   "frmIdentify���������������.frx":078C
         Style           =   1  'Graphical
         TabIndex        =   78
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   6600
         Width           =   300
      End
      Begin VB.CommandButton cmd������㷽ʽ 
         Height          =   300
         Left            =   4365
         Picture         =   "frmIdentify���������������.frx":07EC
         Style           =   1  'Graphical
         TabIndex        =   75
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   6600
         Width           =   300
      End
      Begin VB.Label lab����Ա�������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����Ա��������"
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
         TabIndex        =   43
         Top             =   4260
         Width           =   1680
      End
      Begin VB.Label lab��ͨ���﹫��Ա�����ۼ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ͨ���﹫��Ա�����ۼ�"
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
         Left            =   4755
         TabIndex        =   52
         Top             =   4725
         Width           =   2310
      End
      Begin VB.Label lab����Ա�����𸶱�׼ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����Ա�����𸶱�׼"
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
         Left            =   5175
         TabIndex        =   46
         Top             =   4260
         Width           =   1890
      End
      Begin VB.Label lab������޶��Ա���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������޶��Ա����"
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
         Left            =   180
         TabIndex        =   49
         Top             =   4725
         Width           =   2100
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0080FFFF&
         X1              =   -150
         X2              =   15850
         Y1              =   1710
         Y2              =   1710
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   -150
         X2              =   15850
         Y1              =   1695
         Y2              =   1695
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
         Left            =   1860
         TabIndex        =   9
         Top             =   1215
         Width           =   420
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
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
         Left            =   6435
         TabIndex        =   11
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label Label4 
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
         Left            =   1860
         TabIndex        =   1
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
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
         Left            =   1650
         TabIndex        =   5
         Top             =   765
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҳID"
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
         Left            =   6435
         TabIndex        =   7
         Top             =   825
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����"
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
         Left            =   6435
         TabIndex        =   3
         Top             =   360
         Width           =   630
      End
      Begin VB.Label labͳ��֧�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ��֧��"
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
         Left            =   1440
         TabIndex        =   25
         Top             =   2865
         Width           =   840
      End
      Begin VB.Label labͳ���Ը� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ���Ը�"
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
         Left            =   6225
         TabIndex        =   28
         Top             =   2865
         Width           =   840
      End
      Begin VB.Label labȫ�Ը� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ȫ�Ը�"
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
         Left            =   1650
         TabIndex        =   13
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label lab���� 
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
         Left            =   1650
         TabIndex        =   19
         Top             =   2415
         Width           =   630
      End
      Begin VB.Label lab�����Ը� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ը�"
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
         Left            =   6225
         TabIndex        =   22
         Top             =   2415
         Width           =   840
      End
      Begin VB.Label lab�ҹ��Ը� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ҹ��Ը�"
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
         Left            =   6225
         TabIndex        =   16
         Top             =   1950
         Width           =   840
      End
      Begin VB.Label lab����Ա���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����Ա����"
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
         Left            =   6015
         TabIndex        =   40
         Top             =   3795
         Width           =   1050
      End
      Begin VB.Label lab�����ܷ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ܷ���"
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
         TabIndex        =   55
         Top             =   5205
         Width           =   1050
      End
      Begin VB.Label lab��ͳ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ͳ��"
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
         Left            =   1440
         TabIndex        =   31
         Top             =   3330
         Width           =   840
      End
      Begin VB.Label lab�����Ը� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ը�"
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
         Left            =   1440
         TabIndex        =   37
         Top             =   3795
         Width           =   840
      End
      Begin VB.Label labҽ���ܷ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ���ܷ���"
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
         Left            =   6015
         TabIndex        =   58
         Top             =   5175
         Width           =   1050
      End
      Begin VB.Label lab���Ը� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���Ը�"
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
         Left            =   6225
         TabIndex        =   34
         Top             =   3330
         Width           =   840
      End
      Begin VB.Label lab����˳��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����˳���"
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
         TabIndex        =   67
         Top             =   6180
         Width           =   1050
      End
      Begin VB.Label lab�������� 
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
         Left            =   6225
         TabIndex        =   70
         Top             =   6180
         Width           =   840
      End
      Begin VB.Label labHIS�ܷ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "HIS�ܷ���"
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
         Left            =   1335
         TabIndex        =   61
         Top             =   5685
         Width           =   945
      End
      Begin VB.Label lab������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������"
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
         Left            =   6225
         TabIndex        =   64
         Top             =   5685
         Width           =   840
      End
      Begin VB.Label lab�������˵�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������˵��"
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
         Left            =   5805
         TabIndex        =   76
         Top             =   6645
         Width           =   1260
      End
      Begin VB.Label lab������㷽ʽ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������㷽ʽ"
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
         TabIndex        =   73
         Top             =   6645
         Width           =   1260
      End
   End
   Begin VB.PictureBox P2 
      Height          =   495
      Left            =   1440
      Picture         =   "frmIdentify���������������.frx":084C
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   82
      Top             =   7335
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox P1 
      Height          =   495
      Left            =   75
      Picture         =   "frmIdentify���������������.frx":092A
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   81
      Top             =   7335
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
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
      Left            =   8610
      TabIndex        =   80
      Top             =   7365
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&0)"
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
      Left            =   7110
      TabIndex        =   79
      Top             =   7365
      Width           =   1335
   End
End
Attribute VB_Name = "frmIdentify���������������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_OkCancel            As Boolean

Public Property Get OkCancel() As Boolean
    OkCancel = m_OkCancel
End Property

Private Sub cmdCancel_Click()
    With g�������
        .blnYn = False
    End With
    Unload Me
End Sub

Private Sub cmdHIS�ܷ���_Click()
    txtHIS�ܷ���.Text = Format(Val(txtHIS�ܷ���.Text), "0.00")
    cmdHIS�ܷ���.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmdHIS�ܷ���.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txtHIS�ܷ���.SetFocus
    txtHIS�ܷ���.SelStart = 0
    txtHIS�ܷ���.SelLength = Len(txtHIS�ܷ���.Text)
End Sub

Private Sub cmdOK_Click()
    If Val(txtȫ�Ը�.Text) < 0 Then
        txtȫ�Ը�.SetFocus
        Exit Sub
    End If
    If Val(txt�ҹ��Ը�.Text) < 0 Then
        txt�ҹ��Ը�.SetFocus
        Exit Sub
    End If
    If Val(txt����.Text) < 0 Then
        txt����.SetFocus
        Exit Sub
    End If
    
    If Val(txt�����Ը�.Text) < 0 Then
        txt�����Ը�.SetFocus
        Exit Sub
    End If
    If Val(txtͳ��֧��.Text) < 0 Then
        txtͳ��֧��.SetFocus
        Exit Sub
    End If
    If Val(txtͳ���Ը�.Text) < 0 Then
        txtͳ���Ը�.SetFocus
        Exit Sub
    End If
    If Val(txt��ͳ��.Text) < 0 Then
        txt��ͳ��.SetFocus
        Exit Sub
    End If
    If Val(txt���Ը�.Text) < 0 Then
        txt���Ը�.SetFocus
        Exit Sub
    End If
    If Val(txt�����Ը�.Text) < 0 Then
        txt�����Ը�.SetFocus
        Exit Sub
    End If
    If Val(txtҽ���ܷ���.Text) < 0 Then
        txtҽ���ܷ���.SetFocus
        Exit Sub
    End If
    If Val(txt����Ա����.Text) < 0 Then
        txt����Ա����.SetFocus
        Exit Sub
    End If
    
    If Val(txt����Ա��������.Text) < 0 Then
        txt����Ա��������.SetFocus
        Exit Sub
    End If
    
    If Val(txt����Ա�����𸶱�׼.Text) < 0 Then
        txt����Ա�����𸶱�׼.SetFocus
        Exit Sub
    End If
    
    If Val(txt������޶��Ա����.Text) < 0 Then
        txt������޶��Ա����.SetFocus
        Exit Sub
    End If
    
    If Val(txt��ͨ���﹫��Ա�����ۼ�.Text) < 0 Then
        txt��ͨ���﹫��Ա�����ۼ�.SetFocus
        Exit Sub
    End If
    
    If Val(txt�����ܷ���.Text) < 0 Then
        txt�����ܷ���.SetFocus
        Exit Sub
    End If
    If Val(txtHIS�ܷ���.Text) < 0 Then
        MsgBox "HIS�ܷ��ñ������0", vbCritical, gstrSysName
        txtHIS�ܷ���.SetFocus
        Exit Sub
    End If
    If Len(txt������.Text) <= 0 Then
        MsgBox "�����Ų���Ϊ�գ�", vbCritical, gstrSysName
        txt������.SetFocus
        Exit Sub
    End If
    If Len(txt����˳���.Text) < 0 Then
        MsgBox "����˳��Ų���Ϊ�գ�", vbCritical, gstrSysName
        txt����˳���.SetFocus
        Exit Sub
    End If
    If Not IsDate(txt��������.Text) Then
        MsgBox "�������ڱ���Ϊ�������ͣ�", vbCritical, gstrSysName
        txt��������.SetFocus
        Exit Sub
    End If
    With g�������
        .blnYn = True
        .m_ȫ�Ը� = Val(txtȫ�Ը�.Text)
        .m_�ҹ��Ը� = Val(txt�ҹ��Ը�.Text)
        .m_���� = Val(txt����.Text)
        .m_�����Ը� = Val(txt�����Ը�.Text)
        .m_ͳ��֧�� = Val(txtͳ��֧��.Text)
        .m_ͳ���Ը� = Val(txtͳ���Ը�.Text)
        .m_��ͳ�� = Val(txt��ͳ��.Text)
        .m_���Ը� = Val(txt���Ը�.Text)
        .m_�����Ը� = Val(txt�����Ը�.Text)
        .m_ҽ���ܷ��� = Val(txtҽ���ܷ���.Text)
        .m_����Ա���� = Val(txt����Ա����.Text)
        .m_�����ܷ��� = Val(txt�����ܷ���.Text)
        .m_HIS�ܷ��� = Val(txtHIS�ܷ���.Text)
        .m_������ = txt������.Text
        .m_����˳��� = txt����˳���.Text
        .m_�������� = Format(txt��������.Text, "yyyy-mm-dd hh:mm:ss")
        .m_����Ա�����𸶱�׼ = Val(txt����Ա�����𸶱�׼.Text)
        .m_����Ա�������� = Val(txt����Ա��������.Text)
        .m_��ͨ���﹫��Ա�����ۼ� = Val(txt��ͨ���﹫��Ա�����ۼ�.Text)
        .m_������޶��Ա���� = Val(txt������޶��Ա����.Text)
        .m_������㷽ʽ = txt������㷽ʽ.Text
        .m_�������˵�� = txt�������˵��.Text
    End With
    Unload Me
End Sub

Private Sub cmd������޶��Ա����_Click()
    txt������޶��Ա����.Text = Format(Val(txt������޶��Ա����.Text), "0.00")
    cmd������޶��Ա����.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd������޶��Ա����.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txt������޶��Ա����.SetFocus
    txt������޶��Ա����.SelStart = 0
    txt������޶��Ա����.SelLength = Len(txt������޶��Ա����.Text)
End Sub

Private Sub cmd�����Ը�_Click()
    txt�����Ը�.Text = Format(Val(txt�����Ը�.Text), "0.00")
    cmd�����Ը�.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd�����Ը�.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txt�����Ը�.SetFocus
    txt�����Ը�.SelStart = 0
    txt�����Ը�.SelLength = Len(txt�����Ը�.Text)
End Sub

Private Sub cmd��ͳ��_Click()
    txt��ͳ��.Text = Format(Val(txt��ͳ��.Text), "0.00")
    cmd��ͳ��.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd��ͳ��.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txt��ͳ��.SetFocus
    txt��ͳ��.SelStart = 0
    txt��ͳ��.SelLength = Len(txt��ͳ��.Text)
End Sub

Private Sub cmd���Ը�_Click()
    txt���Ը�.Text = Format(Val(txt���Ը�.Text), "0.00")
    cmd���Ը�.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd���Ը�.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txt���Ը�.SetFocus
    txt���Ը�.SelStart = 0
    txt���Ը�.SelLength = Len(txt���Ը�.Text)
End Sub

Private Sub cmd����Ա����_Click()
    txt����Ա����.Text = Format(Val(txt����Ա����.Text), "0.00")
    cmd����Ա����.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd����Ա����.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txt����Ա����.SetFocus
    txt����Ա����.SelStart = 0
    txt����Ա����.SelLength = Len(txt����Ա����.Text)
End Sub

Private Sub cmd����Ա�����𸶱�׼_Click()
    txt����Ա�����𸶱�׼.Text = Format(Val(txt����Ա�����𸶱�׼.Text), "0.00")
    cmd����Ա�����𸶱�׼.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd����Ա�����𸶱�׼.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txt����Ա�����𸶱�׼.SetFocus
    txt����Ա�����𸶱�׼.SelStart = 0
    txt����Ա�����𸶱�׼.SelLength = Len(txt����Ա�����𸶱�׼.Text)
End Sub

Private Sub cmd����Ա��������_Click()
    txt����Ա��������.Text = Format(Val(txt����Ա��������.Text), "0.00")
    cmd����Ա��������.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd����Ա��������.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txt����Ա��������.SetFocus
    txt����Ա��������.SelStart = 0
    txt����Ա��������.SelLength = Len(txt����Ա��������.Text)
End Sub

Private Sub cmd�ҹ��Ը�_Click()
    txt�ҹ��Ը�.Text = Format(Val(txt�ҹ��Ը�.Text), "0.00")
    cmd�ҹ��Ը�.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd�ҹ��Ը�.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txt�ҹ��Ը�.SetFocus
    txt�ҹ��Ը�.SelStart = 0
    txt�ҹ��Ը�.SelLength = Len(txt�ҹ��Ը�.Text)
End Sub

Private Sub cmd�����Ը�_Click()
    txt�����Ը�.Text = Format(Val(txt�����Ը�.Text), "0.00")
    cmd�����Ը�.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd�����Ը�.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txt�����Ը�.SetFocus
    txt�����Ը�.SelStart = 0
    txt�����Ը�.SelLength = Len(txt�����Ը�.Text)
End Sub

Private Sub cmd������_Click()

    cmd������.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd������.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    
    txt������.SetFocus
    txt������.SelStart = 0
    txt������.SelLength = Len(txt������.Text)
End Sub

Private Sub cmd��������_Click()
    txt��������.Text = Format(txt��������.Text, "yyyy-mm-dd hh:mm:ss")
    cmd��������.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd��������.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    cmd��������.Picture = P1.Picture
    txt��������.SetFocus
    txt��������.SelStart = 0
    txt��������.SelLength = Len(txt��������.Text)
End Sub

Private Sub cmd�����ܷ���_Click()
    txt�����ܷ���.Text = Format(Val(txt�����ܷ���.Text), "0.00")
    cmd�����ܷ���.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd�����ܷ���.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    cmd�����ܷ���.Picture = P1.Picture
    txt�����ܷ���.SetFocus
    txt�����ܷ���.SelStart = 0
    txt�����ܷ���.SelLength = Len(txt�����ܷ���.Text)
End Sub

Private Sub cmd����˳���_Click()
    cmd����˳���.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd����˳���.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txt����˳���.SetFocus
    txt����˳���.SelStart = 0
    txt����˳���.SelLength = Len(txt����˳���.Text)
End Sub

Private Sub cmd��ͨ���﹫��Ա�����ۼ�_Click()
    txt��ͨ���﹫��Ա�����ۼ�.Text = Format(Val(txt��ͨ���﹫��Ա�����ۼ�.Text), "0.00")
    cmd��ͨ���﹫��Ա�����ۼ�.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd��ͨ���﹫��Ա�����ۼ�.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    cmd��ͨ���﹫��Ա�����ۼ�.Picture = P1.Picture
    txt��ͨ���﹫��Ա�����ۼ�.SetFocus
    txt��ͨ���﹫��Ա�����ۼ�.SelStart = 0
    txt��ͨ���﹫��Ա�����ۼ�.SelLength = Len(txt��ͨ���﹫��Ա�����ۼ�.Text)
End Sub

Private Sub cmd����_Click()
    txt����.Text = Format(Val(txt����.Text), "0.00")
    cmd����.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd����.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txt����.SetFocus
    txt����.SelStart = 0
    txt����.SelLength = Len(txt����.Text)
End Sub

Private Sub cmdȫ�Ը�_Click()

    txtȫ�Ը�.Text = Format(Val(txtȫ�Ը�.Text), "0.00")
    cmdȫ�Ը�.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmdȫ�Ը�.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txtȫ�Ը�.SetFocus
    txtȫ�Ը�.SelStart = 0
    txtȫ�Ը�.SelLength = Len(txtȫ�Ը�.Text)

End Sub

Private Sub cmd������㷽ʽ_Click()
    cmd������㷽ʽ.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd������㷽ʽ.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txt������㷽ʽ.SetFocus
    txt������㷽ʽ.SelStart = 0
    txt������㷽ʽ.SelLength = Len(txt������㷽ʽ.Text)
End Sub

Private Sub cmd�������˵��_Click()

    cmd�������˵��.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmd�������˵��.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
 
    cmd�������˵��.Picture = P1.Picture
    txt�������˵��.SetFocus
    txt�������˵��.SelStart = 0
    txt�������˵��.SelLength = Len(txt�������˵��.Text)
 
End Sub

Private Sub cmdͳ��֧��_Click()
    txtͳ��֧��.Text = Format(Val(txtͳ��֧��.Text), "0.00")
    cmdͳ��֧��.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmdͳ��֧��.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    cmdͳ��֧��.Picture = P1.Picture
    txtͳ��֧��.SetFocus
    txtͳ��֧��.SelStart = 0
    txtͳ��֧��.SelLength = Len(txtͳ��֧��.Text)
End Sub

Private Sub cmdͳ���Ը�_Click()
    txtͳ���Ը�.Text = Format(Val(txtͳ���Ը�.Text), "0.00")
    cmdͳ���Ը�.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmdͳ���Ը�.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txtͳ���Ը�.SetFocus
    txtͳ���Ը�.SelStart = 0
    txtͳ���Ը�.SelLength = Len(txtͳ���Ը�.Text)
End Sub

Private Sub cmdҽ���ܷ���_Click()
    txtҽ���ܷ���.Text = Format(Val(txtҽ���ܷ���.Text), "0.00")
    cmdҽ���ܷ���.Tag = IIf(cmd������.Tag = "1", "2", "1")
    cmdҽ���ܷ���.Picture = IIf(cmd������.Tag = "1", P2.Picture, P1.Picture)
    txtҽ���ܷ���.SetFocus
    txtҽ���ܷ���.SelStart = 0
    txtҽ���ܷ���.SelLength = Len(txtҽ���ܷ���.Text)
End Sub

Private Sub Form_Load()
    txt����.Text = g�������.str����
    txtҽ����.Text = g�������.strҽ����
    txt����ID.Text = g�������.lng����ID
    txt��ҳID.Text = g�������.lng��ҳID
    txt����.Text = g�������.str����
    txtסԺ��.Text = g�������.strסԺ��
    
End Sub

Private Sub txtHIS�ܷ���_LostFocus()
    txtHIS�ܷ���.Text = Format(Val(txtHIS�ܷ���.Text), "0.00")
End Sub

Private Sub txt������޶��Ա����_LostFocus()
    txt������޶��Ա����.Text = Format(Val(txt������޶��Ա����.Text), "0.00")
End Sub

Private Sub txt�����Ը�_LostFocus()
    txt�����Ը�.Text = Format(Val(txt�����Ը�.Text), "0.00")
End Sub

Private Sub txt��ͳ��_LostFocus()
    txt��ͳ��.Text = Format(Val(txt��ͳ��.Text), "0.00")
End Sub

Private Sub txt���Ը�_LostFocus()
    txt���Ը�.Text = Format(Val(txt���Ը�.Text), "0.00")
End Sub

Private Sub txt����Ա����_LostFocus()
    txt����Ա����.Text = Format(Val(txt����Ա����.Text), "0.00")
End Sub

Private Sub txt����Ա�����𸶱�׼_LostFocus()
    txt����Ա�����𸶱�׼.Text = Format(Val(txt����Ա�����𸶱�׼.Text), "0.00")
End Sub

Private Sub txt����Ա��������_LostFocus()
    txt����Ա��������.Text = Format(Val(txt����Ա��������.Text), "0.00")
End Sub

Private Sub txt�ҹ��Ը�_LostFocus()
    txt�ҹ��Ը�.Text = Format(Val(txt�ҹ��Ը�.Text), "0.00")
End Sub

Private Sub txt�����Ը�_LostFocus()
    txt�����Ը�.Text = Format(Val(txt�����Ը�.Text), "0.00")
End Sub

Private Sub txt�����ܷ���_LostFocus()
    txt�����ܷ���.Text = Format(Val(txt�����ܷ���.Text), "0.00")
End Sub

Private Sub txt��ͨ���﹫��Ա�����ۼ�_LostFocus()
    txt��ͨ���﹫��Ա�����ۼ�.Text = Format(Val(txt��ͨ���﹫��Ա�����ۼ�.Text), "0.00")
End Sub

Private Sub txt����_LostFocus()
    txt����.Text = Format(Val(txt����.Text), "0.00")
End Sub

Private Sub txtȫ�Ը�_LostFocus()
    txtȫ�Ը�.Text = Format(Val(txtȫ�Ը�.Text), "0.00")
End Sub

Private Sub txtͳ��֧��_LostFocus()
    txtͳ��֧��.Text = Format(Val(txtͳ��֧��.Text), "0.00")
End Sub

Private Sub txtͳ���Ը�_LostFocus()
    txtͳ���Ը�.Text = Format(Val(txtͳ���Ը�.Text), "0.00")
End Sub

Private Sub txtҽ���ܷ���_LostFocus()
    txtҽ���ܷ���.Text = Format(Val(txtҽ���ܷ���.Text), "0.00")
End Sub

