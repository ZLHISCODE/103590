VERSION 5.00
Object = "{D01C2596-4FE0-4EA9-9EE8-D97BE62A1165}#4.0#0"; "ZlPatiAddress.ocx"
Begin VB.Form frmEditPati 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ϣ�޸�"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   Icon            =   "frmEditPati.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   480
      TabIndex        =   62
      Top             =   8565
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7770
      TabIndex        =   61
      Top             =   8565
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6495
      TabIndex        =   60
      Top             =   8565
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   90
      TabIndex        =   63
      Top             =   0
      Width           =   8955
      Begin VB.CheckBox chk��� 
         Caption         =   "�Ƿ����"
         Height          =   195
         Left            =   7605
         TabIndex        =   6
         Top             =   660
         Width           =   1020
      End
      Begin VB.TextBox txt���� 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txt��Ժʱ�� 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   7215
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1530
      End
      Begin VB.TextBox txt�ȼ� 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   960
         Width           =   3300
      End
      Begin VB.TextBox txt��λ 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   600
         Width           =   1380
      End
      Begin VB.TextBox txt���� 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   3300
      End
      Begin VB.TextBox txt�Ա� 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   660
      End
      Begin VB.TextBox txt���� 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   3255
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txtסԺ�� 
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ȼ�"
         Height          =   180
         Left            =   4965
         TabIndex        =   72
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   6450
         TabIndex        =   71
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ȼ�"
         Height          =   180
         Left            =   390
         TabIndex        =   70
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   5325
         TabIndex        =   69
         Top             =   660
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         Height          =   180
         Left            =   390
         TabIndex        =   68
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   5325
         TabIndex        =   67
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2850
         TabIndex        =   66
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   570
         TabIndex        =   65
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   7095
      Left            =   90
      TabIndex        =   64
      Top             =   1320
      Width           =   8955
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   7290
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdת�� 
         Caption         =   "��"
         Height          =   240
         Left            =   8415
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   4700
         Width           =   300
      End
      Begin VB.TextBox txtת�� 
         Height          =   300
         Left            =   4200
         TabIndex        =   114
         Top             =   4680
         Width           =   4550
      End
      Begin VB.ComboBox cboIDNumber 
         Height          =   300
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   975
         Width           =   1695
      End
      Begin VB.TextBox txtLinkManInfo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   7275
         MaxLength       =   90
         TabIndex        =   40
         Top             =   3945
         Width           =   1455
      End
      Begin VB.ComboBox cbo��Ժ��ʽ 
         Height          =   300
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   4680
         Width           =   1860
      End
      Begin VB.PictureBox picDoctor 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   240
         ScaleHeight     =   975
         ScaleWidth      =   8655
         TabIndex        =   102
         Top             =   5040
         Width           =   8655
         Begin VB.Frame Frame2 
            Height          =   30
            Left            =   -120
            TabIndex        =   111
            Top             =   0
            Width           =   8685
         End
         Begin VB.Frame Frame4 
            Height          =   30
            Left            =   -120
            TabIndex        =   110
            Top             =   840
            Width           =   8685
         End
         Begin VB.Frame Frame5 
            Height          =   45
            Left            =   0
            TabIndex        =   109
            Top             =   -120
            Width           =   8085
         End
         Begin VB.ComboBox cbo����ҽʦ 
            Height          =   300
            Left            =   915
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   480
            Width           =   1500
         End
         Begin VB.ComboBox cbo���λ�ʿ 
            Height          =   300
            Left            =   3645
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   480
            Width           =   1440
         End
         Begin VB.ComboBox cbo����ҽʦ 
            Height          =   300
            Left            =   7080
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   480
            Width           =   1455
         End
         Begin VB.ComboBox cbo����ҽʦ 
            Height          =   300
            Left            =   7065
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   60
            Width           =   1455
         End
         Begin VB.ComboBox cboסԺҽʦ 
            Height          =   300
            Left            =   3645
            Style           =   2  'Dropdown List
            TabIndex        =   45
            ToolTipText     =   "����ҽʦ"
            Top             =   60
            Width           =   1455
         End
         Begin VB.TextBox txtҽ��С�� 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   915
            MaxLength       =   6
            TabIndex        =   44
            Top             =   60
            Width           =   1500
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҽʦ"
            Height          =   180
            Left            =   165
            TabIndex        =   108
            Top             =   540
            Width           =   720
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���λ�ʿ"
            Height          =   180
            Left            =   2865
            TabIndex        =   107
            Top             =   540
            Width           =   720
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺҽʦ"
            Height          =   180
            Left            =   2850
            TabIndex        =   106
            Top             =   120
            Width           =   720
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҽʦ"
            Height          =   180
            Left            =   6285
            TabIndex        =   105
            Top             =   120
            Width           =   720
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(������)ҽʦ"
            Height          =   180
            Left            =   5565
            TabIndex        =   104
            Top             =   540
            Width           =   1440
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ��С��"
            Height          =   180
            Left            =   165
            TabIndex        =   103
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.TextBox txt��ϵ�����֤�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1155
         MaxLength       =   18
         TabIndex        =   38
         Top             =   3945
         Width           =   4005
      End
      Begin VB.CommandButton cmd���ڵ�ַ 
         Caption         =   "��"
         Height          =   240
         Left            =   5610
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F3"
         Top             =   2430
         Width           =   285
      End
      Begin VB.TextBox txt���ڵ�ַ�ʱ� 
         Height          =   300
         Left            =   7290
         MaxLength       =   6
         TabIndex        =   31
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   240
         Left            =   8460
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F3"
         Top             =   1730
         Width           =   255
      End
      Begin VB.TextBox txt��ע 
         Height          =   300
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   52
         Top             =   6705
         Width           =   7590
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   240
         Left            =   8460
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F3"
         Top             =   1020
         Width           =   255
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   300
         Left            =   7290
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   3165
         Width           =   1215
      End
      Begin VB.ComboBox cbo��Ժ���� 
         Height          =   300
         Left            =   7290
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox cboְҵ 
         Height          =   300
         Left            =   7290
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cbo��ϵ�˹�ϵ 
         Height          =   300
         Left            =   5730
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   3945
         Width           =   1455
      End
      Begin VB.TextBox txt��ͥ��ַ�ʱ� 
         Height          =   300
         Left            =   7290
         MaxLength       =   6
         TabIndex        =   28
         Top             =   2050
         Width           =   1455
      End
      Begin VB.TextBox txt���֤�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1155
         TabIndex        =   16
         Top             =   975
         Width           =   2085
      End
      Begin VB.CommandButton cmd�����ص� 
         Caption         =   "��"
         Height          =   240
         Left            =   5595
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F3"
         Top             =   1350
         Width           =   300
      End
      Begin VB.ComboBox cbo���䵥λ 
         Height          =   300
         Left            =   6000
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   580
      End
      Begin VB.TextBox txt��ҽ��� 
         Height          =   300
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   51
         Top             =   6360
         Width           =   6165
      End
      Begin VB.CommandButton cmd��ϵ�˵�ַ 
         Caption         =   "��"
         Height          =   240
         Left            =   8415
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F3"
         Top             =   4350
         Width           =   300
      End
      Begin VB.TextBox txt��Ժ��� 
         Height          =   300
         Left            =   1155
         MaxLength       =   200
         TabIndex        =   50
         Top             =   6015
         Width           =   6165
      End
      Begin VB.TextBox txt��ϵ�˵�ַ 
         Height          =   300
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   41
         Top             =   4320
         Width           =   7590
      End
      Begin VB.ComboBox cbo�ѱ� 
         Height          =   300
         Left            =   3195
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1290
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   840
      End
      Begin VB.ComboBox cbo����״�� 
         Height          =   300
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   630
         Width           =   1425
      End
      Begin VB.TextBox txt��ϵ�˵绰 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4035
         MaxLength       =   20
         TabIndex        =   37
         Top             =   3585
         Width           =   1890
      End
      Begin VB.TextBox txt��ϵ������ 
         Height          =   300
         Left            =   1155
         MaxLength       =   64
         TabIndex        =   36
         Top             =   3585
         Width           =   1890
      End
      Begin VB.TextBox txt��ͥ�绰 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   27
         Top             =   2050
         Width           =   4785
      End
      Begin VB.CommandButton cmd��ͥ��ַ 
         Caption         =   "��"
         Height          =   240
         Left            =   5610
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F3"
         Top             =   1730
         Width           =   285
      End
      Begin VB.TextBox txt��ͥ��ַ 
         Height          =   300
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   23
         Top             =   1700
         Width           =   4785
      End
      Begin VB.TextBox txt��λ�绰 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4035
         MaxLength       =   20
         TabIndex        =   34
         Top             =   3165
         Width           =   1890
      End
      Begin VB.CommandButton cmd��λ��ַ 
         Caption         =   "��"
         Height          =   240
         Left            =   8415
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F3"
         Top             =   2820
         Width           =   285
      End
      Begin VB.TextBox txt��λ��ַ 
         Height          =   300
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   32
         Top             =   2790
         Width           =   7575
      End
      Begin VB.ComboBox cboѧ�� 
         Height          =   300
         ItemData        =   "frmEditPati.frx":058A
         Left            =   4515
         List            =   "frmEditPati.frx":058C
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   630
         Width           =   1410
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1425
      End
      Begin VB.TextBox txt��λ�ʱ� 
         Height          =   300
         Left            =   1155
         MaxLength       =   6
         TabIndex        =   33
         Top             =   3165
         Width           =   1890
      End
      Begin VB.CheckBox chk����Ժ 
         Caption         =   "����Ժ"
         Height          =   255
         Left            =   5070
         TabIndex        =   18
         ToolTipText     =   "�ٴ���ס��ͬ���ƿ�Ŀ������ٴ�����"
         Top             =   998
         Width           =   855
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   7290
         MaxLength       =   50
         TabIndex        =   19
         Top             =   990
         Width           =   1455
      End
      Begin VB.TextBox txt���ڵ�ַ 
         Height          =   300
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   29
         Top             =   2400
         Width           =   4785
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   300
         Index           =   5
         Left            =   1155
         TabIndex        =   42
         Tag             =   "��ϵ�˵�ַ"
         Top             =   4320
         Visible         =   0   'False
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   100
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   300
         Index           =   4
         Left            =   1140
         TabIndex        =   30
         Tag             =   "���ڵ�ַ"
         Top             =   2400
         Visible         =   0   'False
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   100
      End
      Begin VB.TextBox txt�����ص� 
         Height          =   300
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   20
         Top             =   1320
         Width           =   4785
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   300
         Index           =   1
         Left            =   1140
         TabIndex        =   21
         Tag             =   "�����ص�"
         Top             =   1320
         Visible         =   0   'False
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Items           =   3
         MaxLength       =   100
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   6525
         MaxLength       =   50
         TabIndex        =   25
         Top             =   1700
         Width           =   2220
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   300
         Index           =   2
         Left            =   6525
         TabIndex        =   26
         Tag             =   "����"
         Top             =   1700
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Items           =   2
         MaxLength       =   100
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   300
         Index           =   3
         Left            =   1140
         TabIndex        =   24
         Tag             =   "��סַ"
         Top             =   1700
         Visible         =   0   'False
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   100
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   6840
         TabIndex        =   117
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblInFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ת��"
         Height          =   180
         Left            =   3720
         TabIndex        =   113
         Top             =   4740
         Width           =   360
      End
      Begin VB.Label lbl��Ժ��ʽ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ��ʽ"
         Height          =   180
         Left            =   360
         TabIndex        =   112
         Top             =   4740
         Width           =   720
      End
      Begin VB.Label lbl��ϵ�����֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�����֤"
         Height          =   180
         Left            =   45
         TabIndex        =   101
         Top             =   4005
         Width           =   1080
      End
      Begin VB.Label lbl���ڵ�ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڵ�ַ"
         Height          =   180
         Left            =   405
         TabIndex        =   100
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label lbl���ڵ�ַ�ʱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڵ�ַ�ʱ�"
         Height          =   180
         Left            =   6135
         TabIndex        =   99
         Top             =   2460
         Width           =   1080
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   6135
         TabIndex        =   98
         Top             =   1760
         Width           =   360
      End
      Begin VB.Label lbl��ע 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע"
         Height          =   180
         Left            =   765
         TabIndex        =   97
         Top             =   6765
         Width           =   360
      End
      Begin VB.Label lblPatiColor 
         BackColor       =   &H80000012&
         Height          =   255
         Left            =   8520
         TabIndex        =   96
         Top             =   3180
         Width           =   225
      End
      Begin VB.Label lblPatiType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   6495
         TabIndex        =   95
         Top             =   3225
         Width           =   720
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ����"
         Height          =   180
         Left            =   6495
         TabIndex        =   94
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   6855
         TabIndex        =   93
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label lbl�����ص� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ص�"
         Height          =   180
         Left            =   405
         TabIndex        =   92
         Top             =   1395
         Width           =   720
      End
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   180
         Left            =   405
         TabIndex        =   91
         Top             =   1035
         Width           =   720
      End
      Begin VB.Label lbl��ҽ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҽ���"
         Height          =   180
         Left            =   405
         TabIndex        =   90
         Top             =   6420
         Width           =   720
      End
      Begin VB.Label lbl��Ժ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   405
         TabIndex        =   89
         Top             =   6075
         Width           =   720
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵�ַ"
         Height          =   180
         Left            =   225
         TabIndex        =   88
         Top             =   4380
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
         Height          =   180
         Left            =   2730
         TabIndex        =   87
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4695
         TabIndex        =   86
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��"
         Height          =   180
         Left            =   405
         TabIndex        =   85
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�绰"
         Height          =   180
         Left            =   3630
         TabIndex        =   84
         Top             =   3645
         Width           =   360
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ"
         Height          =   180
         Left            =   5295
         TabIndex        =   83
         Top             =   4005
         Width           =   360
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ��"
         Height          =   180
         Left            =   585
         TabIndex        =   82
         Top             =   3645
         Width           =   540
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ�ʱ�"
         Height          =   180
         Left            =   6135
         TabIndex        =   81
         Top             =   2115
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ�绰"
         Height          =   180
         Left            =   405
         TabIndex        =   80
         Top             =   2110
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��סַ"
         Height          =   180
         Left            =   585
         TabIndex        =   79
         Top             =   1755
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�绰"
         Height          =   180
         Left            =   3630
         TabIndex        =   77
         Top             =   3225
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������λ"
         Height          =   180
         Left            =   405
         TabIndex        =   76
         Top             =   2850
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѧ��"
         Height          =   180
         Left            =   4050
         TabIndex        =   75
         Top             =   690
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         Height          =   180
         Left            =   405
         TabIndex        =   74
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ"
         Height          =   180
         Left            =   6855
         TabIndex        =   73
         Top             =   675
         Width           =   360
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʱ�"
         Height          =   180
         Left            =   405
         TabIndex        =   78
         Top             =   3225
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmEditPati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mstrPrivs As String
Public mlngUnit As Long, mstrUnit As String
Public mlng����ID As Long, mlng��ҳID As Long

Private mrsPati As ADODB.Recordset
Private mfrmParent As Object
Private mstrPatiPlus    As String     '�ӱ���Ϣ:��Ϣ��1:��Ϣֵ1,��Ϣ��2:��Ϣֵ2
Private mblnEMPI As Boolean       'T-����EMPIƽ̨����,F-δ����EMPIƽ̨����
Private mstrBirthDay As String

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cbo��������_Click()
    If cbo��������.ListCount > 0 And cbo��������.ListIndex <> -1 Then
        lblPatiColor.BackColor = zlDatabase.GetPatiColor(zlCommFun.GetNeedName(cbo��������.Text))
    End If
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    Call cbo.SetIndex(cbo��������.hWnd, cbo.MatchIndex(cbo��������.hWnd, KeyAscii))
End Sub

Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo�ѱ�.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cbo�ѱ�.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cbo�ѱ�.ListIndex = lngIdx
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cbo����.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
End Sub

Private Sub cbo����״��_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����״��.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cbo����״��.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cbo����״��.ListIndex = lngIdx
End Sub

Private Sub cbo��ϵ�˹�ϵ_Click()
    If zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.Text) = "����" Then
        txtLinkManInfo.Enabled = True: txtLinkManInfo.BackColor = &H80000005
    Else
        txtLinkManInfo.Enabled = False: txtLinkManInfo.Text = "": txtLinkManInfo.BackColor = &HE0E0E0
    End If
End Sub

Private Sub cbo���䵥λ_LostFocus()
    If Not CheckOldData(txt����, cbo���䵥λ) Then Exit Sub
End Sub

Private Sub cbo��Ժ��ʽ_Click()
    If zlCommFun.GetNeedName(cbo��Ժ��ʽ.Text) = "ת��" Then
        txtת��.Enabled = True: txtת��.BackColor = &H80000005
        cmdת��.Enabled = True: cmdת��.BackColor = &H80000005
        lblInFrom.Enabled = True
    Else
        cmdת��.Enabled = False: cmdת��.BackColor = &HE0E0E0
        txtת��.Enabled = False: txtת��.Text = "": txtת��.BackColor = &HE0E0E0
        lblInFrom.Enabled = False
    End If
End Sub

Private Sub cbo����ҽʦ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo����ҽʦ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo����ҽʦ.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo����ҽʦ.ListIndex = lngIdx
    ElseIf cbo����ҽʦ.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cboסԺҽʦ_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If cboסԺҽʦ.Text = "����..." Then
        Set rsTmp = GetSelectPersonal("ҽ��", "����ҽʦ,������ҽʦ,����ҽʦ,ҽʦ,ҽʿ", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cboסԺҽʦ.ListCount - 1
                If cboסԺҽʦ.List(i) = rsTmp!���� & "-" & rsTmp!���� Then
                    cboסԺҽʦ.ListIndex = i: Exit Sub
                End If
            Next
            cboסԺҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����, cboסԺҽʦ.ListCount - 1
            cboסԺҽʦ.ListIndex = cboסԺҽʦ.NewIndex
            cboסԺҽʦ.ItemData(cboסԺҽʦ.NewIndex) = rsTmp!�ϼ�ID
        Else
            cboסԺҽʦ.ListIndex = -1
        End If
    End If
End Sub
Private Sub cbo����ҽʦ_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If cbo����ҽʦ.Text = "����..." Then
        Set rsTmp = GetSelectPersonal("ҽ��", "����ҽʦ,������ҽʦ,����ҽʦ", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo����ҽʦ.ListCount - 1
                If cbo����ҽʦ.List(i) = rsTmp!���� & "-" & rsTmp!���� Then
                    cbo����ҽʦ.ListIndex = i: Exit Sub
                End If
            Next
            cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����ҽʦ.ListCount - 1
            cbo����ҽʦ.ListIndex = cbo����ҽʦ.NewIndex
            cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
        Else
            cbo����ҽʦ.ListIndex = -1
        End If
    End If
End Sub
Private Sub cbo����ҽʦ_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If cbo����ҽʦ.Text = "����..." Then
        Set rsTmp = GetSelectPersonal("ҽ��", "����ҽʦ,������ҽʦ", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo����ҽʦ.ListCount - 1
                If cbo����ҽʦ.List(i) = rsTmp!���� & "-" & rsTmp!���� Then
                    cbo����ҽʦ.ListIndex = i: Exit Sub
                End If
            Next
            cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����ҽʦ.ListCount - 1
            cbo����ҽʦ.ListIndex = cbo����ҽʦ.NewIndex
            cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
        Else
            cbo����ҽʦ.ListIndex = -1
        End If
    End If
End Sub

Private Sub cbo���λ�ʿ_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If cbo���λ�ʿ.Text = "����..." Then
        Set rsTmp = GetSelectPersonal("��ʿ", "", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo���λ�ʿ.ListCount - 1
                If cbo���λ�ʿ.List(i) = rsTmp!���� & "-" & rsTmp!���� Then
                    cbo���λ�ʿ.ListIndex = i: Exit Sub
                End If
            Next
            cbo���λ�ʿ.AddItem rsTmp!���� & "-" & rsTmp!����, cbo���λ�ʿ.ListCount - 1
            cbo���λ�ʿ.ListIndex = cbo���λ�ʿ.NewIndex
            cbo���λ�ʿ.ItemData(cbo���λ�ʿ.NewIndex) = rsTmp!ID
        Else
            cbo���λ�ʿ.ListIndex = -1
        End If
    End If
End Sub

Private Sub cbo��ϵ�˹�ϵ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo��ϵ�˹�ϵ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cbo��ϵ�˹�ϵ.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cbo��ϵ�˹�ϵ.ListIndex = lngIdx
End Sub

Private Sub cbo����ҽʦ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����ҽʦ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cbo����ҽʦ.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cbo����ҽʦ.ListIndex = lngIdx
End Sub

Private Sub cboѧ��_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboѧ��.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cboѧ��.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cboѧ��.ListIndex = lngIdx
End Sub

Private Sub cbo���λ�ʿ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo���λ�ʿ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cbo���λ�ʿ.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cbo���λ�ʿ.ListIndex = lngIdx
End Sub

Private Sub cboְҵ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboְҵ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cboְҵ.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cboְҵ.ListIndex = lngIdx
End Sub

Private Sub cboסԺҽʦ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboסԺҽʦ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cboסԺҽʦ.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cboסԺҽʦ.ListIndex = lngIdx
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, "frmHosReg"
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, strSQL_Recalc As String, blnTrans As Boolean
    Dim lng��ҽ����ID As Long, lng��ҽ����ID As Long
    Dim lng��ҽ���ID As Long, lng��ҽ���ID As Long, str���� As String
    Dim arrSQL() As String, i As Integer
    Dim arrTmp As Variant
    Dim lngTmp As Long
    Dim strBeginDate As String, strEndDate As String
    Dim str�Ա� As String, strAge As String, str�������� As String, strErrInfo As String
    Dim bln������Ϣ���� As Boolean, blnMod As Boolean
    Dim strMsg As String
    
    If cbo�ѱ�.ListIndex = -1 Then
        MsgBox "��ȷ�����˵ķѱ�", vbInformation, gstrSysName
        cbo�ѱ�.SetFocus: Exit Sub
    End If
    
    If cbo����.ListIndex = -1 Then
        MsgBox "��ȷ�����˵Ĺ�����", vbInformation, gstrSysName
        If CanFocus(cbo����) = True Then  cbo����.SetFocus: Exit Sub
    End If
    
    '�ʱ���
    If ((Not IsNumeric(txt���ڵ�ַ�ʱ�.Text)) Or Len(txt���ڵ�ַ�ʱ�.Text) > 6 Or InStr(txt���ڵ�ַ�ʱ�.Text, ".") > 0) And txt���ڵ�ַ�ʱ�.Text <> "" Then
        MsgBox "�ʱ��ʽ����,��������ȷ���ʱ�!" & vbCrLf & "����ȷ�ʱ��ʽΪ��λ�����ֱ��롿", vbInformation, gstrSysName
        If CanFocus(txt���ڵ�ַ�ʱ�) = True Then  txt���ڵ�ַ�ʱ�.SetFocus: Exit Sub
    End If
    If ((Not IsNumeric(txt��λ�ʱ�.Text)) Or Len(txt��λ�ʱ�.Text) > 6 Or InStr(txt��λ�ʱ�.Text, ".") > 0) And txt��λ�ʱ�.Text <> "" Then
        MsgBox "�ʱ��ʽ����,��������ȷ���ʱ�!" & vbCrLf & "����ȷ�ʱ��ʽΪ��λ�����ֱ��롿", vbInformation, gstrSysName
        If CanFocus(txt��λ�ʱ�) = True Then txt��λ�ʱ�.SetFocus: Exit Sub
    End If
    If ((Not IsNumeric(txt��ͥ��ַ�ʱ�.Text)) Or Len(txt��ͥ��ַ�ʱ�.Text) > 6 Or InStr(txt��ͥ��ַ�ʱ�.Text, ".") > 0) And txt��ͥ��ַ�ʱ�.Text <> "" Then
        MsgBox "�ʱ��ʽ����,��������ȷ���ʱ�!" & vbCrLf & "����ȷ�ʱ��ʽΪ��λ�����ֱ��롿", vbInformation, gstrSysName
        If CanFocus(txt��ͥ��ַ�ʱ�) = True Then  txt��ͥ��ַ�ʱ�.SetFocus: Exit Sub
    End If
    '��ϵ�˼��
    If Trim(txt��ϵ������.Text) = "" And (cbo��ϵ�˹�ϵ.ListIndex >= 0 Or Trim(txt��ϵ�˵绰.Text) <> "" Or Trim(txt��ϵ�˵�ַ.Text) <> "" Or Trim(txt��ϵ�����֤��.Text) <> "") Then
        MsgBox "����¼����ϵ������!", vbInformation, gstrSysName
        If CanFocus(txt��ϵ������) = True Then txt��ϵ������.SetFocus: Exit Sub
    End If
    '�ѱ����ÿ���
    If Not Check�ѱ����ÿ���(zlCommFun.GetNeedName(cbo�ѱ�.Text), Val(txt����.Tag)) Then
        MsgBox "��ǰ�ѱ�Բ��˿��Ҳ�����,������ѡ��ѱ�!", vbInformation, gstrSysName
        cbo�ѱ�.SetFocus: Exit Sub
    End If
    
    '��Ժ���
    If Not CheckLen(txt��Ժ���, txt��Ժ���.MaxLength) Then Exit Sub
    If Not CheckLen(txt��ҽ���, txt��ҽ���.MaxLength) Then Exit Sub
    If Not IsNull(mrsPati!����) Then
        If gclsInsure.GetCapability(support����¼��������, mlng����ID, mrsPati!����) Then
            If txt��Ժ���.Text = "" Then
                MsgBox "����д�ò��˵���Ժ��ϣ�", vbInformation, gstrSysName
                txt��Ժ���.SetFocus: Exit Sub
            End If
        End If
    End If
    
    
    If Not CheckTextLength("����", txt����) Then Exit Sub
    If Not CheckOldData(txt����, cbo���䵥λ) Then Exit Sub
    If Not CheckLen(txt�����ص�, txt�����ص�.MaxLength) Then Exit Sub
    If Not CheckLen(txt���ڵ�ַ, txt���ڵ�ַ.MaxLength) Then Exit Sub
    If Not CheckLen(txt��ͥ��ַ, txt��ͥ��ַ.MaxLength) Then Exit Sub
    If Not CheckLen(txt��ϵ������, txt��ϵ������.MaxLength) Then Exit Sub
    If Not CheckLen(txt��ϵ�˵�ַ, txt��ϵ�˵�ַ.MaxLength) Then Exit Sub
    If Not CheckLen(txt��λ��ַ, txt��λ��ַ.MaxLength) Then Exit Sub
    If Trim(zlCommFun.GetNeedName(cbo����.Text)) = "�й�" Then
        If Not CheckLen(txt���֤��, 18) Then Exit Sub
    End If
    If Not CheckLen(txt��ϵ�����֤��, 18) Then Exit Sub
    If Not CheckLen(txtLinkManInfo, 100) Then Exit Sub
    
    If zlStr.NeedName(cbo��Ժ��ʽ.Text) = "ת��" Then
        If Not zlControl.TxtCheckInput(txtת��, "ת��", 100) Then Exit Sub
    End If
    
    '����27351 by lesfeng 2010-01-12
    If Not CheckLen(txt��ע, txt��ע.MaxLength) Then Exit Sub
        
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    '--46119,������,2012-08-16,�����������֤�Ա��Ƿ�Ͳ����Ա�һ��
    '--81012,��ΰ��,2014-12-23,���֤��Ϣͬ�����Ա�һ��ʱ���������Ա���С�������Ϣ������Ȩ��ʱ�����Զ����������Ա�
    If Trim(zlCommFun.GetNeedName(cbo����.Text)) = "�й�" Then
        lngTmp = LenB(StrConv(Trim(txt���֤��.Text), vbFromUnicode))
        If lngTmp > 0 Then
            If CreatePublicPatient() Then
                If gobjPublicPatient.CheckPatiIdcard(Trim(txt���֤��.Text), str��������, strAge, str�Ա�, strErrInfo, CDate(txt��Ժʱ��.Text)) Then
                    '���޻�����Ϣ����Ȩ��
                    bln������Ϣ���� = InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";������Ϣ����;") > 0
                    If Format(mstrBirthDay, "HH:MM") <> "00:00" Then
                        str�������� = str�������� & " " & Format(mstrBirthDay, "HH:MM")
                    End If
                    '����
                    If Not str���� Like "Լ*" Or str���� <> "����" Then
                        strMsg = ""
                        If str���� <> strAge Then
                            strMsg = "���֤�����е�����[" & strAge & "]" & "�Ͳ�������[" & str���� & "]��һ��"
                            If str���� Like "*Сʱ*����" Or str���� Like "*����" Or str���� Like "*��*Сʱ" Or str���� Like "*Сʱ" Then
                                strAge = str����
                            End If
                        ElseIf InStr(txt�Ա�.Text, str�Ա�) = 0 Then '�Ա�
                            strMsg = "���֤�����е��Ա�[" & str�Ա� & "]�Ͳ����Ա�[" & txt�Ա�.Text & "]��һ��"
                        End If
                    End If
                    If strMsg <> "" Then
                        If MsgBox(strMsg & ",�Ƿ������" & vbCrLf & IIf(bln������Ϣ����, "ѡ���ǡ�,�����֤����Ϣ�滻���˵���Ϣ��", ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            If CanFocus(txt���֤��) = True Then txt���֤��.SetFocus: Exit Sub
                        Else
                            blnMod = True
                        End If
                    End If
                Else
                    MsgBox strErrInfo, vbInformation + vbOKOnly, gstrSysName
                    If CanFocus(txt���֤��) = True Then txt���֤��.SetFocus: Exit Sub
                End If
            End If
        End If
    End If
    '������ҳ�ӱ���Ϣ
    mstrPatiPlus = ""
    mstrPatiPlus = mstrPatiPlus & "," & "��ϵ�˸�����Ϣ:" & Trim(txtLinkManInfo.Text)
    mstrPatiPlus = mstrPatiPlus & "," & "��Ժת��:" & Trim(zlStr.NeedName(txtת��.Text))
    If Trim(zlCommFun.GetNeedName(cbo����.Text)) = "�й�" Then
        mstrPatiPlus = mstrPatiPlus & "," & "���֤��״̬:" & Trim(zlCommFun.GetNeedName(cboIDNumber.Text))
        mstrPatiPlus = mstrPatiPlus & "," & "�⼮���֤��:"
    Else
        If Trim(txt���֤��.Text) <> "" Then
            mstrPatiPlus = mstrPatiPlus & "," & "�⼮���֤��:" & txt���֤��.Text
            mstrPatiPlus = mstrPatiPlus & "," & "���֤��״̬:"
            txt���֤��.Text = ""
        Else
            mstrPatiPlus = mstrPatiPlus & "," & "���֤��״̬:" & Trim(zlCommFun.GetNeedName(cboIDNumber.Text))
            mstrPatiPlus = mstrPatiPlus & "," & "�⼮���֤��:"
        End If
    End If
    If mstrPatiPlus <> "" Then mstrPatiPlus = Mid(mstrPatiPlus, 2)
    
    If InStr(1, txt��Ժ���.Tag, ";") <= 0 Then
        lng��ҽ����ID = Val(txt��Ժ���.Tag)
    Else
        lng��ҽ���ID = Val(txt��Ժ���.Tag)
    End If
    If InStr(1, txt��ҽ���.Tag, ";") <= 0 Then
        lng��ҽ����ID = Val(txt��ҽ���.Tag)
    Else
        lng��ҽ���ID = Val(txt��ҽ���.Tag)
    End If
    '����27351 by lesfeng 2010-01-12
    '����24463 by lesfeng 2010-03-22 �������
    '����51167,������,2012-07-09,����"��ϵ�����֤��"
    strSQL = "zl_סԺ������ҳ_Update(" & mlng����ID & "," & mlng��ҳID & ",'" & str���� & "'," & _
        "'" & zlCommFun.GetNeedName(cbo�ѱ�.Text) & "','" & zlCommFun.GetNeedName(cbo����״��.Text) & "','" & zlCommFun.GetNeedName(cboѧ��.Text) & "'," & _
        "'" & zlCommFun.GetNeedName(cboְҵ.Text) & "','" & zlCommFun.GetNeedName(cbo����.Text) & "','" & txt��λ��ַ.Text & "'," & _
        Val(txt��λ��ַ.Tag) & ",'" & txt��λ�绰.Text & "','" & txt��λ�ʱ�.Text & "','" & txt��ͥ��ַ.Text & "'," & _
        "'" & txt��ͥ�绰.Text & "','" & txt��ͥ��ַ�ʱ�.Text & "','" & txt���ڵ�ַ.Text & "','" & txt���ڵ�ַ�ʱ�.Text & "'," & _
        "'" & txt��ϵ������.Text & "','" & zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.Text) & "','" & txt��ϵ�˵绰.Text & "','" & txt��ϵ�˵�ַ.Text & "'," & _
        "'" & zlCommFun.GetNeedName(cbo���λ�ʿ.Text) & "','" & zlCommFun.GetNeedName(cbo����ҽʦ.Text) & "'," & _
        "'" & zlCommFun.GetNeedName(cboסԺҽʦ.Text) & "'," & _
        ZVal(lng��ҽ����ID) & "," & ZVal(lng��ҽ���ID) & ",'" & Replace(txt��Ժ���.Text, "'", "''") & "'," & _
        ZVal(lng��ҽ����ID) & "," & ZVal(lng��ҽ���ID) & ",'" & Replace(txt��ҽ���.Text, "'", "''") & "'," & _
        "'" & UserInfo.��� & "','" & UserInfo.���� & "','" & zlCommFun.GetNeedName(cbo����ҽʦ.Text) & "','" & zlCommFun.GetNeedName(cbo����ҽʦ.Text) & "'," & _
        chk����Ժ.Value & ",'" & Trim(txt���֤��.Text) & "','" & Trim(txt�����ص�.Text) & "','" & zlCommFun.GetNeedName(txt����.Text) & "','" & zlCommFun.GetNeedName(txt����.Text) & "','" & _
        zlCommFun.GetNeedName(cbo��Ժ����.Text) & "','" & zlCommFun.GetNeedName(cbo��������.Text) & "','" & zlCommFun.GetNeedName(cbo��Ժ��ʽ.Text) & "'," & _
        IIf(Trim(txt��ע.Text) = "", "Null", "'" & Trim(txt��ע.Text) & "'") & "," & chk���.Value & ",'" & Trim(txt��ϵ�����֤��.Text) & "','" & zlCommFun.GetNeedName(cbo����.Text) & "')"
        
    ReDim Preserve arrSQL(0)
    arrSQL(UBound(arrSQL)) = strSQL
    
    '������ҳ�ӱ���Ϣ����
    If mstrPatiPlus <> "" Then
        arrTmp = Split(mstrPatiPlus, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If InStr(",��ϵ�˸�����Ϣ,��Ժת��,���֤��״̬,�⼮���֤��,", "," & Split(arrTmp(i), ":")(0) & ",") > 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_������ҳ�ӱ�_��ҳ����(" & mlng����ID & "," & mlng��ҳID & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "')"
            End If
            If InStr(",��ϵ�˸�����Ϣ,���֤��״̬,�⼮���֤��,", "," & Split(arrTmp(i), ":")(0) & ",") > 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "','')"
            End If
        Next
    End If
    
    If Val(cbo�ѱ�.Tag) <> cbo�ѱ�.ListIndex And InStr(";" & mstrPrivs & ";", ";�������;") > 0 And Nvl(mrsPati!����, 0) = 0 Then
        If MsgBox("���˷ѱ𱻸ı䣬Ҫ���ò��˵�δ����ð��µķѱ�������?" & vbCrLf & vbCrLf & _
            "�������������˵�ǰ�ѱ��Ӧ���Żݱ��ʶ�δ��������½��д��ۼ���!", vbInformation + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
            
            strSQL_Recalc = "Zl_����δ�����_Recalc(" & mlng����ID & "," & mlng��ҳID & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL_Recalc
        End If
    End If
    
    If gbln���ýṹ����ַ Then
        Call CreateStructAddressSQL(mlng����ID, mlng��ҳID, arrSQL, PatiAddress, 1)
    End If
    
    '���ֻ����סԺҽʦ����ֻ����סԺҽʦ�䶯���������ҽ��С�飬��ֻ����ҽ��С��䶯��
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    strBeginDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    For i = LBound(arrSQL) To UBound(arrSQL)
        zlDatabase.ExecuteProcedure arrSQL(i), Me.Caption
    Next
    strEndDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    
    '����ҽ��������Ϣ�޸Ľӿ�
    If Not IsNull(mrsPati!����) Then
        If Not gclsInsure.ModiPatiSwap(mlng����ID, mlng��ҳID, mrsPati!����, "1") Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
    End If
    '����EMPIƽ̨������Ϣ
    strMsg = ""
    If Not EMPI_AddORUpdatePati(mlng����ID, mlng��ҳID, strMsg) Then
        gcnOracle.RollbackTrans
        MsgBox strMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    gcnOracle.CommitTrans: blnTrans = False
    gblnOK = True
    '����96847��118004
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(2, mlng����ID, mlng��ҳID) <> 1 Then
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������Ӱ����Ϣϵͳ�ӿ�(HISModPati)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISModPati)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
    End If
    On Error Resume Next
    '���˻�����Ϣ�����Ƿ�ɹ�������Ӱ�첡����Ϣ����
    If bln������Ϣ���� And blnMod Then
        strErrInfo = ""
        Call gobjPublicPatient.SavePatiBaseInfo(mlng����ID, mlng��ҳID, Trim(txt����.Text), str�Ա�, strAge, str��������, Me.Caption, IIf(mlng��ҳID <> 0, 2, 1), strErrInfo, True, True)
        '��ʾ
        If strErrInfo <> "" Then
            MsgBox strErrInfo, vbInformation + vbOKOnly, Me.Caption
        End If
    End If

    '����仯�󴥷���Ϣ
    If zlCommFun.GetNeedName(cbo����.Text) <> Nvl(mrsPati!��ǰ����) Then
        Call PatiInfoChange(13, strBeginDate, strEndDate)
    End If
    'סԺҽʦ�䶯�󴥷���Ϣ
    If zlCommFun.GetNeedName(cboסԺҽʦ.Text) <> Nvl(mrsPati!סԺҽʦ) Then
        Call PatiInfoChange(7, strBeginDate, strEndDate)
    End If
    '���λ�ʿ�䶯�󴥷���Ϣ
    If zlCommFun.GetNeedName(cbo���λ�ʿ.Text) <> Nvl(mrsPati!���λ�ʿ) Then
        Call PatiInfoChange(8, strBeginDate, strEndDate)
    End If
    '����ҽʦ�䶯�󴥷���Ϣ
    If zlCommFun.GetNeedName(cbo����ҽʦ.Text) <> Nvl(mrsPati!����ҽʦ) Then
        Call PatiInfoChange(11, strBeginDate, strEndDate)
    End If
    '����ҽʦ�䶯�󴥷���Ϣ
    If zlCommFun.GetNeedName(cbo����ҽʦ.Text) <> Nvl(mrsPati!����ҽʦ) Then
        Call PatiInfoChange(12, strBeginDate, strEndDate)
    End If
    
    If Err <> 0 Then Err.Clear
    
    Unload Me
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd�����ص�_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetAddress(Me, txt�����ص�, True)
    If Not rsTmp Is Nothing Then
        txt�����ص�.Text = rsTmp!����
        txt�����ص�.SelStart = Len(txt�����ص�.Text)
        txt�����ص�.SetFocus
    End If
End Sub

Private Sub cmd���ڵ�ַ_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetAddress(Me, txt���ڵ�ַ, True)
    If Not rsTmp Is Nothing Then
        txt���ڵ�ַ.Text = rsTmp!����
        txt���ڵ�ַ.SelStart = Len(txt���ڵ�ַ.Text)
        txt���ڵ�ַ.SetFocus
    End If
End Sub

Private Sub cmd����_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt����, True)
    If Not rsTmp Is Nothing Then
        txt����.Text = rsTmp!����
        txt����.SelStart = Len(txt����.Text)
        txt����.SetFocus
    Else
        zlControl.TxtSelAll txt����
        txt����.SetFocus
    End If
End Sub

Private Sub cmd����_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt����, True)
    If Not rsTmp Is Nothing Then
        txt����.Text = rsTmp!����
        txt����.SelStart = Len(txt����.Text)
        txt����.SetFocus
    Else
        zlControl.TxtSelAll txt����
        txt����.SetFocus
    End If
End Sub

Private Sub cmdת��_Click()
    Dim vPoint As POINTAPI
    On Error GoTo errH
    vPoint = GetCoordPos(txtת��.Container.hWnd, txtת��.Left, txtת��.Top)
    Call Getҽ�ƻ���(txtת��, Me, 2, "ҽ�ƻ���", "�ֵ������", vPoint, False)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
        Case vbKeyReturn
            If InStr(UCase(",txt��λ��ַ,txt���ڵ�ַ,txt�����ص�,txt��ͥ��ַ,txt��ϵ�˵�ַ,txt��Ժ���,txt��ҽ���,txt����,txt����,PatiAddress,"), UCase("," & ActiveControl.Name & ",")) = 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case vbKeyF3
            If Me.ActiveControl.Name = txt��λ��ַ.Name Then
                cmd��λ��ַ_Click
            ElseIf Me.ActiveControl.Name = txt��ͥ��ַ.Name Then
                cmd��ͥ��ַ_Click
            ElseIf Me.ActiveControl.Name = txt�����ص�.Name Then
                cmd�����ص�_Click
            ElseIf Me.ActiveControl.Name = txt���ڵ�ַ.Name Then
                cmd���ڵ�ַ_Click
            ElseIf Me.ActiveControl.Name = txt��ϵ�˵�ַ.Name Then
                cmd��ϵ�˵�ַ_Click
            ElseIf Me.ActiveControl.Name = txt����.Name Then
                cmd����_Click
            ElseIf Me.ActiveControl.Name = txt����.Name Then
                cmd����_Click
            End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") And Not (Me.ActiveControl Is txt��Ժ��� Or Me.ActiveControl Is txt��ҽ���) Then KeyAscii = 0      '��������п�����'��
    '��ϵ�˹�ϵ˵����ת�벻����¼�붺�ź�ð��,��Ϊ �ö���mstrPatiPlus�� �ķָ��� ����ð�źͶ���
    If Me.ActiveControl Is txtת�� Or Me.ActiveControl Is txtLinkManInfo Then
        If InStr(":��,��", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String, strTmp As String
    Dim rsDiagnosis As ADODB.Recordset, rsBeds As ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    
    gblnOK = False
    mblnEMPI = False
    Call InitStructAddress
    '����27351 by lesfeng 2010-01-12 ,A.��ע
    '����24463 by lesfeng 2010-03-22 �������
    On Error GoTo errH
    strSQL = "Select NVL(A.����,D.����) ����,NVL(A.�Ա�,D.�Ա�) �Ա�, NVL(A.����,D.����) ����,To_Char(A.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��,E.���� as ��ǰ����,A.��Ժ����id as ��ǰ����ID,H.���� as ��ǰ����,A.��ǰ����Id, A.ҽ��С��id, g.���� as ҽ��С��, " & vbNewLine & _
            "A.סԺ��,A.���λ�ʿ, A.����ҽʦ, A.סԺҽʦ, B.��Ϣֵ ����ҽʦ, C.��Ϣֵ ����ҽʦ, A.�ѱ�, A.����״��, A.ѧ��," & vbNewLine & _
            "       A.ְҵ, A.��ǰ����, A.��λ��ַ, A.��λ�ʱ�, A.��λ�绰, A.��ͥ��ַ, A.��ͥ�绰, A.��ͥ��ַ�ʱ�, A.���ڵ�ַ, A.���ڵ�ַ�ʱ�, A.��ϵ�˵�ַ," & vbNewLine & _
            "       A.��ϵ�˵绰, A.��ϵ������, A.��ϵ�˹�ϵ,A.��ϵ�����֤��, A.����Ժ, A.��������, A.����,D.����, D.���֤��, D.����, D.����, D.�����ص�," & vbNewLine & _
            "       D.��������, A.��Ժ����,D.��ͬ��λid, F.���� As ����ȼ�,Nvl(A.��������,Decode(A.����,Null,'��ͨ����','ҽ������')) ��������,A.��Ժ��ʽ,A.��ע,A.�Ƿ����" & vbNewLine & _
            "From ������ҳ A, ������ҳ�ӱ� B, ������ҳ�ӱ� C, ������Ϣ D,���ű� E,���ű� H,�շ���ĿĿ¼ F, �ٴ�ҽ��С�� G " & vbNewLine & _
            "Where A.����id = [1] And A.��ҳid = [2] And A.����id = B.����id(+) And A.��ҳid = B.��ҳid(+) And A.����id = C.����id(+) And" & vbNewLine & _
            "      A.��ҳid = C.��ҳid(+) And A.ҽ��С��id = G.id(+) And B.��Ϣ��(+) = '����ҽʦ' And C.��Ϣ��(+) = '����ҽʦ' And A.����id = D.����id And A.��Ժ����id = E.id And A.��ǰ����Id=H.id(+)" & vbNewLine & _
            " And A.����ȼ�id = F.ID(+)"
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        
    With mrsPati
        txt����.Text = !����
        txtסԺ��.Text = "" & !סԺ��
        txt�Ա�.Text = "" & !�Ա�
        txt��Ժʱ��.Text = !��Ժʱ��
        txt����.Text = "" & !��ǰ����
        txt����.Tag = Val("" & !��ǰ����id)
        '�Ƿ�����ҽ��
        txt��ҽ���.Enabled = (InStr(1, "," & GetDepCharacter(Val("" & !��ǰ����id)) & ",", ",��ҽ��,") > 0)
        txt��ҽ���.ToolTipText = "ֻ�е��������ڿ��ҵ�����Ϊ��ҽ��ʱ������������ҽ���!"
        
        txt����.Text = "" & !����ȼ�
        txt��λ��ַ.Tag = Val("" & !��ͬ��λID)
        txtҽ��С��.Text = "" & !ҽ��С��
        txtҽ��С��.Tag = "" & !ҽ��С��id
        mstrBirthDay = "" & !��������
    End With
    
    Set rsBeds = GetPatiBeds(mlng����ID)
    With rsBeds
        If .RecordCount = 0 Then
            txt��λ.Text = "��ͥ����"
            txt�ȼ�.Text = "��"
        Else
            strTmp = ""
            Do While Not .EOF
                txt��λ.Text = txt��λ.Text & "," & !����
                If InStr("," & strTmp & ",", "," & !��λ�ȼ� & ",") = 0 Then strTmp = strTmp & "," & !��λ�ȼ�
                .MoveNext
            Loop
            txt��λ.Text = Mid(txt��λ.Text, 2)
            txt�ȼ�.Text = Mid(strTmp, 2)
        End If
    End With

    Call InitDicts
    Call LoadOldData("" & mrsPati!����, txt����, cbo���䵥λ)
    txt���֤��.Text = "" & mrsPati!���֤��
    cboIDNumber.Enabled = txt���֤��.Text = ""
    txt����.Text = Nvl(mrsPati!����)
    
    cbo����.ListIndex = cbo.FindIndex(cbo����, IIf(IsNull(mrsPati!����), "", mrsPati!����))
    If cbo����.ListIndex = -1 Then Call SetCboDefault(cbo����)
    
    If InStr(mstrPrivs, "��������ҽʦ") = 0 Then cbo����ҽʦ.Enabled = False
    cboסԺҽʦ.ListIndex = cbo.FindIndex(cboסԺҽʦ, IIf(IsNull(mrsPati!סԺҽʦ), "", mrsPati!סԺҽʦ), True)
    cbo����ҽʦ.ListIndex = cbo.FindIndex(cbo����ҽʦ, IIf(IsNull(mrsPati!����ҽʦ), "", mrsPati!����ҽʦ), True)
    
    cbo���λ�ʿ.ListIndex = cbo.FindIndex(cbo���λ�ʿ, IIf(IsNull(mrsPati!���λ�ʿ), "", mrsPati!���λ�ʿ), True)
    cbo����ҽʦ.ListIndex = cbo.FindIndex(cbo����ҽʦ, IIf(IsNull(mrsPati!����ҽʦ), "", mrsPati!����ҽʦ), True)
    cbo����ҽʦ.ListIndex = cbo.FindIndex(cbo����ҽʦ, IIf(IsNull(mrsPati!����ҽʦ), "", mrsPati!����ҽʦ), True)
            
    cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, IIf(IsNull(mrsPati!�ѱ�), "", mrsPati!�ѱ�), True)
    cbo�ѱ�.Tag = cbo�ѱ�.ListIndex '��¼ԭʼ�ѱ����ڱ���ʱ�ж��Ƿ�����������
    cbo�ѱ�.Enabled = InStr(mstrPrivs, "�������˷ѱ�") > 0
    
    cbo����״��.ListIndex = cbo.FindIndex(cbo����״��, IIf(IsNull(mrsPati!����״��), "", mrsPati!����״��), True)
    cboѧ��.ListIndex = cbo.FindIndex(cboѧ��, IIf(IsNull(mrsPati!ѧ��), "", mrsPati!ѧ��), True)
    cboְҵ.ListIndex = cbo.FindIndex(cboְҵ, IIf(IsNull(mrsPati!ְҵ), "", mrsPati!ְҵ), True)
    cbo����.ListIndex = cbo.FindIndex(cbo����, IIf(IsNull(mrsPati!��ǰ����), "", mrsPati!��ǰ����), True)
    cbo��Ժ����.ListIndex = cbo.FindIndex(cbo��Ժ����, IIf(IsNull(mrsPati!��Ժ����), "", mrsPati!��Ժ����), True)
    cbo��Ժ��ʽ.ListIndex = cbo.FindIndex(cbo��Ժ��ʽ, IIf(IsNull(mrsPati!��Ժ��ʽ), "", mrsPati!��Ժ��ʽ), True)
    
    txt��λ��ַ.Text = IIf(IsNull(mrsPati!��λ��ַ), "", mrsPati!��λ��ַ)
    txt��λ�ʱ�.Text = IIf(IsNull(mrsPati!��λ�ʱ�), "", mrsPati!��λ�ʱ�)
    txt��λ�绰.Text = IIf(IsNull(mrsPati!��λ�绰), "", mrsPati!��λ�绰)
    cbo��������.ListIndex = cbo.FindIndex(cbo��������, mrsPati!��������, True)
    If InStr(mstrPrivs, "������������") = 0 Then cbo��������.Enabled = False
        
    txt��ͥ�绰.Text = IIf(IsNull(mrsPati!��ͥ�绰), "", mrsPati!��ͥ�绰)
    txt��ͥ��ַ�ʱ�.Text = IIf(IsNull(mrsPati!��ͥ��ַ�ʱ�), "", mrsPati!��ͥ��ַ�ʱ�)
    txt���ڵ�ַ�ʱ�.Text = IIf(IsNull(mrsPati!���ڵ�ַ�ʱ�), "", mrsPati!���ڵ�ַ�ʱ�)

    txt��ϵ�˵绰.Text = IIf(IsNull(mrsPati!��ϵ�˵绰), "", mrsPati!��ϵ�˵绰)
    txt��ϵ������.Text = IIf(IsNull(mrsPati!��ϵ������), "", mrsPati!��ϵ������)
    txt��ϵ�����֤��.Text = IIf(IsNull(mrsPati!��ϵ�����֤��), "", mrsPati!��ϵ�����֤��)
    cbo��ϵ�˹�ϵ.ListIndex = cbo.FindIndex(cbo��ϵ�˹�ϵ, IIf(IsNull(mrsPati!��ϵ�˹�ϵ), "", mrsPati!��ϵ�˹�ϵ), True)
    
    '�����������ýṹ����ַ,������Ϣ,������ҳ���ᱣ���ַ��Ϣ
    If gbln���ýṹ����ַ Then
        Call ReadStructAddress(mlng����ID, mlng��ҳID, PatiAddress)
        txt�����ص�.Text = PatiAddress(E_IX_�����ص�).Value
        txt����.Text = PatiAddress(E_IX_����).Value
        txt��ͥ��ַ.Text = PatiAddress(E_IX_��סַ).Value
        txt���ڵ�ַ.Text = PatiAddress(E_IX_���ڵ�ַ).Value
        txt��ϵ�˵�ַ.Text = PatiAddress(E_IX_��ϵ�˵�ַ).Value
    Else
        txt�����ص�.Text = "" & mrsPati!�����ص�
        txt����.Text = Nvl(mrsPati!����)
        txt��ͥ��ַ.Text = IIf(IsNull(mrsPati!��ͥ��ַ), "", mrsPati!��ͥ��ַ)
        txt���ڵ�ַ.Text = IIf(IsNull(mrsPati!���ڵ�ַ), "", mrsPati!���ڵ�ַ)
        txt��ϵ�˵�ַ.Text = IIf(IsNull(mrsPati!��ϵ�˵�ַ), "", mrsPati!��ϵ�˵�ַ)
    End If
    
    '����27351 by lesfeng 2010-01-12
    txt��ע.Text = IIf(IsNull(mrsPati!��ע), "", mrsPati!��ע)
    '����24463 by lesfeng 2010-03-22 �������
    chk���.Value = IIf(IsNull(mrsPati!�Ƿ����), 0, mrsPati!�Ƿ����)
    
     '��ʾ������ϼ�¼
    Set rsDiagnosis = GetDiagnosticInfo(mlng����ID, mlng��ҳID, "1,11,2,12", "2")
    If Not rsDiagnosis Is Nothing Then
        'a.��ҽ���
         rsDiagnosis.Filter = "�������=2"        '��ȡ��ǰ�������Ժ���
         If Not rsDiagnosis.EOF Then
             txt��Ժ���.Text = Nvl(rsDiagnosis!�������): txt��Ժ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��Ժ���.Tag = txt��Ժ���.Text
         Else
             rsDiagnosis.Filter = "�������=1"    '��ȡ��Ժ�Ǽǵ��������
             If Not rsDiagnosis.EOF Then
                 txt��Ժ���.Text = Nvl(rsDiagnosis!�������): txt��Ժ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��Ժ���.Tag = txt��Ժ���.Text
             End If
         End If
     
        'b.��ҽ���
        If txt��ҽ���.Enabled Then
            rsDiagnosis.Filter = "�������=12"        '��ȡ��ǰ�������Ժ���
            If Not rsDiagnosis.EOF Then
                txt��ҽ���.Text = Nvl(rsDiagnosis!�������): txt��ҽ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��ҽ���.Tag = txt��ҽ���.Text
            Else
                rsDiagnosis.Filter = "�������=11"    '��ȡ��Ժ�Ǽǵ��������
                If Not rsDiagnosis.EOF Then
                    txt��ҽ���.Text = Nvl(rsDiagnosis!�������): txt��ҽ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��ҽ���.Tag = txt��ҽ���.Text
                End If
            End If
        End If
    End If
    chk����Ժ.Value = Val("" & mrsPati!����Ժ)
    
    
    '54045:������,2012-09-27,�����ҳ�ж�Ӧ��ҽʦǩ���������޸�
'    strSql = "Select ��Ϣ��,��Ϣֵ From ������ҳ�ӱ� Where ����ID=[1] And ��ҳID=[2] And ��Ϣֵ is Not Null"
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mlng��ҳID)
    Set rsTmp = Get������ҳ�ӱ�(mlng����ID, mlng��ҳID, "")
    rsTmp.Filter = "��Ϣ��='סԺҽʦǩ��'"
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!��Ϣֵ) Then
            cboסԺҽʦ.Enabled = False
            cboסԺҽʦ.BackColor = &HE0E0E0
        End If
    End If
    rsTmp.Filter = "��Ϣ��='����ҽʦǩ��'"
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!��Ϣֵ) Then
            cbo����ҽʦ.Enabled = False
            cbo����ҽʦ.BackColor = &HE0E0E0
        End If
    End If
    rsTmp.Filter = "��Ϣ��='����ҽʦǩ��'"
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!��Ϣֵ) Then
            cbo����ҽʦ.Enabled = False
            cbo����ҽʦ.BackColor = &HE0E0E0
        End If
    End If
    rsTmp.Filter = "��Ϣ��='��ϵ�˸�����Ϣ'"
    If Not rsTmp.EOF Then txtLinkManInfo.Text = rsTmp!��Ϣֵ & ""
    
    rsTmp.Filter = "��Ϣ��='��Ժת��'"
    If Not rsTmp.EOF Then txtת��.Text = rsTmp!��Ϣֵ & ""
    
    '���˴ӱ�
    Set rsTmp = Get������Ϣ�ӱ�(mlng����ID, "���֤��״̬")
    rsTmp.Filter = "��Ϣ��='���֤��״̬'"
    If Not rsTmp.EOF Then
        Call cbo.Locate(cboIDNumber, zlCommFun.GetNeedName(rsTmp!��Ϣֵ) & "")
    End If
    If Trim(zlCommFun.GetNeedName(cbo����.Text)) <> "�й�" And Trim(txt���֤��.Text) = "" Then
        If Trim(zlCommFun.GetNeedName(cboIDNumber.Text)) = "" Then
            Set rsTmp = Get������Ϣ�ӱ�(mlng����ID, "�⼮���֤��")
            rsTmp.Filter = "��Ϣ��='�⼮���֤��'"
            If Not rsTmp.EOF Then
                txt���֤��.Text = "" & rsTmp!��Ϣֵ
            End If
        End If
    End If
    '����EMPIƽ̨������Ϣ
    Call EMPI_LoadPati
    '������Ϣ����
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, P�����������, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'ж����Ϣ����
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
End Sub


Private Sub PatiAddress_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(True) '���������뷨
End Sub

Private Sub PatiAddress_LostFocus(Index As Integer)
'����:
    Select Case Index
    
    Case E_IX_��סַ
        txt��ͥ��ַ.Text = PatiAddress(Index).Value
    Case E_IX_�����ص�
        txt�����ص�.Text = PatiAddress(Index).Value
    Case E_IX_���ڵ�ַ
        txt���ڵ�ַ.Text = PatiAddress(Index).Value
    Case E_IX_����
        txt����.Text = PatiAddress(Index).Value
    Case E_IX_��ϵ�˵�ַ
        txt��ϵ�˵�ַ.Text = PatiAddress(Index).Value
    End Select
    Call zlCommFun.OpenIme '�ر��������뷨
End Sub

Private Sub PatiAddress_Validate(Index As Integer, Cancel As Boolean)
    Dim lngLen As Long
    
    lngLen = PatiAddress(Index).MaxLength
    If LenB(StrConv(PatiAddress(Index).Value, vbFromUnicode)) > lngLen Then
        MsgBox PatiAddress(Index).Tag & "ֻ�������� " & lngLen & " ���ַ��� " & lngLen \ 2 & " �����֣�", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub txtLinkManInfo_GotFocus()
    zlControl.TxtSelAll txtLinkManInfo
    Call zlCommFun.OpenIme(True)
End Sub


Private Sub txtLinkManInfo_LostFocus()
     Call zlCommFun.OpenIme
End Sub

'����27351 by lesfeng 2010-01-12  b
Private Sub txt��ע_GotFocus()
    Call zlControl.TxtSelAll(txt��ע)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ע_KeyPress(KeyAscii As Integer)
    CheckInputLen txt��ע, KeyAscii
End Sub

Private Sub txt��ע_LostFocus()
    Call zlCommFun.OpenIme
End Sub
'����27351 by lesfeng 2010-01-12 e
Private Sub txt�����ص�_GotFocus()
    zlControl.TxtSelAll txt�����ص�
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt�����ص�_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt�����ص�.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt�����ص�)
            If Not rsTmp Is Nothing Then
                txt�����ص�.Text = rsTmp!����
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt�����ص�, KeyAscii
    End If
End Sub

Private Sub txt�����ص�_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt��λ�绰_GotFocus()
    zlControl.TxtSelAll txt��λ�绰
End Sub

Private Sub txt��λ�绰_KeyPress(KeyAscii As Integer)
    If InStr("01234567890()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txt��λ�ʱ�_GotFocus()
    zlControl.TxtSelAll txt��λ�ʱ�
End Sub

Private Sub txt��λ�ʱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If ((Not IsNumeric(txt��λ�ʱ�.Text)) Or Len(txt��λ�ʱ�.Text) > 6 Or InStr(txt��λ�ʱ�.Text, ".") > 0) And txt��λ�ʱ�.Text <> "" Then
            Call SelectYouBian(txt��λ�ʱ�)
        End If
    End If
End Sub

Private Sub txt���ڵ�ַ_GotFocus()
    zlControl.TxtSelAll txt���ڵ�ַ
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt���ڵ�ַ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt���ڵ�ַ.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt���ڵ�ַ)
            If Not rsTmp Is Nothing Then
                txt���ڵ�ַ.Text = rsTmp!����
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt���ڵ�ַ, KeyAscii
    End If
End Sub

Private Sub txt���ڵ�ַ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt���ڵ�ַ�ʱ�_GotFocus()
    zlControl.TxtSelAll txt���ڵ�ַ�ʱ�
End Sub

Private Sub txt���ڵ�ַ�ʱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If ((Not IsNumeric(txt���ڵ�ַ�ʱ�.Text)) Or Len(txt���ڵ�ַ�ʱ�.Text) > 6 Or InStr(txt���ڵ�ַ�ʱ�.Text, ".") > 0) And txt���ڵ�ַ�ʱ�.Text <> "" Then
            Call SelectYouBian(txt���ڵ�ַ�ʱ�)
        End If
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt����.Text <> "" Then
            Set rsTmp = GetArea(Me, txt����)
            If Not rsTmp Is Nothing Then
                txt����.Text = rsTmp!����
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                zlControl.TxtSelAll txt����
                txt����.SetFocus
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt����, KeyAscii
    End If
End Sub

Private Sub txt����_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��ͥ��ַ�ʱ�_GotFocus()
    zlControl.TxtSelAll txt��ͥ��ַ�ʱ�
End Sub

Private Sub txt��ͥ��ַ�ʱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If ((Not IsNumeric(txt��ͥ��ַ�ʱ�.Text)) Or Len(txt��ͥ��ַ�ʱ�.Text) > 6 Or InStr(txt��ͥ��ַ�ʱ�.Text, ".") > 0) And txt��ͥ��ַ�ʱ�.Text <> "" Then
            Call SelectYouBian(txt��ͥ��ַ�ʱ�)
        End If
    End If
End Sub

Private Sub txt��ͥ��ַ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt��ͥ�绰_GotFocus()
    zlControl.TxtSelAll txt��ͥ�绰
End Sub

Private Sub txt��ͥ�绰_KeyPress(KeyAscii As Integer)
    If InStr("01234567890()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txt��ϵ�˵绰_GotFocus()
    zlControl.TxtSelAll txt��ϵ�˵绰
End Sub

Private Sub txt��ϵ�˵绰_KeyPress(KeyAscii As Integer)
    If InStr("01234567890()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub


Private Sub txt��ϵ������_GotFocus()
    zlControl.TxtSelAll txt��ϵ������
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ϵ������_KeyPress(KeyAscii As Integer)
    CheckInputLen txt��ϵ������, KeyAscii
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txt��ϵ������_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt����_GotFocus()
    Call zlCommFun.OpenIme
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbo���䵥λ.Visible = False And IsNumeric(txt����.Text) Then
            Call txt����_Validate(False)
            Call cbo���䵥λ.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt����.Text) Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt��λ��ַ_Change()
    If txt��λ��ַ.Text = "" Then txt��λ��ַ.Tag = ""
End Sub

Private Sub txt��λ��ַ_GotFocus()
    zlControl.TxtSelAll txt��λ��ַ
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��λ��ַ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt��λ��ַ.Text <> "" Then
            Set rsTmp = GetOrgAddress(Me, txt��λ��ַ)
            If Not rsTmp Is Nothing Then
                txt��λ��ַ.Text = rsTmp!����
                txt��λ��ַ.Tag = rsTmp!ID
                txt��λ�绰.Text = Trim(rsTmp!�绰 & "")
            Else
                txt��λ��ַ.Tag = ""
            End If
        Else
            txt��λ��ַ.Tag = ""
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt��λ��ַ, KeyAscii
    End If
End Sub

Private Sub txt��λ��ַ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt��ͥ��ַ_GotFocus()
    zlControl.TxtSelAll txt��ͥ��ַ
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ͥ��ַ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt��ͥ��ַ.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt��ͥ��ַ)
            If Not rsTmp Is Nothing Then
                txt��ͥ��ַ.Text = rsTmp!����
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt��ͥ��ַ, KeyAscii
    End If
End Sub

Private Sub txt��ϵ�˵�ַ_GotFocus()
    zlControl.TxtSelAll txt��ϵ�˵�ַ
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ϵ�˵�ַ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt��ϵ�˵�ַ.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt��ϵ�˵�ַ)
            If Not rsTmp Is Nothing Then
                txt��ϵ�˵�ַ.Text = rsTmp!����
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt��ϵ�˵�ַ, KeyAscii
    End If
End Sub

Private Sub txt��ϵ�˵�ַ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub cmd��λ��ַ_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetOrgAddress(Me, txt��λ��ַ, True)
    If Not rsTmp Is Nothing Then
        txt��λ��ַ.Tag = rsTmp!ID
        txt��λ��ַ.Text = rsTmp!����
        txt��λ��ַ.SelStart = Len(txt��λ��ַ.Text)
        txt��λ�绰.Text = Trim(rsTmp!�绰 & "")
        txt��λ��ַ.SetFocus
    End If
End Sub

Private Sub cmd��ͥ��ַ_Click()
    Dim rsTmp As ADODB.Recordset
   Set rsTmp = GetAddress(Me, txt��ͥ��ַ, True)
    If Not rsTmp Is Nothing Then
        txt��ͥ��ַ.Text = rsTmp!����
        txt��ͥ��ַ.SelStart = Len(txt��ͥ��ַ.Text)
        txt��ͥ��ַ.SetFocus
    End If
End Sub

Private Sub cmd��ϵ�˵�ַ_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetAddress(Me, txt��ϵ�˵�ַ, True)
    If Not rsTmp Is Nothing Then
        txt��ϵ�˵�ַ.Text = rsTmp!����
        txt��ϵ�˵�ַ.SelStart = Len(txt��ϵ�˵�ַ.Text)
        txt��ϵ�˵�ַ.SetFocus
    End If
End Sub

Private Sub InitDicts()
    Dim strSQL As String, i As Integer
    Dim strSQLҽ��С�� As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    cbo���䵥λ.Clear
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.ListIndex = 0
    txt����.MaxLength = GetColumnLength("������Ϣ", "����")
    
    Call ReadDict("�ѱ�", cbo�ѱ�)
    Call ReadDict("����", cbo����)
    Call ReadDict("ѧ��", cboѧ��)
    Call ReadDict("����״��", cbo����״��)
    Call ReadDict("ְҵ", cboְҵ)
    Call ReadDict("����ϵ", cbo��ϵ�˹�ϵ)
    Call ReadDict("��Ժ����", cbo��Ժ����)
    Call ReadDict("��Ժ��ʽ", cbo��Ժ��ʽ)
    Call ReadDict("��������", cbo��������, "��������")
    Call ReadDict("���֤δ¼ԭ��", cboIDNumber)
    Call ReadDict("����", cbo����)
    
    mstrUnit = Get����IDs(mlngUnit) & "," & mlngUnit

    'ҽ��С��
    strSQL = "Select Distinct A.ID, A.���, A.����, A.����" & vbNewLine & _
                        " From ��Ա�� A, ��Ա����˵�� B, ������Ա C" & vbNewLine & _
                        " Where A.ID = B.��Աid And A.ID = C.��Աid And B.��Ա���� = 'ҽ��' And" & vbNewLine & _
                        "      (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & vbNewLine & _
                        "      (Instr(',' || [1] || ',', ',' || C.����id || ',') > 0 Or A.����=[2]) And Instr(',' || [3] || ',', ',' || A.רҵ����ְ�� || ',') > 0" & vbNewLine & _
                        "      And (A.վ��=[4] Or A.վ�� is Null)" & _
                        " Order By A.����"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit, CStr("" & mrsPati!סԺҽʦ), "����ҽʦ,������ҽʦ,����ҽʦ,ҽʦ,ҽʿ", gstrNodeNo)
    cboסԺҽʦ.Clear
    Do Until rsTmp.EOF
        cboסԺҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
        cboסԺҽʦ.ItemData(cboסԺҽʦ.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit, CStr("" & mrsPati!����ҽʦ), "����ҽʦ,������ҽʦ,����ҽʦ", gstrNodeNo)
    cbo����ҽʦ.Clear
    Do Until rsTmp.EOF
        cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
        cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit, CStr("" & mrsPati!����ҽʦ), "����ҽʦ,������ҽʦ", gstrNodeNo)
    Do Until rsTmp.EOF
        cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
        cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop

    '����ҽʦ
    Set rsTmp = GetDoctorOrNurse(0)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Next
    End If
    
    'by lesfeng 2010-01-12 �����Ż�
    'סԺ��ʿ
    strSQL = _
        "Select Distinct A.ID,A.���,A.����,A.����" & _
        " From ��Ա�� A,��Ա����˵�� B,������Ա C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And B.��Ա����=[1] And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (Instr(','||[2]||',',','||C.����ID||',')>0  Or A.����=[3])" & _
        " And (A.վ��=[4] Or A.վ�� is Null)" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "��ʿ", mstrUnit, CStr("" & mrsPati!���λ�ʿ), gstrNodeNo)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo���λ�ʿ.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo���λ�ʿ.ItemData(cbo���λ�ʿ.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Next
    End If
    
    cboסԺҽʦ.AddItem "����..."
    cbo����ҽʦ.AddItem "����..."
    cbo����ҽʦ.AddItem "����..."
    cbo���λ�ʿ.AddItem "����..."
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ReadDict(strDict As String, cboInput As ComboBox, Optional strClass As String) As Boolean
'���ܣ���ʼ��ָ���ʵ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim lngMaxW As Long
    Dim strTemp As String

    On Error GoTo errH
    'by lesfeng 2010-01-12 �����Ż�
    If strDict = "�ѱ�" Then
        If Nvl(mrsPati!��������, 0) = 1 Then
            strTemp = "1,3" '�������۲���
        Else
            strTemp = "2,3"
        End If
'        strSql = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �ѱ� Where Nvl(�������,3) IN(" & strTemp & ") And  Sysdate Between NVL(��Ч��ʼ,Sysdate-1) and NVL(��Ч����,Sysdate+1) Order by ����"
        strSQL = "Select A.����,A.����,A.����,Nvl(A.ȱʡ��־,0) as ȱʡ From �ѱ� A,Table(Cast(f_Num2List([1]) As zlTools.t_Numlist)) B " & _
                 " Where (A.������� = B.Column_Value or A.������� is null) And (a.��Ч��ʼ Is Null And a.��Ч���� Is Null Or Trunc(Sysdate) Between a.��Ч��ʼ And a.��Ч����) Order by A.����"
    ElseIf InStr(",����,", "," & strDict & ",") > 0 Then
        strSQL = "Select ����,����,����,0 as ȱʡ From " & strDict & " Order by ����"
    ElseIf strDict = "��������" Then
        strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ,��ɫ From �������� Order by ����"
    Else
        strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strDict & " Order by ����"
    End If
'    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption, strTemp)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strTemp)
    
    cboInput.Clear
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If strDict = "ְҵ" Then
                cboInput.AddItem rsTmp!���� & "-" & Chr(&HA) & rsTmp!����
            Else
                cboInput.AddItem rsTmp!���� & "-" & rsTmp!����
            End If
            If rsTmp!ȱʡ = 1 Then
                cboInput.ListIndex = cboInput.NewIndex
                cboInput.ItemData(cboInput.NewIndex) = 1
            End If
            If TextWidth(cboInput.List(cboInput.NewIndex) & "��") > lngMaxW Then lngMaxW = TextWidth(cboInput.List(cboInput.NewIndex) & "��")
            rsTmp.MoveNext
        Next
    End If
    ReadDict = True
    If cbo.ListWidth(cboInput.hWnd) < lngMaxW Then cbo.SetListWidth cboInput.hWnd, lngMaxW
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt����_Validate(Cancel As Boolean)
    If Not IsNumeric(txt����.Text) And Trim(txt����.Text) <> "" Then
        If Not InStr(Trim(txt����.Text), "Լ") > 0 And Trim(txt����.Text) <> "����" Then
            cbo���䵥λ.ListIndex = -1: cbo���䵥λ.Visible = False
        End If
    ElseIf cbo���䵥λ.Visible = False Then
        cbo���䵥λ.ListIndex = 0: cbo���䵥λ.Visible = True
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt����.Text <> "" Then
            Set rsTmp = GetArea(Me, txt����)
            If Not rsTmp Is Nothing Then
                txt����.Text = rsTmp!����
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                zlControl.TxtSelAll txt����
                txt����.SetFocus
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt����, KeyAscii
    End If
End Sub

Private Sub txt����_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��Ժ���_GotFocus()
    zlControl.TxtSelAll txt��Ժ���
End Sub

Private Sub txt���֤��_GotFocus()
    zlControl.TxtSelAll txt���֤��
End Sub

Private Sub txt���֤��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt��ϵ�����֤��_GotFocus()
    zlControl.TxtSelAll txt���֤��
End Sub

Private Sub txt��ϵ�����֤��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt���֤��_LostFocus()
    If Trim(txt���֤��.Text) = "" Then
        cboIDNumber.Enabled = True
        cboIDNumber.SetFocus
    Else
        cboIDNumber.Enabled = False
        cboIDNumber.ListIndex = -1
    End If
End Sub

Private Sub txt��ҽ���_GotFocus()
    zlControl.TxtSelAll txt��ҽ���
End Sub

Private Sub txt��Ժ���_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            '����25785 by lesfeng 2009-10-20 ������������¼�����
            '************************************************
            If gintסԺ������� = 1 Then
                strInput = UCase(txt��Ժ���.Text)
                strSex = zlCommFun.GetNeedName(txt�Ա�.Text)
                
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "���� Like [2] or '('||����||')'||���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(gbytCode = 0, "����", "�����") & " Like [2]"
                End If
                
                strSQL = _
                        " Select ID,ID as ��ĿID,����,����,����," & IIf(gbytCode = 0, "����", "����� as ����") & ",˵��" & _
                        " From ��������Ŀ¼ Where Instr([3],���)>0 And (" & strSQL & ")" & _
                        IIf(strSex <> "", " And (�Ա�����=[4] Or �Ա����� is NULL)", "") & _
                        " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by ����"
                '����27613 by lesfeng 2010-01-21
                '����¼��ʱ�ж��ƥ��(����)������ѡ��,���ּ���ĸ�����ѡ��
                If zlCommFun.IsCharChinese(strInput) Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", "D", strSex, gbytCode + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                    End If
                Else
                    vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt��Ժ���.Left, txt��Ժ���.Top)
                    strInput = UCase(txt��Ժ���.Text)
                    strSex = zlCommFun.GetNeedName(txt�Ա�.Text)
                    lngTxtHeight = txt��Ժ���.Height
                    Set rsTmp = GetDiseaseCodeNew(Me, blnCancel, strInput, strSex, "D", vPoint.X, vPoint.Y, lngTxtHeight)
                    If Not rsTmp Is Nothing Then
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        End If
                    End If
                End If
                If Not rsTmp Is Nothing Then
                    '���ݿ���ֻ��һ��ƥ����Ŀ�����Ը�ƥ�����ĿΪ׼
                    txt��Ժ���.Tag = rsTmp!ID
                    txt��Ժ���.Text = "(" & rsTmp!���� & ")" & rsTmp!���� '
                    lbl��Ժ���.Tag = txt��Ժ���.Text '���ڻָ���ʾ
                Else
                    '���������ƥ����Ŀʱ���������Ϊ׼
                    txt��Ժ���.Tag = ""
                    lbl��Ժ���.Tag = txt��Ժ���.Text '���ڻָ���ʾ
                End If
            End If
            '************************************************
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt��Ժ���.Text = lbl��Ժ���.Tag And txt��Ժ���.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt��Ժ���.Text = "" Then
            txt��Ժ���.Tag = "": lbl��Ժ���.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt��Ժ���.Left, txt��Ժ���.Top)
            strInput = UCase(txt��Ժ���.Text)
            strSex = zlCommFun.GetNeedName(txt�Ա�.Text)
            lngTxtHeight = txt��Ժ���.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "D", vPoint.X, vPoint.Y, lngTxtHeight)
            If Not rsTmp Is Nothing Then
                txt��Ժ���.Tag = rsTmp!ID
                txt��Ժ���.Text = "(" & rsTmp!���� & ")" & rsTmp!����
                lbl��Ժ���.Tag = txt��Ժ���.Text '���ڻָ���ʾ
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
                End If
                If lbl��Ժ���.Tag <> "" Then txt��Ժ���.Text = lbl��Ժ���.Tag
                Call txt��Ժ���_GotFocus
                txt��Ժ���.SetFocus
            End If
        End If
    Else
        CheckInputLen txt��Ժ���, KeyAscii
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt��ҽ���_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            '����25785 by lesfeng 2009-10-20 ������������¼�����
            '************************************************
            If gintסԺ������� = 1 Then
                strInput = UCase(txt��ҽ���.Text)
                strSex = zlCommFun.GetNeedName(txt�Ա�.Text)
                
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "���� Like [2] or '('||����||')'||���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(gbytCode = 0, "����", "�����") & " Like [2]"
                End If
                
                strSQL = _
                        " Select ID,ID as ��ĿID,����,����,����," & IIf(gbytCode = 0, "����", "����� as ����") & ",˵��" & _
                        " From ��������Ŀ¼ Where Instr([3],���)>0 And (" & strSQL & ")" & _
                        IIf(strSex <> "", " And (�Ա�����=[4] Or �Ա����� is NULL)", "") & _
                        " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by ����"
                '����27613 by lesfeng 2010-01-21
                '����¼��ʱ�ж��ƥ��(����)������ѡ��,���ּ���ĸ�����ѡ��
                If zlCommFun.IsCharChinese(strInput) Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", "B", strSex, gbytCode + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                    End If
                Else
                    vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt��ҽ���.Left, txt��ҽ���.Top)
                    strInput = UCase(txt��ҽ���.Text)
                    strSex = zlCommFun.GetNeedName(txt�Ա�.Text)
                    lngTxtHeight = txt��ҽ���.Height
                    Set rsTmp = GetDiseaseCodeNew(Me, blnCancel, strInput, strSex, "B", vPoint.X, vPoint.Y, lngTxtHeight)
                    If Not rsTmp Is Nothing Then
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        End If
                    End If
                End If
                If Not rsTmp Is Nothing Then
                    '���ݿ���ֻ��һ��ƥ����Ŀ�����Ը�ƥ�����ĿΪ׼
                    txt��ҽ���.Tag = rsTmp!ID
                    txt��ҽ���.Text = "(" & rsTmp!���� & ")" & rsTmp!���� '
                    lbl��ҽ���.Tag = txt��ҽ���.Text '���ڻָ���ʾ
                Else
                    '���������ƥ����Ŀʱ���������Ϊ׼
                    txt��ҽ���.Tag = ""
                    lbl��ҽ���.Tag = txt��ҽ���.Text '���ڻָ���ʾ
                End If
            End If
            '************************************************
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt��ҽ���.Text = lbl��ҽ���.Tag And txt��ҽ���.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt��ҽ���.Text = "" Then
            txt��ҽ���.Tag = "": lbl��ҽ���.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt��ҽ���.Left, txt��ҽ���.Top)
            strInput = UCase(txt��ҽ���.Text)
            strSex = zlCommFun.GetNeedName(txt�Ա�.Text)
            lngTxtHeight = txt��ҽ���.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "B", vPoint.X, vPoint.Y, lngTxtHeight)
            If Not rsTmp Is Nothing Then
                txt��ҽ���.Tag = rsTmp!ID
                txt��ҽ���.Text = "(" & rsTmp!���� & ")" & rsTmp!����
                lbl��ҽ���.Tag = txt��ҽ���.Text '���ڻָ���ʾ
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
                End If
                If lbl��ҽ���.Tag <> "" Then txt��ҽ���.Text = lbl��ҽ���.Tag
                Call txt��ҽ���_GotFocus
                txt��ҽ���.SetFocus
            End If
        End If
    Else
        CheckInputLen txt��ҽ���, KeyAscii
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt��Ժ���_Validate(Cancel As Boolean)
    If Val(txt��Ժ���.Tag) > 0 And txt��Ժ���.Text <> lbl��Ժ���.Tag Then
        txt��Ժ���.Text = lbl��Ժ���.Tag
    ElseIf Val(txt��Ժ���.Tag) = 0 And RequestCode Then
        txt��Ժ���.Text = ""
    End If
End Sub

Private Sub txt��ҽ���_Validate(Cancel As Boolean)
    If Val(txt��ҽ���.Tag) > 0 And txt��ҽ���.Text <> lbl��ҽ���.Tag Then
        txt��ҽ���.Text = lbl��ҽ���.Tag
    ElseIf Val(txt��ҽ���.Tag) = 0 And RequestCode Then
        txt��ҽ���.Text = ""
    End If
End Sub

Private Function RequestCode() As Boolean
    RequestCode = gintסԺ������� = 2 Or (gintסԺ������� = 3 And Not IsNull(mrsPati!����))
End Function

Public Function ShowMe(frmParent As Object, ByVal lngUnit As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String) As Boolean
    Set mfrmParent = frmParent
    mlngUnit = lngUnit
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
    
    ShowMe = gblnOK
End Function

Private Function PatiInfoChange(ByVal intTYPE As Integer, ByVal strBeginDate As String, ByVal strEndDate As String) As Boolean
'����:���顢���λ�ʿ��סԺҽʦ������ҽʦ������ҽʦ�䶯�󴥷���Ϣ
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    Select Case intTYPE
    Case 13 '����䶯
        If mclsMipModule.IsConnect = True Then
            mclsXML.ClearXmlText '��������е�XML
            '--������Ϣ��װ
            '������Ϣ
            mclsXML.AppendNode "in_patient"
            'patient_id      ����id  1   N
            mclsXML.appendData "patient_id", mlng����ID, xsNumber  '����ID
            'page_id     ��ҳid  1   N
            mclsXML.appendData "page_id", mlng��ҳID, xsNumber '��ҳID
            'patient_name        ����    1   S
            mclsXML.appendData "patient_name", txt����.Text, xsString '����
            'patient_sex     �Ա�    0..1    S
            mclsXML.appendData "patient_sex", txt�Ա�.Text, xsString '�Ա�
            'in_number       סԺ��  1   S
            mclsXML.appendData "in_number", txtסԺ��.Text, xsString  'סԺ��
            mclsXML.AppendNode "in_patient", True
            
            '��ǰ���
            'current_state       ��ǰ���    1
            mclsXML.AppendNode "current_state"
            'current_area_id     ��ǰ����id  0..1    N
            mclsXML.appendData "current_area_id", Val(Nvl(mrsPati!��ǰ����ID)), xsNumber
            'current_area_title      ��ǰ����    0..1    S
            mclsXML.appendData "current_area_title", Nvl(mrsPati!��ǰ����), xsString
            'current_dept_id     ��ǰ����id  1   N
            mclsXML.appendData "current_dept_id", Val(txt����.Tag), xsNumber
            'current_dept_title      ��ǰ����    1   S
            mclsXML.appendData "current_dept_title", txt����.Text, xsString
            'current_situation       ��ǰ����    1    S
            mclsXML.appendData "current_situation", Nvl(mrsPati!��ǰ����), xsString
            mclsXML.AppendNode "current_state", True
            
            strSQL = "Select ID �䶯ID,��ʼʱ�� �䶯ʱ�� From ���˱䶯��¼ where ����ID=[1] And ��ҳId=[2] And ��ʼԭ��=[3] And NVL(���Ӵ�λ,0)=0 And ��ʼʱ��+0 between [4] And��[5]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˱䶯��¼", mlng����ID, mlng��ҳID, intTYPE, CDate(strBeginDate), CDate(strEndDate))
            '�����Ϣ
            'change_state        �����Ϣ    1
            mclsXML.AppendNode "change_state"
            'change_id       ���id  1   N
            mclsXML.appendData "change_id", rsTmp!�䶯ID, xsNumber
            'change_date     ���ʱ��    1   S
            mclsXML.appendData "change_date", Format(Nvl(rsTmp!�䶯ʱ��), "YYYY-MM-DD HH:mm:ss"), xsString
            'change_situation        �������    0..1    S
            mclsXML.appendData "change_situation", zlCommFun.GetNeedName(cbo����.Text), xsString
            'change_operator         ����Ա      1   S
            mclsXML.appendData "change_operator", UserInfo.����, xsString
            mclsXML.AppendNode "change_state", True
    
            PatiInfoChange = mclsMipModule.CommitMessage("ZLHIS_PATIENT_005", mclsXML.XmlText)
        End If
    
    Case 7 'סԺҽʦ
        If mclsMipModule.IsConnect = True Then
            mclsXML.ClearXmlText '��������е�XML
            '--������Ϣ��װ
            '������Ϣ
            mclsXML.AppendNode "in_patient"
            'patient_id      ����id  1   N
            mclsXML.appendData "patient_id", mlng����ID, xsNumber  '����ID
            'page_id     ��ҳid  1   N
            mclsXML.appendData "page_id", mlng��ҳID, xsNumber '��ҳID
            'patient_name        ����    1   S
            mclsXML.appendData "patient_name", txt����.Text, xsString '����
            'patient_sex     �Ա�    0..1    S
            mclsXML.appendData "patient_sex", txt�Ա�.Text, xsString '�Ա�
            'in_number       סԺ��  1   S
            mclsXML.appendData "in_number", txtסԺ��.Text, xsString  'סԺ��
            mclsXML.AppendNode "in_patient", True
            
            '��ǰ���
            'current_state       ��ǰ���    1
            mclsXML.AppendNode "current_state"
            'current_area_id     ��ǰ����id  0..1    N
            mclsXML.appendData "current_area_id", Val(Nvl(mrsPati!��ǰ����ID)), xsNumber
            'current_area_title      ��ǰ����    0..1    S
            mclsXML.appendData "current_area_title", Nvl(mrsPati!��ǰ����), xsString
            'current_dept_id     ��ǰ����id  1   N
            mclsXML.appendData "current_dept_id", Val(txt����.Tag), xsNumber
            'current_dept_title      ��ǰ����    1   S
            mclsXML.appendData "current_dept_title", txt����.Text, xsString
            'curren_in_doctor        סԺҽʦ    1   S
            mclsXML.appendData "curren_in_doctor", Nvl(mrsPati!סԺҽʦ), xsString
            'curren_director_doctor      ����ҽʦ    1   S
            mclsXML.appendData "curren_director_doctor", Nvl(mrsPati!����ҽʦ), xsString
            'curren_treat_doctor     ����ҽʦ    1   S
            mclsXML.appendData "curren_treat_doctor", Nvl(mrsPati!����ҽʦ), xsString
            'curren_duty_nurse       ���λ�ʿ    1   S
            mclsXML.appendData "curren_duty_nurse", Nvl(mrsPati!���λ�ʿ), xsString
            mclsXML.AppendNode "current_state", True
            
            strSQL = "Select ID �䶯ID,��ʼʱ�� �䶯ʱ�� From ���˱䶯��¼ where ����ID=[1] And ��ҳId=[2] And ��ʼԭ��=[3] And NVL(���Ӵ�λ,0)=0 And ��ʼʱ��+0 between [4] And��[5]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˱䶯��¼", mlng����ID, mlng��ҳID, intTYPE, CDate(strBeginDate), CDate(strEndDate))
            '�����Ϣ
            'change_state        �����Ϣ    1
            mclsXML.AppendNode "change_state"
            'change_id       ���id  1   N
            mclsXML.appendData "change_id", rsTmp!�䶯ID, xsNumber
            'change_date     ���ʱ��    1   S
            mclsXML.appendData "change_date", Format(Nvl(rsTmp!�䶯ʱ��), "YYYY-MM-DD HH:mm:ss"), xsString
            'change_in_doctor        סԺҽʦ    1   S
            mclsXML.appendData "change_in_doctor", zlCommFun.GetNeedName(cboסԺҽʦ.Text), xsString
            'change_director_doctor      ����ҽʦ    1   S
            mclsXML.appendData "change_director_doctor", Nvl(mrsPati!����ҽʦ), xsString
            'change_treat_doctor     ����ҽʦ    1   S
            mclsXML.appendData "change_treat_doctor", Nvl(mrsPati!����ҽʦ), xsString
            'change_duty_nurse       ���λ�ʿ    1   S
            mclsXML.appendData "change_duty_nurse", Nvl(mrsPati!���λ�ʿ), xsString
            'change_operator         ����Ա      1   S
            mclsXML.appendData "change_operator", UserInfo.����, xsString
            mclsXML.AppendNode "change_state", True
    
            PatiInfoChange = mclsMipModule.CommitMessage("ZLHIS_PATIENT_007", mclsXML.XmlText)
        End If
    Case 8 '���λ�ʿ
        If mclsMipModule.IsConnect = True Then
            mclsXML.ClearXmlText '��������е�XML
            '--������Ϣ��װ
            '������Ϣ
            mclsXML.AppendNode "in_patient"
            'patient_id      ����id  1   N
            mclsXML.appendData "patient_id", mlng����ID, xsNumber  '����ID
            'page_id     ��ҳid  1   N
            mclsXML.appendData "page_id", mlng��ҳID, xsNumber '��ҳID
            'patient_name        ����    1   S
            mclsXML.appendData "patient_name", txt����.Text, xsString '����
            'patient_sex     �Ա�    0..1    S
            mclsXML.appendData "patient_sex", txt�Ա�.Text, xsString '�Ա�
            'in_number       סԺ��  1   S
            mclsXML.appendData "in_number", txtסԺ��.Text, xsString  'סԺ��
            mclsXML.AppendNode "in_patient", True
            
            '��ǰ���
            'current_state       ��ǰ���    1
            mclsXML.AppendNode "current_state"
            'current_area_id     ��ǰ����id  0..1    N
            mclsXML.appendData "current_area_id", Val(Nvl(mrsPati!��ǰ����ID)), xsNumber
            'current_area_title      ��ǰ����    0..1    S
            mclsXML.appendData "current_area_title", Nvl(mrsPati!��ǰ����), xsString
            'current_dept_id     ��ǰ����id  1   N
            mclsXML.appendData "current_dept_id", Val(txt����.Tag), xsNumber
            'current_dept_title      ��ǰ����    1   S
            mclsXML.appendData "current_dept_title", txt����.Text, xsString
            'curren_in_doctor        סԺҽʦ    1   S
            mclsXML.appendData "curren_in_doctor", Nvl(mrsPati!סԺҽʦ), xsString
            'curren_director_doctor      ����ҽʦ    1   S
            mclsXML.appendData "curren_director_doctor", Nvl(mrsPati!����ҽʦ), xsString
            'curren_treat_doctor     ����ҽʦ    1   S
            mclsXML.appendData "curren_treat_doctor", Nvl(mrsPati!����ҽʦ), xsString
            'curren_duty_nurse       ���λ�ʿ    1   S
            mclsXML.appendData "curren_duty_nurse", Nvl(mrsPati!���λ�ʿ), xsString
            mclsXML.AppendNode "current_state", True
            
            strSQL = "Select ID �䶯ID,��ʼʱ�� �䶯ʱ�� From ���˱䶯��¼ where ����ID=[1] And ��ҳId=[2] And ��ʼԭ��=[3] And NVL(���Ӵ�λ,0)=0 And ��ʼʱ��+0 between [4] And��[5]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˱䶯��¼", mlng����ID, mlng��ҳID, intTYPE, CDate(strBeginDate), CDate(strEndDate))
            '�����Ϣ
            'change_state        �����Ϣ    1
            mclsXML.AppendNode "change_state"
            'change_id       ���id  1   N
            mclsXML.appendData "change_id", rsTmp!�䶯ID, xsNumber
            'change_date     ���ʱ��    1   S
            mclsXML.appendData "change_date", Format(Nvl(rsTmp!�䶯ʱ��), "YYYY-MM-DD HH:mm:ss"), xsString
            'change_in_doctor        סԺҽʦ    1   S
            mclsXML.appendData "change_in_doctor", zlCommFun.GetNeedName(cboסԺҽʦ.Text), xsString
            'change_director_doctor      ����ҽʦ    1   S
            mclsXML.appendData "change_director_doctor", Nvl(mrsPati!����ҽʦ), xsString
            'change_treat_doctor     ����ҽʦ    1   S
            mclsXML.appendData "change_treat_doctor", Nvl(mrsPati!����ҽʦ), xsString
            'change_duty_nurse       ���λ�ʿ    1   S
            mclsXML.appendData "change_duty_nurse", zlCommFun.GetNeedName(cbo���λ�ʿ.Text), xsString
            'change_operator         ����Ա      1   S
            mclsXML.appendData "change_operator", UserInfo.����, xsString
            mclsXML.AppendNode "change_state", True
    
            PatiInfoChange = mclsMipModule.CommitMessage("ZLHIS_PATIENT_007", mclsXML.XmlText)
        End If
    Case 11 '����ҽʦ
        If mclsMipModule.IsConnect = True Then
            mclsXML.ClearXmlText '��������е�XML
            '--������Ϣ��װ
            '������Ϣ
            mclsXML.AppendNode "in_patient"
            'patient_id      ����id  1   N
            mclsXML.appendData "patient_id", mlng����ID, xsNumber  '����ID
            'page_id     ��ҳid  1   N
            mclsXML.appendData "page_id", mlng��ҳID, xsNumber '��ҳID
            'patient_name        ����    1   S
            mclsXML.appendData "patient_name", txt����.Text, xsString '����
            'patient_sex     �Ա�    0..1    S
            mclsXML.appendData "patient_sex", txt�Ա�.Text, xsString '�Ա�
            'in_number       סԺ��  1   S
            mclsXML.appendData "in_number", txtסԺ��.Text, xsString  'סԺ��
            mclsXML.AppendNode "in_patient", True
            
            '��ǰ���
            'current_state       ��ǰ���    1
            mclsXML.AppendNode "current_state"
            'current_area_id     ��ǰ����id  0..1    N
            mclsXML.appendData "current_area_id", Val(Nvl(mrsPati!��ǰ����ID)), xsNumber
            'current_area_title      ��ǰ����    0..1    S
            mclsXML.appendData "current_area_title", Nvl(mrsPati!��ǰ����), xsString
            'current_dept_id     ��ǰ����id  1   N
            mclsXML.appendData "current_dept_id", Val(txt����.Tag), xsNumber
            'current_dept_title      ��ǰ����    1   S
            mclsXML.appendData "current_dept_title", txt����.Text, xsString
            'curren_in_doctor        סԺҽʦ    1   S
            mclsXML.appendData "curren_in_doctor", Nvl(mrsPati!סԺҽʦ), xsString
            'curren_director_doctor      ����ҽʦ    1   S
            mclsXML.appendData "curren_director_doctor", Nvl(mrsPati!����ҽʦ), xsString
            'curren_treat_doctor     ����ҽʦ    1   S
            mclsXML.appendData "curren_treat_doctor", Nvl(mrsPati!����ҽʦ), xsString
            'curren_duty_nurse       ���λ�ʿ    1   S
            mclsXML.appendData "curren_duty_nurse", Nvl(mrsPati!���λ�ʿ), xsString
            mclsXML.AppendNode "current_state", True
            
            strSQL = "Select ID �䶯ID,��ʼʱ�� �䶯ʱ�� From ���˱䶯��¼ where ����ID=[1] And ��ҳId=[2] And ��ʼԭ��=[3] And NVL(���Ӵ�λ,0)=0 And ��ʼʱ��+0 between [4] And��[5]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˱䶯��¼", mlng����ID, mlng��ҳID, intTYPE, CDate(strBeginDate), CDate(strEndDate))
            '�����Ϣ
            'change_state        �����Ϣ    1
            mclsXML.AppendNode "change_state"
            'change_id       ���id  1   N
            mclsXML.appendData "change_id", rsTmp!�䶯ID, xsNumber
            'change_date     ���ʱ��    1   S
            mclsXML.appendData "change_date", Format(Nvl(rsTmp!�䶯ʱ��), "YYYY-MM-DD HH:mm:ss"), xsString
            'change_in_doctor        סԺҽʦ    1   S
            mclsXML.appendData "change_in_doctor", zlCommFun.GetNeedName(cboסԺҽʦ.Text), xsString
            'change_director_doctor      ����ҽʦ    1   S
            mclsXML.appendData "change_director_doctor", Nvl(mrsPati!����ҽʦ), xsString
            'change_treat_doctor     ����ҽʦ    1   S
            mclsXML.appendData "change_treat_doctor", zlCommFun.GetNeedName(cbo����ҽʦ.Text), xsString
            'change_duty_nurse       ���λ�ʿ    1   S
            mclsXML.appendData "change_duty_nurse", zlCommFun.GetNeedName(cbo���λ�ʿ.Text), xsString
            'change_operator         ����Ա      1   S
            mclsXML.appendData "change_operator", UserInfo.����, xsString
            mclsXML.AppendNode "change_state", True
    
            PatiInfoChange = mclsMipModule.CommitMessage("ZLHIS_PATIENT_007", mclsXML.XmlText)
        End If
    Case 12 '����ҽʦ
        If mclsMipModule.IsConnect = True Then
            mclsXML.ClearXmlText '��������е�XML
            '--������Ϣ��װ
            '������Ϣ
            mclsXML.AppendNode "in_patient"
            'patient_id      ����id  1   N
            mclsXML.appendData "patient_id", mlng����ID, xsNumber  '����ID
            'page_id     ��ҳid  1   N
            mclsXML.appendData "page_id", mlng��ҳID, xsNumber '��ҳID
            'patient_name        ����    1   S
            mclsXML.appendData "patient_name", txt����.Text, xsString '����
            'patient_sex     �Ա�    0..1    S
            mclsXML.appendData "patient_sex", txt�Ա�.Text, xsString '�Ա�
            'in_number       סԺ��  1   S
            mclsXML.appendData "in_number", txtסԺ��.Text, xsString  'סԺ��
            mclsXML.AppendNode "in_patient", True
            
            '��ǰ���
            'current_state       ��ǰ���    1
            mclsXML.AppendNode "current_state"
            'current_area_id     ��ǰ����id  0..1    N
            mclsXML.appendData "current_area_id", Val(Nvl(mrsPati!��ǰ����ID)), xsNumber
            'current_area_title      ��ǰ����    0..1    S
            mclsXML.appendData "current_area_title", Nvl(mrsPati!��ǰ����), xsString
            'current_dept_id     ��ǰ����id  1   N
            mclsXML.appendData "current_dept_id", Val(txt����.Tag), xsNumber
            'current_dept_title      ��ǰ����    1   S
            mclsXML.appendData "current_dept_title", txt����.Text, xsString
            'curren_in_doctor        סԺҽʦ    1   S
            mclsXML.appendData "curren_in_doctor", Nvl(mrsPati!סԺҽʦ), xsString
            'curren_director_doctor      ����ҽʦ    1   S
            mclsXML.appendData "curren_director_doctor", Nvl(mrsPati!����ҽʦ), xsString
            'curren_treat_doctor     ����ҽʦ    1   S
            mclsXML.appendData "curren_treat_doctor", Nvl(mrsPati!����ҽʦ), xsString
            'curren_duty_nurse       ���λ�ʿ    1   S
            mclsXML.appendData "curren_duty_nurse", Nvl(mrsPati!���λ�ʿ), xsString
            mclsXML.AppendNode "current_state", True
            
            strSQL = "Select ID �䶯ID,��ʼʱ�� �䶯ʱ�� From ���˱䶯��¼ where ����ID=[1] And ��ҳId=[2] And ��ʼԭ��=[3] And NVL(���Ӵ�λ,0)=0 And ��ʼʱ��+0 between [4] And��[5]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˱䶯��¼", mlng����ID, mlng��ҳID, intTYPE, CDate(strBeginDate), CDate(strEndDate))
            '�����Ϣ
            'change_state        �����Ϣ    1
            mclsXML.AppendNode "change_state"
            'change_id       ���id  1   N
            mclsXML.appendData "change_id", rsTmp!�䶯ID, xsNumber
            'change_date     ���ʱ��    1   S
            mclsXML.appendData "change_date", Format(Nvl(rsTmp!�䶯ʱ��), "YYYY-MM-DD HH:mm:ss"), xsString
            'change_in_doctor        סԺҽʦ    1   S
            mclsXML.appendData "change_in_doctor", zlCommFun.GetNeedName(cboסԺҽʦ.Text), xsString
            'change_director_doctor      ����ҽʦ    1   S
            mclsXML.appendData "change_director_doctor", zlCommFun.GetNeedName(cbo����ҽʦ.Text), xsString
            'change_treat_doctor     ����ҽʦ    1   S
            mclsXML.appendData "change_treat_doctor", zlCommFun.GetNeedName(cbo����ҽʦ.Text), xsString
            'change_duty_nurse       ���λ�ʿ    1   S
            mclsXML.appendData "change_duty_nurse", zlCommFun.GetNeedName(cbo���λ�ʿ.Text), xsString
            'change_operator         ����Ա      1   S
            mclsXML.appendData "change_operator", UserInfo.����, xsString
            mclsXML.AppendNode "change_state", True
    
            PatiInfoChange = mclsMipModule.CommitMessage("ZLHIS_PATIENT_007", mclsXML.XmlText)
        End If
    End Select
End Function

Private Function CanFocus(ctlError As Control) As Boolean
    CanFocus = ctlError.Enabled And ctlError.Visible
End Function

Private Sub InitStructAddress()
'����:�����Ƿ����ýṹ����ַ��������
    Dim i As Long
    
    If gbln���ýṹ����ַ Then
        For i = PatiAddress.LBound To PatiAddress.UBound
             PatiAddress(i).Visible = True
             PatiAddress(i).ShowTown = gbln��ʾ����
        Next
        txt��ͥ��ַ.Visible = False
        cmd��ͥ��ַ.Visible = False
        txt�����ص�.Visible = False
        cmd�����ص�.Visible = False
        txt���ڵ�ַ.Visible = False
        cmd���ڵ�ַ.Visible = False
        txt����.Visible = False
        cmd����.Visible = False
        txt��ϵ�˵�ַ.Visible = False
        cmd��ϵ�˵�ַ.Visible = False
    Else
        For i = PatiAddress.LBound To PatiAddress.UBound
             PatiAddress(i).Visible = False
        Next
        
        txt��ͥ��ַ.Visible = True
        cmd��ͥ��ַ.Visible = True
        txt�����ص�.Visible = True
        cmd�����ص�.Visible = True
        txt���ڵ�ַ.Visible = True
        cmd���ڵ�ַ.Visible = True
        txt����.Visible = True
        cmd����.Visible = True
        txt��ϵ�˵�ַ.Visible = True
        cmd��ϵ�˵�ַ.Visible = True
    End If
End Sub

Private Sub EMPI_LoadPati()
'����:��EMPI�������Ĳ�����Ϣ���µ�����
    Dim rsPatiIn As ADODB.Recordset
    Dim rsPatiOut As ADODB.Recordset
    Dim str�������� As String
    Dim blnRet As Boolean
    
    If CreatePlugInOK(glngModul) Then
        '��֯���˻�����Ϣ
        Set rsPatiIn = New ADODB.Recordset
        With rsPatiIn.Fields
            .Append "����ID", adBigInt
            .Append "��ҳID", adBigInt
            .Append "�Һ�ID", adBigInt
            '-------------------------------
            .Append "�����", adVarChar, 18
            .Append "סԺ��", adVarChar, 18
            .Append "ҽ����", adVarChar, 30
            .Append "���֤��", adVarChar, 18
            .Append "����֤��", adVarChar, 20
            .Append "����", adVarChar, 100
            .Append "�Ա�", adVarChar, 4
            .Append "��������", adVarChar, 20 '���ڸ�ʽ��YYYY-MM-DD HH:MM:SS
            .Append "�����ص�", adVarChar, 100
            .Append "����", adVarChar, 30
            .Append "����", adVarChar, 20
            .Append "ѧ��", adVarChar, 10
            .Append "ְҵ", adVarChar, 80
            .Append "������λ", adVarChar, 100
            .Append "����", adVarChar, 30
            .Append "����״��", adVarChar, 4
            .Append "��ͥ�绰", adVarChar, 20
            .Append "��ϵ�˵绰", adVarChar, 20
            .Append "��λ�绰", adVarChar, 20
            .Append "��ͥ��ַ", adVarChar, 100
            .Append "��ͥ��ַ�ʱ�", adVarChar, 6
            .Append "���ڵ�ַ", adVarChar, 100
            .Append "���ڵ�ַ�ʱ�", adVarChar, 6
            .Append "��λ�ʱ�", adVarChar, 6
            .Append "��ϵ�˵�ַ", adVarChar, 100
            .Append "��ϵ�˹�ϵ", adVarChar, 30
            .Append "��ϵ������", adVarChar, 64
        End With
        rsPatiIn.CursorLocation = adUseClient
        rsPatiIn.LockType = adLockOptimistic
        rsPatiIn.CursorType = adOpenStatic
        rsPatiIn.Open

        With rsPatiIn
            .AddNew
            !����ID = mlng����ID
            !��ҳID = mlng��ҳID
            !סԺ�� = Trim(txtסԺ��.Text)
            '-Ҫ���µ��ֶ�--------------------------------------------
            !���֤�� = Trim(txt���֤��.Text)
            !���� = Trim(txt����.Text)
            !�Ա� = zlCommFun.GetNeedName(txt�Ա�.Text)
            !�����ص� = Trim(txt�����ص�.Text)
            !ѧ�� = zlCommFun.GetNeedName(cboѧ��.Text)
            !ְҵ = zlCommFun.GetNeedName(cboְҵ.Text)
            !������λ = Trim(txt��λ��ַ.Text)
            !����״�� = zlCommFun.GetNeedName(cbo����״��.Text)
            !��ͥ�绰 = Trim(txt��ͥ�绰.Text)
            !��ϵ�˵绰 = Trim(txt��ϵ�˵绰.Text)
            !��λ�绰 = Trim(txt��λ�绰.Text)
            !��ͥ��ַ = Trim(txt��ͥ��ַ.Text)
            !��ͥ��ַ�ʱ� = Trim(txt��ͥ��ַ�ʱ�.Text)
            !���ڵ�ַ = Trim(txt���ڵ�ַ.Text)
            !���ڵ�ַ�ʱ� = Trim(txt���ڵ�ַ�ʱ�.Text)
            !��λ�ʱ� = Trim(txt��λ�ʱ�.Text)
            !��ϵ�˵�ַ = Trim(txt��ϵ�˵�ַ.Text)
            !��ϵ�˹�ϵ = zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.Text)
            !��ϵ������ = Trim(txt��ϵ������.Text)
            .Update
            '-------------------------------------------------------
        End With
        
        '���ò�ѯ�ӿ�
        On Error Resume Next
        blnRet = gobjPlugIn.EMPI_QueryPatiInfo(glngSys, glngModul, rsPatiIn, rsPatiOut)
        Call zlPlugInErrH(Err, "EMPI_QueryPatiInfo")
        Err.Clear: On Error GoTo 0
        If Not blnRet Then Exit Sub
        If rsPatiOut Is Nothing Then Exit Sub
        If rsPatiOut.RecordCount = 0 Then Exit Sub
        '�ҵ����ˣ����������µ���Ϣ���µ�����
        mblnEMPI = True
        With rsPatiOut
            Call cbo.Locate(cboѧ��, !ѧ�� & "")
            Call cbo.SeekIndex(cboְҵ, !ְҵ & "")
            Call cbo.Locate(cbo����״��, !����״�� & "")
            Call cbo.Locate(cbo��ϵ�˹�ϵ, !��ϵ�˹�ϵ & "")
            
            If gbln���ýṹ����ַ Then
                PatiAddress(E_IX_�����ص�).Value = !�����ص� & ""
                PatiAddress(E_IX_��סַ).Value = !��ͥ��ַ & ""
                PatiAddress(E_IX_���ڵ�ַ).Value = !���ڵ�ַ & ""
                PatiAddress(E_IX_��ϵ�˵�ַ).Value = !��ϵ�˵�ַ & ""
            End If
            '����,�Ա�,����,�������� Ҫ�в��˻�����Ϣ�޸�Ȩ�޲��������
            txt�����ص�.Text = !�����ص� & ""
            txt��ͥ��ַ.Text = !��ͥ��ַ & ""
            txt���ڵ�ַ.Text = !���ڵ�ַ & ""
            txt��ϵ�˵�ַ.Text = !��ϵ�˵�ַ & ""
            txt���֤��.Text = !���֤�� & ""
            txt����.Text = !���� & ""
            txt��λ��ַ.Text = !������λ & ""
            txt��ͥ�绰.Text = !��ͥ�绰 & ""
            txt��ϵ�˵绰.Text = !��ϵ�˵绰 & ""
            txt��λ�绰.Text = !��λ�绰 & ""
            txt��ͥ��ַ�ʱ�.Text = !��ͥ��ַ�ʱ� & ""
            txt���ڵ�ַ�ʱ�.Text = !���ڵ�ַ�ʱ� & ""
            txt��λ�ʱ�.Text = !��λ�ʱ� & ""
            txt��ϵ������.Text = !��ϵ������ & ""
        End With
    End If
End Sub

Private Function EMPI_AddORUpdatePati(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByRef strErr As String) As Boolean
'����:���ӻ����EMPI������Ϣ
    Dim lngRet  As Long
    Dim strPlugErr As String
    Dim strTmp As String
    
    lngRet = 1 'Ĭ�ϳɹ� ���� �ϰ�zlPlug����֧�ִ˽ӿڴ����:438
    If CreatePlugInOK(glngModul) Then
        If Not mblnEMPI Then
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_AddPatiInfo(glngSys, glngModul, lngPatiID, lngPageID, 0, strErr) '1=�ɹ�;0-ʧ��
            Call zlPlugInErrH(Err, "EMPI_AddPatiInfo", strPlugErr)
            Err.Clear: On Error GoTo 0
            strTmp = "��EMPIƽ̨����������Ϣʧ�ܣ�"
        Else
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_ModifyPatiInfo(glngSys, glngModul, lngPatiID, lngPageID, 0, strErr) '1=�ɹ�;0-ʧ��
            Call zlPlugInErrH(Err, "EMPI_ModifyPatiInfo", strPlugErr)
            Err.Clear: On Error GoTo 0
            strTmp = "��EMPIƽ̨���²�����Ϣʧ�ܣ�"
        End If
        If strPlugErr <> "" Then
            strErr = strTmp & vbCrLf & strPlugErr
             Exit Function
        ElseIf lngRet = 0 Then
            strErr = strTmp & vbCrLf & strErr
            Exit Function
        End If
    End If
    
    EMPI_AddORUpdatePati = True
End Function



Private Sub PatiAddress_SetInput(Index As Integer, ByVal intLevel As Integer, rsReturn As ADODB.Recordset)
    '���ܣ������벡�˽ṹ����ַ��ʱ��,�����ʱ�
    If (Not rsReturn Is Nothing) And intLevel = 2 Then
        If Index = 3 Then
            txt��ͥ��ַ�ʱ�.Text = rsReturn!�ʱ� & ""
        End If
        If Index = 4 Then
            txt���ڵ�ַ�ʱ�.Text = rsReturn!�ʱ� & ""
        End If
    End If
End Sub

Public Sub SelectYouBian(objTextBox As TextBox)
    '���ܣ��ʱ�ѡ����
    Dim strInput As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI

    strInput = objTextBox.Text
    If strInput <> "" Then
        If zlCommFun.IsCharChinese(strInput) Then
            strSQL = strSQL & " And A.���� Like [1] "
        Else
            strSQL = strSQL & " And A.���� Like [1] "
        End If
    Else
        Exit Sub
    End If
    strSQL = "Select Rownum as ID,����,����,�ʱ�  From ���� A " & _
             "Where �ʱ� is not null " & strSQL & " Order by ����"
    vPoint = GetCoordPos(objTextBox.hWnd, 0, 0)
    Set rsTmp = zlDatabase.ShowSQLSelect(objTextBox.Parent, strSQL, 0, "�ʱ�", False, "", "", False, _
        False, True, vPoint.X, vPoint.Y, objTextBox.Height, False, False, False, UCase(strInput) & "%")
    If Not rsTmp Is Nothing Then
        objTextBox.Text = rsTmp!�ʱ� & ""
    End If
End Sub


Private Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Private Sub txtת��_GotFocus()
    zlControl.TxtSelAll txtת��
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtת��_KeyPress(KeyAscii As Integer)
    Dim vPoint As POINTAPI
    On Error GoTo errH
    If KeyAscii = 13 Then
        KeyAscii = 0
        vPoint = GetCoordPos(txtת��.Container.hWnd, txtת��.Left, txtת��.Top)
        Call GetSpcҽ�ƻ���(txtת��, Me, "ҽ�ƻ���", False, False, False, True, vPoint)
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtת��_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtת��_Validate(Cancel As Boolean)
    Dim vPoint As POINTAPI
    vPoint = GetCoordPos(txtת��.Container.hWnd, txtת��.Left, txtת��.Top)
    Call GetSpcҽ�ƻ���(txtת��, Me, "ҽ�ƻ���", False, False, False, True, vPoint)
    Exit Sub
End Sub
