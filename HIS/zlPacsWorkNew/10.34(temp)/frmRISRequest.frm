VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#2.2#0"; "zlIDKind.ocx"
Begin VB.Form frmRISRequest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���Ǽ�"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleMode       =   0  'User
   ScaleWidth      =   11506.24
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdPetitionCapture 
      Caption         =   "���뵥"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   34
      ToolTipText     =   "����(F2)"
      Top             =   7170
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   3525
      Left            =   135
      TabIndex        =   49
      Top             =   375
      Width           =   11235
      Begin VB.TextBox txt�ͼ�ҽ�� 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4995
         TabIndex        =   10
         Top             =   1485
         Width           =   2280
      End
      Begin VB.TextBox txt�ͼ쵥λ 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1425
         TabIndex        =   9
         Top             =   1485
         Width           =   2280
      End
      Begin VB.ComboBox cboҽ��2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8820
         TabIndex        =   13
         Text            =   "cboҽ��2"
         Top             =   1485
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.ComboBox cbo��ʦ�� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8820
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3120
         Width           =   2325
      End
      Begin VB.ComboBox cboִ�п��� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         ItemData        =   "frmRISRequest.frx":0000
         Left            =   4995
         List            =   "frmRISRequest.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1095
         Width           =   2280
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8820
         MaxLength       =   5
         TabIndex        =   2
         Top             =   195
         Width           =   1335
      End
      Begin VB.ComboBox cboAge 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmRISRequest.frx":0004
         Left            =   10215
         List            =   "frmRISRequest.frx":0011
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   195
         Width           =   915
      End
      Begin VB.ComboBox cbo��ʦһ 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8820
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2700
         Width           =   2325
      End
      Begin VB.TextBox txtҽ������ 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1410
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1905
         Width           =   5595
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   375
         Left            =   7020
         TabIndex        =   60
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(*)"
         Top             =   1905
         Width           =   260
      End
      Begin VB.TextBox Txt��λ���� 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1145
         Left            =   1395
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   59
         Top             =   2310
         Width           =   5895
      End
      Begin VB.ComboBox cboҽ��1 
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
         Left            =   8820
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1485
         Width           =   2325
      End
      Begin VB.ComboBox cbo�������� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmRISRequest.frx":0021
         Left            =   8820
         List            =   "frmRISRequest.frx":0023
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1080
         Width           =   2325
      End
      Begin VB.ComboBox cbo���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1425
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   2280
      End
      Begin VB.TextBox Txt���֤�� 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4995
         TabIndex        =   5
         Top             =   690
         Width           =   2280
      End
      Begin VB.TextBox Txt�绰 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8820
         TabIndex        =   6
         Top             =   660
         Width           =   2295
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1425
         MaxLength       =   64
         TabIndex        =   0
         ToolTipText     =   "����Ϊ���￨�š���������ͷΪ����ID��������סԺ�š���*������š���.���Һŵ��š���/���շѵ��ݺ�"
         Top             =   225
         Width           =   2280
      End
      Begin VB.TextBox TxtӢ���� 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1425
         TabIndex        =   4
         Top             =   675
         Width           =   2280
      End
      Begin VB.ComboBox cbo�Ա� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         ItemData        =   "frmRISRequest.frx":0025
         Left            =   4995
         List            =   "frmRISRequest.frx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   2280
      End
      Begin zlIDKind.IDKind IDKind 
         Height          =   360
         Left            =   690
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   225
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   635
         IDKindStr       =   "��|����|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;ס|סԺ��|0;��|�Һŵ�|0;��|�շѵ��ݺ�|0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   330
         Index           =   0
         Left            =   8820
         TabIndex        =   14
         Top             =   1905
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483643
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   249430019
         CurrentDate     =   38222
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   330
         Index           =   1
         Left            =   8820
         TabIndex        =   15
         Top             =   2310
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   249430019
         CurrentDate     =   38222
      End
      Begin VB.Label lbl�ͼ쵥λ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͼ쵥λ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   92
         Top             =   1530
         Width           =   1140
      End
      Begin VB.Label lbl�ͼ�ҽ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͼ�ҽ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3780
         TabIndex        =   91
         Top             =   1530
         Width           =   1140
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��鼼ʦ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7335
         TabIndex        =   88
         Top             =   3120
         Width           =   1425
      End
      Begin VB.Label labִ�п��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�п���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3765
         TabIndex        =   87
         Top             =   1125
         Width           =   1140
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��   ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   7335
         TabIndex        =   66
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��鼼ʦһ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7335
         TabIndex        =   65
         Top             =   2730
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  ����ʱ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   7335
         TabIndex        =   64
         Top             =   1935
         Width           =   1440
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  ���ʱ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   7335
         TabIndex        =   63
         Top             =   2340
         Width           =   1440
      End
      Begin VB.Label Lbl��λ���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   62
         Top             =   2340
         Width           =   1155
      End
      Begin VB.Label lblҽ������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ŀ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   135
         TabIndex        =   61
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  ����ҽ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7335
         TabIndex        =   58
         Top             =   1530
         Width           =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  �������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7335
         TabIndex        =   57
         Top             =   1125
         Width           =   1440
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   56
         Top             =   1095
         Width           =   1155
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��   ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7335
         TabIndex        =   55
         Top             =   705
         Width           =   1425
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3765
         TabIndex        =   54
         Top             =   705
         Width           =   1140
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӣ �� ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   53
         Top             =   675
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��   ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   3765
         TabIndex        =   52
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   90
         TabIndex        =   51
         Top             =   255
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   135
      TabIndex        =   67
      Top             =   3780
      Width           =   11235
      Begin VB.ComboBox cboRoom 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4995
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   195
         Width           =   2280
      End
      Begin VB.ComboBox cboDevice 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   8820
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   180
         Width           =   2310
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1380
         MaxLength       =   18
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   180
         Width           =   2280
      End
      Begin VB.Label lblItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����豸"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   7515
         TabIndex        =   70
         Top             =   210
         Width           =   1140
      End
      Begin VB.Label lblRoom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ �� ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3750
         TabIndex        =   69
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -180
         TabIndex        =   68
         Top             =   210
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   150
      TabIndex        =   41
      Top             =   0
      Width           =   11190
      Begin VB.CheckBox chk���� 
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   9570
         TabIndex        =   48
         Top             =   75
         Width           =   1545
      End
      Begin VB.TextBox txtBed 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7350
         TabIndex        =   45
         Top             =   105
         Width           =   1890
      End
      Begin VB.TextBox txtID 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   105
         Width           =   1935
      End
      Begin VB.TextBox txtPatientDept 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1350
         TabIndex        =   42
         Top             =   105
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� ʶ ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3345
         TabIndex        =   47
         Top             =   90
         Width           =   1155
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6720
         TabIndex        =   46
         Top             =   90
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���˿���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   43
         Top             =   90
         Width           =   1140
      End
   End
   Begin VB.CheckBox chkRoom 
      Caption         =   "ִ�м����(&R)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1935
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7170
      Width           =   1695
   End
   Begin VB.CommandButton CmdOK 
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
      Height          =   375
      Left            =   9015
      TabIndex        =   36
      ToolTipText     =   "����(F2)"
      Top             =   7170
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancle 
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
      Height          =   375
      Left            =   10245
      TabIndex        =   37
      Top             =   7170
      Width           =   1125
   End
   Begin MSComctlLib.ListView lvwRoom 
      Height          =   1410
      Left            =   1470
      TabIndex        =   39
      Top             =   7710
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   2487
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img16"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame frm������Ϣ 
      Height          =   2730
      Left            =   135
      TabIndex        =   71
      Top             =   4335
      Width           =   11235
      Begin VB.ComboBox cbo���ʽ 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4830
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   2265
         Width           =   1905
      End
      Begin VB.ComboBox cbo�ѱ� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2280
         Width           =   1905
      End
      Begin VB.ComboBox cbo��Ӱ�� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmRISRequest.frx":0044
         Left            =   1335
         List            =   "frmRISRequest.frx":0046
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1875
         Width           =   1905
      End
      Begin VB.TextBox Txt��Ӱ���� 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4830
         TabIndex        =   30
         Top             =   1890
         Width           =   1890
      End
      Begin VB.TextBox Txt��ӰŨ�� 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8580
         TabIndex        =   31
         Top             =   1860
         Width           =   2190
      End
      Begin VB.TextBox txt�������� 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1335
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   1455
         Width           =   9435
      End
      Begin VB.TextBox Txt��ϵ��ַ 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1335
         TabIndex        =   27
         Top             =   1035
         Width           =   9435
      End
      Begin VB.TextBox Txt�ʱ� 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8565
         TabIndex        =   26
         Top             =   630
         Width           =   2205
      End
      Begin VB.ComboBox cboְҵ 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4770
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   615
         Width           =   1830
      End
      Begin VB.ComboBox cbo���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   615
         Width           =   1830
      End
      Begin VB.TextBox Txt���� 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8565
         TabIndex        =   23
         Top             =   195
         Width           =   2205
      End
      Begin VB.TextBox Txt��� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4770
         TabIndex        =   22
         Top             =   210
         Width           =   1830
      End
      Begin MSComCtl2.DTPicker dtp�������� 
         Height          =   300
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483643
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   249430019
         CurrentDate     =   38222
      End
      Begin VB.Label Label31 
         Caption         =   "KG"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10830
         TabIndex        =   90
         Top             =   255
         Width           =   240
      End
      Begin VB.Label Label24 
         Caption         =   "CM"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6645
         TabIndex        =   89
         Top             =   255
         Width           =   315
      End
      Begin VB.Label lblCash 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   8565
         TabIndex        =   86
         Top             =   2310
         Width           =   2160
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7215
         TabIndex        =   85
         Top             =   2295
         Width           =   1170
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3570
         TabIndex        =   84
         Top             =   2280
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   45
         TabIndex        =   83
         Top             =   2295
         Width           =   1170
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "�� Ӱ ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   75
         TabIndex        =   82
         Top             =   1890
         Width           =   1155
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "��Ӱ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3270
         TabIndex        =   81
         Top             =   1905
         Width           =   1455
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "��Ӱ��Ũ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6930
         TabIndex        =   80
         Top             =   1890
         Width           =   1560
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         TabIndex        =   79
         Top             =   1470
         Width           =   1140
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ��ַ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         TabIndex        =   78
         Top             =   1050
         Width           =   1140
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         TabIndex        =   77
         Top             =   615
         Width           =   870
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְ  ҵ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3810
         TabIndex        =   76
         Top             =   645
         Width           =   870
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   45
         TabIndex        =   75
         Top             =   630
         Width           =   1170
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         TabIndex        =   74
         Top             =   225
         Width           =   870
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3810
         TabIndex        =   73
         Top             =   225
         Width           =   870
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         TabIndex        =   72
         Top             =   270
         Width           =   1140
      End
   End
   Begin VB.Label lblִ�м� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ִ �� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   40
      Top             =   7710
      Width           =   1395
   End
End
Attribute VB_Name = "frmRISRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ģ�����----�Դ�ֵ���ⲿ����
Public mstrPrivs As String          '�����ߵ�Ȩ��
Public mlngModul As Long            '��˭����
Public mlngAdviceID As Long         'ҽ��ID
Public mlngSendNo As Long           '���ͺ�
Public mintEditMode As Integer      '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
Public mlngCurDeptId As Long        '��ǰ����ID
Public mstrCur���� As String        '���ұ��������
Public mstrTechnicRoom As String    '��ǰ����ִ�м�
Public mlngResultState As Long      '�����ȡ��,0-ʧ�ܣ� 1-�Ǽǳɹ���2-�����ɹ���3-�޸ĳɹ���4-����ɹ������������Ǽ�ʱ���أ�
'Public mlngQueueWay As Long        '�Ŷӷ�ʽ
Public mblnIsAllDepartment As Boolean '�Ƿ����в���

Public mintImgCount As Integer      '��ɨ��ͼ������
Public mblnIsRelationImage As Boolean '�ж��Ƿ������ͼ���������

Private frmPetitionCap As frmPetitionCapture      'ɨ�����뵥�������

'����ģ�����------����ֵ�Ӳ�������ȡ��
Private mblnChangeNo As Boolean     '�ֹ���������
Private mblnCanOverWrite            '��������ظ�
Private mblnLike As Boolean, mlngLike As Long    '����ģ������,��������
Private mBeforeDays As Integer      '��������
Private mlngTypeSuit As Long        '��ǰ���еļ�飬ƥ����ͼ��ʽ  0-���� 1-����/סԺ��  2-����ʶ��
Private mlngGoOnReg As Long         '�����Ǽ� 0-������,1-����
Private mblnAutoPrint As Boolean    '�������Զ���ӡ���뵥
Private mlngUnicode As Long         '���߼��ű��ֲ���,1-���ּ��Ų��䣻0-������ˮ����
Private mlngUnicodeType As Long     '���ű��ֲ������,������� 0-����𲻱� 1-�����Ҳ���;
Private mlngBuildType As Long       '�������ɷ�ʽ,0-�������� 1-�����ҵ���
Private mlngAutoIncNum As Long      '�Զ���������,0-���Զ����ɣ�1-�Զ�����
Private mblnRegToCheck As Boolean   '�Ǽ�ֱ�Ӽ��
Private mblnNoshowReagent As Boolean '����ʾ��Ӱ��
Private mblnNoshowAddons As Boolean '����ʾ��������
Private mblnInputOutInfo As Boolean  '¼����Ժ��Ϣ
Private mintCheckInMode As Integer  '�Ǽ�ģʽ 1--����ģʽ��2--����ģʽ
Private mblnUseReferencePatient     'ʹ�ù�������ģʽ
Private mintCapital As Integer      'ƴ������Сд
Private mblnUseSplitter As Boolean  'ƴ�����ָ���
Private mblnAllPatientIsOutside As Boolean '���еǼǲ��˱��Ϊ����
Private mlngMoneyExeModle As Long   'Ӱ��ɼ��ķ���ִ��ģʽ 0-����ʱִ�У�1-���ʱִ�У�2-����ʱִ��

'����ģ�����------���������и�ֵ
Private mintSourceType As Integer   '������Դ 1-���� 2-סԺ 3-���� 4-���
Private mlngPatiId As Long, mlngPageID As Long  '����ID,��ҳID
Private mstrItemType As String      'Ӱ�����
Private mlngClinicID As Long        '������ĿID
'Private mstrItemIDS As String       '�շ�ϸĿID
Private mInputType As Integer       '��ȡ���˷�ʽ��0-���￨ 1-����ID 2-סԺ�� 3-����� 4-�Һŵ� 5-�շѵ��ݺ� 6-���� 7-ҽ���� 8-���֤�� 9-IC����
Private mstrExtData  As String      '�Ǽǵ�������Ŀ��λ������ ���="��λ��1;������1,������2|��λ��2;������1,������2|...<vbTab>0-����/1-����/2-����"
Private mstrAppend As String        '���="��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>..."
Private mstrOutNo As String         '�����
Private mstrCardNo As String        '���￨��
Private mstrCardPass As String      '����֤��
Private mstrChargeNo As String      '�շѵ���
Private mstrRegNo As String         '�Һŵ���
Private arrSQL() As Variant
Private mstrNextCheckNo As String     '��¼���λ�ȡ������һ������

Private mobjIcCard As Object        'IC��ʶ�����
Private WithEvents mobjIdCard As clsIDCard '���ʶ�����
Attribute mobjIdCard.VB_VarHelpID = -1
Private mobjSquareCard As Object    'һ��ͨ�������㲿��
Private oneSquardCard As TSquardCard

Private mlngBaby As Long            '�Ƿ�Ӥ����0--����Ӥ����1-9��ʾӤ�����

Private mblnIsOutSideHosp As Boolean     '�Ƿ�����Ժ����
Private mblnIsPetitionScan As Boolean    '�Ƿ��������뵥ɨ��
Private mblnIsSamePatient As Boolean     '�Ƿ������ͬ����
Private mblnUsePacsQueue As Boolean          '�Ƿ������Ŷӽк�


Private mblnExamineDoctorVerify As Boolean '�Ƿ�ʦȷ��
Private mstrExamineDoctorName As String    '��ʦ����
Private mstrExamineDoctorFst As String     '��鼼ʦһ
Private mstrExamineDoctorSed As String     '��鼼ʦ��

Private mlngInsureCheckType As Long         'ҽ������������ 0-����飬 1-����ʾ��2-��ֹ
Private mobjInsure As Object

Private mfrmParent As Form          '������




Public Function zlShowMe(frmParent As Form, ByVal blnBigFont As Boolean, Optional ByVal blnIsAllDepartment As Boolean = False, _
    Optional ByVal lngCopyAdviceId As Long, Optional ByVal lngCopySendNo As Long) As Boolean
    Set mfrmParent = frmParent
    
    mlngResultState = 0
    mblnIsRelationImage = False
    mstrTechnicRoom = ""
    
    mblnIsAllDepartment = blnIsAllDepartment
    
    Call ConfigExeDepartment(mblnIsAllDepartment)
    
    Call InitParameter
    Call InitEdit(False)  '��ʼ����������
    
    Call SetFontSize(blnBigFont)
    
    '��ȡ������Ϣ
    If mintEditMode <> 0 And mlngAdviceID <> 0 Then Call RefreshPatiInfor(mintEditMode = 2)
    
    '���ƴ��ݵĵǼ���Ϣ
    If lngCopyAdviceId <> 0 And lngCopySendNo <> 0 Then Call CopyCheck(lngCopyAdviceId, lngCopySendNo)
    
    Me.Show 1, mfrmParent
End Function



Private Sub SetFontSize(ByVal blnIsBigFont As Boolean)
    Dim objControl As Object
    Dim lngLabFontSize As Long
    Dim lngTxtFontSize As Long
    
    lngLabFontSize = IIf(blnIsBigFont, 14, 12)
    lngTxtFontSize = IIf(blnIsBigFont, 12, 10.5)
    
    For Each objControl In Me.Controls
        If TypeName(objControl) = "Label" Then
            If objControl.Name <> "Label24" And objControl.Name <> "Label31" Then
                objControl.Font.Size = lngLabFontSize
            End If
        Else
            objControl.Font.Size = lngTxtFontSize
        End If
    Next
    
    lblCash.FontSize = lngTxtFontSize
    chk����.FontSize = lngLabFontSize
End Sub




Private Sub SaveAdviceData()
'------------------------------------------------
'���ܣ�����ҽ��
'������ ��
'���أ���
'------------------------------------------------
    Dim str���ʱ�� As String, str����ʱ�� As String, curDate As String
    Dim strNO As String, lngAdviceID As Long, lngSendNO As Long
    Dim IntSeq As Integer   '����ҽ����¼.���
    Dim str��λ As String, str���� As String
    Dim i As Integer, j As Integer, strTmp���� As String, str��λ���� As String
    Dim lng��������ID As Long, lng����ID As Long, strDoctor As String
    Dim strִ�п���ID As String, lngTmpID As Long, arrAppend
    Dim rsTemp As ADODB.Recordset
    Dim lngMasSeq As Long   '����ҽ������.��¼��ţ���ҽ���е�
    Dim lngSonSeq As Long   '����ҽ������.��¼��ţ�����ҽ���еģ�Ҫ����
    

    On Error GoTo errHand
    
    curDate = To_Date(zlDatabase.Currentdate)
    str���ʱ�� = To_Date(dtp(1))
    str����ʱ�� = To_Date(dtp(0))
    
    '�²��ˣ�Ҫ��Ӳ�����Ϣ
    If mlngPatiId <= 0 Then
        '��ȡ�µĲ���ID
        mlngPatiId = zlDatabase.GetNextNo(1)
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_�ҺŲ��˲���_INSERT(1," & mlngPatiId & ",''," & _
            "'',''," & _
            "'" & Trim(txt����.Text) & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & IIf(cboAge.Visible, cboAge.Text, "") & "'," & _
            "'" & NeedName(cbo�ѱ�.Text) & "','" & NeedName(cbo���ʽ.Text) & "'," & _
            "'','" & NeedName(cbo����.Text) & "','" & NeedName(cbo����.Text) & "'," & _
            "'" & NeedName(cboְҵ.Text) & "','" & ToVarchar(Txt���֤��, 18) & "',''," & Val(Label22.Tag) & ",'','','" & ToVarchar(Txt��ϵ��ַ.Text, 50) & _
            "','" & ToVarchar(Txt�绰, 20) & "','" & ToVarchar(Txt�ʱ�, 6) & "'," & curDate & ",'','" & mstrRegNo & "'," & To_Date(dtp��������.value) & ",NULL)"
    End If
    
    '����ҽ��������
    str��λ���� = Split(mstrExtData, Chr(9))(0)
    lng��������ID = Me.cbo��������.ItemData(Me.cbo��������.ListIndex)
    strDoctor = IIf(Me.cboҽ��1.Visible, NeedName(Me.cboҽ��1.Text), NeedName(Me.cboҽ��2.Text))
    strִ�п���ID = mlngCurDeptId
    lngAdviceID = zlDatabase.GetNextId("����ҽ����¼")
    lngSendNO = zlDatabase.GetNextNo(10) 'ҽ�����ͺ�
    
    '�շѵ���Ϊ�գ���ȡ��һ���շѵ��ݺ�
    If mstrChargeNo = "" Then
        strNO = zlDatabase.GetNextNo(IIf(mintSourceType <> 2, 13, 14)) '����ȡ�շѵ��ݺ�,סԺȡ���ʵ��ݺ�
        lngMasSeq = 1
        lngSonSeq = 1
    Else    '���շѵ��ݺ�
        strNO = mstrChargeNo
        '���շѵ���,����NO��ȡ��ǰ������+1��ʼ,���ڲ���ҽ������,��ҽ�������������ٴεݼ�
        gstrSQL = "Select Max(��¼���) as ��� From ����ҽ������ Where No=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰNO������", CStr(mstrChargeNo))
        If rsTemp.EOF Then
            lngMasSeq = 1
            lngSonSeq = 1
        Else
            lngMasSeq = Nvl(rsTemp!���, 0) + 1
            lngSonSeq = lngMasSeq
        End If
    End If
    
    '������Ժ��Ϣ����Ҫ���ͼ쵥λ���ͼ�ҽ��
    If mblnInputOutInfo Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_������Ϣ�ӱ�_Update(" & mlngPatiId & ",'�ͼ쵥λ','" & Trim(Nvl(txt�ͼ쵥λ.Text)) & "'," & lngAdviceID & ")"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_������Ϣ�ӱ�_Update(" & mlngPatiId & ",'�ͼ�ҽ��','" & Trim(Nvl(txt�ͼ�ҽ��.Text)) & "'," & lngAdviceID & ")"
    End If
    
    '������ҽ��
    IntSeq = IntSeq + 1     '����ҽ����¼.��ţ�����
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & lngAdviceID & ",NULL," & _
                    IntSeq & "," & mintSourceType & "," & mlngPatiId & "," & IIf(mintSourceType = 2, mlngPageID, "NULL") & "," & _
                    "0,1,1,'D'," & mlngClinicID & ",NULL,NULL,NULL,1," & _
                    "'" & Me.txtҽ������ & "," & Decode(Txt��λ����.Tag, 1, "����", 2, "����", "����") & "ִ��:" & _
                    get��λ����(mstrExtData) & "',Null,Null,'һ����',NULL,NULL,NULL,NULL,2," & _
                    strִ�п���ID & ",3," & chk����.value & "," & str���ʱ�� & "," & str���ʱ�� & "," & _
                    IIf(Val(Me.txtPatientDept.Tag) = 0, lng��������ID, Val(Me.txtPatientDept.Tag)) & "," & lng��������ID & _
                    ",'" & strDoctor & "'," & curDate & ",'" & mstrRegNo & "',Null,Null," & Txt��λ����.Tag & ",NULL,NULL,'" & UserInfo.���� & "')"
    
    'ѭ����λ���������븽��ҽ��
    For i = 0 To UBound(Split(str��λ����, "|")) '��λ1;����1,����2,����3|��λn;����1,����2,����3---
        str��λ = Split(Split(str��λ����, "|")(i), ";")(0)
        strTmp���� = Split(Split(str��λ����, "|")(i), ";")(1)
        For j = 0 To UBound(Split(strTmp����, ","))
            IntSeq = IntSeq + 1     '����ҽ����¼.��ţ�����
            str���� = Split(strTmp����, ",")(j)
            lngTmpID = zlDatabase.GetNextId("����ҽ����¼")
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                 IntSeq & "," & mintSourceType & "," & mlngPatiId & "," & IIf(mintSourceType = 2, mlngPageID, "NULL") & "," & _
                 "0,1,1,'D'," & mlngClinicID & ",NULL,NULL,NULL,1," & _
                 "'" & Replace(Me.txtҽ������, "'", "") & "',NULL," & _
                 "'" & str��λ & "','һ����',NULL,NULL,NULL,NULL,2," & _
                 strִ�п���ID & ",3," & chk����.value & "," & str���ʱ�� & "," & str���ʱ�� & "," & _
                 IIf(Val(Me.txtPatientDept.Tag) = 0, lng��������ID, Val(Me.txtPatientDept.Tag)) & "," & lng��������ID & _
                 ",'" & strDoctor & "'," & curDate & ",'" & mstrRegNo & "',Null,'" & str���� & "'," & Txt��λ����.Tag & ",NULL,NULL,'" & UserInfo.���� & "')"
            
            '���͸���ҽ��
            '���շѵ��ݺŵ�Ϊ�ѼƷ�,�޵�Ϊδ�Ʒ�
            lngSonSeq = lngSonSeq + 1       '����ҽ������.��¼��ţ�����ҽ���еģ�Ҫ����
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            '����ҽ����ʱ�򣬲���д�״�ʱ���ĩ��ʱ�䣬������ʱ�����д
            arrSQL(UBound(arrSQL)) = "ZL_����ҽ������_Insert(" & _
                lngTmpID & "," & lngSendNO & "," & IIf(mintSourceType = 2, 2, 1) & ",'" & strNO & "'," & _
                lngSonSeq & ",1,NULL,NULL," & str����ʱ�� & ",0," & strִ�п���ID & "," & _
                IIf(mstrChargeNo = "", 0, 1) & ",0,Null,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Next
    Next
    
    '������ҽ��
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    '����ҽ����ʱ�򣬲���д�״�ʱ���ĩ��ʱ�䣬������ʱ�����д
    arrSQL(UBound(arrSQL)) = "ZL_����ҽ������_Insert(" & _
            lngAdviceID & "," & lngSendNO & "," & IIf(mintSourceType = 2, 2, 1) & ",'" & strNO & "'," & _
            lngMasSeq & ",1,NULL,NULL," & str����ʱ�� & ",0," & strִ�п���ID & "," & _
            IIf(mstrChargeNo = "", 0, 1) & ",1,Null,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
    
    '���벡��ҽ������ '     ���="��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>..."
    If mstrAppend <> "" Then
        arrAppend = Split(mstrAppend, "<Split1>")
        For i = 0 To UBound(arrAppend)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngAdviceID & _
                ",'" & Split(arrAppend(i), "<Split2>")(0) & "'," & Val(Split(arrAppend(i), "<Split2>")(1)) & "," & _
                i + 1 & "," & ZVal(Split(arrAppend(i), "<Split2>")(2)) & ",'" & Replace(Split(arrAppend(i), "<Split2>")(3), "'", "''") & "'" & _
                            IIf(i = 0, ",1", "") & ")"
        Next
    End If
    
'    '���շѵ��ݺŵģ����÷��ü�¼��ҽ���Ĺ�����ϵ
'    If mstrChargeNo <> "" Then
'        If mstrItemIDS = "" Then    'mstrItemIDS �շ�ϸĿIDΪ�գ�
'            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
'            arrSQL(UBound(arrSQL)) = "zl_���˷��ü�¼_ҽ��('" & strNO & "',1," & lngAdviceID & ")"
'        Else
'            For i = 0 To UBound(Split(mstrItemIDS, ","))
'                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
'                arrSQL(UBound(arrSQL)) = "zl_���˷��ü�¼_ҽ��('" & strNO & "',1," & lngAdviceID & "," & Split(mstrItemIDS, ",")(i) & ")"
'            Next
'        End If
'    End If
    
    
    '���շѵ��ݺŵģ����÷��ü�¼��ҽ���Ĺ�����ϵ
    If mstrChargeNo <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_���˷��ü�¼_ҽ��('" & strNO & "',1," & lngAdviceID & ")"
    End If
    
    
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ConfigExeDepartment(ByVal blnIsAllDepartment As Boolean)
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strFrom As String
    Dim lngDefaultDeptIndex As Long
    
    labִ�п���.Visible = blnIsAllDepartment
    cboִ�п���.Visible = blnIsAllDepartment
    
    Call cboִ�п���.Clear
    
    If Not blnIsAllDepartment Then Exit Sub
    
    strFrom = "1,2,3"
    strSql = " Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where B.����ID = A.ID " & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null ) " & _
        " And instr([1],','||B.�������||',')> 0 And B.�������� IN('���')" & _
        " Order by A.����"
        
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr("," & strFrom & ","))
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    lngDefaultDeptIndex = 0
    
    While Not rsData.EOF
        cboִ�п���.AddItem (Nvl(rsData!����) & "-" & Nvl(rsData!����))
        cboִ�п���.ItemData(cboִ�п���.ListCount - 1) = Nvl(rsData!ID)
        
        If Nvl(rsData!ID) = mlngCurDeptId Then lngDefaultDeptIndex = cboִ�п���.ListCount - 1
        
        rsData.MoveNext
    Wend
    
    If cboִ�п���.ListCount > 0 Then cboִ�п���.ListIndex = lngDefaultDeptIndex
End Sub


Private Sub cboAge_LostFocus()
    If Not CheckOldData(txt����, cboAge) Then Exit Sub
    If IsNumeric(txt����.Text) Then dtp��������.value = ReCalcBirth(Val(txt����.Text), cboAge.Text)
End Sub


Private Sub cbo��ʦ��_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo��ʦ��.hWnd, zlControl.CboMatchIndex(cbo��ʦ��.hWnd, KeyAscii))
End Sub

Private Sub cbo��ʦһ_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo��ʦһ.hWnd, zlControl.CboMatchIndex(cbo��ʦһ.hWnd, KeyAscii))
End Sub




Private Sub cboִ�п���_Click()
On Error GoTo ErrHandle
    mlngCurDeptId = cboִ�п���.ItemData(cboִ�п���.ListIndex)
    
    txtҽ������.Text = ""
    Txt��λ����.Text = ""
    
    Call InitParameter
    Call InitEdit(True)  '��ʼ����������
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkRoom_Click()
    If chkRoom.value = 1 Then
        Me.Height = Me.Height + lvwRoom.Height + 300
        InitRoomPati
    Else
        Me.Height = Me.Height - lvwRoom.Height - 300
    End If
End Sub

'��ҽ��ģ���У����ƹ����ļ�麯��
Public Function CheckAdviceInsure(ByVal int���� As Integer, ByVal bln���Ѷ��� As Boolean, ByVal lng����ID As Long, ByVal lng�������� As Long, _
   ByVal strIDs1 As String, ByVal strIDs2 As String, ByVal strҽ������ As String, Optional ByVal lng���˲���ID As Long) As String
'���ܣ�ҽ�������´�ҽ��ʱ��ҽ��¼��󣬶�ҽ���漰�ļƼ���Ŀ�ı��ն���������м��
'������strIDs1:ҩƷ���ĵ��շ�ϸĿID�ַ�����һ��ҽ�����磺��ù��+�����ǣ�:�շ�ϸĿID1,�շ�ϸĿID2,������
'      strIDs2 ������������Ŀ��������ĿID��һ��ҽ�����磺��Ѫ��Ŀ+��Ѫ;����:ִ�п����ַ��� ������ĿID1:ִ�п���1,������ĿID2:ִ�п���2,������
'      lng��������=1���=2סԺ
'      strҽ�����ݣ��û���ʾʱ��ʾ��ҽ������
'      bln���Ѷ���=False ��ʾ��ǰ��������飬=True �������
'���أ���ʾ��Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    
    If mlngInsureCheckType = 0 Or int���� = 0 Or Not bln���Ѷ��� Then Exit Function
    If mobjInsure.GetCapability(12, lng����ID, int����) Then Exit Function '12:support��������ҽ����Ŀ
    
    
    If strIDs1 = "" And strIDs2 = "" Then Exit Function
    
    If strIDs1 <> "" Then
        If Mid(strIDs1, 1, 1) = "," Then strIDs1 = Mid(strIDs1, 2)
        strSql = "Select Column_Value as �շ���ĿID From Table(f_Num2list([1]))"
    End If
    If strIDs2 <> "" Then
        If Mid(strIDs2, 1, 1) = "," Then strIDs2 = Mid(strIDs2, 2)
        If strIDs1 <> "" Then strSql = strSql & " Union All "
        '����û�мӲ�λ������������Ҫ��Distinct
        strSql = strSql & "Select �շ���ĿID From (" & _
                "Select Distinct C.�շ���ĿID,C.���ÿ���id" & _
                " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
                " From �����շѹ�ϵ C,Table(f_Num2list2([2])) D Where C.������ĿID=D.c1" & _
                "      And (C.���ÿ���ID is Null or C.���ÿ���ID = Nvl(D.c2,[4]) And C.������Դ = " & IIf(lng�������� = 1, 1, 2) & ")" & _
                " ) Where Nvl(���ÿ���id, 0) = Top"
    End If
    
    strSql = "Select /*+ RULE */ Distinct C.����,B.�շ�ϸĿID" & _
        " From (" & strSql & ") A,����֧����Ŀ B,�շ���ĿĿ¼ C" & _
        " Where A.�շ���ĿID=B.�շ�ϸĿID(+) And A.�շ���ĿID=C.ID" & _
        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
        " And B.����(+)=[3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckAdviceInsure", strIDs1, strIDs2, int����, lng���˲���ID)
    strSql = "": i = 0
    Do While Not rsTmp.EOF
        If IsNull(rsTmp!�շ�ϸĿID) Then
            If i = 8 Then
                strSql = strSql & vbCrLf & "�� ��"
                Exit Do
            End If
            strSql = strSql & vbCrLf & "��" & rsTmp!����
            i = i + 1
        End If
        rsTmp.MoveNext
    Loop
    If strSql <> "" Then
        CheckAdviceInsure = "��ǰ������ҽ�����ˣ���ҽ�������¼Ƽ���Ŀû�����ö�Ӧ�ı�����Ŀ��" & vbCrLf & vbCrLf & _
            "ҽ�����ݣ�" & vbCrLf & strҽ������ & vbCrLf & vbCrLf & "�Ƽ���Ŀ��" & strSql
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub cmdOK_Click()
    Dim l As Long, blnTran As Boolean, rsTmp As New ADODB.Recordset
    Dim rsMother As New ADODB.Recordset
    Dim rsPatiInfo As New ADODB.Recordset
    Dim int��¼���� As Integer     '����ҽ������.��¼���ʣ�����ҽ���ļ�¼���ʣ�1-�շѼ�¼��2-���ʼ�¼
    Dim int������� As Integer     '����ҽ������.������ʣ������סԺҽ��վ����Ϊ�������ʱ��Ϊ1,��������������ʺ�סԺ���ʣ������Ķ���Ϊ��
    Dim str������� As String
    Dim lng���ͺ� As Long
    Dim str���ݺ� As String
    Dim strҽ��IDs As String
    Dim strMsg As String
    Dim lngCurFromType As Long
    Dim lngMsgResult As Long

    On Error GoTo ErrHandle
    
    '������������Ƿ�Ϸ������Ϸ����˳�
    If ValidData = False Then Exit Sub
    
'    '�ǼǺ�ֱ�Ӽ�飬����ʱ��������ű��޸��ˣ�Ҫ�жϼ����Ƿ��ظ�
'    If (mblnRegToCheck And mintEditMode = 0) Or mintEditMode = 2 Or txt����.Tag <> txt����.Text Then
'        '�жϼ����Ƿ��ظ�����������ظ������Ҳ������Ʋ������ظ������˳�
'        If CheckNoValidate = False Then
'            Exit Sub
'        End If
'    End If
    
    
    arrSQL = Array()
    
    lngCurFromType = mintSourceType
    If mblnAllPatientIsOutside Then mintSourceType = 3  '���еǼǲ��˱��Ϊ����
    
    '�����Ӥ��ҽ��,�������޸���Ϣ���߱�������ʱ����Ϣ�ָ���ĸ�׵���Ϣ
    'mlngBaby : 0--����Ӥ����1-9��ʾӤ�����
    'mintEditMode : 0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
    If mlngBaby <> 0 And mintEditMode <> 0 Then
        gstrSQL = "SELECT B.����,B.�Ա�,B.����,B.�������� FROM ����ҽ����¼ A, ������Ϣ B " & _
                " Where A.ID=[1] And A.����ID=B.����ID"
        Set rsMother = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡĸ����Ϣ", mlngAdviceID)
        
        txt���� = Nvl(rsMother!����)
        Call SeekIndex(cbo�Ա�, Nvl(rsMother!�Ա�), True)
        If Nvl(rsMother!����) <> "" Then
            LoadOldData rsMother!����, txt����, cboAge
        Else
            ReCalcOld Format(Nvl(rsMother!��������, zlDatabase.Currentdate), "yyyy-mm-dd"), cboAge
        End If
        dtp��������.value = Format(Nvl(rsMother!��������, ReCalcBirth(txt����.Text, cboAge.Text)), "yyyy-mm-dd")
    End If
    
    ' ����ǵǼǣ��򱣴�ҽ��
    If mintEditMode = 0 Then
        If (lngCurFromType = 1 Or lngCurFromType = 2) And mlngInsureCheckType <> 0 Then
            'ֻ�д������סԺ��������ҽ�����˲Ž���ҽ��������
            gstrSQL = "select ���� from ������Ϣ Where ����ID = [1]"
            Set rsPatiInfo = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����������Ϣ", mlngPatiId)
            
            'ҽ��������
            strMsg = CheckAdviceInsure(Val(Nvl(rsPatiInfo!����)), True, mlngPatiId, mintSourceType, _
                                        "", mlngClinicID & ":" & mlngCurDeptId, "��ǰ��Ŀ")
                                        
            If strMsg <> "" Then
                If mlngInsureCheckType = 1 Then 'ֻ��ʾ
                    lngMsgResult = MsgBoxD(Me, strMsg & vbCrLf & vbCrLf & "Ҫ��������ҽ����", vbYesNo, "��ʾ��Ϣ")
                    If lngMsgResult = vbNo Then Exit Sub
                Else    '����
                    MsgBox strMsg & vbCrLf & vbCrLf & "���Ⱥ������Ա��ϵ��������ҽ�����������档", vbInformation, "��ʾ��Ϣ"
                    Exit Sub
                End If
            End If
        End If
        
        Call SaveAdviceData
    End If

    '���ǵǼ�,���������ﲡ�ˣ������ǵǼǺ�ֱ�ӱ�������Ҫ�޸Ĳ��˵���Ϣ�����ﲡ�˵���Ϣ�Ƚ϶�
    If mintEditMode <> 0 Or mintSourceType = 3 Or (mblnRegToCheck And mintEditMode = 0) Then
        gstrSQL = "Select ����,������λ,��λ�绰,��λ�ʱ�,ҽ����,IC����,����,����,���ڵ�ַ,���ڵ�ַ�ʱ�,��ϵ�����֤��,��ϵ������,��ϵ�˵绰,��ϵ�˹�ϵ,�໤�� From ������Ϣ Where ����ID = [1]"
        Set rsPatiInfo = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����������Ϣ", mlngPatiId)
        If rsPatiInfo.EOF = False Then
        
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_�ҺŲ��˲���_INSERT(3," & mlngPatiId & ",'" & IIf(mstrOutNo = 0, "", mstrOutNo) & "'," & "'" & mstrCardNo & "','" & mstrCardPass & "'," & _
                "'" & Trim(txt����.Text) & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & cboAge.Text & "'," & _
                "'" & NeedName(cbo�ѱ�.Text) & "','" & NeedName(cbo���ʽ.Text) & "','" & Nvl(rsPatiInfo!����) & "','" & NeedName(cbo����.Text) & "'," & _
                "'" & NeedName(cbo����.Text) & "','" & NeedName(cboְҵ.Text) & "','" & ToVarchar(Txt���֤��, 18) & "','" & _
                Nvl(rsPatiInfo!������λ) & "'," & Val(Label22.Tag) & ",'" & Nvl(rsPatiInfo!��λ�绰) & "','" & Nvl(rsPatiInfo!��λ�ʱ�) & "'," & _
                "'" & ToVarchar(Txt��ϵ��ַ.Text, 50) & "','" & ToVarchar(Txt�绰, 20) & "','" & ToVarchar(Txt�ʱ�, 6) & _
                "',Null,Null,Null," & To_Date(CDate(dtp��������.value)) & ",'" & Nvl(rsPatiInfo!ҽ����) & "','" & Nvl(rsPatiInfo!IC����) & _
                "'," & IIf(IsNull(rsPatiInfo!����), "Null", rsPatiInfo!����) & ",'" & Nvl(rsPatiInfo!����) & "','" & Nvl(rsPatiInfo!���ڵ�ַ) & "'," & _
                "'" & Nvl(rsPatiInfo!���ڵ�ַ�ʱ�) & "','" & Nvl(rsPatiInfo!��ϵ�����֤��) & "','" & Nvl(rsPatiInfo!��ϵ������) & "','" & Nvl(rsPatiInfo!��ϵ�˵绰) & "'," & _
                "'" & Nvl(rsPatiInfo!��ϵ�˹�ϵ) & "','" & Nvl(rsPatiInfo!�໤��) & "')"
        End If
    End If
    
    '���� �� �������޸ġ��򡡵ǼǺ�ֱ�Ӽ��
    If mintEditMode = 2 Or (mblnRegToCheck And mintEditMode = 0) Then
        '�������豸
        If cboRoom.ListCount > 0 Then   '�����ִ�м�
            If NeedName(cboRoom.list(cboRoom.ListIndex)) = "" Then 'ִ�м�δ��Ӧ����豸 , ����豸��Ӱ�����ȷ��
                InitDevice mstrItemType
            End If
        Else                          '��ִ�м�, ����豸��Ӱ�����ȷ��
            InitDevice mstrItemType
        End If
        
        '�������Լ�һ��ͨ�Ĵ���
        'ҵ���߼��ǣ�
        '1�������߼�û���շѵĲ��ܱ�������������С�δ�ɷѱ�����Ȩ�޵ģ������ڲ�ʹ��һ��ͨ���̵�����±�����
        '   ��ˢ����Ϣ��ʱ���Ѿ����Ʊ�����ȷ����ť��
        '2���Թ�������������֧�֣�
        '       ������28--����һ��ͨ�����Ѽ���ʣ����ʱ�Ƿ���Ҫ��֤
        '       ������81--ִ�к��Զ����
        '       ������163--����һ��ͨ����Ŀִ��ǰ�������շѻ��ȼ������
        '3���ȴ�����Ҫһ��ͨ����ȷ�ϵģ�����������֮һ
        '       ��1����¼����=1
        '       ��2��ִ�к��Զ����=False����¼����=2���� ����Դ<>סԺ��  ���� ����Դ=סԺ��������ʡ���
        '   ���һ��ͨ����ȷ�ϳɹ�������Ա��������һ��ͨ����ȷ�ϲ��ɹ��������С�δ�ɷѱ�����Ȩ�ޣ�Ҳ���ܱ�����
        '4���ٴ���һ��ͨ���ü�����֤�ģ�ֻ������˵ģ������ǣ�
        '       ��1����¼����=2��ִ�к��Զ����=True
        '       ��2����δ��˷���
        '
        '
        '
        gstrSQL = "Select A.��¼����,A.�������,A.���ͺ�,A.NO,B.������� from ����ҽ������ A,����ҽ����¼ B  where A.ҽ��ID=B.ID and  B.ID =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACS�������Ҽ�¼����", mlngAdviceID)
        If rsTmp.EOF = False Then
            int��¼���� = Nvl(rsTmp!��¼����, 0)
            int������� = Nvl(rsTmp!�������, 0)
            str������� = Nvl(rsTmp!�������)
            lng���ͺ� = rsTmp!���ͺ�
            str���ݺ� = Nvl(rsTmp!NO)
        End If
        
        If int��¼���� = 1 Or _
            (gblnִ�к���� = False And int��¼���� = 2 And (mintSourceType <> 2 Or (mintSourceType = 2 And int������� = 1))) Then
            
            If Not ItemHaveCash(mintSourceType, False, mlngAdviceID, 0, lng���ͺ�, str�������, str���ݺ�, int��¼����, _
                int�������, 0) Then
                If gblnִ��ǰ�Ƚ��� Then
                    '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������,�������ݺţ�����ҽ��ID��ȡ����δ�շѵ��ݻ�δ��˵ļ��ʵ�
                    '��ȡҽ��ID��
                    strҽ��IDs = mlngAdviceID
                    gstrSQL = "Select Id  from ����ҽ����¼ where ���ID = [1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ��ID��", mlngAdviceID)
                    While rsTmp.EOF = False
                        strҽ��IDs = strҽ��IDs & "," & rsTmp!ID
                        rsTmp.MoveNext
                    Wend
                    
                    If mobjSquareCard.zlSquareAffirm(Me, mlngModul, mstrPrivs, mlngPatiId, 0, False, , , strҽ��IDs) = False Then
                        MsgBoxD Me, "�ɷѲ��ɹ����ò��˻�����δ�շѵķ��ã��޷����������顣", vbOKOnly, "�ɷ�ʧ��"
                        Exit Sub
                    End If
                Else
                    '����С�δ�ɷѱ�����Ȩ�ޣ�����ʾ�Ƿ�ȷ��δ�շѿ��Ա�����
                    If InStr(mstrPrivs, "δ�ɷѱ���") > 0 Then
                        If MsgBoxD(Me, "�ò��˻�����δ�շѵķ��ã��Ƿ����������", vbYesNo, "��ʾ��Ϣ") = vbNo Then
                            Exit Sub
                        End If
                    Else
                        MsgBoxD Me, "�ò��˻�����δ�շѵķ��ã����顣", vbOKOnly, "��ʾ��Ϣ"
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        
        If gblnִ�к���� And int��¼���� = 2 Then
            'ȡ�����˵�ǰ���۷��ã���ִ�к��Զ���˻��۵�����Чʱ��
            Dim curMoney As Currency, str��� As String, str����� As String
            
            curMoney = GetAdviceMoney(mlngAdviceID, mintSourceType, str���, str�����)
            '�����ò�Ϊ0ʱ������Ƿ�һ��ͨˢ�����Ƿ���Ҫ���˱���
            If curMoney <> 0 Then
                '���˱���
                If Not FinishBillingWarn(Me, "", mlngPatiId, mlngPageID, Val(lblCash.Tag), curMoney, str���, str�����) Then
                    Exit Sub
                End If
                
                '���⣺34856
                '����һ��ͨ���������֤
                '����28--����һ��ͨ���Ѽ���ʣ����ʱ�Ƿ���Ҫ��֤
                '����81--ִ�к��Զ����
                If Val(zlDatabase.GetPara(28, glngSys)) <> 0 And gblnִ�к���� _
                    And curMoney > 0 And mintSourceType = 1 Then
                    If Not zlDatabase.PatiIdentify(Me, glngSys, mlngPatiId, curMoney) Then Exit Sub
                End If
            End If
        End If
        
        
        mstrTechnicRoom = ""
        If cboRoom.Text <> "����ʱָ��" Then mstrTechnicRoom = NeedNo(cboRoom.Text)
        
        
        '��ʼ���
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_Ӱ����_BEGIN('" & mstrTechnicRoom & "'," & Val(txt����.Text) & "," & mlngAdviceID & "," & mlngSendNo & ",'" & mstrItemType & "','" & _
            Trim(Me.txt����) & "','" & Trim(TxtӢ����.Text) & "','" & NeedName(cbo�Ա�.Text) & "','" & _
            Val(txt����.Text) & IIf(cboAge.Visible, cboAge.Text, "") & "'," & To_Date(dtp��������.value) & ",'" & ToVarchar(Txt���, 5) & "','" & _
            ToVarchar(Txt����, 5) & "',Null,Null,'" & NeedNo(cboDevice.Text) & "','" & NeedName(cbo��ʦһ.Text) & "','" & NeedName(cbo��ʦ��.Text) & "','" & txt��������.Text & "'," & To_Date(CDate(dtp(1).value)) & "," & mlngCurDeptId & ")"
        
        '����Ӱ�����¼--ִ�й���Ϊ-�ѱ���������ʱ������˵ķ���
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_Ӱ����_State(" & mlngAdviceID & "," & mlngSendNo & ",2,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCurDeptId & ")"
        
        '����ʱִ�з��û�ΪӰ��ɼ�ϵͳʱִ�з���
        If mlngModul <> G_LNG_VIDEOSTATION_MODULE Or mlngMoneyExeModle = 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_Ӱ�����ִ��(" & mlngAdviceID & "," & mlngSendNo & ",2,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngCurDeptId & ")"
        End If
        
        '��д������Ӱ��
        If Trim(cbo��Ӱ��.Text) <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_������Ӱ��_INSERT(" & mlngAdviceID & ",'" & ToVarchar(cbo��Ӱ��.Text, 30) & "','" & ToVarchar(Txt��Ӱ����.Text, 30) & "','" & ToVarchar(Txt��ӰŨ��.Text, 30) & "')"
        End If
    End If
    
    
    
    '�������޸�
    If mintEditMode = 3 Then
    
         '�������豸
        If cboRoom.ListCount > 0 Then   '�����ִ�м�
            If NeedName(cboRoom.list(cboRoom.ListIndex)) = "" Then 'ִ�м�δ��Ӧ����豸 , ����豸��Ӱ�����ȷ��
                InitDevice mstrItemType
            End If
        Else                          '��ִ�м�, ����豸��Ӱ�����ȷ��
            InitDevice mstrItemType
        End If
        
        '�޸Ĳ�����Ϣ
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_Ӱ�����¼_UPDATE(" & mlngAdviceID & ", " & mlngSendNo & "," & Val(txt����.Text) & ",'" & _
            Trim(Me.txt����.Text) & "','" & Trim(TxtӢ����.Text) & "','" & NeedName(cbo�Ա�.Text) & "','" & _
            Val(txt����.Text) & IIf(cboAge.Visible, cboAge.Text, "") & "'," & To_Date(dtp��������.value) & ",'" & ToVarchar(Txt���, 5) & "','" & _
            ToVarchar(Txt����, 5) & "','" & NeedNo(cboDevice.Text) & "','" & NeedName(cbo��ʦһ.Text) & "','" & NeedName(cbo��ʦ��.Text) & "','" & txt��������.Text & "','" & NeedNo(cboRoom.Text) & "'," & To_Date(dtp(1).value) & ")"

        
        '��д������Ӱ��
        If Trim(cbo��Ӱ��.Text) <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_������Ӱ��_INSERT(" & mlngAdviceID & ",'" & ToVarchar(cbo��Ӱ��.Text, 30) & "','" & ToVarchar(Txt��Ӱ����.Text, 30) & "','" & ToVarchar(Txt��ӰŨ��.Text, 30) & "')"
        End If
    
    End If
    
    'ִ������д��ǰ�����жϼ����Ƿ��ظ���,�ڱ����͵ǼǺ�ֱ�Ӽ��ʱ��Ҫ�жϼ����Ƿ��ظ�
    If (mblnRegToCheck And mintEditMode = 0) Or mintEditMode = 2 Or txt����.Tag <> txt����.Text Then
        If CheckNoValidate = False Then
            Exit Sub
        End If
    End If
        
    '--------------------------ִ�й��̣�д������
    gcnOracle.BeginTrans
    blnTran = True
    For l = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(l)), "д������")
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
        
    '����,��ǼǺ�ֱ�Ӽ�飬 �ĺ�������
    If mintEditMode = 2 Or (mblnRegToCheck And mintEditMode = 0) Then
        
        '���ٵ���һ��NextNO����¼��ǰ�ļ��ţ�ȷ�����ſ��Ե���
        gstrSQL = "Select NextNO([1],[2],[3],[4]) as NO From Dual"
        Call zlDatabase.OpenSQLRecord(gstrSQL, "���ǰ����", CLng(123), mlngCurDeptId, mstrItemType, Val(txt����.Text))
        
        '������ǰ���еļ�飬���չ���ƥ�����ͼ��
        gstrSQL = "Select A.���UID As ID From Ӱ����ʱ��¼ a Where a.����=[1] And a.Ӱ�����=[2]"
        Select Case mlngTypeSuit
            Case 0 '����
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(txt����.Text), mstrItemType)
            Case 1 '����/סԺ��
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Me.txtID.Text), mstrItemType)
            Case 2 '����ʶ�ţ�ҽ��ID��
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngAdviceID, mstrItemType)
        End Select
        
        '�ҵ�ƥ�����ʱͼ���¼����ͼ��ͼ���Զ�ƥ��
        If rsTmp.RecordCount = 1 Then
            gstrSQL = "ZL_Ӱ����_SET(" & mlngAdviceID & "," & mlngSendNo & ",'" & rsTmp("ID") & "')"
            zlDatabase.ExecuteProcedure gstrSQL, "��ǰ���ƥ��"
            
            mblnIsRelationImage = True
        End If
        
'        '�����ŶӶ���
'        gstrSQL = "Select ID From ���ű� Where Id = [1]"
'        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�Ŷ���Ϣ", mlngCurDeptId)
'        If rsTmp.EOF = False Then
'            If mlngQueueWay = 0 Then '��ִ�м��Ŷ�
'                mfrmParent.zlInQueue Nvl(rsTmp!ID) & ":" & NeedNo(cboRoom.Text), 1, mlngAdviceID, mlngCurDeptId, Trim(txt����.Text), mlngPatiId, NeedNo(cboRoom.Text), ""
'            Else    '�������Ŷ�
'                mfrmParent.zlInQueue mstrCur����, 1, mlngAdviceID, mlngCurDeptId, Trim(txt����.Text), mlngPatiId, "", ""
'            End If
'        End If
        
        '��ӡ���뵥
        AutoPrintApplication
    End If
    
   '�������뵥ͼ��   �ͷ� ����
   If Not frmPetitionCap Is Nothing Then
        If mintEditMode = 0 Then
            Call frmPetitionCap.subSaveImage(, mlngAdviceID)
            'ж��ɨ�����뵥�������
            Set frmPetitionCap = Nothing
        End If
   End If


    '���÷���״̬
    Select Case mintEditMode
        Case 0
            If mblnRegToCheck Then
                mlngResultState = 2
            Else
                mlngResultState = 1
            End If
        Case 1, 3
            mlngResultState = 3
        Case 2
            mlngResultState = 2
    End Select
    
    
    '����������Ǽǣ����Ҵ��ڵǼ�״̬���򲻹رմ��ڡ�
    If mlngGoOnReg = 1 And mintEditMode = 0 Then
        Call InitMvar '��ʼ��ģ�����
        Call ClearFaceData
        'InitEdit '��ʼ������ '���δ���䣬����Ҫÿ�����¼���combobox����
        Me.txt����.SetFocus
    Else
        '������ڱ���״̬,���ߵǼǺ�ֱ�ӱ����������Ƿ���ʾ��������
        If (mintEditMode = 2 Or (mblnRegToCheck And mintEditMode = 0)) And mblnUseReferencePatient = True Then
            frmReferencePatient.zlShowMe mlngAdviceID, Trim(txt����.Text), Me, False, mlngCurDeptId
        End If
        
        Unload Me
        
    End If
    
    Exit Sub
ErrHandle:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub AutoPrintApplication()
'����:�����������Զ���ӡ���뵥
Dim rsTemp As ADODB.Recordset, strBillNo As String, strExseNo As String, intExseKind As Integer

On Error GoTo errHand

    If Not mblnAutoPrint Then Exit Sub
    gstrSQL = "select NO,��¼���� from ����ҽ������ where ҽ��ID=[1] and ���ͺ�=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡNO", mlngAdviceID, mlngSendNo)
    If rsTemp.EOF Then Exit Sub
    strExseNo = rsTemp!NO: intExseKind = rsTemp!��¼����
    
    gstrSQL = "Select B.ID, B.���" & vbNewLine & _
                "From ��������Ӧ�� A, �����ļ��б� B" & vbNewLine & _
                "Where A.������Ŀid =[1] And A.Ӧ�ó��� =[2] And A.�����ļ�id = B.ID And B.���� = 7"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ݱ��", mlngClinicID, CLng(Decode(mintSourceType, 1, 1, 2, 2, 1)))
    If rsTemp.EOF Then Exit Sub
    strBillNo = "ZLCISBILL" & Format(rsTemp!���, "00000") & "-1"
    ReportOpen gcnOracle, glngSys, strBillNo, Me, "NO=" & strExseNo, "����=" & intExseKind, "ҽ��ID=" & mlngAdviceID, 2
    Exit Sub

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdPetitionCapture_Click()
On Error GoTo errHand
    
    If frmPetitionCap Is Nothing Then
        Set frmPetitionCap = New frmPetitionCapture
    End If


     '��ɨ�����뵥����
    Call frmPetitionCap.ShowPetitionCaptureWind(mstrPrivs, _
                                            mlngCurDeptId, _
                                            Nvl(Mid(cbo��������.Text, InStr(cbo��������.Text, "-") + 1, Len(cbo��������.Text))), _
                                            Nvl(Trim(txt����.Text)), _
                                            Nvl(txt����.Text), _
                                            Nvl(Mid(cbo�Ա�.Text, InStr(cbo�Ա�.Text, "-") + 1, Len(cbo�Ա�.Text))), _
                                            Nvl(txtҽ������.Text), _
                                            Nvl(Txt��λ����.Text), _
                                            IIf(InStr(mstrPrivs, "���Ǽ�") <= 0, True, False), _
                                            IIf(mintEditMode = 0, True, False), _
                                            IIf(mintEditMode = 0, 0, mlngAdviceID))

    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSel_Click()
Dim rsTmp As ADODB.Recordset
    
    With txtҽ������
        .Text = ""
        Set rsTmp = SelectDiagItem() '��ȡ��Ŀ
        If rsTmp Is Nothing Then 'ȡ����������
            '�ָ�ԭֵ
            .Text = .Tag
            zlControl.TxtSelAll txtҽ������
            .SetFocus
            Exit Sub
        Else
            If AdviceInput(rsTmp) Then '����ѡ����Ŀ���ò�λ������
                .Tag = .Text
            Else 'ȡ����λ������
                .Text = .Tag
                zlControl.TxtSelAll txtҽ������
                .SetFocus
                Exit Sub
            End If
        End If
    End With
End Sub
Private Function SelectDiagItem() As ADODB.Recordset
'ѡ������Ŀ
    Dim objPoint As RECT
    gstrSQL = "Select Distinct A.ID,A.����,A.����,nvl(A.���㵥λ,'��') as ���㵥λ,nvl(A.�걾��λ,' ') as �걾��λ," & _
                "A.�������� As ��Ŀ����,A.��� As ���ID,A.ID As ������ĿID,nvl(ִ��Ƶ��,0) As ִ��Ƶ��ID," & _
                "nvl(���㷽ʽ,0) As ���㷽ʽID,nvl(ִ�а���,0) As ִ�а���ID,nvl(�Ƽ�����,0) As �Ƽ�����ID," & _
                "nvl(ִ�п���,0) As ִ�п���ID,B.Ӱ�����" & _
              " From ������ĿĿ¼ A,Ӱ������Ŀ B,������Ŀ���� C,����ִ�п��� D" & _
              " Where A.ID=B.������ĿID AND A.ID=C.������ĿID And A.ID=D.������ĿID" & _
                    " And D.ִ�п���ID=" & mlngCurDeptId & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " & _
                    " and (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) " & _
                    " And A.������� IN(" & IIf(mintSourceType = 3, "1,2,4", mintSourceType) & ",3) " & _
                    " And Nvl(A.����Ӧ��,0)=1" & _
                    " And Nvl(A.�����Ա�,0) IN (" & IIf(cbo�Ա�.Text Like "*��*", "1,0)", "2,0)") & _
                    " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" & _
                    " And (" & zlCommFun.GetLike("A", "����", txtҽ������) & _
                            " Or " & zlCommFun.GetLike("A", "����", txtҽ������) & _
                            " Or " & zlCommFun.GetLike("C", "����", txtҽ������) & ")"
    objPoint = GetControlRect(txtҽ������.hWnd)
     Set SelectDiagItem = zlDatabase.ShowSelect(Me, gstrSQL, 0, "ѡ��������Ŀ", True, Me.txtҽ������.Text, "", True, True, True, objPoint.Left, objPoint.Top, Me.txtҽ������.Height, True, True, True)
End Function

Private Function AdviceInput(Optional rsInput As ADODB.Recordset = Nothing) As Boolean
'���ܣ����������������Ŀ(���������)����ȱʡ�Ĳ�λ������
'������rsInput=ѡ�񷵻صļ�¼��
'���أ�mstrExtData "��λ��1;������1,������2|��λ��2;������1,������2|...<vbTab>0-����/1-����/2-����"
    Dim rsTemp As ADODB.Recordset
    Dim strExtData As String, strAppend As String
    Dim blnOk As Boolean
    Dim t_Pati As TYPE_PatiInfoEx
    Dim lngHwnd As Long, int������� As Integer
    
    On Error GoTo ErrHandle
    
    If Not rsInput Is Nothing Then
        txtҽ������.Text = Replace(Replace(rsInput!����, ",", ""), "'", "") '��ʱ��ʾ
    End If
    
    With t_Pati
        .lng����ID = mlngPatiId
        If mintSourceType = 2 Then  'סԺ����д��ҳID
            .lng��ҳID = mlngPageID
        Else
            .str�Һŵ� = mstrRegNo
        End If
        .str�Ա� = NeedName(cbo�Ա�.Text)
    End With
    
    lngHwnd = IIf(mintCheckInMode = 1, Me.txt����.hWnd, Me.Txt��ϵ��ַ.hWnd)
    int������� = IIf(mintSourceType <> 2, 1, 2)
    strExtData = ""
    strAppend = mstrAppend
    
    On Error Resume Next
    '�ӿڸ��죺int����û�д��룬�ִ���0��bytUseType��ǰû�д����ִ�0
    blnOk = frmAdviceEditEx.ShowMe(Me, lngHwnd, t_Pati, 0, 0, 0, 1, int�������, , , , rsInput!������ĿID, strExtData, strAppend)

    If Not blnOk Or strExtData = "" Then Exit Function
    err.Clear
    On Error GoTo ErrHandle
    
    mstrExtData = strExtData        '���� "��λ��1;������1,������2|��λ��2;������1,������2|...<vbTab>0-����/1-����/2-����"
    mstrAppend = strAppend '     ���="��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>..."
    mlngClinicID = rsInput!������ĿID

    
    Txt��λ����.Tag = Split(mstrExtData, Chr(9))(1) 'ִ�б��
    Txt��λ����.Text = Replace(get��λ����(mstrExtData), "),", ")" & vbCrLf)
    Txt��λ����.Text = Txt��λ����.Text & vbCrLf & get������Ŀ(mstrAppend)
    
    
    mstrItemType = rsInput!Ӱ�����
    If mblnRegToCheck And Trim(txt����.Text) = "" Then txt����.Text = Next����: txt����.Tag = txt����.Text '��ʼ����
    
'    mstrItemIDS = "" '���ܸı���Ŀ,���Ե��ȸ�0
'    gstrSQL = "select �շ���ĿID FROM �����շѹ�ϵ��Where ������Ŀid=[1]"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ�ϸĿID", CLng(mlngClinicID))
'    Do Until rsTemp.EOF
'        mstrItemIDS = mstrItemIDS & "," & rsTemp!�շ���ĿID
'        rsTemp.MoveNext
'    Loop
'    mstrItemIDS = Mid(mstrItemIDS, 2)

    AdviceInput = True
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function
Private Function get������Ŀ(ByVal strAppend As String) As String
Dim i As Integer, strReturn As String
    For i = 0 To UBound(Split(strAppend, "<Split1>"))
        strReturn = strReturn & Split(Split(strAppend, "<Split1>")(i), "<Split2>")(0) & ":" & Split(Split(strAppend, "<Split1>")(i), "<Split2>")(3) & vbCrLf
    Next
    get������Ŀ = strReturn
End Function
Private Function get��λ����(ByVal strExtData As String) As String
'��:��λ��1;������1,������2|��λ��2;������1,������2|...<vbTab>0-����/1-����/2-����
'��:��λ��1(������1,������2),��λ��2(������1,������2)-----
Dim i As Integer, strReturn As String, Arr��λ
    Arr��λ = Split(Split(strExtData, Chr(9))(0), "|")
    For i = 0 To UBound(Arr��λ)
        strReturn = strReturn & "," & Split(Arr��λ(i), ";")(0) & "(" & Split(Arr��λ(i), ";")(1) & ")"
    Next
    get��λ���� = Mid(strReturn, 2)
End Function

Private Sub dtp��������_Change()
    txt����.Text = ReCalcOld(dtp��������.value, cboAge)
End Sub

Private Sub RefreshObjEnabled()
'mintEditMode '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
    Dim blnEditableState As Boolean
    
    'ȫ��״̬�µ�ͳһ����
    txtPatientDept.Enabled = False
    txtID.Enabled = False
    txtBed.Enabled = False
    Txt��λ����.Locked = True
    
    'ͨ��Ȩ�������Ʋ��˻�����Ϣ�Ƿ��ܱ��޸�
    blnEditableState = IIf(IIf(InStr(mstrPrivs, "ǿ���޸�סԺ������Ϣ") <= 0, True, False), (mintSourceType = 3), True)
    
    
    '������Ϣ��ֻ��mintSourceType = 3���������¿����޸�
    Call sutSetTxtEnable(txt����, mintSourceType = 3)
    cbo�Ա�.Enabled = mintSourceType = 3: cboAge.Enabled = mintSourceType = 3
    Call sutSetTxtEnable(txt����, mintSourceType = 3)
    dtp��������.Enabled = mintSourceType = 3
    Call sutSetTxtEnable(Txt���֤��, mintSourceType = 3)
            
    cbo�ѱ�.Enabled = blnEditableState
    cbo���ʽ.Enabled = blnEditableState: cbo����.Enabled = blnEditableState
    cboְҵ.Enabled = blnEditableState: cbo����.Enabled = blnEditableState
    
    '��ʦȷ�Ϻ󽫲��ܽ����޸�
    cbo��ʦһ.Enabled = Not mblnExamineDoctorVerify
    cbo��ʦ��.Enabled = Not mblnExamineDoctorVerify
    
    '��������Ϣһֱ�������޸�
    Call sutSetTxtEnable(Txt�绰, True)
    Call sutSetTxtEnable(Txt�ʱ�, True)
    Call sutSetTxtEnable(Txt��ϵ��ַ, True)
    
    Select Case mintEditMode
        Case 0          '0���Ǽ�
            Me.Caption = "���Ǽ�" & IIf(mlngPatiId <= 0, " �� �²��� ��", " �� ��ȡ���� ��")
            
            cboRoom.Enabled = mblnRegToCheck: cbo��ʦһ.Enabled = mblnRegToCheck: cbo��ʦ��.Enabled = mblnRegToCheck:
            cbo��Ӱ��.Enabled = mblnRegToCheck
            
            '�Ǽǵ�ʱ�����������޸�
            Call sutSetTxtEnable(txt����, True)
            cbo�Ա�.Enabled = True: cboAge.Enabled = True
            Call sutSetTxtEnable(txt����, True)
            dtp��������.Enabled = True
            Call sutSetTxtEnable(Txt���֤��, True)
            
            Call sutSetTxtEnable(TxtӢ����, True)
            Call sutSetTxtEnable(txt����, mblnRegToCheck)
            Call sutSetTxtEnable(Txt��Ӱ����, mblnRegToCheck)
            Call sutSetTxtEnable(Txt��ӰŨ��, mblnRegToCheck)
            Call sutSetTxtEnable(Txt���, mblnRegToCheck)
            Call sutSetTxtEnable(Txt����, mblnRegToCheck)
            Call sutSetTxtEnable(txt��������, mblnRegToCheck)
        Case 1          '1���ǼǺ��޸�
            Me.Caption = "�޸���Ϣ"
            
            cboRoom.Enabled = False:  cbo��ʦһ.Enabled = False: cbo��ʦ��.Enabled = False
            cbo��Ӱ��.Enabled = False: dtp(0).Enabled = False
            dtp(1).Enabled = False:  cmdSel.Enabled = False
            chk����.Enabled = False: cbo��������.Enabled = False
            cboҽ��1.Enabled = False: cboҽ��2.Enabled = False
            
            Call sutSetTxtEnable(txt�ͼ쵥λ, False)
            Call sutSetTxtEnable(txt�ͼ�ҽ��, False)
            
            Call sutSetTxtEnable(txtҽ������, False)
            Call sutSetTxtEnable(TxtӢ����, False)
            
            Call sutSetTxtEnable(txt����, False)
            Call sutSetTxtEnable(Txt��Ӱ����, False)
            Call sutSetTxtEnable(Txt��ӰŨ��, False)
            Call sutSetTxtEnable(Txt���, False)
            Call sutSetTxtEnable(Txt����, False)
            Call sutSetTxtEnable(txt��������, False)
        Case 2          '2������
            Me.Caption = "��鱨��"
            
            cbo��ʦһ.Enabled = True
            cbo��ʦ��.Enabled = True
            cbo��������.Enabled = False: cboҽ��1.Enabled = False: cboҽ��2.Enabled = False
            chk����.Enabled = False: dtp(0).Enabled = False
            dtp(1).Enabled = True: cmdSel.Enabled = False
            
            Call sutSetTxtEnable(txt�ͼ쵥λ, False)
            Call sutSetTxtEnable(txt�ͼ�ҽ��, False)
            
            Call sutSetTxtEnable(txtҽ������, False)
            
            Call sutSetTxtEnable(TxtӢ����, False)
            Call sutSetTxtEnable(txt��������, True)
        Case 3          '3���������޸�
            Me.Caption = "�޸���Ϣ"

            cboRoom.Enabled = True
            cbo��Ӱ��.Enabled = True: dtp(0).Enabled = False
            dtp(1).Enabled = True: cmdSel.Enabled = False
            chk����.Enabled = False: cbo��������.Enabled = False
            cboҽ��1.Enabled = False: cboҽ��2.Enabled = False
            
            Call sutSetTxtEnable(txt�ͼ쵥λ, False)
            Call sutSetTxtEnable(txt�ͼ�ҽ��, False)
            
            Call sutSetTxtEnable(txtҽ������, False)
            
            Call sutSetTxtEnable(TxtӢ����, False)
            Call sutSetTxtEnable(txt����, True)
            Call sutSetTxtEnable(Txt��Ӱ����, True)
            Call sutSetTxtEnable(Txt��ӰŨ��, True)
            Call sutSetTxtEnable(Txt���, True)
            Call sutSetTxtEnable(Txt����, True)
            Call sutSetTxtEnable(txt��������, True)
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Set mobjIdCard = New clsIDCard
    Call mobjIdCard.SetParent(Me.hWnd)
    
    '���������㲿��
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    '��ʼ�������㲿��
    mobjSquareCard.zlInitComponents Me, mlngModul, glngSys, gstrDBUser, gcnOracle
    '��ȡIDKindStr
    If Not mobjSquareCard Is Nothing Then
        IDKind.IDKindStr = mobjSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
        'ȡȱʡ��ˢ����ʽ
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
        '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
        '��7λ��,��ֻ��������,��Ȼȡ������
        oneSquardCard.blnȱʡ�������� = Trim(IDKind.GetKindItem(7, IDKind.GetDefaultIDKindIndex)) <> ""
        oneSquardCard.lngȱʡ�����ID = Val(IDKind.GetKindItem(IDKindItem_�����ID, IDKind.GetDefaultIDKindIndex))
    End If
    
    
    '��Ĭ��ֵ
    mlngUnicode = 0
    mlngTypeSuit = 0
    mblnLike = False
    mlngLike = 0
    mblnChangeNo = False
    mBeforeDays = 2
    If mintEditMode = 0 Then mlngBaby = 0        '����Ĭ��ֵ������Ӥ��,ֻ�еǼ�ģʽ������
    
    '��ע���ȡ�ü�鼼ʦһ ����ֵ
    mstrExamineDoctorFst = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��鼼ʦһ", "")
    mstrExamineDoctorSed = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��鼼ʦ��", "")

    
    Call ClearFaceData
End Sub

Private Sub InitParameter()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select �Ƿ�ʦȷ��,��鼼ʦ from Ӱ�����¼ where ҽ��id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    '��ʦ�Ƿ�ȷ��
    If rsTemp.RecordCount > 0 Then
        mblnExamineDoctorVerify = Nvl(rsTemp!�Ƿ�ʦȷ��, 0) = 1
        mstrExamineDoctorName = Nvl(rsTemp!��鼼ʦ)
    End If
    
    mlngGoOnReg = Val(zlDatabase.GetPara("�����Ǽ�����", glngSys, mlngModul, 0)) '�����Ǽ�
    mblnRegToCheck = (Val(GetDeptPara(mlngCurDeptId, "�ǼǺ�ֱ�Ӽ��", 0)) = 1) '�ǼǺ�ֱ�Ӽ��
    mblnAutoPrint = Val(zlDatabase.GetPara("�������Զ���ӡ���뵥", glngSys, mlngModul, 0)) '�������Զ���ӡ���뵥
    mblnAllPatientIsOutside = IIf(Val(GetDeptPara(mlngCurDeptId, "���еǼǲ��˱��Ϊ����", 0)) = 0, False, True)
    mblnUsePacsQueue = IIf(Val(GetDeptPara(mlngCurDeptId, "�����Ŷӽк�", 0)) = 0, False, True)
    
    If mlngModul = G_LNG_VIDEOSTATION_MODULE Then
    'Ӱ��ɼ�ϵͳ����Ҫ���ݲ�ͬ�ķ���ִ��ģʽ���д���
        mlngMoneyExeModle = Val(zlDatabase.GetPara("�ɼ�����ִ��ģʽ", glngSys, mlngModul, 0))
    End If
    
    mlngInsureCheckType = Val(zlDatabase.GetPara(59, glngSys))  '��ȡҽ������������
    If mlngInsureCheckType <> 0 Then
        Set mobjInsure = CreateObject("zl9Insure.clsInsure")
    End If
    
    strSql = "select ID ,����ID,������,����ֵ from Ӱ�����̲��� where ����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId)
    
    While Not rsTemp.EOF
        Select Case rsTemp!������
            Case "���߼��ű��ֲ���"
                mlngUnicode = Nvl(rsTemp!����ֵ, 0)
            Case "���ű��ֲ������"
                mlngUnicodeType = Nvl(rsTemp!����ֵ, 0)
            Case "�������ɷ�ʽ"
                mlngBuildType = Nvl(rsTemp!����ֵ, 0)
            Case "�Զ���������"
                mlngAutoIncNum = Nvl(rsTemp!����ֵ, 1)
            Case "ƥ�����ݿ���Ŀ"
                mlngTypeSuit = Nvl(rsTemp!����ֵ, 0)
            Case "�Ǽ�ʱ����ģ����������"
                mblnLike = IIf(Nvl(rsTemp!����ֵ, 0) <> 0, True, False)
                mlngLike = Abs(Nvl(rsTemp!����ֵ, 0))
            Case "�ֹ���������"
                mblnChangeNo = Nvl(rsTemp!����ֵ, 0) = 1
            Case "Ĭ�Ϲ�������"
                mBeforeDays = Val(Nvl(rsTemp!����ֵ, 2))
                If mBeforeDays > 15 Or mBeforeDays <= 0 Then
                    mBeforeDays = 2
                End If
            Case "��������ظ�"
                mblnCanOverWrite = Nvl(rsTemp!����ֵ, 0) = 1
            Case "������������"
                mblnUseReferencePatient = Nvl(rsTemp!����ֵ, 0) = 1
            Case "ƴ������Сд"
                mintCapital = Nvl(rsTemp!����ֵ, 0)
            Case "ƴ�����ָ���"
                mblnUseSplitter = Nvl(rsTemp!����ֵ, 0) = 0
        End Select
        rsTemp.MoveNext
    Wend
    
    '������Զ��������ţ�����Ҫ�����ֹ�����
    If mlngAutoIncNum = 0 Then mblnChangeNo = True
    
    Call InitFaceScheme
End Sub

Public Sub InitMvar()
    mintSourceType = 3
    mlngPatiId = 0
    mlngPageID = 0
    mstrItemType = ""
    mInputType = 6
    mstrChargeNo = ""
    mstrRegNo = ""
    mstrExtData = ""
    mlngClinicID = 0
'    mstrItemIDS = ""
    mstrAppend = ""
    mstrOutNo = 0
    mstrCardNo = ""
    mstrCardPass = ""
End Sub
Private Function Next����() As String
    Dim rsTemp As ADODB.Recordset
'mlngUnicode, mlngUnicodeType, mlngBuildType '���߼��ű��ֲ���;������� 0-����𲻱� 1-�����Ҳ���;0-�������� 1-�����ҵ���
    
    On Error GoTo errH
    
    If mlngUnicode = 1 Then '�����Ĳ��˱��ֲ���
        If mlngUnicodeType = 0 Then '0-����𲻱� 1-�����Ҳ���
            gstrSQL = "Select Max(B.����) ������" & vbNewLine & _
                        " From ����ҽ����¼ A, Ӱ�����¼ B" & vbNewLine & _
                        " Where A.����id = [1] And A.���id Is Null And A.ID = B.ҽ��id And B.Ӱ����� = [2]"
        Else
            gstrSQL = "Select Max(C.����) ������" & vbNewLine & _
                        " From ����ҽ����¼ A, Ӱ�����¼ C" & vbNewLine & _
                        " Where A.����id = [1] And A.���id Is Null And A.ִ�п���id = [3] And A.ID = C.ҽ��id"
        End If
    Else
BuildNo: '�����ɹ�����ȡ
        If mlngAutoIncNum = 0 Then Exit Function
        
        Next���� = zlDatabase.GetNextNo(123, mlngCurDeptId, mstrItemType)
        mstrNextCheckNo = Next����
        Exit Function
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������ȡ", mlngPatiId, mstrItemType, mlngCurDeptId)
    If Nvl(rsTemp!������, 0) = 0 Then
        GoTo BuildNo
    End If
    
    Next���� = CStr(rsTemp!������)
    mstrNextCheckNo = Next����
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ReCalcBirth(ByVal strOld As String, ByVal str���䵥λ As String) As String
'����:������������䵥λ���㲡�˵ĳ�������,���䵥λΪ��ʱ,�������ռٶ�Ϊ1��1��,���䵥λΪ��ʱ,�������ڼٶ�Ϊ1��
'����:��������
    Dim strTmp As String, strFormat As String, lngDays As Long
    Dim curDate As Date
    
    curDate = zlDatabase.Currentdate
    
    strTmp = "____-__-__"
    If str���䵥λ = "" Then
        strFormat = "YYYY-MM-DD"
        If strOld Like "*��*��" Or strOld Like "*��*����" Then
            strFormat = "YYYY-MM-01"
            lngDays = 365 * Val(strOld) + 30 * Val(Mid(strOld, InStr(1, strOld, "��") + 1))
        ElseIf strOld Like "*��*��" Or strOld Like "*����*��" Then
            lngDays = 30 * Val(strOld) + Val(Mid(strOld, InStr(1, strOld, "��") + 1))
        ElseIf strOld Like "*��" Or IsNumeric(strOld) Then
            strFormat = "YYYY-01-01"
            lngDays = 365 * Val(strOld)
        ElseIf strOld Like "*��" Or strOld Like "*����" Then
            strFormat = "YYYY-MM-01"
            lngDays = 30 * Val(strOld)
        ElseIf strOld Like "*��" Then
            lngDays = Val(strOld)
        End If
        If lngDays <> 0 Then strTmp = Format(DateAdd("d", lngDays * -1, curDate), strFormat)
    ElseIf strOld <> "" Then
        Select Case str���䵥λ
            Case "��"
                If Val(strOld) > 200 Then lngDays = -1
            Case "��"
                If Val(strOld) > 2400 Then lngDays = -1
            Case "��"
                If Val(strOld) > 73000 Then lngDays = -1
        End Select
        
        If lngDays = 0 Then
            strTmp = Switch(str���䵥λ = "��", "yyyy", str���䵥λ = "��", "m", str���䵥λ = "��", "d")
            strTmp = Format(DateAdd(strTmp, Val(strOld) * -1, curDate), "YYYY-MM-DD")
            
            If str���䵥λ = "��" Then
                strTmp = Format(strTmp, "YYYY-01-01")
            ElseIf str���䵥λ = "��" Then
                strTmp = Format(strTmp, "YYYY-MM-01")
            End If
        End If
    End If
    If strTmp = "____-__-__" Then strTmp = Format(curDate, "YYYY-MM-DD")
    ReCalcBirth = strTmp
End Function
Function CheckOldData(ByRef txt���� As TextBox, ByRef cbo���䵥λ As ComboBox) As Boolean
'���ܣ������������ֵ����Ч��
'���أ�
    If Not IsNumeric(txt����.Text) Then CheckOldData = True: Exit Function
    
    Select Case cbo���䵥λ.Text
        Case "��"
            If Val(txt����.Text) > 200 Then
                MsgBoxD Me, "���䲻�ܴ���200��!", vbInformation, gstrSysName
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "��"
            If Val(txt����.Text) > 2400 Then
                MsgBoxD Me, "���䲻�ܴ���2400��!", vbInformation, gstrSysName
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "��"
            If Val(txt����.Text) > 73000 Then
                MsgBoxD Me, "���䲻�ܴ���73000��!", vbInformation, gstrSysName
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                CheckOldData = False: Exit Function
            End If
    End Select
    CheckOldData = True
End Function
Private Function ReCalcOld(ByVal DateBir As Date, ByRef cbo���䵥λ As ComboBox, Optional ByVal lng����ID As Long) As String
'����:���ݳ����������¼��㲡�˵�����,�������䵥λ
'����:����,���䵥λ
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strTmp As String
 
    strSql = "Select Zl_Age_Calc([1],[2],Null) old From Dual"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng����ID, IIf(DateBir = CDate("0"), Null, DateBir))
    If Not IsNull(rsTmp!old) Then
        If rsTmp!old Like "*��" Or rsTmp!old Like "*��" Or rsTmp!old Like "*��" Then
            strTmp = Mid(rsTmp!old, 1, Len(rsTmp!old) - 1)
            If IsNumeric(strTmp) Then
                Call zlControl.CboLocate(cbo���䵥λ, Mid(rsTmp!old, Len(rsTmp!old), 1))
            Else
                strTmp = rsTmp!old
                cbo���䵥λ.ListIndex = -1
            End If
        Else
            strTmp = rsTmp!old
            If IsNumeric(strTmp) Then
                cbo���䵥λ.ListIndex = 0
            Else
                cbo���䵥λ.ListIndex = -1
            End If
        End If
    End If
    If cbo���䵥λ.ListIndex = -1 Then
        cbo���䵥λ.Visible = False
    Else
        If cbo���䵥λ.Visible = False Then cbo���䵥λ.Visible = True
    End If
    
    ReCalcOld = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPatient(strCode As String, blnCard As Boolean) As ADODB.Recordset
'���ܣ���ȡ������Ϣ������ʾ�ò��˴��ڵ�ҽ��ʱ��
    Dim strNO As String, strSeek As String
    Dim objRect As RECT, blnCancel As Boolean
    Dim lng�����ID As Long
    Dim lng����ID As Long
    Dim rsTemp As ADODB.Recordset
    
    'mInputType   1-����ID 2-סԺ�� 3-����� 4-�Һŵ� 5-�շѵ��ݺ� 6-���� 7-ҽ���� 8-���֤�� 9-IC����
    'һ��ͨ�޸�֮��mInputType�в����ھ��￨�ˣ����￨�㵽���ж�̬��֮�У�ͨ������ID��ȡ��Ϣ
    On Error GoTo errH

    mstrChargeNo = "": mstrRegNo = ""
    strSeek = strCode
    '�жϵ�ǰ����ģʽ
    Select Case IDKind.IDKind
        Case IDKind.GetKindIndex(IDKind_ҽ����)
            mInputType = 7
            strSeek = strCode
        Case IDKind.GetKindIndex(IDKind_���֤��)
            mInputType = 8
            strSeek = strCode
        Case IDKind.GetKindIndex(IDKind_IC����)
            mInputType = 9
            strSeek = strCode
        Case IDKind.GetKindIndex(IDKind_�����)
            mInputType = 3
            strSeek = Val(strCode)
        Case IDKind.GetKindIndex(IDKind_סԺ��)
            mInputType = 2
            strSeek = Val(strCode)
        Case IDKind.GetKindIndex(IDKind_�Һŵ�)
            mInputType = 4
            strSeek = strCode
        Case IDKind.GetKindIndex(IDKind_�շѵ��ݺ�)
            mInputType = 5
            strSeek = strCode
        Case Else       'ʹ��������ʱ�򣬾���ֱ��ˢ��������������ˢ���ķ���һ����
            
            If IDKind.IDKind = IDKind.GetKindIndex(IDKind_����) And blnCard = False Then
                '�����������ǲ���ˢ����
                If Left(strCode, 1) = "-" And IsNumeric(Mid(strCode, 2)) Then    '����ID
                    mInputType = 1
                    strSeek = Mid(strCode, 2)
                ElseIf Left(strCode, 1) = "+" And IsNumeric(Mid(strCode, 2)) Then 'סԺ��
                    mInputType = 2
                    strSeek = Mid(strCode, 2)
                ElseIf Left(strCode, 1) = "*" And IsNumeric(Mid(strCode, 2)) Then '�����
                    mInputType = 3
                    strSeek = Mid(strCode, 2)
                ElseIf Left(strCode, 1) = "." Then '�Һŵ�
                    mInputType = 4
                    strSeek = Mid(strCode, 2)
                ElseIf Left(strCode, 1) = "/" Then '�շѵ��ݺ�
                    mInputType = 5
                    strSeek = Mid(strCode, 2)
                ElseIf Not IsNumeric(Mid(strCode, 2)) Then '��������
                    mInputType = 6
                    strSeek = strCode
                End If
            Else
                '����̬���ֵ�ҽ�ƿ�
                '�������ģ���ȡ��صĲ���ID
                '��������,��ȡ��صĲ���ID
                '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
                '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
                '��7λ��,��ֻ��������,��Ȼȡ������
                If IDKind.IDKind = IDKind.GetKindIndex(IDKind_����) And blnCard Then
                    lng�����ID = Val(IDKind.GetKindItem(IDKindItem_�����ID, IDKind.GetDefaultIDKindIndex))
                Else
                    lng�����ID = Val(IDKind.GetKindItem(IDKindItem_�����ID))
                End If
                
                If lng�����ID <> 0 Then
                    If mobjSquareCard.zlGetPatiID(lng�����ID, strCode, False, lng����ID) = False Then
                        lng����ID = 0
                    End If
                Else
                    If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("ȫ��"), strCode, False, lng����ID) = False Then
                        lng����ID = 0
                    End If
                End If
                '��ǲ��ҷ�ʽʹ�ò���ID
                mInputType = 1
                strSeek = lng����ID
            End If
    End Select
    
    '����ID ���� �Ա� ���� ��Դ ���˿��� ��ҳid ���˿���ID ҽ�� סԺ�� ����� ��ǰ����
    '    �ѱ� ҽ�Ƹ��ʽ ���֤�� ���� ְҵ ����״�� �绰 �ʱ� ��ַ
    If mInputType = 1 Then '����ID
        gstrSQL = "select ����id,����,�Ա�,����,��������,��ԴID,��ҳID,���˿���ID,ҽ��,�����,סԺ��,���￨��,����֤��,��ǰ����,�ѱ�" & _
                        ",ҽ�Ƹ��ʽ,���֤��,����,ְҵ,����״��,�绰,�ʱ�,��ַ,��ͬ��λID, �²���" & _
                    " From(Select distinct A.����id,A.����,A.�Ա�,A.����,to_char(A.��������,'yyyy-mm-dd') ��������,Decode(A.��ǰ����id,Null,1,2) As ��ԴID,Nvl(A.סԺ����,0) As ��ҳID," & _
                        "Decode(A.��ǰ����id,Null,Nvl(B.ִ�в���ID,0),A.��ǰ����id) As ���˿���ID,nvl(B.ִ����,'') As ҽ��,A.�����,A.סԺ��,A.���￨��,A.����֤��,A.��ǰ����," & _
                        "A.�ѱ�,A.ҽ�Ƹ��ʽ,A.���֤��,A.����,A.ְҵ,A.����״��,nvl(A.��ͥ�绰,A.��ϵ�˵绰) �绰," & _
                        "nvl(A.��ͥ��ַ�ʱ�,A.��λ�ʱ�) �ʱ�,nvl(A.��ͥ��ַ,A.������λ) ��ַ,A.��ͬ��λID, 0 as �²���,B.�Ǽ�ʱ��" & _
                  " From ������Ϣ A,���˹Һż�¼ B Where A.����ID=[2] And A.����ID=B.����ID(+) And A.�����=B.�����(+) and '%'='%' " & _
                  " order by B.�Ǽ�ʱ�� desc) where rownum=1" 'Ϊ���һ��Ҳ��������������%,%��ShowSQLSelect������
    ElseIf mInputType = 2 Then 'סԺ��
        gstrSQL = "Select distinct A.����id,A.����,A.�Ա�,A.����,to_char(A.��������,'yyyy-mm-dd') ��������,Decode(A.��ǰ����id,Null,1,2) As ��ԴID,Nvl(A.סԺ����,0) As ��ҳID," & _
                        "Decode(A.��ǰ����id,Null,Nvl(B.��Ժ����ID,0),A.��ǰ����id) As ���˿���ID,B.סԺҽʦ As ҽ��,A.�����,A.סԺ��,A.���￨��,A.����֤��,A.��ǰ����," & _
                        "A.�ѱ�,A.ҽ�Ƹ��ʽ,A.���֤��,A.����,A.ְҵ,A.����״��,nvl(A.��ͥ�绰,A.��ϵ�˵绰) �绰," & _
                        "nvl(A.��ͥ��ַ�ʱ�,A.��λ�ʱ�) �ʱ�,nvl(A.��ͥ��ַ,A.������λ) ��ַ,A.��ͬ��λID, 0 as �²���" & _
                  " From ������Ϣ A,������ҳ B " & _
                  " Where A.סԺ��=[1] And A.����ID=B.����ID and A.��Ժʱ�� Is Null and '%'='%'" 'Ϊ���һ��Ҳ��������������%,%��ShowSQLSelect������
    ElseIf mInputType = 3 Then '�����,��������ŵģ���Ϊ�����ﲡ��
        gstrSQL = "select ����id,����,�Ա�,����,��������,��ԴID,��ҳID,���˿���ID,ҽ��,�����,סԺ��,���￨��,����֤��,��ǰ����,�ѱ�" & _
                        ",ҽ�Ƹ��ʽ,���֤��,����,ְҵ,����״��,�绰,�ʱ�,��ַ,��ͬ��λID, �²���" & _
                    " From (Select distinct A.����id,A.����,A.�Ա�,A.����,to_char(A.��������,'yyyy-mm-dd') ��������,Decode(A.��ǰ����id,Null,1,2) As ��ԴID,Nvl(A.סԺ����,0) As ��ҳID," & _
                        "Decode(A.��ǰ����id,Null,Nvl(B.ִ�в���ID,0),A.��ǰ����id) As ���˿���ID,B.ִ���� As ҽ��,A.�����,A.סԺ��,A.���￨��,A.����֤��,A.��ǰ����," & _
                        "A.�ѱ�,A.ҽ�Ƹ��ʽ,A.���֤��,A.����,A.ְҵ,A.����״��,nvl(A.��ͥ�绰,A.��ϵ�˵绰) �绰," & _
                        "nvl(A.��ͥ��ַ�ʱ�,A.��λ�ʱ�) �ʱ�,nvl(A.��ͥ��ַ,A.������λ) ��ַ,B.�Ǽ�ʱ��,A.��ͬ��λID, 0 as �²���" & _
                        " From ������Ϣ A,���˹Һż�¼ B Where A.�����=[1] And A.����ID=B.����ID(+) And A.�����=B.�����(+) and B.��¼����=1 and B.��¼״̬=1 Order By B.�Ǽ�ʱ�� Desc)" & _
                    " where Rownum=1 and '%'='%'" 'Ϊ���һ��Ҳ��������������%,%��ShowSQLSelect������
    ElseIf mInputType = 4 Then '�Һŵ�
        strNO = GetFullNO(strSeek, 12)
        txt����.Text = strNO
'        mstrRegNo = strNO
        gstrSQL = "Select Distinct A.����id, A.����, A.�Ա�, A.����, To_Char(A.��������, 'yyyy-mm-dd') ��������, Decode(Nvl(A.��Ժ, 0), 0, 1, 2) As ��Դid," & vbNewLine & _
                    "                Nvl(A.סԺ����, 0) As ��ҳid, Nvl(B.ִ�в���id, B.ת�����id) As ���˿���id, B.ִ���� As ҽ��, Nvl(A.�����, B.�����) �����, A.סԺ��," & vbNewLine & _
                    "                A.���￨��, A.����֤��, A.��ǰ����, A.�ѱ�, A.ҽ�Ƹ��ʽ, A.���֤��, A.����, A.ְҵ, A.����״��, Nvl(A.��ͥ�绰, A.��ϵ�˵绰) �绰," & vbNewLine & _
                    "                Nvl(A.��ͥ��ַ�ʱ�, A.��λ�ʱ�) �ʱ�, Nvl(A.��ͥ��ַ, A.������λ) ��ַ, A.��ͬ��λid, 0 as �²���" & vbNewLine & _
                    "From ������Ϣ A, ���˹Һż�¼ B" & vbNewLine & _
                    "Where B.NO = [3] And B.����id = A.����id and B.��¼����=1 and B.��¼״̬=1 and '%'='%'"  'Ϊ���һ��Ҳ��������������%,%��ShowSQLSelect������
    ElseIf mInputType = 5 Then '�շѵ��ݺţ������շѵ��ݺŵģ���Ϊ�����ﲡ��
        strNO = GetFullNO(strSeek, 13)
        txt����.Text = strNO
        mstrChargeNo = strNO
        
        '������ü�¼��NO=���˹Һż�¼��NO������ʹ���շѵ��ݺ���ȡ���˵�ʱ��ͬʱ��¼�Һŵ���
        '���û�йҺŵ�Ϊ�գ���ͨ���շѵ��ݺ���ȡ���Ǽǵ����ﲡ�ˣ�������ҽ�����ݡ�
'        mstrRegNo = strNO
        
        gstrSQL = "Select Distinct Nvl(A.����id, 0) ����id, Nvl(A.����, B.����) ����, Nvl(A.�Ա�, B.�Ա�) �Ա�, Nvl(A.����, B.����) ����," & vbNewLine & _
                    "                To_Char(A.��������, 'yyyy-mm-dd') ��������, Decode(Nvl(A.��Ժ, 0), 0, 1, 2) As ��Դid, Nvl(A.סԺ����, 0) As ��ҳid," & vbNewLine & _
                    "                Nvl(B.��������id, B.���˿���id) As ���˿���id, Nvl(B.������, B.ִ����) As ҽ��, Nvl(A.�����, B.��ʶ��) �����, A.סԺ��, A.���￨��, A.����֤��," & vbNewLine & _
                    "                A.��ǰ����, A.�ѱ�, A.ҽ�Ƹ��ʽ, A.���֤��, A.����, A.ְҵ, A.����״��, Nvl(A.��ͥ�绰, A.��ϵ�˵绰) �绰, Nvl(A.��ͥ��ַ�ʱ�, A.��λ�ʱ�) �ʱ�," & vbNewLine & _
                    "                Nvl(A.��ͥ��ַ, A.������λ) ��ַ, A.��ͬ��λid, 0 as �²���" & vbNewLine & _
                    "From ������Ϣ A, ������ü�¼ B" & vbNewLine & _
                    "Where B.NO = [3] And B.��¼���� = 1 And B.��¼״̬ = 1 And B.ִ��״̬ <>9 And B.����id = A.����id(+) And '%' = '%'" 'Ϊ���һ��Ҳ��������������%,%��ShowSQLSelect������
    ElseIf mInputType = 6 Then '��������
            gstrSQL = "Select distinct A.����id,A.����,A.�Ա�,A.����,to_char(A.��������,'yyyy-mm-dd') ��������,Decode(A.��ǰ����id,Null,1,2) As ��ԴID,Nvl(A.סԺ����,0) As ��ҳID," & _
                        "NVL(A.��ǰ����id,0) As ���˿���ID,'' As ҽ��,A.�����,A.סԺ��,A.���￨��,A.����֤��,A.��ǰ����," & _
                        "A.�ѱ�,A.ҽ�Ƹ��ʽ,A.���֤��,A.����,A.ְҵ,A.����״��,nvl(A.��ͥ�绰,A.��ϵ�˵绰) �绰," & _
                        "nvl(A.��ͥ��ַ�ʱ�,A.��λ�ʱ�) �ʱ�,nvl(A.��ͥ��ַ,A.������λ) ��ַ,A.��ͬ��λID, 0 as �²���" & _
                " From ������Ϣ A where " & IIf(mblnLike = False, "A.����=[1]", IIf(mlngLike = 0, "instr(A.����,[1])>0", "A.�Ǽ�ʱ�� Between sysdate-" & mlngLike & " and sysdate and instr(A.����,[1])>0"))
    
    ElseIf mInputType = 7 Then 'ҽ����
        gstrSQL = "Select distinct A.����id,A.����,A.�Ա�,A.����,to_char(A.��������,'yyyy-mm-dd') ��������,Decode(A.��ǰ����id,Null,1,2) As ��ԴID,Nvl(A.סԺ����,0) As ��ҳID," & _
                        "NVL(A.��ǰ����id,0) As ���˿���ID,'' As ҽ��,A.�����,A.סԺ��,A.���￨��,A.����֤��,A.��ǰ����," & _
                        "A.�ѱ�,A.ҽ�Ƹ��ʽ,A.���֤��,A.����,A.ְҵ,A.����״��,nvl(A.��ͥ�绰,A.��ϵ�˵绰) �绰," & _
                        "nvl(A.��ͥ��ַ�ʱ�,A.��λ�ʱ�) �ʱ�,nvl(A.��ͥ��ַ,A.������λ) ��ַ,A.��ͬ��λID, 0 as �²���" & _
                  " From ������Ϣ A Where A.ҽ����=[1] and '%'='%'" 'Ϊ���һ��Ҳ��������������%,%��ShowSQLSelect������
    ElseIf mInputType = 8 Then '���֤��
        gstrSQL = "Select distinct A.����id,A.����,A.�Ա�,A.����,to_char(A.��������,'yyyy-mm-dd') ��������,Decode(A.��ǰ����id,Null,1,2) As ��ԴID,Nvl(A.סԺ����,0) As ��ҳID," & _
                        "NVL(A.��ǰ����id,0) As ���˿���ID,'' As ҽ��,A.�����,A.סԺ��,A.���￨��,A.����֤��,A.��ǰ����," & _
                        "A.�ѱ�,A.ҽ�Ƹ��ʽ,A.���֤��,A.����,A.ְҵ,A.����״��,nvl(A.��ͥ�绰,A.��ϵ�˵绰) �绰," & _
                        "nvl(A.��ͥ��ַ�ʱ�,A.��λ�ʱ�) �ʱ�,nvl(A.��ͥ��ַ,A.������λ) ��ַ,A.��ͬ��λID, 0 as �²���" & _
                  " From ������Ϣ A Where A.���֤��=[1] and '%'='%'" 'Ϊ���һ��Ҳ��������������%,%��ShowSQLSelect������
    ElseIf mInputType = 9 Then 'IC����
        gstrSQL = "Select distinct A.����id,A.����,A.�Ա�,A.����,to_char(A.��������,'yyyy-mm-dd') ��������,Decode(A.��ǰ����id,Null,1,2) As ��ԴID,Nvl(A.סԺ����,0) As ��ҳID," & _
                        "NVL(A.��ǰ����id,0) As ���˿���ID,'' As ҽ��,A.�����,A.סԺ��,A.���￨��,A.����֤��,A.��ǰ����," & _
                        "A.�ѱ�,A.ҽ�Ƹ��ʽ,A.���֤��,A.����,A.ְҵ,A.����״��,nvl(A.��ͥ�绰,A.��ϵ�˵绰) �绰," & _
                        "nvl(A.��ͥ��ַ�ʱ�,A.��λ�ʱ�) �ʱ�,nvl(A.��ͥ��ַ,A.������λ) ��ַ,A.��ͬ��λID, 0 as �²���" & _
                  " From ������Ϣ A Where A.IC����=[1] and '%'='%'" 'Ϊ���һ��Ҳ��������������%,%��ShowSQLSelect������
    End If
    
    gstrSQL = gstrSQL & " Union " & _
                "Select 0 ����ID,'�²���' ����,'δ֪' �Ա�,'' ����,null ��������,3 As ��ԴID,0 As ��ҳID," & _
                        "0 As ���˿���ID,'' As ҽ��,0 as �����,0 as סԺ��,'' as ���￨��,'' ����֤��,'' as ��ǰ����," & _
                        "'' as �ѱ�,'' as ҽ�Ƹ��ʽ,'' as ���֤��,'��' as ����,'' as  ְҵ,'δ��' as ����״��,'' �绰,'' �ʱ�,'' ��ַ,0 ��ͬ��λID, 1 as �²���" & _
             " From dual where '%'='%'"
    gstrSQL = "select RowNum as ID,����id,����,�Ա�,����,��������,��ԴID,��ҳID,���˿���ID,ҽ��,�����," & _
                "סԺ��,���￨��,����֤��,��ǰ����,�ѱ�,ҽ�Ƹ��ʽ,���֤��,����,ְҵ,����״��,�绰,�ʱ�,��ַ,��ͬ��λID" & _
                " From (" & gstrSQL & ") Order by �²��� asc,����ID desc"
    objRect = GetControlRect(txt����.hWnd)
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ������ͬ����", CStr(strSeek), Val(strSeek), strNO)
    mblnIsSamePatient = IIf(rsTemp.RecordCount > 1, True, False)
    
    Set GetPatient = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "�鲡����Ϣ", False, "����ID", "", False, False, True, objRect.Left, objRect.Top, txt����.Height, blnCancel, True, False, CStr(strSeek), Val(strSeek), strNO)
    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function GetDictData(strDict As String) As ADODB.Recordset
'���ܣ���ָ�����ֵ��ж�ȡ����
'������strDict=�ֵ��Ӧ�ı���
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
        
    strSql = "Select ����,nvl(����,'δ֪') as ����,Nvl(ȱʡ��־,0) as ȱʡ From " & strDict & " Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ" & strDict)
    
    If Not rsTmp.EOF Then Set GetDictData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub InitDoctors(ByVal lng����ID As Long)
'���ܣ���ȡ��ǰ���������а�����������Ա
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    
    strSql = "Select /*+RULE*/" & vbNewLine & _
                "Distinct b.id,b.����, Upper(b.����) As ����" & vbNewLine & _
                " From ������Ա a, ��Ա�� b, ��Ա����˵�� c" & vbNewLine & _
                " Where a.����id = [1] And a.��Աid = b.Id And b.Id = c.��Աid And c.��Ա���� = 'ҽ��' And" & vbNewLine & _
                "      (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null)" & vbNewLine & _
                " Order By ���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
    
    If mblnIsOutSideHosp Then
        cboҽ��2.Clear
        If Not rsTmp.EOF Then
            Do Until rsTmp.EOF
                cboҽ��2.AddItem rsTmp!���� & "-" & rsTmp!����
                If rsTmp!ID = UserInfo.ID Then cboҽ��2.ListIndex = cboҽ��2.NewIndex
                rsTmp.MoveNext
            Loop
            If cboҽ��2.ListCount > 0 And cboҽ��2.ListIndex = -1 Then cboҽ��2.ListIndex = 0
            cboҽ��2.Enabled = True
        End If
    Else
        cboҽ��1.Clear
        If Not rsTmp.EOF Then
            Do Until rsTmp.EOF
                cboҽ��1.AddItem rsTmp!���� & "-" & rsTmp!����
                If rsTmp!ID = UserInfo.ID Then cboҽ��1.ListIndex = cboҽ��1.NewIndex
                rsTmp.MoveNext
            Loop
            If cboҽ��1.ListCount > 0 And cboҽ��1.ListIndex = -1 Then cboҽ��1.ListIndex = 0
            cboҽ��1.Enabled = True
        End If
    End If
    
End Sub
Private Sub InitInput()
    Dim i As Integer, strInput As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select ID ,����ID,����ֵ from Ӱ�����̲��� where ����ID = [1] and ������ = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId, CStr("�������"))
    If Not rsTemp.EOF Then
        strInput = Nvl(rsTemp!����ֵ)
    End If
    
    For i = 0 To UBound(Split(strInput, "|"))
        Select Case Split(strInput, "|")(i)
            Case "Ӣ����"
                TxtӢ����.TabStop = False
            Case "�Ա�"
                cbo�Ա�.TabStop = False
            Case "����"
                txt����.TabStop = False
                cboAge.TabStop = False
            Case "��������"
                dtp��������.TabStop = False
            Case "���"
                Txt���.TabStop = False
            Case "����"
                Txt����.TabStop = False
            Case "�ѱ�"
                cbo�ѱ�.TabStop = False
            Case "���ʽ"
                cbo���ʽ.TabStop = False
            Case "���֤��"
                Txt���֤��.TabStop = False
            Case "����"
                cbo����.TabStop = False
            Case "ְҵ"
                cboְҵ.TabStop = False
            Case "����"
                cbo����.TabStop = False
            Case "�绰"
                Txt�绰.TabStop = False
            Case "�ʱ�"
                Txt�ʱ�.TabStop = False
            Case "��ַ"
                Txt��ϵ��ַ.TabStop = False
            Case "ִ�м�"
                cboRoom.TabStop = False
            Case "����"
                chk����.TabStop = False
            Case "����豸"
                cboDevice.TabStop = False
            Case "����"
                txt����.TabStop = False
            Case "����ʱ��"
                dtp(0).TabStop = False
            Case "���ʱ��"
                dtp(1).TabStop = False
            Case "��Ӱ��"
                cbo��Ӱ��.TabStop = False
                Txt��Ӱ����.TabStop = False
                Txt��ӰŨ��.TabStop = False
            Case "��鼼ʦ"
                cbo��ʦһ.TabStop = False
            Case "��鼼ʦ��"
                cbo��ʦ��.TabStop = False
        End Select
    Next
End Sub
Public Sub InitRoomPati()
Dim rsTemp As ADODB.Recordset, i As Integer, lst As ListItem
    On Error GoTo errH:
    If cboRoom.ListCount < 1 Then 'û��ִ�м�
        Exit Sub
    End If
    With lvwRoom
        With .ColumnHeaders
            .Clear
            .Add , , "ִ�м�", 2800
            .Add , , "��������", 1400, 1
            .Add , , "�ѱ���", 1400, 1
            .Add , , "������", 1400, 1
        End With
        .ListItems.Add , , "Temp"
        .ListItems.Clear
    End With
    
    gstrSQL = "Select Count(ID) ����, ִ�м�, ״̬" & vbNewLine & _
                "From (Select /*+rule*/" & vbNewLine & _
                "        A.ID, Decode(Nvl(B.ִ�м�, ''), '', 'δ��ִ�м�', B.ִ�м�) ִ�м�," & vbNewLine & _
                "        Decode(Nvl(D.����id, 0), 0, '������', '�ѱ���') ״̬" & vbNewLine & _
                "       From ����ҽ����¼ A, ����ҽ������ B, Ӱ�����¼ C, ����ҽ������ D" & vbNewLine & _
                "       Where A.���id Is Null And A.ִ�п���id = [1] And" & vbNewLine & _
                "             A.��ʼִ��ʱ�� Between To_Date(To_Char(Sysdate-" & (mBeforeDays - 1) & ", 'yyyy-mm-dd'), 'yyyy-mm-dd hh24:mi:ss') And Sysdate And" & vbNewLine & _
                "             A.ID = B.ҽ��id And B.ҽ��id = C.ҽ��id And B.���ͺ� = C.���ͺ� And A.ID = D.ҽ��id(+))" & vbNewLine & _
                "Group By ִ�м�, ״̬" & vbNewLine & _
                "Order By ִ�м�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡִ�м䲡�����", mlngCurDeptId)

    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    For i = 0 To cboRoom.ListCount - 1
        Set lst = lvwRoom.ListItems.Add(, "_" & NeedNo(cboRoom.list(i)), NeedNo(cboRoom.list(i)))
        rsTemp.Filter = "ִ�м�='" & NeedNo(cboRoom.list(i)) & "'"
        Do Until rsTemp.EOF
            If rsTemp!״̬ = "�ѱ���" Then
                lst.SubItems(2) = rsTemp!����
            Else
                lst.SubItems(3) = rsTemp!����
            End If
            lst.SubItems(1) = Val(Nvl(lst.SubItems(1), 0)) + rsTemp!����
            rsTemp.MoveNext
        Loop
    Next
    
    rsTemp.Filter = "ִ�м�='δ��ִ�м�'"
    If Not rsTemp.EOF Then Set lst = lvwRoom.ListItems.Add(, "_δ��ִ�м�", "δ��ִ�м�")
    Do Until rsTemp.EOF
        If rsTemp!״̬ = "�ѱ���" Then
            lst.SubItems(2) = rsTemp!����
        Else
            lst.SubItems(3) = rsTemp!����
        End If
        lst.SubItems(1) = Val(Nvl(lst.SubItems(1), 0)) + rsTemp!����
        rsTemp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub InitFaceScheme()
    '��ȡ����
    mblnNoshowReagent = Val(zlDatabase.GetPara("����ʾ��Ӱ��", glngSys, mlngModul, 0)) = 1
    mblnNoshowAddons = Val(zlDatabase.GetPara("����ʾ��������", glngSys, mlngModul, 0)) = 1
    mblnInputOutInfo = Val(zlDatabase.GetPara("¼����Ժ��Ϣ", glngSys, mlngModul, 0)) = 1
    mintCheckInMode = Val(zlDatabase.GetPara("�Ǽ�ģʽ", glngSys, mlngModul, 2))
    
    mblnIsPetitionScan = IIf(Val(GetDeptPara(mlngCurDeptId, "�������뵥ɨ��", 1)) = 1, True, False)   '��ȡ�������뵥ɨ�����
    Me.cmdPetitionCapture.Visible = mblnIsPetitionScan
    
    If mintCheckInMode <> 1 Then mintCheckInMode = 2
    
    If Not mblnInputOutInfo Then
        lbl�ͼ쵥λ.Visible = False
        txt�ͼ쵥λ.Visible = False
        lbl�ͼ�ҽ��.Visible = False
        txt�ͼ�ҽ��.Visible = False
        
        lblҽ������.Top = 1530
        txtҽ������.Top = 1515
        cmdSel.Top = 1500
        Lbl��λ����.Top = 2040
        Txt��λ����.Top = 2010
        Txt��λ����.Height = 1400
    End If
    
    '��Ϊ������������Ӱ�����Ϸ���ʾ�������ȴ���������
    If mblnNoshowAddons And Label29.Visible = True Then '����ʾ�������ߣ��Ҹ��������Ѿ�����ʾ����ر���ʾ��������
        Label29.Visible = False: txt��������.Visible = False: txt��������.Enabled = False
        '��������ؼ���λ��
        Label26.Top = Label26.Top - 350: cbo��Ӱ��.Top = cbo��Ӱ��.Top - 370
        Label27.Top = Label27.Top - 350: Txt��Ӱ����.Top = Txt��Ӱ����.Top - 370
        Label28.Top = Label28.Top - 350: Txt��ӰŨ��.Top = Txt��ӰŨ��.Top - 370
        Label1.Top = Label1.Top - 370: cbo�ѱ�.Top = cbo�ѱ�.Top - 370
        Label13.Top = Label13.Top - 370: cbo���ʽ.Top = cbo���ʽ.Top - 370
        Label12.Top = Label12.Top - 370: lblCash.Top = lblCash.Top - 370
        frm������Ϣ.Height = frm������Ϣ.Height - 400
        CmdOK.Top = CmdOK.Top - 400: CmdCancle.Top = CmdOK.Top: chkRoom.Top = CmdOK.Top: cmdPetitionCapture.Top = CmdOK.Top
        lvwRoom.Top = lvwRoom.Top - 400: lblִ�м�.Top = lvwRoom.Top
        Me.Height = Me.Height - 400
    End If
    
    If mblnNoshowReagent And Label26.Visible = True Then    '����ʾ��Ӱ��������Ӱ���Ѿ�����ʾ����ر���Ӱ������ʾ
        Label26.Visible = False: Label27.Visible = False: Label28.Visible = False
        cbo��Ӱ��.Visible = False: cbo��Ӱ��.Enabled = False
        Txt��ӰŨ��.Visible = False: Txt��ӰŨ��.Visible = False
        Txt��Ӱ����.Visible = False: Txt��Ӱ����.Visible = False
        '��������Ŀؼ�λ��
        Label1.Top = Label1.Top - 370: cbo�ѱ�.Top = cbo�ѱ�.Top - 370
        Label13.Top = Label13.Top - 370: cbo���ʽ.Top = cbo���ʽ.Top - 370
        Label12.Top = Label12.Top - 370: lblCash.Top = lblCash.Top - 370
        frm������Ϣ.Height = frm������Ϣ.Height - 400
        CmdOK.Top = CmdOK.Top - 400: CmdCancle.Top = CmdOK.Top: chkRoom.Top = CmdOK.Top: cmdPetitionCapture.Top = CmdOK.Top
        lvwRoom.Top = lvwRoom.Top - 400: lblִ�м�.Top = lvwRoom.Top
        Me.Height = Me.Height - 400
    End If
    
    If mintCheckInMode = 1 Then     '����ģʽ
        frm������Ϣ.Visible = False
        CmdOK.Top = CmdOK.Top - frm������Ϣ.Height: CmdCancle.Top = CmdOK.Top: chkRoom.Top = CmdOK.Top: cmdPetitionCapture.Top = CmdOK.Top
        lvwRoom.Top = lvwRoom.Top - frm������Ϣ.Height: lblִ�м�.Top = lvwRoom.Top
        Me.Height = Me.Height - frm������Ϣ.Height
    End If
End Sub

Private Sub ClearFaceData()
    Dim curDate As Date
    
    txt����.Text = "":      txt����.Tag = ""
    TxtӢ����.Text = "":    TxtӢ����.Tag = ""
    txt����.Text = "":      cboAge.Visible = True
    Txt���.Text = "":      Txt����.Text = ""
    Txt���֤��.Text = "":  Txt�绰.Text = ""
    Txt�ʱ�.Text = "":      Txt��ϵ��ַ = ""
    txtPatientDept.Text = "":  txtID.Text = ""
    txtBed.Text = ""
    txt����.Text = "":    txt����.Tag = ""
    Txt��Ӱ����.Text = "":  Txt��ӰŨ��.Text = ""
    txtҽ������.Text = "":  txtҽ������.Tag = ""
    Txt��λ����.Text = "":  Txt��λ����.Tag = ""
    
    curDate = zlDatabase.Currentdate
    
    dtp��������.value = Format(curDate, "yyyy-mm-dd")
    dtp(0).value = curDate
    dtp(1).value = curDate
    
    cboAge.ListIndex = 0
    
End Sub

Private Sub InitEdit(ByVal blnIsChangeDept As Boolean)
    Dim strSql As String, rsTmp As ADODB.Recordset, i As Integer
    
    On Error GoTo DBError
    
    If Not blnIsChangeDept Then
        cboAge.ListIndex = 0
        
        '�Ա�
        Set rsTmp = GetDictData("�Ա�")
        cbo�Ա�.Clear
        If Not rsTmp Is Nothing Then
            For i = 1 To rsTmp.RecordCount
                cbo�Ա�.AddItem rsTmp!���� & "-" & rsTmp!����
                If rsTmp!ȱʡ = 1 Then
                    cbo�Ա�.ItemData(cbo�Ա�.NewIndex) = 1
                    cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
                End If
                rsTmp.MoveNext
            Next
        End If
        
        '���ݴ����ͼ���������жϸı䰴ť������
        If mintEditMode > 0 Then cmdPetitionCapture.Caption = IIf(mintImgCount = 0, "���뵥", "���뵥(" & mintImgCount & "��)")
        
        '�ѱ�
        Set rsTmp = GetDictData("�ѱ�")
        cbo�ѱ�.Clear
        If Not rsTmp Is Nothing Then
            Do Until rsTmp.EOF
                cbo�ѱ�.AddItem rsTmp!���� & "-" & rsTmp!����
                If rsTmp!ȱʡ = 1 Then
                    cbo�ѱ�.ItemData(cbo�ѱ�.NewIndex) = 1
                    cbo�ѱ�.ListIndex = cbo�ѱ�.NewIndex
                End If
                rsTmp.MoveNext
            Loop
        End If
        
        '���ʽ
        Set rsTmp = GetDictData("ҽ�Ƹ��ʽ")
        cbo���ʽ.Clear
        If Not rsTmp Is Nothing Then
            Do Until rsTmp.EOF
                cbo���ʽ.AddItem rsTmp!���� & "-" & rsTmp!����
                If rsTmp!ȱʡ = 1 Then
                    cbo���ʽ.ItemData(cbo���ʽ.NewIndex) = 1
                    cbo���ʽ.ListIndex = cbo���ʽ.NewIndex
                End If
                rsTmp.MoveNext
            Loop
        End If
        
        '����
        Set rsTmp = GetDictData("����")
        cbo����.Clear
        If Not rsTmp Is Nothing Then
            Do Until rsTmp.EOF
                cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
                If rsTmp!ȱʡ = 1 Then
                    cbo����.ItemData(cbo����.NewIndex) = 1
                    cbo����.ListIndex = cbo����.NewIndex
                End If
                rsTmp.MoveNext
            Loop
        End If
        
        'ְҵ
        Set rsTmp = GetDictData("ְҵ")
        cboְҵ.Clear
        If Not rsTmp Is Nothing Then
            Do Until rsTmp.EOF
                cboְҵ.AddItem rsTmp!���� & "-" & rsTmp!����
                If rsTmp!ȱʡ = 1 Then
                    cboְҵ.ItemData(cboְҵ.NewIndex) = 1
                    cboְҵ.ListIndex = cboְҵ.NewIndex
                End If
                rsTmp.MoveNext
            Loop
        End If
        
        '����״��
        Set rsTmp = GetDictData("����״��")
        cbo����.Clear
        If Not rsTmp Is Nothing Then
            Do Until rsTmp.EOF
                cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
                If rsTmp!ȱʡ = 1 Then
                    cbo����.ItemData(cbo����.NewIndex) = 1
                    cbo����.ListIndex = cbo����.NewIndex
                End If
                rsTmp.MoveNext
            Loop
        End If
        
        '��������
        strSql = " Select Distinct A.ID,A.����,A.����" & _
                    " From ���ű� A,��������˵�� B " & _
                    " Where B.����ID = A.ID " & _
                    " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
                    " And (B.�������� IN('�ٴ�','���','���'))" & _
                    " Order by A.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        cbo��������.Clear
        Do Until rsTmp.EOF
            cbo��������.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo��������.ItemData(cbo��������.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlngCurDeptId Then cbo��������.ListIndex = cbo��������.NewIndex
            rsTmp.MoveNext
        Loop
        If cbo��������.ListCount > 0 And Me.cbo��������.ListIndex = -1 Then cbo��������.ListIndex = 0
        
        '��Ӱ��
        strSql = "select ���� from ��Ӱ��"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        cbo��Ӱ��.Clear
        cbo��Ӱ��.AddItem "                 "
        Do Until rsTmp.EOF
            cbo��Ӱ��.AddItem rsTmp!����
            rsTmp.MoveNext
        Loop
    End If
    
    '��鼼ʦ
    strSql = "Select /*+RULE*/" & vbNewLine & _
                "Distinct b.id,b.����, Upper(b.����) As ����" & vbNewLine & _
                " From ������Ա a, ��Ա�� b " & vbNewLine & _
                " Where a.����id = [1] And a.��Աid = b.Id And " & vbNewLine & _
                "      (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null)" & vbNewLine & _
                " Order By ���� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId)
    
    '���ؼ�鼼ʦһ
    cbo��ʦһ.Clear
    Do Until rsTmp.EOF
        cbo��ʦһ.AddItem rsTmp!���� & "-" & rsTmp!����
        If rsTmp!ID = UserInfo.ID Then cbo��ʦһ.ListIndex = cbo��ʦһ.NewIndex
        rsTmp.MoveNext
    Loop
    If cbo��ʦһ.ListCount > 0 And cbo��ʦһ.ListIndex = -1 And mintEditMode = 2 Then cbo��ʦһ.ListIndex = 0
    
    '���ؼ�鼼ʦ��
    cbo��ʦ��.Clear
    
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        Do Until rsTmp.EOF
            cbo��ʦ��.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ID = UserInfo.ID Then cbo��ʦ��.ListIndex = cbo��ʦ��.NewIndex
            rsTmp.MoveNext
        Loop
    End If
    
    '�Ա�ע��� �� ���ݿ��е�ֵ��������鼼ʦһ ����ѡ�λ
    If mintEditMode > 1 Or mblnRegToCheck Then
        If mstrExamineDoctorFst <> "" Or mstrExamineDoctorSed <> "" Then
        
            strSql = "select ��鼼ʦ,��鼼ʦ�� from Ӱ�����¼ where ҽ��ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�õ���鼼ʦ", mlngAdviceID)
            
            For i = 0 To cbo��ʦһ.ListCount - 1
                If NeedName(cbo��ʦһ.list(i)) = IIf(mblnExamineDoctorVerify, mstrExamineDoctorName, mstrExamineDoctorFst) Then
                    cbo��ʦһ.ListIndex = i
                    Exit For
                Else
                    cbo��ʦһ.ListIndex = -1
                End If
            Next i
            
            If rsTmp.RecordCount > 0 And mstrExamineDoctorFst = "" Then
                For i = 0 To cbo��ʦһ.ListCount - 1
                    If NeedName(cbo��ʦһ.list(i)) = Nvl(rsTmp!��鼼ʦ) Then
                        cbo��ʦһ.ListIndex = i
                        Exit For
                    Else
                        cbo��ʦһ.ListIndex = -1
                    End If
                Next i
            End If
            
            
            For i = 0 To cbo��ʦ��.ListCount - 1
                If NeedName(cbo��ʦ��.list(i)) = mstrExamineDoctorSed Then
                    cbo��ʦ��.ListIndex = i
                    Exit For
                Else
                    cbo��ʦ��.ListIndex = -1
                End If
            Next i
            
            If rsTmp.RecordCount > 0 And mstrExamineDoctorSed = "" Then
                For i = 0 To cbo��ʦ��.ListCount - 1
                    If NeedName(cbo��ʦ��.list(i)) = Nvl(rsTmp!��鼼ʦ��) Then
                        cbo��ʦ��.ListIndex = i
                        Exit For
                    Else
                        cbo��ʦ��.ListIndex = -1
                    End If
                Next i
            End If
            
        Else
            
            cbo��ʦһ.ListIndex = -1
            cbo��ʦ��.ListIndex = -1
            
        End If
    End If
    

    '��ʼ��ִ�м�
    If mlngCurDeptId = 0 Then
        strSql = "Select ִ�м�,����豸 From ҽ��ִ�з���"
    Else
        strSql = "Select ִ�м�,����豸 From ҽ��ִ�з��� Where ����id = [1]"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngCurDeptId)
    cboRoom.Clear
    Do While Not rsTmp.EOF
        cboRoom.AddItem rsTmp!ִ�м� & "-" & Nvl(rsTmp!����豸)
        rsTmp.MoveNext
    Loop
    
    If mblnUsePacsQueue Then cboRoom.AddItem "����ʱָ��"
    
    If cboRoom.ListCount <= 0 Then
        cboRoom.Enabled = False
    Else
        Call InitDevice
        strSql = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & mlngCurDeptId & "\" & Me.Name, "��ǰִ�м�", "") '��ȡ�ϴεǼ�ʱ��ִ�м�
        If mintEditMode >= 2 Or mblnRegToCheck Then Call SeekIndexWithNo(cboRoom, strSql, True)
    End If

    InitInput '��꾭��λ��
    
    '�Ǽǵ��������Ҫ���ƿؼ��Ŀ�����
    If mintEditMode = 0 Then Call RefreshObjEnabled
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub InitDevice(Optional ByVal CheckType As String)
'------------------------------------------------
'���ܣ���ʼ�������Ӱ���豸
'������ CheckType -Ӱ�����
'���أ���
'------------------------------------------------
Dim rsTmp As ADODB.Recordset
    
    cboDevice.Clear
    gstrSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where " & IIf(CheckType <> "", "Ӱ�����=[1] AND ", "") & "����=4 AND  ״̬=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CheckType)
    Do Until rsTmp.EOF
        cboDevice.AddItem rsTmp!�豸�� & "-" & Nvl(rsTmp!�豸��)
        rsTmp.MoveNext
    Loop
End Sub
Private Sub LoadOldData(ByVal strOld As String, ByRef txt���� As TextBox, ByRef cbo���䵥λ As ComboBox)
'����:�����ݿ��б�������䰴�淶�ĸ�ʽ���ص�����,���淶��ԭ����ʾ
    Dim strTmp As String, lngIdx As Long
    
    If Trim(strOld) = "" Then Exit Sub
    
    lngIdx = -1
    strTmp = strOld
    If InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngIdx = 0
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngIdx = 1
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngIdx = 2
        End If
    ElseIf IsNumeric(strOld) Then
        lngIdx = 0
    End If
    
    If strTmp = "" Then strTmp = 0
    txt����.Text = strTmp
    If cbo���䵥λ.ListCount > 0 Then Call zlControl.CboSetIndex(cbo���䵥λ.hWnd, lngIdx)
    If lngIdx = -1 Then
        cbo���䵥λ.Visible = False
    Else
        If cbo���䵥λ.Visible = False Then cbo���䵥λ.Visible = True
    End If
End Sub

Private Function CopyCheck(ByVal lngAdviceID As Long, ByVal lngSendNO As Long) As Boolean
'����:���ڸ��ƵǼǣ�ͬһ������ͬ��Ŀ����ͬ��λ
'���أ� True--���Ƴɹ���False--������Ϣ������

    Dim rsTemp As New ADODB.Recordset
    Dim curDate As Date

    On Error GoTo errHand
    CopyCheck = False
    
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "SELECT nvl(B.����,E.����) ����,nvl(B.�Ա�,E.�Ա�) �Ա�,nvl(B.����,E.����) ����,B.��������,B.�ѱ�,B.ҽ�Ƹ��ʽ,B.���֤��,B.����,B.ְҵ,Nvl(E.Ӣ����,'') Ӣ����,E.���,E.����" & _
                    ",B.����״��,Nvl(B.��ͥ�绰,B.��ϵ�˵绰) �绰,Nvl(B.��ͥ��ַ�ʱ�,B.��λ�ʱ�) �ʱ�,nvl(B.��ͥ��ַ,B.������λ) ��ַ,B.��ͬ��λID,B.�����,B.���￨��,B.����֤��" & _
                    ",NVL(D.����,'') AS ���˿���,A.���˿���ID,Decode(A.������Դ,2,B.סԺ��,B.�����) As ���˺�,Decode(B.סԺ��,NULL,NULL,B.��ǰ����) As ����" & _
                    ",F.����ʱ�� ����ʱ��,NVL(C.����,0) ���ұ���,NVL(C.����,'δ֪') AS ��������,A.����ҽ��,A.������־,F.�״�ʱ��,F.ִ�м�,E.����豸,A.ҽ������,E.����,E.��鼼ʦ,E.��鼼ʦ�� " & _
                    ",DECODE(A.������Դ,2,2,1,1,4,4,3) AS ������Դ,Nvl(E.Ӱ�����,G.Ӱ�����) As Ӱ�����,B.����id,A.��ҳid,A.������ĿID,E.��������" & _
                " FROM ����ҽ������ F,����ҽ����¼ A, ������Ϣ B,���ű� C,���ű� D,Ӱ�����¼ E,Ӱ������Ŀ G " & _
                " Where F.ҽ��ID=[1] And F.���ͺ�=[2] AND F.ҽ��ID=A.ID" & _
                        " AND F.ҽ��ID=E.ҽ��ID(+) And F.���ͺ�=E.���ͺ�(+)  And A.����ID=B.����ID" & _
                        " And A.��������ID=C.ID And A.���˿���ID=D.ID And A.������ĿID=G.������ĿID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", lngAdviceID, lngSendNO)

    If rsTemp.EOF Then
        '��鲡����Ϣ��������ԭ�������û�С�����ҽ�����ͼ�¼������ʾ����ҽ���ѱ����˻�����
        gstrSQL = "Select ҽ��ID From ����ҽ������ Where ҽ��ID =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ҽ��״̬", lngAdviceID)
        If rsTemp.EOF Then
            Call MsgBoxD(Me, "���μ��ҽ��û�з��ͼ�¼�������Ǹ�ҽ���Ѿ������˻��������ϣ���ˢ�º���ҽ��״̬��", vbInformation, gstrSysName)
        Else
            Call MsgBoxD(Me, "������Ϣ���������������Ա��ϵ��", vbInformation, gstrSysName)
        End If
        
        mlngResultState = 0
        CmdOK.Enabled = False
        Exit Function
    End If
    
    txt���� = Nvl(rsTemp!����):  TxtӢ���� = Decode(Nvl(rsTemp!Ӣ����), "", zlCommFun.mGetFullPY(txt����, mintCapital, mblnUseSplitter), rsTemp!Ӣ����)
    Call SeekIndex(cbo�Ա�, Nvl(rsTemp!�Ա�), True)
    If Nvl(rsTemp!����) <> "" Then
        LoadOldData rsTemp!����, txt����, cboAge
    Else
        ReCalcOld Format(Nvl(rsTemp!��������, curDate), "yyyy-mm-dd"), cboAge
    End If
    If Trim(txt����) = "" Then txt���� = 0
    Txt��� = Nvl(rsTemp!���): Txt���� = Nvl(rsTemp!����): dtp��������.value = Format(Nvl(rsTemp!��������, ReCalcBirth(txt����.Text, cboAge.Text)), "yyyy-mm-dd")
    Call SeekIndex(cbo�ѱ�, Nvl(rsTemp!�ѱ�), True)
    Call SeekIndex(cbo���ʽ, Nvl(rsTemp!ҽ�Ƹ��ʽ), True)
    Txt���֤�� = Nvl(rsTemp!���֤��)
    Call SeekIndex(cbo����, Nvl(rsTemp!����), True)
    Call SeekIndex(cboְҵ, Nvl(rsTemp!ְҵ), True)
    Call SeekIndex(cbo����, Nvl(rsTemp!����״��), True)
    Txt�绰 = Nvl(rsTemp!�绰): Txt�ʱ� = Nvl(rsTemp!�ʱ�)
    Txt��ϵ��ַ = Nvl(rsTemp!��ַ)
    Label22.Tag = Nvl(rsTemp!��ͬ��λID, 0)
    
    txtPatientDept.Text = Nvl(rsTemp!���˿���)
    txtPatientDept.Tag = Nvl(rsTemp!���˿���ID, 0)
    txtID = Nvl(rsTemp!���˺�): txtBed = Nvl(rsTemp!����)
    dtp(0).value = Format(rsTemp!����ʱ��, "yyyy-mm-dd HH:MM")
    Call SeekIndex(cbo��������, Nvl(rsTemp!���ұ���), True, , True)
    Call SeekIndex(cboҽ��1, Nvl(rsTemp!����ҽ��), True)
    Call SeekIndex(cboҽ��2, Nvl(rsTemp!����ҽ��), True)
    '���Ҳ�������ҽ�����ҿ���ҽ����Ϊ�գ���ֱ����д����ҽ���ֶ�
    
    If Nvl(rsTemp!����ҽ��) <> "" And cboҽ��1.ListIndex = -1 Then
        Me.cboҽ��1.Visible = False
        Me.cboҽ��2.Visible = True
        cboҽ��2.Text = Nvl(rsTemp!����ҽ��)
    End If
    
    chk����.value = Nvl(rsTemp!������־, 0)
    dtp(1).value = Format(Nvl(rsTemp!�״�ʱ��, curDate), "yyyy-mm-dd HH:MM")
    
    Call SeekIndexWithNo(cboRoom, Nvl(rsTemp!ִ�м�), True) 'ƥ��ִ�м�
    
    txt��������.Text = Nvl(rsTemp!��������)
    'ҽ�����ݡ���������,����/����:��λ1(����1),��λ1(����2),��λ2(����1)---
    txtҽ������ = Split(Split(rsTemp!ҽ������, ":")(0), ",")(0)
    Call SeekIndex(cbo��ʦһ, Nvl(rsTemp!��鼼ʦ), True, True)
    Call SeekIndex(cbo��ʦ��, Nvl(rsTemp!��鼼ʦ��), True, True)
    
    mstrOutNo = Nvl(rsTemp!�����, 0)
    mstrCardNo = Nvl(rsTemp!���￨��)
    mstrCardPass = Nvl(rsTemp!����֤��)
    mintSourceType = rsTemp!������Դ
    
    If mblnAllPatientIsOutside Then mintSourceType = 3
    
    mlngPatiId = Nvl(rsTemp!����ID, 0)
    mlngPageID = Nvl(rsTemp!��ҳID, 0)
    mstrItemType = Nvl(rsTemp!Ӱ�����)
    mlngClinicID = Nvl(rsTemp!������ĿID)
    
    If mstrItemType = "" Then
        MsgBoxD Me, "���μ����Ŀδ����Ӱ������Ŀ,����", vbInformation, gstrSysName
        mlngResultState = 0
        CmdOK.Enabled = False
        Exit Function
    End If
    
    '��ʾ�ͼ쵥λ���ͼ�ҽ����Ϣ
    If mblnInputOutInfo Then
        gstrSQL = "select ��Ϣ��,��Ϣֵ from ������Ϣ�ӱ� where ����ID=[1] and ����id=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ������Ϣ", mlngPatiId, mlngAdviceID)
        Do Until rsTemp.EOF
            If Nvl(rsTemp!��Ϣ��) = "�ͼ쵥λ" Then txt�ͼ쵥λ.Text = Nvl(rsTemp!��Ϣֵ)
            If Nvl(rsTemp!��Ϣ��) = "�ͼ�ҽ��" Then txt�ͼ�ҽ��.Text = Nvl(rsTemp!��Ϣֵ)
            rsTemp.MoveNext
        Loop
    End If
    
    gstrSQL = "select ��Ӱ��,����,Ũ�� from ������Ӱ�� where ҽ��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", mlngAdviceID)
    If Not rsTemp.EOF Then
        Call SeekIndex(cbo��Ӱ��, Nvl(rsTemp!��Ӱ��), True)
        Txt��Ӱ����.Text = Nvl(rsTemp!����)
        Txt��ӰŨ��.Text = Nvl(rsTemp!Ũ��)
    End If

    txtҽ������.TabIndex = 0
    
    CopyCheck = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function RefreshPatiInfor(bln���� As Boolean) As Boolean
'����:���ڱ������޸�ʱˢ�²���
'bln����=True���Ǳ������򲿷���Ϣ����ֱ��ʹ��Ĭ����Ϣ
'bln����=False,���޸ģ�����ϢӦ��ȫ��ʹ�����ݿ��е���Ϣ

Dim rsTemp As New ADODB.Recordset
Dim strSql As String
Dim rsBaby As New ADODB.Recordset
Dim lngPatientID As Long
Dim lngPageID As Long
Dim intChargeState As Integer
Dim intChargeType As Integer    '����ҽ������.��¼����---1-�շѼ�¼��2-���ʼ�¼��
Dim curDate As Date

    On Error GoTo errHand
    
    RefreshPatiInfor = False
    
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "SELECT nvl(B.����,E.����) ����,nvl(B.�Ա�,E.�Ա�) �Ա�,nvl(B.����,E.����) ����,B.��������,B.�ѱ�,B.ҽ�Ƹ��ʽ,B.���֤��,B.����,B.ְҵ,Nvl(E.Ӣ����,'') Ӣ����,E.���,E.����" & _
                    ",B.����״��,Nvl(B.��ͥ�绰,B.��ϵ�˵绰) �绰,Nvl(B.��ͥ��ַ�ʱ�,B.��λ�ʱ�) �ʱ�,nvl(B.��ͥ��ַ,B.������λ) ��ַ,B.��ͬ��λID,B.�����,B.���￨��,B.����֤��" & _
                    ",NVL(D.����,'') AS ���˿���,A.���˿���ID,Decode(A.������Դ,2,B.סԺ��,B.�����) As ���˺�,Decode(B.סԺ��,NULL,NULL,B.��ǰ����) As ����,B.��ǰ����ID" & _
                    ",F.����ʱ�� ����ʱ��,NVL(C.����,0) ���ұ���,NVL(C.����,'δ֪') AS ��������,A.����ҽ��,A.������־,F.�״�ʱ��,F.ִ�м�,E.����豸,A.ҽ������,E.����,E.��鼼ʦ" & _
                    ",DECODE(A.������Դ,2,2,1,1,4,4,3) AS ������Դ,Nvl(E.Ӱ�����,G.Ӱ�����) As Ӱ�����,B.����id,A.��ҳid,A.������ĿID,E.��������,Nvl(A.Ӥ��, 0) As Ӥ��" & _
                    ",F.��¼���� " & _
                " FROM ����ҽ������ F,����ҽ����¼ A, ������Ϣ B,���ű� C,���ű� D,Ӱ�����¼ E,Ӱ������Ŀ G " & _
                " Where F.ҽ��ID=[1] And F.���ͺ�=[2] AND F.ҽ��ID=A.ID" & _
                        " AND F.ҽ��ID=E.ҽ��ID(+) And F.���ͺ�=E.���ͺ�(+)  And A.����ID=B.����ID" & _
                        " And A.��������ID=C.ID And A.���˿���ID=D.ID And A.������ĿID=G.������ĿID(+)"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", mlngAdviceID, mlngSendNo)

    If rsTemp.EOF Then
        '��鲡����Ϣ��������ԭ�������û�С�����ҽ�����ͼ�¼������ʾ����ҽ���ѱ����˻�����
        gstrSQL = "Select ҽ��ID From ����ҽ������ Where ҽ��ID =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ҽ��״̬", mlngAdviceID)
        If rsTemp.EOF Then
            Call MsgBoxD(Me, "���μ��ҽ��û�з��ͼ�¼�������Ǹ�ҽ���Ѿ������˻��������ϣ���ˢ�º���ҽ��״̬��", vbInformation, gstrSysName)
        Else
            Call MsgBoxD(Me, "������Ϣ���������������Ա��ϵ��", vbInformation, gstrSysName)
        End If
    
        mlngResultState = 0
        CmdOK.Enabled = False
        Exit Function
    End If
    
    '����Ӥ����Ϣ
    mlngBaby = rsTemp!Ӥ��
    If mlngBaby = 0 Then
Normal:
        txt���� = Nvl(rsTemp!����)
        Call SeekIndex(cbo�Ա�, Nvl(rsTemp!�Ա�), True)
        If Nvl(rsTemp!����) <> "" Then
            LoadOldData rsTemp!����, txt����, cboAge
        Else
            ReCalcOld Format(Nvl(rsTemp!��������, curDate), "yyyy-mm-dd"), cboAge
        End If
        dtp��������.value = Format(Nvl(rsTemp!��������, ReCalcBirth(txt����.Text, cboAge.Text)), "yyyy-mm-dd")
    Else
        lngPatientID = rsTemp!����ID
        lngPageID = Nvl(rsTemp!��ҳID, 0)
        strSql = "Select Decode(a.Ӥ������,Null,b.����||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������,Ӥ���Ա�,����ʱ�� From ������������¼ a,������Ϣ b Where a.����id=[1] And a.��ҳid=[2] And a.����id=b.����id And a.���=[3]"
        Set rsBaby = zlDatabase.OpenSQLRecord(strSql, "��ȡӤ����Ϣ", lngPatientID, lngPageID, mlngBaby)
        If rsBaby.EOF Then
            GoTo Normal
        Else
            txt���� = Nvl(rsBaby!Ӥ������)
            Call SeekIndex(cbo�Ա�, Nvl(rsBaby!Ӥ���Ա�), True)
            ReCalcOld Format(Nvl(rsBaby!����ʱ��, curDate), "yyyy-mm-dd"), cboAge
            dtp��������.value = Format(Nvl(rsBaby!����ʱ��, ReCalcBirth(txt����.Text, cboAge.Text)), "yyyy-mm-dd")
        End If
    End If
    lblCash.Tag = Nvl(rsTemp!��ǰ����ID)
    TxtӢ���� = Decode(Nvl(rsTemp!Ӣ����), "", zlCommFun.mGetFullPY(txt����, mintCapital, mblnUseSplitter), rsTemp!Ӣ����)
    If Trim(txt����) = "" Then txt���� = 0
    Txt��� = Nvl(rsTemp!���): Txt���� = Nvl(rsTemp!����)
    Call SeekIndex(cbo�ѱ�, Nvl(rsTemp!�ѱ�), True)
    Call SeekIndex(cbo���ʽ, Nvl(rsTemp!ҽ�Ƹ��ʽ), True)
    Txt���֤�� = Nvl(rsTemp!���֤��)
    Call SeekIndex(cbo����, Nvl(rsTemp!����), True)
    Call SeekIndex(cboְҵ, Nvl(rsTemp!ְҵ), True)
    Call SeekIndex(cbo����, Nvl(rsTemp!����״��), True)
    Txt�绰 = Nvl(rsTemp!�绰): Txt�ʱ� = Nvl(rsTemp!�ʱ�)
    Txt��ϵ��ַ = Nvl(rsTemp!��ַ)
    Label22.Tag = Nvl(rsTemp!��ͬ��λID, 0)
    
    txtPatientDept.Text = Nvl(rsTemp!���˿���)
    txtPatientDept.Tag = Nvl(rsTemp!���˿���ID, 0)
    txtID = Nvl(rsTemp!���˺�): txtBed = Nvl(rsTemp!����)
    dtp(0).value = Format(rsTemp!����ʱ��, "yyyy-mm-dd HH:MM")
    Call SeekIndex(cbo��������, Nvl(rsTemp!���ұ���), True, , True)
    Call SeekIndex(cboҽ��1, Nvl(rsTemp!����ҽ��), True)
    Call SeekIndex(cboҽ��2, Nvl(rsTemp!����ҽ��), True)
    
    '���Ҳ�������ҽ�����ҿ���ҽ����Ϊ�գ���ֱ����д����ҽ���ֶ�
    If Nvl(rsTemp!����ҽ��) <> "" And cboҽ��1.ListIndex = -1 Then
        Me.cboҽ��1.Visible = False
        Me.cboҽ��2.Visible = True
        cboҽ��2.Text = Nvl(rsTemp!����ҽ��)
    End If

    chk����.value = Nvl(rsTemp!������־, 0)
    dtp(1).value = Format(Nvl(rsTemp!�״�ʱ��, curDate), "yyyy-mm-dd HH:MM")
    
    If Not bln���� Then '�޸���Ϣ
        Call SeekIndexWithNo(cboRoom, Nvl(rsTemp!ִ�м�), True)
    End If
    
    txt��������.Text = Nvl(rsTemp!��������)
    'ҽ�����ݡ���������,����/����:��λ1(����1),��λ1(����2),��λ2(����1)---
    txtҽ������ = Split(Split(rsTemp!ҽ������, ":")(0), ",")(0)
    txtҽ������.Tag = txtҽ������.Text
    If InStr(Nvl(rsTemp!ҽ������, ""), ":") > 0 Then
        Txt��λ���� = Replace(Split(rsTemp!ҽ������, ":")(1), "),", ")" & vbCrLf)
    Else
        Txt��λ���� = Nvl(rsTemp!ҽ������, "")
    End If
    txt����.Text = CStr(Nvl(rsTemp!����)): txt����.Tag = txt����.Text
    
    '������޸Ĳ��� ��ˢ�±�����ֵ
    If mintEditMode = 3 Then mstrNextCheckNo = CStr(Nvl(rsTemp!����))
    
    Call SeekIndex(cbo��ʦһ, Nvl(rsTemp!��鼼ʦ), True, True)
    
    mstrOutNo = Nvl(rsTemp!�����, 0)
    mstrCardNo = Nvl(rsTemp!���￨��)
    mstrCardPass = Nvl(rsTemp!����֤��)
    mintSourceType = rsTemp!������Դ
    mlngPatiId = Nvl(rsTemp!����ID, 0)
    mlngPageID = Nvl(rsTemp!��ҳID, 0)
    mstrItemType = Nvl(rsTemp!Ӱ�����)
    mlngClinicID = Nvl(rsTemp!������ĿID)
    
    intChargeType = Nvl(rsTemp!��¼����, 1)
    
    If mstrItemType = "" Then
        MsgBoxD Me, "���μ����Ŀδ����Ӱ������Ŀ,����", vbInformation, gstrSysName
        mlngResultState = 0
        CmdOK.Enabled = False
        Exit Function
    End If
    
    '��ʾ�ͼ쵥λ���ͼ�ҽ����Ϣ
    If mblnInputOutInfo Then
        gstrSQL = "select ��Ϣ��,��Ϣֵ from ������Ϣ�ӱ� where ����ID=[1] and ����id=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ������Ϣ", mlngPatiId, mlngAdviceID)
        Do Until rsTemp.EOF
            If Nvl(rsTemp!��Ϣ��) = "�ͼ쵥λ" Then txt�ͼ쵥λ.Text = Nvl(rsTemp!��Ϣֵ)
            If Nvl(rsTemp!��Ϣ��) = "�ͼ�ҽ��" Then txt�ͼ�ҽ��.Text = Nvl(rsTemp!��Ϣֵ)
            rsTemp.MoveNext
        Loop
    End If
    
    gstrSQL = "select ��Ӱ��,����,Ũ�� from ������Ӱ�� where ҽ��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", mlngAdviceID)
    If Not rsTemp.EOF Then
        Call SeekIndex(cbo��Ӱ��, Nvl(rsTemp!��Ӱ��), True)
        Txt��Ӱ����.Text = Nvl(rsTemp!����)
        Txt��ӰŨ��.Text = Nvl(rsTemp!Ũ��)
    End If
    
    gstrSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˸���", mlngAdviceID)
    Txt��λ���� = Txt��λ���� & vbCrLf
    Do Until rsTemp.EOF
        Txt��λ���� = Txt��λ���� & rsTemp!��Ŀ & ":" & Nvl(rsTemp!����) & vbCrLf
        rsTemp.MoveNext
    Loop
    
    If mintEditMode = 2 Then
        txt����.Text = Next����: txt����.Tag = txt����.Text
    End If
    
    intChargeState = CheckChargeState(mlngAdviceID, mintSourceType)
    If intChargeState = 1 Then
        lblCash.Caption = "����"
    ElseIf intChargeState = 2 Then
        lblCash.Caption = "��"
    Else
        lblCash.Caption = ""
    End If
    Call RefreshObjEnabled
    
    If bln���� And InStr(mstrPrivs, "δ�ɷѱ���") = 0 And mintSourceType <> 3 Then '24361 ��Ȩ�޲��жϣ����еǼǲ����ƣ�����Ҳ�����ж�
        If lblCash.Caption = "����" Or lblCash.Caption = "��" _
            Or (gblnִ�к���� And intChargeType = 2) _
            Or gblnִ��ǰ�Ƚ��� Then
            ''��Ҫ����ϵͳ�����жϣ� gblnִ�к����=81�Ų�����"ִ�к��Զ���˻��۵�",��ѡ���������û��δ���ѱ���Ȩ��ʱ��ҲӦ�ÿ��ԶԼ��˼�¼���б���
            ''gblnִ��ǰ�Ƚ��� = 163--����һ��ͨ����Ŀִ��ǰ�������շѻ��ȼ������,��ѡ���������û��δ���ѱ���Ȩ��ʱ��ҲӦ�ÿ��Խ��б�����������ʱ���ˢ������
            CmdOK.Enabled = True
        Else
            CmdOK.Enabled = False
        End If

        If CmdOK.Enabled = False Then
            Me.Caption = Me.Caption & "(��ǰ����δ�շѣ����ܱ���)"
        End If
    End If
    
    RefreshPatiInfor = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub CmdCancle_Click()
    mlngResultState = IIf(mlngGoOnReg = 1, 4, 0)
    Unload Me
End Sub

Private Function ValidData() As Boolean
'------------------------------------------------
'���ܣ�����������ݵĺϷ���
'������ ��
'���أ�True--��������ϸ񣬿��Լ�����False --���������벻�ϸ���Ҫ�޸�����
'------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    
    ValidData = False
    
    gstrSQL = "select ID ,����ID,����ֵ from Ӱ�����̲��� where ����ID = [1] and ������ = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngCurDeptId, CStr("��¼����"))
    If Not rsTemp.EOF Then
        If Nvl(rsTemp!����ֵ) <> "" Then
            If InStr(rsTemp!����ֵ, "Ӣ����") > 0 And Trim(TxtӢ����) = "" And TxtӢ����.Enabled = True Then
                MsgBoxD Me, "��������Ӣ���������飡", vbInformation, gstrSysName: DoEvents
                TxtӢ����.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "�Ա�") > 0 And Trim(cbo�Ա�.Text) = "" And cbo�Ա�.Enabled = True Then
                MsgBoxD Me, "���������Ա����飡", vbInformation, gstrSysName: DoEvents
                cbo�Ա�.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "����") > 0 And Trim(txt����) = "" And txt����.Enabled = True Then
                MsgBoxD Me, "�����������䣬���飡", vbInformation, gstrSysName: DoEvents
                txt����.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "��������") > 0 And Trim(dtp��������.value) = "" And dtp��������.Enabled = True Then
                MsgBoxD Me, "��������������ڣ����飡", vbInformation, gstrSysName: DoEvents
                dtp��������.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "���") > 0 And Trim(Txt���) = "" And Txt���.Enabled = True Then
                MsgBoxD Me, "����������ߣ����飡", vbInformation, gstrSysName: DoEvents
                Txt���.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "����") > 0 And Trim(Txt����) = "" And Txt����.Enabled = True Then
                MsgBoxD Me, "�����������أ����飡", vbInformation, gstrSysName: DoEvents
                Txt����.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "�ѱ�") > 0 And Trim(cbo�ѱ�.Text) = "" And cbo�ѱ�.Enabled = True Then
                MsgBoxD Me, "��������ѱ����飡", vbInformation, gstrSysName: DoEvents
                cbo�ѱ�.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "���ʽ") > 0 And Trim(cbo���ʽ.Text) = "" And cbo���ʽ.Enabled = True Then
                MsgBoxD Me, "�������븶�ʽ�����飡", vbInformation, gstrSysName: DoEvents
                cbo���ʽ.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "���֤��") > 0 And Trim(Txt���֤��) = "" And Txt���֤��.Enabled = True Then
                MsgBoxD Me, "�����������֤�ţ����飡", vbInformation, gstrSysName: DoEvents
                Txt���֤��.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "����") > 0 And Trim(cbo����.Text) = "" And cbo����.Enabled = True Then
                MsgBoxD Me, "�����������壬���飡", vbInformation, gstrSysName: DoEvents
                cbo����.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "ְҵ") > 0 And Trim(cboְҵ.Text) = "" And cboְҵ.Enabled = True Then
                MsgBoxD Me, "��������ְҵ�����飡", vbInformation, gstrSysName: DoEvents
                cboְҵ.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "����") > 0 And Trim(cbo����.Text) = "" And cbo����.Enabled = True Then
                MsgBoxD Me, "����������������飡", vbInformation, gstrSysName: DoEvents
                cbo����.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "�绰") > 0 And Trim(Txt�绰) = "" And Txt�绰.Enabled = True Then
                MsgBoxD Me, "��������绰�����飡", vbInformation, gstrSysName: DoEvents
                Txt�绰.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "�ʱ�") > 0 And Trim(Txt�ʱ�) = "" And Txt�ʱ�.Enabled = True Then
                MsgBoxD Me, "���������ʱ࣬���飡", vbInformation, gstrSysName: DoEvents
                Txt�ʱ�.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "��ַ") > 0 And Trim(Txt��ϵ��ַ) = "" And Txt��ϵ��ַ.Enabled = True Then
                MsgBoxD Me, "����������ϵ��ַ�����飡", vbInformation, gstrSysName: DoEvents
                Txt��ϵ��ַ.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "ִ�м�") > 0 And Trim(cboRoom.Text) = "" And cboRoom.Enabled = True Then
                MsgBoxD Me, "��������ִ�м䣬���飡", vbInformation, gstrSysName: DoEvents
                cboRoom.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "��Ӱ��") > 0 And Trim(cbo��Ӱ��.Text) = "" And cbo��Ӱ��.Enabled = True Then
                MsgBoxD Me, "����������Ӱ�������飡", vbInformation, gstrSysName: DoEvents
                cbo��Ӱ��.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "��鼼ʦ") > 0 And Trim(cbo��ʦһ.Text) = "" And cbo��ʦһ.Enabled = True Then
                MsgBoxD Me, "���������鼼ʦ�����飡", vbInformation, gstrSysName: DoEvents
                cbo��ʦһ.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "��鼼ʦ��") > 0 And Trim(cbo��ʦһ.Text) = "" And cbo��ʦ��.Enabled = True Then
                MsgBoxD Me, "���������鼼ʦ�������飡", vbInformation, gstrSysName: DoEvents
                cbo��ʦ��.SetFocus: Exit Function
            ElseIf InStr(rsTemp!����ֵ, "��������") > 0 And Trim(txt��������.Text) = "" And txt��������.Enabled = True Then
                MsgBoxD Me, "�������븽�����������飡", vbInformation, gstrSysName: DoEvents
                txt��������.SetFocus: Exit Function
            End If
        End If
    End If

    On Error Resume Next
    If Len(Trim(Me.txtҽ������.Tag)) = 0 Then
        MsgBoxD Me, "��������������Ŀ��", vbInformation, gstrSysName: DoEvents
        Me.txtҽ������.SetFocus: Exit Function
    End If
    If Me.cbo��������.ListIndex = -1 Then
        MsgBoxD Me, "��ָ��������ң�", vbInformation, gstrSysName: DoEvents
        Me.cbo��������.SetFocus: Exit Function
    End If
    
    If cboҽ��1.Visible Then
        If Len(Trim(Me.cboҽ��1.Text)) = 0 Then
            MsgBoxD Me, "��ָ������ҽ����", vbInformation, gstrSysName: DoEvents
            Me.cboҽ��1.SetFocus: Exit Function
        End If
    Else
        If Len(Trim(Me.cboҽ��2.Text)) = 0 Then
            MsgBoxD Me, "��ָ������ҽ����", vbInformation, gstrSysName: DoEvents
            Me.cboҽ��2.SetFocus: Exit Function
        End If
    End If
   
    
    If dtp(0).value > dtp(1).value Then
        MsgBoxD Me, "����ʱ�䲻�ܴ��ڼ��ʱ�䣡", vbInformation, gstrSysName: DoEvents
        Me.dtp(0).SetFocus: Exit Function
    End If
    
    If Len(Trim(Me.txt����)) = 0 And txt����.Enabled Then
        MsgBoxD Me, "�����벡��������", vbInformation, gstrSysName: DoEvents
        Me.txt����.SetFocus
        Exit Function
    End If
    
    If Trim(TxtӢ����) = "" And TxtӢ����.TabStop And TxtӢ����.Enabled Then
        MsgBoxD Me, "Ӣ��������Ϊ�գ�", vbInformation, gstrSysName: DoEvents
        TxtӢ����.SetFocus
        Exit Function
    End If

    If mintEditMode >= 2 Or mblnRegToCheck Then '����,�򱨵����޸ġ��򡡵ǼǺ�ֱ�Ӽ�� (�Ǽ�ʱ��ǼǺ��޸Ĳ��ж�)
        If Len(Trim(Me.txt����)) = 0 And txt����.Enabled Then
            MsgBoxD Me, "���Ų���Ϊ�գ�", vbInformation, gstrSysName: DoEvents
            txt����.SetFocus
            Exit Function
        End If
        '�жϼ��ŵĵ����������������10������ʾ
        If mlngAutoIncNum <> 0 Then
            If Val(txt����.Text) > Val(mstrNextCheckNo) + 10 Then
                If MsgBoxD(Me, "���Ź��󣬱ȵ�ǰ�����������" & (Val(txt����.Text) - Val(mstrNextCheckNo)) & "����ȷ���Ƿ����" & IIf(mintEditMode = 3, "�޸�", "����") & "��", vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                    txt����.Text = Val(mstrNextCheckNo)
                    txt����.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    
    '������ǼǺ�ֱ�Ӽ���ʱ�򣬱���Ҫ����ִ�м�
    If mintEditMode = 2 Or mintEditMode = 3 Or (mblnRegToCheck And mintEditMode = 0) Then
        If cboRoom.Text = "" And Not mblnUsePacsQueue Then
            MsgBoxD Me, "ִ�м䲻��Ϊ�գ�", vbInformation, gstrSysName: DoEvents
            cboRoom.SetFocus
            Exit Function
        End If
    End If
    
    ValidData = True
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            zlCommFun.PressKey vbKeyTab
        Case vbKeyF2
            If mintEditMode <> 1 Then cmdOK_Click   '�ǼǺ��޸Ķ���F2
        Case vbKeyF4
            If mintEditMode = 1 Then cmdOK_Click   '������F4
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjIdCard Is Nothing Then '���֤ʶ�����
        Call mobjIdCard.SetEnabled(False)
        Set mobjIdCard = Nothing
    End If
    
    Set mobjIcCard = Nothing
    Set mobjSquareCard = Nothing
    
    Set mobjInsure = Nothing
    
    If mintEditMode = 2 Or mintEditMode = 3 Or mblnRegToCheck Then
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & mlngCurDeptId & "\" & Me.Name, "��ǰִ�м�", NeedNo(cboRoom)
    End If
    
    If mintEditMode > 1 Or mblnRegToCheck Then
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��鼼ʦһ", NeedName(cbo��ʦһ.Text)
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��鼼ʦ��", NeedName(cbo��ʦ��.Text)
    End If
    
    
        '�����жϵǼ�ʱɨ��� ���ȡ����ť ɨ�贰���ͷ�
    If Not frmPetitionCap Is Nothing Then
        frmPetitionCap.mblnIsLogin = False
        Call frmPetitionCap.Form_Unload(0)
        Set frmPetitionCap = Nothing
    End If
    
End Sub

Private Sub IDKind_ItemClick(Index As Integer)
    '���ݲ�ͬ������ò�ͬ���ų��ȺͿ���������ʾ
    Dim lngҽ�ƿ�����  As Long
    Dim bln�������� As Boolean
    Dim lng�����ID As Long
    
    
    '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
    '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
    lngҽ�ƿ����� = Val(IDKind.GetKindItem(IDKindItem_���ų���))
    '��7λ��,��ֻ��������,��Ȼȡ������
    bln�������� = Trim(IDKind.GetKindItem(7)) <> ""
    txt����.PasswordChar = IIf(bln��������, "*", "")
    lng�����ID = Val(IDKind.GetKindItem(IDKindItem_�����ID))
    If lng�����ID <> 0 Then
        txt����.MaxLength = lngҽ�ƿ�����
    Else
        txt����.MaxLength = 64
    End If
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txt����.Text <> "" Then txt����.Text = ""
    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
    zlControl.TxtSelAll txt����
End Sub

Private Sub Txt�绰_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��������_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��������_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    txt����.Locked = Not mblnChangeNo
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboAge.Visible = False And IsNumeric(txt����.Text) Then
            Call txt����_Validate(False)
            cboAge.SetFocus
        End If
        If Not IsNumeric(txt����.Text) Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt����_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not CheckOldData(txt����, cboAge) Then Exit Sub
    If IsNumeric(txt����.Text) Then dtp��������.value = ReCalcBirth(Val(txt����.Text), cboAge.Text)
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    If Not IsNumeric(txt����.Text) And Trim(txt����.Text) <> "" Then
        cboAge.ListIndex = -1: cboAge.Visible = False
    ElseIf cboAge.Visible = False Then
        cboAge.ListIndex = 0: cboAge.Visible = True
    End If
End Sub

Private Sub Txt���_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub Txt����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
On Error GoTo err
    Dim rsTmp As ADODB.Recordset
    Dim blnCard As Boolean
    Dim lngAge As Long
    Dim curDate As Date
    
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If IDKind.IDKind = IDKind.GetKindIndex(IDKind_����) Then
        blnCard = zlCommFun.InputIsCard(txt����, KeyAscii, oneSquardCard.blnȱʡ��������)
    ElseIf IDKind.IDKind = IDKind.GetKindIndex(IDKind_�����) Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    End If
                    
    '����ˢ�������߻س����������ʱ��Ҫ��ȡ��Ӧ�Ĳ�����Ϣ
        If ((blnCard And Len(txt����.Text) = IDKind.GetDefaultIDKindLength - 1 And KeyAscii <> 8) _
        Or (KeyAscii = vbKeyReturn And Trim(txt����.Text) <> "")) _
        And mintEditMode = 0 Then  '�Ǽ�ʱ�Ŵ���

        If KeyAscii <> 13 Then  '���ǻس�������ͨ��ˢ�������
            txt����.Text = txt����.Text & Chr(KeyAscii) 'ˢ��ʱ���һλ��δ��ʾ��TEXT��,TEXT���������һλ
            txt����.SelStart = Len(txt����.Text)
        End If
        KeyAscii = 0
        
        Set rsTmp = GetPatient(txt����.Text, blnCard) '����������ȡ������Ϣ
        If Not rsTmp Is Nothing Then
            If Not rsTmp.EOF Then
                If Nvl(rsTmp!����) <> "�²���" Then
                    curDate = zlDatabase.Currentdate
                
                    txt����.Tag = Trim(Nvl(rsTmp!����))
                    txt����.Text = Trim(Nvl(rsTmp!����))
                    Call SeekIndex(cbo�Ա�, Nvl(rsTmp!�Ա�), True)
                    
                    dtp��������.value = Format(Nvl(rsTmp!��������, curDate), "yyyy-mm-dd")
                    
                    If Nvl(rsTmp!��������, "") <> "" Then
                        lngAge = DateDiff("d", dtp��������.value, curDate)
                        
                        If lngAge > 0 Then
                            If lngAge > 365 Then
                                lngAge = Fix(lngAge / 365.25)
                                txt����.Text = lngAge & "��"
                            ElseIf lngAge > 30 Then
                                lngAge = Fix(lngAge / 30)
                                txt����.Text = lngAge & "��"
                            Else
                                txt����.Text = lngAge & "��"
                            End If
                        Else
                            txt����.Text = ""
                        End If
                    Else
                        txt����.Text = ""
                    End If
                    
                    If txt����.Text = "" Then txt����.Text = Nvl(rsTmp!����)
                    
                    If txt����.Text <> "" Then
                        '������䲻Ϊ�գ���ֱ�����txt�����cboage
                        LoadOldData txt����.Text, txt����, cboAge
                    Else
                        txt���� = 0
                        cboAge.Visible = True
                        cboAge.ListIndex = 0
                    End If
                    

                        
                    Call SeekIndex(cbo�ѱ�, Nvl(rsTmp!�ѱ�, "��ͨ"))
                    Call SeekIndex(cbo���ʽ, Nvl(rsTmp!ҽ�Ƹ��ʽ, "�Է�ҽ��"))
                    Txt���֤�� = Nvl(rsTmp!���֤��)
                    Call SeekIndex(cbo����, Nvl(rsTmp!����, "����"))
                    Call SeekIndex(cboְҵ, Nvl(rsTmp!ְҵ, "����"))
                    Call SeekIndex(cbo����, Nvl(rsTmp!����״��, "δ��"))
                    Txt�绰 = Nvl(rsTmp!�绰)
                    Txt�ʱ� = Nvl(rsTmp!�ʱ�)
                    Txt��ϵ��ַ = Nvl(rsTmp!��ַ)
                    Label22.Tag = Nvl(rsTmp!��ͬ��λID, 0)
                    txtID = Decode(Nvl(rsTmp!סԺ��), "", Nvl(rsTmp!�����), Nvl(rsTmp!סԺ��))
                    txtBed = Nvl(rsTmp!��ǰ����)
                    Call SeekIndex(cbo��������, getID_TO_����(Nvl(rsTmp!���˿���ID), "���ű�"), True, , True)
                    Call SeekIndex(cboҽ��1, Nvl(rsTmp!ҽ��))
                    Call SeekIndex(cboҽ��2, Nvl(rsTmp!ҽ��))
                    mlngPatiId = Nvl(rsTmp!����ID, 0)
                    mintSourceType = Nvl(rsTmp!��Դid, 1)
                    
                    '���ڷ�סԺ���ˣ������������ﻹ������
                    If mintSourceType <> 2 Then mintSourceType = getSourceType(rsTmp!����ID)
                    
                    mlngPageID = Nvl(rsTmp!��ҳID, 0)
                    mstrOutNo = Nvl(rsTmp!�����, 0)
                    mstrCardNo = Nvl(rsTmp!���￨��)
                    mstrCardPass = Nvl(rsTmp!����֤��)
                    
                    '��ʾ���˿���
                    txtPatientDept.Text = NeedName(cbo��������)
                    txtPatientDept.Tag = Nvl(rsTmp!���˿���ID)
                    If cbo�Ա�.Enabled = True Then cbo�Ա�.SetFocus
                    
                    Call RefreshObjEnabled
                    
                    '�����ˢ���ģ�txt�����ĳ�������ʾ
                    If txt����.PasswordChar = "*" Then
                        txt����.PasswordChar = ""
                    End If
                    
                    '��ȡ������Ϣ��ɺ� �Զ����㲡�˳�������
                    If IsNumeric(txt����.Text) Then dtp��������.value = ReCalcBirth(Val(txt����.Text), cboAge.Text)
        
                    Exit Sub
                Else
                    If cbo�Ա�.Enabled = True And mblnIsSamePatient Then cbo�Ա�.SetFocus
                    
                End If
            End If
        End If
        
        'û�鵽���µǼǲ�����
        Dim strTmp As String
        strTmp = Trim(txt����.Text)
        
        If IDKind.IDKind <> IDKind.GetKindIndex(IDKind_���֤��) Then '���֤��ȡ�����֤����������д��������Ϣ
            If txt���� <> strTmp Then txt���� = strTmp
            txt����.Tag = Trim(txt����.Text)
            TxtӢ����.Text = zlCommFun.mGetFullPY(txt����, mintCapital, mblnUseSplitter)
        End If
        mlngPatiId = 0
        mintSourceType = 3
        mlngPageID = 0
        'ˢ��������û����ȡ��������Ϣ����Ȼѡ��txt����
        If blnCard Then
            txt����.SelStart = 0
            txt����.SelLength = Len(txt����.Text)
            txt����.SetFocus
        End If
        Call RefreshObjEnabled
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Function getSourceType(ByVal lngPatiID As Long) As Integer
'����:��ȡ������Դ�͹Һŵ�
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If mInputType = 4 Then Exit Function 'Ϊ�Һŵ�ʱ��ȷ��Ϊ���ﲡ��
    'ȱʡΪ��Ժ����
    getSourceType = 3
    
    strSql = "select NO from ���˹Һż�¼ where ����ID=[1] and ִ��״̬=2 order by �Ǽ�ʱ�� desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ������Դ�͹Һŵ�", lngPatiID)
    
    If rsTemp.RecordCount > 0 Then
        getSourceType = 1
        mstrRegNo = Nvl(rsTemp!NO)
    End If
End Function

Private Sub txtҽ������_KeyPress(KeyAscii As Integer)
Dim rsTmp As ADODB.Recordset
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        With txtҽ������
            If .Text = "" Then Call cmdSel_Click
            If Trim(.Text) = .Tag Then Exit Sub
            
            Set rsTmp = SelectDiagItem() '��ȡ��Ŀ
            If rsTmp Is Nothing Then 'ȡ����������
                '�ָ�ԭֵ
                .Text = .Tag
                zlControl.TxtSelAll txtҽ������
                .SetFocus
                Exit Sub
            Else
                If AdviceInput(rsTmp) Then '����ѡ����Ŀ���ò�λ������
                    .Tag = .Text
                Else 'ȡ����λ������
                    .Text = .Tag
                    zlControl.TxtSelAll txtҽ������
                    .SetFocus
                    Exit Sub
                End If
            End If
        End With
    End If
End Sub

Private Sub txtҽ������_Validate(Cancel As Boolean)
    '�ָ���Ϊ�ĸı�,�س�ʱ��ֵ
    If txtҽ������.Text <> txtҽ������.Tag Then
        txtҽ������.Text = txtҽ������.Tag
    End If
End Sub

Private Sub TxtӢ����_LostFocus()
    zlControl.TxtSelAll TxtӢ����
End Sub

Private Sub Txt�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub cbo��������_Click()
    '�ж�ѡ����� �Ƿ�����Ժ����
    If InStr(cbo��������.Text, "��Ժ") > 0 Then
        mblnIsOutSideHosp = True
        
        cboҽ��1.Visible = False
        cboҽ��2.Visible = True
    Else
        mblnIsOutSideHosp = False
    
        cboҽ��1.Visible = True
        cboҽ��2.Visible = False
    End If

    If cbo��������.ListIndex > -1 Then InitDoctors cbo��������.ItemData(cbo��������.ListIndex)
End Sub
Private Sub txt����_LostFocus()
    TxtӢ����.Text = zlCommFun.mGetFullPY(txt����, mintCapital, mblnUseSplitter)
    Call zlCommFun.OpenIme
    If Not mobjIdCard Is Nothing Then mobjIdCard.SetEnabled (False)
End Sub

Private Sub txtҽ������_GotFocus()
    Call zlControl.TxtSelAll(txtҽ������)
End Sub
Private Sub mobjIdCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
        ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    If Me.ActiveControl Is Nothing Then Exit Sub
    If txt����.Text = "" And Me.ActiveControl Is txt���� Then
        IDKind.IDKind = IDKind.GetKindIndex(IDKind_���֤��)
        txt����.Text = strID
        Call txt����_KeyPress(vbKeyReturn)
        
        '�����²���
        If txt����.Text = "" Then
            Txt���֤��.Text = strID
            txt����.Text = strName
            txt����.Tag = strName
            TxtӢ����.Text = zlCommFun.mGetFullPY(txt����, mintCapital, mblnUseSplitter)
            Call SeekIndex(cbo�Ա�, strSex, True)
            Call SeekIndex(cbo����, strNation, True)
            dtp��������.value = Format(datBirthday, "yyyy-mm-dd")
            txt����.Text = Get����(Format(datBirthday, "yyyy-mm-dd"))
            cboAge.Visible = True: cboAge.ListIndex = 0
            Txt��ϵ��ַ.Text = strAddress
            IDKind.IDKind = IDKind.GetKindIndex(IDKind_����)
        End If
    End If
End Sub

Private Sub Txt��ϵ��ַ_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub Txt��ϵ��ַ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt����_Change()
    If Not mobjIdCard Is Nothing Then mobjIdCard.SetEnabled (txt����.Text = "" And Me.ActiveControl Is txt����)
    'ֻ�еǼǵ�ʱ����ȡ�˲��ˣ����޸��������Ż������²���
    If mintEditMode = 0 And mlngPatiId <> 0 And txt����.Text <> "" Then
        MsgBoxD Me, "�����޸������󣬾���Ϊ�²��˴����ˡ�", vbOKOnly, "��ʾ��Ϣ"
        mlngPatiId = 0
        Call txt����_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    Call zlCommFun.OpenIme(gstrIme <> "���Զ�����")
    
    If Not mobjIdCard Is Nothing And txt����.Text = "" Then mobjIdCard.SetEnabled (True)
End Sub
Private Sub IDKind_Click()
    Dim lng�����ID As Long
    Dim strExpand As String
    Dim strOutCardNO As String
    Dim strOutPatiInfoXML As String
    
    If IDKind.IDKind = IDKind.GetKindIndex(IDKind_IC����) Then
            If mobjIcCard Is Nothing Then
                Set mobjIcCard = CreateObject("zlICCard.clsICCard")
                Set mobjIcCard.gcnOracle = gcnOracle
            End If
            If Not mobjIcCard Is Nothing Then
                txt����.Text = mobjIcCard.Read_Card()
                If txt����.Text <> "" Then Call txt����_KeyPress(vbKeyReturn)
            End If
    End If
    
    lng�����ID = Val(IDKind.GetKindItem(IDKindItem_�����ID))
    
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If lng�����ID = 0 Then Exit Sub
    If mobjSquareCard.zlReadCard(Me, mlngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInfoXML) = False Then
        Exit Sub
    End If
    txt����.Text = strOutCardNO
    If txt����.Text <> "" Then
        Call txt����_KeyPress(vbKeyReturn)
    End If

End Sub

Private Sub IDKind_Validate(Cancel As Boolean)
    Select Case IDKind.IDKind
        Case IDKind.GetKindIndex(IDKind_IC����)
            txt����.ToolTipText = "IC��ʶ��"
        Case IDKind.GetKindIndex(IDKind_����)
            txt����.ToolTipText = "����Ϊ���￨�š���������ͷΪ����ID��������סԺ�š���*������š���.���Һŵ��š���/���շѵ��ݺ�"
        Case IDKind.GetKindIndex(IDKind_ҽ����)
            txt����.ToolTipText = "��¼��ҽ����"
        Case IDKind.GetKindIndex(IDKind_���֤��)
            txt����.ToolTipText = "�뽫���֤���ڶ�������"
    End Select
End Sub

Private Sub cboRoom_Click()
    If NeedName(cboRoom.list(cboRoom.ListIndex)) <> "" Then
        Call SeekIndexWithNo(cboDevice, NeedName(cboRoom.list(cboRoom.ListIndex)), True)
    Else
        cboDevice.ListIndex = -1
    End If
End Sub


Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo�ѱ�.hWnd, zlControl.CboMatchIndex(cbo�ѱ�.hWnd, KeyAscii))
End Sub

Private Sub cbo���ʽ_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo���ʽ.hWnd, zlControl.CboMatchIndex(cbo���ʽ.hWnd, KeyAscii))
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo����.hWnd, zlControl.CboMatchIndex(cbo����.hWnd, KeyAscii))
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo��������.hWnd, zlControl.CboMatchIndex(cbo��������.hWnd, KeyAscii))
    
    If KeyAscii = vbKeyReturn Then
        Call cbo��������_Click
    End If
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo����.hWnd, zlControl.CboMatchIndex(cbo����.hWnd, KeyAscii))
End Sub

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cbo�Ա�.hWnd, zlControl.CboMatchIndex(cbo�Ա�.hWnd, KeyAscii))
End Sub
Private Sub cboҽ��1_KeyPress(KeyAscii As Integer)
    '�����������ѡ����� ��Ժ���ң���ô����ҽ���ļ�����ҹ��ܣ�����ҽ������������¼��
    If Not mblnIsOutSideHosp Then
        Call zlControl.CboSetIndex(cboҽ��1.hWnd, zlControl.CboMatchIndex(cboҽ��1.hWnd, KeyAscii))
    End If
End Sub

Private Sub cboְҵ_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cboְҵ.hWnd, zlControl.CboMatchIndex(cboְҵ.hWnd, KeyAscii))
End Sub

Private Function CheckNoValidate() As Boolean
'------------------------------------------------
'���ܣ��жϼ����Ƿ��ظ�������ظ��ˣ��Ƿ���Լ�����
'       ����ʹ���˼������̹�����������ж�
'       1��mlngBuildType---1-�����ҵ�����0-��Ӱ��������
'       2��mlngUnicode --- ���߼��ű��ֲ���,1-���ּ��Ų��䣻0-������ˮ����
'       3��mblnCanOverWrite --- ��������ظ�
'������ ��
'���أ�True--����������False --ֹͣ����
'------------------------------------------------
Dim rsTmp As ADODB.Recordset
On Error GoTo ErrHandle
CheckNoValidate = True

'mintEditMode >= 2---�������������޸�,���� mblnRegToCheck --�Ǽ�ֱ�Ӽ��
If mintEditMode >= 2 Or mblnRegToCheck Then '�жϼ����Ƿ��ظ�
    
    If mlngBuildType = 1 Then
        '1-�����ҵ���,��ѯͬһ�������Ƿ�����ͬ�ļ���
        gstrSQL = "Select A.����,A.�Ա�,A.����,B.����ID From Ӱ�����¼ A,����ҽ����¼ B Where A.ִ�п���ID=[1] AND ����=[2] " _
                    & " AND B.ID=A.ҽ��ID AND B.���ID IS NULL"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngCurDeptId, Val(txt����))
    Else
        '0-��Ӱ��������,��ѯͬһӰ������Ƿ�����ͬ�ļ���
        gstrSQL = "Select A.����,A.�Ա�,A.����,B.����ID From Ӱ�����¼ A,����ҽ����¼ B Where Ӱ�����=[1] AND ����=[2] " _
                    & " AND B.ID=A.ҽ��ID AND B.���ID IS NULL"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrItemType, Val(txt����))
    End If
    
    If Not rsTmp.EOF Then   ' �����ظ��ļ���
        'mlngUnicode = 0--������ˮ����;   rsTmp!����ID <> mlngPatiId --���Ÿ��������˵ļ����ظ���
        If mlngUnicode = 0 Or rsTmp!����ID <> mlngPatiId Then
        
            If mblnCanOverWrite Then    '��������ظ�
                If MsgBoxD(Me, "��ǰ���������л����ظ����Ƿ������" & vbCrLf & "������Ϣ��" & Nvl(rsTmp!����) & " " _
                            & Nvl(rsTmp!�Ա�) & " " & Nvl(rsTmp!����), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    txt����.Text = Next����
                    MsgBoxD Me, "�Ѿ��������ɼ��ţ�" & txt����.Text & "�����ʵ���ٴ�ȷ����", vbInformation, gstrSysName
                    txt����.SetFocus
                    CheckNoValidate = False
                    Exit Function
                End If
            Else        '����������ظ�
                'ǿ�ƽ������滻�������ɵĺ���
                txt����.Text = Next����
                MsgBoxD Me, "��ǰ���������л����ظ������飡" & vbCrLf & "������Ϣ��" & Nvl(rsTmp!����) & " " & Nvl(rsTmp!�Ա�) & " " & Nvl(rsTmp!����) _
                    & vbCrLf & "�Ѿ��������ɼ��ţ�" & txt����.Text & "�����ʵ���ٴ�ȷ����", vbExclamation, gstrSysName
                txt����.SetFocus
                CheckNoValidate = False
                Exit Function
            End If
        End If
    End If
End If
Exit Function

ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub sutSetTxtEnable(thisBox As TextBox, blnEnable As Boolean)
    thisBox.Enabled = blnEnable
    If blnEnable = True Then
        thisBox.BackColor = vbWhite
    Else
        thisBox.BackColor = &H8000000B
    End If
End Sub
