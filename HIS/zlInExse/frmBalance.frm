VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#3.1#0"; "zlidkind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalance 
   AutoRedraw      =   -1  'True
   Caption         =   "���˽��ʵ�"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   Icon            =   "frmBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Picture         =   "frmBalance.frx":08CA
   ScaleHeight     =   8130
   ScaleWidth      =   11790
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picOwnFee 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   4170
      ScaleHeight     =   315
      ScaleWidth      =   1590
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   1620
      Begin VB.Label lblOwnFee 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   75
         TabIndex        =   78
         Top             =   30
         Width           =   120
      End
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   420
      Left            =   10305
      TabIndex        =   26
      ToolTipText     =   "�ȼ�:Esc"
      Top             =   7275
      Width           =   1410
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
      Height          =   420
      Left            =   8895
      TabIndex        =   25
      ToolTipText     =   "�ȼ���F2"
      Top             =   7260
      Width           =   1410
   End
   Begin VB.CommandButton cmd���㿨 
      Caption         =   "���㿨(&V)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7350
      TabIndex        =   74
      ToolTipText     =   "�ȼ���F5"
      Top             =   7275
      Width           =   1410
   End
   Begin VB.CommandButton cmdYB 
      Caption         =   "������֤(&Y)"
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
      Left            =   600
      TabIndex        =   71
      ToolTipText     =   "ҽ�����������֤,�ȼ�F6"
      Top             =   520
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Frame fraTitle 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   0
      TabIndex        =   29
      Top             =   -120
      Width           =   12165
      Begin MSCommLib.MSComm com 
         Left            =   8880
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.PictureBox pic״̬ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2520
         ScaleHeight     =   315
         ScaleWidth      =   3225
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   653
         Visible         =   0   'False
         Width           =   3255
         Begin VB.Label lbl���ʽ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1200
            TabIndex        =   70
            Top             =   30
            Width           =   1920
         End
         Begin VB.Label lbl״̬ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   75
            TabIndex        =   51
            Top             =   30
            Width           =   960
         End
      End
      Begin VB.ComboBox cboNO 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   10050
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   645
         Width           =   1515
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   11595
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F8"
         Top             =   630
         Width           =   465
      End
      Begin VB.TextBox txtInvoice 
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
         Left            =   7680
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   645
         Width           =   1425
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ�ݸ�ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   210
         Left            =   10920
         TabIndex        =   72
         Top             =   240
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ�ݺ�"
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
         Left            =   6900
         TabIndex        =   27
         Top             =   705
         Width           =   720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   30
         X2              =   18000
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   25000
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Label lblFlag 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   10935
         TabIndex        =   42
         Top             =   660
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   9285
         TabIndex        =   32
         Top             =   705
         Width           =   720
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���˽��ʵ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   180
         TabIndex        =   31
         Top             =   180
         Width           =   1875
      End
   End
   Begin VB.Frame fraPatient 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   0
      TabIndex        =   30
      Top             =   825
      Width           =   12165
      Begin zlIDKind.IDKindNew IDKIND 
         Height          =   345
         Left            =   570
         TabIndex        =   75
         Top             =   195
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   609
         Appearance      =   2
         IDKindStr       =   $"frmBalance.frx":0C0C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "����"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         NotContainFastKey=   "F1;F2;CTRL+F4;F6;F8;F9;F11;F12;ESC"
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txt�ѱ� 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   180
         Width           =   930
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
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
         Left            =   10350
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.TextBox txtBed 
         BackColor       =   &H00E0E0E0&
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
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox txt��ʶ�� 
         BackColor       =   &H00E0E0E0&
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
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtOld 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   4590
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   585
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   600
      End
      Begin VB.TextBox txtPatient 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1250
         MaxLength       =   100
         TabIndex        =   2
         ToolTipText     =   "�ȼ���F11"
         Top             =   180
         Width           =   1250
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
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
         Left            =   5235
         TabIndex        =   52
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
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
         Left            =   9850
         TabIndex        =   43
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblBed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8760
         TabIndex        =   38
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lbl��ʶ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
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
         Left            =   6720
         TabIndex        =   37
         Top             =   240
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
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
         Left            =   4095
         TabIndex        =   36
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
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
         Left            =   2930
         TabIndex        =   35
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   80
         TabIndex        =   34
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame fraDate 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7305
      TabIndex        =   56
      Top             =   1305
      Width           =   4860
      Begin VB.Frame fra�����ڼ� 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   105
         TabIndex        =   64
         Top             =   615
         Width           =   4665
         Begin MSMask.MaskEdBox txtEnd 
            Height          =   360
            Left            =   3050
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   14737632
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtBegin 
            Height          =   360
            Left            =   645
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   14737632
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
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
            Left            =   0
            TabIndex        =   68
            Top             =   60
            Width           =   480
         End
         Begin VB.Label lbl�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
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
            Index           =   1
            Left            =   2400
            TabIndex        =   67
            Top             =   60
            Width           =   240
         End
      End
      Begin VB.Frame fra����ʱ�� 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   105
         TabIndex        =   60
         Top             =   1395
         Width           =   4620
         Begin VB.TextBox txt���� 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
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
            Left            =   3060
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   0
            Width           =   645
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   645
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   0
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   14737632
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd hh:mm:ss"
            Mask            =   "####-##-## ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
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
            Left            =   0
            TabIndex        =   63
            Top             =   60
            Width           =   480
         End
         Begin VB.Label lbl�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
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
            Left            =   3870
            TabIndex        =   62
            Top             =   60
            Width           =   240
         End
      End
      Begin VB.OptionButton opt��; 
         Caption         =   "��;����"
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
         Left            =   1620
         TabIndex        =   14
         Top             =   255
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.OptionButton opt��Ժ 
         Caption         =   "��Ժ����"
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
         Left            =   3120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdPar 
         Caption         =   "��������(&S)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   75
         TabIndex        =   13
         ToolTipText     =   "�ȼ���F9"
         Top             =   180
         Width           =   1365
      End
      Begin VB.Frame fraסԺ�ڼ� 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   105
         TabIndex        =   57
         Top             =   1005
         Width           =   4665
         Begin MSMask.MaskEdBox txtPatiEnd 
            Height          =   360
            Left            =   3050
            TabIndex        =   17
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPatiBegin 
            Height          =   360
            Left            =   645
            TabIndex        =   16
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VB.Label lbl�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
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
            Index           =   0
            Left            =   2400
            TabIndex        =   59
            Top             =   60
            Width           =   240
         End
         Begin VB.Label lblסԺ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ"
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
            Left            =   0
            TabIndex        =   58
            Top             =   60
            Width           =   480
         End
      End
   End
   Begin VB.Frame fraBalance 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   7305
      TabIndex        =   39
      Top             =   3000
      Width           =   4860
      Begin VSFlex8Ctl.VSFlexGrid vsfMoney 
         Height          =   870
         Left            =   30
         TabIndex        =   76
         Top             =   1935
         Width           =   4785
         _cx             =   8440
         _cy             =   1535
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.TextBox txtOwe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   3720
         Width           =   1560
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDeposit 
         Height          =   1188
         Left            =   36
         TabIndex        =   19
         Tag             =   "1470"
         Top             =   408
         Width           =   4788
         _ExtentX        =   8440
         _ExtentY        =   2090
         _Version        =   393216
         Rows            =   5
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         HighLight       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.Label lblTicketCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ�����վ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2400
         TabIndex        =   69
         Top             =   3780
         Width           =   2400
      End
      Begin VB.Label lbl�����ʻ� 
         AutoSize        =   -1  'True
         Caption         =   "�ʻ����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2160
         TabIndex        =   47
         Top             =   1665
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblҽ������ 
         AutoSize        =   -1  'True
         Caption         =   "ͳ��֧��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   60
         TabIndex        =   46
         Top             =   1665
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblDeposit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ԥ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   105
         TabIndex        =   45
         Top             =   165
         Width           =   840
      End
      Begin VB.Label lblSpare 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ�����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   105
         TabIndex        =   44
         Top             =   165
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblOwe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
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
         Left            =   120
         TabIndex        =   40
         Top             =   3780
         Width           =   480
      End
   End
   Begin VB.TextBox txtMoney 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E7CFBA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1770
      MaxLength       =   10
      TabIndex        =   41
      Top             =   2220
      Visible         =   0   'False
      Width           =   885
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   4680
      Left            =   30
      TabIndex        =   10
      Top             =   1890
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   8255
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   28
      Top             =   7764
      Width           =   11796
      _ExtentX        =   20796
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2884
            MinWidth        =   882
            Picture         =   "frmBalance.frx":0CA2
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10901
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "�������"
            Object.ToolTipText     =   "�������"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1270
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "LocalParSet"
            Object.ToolTipText     =   "���ز�������"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1270
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1270
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
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
   Begin MSComctlLib.TabStrip tabCard 
      Height          =   5205
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1440
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   9181
      TabFixedWidth   =   1409
      TabFixedHeight  =   616
      HotTracking     =   -1  'True
      TabMinWidth     =   1411
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���ʱ�"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��ϸ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ŀ��ϸ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���±�"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ŀ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���յ���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���շ�Ŀ"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshQuery 
      Height          =   4770
      Left            =   30
      TabIndex        =   11
      Top             =   1815
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   8414
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483631
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      MergeCells      =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fra��ע 
      Height          =   555
      Left            =   30
      TabIndex        =   73
      Top             =   6600
      Width           =   7260
      Begin VB.TextBox txt��ע 
         Height          =   350
         Left            =   480
         MaxLength       =   50
         TabIndex        =   21
         Top             =   150
         Width           =   6735
      End
      Begin VB.Label lbl��ע 
         Caption         =   "��ע"
         Height          =   300
         Left            =   75
         TabIndex        =   20
         Top             =   210
         Width           =   375
      End
   End
   Begin VB.Frame fraAppend 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   15
      TabIndex        =   48
      Top             =   7155
      Width           =   7290
      Begin VB.Frame fra�Ҳ� 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   2880
         TabIndex        =   53
         Top             =   120
         Width           =   4410
         Begin VB.TextBox txt�ɿ� 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   690
            MaxLength       =   12
            TabIndex        =   23
            Text            =   "0.00"
            Top             =   0
            Width           =   1470
         End
         Begin VB.TextBox txt�Ҳ� 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   2940
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   0
            Width           =   1470
         End
         Begin VB.Label lbl�ɿ� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ɿ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   315
            Left            =   0
            TabIndex        =   55
            Top             =   45
            Width           =   690
         End
         Begin VB.Label lbl�Ҳ� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ҳ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   315
            Left            =   2235
            TabIndex        =   54
            Top             =   45
            Width           =   690
         End
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   795
         MaxLength       =   12
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   120
         Width           =   1860
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   60
         TabIndex        =   49
         Top             =   165
         Width           =   690
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Visible         =   0   'False
      Begin VB.Menu mnuFileZero 
         Caption         =   "��ʾ�����(&Z)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuDepositClear 
         Caption         =   "�����Ԥ��(&C)"
      End
      Begin VB.Menu mnuPopuDepositAll 
         Caption         =   "ʹ������Ԥ����(&A)"
      End
      Begin VB.Menu mnuPopuDepositBalance 
         Caption         =   "�����ʽ��ʹ��Ԥ��(&J)"
      End
      Begin VB.Menu mnuPopSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColsVisible 
         Caption         =   "��ʾ��ѡ��(&S)"
         Begin VB.Menu mnuViewToolCols 
            Caption         =   "���ݺ�(&N)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewToolCols 
            Caption         =   "Ʊ�ݺ�(&R)"
            Index           =   1
         End
         Begin VB.Menu mnuViewToolCols 
            Caption         =   "����(&D)"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu mnuViewToolCols 
            Caption         =   "���㷽ʽ(&T)"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu mnuViewToolCols 
            Caption         =   "���(&B)"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu mnuViewToolCols 
            Caption         =   "��Ԥ��(&P)"
            Checked         =   -1  'True
            Index           =   5
         End
      End
   End
End
Attribute VB_Name = "frmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'��ڲ�����
Public mlngPatientID As Long        '��ǰҪ���ʵĲ���ID
Public mbytInState As Byte          '0=����״̬(Ĭ������,����),1=���״̬
Public mbytFunc As Byte              '0-�������;1-סԺ����
Public mblnViewCancel As Boolean    '�Ƿ�鿴�����ϵ���
Public mstrInNO As String           'Ҫ��������ϵĵ��ݺ�
Public mblnNOMoved As Boolean       '�����ĵ����Ƿ��ں����ݱ���
Public mlngBillID As Long           'Ҫ�������ݵ�ID
Public mstrPrivs As String
Public mlngModul As Long
Public mstr��ҳId As String   '��ĳ�η���:0-������;1-��סԺ�ڼ��η���;��Ϊ������
Public mbln����תסԺ As Boolean 'true:����תסԺ���ýӿ�;FalseΪ����
Public mstrPepositDate As String 'ָ���ص��Ԥ������(��Ҫ��Ӧ��������תסԺ����ʱ,ʹ��ת���Ԥ�����н���)
'------------------------------------------------------------
Private mrsInfo As ADODB.Recordset '������Ϣ(����ID,����,�Ա�,����,סԺ��,����,��Ժ��־)
Private mrsBalance As ADODB.Recordset '����δ�Ს����ϸ
Private mrsDeposit As ADODB.Recordset '����ʣ��Ԥ����ϸ
Private mcurSpare As Currency '���˷������
Private mlng����ID As Long
Private mblnDel As Boolean
Private mcurTotal As Currency
Private mcur����� As Currency
Private mblnPrint As Boolean '���ݲ����Ͳ���ѡ������Ƿ��ӡƱ��
Private mstrDec As String   '���ν��ʵķ������С��λ��,ȱʡΪgstrDec
Private mblnNOCancel As Boolean '����������������ʱ��ֹȡ��
Private mintPatientRange As Integer '����������ʱ,�Ƿ�ֻ��ʾδ����õĲ���,0-���ѽ���,1-δ����,2-���δ����,3-סԺδ����
Private mblnSetPar As Boolean '���ν����Ƿ�����˽�����������

Private mblnOneCard As Boolean      '�Ƿ�������һ��ͨ�ӿ�
Private mrsOneCard As ADODB.Recordset
Private mstrOneCard As String       '����ʱ��ѡ���һ��ͨ�ӿڶ�Ӧ�Ľ��㷽ʽ
Private mstr����סԺ���� As String
Private mblnNotClearBill As Boolean 'δ�������
Private mblnNotClick As Boolean
Private mblnNoInsure As Boolean
'ҽ������--------------------
Private mrs���㷽ʽ As ADODB.Recordset
Private mstrȱʡ���� As String 'ȱʡ���㷽ʽ
Private mstrBalance As String 'ҽ�����صĸ��ֽ�����:"���㷽ʽ;���;�Ƿ������޸�|..."

Private mbln���ʽ��� As Boolean '�����Ƿ񷵻��˸��ʽ���
Private mcur������� As Currency '�����ʻ����
Private mcur�����޶� As Currency '�����ʻ�����޶�
Private mcur����͸֧ As Currency '�����ʻ�����͸֧���
Private mstrYBPati As String    'ҽ�����������Ϣ
Private mintInsure As Integer   '����ʱ,��ȡ�ĵ����е�����,�����ж��Ƿ����ֽ�,������
Private mblnҽ������ȫ�� As Boolean     '�Ƿ��в�֧�ֵ����Ͻ��㷽ʽ
Private mbytMCMode As Byte 'ҽ���������֤��ģʽ,����1-����,2-סԺ����ģʽ,0-��ʾ��ҽ��
Private mblnMC_TwoMode As Boolean '�Ƿ�֧�������סԺҽ���������֤������ģʽ
Private mblnUnload As Boolean
'ÿ�����˿�ʼʱ��ʼ(������ʾ�����ô���)
Private mstrAllTime As String '��������δ����סԺ����
Private mstrUnAuditTime As String '��������δ���סԺ����
Private mstrAllUnit As String '��������δ���ʿ���
Private mstrALLItem As String '��������δ���վݷ�Ŀ
Private mstrAllClass As String '��������δ���������
Private mstrALLChargeType As String '��������δ����շ���� '34260
Private mMinDate As Date, mMaxDate As Date
Private mblnDateMoved As Boolean '���˵ĵǼ�ʱ���Ƿ���ת������֮ǰ

'ÿ�����˽�����ʼ(��Ϊ���ʲ���)
Private mstrTime As String  '���˽��ʴ���(��ʼ="",����Ϊ"0,1,2,3...",0��ʾ��ҳIDΪ��)
Private mDateBegin As Date  '���˽��ʵĿ�ʼʱ��,��ʼΪ'1900-01-01'
Private mDateEnd As Date    '���˽��ʵĽ���ʱ��,��ʼΪ'3000-01-01'
Private mstrUnit As String '���˽��ʿ���ID��(��ʼ="",����Ϊ"0,1,2,3...",0��ʾ��������IDΪ��)
Private mstrClass As String  '��������=""-���з���(��δ����),"'����','����',..."
Private mstrChargeType As String '�շ���� '34260
Private mbytBaby As Byte '�Ƿ������Ӥ������(0-���з���,1-���˷���,2������-��mbytbaby-1��Ӥ������)
Private mstrItem As String 'Ҫ����վݷ�Ŀ
Private mbytKind As Byte  '0-����ͨ����,1-��������,2-��ͨ���ú�������

Private Const COL_��־ = 0
Private Const COL_סԺ = 1
Private Const COL_���� = 2
Private Const COL_ʱ�� = 3
Private Const COL_���ݺ� = 4
Private Const COL_��Ŀ = 5
Private Const COL_��Ŀ = 6
Private Const COL_Ӥ���� = 7
Private Const COL_ID = 8
Private Const COL_��� = 9
Private Const COL_��¼���� = 10
Private Const COL_��¼״̬ = 11
Private Const COL_ִ��״̬ = 12
Private Const COL_��ҳID = 13
Private Const COL_��������ID = 14
Private Const COL_�Ǽ�ʱ�� = 15
Private Const COL_δ���� = 16
Private Const COL_���ʽ�� = 17
Private Const COL_���� = 18

'Ԥ���嵥�б���,����ʱ
Private Const mstrDepositHeader = "ID|0|1,���ݺ�|920|1,Ʊ�ݺ�|920|1,����|940|6,���㷽ʽ|640|1,���|980|7,��Ԥ��|980|7"
'Ԥ���嵥�б���,�鿴ʱ
Private Const mstrDepositRHeader = "ID|0|1,���ݺ�|920|1,Ʊ�ݺ�|920|1,����|1160|6,���㷽ʽ|940|1,���|980|7"
Private Enum COLDeposit
    ID = 0
    ���ݺ� = 1
    Ʊ�ݺ� = 2
    ���� = 3
    ���㷽ʽ = 4
    ��� = 5
    ��Ԥ�� = 6
End Enum
Private Enum COLMoney
    C0���� = 0
    C1��� = 1
    C2���� = 2
    C3���� = 3
    C4ȱʡ = 4  '��ȡʱ���и���
End Enum

'��ǰ���������ҽ��֧�ֲ���
Private Type TYPE_MedicarePAR
    '1.���סԺ���㹲�õĲ���
    �ֱҴ��� As Boolean
    
    '2.��������õĲ���
    ���ﲡ�˽������� As Boolean
    ������봫����ϸ As Boolean
    ����Ԥ���� As Boolean
    �������_�������� As Boolean
    
    '3.סԺ�����õĲ���
    δ�����Ժ As Boolean
    ����ʹ�ø����ʻ� As Boolean
    ��Ժ��������Ժ As Boolean
    ��Ժ���˽������� As Boolean
    ��;������������ϴ����� As Boolean
    �������ú���ýӿ� As Boolean
    �������Ϻ��ӡ�ص� As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

Private Type Ty_ModulePara
    int�˿�Ʊ�� As Integer  '0-����ӡ,1-��ʾ��ӡ,2-����ʾ��ӡ;'���˺� ����:27776 ����:2010-02-04 16:49:03
    bln���ʺ�����Ϣ As Boolean    ''���˺� ����:27776 ����:2010-02-04 16:49:03
    bln���ʼ�鲡������ As Boolean '30036
    byt�ɿ�������� As Byte  '
    bytMzDeposit As Byte    '����Ԥ��ȱʡʹ�÷�ʽ:0-ȱʡ��ʹ�ý�;1-�����ʽ��ʹ��Ԥ��;2-ʹ������Ԥ��
    bln�����˿ʽ As Boolean 'True-�����˿�Ĭ�ϰ�Ԥ�����㷽ʽ False-�����˿�Ĭ���ֽ�
End Type
Private mty_ModulePara As Ty_ModulePara

'�������ѿ��Ĵ������
Private Type Ty_SquareCard
    blnExistsObjects As Boolean     '��װ�����ѿ���
    rsSquare As ADODB.Recordset
    dblˢ���ܶ� As Double
    bln������ As Boolean '��ǰ��ȡ�ĵ����ǿ�����
    strˢ������ As String   'ˢ�����㷽ʽ;���;�Ƿ������޸�|..."
End Type
Private mtySquareCard As Ty_SquareCard
Private mobjInPatient As Object
Private mblnFirst As Boolean
'Ʊ�����
Private mlngShareUseID As Long '������������ID
Private mstrUseType As String 'ʹ�����
Private mintInvoiceFormat As Integer  '��ӡ�ķ�Ʊ��ʽ,��Ʊ��ʽ���
Private mintInvoiceMode As Integer '0-����ӡ;1-�Զ���ӡ;2-ѡ���ӡ
Private mblnStartFactUseType As Boolean  '�Ƿ������˶���ʹ������Ʊ��

'-----------------------------------------------------------------------------------
'���㿨���
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Private mintԤ����� As Integer  '0-�����סԺ;1-����;2-סԺ
Private mlngCardTypeID As Long '��ǰˢ������56615
 
Private Sub zlInitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ģ�����
    '����:���˺�
    '����:2010-02-04 16:50:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mty_ModulePara
        '0-����ӡ,1-��ʾ��ӡ,2-����ʾ��ӡ;'���˺� ����:27776 ����:2010-02-04 16:49:03
        .int�˿�Ʊ�� = Val(zlDatabase.GetPara("�˿��վݴ�ӡ", glngSys, mlngModul))
        .bln���ʺ�����Ϣ = IIf(Val(zlDatabase.GetPara("���ʺ������Ϣ", glngSys, mlngModul)) = 1, True, False)
        .bln���ʼ�鲡������ = IIf(Val(zlDatabase.GetPara("���ʼ�鲡������", glngSys, mlngModul)) = 1, True, False) '30036
        '����:43153::0-�����п���;1-������ȡ�ֽ�ʱ,��������ɿ�.
        .byt�ɿ�������� = Val(zlDatabase.GetPara("���ʽɿ��������", glngSys, mlngModul, 0))
        .bytMzDeposit = Val(zlDatabase.GetPara("����Ԥ��ȱʡʹ�÷�ʽ", glngSys, mlngModul, 2))
        .bln�����˿ʽ = IIf(Val(zlDatabase.GetPara("�����˿�ȱʡ��ʽ", glngSys, mlngModul)) = 1, True, False)
    End With
End Sub

Private Sub cmd���㿨_Click()
    Dim dblTotal As Double, rsFeeList As ADODB.Recordset
    Err = 0: On Error GoTo Errhand:

    If mtySquareCard.blnExistsObjects = False Then Exit Sub
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If mrsInfo.EOF Then Exit Sub
    If Not IsNull(mrsInfo!����) Then
        ShowMsgbox "Ŀǰ���㿨��֧��ҽ������,����"
        Exit Sub
    End If

    '���㿨��һЩ��ش���
    dblTotal = Get��ˢ���
    If dblTotal <= 0 Then
         Call MsgBox("û�п�ˢ���㿨�Ľ��,����ˢ��!", vbInformation + vbDefaultButton1, gstrSysName)
         Exit Sub
    End If

    Screen.MousePointer = 11
    If zlSquareCardFeeList(rsFeeList) = False Then Exit Sub

    '���ýӿ�
    'Public Function zlBrushCardSquare(frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, ByVal rsFeeList As ADODB.Recordset, ByVal dbl������� As Double, ByRef rsSquare As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: zlBrushCardSquare (ˢ������ӿ�)
    '���:frmMain:HIS���� ���õ�������
    '     intCallType : HIS���� 0-  ������õ��� 1-  סԺ���ʵ���
    '     rsFeeList: HIS���� ���������൥��,�����е��ݵ���ϸ,�����סԺ���� , ���Ǳ��ν��ʵ�������ϸ
    '     dbl������� :  HIS���� ��ʾˢ�����ܳ����˽��
    '
    '����:rsSquare : �ӿڷ���    ���ؼ�¼��:�ӿڴ���սṹ(�ӿڷ�����ص�����) , �ṹ����:
    '                �ӿڱ�� , ���ѿ�ID, ���㷽ʽ, ������, ���ſ�����, ������ˮ��, ����ʱ��, ��ע
    '     rsSquare˵��:��Ҫ�ǽ��ͬһ����,ˢ���ſ����ѵ����.,�������ˢ���ſ� , ����ӿ����Ѿ�ˢ���Ŀ���Ϣ
    '     rs��̯���:������� ���ѿ�ID,����,���㷽ʽ,��̯��
    '����:true:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:18:38
    '˵��:
    '    1.  �������շѽ���ʱ,HIS�ڵ�"���㿨"ʱ,���ñ��ӿ�
    '    2.  ��סԺ���ʽ���ʱ,HIS�ڵ�"���㿨"ʱ,���ñ��ӿ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlBrushCardSquare(Me, mlngModul, mstrPrivs, rsFeeList, dblTotal, mtySquareCard.rsSquare) = False Then
        GoTo goRestoreMouse:
    End If
    
    If mtySquareCard.rsSquare Is Nothing Then GoTo goRestoreMouse:
    If mtySquareCard.rsSquare.State <> 1 Then GoTo goRestoreMouse:
    '��Ҫ���ݷ��ؽ��,���¼��㵥��
    If mtySquareCard.rsSquare.RecordCount = 0 Then
        Set mtySquareCard.rsSquare = Nothing: GoTo goRestoreMouse:
    End If
    If סԺˢ���㿨() = False Then GoTo goRestoreMouse:


goRestoreMouse:
    Screen.MousePointer = 0
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Screen.MousePointer = 0
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then
            Call FindPati(objCard, True, txtPatient.Text)
        End If
        Exit Sub
    End If
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
    
End Sub
 
Private Sub SetOneCardBalance()
'����: ����һ��ͨ���㷽ʽ
    Dim curOneCard As Currency, strName As String
    
    If mblnOneCard And Not mobjICCard Is Nothing Then
        curOneCard = mobjICCard.GetSpare(strName)
        If curOneCard <> 0 Then
           mrsOneCard.Filter = "����='" & strName & "'"
           If mrsOneCard.RecordCount > 0 Then mstrOneCard = mrsOneCard!���㷽ʽ
        End If
        sta.Panels(2).Text = "�����:" & Format(curOneCard, "0.00") & "Ԫ"
    End If
End Sub
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub
 

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    Call FindPati(objCard, True, objPatiInfor.����)
End Sub
Private Sub mnuPopuDepositAll_Click()
    'Ԥ����ȫ�壬������˸�����
    Call ShowMoney(True, , 2)
End Sub

Private Sub mnuPopuDepositBalance_Click()
    '�����ʽ���Ԥ��
     Call ShowMoney(True, , 1)
End Sub

Private Sub mnuPopuDepositClear_Click()
    '���Ԥ�����
     Call ShowMoney(True, , 0)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    Dim objCard As Card
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Set objCard = IDKIND.GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    Call FindPati(objCard, True, strCardNo)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    Dim objCard As Card
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Set objCard = IDKIND.GetIDKindCard("���֤", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    Call FindPati(objCard, True, strID)
End Sub



Private Sub SetDisibleColor(Optional bln As Boolean = False)
    If Not bln Then
        txtPatient.BackColor = &HE0E0E0
        txtPatiBegin.BackColor = &HE0E0E0
        txtPatiEnd.BackColor = &HE0E0E0
        txtTotal.BackColor = &HE0E0E0
        txtInvoice.BackColor = &HE0E0E0
    Else
        txtPatient.BackColor = &HFFFFFF
        txtPatiBegin.BackColor = &HFFFFFF
        txtPatiEnd.BackColor = &HFFFFFF
        txtTotal.BackColor = &HFFFFFF
        txtInvoice.BackColor = &HFFFFFF
    End If
End Sub

Private Sub InitPatiVariable()
'��ʼ��ÿ�����˽���������صı���
    mstrTime = "":  mstrUnit = "": mstrClass = "": mbytBaby = 0: mstrItem = "": mbytKind = 0
    If mblnNoInsure = False Then mstrChargeType = ""
    mDateBegin = CDate("0:00:00"): mDateEnd = CDate("0:00:00")
End Sub

Private Sub InitBalanceCondition()
'��ʼ��ÿ�����˽���������صı���
    mstrAllTime = "":  mstrAllUnit = "": mstrALLItem = "": mstrAllClass = "": mstrUnAuditTime = ""
    mstrALLChargeType = ""  '34260
    mMinDate = #1/1/1900#: mMaxDate = #1/1/1900#
    mblnSetPar = False
End Sub

Private Sub chkCancel_Click()
    Dim i As Long, blnNew As Boolean
            
    blnNew = (chkCancel.Value = 0)
    IDKIND.Enabled = blnNew
    If blnNew Then cboNO.Text = "": mstrInNO = ""
    
    Call NewBill    '���е�InitBalanceSet������һЩ�ؼ�״̬
    
    txtInvoice.Locked = Not blnNew
    cboNO.Locked = blnNew
    
    fraPatient.Enabled = blnNew
    cmdYB.Visible = blnNew
    cmdPar.Visible = blnNew
    opt��Ժ.Visible = blnNew
    opt��;.Visible = blnNew
    fraסԺ�ڼ�.Enabled = blnNew
    txt��ע.Enabled = blnNew: lbl��ע.Enabled = blnNew
    fra�Ҳ�.Visible = blnNew
    lblSpare.Visible = False
    txtTotal.Locked = (Not blnNew) Or (InStr(mstrPrivs, ";��������;") = 0)
    cmd���㿨.Visible = False ' blnNew And mtySquareCard.blnExistsObjects

    Call SetDisibleColor(blnNew)
        
    If Not blnNew Then
        For i = tabCard.Tabs.Count To 2 Step -1
            tabCard.Tabs.Remove i
        Next
        tabCard.SelectedItem = tabCard.Tabs(1)
        Call tabCard_Click
                
        chkCancel.ForeColor = &HFF&
        txtInvoice.Text = ""
        If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
    Else
        tabCard.Tabs.Add 2, , "��ϸ��"
        tabCard.Tabs.Add 3, , "��Ŀ��ϸ"
        tabCard.Tabs.Add 4, , "�����"
        tabCard.Tabs.Add 5, , "���±�"
        tabCard.Tabs.Add 6, , "��Ŀ��"
        tabCard.Tabs.Add 7, , "���յ���"
        tabCard.Tabs.Add 8, , "���շ�Ŀ"
        
        chkCancel.ForeColor = 0
        Call ReInitPatiInvoice
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    txtMoney.Visible = False
    If mbytInState = 0 Then
        '����:
        If mty_ModulePara.bln���ʺ�����Ϣ And mblnNotClearBill Then
            If mrsInfo Is Nothing Then
                Call NewBill
                mblnNotClearBill = False
                Exit Sub
            ElseIf mrsInfo.State <> 1 Then
                Call NewBill
                 mblnNotClearBill = False
                Exit Sub
            End If
        End If
        
        If chkCancel.Value = Checked And txtPatient.Text <> "" Then
            If MsgBox("ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            If mbytMCMode = 1 Then
                If MsgBox("ȷʵҪȡ����ǰ�������֤����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    If YBIdentifyCancel Then Call NewBill
                    Exit Sub
                    '���˳�����,�Ա�ѡ���������˽��������֤
                End If
            Else
                If Val(txtTotal.Text) <> 0 And mrsInfo.State = adStateOpen Then
                    If MsgBox("�ò�����δȷ������,ȷʵȡ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    Else
                        Call NewBill
                        Exit Sub
                    End If
                ElseIf txtPatient.Text <> "" Then
                    If MsgBox("ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
            End If
        End If
    End If
    Unload Me
End Sub


Private Function YBIdentifyCancel() As Boolean
'���ܣ�ȡ��ҽ�����������֤
'���أ����ؼ�ʱ���˳�������������
    Dim lng����ID As Long
    YBIdentifyCancel = True
    
    If mstrYBPati <> "" Then
        If UBound(Split(mstrYBPati, ";")) >= 8 Then lng����ID = Val(Split(mstrYBPati, ";")(8))
        If lng����ID <> 0 Then YBIdentifyCancel = gclsInsure.IdentifyCancel(0, lng����ID, mrsInfo!����)
    End If
End Function

Private Function GetPatientState() As Integer
'����:��ȡ����״̬
'����:0-��Ժ,1-��Ժ,2-Ԥ��Ժ,-1-�������ݿ����
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    GetPatientState = -1
    On Error GoTo errH
    strSql = "Select A.��ǰ����ID,B.״̬ From ������Ϣ A,������ҳ B " & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And Nvl(B.��ҳID,0)=[2] And A.����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CLng(mrsInfo!����ID), Val("" & mrsInfo!��ҳID))
    
    If rsTmp.RecordCount > 0 Then
        If Not IsNull(rsTmp!��ǰ����id) Then
            If Val("" & rsTmp!״̬) = 3 Then
                GetPatientState = 2
            Else
                GetPatientState = 1
            End If
        Else
            GetPatientState = 0
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub DelBalance()
    Dim blnTrans As Boolean, blnTransMC As Boolean
    Dim strSql As String, i As Long, lng����ID As Long, str���NO As String, strBalance As String, strAdvance As String
    Dim curDeposit As Currency, blnAdded As Boolean, intCashRow As Integer, curRetuCash As Currency
    Dim rsOneCard As ADODB.Recordset, objICCard As Object, strCardNo As String
    Dim strNo As String, lng����ID As Long, lng����ID As Long
    If InStr(1, mstrPrivs, ";Ԥ�����ֽ�;") > 0 Then
        curDeposit = Val(lblDeposit.Tag)
        If curDeposit <> 0 Then
            For i = 1 To vsfMoney.Rows - 1
                If vsfMoney.TextMatrix(i, COLMoney.C3����) = 1 Then intCashRow = i
            Next
            If intCashRow > 0 Then
                curRetuCash = CentMoney(curDeposit)
                If curRetuCash <> 0 Then
                    If MsgBox("��Ҫ������ʱ�����Ԥ����" & curRetuCash & "Ԫ��Ϊ�ֽ���?", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then
                        curDeposit = 0
                    Else
                        If curRetuCash <> curDeposit Then
                            '֮ǰmcur������¼�������ҽ����֧�ֻ������ֽ������
                            mcur����� = mcur����� + (curRetuCash - curDeposit)
                            curDeposit = curRetuCash
                        End If
                    End If
                Else
                    curDeposit = 0
                End If
            Else
                curDeposit = 0
            End If
        End If
    End If
    If mintInsure > 0 Or curDeposit <> 0 Then
        '�ռ��˿ʽ�����
        If Not mblnҽ������ȫ�� Or curDeposit <> 0 Then
            With vsfMoney
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, 1)) <> 0 Then '���㷽ʽ|������|�������||......  �������Ϊ��ʱ,�Կո�ֿ�,�Ա�����|��||,
                       If .TextMatrix(i, COLMoney.C3����) = 1 Then blnAdded = True
                       strBalance = strBalance & "||" & .TextMatrix(i, COLMoney.C0����) & "|" & Val(.TextMatrix(i, COLMoney.C1���)) + IIf(.TextMatrix(i, COLMoney.C3����) = 1, curDeposit, 0) & "|" & _
                                IIf(.TextMatrix(i, COLMoney.C2����) = "", " ", .TextMatrix(i, COLMoney.C2����))
                    End If
                Next
                If Not blnAdded And curDeposit <> 0 Then
                    strBalance = strBalance & "||" & .TextMatrix(intCashRow, COLMoney.C0����) & "|" & curDeposit & "| "
                End If
            End With
        End If
    End If
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    
    strNo = cboNO.Text
    lng����ID = GetBalanceID(cboNO.Text)
    
'''    '���˺� ����:���ѿ����� ����:2010-01-14 09:58:02
'''    If zlIsCheckCanelFee(lng����ID, False) = False Then Exit Sub
    If mblnOneCard Then
        Set rsOneCard = GetOneCardBalance(lng����ID)
        If rsOneCard.RecordCount > 0 Then
            On Error Resume Next
            Set objICCard = CreateObject("zlICCard.clsICCard")
            On Error GoTo 0
            If objICCard Is Nothing Then
                MsgBox "һ��ͨ�ӿڴ���ʧ��,���ܽ����˷�!����ӿ��ļ�.", vbInformation, gstrSysName
                Exit Sub
            End If
            strCardNo = objICCard.Read_Card(Me)
            If strCardNo = "" Then Exit Sub
            If strCardNo <> rsOneCard!��λ�ʺ� Then
                MsgBox "��ǰ������ۿ�Ų�һ��!���ܽ����˷�.", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
        
        
    strSql = "zl_���˽��ʼ�¼_Delete('" & cboNO.Text & "','" & UserInfo.��� & "','" & UserInfo.���� & "'," & mcur����� & _
                "," & "'" & strBalance & "'," & IIf(curDeposit <> 0, "1", "0") & ")"
    
        
    If gblnBillPrint Then
        If gobjBillPrint.zlEraseBill("", lng����ID) = False Then Exit Sub
    End If
    
    
    cmdOK.Enabled = False   '��ֹҽ����ʱ
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        
        '���սӿ�
        blnTransMC = False
        If mintInsure <> 0 Then
            If mbytMCMode = 1 Then
                If MCPAR.���ﲡ�˽������� Then
                    strAdvance = "1|1"
                    If Not gclsInsure.ClinicDelSwap(lng����ID, , mintInsure, strAdvance) Then
                        gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Sub
                    Else
                        blnTransMC = True
                    End If
                End If
            Else
                If Not gclsInsure.SettleDelSwap(lng����ID, mintInsure) Then
                    gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Sub
                Else
                    blnTransMC = True
                End If
            End If
        ElseIf Not rsOneCard Is Nothing Then
            If rsOneCard.RecordCount > 0 Then
                If Not objICCard.ReturnSwap(rsOneCard!��λ�ʺ�, rsOneCard!ҽԺ����, "" & rsOneCard!�������, rsOneCard!���) Then
                    gcnOracle.RollbackTrans
                    MsgBox "һ��ͨ�˷ѽ��׵���ʧ�ܣ��˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
                    cmdOK.Enabled = True: Exit Sub
                End If
            End If
        End If
        
        '4.�����㴦��
        If zlCallSquare_DelFree(lng����ID) = False Then
            '�����������,�ڹ����оͻ�����
            cmdOK.Enabled = True: Exit Sub
        End If
                
    gcnOracle.CommitTrans: blnTrans = False
    If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(mbytMCMode = 1, ����Enum.Busi_ClinicDelSwap, ����Enum.Busi_SettleDelSwap), True, mintInsure)
    cmdOK.Enabled = True   '��ֹҽ����ʱ
    
    If Not gobjTax Is Nothing And gblnTax Then
        gstrTax = gobjTax.zlTaxInErase(gcnOracle, lng����ID)
        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
    End If
    
    '����:35554
    If mintInsure <> 0 Then
        If MCPAR.�������Ϻ��ӡ�ص� And InStr(1, mstrPrivs, ";�����˷ѻص�;") > 0 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_4", Me, "����ID=" & zlGet���ʳ���ID(lng����ID), 2)
        End If
    ElseIf InStr(1, mstrPrivs, ";�����˷ѻص�;") > 0 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_4", Me, "����ID=" & zlGet���ʳ���ID(lng����ID), 2)
    End If
    lng����ID = GetDelBalanceID(strNo, lng����ID)
    Call WriteZYInforToCard(lng����ID, lng����ID, True)
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then
        'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
        If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(mbytMCMode = 1, ����Enum.Busi_ClinicDelSwap, ����Enum.Busi_SettleDelSwap), False, mintInsure)
    End If
    Call SaveErrLog
End Sub

Private Function GetOneCardMoney() As Currency
'���ܣ���ȡһ��ͨ������
    Dim i As Long
    
    For i = 1 To vsfMoney.Rows - 1
        If vsfMoney.TextMatrix(i, COLMoney.C3����) = 7 And Val(vsfMoney.TextMatrix(i, COLMoney.C1���)) <> 0 Then
            mrsOneCard.Filter = "���㷽ʽ='" & vsfMoney.TextMatrix(i, COLMoney.C0����) & "'"
            GetOneCardMoney = Val(vsfMoney.TextMatrix(i, COLMoney.C1���))
            Exit For
        End If
    Next
End Function

Private Function GetOneCardCount() As Integer
'���ܣ���ȡһ��ʹ���˼���һ��ͨ���㷽ʽ
    Dim i As Long
    
    For i = 1 To vsfMoney.Rows - 1
        If vsfMoney.TextMatrix(i, COLMoney.C3����) = 7 And Val(vsfMoney.TextMatrix(i, COLMoney.C1���)) <> 0 Then
            GetOneCardCount = GetOneCardCount + 1
        End If
    Next
End Function

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim lngSaveID As Long, i As Long, strNo As String, Curdate As Date, curDeposit As Currency, cur���ѽ�� As Currency, curOneCard As Currency
    Dim blnOut As Boolean, intState As Integer, strInfo As String, strTmp As String, strTime As String
    Dim bln��ӡ�˿��վ� As Boolean, str����ԭ�� As String
    Dim bln��ӡ������ϸ As Boolean, bln�Է��嵥 As Boolean
    Dim blnPrintBillEmpty As Boolean   '55052
    
    If chkCancel.Value = 1 Then '���Ͻ��ʵ�
        If mintInsure > 0 And Not MCPAR.��Ժ���˽������� And mbytMCMode <> 1 Then
            If Not isYBPati(CLng(txtPatient.Tag), True) Then
                MsgBox "�òα������Ѿ���Ժ���������ϸý��ʵ���", vbInformation, gstrSysName: Exit Sub
            End If
        End If
        If MsgBox("ȷʵҪ������[" & cboNO.Text & "]������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        '���˺�:28947
        If mintInsure <> 0 Then
            If gclsInsure.CheckInsureValid(mintInsure) = False Then
                Exit Sub
            End If
        End If
        Call DelBalance
        chkCancel.Value = 0 '(�������¼�)
    Else '�µ�����
        txtMoney.Visible = False
        
        '1.�����߼����
        If mrsInfo.State = 0 Then
            MsgBox "û��ȷ�����ʲ���,���ܴ��̣�", vbExclamation, gstrSysName
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            Exit Sub
        End If
        
        
        '����סԺʱ����Ч���ж�
        If txtPatiBegin.Text <> "____-__-__" And Not IsDate(txtPatiBegin.Text) Then
            MsgBox "������һ����Ч�Ŀ�ʼʱ�䣡", vbInformation, gstrSysName
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            Exit Sub
        End If
        If txtPatiEnd.Text <> "____-__-__" And Not IsDate(txtPatiEnd.Text) Then
            MsgBox "������һ����Ч�Ľ���ʱ�䣡", vbInformation, gstrSysName
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            Exit Sub
        End If
        If IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
            If txtPatiEnd < txtPatiBegin.Text Then
                MsgBox "����ʱ�䲻��С�ڿ�ʼʱ�䣡", vbInformation, gstrSysName
                If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
                Exit Sub
            End If
        End If
        If IsDate(txtPatiBegin.Text) And Not IsDate(txtPatiEnd.Text) Then
            MsgBox "��һ��������Ч�Ľ���ʱ�䣡", vbInformation, gstrSysName
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            Exit Sub
        End If
        If Not IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
            MsgBox "��һ��������Ч�Ŀ�ʼʱ�䣡", vbInformation, gstrSysName
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            Exit Sub
        End If
            
        If mshDetail.Rows = 2 And mshDetail.TextMatrix(1, 0) = "" Then
            MsgBox "�������²���û����Ҫ���ʵķ��ã�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If CCur(txtOwe.Text) <> 0 Then
            If CCur(txtOwe.Text) > 0 Then
                MsgBox "���˽ɿ��,�밴����ʾ�Ĳ��", vbExclamation, gstrSysName
                vsfMoney.SetFocus: Exit Sub
            Else
                MsgBox "���˽ɿ����,�밴����ʾ�Ĳ��˲��ˣ�", vbExclamation, gstrSysName
                vsfMoney.SetFocus: Exit Sub
            End If
        End If
        '43153
        '�ɿ����:0-�����п���;1-������ȡ�ֽ�ʱ,��������ɿ�.
        If mty_ModulePara.byt�ɿ�������� <> 0 And Val(txt�Ҳ�.Tag) < 0 And Val(txt�ɿ�.Text) = 0 Then
            MsgBox "�㻹δ����ɿ���,���ܼ���", vbExclamation, gstrSysName
            If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
            zlControl.TxtSelAll txt�ɿ�: Exit Sub
        End If
        '���˺�:����:25596
        If zlCommFun.StrIsValid(txt��ע.Text, 50, txt��ע.hWnd, "��ע") = False Then Exit Sub
        
        '2.ҵ�������
        If mbytMCMode <> 1 Then
            intState = GetPatientState
            If Not IsNull(mrsInfo!����) And opt��Ժ.Value Then
                If MCPAR.��Ժ��������Ժ And intState <> 0 Then
                    If IsNull(mrsInfo!��ǰ����) Then
                        MsgBox "�����ڽ����ڼ䱻������Ժ,ҽ�����˳�Ժ����ǰ�����ȳ�Ժ��", vbInformation, gstrSysName
                    Else
                        MsgBox "ҽ�����˳�Ժ����ǰ�����ȳ�Ժ��", vbInformation, gstrSysName
                    End If
                    Exit Sub
                End If
            End If
            
            '�Ƿ���Ժ
            If gbln��Ժ��׼���� And opt��Ժ.Value And (intState = 1 Or intState = 2) Then '  ' 30572:Ԥ��ԺҲ����Ժ.
                MsgBox "��ǰ������Ժ���������Ժ���ʡ�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '����Ƿ��д��շ���δ�˻�����
            If opt��Ժ.Value = True Then
                If PatiHaveStorage(mrsInfo!����ID) Then
                    Exit Sub
                End If
            End If
            
            'gbytAuditing:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
            '����:37369:��;���ʲ����
            If gbytAuditing <> 0 And opt��Ժ.Value Then
                '61345:������,2014-02-11,ֻ�����Ҫ���ʵ�סԺ�����ķ���
'                strHosTimes = ""
'                For i = 0 To frmSetBalance.lstTime.ListCount - 1
'                    If frmSetBalance.lstTime.Selected(i) = True Then strHosTimes = strHosTimes & "," & frmSetBalance.lstTime.ItemData(i)
'                Next i
'                strHosTimes = Mid(strHosTimes, 2)
'                If strHosTimes = "0" Then strHosTimes = ""
                If HaveNOAuditing(mrsInfo!����ID, mstrTime) Then
                    If gbytAuditing = 1 Then
                        '�ڶ�ȡ������Ϣʱ,�Ѿ������
                    ElseIf gbytAuditing = 2 Then
                         Call MsgBox("�ò��˻�����δ��˵ļ��ʷ���,��ֹ����!", vbInformation + vbOKOnly, gstrSysName)
                         Exit Sub
                    End If
                End If
            End If
                        
            '��Ҫ�ٴμ��,�Է������ڼ�����˵Ĳ��˱�ȡ�����
            If (InStr(mstrPrivs, ";δ��˲�����;����;") = 0 And opt��;.Value Or InStr(mstrPrivs, ";δ��˲��˳�Ժ����;") = 0 And opt��Ժ.Value) And mrsInfo!��ҳID <> 0 Then
                strTime = IIf(mstrTime = "", mstrAllTime, mstrTime)
                If strTime <> "" Then
                    For i = 0 To UBound(Split(strTime, ","))
                        strTmp = Split(strTime, ",")(i)
                        If Val(strTmp) <> 0 Then
                            If Not Chk�������(mrsInfo!����ID, Val(strTmp)) Then
                                MsgBox "�����ʷ����а������˵�" & strTmp & "��סԺδ��˵ķ��ü�¼��" & vbCrLf & _
                                    "�㲻�ܶ�δ��˵ķ��ý��н��ʣ�", vbInformation, gstrSysName
                                If cmdPar.Visible And cmdPar.Enabled Then cmdPar.SetFocus
                                Exit Sub
                            End If
                        End If
                    Next
                End If
            End If
        End If
                      
         
         '��鲡���Ƿ���δִ����ɵ�������Ŀ��δ��ҩƷ
        If opt��Ժ.Value Or mbytFunc = 0 Then
            'mbytFunc :0-�������;1-סԺ����
            'ֻ�г�Ժ���ʺ�������ʲż�� Or Not opt��Ժ.Enabled
            '����:45312
            If gbyt���δִ�� <> 0 Then
                strInfo = ExistWaitExe(mrsInfo!����ID, Nvl(mrsInfo!��ҳID, 0))
                If strInfo <> "" Then
                    If gbyt���δִ�� = 1 Then
                        If MsgBox("���ֲ���" & mrsInfo!���� & "������δִ����ɵ����ݣ�" & _
                            vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "Ҫ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
                            Exit Sub
                        End If
                    Else
                        MsgBox "���ֲ���" & mrsInfo!���� & "������δִ����ɵ����ݣ�" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "�������Ժ����.", vbInformation, gstrSysName
                        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
                        Exit Sub
                    End If
                End If
            End If
            
            '����:33048
            If gbyt���δ��ҩ <> 0 Then
                    strInfo = ExistWaitDrug(mrsInfo!����ID, Nvl(mrsInfo!��ҳID, 0), 1)
                    If strInfo <> "" Then
                        If gbyt���δ��ҩ = 1 Then
                            If MsgBox("���ֲ���" & mrsInfo!���� & strInfo & vbCrLf & vbCrLf & "Ҫ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
                                Exit Sub
                            End If
                        Else
                            MsgBox "���ֲ���" & mrsInfo!���� & strInfo & vbCrLf & vbCrLf & "�������Ժ���ʡ�", vbInformation, gstrSysName
                            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
                            Exit Sub
                        End If
                    End If
            End If
        End If
        
        If gblnAutoOut And Not IsNull(mrsInfo!��ǰ����id) And opt��Ժ.Value And mbytMCMode <> 1 Then
            If GetUnAuditReFee(mrsInfo!����ID, Nvl(mrsInfo!��ҳID, 0)) Then
                If MsgBox("����" & txtPatient.Text & "�����������˷ѵ�δ��˵ļ�¼,ȷ��Ҫ���г�Ժ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        

        If Val(txtTotal.Text) <= 0 Then
            If MsgBox("����ʵ��û�пɽ����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
                Exit Sub
            End If
        ElseIf MsgBox("��ȷ��Ҫ�Ըò��˽��н�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            Exit Sub
        End If
        
        If gbln������֤ Then
            curDeposit = 0
            For i = 1 To mshDeposit.Rows - 1
                curDeposit = curDeposit + Val(mshDeposit.TextMatrix(i, COLDeposit.��Ԥ��))
            Next
            strTime = IIf(mstrTime = "", mstrAllTime, mstrTime)
            If strTime = "0" And curDeposit <> 0 Then
                If Not zlDatabase.PatiIdentify(Me, glngSys, mrsInfo!����ID, curDeposit) Then Exit Sub
            End If
        End If
        '30036
        If mty_ModulePara.bln���ʼ�鲡������ And opt��Ժ.Value = True Then
            If IsCheck�����ѽ���(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID))) = False Then
                If MsgBox("���ֲ���" & mrsInfo!���� & "û�н��в������," & _
                    vbCrLf & "Ҫ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
                End If
                str����ԭ�� = ""
                If frmInputBox.InputBox(Me, "����δ��ԭ��", "�����벡��δ��ԭ����Ϣ:", 100, 3, True, False, str����ԭ��) = False Then
                    If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Sub
                End If
            End If
        End If
        If mblnOneCard Then
            If GetOneCardCount > 1 Then
                MsgBox "��֧��һ��ʹ�ö���һ��֧ͨ����", vbInformation, gstrSysName
                Exit Sub
            End If
            cur���ѽ�� = GetOneCardMoney
            If cur���ѽ�� <> 0 Then
                If mstrYBPati <> "" Then
                    MsgBox "��֧��ҽ������ʹ��һ��֧ͨ����", vbInformation, gstrSysName
                    Exit Sub
                End If
                If mobjICCard Is Nothing Or IsNull(mrsInfo!IC����) Then
                    MsgBox "ʹ��һ��֧ͨ�������ȶ�����", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                curOneCard = mobjICCard.GetSpare
                If curOneCard < cur���ѽ�� Then
                    MsgBox "�������" & Format(curOneCard, "0.00") & ",����Ҫ��֧�����" & Format(cur���ѽ��, "0.00"), vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        bln��ӡ�˿��վ� = False
        If mty_ModulePara.int�˿�Ʊ�� <> 0 And InStr(1, mstrPrivs, ";���˽����˿��վ�;") > 0 Then
            '0-����ӡ,1-��ʾ��ӡ,2-����ʾ��ӡ;'���˺� ����:27776 ����:2010-02-04 16:49:03
            If mty_ModulePara.int�˿�Ʊ�� = 1 Then
               If MsgBox("���Ƿ�Ҫ��ӡ�����˽����˿��վݡ���" & vbCrLf & _
                       "   ���ǡ�����ӡ���˽����˿��վ�" & vbCrLf & _
                       "   ���񡻣�����ӡ���˽����˿��վ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    bln��ӡ�˿��վ� = True
                End If
            Else
                bln��ӡ�˿��վ� = True
            End If
        End If
         '����������:�����������ʾ
'        '34681
'        If opt��Ժ.Value Then
'            If zlCheckPatiIsDeath(Val(Nvl(mrsInfo!����ID))) = True Then
'                If MsgBox("ע��:" & vbCrLf & "    �ò����Ѿ�����,�Ƿ��������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'            End If
'        End If

        '3.Ʊ����ؼ��
        '����:27559
        If Not mblnNoInsure Then
            mblnPrint = True
            '���ղ��˸���ʹ�����������ȷ����
            Select Case mintInvoiceMode
            Case 0: mblnPrint = False '����ӡ
            Case 2  '�Զ���ӡ
                If MsgBox("�Ƿ��ӡƱ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    mblnPrint = False
                End If
            End Select
        End If
        bln��ӡ������ϸ = False
        Select Case gbytFeePrintSet
        Case 1  '��ӡ.
            If MsgBox("���Ƿ�Ҫ��ӡ�����˽��ʷ�����ϸ����" & vbCrLf & _
                    "   ���ǡ�����ӡ���˽��ʷ�����ϸ" & vbCrLf & _
                    "   ���񡻣�����ӡ���˽��ʷ�����ϸ", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    bln��ӡ������ϸ = True
            End If
        Case 0  '����ӡ
        Case 2  '��ӡ.������ʾ
            bln��ӡ������ϸ = True
        End Select
        If mblnNoInsure Then
            mblnPrint = Val(zlDatabase.GetPara("�Ƚ��Էѷ��ò���ӡ����Ʊ��", glngSys, mlngModul, "0")) = 0
            Select Case Val(zlDatabase.GetPara("�Էѷ��ô�ӡ��ʽ", glngSys, mlngModul, "0"))
                Case 2  '��ӡ.
                    If MsgBox("���Ƿ�Ҫ��ӡ�������Էѷ����嵥����" & vbCrLf & _
                            "   ���ǡ�����ӡ�����Էѷ����嵥" & vbCrLf & _
                            "   ���񡻣�����ӡ�����Էѷ����嵥", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                            bln�Է��嵥 = True
                    End If
                Case 0  '����ӡ
                Case 1  '��ӡ.������ʾ
                    bln�Է��嵥 = True
            End Select
        End If
        'Ʊ�ݺ�����
        If mblnPrint Then
            If gblnStrictCtrl Then   '�ϸ�Ʊ�ݹ���
                If Trim(txtInvoice.Text) = "" Then
                    MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
                mlng����ID = GetInvoiceGroupID(IIf(gbytInvoiceKind = 0, 3, 1), 1, mlng����ID, mlngShareUseID, txtInvoice.Text, mstrUseType)
                If mlng����ID <= 0 Then
                    Select Case mlng����ID
                        Case 0 '����ʧ��
                        Case -1
                            MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                        Case -2
                            MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                        Case -3
                            MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,����������", vbInformation, gstrSysName
                            txtInvoice.SetFocus
                    End Select
                    Exit Sub
                End If
            Else
                If Len(txtInvoice.Text) <> gbytFactLength And txtInvoice.Text <> "" Then
                    MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytFactLength & " λ��", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
            End If
        End If
        '4.����
        '-------------------------------------------------------------------------------------
        cmdOK.Enabled = False   '��ֹ���ô�ӡ�������ķ�ģ̬����,�Լ�ҽ����ʱ
        lngSaveID = SaveBalance(strNo, Curdate, str����ԭ��)
        If lngSaveID = 0 Then cmdOK.Enabled = True: Exit Sub
        
        If bln��ӡ�˿��վ� Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_3", Me, "����ID=" & lngSaveID, 2)
        End If
        'Ʊ�ݴ�ӡ
        If mblnPrint Then
       '����:44332
RePrint:
            Dim strNotValiedNos As String
            Call frmPrint.ReportPrint(1, strNo, lngSaveID, mlng����ID, mlngShareUseID, mstrUseType, txtInvoice.Text, Curdate, txt�ɿ�.Text, txt�Ҳ�.Text, , mintInvoiceFormat, blnPrintBillEmpty)
           
            If gblnStrictCtrl And blnPrintBillEmpty = False And _
                ((gbytInvoiceKind = 0 And InStr(1, mstrPrivs, ";�վݴ�ӡ;") > 0) _
                   Or (gbytInvoiceKind <> 0 And InStr(1, mstrPrivs, ";��ӡ�����շ�Ʊ��;") > 0)) Then    'blnPrintBillEmpty:55052
                   '60155
                    If zlIsNotSucceedPrintBill(3, strNo, strNotValiedNos) = True Then
                            If MsgBox("���ʵ���Ϊ[" & strNotValiedNos & "]�Ľ���Ʊ�ݴ�ӡδ�ɹ�,�Ƿ����´�ӡ����Ʊ��?", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
                    End If
            End If
        End If
        If bln��ӡ������ϸ Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_3", Me, "����ID=" & Val(Nvl(mrsInfo!����ID)), "����ID=" & lngSaveID, 2)
        End If
        If bln�Է��嵥 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_4", Me, "����ID=" & Val(Nvl(mrsInfo!����ID)), "����ID=" & lngSaveID, 2)
        End If
        '�Զ���Ժ(��Ժ����)
        If gblnAutoOut And Not IsNull(mrsInfo!��ǰ����id) And opt��Ժ.Value And mbytMCMode <> 1 And Not mblnNoInsure Then
            blnOut = True
            If Not IsNull(mrsInfo!����) And Not MCPAR.δ�����Ժ Then
                Set rsTmp = GetMoneyInfo(mrsInfo!����ID, , , 2)
                If Not rsTmp Is Nothing Then
                    If Nvl(rsTmp!�������, 0) <> 0 Then blnOut = False
                End If
            End If
            
            If gblnҽ��������ܳ�Ժ And blnOut Then
                If Not checkҽ���´��Ժҽ��(mrsInfo!����ID, mrsInfo!��ҳID) Then blnOut = False
            End If
            
            If blnOut Then
                frmAutoOut.mlng����ID = mrsInfo!����ID
                frmAutoOut.mlng��ҳID = mrsInfo!��ҳID
                frmAutoOut.mlngDepID = Val("" & mrsInfo!��ǰ����id)
                frmAutoOut.mint���� = Nvl(mrsInfo!����, 0)
                frmAutoOut.mstr�Ա� = Nvl(mrsInfo!�Ա�)
                frmAutoOut.Show 1, Me
            End If
        End If
        
        'סԺ��Ϣд��:56615
        Call WriteZYInforToCard(Val(Nvl(mrsInfo!����ID)), lngSaveID)
        If IsNull(mrsInfo!��ǰ����id) Then
            zlDatabase.SetPara "Ĭ�ϳ�Ժ����", IIf(opt��Ժ.Value = True, "1", "0"), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
        End If
        cmdOK.Enabled = True
    End If
    
    gblnOK = True
    
    
    '���˺�:
    cmdOK.Enabled = False
    cboNO.Text = ""
    
    If mblnNoInsure And mblnSetPar = False And mblnDel = False Then
        If MsgBox("�����Էѷ��ý�����ɣ��Ƿ�������н��㣿", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
            mcurSpare = Get�������(mrsInfo!����ID, 0, mintԤ�����)
            mstrChargeType = ""
            mblnNoInsure = False
            picOwnFee.Visible = False
            If mblnPrint Then Call RefreshFact
            Call ShowBalance(False)
            cmdOK.Enabled = True
            Exit Sub
        End If
    End If
    
    '���˺�:27503
    If mty_ModulePara.bln���ʺ�����Ϣ Then
        Set mrsInfo = New ADODB.Recordset
        If txtInvoice.Tag <> "" And txtInvoice.Text <> txtInvoice.Tag Then txtInvoice.Text = txtInvoice.Tag '��Ҫ��Ҫ������Ϣ,��ȷ������Ҫ�����̶�
         Dim strTemp As String
         strTemp = txtInvoice.Text
        Call ReInitPatiInvoice
        txtInvoice.Text = strTemp   '��Ҫ�ǲ�Ҫ����ϴεķ�Ʊ,�µķ�Ʊ����.tag��,�ڸı䲡��ʱ,ֱ�Ӵ�����ط���ȡ
        mblnNotClearBill = True
    Else
        Call NewBill
        Call ReInitPatiInvoice(Not mblnStartFactUseType)
    End If
    sta.Panels(2) = "������ϣ��������������˱�ʶ��"
    If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
End Sub

Private Sub InitBalanceSet(bln As Boolean)
'����:����ҽ������״̬ʱ,��ؿؼ�����
    chkCancel.Enabled = bln
    cmdYB.Enabled = bln
    txtPatient.Enabled = bln
    cmdPar.Enabled = bln
    txtPatiBegin.Enabled = bln
    txtPatiEnd.Enabled = bln
    
    If bln Then
        opt��;.Enabled = bln
        opt��Ժ.Enabled = bln: opt��Ժ.Caption = "��Ժ����"
        txtTotal.Locked = (InStr(mstrPrivs, ";��������;") = 0)
    Else
        opt��;.Enabled = bln
        opt��Ժ.Enabled = Not bln: opt��Ժ.Caption = "�������": opt��Ժ.Value = True
        txtTotal.Locked = Not bln
        If MCPAR.�������_�������� Then cmdPar.Enabled = True
    End If
End Sub

Private Sub NewBill()
'����:��ʼ�����ʽ���
    If mstrInNO = "" And mbytInState = 0 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
    End If
    
    Set mrsInfo = New ADODB.Recordset '���������Ϣ
    Set mtySquareCard.rsSquare = Nothing
    picOwnFee.Visible = False
    mstrYBPati = "": mbytMCMode = 0
    mstrOneCard = ""
'''    Call zlClear���㿨
    Call ClearDetail
    Call AdjustBalance
    Call AdjustDeposit
    Call HideMoneyInfo
    Call InitBalanceCondition
    Call InitPatiVariable
    Call InitBalanceSet(True)
    
    pic״̬.Visible = False: lbl״̬.Caption = "":  lbl���ʽ.Caption = ""
    mstr����סԺ���� = ""
    txtPatient.Text = "":    txtSex.Text = "":      txtOld.Text = ""
    txt�ѱ�.Text = "":       txt��ʶ��.Text = "":   txtBed.Text = "": txt����.Text = ""
    txtBegin.Text = "____-__-__": txtEnd.Text = "____-__-__"
    txtPatiBegin.Text = "____-__-__": txtPatiEnd.Text = "____-__-__":    txtPatiEnd.Tag = "____-__-__"
    txtDate.Text = "____-__-__ __:__:__": txt����.Text = ""
    txt��ע.Text = ""
    lblBed.Visible = False:     txtBed.Visible = False
    lbl��ʶ��.Visible = False:  txt��ʶ��.Visible = False
    lbl����.Visible = False:    txt����.Visible = False
    
    lblSpare.Caption = "Ԥ�����:"
    lblSpare.Tag = ""
    sta.Panels(3).Text = ""
    lblDeposit.Caption = "��Ԥ��:"
    lblDeposit.Tag = ""
    lblTicketCount.Caption = "Ԥ�����վ�:"
    
    cmdOK.Enabled = True
    
    sta.Panels(2) = ""
End Sub
Private Sub cmdPar_Click()
    Dim blnAll As Boolean, i As Long
    If mrsInfo.State = 0 Then
        MsgBox "û��ȷ�����ʲ���,�������ý��ʲ�����", vbExclamation, gstrSysName
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Sub
    End If
    
    With frmSetBalance
        .mstrUnAuditTime = mstrUnAuditTime
        .mblnNOCancel = mblnNOCancel
        .mlngInsure = Val("" & mrsInfo!����)
        .mlngPatient = mrsInfo!����ID
        .mstrAllTime = mstrAllTime
        .mstrAllUnit = mstrAllUnit
        .mstrALLItem = mstrALLItem
        .mstrALLChargeType = mstrALLChargeType '34260
        .mstrAllClass = mstrAllClass
        .mMinDate = mMinDate
        .mMaxDate = mMaxDate
        .mbytKind = mbytKind
        .mbln������ʽ��� = mbytMCMode = 1
        .mbytFunc = mbytFunc
        .mblnEditFee = Not mblnNoInsure
        .Show 1, Me
    
    
        Me.Refresh
        If .mblnOk Then
            mblnSetPar = True
            'ȡ��������
            Call InitPatiVariable
            '��������
            mstrClass = ""
            If Not .lstClass.Selected(0) Then
                For i = 1 To .lstClass.ListCount - 1
                    If .lstClass.Selected(i) Then
                        mstrClass = mstrClass & ",'" & .lstClass.List(i) & "'"
                    End If
                Next
            End If
            
            If mblnNoInsure = False Then
                '�շ����:34260
                mstrChargeType = ""
                Dim objList As ListItem
                With .lvwChargeType
                    If .ListItems("ALL").Checked = False Then
                        For Each objList In .ListItems
                            If objList.Key <> "ALL" And objList.Checked Then
                                mstrChargeType = mstrChargeType & ",'" & Mid(objList.Key, 2) & "'"
                            End If
                        Next
                    End If
                End With
            End If
            
            'Ӥ����
            mbytBaby = .cboBabyFee.ListIndex
            
            '����
            mbytKind = 0
            If .chkKind(0).Value = 1 And .chkKind(1).Value = 1 Then
                mbytKind = 2
            Else
                If .chkKind(1).Value = 1 Then mbytKind = 1
            End If
            If mbytFunc = 0 Then
                mstrTime = ",0"
            Else
                If .lstTime.ListCount > 0 Then
                    blnAll = True
                    For i = 0 To .lstTime.ListCount - 1
                        If .lstTime.Selected(i) Then
                            mstrTime = mstrTime & "," & .lstTime.ItemData(i)
                        Else
                            blnAll = False
                        End If
                    Next
                    If blnAll And Not gbln����ָ��Ԥ���� Then mstrTime = ""
                End If
             End If
            If .lstUnit.ListCount > 0 Then
                blnAll = True
                For i = 0 To .lstUnit.ListCount - 1
                    If .lstUnit.Selected(i) Then
                        mstrUnit = mstrUnit & "," & .lstUnit.ItemData(i)
                    Else
                        blnAll = False
                    End If
                Next
                If blnAll Then mstrUnit = ""
            End If
            If .lstItem.ListCount > 0 Then
                blnAll = True
                For i = 0 To .lstItem.ListCount - 1
                    If .lstItem.Selected(i) Then
                        mstrItem = mstrItem & ",'" & .lstItem.List(i) & "'"
                    Else
                        blnAll = False
                    End If
                Next
                If blnAll Then mstrItem = ""
            End If
            
            '�õǼ�ʱ���ѯ,����ʱ����ʾ
            '����������ʱ,�����ڼ�
            If .chkKind(0).Value = 0 And .chkKind(1).Value = 1 Then
                mDateBegin = CDate("0:00:00")
                mDateEnd = CDate("0:00:00")
            Else
                mDateBegin = CDate(Format(.dtpBegin.Value, "yyyy-MM-dd 00:00:00"))
                mDateEnd = CDate(Format(.dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
            End If
                
            '��ʾ����ʱ��
            txtEnd.Text = Format(.dtpEnd.Value, txtEnd.Format)
            txtBegin.Text = Format(.dtpBegin.Value, txtBegin.Format)
            
            mstrTime = Mid(mstrTime, 2)
            mstrUnit = Mid(mstrUnit, 2)
            mstrItem = Mid(mstrItem, 2)
            mstrClass = Mid(mstrClass, 2)
            If mstrChargeType <> "" And mblnNoInsure = False Then mstrChargeType = Mid(mstrChargeType, 2)   '34260
            
            '��������ж��סԺ����δ�ᣬ��ֻѡ���ĳ��סԺ���ã�����ݸô�סԺ��Ϣ�����������Ƿ���ҽ������
            If mstrTime <> "" And InStr(1, mstrTime, ",") = 0 And mrsInfo!��ҳID <> mstrTime And InStr(1, mstrAllTime, ",") > 0 Then
                IDKIND.IDKIND = IDKIND.GetKindIndex("����")
                txtPatient.Text = "-" & mrsInfo!����ID
                Call LoadPatientInfo(IDKIND.GetCurCard, False, 0, Val(mstrTime))
            End If
            
            If Not ShowBalance() Then
                cmdOK.Enabled = False
                MsgBox "�������²���û����Ҫ���ʵķ��ã�", vbInformation, gstrSysName
                If cmdPar.Visible And cmdPar.Enabled Then cmdPar.SetFocus
            Else
                If vsfMoney.Visible And vsfMoney.Enabled Then vsfMoney.SetFocus
            End If
        Else
            If mblnSetPar = False And Not IsNull(mrsInfo!����) And MCPAR.�������ú���ýӿ� Then
                cmdOK.Enabled = False
            End If
        End If
    End With
End Sub

Private Sub OutputList(ByVal bytStyle As Byte)
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, lngRow As Long
    
    If mshDetail.TextMatrix(1, 0) = "" Then
        MsgBox "û�����ݣ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    objOut.Title.Text = "����" & tabCard.SelectedItem.Caption
    If tabCard.SelectedItem.Index = 1 Then
        Set objOut.Title.Font = tabCard.Font
        Set objOut.Body = mshDetail
        
        lngRow = mshDetail.Row
    Else
        Set objOut.Title.Font = tabCard.Font
        Set objOut.Body = mshQuery
        
        lngRow = mshQuery.Row
        mshQuery_LeaveCell
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "���ݺ�:" & cboNO.Text
    objRow.Add "ʵ�ʺ�:" & txtInvoice.Text
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "����:" & txtPatient.Text
    objRow.Add "סԺ��:" & txt��ʶ��.Text
    objRow.Add "�ϼ�:" & txtTotal.Text
    objOut.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "��ӡʱ��:" & Format(zlDatabase.Currentdate, "YYYY-MM-DD hh:mm:ss")
    objRow.Add "����ʱ��:" & txtDate.Text
    objOut.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    If mbytInState = 0 Then
        objRow.Add "��ע:δ����"
    ElseIf mbytInState = 1 Then
        If mblnViewCancel Then
            objRow.Add "��ע:���ϵ�"
        Else
            objRow.Add "��ע:"
        End If
    End If
    objOut.BelowAppRows.Add objRow
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
        
    If tabCard.SelectedItem.Index = 1 Then
        mshDetail.Row = lngRow
    Else
        mshQuery.Row = lngRow
        mshQuery_EnterCell
    End If
End Sub

Private Sub Form_Activate()
    '˫����ʾ��������ڵ�ǰ������ʾ֮�������ʾ�����ƶ�����
    If mblnUnload = True Then Unload Me: Exit Sub
    
    mblnFirst = False
    If mstrInNO = "" And mbytInState = 0 And gblnLED And Trim(txtPatient.Text) = "" Then
        zl9LedVoice.DisplayPatient ""
    End If
    
    If mbytInState = 1 Then
        If cmdCancel.Visible And cmdCancel.Enabled Then cmdCancel.SetFocus
    ElseIf mstrInNO <> "" Then
        '����ʱ
        If txtPatient.Text = "" Then Unload Me: Exit Sub
        If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If txtMoney.Visible Then
                txtMoney.Visible = False
                If txtMoney.Left < fraBalance.Left Then
                    mshDetail.SetFocus
                Else
                    mshDeposit.SetFocus
                End If
            Else
                'ȡ����ť
                If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus: Call cmdCancel_Click
            End If
        Case vbKeyF1
            ShowHelp App.ProductName, Me.hWnd, Me.Name
        Case vbKeyF2
            If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus: Call cmdOK_Click
        Case vbKeyF4
            If Shift = vbCtrlMask Then
                If IDKIND.Enabled Then
                    Dim intIndex As Integer
                    intIndex = IDKIND.GetKindIndex("IC����")
                    If intIndex <= 0 Then Exit Sub
                    IDKIND.IDKIND = intIndex: Call IDKind_Click(IDKIND.GetCurCard)
                End If
            ElseIf Me.ActiveControl Is txtPatient Then
                If IDKIND.Enabled Then
                    If Shift = vbShiftMask Then
                        IDKIND.IDKIND = IIf(IDKIND.IDKIND = 0, UBound(Split(IDKIND.IDKindStr, ";")), IDKIND.IDKIND - 1)
                    Else
                        IDKIND.IDKIND = IIf(IDKIND.IDKIND = UBound(Split(IDKIND.IDKindStr, ";")), 0, IDKIND.IDKIND + 1)
                    End If
                End If
            End If
        Case vbKeyF6
            If cmdYB.Enabled And cmdYB.Visible Then cmdYB.SetFocus: Call cmdYB_Click
        Case vbKeyF8 '�˺ſ��
            chkCancel.Value = IIf(chkCancel.Value = 1, 0, 1)
        Case vbKeyF9 '��������
            If cmdPar.Enabled And cmdPar.Visible Then cmdPar.SetFocus: Call cmdPar_Click
        Case vbKeyF11 '��λ�����������
            If Not txtPatient.Locked And txtPatient.Enabled Then txtPatient.SetFocus
        Case vbKeyF12 '��λ�����ſ�
            If Not cboNO.Locked And cboNO.Enabled Then cboNO.SetFocus
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strTmp As String
    If mstrInNO <> "" And mbytInState = 0 Then
        mblnDel = True
    Else
        mblnDel = False
    End If
    mblnFirst = True
    mblnUnload = False
    glngFormW = 11565: glngFormH = 8535
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
       
    mintԤ����� = 2 'ȱʡΪסԺԤ��
    Call RestoreWinState(Me, App.ProductName)
    gblnOK = False
    
    If mbytInState = 0 Then
        Set mrsOneCard = GetOneCard
        mblnOneCard = mrsOneCard.RecordCount > 0
    End If
    If InStr(1, mstrPrivs, ";���ô��۽���;") = 0 Then
        strTmp = "1,2,3,4,5,9"    '7,8:����:48810
    Else
        strTmp = "1,2,3,4,5,6,9"  '7,8:����:48810
    End If
    Set mrs���㷽ʽ = Get���㷽ʽ("����", strTmp)
    If mrs���㷽ʽ.RecordCount = 0 Then
        MsgBox "δ���ý��ʳ��Ͽ��õĽ��㷽ʽ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call InitFace
    
    If mbytInState = 0 And gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.��� & "��Ϊ������", mlngModul, gcnOracle
    End If
 

    
    If mbytInState = 1 Then                 '�鿴
        If Not ReadBalance(mstrInNO) Then mblnUnload = True: Exit Sub
    ElseIf mstrInNO <> "" Then        '����
        chkCancel.Value = 1     '����Click�¼�
        cboNO.Text = mstrInNO
        cboNO_KeyPress (13)
    Else 'ִ�н���
'        If Not CheckErrorItem Then
'            MsgBox "ϵͳ����δ������Ч��������Ŀ�����ȵ������������������á�", vbInformation, gstrSysName
'            mblnUnload = True:  Exit Sub
'        End If
        
        mintPatientRange = Val(zlDatabase.GetPara("��ʾ���岡��", glngSys, mlngModul, 0))
        If mlngPatientID <> 0 Then
            txtPatient.Text = "-" & mlngPatientID
            mstrTime = mstr��ҳId
            Call txtPatient_KeyPress(vbKeyReturn)
            If Val(mstr��ҳId) = "0" Then cmdYB.Enabled = True
            If mrsInfo.State = 0 Then mblnUnload = True: Exit Sub
        End If
    End If
    
    '����:47798
    If mbytInState = 0 Then
        Call GetRegisterItem(g˽��ģ��, Me.Name, "idkind", strTmp)
        Err = 0: On Error Resume Next
        mblnNotClick = True
        IDKIND.IDKIND = Val(strTmp)
        mblnNotClick = False
        Err = 0: On Error GoTo 0
    End If
End Sub

Private Sub RefreshFact()
    '���ܣ�ˢ���շ�Ʊ�ݺ�
    If mintInvoiceMode = 0 Then Exit Sub
    
    If gblnStrictCtrl Then
        mlng����ID = CheckUsedBill(IIf(gbytInvoiceKind = 0, 3, 1), IIf(mlng����ID > 0, mlng����ID, mlngShareUseID), , mstrUseType)
        If mlng����ID <= 0 Then
            Select Case mlng����ID
                Case 0 '����ʧ��
                Case -1
                    MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End Select
            txtInvoice.Text = ""
            txtInvoice.Tag = ""
        Else
            '�ϸ�ȡ��һ������
            txtInvoice.Text = GetNextBill(mlng����ID)
            txtInvoice.Tag = txtInvoice.Text
        End If
    Else
        '��ɢ��ȡ��һ������
        txtInvoice.Text = IncStr(UCase(zlDatabase.GetPara("��ǰ����Ʊ�ݺ�", glngSys, 1137, "")))
        txtInvoice.Tag = txtInvoice.Text
    End If
End Sub

Private Sub InitFace()
    Dim i As Long
    
    If mbytInState = 1 Then
         lblTitle.Caption = gstrUnitName & "���˽��ʵ�"
    Else
         lblTitle.Caption = gstrUnitName & IIf(mbytFunc = 0, "���ﲡ�˽��ʵ�", "סԺ���˽��ʵ�")
    End If
    
    sta.Panels("LocalParSet").Visible = mlngPatientID <> 0  '���˷��ò�ѯ�е���ʱ,�ṩ���ز�������
    
    Call zlInitModulePara
    Call initCardSquareData
    
    mblnStartFactUseType = zlStartFactUseType(IIf(gbytInvoiceKind = 0, 3, 1))
    
    If Not (mbytInState = 0 And mstrInNO <> "") Then Call NewBill    '����ʱ��chkCancel.Value = 1ʱ����
    chkCancel.Visible = (mbytInState = 0 And (InStr(";" & mstrPrivs, ";��������;") > 0))
         
    txtPatient.Width = txtPatient.Width + 400
    
    IDKIND.Enabled = (mbytInState = 0 And mstrInNO = "")
    If mbytInState = 0 And mstrInNO = "" Then
        Call ReInitPatiInvoice(Not mblnStartFactUseType)
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
        If InStr(mstrPrivs, ";���ս���;") > 0 Then
            cmdYB.Visible = True
            
            '�ɶ��ϰ�ҽ��֧�������סԺ���������֤ģʽ
            mblnMC_TwoMode = InStr("," & GetSetting("ZLSOFT", "����ȫ��", "����֧�ֵ�ҽ��", "") & ",", ",20,") > 0
            If mblnMC_TwoMode Then
                cmdYB.Caption = "ˢ"
                txtPatient.Width = txtPatient.Width - 400
                cmdYB.Left = txtPatient.Left + txtPatient.Width + 10
                cmdYB.Top = fraPatient.Top + 180
                cmdYB.Width = 400
                pic״̬.Left = txtPatient.Left
            ElseIf InStr(mstrPrivs, ";������ý���;") = 0 Or mbytFunc = 1 Then    'mbytFunc=1:סԺ����
                cmdYB.Visible = False
                pic״̬.Left = txtPatient.Left
            End If
        Else
            cmdYB.Visible = False
            pic״̬.Left = txtPatient.Left
        End If
    
        If InStr(mstrPrivs, ";��������;") = 0 Then
            cmdPar.Visible = False
            txtTotal.Locked = True
            opt��;.Left = opt��;.Left - cmdPar.Width / 2
            opt��Ժ.Left = opt��Ժ.Left - cmdPar.Width / 2
        End If
        cboNO.Text = ""
        opt��Ժ.Visible = True
        opt��;.Visible = True
        cmd���㿨.Visible = False ' mtySquareCard.blnExistsObjects
        Call InitԤ�����
    ElseIf mbytInState = 1 Then
        If mblnViewCancel Then lblFlag.Visible = True
        cmdOK.Visible = False
        cmdCancel.Caption = "�˳�(&X)"
        txtPatient.Locked = True
        txtTotal.Locked = True
        
        fra�Ҳ�.Visible = False
        txt��ע.Enabled = False: lbl��ע.Enabled = False
        cmdPar.Visible = False
        opt��Ժ.Visible = False
        opt��;.Visible = False
        
        fra�����ڼ�.Top = fra�����ڼ�.Top - cmdPar.Height
        fraסԺ�ڼ�.Top = fraסԺ�ڼ�.Top - cmdPar.Height
        fra����ʱ��.Top = fra����ʱ��.Top - cmdPar.Height
        fraDate.Height = fraDate.Height - cmdPar.Height
        fraBalance.Top = fraBalance.Top - cmdPar.Height
        
        fraTitle.Enabled = False
        fraסԺ�ڼ�.Enabled = False
        Call SetDisibleColor
        cmd���㿨.Visible = False
    End If

End Sub
Private Sub SetSortMoneyData(ByVal BytType As Byte, ByVal blnHaveMoeny As Boolean, ByVal bytEdit As Byte, _
    ByRef k As Integer, ByRef ArrSort() As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���,���ý��㷽ʽ��ʾ˳������
    '���:bytType-����(0-��ҽ��;1-ҽ��)
    '       blnHaveMoeny-true:�н��;False;�޽��
    '       bytEdit-0-�����ֱ༭;1����༭;2���ɱ༭
    '����:K-�������һ��˳����
    '       ArrSort-������������
    '����:
    '����:���˺�
    '����:2010-09-26 15:03:35
    '����:32322
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, bytTemp As Byte   '0��ҽ��;>1ҽ��
    Dim blnTempMoney As Boolean, bytTempEdit As Byte
    For i = 1 To vsfMoney.Rows - 1
        bytTemp = IIf(InStr(1, ",3,4,", Trim(vsfMoney.TextMatrix(i, COLMoney.C3����))) = 0, 0, 1)
        blnTempMoney = Val(vsfMoney.TextMatrix(i, COLMoney.C1���)) <> 0
        bytTempEdit = IIf(bytEdit = 0, 0, IIf(vsfMoney.RowData(i) = 0, 1, 2))
        If bytTemp = BytType And blnHaveMoeny = blnTempMoney And bytTempEdit = bytEdit Then
            '��������
            For j = 0 To vsfMoney.Cols - 1
                ArrSort(k, j) = vsfMoney.TextMatrix(i, j)
            Next
            '��������
            ArrSort(k, vsfMoney.Cols) = vsfMoney.RowData(i)
            vsfMoney.Row = i: vsfMoney.Col = 0
            ArrSort(k, vsfMoney.Cols + 1) = IIf(vsfMoney.CellFontBold, 1, 0)
            ArrSort(k, vsfMoney.Cols + 2) = vsfMoney.CellForeColor
            k = k + 1
        End If
    Next
End Sub
Private Sub SortMoney()
'���ܣ��������㷽ʽ���б�,ʹ�н�������ǰ��
'˵����ͬ����ԭ��˳�򲻱�
    Dim arrCell() As String, blnRedraw As Boolean
    Dim i As Integer, j As Integer, k As Integer
    Dim lngRow As Long, lngCol As Long
    Dim varData As Variant
    Dim arrTemp() As String
    
    ReDim arrTemp(0 To vsfMoney.Cols + 2)
    ReDim arrCell(1 To vsfMoney.Rows - 1, 0 To vsfMoney.Cols + 2)
    lngRow = vsfMoney.Row: lngCol = vsfMoney.Col
    blnRedraw = vsfMoney.Redraw
    vsfMoney.Redraw = False
    '����:32322

    k = 1
    varData = Split(gstr���㷽ʽ��ʾ˳��, ";")
    '��ҽ������-�н��;��ҽ������-�޽��;ҽ������-�н���������޸�;ҽ������-�޽���������޸�;ҽ������-�н���Ҳ������޸�;ҽ������-�޽���Ҳ������޸�
    For i = 0 To UBound(varData)
        Select Case varData(i)
        Case "��ҽ������-�н��"
            Call SetSortMoneyData(0, True, 0, k, arrCell)
        Case "��ҽ������-�޽��"
            Call SetSortMoneyData(0, False, 0, k, arrCell)
        Case "ҽ������-�н���������޸�"
            Call SetSortMoneyData(1, True, 1, k, arrCell)
        Case "ҽ������-�޽���������޸�"
            Call SetSortMoneyData(1, False, 1, k, arrCell)
        Case "ҽ������-�н���Ҳ������޸�"
            Call SetSortMoneyData(1, True, 2, k, arrCell)
        Case "ҽ������-�޽���Ҳ������޸�"
            Call SetSortMoneyData(1, False, 2, k, arrCell)
        Case Else
        End Select
    Next
    'Ԥ��ĳЩ���㷽ʽ������,�������������
    Dim blnFind As Boolean
    With vsfMoney
        For i = 1 To .Rows - 1
            blnFind = False
            For j = 1 To UBound(arrCell)
                If .TextMatrix(i, COLMoney.C0����) = arrCell(j, COLMoney.C0����) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If blnFind = False Then
                'δ�ҵ�����,��Ҫ���¼�����ȥ
                For j = 0 To vsfMoney.Cols - 1
                    arrCell(k, j) = vsfMoney.TextMatrix(i, j)
                Next
                '��������
                arrCell(k, vsfMoney.Cols) = vsfMoney.RowData(i)
                vsfMoney.Row = i: vsfMoney.Col = 0
                arrCell(k, vsfMoney.Cols + 1) = IIf(vsfMoney.CellFontBold, 1, 0)
                arrCell(k, vsfMoney.Cols + 2) = vsfMoney.CellForeColor
                k = k + 1
            End If
        Next
    End With
    
'''    '���㷽ʽ����:-1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����
'''    '����:27374
'''    '1    ?HIS�Ľ��㷽ʽ����ǰ��?
'''    '2������ҽ���ӿڷ��ص���Ϣ�����޸Ľ��㷽ʽ����ǰ�������н��Ľ��㷽ʽ������ǰ��
'''
'''    '��ȡHIS�Ľ��㷽ʽ
'''
'''
'''
'''    '��ȡHIS�����н𲿷�������ǰ��
'''    K = 1
'''    For i = 1 To vsfMoney.Rows - 1
'''        If Val(vsfMoney.TextMatrix(i, COLMoney.C1���)) <> 0 And InStr(1, ",3,4,", Trim(vsfMoney.TextMatrix(i, COLMoney.C3����))) = 0 Then
'''            For j = 0 To vsfMoney.Cols - 1
'''                arrCell(K, j) = vsfMoney.TextMatrix(i, j)
'''            Next
'''            '��������
'''            arrCell(K, vsfMoney.Cols) = vsfMoney.RowData(i)
'''            vsfMoney.Row = i: vsfMoney.Col = 0
'''            arrCell(K, vsfMoney.Cols + 1) = IIf(vsfMoney.CellFontBold, 1, 0)
'''            arrCell(K, vsfMoney.Cols + 2) = vsfMoney.CellForeColor
'''            K = K + 1
'''        End If
'''    Next
'''
'''    '��ȡHIS�޽��Ľ��㷽ʽ
'''    For i = 1 To vsfMoney.Rows - 1
'''        If Val(vsfMoney.TextMatrix(i, COLMoney.C1���)) = 0 And InStr(1, ",3,4,", Trim(vsfMoney.TextMatrix(i, COLMoney.C3����))) = 0 Then
'''            For j = 0 To vsfMoney.Cols - 1
'''                arrCell(K, j) = vsfMoney.TextMatrix(i, j)
'''            Next
'''
'''            '��������
'''            arrCell(K, vsfMoney.Cols) = vsfMoney.RowData(i)
'''            vsfMoney.Row = i: vsfMoney.Col = 0
'''            arrCell(K, vsfMoney.Cols + 1) = IIf(vsfMoney.CellFontBold, 1, 0)
'''            arrCell(K, vsfMoney.Cols + 2) = vsfMoney.CellForeColor
'''
'''            K = K + 1
'''        End If
'''    Next
'''
'''    '-------------------------------------------------------------------------------------------------------------------------------------------------------------
'''    '--ҽ���Ĵ���
'''    '��ȡҽ���ȿ��޸����н��Ľ��㷽ʽ
'''    For i = 1 To vsfMoney.Rows - 1
'''        If Val(vsfMoney.TextMatrix(i, COLMoney.C1���)) <> 0 And vsfMoney.RowData(i) = 0 And InStr(1, ",3,4,", Trim(vsfMoney.TextMatrix(i, COLMoney.C3����))) > 0 Then
'''            For j = 0 To vsfMoney.Cols - 1
'''                arrCell(K, j) = vsfMoney.TextMatrix(i, j)
'''            Next
'''
'''            '��������
'''            arrCell(K, vsfMoney.Cols) = vsfMoney.RowData(i)
'''            vsfMoney.Row = i: vsfMoney.Col = 0
'''            arrCell(K, vsfMoney.Cols + 1) = IIf(vsfMoney.CellFontBold, 1, 0)
'''            arrCell(K, vsfMoney.Cols + 2) = vsfMoney.CellForeColor
'''            K = K + 1
'''        End If
'''    Next
'''
'''    '��ȡҽ���ȿ��޸����޽��Ľ��㷽ʽ
'''    For i = 1 To vsfMoney.Rows - 1
'''        If Val(vsfMoney.TextMatrix(i, COLMoney.C1���)) = 0 And vsfMoney.RowData(i) = 0 And InStr(1, ",3,4,", Trim(vsfMoney.TextMatrix(i, COLMoney.C3����))) > 0 Then
'''            For j = 0 To vsfMoney.Cols - 1
'''                arrCell(K, j) = vsfMoney.TextMatrix(i, j)
'''            Next
'''
'''            '��������
'''            arrCell(K, vsfMoney.Cols) = vsfMoney.RowData(i)
'''            vsfMoney.Row = i: vsfMoney.Col = 0
'''            arrCell(K, vsfMoney.Cols + 1) = IIf(vsfMoney.CellFontBold, 1, 0)
'''            arrCell(K, vsfMoney.Cols + 2) = vsfMoney.CellForeColor
'''            K = K + 1
'''        End If
'''    Next
'''
'''    '��ȡҽ���Ȳ����޸����н��Ľ��㷽ʽ
'''    For i = 1 To vsfMoney.Rows - 1
'''        If Val(vsfMoney.TextMatrix(i, COLMoney.C1���)) <> 0 And vsfMoney.RowData(i) = 1 And InStr(1, ",3,4,", Trim(vsfMoney.TextMatrix(i, COLMoney.C3����))) > 0 Then
'''            For j = 0 To vsfMoney.Cols - 1
'''                arrCell(K, j) = vsfMoney.TextMatrix(i, j)
'''            Next
'''
'''            '��������
'''            arrCell(K, vsfMoney.Cols) = vsfMoney.RowData(i)
'''            vsfMoney.Row = i: vsfMoney.Col = 0
'''            arrCell(K, vsfMoney.Cols + 1) = IIf(vsfMoney.CellFontBold, 1, 0)
'''            arrCell(K, vsfMoney.Cols + 2) = vsfMoney.CellForeColor
'''            K = K + 1
'''        End If
'''    Next
'''    '��ȡҽ���Ȳ����޸����޽��Ľ��㷽ʽ
'''    For i = 1 To vsfMoney.Rows - 1
'''        If Val(vsfMoney.TextMatrix(i, COLMoney.C1���)) = 0 And vsfMoney.RowData(i) = 1 And InStr(1, ",3,4,", Trim(vsfMoney.TextMatrix(i, COLMoney.C3����))) > 0 Then
'''            For j = 0 To vsfMoney.Cols - 1
'''                arrCell(K, j) = vsfMoney.TextMatrix(i, j)
'''            Next
'''            '��������
'''            arrCell(K, vsfMoney.Cols) = vsfMoney.RowData(i)
'''            vsfMoney.Row = i: vsfMoney.Col = 0
'''            arrCell(K, vsfMoney.Cols + 1) = IIf(vsfMoney.CellFontBold, 1, 0)
'''            arrCell(K, vsfMoney.Cols + 2) = vsfMoney.CellForeColor
'''            K = K + 1
'''        End If
'''    Next
'''
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------

''    '��ȡ�н���
''    k = 1
''    For i = 1 To vsfMoney.Rows - 1
''        If Val(vsfMoney.TextMatrix(i, COLMoney.C1���)) <> 0 Then
''            For j = 0 To vsfMoney.Cols - 1
''                arrCell(k, j) = vsfMoney.TextMatrix(i, j)
''            Next
''
''            '��������
''            arrCell(k, vsfMoney.Cols) = vsfMoney.RowData(i)
''            vsfMoney.Row = i: vsfMoney.Col = 0
''            arrCell(k, vsfMoney.Cols + 1) = IIf(vsfMoney.CellFontBold, 1, 0)
''            arrCell(k, vsfMoney.Cols + 2) = vsfMoney.CellForeColor
''
''            k = k + 1
''        End If
''    Next
''
''    '��ȡ�޽���
''    For i = 1 To vsfMoney.Rows - 1
''        If Val(vsfMoney.TextMatrix(i, COLMoney.C1���)) = 0 Then
''            For j = 0 To vsfMoney.Cols - 1
''                arrCell(k, j) = vsfMoney.TextMatrix(i, j)
''            Next
''
''            '��������
''            arrCell(k, vsfMoney.Cols) = vsfMoney.RowData(i)
''            vsfMoney.Row = i: vsfMoney.Col = 0
''            arrCell(k, vsfMoney.Cols + 1) = IIf(vsfMoney.CellFontBold, 1, 0)
''            arrCell(k, vsfMoney.Cols + 2) = vsfMoney.CellForeColor
''
''            k = k + 1
''        End If
''    Next

    '����������ǰ
    For i = 1 To vsfMoney.Rows - 1
        If Val(arrCell(i, COLMoney.C3����)) = 9 Then
            For j = 0 To vsfMoney.Cols + 2
                arrTemp(j) = arrCell(1, j)
            Next
            For j = 0 To vsfMoney.Cols + 2
                arrCell(1, j) = arrCell(i, j)
            Next
            For j = 0 To vsfMoney.Cols + 2
                arrCell(i, j) = arrTemp(j)
            Next
            Exit For
        End If
    Next
    '������д���
    For i = 1 To vsfMoney.Rows - 1
        For j = 0 To vsfMoney.Cols - 1
            vsfMoney.TextMatrix(i, j) = arrCell(i, j)
        Next
        
        '��������
        vsfMoney.RowData(i) = Val(arrCell(i, vsfMoney.Cols))
        vsfMoney.Row = i: vsfMoney.Col = 0
        vsfMoney.CellFontBold = IIf(Val(arrCell(i, vsfMoney.Cols + 1)) = 1, True, False)
        vsfMoney.CellForeColor = Val(arrCell(i, vsfMoney.Cols + 2))
    Next
    vsfMoney.Row = lngRow: vsfMoney.Col = lngCol
    vsfMoney.Redraw = blnRedraw
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COLMoney.C1���)) = 0 And Val(.RowData(i)) = 1 Then
                .RowHidden(i) = True
            Else
                .RowHidden(i) = False
            End If
        Next i
        .Refresh
    End With
End Sub

Private Sub AdjustBalance()
'���ܣ�����������Ŀ�б�
    Dim strSql As String, i As Long
    Dim intDef As Integer, lngW As Long, blnTmp As Boolean
            
    mbln���ʽ��� = False
    mcur������� = 0
    mcur�����޶� = 0
    mcur����͸֧ = 0
    mstrȱʡ���� = ""
    mstrBalance = ""
    
    mrs���㷽ʽ.Filter = ""
    If mrsInfo.State = 1 Then
        If Not IsNull(mrsInfo!����) And mbytMCMode <> 1 Then
            If Not MCPAR.����ʹ�ø����ʻ� Then mrs���㷽ʽ.Filter = "����<>3"
        End If
    End If
    
    With vsfMoney
        blnTmp = .Redraw
        .Redraw = False
        .Rows = 2
        .TextMatrix(0, COLMoney.C0����) = "���㷽ʽ"
        .TextMatrix(0, COLMoney.C1���) = "���"
        .TextMatrix(0, COLMoney.C2����) = "�������"
        .TextMatrix(0, COLMoney.C3����) = "����"
        
        '���ÿ��ý��㷽ʽ
        If Not mrs���㷽ʽ.EOF Then
            .Rows = mrs���㷽ʽ.RecordCount + 1
            For i = 1 To mrs���㷽ʽ.RecordCount
                .TextMatrix(i, COLMoney.C0����) = mrs���㷽ʽ!����
                .TextMatrix(i, COLMoney.C3����) = mrs���㷽ʽ!����
                .Row = i: .Col = 0
                .CellForeColor = vbBlack
                'ȱʡ��ʽ������ʾ
                If mrs���㷽ʽ!ȱʡ = 1 Then
                    mstrȱʡ���� = mrs���㷽ʽ!����
                    .Row = i: .Col = 0
                    .CellFontBold = True
                    intDef = .Row
                ElseIf InStr(",3,4,", mrs���㷽ʽ!����) > 0 Then
                    .Row = i: .Col = 0
                    .CellForeColor = vbBlue
                ElseIf InStr(",9,", mrs���㷽ʽ!����) > 0 Then
                    .Row = i: .Col = 0
                    .CellForeColor = vbRed
                End If
                mrs���㷽ʽ.MoveNext
            Next
        End If
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .Row = 0
        .Col = 0: .CellAlignment = 4
        .Col = 1: .CellAlignment = 4
        .Col = 2: .CellAlignment = 4
        .Col = 3: .CellAlignment = 4
        
        lngW = .Width - 75
        If .Rows > .Height \ .RowHeight(0) Then lngW = lngW - 250
        .ColWidth(0) = lngW * 0.3
        .ColWidth(1) = lngW * 0.3
        .ColWidth(2) = lngW * 0.4
        .ColWidth(3) = 0
        
        .Col = 1
        For i = 1 To .Rows - 1
            .Row = i
            .CellBackColor = txtMoney.BackColor
            If InStr(",3,4,", Val(.TextMatrix(i, 3))) > 0 Then
                .RowData(i) = 1 'ҽ������ȱʡΪ���ɱ༭
            ElseIf Val(.TextMatrix(i, 3)) = 8 Then
                .RowData(i) = 1 '���ѿ����ɱ༭
            ElseIf Val(.TextMatrix(i, 3)) = 9 Then
                .RowData(i) = 1 '���Ѳ��ɱ༭
            Else
            
                .RowData(i) = 0 '��ͨ����ȱʡΪ���Ա༭
            End If
            .TextMatrix(i, 1) = "0.00"
            .TextMatrix(i, 2) = ""
        Next
        If intDef > 0 Then .Row = intDef
        
        txtOwe.Text = "0.00"
        
        .Redraw = blnTmp
    End With
End Sub

Private Sub ClearDetail(Optional blnSetPatiForeColor As Boolean = True)
    Dim i As Long, j As Long
    With mshDetail
        .Redraw = False
        .Clear
        .ClearStructure
        .Rows = 2: .Cols = 2
        .ColWidth(0) = 1000: .ColWidth(1) = 1000
        .Row = 1: .Col = 0
        .Redraw = True
    End With
    txt�ɿ�.Text = "0.00"
    txt�Ҳ�.Text = "0.00"
    txtTotal.Text = gstrDec
    txtTotal.Tag = gstrDec
    mstrDec = gstrDec
    mcurTotal = 0: mcur����� = 0
    If blnSetPatiForeColor Then txtPatient.ForeColor = Me.ForeColor
    With mshQuery
        .Tag = ""
        .Redraw = False
        .Clear
        .ClearStructure
        .Rows = 2
        .Cols = 2
        .Row = 1: .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub Form_Resize()
    Dim lngCancelW As Long
    Dim lngInsureH As Long
    
    On Error Resume Next
    If WindowState = 1 Then Exit Sub
    
    If chkCancel.Visible Or lblFlag.Visible Then lngCancelW = chkCancel.Width
    fraTitle.Width = Me.ScaleWidth - fraTitle.Left
    chkCancel.Left = fraTitle.Width - chkCancel.Width - 60
    lblFlag.Left = chkCancel.Left + (chkCancel.Width - lblFlag.Width) / 2
    
    cboNO.Left = fraTitle.Width - lngCancelW - 60 - cboNO.Width - 30
    lblNO.Left = cboNO.Left - lblNO.Width - 45
    txtInvoice.Left = lblNO.Left - txtInvoice.Width - 200
    lblFact.Left = txtInvoice.Left - lblFact.Width - 45
    
    fraPatient.Width = fraTitle.Width
    
    fraDate.Left = Me.ScaleWidth - fraDate.Width
    fraBalance.Left = fraDate.Left
    
    cmdCancel.Left = fraDate.Left + fraDate.Width - cmdCancel.Width - 50
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    
    tabCard.Width = Me.ScaleWidth - fraDate.Width - tabCard.Left - 30
    
    mshQuery.Width = tabCard.Width - mshQuery.Left - 60
    mshDetail.Width = tabCard.Width - mshDetail.Left - 60
    tabCard.Height = Me.ScaleHeight - tabCard.Top - fraAppend.Height - sta.Height - (fra��ע.Height - 50)
    With fra��ע
        .Width = tabCard.Width
        .Top = tabCard.Top + tabCard.Height - 50
        fraAppend.Top = .Top + .Height - 50
        txt��ע.Width = .Width - txt��ע.Left - .Left - 50
        fraBalance.Height = .Top + .Height - fraBalance.Top
    End With
    
    'fraAppend.Top = tabCard.Top + tabCard.Height
    mshDetail.Height = tabCard.Height - 480
    mshQuery.Height = tabCard.Height - 480
    
    'fraBalance.Height = tabCard.Top + tabCard.Height - fraBalance.Top
    
    cmdOK.Top = fraAppend.Top + (fraAppend.Height - cmdOK.Height) / 2
    cmdCancel.Top = cmdOK.Top
    cmd���㿨.Top = cmdOK.Top
    lngInsureH = IIf(lblҽ������.Visible, lblҽ������.Height + 30, 30)
    
    mshDeposit.Height = (fraBalance.Height - lblDeposit.Height - txtOwe.Height - 240) * 0.45
    lblҽ������.Top = mshDeposit.Top + mshDeposit.Height + 15
    lbl�����ʻ�.Top = lblҽ������.Top
    vsfMoney.Top = mshDeposit.Top + mshDeposit.Height + lngInsureH
    vsfMoney.Height = (fraBalance.Height - lblDeposit.Height - txtOwe.Height - 240) * 0.55 - lngInsureH
    
    txtOwe.Top = vsfMoney.Top + vsfMoney.Height + 15
    lblOwe.Top = txtOwe.Top + (txtOwe.Height - lblOwe.Height) / 2
    lblTicketCount.Top = lblOwe.Top
    
    fraAppend.Width = fra�Ҳ�.Width + lblTotal.Width + txtTotal.Width + 200
    fra�Ҳ�.Left = fraAppend.Width - fra�Ҳ�.Width
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mbytInState = 0 And mstrYBPati <> "" And mstrInNO = "" Then
        If MsgBox("��ǰ���ڶ�ҽ�����˽��ʣ�ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
        If YBIdentifyCancel = False Then        'ȡ��ҽ�����������֤,���ؼ�ʱ���˳�
            Cancel = 1: Exit Sub
        End If
    End If
    
    '�����ڲ���
    mlngPatientID = 0
    mbytInState = 0
    mblnViewCancel = False
    mstrInNO = ""
    mblnNOMoved = False
    mlngBillID = 0
    mstrPrivs = ""
    
    mstrȱʡ���� = "": mstrBalance = ""
    mstrYBPati = "":   mbytMCMode = 0:    mintInsure = 0
    mlng����ID = 0:    mcurTotal = 0:     mcur����� = 0
    mcur������� = 0:  mcur�����޶� = 0:  mcur����͸֧ = 0
    mbln����תסԺ = False: mstr��ҳId = "": mstrPepositDate = ""
    Call InitBalanceCondition
    Call InitPatiVariable
        
    Set mrsBalance = Nothing
    Set mrsDeposit = Nothing
    Set mrsInfo = New ADODB.Recordset
    
    Unload frmSetBalance
    
    If mbytInState = 0 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
        zl9LedVoice.Reset com
    End If
    
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    
    Call SaveWinState(Me, App.ProductName)
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    '����:47798
    If mbytInState = 0 Then
        Call SaveRegisterItem(g˽��ģ��, Me.Name, "idkind", IDKIND.IDKIND)
    End If

End Sub

Private Sub mshDeposit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    If Button = 2 Then
        If mstrInNO <> "" Or cboNO.Text <> "" Or chkCancel.Value = 1 Then  '�鿴������
            mnuViewToolCols(mnuViewToolCols.UBound).Visible = False
            mnuViewToolCols(mnuViewToolCols.UBound - 1).Caption = "���"
        Else
            mnuViewToolCols(mnuViewToolCols.UBound).Visible = True
            mnuViewToolCols(mnuViewToolCols.UBound - 1).Caption = "���"
        End If
                
        For i = 0 To mnuViewToolCols.UBound
            If mnuViewToolCols(i).Visible Then
                If i + 1 < mshDeposit.Cols Then mnuViewToolCols(i).Checked = mshDeposit.ColWidth(i + 1) <> 0
            End If
        Next
        If mbytFunc = 0 Then
            Me.PopupMenu Me.mnuPopu, 0
        Else
            Me.PopupMenu Me.mnuColsVisible, 0
        End If
    End If
End Sub

Private Sub mnuViewToolCols_Click(Index As Integer)
    Dim ArrHeader As Variant, i As Integer, j As Integer
        
    mnuViewToolCols(Index).Checked = Not mnuViewToolCols(Index).Checked
    
    For i = 0 To mnuViewToolCols.UBound
        If mnuViewToolCols(i).Visible And mnuViewToolCols(i).Checked Then j = j + 1
    Next
    If j < 2 Then
        sta.Panels(2).Text = "Ҫ�����ٱ���������ʾ!"
        mnuViewToolCols(Index).Checked = True
    End If
    
    If mnuViewToolCols(Index).Checked Then
        If mstrInNO <> "" Or cboNO.Text <> "" Or chkCancel.Value = 1 Then  '�鿴������
            ArrHeader = Split(mstrDepositRHeader, ",")
        Else
            ArrHeader = Split(mstrDepositHeader, ",")
        End If
        If Index + 1 < mshDeposit.Cols Then mshDeposit.ColWidth(Index + 1) = Split(ArrHeader(Index + 1), "|")(1)
    Else
        If Index + 1 < mshDeposit.Cols Then mshDeposit.ColWidth(Index + 1) = 0
    End If
End Sub

Private Sub mnuFileExcel_Click()
    Call OutputList(3)
End Sub

Private Sub mnuFilePreview_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFilePrint_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFilePrintSetup_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFileZero_Click()
    mnuFileZero.Checked = Not mnuFileZero.Checked
    Call LoadCardData
End Sub
Private Sub vsfMoney_DblClick()
    If Not txtMoney.Visible And vsfMoney.Row >= 1 And vsfMoney.Col > 0 _
        And mbytInState = 0 And chkCancel.Value = Unchecked _
        And mrsInfo.State = adStateOpen Then
                        
        '�����޸ĵĽ��㷽ʽ
        If vsfMoney.RowData(vsfMoney.Row) = 1 Then Exit Sub

        With txtMoney
            .MaxLength = IIf(vsfMoney.Col = 2, 30, 10)
            .Left = fraBalance.Left + vsfMoney.Left + vsfMoney.CellLeft + 15
            .Top = fraBalance.Top + vsfMoney.Top + vsfMoney.CellTop + (vsfMoney.CellHeight - txtMoney.Height) / 2 - 15
            .Width = vsfMoney.CellWidth - 60
            .ForeColor = vsfMoney.CellForeColor
            .BackColor = vsfMoney.CellBackColor
            .Alignment = IIf(vsfMoney.Col = 2, 0, 1)
            .Text = vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.Col)
            .SelStart = 0: .SelLength = Len(.Text)
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub vsfMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsfMoney
        If .Row >= 1 Then
            If .Col < .Cols - 2 Then
                .Col = .Col + 1
            Else
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    .Col = 1
                    If .Row - (.Height \ .RowHeight(0) - 2) > 1 Then
                        .TopRow = .Row - (.Height \ .RowHeight(1) - 2)
                    End If
                Else
                    If txt��ע.Visible And txt��ע.Enabled Then
                        txt��ע.SetFocus
                    ElseIf GetӦ�� > 0 And txt�ɿ�.Visible Then
                        txt�ɿ�.SetFocus
                    ElseIf cmdOK.Visible And cmdOK.Enabled Then
                        cmdOK.SetFocus
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfMoney_KeyPress(KeyAscii As Integer)
    If Not txtMoney.Visible And vsfMoney.Row >= 1 And vsfMoney.Col > 0 _
        And KeyAscii <> 13 And mbytInState = 0 And chkCancel.Value = Unchecked _
        And mrsInfo.State = adStateOpen Then
        
        '�����޸ĵĽ��㷽ʽ
        If vsfMoney.RowData(vsfMoney.Row) = 1 Then Exit Sub
        
        '�������û����
        If vsfMoney.Col = 1 Then If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
        With txtMoney
            .MaxLength = IIf(vsfMoney.Col = 2, 30, 10)
            .Left = fraBalance.Left + vsfMoney.Left + vsfMoney.CellLeft + 15
            .Top = fraBalance.Top + vsfMoney.Top + vsfMoney.CellTop + (vsfMoney.CellHeight - txtMoney.Height) / 2 - 15
            .Width = vsfMoney.CellWidth - 60
            .ForeColor = vsfMoney.CellForeColor
            .BackColor = vsfMoney.CellBackColor
            .Alignment = IIf(vsfMoney.Col = 2, 0, 1)
            .Text = Chr(KeyAscii)
            .SelStart = 1
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub mshDetail_DblClick()
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    If InStr(mstrPrivs, ";��������;") = 0 Then Exit Sub
    If mshDetail.Col <> GetColNum("���ʽ��") Then Exit Sub
     
    If Not txtMoney.Visible And mshDetail.Row >= 1 _
        And mbytInState = 0 And chkCancel.Value = Unchecked _
        And mrsInfo.State = adStateOpen Then
        If IsNull(mrsInfo!ҽ����) And mbytFunc <> 0 Then
            With txtMoney
                .Left = mshDetail.Left + mshDetail.CellLeft + 15
                .Top = mshDetail.Top + mshDetail.CellTop + (mshDetail.CellHeight - txtMoney.Height) / 2 - 15
                .Width = mshDetail.CellWidth - 60
                .ForeColor = mshDetail.CellForeColor
                .BackColor = mshDetail.CellBackColor
                .Alignment = 1
                .Text = mshDetail.TextMatrix(mshDetail.Row, mshDetail.Col)
                .SelStart = 0: .SelLength = Len(.Text)
                .ZOrder: .Visible = True
                .SetFocus
            End With
        End If
    End If
End Sub

Private Sub mshDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        If mshDetail.Row >= 1 Then
            If mshDetail.Col = GetColNum("���ʽ��") Then
                If mshDetail.Row < mshDetail.Rows - 1 Then
                    mshDetail.Row = mshDetail.Row + 1
                    If mshDetail.Row - (mshDetail.Height \ mshDetail.RowHeight(0) - 2) > 1 Then
                        mshDetail.TopRow = mshDetail.Row - (mshDetail.Height \ mshDetail.RowHeight(1) - 2)
                    End If
                Else
                    mshDeposit.SetFocus
                End If
            Else
                mshDetail.Col = mshDetail.Col + 1
            End If
        End If
    End If
End Sub

Private Sub mshDetail_KeyPress(KeyAscii As Integer)
    If InStr(mstrPrivs, ";��������;") = 0 Then Exit Sub
    If mshDetail.Col <> GetColNum("���ʽ��") Then Exit Sub
    
    If Not txtMoney.Visible And mshDetail.Row >= 1 _
        And KeyAscii <> 13 And mbytInState = 0 And chkCancel.Value = Unchecked _
        And mrsInfo.State = adStateOpen Then
        If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        If IsNull(mrsInfo!ҽ����) And mbytFunc <> 0 Then
            With txtMoney
                .Left = mshDetail.Left + mshDetail.CellLeft + 15
                .Top = mshDetail.Top + mshDetail.CellTop + (mshDetail.CellHeight - txtMoney.Height) / 2 - 15
                .Width = mshDetail.CellWidth - 60
                .ForeColor = mshDetail.CellForeColor
                .BackColor = mshDetail.CellBackColor
                .Alignment = 1
                .Text = Chr(KeyAscii)
                .SelStart = 1
                .ZOrder: .Visible = True
                .SetFocus
            End With
        End If
    End If
End Sub

Private Sub mshDetail_LeaveCell()
    txtMoney.Visible = False
End Sub

Private Sub mshDetail_Scroll()
    txtMoney.Visible = False
End Sub

Private Sub mshQuery_EnterCell()
    Dim i As Long, blnPre As Boolean
    Dim intRow As Integer, intCol As Integer
    
    blnPre = mshQuery.Redraw
    intRow = mshQuery.Row: intCol = mshQuery.Col
    mshQuery.Redraw = False
    
    For i = 0 To mshQuery.Cols - 1
        mshQuery.Col = i
        mshQuery.CellBackColor = mshQuery.BackColorSel
        mshQuery.CellForeColor = mshQuery.ForeColorSel
    Next
    
    mshQuery.Row = intRow:  mshQuery.Col = intCol
    mshQuery.Redraw = blnPre
End Sub

Private Sub mshQuery_LeaveCell()
    Dim i As Long, blnPre As Boolean
    
    blnPre = mshQuery.Redraw
    mshQuery.Redraw = False
    
    For i = 0 To mshQuery.Cols - 1
        mshQuery.Col = i
        mshQuery.CellBackColor = mshQuery.BackColor
        mshQuery.CellForeColor = mshQuery.ForeColor
    Next
    
    mshQuery.Redraw = blnPre
End Sub

Private Sub mshQuery_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mnuFileZero.Visible = InStr(",2,4,7,", tabCard.SelectedItem.Index) > 0
        mnuFile_1.Visible = InStr(",2,4,7,", tabCard.SelectedItem.Index) > 0
        PopupMenu mnuFile, 2
    End If
End Sub

Private Sub mshDetail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mnuFileZero.Visible = False
        mnuFile_1.Visible = False
        PopupMenu mnuFile, 2
    End If
End Sub

Private Sub opt��Ժ_Click()
    
    Call zlChangeDefaultTime
    If mshDetail.TextMatrix(1, 0) <> "" Then
        If Not IsNull(mrsInfo!����) And mbytMCMode <> 1 Then Call ShowBalance   'ҽ������Ԥ����
        Call ShowMoney(True, , mty_ModulePara.bytMzDeposit)
    End If
End Sub

Private Sub opt��;_Click()
    Call zlChangeDefaultTime
    If mshDetail.TextMatrix(1, 0) <> "" Then
        If Not IsNull(mrsInfo!����) And mbytMCMode <> 1 Then Call ShowBalance 'ҽ������Ԥ����
        Call ShowMoney(True, , mty_ModulePara.bytMzDeposit)
    End If
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "LocalParSet" Then
        frmSetExpence.mstrPrivs = mstrPrivs
        frmSetExpence.mbytInFun = 1
        frmSetExpence.Show 1, Me
    End If
End Sub

Private Sub tabCard_Click()
    If tabCard.SelectedItem.Index = 1 Then
        mshDetail.ZOrder
        txtMoney.ZOrder
        
        mshDetail.Visible = True
        mshQuery.Visible = False
        
        mshDetail.TopRow = 1
        mshDetail.Row = 1
        mshDetail.Col = GetColNum("���ʽ��") ' mshDetail.Cols - 1
        If mshDetail.Visible Then mshDetail.SetFocus
    Else
        mshQuery.ZOrder
        mshQuery.Visible = True
        
        mshDetail.Visible = False
        
        'û�ж�ȡ���嵥����ʱ��ȡ
        If (mshQuery.TextMatrix(1, 0) = "" And mshQuery.Rows = 2) _
            Or Val(mshQuery.Tag) <> tabCard.SelectedItem.Index Then
            Call LoadCardData
        End If
                
        If mshQuery.Visible And mshQuery.Enabled Then mshQuery.SetFocus
    End If
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0: txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then mshDeposit.SetFocus
End Sub

Private Sub txtInvoice_Change()
    lblFact.Tag = ""
End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or InStr("0123456789" & Chr(8), Chr(KeyAscii)) > 0) Then
        KeyAscii = 0
    ElseIf Len(txtInvoice.Text) = txtInvoice.MaxLength And KeyAscii <> 8 And txtInvoice.SelLength <> Len(txtInvoice) Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtInvoice_GotFocus()
    SelAll txtInvoice
End Sub


Private Sub txtMoney_LostFocus()
    txtMoney.Visible = False
End Sub

Private Sub cboNO_GotFocus()
    If Not cboNO.Locked Then cboNO.SelStart = 0: cboNO.SelLength = Len(cboNO.Text)
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim strOper As String, vDate As Date, bytFlag As Byte
    Dim lng����ID  As Long
    'ת���ɴ�д(���ֲ��ɴ���)
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    ElseIf cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 15)
        
        '�Ƿ���ת������ݱ���
        If zlDatabase.NOMoved("���˽��ʼ�¼", cboNO.Text, , , Me.Caption) Then
            If Not ReturnMovedExes(cboNO.Text, 7, Me.Caption) Then Exit Sub
            mblnNOMoved = False
        End If
    
        '����Ȩ��
        If Not ReadBillInfo(2, cboNO.Text, -1, strOper, vDate) Then
            cboNO.Text = "": If cboNO.Visible Then cboNO.SetFocus
            Exit Sub
        End If
        If Not BillOperCheck(7, strOper, vDate, "����") Then
            cboNO.Text = "": If cboNO.Visible Then cboNO.SetFocus
            Exit Sub
        End If
        'lng����ID:49084
        mintInsure = BalanceExistInsure(cboNO.Text, bytFlag, lng����ID)
        mbytMCMode = bytFlag
        If mintInsure <> 0 Then
            '���ս���Ȩ���ж�
            If InStr(mstrPrivs, ";���ս���;") = 0 Then
                MsgBox "��û��Ȩ�����ϱ��ղ��˵Ľ��ʵ��ݡ�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            MCPAR.�ֱҴ��� = gclsInsure.GetCapability(support�ֱҴ���, lng����ID, mintInsure)
            If mbytMCMode = 1 Then
                MCPAR.���ﲡ�˽������� = gclsInsure.GetCapability(support�����������, lng����ID, mintInsure)
            Else
                MCPAR.��Ժ���˽������� = gclsInsure.GetCapability(support��Ժ���˽�������, lng����ID, mintInsure)
            End If
            MCPAR.�������Ϻ��ӡ�ص� = gclsInsure.GetCapability(support�������Ϻ��ӡ�ص�, lng����ID, mintInsure)
        Else
            If InStr(mstrPrivs, ";��ͨ���˽���;") = 0 Then
                MsgBox "��û��Ȩ��������ͨ���˵Ľ��ʵ��ݡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        If CheckExistsGathering(cboNO.Text) Then
            MsgBox "�ý��ʵ��ݴ����ѽɿ��Ӧ�տ��¼�����˿����ִ�����ϡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If CheckBillBeforIN(cboNO.Text) Then
            If MsgBox("�ý��ʵ��Ǳ���סԺ֮ǰ�����ģ���ȷ��Ҫ���ϸõ�����?", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        '��ȡҪ���ϵĽ��ʵ�
        If Not ReadBalance(cboNO.Text) Then
            cboNO.Text = "": If cboNO.Visible Then cboNO.SetFocus
        Else
            If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
        End If
    Else
           If InStr(mstrPrivs, ";��ͨ���˽���;") = 0 Then
                MsgBox "��û��Ȩ�����ϷǱ��ղ��˵Ľ��ʵ��ݡ�", vbInformation, gstrSysName
                Exit Sub
           End If
    End If
End Sub

Private Function CheckOutBalance(strNo As String) As Boolean
'���ܣ����ָ���Ľ��ʵ���Ӧ�ķ����Ƿ�ȫ��������ʷ���
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1 From סԺ���ü�¼ A, ���˽��ʼ�¼ B" & vbNewLine & _
            "Where A.����id = B.ID And B.NO = [1] And A.�����־ = 2 And Rownum < 2"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    
    CheckOutBalance = rsTmp.RecordCount = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtMoney_Validate(Cancel As Boolean)
    If txtMoney.Visible Then Call txtMoney_KeyPress(13)
End Sub

Private Sub txtOwe_Change()
    If IsNumeric(txtOwe.Text) Then
        If CCur(txtOwe.Text) > 0 Then
            txtOwe.ForeColor = vbBlue
        ElseIf CCur(txtOwe.Text) < 0 Then
            txtOwe.ForeColor = vbRed
        Else
            txtOwe.ForeColor = vbBlack
        End If
    End If
End Sub

Private Sub txtPatiBegin_Change()
    If IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
        txt����.Text = CDate(txtPatiEnd.Text) - CDate(txtPatiBegin.Text)
        If Val(txt����.Text) = 0 Then txt����.Text = 1
    Else
        txt����.Text = ""
    End If
End Sub

Private Sub txtPatiBegin_GotFocus()
    SelAll txtPatiBegin
End Sub

Private Sub txtPatiBegin_Validate(Cancel As Boolean)
    If txtPatiBegin.Text <> "____-__-__" And Not IsDate(txtPatiBegin.Text) Then
        Cancel = True
   End If
End Sub

Private Sub txtPatiEnd_Change()
    If IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
        txt����.Text = CDate(txtPatiEnd.Text) - CDate(txtPatiBegin.Text)
        If Val(txt����.Text) = 0 Then txt����.Text = 1
    Else
        txt����.Text = ""
    End If
End Sub

Private Sub txtPatiEnd_GotFocus()
    SelAll txtPatiEnd
End Sub

Private Sub txtPatiEnd_Validate(Cancel As Boolean)
    If txtPatiEnd.Text <> "____-__-__" And Not IsDate(txtPatiEnd.Text) Then
        Cancel = True
   End If
End Sub

Private Sub txtPatient_Change()
    If Not Me.ActiveControl Is txtPatient Or txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    IDKIND.SetAutoReadCard (txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    SelAll txtPatient
    If txtPatient.Locked Then Exit Sub
    
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    IDKIND.SetAutoReadCard (txtPatient.Text = "")
End Sub

Private Sub LoadPatientInfo(ByVal objCard As Card, ByVal blnCard As Boolean, _
    Optional ByVal intInsure As Integer, _
    Optional ByVal lng��ҳID As Long)
    '����:��ȡ������Ϣ
    '       lng��ҳID=��ȡָ��סԺ�����Ĳ�����Ϣ
    Dim strTmp As String, i As Long, strSql As String
    Dim blnICCard As Boolean, curDue As Currency, blnIDCard As Boolean
        
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset

    txtPatient.ForeColor = Me.ForeColor
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ = True Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.���� Like "*���֤*" And objCard.ϵͳ = True Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    
    sta.Panels(2).Text = ""
    If Not GetPatient(objCard, Trim(txtPatient.Text), blnCard, lng��ҳID) Then
        If txtPatient.Text = "" Then MsgBox "û���ҵ��ò���,�������������Ƿ���ȷ��", vbInformation, gstrSysName
        txtPatient.PasswordChar = "": txtPatient.Text = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        mstr����סԺ���� = ""
        Call ReInitPatiInvoice
        Exit Sub
    Else
        Unload frmSetBalance
        mstr����סԺ���� = ""
        '���￨������
        If (objCard.���� Like "IC��*" Or objCard.���� Like "*���֤*") And objCard.ϵͳ = True And blnCard Then blnCard = False
        If Mid(gstrCardPass, 7, 1) = "1" And (blnCard Or ((blnICCard Or blnIDCard) And mstrPassWord <> "")) Then
            If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!����, mrsInfo!�Ա�, "" & mrsInfo!����) Then
                GoTo ExitHandle
            End If
        End If
        
        '����:27690
        If Val(Nvl(mrsInfo!����)) = 0 Then
                If InStr(1, mstrPrivs, ";��ͨ���˽���;") = 0 Then
                    MsgBox "��û��Ȩ�޶ԷǱ��ղ��˽��н��㡣", vbInformation, gstrSysName
                    GoTo ExitHandle
                End If
        End If
        
        'ҽ������ж�
        If Not IsNull(mrsInfo!����) Then
            If InStr(mstrPrivs, ";���ս���;") = 0 Then
                MsgBox "��û��Ȩ�޶Ա��ղ��˽��н��㡣", vbInformation, gstrSysName
                GoTo ExitHandle
            End If
            
            If mstrYBPati <> "" And intInsure <> mrsInfo!���� Then
                MsgBox "���˵Ǽǵ�������ҽ�������֤�����಻����", vbInformation, gstrSysName
                GoTo ExitHandle
            End If
            
            If mbytMCMode = 1 And Not IsNull(mrsInfo!��ǰ����id) Then
                MsgBox "��Ժ���˲��ܽ�������ҽ�������֤��", vbInformation, gstrSysName
                GoTo ExitHandle
            End If
            
            MCPAR.�ֱҴ��� = gclsInsure.GetCapability(support�ֱҴ���, mrsInfo!����ID, mrsInfo!����)
            If mbytMCMode = 1 Then
                MCPAR.����Ԥ���� = gclsInsure.GetCapability(support����Ԥ��, mrsInfo!����ID, mrsInfo!����)
                MCPAR.������봫����ϸ = gclsInsure.GetCapability(support������봫����ϸ, mrsInfo!����ID, mrsInfo!����)
                MCPAR.�������_�������� = gclsInsure.GetCapability(support�������_�������ú���ýӿ�, mrsInfo!����ID, mrsInfo!����)
            Else
                MCPAR.δ�����Ժ = gclsInsure.GetCapability(supportδ�����Ժ, mrsInfo!����ID, mrsInfo!����)
                MCPAR.����ʹ�ø����ʻ� = gclsInsure.GetCapability(support����ʹ�ø����ʻ�, mrsInfo!����ID, mrsInfo!����)
                MCPAR.��Ժ��������Ժ = gclsInsure.GetCapability(support��Ժ��������Ժ, mrsInfo!����ID, mrsInfo!����)
                MCPAR.��;������������ϴ����� = gclsInsure.GetCapability(support��;������������ϴ�����, mrsInfo!����ID, mrsInfo!����)
                MCPAR.�������ú���ýӿ� = gclsInsure.GetCapability(support����_�������ú���ýӿ�, mrsInfo!����ID, mrsInfo!����)
                MCPAR.�������_�������� = False
            End If
        ElseIf mstrYBPati <> "" Then
            MsgBox "���������֤�ɹ�,�����˵Ǽǵ�����Ϊ�գ�", vbInformation, gstrSysName
                GoTo ExitHandle
        End If
        
        '����:34763 ��鲡���Ƿ���ڱ�ע��Ϣ
        
        If zlCheckPatiIsMemo(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID))) = True Then
            Call zlCallPatiMemoWriteAndRead(Me, mlngModul, mstrPrivs, Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID)), mobjInPatient)
        End If
        
        If lng��ҳID = 0 Then
            If mbytMCMode <> 1 Then
                If mrsInfo!��ҳID <> 0 Then
                    '����:30027:����ȱʡ����;����
                    '       1.��Ժ����,Ĭ��Ϊ��Ժ���� ����:û��"��;����"Ȩ�޵�,ҲĬ��Ϊ��Ժ����
                    '       2.��Ժ����-��ͨ����(�����ϴγ�Ժ���˵�ѡ���Ϊ׼)
                    '              Ĭ�ϳ�Ժ��(���ϴ�ѡ�����;���ʻ�סԺ����)����Ϊtrue,Ĭ��Ϊ��Ժ����,����Ĭ��Ϊ��;����
                    '       3.��Ժ����-ҽ������(������)
                    '           ����ҽ����߲���ȷ��,���,����ԭ���Ĺ���һ��,�������ϴγ�Ժ���˵�ѡ���Ϊ׼!
                    If InStr(mstrPrivs, ";��;����;") = 0 Then
                        opt��Ժ.Value = True: opt��;.Enabled = False
                    ElseIf Not IsNull(mrsInfo!��ǰ����id) And Nvl(mrsInfo!״̬, 0) <> 3 Then  '��Ժ����()
                            If IsNull(mrsInfo!����) Then
                                'ҽ��������Ҫ֧����;����ʱֻ�������ϴ�����,���Բ���
                                If zlDatabase.GetPara("Ĭ�ϳ�Ժ����", glngSys, mlngModul, "1") <> "0" Then
                                    opt��Ժ.Value = True
                                Else
                                    opt��;.Value = True
                                End If
                            End If
                    Else
                            '��Ժ����(����Ԥ��Ժ�Ĳ���)
                             opt��Ժ.Value = True
                    End If
                    opt��Ժ.Enabled = True
                    
                    '��Ժ���˲������Ժ����(Ԥ��Ժ���˿���)
                    If gbln��Ժ��׼���� And Not IsNull(mrsInfo!��ǰ����id) Then         'And Nvl(mrsInfo!״̬, 0) <> 3:30572:Ԥ��ԺҲ����Ժ.
                        If Not opt��;.Enabled Then
                            MsgBox "��Ժ���˲������Ժ����,������û����;���ʵ�Ȩ��,���Բ��ܶԸò��˽���!", vbInformation, gstrSysName
                            GoTo ExitHandle
                        End If
                        If mblnFirst And mlngPatientID <> 0 Then
                            '��һ���Զ���ȡ���˽���ʱ,��ȥ��������
                            '38537:�������Ժ����,�϶���Ҫ����Ϊ��;����
                            opt��;.Value = True: opt��Ժ.Value = False: opt��Ժ.Enabled = False
                        Else
                            If opt��;.Value Then
                                opt��Ժ.Value = False: opt��Ժ.Enabled = False
                            Else
                                If MsgBox("��ǰ������Ժ���������Ժ���ʡ�" & vbCrLf & "����ǳ�Ժ���ʣ����Ƚ����˳�Ժ��" & _
                                    vbCrLf & "��Ҫ�Ըò��˽�����;������?", vbYesNo + vbInformation + vbDefaultButton2, gstrSysName) = vbYes Then
                                    opt��Ժ.Value = False: opt��Ժ.Enabled = False
                                    opt��;.Value = True
                                Else
                                    GoTo ExitHandle
                                End If
                            End If
                        End If
                    End If
                Else
                    '����:47430
                    opt��Ժ.Value = True: opt��Ժ.Enabled = False
                    opt��;.Enabled = False
                End If
            End If
            
            
            '����������
            strTmp = inBlackList(mrsInfo!����ID)
            If strTmp <> "" Then
                If MsgBox("����""" & mrsInfo!���� & """�����ⲡ�������С�" & vbCrLf & vbCrLf & "ԭ��" & vbCrLf & vbCrLf & "����" & strTmp & vbCrLf & vbCrLf & "Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    GoTo ExitHandle
                End If
            End If
                                                                                        
            'gbytAuditing:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
            '����:37369:��;���ʲ����
            If gbytAuditing <> 0 Then
                If HaveNOAuditing(mrsInfo!����ID) Then
                    If gbytAuditing = 1 Then
                        If MsgBox("�ò��˻�����δ��˵ļ��ʷ��ã�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            GoTo ExitHandle
                        End If
                    ElseIf gbytAuditing = 2 Then
                         If MsgBox("�ò��˻�����δ��˵ļ��ʷ��ã�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                GoTo ExitHandle
                         End If
                          If opt��;.Enabled Then opt��;.Value = True
                    End If
                End If
            End If
            
            '�Զ����㲡�˵Ĵ�λ���úͻ�������
            If mrsInfo!��ҳID <> 0 And mbytMCMode <> 1 Then
                strSql = "ZL1_AUTOCPTPATI(" & mrsInfo!����ID & "," & mrsInfo!��ҳID & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
            Call InitԤ�����
            '��ȡ���˷������
            If mintԤ����� = 0 Then
                strSql = "Select Sum(Ԥ�����) As Ԥ�����,Sum(�������) As ������� From ������� Where ����ID= [1] And ����=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsInfo!����ID)))
            Else
                strSql = "Select Ԥ�����,������� From ������� Where ����ID= [1] And ����=1 And ����= [2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsInfo!����ID)), mintԤ�����)
            End If
            mcurSpare = Get�������(mrsInfo!����ID, 0, mintԤ�����)
            lblSpare.Tag = Get�������(mrsInfo!����ID, 1, mintԤ�����)  'ShowBalance��LED��ʾ���õ��˽��
            lblSpare.Caption = "Ԥ�����:" & Format(lblSpare.Tag, "0.00")
            '60615,������,2013-12-20,״̬����ʾԤ�������ý���ʣ�����
            If rsTmp.RecordCount <> 0 Then
                sta.Panels(3).Text = "Ԥ��:" & Format(Nvl(rsTmp!Ԥ�����), "0.00") & _
                                     "/����:" & Format(Nvl(rsTmp!�������), "0.00") & _
                                     "/ʣ��:" & Format(Val(Nvl(rsTmp!Ԥ�����)) - Val(Nvl(rsTmp!�������)), "0.00")
            End If
            
            If InStr(mstrPrivs, ";Ӧ�տ����;") > 0 Then
                curDue = GetPatientDue(Val(mrsInfo!����ID))
                If curDue <> 0 Then
                    MsgBox mrsInfo!���� & ",Ӧ�տ����:" & Format(curDue, "0.00") & "Ԫ", vbInformation, gstrSysName
                    sta.Panels(2).Text = "����Ӧ�տ����:" & Format(curDue, "0.00") & "Ԫ"
                End If
            End If
            
            mblnDateMoved = zlDatabase.DateMoved(mrsInfo!�Ǽ�ʱ��, , , Me.Caption)
        Else
            If IsNull(mrsInfo!��ǰ����id) And Nvl(mrsInfo!״̬, 0) <> 3 Then
                opt��Ժ.Value = True: opt��Ժ.Visible = True: opt��Ժ.Enabled = True
            End If
        End If
        txtPatient.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        
        txtPatient.IMEMode = 0
        txtPatient.Text = mrsInfo!����: txtSex.Text = Nvl(mrsInfo!�Ա�): txtOld.Text = Nvl(mrsInfo!����)
        '��ʾ��������
        '62906
        '�Һ�ʱ,����δ����ҽ����֤ʱ,�����������벡�˺�,������֤ҽ��
        cmdYB.Enabled = IIf(mbytFunc = 0, True, False)
        If Not IsNull(mrsInfo!����) Then
            sta.Panels(2).Text = sta.Panels(2).Text & "  ���ࣺ" & GetInsureName(mrsInfo!����)
            If mbytMCMode = 1 Then Call InitBalanceSet(False)
            cmdOK.Enabled = False
        End If
        txtPatient.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!��������))
        
        lbl״̬.Caption = GetPatiState(mrsInfo!����ID)
        lbl���ʽ.Left = lbl״̬.Left + lbl״̬.Width + 100
        lbl���ʽ.Caption = "" & mrsInfo!ҽ�Ƹ��ʽ
        pic״̬.Width = lbl״̬.Width + lbl���ʽ.Width + 300
        pic״̬.Visible = True
        
        txt�ѱ�.Text = Nvl(mrsInfo!�ѱ�)
        
        '����65105,������:�������ʱ��ʾ�����
        If mbytFunc = 1 Then
            If Not IsNull(mrsInfo!סԺ��) Then
                txt��ʶ��.Text = mrsInfo!סԺ��
                lbl��ʶ��.Visible = True: txt��ʶ��.Visible = True
                lbl��ʶ��.Caption = "סԺ��"
            End If
            If Not IsNull(mrsInfo!��ǰ����) Then
                txtBed.Text = "" & mrsInfo!��ǰ����
                txt����.Text = mrsInfo!��ǰ����
                lblBed.Visible = True: txtBed.Visible = True
                lbl����.Visible = True: txt����.Visible = True
            ElseIf Not IsNull(mrsInfo!��Ժ����) Then
                txtBed.Text = Nvl(mrsInfo!��Ժ����)
                txt����.Text = mrsInfo!��Ժ����
                lblBed.Visible = True: txtBed.Visible = True
                lbl����.Visible = True: txt����.Visible = True
            End If
        ElseIf mbytFunc = 0 Then
            If Not IsNull(mrsInfo!�����) Then
                txt��ʶ��.Text = mrsInfo!�����
                lbl��ʶ��.Visible = True: txt��ʶ��.Visible = True
                lbl��ʶ��.Caption = "�����"
            End If
        End If
        
        '��ʾ����Ҫ��������,����ʼ��������
        '-------------------------------------------------------------------------------------------
        If lng��ҳID = 0 Then
            strTmp = ""
            If Not ShowBalance(True, strTmp) Then
                MsgBox strTmp, vbInformation, gstrSysName
                GoTo ExitHandle
            End If
                    
            Call Led��ӭ��Ϣ
        End If
        
        If vsfMoney.Visible And vsfMoney.Enabled Then vsfMoney.SetFocus
    End If
    
    Call ReInitPatiInvoice
    Call Calc�Ҳ�
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
ExitHandle:
    mcurSpare = 0
    Call NewBill
    If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
    Exit Sub
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKIND.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, strInput As String
    If txtPatient.Locked Then Exit Sub
    '����ѡ����
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        With frmPatiSelect
            .mstrPrivs = mstrPrivs
            .mbytUseType = 3
            Set .mfrmParent = Me
            .Show 1, Me
            mintPatientRange = Val(zlDatabase.GetPara("��ʾ���岡��", glngSys, mlngModul, 0))
        End With
    Else
        If IDKIND.GetCurCard.���� Like "����*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKIND.ShowPassText)
        ElseIf IDKIND.GetCurCard.���� = "�����" Or IDKIND.GetCurCard.���� = "סԺ��" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
            End If
        Else
            txtPatient.PasswordChar = IIf(IDKIND.ShowPassText, "*", "")
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
        End If
    End If
    Me.Refresh
    
    'ˢ����ϻ���������س�
    If blnCard And Len(txtPatient.Text) = IDKIND.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        strInput = txtPatient.Text
        Call FindPati(IDKIND.GetCurCard, blnCard, strInput)
    End If
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call NewBill
    txtPatient.Text = strInput
    '���˺�:27503
    If mty_ModulePara.bln���ʺ�����Ϣ Then
        If txtInvoice.Tag <> "" And txtInvoice.Text <> txtInvoice.Tag Then txtInvoice.Text = txtInvoice.Tag '��Ҫ��Ҫ������Ϣ,��ȷ������Ҫ�����̶�
    End If
    If mblnFirst Then mstrTime = mstr��ҳId
    If mblnOneCard And Not mobjICCard Is Nothing And objCard.���� Like "IC��*" And objCard.ϵͳ Then
        Call SetOneCardBalance  '��ʾһ��ͨ���
    End If
    Call LoadPatientInfo(objCard, blnCard)
End Sub

Private Sub vsfMoney_Scroll()
    txtMoney.Visible = False
End Sub
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional ByVal lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=�Ƿ���￨ˢ��,lng��ҳID=��ȡָ��סԺ�����Ĳ�����Ϣ
    '����:
    '����:�Ƿ��ȡ�ɹ�,�ɹ�ʱmrsInfo�а���������Ϣ,ʧ��ʱmrsInfo=Close,strInput�����������ж��Ƿ�����ʾ��,�����ٴ���ʾû���ҵ�����
    '����:���˺�
    '����:2011-08-03 16:56:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strWhere As String, strField As String, bytMzMode As Byte
    Dim strSql As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim strPati As String, strRange As String
    Dim vRect As RECT
    mstrPassWord = ""
    mlngCardTypeID = 0
    On Error GoTo errH
    strField = ",A.��ǰ����ID"
    bytMzMode = mbytMCMode
    
    If mlngPatientID <> 0 And mblnFirst Then
        '��һ��ȡ��ʱ
        lng��ҳID = Val(mstr��ҳId)
        If Val(mstr��ҳId) = 0 Then '����
            strWhere = strWhere & " And B.��ҳID(+)=-100"
            bytMzMode = IIf(bytMzMode = 0, 0, 1): strField = " ,NULL as ��ǰ����ID"
            If mbytFunc = 1 Then bytMzMode = 2  'סԺ��:44022
        Else    'ָ������
            strWhere = strWhere & "  And B.��ҳID=[3]"
            bytMzMode = 2   'סԺ��
        End If
    Else
        If mbytFunc = 0 Then    '����
            strWhere = strWhere & " And   A.��ҳID=B.��ҳID(+)"
            '����:43730
            bytMzMode = IIf(bytMzMode = 0, 0, 1): strField = " ,NULL as ��ǰ����ID"
        Else
            'ָ������
            '76451,Ƚ����,2014-8-19
            If lng��ҳID <> 0 Then strField = ",Decode(A.��ҳID,[3],A.��ǰ����ID,NULL) as ��ǰ����ID"
            strWhere = IIf(lng��ҳID = 0, " And A.��ҳID=B.��ҳID(+)", " And B.��ҳID=[3]")
            bytMzMode = 2
        End If
    End If
    strSql = _
        "Select A.����ID,Nvl(B.��ҳID,0) as ��ҳID,A.�����,nvl(B.סԺ��,A.סԺ��) as סԺ��,A.��ǰ����,B.��Ժ����," & _
        "       nvl(B.����,A.����) as ����, nvl(B.�Ա�,Nvl(A.�Ա�,'δ֪')) as  �Ա�,Nvl(B.����,A.����) as ����,A.IC����,A.���￨��,A.����֤��," & _
        "       Nvl(B.�ѱ�,A.�ѱ�) as �ѱ�,C.���� as ��ǰ����" & strField & ",D.���� as ��Ժ����,B.��Ժ����ID," & _
                IIf(bytMzMode = 0, "NULL", IIf(bytMzMode = 1, "A.����", "B.����")) & " as ����,E.����,E.ҽ����,E.����," & _
        " A.�Ǽ�ʱ��,Nvl(B.״̬,0) as ״̬,Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,Nvl(B.��˱�־,0) as ��˱�־,B.��Ժ����,B.��Ժ����,B.��������,B.��������" & _
        " From ������Ϣ A,������ҳ B,���ű� C,���ű� D,ҽ�����˵��� E,ҽ�����˹����� F" & _
        " Where A.ͣ��ʱ�� is NULL And A.����ID=B.����ID(+)   " & strWhere & _
        " And A.����ID=F.����ID(+) And F.��־(+)=1 And F.ҽ����=E.ҽ����(+) And F.����=E.����(+) And F.���� = E.����(+)" & _
        " And A.��ǰ����ID=C.ID(+) And B.��Ժ����ID=D.ID(+)"
        
    If blnCard = True And objCard.���� Like "����*" Then    'ˢ��
        If IDKIND.Cards.��ȱʡ������ And Not IDKIND.GetfaultCard Is Nothing Then
            lng�����ID = IDKIND.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        mlngCardTypeID = lng�����ID
        strSql = strSql & " And A.����ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSql = strSql & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSql = strSql & " And A.�����=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strSql = strSql & " And A.����ID=(Select nvl(Max(����ID),0) as ����ID From ������ҳ   Where  סԺ��=[2])"
        strInput = Mid(strInput, 2)
    Else '��������
        mlngCardTypeID = objCard.�ӿ����
        Select Case objCard.����
            Case "����", "��������￨"
                If mrsInfo.State = 1 Then
                    If mrsInfo!���� = Trim(txtPatient.Text) Then
                        GetPatient = True
                        Exit Function
                    End If
                End If
                
                If mintPatientRange > 0 Then
                    Select Case mintPatientRange
                        Case 1  '�κη���δ���岡��
                            strRange = ""
                        Case 2  '���δ����Ĳ���
                            strRange = " And C.��Դ;�� = 4"
                        Case 3  'סԺδ����Ĳ���
                            strRange = " And C.��Դ;�� = 2"
                        Case 4  '����δ����Ĳ���
                            strRange = " And C.��Դ;�� = 1"
                    End Select
                    strPati = " And Exists(Select 1 From ����δ����� C Where C.����id=A.����ID And Nvl(C.��ҳID,0)=A.��ҳID" & strRange & ")"
                End If
                
                 'ͨ����������
                strPati = "" & _
                " Select A.����ID as ID,A.����ID,A.סԺ��, A.�����, nvl(B.�Ա�,Nvl(A.�Ա�,'δ֪')) as  �Ա�, A.����, A.סԺ����, A.��ͥ��ַ, A.������λ," & vbNewLine & _
                "   To_Char(A.��������,'YYYY-MM-DD') as ��������,  To_Char(B.��Ժ����,'YYYY-MM-DD') as ��Ժ����, To_Char(B.��Ժ����,'YYYY-MM-DD') as ��Ժ����" & vbNewLine & _
                " From ������Ϣ A, ������ҳ B" & vbNewLine & _
                " Where A.����id = B.����id(+) And A.��ҳID = B.��ҳid(+) And A.ͣ��ʱ�� Is Null And A.���� = [1] " & vbNewLine & strPati & vbNewLine & _
                " Order By Decode(סԺ��, Null, 1, 0), ��Ժ���� Desc"
                        
                vRect = GetControlRect(txtPatient.hWnd)
                Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput)
                            
                If Not mrsInfo Is Nothing Then
                    strInput = Val(mrsInfo!����ID)
                    strSql = strSql & " And A.����ID=[2]"
                Else
                    Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                strSql = strSql & " And A.ҽ����=[2]"
            Case "���֤��", "�������֤", "���֤"
                If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strInput = "-" & lng����ID
                blnHavePassWord = True
                strSql = strSql & " And A.����ID=[1] "
            Case "IC����"
                If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strInput = "-" & lng����ID
                blnHavePassWord = True
                strSql = strSql & " And A.����ID=[1] "
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.�����=[2]"
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.סԺ��=[2]"
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strSql = strSql & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Mid(strInput, 2)), strInput, lng��ҳID)
    If mrsInfo.EOF Then Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
    
    mstrPassWord = strPassWord
    If Not blnHavePassWord Then
        mstrPassWord = Nvl(mrsInfo!����֤��)
    End If
    
    '����������:�����������ʾ
    '34681:35686
    If zlCheckPatiIsDeath(Val(Nvl(mrsInfo!����ID))) = True Then
        If MsgBox("ע��:" & vbCrLf & "    �ò����Ѿ�����,�Ƿ��������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
        End If
    End If
    
    '��Ҫ�ٴμ��,�Է������ڼ�����˵Ĳ��˱�ȡ�����
    '36209
    If (InStr(mstrPrivs, ";δ��˲�����;����;") = 0 And opt��;.Value Or InStr(mstrPrivs, ";δ��˲��˳�Ժ����;") = 0 And opt��Ժ.Value) And Val(Nvl(mrsInfo!��ҳID)) <> 0 Then
        If Not Chk�������(mrsInfo!����ID, Val(Nvl(mrsInfo!��ҳID))) Then
            If MsgBox("�����ʷ����а������˵�" & Val(Nvl(mrsInfo!��ҳID)) & "��סԺδ��˵ķ��ü�¼��" & vbCrLf & _
                " �Ƿ��������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
            End If
        End If
    End If
    GetPatient = True
    Exit Function
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function

Private Function ShowBillFormat()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰ��¼���շѲ���Ա��ʾ����ʹ���շ�Ʊ�ݸ�ʽ
    '����:���˺�
    '����:2011-01-02 09:47:25
    '����:35142
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intFormat As Integer, strRptName As String
    Dim blnҽ������ As Boolean
    
    lblFormat.Caption = "": blnҽ������ = False
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then blnҽ������ = Not IsNull(mrsInfo!����)
    End If
    
    'gbytInvoiceKind:����Ʊ������,0-סԺƱ��;1-����Ʊ��
    strRptName = IIf(gbytInvoiceKind = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
    intFormat = mintInvoiceFormat
    If intFormat = 0 Then   '��ȱʡƱ�ݸ�ʽ��ʾ
        intFormat = Val(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\zl9Report\LocalSet\" & strRptName, "Format", 1))
    End If
    
    strSql = "Select B.˵�� From zlReports A,zlRptFmts B" & _
        " Where A.ID=B.����ID And A.���=[1] And B.���=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strRptName, intFormat)
    If Not rsTmp.EOF Then
        lblFormat.Caption = "Ʊ��:" & Nvl(rsTmp!˵��)
        lblFormat.Visible = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowBalance(Optional ByVal blnFirst As Boolean, Optional ByRef strMessage As String) As Boolean
'���ܣ���������,��ʾ����Ҫ��������,����ʼ��������
'������blnFirst-�������ȷ��ʱ���ã�strMessage-������ʾ��Ϣ
'˵�����ù��ܿ�������һ�����˽�����ɺ����,Ҳ�����ǵ�һ�������ڽ���ʱ��һ������;����
    Dim i As Long, j As Long, curͳ��֧�� As Currency, cur�����ʻ� As Currency, curTmp As Currency, lngMaxLength As Long, lngP As Long
    Dim rsDetail As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim strMoney As String, strInfo As String, strTime As String
    Dim blnUpload As Boolean, blnZero As Boolean, blnAll As Boolean
    Dim dBegin As Date, dEnd As Date, DatTmp As Date
    Dim dblMoney As Double, strסԺ���� As String
    Dim strSql As String
    
    Call ClearDetail(False)
    Call AdjustBalance
    Call AdjustDeposit
    
    If mrsInfo.State <> 1 Then Exit Function
    Screen.MousePointer = 11
    Me.Refresh
    
    blnZero = gblnZero
    
    If Not IsNull(mrsInfo!����) And mbytMCMode <> 1 Then
        If opt��;.Value And MCPAR.��;������������ϴ����� Then blnUpload = True
    End If
    
    If IsNull(mrsInfo!����) Then
        mblnNoInsure = False
        picOwnFee.Visible = False
    End If
    If Not IsNull(mrsInfo!����) Then
        If blnFirst Then
            mstrChargeType = zlDatabase.GetPara("ҽ������ǰ�Ƚ��Էѷ���", glngSys, mlngModul, "")
            If mstrChargeType <> "" Then
                mblnNoInsure = True
                picOwnFee.Visible = True
                picOwnFee.Left = lblTitle.Left + lblTitle.Width + 150
                lblOwnFee.Caption = ""
                strSql = "Select ��� From �շ���� Where ���� In (Select Column_Value From Table(f_Str2list([1])))"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstrChargeType)
                Do While Not rsTmp.EOF
                    lblOwnFee.Caption = lblOwnFee.Caption & "," & rsTmp!��� & "��"
                    rsTmp.MoveNext
                Loop
                If lblOwnFee.Caption <> "" Then
                    lblOwnFee.Caption = Mid(lblOwnFee.Caption, 2)
                    picOwnFee.Width = lblOwnFee.Width + 150
                End If
                mstrChargeType = "'" & Replace(mstrChargeType, ",", "','") & "'"
            End If
        End If
    End If
    
    Set mrsBalance = GetBalance(mbytFunc, mstrPrivs, mrsInfo!����ID, IIf(mbytFunc = 0, "0", mstrTime), mstrUnit, mstrClass, mDateBegin, mDateEnd, mbytBaby, mstrItem, blnUpload, blnZero, mblnDateMoved, mbytMCMode = 1, mbytKind, mtySquareCard.blnExistsObjects, mstrChargeType)
    If mrsBalance Is Nothing Then Screen.MousePointer = 0: Exit Function
    If mrsBalance.RecordCount = 0 And mblnNoInsure = True Then
        mblnNoInsure = False
        picOwnFee.Visible = False
        mstrChargeType = ""
        Set mrsBalance = GetBalance(mbytFunc, mstrPrivs, mrsInfo!����ID, IIf(mbytFunc = 0, "0", mstrTime), mstrUnit, mstrClass, mDateBegin, mDateEnd, mbytBaby, mstrItem, blnUpload, blnZero, mblnDateMoved, mbytMCMode = 1, mbytKind, mtySquareCard.blnExistsObjects, mstrChargeType)
        If mrsBalance Is Nothing Then Screen.MousePointer = 0: Exit Function
    End If
    
    If blnFirst And mrsBalance.RecordCount = 0 And mbytFunc = 0 Then
        mbytKind = 1 'ȱʡֻȡ��ͨ���ã����û���ټ��ֻ���������������
        Set mrsBalance = GetBalance(mbytFunc, mstrPrivs, mrsInfo!����ID, IIf(mbytFunc = 0, "0", mstrTime), mstrUnit, mstrClass, mDateBegin, mDateEnd, mbytBaby, mstrItem, blnUpload, blnZero, mblnDateMoved, mbytMCMode = 1, mbytKind, mtySquareCard.blnExistsObjects, mstrChargeType)
        If mrsBalance Is Nothing Then
            Screen.MousePointer = 0: Exit Function
        ElseIf mrsBalance.RecordCount > 0 Then
            If MsgBox("�ò�����ͨ�����ѽ���,Ҫ�������ý��н�����?", vbInformation, Me.Caption) = vbNo Then
                Set mrsBalance = Nothing
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    
    If mrsBalance.RecordCount = 0 Then
        If blnFirst Then strMessage = "�ò���û����Ҫ���ʵķ��ã�"
        Screen.MousePointer = 0: Exit Function
    End If
    
    If blnFirst Then
        Call GetStateIF
        If InStr(mstrPrivs, ";δ��˲�����;����;") = 0 And InStr(mstrPrivs, ";δ��˲��˳�Ժ����;") = 0 And mrsInfo!��ҳID <> 0 Then
            If CStr(mrsInfo!��ҳID) = mstrAllTime Then
                If mrsInfo!��˱�־ = 0 And mrsInfo!��ҳID <> 0 Then
                    strMessage = "��ǰ����δ��ˣ��㲻�ܶ�δ��˵Ĳ��˽��н��ʡ�"
                    Screen.MousePointer = 0: Exit Function
                End If
            Else
                blnAll = True
                For i = 0 To UBound(Split(mstrAllTime, ","))
                    strTime = Split(mstrAllTime, ",")(i)
                    If Val(strTime) <> 0 Then
                        If Not Chk�������(mrsInfo!����ID, Val(strTime)) Then
                            mstrUnAuditTime = IIf(mstrUnAuditTime = "", strTime, mstrUnAuditTime & "," & strTime)
                        Else
                            blnAll = False
                        End If
                    Else
                        blnAll = False
                    End If
                Next
                If blnAll Then
                    strMessage = "�ò�������סԺ���ö�û����ˣ����ܽ��н��ʣ�"
                    Screen.MousePointer = 0: Exit Function
                End If
            End If
        End If
        If cmdPar.Enabled Then
            If (gbln���סԺ������������ And InStr(1, mstrAllTime, ",") > 0 Or Not IsNull(mrsInfo!����) And MCPAR.�������ú���ýӿ�) Or MCPAR.�������_�������� Then
                '---------------------------------------------------------------------------------------
                '34260:��Ѫ�Ѽ��
                If gbyt����ʱ��Ѫ�Ѽ�� = 1 And InStr(1, "," & mstrALLChargeType & ",", ",'K',") > 0 Then     '0:�����;1-��鲢��ʾ
                    Call MsgBox("ע��:" & vbCrLf & "    �ò���δ������а�������Ѫ��,��ע�����Ѫ�ѽ��н���!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName)
                End If
                Screen.MousePointer = 0
                mblnNOCancel = True
                Call cmdPar_Click
                mblnNOCancel = False
                ShowBalance = True  '���������������û�д�����ã��Է��سɹ��������ٴ�ѡ��
                Exit Function
            End If
        End If
        '---------------------------------------------------------------------------------------
        '34260:��Ѫ�Ѽ��
        If gbyt����ʱ��Ѫ�Ѽ�� = 1 Then '0:�����;1-��鲢��ʾ
            If InStr(1, "," & mstrALLChargeType & ",", ",'K',") > 0 Then  '34260
                If MsgBox("ע��:" & vbCrLf & "    �ò���δ������а�������Ѫ��,�����Ƿ�ֻ����Ѫ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mstrChargeType = "'K'"
                     If ShowBalance(False) Then
                        ShowBalance = True
                     End If
                    Exit Function
                End If
            End If
        End If
        '---------------------------------------------------------------------------------------
    End If
    '78317:ҽ������Ĭ��ֻ��ȡ���һ��סԺ������
    If Val(Nvl(mrsInfo!����)) <> 0 And mstrTime = "" Then
        mstrTime = Split(mstrAllTime & ",", ",")(0)
        Set mrsBalance = GetBalance(mbytFunc, mstrPrivs, mrsInfo!����ID, IIf(mbytFunc = 0, "0", mstrTime), mstrUnit, mstrClass, mDateBegin, mDateEnd, mbytBaby, mstrItem, blnUpload, blnZero, mblnDateMoved, mbytMCMode = 1, mbytKind, mtySquareCard.blnExistsObjects, mstrChargeType)
    End If
    
    '����ʾ������ϸ
    '��־,סԺ,����,ʱ��,[���ݺ�],��Ŀ,��Ŀ,Ӥ����,[ID],[���],[��¼����],[��¼״̬],[ִ��״̬],[��ҳID],[��������ID],[�Ǽ�ʱ��],δ����,���ʽ��
    
    With mshDetail
        .Redraw = False
        Set .DataSource = mrsBalance
        .Cols = 18 '  .Cols - 1 '����ʾ��������
        .ToolTipText = "��" & mrsBalance.RecordCount & "����ϸ��¼!"
        
        '������ϸ��ʽ
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4
            Select Case .TextMatrix(0, i)
                Case "סԺ", "Ӥ����", "���ݺ�"
                    .ColAlignment(i) = 4
                Case "����", "��Ŀ", "��Ŀ", "ʱ��"
                    .ColAlignment(i) = 1
                Case "δ����", "���ʽ��"
                    .ColAlignment(i) = 7
            End Select
            Select Case .TextMatrix(0, i)
                Case "ID", "��־", "���", "��¼����", "��ҳID", "��������ID", "��¼״̬", "ִ��״̬", "����", "סԺ", "�Ǽ�ʱ��", _
                     "�ѱ�", "ִ�в���ID", "�շ����", "������", "����", "�۸�", "ͳ����", "���մ���ID", "�շ�ϸĿID", "���㵥λ"
                    .ColWidth(i) = 0
                Case "Ӥ����"
                    .ColWidth(i) = 520
                    .TextMatrix(0, i) = "Ӥ��"
                Case "��Ŀ"
                    .ColWidth(i) = 800
                Case "���ݺ�"
                    .ColWidth(i) = 950
                Case "δ����", "���ʽ��"
                    .ColWidth(i) = 930
                Case "ʱ��"
                    .ColWidth(i) = 1130
                Case "��Ŀ"
                    .ColWidth(i) = 1500
            End Select
            .ColData(i) = .ColWidth(i)
        Next
        
        lngMaxLength = Len(Mid(gstrDec, 3))
        If mrsBalance.RecordCount > 0 Then
            For i = 1 To mrsBalance.RecordCount
                lngP = InStr(1, CStr(mrsBalance!���ʽ��), ".")
                If lngP > 0 Then
                    lngP = Len(Mid(CStr(mrsBalance!���ʽ��), lngP + 1))
                    If lngP > lngMaxLength Then lngMaxLength = lngP
                End If
                mrsBalance.MoveNext
            Next
            mrsBalance.MoveFirst
        End If
        mstrDec = "0." & String(lngMaxLength, "0")
        
        For i = 1 To .Rows - 1
            .Row = i
            .Col = .Cols - 1
            If mbytFunc = 0 Then
                .CellBackColor = 12900351
            Else
                .CellBackColor = txtMoney.BackColor
            End If
            .Col = .Cols - 2
            .CellBackColor = 12900351
            .TextMatrix(i, COL_δ����) = LTrim(Format(.TextMatrix(i, COL_δ����), mstrDec))
            .TextMatrix(i, COL_���ʽ��) = LTrim(Format(.TextMatrix(i, COL_���ʽ��), mstrDec))
        Next
        .Redraw = True
    End With
    'ҽ��Ԥ����֮ǰ����ʾ���ʽ��ϼ�
    txtTotal.Text = Format(GetBalanceSum, mstrDec)
    txtTotal.Tag = txtTotal.Text
    dblMoney = Val(txtTotal.Text)
    '��ʾԤ����ϸ
    'mbln����תסԺ:36984
    strסԺ���� = ""
    If mbytFunc <> 0 Then
        strסԺ���� = IIf(gbln����ָ��Ԥ���� And mbln����תסԺ = False, IIf(mstrTime = "", mstrAllTime, mstrTime), "")
    End If
    
    Set mrsDeposit = GetDeposit(mrsInfo!����ID, mblnDateMoved, strסԺ����, mbln����תסԺ, mstrPepositDate, mintԤ�����)
    If Not mrsDeposit.EOF Then
        With mshDeposit
            .Redraw = False
            .Rows = mrsDeposit.RecordCount + 1
            For i = 1 To mrsDeposit.RecordCount
                .Row = i
                .Col = COLDeposit.��Ԥ��: .CellBackColor = txtMoney.BackColor
                .Col = COLDeposit.���: .CellBackColor = 12900351
                
                .RowData(i) = IIf(IsNull(mrsDeposit!��¼״̬), 0, mrsDeposit!��¼״̬)
                
                .TextMatrix(i, COLDeposit.ID) = mrsDeposit!ID
                .TextMatrix(i, COLDeposit.���ݺ�) = mrsDeposit!NO
                .TextMatrix(i, COLDeposit.Ʊ�ݺ�) = "" & mrsDeposit!Ʊ�ݺ�
                .TextMatrix(i, COLDeposit.����) = Format(mrsDeposit!����, "yyyy-MM-dd")
                .TextMatrix(i, COLDeposit.���㷽ʽ) = IIf(IsNull(mrsDeposit!���㷽ʽ), "", mrsDeposit!���㷽ʽ)
                .TextMatrix(i, COLDeposit.���) = Format(mrsDeposit!���, "0.00")
                If mbln����תסԺ Then
                    If Val(Nvl(mrsDeposit!���)) <= dblMoney Then
                        .TextMatrix(i, COLDeposit.��Ԥ��) = Format(mrsDeposit!���, "0.00")
                        dblMoney = dblMoney - Round(Val(Nvl(mrsDeposit!���)), 2)
                    ElseIf dblMoney <> 0 Then
                        .TextMatrix(i, COLDeposit.��Ԥ��) = Format(dblMoney, "0.00")
                        dblMoney = 0
                    End If
                Else
                    .TextMatrix(i, COLDeposit.��Ԥ��) = Format(mrsDeposit!���, "0.00")
                End If
                mrsDeposit.MoveNext
            Next
            .Row = 1: .Col = .Cols - 1
            .Redraw = True
        End With
        lblTicketCount.Caption = "Ԥ�����վ�:" & mrsDeposit.RecordCount & "��"
    End If
 

                                
    '���˺�:30043
    If IIf(mstrTime = "", mstrAllTime, mstrTime) <> "" Then
        Call zlSetDefaultTime(Val(Nvl(mrsInfo!����ID)))
    End If
        
    
    Call GetPatiDate(dBegin, dEnd)
    
    
    txtPatiBegin.Text = Format(dBegin, txtPatiBegin.Format)
    txtPatiEnd.Text = Format(dEnd, txtPatiEnd.Format)
    txtPatiEnd.Tag = Format(dEnd, txtPatiEnd.Format)
    Call zlChangeDefaultTime
    'ҽ��Ԥ����
    If Not IsNull(mrsInfo!����) And (Not MCPAR.�������ú���ýӿ� Or MCPAR.�������ú���ýӿ� And mblnSetPar) And Not mblnNoInsure Then
        '��ȡ������ϸ
        Set rsDetail = GetVBalance(mbytFunc, mstrPrivs, mrsInfo!����, mrsInfo!����ID, IIf(mbytFunc = 0, "0", mstrTime), mDateBegin, mDateEnd, blnUpload, mblnDateMoved, mbytBaby, mbytMCMode = 1, mbytKind, mstrItem, mstrUnit, mstrClass, mstrChargeType)
        
        'ҽ���ӿ�:���ظ��ֱ������
        If mbytMCMode = 1 Then
            If MCPAR.����Ԥ���� Then
                If rsDetail.RecordCount = 0 Then
                    MsgBox "��ȡҽ��Ԥ��������ʧ��!", vbInformation, gstrSysName
                    Screen.MousePointer = 0: Exit Function
                End If
            
                mstrBalance = ""
                If Not gclsInsure.ClinicPreSwap(rsDetail, mstrBalance, mrsInfo!����, "1|1") Then
                    MsgBox "����ҽ��Ԥ����ʧ��!", vbInformation, gstrSysName
                    Screen.MousePointer = 0: Exit Function
                End If
            End If
        Else
            mstrBalance = gclsInsure.WipeoffMoney(rsDetail, mrsInfo!����ID, "" & mrsInfo!ҽ����, "1", mrsInfo!����, "|" & IIf(opt��;.Value, 0, 1))
        End If
        
        '��ʾ����ͳ�ﱨ���ܶ�
        curͳ��֧�� = 0: cur�����ʻ� = 0
        For i = 0 To UBound(Split(mstrBalance, "|"))
            strMoney = Split(mstrBalance, "|")(i)
            j = GetBalanceNature(Split(strMoney, ";")(0))
            If j = 3 Then
                cur�����ʻ� = cur�����ʻ� + Val(Split(strMoney, ";")(1))
            ElseIf j = 4 Then
                curͳ��֧�� = curͳ��֧�� + Val(Split(strMoney, ";")(1))
            End If
        Next
        lblҽ������.Caption = "ͳ��֧��:" & Format(curͳ��֧��, "0.00")
        lblҽ������.Visible = True
        
        '��ʾ�������
        mcur������� = gclsInsure.SelfBalance(mrsInfo!����ID, "" & mrsInfo!ҽ����, IIf(mbytMCMode = 1, 10, 40), mcur����͸֧, mrsInfo!����)
        lbl�����ʻ�.Caption = "�ʻ����:" & Format(mcur�������, "0.00")
        lbl�����ʻ�.Visible = True
        
        Call Form_Resize
        txtTotal.Enabled = False
        cmdOK.Enabled = mstrBalance <> "" Or (mbytMCMode = 1 And Not MCPAR.����Ԥ����)
        
        If gblnLED Then
            zl9LedVoice.DisplayBank "ҽ������:", "�ʻ����" & Format(mcur�������, "0.00"), "�ʻ�֧��" & Format(cur�����ʻ�, "0.00"), "ͳ��֧��" & Format(curͳ��֧��, "0.00")
            DatTmp = Time
            Do While Time < DateAdd("s", 4, DatTmp)
            Loop
        End If
    Else
        Call HideMoneyInfo
        
        txtTotal.Enabled = True
        cmdOK.Enabled = True
    End If
    
    strInfo = ShowMoney(True, , mty_ModulePara.bytMzDeposit)
    Call SortMoney
    
    mcurTotal = Val(txtTotal.Text) '�������õ������
    txtDate.Text = Format(zlDatabase.Currentdate, txtDate.Format)
    
    If tabCard.SelectedItem.Index <> 1 Then Call LoadCardData
    Screen.MousePointer = 0
        
    '��ʾδ���õĽ��㷽ʽ
    If strInfo <> "" Then
        Me.Refresh
        MsgBox strInfo, vbInformation, gstrSysName
    End If
    
    ShowBalance = True
End Function

Private Function GetBalanceNature(ByVal strName As String) As Integer
'����:����ָ���Ľ��㷽ʽ����,���ؽ�������,û���ҵ�ʱ,����0
    Dim i As Long
    
    For i = 1 To vsfMoney.Rows - 1
        If vsfMoney.TextMatrix(i, COLMoney.C0����) = strName Then
            GetBalanceNature = Val(vsfMoney.TextMatrix(i, COLMoney.C3����))
            Exit For
        End If
    Next
End Function

Private Sub GetStateIF()
'���ܣ���ȡ���˵�סԺ���������ÿ���,������Ŀ,��������,��С��������ʱ��
    Dim i As Long, DateThis As Date
    
    Call InitBalanceCondition
    
    mrsBalance.MoveFirst
    For i = 1 To mrsBalance.RecordCount
                
        '���Ϊ��,���ʾ�������
        If InStr("," & mstrAllTime & ",", "," & Nvl(mrsBalance!��ҳID, 0) & ",") = 0 Then
            mstrAllTime = mstrAllTime & "," & Nvl(mrsBalance!��ҳID, 0)
        End If
        
        If Trim(Nvl(mrsBalance!��������ID, "")) <> "" Then
            If Not IsNull(mrsBalance!����) Then
                If InStr("," & mstrAllUnit & ",", "," & mrsBalance!��������ID & ":" & mrsBalance!���� & ",") = 0 Then
                    mstrAllUnit = mstrAllUnit & "," & mrsBalance!��������ID & ":" & mrsBalance!����
                End If
            End If
        End If
        
        If Trim(Nvl(mrsBalance!��Ŀ, "")) <> "" Then
            If InStr("," & mstrALLItem & ",", ",'" & mrsBalance!��Ŀ & "',") = 0 Then
                mstrALLItem = mstrALLItem & ",'" & mrsBalance!��Ŀ & "'"
            End If
        End If
        If Trim(Nvl(mrsBalance!�շ����)) <> "" Then '34260
            If InStr("," & mstrALLChargeType & ",", ",'" & mrsBalance!�շ���� & "',") = 0 Then
                mstrALLChargeType = mstrALLChargeType & ",'" & mrsBalance!�շ���� & "'"
            End If
        End If
        '���Ϊ��,ָû�����÷�������
        If InStr("," & mstrAllClass & ",", ",'" & Nvl(mrsBalance!����, "��") & "',") = 0 Then
            mstrAllClass = mstrAllClass & ",'" & Nvl(mrsBalance!����, "��") & "'"
        End If
        
        '�Ƚ�ȡ�����Сֵ
        If gint����ʱ�� = 0 Then
            DateThis = mrsBalance!�Ǽ�ʱ��
        Else
            DateThis = mrsBalance!ʱ��
        End If
        If i = 1 Then
            mMinDate = DateThis
            mMaxDate = DateThis
        Else
            If DateThis < mMinDate Then mMinDate = DateThis
            If DateThis > mMaxDate Then mMaxDate = DateThis
        End If
        
        mrsBalance.MoveNext
    Next
    mstrAllTime = Mid(mstrAllTime, 2)
    mstrAllUnit = Mid(mstrAllUnit, 2)
    mstrALLItem = Mid(mstrALLItem, 2)
    If mstrALLChargeType <> "" Then mstrALLChargeType = Mid(mstrALLChargeType, 2) '34260
    mstrAllClass = Mid(mstrAllClass, 2)
    
    '��ʾ����ʱ��
    txtEnd.Text = Format(mMaxDate, txtEnd.Format)
    txtBegin.Text = Format(mMinDate, txtBegin.Format)
    mrsBalance.MoveFirst
End Sub
Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKIND.SetAutoReadCard (False)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If mrsInfo.State = 1 Then
        If txtPatient.Text <> mrsInfo!���� Then txtPatient.Text = mrsInfo!����
    End If
End Sub

Private Sub txtTotal_GotFocus()
    SelAll txtTotal
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
    Dim curMoney As Currency, i As Long
    
    If txtTotal.Locked Then Exit Sub
    If mrsInfo.State = 0 Then KeyAscii = 0: Exit Sub
    If mshDetail.TextMatrix(1, 0) = "" Then KeyAscii = 0: Exit Sub

    If KeyAscii = 13 Then
        If Not IsNumeric(txtTotal.Text) Then
            sta.Panels(2) = "�������": Beep
            txtTotal.Text = txtTotal.Tag
            SelAll txtTotal
        ElseIf Val(txtTotal.Text) <> 0 And Val(txtTotal.Text) > mcurTotal Then
            sta.Panels(2) = "������ܴ��ڱ��ν��ʵĽ��:" & Format(mcurTotal, mstrDec): Beep
            txtTotal.Text = txtTotal.Tag
            SelAll txtTotal
        Else
            '�Զ�����ϼƷ���
            sta.Panels(2) = ""
            curMoney = Format(txtTotal.Text, mstrDec)
            mshDetail.Redraw = False
            For i = mshDetail.Rows - 1 To 1 Step -1
                If curMoney = 0 Then
                    mshDetail.TextMatrix(i, COL_���ʽ��) = mstrDec
                Else
                    If Val(mshDetail.TextMatrix(i, COL_δ����)) >= curMoney Then
                        mshDetail.TextMatrix(i, COL_���ʽ��) = Format(curMoney, mstrDec)
                    ElseIf Val(mshDetail.TextMatrix(i, COL_δ����)) < curMoney Then
                        mshDetail.TextMatrix(i, COL_���ʽ��) = Format(mshDetail.TextMatrix(i, COL_δ����), mstrDec)
                    End If
                    curMoney = curMoney - Val(mshDetail.TextMatrix(i, COL_���ʽ��))
                End If
            Next
            If curMoney <> 0 Then
                mshDetail.TextMatrix(1, COL_���ʽ��) = Format(Val(mshDetail.TextMatrix(1, COL_���ʽ��)) + curMoney, mstrDec)
            End If
            Call ShowMoney(True, , mty_ModulePara.bytMzDeposit)
            
            mshDetail.Redraw = True
            mshDeposit.SetFocus
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtTotal_LostFocus()
    If mbytInState = 1 Then Exit Sub
    If Not IsNumeric(txtTotal.Text) Then
        txtTotal.SetFocus
    ElseIf CCur(txtTotal.Tag) <> CCur(txtTotal.Text) Then
        txtTotal.Text = Format(txtTotal.Tag, mstrDec)
    End If
End Sub

Private Sub AdjustDeposit()
'����:��ʼ��Ԥ�����б�
    Dim i As Integer
    
    Call zlControl.MshSetFormat(mshDeposit, IIf(mstrInNO <> "" Or chkCancel.Value = 1, mstrDepositRHeader, mstrDepositHeader), App.ProductName & "\" & Me.Name, , , Not Visible)
    mshDeposit.FixedAlignment(COLDeposit.���㷽ʽ) = 1  '���ǵ�800*600���й�����ʱ�Բ���,�����
    
    '��0����ID
    For i = 0 To mnuViewToolCols.UBound
        If Not mnuViewToolCols(i).Checked And mnuViewToolCols(i).Visible Then
            If i + 1 < mshDeposit.Cols Then mshDeposit.ColWidth(i + 1) = 0
        End If
    Next
End Sub

Private Sub mshDeposit_DblClick()
    If Not txtMoney.Visible And mshDeposit.Row >= 1 And mshDeposit.Col = COLDeposit.��Ԥ�� _
        And mbytInState = 0 And chkCancel.Value = Unchecked _
        And mrsInfo.State = adStateOpen Then
        With txtMoney
            .Left = fraBalance.Left + mshDeposit.Left + mshDeposit.CellLeft + 15
            .Top = fraBalance.Top + mshDeposit.Top + mshDeposit.CellTop + (mshDeposit.CellHeight - txtMoney.Height) / 2 - 15
            .Width = mshDeposit.CellWidth - 60
            .ForeColor = mshDeposit.CellForeColor
            .BackColor = mshDeposit.CellBackColor
            .Alignment = 1
            .Text = mshDeposit.TextMatrix(mshDeposit.Row, mshDeposit.Col)
            .SelStart = 0: .SelLength = Len(.Text)
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub mshDeposit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mshDeposit.Row >= 1 Then
            If mshDeposit.Row < mshDeposit.Rows - 1 Then
                mshDeposit.Row = mshDeposit.Row + 1
                mshDeposit.Col = mshDeposit.Cols - 1
                If mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(0) - 2) > 1 Then
                    mshDeposit.TopRow = mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(1) - 2)
                End If
            Else
                vsfMoney.SetFocus
            End If
        End If
    End If
End Sub

Private Sub mshDeposit_KeyPress(KeyAscii As Integer)
    If Not txtMoney.Visible And mshDeposit.Row >= 1 And mshDeposit.Col = COLDeposit.��Ԥ�� _
        And KeyAscii <> 13 And mbytInState = 0 And chkCancel.Value = Unchecked _
        And mrsInfo.State = adStateOpen Then
        If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        With txtMoney
            .Left = fraBalance.Left + mshDeposit.Left + mshDeposit.CellLeft + 15
            .Top = fraBalance.Top + mshDeposit.Top + mshDeposit.CellTop + (mshDeposit.CellHeight - txtMoney.Height) / 2 - 15
            .Width = mshDeposit.CellWidth - 60
            .ForeColor = mshDeposit.CellForeColor
            .BackColor = mshDeposit.CellBackColor
            .Alignment = 1
            .Text = Chr(KeyAscii)
            .SelStart = 1
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub mshDeposit_LeaveCell()
    txtMoney.Visible = False
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    Dim blnCent As Boolean, i As Long
    
    If KeyAscii <> 13 Then
        '��������
        If Not (txtMoney.Left > fraBalance.Left And txtMoney.Top > vsfMoney.Top + fraBalance.Top And vsfMoney.Col = 2) Then
            If InStr(txtMoney.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0: Beep: Exit Sub
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        '�������
        Else
            If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Else
        KeyAscii = 0
        sta.Panels(2) = ""
        If Not (txtMoney.Left > fraBalance.Left And txtMoney.Top > vsfMoney.Top + fraBalance.Top And vsfMoney.Col = 2) Then
            If Trim(txtMoney.Text) = "" Then
                sta.Panels(2) = "���������"
                SelAll txtMoney: Call Beep: Exit Sub
            ElseIf Not IsNumeric(Trim(txtMoney.Text)) Then
                sta.Panels(2) = "�����˷Ƿ���"
                SelAll txtMoney: Call Beep: Exit Sub
            End If
        Else '�����������������ַ�
            If InStr(txtMoney.Text, "'") > 0 Or InStr(txtMoney.Text, "|") > 0 Or InStr(txtMoney.Text, ",") > 0 Then
                Call Beep: Exit Sub
            End If
        End If
        If txtMoney.Left < fraBalance.Left Then
            '�ڷ�����ϸ�б���:����ϵͳ������С������λ��
            txtMoney.Text = Format(Val(txtMoney.Text), mstrDec)
            
            '�޸Ĳ��ܳ�������
            If Val(txtMoney.Text) > Val(mshDetail.TextMatrix(mshDetail.Row, COL_δ����)) Then
                txtMoney.Text = Val(mshDetail.TextMatrix(mshDetail.Row, COL_δ����))
            End If
            
            mshDetail.TextMatrix(mshDetail.Row, mshDetail.Col) = Format(Val(txtMoney.Text), mstrDec)
            
            txtMoney.Visible = False
'''            Call zlClear���㿨
            Call ShowMoney(True, , mty_ModulePara.bytMzDeposit)
            
            If mshDetail.Row = mshDetail.Rows - 1 Then
                '��һ�ؼ�����
                mshDeposit.SetFocus
            Else
                '��һ�д���
                mshDetail.Row = mshDetail.Row + 1
                If mshDetail.Row - (mshDetail.Height \ mshDetail.RowHeight(0) - 2) > 1 Then
                    mshDetail.TopRow = mshDetail.Row - (mshDetail.Height \ mshDetail.RowHeight(1) - 2)
                End If
                mshDetail.Col = GetColNum("���ʽ��") ' mshDetail.Cols - 1
                mshDetail.SetFocus
            End If
        ElseIf txtMoney.Top > fraBalance.Top + vsfMoney.Top Then
            '�ڽ������б���
            If vsfMoney.Col <> 1 Then
                '��������
                vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.Col) = Trim(txtMoney.Text)
                Call Calc�Ҳ�
            Else
                '���������:����䵽0.00
                txtMoney.Text = Format(Val(txtMoney.Text), "0.00")
                
                If Val(txtMoney.Text) <> 0 Then
                    If Val(vsfMoney.TextMatrix(vsfMoney.Row, COLMoney.C3����)) = 1 Then
                        '��������ֽ���������,�����Ҫ����ֱ���ֻ׼�䵽0.0
                        blnCent = True
                        If gBytMoney = 0 Then blnCent = False
                        If blnCent And Not IsNull(mrsInfo!����) Then
                            If Not MCPAR.�ֱҴ��� Then blnCent = False
                        End If
                        If blnCent Then txtMoney.Text = Format(CentMoney(Val(txtMoney.Text)), "0.00")
                    ElseIf Val(vsfMoney.TextMatrix(vsfMoney.Row, COLMoney.C3����)) = 3 Then
                        '�����ʻ����
                        If Val(txtMoney.Text) < 0 Then
                            MsgBox "�����ʻ����������Ϊ������", vbInformation, gstrSysName
                            Call zlControl.TxtSelAll(txtMoney):  Exit Sub
                        End If
                        '�����������ص�ԭʼ�����޶�(�����ʻ�����͸֧ʱ���ж�)
                        If Val(txtMoney.Text) > mcur�����޶� And mcur�����޶� <> 0 And mcur����͸֧ = 0 And mbln���ʽ��� Then
                            MsgBox "����Ľ������˲��˿�֧���ĸ����ʻ��޶�:" & Format(mcur�����޶�, "0.00") & "��", vbInformation, gstrSysName
                            Call zlControl.TxtSelAll(txtMoney):  Exit Sub
                        End If
                        '������������͸֧���
                        If mcur������� - Val(txtMoney.Text) < -1 * mcur����͸֧ Then
                            MsgBox "�ʻ����:" & Format(mcur�������, "0.00") & _
                                IIf(mcur����͸֧ = 0, "", "(" & "����͸֧:" & Format(mcur����͸֧, "0.00") & ")") & _
                                "����Ҫ����Ľ�", vbInformation, gstrSysName
                            Call zlControl.TxtSelAll(txtMoney):  Exit Sub
                        End If
                    End If
                End If
            
                vsfMoney.TextMatrix(vsfMoney.Row, vsfMoney.Col) = Format(Val(txtMoney.Text), "0.00")
                Call ShowMoney(False, GetDefaultRow <> vsfMoney.Row, mty_ModulePara.bytMzDeposit)   '�޸ĺ��Զ���ƽ,���ǵ�ǰ����ȱʡ���㷽ʽ��
            End If
            
            txtMoney.Visible = False
            
            If vsfMoney.Col < vsfMoney.Cols - 2 Then
                vsfMoney.Col = vsfMoney.Col + 1
                vsfMoney.SetFocus
            Else
                If vsfMoney.Row = vsfMoney.Rows - 1 Then
                    '��һ�ؼ�����
                    If GetӦ�� > 0 And txt�ɿ�.Visible Then
                        txt�ɿ�.SetFocus
                    ElseIf cmdOK.Visible And cmdOK.Enabled Then
                        cmdOK.SetFocus
                    End If
                Else
                    '��һ�д���
                    vsfMoney.Row = vsfMoney.Row + 1
                    vsfMoney.Col = 1
                    If vsfMoney.Row - (vsfMoney.Height \ vsfMoney.RowHeight(0) - 2) > 1 Then
                        vsfMoney.TopRow = vsfMoney.Row - (vsfMoney.Height \ vsfMoney.RowHeight(1) - 2)
                    End If
                    vsfMoney.SetFocus
                End If
            End If
        Else
            '�ڳ�Ԥ���б���:����䵽0.00
            txtMoney.Text = Format(Val(txtMoney.Text), "0.00")
            
            '�޸Ĳ��ܳ�������
            If Val(txtMoney.Text) > Val(mshDeposit.TextMatrix(mshDeposit.Row, COLDeposit.���)) Then
                txtMoney.Text = Val(mshDeposit.TextMatrix(mshDeposit.Row, COLDeposit.���))
            End If
            mshDeposit.TextMatrix(mshDeposit.Row, mshDeposit.Col) = Format(Val(txtMoney.Text), "0.00")
            
            txtMoney.Visible = False
            Call ShowMoney(False, , mty_ModulePara.bytMzDeposit)
            
            If mshDeposit.Row = mshDeposit.Rows - 1 Then
                '��һ�ؼ�����
                vsfMoney.SetFocus
            Else
                '��һ�д���
                mshDeposit.Row = mshDeposit.Row + 1
                If mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(0) - 2) > 1 Then
                    mshDeposit.TopRow = mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(1) - 2)
                End If
                mshDeposit.Col = mshDeposit.Cols - 1
                mshDeposit.SetFocus
            End If
        End If
    End If
End Sub

Private Function ReadBalance(strNo As String) As Boolean
'���ܣ��鿴������ʱ,��ȡ����ʾ���ʵ�
'������strfullno=���ݺ�
'���أ�
'     -1:�ɹ�
'      0:ʧ��
'      1:�õ��ݲ�����
'      2:�õ���������(mblnViewCancel=Trueʱ��Ч)
'      3:�������ݲ�����
    Dim rsTmp As ADODB.Recordset, strFullNO As String
    Dim lngID As Long, i As Long, j As Long, lngDefault As Long
    Dim strSql As String, dMax As Date, dMin As Date, blnUndo As Boolean
    Dim curTmp As Currency, curMoney As Currency, curDeposit As Currency
    Dim lngMaxLength As Long, lngP As Long, lng����ID As Long
    Dim rsUnit As ADODB.Recordset, rsFee As New ADODB.Recordset
    Dim strTable As String
    
    On Error GoTo errH
    
    '��������
    strFullNO = GetFullNO(strNo, 15)
    
    strTable = IIf(mblnNOMoved, "H", "") & "���˽��ʼ�¼"
    strSql = _
    "Select A.ID,A.ʵ��Ʊ�� as Ʊ�ݺ�,B.����ID,B.�����,B.סԺ��,Nvl(D.��Ժ����,B.��ǰ����) as ��ǰ����, " & _
    "       Nvl(E.����,C.����) as ��ǰ����," & _
    "       Nvl(D.�ѱ�,B.�ѱ�) as �ѱ�,nvl(D.����,B.����) as ����,nvl(D.�Ա�,B.�Ա�) as �Ա�,B.����,A.�շ�ʱ��,A.��ʼ����,A.��������,A.��ע,A.ԭ��,A.��������" & _
    " From " & strTable & " A,������Ϣ B,���ű� C,������ҳ D,���ű� E" & _
    " Where A.����ID=B.����ID(+) And B.��ǰ����ID=C.ID(+) And D.��Ժ����ID=E.ID(+)" & _
    "       And B.����ID=D.����ID(+) And Nvl(B.��ҳID,0)=D.��ҳID(+) " & _
    "       And A.NO=[1] And A.��¼״̬ " & IIf(mblnViewCancel, "= 2", "In(1,3)")
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strFullNO)
    If rsTmp.EOF Then
        MsgBox "û�з��ָý��ʵ���,�����Ѿ����ϣ�", vbInformation, gstrSysName
        Exit Function
    End If
    If Not GetMinMaxDate(rsTmp!ID, dMin, dMax, mblnNOMoved) Then
        MsgBox "�ý��ʵ������ݲ���ȷ��û�з��ֽ��ʵķ�����ϸ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    cboNO.Text = strFullNO
    txtInvoice.Text = Nvl(rsTmp!Ʊ�ݺ�)
    
    lng����ID = Val(Nvl(rsTmp!����ID))
    If Val(Nvl(rsTmp!��������)) = 0 Then
        lblTitle.Caption = gstrUnitName & "���˽��ʵ�"
    ElseIf Val(Nvl(rsTmp!��������)) = 1 Then
        lblTitle.Caption = gstrUnitName & "���ﲡ�˽��ʵ�"
    Else
        lblTitle.Caption = gstrUnitName & "סԺ���˽��ʵ�"
    End If
    
    '��ȡ�������
    If Val(Nvl(rsTmp!��������)) = 0 Then
        strSql = "Select Sum(Ԥ�����) As Ԥ�����,Sum(�������) As ������� From ������� Where ����ID= [1] And ����=1"
        Set rsFee = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
    Else
        strSql = "Select Ԥ�����,������� From ������� Where ����ID= [1] And ����=1 And ����= [2]"
        Set rsFee = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, Val(Nvl(rsTmp!��������)))
    End If
    '60615,������,2013-12-20,״̬����ʾԤ�������ý���ʣ�����
    If rsFee.RecordCount <> 0 Then
        sta.Panels(3).Text = "Ԥ��:" & Format(Nvl(rsFee!Ԥ�����), "0.00") & _
                             "/����:" & Format(Nvl(rsFee!�������), "0.00") & _
                             "/ʣ��:" & Format(Val(Nvl(rsFee!Ԥ�����)) - Val(Nvl(rsFee!�������)), "0.00")
    End If
    
    '����Ƿ��Լ��λ����:����:35090
    If Val(Nvl(rsTmp!����ID)) = 0 Then
        If Nvl(rsTmp!ԭ��) <> "" Then
            txtPatient.Text = Nvl(rsTmp!ԭ��)
        Else
            strSql = "" & _
            "   Select  D.���� " & _
            "   From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A, ������Ϣ C, ��Լ��λ D " & _
            "   Where A.����ID=[1]  And A.����ID=C.����ID And C.��ͬ��λid = D.ID(+) and Rownum=1 " & _
            "    Union ALL " & _
            "   Select  D.���� " & _
            "   From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ A, ������Ϣ C, ��Լ��λ D " & _
            "   Where A.����ID=[1] And C.��ͬ��λid = D.ID(+) and Rownum=1 " & _
            "   "
            Set rsUnit = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(rsTmp!ID)))
            If Not rsUnit.EOF Then
                txtPatient.Text = Nvl(rsUnit!����)
            Else
                txtPatient.Text = "δ�ҵ���Լ��λ"
            End If
        End If
        txtPatient.Tag = "��Լ��λ"
    Else
        txtPatient.Text = Nvl(rsTmp!����)
        txtPatient.Tag = Val(Nvl(rsTmp!����ID))
    End If
    
    txtSex.Text = Nvl(rsTmp!�Ա�)
    txtOld.Text = Nvl(rsTmp!����)
    txt�ѱ�.Text = Nvl(rsTmp!�ѱ�)
    txtDate.Text = Format(rsTmp!�շ�ʱ��, "yyyy-MM-dd HH:mm:ss")
    
    '����65105,������:���˲�������������������ʾ
    Select Case Val(Nvl(rsTmp!��������))
        '10.29��ǰ�����ͣ���������
        Case 0
        Case 1
            txt��ʶ��.Text = Nvl(rsTmp!�����)
            txt��ʶ��.Visible = True
            lbl��ʶ��.Visible = True
            lbl��ʶ��.Caption = "�����"
        Case 2
            txt��ʶ��.Text = Nvl(rsTmp!סԺ��)
            txt��ʶ��.Visible = True
            lbl��ʶ��.Visible = True
            lbl��ʶ��.Caption = "סԺ��"
            
            If Not IsNull(rsTmp!��ǰ����) Then
                txtBed.Text = rsTmp!��ǰ����
                txtBed.Visible = True
                lblBed.Visible = True
            End If
            
            If Not IsNull(rsTmp!��ǰ����) Then
                txt����.Text = rsTmp!��ǰ����
                txt����.Visible = True
                lbl����.Visible = True
            End If
    End Select
    
    txtBegin.Text = Format(dMin, txtBegin.Format)
    txtEnd.Text = Format(dMax, txtEnd.Format)
    txt��ע.Text = Nvl(rsTmp!��ע)
    If Not IsNull(rsTmp!��ʼ����) Then
        txtPatiBegin.Text = Format(rsTmp!��ʼ����, "yyyy-MM-dd")
    End If
    
    If Not IsNull(rsTmp!��������) Then
        txtPatiEnd.Text = Format(rsTmp!��������, "yyyy-MM-dd")
    End If
    
    lngID = rsTmp!ID
    
    '��Ԥ���嵥
    Me.lblSpare.Visible = False
    Call zlControl.MshSetFormat(mshDeposit, mstrDepositRHeader, App.ProductName & "\" & Me.Name, , , Not Visible)
    '��0����ID
    For i = 1 To mshDeposit.Cols - 1
        If Not mnuViewToolCols(i - 1).Checked Then mshDeposit.ColWidth(i) = 0
    Next
    
    Set rsTmp = GetBalanceDeposit(lngID, mblnNOMoved)
    If Not rsTmp.EOF Then Set mshDeposit.DataSource = rsTmp
    
    curDeposit = 0
    For i = 1 To mshDeposit.Rows - 1
        curDeposit = curDeposit + Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 1))
    Next
    lblDeposit.Caption = "��Ԥ��:" & Format(curDeposit, "0.00")
    lblDeposit.Tag = curDeposit
    lblTicketCount.Caption = "Ԥ�����վ�:" & rsTmp.RecordCount & "��"
    '���ʲ����嵥,δ�õĽ��㷽ʽҲ�г�,�Ա�����ʱ,�������ҽ���������ֽ�
    '---------------------------------------------------------------------------------------------------------------------
    mrs���㷽ʽ.Filter = ""
    With vsfMoney
        .Redraw = False
        .Clear
        .Rows = 2: .Cols = 5
        
        .TextMatrix(0, COLMoney.C0����) = "���㷽ʽ"
        .TextMatrix(0, COLMoney.C1���) = "���"
        .TextMatrix(0, COLMoney.C2����) = "�������"
        .TextMatrix(0, COLMoney.C3����) = "����"
        .TextMatrix(0, COLMoney.C4ȱʡ) = "ȱʡ"
        
        .Rows = mrs���㷽ʽ.RecordCount + 1
        For i = 1 To mrs���㷽ʽ.RecordCount
            .TextMatrix(i, COLMoney.C0����) = mrs���㷽ʽ!����
            .TextMatrix(i, COLMoney.C3����) = mrs���㷽ʽ!����
            .TextMatrix(i, COLMoney.C4ȱʡ) = mrs���㷽ʽ!ȱʡ
            mrs���㷽ʽ.MoveNext
        Next
        
        .FixedAlignment(0) = 4: .ColAlignment(0) = 1: .ColWidth(0) = 1200
        .FixedAlignment(1) = 4: .ColAlignment(1) = 7: .ColWidth(1) = 1100
        .FixedAlignment(2) = 4: .ColAlignment(2) = 1: .ColWidth(2) = 1450
        .FixedAlignment(3) = 4: .ColAlignment(3) = 1: .ColWidth(3) = 0
        .FixedAlignment(4) = 4: .ColAlignment(4) = 1: .ColWidth(4) = 0
        
        .Redraw = True
        
        '�����嵥
        Me.lblSpare.Visible = False
        Set rsTmp = GetBalancePay(lngID, mblnNOMoved)
        
        For i = 1 To rsTmp.RecordCount
            For j = 1 To .Rows - 1
                If rsTmp!���㷽ʽ = .TextMatrix(j, COLMoney.C0����) Then
                    .TextMatrix(j, COLMoney.C1���) = Format(rsTmp!���, "0.00")
                    .TextMatrix(j, COLMoney.C2����) = "" & rsTmp!�������
                    Exit For
                End If
            Next
            rsTmp.MoveNext
        Next
        For i = 1 To .Rows - 1
            If Nvl(.TextMatrix(i, COLMoney.C3����)) = 9 Then
                .Row = i
                .Col = 0
                .CellForeColor = vbRed
                Exit For
            End If
        Next
        
        '��ҽ����������ʱ,����֧�ֻ��˵�ҽ�������Ƶ�ȱʡ���㷽ʽ��
        mblnҽ������ȫ�� = True
        If mbytInState = 0 And mintInsure <> 0 Then        '
            For i = 1 To .Rows - 1
                If Nvl(.TextMatrix(i, COLMoney.C4ȱʡ)) = 1 Then lngDefault = i: Exit For
                If Nvl(.TextMatrix(i, COLMoney.C3����)) = 1 Then lngDefault = i: Exit For
            Next
            If lngDefault = 0 Then MsgBox "û������ȱʡ���㷽ʽ,���ʳ���Ҳû���ֽ���㷽ʽ����,�޷�����ҽ����������!", vbInformation, gstrSysName: Exit Function
                    
            .Row = lngDefault: .Col = 0
            .CellFontBold = True
            'ҽ����֧�����ϵĽ��㷽ʽ��Ϊȱʡ����
            For i = 1 To .Rows - 1
                If (.TextMatrix(i, COLMoney.C3����) = 3 Or .TextMatrix(i, COLMoney.C3����) = 4) And Val(.TextMatrix(i, COLMoney.C1���)) <> 0 Then
                    '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                    If mbytMCMode = 1 And Not MCPAR.���ﲡ�˽������� Then
                        blnUndo = Val(.TextMatrix(i, COLMoney.C3����)) = 3
                    Else
                       'lng����ID:49084
                        blnUndo = Not gclsInsure.GetCapability(IIf(mbytMCMode = 1, support�����������, supportסԺ��������), lng����ID, mintInsure, .TextMatrix(i, COLMoney.C0����))
                    End If
                    If blnUndo Then
                        .TextMatrix(lngDefault, COLMoney.C1���) = Format(Val(.TextMatrix(lngDefault, COLMoney.C1���)) + Val(.TextMatrix(i, COLMoney.C1���)), "0.00")
                        .TextMatrix(i, COLMoney.C1���) = ""
                        mblnҽ������ȫ�� = False
                    Else
                        .Row = i: .Col = 0: .CellBackColor = txtMoney.BackColor
                        .Col = 1: .CellBackColor = txtMoney.BackColor
                        .Col = 2: .CellBackColor = txtMoney.BackColor
                    End If
                End If
            Next
            If Not mblnҽ������ȫ�� Then
                '������ֽ�,���зֱҴ���
                If .TextMatrix(lngDefault, COLMoney.C3����) = 1 And Val(.TextMatrix(lngDefault, COLMoney.C1���)) <> 0 And MCPAR.�ֱҴ��� Then
                    .TextMatrix(lngDefault, COLMoney.C1���) = Format(CentMoney(Val(.TextMatrix(lngDefault, COLMoney.C1���))), "0.00")
                End If
                For i = 1 To .Rows - 1
                    curMoney = curMoney + Val(.TextMatrix(i, COLMoney.C1���))
                Next
            End If
        End If
    End With
    
    
    
    
    '������ϸ
    'סԺ���ü�¼��[סԺ],[����],ʱ��,[���ݺ�],��Ŀ,��Ŀ,[Ӥ����],���ʽ��
    '------------------------------------------------------------------------------------
    strSql = "" & _
    "   Select  '����' as סԺ,A.����ʱ��,A.NO,A.���,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.Ӥ����,A.���ʽ��,A.��������ID " & _
    "   From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A" & _
    "   Where A.����ID=[1]" & _
    "    Union ALL " & _
    "   Select  Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��') as סԺ,A.����ʱ��,A.NO,A.���,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.Ӥ����,A.���ʽ��,A.��������ID " & _
    "   From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ A" & _
    "   Where A.����ID=[1] " & _
    "   "
    
    
    strSql = _
    "  Select   A.סԺ," & _
    "            Nvl(B.����,'δ֪') as ����,To_Char(A.����ʱ��,'YYYY-MM-DD') as ʱ��," & _
    "            A.NO as ���ݺ�,Nvl(E.����,D.����) as ��Ŀ,A.�վݷ�Ŀ as ��Ŀ," & _
    "            Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.���ʽ��" & _
    " From (" & strSql & ") A,���ű� B,�շ���ĿĿ¼ D,�շ���Ŀ���� E" & _
    " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=D.ID" & _
    "           And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    "            " & _
    " Order by סԺ Desc,ʱ�� Desc,���ݺ� Desc,���"
'
'
'    strSQL = _
'    " Select Decode(A.��ҳID,NULL,'����','��'||A.��ҳID||'��') as סԺ," & _
'    "       Nvl(B.����,'δ֪') as ����,To_Char(A.����ʱ��,'YYYY-MM-DD') as ʱ��," & _
'    "       A.NO as ���ݺ�,Nvl(E.����,D.����) as ��Ŀ,A.�վݷ�Ŀ as ��Ŀ," & _
'    "       Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.���ʽ��" & _
'    " From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ A,���ű� B,�շ���ĿĿ¼ D,�շ���Ŀ���� E" & _
'    " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=D.ID" & _
'    "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
'    "       And A.����ID=[1] " & vbCrLf & _
'    " Union ALL " & _
'    " Select  '����' as סԺ," & _
'    "       Nvl(B.����,'δ֪') as ����,To_Char(A.����ʱ��,'YYYY-MM-DD') as ʱ��," & _
'    "       A.NO as ���ݺ�,Nvl(E.����,D.����) as ��Ŀ,A.�վݷ�Ŀ as ��Ŀ," & _
'    "       Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����,A.���ʽ��" & _
'    " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,���ű� B,�շ���ĿĿ¼ D,�շ���Ŀ���� E" & _
'    " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=D.ID" & _
'    "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
'    "       And A.����ID=[1]" & _
'    " Order by סԺ Desc,ʱ�� Desc,���ݺ� Desc,���"
'
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
    If rsTmp.EOF Then Exit Function
    
    With mshDetail
        .Redraw = False
        Call ClearDetail
        If Not rsTmp.EOF Then Set .DataSource = rsTmp
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4
            If i <= 4 Then .MergeCol(i) = True
            Select Case .TextMatrix(0, i)
                Case "סԺ", "Ӥ����", "���ݺ�"
                    .ColAlignment(i) = 4
                Case "����", "��Ŀ", "��Ŀ", "ʱ��"
                    .ColAlignment(i) = 1
                Case "���ʽ��"
                    .ColAlignment(i) = 7
            End Select
            
            Select Case .TextMatrix(0, i)
                Case "����", "סԺ"
                    .ColWidth(i) = 0
                Case "Ӥ����"
                    .ColWidth(i) = 750
                Case "��Ŀ"
                    .ColWidth(i) = 800
                Case "���ʽ��", "���ݺ�"
                    .ColWidth(i) = 950
                Case "ʱ��"
                    .ColWidth(i) = 1130
                Case "��Ŀ"
                    .ColWidth(i) = 2300
            End Select
            .ColData(i) = .ColWidth(i)
        Next
        
        lngMaxLength = Len(Mid(gstrDec, 3))
        If rsTmp.RecordCount > 0 Then
            For i = 1 To rsTmp.RecordCount
                lngP = InStr(1, CStr(rsTmp!���ʽ��), ".")
                If lngP > 0 Then
                    lngP = Len(Mid(CStr(rsTmp!���ʽ��), lngP + 1))
                    If lngP > lngMaxLength Then lngMaxLength = lngP
                End If
                rsTmp.MoveNext
            Next
            rsTmp.MoveFirst
        End If
        mstrDec = "0." & String(lngMaxLength, "0")
        
        curTmp = 0
        For i = 1 To .Rows - 1
            .TextMatrix(i, .Cols - 1) = Format(.TextMatrix(i, .Cols - 1), mstrDec)
            curTmp = curTmp + Val(.TextMatrix(i, .Cols - 1))
        Next
        txtTotal.Text = Format(curTmp, mstrDec)
        curTmp = Val(txtTotal.Text)
        .Redraw = True
        
        If mbytInState = 0 And mintInsure <> 0 And Not mblnҽ������ȫ�� Then
            '����
            mcur����� = curDeposit + curMoney - curTmp
            vsfMoney.ToolTipText = "��������,�����:" & Format(mcur�����, mstrDec)
        Else
            mcur����� = 0
        End If
    End With
    
    If mbytInState = 0 Then
        mtySquareCard.bln������ = zlIsExistsSquareCard(strNo)
    Else
        mtySquareCard.bln������ = False
    End If
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COLMoney.C1���)) = 0 Then
                .RowHidden(i) = True
            Else
                .RowHidden(i) = False
            End If
        Next i
        .Refresh
    End With
    ReadBalance = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetDefaultRow() As Long
'���ܣ���ȡ��ǰȱʡ���㷽ʽ�к�
    Dim i As Long, lngDefaultRow As Long, curBalance As Currency, curDeposit As Currency
    Dim strסԺ���� As String, strSql As String, rsTmp As ADODB.Recordset
    
    If mblnOneCard And mstrOneCard <> "" Then
        For i = 1 To vsfMoney.Rows - 1
            If vsfMoney.TextMatrix(i, COLMoney.C0����) = mstrOneCard Then
                lngDefaultRow = i: Exit For
            End If
        Next
    Else
        If mstrȱʡ���� <> "" Then
            For i = 1 To vsfMoney.Rows - 1
                If vsfMoney.TextMatrix(i, COLMoney.C0����) = mstrȱʡ���� Then
                    lngDefaultRow = i: Exit For
                End If
            Next
        Else
            '78882:�����˿�ȱʡ��Ԥ���ɿ���㷽ʽ�˿���û��ѡ�����������ȱʡ���ֽ��˿�
            '���Ԥ���ɿ��ж��ֽ��㷽ʽ��������˳����
            '        1.���п�(�ֹ���������п�,����Ϊ2���ҷ�֧Ʊ�Ľ��㷽ʽ)
            '        2.�ֽ�
            '        3.֧Ʊ
            '        4.�������㷽ʽ
            If mbytFunc = 1 Then
                curBalance = GetBalanceSum
                For i = 1 To mshDeposit.Rows - 1
                    curDeposit = curDeposit + Val(mshDeposit.TextMatrix(i, COLDeposit.��Ԥ��))
                Next i
                If curDeposit > curBalance Then
                    If mty_ModulePara.bln�����˿ʽ = False Then
                        'ȱʡ���ֽ�
                        For i = 1 To vsfMoney.Rows - 1
                            If Val(vsfMoney.TextMatrix(i, COLMoney.C3����)) = 1 Then  'û��ָ��ȱʡʱ���ֽ�Ϊȱʡ��
                                 lngDefaultRow = i
                                 GetDefaultRow = lngDefaultRow
                                 Exit Function
                            End If
                        Next
                    Else
                        'ȱʡ��Ԥ���ɿ���㷽ʽ
                        strסԺ���� = ""
                        If mbytFunc = 1 Then
                            strסԺ���� = IIf(gbln����ָ��Ԥ���� And mbln����תסԺ = False, IIf(mstrTime = "", mstrAllTime, mstrTime), "")
                        End If
                        
                        strSql = " Select a.���㷽ʽ, Decode(Nvl(b.����,0), 7, 1, 2, Decode(a.���㷽ʽ,'֧Ʊ',4,2), 1, 3, 5) As ˳�� From ����Ԥ����¼ A,���㷽ʽ B " & _
                                 " Where a.��¼���� = 1 And a.����id = [1] And a.Ԥ����� = 2 And a.���㷽ʽ = b.����(+) " & _
                                 IIf(strסԺ���� = "", "", " And a.��ҳID In (Select Column_Value From Table(f_str2list([2]))) ") & _
                                 " Order By ˳�� "
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(Nvl(mrsInfo!����ID)), strסԺ����)
                        If rsTmp.RecordCount <> 0 Then
                            For i = 1 To vsfMoney.Rows - 1
                                If vsfMoney.TextMatrix(i, COLMoney.C0����) = Nvl(rsTmp!���㷽ʽ) Then
                                     lngDefaultRow = i
                                     GetDefaultRow = lngDefaultRow
                                     Exit Function
                                End If
                            Next
                        End If
                    End If
                End If
            End If
            For i = 1 To vsfMoney.Rows - 1
                If Val(vsfMoney.TextMatrix(i, COLMoney.C3����)) = 1 Then  'û��ָ��ȱʡʱ���ֽ�Ϊȱʡ��
                     lngDefaultRow = i: Exit For
                End If
            Next
        End If
    End If
    
    GetDefaultRow = lngDefaultRow
End Function

Private Function GetBalanceSum() As Currency
    Dim i As Long, cur���ʺϼ� As Currency
    Dim lngCol As Long
    lngCol = GetColNum("���ʽ��")
    
    If lngCol <> COL_���ʽ�� Then Exit Function
    
    For i = 1 To mshDetail.Rows - 1
        cur���ʺϼ� = cur���ʺϼ� + Val(mshDetail.TextMatrix(i, lngCol))
    Next
    GetBalanceSum = cur���ʺϼ�
End Function

Private Function ShowMoney(blnFirst As Boolean, _
    Optional blnAutoCalc As Boolean = True, Optional bytMzDeposit As Byte = 2) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ú���ʾ����ĸ��ֽ��
    '���:blnFirst=�Ƿ����´��������ϸ,��Ԥ����,ҽ�����㲿��,�����һ�ε��ñ�����һ��
    '     blnAutoCalc=���ݲ���Զ���ƽ������
    '     bytMzDeposit-������������Ч,0-��ʾȫ��;1-������ݽ��ʽ������̯Ԥ��;2-Ԥ����ȫ��
    '����:
    '����:ҽ���ɱ������㲿��δ��������ʾ��
    '����:���˺�
    '����:2014-05-23 16:11:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngȱʡRow As Long, blnȱʡ�ֽ� As Boolean, i As Long, j As Long, lng��� As Long
    Dim cur���ʺϼ� As Currency, curMoney As Currency, curTemp As Currency
    Dim strMoney As String, strNone As String, strHave As String
    Dim blnCent As Boolean, curOwn As Currency, curTmp As Currency
    
    '�ж�ȱʡ���㷽ʽ�Ƿ��ֽ����ֽ����Զ���ƽʱ����ֱң�������������
    '���û������ȱʡ���㷽ʽ�����ֽ���Ϊȱʡ�Ĳ�ƽ���㷽ʽ(�����)
    '-----------------------------------------------------------------------------------------------------
    lngȱʡRow = GetDefaultRow
    For i = 1 To vsfMoney.Rows - 1
        If Val(vsfMoney.TextMatrix(i, COLMoney.C3����)) = 9 Then
            vsfMoney.TextMatrix(i, COLMoney.C1���) = 0
            lng��� = i: Exit For
        End If
    Next i
    If lngȱʡRow > 0 Then blnȱʡ�ֽ� = (Val(vsfMoney.TextMatrix(lngȱʡRow, COLMoney.C3����)) = 1)
    
    '�ж��Ƿ�Ӧ�ý��зֱҴ���
    blnCent = True
    If gBytMoney = 0 Then blnCent = False
    If Not IsNull(mrsInfo!����) And Not MCPAR.�ֱҴ��� Then blnCent = False
    
    '��ʾ���ʺϼƼ����ó�Ԥ���͸��ֽ�����
    '-----------------------------------------------------------------------------------------------------
    If blnFirst Then
        'ͳ�Ʋ���ʾ���ʽ��ϼ�
        cur���ʺϼ� = GetBalanceSum
        txtTotal.Text = Format(cur���ʺϼ�, mstrDec)
        txtTotal.Tag = txtTotal.Text
            
        '����ҽ�����㲿�ֽ��
        For i = 0 To UBound(Split(mstrBalance, "|"))
            strMoney = Split(mstrBalance, "|")(i)
            For j = 1 To vsfMoney.Rows - 1
                If vsfMoney.TextMatrix(j, COLMoney.C0����) = CStr(Split(strMoney, ";")(0)) _
                    And InStr(",3,4,", Val(vsfMoney.TextMatrix(j, COLMoney.C3����))) > 0 Then
                    '�����ʻ����������
                    If Val(vsfMoney.TextMatrix(j, COLMoney.C3����)) = 3 Then
                        '�����ʻ����֧�����
                        mbln���ʽ��� = True
                        mcur�����޶� = CCur(Split(strMoney, ";")(1))
                        
                        'ȱʡ���ܳ��������ʻ���������͸֧���
                        If mcur������� - CCur(Split(strMoney, ";")(1)) >= -1 * mcur����͸֧ Then
                            vsfMoney.TextMatrix(j, COLMoney.C1���) = Format(CCur(Split(strMoney, ";")(1)), "0.00") '������͸֧��Χ���㹻(����͸֧0Ϊ����)
                        Else
                            vsfMoney.TextMatrix(j, COLMoney.C1���) = "0.00"
                            MsgBox "�����ʻ������δ����,������ҽ������!", vbInformation, Me.Caption
                            cmdOK.Enabled = False
                        End If
                    Else
                        vsfMoney.TextMatrix(j, COLMoney.C1���) = Format(CCur(Split(strMoney, ";")(1)), "0.00")
                    End If
                    
                    If Val(Split(strMoney, ";")(2)) = 0 Then
                        vsfMoney.RowData(j) = 1 '�ý�����ɸ���
                    Else
                        vsfMoney.RowData(j) = 0 '�ý�������Ը���
                    End If
                    
                    '����ҽ���Ѵ���Ľ���
                    cur���ʺϼ� = cur���ʺϼ� - Format(Val(vsfMoney.TextMatrix(j, COLMoney.C1���)), "0.00")
                    strHave = strHave & ";" & CStr(Split(strMoney, ";")(0))
                    Exit For
                End If
            Next
            'δ����ҽ���ɱ������㷽ʽ
            If j = vsfMoney.Rows Then
                strNone = strNone & vbCrLf & vbTab & CStr(Split(strMoney, ";")(0)) & ":" & Format(CCur(Split(strMoney, ";")(1)), "0.00")
            End If
        Next
        
        '���˺�:��Խ��㿨���д���
        Call zlReCalcRequare(cur���ʺϼ�, strNone)
        
        '���ó�Ԥ��(���ʺϼ� - ���պϼ�)
        If mshDeposit.TextMatrix(1, COLDeposit.ID) <> "" Then
    
            If (mbytFunc <> 0 And (opt��Ժ.Value Or gbln��;������Ԥ��)) _
                Or (mbytFunc = 0 And bytMzDeposit = 2) Then
                'ȫ��������(����˾��˸�����)
                '1.��Ժ����
                '2.�������ȫ��
                For i = 1 To mshDeposit.Rows - 1
                    mshDeposit.TextMatrix(i, COLDeposit.��Ԥ��) = Format(Val(mshDeposit.TextMatrix(i, COLDeposit.���)), "0.00")
                    cur���ʺϼ� = cur���ʺϼ� - Val(mshDeposit.TextMatrix(i, COLDeposit.��Ԥ��))
                Next
            ElseIf (mbytFunc = 0 And bytMzDeposit = 0) Then
                '�������,��ʹ��Ԥ��
                For i = 1 To mshDeposit.Rows - 1
                    mshDeposit.TextMatrix(i, COLDeposit.��Ԥ��) = "0.00"
                Next
            Else
                '1.��;����ֻ���㹻��
                '2.�������ֻ���㹻��
                For i = 1 To mshDeposit.Rows - 1
                    If cur���ʺϼ� = 0 Then
                        mshDeposit.TextMatrix(i, COLDeposit.��Ԥ��) = "0.00"
                    Else
                        If Val(mshDeposit.TextMatrix(i, COLDeposit.���)) <= Format(cur���ʺϼ�, "0.00") Then
                            mshDeposit.TextMatrix(i, COLDeposit.��Ԥ��) = Format(Val(mshDeposit.TextMatrix(i, COLDeposit.���)), "0.00")
                        Else
                            mshDeposit.TextMatrix(i, COLDeposit.��Ԥ��) = Format(cur���ʺϼ�, "0.00")
                        End If
                        cur���ʺϼ� = cur���ʺϼ� - Val(mshDeposit.TextMatrix(i, COLDeposit.��Ԥ��))
                    End If
                Next
            End If
        End If
                    
        'ʣ��Ӧ�ɲ��ݳ������õ�ȱʡ���㷽ʽ
        If lngȱʡRow <> 0 Then
            If blnȱʡ�ֽ� And blnCent Then '�ֽ�ʱҪ���зֱҴ���
                vsfMoney.TextMatrix(lngȱʡRow, COLMoney.C1���) = Format(CentMoney(cur���ʺϼ�), "0.00")
            Else
                vsfMoney.TextMatrix(lngȱʡRow, COLMoney.C1���) = Format(cur���ʺϼ�, "0.00")
            End If
            cur���ʺϼ� = 0
        End If
    End If
    
    '��ʾ��ǰ��Ԥ������
    '-----------------------------------------------------------------------------------------------------
    curMoney = GetPaySum
    
    '�����ǲ��,��һ�����ֽ�,���Բ�����ֱ�
    curOwn = Val(txtTotal.Text) - curMoney
    txtOwe.Text = Format(curOwn, "0.00")
    
    '���ݲ���Զ���ƽ������'ʣ�ಿ�ݳ������õ�ȱʡ���㷽ʽ��
    '-----------------------------------------------------------------------------------------------------
    If blnAutoCalc And Val(txtOwe.Text) <> 0 And lngȱʡRow <> 0 Then
        curTmp = Val(vsfMoney.TextMatrix(lngȱʡRow, COLMoney.C1���)) + curOwn
        If Abs(curTmp) >= 0.01 Then
            If blnȱʡ�ֽ� And blnCent Then
                vsfMoney.TextMatrix(lngȱʡRow, COLMoney.C1���) = Format(CentMoney(curTmp), "0.00")
            Else
                vsfMoney.TextMatrix(lngȱʡRow, COLMoney.C1���) = Format(curTmp, "0.00")
            End If
        Else
            vsfMoney.TextMatrix(lngȱʡRow, COLMoney.C1���) = "0.00"
        End If
        txtOwe.Text = "0.00"
    End If
    
    '���������(������-���ʽ��)
    '-----------------------------------------------------------------------------------------------------
    curMoney = GetPaySum
    If lng��� <> 0 Then
        'mcur����� = Format(vsfmoney.TextMatrix(lng���, COLMoney.C1���), mstrDec)
        vsfMoney.TextMatrix(lng���, COLMoney.C1���) = Format(Val(txtTotal.Text) - curMoney, mstrDec)
    Else
        mcur����� = Format(curMoney - Val(txtTotal.Text), mstrDec)
    End If
    
    '�п���Ӧ����������Ǵ���ֱҵ�����,�Ͳ���ʾ��
    If Val(txtOwe.Text) <> 0 And lngȱʡRow <> 0 And blnȱʡ�ֽ� And blnCent Then
        If Abs(Val(txtOwe.Text)) < 0.1 Or gBytMoney = 5 And Abs(Val(txtOwe.Text)) < 0.3 Then
            If CentMoney(Val(vsfMoney.TextMatrix(lngȱʡRow, COLMoney.C1���)) + Val(txtOwe.Text)) = Val(vsfMoney.TextMatrix(lngȱʡRow, COLMoney.C1���)) Then
                txtOwe.Text = "0.00"
            End If
        End If
    End If
    
    '����Ӧ��������С�������������,�����������С��1��,�Ͳ���ʾ��
    If Val(txtOwe.Text) <> 0 And mcur����� + curOwn = 0 And Abs(curOwn) <= 0.005 Then
        txtOwe.Text = "0.00"
    End If
    'txtOwe.ToolTipText = "�����:" & Format(mcur�����, mstrDec)
    
    curMoney = 0
    If mshDeposit.TextMatrix(1, COLDeposit.ID) <> "" Then
        For i = 1 To mshDeposit.Rows - 1
            curMoney = curMoney + Val(mshDeposit.TextMatrix(i, COLDeposit.��Ԥ��))
        Next
    End If
    lblDeposit.Caption = "��Ԥ��:" & Format(curMoney, "0.00")

    Call Calc�Ҳ�
    If gblnLED Then
        curTmp = GetӦ��
        zl9LedVoice.DisplayBank "�ܷ���" & Format(txtTotal.Text, "0.00"), "Ԥ����" & Format(lblSpare.Tag, "0.00"), _
                "��Ԥ��" & Format(curMoney, "0.00"), IIf(curTmp < 0, "�Ҳ�", "Ӧ��") & Format(Abs(curTmp), "0.00")
    End If
    
    '������ʾ
    '-----------------------------------------------------------------------------------------------------
    If strNone <> "" Then
        ShowMoney = "���ʳ��ϵı��ս��㷽ʽδ������ȫ,�ò��˻������±��ս��㷽ʽ���Ա�����" & _
            vbCrLf & strNone & vbCrLf & vbCrLf & "�����Ե����û�����Ŀ\���㷽ʽ������ȥ������Щ���㷽ʽ��"
    End If
End Function

Private Function GetPaySum() As Currency
'���ܣ���ȡ����ϼƣ�������Ԥ��������ĸ��ָ��ʽ���
    Dim i As Long, curMoney As Currency
    
    If mshDeposit.TextMatrix(1, COLDeposit.ID) <> "" Then
        For i = 1 To mshDeposit.Rows - 1
            curMoney = curMoney + Val(mshDeposit.TextMatrix(i, COLDeposit.��Ԥ��))
        Next
    End If
    
    For i = 1 To vsfMoney.Rows - 1
        If IsNumeric(vsfMoney.TextMatrix(i, COLMoney.C1���)) Then
            curMoney = curMoney + Val(vsfMoney.TextMatrix(i, COLMoney.C1���))
        End If
    Next

    GetPaySum = curMoney
End Function
Public Function Zl���˷�����Դ() As Byte
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ���˷���Դ��Ϣ
    '���أ�0-Ȩ����;1-��סԺ;2-�����סԺ(�ݲ����޴�����)
    '���ƣ����˺�
    '���ڣ�2010-03-09 17:39:26
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim BytType As Byte
    '��ȡ���û�ȡ��Χ����:'bytKind: 0-����ͨ����,1-��������,2-��ͨ���ú�������
    If mbytFunc = 0 Then BytType = 0
    If mbytFunc = 1 Then BytType = 1
    '���˺�:����ֻ�������סԺ����;���,ȡ�������ж�
'''    If mbytKind = 1 Then '��������
'''        BytType = 0
'''    ElseIf (InStr(mstrPrivs, "סԺ���ý���") = 0 Or mbytMCMode = 1) Then  '���ﲿ�ֵĴ���
'''            If InStr(mstrPrivs, "������ý���") = 0 Then
'''                '��Ȩ��,�ִ�������������ݵ�:
'''                ' a: 3-����(���￨�ȶ�����շ�);4-���
'''                BytType = IIf(mbytKind = 0, 1, 0) '����Ǿ��￨,�Ͷ�סԺ���ü�¼,�����������ü�¼
'''            Else
'''                '���������Ȩ��
'''                'a: 1-����,3-����(���￨�ȶ�����շ�);4-���
'''                BytType = IIf(mbytKind = 0, 2, 0)
'''            End If
'''    ElseIf InStr(mstrPrivs, "������ý���") = 0 Then    'סԺ����,�����ܽ��������
'''        '2-סԺ;3-����(���￨�ȶ�����շ�);4-���
'''        BytType = IIf(mbytKind = 0, 1, 2)
'''    Else  '�����סԺ
'''        '1-����;2-סԺ;3-����(���￨�ȶ�����շ�);4-���
'''        BytType = 2
'''    End If
    Zl���˷�����Դ = BytType
End Function
Private Function Is��������(ByVal lng����ID As Long, ByRef lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ�����Ƿ����������۲��˷����ڼ�
    '���:lng����ID
    '����:lng��ҳID-���ص�ǰ����ID(�ڼ������۵�)
    '����:
    '����:���˺�
    '����:2012-01-10 12:07:52
    '����:45302
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, dtStartDate As Date, dtEndDate As Date
    Dim strʱ�� As String, strCond As String, rsTemp As ADODB.Recordset
    strʱ�� = IIf(gint����ʱ�� = 0, "A.�Ǽ�ʱ��", "A.����ʱ��")
    strCond = "": dtStartDate = CDate("1901-01-01"): dtEndDate = dtStartDate
    If Not mDateBegin = CDate("0:00:00") Then
        strCond = " " & strʱ�� & " Between [3] And [4]"
        dtStartDate = CDate(Format(mDateBegin, "yyyy-MM-dd 00:00:00"))
        dtEndDate = CDate(Format(mDateEnd, "yyyy-MM-dd 23:59:59"))
    End If
    gstrSQL = "" & _
    "Select A.��ҳid " & _
    "   From ������ҳ A, " & _
    "        (Select Min(" & strʱ�� & ") As ��С����ʱ��, Max(" & strʱ�� & " ) ������ʱ�� " & _
    "          From ������ü�¼ A " & _
    "          Where  ����id = 728932 " & strCond & ") B " & _
    "   Where A.����id = 728932 And A.�������� = 1  " & _
    "       And (B.��С����ʱ�� Between A.��Ժ���� And Nvl(A.��Ժ����, Sysdate) Or " & _
    "                B.������ʱ�� Between A.��Ժ���� And Nvl(A.��Ժ����, Sysdate) Or " & _
    "                A.��Ժ���� Between B.��С����ʱ�� And B.������ʱ�� Or " & _
    "                Nvl(A.��Ժ����, Sysdate) Between B.��С����ʱ�� And B.������ʱ��)" & _
    "   Order by ��ҳID Desc"
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, dtStartDate, dtEndDate)
    If rsTemp.EOF Then rsTemp.Close: Set rsTemp = Nothing: Exit Function
    lng��ҳID = Val(Nvl(rsTemp!��ҳID))
    rsTemp.Close: Set rsTemp = Nothing
    Is�������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function SaveBalance(ByRef strNo As String, ByRef Curdate As Date, str����ԭ�� As String) As Long
'���ܣ��Ե�ǰ���ʵ����̴���
'���أ�����ID
    Dim arrSQL() As Variant
    Dim lng����ID As Long, i As Long, j As Long, lngTmp As Long, intInsure As Integer
    Dim str����IDs As String, str����ID As String, str���NO As String, strTmp As String
    Dim cur���ʽ��ϼ� As Currency, str���ս��� As String, str������Ϣ As String, strAdvance As String
    Dim blnҽ������У�� As Boolean, blnTrans As Boolean, blnTransMC As Boolean
    Dim cur�����ʻ� As Currency, curҽ������ As Currency, intMaxTime As Integer
    Dim cur�ɿ� As Currency, cur�Ҳ� As Currency, curԤ����� As Currency, cur��Ԥ�� As Currency, curԤ�����ϼ� As Currency, cur��Ԥ���ϼ� As Currency
    Dim lng��ҳID As Long
    Dim curOneCard As Currency, dblOneCardBalance As Double
    Dim strCardNo  As String, intCardType As Integer, strTransFlow As String
    Dim BytType As Byte, strסԺ���� As String
    
    Dim rsDeposit As ADODB.Recordset
    
    Screen.MousePointer = 11
    On Error GoTo Errhand:
    arrSQL = Array()
    strNo = zlDatabase.GetNextNo(15)
    lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    Curdate = zlDatabase.Currentdate
    intInsure = Nvl(mrsInfo!����, 0)
    If intInsure <> 0 Then str������Ϣ = Nvl(mrsInfo!����, " ") & "," & Nvl(mrsInfo!����, " ") & "," & Nvl(mrsInfo!ҽ����, " ")
    intMaxTime = GetMinMaxTime(1)
    cur�ɿ� = Val(txt�ɿ�.Text)
    cur�Ҳ� = Val(txt�Ҳ�.Text)
    
    '0-������;1-��סԺ;2-�����סԺ
    BytType = zlGetPatiSource
 
    '1.���˽��ʼ�¼
    '����:25596
    ' Zl_���˽��ʼ�¼_Insert
    strTmp = "zl_���˽��ʼ�¼_Insert("
    '  Id_In           ���˽��ʼ�¼.ID%Type,
    strTmp = strTmp & "" & lng����ID & ","
    '  ���ݺ�_In       ���˽��ʼ�¼.NO%Type,
    strTmp = strTmp & "'" & strNo & "',"
    '  ����id_In       ���˽��ʼ�¼.����id%Type,
    strTmp = strTmp & "" & Val(Nvl(mrsInfo!����ID)) & ","
    '  �շ�ʱ��_In     ���˽��ʼ�¼.�շ�ʱ��%Type,
    strTmp = strTmp & "To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  ��ʼ����_In     ���˽��ʼ�¼.��ʼ����%Type,
    strTmp = strTmp & "" & IIf(IsDate(txtPatiBegin.Text), "To_Date('" & txtPatiBegin.Text & "','YYYY-MM-DD')", "NULL") & ","
    '  ��������_In     ���˽��ʼ�¼.��������%Type,
    strTmp = strTmp & "" & IIf(IsDate(txtPatiEnd.Text), "To_Date('" & txtPatiEnd.Text & "','YYYY-MM-DD')", "NULL") & ","
    '  ��;����_In     ���˽��ʼ�¼.��;����%Type := 0,
    strTmp = strTmp & "" & IIf(opt��;.Value, 1, 0) & ","
    '  �ಡ�˽���_In   Number := 0,
    strTmp = strTmp & "" & 0 & ","
    '  �����ʴ���_In Number := 0,
    strTmp = strTmp & "" & intMaxTime & ","
    '  ��ע_In         ���˽��ʼ�¼.��ע%Type := Null
    strTmp = strTmp & "" & IIf(Trim(txt��ע.Text) = "", "NULL", "'" & Trim(txt��ע.Text) & "'") & ","
    '   ��Դ_In         Number := 1,1-����;2-סԺ
    strTmp = strTmp & "" & BytType & ","
    '  ԭ��_In         ���˽��ʼ�¼.ԭ��%Type := Null
    strTmp = strTmp & "" & IIf(Trim(str����ԭ��) = "", "NULL", "'" & Trim(str����ԭ��) & "'") & ","
    '    ��������_In     ���˽��ʼ�¼.��������%type:=2
    strTmp = strTmp & "" & IIf(mbytFunc = 0, 1, 2) & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strTmp: strTmp = ""
       
    '2.����Ԥ����¼-��Ԥ����[ID],[NO],����,���㷽ʽ,���,���
    With mshDeposit
        If .TextMatrix(1, COLDeposit.ID) <> "" Then
            '�ض�����Ԥ��,���������ж�
            Set rsDeposit = GetDeposit(mrsInfo!����ID, mblnDateMoved, IIf(gbln����ָ��Ԥ����, IIf(mstrTime = "", mstrAllTime, mstrTime), ""), , , mintԤ�����)
            For i = 1 To .Rows - 1
                curԤ����� = Val(.TextMatrix(i, COLDeposit.���))
                cur��Ԥ�� = Val(.TextMatrix(i, COLDeposit.��Ԥ��))
                If cur��Ԥ�� <> 0 Then
                    rsDeposit.Filter = "ID=" & CLng(.TextMatrix(i, COLDeposit.ID)) & " And NO='" & .TextMatrix(i, COLDeposit.���ݺ�) & "' And ��¼״̬=" & .RowData(i) & " And ���=" & curԤ�����
                    If rsDeposit.RecordCount = 0 Then
                        Call MsgBox("���ڲ�������,����Ԥ�����ѷ����仯,��������ȡ���˽���!", vbInformation, gstrSysName)
                        Screen.MousePointer = 0
                        Exit Function
                    End If
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "zl_����Ԥ����¼_Insert(" & CLng(.TextMatrix(i, COLDeposit.ID)) & "," & _
                        "'" & .TextMatrix(i, COLDeposit.���ݺ�) & "'," & .RowData(i) & "," & _
                        cur��Ԥ�� & "," & lng����ID & "," & mrsInfo!����ID & ")"
                    cur��Ԥ���ϼ� = cur��Ԥ���ϼ� + cur��Ԥ��
                End If
                curԤ�����ϼ� = curԤ�����ϼ� + curԤ�����
            Next
            '���ʳ����Ԥ��������Ԥ��������б����Ϻ�,����ָ���Ԥ������
            If cur��Ԥ���ϼ� > curԤ�����ϼ� And cur��Ԥ���ϼ� <> 0 Then
                Call MsgBox("����Ԥ����������!", vbInformation, gstrSysName)
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
    End With
    
    '3.����Ԥ����¼-���ʲ������㷽ʽ,���,�������
    With vsfMoney
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COLMoney.C1���)) <> 0 Then
                'ҽ���洢:�ɿλ=�������,��λ������=����,��λ�ʺ�=ҽ����
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                lng��ҳID = Val(Nvl(mrsInfo!��ҳID))
                If lng��ҳID = 0 Or mbytMCMode = 1 Or mbytFunc = 0 Then
                    '��������,��Ҫ������ҳID
                    '����:45302
                    If Nvl(mrsInfo!��������, 0) <> 1 And lng��ҳID <> 0 Then
                        '��ǰ���˲�������
                          If Not Is��������(mrsInfo!����ID, lng��ҳID) Then
                                lng��ҳID = 0
                          End If
                    End If
                End If
                
                arrSQL(UBound(arrSQL)) = _
                    "zl_���ʽɿ��¼_Insert('" & strNo & "'," & mrsInfo!����ID & "," & lng��ҳID & "," & _
                    IIf(IsNull(mrsInfo!��ǰ����id), 0, mrsInfo!��ǰ����id) & "," & _
                    "'" & .TextMatrix(i, COLMoney.C0����) & "','" & .TextMatrix(i, COLMoney.C2����) & "'," & _
                    CCur(.TextMatrix(i, COLMoney.C1���)) & "," & lng����ID & ",'" & UserInfo.��� & "'," & _
                    "'" & UserInfo.���� & "',To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    IIf(InStr(",3,4,", Val(.TextMatrix(i, COLMoney.C3����))) > 0, IIf(IsNull(mrsInfo!����), "NULL", mrsInfo!����), "NULL") & "," & _
                    IIf(InStr(",3,4,", Val(.TextMatrix(i, COLMoney.C3����))) > 0, "'" & IIf(IsNull(mrsInfo!ҽ����), "", mrsInfo!ҽ����) & "'", "NULL") & "," & _
                    IIf(InStr(",3,4,", Val(.TextMatrix(i, COLMoney.C3����))) > 0, "'" & IIf(IsNull(mrsInfo!����), "", mrsInfo!����) & "'", "NULL") & _
                    IIf(cur�ɿ� <> 0, "," & cur�ɿ� & "," & cur�Ҳ�, ",Null,Null") & ")"
                    
                    cur�ɿ� = 0
                If intInsure <> 0 And Not mblnNoInsure Then
                    '"���㷽ʽ|������||....."
                    If InStr(",3,4,", Val(.TextMatrix(i, COLMoney.C3����))) > 0 Then str���ս��� = str���ս��� & "||" & .TextMatrix(i, COLMoney.C0����) & "|" & Val(.TextMatrix(i, COLMoney.C1���))
                    If Val(.TextMatrix(i, COLMoney.C3����)) = 3 Then cur�����ʻ� = cur�����ʻ� + Val(.TextMatrix(i, COLMoney.C1���))
                    If Val(.TextMatrix(i, COLMoney.C3����)) = 4 Then curҽ������ = curҽ������ + Val(.TextMatrix(i, COLMoney.C1���))
                End If
                
                If mblnOneCard And Not mobjICCard Is Nothing Then
                    If .TextMatrix(i, COLMoney.C0����) = mrsOneCard!���㷽ʽ Then '�ڱ���֮ǰ���,ֻ��ʹ��һ��һ��ͨ���㷽ʽ
                        curOneCard = CCur(.TextMatrix(i, COLMoney.C1���))
                    End If
                End If
            End If
        Next
    End With
    If str���ս��� <> "" Then str���ս��� = Mid(str���ս���, 3)
    
    '4.סԺ���ü�¼��סԺ,�ڼ�,����,����,[���ݺ�],��Ŀ,��Ŀ,Ӥ����,[ID],[���],[��¼����],[��¼״̬],[ִ��״̬],[A.��ҳID],[A.��������ID],δ����,���ʽ��
    With mshDetail
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_���ʽ��)) <> 0 Or Val(.TextMatrix(i, COL_δ����)) = 0 Then
                'a.��ʣ����,���״ν��ʵ����ֽ�
                If Val(.TextMatrix(i, COL_ID)) = 0 Or Val(.TextMatrix(i, COL_δ����)) <> Val(.TextMatrix(i, COL_���ʽ��)) Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "zl_���ʷ��ü�¼_Insert(" & .TextMatrix(i, COL_ID) & "," & _
                        "'" & .TextMatrix(i, COL_���ݺ�) & "'," & .TextMatrix(i, COL_��¼����) & "," & _
                         .TextMatrix(i, COL_��¼״̬) & "," & Val(.TextMatrix(i, COL_ִ��״̬)) & "," & _
                         .TextMatrix(i, COL_���) & "," & CCur(.TextMatrix(i, COL_���ʽ��)) & "," & _
                         lng����ID & ")"
                Else
                'b.�״ν��ʲ���ȫ��
                    str����IDs = str����IDs & .TextMatrix(i, COL_ID) & ","
                End If
                If intInsure <> 0 And Not mblnNoInsure Then cur���ʽ��ϼ� = cur���ʽ��ϼ� + CCur(.TextMatrix(i, COL_���ʽ��))
            End If
        Next
                
        While str����IDs <> ""
            If Len(str����IDs) > 3998 Then
                lngTmp = InStrRev(Mid(str����IDs, 1, 3998), ",")
                str����ID = Mid(str����IDs, 1, lngTmp - 1)
                str����IDs = Mid(str����IDs, lngTmp + 1)
            Else
                str����ID = Mid(str����IDs, 1, Len(str����IDs) - 1)
                str����IDs = ""
            End If
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_���ʷ��ü�¼_Batch('" & str����ID & "'," & mrsInfo!����ID & "," & lng����ID & ")"
        Wend
    End With
    
    '5.��д��ʼƱ�ݺ�
    If mblnPrint And Trim(txtInvoice.Text) <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_Ʊ����ʼ��_Update('" & strNo & "','" & Trim(txtInvoice.Text) & "',3)"
    End If
        
    '���ִ��ǰ���������ж�
    '------------------------------------------------------------------------------
    '6.�����ʲ����ڼ�,���˷�������Ƿ����仯.
    If opt��Ժ.Value Then
        If mcurSpare <> Get�������(mrsInfo!����ID, 0, mintԤ�����) Then
        '���˺� ����:����:34244    ����:2010-11-19 15:06:09
        Call MsgBox("����Ҫ���ʵķ��������ʵ�ʵķ�����һ��!" & vbCrLf & _
        "�����ǽ��ʹ�����,�����˲�����Ϣ��,�����޸��˲��˷���!" & vbCrLf & _
        "�����ȷ������,ϵͳ��ǿ�����¶�ȡ���˷���!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName)
            If mDateBegin = CDate("0:00:00") Then
                txtPatient_KeyPress (13)  '������txt�������ֶ���������������,��ΪmrsInfo�Ǵ򿪵�,�����ض�������Ϣ
            Else
                Call ShowBalance
            End If
            Screen.MousePointer = 0
            Exit Function
        End If
    End If
    
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        blnTransMC = False
        If intInsure <> 0 And Not mblnNoInsure Then
            If mbytMCMode = 1 Then  '����ҽ������
                If cur�����ʻ� <> 0 Or curҽ������ <> 0 Or MCPAR.������봫����ϸ Then
                    If Not gclsInsure.ClinicSwap(lng����ID, cur�����ʻ�, curҽ������, 0, 0, intInsure, strAdvance) Then
                        gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
                    Else
                        blnTransMC = True
                    End If
                End If
            Else                    'סԺҽ������
                If Not gclsInsure.SettleSwap(lng����ID, intInsure, strAdvance) Then
                    gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
                Else
                    blnTransMC = True
                End If
            End If
        Else
            'һ��ͨ����
            If mblnOneCard And Not mobjICCard Is Nothing Then
                If curOneCard <> 0 Then
                    If Not mobjICCard.PaymentSwap(curOneCard, dblOneCardBalance, intCardType, Val("" & mrsOneCard!ҽԺ����), strCardNo, strTransFlow, lng����ID, mrsInfo!����ID) Then
                        gcnOracle.RollbackTrans
                        MsgBox "һ��ͨ����ʧ��", vbInformation, gstrSysName
                        Exit Function
                    Else
                        gstrSQL = "zl_һ��ͨ����_Update(" & lng����ID & ",'" & mrsOneCard!���㷽ʽ & "','" & strCardNo & "','" & intCardType & "','" & strTransFlow & "'," & dblOneCardBalance & ")"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    End If
                End If
            End If
        End If
        '���˺�;
        If zlSequareBlance(lng����ID) = False Then
            gcnOracle.RollbackTrans
            MsgBox "���ѿ�����ʧ��!", vbInformation, gstrSysName
            Exit Function
        End If
        
    gcnOracle.CommitTrans: blnTrans = False
    If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(mbytMCMode = 1, ����Enum.Busi_ClinicSwap, ����Enum.Busi_SettleSwap), True, intInsure)
    Screen.MousePointer = 0
    
    'ҽ������У��
    If strAdvance <> "" And str���ս��� <> strAdvance And Not mblnNoInsure Then
        blnҽ������У�� = True
        If UBound(Split(str���ս���, "||")) = UBound(Split(strAdvance, "||")) Then
            For i = 0 To UBound(Split(str���ս���, "||"))
                blnҽ������У�� = True
                strTmp = Split(str���ս���, "||")(i)
                For j = 0 To UBound(Split(strAdvance, "||"))
                    If Split(strTmp, "|")(0) = Split(Split(strAdvance, "||")(j), "|")(0) Then
                        If Val(Split(strTmp, "|")(1)) = Val(Split(Split(strAdvance, "||")(j), "|")(1)) Then
                            blnҽ������У�� = False
                        End If
                    End If
                Next
                If blnҽ������У�� Then Exit For
            Next
        End If
        '��ʽ����ǰ��,���㷽ʽ�ͽ�����δ�����仯ʱ��У��
        If blnҽ������У�� Then
            cur�ɿ� = Val(txt�ɿ�.Text)
            strסԺ���� = ""
            If mbytFunc <> 0 Then
                strסԺ���� = IIf(gbln����ָ��Ԥ���� And mbln����תסԺ = False, IIf(mstrTime = "", mstrAllTime, mstrTime), "")
            End If

            blnҽ������У�� = frmMedicareReckoning.ShowMe(Me, _
                lng����ID, mrsInfo!����ID, opt��;.Value, cur���ʽ��ϼ�, strAdvance, str������Ϣ, _
                intInsure, mstrDec, gBytMoney, cur�ɿ�, "" & mrsInfo!ҽ����, mbytMCMode, strסԺ����, mintԤ�����)
                                    
            If Not blnҽ������У�� Then
                MsgBox "����[" & strNo & "]����ҽ������У��ʧ��,���ʽ����ܲ���ȷ!" & _
                    vbCrLf & vbCrLf & "������ӡƱ��,�뵽[���ս������]������У�Ժ��ٴ�ӡ!", vbInformation, gstrSysName
                mblnPrint = False
            End If
        End If
    End If
    
    '���뵥����ʷ��¼(�������͵���)
    strTmp = strNo
    For i = 0 To cboNO.ListCount - 1
        strTmp = strTmp & "," & cboNO.List(i)
    Next
    cboNO.Clear
    For i = 0 To UBound(Split(strTmp, ","))
        cboNO.AddItem Split(strTmp, ",")(i)
        If i = 9 Then Exit For 'ֻ��ʾ10��
    Next
    
    Set mtySquareCard.rsSquare = Nothing
    SaveBalance = lng����ID
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    
    If blnTrans Then
        'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
        If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(mbytMCMode = 1, ����Enum.Busi_ClinicSwap, ����Enum.Busi_SettleSwap), False, intInsure)
    End If
    
    Screen.MousePointer = 0
    Call SaveErrLog
    Exit Function
Errhand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Screen.MousePointer = 99
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function ExecuteSquareUpdate(ByVal rsSquare As ADODB.Recordset, ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ���
    '���:rsSquare-ˢ����������
    '����:
    '����:
    '����:���˺�
    '����:2010-01-09 22:47:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strTemp As String
    
     With rsSquare
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            'Zl_���˿������¼_Insert
            strSql = "Zl_���˿������¼_Insert("
            '  �ӿڱ��_In   In ���˿������¼.�ӿڱ��%Type,
            strSql = strSql & "" & Val(Nvl(!�ӿڱ��)) & ","
            '  ���ѿ�id_In   In ���˿������¼.���ѿ�id%Type,
            strSql = strSql & "" & IIf(Val(Nvl(!���ѿ�ID)) = 0, "NULL", Val(Nvl(!���ѿ�ID))) & ","
            '  ���㷽ʽ_In   In ���˿������¼.���㷽ʽ%Type,
            strSql = strSql & "'" & Trim(Nvl(!���㷽ʽ)) & "',"
            '  ������_In   In ���˿������¼.������%Type,
            strSql = strSql & "" & Val(Nvl(!������)) & ","
            '  ����_In       In ���˿������¼.����%Type,
            strSql = strSql & "'" & Trim(Nvl(!����)) & "',"
            '  ������ˮ��_In In ���˿������¼.������ˮ��%Type,
            
            strSql = strSql & "'" & Trim(Nvl(!������ˮ��)) & "',"
            '  ����ʱ��_In   In ���˿������¼.����ʱ��%Type,
            strTemp = Format(!����ʱ��, "yyyy-mm-dd HH:MM:SS")
            If strTemp = "" Then
                strSql = strSql & "NULL,"
            Else
                strSql = strSql & "to_date('" & strTemp & "','yyyy-mm-dd hh24:mi:ss'),"
            End If
            '  ��ע_In       In ���˿������¼.��ע%Type,
            strSql = strSql & "'" & Trim(Nvl(!��ע)) & "',"
            '  ����id_In     In Varchar2
            strSql = strSql & "'" & lng����ID & "')"
            
            zlDatabase.ExecuteProcedure strSql, Me.Caption
            .MoveNext
        Loop
     End With
    ExecuteSquareUpdate = True
End Function

Private Function zlSequareBlance(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ѿ�����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsSquare As ADODB.Recordset
    If mbytInState <> 0 Then GoTo goEnd:

    '���˺�:
    If Not mtySquareCard.blnExistsObjects Then GoTo goEnd:
    If gobjSquare.objSquareCard Is Nothing Then GoTo goEnd:
    If mtySquareCard.rsSquare Is Nothing Then GoTo goEnd:
    If mtySquareCard.rsSquare.State <> 1 Then GoTo goEnd:
    If mtySquareCard.rsSquare.RecordCount = 0 Then GoTo goEnd:

    Set rsSquare = zlDatabase.CopyNewRec(mtySquareCard.rsSquare)
    If rsSquare Is Nothing Then GoTo goEnd:
    If rsSquare.State <> 1 Then GoTo goEnd:
    If ExecuteSquareUpdate(rsSquare, lng����ID) = False Then Exit Function

    '������Ӧ�Ľ���ӿ�
    '���ýӿ�
    'Public Function zlSquareFee(frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, ByVal str����ID_IN As String, ByVal rsSquare As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: zlSquareFee (����ӿ�)
    '���:frmMain:HIS���� ���õ�������
    '     intCallType : HIS���� 0-  ������õ��� 1-  סԺ���ʵ���
    '     str����ID_IN: HIS���� ���ν��ʵĽ���ID��
    '     rsSquare :  ����Ӧˢ���Ľ���
    '����:
    '����:true:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:18:38
    '˵��:
    '    1. ��"�����շ�"�����"ȷ��"ʱ,���ñ��ӿ�
    '    2. ��"סԺ����"�����"ȷ��"ʱ,���ñ��ӿ�
    'ע:
    '  �˽ӿ���������HIS������ , ��˲����ڴ˽ӿڴ������û������Ĳ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
     If gobjSquare.objSquareCard.zlSquareFee(Me, mlngModul, mstrPrivs, lng����ID, mtySquareCard.rsSquare) = False Then
          Exit Function
     End If
goEnd:
    zlSequareBlance = True
    Exit Function
End Function

Private Function LoadCardData() As Boolean
'���ܣ����ݵ�ǰѡ��Ĳ��˷�����Ŀ��Ƭ����ȡ�����÷����嵥
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long, j As Long
    Dim strInfo As String, strPre As String
    Dim strMoney As String, strTmp As String, strTmpSql As String
    Dim arrTotal() As Currency
    Dim strCond As String, BytType As Byte '0-����;1-סԺ;2-�����סԺ
    Dim DateBegin As Date, DateEnd As Date
    Dim strTable As String
    
    On Error GoTo errH
    
    If mbytInState = 0 And mrsInfo.State = 0 Then Exit Function
    
    strPre = sta.Panels(2).Text
    sta.Panels(2).Text = "���ڶ�ȡ����,���Ժ� ����"
    Screen.MousePointer = 11
    mshQuery.Redraw = False
    Me.Refresh
    
    If mbytInState = 0 Then
        strCond = ""
        strCond = strCond & IIf(mstrTime = "", "", " And Instr([2],','||Nvl(A.��ҳID,0)||',')>0")
        If mDateBegin <> CDate("0:00:00") Then
            strCond = strCond & " And " & IIf(gint����ʱ�� = 0, "A.�Ǽ�ʱ��", "A.����ʱ��") & " Between [3] And [4]"
            DateBegin = CDate(Format(mDateBegin, "yyyy-MM-dd 00:00:00"))
            DateEnd = CDate(Format(mDateEnd, "yyyy-MM-dd 23:59:59"))
        End If
        strCond = strCond & IIf(mstrUnit = "", "", " And Instr([5],','||A.��������ID||',')>0")
        strCond = strCond & IIf(mbytBaby = 0, "", IIf(mbytBaby = 1, " And Nvl(A.Ӥ����,0)=0", " And A.Ӥ����=[6]"))
        strCond = strCond & IIf(mstrItem = "", "", " And Instr([7],','''||A.�վݷ�Ŀ||''',')>0")
        
        If mbytKind = 1 Then
            strCond = strCond & " And A.�����־=4"
        Else
            If InStr(mstrPrivs, ";סԺ���ý���;") = 0 Or mbytMCMode = 1 Then strCond = strCond & " And A.�����־<>2"
            If InStr(mstrPrivs, ";������ý���;") = 0 Then strCond = strCond & " And A.�����־<>1"
            If mbytKind = 0 Then strCond = strCond & " And A.�����־<>4"
        End If
        
        BytType = Zl���˷�����Դ
    
        '���ü�¼״̬,ֻȡ��δ����ĵ���(δ��ϸ�����,Ҫ��ʾ�����˷���)
        If Not mnuFileZero.Checked Then
            strSql = _
            " Select NO,Mod(��¼����,10) as ��¼����, Nvl(Sum(ʵ�ս��),0) as ʵ�ս��,Nvl(Sum(���ʽ��),0) as ���ʽ��" & _
            " From סԺ���ü�¼ A" & _
            " Where ��¼״̬<>0 And ���ʷ���=1 " & strCond & _
            "       And ����ID=[1]" & _
            " Group by NO,Mod(��¼����,10) " & _
            " Having Nvl(Sum(ʵ�ս��),0)-Nvl(Sum(���ʽ��),0)<>0"
            
            strSql = _
                " Select Mod(A.��¼����,10) as ��¼����,A.����ʱ��,A.�Ǽ�ʱ��,A.NO,A.�շ����,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��������ID,A.���㵥λ," & _
                "        A.����,Nvl(A.����,1) as ����,A.��׼����,Sum(A.ʵ�ս��) As ʵ�ս��,Sum(A.���ʽ��) As ���ʽ��,A.����Ա����,A.��������" & _
                " From סԺ���ü�¼ A,(" & strSql & ") B" & _
                " Where A.NO=B.NO And Mod(A.��¼����,10)=B.��¼����" & _
                "       And A.��¼״̬<>0 And A.���ʷ���=1 And Nvl(A.ʵ�ս��,0)<>Nvl(A.���ʽ��,0)" & _
                "       And A.����ID+0=[1] " & strCond & _
                " Having Nvl(Sum(A.ʵ�ս��),0)-Nvl(Sum(A.���ʽ��),0)<>0" & _
                " Group by Mod(A.��¼����,10),A.����ʱ��,A.�Ǽ�ʱ��,A.NO,A.�շ����,Nvl(A.�۸񸸺�,A.���),A.�շ�ϸĿID," & _
                "       A.�վݷ�Ŀ,A.��������ID,A.���㵥λ,A.����,Nvl(A.����,1),A.��׼����,A.����Ա����,A.�������� "
            
            If mblnDateMoved Then
                strSql = strSql & " Union All " & Replace(strSql, "סԺ���ü�¼", "HסԺ���ü�¼")
            End If
        Else
            strSql = _
                " Select Mod(A.��¼����,10) as ��¼����,A.����ʱ��,A.�Ǽ�ʱ��,A.NO,A.�շ����,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��������ID,A.���㵥λ," & _
                "       A.����,Nvl(A.����,1) as ����,A.��׼����,Sum(A.ʵ�ս��) As ʵ�ս��,Sum(A.���ʽ��) As ���ʽ��,A.����Ա����,A.��������" & _
                " From " & IIf(mblnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & "" & _
                " Where A.��¼״̬<>0 And A.���ʷ���=1  And A.����ID=[1] " & strCond & _
                "       And (Nvl(A.ʵ�ս��,0)<>Nvl(A.���ʽ��,0) Or Nvl(A.ʵ�ս��,0)=0 And A.����ID is Null)" & _
                " Having Nvl(Sum(A.ʵ�ս��),0)-Nvl(Sum(A.���ʽ��),0)<>0 Or Sum(Nvl(A.ʵ�ս��,0))=0 And Sum(A.���ʽ��) is Null" & _
               "  Group by Mod(A.��¼����,10),A.����ʱ��,A.�Ǽ�ʱ��,A.NO,A.�շ����,Nvl(A.�۸񸸺�,A.���),A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��������ID,A.���㵥λ,A.����,Nvl(A.����,1),A.��׼����,A.����Ա����,A.�������� "
        End If
        
        Select Case BytType
        Case 0 '����
            strSql = Replace(Replace(strSql, "סԺ���ü�¼", "������ü�¼"), " And Instr([2],','||Nvl(A.��ҳID,0)||',')>0", "")
            If Not mnuFileZero.Checked Then
                strTmpSql = _
                " Select NO,Mod(��¼����,10) as ��¼����, Nvl(Sum(ʵ�ս��),0) as ʵ�ս��,Nvl(Sum(���ʽ��),0) as ���ʽ��" & _
                " From סԺ���ü�¼ A" & _
                " Where ��¼״̬<>0 And ���ʷ���=1 And Mod(��¼����,10)=5 And ��ҳID Is Null " & strCond & _
                "       And ����ID=[1]" & _
                " Group by NO,Mod(��¼����,10) " & _
                " Having Nvl(Sum(ʵ�ս��),0)-Nvl(Sum(���ʽ��),0)<>0"
                
                strTmpSql = _
                " Select Mod(A.��¼����,10) as ��¼����,A.����ʱ��,A.�Ǽ�ʱ��,A.NO,A.�շ����,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��������ID,A.���㵥λ," & _
                "        A.����,Nvl(A.����,1) as ����,A.��׼����,Sum(A.ʵ�ս��) As ʵ�ս��,Sum(A.���ʽ��) As ���ʽ��,A.����Ա����,A.��������" & _
                " From סԺ���ü�¼ A,(" & strTmpSql & ") B" & _
                " Where A.NO=B.NO And Mod(A.��¼����,10)=B.��¼����" & _
                "       And A.��¼״̬<>0 And A.���ʷ���=1 And Mod(A.��¼����,10)=5 And A.��ҳID Is Null And Nvl(A.ʵ�ս��,0)<>Nvl(A.���ʽ��,0)" & _
                "       And A.����ID+0=[1] " & strCond & _
                " Having Nvl(Sum(A.ʵ�ս��),0)-Nvl(Sum(A.���ʽ��),0)<>0" & _
                " Group by Mod(A.��¼����,10),A.����ʱ��,A.�Ǽ�ʱ��,A.NO,A.�շ����,Nvl(A.�۸񸸺�,A.���),A.�շ�ϸĿID," & _
                "       A.�վݷ�Ŀ,A.��������ID,A.���㵥λ,A.����,Nvl(A.����,1),A.��׼����,A.����Ա����,A.�������� "
                If mblnDateMoved Then
                    strTmpSql = strTmpSql & " Union All " & Replace(strTmpSql, "סԺ���ü�¼", "HסԺ���ü�¼")
                End If
            Else
                strTmpSql = _
                " Select Mod(A.��¼����,10) as ��¼����,A.����ʱ��,A.�Ǽ�ʱ��,A.NO,A.�շ����,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��������ID,A.���㵥λ," & _
                "       A.����,Nvl(A.����,1) as ����,A.��׼����,Sum(A.ʵ�ս��) As ʵ�ս��,Sum(A.���ʽ��) As ���ʽ��,A.����Ա����,A.��������" & _
                " From " & IIf(mblnDateMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & "" & _
                " Where A.��¼״̬<>0 And A.���ʷ���=1 And  Mod(A.��¼����,10)=5 And A.��ҳID Is Null And A.����ID=[1] " & strCond & _
                "       And (Nvl(A.ʵ�ս��,0)<>Nvl(A.���ʽ��,0) Or Nvl(A.ʵ�ս��,0)=0 And A.����ID is Null)" & _
                " Having Nvl(Sum(A.ʵ�ս��),0)-Nvl(Sum(A.���ʽ��),0)<>0 Or Sum(Nvl(A.ʵ�ս��,0))=0 And Sum(A.���ʽ��) is Null" & _
               "  Group by Mod(A.��¼����,10),A.����ʱ��,A.�Ǽ�ʱ��,A.NO,A.�շ����,Nvl(A.�۸񸸺�,A.���),A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��������ID,A.���㵥λ,A.����,Nvl(A.����,1),A.��׼����,A.����Ա����,A.�������� "
            End If
            strTmpSql = Replace(strTmpSql, " And Instr([2],','||Nvl(A.��ҳID,0)||',')>0", "")
            strSql = strSql & " Union All " & strTmpSql
        Case 1 'סԺ
        Case Else
            '�����סԺ
             strSql = strSql & " Union All " & Replace(Replace(strSql, "סԺ���ü�¼", "������ü�¼"), " And Instr([2],','||Nvl(A.��ҳID,0)||',')>0", "")
        End Select
        strTable = "(" & strSql & ") "
        
            
        'δ������嵥
        Select Case tabCard.SelectedItem.Index
            Case 2 '��ϸ�嵥
                strSql = _
                " SELECT To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�," & _
                "       B.���� as ����,Nvl(D.����,C.����) as ��Ŀ,C.���,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')||A.����||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(Nvl(A.��׼����,0),'999999999" & gstrFeePrecisionFmt & "')) as ��׼����," & _
                "       Ltrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),4),'999999999" & mstrDec & "')) as ��׼���," & _
                "       Ltrim(To_Char(Nvl(A.ʵ�ս��,0)-Nvl(A.���ʽ��,0),'999999999" & mstrDec & "')) as δ����,A.����Ա���� as ����Ա" & _
                " FROM " & strTable & " A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID " & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Order by ��������,���ݺ�,��Ŀ"
                strMoney = "4,4,1,1,1,1,1,7,7,7,1"
            Case 3 '����Ŀ��ϸ
                strSql = _
                " SELECT To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�," & _
                "       B.���� as ��������,Nvl(D.����,C.����) as ��Ŀ,Nvl(C.���,' ') ���,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')||A.����||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(Nvl(A.��׼����,0),'999999999" & gstrFeePrecisionFmt & "')) as ��׼����," & _
                "       Ltrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),4),'999999999" & mstrDec & "')) as ��׼���," & _
                "       Ltrim(To_Char(Nvl(A.ʵ�ս��,0)-Nvl(A.���ʽ��,0),'999999999" & mstrDec & "')) as δ����," & _
                "       Nvl(A.��������,C.��������) as ����,A.����Ա���� as ����Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��" & _
                " FROM " & strTable & " A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1)
                
               strSql = strSql & _
                " Union All" & _
                " SELECT NULL as ��������,NULL as ���ݺ�,NULL as ��������," & _
                "       Nvl(D.����,C.����) as ��Ŀ,Nvl(C.���,' ')||'ZZZZZ' as ���,NULL,to_char(sum(Nvl(A.����,1)*Nvl(A.����,1)))||' '||A.���㵥λ as ����,NULL as ��׼����," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "       Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as δ����," & _
                "       NULL as ����,NULL as ����Ա,NULL as �Ǽ�ʱ��" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.�շ�ϸĿID=C.ID And A.�շ�ϸĿID=D.�շ�ϸĿID(+)" & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                "              And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Group by Nvl(D.����,C.����),C.���,A.���㵥λ" & _
                " Order by ��Ŀ,���,��������,���ݺ�"
                
                strMoney = "4,4,1,1,1,1,1,7,7,7,1,1,1"
            Case 4 '������ϸ
                strSql = _
                " SELECT To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�," & _
                "       B.���� as ����,Nvl(D.����,C.����) as ��Ŀ,C.���,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')||A.����||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(Nvl(A.��׼����,0),'999999999" & gstrFeePrecisionFmt & "')) as ��׼����," & _
                "       Ltrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),4),'999999999" & mstrDec & "')) as ��׼���," & _
                "       Ltrim(To_Char(Nvl(A.ʵ�ս��,0)-Nvl(A.���ʽ��,0),'999999999" & mstrDec & "')) as δ����,A.����Ա���� as ����Ա " & _
                " FROM " & strTable & " A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID " & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Union All" & _
                " SELECT NULL as ��������,NULL as ���ݺ�,NULL as ����,NULL as ��Ŀ,Null as ���,A.�վݷ�Ŀ||'ZZZZZ' as ��Ŀ," & _
                "        NULL as ����,NULL as ��׼����," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as δ����,NULL as ����Ա" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Group by A.�վݷ�Ŀ||'ZZZZZ'" & _
                " Order by ��Ŀ,��������,���ݺ�"
                strMoney = "4,4,1,1,1,1,1,7,7,7,1"
            Case 5 '�����嵥
                strSql = _
                " SELECT B.�ڼ�,A.�վݷ�Ŀ as ��Ŀ," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as δ����" & _
                "        FROM " & strTable & " A,�ڼ�� B,�շ���ĿĿ¼ C" & _
                " Where A.�Ǽ�ʱ�� Between Trunc(B.��ʼ����) and Trunc(B.��ֹ����)+1-1/24/60/60 " & _
                "       And A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Group by B.�ڼ�,A.�վݷ�Ŀ" & _
                " Union All" & _
                " SELECT B.�ڼ�||'ZZZZZ',NULL as ��Ŀ," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as δ����" & _
                " FROM " & strTable & " A,�ڼ�� B,�շ���ĿĿ¼ C" & _
                " Where A.�Ǽ�ʱ�� Between Trunc(B.��ʼ����) and Trunc(B.��ֹ����)+1-1/24/60/60 " & _
                "       And A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Group by B.�ڼ�||'ZZZZZ'" & _
                " Order by �ڼ�,��Ŀ"
                strMoney = "4,4,7,7"
                
            Case 6 '�����嵥
                strSql = _
                " SELECT A.�վݷ�Ŀ as ��Ŀ," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "       Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as δ����" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Group by A.�վݷ�Ŀ Order by ��Ŀ"
                strMoney = "4,7,7"
            Case 7 '���շ���
                strSql = _
                " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�,A.�վݷ�Ŀ as ������Ŀ," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as δ����," & _
                "        A.����Ա���� as ����Ա,A.��¼����" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Group by A.��¼����,TO_Char(A.����ʱ��,'YYYY-MM-DD'),A.NO,A.�վݷ�Ŀ,A.����Ա����"
                strSql = strSql & _
                " Union All" & _
                " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO||'ZZZZZ' as ���ݺ�,NULL as ������Ŀ," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as δ����,NULL as ����Ա,A.��¼����" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Having Nvl(Sum(A.ʵ�ս��), 0) - Nvl(Sum(A.���ʽ��), 0) <> 0" & _
                " Group by A.��¼����,TO_Char(A.����ʱ��,'YYYY-MM-DD'),A.NO" & _
                " Union All" & _
                " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD')||'ZZZZZ' as ��������,NULL as ���ݺ�,NULL as ������Ŀ," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as δ����,NULL as ����Ա,-1" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Having Nvl(Sum(A.ʵ�ս��), 0) - Nvl(Sum(A.���ʽ��), 0) <> 0" & _
                " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD')" & _
                " Order by ��������,��¼���� desc,���ݺ�,������Ŀ"
                
                strMoney = "4,4,4,7,7,1"
            Case 8 '���շ�Ŀ
                strSql = _
                " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.�վݷ�Ŀ as ������Ŀ," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as δ����" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD'),A.�վݷ�Ŀ" & _
                " Union All" & _
                " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD')||'ZZZZZ' as ��������,NULL as ������Ŀ," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as δ����" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'��'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Having Nvl(Sum(A.ʵ�ս��), 0) - Nvl(Sum(A.���ʽ��), 0) <> 0" & _
                " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD')" & _
                " Order by ��������,������Ŀ"
                strMoney = "4,4,7,7"
        End Select
                
        mshQuery.MergeCells = flexMergeFree
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CLng(mrsInfo!����ID), "," & mstrTime & ",", DateBegin, DateEnd, _
                    "," & mstrUnit & ",", mbytBaby - 1, "," & mstrItem & ",", "," & mstrClass & ",", "," & mstrChargeType & ",")
        If rsTmp.RecordCount > 0 Then
            Set mshQuery.DataSource = rsTmp
        Else
            Call BandRectoGrid(mshQuery, rsTmp)
        End If
        
        
        mshQuery.Tag = tabCard.SelectedItem.Index
        For i = 0 To mshQuery.Cols - 1
            mshQuery.MergeCol(i) = False
        Next
        
        '��ϼ�(С��)
        Select Case tabCard.SelectedItem.Index
            Case 2, 4  '��ϸ�嵥��������ϸ
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 5) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 8))
                            If IsNumeric(mshQuery.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 9))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.Row = i
                            mshQuery.MergeRow(i) = True
                            strTmp = mshQuery.TextMatrix(i, 5)
                            For j = 0 To 7
                                mshQuery.Col = j: mshQuery.CellAlignment = 4
                                mshQuery.TextMatrix(i, j) = "С ��:" & Left(strTmp, Len(strTmp) - 5)
                            Next
                            For j = 8 To mshQuery.Cols - 2
                                mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 7
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "�� ��"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 8) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 9) = Format(arrTotal(1), " " & mstrDec)
                End If
            Case 3 '����Ŀ��ϸ
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 4) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 8))
                            If IsNumeric(mshQuery.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 9))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.Row = i
                            mshQuery.MergeRow(i) = True
                            strTmp = mshQuery.TextMatrix(i, 3)
                            For j = 0 To 5
                                mshQuery.Col = j: mshQuery.CellAlignment = 4
                                mshQuery.TextMatrix(i, j) = "С ��:" & strTmp
                            Next
                            mshQuery.TextMatrix(i, 7) = " " '������
                            For j = 8 To mshQuery.Cols - 2
                                mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 7
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "�� ��"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 8) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 9) = Format(arrTotal(1), " " & mstrDec)
                End If
            Case 5 '�����嵥
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 2))
                            If IsNumeric(mshQuery.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 3))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.Row = i
                            mshQuery.MergeRow(i) = True
                            For j = 0 To 1
                                mshQuery.Col = j: mshQuery.CellAlignment = 4
                                mshQuery.TextMatrix(i, j) = "С��:" & mshQuery.TextMatrix(i - 1, 0)
                            Next
                            For j = 2 To mshQuery.Cols - 1
                                mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 1
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "�� ��"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 2) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 3) = Format(arrTotal(1), " " & mstrDec)
                End If
            Case 6 '�����嵥
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If IsNumeric(mshQuery.TextMatrix(i, 1)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 1))
                        If IsNumeric(mshQuery.TextMatrix(i, 2)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 2))
                        mshQuery.MergeRow(i) = False
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.Col = 0: mshQuery.CellAlignment = 4
                    mshQuery.TextMatrix(mshQuery.Row, 0) = "�� ��"
                    mshQuery.TextMatrix(mshQuery.Row, 1) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 2) = Format(arrTotal(1), " " & mstrDec)
                End If
            Case 7 '���յ���
                If rsTmp.RecordCount > 0 Then
                    mshQuery.MergeCol(0) = True
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 1) Like "*ZZZZZ") And Not (mshQuery.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 3)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 3))
                            If IsNumeric(mshQuery.TextMatrix(i, 4)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 4))
                            mshQuery.MergeRow(i) = False
                        Else
                            If mshQuery.TextMatrix(i, 1) Like "*ZZZZZ" Then
                                mshQuery.Row = i
                                mshQuery.MergeRow(i) = True
                                For j = 1 To 2
                                    mshQuery.Col = j: mshQuery.CellAlignment = 4
                                    mshQuery.TextMatrix(i, j) = "С��:" & mshQuery.TextMatrix(i - 1, 1)
                                Next
                                For j = 3 To mshQuery.Cols - 2
                                    mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                                Next
                            Else
                                mshQuery.Row = i
                                mshQuery.MergeRow(i) = True
                                For j = 0 To 2
                                    mshQuery.Col = j: mshQuery.CellAlignment = 4
                                    mshQuery.TextMatrix(i, j) = "С��:" & mshQuery.TextMatrix(i - 1, 0)
                                Next
                                For j = 3 To mshQuery.Cols - 2
                                    mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                                Next
                            End If
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 2
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "�� ��"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 3) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 4) = Format(arrTotal(1), " " & mstrDec)
                    
                    'ɾ��ֻ��һ�е��ݵ�С����
                    j = 0
                    For i = 1 To mshQuery.Rows - 1
                        If mshQuery.TextMatrix(i, 1) Like "*С��*" Then
                            If j = 1 Then mshQuery.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                End If
            Case 8 '���շ�Ŀ
                If rsTmp.RecordCount > 0 Then
                    mshQuery.MergeCol(0) = True
                    ReDim arrTotal(1)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 2))
                            If IsNumeric(mshQuery.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 3))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.MergeRow(i) = True
                            mshQuery.Row = i
                            mshQuery.Col = 1: mshQuery.CellAlignment = 4
                            mshQuery.TextMatrix(i, 0) = "С��:" & mshQuery.TextMatrix(i - 1, 0)
                            mshQuery.TextMatrix(i, 1) = mshQuery.TextMatrix(i, 0)
                            For j = 2 To mshQuery.Cols - 2
                                mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 1
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "�� ��"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 2) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 3) = Format(arrTotal(1), " " & mstrDec)
                
                    'ɾ��ֻ��һ�з��õ�С����
                    j = 0
                    For i = 1 To mshQuery.Rows - 1
                        If mshQuery.TextMatrix(i, 1) Like "*С��*" Then
                            If j = 1 Then mshQuery.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                End If
        End Select
    Else
        strSql = "Select ����ʱ��,�Ǽ�ʱ��,NO,�վݷ�Ŀ,��������,����,����,���㵥λ,��׼����,���ʽ��,����Ա����,��������ID,�շ�ϸĿID,����ID From סԺ���ü�¼  where ����ID= [1]  Union ALL " & _
                 "Select ����ʱ��,�Ǽ�ʱ��,NO,�վݷ�Ŀ,��������,����,����,���㵥λ,��׼����,���ʽ��,����Ա����,��������ID,�շ�ϸĿID,����ID From ������ü�¼  where ����ID= [1]"
        
        If mblnNOMoved Then
            strSql = Replace(Replace(strSql, "סԺ���ü�¼", "HסԺ���ü�¼"), "������ü�¼", "H������ü�¼")
        End If
        strSql = "(" & strSql & ")"
        
        '��ȡ���ʵ�ʱ,����ʷ�����ϸ
        Select Case tabCard.SelectedItem.Index
            Case 2 '��ϸ
                '��������,���ݺ�,����,��Ŀ,��Ŀ,����,����,��׼���,���ʽ��,����Ա
                strSql = _
                " Select To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�," & _
                "       Nvl(B.����,'δ֪') as ����,Nvl(D.����,C.����) as ��Ŀ,C.���,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')||A.����||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(A.��׼����,'99999" & gstrFeePrecisionFmt & "')) as ����," & _
                "       Ltrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),4),'999999999" & mstrDec & "')) as ��׼���," & _
                "       Ltrim(To_Char(A.���ʽ��,'999999999" & mstrDec & "')) as ���ʽ��,A.����Ա���� as ����Ա" & _
                " From " & strSql & " A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.��������ID = B.ID(+) And A.�շ�ϸĿID=C.ID" & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Order by ��������,���ݺ�,��Ŀ"
                
                '�嵥��ʽ����
               strMoney = "4,4,1,1,1,4,1,7,7,7,1"
            Case 3 '����Ŀ��ϸ
                '��������,���ݺ�,����,��Ŀ,���,��Ŀ,����,����,��׼���,���ʽ��,����,����Ա
                strSql = _
                " SELECT To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�," & _
                "       B.���� as ��������,Nvl(D.����,C.����) as ��Ŀ,Nvl(C.���,' ') as ���,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')||A.����||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(Nvl(A.��׼����,0),'999999999" & gstrFeePrecisionFmt & "')) as ��׼����," & _
                "       Ltrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),4),'999999999" & mstrDec & "')) as ��׼���," & _
                "       Ltrim(To_Char(Nvl(A.���ʽ��,0),'999999999" & mstrDec & "')) as ���ʽ��," & _
                "       Nvl(A.��������,C.��������) as ����,A.����Ա���� as ����Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��" & _
                " FROM " & strSql & " A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID" & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Union All" & _
                " SELECT NULL as ��������,NULL as ���ݺ�,NULL as ��������,Nvl(D.����,C.����) as ��Ŀ,Nvl(C.���,' ')||'ZZZZZ' as ���," & _
                "        NULL as ��Ŀ,to_char(sum(Nvl(A.����,1)*Nvl(A.����,1)))||' '||A.���㵥λ as ����,NULL as ��׼����," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "        Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as ���ʽ��,NULL as ����,NULL as ����Ա,NULL as �Ǽ�ʱ��" & _
                " FROM " & strSql & " A,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Group by Nvl(D.����,C.����),C.���,A.���㵥λ" & _
                " Order by ��Ŀ,���,��������,���ݺ�"
                strMoney = "4,4,1,1,1,4,1,7,7,7,1,1,1"
            Case 4 '������ϸ
                strSql = _
                " SELECT To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�," & _
                "       B.���� as ����,Nvl(D.����,C.����) as ��Ŀ,C.���,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')||A.����||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(Nvl(A.��׼����,0),'999999999" & gstrFeePrecisionFmt & "')) as ��׼����," & _
                "       Ltrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),4),'999999999" & mstrDec & "')) as ��׼���," & _
                "       Ltrim(To_Char(Nvl(A.���ʽ��,0),'999999999" & mstrDec & "')) as ���ʽ��,A.����Ա���� as ����Ա " & _
                " FROM " & strSql & " A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID" & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And ����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Union All" & _
                " SELECT NULL as ��������,NULL as ���ݺ�,NULL as ����,NULL as ��Ŀ,Null as ���,A.�վݷ�Ŀ||'ZZZZZ' as ��Ŀ," & _
                "       NULL as ����,NULL as ��׼����," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "       Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as ���ʽ��,NULL as ����Ա" & _
                " FROM " & strSql & " A,���ű� B,�շ���ĿĿ¼ C" & _
                " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID " & _
                " Group by A.�վݷ�Ŀ||'ZZZZZ' " & _
                " Order by ��Ŀ,��������,���ݺ�"
                strMoney = "4,4,1,1,1,1,1,7,7,7,1"
            Case 5 '�����嵥
                strSql = _
                " SELECT B.�ڼ�,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "       Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as ���ʽ��" & _
                " FROM " & strSql & " A,�ڼ�� B" & _
                " Where A.�Ǽ�ʱ�� Between Trunc(B.��ʼ����) and Trunc(B.��ֹ����)+1-1/24/60/60 " & _
                " Group by B.�ڼ�,A.�վݷ�Ŀ" & _
                " Union All" & _
                " SELECT B.�ڼ�||'ZZZZZ',NULL as ��Ŀ," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "       Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as ���ʽ��" & _
                " FROM " & strSql & " A,�ڼ�� B" & _
                " Where A.�Ǽ�ʱ�� Between Trunc(B.��ʼ����) and Trunc(B.��ֹ����)+1-1/24/60/60 " & _
                " Group by B.�ڼ�||'ZZZZZ'" & _
                " Order by �ڼ�,��Ŀ"
                strMoney = "4,4,7,7"
            Case 6 '�����嵥
                strSql = _
                " SELECT A.�վݷ�Ŀ as ��Ŀ," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "       Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as ���ʽ��" & _
                " FROM " & strSql & " A" & _
                " Group by A.�վݷ�Ŀ Order by ��Ŀ"
                strMoney = "4,7,7"
            Case 7 '���յ���
                strSql = _
                    " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�,A.�վݷ�Ŀ as ������Ŀ," & _
                    "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                    "        Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as ���ʽ��,A.����Ա���� as ����Ա " & _
                    " FROM " & strSql & " A" & _
                    " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD'),A.NO,A.�վݷ�Ŀ,A.����Ա����" & _
                    " Union All" & _
                    " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO||'ZZZZZ' as ���ݺ�,NULL as ������Ŀ," & _
                    "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                    "        Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as ���ʽ��, NULL as ����Ա  " & _
                    " FROM " & strSql & " A" & _
                    " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD'),A.NO" & vbCrLf & _
                    " Union All" & _
                    " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,'ZZZZZAAAAA' as ���ݺ�,NULL as ������Ŀ," & _
                    "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                    "        Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as ���ʽ��,NULL as ����Ա " & _
                    " FROM  " & strSql & " A" & _
                    " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD')" & _
                    " Order by ��������,���ݺ�,������Ŀ"
                strMoney = "4,4,4,7,7,1"
            Case 8 '���շ�Ŀ
                strSql = _
                " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.�վݷ�Ŀ as ������Ŀ," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "       Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as ���ʽ��" & _
                " FROM " & strSql & " A " & _
                " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD'),A.�վݷ�Ŀ" & _
                " Union All" & _
                " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD')||'ZZZZZ' as ��������,NULL as ������Ŀ," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),4)),'999999999" & mstrDec & "')) as ��׼���," & _
                "        Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & mstrDec & "')) as ���ʽ��" & _
                " FROM " & strSql & " A" & _
                " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD')" & _
                " Order by ��������,������Ŀ"
                strMoney = "4,4,7,7"
        End Select
        
        mshQuery.MergeCells = flexMergeFree
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngBillID)
        If rsTmp.RecordCount > 0 Then
            Set mshQuery.DataSource = rsTmp
        Else
            Call BandRectoGrid(mshQuery, rsTmp)
        End If

        mshQuery.Tag = tabCard.SelectedItem.Index
        For i = 0 To mshQuery.Cols - 1
            mshQuery.MergeCol(i) = False
        Next

        Select Case tabCard.SelectedItem.Index
            Case 2, 4  '��ϸ�嵥��������ϸ
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 5) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 8))
                            If IsNumeric(mshQuery.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 9))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.Row = i
                            mshQuery.MergeRow(i) = True
                            strTmp = mshQuery.TextMatrix(i, 5)
                            For j = 0 To 7
                                mshQuery.Col = j: mshQuery.CellAlignment = 4
                                mshQuery.TextMatrix(i, j) = "С ��:" & Left(strTmp, Len(strTmp) - 5)
                            Next
                            For j = 8 To mshQuery.Cols - 2
                                mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 7
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "�� ��"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 8) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 9) = Format(arrTotal(1), " " & mstrDec)
                End If
            Case 3 '����Ŀ��ϸ
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 4) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 8))
                            If IsNumeric(mshQuery.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 9))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.Row = i
                            mshQuery.MergeRow(i) = True
                            strTmp = mshQuery.TextMatrix(i, 3)
                            For j = 0 To 5
                                mshQuery.Col = j: mshQuery.CellAlignment = 4
                                mshQuery.TextMatrix(i, j) = "С ��:" & strTmp
                            Next
                            mshQuery.TextMatrix(i, 7) = " " '������
                            For j = 8 To mshQuery.Cols - 2
                                mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 7
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "�� ��"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 8) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 9) = Format(arrTotal(1), " " & mstrDec)
                End If
             Case 5 '�����嵥
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 2))
                            If IsNumeric(mshQuery.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 3))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.Row = i
                            mshQuery.MergeRow(i) = True
                            For j = 0 To 1
                                mshQuery.Col = j: mshQuery.CellAlignment = 4
                                mshQuery.TextMatrix(i, j) = "С��:" & mshQuery.TextMatrix(i - 1, 0)
                            Next
                            For j = 2 To mshQuery.Cols - 1
                                mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 1
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "�� ��"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 2) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 3) = Format(arrTotal(1), " " & mstrDec)
                End If
             Case 6 '�����嵥
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If IsNumeric(mshQuery.TextMatrix(i, 1)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 1))
                        If IsNumeric(mshQuery.TextMatrix(i, 2)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 2))
                        mshQuery.MergeRow(i) = False
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.Col = 0: mshQuery.CellAlignment = 4
                    mshQuery.TextMatrix(mshQuery.Row, 0) = "�� ��"
                    mshQuery.TextMatrix(mshQuery.Row, 1) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 2) = Format(arrTotal(1), " " & mstrDec)
                End If
            Case 7
                For i = 0 To mshQuery.Cols - 1
                    mshQuery.MergeCol(i) = False
                Next
                If rsTmp.RecordCount > 0 Then
                    mshQuery.MergeCol(0) = True
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not (mshQuery.TextMatrix(i, 1) Like "*ZZZZZ") And Not (mshQuery.TextMatrix(i, 1) Like "*AAAAA") Then
                            If IsNumeric(mshQuery.TextMatrix(i, 3)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 3))
                            If IsNumeric(mshQuery.TextMatrix(i, 4)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 4))
                            mshQuery.MergeRow(i) = False
                        Else
                            If Not (mshQuery.TextMatrix(i, 1) Like "*AAAAA") Then
                                mshQuery.Row = i
                                mshQuery.MergeRow(i) = True
                                For j = 1 To 2
                                    mshQuery.Col = j: mshQuery.CellAlignment = 4
                                    mshQuery.TextMatrix(i, j) = "����С��:" & mshQuery.TextMatrix(i - 1, 1)
                                Next
                                For j = 3 To mshQuery.Cols - 2
                                    mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                                Next
                            Else
                                mshQuery.Row = i
                                mshQuery.MergeRow(i) = True
                                For j = 1 To 2
                                    mshQuery.Col = j: mshQuery.CellAlignment = 4
                                    mshQuery.TextMatrix(i, j) = "��С��"
                                Next
                                For j = 3 To mshQuery.Cols - 2
                                    mshQuery.TextMatrix(i, j) = Space(j Mod 2) & mshQuery.TextMatrix(i, j)
                                Next
                            End If
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 2
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "�� ��"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 3) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 4) = Format(arrTotal(1), " " & mstrDec)
                    
                    'ɾ��ֻ��һ�е��ݵ�С����
                    j = 0
                    For i = 1 To mshQuery.Rows - 1
                        If mshQuery.TextMatrix(i, 1) Like "*С��*" Then
                            If j = 1 Then mshQuery.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                End If
            Case 8
                For i = 0 To mshQuery.Cols - 1
                    mshQuery.MergeCol(i) = False
                Next
                If rsTmp.RecordCount > 0 Then
                    mshQuery.MergeCol(0) = True
                    ReDim arrTotal(3)
                    For i = 1 To mshQuery.Rows - 1
                        If Not mshQuery.TextMatrix(i, 0) Like "*ZZZZZ" Then
                            If IsNumeric(mshQuery.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(mshQuery.TextMatrix(i, 2))
                            If IsNumeric(mshQuery.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(mshQuery.TextMatrix(i, 3))
                            mshQuery.MergeRow(i) = False
                        Else
                            mshQuery.Row = i
                            mshQuery.MergeRow(i) = False
                            mshQuery.Col = 0: mshQuery.CellAlignment = 4
                            mshQuery.TextMatrix(i, 0) = Left(mshQuery.TextMatrix(i, 0), Len(mshQuery.TextMatrix(i, 0)) - 5)
                            mshQuery.TextMatrix(i, 1) = "��С��"
                        End If
                    Next
                    mshQuery.Rows = mshQuery.Rows + 1
                    mshQuery.Row = mshQuery.Rows - 1
                    mshQuery.MergeRow(mshQuery.Row) = True
                    For i = 0 To 1
                        mshQuery.Col = i: mshQuery.CellAlignment = 4
                        mshQuery.TextMatrix(mshQuery.Row, i) = "�� ��"
                    Next
                    mshQuery.TextMatrix(mshQuery.Row, 2) = Format(arrTotal(0), mstrDec)
                    mshQuery.TextMatrix(mshQuery.Row, 3) = Format(arrTotal(1), " " & mstrDec)
                    
                    'ɾ��ֻ��һ�е��ݵ�С����
                    j = 0
                    For i = 1 To mshQuery.Rows - 1
                        If mshQuery.TextMatrix(i, 1) Like "*С��*" Then
                            If j = 1 Then mshQuery.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                End If
        End Select
    End If
    
    '�ܵĸ�ʽ����
    If mshQuery.Rows = 1 Then mshQuery.Rows = 2
    
    For i = 0 To mshQuery.Cols - 1
        mshQuery.FixedAlignment(i) = 4
    Next
    
    '���ȡ��,����û�����ó�ʼ�п�,��ӡ���쳣
    Call SetGridWidth(mshQuery, Me)
    
    '�и���¼������
    If tabCard.SelectedItem.Index = 7 And mbytInState = 0 Then
        mshQuery.ColWidth(mshQuery.Cols - 1) = 0
    End If
    
    For i = 0 To UBound(Split(strMoney, ","))
        mshQuery.ColAlignment(i) = Split(strMoney, ",")(i)
    Next
    
    mshQuery.Row = 1: mshQuery.Col = 0: mshQuery_EnterCell
    
    sta.Panels(2).Text = strPre
    
    mshQuery.Redraw = True
    mshQuery.Refresh
    Screen.MousePointer = 0
    LoadCardData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    mshQuery.Redraw = True
    If ErrCenter() = 1 Then
        mshQuery.Redraw = False
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
    sta.Panels(2).Text = strPre
End Function

Private Function GetMinMaxTime(ByVal bytMode As Byte) As Integer
'����:ȡδ������е���С������סԺ����,���ܷ���0
'����:bytMode,0-��С����,1-������
    Dim strTime As String, arrTmp As Variant
    Dim i As Long, intTime As Integer
    
    strTime = IIf(mstrTime = "", mstrAllTime, mstrTime)
    arrTmp = Split(strTime, ",")
    For i = 0 To UBound(arrTmp)
        If i = 0 Then intTime = Val(arrTmp(i))
        If bytMode = 0 Then
            If intTime > Val(arrTmp(i)) Then intTime = Val(arrTmp(i))
        Else
            If intTime < Val(arrTmp(i)) Then intTime = Val(arrTmp(i))
        End If
    Next
    
    GetMinMaxTime = intTime
End Function

Private Sub GetFeeDate(dBegin As Date, dEnd As Date)
'���ܣ���ȡ���˵���С��������ʱ��
    Dim i As Long, DateThis As Date
    
    mrsBalance.MoveFirst
    For i = 1 To mrsBalance.RecordCount
        If gint����ʱ�� = 0 Then
            DateThis = mrsBalance!�Ǽ�ʱ��
        Else
            DateThis = mrsBalance!ʱ��
        End If
        If i = 1 Then
            dBegin = DateThis
            dEnd = DateThis
        Else
            If DateThis < dBegin Then dBegin = DateThis
            If DateThis > dEnd Then dEnd = DateThis
        End If
        
        mrsBalance.MoveNext
    Next
    mrsBalance.MoveFirst
End Sub

Private Function GetPatiDate(dBegin As Date, dEnd As Date) As Boolean
'���ܣ���ȡ���˵����Ժʱ��,���ﲡ��ȡ������С����ʱ��
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim lng��ҳID As Long

    Call GetFeeDate(dBegin, dEnd)
    If mrsInfo!��ҳID <> 0 Then
        lng��ҳID = GetMinMaxTime(0)
        If lng��ҳID > 0 Then
            If lng��ҳID = mrsInfo!��ҳID Then
                dBegin = mrsInfo!��Ժ����
                If IsDate(mstr����סԺ����) Then    '����:30043
                    If Format(dBegin, "yyyy-mm-dd") < mstr����סԺ���� Then dBegin = CDate(mstr����סԺ����)
                End If
                If Not IsNull(mrsInfo!��Ժ����) Then
                    dEnd = mrsInfo!��Ժ����
                Else
                    dEnd = zlDatabase.Currentdate
                End If
            Else
                If CStr(lng��ҳID) = IIf(mstrTime = "", mstrAllTime, mstrTime) Then '�����ǽ���ǰĳ��סԺ����
                    On Error GoTo errH
                    strSql = "Select ��Ժ����,Nvl(��Ժ����,Sysdate) as ��Ժ���� From ������ҳ" & _
                            " Where ����ID=[1] And ��ҳID=[2]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CLng(mrsInfo!����ID), lng��ҳID)
                    dBegin = rsTmp!��Ժ����
                    If IsDate(mstr����סԺ����) Then
                        If Format(dBegin, "yyyy-mm-dd") < mstr����סԺ���� Then dBegin = CDate(mstr����סԺ����)
                    End If
                    dEnd = rsTmp!��Ժ����
                End If
            End If
        End If
    End If
    
    GetPatiDate = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetColNum(strHead As String) As Integer
    Dim i As Long
    For i = 0 To mshDetail.Cols - 1
        If mshDetail.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub cmdYB_Click()
'���ܣ����ﲡ�˽���ǰ�������֤(�ɶ�ҽ����֧��סԺ����ҽ�������֤)
    Dim lng����ID As Long, bytMode As Byte
    Dim strMessage As String, intInsure As Integer
    lng����ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    Call NewBill
    bytMode = 0
    If mblnMC_TwoMode Then
        If InStr(mstrPrivs, ";������ý���;") = 0 Then
            bytMode = 4
        Else
            If zlCommFun.ShowMsgbox("ҽ����֤��֤", "��ѡ���������֤ģʽ��", "!סԺҽ��(&Z),����ҽ��(&M)", Me, vbInformation) = "סԺҽ��" Then
                bytMode = 4
            End If
        End If
    End If
        
    '���˺�:����תסԺ����ʱ����
    mstrYBPati = gclsInsure.Identify(bytMode, lng����ID, intInsure)
    If mstrYBPati = "" Then GoTo ExceptionHand
    cmdOK.Enabled = False   '����:43776
    
    mbytMCMode = IIf(bytMode = 0, 1, 2) '������LoadPatientInfo֮ǰ
    If mbytMCMode = 1 Then
        '        'lng����ID:49084
        If Not gclsInsure.GetCapability(support�������, lng����ID, intInsure) Then
            strMessage = "���˵�ǰ���಻֧������ҽ�����ʡ�": GoTo ExceptionHand
        End If
    End If
    
    'New:�ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
    If UBound(Split(mstrYBPati, ";")) >= 8 Then lng����ID = Val(Split(mstrYBPati, ";")(8))
    If lng����ID <> 0 Then
        txtPatient.Text = "-" & lng����ID
        Call LoadPatientInfo(IDKIND.GetCurCard, False, intInsure)
        If mrsInfo.State = 0 Then GoTo ExceptionHand
    Else
        strMessage = "���������֤�ɹ�,��δ���ֲ��˵��ʻ���Ϣ!" & vbCrLf & "�����ǲ�����Ժʱû�н�����֤,���ܽ��б��ս��㣡"
        GoTo ExceptionHand
    End If
    Exit Sub
ExceptionHand:
    If strMessage <> "" Then Call MsgBox(strMessage, vbInformation, gstrSysName)
    Set mrsInfo = New ADODB.Recordset
    mstrYBPati = "": mbytMCMode = 0
    txtPatient.Text = "": txtPatient.SetFocus
    cmdOK.Enabled = True
End Sub

Private Sub HideMoneyInfo()
    lblҽ������.Caption = "ͳ��֧��:"
    lblҽ������.Visible = False
    lbl�����ʻ�.Caption = "�ʻ����:"
    lbl�����ʻ�.Visible = False
    Form_Resize
End Sub

Private Sub txtTotal_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txtTotal.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtTotal.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtTotal_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtTotal.hWnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Function GetPatiState(lng����ID As Long) As String
'���ܣ����ز���״̬˵��
'��ͨ��Ժ,������Ժ,ҽ����Ժ;��ͨ��Ժ,���۳�Ժ,ҽ����Ժ;������ͨ,��������,����ҽ��
    Dim lng��ҳID As Long
    If mrsInfo!��ҳID = 0 Or mbytMCMode = 1 Then
        If IsNull(mrsInfo!����) Then
            GetPatiState = "������ͨ"
        Else
            GetPatiState = "����ҽ��"
        End If
    Else
        If Nvl(mrsInfo!��������, 0) = 1 Then
            GetPatiState = "��������"
        Else
            If Not IsNull(mrsInfo!����) Then
                GetPatiState = "ҽ��"
            ElseIf Nvl(mrsInfo!��������, 0) = 2 Then
                GetPatiState = "����"
            Else
                GetPatiState = "��ͨ"
            End If
            If mbytFunc = 0 Then
                If Is��������(mrsInfo!����ID, lng��ҳID) Then
                     GetPatiState = "��������"
                Else
                    GetPatiState = "����" & GetPatiState
                End If
            Else
                If IsNull(mrsInfo!��Ժ����) Then
                    GetPatiState = GetPatiState & "��Ժ"
                Else
                    GetPatiState = GetPatiState & "��Ժ"
                End If
            End If
        End If
        If Nvl(mrsInfo!״̬, 0) = 3 Then
            GetPatiState = GetPatiState & "(Ԥ��Ժ)"
        End If
    End If
End Function

Private Function GetӦ��() As Currency
    Dim i As Long
    
    For i = 1 To vsfMoney.Rows - 1
        If Val(vsfMoney.TextMatrix(i, COLMoney.C3����)) = 1 Then
            GetӦ�� = Val(vsfMoney.TextMatrix(i, COLMoney.C1���))
            Exit Function
        End If
    Next
End Function

Private Sub txt��ע_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt��ע
End Sub

Private Sub txt��ע_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If GetӦ�� > 0 And txt�ɿ�.Visible Then
        txt�ɿ�.SetFocus
    ElseIf cmdOK.Visible And cmdOK.Enabled Then
        cmdOK.SetFocus
    End If
End Sub
Private Sub txt��ע_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt��ע, KeyAscii, m�ı�ʽ
End Sub
Private Sub txt��ע_LostFocus()
   zlCommFun.OpenIme False
End Sub

Private Sub txt�ɿ�_Change()
    
    If Val(txt�ɿ�.Text) = 0 Then txt�Ҳ�.Text = "0.00"
    Call Calc�Ҳ�
    
'    txt�Ҳ�.Text = Format(Val(txt�ɿ�.Text) - GetӦ��, "0.00")
End Sub

Private Sub txt�ɿ�_GotFocus()
    '#21 1234.56   --��������һǧ������ʮ�ĵ�����Ԫ  J
    '#22 1234.56   --Ԥ��һǧ������ʮ�ĵ�����Ԫ Y
    '#23 1234.56   --����һǧ������ʮ�ĵ�����Ԫ Z
    Dim curTotal As Currency
    
    Call zlControl.TxtSelAll(txt�ɿ�)
    If gblnLED Then
        zl9LedVoice.DisplayBank (" ")
        curTotal = GetӦ��
        If curTotal > 0 Then
            zl9LedVoice.Speak "#21 " & curTotal
        Else
            zl9LedVoice.Speak "#23 " & Abs(curTotal)
        End If
    End If
End Sub

Private Sub Led��ӭ��Ϣ()
    'LED��ʼ��
    If mbytInState = 0 And gblnLED Then
        If gblnLedWelcome Then
            zl9LedVoice.Reset com
            zl9LedVoice.Speak "#1"
            zl9LedVoice.Init UserInfo.��� & "�� Ϊ������", mlngModul, gcnOracle
        End If
        
        zl9LedVoice.DisplayPatient txtPatient.Text & " " & txtSex.Text & " " & txtOld.Text, Val("" & mrsInfo!����ID)
    End If
End Sub

Private Sub txt�ɿ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If KeyAscii = Asc(".") And InStr(txt�ɿ�.Text, ".") > 0 Then KeyAscii = 0
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt�ɿ�_LostFocus()
    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
End Sub

Private Sub txt�ɿ�_Validate(Cancel As Boolean)
    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
    
    If Val(txt�ɿ�.Text) <> 0 Then
        If CSng(txt�Ҳ�.Tag) < 0 Then
            MsgBox "�ɿ����,�벹��Ӧ�ɽ�", vbInformation, gstrSysName
            Call SelAll(txt�ɿ�): txt�ɿ�.SetFocus
            Cancel = True: Exit Sub
        End If
                
        If gblnLED Then
            zl9LedVoice.DispCharge Format(GetӦ��, "0.00"), txt�ɿ�.Text, txt�Ҳ�.Text
            zl9LedVoice.Speak "#22 " & txt�ɿ�.Text
            zl9LedVoice.Speak "#23 " & CSng(txt�Ҳ�.Tag)
            zl9LedVoice.Speak "#3"                  '#3  --�뵱�����, лл!
        End If
    End If
    
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    '�˿ؼ���ý���,��Ϊ��ʹǰһ�ؼ�:����ʱ�������,������Ԥ�������봦,�������������Ԥ�����.
    If KeyAscii = vbKeyReturn Then Call SendKeys("{Tab}")
End Sub
Private Sub Calc�Ҳ�()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼����Ҳ�
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2010-01-12 17:41:47
    '����:27360
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl�Ҳ� As Double
    If Val(txt�ɿ�.Text) = 0 Then txt�Ҳ�.Text = "0.00"
    dbl�Ҳ� = Round(Val(txt�ɿ�.Text) - GetӦ��, 2)
    txt�Ҳ�.Text = Format(Abs(dbl�Ҳ�), "0.00")
    txt�Ҳ�.Tag = dbl�Ҳ�
    If dbl�Ҳ� <= 0 Then
        lbl�Ҳ�.Caption = "�տ�"
        lbl�Ҳ�.ForeColor = &H0&
    Else
        lbl�Ҳ�.Caption = "�Ҳ�"
        lbl�Ҳ�.ForeColor = vbRed   '35830
    End If
    txt�Ҳ�.ForeColor = lbl�Ҳ�.ForeColor
End Sub
Private Sub txt�Ҳ�_Change()
    txt�Ҳ�.Tag = ""
End Sub

Private Function Get��ˢ���() As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨�Ŀ�ˢ���
    '����:
    '����:���˺�
    '����:2010-02-08 13:49:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, intCol As Integer
    Dim dbl��ˢ��� As Double, dbl��Ԥ�� As Double
    Dim dbl�ܶ� As Double
    
    dbl�ܶ� = GetBalanceSum
    dbl��ˢ��� = 0
    For i = 1 To vsfMoney.Rows - 1
        If InStr(1, ";8;1;", ";" & vsfMoney.TextMatrix(i, COLMoney.C3����) & ";") = 0 And Val(vsfMoney.TextMatrix(i, COLMoney.C1���)) <> 0 Then
            dbl��ˢ��� = dbl��ˢ��� + Val(vsfMoney.TextMatrix(i, COLMoney.C1���))
        End If
    Next
    
    dbl��Ԥ�� = 0
    For i = 1 To mshDeposit.Rows - 1
        dbl��Ԥ�� = dbl��Ԥ�� + Val(mshDeposit.TextMatrix(i, COLDeposit.��Ԥ��))
    Next
            
    dbl��ˢ��� = dbl�ܶ� - dbl��Ԥ�� - dbl��ˢ���
    If dbl��ˢ��� < 0 Then dbl��ˢ��� = 0
    Get��ˢ��� = Format(dbl��ˢ���, gstrDec)
End Function

Private Function zlSquareCardFeeList(ByRef rsFeeList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨��ϸ��Ϣ
    '���:
    '����:rsFreeList-������ϸ����
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-05 16:02:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim rsTemp As ADODB.Recordset, strDate As String, strInvoice As String
    Dim i As Long
    
    If mrsInfo Is Nothing Then Exit Function
    If mrsBalance Is Nothing Then Exit Function
    
    If zlCreateFeeListStruc(rsFeeList) = False Then Exit Function
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    
    Set rsTemp = mrsBalance  'GetVBalance(mstrPrivs, mrsInfo!����, mrsInfo!����ID, mstrTime, mDateBegin, mDateEnd, False, mblnDateMoved, mbytBaby, mbytMCMode = 1, mbytKind, mstrItem, mstrUnit, mstrClass)
    rsTemp.Filter = 0
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
          rsFeeList.AddNew
          rsFeeList!������� = 1
          rsFeeList!�ѱ� = Nvl(rsTemp!�ѱ�)
          rsFeeList!NO = Nvl(rsTemp!���ݺ�)
          rsFeeList!ʵ��Ʊ�� = txtInvoice.Text
          rsFeeList!����ʱ�� = CDate(strDate)
          rsFeeList!����ID = Val(Nvl(mrsInfo!����ID))
          rsFeeList!��ҳID = Val(Nvl(rsTemp!��ҳID))
          rsFeeList!�շ���� = Nvl(rsTemp!�շ����)
          If Nvl(rsTemp!��Ŀ) <> "" Then
              rsFeeList!�վݷ�Ŀ = Nvl(rsTemp!��Ŀ)
          Else
              rsFeeList!�վݷ�Ŀ = Null
          End If
          rsFeeList!������ = Nvl(rsTemp!������)
          rsFeeList!�շ�ϸĿID = Val(Nvl(rsTemp!�շ�ϸĿID))
          rsFeeList!���㵥λ = Nvl(rsTemp!���㵥λ)
          rsFeeList!���� = Val(Nvl(rsTemp!����))
          rsFeeList!���� = Format(Val(Nvl(rsTemp!�۸�)), gstrFeePrecisionFmt)
          rsFeeList!ʵ�ս�� = Format(Val(Nvl(rsTemp!δ����)), gstrDec)
          rsFeeList!ͳ���� = Format(Val(Nvl(rsTemp!ͳ����)), gstrDec)
          rsFeeList!����֧������ID = IIf(Val(Nvl(rsTemp!���մ���ID)) = 0, Null, Val(Nvl(rsTemp!���մ���ID)))
          rsFeeList!�Ƿ�ҽ�� = 0 ' Val(Nvl(rsTemp!�Ƿ�ҽ��))
          rsFeeList!���ձ��� = Null ' Nvl(rsTemp!���ձ���)
          rsFeeList!ժҪ = Null ' Nvl(rsTemp!ժҪ)
          rsFeeList!�Ƿ��� = 0 ' Val(Nvl(rsTemp!�Ƿ���))
          rsFeeList!��������ID = Val(Nvl(rsTemp!��������ID))
          rsFeeList!ִ�в���ID = Val(Nvl(rsTemp!ִ�в���ID))
          rsFeeList!���ν��� = 0
          rsFeeList.Update
          rsTemp.MoveNext
    Loop
     If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    zlSquareCardFeeList = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function סԺˢ���㿨() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:סԺˢ���㿨
     '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-06 09:32:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String, cll����ϼ� As Collection, strTemp As String, strNone As String
    Dim dblTemp As Double
    Dim arrPage As Variant, arrBalance() As String, strBalance As String
    Dim cur���ʺϼ� As Currency, cur���� As Currency, cur������ As Currency, cur�ɷ���� As Currency
    Dim i As Integer, j As Integer, k As Integer, P As Integer
    Dim strDate As String, strAdvance As String, strInvoice As String, str���㷽ʽ As String
                
    strInvoice = Trim(txtInvoice.Text)
    
    On Error GoTo errH
    strTemp = "": strNone = ""
    mtySquareCard.strˢ������ = ""
    Set cll����ϼ� = New Collection
    '
    '���㷽ʽ;���;�Ƿ������޸�|..."
    '�ȼ����ֽ��㷽ʽ�Ƿ����?
    ''"�ӿڱ��" "���ѿ�ID",  "����", "���㷽ʽ", "������",   "���",  "������"  "����ʱ��",  "��ע",  "�����־"
    With mtySquareCard.rsSquare
        .Filter = 0: If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '���������øý��㷽ʽ,��Ϊ���㿨�Ľ��㷽ʽ
            str���㷽ʽ = Nvl(!���㷽ʽ)
            mrs���㷽ʽ.Filter = "����='" & str���㷽ʽ & "' And ����=8"
            If mrs���㷽ʽ.EOF Then
               If InStr(strNone & ",", "," & str���㷽ʽ & ",") = 0 Then
                   strNone = strNone & "," & str���㷽ʽ
               End If
            End If
            If InStr(1, strTemp & ",", "," & str���㷽ʽ & ",") > 0 Then
                dblTemp = Val(cll����ϼ�("K" & str���㷽ʽ)(0)) + Val(Nvl(!������))
                cll����ϼ�.Remove "K" & str���㷽ʽ
            Else
                dblTemp = Val(Nvl(!������))
            End If
            cll����ϼ�.Add Array(dblTemp, str���㷽ʽ), "K" & str���㷽ʽ
            strTemp = strTemp & "," & str���㷽ʽ
            .MoveNext
        Loop
    End With
    
    If strNone <> "" Then
        strNone = Mid(strNone, 2)
        MsgBox "��ǰ���㿨�Ľ���ʹ�õĽ��㷽ʽ" & vbCrLf & vbCrLf & vbTab & strNone & vbCrLf & vbCrLf & _
        "�ڽ���δ���ã����ȵ����㷽ʽ������������Щ���㷽ʽ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    str���㷽ʽ = ""
    For i = 1 To cll����ϼ�.Count
        str���㷽ʽ = cll����ϼ�(i)(1)
        If InStr(1, mtySquareCard.strˢ������, ";" & str���㷽ʽ & ";") = 0 Then
            dblTemp = 0
            For j = 1 To cll����ϼ�.Count
                If str���㷽ʽ = cll����ϼ�(j)(1) Then
                    dblTemp = dblTemp + Val(cll����ϼ�(i)(0))
                End If
            Next
            mtySquareCard.strˢ������ = ";" & str���㷽ʽ & ";" & dblTemp & ";0|"
        End If
    Next
    If mtySquareCard.strˢ������ <> "" Then
        mtySquareCard.strˢ������ = Mid(mtySquareCard.strˢ������, 2)
        mtySquareCard.strˢ������ = Mid(mtySquareCard.strˢ������, 1, Len(mtySquareCard.strˢ������) - 1)
    End If
    ShowMoney True, , mty_ModulePara.bytMzDeposit
    Screen.MousePointer = 0
    סԺˢ���㿨 = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlReCalcRequare(ByRef cur������� As Currency, ByRef strNotBlance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ý��ʿ����ֽ��
    '���:
    '����:cur�������-���ص�ǰ�����Ľ������
    '     strNotBlance-����δ���ý������Ϣ
    '����:����ɹ���,����true,���򷵻�Flase
    '����:���˺�
    '����:2010-02-08 14:27:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varBalace As Variant, i As Long, j As Long
    Dim varItem As Variant, strMoney As String
    
    If mtySquareCard.strˢ������ = "" Then zlReCalcRequare = True: Exit Function
    '���㷽ʽ;���;�Ƿ������޸�|..."
    varBalace = Split(mtySquareCard.strˢ������, "|")
    With vsfMoney
        '���ý��ʿ����ֽ��
        For i = 0 To UBound(varBalace)
            strMoney = varBalace(i) '���㷽ʽ;���;�Ƿ������޸�|....
            varItem = Split(strMoney, ";")  '���㷽ʽ;���;�Ƿ������޸�
            For j = 1 To .Rows - 1
                If .TextMatrix(j, COLMoney.C0����) = CStr(varItem(0)) And InStr(",8,", Val(vsfMoney.TextMatrix(j, COLMoney.C3����))) > 0 Then
                     .TextMatrix(j, COLMoney.C1���) = Format(CCur(varItem(1)), "0.00")
                    If Val(varItem(2)) = 0 Then
                        vsfMoney.RowData(j) = 1 '�ý�����ɸ���
                    Else
                        vsfMoney.RowData(j) = 0 '�ý�������Ը���
                    End If
                    '������㿨�Ѵ���Ľ���
                    cur������� = cur������� - Format(Val(vsfMoney.TextMatrix(j, COLMoney.C1���)), "0.00")
                    Exit For
                End If
            Next
            'δ����ҽ���ɱ������㷽ʽ
            If j = vsfMoney.Rows Then
                mrs���㷽ʽ.Filter = "���㷽ʽ='" & varItem(0) & "'"
                If mrs���㷽ʽ.EOF Then
                    strNotBlance = strNotBlance & vbCrLf & vbTab & CStr(Split(strMoney, ";")(0)) & ":" & Format(CCur(Split(strMoney, ";")(1)), "0.00")
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, COLMoney.C1���) = Format(CCur(varItem(1)), "0.00")
                    .TextMatrix(.Rows - 1, COLMoney.C0����) = varItem(0)
                    .TextMatrix(.Rows - 1, COLMoney.C3����) = Nvl(mrs���㷽ʽ!����)
                    If Val(varItem(2)) = 0 Then
                        vsfMoney.RowData(.Rows - 1) = 1 '�ý�����ɸ���
                    Else
                        vsfMoney.RowData(.Rows - 1) = 0 '�ý�������Ը���
                    End If
                    '������㿨�Ѵ���Ľ���
                    cur������� = cur������� - Format(Val(vsfMoney.TextMatrix(.Rows - 1, COLMoney.C1���)), "0.00")
                End If
            End If
        Next
    End With
End Function


Private Function zlCallSquare_DelFree(ByVal str����ID_In As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�п������˷�
    '���:str����ID_In��ԭ����ID
    '����:
    '����:������óɹ�,����true,���򷵻�False,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-12 14:19:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Err = 0: On Error GoTo Errhand:
    '���ŵ��ݲ����ڿ�����,�˳�
    If Not mtySquareCard.bln������ Then zlCallSquare_DelFree = True: Exit Function

    'Zl_���˿������¼_Strike(����id_In In Varchar2)
    strSql = "Zl_���˿������¼_Strike(" & str����ID_In & ")"
    zlDatabase.ExecuteProcedure strSql, Me.Caption

    'Public Function zlDelSquareFee(frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, ByVal str����ID_IN As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����: zlSquareFee (����ӿ�)
    '    '���:frmMain:HIS���� ���õ�������
    '    '     intCallType : HIS���� 0-  ������õ��� 1-  סԺ���ʵ���
    '    '     str����ID_IN: HIS���� ���ν��ʵĽ���ID��
    '    '����:
    '    '����:true:���óɹ�,False:����ʧ��
    '    '����:���˺�
    '    '����:2009-12-15 15:18:38
    '    '˵��:
    '    '    1. "�����շѹ���"��"סԺ���ʹ���"������ʱ,���ô˽ӿ�
    '    'ע:
    '    '  �˽ӿ���������HIS������ , ��˲����ڴ˽ӿڴ������û������Ĳ���
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlDelSquareFee(Me, mlngModul, mstrPrivs, str����ID_In) = False Then
        zlCallSquare_DelFree = False
        gcnOracle.RollbackTrans
    Else
        zlCallSquare_DelFree = True
    End If
    Exit Function
Errhand:
    gcnOracle.RollbackTrans
    Call ErrCenter
End Function
Private Function zlIsCheckCanelFee(ByVal str����ID_In As String, ByVal bln�����˷� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷��Ƿ�Ϸ�,�Ϸ�������true,���򷵻�False
    '���:str����ID_IN-����ID_IN
    '����:
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2010-01-14 09:45:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If mtySquareCard.bln������ = False Then zlIsCheckCanelFee = True: Exit Function
    '���˷�,����Ҫ�����㿨�Ƿ�װ����
    If gobjSquare.objSquareCard Is Nothing Then
        ShowMsgbox ("ע��:" & vbCrLf & "    ��ǰû�а�װ�����㲿�������ܽ����˷�,���飡")
        Exit Function
    End If
    If bln�����˷� Then
        ShowMsgbox ("ע��:" & vbCrLf & "    ˢ��ʱ�ķ��õ������ܽ��в����˷�,���飡")
        Exit Function
    End If
    If str����ID_In = "" Then
        ShowMsgbox ("ע��:" & vbCrLf & "    δѡ���˷ѵĵ��ݣ����ܽ����˷�,���飡")
        Exit Function
    End If

    'Public Function zlCheckDelSquareValied(frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, ByVal str����ID_IN As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:��ִ���˷�ʱ,�����صĽӿڲ����Ƿ�����
    '    '���:
    '    '����:
    '    '����:����,����true,���򷵻�False
    '    '����:���˺�
    '    '����:2009-12-31 16:39:47
    '    '˵��;
    '    '     ���˷�ʱ����Ҫ������صļ��
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlCheckDelSquareValied(Me, mlngModul, mstrPrivs, str����ID_In) = False Then
        Exit Function
    End If
    zlIsCheckCanelFee = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub zlClear���㿨()
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '����:������㿨�������Ϣ
        '����:���˺�
        '����:2010-01-11 11:26:20
        '---------------------------------------------------------------------------------------------------------------------------------------------
        Dim j As Long
        If cmd���㿨.Visible = False Then Exit Sub
        cmd���㿨.TabStop = True
        '��Ҫ����ˢ������
        Set mtySquareCard.rsSquare = Nothing
        mtySquareCard.strˢ������ = ""
        '��Ҫ�������е�ˢ������
        With vsfMoney
            '���ý��ʿ����ֽ��
            For j = 1 To .Rows - 1
                If InStr(",8,", Val(vsfMoney.TextMatrix(j, COLMoney.C3����))) > 0 Then
                     .TextMatrix(j, COLMoney.C1���) = "0.00"
                End If
            Next
        End With
    End Sub
Private Function IsCheck�����ѽ���(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡���Ƿ��Ѿ�����
    '���:
    '����:
    '����:�ѽ��շ���True,���򷵻�False
    '����:���˺�
    '����:2010-05-24 16:39:47
    '˵��;
    '     ����:30036
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSql = "select nvl(��Ϣֵ,0) as �������� from ������ҳ�ӱ� where ����id=[1] and ��ҳid=[2] and ��Ϣ��='��������'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lng��ҳID)
    If Not rsTemp.EOF Then
            IsCheck�����ѽ��� = Val(Nvl(rsTemp!��������)) = 1
    Else
            IsCheck�����ѽ��� = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub zlSetDefaultTime(ByVal lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ��סԺ����
    '���:lng����ID-����ID
    '       lng��ҳID-��ҳID
    '����:
    '����:���˺�
    '����:2010-05-24 16:39:47
    '˵��;
    '     ����:30043
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim strDate As String
    
    strSql = "" & _
    "   Select to_char( Max(��������)+1,'yyyy-mm-dd') as �������� " & _
    "   From ���˽��ʼ�¼ " & _
    "   Where  ��¼״̬=1  And ����iD=[1] and nvl(��;����,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
    If Not rsTemp.EOF Then
        strDate = Nvl(rsTemp!��������)
    Else
        strDate = ""
    End If
    mstr����סԺ���� = strDate
End Sub

Private Sub zlChangeDefaultTime()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��ı�ȱʡ����
    '���ƣ����˺�
    '���ڣ�2010-05-25 10:25:18
    '˵����30043
    '------------------------------------------------------------------------------------------------------------------------
    If opt��Ժ.Value Then
        txtPatiEnd.Text = txtPatiEnd.Tag
    Else
        txtPatiEnd.Text = Format(zlDatabase.Currentdate - 1, "yyyy-mm-dd")
        If txtPatiEnd.Text < txtPatiBegin.Text Then
            txtPatiEnd.Text = txtPatiEnd.Tag
        End If
    End If
End Sub
Private Function zlGetPatiSource() As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Դ(��ҪӦ�����Ƿ���λ��)
    '����:1-����;2-סԺ
    '����:���˺�
    '����:2011-03-14 18:01:36
    '�����:36121
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str����IDs As String, rsTemp As ADODB.Recordset
    Dim bln���� As Boolean, blnסԺ As Boolean
    Dim strTable As String, strSql As String
    Dim BytType As Byte
    Dim i As Long
    
    Err = 0: On Error GoTo Errhand:
    '0-Ȩ����;1-��סԺ;2-�����סԺ
    BytType = Zl���˷�����Դ
    '���Ѵ�Ź���:
    '���ֻ�������,����������ü�¼��;
    '���������סԺ���ʵ�,�����סԺ���ü�¼��;
    If BytType <> 2 Then
        'ֱ��ȷ�����˵�,�򷵻�
        zlGetPatiSource = IIf(BytType = 0, 1, 2): Exit Function
    End If
    '������ֲ�������,����Ҫ���������Ǳߵ�,
    '���������סԺ(��������Ҳ��סԺ��),��������סԺ;
    '������ý��������,������������
    With mshDetail
        For i = 1 To .Rows - 1
            If blnסԺ Then
                zlGetPatiSource = 2: Exit Function
            End If
            If Val(.TextMatrix(i, COL_��־)) = 1 Then
                bln���� = True
            Else
                blnסԺ = True
            End If
        Next
    End With
    If bln���� And blnסԺ = False Then
        zlGetPatiSource = 1
    Else
        zlGetPatiSource = 2
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���³�ʼ�����˷�Ʊ��Ϣ
    '����:���˺�
    '����:2011-04-29 14:17:33
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    Dim intInsure As Integer
    intInsure = mintInsure
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng����ID = Val(Nvl(mrsInfo!����ID)): lng��ҳID = Val(Nvl(mrsInfo!��ҳID))
            intInsure = Val(Nvl(mrsInfo!����))
        End If
    End If
    If mblnStartFactUseType Then mlng����ID = 0
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(lng����ID, lng��ҳID, intInsure)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModul, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModul, mstrUseType)
    mintInvoiceMode = zl_GetInvoicePrintMode(mlngModul, mstrUseType)
    If blnFact Then Call RefreshFact
    Call ShowBillFormat
End Sub
Private Function zlGetInvoiceGroupUseID(ByRef lng����ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵ�����ID
    '���:lng����ID-����id
    '       intNum-ҳ��
    '       strInvoiceNO-����ķ�Ʊ��
    '����:lng����ID-����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    lng����ID = GetInvoiceGroupID(1, intNum, lng����ID, mlngShareUseID, strInvoiceNO, mstrUseType)
    If lng����ID <= 0 Then
        Select Case lng����ID
            Case 0 '����ʧ��
            Case -1
                If Trim(mstrUseType) = "" Then
                    MsgBox "��û�����ú͹��õ��շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "��û�����ú͹��õġ�" & mstrUseType & "���շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mstrUseType) = "" Then
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "���صĹ���Ʊ�ݵġ�" & mstrUseType & "���շ�Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
                Exit Function
            Case -3
                MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,���������룡", vbInformation, gstrSysName
                If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus
                Exit Function
        End Select
    End If
    zlGetInvoiceGroupUseID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨����������Ϣ
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytInState = 1 Then Exit Sub
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKIND.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "")
    Set objCard = IDKIND.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKIND.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
    Else
        gobjSquare.blnȱʡ�������� = IDKIND.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    gobjSquare.bln��ȱʡ������ = IDKIND.Cards.��ȱʡ������
    mtySquareCard.blnExistsObjects = isExistsThreeSwap
    
End Sub
Private Sub InitԤ�����()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��Ԥ�����
    '����:���˺�
    '����:2011-09-05 01:53:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int��� As Integer, varPage As Variant
    Dim i As Integer
    mintԤ����� = IIf(mbytFunc = 0, 1, 2)
'    mintԤ����� = 2
'    If InStr(1, "," & mstrTime & ",", ",0,") > 0 Then
'        varPage = Split(mstrTime, ",")
'         mintԤ����� = 1
'        For i = 0 To UBound(varPage)
'            '�����סԺ,ֻ��ȫ��ʾ������
'            If Val(varPage(i)) > 0 Then mintԤ����� = 0: Exit For
'        Next
'    End If
End Sub
Private Function isExistsThreeSwap() As Boolean
    Dim strPayType As String, varData As Variant, varTemp As Variant
    Dim i As Long, j As Long
    If gobjSquare Is Nothing Then Exit Function
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    varData = Split(strPayType, ";")
    For i = 0 To UBound(varData)
        If InStr(1, varData(i), "|") <> 0 Then
            varTemp = Split(varData(i), "|")
            If Val(varTemp(5)) = 1 Then
                'Ŀǰֻ������ѿ�
                isExistsThreeSwap = True: Exit Function
            End If
            j = j + 1
        End If
    Next
End Function
Private Sub WriteZYInforToCard(ByVal lng����ID As Long, ByVal lng����ID As Long, Optional blnDelete As Boolean = False)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��סԺ��Ϣд�뿨��
    '���:blnDelete-�Ƿ��˷�
    '����:���˺�
    '����:2012-12-14 17:06:27
    '˵��:
    '����:56615
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, strExpend As String
    'δȷ��ˢ�����,ֱ���˳�
    If InStr(1, mstrPrivs, ";סԺ��Ϣд��;") = 0 Then Exit Sub
    If lng����ID = 0 Then Exit Sub
    If mlngCardTypeID = 0 Then
        If blnDelete Then GoTo goDelete:
        Exit Sub
    End If
    Dim objCard As Card
    If IDKIND.GetCurCard.�ӿ���� = mlngCardTypeID Then
        Set objCard = IDKIND.GetCurCard
    Else
        Set objCard = IDKIND.GetIDKindCard(mlngCardTypeID, CardTypeID)
    End If
    If objCard Is Nothing Then Exit Sub
    If objCard.�Ƿ�д�� = False Or objCard.�ӿ���� <= 0 Then Exit Sub '��׼д����,�����ýӿ�
    lngCardTypeID = objCard.�ӿ����
goDelete:
    If mbytFunc = 0 Then
        Call gobjSquare.objSquareCard.zlMzInforWriteToCard(Me, mlngModul, lngCardTypeID, _
        lng����ID, lng����ID, strExpend)
    Else
        Call gobjSquare.objSquareCard.zlzyInforWriteToCard(Me, mlngModul, lngCardTypeID, _
        lng����ID, lng����ID, strExpend)
    End If
End Sub

Private Function GetDelBalanceID(ByVal strNo As String, ByRef lng����ID As Long) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ϵĽ���ID
    '����:lng����ID-���ز���ID
    '����:�������ϵĽ���ID
    '����:���˺�
    '����:2012-12-14 18:52:31
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSql = "Select ID,����ID From ���˽��ʼ�¼ Where  NO=[1] and ��¼״̬=2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    If rsTemp.EOF Then Exit Function
    lng����ID = Val(Nvl(rsTemp!����ID))
    GetDelBalanceID = Val(Nvl(rsTemp!ID))
    GetDelBalanceID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
