VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDiseaseRegist 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ⱦ�����Խ��������"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10680
   Icon            =   "frmDiseaseRegist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13455.55
   ScaleMode       =   0  'User
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11175
      TabIndex        =   54
      Top             =   0
      Width           =   11175
      Begin XtremeCommandBars.CommandBars cbsMain 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.VScrollBar vsbReport 
      Height          =   7335
      LargeChange     =   50
      Left            =   0
      Max             =   100
      SmallChange     =   10
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar hsbReport 
      Height          =   255
      LargeChange     =   500
      Left            =   0
      Max             =   100
      SmallChange     =   10
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.Frame frmMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9495
      Left            =   240
      TabIndex        =   52
      Top             =   360
      Width           =   10455
      Begin MSComCtl2.MonthView dtpDate 
         Bindings        =   "frmDiseaseRegist.frx":6852
         Height          =   2220
         Left            =   4440
         TabIndex        =   51
         Top             =   4200
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   3916
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         StartOfWeek     =   41091073
         TitleBackColor  =   -2147483636
         TitleForeColor  =   -2147483634
         TrailingForeColor=   -2147483637
         CurrentDate     =   37904
      End
      Begin VB.CommandButton cmdInfo 
         Height          =   250
         Index           =   0
         Left            =   3910
         Picture         =   "frmDiseaseRegist.frx":6866
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   4435
         Width           =   250
      End
      Begin VB.CommandButton cmdInfo 
         Height          =   250
         Index           =   5
         Left            =   3090
         Picture         =   "frmDiseaseRegist.frx":6B80
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2435
         Width           =   250
      End
      Begin VB.CommandButton cmdInfo 
         Height          =   250
         Index           =   4
         Left            =   9810
         Picture         =   "frmDiseaseRegist.frx":6E9A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2045
         Width           =   250
      End
      Begin VB.CommandButton cmdInfo 
         Height          =   250
         Index           =   3
         Left            =   6450
         Picture         =   "frmDiseaseRegist.frx":71B4
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2045
         Width           =   250
      End
      Begin VB.CommandButton cmdInfo 
         Height          =   250
         Index           =   2
         Left            =   6450
         Picture         =   "frmDiseaseRegist.frx":74CE
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2435
         Width           =   250
      End
      Begin VB.CommandButton cmdInfo 
         Height          =   250
         Index           =   1
         Left            =   3090
         Picture         =   "frmDiseaseRegist.frx":77E8
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2035
         Width           =   250
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   6
         Left            =   4660
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2045
         Width           =   1725
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   11
         Left            =   8020
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2045
         Width           =   1725
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   16
         Left            =   8265
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   6180
         Width           =   1725
      End
      Begin VB.Frame fraIdea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   0
         TabIndex        =   53
         Top             =   6570
         Width           =   10700
         Begin VB.ComboBox cboReport 
            BackColor       =   &H8000000E&
            Height          =   300
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1800
            Width           =   3255
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1440
            Index           =   18
            Left            =   1350
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Top             =   180
            Width           =   8685
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
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
            Index           =   9
            Left            =   8265
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   2280
            Width           =   1725
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
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
            Index           =   10
            Left            =   8265
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   1800
            Width           =   1725
         End
         Begin VB.Label lblReport 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�����������ı��濨:"
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
            Height          =   255
            Left            =   300
            TabIndex        =   43
            Top             =   1850
            Width           =   2175
         End
         Begin VB.Line Line2 
            BorderStyle     =   2  'Dash
            X1              =   -30
            X2              =   11130
            Y1              =   45
            Y2              =   45
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   300
            TabIndex        =   41
            Top             =   195
            Width           =   900
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ȷ��ҽʦ"
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
            Index           =   10
            Left            =   7320
            TabIndex        =   45
            Top             =   1800
            Width           =   840
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
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
            Index           =   11
            Left            =   7320
            TabIndex        =   47
            Top             =   2325
            Width           =   840
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   8160
            X2              =   10035
            Y1              =   2505
            Y2              =   2505
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   8160
            X2              =   10035
            Y1              =   2025
            Y2              =   2025
         End
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Index           =   17
         Left            =   1350
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   2845
         Width           =   8685
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   7
         Left            =   6340
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1245
         Width           =   1100
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   5
         Left            =   6340
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1645
         Width           =   1100
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   4
         Left            =   1300
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1645
         Width           =   1100
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   3
         Left            =   3820
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1245
         Width           =   1100
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   2
         Left            =   1300
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1245
         Width           =   1100
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   1
         Left            =   3820
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1645
         Width           =   1100
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   19
         Left            =   1300
         TabIndex        =   16
         Top             =   2045
         Width           =   1725
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   12
         Left            =   4660
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2445
         Width           =   1725
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   8
         Left            =   1320
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   4445
         Width           =   2525
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   13
         Left            =   8265
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   5655
         Width           =   1725
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   14
         Left            =   8265
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   5130
         Width           =   1725
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   15
         Left            =   1300
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2445
         Width           =   1725
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Index           =   0
         Left            =   8860
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1245
         Width           =   1100
      End
      Begin VB.Line Line1 
         Index           =   17
         X1              =   4605
         X2              =   6510
         Y1              =   2275
         Y2              =   2275
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ͼ����"
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
         Index           =   16
         Left            =   3720
         TabIndex        =   18
         Top             =   2045
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ͼ�ҽ��"
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
         Index           =   6
         Left            =   7080
         TabIndex        =   21
         Top             =   2045
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   7965
         X2              =   9870
         Y1              =   2275
         Y2              =   2275
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽ���"
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
         Index           =   36
         Left            =   7530
         TabIndex        =   35
         Top             =   5160
         Width           =   630
      End
      Begin VB.Line Line1 
         Index           =   31
         X1              =   8175
         X2              =   10030
         Y1              =   6405
         Y2              =   6405
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽ�ʱ��"
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
         Index           =   34
         Left            =   7320
         TabIndex        =   39
         Top             =   6180
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽǿ���"
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
         Index           =   33
         Left            =   7320
         TabIndex        =   37
         Top             =   5655
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   1245
         X2              =   3960
         Y1              =   4675
         Y2              =   4675
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   13
         Left            =   300
         TabIndex        =   30
         Top             =   2845
         Width           =   900
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���Ƽ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   19
         Left            =   300
         TabIndex        =   32
         Top             =   4445
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   6285
         X2              =   7515
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   8805
         X2              =   10035
         Y1              =   1475
         Y2              =   1475
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   6285
         X2              =   7515
         Y1              =   1475
         Y2              =   1475
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   3765
         X2              =   4995
         Y1              =   1475
         Y2              =   1475
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   1245
         X2              =   2475
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   1245
         X2              =   2475
         Y1              =   1475
         Y2              =   1475
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   3765
         X2              =   4995
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3240
         X2              =   7080
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ⱦ�����Խ��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   3270
         TabIndex        =   0
         Top             =   360
         Width           =   3750
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
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
         Index           =   5
         Left            =   5400
         TabIndex        =   13
         Top             =   1645
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
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
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Top             =   1645
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
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
         Index           =   2
         Left            =   360
         TabIndex        =   1
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ס Ժ ��"
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
         Index           =   1
         Left            =   2880
         TabIndex        =   11
         Top             =   1645
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
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
         Index           =   7
         Left            =   5400
         TabIndex        =   5
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
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
         Index           =   3
         Left            =   2880
         TabIndex        =   3
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ͼ�ʱ��"
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
         Index           =   35
         Left            =   360
         TabIndex        =   15
         Top             =   2045
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   16
         X1              =   1245
         X2              =   3150
         Y1              =   2275
         Y2              =   2275
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
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
         Index           =   8
         Left            =   7920
         TabIndex        =   7
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�걾����"
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
         Index           =   9
         Left            =   3750
         TabIndex        =   27
         Top             =   2445
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   4605
         X2              =   6510
         Y1              =   2675
         Y2              =   2675
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   8175
         X2              =   10030
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line1 
         Index           =   15
         X1              =   8175
         X2              =   10030
         Y1              =   5355
         Y2              =   5355
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   1245
         X2              =   3150
         Y1              =   2675
         Y2              =   2675
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ʱ��"
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
         Index           =   12
         Left            =   360
         TabIndex        =   24
         Top             =   2445
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmDiseaseRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************************
'������֮���԰�CommandBars����PictureBox���棬����Ϊ�ô����������ط��ᱻ��Ƕ��TabControl���棬
'����CommandBars����PictureBox����Ļ��������
'***************************************************************************************************
Public Event PatiTransfer(ByVal lng����ID As Long, ByVal str�Һ�No As String) 'ת��
Public Event Closed(ByVal lngFunID As Long, ByVal strTag As String) '�����д����ϣ��رմ���ʱ����lngFunID�̶�Ϊ0��strTag ��չ����δʹ�� ""��

Private Enum mCtlID '�����ϵĿؼ�����ֵ
    txt���� = 0
    txtסԺ�� = 1
    txt���� = 2
    txt�Ա� = 3
    txt���� = 4
    txt���� = 5
    txt�ͼ���� = 6
    txt���� = 7
    txt���Ƽ��� = 8
    txt����ʱ�� = 9
    txtȷ��ҽʦ = 10
    txt�ͼ�ҽ�� = 11
    txt�걾���� = 12
    txt�Ǽǿ��� = 13
    txt�Ǽ��� = 14
    txt���ʱ�� = 15
    txt�Ǽ�ʱ�� = 16
    txt������� = 17
    txt������� = 18
    txt�ͼ�ʱ�� = 19
    
    cmd���Ƽ��� = 0
    cmd�ͼ�ʱ�� = 1
    cmd�걾���� = 2
    cmd�ͼ���� = 3
    cmd�ͼ�ҽ�� = 4
    cmd���ʱ�� = 5
End Enum

Private mlngID As Long   '�������Լ�¼ ���ID
Private mint���� As Integer    '0-סԺ��1-����
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mstr�Һ�NO As String
Private mlng�Һ�ID As Long
Private mlng����ID As Long '����ID���� ������ҳ.��ǰ����ID
Private mstr���� As String '���ﲡ������ ���˹Һż�¼.����
Private mlng�Ǽǿ���ID As Long
Private mdat�ͼ�ʱ�� As Date
Private mdat���ʱ�� As Date
Private mlng�ͼ����ID As Long
Private mstr�ͼ�ҽ�� As String
Private mstr�걾���� As String
Private mstr������� As String
Private mstr���ƴ�Ⱦ�� As String
Private mintType As Integer  '0��ʾ��д��ֻ��ʾ�ϰ벿�֣���1-��ʾҽ������ֻ���°벿�ֿɱ༭����2-�鿴�����в��ɱ༭���ɲ鿴�°벿�֣�,3-�޸ģ��ɱ༭���ɲ鿴�ϰ벿�֣�
Private mlng����ID As Long
Private mIntState As Integer '�������ĵ�ǰ״̬��0-������д��1-��ҽ��ȷ�ϣ�2-ҽ���Ѵ���,3-�Ǵ�Ⱦ����4-ת�ƴ�����
Private mblnOk As Boolean

Private mdat�Ǽ�ʱ�� As Date
Private mstr�Ǽ��� As String
Private mstr������� As String

Private mstr����ʱ�� As Date
Private mintResult As Integer   '1-���ͣ�2-��ɣ�3-ת��
Private WithEvents mclsDiagEdit As zlMedRecPage.clsDiagEdit
Attribute mclsDiagEdit.VB_VarHelpID = -1
Private mclsMipModule As zl9ComLib.clsMipModule
Private mblnDiagnose As Boolean     '�Ƿ���д�����

Private mblnDialog As Boolean        '�Ƿ���ʾΪ����
Private mlngTop As Long              '����ʾΪ����ʱ���ϱ߾�
Private mblnSbSisible As Boolean     '����ʾΪ����ʱ���������Ƿ�ɼ�
Private mlngҽ��ID As Long           '��д�÷�������ҽ��ID
Private mstr��� As String           '���Խ����Ӧ��ҽ������� D/E   ���/����(�ɼ���ʽ)
Private mblnNoID As Boolean

Public Function ShowDiseaseRegist(ByRef frmParent As Object, ByVal intType As Integer, Optional ByVal lngID As Long, _
                Optional ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long, Optional ByVal str�Һ�No As String, _
                Optional ByVal lngҽ��id As Long, Optional ByVal var�Ǽǿ��� As Variant, Optional ByVal dat�ͼ�ʱ�� As Date, Optional ByVal var�ͼ���� As Variant, _
                Optional ByVal str�ͼ�ҽ�� As String, Optional ByVal str�걾���� As String, Optional ByVal str������� As String, _
                Optional ByVal dat���ʱ�� As Date, Optional ByVal str���ƴ�Ⱦ�� As String, Optional ByRef objMip As Object, Optional ByVal str�Ǽ��� As String) As Integer
'���ܣ����ô�Ⱦ�����Խ��������
'������intType 0��ʾ��д��ֻ��ʾ�ϰ벿�֣���1-��ʾҽ������ֻ���°벿�ֿɱ༭����2-�鿴�����в��ɱ༭���ɲ鿴�°벿�֣�,3-�޸ģ��ɱ༭���ɲ鿴�ϰ벿�֣�
'      lngID  = �������Լ�¼ ID
'      lng����ID = ����ID
'      lng��ҳID=סԺ:��ҳID
'      str�Һ�No =����Һŵ�NO
'      lng�Ǽǿ���ID �� lng�ͼ����ID ֮�����ǿɱ��������� ��Ϊ�˼���LIS������װ
    mintType = intType
    mlngID = lngID
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstr�Һ�NO = str�Һ�No
    mlngҽ��ID = lngҽ��id
    mdat�ͼ�ʱ�� = dat�ͼ�ʱ��
    mstr�ͼ�ҽ�� = str�ͼ�ҽ��
    mstr�걾���� = str�걾����
    mstr������� = str�������
    mdat���ʱ�� = dat���ʱ��
    mstr���ƴ�Ⱦ�� = str���ƴ�Ⱦ��
    mstr�Ǽ��� = str�Ǽ���
    
    If TypeName(var�ͼ����) = "String" Then         '���ı���
        mlng�ͼ����ID = GetDeptID(var�ͼ����)
    ElseIf IsNumeric(var�ͼ����) Then
        mlng�ͼ����ID = Val(var�ͼ����)
    Else
        mlng�ͼ����ID = 0
    End If
    
    If TypeName(var�Ǽǿ���) = "String" Then
        mlng�Ǽǿ���ID = GetDeptID(var�Ǽǿ���)
    ElseIf IsNumeric(var�Ǽǿ���) Then
        mlng�Ǽǿ���ID = Val(var�Ǽǿ���)
    Else
        mlng�Ǽǿ���ID = 0
    End If
    
    mintResult = 0
    
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip

     '�ж���סԺ���˻������ﲡ��
    If intType = 0 Then
        If mlng��ҳID = 0 And str�Һ�No <> "" Then
            mint���� = 1
        ElseIf mlng��ҳID <> 0 And str�Һ�No = "" Then
            mint���� = 0
        Else
            Call MsgBox("����������ͳ�Ժ����Ĳ���!", vbInformation, gstrSysName)
            Exit Function
        End If
    End If
    mblnDialog = True
    On Error Resume Next
    Me.Show 1, frmParent
        
    ShowDiseaseRegist = mintResult
End Function


Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox

    '������----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = gobjComlib.zlCommFun.GetPubIcons

    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Send, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_OK, "ȷ��Ϊ��Ⱦ��")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_NO, "�Ǵ�Ⱦ��")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_ViewReport, "����鿴"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Transfer, "ת��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
End Sub

Private Function SetDiagnose() As Boolean
    Dim lng����ID As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnCancle As Boolean
    Dim mstr���IDs As String, str��� As String
On Error GoTo errH

    If mclsDiagEdit Is Nothing Then
        Set mclsDiagEdit = New zlMedRecPage.clsDiagEdit
        Call mclsDiagEdit.InitDiagEdit(gcnOracle, glngSys, IIf(mint���� = 1, 1260, 1261), gclsMipModule)
    End If
    
    lng����ID = IIf(mint���� = 0, mlng��ҳID, mlng�Һ�ID)

    strSQL = "Select rownum as ID, a.����id, a.���id, b.���� || '-' || b.���� As ����, c.���� || '-' || c.���� As ���, a.���没��" & vbNewLine & _
            "From ��������ǰ�� A, ��������Ŀ¼ B, �������Ŀ¼ C" & vbNewLine & _
            "Where a.����id = b.Id(+) And a.���id = c.Id(+) And a.���没�� = [1]"
    Set rsTemp = gobjComlib.FS.ShowSQLSelectEx(Me, txtInfo(2), strSQL, 0, "���", False, "", "���Ƽ��������ѡ��", False, False, False, blnCancle, True, False, True, "MultiCheckReturn=1", mstr���ƴ�Ⱦ��)
    
    If Not blnCancle Then
        If rsTemp Is Nothing Then
            SetDiagnose = mclsDiagEdit.ShowDiagEdit(Me, mlngID, mlng����ID, lng����ID, IIf(mint���� = 1, 1, 2), mlng����ID, mstr���IDs, str���, 0)
        Else
            SetDiagnose = mclsDiagEdit.ConfirmInfectiousDiseases(Me, mlngID, mlng����ID, lng����ID, IIf(mint���� = 1, 1, 2), mlng����ID, rsTemp)
        End If
    End If
    SetDiagnose = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_Print
            If mintType = 0 Or mintType = 3 Then
                If CheckData Then
                    If SendDiseaseRecord Then
                        Call PrintDiseaseRegist(2, mlngID, Me)
                        mintType = 3
                    End If
                End If
            ElseIf mintType = 1 Then
                If SaveDisProcessData(6) Then
                    Call PrintDiseaseRegist(2, mlngID, Me)
                End If
            Else
                Call PrintDiseaseRegist(2, mlngID, Me)
            End If
        Case conMenu_Tool_Send
            If CheckData Then
                If SendDiseaseRecord Then
                    If mintType = 0 Then
                        Call SendMsg
                    End If
                    mintResult = 1
                    Unload Me
                End If
            End If
        Case conMenu_Tool_OK
            If SaveDisProcessData(2) Then
                 mintResult = 2
                 Call SetDiagnose
                 mblnOk = True
                 Unload Me
            End If
        Case conMenu_Tool_NO
            If cboReport.Text <> "" Then
                If MsgBox("�÷������Ѿ������˱��濨��ȷ��Ϊ�Ǵ�Ⱦ����ȡ���������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
            If SaveDisProcessData(3) Then
                mintResult = 2
                Unload Me
            End If
        Case conMenu_Tool_Transfer
            If SaveDisProcessData(4) Then
                RaiseEvent PatiTransfer(mlng����ID, mstr�Һ�NO)
                mintResult = 3
                Unload Me
            End If
        Case conMenu_Tool_ViewReport
            Call ViewEPRReport
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Not mblnDialog Then Exit Sub
    Select Case Control.ID
        Case conMenu_Tool_Send
            Control.Visible = (mintType = 0 Or mintType = 3)
        Case conMenu_Tool_OK
            Control.Visible = (mintType = 1)
        Case conMenu_Tool_NO
            Control.Visible = (mintType = 1)
        Case conMenu_Tool_Transfer
            Control.Visible = (mintType = 1)
            If Control.Visible Then Control.Visible = (mint���� = 1)
        Case conMenu_Tool_ViewReport
            Control.Visible = (mintType = 1)
    End Select
End Sub

Private Sub cmdInfo_Click(Index As Integer)
    Select Case Index
        Case cmd���Ƽ���
            Call GetDiseaseList(1)
        Case cmd�ͼ�ʱ��
            If IsDate(txtInfo(txt�ͼ�ʱ��).Text) Then
                dtpDate.Value = CDate(txtInfo(txt�ͼ�ʱ��).Text)
            Else
                dtpDate.Value = gobjComlib.zlDatabase.Currentdate
            End If
            dtpDate.Tag = "�ͼ�ʱ��"
            dtpDate.Left = txtInfo(txt�ͼ�ʱ��).Left
            dtpDate.Top = txtInfo(txt�ͼ�ʱ��).Top + txtInfo(txt�ͼ�ʱ��).Height
            dtpDate.Visible = True
            dtpDate.SetFocus
        Case cmd�걾����
            Call GetSampleList(1)
        Case cmd�ͼ����
            Call GetInspectDept(1)
        Case cmd�ͼ�ҽ��
            Call GetInspectDoctor(1)
        Case cmd���ʱ��
            If IsDate(txtInfo(txt���ʱ��).Text) Then
                dtpDate.Value = CDate(txtInfo(txt���ʱ��).Text)
            Else
                dtpDate.Value = gobjComlib.zlDatabase.Currentdate
            End If
            dtpDate.Tag = "���ʱ��"
            dtpDate.Left = txtInfo(txt���ʱ��).Left
            dtpDate.Top = txtInfo(txt���ʱ��).Top + txtInfo(txt���ʱ��).Height
            dtpDate.Visible = True
            dtpDate.SetFocus
    End Select
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String
    
    If dtpDate.Tag = "�ͼ�ʱ��" Then
        'ȡֵ
        If IsDate(txtInfo(txt�ͼ�ʱ��).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(txt�ͼ�ʱ��).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        txtInfo(txt�ͼ�ʱ��).Text = strDate
        txtInfo(txt�ͼ�ʱ��).Tag = strDate
        mdat�ͼ�ʱ�� = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        txtInfo(txt�ͼ�ʱ��).SetFocus
    ElseIf dtpDate.Tag = "���ʱ��" Then
        'ȡֵ
        If IsDate(txtInfo(txt���ʱ��).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(txt���ʱ��).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        txtInfo(txt���ʱ��).Text = strDate
        txtInfo(txt���ʱ��).Tag = strDate
        mdat���ʱ�� = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        txtInfo(txt���ʱ��).SetFocus
    End If
End Sub

Private Sub Form_Load()
    picMenu.Visible = mblnDialog
    mblnNoID = False
    If mblnDialog Then
        Me.BorderStyle = 3
        lblReport.Visible = True
        cboReport.Visible = True
        Set mclsDiagEdit = New zlMedRecPage.clsDiagEdit
        Call mclsDiagEdit.InitDiagEdit(gcnOracle, glngSys, IIf(mint���� = 1, 1260, 1261), gclsMipModule)
        Call InitCommandBar
        Call SetFormState(mintType)
        Call LoadPatiInfo
        If mblnNoID Then
            Unload Me
        Else
            Call SaveDisProcessData(1)
        End If
    Else
        lblReport.Visible = False
        cboReport.Visible = False
    End If
End Sub

'��ȡ�Ǽ��ˣ��Ǽ�ʱ�䣬�Ǽǿ���
Private Sub GetRegistInfo()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
On Error GoTo errH
    
    '��ȡ�Ǽǿ���
    strSQL = "Select a.Id, a.���� From ���ű� A Where ID = [1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�Ǽǿ���ID)
    If rsTmp.RecordCount > 0 Then
        mlng�Ǽǿ���ID = Val(rsTmp!ID)
        txtInfo(txt�Ǽǿ���).Text = rsTmp!���� & ""
    Else
        mlng�Ǽǿ���ID = UserInfo.����ID
        txtInfo(txt�Ǽǿ���).Text = UserInfo.������
    End If

    If mstr�Ǽ��� = "" Then mstr�Ǽ��� = UserInfo.����
    mdat�Ǽ�ʱ�� = gobjComlib.zlDatabase.Currentdate
    txtInfo(txt�Ǽ���).Text = mstr�Ǽ���
    txtInfo(txt�Ǽ�ʱ��).Text = Format(mdat�Ǽ�ʱ��, "yyyy-MM-dd HH:mm")

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDiseaseList(ByVal intType As Integer)
'���ܣ���ȡ���Ƽ���Ŀ¼
'������0 �ı��򰴻س���1 �㰴ť
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(txt���Ƽ���).Tag = txtInfo(txt���Ƽ���).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(txt���Ƽ���).Text = "" Then  '�൱�����������Ŀ
            txtInfo(txt���Ƽ���).Tag = ""
            mstr���ƴ�Ⱦ�� = ""
            Call SeekNextCtl
            Exit Sub
        End If
    End If
    
    strSQL = "select A.���� as ID,A.����,A.���� from ��Ⱦ��Ŀ¼ A" & IIf(intType = 0, " where A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2] ", "") & " order by A.����"
        
    strInput = Trim(UCase(txtInfo(txt���Ƽ���).Text))
    vRect = gobjComlib.zlControl.GetControlRect(txtInfo(txt���Ƽ���).hwnd)
    
    Set rsTmp = gobjComlib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��Ⱦ��", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(txt���Ƽ���).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call MsgBox("û���ҵ�ƥ��ļ���!", vbInformation, gstrSysName)
            txtInfo(txt���Ƽ���).SetFocus
            mstr���ƴ�Ⱦ�� = ""
            gobjComlib.zlControl.TxtSelAll txtInfo(txt���Ƽ���)
        End If
        Exit Sub
    Else
        mstr���ƴ�Ⱦ�� = rsTmp!���� & ""
        txtInfo(txt���Ƽ���).Text = rsTmp!���� & ""
        txtInfo(txt���Ƽ���).Tag = rsTmp!���� & ""
        txtInfo(txt���Ƽ���).SetFocus
        Call SeekNextCtl
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPatiInfo()
'���ܣ���ȡ���˻�����Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    '��ȡ���������Ϣ
    If mintType <> 0 Then
        strSQL = "Select a.Id, a.����id, a.��ҳid, a.�Һŵ�, a.ҽ��ID,a.�ͼ�ʱ��, a.�ͼ����id, a.�ͼ�ҽ��, a.�걾����, a.�������, a.��Ⱦ������ As ���Ƽ���, " & vbNewLine & _
                " a.���ʱ��, a.�Ǽ�ʱ��, a.�Ǽ���, a.�Ǽǿ���id,a.��¼״̬, a.������, a.����ʱ��, a.�������˵��,b.�������" & vbNewLine & _
                "From �������Լ�¼ A,����ҽ����¼ b Where a.ҽ��ID=b.id(+) and a.Id = [1]"

        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        
        If rsTmp.RecordCount > 0 Then
            mlng����ID = Val(rsTmp!����ID & "")
            mlng��ҳID = Val(rsTmp!��ҳID & "")
            mstr�Һ�NO = rsTmp!�Һŵ� & ""
            mlngҽ��ID = Val(rsTmp!ҽ��id & "")
            mIntState = Val(rsTmp!��¼״̬ & "")
            mstr��� = rsTmp!������� & ""
            If mlng��ҳID = 0 And mstr�Һ�NO <> "" Then
                mint���� = 1
            Else
                mint���� = 0
            End If

            If IsDate(rsTmp!�Ǽ�ʱ�� & "") Then
                txtInfo(txt�Ǽ�ʱ��).Text = Format(rsTmp!�Ǽ�ʱ�� & "", "YYYY-MM-DD HH:mm")
            End If
            
            If IsDate(rsTmp!����ʱ�� & "") Then
                txtInfo(txt����ʱ��).Text = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:mm")
            End If
            
            If IsDate(rsTmp!�ͼ�ʱ�� & "") Then
                mdat�ͼ�ʱ�� = Format(rsTmp!�ͼ�ʱ�� & "", "YYYY-MM-DD HH:mm")
            End If
            
            If IsDate(rsTmp!���ʱ�� & "") Then
                mdat���ʱ�� = Format(rsTmp!���ʱ�� & "", "YYYY-MM-DD HH:mm")
            End If
            
            txtInfo(txt�Ǽ���).Text = rsTmp!�Ǽ��� & ""
            txtInfo(txtȷ��ҽʦ).Text = rsTmp!������ & ""
            txtInfo(txt�������).Text = rsTmp!�������˵�� & ""
            mstr���ƴ�Ⱦ�� = rsTmp!���Ƽ��� & ""
            mlng�ͼ����ID = Val(rsTmp!�ͼ����ID & "")
            mlng�Ǽǿ���ID = Val(rsTmp!�Ǽǿ���ID & "")
            mstr�ͼ�ҽ�� = rsTmp!�ͼ�ҽ�� & ""
            mstr�걾���� = rsTmp!�걾���� & ""
            mstr������� = rsTmp!������� & ""
        Else
            mblnNoID = True
        End If
    End If

    If mint���� = 0 Then
        strSQL = "Select A.סԺ��, Nvl(C.����, A.����) ����, Nvl(C.�Ա�, A.�Ա�) �Ա�, Nvl(C.����, A.����) ����,B.ID as ����ID, B.���� As ����, C.��Ժ���� As ��ǰ����, c.��ǰ���� as ����," & _
                "C.��Ժ���� as ����ʱ��, C.����,c.��ǰ����ID,c.��Ժ����ID" & vbNewLine & _
                "From ������Ϣ A, ���ű� B, ������ҳ C" & vbNewLine & _
                "Where C.��Ժ����id = B.Id And A.����id = C.����id And A.��ҳid = C.��ҳid And C.����id = [1] And C.��ҳid = [2]"
        
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    Else
        strSQL = "Select a.id,A.����,A.�Ա�,A.����,a.no,a.����� as סԺ��,B.ID as ����ID, b.���� as ����, null as ����,a.ִ��ʱ�� as ����ʱ��,a.����" & _
                " From ���˹Һż�¼ A,���ű� b " & _
                " Where A.NO=[1] And a.��¼����=1 And a.��¼״̬=1 And A.����ID+0=[2] and a.ִ�в���id=b.id"
        
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һ�NO, mlng����ID)
    End If
    
    If rsTmp.RecordCount > 0 Then
        txtInfo(txt����).Text = rsTmp!���� & ""
        txtInfo(txt�Ա�).Text = rsTmp!�Ա� & ""
        txtInfo(txt����).Text = rsTmp!���� & ""
        txtInfo(txt����).Text = rsTmp!���� & ""
        txtInfo(txt����).Text = rsTmp!���� & ""
        If mint���� = 0 Then
            txtInfo(txt����).Text = rsTmp!��ǰ���� & ""
            mlng����ID = Val(rsTmp!��ǰ����ID & "")
        Else
            txtInfo(txt����).Text = ""
            mlng�Һ�ID = Val(rsTmp!ID & "")
            mstr���� = rsTmp!���� & ""
        End If
        If IsDate(rsTmp!����ʱ�� & "") Then
             mstr����ʱ�� = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:mm")
        End If
        txtInfo(txtסԺ��).Text = rsTmp!סԺ�� & ""
        mlng����ID = Val(rsTmp!����ID & "")
    End If
        
    If mdat�ͼ�ʱ�� <> CDate(0) Then
        txtInfo(txt�ͼ�ʱ��).Text = Format(mdat�ͼ�ʱ��, "yyyy-MM-dd HH:mm")
        txtInfo(txt�ͼ�ʱ��).Tag = txtInfo(txt�ͼ�ʱ��).Text
    End If
    
    '��ȡ�ͼ����
    If mlng�ͼ����ID <> 0 Then
        strSQL = "Select a.Id, a.���� From ���ű� A Where ID = [1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�ͼ����ID)
        If rsTmp.RecordCount > 0 Then
            txtInfo(txt�ͼ����).Text = rsTmp!���� & ""
            txtInfo(txt�ͼ����).Tag = txtInfo(txt�ͼ����).Text
        End If
    End If
    
    If mdat���ʱ�� <> CDate(0) Then
        txtInfo(txt���ʱ��).Text = Format(mdat���ʱ��, "yyyy-MM-dd HH:mm")
        txtInfo(txt���ʱ��).Tag = txtInfo(txt���ʱ��).Text
    End If
    
    txtInfo(txt�ͼ�ҽ��).Text = mstr�ͼ�ҽ��
    txtInfo(txt�걾����).Text = mstr�걾����
    txtInfo(txt�������).Text = mstr�������
    txtInfo(txt���Ƽ���).Text = mstr���ƴ�Ⱦ��
    txtInfo(txt�ͼ�ҽ��).Tag = txtInfo(txt�ͼ�ҽ��).Text
    txtInfo(txt�걾����).Tag = txtInfo(txt�걾����).Text
    txtInfo(txt���Ƽ���).Tag = txtInfo(txt���Ƽ���).Text
    
    If mintType = 1 Or mintType = 2 Then
        '��ȡ�Ǽǿ���
        If mlng�Ǽǿ���ID <> 0 Then
            strSQL = "Select a.Id, a.���� From ���ű� A Where ID = [1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)"
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�Ǽǿ���ID)
            If rsTmp.RecordCount > 0 Then
                txtInfo(txt�Ǽǿ���).Text = rsTmp!���� & ""
            End If
        End If
    ElseIf mintType = 0 Or mintType = 3 Then
         Call GetRegistInfo
    End If
    
    If mintType = 1 Then
        glngOpenedID = mlngID
        txtInfo(txtȷ��ҽʦ).Text = UserInfo.����
        txtInfo(txt����ʱ��).Text = Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    End If

    If mblnDialog Then
        Call SetCboReportData(mlng����ID, mstr���ƴ�Ⱦ��, mlngID)
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetSampleList(ByVal intType As Integer)
'���ܣ���ȡ�걾����
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(txt�걾����).Tag = txtInfo(txt�걾����).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(txt�걾����).Text = "" Then '�൱�����������Ŀ
            txtInfo(txt�걾����).Tag = ""
            mstr�걾���� = ""
            Call SeekNextCtl
            Exit Sub
        End If
    End If
    
    strSQL = "select A.���� as ID,A.���� from ���Ƽ���걾 A" & IIf(intType = 0, " where A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2]", "") & " order by A.����"
        
    strInput = Trim(UCase(txtInfo(txt�걾����).Text))
    vRect = gobjComlib.zlControl.GetControlRect(txtInfo(txt�걾����).hwnd)
    
    Set rsTmp = gobjComlib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��Ⱦ��", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(txt�걾����).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call MsgBox("û���ҵ�ƥ��ı걾!", vbInformation, gstrSysName)
            txtInfo(txt�걾����).SetFocus
            mstr�걾���� = ""
            gobjComlib.zlControl.TxtSelAll txtInfo(txt�걾����)
        End If
        Exit Sub
    Else
        txtInfo(txt�걾����).Text = rsTmp!���� & ""
        txtInfo(txt�걾����).Tag = rsTmp!���� & ""
        txtInfo(txt�걾����).SetFocus
        mstr�걾���� = rsTmp!���� & ""
        Call SeekNextCtl
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetInspectDept(ByVal intType As Integer)
'���ܣ���ȡ�ͼ����
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(txt�ͼ����).Tag = txtInfo(txt�ͼ����).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(txt�ͼ����).Text = "" Then '�൱�����������Ŀ
            txtInfo(txt�ͼ����).Tag = ""
            mlng�ͼ����ID = 0
            Call SeekNextCtl
            Exit Sub
        End If
    End If
    If mint���� = 0 Then
        strTemp = " and B.������� in (2,3) "
    ElseIf mint���� = 1 Then
        strTemp = " and B.������� in (1,3) "
    End If
    
    strInput = Trim(UCase(txtInfo(txt�ͼ����).Text))
    vRect = gobjComlib.zlControl.GetControlRect(txtInfo(txt�ͼ����).hwnd)
        
    If mstr�ͼ�ҽ�� = "" Then
        strSQL = "Select Distinct A.ID,A.����,A.���� as ����,A.���� From ���ű� A,��������˵�� B " & _
                " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) And a.Id = b.����id" & _
                IIf(intType = 0, " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])", "") & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) And B.��������= '�ٴ�' " & strTemp & "  Order by A.����"
    Else
        strSQL = "Select Distinct d.Id, d.����, d.���� As ����, d.���� " & vbNewLine & _
                "From ��Ա�� A, ��������˵�� B,������Ա C, ���ű� D " & vbNewLine & _
                "Where (D.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� is NULL) " & vbNewLine & _
                 IIf(intType = 0, " And (D.���� Like [1] Or D.���� Like [2] Or D.���� Like [2])", "") & vbNewLine & _
                "and a.Id = c.��Աid And d.Id = B.����id And c.����id = d.Id  And a.���� = [3]" & vbNewLine & _
                " And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null) And B.��������= '�ٴ�' " & strTemp & "  Order by D.����"
    End If
    
    Set rsTmp = gobjComlib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��Ⱦ��", False, "", "", False, False, True, _
                vRect.Left, vRect.Top, txtInfo(txt�ͼ����).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%", mstr�ͼ�ҽ��)
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call MsgBox("û���ҵ�ƥ��Ŀ���!", vbInformation, gstrSysName)
            mlng�ͼ����ID = 0
            txtInfo(txt�ͼ����).SetFocus
            gobjComlib.zlControl.TxtSelAll txtInfo(txt�ͼ����)
        End If
        Exit Sub
    Else
        mlng�ͼ����ID = Val(rsTmp!ID)
        txtInfo(txt�ͼ����).Text = rsTmp!���� & ""
        txtInfo(txt�ͼ����).Tag = rsTmp!���� & ""
        txtInfo(txt�ͼ����).SetFocus
        Call SeekNextCtl
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetInspectDoctor(ByVal intType As Integer)
'���ܣ���ȡ�ͼ�ҽ��,��Ա����Ϊ ҽ�� ����������Ϊ �ٴ���������סԺ�������� �������Ա
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean
    Dim lngDeptId As Long
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(txt�ͼ�ҽ��).Tag = txtInfo(txt�ͼ�ҽ��).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(txt�ͼ�ҽ��).Text = "" Then '�൱�����������Ŀ
            txtInfo(txt�ͼ�ҽ��).Tag = ""
             mstr�ͼ�ҽ�� = ""
            Call SeekNextCtl
            Exit Sub
        End If
    End If
    
    strSQL = "Select Distinct a.Id, a.���, a.����, a.����, d.���� As ���� ,d.ID as ����ID" & vbNewLine & _
            "From ��Ա�� A, ��Ա����˵�� B, ������Ա C, ���ű� D, ��������˵�� E " & vbNewLine & _
            "Where a.Id = b.��Աid And b.��Ա���� = 'ҽ��' And a.Id = c.��Աid  " & vbNewLine & _
             IIf(mlng�ͼ����ID = 0, "And c.ȱʡ = 1 ", "And c.����id = [1] ") & vbNewLine & _
             IIf(intType = 0, " And (A.��� Like [2] Or A.���� Like [3] Or A.���� Like [3]) ", "") & vbNewLine & _
            "And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And " & vbNewLine & _
            "(d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.����ʱ�� Is Null) And c.����id = d.Id And d.Id = e.����id  " & vbNewLine & _
            IIf(mint���� = 0, "and e.������� In (2, 3) ", "and e.������� In (1, 3) ") & vbNewLine & _
            "And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null)  Order By a.���"
    
    strInput = Trim(UCase(txtInfo(txt�ͼ�ҽ��).Text))
    vRect = gobjComlib.zlControl.GetControlRect(txtInfo(txt�ͼ�ҽ��).hwnd)
    
    Set rsTmp = gobjComlib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��Ⱦ��", False, "", "", False, False, True, _
                vRect.Left, vRect.Top, txtInfo(txt�ͼ�ҽ��).Height, blnCancel, False, True, mlng�ͼ����ID, strInput & "%", gstrLike & strInput & "%")
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call MsgBox("û���ҵ�ƥ���ҽ��!", vbInformation, gstrSysName)
             mstr�ͼ�ҽ�� = ""
            txtInfo(txt�ͼ�ҽ��).SetFocus
            gobjComlib.zlControl.TxtSelAll txtInfo(txt�ͼ�ҽ��)
        End If
        Exit Sub
    Else
        txtInfo(txt�ͼ�ҽ��).Text = rsTmp!���� & ""
        txtInfo(txt�ͼ�ҽ��).Tag = rsTmp!���� & ""
        mstr�ͼ�ҽ�� = rsTmp!���� & ""
        If (mlng�ͼ����ID = 0) Then
            txtInfo(txt�ͼ����).Text = rsTmp!���� & ""
            txtInfo(txt�ͼ����).Tag = rsTmp!���� & ""
            mlng�ͼ����ID = Val(rsTmp!����ID & "")
        End If
        
        txtInfo(txt�ͼ�ҽ��).SetFocus
        Call SeekNextCtl
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetFormState(ByVal intType As Integer)
    Dim objControl As Object
    
    If intType = 0 Or intType = 3 Then
        SetControlEnabled txtInfo(txt����), False
        SetControlEnabled txtInfo(txt�Ա�), False
        SetControlEnabled txtInfo(txt����), False
        SetControlEnabled txtInfo(txt����), False
        SetControlEnabled txtInfo(txt����), False
        SetControlEnabled txtInfo(txtסԺ��), False
        SetControlEnabled txtInfo(txt����), False

        SetControlEnabled txtInfo(txt�Ǽ���), False, False
        SetControlEnabled txtInfo(txt�Ǽǿ���), False, False
        SetControlEnabled txtInfo(txt�Ǽ�ʱ��), False, False
        If intType = 0 Then
            SetControlEnabled txtInfo(txt�ͼ����), mlng�ͼ����ID = 0
            SetControlEnabled txtInfo(txt�ͼ�ҽ��), mstr�ͼ�ҽ�� = ""
            SetControlEnabled txtInfo(txt�걾����), mstr�걾���� = ""
            SetControlEnabled txtInfo(txt�ͼ�ʱ��), mdat�ͼ�ʱ�� = 0
            SetControlEnabled txtInfo(txt���ʱ��), mdat���ʱ�� = 0
            SetControlEnabled txtInfo(txt���Ƽ���), mstr���ƴ�Ⱦ�� = ""
            
            SetControlEnabled cmdInfo(cmd�ͼ����), mlng�ͼ����ID = 0
            SetControlEnabled cmdInfo(cmd�ͼ�ҽ��), mstr�ͼ�ҽ�� = ""
            SetControlEnabled cmdInfo(cmd�걾����), mstr�걾���� = ""
            SetControlEnabled cmdInfo(cmd�ͼ�ʱ��), mdat�ͼ�ʱ�� = 0
            SetControlEnabled cmdInfo(cmd���ʱ��), mdat���ʱ�� = 0
            SetControlEnabled cmdInfo(cmd���Ƽ���), mstr���ƴ�Ⱦ�� = ""
        End If
        fraIdea.Visible = False
        Me.Height = 7800
    ElseIf intType = 1 Then
        For Each objControl In Me.Controls
            SetControlEnabled objControl, False
        Next
        SetControlEnabled txtInfo(txt�������), True
        SetControlEnabled cboReport, True
    ElseIf intType = 2 Then
        For Each objControl In Me.Controls
            SetControlEnabled objControl, False
        Next
    End If
    
    lblInfo(txtסԺ��).Caption = IIf(mint���� = 0, "ס Ժ ��", "�� �� ��")
    lblInfo(txt����).Visible = (mint���� = 0)
    txtInfo(txt����).Visible = (mint���� = 0)
    Line1(txt����).Visible = (mint���� = 0)
End Sub

Private Sub SetControlEnabled(objControl As Object, ByVal blnEnabled As Boolean, Optional blnColor As Boolean = True)
'���ܣ����ÿؼ��Ŀ�����
    Select Case TypeName(objControl)
        Case "TextBox"
            objControl.Locked = Not blnEnabled
            objControl.TabStop = blnEnabled
            If blnColor Then objControl.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
        Case "CommandButton", "ComboBox"
            objControl.Enabled = blnEnabled
    End Select
End Sub

Private Function SeekNextCtl() As Boolean
'���ܣ���λ����һ������Ŀؼ���
    Call gobjComlib.zlCommFun.PressKey(vbKeyTab)
    SeekNextCtl = True
End Function

Private Function SendDiseaseRecord() As Boolean
    '����
    Dim strSQL As String
    Dim str��ҳID As String
    Dim str�Һ�No As String
    Dim str�ͼ�ʱ�� As String
    Dim str�ͼ�ҽ�� As String
    Dim str�걾���� As String
    Dim str������� As String
    Dim str���ƴ�Ⱦ�� As String
    Dim str���ʱ�� As String
    Dim str�Ǽ�ʱ�� As String
    Dim str�Ǽ��� As String
 On Error GoTo errH
    If mint���� = 0 Then
        str�Һ�No = "NULL"
        str��ҳID = CStr(mlng��ҳID)
    ElseIf mint���� = 1 Then
        str��ҳID = "NULL"
        str�Һ�No = "'" & mstr�Һ�NO & "'"
    End If
    
    If mdat�ͼ�ʱ�� = CDate(0) Then
        str�ͼ�ʱ�� = "NULL"
    Else
        str�ͼ�ʱ�� = "to_date('" & Format(mdat�ͼ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
    End If
    
    If mdat���ʱ�� = CDate(0) Then
        str���ʱ�� = "NULL"
    Else
        str���ʱ�� = "to_date('" & Format(mdat���ʱ��, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
    End If
    
    If mdat�Ǽ�ʱ�� = CDate(0) Then
        str�Ǽ�ʱ�� = "NULL"
    Else
        str�Ǽ�ʱ�� = "to_date('" & Format(mdat�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
    End If
    
    str�ͼ�ҽ�� = "'" & mstr�ͼ�ҽ�� & "'"
    str�걾���� = "'" & mstr�걾���� & "'"
    str������� = "'" & mstr������� & "'"
    str���ƴ�Ⱦ�� = "'" & mstr���ƴ�Ⱦ�� & "'"
    str�Ǽ��� = "'" & mstr�Ǽ��� & "'"
    If mintType = 0 Then
        mlngID = gobjComlib.zlDatabase.GetNextId("�������Լ�¼")       '��ȡID
        strSQL = "Zl_�������Լ���¼_Insert(" & mlngID & "," & mlng����ID & "," & str��ҳID & "," & str�Һ�No & "," & IIf(mlngҽ��ID = 0, "NULL", mlngҽ��ID) & "," _
                & str�ͼ�ʱ�� & "," & IIf(mlng�ͼ����ID = 0, "NULL", mlng�ͼ����ID) & "," & str�ͼ�ҽ�� & "," & str�걾���� & "," & str������� & "," _
                & str���ƴ�Ⱦ�� & "," & str���ʱ�� & "," & str�Ǽ�ʱ�� & "," & str�Ǽ��� & "," & IIf(mlng�Ǽǿ���ID = 0, "NULL", mlng�Ǽǿ���ID) & "," & 1 & ")"
    ElseIf mintType = 3 Then
        strSQL = "Zl_�������Լ���¼_Update(" & 4 & "," & mlngID & ",NULL,NULL,NULL,NULL,NULL," & str�ͼ�ʱ�� & "," & IIf(mlng�ͼ����ID = 0, "NULL", mlng�ͼ����ID) & "," & str�ͼ�ҽ�� & "," & str�걾���� & "," & str������� & "," _
                & str���ƴ�Ⱦ�� & "," & str���ʱ�� & "," & str�Ǽ�ʱ�� & "," & str�Ǽ��� & "," & IIf(mlng�Ǽǿ���ID = 0, "NULL", mlng�Ǽǿ���ID) & ")"
    End If
    Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, Me.Caption)

    SendDiseaseRecord = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function SaveDisProcessData(ByVal intType As Integer) As Boolean
'���ܣ���������
'������intType   1-�ʼ����ʱ�ı�״̬��ʱ�䣬ҽ����2-ȷ����3-�Ǵ�Ⱦ����4-ת�ƣ�5-�رգ�6-��ӡǰ����
    Dim strSQL As String
    Dim str����ʱ�� As String
    Dim str����ҽ�� As String
    Dim str������� As String, strTmp As String
    Dim lngReportID As Long
    Dim intDisState As Integer      '�������Լ�¼.��¼״̬,�������ĵ�ǰ״̬��1-��ҽ��ȷ�ϣ�2-ҽ���Ѵ���(���ȷ��),3-�Ǵ�Ⱦ����4-ת�ƴ�����
    
    On Error GoTo errH
    
    If mintType <> 1 Then Exit Function
    
    strTmp = Trim(txtInfo(txt�������).Text)
    str����ʱ�� = "to_date('" & Format(txtInfo(txt����ʱ��).Text, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
    str����ҽ�� = "'" & txtInfo(txtȷ��ҽʦ).Text & "'"
    str������� = "'" & strTmp & "'"


    If cboReport.ListCount > 1 Then
        lngReportID = cboReport.ItemData(cboReport.ListIndex)
    End If

    If (mIntState = 1 Or mIntState = 4) And (intType = 1 Or intType = 2) Or (mIntState = 3 And intType = 2) Then
        strSQL = "Zl_�������Լ���¼_update(1," & mlngID & "," & IIf(lngReportID = 0, "NULL", lngReportID) & "," & "2" & "," & str����ҽ�� & "," & str����ʱ�� & "," & str������� & ")"
        Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        mIntState = 2
    ElseIf intType = 3 Or intType = 4 Then
        If strTmp = "" Then
            MsgBox "�������˵������Ϊ�գ�������д�������˵����", vbInformation, gstrSysName
            Exit Function
        End If
        If intType = 3 Then lngReportID = 0
        strSQL = "Zl_�������Լ���¼_update(1," & mlngID & "," & IIf(lngReportID = 0, "NULL", lngReportID) & "," & CStr(intType) & "," & str����ҽ�� & "," & str����ʱ�� & "," & str������� & ")"
        Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        mIntState = intType
    ElseIf intType = 5 Then
        If strTmp = "" Then
            MsgBox "�������˵������Ϊ�գ�������д�������˵����", vbInformation, gstrSysName
            Exit Function
        End If
        strSQL = "Zl_�������Լ���¼_update(1," & mlngID & "," & IIf(lngReportID = 0, "NULL", lngReportID) & "," & CStr(mIntState) & "," & str����ҽ�� & "," & str����ʱ�� & "," & str������� & ")"
        Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    ElseIf intType = 6 Then
        strSQL = "Zl_�������Լ���¼_update(1," & mlngID & "," & IIf(lngReportID = 0, "NULL", lngReportID) & "," & CStr(mIntState) & "," & str����ҽ�� & "," & str����ʱ�� & "," & str������� & ")"
        Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    If (intType = 2 Or intType = 3) And mstr��� = "E" Then
        Call InitObjLis(1278)
        If Not gobjLIS Is Nothing Then
            strTmp = ""
            Call gobjLIS.WriteInLisNotify(2, CStr(mlngҽ��ID), , strTmp)
            If strTmp <> "" Then MsgBox "zl9LisInsideComm����(WriteInLisNotify)��������" & vbCrLf & strTmp, vbInformation, gstrSysName
        End If
    End If
    SaveDisProcessData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckData() As Boolean
'���ܣ����������ȷ��
    Dim strMsg As String
    If mintType = 0 Then
        '����¼�����Ƽ���
        If txtInfo(txt���Ƽ���).Text = "" Then
            MsgBox "û��ȷ�����Ƽ�����", vbInformation, gstrSysName
            If txtInfo(txt���Ƽ���).Enabled Then txtInfo(txt���Ƽ���).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt�������).Text = "" Then
            MsgBox "û����д���������", vbInformation, gstrSysName
            If txtInfo(txt�������).Enabled Then txtInfo(txt�������).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt���ʱ��).Text = "" Then
            MsgBox "û����д���ʱ�䡣", vbInformation, gstrSysName
            Exit Function
        Else
            If Not Checkʱ��("���ʱ��", txtInfo(txt���ʱ��).Text, strMsg) Then
                MsgBox strMsg, vbInformation, gstrSysName
                If txtInfo(txt���ʱ��).Enabled Then txtInfo(txt���ʱ��).SetFocus
                Exit Function
            End If
        End If
        
        If txtInfo(txt�ͼ�ʱ��).Text <> "" Then
            If Not Checkʱ��("�ͼ�ʱ��", txtInfo(txt�ͼ�ʱ��).Text, strMsg) Then
                MsgBox strMsg, vbInformation, gstrSysName
                If txtInfo(txt�ͼ�ʱ��).Enabled Then txtInfo(txt�ͼ�ʱ��).SetFocus
                Exit Function
            End If
        End If
    End If
    CheckData = True
End Function

Private Sub Form_Resize()
On Error Resume Next
    If mblnDialog Then
        frmMain.Top = 800
        frmMain.Left = Me.ScaleLeft + (Me.ScaleWidth / 2) - (frmMain.Width / 2)
    Else
        frmMain.Top = mlngTop
        frmMain.Left = Me.ScaleLeft + (Me.ScaleWidth / 2) - (frmMain.Width / 2)
    
        If mblnSbSisible Then
            If Me.ScaleWidth < frmMain.Width Then
                hsbReport.Visible = True
            Else
                hsbReport.Visible = False
            End If
        
            If Me.ScaleHeight < frmMain.Height Then
                vsbReport.Visible = True
            Else
                vsbReport.Visible = False
            End If
            vsbReport.Top = Me.ScaleTop
            vsbReport.Left = Me.ScaleLeft + Me.ScaleWidth - vsbReport.Width
            vsbReport.Height = Me.ScaleHeight - IIf(hsbReport.Visible = True, hsbReport.Height, 0)
            vsbReport.LargeChange = 100 / ((frmMain.Height + 800) / Me.ScaleHeight)
            vsbReport.SmallChange = vsbReport.LargeChange
            
            hsbReport.Top = vsbReport.Top + vsbReport.Height
            hsbReport.Left = Me.ScaleLeft
            hsbReport.Width = Me.ScaleLeft + Me.ScaleWidth
            hsbReport.LargeChange = 100 / (frmMain.Width / Me.ScaleWidth)
            hsbReport.SmallChange = hsbReport.LargeChange
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mintType = 1 And Not mblnOk Then
        If Not SaveDisProcessData(5) Then
            Cancel = True
            Exit Sub
        End If
    End If
    If mblnDiagnose Then
        RaiseEvent Closed(0, "")
    End If
    If Not mclsDiagEdit Is Nothing Then Set mclsDiagEdit = Nothing
    If Not mclsMipModule Is Nothing Then Set mclsMipModule = Nothing
    If mintType = 1 Then
        glngOpenedID = 0
    End If
End Sub

Private Sub mclsDiagEdit_Closed(ByVal blnEditCancel As Boolean, ByVal str����ID As String, ByVal str���ID As String, ByVal strTag As String)
    Dim clsDisease As New cDockDisease
    Dim strName As String, strReason As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim frmDisStation As frmDiseaseStation
    Dim blnNotView As Boolean
    
    On Error GoTo errH
     If Not blnEditCancel Then
        strName = txtInfo(txt����).Text
        mblnDiagnose = True
        If str���ID = "" And str����ID = "" Then Exit Sub
        
        If InStr(";" & gobjComlib.GetPrivFunc(glngSys, 1249) & ";", ";������д;") <= 0 Then
            Exit Sub
        End If
        
        Set rsTemp = clsDisease.SatisfyEditDiseaseDoc(mlng����ID, mlng��ҳID, IIf(mint���� = 0, 2, 1), mlng����ID, str����ID, str���ID)
        
        If rsTemp Is Nothing Then
            Exit Sub
        ElseIf rsTemp.RecordCount = 0 Then
            Exit Sub          '�����ϼ�������ǰ�ᣬ�˳�
        End If
        If cboReport.ListCount > 1 Then
            If cboReport.ListIndex > 0 Then
                If MsgBox("�÷������Ѿ�������һ�ݼ������浥���Ƿ���д�µı��浥����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            Else
                 If MsgBox("�ò����Ѿ���д��" & "��" & mstr���ƴ�Ⱦ�� & "�������ı��浥���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    cboReport.ListIndex = 1
                    Call SaveDisProcessData(2)
                    Exit Sub
                End If
            End If
        End If
        Set frmDisStation = New frmDiseaseStation

        'û�в��ҵ�һ��֮�ڵ��ظ��ı��濨
        If Not frmDisStation.ShowDiseaseStation(Me, mlng����ID, IIf(mint���� = 0, mlng��ҳID, mlng�Һ�ID), IIf(mint���� = 0, 2, 1), _
                                    mlng����ID, str����ID, str���ID, blnNotView) Then
            Call clsDisease.EditDiseaseReport(Me, rsTemp, mlng����ID, IIf(mint���� = 0, mlng��ҳID, mlng�Һ�ID), IIf(mint���� = 0, 2, 1), mlng����ID, str����ID, str���ID, strReason)
            If strReason <> "" Then txtInfo(txt�������).Text = strReason
        ElseIf blnNotView Then
            Call clsDisease.EditDiseaseReport(Me, rsTemp, mlng����ID, IIf(mint���� = 0, mlng��ҳID, mlng�Һ�ID), IIf(mint���� = 0, 2, 1), mlng����ID, str����ID, str���ID, strReason)
            If strReason <> "" Then txtInfo(txt�������).Text = strReason
        End If
        Call SetCboReportData(mlng����ID, mstr���ƴ�Ⱦ��, mlngID)
        
        If Not frmDisStation Is Nothing Then
            Unload frmDisStation
            Set frmDisStation = Nothing
        End If
    End If
    Set clsDisease = Nothing
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SetCboReportData(ByVal lng����ID As Long, ByVal str���ƴ�Ⱦ�� As String, ByVal lngID As Long) As Boolean
'���ܣ���ѯ���Է����������ļ�������
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
On Error GoTo errH
    cboReport.Clear
    strSQL = "Select rowNum as NO,A.ID,B.ID as ������ID,A.����ʱ��,A.��������,B.��Ⱦ������  from ���Ӳ�����¼ A, �������Լ�¼ B where A.ID = B.�ļ�ID and A.����ID = B.����ID and B.����ID = [1] and B.��Ⱦ������ = [2]"
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ���Է����������ļ�������", lng����ID, str���ƴ�Ⱦ��)
    
    If rsTemp.RecordCount > 0 Then
        cboReport.AddItem ""
        cboReport.ItemData(cboReport.NewIndex) = 0
        cboReport.ListIndex = 0
        For i = 1 To rsTemp.RecordCount
            cboReport.AddItem rsTemp!NO & "-" & rsTemp!�������� & "(" & rsTemp!��Ⱦ������ & ")"
            cboReport.ItemData(cboReport.NewIndex) = rsTemp!ID
            If lngID = rsTemp!������ID Then
                cboReport.ListIndex = i
            End If
            rsTemp.MoveNext
        Next
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtInfo_Change(Index As Integer)
    Select Case Index
        Case txt�������
            mstr������� = txtInfo(txt�������).Text
    End Select
End Sub

Private Function Checkʱ��(ByVal strTimeType As String, ByVal strʱ�� As String, Optional ByRef strMsg As String) As Boolean
'���ܣ���������ʱ���Ƿ�Ϸ�
    Dim strInDate As String
    Dim datCurrent As Date
    
    datCurrent = gobjComlib.zlDatabase.Currentdate
    strInDate = mstr����ʱ��
    If Not IsDate(strʱ��) Then
        strMsg = "�����" & strTimeType & "��Ч��"
        Exit Function
    End If

    If mint���� = 0 Then
        If Format(strʱ��, "yyyy-MM-dd HH:mm") < Format(mstr����ʱ��, "yyyy-MM-dd HH:mm") Then
            strMsg = strTimeType & "����С�ڲ��˵���Ժʱ�� " & strInDate & " ��"
            Exit Function
        End If
    Else
        If Format(strʱ��, "yyyy-MM-dd HH:mm") < Format(mstr����ʱ��, "yyyy-MM-dd HH:mm") Then
            strMsg = strTimeType & "����С�ڲ��˵ľ���ʱ�� " & strInDate & " ��"
            Exit Function
        End If
    End If
    
    If Format(strʱ��, "yyyy-MM-dd HH:mm") > Format(datCurrent, "yyyy-MM-dd HH:mm") Then
         strMsg = strTimeType & "���ܴ��ڵ�ǰʱ�� " & Format(datCurrent, "yyyy-MM-dd HH:mm") & " ��"
         Exit Function
     End If
        
    Checkʱ�� = True
End Function

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
'�����¼�����ģ����
    If Asc("'") = KeyAscii Or Asc(";") = KeyAscii Or Asc("%") = KeyAscii Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        Select Case Index
            Case txt�ͼ����
                Call GetInspectDept(0)
            Case txt�ͼ�ҽ��
                Call GetInspectDoctor(0)
            Case txt���Ƽ���
                Call GetDiseaseList(0)
            Case txt�걾����
                Call GetSampleList(0)
        End Select
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Select Case Index
        Case txt�ͼ����
            Call GetInspectDept(1)
        Case txt�ͼ�ҽ��
            Call GetInspectDoctor(1)
        Case txt���Ƽ���
            Call GetDiseaseList(1)
        Case txt�걾����
            Call GetSampleList(1)
        End Select
    End If
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    Dim strMsg As String
    
    If mintType <> 0 Then
        Exit Sub
    End If
    Select Case Index
        Case txt�ͼ����, txt�ͼ�ҽ��, txt�걾����, txt���Ƽ���
            If txtInfo(Index).Text <> txtInfo(Index).Tag Then
                If txtInfo(Index).Text = "" Then
                    txtInfo(Index).Tag = ""
                    If Index = txt�ͼ���� Then
                        mlng�ͼ����ID = 0
                    ElseIf Index = txt�ͼ�ҽ�� Then
                        mstr�ͼ�ҽ�� = ""
                    ElseIf Index = txt�걾���� Then
                        mstr�걾���� = ""
                    ElseIf Index = txt���Ƽ��� Then
                        mstr���ƴ�Ⱦ�� = ""
                    End If
                Else
                    txtInfo(Index).Text = txtInfo(Index).Tag
                    If txtInfo(Index).Enabled Then
                        txtInfo(Index).SetFocus
                        gobjComlib.zlControl.TxtSelAll txtInfo(Index)
                    End If
                End If
            End If
            
        Case txt�ͼ�ʱ��, txt���ʱ��
            If Not IsDate(txtInfo(Index).Text) Then
                txtInfo(Index).Text = txtInfo(Index).Tag
            Else
                txtInfo(Index).Tag = txtInfo(Index).Text
                If Index = txt�ͼ�ʱ�� Then
                    mdat�ͼ�ʱ�� = Format(txtInfo(txt�ͼ�ʱ��).Text, "yyyy-MM-dd HH:mm")
                ElseIf Index = txt���ʱ�� Then
                    mdat���ʱ�� = Format(txtInfo(txt���ʱ��).Text, "yyyy-MM-dd HH:mm")
                End If
            End If
    End Select
End Sub

Private Sub SendMsg()
'���ܣ����� ��Ⱦ�����Խ�� ��Ϣ
    Dim strXML As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String

    On Error GoTo errH
    strXML = "<patient_info><patient_id>" & mlng����ID & "</patient_id><patient_name>" & txtInfo(txt����).Text & "</patient_name>"
    If mint���� = 0 Then
        strXML = strXML & "<in_number>" & txtInfo(txtסԺ��).Text & "</in_number>"
    Else
        strXML = strXML & "<out_number>" & txtInfo(txtסԺ��).Text & "</out_number>"
    End If
    strXML = strXML & "</patient_info><patient_clinic><patient_source>" & IIf(mint���� = 0, 2, 1) & "</patient_source>"
    strXML = strXML & "<clinic_id>" & IIf(mint���� = 0, mlng��ҳID, mlng�Һ�ID) & "</clinic_id>"
    If mlng����ID <> 0 Then
        strXML = strXML & "<clinic_area_id>" & mlng����ID & "</clinic_area_id>"
        strTmp = ""
        strTmp = gobjComlib.Sys.RowValue("���ű�", mlng����ID, "����")
        If strTmp <> "" Then
            strXML = strXML & "<clinic_area_title>" & strTmp & "</clinic_area_title>"
        End If
    End If
    strXML = strXML & "<clinic_dept_id>" & mlng����ID & "</clinic_dept_id>"
    strTmp = "" & gobjComlib.Sys.RowValue("���ű�", mlng����ID, "����")
    strXML = strXML & "<clinic_dept_title>" & strTmp & "</clinic_dept_title>"
    strXML = strXML & "<clinic_room>" & mstr���� & "</clinic_room>"
    If "" <> txtInfo(txt����).Text Then
        strXML = strXML & "<clinic_bed>" & strTmp & "</clinic_bed>"
    End If
    strXML = strXML & "</patient_clinic><positive_info><info_id>" & mlngID & "</info_id>"
    strXML = strXML & "<sample_name>" & mstr�걾���� & "</sample_name>"
    strXML = strXML & "<disease_name>" & mstr���ƴ�Ⱦ�� & "</disease_name>"
    strXML = strXML & "<create_time>" & Format(mdat�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "</create_time>"
    strXML = strXML & "<create_doctor>" & mstr�Ǽ��� & "</create_doctor>"
    strXML = strXML & "<create_dept_id>" & IIf(mlng�Ǽǿ���ID = 0, "NULL", mlng�Ǽǿ���ID) & "</create_dept_id>"

    strTmp = ""
    strTmp = "" & gobjComlib.Sys.RowValue("���ű�", IIf(mlng�Ǽǿ���ID = 0, "NULL", mlng�Ǽǿ���ID), "����")
    If strTmp <> "" Then
        strXML = strXML & "<create_dept>" & strTmp & "</create_dept>"
    End If

    strXML = strXML & "</positive_info>"

    If Not mclsMipModule Is Nothing Then
        If mclsMipModule.IsConnect Then Call mclsMipModule.CommitMessage("ZLHIS_CIS_032", strXML)
    End If

    Call gobjComlib.zlDatabase.SendMsg("ZLHIS_CIS_032", strXML)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub SetFrmInset(ByVal blnSbSisible As Boolean)
    mblnSbSisible = blnSbSisible
    Me.Appearance = 1
    Me.BackColor = &HC0C0C0
    Me.BorderStyle = 0
    Me.Caption = Me.Caption
    mblnDialog = False
End Sub

Public Sub zlRefresh(ByVal lngID As Long)
    mintType = 2
    mlngID = lngID
    Call LoadPatiInfo
    Call SetFormState(2)
End Sub

Private Sub hsbReport_Change()
    frmMain.Left = -((frmMain.Width - Me.Width) * (hsbReport.Value / 100))
End Sub

Public Sub SetReportTop(ByVal lngTop As Long)
     frmMain.Top = lngTop
     mlngTop = lngTop
End Sub

Private Sub vsbReport_Change()
    frmMain.Top = 200 - ((frmMain.Height + 800 - Me.Height) * (vsbReport.Value / 100))
End Sub

Private Sub ViewEPRReport()
'���ܣ����ı���
    Dim lng����ID As Long
    Dim str��鱨��ID As String
    Dim objPublicPACS As Object

    '���ж��Ƿ���Լ�������
    If mlngҽ��ID = 0 Then
        MsgBox "�÷�������Ӧ��ҽ��Ϊ�գ��޷��鿴�����鱨�棡", vbInformation, gstrSysName
        Exit Sub
    End If
    Select Case CheckEPRReport(mlngҽ��ID, lng����ID, str��鱨��ID)
    Case 0
        MsgBox "��ҽ���ı���û����д��", vbInformation, gstrSysName
        Exit Sub
    Case 2
        If InStr(gobjComlib.GetPrivFunc(glngSys, 1253), "����δ��ɱ���") > 0 Then
            MsgBox "ע�⣺��ҽ���ı��滹û����ʽǩ����", vbInformation, gstrSysName
        Else
            MsgBox "��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)����û��Ȩ�޲�����", vbInformation, gstrSysName
            Exit Sub
        End If
    End Select

    'ִ�в���
    '�°�PACS���棬ֱ��ǿ��ʹ���°�PACS����༭��
    If str��鱨��ID <> "" Then
        Call CreateObjectPacs(objPublicPACS)
        Call objPublicPACS.zlDocShowReport(mlngҽ��ID, , False, Me, True)
    Else
        Call gObjRichEPR.ViewDocument(Me, lng����ID)
        '���ı���
    End If
    If objPublicPACS Is Nothing Then Set objPublicPACS = Nothing
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function CreateObjectPacs(objPublicPACS As Object) As Boolean
    If objPublicPACS Is Nothing Then
        On Error Resume Next
        Set objPublicPACS = CreateObject("zlPublicPACS.clsPublicPACS")
        Err.Clear: On Error GoTo 0
        If Not objPublicPACS Is Nothing Then
            Call objPublicPACS.InitInterface(gcnOracle, UserInfo.�û���)
        End If
        If objPublicPACS Is Nothing Then
            MsgBox "PACS��������δ�����ɹ���", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreateObjectPacs = True
End Function


Private Function CheckEPRReport(ByVal lngҽ��id As Long, ByRef lng����ID As Long, ByRef str��鱨��ID As String) As Integer
'���ܣ�����Ӧ��Ŀ�ı�����д���
'������lngҽ��ID=�ɼ��е�ҽ��ID
'      lng����ID=���Դ��룬��Ҫ���ڷ��ر��没��ID
'      intִ��״̬=���ڼ������ʱ�������ۺϵ�ִ��״̬
'���أ�0-���滹û����д
'      1-��������д���(��ǩ��,�����޶���ǩ��,����ִ�����)
'      2-����δ��д���(δǩ��,���޶���δǩ��,��δִ�����)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '��鱨���Ƿ�����д
    strSQL = "Select ����ID,��鱨��ID || ''  as ��鱨��ID From ����ҽ������ Where ҽ��ID=[1]"
    
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lngҽ��id)
    If Not rsTmp.EOF Then lng����ID = Val(rsTmp!����id & ""): str��鱨��ID = rsTmp!��鱨��ID & ""
    If lng����ID = 0 And str��鱨��ID = "" Then
        CheckEPRReport = 0: Exit Function
    End If
    
    '��鱨��ִ�й���(5-���;6-�������)��״̬(1-���)
    '���鱨���ǹ������ɼ���ʽ����ģ����ɼ���ʽ����Ϊ����δ�������ͼ�¼
    strSQL = _
        " Select 2 as ����,ҽ��ID,ִ�й���,ִ��״̬,����ʱ�� From ����ҽ������ Where ҽ��ID=[1]" & _
        " Union ALL" & _
        " Select ����,ҽ��ID,ִ�й���, ִ��״̬,����ʱ��" & _
        " From (" & _
            " Select 1 as ����,B.ҽ��ID,B.ִ�й���,B.ִ��״̬,B.����ʱ�� From ����ҽ����¼ A,����ҽ������ B" & _
            " Where A.ID=B.ҽ��ID And A.���ID=(" & _
                " Select A.ID From ����ҽ����¼ A,������ĿĿ¼ B Where A.ID=[1] And A.������ĿID=B.ID And A.�������='E' And B.��������='6')" & _
            " Order by A.���" & _
        " ) Where Rownum=1" & _
        " Order by ����,����ʱ�� Desc"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lngҽ��id, 0)
    If NVL(rsTmp!ִ�й���, 0) >= 5 Or NVL(rsTmp!ִ��״̬, 0) = 1 Then
        CheckEPRReport = 1
    Else
        CheckEPRReport = 2
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitObjLis(ByVal lngProgram As Long)
'�ж�����°�LIS����Ϊ�վͳ�ʼ��
    Dim strErr As String
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, lngProgram, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS������ʼ������" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub
