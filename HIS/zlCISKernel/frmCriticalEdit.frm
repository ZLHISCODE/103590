VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCriticalEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Σ��ֵ�Ǽǵ�"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11370
   Icon            =   "frmCriticalEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComCtl2.MonthView dtpDate 
      Height          =   2220
      Left            =   9570
      TabIndex        =   40
      Top             =   8880
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   117702657
      TitleBackColor  =   -2147483636
      TitleForeColor  =   -2147483634
      TrailingForeColor=   -2147483637
      CurrentDate     =   37904
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7665
      Left            =   450
      ScaleHeight     =   7665
      ScaleWidth      =   10500
      TabIndex        =   0
      Top             =   660
      Width           =   10500
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
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "235"
         Top             =   1785
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
         Index           =   0
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "����"
         Top             =   1230
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
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "��"
         Top             =   1230
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
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "�ڿ�"
         Top             =   1785
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
         Left            =   8700
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "6"
         Top             =   1785
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
         Left            =   8700
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "28��"
         Top             =   1230
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
         Index           =   6
         Left            =   1260
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "2017-10-11 12:00"
         Top             =   4080
         Width           =   1725
      End
      Begin VB.PictureBox picCL 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   -30
         ScaleHeight     =   2895
         ScaleWidth      =   10455
         TabIndex        =   3
         Top             =   4530
         Width           =   10455
         Begin VB.CommandButton cmdSel 
            Caption         =   "��"
            Height          =   270
            Index           =   13
            Left            =   9705
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   2250
            Width           =   270
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "��"
            Height          =   270
            Index           =   12
            Left            =   6570
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   2265
            Width           =   270
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
            Left            =   1260
            TabIndex        =   8
            Text            =   "2013-06-20 18:00"
            Top             =   2280
            Width           =   1695
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
            Index           =   10
            Left            =   1275
            MaxLength       =   4000
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   135
            Width           =   8685
         End
         Begin VB.PictureBox picInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Index           =   0
            Left            =   1200
            ScaleHeight     =   390
            ScaleWidth      =   1965
            TabIndex        =   4
            Top             =   1665
            Width           =   1965
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "��"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   75
               TabIndex        =   6
               Top             =   60
               Value           =   -1  'True
               Width           =   600
            End
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "��"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   825
               TabIndex        =   5
               Top             =   75
               Width           =   600
            End
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
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
            Left            =   8775
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Text            =   "����"
            Top             =   2280
            Width           =   1125
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
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
            Left            =   5505
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "�ڿ�"
            Top             =   2280
            Width           =   1290
         End
         Begin VB.Line Line2 
            BorderStyle     =   2  'Dash
            X1              =   -105
            X2              =   11055
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "ȷ��ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   11
            Left            =   315
            TabIndex        =   17
            Top             =   2295
            Width           =   840
         End
         Begin VB.Line Line1 
            Index           =   11
            X1              =   1260
            X2              =   2910
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Line Line1 
            Index           =   12
            X1              =   5490
            X2              =   6795
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "ȷ�Ͽ���"
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
            Left            =   4560
            TabIndex        =   16
            Top             =   2295
            Width           =   840
         End
         Begin VB.Line Line1 
            Index           =   13
            X1              =   8700
            X2              =   9960
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "ȷ �� ��"
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
            Left            =   7830
            TabIndex        =   15
            Top             =   2295
            Width           =   840
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   10
            Left            =   315
            TabIndex        =   14
            Top             =   150
            Width           =   900
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "�Ƿ���Σֵ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   14
            Left            =   105
            TabIndex        =   13
            Top             =   1725
            Width           =   1050
         End
         Begin VB.Image imgDate 
            Height          =   240
            Index           =   11
            Left            =   3015
            Picture         =   "frmCriticalEdit.frx":6852
            Top             =   2280
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   270
         Index           =   7
         Left            =   6480
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4050
         Width           =   270
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   270
         Index           =   8
         Left            =   9660
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4050
         Width           =   270
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
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "B����"
         Top             =   4080
         Width           =   1185
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
         Left            =   8700
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "������"
         Top             =   4080
         Width           =   1230
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
         Index           =   9
         Left            =   1290
         MaxLength       =   4000
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   2385
         Width           =   8685
      End
      Begin VB.Label lblHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Σ��ֵ�Ǽǵ�"
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
         Left            =   4065
         TabIndex        =   39
         Top             =   120
         Width           =   2250
      End
      Begin VB.Line linHead 
         X1              =   4065
         X2              =   6315
         Y1              =   495
         Y2              =   495
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
         Index           =   1
         Left            =   4560
         TabIndex        =   38
         Top             =   1230
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
         Left            =   7785
         TabIndex        =   37
         Top             =   1230
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
         Index           =   4
         Left            =   4560
         TabIndex        =   36
         Top             =   1785
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
         Index           =   0
         Left            =   360
         TabIndex        =   35
         Top             =   1230
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
         Left            =   360
         TabIndex        =   34
         Top             =   1785
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
         Index           =   5
         Left            =   7785
         TabIndex        =   33
         Top             =   1785
         Width           =   840
      End
      Begin VB.Line linL 
         Index           =   4
         X1              =   5505
         X2              =   6735
         Y1              =   2025
         Y2              =   2025
      End
      Begin VB.Line linL 
         Index           =   0
         X1              =   1275
         X2              =   2505
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line linL 
         Index           =   3
         X1              =   1275
         X2              =   2505
         Y1              =   2025
         Y2              =   2025
      End
      Begin VB.Line linL 
         Index           =   1
         X1              =   5505
         X2              =   6735
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line linL 
         Index           =   2
         X1              =   8700
         X2              =   9930
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line linL 
         Index           =   5
         X1              =   8700
         X2              =   9930
         Y1              =   2025
         Y2              =   2025
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
         Index           =   6
         Left            =   360
         TabIndex        =   32
         Top             =   4065
         Width           =   840
      End
      Begin VB.Line linL 
         Index           =   6
         X1              =   1275
         X2              =   2940
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Σ��ֵ����"
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
         Index           =   9
         Left            =   105
         TabIndex        =   31
         Top             =   2370
         Width           =   1125
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� ��"
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
         Left            =   7785
         TabIndex        =   30
         Top             =   4065
         Width           =   840
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   7
         Left            =   4560
         TabIndex        =   29
         Top             =   4065
         Width           =   840
      End
      Begin VB.Line linL 
         Index           =   8
         X1              =   8700
         X2              =   9930
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line linL 
         Index           =   7
         X1              =   5505
         X2              =   6735
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label lblType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   7785
         TabIndex        =   28
         Top             =   510
         Width           =   900
      End
      Begin VB.Image imgDate 
         Height          =   240
         Index           =   6
         Left            =   2985
         Picture         =   "frmCriticalEdit.frx":D0A4
         Top             =   4080
         Width           =   240
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   525
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmCriticalEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum CtlID
    txt���� = 0
    txt�Ա� = 1
    txt���� = 2
    txt���� = 3
    txtסԺ�� = 4
    txt���� = 5
    txt����ʱ�� = 6
    txt������� = 7
    txt������ = 8
     
    txtΣ��ֵ���� = 9
    txt������� = 10
    txtȷ��ʱ�� = 11
    txtȷ�Ͽ��� = 12
    txtȷ���� = 13
End Enum

Private mclsMipModule As zl9ComLib.clsMipModule '��Ϣ����
Private mblnModal As Boolean '��ʾ��ʽ��ģ̬����ģ̬
Private mfrmParent As Object '�����ڶ���
Private mint�������� As Integer  '1-����,2-סԺ,3-������Դ����
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mstr�Һŵ� As String
Private mlngҽ��ID As Long
Private mintType As Integer '0-������1-�޸ģ�2-�鿴��3-����
Private mstr����ʱ�� As String
Private mlngΣ��ֵID As Long
Private mrsΣ��ֵ As ADODB.Recordset
Private mlng�걾ID As Long
Private mstrΣ��ֵ���� As String
Private mdat����ʱ�� As Date
Private mlng�������ID As Long
Private mstr������ As String
Private mstr������� As String
Private mdatȷ��ʱ�� As Date
Private mstrȷ���� As String
Private mlngȷ�Ͽ���ID As Long
Private mobjReport As Object
Private mlngӤ�� As Long
Private mblnOK As Boolean
 
Private mblnChange As Boolean

Public Function ShowMe(frmParent As Object, ByVal blnModal As Boolean, ByVal intType As Integer, ByVal int�������� As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String, ByVal lngӤ�� As Long, ByRef lngΣ��ֵID As Long, ByVal lngҽ��ID As Long, _
    Optional ByVal lng�걾id As Long, Optional ByVal strΣ��ֵ���� As String, Optional ByVal dat����ʱ�� As Date, Optional ByVal lng�������ID As Long, Optional ByVal str������ As String, Optional ByRef objMip As Object) As Boolean
'���ܣ���ʾ���壬����һ��Σ��ֵ
'���أ�true �ύ�����ݣ� false δ����
'������frmParent ������
'      blnModal ��ʾ��ʽ��1-ģ̬��0-��ģ̬
'      intType  0-������1-�޸ģ�2-�鿴��3-����
'      int�������� 1-���ﲡ��,2-סԺ����,3-��������
'      lng����ID,lng��ҳID,str�Һŵ�,�������
'      lngΣ��ֵID ��ǰ��¼ID,�ɷ���
'      lngҽ��ID ��Ӧ��ҽ����Ŀ
'      lng�걾ID LIS�����Ǵ���

'      strΣ��ֵ���� ����ʱ����ȱʡֵ
'      dat����ʱ��   ����ʱ����ȱʡֵ
'      lng�������ID ����ʱ����ȱʡֵ
'      str������     ����ʱ����ȱʡֵ

'      objMip ���ڷ�����Ϣ�Ķ��� zl9ComLib.clsMipModule

    Set mfrmParent = frmParent
    mblnModal = blnModal
    mintType = intType
    mint�������� = int��������
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstr�Һŵ� = str�Һŵ�
    
    If mlng��ҳID = 0 And mstr�Һŵ� = "" Then
        mint�������� = 3
    End If
    
    mlngӤ�� = lngӤ��
    mlngΣ��ֵID = lngΣ��ֵID
    mlngҽ��ID = lngҽ��ID
    mlng�걾ID = lng�걾id
    mstrΣ��ֵ���� = strΣ��ֵ����
    mdat����ʱ�� = dat����ʱ��
    mlng�������ID = lng�������ID
    mstr������ = str������
    
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    
    Me.Show IIF(blnModal, 1, 0), frmParent
    
    lngΣ��ֵID = mlngΣ��ֵID
    
    ShowMe = mblnOK
    
End Function

Public Function ShowApp(frmParent As Object, ByVal lngΣ��ֵID As Long)
'���ܣ��鿴Σ��ֵ��¼
    Set mfrmParent = frmParent
    mlngΣ��ֵID = lngΣ��ֵID
    mintType = 2
    Call InitBaseBy��¼ID(lngΣ��ֵID)
    Me.Show 1, frmParent
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_PrintSet '��ӡ����
             Call ReportPrintSet(gcnOracle, glngSys, "ZL1_INSIDE_1254_20", Me)
        Case conMenu_File_Preview 'Ԥ��
            Call PrintApply(1)
        Case conMenu_File_Print '��ӡ
            Call PrintApply(2)
        Case conMenu_Edit_Save
            If CheckData() Then
                Call SaveData
            End If
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnVisible As Boolean
    
    blnVisible = True
    Select Case Control.ID
        Case conMenu_Edit_Save
            Control.Enabled = mblnChange
            If mintType = 2 Then
                blnVisible = False
            End If
    End Select
    Control.Visible = blnVisible
End Sub

Private Sub cmdSel_Click(Index As Integer)
    Select Case Index
    Case txt�������
    Case txt������
        Call GetItem������(1)
    Case txtȷ�Ͽ���
    Case txtȷ����
    End Select
End Sub

 
Private Sub Form_Load()
    mblnOK = False
    Call InitCommandBar
    Call LoadBaseInfo
    If mintType <> 0 Then
        Call LoadInputData
    End If
    Call SetFaceCtrl
    mblnChange = False
End Sub

Private Sub SetFaceCtrl()
    Select Case mintType
    Case 0
        txtInfo(txt����ʱ��).Locked = False
        txtInfo(txt�������).Locked = True
        txtInfo(txt������).Locked = False
        txtInfo(txtΣ��ֵ����).Locked = False
        
        cmdSel(txt�������).Visible = False
        
        picCL.Visible = False
    Case 1
        txtInfo(txt����ʱ��).Locked = False
        txtInfo(txt�������).Locked = True
        txtInfo(txt������).Locked = False
        txtInfo(txtΣ��ֵ����).Locked = False
        
        cmdSel(txt�������).Visible = False
        
        picCL.Visible = False
    Case 2
        txtInfo(txt����ʱ��).Locked = True
        txtInfo(txt�������).Locked = True
        txtInfo(txt������).Locked = True
        txtInfo(txtΣ��ֵ����).Locked = True
        imgDate(txt����ʱ��).Visible = False
        cmdSel(txt�������).Visible = False
        cmdSel(txt������).Visible = False
        
        
        txtInfo(txt�������).Locked = True
        txtInfo(txtȷ��ʱ��).Locked = True
        txtInfo(txtȷ�Ͽ���).Locked = True
        txtInfo(txtȷ����).Locked = True
        imgDate(txtȷ��ʱ��).Visible = False
        cmdSel(txtȷ�Ͽ���).Visible = False
        cmdSel(txtȷ����).Visible = False
        
        picInfo(0).Enabled = False
        
        '���ҽ��δ��������ʾ���沿��
        If txtInfo(txt�������).Text = "" Then
            picCL.Visible = False
        Else
            picCL.Visible = True
        End If
    Case 3
        txtInfo(txt����ʱ��).Locked = True
        txtInfo(txt�������).Locked = True
        txtInfo(txt������).Locked = True
        txtInfo(txtΣ��ֵ����).Locked = True
        imgDate(txt����ʱ��).Visible = False
        cmdSel(txt�������).Visible = False
        cmdSel(txt������).Visible = False
        cmdSel(txtȷ�Ͽ���).Visible = False
        cmdSel(txtȷ����).Visible = False
        
        txtInfo(txt�������).Locked = False
        txtInfo(txtȷ��ʱ��).Locked = False
        txtInfo(txtȷ�Ͽ���).Locked = True
        txtInfo(txtȷ����).Locked = True
        picCL.Visible = True
    End Select
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    picMain.Top = 525
    picMain.Left = Me.ScaleLeft + (Me.ScaleWidth / 2) - (picMain.Width / 2)
    picCL.Left = 0
    picCL.Width = picMain.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjReport Is Nothing Then Set mobjReport = Nothing
End Sub

Private Sub optInfo_Click(Index As Integer)
    mblnChange = True
End Sub

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
    Set cbsMain.Icons = zlCommFun.GetPubIcons

    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����")
            objControl.IconId = 815
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, " ����")
            objControl.BeginGroup = True
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

Private Sub LoadBaseInfo()
'���ܣ����س�ʼ����
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim i As Long
    
    '���
    For i = 0 To txt����ʱ��
        txtInfo(i).Text = ""
    Next
    
    txtInfo(txt������).Text = ""
    txtInfo(txt�������).Text = ""
    txtInfo(txt����ʱ��).Text = ""
    txtInfo(txtΣ��ֵ����).Text = ""
    txtInfo(txt�������).Text = ""
    txtInfo(txtȷ��ʱ��).Text = ""
    txtInfo(txtȷ�Ͽ���).Text = ""
    txtInfo(txtȷ����).Text = ""
    
    On Error GoTo errH
    
    If mint�������� = 1 Then
        strSql = "select a.����, a.�Ա�, a.����,b.���� as ����,a.����� as סԺ��,null as ����,a.����ʱ�� as ����ʱ��,b.id as ����ID,a.���� from ���˹Һż�¼ a,���ű� b where a.ִ�в���id=b.id and a.no=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, mstr�Һŵ�)
        lblInfo(txtסԺ��).Caption = "�� �� ��"
        lblInfo(txt����).Caption = "��    ��"
        txtInfo(txt����).Text = IIF(Val(rsTmp!���� & "") = 1, "��", "��")
    ElseIf mint�������� = 2 Then
        strSql = "Select a.����, a.�Ա�, a.����, b.���� As ����, a.סԺ��, a.��Ժ���� As ����,a.��Ժ���� as ����ʱ��,b.id as ����ID" & vbNewLine & _
            "From ������ҳ A, ���ű� B Where a.��Ժ����id = b.Id And a.����id = [1] And a.��ҳid = [2]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mlng��ҳID)
        txtInfo(txt����).Text = rsTmp!���� & ""
    ElseIf mint�������� = 3 Then
        strSql = "select  a.����, a.�Ա�, a.���� ,b.���� as ����,null as סԺ��,null as ����,a.��ʼִ��ʱ�� as ����ʱ��,b.id as ����ID,null as ����  from ����ҽ����¼ a,���ű� b where a.���˿���id=b.id and a.id=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, mlngҽ��ID)
        lblInfo(txtסԺ��).Caption = "�� �� ��"
        lblInfo(txt����).Caption = "��    ��"
        txtInfo(txt����).Text = IIF(Val(rsTmp!���� & "") = 1, "��", "��")
    End If
    
    txtInfo(txt����).Text = rsTmp!���� & ""
    txtInfo(txt�Ա�).Text = rsTmp!�Ա� & ""
    txtInfo(txt����).Text = rsTmp!���� & ""
    txtInfo(txt����).Text = rsTmp!���� & ""
        txtInfo(txt����).Tag = Val(rsTmp!����ID & "")
    txtInfo(txtסԺ��).Text = rsTmp!סԺ�� & ""
    
    mstr����ʱ�� = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:mm")
     
    
    '����
    strSql = "select b.����,b.id as �Ǽǿ���ID,a.������� from ����ҽ����¼ a,���ű� b where a.ִ�п���id=b.id and a.id=[1]"
    Set rsAdvice = zldatabase.OpenSQLRecord(strSql, Me.Caption, mlngҽ��ID)
    If Not rsAdvice.EOF Then
        mlng�������ID = Val(rsAdvice!�Ǽǿ���id & "")
        txtInfo(txt�������).Text = rsAdvice!���� & ""
        If rsAdvice!������� & "" = "D" Then
            lblType.Caption = "�����"
        Else
            lblType.Caption = "������"
        End If
    End If
         
    If mintType = 0 Then
        txtInfo(txt������).Text = UserInfo.����
        txtInfo(txt������).Tag = UserInfo.����
        
        txtInfo(txt����ʱ��).Text = Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm")
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtInfo(Index)
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 39 Then KeyAscii = 0 '�����ű���
    If KeyAscii = 13 Then
        Select Case Index
        Case txt������
            Call GetItem������(0)
        End Select
    End If
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case txt����ʱ��
            If Not IsDate(txtInfo(Index).Text) Then
                txtInfo(Index).Text = txtInfo(Index).Tag
            Else
                txtInfo(Index).Tag = txtInfo(Index).Text
            End If
    End Select
End Sub

Private Sub imgDate_Click(Index As Integer)
    Select Case Index
        Case txt����ʱ��
            If IsDate(txtInfo(txt����ʱ��).Text) Then
                dtpDate.value = CDate(txtInfo(txt����ʱ��).Text)
            Else
                dtpDate.value = zldatabase.Currentdate
            End If
            dtpDate.Tag = "����ʱ��"
            dtpDate.Left = txtInfo(txt����ʱ��).Left + picMain.Left
            dtpDate.Top = txtInfo(txt����ʱ��).Top + txtInfo(txt����ʱ��).Height + picMain.Top + 20
            dtpDate.Visible = True
            dtpDate.SetFocus
        Case txtȷ��ʱ��
            If IsDate(txtInfo(Index).Text) Then
                dtpDate.value = CDate(txtInfo(Index).Text)
            Else
                dtpDate.value = zldatabase.Currentdate
            End If
            dtpDate.Tag = "ȷ��ʱ��"
            dtpDate.Left = txtInfo(Index).Left + picMain.Left
            dtpDate.Top = txtInfo(Index).Top + txtInfo(Index).Height + picCL.Top + picMain.Top + 20
            dtpDate.Visible = True
            dtpDate.SetFocus
    End Select
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String
    
    If dtpDate.Tag = "����ʱ��" Then
        'ȡֵ
        If IsDate(txtInfo(txt����ʱ��).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(txt����ʱ��).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        txtInfo(txt����ʱ��).Text = strDate
        txtInfo(txt����ʱ��).Tag = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        txtInfo(txt����ʱ��).SetFocus
    ElseIf dtpDate.Tag = "ȷ��ʱ��" Then
        'ȡֵ
        If IsDate(txtInfo(txtȷ��ʱ��).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(txtȷ��ʱ��).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        txtInfo(txtȷ��ʱ��).Text = strDate
        txtInfo(txtȷ��ʱ��).Tag = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        txtInfo(txtȷ��ʱ��).SetFocus
    End If
End Sub

Private Function CheckData() As Boolean
'���ܣ����������ȷ��
    Dim strMsg As String
    Dim strTmp As String
    
    If mintType = 0 Or mintType = 1 Then
        If txtInfo(txtΣ��ֵ����).Text = "" Then
            MsgBox "û����д Σ��ֵ���� ��", vbInformation, gstrSysName
            If txtInfo(txtΣ��ֵ����).Enabled Then txtInfo(txtΣ��ֵ����).SetFocus
            Exit Function
        End If
        
        strTmp = txtInfo(txtΣ��ֵ����).Text
        If zlCommFun.ActualLen(strTmp) > txtInfo(txtΣ��ֵ����).MaxLength Then
            strMsg = "Σ��ֵ����-����̫��(����¼��" & txtInfo(txtΣ��ֵ����).MaxLength & "���ַ���" & txtInfo(txtΣ��ֵ����).MaxLength \ 2 & "������)��"
            MsgBox strMsg, vbInformation, gstrSysName
            If txtInfo(txtΣ��ֵ����).Enabled Then txtInfo(txtΣ��ֵ����).SetFocus
            Exit Function
        End If
                
        If txtInfo(txt����ʱ��).Text = "" Then
            MsgBox "û����д ����ʱ�� ��", vbInformation, gstrSysName
            If txtInfo(txt����ʱ��).Enabled Then txtInfo(txt����ʱ��).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt����ʱ��).Text <> "" Then
            If Not Checkʱ��("����ʱ��", txtInfo(txt����ʱ��).Text, strMsg) Then
                MsgBox strMsg, vbInformation, gstrSysName
                If txtInfo(txt����ʱ��).Enabled Then txtInfo(txt����ʱ��).SetFocus
                Exit Function
            End If
        End If
    End If
    
    If mintType = 3 Then
        If txtInfo(txt�������).Text = "" Then
            MsgBox "û����д ������� ��", vbInformation, gstrSysName
            If txtInfo(txt�������).Enabled Then txtInfo(txt�������).SetFocus
            Exit Function
        End If
        
        strTmp = txtInfo(txt�������).Text
        If zlCommFun.ActualLen(strTmp) > txtInfo(txt�������).MaxLength Then
            strMsg = "�������-����̫��(����¼��" & txtInfo(txt�������).MaxLength & "���ַ���" & txtInfo(txt�������).MaxLength \ 2 & "������)��"
            MsgBox strMsg, vbInformation, gstrSysName
            If txtInfo(txt�������).Enabled Then txtInfo(txt�������).SetFocus
            Exit Function
        End If
        
        If txtInfo(txtȷ��ʱ��).Text = "" Then
            MsgBox "û����д ȷ��ʱ�� ��", vbInformation, gstrSysName
            If txtInfo(txtȷ��ʱ��).Enabled Then txtInfo(txtȷ��ʱ��).SetFocus
            Exit Function
        End If
        
        If txtInfo(txtȷ��ʱ��).Text <> "" Then
            If Not Checkʱ��("ȷ��ʱ��", txtInfo(txtȷ��ʱ��).Text, strMsg) Then
                MsgBox strMsg, vbInformation, gstrSysName
                If txtInfo(txtȷ��ʱ��).Enabled Then txtInfo(txtȷ��ʱ��).SetFocus
                Exit Function
            End If
        End If
        
    End If
    
    CheckData = True
End Function

Private Function Checkʱ��(ByVal strTimeType As String, ByVal strʱ�� As String, Optional ByRef strMsg As String) As Boolean
'���ܣ���������ʱ���Ƿ�Ϸ�
    Dim strInDate As String
    Dim datCurrent As Date
    
    If Not IsDate(strʱ��) Then
        strMsg = "�����" & strTimeType & "��Ч��"
        Exit Function
    End If
    
    If "����ʱ��" = strTimeType Then
        datCurrent = zldatabase.Currentdate
        strInDate = mstr����ʱ��
        
    
        If Format(strʱ��, "yyyy-MM-dd HH:mm") < Format(mstr����ʱ��, "yyyy-MM-dd HH:mm") Then
            strMsg = strTimeType & "����С�ڲ��˵���Ժʱ�� " & strInDate & " ��"
            Exit Function
        End If
       
        If Format(strʱ��, "yyyy-MM-dd HH:mm") > Format(datCurrent, "yyyy-MM-dd HH:mm") Then
            strMsg = strTimeType & "���ܴ��ڵ�ǰʱ�� " & Format(datCurrent, "yyyy-MM-dd HH:mm") & " ��"
            Exit Function
        End If
    ElseIf "ȷ��ʱ��" = strTimeType Then
        'ȷ��ʱ��Ӧ�ô��ڱ���ʱ��
        If Format(strʱ��, "yyyy-MM-dd HH:mm") < txtInfo(txt����ʱ��).Text Then
            strMsg = strTimeType & "����С�ڱ���ʱ�� " & strInDate & " ��"
            Exit Function
        End If
    End If
    Checkʱ�� = True
End Function

Private Function SaveData() As Boolean
'���ܣ���������
    Dim strSql As String
    Dim strPars As String
    Dim int�Ƿ�Σ��ֵ As Integer
    
    On Error GoTo errH
    
    mstrΣ��ֵ���� = txtInfo(txtΣ��ֵ����).Text
    mdat����ʱ�� = CDate(txtInfo(txt����ʱ��).Text)
    mstr������ = txtInfo(txt������).Text
    
    If mintType = 0 Then
        mlngΣ��ֵID = zldatabase.GetNextID("����Σ��ֵ��¼")        '��ȡΣ��ֵ��¼ID
    End If
    
    strPars = "(" & mlngΣ��ֵID & ",null," & mlng����ID & "," & ZVal(mlng��ҳID) & "," & IIF(mstr�Һŵ� = "", "null", "'" & mstr�Һŵ� & "'") & "," & mlngӤ�� & ","
    strPars = strPars & "'" & txtInfo(txt����).Text & "','" & txtInfo(txt�Ա�).Text & "','" & txtInfo(txt����).Text & "'," & mlngҽ��ID & "," & mlng�걾ID & ",'" & mstrΣ��ֵ���� & "',"
    strPars = strPars & "to_date('" & Format(mdat����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),"
    strPars = strPars & mlng�������ID & ",'" & mstr������ & "')"
    
    If mintType = 0 Then
        strSql = "Zl_����Σ��ֵ��¼_Insert" & strPars
        Call zldatabase.ExecuteProcedure(strSql, Me.Caption)
        If lblType.Caption = "�����" Then
            Call SendΣ��ֵ��Ϣ(0)
        Else
            Call SendΣ��ֵ��Ϣ(1)
        End If
        '����״̬������Ϊ�޸�״̬
        mintType = 1
    ElseIf mintType = 1 Then
        strSql = "Zl_����Σ��ֵ��¼_Update" & strPars
        Call zldatabase.ExecuteProcedure(strSql, Me.Caption)
    ElseIf mintType = 3 Then
        mdatȷ��ʱ�� = CDate(txtInfo(txtȷ��ʱ��).Text)
        mstr������� = txtInfo(txt�������).Text
        mstrȷ���� = txtInfo(txtȷ����).Text
        If optInfo(0).value Then
            int�Ƿ�Σ��ֵ = 1
        Else
            int�Ƿ�Σ��ֵ = 0
        End If
        strSql = "Zl_����Σ��ֵ��¼_����(" & mlngΣ��ֵID & ",'" & mstr������� & "',to_date('" & Format(mdatȷ��ʱ��, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),'" & mstrȷ���� & "'," & mlngȷ�Ͽ���ID & "," & int�Ƿ�Σ��ֵ & ")"
        Call zldatabase.ExecuteProcedure(strSql, Me.Caption)
    End If
    mblnOK = True
    mblnChange = False
    SaveData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitBaseBy��¼ID(ByVal lngIdin As Long)
'���ܣ�ͨ��Σ��ֵ��ʼ������Ϣ
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
     
    On Error GoTo errH
    mlng����ID = 4
    strSql = "Select a.Id, a.������Դ, a.����id, a.��ҳid, a.�Һŵ�, a.Ӥ��, a.����, a.�Ա�, a.����, a.ҽ��id, a.�걾id, a.Σ��ֵ����, a.����ʱ��, a.�������id, a.������,a.�������, a.ȷ��ʱ��, a.ȷ����, a.ȷ�Ͽ���id, a.״̬, a.�Ƿ�Σ��ֵ " & _
        " From ����Σ��ֵ��¼ A where a.id=[1]"
    Set mrsΣ��ֵ = zldatabase.OpenSQLRecord(strSql, Me.Caption, lngIdin)
    
    mlng����ID = Val(mrsΣ��ֵ!����ID & "")
    mlng��ҳID = Val(mrsΣ��ֵ!��ҳID & "")
    mstr�Һŵ� = mrsΣ��ֵ!�Һŵ� & ""
    mint�������� = IIF(mstr�Һŵ� = "", 2, 1)
    mlngҽ��ID = Val(mrsΣ��ֵ!ҽ��ID & "")
    mlngӤ�� = Val(mrsΣ��ֵ!Ӥ�� & "")
    
    mlng�걾ID = Val(mrsΣ��ֵ!�걾ID & "")
    mstrΣ��ֵ���� = Val(mrsΣ��ֵ!Σ��ֵ���� & "")
    
    If Not IsNull(mrsΣ��ֵ!����ʱ��) Then
        mdat����ʱ�� = Format(mrsΣ��ֵ!����ʱ�� & "", "yyyy-MM-dd HH:mm:ss")
    End If
    
    mlng�������ID = Val(mrsΣ��ֵ!�������id & "")
    mstr������ = Val(mrsΣ��ֵ!������ & "")
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtpDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        If dtpDate.Tag = "����ʱ��" Then
            txtInfo(txt����ʱ��).SetFocus
        ElseIf dtpDate.Tag = "ȷ��ʱ��" Then
            txtInfo(txtȷ��ʱ��).SetFocus
        End If
        dtpDate.Tag = ""
        dtpDate.Visible = False
    End If
End Sub

Private Sub LoadInputData()
'���ܣ��鿴ʱ���ؽ�������ʱ��д����Ϣ
    
    Call InitBaseBy��¼ID(mlngΣ��ֵID)
    
    txtInfo(txt����ʱ��).Text = Format(mrsΣ��ֵ!����ʱ��, "yyyy-MM-dd HH:mm")
    txtInfo(txtΣ��ֵ����).Text = mrsΣ��ֵ!Σ��ֵ���� & ""
    txtInfo(txt�������).Text = mrsΣ��ֵ!������� & ""
    txtInfo(txt������).Text = mrsΣ��ֵ!������ & ""
    txtInfo(txt�������).Text = Sys.RowValue("���ű�", Val(mrsΣ��ֵ!�������id & ""), "����")
    txtInfo(txt����ʱ��).Text = Format(mrsΣ��ֵ!����ʱ��, "yyyy-MM-dd HH:mm")
    
    
    Select Case mintType
    Case 0
    Case 1
    Case 2
        If txtInfo(txt�������).Text <> "" Then
            txtInfo(txtȷ��ʱ��).Text = Format(mrsΣ��ֵ!ȷ��ʱ��, "yyyy-MM-dd HH:mm")
            txtInfo(txtȷ����).Text = mrsΣ��ֵ!ȷ���� & ""
            mlngȷ�Ͽ���ID = Val(mrsΣ��ֵ!ȷ�Ͽ���ID & "")
            txtInfo(txtȷ�Ͽ���).Text = Sys.RowValue("���ű�", mlngȷ�Ͽ���ID, "����")
            If Val(mrsΣ��ֵ!�Ƿ�Σ��ֵ & "") = 1 Then
                optInfo(0).value = True
                optInfo(1).value = False
            Else
                optInfo(0).value = False
                optInfo(1).value = True
            End If
        End If
    Case 3
        If txtInfo(txt�������).Text = "" Then
            '���û����д�������������ȱʡֵ��ȷ��ʱ��Ĭ��Ϊ����ʱ�䣬ȷ�Ͽ���Ϊ���˿��ң�ȷ����Ϊ��ǰ����Ա��Ĭ��Ϊ��Σ��ֵ
            txtInfo(txtȷ��ʱ��).Text = Format(mrsΣ��ֵ!����ʱ��, "yyyy-MM-dd HH:mm")
            txtInfo(txtȷ����).Text = UserInfo.����
            mlngȷ�Ͽ���ID = Val(txtInfo(txt����).Tag)
            txtInfo(txtȷ�Ͽ���).Text = txtInfo(txt����).Text
            optInfo(0).value = True
            optInfo(1).value = False
        Else
            txtInfo(txtȷ��ʱ��).Text = Format(mrsΣ��ֵ!ȷ��ʱ��, "yyyy-MM-dd HH:mm")
            txtInfo(txtȷ����).Text = mrsΣ��ֵ!ȷ���� & ""
            mlngȷ�Ͽ���ID = Val(mrsΣ��ֵ!ȷ�Ͽ���ID & "")
            txtInfo(txtȷ�Ͽ���).Text = Sys.RowValue("���ű�", mlngȷ�Ͽ���ID, "����")
            If Val(mrsΣ��ֵ!�Ƿ�Σ��ֵ & "") = 1 Then
                optInfo(0).value = True
                optInfo(1).value = False
            Else
                optInfo(0).value = False
                optInfo(1).value = True
            End If
        End If
    End Select
    
End Sub

Private Sub SendΣ��ֵ��Ϣ(ByVal intType As Integer)
'���ܣ�����PACSΣֵ��Ϣ,ZLHIS_PACS_005,LISΣ��ֵ��Ϣ ZLHIS_LIS_003
'������intType��0-ZLHIS_PACS_005��1-ZLHIS_LIS_003
    Dim strXML As String
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim str���� As String
    Dim str����� As String
    Dim strסԺ�� As String
    Dim int������Դ As Integer
    Dim lng����ID As Long
    Dim lng����ID As Long '������ҳ.��ǰ����ID
    Dim lng����id As Long '��������ID
    Dim lng��Ŀid As Long
    Dim strҽ������ As String
    Dim lngִ�п���ID As Long
    Dim str����IDs As String
    
    On Error GoTo errH
    
    str���� = txtInfo(txt����).Text
    
    If mint�������� = 1 Then
        str����� = txtInfo(txtסԺ��).Text
        strSql = "select b.id as ����ID,a.������Դ,a.��������id,a.������Ŀid,a.ҽ������,a.ִ�п���id from ����ҽ����¼ a,���˹Һż�¼ b" & _
            "  where a.�Һŵ�=b.no and a.�Һŵ�=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, mstr�Һŵ�)
        If Not rsTmp.EOF Then
            int������Դ = Val(rsTmp!������Դ & "")
            lng����id = Val(rsTmp!��������id & "")
            lng��Ŀid = Val(rsTmp!������ĿID & "")
            strҽ������ = rsTmp!ҽ������ & ""
            lngִ�п���ID = Val(rsTmp!ִ�п���ID & "")
            lng����ID = Val(rsTmp!����ID & "")
        End If
    Else
        strסԺ�� = txtInfo(txtסԺ��).Text
        lng����ID = mlng��ҳID
        strSql = "select a.������Դ,b.��ǰ����id,a.��������id,a.������Ŀid,a.ҽ������,a.ִ�п���id from ����ҽ����¼ a,������ҳ b" & _
            "  where a.����id=b.����id and a.��ҳid=b.��ҳid and a.id=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, mlngҽ��ID)
        If Not rsTmp.EOF Then
            int������Դ = Val(rsTmp!������Դ & "")
            lng����ID = Val(rsTmp!��ǰ����ID & "")
            lng����id = Val(rsTmp!��������id & "")
            lng��Ŀid = Val(rsTmp!������ĿID & "")
            strҽ������ = rsTmp!ҽ������ & ""
            lngִ�п���ID = Val(rsTmp!ִ�п���ID & "")
        End If
    End If
    
    If intType = 0 Then
        strXML = "<patient_info>" & vbNewLine & _
            "              <patient_id>" & mlng����ID & "</patient_id>" & vbNewLine & _
            "              <patient_name>" & str���� & "</patient_name>" & vbNewLine & _
            "              <in_number>" & strסԺ�� & "</in_number>" & vbNewLine & _
            "              <out_number>" & str����� & "</out_number>" & vbNewLine & _
            "          </patient_info>" & vbNewLine & _
            "          <patient_clinic>" & vbNewLine & _
            "              <patient_source>" & int������Դ & "</patient_source>" & vbNewLine & _
            "              <clinic_id>" & lng����ID & "</clinic_id>" & vbNewLine & _
            "              <clinic_area_id>" & lng����ID & "</clinic_area_id>" & vbNewLine & _
            "              <clinic_dept_id>" & lng����id & "</clinic_dept_id>" & vbNewLine & _
            "          </patient_clinic>" & vbNewLine & _
            "          <check_order>" & vbNewLine & _
            "              <order_id>" & mlngҽ��ID & "</order_id>" & vbNewLine & _
            "              <check_item_id>" & lng��Ŀid & "</check_item_id>" & vbNewLine & _
            "              <check_item_title>" & strҽ������ & "</check_item_title>" & vbNewLine & _
            "              <study_execute_id>" & lngִ�п���ID & "</study_execute_id>" & vbNewLine & _
            "          </check_order>"
        If Not (mclsMipModule Is Nothing) Then
            If mclsMipModule.IsConnect Then
                Call mclsMipModule.CommitMessage("ZLHIS_PACS_005", strXML)
            End If
        End If
        Call zldatabase.SendMsg("ZLHIS_PACS_005", strXML)
    Else
        '�����¿���Ϣ��LIS��Ϣ�Ȳ�����Ϣƽ̨����
        str����IDs = lng����id
        If lng����ID <> 0 Then
            If lng����ID <> lng����id Then
                str����IDs = str����IDs & "," & lng����ID
            End If
        End If
        strSql = "Zl_ҵ����Ϣ�嵥_Insert(" & mlng����ID & "," & lng����ID & "," & lng����id & ","
        strSql = strSql & IIF(lng����ID = 0, "NULL", lng����ID) & "," & int������Դ & ","
        strSql = strSql & "'" & mstrΣ��ֵ���� & "','1110','ZLHIS_LIS_003','" & mlngҽ��ID & "',3,0,sysdate,'" & str����IDs & "',null)"
        Call zldatabase.ExecuteProcedure(strSql, Me.Caption)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim strMsg As String
    
    If mblnChange Then
        
        strMsg = "��ǰ���ݱ༭����δ���棬ȷʵҪ�˳���"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True
            Exit Sub
        End If
        
    End If
End Sub

Private Sub PrintApply(ByVal intType As Integer)
'���ܴ�ӡԤ�����뵥
'������intType:1-Ԥ����2-��ӡ
    '�ж������δ�������ȱ����ٴ�ӡ
    If mblnChange Then
        If CheckData = False Then Exit Sub
        If SaveData() Then
            mblnOK = True
        End If
    End If
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1254_20", Me, "��¼ID=" & mlngΣ��ֵID, intType)
End Sub

Private Sub GetItem������(ByVal intType As Integer)
'���ܣ���ȡ����ҽ����Ŀ
'������0 �ı��򰴻س���1 �㰴ť
    Dim strSql As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean, strTmp As String
    Dim blnDo As Boolean, str���� As String
    Dim lng����ID As Long, lng��Աid As Long
    Dim i As Integer
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(txt������).Tag = txtInfo(txt������).Text Then
'            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(txt������).Text = "" Then '�൱�����������Ŀ
            txtInfo(txt������).Tag = ""
'            Call SeekNextCtl
            Exit Sub
        End If
    End If
            
    strInput = Trim(UCase(txtInfo(txt������).Text))   '�����ֵ����ǰ׺�ո�
    
    strSql = "Select A.ID,A.���,A.����,A.����,A.�����ȼ�" & _
        " From ��Ա�� A,��Ա����˵�� B" & _
        " Where A.ID=B.��ԱID And B.��Ա����='ҽ��'" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        IIF(intType = 0, " And (A.��� Like [1] Or A.���� Like [2] Or A.���� Like [2])", "") & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.���"
    vRect = zlControl.GetControlRect(txtInfo(txt������).Hwnd)
    Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, "������", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(txt������).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
        
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            If MsgBox("û���ҵ�ƥ���ҽ������ȷ��Ҫ����û�н�����Ա������ҽ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                blnDo = True
                strTmp = strInput
            Else
                blnDo = False
            End If
        Else
            Call MsgBox("û���ҵ�ƥ���ҽ��!", vbInformation, gstrSysName)
            blnDo = False
        End If
    Else
        blnDo = True
        txtInfo(txt������).Text = rsTmp!���� & ""
        txtInfo(txt������).Tag = rsTmp!���� & ""
        lng��Աid = rsTmp!ID
        txtInfo(txt������).SetFocus
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
