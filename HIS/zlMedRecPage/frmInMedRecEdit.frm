VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{D01C2596-4FE0-4EA9-9EE8-D97BE62A1165}#4.3#0"; "ZlPatiAddress.ocx"
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Begin VB.Form frmInMedRecEdit 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "סԺ��ҳ"
   ClientHeight    =   49995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   49995
   ScaleWidth      =   16005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTop 
      Appearance      =   0  'Flat
      Height          =   500
      Left            =   0
      Picture         =   "frmInMedRecEdit.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   372
      ToolTipText     =   "�ض���"
      Top             =   1000
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.VScrollBar vsbMain 
      Height          =   7335
      LargeChange     =   100
      Left            =   0
      Max             =   1000
      SmallChange     =   10
      TabIndex        =   373
      TabStop         =   0   'False
      Top             =   1800
      Width           =   255
   End
   Begin VB.HScrollBar hsbMain 
      Height          =   255
      LargeChange     =   25
      Left            =   1000
      Max             =   100
      TabIndex        =   397
      Top             =   0
      Width           =   7935
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   48700
      Left            =   600
      ScaleHeight     =   48675
      ScaleWidth      =   12465
      TabIndex        =   374
      TabStop         =   0   'False
      Top             =   300
      Width           =   12500
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6345
         Index           =   1
         Left            =   732
         ScaleHeight     =   6345
         ScaleWidth      =   10995
         TabIndex        =   376
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   2180
         Width           =   11000
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   64
            Left            =   1515
            MaxLength       =   20
            TabIndex        =   91
            Top             =   4238
            Width           =   2240
         End
         Begin VB.PictureBox picRelation 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4680
            ScaleHeight     =   240
            ScaleMode       =   0  'User
            ScaleWidth      =   1445
            TabIndex        =   403
            TabStop         =   0   'False
            Top             =   3480
            Visible         =   0   'False
            Width           =   1455
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               Index           =   41
               Left            =   0
               MaxLength       =   100
               TabIndex        =   80
               Top             =   0
               Width           =   1445
            End
            Begin VB.Line lineRelation 
               X1              =   0
               X2              =   1440.034
               Y1              =   225
               Y2              =   225
            End
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "��"
            Height          =   240
            Index           =   42
            Left            =   7920
            TabIndex        =   399
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   4845
            Width           =   270
         End
         Begin VB.CommandButton cmdDateInfo 
            Appearance      =   0  'Flat
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
            Height          =   240
            Index           =   3
            Left            =   2940
            Picture         =   "frmInMedRecEdit.frx":0359
            Style           =   1  'Graphical
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   6045
            Visible         =   0   'False
            Width           =   270
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   225
            Index           =   3
            Left            =   1335
            TabIndex        =   115
            TabStop         =   0   'False
            Tag             =   "####-##-## ##:##"
            Top             =   6060
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   225
            Index           =   2
            Left            =   1335
            TabIndex        =   98
            TabStop         =   0   'False
            Tag             =   "####-##-## ##:##"
            Top             =   5265
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "��"
            Height          =   240
            Index           =   0
            Left            =   5850
            TabIndex        =   40
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   1470
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "��"
            Height          =   240
            Index           =   2
            Left            =   5850
            TabIndex        =   58
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   2270
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "��"
            Height          =   240
            Index           =   3
            Left            =   5850
            TabIndex        =   66
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   2670
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "��"
            Height          =   240
            Index           =   4
            Left            =   5850
            TabIndex        =   86
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   3870
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "��"
            Height          =   240
            Index           =   1
            Left            =   8400
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   1440
            Width           =   270
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   225
            Index           =   0
            Left            =   5055
            TabIndex        =   15
            TabStop         =   0   'False
            Tag             =   "####-##-## ##:##"
            Top             =   180
            Visible         =   0   'False
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   16
            Mask            =   "####-##-## ##:##"
            PromptChar      =   "_"
         End
         Begin ZlPatiAddress.PatiAddress padrInfo 
            Height          =   225
            Index           =   4
            Left            =   1095
            TabIndex        =   84
            Top             =   3885
            Width           =   5025
            _ExtentX        =   8864
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   1
            MaxLength       =   100
         End
         Begin ZlPatiAddress.PatiAddress padrInfo 
            Height          =   225
            Index           =   3
            Left            =   1095
            TabIndex        =   64
            Top             =   2685
            Width           =   5025
            _ExtentX        =   8864
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   1
            MaxLength       =   100
         End
         Begin ZlPatiAddress.PatiAddress padrInfo 
            Height          =   225
            Index           =   2
            Left            =   1095
            TabIndex        =   56
            Top             =   2280
            Width           =   5025
            _ExtentX        =   8864
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   1
            MaxLength       =   100
         End
         Begin ZlPatiAddress.PatiAddress padrInfo 
            Height          =   225
            Index           =   1
            Left            =   6960
            TabIndex        =   42
            Top             =   1485
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   397
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
            Style           =   1
            MaxLength       =   100
         End
         Begin ZlPatiAddress.PatiAddress padrInfo 
            Height          =   225
            Index           =   0
            Left            =   1095
            TabIndex        =   38
            Top             =   1485
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   397
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
            Style           =   1
            MaxLength       =   100
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "��"
            Height          =   240
            Index           =   5
            Left            =   8460
            TabIndex        =   89
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   3870
            Width           =   270
         End
         Begin VB.CommandButton cmdDateInfo 
            Appearance      =   0  'Flat
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
            Height          =   240
            Index           =   2
            Left            =   2940
            Picture         =   "frmInMedRecEdit.frx":044F
            Style           =   1  'Graphical
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   5250
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "��"
            Height          =   240
            Index           =   28
            Left            =   5700
            TabIndex        =   110
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   5655
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "��"
            Height          =   240
            Index           =   27
            Left            =   2940
            TabIndex        =   107
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   5655
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "��"
            Height          =   240
            Index           =   6
            Left            =   5850
            TabIndex        =   71
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   3070
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "��"
            Height          =   240
            Index           =   29
            Left            =   8145
            TabIndex        =   113
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   5655
            Width           =   270
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   29
            Left            =   6705
            MaxLength       =   100
            TabIndex        =   112
            Top             =   5655
            Width           =   1695
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   9570
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   140
            Width           =   1125
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   0
            Left            =   5055
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   180
            Width           =   1500
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   3
            Left            =   1335
            Locked          =   -1  'True
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   6060
            Width           =   1815
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   2
            Left            =   1335
            Locked          =   -1  'True
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   5265
            Width           =   1815
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   6
            Left            =   1095
            MaxLength       =   100
            TabIndex        =   70
            ToolTipText     =   "��*����ʾ��Լ��λ�б�"
            Top             =   3078
            Width           =   5025
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   3
            Left            =   1095
            MaxLength       =   100
            TabIndex        =   65
            Top             =   2678
            Width           =   5025
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   2
            Left            =   1095
            MaxLength       =   100
            TabIndex        =   57
            ToolTipText     =   "��*����ʾ�����б�"
            Top             =   2278
            Width           =   5025
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   1
            Left            =   6960
            MaxLength       =   100
            TabIndex        =   43
            ToolTipText     =   "��*����ʾ�����б�"
            Top             =   1478
            Width           =   1740
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   0
            Left            =   1095
            MaxLength       =   100
            TabIndex        =   39
            Top             =   1478
            Width           =   5025
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   27
            Left            =   1335
            MaxLength       =   100
            TabIndex        =   106
            Top             =   5655
            Width           =   1815
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   28
            Left            =   4265
            MaxLength       =   100
            TabIndex        =   109
            Top             =   5655
            Width           =   1695
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   8
            Left            =   6705
            MaxLength       =   100
            TabIndex        =   104
            Top             =   5265
            Width           =   1695
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   7
            Left            =   4265
            Locked          =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   5265
            Width           =   1695
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   7
            ItemData        =   "frmInMedRecEdit.frx":0545
            Left            =   1340
            List            =   "frmInMedRecEdit.frx":0547
            Style           =   2  'Dropdown List
            TabIndex        =   93
            Top             =   4815
            Width           =   1900
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   6
            Left            =   7455
            MaxLength       =   5
            TabIndex        =   24
            Top             =   578
            Width           =   600
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   8
            Left            =   9570
            MaxLength       =   5
            TabIndex        =   27
            Top             =   578
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   4
            Left            =   4365
            MaxLength       =   20
            TabIndex        =   50
            Top             =   1878
            Width           =   1755
         End
         Begin VB.ComboBox cboSpecificInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   15
            Left            =   7860
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   150
            Width           =   765
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   19
            Left            =   9660
            Locked          =   -1  'True
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   6060
            Width           =   945
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   10
            Left            =   6705
            MaxLength       =   100
            TabIndex        =   121
            Top             =   6060
            Width           =   1695
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   9
            Left            =   4265
            Locked          =   -1  'True
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   6060
            Width           =   1695
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   14
            Left            =   6960
            MaxLength       =   20
            TabIndex        =   82
            Top             =   3478
            Width           =   1740
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   6
            Left            =   1095
            MaxLength       =   64
            TabIndex        =   77
            Top             =   3478
            Width           =   1065
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   2
            Left            =   9570
            MaxLength       =   6
            TabIndex        =   75
            Top             =   3078
            Width           =   1125
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   6960
            MaxLength       =   20
            TabIndex        =   73
            Top             =   3078
            Width           =   1740
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   4
            Left            =   9570
            MaxLength       =   6
            TabIndex        =   62
            Top             =   2278
            Width           =   1125
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   6960
            MaxLength       =   20
            TabIndex        =   60
            Top             =   2278
            Width           =   1740
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   5
            ItemData        =   "frmInMedRecEdit.frx":0549
            Left            =   9570
            List            =   "frmInMedRecEdit.frx":054B
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   1440
            Width           =   1125
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   3
            ItemData        =   "frmInMedRecEdit.frx":054D
            Left            =   6960
            List            =   "frmInMedRecEdit.frx":054F
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   1840
            Width           =   1740
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            ItemData        =   "frmInMedRecEdit.frx":0551
            Left            =   9570
            List            =   "frmInMedRecEdit.frx":0553
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   1840
            Width           =   1125
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   15
            Left            =   7455
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   178
            Width           =   360
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            ItemData        =   "frmInMedRecEdit.frx":0555
            Left            =   2970
            List            =   "frmInMedRecEdit.frx":0557
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   140
            Width           =   885
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   3
            Left            =   980
            Locked          =   -1  'True
            MaxLength       =   64
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   178
            Width           =   1140
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   5
            Left            =   6960
            MaxLength       =   30
            TabIndex        =   88
            ToolTipText     =   "��*����ʾ�����б�"
            Top             =   3878
            Width           =   1740
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   16
            Left            =   3090
            MaxLength       =   20
            TabIndex        =   30
            Top             =   973
            Width           =   360
         End
         Begin VB.ComboBox cboSpecificInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   16
            Left            =   3810
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   945
            Width           =   765
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   18
            Left            =   9090
            MaxLength       =   25
            TabIndex        =   36
            Top             =   975
            Width           =   1155
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��Ժǰ����Ժ����"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   1
            Left            =   8775
            TabIndex        =   96
            Top             =   4845
            Width           =   1890
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   5
            Left            =   9570
            MaxLength       =   6
            TabIndex        =   68
            Top             =   2678
            Width           =   1125
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   17
            Left            =   6090
            MaxLength       =   25
            TabIndex        =   34
            Top             =   978
            Width           =   1155
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   61
            ItemData        =   "frmInMedRecEdit.frx":0559
            Left            =   1095
            List            =   "frmInMedRecEdit.frx":055B
            TabIndex        =   48
            Top             =   1840
            Width           =   2340
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   42
            Left            =   4265
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   4860
            Width           =   3700
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   4
            Left            =   1095
            MaxLength       =   100
            TabIndex        =   85
            Top             =   3878
            Width           =   5025
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   6
            ItemData        =   "frmInMedRecEdit.frx":055D
            Left            =   3060
            List            =   "frmInMedRecEdit.frx":055F
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   3440
            Width           =   1605
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����Ժ"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   5520
            TabIndex        =   22
            Top             =   565
            Width           =   960
         End
         Begin VB.TextBox txtSpecificInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   52
            Left            =   3570
            MaxLength       =   2
            TabIndex        =   31
            Top             =   840
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�໤�����֤��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   129
            Left            =   120
            TabIndex        =   90
            Top             =   4260
            Width           =   1260
         End
         Begin VB.Label lblSpecificInfo 
            AutoSize        =   -1  'True
            Caption         =   "30"
            Height          =   180
            Index           =   52
            Left            =   3570
            TabIndex        =   402
            Top             =   1125
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Line lineH 
            Index           =   0
            X1              =   0
            X2              =   14200
            Y1              =   1350
            Y2              =   1350
         End
         Begin VB.Line lineH 
            Index           =   1
            X1              =   0
            X2              =   14200
            Y1              =   4725
            Y2              =   4725
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   9135
            TabIndex        =   20
            Top             =   195
            Width           =   360
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������������              ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   17
            Left            =   4800
            TabIndex        =   33
            Top             =   1005
            Width           =   2700
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ;��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   540
            TabIndex        =   92
            Top             =   4875
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   6
            Left            =   7020
            TabIndex        =   23
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CM"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   8100
            TabIndex        =   25
            Top             =   600
            Width           =   180
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   8
            Left            =   9135
            TabIndex        =   26
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "KG"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   9
            Left            =   10290
            TabIndex        =   28
            Top             =   600
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����֤��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   3570
            TabIndex        =   49
            Top             =   1905
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   19
            Left            =   8865
            TabIndex        =   122
            Top             =   6075
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   10
            Left            =   6270
            TabIndex        =   120
            Top             =   6075
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   9
            Left            =   3825
            TabIndex        =   118
            Top             =   6075
            Width           =   360
         End
         Begin VB.Label lblDateInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժʱ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   540
            TabIndex        =   114
            Top             =   6075
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   8
            Left            =   6270
            TabIndex        =   103
            Top             =   5280
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   3840
            TabIndex        =   101
            Top             =   5280
            Width           =   360
         End
         Begin VB.Label lblDateInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժʱ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   540
            TabIndex        =   97
            Top             =   5280
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�绰"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   14
            Left            =   6525
            TabIndex        =   81
            Top             =   3495
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   6
            Left            =   2625
            TabIndex        =   78
            Top             =   3495
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ������"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   76
            Top             =   3495
            Width           =   900
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�ʱ�"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   9135
            TabIndex        =   74
            Top             =   3105
            Width           =   360
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�绰"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   6525
            TabIndex        =   72
            Top             =   3105
            Width           =   360
         End
         Begin VB.Label lblAdressInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������λ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   6
            Left            =   300
            TabIndex        =   69
            Top             =   3105
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�ʱ�"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   9135
            TabIndex        =   61
            Top             =   2295
            Width           =   360
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�绰"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   6525
            TabIndex        =   59
            Top             =   2295
            Width           =   360
         End
         Begin VB.Label lblAdressInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��סַ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   480
            TabIndex        =   55
            Top             =   2295
            Width           =   540
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���֤��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   61
            Left            =   300
            TabIndex        =   47
            Top             =   1905
            Width           =   720
         End
         Begin VB.Label lblAdressInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   480
            TabIndex        =   37
            Top             =   1500
            Width           =   540
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   5
            Left            =   9135
            TabIndex        =   45
            Top             =   1500
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ְҵ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   6525
            TabIndex        =   51
            Top             =   1905
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   9135
            TabIndex        =   53
            Top             =   1905
            Width           =   360
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   15
            Left            =   7020
            TabIndex        =   17
            Top             =   195
            Width           =   360
         End
         Begin VB.Label lblDateInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   4260
            TabIndex        =   14
            Top             =   195
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�Ա�"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   2535
            TabIndex        =   12
            Top             =   195
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   540
            TabIndex        =   10
            Top             =   200
            Width           =   360
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(���䲻��һ�����)����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   16
            Left            =   915
            TabIndex        =   29
            Top             =   1005
            Width           =   2100
         End
         Begin VB.Label lblAdressInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   6525
            TabIndex        =   41
            Top             =   1500
            Width           =   360
         End
         Begin VB.Label lblAdressInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���ڵ�ַ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   300
            TabIndex        =   63
            Top             =   2700
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�ʱ�"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   5
            Left            =   9120
            TabIndex        =   67
            Top             =   2700
            Width           =   360
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������Ժ����               ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   18
            Left            =   7740
            TabIndex        =   35
            Top             =   1005
            Width           =   2790
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ת��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   27
            Left            =   540
            TabIndex        =   105
            Top             =   5685
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   28
            Left            =   3840
            TabIndex        =   108
            Top             =   5685
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   29
            Left            =   6270
            TabIndex        =   111
            Top             =   5685
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ת��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   42
            Left            =   3825
            TabIndex        =   94
            Top             =   4875
            Width           =   360
         End
         Begin VB.Label lblAdressInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ�˵�ַ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   83
            Top             =   3900
            Width           =   900
         End
         Begin VB.Label lblAdressInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   5
            Left            =   6525
            TabIndex        =   87
            Top             =   3900
            Width           =   360
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2180
         Index           =   0
         Left            =   732
         ScaleHeight     =   2175
         ScaleWidth      =   10995
         TabIndex        =   375
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   0
         Width           =   11000
         Begin VB.PictureBox PicInNum 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   9375
            ScaleHeight     =   240
            ScaleMode       =   0  'User
            ScaleWidth      =   1593.969
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1770
            Width           =   1600
            Begin VB.TextBox txtSpecificInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   225
               Index           =   20
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   0
               Width           =   1611
            End
            Begin VB.Line lineInNum 
               X1              =   0
               X2              =   2145.155
               Y1              =   225
               Y2              =   225
            End
         End
         Begin VB.Frame fraCbo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   4440
            TabIndex        =   396
            Top             =   1080
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   63
            Left            =   800
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1778
            Width           =   1620
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   0
            Left            =   1160
            TabIndex        =   2
            Text            =   "Combo1"
            Top             =   1340
            Width           =   2300
         End
         Begin VB.TextBox txtSpecificInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1755
            Width           =   375
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   20
            Left            =   8760
            TabIndex        =   7
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��     ��סԺ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   11
            Left            =   4680
            TabIndex        =   5
            Top             =   1800
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   63
            Left            =   0
            TabIndex        =   3
            Top             =   1800
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ�Ƹ��ѷ�ʽ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   1
            Top             =   1400
            Width           =   1080
         End
         Begin VB.Label lblHead 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ס Ժ �� ҳ"
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
            Left            =   4300
            TabIndex        =   0
            Tag             =   "241,88"
            Top             =   360
            Width           =   2085
         End
      End
      Begin VB.PictureBox picInfectInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   960
         ScaleHeight     =   3465
         ScaleWidth      =   4065
         TabIndex        =   128
         Top             =   0
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ListBox lstInfectParts 
            Appearance      =   0  'Flat
            Height          =   2340
            ItemData        =   "frmInMedRecEdit.frx":0561
            Left            =   240
            List            =   "frmInMedRecEdit.frx":0568
            Style           =   1  'Checkbox
            TabIndex        =   132
            Top             =   840
            Width           =   3615
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            Index           =   10
            ItemData        =   "frmInMedRecEdit.frx":057C
            Left            =   1680
            List            =   "frmInMedRecEdit.frx":057E
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ⱦ��λ"
            Height          =   180
            Index           =   128
            Left            =   120
            TabIndex        =   131
            Top             =   480
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��Ⱦ�������Ĺ�ϵ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   10
            Left            =   120
            TabIndex        =   129
            Top             =   180
            Width           =   1440
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3815
         Index           =   18
         Left            =   732
         ScaleHeight     =   3810
         ScaleWidth      =   10995
         TabIndex        =   393
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   44220
         Width           =   11000
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "1.CT"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   13
            Left            =   645
            TabIndex        =   357
            Top             =   2355
            Width           =   795
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "2.MRI"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   14
            Left            =   1740
            TabIndex        =   358
            Top             =   2355
            Width           =   765
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "3.��ɫ������"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   15
            Left            =   2940
            TabIndex        =   359
            Top             =   2355
            Width           =   1545
         End
         Begin VB.ListBox lstInfection 
            Appearance      =   0  'Flat
            Height          =   1290
            ItemData        =   "frmInMedRecEdit.frx":0580
            Left            =   0
            List            =   "frmInMedRecEdit.frx":0587
            Style           =   1  'Checkbox
            TabIndex        =   355
            Top             =   600
            Width           =   5200
         End
         Begin VB.PictureBox PicAdvEvent 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   3015
            Left            =   5400
            ScaleHeight     =   2985
            ScaleWidth      =   5475
            TabIndex        =   362
            TabStop         =   0   'False
            Top             =   600
            Width           =   5500
            Begin VB.ListBox lstAdvEvent 
               Appearance      =   0  'Flat
               Height          =   1500
               ItemData        =   "frmInMedRecEdit.frx":0599
               Left            =   -15
               List            =   "frmInMedRecEdit.frx":05A0
               Style           =   1  'Checkbox
               TabIndex        =   363
               Top             =   -15
               Width           =   5500
            End
            Begin VB.ComboBox cboBaseInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   46
               ItemData        =   "frmInMedRecEdit.frx":05B1
               Left            =   3795
               List            =   "frmInMedRecEdit.frx":05B3
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   367
               TabStop         =   0   'False
               Top             =   1640
               Width           =   1620
            End
            Begin VB.ComboBox cboBaseInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   48
               ItemData        =   "frmInMedRecEdit.frx":05B5
               Left            =   1460
               List            =   "frmInMedRecEdit.frx":05B7
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   371
               TabStop         =   0   'False
               Top             =   2540
               Width           =   3960
            End
            Begin VB.ComboBox cboBaseInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   47
               ItemData        =   "frmInMedRecEdit.frx":05B9
               Left            =   1460
               List            =   "frmInMedRecEdit.frx":05BB
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   369
               TabStop         =   0   'False
               Top             =   2090
               Width           =   3960
            End
            Begin VB.ComboBox cboBaseInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   45
               ItemData        =   "frmInMedRecEdit.frx":05BD
               Left            =   1460
               List            =   "frmInMedRecEdit.frx":05BF
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   365
               TabStop         =   0   'False
               Top             =   1640
               Width           =   1680
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "ѹ�������ڼ�"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   45
               Left            =   300
               TabIndex        =   364
               Top             =   1700
               Width           =   1080
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������׹��ԭ��"
               Height          =   180
               Index           =   48
               Left            =   120
               TabIndex        =   370
               Top             =   2600
               Width           =   1260
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������׹���˺�"
               Height          =   180
               Index           =   47
               Left            =   120
               TabIndex        =   368
               Top             =   2150
               Width           =   1260
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Index           =   46
               Left            =   3360
               TabIndex        =   366
               Top             =   1695
               Width           =   360
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsTSJC 
            Height          =   930
            Left            =   0
            TabIndex        =   360
            Top             =   2685
            Width           =   5200
            _cx             =   9172
            _cy             =   1640
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit.frx":05C1
            ScrollTrack     =   -1  'True
            ScrollBars      =   0
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
            Editable        =   2
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
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҳ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   18
            Left            =   0
            TabIndex        =   353
            Top             =   0
            Width           =   450
         End
         Begin VB.Label lblTSJC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���������"
            Height          =   180
            Left            =   0
            TabIndex        =   356
            Top             =   2040
            Width           =   1080
         End
         Begin VB.Line lineCheck 
            X1              =   1080
            X2              =   5160
            Y1              =   2130
            Y2              =   2130
         End
         Begin VB.Label lblInfection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��Ⱦ����"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   354
            Top             =   360
            Width           =   720
         End
         Begin VB.Line lineInfection 
            X1              =   720
            X2              =   5160
            Y1              =   450
            Y2              =   450
         End
         Begin VB.Line lineAdvEvent 
            X1              =   6120
            X2              =   11040
            Y1              =   450
            Y2              =   450
         End
         Begin VB.Label lblAdvEvent 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�����¼�"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   5400
            TabIndex        =   361
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3560
         Index           =   2
         Left            =   732
         ScaleHeight     =   3555
         ScaleWidth      =   10995
         TabIndex        =   377
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   8760
         Width           =   11000
         Begin VB.OptionButton optDiag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������ϱ�׼����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   6000
            TabIndex        =   125
            TabStop         =   0   'False
            Top             =   55
            Value           =   -1  'True
            Width           =   1770
         End
         Begin VB.OptionButton optDiag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���ݼ�����������"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   8040
            TabIndex        =   126
            TabStop         =   0   'False
            Top             =   55
            Width           =   1770
         End
         Begin VB.CommandButton cmdDiagMove 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Height          =   375
            Index           =   1
            Left            =   10620
            Picture         =   "frmInMedRecEdit.frx":0632
            Style           =   1  'Graphical
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton cmdDiagMove 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Height          =   375
            Index           =   0
            Left            =   10620
            Picture         =   "frmInMedRecEdit.frx":2D5D
            Style           =   1  'Graphical
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   1320
            Width           =   375
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
            Height          =   3000
            Left            =   0
            TabIndex        =   127
            Top             =   360
            Width           =   10500
            _cx             =   18521
            _cy             =   5292
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   9
            Cols            =   24
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmInMedRecEdit.frx":5305
            ScrollTrack     =   -1  'True
            ScrollBars      =   1
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
            Editable        =   2
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
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ҽ���"
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
            Index           =   2
            Left            =   0
            TabIndex        =   124
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   732
         ScaleHeight     =   255
         ScaleWidth      =   10995
         TabIndex        =   385
         TabStop         =   0   'False
         Tag             =   "false"
         Top             =   30000
         Visible         =   0   'False
         Width           =   11000
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ����"
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
            Left            =   0
            TabIndex        =   294
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2700
         Index           =   4
         Left            =   732
         ScaleHeight     =   2700
         ScaleWidth      =   10995
         TabIndex        =   379
         TabStop         =   0   'False
         Top             =   16290
         Visible         =   0   'False
         Width           =   11000
         Begin VB.CommandButton cmdDiagMove 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Height          =   375
            Index           =   3
            Left            =   10620
            Picture         =   "frmInMedRecEdit.frx":565F
            Style           =   1  'Graphical
            TabIndex        =   186
            TabStop         =   0   'False
            Top             =   1335
            Width           =   375
         End
         Begin VB.CommandButton cmdDiagMove 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Height          =   375
            Index           =   2
            Left            =   10620
            Picture         =   "frmInMedRecEdit.frx":7D8A
            Style           =   1  'Graphical
            TabIndex        =   185
            TabStop         =   0   'False
            Top             =   855
            Width           =   375
         End
         Begin VB.OptionButton optDiag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���ݼ�����������"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   8040
            TabIndex        =   183
            TabStop         =   0   'False
            Top             =   150
            Width           =   1770
         End
         Begin VB.OptionButton optDiag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������ϱ�׼����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   6000
            TabIndex        =   182
            TabStop         =   0   'False
            Top             =   150
            Value           =   -1  'True
            Width           =   1890
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
            Height          =   2100
            Left            =   0
            TabIndex        =   184
            Top             =   405
            Width           =   10500
            _cx             =   18521
            _cy             =   3704
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   24
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmInMedRecEdit.frx":A332
            ScrollTrack     =   -1  'True
            ScrollBars      =   1
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
            Editable        =   2
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
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ҽ���"
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
            Index           =   4
            Left            =   0
            TabIndex        =   181
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3975
         Index           =   3
         Left            =   732
         ScaleHeight     =   3975
         ScaleWidth      =   10995
         TabIndex        =   378
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   12315
         Width           =   11000
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   20
            ItemData        =   "frmInMedRecEdit.frx":A626
            Left            =   4575
            List            =   "frmInMedRecEdit.frx":A628
            Style           =   2  'Dropdown List
            TabIndex        =   400
            Top             =   3180
            Width           =   1470
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   225
            Index           =   5
            Left            =   1560
            TabIndex        =   172
            TabStop         =   0   'False
            Tag             =   "####-##-## ##:##:##"
            Top             =   3225
            Visible         =   0   'False
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   21
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   165
            Top             =   2718
            Width           =   600
         End
         Begin VB.CommandButton cmdDateInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   240
            Index           =   4
            Left            =   10260
            Picture         =   "frmInMedRecEdit.frx":A62A
            Style           =   1  'Graphical
            TabIndex        =   143
            TabStop         =   0   'False
            Top             =   310
            Width           =   285
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   225
            Index           =   4
            Left            =   8430
            TabIndex        =   141
            TabStop         =   0   'False
            Tag             =   "####-##-## ##:##:##"
            Top             =   315
            Visible         =   0   'False
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483633
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   240
            Index           =   21
            Left            =   10290
            TabIndex        =   157
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   1710
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   240
            Index           =   22
            Left            =   10290
            TabIndex        =   170
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   2710
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   240
            Index           =   20
            Left            =   10290
            TabIndex        =   176
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   3210
            Width           =   270
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   60
            ItemData        =   "frmInMedRecEdit.frx":A720
            Left            =   1950
            List            =   "frmInMedRecEdit.frx":A722
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   178
            TabStop         =   0   'False
            Top             =   3580
            Width           =   765
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   16
            ItemData        =   "frmInMedRecEdit.frx":A724
            Left            =   1680
            List            =   "frmInMedRecEdit.frx":A726
            Style           =   2  'Dropdown List
            TabIndex        =   149
            Top             =   1180
            Width           =   1470
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�Ƿ�ȷ��"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   2
            Left            =   5235
            TabIndex        =   139
            Top             =   305
            Width           =   1170
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   137
            Top             =   280
            Width           =   1470
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   19
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   159
            TabStop         =   0   'False
            Top             =   2218
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   18
            ItemData        =   "frmInMedRecEdit.frx":A728
            Left            =   9090
            List            =   "frmInMedRecEdit.frx":A72A
            Style           =   2  'Dropdown List
            TabIndex        =   163
            Top             =   2180
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   21
            ItemData        =   "frmInMedRecEdit.frx":A72C
            Left            =   5580
            List            =   "frmInMedRecEdit.frx":A72E
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   180
            TabStop         =   0   'False
            Top             =   3580
            Width           =   1470
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   22
            Left            =   5580
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   169
            TabStop         =   0   'False
            Top             =   2718
            Width           =   4980
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   22
            Left            =   3420
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   167
            TabStop         =   0   'False
            Top             =   2718
            Width           =   600
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�·�����"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   5
            Left            =   3615
            TabIndex        =   138
            Top             =   305
            Width           =   1170
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ҽԺ��Ⱦ����ԭѧ���"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   3
            Left            =   1185
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   1705
            Width           =   2370
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   21
            Left            =   5580
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   156
            TabStop         =   0   'False
            Top             =   1718
            Width           =   4980
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   13
            ItemData        =   "frmInMedRecEdit.frx":A730
            Left            =   5580
            List            =   "frmInMedRecEdit.frx":A732
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   147
            TabStop         =   0   'False
            Top             =   780
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   12
            ItemData        =   "frmInMedRecEdit.frx":A734
            Left            =   1680
            List            =   "frmInMedRecEdit.frx":A736
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   145
            TabStop         =   0   'False
            Top             =   780
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   19
            ItemData        =   "frmInMedRecEdit.frx":A738
            Left            =   5580
            List            =   "frmInMedRecEdit.frx":A73A
            Style           =   2  'Dropdown List
            TabIndex        =   161
            Top             =   2180
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   15
            ItemData        =   "frmInMedRecEdit.frx":A73C
            Left            =   9090
            List            =   "frmInMedRecEdit.frx":A73E
            Style           =   2  'Dropdown List
            TabIndex        =   153
            Top             =   1180
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   14
            ItemData        =   "frmInMedRecEdit.frx":A740
            Left            =   5580
            List            =   "frmInMedRecEdit.frx":A742
            Style           =   2  'Dropdown List
            TabIndex        =   151
            Top             =   1180
            Width           =   1470
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            HelpContextID   =   20
            Index           =   20
            Left            =   7020
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   175
            TabStop         =   0   'False
            Top             =   3218
            Width           =   3540
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   4
            Left            =   8430
            Locked          =   -1  'True
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   315
            Width           =   1830
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   5
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   173
            TabStop         =   0   'False
            Top             =   3218
            Width           =   1830
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����ڼ�"
            Height          =   180
            Index           =   20
            Left            =   3780
            TabIndex        =   401
            Top             =   3240
            Width           =   720
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ҽ��Ϸ������"
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
            Index           =   3
            Left            =   0
            TabIndex        =   135
            Top             =   0
            Width           =   1800
         End
         Begin VB.Line lineH 
            Index           =   6
            X1              =   0
            X2              =   14400
            Y1              =   3080
            Y2              =   3080
         End
         Begin VB.Line lineH 
            Index           =   5
            X1              =   0
            X2              =   14400
            Y1              =   2580
            Y2              =   2580
         End
         Begin VB.Line lineH 
            Index           =   4
            X1              =   0
            X2              =   14400
            Y1              =   2080
            Y2              =   2080
         End
         Begin VB.Line lineH 
            Index           =   3
            X1              =   0
            X2              =   14400
            Y1              =   1580
            Y2              =   1580
         End
         Begin VB.Line lineH 
            Index           =   2
            X1              =   0
            X2              =   14200
            Y1              =   680
            Y2              =   680
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������ʬ��"
            Height          =   180
            Index           =   60
            Left            =   795
            TabIndex        =   177
            Top             =   3640
            Width           =   1080
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������Ժ"
            Height          =   180
            Index           =   16
            Left            =   705
            TabIndex        =   148
            Top             =   1245
            Width           =   900
         End
         Begin VB.Label lblDateInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ҫ���ȷ������"
            Height          =   180
            Index           =   4
            Left            =   6915
            TabIndex        =   140
            Top             =   345
            Width           =   1440
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ���"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   11
            Left            =   885
            TabIndex        =   136
            Top             =   345
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   19
            Left            =   945
            TabIndex        =   158
            Top             =   2235
            Width           =   540
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���Ȳ���"
            Height          =   180
            Index           =   22
            Left            =   4785
            TabIndex        =   168
            Top             =   2745
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ɹ�����"
            Height          =   180
            Index           =   22
            Left            =   2625
            TabIndex        =   166
            Top             =   2745
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ȴ���"
            Height          =   180
            Index           =   21
            Left            =   765
            TabIndex        =   164
            Top             =   2745
            Width           =   720
         End
         Begin VB.Label lblDateInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
            Height          =   180
            Index           =   5
            Left            =   765
            TabIndex        =   171
            Top             =   3240
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ԭ��"
            Height          =   180
            Index           =   20
            Left            =   6240
            TabIndex        =   174
            Top             =   3240
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽԺ��Ⱦ��ԭѧ���"
            Height          =   180
            Index           =   21
            Left            =   3885
            TabIndex        =   155
            Top             =   1740
            Width           =   1620
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����������"
            Height          =   180
            Index           =   13
            Left            =   4425
            TabIndex        =   146
            Top             =   840
            Width           =   1080
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ֻ��̶�"
            Height          =   180
            Index           =   12
            Left            =   885
            TabIndex        =   144
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ٴ���ʬ��"
            Height          =   180
            Index           =   21
            Left            =   4605
            TabIndex        =   179
            Top             =   3640
            Width           =   900
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ٴ��벡��"
            Height          =   180
            Index           =   19
            Left            =   4605
            TabIndex        =   160
            Top             =   2235
            Width           =   900
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����벡��"
            Height          =   180
            Index           =   18
            Left            =   8115
            TabIndex        =   162
            Top             =   2235
            Width           =   900
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ���Ժ"
            Height          =   180
            Index           =   15
            Left            =   8115
            TabIndex        =   152
            Top             =   1245
            Width           =   900
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������Ժ"
            Height          =   180
            Index           =   14
            Left            =   4605
            TabIndex        =   150
            Top             =   1245
            Width           =   900
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2220
         Index           =   9
         Left            =   732
         ScaleHeight     =   2220
         ScaleWidth      =   10995
         TabIndex        =   384
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   27780
         Width           =   11000
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   17
            ItemData        =   "frmInMedRecEdit.frx":A744
            Left            =   1200
            List            =   "frmInMedRecEdit.frx":A746
            Style           =   2  'Dropdown List
            TabIndex        =   285
            Top             =   290
            Width           =   1470
         End
         Begin VB.CommandButton cmdAutoLoad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�Զ���ȡ"
            Height          =   330
            Index           =   1
            Left            =   9480
            TabIndex        =   290
            TabStop         =   0   'False
            Top             =   240
            Width           =   1000
         End
         Begin VB.CommandButton cmdOPSMove 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Height          =   375
            Index           =   0
            Left            =   10620
            Picture         =   "frmInMedRecEdit.frx":A748
            Style           =   1  'Graphical
            TabIndex        =   292
            TabStop         =   0   'False
            Top             =   1200
            Width           =   375
         End
         Begin VB.CommandButton cmdOPSMove 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Height          =   375
            Index           =   1
            Left            =   10620
            Picture         =   "frmInMedRecEdit.frx":CCF0
            Style           =   1  'Graphical
            TabIndex        =   293
            TabStop         =   0   'False
            Top             =   1680
            Width           =   375
         End
         Begin VB.CheckBox chkParaOPSInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "δ�ҵ�ʱ��������¼��"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   6480
            TabIndex        =   288
            Top             =   150
            Width           =   2115
         End
         Begin VB.OptionButton OptParaOPSInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����ICD9-CM3����"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   1
            Left            =   4710
            MaskColor       =   &H80000005&
            TabIndex        =   287
            TabStop         =   0   'False
            Top             =   150
            Width           =   1770
         End
         Begin VB.OptionButton OptParaOPSInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����������Ŀ����"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   2880
            TabIndex        =   286
            TabStop         =   0   'False
            Top             =   150
            Value           =   -1  'True
            Width           =   1845
         End
         Begin VSFlex8Ctl.VSFlexGrid vsOPS 
            Height          =   1200
            Left            =   0
            TabIndex        =   291
            Top             =   720
            Width           =   10500
            _cx             =   18521
            _cy             =   2117
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   43
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmInMedRecEdit.frx":F41B
            ScrollTrack     =   -1  'True
            ScrollBars      =   1
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
            Editable        =   2
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
            Begin VB.PictureBox picCopy 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               Picture         =   "frmInMedRecEdit.frx":FAC9
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   395
               TabStop         =   0   'False
               Top             =   300
               Visible         =   0   'False
               Width           =   255
            End
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ǰ������"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   17
            Left            =   270
            TabIndex        =   284
            Top             =   350
            Width           =   900
         End
         Begin VB.Label lblAutoInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��ʾ��Ϣ"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   2880
            TabIndex        =   289
            Top             =   480
            Visible         =   0   'False
            Width           =   6375
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������¼"
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
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   283
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1745
         Index           =   17
         Left            =   732
         ScaleHeight     =   1740
         ScaleWidth      =   10995
         TabIndex        =   392
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   42465
         Width           =   11000
         Begin VSFlex8Ctl.VSFlexGrid vsfMain 
            Height          =   1200
            Left            =   0
            TabIndex        =   352
            Top             =   345
            Width           =   10990
            _cx             =   19385
            _cy             =   2117
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483642
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   100
            ColWidthMax     =   2400
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit.frx":107BB
            ScrollTrack     =   0   'False
            ScrollBars      =   1
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
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
            Editable        =   2
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
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����������Ŀ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   17
            Left            =   0
            TabIndex        =   351
            Top             =   0
            Width           =   1350
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1760
         Index           =   16
         Left            =   732
         ScaleHeight     =   1755
         ScaleWidth      =   10995
         TabIndex        =   391
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   40710
         Width           =   11000
         Begin VSFlex8Ctl.VSFlexGrid vsFlxAddICU 
            Height          =   1200
            Left            =   0
            TabIndex        =   350
            Top             =   360
            Width           =   10990
            _cx             =   19385
            _cy             =   2117
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit.frx":108CB
            ScrollTrack     =   0   'False
            ScrollBars      =   1
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
            Editable        =   2
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
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��֢�໤���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   16
            Left            =   0
            TabIndex        =   349
            Top             =   0
            Width           =   1350
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1850
         Index           =   13
         Left            =   732
         ScaleHeight     =   1845
         ScaleWidth      =   10995
         TabIndex        =   388
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   35205
         Width           =   11000
         Begin VSFlex8Ctl.VSFlexGrid vsRadioth 
            Height          =   1200
            Left            =   0
            TabIndex        =   343
            Top             =   480
            Width           =   10990
            _cx             =   19385
            _cy             =   2117
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit.frx":109FD
            ScrollTrack     =   0   'False
            ScrollBars      =   1
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
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   -1  'True
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
         Begin VB.Label lblEdit 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ʾ��Ϣ"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   1
            Left            =   2760
            TabIndex        =   342
            Top             =   180
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���Ƽ�¼��Ϣ"
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
            Left            =   0
            TabIndex        =   341
            Top             =   0
            Width           =   1350
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1850
         Index           =   12
         Left            =   732
         ScaleHeight     =   1845
         ScaleWidth      =   10995
         TabIndex        =   387
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   33360
         Width           =   11000
         Begin VSFlex8Ctl.VSFlexGrid vsChemoth 
            Height          =   1200
            Left            =   0
            TabIndex        =   340
            Top             =   480
            Width           =   10990
            _cx             =   19385
            _cy             =   2117
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit.frx":10B45
            ScrollTrack     =   -1  'True
            ScrollBars      =   1
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
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   -1  'True
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
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���Ƽ�¼��Ϣ"
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
            Index           =   12
            Left            =   0
            TabIndex        =   338
            Top             =   0
            Width           =   1350
         End
         Begin VB.Label lblEdit 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ʾ��Ϣ"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   0
            Left            =   2880
            TabIndex        =   339
            Top             =   180
            Visible         =   0   'False
            Width           =   720
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3100
         Index           =   11
         Left            =   732
         ScaleHeight     =   3105
         ScaleWidth      =   10995
         TabIndex        =   386
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   30255
         Width           =   11000
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   225
            Index           =   24
            Left            =   10485
            TabIndex        =   317
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   1320
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   240
            Index           =   17
            Left            =   6540
            TabIndex        =   310
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   790
            Width           =   270
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   225
            Index           =   6
            Left            =   9075
            TabIndex        =   302
            Tag             =   "####-##-##"
            Top             =   405
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            MaxLength       =   10
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   6
            Left            =   9075
            MaxLength       =   30
            TabIndex        =   304
            Top             =   405
            Visible         =   0   'False
            Width           =   1065
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   225
            Index           =   7
            Left            =   10155
            TabIndex        =   303
            Tag             =   "##:##"
            Top             =   405
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   7
            Left            =   10155
            MaxLength       =   30
            TabIndex        =   305
            Top             =   405
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���в���"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   10
            Left            =   5010
            TabIndex        =   299
            Top             =   385
            Width           =   1170
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   34
            ItemData        =   "frmInMedRecEdit.frx":10C93
            Left            =   1155
            List            =   "frmInMedRecEdit.frx":10C95
            Style           =   2  'Dropdown List
            TabIndex        =   297
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ʾ�̲���"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   9
            Left            =   3690
            TabIndex        =   298
            Top             =   385
            Width           =   1050
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   18
            Left            =   8355
            MaxLength       =   100
            TabIndex        =   312
            Top             =   798
            Width           =   2400
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   17
            Left            =   4275
            MaxLength       =   100
            TabIndex        =   309
            Top             =   798
            Width           =   2535
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            Index           =   43
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   307
            Top             =   760
            Width           =   1815
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���Ѳ���"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   11
            Left            =   6330
            TabIndex        =   300
            Top             =   385
            Width           =   1050
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   12
            Left            =   4170
            TabIndex        =   334
            Top             =   2680
            Width           =   690
         End
         Begin VB.ComboBox cboSpecificInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   41
            ItemData        =   "frmInMedRecEdit.frx":10C97
            Left            =   7380
            List            =   "frmInMedRecEdit.frx":10C99
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   337
            TabStop         =   0   'False
            Top             =   2670
            Width           =   735
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   41
            Left            =   6570
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   336
            TabStop         =   0   'False
            Top             =   2698
            Width           =   765
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   44
            ItemData        =   "frmInMedRecEdit.frx":10C9B
            Left            =   1155
            List            =   "frmInMedRecEdit.frx":10C9D
            Style           =   2  'Dropdown List
            TabIndex        =   314
            Top             =   1260
            Width           =   1815
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   24
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   316
            TabStop         =   0   'False
            Top             =   1298
            Width           =   5475
         End
         Begin VB.OptionButton optInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   3090
            TabIndex        =   320
            Top             =   1780
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton optInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�У�Ŀ�ģ�"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   1
            Left            =   4020
            TabIndex        =   321
            Top             =   1780
            Width           =   1200
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   49
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   319
            Top             =   1760
            Width           =   1815
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   35
            Left            =   3120
            MaxLength       =   6
            TabIndex        =   325
            Top             =   2298
            Width           =   675
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   38
            Left            =   7680
            MaxLength       =   6
            TabIndex        =   329
            Top             =   2298
            Width           =   675
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   34
            Left            =   1155
            MaxLength       =   6
            TabIndex        =   333
            Top             =   2698
            Width           =   1035
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   40
            Left            =   9780
            MaxLength       =   6
            TabIndex        =   331
            Top             =   2298
            Width           =   675
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   39
            Left            =   8640
            MaxLength       =   6
            TabIndex        =   330
            Top             =   2298
            Width           =   675
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   37
            Left            =   5220
            MaxLength       =   6
            TabIndex        =   327
            Top             =   2298
            Width           =   675
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   36
            Left            =   4050
            MaxLength       =   6
            TabIndex        =   326
            Top             =   2298
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   25
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   322
            Top             =   1798
            Width           =   5475
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ���"
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
            Index           =   11
            Left            =   0
            TabIndex        =   295
            Top             =   0
            Width           =   900
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ҽѧ��ʾ"
            Height          =   180
            Index           =   17
            Left            =   3480
            TabIndex        =   308
            Top             =   825
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Index           =   34
            Left            =   360
            TabIndex        =   296
            Top             =   420
            Width           =   720
         End
         Begin VB.Label lblDateInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
            Height          =   180
            Index           =   6
            Left            =   8280
            TabIndex        =   301
            Top             =   420
            Width           =   720
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ҽѧ��ʾ"
            Height          =   180
            Index           =   18
            Left            =   7200
            TabIndex        =   311
            Top             =   825
            Width           =   1080
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����״��"
            Height          =   180
            Index           =   43
            Left            =   360
            TabIndex        =   306
            Top             =   825
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   41
            Left            =   5760
            TabIndex        =   335
            Top             =   2715
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ��ʽ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   44
            Left            =   360
            TabIndex        =   313
            Top             =   1320
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ת��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   24
            Left            =   4860
            TabIndex        =   315
            Top             =   1320
            Width           =   360
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "  ��Ժ��        ��         Сʱ        ����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   38
            Left            =   6945
            TabIndex        =   328
            Top             =   2325
            Width           =   3870
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������ʹ��            Сʱ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   34
            Left            =   240
            TabIndex        =   332
            Top             =   2715
            Width           =   2340
         End
         Begin VB.Label lblNumInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժǰ        ��         Сʱ        ����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   2565
            TabIndex        =   324
            Top             =   2325
            Width           =   3690
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�Ƿ���"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   49
            Left            =   540
            TabIndex        =   318
            Top             =   1815
            Width           =   540
         End
         Begin VB.Label lblNumInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "­�����˻��߻���ʱ��:"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   35
            Left            =   270
            TabIndex        =   323
            Top             =   2325
            Width           =   1890
         End
         Begin VB.Line lineH 
            Index           =   11
            X1              =   0
            X2              =   14000
            Y1              =   1160
            Y2              =   1160
         End
         Begin VB.Line lineH 
            Index           =   12
            X1              =   0
            X2              =   14000
            Y1              =   1660
            Y2              =   1660
         End
         Begin VB.Line lineH 
            BorderStyle     =   2  'Dash
            DrawMode        =   1  'Blackness
            Index           =   13
            X1              =   0
            X2              =   14000
            Y1              =   2160
            Y2              =   2160
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2080
         Index           =   8
         Left            =   732
         ScaleHeight     =   2085
         ScaleWidth      =   10995
         TabIndex        =   383
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   25680
         Width           =   11000
         Begin VB.CommandButton cmdDateInfo 
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
            Height          =   240
            Index           =   8
            Left            =   10380
            Picture         =   "frmInMedRecEdit.frx":10C9F
            Style           =   1  'Graphical
            TabIndex        =   282
            TabStop         =   0   'False
            Top             =   1665
            Width           =   270
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   225
            Index           =   8
            Left            =   9225
            TabIndex        =   280
            Tag             =   "####-##-##"
            Top             =   1680
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   8
            Left            =   9225
            MaxLength       =   30
            TabIndex        =   281
            Top             =   1673
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   8
            ItemData        =   "frmInMedRecEdit.frx":10D95
            Left            =   3705
            List            =   "frmInMedRecEdit.frx":10D97
            TabIndex        =   276
            Top             =   1635
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   9
            ItemData        =   "frmInMedRecEdit.frx":10D99
            Left            =   6465
            List            =   "frmInMedRecEdit.frx":10D9B
            TabIndex        =   278
            Top             =   1640
            Width           =   1185
         End
         Begin VB.CommandButton cmdSign 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ǩ��"
            Enabled         =   0   'False
            Height          =   300
            Index           =   3
            Left            =   10440
            TabIndex        =   262
            TabStop         =   0   'False
            Top             =   345
            Width           =   520
         End
         Begin VB.CommandButton cmdSign 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ǩ��"
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   4920
            TabIndex        =   256
            TabStop         =   0   'False
            Top             =   345
            Width           =   520
         End
         Begin VB.CommandButton cmdSign 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ǩ��"
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   7680
            TabIndex        =   259
            TabStop         =   0   'False
            Top             =   345
            Width           =   520
         End
         Begin VB.CommandButton cmdSign 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ǩ��"
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   2040
            TabIndex        =   253
            TabStop         =   0   'False
            Top             =   345
            Width           =   520
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   7
            ItemData        =   "frmInMedRecEdit.frx":10D9D
            Left            =   6465
            List            =   "frmInMedRecEdit.frx":10D9F
            TabIndex        =   268
            Top             =   735
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   6
            ItemData        =   "frmInMedRecEdit.frx":10DA1
            Left            =   9225
            List            =   "frmInMedRecEdit.frx":10DA3
            TabIndex        =   270
            Top             =   735
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   3
            ItemData        =   "frmInMedRecEdit.frx":10DA5
            Left            =   3705
            List            =   "frmInMedRecEdit.frx":10DA7
            TabIndex        =   266
            Top             =   735
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   5
            ItemData        =   "frmInMedRecEdit.frx":10DA9
            Left            =   9225
            List            =   "frmInMedRecEdit.frx":10DAB
            TabIndex        =   261
            Top             =   340
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   4
            ItemData        =   "frmInMedRecEdit.frx":10DAD
            Left            =   6465
            List            =   "frmInMedRecEdit.frx":10DAF
            TabIndex        =   258
            Top             =   340
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            ItemData        =   "frmInMedRecEdit.frx":10DB1
            Left            =   3705
            List            =   "frmInMedRecEdit.frx":10DB3
            TabIndex        =   255
            Top             =   345
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            ItemData        =   "frmInMedRecEdit.frx":10DB5
            Left            =   825
            List            =   "frmInMedRecEdit.frx":10DB7
            TabIndex        =   252
            Top             =   340
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   10
            ItemData        =   "frmInMedRecEdit.frx":10DB9
            Left            =   825
            List            =   "frmInMedRecEdit.frx":10DBB
            TabIndex        =   264
            Top             =   740
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            ItemData        =   "frmInMedRecEdit.frx":10DBD
            Left            =   825
            List            =   "frmInMedRecEdit.frx":10DBF
            TabIndex        =   272
            Top             =   1140
            Width           =   1185
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   33
            ItemData        =   "frmInMedRecEdit.frx":10DC1
            Left            =   825
            List            =   "frmInMedRecEdit.frx":10DC3
            Style           =   2  'Dropdown List
            TabIndex        =   274
            Top             =   1640
            Width           =   1185
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ǩ����Ϣ"
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
            Index           =   8
            Left            =   0
            TabIndex        =   250
            Top             =   0
            Width           =   900
         End
         Begin VB.Label lblManInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(������)    ҽʦ"
            Height          =   360
            Index           =   2
            Left            =   2640
            TabIndex        =   254
            Top             =   315
            Width           =   1080
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblManInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ʿ�ҽʦ"
            Height          =   180
            Index           =   8
            Left            =   2910
            TabIndex        =   275
            Top             =   1695
            Width           =   720
         End
         Begin VB.Label lblManInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ʿػ�ʿ"
            Height          =   180
            Index           =   9
            Left            =   5670
            TabIndex        =   277
            Top             =   1695
            Width           =   720
         End
         Begin VB.Label lblManInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ʵϰҽʦ"
            Height          =   180
            Index           =   7
            Left            =   5670
            TabIndex        =   267
            Top             =   795
            Width           =   720
         End
         Begin VB.Label lblManInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�о���ҽʦ"
            Height          =   180
            Index           =   6
            Left            =   8250
            TabIndex        =   269
            Top             =   795
            Width           =   900
         End
         Begin VB.Label lblManInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҽʦ"
            Height          =   180
            Index           =   3
            Left            =   2910
            TabIndex        =   265
            Top             =   795
            Width           =   720
         End
         Begin VB.Label lblManInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺҽʦ"
            Height          =   180
            Index           =   5
            Left            =   8430
            TabIndex        =   260
            Top             =   405
            Width           =   720
         End
         Begin VB.Label lblManInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҽʦ"
            Height          =   180
            Index           =   4
            Left            =   5670
            TabIndex        =   257
            Top             =   405
            Width           =   720
         End
         Begin VB.Label lblManInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   180
            Index           =   1
            Left            =   210
            TabIndex        =   251
            Top             =   405
            Width           =   540
         End
         Begin VB.Label lblManInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���λ�ʿ"
            Height          =   180
            Index           =   10
            Left            =   30
            TabIndex        =   263
            Top             =   795
            Width           =   720
         End
         Begin VB.Label lblManInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҽʦ"
            Height          =   180
            Index           =   0
            Left            =   30
            TabIndex        =   271
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Index           =   33
            Left            =   30
            TabIndex        =   273
            Top             =   1695
            Width           =   720
         End
         Begin VB.Label lblDateInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�ʿ�����"
            Height          =   180
            Index           =   8
            Left            =   8430
            TabIndex        =   279
            Top             =   1695
            Width           =   720
         End
         Begin VB.Line lineH 
            Index           =   10
            X1              =   0
            X2              =   14400
            Y1              =   1540
            Y2              =   1540
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2240
         Index           =   7
         Left            =   732
         ScaleHeight     =   2235
         ScaleWidth      =   10995
         TabIndex        =   382
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   23460
         Width           =   11000
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   38
            ItemData        =   "frmInMedRecEdit.frx":10DC5
            Left            =   5580
            List            =   "frmInMedRecEdit.frx":10DC7
            Style           =   2  'Dropdown List
            TabIndex        =   233
            Top             =   225
            Width           =   1200
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   36
            ItemData        =   "frmInMedRecEdit.frx":10DC9
            Left            =   1920
            List            =   "frmInMedRecEdit.frx":10DCB
            Style           =   2  'Dropdown List
            TabIndex        =   223
            Top             =   225
            Width           =   1425
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   40
            ItemData        =   "frmInMedRecEdit.frx":10DCD
            Left            =   1920
            List            =   "frmInMedRecEdit.frx":10DCF
            Style           =   2  'Dropdown List
            TabIndex        =   227
            Top             =   1005
            Width           =   1425
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   42
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   231
            Top             =   1800
            Width           =   1410
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   35
            ItemData        =   "frmInMedRecEdit.frx":10DD1
            Left            =   9210
            List            =   "frmInMedRecEdit.frx":10DD3
            Style           =   2  'Dropdown List
            TabIndex        =   243
            Top             =   600
            Width           =   1425
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   37
            ItemData        =   "frmInMedRecEdit.frx":10DD5
            Left            =   9210
            List            =   "frmInMedRecEdit.frx":10DD7
            Style           =   2  'Dropdown List
            TabIndex        =   245
            Top             =   1005
            Width           =   1425
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   39
            ItemData        =   "frmInMedRecEdit.frx":10DD9
            Left            =   9210
            List            =   "frmInMedRecEdit.frx":10DDB
            Style           =   2  'Dropdown List
            TabIndex        =   247
            Top             =   1395
            Width           =   1425
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   41
            ItemData        =   "frmInMedRecEdit.frx":10DDD
            Left            =   1920
            List            =   "frmInMedRecEdit.frx":10DDF
            Style           =   2  'Dropdown List
            TabIndex        =   229
            Top             =   1395
            Width           =   1425
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   23
            Left            =   9210
            MaxLength       =   30
            TabIndex        =   249
            Top             =   1845
            Width           =   1380
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   33
            Left            =   5580
            MaxLength       =   5
            TabIndex        =   241
            Top             =   1845
            Width           =   1140
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   32
            Left            =   5580
            MaxLength       =   5
            TabIndex        =   239
            Top             =   1440
            Width           =   1140
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   31
            Left            =   5580
            MaxLength       =   5
            TabIndex        =   237
            Top             =   1035
            Width           =   1140
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   30
            Left            =   5580
            MaxLength       =   5
            TabIndex        =   235
            Top             =   645
            Width           =   1140
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   29
            Left            =   1920
            MaxLength       =   5
            TabIndex        =   225
            Top             =   645
            Width           =   1140
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��Ѫ��Ϣ"
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
            Index           =   7
            Left            =   0
            TabIndex        =   221
            Top             =   0
            Width           =   900
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ѫ��"
            Height          =   180
            Index           =   36
            Left            =   1485
            TabIndex        =   222
            Top             =   285
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rh"
            Height          =   180
            Index           =   38
            Left            =   5325
            TabIndex        =   232
            Top             =   285
            Width           =   180
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Һ��Ӧ"
            Height          =   180
            Index           =   40
            Left            =   1125
            TabIndex        =   226
            Top             =   1065
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ѫǰ��9����"
            Height          =   180
            Index           =   42
            Left            =   495
            TabIndex        =   230
            Top             =   1860
            Width           =   1350
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "HB&sAg"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   35
            Left            =   8685
            TabIndex        =   242
            Top             =   660
            Width           =   450
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HC&V-Ab"
            Height          =   180
            Index           =   37
            Left            =   8595
            TabIndex        =   244
            Top             =   1065
            Width           =   540
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "H&IV-Ab"
            Height          =   180
            Index           =   39
            Left            =   8595
            TabIndex        =   246
            Top             =   1455
            Width           =   540
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ѫ��Ӧ"
            Height          =   180
            Index           =   41
            Left            =   1125
            TabIndex        =   228
            Top             =   1455
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   180
            Index           =   23
            Left            =   8595
            TabIndex        =   248
            Top             =   1860
            Width           =   540
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ѫ��               ml"
            Height          =   180
            Index           =   31
            Left            =   4980
            TabIndex        =   236
            Top             =   1065
            Width           =   2070
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���ϸ��              ��λ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   29
            Left            =   1110
            TabIndex        =   224
            Top             =   660
            Width           =   2340
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ѪС��              ��λ"
            Height          =   180
            Index           =   30
            Left            =   4830
            TabIndex        =   234
            Top             =   660
            Width           =   2340
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ȫѪ               ml"
            Height          =   180
            Index           =   32
            Left            =   4980
            TabIndex        =   238
            Top             =   1455
            Width           =   2070
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������               ml"
            Height          =   180
            Index           =   33
            Left            =   4800
            TabIndex        =   240
            Top             =   1860
            Width           =   2250
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1850
         Index           =   6
         Left            =   732
         ScaleHeight     =   1845
         ScaleWidth      =   10995
         TabIndex        =   381
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   21615
         Width           =   11000
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "�޹�����¼"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   30
            Left            =   1800
            TabIndex        =   217
            Top             =   150
            Width           =   1215
         End
         Begin VB.CommandButton cmdAutoLoad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�Զ���ȡ"
            Height          =   330
            Index           =   2
            Left            =   9860
            TabIndex        =   398
            TabStop         =   0   'False
            Top             =   60
            Width           =   1000
         End
         Begin VB.OptionButton optAller 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���ݹ���Դ����"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   1
            Left            =   7950
            TabIndex        =   219
            TabStop         =   0   'False
            Top             =   150
            Width           =   1650
         End
         Begin VB.OptionButton optAller 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����ҩƷĿ¼����"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   6120
            TabIndex        =   218
            TabStop         =   0   'False
            Top             =   150
            Width           =   1770
         End
         Begin VSFlex8Ctl.VSFlexGrid vsAller 
            Height          =   1200
            Left            =   0
            TabIndex        =   220
            Top             =   450
            Width           =   10990
            _cx             =   19385
            _cy             =   2117
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit.frx":10DE1
            ScrollTrack     =   -1  'True
            ScrollBars      =   0
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
            Editable        =   2
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
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ҩ�����"
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
            Index           =   6
            Left            =   0
            TabIndex        =   216
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2620
         Index           =   5
         Left            =   732
         ScaleHeight     =   2625
         ScaleWidth      =   10995
         TabIndex        =   380
         TabStop         =   0   'False
         Top             =   18990
         Visible         =   0   'False
         Width           =   11000
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   23
            ItemData        =   "frmInMedRecEdit.frx":10E93
            Left            =   6540
            List            =   "frmInMedRecEdit.frx":10E95
            Style           =   2  'Dropdown List
            TabIndex        =   191
            Top             =   280
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   22
            ItemData        =   "frmInMedRecEdit.frx":10E97
            Left            =   2880
            List            =   "frmInMedRecEdit.frx":10E99
            Style           =   2  'Dropdown List
            TabIndex        =   189
            Top             =   280
            Width           =   1635
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Σ��"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   6
            Left            =   2115
            TabIndex        =   193
            Top             =   805
            Width           =   900
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   8
            Left            =   6510
            TabIndex        =   195
            Top             =   805
            Width           =   900
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��֢"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   7
            Left            =   4320
            TabIndex        =   194
            Top             =   805
            Width           =   900
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   28
            ItemData        =   "frmInMedRecEdit.frx":10E9B
            Left            =   2280
            List            =   "frmInMedRecEdit.frx":10E9D
            Style           =   2  'Dropdown List
            TabIndex        =   211
            Top             =   2180
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   30
            ItemData        =   "frmInMedRecEdit.frx":10E9F
            Left            =   5820
            List            =   "frmInMedRecEdit.frx":10EA1
            Style           =   2  'Dropdown List
            TabIndex        =   213
            Top             =   2175
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   32
            ItemData        =   "frmInMedRecEdit.frx":10EA3
            Left            =   9090
            List            =   "frmInMedRecEdit.frx":10EA5
            Style           =   2  'Dropdown List
            TabIndex        =   215
            Top             =   2175
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   24
            ItemData        =   "frmInMedRecEdit.frx":10EA7
            Left            =   2280
            List            =   "frmInMedRecEdit.frx":10EA9
            Style           =   2  'Dropdown List
            TabIndex        =   198
            Top             =   1280
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   25
            ItemData        =   "frmInMedRecEdit.frx":10EAB
            Left            =   5820
            List            =   "frmInMedRecEdit.frx":10EAD
            Style           =   2  'Dropdown List
            TabIndex        =   200
            Top             =   1275
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   26
            ItemData        =   "frmInMedRecEdit.frx":10EAF
            Left            =   9090
            List            =   "frmInMedRecEdit.frx":10EB1
            Style           =   2  'Dropdown List
            TabIndex        =   202
            Top             =   1275
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   31
            ItemData        =   "frmInMedRecEdit.frx":10EB3
            Left            =   9090
            List            =   "frmInMedRecEdit.frx":10EB5
            Style           =   2  'Dropdown List
            TabIndex        =   209
            Top             =   1785
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   29
            ItemData        =   "frmInMedRecEdit.frx":10EB7
            Left            =   5820
            List            =   "frmInMedRecEdit.frx":10EB9
            Style           =   2  'Dropdown List
            TabIndex        =   207
            Top             =   1785
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   27
            ItemData        =   "frmInMedRecEdit.frx":10EBB
            Left            =   2280
            List            =   "frmInMedRecEdit.frx":10EBD
            Style           =   2  'Dropdown List
            TabIndex        =   205
            Top             =   1780
            Width           =   1635
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ҽ��Ϸ������"
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
            Index           =   5
            Left            =   0
            TabIndex        =   187
            Top             =   0
            Width           =   1800
         End
         Begin VB.Line lineH 
            Index           =   9
            X1              =   0
            X2              =   14400
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line lineH 
            Index           =   7
            X1              =   0
            X2              =   14400
            Y1              =   680
            Y2              =   680
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ���Ժ"
            Height          =   180
            Index           =   23
            Left            =   5565
            TabIndex        =   190
            Top             =   345
            Width           =   900
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�������Ժ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   22
            Left            =   1905
            TabIndex        =   188
            Top             =   345
            Width           =   900
         End
         Begin VB.Label lblZY 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ�ڼ䲡��:"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   192
            Top             =   840
            Width           =   1170
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ʹ����ҽ�����豸"
            Height          =   180
            Index           =   28
            Left            =   765
            TabIndex        =   210
            Top             =   2235
            Width           =   1440
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ʹ����ҽ���Ƽ���"
            Height          =   180
            Index           =   30
            Left            =   4305
            TabIndex        =   212
            Top             =   2235
            Width           =   1440
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��֤ʩ��"
            Height          =   180
            Index           =   32
            Left            =   8295
            TabIndex        =   214
            Top             =   2235
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��֤"
            Height          =   180
            Index           =   24
            Left            =   1845
            TabIndex        =   197
            Top             =   1335
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�η�"
            Height          =   180
            Index           =   25
            Left            =   5385
            TabIndex        =   199
            Top             =   1335
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩ"
            Height          =   180
            Index           =   26
            Left            =   8655
            TabIndex        =   201
            Top             =   1335
            Width           =   360
         End
         Begin VB.Label lblZY 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "׼ȷ��:"
            Height          =   180
            Index           =   1
            Left            =   540
            TabIndex        =   196
            Top             =   1335
            Width           =   630
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������ҩ�Ƽ�"
            Height          =   180
            Index           =   31
            Left            =   7935
            TabIndex        =   208
            Top             =   1845
            Width           =   1080
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ȷ���"
            Height          =   180
            Index           =   29
            Left            =   5025
            TabIndex        =   206
            Top             =   1845
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������"
            Height          =   180
            Index           =   27
            Left            =   1485
            TabIndex        =   204
            Top             =   1845
            Width           =   720
         End
         Begin VB.Label lblZY 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���Ʒ���:"
            Height          =   180
            Index           =   2
            Left            =   360
            TabIndex        =   203
            Top             =   1845
            Width           =   810
         End
         Begin VB.Line lineH 
            Index           =   8
            X1              =   0
            X2              =   14400
            Y1              =   1180
            Y2              =   1180
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1750
         Index           =   15
         Left            =   732
         ScaleHeight     =   1755
         ScaleWidth      =   10995
         TabIndex        =   390
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   38955
         Width           =   11000
         Begin VSFlex8Ctl.VSFlexGrid vsSpirit 
            Height          =   1200
            Left            =   0
            TabIndex        =   348
            Top             =   350
            Width           =   10990
            _cx             =   19385
            _cy             =   2117
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit.frx":10EBF
            ScrollTrack     =   0   'False
            ScrollBars      =   1
            ScrollTips      =   0   'False
            MergeCells      =   120
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
            Editable        =   2
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
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�������������"
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
            Index           =   15
            Left            =   0
            TabIndex        =   347
            Top             =   0
            Width           =   1800
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1900
         Index           =   14
         Left            =   732
         ScaleHeight     =   1905
         ScaleWidth      =   10995
         TabIndex        =   389
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   37065
         Width           =   11000
         Begin VB.CommandButton cmdAutoLoad 
            Appearance      =   0  'Flat
            Caption         =   "�Զ���ȡ"
            Height          =   330
            Index           =   0
            Left            =   9840
            TabIndex        =   345
            TabStop         =   0   'False
            Top             =   90
            Width           =   1000
         End
         Begin VSFlex8Ctl.VSFlexGrid vsKSS 
            Height          =   1200
            Left            =   0
            TabIndex        =   346
            Top             =   500
            Width           =   10990
            _cx             =   19385
            _cy             =   2117
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
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
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit.frx":10FA0
            ScrollTrack     =   -1  'True
            ScrollBars      =   1
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
            Editable        =   2
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
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ҩ��ʹ���������DDD���������У�"
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
            Index           =   14
            Left            =   0
            TabIndex        =   344
            Top             =   0
            Width           =   3960
         End
      End
      Begin MSComCtl2.MonthView monInfo 
         Height          =   2160
         Left            =   960
         TabIndex        =   394
         TabStop         =   0   'False
         Top             =   3000
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   3810
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         StartOfWeek     =   188678145
         TitleBackColor  =   8421504
         TitleForeColor  =   16777215
         CurrentDate     =   38003
      End
   End
   Begin zlSubclass.Subclass subcMain 
      Left            =   0
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Image imgButtonNew 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   500
      Picture         =   "frmInMedRecEdit.frx":110E4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgButtonDel 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   0
      Picture         =   "frmInMedRecEdit.frx":1166E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmInMedRecEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboManInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call ManInfoKeyDown(Index, KeyCode)
End Sub

Private Sub cmdTop_Click()
    Call cmdTopClick
End Sub

Private Sub cmdTop_GotFocus()
    Call cmdTopGotFocus
End Sub

Private Sub Form_Load()
    Call FormLoad
End Sub

Private Sub Form_Resize()
    Call FormResize
End Sub

Private Sub hsbMain_Change()
    Call hsbMainChange
End Sub

Private Sub lstAdvEvent_LostFocus()
    Call lstLostFocus(lstAdvEvent)
End Sub

Private Sub lstInfection_LostFocus()
    Call lstLostFocus(lstInfection)
End Sub

Private Sub lstInfectParts_LostFocus()
    Call lstLostFocus(lstInfectParts)
End Sub

Private Sub padrInfo_Change(Index As Integer)
    Call CheckValueChange
End Sub

Private Sub padrInfo_SetInput(Index As Integer, ByVal intLevel As Integer, rsReturn As ADODB.Recordset)
    Call SetYoubian(Index, intLevel, rsReturn)
End Sub

Private Sub txtAdressInfo_Change(Index As Integer)
    Call txtAdressInfoChange(Index)
End Sub

Private Sub txtDateInfo_Change(Index As Integer)
    Call CheckValueChange(txtDateInfo(Index))
End Sub

Private Sub txtDateInfo_GotFocus(Index As Integer)
    Call txtDateInfoGotFocus(Index)
End Sub

Private Sub vsAller_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsAller)
End Sub

Private Sub vsbMain_Change()
   Call vsbMainChange
End Sub

Private Sub subcMain_WndProc(msg As Long, wParam As Long, lParam As Long, Result As Long)
    Call SubCMainWndProc(msg, wParam, lParam, Result)
End Sub

Private Sub cboBaseInfo_Click(Index As Integer)
    Call CboBaseInfoClick(Index)
End Sub

Private Sub cboBaseInfo_GotFocus(Index As Integer)
    Call CboBaseInfoGotFocus(Index)
End Sub

Private Sub cboBaseInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CboBaseInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub cboBaseInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CboBaseInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub cboBaseInfo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CboBaseInfoKeyUp(Index, KeyCode, Shift)
End Sub

Private Sub cboBaseInfo_Validate(Index As Integer, Cancel As Boolean)
    Call cboBaseInfoValidate(Index, Cancel)
End Sub

Private Sub cboManInfo_Click(Index As Integer)
    Call ManInfoClick(Index)
End Sub

Private Sub cboManInfo_DropDown(Index As Integer)
    Call ManInfoDropDown(Index)
End Sub

Private Sub cboManInfo_GotFocus(Index As Integer)
    Call ManInfoGotFocus(Index)
End Sub

Private Sub cboManInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call ManInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub cboManInfo_LostFocus(Index As Integer)
    Call ManInfoLostFocus(Index)
End Sub

Private Sub cboManInfo_Validate(Index As Integer, Cancel As Boolean)
    Call ManInfoValidate(Index, Cancel)
End Sub

Private Sub cboBaseInfo_Change(Index As Integer)
    Call CboBaseInfoChange(Index)
End Sub

Private Sub CboSpecificInfo_Click(Index As Integer)
    Call CboSpecificInfoClick(Index)
End Sub

Private Sub cboSpecificInfo_GotFocus(Index As Integer)
    Call CboSpecificInfoGotFocus(Index)
End Sub

Private Sub cboSpecificInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CboSpecificInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub cboSpecificInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call cboSpecificInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub chkInfo_Click(Index As Integer)
    Call chkInfoClick(Index)
End Sub

Private Sub chkInfo_GotFocus(Index As Integer)
    Call ChkInfoGotFocus(Index)
End Sub

Private Sub chkInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call ChkInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub chkParaOPSInfo_Click(Index As Integer)
    Call ChkParaOPSInfoClick(Index)
End Sub

Private Sub chkParaOPSInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call ChkParaOPSInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub cmdAdressInfo_Click(Index As Integer)
    Call CmdAdressInfoClick(Index)
End Sub

Private Sub cmdAutoLoad_Click(Index As Integer)
    Call CmdAutoLoadClick(Index)
End Sub

Private Sub cmdDateInfo_Click(Index As Integer)
    Call DateInfoClick(Index)
End Sub

Private Sub cmdDiagMove_Click(Index As Integer)
    Call CmdDiagMoveClick(Index)
End Sub

Private Sub cmdDiagMove_GotFocus(Index As Integer)
    Call CmdDiagMoveGotFocus(Index)
End Sub

Private Sub cmdInfo_Click(Index As Integer)
    Call CmdInfoClick(Index)
End Sub

Private Sub cmdOPSMove_Click(Index As Integer)
    Call cmdOPSMoveClick(Index)
End Sub

Private Sub cmdSign_Click(Index As Integer)
    Call CmdSignClick(Index)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FormKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call FormKeyPress(KeyAscii)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call FormUnLoad(Cancel)
End Sub

Private Sub lstAdvEvent_GotFocus()
    Call LstGotFocus(lstAdvEvent)
End Sub

Private Sub lstAdvEvent_ItemCheck(Item As Integer)
    Call LstItemCheck(lstAdvEvent, Item)
End Sub

Private Sub lstAdvEvent_KeyDown(KeyCode As Integer, Shift As Integer)
    Call LstKeyDown(lstAdvEvent, KeyCode, Shift)
End Sub

Private Sub lstAdvEvent_KeyPress(KeyAscii As Integer)
    Call LstKeyPress(lstAdvEvent, KeyAscii)
End Sub

Private Sub lstInfection_GotFocus()
    Call LstGotFocus(lstInfection)
End Sub

Private Sub lstInfection_ItemCheck(Item As Integer)
    Call LstItemCheck(lstInfection, Item)
End Sub

Private Sub lstInfection_KeyDown(KeyCode As Integer, Shift As Integer)
    Call LstKeyDown(lstInfection, KeyCode, Shift)
End Sub

Private Sub lstInfection_KeyPress(KeyAscii As Integer)
    Call LstKeyPress(lstInfection, KeyAscii)
End Sub

Private Sub lstInfectParts_GotFocus()
    Call LstGotFocus(lstInfectParts)
End Sub

Private Sub lstInfectParts_ItemCheck(Item As Integer)
    Call LstItemCheck(lstInfectParts, Item)
End Sub

Private Sub lstInfectParts_KeyDown(KeyCode As Integer, Shift As Integer)
    Call LstKeyDown(lstInfectParts, KeyCode, Shift)
End Sub

Private Sub lstInfectParts_KeyPress(KeyAscii As Integer)
    Call LstKeyPress(lstInfectParts, KeyAscii)
End Sub

Private Sub monInfo_DateClick(ByVal DateClicked As Date)
    Call monInfoDateClick(DateClicked)
End Sub

Private Sub monInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call monInfoKeyDown(KeyCode, Shift)
End Sub

Private Sub monInfo_KeyPress(KeyAscii As Integer)
    Call monInfoKeyPress(KeyAscii)
End Sub

Private Sub monInfo_Validate(Cancel As Boolean)
    Call monInfoValidate(Cancel)
End Sub

Private Sub mskDateInfo_Change(Index As Integer)
    Call DateInfoChange(Index)
End Sub

Private Sub mskDateInfo_GotFocus(Index As Integer)
    Call DateInfoGotFocus(Index)
End Sub

Private Sub mskDateInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call DateInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub mskDateInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call DateInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub mskDateInfo_Validate(Index As Integer, Cancel As Boolean)
    Call DateInfoValidate(Index, Cancel)
End Sub

Private Sub optAller_Click(Index As Integer)
    Call OptAllerClick(Index)
End Sub

Private Sub optAller_KeyPress(Index As Integer, KeyAscii As Integer)
    Call OptAllerKeyPress(Index, KeyAscii)
End Sub

Private Sub optInput_Click(Index As Integer)
    Call OptInputClick(Index)
End Sub

Private Sub optInput_KeyPress(Index As Integer, KeyAscii As Integer)
    Call OptInputKeyPress(Index, KeyAscii)
End Sub

Private Sub optDiag_Click(Index As Integer)
    Call optDiagClick(Index)
End Sub

Private Sub optDiag_GotFocus(Index As Integer)
    Call optDiagGotFocus(Index)
End Sub

Private Sub optDiag_KeyPress(Index As Integer, KeyAscii As Integer)
    Call optDiagKeyPress(Index, KeyAscii)
End Sub

Private Sub optParaOPSInfo_Click(Index As Integer)
    Call OptParaOPSInfoClick(Index)
End Sub

Private Sub optParaOPSInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call OptParaOPSInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub picCopy_Click()
    Call picCopyClick
End Sub

Private Sub txtAdressInfo_GotFocus(Index As Integer)
    Call txtAdressInfoGotFocus(Index)
End Sub

Private Sub txtAdressInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call txtAdressInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub txtAdressInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call txtAdressInfoMouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub txtAdressInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call txtAdressInfoMouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub txtInfo_Change(Index As Integer)
    Call TxtInfoChange(Index)
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Call TxtInfoGotFocus(Index)
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call TxtInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call TxtInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub txtInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call TxtInfoMouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub txtInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call TxtInfoMouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    Call TxtInfoValidate(Index, Cancel)
End Sub

Private Sub txtSpecificInfo_Change(Index As Integer)
    Call SpecificInfoChange(Index)
End Sub

Private Sub txtSpecificInfo_GotFocus(Index As Integer)
    Call SpecificInfoGotFocus(Index)
End Sub

Private Sub txtSpecificInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call SpecificInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub txtSpecificInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call SpecificInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub txtSpecificInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SpecificInfoMouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub txtSpecificInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SpecificInfoMouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub txtSpecificInfo_Validate(Index As Integer, Cancel As Boolean)
    Call SpecificInfoValidate(Index, Cancel)
End Sub

Private Sub vsAller_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call AllerAfterEdit(vsAller, Row, Col)
End Sub

Private Sub vsAller_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call AllerAfterRowColChange(vsAller, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsAller_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call AllerCellButtonClick(vsAller, Row, Col)
End Sub

Private Sub vsAller_KeyDown(KeyCode As Integer, Shift As Integer)
    Call AllerKeyDown(vsAller, KeyCode, Shift)
End Sub

Private Sub vsAller_KeyPress(KeyAscii As Integer)
    Call AllerKeyPress(vsAller, KeyAscii)
End Sub

Private Sub vsAller_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call AllerKeyPressEdit(vsAller, Row, Col, KeyAscii)
End Sub

Private Sub vsAller_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call AllerSetupEditWindow(vsAller, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsAller_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call AllerStartEdit(vsAller, Row, Col, Cancel)
End Sub

Private Sub vsAller_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call AllerValidateEdit(vsAller, Row, Col, Cancel)
End Sub

Private Sub vsChemoth_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call ChemothAfterEdit(vsChemoth, Row, Col)
End Sub

Private Sub vsChemoth_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call ChemothAfterRowColChange(vsChemoth, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsChemoth_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsChemoth)
End Sub

Private Sub vsChemoth_GotFocus()
    Call VSFlxGotFocus(vsChemoth)
End Sub

Private Sub vsChemoth_KeyDown(KeyCode As Integer, Shift As Integer)
    Call ChemothKeyDown(vsChemoth, KeyCode, Shift)
End Sub

Private Sub vsChemoth_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Call ChemothKeyDownEdit(vsChemoth, Row, Col, KeyCode, Shift)
End Sub

Private Sub vsChemoth_KeyPress(KeyAscii As Integer)
    Call ChemothKeyPress(vsChemoth, KeyAscii)
End Sub

Private Sub vsChemoth_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call ChemothKeyPressEdit(vsChemoth, Row, Col, KeyAscii)
End Sub

Private Sub vsChemoth_LostFocus()
    Call ChemothLostFocus(vsChemoth)
End Sub

Private Sub vsChemoth_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call ChemothValidateEdit(vsChemoth, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterEdit(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call DiagAfterRowColChange(vsDiagXY, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsDiagXY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagBeforeUserResize(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call DiagCellButtonClick(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsDiagXY)
End Sub

Private Sub vsDiagXY_Click()
    Call DiagClick(vsDiagXY)
End Sub

Private Sub vsDiagXY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Call DiagComboDropDown(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_DblClick()
    Call DiagDblClick(vsDiagXY)
End Sub

Private Sub vsDiagXY_GotFocus()
    Call DiagGotFocus(vsDiagXY)
End Sub

Private Sub vsDiagXY_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiagKeyDown(vsDiagXY, KeyCode, Shift)
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    Call DiagKeyPress(vsDiagXY, KeyAscii)
End Sub

Private Sub vsDiagXY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call DiagKeyPressEdit(vsDiagXY, Row, Col, KeyAscii)
End Sub

Private Sub vsDiagXY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call DiagSetupEditWindow(vsDiagXY, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsDiagXY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagStartEdit(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagValidateEdit(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterEdit(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call DiagAfterRowColChange(vsDiagZY, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsDiagZY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagBeforeUserResize(vsDiagZY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call DiagCellButtonClick(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsDiagZY)
End Sub

Private Sub vsDiagZY_Click()
    Call DiagClick(vsDiagZY)
End Sub

Private Sub vsDiagZY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Call DiagComboDropDown(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_DblClick()
    Call DiagDblClick(vsDiagZY)
End Sub

Private Sub vsDiagZY_GotFocus()
    Call DiagGotFocus(vsDiagZY)
End Sub

Private Sub vsDiagZY_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiagKeyDown(vsDiagZY, KeyCode, Shift)
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    Call DiagKeyPress(vsDiagZY, KeyAscii)
End Sub

Private Sub vsDiagZY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call DiagKeyPressEdit(vsDiagZY, Row, Col, KeyAscii)
End Sub

Private Sub vsDiagZY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call DiagSetupEditWindow(vsDiagZY, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsDiagZY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagStartEdit(vsDiagZY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagValidateEdit(vsDiagZY, Row, Col, Cancel)
End Sub

Private Sub vsFlxAddICU_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call FlxAddICUAfterEdit(vsFlxAddICU, Row, Col)
End Sub

Private Sub vsFlxAddICU_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call FlxAddICUCellButtonClick(vsFlxAddICU, Row, Col)
End Sub

Private Sub vsFlxAddICU_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsFlxAddICU)
End Sub

Private Sub vsFlxAddICU_EnterCell()
    Call FlxAddICUEnterCell(vsFlxAddICU)
End Sub

Private Sub vsFlxAddICU_GotFocus()
    Call VSFlxGotFocus(vsFlxAddICU)
End Sub

Private Sub vsFlxAddICU_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FlxAddICUKeyDown(vsFlxAddICU, KeyCode, Shift)
End Sub

Private Sub vsFlxAddICU_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Call FlxAddICUKeyDownEdit(vsFlxAddICU, Row, Col, KeyCode, Shift)
End Sub

Private Sub vsfMain_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsfMain)
End Sub

Private Sub vsfMain_EnterCell()
    Call vsfMainEnterCell(vsfMain)
End Sub

Private Sub vsfMain_GotFocus()
    Call VSFlxGotFocus(vsfMain)
End Sub

Private Sub vsfMain_KeyPress(KeyAscii As Integer)
    Call vsfMainKeyPress(vsfMain, KeyAscii)
End Sub

Private Sub vsfMain_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call vsfMainStartEdit(vsfMain, Row, Col, Cancel)
End Sub

Private Sub vsfMain_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call vsfMainValidateEdit(vsfMain, Row, Col, Cancel)
End Sub

Private Sub vsKSS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call KSSAfterEdit(vsKSS, Row, Col)
End Sub

Private Sub vsKSS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call KSSAfterRowColChange(vsKSS, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsKSS_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call KSSCellButtonClick(vsKSS, Row, Col)
End Sub

Private Sub vsKSS_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsKSS)
End Sub

Private Sub vsKSS_GotFocus()
    Call VSFlxGotFocus(vsKSS)
End Sub

Private Sub vsKSS_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KSSKeyDown(vsKSS, KeyCode, Shift)
End Sub

Private Sub vsKSS_KeyPress(KeyAscii As Integer)
    Call KSSKeyPress(vsKSS, KeyAscii)
End Sub

Private Sub vsKSS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call KSSKeyPressEdit(vsKSS, Row, Col, KeyAscii)
End Sub

Private Sub vsKSS_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call KSSSetupEditWindow(vsKSS, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsKSS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call KSSValidateEdit(vsKSS, Row, Col, Cancel)
End Sub

Private Sub vsOPS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call OPSAfterEdit(vsOPS, Row, Col)
End Sub

Private Sub vsOPS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call OPSAfterRowColChange(vsOPS, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsOPS_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call OPSBeforeUserResize(vsOPS, Row, Col, Cancel)
End Sub

Private Sub vsOPS_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call OPSCellButtonClick(vsOPS, Row, Col)
End Sub

Private Sub vsOPS_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsOPS)
End Sub

Private Sub vsOPS_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Call OPSComboDropDown(vsOPS, Row, Col)
End Sub

Private Sub vsOPS_DblClick()
    Call OPSDblClick(vsOPS)
End Sub

Private Sub vsOPS_KeyDown(KeyCode As Integer, Shift As Integer)
    Call OPSKeyDown(vsOPS, KeyCode, Shift)
End Sub

Private Sub vsOPS_KeyPress(KeyAscii As Integer)
    Call OPSKeyPress(vsOPS, KeyAscii)
End Sub

Private Sub vsOPS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call OPSKeyPressEdit(vsOPS, Row, Col, KeyAscii)
End Sub

Private Sub vsOPS_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call OPSSetupEditWindow(vsOPS, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsOPS_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call OPSStartEdit(vsOPS, Row, Col, Cancel)
End Sub

Private Sub vsOPS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call OPSValidateEdit(vsOPS, Row, Col, Cancel)
End Sub

Private Sub vsRadioth_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call RadiothAfterEdit(vsRadioth, Row, Col)
End Sub

Private Sub vsRadioth_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call RadiothAfterRowColChange(vsRadioth, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsRadioth_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsRadioth)
End Sub

Private Sub vsRadioth_GotFocus()
    Call VSFlxGotFocus(vsRadioth)
End Sub

Private Sub vsRadioth_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RadiothKeyDown(vsRadioth, KeyCode, Shift)
End Sub

Private Sub vsRadioth_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Call RadiothKeyDownEdit(vsRadioth, Row, Col, KeyCode, Shift)
End Sub

Private Sub vsRadioth_KeyPress(KeyAscii As Integer)
    Call RadiothKeyPress(vsRadioth, KeyAscii)
End Sub

Private Sub vsRadioth_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call RadiothKeyPressEdit(vsRadioth, Row, Col, KeyAscii)
End Sub

Private Sub vsRadioth_LostFocus()
    Call RadiothLostFocus(vsRadioth)
End Sub

Private Sub vsRadioth_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call RadiothValidateEdit(vsRadioth, Row, Col, Cancel)
End Sub

Private Sub vsSpirit_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call SpiritAfterRowColChange(vsSpirit, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsSpirit_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsSpirit)
End Sub

Private Sub vsSpirit_GotFocus()
    Call VSFlxGotFocus(vsSpirit)
End Sub

Private Sub vsSpirit_KeyDown(KeyCode As Integer, Shift As Integer)
    Call SpiritKeyDown(vsSpirit, KeyCode, Shift)
End Sub

Private Sub vsSpirit_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Call SpiritKeyDownEdit(vsSpirit, Row, Col, KeyCode, Shift)
End Sub

Private Sub vsSpirit_KeyPress(KeyAscii As Integer)
    Call SpiritKeyPress(vsSpirit, KeyAscii)
End Sub

Private Sub vsSpirit_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call SpiritStartEdit(vsSpirit, Row, Col, Cancel)
End Sub

Private Sub vsSpirit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call SpiritValidateEdit(vsSpirit, Row, Col, Cancel)
End Sub

Private Sub vsTSJC_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call TSJCAfterEdit(vsTSJC, Row, Col)
End Sub

Private Sub vsTSJC_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call TSJCAfterRowColChange(vsTSJC, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsTSJC_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call TSJCCellButtonClick(vsTSJC, Row, Col)
End Sub

Private Sub vsTSJC_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange(vsTSJC)
End Sub

Private Sub vsTSJC_GotFocus()
    Call VSFlxGotFocus(vsTSJC)
End Sub

Private Sub vsTSJC_KeyDown(KeyCode As Integer, Shift As Integer)
    Call TSJCKeyDown(vsTSJC, KeyCode, Shift)
End Sub

Private Sub vsTSJC_KeyPress(KeyAscii As Integer)
    Call TSJCKeyPress(vsTSJC, KeyAscii)
End Sub

Private Sub vsTSJC_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call TSJCKeyPressEdit(vsTSJC, Row, Col, KeyAscii)
End Sub

Private Sub vsTSJC_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call TSJCSetupEditWindow(vsTSJC, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsTSJC_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call TSJCValidateEdit(vsTSJC, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DiagMouseDown(vsDiagXY, Button, Shift, X, Y)
End Sub

Private Sub vsDiagXY_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DiagMouseUp(vsDiagXY, Button, Shift, X, Y)
End Sub

Private Sub vsDiagZY_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DiagMouseDown(vsDiagZY, Button, Shift, X, Y)
End Sub

Private Sub vsDiagZY_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DiagMouseUp(vsDiagZY, Button, Shift, X, Y)
End Sub
