VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Object = "{D01C2596-4FE0-4EA9-9EE8-D97BE62A1165}#4.0#0"; "ZlPatiAddress.ocx"
Begin VB.Form frmInMedRecEdit_SC 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "住院首页"
   ClientHeight    =   55995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   55995
   ScaleWidth      =   16005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdTop 
      Appearance      =   0  'Flat
      Height          =   500
      Left            =   0
      Picture         =   "frmInMedRecEdit_SC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   412
      ToolTipText     =   "回顶部"
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
      TabIndex        =   413
      Top             =   1800
      Width           =   255
   End
   Begin VB.HScrollBar hsbMain 
      Height          =   255
      LargeChange     =   25
      Left            =   1000
      Max             =   100
      TabIndex        =   437
      Top             =   0
      Width           =   7935
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   55000
      Left            =   600
      ScaleHeight     =   54975
      ScaleWidth      =   12465
      TabIndex        =   414
      TabStop         =   0   'False
      Top             =   300
      Width           =   12500
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6840
         Index           =   1
         Left            =   732
         ScaleHeight     =   6840
         ScaleWidth      =   10995
         TabIndex        =   416
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
            TabIndex        =   97
            Top             =   4680
            Width           =   2240
         End
         Begin VB.TextBox txtSpecificInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   52
            Left            =   3120
            MaxLength       =   2
            TabIndex        =   31
            Top             =   840
            Width           =   180
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Height          =   240
            Index           =   42
            Left            =   8160
            TabIndex        =   441
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   5250
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
            Width           =   1680
            _ExtentX        =   2963
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
         Begin VB.CommandButton cmdDateInfo 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
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
            Picture         =   "frmInMedRecEdit_SC.frx":0359
            Style           =   1  'Graphical
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   5655
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.CommandButton cmdDateInfo 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
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
            Picture         =   "frmInMedRecEdit_SC.frx":044F
            Style           =   1  'Graphical
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   6450
            Visible         =   0   'False
            Width           =   270
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   225
            Index           =   3
            Left            =   1335
            TabIndex        =   121
            TabStop         =   0   'False
            Tag             =   "####-##-## ##:##"
            Top             =   6465
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
               Name            =   "宋体"
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
            TabIndex        =   104
            TabStop         =   0   'False
            Tag             =   "####-##-## ##:##"
            Top             =   5655
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
               Name            =   "宋体"
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
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Height          =   240
            Index           =   27
            Left            =   2960
            TabIndex        =   113
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   6045
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Height          =   240
            Index           =   29
            Left            =   8415
            TabIndex        =   119
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   6045
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Height          =   240
            Index           =   28
            Left            =   5690
            TabIndex        =   116
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   6045
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Height          =   240
            Index           =   5
            Left            =   8430
            TabIndex        =   91
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   3870
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Height          =   240
            Index           =   0
            Left            =   5850
            TabIndex        =   40
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   1470
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Height          =   240
            Index           =   6
            Left            =   5850
            TabIndex        =   72
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   3075
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Height          =   240
            Index           =   2
            Left            =   5850
            TabIndex        =   58
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   2270
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Height          =   240
            Index           =   3
            Left            =   5850
            TabIndex        =   66
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   2670
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Height          =   240
            Index           =   4
            Left            =   5850
            TabIndex        =   88
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   3870
            Width           =   270
         End
         Begin VB.CommandButton cmdAdressInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Height          =   240
            Index           =   1
            Left            =   8430
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   1470
            Width           =   270
         End
         Begin ZlPatiAddress.PatiAddress padrInfo 
            Height          =   225
            Index           =   6
            Left            =   1095
            TabIndex        =   70
            Top             =   3090
            Width           =   5025
            _ExtentX        =   8864
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
               Name            =   "宋体"
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
            Index           =   4
            Left            =   1095
            TabIndex        =   86
            Top             =   3885
            Width           =   5025
            _ExtentX        =   8864
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
               Name            =   "宋体"
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
               Name            =   "宋体"
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
            Index           =   0
            Left            =   1095
            TabIndex        =   38
            Top             =   1478
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   397
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   3470
            Visible         =   0   'False
            Width           =   1455
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               Index           =   41
               Left            =   0
               MaxLength       =   100
               TabIndex        =   82
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
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   29
            Left            =   6950
            MaxLength       =   100
            TabIndex        =   118
            Top             =   6060
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
            Width           =   1680
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   3
            Left            =   1335
            Locked          =   -1  'True
            TabIndex        =   122
            TabStop         =   0   'False
            Top             =   6465
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
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   5655
            Width           =   1815
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   6
            Left            =   1095
            MaxLength       =   100
            TabIndex        =   71
            ToolTipText     =   "按*键显示合约单位列表"
            Top             =   3075
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
            Top             =   2685
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
            ToolTipText     =   "按*键显示地区列表"
            Top             =   2280
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
            ToolTipText     =   "按*键显示地区列表"
            Top             =   1485
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
            Top             =   1485
            Width           =   5025
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   27
            Left            =   1335
            MaxLength       =   100
            TabIndex        =   112
            Top             =   6060
            Width           =   1815
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   28
            Left            =   4265
            MaxLength       =   100
            TabIndex        =   115
            Top             =   6060
            Width           =   1695
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   8
            Left            =   6950
            MaxLength       =   100
            TabIndex        =   110
            Top             =   5655
            Width           =   1395
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   7
            Left            =   4265
            Locked          =   -1  'True
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   5655
            Width           =   1695
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   7
            ItemData        =   "frmInMedRecEdit_SC.frx":0545
            Left            =   1340
            List            =   "frmInMedRecEdit_SC.frx":0547
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Top             =   5220
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
            Left            =   8100
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
            TabIndex        =   129
            TabStop         =   0   'False
            Top             =   6465
            Width           =   945
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   10
            Left            =   6950
            MaxLength       =   100
            TabIndex        =   127
            Top             =   6465
            Width           =   1395
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   9
            Left            =   4265
            Locked          =   -1  'True
            TabIndex        =   125
            TabStop         =   0   'False
            Top             =   6465
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
            TabIndex        =   84
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
            TabIndex        =   78
            Top             =   3478
            Width           =   1305
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   2
            Left            =   9690
            MaxLength       =   6
            TabIndex        =   76
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
            TabIndex        =   74
            Top             =   3078
            Width           =   1740
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   4
            Left            =   9690
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
            ItemData        =   "frmInMedRecEdit_SC.frx":0549
            Left            =   9690
            List            =   "frmInMedRecEdit_SC.frx":054B
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
            ItemData        =   "frmInMedRecEdit_SC.frx":054D
            Left            =   6960
            List            =   "frmInMedRecEdit_SC.frx":054F
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
            ItemData        =   "frmInMedRecEdit_SC.frx":0551
            Left            =   9690
            List            =   "frmInMedRecEdit_SC.frx":0553
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
            Width           =   600
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            ItemData        =   "frmInMedRecEdit_SC.frx":0555
            Left            =   3090
            List            =   "frmInMedRecEdit_SC.frx":0557
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
            Width           =   1260
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   5
            Left            =   6960
            MaxLength       =   30
            TabIndex        =   90
            ToolTipText     =   "按*键显示地区列表"
            Top             =   3878
            Width           =   1740
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   16
            Left            =   2610
            MaxLength       =   20
            TabIndex        =   30
            Top             =   983
            Width           =   360
         End
         Begin VB.ComboBox cboSpecificInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   16
            Left            =   3495
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   955
            Width           =   765
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   18
            Left            =   9210
            MaxLength       =   25
            TabIndex        =   36
            Top             =   983
            Width           =   1155
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "入院前经外院治疗"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   1
            Left            =   8775
            TabIndex        =   102
            Top             =   5250
            Width           =   1890
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   5
            Left            =   9690
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
            Left            =   5850
            MaxLength       =   25
            TabIndex        =   34
            Top             =   983
            Width           =   1155
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   61
            ItemData        =   "frmInMedRecEdit_SC.frx":0559
            Left            =   1095
            List            =   "frmInMedRecEdit_SC.frx":055B
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
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   5265
            Width           =   3885
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   4
            Left            =   1095
            MaxLength       =   100
            TabIndex        =   87
            Top             =   3885
            Width           =   5025
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   6
            ItemData        =   "frmInMedRecEdit_SC.frx":055D
            Left            =   2940
            List            =   "frmInMedRecEdit_SC.frx":055F
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   3440
            Width           =   1725
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "再入院"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   5760
            TabIndex        =   22
            Top             =   565
            Width           =   840
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   36
            Left            =   1080
            MaxLength       =   30
            TabIndex        =   93
            Top             =   4278
            Width           =   5025
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   47
            Left            =   6960
            MaxLength       =   12
            TabIndex        =   95
            Top             =   4278
            Width           =   1740
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "监护人身份证号"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   129
            Left            =   120
            TabIndex        =   96
            Top             =   4695
            Width           =   1260
         End
         Begin VB.Label lblSpecificInfo 
            AutoSize        =   -1  'True
            Caption         =   "30"
            Height          =   180
            Index           =   52
            Left            =   3120
            TabIndex        =   442
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
            Y1              =   5115
            Y2              =   5115
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "国籍"
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
            Caption         =   "新生儿出生体重              克"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   17
            Left            =   4530
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
            Caption         =   "入院途径"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   540
            TabIndex        =   98
            Top             =   5280
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "身高"
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
            Caption         =   "体重"
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
            Caption         =   "其他证件"
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
            Caption         =   "住院天数"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   19
            Left            =   8865
            TabIndex        =   128
            Top             =   6480
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "病房"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   10
            Left            =   6510
            TabIndex        =   126
            Top             =   6480
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "科室"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   9
            Left            =   3825
            TabIndex        =   124
            Top             =   6480
            Width           =   360
         End
         Begin VB.Label lblDateInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "出院时间"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   540
            TabIndex        =   120
            Top             =   6480
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "病房"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   8
            Left            =   6510
            TabIndex        =   109
            Top             =   5685
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "科室"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   3825
            TabIndex        =   107
            Top             =   5685
            Width           =   360
         End
         Begin VB.Label lblDateInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "入院时间"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   540
            TabIndex        =   103
            Top             =   5685
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "电话"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   14
            Left            =   6525
            TabIndex        =   83
            Top             =   3495
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "关系"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   6
            Left            =   2505
            TabIndex        =   79
            Top             =   3495
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "联系人姓名"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   77
            Top             =   3495
            Width           =   900
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "邮编"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   9255
            TabIndex        =   75
            Top             =   3105
            Width           =   360
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "电话"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   6525
            TabIndex        =   73
            Top             =   3105
            Width           =   360
         End
         Begin VB.Label lblAdressInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "工作单位"
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
            Caption         =   "邮编"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   9255
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
            Caption         =   "电话"
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
            Caption         =   "现住址"
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
            Caption         =   "身份证号"
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
            Caption         =   "出生地"
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
            Caption         =   "民族"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   5
            Left            =   9255
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
            Caption         =   "职业"
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
            Caption         =   "婚姻"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   9255
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
            Caption         =   "年龄"
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
            Caption         =   "出生日期"
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
            Caption         =   "性别"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   2655
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
            Caption         =   "姓名"
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
            Caption         =   "(年龄不足一周岁的)年龄"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   16
            Left            =   435
            TabIndex        =   29
            Top             =   1005
            Width           =   1980
         End
         Begin VB.Label lblAdressInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "籍贯"
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
            Caption         =   "户口地址"
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
            Caption         =   "邮编"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   5
            Left            =   9240
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
            Caption         =   "新生儿入院体重               克"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   18
            Left            =   7860
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
            Caption         =   "转科"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   27
            Left            =   540
            TabIndex        =   111
            Top             =   6075
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "→"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   28
            Left            =   3825
            TabIndex        =   114
            Top             =   6075
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "→"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   29
            Left            =   6510
            TabIndex        =   117
            Top             =   6075
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "转入"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   42
            Left            =   3825
            TabIndex        =   100
            Top             =   5280
            Width           =   360
         End
         Begin VB.Label lblAdressInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "联系人地址"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   85
            Top             =   3900
            Width           =   900
         End
         Begin VB.Label lblAdressInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "区域"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   5
            Left            =   6525
            TabIndex        =   89
            Top             =   3900
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   36
            Left            =   570
            TabIndex        =   92
            Top             =   4300
            Width           =   450
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "QQ"
            Height          =   180
            Index           =   47
            Left            =   6705
            TabIndex        =   94
            Top             =   4305
            Width           =   180
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
         TabIndex        =   415
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
            Left            =   9360
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
            Caption         =   "住院号"
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
            Caption         =   "第     次住院"
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
            Caption         =   "健康卡号"
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
            Caption         =   "医疗付费方式"
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
            Caption         =   "住 院 首 页"
            BeginProperty Font 
               Name            =   "宋体"
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
         TabIndex        =   134
         Top             =   0
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ListBox lstInfectParts 
            Appearance      =   0  'Flat
            Height          =   2340
            ItemData        =   "frmInMedRecEdit_SC.frx":0561
            Left            =   240
            List            =   "frmInMedRecEdit_SC.frx":0568
            Style           =   1  'Checkbox
            TabIndex        =   138
            Top             =   840
            Width           =   3615
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            Index           =   10
            ItemData        =   "frmInMedRecEdit_SC.frx":057C
            Left            =   1680
            List            =   "frmInMedRecEdit_SC.frx":057E
            Style           =   2  'Dropdown List
            TabIndex        =   136
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "感染部位"
            Height          =   180
            Index           =   128
            Left            =   120
            TabIndex        =   137
            Top             =   480
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "感染与死亡的关系"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   10
            Left            =   120
            TabIndex        =   135
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
         Height          =   7245
         Index           =   19
         Left            =   732
         ScaleHeight     =   7245
         ScaleWidth      =   10995
         TabIndex        =   434
         Tag             =   "true"
         Top             =   47370
         Width           =   11000
         Begin VB.PictureBox PicCareInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2295
            Left            =   0
            ScaleHeight     =   2295
            ScaleWidth      =   5205
            TabIndex        =   395
            Top             =   4920
            Width           =   5200
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   225
               Index           =   39
               Left            =   3480
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   402
               TabStop         =   0   'False
               Top             =   825
               Width           =   1600
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   225
               Index           =   38
               Left            =   1080
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   400
               TabStop         =   0   'False
               Top             =   825
               Width           =   1455
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "住院期间身体约束"
               ForeColor       =   &H80000008&
               Height          =   250
               Index           =   26
               Left            =   1080
               TabIndex        =   403
               Top             =   1245
               Width           =   2175
            End
            Begin VB.ComboBox cboBaseInfo 
               Appearance      =   0  'Flat
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   40
               ItemData        =   "frmInMedRecEdit_SC.frx":0580
               Left            =   1065
               List            =   "frmInMedRecEdit_SC.frx":0582
               Style           =   2  'Dropdown List
               TabIndex        =   398
               Top             =   360
               Width           =   1455
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               Index           =   40
               Left            =   3000
               MaxLength       =   100
               TabIndex        =   405
               Top             =   1673
               Width           =   2115
            End
            Begin VB.Line lineCareInfo 
               X1              =   1080
               X2              =   13080
               Y1              =   120
               Y2              =   120
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "临床表现"
               Height          =   210
               Index           =   39
               Left            =   2685
               TabIndex        =   401
               Top             =   825
               Width           =   705
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "引发药物"
               Height          =   210
               Index           =   38
               Left            =   285
               TabIndex        =   399
               Top             =   825
               Width           =   705
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "离院时透析（血透、腹透）尿素氮值"
               Height          =   210
               Index           =   40
               Left            =   0
               TabIndex        =   404
               Top             =   1680
               Width           =   2880
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输液反应"
               Height          =   180
               Index           =   40
               Left            =   270
               TabIndex        =   397
               Top             =   420
               Width           =   720
            End
            Begin VB.Label lblCareInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "护理相关情况"
               Height          =   180
               Left            =   0
               TabIndex        =   396
               Top             =   0
               Width           =   1080
            End
         End
         Begin VB.PictureBox PicAdvEvent 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   4095
            Left            =   5400
            ScaleHeight     =   4065
            ScaleWidth      =   5475
            TabIndex        =   385
            TabStop         =   0   'False
            Top             =   720
            Width           =   5500
            Begin VB.ListBox lstAdvEvent 
               Appearance      =   0  'Flat
               Height          =   2550
               ItemData        =   "frmInMedRecEdit_SC.frx":0584
               Left            =   -15
               List            =   "frmInMedRecEdit_SC.frx":058B
               Style           =   1  'Checkbox
               TabIndex        =   386
               Top             =   -15
               Width           =   5500
            End
            Begin VB.ComboBox cboBaseInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   46
               ItemData        =   "frmInMedRecEdit_SC.frx":059C
               Left            =   3795
               List            =   "frmInMedRecEdit_SC.frx":059E
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   390
               TabStop         =   0   'False
               Top             =   2715
               Width           =   1620
            End
            Begin VB.ComboBox cboBaseInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   48
               ItemData        =   "frmInMedRecEdit_SC.frx":05A0
               Left            =   1460
               List            =   "frmInMedRecEdit_SC.frx":05A2
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   394
               TabStop         =   0   'False
               Top             =   3615
               Width           =   3960
            End
            Begin VB.ComboBox cboBaseInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   47
               ItemData        =   "frmInMedRecEdit_SC.frx":05A4
               Left            =   1460
               List            =   "frmInMedRecEdit_SC.frx":05A6
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   392
               TabStop         =   0   'False
               Top             =   3165
               Width           =   3960
            End
            Begin VB.ComboBox cboBaseInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   45
               ItemData        =   "frmInMedRecEdit_SC.frx":05A8
               Left            =   1460
               List            =   "frmInMedRecEdit_SC.frx":05AA
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   388
               TabStop         =   0   'False
               Top             =   2715
               Width           =   1560
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "压疮发生期间"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   45
               Left            =   300
               TabIndex        =   387
               Top             =   2775
               Width           =   1080
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "跌倒或坠床原因"
               Height          =   180
               Index           =   48
               Left            =   120
               TabIndex        =   393
               Top             =   3675
               Width           =   1260
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "跌倒或坠床伤害"
               Height          =   180
               Index           =   47
               Left            =   120
               TabIndex        =   391
               Top             =   3225
               Width           =   1260
            End
            Begin VB.Label lblBaseInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "分期"
               Height          =   180
               Index           =   46
               Left            =   3360
               TabIndex        =   389
               Top             =   2775
               Width           =   360
            End
         End
         Begin VB.PictureBox PicPath 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1335
            Left            =   0
            ScaleHeight     =   1305
            ScaleWidth      =   5175
            TabIndex        =   371
            Top             =   720
            Width           =   5200
            Begin VB.Frame fraCbo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               ForeColor       =   &H80000008&
               Height          =   225
               Index           =   0
               Left            =   3120
               TabIndex        =   380
               Top             =   975
               Visible         =   0   'False
               Width           =   1965
               Begin VB.ComboBox cboBaseInfo 
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Index           =   62
                  ItemData        =   "frmInMedRecEdit_SC.frx":05AC
                  Left            =   -30
                  List            =   "frmInMedRecEdit_SC.frx":05AE
                  Style           =   2  'Dropdown List
                  TabIndex        =   381
                  Top             =   -25
                  Width           =   2035
               End
            End
            Begin VB.CommandButton cmdAutoLoad 
               Appearance      =   0  'Flat
               Caption         =   "自动提取"
               Height          =   330
               Index           =   3
               Left            =   1200
               TabIndex        =   373
               TabStop         =   0   'False
               Top             =   100
               Width           =   1000
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   225
               Index           =   31
               Left            =   3120
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   379
               TabStop         =   0   'False
               Top             =   968
               Width           =   1965
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   225
               Index           =   30
               Left            =   3120
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   376
               TabStop         =   0   'False
               Top             =   568
               Width           =   1965
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "变异"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   250
               Index           =   21
               Left            =   960
               TabIndex        =   377
               TabStop         =   0   'False
               Top             =   955
               Width           =   930
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "完成路径"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   250
               Index           =   20
               Left            =   960
               TabIndex        =   374
               TabStop         =   0   'False
               Top             =   550
               Width           =   1170
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "进入路径"
               ForeColor       =   &H80000008&
               Height          =   250
               Index           =   19
               Left            =   0
               TabIndex        =   372
               Top             =   150
               Width           =   1050
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "变异原因"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   31
               Left            =   2280
               TabIndex        =   378
               Top             =   990
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "退出原因"
               Height          =   180
               Index           =   30
               Left            =   2280
               TabIndex        =   375
               Top             =   585
               Width           =   720
            End
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "上一次住本院与本次住院是因同一疾病(主要诊断)"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   23
            Left            =   5400
            TabIndex        =   409
            Tag             =   "0"
            Top             =   5640
            Width           =   4545
         End
         Begin VB.CommandButton cmdLastDiag 
            Caption         =   "获取上次住院主要诊断"
            Height          =   330
            Left            =   5400
            TabIndex        =   410
            TabStop         =   0   'False
            Top             =   6045
            Width           =   4020
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   51
            Left            =   7200
            MaxLength       =   5
            TabIndex        =   408
            Top             =   5220
            Width           =   700
         End
         Begin VSFlex8Ctl.VSFlexGrid vsTSJC 
            Height          =   2415
            Left            =   0
            TabIndex        =   383
            Top             =   2400
            Width           =   5200
            _cx             =   9172
            _cy             =   4260
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            Rows            =   8
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_SC.frx":05B0
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
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "附页2"
            BeginProperty Font 
               Name            =   "宋体"
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
            Left            =   0
            TabIndex        =   369
            Top             =   0
            Width           =   570
         End
         Begin VB.Label lblTSJC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "检查情况"
            Height          =   180
            Left            =   0
            TabIndex        =   382
            Top             =   2115
            Width           =   720
         End
         Begin VB.Line lineCheck 
            X1              =   720
            X2              =   5160
            Y1              =   2205
            Y2              =   2205
         End
         Begin VB.Line lineAdvEvent 
            X1              =   6120
            X2              =   11510
            Y1              =   525
            Y2              =   525
         End
         Begin VB.Label lblAdvEvent 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "不良事件"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   5400
            TabIndex        =   384
            Top             =   435
            Width           =   720
         End
         Begin VB.Line lineH 
            Index           =   15
            X1              =   0
            X2              =   15000
            Y1              =   300
            Y2              =   300
         End
         Begin VB.Label lblPath 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "临床路径"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   370
            Top             =   480
            Width           =   720
         End
         Begin VB.Line linePath 
            X1              =   720
            X2              =   5160
            Y1              =   570
            Y2              =   570
         End
         Begin VB.Label lblDiagInfo 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "上次诊断"
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   5400
            TabIndex        =   411
            Top             =   6435
            Visible         =   0   'False
            Width           =   5460
         End
         Begin VB.Label lblSpecificInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "距上一次住本院的时间        天"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   51
            Left            =   5400
            TabIndex        =   407
            Top             =   5235
            Width           =   2700
         End
         Begin VB.Line lineIn 
            X1              =   7080
            X2              =   11880
            Y1              =   5040
            Y2              =   5040
         End
         Begin VB.Label lblIn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "本次住院与上次住院"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   5400
            TabIndex        =   406
            Top             =   4920
            Width           =   1620
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
         TabIndex        =   417
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   9165
         Width           =   11000
         Begin VB.OptionButton optDiag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "根据诊断标准输入"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   6600
            TabIndex        =   131
            TabStop         =   0   'False
            Top             =   55
            Value           =   -1  'True
            Width           =   1770
         End
         Begin VB.OptionButton optDiag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "根据疾病编码输入"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   1
            Left            =   8760
            TabIndex        =   132
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
            Picture         =   "frmInMedRecEdit_SC.frx":063D
            Style           =   1  'Graphical
            TabIndex        =   140
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
            Picture         =   "frmInMedRecEdit_SC.frx":2D68
            Style           =   1  'Graphical
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   1320
            Width           =   375
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
            Height          =   3000
            Left            =   0
            TabIndex        =   133
            Top             =   360
            Width           =   10500
            _cx             =   18521
            _cy             =   5292
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            FormatString    =   $"frmInMedRecEdit_SC.frx":5310
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
            Caption         =   "西医诊断"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   130
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
         TabIndex        =   425
         Tag             =   "false"
         Top             =   30075
         Visible         =   0   'False
         Width           =   11000
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "住院费用"
            BeginProperty Font 
               Name            =   "宋体"
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
         TabIndex        =   419
         TabStop         =   0   'False
         Top             =   16935
         Visible         =   0   'False
         Width           =   11000
         Begin VB.CommandButton cmdDiagMove 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Height          =   375
            Index           =   3
            Left            =   10620
            Picture         =   "frmInMedRecEdit_SC.frx":566A
            Style           =   1  'Graphical
            TabIndex        =   192
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
            Picture         =   "frmInMedRecEdit_SC.frx":7D95
            Style           =   1  'Graphical
            TabIndex        =   191
            TabStop         =   0   'False
            Top             =   855
            Width           =   375
         End
         Begin VB.OptionButton optDiag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "根据疾病编码输入"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   8760
            TabIndex        =   189
            TabStop         =   0   'False
            Top             =   150
            Width           =   1770
         End
         Begin VB.OptionButton optDiag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "根据诊断标准输入"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   6720
            TabIndex        =   188
            TabStop         =   0   'False
            Top             =   150
            Value           =   -1  'True
            Width           =   1890
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
            Height          =   2100
            Left            =   0
            TabIndex        =   190
            Top             =   400
            Width           =   10500
            _cx             =   18521
            _cy             =   3704
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            FormatString    =   $"frmInMedRecEdit_SC.frx":A33D
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
            Caption         =   "中医诊断"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   187
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
         Height          =   4215
         Index           =   3
         Left            =   732
         ScaleHeight     =   4215
         ScaleWidth      =   10995
         TabIndex        =   418
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   12720
         Width           =   11000
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   20
            ItemData        =   "frmInMedRecEdit_SC.frx":A631
            Left            =   4380
            List            =   "frmInMedRecEdit_SC.frx":A633
            Style           =   2  'Dropdown List
            TabIndex        =   439
            Top             =   3180
            Width           =   1470
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   225
            Index           =   5
            Left            =   1200
            TabIndex        =   178
            TabStop         =   0   'False
            Tag             =   "####-##-## ##:##:##"
            Top             =   3218
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
               Name            =   "宋体"
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
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   225
            Index           =   4
            Left            =   8550
            TabIndex        =   147
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
               Name            =   "宋体"
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
            Left            =   1200
            MaxLength       =   5
            TabIndex        =   171
            Top             =   2718
            Width           =   600
         End
         Begin VB.CommandButton cmdDateInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   10380
            Picture         =   "frmInMedRecEdit_SC.frx":A635
            Style           =   1  'Graphical
            TabIndex        =   149
            TabStop         =   0   'False
            Top             =   310
            Width           =   285
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   240
            Index           =   21
            Left            =   10410
            TabIndex        =   163
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   1710
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   240
            Index           =   22
            Left            =   10410
            TabIndex        =   176
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   2710
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   240
            Index           =   20
            Left            =   10410
            TabIndex        =   182
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   3210
            Width           =   270
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   60
            ItemData        =   "frmInMedRecEdit_SC.frx":A72B
            Left            =   1550
            List            =   "frmInMedRecEdit_SC.frx":A72D
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   184
            TabStop         =   0   'False
            Top             =   3580
            Width           =   675
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   16
            ItemData        =   "frmInMedRecEdit_SC.frx":A72F
            Left            =   1800
            List            =   "frmInMedRecEdit_SC.frx":A731
            Style           =   2  'Dropdown List
            TabIndex        =   155
            Top             =   1180
            Width           =   1470
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "是否确诊"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   2
            Left            =   5475
            TabIndex        =   145
            Top             =   305
            Width           =   1170
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   143
            Top             =   280
            Width           =   1470
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   19
            Left            =   1200
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   2218
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   18
            ItemData        =   "frmInMedRecEdit_SC.frx":A733
            Left            =   9210
            List            =   "frmInMedRecEdit_SC.frx":A735
            Style           =   2  'Dropdown List
            TabIndex        =   169
            Top             =   2180
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   21
            ItemData        =   "frmInMedRecEdit_SC.frx":A737
            Left            =   4380
            List            =   "frmInMedRecEdit_SC.frx":A739
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   186
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
            Left            =   5460
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   175
            TabStop         =   0   'False
            Top             =   2718
            Width           =   5220
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   22
            Left            =   3540
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   173
            TabStop         =   0   'False
            Top             =   2718
            Width           =   600
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "新发肿瘤"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   5
            Left            =   3720
            TabIndex        =   144
            Top             =   305
            Width           =   1050
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "医院感染作病原学检查"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   3
            Left            =   825
            TabIndex        =   160
            TabStop         =   0   'False
            Top             =   1705
            Width           =   2130
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   21
            Left            =   5460
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   162
            TabStop         =   0   'False
            Top             =   1718
            Width           =   5220
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   13
            ItemData        =   "frmInMedRecEdit_SC.frx":A73B
            Left            =   5460
            List            =   "frmInMedRecEdit_SC.frx":A73D
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   153
            TabStop         =   0   'False
            Top             =   780
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   12
            ItemData        =   "frmInMedRecEdit_SC.frx":A73F
            Left            =   1800
            List            =   "frmInMedRecEdit_SC.frx":A741
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   151
            TabStop         =   0   'False
            Top             =   780
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   19
            ItemData        =   "frmInMedRecEdit_SC.frx":A743
            Left            =   5460
            List            =   "frmInMedRecEdit_SC.frx":A745
            Style           =   2  'Dropdown List
            TabIndex        =   167
            Top             =   2180
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   15
            ItemData        =   "frmInMedRecEdit_SC.frx":A747
            Left            =   9210
            List            =   "frmInMedRecEdit_SC.frx":A749
            Style           =   2  'Dropdown List
            TabIndex        =   159
            Top             =   1180
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   14
            ItemData        =   "frmInMedRecEdit_SC.frx":A74B
            Left            =   5460
            List            =   "frmInMedRecEdit_SC.frx":A74D
            Style           =   2  'Dropdown List
            TabIndex        =   157
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
            Left            =   7140
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   181
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
            Left            =   8550
            Locked          =   -1  'True
            TabIndex        =   148
            TabStop         =   0   'False
            Top             =   315
            Width           =   1830
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   5
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   179
            TabStop         =   0   'False
            Top             =   3218
            Width           =   1830
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "死亡期间"
            Height          =   180
            Index           =   20
            Left            =   3585
            TabIndex        =   440
            Top             =   3240
            Width           =   720
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "西医诊断符合情况"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   141
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
            Caption         =   "死亡患者尸检"
            Height          =   180
            Index           =   60
            Left            =   405
            TabIndex        =   183
            Top             =   3645
            Width           =   1080
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊与入院"
            Height          =   180
            Index           =   16
            Left            =   825
            TabIndex        =   154
            Top             =   1245
            Width           =   900
         End
         Begin VB.Label lblDateInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "主要诊断确认日期"
            Height          =   180
            Index           =   4
            Left            =   7035
            TabIndex        =   146
            Top             =   345
            Width           =   1440
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "入院情况"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   11
            Left            =   1005
            TabIndex        =   142
            Top             =   345
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病理号"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   19
            Left            =   585
            TabIndex        =   164
            Top             =   2235
            Width           =   540
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "抢救病因"
            Height          =   180
            Index           =   22
            Left            =   4665
            TabIndex        =   174
            Top             =   2745
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "成功次数"
            Height          =   180
            Index           =   22
            Left            =   2745
            TabIndex        =   172
            Top             =   2745
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "抢救次数"
            Height          =   180
            Index           =   21
            Left            =   405
            TabIndex        =   170
            Top             =   2745
            Width           =   720
         End
         Begin VB.Label lblDateInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "死亡时间"
            Height          =   180
            Index           =   5
            Left            =   405
            TabIndex        =   177
            Top             =   3240
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "死亡原因"
            Height          =   180
            Index           =   20
            Left            =   6360
            TabIndex        =   180
            Top             =   3240
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医院感染病原学诊断"
            Height          =   180
            Index           =   21
            Left            =   3765
            TabIndex        =   161
            Top             =   1740
            Width           =   1620
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "最高诊断依据"
            Height          =   180
            Index           =   13
            Left            =   4305
            TabIndex        =   152
            Top             =   840
            Width           =   1080
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "分化程度"
            Height          =   180
            Index           =   12
            Left            =   1005
            TabIndex        =   150
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "临床与尸检"
            Height          =   180
            Index           =   21
            Left            =   3405
            TabIndex        =   185
            Top             =   3645
            Width           =   900
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "临床与病理"
            Height          =   180
            Index           =   19
            Left            =   4485
            TabIndex        =   166
            Top             =   2235
            Width           =   900
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "放射与病理"
            Height          =   180
            Index           =   18
            Left            =   8235
            TabIndex        =   168
            Top             =   2235
            Width           =   900
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "入院与出院"
            Height          =   180
            Index           =   15
            Left            =   8235
            TabIndex        =   158
            Top             =   1245
            Width           =   900
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊与出院"
            Height          =   180
            Index           =   14
            Left            =   4485
            TabIndex        =   156
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
         Height          =   2025
         Index           =   9
         Left            =   732
         ScaleHeight     =   2025
         ScaleWidth      =   10995
         TabIndex        =   424
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   28050
         Width           =   11000
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   17
            ItemData        =   "frmInMedRecEdit_SC.frx":A74F
            Left            =   1320
            List            =   "frmInMedRecEdit_SC.frx":A751
            Style           =   2  'Dropdown List
            TabIndex        =   285
            Top             =   290
            Width           =   1470
         End
         Begin VB.CommandButton cmdAutoLoad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "自动提取"
            Height          =   330
            Index           =   1
            Left            =   9600
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
            Picture         =   "frmInMedRecEdit_SC.frx":A753
            Style           =   1  'Graphical
            TabIndex        =   292
            TabStop         =   0   'False
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdOPSMove 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Height          =   375
            Index           =   1
            Left            =   10620
            Picture         =   "frmInMedRecEdit_SC.frx":CCFB
            Style           =   1  'Graphical
            TabIndex        =   293
            TabStop         =   0   'False
            Top             =   1560
            Width           =   375
         End
         Begin VB.CheckBox chkParaOPSInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "未找到时允许自由录入"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   6720
            TabIndex        =   288
            Top             =   150
            Width           =   2115
         End
         Begin VB.OptionButton OptParaOPSInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "根据ICD9-CM3输入"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   1
            Left            =   4950
            MaskColor       =   &H80000005&
            TabIndex        =   287
            TabStop         =   0   'False
            Top             =   150
            Width           =   1770
         End
         Begin VB.OptionButton OptParaOPSInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "根据诊疗项目输入"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   3120
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
               Name            =   "宋体"
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
            FormatString    =   $"frmInMedRecEdit_SC.frx":F426
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
               Picture         =   "frmInMedRecEdit_SC.frx":FAD4
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   435
               TabStop         =   0   'False
               Top             =   300
               Visible         =   0   'False
               Width           =   255
            End
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "术前与术后"
            Height          =   180
            Index           =   17
            Left            =   390
            TabIndex        =   284
            Top             =   350
            Width           =   900
         End
         Begin VB.Label lblAutoInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "提示信息"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   3120
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
            Caption         =   "手术记录"
            BeginProperty Font 
               Name            =   "宋体"
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
         TabIndex        =   432
         Tag             =   "true"
         Top             =   43440
         Width           =   11000
         Begin VSFlex8Ctl.VSFlexGrid vsfMain 
            Height          =   1200
            Left            =   0
            TabIndex        =   363
            Top             =   345
            Width           =   10990
            _cx             =   19385
            _cy             =   2117
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            FormatString    =   $"frmInMedRecEdit_SC.frx":107C6
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
            Caption         =   "病案附加项目"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   362
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
         Height          =   3600
         Index           =   16
         Left            =   732
         ScaleHeight     =   3600
         ScaleWidth      =   10995
         TabIndex        =   431
         Tag             =   "true"
         Top             =   39840
         Width           =   11000
         Begin VSFlex8Ctl.VSFlexGrid vsFlxAddICU 
            Height          =   1320
            Left            =   0
            TabIndex        =   359
            Top             =   585
            Width           =   10990
            _cx             =   19385
            _cy             =   2328
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            Rows            =   3
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_SC.frx":108D6
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
         Begin VSFlex8Ctl.VSFlexGrid vsICUInstruments 
            Height          =   1155
            Left            =   0
            TabIndex        =   361
            Top             =   2250
            Width           =   10990
            _cx             =   19385
            _cy             =   2037
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            Rows            =   3
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_SC.frx":109BA
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
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "重症监护情况"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   357
            Top             =   0
            Width           =   1350
         End
         Begin VB.Label lblICUInstruments 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "患者入住重症监护室期间器械使用情况"
            Height          =   180
            Left            =   0
            TabIndex        =   360
            Top             =   2040
            Width           =   3060
         End
         Begin VB.Label lblAddICU 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "患者入住重症监护病房记录"
            Height          =   180
            Left            =   0
            TabIndex        =   358
            Top             =   360
            Width           =   2415
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
         TabIndex        =   428
         Tag             =   "true"
         Top             =   35775
         Width           =   11000
         Begin VSFlex8Ctl.VSFlexGrid vsRadioth 
            Height          =   1200
            Left            =   0
            TabIndex        =   352
            Top             =   450
            Width           =   10990
            _cx             =   19385
            _cy             =   2117
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            FormatString    =   $"frmInMedRecEdit_SC.frx":10AE2
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
            Caption         =   "提示信息"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   1
            Left            =   3000
            TabIndex        =   351
            Top             =   180
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "放疗记录信息"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   350
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
         TabIndex        =   427
         Tag             =   "true"
         Top             =   33930
         Width           =   11000
         Begin VSFlex8Ctl.VSFlexGrid vsChemoth 
            Height          =   1200
            Left            =   0
            TabIndex        =   349
            Top             =   450
            Width           =   10990
            _cx             =   19385
            _cy             =   2117
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            FormatString    =   $"frmInMedRecEdit_SC.frx":10C2A
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
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "化疗记录信息"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   347
            Top             =   0
            Width           =   1350
         End
         Begin VB.Label lblEdit 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "提示信息"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   0
            Left            =   2880
            TabIndex        =   348
            Top             =   200
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
         Height          =   3600
         Index           =   11
         Left            =   732
         ScaleHeight     =   3600
         ScaleWidth      =   10995
         TabIndex        =   426
         Tag             =   "true"
         Top             =   30330
         Width           =   11000
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "住院期间告病重或病危"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   22
            Left            =   3480
            TabIndex        =   342
            Top             =   3185
            Width           =   2130
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   225
            Index           =   6
            Left            =   9195
            TabIndex        =   304
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
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   225
            Index           =   7
            Left            =   10275
            TabIndex        =   305
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
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   225
            Index           =   24
            Left            =   10605
            TabIndex        =   319
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   1298
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "…"
            Height          =   240
            Index           =   17
            Left            =   5970
            TabIndex        =   312
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   790
            Width           =   270
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   6
            Left            =   9195
            MaxLength       =   30
            TabIndex        =   306
            Top             =   405
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txtDateInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   7
            Left            =   10275
            MaxLength       =   30
            TabIndex        =   307
            Top             =   405
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "科研病案"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   10
            Left            =   6105
            TabIndex        =   301
            Top             =   385
            Width           =   1050
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   34
            ItemData        =   "frmInMedRecEdit_SC.frx":10D78
            Left            =   915
            List            =   "frmInMedRecEdit_SC.frx":10D7A
            Style           =   2  'Dropdown List
            TabIndex        =   297
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "示教病案"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   9
            Left            =   5010
            TabIndex        =   300
            Top             =   385
            Width           =   1050
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   18
            Left            =   7995
            MaxLength       =   100
            TabIndex        =   314
            Top             =   798
            Width           =   2880
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   17
            Left            =   3795
            MaxLength       =   100
            TabIndex        =   311
            Top             =   798
            Width           =   2175
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            Index           =   43
            Left            =   915
            Style           =   2  'Dropdown List
            TabIndex        =   309
            Top             =   760
            Width           =   1815
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "疑难病例"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   11
            Left            =   7170
            TabIndex        =   302
            Top             =   385
            Width           =   1050
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "随诊"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   12
            Left            =   7170
            TabIndex        =   343
            Top             =   3185
            Width           =   690
         End
         Begin VB.ComboBox cboSpecificInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   41
            ItemData        =   "frmInMedRecEdit_SC.frx":10D7C
            Left            =   9780
            List            =   "frmInMedRecEdit_SC.frx":10D7E
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   346
            TabStop         =   0   'False
            Top             =   3170
            Width           =   735
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   41
            Left            =   8970
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   345
            TabStop         =   0   'False
            Top             =   3198
            Width           =   765
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   44
            ItemData        =   "frmInMedRecEdit_SC.frx":10D80
            Left            =   915
            List            =   "frmInMedRecEdit_SC.frx":10D82
            Style           =   2  'Dropdown List
            TabIndex        =   316
            Top             =   1260
            Width           =   1815
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   24
            Left            =   5520
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   318
            TabStop         =   0   'False
            Top             =   1298
            Width           =   5355
         End
         Begin VB.OptionButton optInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "无"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   3330
            TabIndex        =   328
            Top             =   2285
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "有，目的："
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   1
            Left            =   4260
            TabIndex        =   329
            Top             =   2285
            Width           =   1200
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   49
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   327
            Top             =   2260
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
            TabIndex        =   333
            Top             =   2798
            Width           =   675
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   38
            Left            =   7800
            MaxLength       =   6
            TabIndex        =   337
            Top             =   2798
            Width           =   675
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   34
            Left            =   1275
            MaxLength       =   6
            TabIndex        =   341
            Top             =   3198
            Width           =   1035
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   40
            Left            =   9900
            MaxLength       =   6
            TabIndex        =   339
            Top             =   2798
            Width           =   675
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   39
            Left            =   8760
            MaxLength       =   6
            TabIndex        =   338
            Top             =   2798
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
            TabIndex        =   335
            Top             =   2798
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
            TabIndex        =   334
            Top             =   2798
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   25
            Left            =   5520
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   330
            Top             =   2298
            Width           =   5355
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   56
            Left            =   3580
            Style           =   2  'Dropdown List
            TabIndex        =   299
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   37
            Left            =   5520
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   325
            TabStop         =   0   'False
            Top             =   1798
            Width           =   5355
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   50
            Left            =   4275
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   324
            TabStop         =   0   'False
            Top             =   1798
            Width           =   400
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            IMEMode         =   3  'DISABLE
            Index           =   49
            Left            =   2520
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   322
            TabStop         =   0   'False
            Top             =   1798
            Width           =   400
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "会诊情况"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   29
            Left            =   480
            TabIndex        =   320
            Top             =   1780
            Width           =   1155
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "住院情况"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "医学警示"
            Height          =   180
            Index           =   17
            Left            =   3000
            TabIndex        =   310
            Top             =   825
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病例分型"
            Height          =   180
            Index           =   34
            Left            =   120
            TabIndex        =   296
            Top             =   420
            Width           =   720
         End
         Begin VB.Label lblDateInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发病时间"
            Height          =   180
            Index           =   6
            Left            =   8400
            TabIndex        =   303
            Top             =   420
            Width           =   720
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "其他医学警示"
            Height          =   180
            Index           =   18
            Left            =   6840
            TabIndex        =   313
            Top             =   825
            Width           =   1080
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "生育状况"
            Height          =   180
            Index           =   43
            Left            =   120
            TabIndex        =   308
            Top             =   825
            Width           =   720
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "随诊期限"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   41
            Left            =   8160
            TabIndex        =   344
            Top             =   3225
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "离院方式"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   44
            Left            =   120
            TabIndex        =   315
            Top             =   1320
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "转入"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   24
            Left            =   5100
            TabIndex        =   317
            Top             =   1320
            Width           =   360
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "  入院后        天         小时        分钟"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   38
            Left            =   7065
            TabIndex        =   336
            Top             =   2820
            Width           =   3870
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "呼吸机使用            小时"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   34
            Left            =   360
            TabIndex        =   340
            Top             =   3225
            Width           =   2340
         End
         Begin VB.Label lblNumInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "入院前        天         小时        分钟"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   2565
            TabIndex        =   332
            Top             =   2820
            Width           =   3690
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "是否有"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   49
            Left            =   420
            TabIndex        =   326
            Top             =   2325
            Width           =   540
         End
         Begin VB.Label lblNumInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "颅脑损伤患者昏迷时间:"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   35
            Left            =   390
            TabIndex        =   331
            Top             =   2820
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
            Index           =   13
            X1              =   0
            X2              =   14000
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "TNM分期"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   56
            Left            =   2880
            TabIndex        =   298
            Top             =   420
            Width           =   630
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "外院会诊      次，其他"
            Height          =   180
            Index           =   50
            Left            =   3480
            TabIndex        =   323
            Top             =   1815
            Width           =   1980
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "院内会诊      次"
            Height          =   180
            Index           =   49
            Left            =   1740
            TabIndex        =   321
            Top             =   1815
            Width           =   1440
         End
         Begin VB.Line lineH 
            Index           =   12
            X1              =   0
            X2              =   14000
            Y1              =   1660
            Y2              =   1660
         End
         Begin VB.Line lineH 
            Index           =   14
            X1              =   0
            X2              =   14000
            Y1              =   2660
            Y2              =   2660
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
         TabIndex        =   423
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   25965
         Width           =   11000
         Begin VB.CommandButton cmdDateInfo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
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
            Picture         =   "frmInMedRecEdit_SC.frx":10D84
            Style           =   1  'Graphical
            TabIndex        =   282
            TabStop         =   0   'False
            Top             =   1672
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
               Name            =   "宋体"
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
            Top             =   1680
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   8
            ItemData        =   "frmInMedRecEdit_SC.frx":10E7A
            Left            =   3705
            List            =   "frmInMedRecEdit_SC.frx":10E7C
            TabIndex        =   276
            Top             =   1640
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   9
            ItemData        =   "frmInMedRecEdit_SC.frx":10E7E
            Left            =   6465
            List            =   "frmInMedRecEdit_SC.frx":10E80
            TabIndex        =   278
            Top             =   1640
            Width           =   1185
         End
         Begin VB.CommandButton cmdSign 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "签名"
            Enabled         =   0   'False
            Height          =   300
            Index           =   3
            Left            =   10440
            TabIndex        =   262
            TabStop         =   0   'False
            Top             =   333
            Width           =   520
         End
         Begin VB.CommandButton cmdSign 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "签名"
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   4920
            TabIndex        =   256
            TabStop         =   0   'False
            Top             =   333
            Width           =   520
         End
         Begin VB.CommandButton cmdSign 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "签名"
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   7680
            TabIndex        =   259
            TabStop         =   0   'False
            Top             =   333
            Width           =   520
         End
         Begin VB.CommandButton cmdSign 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "签名"
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   2040
            TabIndex        =   253
            TabStop         =   0   'False
            Top             =   333
            Width           =   520
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   7
            ItemData        =   "frmInMedRecEdit_SC.frx":10E82
            Left            =   6465
            List            =   "frmInMedRecEdit_SC.frx":10E84
            TabIndex        =   268
            Top             =   740
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   3
            ItemData        =   "frmInMedRecEdit_SC.frx":10E86
            Left            =   3705
            List            =   "frmInMedRecEdit_SC.frx":10E88
            TabIndex        =   266
            Top             =   740
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   5
            ItemData        =   "frmInMedRecEdit_SC.frx":10E8A
            Left            =   9225
            List            =   "frmInMedRecEdit_SC.frx":10E8C
            TabIndex        =   261
            Top             =   340
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   4
            ItemData        =   "frmInMedRecEdit_SC.frx":10E8E
            Left            =   6465
            List            =   "frmInMedRecEdit_SC.frx":10E90
            TabIndex        =   258
            Top             =   340
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            ItemData        =   "frmInMedRecEdit_SC.frx":10E92
            Left            =   3705
            List            =   "frmInMedRecEdit_SC.frx":10E94
            TabIndex        =   255
            Top             =   340
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            ItemData        =   "frmInMedRecEdit_SC.frx":10E96
            Left            =   825
            List            =   "frmInMedRecEdit_SC.frx":10E98
            TabIndex        =   252
            Top             =   340
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   10
            ItemData        =   "frmInMedRecEdit_SC.frx":10E9A
            Left            =   825
            List            =   "frmInMedRecEdit_SC.frx":10E9C
            TabIndex        =   264
            Top             =   740
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            ItemData        =   "frmInMedRecEdit_SC.frx":10E9E
            Left            =   825
            List            =   "frmInMedRecEdit_SC.frx":10EA0
            TabIndex        =   272
            Top             =   1140
            Width           =   1185
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   33
            ItemData        =   "frmInMedRecEdit_SC.frx":10EA2
            Left            =   825
            List            =   "frmInMedRecEdit_SC.frx":10EA4
            Style           =   2  'Dropdown List
            TabIndex        =   274
            Top             =   1640
            Width           =   1185
         End
         Begin VB.ComboBox cboManInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   12
            ItemData        =   "frmInMedRecEdit_SC.frx":10EA6
            Left            =   9225
            List            =   "frmInMedRecEdit_SC.frx":10EA8
            TabIndex        =   270
            Top             =   735
            Width           =   1185
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "签名信息"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "主任(副主任)    医师"
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
            Caption         =   "质控医师"
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
            Caption         =   "质控护士"
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
            Caption         =   "实习医师"
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
            Caption         =   "进修医师"
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
            Caption         =   "住院医师"
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
            Caption         =   "主治医师"
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
            Caption         =   "科主任"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   251
            Top             =   405
            Width           =   540
         End
         Begin VB.Label lblManInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "责任护士"
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
            Caption         =   "门诊医师"
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
            Caption         =   "病案质量"
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
            Caption         =   "质控日期"
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
         Begin VB.Label lblManInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "主诊医师"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   12
            Left            =   8430
            TabIndex        =   269
            Top             =   795
            Width           =   720
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1865
         Index           =   7
         Left            =   732
         ScaleHeight     =   1860
         ScaleWidth      =   10995
         TabIndex        =   422
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   24105
         Width           =   11000
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   38
            ItemData        =   "frmInMedRecEdit_SC.frx":10EAA
            Left            =   5580
            List            =   "frmInMedRecEdit_SC.frx":10EAC
            Style           =   2  'Dropdown List
            TabIndex        =   237
            Top             =   225
            Width           =   1200
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   36
            ItemData        =   "frmInMedRecEdit_SC.frx":10EAE
            Left            =   1920
            List            =   "frmInMedRecEdit_SC.frx":10EB0
            Style           =   2  'Dropdown List
            TabIndex        =   229
            Top             =   225
            Width           =   1425
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   42
            Left            =   9360
            Style           =   2  'Dropdown List
            TabIndex        =   247
            Top             =   1025
            Width           =   1410
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   41
            ItemData        =   "frmInMedRecEdit_SC.frx":10EB2
            Left            =   9360
            List            =   "frmInMedRecEdit_SC.frx":10EB4
            Style           =   2  'Dropdown List
            TabIndex        =   245
            Top             =   625
            Width           =   1425
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   23
            Left            =   9360
            MaxLength       =   30
            TabIndex        =   249
            Top             =   1463
            Width           =   1380
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   33
            Left            =   5580
            MaxLength       =   5
            TabIndex        =   243
            Top             =   1463
            Width           =   1140
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   32
            Left            =   5580
            MaxLength       =   5
            TabIndex        =   241
            Top             =   1063
            Width           =   1140
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   31
            Left            =   5580
            MaxLength       =   5
            TabIndex        =   239
            Top             =   663
            Width           =   1140
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   30
            Left            =   1920
            MaxLength       =   5
            TabIndex        =   235
            Top             =   1463
            Width           =   1140
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   29
            Left            =   1920
            MaxLength       =   5
            TabIndex        =   231
            Top             =   663
            Width           =   1140
         End
         Begin VB.TextBox txtSpecificInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   48
            Left            =   1920
            MaxLength       =   5
            TabIndex        =   233
            Top             =   1063
            Width           =   1140
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "输血信息"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   227
            Top             =   0
            Width           =   900
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "血型"
            Height          =   180
            Index           =   36
            Left            =   1485
            TabIndex        =   228
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
            TabIndex        =   236
            Top             =   285
            Width           =   180
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输血前的9项检查"
            Height          =   180
            Index           =   42
            Left            =   7935
            TabIndex        =   246
            Top             =   1080
            Width           =   1350
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输血反应"
            Height          =   180
            Index           =   41
            Left            =   8565
            TabIndex        =   244
            Top             =   690
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输其他"
            Height          =   180
            Index           =   23
            Left            =   8745
            TabIndex        =   248
            Top             =   1485
            Width           =   540
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输血浆               ml"
            Height          =   180
            Index           =   31
            Left            =   4980
            TabIndex        =   238
            Top             =   690
            Width           =   2070
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "输红细胞              单位"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   29
            Left            =   1110
            TabIndex        =   230
            Top             =   690
            Width           =   2340
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输血小板              单位"
            Height          =   180
            Index           =   30
            Left            =   1110
            TabIndex        =   234
            Top             =   1485
            Width           =   2340
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输全血               ml"
            Height          =   180
            Index           =   32
            Left            =   4980
            TabIndex        =   240
            Top             =   1080
            Width           =   2070
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "自体回收               ml"
            Height          =   180
            Index           =   33
            Left            =   4800
            TabIndex        =   242
            Top             =   1485
            Width           =   2250
         End
         Begin VB.Label lblSpecificInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输白蛋白              g"
            Height          =   180
            Index           =   48
            Left            =   1110
            TabIndex        =   232
            Top             =   1080
            Width           =   2070
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
         TabIndex        =   421
         TabStop         =   0   'False
         Tag             =   "true"
         Top             =   22260
         Width           =   11000
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "无过敏记录"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   30
            Left            =   1320
            TabIndex        =   223
            Top             =   150
            Width           =   1215
         End
         Begin VB.CommandButton cmdAutoLoad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "自动提取"
            Height          =   330
            Index           =   2
            Left            =   9860
            TabIndex        =   438
            TabStop         =   0   'False
            Top             =   60
            Width           =   1000
         End
         Begin VB.OptionButton optAller 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "根据过敏源输入"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   1
            Left            =   7710
            TabIndex        =   225
            TabStop         =   0   'False
            Top             =   150
            Width           =   1800
         End
         Begin VB.OptionButton optAller 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "根据药品目录输入"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   5760
            TabIndex        =   224
            TabStop         =   0   'False
            Top             =   150
            Width           =   1800
         End
         Begin VSFlex8Ctl.VSFlexGrid vsAller 
            Height          =   1200
            Left            =   0
            TabIndex        =   226
            Top             =   450
            Width           =   10990
            _cx             =   19385
            _cy             =   2117
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            FormatString    =   $"frmInMedRecEdit_SC.frx":10EB6
            ScrollTrack     =   -1  'True
            ScrollBars      =   1
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
            Caption         =   "药物过敏"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   222
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
         TabIndex        =   420
         TabStop         =   0   'False
         Top             =   19635
         Visible         =   0   'False
         Width           =   11000
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   23
            ItemData        =   "frmInMedRecEdit_SC.frx":10F68
            Left            =   6540
            List            =   "frmInMedRecEdit_SC.frx":10F6A
            Style           =   2  'Dropdown List
            TabIndex        =   197
            Top             =   280
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   22
            ItemData        =   "frmInMedRecEdit_SC.frx":10F6C
            Left            =   2880
            List            =   "frmInMedRecEdit_SC.frx":10F6E
            Style           =   2  'Dropdown List
            TabIndex        =   195
            Top             =   280
            Width           =   1635
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "危重"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   6
            Left            =   2475
            TabIndex        =   199
            Top             =   805
            Width           =   800
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "疑难"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   8
            Left            =   6870
            TabIndex        =   201
            Top             =   805
            Width           =   800
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "急症"
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   7
            Left            =   4680
            TabIndex        =   200
            Top             =   805
            Width           =   800
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   28
            ItemData        =   "frmInMedRecEdit_SC.frx":10F70
            Left            =   2160
            List            =   "frmInMedRecEdit_SC.frx":10F72
            Style           =   2  'Dropdown List
            TabIndex        =   217
            Top             =   2175
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   30
            ItemData        =   "frmInMedRecEdit_SC.frx":10F74
            Left            =   5820
            List            =   "frmInMedRecEdit_SC.frx":10F76
            Style           =   2  'Dropdown List
            TabIndex        =   219
            Top             =   2175
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   32
            ItemData        =   "frmInMedRecEdit_SC.frx":10F78
            Left            =   9210
            List            =   "frmInMedRecEdit_SC.frx":10F7A
            Style           =   2  'Dropdown List
            TabIndex        =   221
            Top             =   2175
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   24
            ItemData        =   "frmInMedRecEdit_SC.frx":10F7C
            Left            =   2160
            List            =   "frmInMedRecEdit_SC.frx":10F7E
            Style           =   2  'Dropdown List
            TabIndex        =   204
            Top             =   1275
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   25
            ItemData        =   "frmInMedRecEdit_SC.frx":10F80
            Left            =   5820
            List            =   "frmInMedRecEdit_SC.frx":10F82
            Style           =   2  'Dropdown List
            TabIndex        =   206
            Top             =   1275
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   26
            ItemData        =   "frmInMedRecEdit_SC.frx":10F84
            Left            =   9210
            List            =   "frmInMedRecEdit_SC.frx":10F86
            Style           =   2  'Dropdown List
            TabIndex        =   208
            Top             =   1275
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   31
            ItemData        =   "frmInMedRecEdit_SC.frx":10F88
            Left            =   9210
            List            =   "frmInMedRecEdit_SC.frx":10F8A
            Style           =   2  'Dropdown List
            TabIndex        =   215
            Top             =   1785
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   29
            ItemData        =   "frmInMedRecEdit_SC.frx":10F8C
            Left            =   5820
            List            =   "frmInMedRecEdit_SC.frx":10F8E
            Style           =   2  'Dropdown List
            TabIndex        =   213
            Top             =   1785
            Width           =   1635
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   27
            ItemData        =   "frmInMedRecEdit_SC.frx":10F90
            Left            =   2160
            List            =   "frmInMedRecEdit_SC.frx":10F92
            Style           =   2  'Dropdown List
            TabIndex        =   211
            Top             =   1785
            Width           =   1635
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "中医诊断符合情况"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   193
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
            Caption         =   "入院与出院"
            Height          =   180
            Index           =   23
            Left            =   5565
            TabIndex        =   196
            Top             =   345
            Width           =   900
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "门诊与出院"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   22
            Left            =   1905
            TabIndex        =   194
            Top             =   345
            Width           =   900
         End
         Begin VB.Label lblZY 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "住院期间病情:"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   198
            Top             =   840
            Width           =   1170
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "使用中医诊疗设备"
            Height          =   180
            Index           =   28
            Left            =   645
            TabIndex        =   216
            Top             =   2235
            Width           =   1440
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "使用中医诊疗技术"
            Height          =   180
            Index           =   30
            Left            =   4305
            TabIndex        =   218
            Top             =   2235
            Width           =   1440
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "辨证施护"
            Height          =   180
            Index           =   32
            Left            =   8415
            TabIndex        =   220
            Top             =   2235
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "辨证"
            Height          =   180
            Index           =   24
            Left            =   1725
            TabIndex        =   203
            Top             =   1335
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "治法"
            Height          =   180
            Index           =   25
            Left            =   5385
            TabIndex        =   205
            Top             =   1335
            Width           =   360
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "方药"
            Height          =   180
            Index           =   26
            Left            =   8775
            TabIndex        =   207
            Top             =   1335
            Width           =   360
         End
         Begin VB.Label lblZY 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "准确度:"
            Height          =   180
            Index           =   1
            Left            =   540
            TabIndex        =   202
            Top             =   1335
            Width           =   630
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "自制中药制剂"
            Height          =   180
            Index           =   31
            Left            =   8055
            TabIndex        =   214
            Top             =   1845
            Width           =   1080
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "抢救方法"
            Height          =   180
            Index           =   29
            Left            =   5025
            TabIndex        =   212
            Top             =   1845
            Width           =   720
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "治疗类别"
            Height          =   180
            Index           =   27
            Left            =   1365
            TabIndex        =   210
            Top             =   1845
            Width           =   720
         End
         Begin VB.Label lblZY 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "治疗方法:"
            Height          =   180
            Index           =   2
            Left            =   360
            TabIndex        =   209
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
         Height          =   315
         Index           =   15
         Left            =   732
         ScaleHeight     =   315
         ScaleWidth      =   10995
         TabIndex        =   430
         Tag             =   "true"
         Top             =   39525
         Width           =   11000
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "抗精神病治疗情况"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   356
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
         TabIndex        =   429
         Tag             =   "true"
         Top             =   37635
         Width           =   11000
         Begin VB.CommandButton cmdAutoLoad 
            Appearance      =   0  'Flat
            Caption         =   "自动提取"
            Height          =   330
            Index           =   0
            Left            =   9840
            TabIndex        =   354
            TabStop         =   0   'False
            Top             =   120
            Width           =   1000
         End
         Begin VSFlex8Ctl.VSFlexGrid vsKSS 
            Height          =   1200
            Left            =   0
            TabIndex        =   355
            Top             =   500
            Width           =   10990
            _cx             =   19385
            _cy             =   2117
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_SC.frx":10F94
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
            Caption         =   "抗菌药物使用情况（按DDD数降序排列）"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   353
            Top             =   0
            Width           =   3960
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2180
         Index           =   18
         Left            =   732
         ScaleHeight     =   2175
         ScaleWidth      =   10995
         TabIndex        =   433
         Tag             =   "true"
         Top             =   45195
         Width           =   11000
         Begin VSFlex8Ctl.VSFlexGrid vsSample 
            Height          =   1380
            Left            =   5520
            TabIndex        =   368
            Top             =   600
            Width           =   5400
            _cx             =   9525
            _cy             =   2434
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            Rows            =   2
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_SC.frx":110D8
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
         Begin VSFlex8Ctl.VSFlexGrid vsInfect 
            Height          =   1380
            Left            =   0
            TabIndex        =   366
            Top             =   600
            Width           =   5400
            _cx             =   9525
            _cy             =   2434
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            Rows            =   2
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_SC.frx":11198
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
            Caption         =   "附页1"
            BeginProperty Font 
               Name            =   "宋体"
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
            TabIndex        =   364
            Top             =   0
            Width           =   570
         End
         Begin VB.Label lblInfect 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医院感染情况"
            Height          =   180
            Index           =   0
            Left            =   50
            TabIndex        =   365
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lblSample 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "标本来源"
            Height          =   180
            Left            =   5640
            TabIndex        =   367
            Top             =   360
            Width           =   720
         End
      End
      Begin MSComCtl2.MonthView monInfo 
         Height          =   2160
         Left            =   960
         TabIndex        =   436
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
         StartOfWeek     =   188547073
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
   Begin VB.Image imgButtonDel 
      Height          =   240
      Left            =   0
      Picture         =   "frmInMedRecEdit_SC.frx":11230
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgButtonNew 
      Height          =   240
      Left            =   500
      Picture         =   "frmInMedRecEdit_SC.frx":17A82
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmInMedRecEdit_SC"
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

Private Sub lstInfectParts_LostFocus()
    Call lstLostFocus(lstInfectParts)
End Sub

Private Sub padrInfo_SetInput(Index As Integer, ByVal intLevel As Integer, rsReturn As ADODB.Recordset)
    Call SetYoubian(Index, intLevel, rsReturn)
End Sub

Private Sub PicPage_Resize(Index As Integer)
    Call PicPageResize(Index)
End Sub

Private Sub padrInfo_Change(Index As Integer)
    Call CheckValueChange
End Sub

Private Sub txtAdressInfo_Change(Index As Integer)
    Call txtAdressInfoChange(Index)
End Sub

Private Sub txtAdressInfo_GotFocus(Index As Integer)
    Call txtAdressInfoGotFocus(Index)
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

Private Sub cmdLastDiag_Click()
    Call cmdLastDiagClick
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
    Call CheckValueChange
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

Private Sub vsFlxAddICU_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call FlxAddICUStartEdit(vsFlxAddICU, Row, Col, Cancel)
End Sub

Private Sub vsFlxAddICU_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call FlxAddICUValidateEdit(vsFlxAddICU, Row, Col, Cancel)
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

Private Sub vsICUInstruments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsICUInstrumentsAfterEdit(vsICUInstruments, Row, Col)
End Sub

Private Sub vsICUInstruments_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsICUInstrumentsAfterRowColChange(vsICUInstruments, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsICUInstruments_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange
End Sub

Private Sub vsICUInstruments_KeyDown(KeyCode As Integer, Shift As Integer)
    Call vsICUInstrumentsKeyDown(vsICUInstruments, KeyCode, Shift)
End Sub

Private Sub vsICUInstruments_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Call vsICUInstrumentsKeyDownEdit(vsICUInstruments, Row, Col, KeyCode, Shift)
End Sub

Private Sub vsICUInstruments_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call vsICUInstrumentsStartEdit(vsICUInstruments, Row, Col, Cancel)
End Sub

Private Sub vsICUInstruments_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call vsICUInstrumentsValidateEdit(vsICUInstruments, Row, Col, Cancel)
End Sub

Private Sub vsInfect_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsInfectAfterEdit(vsInfect, Row, Col)
End Sub

Private Sub vsInfect_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsInfectAfterRowColChange(vsInfect, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsInfect_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call vsInfectCellButtonClick(vsInfect, Row, Col)
End Sub

Private Sub vsInfect_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange
End Sub

Private Sub vsInfect_GotFocus()
    Call VSFlxGotFocus(vsInfect)
End Sub

Private Sub vsInfect_KeyDown(KeyCode As Integer, Shift As Integer)
    Call vsInfectKeyDown(vsInfect, KeyCode, Shift)
End Sub

Private Sub vsInfect_KeyPress(KeyAscii As Integer)
    Call vsInfectKeyPress(vsInfect, KeyAscii)
End Sub

Private Sub vsInfect_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call vsInfectKeyPressEdit(vsInfect, Row, Col, KeyAscii)
End Sub

Private Sub vsInfect_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call vsInfectStartEdit(vsInfect, Row, Col, Cancel)
End Sub

Private Sub vsInfect_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call vsInfectValidateEdit(vsInfect, Row, Col, Cancel)
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

Private Sub vsSample_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsSampleAfterEdit(vsSample, Row, Col)
End Sub

Private Sub vsSample_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange
End Sub

Private Sub vsSample_GotFocus()
    Call VSFlxGotFocus(vsSample)
End Sub

Private Sub vsSample_KeyDown(KeyCode As Integer, Shift As Integer)
    Call vsSampleKeyDown(vsSample, KeyCode, Shift)
End Sub

Private Sub vsSample_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call vsSampleStartEdit(vsSample, Row, Col, Cancel)
End Sub

Private Sub vsSample_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call vsSampleValidateEdit(vsSample, Row, Col, Cancel)
End Sub

Private Sub vsTSJC_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Call CheckValueChange
End Sub

Private Sub vsTSJC_GotFocus()
    Call VSFlxGotFocus(vsTSJC)
End Sub

Private Sub vsTSJC_KeyPress(KeyAscii As Integer)
    Call TSJCKeyPress(vsTSJC, KeyAscii)
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
