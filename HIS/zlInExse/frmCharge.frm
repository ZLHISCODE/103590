VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Object = "*\A..\ZlBillEdit\zl9BillEdit.vbp"
Begin VB.Form frmCharge 
   AutoRedraw      =   -1  'True
   Caption         =   "סԺ���ʴ���"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13290
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   13290
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picStatuPancl 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   9750
      ScaleHeight     =   300
      ScaleWidth      =   2340
      TabIndex        =   62
      Top             =   7605
      Width           =   2340
      Begin VB.Label lblStatuPati 
         Caption         =   "����Ƿ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   0
         TabIndex        =   63
         Top             =   45
         Width           =   855
      End
   End
   Begin VB.Frame fraTitle 
      Height          =   1095
      Left            =   30
      TabIndex        =   33
      ToolTipText     =   "���:F6"
      Top             =   -120
      Width           =   13065
      Begin VB.CommandButton cmdSelWholeSet 
         Caption         =   "����(&T)"
         Height          =   375
         Left            =   3405
         TabIndex        =   65
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   630
         Width           =   1080
      End
      Begin VB.CommandButton cmdSaveWholeSet 
         Caption         =   "����Ϊ�����շ���Ŀ(&W)"
         Height          =   375
         Left            =   4530
         TabIndex        =   64
         Top             =   630
         Width           =   2715
      End
      Begin VB.Timer tmrStatuPati 
         Interval        =   100
         Left            =   1000
         Top             =   1005
      End
      Begin VB.CommandButton cmd�䷽ 
         Caption         =   "�䷽(&R)"
         Height          =   375
         Left            =   2280
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F11"
         Top             =   630
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox txtIn 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   690
         MaxLength       =   8
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   630
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CheckBox chkIn 
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
         Height          =   375
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "������ʵ�:F3"
         Top             =   630
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9870
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   630
         Width           =   1425
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
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   11325
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F8"
         Top             =   630
         Width           =   495
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   18000
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   30
         X2              =   18000
         Y1              =   570
         Y2              =   570
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
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   11325
         TabIndex        =   43
         Top             =   645
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "סԺ���ʵ�"
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
         Left            =   225
         TabIndex        =   37
         ToolTipText     =   "���:F6"
         Top             =   180
         Width           =   1875
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "���ݺ�"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   9090
         TabIndex        =   34
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame fraUnit 
      Height          =   1065
      Left            =   9555
      TabIndex        =   32
      Top             =   855
      Width           =   2325
      Begin VB.ComboBox cbo�������� 
         Height          =   360
         Left            =   90
         TabIndex        =   10
         Text            =   "cbo��������"
         Top             =   600
         Width           =   2160
      End
      Begin VB.Label lbl�������� 
         Caption         =   "��������"
         Height          =   240
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1065
      Left            =   30
      TabIndex        =   31
      Top             =   855
      Width           =   9525
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   690
         TabIndex        =   66
         Top             =   210
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   635
         Appearance      =   2
         IDKindStr       =   $"frmCharge.frx":08CA
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
         DefaultCardType =   "0"
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   210
         Width           =   1380
      End
      Begin VB.TextBox txt������ 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   615
         Width           =   1380
      End
      Begin VB.TextBox txt������ 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6280
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   615
         Width           =   915
      End
      Begin VB.ComboBox cboҽ�Ƹ��� 
         Height          =   360
         Left            =   3705
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   615
         Width           =   1695
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6280
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   915
      End
      Begin VB.TextBox txtPatient 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1330
         MaxLength       =   100
         TabIndex        =   1
         Top             =   210
         Width           =   1270
      End
      Begin VB.ComboBox cboSex 
         Height          =   360
         Left            =   3240
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   210
         Width           =   975
      End
      Begin VB.TextBox txtOld 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   630
      End
      Begin VB.ComboBox cbo�ѱ� 
         Height          =   360
         Left            =   675
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   615
         Width           =   1950
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   240
         Left            =   7275
         TabIndex        =   56
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Left            =   7275
         TabIndex        =   55
         Top             =   675
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Left            =   5510
         TabIndex        =   54
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lblҽ�Ƹ��� 
         Caption         =   "���ʽ"
         Height          =   240
         Index           =   0
         Left            =   2715
         TabIndex        =   53
         Top             =   675
         Width           =   960
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   5750
         TabIndex        =   45
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         Caption         =   "����"
         ForeColor       =   &H80000007&
         Height          =   240
         Index           =   7
         Left            =   165
         TabIndex        =   41
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   240
         Index           =   8
         Left            =   2715
         TabIndex        =   40
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   240
         Index           =   9
         Left            =   4260
         TabIndex        =   39
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�ѱ�"
         Height          =   240
         Index           =   12
         Left            =   165
         TabIndex        =   38
         Top             =   675
         Width           =   480
      End
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   2580
      Left            =   -15
      TabIndex        =   11
      Top             =   2520
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   4551
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      TxtCheck        =   -1  'True
      TxtCheck        =   -1  'True
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Active          =   -1  'True
      Cols            =   2
      RowHeight0      =   360
      RowHeightMin    =   360
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin MSCommLib.MSComm com 
      Left            =   120
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox picAppend 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   0
      ScaleHeight     =   2745
      ScaleWidth      =   13290
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   5190
      Width           =   13290
      Begin VB.ComboBox cboTemp 
         Height          =   360
         Left            =   7320
         TabIndex        =   67
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   4000
         Width           =   1575
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   11415
         Top             =   1020
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   1
         ImageHeight     =   18
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCharge.frx":0967
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȡ��(&C)"
         Height          =   420
         Left            =   10095
         TabIndex        =   28
         ToolTipText     =   "�ȼ�:Esc"
         Top             =   1575
         Width           =   1575
      End
      Begin VB.Frame fraAppend 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   47
         ToolTipText     =   "���:F6"
         Top             =   -105
         Width           =   13065
         Begin VB.ComboBox cboBaby 
            Height          =   360
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   180
            Width           =   1800
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�������"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   4320
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CheckBox chk�Ӱ� 
            Caption         =   "�Ӱ�(&A)"
            Height          =   270
            Left            =   120
            TabIndex        =   12
            Top             =   225
            Width           =   1170
         End
         Begin VB.ComboBox cbo������ 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   6495
            TabIndex        =   16
            Top             =   180
            Width           =   2205
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   10635
            TabIndex        =   17
            Top             =   180
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   635
            _Version        =   393216
            AutoTab         =   -1  'True
            HideSelection   =   0   'False
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
         Begin VB.Label lblBaby 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ӥ����(&B)"
            Height          =   240
            Left            =   1320
            TabIndex        =   13
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl������ 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   240
            Left            =   5730
            TabIndex        =   49
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "ʱ��"
            Height          =   240
            Left            =   10095
            TabIndex        =   48
            Top             =   240
            Width           =   480
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
         Height          =   1635
         Left            =   0
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   2884
         _Version        =   393216
         Rows            =   5
         FixedCols       =   0
         RowHeightMin    =   320
         BackColorBkg    =   15466495
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         FormatString    =   "^         ��Ŀ|^          ���"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmdPrice 
         BackColor       =   &H00C0C0C0&
         Caption         =   "���۵�(&I)"
         Height          =   420
         Left            =   6915
         TabIndex        =   26
         ToolTipText     =   "����Ϊ���۵�"
         Top             =   1575
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame fraDrawDept 
         Height          =   720
         Left            =   0
         TabIndex        =   58
         Top             =   345
         Width           =   13575
         Begin VB.ComboBox cboִ������ 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   9375
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   247
            Width           =   1725
         End
         Begin VB.TextBox txt���˱�ע 
            BackColor       =   &H00E0E0E0&
            Height          =   360
            Left            =   5445
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   225
            Width           =   2700
         End
         Begin VB.ComboBox cboDrawDept 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   247
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label lblִ������ 
            AutoSize        =   -1  'True
            Caption         =   "ִ������"
            Height          =   240
            Left            =   8325
            TabIndex        =   20
            Top             =   307
            Width           =   960
         End
         Begin VB.Label lbl���˱�ע 
            Caption         =   "���˱�ע"
            Height          =   225
            Left            =   4455
            TabIndex        =   22
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label lblDrawDrugDept 
            AutoSize        =   -1  'True
            Caption         =   "��ҩ����"
            Height          =   255
            Left            =   255
            TabIndex        =   18
            Top             =   300
            Visible         =   0   'False
            Width           =   960
         End
      End
      Begin VB.Frame fraStat 
         Height          =   1770
         Left            =   3510
         TabIndex        =   50
         Top             =   975
         Width           =   3240
         Begin VB.TextBox txtPreNO 
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
            ForeColor       =   &H00C00000&
            Height          =   450
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   1252
            Width           =   1845
         End
         Begin VB.TextBox txtʵ�� 
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
            ForeColor       =   &H00C00000&
            Height          =   450
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   750
            Width           =   1845
         End
         Begin VB.TextBox txtӦ�� 
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
            Height          =   450
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   250
            Width           =   1845
         End
         Begin VB.Label lblPreNO 
            AutoSize        =   -1  'True
            Caption         =   "����"
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
            Left            =   225
            TabIndex        =   60
            Top             =   1320
            Width           =   690
         End
         Begin VB.Label lblʵ�� 
            AutoSize        =   -1  'True
            Caption         =   "ʵ��"
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
            Left            =   225
            TabIndex        =   59
            Top             =   818
            Width           =   690
         End
         Begin VB.Label lblӦ�� 
            AutoSize        =   -1  'True
            Caption         =   "Ӧ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   225
            TabIndex        =   51
            Top             =   318
            Width           =   690
         End
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȷ��(&O)"
         Height          =   420
         Left            =   8505
         TabIndex        =   27
         ToolTipText     =   "�ȼ���F2"
         Top             =   1575
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   35
      Top             =   7935
      Width           =   13290
      _ExtentX        =   23442
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmCharge.frx":0A59
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13970
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   88
            Key             =   "�������"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   2
            Key             =   "MedicareType"
            Object.ToolTipText     =   "���մ���"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   951
            MinWidth        =   951
            Picture         =   "frmCharge.frx":12ED
            Key             =   "Drugstore"
            Object.Tag             =   "Drugstore"
            Object.ToolTipText     =   "ҩ������"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   952
            MinWidth        =   952
            Picture         =   "frmCharge.frx":1607
            Key             =   "BarCode"
            Object.Tag             =   "BarCode"
            Object.ToolTipText     =   "��ʾ�������"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCharge.frx":1D31
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCharge.frx":236B
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraBarCode 
      Height          =   630
      Left            =   30
      TabIndex        =   68
      Top             =   1815
      Width           =   11850
      Begin VB.TextBox txtBarCode 
         Height          =   360
         Left            =   705
         TabIndex        =   69
         Top             =   195
         Width           =   11040
      End
      Begin VB.Label lblBarCode 
         Caption         =   "����"
         Height          =   300
         Left            =   150
         TabIndex        =   70
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ϼ�:"
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
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   945
   End
End
Attribute VB_Name = "frmCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'����������������������������������������������������������������������������������������������������������������������������������������
'��ڲ�����
'2.����ʼ״̬������
Public mbytInState As Byte '0-ִ��,1-����,2-����,3-����
Public mblnCopyBill As Boolean '�Ƿ��Զ����Ʋ����ĵ���
Public mstrInNO As String '�������ĵ��ݺ�
Public mbytNOType As Byte '��������:2-�˹����ʵ�,3-�Զ����ʵ�   �����ʹ���ʹ��
Public mblnNOMoved As Boolean '�����ĵ����Ƿ��ں����ݱ���
Public mlngҽ��ID As Long 'Nvl(���ID,ID),��ʿվֱ������ʱ,���ڴ���ȱʡ��ѡ���ʵ���
Public mstrTime As String '�����������ݵĵǼ�ʱ��
Public mblnDelete As Boolean '�Ƿ�����˷ѵ���
Public mlngDelRow As Long '���ⲿ��������ʱ��ȱʡ���ʵķ��ü�¼
Public mlngUnitID As Long '��ǰ���ʲ���,Ϊ0ʱ��ʾ���в���
Public mlngDeptID As Long '��ǰ���ʿ���,Ϊ0ʱ��ʾ���п���
Public mbytUseType As Byte '���ʵ���;,0-��ͨ����,1-�����ҷ�ɢ����,2-ҽ�����Ҽ���
Public mlng����ID As Long '���ҷ�ɢ������
Public mlng��ҳID As Long '����ʹ��
Public mstrPrivs As String
Public mlng����ҽ�� As Long

Public mlngModule As Long
Public mbln���� As Boolean '33744
Public mstr���ת��ʱ�� As String


'����������������������������������������������������������������������������������������������������������������������������������������
Private mstrPrivsOpt As String '���ʲ���1150ģ�����Ȩ����
Private mstr������Ŀ As String '������Ŀ
'���ݶ���
Private mrsClass As ADODB.Recordset '���ݲ�����ȡ�ĵ�ǰ���õ��շ����
Private mrsUnit As ADODB.Recordset '��ѡ���ִ�п���
Private mrsInfo As New ADODB.Recordset '������Ϣ
Private mrsMedAudit As ADODB.Recordset  '�����������ķ�����Ŀ
Private mrsWork As New ADODB.Recordset '�����ϰ��ҩ��
Private mrsWarn As ADODB.Recordset  '����������
Private mrsMedPayMode As ADODB.Recordset '���п��õ�ҽ�Ƹ��ʽ
Private mrs�������� As ADODB.Recordset '��������
Private mrs�������� As ADODB.Recordset  '��ѡ�Ŀ�������
Private mrs������ As ADODB.Recordset    '��ѡҽ���ͻ�ʿ
Private mrs��ҩ���� As ADODB.Recordset
Private mobjBaseItem As Object    '������Ŀ���ò���
'�������
Private mobjBill As ExpenseBill '������õ��ݶ������
Private mcolBillDetails As BillDetails '���ݵ��շ�ϸĿ��
Private mobjBillDetail As BillDetail '���ݵ��շ�ϸĿ����
Private mcolBillInComes As BillInComes '�շ�ϸĿ��������Ŀ��
Private mobjBillIncome As BillInCome '�շ�ϸĿ��������Ŀ����
Private mobjDetail As Detail '�������շ�ϸĿ����
Private mcolDetails As Details '�������շ�ϸĿ����
Private mcolMoneys As BillInComes  '���������Ŀ���ܼ���(��ʾ����ӡʱʹ��)���

Private Enum BillColType       '���ݿؼ���������
    CheckBox = -1
    Text_UnModify = 0
    CommandButton = 1
    Date = 2
    ComboBox = 3
    Text = 4
    UnFocus = 5
End Enum

Private Enum BillCol
    �� = 0
    ��� = 1
    ��Ŀ = 2
    ��Ʒ�� = 3
    ��� = 4
    ��λ = 5
    ���� = 6
    ���� = 7
    ���� = 8
    Ӧ�ս�� = 9
    ʵ�ս�� = 10
    ִ�п��� = 11
    ��־ = 12
    ���� = 13
End Enum

'�������
Public mlngBill����ID As Long   '����ʱʹ��
Public mlngBill��ҳID As Long

Private mblncboEnterCell As Boolean '����ѭ������
Private mblncboClick  As Boolean    '����ѭ������
Private mlngPreRow As Long '��ǰ�к�,�����иı�ʱ�ж�

Private mbln����ְ���� As Boolean     '�Ƿ���д���ְ����
Private mbln����������� As Boolean     '�Ƿ���д����������

Private mblnWork As Boolean '��ǰ�Ƿ��������ϰ��ҩ��
Private mlngҩƷ���ID As Long '��ǰ���ݲ�����ҩƷ������ID
Private mlng�������ID As Long '��ǰ���ݲ���������������ID

Private mcurModiMoney As Currency '�޸ĵ���ʱԭ���ݵĽ��
Private mstrUnitIDs As String   '��ǰ����Ա�����в���ID
Private mstrWarn As String '�Ѿ���������ѡ����������
Private mblnSavePrice As Boolean    'Ƿ��ʱ����Ϊ���۵�
Private mblnSendMateria As Boolean  '���ʺ��Զ���ҩ
Private mcolStock1 As Collection '��Ÿ���ҩƷ�ⷿ�ĳ����鷽ʽ
Private mcolStock2 As Collection '��Ÿ������Ŀ�ĳ����鷽ʽ
Private mblnSetControl As Boolean

Private mbln������۸� As Boolean     '���޸ĺ͵��뵥��ʱ,���÷ѱ�ʱ������۸�,����ʱ����,����Ҳ������

Private mblnDrop As Boolean '��KeyDown���ж�cbo�����˵�ǰ�Ƿ񵯳�
Private mblnFirst As Boolean
Private mblnValid As Boolean
Private mblnNewRow As Boolean
Private mblnPrint As Boolean '��ȡ��˵�ʱ�Ƿ����Ҫ��ӡ���շ����
Private mblnOne As Boolean '�Ƿ�ֻ��һ�������շ����
Private marrColData() As Integer '��ǰ���ݱ༭����ӳ��
Private mdblItemNum As Double '���ݿ��е�ǰ�����Ŀ������
Private mblnSelect As Boolean '���ڿ����շ�ϸĿ�����Ƿ��������б�ѡ���ѡ����
Private mblnNotClick As Boolean
Private WithEvents mobjBrushCheck As clsBrushCardInput
Attribute mobjBrushCheck.VB_VarHelpID = -1
Private mobjCard As New Card
Private mbln����ˢ�� As Boolean
Private mlng���� As Long

Private Const STR_HEAD = "��,450,4;���,750,1;��Ŀ,2175,1;��Ʒ��,1800,1;���,1105,1;��λ,520,4;����,520,1;����,570,1;����,1055,7;" & "Ӧ�ս��,1030,7;ʵ�ս��,1080,7;ִ�п���,1255,1;��־,520,4;����,520,1"

'��ǰ���������ҽ��֧�ֲ���
Private Type TYPE_MedicarePAR
    �������� As Boolean
    �����ϴ� As Boolean
    ������ɺ��ϴ� As Boolean
    ���������ϴ� As Boolean
    ʵʱ��� As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private Enum Pan
    C2��ʾ��Ϣ = 2
End Enum
'-----------------------------------------------------------------------------------
'���㿨���
Private mstrPassWord As String 'ˢ��ʱ������
'-----------------------------------------------------------------------------------
Private mstrҩƷ�۸�ȼ� As String, mstr���ļ۸�ȼ� As String, mstr��ͨ�۸�ȼ� As String
Private mblnShowBarCode As Boolean '��ʾ���������

Private Sub Bill_cboKeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    '���˺� ����:27378 ����:2010-01-27 13:35:37
    If Bill.cboStyle = DropOlnyDown Then Exit Sub
    
    Select Case Bill.TextMatrix(0, Bill.Col)
        Case "ִ�п���"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case "��ҩҩ��"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case Else
        Exit Sub
    End Select
    lngRow = Bill.Row
    If mobjBill.Details.Count < lngRow Then Exit Sub
    
    With mobjBill.Details(lngRow)
        If InStr(",4,5,6,7,", .�շ����) > 0 Then
            If mrsWork Is Nothing Then Exit Sub
            If mrsWork.State <> 1 Then Exit Sub
            If zlSelectDept(Me, mlngModule, Bill.cboObj, mrsWork, Bill.CboText, True, , False) = False Then Exit Sub
        Else
            If mrsUnit Is Nothing Then Exit Sub
            If mrsUnit.State <> 1 Then Exit Sub
            If zlSelectDept(Me, mlngModule, Bill.cboObj, mrsUnit, Bill.CboText, True, , False) = False Then Exit Sub
        End If
    End With
    Exit Sub
End Sub

Private Sub cboDrawDept_Click()
    Dim lng��ҩ����ID As Long
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If cboDrawDept.ListIndex <> -1 Then lng��ҩ����ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
    If mobjBill.��ҩ����ID = lng��ҩ����ID Then Exit Sub
    mobjBill.��ҩ����ID = lng��ҩ����ID
End Sub

Private Sub cboDrawDept_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 And Not cboDrawDept.Locked Then
        lngIdx = zlControl.CboMatchIndex(cboDrawDept.hWnd, KeyAscii)
        If lngIdx = -1 And cboDrawDept.ListCount > 0 Then lngIdx = 0
        cboDrawDept.ListIndex = lngIdx
    ElseIf KeyAscii = 13 Then
        If cboDrawDept.ListIndex = -1 Then
            Beep
        Else
            mobjBill.��ҩ����ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cbo��������_GotFocus()
    zlControl.TxtSelAll cbo��������
End Sub

Private Sub cbo��������_LostFocus()
    cbo��������.SelLength = 0
End Sub

Private Sub cbo��������_Validate(Cancel As Boolean)
    If cbo��������.Text <> "" And cbo��������.ListIndex < 0 Then cbo��������.Text = ""
End Sub

Private Sub cboҽ�Ƹ���_Click()
    On Error GoTo errHandler
    If mbytInState <> 0 Then Exit Sub
    If cboҽ�Ƹ���.ListIndex = -1 Then Exit Sub
    If gintPriceGradeStartType < 2 Then Exit Sub
    
    If mrsInfo.State = adStateOpen Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID)), zlStr.NeedName(cboҽ�Ƹ���.Text), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    Else
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, zlStr.NeedName(cboҽ�Ƹ���.Text), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    End If
    
    If mbln������۸� Then Exit Sub
    If mobjBill.Details.Count = 0 Then Exit Sub
    
    '���¼���۸�
    Call CalcMoneys
    Call ShowDetails
    Call ShowMoney
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboִ������_Click()
    Dim i As Long
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If mobjBill Is Nothing Then Exit Sub
    
    mobjBill.ִ������ = cboִ������.ItemData(cboִ������.ListIndex)
    

End Sub

Private Sub cboִ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub



Private Sub cmdSaveWholeSet_Click()
    Dim i As Long, strItems As String, lngִ�п���ID As Long
    Dim rsTemp As ADODB.Recordset, dbl���� As Double, dbl�۸� As Double
    Dim strSQL As String
    
    '����Ϊ�����շ���Ŀ
    '����:27327
    
    Err = 0: On Error Resume Next
    If mobjBaseItem Is Nothing Then
        Set mobjBaseItem = CreateObject("zl9BaseItem.clsBaseItem")
    End If
    If mobjBaseItem Is Nothing Then Exit Sub
    'OpenEditWholeSetItem(ByVal frmMain As Object, ByVal cnOracle As ADODB.Connection,
    '      ByVal lngSys As Long, ByVal lngModule As Long, ByVal strPrivs As String, ByVal strItems As String) As Boolean
    'strItems:���,����,�շ�ϸĿID,����,����,ִ�п���|���,����,�շ�ϸĿID,����,����,ִ�п���|��
    Err = 0: On Error GoTo ErrHand:
   If mbytInState = 1 Then
        '�鿴
         strSQL = _
        " Select Nvl(A.�۸񸸺�,A.���) as ���,A.�շ����,A.��������,A.�շ�ϸĿID,A.ִ�в���ID," & _
        "       ��   Avg(Nvl(A.����,1)) as ����, Avg(A.����) ����, Sum(A.��׼����) as ����,B.ִ�п���, B.�Ƿ���,M.��������" & _
        " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼  A") & ",�շ���ĿĿ¼ B,�������� M" & _
        " Where  A.��¼״̬  IN(0,1,3)  And A.NO=[1]  And A.��¼����=[2] And A.�����־=2 And Nvl(A.�ಡ�˵�,0)=0  " & _
        "               And a.�շ�ϸĿID=b.ID And a.�շ�ϸĿID=M.����ID(+) " & _
                        IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[3]", "") & _
        "  Group by Nvl(A.�۸񸸺�,A.���),A.�շ����,A.�շ�ϸĿID,A.��������,A.ִ�в���id,B.ִ�п���,B.�Ƿ���,M.��������" & _
        " Order by ���"
        If mstrTime <> "" Then
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrInNO, 2, CDate(mstrTime))
        Else
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrInNO, 2)
        End If
        With rsTemp
            Do While Not .EOF
                 '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
                If InStr(1, ",4,5,6,7,", "," & Nvl(!�շ����)) > 0 Then
                    lngִ�п���ID = 0
                ElseIf InStr(1, ",0,4", Val(Nvl(!ִ�п���))) > 0 Then
                    lngִ�п���ID = Val(Nvl(!ִ�в���ID))
                Else
                    lngִ�п���ID = 0
                End If
                
                dbl�۸� = 0
                If Val(Nvl(!�Ƿ���)) = 1 Then
                    If InStr(1, "5,6,7", Nvl(!�շ����)) > 0 Or (Nvl(!�շ����) = "4" And Val(Nvl(!��������)) = 1) Then
                        'ҩƷ,����������Ϊ��ȱʡ�۸�,���Բ�����(ͨ��������)
                        dbl�۸� = 0
                    Else
                        dbl�۸� = Val(Nvl(!����))
                    End If
                End If
                strItems = strItems & "|" & Val(Nvl(!���)) & "," & Val(Nvl(!��������)) & "," & Val(Nvl(!�շ�ϸĿID)) & "," & Val(Nvl(!����)) & "," & Val(Nvl(!����)) & "," & dbl�۸� & "," & lngִ�п���ID
                .MoveNext
            Loop
        End With
         If strItems = "" Then
            MsgBox "����δ�����κ���Ϣ,���ܱ���Ϊ�����շ���Ŀ,����!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
            Exit Sub
        End If
        strItems = Mid(strItems, 2)
   Else
        With mobjBill
            strItems = ""
            For i = 1 To .Details.Count
                 '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
                If InStr(1, ",4,5,6,7,", "," & .Details(i).Detail.���) > 0 Then
                    lngִ�п���ID = 0
                ElseIf InStr(1, ",0,4", .Details(i).Detail.ִ�п���) > 0 Then
                    lngִ�п���ID = .Details(i).ִ�в���ID
                Else
                    lngִ�п���ID = 0
                End If
                '����:52349
                dbl���� = .Details(i).����
                dbl�۸� = IIf(.Details(i).Detail.���, .Details(i).InComes(1).��׼����, 0)
                If InStr(",5,6,7,", .Details(i).�շ����) > 0 And gblnסԺ��λ Then
                    dbl���� = Format(dbl���� * .Details(i).Detail.סԺ��װ, gstrFeePrecisionFmt)
                    dbl�۸� = Format(dbl�۸� / .Details(i).Detail.סԺ��װ, gstrFeePrecisionFmt)
                End If
                strItems = strItems & "|" & .Details(i).��� & "," & .Details(i).�������� & "," & .Details(i).�շ�ϸĿID & "," & .Details(i).���� & "," & dbl���� & "," & dbl�۸� & "," & lngִ�п���ID
             Next
             If strItems = "" Then
                MsgBox "����δ�����κ���Ϣ,���ܱ���Ϊ�����շ���Ŀ,����!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                Exit Sub
            End If
            strItems = Mid(strItems, 2)
        End With
    End If
    tmrStatuPati.Enabled = False
    Call mobjBaseItem.OpenEditWholeSetItem(Me, gcnOracle, glngSys, 1150, mstrPrivsOpt, strItems)
    tmrStatuPati.Enabled = True
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
 Private Sub ReSetDefaultִ�п���(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������ȱʡ��ִ�п���
    '����:���˺�
    '����:2010-09-03 16:21:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng���˿���ID As Long, lngDoUnit As Long, strҩ��IDs As String
    
    Dim dblStock As Double
    Err = 0: On Error GoTo ErrHand:
    With mobjBill.Details(lngRow)
         '���ĺ�ҩƷ����
         '����ִ�п���ȱʡΪ���˲���,�������ָ����,��Ϊָ������
        If .Detail.��� = "4" Then
            lngDoUnit = IIf(glng���ϲ��� > 0, glng���ϲ���, mobjBill.����ID)
            If lngDoUnit = 0 Then lngDoUnit = Get��������ID
        End If
        '���˿���ID
        lng���˿���ID = mobjBill.����ID
        If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)
        
        lngDoUnit = Get�շ�ִ�п���ID(.Detail.���, .Detail.ID, _
             .Detail.ִ�п���, lng���˿���ID, Get��������ID, Get������Դ, lngDoUnit, mobjBill.����ID, .ִ�в���ID)
       .ִ�в���ID = lngDoUnit
        
        If InStr(",5,6,7,", .Detail.���) > 0 Then
            '��ǰ��ҩƷ���
            If Not gbln���뷢ҩ Then
                dblStock = GetStock(.Detail.ID, lngDoUnit)
                If gblnסԺ��λ Then
                    dblStock = dblStock / .Detail.סԺ��װ
                End If
                  .Detail.��� = dblStock
                Call ShowStock(.Detail.����, .Detail.���)
            Else
                strҩ��IDs = Decode(.Detail.���, "5", gstr��ҩ��, "6", gstr��ҩ��, "7", gstr��ҩ��)
                If strҩ��IDs <> "" Then
                    dblStock = GetMultiStock(.Detail.ID, strҩ��IDs)
                    If gblnסԺ��λ Then
                        dblStock = dblStock / .Detail.סԺ��װ
                    End If
                    .Detail.��� = dblStock
                    Call ShowStock(.Detail.����, .Detail.���)
                End If
            End If
        ElseIf .Detail.��� = "4" And .Detail.�������� Then
            dblStock = GetStock(.Detail.ID, lngDoUnit, .Detail.����)
            .Detail.��� = dblStock
            Call ShowStock(.Detail.����, .Detail.���)
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
 End Sub
Private Sub cmdSelWholeSet_Click()
    'ѡ������Ŀ
    Dim rsSel As ADODB.Recordset, lng����ID As Long, lng��������ID As Long
    Dim tmpBill As New ExpenseBill, bytӤ���� As Byte, Curdate As Date
    Dim curTotal  As Currency, rsTmp As ADODB.Recordset, i As Long
    Dim intInsure As Integer
    intInsure = 0
    If mobjBill Is Nothing Then
        If mrsInfo Is Nothing Then
            MsgBox "����ѡ����,����!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        End If
        lng����ID = Val(Nvl(mrsInfo!����ID))
        intInsure = Val(Nvl(mrsInfo!����))
        If cbo��������.ListIndex < 0 Then
            lng��������ID = 0
        Else
            lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        End If
        If cboBaby.ListIndex < 0 Then
            bytӤ���� = 0
        Else
            bytӤ���� = cboBaby.ItemData(cboBaby.ListIndex)
        End If
    Else
        lng����ID = mobjBill.����ID: lng��������ID = mobjBill.��������ID: bytӤ���� = mobjBill.Ӥ����
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then intInsure = Val(Nvl(mrsInfo!����))
        End If
    End If
    If lng����ID = 0 Then
            MsgBox "����ѡ����,����!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
    End If
    
    If zlSelectWholeItems(Me, mlngModule, mstrPrivsOpt, rsSel) = False Then Exit Sub
    If rsSel Is Nothing Then Exit Sub
    Err = 0: On Error GoTo ErrHand:
    Screen.MousePointer = 11
    Set tmpBill = ImportWholeSet(Me, intInsure, rsSel, lng����ID, gblnסԺ��λ, lng��������ID, bytӤ����, 2, chk�Ӱ�.Value = 1, _
        0, Get������Դ, UserInfo.����, zlStr.NeedName(cbo������.Text), , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�, _
        IIf(mbln���� And mlng��ҳID <> 0, mlng��ҳID, 0), IIf(mbln���� And mstr���ת��ʱ�� <> "", mlngDeptID, 0), IIf(mbln���� And mstr���ת��ʱ�� <> "", mlngUnitID, 0))
    '��������
    '�������Ĳ�����Ϣ
    Set mobjBill = New ExpenseBill
    Set mobjBill = tmpBill
    Dim bln��ҩ As Boolean
    bln��ҩ = False
    With mobjBill
        For i = 1 To .Details.Count - 1
            If .Details(i).�շ���� = "7" Then
                bln��ҩ = True
                Exit For
            End If
            Exit For
        Next
    End With
    Curdate = zlDatabase.Currentdate
    mobjBill.NO = cboNO.Text
    mobjBill.�Ǽ�ʱ�� = Curdate
    mobjBill.����Ա��� = UserInfo.���
    mobjBill.����Ա���� = UserInfo.����
    mobjBill.�Ӱ��־ = chk�Ӱ�.Value
    mobjBill.Ӥ���� = cboBaby.ItemData(cboBaby.ListIndex)
    
    'ȡ��ǰʱ��:33744
    If mbln���� Then
        If mstr���ת��ʱ�� <> "" Then
            txtDate.Text = Format(CDate(mstr���ת��ʱ��) - 1 / 24 / 60, "yyyy-mm-dd HH:MM:SS")
            txtDate.ForeColor = vbBlue
        End If
    Else
        txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
    End If
    
    Bill.Redraw = False
    Bill.ClearBill
    Bill.Rows = mobjBill.Details.Count + 1
    
    Call InitBillColumnColor
    '���ʷ��౨��
    mstrWarn = ""
        
    Call Set�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mobjBill.������, mobjBill.��������ID)
        
    '������Ķ����˺�ȷ���ѱ��,�ټ���۸�
    Call CalcMoneys
    Call ShowDetails
    Call ShowMoney
    With Bill
        For i = 1 To .Rows - 1
            .TextMatrix(i, BillCol.��) = i
        Next
    End With

    Bill.Redraw = True
    'ˢ�²��˷�����Ϣ
    If mrsInfo.State = 1 Then
        'ˢ�²���Ԥ������Ϣ
        curTotal = GetBillTotal(mobjBill)
        Set rsTmp = GetMoneyInfo(mrsInfo!����ID, 0, True, 2)
        If Not rsTmp Is Nothing Then
            cmdOK.Tag = rsTmp!Ԥ�����
            cmdCancel.Tag = rsTmp!�������
            txtʵ��.Tag = rsTmp!Ԥ����� - rsTmp!�������
        Else
            cmdOK.Tag = 0: cmdCancel.Tag = 0: txtʵ��.Tag = 0
        End If
        Call SetStatuPatiInfor(Val(cmdOK.Tag), Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), Val(txtʵ��.Tag) - IIf(gbytBilling = 0, curTotal, 0))
    End If
    '���¼���ͳ����
    Call ReCalcInsure
    Call SetDrawDrugDeptEnabled
    Screen.MousePointer = 0
    If bln��ҩ Then
         cmd�䷽_Click
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
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
        If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And Not mblnNotClick Then
        txtPatient.Text = ""
    End If
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("���֤��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub


Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, bytsubs As Byte
    Dim bln��������ۿ� As Boolean
 
    Dim lngMainRow As Long
    
    If mbytInState <> 0 Or chkCancel.Value = 1 Then Cancel = True: Exit Sub
    
     
    If mobjBill.Details.Count >= Row Then
        '��������Ŀ����ɾ��ȷ��
        For i = Row + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).�������� = Row Then bytsubs = bytsubs + 1
        Next
        If bytsubs > 0 Then
            If MsgBox("����Ŀ���� " & bytsubs & " ��������Ŀ,ɾ������ĿҲ��ɾ�����Ĵ�����Ŀ,������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
        ElseIf mobjBill.Details(Row).�������� <> 0 Then '������Ŀɾ��ȷ��
            If MsgBox("����Ŀ��[" & mobjBill.Details(mobjBill.Details(Row).��������).Detail.���� & "]�Ĵ�����Ŀ,ȷ��Ҫɾ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            Else
                bln��������ۿ� = gbln��������ۿ�
            End If
        ElseIf MsgBox("ȷʵҪɾ�����շ���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
        
        If bln��������ۿ� Then lngMainRow = mobjBill.Details(Bill.Row).�������� '����Ǵ���,ɾ��֮ǰ���´���Ĵ�������,���������,����ɾ��,��������
        

        
        'ɾ������
        For i = mobjBill.Details.Count To Row + 1 Step -1
            If mobjBill.Details(i).�������� = Row Then
                Call DeleteDetail(i) '��˳��ɾ���������
            End If
        Next
        Call DeleteDetail(Row) 'ɾ������
        
        '���¼��㲢ˢ��
        If bln��������ۿ� Then
            If CheckItemHaveSub(lngMainRow) Then
                Call Calc��������ʵ��(lngMainRow)
            Else
                Call CalcMoney(lngMainRow, False) 'ֻ��һ��������,����ȫ����ɾ��ʱ,������ͨ���������
            End If
        End If
            
        Call ShowDetails
        Call ShowMoney
                
        Bill.TxtVisible = False
        Bill.CmdVisible = False
        Bill.CboVisible = False
        
        Cancel = True '���ÿؼ�������ɾ��
        
        mlngPreRow = 0    '��ʾ�иı���
        Call Bill_EnterCell(Bill.Row, Bill.Col)
        Call SetDrawDrugDeptEnabled
    ElseIf Row = 1 Then
        For i = 1 To Bill.Cols - 1
            Bill.TextMatrix(Row, i) = ""
        Next
        Cancel = True
    End If
    Call SetColNum(Row)
End Sub

Private Sub ShowStock(strҩƷ As String, dbl��� As Double)
'���ܣ���ʾҩƷ�����ĵĿ��
    If InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0 Then
        sta.Panels(Pan.C2��ʾ��Ϣ).Text = "[" & strҩƷ & "]���ÿ��:" & dbl���
    Else
        sta.Panels(Pan.C2��ʾ��Ϣ).Text = "[" & strҩƷ & "]" & IIf(dbl��� > 0, "��", "��") & "���."
    End If
End Sub

Private Sub Bill_cboClick(ListIndex As Long)
    Dim dblStock As Double
    Dim lngִ�п��� As Long, strִ�п��� As String
    If mblncboClick Then Exit Sub  '����ͬһ������������bill��ֵѭ������,ע�����κ�exit sub ֮ǰ����mblncboClick = False
    mblncboClick = True
    'ҩƷ�����
    If ListIndex <> -1 And Bill.TextMatrix(0, Bill.Col) = "ִ�п���" Then
        If mobjBill.Details.Count >= Bill.Row Then
            With mobjBill.Details(Bill.Row)
                If .ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
                    lngִ�п��� = .ִ�в���ID: strִ�п��� = Bill.TextMatrix(Bill.Row, Bill.Col)
                    .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                    Bill.TextMatrix(Bill.Row, Bill.Col) = Bill.CboText
                    
                    If InStr(",5,6,7,", .�շ����) > 0 And Not gbln���뷢ҩ Then
                        'ȡ���
                        dblStock = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If gblnסԺ��λ Then
                            dblStock = dblStock / .Detail.סԺ��װ
                        End If
                        .Detail.��� = dblStock  '��¼��ǰ��ҩƷ���
                        Call ShowStock(.Detail.����, .Detail.���)
                        
                        'ҩ���ı�,ʵ��ҩƷ���¼���۸�
                        'If .Detail.��� Then    '����ѱ�ļ��㷽ʽ�ǳɱ��ۼ��շ�,����Ҫ����۸�,����򻯲����ж�
                            Call CalcMoneys(Bill.Row)   '�����Ҫ���ܼ���,����������ʵ��
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney
                        'End If
                    ElseIf .�շ���� = "4" And .Detail.�������� Then
                        'ȡ���
                        dblStock = GetStock(.�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                        .Detail.��� = dblStock
                        Call ShowStock(.Detail.����, .Detail.���)
                        
                        '���ϲ��Ÿı�,ʱ���������¼���۸�
                        If .Detail.��� Then
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney
                        End If
                        
                    ElseIf InStr(",4,5,6,7,", .�շ����) = 0 Then
                        If CheckItemHaveSub(Bill.Row) Then Call SetSubItemDept(Bill.Row) '������ڴ���,��ı��ҩƷ�е�ִ�п���
                    End If
                    If mrsInfo.State = 1 Then
                        If Not IsNull(mrsInfo!����) And MCPAR.ʵʱ��� And mobjBill.Details(Bill.Row).���� <> 0 Then
                            If gclsInsure.CheckItem(Val(mrsInfo!����), 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, IIf(gbytBilling = 1 Or mblnSavePrice, 1, 0), Bill.Row)) = False Then
                                Bill.Text = "": Bill.TxtVisible = False
                                Bill.cboObj.Text = strִ�п���: .ִ�в���ID = lngִ�п���
                                mblncboClick = False: Exit Sub
                            End If
                        End If
                    End If
                        
                    If mobjBill.Details(Bill.Row).���� <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, IIf(gbytBilling = 1 Or mblnSavePrice, 1, 0), Bill.Row)) = False Then
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.cboObj.Text = strִ�п���: .ִ�в���ID = lngִ�п���
                            mblncboClick = False: Exit Sub
                        End If
                    End If
                End If
            End With
        End If
    End If
    
    mblncboClick = False
End Sub

Private Sub SelALLRow()
'���ܣ�ʵ���˷�ʱ��ȫѡ
    Dim i As Long
    
    If Bill.TextMatrix(0, Bill.Cols - 1) = "����" Then
        For i = 1 To Bill.Rows - 1
            If Bill.TextMatrix(i, BillCol.��Ŀ) <> "" Then
                Bill.TextMatrix(i, Bill.Cols - 1) = "��"
            End If
        Next
    End If
End Sub

Private Sub ClearALLRow()
'���ܣ�ʵ���˷�ʱ��ȫ��
    Dim i As Long
    
    If Bill.TextMatrix(0, Bill.Cols - 1) = "����" Then
        For i = 1 To Bill.Rows - 1
            Bill.TextMatrix(i, Bill.Cols - 1) = ""
        Next
    End If
End Sub

Private Sub Bill_CellCheck(Row As Long, Col As Long)
'˵��������ȫ��Ϊ��Ҫ����,������ȫ��Ϊ��������
    Dim i As Long, strCheck As String, bytTime As Byte
    Dim blnReSet As Boolean
    If Bill.TextMatrix(Row, BillCol.��Ŀ) = "" Then Bill.TextMatrix(Row, Col) = "": Exit Sub
    '����������,����û�б�Ҫ���к���Ĵ���
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then Exit Sub
    
    
    '������δ��������Ч
    If mobjBill.Details.Count < Row Then
        Bill.TextMatrix(Row, Col) = "": Exit Sub
    End If
    
    strCheck = Bill.TextMatrix(Row, Col)
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�շ���� = "F" And mobjBill.Details(i).���ӱ�־ = 0 And i <> Row Then bytTime = bytTime + 1
    Next
    
    blnReSet = bytTime > 0
    If blnReSet = False Then     '����ֻ���ڸ����������ָĳ���������,��Ҫ���¼ƴ���:25495
        blnReSet = (strCheck = "" And mobjBill.Details(Row).�շ���� = "F" And mobjBill.Details(Row).���ӱ�־ = 1)
    End If
    
    If blnReSet Then
        With mobjBill.Details(Row)
            .���ӱ�־ = IIf(strCheck = "", 0, 1)
            Call CalcMoneys(Row)
            Call ShowDetails(Row)
        End With
        Call ShowMoney
    ElseIf strCheck <> "" Then
        Bill.TextMatrix(Row, Col) = ""
        MsgBox "�����б�Ȼ��һ���������Ǹ���������", vbInformation, gstrSysName
        Exit Sub
    End If
    
End Sub

Private Sub Bill_CommandClick()
    Dim lng��Ŀid As Long, blnCancel As Boolean, bln��ʿ As Boolean
    Dim str��� As String, str��׼��Ŀ As String
    Dim int������Դ As Integer, int���� As Integer
    Dim str�ų���� As String
    Call GetOperatorInfo(mrs������, mobjBill.������, bln��ʿ)
    If gbln�շ���� Then
        If Bill.RowData(Bill.Row) <> 0 Then
            str��� = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
        Else
            str��� = IIf(bln��ʿ, "'E','M','4'", gstr�շ����)
        End If
    Else
        str��� = IIf(bln��ʿ, "'E','M','4'", gstr�շ����)
    End If
    '--0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
    If mrsInfo.State = 1 Then
        If Not IsNull(mrsInfo!����) Then
            int���� = mrsInfo!����
            '���˺�:24862
            If zl_Check��׼��Ŀ(gclsInsure, int����, Val(Nvl(mrsInfo!����ID)), False) Then str��׼��Ŀ = Get������׼��Ŀ(Val(Nvl(mrsInfo!����ID)), "A.ID")
        End If
        If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
            int������Դ = 2
        ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
            int������Դ = 1
        End If
    Else
        int������Դ = 2
    End If
   If zlCheckBill���ڷ�ɢװ��ҩ() = True Then
        mblnSelect = False: Exit Sub
    End If
    mlng���� = -1
    lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, int������Դ, int����, gblnסԺ��λ, str���, , , str��׼��Ŀ, _
        0, str�ų����, False, mbln����ˢ��, mlng����, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    If lng��Ŀid <> 0 Then
        Bill.Text = lng��Ŀid
        mblnSelect = True
        Call Bill_KeyDown(13, 0, blnCancel)
        Bill.SetFocus
        If Not blnCancel Then
            Bill.Text = "": Bill.TxtVisible = False
            Call zlCommFun.PressKey(13)
        End If
    End If
    mblnSelect = False
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
'���ܣ�����������
    Dim rsҩƷ��Ϣ As ADODB.Recordset
    Dim lng��Ŀid As Long, str��� As String, bln��ʿ As Boolean
    Dim str��׼��Ŀ As String, int������Դ As Integer, lng���˿���ID As Long, int���� As Integer
    Dim dblStock As Double, strScope As String, i As Long
    Dim dblPreTime As Double, dblPreMoney As Double, dblNum As Double, lngOld���� As Long
    Dim blnSkip As Boolean, curTotal As Currency
    Dim lngDoUnit As Long, strժҪ As String, blnInput As Boolean
    Dim strҩ��IDs As String, bln�������� As Boolean, cur��� As Currency
    Dim curItemMoney As Currency
    Dim colStock As Collection, str�ų���� As String
    Dim strPriceGrade As String
    
    On Error GoTo errH
    
    If KeyCode = 13 And Bill.Active Then
        If mbytInState = 2 Then
            If Bill.Col = Bill.Cols - 1 And Bill.Row = Bill.Rows - 1 Then
                Cancel = True: Exit Sub
            ElseIf Bill.TextMatrix(0, Bill.Col) <> "ִ�п���" Then
                Exit Sub
            End If
        End If
        If Bill.ColData(Bill.Col) = BillColType.Text_UnModify Then Exit Sub
                        
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "���"
                If Bill.ListIndex <> -1 Then '���������ʱ���ᶨλ�������
                    If Bill.RowData(Bill.Row) <> Bill.ItemData(Bill.ListIndex) Then
                        'һ���ĸ��շ����,�����(����)ԭ�и���Ŀ����
                        For i = 2 To Bill.Cols - 1
                            Bill.TextMatrix(Bill.Row, i) = ""
                        Next
                        If mobjBill.Details.Count >= Bill.Row Then
                            Set mobjBill.Details(Bill.Row).Detail = New Detail
                            Set mobjBill.Details(Bill.Row).InComes = New BillInComes
                            With mobjBill.Details(Bill.Row)
                                .�շ�ϸĿID = 0: .�շ���� = ""
                            End With
                            Call CalcMoneys
                            Call ShowMoney
                        End If
                    End If
                    Bill.RowData(Bill.Row) = Bill.ItemData(Bill.ListIndex) '��ʱ��RowData��¼��ѡ����շ����
                    ''���������ǲ�ҩ���,(���ܴ���ѡ������,���,�ݲ�֧�����ַ�ʽ)
                    'If Chr(Bill.ItemData(Bill.ListIndex)) = "7" Then
                    'Call cmd�䷽_Click
                    'End If
                    
                End If
            Case "��Ŀ"
                '����Ŀȷ��,���շ�ϸĿ��Ӧ�ĳ�����������,ͬʱ���ﴦ���շѴ�����Ŀ
                If Bill.Text <> "" Then
                    '��������������Ŀ�ϰ��س�,��ѡ����ѡ��
                    If mobjBill.Details.Count >= Bill.Row Then
                        'ͨ����ťѡ���Ƿ��ص�ID,�����������ı�,�����һ����,�򲻸ı�
                        If Bill.TextMatrix(Bill.Row, BillCol.��Ŀ) = Bill.Text Then
                            Bill.TxtVisible = False
                            Bill.CmdVisible = False
                            Exit Sub
                        End If
                    End If
                
                    sta.Panels(2).Text = ""
                    sta.Panels("MedicareType").Text = ""
                    blnInput = True
                    If mblnSelect Then
                        mblnSelect = False '��������ñ�־
                        Set mobjDetail = GetInputDetail(Val(Bill.Text))
                    Else
                        If gbln�շ���� Then
                            If Bill.RowData(Bill.Row) = 0 Then
                                sta.Panels(2) = "û��ȷ���������,�����������"
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                            str��� = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
                        Else
                            Call GetOperatorInfo(mrs������, mobjBill.������, bln��ʿ)
                            str��� = IIf(bln��ʿ, "'E','M','4'", gstr�շ����)
                            
                        End If
                        
                        If mrsInfo.State = 1 Then
                            If Not IsNull(mrsInfo!����) Then
                                int���� = mrsInfo!����
                                '���˺�:24862
                                If zl_Check��׼��Ŀ(gclsInsure, int����, Val(Nvl(mrsInfo!����ID)), False) Then str��׼��Ŀ = Get������׼��Ŀ(Val(Nvl(mrsInfo!����ID)), "A.ID")
                            End If
                            If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
                                int������Դ = 2
                            ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
                                int������Դ = 1
                            End If
                        Else
                            int������Դ = 2
                        End If
                        If zlCheckBill���ڷ�ɢװ��ҩ Then
                            '���ڷ�ɢװ��,�����оͲ��ܽ���¼��
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                        mlng���� = -1
                        lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, int������Դ, int����, gblnסԺ��λ, str���, _
                            Bill.Text, Bill.TxtHwnd, str��׼��Ŀ, 0, , False, mbln����ˢ��, mlng����, _
                             mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
                        If lng��Ŀid <> 0 Then
                            Set mobjDetail = GetInputDetail(lng��Ŀid)
                            
                            If int���� <> 0 Then sta.Panels("MedicareType").Text = Getҽ������(lng��Ŀid, int����)
                        Else
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    Bill.TxtVisible = False '(���Ӳ���)
                    
                    '�������ò��˲�������
                    If InStr(",5,6,7,", mobjDetail.���) = 0 And mrsInfo.State = 1 Then
                        If Not CheckFeeItemLimitDept(mobjDetail.ID, IIf(mbytUseType = 2, UserInfo.����ID, mobjBill.����ID), IIf(mbytUseType = 2, UserInfo.����ID, mobjBill.����ID)) Then
                            If mbytUseType = 2 Then
                                MsgBox "���շ���Ŀ�Ե�ǰ���˲����Ϳ��Ҳ����ã�", vbInformation, gstrSysName
                            Else
                                MsgBox "���շ���Ŀ�Ե�ǰ���˲����Ϳ��Ҳ����ã�", vbInformation, gstrSysName
                            End If
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    'ҽ�����˷�������
                    If mrsInfo.State = 1 Then
                        If Not IsNull(mrsInfo!����) Then
                            If mobjDetail.Ҫ������ And Not mrsMedAudit Is Nothing Then
                                mrsMedAudit.Filter = "��ĿID=" & mobjDetail.ID
                                If mrsMedAudit.RecordCount = 0 Then
                                    MsgBox "��ǰ����δ����׼ʹ��[" & mobjDetail.���� & "]��", vbInformation, gstrSysName
                                    Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                                ElseIf Not IsNull(mrsMedAudit!��������) Then
                                    If mrsMedAudit!�������� <= 0 Then
                                        MsgBox "��ǰ����ʹ��[" & mobjDetail.���� & "]�Ѵﵽ��׼��ʹ������" & FormatEx(mrsMedAudit!ʹ������ / IIf(gblnסԺ��λ, mobjDetail.סԺ��װ, 1), 5) & "��", vbInformation, gstrSysName
                                        Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    '�շ��뷢ҩ����ʱ����������ʱ�ۼ�����ҩƷ
                    If InStr(",5,6,7,", mobjDetail.���) > 0 And gbln���뷢ҩ Then
                        If mobjDetail.��� Or mobjDetail.���� Then
                            MsgBox "��ҩ���봦��ʱ��������ʱ�ۻ����ҩƷ��", vbInformation, gstrSysName
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '��鶾�����ͼ�ֵ����Ȩ��
                    If InStr(",5,6,7,", mobjDetail.���) > 0 Then
                        Set rsҩƷ��Ϣ = ReadҩƷ��Ϣ(mobjDetail.ID)
                        If Not rsҩƷ��Ϣ Is Nothing Then
                            If IIf(IsNull(rsҩƷ��Ϣ!�������), "", rsҩƷ��Ϣ!�������) = "����ҩ" _
                                And InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then
                                MsgBox """" & mobjDetail.���� & """Ϊ����ҩƷ����û��Ȩ�޶Ը���ҩƷ���ʣ�", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            ElseIf IIf(IsNull(rsҩƷ��Ϣ!�������), "", rsҩƷ��Ϣ!�������) = "����ҩ" _
                                And InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then
                                MsgBox """" & mobjDetail.���� & """Ϊ����ҩƷ����û��Ȩ�޶Ը���ҩƷ���ʣ�", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            ElseIf (IIf(IsNull(rsҩƷ��Ϣ!��ֵ����), "", rsҩƷ��Ϣ!��ֵ����) = "����" _
                                Or IIf(IsNull(rsҩƷ��Ϣ!��ֵ����), "", rsҩƷ��Ϣ!��ֵ����) = "����") _
                                And InStr(mstrPrivsOpt, ";����ҩƷ����;") = 0 Then
                                MsgBox """" & mobjDetail.���� & """Ϊ���ػ򰺹�ҩƷ����û��Ȩ�޶Ը���ҩƷ���ʣ�", vbInformation, gstrSysName
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    '���ҩƷ�����Ƿ��ظ�:������ʱ��ͬһҩ���������ظ�(����ֻ����)
                    If InStr(",5,6,7,", mobjDetail.���) > 0 Or _
                        (mobjDetail.��� = "4" And mobjDetail.��������) Then
                        If PhysicExist(mobjDetail, Bill.Row) Then
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '��鴦��ְ��
                    If InStr(",5,6,7,", mobjDetail.���) > 0 And mbln����ְ���� Then
                        mobjDetail.����ְ�� = Get����ְ��(mobjDetail.ID)
                        If cboҽ�Ƹ���.ListIndex <> -1 Then
                            'ҽ���򹫷Ѳ���
                            '����:45605
                            If zlIsCheckMedicinePayMode(zlStr.NeedName(cboҽ�Ƹ���)) Then
                                If CheckDuty(mobjDetail, False) > 0 Then
                                    Bill.TxtSetFocus: Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        '���в���
                        If CheckDuty(mobjDetail, True) > 0 Then
                            Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '��ȡҩƷ�����Ϣ
                    '����ִ�п���ȱʡΪ���˲���,�������ָ����,��Ϊָ������
                    If mobjDetail.��� = "4" Then
                        lngDoUnit = IIf(glng���ϲ��� > 0, glng���ϲ���, mobjBill.����ID)
                        If lngDoUnit = 0 Then lngDoUnit = Get��������ID
                    End If
                    
                    '���˿���ID
                    lng���˿���ID = mobjBill.����ID
                    If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)
                    
                    lngDoUnit = Get�շ�ִ�п���ID(mobjDetail.���, mobjDetail.ID, _
                        mobjDetail.ִ�п���, lng���˿���ID, Get��������ID, Get������Դ, lngDoUnit, mobjBill.����ID)
                    
                    
                    If ReadDrugAndStuffStock(lngDoUnit, mobjDetail) = False Then
                        Bill.TxtSetFocus: Cancel = True: Exit Sub
                    End If
                    
                     '��������
                    If InStr(",5,6,7,", mobjDetail.���) > 0 And mbln����������� Then
                        mobjDetail.�������� = Get��������(mobjDetail.ID)
                    End If
                    
                    '������Ŀ��Ӧ���
                    If mrsInfo.State = 1 Then
                        If Not IsNull(mrsInfo!����) Then
                            If InStr(",5,6,7,", mobjDetail.���) > 0 Then
                                strPriceGrade = mstrҩƷ�۸�ȼ�
                            ElseIf mobjDetail.��� = "4" Then
                                strPriceGrade = mstr���ļ۸�ȼ�
                            Else
                                strPriceGrade = mstr��ͨ�۸�ȼ�
                            End If
                            If Not CheckMediCareItem(mobjDetail.ID, mrsInfo!����, mobjDetail.����, _
                                mobjDetail.��� = False, , strPriceGrade) Then
                                Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    '����ժҪ(ȡ���е����Ա��޸�)
                    If mobjBill.Details.Count >= Bill.Row Then
                        If mobjBill.Details(Bill.Row).Detail.ID = mobjDetail.ID Then
                            strժҪ = mobjBill.Details(Bill.Row).ժҪ
                        End If
                    End If
                    
                    '������޸ĸ��շ�ϸĿ��
                    Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                    '59051:�ȵ���GetItemInfor
                    '����ժҪ(������������и���ժҪ)
                    If mobjBill.Details(Bill.Row).Detail.����ժҪ Then
                        If frmInputBox.InputBox(Me, "ժҪ", "������""" & mobjBill.Details(Bill.Row).Detail.���� & """��ժҪ��Ϣ:", 200, 3, True, False, strժҪ) Then
                            mobjBill.Details(Bill.Row).ժҪ = strժҪ
                        End If
                    Else
                        If mrsInfo.State = 1 Then '90304
                            strժҪ = gclsInsure.GetItemInfo(Val(Nvl(mrsInfo!����)), mrsInfo!����ID, mobjBill.Details(Bill.Row).�շ�ϸĿID, strժҪ, 2)
                        Else
                            strժҪ = gclsInsure.GetItemInfo(0, 0, mobjBill.Details(Bill.Row).�շ�ϸĿID, strժҪ, 2)
                        End If
                        mobjBill.Details(Bill.Row).ժҪ = strժҪ
                    End If
                    Call CalcMoney(Bill.Row)                        '��ʱ,��ʹ�������������,���û������
                    '�����ҽ��Calcmoney�п��ܷ���ժҪ
                    If mobjBill.Details(Bill.Row).ժҪ <> "" Then strժҪ = mobjBill.Details(Bill.Row).ժҪ

                    
                    '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                    mrsWarn.Filter = ""
                    If mrsWarn.RecordCount > 0 And mrsInfo.State = 1 And mobjBill.Details.Count = Bill.Row Then
                        curTotal = GetBillTotal(mobjBill)
                        If curTotal > 0 Then
                            '���˺�:24491
                            curItemMoney = GetBillRowTotal(mobjBill.Details(Bill.Row).InComes)
                            cur��� = Val(txtʵ��.Tag)
                            If gbln�����������۷��� Then cur��� = Val(txtʵ��.Tag) - GetPriceMoneyTotal(1, mrsInfo!����ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                            gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!���� & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), Val("" & mrsInfo!����ID), mrsInfo!���ò���, mrsWarn, cur���, mrsInfo!���ն� - mcurModiMoney, curTotal, _
                                        IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, _
                                        mstrWarn, , gblnPrice And (gbytBilling = 0 And mstrInNO = "" Or gbytBilling = 1), curItemMoney)
                            If gbytWarn = 2 Or gbytWarn = 3 Then
                                mobjBill.Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                                Bill.Text = "": Cancel = True: Exit Sub
                            ElseIf gbytBilling = 0 And gblnPrice And mstrInNO = "" Then '����ģʽ�޸�ʱ��Ч
                                If gbytWarn = 1 Or gbytWarn = 4 Then
                                    cmdPrice.Visible = True: cmdOK.Visible = True: Call SetButtonPlace
                                ElseIf gbytWarn = 5 Then
                                    cmdPrice.Visible = True: cmdOK.Visible = False: Call SetButtonPlace
                                End If
                            End If
                        End If
                    End If
                    
                    If mrsInfo.State = 1 Then
                        If Not IsNull(mrsInfo!����) And MCPAR.ʵʱ��� And mobjBill.Details(Bill.Row).���� <> 0 Then
                            If gclsInsure.CheckItem(Val(mrsInfo!����), 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, IIf(gbytBilling = 1 Or mblnSavePrice, 1, 0), Bill.Row)) = False Then
                                mobjBill.Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    If mobjBill.Details(Bill.Row).���� <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, IIf(gbytBilling = 1 Or mblnSavePrice, 1, 0), Bill.Row)) = False Then
                            mobjBill.Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                            Bill.Text = "": Cancel = True: Exit Sub
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                    '�������ͼ��
                    Call Check��������(Bill.Row)
                    '�������
                    If gcurMaxMoney > 0 Then
                        If Bill.TextMatrix(Bill.Row, BillCol.����) * Bill.TextMatrix(Bill.Row, BillCol.����) * Bill.TextMatrix(Bill.Row, BillCol.����) > gcurMaxMoney Then
                            If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                mobjBill.Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                                Exit Sub
                            End If
                        End If
                    End If
                    Call SetDrawDrugDeptEnabled
                    Bill.Text = "": Bill.SetFocus
                End If
                
                If mobjBill.Details.Count >= Bill.Row Then
                    mlngPreRow = 0  '�޸�������ʱ,�ָ���ֵ,�Ա���ʾ���
                    With mobjBill.Details(Bill.Row)
                        '��һ�е�����ȷ��
                        If .�շ���� = "7" And gblnPay Then Bill.ColData(BillCol.����) = BillColType.Text  '����
                        If .�շ���� = "F" Then Bill.ColData(BillCol.��־) = BillColType.CheckBox '���ӱ�־
                        
                        '���������������
                        If .Detail.��� And InStr(",5,6,7,", .�շ����) = 0 _
                            And Not (.�շ���� = "4" And .Detail.��������) Then
                            Bill.ColData(BillCol.����) = IIf(gblnTime, BillColType.Text, BillColType.UnFocus) '����
                            Bill.ColData(BillCol.����) = BillColType.Text '����
                        Else
                            Bill.ColData(BillCol.����) = BillColType.Text '����
                            Bill.ColData(BillCol.����) = BillColType.UnFocus '����
                        End If
                        
                        'ִ�п���
                        If InStr(",5,6,7,", .�շ����) > 0 And gbln���뷢ҩ Then
                            Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus: .Key = 1
                        Else
                             '��FillBillComboBox������ListIndexʱ����CboClick�¼�
                            mblncboEnterCell = True: Bill.Col = BillCol.ִ�п���: mblncboEnterCell = False
                            Call FillBillComboBox(Bill.Row, BillCol.ִ�п���, Not blnInput)  'ֱ�ӻس�ʱ����ִ�п���
                            mblncboEnterCell = True: Bill.Col = BillCol.��Ŀ: mblncboEnterCell = False
                            
                            blnSkip = Bill.ListCount = 1
                            If Not blnSkip And InStr(",4,5,6,7,", .�շ����) > 0 Then
                                'ָ���˹̶�ҩ��ʱ,��������ѡ��
                                Select Case .�շ����
                                    Case "4"
                                        blnSkip = glng���ϲ��� > 0 And .ִ�в���ID = glng���ϲ���
                                    Case "5"
                                        blnSkip = glng��ҩ�� > 0 And .ִ�в���ID = glng��ҩ��
                                    Case "6"
                                        blnSkip = glng��ҩ�� > 0 And .ִ�в���ID = glng��ҩ��
                                    Case "7"
                                        blnSkip = glng��ҩ�� > 0 And .ִ�в���ID = glng��ҩ��
                                End Select
                            End If
                            If blnSkip Then
                                Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus: .Key = 1
                            Else
                                Bill.ColData(BillCol.ִ�п���) = BillColType.ComboBox: .Key = Bill.ListCount
                            End If
                        End If
                        
                        If .ִ�в���ID <> lngDoUnit Then
                             If ReadDrugAndStuffStock(.ִ�в���ID, mobjBill.Details(Bill.Row).Detail) = False Then
                                 Bill.TxtSetFocus: Cancel = True: Exit Sub
                             End If
                        End If
                        
                        '����������ϵ����Ч��,��ȷ��ִ�п���֮��
                        If .�շ���� = "4" And .Detail.�������� Then
                            Call CheckValidity(.�շ�ϸĿID, .ִ�в���ID, .����, False) '��ȷ������,��������
                        End If
                        
                         '������Ŀ����,�������շ���Ŀ�д�����Ŀ����δȡ��ȡ,ҩƷ�����ж�,ҩƷ��������������
                        If Bill.TextMatrix(0, Bill.Col) = "��Ŀ" And InStr(",5,6,7,", .�շ����) = 0 Then
                            If (gbln��������ۿ� And mobjBill.Details(Bill.Row).�������� = 0) Or Not gbln��������ۿ� Then  '(����м���,ֻȡһ��)
                                If ShouldDO(Bill.Row) Then
                                   Call SetSubItem
                                   mlngPreRow = 0 'ͨ���б仯��־������ȷ��������
                                End If
                            End If
                        End If
                        
                    End With
                End If
                'ȡ����һ�������ҩ�󵯳��䷽����:38328
'                If mobjBill.Details.Count >= Bill.Row And Bill.Active And Visible Then
'                    If mobjBill.Details(Bill.Row).�շ���� = "7" Then
'                         Call cmd�䷽_Click
'                         Exit Sub
'                    End If
'                End If
                'ֻ����һ�θ���
                If mobjBill.Details.Count >= Bill.Row And Bill.Row >= 2 And Bill.Active And Visible Then
                    If mobjBill.Details(Bill.Row).�շ���� = "7" Then
                        For i = 1 To Bill.Row - 1
                            If mobjBill.Details(i).�շ���� = "7" Then
                                '����ִ�иù��̣�����ᶨλ��һ����Ԫ,�ȶ�λ������,����һ����Ԫ������
                                'ѡ����øù��̣����ú���͸��س������ﲻ���ٻس��������������س���Ч��(�ؼ�ԭ��)��
                                Bill.Col = BillCol.����: Exit For
                            End If
                        Next
                    End If
                End If
            Case "����"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '���ֺϷ���
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) <= 0 Or Val(Bill.Text) <> Int(Val(Bill.Text)) Then
                        MsgBox "����Ӧ��Ϊ����������", vbInformation, gstrSysName
                        Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                    End If
                    
                    '�������
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� * Bill.TextMatrix(Bill.Row, BillCol.����) > gcurMaxMoney Then
                            If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                
                    '����ҩ���Ǵ�����Ŀ�ſɸ��ĸ���(������ı�,����Ҳ��)
                    If mobjBill.Details(Bill.Row).�շ���� = "7" Then
                        '������ʱ��ҩƷ�����ֹ����(û�з�����ʱ��ҩƷ�����޸ĸ���������)
                        If mobjBill.Details(Bill.Row).Detail.���� Or mobjBill.Details(Bill.Row).Detail.��� Then
                            If CSng(Bill.Text) * mobjBill.Details(Bill.Row).���� > mobjBill.Details(Bill.Row).Detail.��� Then
                                MsgBox """" & mobjBill.Details(Bill.Row).Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                        End If
                              
                        '�������ʱ�ۻ������ҩ���ĸ��������Ƿ��㹻
                        For i = 1 To mobjBill.Details.Count
                            If i <> Bill.Row And mobjBill.Details(i).�շ���� = "7" _
                                And (mobjBill.Details(i).Detail.��� Or mobjBill.Details(i).Detail.����) Then
                                If Val(Bill.Text) * mobjBill.Details(i).���� > mobjBill.Details(i).Detail.��� Then
                                    MsgBox "�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    Bill.Text = mobjBill.Details(Bill.Row).����: Cancel = True: Exit Sub
                                End If
                            End If
                        Next
                                                
                        lngOld���� = mobjBill.Details(Bill.Row).����
                        '���㲢ˢ�¸���
                        mobjBill.Details(Bill.Row).���� = Bill.Text
                        Call CalcMoneys(Bill.Row)
                        
                        
                        If mrsInfo.State = 1 Then
                            If Not IsNull(mrsInfo!����) And MCPAR.ʵʱ��� And mobjBill.Details(Bill.Row).���� <> 0 Then
                                If gclsInsure.CheckItem(Val(mrsInfo!����), 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, IIf(gbytBilling = 1 Or mblnSavePrice, 1, 0), Bill.Row)) = False Then
                                    mobjBill.Details(Bill.Row).���� = lngOld����
                                    Call CalcMoneys(Bill.Row)
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        If mobjBill.Details(Bill.Row).���� <> 0 Then
                            If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                                MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, IIf(gbytBilling = 1 Or mblnSavePrice, 1, 0), Bill.Row)) = False Then
                                mobjBill.Details(Bill.Row).���� = lngOld����
                                Call CalcMoneys(Bill.Row)
                                Bill.Text = "": Bill.TxtVisible = False
                                Cancel = True: Exit Sub
                            End If
                        End If
                        
                        Call ShowDetails(Bill.Row)
                        
                        '����������ҩ����,����Ƕ�����,���޸������Ǵ����,����Ǵ���,���޸�ͬһ����Ĵ����.��Ϊ�޶�Ϊ�в�ҩ,������������
                        For i = 1 To mobjBill.Details.Count
                            If i <> Bill.Row And mobjBill.Details(i).�շ���� = "7" And mobjBill.Details(i).�������� = mobjBill.Details(Bill.Row).�������� Then
                                If mobjBill.Details(i).�������� = 0 Or (mobjBill.Details(i).�������� <> 0 And mobjBill.Details(i).Detail.���д��� = 0) Then     '1��2�̶��Ͱ������Ĳ���
                                    mobjBill.Details(i).���� = Bill.Text
                                    Call CalcMoneys(i)
                                    Call ShowDetails(i)
                                End If
                            End If
                        Next
                                                
                        Call ShowMoney
                    Else
                        sta.Panels(2) = "������Ŀ�ĸ������ܸ��ģ�"
                        Bill.Text = mobjBill.Details(Bill.Row).����: Beep '�ָ�ԭ�и���ֵ
                    End If
                End If
            Case "����"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                     With mobjBill.Details(Bill.Row)
                         '��ҩ�������ת��
                        If .�շ���� = "7" Then Bill.Text = ConvertABCtoNUM(Bill.Text)
                        '���ֺϷ���
                        If Not IsNumeric(Bill.Text) Then
                            MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                            Bill.Text = .����: Cancel = True: Exit Sub
                        End If
                        If Val(Bill.Text) = 0 Then
                            If MsgBox("��������Ϊ�㣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Bill.Text = .����: Cancel = True: Exit Sub
                            End If
                        End If
                        'ҩƷ����С��
                        If InStr(",5,6,7,", .�շ����) > 0 Then
                            If Val(Bill.Text) - Int(Val(Bill.Text)) <> 0 And InStr(mstrPrivsOpt, ";ҩƷ����С��;") = 0 Then
                                MsgBox "��û��Ȩ������С����", vbInformation, gstrSysName
                                Bill.Text = .����: Cancel = True: Exit Sub
                            End If
                        End If
                        '�������
                        If gcurMaxMoney > 0 Then
                            If CSng(Bill.Text) * .���� * Bill.TextMatrix(Bill.Row, BillCol.����) > gcurMaxMoney Then
                                If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                    Bill.Text = .����: Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        Bill.Text = FormatEx(Bill.Text, 5)
                        If InStr(",5,6,7,", .�շ����) > 0 And gblnסԺ��λ Then
                            dblNum = Val(Bill.Text) * .���� * .Detail.סԺ��װ
                        Else
                            dblNum = Val(Bill.Text) * .����
                        End If
                            
                        '�����Ϸ��Լ��
                        If Val(Bill.Text) * .���� < 0 Then
                            'Ȩ��
                            bln�������� = True
                            If InStr(",5,6,", .�շ����) > 0 Then
                                bln�������� = (InStr(mstrPrivsOpt, ";��ҩ��������;") > 0)
                            ElseIf InStr(",7,", .�շ����) > 0 Then
                                bln�������� = (InStr(mstrPrivsOpt, ";��ҩ��������;") > 0)
                            Else
                                bln�������� = (InStr(mstrPrivsOpt, ";���Ƹ�������;") > 0)
                            End If
                        
                            If Not bln�������� Then
                                MsgBox "��û��Ȩ�����븺����", vbInformation, gstrSysName
                                Bill.Text = .����: Cancel = True: Exit Sub
                            Else
                                If .Detail.���� Then
                                    MsgBox "����ҩƷ���������븺����", vbInformation, gstrSysName
                                    Bill.Text = .����: Cancel = True: Exit Sub
                                End If
                                If mrsInfo.State = 1 Then
                                    If Not IsNull(mrsInfo!����) Then
                                        If Not MCPAR.�������� Then
                                            MsgBox "����ҽ����֧�ֶ�ҽ�����˽��и������ʣ�", vbInformation, gstrSysName
                                            Bill.Text = .����: Cancel = True: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                            
                            '���������������
                            If Not (InStr(",5,6,7,", .�շ����) > 0 And gbln���뷢ҩ) Then
                                If Not CheckNegative(mobjBill.����ID, mobjBill.��ҳID, .�շ�ϸĿID, .ִ�в���ID, dblNum, .Detail.סԺ��װ, mstrPrivsOpt) Then
                                    Bill.Text = .����: Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        'ҩƷ�����
                        If (.�շ���� = "4" And .Detail.��������) Or (InStr(",5,6,7,", .�շ����) > 0 And Not gbln���뷢ҩ) Then
                            If .Detail.���� Or .Detail.��� Then
                                '������ʱ��ҩƷ�����ֹ����
                                If .���� * Val(Bill.Text) > .Detail.��� Then
                                    If .�շ���� = "4" Then
                                        MsgBox """" & .Detail.���� & """Ϊ������ʱ����������,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    Else
                                        MsgBox """" & .Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    End If
                                    Bill.Text = .����: Cancel = True: Exit Sub
                                End If
                            Else
                                Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
                                If colStock("_" & .ִ�в���ID) <> 0 And Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus Then
                                    If .���� * Val(Bill.Text) > .Detail.��� Then
                                        If colStock("_" & .ִ�в���ID) = 1 Then
                                            If MsgBox("""" & .Detail.���� & """�ĵ�ǰ���ÿ�治����������,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                Bill.Text = .����: Cancel = True: Exit Sub
                                            End If
                                        ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                                            MsgBox """" & .Detail.���� & """�ĵ�ǰ���ÿ�治������������", vbInformation, gstrSysName
                                            Bill.Text = .����: Cancel = True: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf InStr(",5,6,7,", .�շ����) > 0 And gbln���뷢ҩ Then
                            'û��Ȩ��ʱ���̶�����ʾ��ʽ���
                            strҩ��IDs = Decode(.�շ����, "5", gstr��ҩ��, "6", gstr��ҩ��, "7", gstr��ҩ��)
                            If strҩ��IDs <> "" And .���� * Val(Bill.Text) > .Detail.��� Then
                                If gblnStock Then
                                    MsgBox "[" & .Detail.���� & "]�ĵ�ǰ���ÿ�治����������!", vbInformation, gstrSysName
                                    Bill.Text = .����: Cancel = True: Exit Sub
                                Else
                                    If MsgBox("[" & .Detail.���� & "]�ĵ�ǰ���ÿ�治����������,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        Bill.Text = .����: Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        
                        dblPreTime = .����
                        .���� = Bill.Text
                        
                        '�����������
                        If mbln����������� And Not gbln�������� Then
                            If Not CheckLimit(mobjBill, Bill.Row, gblnסԺ��λ) Then
                                .���� = dblPreTime: Bill.Text = dblPreTime
                                Cancel = True: Exit Sub
                            End If
                        End If
                        If .Detail.¼������ > 0 And dblNum > .Detail.¼������ Then
                            If MsgBox("��������γ�����¼������" & FormatEx(.Detail.¼������ / IIf(gblnסԺ��λ, .Detail.סԺ��װ, 1), 5) & ",�Ƿ����?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                                .���� = dblPreTime: Bill.Text = dblPreTime
                                Cancel = True: Exit Sub
                            End If
                        End If
                        '����ʹ������
                        If mrsInfo.State = 1 Then
                            If .Detail.Ҫ������ And Not IsNull(mrsInfo!����) And Not mrsMedAudit Is Nothing Then
                                mrsMedAudit.Filter = "��ĿID=" & .�շ�ϸĿID
                                If mrsMedAudit.RecordCount > 0 Then
                                    If Not IsNull(mrsMedAudit!��������) Then
                                        If dblNum > mrsMedAudit!�������� Then
                                            MsgBox "��������γ�������׼�Ŀ�������" & FormatEx(mrsMedAudit!�������� / IIf(gblnסԺ��λ, .Detail.סԺ��װ, 1), 5) & "��", vbInformation, gstrSysName
                                            .���� = dblPreTime: Bill.Text = dblPreTime
                                            Cancel = True: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        '���д������ܸ�������(����Ŀ���θı�,���д���������Ҳ��)
                        If .�������� <> 0 And .Detail.���д��� <> 0 Then
                            sta.Panels(2) = "����Ŀ�ǹ��д�����Ŀ,�����β��ܹ����ġ�"
                            .���� = dblPreTime: Bill.Text = dblPreTime
                            Exit Sub
                        End If
                                            
                        Call CalcMoneys(Bill.Row)
                        
                        '����������(���Ѿ�������з��õ�δ��ʾǰ)
                        If MoneyOverFlow(mobjBill) Then
                            MsgBox "�����������µ��ݽ����������ʵ�������", vbInformation, gstrSysName
                            .���� = dblPreTime
                            Call CalcMoneys(Bill.Row)
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                        
                        '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                        mrsWarn.Filter = ""
                        If mrsWarn.RecordCount > 0 And mrsInfo.State = 1 Then
                            curTotal = GetBillTotal(mobjBill)
                            If curTotal > 0 Then
                                cur��� = Val(txtʵ��.Tag)
                                '���˺�:24491
                                curItemMoney = 0
                                If mobjBill.Details(Bill.Row).�շ���� = "F" Then   '���ܴ��ڸ����������,���ֻ����ʾ
                                    curItemMoney = GetBillRowTotal(mobjBill.Details(Bill.Row).InComes)
                                End If
                                If gbln�����������۷��� Then cur��� = Val(txtʵ��.Tag) - GetPriceMoneyTotal(1, mrsInfo!����ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                                gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!���� & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), Val("" & mrsInfo!����ID), mrsInfo!���ò���, mrsWarn, cur���, mrsInfo!���ն� - mcurModiMoney, _
                                            curTotal, IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), .�շ����, .Detail.�������, mstrWarn, , gblnPrice And (gbytBilling = 0 And mstrInNO = "" Or gbytBilling = 1), curItemMoney)
                                            
                                If gbytWarn = 2 Or gbytWarn = 3 Then
                                    .���� = dblPreTime
                                    Bill.Text = ""
                                    Call CalcMoneys(Bill.Row)
                                    Cancel = True: Bill.TxtVisible = False: Exit Sub
                                ElseIf gbytBilling = 0 And gblnPrice And mstrInNO = "" Then
                                    If gbytWarn = 1 Or gbytWarn = 4 Then
                                        cmdPrice.Visible = True: cmdOK.Visible = True: Call SetButtonPlace
                                    ElseIf gbytWarn = 5 Then
                                        cmdPrice.Visible = True: cmdOK.Visible = False: Call SetButtonPlace
                                    End If
                                End If
                            End If
                        End If
                        
                        
                        If mrsInfo.State = 1 Then
                            If Not IsNull(mrsInfo!����) And MCPAR.ʵʱ��� And mobjBill.Details(Bill.Row).���� <> 0 Then
                                If gclsInsure.CheckItem(Val(mrsInfo!����), 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, IIf(gbytBilling = 1 Or mblnSavePrice, 1, 0), Bill.Row)) = False Then
                                    .���� = dblPreTime
                                    Bill.Text = ""
                                    Call CalcMoneys(Bill.Row)
                                    Cancel = True: Bill.TxtVisible = False: Exit Sub
                                End If
                            End If
                        End If
                        
                        If .���� <> 0 Then
                            If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                                MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, IIf(gbytBilling = 1 Or mblnSavePrice, 1, 0), Bill.Row)) = False Then
                                .���� = dblPreTime
                                Call CalcMoneys(Bill.Row)
                                Bill.Text = "": Bill.TxtVisible = False
                                Cancel = True: Exit Sub
                            End If
                        End If
                    End With
                        
                    Call ShowDetails(Bill.Row)
                    '��������д���������
                    For i = Bill.Row + 1 To mobjBill.Details.Count
                        If mobjBill.Details(i).�������� = Bill.Row Then
                            '28136
                            '���������ĸ���,��Ҫ���¼��еĸ������и��³ɸ���
                            With mobjBill.Details(i)
                                If .Detail.���д��� = 0 Then  '�ǹ��д���
                                    If Abs(.����) <> Abs(.Detail.��������) Then GoTo NotCalc:
                                    .���� = IIf(Val(Bill.Text) < 0, -1, 1) * .Detail.��������
                                ElseIf .Detail.���д��� = 1 Then '�̶��Ĺ��д���
                                    .���� = IIf(Val(Bill.Text) < 0, -1, 1) * IIf(.Detail.�������� = 0, 1, .Detail.��������)
                                ElseIf .Detail.���д��� = 2 Then   '�������Ĺ��д���
                                    .���� = Val(Bill.Text) * .Detail.��������
                                Else
                                     GoTo NotCalc:
                                End If
                            End With
                            
                            Call CalcMoneys(i)
                            Call ShowDetails(i)
NotCalc:
                        End If
                    Next
 
                    
                    Call ShowMoney
                ElseIf mobjBill.Details.Count >= Bill.Row Then
                    If Val(Bill.TextMatrix(Bill.Row, Bill.Col)) = 0 Then
                        If MsgBox("��������Ϊ�㣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True: Exit Sub
                        End If
                    End If
                End If
                    
                If Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus Then
                    If CheckItemHaveSub(Bill.Row) Then
                        KeyCode = 0
                        Call LocateMainItemNextRow(Bill.Row)
                    End If
                End If
            Case "����"
                If mobjBill.Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '���ֺϷ���
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    If Val(Bill.Text) < 0 Then
                        MsgBox "��Ŀ�۸�Ӧ��Ϊ������Ҫ�������ã������븺��������ʵ�֣�", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    '�������
                    If gcurMaxMoney > 0 Then
                        If Val(Bill.Text) * mobjBill.Details(Bill.Row).���� * mobjBill.Details(Bill.Row).���� > gcurMaxMoney Then
                            If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If

                    Bill.Text = FormatEx(Bill.Text, 5)
                    
                    '���û�ж�Ӧ��������Ŀ,���޷�����
                    If mobjBill.Details(Bill.Row).Detail.��� And mobjBill.Details(Bill.Row).InComes.Count > 0 Then
                        If Not (mobjBill.Details(Bill.Row).InComes(1).�ּ� = 0 And mobjBill.Details(Bill.Row).InComes(1).ԭ�� = 0) Then
                            strScope = CheckScope(mobjBill.Details(Bill.Row).InComes(1).ԭ��, mobjBill.Details(Bill.Row).InComes(1).�ּ�, CCur(Bill.Text))
                            If strScope <> "" Then
                                sta.Panels(2) = strScope
                                If Bill.TxtVisible And Len(Bill.Text) > 9 Then Bill.Text = mobjBill.Details(Bill.Row).InComes(1).��׼����
                                If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                                Cancel = True: Beep: Exit Sub
                            End If
                        End If
                        
                        dblPreMoney = mobjBill.Details(Bill.Row).InComes(1).��׼����
                        
                        mobjBill.Details(Bill.Row).InComes(1).��׼���� = Bill.Text '�����շ�ϸĿֻ�ܶ�Ӧһ��������Ŀ
                        Call CalcMoneys(Bill.Row)
                        
                        '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                        mrsWarn.Filter = ""
                        If mrsWarn.RecordCount > 0 And mrsInfo.State = 1 Then
                            curTotal = GetBillTotal(mobjBill)
                            If curTotal > 0 Then
                                cur��� = Val(txtʵ��.Tag)
                                If gbln�����������۷��� Then cur��� = Val(txtʵ��.Tag) - GetPriceMoneyTotal(1, mrsInfo!����ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                                gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!���� & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), Val("" & mrsInfo!����ID), mrsInfo!���ò���, mrsWarn, cur���, mrsInfo!���ն� - mcurModiMoney, _
                                            curTotal, IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), mobjBill.Details(Bill.Row).�շ����, mobjBill.Details(Bill.Row).Detail.�������, _
                                            mstrWarn, , gblnPrice And (gbytBilling = 0 And mstrInNO = "" Or gbytBilling = 1))
                                If gbytWarn = 2 Or gbytWarn = 3 Then
                                    mobjBill.Details(Bill.Row).InComes(1).��׼���� = dblPreMoney
                                    Bill.Text = ""
                                    Call CalcMoneys(Bill.Row)
                                    Cancel = True: Bill.TxtVisible = False: Exit Sub
                                ElseIf gbytBilling = 0 And gblnPrice And mstrInNO = "" Then
                                    If gbytWarn = 1 Or gbytWarn = 4 Then
                                        cmdPrice.Visible = True: cmdOK.Visible = True: Call SetButtonPlace
                                    ElseIf gbytWarn = 5 Then
                                        cmdPrice.Visible = True: cmdOK.Visible = False: Call SetButtonPlace
                                    End If
                                End If
                            End If
                        End If
                        
                        Call ShowDetails(Bill.Row)
                        Call ShowMoney
                    Else
                        Bill.Text = "0"
                        sta.Panels(2) = "����Ŀ�������ö�Ӧ�ķ�Ŀ�������޷�������ã�"
                        Beep
                    End If
                End If
            Case "ִ�п���"
                If mobjBill.Details.Count >= Bill.Row And Bill.ListIndex <> -1 Then
                    With mobjBill.Details(Bill.Row)
                        If .ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
                            .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                            If CheckItemHaveSub(Bill.Row) Then Call SetSubItemDept(Bill.Row) '������ڴ���,��ı��ҩƷ�е�ִ�п���
                        End If
                
                        'ҩƷ�����:��̬ҩ��,������ʱ��ҩƷҲҪ�����
                        If (.�շ���� = "4" And .Detail.��������) Or (InStr(",5,6,7,", .�շ����) > 0 And Not gbln���뷢ҩ) Then
                            If .Detail.���� Or .Detail.��� Then '������ʱ��ҩƷ��治���ֹ����
                                If .���� * .���� > .Detail.��� Then
                                    If .�շ���� = "4" Then
                                        MsgBox "[" & .Detail.���� & "]Ϊ������ʱ����������,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    Else
                                        MsgBox "[" & .Detail.���� & "]Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    End If
                                    Cancel = True
                                End If
                            Else
                                Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
                                If colStock("_" & .ִ�в���ID) <> 0 Then
                                    If .���� * .���� > .Detail.��� Then
                                        If colStock("_" & .ִ�в���ID) = 1 Then
                                            If MsgBox("[" & .Detail.���� & "]�ĵ�ǰ���ÿ�治����������,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                Cancel = True
                                            End If
                                        ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                                            MsgBox "[" & .Detail.���� & "]�ĵ�ǰ���ÿ�治������������", vbInformation, gstrSysName
                                            Cancel = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        '����������ϵ����Ч��,��ȷ��ִ�п���֮��
                        If .�շ���� = "4" And .Detail.�������� Then
                            Call CheckValidity(.�շ�ϸĿID, .ִ�в���ID, .����, False) '��ȷ������,��������
                        End If
                        
                        If CheckItemHaveSub(Bill.Row) Then
                            KeyCode = 0
                            Call LocateMainItemNextRow(Bill.Row)
                        End If
                        If mrsInfo.State = 1 Then
                            If Not IsNull(mrsInfo!����) And MCPAR.ʵʱ��� And mobjBill.Details(Bill.Row).���� <> 0 Then
                                If gclsInsure.CheckItem(Val(mrsInfo!����), 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, IIf(gbytBilling = 1 Or mblnSavePrice, 1, 0), Bill.Row)) = False Then
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        
                        If mobjBill.Details(Bill.Row).���� <> 0 Then
                            If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 0, _
                                MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, IIf(gbytBilling = 1 Or mblnSavePrice, 1, 0), Bill.Row)) = False Then
                                Bill.Text = "": Bill.TxtVisible = False
                                Cancel = True: Exit Sub
                            End If
                        End If
                    End With
                End If
        Case "��־"
        End Select
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Cancel = True
End Sub

Private Sub LocateMainItemNextRow(ByVal lngRow As Long)
    Dim i As Long
    
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�������� = lngRow Then
            If mobjBill.Details(i).Detail.���д��� = 0 Then Exit For
        End If
    Next
    
    If i <= mobjBill.Details.Count Then
        Bill.Col = BillCol.����
        Bill.Row = i: Bill.MsfObj.TopRow = i
    Else
        Call LocateNewRow
    End If
End Sub

Private Sub LocateNewRow()
    If mobjBill.Details.Count >= Bill.Rows - 1 Then
        Bill.Rows = Bill.Rows + 1
        mblnNewRow = True
        Call bill_AfterAddRow(Bill.Rows - 1)
        mblnNewRow = False
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.���
    Else
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.���
    End If
    '����:27792
    If Not Me.ActiveControl Is Bill Then
        If Bill.Active And Bill.Visible Then Bill.SetFocus
    End If
End Sub
Private Sub SetDetailtStock(ByVal lngִ�п���ID As Long, ByRef objDetail As Detail)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�������ϸ�Ŀ������
    '���ƣ����˺�
    '���ڣ�2010-07-12 14:27:51
    '˵����
    '      bug:31374
    '------------------------------------------------------------------------------------------------------------------------
    Dim strҩ��IDs As String, dblStock As Double
    
    '��ȡ���
    '�������ҩƷ������
    If InStr(1, "5,6,7,4", objDetail.���) = 0 Then Exit Sub
    If objDetail.��� = "4" And objDetail.�������� = False Then Exit Sub
    If objDetail.��� = "4" Then
        '����
        dblStock = GetStock(objDetail.ID, lngִ�п���ID, objDetail.����)
        objDetail.��� = dblStock
        Exit Sub
    End If
    
    If Not gbln���뷢ҩ Then
        dblStock = GetStock(objDetail.ID, lngִ�п���ID)
        If gblnסԺ��λ Then
            dblStock = dblStock / objDetail.סԺ��װ
        End If
        objDetail.��� = dblStock  '��¼��ǰ��ҩƷ���
        Exit Sub
    End If
    strҩ��IDs = Decode(mobjDetail.���, "5", gstr��ҩ��, "6", gstr��ҩ��, "7", gstr��ҩ��)
    If strҩ��IDs <> "" Then
        dblStock = GetMultiStock(mobjDetail.ID, strҩ��IDs)
        If gblnסԺ��λ Then
            dblStock = dblStock / mobjDetail.סԺ��װ
        End If
        mobjDetail.��� = dblStock
    End If
End Sub
Private Sub SetSubItem()
'����:�����շ���Ŀ��,���ص�ǰ�շ���Ŀ�Ĵ�����Ŀ�����ü�����,����ʾ�ڵ��ݿؼ���
'����:
'������:Bill_KeyDown��������Ŀ��
Dim i As Integer, j As Integer, lngMainRow As Long
Dim lngDoUnit As Long, lng���˿���ID As Long
Dim bln��������ۿ� As Boolean
Dim strժҪ As String
Dim dblStock As Double, strPriceGrade As String
lngMainRow = Bill.Row               '�������
If gbln��������ۿ� Then            '����������ηѱ�,����ܼ����ۿ۲�����Ч,�����ܼ���
    bln��������ۿ� = Not mobjBill.Details(lngMainRow).Detail.���ηѱ�
End If

lng���˿���ID = mobjBill.����ID
If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)

With mobjBill.Details(lngMainRow)
    Set mcolDetails = New Details
    Set mcolDetails = GetSubDetails(.�շ�ϸĿID)
    For i = 1 To mcolDetails.Count
        If mobjBill.Details.Count >= Bill.Rows - 1 Then
            Bill.Rows = Bill.Rows + 1
            mblnNewRow = True
            Call bill_AfterAddRow(Bill.Rows - 1)
            mblnNewRow = False
        End If
        Bill.TextMatrix(Bill.Rows - 1, BillCol.���) = "" '�б�Ҫ����
        
        'a.������ĿΪ��ҩƷ��Ŀ��ִ�п���
        lngDoUnit = 0
        If InStr(",4,5,6,7,", mcolDetails(i).���) = 0 Then
             If mcolDetails(i).��� = .�շ���� Or mcolDetails(i).ִ�п��� = 0 Then
                '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                '2.��������Ϊ����ȷ���ҵ�,ȱʡ������ִ�п�����ͬ��
                lngDoUnit = .ִ�в���ID
             Else
                '3.������ҩ��Ŀ��ִ�п���
                lngDoUnit = Get�շ�ִ�п���ID(mcolDetails(i).���, mcolDetails(i).ID, _
                    mcolDetails(i).ִ�п���, lng���˿���ID, Get��������ID, Get������Դ, , mobjBill.����ID)
             End If
        'b.������ĿΪҩƷ,���ĵ�ִ�п���
        Else
            lngDoUnit = Get�շ�ִ�п���ID(mcolDetails(i).���, mcolDetails(i).ID, _
                mcolDetails(i).ִ�п���, lng���˿���ID, Get��������ID, Get������Դ, .ִ�в���ID, mobjBill.����ID) '���Ĵ���ȱʡ������ִ�п�����ͬ
        End If
        
        '���»�ȡ���
        Call SetDetailtStock(lngDoUnit, mcolDetails(i))
 
                   
         '������Ŀ��Ӧ���
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!����) Then
                If InStr(",5,6,7,", mcolDetails(i).���) > 0 Then
                    strPriceGrade = mstrҩƷ�۸�ȼ�
                ElseIf mcolDetails(i).��� = "4" Then
                    strPriceGrade = mstr���ļ۸�ȼ�
                Else
                    strPriceGrade = mstr��ͨ�۸�ȼ�
                End If
                If Not CheckMediCareItem(mcolDetails(i).ID, mrsInfo!����, mcolDetails(i).����, _
                    mcolDetails(i).��� = False, , strPriceGrade) Then
                    Exit Sub
                End If
            End If
        End If
        
        Call SetDetail(mcolDetails(i), Bill.Rows - 1, lngDoUnit, Bill.Row)
        
        Call CalcMoney(Bill.Rows - 1, bln��������ۿ�)
        Call ShowDetails(Bill.Rows - 1)
        
        'CalcMoney���ȵ���GetuItemInsure���ܷ���ժҪ
        strժҪ = mobjBill.Details(Bill.Rows - 1).ժҪ
        If mrsInfo.State = 1 Then '90304
            strժҪ = gclsInsure.GetItemInfo(Val(Nvl(mrsInfo!����)), mrsInfo!����ID, mcolDetails(i).ID, strժҪ, 2)
        Else
            strժҪ = gclsInsure.GetItemInfo(0, 0, mcolDetails(i).ID, strժҪ, 2)
        End If
        mobjBill.Details(Bill.Rows - 1).ժҪ = strժҪ
    Next
    
    If bln��������ۿ� Then
        Call CalcMoney(lngMainRow, bln��������ۿ�) '�����������Ӧ����ʵ��,��Ϊ��û�м������ǰ�����ǰ������������.
        
        Call Calc��������ʵ��(lngMainRow)
    End If
    
    Call ShowMoney
End With
End Sub

Private Sub Calc��������ʵ��(ByVal lngMainRow As Long)
'����:����������ۿ�ʱ,����ָ�����������ID�ĵ�һ��������Ŀ���������ʵ�ս��
'����:  lngMainRow-������ID

Dim i As Long, j As Long
Dim cur����ǰӦ�պϼ� As Currency     '��¼�����������Ӧ�պϼ�
Dim cur���ۺ�ʵ�� As Currency


With mobjBill
    For i = lngMainRow To .Details.Count
        'If i <> lngMainRow And .Details(i).�������� <> lngMainRow Then Exit For    '��ȻĿǰ�����˲������ڴ����м������������,����һ�ŵ�����������,Ϊ�˽������ܵ�����,����ȫ��ɨ��
        
        If i = lngMainRow Or .Details(i).�������� = lngMainRow Then
            For j = 1 To .Details(i).InComes.Count
                cur����ǰӦ�պϼ� = cur����ǰӦ�պϼ� + .Details(i).InComes(j).Ӧ�ս��
            Next
        End If
    Next
   
    cur���ۺ�ʵ�� = CCur(Format(ActualMoney(.�ѱ�, .Details(lngMainRow).InComes(1).������ĿID, cur����ǰӦ�պϼ�, 0, 0, 0, 0), gstrDec))
    cur���ۺ�ʵ�� = cur���ۺ�ʵ�� - cur����ǰӦ�պϼ� + .Details(lngMainRow).InComes(1).Ӧ�ս��
    .Details(lngMainRow).InComes(1).ʵ�ս�� = Format(cur���ۺ�ʵ��, gstrDec)
    
    Call ShowDetails(lngMainRow)
End With
End Sub

Private Sub SetSubItemDept(ByVal lngRow As Long)
'����:��������ִ�п��ҵı仯,ˢ�·�ҩ�����ִ�п���

    Dim i As Long, j As Long, lng���˿���ID As Long
    
    With mobjBill
        '��ȡ���д����ִ�п�������,������ȡ(��Ϊ�����ϵĴ�����Ϣ�������޸Ĺ���)
        Set mcolDetails = GetSubDetails(.Details(lngRow).�շ�ϸĿID)
        
        lng���˿���ID = .����ID
        If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)

        For i = lngRow + 1 To .Details.Count
            If .Details(i).�������� = lngRow Then
                '������ΪҩƷ�����ĵ���Ŀ��ִ�п��Ҳ�������䶯
                If InStr(",4,5,6,7,", .Details(i).�շ����) = 0 Then
                    If .Details(i).�շ���� = .Details(lngRow).�շ���� Then
                        '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                        .Details(i).ִ�в���ID = .Details(lngRow).ִ�в���ID
                    Else
                        For j = 1 To mcolDetails.Count
                            If mcolDetails.Item(j).ID = .Details(i).Detail.ID Then
                                Exit For
                            End If
                        Next
                        If j <= mcolDetails.Count Then
                            If mcolDetails.Item(j).ִ�п��� = 0 Then
                                '2.��������Ϊ����ȷ���ҵ�,ȱʡ������ִ�п�����ͬ��
                                 .Details(i).ִ�в���ID = .Details(lngRow).ִ�в���ID
                            Else
                                '3.������ҩ��Ŀ��ִ�п���
                                .Details(i).ִ�в���ID = Get�շ�ִ�п���ID(mcolDetails(j).���, mcolDetails(j).ID, _
                                    mcolDetails(j).ִ�п���, lng���˿���ID, Get��������ID, Get������Դ, , mobjBill.����ID)
                            End If
                        End If
                    End If
                    
                    'ˢ����ʾ����ִ�п���
                    If .Details(i).ִ�в���ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .Details(i).ִ�в���ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.ִ�п���) = mrsUnit!���� & "-" & mrsUnit!����
                            Else
                                Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.Details(i).ִ�в���ID, mrsUnit)
                            End If
                        Else
                            '�������ֻ(��)��ʾ����
                            Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.Details(i).ִ�в���ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.ִ�п���) = ""
                    End If
                    
                End If
            End If
        Next
    End With
End Sub

Private Function CheckItemHaveSub(ByVal lngRow As Long) As Boolean
'���ܣ��жϵ�ǰ�е���Ŀ�Ƿ���д�����Ŀ
    Dim i As Long
    
    If mobjBill.Details.Count >= lngRow Then
        For i = lngRow + 1 To mobjBill.Details.Count
            If mobjBill.Details(i).�������� = lngRow Then
                CheckItemHaveSub = True: Exit Function
            End If
        Next
    End If
End Function

Private Sub Bill_EnterCell(Row As Long, Col As Long)
'ע��:���κ�exit sub ֮ǰ����mblncboClick = False,����,�޷�������
    
    Dim strStock As String, i As Long
    Dim strҩ��IDs As String
        
    If Not Bill.Active Then Exit Sub
    If Bill.ColData(Col) = BillColType.UnFocus Then Exit Sub
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        '����б༭����������ɫ
        Bill.SetColColor BillCol.���, &HE7CFBA  '��ȻҪ�ɰ�ɫ
        Exit Sub
    End If
    
    If mblncboEnterCell Then Exit Sub  '����ͬһ������������bill��ֵ��ѭ������,ע�����κ�exit sub ֮ǰ����mblncboClick = False
    mblncboEnterCell = True
        
    '--------------------------------------------------------------------------
    '1.�иı��������ݴ��������     mlngPreRow    ��ǰ���Ƿ�ı�
    If zlCheckBill���ڷ�ɢװ��ҩ = True Then
        '��������д��ڷ�ɢװ��,��������
        Call SetBill�в�ҩEditEnabled
        mblncboEnterCell = False
         Exit Sub
    End If
   
    If mobjBill.Details.Count >= Bill.Row And mlngPreRow <> Row Then
        With mobjBill.Details(Bill.Row)
            '��ʾ���
            If InStr(",5,6,7,", .�շ����) > 0 And .�շ�ϸĿID <> 0 Then
                If Not gbln���뷢ҩ Then
                    If gbln����ҩ�� Or gbln����ҩ�� Then
                        strStock = GetStockInfo(.�շ�ϸĿID, gbln����ҩ��, gbln����ҩ��, gblnסԺ��λ)
                        If strStock <> "" Then
                            If InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0 Then
                                sta.Panels(Pan.C2��ʾ��Ϣ) = "��" & Bill.Row & "�п��:" & strStock
                            Else
                                sta.Panels(Pan.C2��ʾ��Ϣ) = "��" & Bill.Row & "���п��."
                            End If
                        End If
                    End If
                    If strStock = "" Then
                        '���¿����ʾ
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If gblnסԺ��λ Then
                            .Detail.��� = .Detail.��� / .Detail.סԺ��װ
                        End If
                        Call ShowStock(.Detail.����, .Detail.���)
                    End If
                Else
                    strҩ��IDs = Decode(.�շ����, "5", gstr��ҩ��, "6", gstr��ҩ��, "7", gstr��ҩ��)
                    If strҩ��IDs <> "" Then
                        .Detail.��� = GetMultiStock(.�շ�ϸĿID, strҩ��IDs)
                        If gblnסԺ��λ Then
                            .Detail.��� = .Detail.��� / .Detail.סԺ��װ
                        End If
                        Call ShowStock(.Detail.����, .Detail.���)
                    Else
                        sta.Panels(2) = ""
                    End If
                End If
            ElseIf .�շ���� = "4" And .Detail.�������� And .�շ�ϸĿID <> 0 Then
                .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                Call ShowStock(.Detail.����, .Detail.���)
            Else
                sta.Panels(2) = ""
            End If
                     
            Bill.ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus)
            Bill.ColData(BillCol.��Ŀ) = BillColType.CommandButton
            
             '����Ǵ�����Ŀ������Ŀ�����,���������������Ŀ
            If CheckItemHaveSub(Row) Or .�������� > 0 Then
                Bill.ColData(BillCol.���) = BillColType.Text_UnModify
                Bill.ColData(BillCol.��Ŀ) = BillColType.Text_UnModify
            End If
            
            
            '����Ƿǵ���״̬
            If mbytInState <> 2 Then
                If .�շ���� = "7" And gblnPay Then
                    Bill.ColData(BillCol.����) = BillColType.Text
                Else
                    Bill.ColData(BillCol.����) = BillColType.UnFocus
                End If
                
                '���������������
                If .Detail.��� And InStr(",5,6,7,", .�շ����) = 0 _
                    And Not (.�շ���� = "4" And .Detail.��������) Then
                    Bill.ColData(BillCol.����) = IIf(gblnTime, BillColType.Text, BillColType.UnFocus) '����
                    Bill.ColData(BillCol.����) = BillColType.Text '���
                Else
                    Bill.ColData(BillCol.����) = BillColType.Text
                    Bill.ColData(BillCol.����) = BillColType.UnFocus
                End If
                
                If .Key = "1" Then    'ָ���˹̶�ҩ��ʱ,��������ѡ��ִ�п���
                    Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus
                Else
                    Bill.ColData(BillCol.ִ�п���) = BillColType.ComboBox
                End If
                    
                If .�շ���� = "F" Then
                    Bill.ColData(BillCol.��־) = BillColType.CheckBox
                Else
                    Bill.ColData(BillCol.��־) = BillColType.UnFocus
                End If
                
                'ֻ����һ�����,������ѡ�����
                If mblnOne Then Bill.ColData(BillCol.���) = BillColType.UnFocus
            End If
        End With
    
        '��ʾժҪ
        If mobjBill.Details(Bill.Row).ժҪ <> "" Then
            sta.Panels(2) = sta.Panels(2) & "  ժҪ:" & mobjBill.Details(Bill.Row).ժҪ
        End If
    End If
    
    '������δ�������,��ָ��е�����
    If mobjBill.Details.Count < Bill.Row Then
        Bill.ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus)  '�����,��������ʱ�ᱻ�ı�
        Bill.ColData(BillCol.��Ŀ) = BillColType.CommandButton   '��Ŀ��,��������ʱ�ᱻ�ı�
    End If
    
    
    '-----------------------------------------------------------------
    '2.�иı�������ݴ������ʾ����
    If Bill.ColData(Bill.Col) = BillColType.ComboBox Then  '���ص�ǰ�е�����������
        Call FillBillComboBox(Bill.Row, Bill.Col, True)
    End If
    
    If gbln�շ���� And Bill.TextMatrix(Row, BillCol.���) = "" And mblnOne Then
        mrsClass.Filter = "����=" & gstr�շ����
        Bill.TextMatrix(Row, BillCol.���) = mrsClass!���
        Bill.RowData(Row) = Asc(mrsClass!����)
    End If
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "���" '���������ʱ���ᶨλ�������
            SetWidth Bill.cboHwnd, 70
            If Bill.TextMatrix(Row, Col) = "" Then
                If mblnOne Then
                    mrsClass.Filter = "����=" & gstr�շ����
                    Bill.TextMatrix(Row, Col) = mrsClass!���
                    Bill.RowData(Row) = Asc(mrsClass!����)
                ElseIf Row > 1 Then
                    Bill.ListIndex = -1
                    For i = 0 To Bill.ListCount - 1
                        If InStr(Bill.List(i), Bill.TextMatrix(Row - 1, Col)) > 0 Then Bill.ListIndex = i: Exit For
                    Next
                End If
            ElseIf Row >= 1 And Bill.TextMatrix(Row, Col) <> "" Then
                For i = 0 To Bill.ListCount - 1
                    If InStr(Bill.List(i), Bill.TextMatrix(Row, Col)) > 0 Then
                        Bill.ListIndex = i: Exit For
                    End If
                Next
                If Bill.ListIndex = -1 Then
                    Bill.ListIndex = SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal Bill.TextMatrix(Row - 1, Col))
                End If
            End If
        Case "ִ�п���"
            SetWidth Bill.cboHwnd, 130
        Case "����"
            Bill.TextLen = 3
            Bill.TextMask = "0123456789" & Chr(8)
        Case "����"
            Bill.TextLen = 8
            Bill.TextMask = "0123456789." & Chr(8)
            If mobjBill.Details.Count >= Bill.Row Then
                If InStr(",5,6,7,", mobjBill.Details(Bill.Row).�շ����) > 0 Then
                    If InStr(mstrPrivsOpt, ";ҩƷ����С��;") = 0 Then
                        Bill.TextMask = Replace(Bill.TextMask, ".", "")
                    End If
                End If
                
                '�ɷ����븺��
                If Not mobjBill.Details(Bill.Row).Detail.���� Then
                    If InStr(",5,6,", mobjBill.Details(Bill.Row).�շ����) > 0 Then
                        If InStr(mstrPrivsOpt, ";��ҩ��������;") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    ElseIf InStr(",7,", mobjBill.Details(Bill.Row).�շ����) > 0 Then
                        If InStr(mstrPrivsOpt, ";��ҩ��������;") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    Else
                        If InStr(mstrPrivsOpt, ";���Ƹ�������;") > 0 Then Bill.TextMask = "-" & Bill.TextMask
                    End If
                    
                    If InStr(Bill.TextMask, "-") > 0 And mrsInfo.State = 1 Then
                        If Not IsNull(mrsInfo!����) Then
                            If Not MCPAR.�������� Then
                                Bill.TextMask = Replace(Bill.TextMask, "-", "")
                            End If
                        End If
                    End If
                End If
                
                '��ҩ�������
                If mobjBill.Details(Bill.Row).�շ���� = "7" Then
                        Bill.TextMask = Bill.TextMask & gstrABC & LCase(gstrABC)
                End If
                
                
            End If
        Case "����"
            Bill.TextLen = 10
            Bill.TextMask = "0123456789." & Chr(8)
    End Select
    
    '����,����������е����ʱ,�������л�û�п�ʼ
    If Bill.TextMatrix(Row, BillCol.��Ŀ) = "" Then
        mlngPreRow = 0
    ElseIf mobjBill.Details.Count >= Row Then
        mlngPreRow = Row
    End If
    
    mblncboEnterCell = False
End Sub

Private Sub Bill_LostFocus()
    Bill.TxtVisible = False
    Bill.CmdVisible = False
    Bill.CboVisible = False
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Bill.ToolTipText = Bill.TextMatrix(Bill.MouseRow, Bill.MouseCol)
End Sub

Private Sub cboBaby_Click()
    mobjBill.Ӥ���� = cboBaby.ItemData(cboBaby.ListIndex)
End Sub

Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�ѱ�_Click()
    If cbo�ѱ�.ListIndex <> -1 And Not mobjBill Is Nothing And Bill.Active Then
        If mobjBill.�ѱ� <> zlStr.NeedName(cbo�ѱ�.Text) And Not mbln������۸� Then
            mobjBill.�ѱ� = zlStr.NeedName(cbo�ѱ�.Text)
            
            If mbytInState = 0 And mobjBill.Details.Count > 0 Then
                '���¼���۸�
                Call CalcMoneys
                Call ShowDetails
                Call ShowMoney
            End If
        End If
    End If
End Sub

Private Sub SetDefaultDoctor()
'����:����ȱʡ������
    If cbo������.ListCount = 0 Then Exit Sub
    
    If cbo������.ListCount = 1 Then
        cbo������.ListIndex = 0
    Else
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!סԺҽʦ) Then
                Call zlControl.CboSetIndex(cbo������.hWnd, cbo.FindIndex(cbo������, mrsInfo!סԺҽʦ, True))
            End If
        End If
    End If
End Sub

Private Sub cbo��������_Click()
    Dim i As Long, lng��������ID As Long
    
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If cbo��������.ListIndex <> -1 Then lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    If mobjBill.��������ID = lng��������ID Then Exit Sub
    
    '����:
    
    If mrs��ҩ����.RecordCount <> 0 Then
        For i = 0 To cboDrawDept.ListCount - 1
             If cboDrawDept.ItemData(i) = lng��������ID Then
                mobjBill.��ҩ����ID = lng��������ID
                cboDrawDept.ListIndex = i: Exit For
             End If
        Next
    End If
    
    mobjBill.��������ID = lng��������ID
        
    '��������ȷ��ҽ��
    If Not gblnFromDr Then
        If cbo��������.ListIndex <> -1 Then
            If gbln������ Then
                Call FillDoctor(cbo������, mrs������)
            Else
                Call FillDoctor(cbo������, mrs������, lng��������ID)
            End If
            Call SetDefaultDoctor
        Else
            cbo������.Clear
        End If
        Call cbo������_Click
    End If
    
    
    '�������������Ŀ��ִ�п���
    If cbo��������.ListIndex <> -1 And cbo��������.Visible Then
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                '�������շ���Ŀ
                If InStr(",4,5,6,7,", .Detail.���) = 0 And .Detail.ִ�п��� = 6 Then '6-�����˿���
                    .ִ�в���ID = cbo��������.ItemData(cbo��������.ListIndex)
                    'ˢ����ʾ����ִ�п���
                    If i <= Bill.Rows - 1 And .ִ�в���ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .ִ�в���ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.ִ�п���) = mrsUnit!���� & "-" & mrsUnit!����
                            Else
                                Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.ִ�в���ID, mrsUnit)
                            End If
                        Else
                            '�������ֻ(��)��ʾ����
                            Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.ִ�в���ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.ִ�п���) = ""
                    End If
                End If
            End With
        Next
    End If
End Sub

Private Sub cbo������_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub cbo������_Validate(Cancel As Boolean)
    If cbo������.Text <> "" Then
        If cbo.FindIndex(cbo������, zlStr.NeedName(cbo������.Text), True) = -1 Then cbo������.ListIndex = -1: cbo������.Text = ""
    End If
    If cbo������.Text = "" Then Call cbo������_KeyPress(vbKeyReturn)
    '����������ȷ��������ʱ,���ܴ�ʱ��ѡ������,��ȥ�����������Һ�����ѡ
    If gblnFromDr And gbln������ And cbo������.ListIndex = -1 And txtPatient.Text <> "" Then Cancel = True
End Sub

Private Sub cbo������_Click()
    Dim lng������ID As Long
    
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If mobjBill.������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text)) Then Exit Sub
    
    mobjBill.������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text))
    If gblnFromDr Then
        If cbo������.ListIndex <> -1 Then
            lng������ID = cbo������.ItemData(cbo������.ListIndex)
            
            Call FillDept(cbo��������, mrs��������, mrs������, mstrPrivs, mbytUseType, mlngDeptID, lng������ID)
            Call SetDefaultDept(cbo��������, mrs��������, mrs������, lng������ID)
        Else
            cbo��������.Clear
        End If
        Call cbo��������_Click
    End If
                        
    '��ʿ���
    If Bill.Active Then
        If mobjBill.Details.Count < Bill.Rows - 1 And Bill.Row = Bill.Rows - 1 _
            And Bill.RowData(Bill.Rows - 1) <> 0 Then
            '�����Ч����
            Bill.TextMatrix(Bill.Rows - 1, BillCol.���) = ""
            Bill.RowData(Bill.Rows - 1) = 0
        ElseIf Bill.Col = BillCol.��� Then
            Call Bill_EnterCell(Bill.Row, Bill.Col) 'ˢ��
        End If
    End If
    
    '��ʿ���:�жϷǷ�����
    If CheckInhibitiveByNurse(mobjBill, mrs������) Then
        MsgBox "��ʿֻ���������Ƽ�������Ŀ,�������д����������͵���Ŀ��", vbInformation, gstrSysName
    End If
End Sub

Private Sub cbo������_KeyDown(KeyCode As Integer, Shift As Integer)
    If cbo������.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo������.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub

Private Sub cboҽ�Ƹ���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cboҽ�Ƹ���.ListIndex <> -1 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub


Private Sub chkCancel_Click()
    Dim i As Long
    
    '�������
    chk����.Value = 0: chk����.Visible = False
    sta.Panels(2).Text = ""
    
    mstrInNO = ""
    Call NewBill
    Call ClearRows
    Call Bill.ClearBill: Call SetColNum
    Call ClearMoney
    
    Bill.AllowAddRow = (chkCancel.Value = 0)
    
    Call SetDrawDrugDeptVisible
    
    If chkCancel.Value = 1 Then
        chkCancel.ForeColor = &HFF&
        
        chkIn.Enabled = False
        
        fraInfo.Enabled = False
        fraUnit.Enabled = False
        
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = BillColType.Text_UnModify
        Next
        Call ShowDeleteCol(True)
        Bill.SetColColor BillCol.���, &HE7CFBA  '��ȻҪ�ɰ�ɫ
        Bill.Active = True
        
        If InStr(mstrPrivsOpt, ";ҽ����ѯ;") = 0 Then cbo������.Visible = False: lbl������.Visible = False
        Call SetDisible
        cmd�䷽.Enabled = False
        
        fraDrawDept.Enabled = False
        
        cboNO.Locked = False
        cboNO.SetFocus
    Else
        chkCancel.ForeColor = 0
        
        If InStr(mstrPrivsOpt, ";ҽ����ѯ;") = 0 Then cbo������.Visible = True: lbl������.Visible = True
        Call cbo��������_Click
        
        
        If gbytBilling = 2 Then  '���ʱ
            Call SetDisible
            cboNO.Locked = False
        Else
            Call SetDisible(True)
            cmd�䷽.Enabled = True
        End If
        
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = marrColData(i)
        Next
        Call ShowDeleteCol(False)
        Bill.SetColColor BillCol.���, &HE7CFBA  '��ȻҪ�ɰ�ɫ
        
        If gbytBilling = 2 Then
            Bill.Active = False
            cboNO.Locked = False
            cboNO.SetFocus
        Else
            fraInfo.Enabled = True
            fraUnit.Enabled = True
            fraDrawDept.Enabled = False
                    
            fraAppend.Enabled = True
            Bill.Active = True
            cboNO.Locked = True
            chkIn.Enabled = True
            If mbytUseType = 1 And mlng����ID > 0 Then
                txtPatient.Text = "-" & mlng����ID
                Call txtPatient_KeyPress(13)
                Bill.SetFocus
            Else
               If txtPatient.Enabled Then txtPatient.SetFocus
            End If
        End If
    End If
    

End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�Ӱ�_Click()
    If mbytInState = 1 Or chkCancel.Value = Checked Or gbytBilling = 2 Then Exit Sub
    If mbytInState = 2 Then Exit Sub
    If Not chk�Ӱ�.Visible Then Exit Sub
    
    Dim blnAdd As Boolean
    
    blnAdd = OverTime(zlDatabase.Currentdate)
    If chk�Ӱ�.Value = Unchecked And blnAdd Then
        If MsgBox("��ǰ���ڼӰ�ʱ�䷶Χ��,Ҫȡ���Ӱ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = Checked
        End If
    End If
    If chk�Ӱ�.Value = Checked And Not blnAdd Then
        If MsgBox("��ǰ�����ڼӰ�ʱ�䷶Χ��,Ҫִ�мӰ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = Unchecked
        End If
    End If
    mobjBill.�Ӱ��־ = IIf(chk�Ӱ�.Value = Checked, 1, 0)
    
    '���¼���۸�
    If Not mobjBill.Details.Count = 0 Then
        Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
    End If
End Sub

Private Sub chk�Ӱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    '33744
    If mbln���� And mobjBill.Details.Count = 0 Then
        Unload Me: Exit Sub
    End If
    
    If (mobjBill.Details.Count > 0 Or txtPatient.Text <> "") And Bill.Active And mbytInState = 0 And mstrInNO = "" And Not mblnCopyBill Then
    
        If MsgBox("ȷʵҪ�����ǰ�����е�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
       
        '�������
        chk����.Value = 0: chk����.Visible = False
        If chkCancel.Value = 1 Then '�˾ݵ�״̬
            Call ClearRows: Call Bill.ClearBill
            Call SetColNum: Call ClearMoney
            chkCancel.Value = Unchecked
            Call NewBill: Call SetDisible(True)
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        ElseIf Bill.Active Then '�������뵥��״̬'(����������²��˵���)
            mstrInNO = ""
            If Not mbln���� Then
                txtPatient.Text = "": txtOld.Text = ""
                txt����.Text = "": txtסԺ��.Text = ""
            End If
            txtʵ��.Text = gstrDec: txtӦ��.Text = gstrDec
            
            Call ClearRows: Call Bill.ClearBill
            Call SetColNum: Call ClearMoney
            Call NewBill(IIf(mbln����, False, True))
            If mbln���� Then
                txtPatient.Tag = "-" & mrsInfo!����ID
                With mobjBill
                    .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
                    .��ҳID = IIf(mbln���� And mlng��ҳID <> 0, mlng��ҳID, IIf(IsNull(mrsInfo!��ҳID), 0, mrsInfo!��ҳID))
                    
                    .����ID = IIf(mbln���� And mlngUnitID <> 0, mlngUnitID, Val(Nvl(mrsInfo!����ID)))
                    .����ID = IIf(mbln���� And mlngDeptID <> 0, mlngDeptID, Val(Nvl(mrsInfo!����ID)))
                    
                    .���� = "" & mrsInfo!����
                    .��ʶ�� = IIf(IsNull(mrsInfo!סԺ��), 0, mrsInfo!סԺ��)
                    .���� = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
                    .�Ա� = IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
                    .���� = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
                    .�ѱ� = IIf(IsNull(mrsInfo!�ѱ�), "", mrsInfo!�ѱ�)
                    .Ӥ���� = cboBaby.ItemData(cboBaby.ListIndex)
                    .������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text))
                End With
                 Bill.SetFocus
                
            End If
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        ElseIf Not Bill.Active Then '��ȡ���۵�����״̬
            Call ClearRows: Call Bill.ClearBill
            Call SetColNum: Call ClearMoney
            Call NewBill: Call SetDisible(True)
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        End If
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Function CheckBillNegative() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥��������ĸ����������˻ؿ����Ƿ���ȷ
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-12-29 12:13:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim strItems As String, str���� As String
    Dim str��λ As String, dbl���� As Double, dbl���κϼ� As Double
    Dim dbl�ѽ����� As Double
    
    CheckBillNegative = True
    If mobjBill.����ID = 0 Then Exit Function
    '����:26951
    If InStr(1, mstrPrivsOpt, ";�������ʲ���鷢����Ŀ;") > 0 Then
        '���ڸ�������ʱ����鱾��סԺ��������Ŀ����,�д�Ȩ��,����¼�벡��δ�������ķ�����Ŀ���г���,�����鱾��סԺ��������Ŀ�������ܳ���
        CheckBillNegative = True: Exit Function
    End If
    
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).���� < 0 And mobjBill.Details(i).ִ�в���ID <> 0 Then strItems = strItems & "," & mobjBill.Details(i).�շ�ϸĿID
    Next
    If strItems = "" Then Exit Function
    strItems = Mid(strItems, 2)
    strSQL = " " & _
    "     Select A.�շ�ϸĿid, A.ִ�в���id, Nvl(Sum(Decode(A.��¼����, 2, 1, 3, 1, 0) * Nvl(A.����, 1) * A.����), 0) As ����, " & _
    "            Sum(Decode(nvl(Mod(M.��¼״̬, 3),1), 0, 1, 1, 1, -1) * Decode(����id, Null, 0, 1) * Nvl(����, 1) * ����) As �������� " & _
    "     From סԺ���ü�¼ A, ���˽��ʼ�¼ M " & _
    "     Where  A.����id = M.ID(+)  And A.���ʷ���=1 And A.�۸񸸺� Is Null   " & IIf(gbytBilling = 0, " And A.��¼״̬<>0", "") & _
    "            And A.����id = [1] And A.��ҳid = [2]    " & _
    "            And Instr(',' ||[3]|| ',', ',' || �շ�ϸĿid || ',') > 0   " & IIf(mstrInNO <> "", " And NO<>[4]", "") & _
    "     Group By �շ�ϸĿid, ִ�в���id"
    'strSQL = _
    " Select �շ�ϸĿID,ִ�в���ID,Sum(Nvl(����,1)*����) as ����," & _
    "           Sum(decode(����ID,NULL,0,1)* Nvl(����,1)*����) as ��������  " & _
    " From סԺ���ü�¼" & _
    " Where  And �۸񸸺� is NULL And ����ID=[1] And ��ҳID=[2] And Instr(','||[3]||',',','||�շ�ϸĿID||',')>0" & _
            IIf(gbytBilling = 0, " And ��¼״̬<>0", "") & _
            IIf(mstrInNO <> "", " And NO<>[4]", "") & _
    " Group by �շ�ϸĿID,ִ�в���ID"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.����ID, mobjBill.��ҳID, strItems, mstrInNO)
    
    For i = 1 To mobjBill.Details.Count
        With mobjBill.Details(i)
            If .���� < 0 And .ִ�в���ID <> 0 Then
                rsTmp.Filter = "�շ�ϸĿID=" & .�շ�ϸĿID & " And ִ�в���ID=" & .ִ�в���ID
                If Not rsTmp.EOF Then
                    If InStr(",5,6,7,", .�շ����) > 0 And gblnסԺ��λ Then
                        str��λ = .Detail.סԺ��λ
                        dbl���� = Nvl(rsTmp!����, 0) / .Detail.סԺ��װ
                        dbl���κϼ� = Abs(.����) * .����
                        dbl�ѽ����� = Nvl(rsTmp!��������, 0) / .Detail.סԺ��װ
                    Else
                        str��λ = .Detail.���㵥λ
                        dbl���� = Nvl(rsTmp!����, 0)
                        dbl���κϼ� = Abs(.����) * .����
                        dbl�ѽ����� = Nvl(rsTmp!��������, 0)
                        
                        '���ܴ���������ͬ�ļ�¼
                        '����:29412
                        For j = i + 1 To mobjBill.Details.Count
                             If .�շ�ϸĿID = mobjBill.Details(j).�շ�ϸĿID _
                                And mobjBill.Details(j).���� < 0 And mobjBill.Details(j).ִ�в���ID = .ִ�в���ID Then
                                dbl���κϼ� = dbl���κϼ� + Abs(.����) * .����
                             End If
                        Next
                    End If
                    
                    If dbl���κϼ� > dbl���� - dbl�ѽ����� Then
                            Select Case gbytBillOpt '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
                            Case 0  '����
                                If dbl���κϼ� > dbl���� Then
                                        str���� = GET��������(.ִ�в���ID, mrsUnit)
                                        MsgBox "�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                            " ���ڿ��������� " & FormatEx(dbl����, 5) & str��λ & "��", vbInformation, gstrSysName
                                        CheckBillNegative = False: Exit Function
                                End If
                            Case 1   '����
                                str���� = GET��������(.ִ�в���ID, mrsUnit)
                                If dbl���κϼ� > dbl���� Then
                                        MsgBox "�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                            " ���ڿ��������� " & FormatEx(dbl����, 5) & str��λ & "��", vbInformation, gstrSysName
                                        CheckBillNegative = False: Exit Function
                                End If
                                
                                If MsgBox("�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                    " �а������ѽᲿ��(δ��:" & FormatEx(dbl���� - dbl�ѽ�����, 5) & str��λ & "; �ѽ�:" & FormatEx(dbl�ѽ�����, 5) & str��λ & ") ��" & vbCrLf & _
                                    " �Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                    CheckBillNegative = False: Exit Function
                                End If
                            Case 2   '��ֹ
                                str���� = GET��������(.ִ�в���ID, mrsUnit)
                                MsgBox "�� " & i & " ��[" & .Detail.���� & "]�˻�" & str���� & "������ " & FormatEx(dbl���κϼ�, 5) & str��λ & _
                                    " ���ڿ��������� " & FormatEx(dbl���� - dbl�ѽ�����, 5) & str��λ & "��", vbInformation, gstrSysName
                                    CheckBillNegative = False: Exit Function
                            End Select
                    End If
                Else
                    MsgBox "�� " & i & " ��[" & .Detail.���� & "]����������Ϊ�㣬�����������", vbInformation, gstrSysName
                    CheckBillNegative = False: Exit Function
                End If
            End If
        End With
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function CheckMainOperation() As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������������(�����������Ҫ����,�����ڸ�������,���ֹ
    '���:
    '����:lngRow-���ظ�����������
    '����:������������û�����븽������,����true,���򷵻�False
    '����:
    '�޸�:���˺�(�˺�ʱ,���Ӷ�λ����),���Ӳ���;strBackNo
    '����:2009/7/10
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCount As Long, lngRow As Long   'ָ����
    Dim i As Long
    
    lngCount = 0
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�շ���� = "F" Then
           If mobjBill.Details(i).���ӱ�־ = 0 Then CheckMainOperation = True: Exit Function     '������Ҫ����,�򲻼��,ֱ�ӷ���true
           lngCount = lngCount + 1  '��ʾ��������
           If lngRow <= 0 Then lngRow = i
        End If
    Next
    If lngCount <> 0 Then
          MsgBox "�����в�����Ҫ����,�����ڸ�������,���飡", vbInformation, gstrSysName
          If Bill.Rows > lngRow Then Bill.Row = lngRow
          If Bill.Visible Then Bill.SetFocus
          Exit Function
    End If
    CheckMainOperation = True
End Function

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset, rsFeeItem As ADODB.Recordset
    Dim strInfo As String, strSQL As String, strTmp As String, strRows As String, str���ܺ� As String
    Dim strAddDate As String '���ʷ���,�Զ���ҩ,���ϵ�ʱ��
    Dim i As Long, j As Long, lng����ID As Long
    Dim cur���ն� As Currency, bln�������� As Boolean
    Dim curTotal As Currency, intInsure As Integer, cur��� As Currency, dbl���� As Double
    Dim dblTotal As Double, Curdate As Date, blnTrans As Boolean
    Dim colStock As Collection
    Dim arrSMSQL As Variant, str��������IDs As String, str������s As String
    Dim cllPro As Collection
    Dim rsItems As ADODB.Recordset
    
    '���ʹ���
    If mbytInState = 3 Or (mbytInState = 0 And chkCancel.Visible And chkCancel.Value = 1) Then
        If mbytInState = 0 And mstrInNO = "" Then
            MsgBox "û�ж�ȡ��������,�������ʣ�", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If
        
        For i = 1 To Bill.Rows - 1
            If Bill.TextMatrix(i, Bill.Cols - 1) = "��" And Bill.RowData(i) > 0 Then
                strRows = strRows & "," & Bill.RowData(i)
            End If
        Next
        If strRows = "" Then
            MsgBox "������ѡ��һ��Ҫ���ʵķ��ã�", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        If zlCheckIsExistsApplied(mstrInNO, strRows, str��������IDs, str������s) Then
            '����:47416
            If MsgBox("ע��:" & vbCrLf & "    ����" & mstrInNO & "�д����������ʵ���Ŀ,���ʺ�,�����Զ�ȡ��" & vbCrLf & "�����˵�������Ŀ,�Ƿ��������?" & vbCrLf & "����������: " & str������s, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        '������ѡ����
        strRows = Mid(strRows, 2)
        i = GetBillRows(mstrInNO, mbytNOType)
        If UBound(Split(strRows, ",")) + 1 = i Then strRows = ""
                
        If strRows <> "" And InStr(1, mstrPrivsOpt, ";��������;") = 0 Then
            MsgBox "��û�в������ʵ�Ȩ�ޣ�ֻ�ܶԸõ���ȫ�����ʣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        'ҽ�����������ϴ�(ע���ж�˳��)
        If gbytBilling = 0 Then '��������ʱ����
            intInsure = BillExistInsure(mstrInNO, mstrTime, , mbytNOType) '�ж��Ƿ�ҽ�����˼ǵ���
            If intInsure > 0 Then
                MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, , intInsure)
                MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure)
                
                If MCPAR.���������ϴ� Then
                    'ȥ����ҽ������ƥ����
                    If Not gclsInsure.GetCapability(support�����ݳ�������, , intInsure) And strRows <> "" Then '���ܲ�������
                        MsgBox "��Ϊҽ��������Ҫ,�õ����е���Ŀ����ȫ�����ʣ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        If zlPatiIS�����ѱ�Ŀ(mlngBill����ID, mlngBill��ҳID) = True Then     '����:28725
            Exit Sub
        End If
        If zlIsAllowFeeChange(mlngBill����ID, mlngBill��ҳID) = False Then
            Exit Sub
        End If

        Set rsFeeItem = GetNOFeeItem(mstrInNO, mbytNOType, strRows)
        Set rsTmp = GetPatientFeeItemTotal(mlngBill����ID, mlngBill��ҳID, mstrInNO)
        If rsFeeItem.RecordCount > 0 And rsTmp.RecordCount > 0 Then
            For i = 1 To Bill.Rows - 1
                rsFeeItem.Filter = "���=" & Bill.RowData(i)
                If rsFeeItem.RecordCount > 0 Then
                    If Not (InStr(",5,6,7,", rsFeeItem!�շ����) > 0 And gbln���뷢ҩ) Then
                        rsTmp.Filter = "�շ�ϸĿid=" & rsFeeItem!�շ�ϸĿID & " And ִ�в���id=" & rsFeeItem!ִ�в���ID
                        If rsTmp.RecordCount > 0 Then
                            If Bill.TextMatrix(i, BillCol.����) * Bill.TextMatrix(i, BillCol.����) > rsTmp!���� Then
                                MsgBox "��" & i & "�������������ڼ�������" & rsTmp!���� & "��", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        Else
                            MsgBox "��" & i & "�п����ʵ�����Ϊ�㡣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
            Next
        End If
        
         '����:47416
        Set cllPro = New Collection
        If str��������IDs <> "" Then
            strSQL = "zl_���˷�������_Delete('" & str��������IDs & "')"
            zlAddArray cllPro, strSQL
        End If
        strSQL = "zl_סԺ���ʼ�¼_DELETE('" & mstrInNO & "','" & strRows & "','" & UserInfo.��� & "','" & UserInfo.���� & "'," & mbytNOType & ")"
        zlAddArray cllPro, strSQL
        cmdOK.Enabled = False
        On Error GoTo errH
            blnTrans = True
            zlExecuteProcedureArrAy cllPro, Me.Caption, True
            'ҽ�����������ϴ�
            If gbytBilling = 0 And intInsure <> 0 Then
                If MCPAR.���������ϴ� And Not MCPAR.������ɺ��ϴ� Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, mbytNOType, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Sub
                    End If
                End If
            End If
        gcnOracle.CommitTrans: blnTrans = False
        
        'ҽ�����������ϴ�
        If gbytBilling = 0 And intInsure <> 0 Then
            If MCPAR.���������ϴ� And MCPAR.������ɺ��ϴ� Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, mbytNOType, 2, "", , intInsure) Then
                    MsgBox "����""" & mstrInNO & """������������ҽ������ʧ�ܣ��õ��������ʡ�", vbInformation, gstrSysName
                End If
            End If
        End If
        
        cmdOK.Enabled = True
        On Error GoTo 0
        
        If mbytInState = 0 Then
            txtPreNO.Text = mstrInNO
            mstrInNO = "": cboNO.Text = ""
            txtPatient.Text = "": txtOld.Text = ""
            txtʵ��.Text = gstrDec: txtӦ��.Text = gstrDec
            Call ClearRows: Call Bill.ClearBill: Call SetColNum
            Call ClearMoney: Call NewBill
            Call SetMoneyList
            
            chkCancel.Value = 0
            
            If gbytBilling = 2 Then
                cboNO.SetFocus
            Else
                txtPatient.SetFocus
            End If
        Else
           gblnOK = True: Unload Me: Exit Sub
        End If
        
    ElseIf mbytInState = 2 Then
        If Not IsDate(txtDate.Text) Then
            MsgBox "������Ϸ��ķ���ʱ�䣡", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        strInfo = Check����ʱ��(CDate(txtDate.Text), cboNO.Text, IIf(mbln����, mlng��ҳID, 0))
        If strInfo <> "" Then
            MsgBox strInfo, vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        If Not SaveModi() Then Exit Sub
        gblnOK = True: Unload Me: Exit Sub
    ElseIf Bill.Active And chkCancel.Value = 0 Then '�������뵥��״̬
        If mrsInfo.State = adStateClosed Then
            MsgBox "û�з��ֲ�����Ϣ,��ȷ��������Ϣ��", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Sub
        End If
        If cbo�ѱ�.ListIndex = -1 Or mobjBill.�ѱ� = "" Then
            MsgBox "��ѡ���˷ѱ�", vbInformation, gstrSysName
            cbo�ѱ�.SetFocus: Exit Sub
        End If
        If mobjBill.Details.Count = 0 Then
            MsgBox "������û���κ�����,����ȷ���뵥�����ݣ�", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        i = Checkִ�п���
        If i <> 0 Then
            MsgBox "�����е� " & i & " ����Ŀû��ָ��ִ�п��ң�", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        If mobjBill.��������ID = 0 Then
            MsgBox "��ȷ���������ң�", vbInformation, gstrSysName
            cbo��������.SetFocus
            Exit Sub
        End If

        
        If mobjBill.������ = "" And gbln������ Then
            MsgBox "�����뿪���ˣ�", vbInformation, gstrSysName
            cbo������.SetFocus: Exit Sub
        End If
        
        '��ʿ���:�жϷǷ�����
        If CheckInhibitiveByNurse(mobjBill, mrs������) Then
            MsgBox "��ʿֻ���������Ƽ�������Ŀ,�������д����������͵���Ŀ��", vbInformation, gstrSysName
            Exit Sub
        End If
                
        '����ʱ����
        If Not IsDate(txtDate.Text) Then
            MsgBox "��������ȷ�ķ������ڣ�", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        '��鷢��ʱ��
        strInfo = Check����ʱ��(CDate(txtDate.Text), mrsInfo!����ID, mlng��ҳID)
        If strInfo <> "" Then
            MsgBox strInfo, vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        '33744
        If mbln���� Then
            If txtDate.Text > mstr���ת��ʱ�� And mstr���ת��ʱ�� <> "" Then
                MsgBox "ע��:" & vbCrLf & "    �ò��˲�¼�ķ���ʱ�䳬�������ת����ʱ��(" & mstr���ת��ʱ�� & "),���ܽ��в��Ѳ���!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
            If cbo��������.ItemData(cbo��������.ListIndex) <> mlngDeptID Then
                MsgBox "ע��:" & vbCrLf & "    �������Ҳ��ǲ���ת�ƵĿ���,���ܽ��в��Ѳ���!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                Exit Sub
            End If
        End If
        
        '��Ժǿ�Ƽ���Ȩ�޼��
        If Not PatiCanBilling(mrsInfo!����ID, Nvl(mrsInfo!��ҳID, 0), mstrPrivsOpt) Then Exit Sub
        If zlIsAllowFeeChange(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID))) = False Then
            Exit Sub
        End If
        If zlPatiIS�����ѱ�Ŀ(mrsInfo!����ID, Nvl(mrsInfo!��ҳID, 0)) = True Then     '����:28725
            Exit Sub
        End If
        '����ʱ����
        If Not IsNull(mrsInfo!��Ժ����) Then
            If Format(txtDate.Text, txtDate.Format) > Format(mrsInfo!��Ժ����, txtDate.Format) Then
                MsgBox "ǿ�ƶԳ�Ժ���˼���ʱ������ʱ�䲻�ܴ��ڲ��˳�Ժʱ��:" & Format(mrsInfo!��Ժ����, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        If Not IsNull(mrsInfo!����) And Not IsNull(mrsInfo!��Ժ����) Then
            If Format(txtDate.Text, txtDate.Format) < Format(mrsInfo!��Ժ����, txtDate.Format) Then
                MsgBox "���õķ���ʱ�䲻��С��ҽ�����˵���Ժʱ��:" & Format(mrsInfo!��Ժ����, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        
        '�Ƿ���
        dbl���� = 0
        For i = 1 To mobjBill.Details.Count
            '27467,52828
            If mobjBill.Details(i).���� <> 0 And dbl���� = 0 Then
                dbl���� = mobjBill.Details(i).����
            End If
            If mobjBill.Details(i).�շ�ϸĿID = 0 Then
                MsgBox "�����е� " & i & " ��û����ȷ��������,��������ɾ�����У�", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
            ElseIf InStr(1, ",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
                '�ռ�ҩƷ�ķ�ҩҩ��
                strTmp = strTmp & "," & mobjBill.Details(i).�շ�ϸĿID
            End If
        Next
        '27467,52828
        If mbytInState = 0 And FormatEx(dbl����, 7) = 0 Then
            MsgBox "����������Ҫ��һ����Ϊ�������,���飡", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        '���ҩƷ�ķ�ҩҩ����Ӧ�ķ������(�洢�ⷿ)
        If strTmp <> "" And Not gbln���뷢ҩ Then
            strTmp = Mid(strTmp, 2)
            Set rsTmp = GetServiceDept(strTmp)
            If Not rsTmp Is Nothing Then
                strTmp = ""
                For i = 1 To mobjBill.Details.Count
                    If InStr(1, ",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
                        strInfo = mobjBill.Details(i).�շ�ϸĿID
                        '�ȼ���Ƿ�������Ĵ洢�ⷿ
                        rsTmp.Filter = "�շ�ϸĿID=" & strInfo & " And ִ�п���id=" & mobjBill.Details(i).ִ�в���ID
                        If rsTmp.RecordCount = 0 Then
                            strTmp = strTmp & "," & i
                        Else
                            '�ټ���Ƿ�������ķ������(û�����÷�����ҵ�,��������IDΪ��)
                            rsTmp.Filter = "(" & rsTmp.Filter & " And ��������ID=" & mobjBill.����ID & ") Or (" & rsTmp.Filter & " And ��������ID=0)"
                            If rsTmp.RecordCount = 0 Then
                                strTmp = strTmp & "," & i
                            End If
                        End If
                    End If
                Next
                If strTmp <> "" Then
                    strTmp = Mid(strTmp, 2)
                    MsgBox "����,��" & strTmp & "��ҩƷ�Ƿ�Υ�����¹���:" & vbCrLf & vbCrLf & _
                        "A.ѡ���ִ�п��Ҳ���ҩƷ�Ĵ洢�ⷿ" & vbCrLf & _
                        "B.���˿���[" & GET��������(mobjBill.����ID, mrs��������) & "]������ҩƷ�ڴ˴洢�ⷿ�ķ������.", _
                        vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        
        'ҽ���������ʼ��    ��Ϊ����Ա�������䵥��,��ȷ������,����Ҫ�ټ��һ��(�˴������ж�Ȩ��,��Ϊ��Ȩ�޲ſ����Ǹ���)
        If InStr(mstrPrivsOpt, ";��������;") > 0 Then    '����������һ�ָ�������Ȩ��,�ſ����Ǹ���
            If Not IsNull(mrsInfo!����) Then
                If Not MCPAR.�������� Then
                    For i = 1 To mobjBill.Details.Count
                        If mobjBill.Details(i).���� * mobjBill.Details(i).���� < 0 Then
                                MsgBox "�����е� " & i & " ���Ǹ���,����ҽ����֧�ָ������ʣ�", vbInformation, gstrSysName
                                Bill.SetFocus: Exit Sub
                        End If
                    Next
                End If
            End If
        End If
        
        If Not IsNull(mrsInfo!����) And MCPAR.ʵʱ��� Then
            If gclsInsure.CheckItem(Val(mrsInfo!����), 1, 2, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, IIf(gbytBilling = 1 Or mblnSavePrice, 1, 0))) = False Then
                Bill.SetFocus: Exit Sub
            End If
        End If
        '����ְ����
        If cboҽ�Ƹ���.ListIndex <> -1 Then
            'ҽ���򹫷Ѳ���
            '����:45605
            If zlIsCheckMedicinePayMode(zlStr.NeedName(cboҽ�Ƹ���)) Then
                i = CheckDuty(, False)
                If i > 0 Then
                    Bill.Row = i: Bill.MsfObj.TopRow = i
                    Bill.Col = BillCol.��Ŀ: Bill.SetFocus
                    Exit Sub
                End If
            End If
        End If
        '���в�����Ŀ
        i = CheckDuty(, True)
        If i > 0 Then
            Bill.Row = i: Bill.MsfObj.TopRow = i
            Bill.Col = BillCol.��Ŀ: Bill.SetFocus
            Exit Sub
        End If
        
        'Ҫ������,ҽ��������Ŀ�Ƿ��������,����ʱ�Ѽ�飬����ʱ�ټ������Ϊ��
        '1.���䵥����ȷ��ҽ����ݣ�2.�������������ʱֻ��������3.���뵥��ʱδ���,4.ͨ���䷽����ʱδ���
        If Not IsNull(mrsInfo!����) And Not mrsMedAudit Is Nothing Then
            If Not CheckExamine(mobjBill.Details, mrsMedAudit, mrsInfo!����) Then Exit Sub
        End If
        
        '�������ò��˲�������
        For i = 1 To mobjBill.Details.Count
            If InStr(",5,6,7,", mobjBill.Details(i).�շ����) = 0 Then
                If CheckItemHaveSub(i) Then
                    If Not CheckFeeItemLimitDept(mobjBill.Details(i).�շ�ϸĿID, IIf(mbytUseType = 2, UserInfo.����ID, mobjBill.����ID), IIf(mbytUseType = 2, UserInfo.����ID, mobjBill.����ID)) Then
                        If mbytUseType = 2 Then
                            MsgBox "��" & i & "�е��շ���Ŀ�������ڵĿ��Ҳ����ã�", vbInformation, gstrSysName
                        Else
                            MsgBox "��" & i & "�е��շ���Ŀ�Ե�ǰ���˲����Ϳ��Ҳ����ã�", vbInformation, gstrSysName
                        End If
                        Bill.Row = i: Bill.MsfObj.TopRow = i
                        Bill.Col = BillCol.��Ŀ: Bill.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        Next
        
        '�������ͼ��
        If Not Check�������� Then Exit Sub
        
        '���ʷ��౨��:����ʱ��Ϊ���۵����ٱ���
        mrsWarn.Filter = ""
        If mrsWarn.RecordCount > 0 And mstrWarn <> "-" And Not mblnSavePrice Then
            curTotal = CalcGridToTal '���ݷ���
            If curTotal > 0 Then
                'ˢ�²���Ԥ������Ϣ
                Set rsTmp = GetMoneyInfo(mrsInfo!����ID, IIf(gbytBilling = 0, mcurModiMoney, 0), True, 2)
                If Not rsTmp Is Nothing Then
                    cmdOK.Tag = rsTmp!Ԥ�����
                    cmdCancel.Tag = rsTmp!�������
                    txtʵ��.Tag = rsTmp!Ԥ����� - rsTmp!�������
                Else
                    cmdOK.Tag = 0: cmdCancel.Tag = 0: txtʵ��.Tag = 0
                End If
                '����ʱ��ʾ���㵱ǰ���ݷ���,�����۱���Ҫ��
                '����:30604
                Call SetStatuPatiInfor(Val(cmdOK.Tag), Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), Val(txtʵ��.Tag) - IIf(gbytBilling = 0, curTotal, 0))
                '���¶�ȡ���ն�
                cur���ն� = GetPatiDayMoney(mrsInfo!����ID)
                                
                cur��� = Val(txtʵ��.Tag)
                If gbln�����������۷��� Then cur��� = Val(txtʵ��.Tag) - GetPriceMoneyTotal(1, mrsInfo!����ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
                
                '����ȷ���Ǽ��ʱ���ʱ,�������ķ�ʽ������
                '����ǻ���ģʽ,��Ϊ�ް�ť����,��������µķ�ʽ������
                For i = 1 To mobjBill.Details.Count
                    gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!���� & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), Val("" & mrsInfo!����ID), mrsInfo!���ò���, mrsWarn, cur���, cur���ն� - mcurModiMoney, curTotal, _
                                IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), mobjBill.Details(i).�շ����, mobjBill.Details(i).Detail.�������, mstrWarn, , gblnPrice And gbytBilling = 1)
                    If gbytWarn = 2 Or gbytWarn = 3 Then Exit Sub
                Next
            End If
        End If
        
        'ҩƷ���ɼ��
        strInfo = CheckDisable(mobjBill)
        If strInfo <> "" Then
            If strInfo Like "*(�������)*" Then
                MsgBox strInfo, vbInformation, gstrSysName
                Exit Sub
            Else
                If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
                    
        '�����������
        If Not gbln�������� And mbln����������� Then
            If Not CheckLimit(mobjBill, , gblnסԺ��λ) Then Exit Sub
        End If
        
        '��������ʱ��ҩƷͬһҩ���Ƿ����ظ�����
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If (.Detail.���� Or .Detail.���) _
                    And (InStr(",5,6,7,", .�շ����) > 0 Or .�շ���� = "4" And .Detail.��������) Then
                    For j = 1 To mobjBill.Details.Count
                        If i <> j And .�շ�ϸĿID = mobjBill.Details(j).�շ�ϸĿID And .ִ�в���ID = mobjBill.Details(j).ִ�в���ID Then
                            If .�շ���� = "4" Then
                                If .Detail.���� = mobjBill.Details(j).Detail.���� And .Detail.���� > 0 Then
                                    MsgBox "�� " & j & " �з�����������""" & .Detail.���� & """ ��ͬһ�����ϲ��ű��ظ�������ͬ���Σ���ϲ���", vbInformation, gstrSysName
                                    Exit Sub
                                ElseIf .Detail.���� <= 0 Then
                                    MsgBox "�� " & j & " �еķ�����ʱ����������""" & .Detail.���� & """��ͬһ�����ϲ��ű��ظ����룬��ϲ���", vbInformation, gstrSysName
                                    Exit Sub
                                End If
                            Else
                                MsgBox "�� " & j & " �еķ�����ʱ��ҩƷ""" & .Detail.���� & """��ͬһ��ҩ�����ظ����룬��ϲ���", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            
                        End If
                    Next
                End If
            End With
        Next
        
        'ҩƷ�����,71188:������,2014-04-03,�Բ������ѵ�ҲҪ���м��
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
                If InStr(",5,6,7,", .�շ����) > 0 And Not gbln���뷢ҩ Then
                    If .Detail.���� Or .Detail.��� Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If gblnסԺ��λ Then .Detail.��� = .Detail.��� / .Detail.סԺ��װ
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ��ʱ�ۻ����ҩƷ""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If gblnסԺ��λ Then .Detail.��� = .Detail.��� / .Detail.סԺ��װ
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ��ҩƷ""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .ִ�в���ID) = 1 Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If gblnסԺ��λ Then .Detail.��� = .Detail.��� / .Detail.סԺ��װ
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblTotal > .Detail.��� Then
                            If MsgBox("�� " & i & " ��ҩƷ""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """,Ҫ������?", vbInformation + vbYesNo, gstrSysName) = vbNo Then Exit Sub
                        End If
                    End If
                ElseIf InStr(",5,6,7,", .�շ����) > 0 And gbln���뷢ҩ And gblnStock Then
                    '���ݶ���Ŀ���Ǳ��ز���ָ����ҩ���Ŀ��֮��
                    strInfo = Decode(.Detail.���, "5", gstr��ҩ��, "6", gstr��ҩ��, "7", gstr��ҩ��)
                    If strInfo <> "" Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, 0)
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, 0)
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ��ҩƷ""" & .Detail.���� & "]�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & _
                                dblTotal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                ElseIf .�շ���� = "4" And .Detail.�������� Then
                    If .Detail.���� Or .Detail.��� Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ��ʱ�ۻ������������""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblTotal > .Detail.��� Then
                            MsgBox "�� " & i & " ����������""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    ElseIf colStock("_" & .ִ�в���ID) = 1 Then
                        dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblTotal > .Detail.��� Then
                            If MsgBox("�� " & i & " ����������""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivsOpt, ";��ʾ���;") > 0, .Detail.���, "") & "������������""" & dblTotal & """,Ҫ������?", vbInformation + vbYesNo, gstrSysName) = vbNo Then Exit Sub
                        End If
                    End If
                End If
            End With
        Next
    
        '���ۼ��,105875
        If Not gobjPublicDrug Is Nothing Then
            'Private Function zlCheckPriceAdjustBySell(ByVal lngҩƷid As Long, ByVal lngҩ��id As Long) As Boolean
            '���۹���ģʽʱ���жϼ۸��Ƿ��������۹���Ҫ���ɱ��ۺ��ۼ�һ�£�
            '����ҩƷ���ۼ��ǹ̶��ģ��Ƚ�����ҩ���ĳɱ��ۣ�������ڲ�һ�µľͲ������۳���
            'ʱ��ҩƷ���Ƚ�ҩ������¼�����ۼۺͳɱ��ۣ�������ڲ�һ�µľͲ������۳���
            '���۳���ʱֻ�ж�ҩ��
            '���أ�True-�����������۳��⣻false-���ܽ������۳���
            For i = 1 To mobjBill.Details.Count
                With mobjBill.Details(i)
                    If InStr(",5,6,7,", .�շ����) > 0 Then
                        If gobjPublicDrug.zlCheckPriceAdjustBySell(.�շ�ϸĿID, .ִ�в���ID) = False Then
                            Exit Sub
                        End If
                    End If
                End With
            Next
        End If
        
        '���˺�:22441,����������͸����������
        If CheckMainOperation = False Then Exit Sub
        
        
        '��Ŀ���������(��Ҫ��Ϊ�����������۲���)
        If Check������� > 0 Then Exit Sub
        
        '�����˷Ѽ��
        If Not CheckBillNegative Then Exit Sub
        
        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModule, 1, 1, _
            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 2, IIf(gbytBilling = 1 Or mblnSavePrice, 1, 0))) = False Then
            Exit Sub
        End If
        
        '����������ϵ����Ч��
        '���ʺ��Զ���ҩ
        mblnSendMateria = False
        For i = 1 To mobjBill.Details.Count
            With mobjBill.Details(i)
                If .�շ���� = "4" And .Detail.�������� Then
                    dblTotal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                    If Not CheckValidity(.�շ�ϸĿID, .ִ�в���ID, dblTotal) Then Exit Sub
                    
                ElseIf InStr(1, ",5,6,7,", .�շ����) > 0 Then
                    '��ӡ��ҩ��,����ͨ����,�һ��۵�����
                    If gbytSendMateria <> 0 And mbytUseType = 0 And gbytBilling = 0 And Not mblnSavePrice Then
                        'ȫ��ҩƷ��ȷ����ҩ���Ĳ��Զ���ҩ(���뷢ҩʱ,û��ȷ��ҩ��)
                        mblnSendMateria = .ִ�в���ID <> 0
                    End If
                End If
            End With
        Next
        If InStr(mstrPrivsOpt, ";ҩƷ��ҩ;") = 0 Then mblnSendMateria = False
        
        If mstrInNO <> "" Then
            If HaveExecute(2, mstrInNO, 2) Then
                MsgBox "�õ��ݰ�����ȫִ�л򲿷�ִ�е���Ŀ,�������޸ġ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        If mblnSendMateria And gbytSendMateria = 2 Then
            If MsgBox("������ɺ��Զ�ִ�з�ҩ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnSendMateria = False
            End If
        End If
        
        If IsDate(txtDate.Text) Then mobjBill.����ʱ�� = CDate(txtDate.Text)
        mobjBill.�Ǽ�ʱ�� = zlDatabase.Currentdate      'ע��:��ӡ��ҩ��ʱҪ�õ����ʱ��
        If zlGetSaveDataItems_Plugin(mobjBill, rsItems) = False Then Exit Sub
        If zlChargeSaveValied_Plugin(mlngModule, 2, False, gbytBilling = 1, "", rsItems) = False Then Exit Sub
        
        cmdOK.Enabled = False
        If Not SaveBill Then
            cmdOK.Enabled = True
            If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
            Exit Sub
        Else
            Call zlChargeSaveAfter_Plugin(mlngModule, mobjBill.����ID, mobjBill.��ҳID, False, 2, mobjBill.NO)
            If gbytBilling = 0 And Not mblnSavePrice And gbln���ʴ�ӡ Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_113" & 3 + mbytUseType, Me, "NO=" & mobjBill.NO, "�Ǽ�ʱ��=" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), "ҩƷ��λ=" & IIf(gblnסԺ��λ, 1, 0), "PrintEmpty=0", "�ش�=0", 2)
            ElseIf (gbytBilling = 1 Or mblnSavePrice) And gbln���۴�ӡ Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & mobjBill.NO, "�Ǽ�ʱ��=" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), "ҩƷ��λ=" & IIf(gblnסԺ��λ, 1, 0), "PrintEmpty=0", "�ش�=0", 2)
            End If
            
            '��ӡ��ҩ��
            If mblnSendMateria Then
                If MsgBox("����""" & mobjBill.NO & """��ҩ��ɣ�Ҫ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133_2", Me, "���ݺ�=" & mobjBill.NO, "�Ǽ�ʱ��=" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), 1)
                End If
            End If
                       
            cmdOK.Enabled = True
            If mstrInNO = "" And Not mblnCopyBill Then
                txtPreNO.Text = mobjBill.NO
                Call ClearRows: Call Bill.ClearBill: Call SetColNum
                
                '����ʱ������������û�����
                If gbytBilling = 0 Then Call ClearMoney
                
                Call SetMoneyList
                mstrInNO = ""

                If mrsInfo.State = 1 Then
                    '����ʱ��Ϊ���۵���,Ҫˢ�·������(NewBillǰ)
                    If mblnSavePrice Then
                        Set rsTmp = GetMoneyInfo(mrsInfo!����ID, , True, 2)
                        If Not rsTmp Is Nothing Then
                            cmdOK.Tag = rsTmp!Ԥ�����
                            cmdCancel.Tag = rsTmp!�������
                            txtʵ��.Tag = rsTmp!Ԥ����� - rsTmp!�������
                        Else
                            cmdOK.Tag = 0: cmdCancel.Tag = 0: txtʵ��.Tag = 0
                        End If
'                        sta.Panels(3).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
'                        sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(Val(cmdCancel.Tag), gstrDec)
'                        sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(Val(txtʵ��.Tag), "0.00")
                        '����:30604
                        Call SetStatuPatiInfor(Val(cmdOK.Tag), Val(cmdCancel.Tag), Val(txtʵ��.Tag))
                        
                    End If
                    
                    Call NewBill(False)
                    txtPatient.Tag = "-" & mrsInfo!����ID
                    
                    With mobjBill
                        .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
                        .��ҳID = IIf(mbln���� And mlng��ҳID <> 0, mlng��ҳID, IIf(IsNull(mrsInfo!��ҳID), 0, mrsInfo!��ҳID))
                        
                        .����ID = IIf(mbln���� And mlngUnitID <> 0, mlngUnitID, Val(Nvl(mrsInfo!����ID)))
                        .����ID = IIf(mbln���� And mlngDeptID <> 0, mlngDeptID, Val(Nvl(mrsInfo!����ID)))
                        
                        .���� = "" & mrsInfo!����
                        .��ʶ�� = IIf(IsNull(mrsInfo!סԺ��), 0, mrsInfo!סԺ��)
                        .���� = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
                        .�Ա� = IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
                        .���� = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
                        .�ѱ� = IIf(IsNull(mrsInfo!�ѱ�), "", mrsInfo!�ѱ�)
                        
                        .Ӥ���� = cboBaby.ItemData(cboBaby.ListIndex)
                        .������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text))
                    End With
                    
                    If mbytUseType = 1 Then
                        Call txtPatient_KeyPress(13) 'ˢ��һЩ������Ϣ
                        Bill.SetFocus
                    Else
                      If txtPatient.Enabled Then txtPatient.SetFocus
                      If mbln���� Then Bill.SetFocus
                    End If
                Else
                    Call NewBill
                    txtPatient.SetFocus
                End If
            ElseIf mstrInNO <> "" Then '�޸�
                gblnOK = True: Unload Me: Exit Sub
            ElseIf mblnCopyBill Then '����
                gblnOK = True: Unload Me: Exit Sub
            End If
        End If
    ElseIf Not Bill.Active Then '���סԺ����״̬
        If mstrInNO = "" Then
            MsgBox "û��סԺ���۵���,�������룡", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If
        
        'ȡ������˵������
        strSQL = ""
        For i = 1 To Bill.Rows - 1
            If Bill.RowData(i) > 0 Then
                strSQL = strSQL & "," & Bill.RowData(i)
            End If
        Next
        strSQL = Mid(strSQL, 2)
        i = GetBillRows(mstrInNO, 2)
        If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
                
        'ҽ�����
        intInsure = BillExistInsure(mstrInNO, , True)
        If intInsure > 0 Then
            'ȥ����ҽ������ƥ����
            MCPAR.�����ϴ� = gclsInsure.GetCapability(support�����ϴ�, , intInsure)
            MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure)
        End If
                
        '���ñ���
        mrsWarn.Filter = ""
        If mrsWarn.RecordCount > 0 Then
            If Not AuditingWarn(mstrPrivsOpt, mrsWarn, mstrInNO, strSQL) Then Exit Sub
        End If
        
        '���ʺ��Զ���ҩ
        mblnSendMateria = False
        If gbytSendMateria <> 0 And mbytUseType = 0 And InStr(mstrPrivsOpt, ";ҩƷ��ҩ;") > 0 Then
            For i = 1 To Bill.Rows - 1
                If InStr(",����ҩ,�г�ҩ,�в�ҩ,", "," & Bill.TextMatrix(i, BillCol.���) & ",") > 0 Then '���ȡ����ʱû�д洢������,��Ϊ���������ж�
                    'ȫ��ҩƷ��ȷ����ҩ���Ĳ��Զ���ҩ(���뷢ҩʱ,û��ȷ��ҩ��)
                    mblnSendMateria = Trim(Bill.TextMatrix(i, BillCol.ִ�п���)) <> ""
                End If
            Next
        End If
        If mblnSendMateria And gbytSendMateria = 2 Then
            If MsgBox("������˺��Զ�ִ�з�ҩ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnSendMateria = False
            End If
        End If
        
        cmdOK.Enabled = False
        arrSMSQL = Array()
        Curdate = zlDatabase.Currentdate
        strAddDate = "To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        strSQL = "zl_סԺ���ʼ�¼_Verify('" & mstrInNO & "','" & UserInfo.��� & "','" & UserInfo.���� & "','" & strSQL & "',NULL," & strAddDate & ")"
        str���ܺ� = zlDatabase.GetNextNo(20)
                    
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            '׼���Զ���ҩ(����ͨ����),�����������в��ܶ�������
            If mblnSendMateria Then
                Set rsTmp = Get����ҩ�嵥(mstrInNO, Format(Curdate, "yyyy-MM-dd HH:mm:ss"), False)
                If rsTmp.RecordCount > 0 Then
                    ReDim arrSMSQL(rsTmp.RecordCount - 1)
                    For i = 0 To rsTmp.RecordCount - 1
                        arrSMSQL(i) = "ZL_ҩƷ�շ���¼_���ŷ�ҩ(" & rsTmp!�ⷿID & "," & rsTmp!ID & ",'" & UserInfo.���� & "'," & strAddDate & ",Null,Null,Null," & str���ܺ� & ")"
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Close
            End If
            'ִ���Զ���ҩ
            For i = 0 To UBound(arrSMSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSMSQL(i)), Me.Caption)
            Next
            
            'ҽ���ϴ�
            If intInsure <> 0 Then
                'ҽ�����������ϸ
                If MCPAR.�����ϴ� And Not MCPAR.������ɺ��ϴ� Then
                    strInfo = ""
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 1, strInfo, , intInsure) Then
                        gcnOracle.RollbackTrans
                        If strInfo <> "" Then MsgBox strInfo, vbInformation, gstrSysName
                        cmdOK.Enabled = True
                        Exit Sub
                    End If
                End If
            End If
        gcnOracle.CommitTrans: blnTrans = False
        
        'ҽ���ϴ�
        If intInsure <> 0 Then
            'ҽ�����������ϸ
            If MCPAR.�����ϴ� And MCPAR.������ɺ��ϴ� Then
                strInfo = ""
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 1, strInfo, , intInsure) Then
                    If strInfo <> "" Then
                        MsgBox strInfo, vbInformation, gstrSysName
                    Else
                        MsgBox "����""" & mstrInNO & """��������ҽ������ʧ��,�õ�������ˣ�", vbInformation, gstrSysName
                    End If
                    cmdOK.Enabled = True
                    Exit Sub
                End If
            End If
        End If
        
        On Error GoTo 0
        
        If gbytBilling = 2 And gbln��˴�ӡ And mblnPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & mstrInNO, "�Ǽ�ʱ��=" & Format(Curdate, "yyyy-MM-dd HH:mm:ss"), "ҩƷ��λ=" & IIf(gblnסԺ��λ, 1, 0), "PrintEmpty=0", "�ش�=0", 2)
        End If
        
        '��ӡ��ҩ��
        If mblnSendMateria Then
            If MsgBox("����""" & mstrInNO & """��ҩ��ɣ�Ҫ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133_2", Me, "���ݺ�=" & mstrInNO, "�Ǽ�ʱ��=" & Format(Curdate, "yyyy-MM-dd HH:mm:ss"), 1)
            End If
        End If
        cmdOK.Enabled = True
        
        txtPreNO.Text = mstrInNO
        mstrInNO = "": cboNO.Text = ""
        txtPatient.Text = "": txtOld.Text = ""
        txtʵ��.Text = gstrDec: txtӦ��.Text = gstrDec
        Call ClearRows: Call Bill.ClearBill: Call SetColNum
        Call ClearMoney: Call NewBill
        Call SetMoneyList
        cboNO.Locked = False: cboNO.SetFocus
    End If
    Call SetDrawDrugDeptEnabled
    gblnOK = True
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
    cmdOK.Enabled = True
End Sub

Private Sub cmdOK_GotFocus()
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        Bill.Row = 1: Bill.Col = Bill.Cols - 1
    End If
End Sub
Private Sub cmdPrice_Click()
    mblnSavePrice = True
    Call cmdOK_Click
    mblnSavePrice = False
End Sub
Private Sub SetBill�в�ҩEditEnabled()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ������в�ҩ�ı༭״̬
    '���ƣ����˺�
    '���ڣ�2010-08-06 10:58:45
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With Bill
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "��Ŀ" Then
                .ColData(i) = 0
            Else
                .ColData(i) = 5
            End If
        Next
    End With
End Sub

Private Sub cmd�䷽_Click()
'������ҩ�䷽���빦��
    Dim objDetails As BillDetails
    Dim lng����ID As Long, int������Դ As Integer, int���� As Integer
    Dim int���ʽ As Integer, i As Long, lngMax��� As Long
    Dim blnҽ�� As Boolean
    
    If Not (Bill.Active And mbytInState = 0) Then Exit Sub
    '����Ƿ��з���ҩ
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�շ���� <> "7" Then
            Call MsgBox("�ڵ�ǰ�����д��ڲ����в�ҩ���շ���Ŀ����ɾ�����в�ҩ�շ���Ŀ��,�ٽ����䷽!", vbInformation + vbDefaultButton1, gstrSysName)
            Exit Sub
        End If
    Next
    
    '���˿��һ򿪵�����ID
    lng����ID = mobjBill.����ID
    If lng����ID = 0 Then lng����ID = Get��������ID
    
    'ҽ�Ƹ��ʽ
    If cboҽ�Ƹ���.ListIndex <> -1 Then
        int���ʽ = cboҽ�Ƹ���.ItemData(cboҽ�Ƹ���.ListIndex)
        '����:45605
        Call zlIsCheckMedicinePayMode(zlStr.NeedName(cboҽ�Ƹ���), blnҽ��)
    End If
  
      'ȷ��������Դ
    If mrsInfo.State = 1 Then
        If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
            int������Դ = 2
        ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
            int������Դ = 1
        End If
        'And int���ʽ <> 0:������˵:��Ժ��ҽ�������,���,ֻ�ܼ�鸶�ʽΪҽ����
        '����:45605
        int���� = IIf(Nvl(mrsInfo!����) <> "" And blnҽ��, Nvl(mrsInfo!����), 0)   'ҽ�����˿��ܱ��μ��ʲ���ҽ��
    Else
        int������Դ = 2
        int���� = 0
    End If
    '���ô���,���������,������̬�ѱ�,����mlng��ҩ��,����mbytInFun,�ഫ������Դ����,���ۻ��Ǽ���gbytBilling
    Set objDetails = frmCHRecipe.ShowMe(Me, mstrPrivs, gbytBilling, mcurModiMoney, mobjBill.����ID, int������Դ, lng����ID, Get��������ID, _
                    glng��ҩ��, mobjBill.Details, zlStr.NeedName(cbo�ѱ�.Text), int����, chk�Ӱ�.Value = 1, mobjBill.�巨, mrsWarn, mcolStock1, zl��ȡ��ҩ��̬(Bill.Row, True))
                
    If Not objDetails Is Nothing Then
        Screen.MousePointer = 11
        '���ԭ�����е��в�ҩ
        For i = mobjBill.Details.Count To 1 Step -1
            If mobjBill.Details(i).�շ���� = "7" Then
                mobjBill.Details.Remove i
            End If
        Next
        
        lngMax��� = mobjBill.Details.Count
        '��ӱ༭����в�ҩ
        For i = 1 To objDetails.Count
            With objDetails(i)
                Call mobjBill.Details.Add(.Detail, .�շ�ϸĿID, lngMax��� + .���, lngMax��� + .��������, .����ID, .��ҳID, .����ID, .����ID, _
                .����, .�Ա�, .����, .סԺ��, .����, .�ѱ�, .��������, .�շ����, .���㵥λ, .��ҩ����, .����, .����, _
                .���ӱ�־, .ִ�в���ID, .InComes, .���￨��, "", .������, .ҽ�Ƹ���, .������Ŀ��, .���մ���ID, .���ձ���, .ժҪ, .ԭʼ����, .ԭʼִ�в���ID)
            End With
        Next
         '������ҩ�巨
        mobjBill.�巨 = frmCHRecipe.mstr�巨
        
        'ˢ�µ�ǰ�����е���ʾ
        With Bill
            .Redraw = False
            .ClearBill
            .Rows = mobjBill.Details.Count + 1
        End With
        Call InitBillColumnColor
        Call ShowDetails
        Call ShowMoney
        Call SetBill�в�ҩEditEnabled
        Bill.Redraw = True
        Screen.MousePointer = 0
        Bill.SetFocus
    Else
        Bill.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    Dim objTemp As Object
    If Not mblnFirst Then Exit Sub
    
    '������ҩ����
    Call SetDrawDrugDeptVisible
    
    mblnFirst = False
    On Error Resume Next
    If mbytUseType = 1 And mlng����ID <> 0 And mbytInState = 0 Then
        If mblnCopyBill Then
            cmdOK.SetFocus
        ElseIf gblnFromDr Then
            cbo������.SetFocus
        Else
            Bill.SetFocus
        End If
 
    ElseIf gbytBilling = 2 Then
        cboNO.SetFocus
    ElseIf mbytInState = 1 Then
        cmdCancel.SetFocus
    ElseIf mbytInState = 2 Then
        txtDate.SetFocus
    ElseIf mbytInState = 3 Then
        cmdOK.SetFocus
    ElseIf mstrInNO <> "" Then
        Bill.SetFocus
    End If
    If Not Me.ActiveControl Is cbo�������� Then
        cbo��������.SelLength = 0
    End If
    Call SetDrawDrugDeptEnabled
    '101218
    If mblnSetControl Then
        mblnSetControl = False
        Set objTemp = Me.ActiveControl
        If cboTemp.Visible And cboTemp.Enabled Then cboTemp.SetFocus
        If objTemp.Visible And objTemp.Enabled Then objTemp.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl Is Bill Or Me.ActiveControl Is txt���˱�ע Then Exit Sub
    If Me.ActiveControl Is txtBarCode Then Exit Sub
    
    If InStr("',|~" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    '����:29464
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub  '���ܴ������Ƶ�ˢ��:   ;1088029?
    
End Sub

Private Sub Form_Load()
    Dim tmpBill As ExpenseBill
    Dim i As Long, lngPre As Long, strPre As String, strTmp As String, strҩ��IDs As String
    glngFormW = 12000: glngFormH = 7710
    If Not OS.IsDesinMode Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    RestoreWinState Me, App.ProductName, mbytInState
    sta.Visible = True
    mblnSetControl = True
    Call initCardSquareData
    '����:47798
    If mbytInState = 0 Then
        Call GetRegisterItem(g˽��ģ��, Me.Name, "idkind", strTmp)
        Err = 0: On Error Resume Next
        mblnNotClick = True
        IDKind.IDKind = Val(strTmp)
        mblnNotClick = False
        Err = 0: On Error GoTo 0
    End If
    
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p���ʲ���)
    gblnOK = False: mblnValid = False: mblnFirst = True: gbln�������� = False: mbln������۸� = False
    
    If mbytInState <> 3 And mbytInState <> 1 Then mbytNOType = 2     '���鿴������ʱ�Żᴫ��
    If mbytNOType = 0 Then mbytNOType = 2
    
    '��ʼ����������
    Set mobjBill = New ExpenseBill
    If mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3 Then
        If Not InitData Then Unload Me: Exit Sub
    Else
        If Init�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mstrPrivs, mbytUseType, mlngDeptID) = False Then
            Exit Sub
        End If
    End If
    mstrUnitIDs = GetUserUnits
    
    If mbytInState = 0 And (gbytBilling = 0 Or gbytBilling = 1) Then
        chkIn.Visible = True
        txtIn.Visible = True
    End If
    
    '����:????������ҩ����
    Call zlLoadDrawDeptData(mbytUseType, mlngDeptID)
    Call InitFace
    Call NewBill
    

    If mbytInState <> 0 Then '��ʾ�����������ʵ���(1,2,3)
        If Not ReadBill(mstrInNO, (mbytInState = 3), mbytNOType) Then Unload Me: Exit Sub
        cboNO.Text = mstrInNO
        If InStr(mstrPrivsOpt, ";ҽ����ѯ;") = 0 Then cbo������.Visible = False: lbl������.Visible = False
    Else '����,�޸�
        mstrҩƷ�۸�ȼ� = gstrҩƷ�۸�ȼ�
        mstr���ļ۸�ȼ� = gstr���ļ۸�ȼ�
        mstr��ͨ�۸�ȼ� = gstr��ͨ�۸�ȼ�
        '��ȡ�õ��ݵ�����
        If mstrInNO <> "" Then '�޸ĵ���
            Set mobjBill = ImportBill(mstrInNO, False, Me, True, gblnסԺ��λ, , , False, _
                 mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
            
            If mobjBill.NO = "" Then
                MsgBox "��ȡ����ʧ�ܡ�", vbInformation, gstrSysName
                Unload Me: Exit Sub
            Else
                If Not mblnCopyBill Then
                    mcurModiMoney = GetBillMoney(2, mobjBill.NO) 'Ҫ�ڶ�ȡ������Ϣǰ�ȶ�
                Else
                    mstrInNO = "" '�������ݺ����,�������޸�
                    If InStr(mstrPrivsOpt, ";ҽ����ѯ;") = 0 Then mobjBill.������ = ""
                End If
                
                lngPre = mobjBill.��������ID    'txtpatient_keypress�л�Ķ�
                strPre = mobjBill.������
                
                mbln������۸� = True
                txtPatient.Text = "-" & mobjBill.����ID
                Call txtPatient_KeyPress(13)
                
                '����:50822
                If mrsInfo Is Nothing Then
                    Unload Me: Exit Sub
                End If
                If mrsInfo.State <> 1 Then
                    Unload Me: Exit Sub
                End If
                
                mbln������۸� = False
                '���¼���ͳ����
                Call ReCalcInsure
                
                If Not mblnCopyBill Then
                    '��ʾ����ԭ���ݺ�,��������µ��ݺ�
                    cboNO.Text = mobjBill.NO
                End If
                zlControl.CboLocate cboִ������, mobjBill.ִ������, True
                
                Bill.ClearBill: Call SetColNum
                Bill.Rows = mobjBill.Details.Count + 1
                Call InitBillColumnColor
                '����55420,���Ƶ���Ĭ��ʱ��Ϊ��ǰʱ��
                If mblnCopyBill = True Then
                    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                Else
                    txtDate.Text = Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss")
                End If
                chk�Ӱ�.Value = mobjBill.�Ӱ��־
                
                mobjBill.��������ID = lngPre
                mobjBill.������ = strPre
                Call Set�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mobjBill.������, mobjBill.��������ID)
                Call zlControl.CboLocate(cboBaby, mobjBill.Ӥ����, True)
                
                '�޸�ʱӦ���浱ǰ����Ա������
                mobjBill.����Ա��� = UserInfo.���
                mobjBill.����Ա���� = UserInfo.����
                
                If gintPriceGradeStartType < 2 Then
                    If gbln��������ۿ� Then Call CalcMoneys
                Else
                    Call CalcMoneys
                End If
                Call ShowDetails
                Call ShowMoney
                
                Call SetColNum
            End If
        Else
            If mlng����ID <> 0 Then
                txtPatient.Text = "-" & mlng����ID
                Call txtPatient_KeyPress(13)
            End If
        End If
    End If
   
End Sub

Private Sub Form_Resize()
    Dim lngCancelW As Long
    
    On Error Resume Next
    '��������
    fraBarCode.Width = Me.ScaleWidth - fraBarCode.Left
    txtBarCode.Width = ScaleWidth - txtBarCode.Left - 100
    
    Bill.Top = IIf(Not mblnShowBarCode, fraInfo.Top + fraInfo.Height, fraBarCode.Top + fraBarCode.Height)
    Bill.Height = Me.ScaleHeight - picAppend.Height - sta.Height - Bill.Top - 50
   
    picAppend.Top = Me.ScaleHeight - picAppend.Height - sta.Height
    picAppend.Left = Me.ScaleLeft
    picAppend.Width = Me.ScaleWidth - picAppend.Left
    
    If chkCancel.Visible Or lblFlag.Visible Then lngCancelW = chkCancel.Width
    
    
    fraTitle.Width = Me.ScaleWidth - fraTitle.Left
    chkCancel.Left = fraTitle.Width - chkCancel.Width - 60
    lblFlag.Left = chkCancel.Left + (chkCancel.Width - lblFlag.Width) / 2
    
    cboNO.Left = fraTitle.Width - lngCancelW - 60 - cboNO.Width - 30
    lblNO.Left = cboNO.Left - lblNO.Width - 45
        
    fraUnit.Left = Me.ScaleWidth - fraUnit.Width
    fraInfo.Width = Me.ScaleWidth - fraUnit.Width - fraInfo.Left
    
    Bill.Width = Me.ScaleWidth - Bill.Left
    
    fraAppend.Width = Me.ScaleWidth - fraAppend.Left
    
    txtDate.Left = fraAppend.Width - txtDate.Width - 90
    lblDate.Left = txtDate.Left - lblDate.Width - 45
            
    If cbo������.Container Is fraUnit Then
        cbo��������.Left = lblDate.Left - cbo��������.Width - 200
        lbl��������.Left = cbo��������.Left - lbl��������.Width - 45
    Else
        cbo������.Left = lblDate.Left - cbo������.Width - 200
        lbl������.Left = cbo������.Left - lbl������.Width - 45
    End If
    Me.Refresh
    Call MoveStatuPatiInfor
    Call SetButtonPlace
    Me.Refresh
End Sub

Private Sub SetButtonPlace()
'���ܣ����ݹ��ܰ�ť����,���ð�ťλ��
    If cmdOK.Visible And cmdCancel.Visible And cmdPrice.Visible Then
        cmdPrice.Left = fraStat.Left + fraStat.Width + (Me.ScaleWidth - (fraStat.Left + fraStat.Width) - cmdOK.Width - cmdCancel.Width - cmdPrice.Width) / 2
        cmdOK.Left = cmdPrice.Left + cmdPrice.Width
        cmdCancel.Left = cmdOK.Left + cmdOK.Width
    ElseIf cmdOK.Visible And cmdCancel.Visible Then
        cmdOK.Left = fraStat.Left + fraStat.Width + (Me.ScaleWidth - (fraStat.Left + fraStat.Width) - cmdOK.Width - cmdCancel.Width) / 2
        cmdCancel.Left = cmdOK.Left + cmdOK.Width
    ElseIf cmdPrice.Visible And cmdCancel.Visible Then
        cmdPrice.Left = fraStat.Left + fraStat.Width + (Me.ScaleWidth - (fraStat.Left + fraStat.Width) - cmdPrice.Width - cmdCancel.Width) / 2
        cmdCancel.Left = cmdPrice.Left + cmdPrice.Width
        cmdPrice.ToolTipText = cmdOK.ToolTipText
    ElseIf cmdCancel.Visible Then
        cmdCancel.Left = fraStat.Left + fraStat.Width + (Me.ScaleWidth - (fraStat.Left + fraStat.Width) - cmdOK.Width) / 2
    End If
    cmdPrice.TabStop = Not cmdOK.Visible
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mbytInState
    mbytInState = 0
    mblnCopyBill = False
    mstrInNO = ""
    mblnNOMoved = False '�����˳������,����Ӱ���������
    mlngҽ��ID = 0
    mlng����ҽ�� = 0
    
    mlngDelRow = 0
    mlngUnitID = 0
    mstrTime = ""
    mblnDelete = False
    gbytBilling = 0
    mbytUseType = 0
    mlngDeptID = 0
    mlng����ID = 0
    
    mlngҩƷ���ID = 0
    mlng�������ID = 0
    mstr���ת��ʱ�� = ""
    mbln���� = False
    Set mrs�������� = Nothing
    Set mrs�������� = Nothing
    Set mrs������ = Nothing
    Set mrsWarn = Nothing
    Set mrsMedAudit = Nothing
    Set mrsMedPayMode = Nothing
    Set mobjCard = Nothing
    Set mobjBrushCheck = Nothing
    
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    Set mobjBaseItem = Nothing
    If Not OS.IsDesinMode Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    If mbytInState = 0 Then
        Call SaveRegisterItem(g˽��ģ��, Me.Name, "idkind", IDKind.IDKind)
    End If
End Sub

Private Sub mobjBrushCheck_ReadCardNoed(ByVal strCardNo As String, ByVal blnBrushCard As Boolean)
    If blnBrushCard Then
        mbln����ˢ�� = True
    Else
        mbln����ˢ�� = False
    End If
End Sub

Private Sub zlMoveDrawControl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ƶ���ҩ���ſؼ�λ��
    '����:���˺�
    '����:2009-07-29 14:37:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���˺� ����:26953 ����:2009-12-25 14:38:12
    Dim lngLeft As Long
    
    lngLeft = IIf(lblDrawDrugDept.Visible, cboDrawDept.Left + cboDrawDept.Width + 50, lblDrawDrugDept.Left)
    If lblִ������.Visible Then
        '����:27383
        lblִ������.Left = lngLeft: lngLeft = lblִ������.Left + lblִ������.Width + 20
        cboִ������.Left = lngLeft: lngLeft = cboִ������.Left + cboִ������.Width + 50
    End If
    lbl���˱�ע.Left = lngLeft
    
    txt���˱�ע.Left = lbl���˱�ע.Left + lbl���˱�ע.Width + 20
    txt���˱�ע.Width = picAppend.ScaleWidth - txt���˱�ע.Left - 100
    
    fraStat.Top = mshMoney.Top - 120
    cmdOK.Top = mshMoney.Top + (mshMoney.Height - cmdOK.Height) \ 2
    cmdCancel.Top = cmdOK.Top
    cmdPrice.Top = cmdOK.Top
    
    Call Form_Resize
End Sub
Private Sub zlReSetDrawDrugDept()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ӧ�Ĺ���,���»�ȡ��ҩ����
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-07-29 18:23:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '3)  ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
    '4)  סԺ���ʡ����ҷ�ɢ���ʣ������ɲ���ʹ�ã�Ҳ������ҽ������ʹ�á�
    '    a)  �жϵ�ǰ����Ա�������ң����������ҽ�����ʵĿ��ң�����ҩ���Ź̶�Ϊ���˲�����(��顢���顢���������ơ�Ӫ��)
    '    b)  �������Ա����ҽ�����ʵĿ��ң����ڵ��ݽ���������"��ҩ����"ѡ��򣬿�ѡ��ΧΪ����Ա������ҽ�����ʵĿ���(���ܶ��)��ȱʡ�뿪��������ͬ��
    If mbytUseType = 2 Then
        'ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
        mobjBill.��ҩ����ID = mlngDeptID: Exit Sub
    End If
    If mrs��ҩ����.RecordCount = 0 Then
        '�жϵ�ǰ����Ա�������ң����������ҽ�����ʵĿ��ң�����ҩ���Ź̶�Ϊ���˲�����(��顢���顢���������ơ�Ӫ��)
        mobjBill.��ҩ����ID = mobjBill.����ID: Exit Sub
    End If
    '�������Ա����ҽ�����ʵĿ��ң����ڵ��ݽ���������"��ҩ����"ѡ��򣬿�ѡ��ΧΪ����Ա������ҽ�����ʵĿ���(���ܶ��)��ȱʡ�뿪��������ͬ��
    If mrs��ҩ����.RecordCount = 1 Then
        'ֻ��һ������,�϶�����
        If mrs��ҩ����.EOF Then mrs��ҩ����.MoveFirst
         mobjBill.��ҩ����ID = Val(Nvl(mrs��ҩ����!ID)): Exit Sub
    End If
    'ѡ��Ŀ������ĸ������ĸ�
    With cboDrawDept
        If .ListIndex < 0 Then Exit Sub
        If mobjBill.��ҩ����ID <> .ItemData(.ListIndex) Then mobjBill.��ҩ����ID = .ItemData(.ListIndex): Exit Sub
    End With
End Sub

Private Sub zlLoadDrawDeptData(ByVal bytUseType As Byte, Optional ByVal lngDeptID As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:
    '���:bytUseType:���ʵ���;,0-��ͨ����,1-�����ҷ�ɢ����,2-ҽ�����Ҽ���
    '����:
    '����:
    '����:24729,24731
    '����:���˺�
    '����:2009-07-29 15:05:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    
    '3)  ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
    '4)  סԺ���ʡ����ҷ�ɢ���ʣ������ɲ���ʹ�ã�Ҳ������ҽ������ʹ�á�
    '    a)  �жϵ�ǰ����Ա�������ң����������ҽ�����ʵĿ��ң�����ҩ���Ź̶�Ϊ���˲�����(��顢���顢���������ơ�Ӫ��)
    '    b)  �������Ա����ҽ�����ʵĿ��ң����ڵ��ݽ���������"��ҩ����"ѡ��򣬿�ѡ��ΧΪ����Ա������ҽ�����ʵĿ���(���ܶ��)��ȱʡ�뿪��������ͬ��
    
    On Error GoTo errHandle
    
    'ҽ������
    If bytUseType = 2 Then
        '3)  ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
        strSQL = "Select ID,����,���� From ���ű� where id=[2]"
    Else
        strSQL = _
            " Select distinct  A.ID, A.����,A.����   " & vbNewLine & _
            " From ���ű� A, ��������˵�� B,������Ա C" & vbNewLine & _
            " Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)  " & _
            "       And A.ID = B.����id and a.id=C.����ID and C.��Աid=[1] " & vbNewLine & _
            "       And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            "       AND B.�������� IN('���','����','����','����','Ӫ��') " & _
            " Order by ����"
    End If
    Set mrs��ҩ���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, lngDeptID)
    With mrs��ҩ����
        cboDrawDept.Clear
        Do While Not .EOF
            cboDrawDept.AddItem IIf(zlIsShowDeptCode, Nvl(!����) & "-", "") & Nvl(!����)
            cboDrawDept.ItemData(cboDrawDept.NewIndex) = Val(Nvl(!ID))
            If Val(Nvl(!ID)) = UserInfo.����ID Then cboDrawDept.ListIndex = cboDrawDept.NewIndex
            .MoveNext
        Loop
        If .RecordCount <> 0 And cboDrawDept.ListIndex < 0 Then cboDrawDept.ListIndex = 0
        If .RecordCount <> 0 Then .MoveFirst
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetDrawDrugDeptVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ҩ���ŵ�visibled����
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-07-29 19:07:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    ' mbytUseType As Byte '���ʵ���;,0-��ͨ����,1-�����ҷ�ɢ����,2-ҽ�����Ҽ���
    
    '3)  ҽ�����Ҽ���ʱ����Ӧ����ҩ���Ź̶�ȷ��Ϊ��������ѡ����ҽ�����ҡ�(������Ӧֻ�ṩ��������ҺͲ��˿��ҿ�ѡ)
    If mbytUseType = 2 Then
        cboDrawDept.Visible = False
    ElseIf chkCancel.Value = 1 Then
        '����Ҳ���ܿ���
        cboDrawDept.Visible = False
    Else
        'mbytInState As Byte '0-ִ��,1-����,2-����,3-����
        'gbytBilling As Byte '0-����,1-����,2-���
        cboDrawDept.Visible = mrs��ҩ����.RecordCount > 1 And (mbytInState = 0 And gbytBilling <> 2)     '
    End If
    lblDrawDrugDept.Visible = cboDrawDept.Visible
    Call zlMoveDrawControl
End Sub
Private Sub SetDrawDrugDeptEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ҩ���ŵ�Enabled����
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-07-31 11:55:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnHaveDrug As Boolean '����ҩƷ
    
    '���û�����ò��ŵ�ѡ��,��ֱ���˳�
    If cboDrawDept.Visible = False Then cboDrawDept.Enabled = False: lblDrawDrugDept.Enabled = False: Exit Sub
    blnHaveDrug = False
    For i = 1 To mobjBill.Details.Count
        If InStr(1, ",5,6,7,", "," & mobjBill.Details(i).�շ���� & ",") > 0 Then
            blnHaveDrug = True
            Exit For
        End If
    Next
    cboDrawDept.Enabled = blnHaveDrug: lblDrawDrugDept.Enabled = blnHaveDrug
End Sub

Private Sub picAppend_Resize()
    Err = 0: On Error Resume Next
    With picAppend
        fraDrawDept.Left = 0
        fraDrawDept.Width = .ScaleWidth + 50
        txt���˱�ע.Width = .ScaleWidth - txt���˱�ע.Left - 100
    End With
End Sub

Private Sub txtBarCode_GotFocus()
    zlCommFun.OpenIme False
    zlControl.TxtSelAll txtBarCode
End Sub

Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode <> vbKeyReturn Then Exit Sub
   
   If AddStuffItemFromBarCode(txtBarCode.Text) = False Then
        If txtBarCode.Enabled And txtBarCode.Visible Then txtBarCode.SetFocus
        zlControl.TxtSelAll txtBarCode: Exit Sub
   End If
   txtBarCode.Text = ""
   If txtBarCode.Enabled And txtBarCode.Visible Then txtBarCode.SetFocus
End Sub
 
Private Sub txtBarCode_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub ShowAndHideBarCodeInput()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ���������������
    '����:���˺�
    '����:2017-11-22 11:42:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    fraBarCode.Visible = mblnShowBarCode
    txtBarCode.Visible = mblnShowBarCode
    lblBarCode.Visible = mblnShowBarCode
    Call Form_Resize
 End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)

    Select Case Panel.Key
    Case "BarCode"
        '��ʾ����
        mblnShowBarCode = Not mblnShowBarCode
        Panel.Bevel = IIf(Not mblnShowBarCode, sbrRaised, sbrInset)
        Panel.ToolTipText = IIf(Not mblnShowBarCode, "��ʾ���������", "�������������")
        Call ShowAndHideBarCodeInput
        If txtBarCode.Enabled And txtBarCode.Visible Then txtBarCode.SetFocus
        Call zlDatabase.SetPara("�ϴ�ѡ���������", IIf(mblnShowBarCode, 1, 0), glngSys, 1150)
        Exit Sub
    Case "PY", "WB"
        If Panel.Bevel = sbrRaised And gbln�����л� Then
            '�л����������ƥ�䷽ʽ
            Panel.Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            If Panel.Key = "PY" Then
                sta.Panels("WB").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            Else
                sta.Panels("PY").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            End If
            zlDatabase.SetPara "���뷽ʽ", IIf(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIf(sta.Panels("WB").Bevel = sbrInset, 1, 0))
            gbytCode = Val(zlDatabase.GetPara("���뷽ʽ", , , 0))
        End If
    Case "Drugstore"
        With frmSetExpence
            .mlngModul = mlngModule
            .mstrPrivs = mstrPrivs
            
            '���ʵ���;,0-��ͨ����,1-�����ҷ�ɢ����,2-ҽ�����Ҽ���
            '           0:��ͨ����,1-���ҷ�ɢ����,2-ҽ�����Ҽ���
            .mbytInFun = 0
            .mbytUseType = mbytUseType
            .mblnOnlyDrugStock = True
            .Show 1, Me
        End With
    End Select
End Sub
 
Private Sub tmrStatuPati_Timer()
    If picStatuPancl.Visible Then Call MoveStatuPatiInfor
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cboNO_GotFocus()
    zlControl.TxtSelAll cboNO
    If gbytBilling = 2 Or chkCancel.Value = Checked Then
        cboNO.Locked = False
    Else
        cboNO.Locked = True
    End If
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim blnRead As Boolean, strOper As String
    Dim vDate As Date, intTmp As Integer
    Dim strInfo As String, intInsure As Integer, blnFlagPrint As Boolean
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    End If
    
    If KeyAscii = 13 And cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 14)
        
        If chkCancel.Value = 1 Then
            '����
            
            If gbytBilling = 0 Then
                '�Ƿ���ת������ݱ���
                If zlDatabase.NOMoved("סԺ���ü�¼", cboNO.Text, , 2, Me.Caption) Then
                    If Not ReturnMovedExes(cboNO.Text, 2, Me.Caption) Then Exit Sub
                    mblnNOMoved = False
                End If
            End If
        
            '�����˻���ȫ��˵Ĳ���������
            If Not BillIdentical(cboNO.Text) Then
                MsgBox "�����а�������δ��ȫ��˻�ֶ����˵����ݣ����������������ʡ�" & _
                    vbCrLf & "���˻ع��������˳���Ӧ�ĵ������ݣ�Ȼ�������ʡ�", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        
            '����Ȩ��
            If Not ReadBillInfo(2, cboNO.Text, 2, strOper, vDate) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            If mbytUseType = 0 And InStr(mstrPrivs, ";���в���Ա;") <= 0 Then
                If UserInfo.���� <> strOper Then
                    MsgBox "��û��""���в���Ա""Ȩ��,���ܶ�" & strOper & "�ĵ��ݽ�������!", vbInformation, gstrSysName
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            If Not BillOperCheck(5, strOper, vDate, "����", cboNO.Text) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            
            '��Ŀ����Ȩ��
            If mbytUseType = 0 Or mbytUseType = 1 Then
                If Not CheckDelPriv(cboNO.Text, mstrPrivsOpt) Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
                
            '���۲���Ȩ��
            strInfo = Check���۲���(cboNO.Text, mstrPrivsOpt)
            If strInfo <> "" Then
                MsgBox "�����а���" & strInfo & ",��û��Ȩ�޶Ըõ��ݽ��в�����", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            
            '�Ƿ���ִ��
            intTmp = BillCanDelete(cboNO.Text, 2, , , mstrPrivsOpt, blnFlagPrint)
            If intTmp <> 0 Then
                Select Case intTmp
                    Case 1 '�õ��ݲ�����
                        MsgBox "ָ�������е����ݲ�����,������û������շ���Ŀ������Ȩ�ޣ�", vbInformation, gstrSysName
                    Case 2 '�Ѿ�ȫ����ȫִ��
                        MsgBox "ָ�������е������Ѿ�ȫ����ȫִ�У�", vbInformation, gstrSysName
                    Case 3 'δ��ȫִ�в���ʣ������Ϊ0
                        MsgBox "ָ�������е�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�", vbInformation, gstrSysName
                End Select
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            If blnFlagPrint Then
                If MsgBox("ע��:����ҽ���������Ѵ�ӡ���Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
                    
            '��Ժ���˲���Ȩ���ж�
            If Not BillCanBeOperate(cboNO.Text, mstrPrivsOpt, "����") Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            
            '�Ƿ��ѽ���
            intInsure = BillExistInsure(cboNO.Text)
            intTmp = HaveBilling(2, cboNO.Text, False)
            If intTmp <> 0 Then
                If intInsure <> 0 Then
                    If Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, , intInsure) Then
                        'ҽ�����˵ĵ���,�̶�Ϊ�ѽ��ʵĽ�ֹ����
                        If intTmp = 1 Then
                            MsgBox "��ҽ�����ʵ���δ���ʲ����Ѿ�����,�������ʣ�", vbInformation, gstrSysName
                            cboNO.Text = "": cboNO.SetFocus: Exit Sub
                        Else
                            MsgBox "��ҽ�����ʵ��ݰ����Ѿ����ʵ�����,ֻ�ܶ�δ���ʲ��ֽ������ʣ�", vbInformation, gstrSysName
                        End If
                    End If
                Else
                    Select Case gbytBillOpt
                        Case 0
                        Case 1
                            If MsgBox("�ü��ʵ��ݰ����Ѿ����ʵ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                cboNO.Text = "": cboNO.SetFocus: Exit Sub
                            End If
                        Case 2
                            If intTmp = 1 Then
                                MsgBox "�ü��ʵ���δ���ʲ����Ѿ�����,�������ʣ�", vbInformation, gstrSysName
                                cboNO.Text = "": cboNO.SetFocus: Exit Sub
                            Else
                                MsgBox "�ü��ʵ��ݰ����Ѿ����ʵ�����,ֻ�ܶ�δ���ʲ��ֽ������ʣ�", vbInformation, gstrSysName
                            End If
                    End Select
                End If
            End If
            
            'ҽ�����ʲ�����Ը�����¼��������
            If intInsure <> 0 Then
                If CheckNONegative(cboNO.Text) Then
                    MsgBox "�õ��ݴ��ڸ������ʼ�¼,���������ҽ�����ʲ�����", vbInformation, gstrSysName
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
                        
            '�Ƿ������������¼
            If CheckRecalcRecord(cboNO.Text) Then
                MsgBox "���ָü��ʵ��ݴ��ڰ��ѱ�����Ĵ��۳����¼!" & vbCrLf & _
                    "����ǰ�밴�ѱ�������ã������˽����������ʵ��ݵĴ����Żݽ�", vbInformation, Me.Caption
            End If
        ElseIf mobjBill.Details.Count = 0 Then
            '���ʻ��۵�(�������)
            If Not BillExistMoney(cboNO.Text, 2) Then
                MsgBox "�õ��ݷ����Ѿ�ȫ�����ʻ򵥾ݲ����ڣ�", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            
            '��Ժ���˲���Ȩ���ж�
            If Not BillCanBeOperate(cboNO.Text, mstrPrivsOpt, "���") Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        End If
        
        If chkCancel.Value = 1 Then '��ȡ�˷ѵ�
            blnRead = ReadBill(cboNO.Text, True)
        ElseIf mobjBill.Details.Count = 0 Then '��ȡסԺ���۵�
            blnRead = ReadBill(cboNO.Text, False)
        End If
        
        If blnRead Then
            
            mstrInNO = cboNO.Text 'ȷ��ʱ��mstrInNOΪ׼
            If chkCancel.Value = 0 Then
                '���۵�
                Bill.Active = False
            Else
                '����
                Bill.Active = True
            End If
            cmdOK.SetFocus
            If gbytBilling = 2 Then  '���ʱ
                Call SetDisible
                cboNO.Locked = False
            End If
        Else
            mstrInNO = "": cboNO.Text = "": cboNO.SetFocus
        End If
    End If

End Sub

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
    If IsDate(txtDate.Text) Then mobjBill.����ʱ�� = CDate(txtDate.Text)
End Sub

Private Sub txtOld_Gotfocus()
    zlControl.TxtSelAll txtOld
End Sub

Private Sub txtOld_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mobjBill.���� = txtOld.Text
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep
End Sub

Private Sub txtPatient_Change()
    txtPatient.Tag = ""
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "")
End Sub


Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If txtPatient.Locked = True Then Exit Sub
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    Dim i As Long

    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        Bill.Row = 1: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    
    With Bill
        '������ʱ,�������ÿ����Ѿ������ĵĿɱ������е���ֵ
        If mbytInState <> 2 Then
            .ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus)  '�����,��������ʱ�ᱻ�ı�
            .ColData(BillCol.��Ŀ) = BillColType.CommandButton    '��Ŀ��,��������ʱ�ᱻ�ı�
            .ColData(BillCol.����) = BillColType.UnFocus   '����ȱʡ����(=1),�����Ϊ��ҩʱ,��Ϊ����(4)(��ֵ,һ��ȫ��)
            .ColData(BillCol.����) = BillColType.UnFocus  '����ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
            .ColData(BillCol.��־) = BillColType.UnFocus  '��־ȱʡ����,��Ϊ����ʱ,��Ϊ��ѡ(-1)
        End If
        
        '����б༭����������ɫ
        .SetColColor BillCol.���, &HE7CFBA
        .SetColColor BillCol.��Ŀ, &HE7CFBA
        .SetColColor BillCol.����, &HE7CFBA
        .SetColColor BillCol.ִ�п���, &HE7CFBA
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.��־, &HE0E0E0
        
        .TextMatrix(Row, BillCol.��) = Row
        
        '����ط��ֶ����ò�ִ��
        If Visible And Bill.Active And Row > 0 And .ColData(BillCol.���) <> BillColType.UnFocus And Not mblnNewRow Then
            Call zlCommFun.PressKey(13)
        End If
    End With
End Sub

Private Sub cboSex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboSex.ListIndex <> -1 Then mobjBill.�Ա� = Mid(cboSex.Text, InStr(cboSex.Text, "-") + 1)
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cbo�ѱ�.ListIndex <> -1 Then
        mobjBill.�ѱ� = zlStr.NeedName(cbo�ѱ�.Text)
        If mbytInState = 0 And mstrInNO <> "" And mobjBill.Details.Count > 0 Then
            '���¼���۸�
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)

   Dim lngIdx As Long, lngҽ��ID As Long
    
    If KeyAscii <> 13 Then Exit Sub
    If cbo��������.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cbo������.ListIndex >= 0 Then lngҽ��ID = cbo������.ItemData(cbo������.ListIndex)
    If mrs�������� Is Nothing Then Call FillDept(cbo��������, mrs��������, mrs������, mstrPrivs, mbytUseType, mlngDeptID, lngҽ��ID)
    
    If zlSelectDept(Me, mlngModule, cbo��������, mrs��������, cbo��������.Text) = False Then
        Call Beep: mobjBill.��������ID = 0
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub
Private Function isCheck������Exists(ByVal str���� As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ��ڿ����������б���.
    '���:str����-����
    '     blnLocateItem:�Ƿ�ֱ�Ӷ�λ
    '����:
    '����:
    '����:���˺�
    '����:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cbo������.ListCount - 1
        If zlStr.NeedName(cbo������.List(i)) = str���� Then
            If blnLocateItem Then cbo������.ListIndex = i
            isCheck������Exists = True
            Exit Function
        End If
    Next
End Function

Private Sub cbo������_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim rsTemp As ADODB.Recordset
    Dim strAdded As String
    If KeyAscii = 13 Then
        If cbo������.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = UCase(cbo������.Text)
        If cbo������.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cbo������.List(cbo������.ListIndex) Then Call zlControl.CboSetIndex(cbo������.hWnd, -1)
        End If
        If strText = "" Then
            cbo������.ListIndex = -1
        ElseIf cbo������.ListIndex = -1 Then
            intIdx = -1
            strFilter = IIf(gbln��ʿ, "��Ա����<>''", "��Ա����<>'��ʿ'")
            '���˺�:22383
            '�ȸ��Ƽ�¼��
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrs������)
            Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
            Dim strCompents As String 'ƥ�䴮
            
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrs������.Filter = strFilter: iCount = 0
            With mrs������
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrs������.EOF
                    Select Case intInputType
                    Case 0  '�������ȫ����
                        '������������,��Ҫ���:
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                        
                        
                        '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                        If Nvl(!���) = strText Then strResult = Nvl(!����): iCount = 0: Exit Do
                        
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                        If Val(Nvl(!���)) = Val(strText) Then
                            If iCount = 0 Then strResult = Nvl(!����)
                            iCount = iCount + 1
                        End If
                        
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                         If Val(Nvl(!���)) Like strText & "*" Then
                            If isCheck������Exists(Nvl(!����)) And InStr(strAdded, "," & Nvl(!���) & ",") = 0 Then
                                Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                                strAdded = strAdded & "," & Nvl(!���) & ","
                            End If
                         End If
                    Case 1  '�������ȫ��ĸ
                        '����:
                        ' 1.����ļ������,��ֱ�Ӷ�λ
                        ' 2.���ݲ�����ƥ����ͬ����
                        
                        '1.����ļ������,��ֱ�Ӷ�λ
                        If Trim(Nvl(!����)) = strText Then
                            If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.���ݲ�����ƥ����ͬ����
                        If Trim(Nvl(!����)) Like strCompents Then
                            If isCheck������Exists(Nvl(!����)) And InStr(strAdded, "," & Nvl(!���) & ",") = 0 Then
                                Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                                strAdded = strAdded & "," & Nvl(!���) & ","
                            End If
                        End If
                    Case Else  ' 2-����
                        '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                        '1.����\�������,ֱ�Ӷ�λ
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        
                        '1.����\�������,ֱ�Ӷ�λ
                        If Trim(!���) = strText Or Trim(!����) = strText Or Trim(!����) = strText Then
                            If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        If Trim(!���) Like strText & "*" Or Trim(Nvl(!����)) Like strCompents Or Trim(Nvl(!����)) Like strCompents Then
                            If isCheck������Exists(Nvl(!����)) And InStr(strAdded, "," & Nvl(!���) & ",") = 0 Then
                                Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                                strAdded = strAdded & "," & Nvl(!���) & ","
                            End If
                        End If
                    End Select
                    mrs������.MoveNext
                Loop
            End With
             If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!����)
            '���˺�:ֱ�Ӷ�λ
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheck������Exists(strResult, True) Then zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '��Ҫ����Ƿ��ж������������ļ�¼
            If rsTemp.RecordCount <> 0 Then
                '�Ȱ�ĳ�ַ�ʽ��������
                Select Case intInputType
                Case 0 '����ȫ����
                    rsTemp.Sort = "���"
                Case 1 '����ȫƴ��
                    rsTemp.Sort = "����"
                Case Else
                    '����ѡ������
                    If gbyt��������ʾ = 1 Then '����
                        rsTemp.Sort = "����"
                    Else
                        rsTemp.Sort = "���"
                    End If
                End Select
                '����ѡ����
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, 1133, cbo������, rsTemp, True, "", "ȱʡ,ְ��,���ȼ���", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '���ж�λ
                            If isCheck������Exists(Nvl(rsReturn!����), True) Then
                                'zlCommFun.PressKey vbKeyTab
                            End If
                        End If
                    End If
                End If
            Else
                'δ�ҵ�
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: zlControl.TxtSelAll cbo������: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
             
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call cbo������_Click
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cbo������.ListIndex = -1 Then
            cbo������.Text = ""
            mobjBill.������ = ""
            If gblnFromDr Then Exit Sub
        Else
            mobjBill.������ = zlStr.NeedName(cbo������.Text)
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
                Call cbo������_Click
            ElseIf intIdx <> cbo������.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cbo������.SetFocus
                Call zlCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
                Call cbo������_Click
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If mbytInState = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyF1
            ShowHelp App.ProductName, Me.hWnd, Me.Name
        Case vbKeyF2
            If ActiveControl Is txtPatient Then Call txtPatient_Validate(False)
            If ActiveControl Is cbo������ Then Call cbo������_KeyPress(vbKeyReturn)
            If cmdOK.Enabled And cmdOK.Visible Then
                Call cmdOK.SetFocus
                Call cmdOK_Click
            ElseIf cmdPrice.Enabled And cmdPrice.Visible Then
                Call cmdPrice.SetFocus
                Call cmdPrice_Click
            End If
        Case vbKeyF3    '���뵥��
            If chkIn.Visible And chkIn.Enabled Then chkIn.Value = IIf(chkIn.Value = 1, 0, 1)
        Case vbKeyF4
            If Shift = vbCtrlMask And IDKind.Enabled Then
                Dim intIndex As Integer
                intIndex = IDKind.GetKindIndex("IC����")
                If intIndex <= 0 Then Exit Sub
                IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
            End If
        Case vbKeyF6    '��λ�����������
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            Call zlControl.TxtSelAll(txtPatient)
        Case vbKeyF7    '�л����뷨
            If gbln�����л� Then
                If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                    If sta.Panels("WB").Bevel = sbrRaised Then
                        Call sta_PanelClick(sta.Panels("WB"))
                    Else
                        Call sta_PanelClick(sta.Panels("PY"))
                    End If
                End If
            End If
        Case vbKeyF8    '��(�Զ������¼�)
            If chkCancel.Visible And chkCancel.Enabled Then chkCancel.Value = IIf(chkCancel.Value = Checked, Unchecked, Checked)
        Case vbKeyF9 '��λ�����ݺ������
            cboNO.SetFocus
            Call zlControl.TxtSelAll(cboNO)
        Case vbKeyF11
            If cmd�䷽.Enabled And cmd�䷽.Visible Then Call cmd�䷽_Click
        Case vbKeyF12
            If Shift = vbAltMask Then
                Call sta_PanelClick(sta.Panels("Drugstore"))
            End If
              
        Case vbKeyA, vbKeyR
            'ȫѡ��ȫ��
            If Shift = vbCtrlMask Then
                If KeyCode = vbKeyA Then
                    Call SelALLRow
                ElseIf KeyCode = vbKeyR Then
                    Call ClearALLRow
                End If
            End If
        Case vbKeyQ
            If Shift = vbCtrlMask Then
                Call LocateNewRow
            End If
        Case vbKeyEscape
            If Bill.TxtVisible Then
                Bill.Text = "": Bill.TxtVisible = False: Bill.SetFocus
            Else
                Call cmdCancel_Click
            End If
    End Select
End Sub

Private Sub SetMoneyList()
'����:���ݵ�ǰ������Ŀ�����������п�
    Dim lngW As Long
    lngW = mshMoney.Width - 60
    If mshMoney.Rows > mshMoney.Height / mshMoney.RowHeight(0) Then
        lngW = lngW - 250
    End If
    mshMoney.ColWidth(0) = lngW * 0.5
    mshMoney.ColWidth(1) = lngW * 0.5
    
    mshMoney.ColAlignment(0) = 1
    mshMoney.ColAlignment(1) = 7
    
    mshMoney.TextMatrix(0, 0) = "��Ŀ"
    mshMoney.TextMatrix(0, 1) = "���"
    mshMoney.Row = 0
    mshMoney.ColAlignmentFixed(0) = 4
    mshMoney.ColAlignmentFixed(1) = 4
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim Curdate As Date     '��������ǰʱ��
    On Error GoTo errH
   
    If mbytInState = 0 And mstrInNO = "" Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    
    '��ȡ��ҩ������
    If mbytUseType = 0 Or mbytUseType = 1 Then Call ReadABCNum(mstrPrivsOpt)
    
    '��ͬҩ��ҩƷ�����鷽ʽ
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    

    '------------------������ȡ------------------
    
    '��ѡ�Ա�,ҽ�Ƹ��ʽ,���㷽ʽ
    strSQL = " Select '�Ա�' as ���,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �Ա� Union All " & _
             " Select '�ѱ�' as ���,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �ѱ� Union All " & _
             " Select 'ҽ�Ƹ��ʽ' as ���,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From ҽ�Ƹ��ʽ "
    
    strSQL = strSQL & " Order by ���,����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    '1.�Ա�
    rsTmp.Filter = "���='�Ա�'"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboSex.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then cboSex.ListIndex = cboSex.NewIndex
            rsTmp.MoveNext
        Next
    End If
    '2.�ѱ�,�����й̶��ѱ�,�뿪�������޹�
    rsTmp.Filter = "���='�ѱ�'"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo�ѱ�.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 And cbo�ѱ�.ListIndex = -1 Then cbo�ѱ�.ListIndex = cbo�ѱ�.NewIndex
            rsTmp.MoveNext
        Next
    Else
        MsgBox "û�г�ʼ���ѱ����ȵ��ѱ�����н������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '3.ҽ�Ƹ��ʽ
    rsTmp.Filter = "���='ҽ�Ƹ��ʽ'"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboҽ�Ƹ���.AddItem rsTmp!���� & "-" & rsTmp!����
            cboҽ�Ƹ���.ItemData(cboҽ�Ƹ���.NewIndex) = Val(rsTmp!����)
            If rsTmp!ȱʡ = 1 Then
                cboҽ�Ƹ���.ListIndex = cboҽ�Ƹ���.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    
    strSQL = " Select '����ְ��' As ����,count(ҩ��ID) As num From ҩƷ���� Where ����ְ��<>'00' Union All " & _
             " Select '��������' As ����,count(ҩ��ID) As num From ҩƷ���� Where ��������>0    "
    
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    rsTmp.Filter = "����='����ְ��'"
    If Not rsTmp.EOF Then mbln����ְ���� = (rsTmp!Num > 0)
    
    rsTmp.Filter = "����='��������'"
    If Not rsTmp.EOF Then mbln����������� = (rsTmp!Num > 0)

    
    '------------------������ȡ------------------
            
    If Init�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mstrPrivs, mbytUseType, mlngDeptID) = False Then
        Exit Function
    End If
    
    If gstr�շ���� = "" Then
        strSQL = "Select ����,���� as ��� from �շ���Ŀ��� Where ����<>'1' Order by ���"
    Else
        strSQL = "" & _
        "   Select /*+ RULE */   A.����,A.���� as ��� " & _
        "   From �շ���Ŀ��� A," & _
        "          (Select Column_Value From Table(Cast(f_str2list([1]) As Zltools.t_strlist))) J " & _
        "   Where A.����=J. Column_Value " & _
        "   Order by ���"
    End If
    Set mrsClass = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(gstr�շ����, "'", ""))
    
    If mrsClass.EOF Then
        MsgBox "û�����ÿ��õ��շ����,�����ڱ��ز��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    '��ֻ��һ�ֿ�ѡ�շ����ʱ,�����û�ѡ��
    mblnOne = (mrsClass.RecordCount = 1)
    If InStr(gstr�շ����, "'5'") > 0 Or InStr(gstr�շ����, "'6'") > 0 _
        Or InStr(gstr�շ����, "'7'") > 0 Or gstr�շ���� = "" Then
        mlngҩƷ���ID = ExistIOClass(9)
        If mlngҩƷ���ID = 0 Then
            MsgBox "����ȷ���������ݵ�������,���ȵ����������������ã�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If InStr(gstr�շ����, "'4'") > 0 Or gstr�շ���� = "" Then
        mlng�������ID = ExistIOClass(41)
        If mlng�������ID = 0 Then
            MsgBox "����ȷ�����ĵ��ݵ�������,���ȵ����������������ã�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    'ִ�в���
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,A.����,B.��������,B.������� " & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID and B.������� IN(2,3) " & _
        " Order by B.�������,A.����"
    Set mrsUnit = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrsUnit, strSQL, Me.Caption)
    If mrsUnit.EOF Then
        MsgBox "û�г�ʼ��������Ϣ,�����޷�����ִ�в��š����ȵ����Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Curdate = zlDatabase.Currentdate
    'ȡ��ǰʱ��:33744
    If mbln���� And mstr���ת��ʱ�� <> "" Then
        txtDate.Text = Format(CDate(mstr���ת��ʱ��) - 1 / 24 / 60, "yyyy-mm-dd HH:MM:SS")
    Else
        txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
     End If
    '�Զ�ʶ��Ӱ�
    If mbytInState <> 2 And mstrInNO = "" Then
        If OverTime(Curdate) Then chk�Ӱ�.Value = Checked
    End If
    
    If mbytInState = 0 Then Set mrsWarn = GetUnitWarn
    Set mrsInfo = New ADODB.Recordset
    
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetLastDeptID(ByVal str��� As String, ByVal lngRow As Long, ByVal strDeptIDs As String) As Long
'���ܣ���ȡ����������ͬ�����Ŀ��ִ�п���ID
    Dim i As Long
    
    For i = lngRow - 1 To 1 Step -1
        If mobjBill.Details(i).�շ���� = str��� _
            And mobjBill.Details(i).ִ�в���ID <> 0 Then
            If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).ִ�в���ID & ",") > 0 Then
                GetLastDeptID = mobjBill.Details(i).ִ�в���ID
                Exit Function
            End If
        End If
    Next
    
    '�������������,��ȡ��������������ƥ���ִ�п���
    If str��� = "4" Then
        For i = lngRow - 1 To 1 Step -1
            If mobjBill.Details(i).ִ�в���ID <> 0 Then
                If InStr("," & strDeptIDs & ",", "," & mobjBill.Details(i).ִ�в���ID & ",") > 0 Then
                    GetLastDeptID = mobjBill.Details(i).ִ�в���ID
                    Exit Function
                End If
            End If
        Next
    End If
End Function

Private Sub FillBillComboBox(lngRow As Long, lngCol As Long, Optional blnEnter As Boolean)
'���ܣ����ݵ��������������б������
'������blnEnter=�Ƿ񰴽�����д���,����ִ�п��ұ��ֲ���
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String, bln��ʿ As Boolean
    Dim strSQL As String, strIDs As String, i As Long
    Dim lng����ID As Long, lng����ID As Long, j As Long
    Dim bln��ҩ��� As Boolean '�Ƿ����������ҩ���
    
    Bill.Clear
    
    On Error GoTo errHandle
    
    Select Case Bill.TextMatrix(0, lngCol)
        Case "���"
            Call GetOperatorInfo(mrs������, mobjBill.������, bln��ʿ)
            mrsClass.Filter = 0
            If mrsClass.RecordCount <> 0 Then
                mrsClass.MoveFirst
                j = 1
                For i = 1 To mrsClass.RecordCount
                    '��ʿ���:����
                    If Not (bln��ʿ And InStr(",E,M,4,", mrsClass!����) = 0) Then
                        Bill.AddItem j & "-" & mrsClass!���
                        Bill.ItemData(Bill.NewIndex) = Asc(mrsClass!����)  '����������ASCII��
                        j = j + 1
                    End If
                    mrsClass.MoveNext
                Next
            End If
            Bill.cboStyle = DropOlnyDown
        Case "ִ�п���"
            Bill.cboStyle = DropDownAndEdit
            '���ݵ�ǰ��Ŀִ�п�������,��̬���ÿ�ѡ����
            If mobjBill.Details.Count >= lngRow Then
                With mobjBill.Details(lngRow)
                    If InStr(",4,5,6,7,", .�շ����) > 0 Then
                        Call GetWorkUnit(.�շ�ϸĿID, .�շ����)
                        If mrsWork.RecordCount > 0 Then
                            'ȡ��һ��ҩ��ҩ��
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                strIDs = strIDs & "," & mrsWork!ID
                                mrsWork.MoveNext
                            Next
                            If Not blnEnter Then '�������ʱ������ȷ��ֵ����
                                lng����ID = GetLastDeptID(.�շ����, lngRow, Mid(strIDs, 2))
                            End If
                            If lng����ID = 0 Then lng����ID = .ִ�в���ID
                            
                            'ȷ����ǰ�е�ҩ��
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                Bill.AddItem IIf(zlIsShowDeptCode, mrsWork!���� & "-", "") & mrsWork!����
                                Bill.ItemData(Bill.NewIndex) = mrsWork!ID
                                If mrsWork!ID = lng����ID Then Bill.ListIndex = Bill.NewIndex
                                mrsWork.MoveNext
                            Next
                        End If
                    Else
                        Bill.TextMatrix(lngRow, lngCol) = ""
                        
                        lng����ID = mobjBill.����ID
                        If lng����ID = 0 Then lng����ID = Get��������ID
                        
                        lng����ID = mobjBill.����ID
                        If lng����ID = 0 Then lng����ID = Get����ID(lng����ID)
                        If lng����ID = 0 Then lng����ID = lng����ID
                        
                        '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
                        Select Case .Detail.ִ�п���
                            Case 0 '����ȷ
                                mrsUnit.Filter = 0
                                '101736,�ֹ�����ȱʡִ�п���
                                If Get������Դ = 2 And Not blnEnter Then
                                    '1 ������Ŀѡ���Ҵ���ȱʡ��ִ�п��ҵ� ������Ŀ��ִ�в���ID
                                    '   (���ﲻ�����ǳ�����Ŀ)
                                    '2 �շ���Ŀ.ȱʡ����(�ֹ�����ȱʡִ�п���)
                                    strSQL = "Select a.ִ�п���id" & vbNewLine & _
                                            " From �շ�ִ�п��� A, ���ű� C" & vbNewLine & _
                                            " Where a.ִ�п���id + 0 = c.Id And a.�շ�ϸĿid = [1]" & vbNewLine & _
                                            "       And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)" & vbNewLine & _
                                            "       And (c.վ�� = '" & gstrNodeNo & "' Or c.վ�� Is Null)" & vbNewLine & _
                                            "       And a.������Դ = [2] And a.��������id Is Null"
                                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .�շ�ϸĿID, 2)
                                    If Not rsTmp.EOF Then lng����ID = Val(Nvl(rsTmp!ִ�п���ID))
                                    '3 ���˿���
                                    If lng����ID = 0 Then lng����ID = mobjBill.����ID
                                    '4 ��������
                                    If lng����ID = 0 Then lng����ID = Get��������ID
                                    '5 ����Ա��������ID
                                    If lng����ID = 0 Then lng����ID = UserInfo.����ID
                                End If
                            Case 1 '���˿���
                                mrsUnit.Filter = "ID=" & lng����ID & " Or ID=" & .ִ�в���ID
                            Case 2 '���˲���
                                mrsUnit.Filter = "ID=" & lng����ID & " Or ID=" & .ִ�в���ID
                            Case 3 '����Ա����
                                mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                            Case 4 'ָ������
                                strSQL = "" & _
                                "   Select Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                                "   From �շ�ִ�п��� A,���ű� C" & _
                                "   Where A.�շ�ϸĿID=[1]��And A.ִ�п���ID+0=C.ID " & _
                                "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                                "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
                                "       And (A.������Դ is NULL Or A.������Դ=[2])" & _
                                "       And (A.��������ID is NULL Or A.��������ID=[3])" & _
                                " Order by Decode(A.������Դ,Null,2,1)" 'Ĭ�Ͽ�������
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .�շ�ϸĿID, Get������Դ, lng����ID)
                                If Not rsTmp.EOF Then
                                    For i = 1 To rsTmp.RecordCount
                                        strTmp = strTmp & "ID=" & rsTmp!ִ�п���ID & " OR "
                                        rsTmp.MoveNext
                                    Next
                                    strTmp = strTmp & "ID=" & .ִ�в���ID & " OR "
                                    strTmp = Left(strTmp, Len(strTmp) - 4)
                                    mrsUnit.Filter = strTmp
                                Else
                                    mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                                End If
                            Case 5 'Ժ��ִ��(Ԥ��,������δ��)
                            Case 6 '�����˿���
                               mrsUnit.Filter = "ID=" & Get��������ID & " Or ID=" & .ִ�в���ID
                        End Select
                        If mrsUnit.EOF Then mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                        If Not mrsUnit.EOF Then
                            For i = 1 To mrsUnit.RecordCount
                                strTmp = IIf(zlIsShowDeptCode, mrsUnit!���� & "-", "") & mrsUnit!����
                                '���˺�:28947
                                If zlCboFindItem(Bill.cboObj, Val(Nvl(mrsUnit!ID))) = False Then
                                'If Not (SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                    Bill.AddItem strTmp
                                    Bill.ItemData(Bill.ListCount - 1) = mrsUnit!ID
                                    
                                    '����ȱʡִ�п���
                                    If Not blnEnter Then '�������ʱ������ȷ��ֵ����
                                        If lngRow = 1 Then
                                            If mrsUnit!ID = lng����ID Then Bill.ListIndex = Bill.NewIndex
                                        ElseIf lngRow > 1 Then
                                            '����һ�з�ҩƷ��ͬ
                                            If mrsUnit!ID = mobjBill.Details(lngRow - 1).ִ�в���ID And mobjBill.Details(lngRow - 1).Detail.ִ�п��� = .Detail.ִ�п��� _
                                                And InStr(",5,6,7,", mobjBill.Details(lngRow - 1).�շ����) = 0 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            ElseIf mrsUnit!ID = lng����ID And Bill.ListIndex = -1 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            End If
                                        End If
                                    End If
                                End If
                                mrsUnit.MoveNext
                            Next
                            
                            If Not blnEnter And .Detail.ִ�п��� = 4 Then    'ִ�п���Ϊָ�����ҵ�,ȱʡΪ����Ա���ڿ���
                                For i = 0 To Bill.ListCount - 1
                                    If Bill.ItemData(i) = UserInfo.����ID Then Bill.ListIndex = i: Exit For
                                Next
                            End If
                            
                            If Bill.ListIndex = -1 Then '���û����ȡ���е�ִ�п���
                                For i = 0 To Bill.ListCount - 1
                                    If Bill.ItemData(i) = .ִ�в���ID Then Bill.ListIndex = i: Exit For
                                Next
                            End If
                        End If
                        
                        If Bill.ListIndex = -1 And Bill.ListCount > 0 Then Bill.ListIndex = 0
                    End If
                    
                    If Bill.ListIndex <> -1 Then
                        .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                        Bill.TextMatrix(lngRow, lngCol) = Bill.List(Bill.ListIndex)
                    Else
                        .ִ�в���ID = 0
                    End If
                End With
            End If
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitFace()
'���ܣ����ݱ�Ҫ��ɵĹ������ý��沼��
    Dim arrHead() As String, i As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim blnStatusIn As Boolean
    
    '27383
    With cboִ������
        .Clear
        .AddItem "����"
        .ItemData(.NewIndex) = 0: .ListIndex = .NewIndex
        .AddItem "��Ժ��ҩ"
        .ItemData(.NewIndex) = 3
        .AddItem "��ȡҩ"
        .ItemData(.NewIndex) = 4
    End With
            
    '���õ��ݱ��ʽ
    With Bill
        .Font.Size = 10.5
        .CboFont.Size = 11
        .TxtEditFont.Size = 11
        
        arrHead = Split(STR_HEAD, ";")
        .Cols = UBound(arrHead) + 1
        
        .MsfObj.FixedCols = 1
        .MsfObj.ScrollBars = flexScrollBarVertical
        .LocateCol = BillCol.��Ŀ
        .PrimaryCol = BillCol.��Ŀ
        .MsfObj.ColAlignmentFixed(0) = 4
        .TextMatrix(1, BillCol.��) = 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
                
        If mbytInState = 0 And gbytBilling <> 2 Then
            .ColData(BillCol.��) = BillColType.UnFocus
            
            .ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus)
            If mblnOne Then .ColData(BillCol.���) = BillColType.UnFocus
            
            .ColData(BillCol.��Ŀ) = BillColType.CommandButton  '��Ŀ����,��Ť��ѡ
            .ColData(BillCol.����) = BillColType.Text '��/������
            '���˺�:27990 2010-02-22 17:15:47
            .ColData(BillCol.��Ʒ��) = BillColType.UnFocus    '��Ʒ������
            .ColData(BillCol.���) = BillColType.UnFocus    '�������
            .ColData(BillCol.��λ) = BillColType.UnFocus  '��λ����
            .ColData(BillCol.����) = BillColType.UnFocus  '����ȱʡ����(=1),�����Ϊ��ҩʱ,��Ϊ����(4)(��ֵ,һ��ȫ��)
            .ColData(BillCol.����) = BillColType.UnFocus '����ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
            .ColData(BillCol.Ӧ�ս��) = BillColType.UnFocus  'Ӧ�ս������
            .ColData(BillCol.ʵ�ս��) = BillColType.UnFocus   'ʵ�ս������
            .ColData(BillCol.ִ�п���) = BillColType.ComboBox 'Ĭ��ȡ�������һ���һ����
            .ColData(BillCol.��־) = BillColType.UnFocus '��־ȱʡ����,��Ϊ����ʱ,��Ϊ��ѡ(-1)
            .ColData(BillCol.����) = BillColType.UnFocus  '����ȱʡ����
        End If
        .SetColColor BillCol.���, &HE7CFBA
        .SetColColor BillCol.��Ŀ, &HE7CFBA
        .SetColColor BillCol.����, &HE7CFBA
        .SetColColor BillCol.ִ�п���, &HE7CFBA
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.��־, &HE0E0E0
        
        ReDim marrColData(.Cols - 1)
        For i = 0 To .Cols - 1
            marrColData(i) = .ColData(i)
        Next
        
        If mbytInState = 3 Then .AllowAddRow = False
    End With
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name & mbytInState)
    If gTy_System_Para.bytҩƷ������ʾ <> 2 Then
        '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
        Bill.ColWidth(BillCol.��Ʒ��) = 0
    Else
        If Bill.ColWidth(BillCol.��Ʒ��) = 0 Then
             Bill.ColWidth(BillCol.��Ʒ��) = GetOrigColWidth(BillCol.��Ʒ��)
        End If
    End If
    
    Me.KeyPreview = True
    Set mobjBrushCheck = New clsBrushCardInput
    mobjBrushCheck.OnlyLegalCardNo = False
'    mobjCard.���ų��� = 18
'    mobjCard.������С���� = 8
'    mobjCard.������Ч�ַ� = Asc("=")
    'mobjCard.���Ž����� = Asc("=")
    'mobjCard.ˢ�������� = 13
    'mobjCard.�������Ĺ��� = "1-3"
'    mobjCard.������Ч�ַ� = "0" '��������(0-�����ַ�,1-����,2-��ĸ;3-���ֻ���ĸ;4-ָ���ַ�)|Ascii��1��Ascii��2....
    Call mobjBrushCheck.InitCompents(Me, Bill, mobjCard)
    
    Call SetMoneyList
    
    IDKind.Enabled = (mbytInState = 0 And mstrInNO = "")
    
    '��ȡ����ƥ�䷽ʽ
    sta.Panels("MedicareType").Visible = mbytInState = 0
    sta.Panels("PY").Visible = mbytInState = 0 And gbln�����л� '35242
    sta.Panels("WB").Visible = mbytInState = 0 And gbln�����л�
    sta.Panels("BarCode").Visible = mbytInState = 0
    If mbytInState = 0 Then
        '����ƥ�䷽ʽ��0-ƴ��,1-���,2-����
        If gbytCode = 0 Then
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrRaised
        ElseIf gbytCode = 1 Then
            sta.Panels("PY").Bevel = sbrRaised
            sta.Panels("WB").Bevel = sbrInset
        Else
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrInset
        End If
        mblnShowBarCode = Val(zlDatabase.GetPara("�ϴ�ѡ���������", glngSys, 1150))
        sta.Panels("BarCode").Bevel = IIf(Not mblnShowBarCode, sbrRaised, sbrInset)
        sta.Panels("BarCode").ToolTipText = IIf(Not mblnShowBarCode, "��ʾ���������", "�������������")
        Call ShowAndHideBarCodeInput
    End If
        
    'mbytUseType:=
    If gbln���뷢ҩ Or Not (mbytInState = 0) Then
        sta.Panels("Drugstore").Visible = False
    End If
            
    '����
    Select Case gbytBilling
        Case 0
            lblTitle.Caption = gstrUnitName & "סԺ���ʵ�"
        Case 1
            lblTitle.Caption = gstrUnitName & "סԺ���ʵ�(����)"
        Case 2
            lblTitle.Caption = gstrUnitName & "סԺ���ʵ�(���)"
    End Select
    
    If mbln���� Then
        If mlng��ҳID <> 0 Then
            strSQL = "Select ��ǰ����ID,��Ժ����ID,��Ժ���� From ������ҳ Where ����ID = [1] And ��ҳID = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
            If Not rsTemp.EOF Then
                If mlngDeptID = 0 Then
                    mlngDeptID = Val(Nvl(rsTemp!��Ժ����ID))
                End If
                If mlngUnitID = 0 Then
                    mlngUnitID = Val(Nvl(rsTemp!��ǰ����ID))
                End If
                blnStatusIn = IsNull(rsTemp!��Ժ����)
            End If
        End If
        If blnStatusIn Or mlng��ҳID = 0 Or rsTemp.EOF Then
            lblTitle.Caption = lblTitle.Caption & "(" & "����" & ")"
        Else
            lblTitle.Caption = lblTitle.Caption & "(��" & mlng��ҳID & "�β���" & ")"
        End If
    End If
    
    txtӦ��.Text = gstrDec: txtʵ��.Text = gstrDec
    
    cmdSelWholeSet.Visible = (gbytBilling = 0 Or gbytBilling = 1) And mbytInState = 0
    cmdSaveWholeSet.Visible = zlStr.IsHavePrivs(mstrPrivsOpt, "���ӳ�����Ŀ")
    Select Case mbytInState
        Case 0 'ִ��
            If mstrInNO <> "" Or _
            (InStr(mstrPrivsOpt, ";ҩƷ����;") = 0 _
                And InStr(mstrPrivsOpt, ";��������;") = 0 _
                And InStr(mstrPrivsOpt, ";��������;") = 0) Then
                chkCancel.Visible = False
            End If
            Select Case gbytBilling
                Case 0, 1 'ִ�м��ʡ�����
                    Call SetShowCol
                    '��ͨ���ʺͿ��ҷ�ɢ���ʻ򻮼�ʱ,�������޸Ĳ���������������ҩ�䷽,���ʱ��ݲ��ṩ
                    cmd�䷽.Visible = (mbytUseType = 0 Or mbytUseType = 1 Or mbytUseType = 2)
                    txtPatient.Enabled = (mstrInNO = "")
                    cboִ������.Visible = True
                    lblִ������.Visible = True
                Case 2 'ִ�����
                    Call SetDisible
                    cboNO.Locked = False
                    fraInfo.Enabled = False
                    fraUnit.Enabled = False
                    fraAppend.Enabled = False
                    fraDrawDept.Enabled = False
                    cmdSaveWholeSet.Left = fraTitle.Left + 50
            End Select
        Case 1 '����
            Call SetDisible
            chkCancel.Visible = False
            If mblnDelete Then lblFlag.Visible = True
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            fraAppend.Enabled = False
            fraDrawDept.Enabled = False
            cmdOK.Visible = False
            cmdCancel.Caption = "�˳�(&X)"
        Case 2 '����
            Call SetDisible
            txtDate.Enabled = True
            chkCancel.Visible = False
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            chk����.Enabled = False
            fraDrawDept.Enabled = False
        Case 3 '����
            Call SetDisible
            chkCancel.Visible = False
            fraTitle.Enabled = False
            fraInfo.Enabled = False
            fraUnit.Enabled = False
            fraAppend.Enabled = False
            fraDrawDept.Enabled = False
            Call ShowDeleteCol(True)
            Bill.Active = True      '����������
    End Select
    
    If fraTitle.Enabled = False Then
        Set cmdSaveWholeSet.Container = Me
        cmdSaveWholeSet.Left = fraTitle.Left + 50
        cmdSaveWholeSet.Top = fraTitle.Height - cmdSelWholeSet.Height * 1.6
    End If
    
    If mbytInState <> 0 Then
        lblPreNO.Visible = False: txtPreNO.Visible = False
        lblӦ��.Top = lblӦ��.Top + txtPreNO.Height / 2
        txtӦ��.Top = txtӦ��.Top + txtPreNO.Height / 2
        lblʵ��.Top = lblʵ��.Top + txtPreNO.Height * 0.75
        txtʵ��.Top = txtʵ��.Top + txtPreNO.Height * 0.75
    End If
    
    '�������������뿪����λ��
    If gblnFromDr Then
        Call ExChangeLocate(cbo��������, cbo������)
        Call ExChangeLocate(lbl��������, lbl������)
        cbo��������.TabStop = False
    End If
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
'��������Ϊ�����޸�״̬
    cboNO.Locked = Not bln
    txtPatient.Locked = Not bln
    cbo��������.Locked = Not bln
    cbo������.Locked = Not bln
    
    chk�Ӱ�.Enabled = bln
    cboBaby.Enabled = bln
    txtDate.Enabled = bln
    Bill.Active = bln
    
    If Not bln Then
        txtPatient.BackColor = &HE0E0E0
        txtOld.BackColor = &HE0E0E0
    Else
        txtPatient.BackColor = &HFFFFFF
        txtOld.BackColor = &HFFFFFF
    End If
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Then Exit Sub
    Call IDKind.ActiveFastKey
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    
    If txtPatient.Locked Then Exit Sub
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        With frmPatiSelect
            If (mbytUseType = 0 Or mbytUseType = 1) Then
                .mlngUnitID = mlngUnitID
            Else
                .mlngUnitID = mlngDeptID
            End If
            .mbytUseType = mbytUseType
            .mstrPrivs = mstrPrivs
            Set .mfrmParent = Me
            .Show 1, Me
        End With
    Else
        If IDKind.GetCurCard.���� Like "����*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        ElseIf IDKind.GetCurCard.���� = "�����" Or IDKind.GetCurCard.���� = "סԺ��" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
        Else
            txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
        End If
    End If
    
    
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        
      
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        ElseIf IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
            KeyAscii = 0
            
            'ˢ�²�����Ϣ:"-����ID"
            Call GetPatient(IDKind.GetCurCard, txtPatient.Tag, False)
            If mrsInfo.State = 0 Then   '��������ʱ�����ܴ�ʱ����������˷��ã�������Աû��"��Ժδ��ǿ�Ƽ���"Ȩ�ޣ�����������
                txtPatient.Text = "": txtOld.Text = ""
                txt����.Text = "": txtסԺ��.Text = ""
                Exit Sub
            End If
            
            '����:27658
            If "-" & Val(Nvl(mrsInfo!����ID)) <> txtPatient.Tag Then
                txtPreNO.Text = ""
            End If
            
            'ˢ�²���Ԥ������Ϣ
            curTotal = GetBillTotal(mobjBill)
            Set rsTmp = GetMoneyInfo(mrsInfo!����ID, IIf(gbytBilling = 0, mcurModiMoney, 0), True, 2)
            If Not rsTmp Is Nothing Then
                cmdOK.Tag = rsTmp!Ԥ�����
                cmdCancel.Tag = rsTmp!�������
                txtʵ��.Tag = rsTmp!Ԥ����� - rsTmp!�������
            Else
                cmdOK.Tag = 0: cmdCancel.Tag = 0: txtʵ��.Tag = 0
            End If
            '����ʱ��ʾ���㵱ǰ���ݷ���,�����۱���Ҫ��
            strInfo = GetPatientDue(Val(mrsInfo!����ID))
            ' If Val(strInfo) <> 0 Then sta.Panels(3).Text = sta.Panels(3).Text & "/Ӧ�տ�:" & Format(strInfo, "0.00")
            '����:30604
            Call SetStatuPatiInfor(Val(cmdOK.Tag), Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), Val(txtʵ��.Tag) - IIf(gbytBilling = 0, curTotal, 0), Val(strInfo))
            
            Call LoadPatientBaby(cboBaby, mrsInfo!����ID, mrsInfo!��ҳID)
                                    
            If Not mblnValid Then Bill.SetFocus
            txtPatient.PasswordChar = ""
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
            Exit Sub
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, txtPatient.Text)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMsg As Boolean, blnICCard As Boolean, blnIDCard As Boolean
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    
    '�ɶ���������
    If mobjBill.Details.Count = 0 Then
        Call ClearMoney
        txtʵ��.Text = gstrDec: txtӦ��.Text = gstrDec
    End If
        
    '��ȡ������Ϣ
    If Not (mbytInState = 0 And mbytUseType = 1 And sta.Panels(2) Like "��һ��*") Then
        sta.Panels(2) = ""
    End If
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.���� Like "*���֤*" And objCard.ϵͳ Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    
    If Not GetPatient(objCard, strInput, blnCard, blnMsg) Then
        chk����.Value = 0: chk����.Visible = False
        If blnCard Then
            If Not blnMsg Then MsgBox "����ȷ��������Ϣ�������Ƿ���ȷˢ����", vbInformation, gstrSysName
            txtPatient.Text = "": txtOld.Text = ""
            txt����.Text = "": txtסԺ��.Text = ""
            Exit Sub
        End If
        If Not blnMsg Then MsgBox "���ܶ�ȡ������Ϣ��", vbInformation, gstrSysName
        zlControl.TxtSelAll txtPatient
        If mstrInNO = "" Then
            txtOld.Text = "": txt����.Text = "": txtסԺ��.Text = ""
        End If
        Exit Sub
    End If
    '��ȡ�ɹ�
     '���￨������
     If (objCard.���� Like "IC��*" Or objCard.���� Like "*���֤*") And objCard.ϵͳ = True And blnCard Then blnCard = False
     If Mid(gstrCardPass, 6, 1) = "1" _
        And (blnCard Or (blnICCard And mstrPassWord <> "") _
               Or (blnIDCard And mstrPassWord <> "") Or (IDKind.GetCurCard.�ӿ���� <> 0 And mstrPassWord <> "")) Then
         If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!����, mrsInfo!�Ա�, "" & mrsInfo!����) Then
             Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
         End If
     End If

    If Not IsNull(mrsInfo!����) Then
        chk����.Value = 0: chk����.Visible = True
        MCPAR.�������� = gclsInsure.GetCapability(support��������, mrsInfo!����ID, mrsInfo!����)
        MCPAR.�����ϴ� = gclsInsure.GetCapability(support�����ϴ�, mrsInfo!����ID, mrsInfo!����)
        MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, mrsInfo!����ID, mrsInfo!����)
        MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, mrsInfo!����ID, mrsInfo!����)
        MCPAR.ʵʱ��� = gclsInsure.GetCapability(supportʵʱ���, mrsInfo!����ID, mrsInfo!����)
    Else
        chk����.Value = 0: chk����.Visible = False
    End If
         
    '����:27658
    If Val(Nvl(mrsInfo!����ID)) <> mlng����ID Then
        txtPreNO.Text = ""
    End If
    
    If mbytUseType = 1 And mrsInfo!����ID <> mlng����ID Then mlng����ID = 0

    '�Զ����ÿ�������(ͬʱ���ü��ʱ�����Ϣ),ҽ�����ʲ��˿��Ҳ�һ���ǿ�������
    If mbytUseType = 2 Then lngUnit = cbo��������.ListIndex

    If gblnFromDr Then
        If Not IsNull(mrsInfo!סԺҽʦ) Then
            cbo������.ListIndex = -1
            cbo������.ListIndex = cbo.FindIndex(cbo������, mrsInfo!סԺҽʦ, True)
        End If
    Else
        '33744
        If mbln���� Then
            If cbo��������.ListIndex >= 0 Then
                If cbo��������.ItemData(cbo��������.ListIndex) <> mlngDeptID And mlngDeptID <> 0 Then
                    cbo��������.ListIndex = cbo.FindIndex(cbo��������, mlngDeptID)
                ElseIf mlngDeptID = 0 Then
                
                    cbo��������.ListIndex = cbo.FindIndex(cbo��������, IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID))
                End If
            ElseIf mlngDeptID = 0 Then
                cbo��������.ListIndex = cbo.FindIndex(cbo��������, IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID))
            Else
                cbo��������.ListIndex = cbo.FindIndex(cbo��������, mlngDeptID)
            End If
        Else
            cbo��������.ListIndex = -1
            cbo��������.ListIndex = cbo.FindIndex(cbo��������, IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID))
        End If
        If cbo��������.ListIndex <> -1 Then
            mobjBill.��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        ElseIf mbytUseType = 2 And lngUnit <> -1 Then
            cbo��������.ListIndex = cbo.FindIndex(cbo��������, lngUnit)
        End If
    End If
            
    Call LoadPatientBaby(cboBaby, mrsInfo!����ID, mrsInfo!��ҳID)
            
    '����Ԥ������Ϣ
    curTotal = GetBillTotal(mobjBill)
    Set rsTmp = GetMoneyInfo(mrsInfo!����ID, IIf(gbytBilling = 0, mcurModiMoney, 0), True, 2)
    If Not rsTmp Is Nothing Then
        cmdOK.Tag = rsTmp!Ԥ�����
        cmdCancel.Tag = rsTmp!�������
        txtʵ��.Tag = rsTmp!Ԥ����� - rsTmp!�������
    Else
        cmdOK.Tag = 0: cmdCancel.Tag = 0: txtʵ��.Tag = 0
    End If
            
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------
    '���˺�:26952
    Dim cur��� As Currency, curItemMoney As Currency
    
    cur��� = Val(txtʵ��.Tag)
    curItemMoney = 0
    
    If gbln�����������۷��� Then cur��� = Val(txtʵ��.Tag) - GetPriceMoneyTotal(1, mrsInfo!����ID) + IIf(gbytBilling = 1, mcurModiMoney, 0)
    
    gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!���� & IIf(Nvl(mrsInfo!סԺ��) = "", "", "(סԺ��:" & mrsInfo!סԺ�� & " ����:" & mrsInfo!���� & ")"), Val("" & mrsInfo!����ID), mrsInfo!���ò���, mrsWarn, cur���, mrsInfo!���ն� - mcurModiMoney, curTotal, _
                IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), "", "", _
                 mstrWarn, , gblnPrice And (gbytBilling = 0 And mstrInNO = "" Or gbytBilling = 1), curItemMoney, True)
    '����:0;û�б���,����
    '     1:������ʾ���û�ѡ�����
    '     2:������ʾ���û�ѡ���ж�
    '     3:������ʾ�����ж�
    '     4:ǿ�Ƽ��ʱ���,����
    '     5.������ʾ���û�ѡ�����,��ֻ�������Ϊ���۵�
    If gbytWarn = 2 Or gbytWarn = 3 Then
        Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "":
        mlng����ID = 0
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Sub
    ElseIf gbytBilling = 0 And gblnPrice And mstrInNO = "" Then
          '���˺�:2010-03-15 14:59:37:27963
            If gbytWarn = 1 Or gbytWarn = 4 Then
                cmdPrice.Visible = True: cmdOK.Visible = True: Call SetButtonPlace
            ElseIf gbytWarn = 5 Then
                cmdPrice.Visible = True: cmdOK.Visible = False: Call SetButtonPlace
            End If
    End If
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����ʱ��ʾ���㵱ǰ���ݷ���,�����۱���Ҫ��
    'sta.Panels(3).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
    'sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
    'sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(Val(txtʵ��.Tag) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
    strInfo = GetPatientDue(Val(mrsInfo!����ID))
    'If Val(strInfo) <> 0 Then sta.Panels(3).Text = sta.Panels(3).Text & "/Ӧ�տ�:" & Format(strInfo, "0.00")
    
    Call SetStatuPatiInfor(Val(cmdOK.Tag), Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), Val(txtʵ��.Tag) - IIf(gbytBilling = 0, curTotal, 0), Val(strInfo))
                
    '������Ϣ
    txtPatient.Text = Nvl(mrsInfo!����)
    cboSex.ListIndex = cbo.FindIndex(cboSex, Nvl(mrsInfo!�Ա�), True)
    txtOld.Text = Nvl(mrsInfo!����)
    cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, Nvl(mrsInfo!�ѱ�), True)
    cboҽ�Ƹ���.ListIndex = cbo.FindIndex(cboҽ�Ƹ���, Nvl(mrsInfo!ҽ�Ƹ��ʽ), True)
    txt����.Text = "" & mrsInfo!����
    txtסԺ��.Text = Nvl(mrsInfo!סԺ��)
    txt������.Text = Nvl(mrsInfo!������)
    txt������.Text = Format(Nvl(mrsInfo!������), "0.00")
    '���˺�:d
    txt���˱�ע.Text = Nvl(mrsInfo!��ע)
    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    With mobjBill
        .����ID = Nvl(mrsInfo!����ID, 0)
        .��ҳID = IIf(mbln���� And mlng��ҳID <> 0, mlng��ҳID, Nvl(mrsInfo!��ҳID, 0))
        .����ID = IIf(mbln���� And mlngUnitID <> 0, mlngUnitID, Nvl(mrsInfo!����ID, 0))
        .����ID = IIf(mbln���� And mlngDeptID <> 0, mlngDeptID, Nvl(mrsInfo!����ID, 0))
        
        .���� = "" & mrsInfo!����
        .��ʶ�� = Nvl(mrsInfo!סԺ��, 0)
        .���� = Nvl(mrsInfo!����)
        .�Ա� = Nvl(mrsInfo!�Ա�)
        .���� = Nvl(mrsInfo!����)
        .�ѱ� = Nvl(mrsInfo!�ѱ�)
    End With
    If Not IsNull(mrsInfo!��Ժ����) Then
        MsgBox "��������" & vbCrLf & vbCrLf & "�ò������� " & Format(mrsInfo!��Ժ����, "yyyy-MM-dd") & " ��Ժ�����ڶԸò���ǿ�ƽ��м��ʣ�", vbInformation, gstrSysName
        txtDate.Text = Format(mrsInfo!��Ժ����, "yyyy-MM-dd HH:mm:ss")
    Else
        txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    End If
    'ȡ��ǰʱ��:33744
    If mbln���� And mstr���ת��ʱ�� <> "" Then
        txtDate.Text = Format(CDate(mstr���ת��ʱ��) - 1 / 24 / 60, "yyyy-mm-dd HH:MM:SS")
        txtDate.ForeColor = vbBlue
    End If
    
    If Not (mbytInState = 0 And mbytUseType = 1 And sta.Panels(2) Like "��һ��*") Then
        If Not IsNull(mrsInfo!��Ժ����) Then
            sta.Panels(2).Text = "��Ժ����:" & Format(mrsInfo!��Ժ����, "yyyy-MM-dd")
            strInfo = GetInsureInfo(mrsInfo!����ID)
            If strInfo <> "" Then sta.Panels(2).Text = sta.Panels(2).Text & "/�ʺ�:" & Split(strInfo, ";")(1)
        End If
    End If
    
    If mbytInState = 0 And mobjBill.Details.Count > 0 And Not mbln������۸� Then
        '���¼���۸�
        Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
    End If
            
    If gblnFromDr Then
        If cbo������.Visible And cbo������.Enabled Then cbo������.SetFocus
    Else
        If cbo��������.Visible And cbo��������.Enabled Then cbo��������.SetFocus
    End If
    '33744
    If mbln���� Then
        Call Set���˲��ѱ༭����
    End If
End Sub


Private Sub Set���˲��ѱ༭����()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ò��˲���ʱ�ı༭����
    '����:���˺�
    '����:2010-12-10 14:54:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbln���� = False Then Exit Sub
    txtPatient.Enabled = False
    cbo��������.Enabled = False
    cboSex.Enabled = False
    IDKind.Enabled = False
    chkCancel.Visible = False
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=�Ƿ���￨ˢ��
    '����:blnOutMsg-�Ѿ���ʾ,�������ⲿ����ʾ
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-03 16:54:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim strIF As String, strWhere As String
    Dim rsOutSel As ADODB.Recordset, bln���в��� As Boolean
    On Error GoTo errH
        
    'a.�Ƿ����ǿ�Ƽ���Ȩ��
    If InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        strIF = ""
    ElseIf InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 Then
        strIF = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)<>0)"
    ElseIf InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        strIF = " And ((B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3) Or Nvl(X.�������,0)=0)"
    Else
        strIF = " And B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3"
    End If
    
    'b.�Ƿ���Լ����в�������
    bln���в��� = True
    If (mbytUseType = 0 Or mbytUseType = 1) And InStr(mstrPrivs, ";���в���;") <= 0 Then
        bln���в��� = False
        If InStr(1, mstrUnitIDs, ",") = 0 Then
            strIF = strIF & " And B.��ǰ����ID+0=[3]"
        Else
            strIF = strIF & " And B.��ǰ����ID+0 IN(Select Column_Value From Table(Cast(f_num2list([4]) As zlTools.t_numlist)))"
        End If
    End If
       
    'c.�Ƿ����۲��˼���Ȩ��
    If (InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln��������) And (InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ����) Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,1,2)"
    ElseIf InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,1)"
    ElseIf InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ���� Then
        strIF = strIF & " And Nvl(B.��������,0) IN(0,2)"
    Else
        strIF = strIF & " And Nvl(B.��������,0)=0"
    End If
    
    strSQL = _
            "Select A.����ID,B.��ҳID,B.��ǰ����ID as ����ID,B.��Ժ����ID as ����ID,B.��Ժ����,B.��Ժ����," & _
            "   A.���￨��,A.����֤��,A.סԺ��,B.��Ժ���� as ����,X.�������,B.״̬, " & _
            "   nvl(B.����,A.����) as ����,nvl(b.�Ա�,A.�Ա�) as �Ա�,nvl(b.����,A.����) as ����,B.�ѱ�,B.סԺҽʦ,B.ҽ�Ƹ��ʽ," & _
            "   A.������,Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������,zl_PatiDayCharge(A.����ID) as ���ն�," & _
            "   Zl_Patiwarnscheme(B.����id, B.��ҳid) As ���ò���,B.����,Nvl(B.��������,0) as ��������,B.��������,B.��ע,B.��˱�־" & _
            " From ������Ϣ A,������ҳ B,������� X" & _
            " Where A.����ID=B.����ID And " & IIf(mbln���� And mlng��ҳID <> 0, " B.��ҳID = [5] ", "A.��ҳID=B.��ҳID") & _
            " And Nvl(B.��ҳID,0)<>0 And A.����ID=X.����ID(+) and X.����(+)=1 and X.����(+)=2 And A.ͣ��ʱ�� is NULL " & strIF
            
    If blnCard = True And objCard.���� Like "����*" Then   'ˢ��
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strWhere = strWhere & " And A.����ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strWhere = strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "/" Then  '��λ��
        '41654 And IsNumeric(Mid(strInput, 2))
        strInput = Mid(strInput, 2)
        If mlngUnitID = 0 Then '������ȷ��������ͨ������ȷ������
            Set mrsInfo = New ADODB.Recordset: Exit Function
        End If
        strSQL = _
            "Select A.����ID,B.��ҳID,B.��ǰ����ID as ����ID,B.��Ժ����ID as ����ID,B.��Ժ����,B.��Ժ����," & _
            "   A.���￨��,A.����֤��,A.סԺ��,B.��Ժ���� as ����,X.�������,B.״̬," & _
            "   nvl(B.����,A.����) as ����,nvl(b.�Ա�,A.�Ա�) as �Ա�,nvl(b.����,A.����) as  ����,B.�ѱ�,B.סԺҽʦ,B.ҽ�Ƹ��ʽ," & _
            "   A.������,Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,B.��ҳID)) ������,zl_PatiDayCharge(A.����ID) as ���ն�," & _
            "   Zl_Patiwarnscheme(B.����id, B.��ҳid) As ���ò���,B.����,Nvl(B.��������,0) as ��������,B.��������,B.��˱�־,B.��ע" & _
            " From ������Ϣ A,������ҳ B,��λ״����¼ C,������� X" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID" & _
            " And Nvl(B.��ҳID,0)<>0 And A.����ID=C.����ID And A.����ID=X.����ID(+) And X.����(+)=1 And X.����(+)=2 And A.ͣ��ʱ�� is NULL " & _
            " And C.����ID=[3] And C.����=[2] " & strIF
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��(������Ժ)
        strWhere = strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strWhere = strWhere & " And A.�����=[1]"
    Else '��������
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then
                If Not mrsInfo.EOF Then
                    If mrsInfo!���� = Trim(txtPatient.Text) Then GetPatient = True: Exit Function
                End If
            End If
        End If
        Select Case objCard.����
            Case "����", "��������￨"
                If zlSelectChargePatiFromInputName(Me, mstrPrivsOpt, strInput, bln���в���, mstrUnitIDs, gintOutDay, lng����ID, strErrMsg, txtPatient.hWnd, txtPatient.Height) = False Then
                    If strErrMsg = "" Then blnOutMsg = True: Set mrsInfo = New Recordset: Exit Function
                    If mbytUseType = 2 And InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then GoTo GoYJReadPati:
                    MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    blnOutMsg = True: Set mrsInfo = New Recordset: Exit Function
                End If
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "ҽ����"
                strInput = UCase(strInput)
                strWhere = strWhere & " And A.ҽ����=[2]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.�����=[2]"
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])"
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
                strWhere = strWhere & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    
    strSQL = strSQL & vbCrLf & strWhere
    txtPatient.ForeColor = Me.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, mlngUnitID, mstrUnitIDs, mlng��ҳID)
    
    If Not mrsInfo.EOF Then
        txtPatient.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!��������))
        
        If zlPatiIS�����ѱ�Ŀ(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID))) = True Then    '����:28725
            Set mrsInfo = New ADODB.Recordset
            Set mrsMedAudit = Nothing   'ҽ�����˱�����Ժ�ż���������
            blnOutMsg = True
            Exit Function
        End If
        If zlIsAllowFeeChange(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID)), Val(Nvl(mrsInfo!��˱�־))) = False Then
            Set mrsInfo = New ADODB.Recordset
            Set mrsMedAudit = Nothing
            blnOutMsg = True
            Exit Function
        End If
        
        If mrsInfo!����ID <> mobjBill.����ID Or mbytInState = 0 And mstrInNO <> "" Then    'ͬһ���˲����ظ���ȡ
            If GetMedPayMode("" & mrsInfo!ҽ�Ƹ��ʽ, mrsMedPayMode) = 1 Then
                Set mrsMedAudit = GetAuditRecord(mrsInfo!����ID, mrsInfo!��ҳID)
            Else
                Set mrsMedAudit = Nothing
            End If
        End If
        
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then
           mstrPassWord = Nvl(mrsInfo!����֤��)
        End If
        If mlng����ID <> mrsInfo!����ID Then mlng����ҽ�� = 0
        GetPatient = True
        Exit Function
    Else
        Set mrsMedAudit = Nothing   'ҽ�����˱�����Ժ�ż���������
    End If
    
        
    'ҽ�����Ҽ��ʣ�û�з���סԺ(��Ժ���Ժ)����,�����ﲡ�˶�
    If mbytUseType = 2 And InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
GoYJReadPati:
        '76451,Ƚ����,2014-8-19
        strSQL = _
            "Select A.����ID,Nvl(A.��ҳID,0) as ��ҳID,A.��ǰ����ID as ����ID,A.��ǰ����ID as ����ID,'' as ����,'' as ����," & _
            " A.��Ժʱ�� as ��Ժ����,A.���￨��,A.����֤��,A.סԺ��,A.��ǰ���� as ����,A.����,A.�Ա�,A.����," & _
            " A.��Ժʱ�� as ��Ժ����,A.�ѱ�,Decode(A.������,null,A.������,Zl_Patientsurety(A.����ID,null)) ������,Zl_Patiwarnscheme(A.����id) As ���ò���," & _
            "A.������,NULL as סԺҽʦ,A.ҽ�Ƹ��ʽ,zl_PatiDayCharge(A.����ID) as ���ն�,A.����,-1 as ��������,'' ��ע" & _
            " From ������Ϣ A Where A.ͣ��ʱ�� is NULL "
        If blnCard Or IDKind.IDKind = IDKind.GetKindIndex("���￨") Then '���￨��
            strSQL = strSQL & " And A.���￨��=[2]"
        ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
            strSQL = strSQL & " And A.����ID=[1]"
        ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����(ҽ������)
            strSQL = strSQL & " And A.�����=[1]"
        Else '��������
            Select Case IDKind.IDKind
                Case IDKind.GetKindIndex("����")
                  strSQL = strSQL & " And A.����=[2]"
                Case IDKind.GetKindIndex("ҽ����")
                    strSQL = strSQL & " And A.ҽ����=[2]"
                Case IDKind.GetKindIndex("���֤��")
                    strSQL = strSQL & " And A.���֤��=[2]"
                Case IDKind.GetKindIndex("IC����")
                    strSQL = strSQL & " And A.IC����=[2]"
                Case IDKind.GetKindIndex("�����")
                    If Not IsNumeric(strInput) Then strInput = "0"
                    strSQL = strSQL & " And A.�����=[2]"
                Case IDKind.GetKindIndex("סԺ��")
                    If Not IsNumeric(strInput) Then strInput = "0"
                    strSQL = strSQL & " And A.סԺ��=[2]"
            End Select
        End If
        
        'Val(Mid(strInput, 2)):29787
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
        If Not mrsInfo.EOF Then
            If zlPatiIS�����ѱ�Ŀ(Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID))) = True Then
                Set mrsInfo = New ADODB.Recordset
                blnOutMsg = True
                Exit Function
            End If
            If mlng����ID <> mrsInfo!����ID Then mlng����ҽ�� = 0
            GetPatient = True
            Exit Function
        End If
        Set mrsInfo = New ADODB.Recordset
        Exit Function
    End If
    
    Set mrsMedAudit = Nothing   'ҽ�����˱�����Ժ�ż���������'
    Set mrsInfo = New ADODB.Recordset
    If strWhere = "" Then Exit Function '������������ֱ���˳�
    
    
    
    'δ�ҵ����ˣ���Ҫ�Ըò��˵ľ��������Ϣ������ʾ
    strSQL = _
    " Select A.����ID,B.��ҳID,B.��ǰ����ID as ����ID,B.��Ժ����ID as ����ID,a.��Ժ,B.��Ժ����,B.��Ժ����,X.�������,B.״̬, " & _
    "       nvl(B.����,A.����) as ����,nvl(b.�Ա�,A.�Ա�) as �Ա�,nvl(b.����,A.����) as ����,B.�ѱ�,Nvl(B.��������,0) as ��������,B.��������" & _
    " From ������Ϣ A,������ҳ B,������� X" & _
    " Where A.����ID=B.����ID And " & IIf(mbln���� And mlng��ҳID <> 0, " B.��ҳID = [3] ", "A.��ҳID=B.��ҳID") & _
    "   And Nvl(B.��ҳID,0)<>0 And A.����ID=X.����ID(+) and X.����(+)=1 and X.����(+)=2 And A.ͣ��ʱ�� is NULL " & strWhere
    
    Set rsOutSel = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, mlng��ҳID)
    If rsOutSel.EOF Then Exit Function
    
    '1.�������
    If (mbytUseType = 0 Or mbytUseType = 1) And InStr(mstrPrivs, ";���в���;") <= 0 Then
        If InStr(1, "," & mstrUnitIDs & ",", "," & Val(rsOutSel!����ID) & ",") = 0 Then
            MsgBox "����:��" & Nvl(rsOutSel!����) & "�������㸺��Ĳ���,���ܶԸò��˽��м��˲���!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    End If
    
    '2.���۲��˼��(�Ƿ����۲��˼���Ȩ��)
    If (InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln��������) And (InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ����) Then
        '0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
    ElseIf InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
        If Val(Nvl(rsOutSel!��������)) = 2 Then
            MsgBox "����:��" & Nvl(rsOutSel!����) & "��ΪסԺ���۲���,�㲻�߱���סԺ���ۼ��ʡ�Ȩ��,���ܶԸò��˽��м��˲���!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    ElseIf InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ���� Then
        If Val(Nvl(rsOutSel!��������)) = 1 Then
            MsgBox "����:��" & Nvl(rsOutSel!����) & "��Ϊ�������۲���,�㲻�߱����������ۼ��ʡ�Ȩ��,���ܶԸò��˽��м��˲���!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    Else
        If Val(Nvl(rsOutSel!��������)) <> 0 Then
            MsgBox "����:��" & Nvl(rsOutSel!����) & "��Ϊ" & IIf(Val(Nvl(rsOutSel!��������)) = 1, "����", "סԺ") & "���۲���,�㲻�߱��������סԺ ���ۼ��ʡ�Ȩ��,���ܶԸò��˽��м��˲���!", vbInformation + vbOKOnly, gstrSysName
            blnOutMsg = True
            Exit Function
        End If
    End If
    '124007
    If InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        strErrMsg = ""
    ElseIf InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 Then
        If Not (Val(Nvl(rsOutSel!״̬)) <> 3 And IsNull(rsOutSel!��Ժ����) Or Val(Nvl(rsOutSel!�������)) <> 0) Then
              
                If Val(Nvl(rsOutSel!״̬)) = 3 And IsNull(rsOutSel!��Ժ����) Then
                    strErrMsg = "�����Ѿ�Ԥ��Ժ�����ܶԲ��˽��м��˲���!"
                Else
                    strErrMsg = "������" & Format(rsOutSel!��Ժ����, "yyyy��mm��DD��") & " ��Ժ�����ܶԲ��˽��м��˲���!"
                End If
        End If
    ElseIf InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
        If Not (Val(Nvl(rsOutSel!״̬)) <> 3 And IsNull(rsOutSel!��Ժ����) Or Val(Nvl(rsOutSel!�������)) = 0) Then
                If Val(Nvl(rsOutSel!״̬)) = 3 And IsNull(rsOutSel!��Ժ����) Then
                strErrMsg = "�����Ѿ�Ԥ��Ժ�����ܶԲ��˽��м��˲���!"
                Else
                strErrMsg = "������" & Format(rsOutSel!��Ժ����, "yyyy��mm��DD��") & " ��Ժ�����ܶԲ��˽��м��˲���!"
                End If
        End If
    Else
        If Not (Val(Nvl(rsOutSel!״̬)) <> 3 And IsNull(rsOutSel!��Ժ����)) Then
            If Val(Nvl(rsOutSel!״̬)) = 3 And IsNull(rsOutSel!��Ժ����) Then
                strErrMsg = "�����Ѿ�Ԥ��Ժ�����ܶԲ��˽��м��˲���!"
            Else
                strErrMsg = "������" & Format(rsOutSel!��Ժ����, "yyyy��mm��DD��") & " ��Ժ�����ܶԲ��˽��м��˲���!"
            End If
        End If
    End If
    
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbInformation, gstrSysName
        Set mrsMedAudit = Nothing   'ҽ�����˱�����Ժ�ż���������'
        blnOutMsg = True
        Exit Function
    End If
    
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

Private Sub CalcMoneys(Optional lngRow As Long = 0)
'���ܣ���������¼���ָ���л������еĽ��
'������lngRow=ָ����,Ϊ0��ʾ����������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long
    Dim strMainRows As String
    Dim bln��������ۿ� As Boolean
    
    If mobjBill.Details.Count = 0 Then Exit Sub
    
    For i = IIf(lngRow = 0, 1, lngRow) To IIf(lngRow = 0, mobjBill.Details.Count, lngRow)
        
        bln��������ۿ� = False
        If gbln��������ۿ� Then                    '����������ηѱ�,����ܼ����ۿ۲�����Ч,�����ܼ���
            If mobjBill.Details(i).�������� > 0 Then    '����
                bln��������ۿ� = Not mobjBill.Details(mobjBill.Details(i).��������).Detail.���ηѱ�
                If bln��������ۿ� And lngRow <> 0 Then strMainRows = "," & mobjBill.Details(i).��������      '��������һ�е�ʱ��
            Else
                If CheckItemHaveSub(i) Then                          '����������
                     bln��������ۿ� = Not mobjBill.Details(i).Detail.���ηѱ�
                     If bln��������ۿ� Then strMainRows = strMainRows & "," & i  'һҳ�����ж��������,�ȼ�¼�����к�,���������������ۿ�
                End If
            End If
        End If
                    
        Call CalcMoney(i, bln��������ۿ�)
    Next
    
    '������������,������bln��������ۿ۱���,��Ϊ�������������Ǵ������ʱ�Ѹı�
    If gbln��������ۿ� Then
        For i = 1 To UBound(Split(strMainRows, ","))
            Call Calc��������ʵ��(Split(strMainRows, ",")(i))
        Next
    End If
    
End Sub

Private Sub CalcMoney(lngRow As Long, Optional bln��������ۿ� As Boolean)
'���ܣ���������¼���ָ���еĽ��
'������lngRow=ָ����
'˵����1.ExpenseBill���ϵ�������Ӧ���ݵ��к�
'      2.���ֻ�ܶ�Ӧһ��������Ŀ:mobjBill.Details(lngRow).InComes(1)
'      3.������ϸĿδ�����������Ŀ(��һ�μ���),��ʹ��Ĭ���ּ�
'      4.������ϸĿ�Ѿ������������Ŀ(����2��),���ֶ�����(Ҳ����δ��)�˵���,�򰴸õ��ۼ��㡣
    Dim rsTmp As ADODB.Recordset
    Dim strInfo As String, strSQL As String, i As Long
    Dim dblMoney As Double '�û�����ı�۽��
    
    Dim dblAllTime As Double, dbl�Ӱ�Ӽ��� As Double
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dblʣ������ As Double
    Dim strPriceGrade As String, strWherePriceGrade As String
    
    On Error GoTo errH
    If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
        strPriceGrade = mstrҩƷ�۸�ȼ�
    ElseIf mobjBill.Details(lngRow).�շ���� = "4" Then
        strPriceGrade = mstr���ļ۸�ȼ�
    Else
        strPriceGrade = mstr��ͨ�۸�ȼ�
    End If
    
    If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
        Call AdjustCpt(mobjBill.Details(lngRow).�շ�ϸĿID)
    End If
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "       And (b.�۸�ȼ� = [2]" & vbNewLine & _
            "            Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "                And Not Exists(Select 1" & vbNewLine & _
            "                               From �շѼ�Ŀ" & vbNewLine & _
            "                               Where b.�շ�ϸĿId = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
            "                                     And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null"
    End If
    strSQL = _
        " Select B.������ĿID,C.����,C.�վݷ�Ŀ,B.�ּ�,B.ԭ��,B.�Ӱ�Ӽ���,B.�����շ���,B.ȱʡ�۸� " & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C " & _
        " Where B.�շ�ϸĿID = A.ID And C.ID = B.������ĿID " & _
        " And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD')) " & _
        " And A.ID=[1]" & vbNewLine & _
        strWherePriceGrade
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details(lngRow).�շ�ϸĿID, strPriceGrade)
    If rsTmp.EOF Then
        '���û��������Ŀ,�������Ӧ�ĳ������
        Set mobjBill.Details(lngRow).InComes = New BillInComes
        Exit Sub
    End If
    
    '�Ȼ�ȡ����Ա��ǰ����ı�۽��
    With mobjBill.Details(lngRow)
        If InStr(",5,6,7,", .�շ����) > 0 Or (.�շ���� = "4" And .Detail.��������) Then
            '����ҩƷʱ��(�����򲻷���)
            '��Ȼ�м�¼(�������Ŀʱ���ж�)
            dblAllTime = .���� * .����
            If gblnסԺ��λ And InStr(",5,6,7,", .�շ����) > 0 Then
                dblAllTime = dblAllTime * .Detail.סԺ��װ '���ʱ�۰��ۼ��������м���
            End If
            If dblAllTime <> 0 Or Not .Detail.��� Then
                If .Detail.���� <= 0 Then
                    strSQL = "Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual"
                Else
                    strSQL = "Select Zl_Fun_Getprice([1],[2],[3],[4],[5]) As Price From Dual"
                End If
                Set rsPrice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .�շ�ϸĿID, .ִ�в���ID, dblAllTime, 0, .Detail.����)
                If rsPrice.EOF Then
                    '��ȡ�۸�ʧ��
                    If InStr(",5,6,7,", .�շ����) > 0 Then
                        MsgBox "�� " & lngRow & " ��ҩƷ""" & .Detail.���� & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                    Else
                        MsgBox "�� " & lngRow & " ����������""" & .Detail.���� & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                    End If
                Else
                    strPrice = Nvl(rsPrice!Price) & "|||"
                    varPrice = Split(strPrice, "|")
                    dblMoney = Val(varPrice(0))
                    dblʣ������ = Val(varPrice(2))
                    
                    If dblʣ������ <> 0 And .Detail.��� Then
                        '����δ�ֽ����
                        If InStr(",5,6,7,", .�շ����) > 0 Then
                            MsgBox "�� " & lngRow & " ��ʱ��ҩƷ""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                        Else
                            MsgBox "�� " & lngRow & " ��ʱ����������""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                        End If
                        dblMoney = 0
                    End If
                End If
            Else
                dblMoney = 0
            End If
        Else
            If .Detail.��� Then
                If .InComes.Count = 0 Then '��һ�μ�����ȡȱʡֵ
                    dblMoney = Val(Nvl(rsTmp!ȱʡ�۸�))
                Else                        '��ȡ����Ա��ǰ����ı�۽��
                    dblMoney = .InComes(1).��׼����
                    '����û�����ı�۲������۷�Χ����ȡȱʡֵ
                    If CheckScope(Val(Nvl(rsTmp!ԭ��)), Val(Nvl(rsTmp!�ּ�)), dblMoney) <> "" Then
                        dblMoney = Val(Nvl(rsTmp!ȱʡ�۸�))
                    End If
                End If
            End If
        End If
    End With
    
    '�����ԭ�м�¼
    Set mobjBill.Details(lngRow).InComes = New BillInComes
    
    '��д���з��ü�¼
    For i = 1 To rsTmp.RecordCount
        Set mobjBillIncome = New BillInCome
        With mobjBillIncome
            .������ĿID = rsTmp!������ĿID
            .������Ŀ = rsTmp!����
            .�վݷ�Ŀ = Nvl(rsTmp!�վݷ�Ŀ)
            .ԭ�� = Val(Nvl(rsTmp!ԭ��))
            .�ּ� = Val(Nvl(rsTmp!�ּ�))
            
            If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
                If gblnסԺ��λ Then
                    .��׼���� = Format(dblMoney * mobjBill.Details(lngRow).Detail.סԺ��װ, gstrFeePrecisionFmt)
                Else
                    .��׼���� = Format(dblMoney, gstrFeePrecisionFmt)
                End If
            Else
                If mobjBill.Details(lngRow).Detail.��� Then
                    .��׼���� = Format(dblMoney, gstrFeePrecisionFmt)
                Else
                    .��׼���� = Format(Nvl(rsTmp!�ּ�, 0), gstrFeePrecisionFmt)
                End If
            End If
            
            'Ӧ�ս��=���� * ���� * ����
            .Ӧ�ս�� = .��׼���� * IIf(mobjBill.Details(lngRow).���� = 0, 1, mobjBill.Details(lngRow).����) * mobjBill.Details(lngRow).����
            
            '�������������ü���(����������Ŀ)
            If mobjBill.Details(lngRow).���ӱ�־ = 1 And mobjBill.Details(lngRow).�շ���� = "F" Then
                .Ӧ�ս�� = .Ӧ�ս�� * IIf(IsNull(rsTmp!�����շ���), 1, rsTmp!�����շ��� / 100)
            End If
            '�Ӱ�����ʼ���
            dbl�Ӱ�Ӽ��� = 0
            If mobjBill.�Ӱ��־ = 1 And mobjBill.Details(lngRow).Detail.�Ӱ�Ӽ� Then
                dbl�Ӱ�Ӽ��� = IIf(IsNull(rsTmp!�Ӱ�Ӽ���), 0, rsTmp!�Ӱ�Ӽ��� / 100)
                .Ӧ�ս�� = .Ӧ�ս�� + .Ӧ�ս�� * dbl�Ӱ�Ӽ���
            End If
            
            .Ӧ�ս�� = CCur(Format(.Ӧ�ս��, gstrDec))
            dblAllTime = mobjBill.Details(lngRow).���� * mobjBill.Details(lngRow).����
            If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 Then
                If gblnסԺ��λ Then dblAllTime = dblAllTime * mobjBill.Details(lngRow).Detail.סԺ��װ
            End If
            
            If mobjBill.Details(lngRow).Detail.���ηѱ� Or bln��������ۿ� Or .Ӧ�ս�� = 0 Then
                .ʵ�ս�� = .Ӧ�ս��
            Else
                If .Ӧ�ս�� = 0 Then
                    .ʵ�ս�� = 0
                    mobjBill.Details(lngRow).�ѱ� = mobjBill.�ѱ�
                Else
                     'ҩƷ���ɱ��ۼ���,��������
                    .ʵ�ս�� = CCur(Format(ActualMoney(mobjBill.�ѱ�, .������ĿID, .Ӧ�ս��, _
                         mobjBill.Details(lngRow).�շ�ϸĿID, mobjBill.Details(lngRow).ִ�в���ID, dblAllTime, dbl�Ӱ�Ӽ���), gstrDec))
                End If
            End If
            
            '��ȡ��Ŀ������Ϣ,ҽ�����˲Ŵ���,����Ҫ����ҽ��
            If mrsInfo.State = 1 Then
                If Not IsNull(mrsInfo!����) Then
                    strInfo = gclsInsure.GetItemInsure(mobjBill.����ID, mobjBill.Details(lngRow).�շ�ϸĿID, .ʵ�ս��, False, mrsInfo!����, _
                         mobjBill.Details(lngRow).ժҪ & "||" & dblAllTime)
                    If strInfo <> "" Then
                        mobjBill.Details(lngRow).������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                        mobjBill.Details(lngRow).���մ���ID = Val(Split(strInfo, ";")(1))
                        .ͳ���� = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                        mobjBill.Details(lngRow).���ձ��� = CStr(Split(strInfo, ";")(3))
                        
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(lngRow).ժҪ = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(lngRow).Detail.���� = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                End If
            End If
            
            'ʵ�ս�����Key��,�Դ���ֱ�����(��Key�д��ԭʼʵ�ս��,����)
            mobjBill.Details(lngRow).InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��, .ͳ����
        End With
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowDetails(Optional lngRow As Long = 0)
'���ܣ�ˢ����ʾָ���л������е�����
'������lngRow=ָ����,Ϊ0��ʾ��ʾ������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long, curTotal As Currency
    
    Bill.Redraw = False
    If lngRow = 0 Then
        For i = 1 To mobjBill.Details.Count
        
            ShowDetail i
        Next
    Else
        ShowDetail lngRow
    End If
    Bill.Redraw = True
    
    curTotal = GetBillTotal(mobjBill)
        
    If IsNumeric(cmdOK.Tag) Then
        '����ʱ��ʾ���㵱ǰ���ݷ���,�����۱���Ҫ��
        'sta.Panels(3).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
        'sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
        'sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(Val(txtʵ��.Tag) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
        Call SetStatuPatiInfor(Val(cmdOK.Tag), Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), Val(txtʵ��.Tag) - IIf(gbytBilling = 0, curTotal, 0))
    End If
End Sub

Private Sub ShowDetail(lngRow As Long)
'���ܣ�ˢ����ʾָ���е�����
'������lngRow=ָ����
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim dbl���� As Double, cur��� As Currency
    Dim i As Long, j As Long
    
    If lngRow > Bill.Rows - 1 Then Exit Sub
    
    '���������
    For i = 1 To Bill.Cols - 1
        '����ʱ�շ�������
        If Not (i = 1 And Bill.TextMatrix(lngRow, i) <> "") Then Bill.TextMatrix(lngRow, i) = ""
    Next
    
    If mobjBill.Details(lngRow).�շ���� <> "" Then
        Bill.RowData(lngRow) = Asc(mobjBill.Details(lngRow).�շ����)
    End If
    
    'ˢ�µ�����
    For i = 1 To Bill.Cols - 1
        Select Case Bill.TextMatrix(0, i)
            Case "���"
                '������ݻ������Ŀֻ(��)��ʾ����
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.�������
            Case "��Ŀ"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.����
            Case "���"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.���
            Case "��Ʒ��"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.��Ʒ��
            Case "��λ"
                If InStr(",5,6,7,", mobjBill.Details(lngRow).�շ����) > 0 And gblnסԺ��λ Then
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.סԺ��λ
                Else
                    Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.���㵥λ
                End If
            Case "����"
                Bill.TextMatrix(lngRow, i) = IIf(mobjBill.Details(lngRow).���� = 0, 1, mobjBill.Details(lngRow).����)
            Case "����"
                '�����ڵ�һ����ʾʱ��Ĭ������Ϊ1
                Bill.TextMatrix(lngRow, i) = FormatEx(mobjBill.Details(lngRow).����, 5)
            Case "����"
                '�����Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                '��һ�μ���ʱ����Ĭ������Ϊ1�Ļ����ϼ��������
                dbl���� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        dbl���� = dbl���� + mobjBill.Details(lngRow).InComes(j).��׼����
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(dbl����, gstrFeePrecisionFmt)
            Case "Ӧ�ս��"
                'Ӧ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                cur��� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur��� = cur��� + mobjBill.Details(lngRow).InComes(j).Ӧ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur���, gstrDec)
            Case "ʵ�ս��"
                'ʵ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                cur��� = 0
                If mobjBill.Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Details(lngRow).InComes.Count
                        cur��� = cur��� + mobjBill.Details(lngRow).InComes(j).ʵ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur���, gstrDec)
            Case "ִ�п���"
                If mobjBill.Details(lngRow).ִ�в���ID <> 0 Then
                    If mbytInState = 0 Then
                        mrsUnit.Filter = "ID=" & mobjBill.Details(lngRow).ִ�в���ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(lngRow, i) = mrsUnit!���� & "-" & mrsUnit!����
                        Else
                            Bill.TextMatrix(lngRow, i) = GET��������(mobjBill.Details(lngRow).ִ�в���ID, mrsUnit)
                        End If
                    Else
                        '�������ֻ(��)��ʾ����
                        Bill.TextMatrix(lngRow, i) = GET��������(mobjBill.Details(lngRow).ִ�в���ID, mrsUnit)
                    End If
                Else
                    Bill.TextMatrix(lngRow, i) = ""
                End If
            Case "��־"
                If mobjBill.Details(lngRow).�շ���� = "F" And mobjBill.Details(lngRow).���ӱ�־ = 1 Then
                    Bill.TextMatrix(lngRow, i) = "��"
                End If
            Case "����"
                Bill.TextMatrix(lngRow, i) = mobjBill.Details(lngRow).Detail.����
        End Select
    Next
    Bill.Text = Bill.MsfObj.Text
End Sub

Public Sub ShowMoney()
'���ܣ�ˢ����ʾ������Ŀ������
    Dim i As Long, j As Long, k As Long
    Dim blnExist As Boolean, curTotal As Currency, curӦ��Total As Currency
    
    mshMoney.Redraw = False
    
    '�������ܷ�Ŀ
    Set mcolMoneys = New BillInComes
    For i = 1 To mobjBill.Details.Count
        For j = 1 To mobjBill.Details(i).InComes.Count
            '�����Ƿ��Ѿ��������������Ŀ,������ϼ�,��������
            blnExist = False
            For k = 1 To mcolMoneys.Count
                If mcolMoneys(k).������ĿID = mobjBill.Details(i).InComes(j).������ĿID Then
                    blnExist = True: Exit For
                End If
            Next
            
            If blnExist Then
                mcolMoneys(k).ʵ�ս�� = mcolMoneys(k).ʵ�ս�� + mobjBill.Details(i).InComes(j).ʵ�ս��
                mcolMoneys(k).Ӧ�ս�� = mcolMoneys(k).Ӧ�ս�� + mobjBill.Details(i).InComes(j).Ӧ�ս��
            Else
                With mobjBill.Details(i).InComes(j)
                    mcolMoneys.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��
                End With
            End If
        Next
    Next
    
    'ˢ����ʾ
    If mcolMoneys.Count > 0 Then
        mshMoney.Rows = mcolMoneys.Count + 1
    End If
    If mshMoney.Rows < 5 Then mshMoney.Rows = 5

    Call SetMoneyList
    
    For i = 1 To mcolMoneys.Count
        mshMoney.TextMatrix(i, 0) = mcolMoneys(i).������Ŀ
        mshMoney.TextMatrix(i, 1) = Format(mcolMoneys(i).ʵ�ս��, gstrDec)
        curTotal = curTotal + mcolMoneys(i).ʵ�ս��
        curӦ��Total = curӦ��Total + mcolMoneys(i).Ӧ�ս��
    Next
    
    txtӦ��.Text = Format(curӦ��Total, gstrDec)
    txtʵ��.Text = Format(curTotal, gstrDec)
    
    For i = 1 To mshMoney.Rows - 1
        mshMoney.TopRow = i
    Next
    mshMoney.Redraw = True
End Sub

Private Function GetCurӦ��() As Currency
'���ܣ���ȡ���˵�ǰ���ݺϼƽ��(�շѲ����ۼӵ���ʱ��)
    Dim i As Long
    For i = 1 To mcolMoneys.Count
        GetCurӦ�� = GetCurӦ�� + mcolMoneys(i).Ӧ�ս��
    Next
End Function

Private Function GetInputDetail(ByVal lng��Ŀid As Long) As Detail
'���ܣ���ȡ�շ���Ŀ��Ϣ
    Dim objDetail As New Detail
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, lngMediCareNO As Long
        
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!����)
    If lngMediCareNO <> 0 Then
        strSQL = _
        " Select A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ," & _
        "       A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.��������,A.����ժҪ,A.�������,M.Ҫ������," & _
        "       Decode(A.���,'4',D.����ID,C.ҩ��ID) as ҩ��ID," & _
        "       Decode(A.���,'4',D.���÷���,C.ҩ������) as ����," & _
        "       Decode(A.���,'4',1,C.סԺ��װ) as סԺ��װ," & _
        "       Decode(A.���,'4',A.���㵥λ,C.סԺ��λ) as סԺ��λ,D.��������,A.¼������,C.��ҩ��̬,M1.���� as ��������,M1.���㵥λ as ������λ,C.����ϵ��" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,ҩƷ��� C,�������� D,�շ���Ŀ���� E,�շ���Ŀ���� E1,����֧����Ŀ M,������ĿĿ¼ M1" & _
        " Where A.���=B.���� And A.ID=C.ҩƷID(+) And C.ҩ��ID=M1.id(+) And A.ID=D.����ID(+)" & _
        "       And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "       And A.ID=M.�շ�ϸĿID(+) And M.����(+)=[2]" & vbNewLine & _
        "       And A.ID=[1]"
    Else
        strSQL = _
        " Select A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ," & _
        "       A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.��������,A.����ժҪ,A.�������,0 as Ҫ������," & _
        "       Decode(A.���,'4',D.����ID,C.ҩ��ID) as ҩ��ID," & _
        "       Decode(A.���,'4',D.���÷���,C.ҩ������) as ����," & _
        "       Decode(A.���,'4',1,C.סԺ��װ) as סԺ��װ," & _
        "       Decode(A.���,'4',A.���㵥λ,C.סԺ��λ) as סԺ��λ,D.��������,A.¼������,C.��ҩ��̬,M1.���� as ��������,M1.���㵥λ as ������λ,C.����ϵ��" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,ҩƷ��� C,�������� D,�շ���Ŀ���� E,�շ���Ŀ���� E1,������ĿĿ¼ M1" & _
        " Where A.���=B.���� And A.ID=C.ҩƷID(+) And C.ҩ��ID=M1.id(+) And A.ID=D.����ID(+)" & _
        "       And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "       And A.ID=[1]"
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid, lngMediCareNO)
    With objDetail
        .ID = rsTmp!ID
        .ҩ��ID = Nvl(rsTmp!ҩ��ID, 0) '�����ж������ظ�
        .��� = rsTmp!���
        .������� = rsTmp!�������
        .���� = rsTmp!����
        .���� = rsTmp!����
        .��� = Nvl(rsTmp!���)
        .���㵥λ = Nvl(rsTmp!���㵥λ)
        .סԺ��λ = Nvl(rsTmp!סԺ��λ)
        .סԺ��װ = Nvl(rsTmp!סԺ��װ, 1)
        .���� = Nvl(rsTmp!����, 0) = 1 '�Ƿ�ҩ������
        .��� = Nvl(rsTmp!�Ƿ���, 0) = 1 '��ҩƷ�����Ƿ�ʱ��
        .���� = Nvl(rsTmp!��������)
        .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
        .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
        .ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
        .������� = Nvl(rsTmp!�������, 0)
        .����ժҪ = Nvl(rsTmp!����ժҪ, 0) = 1
        .�������� = Nvl(rsTmp!��������, 0) = 1
        .Ҫ������ = Nvl(rsTmp!Ҫ������, 0) = 1
        .¼������ = Val("" & rsTmp!¼������)
        .��ҩ��̬ = Val(Nvl(rsTmp!��ҩ��̬))
        .��Ʒ�� = Nvl(rsTmp!��Ʒ��)
        .�������� = Nvl(rsTmp!��������)
        .������λ = Nvl(rsTmp!������λ)
        .����ϵ�� = Val(Nvl(rsTmp!����ϵ��))
        .���� = mlng����
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetDetail(Detail As Detail, lngRow As Long, lngDoUnit As Long, Optional bytParent As Byte = 0)
'���ܣ�����ָ�����շ�ϸĿ�����趨����ָ�㶨�е��շ�ϸĿ(�����Ļ��޸�)
'˵����
'      1.���������������շ�ϸĿ�У�����
'      2.��bytParent<>0ʱ,��Ϊ���ô�����Ŀ,������Ŀһ����������,������Ŀһ������

    Dim tmpIncomes As New BillInComes
    Dim intPay As Integer, i As Long, dblTime As Double
    
    'ȡ������ҩ�ĸ���
    intPay = GetPay(lngRow)
    If Detail.��� <> "7" Then intPay = 1
    
    If mobjBill.Details.Count < lngRow Then
        '������ж�Ӧ�ĳ��������δ��ʼ,�����
        With Detail
            '���=�к�,����=0
            '����=1,������Ŀ�Ĵ������������ȷ��
            'ִ�в���ID:����ϸĿִ�п��ұ�־ȡ
            '���ӱ�־:�Ե�һ��Ϊ��,����Ϊ������Ȩ
            '���뼯=��
            If bytParent <> 0 Then
                '���ø���RowData
                Bill.RowData(lngRow) = Asc(Detail.���)
                '��ʼ����
                If Detail.���д��� = 0 Then '�ǹ��д���
                    dblTime = Detail.��������
                ElseIf Detail.���д��� = 1 Then '�̶��Ĺ��д���
                    dblTime = IIf(Detail.�������� = 0, 1, Detail.��������)
                ElseIf Detail.���д��� = 2 Then '�������Ĺ��д���
                    dblTime = Detail.�������� * mobjBill.Details(bytParent).����
                End If
            Else
                
                If InStr(",5,6,7,", Detail.���) > 0 Then
                    dblTime = 0
                                     
                Else
                    dblTime = 1
                End If
            End If
            mobjBill.Details.Add Detail, .ID, CByte(lngRow), CInt(bytParent), 0, 0, 0, 0, "", "", "", _
            0, 0, mobjBill.�ѱ�, 0, .���, .���㵥λ, "", intPay, dblTime, 0, lngDoUnit, tmpIncomes
        End With
    Else '��������Ѿ�����,���޸�
        
        If InStr(",5,6,7,", Detail.���) > 0 Then
            dblTime = 0
        Else
            dblTime = 1
        End If
        
        With mobjBill.Details(lngRow)
            Set .Detail = Detail
            Set .InComes = tmpIncomes
            .�ѱ� = mobjBill.�ѱ�
            .���� = intPay
            .���ӱ�־ = 0
            .���㵥λ = Detail.���㵥λ
            .�շ���� = Detail.���
            .�շ�ϸĿID = Detail.ID
            .���� = dblTime
            .��� = lngRow
            .�������� = 0
            .ִ�в���ID = lngDoUnit
        End With
    End If
End Sub

Private Function ShouldDO(lngRow As Long) As Boolean
'���ܣ��жϸ����Ƿ�Ӧ��ȡ������Ŀ
'˵�����������շ���Ŀ�д�����Ŀ����δȡ��ȡ��
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select count(����ID) as NUM From �շѴ�����Ŀ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details(lngRow).�շ�ϸĿID)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!Num) Then
            ShouldDO = False
        ElseIf rsTmp!Num = 0 Then
            ShouldDO = False
        Else
            blnExist = False
            For i = lngRow + 1 To mobjBill.Details.Count
                If mobjBill.Details(i).�������� = lngRow Then
                    blnExist = True: Exit For
                End If
            Next
            If Not blnExist Then
                ShouldDO = True
            Else
                ShouldDO = False
            End If
        End If
    Else
        ShouldDO = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetSubDetails(ByVal lng��Ŀid As Long) As Details
'���ܣ�����һ���շ�ϸĿ�Ĵ�����Ŀ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objDetail As New Detail, lngMediCareNO As Long
    Dim dblStock As Double
    
    Set GetSubDetails = New Details
    
    If mrsInfo.State = 1 Then lngMediCareNO = Val(Nvl(mrsInfo!����))
    If lngMediCareNO > 0 Then
        strSQL = _
        " Select A.ID,Decode(A.���,'4',E.����ID,D.ҩ��ID) as ҩ��ID,A.���,B.���� as �������," & _
        "       A.��������,A.����,Nvl(F.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ,A.���ηѱ�,G.Ҫ������," & _
        "       Decode(A.���,'4',E.���÷���,D.ҩ������) as ����,A.�Ƿ���," & _
        "       Decode(A.���,'4',1,D.סԺ��װ) as סԺ��װ,A.�������," & _
        "       Decode(A.���,'4',A.���㵥λ,D.סԺ��λ) as סԺ��λ," & _
        "       A.�Ӱ�Ӽ�,A.ִ�п���,C.���д���,C.��������,E.��������,D.��ҩ��̬,M1.���� as ��������,M1.���㵥λ as ������λ,D.����ϵ��" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�շѴ�����Ŀ C,ҩƷ��� D,�������� E,�շ���Ŀ���� F,�շ���Ŀ���� E1,����֧����Ŀ G,������ĿĿ¼ M1" & _
        " Where B.����=A.��� And C.����ID=A.ID And A.ID=D.ҩƷID(+) And D.ҩ��ID=M1.id(+) And A.ID=E.����ID(+)" & _
        "       And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        "       And A.ID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "       And C.����ID=[1] And A.ID=G.�շ�ϸĿID(+) And G.����(+)=[2] " & _
        " Order by ����"
    Else
        strSQL = _
        "Select A.ID,Decode(A.���,'4',E.����ID,D.ҩ��ID) as ҩ��ID,A.���,B.���� as �������," & _
        "       A.��������,A.����,Nvl(F.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.���㵥λ,A.���ηѱ�,0 as Ҫ������," & _
        "       Decode(A.���,'4',E.���÷���,D.ҩ������) as ����,A.�Ƿ���," & _
        "       Decode(A.���,'4',1,D.סԺ��װ) as סԺ��װ,A.�������," & _
        "       Decode(A.���,'4',A.���㵥λ,D.סԺ��λ) as סԺ��λ," & _
        "       A.�Ӱ�Ӽ�,A.ִ�п���,C.���д���,C.��������,E.��������,D.��ҩ��̬,M1.���� as ��������,M1.���㵥λ as ������λ,D.����ϵ��" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�շѴ�����Ŀ C,ҩƷ��� D,�������� E,�շ���Ŀ���� F,�շ���Ŀ���� E1,������ĿĿ¼ M1" & _
        " Where B.����=A.��� And C.����ID=A.ID And A.ID=D.ҩƷID(+) And D.ҩ��ID=M1.id(+)  And A.ID=E.����ID(+)" & _
        "   And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        "   And A.ID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "   And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        "   And C.����ID=[1] " & _
        " Order by ����"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid, lngMediCareNO)
    For i = 1 To rsTmp.RecordCount
        Set objDetail = New Detail
        With objDetail
            .ID = rsTmp!ID
            .ҩ��ID = Nvl(rsTmp!ҩ��ID, 0)
            .���� = rsTmp!����
            .��� = Nvl(rsTmp!�Ƿ���, 0) = 1
            .��� = Nvl(rsTmp!���)
            .סԺ��װ = Nvl(rsTmp!סԺ��װ, 1)
            .סԺ��λ = Nvl(rsTmp!סԺ��λ)
            .���㵥λ = Nvl(rsTmp!���㵥λ)
            .���� = Nvl(rsTmp!����, 0) = 1
            .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
            .��� = rsTmp!���
            .������� = rsTmp!�������
            .���� = rsTmp!����
            .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
            .ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
            .������� = Nvl(rsTmp!�������, 0)
            .���д��� = Nvl(rsTmp!���д���, 0)
            .�������� = Nvl(rsTmp!��������, 1)
            .���� = Nvl(rsTmp!��������)
            .�������� = Nvl(rsTmp!��������, 0) = 1
            .Ҫ������ = Nvl(rsTmp!Ҫ������, 0) = 1
            .��ҩ��̬ = Val(Nvl(rsTmp!��ҩ��̬))
            .��Ʒ�� = Nvl(rsTmp!��Ʒ��)
            .�������� = Nvl(rsTmp!��������)
            .������λ = Nvl(rsTmp!������λ)
            .����ϵ�� = Val(Nvl(rsTmp!����ϵ��))
            GetSubDetails.Add .ID, .ҩ��ID, .���, .�������, .����, .����, .����, .����, .���, .���㵥λ, .˵��, .���ηѱ�, _
                .סԺ��װ, .סԺ��λ, .����, .���, .�Ӱ�Ӽ�, .ִ�п���, .�������, .����, .����ժҪ, .���д���, .��������, .��������, , , , , , .Ҫ������, , .��ҩ��̬, .��Ʒ��, .��������, .������λ, .����ϵ��
        End With
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub DeleteDetail(lngRow As Long)
'���ܣ�ɾ��ָ���շ���Ŀ��
'˵������ʱ����������е�ɾ��,��Ҫ�����������д�����ϵ����Ӧ�ĵ���
    Dim i As Long
    For i = lngRow + 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�������� <> 0 And mobjBill.Details(i).�������� > lngRow Then
            mobjBill.Details(i).�������� = mobjBill.Details(i).�������� - 1
        End If
        mobjBill.Details(i).��� = mobjBill.Details(i).��� - 1 '������кŶ�Ӧ
    Next
    mobjBill.Details.Remove lngRow
    If lngRow = 1 And mobjBill.Details.Count = 0 And Bill.Rows = 2 Then
        For i = 1 To Bill.Cols - 1
            Bill.TextMatrix(lngRow, i) = ""
            Bill.RowData(lngRow) = 0
        Next
    Else
        Bill.RemoveMSFItem lngRow
    End If
End Sub

Private Sub NewBill(Optional blnPati As Boolean = True)
'���ܣ���ʼ��һ���µĵ���(�������)
'������blnPati=�Ƿ��ʼ��������Ϣ
    Dim blnKeepDate As Boolean
    Dim Curdate As Date     '��������ǰʱ��
    
    mcurModiMoney = 0
    mlngPreRow = 0
    
    If mrsInfo.State = 0 Then txtPatient.ForeColor = Me.ForeColor
        
    If blnPati Then
        sta.Panels(3).Text = "": lblStatuPati.Caption = "": picStatuPancl.Visible = False
        cmdOK.Tag = "": cmdCancel.Tag = "": txtʵ��.Tag = ""
        
        Set mrsMedAudit = Nothing
        Set mrsInfo = New ADODB.Recordset
        txtPatient.Text = "": txtOld.Text = ""
        txt����.Text = "": txtסԺ��.Text = "": txt���˱�ע.Text = ""
        txt������.Text = "": txt������.Text = ""
        cboSex.ListIndex = -1: cbo�ѱ�.ListIndex = -1: cboҽ�Ƹ���.ListIndex = -1
    End If
    
    mstrWarn = ""
    cboNO.Text = ""
    Set mobjBill = New ExpenseBill
    Bill.ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus)
    Curdate = zlDatabase.Currentdate
    chk�Ӱ�.Value = IIf(OverTime(Curdate), 1, 0)
    
    If Not blnPati And mrsInfo.State = 1 Then
        If mrsInfo!��Ժ���� < Curdate Then blnKeepDate = True
    End If
    If Not blnKeepDate Then
        txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
        'ȡ��ǰʱ��:33744
        If mbln���� And mstr���ת��ʱ�� <> "" Then
            txtDate.Text = Format(CDate(mstr���ת��ʱ��) - 1 / 24 / 60, "yyyy-mm-dd HH:MM:SS")
            txtDate.ForeColor = vbBlue
        End If
    End If
    


    Call LoadPatientBaby(cboBaby, 0, 0)
    Call cbo��������_Click
    
    mblnSavePrice = False
    cmdPrice.Visible = False
    cmdOK.Visible = mbytInState <> 1
    Call SetButtonPlace
    
    With mobjBill
        .�����־ = 2
        .������ = UserInfo.����
        .������ = zlStr.NeedName(cbo������.Text)
        .����Ա��� = UserInfo.���
        .����Ա���� = UserInfo.����
        .����ʱ�� = CDate(txtDate.Text)
        .�Ӱ��־ = chk�Ӱ�.Value
        .Ӥ���� = 0
        
        If cbo��������.ListIndex = -1 Then
            .��������ID = 0
        Else
            .��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        End If
        If cboDrawDept.ListIndex = -1 Then
            .��ҩ����ID = 0
        Else
            .��ҩ����ID = cboDrawDept.ItemData(cboDrawDept.ListIndex)
        End If
    End With
End Sub

Private Sub ClearMoney()
'���ܣ����������ʾ��
    Dim i As Long, j As Long
    mshMoney.Redraw = False
    For i = 1 To mshMoney.Rows - 1
        For j = 0 To mshMoney.Cols - 1
            mshMoney.TextMatrix(i, j) = ""
        Next
    Next
    mshMoney.Rows = 5
    mshMoney.Redraw = True
End Sub

Private Function SaveBill() As Boolean
'����:���浱ǰ����ļ��ʵ���(����סԺ���ʡ����ۡ�������ߵ��޸�)
'���:mobjBill=���ݶ���
'����:�����Ƿ�ɹ�
    Dim i As Long, j As Long, arrSQL As Variant, arrSMSQL As Variant
    Dim int��� As Integer, int�к� As Integer, strNO As String, strTmp As String, str���ܺ� As String
    Dim intParent As Integer, intParentNO As Integer
    Dim str��Ϣ As String, intInsure As Integer
    Dim dbl���� As Double, dbl���� As Double
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim strSQL As String, strStuffDept As String '��¼���Ϸ��ϲ���
    Dim strAddDate As String '���ʷ���,�Զ���ҩ,���ϵ�ʱ��
    Dim blnTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim str��ҩ��̬ As String
     
    mobjBill.NO = zlDatabase.GetNextNo(14)
    strAddDate = "To_Date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    gstrModiNO = mobjBill.NO
    arrSQL = Array()
    arrSMSQL = Array()
        
    '�Ƿ�ҽ������,ȡҽ����Ϣ
    If mstrInNO <> "" Then
        Call BillisAdviceMoney(mstrInNO, 2, lngҽ��ID, lng���ͺ�)
    End If
    If mlng����ҽ�� <> 0 And lngҽ��ID = 0 Then lngҽ��ID = mlng����ҽ��
        
    '���˺�:���»�ȡ��ҩ����
    Call zlReSetDrawDrugDept

    For Each mobjBillDetail In mobjBill.Details
        If mobjBillDetail.���� <> 0 Then
            intParent = 0: intParentNO = int���
            For Each mobjBillIncome In mobjBillDetail.InComes
                int��� = int��� + 1 '��ǰ��¼���
                '��������
                With mobjBill
                    gstrSQL = "zl_סԺ���ʼ�¼_INSERT('" & .NO & "'," & int��� & "," & .����ID & "," & IIf(.��ҳID = 0, "NULL", .��ҳID) & "," & _
                        IIf(Val(.��ʶ��) = 0, "NULL", .��ʶ��) & "," & "'" & .���� & "','" & .�Ա� & "','" & .���� & "','" & .���� & "','" & .�ѱ� & "'," & _
                             IIf(.����ID = 0, .��������ID, .����ID) & "," & IIf(.����ID = 0, .��������ID, .����ID) & "," & .�Ӱ��־ & "," & .Ӥ���� & "," & .��������ID & ",'" & .������ & "',"
                End With
                
                '�շ�ϸĿ����
                With mobjBillDetail
                    '�����������
                    If .��� <> int�к� Then
                        int�к� = .���
                        
                        '���´����������
                        If mobjBill.Details(.���).�������� = 0 Then    'ֻ�д��ڸ���ʱ,�Ż���´�����
                            For i = .��� + 1 To mobjBill.Details.Count
                                If mobjBill.Details(i).�������� = .��� Then
                                    mobjBill.Details(i).�������� = int��� '������Ŀ�ж��������Ŀ(������)ʱ,ȡ��һ�����
                                End If
                            Next
                        End If
                    End If
                    gstrSQL = gstrSQL & .�������� & "," & .�շ�ϸĿID & ",'" & .�շ���� & "','" & .���㵥λ & "',"
                    
                    gstrSQL = gstrSQL & IIf(.������Ŀ��, 1, 0) & "," & IIf(.���մ���ID = 0, "NULL", .���մ���ID) & ",'" & .���ձ��� & "',"
                    
                    dbl���� = .����
                    If InStr(",5,6,7,", .�շ����) > 0 And gblnסԺ��λ Then
                        dbl���� = Format(.���� * .Detail.סԺ��װ, gstrFeePrecisionFmt)
                    End If
                    gstrSQL = gstrSQL & IIf(.���� = 0, 1, .����) & "," & dbl���� & "," & .���ӱ�־ & "," & IIf(.ִ�в���ID = 0, "NULL", .ִ�в���ID) & ","
                    
                    '�ռ����Ϸ��ϲ���,�Ա��Զ�����
                    If Not (gbytBilling = 1 Or mblnSavePrice) And gint���ķ��Ͽ��� <> 0 Then
                        'gint���ķ��Ͽ���:0-���Զ����ϣ�1-�Զ����ϣ�2-�����ҿ���ʱ�Զ�����
                        If .ִ�в���ID <> 0 And .�շ���� = "4" And .Detail.�������� _
                            And ((gint���ķ��Ͽ��� = 2 And .ִ�в���ID = mobjBill.��������ID) Or gint���ķ��Ͽ��� = 1) Then
                            If InStr("," & strStuffDept, "," & .ִ�в���ID & ",") = 0 Then
                                strStuffDept = strStuffDept & "," & .ִ�в���ID
                            End If
                        End If
                    End If
                End With
                
                '������Ŀ����
                With mobjBillIncome
                    intParent = intParent + 1
                    dbl���� = .��׼����
                    If InStr(",5,6,7,", mobjBillDetail.�շ����) > 0 And gblnסԺ��λ Then
                        dbl���� = Format(.��׼���� / mobjBillDetail.Detail.סԺ��װ, gstrFeePrecisionFmt)
                    End If
                    gstrSQL = gstrSQL & IIf(intParent = 1, "Null", intParentNO + 1) & "," & .������ĿID & "," & _
                        "'" & .�վݷ�Ŀ & "'," & dbl���� & "," & .Ӧ�ս�� & "," & .ʵ�ս�� & "," & _
                        IIf(.ͳ���� = 0, "NULL", .ͳ����) & ","
                End With
                '���˺� ����:27383 ����:2010-02-01 17:02:08
                If cboִ������.ListIndex < 0 Or cboִ������.Enabled = False Then
                    strTmp = "NULL,NULL"
                ElseIf cboִ������.ItemData(cboִ������.ListIndex) = 0 Then
                    strTmp = "NULL,NULL"
                Else
                    strTmp = "1," & cboִ������.ItemData(cboִ������.ListIndex)
                End If
               
                If mobjBillDetail.�շ���� = "7" Then
                    str��ҩ��̬ = "'" & mobjBillDetail.Detail.��ҩ��̬ & "'"
                Else
                    str��ҩ��̬ = "NULL"
                End If
                
                '��ҩ��̬_In       סԺ���ü�¼.����%Type := Null
                '��������
                '�����:117445,����,2017/12/4,�����ҩƷ�����ĵ�����Ϊ0ʱû�������ݴ���
                gstrSQL = gstrSQL & _
                    "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & strAddDate & "," & _
                    "'" & mstrInNO & "'," & IIf(gbytBilling = 1 Or mblnSavePrice, 1, 0) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                    "0," & IIf(mobjBillDetail.�շ���� = "4", mlng�������ID, mlngҩƷ���ID) & "," & _
                    "NULL,'" & mobjBillDetail.ժҪ & "'," & chk����.Value & "," & ZVal(lngҽ��ID) & _
                    ",Null,Null,'|" & mobjBill.�巨 & "', " & strTmp & ",NULL,'" & mobjBillDetail.Detail.���� & "',0," & mobjBill.��ҩ����ID & "," & _
                    str��ҩ��̬ & ",-1,0," & IIf(mobjBillDetail.Detail.���� = -1 Or mobjBillDetail.Detail.���� = 0, "Null", mobjBillDetail.Detail.����) & ")"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mobjBillDetail.�շ�ϸĿID & ";" & gstrSQL
            Next
        End If
    Next
    
    '�޸�ǰ�˳�ԭ����
    '---------------------------------------------------------------
    If mstrInNO <> "" Then
        '���ж��Ƿ�ҽ�����˼ǵ���,�����Ϸ��Լ��(�����޸�ʱ������һ������ж�)
        If gbytBilling = 0 Then '�޸Ļ��۵�ʱ����
            intInsure = BillExistInsure(mstrInNO)
            If intInsure > 0 Then
                'ȥ����ҽ������ƥ����
            End If
        End If
        gstrSQL = "zl_סԺ���ʼ�¼_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        If gstrSQL <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
        End If
    End If

    '������޸�ҽ���ĸ���,���µ�NO���ڸ�����
    '---------------------------------------------------------------
    If lngҽ��ID <> 0 And lng���ͺ� <> 0 Then
        gstrSQL = "ZL_����ҽ������_Insert(" & lngҽ��ID & "," & lng���ͺ� & ",2,'" & mobjBill.NO & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
    End If
    

    If UBound(arrSQL) >= 0 Then
        '��SQL���а��շ�ϸĿID����
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                    strTmp = CStr(arrSQL(j))
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next
        
        'ִ��SQL���
        On Error GoTo errH
        gcnOracle.BeginTrans
            blnTrans = True
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
            Next
            
            'ִ���Զ�����
            If strStuffDept <> "" Then
                strStuffDept = Mid(strStuffDept, 2)
                For i = 0 To UBound(Split(strStuffDept, ","))
                    strSQL = "zl_�����շ���¼_��������(" & Split(strStuffDept, ",")(i) & ",25,'" & mobjBill.NO & "','" & _
                        UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1," & strAddDate & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                Next
            End If
            
            '׼���Զ���ҩ(����ͨ����),�����������в��ܶ�������
            If mblnSendMateria Then
                Set rsTmp = Get����ҩ�嵥(mobjBill.NO, Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"), False)
                If rsTmp.RecordCount > 0 Then
                    str���ܺ� = zlDatabase.GetNextNo(20)
                    ReDim arrSMSQL(rsTmp.RecordCount - 1)
                    For i = 0 To rsTmp.RecordCount - 1
                        arrSMSQL(i) = "ZL_ҩƷ�շ���¼_���ŷ�ҩ(" & rsTmp!�ⷿID & "," & rsTmp!ID & ",'" & UserInfo.���� & "'," & strAddDate & ",Null,Null,Null," & str���ܺ� & ")"
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Close
            End If
            'ִ���Զ���ҩ
            For i = 0 To UBound(arrSMSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSMSQL(i)), Me.Caption)
            Next

            
            'ҽ���ӿ�
            '1.ҽ�����������ϴ�
            If mstrInNO <> "" And gbytBilling = 0 And intInsure <> 0 Then
                If MCPAR.���������ϴ� And Not MCPAR.������ɺ��ϴ� Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Function
                    End If
                End If
            End If
            
            '2.����ʵʱ�ϴ�
            If gbytBilling = 0 And Not mblnSavePrice And Not IsNull(mrsInfo!����) Then
                'ҽ�����������ϸ
                If MCPAR.�����ϴ� And Not MCPAR.������ɺ��ϴ� Then
                    str��Ϣ = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str��Ϣ, , mrsInfo!����) Then
                        gcnOracle.RollbackTrans
                        If str��Ϣ <> "" Then MsgBox str��Ϣ, vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        gcnOracle.CommitTrans
        blnTrans = False
        
        'ҽ���ӿ�
        '1.ҽ�����������ϴ�
        If mstrInNO <> "" And gbytBilling = 0 And intInsure <> 0 Then
            If MCPAR.���������ϴ� And MCPAR.������ɺ��ϴ� Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "����""" & mstrInNO & """������������ҽ������ʧ��,�õ��������ʣ�", vbInformation, gstrSysName
                End If
            End If
        End If
        
        '2.����ʵʱ�ϴ�
        If gbytBilling = 0 And Not mblnSavePrice And Not IsNull(mrsInfo!����) Then
            'ҽ�����������ϸ
            If MCPAR.�����ϴ� And MCPAR.������ɺ��ϴ� Then
                str��Ϣ = ""
                If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str��Ϣ, , mrsInfo!����) Then
                    If str��Ϣ <> "" Then
                        MsgBox str��Ϣ, vbInformation, gstrSysName
                    Else
                        MsgBox "����""" & mobjBill.NO & """��������ҽ������ʧ��,�õ����ѱ��棡", vbInformation, gstrSysName
                    End If
                End If
            End If
        End If
        
        '���뵥����ʷ��¼(�������͵���)
        For i = 0 To cboNO.ListCount - 1
            strNO = strNO & "," & cboNO.List(i)
        Next
        strNO = mobjBill.NO & strNO
        cboNO.Clear
        For i = 0 To UBound(Split(strNO, ","))
            cboNO.AddItem Split(strNO, ",")(i)
            If i = 9 Then Exit For 'ֻ��ʾ10��
        Next
        
        'ҽ���ӿ�
        If str��Ϣ <> "" Then MsgBox str��Ϣ, vbInformation, gstrSysName
    End If
    SaveBill = True
    Exit Function
errH:
    If Err.Description Like "*��ǰ���㵥�۲�һ��*" Then
       If blnTrans Then gcnOracle.RollbackTrans
       
       If MsgBox("ĳЩ����ҩƷ�۸��ѷ����仯��Ҫ�Զ�����۸���", vbYesNo + vbQuestion + vbDefaultButton1, App.ProductName) = vbYes Then
           Call CalcMoneys
           Call ShowDetails
           Call ShowMoney
           Exit Function
       End If
    Else
        If blnTrans Then gcnOracle.RollbackTrans
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Function
 


Private Function ReadBill(ByVal strNO As String, Optional blnDelete As Boolean, Optional ByVal bytType As Byte = 2) As Integer
'���ܣ����ݵ��ݺŶ�ȡһ�ŵ��ݲ�����������
'������strNO=���ݺ�
'      blnDelete=True:���ʵ���ʱ����,False:���ĵ���ʱ����
'      bytType=2:��ͨ���ʵ�,3-�Զ����ʵ�
    Dim rsTmp As ADODB.Recordset
    Dim rsPatiMoney As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String, strSQL2 As String
    Dim strAdvice As String, strFeeKind As String, strҽ�Ƹ��� As String, strUserUnitIDs As String
    Dim rsִ������ As ADODB.Recordset, intִ������ As Integer
    Dim i As Long, intSign As Integer, intInsure As Integer
    Dim curTotal As Currency, curӦ��Total As Currency
    Dim blnDo As Boolean
    
    On Error GoTo errH
    
    mblnPrint = False
    
    '������֮ǰ�Ѽ��,������һ������Ȩ��
    If blnDelete Then
        '55380
        Dim blnYP As Boolean, blnZL As Boolean, blnWC As Boolean
        blnYP = zlStr.IsHavePrivs(mstrPrivsOpt, "ҩƷ����")
        blnZL = zlStr.IsHavePrivs(mstrPrivsOpt, "��������")
        blnWC = zlStr.IsHavePrivs(mstrPrivsOpt, "��������")
        If blnYP And blnWC And blnZL Then
            '����,������
        ElseIf blnYP And blnWC And Not blnZL Then
            strFeeKind = " And �շ����   In('4','5','6','7')"
        ElseIf blnYP And Not blnWC And blnZL Then
            strFeeKind = " And �շ����   <>'4'"
        ElseIf blnYP And Not blnWC And Not blnZL Then
            strFeeKind = " And �շ���� In('5','6','7')"
        ElseIf Not blnYP And blnWC And blnZL Then
            strFeeKind = " And �շ���� Not In('5','6','7')"
        ElseIf Not blnYP And Not blnWC And blnZL Then
            strFeeKind = " And �շ���� Not In('4','5','6','7')"
        ElseIf Not blnYP And blnWC And Not blnZL Then
            strFeeKind = " And �շ���� ='4'"
        End If
    End If
    
    Call ClearRows: Call Bill.ClearBill: Call SetColNum: Call ClearMoney
    
    '��ȡ��������
    strNO = GetFullNO(strNO, IIf(bytType = 2, 14, 17))
   
    strSQL = _
        " Select A.����ID,Nvl(A.��ҳID,0) as ��ҳID,A.����,A.�Ա�,A.����,A.�ѱ�,A.����,A.��ʶ��," & _
        " A.���˲���ID,A.��������ID,Nvl(A.�Ӱ��־,0) as �Ӱ��־,Nvl(A.Ӥ����,0) as Ӥ����," & _
        " A.������,A.������,A.����Ա����,A.����ʱ��,A.����ID,B.������,B.������," & _
        " Nvl(A.�Ƿ���,0) as �Ƿ���,B1.��ע as ���˱�ע" & _
        " From  " & _
                 IIf(mblnNOMoved And gbytBilling = 0, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & "," & _
        "        ������Ϣ B,��Ա�� C,������ҳ B1 " & _
        " Where A.NO=[1] And A.��¼����=[2] And A.�����־=2 And Nvl(A.�ಡ�˵�,0)=0 " & _
        "       And Nvl(A.����Ա����,A.������)=C.���� and A.����ID=B1.����ID(+) and A.��ҳID=B1.��ҳID(+)" & _
        "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & vbNewLine & _
        "       And A.����ID=B.����ID And Rownum=1 And A.��¼״̬" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
        IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[3]", "") & _
        IIf(mbytInState = 0 And gbytBilling = 0, " And A.����Ա���� is Not Null", "") & _
        IIf(mbytInState = 0 And gbytBilling = 1, " And A.����Ա���� is Null And A.������ is Not NULL", "") & _
        IIf(mbytInState = 0 And gbytBilling = 2, " And A.����Ա���� is Null And A.������ is Not NULL", "")
    
    
    
    
    
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, bytType, CDate(mstrTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, bytType)
    End If
    
     
    If rsTmp.EOF Then
        MsgBox "û���ҵ��õ��ݣ�����õ����Ƿ�����סԺ���ʵ�.", vbInformation, gstrSysName
        Exit Function
    Else
        If blnDelete Then
            If InStr(mstrPrivsOpt, ";ȫԺ����;") = 0 Then
                strUserUnitIDs = GetUserUnits(True)
                If InStr("," & strUserUnitIDs & ",", "," & rsTmp!��������ID & ",") = 0 Then
                    MsgBox "��û��Ȩ�޶��������ҵĵ������ʣ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            If mbytUseType = 0 Or mbytUseType = 1 Then
                If InStr(mstrPrivs, ";���в���;") = 0 And mlngUnitID > 0 Then
                    If InStr(1, "," & mstrUnitIDs & ",", "," & IIf(IsNull(rsTmp!���˲���ID), 0, rsTmp!���˲���ID) & ",") = 0 Then
                        MsgBox "��û��Ȩ�޶�ȡ���������ĵ��ݣ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            ElseIf mbytUseType = 2 Then
                If InStr(mstrPrivs, ";���п���;") = 0 And mlngDeptID > 0 Then
                    If IIf(IsNull(rsTmp!��������ID), 0, rsTmp!��������ID) <> mlngDeptID Then
                        MsgBox "��û��Ȩ�޶�ȡ�������ҿ����ĵ��ݣ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    '���˺� ����:27383 ����:2010-02-01 16:58:14
    gstrSQL = "" & _
        "   Select Max(����) as ִ������,Count(*) as ��¼�� " & _
        "     From ҩƷ�շ���¼ " & _
        "     Where NO = [1] And ���� =9 "
    Set rsִ������ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNO, bytType)
    
    If Val(Nvl((rsִ������!��¼��))) = 0 Then
        If mbytInState = 1 Then cboִ������.Visible = False: lblִ������.Visible = False
    Else
        intִ������ = Val(Mid(Nvl(rsִ������!ִ������) & "00", 2, 1))
    End If
    
    zlControl.CboLocate cboִ������, intִ������, True
    
    If blnDelete Then
        mlngBill����ID = rsTmp!����ID
        mlngBill��ҳID = rsTmp!��ҳID
    End If

    '���ݺ�
    cboNO.Text = strNO

    '����
    txtPatient.Text = Nvl(rsTmp!����)
    
    '�Ա�
    cboSex.ListIndex = cbo.FindIndex(cboSex, Nvl(rsTmp!�Ա�), True)
    If cboSex.ListIndex = -1 And Not IsNull(rsTmp!�Ա�) Then
        cboSex.AddItem rsTmp!�Ա�, 0
        cboSex.ListIndex = 0
    End If
    '����
    txtOld.Text = Nvl(rsTmp!����)
    
    '����
    txt����.Text = "" & rsTmp!����
    txtסԺ��.Text = Nvl(rsTmp!��ʶ��)
    txt���˱�ע.Text = Nvl(rsTmp!���˱�ע)
    txt������.Text = Nvl(rsTmp!������)
    txt������.Text = Format(Nvl(rsTmp!������), "0.00")
    
    '�ѱ�
    cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, Nvl(rsTmp!�ѱ�), True)
    If cbo�ѱ�.ListIndex = -1 And Not IsNull(rsTmp!�ѱ�) Then
        cbo�ѱ�.AddItem rsTmp!�ѱ�, 0
        cbo�ѱ�.ListIndex = 0
    End If
    
    'ҽ�Ƹ��ʽ
    strҽ�Ƹ��� = Get����ҽ�Ƹ��ʽ(rsTmp!����ID, rsTmp!��ҳID)
    cboҽ�Ƹ���.ListIndex = cbo.FindIndex(cboҽ�Ƹ���, strҽ�Ƹ���, True)
    If cboҽ�Ƹ���.ListIndex = -1 And strҽ�Ƹ��� <> "" Then
        cboҽ�Ƹ���.AddItem strҽ�Ƹ���, 0
        cboҽ�Ƹ���.ListIndex = 0
    End If
        
    '�Ƿ���
    If rsTmp!�Ƿ��� = 1 Then chk����.Value = 1: chk����.Visible = True
    
    txtDate.Text = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    chk�Ӱ�.Value = IIf(IsNull(rsTmp!�Ӱ��־), 0, rsTmp!�Ӱ��־)
    
    'Ӥ����
    Call LoadPatientBaby(cboBaby, rsTmp!����ID, rsTmp!��ҳID)
    Call zlControl.CboLocate(cboBaby, rsTmp!Ӥ����, True)
    Call Set�����˿�������(cbo������, cbo��������, mrs������, mrs��������, Nvl(rsTmp!������), Nvl(rsTmp!��������ID, 0))
    '���˷�����Ϣ
    If Not IsNull(rsTmp!����ID) Then
        Set rsPatiMoney = GetMoneyInfo(rsTmp!����ID, , True, 2)
        If Not rsPatiMoney Is Nothing Then
           ' sta.Panels(3).Text = "Ԥ��:" & Format(rsPatiMoney!Ԥ�����, "0.00") & _
            "/����:" & Format(rsPatiMoney!�������, gstrDec) & _
            "/ʣ��:" & Format(rsPatiMoney!Ԥ����� - rsPatiMoney!�������, "0.00")
            Call SetStatuPatiInfor(Val(Nvl(rsPatiMoney!Ԥ�����)), Val(Nvl(rsPatiMoney!�������)), Val(Nvl(rsPatiMoney!Ԥ�����)) - Val(Nvl(rsPatiMoney!�������)))
        End If
    End If
    
    '------------------------------------------------------------------------------------
    If blnDelete Then
        '���ʵ���ȡʱ�������߱�,��ʱ����ں󱸱�,�ѽ���
        '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))
        
        '��ȡ������ԭʼ��¼�ķ���ID
        strSQL1 = _
            " Select A.ID,A.���,A.�շ�ϸĿID," & _
            " Nvl(A.����,1)*A.����" & IIf(gblnסԺ��λ, "/Nvl(B.סԺ��װ,1)", "") & " as ԭʼ����" & _
            " From סԺ���ü�¼ A,ҩƷ��� B" & _
            " Where A.NO=[1] And A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
            " And A.�շ�ϸĿID=B.ҩƷID(+) And A.��¼����=[2] And A.�����־=2 And Nvl(A.�ಡ�˵�,0)=0" & _
            IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[3]", "")

        '��ȡҩƷ�շ���¼�е�׼����
        strSQL2 = _
            " Select A.����ID,Sum(Nvl(A.����,1)*A.ʵ������" & IIf(gblnסԺ��λ, "/Nvl(B.סԺ��װ,1)", "") & ") as ׼������" & _
            " From ҩƷ�շ���¼ A,ҩƷ��� B" & _
            " Where A.NO=[1] And MOD(A.��¼״̬,3)=1" & _
            " And A.ҩƷID=B.ҩƷID(+) And A.���� IN(9,25) And A.����� is NULL" & _
            " Group by A.����ID"
        
        '���ŵ��ݻ��ܽ��(��ϸ���շ�ϸĿ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
        strSQL = "Select Nvl(�۸񸸺�,���) From סԺ���ü�¼ " & _
            " Where ��¼����=[2] And �����־=2 And Nvl(�ಡ�˵�,0)=0" & _
            " And ��¼״̬ IN(0,1,3) And NO=[1] And Nvl(ִ��״̬,0)<>1" & _
            IIf(mstrTime <> "", " And �Ǽ�ʱ��=[3]", "") & strFeeKind
        
        '����ѽ��ʵ��ݽ�ֹ����,����ҽ�����ʵĵ��ݡ�����ԭʼ��������ֻȡδ���ʲ���
        intInsure = BillExistInsure(strNO, , , bytType)
        If intInsure <> 0 Then
            blnDo = Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, , intInsure)
        Else
            blnDo = gbytBillOpt = 2
        End If
        If blnDo Then
            strSQL = strSQL & " And Nvl(�۸񸸺�,���) IN" & _
                " (" & _
                " Select Nvl(�۸񸸺�,���) as ���" & _
                " From סԺ���ü�¼ " & _
                " Where NO=[1] And mod(��¼����,10)=[2]" & _
                " Group by Nvl(�۸񸸺�,���)" & _
                " Having Sum(Nvl(���ʽ��,0))=0" & _
                " )"
        End If
        
        '��Ϊ�ǽ�Ҫ��������ʣ�������ģ����Բ�����ֱ����ʱ�����ƣ����������
        strSQL = _
            " Select A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���) as ���," & _
            " A.�շ�ϸĿID,C.����,C.���� as ���,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
            IIf(gblnסԺ��λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X.סԺ��λ)", "A.���㵥λ") & " as ���㵥λ," & _
            " Avg(Nvl(A.����,1)) as ����," & _
            " Avg(A.����" & IIf(gblnסԺ��λ, "/Nvl(X.סԺ��װ,1)", "") & ") as ����," & _
            " Sum(A.��׼����" & IIf(gblnסԺ��λ, "*Nvl(X.סԺ��װ,1)", "") & ") as ����," & _
            " Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, " & _
            " D.���� as ִ�в���,A.���ӱ�־,A.ҽ�����" & _
            " From סԺ���ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� X" & _
            " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID(+)" & _
            " And A.�շ�ϸĿID=X.ҩƷID(+) And A.��¼����=[2] And A.�����־=2 And Nvl(A.���ӱ�־,0)<>9 And Nvl(A.�ಡ�˵�,0)=0" & _
            " And A.NO=[1] And Nvl(A.�۸񸸺�,A.���) IN(" & strSQL & ")" & _
            " Group by A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���),A.�շ�ϸĿID,C.����,C.����,B.����," & _
            " B.���,Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־,A.ҽ�����,X.ҩƷID,X.סԺ��λ"
            
        '��������
        '��"׼������=ԭʼ����"ʱ,�����ű���
        '�ſ��Ѿ�ȫ���˷ѵ���(ִ��״̬=0��һ�ֿ���)
        '��ʣ��������׼�������������������
            '1.�޶�Ӧ���շ���¼(����ͨ���û򲻸������õ�����),��ʱӦ��ʣ������
            '2.�շ���¼����ȫ������,����ȫ��ִ��,SQL���ų����ּ�¼
        strSQL = _
            " Select A.���,A.�շ�ϸĿID,A.����,A.���,A.����,A.���,A.��������,A.���㵥λ," & _
            " Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Avg(A.����),1) as ׼�˸���," & _
            " Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Sum(A.����),Nvl(C.׼������,Sum(A.����*A.����))) as ׼������," & _
            " Nvl(C.׼������,Sum(A.����*A.����)) as ׼������,Sum(A.����*A.����) as ʣ������," & _
            " A.����,Sum(A.Ӧ�ս��) as ʣ��Ӧ��,Sum(A.ʵ�ս��) as ʣ��ʵ��,A.ִ�в���,A.���ӱ�־,A.ҽ�����" & _
            " From (" & strSQL & ") A,(" & strSQL1 & ") B,(" & strSQL2 & ") C" & _
            " Where A.���=B.��� And B.ID=C.����ID(+)" & _
            " Group by A.���,A.�շ�ϸĿID,A.����,A.���,A.����,A.���,A.��������," & _
            " A.���㵥λ,A.����,B.ԭʼ����,C.׼������,A.ִ�в���,A.���ӱ�־,A.ҽ�����" & _
            " Having Sum(A.����*A.����)<>0"
        If intInsure <> 0 Then
            'ҽ�����ݿ��ܲ�������,��������������(׼������=ԭʼ����)
            If Not gclsInsure.GetCapability(support�����ֳ�����ϸ, , intInsure) Then
                strSQL = strSQL & " And Nvl(C.׼������,Sum(A.����*A.����))=B.ԭʼ����"
            End If
        End If
            
        strSQL = _
        " Select A.���,A.����,A.���,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��,A.���," & _
        "       A.��������,A.���㵥λ,A.׼�˸��� as ����,A.׼������ as ����,A.����," & _
        "       A.ʣ��Ӧ��*(A.׼������/A.ʣ������) as Ӧ�ս��," & _
        "       A.ʣ��ʵ��*(A.׼������/A.ʣ������) as ʵ�ս��," & _
        "       A.ִ�в���,A.���ӱ�־,A.ҽ�����" & _
        " From (" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1" & _
        " Where A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       and A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        " Order by A.���"
    ElseIf mbytInState = 0 And gbytBilling = 2 Then
        '��ȡ���ʻ��۵�(�������),ֻ��ȡδ��˲���
        '���ÿ��ǴӺ󱸱�ȡ��
        strSQL = _
            " Select" & _
            " Nvl(A.�۸񸸺�,A.���) as ���,A.�շ�ϸĿID,C.����,C.���� as ���,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
            IIf(gblnסԺ��λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X.סԺ��λ)", "A.���㵥λ") & " as ���㵥λ," & _
            " Avg(Nvl(A.����,1)) as ����," & _
            " Avg(A.����" & IIf(gblnסԺ��λ, "/Nvl(X.סԺ��װ,1)", "") & ") as ����," & _
            " Sum(A.��׼����" & IIf(gblnסԺ��λ, "*Nvl(X.סԺ��װ,1)", "") & ") as ����," & _
            " Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, " & _
            " D.���� as ִ�в���,A.���ӱ�־" & _
            " From סԺ���ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� X" & _
            " Where A.��¼״̬=0 And A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID(+)" & _
            " And A.�շ�ϸĿID=X.ҩƷID(+) And A.��¼����=2 And Nvl(A.�ಡ�˵�,0)=0 And �����־=2 And A.NO=[1]" & _
            " Group by Nvl(A.�۸񸸺�,A.���),A.��¼״̬,A.�շ�ϸĿID,C.����,C.����,B.����,B.���," & _
            " Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־,X.ҩƷID,X.סԺ��λ"
        
        strSQL = _
        " Select A.���,A.����,A.���,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��,A.���," & _
        "   A.��������,A.���㵥λ,A.����,A.����,A.����,A.Ӧ�ս��,A.ʵ�ս��,A.ִ�в���,A.���ӱ�־" & _
        " From(" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1" & _
        " Where A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "   and A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        " Order by A.���"
    Else
        '��ȡ����ԭʼ����
        intSign = IIf(mblnDelete, -1, 1) '����,�����������
        strSQL = _
            " Select Nvl(A.�۸񸸺�,A.���) as ���,A.�շ�ϸĿID,C.����,C.���� as ���,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
            IIf(gblnסԺ��λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X.סԺ��λ)", "A.���㵥λ") & " as ���㵥λ," & _
            " Avg(Nvl(A.����,1)) as ����," & _
            " Avg(" & intSign & "*A.����" & IIf(gblnסԺ��λ, "/Nvl(X.סԺ��װ,1)", "") & ") as ����," & _
            " Sum(A.��׼����" & IIf(gblnסԺ��λ, "*Nvl(X.סԺ��װ,1)", "") & ") as ����," & _
            " Sum(" & intSign & "*A.Ӧ�ս��) as Ӧ�ս��,Sum(" & intSign & "*A.ʵ�ս��) as ʵ�ս��, " & _
            " D.���� as ִ�в���,A.���ӱ�־" & _
            " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼  A") & ",�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� X" & _
            " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID(+)" & _
            " And A.�շ�ϸĿID=X.ҩƷID(+) And A.��¼����=[2] And Nvl(A.�ಡ�˵�,0)=0 And �����־=2" & _
            " And A.��¼״̬" & IIf(mblnDelete, "=2", " IN(0,1,3)") & " And A.NO=[1]" & _
            IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[3]", "") & _
            " Group by Nvl(A.�۸񸸺�,A.���),A.�շ�ϸĿID,C.����,C.����,B.����,B.���," & _
            " Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.���ӱ�־,X.ҩƷID,X.סԺ��λ"
            
        strSQL = "" & _
        " Select A.���,A.����,A.���,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��,A.���," & _
        "       A.��������,A.���㵥λ,A.����,A.����,A.����,A.Ӧ�ս��,A.ʵ�ս��,A.ִ�в���,A.���ӱ�־" & _
        " From(" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1" & _
        " Where A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       and A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
        " Order by A.���"
    End If
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, bytType, CDate(mstrTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, bytType)
    End If
    If rsTmp.EOF Then Exit Function
    
    '��ʿվ����ʱ����ȡҽ����ӦҪȱʡ���ʵķ�����
    If blnDelete And mlngҽ��ID <> 0 Then
        strAdvice = GetAdviceIDs(mlngҽ��ID)
    End If
    
    Bill.Redraw = False
    Bill.Rows = rsTmp.RecordCount + 1
    For i = 1 To rsTmp.RecordCount
        If gbytBilling = 2 And Not mblnPrint Then mblnPrint = True
        Bill.RowData(i) = rsTmp!��� '���ڼ��������Լ��������
        
        Bill.TextMatrix(i, BillCol.���) = rsTmp!���
        Bill.TextMatrix(i, BillCol.��Ŀ) = rsTmp!����
        Bill.TextMatrix(i, BillCol.��Ʒ��) = Nvl(rsTmp!��Ʒ��)
        Bill.TextMatrix(i, BillCol.���) = IIf(IsNull(rsTmp!���), "", rsTmp!���)
        Bill.TextMatrix(i, BillCol.��λ) = IIf(IsNull(rsTmp!���㵥λ), "", rsTmp!���㵥λ)
        Bill.TextMatrix(i, BillCol.����) = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        Bill.TextMatrix(i, BillCol.����) = FormatEx(Val(Nvl(rsTmp!����)), 5)
        Bill.TextMatrix(i, BillCol.����) = Format(Val(Nvl(rsTmp!����)), gstrFeePrecisionFmt)
        Bill.TextMatrix(i, BillCol.Ӧ�ս��) = Format(Val(Nvl(rsTmp!Ӧ�ս��)), gstrDec)
        Bill.TextMatrix(i, BillCol.ʵ�ս��) = Format(Val(Nvl(rsTmp!ʵ�ս��)), gstrDec)
        Bill.TextMatrix(i, BillCol.ִ�п���) = Nvl(rsTmp!ִ�в���)
        Bill.TextMatrix(i, BillCol.��־) = IIf(rsTmp!���ӱ�־ = 1, "��", "")
        Bill.TextMatrix(i, BillCol.����) = IIf(IsNull(rsTmp!��������), "", rsTmp!��������)
        
        '�������ʱ�־
        If Bill.TextMatrix(0, Bill.Cols - 1) = "����" Then
            If strAdvice <> "" Then
                If InStr("," & strAdvice & ",", "," & rsTmp!ҽ����� & ",") > 0 Then
                    Bill.TextMatrix(i, Bill.Cols - 1) = "��"
                End If
            Else
                If mlngDelRow = 0 Or mlngDelRow <> 0 And mlngDelRow = rsTmp!��� Then
                    Bill.TextMatrix(i, Bill.Cols - 1) = "��"
                End If
            End If
        End If
        
        rsTmp.MoveNext
    Next
    '����б༭����������ɫ
    Call InitBillColumnColor
    Call SetColNum
    Bill.Redraw = True
    
    '----------------------------------------------------------------------------
    '������Ŀ�б���ʾ
    If blnDelete Then
         '�˷ѵ����迼�Ǻ󱸱�,ǰ��Ĳ����ѽ�ֹ
        '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))

        '��ȡҩƷ�շ���¼�е�׼����
        strSQL1 = _
            " Select A.����ID,Sum(Nvl(A.����,1)*A.ʵ������" & IIf(gblnסԺ��λ, "/Nvl(B.סԺ��װ,1)", "") & ") as ׼������" & _
            " From ҩƷ�շ���¼ A,ҩƷ��� B" & _
            " Where A.NO=[1] And MOD(A.��¼״̬,3)=1" & _
            " And A.ҩƷID=B.ҩƷID(+) And A.����� is NULL And A.���� IN(9,25)" & _
            " Group by A.����ID"
        
        '���ŷ��õ���(��ϸ��������Ŀ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        strSQL = "Select Nvl(�۸񸸺�,���) From סԺ���ü�¼ " & _
            " Where ��¼����=[2] And �����־=2 And Nvl(�ಡ�˵�,0)=0" & _
            " And ��¼״̬ IN(0,1,3) And NO=[1] And Nvl(ִ��״̬,0)<>1" & _
            IIf(mstrTime <> "", " And �Ǽ�ʱ��=[3]", "") & strFeeKind
            
        If blnDo Then
            strSQL = strSQL & " And Nvl(�۸񸸺�,���) IN" & _
                " (" & _
                " Select Nvl(�۸񸸺�,���) as ���" & _
                " From סԺ���ü�¼ " & _
                " Where NO=[1] And mod(��¼����,10)=[2]" & _
                " Group by Nvl(�۸񸸺�,���)" & _
                " Having Sum(Nvl(���ʽ��,0))=0" & _
                " )"
        End If
            
        strSQL = _
            " Select Sum(A.ID) as ID,A.���,A.����,A.�շ����," & _
            " Sum(A.����) as ʣ������,Sum(A.Ӧ�ս��) as ʣ��Ӧ��," & _
            " Sum(A.ʵ�ս��) as ʣ��ʵ�� From (" & _
            " Select Decode(A.��¼״̬,2,0,A.ID) as ID,A.���,B.����,A.�շ����," & _
            " Nvl(A.����,1)*A.����" & IIf(gblnסԺ��λ, "/Nvl(X.סԺ��װ,1)", "") & " as ����," & _
            " A.Ӧ�ս��,A.ʵ�ս��" & _
            " From סԺ���ü�¼ A,������Ŀ B,ҩƷ��� X" & _
            " Where A.��¼����=[2] And A.NO=[1]" & _
            " And A.������ĿID=B.ID And Nvl(A.�۸񸸺�,A.���) IN(" & strSQL & ")" & _
            " And A.�շ�ϸĿID=X.ҩƷID(+) And A.�����־=2 And Nvl(A.�ಡ�˵�,0)=0) A" & _
            " Group by A.���,A.����,A.�շ����" & _
            " Having Sum(A.����)<>0"
                    
        '��������
        '��ʣ��������׼�������������������
            '1.�޶�Ӧ���շ���¼(����ͨ���û򲻸������õ�����),��ʱӦ��ʣ������
            '2.�շ���¼����ȫ������,����ȫ��ִ��,SQL���ų����ּ�¼
        strSQL = _
            " Select A.����,Sum(A.ʣ��Ӧ��*(A.׼������/A.ʣ������)) as Ӧ�ս��," & _
            " Sum(ʣ��ʵ��*(A.׼������/A.ʣ������)) as ʵ�ս�� From (" & _
            " Select A.����,A.ʣ������,A.ʣ��Ӧ��,A.ʣ��ʵ��," & _
            " Decode(Instr(',4,5,6,7,',A.�շ����),0,A.ʣ������,Nvl(B.׼������,A.ʣ������)) as ׼������" & _
            " From (" & strSQL & ") A,(" & strSQL1 & ") B" & _
            " Where A.ID=B.����ID(+)" & _
            " ) A Group by A.����"
    ElseIf mbytInState = 0 And gbytBilling = 2 Then
        '��ȡ���ʻ��۵�(�������),ֻ��ȡδ��˲���
        '���ÿ��Ǻ󱸱�ȡ��
        strSQL = _
            "Select B.����,Sum(A.Ӧ�ս��) as Ӧ�ս��," & _
            " Sum(A.ʵ�ս��) as ʵ�ս�� " & _
            " From סԺ���ü�¼ A,������Ŀ B" & _
            " Where A.��¼����=2 And A.�����־=2 And Nvl(A.�ಡ�˵�,0)=0" & _
            " And A.��¼״̬=0 And A.NO=[1] And A.������ĿID=B.ID" & _
            " Group By B.����"
    Else
        '��ȡ����ԭʼ����
        intSign = IIf(mblnDelete, -1, 1) '����,�����������
        strSQL = _
            "Select B.����,Sum(" & intSign & "*A.Ӧ�ս��) as Ӧ�ս��," & _
            " Sum(" & intSign & "*A.ʵ�ս��) as ʵ�ս�� " & _
            " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & ",������Ŀ B" & _
            " Where A.��¼״̬" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
            " And A.��¼����=[2] And A.�����־=2 And Nvl(A.�ಡ�˵�,0)=0" & _
            IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[3]", "") & _
            " And A.NO=[1] And A.������ĿID=B.ID Group By B.����"
    End If
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, bytType, CDate(mstrTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, bytType)
    End If
    If rsTmp.EOF Then Exit Function
    
    'ˢ����ʾ(�շ�Ҫ����)
    mshMoney.Rows = rsTmp.RecordCount + 1
    If mshMoney.Rows < 5 Then mshMoney.Rows = 5
    Call SetMoneyList
    
    For i = 1 To rsTmp.RecordCount
        mshMoney.TextMatrix(i, 0) = rsTmp!����
        mshMoney.TextMatrix(i, 1) = Format(Val(Nvl(rsTmp!ʵ�ս��)), gstrDec)
        curTotal = curTotal + Val(Nvl(rsTmp!ʵ�ս��))
        curӦ��Total = curӦ��Total + Val(Nvl(rsTmp!Ӧ�ս��))
        rsTmp.MoveNext
    Next
    
    txtʵ��.Text = Format(curTotal, gstrDec)
    txtӦ��.Text = Format(curӦ��Total, gstrDec)
    Debug.Print Me.txtPatient.Text
    
    ReadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Function GetAdviceIDs(ByVal lngҽ��ID As Long) As String
'���ܣ���ȡһ��ҽ��������ҽ����¼ID��
'������lngҽ��ID=һ��ҽ����¼����ID:Nvl(���ID,ID)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ID From ����ҽ����¼ Where ID=[1] Or ���ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    
    GetAdviceIDs = Mid(strSQL, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetShowCol()
'���ܣ������еĿ���(���ʱչ��)
    mrsClass.Filter = "����='7'"
    If mrsClass.RecordCount = 0 Then
        Bill.ColWidth(BillCol.����) = 0
    ElseIf Bill.ColWidth(BillCol.����) = 0 Then
        Bill.ColWidth(BillCol.����) = 520
    End If
End Sub
Private Sub InitBillColumnColor()
    Bill.SetColColor BillCol.���, &HE7CFBA
    Bill.SetColColor BillCol.��Ŀ, &HE7CFBA
    Bill.SetColColor BillCol.����, &HE7CFBA
    Bill.SetColColor BillCol.ִ�п���, &HE7CFBA
    Bill.SetColColor BillCol.����, &HE0E0E0
    Bill.SetColColor BillCol.����, &HE0E0E0
    Bill.SetColColor BillCol.��־, &HE0E0E0
End Sub

Private Sub ClearRows()
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
    Next
End Sub

Private Function GetPay(lngRow As Long) As Integer
    Dim i As Long
    'ȡ������ҩ�ĸ���
    GetPay = 1
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).�շ���� = "7" And i <> lngRow Then
            GetPay = mobjBill.Details(i).����
            Exit For
        End If
    Next
End Function

Private Function GetDetailNum(lngRow As Long) As Double
'���ܣ���ȡ����ָ��ϸĿ���ܼ�������(����������)
'������lngRow=��ǰ������
    Dim rsTmp As ADODB.Recordset
    Dim lngNum As Long, i As Long
    Dim strSQL As String
        
    If lngRow <= mobjBill.Details.Count And mrsInfo.State = 1 Then
        '��ǰ�����е�����
        For i = 1 To mobjBill.Details.Count
            If i <> lngRow And mobjBill.Details(i).�շ�ϸĿID = mobjBill.Details(lngRow).�շ�ϸĿID Then
                lngNum = lngNum + mobjBill.Details(i).���� * IIf(mobjBill.Details(i).���� = 0, 1, mobjBill.Details(i).����)
            End If
        Next
        
        '���ݿ��е�����
        strSQL = _
            "Select Sum(A.����*Nvl(A.����,1)" & IIf(gblnסԺ��λ, "/Nvl(B.סԺ��װ,1)", "") & ") as Num" & _
            " From סԺ���ü�¼ A,ҩƷ��� B" & _
            " Where A.�۸񸸺� is Null And A.���ʷ���=1" & _
            IIf(gbytBilling = 0, " And A.��¼״̬<>0", "") & _
            " And A.����ID=[1] And Nvl(A.��ҳID,0)=[2] And A.�շ�ϸĿID=B.ҩƷID(+) And A.�շ�ϸĿID+0=[3]"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(mrsInfo!����ID), Val("" & mrsInfo!��ҳID), mobjBill.Details(lngRow).�շ�ϸĿID)
        If Not rsTmp.EOF Then
            lngNum = lngNum + Nvl(rsTmp!Num, 0)
        End If
        GetDetailNum = lngNum
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetWorkUnit(ByVal lngҩƷID As Long, ByVal str��� As String) As Boolean
'���ܣ�ȡ���пɹ�ѡ���ҩ��
    Dim strSQL As String, strҩ�� As String, bytDay As Byte
    Dim int������� As Integer, str������� As String
    Dim int������Դ As Integer, lng��������ID As Long
    
    '������Ŀ��Ȩ��ȷ��ҩ���ķ������
    int������� = Get�������(lngҩƷID)
    
    If int������� = 1 Then
        str������� = "1,3"
    ElseIf int������� = 2 Then
        str������� = "2,3"
    ElseIf int������� = 3 Then
        If InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
            str������� = "1,2,3"
        Else
            str������� = "2,3"
        End If
    Else
            str������� = "2,3"
    End If
    
    'ȷ��������Դ
    If mrsInfo.State = 1 Then
        If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
            int������Դ = 2
        ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
            int������Դ = 1
        End If
    Else
        int������Դ = 2
    End If
    
    lng��������ID = mobjBill.����ID
    If lng��������ID = 0 And cbo��������.ListIndex <> -1 Then lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
       
    If str��� = "4" Then
        strSQL = _
        "Select Distinct c.Id, c.����, c.����, c.����, b.��������, b.�������, Decode(a.��������id,[2],0,1) As ����" & vbNewLine & _
        "From �շ�ִ�п��� A, ��������˵�� B, ���ű� C" & vbNewLine & _
        "Where a.ִ�п���id + 0 = b.����id And b.�������� = '���ϲ���' And b.������� IN(" & str������� & ") And b.����id = c.Id And" & vbNewLine & _
        "      (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And (c.վ�� = '" & gstrNodeNo & "' Or c.վ�� Is Null) And" & vbNewLine & _
        "      (a.������Դ Is Null Or a.������Դ = [1]) And" & vbNewLine & _
        "      (a.��������id Is Null Or a.��������id = [2] Or Exists (Select 1 From �������Ҷ�Ӧ Where ����id = [2] And a.��������id = ����id)) And a.�շ�ϸĿid = [3]" & vbNewLine & _
        "Order By b.�������, ����, c.����"

    Else
        '��ҩƷ����ȷ��ҩ������
        Select Case str���
            Case "5"
                strҩ�� = "��ҩ��"
            Case "6"
                strҩ�� = "��ҩ��"
            Case "7"
                strҩ�� = "��ҩ��"
        End Select
        
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If Not gblnҩ���ϰల�� Then
            strSQL = _
            " Select Distinct C.ID,C.����,C.����,C.����,B.��������,B.������� " & _
            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������=[4]" & _
            "       And B.������� IN(" & str������� & ") And B.����ID=C.ID" & _
            "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
            "       And (A.������Դ is NULL Or A.������Դ=[1])" & _
            "       And (A.��������ID is NULL Or A.��������ID=[2])" & _
            "       And A.�շ�ϸĿID=[3]" & _
            " Order by B.�������,C.����"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSQL = _
            " Select Distinct C.ID,C.����,C.����,C.����,B.��������,B.������� " & _
            " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������=[4]" & _
            "       And B.������� IN(" & str������� & ") And B.����ID=C.ID" & _
            "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
            "       And D.����ID=C.ID And D.����=[5]" & _
            "       And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
            "       And (A.������Դ is NULL Or A.������Դ=[1])" & _
            "       And (A.��������ID is NULL Or A.��������ID=[2])" & _
            "       And A.�շ�ϸĿID=[3]" & _
            " Order by B.�������,C.����"
        End If
    End If
    
    On Error GoTo errH
    Set mrsWork = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, int������Դ, lng��������ID, lngҩƷID, strҩ��, bytDay)
    GetWorkUnit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function CalcGridToTal(Optional blnӦ�� As Boolean) As Currency
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome
    Dim i As Long, intCol As Integer

    If mobjBill.Details.Count > 0 Then
        For Each objTmpDetail In mobjBill.Details
            For Each objTmpIncome In objTmpDetail.InComes
                If blnӦ�� Then
                    CalcGridToTal = CalcGridToTal + objTmpIncome.Ӧ�ս��
                Else
                    CalcGridToTal = CalcGridToTal + objTmpIncome.ʵ�ս��
                End If
            Next
        Next
    Else
        For i = 1 To Bill.Cols - 1
            If blnӦ�� Then
                If Bill.TextMatrix(0, i) = "Ӧ�ս��" Then intCol = i: Exit For
            Else
                If Bill.TextMatrix(0, i) = "ʵ�ս��" Then intCol = i: Exit For
            End If
        Next
    
        For i = 1 To Bill.Rows - 1
            CalcGridToTal = CalcGridToTal + Val(Bill.TextMatrix(i, intCol))
        Next
    End If
End Function

Private Sub ShowDeleteCol(blnShow As Boolean)
'���ܣ���ʾ\�������ʱ�־��
    Dim i As Long, blnACT As Boolean
    If blnShow Then
        If Bill.TextMatrix(0, Bill.Cols - 1) <> "����" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols + 1
            Bill.TextMatrix(0, Bill.Cols - 1) = "����"
            Bill.ColAlignment(Bill.Cols - 1) = 4
            Bill.ColWidth(Bill.Cols - 1) = 550
            Bill.ColData(Bill.Cols - 1) = BillColType.CheckBox
            
            blnACT = Bill.Active: Bill.Active = False
            Bill.Row = 0: Bill.Col = Bill.Cols - 1: Bill.MsfObj.CellForeColor = vbRed
            Bill.Row = 1: Bill.Col = Bill.Cols - 1
            Bill.Active = blnACT
            
            Bill.ColWidth(BillCol.���) = GetOrigColWidth(BillCol.���) - 120
            Bill.ColWidth(BillCol.��Ŀ) = GetOrigColWidth(BillCol.��Ŀ) - 100
            Bill.ColWidth(BillCol.ִ�п���) = GetOrigColWidth(BillCol.ִ�п���) - 200
            
            Bill.ColWidth(BillCol.����) = GetOrigColWidth(BillCol.����) - 50
            Bill.ColWidth(BillCol.Ӧ�ս��) = GetOrigColWidth(BillCol.Ӧ�ս��) - 50
            Bill.ColWidth(BillCol.ʵ�ս��) = GetOrigColWidth(BillCol.ʵ�ս��) - 50
            Bill.Redraw = True
        End If
    Else
        If Bill.TextMatrix(0, Bill.Cols - 1) = "����" Then
            Bill.Redraw = False
            Bill.Cols = Bill.Cols - 1
            Bill.ColWidth(BillCol.���) = GetOrigColWidth(BillCol.���)
            Bill.ColWidth(BillCol.��Ŀ) = GetOrigColWidth(BillCol.��Ŀ)
            Bill.ColWidth(BillCol.ִ�п���) = GetOrigColWidth(BillCol.ִ�п���)
            
            Bill.ColWidth(BillCol.����) = GetOrigColWidth(BillCol.����)
            Bill.ColWidth(BillCol.Ӧ�ս��) = GetOrigColWidth(BillCol.Ӧ�ս��)
            Bill.ColWidth(BillCol.ʵ�ս��) = GetOrigColWidth(BillCol.ʵ�ս��)
            Bill.Redraw = True
        End If
    End If
End Sub

Private Function GetOrigColWidth(ByVal intIdx As Integer) As Long
'���ܣ���ȡָ���е�ԭʼ�п�
    GetOrigColWidth = Val(Split(Split(STR_HEAD, ";")(intIdx), ",")(1))
End Function

Private Function SaveModi() As Boolean
'���ܣ����浱ǰ�޸ĵķ��õ���
    Dim strSQL As String
    
    strSQL = "zl_���˷��ü�¼_Update('" & cboNO.Text & "',2,'" & zlStr.NeedName(cbo������.Text) & "'," & _
        "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'),NULL,2)"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveModi = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetColNum(Optional intRow As Long = 1)
'���ܣ�������ʾ���е��к�
'������intRow=�Ӹ��п�ʼ
    Dim bln As Boolean, i As Long
    
    Bill.Redraw = False
    For i = intRow To Bill.Rows - 1
        Bill.TextMatrix(i, BillCol.��) = i
    Next
    Bill.Redraw = True
End Sub

Private Function CheckDuty(Optional tmpDetail As Detail, Optional blnCommon As Boolean = True) As Integer
'���ܣ����ָ��ҩƷ�е�ְ���Ƿ��뵱ǰҽ����ְ����ƥ��
'������tmpDetail=�������Ŀ,����Ϊ������,blnCommon=�Ƿ��������ж�,����Ϊҽ���򹫷Ѳ��˵��ж�
'���أ���ƥ�����,0Ϊ��ȷ
'˵����ְ��1=����,2=����,3=�м�,4=����/ʦ��,5=Ա/ʿ,9=��Ƹ
    Dim i As Long, intְ��A As Integer, intְ��B As Integer
    Dim strTmp As String, strAllDuty As String
    
    If cbo������.ListIndex = -1 Then Exit Function
    strAllDuty = "����,����,�м�,����/ʦ��,Ա/ʿ,,,,��Ƹ"
    Call GetOperatorInfo(mrs������, mobjBill.������, , intְ��A)
        
    If tmpDetail Is Nothing Then
        For i = 1 To mobjBill.Details.Count
            If InStr(",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
                If Not blnCommon Then
                    intְ��B = Val(Right(mobjBill.Details(i).Detail.����ְ��, 1))
                    If intְ��B > 0 Then
                        If intְ��A = 0 Then
                            strTmp = "��ҽ���򹫷Ѳ���,�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                            CheckDuty = 1
                        ElseIf intְ��B < intְ��A Then
                            strTmp = "��ҽ���򹫷Ѳ���,�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��A - 1) & """��"
                            CheckDuty = i: Exit For
                        End If
                    End If
                Else
                    intְ��B = Val(Left(mobjBill.Details(i).Detail.����ְ��, 1))
                    If intְ��B > 0 Then
                        If intְ��A = 0 Then
                            strTmp = "�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                            CheckDuty = 1
                        ElseIf intְ��B < intְ��A Then
                            strTmp = "�� " & i & " ��ҩƷ""" & mobjBill.Details(i).Detail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��A - 1) & """��"
                            CheckDuty = i: Exit For
                        End If
                    End If
                End If
            End If
        Next
    Else
        If InStr(",5,6,7,", tmpDetail.���) = 0 Then Exit Function
        If Not blnCommon Then
            intְ��B = Val(Right(tmpDetail.����ְ��, 1))
            If intְ��B > 0 Then
                If intְ��A = 0 Then
                    strTmp = "��ҽ���򹫷Ѳ���,ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                    CheckDuty = 1
                ElseIf intְ��B < intְ��A Then
                    strTmp = "��ҽ���򹫷Ѳ���,ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��A - 1) & """��"
                    CheckDuty = 1
                End If
            End If
        Else
            intְ��B = Val(Left(tmpDetail.����ְ��, 1))
            If intְ��B > 0 Then
                If intְ��A = 0 Then
                    strTmp = "ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ������Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """,����ǰҽ��δ����ְ��"
                    CheckDuty = 1
                ElseIf intְ��B < intְ��A Then
                    strTmp = "ҩƷ""" & tmpDetail.���� & """Ҫ��ҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��B - 1) & """����,����ǰҽ��ְ��Ϊ""" & Split(strAllDuty, ",")(intְ��A - 1) & """��"
                    CheckDuty = 1
                End If
            End If
        End If
    End If
    If CheckDuty > 0 Then MsgBox strTmp, vbInformation, gstrSysName
End Function

Private Function PhysicExist(objDetail As Detail, intRow As Integer, Optional blnCancel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ��ҩƷ�ڵ������Ƿ��Ѿ�����
    '���:objDetail=��Ŀ,intRow=Ҫ�жϵ���
    '����:blnCancel-�����Ҫǿ��ȡ��������true
    '����:true��ʾ��������ʾ��false-��ʾ�Ϸ�
    '����:���˺�
    '����:2017-11-22 17:23:06
    '˵��:ʱ�ۻ����ҩƷ��ͬһҩ����ֹ�ظ�����(�������ʾ,����ʱ��ֹ(blnCancel=trueʱ����))
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    blnCancel = False
     
    For i = 1 To mobjBill.Details.Count
        If i <> intRow And InStr(",4,5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
            If mobjBill.Details(i).Detail.ID = objDetail.ID Then
                If (mobjBill.Details(i).Detail.���� Or mobjBill.Details(i).Detail.���) _
                    And (objDetail.���� Or objDetail.���) Then
                    If objDetail.��� = "4" Then
                        If mobjBill.Details(i).Detail.���� <> objDetail.���� And objDetail.���� >= 0 Then
                           'ɨ����ģ�����ˢ������ε�
                        Else
                            If mobjBill.Details(i).Detail.���� = objDetail.���� And objDetail.���� > 0 Then
                                Call MsgBox("��������""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�������ͬ������,��ֹ����!", vbInformation + vbOKOnly, gstrSysName)
                                blnCancel = True
                                PhysicExist = True
                                Exit Function
                            End If
                            If MsgBox("��������""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������" & _
                                vbCrLf & vbCrLf & "ע�⣺����������Ϊ������ʱ�۲���,�ظ�����ʱ���뱣֤���ǵķ��ϲ��Ų�ͬ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                PhysicExist = True
                            End If
                        End If
                    Else
                        If MsgBox("ҩƷ""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������" & _
                            vbCrLf & vbCrLf & "ע�⣺��ҩƷΪ������ʱ��ҩƷ,�ظ�����ʱ���뱣֤���ǵ�ִ��ҩ����ͬ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                Else
                    If objDetail.��� = "4" Then
                        If MsgBox("��������""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    Else
                        If MsgBox("ҩƷ""" & objDetail.���� & """�ڵ� " & i & " ���Ѿ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            PhysicExist = True
                        End If
                    End If
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function Check��������(Optional intRow As Integer) As Boolean
'���ܣ����ݵ�ǰ���˵������ж�ָ���е���Ŀ�Ƿ��������,����������������Ŀ
    Dim strSQL As String, bytType As Byte
    Dim i As Integer
    Dim blnҽ�� As Boolean, bln���� As Boolean
    
    Check�������� = True
    
    On Error GoTo errHandle
    

    '�޷����
    If cboҽ�Ƹ���.ListIndex = -1 Then Exit Function
    'ҽ���򹫷Ѳ���
    '����:45605
    'ֻ���ҽ�����˺͹��Ѳ���
    If zlIsCheckMedicinePayMode(zlStr.NeedName(cboҽ�Ƹ���), blnҽ��, bln����) = False Then Exit Function
    'ȷ����������
    bytType = IIf(blnҽ��, 1, 2)
    
    '��ȡ�������
    If mrs�������� Is Nothing Then
        strSQL = " Select 'ҽ��' As ���,����,���� From �������� Where ���� In(" & gstrҽ���������� & ") Union All " & _
                 " Select '����' As ���,����,���� From �������� Where ���� In(" & gstr���ѷ������� & ") "
        Set mrs�������� = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(mrs��������, strSQL, Me.Caption)
    End If
    mrs��������.Filter = ""
    If mrs��������.RecordCount = 0 Then Exit Function
        
    If bytType = 1 Then
        strSQL = " And ���='ҽ��'"
    Else
        strSQL = " And ���='����'"
    End If
    
    If intRow > 0 Then
        If mobjBill.Details(intRow).Detail.���� = "" Then
            MsgBox """" & mobjBill.Details(intRow).Detail.���� & """�ķ�������δ���ã�", vbInformation, gstrSysName
            Check�������� = False
        Else
            mrs��������.Filter = "����='" & mobjBill.Details(intRow).Detail.���� & "'" & strSQL
            If mrs��������.EOF Then
                MsgBox """" & mobjBill.Details(intRow).Detail.���� & """�ķ�������Ϊ""" & _
                    mobjBill.Details(intRow).Detail.���� & """,����" & _
                    IIf(bytType = 1, "ҽ��", "����") & "�������ͣ�", vbInformation, gstrSysName
                Check�������� = False
            End If
        End If
    Else
        For i = 1 To mobjBill.Details.Count
            If mobjBill.Details(i).Detail.���� = "" Then
                If MsgBox("�����е� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """�ķ�������δ���ã�" & vbCrLf & "ȷʵҪ���浥����", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Check�������� = False: Exit For
                End If
            Else
                mrs��������.Filter = "����='" & mobjBill.Details(i).Detail.���� & "'" & strSQL
                If mrs��������.EOF Then
                    If MsgBox("�����е� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """�ķ�������Ϊ""" & _
                        mobjBill.Details(i).Detail.���� & """,����" & _
                        IIf(bytType = 1, "ҽ��", "����") & "�������ͣ�" & vbCrLf & "ȷʵҪ���浥����", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Check�������� = False: Exit For
                    End If
                End If
            End If
        Next
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ReCalcInsure()
'���ܣ��޸ĵ���ʱ,���¼���ͳ������������Ϣ
    Dim i As Long, j As Long, dblAllTime As Double
    Dim strInfo As String
    
    If mrsInfo.State = 1 Then
        If Not IsNull(mrsInfo!����) Then
            For i = 1 To mobjBill.Details.Count
                For j = 1 To mobjBill.Details(i).InComes.Count
                    dblAllTime = mobjBill.Details(i).���� * mobjBill.Details(i).����
                    If InStr(",5,6,7,", mobjBill.Details(i).�շ����) > 0 Then
                        If gblnסԺ��λ Then dblAllTime = dblAllTime * mobjBill.Details(i).Detail.סԺ��װ
                    End If
                    
                    strInfo = gclsInsure.GetItemInsure(mobjBill.����ID, mobjBill.Details(i).�շ�ϸĿID, mobjBill.Details(i).InComes(j).ʵ�ս��, False, mrsInfo!����, _
                        mobjBill.Details(i).ժҪ & "||" & dblAllTime)
                    If strInfo <> "" Then
                        mobjBill.Details(i).������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                        mobjBill.Details(i).���մ���ID = Val(Split(strInfo, ";")(1))
                        mobjBill.Details(i).InComes(j).ͳ���� = Val(Split(strInfo, ";")(2))
                        mobjBill.Details(i).���ձ��� = CStr(Split(strInfo, ";")(3))
                        
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(i).ժҪ = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(i).Detail.���� = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                Next
            Next
        End If
    End If
End Sub

Private Sub chkIn_Click()
    sta.Panels(2) = ""
    If chkIn.Value = Checked Then
        txtIn.Enabled = True
        txtIn.BackColor = &H80000005
        sta.Panels(2) = "������Ҫ����ļ��ʵ����ݺ���"
        txtIn.SetFocus
    Else
        txtIn.Text = ""
        txtIn.Enabled = False
        txtIn.BackColor = &HE0E0E0
        Bill.SetFocus
    End If
End Sub


Private Sub txtIn_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim tmpBill As New ExpenseBill
    Dim i As Long, strSQL As String
    Dim lng����ID As Long, curTotal As Currency
    Dim lngPre As Long, strPre As String
    Dim blnHavePatient As Boolean
    Dim Curdate As Date     '��������ǰʱ��
    
    On Error GoTo errH
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtIn, KeyAscii)
    Else
        txtIn.Text = GetFullNO(txtIn.Text, 14)
        
        Set tmpBill = ImportBill(txtIn.Text, False, Me, False, gblnסԺ��λ, , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
        If tmpBill.NO = "" Then
            MsgBox "��ȡ����ʧ�ܡ�", vbExclamation, gstrSysName
            txtIn.Text = "": txtIn.SetFocus: Exit Sub
        Else
            '�����޸ļ���ʾ
            Screen.MousePointer = 11
                        
            lng����ID = tmpBill.����ID
            lngPre = tmpBill.��������ID
            strPre = tmpBill.������
            If InStr(mstrPrivsOpt, ";ҽ����ѯ;") = 0 Then strPre = ""
            
            '�������Ĳ�����Ϣ
            tmpBill.����ID = 0
            tmpBill.��ҳID = 0
            tmpBill.���� = ""
            tmpBill.��ʶ�� = 0
            tmpBill.���� = ""
            tmpBill.�Ա� = ""
            tmpBill.���� = ""
            tmpBill.�ѱ� = ""
            tmpBill.����ID = 0
            tmpBill.����ID = 0
            
            '���˺�:25882
            For i = 1 To tmpBill.Details.Count
                tmpBill.Details(i).����ID = 0
                tmpBill.Details(i).��ҳID = 0
                tmpBill.Details(i).���� = ""
                tmpBill.Details(i).�Ա� = ""
                tmpBill.Details(i).���� = ""
                tmpBill.Details(i).�ѱ� = ""
                tmpBill.Details(i).����ID = 0
                tmpBill.Details(i).����ID = 0
            Next
            
            '�������в�����Ϣ
            If Not mobjBill Is Nothing Then
                If mobjBill.����ID > 0 Then
                    lng����ID = mobjBill.����ID
                    lngPre = mobjBill.��������ID
                    strPre = mobjBill.������
                    blnHavePatient = True
                End If
            End If
            
            Set mobjBill = New ExpenseBill
            Set mobjBill = tmpBill
            
            Curdate = zlDatabase.Currentdate
            mobjBill.NO = cboNO.Text
            mobjBill.�Ǽ�ʱ�� = Curdate
            mobjBill.����Ա��� = UserInfo.���
            mobjBill.����Ա���� = UserInfo.����
            mobjBill.�Ӱ��־ = chk�Ӱ�.Value
            mobjBill.Ӥ���� = cboBaby.ItemData(cboBaby.ListIndex)
            
            'ȡ��ǰʱ��
            txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
            'ȡ��ǰʱ��:33744
            If mbln���� And mstr���ת��ʱ�� <> "" Then
                txtDate.Text = Format(CDate(mstr���ת��ʱ��) - 1 / 24 / 60, "yyyy-mm-dd HH:MM:SS")
                txtDate.ForeColor = vbBlue
            End If
            
            Bill.Redraw = False
            Bill.ClearBill
            Bill.Rows = mobjBill.Details.Count + 1
            
            Call InitBillColumnColor
            
            '���ʷ��౨��
            mstrWarn = ""
            
            If lng����ID <> 0 Then
                mbln������۸� = True
                txtPatient.Text = "-" & lng����ID
                Call txtPatient_KeyPress(13)            '���ܻ�ı俪���˺Ϳ�������
                mbln������۸� = False
            End If
            
            mobjBill.��������ID = lngPre
            mobjBill.������ = strPre
            Call Set�����˿�������(cbo������, cbo��������, mrs������, mrs��������, mobjBill.������, mobjBill.��������ID)
            
            
            '������Ķ����˺�ȷ���ѱ��,�ټ���۸�
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
            Bill.Redraw = True
            chkIn.Value = 0
            
            
            'ˢ�²��˷�����Ϣ
            If mrsInfo.State = 1 Then
                'ˢ�²���Ԥ������Ϣ
                curTotal = GetBillTotal(mobjBill)
                Set rsTmp = GetMoneyInfo(mrsInfo!����ID, 0, True, 2)
                If Not rsTmp Is Nothing Then
                    cmdOK.Tag = rsTmp!Ԥ�����
                    cmdCancel.Tag = rsTmp!�������
                    txtʵ��.Tag = rsTmp!Ԥ����� - rsTmp!�������
                Else
                    cmdOK.Tag = 0: cmdCancel.Tag = 0: txtʵ��.Tag = 0
                End If
                '����ʱ��ʾ���㵱ǰ���ݷ���,�����۱���Ҫ��
                'sta.Panels(3).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
               ' sta.Panels(3).Text = sta.Panels(3).Text & "/����:" & Format(Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), gstrDec)
               ' sta.Panels(3).Text = sta.Panels(3).Text & "/ʣ��:" & Format(Val(txtʵ��.Tag) - IIf(gbytBilling = 0, curTotal, 0), "0.00")
                Call SetStatuPatiInfor(Val(cmdOK.Tag), Val(cmdCancel.Tag) + IIf(gbytBilling = 0, curTotal, 0), Val(txtʵ��.Tag) - IIf(gbytBilling = 0, curTotal, 0))
            End If
            
            
            '���¼���ͳ����
            Call ReCalcInsure
            Call SetDrawDrugDeptEnabled
            Screen.MousePointer = 0
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Checkִ�п���() As Integer
    Dim i As Long
    For i = 1 To mobjBill.Details.Count
        If mobjBill.Details(i).ִ�в���ID = 0 Or Bill.TextMatrix(i, BillCol.ִ�п���) = "" Then
            If Not (InStr(",5,6,7,", mobjBill.Details(i).�շ����) > 0 And gbln���뷢ҩ) Then
                Checkִ�п��� = i: Exit Function
            End If
        End If
    Next
End Function

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Function Check�������() As Integer
'���ܣ���鵱ǰ���˵ļ��ʷ�����Ŀ�ķ�������Ƿ�һ��
'˵������Ϊ�������������۲���,�����д˼��
'���أ���һ�µķ�����,Ϊ0ʱ����
    Dim i As Integer
    
    If mrsInfo.State = 0 Then Exit Function
    For i = 1 To mobjBill.Details.Count
        If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
            'סԺ���˻�סԺ���۲���,������ֻ�������������Ŀ
            If mobjBill.Details(i).Detail.������� = 1 Then
                MsgBox "�� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """������������,�ò��˲���ʹ��.", vbInformation, gstrSysName
                Check������� = i: Exit Function
            End If
        ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
            '������Ժ����(ҽ������)���������۲���,������ֻ������סԺ����Ŀ
            If mobjBill.Details(i).Detail.������� = 2 Then
                MsgBox "�� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """��������סԺ,�ò��˲���ʹ��.", vbInformation, gstrSysName
                Check������� = i: Exit Function
            End If
        End If
        If mobjBill.Details(i).Detail.������� = 0 Then
            MsgBox "�� " & i & " ����Ŀ""" & mobjBill.Details(i).Detail.���� & """�������ڲ���,�ò��˲���ʹ��.", vbInformation, gstrSysName
            Check������� = i: Exit Function
        End If
    Next
End Function

Private Sub txtPatient_Validate(Cancel As Boolean)
    If IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
        mblnValid = True
        Call txtPatient_KeyPress(13)
        mblnValid = False
    End If
End Sub
Private Function Get��������ID() As Long
    If cbo��������.ListIndex <> -1 Then
        Get��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    Else
        Get��������ID = UserInfo.����ID
    End If
End Function
Private Function Get������Դ() As Integer
'���ܣ���ȡ��ǰ���˵���Դ(��Ϊ���Զ��������۲��˼���)
    If mrsInfo.State = 1 Then
        If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
            Get������Դ = 2
        ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
            Get������Դ = 1 '���ﲡ��(ҽ������)���������۲���
        End If
    Else
        Get������Դ = 2 'ȱʡΪ2
    End If
End Function
Public Function zl��ȡ��ҩ��̬(Optional ByVal lngRow As Long = -1, Optional blnOnly�г�ҩ As Boolean = False) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����Ƿ�¼�����в�ҩ��
    '���:blnOnly�г�ҩ-���ж��Ƿ����г�ҩ(���䷽ʱ�ж���Ч):ԭ�����г�ҩ���䷽���Ѿ�����,�Ͳ���Ҫ���
    '     lngRow-��ǰ��������
    '����:
    '����:¼�����в�ҩ��,�򷵻���ҩ��̬����(0-ɢװ,1-��Ƭ,2-����),���򷵻�-1 ��ʾ��û��¼����ҩ��̬��Ŀ
    '����:���˺�
    '����:2010-02-02 11:44:17
    '����:27816
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    
    zl��ȡ��ҩ��̬ = -1
    '���δָ��ҳ,���õ�ǰҳ
    If mobjBill Is Nothing Then Exit Function
    strTemp = IIf(blnOnly�г�ҩ, ",6,", ",6,7,")
    With mobjBill.Details
        For i = 1 To .Count
            If InStr(1, strTemp, "," & .Item(i).�շ���� & ",") > 0 And .Item(i).�շ�ϸĿID <> 0 And i <> lngRow Then
                zl��ȡ��ҩ��̬ = .Item(i).Detail.��ҩ��̬
                Exit Function
            End If
        Next
    End With
End Function
Public Sub SetStatuPatiInfor(ByVal dblԤ�� As Double, dblFee As Double, dblʣ�� As Double, Optional dblӦ�� As Double = 0)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����״̬����Ϣ
    '���ƣ����˺�
    '���ڣ�2010-06-23 11:28:31
    '˵����30604
    '------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    picStatuPancl.Visible = False
    '78082:���ϴ�,2014/10/10,Ԥ�������ʾ
    strTemp = "Ԥ��:" & Format(Val(dblԤ��), "0.00")
    strTemp = strTemp & "/����:" & Format(dblFee, gstrDec)
    strTemp = strTemp & "/ʣ��:" & Format(dblʣ��, "0.00")
    If dblӦ�� <> 0 Then
        strTemp = strTemp & "/Ӧ�տ�:" & Format(dblӦ��, "0.00")
    End If
    
    sta.Panels(3).Text = strTemp
    Call MoveStatuPatiInfor
    If dblʣ�� <= 0 Then
        lblStatuPati.Caption = strTemp
        lblStatuPati.AutoSize = True
        picStatuPancl.Visible = True
    End If
    Err = 0
End Sub
Private Sub MoveStatuPatiInfor()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��ƶ�״̬���Ĳ���Ƿ����Ϣ
    '��Σ�
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-06-23 13:51:45
    '˵����30604
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    With picStatuPancl
        .Left = sta.Panels(3).Left + 50
        .Width = sta.Panels(3).Width - 10
        .Top = Me.ScaleHeight - .Height - 10
    End With
End Sub
Private Function zlCheckBill���ڷ�ɢװ��ҩ() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥���д��ڷ�ɢװ��ҩ��̬
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-26 10:19:46
    '����:38328
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If mobjBill Is Nothing Then Exit Function
    If mobjBill.Details.Count = 0 Then Exit Function
    With mobjBill
        For i = 1 To mobjBill.Details.Count
            If .Details(i).�շ���� = "7" Then
                If .Details(i).Detail.��ҩ��̬ <> 0 Then    '0-ɢװ;1-��ҩ��Ƭ;2-����
                    zlCheckBill���ڷ�ɢװ��ҩ = True: Exit Function
                End If
            End If
        Next
    End With
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
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Set objCard = IDKind.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    gobjSquare.bln��ȱʡ������ = IDKind.Cards.��ȱʡ������
      
End Sub

Private Function AddStuffItemFromBarCode(ByVal strBarCode As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������������Ŀ
    '���:strBarCode-��������
    '����:
    '����:���ӳɹ�������True,���򷵻�False
    '����:���˺�
    '����:2017-11-22 13:58:00
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng��Ŀid As Long, str��� As String, lng���� As Long
    Dim intInsure As Integer, lng����ID As Long, dblStock As Double
    Dim lng���˿���ID As Long, lngDoUnit As Long, strժҪ As String
    Dim lngRow As Long, blnCancel As Boolean, str��׼��Ŀ As String, int������Դ As Integer
    Dim blnAdd As Boolean, bln��ʿ As Boolean
    
    On Error GoTo errHandle
    If Trim(strBarCode) = "" Then Exit Function
    
    strBarCode = Trim(strBarCode)
    
    str��� = "'4'"
    Call GetOperatorInfo(mrs������, mobjBill.������, bln��ʿ)
    If bln��ʿ = False Then
        If InStr(gstr�շ����, "'4'") = 0 And gstr�շ���� <> "" Then
            MsgBox "��ǰվ�㲻�߱����������Ͻ����շѻ���ʵ�Ȩ�ޣ���ϵͳ����Ա��ϵ,�ڲ��������п����������ϡ�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    int������Դ = 2
    If Not mrsInfo Is Nothing Then
       If mrsInfo.State = 1 Then
            intInsure = Val(Nvl(mrsInfo!����))
            lng����ID = Val(Nvl(mrsInfo!����ID))
            lngDoUnit = Val(Nvl(mrsInfo!����ID))
            If mrsInfo!�������� = 0 Or mrsInfo!�������� = 2 Then
                int������Դ = 2
            ElseIf mrsInfo!�������� = 1 Or mrsInfo!�������� = -1 Then
                int������Դ = 1
            End If
       End If
    End If
    
    If intInsure <> 0 Then
        If zl_Check��׼��Ŀ(gclsInsure, intInsure, lng����ID, False) Then str��׼��Ŀ = Get������׼��Ŀ(lng����ID, "A.ID")
    End If
    
    If zlCheckBill���ڷ�ɢװ��ҩ() = True Then mblnSelect = False: Exit Function
 
    mlng���� = -1
    lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, int������Դ, intInsure, gblnסԺ��λ, str���, strBarCode, txtBarCode.hWnd, str��׼��Ŀ, -1, "", False, True, lng����)
    If lng��Ŀid = 0 Then Exit Function
    mlng���� = lng����
    
    blnAdd = False
    lngRow = mobjBill.Details.Count
    If lngRow >= Bill.Rows - 1 Then
        Bill.MsfObj.Rows = Bill.MsfObj.Rows + 1
        Bill.Row = Bill.Rows - 1
        Call bill_AfterAddRow(Bill.Row)
        Bill.Col = BillCol.��Ŀ
        blnAdd = True
    End If
        
    Bill.Col = BillCol.��Ŀ
    Bill.SetFocus
    Bill.TxtVisible = True: Bill.Text = lng��Ŀid
    
    mblnSelect = True
    Call Bill_KeyDown(13, 0, blnCancel)
    
    Bill.SetFocus
    If blnCancel Then
        Bill.Text = "": Bill.TxtVisible = False: mblnSelect = False
        If blnAdd And Bill.Rows >= 2 Then
            Bill.Rows = Bill.Rows - 1
            Bill.Row = Bill.Rows - 1
        End If
        AddStuffItemFromBarCode = False: Exit Function
    End If
    mblnSelect = False
    AddStuffItemFromBarCode = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ReadDrugAndStuffStock(ByVal lng�ⷿID As Long, ByRef objDetail As Detail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҩƷ�������ϵĿ����Ϣ
    '���:lng�ⷿID-�ⷿID
    '����:objDetail-Detail����
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-10 09:34:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblStock As Double, strҩ��IDs As String
    On Error GoTo errHandle
    If objDetail Is Nothing Then Exit Function
    '��ҩƷ���������ϵģ�ֱ�ӷ���True
    If InStr(",5,6,7,4,", objDetail.���) = 0 Then ReadDrugAndStuffStock = True: Exit Function
    If objDetail.��� = "4" And objDetail.�������� = False Then ReadDrugAndStuffStock = True: Exit Function
   

    If InStr(",5,6,7,", objDetail.���) > 0 Then
        '��ǰ��ҩƷ���
        If Not gbln���뷢ҩ Then
            dblStock = GetStock(objDetail.ID, lng�ⷿID)
            If gblnסԺ��λ Then
                dblStock = dblStock / objDetail.סԺ��װ
            End If
            objDetail.��� = dblStock
            Call ShowStock(objDetail.����, objDetail.���)
        Else
            strҩ��IDs = Decode(objDetail.���, "5", gstr��ҩ��, "6", gstr��ҩ��, "7", gstr��ҩ��)
            If strҩ��IDs <> "" Then
                dblStock = GetMultiStock(objDetail.ID, strҩ��IDs)
                
                If dblStock = 0 And gblnStock Then
                   MsgBox "[" & objDetail.���� & "]�Ŀ��ÿ��Ϊ��!", vbInformation, gstrSysName
                   Exit Function
                End If
                
                If gblnסԺ��λ Then
                    dblStock = dblStock / objDetail.סԺ��װ
                End If
                objDetail.��� = dblStock
                Call ShowStock(objDetail.����, objDetail.���)
            End If
        End If
        ReadDrugAndStuffStock = True
        Exit Function
    End If
    If objDetail.��� = "4" And objDetail.�������� Then
        dblStock = GetStock(objDetail.ID, lng�ⷿID, objDetail.����)
        objDetail.��� = dblStock
        Call ShowStock(objDetail.����, objDetail.���)
    End If
    ReadDrugAndStuffStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

