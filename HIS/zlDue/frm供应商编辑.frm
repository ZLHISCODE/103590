VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm供应商编辑 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "供应商编辑"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12930
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd照片 
      Caption         =   "导入照片(&F)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   240
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton cmd照片 
      Caption         =   "清除照片(&L)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   1680
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   9120
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Index           =   1
      Left            =   -180
      TabIndex        =   36
      Top             =   645
      Width           =   10155
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   8880
      Width           =   13185
   End
   Begin TabDlg.SSTab sstab 
      Height          =   7935
      Left            =   120
      TabIndex        =   32
      Top             =   720
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   13996
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "常规信息(&0)"
      TabPicture(0)   =   "frm供应商编辑.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chk末级"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "pic基本"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Picture1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkCodeLen"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "附加信息(&1)"
      TabPicture(1)   =   "frm供应商编辑.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblEdit(9)"
      Tab(1).Control(1)=   "lblEdit(10)"
      Tab(1).Control(2)=   "lblEdit(8)"
      Tab(1).Control(3)=   "lblEdit(7)"
      Tab(1).Control(4)=   "lblEdit(6)"
      Tab(1).Control(5)=   "lblEdit(0)"
      Tab(1).Control(6)=   "lblEdit(4)"
      Tab(1).Control(7)=   "lblEdit(5)"
      Tab(1).Control(8)=   "Lbl许可证号"
      Tab(1).Control(9)=   "Lbl许可证效期"
      Tab(1).Control(10)=   "lbl执照号"
      Tab(1).Control(11)=   "Lbl执照效期"
      Tab(1).Control(12)=   "lbl(0)"
      Tab(1).Control(13)=   "lbl(1)"
      Tab(1).Control(14)=   "Label2"
      Tab(1).Control(15)=   "Label3"
      Tab(1).Control(16)=   "dtp授权期"
      Tab(1).Control(17)=   "Dtp执照效期"
      Tab(1).Control(18)=   "Dtp许可证效期"
      Tab(1).Control(19)=   "TxtEdit(11)"
      Tab(1).Control(20)=   "TxtEdit(10)"
      Tab(1).Control(21)=   "TxtEdit(3)"
      Tab(1).Control(22)=   "TxtEdit(2)"
      Tab(1).Control(23)=   "TxtEdit(8)"
      Tab(1).Control(24)=   "TxtEdit(7)"
      Tab(1).Control(25)=   "TxtEdit(9)"
      Tab(1).Control(26)=   "TxtEdit(4)"
      Tab(1).Control(27)=   "TxtEdit(5)"
      Tab(1).Control(28)=   "TxtEdit(6)"
      Tab(1).Control(29)=   "TxtEdit(15)"
      Tab(1).ControlCount=   30
      TabCaption(2)   =   "其他辅助信息(&2)"
      TabPicture(2)   =   "frm供应商编辑.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl首营品种"
      Tab(2).Control(1)=   "lbl备注"
      Tab(2).Control(2)=   "txt首营品种"
      Tab(2).Control(3)=   "txt备注"
      Tab(2).Control(4)=   "Picture2"
      Tab(2).Control(5)=   "Picture3"
      Tab(2).Control(6)=   "Picture4"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "许可证照片(&3)"
      TabPicture(3)   =   "frm供应商编辑.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pic照片(0)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "执照号照片(&4)"
      TabPicture(4)   =   "frm供应商编辑.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "pic照片(1)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "授权号照片(&5)"
      TabPicture(5)   =   "frm供应商编辑.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "pic照片(2)"
      Tab(5).ControlCount=   1
      Begin VB.CheckBox chkCodeLen 
         Caption         =   "允许更改编码长度，并按此调整各同级编码(&L)"
         Height          =   285
         Left            =   480
         TabIndex        =   87
         Top             =   3840
         Width           =   4290
      End
      Begin VB.PictureBox pic照片 
         BorderStyle     =   0  'None
         Height          =   7335
         Index           =   2
         Left            =   -74880
         ScaleHeight     =   7335
         ScaleWidth      =   12375
         TabIndex        =   81
         Top             =   480
         Width           =   12375
         Begin VB.Image img照片 
            Height          =   1650
            Index           =   2
            Left            =   600
            Stretch         =   -1  'True
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.PictureBox pic照片 
         BorderStyle     =   0  'None
         Height          =   7335
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   7335
         ScaleWidth      =   12375
         TabIndex        =   80
         Top             =   480
         Width           =   12375
         Begin VB.Image img照片 
            Height          =   1650
            Index           =   1
            Left            =   600
            Stretch         =   -1  'True
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.PictureBox pic照片 
         BorderStyle     =   0  'None
         Height          =   7335
         Index           =   0
         Left            =   -74880
         ScaleHeight     =   7335
         ScaleWidth      =   12375
         TabIndex        =   79
         Top             =   480
         Width           =   12375
         Begin VB.Image img照片 
            Height          =   1650
            Index           =   0
            Left            =   600
            Stretch         =   -1  'True
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -74760
         ScaleHeight     =   615
         ScaleWidth      =   7575
         TabIndex        =   73
         Top             =   1920
         Width           =   7575
         Begin VB.TextBox TxtEdit 
            Height          =   300
            Index           =   14
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   75
            Tag             =   "药监局备案信息中的证号"
            Top             =   240
            Width           =   2310
         End
         Begin MSComCtl2.DTPicker dtp药监局备案 
            Height          =   300
            Left            =   4680
            TabIndex        =   74
            Top             =   240
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy年MM月dd日"
            DateIsNull      =   -1  'True
            Format          =   139198467
            CurrentDate     =   37994
         End
         Begin VB.Label Label7 
            Caption         =   "药监局备案信息"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   0
            Width           =   3015
         End
         Begin VB.Label lbl 
            Caption         =   "证  号(&V)"
            Height          =   225
            Index           =   6
            Left            =   120
            TabIndex        =   77
            Top             =   300
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "日期(&R)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   4050
            TabIndex        =   76
            Top             =   300
            Width           =   630
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -74760
         ScaleHeight     =   615
         ScaleWidth      =   7575
         TabIndex        =   67
         Top             =   1200
         Width           =   7575
         Begin VB.TextBox TxtEdit 
            Height          =   300
            Index           =   13
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   68
            Tag             =   "质量认证信息中的证号"
            Top             =   240
            Width           =   2310
         End
         Begin MSComCtl2.DTPicker dtp质量认证 
            Height          =   300
            Left            =   4680
            TabIndex        =   69
            Top             =   240
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy年MM月dd日"
            DateIsNull      =   -1  'True
            Format          =   139198467
            CurrentDate     =   37994
         End
         Begin VB.Label Label6 
            Caption         =   "质量认证信息"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "日期(&L)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   4050
            TabIndex        =   71
            Top             =   300
            Width           =   630
         End
         Begin VB.Label lbl 
            Caption         =   "证  号(&Z)"
            Height          =   225
            Index           =   5
            Left            =   120
            TabIndex        =   70
            Top             =   285
            Width           =   915
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -74760
         ScaleHeight     =   615
         ScaleWidth      =   7575
         TabIndex        =   61
         Top             =   480
         Width           =   7575
         Begin VB.TextBox TxtEdit 
            Height          =   300
            Index           =   12
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   62
            Tag             =   "委托书姓名"
            Top             =   240
            Width           =   2310
         End
         Begin MSComCtl2.DTPicker Dtp委托书日期 
            Height          =   300
            Left            =   4680
            TabIndex        =   63
            Top             =   240
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy年MM月dd日"
            DateIsNull      =   -1  'True
            Format          =   139198467
            CurrentDate     =   37994
         End
         Begin VB.Label Label5 
            Caption         =   "销售人员委托书信息"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label lbl 
            Caption         =   "姓  名(&N)"
            Height          =   225
            Index           =   3
            Left            =   120
            TabIndex        =   65
            Top             =   315
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "日期(&D)"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   4050
            TabIndex        =   64
            Top             =   300
            Width           =   630
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   240
         ScaleHeight     =   1335
         ScaleWidth      =   7455
         TabIndex        =   54
         Top             =   2520
         Width           =   7455
         Begin VB.CheckBox chkType 
            Caption         =   "卫生材料(&W)"
            Height          =   315
            Index           =   4
            Left            =   3120
            TabIndex        =   55
            Top             =   360
            Width           =   1410
         End
         Begin VB.CheckBox chkType 
            Caption         =   "药品(&Y)"
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   59
            Top             =   0
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "物资(&M)"
            Height          =   315
            Index           =   1
            Left            =   2040
            TabIndex        =   58
            Top             =   360
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "设备(&J)"
            Height          =   315
            Index           =   2
            Left            =   915
            TabIndex        =   57
            Top             =   360
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "其它(&E)"
            Height          =   315
            Index           =   3
            Left            =   2040
            TabIndex        =   56
            Top             =   0
            Width           =   1125
         End
         Begin VB.Label Label4 
            Caption         =   "类型"
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox pic基本 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   240
         ScaleHeight     =   1935
         ScaleWidth      =   7455
         TabIndex        =   41
         Top             =   480
         Width           =   7455
         Begin VB.TextBox TxtEdit 
            Height          =   300
            Index           =   1
            Left            =   915
            MaxLength       =   10
            TabIndex        =   48
            Tag             =   "简码"
            Top             =   1110
            Width           =   1905
         End
         Begin VB.TextBox TxtEdit 
            Height          =   300
            Index           =   0
            Left            =   915
            MaxLength       =   80
            TabIndex        =   46
            Tag             =   "名称"
            Top             =   750
            Width           =   5655
         End
         Begin VB.TextBox txtParent 
            Height          =   300
            Left            =   915
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   45
            Top             =   0
            Width           =   5385
         End
         Begin VB.CommandButton cmd上级 
            Caption         =   "&P"
            Height          =   300
            Left            =   6300
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   0
            Width           =   285
         End
         Begin VB.ComboBox cmbStationNo 
            Height          =   300
            Left            =   915
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1500
            Width           =   1905
         End
         Begin VB.TextBox txtCode 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   42
            Tag             =   "编码"
            Text            =   "111111"
            Top             =   420
            Width           =   1755
         End
         Begin VB.TextBox txtUpCode 
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   915
            MaxLength       =   10
            TabIndex        =   47
            TabStop         =   0   'False
            Tag             =   "编码"
            Text            =   "11"
            Top             =   375
            Width           =   1905
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "上级(&U)"
            Height          =   180
            Index           =   11
            Left            =   240
            TabIndex        =   53
            Top             =   75
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "编码(&D)"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   52
            Top             =   435
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "名称(&N)"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   51
            Top             =   795
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "简码(&S)"
            Height          =   180
            Index           =   3
            Left            =   240
            TabIndex        =   50
            Top             =   1155
            Width           =   630
         End
         Begin VB.Label lblStationNo 
            AutoSize        =   -1  'True
            Caption         =   "院区(&B)"
            Height          =   180
            Left            =   240
            TabIndex        =   49
            Top             =   1545
            Width           =   630
         End
      End
      Begin VB.TextBox txt备注 
         Height          =   2265
         Left            =   -73560
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   3240
         Width           =   5805
      End
      Begin VB.TextBox txt首营品种 
         Height          =   585
         Left            =   -73560
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   2520
         Width           =   5805
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   15
         Left            =   -73740
         MaxLength       =   16
         TabIndex        =   17
         Tag             =   "授权号"
         Top             =   2070
         Width           =   2205
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   6
         Left            =   -73740
         MaxLength       =   16
         TabIndex        =   13
         Tag             =   "执照号"
         Top             =   1680
         Width           =   2205
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   5
         Left            =   -73740
         MaxLength       =   16
         TabIndex        =   9
         Tag             =   "许可证号"
         Top             =   1290
         Width           =   2205
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   4
         Left            =   -69900
         MaxLength       =   16
         TabIndex        =   7
         Tag             =   "电话"
         Top             =   900
         Width           =   2625
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   9
         Left            =   -73740
         MaxLength       =   50
         TabIndex        =   29
         Tag             =   "地址"
         Top             =   3255
         Width           =   6450
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   7
         Left            =   -69900
         MaxLength       =   20
         TabIndex        =   3
         Tag             =   "税务登记号"
         Top             =   510
         Width           =   2640
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   8
         Left            =   -73740
         MaxLength       =   50
         TabIndex        =   27
         Tag             =   "开户银行"
         Top             =   2880
         Width           =   6450
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   2
         Left            =   -73740
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "帐号"
         Top             =   510
         Width           =   2205
      End
      Begin VB.TextBox TxtEdit 
         Height          =   300
         Index           =   3
         Left            =   -73740
         MaxLength       =   20
         TabIndex        =   5
         Tag             =   "联系人"
         Top             =   900
         Width           =   2205
      End
      Begin VB.TextBox TxtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   10
         Left            =   -73740
         MaxLength       =   6
         TabIndex        =   21
         Tag             =   "信用期"
         Top             =   2460
         Width           =   2055
      End
      Begin VB.TextBox TxtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   11
         Left            =   -69900
         MaxLength       =   8
         TabIndex        =   24
         Tag             =   "信用额"
         Top             =   2460
         Width           =   2430
      End
      Begin VB.CheckBox chk末级 
         Caption         =   "末级(&M)"
         Height          =   180
         Left            =   510
         TabIndex        =   33
         Top             =   8040
         Visible         =   0   'False
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker Dtp许可证效期 
         Height          =   300
         Left            =   -69900
         TabIndex        =   11
         Top             =   1290
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy年MM月dd日"
         DateIsNull      =   -1  'True
         Format          =   141033475
         CurrentDate     =   37994
      End
      Begin MSComCtl2.DTPicker Dtp执照效期 
         Height          =   300
         Left            =   -69900
         TabIndex        =   15
         Top             =   1680
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy年MM月dd日"
         DateIsNull      =   -1  'True
         Format          =   141099011
         CurrentDate     =   37994
      End
      Begin MSComCtl2.DTPicker dtp授权期 
         Height          =   300
         Left            =   -69900
         TabIndex        =   19
         Top             =   2070
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy年MM月dd日"
         DateIsNull      =   -1  'True
         Format          =   141033475
         CurrentDate     =   37994
      End
      Begin VB.Label lbl备注 
         AutoSize        =   -1  'True
         Caption         =   "备  注(&B)"
         Height          =   180
         Left            =   -74640
         TabIndex        =   40
         Top             =   3240
         Width           =   810
      End
      Begin VB.Label lbl首营品种 
         AutoSize        =   -1  'True
         Caption         =   "首营品种(&S)"
         Height          =   180
         Left            =   -74640
         TabIndex        =   39
         Top             =   2640
         Width           =   990
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "授权号(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74550
         TabIndex        =   16
         Top             =   2130
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "授权期(&Q)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -70740
         TabIndex        =   18
         Top             =   2130
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "月"
         Height          =   180
         Index           =   1
         Left            =   -71670
         TabIndex        =   22
         Top             =   2520
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "元"
         Height          =   180
         Index           =   0
         Left            =   -67440
         TabIndex        =   25
         Top             =   2520
         Width           =   180
      End
      Begin VB.Label Lbl执照效期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "执照效期(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -70920
         TabIndex        =   14
         Top             =   1740
         Width           =   990
      End
      Begin VB.Label lbl执照号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "执照号(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74550
         TabIndex        =   12
         Top             =   1740
         Width           =   810
      End
      Begin VB.Label Lbl许可证效期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "许可证效期(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -71100
         TabIndex        =   10
         Top             =   1350
         Width           =   1170
      End
      Begin VB.Label Lbl许可证号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "许可证号(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74730
         TabIndex        =   8
         Top             =   1350
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "电话(&T)"
         Height          =   180
         Index           =   5
         Left            =   -70560
         TabIndex        =   6
         Top             =   960
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "地址(&A)"
         Height          =   180
         Index           =   4
         Left            =   -74370
         TabIndex        =   28
         Top             =   3315
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "税务登记号(&K)"
         Height          =   180
         Index           =   0
         Left            =   -71100
         TabIndex        =   2
         Top             =   570
         Width           =   1170
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "开户银行(&G)"
         Height          =   180
         Index           =   6
         Left            =   -74730
         TabIndex        =   26
         Top             =   2940
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "帐  号(&Z)"
         Height          =   180
         Index           =   7
         Left            =   -74550
         TabIndex        =   0
         Top             =   570
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "联系人(&L)"
         Height          =   180
         Index           =   8
         Left            =   -74550
         TabIndex        =   4
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "信用期(&Y)"
         Height          =   180
         Index           =   10
         Left            =   -74550
         TabIndex        =   20
         Top             =   2520
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "信用额(&E)"
         Height          =   180
         Index           =   9
         Left            =   -70740
         TabIndex        =   23
         Top             =   2520
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   10080
      TabIndex        =   30
      Top             =   9120
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   11400
      TabIndex        =   31
      Top             =   9120
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog cdl照片 
      Left            =   7920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl图片说明 
      Height          =   210
      Index           =   2
      Left            =   3120
      TabIndex        =   86
      Top             =   9180
      Visible         =   0   'False
      Width           =   2600
   End
   Begin VB.Label lbl图片说明 
      Height          =   210
      Index           =   1
      Left            =   3120
      TabIndex        =   85
      Top             =   9187
      Visible         =   0   'False
      Width           =   2600
   End
   Begin VB.Label lbl图片说明 
      Height          =   210
      Index           =   0
      Left            =   3120
      TabIndex        =   84
      Top             =   9187
      Visible         =   0   'False
      Width           =   2600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "对药品(物资、设备等)供应商建立档案或修改档案.同时可加长或减少已有编码的长度。"
      Height          =   180
      Left            =   600
      TabIndex        =   35
      Top             =   345
      Width           =   6930
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frm供应商编辑.frx":00A8
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frm供应商编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrID As String         '当前编辑的单位ID
Dim mlng上级id As Long       '上级单位ID
Dim mintSuccess As Integer
Dim mintEditType As gEditType    '编辑类型
Dim mblnChange As Boolean
Dim mstrPrivs As String         '权限串
Const mintMaxLen = 8        '编码长度
Dim mblnFist As Boolean

Private Enum picType
    许可证照片 = 0
    执照照片 = 1
    授权号照片 = 2
End Enum

Private Type picCon
    mblnExistPic(0 To 2) As Boolean     '当前是否有图片信息
    mblnIsModify(0 To 2) As Boolean     '当照片发生更改时才为True
End Type
Private myPicCon As picCon

Private Sub InitDefaultLen()
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置编辑的默认长度
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-10-23 14:31:25
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, i As Long, j As Long
    Dim strSQL As String
    On Error GoTo errHandle
    strSQL = "Select 税务登记号,许可证号,执照号,授权号 From 供应商 where id=0"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    For i = 0 To rsTemp.Fields.Count - 1
        For j = 0 To TxtEdit.UBound
            If rsTemp.Fields(i).Name = TxtEdit(j).Tag Then
                TxtEdit(j).MaxLength = rsTemp.Fields(i).DefinedSize
                Exit For
            End If
        Next
    Next
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub chkCodeLen_Click()
    If chkCodeLen.Visible = False Then Exit Sub
    If Me.chkCodeLen.Value = 1 Then
        Me.txtCode.MaxLength = mintMaxLen - Len(Me.txtUpCode.Text)
    Else
        Me.txtCode.MaxLength = Me.txtCode.Tag
        Me.txtCode.Text = Mid(Me.txtCode.Text, 1, Me.txtCode.MaxLength)
    End If
    If sstab.Tab = 0 Then
        If Me.txtCode.Enabled Then txtCode.SetFocus
    End If
End Sub

Private Sub chkCodeLen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
         zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkType_Click(Index As Integer)
    mblnChange = True
    setCtlEn
End Sub

Private Sub chkType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
    Case 3
        '需打开另外一页信息,需输入
        sstab.Tab = 1
                
        If TxtEdit(2).Enabled Then
            TxtEdit(2).SetFocus
        Else
            zlCommFun.PressKey vbKeyTab
        End If
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub cmbStationNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub CmdCancel_Click()
    Dim blnYes As Boolean
    If mblnChange = False Then
        Unload Me
        Exit Sub
    End If
    ShowMsgbox "你已经更改了档案信息,你这样退出的话," & vbCrLf & "所更改的数据将不能保存,真的要退出吗?", True, blnYes
    If blnYes = True Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub cmdOK_Click()
    Dim intIndex As Integer
    
    If IsValid() = False Then Exit Sub
    If Save单位() = False Then Exit Sub
    
    mintSuccess = mintSuccess + 1
    
    If mstrID <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    
    cmdOK.Enabled = False
    mstrID = ""
    For intIndex = 0 To 15
        TxtEdit(intIndex).Text = ""
    Next
    zlChangeCode "供应商", mlng上级id, txtUpCode, txtCode, chkCodeLen, Me.Caption
     mblnChange = False
    sstab.Tab = 0
    If TxtEdit(0).Enabled And TxtEdit(0).Visible Then
        TxtEdit(0).SetFocus
    End If
End Sub

Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:验证数据的合法性
    '--入参数:
    '--出参数:
    '--返  回:验证合法,返回True,否则=false
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    
    Dim strTemp As String
    
    strTemp = Trim(txtCode.Text)
     
    If strTemp = "" Then
        ShowMsgbox "编码必须输入!"
        If txtCode.Enabled Then txtCode.SetFocus
        Exit Function
    End If
    If InStr(1, strTemp, "'") <> 0 Then
        ShowMsgbox "编码不能输入单引号!"
        If txtCode.Enabled Then txtCode.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(strTemp) Then
        ShowMsgbox "编码必需由数字组成,请重输!"
        If txtCode.Enabled Then txtCode.SetFocus
        Exit Function
    End If
    If Len(txtUpCode.Text & strTemp) > 8 Then
        ShowMsgbox "编码长度不能超过8位,请重输!"
        If txtCode.Enabled Then txtCode.SetFocus
        Exit Function
    End If
    If LenB(txt首营品种.Text) > 200 Then
        ShowMsgbox "首营品种不能超过200位字符或100个汉字,请重输!"
        sstab.Tab = 2
        If txt首营品种.Enabled Then txt首营品种.SetFocus
        Exit Function
    End If
    If LenB(txt备注.Text) > 200 Then
        ShowMsgbox "备注不能超过200位字符或100个汉字,请重输!"
        sstab.Tab = 2
        If txt备注.Enabled Then txt备注.SetFocus
        Exit Function
    End If
    For intIndex = 0 To 15
        strTemp = Trim(TxtEdit(intIndex).Text)
        If intIndex = 0 Then
            If strTemp = "" Then
                ShowMsgbox TxtEdit(intIndex).Tag & "必需输入!"
                If TxtEdit(intIndex).Enabled Then TxtEdit(intIndex).SetFocus
                Exit Function
            End If
        End If
        
        If strTemp <> "" Then
            If LenB(StrConv(strTemp, vbFromUnicode)) > TxtEdit(intIndex).MaxLength Then
                ShowMsgbox TxtEdit(intIndex).Tag & "超长,最多能输入" & TxtEdit(intIndex).MaxLength / 2 & "个汉字或" & TxtEdit(intIndex).MaxLength & "个字符!"
                If TxtEdit(intIndex).Enabled Then TxtEdit(intIndex).SetFocus
                Exit Function
            End If
            If InStr(1, strTemp, "'") <> 0 Then
                ShowMsgbox TxtEdit(intIndex).Tag & "不能输入单引号!"
                If TxtEdit(intIndex).Enabled Then TxtEdit(intIndex).SetFocus
                Exit Function
            End If
            
            Select Case TxtEdit(intIndex).Tag
            Case "信用期", "信用额"
                If Not IsNumeric(strTemp) Then
                    ShowMsgbox TxtEdit(intIndex).Tag & "不是数据型,请重输!"
                    If TxtEdit(intIndex).Enabled Then TxtEdit(intIndex).SetFocus
                    Exit Function
                End If
                If TxtEdit(intIndex).Tag = "信用额" Then
                    If Val(strTemp) > 99999999 Then
                        ShowMsgbox TxtEdit(intIndex).Tag & "大于了99999999,请重输!"
                        If TxtEdit(intIndex).Enabled Then TxtEdit(intIndex).SetFocus
                        Exit Function
                    End If
                Else
                    If Val(strTemp) > 999999 Then
                        ShowMsgbox TxtEdit(intIndex).Tag & "大于了999999,请重输!"
                        If TxtEdit(intIndex).Enabled Then TxtEdit(intIndex).SetFocus
                        Exit Function
                    End If
                End If
            End Select
        End If
    Next
    Dim blnTrue As Boolean
    Dim i As Byte
    For i = 0 To 4
        If chkType(i).Value = 1 Then
            blnTrue = True
        End If
    Next
    If blnTrue = False And chk末级.Value = 1 Then
        ShowMsgbox "没选择是属于哪种供应商,请选择!"
        Exit Function
    End If
    IsValid = True
End Function

Private Function Save单位() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:保存数据
    '--入参数:
    '--出参数:
    '--返  回:保存成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim lngPriType As Long
    Dim lngID As Long
    Dim strTmp As String
    Dim intIndex As Integer
    Dim blnTran As Boolean
    Dim i As Integer
    Dim rsDepend As New ADODB.Recordset
    
    On Error GoTo errHandle
    strTmp = ""
    For intIndex = 0 To 4
        strTmp = strTmp & IIf(chkType(intIndex).Value = 1, 1, 0)
    Next
    
    If mintEditType = g新增 Then
        gstrSQL = "Select 名称 From 供应商 "
    ElseIf mintEditType = g修改 Then
        gstrSQL = "Select 名称 From 供应商 Where ID <> [1]  "
    End If
    
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "检查名称是否重复", Val(mstrID))
    Do While Not rsDepend.EOF
        If TxtEdit(0) = rsDepend!名称 Then
            MsgBox "名称重复请重新输入！", vbInformation, gstrSysName
            Exit Function
        End If
        rsDepend.MoveNext
    Loop
        
    If mstrID = "" Then
        lngID = zldatabase.GetNextId("供应商")
        gstrSQL = "zl_供应商_insert ( "
    Else
        lngID = Val(mstrID)
        gstrSQL = "zl_供应商_update ( "
    End If
    
    '过程参数如下:
    '   ID_IN,上级ID_IN,编码_IN,名称_IN,简码_IN,地址_IN,电话_IN,开户银行_IN,帐号_IN,联系人_IN,
    '   税务登记号_IN,许可证号_IN,许可证效期_IN,执照号_IN,执照效期_IN,授权号_IN,授权期_IN,供应商类型_IN,信用期_IN,
    '   信用额_IN,销售委托人_IN ,销售委托日期_IN,质量认证号_IN,质量认证日期_IN, 药监局备案号_IN,药监局备案日期_IN
    '     站点_In,末级_IN , 改变编码长度,首营品种_In,备注_In
    
    gstrSQL = gstrSQL & "" & _
            lngID & "," & _
            IIf(mlng上级id = 0, "Null", mlng上级id) & ",'" & _
            txtUpCode.Text & txtCode.Text & "','" & _
            Trim(TxtEdit(0).Text) & "','" & _
            Trim(TxtEdit(1).Text) & "'," & _
            IIf(Trim(TxtEdit(9).Text) = "", "NULL", "'" & Trim(TxtEdit(9).Text) & "'") & "," & _
            IIf(Trim(TxtEdit(4).Text) = "", "NULL", "'" & Trim(TxtEdit(4).Text) & "'") & "," & _
            IIf(Trim(TxtEdit(8).Text) = "", "NULL", "'" & Trim(TxtEdit(8).Text) & "'") & "," & _
            IIf(Trim(TxtEdit(2).Text) = "", "NULL", "'" & Trim(TxtEdit(2).Text) & "'") & "," & _
            IIf(Trim(TxtEdit(3).Text) = "", "NULL", "'" & Trim(TxtEdit(3).Text) & "'") & "," & _
            IIf(Trim(TxtEdit(7).Text) = "", "NULL", "'" & Trim(TxtEdit(7).Text) & "'") & "," & _
            IIf(Trim(TxtEdit(5).Text) = "", "NULL", "'" & Trim(TxtEdit(5).Text) & "'") & "," & _
            IIf(Dtp许可证效期.Value = "" Or IsNull(Dtp许可证效期.Value), "NULL", "to_Date('" & Format(Dtp许可证效期.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
            IIf(Trim(TxtEdit(6).Text) = "", "NULL", "'" & Trim(TxtEdit(6).Text) & "'") & "," & _
            IIf(Dtp执照效期.Value = "" Or IsNull(Dtp执照效期.Value), "NULL", "to_Date('" & Format(Dtp执照效期.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
            IIf(Trim(TxtEdit(15).Text) = "", "NULL", "'" & Trim(TxtEdit(15).Text) & "'") & "," & _
            IIf(dtp授权期.Value = "" Or IsNull(dtp授权期.Value), "NULL", "to_Date('" & Format(dtp授权期.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
            "'" & strTmp & "'," & _
            IIf(Trim(TxtEdit(10).Text) = "", "NULL", Val(TxtEdit(10).Text)) & "," & _
            IIf(Trim(TxtEdit(11).Text) = "", "NULL", Val(TxtEdit(11).Text)) & ","
        gstrSQL = gstrSQL & _
            IIf(Trim(TxtEdit(12).Text) = "", "NULL", "'" & Trim(TxtEdit(12).Text) & "'") & "," & _
            IIf(Dtp委托书日期.Value = "" Or IsNull(Dtp委托书日期.Value), "NULL", "to_Date('" & Format(Dtp委托书日期.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
            IIf(Trim(TxtEdit(13).Text) = "", "NULL", "'" & Trim(TxtEdit(13).Text) & "'") & "," & _
            IIf(dtp质量认证.Value = "" Or IsNull(dtp质量认证.Value), "NULL", "to_Date('" & Format(dtp质量认证.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
            IIf(Trim(TxtEdit(14).Text) = "", "NULL", "'" & Trim(TxtEdit(14).Text) & "'") & "," & _
            IIf(dtp药监局备案.Value = "" Or IsNull(dtp药监局备案.Value), "NULL", "to_Date('" & Format(dtp药监局备案.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
            IIf(cmbStationNo.Visible = True And Trim(cmbStationNo.Text) <> "", "'" & Me.cmbStationNo.ItemData(Me.cmbStationNo.ListIndex) & "'", "NULL") & "," & chk末级.Value & "," & _
            IIf(Me.chkCodeLen.Value = 1, 1, 0) & "," & _
            IIf(Trim(txt首营品种.Text) = "", "NULL", "'" & txt首营品种.Text & "'") & "," & _
            IIf(Trim(txt备注.Text) = "", "NULL", "'" & txt备注.Text & "'") & _
            ")"
    
    gcnOracle.BeginTrans: blnTran = True
    zldatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    '处理照片
    For i = 0 To 2
        If myPicCon.mblnIsModify(i) = True Then
            '只有发生了更改才需要处理
            Call zldatabase.ExecuteProcedure("Zl_供应商照片_Delete(" & lngID & "," & i & ")", Me.Caption)
            
            If myPicCon.mblnExistPic(i) = True And img照片(i).Tag <> "" Then
                '保存
                If sys.Savelob(100, 23, lngID & "," & i, img照片(i).Tag) = False Then
                    gcnOracle.RollbackTrans
                    MsgBox "照片保存失败。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Next
    
    gcnOracle.CommitTrans: blnTran = False
    
    Save单位 = True
    Exit Function
errHandle:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function 编辑单位(ByVal FrmMain As Object, ByVal lng上级id As Long, _
    intEditType As gEditType, Optional strID As String = "", Optional ByVal bln末级 As Boolean = False, _
    Optional strPrivs As String = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:编辑供应商档案
    '--入参数:frmMain-调用的主窗体
    '--       lng上级id-上级id
    '--       intEditType -编辑类型
    '--       strID-编辑档案的当前ID
    '--       bln末级-是否是未级项目
    '--出参数:
    '--返  回:编辑成功,返回ture,否则false
    '-----------------------------------------------------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    Dim intTemp As Byte, i As Integer
    Dim strTemp As String
    Dim strTempFile As String
   

    mintSuccess = 0
    
    mstrID = strID
    mlng上级id = lng上级id
    mintEditType = intEditType
    mstrPrivs = strPrivs
    On Error GoTo errHandle
    '初始化院区信息
    gstrSQL = "Select 编号, 名称 From Zltools.Zlnodelist "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "获取站点")
    With cmbStationNo
        .Clear
        .AddItem ""
        Do While Not rsTemp.EOF
            .AddItem rsTemp!编号 & "-" & rsTemp!名称
            .ItemData(.NewIndex) = rsTemp!编号
            rsTemp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    If mlng上级id <> 0 Then
        '求出上级编码及名称
        'by lesfeng 2009-12-2 性能优化
        gstrSQL = "Select 编码,名称 From 供应商 where id=[1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng上级id)
        If rsTemp.EOF Then
            ShowMsgbox "上级分类已被他人删除,不能再增加该分类的下级项目!"
            Exit Function
        End If
        txtParent.Text = "[" & Nvl(rsTemp!编码, "..") & "]" & Nvl(rsTemp!名称, "..")
        txtUpCode.Text = Nvl(rsTemp!编码)
        mlng上级id = lng上级id
    Else
        txtParent.Text = "无"
        txtUpCode.Text = ""
    End If
    If mintEditType <> g新增 Then
        '需确定本级需操作的项目
        'by lesfeng 2009-12-2 性能优化
        gstrSQL = "Select ID,上级ID,编码,名称,简码,末级,许可证号,许可证效期,执照号,执照效期,税务登记号,地址,电话,开户银行," & _
                  "       帐号,联系人,建档时间,撤档时间,类型,信用期,信用额,销售委托人,销售委托日期,质量认证号,质量认证日期," & _
                  "       药监局备案号,药监局备案日期,授权号,授权期,站点,首营品种,备注 " & _
                  "  From 供应商 where id=[1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "获取被编辑的供应商!", Val(strID))
        
        If mintEditType = g查看 Then
        Else
            If SetEditPro(Nvl(rsTemp!类型)) = False Then Exit Function
        End If
        With rsTemp
            txtCode.Text = Mid(Nvl(!编码), Len(txtUpCode.Text) + 1)
            txtCode.MaxLength = Len(txtCode.Text)
            txtCode.Tag = .Fields("编码").DefinedSize
            Dim intIndex As Long
            For intIndex = 0 To 11
                strTemp = TxtEdit(intIndex).Tag
                Select Case strTemp
                Case "信用额"
                        TxtEdit(intIndex).Text = Format(Nvl(.Fields(strTemp), 0), "####0.00;####0.00; ;")
                Case Else
                    TxtEdit(intIndex).Text = Nvl(.Fields(strTemp))
                End Select
            Next
            If IsNull(!许可证效期) Then
                Dtp许可证效期.Value = ""
            Else
                Dtp许可证效期.Value = Format(!许可证效期, "yyyy-mm-dd")
            End If
            If IsNull(!执照效期) Then
                Dtp执照效期.Value = ""
            Else
                Dtp执照效期.Value = Format(!执照效期, "yyyy-mm-dd")
            End If
            
            TxtEdit(15).Text = Nvl(!授权号)
            If IsNull(!授权期) Then
                dtp授权期.Value = ""
            Else
                dtp授权期.Value = Format(!授权期, "yyyy-mm-dd")
            End If
                        
            If IsNull(!销售委托日期) Then
                Dtp委托书日期.Value = ""
            Else
                Dtp委托书日期.Value = Format(!销售委托日期, "yyyy-mm-dd")
            End If
            
            If IsNull(!质量认证日期) Then
                dtp质量认证.Value = ""
            Else
                dtp质量认证.Value = Format(!质量认证日期, "yyyy-mm-dd")
            End If
            If IsNull(!药监局备案日期) Then
                dtp药监局备案.Value = ""
            Else
                dtp药监局备案.Value = Format(!药监局备案日期, "yyyy-mm-dd")
            End If
                        
            TxtEdit(12).Text = Nvl(!销售委托人)
            TxtEdit(13).Text = Nvl(!质量认证号)
            TxtEdit(14).Text = Nvl(!药监局备案号)
            
            txt首营品种.Text = Nvl(!首营品种)
            txt备注.Text = Nvl(!备注)
            
            '加载站点信息
            With cmbStationNo
                For i = 0 To .ListCount - 1
                    If Mid(.List(i), 1, 1) = Nvl(rsTemp!站点) Then
                        .ListIndex = i
                        Exit For
                    End If
                Next
            End With
            
            If !末级 = 1 Then
                chk末级.Value = 1
            Else
                chk末级.Value = 0
            End If
            strTemp = Nvl(!类型)
            
            '获取类型
            If Len(strTemp) >= 4 Then
                For intTemp = 0 To 4
                    If intTemp > Len(strTemp) - 1 Then
                        chkType(intTemp).Value = 0
                    Else
                        chkType(intTemp).Value = Mid(strTemp, intTemp + 1, 1)
                    End If
                Next
            End If
        End With
    Else
        '新增
        zlChangeCode "供应商", mlng上级id, txtUpCode, txtCode, chkCodeLen, Me.Caption
        If bln末级 Then
            chk末级.Value = 1
        Else
            chk末级.Value = 0
        End If
        For intTemp = 0 To 4
            chkType(intTemp).Value = 0
        Next
    End If
    
    If chk末级.Value <> 1 Then
        Set pic基本.Container = Me
        pic基本.Top = sstab.Top
        chkCodeLen.Top = pic基本.Top + pic基本.Height + 100
        fra(0).Top = chkCodeLen.Top + chkCodeLen.Height + 100
        cmdCancel.Top = fra(0).Top + fra(0).Height + 100
        cmdOK.Top = cmdCancel.Top
        sstab.Visible = False
        frm供应商编辑.Height = cmdCancel.Top + cmdCancel.Height + 600
    Else
        sstab.Visible = True
    End If
    mblnChange = False
    Me.chkCodeLen.Visible = InStr(1, mstrPrivs, "改变编码长度") <> 0
    If chk末级.Value <> 1 Then
        Me.Caption = "分类编辑"
        Label1.Caption = "对供应商分类进行设置.同时可加长或减少已有编码的长度。"
    End If
    
    '处理图片
    For i = 0 To 2
        strTempFile = sys.Readlob(100, 23, strID & "," & i)
        img照片(i).Picture = LoadPicture(strTempFile)
        myPicCon.mblnIsModify(i) = False
        myPicCon.mblnExistPic(i) = (strTempFile <> "")
        lbl图片说明(i) = GetPictureInfo(img照片(i).Picture)
        '删除该临时文件
        If lbl图片说明(i) <> "无照片" Then
            Kill strTempFile
        End If
    Next
    
    frm供应商编辑.Show 1, FrmMain
    编辑单位 = mintSuccess > 0
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub setCtlEn()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:设置控件的Enable属性
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    If mintEditType = g查看 Then
        txtCode.Enabled = False
        txtUpCode.Enabled = False
        txtParent.Enabled = False
        cmd上级.Enabled = False
        cmdOK.Visible = False
        cmbStationNo.Enabled = False
        For intIndex = 0 To 4
            chkType(intIndex).Enabled = False
        Next
        
        For intIndex = 0 To TxtEdit.UBound
            TxtEdit(intIndex).Enabled = False
        Next
        Dtp许可证效期.Enabled = False
        Dtp执照效期.Enabled = False
        Dtp委托书日期.Enabled = False
        dtp质量认证.Enabled = False
        dtp药监局备案.Enabled = False
'        cmbStationNo.Enabled = False
        chkCodeLen.Enabled = False
        dtp授权期.Enabled = False
        txt首营品种.Enabled = False
        txt备注.Enabled = False
        
    End If
    cmdOK.Enabled = mblnChange And Trim(TxtEdit(0).Text) <> "" And Trim(txtCode.Text) <> ""
End Sub

Private Sub cmd上级_Click()
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim strID As String
    Dim str编码 As String
    Dim int编码  As Integer
    
    gstrSQL = "select ID,上级ID,名称,编码 from 供应商  " & _
        "where 末级 <> 1 start with 上级ID is null connect by prior ID =上级ID"
    strID = IIf(mlng上级id = 0, "", mlng上级id)
    
    str名称 = TxtEdit(0).Text
    str编码 = txtUpCode.Text
    blnRe = frm树型选择.ShowTree(gstrSQL, strID, str名称, str编码, mstrID, "供应商", "所有供应商")
    '成功返回
    If blnRe Then
        '新的本级的宽度
        txtParent.Text = str名称
        mlng上级id = Val(strID)
        '设置编码
        zlChangeCode "供应商", mlng上级id, txtUpCode, txtCode, chkCodeLen, Me.Caption
        setCtlEn
    End If
End Sub

Private Sub cmd照片_Click(Index As Integer)
    Dim intPicIndex As Integer
    
    intPicIndex = sstab.Tab - 3
    
    Select Case Index
        Case 0 '文件
            With cdl照片
                .CancelError = True
                .Filter = "图片文件(*.bmp,*.gif,*.jpg)|*.bmp;*.gif;*.jpg"
                
                On Error Resume Next
                .ShowOpen
                If Err <> 0 Then
                    '没选中文件
                    Err.Clear
                Else
                    img照片(intPicIndex).Picture = LoadPicture(.FileName)
'                    img照片.Left = pic镜框.ScaleLeft
'                    img照片.Top = pic镜框.ScaleTop
                    
'                    DoEvents
                    If Err <> 0 Then
                        MsgBox "图片文件无效，或文件不存在。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    lbl图片说明(intPicIndex) = GetPictureInfo(img照片(intPicIndex).Picture)
                    img照片(intPicIndex).Tag = .FileName
                    myPicCon.mblnExistPic(intPicIndex) = True
                    myPicCon.mblnIsModify(intPicIndex) = True
                End If
            End With
        Case 1 '清除
            myPicCon.mblnExistPic(intPicIndex) = False
            myPicCon.mblnIsModify(intPicIndex) = True
            Call 显示空图片(intPicIndex)
    End Select
    
    
End Sub

Private Sub 显示空图片(ByVal intPicIndex As Integer)
    '在图片框中显示无图片信息
    If myPicCon.mblnExistPic(intPicIndex) = False Then
        img照片(intPicIndex).Picture = Nothing
        img照片(intPicIndex).Tag = ""
        lbl图片说明(intPicIndex) = "无照片"
    End If
End Sub
Private Sub dtp授权期_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub dtp授权期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Dtp委托书日期_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub Dtp委托书日期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Dtp许可证效期_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub Dtp许可证效期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtp药监局备案_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub dtp药监局备案_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Dtp执照效期_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub Dtp执照效期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtp质量认证_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub dtp质量认证_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Activate()
    If mblnFist = False Then Exit Sub
    mblnFist = False
    '初始站点
'    cmbStationNo.Visible = gSystemPara.bln存在站点 And chk末级.Value = 1
'    lblStationNo.Visible = cmbStationNo.Visible
    
    If Me.TxtEdit(0).Enabled Then Me.TxtEdit(0).SetFocus
    Call 权限控制
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnFist = True
    Call InitDefaultLen
End Sub



Private Sub Form_Resize()
    Dim i As Integer
    
    On Error Resume Next
    
    For i = 0 To 2
        img照片(i).Move pic照片(i).ScaleLeft, pic照片(i).ScaleTop, pic照片(i).ScaleWidth, pic照片(i).ScaleHeight
    Next
End Sub

Private Sub sstab_Click(PreviousTab As Integer)
    If sstab.Tab >= 3 And sstab.Tab <= 5 Then
        cmd照片(0).Enabled = True
        cmd照片(1).Enabled = True
    Else
        cmd照片(0).Enabled = False
        cmd照片(1).Enabled = False
    End If
    
    lbl图片说明(0).Visible = (sstab.Tab = 3)
    lbl图片说明(1).Visible = (sstab.Tab = 4)
    lbl图片说明(2).Visible = (sstab.Tab = 5)
End Sub

Private Sub TxtCode_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtCode, KeyAscii, m数字式
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 0 Then
        TxtEdit(1).Text = zlCommFun.SpellCode(TxtEdit(0).Text)
    End If
    setCtlEn
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Dim blnOpen As Boolean
    
    Select Case TxtEdit(Index).Tag
    Case "信用期", "信用额", "简码"
            blnOpen = False
    Case Else
            blnOpen = True
    End Select
    SetTxtGotFocus TxtEdit(Index), blnOpen
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Select Case TxtEdit(Index).Tag
        Case "名称"
            If cmbStationNo.Visible And cmbStationNo.Enabled Then
                cmbStationNo.SetFocus
            ElseIf chkType(0).Enabled And chkType(0).Visible Then
                chkType(0).SetFocus
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        Case Else
            zlCommFun.PressKey vbKeyTab
        End Select
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case TxtEdit(Index).Tag
    Case "信用期"
            zlControl.TxtCheckKeyPress TxtEdit(Index), KeyAscii, m数字式
    Case "信用额"
            zlControl.TxtCheckKeyPress TxtEdit(Index), KeyAscii, m金额式
    Case "帐号"
        If LenB(StrConv(TxtEdit(Index).Text, vbFromUnicode)) >= 50 And (KeyAscii <> vbKeyBack) Then
            KeyAscii = 0
        End If
    Case Else
            zlControl.TxtCheckKeyPress TxtEdit(Index), KeyAscii, m文本式
    End Select
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If TxtEdit(Index).Tag = "信用额" Then
        TxtEdit(Index).Text = Format(Val(TxtEdit(Index).Text), "####0.00;-####0.00; ;")
    End If
    ImeLanguage False
End Sub

Private Sub txtParent_Change()
    mblnChange = True
    setCtlEn
End Sub

Private Sub txtParent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtUpCode_Change()
    Me.txtCode.Width = txtUpCode.Width - TextWidth(txtUpCode.Text) - 120
    Me.txtCode.Left = txtUpCode.Left + TextWidth(txtUpCode.Text) + 60
End Sub

Private Sub 权限控制()
    '权限控制
    Dim bln药品 As Boolean
    Dim bln物资 As Boolean
    Dim bln设备 As Boolean
    Dim bln其他 As Boolean
    Dim bln卫材 As Boolean
    
    bln药品 = InStr(1, mstrPrivs, "药品供应商") <> 0
    bln物资 = InStr(1, mstrPrivs, "物资供应商") <> 0
    bln设备 = InStr(1, mstrPrivs, "设备供应商") <> 0
    bln其他 = InStr(1, mstrPrivs, "其他供应商") <> 0
    bln卫材 = InStr(1, mstrPrivs, "卫材供应商") <> 0
    
    chkType(0).Enabled = bln药品 And mintEditType <> g查看
    chkType(1).Enabled = bln物资 And mintEditType <> g查看
    chkType(2).Enabled = bln设备 And mintEditType <> g查看
    chkType(3).Enabled = bln其他 And mintEditType <> g查看
    chkType(4).Enabled = bln卫材 And mintEditType <> g查看
End Sub
Private Function SetEditPro(ByVal str类型 As String) As Boolean
    '设置编辑权限
    
    Dim bln药品 As Boolean
    Dim bln物资 As Boolean
    Dim bln设备 As Boolean
    Dim bln其他 As Boolean
    Dim bln卫材 As Boolean
    
    bln药品 = InStr(1, mstrPrivs, "药品供应商") <> 0
    bln物资 = InStr(1, mstrPrivs, "物资供应商") <> 0
    bln设备 = InStr(1, mstrPrivs, "设备供应商") <> 0
    bln其他 = InStr(1, mstrPrivs, "其他供应商") <> 0
    bln卫材 = InStr(1, mstrPrivs, "卫材供应商") <> 0
    
    Err = 0: On Error GoTo ErrHand:
    SetEditPro = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

