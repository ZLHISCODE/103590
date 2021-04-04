VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPresSet 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "人员设置"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   Icon            =   "frmPresSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6555
      TabIndex        =   33
      Top             =   375
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6555
      TabIndex        =   34
      Top             =   975
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   6555
      TabIndex        =   35
      Top             =   7335
      Width           =   1100
   End
   Begin VB.Frame fra页 
      BorderStyle     =   0  'None
      Height          =   7200
      Index           =   0
      Left            =   105
      TabIndex        =   36
      Top             =   405
      Width           =   6255
      Begin MSComctlLib.TreeView tvw执业类别 
         Height          =   3735
         Left            =   3120
         TabIndex        =   39
         Tag             =   "1000"
         Top             =   7365
         Visible         =   0   'False
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   6588
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ils16"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.TextBox txt别名扩展 
         Height          =   270
         Left            =   1560
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1920
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "删除"
         Enabled         =   0   'False
         Height          =   350
         Left            =   4395
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "减少一个所属部门"
         Top             =   4305
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "增加"
         Height          =   350
         Left            =   3240
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "新增一个所属部门"
         Top             =   4305
         Width           =   1100
      End
      Begin VB.ComboBox cmb编码 
         Height          =   300
         Index           =   2
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   6045
         Width           =   1785
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   9
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   10
         Top             =   5265
         Width           =   1785
      End
      Begin VB.TextBox txt资格证书编号 
         Height          =   270
         Left            =   150
         MaxLength       =   30
         TabIndex        =   9
         Top             =   4905
         Width           =   2835
      End
      Begin VB.ListBox lst编码 
         Height          =   1320
         Index           =   8
         Left            =   600
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   3240
         Width           =   2445
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   0
         Top             =   120
         Width           =   1785
      End
      Begin VB.ListBox lst编码 
         Height          =   1740
         Index           =   7
         Left            =   3255
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   360
         Width           =   2790
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1200
         Width           =   1785
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   1
         Top             =   480
         Width           =   1785
      End
      Begin VB.ComboBox cmb编码 
         Height          =   300
         Index           =   4
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   6795
         Width           =   1785
      End
      Begin VB.ComboBox cmbStationNo 
         Height          =   300
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   5640
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   8
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2280
         Width           =   1785
      End
      Begin VB.CheckBox chk处方权标志 
         Caption         =   "处方权(&J)"
         Height          =   180
         Left            =   3240
         TabIndex        =   17
         Top             =   5715
         Width           =   1170
      End
      Begin VB.ComboBox cmbKss 
         Height          =   300
         Index           =   0
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   6045
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.ComboBox cmb编码 
         Height          =   300
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   1785
      End
      Begin VB.TextBox txt编码 
         Height          =   300
         Left            =   1200
         TabIndex        =   7
         Top             =   2640
         Width           =   1525
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&P"
         Height          =   300
         Left            =   2720
         TabIndex        =   38
         Top             =   2640
         Width           =   285
      End
      Begin VB.ComboBox cboSS 
         Height          =   300
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   6795
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.ComboBox cmbKss 
         Height          =   300
         Index           =   1
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   6420
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmd别名 
         Height          =   250
         Left            =   2685
         Picture         =   "frmPresSet.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1920
         Width           =   270
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   11
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1560
         Width           =   1785
      End
      Begin MSComctlLib.ListView lvw部门 
         Height          =   1605
         Left            =   3240
         TabIndex        =   16
         Top             =   2640
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   2831
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "所属部门"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "缺省标志"
            Object.Width           =   1147
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtp执业时间 
         Height          =   345
         Left            =   1200
         TabIndex        =   11
         ToolTipText     =   "指执业开始时间"
         Top             =   5625
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   228655105
         CurrentDate     =   40387
      End
      Begin zl9BaseItem.cboTree cbo技术职务 
         Height          =   300
         Left            =   1200
         TabIndex        =   13
         Top             =   6420
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "宋体"
         FontSize        =   9
         SplitString     =   "."
         sngSelDownWidth =   3980
         TopShowDown     =   -1  'True
      End
      Begin MSComctlLib.ImageList ils16 
         Left            =   2400
         Top             =   7350
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPresSet.frx":0B46
               Key             =   "Root"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPresSet.frx":1192
               Key             =   "Nature"
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   7
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   40
         Top             =   1920
         Width           =   1785
      End
      Begin VB.Label lbl说明 
         BackStyle       =   0  'Transparent
         Caption         =   "    说明：人员可以隶属于多个部门，但缺省部门有且只能有一个。双击或使用空格键可使用指定部门成为缺省部门。"
         Height          =   915
         Left            =   3240
         TabIndex        =   65
         Top             =   4800
         Width           =   2790
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "专业职务(&Z)"
         Height          =   180
         Index           =   16
         Left            =   150
         TabIndex        =   64
         Top             =   6465
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "管理职务(&M)"
         Height          =   180
         Index           =   15
         Left            =   150
         TabIndex        =   63
         Top             =   6090
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "执业时间(&I)"
         Height          =   180
         Index           =   27
         Left            =   150
         TabIndex        =   62
         ToolTipText     =   "指执业开始时间"
         Top             =   5715
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "执业证号(&Y)"
         Height          =   180
         Index           =   25
         Left            =   150
         TabIndex        =   61
         Top             =   5325
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "编号(&U)"
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   60
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lbl性质 
         AutoSize        =   -1  'True
         Caption         =   "工作性质(&R)"
         Height          =   180
         Left            =   3255
         TabIndex        =   59
         Top             =   120
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "所属部门(&D)"
         Height          =   180
         Index           =   4
         Left            =   3255
         TabIndex        =   58
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "简码(&S)"
         Height          =   180
         Index           =   3
         Left            =   510
         TabIndex        =   57
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "姓名(&N)"
         Height          =   180
         Index           =   2
         Left            =   510
         TabIndex        =   56
         Top             =   540
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "执业类别(&K)"
         Height          =   180
         Index           =   12
         Left            =   150
         TabIndex        =   55
         Top             =   2700
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "聘任职务(&T)"
         Height          =   180
         Index           =   18
         Left            =   150
         TabIndex        =   54
         Top             =   6840
         Width           =   990
      End
      Begin VB.Label lbl执业分类 
         AutoSize        =   -1  'True
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   1320
         TabIndex        =   53
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label lblStationNo 
         AutoSize        =   -1  'True
         Caption         =   "院区(&B)"
         Height          =   180
         Left            =   4560
         TabIndex        =   52
         Top             =   5715
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "别名(&A)"
         Height          =   180
         Index           =   13
         Left            =   510
         TabIndex        =   51
         Top             =   1980
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "签名(&Q)"
         Height          =   180
         Index           =   17
         Left            =   510
         TabIndex        =   50
         Top             =   2340
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "住院抗菌药物权限"
         Height          =   180
         Index           =   28
         Left            =   3240
         TabIndex        =   49
         Top             =   6090
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "性别(&X)"
         Height          =   180
         Index           =   6
         Left            =   510
         TabIndex        =   48
         Top             =   900
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "执业范围(&G)"
         Height          =   180
         Index           =   14
         Left            =   150
         TabIndex        =   47
         Top             =   3000
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "资格证书编号(&N)"
         Height          =   180
         Index           =   26
         Left            =   150
         TabIndex        =   46
         Top             =   4680
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "手术等级"
         Height          =   180
         Index           =   29
         Left            =   3960
         TabIndex        =   45
         Top             =   6840
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊抗菌药物权限"
         Height          =   180
         Index           =   30
         Left            =   3240
         TabIndex        =   44
         Top             =   6465
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "顺序(&R)"
         Height          =   180
         Index           =   32
         Left            =   510
         TabIndex        =   43
         Top             =   1620
         Width           =   630
      End
   End
   Begin VB.Frame fra页 
      BorderStyle     =   0  'None
      Height          =   7200
      Index           =   1
      Left            =   105
      TabIndex        =   66
      Top             =   405
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1305
         MaxLength       =   18
         TabIndex        =   24
         Top             =   1005
         Width           =   2325
      End
      Begin VB.ComboBox cmb编码 
         Height          =   300
         Index           =   1
         Left            =   1305
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   215
         Width           =   2325
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   5
         Left            =   1305
         MaxLength       =   20
         TabIndex        =   30
         Top             =   3375
         Width           =   2325
      End
      Begin VB.ComboBox cmb编码 
         Height          =   300
         Index           =   5
         Left            =   1305
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1425
         Width           =   2325
      End
      Begin VB.ComboBox cmb编码 
         Height          =   300
         Index           =   6
         Left            =   1305
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1815
         Width           =   2325
      End
      Begin VB.PictureBox pic外框 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   3945
         ScaleHeight     =   140
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   120
         Width           =   1920
         Begin VB.PictureBox pic镜框 
            Height          =   1755
            Left            =   60
            ScaleHeight     =   1695
            ScaleWidth      =   1725
            TabIndex        =   74
            Top             =   60
            Width           =   1785
            Begin VB.Image img照片 
               Appearance      =   0  'Flat
               Height          =   1185
               Left            =   15
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1635
            End
         End
         Begin VB.Label lbl图片说明 
            Alignment       =   2  'Center
            Height          =   210
            Left            =   135
            TabIndex        =   75
            Top             =   1860
            Width           =   1560
         End
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   4
         Left            =   1305
         MaxLength       =   20
         TabIndex        =   28
         Top             =   2610
         Width           =   2325
      End
      Begin VB.TextBox txtEdit 
         Height          =   2925
         Index           =   6
         Left            =   1290
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   3870
         Width           =   4560
      End
      Begin VB.CommandButton cmd照片 
         Caption         =   "文件(&F)"
         Height          =   345
         Index           =   0
         Left            =   3930
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2250
         Width           =   855
      End
      Begin VB.CommandButton cmd照片 
         Caption         =   "清除(&L)"
         Height          =   345
         Index           =   1
         Left            =   5025
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2250
         Width           =   825
      End
      Begin VB.TextBox txtDate 
         Height          =   300
         Index           =   0
         Left            =   1305
         TabIndex        =   23
         Top             =   600
         Width           =   2325
      End
      Begin VB.TextBox txtDate 
         Height          =   300
         Index           =   1
         Left            =   1305
         TabIndex        =   27
         Top             =   2190
         Width           =   2325
      End
      Begin VB.PictureBox pic签名图片 
         AutoRedraw      =   -1  'True
         Height          =   810
         Left            =   3945
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   70
         Top             =   2865
         Width           =   810
      End
      Begin VB.CommandButton cmd签名 
         Caption         =   "清除(&N)"
         Height          =   345
         Index           =   1
         Left            =   5025
         TabIndex        =   69
         Top             =   3330
         Width           =   825
      End
      Begin VB.PictureBox picSign 
         AutoRedraw      =   -1  'True
         Height          =   210
         Left            =   4635
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   68
         Top             =   2625
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd签名 
         Caption         =   "文件(&I)"
         Height          =   345
         Index           =   0
         Left            =   5025
         TabIndex        =   67
         Top             =   2865
         Width           =   825
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   10
         Left            =   1305
         MaxLength       =   11
         TabIndex        =   29
         Top             =   3000
         Width           =   2325
      End
      Begin MSComDlg.CommonDialog cdl照片 
         Left            =   4290
         Top             =   1620
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "身份证号(&G)"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   86
         Top             =   1065
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "民族(&K)"
         Height          =   180
         Index           =   7
         Left            =   630
         TabIndex        =   85
         Top             =   275
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "电子邮件(&M)"
         Height          =   180
         Index           =   9
         Left            =   240
         TabIndex        =   84
         Top             =   3435
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "参加工作(&W)"
         Height          =   180
         Index           =   11
         Left            =   270
         TabIndex        =   83
         Top             =   2250
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "学历(&S)"
         Height          =   180
         Index           =   19
         Left            =   630
         TabIndex        =   82
         Top             =   1485
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "所学专业(&P)"
         Height          =   180
         Index           =   20
         Left            =   270
         TabIndex        =   81
         Top             =   1875
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "座机电话(&T)"
         Height          =   180
         Index           =   8
         Left            =   240
         TabIndex        =   80
         Top             =   2670
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "个人简介(&D)"
         Height          =   180
         Index           =   10
         Left            =   270
         TabIndex        =   79
         Top             =   3930
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "出生日期(&B)"
         Height          =   180
         Index           =   5
         Left            =   270
         TabIndex        =   78
         Top             =   660
         Width           =   990
      End
      Begin VB.Label lbl签名说明 
         Caption         =   "签名图片"
         Height          =   810
         Left            =   4800
         TabIndex        =   77
         Top             =   2925
         Width           =   210
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "移动电话(&M)"
         Height          =   180
         Index           =   31
         Left            =   240
         TabIndex        =   76
         Top             =   3060
         Width           =   990
      End
   End
   Begin VB.Frame fra页 
      BorderStyle     =   0  'None
      Height          =   7200
      Index           =   2
      Left            =   105
      TabIndex        =   87
      Top             =   405
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txt数 
         Height          =   300
         Index           =   0
         Left            =   1410
         MaxLength       =   2
         TabIndex        =   32
         Top             =   270
         Width           =   525
      End
      Begin VB.ListBox lst编码 
         Height          =   1950
         Index           =   9
         Left            =   360
         Style           =   1  'Checkbox
         TabIndex        =   90
         Top             =   930
         Width           =   2085
      End
      Begin VB.ListBox lst编码 
         Height          =   1530
         Index           =   10
         Left            =   2970
         Style           =   1  'Checkbox
         TabIndex        =   89
         Top             =   930
         Width           =   2895
      End
      Begin VB.ListBox lst编码 
         Height          =   1740
         Index           =   11
         Left            =   360
         Style           =   1  'Checkbox
         TabIndex        =   88
         Top             =   3390
         Width           =   2115
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "接受培训(&R)"
         Height          =   180
         Index           =   23
         Left            =   2970
         TabIndex        =   94
         Top             =   690
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "留学时间(&I)        月"
         Height          =   180
         Index           =   21
         Left            =   360
         TabIndex        =   93
         Top             =   330
         Width           =   1890
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "留学渠道(&U)"
         Height          =   180
         Index           =   22
         Left            =   360
         TabIndex        =   92
         Top             =   690
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "科研课题(&J)"
         Height          =   180
         Index           =   24
         Left            =   360
         TabIndex        =   91
         Top             =   3120
         Width           =   990
      End
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   7620
      Left            =   75
      TabIndex        =   95
      Top             =   75
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   13441
      TabWidthStyle   =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "管理信息"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "详细情况"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "其它内容"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPresSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum const编码
    code性别 = 0
    code民族 = 1
    code管理职务 = 2
    '刘兴宏:2007/06/05屏蔽,解决易用性问题
    'code专业技术职务 = 3
    code聘任技术职务 = 4
    code学历 = 5
    code所学专业 = 6
    '以下的使用ListBox控件
    code人员性质 = 7
    code执业范围 = 8
    code留学渠道 = 9
    code接受培训 = 10
    code科研课题 = 11
End Enum

Private Enum const数
    Number留学时间 = 0
End Enum

Private Enum const文本
    Text身份证号 = 0
    Text编号 = 1
    Text姓名 = 2
    text简码 = 3
    Text电话 = 4
    Text电子邮件 = 5
    Text个人简介 = 6
    text别名 = 7
    Text签名 = 8
    text执业证号 = 9
    Text移动电话 = 10
    Text顺序 = 11
End Enum

Private Enum const日期
    Date出生日期 = 0
    Date参加工作 = 1
End Enum

Private mstrID As String             '当前编辑的人员ID
Private mblnChange As Boolean        '是否改变了
Private mbln照片 As Boolean          '当前人员是否有图片信息
Private mbln照片更改 As Boolean      '当照片发生更改时才为True
Private mbln签名图 As Boolean        '签名图是否有图
Private mbln签名图更改 As Boolean    '签名图更改时设为True
Private mblnLoad As Boolean          '为TRUE表示刚装入
Private mcol分类 As New Collection   '表示某种执业类别的分类
Private msngStartX As Single, msngStartY As Single    '移动前鼠标的位置
Private mstrKssZY As String           '进入窗体后抗生素权限的值，用于判断是否改变
Private mstrKssMZ As String
Private mstr民族 As String
Private mrs民族 As ADODB.Recordset   '记录查询中有多条记录的集合
Private mblnClick职务 As Boolean     '专业职务是否被点击
Private mbln抗菌药物 As Boolean         '抗菌药物修改权限，true-允许修改， false-不允许修改
Private mblnPACSInterface As Boolean        '启用影像信息系统接口
Private Sub IniStationNo()
    Dim strSQL As String
    Dim rsRecord As ADODB.Recordset
'    lblStationNo.Visible = False
'    cmbStationNo.Visible = False
'
'    If gstrNodeNo <> "-" Then
        On Error GoTo ErrHandle
        lblStationNo.Visible = True
        cmbStationNo.Visible = True
        
        strSQL = "select 编号,名称 from zlnodelist"
        Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "站点查询")
        With cmbStationNo
            .AddItem ""
            Do While Not rsRecord.EOF
                .AddItem rsRecord!编号 & "-" & rsRecord!名称
                rsRecord.MoveNext
            Loop
        End With
        
'        With cmbStationNo
'            .Clear
'            .AddItem ""
'            .AddItem "0"
'            .AddItem "1"
'            .AddItem "2"
'            .AddItem "3"
'            .AddItem "4"
'            .AddItem "5"
'            .AddItem "6"
'            .AddItem "7"
'            .AddItem "8"
'            .AddItem "9"
'
'            .ListIndex = 0
'        End With
'    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetStationNo(ByVal strNo As String)
    Dim n As Integer
    
'    If gstrNodeNo = "-" Then Exit Sub
    
    If strNo = "" Then
        cmbStationNo.ListIndex = 0
    Else
        For n = 1 To cmbStationNo.ListCount - 1
            If Mid(cmbStationNo.List(n), 1, InStr(1, cmbStationNo.List(n), "-") - 1) = strNo Then
                cmbStationNo.ListIndex = n
            End If
        Next
    End If
        
End Sub

Private Sub cbo技术职务_Change()
    Dim i As Long
    
    mblnClick职务 = False
    If mstrID <> "" And cmbKss(0).ListIndex > 0 Then
        Call CheckWorkNature
        Exit Sub
    End If
    If mstrID <> "" And cmbKss(1).ListIndex > 0 Then
        Call CheckWorkNature
        Exit Sub
    End If
    
    For i = 0 To 1
        If Mid(cbo技术职务.Text, InStr(cbo技术职务.Text, ".") + 1) = "医师" Then
            cmbKss(i).Text = "非限制使用"
            cmbKss(i).Enabled = True
        ElseIf Mid(cbo技术职务.Text, InStr(cbo技术职务.Text, ".") + 1) = "主治医师" Then
            cmbKss(i).Text = "限制使用"
            cmbKss(i).Enabled = True
        ElseIf Mid(cbo技术职务.Text, InStr(cbo技术职务.Text, ".") + 1) = "副主任医师" Or Mid(cbo技术职务.Text, InStr(cbo技术职务.Text, ".") + 1) = "主任医师" Then
            cmbKss(i).Text = "特殊使用"
            cmbKss(i).Enabled = True
        Else
            Call CheckWorkNature
            cmbKss(i).ListIndex = 0
        End If
    Next
End Sub

Private Sub cbo技术职务_DownClick()
    mblnClick职务 = True
End Sub

Private Sub cbo技术职务_LostFocus()
    mblnClick职务 = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub ClearContext()
    Dim lngCount As Integer
    Dim lngList As Integer
    
    mstrID = ""
    For lngCount = txtEdit.LBound To txtEdit.UBound
        txtEdit(lngCount).Text = ""
    Next
    For lngCount = txtDate.LBound To txtDate.UBound
        txtDate(lngCount).Text = ""
    Next
    For lngCount = txt数.LBound To txt数.UBound
        txt数(lngCount).Text = ""
    Next
    For lngCount = cmb编码.LBound To cmb编码.UBound
        If lngCount <> 3 Then
            cmb编码(lngCount).ListIndex = -1
            For lngList = 0 To cmb编码(lngCount).ListCount - 1
                '设置缺省值
                If cmb编码(lngCount).ItemData(lngList) = 1 Then
                    cmb编码(lngCount).ListIndex = lngList
                    Exit For
                End If
            Next
        End If
    Next
    For lngCount = lst编码.LBound To lst编码.UBound
        For lngList = 0 To lst编码(lngCount).ListCount - 1
            lst编码(lngCount).Selected(lngList) = False
        Next
    Next
    
    mbln照片 = False:   mbln照片更改 = False
    Call 显示空图片
    mbln签名图 = False: mbln签名图更改 = False
    Set pic签名图片.Picture = Nothing: pic签名图片.Tag = "": pic签名图片.Cls
    txtEdit(Text编号).Text = Sys.MaxCode("人员表", "编号", 6)
    mblnChange = False
End Sub

Private Sub cmdOK_Click()
    If IsValid() = False Then Exit Sub
    If Save人员() = False Then Exit Sub
    If mstrID <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    
    Call ClearContext
    tabMain.Tabs(1).Selected = True
    txtEdit(Text姓名).SetFocus
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim ctlError As Control
    
    Set ctlError = GetErrorObject
    
    If ctlError Is Nothing Then
        '没有检查到错误
        IsValid = True
    Else
        '显示到错误控件处
        lngCount = ctlError.Container.Index
        tabMain.Tabs(lngCount + 1).Selected = True
        ctlError.SetFocus
    End If
    
End Function

Private Function GetErrorObject() As Control
    Dim i As Integer
    Dim strTemp As String
    
    '检查文本型字段
    For i = txtEdit.LBound To txtEdit.UBound
        strTemp = Trim(txtEdit(i).Text)
        If zlCommFun.StrIsValid(strTemp, txtEdit(i).MaxLength) = False Then
            Set GetErrorObject = txtEdit(i)
            Exit Function
        End If
    Next
    '检查数字型字段
    For i = txt数.LBound To txt数.UBound
        If IntegerIsValid(Trim(txt数(i).Text), txt数(i).MaxLength) = False Then
            Set GetErrorObject = txt数(i)
            Exit Function
        End If
    Next
    
    '检查日期型控件
    For i = txtDate.LBound To txtDate.UBound
        If txtDate(i).Text <> "" Then
            If CDate(txtDate(i)) > Date Then
                MsgBox "输入日期超过当前日期。", vbInformation, gstrSysName
                Set GetErrorObject = txtDate(i)
                Exit Function
            End If
        End If
    Next
    If txtDate(Date参加工作) <> "" And txtDate(Date出生日期) <> "" Then
        If CDate(txtDate(Date参加工作)) <= CDate(txtDate(Date出生日期)) Then
            MsgBox "参加工作的日期须大于出生日期。", vbInformation, gstrSysName
            Set GetErrorObject = txtDate(Date参加工作)
            Exit Function
        End If
    End If
    
    '检查列表框控件
    For i = code执业范围 To code科研课题
        If lst编码(i).SelCount > 3 Then
            MsgBox "选择的项目不能超过3个。", vbExclamation, gstrSysName
            Set GetErrorObject = lst编码(i)
            Exit Function
        End If
    Next
        
    If Len(Trim(txtEdit(Text编号).Text)) = 0 Then
        txtEdit(Text编号).Text = ""
        MsgBox "编号不能为空。", vbExclamation, gstrSysName
        Set GetErrorObject = txtEdit(Text编号)
        Exit Function
    End If
    If Len(Trim(txtEdit(Text姓名).Text)) = 0 Then
        MsgBox "姓名不能为空。", vbExclamation, gstrSysName
        txtEdit(Text姓名).Text = ""
        Set GetErrorObject = txtEdit(Text姓名)
        Exit Function
    End If
        
    
    '对身份证号进行验证
    strTemp = txtEdit(Text身份证号)
    If strTemp <> "" Then
        '如果输入了身份证号
        If IntegerIsValid(Trim(Mid(strTemp, 1, Len(strTemp) - 1)), 17) = False Then
            Set GetErrorObject = txtEdit(Text身份证号)
            Exit Function
        End If
        
        Dim str出生日期 As String
        Dim lng性别 As Long
        
        If Len(strTemp) <> 15 And Len(strTemp) <> 18 Then
            Set GetErrorObject = txtEdit(Text身份证号)
            MsgBox "身份证号码长度不对。", vbInformation, gstrSysName
            Exit Function
        End If
        
        If Len(strTemp) = 15 Then
            '老式
            str出生日期 = Mid(strTemp, 7, 6)
            str出生日期 = zlCommFun.AddDate(str出生日期)
            
            lng性别 = Val(Right(strTemp, 1))
        Else
            '新式
            str出生日期 = Mid(strTemp, 7, 8)
            str出生日期 = zlCommFun.AddDate(str出生日期)
            
            lng性别 = Val(Mid(strTemp, 17, 1))
        End If
        If Not IsDate(str出生日期) Then
            Set GetErrorObject = txtEdit(Text身份证号)
            MsgBox "身份证号码中出生日期信息不正确。", vbInformation, gstrSysName
            Exit Function
        End If
        If Not IsDate(txtDate(Date出生日期).Text) Then
            Set GetErrorObject = txtDate(Date出生日期)
            MsgBox "请确认该出生日期是否正确。", vbInformation, gstrSysName
            Exit Function
        End If
        If CDate(str出生日期) <> CDate(txtDate(Date出生日期).Text) Then
            Set GetErrorObject = txtEdit(Text身份证号)
            MsgBox "身份证号码中出生日期信息与出生日期不对。", vbInformation, gstrSysName
            Exit Function
        End If
        If (lng性别 Mod 2 = 1 And InStr(cmb编码(code性别).Text, "女") > 0) Or (lng性别 Mod 2 = 0 And InStr(cmb编码(code性别).Text, "男") > 0) Then
            Set GetErrorObject = txtEdit(Text身份证号)
            MsgBox "身份证号码中性别信息不正确。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '医生性质人员必须选择执业类别
    For i = 0 To lst编码(code人员性质).ListCount - 1
        If lst编码(code人员性质).Selected(i) = True Then
            If lst编码(code人员性质).List(i) = "医生" And txt编码.Text = "" Then
                Set GetErrorObject = txt编码
                MsgBox "该人员具有医生性质，必须选择执业类别。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Next
'    '屏蔽该段代码，不进行检查
'    '问题27371、27388 by lesfeng 2010-01-19 增加设备及物资领用分类 新增 部门性质分类
'    '设备领用员以及物资领用员性质人员必须检查所选部门是具有设备领用及物资领用部门性质分类
'    For i = 0 To lst编码(code人员性质).ListCount - 1
'        If lst编码(code人员性质).Selected(i) = True Then
'            strTemp = Trim(lst编码(code人员性质).List(i))
'            If strTemp = "设备领用员" Then
'                If DeptIsValid(strTemp, 1) Then
'                    Set GetErrorObject = lst编码(code人员性质)
'                    MsgBox "该人员所属部门不具有‘设备领用’工作性质，请取消设备领用员设置。", vbInformation, gstrSysName
'                    Exit Function
'                End If
'            End If
'            If strTemp = "物资领用员" Then
'                If DeptIsValid(strTemp, 2) Then
'                    Set GetErrorObject = lst编码(code人员性质)
'                    MsgBox "该人员所属部门不具有‘物资领用’工作性质，请取消物资领用员设置。", vbInformation, gstrSysName
'                    Exit Function
'                End If
'            End If
'        End If
'    Next
    
End Function

Private Function IntegerIsValid(ByVal strInput As String, ByVal lng长度 As Long) As Boolean
    Dim sngTemp As Long
    
    If strInput = "" Then
        IntegerIsValid = True
        Exit Function
    End If
    If Not IsNumeric(strInput) Then
        MsgBox "请输入一个正确的数值。", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(strInput) < 0 Then
        MsgBox "请输入一个正数。", vbInformation, gstrSysName
        Exit Function
    End If
    '数值太大会出错
    On Error Resume Next
    sngTemp = Fix(Val(strInput))
    If Err <> 0 Then
        Err.Clear
        If lng长度 < 10 Then
            MsgBox "数据值过大。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    'Tag属性保存整据整数部分的长度
    If lng长度 < 10 Then
        If Len(CStr(sngTemp)) > lng长度 Then
            MsgBox "数据值过大。", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        If Len(strInput) > lng长度 Then
            MsgBox "数据值过大。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    IntegerIsValid = True
End Function

Private Function DeptIsValid(ByVal strInput As String, ByVal IntFlag As Integer) As Boolean
'--问题27371、27388 by lesfeng 2010-01-19
    Dim i As Integer
    Dim str部门ID As String
    
    DeptIsValid = True
    For i = 1 To lvw部门.ListItems.Count
        str部门ID = Mid(lvw部门.ListItems(i).Key, 2)
        DeptIsValid = DeptSQLIsValid(str部门ID, IntFlag)
        If Not DeptIsValid Then Exit Function
    Next
End Function

Private Function DeptSQLIsValid(ByVal strInput As String, ByVal IntFlag As Integer) As Boolean
'--问题27371、27388 by lesfeng 2010-01-19
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim int计数 As Integer
    
    
    DeptSQLIsValid = True
    On Error GoTo ErrHandle
    If IntFlag = 1 Then '10 设备
        strTemp = " And A.部门id = [1] And B.编码 = '10' "
    Else '11 物资
        strTemp = " And A.部门id = [1] And B.编码 = '11' "
    End If
    
    strSQL = " Select Count(A.部门id) As 计数 " & _
             "   From 部门性质说明 A, 部门性质分类 B " & _
             "  Where A.工作性质 = B.名称" & strTemp
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strInput))
    
    If Not rsTemp.EOF Then
        int计数 = IIF(IsNull(rsTemp!计数), 0, rsTemp!计数)
        If int计数 > 0 Then
            DeptSQLIsValid = False
        End If
    End If
    rsTemp.Close
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Function Save人员() As Boolean
    Dim lng人员id  As Long
    Dim str部门ID As String, str科室名称 As String
    Dim str人员性质 As String
    Dim i As Integer, int对象 As Integer
    Dim nod As Node
    Dim lst As ListItem
    Dim str专业技术职务 As String, blnTran As Boolean
    Dim str站点 As String
    Dim strSQL As String
    Dim curDate As Date
    Dim str编码 As String
    
    On Error GoTo ErrHandle
    
    If Trim(txtEdit(text别名).Text) = "" Then
        txtEdit(text别名).Text = txtEdit(Text姓名).Text
    End If
        
    '把所有部门做成一个串，选中的为1
    For i = 1 To lvw部门.ListItems.Count
        str部门ID = str部门ID & Mid(lvw部门.ListItems(i).Key, 2) & ":"
        If lvw部门.ListItems(i).SubItems(1) = "√" Then
            str部门ID = str部门ID & "1;"
            
            str科室名称 = Mid(lvw部门.ListItems(i).Text, InStr(lvw部门.ListItems(i).Text, "】") + 1)
        Else
            str部门ID = str部门ID & "0;"
        End If
    Next
    
    '把所有选中的工作性质做成一个串
    For i = 0 To lst编码(code人员性质).ListCount - 1
        If lst编码(code人员性质).Selected(i) = True Then
            str人员性质 = str人员性质 & lst编码(code人员性质).List(i) & ";"
        End If
    Next
    
    str专业技术职务 = cbo技术职务.Text
    If str专业技术职务 <> "" Then
        str专业技术职务 = "'" & Mid(str专业技术职务, InStr(1, str专业技术职务, ".") + 1) & "'"
    Else
        str专业技术职务 = "NULL"
    End If
    
    gcnOracle.BeginTrans: blnTran = True
    
    If cmbStationNo.Text = "" Then
        str站点 = "Null"
    Else
        str站点 = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
    End If
    If txt编码.Text <> "" Then
        str编码 = "'" & Mid(txt编码.Text, 1, InStr(1, txt编码.Text, ".") - 1) & "'"
    Else
        str编码 = "Null"
    End If
    
    '正式处理
    If mstrID = "" Then       '新增一条记录
        lng人员id = Sys.NextId("人员表")
        gstrSQL = "zl_人员表_新增(" & lng人员id & _
            ",'" & txtEdit(Text编号).Text & "','" & txtEdit(Text姓名).Text & "','" & txtEdit(text简码).Text & "','" & _
            txtEdit(Text身份证号).Text & "'," & IIF(txtDate(Date出生日期).Text = "", "null", "to_date('" & txtDate(Date出生日期).Text & "','yyyy-MM-dd')") & "," & _
            GetTextFromCombo(cmb编码(code性别), True, ".") & "," & GetTextFromCombo(cmb编码(code民族), True, ".") & "," & _
            IIF(txtDate(Date参加工作).Text = "", "null", "to_date('" & txtDate(Date参加工作).Text & "','yyyy-MM-dd')") & ",'" & _
            txtEdit(Text电话).Text & "','" & txtEdit(Text电子邮件).Text & "'," & _
            str编码 & "," & GetTextFromList(lst编码(code执业范围)) & "," & _
            GetTextFromCombo(cmb编码(code管理职务), True, ".") & "," & str专业技术职务 & "," & _
            GetTextFromCombo(cmb编码(code聘任技术职务), False, ".") & "," & GetTextFromCombo(cmb编码(code学历), True, ".") & "," & _
            GetTextFromCombo(cmb编码(code所学专业), False, ".") & ",'" & txt数(Number留学时间).Text & "'," & _
            GetTextFromList(lst编码(code留学渠道)) & "," & GetTextFromList(lst编码(code接受培训)) & "," & _
            GetTextFromList(lst编码(code科研课题)) & ",'" & txtEdit(Text个人简介).Text & "','" & _
            str部门ID & "','" & str人员性质 & "','" & txtEdit(text别名).Text & "'," & IIF(cmbStationNo.Visible = False Or cmbStationNo.Text = "", "Null", str站点) & _
            ",'" & txtEdit(Text签名).Text & "','" & txtEdit(text执业证号).Text & "','" & Me.txt资格证书编号.Text & _
            IIF(Me.dtp执业时间.value = Null, "',null", "',to_date('" & Me.dtp执业时间.value & "','yyyy-MM-dd')") & _
            "," & Me.chk处方权标志.value & _
            ",'" & Me.cboSS.Text & "','" & txtEdit(Text移动电话).Text & "'," & _
            IIF(txtEdit(Text顺序).Text = "", "Null", txtEdit(Text顺序).Text) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Else
        '修改
        lng人员id = Val(mstrID)
        gstrSQL = "zl_人员表_修改(" & lng人员id & _
            ",'" & txtEdit(Text编号).Text & "','" & txtEdit(Text姓名).Text & "','" & txtEdit(text简码).Text & "','" & _
            txtEdit(Text身份证号).Text & "'," & IIF(txtDate(Date出生日期).Text = "", "null", "to_date('" & txtDate(Date出生日期).Text & "','yyyy-MM-dd')") & "," & _
            GetTextFromCombo(cmb编码(code性别), True, ".") & "," & GetTextFromCombo(cmb编码(code民族), True, ".") & "," & _
            IIF(txtDate(Date参加工作).Text = "", "null", "to_date('" & txtDate(Date参加工作).Text & "','yyyy-MM-dd')") & ",'" & _
            txtEdit(Text电话).Text & "','" & txtEdit(Text电子邮件).Text & "'," & _
            str编码 & "," & GetTextFromList(lst编码(code执业范围)) & "," & _
            GetTextFromCombo(cmb编码(code管理职务), True, ".") & "," & str专业技术职务 & "," & _
            GetTextFromCombo(cmb编码(code聘任技术职务), False, ".") & "," & GetTextFromCombo(cmb编码(code学历), True, ".") & "," & _
            GetTextFromCombo(cmb编码(code所学专业), False, ".") & ",'" & txt数(Number留学时间).Text & "'," & _
            GetTextFromList(lst编码(code留学渠道)) & "," & GetTextFromList(lst编码(code接受培训)) & "," & _
            GetTextFromList(lst编码(code科研课题)) & ",'" & txtEdit(Text个人简介).Text & "','" & _
            str部门ID & "','" & str人员性质 & "','" & txtEdit(text别名).Text & "'," & IIF(cmbStationNo.Visible = False Or cmbStationNo.Text = "", "Null", str站点) & _
            ",'" & txtEdit(Text签名).Text & "','" & txtEdit(text执业证号).Text & "','" & Me.txt资格证书编号.Text & _
            IIF(Me.dtp执业时间.value = Null, "',null", "',to_date('" & Me.dtp执业时间.value & "','yyyy-MM-dd')") & _
            "," & Me.chk处方权标志.value & _
            ",'" & Me.cboSS.Text & "','" & txtEdit(Text移动电话).Text & "'," & _
            IIF(txtEdit(Text顺序).Text = "", "Null", txtEdit(Text顺序).Text) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    '处理人员照片
    If mbln照片更改 = True Then
        '只有发生了更改才需要处理
        'zlDatabase.ExecuteProcedure "delete from 人员照片 where 人员ID=" & lng人员id, Me.Caption
        Call zlDatabase.ExecuteProcedure("zl_人员照片_Delete(" & lng人员id & ")", Me.Caption)
        If mbln照片 = True Then
            '保存
            If Sys.SaveLob(100, 16, lng人员id, img照片.Tag) = False Then
                gcnOracle.RollbackTrans
                MsgBox "照片保存失败。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    If mbln签名图更改 Then
        If Save签名图片(lng人员id, True) = False Then
            gcnOracle.RollbackTrans
            MsgBox "照片清除失败。", vbInformation, gstrSysName
            Exit Function
        End If
        If mbln签名图 Then
            If Save签名图片(lng人员id, False) = False Then
                gcnOracle.RollbackTrans
                MsgBox "照片保存失败。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '住院抗菌药物授权
    If mbln抗菌药物 = True Then
        If mstrKssZY <> cmbKss(0).Text Then
            If mstrID <> "" Or cmbKss(0).Text <> "" Then
                curDate = Sys.Currentdate
                strSQL = "Zl_人员抗菌药物权限_Update('" & _
                         lng人员id & "'," & _
                         cmbKss(0).ListIndex & ",'" & _
                         gstrUserName & "'," & _
                         "to_date('" & curDate & "','YYYY-MM-DD HH24:MI:SS'), 1) "
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
        End If
    End If
    '门诊抗菌药物授权
    If mbln抗菌药物 = True Then
        If mstrKssMZ <> cmbKss(1).Text Then
            If mstrID <> "" Or cmbKss(1).Text <> "" Then
                curDate = Sys.Currentdate
                strSQL = "Zl_人员抗菌药物权限_Update('" & _
                         lng人员id & "'," & _
                         cmbKss(1).ListIndex & ",'" & _
                         gstrUserName & "'," & _
                         "to_date('" & curDate & "','YYYY-MM-DD HH24:MI:SS'), 2) "
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
        End If
    End If
    
    '新网RIS接口，修改人员信息（因为主要是同步用户信息，不处理新增人员）；标准版，启用参数，部门性质为“检查”的部门人员，接口部件有效的前提下
    If Int(glngSys / 100) = 1 And mblnPACSInterface = True And mstrID <> "" Then
        If IsCheckDeptPres(lng人员id) Then
            If Not gobjRIS Is Nothing Then
                If gobjRIS.HISBasicDictTable(RISBaseItemType.Personnel, RISBaseItemOper.Modify, lng人员id) <> 1 Then
                    gcnOracle.RollbackTrans
                    
                    '出错时提示接口错误信息
                    If gobjRIS.LastErrorInfo <> "" Then
                        MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                    Else
                        MsgBox "调用RIS接口错误，不能继续当前操作！请与系统管理员联系", vbInformation, gstrSysName
                    End If
                    
                    Exit Function
                End If
            Else
                gcnOracle.RollbackTrans
                
                '接口部件无效时禁止并提示
                MsgBox "RIS接口创建失败，不能继续当前操作！可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
                
                Exit Function
            End If
        End If
    End If
    
    gcnOracle.CommitTrans: blnTran = False
    
    frmPresManage.FillList frmPresManage.tvwMain_S.SelectedItem.Key
    Save人员 = True
    Exit Function
ErrHandle:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function IsCheckDeptPres(ByVal lngPres As Long) As Boolean
    '是否检查科室人员
    Dim rsData  As ADODB.Recordset
    
    gstrSQL = "Select 1 From 部门人员 A, 部门性质说明 B Where a.部门id = b.部门id And 工作性质 = '检查' And a.人员id = [1] "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "IsCheckDeptPres", lngPres)
    
    IsCheckDeptPres = Not rsData.EOF
End Function
Public Function 编辑人员(Optional strID As String = "", Optional ByVal str部门ID As String = "") As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim varTemp As Variant
    Dim strTempFile As String
    Dim strSQL As String
    Dim strTemp As String
    Dim j As Integer
    Dim rs人员性质 As New ADODB.Recordset
    Dim blnKind As Boolean
    
    rsTemp.CursorLocation = adUseClient
   
    img照片.ToolTipText = "镜框大小：" & img照片.Width & "×" & img照片.Height
    Call InitEnv
   
    Call IniStationNo
    
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    
    mbln照片 = False:    mbln照片更改 = False
    mbln签名图 = False:  mbln签名图更改 = False
    On Error GoTo ErrHandle
    mstrID = strID
    If strID <> "" Then
        Dim i As Integer, varValue As Variant
        gstrSQL = "Select  a.ID, a.编号,b.编码 as 执业类别编码, a.姓名, a.简码, a.身份证号, a.出生日期, a.性别, a.民族, a.工作日期, a.办公室电话,a.移动电话, a.电子邮件,b.名称 as 执业类别, a.执业范围, a.管理职务, a.专业技术职务,a.聘任技术职务, a.学历, a.所学专业, a.留学时间, a.留学渠道," & _
                          " a.接受培训 , a.科研课题, a.个人简介, a.别名, a.站点, a.签名,a.执业证号, a.资格证书号, a.执业开始日期, a.处方权标志,a.手术等级,a.顺序 " & _
                   " From 人员表 a,执业类别 b" & _
                   " Where a.ID = [1] and a.执业类别=b.编码(+) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
                
        txtEdit(Text编号).Text = rsTemp("编号")
        txtEdit(Text姓名).Text = rsTemp("姓名")
        txtEdit(text简码).Text = IIF(IsNull(rsTemp("简码")), "", rsTemp("简码"))
        txtEdit(text别名).Text = IIF(IsNull(rsTemp("别名")), "", rsTemp("别名"))
        txtEdit(Text签名).Text = IIF(IsNull(rsTemp("签名")), "", rsTemp("签名"))
        txtEdit(Text顺序).Text = IIF(IsNull(rsTemp("顺序")), "", rsTemp("顺序"))
        
        txtEdit(Text身份证号).Text = IIF(IsNull(rsTemp("身份证号")), "", rsTemp("身份证号"))
        txtEdit(Text电话).Text = IIF(IsNull(rsTemp("办公室电话")), "", rsTemp("办公室电话"))
        txtEdit(Text移动电话).Text = IIF(IsNull(rsTemp("移动电话")), "", rsTemp("移动电话"))
        txtEdit(Text电子邮件).Text = IIF(IsNull(rsTemp("电子邮件")), "", rsTemp("电子邮件"))
        txtEdit(Text个人简介).Text = IIF(IsNull(rsTemp("个人简介")), "", rsTemp("个人简介"))
        
        txt数(Number留学时间).Text = IIF(IsNull(rsTemp("留学时间")), "", rsTemp("留学时间"))
        
        SetComboByText cmb编码(code性别), IIF(IsNull(rsTemp("性别")), "", rsTemp("性别")), True, "."
        SetComboByText cmb编码(code民族), IIF(IsNull(rsTemp("民族")), "", rsTemp("民族")), True, "."
        SetComboByText cmb编码(code学历), IIF(IsNull(rsTemp("学历")), "", rsTemp("学历")), True, "."
        
        
        'SetComboByText cmb编码(code专业技术职务), IIF(IsNull(rsTemp("专业技术职务")), "", rsTemp("专业技术职务")), True, "."
        
        SetComboByText cmb编码(code管理职务), IIF(IsNull(rsTemp("管理职务")), "", rsTemp("管理职务")), True, "."
        SetComboByText cmb编码(code所学专业), IIF(IsNull(rsTemp("所学专业")), "", rsTemp("所学专业")), False, "."
        
        txt编码.Text = IIF(IsNull(rsTemp("执业类别")), "", rsTemp!执业类别编码 & "." & rsTemp("执业类别"))
        SetComboByText cmb编码(code聘任技术职务), IIF(IsNull(rsTemp("聘任技术职务")), "", rsTemp("聘任技术职务")), False, "."
        
        SetListByText lst编码(code执业范围), IIF(IsNull(rsTemp("执业范围")), "", rsTemp("执业范围"))
        SetListByText lst编码(code留学渠道), IIF(IsNull(rsTemp("留学渠道")), "", rsTemp("留学渠道"))
        SetListByText lst编码(code接受培训), IIF(IsNull(rsTemp("接受培训")), "", rsTemp("接受培训"))
        SetListByText lst编码(code科研课题), IIF(IsNull(rsTemp("科研课题")), "", rsTemp("科研课题"))
        
        txtDate(Date出生日期).Text = Format(rsTemp("出生日期"), "yyyy-MM-dd")
        txtDate(Date参加工作).Text = Format(rsTemp("工作日期"), "yyyy-MM-dd")
        
        txtEdit(text执业证号).Text = IIF(IsNull(rsTemp!执业证号), "", rsTemp!执业证号)
        txt资格证书编号.Text = IIF(IsNull(rsTemp!资格证书号), "", rsTemp!资格证书号)
        chk处方权标志.value = IIF(IsNull(rsTemp!处方权标志), 0, rsTemp!处方权标志)
        dtp执业时间.value = IIF(IsNull(rsTemp!执业开始日期), Null, rsTemp!执业开始日期)
        If NVL(rsTemp!手术等级) <> "" Then
            cboSS.Text = NVL(rsTemp!手术等级)
        Else
            cboSS.ListIndex = 0
        End If
        
        
        SetStationNo (IIF(IsNull(rsTemp("站点")), "", rsTemp("站点")))
        
        strTempFile = Sys.ReadLob(100, 15, Val(strID))
        If strTempFile <> "" Then
            picSign.Picture = LoadPicture(strTempFile)
            pic签名图片.PaintPicture picSign.Picture, 0, 0, pic签名图片.ScaleX(pic签名图片.Width, vbTwips, vbPixels), pic签名图片.ScaleY(pic签名图片.Height, vbTwips, vbPixels)
            Kill strTempFile
        End If
        
        strTempFile = Trim(NVL(rsTemp!专业技术职务))
        If strTempFile <> "" Then
            gstrSQL = "Select 编码 From 专业技术职务 where 名称 =[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTempFile)
            If rsTemp.EOF Then
            Else
                cbo技术职务.SelectItemID = NVL(rsTemp!编码)
                cbo技术职务.Text = strTempFile
            End If
        End If
        
        '处理部门列表
'        If rsTemp.State = adStateOpen Then rsTemp.Close
        gstrSQL = "select C.部门ID,b.名称 as 部门,b.编码 as 部门编码,c.缺省" & _
                    "  from 部门表 b,部门人员 C " & _
                    " where C.部门ID=B.ID and C.人员id=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
        lvw部门.ListItems.Clear
        Do Until rsTemp.EOF
            lvw部门.ListItems.Add , "C" & rsTemp("部门ID"), "【" & rsTemp("部门编码") & "】" & rsTemp("部门")
            If rsTemp("缺省") = 1 Then
                lvw部门.ListItems("C" & rsTemp("部门ID")).SubItems(1) = "√"
            End If
            rsTemp.MoveNext
        Loop
        
        '处理图片
        strTempFile = Sys.ReadLobV2("人员照片", "照片", "人员ID=[1]", "", Val(strID))
        img照片.Picture = LoadPicture(strTempFile)
        mbln照片 = True
        lbl图片说明 = GetPictureInfo(img照片.Picture)
        '删除该临时文件
        If lbl图片说明 <> "无照片" Then
            Kill strTempFile
        End If
        
        '住院抗菌药物权限
        strSQL = "Select Max(级别) 级别 from 人员抗菌药物权限 where 人员ID=[1] And 记录状态=1 and 场合 = 1 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
        
        Call cbo.SetIndex(cmbKss(0).hwnd, Val(rsTemp!级别 & ""))
        '门诊抗菌药物权限
        strSQL = "Select Max(级别) 级别 from 人员抗菌药物权限 where 人员ID=[1] And 记录状态=1 and 场合 = 2 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
        Call cbo.SetIndex(cmbKss(1).hwnd, Val(rsTemp!级别 & ""))
        
    Else
        txtEdit(Text编号).Text = Sys.MaxCode("人员表", "编号", 6)
        
        '读出缺省的部门表
        gstrSQL = "select a.ID,a.名称 ,a.编码  from 部门表 A  where A.ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(str部门ID))
                
        lvw部门.ListItems.Clear
        lvw部门.ListItems.Add , "C" & rsTemp("ID"), "【" & rsTemp("编码") & "】" & rsTemp("名称")
        lvw部门.ListItems("C" & rsTemp("ID")).SubItems(1) = "√"
    End If
    
    '记录初始的抗生素值
    mstrKssZY = cmbKss(0).Text
    mstrKssMZ = cmbKss(1).Text
    
    If Not lvw部门.SelectedItem Is Nothing Then
        If InStr(frmPresManage.mstrPrivs, "所有部门") = 0 Then
            If Val(mstrID) = glngUserId Then
                cmdRemove.Enabled = False
                cmdAdd.Enabled = False
            Else
                If CheckDeptPermission(1, Mid(lvw部门.SelectedItem.Key, 2)) = False Then
                    cmdRemove.Enabled = False
                Else
                    cmdRemove.Enabled = lvw部门.SelectedItem.SubItems(1) = ""
                End If
            End If
        Else
            cmdRemove.Enabled = lvw部门.SelectedItem.SubItems(1) = ""
        End If
    End If
    
    '列出该人员的性质
    If rsTemp.State = 1 Then rsTemp.Close
    If strID = "" Then
         gstrSQL = "select 名称,null as 人员性质 from 人员性质分类 order by 编码"
    Else
         gstrSQL = "select A.名称,B.人员性质 from 人员性质分类 A,人员性质说明 B where A.名称=B.人员性质(+) and b.人员ID(+)=[1] order by decode(人员性质,null,1,0),A.编码"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
    Dim lst As ListItem
    Do Until rsTemp.EOF
        lst编码(code人员性质).AddItem rsTemp("名称")
        If Not IsNull(rsTemp("人员性质")) Then lst编码(code人员性质).Selected(lst编码(code人员性质).NewIndex) = True
        rsTemp.MoveNext
    Loop
    
    For j = 0 To lst编码(code人员性质).ListCount - 1
        If lst编码(code人员性质).Selected(j) = True And (lst编码(code人员性质).List(j) = "医生" Or lst编码(code人员性质).List(j) = "护士") Then
            strTemp = lst编码(code人员性质).List(j)
        End If
    Next
    
    '根据权限判断是否可以修改
    If strID <> "" And InStr(frmPresManage.mstrPrivs, ";修改时不限定人员性质;") = 0 Then
        gstrSQL = "Select 人员性质 From 人员性质说明 Where 人员id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "操作员性质查询", glngUserId)
        If rsTemp.RecordCount <> 0 Then
            For j = 0 To lst编码(code人员性质).ListCount - 1
                rsTemp.MoveFirst
                If lst编码(code人员性质).Selected(j) = True Then
                    Do While Not rsTemp.EOF
                        If lst编码(code人员性质).List(j) = rsTemp!人员性质 Then
                            blnKind = True
                            Exit Do
                        End If
                        If Not rsTemp.EOF Then
                            rsTemp.MoveNext
                        End If
                    Loop
                End If
            Next
            If blnKind = False Then
                fra页(0).Enabled = False
                fra页(1).Enabled = False
                fra页(2).Enabled = False
            Else
                For j = 1 To txtEdit.UBound
                    txtEdit(j).Enabled = False
                Next
                txt编码.Enabled = False
                txt资格证书编号.Enabled = False
                cmb编码(0).Enabled = False
                cmb编码(2).Enabled = False
                cmb编码(4).Enabled = False
                dtp执业时间.Enabled = False
                cbo技术职务.Enabled = False
                cmdSelect.Enabled = False
                lst编码(8).Enabled = False
                lst编码(7).Enabled = False
                chk处方权标志.Enabled = False
                cmbStationNo.Enabled = False
                cmbKss(0).Enabled = False
                cmbKss(1).Enabled = False
                cboSS.Enabled = False
                
                fra页(1).Enabled = False
                fra页(2).Enabled = False
            End If
        End If
    End If
    
    gstrSQL = " Select 编码 as ID,decode( substr(编码,1,2),编码,null ,substr(编码,1,2)) 上级ID,编码,名称,简码 From 专业技术职务 order by 编码"
    zlDatabase.OpenRecordset rs人员性质, gstrSQL, Me.Caption
    With rs人员性质
        If .EOF Then
            MsgBox "专业技术职务未安装,请找系统管理员！", vbInformation, gstrSysName
            Exit Function
        End If
        
        If cbo技术职务.FullCboData(rs人员性质, "", "编码,名称,简码", "编码|1000,名称|2000,简码|800", "", strTemp) = False Then
            MsgBox "数据加载有误,请查看专业技术职是否正确！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    Call 显示空图片
    '初始化完成
    mblnChange = False
    frmPresSet.Show vbModal
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdAdd_Click()
    Dim blnRe As Boolean
    Dim lngCount As Long
    Dim strID As String, str名称 As String, str编码 As String
    
    If InStr(frmPresManage.mstrPrivs, "所有部门") = 0 Then
        blnRe = frmTreeSel.ShowTreePrivs(IIF(mstrID = "", glngUserId, mstrID), strID, str名称, str编码)
    Else
        gstrSQL = "select id,上级id,名称,编码,Upper(简码) as 简码 from 部门表 where 撤档时间=to_date('3000-01-01','YYYY-MM-DD')  start with 上级id is null connect by prior id =上级id "
        blnRe = frmTreeSel.ShowTree(gstrSQL, strID, str名称, str编码, "", "人员表", "所有部门", False)
    End If
    DoEvents
    If blnRe Then
        For lngCount = 1 To lvw部门.ListItems.Count
            If Mid(lvw部门.ListItems(lngCount).Key, 2) = strID Then
                MsgBox "“" & str名称 & "”已经是该人员的所属部门了。", vbExclamation, gstrSysName
                Exit Sub
            End If
        Next
        lvw部门.ListItems.Add , "C" & strID, "【" & str编码 & "】" & str名称
        lvw部门.Refresh
        mblnChange = True
    End If
    
    If CheckOrder = True Then
        txtEdit(Text顺序).SetFocus
        Exit Sub
    End If
    
    Call lvw部门_ItemClick(lvw部门.SelectedItem)
End Sub

Private Sub cmdRemove_Click()
    If lvw部门.SelectedItem Is Nothing Then Exit Sub
    
    lvw部门.ListItems.Remove lvw部门.SelectedItem.Key
    lvw部门.ListItems(1).Selected = True
    
    Call lvw部门_ItemClick(lvw部门.SelectedItem)
End Sub

Private Sub cmdSelect_Click()
    With tvw执业类别
        .Top = txt编码.Top + txt编码.Height
        .Left = txt编码.Left
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub cmd别名_Click()
    If txt别名扩展.Visible = False Then
        txt别名扩展.Visible = True
        txt别名扩展.SetFocus
    End If
End Sub

Private Sub cmd签名_Click(Index As Integer)
Dim sZoom As Single, lDesWidth As Long, lDesHeight As Long
    If Index = 1 Then '-清空图片
        Set pic签名图片.Picture = Nothing
        pic签名图片.Tag = ""
        pic签名图片.Cls
        mbln签名图更改 = True:   mbln签名图 = False
    Else
        With cdl照片
            .CancelError = True
            .Filter = "图片文件(*.bmp,*.gif,*.jpg)|*.bmp;*.gif;*.jpg"
            
            On Error Resume Next
            .ShowOpen
            If Err <> 0 Then
                '没选中文件
                Err.Clear
            Else
                pic签名图片.Cls
                Set pic签名图片.Picture = Nothing
                pic签名图片.Picture = LoadPicture(.FileName)
                '高度不超过50像素,保持纵横比,另存图片后缀加.PIC
                If pic签名图片.ScaleY(pic签名图片.Picture.Height, vbHimetric, vbPixels) <= 50 Then
                    lDesWidth = pic签名图片.ScaleX(pic签名图片.Picture.Width, vbHimetric, vbPixels)
                    lDesHeight = pic签名图片.ScaleY(pic签名图片.Picture.Height, vbHimetric, vbPixels)
                Else
                    sZoom = pic签名图片.ScaleY(pic签名图片.Picture.Height, vbHimetric, vbPixels) / pic签名图片.ScaleX(pic签名图片.Picture.Width, vbHimetric, vbPixels)
                    lDesHeight = 50: lDesWidth = 50 / sZoom
                End If
                pic签名图片.PaintPicture pic签名图片.Picture, 0, 0, lDesWidth, lDesHeight
                picSign.Cls: Set picSign.Picture = Nothing
                picSign.Width = picSign.ScaleX(lDesWidth, vbPixels, vbTwips) + 45: picSign.Height = picSign.ScaleY(lDesHeight, vbPixels, vbTwips) + 45
                picSign.PaintPicture pic签名图片.Picture, 0, 0, lDesWidth, lDesHeight
                SavePicture picSign.Image, Mid(.FileName, 1, Len(.FileName) - 3) & "PIC"
                If Err <> 0 Then
                    MsgBox "图片文件无效，或文件不存在。", vbInformation, ""
                    Err.Clear
                    Exit Sub
                End If
                pic签名图片.Tag = Mid(.FileName, 1, Len(.FileName) - 3) & "PIC"
                mbln签名图更改 = True:   mbln签名图 = True
            End If
        End With
    End If
End Sub

Private Sub img照片_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
        msngStartY = Y
    End If
End Sub

Private Sub img照片_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngLeft As Single
    Dim sngTop As Single
    
    '缩放状态不处理
    If img照片.Stretch = True Then Exit Sub
    If Button = 1 Then
        '首先求出可能的
        sngLeft = img照片.Left + X - msngStartX
        sngTop = img照片.Top + Y - msngStartY
        
        '设置可能的左边距
        If img照片.Width < pic镜框.ScaleWidth Or sngLeft > pic镜框.ScaleLeft Then
            sngLeft = pic镜框.ScaleLeft
        Else
            If sngLeft + img照片.Width < pic镜框.ScaleWidth Then
                sngLeft = pic镜框.ScaleWidth - img照片.Width
            End If
        End If
        '设置可能的顶边距
        If img照片.Height < pic镜框.ScaleHeight Or sngTop > pic镜框.ScaleTop Then
            sngTop = pic镜框.ScaleTop
        Else
            If sngTop + img照片.Height < pic镜框.ScaleHeight Then
                sngTop = pic镜框.ScaleHeight - img照片.Height
            End If
        End If
        img照片.Left = sngLeft
        img照片.Top = sngTop
    End If
End Sub

Private Sub cmd照片_Click(Index As Integer)
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
                    img照片.Picture = LoadPicture(.FileName)
                    img照片.Left = pic镜框.ScaleLeft
                    img照片.Top = pic镜框.ScaleTop
                    
                    If Err <> 0 Then
                        MsgBox "图片文件无效，或文件不存在。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    lbl图片说明 = GetPictureInfo(img照片.Picture)
                    img照片.Tag = .FileName
                    mbln照片 = True
                    mbln照片更改 = True
                End If
            End With
        Case 1 '清除
            mbln照片 = False
            mbln照片更改 = True
            Call 显示空图片
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab
            If Shift = vbCtrlMask Then
                If tabMain.SelectedItem.Index = tabMain.Tabs.Count Then
                    tabMain.Tabs(1).Selected = True
                Else
                    tabMain.Tabs(tabMain.SelectedItem.Index + 1).Selected = True
                End If
            ElseIf Shift = (vbCtrlMask Or vbShiftMask) Then
                If tabMain.SelectedItem.Index = 1 Then
                    tabMain.Tabs(tabMain.Tabs.Count).Selected = True
                Else
                    tabMain.Tabs(tabMain.SelectedItem.Index - 1).Selected = True
                End If
            End If
        Case vbKeyPageDown
            Call OS.PressKeyEx(vbKeyTab, vbKeyShift)
            Exit Sub
        Case vbKeyPageUp
            Call OS.PressKeyEx(vbKeyTab, vbKeyShift)
            Exit Sub
        Case vbKeyEscape
            If mblnClick职务 = True Then
                mblnClick职务 = False
            Else
                Unload Me
                Exit Sub
            End If
    End Select
    
    If KeyCode = vbKeyReturn Then
        If ActiveControl Is lvw部门 Then
            tabMain.Tabs(2).Selected = True
        Else
            If Shift = 0 Then
               ' KeyCode = 0
                OS.PressKey vbKeyTab
            End If
        End If
        Exit Sub
    End If
    
    If Left(ActiveControl.Name, 3) = "dtp" Then
        Select Case KeyCode
            Case vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9, vbKeyReturn, vbKeyEscape, vbKeyDelete
            
            Case Else
                KeyCode = 0
        End Select
    End If
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call tabMain_Click
    End If
    
    mblnLoad = False
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim blnSign As Boolean
    Dim blnModify As Boolean
    Dim strPrivs As String
    Dim j As Integer
    
    tvw执业类别.Visible = False
    mblnLoad = True
    For j = 0 To 1
        cmbKss(j).Enabled = False
        For i = 0 To lst编码(code人员性质).ListCount - 1
            If lst编码(code人员性质).List(i) = "医生" Then
                If lst编码(code人员性质).Selected(i) Then
                    cmbKss(j).Enabled = True
                    Exit For
                End If
            End If
        Next
        If Mid(cbo技术职务.Text, InStr(cbo技术职务.Text, ".") + 1) = "医师" Or Mid(cbo技术职务.Text, InStr(cbo技术职务.Text, ".") + 1) = "主治医师" Or Mid(cbo技术职务.Text, InStr(cbo技术职务.Text, ".") + 1) = "副主任医师" Or Mid(cbo技术职务.Text, InStr(cbo技术职务.Text, ".") + 1) = "主任医师" Then
            cmbKss(j).Enabled = True
        End If
    Next
    
    strPrivs = ";" & zl9ComLib.GetPrivFunc(glngSys, 1002) & ";"
    blnSign = (Val(zlDatabase.GetPara("电子签名认证中心", glngSys)) = 0)
    blnModify = InStr(strPrivs, ";修改电子签名图片;") > 0
    If blnSign Then
        cmd签名(0).Visible = True
        cmd签名(1).Visible = True
    Else
        If blnModify Then
            cmd签名(0).Visible = True
            cmd签名(1).Visible = True
        Else
            cmd签名(0).Visible = False
            cmd签名(1).Visible = False
        End If
    End If
End Sub
Private Sub Form_Resize()
    Dim intFra As Integer
    
'    tabMain.Left = 120
'    tabMain.Top = 120
'    tabMain.Width = ScaleWidth - 240
'    tabMain.Height = cmdOK.Top - 240
    
    For intFra = 0 To 1
        fra页(intFra).Left = tabMain.ClientLeft
        fra页(intFra).Top = tabMain.ClientTop
        fra页(intFra).Height = tabMain.ClientHeight
        fra页(intFra).Width = tabMain.ClientWidth
        fra页(intFra).Visible = False
    Next
    
    With txt别名扩展
        .Visible = False
        .Top = lblEdit(13).Top + lblEdit(13).Height + 50
        .Left = lblEdit(13).Left
        .Width = txtEdit(7).Left - lblEdit(13).Left + txtEdit(7).Width + 50
        .Height = 2000
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
    If Cancel <> 1 Then Set mcol分类 = Nothing
    mstr民族 = ""
    Set mrs民族 = Nothing
    
End Sub

Private Sub lst编码_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub lst编码_ItemCheck(Index As Integer, Item As Integer)
    If Index = code人员性质 Then
        If lst编码(code人员性质).List(Item) = "医生" Then
            If lst编码(code人员性质).Selected(Item) Then
                cmbKss(0).Enabled = True
                cmbKss(1).Enabled = True
                cboSS.Enabled = True
            Else
                If Mid(cbo技术职务.Text, InStr(cbo技术职务.Text, ".") + 1) = "医师" Or Mid(cbo技术职务.Text, InStr(cbo技术职务.Text, ".") + 1) = "主治医师" Or Mid(cbo技术职务.Text, InStr(cbo技术职务.Text, ".") + 1) = "副主任医师" Or Mid(cbo技术职务.Text, InStr(cbo技术职务.Text, ".") + 1) = "主任医师" Then
                    cmbKss(0).Enabled = True
                    cmbKss(1).Enabled = True
                Else
                    cmbKss(0).Enabled = False
                    cmbKss(1).Enabled = False
                End If
                cboSS.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub lvw部门_DblClick()
    Dim lst As ListItem
    
    If lvw部门.SelectedItem Is Nothing Then Exit Sub
    If InStr(frmPresManage.mstrPrivs, "所有部门") = 0 And CheckDeptPermission(1, Mid(lvw部门.SelectedItem.Key, 2)) = False Then
        cmdRemove.Enabled = False
        Exit Sub
    End If
    For Each lst In lvw部门.ListItems
        If lst Is lvw部门.SelectedItem Then
            lvw部门.SelectedItem.SubItems(1) = "√"
        Else
            lst.SubItems(1) = ""
        End If
    Next
    cmdRemove.Enabled = lvw部门.SelectedItem.SubItems(1) = ""
End Sub

Private Sub lvw部门_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If InStr(frmPresManage.mstrPrivs, "所有部门") = 0 Then
        If Val(mstrID) = glngUserId Then
            cmdRemove.Enabled = False
            cmdAdd.Enabled = False
            Exit Sub
        End If
        If CheckDeptPermission(1, Mid(Item.Key, 2)) = False Then
            cmdRemove.Enabled = False
        Else
            cmdRemove.Enabled = Item.SubItems(1) = ""
        End If
    Else
        cmdRemove.Enabled = Item.SubItems(1) = ""
    End If
End Sub

Private Sub lvw部门_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        lvw部门_DblClick
    End If
End Sub

Private Sub pic外框_Paint()
    Dim r As RECT
    
    With r
        .Left = 0
        .Right = pic外框.ScaleWidth
        .Top = 0
        .Bottom = pic外框.ScaleHeight
    End With
    DrawEdge pic外框.hdc, r, BDR_RAISEDINNER, BF_RECT
End Sub

Private Sub tabMain_Click()
    Dim lngIndex As Long
    
    fra页(0).Visible = False
    fra页(1).Visible = False
    fra页(2).Visible = False
    
    lngIndex = Val(tabMain.SelectedItem.Index - 1)
    fra页(lngIndex).Visible = True
    fra页(lngIndex).ZOrder
    Select Case lngIndex
        Case 0
            If txtEdit(Text姓名).Enabled = True Then
                txtEdit(Text姓名).SetFocus
            End If
        Case 1
            If cmb编码(code民族).Enabled = True Then
                cmb编码(code民族).SetFocus
            End If
        Case 2
            If txt数(Number留学时间).Enabled = True Then
                txt数(Number留学时间).SetFocus
            End If
    End Select
End Sub

Private Sub tvw执业类别_LostFocus()
    tvw执业类别.Visible = False
End Sub

Private Sub tvw执业类别_NodeClick(ByVal Node As MSComctlLib.Node)
    With tvw执业类别
        If InStr(1, Node.Key, "C") > 0 Then
            txt编码.Text = Node.Text
            lblEdit(12).Tag = Node.Key
            .Visible = False
        End If
    End With
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
    Dim strDate As String
    
    strDate = zlCommFun.AddDate(txtDate(Index).Text)
    If Not IsDate(strDate) And strDate <> "" Then
        MsgBox "请按以下格式输入日期：2000-01-01。", vbInformation, gstrSysName
        Cancel = True
        zlControl.TxtSelAll txtDate(Index)
        Exit Sub
    End If
    If strDate <> "" Then
        txtDate(Index).Text = Format(CDate(strDate), "yyyy-MM-dd")
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text姓名 Then
        txtEdit(text简码).Text = zlStr.GetCodeByVB(txtEdit(Text姓名).Text)
    ElseIf Index = text别名 Then
        txt别名扩展.Text = txtEdit(text别名).Text
    End If
End Sub
Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
        Case Text姓名, Text个人简介, text别名, Text签名
            OS.OpenIme True
        Case Else
            OS.OpenIme False
    End Select
End Sub

'
Private Sub cmb编码_Click(Index As Integer)
    mblnChange = True
    
    Dim str编码 As String
    Dim lngCount As Long
    
'    If Index = code执业类别 Then
'        tvw执业类别.Visible = True
'        tvw执业类别.Top = cmb编码(7).Top
''        If cmb编码(code执业类别).Text = "" Then
''            lbl执业分类.Caption = ""
''        Else
''            str编码 = Mid(cmb编码(code执业类别), 1, InStr(cmb编码(code执业类别), ".") - 1)
''            lbl执业分类.Caption = mcol分类("K" & str编码)
''        End If
''
''        lst编码(code执业范围).Enabled = (lbl执业分类.Caption = "执业医师" Or lbl执业分类.Caption = "执业助理医师")
''        If lst编码(code执业范围).Enabled = False Then
''            For lngCount = 0 To lst编码(code执业范围).ListCount - 1
''                lst编码(code执业范围).Selected(lngCount) = False
''            Next
''        End If
'    End If
End Sub

Private Sub cmb编码_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngIdx As Long
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset
    
'    If SendMessage(cmb编码(Index).hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then OS.PressKey vbKeyF4
    If cmb编码(Index).Locked = True Then Exit Sub
    
    On Error GoTo ErrHandle
    If Index = code民族 Then
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
            If Chr(KeyAscii) <> mstr民族 Then
               gstrSQL = "select 编码,名称,简码 from 民族 where 简码 like [1]"
               Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "民族查询", UCase(Chr(KeyAscii)) & "%")
               If rsTemp.RecordCount > 0 Then
                  cmb编码(code民族).Text = rsTemp!编码 & "." & rsTemp!名称
                  mstr民族 = Chr(KeyAscii)
                  If Not rsTemp.EOF Then
                        Set mrs民族 = rsTemp
                        mrs民族.MoveNext
                  End If
               End If
            ElseIf Chr(KeyAscii) = mstr民族 And Not mrs民族.EOF Then '相同并且集合还没有到最后
                cmb编码(code民族).Text = mrs民族!编码 & "." & mrs民族!名称
                If Not mrs民族.EOF Then
                    mrs民族.MoveNext
                End If
            End If
        End If
    Else
        lngIdx = cbo.MatchIndex(cmb编码(Index).hwnd, KeyAscii)
        If lngIdx <> -2 Then cmb编码(Index).ListIndex = lngIdx
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmb编码_GotFocus(Index As Integer)
    OS.OpenIme False
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = Text个人简介 Then
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
        End If
    End If
    
    If Index = Text姓名 Then
        If KeyAscii = vbKeyReturn Then
            txtEdit(text别名).Text = txtEdit(Text姓名).Text
            txtEdit(Text签名).Text = txtEdit(Text姓名).Text
        End If
    End If
    
    If Index = Text姓名 Or Index = text别名 Or Index = Text签名 Or Index = text执业证号 Then
        If InStr("';/", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    ElseIf Index = Text编号 Or Index = text简码 Then
        If InStr("';", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
    
    If Index = text简码 Or Index = Text编号 Then
        If InStr(1, "0123456789qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM_", Chr(KeyAscii)) = 0 And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    End If
    
    If Index = Text移动电话 Then
        If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    End If
    
    If Index = Text顺序 Then
        If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    End If
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    If Index = Text顺序 Then
        If CheckOrder = True Then
            Cancel = True
        End If
    End If
End Sub

Private Function CheckOrder() As Boolean
    Dim rsTemp As Recordset
    Dim intOrder As Integer
    Dim i As Integer
    Dim str部门ID As String
    
    On Error GoTo ErrHandle
    
    If Val(txtEdit(Text顺序).Text) = 0 Then Exit Function
    CheckOrder = False
    
    For i = 1 To lvw部门.ListItems.Count
        str部门ID = str部门ID & "," & Mid(lvw部门.ListItems(i).Key, 2)
    Next
    
    If mstrID = "" Then '新增
        gstrSQL = "Select 1 From 人员表 A, 部门人员 B" & vbNewLine & _
                "Where a.Id = b.人员id" & vbNewLine & _
                "  And b.部门id In (Select /*+cardinality(a,10)*/ Column_Value" & vbNewLine & _
                "            From Table(f_Num2list([1])))" & vbNewLine & _
                "  And 顺序 = [2] And Rownum < 2"
    Else
        gstrSQL = "Select 1 From 人员表 A, 部门人员 B" & vbNewLine & _
                "Where a.Id = b.人员id" & vbNewLine & _
                "  And b.部门id In (Select /*+cardinality(a,10)*/ Column_Value" & vbNewLine & _
                "            From Table(f_Num2list([1])))" & vbNewLine & _
                "  And 顺序 = [2] And ID <> [3] And Rownum < 2"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询人员顺序", Mid(str部门ID, 2), Val(txtEdit(Text顺序).Text), Val(mstrID))
    
    If Not rsTemp.EOF Then
        gstrSQL = "Select Max(Nvl(a.顺序, 0)) As 最大顺序 From 人员表 A, 部门人员 B" & vbNewLine & _
                "Where a.Id = b.人员id" & vbNewLine & _
                "  And b.部门id In (Select /*+cardinality(a,10)*/ Column_Value" & vbNewLine & _
                "            From Table(f_Num2list([1])))"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询最大顺序", Mid(str部门ID, 2))
        
        MsgBox "所属部门下顺序为‘" & Val(txtEdit(Text顺序).Text) & "’的人员已存在，且最大顺序为‘" & rsTemp!最大顺序 & "’" & "，请重新输入人员顺序！", vbInformation, gstrSysName
        CheckOrder = True
    End If
    
    rsTemp.Close
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Txt编码_GotFocus()
    txt编码.SelStart = 0
    txt编码.SelLength = Len(txt编码.Text)
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    Else
        txt编码.Text = ""
    End If
End Sub

Private Sub txt别名扩展_Change()
    txtEdit(text别名).Text = txt别名扩展.Text
End Sub

Private Sub txt别名扩展_LostFocus()
    txt别名扩展.Visible = False
End Sub


Private Sub txt数_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt数(Index)
    OS.OpenIme False
End Sub

Private Sub txt数_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
'对于日期，不用输汉字
    OS.OpenIme False
    zlControl.TxtSelAll txtDate(Index)
End Sub

Private Sub InitEnv()
    Dim rsTemp As New ADODB.Recordset
    Dim strPrivs As String
    Dim strTemp As String
    Dim i As Integer
    
    Err = 0: On Error GoTo ErrHand:
    img照片.Stretch = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & frmPresManage.Name, "照片自动缩放", 1)) = 1)
    img照片.Left = pic镜框.ScaleLeft
    img照片.Top = pic镜框.ScaleTop
    If img照片.Stretch = True Then
        '不需要调整位置
        img照片.MousePointer = vbArrow
        img照片.Width = pic镜框.ScaleWidth
        img照片.Height = pic镜框.ScaleHeight
    Else
        img照片.MousePointer = vbSizeAll
    End If
    
    lbl说明.Caption = "    说明：人员可以隶属于多个部门，但缺省部门有且只能有一个。双击或使用空格键可使用指定部门成为缺省部门。"
    
    LoadComboFromSQL "select 编码,名称,缺省标志 from 性别 order by 编码", cmb编码(code性别)
    LoadComboFromSQL "select 编码,名称,缺省标志 from 民族 order by 编码", cmb编码(code民族)
    LoadComboFromSQL "select 编码,名称,缺省标志 from 学历 order by 编码", cmb编码(code学历)
    
   ' LoadComboFromSQL "select 编码,名称,0 as 缺省标志 from 专业技术职务 where 是否选择=1 order by 编码", cmb编码(code专业技术职务)
    
    LoadComboFromArray Array("1.正高", "2.副高", "3.中级", "4.助理/师级", "5.员/士", "9.待聘"), cmb编码(code聘任技术职务): cmb编码(code聘任技术职务).ListIndex = -1
    LoadComboFromArray Array("11.医疗(西医)", "12.中医", "13.口腔", "14.护理", "15.公共卫生", "16.药学", "17.检验" _
                              , "21.工程", "22.信息/计算机", "23.经济", "24.统计", "25.会计", "26.法律", "99.其他"), cmb编码(code所学专业): cmb编码(code所学专业).ListIndex = -1
    
    LoadComboFromArray Array("11.内科专业", "12.外科专业", "13.妇产科专业", "14.儿科专业", "15.眼耳鼻咽喉科专业", "16.皮肤病与性病专业", "17.精神卫生专业", "18.职业病专业", _
                             "19.医学影像和放射治疗专业", "20.医学检验、病理专业", "21.全科医学专业", "22.急救医学专业", "23.康复医学专业", "24.预防保健专业", "25.特种医学与军事医学专业", "26.计划生育技术服务专业", _
                             "31.口腔科专业", "41.公共卫生类别专业", "51.中医专业", "52.中西医结合专业", "53.蒙医专业", "54.藏医专业", "55.维医专业", "56.傣医专业"), lst编码(code执业范围)
                             
    LoadComboFromArray Array("1.世界卫生组织奖学金", "2.世界医学奖学金", "3.世界银行贷款", "4.教育部", "5.省市院校双边交流", "6.单位/自费公派", "7.自费", "9.其他"), lst编码(code留学渠道)
    LoadComboFromArray Array("1.住院医师规范化培训已合格", "2.正在接受住院医师规范化培训", "3.接受继续医学教育>=25学分", _
                             "4.接受继续医学教育<25学分", "5.其他岗位培训", "6.进修半年以上"), lst编码(code接受培训)
    LoadComboFromArray Array("1.自然科学基金", "2.国家科技攻关计划", "3.863计划", "4.973计划", _
                             "5.其他国家科技计划", "6.卫生部科技专项", "7.省级科技计划", "9.其他"), lst编码(code科研课题)
    
    gstrSQL = "select distinct 分类 from 执业类别"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Set mcol分类 = Nothing
'    cmb编码(code执业类别).Clear
    tvw执业类别.Nodes.Clear
    tvw执业类别.Nodes.Add , , "Root", "所有分类", "Root", "Root"
    tvw执业类别.Nodes("Root").Sorted = True
    Do Until rsTemp.EOF
        With tvw执业类别
            .Nodes.Add "Root", tvwChild, "K" & rsTemp!分类, rsTemp!分类, "Root"
'            cmb编码(code执业类别).AddItem rsTemp("编码") & "." & rsTemp("名称")
'            mcol分类.Add CStr(IIF(IsNull(rsTemp("分类")), "", rsTemp("分类"))), "K" & rsTemp("编码")
            rsTemp.MoveNext
        End With
    Loop
    gstrSQL = "select 编码,名称,分类 from 执业类别 order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询执行类别")
    Do Until rsTemp.EOF
        With tvw执业类别
            .Nodes.Add "K" & rsTemp!分类, tvwChild, "C" & rsTemp!编码, rsTemp!编码 & "." & rsTemp!名称, "Nature"
            rsTemp.MoveNext
        End With
    Loop
    tvw执业类别.Nodes.Item("Root").Expanded = True
    
    '管理职务
    gstrSQL = "Select 编码,名称 From 管理职务 Order By 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    cmb编码(code管理职务).Clear
    Do Until rsTemp.EOF
        cmb编码(code管理职务).AddItem rsTemp("编码") & "." & rsTemp("名称")
        rsTemp.MoveNext
    Loop
        
    '刘兴宏:2007/06/05:主要是将Combox控件改为了自定义的控件,主要是控制输入的编码\简码和名称
    
    'LoadComboFromSQL "select 编码,名称,0 as 缺省标志 from 专业技术职务 where 是否选择=1 order by 编码", cmb编码(code专业技术职务)
    
'    For i = 0 To lst编码(code人员性质).ListCount - 1
'        If lst编码(code人员性质).Selected(i) = True And (lst编码(code人员性质).List(i) = "医生" Or lst编码(code人员性质).List(i) = "护士") Then
'            strTemp = lst编码(code人员性质).List(i)
'        End If
'    Next
'
'    gstrSQL = " Select 编码 as ID,decode( substr(编码,1,2),编码,null ,substr(编码,1,2)) 上级ID,编码,名称,简码 From 专业技术职务 order by 编码"
'    zlDatabase.OpenRecordset rstemp, gstrSQL, Me.Caption
'    With rstemp
'        If .EOF Then
'            MsgBox "专业技术职务未安装,请找系统管理员！", vbInformation, gstrSysName
'            Exit Sub
'        End If
'
'        If cbo技术职务.FullCboData(rstemp, "", "编码,名称,简码", "编码|1000,名称|2000,简码|800", "", strTemp) = False Then
'            MsgBox "数据加载有误,请查看专业技术职是否正确！", vbInformation, gstrSysName
'            Exit Sub
'        End If
'    End With
    
    '住院、门诊抗菌药物授权
    For i = 0 To 1
        cmbKss(i).Clear
        cmbKss(i).AddItem ""
        cmbKss(i).AddItem "非限制使用"
        cmbKss(i).AddItem "限制使用"
        cmbKss(i).AddItem "特殊使用"
    Next
    If Int(glngSys / 100) = 1 Then
        strPrivs = ";" & zl9ComLib.GetPrivFunc(glngSys, 1024) & ";"
        If InStr(strPrivs, ";非限制使用;") > 0 And InStr(strPrivs, ";限制使用;") > 0 And InStr(strPrivs, ";特殊使用;") > 0 Then
            cmbKss(0).Visible = True
            cmbKss(1).Visible = True
            lblEdit(28).Visible = True
            lblEdit(30).Visible = True
            mbln抗菌药物 = True
        Else
            lblEdit(29).Top = lblEdit(16).Top
            cboSS.Top = cbo技术职务.Top
        End If
        
        lblEdit(29).Visible = True
        cboSS.Visible = True
    End If
    
    gstrSQL = "Select 名称 From 手术类型"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询手术类型")
    With cboSS
        .Clear
        .Enabled = False
        .AddItem ""
        Do While Not rsTemp.EOF
            .AddItem rsTemp!名称
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function LoadComboFromSQL(ByVal strSQL As String, cmbTemp As Variant, Optional ByVal blnID As Boolean = False) As Boolean
'本函数的功能是从数据库中读出列表值并装到下拉框中
    Dim rsTemp As New ADODB.Recordset
    Dim intCount As Long
    
    On Error GoTo ErrHandle
    rsTemp.CursorLocation = adUseClient
    rsTemp.CursorType = adOpenForwardOnly
    rsTemp.LockType = adLockReadOnly
'    Set rstemp.ActiveConnection = gcnOracle
    
'    Call SQLTest(App.ProductName, Me.Caption, strSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "LoadComboFromSQL")
'    Call SQLTest
    
    '下拉框数组
    If IsArray(cmbTemp) Then
        For intCount = LBound(cmbTemp) To UBound(cmbTemp)
            cmbTemp(intCount).Clear
            Do Until rsTemp.EOF
                If IsNull(rsTemp("编码")) Then
                    cmbTemp(intCount).AddItem rsTemp.AbsolutePosition & "." & rsTemp("名称")
                Else
                    cmbTemp(intCount).AddItem rsTemp("编码") & "." & rsTemp("名称")
                End If
                If blnID = True Then cmbTemp(intCount).ItemData(cmbTemp(intCount).NewIndex) = rsTemp("ID")
                If rsTemp("缺省标志") = 1 Then
                    cmbTemp(intCount).ListIndex = cmbTemp(intCount).NewIndex
                    cmbTemp(intCount).ItemData(cmbTemp(intCount).NewIndex) = 1
                End If
                rsTemp.MoveNext
            Loop
            rsTemp.MoveFirst
            If blnID = True Then cmbTemp(intCount).ListIndex = 0
        Next
         
    Else
        cmbTemp.Clear
        Do Until rsTemp.EOF
            If IsNull(rsTemp("编码")) Then
                cmbTemp.AddItem rsTemp.AbsolutePosition & "." & rsTemp("名称")
            Else
                cmbTemp.AddItem rsTemp("编码") & "." & rsTemp("名称")
            End If
            If blnID = True Then cmbTemp.ItemData(cmbTemp.NewIndex) = rsTemp("ID")
            If rsTemp("缺省标志") = 1 Then
                cmbTemp.ListIndex = cmbTemp.NewIndex
                cmbTemp.ItemData(cmbTemp.NewIndex) = 1
            End If
            rsTemp.MoveNext
        Loop
        If blnID = True Then cmbTemp.ListIndex = 0
    End If
    
    LoadComboFromSQL = True
    Exit Function
ErrHandle:
    LoadComboFromSQL = False
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadComboFromArray(ByVal varArray As Variant, cmbTemp As Variant) As Boolean
'本函数的功能是数组中读出列表值装到下拉框中
    
    Dim intArray As Long
    Dim intCount As Long
    
    On Error GoTo ErrHandle
    
    If IsArray(cmbTemp) Then
        For intCount = LBound(cmbTemp) To UBound(cmbTemp)
            cmbTemp(intCount).Clear
            For intArray = LBound(varArray) To UBound(varArray)
                cmbTemp(intCount).AddItem varArray(intArray)
            Next
            cmbTemp(intCount).ListIndex = 0
        Next
    Else
        cmbTemp.Clear
        For intArray = LBound(varArray) To UBound(varArray)
            cmbTemp.AddItem varArray(intArray)
        Next
        cmbTemp.ListIndex = 0
    End If
    LoadComboFromArray = True
    Exit Function
ErrHandle:
    LoadComboFromArray = False
End Function

Private Sub 显示空图片()
'在图片框中显示无图片信息
    If mbln照片 = False Then
        img照片.Picture = Nothing
        img照片.Tag = ""
        lbl图片说明 = "无照片"
    End If
End Sub
Private Function Save签名图片(ByVal lng人员id As Long, ByVal blnClear As Boolean) As Boolean
Dim rsTemp As New ADODB.Recordset, blnOk As Boolean
    
    On Error GoTo ErrHandle

    If blnClear Then
        gstrSQL = "Update 人员表 Set 签名图片 = Null Where ID = " & lng人员id
        gcnOracle.Execute gstrSQL
        blnOk = True
    Else
        blnOk = Sys.SaveLob(100, 15, lng人员id, pic签名图片.Tag)
        Kill pic签名图片.Tag
    End If
    
    Save签名图片 = blnOk
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt资格证书编号_Change()
    If Len(txt资格证书编号.Text) > 0 Then
        lblEdit(26).Caption = "资格证书编号(&N)" & "   共" & Len(txt资格证书编号.Text) & "位"
    Else
        lblEdit(26).Caption = "资格证书编号(&N)"
    End If
End Sub

Private Sub txt资格证书编号_KeyPress(KeyAscii As Integer)
    If InStr(" ';`/.,\][`-=~!@#$%^&*()_+{}:|<>?", Chr(KeyAscii)) > 0 Or KeyAscii = 34 Then KeyAscii = 0
End Sub

Private Sub CheckWorkNature()
'功能：检查是否有医生的工作性质，并设备抗菌药物权限

    Dim i As Integer
    Dim blnDuty As Boolean
    Dim strDuty As String
    
    If cmbKss(0).Visible = False Then Exit Sub
    
    strDuty = Mid(cbo技术职务.Text, InStr(cbo技术职务.Text, ".") + 1)
    blnDuty = strDuty = "医师" Or strDuty = "主治医师" Or strDuty = "副主任医师" Or strDuty = "主任医师"
    
    For i = 0 To lst编码(code人员性质).ListCount - 1
        If lst编码(code人员性质).List(i) = "医生" Then
            cmbKss(0).Enabled = lst编码(code人员性质).Selected(i) Or blnDuty
            cmbKss(1).Enabled = lst编码(code人员性质).Selected(i) Or blnDuty
            Exit For
        End If
    Next
End Sub

