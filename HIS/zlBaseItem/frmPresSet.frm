VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPresSet 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ա����"
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
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6555
      TabIndex        =   33
      Top             =   375
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6555
      TabIndex        =   34
      Top             =   975
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   6555
      TabIndex        =   35
      Top             =   7335
      Width           =   1100
   End
   Begin VB.Frame fraҳ 
      BorderStyle     =   0  'None
      Height          =   7200
      Index           =   0
      Left            =   105
      TabIndex        =   36
      Top             =   405
      Width           =   6255
      Begin MSComctlLib.TreeView tvwִҵ��� 
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
      Begin VB.TextBox txt������չ 
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
         Caption         =   "ɾ��"
         Enabled         =   0   'False
         Height          =   350
         Left            =   4395
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "����һ����������"
         Top             =   4305
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����"
         Height          =   350
         Left            =   3240
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "����һ����������"
         Top             =   4305
         Width           =   1100
      End
      Begin VB.ComboBox cmb���� 
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
      Begin VB.TextBox txt�ʸ�֤���� 
         Height          =   270
         Left            =   150
         MaxLength       =   30
         TabIndex        =   9
         Top             =   4905
         Width           =   2835
      End
      Begin VB.ListBox lst���� 
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
      Begin VB.ListBox lst���� 
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
      Begin VB.ComboBox cmb���� 
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
      Begin VB.CheckBox chk����Ȩ��־ 
         Caption         =   "����Ȩ(&J)"
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
      Begin VB.ComboBox cmb���� 
         Height          =   300
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   1785
      End
      Begin VB.TextBox txt���� 
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
      Begin VB.CommandButton cmd���� 
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
      Begin MSComctlLib.ListView lvw���� 
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
            Text            =   "��������"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "ȱʡ��־"
            Object.Width           =   1147
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpִҵʱ�� 
         Height          =   345
         Left            =   1200
         TabIndex        =   11
         ToolTipText     =   "ִָҵ��ʼʱ��"
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
      Begin zl9BaseItem.cboTree cbo����ְ�� 
         Height          =   300
         Left            =   1200
         TabIndex        =   13
         Top             =   6420
         Width           =   1785
         _ExtentX        =   3149
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
         FontName        =   "����"
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
      Begin VB.Label lbl˵�� 
         BackStyle       =   0  'Transparent
         Caption         =   "    ˵������Ա���������ڶ�����ţ���ȱʡ��������ֻ����һ����˫����ʹ�ÿո����ʹ��ָ�����ų�Ϊȱʡ���š�"
         Height          =   915
         Left            =   3240
         TabIndex        =   65
         Top             =   4800
         Width           =   2790
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "רҵְ��(&Z)"
         Height          =   180
         Index           =   16
         Left            =   150
         TabIndex        =   64
         Top             =   6465
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����ְ��(&M)"
         Height          =   180
         Index           =   15
         Left            =   150
         TabIndex        =   63
         Top             =   6090
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ִҵʱ��(&I)"
         Height          =   180
         Index           =   27
         Left            =   150
         TabIndex        =   62
         ToolTipText     =   "ִָҵ��ʼʱ��"
         Top             =   5715
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ִҵ֤��(&Y)"
         Height          =   180
         Index           =   25
         Left            =   150
         TabIndex        =   61
         Top             =   5325
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���(&U)"
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   60
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "��������(&R)"
         Height          =   180
         Left            =   3255
         TabIndex        =   59
         Top             =   120
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��������(&D)"
         Height          =   180
         Index           =   4
         Left            =   3255
         TabIndex        =   58
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&S)"
         Height          =   180
         Index           =   3
         Left            =   510
         TabIndex        =   57
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   2
         Left            =   510
         TabIndex        =   56
         Top             =   540
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ִҵ���(&K)"
         Height          =   180
         Index           =   12
         Left            =   150
         TabIndex        =   55
         Top             =   2700
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "Ƹ��ְ��(&T)"
         Height          =   180
         Index           =   18
         Left            =   150
         TabIndex        =   54
         Top             =   6840
         Width           =   990
      End
      Begin VB.Label lblִҵ���� 
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
         Caption         =   "Ժ��(&B)"
         Height          =   180
         Left            =   4560
         TabIndex        =   52
         Top             =   5715
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&A)"
         Height          =   180
         Index           =   13
         Left            =   510
         TabIndex        =   51
         Top             =   1980
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ǩ��(&Q)"
         Height          =   180
         Index           =   17
         Left            =   510
         TabIndex        =   50
         Top             =   2340
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "סԺ����ҩ��Ȩ��"
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
         Caption         =   "�Ա�(&X)"
         Height          =   180
         Index           =   6
         Left            =   510
         TabIndex        =   48
         Top             =   900
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ִҵ��Χ(&G)"
         Height          =   180
         Index           =   14
         Left            =   150
         TabIndex        =   47
         Top             =   3000
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�ʸ�֤����(&N)"
         Height          =   180
         Index           =   26
         Left            =   150
         TabIndex        =   46
         Top             =   4680
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����ȼ�"
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
         Caption         =   "���￹��ҩ��Ȩ��"
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
         Caption         =   "˳��(&R)"
         Height          =   180
         Index           =   32
         Left            =   510
         TabIndex        =   43
         Top             =   1620
         Width           =   630
      End
   End
   Begin VB.Frame fraҳ 
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
      Begin VB.ComboBox cmb���� 
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
      Begin VB.ComboBox cmb���� 
         Height          =   300
         Index           =   5
         Left            =   1305
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1425
         Width           =   2325
      End
      Begin VB.ComboBox cmb���� 
         Height          =   300
         Index           =   6
         Left            =   1305
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1815
         Width           =   2325
      End
      Begin VB.PictureBox pic��� 
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
         Begin VB.PictureBox pic���� 
            Height          =   1755
            Left            =   60
            ScaleHeight     =   1695
            ScaleWidth      =   1725
            TabIndex        =   74
            Top             =   60
            Width           =   1785
            Begin VB.Image img��Ƭ 
               Appearance      =   0  'Flat
               Height          =   1185
               Left            =   15
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1635
            End
         End
         Begin VB.Label lblͼƬ˵�� 
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
      Begin VB.CommandButton cmd��Ƭ 
         Caption         =   "�ļ�(&F)"
         Height          =   345
         Index           =   0
         Left            =   3930
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2250
         Width           =   855
      End
      Begin VB.CommandButton cmd��Ƭ 
         Caption         =   "���(&L)"
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
      Begin VB.PictureBox picǩ��ͼƬ 
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
      Begin VB.CommandButton cmdǩ�� 
         Caption         =   "���(&N)"
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
      Begin VB.CommandButton cmdǩ�� 
         Caption         =   "�ļ�(&I)"
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
      Begin MSComDlg.CommonDialog cdl��Ƭ 
         Left            =   4290
         Top             =   1620
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���֤��(&G)"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   86
         Top             =   1065
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&K)"
         Height          =   180
         Index           =   7
         Left            =   630
         TabIndex        =   85
         Top             =   275
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����ʼ�(&M)"
         Height          =   180
         Index           =   9
         Left            =   240
         TabIndex        =   84
         Top             =   3435
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�μӹ���(&W)"
         Height          =   180
         Index           =   11
         Left            =   270
         TabIndex        =   83
         Top             =   2250
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ѧ��(&S)"
         Height          =   180
         Index           =   19
         Left            =   630
         TabIndex        =   82
         Top             =   1485
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ѧרҵ(&P)"
         Height          =   180
         Index           =   20
         Left            =   270
         TabIndex        =   81
         Top             =   1875
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����绰(&T)"
         Height          =   180
         Index           =   8
         Left            =   240
         TabIndex        =   80
         Top             =   2670
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���˼��(&D)"
         Height          =   180
         Index           =   10
         Left            =   270
         TabIndex        =   79
         Top             =   3930
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��������(&B)"
         Height          =   180
         Index           =   5
         Left            =   270
         TabIndex        =   78
         Top             =   660
         Width           =   990
      End
      Begin VB.Label lblǩ��˵�� 
         Caption         =   "ǩ��ͼƬ"
         Height          =   810
         Left            =   4800
         TabIndex        =   77
         Top             =   2925
         Width           =   210
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�ƶ��绰(&M)"
         Height          =   180
         Index           =   31
         Left            =   240
         TabIndex        =   76
         Top             =   3060
         Width           =   990
      End
   End
   Begin VB.Frame fraҳ 
      BorderStyle     =   0  'None
      Height          =   7200
      Index           =   2
      Left            =   105
      TabIndex        =   87
      Top             =   405
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txt�� 
         Height          =   300
         Index           =   0
         Left            =   1410
         MaxLength       =   2
         TabIndex        =   32
         Top             =   270
         Width           =   525
      End
      Begin VB.ListBox lst���� 
         Height          =   1950
         Index           =   9
         Left            =   360
         Style           =   1  'Checkbox
         TabIndex        =   90
         Top             =   930
         Width           =   2085
      End
      Begin VB.ListBox lst���� 
         Height          =   1530
         Index           =   10
         Left            =   2970
         Style           =   1  'Checkbox
         TabIndex        =   89
         Top             =   930
         Width           =   2895
      End
      Begin VB.ListBox lst���� 
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
         Caption         =   "������ѵ(&R)"
         Height          =   180
         Index           =   23
         Left            =   2970
         TabIndex        =   94
         Top             =   690
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ѧʱ��(&I)        ��"
         Height          =   180
         Index           =   21
         Left            =   360
         TabIndex        =   93
         Top             =   330
         Width           =   1890
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ѧ����(&U)"
         Height          =   180
         Index           =   22
         Left            =   360
         TabIndex        =   92
         Top             =   690
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���п���(&J)"
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
            Caption         =   "������Ϣ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��ϸ���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��������"
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
Private Enum const����
    code�Ա� = 0
    code���� = 1
    code����ְ�� = 2
    '���˺�:2007/06/05����,�������������
    'codeרҵ����ְ�� = 3
    codeƸ�μ���ְ�� = 4
    codeѧ�� = 5
    code��ѧרҵ = 6
    '���µ�ʹ��ListBox�ؼ�
    code��Ա���� = 7
    codeִҵ��Χ = 8
    code��ѧ���� = 9
    code������ѵ = 10
    code���п��� = 11
End Enum

Private Enum const��
    Number��ѧʱ�� = 0
End Enum

Private Enum const�ı�
    Text���֤�� = 0
    Text��� = 1
    Text���� = 2
    text���� = 3
    Text�绰 = 4
    Text�����ʼ� = 5
    Text���˼�� = 6
    text���� = 7
    Textǩ�� = 8
    textִҵ֤�� = 9
    Text�ƶ��绰 = 10
    Text˳�� = 11
End Enum

Private Enum const����
    Date�������� = 0
    Date�μӹ��� = 1
End Enum

Private mstrID As String             '��ǰ�༭����ԱID
Private mblnChange As Boolean        '�Ƿ�ı���
Private mbln��Ƭ As Boolean          '��ǰ��Ա�Ƿ���ͼƬ��Ϣ
Private mbln��Ƭ���� As Boolean      '����Ƭ��������ʱ��ΪTrue
Private mblnǩ��ͼ As Boolean        'ǩ��ͼ�Ƿ���ͼ
Private mblnǩ��ͼ���� As Boolean    'ǩ��ͼ����ʱ��ΪTrue
Private mblnLoad As Boolean          'ΪTRUE��ʾ��װ��
Private mcol���� As New Collection   '��ʾĳ��ִҵ���ķ���
Private msngStartX As Single, msngStartY As Single    '�ƶ�ǰ����λ��
Private mstrKssZY As String           '���봰�������Ȩ�޵�ֵ�������ж��Ƿ�ı�
Private mstrKssMZ As String
Private mstr���� As String
Private mrs���� As ADODB.Recordset   '��¼��ѯ���ж�����¼�ļ���
Private mblnClickְ�� As Boolean     'רҵְ���Ƿ񱻵��
Private mbln����ҩ�� As Boolean         '����ҩ���޸�Ȩ�ޣ�true-�����޸ģ� false-�������޸�
Private mblnPACSInterface As Boolean        '����Ӱ����Ϣϵͳ�ӿ�
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
        
        strSQL = "select ���,���� from zlnodelist"
        Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "վ���ѯ")
        With cmbStationNo
            .AddItem ""
            Do While Not rsRecord.EOF
                .AddItem rsRecord!��� & "-" & rsRecord!����
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

Private Sub cbo����ְ��_Change()
    Dim i As Long
    
    mblnClickְ�� = False
    If mstrID <> "" And cmbKss(0).ListIndex > 0 Then
        Call CheckWorkNature
        Exit Sub
    End If
    If mstrID <> "" And cmbKss(1).ListIndex > 0 Then
        Call CheckWorkNature
        Exit Sub
    End If
    
    For i = 0 To 1
        If Mid(cbo����ְ��.Text, InStr(cbo����ְ��.Text, ".") + 1) = "ҽʦ" Then
            cmbKss(i).Text = "������ʹ��"
            cmbKss(i).Enabled = True
        ElseIf Mid(cbo����ְ��.Text, InStr(cbo����ְ��.Text, ".") + 1) = "����ҽʦ" Then
            cmbKss(i).Text = "����ʹ��"
            cmbKss(i).Enabled = True
        ElseIf Mid(cbo����ְ��.Text, InStr(cbo����ְ��.Text, ".") + 1) = "������ҽʦ" Or Mid(cbo����ְ��.Text, InStr(cbo����ְ��.Text, ".") + 1) = "����ҽʦ" Then
            cmbKss(i).Text = "����ʹ��"
            cmbKss(i).Enabled = True
        Else
            Call CheckWorkNature
            cmbKss(i).ListIndex = 0
        End If
    Next
End Sub

Private Sub cbo����ְ��_DownClick()
    mblnClickְ�� = True
End Sub

Private Sub cbo����ְ��_LostFocus()
    mblnClickְ�� = False
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
    For lngCount = txt��.LBound To txt��.UBound
        txt��(lngCount).Text = ""
    Next
    For lngCount = cmb����.LBound To cmb����.UBound
        If lngCount <> 3 Then
            cmb����(lngCount).ListIndex = -1
            For lngList = 0 To cmb����(lngCount).ListCount - 1
                '����ȱʡֵ
                If cmb����(lngCount).ItemData(lngList) = 1 Then
                    cmb����(lngCount).ListIndex = lngList
                    Exit For
                End If
            Next
        End If
    Next
    For lngCount = lst����.LBound To lst����.UBound
        For lngList = 0 To lst����(lngCount).ListCount - 1
            lst����(lngCount).Selected(lngList) = False
        Next
    Next
    
    mbln��Ƭ = False:   mbln��Ƭ���� = False
    Call ��ʾ��ͼƬ
    mblnǩ��ͼ = False: mblnǩ��ͼ���� = False
    Set picǩ��ͼƬ.Picture = Nothing: picǩ��ͼƬ.Tag = "": picǩ��ͼƬ.Cls
    txtEdit(Text���).Text = Sys.MaxCode("��Ա��", "���", 6)
    mblnChange = False
End Sub

Private Sub cmdOK_Click()
    If IsValid() = False Then Exit Sub
    If Save��Ա() = False Then Exit Sub
    If mstrID <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    
    Call ClearContext
    tabMain.Tabs(1).Selected = True
    txtEdit(Text����).SetFocus
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim ctlError As Control
    
    Set ctlError = GetErrorObject
    
    If ctlError Is Nothing Then
        'û�м�鵽����
        IsValid = True
    Else
        '��ʾ������ؼ���
        lngCount = ctlError.Container.Index
        tabMain.Tabs(lngCount + 1).Selected = True
        ctlError.SetFocus
    End If
    
End Function

Private Function GetErrorObject() As Control
    Dim i As Integer
    Dim strTemp As String
    
    '����ı����ֶ�
    For i = txtEdit.LBound To txtEdit.UBound
        strTemp = Trim(txtEdit(i).Text)
        If zlCommFun.StrIsValid(strTemp, txtEdit(i).MaxLength) = False Then
            Set GetErrorObject = txtEdit(i)
            Exit Function
        End If
    Next
    '����������ֶ�
    For i = txt��.LBound To txt��.UBound
        If IntegerIsValid(Trim(txt��(i).Text), txt��(i).MaxLength) = False Then
            Set GetErrorObject = txt��(i)
            Exit Function
        End If
    Next
    
    '��������Ϳؼ�
    For i = txtDate.LBound To txtDate.UBound
        If txtDate(i).Text <> "" Then
            If CDate(txtDate(i)) > Date Then
                MsgBox "�������ڳ�����ǰ���ڡ�", vbInformation, gstrSysName
                Set GetErrorObject = txtDate(i)
                Exit Function
            End If
        End If
    Next
    If txtDate(Date�μӹ���) <> "" And txtDate(Date��������) <> "" Then
        If CDate(txtDate(Date�μӹ���)) <= CDate(txtDate(Date��������)) Then
            MsgBox "�μӹ�������������ڳ������ڡ�", vbInformation, gstrSysName
            Set GetErrorObject = txtDate(Date�μӹ���)
            Exit Function
        End If
    End If
    
    '����б��ؼ�
    For i = codeִҵ��Χ To code���п���
        If lst����(i).SelCount > 3 Then
            MsgBox "ѡ�����Ŀ���ܳ���3����", vbExclamation, gstrSysName
            Set GetErrorObject = lst����(i)
            Exit Function
        End If
    Next
        
    If Len(Trim(txtEdit(Text���).Text)) = 0 Then
        txtEdit(Text���).Text = ""
        MsgBox "��Ų���Ϊ�ա�", vbExclamation, gstrSysName
        Set GetErrorObject = txtEdit(Text���)
        Exit Function
    End If
    If Len(Trim(txtEdit(Text����).Text)) = 0 Then
        MsgBox "��������Ϊ�ա�", vbExclamation, gstrSysName
        txtEdit(Text����).Text = ""
        Set GetErrorObject = txtEdit(Text����)
        Exit Function
    End If
        
    
    '�����֤�Ž�����֤
    strTemp = txtEdit(Text���֤��)
    If strTemp <> "" Then
        '������������֤��
        If IntegerIsValid(Trim(Mid(strTemp, 1, Len(strTemp) - 1)), 17) = False Then
            Set GetErrorObject = txtEdit(Text���֤��)
            Exit Function
        End If
        
        Dim str�������� As String
        Dim lng�Ա� As Long
        
        If Len(strTemp) <> 15 And Len(strTemp) <> 18 Then
            Set GetErrorObject = txtEdit(Text���֤��)
            MsgBox "���֤���볤�Ȳ��ԡ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        If Len(strTemp) = 15 Then
            '��ʽ
            str�������� = Mid(strTemp, 7, 6)
            str�������� = zlCommFun.AddDate(str��������)
            
            lng�Ա� = Val(Right(strTemp, 1))
        Else
            '��ʽ
            str�������� = Mid(strTemp, 7, 8)
            str�������� = zlCommFun.AddDate(str��������)
            
            lng�Ա� = Val(Mid(strTemp, 17, 1))
        End If
        If Not IsDate(str��������) Then
            Set GetErrorObject = txtEdit(Text���֤��)
            MsgBox "���֤�����г���������Ϣ����ȷ��", vbInformation, gstrSysName
            Exit Function
        End If
        If Not IsDate(txtDate(Date��������).Text) Then
            Set GetErrorObject = txtDate(Date��������)
            MsgBox "��ȷ�ϸó��������Ƿ���ȷ��", vbInformation, gstrSysName
            Exit Function
        End If
        If CDate(str��������) <> CDate(txtDate(Date��������).Text) Then
            Set GetErrorObject = txtEdit(Text���֤��)
            MsgBox "���֤�����г���������Ϣ��������ڲ��ԡ�", vbInformation, gstrSysName
            Exit Function
        End If
        If (lng�Ա� Mod 2 = 1 And InStr(cmb����(code�Ա�).Text, "Ů") > 0) Or (lng�Ա� Mod 2 = 0 And InStr(cmb����(code�Ա�).Text, "��") > 0) Then
            Set GetErrorObject = txtEdit(Text���֤��)
            MsgBox "���֤�������Ա���Ϣ����ȷ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    'ҽ��������Ա����ѡ��ִҵ���
    For i = 0 To lst����(code��Ա����).ListCount - 1
        If lst����(code��Ա����).Selected(i) = True Then
            If lst����(code��Ա����).List(i) = "ҽ��" And txt����.Text = "" Then
                Set GetErrorObject = txt����
                MsgBox "����Ա����ҽ�����ʣ�����ѡ��ִҵ���", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Next
'    '���θöδ��룬�����м��
'    '����27371��27388 by lesfeng 2010-01-19 �����豸���������÷��� ���� �������ʷ���
'    '�豸����Ա�Լ���������Ա������Ա��������ѡ�����Ǿ����豸���ü��������ò������ʷ���
'    For i = 0 To lst����(code��Ա����).ListCount - 1
'        If lst����(code��Ա����).Selected(i) = True Then
'            strTemp = Trim(lst����(code��Ա����).List(i))
'            If strTemp = "�豸����Ա" Then
'                If DeptIsValid(strTemp, 1) Then
'                    Set GetErrorObject = lst����(code��Ա����)
'                    MsgBox "����Ա�������Ų����С��豸���á��������ʣ���ȡ���豸����Ա���á�", vbInformation, gstrSysName
'                    Exit Function
'                End If
'            End If
'            If strTemp = "��������Ա" Then
'                If DeptIsValid(strTemp, 2) Then
'                    Set GetErrorObject = lst����(code��Ա����)
'                    MsgBox "����Ա�������Ų����С��������á��������ʣ���ȡ����������Ա���á�", vbInformation, gstrSysName
'                    Exit Function
'                End If
'            End If
'        End If
'    Next
    
End Function

Private Function IntegerIsValid(ByVal strInput As String, ByVal lng���� As Long) As Boolean
    Dim sngTemp As Long
    
    If strInput = "" Then
        IntegerIsValid = True
        Exit Function
    End If
    If Not IsNumeric(strInput) Then
        MsgBox "������һ����ȷ����ֵ��", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(strInput) < 0 Then
        MsgBox "������һ��������", vbInformation, gstrSysName
        Exit Function
    End If
    '��ֵ̫������
    On Error Resume Next
    sngTemp = Fix(Val(strInput))
    If Err <> 0 Then
        Err.Clear
        If lng���� < 10 Then
            MsgBox "����ֵ����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    'Tag���Ա��������������ֵĳ���
    If lng���� < 10 Then
        If Len(CStr(sngTemp)) > lng���� Then
            MsgBox "����ֵ����", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        If Len(strInput) > lng���� Then
            MsgBox "����ֵ����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    IntegerIsValid = True
End Function

Private Function DeptIsValid(ByVal strInput As String, ByVal IntFlag As Integer) As Boolean
'--����27371��27388 by lesfeng 2010-01-19
    Dim i As Integer
    Dim str����ID As String
    
    DeptIsValid = True
    For i = 1 To lvw����.ListItems.Count
        str����ID = Mid(lvw����.ListItems(i).Key, 2)
        DeptIsValid = DeptSQLIsValid(str����ID, IntFlag)
        If Not DeptIsValid Then Exit Function
    Next
End Function

Private Function DeptSQLIsValid(ByVal strInput As String, ByVal IntFlag As Integer) As Boolean
'--����27371��27388 by lesfeng 2010-01-19
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim int���� As Integer
    
    
    DeptSQLIsValid = True
    On Error GoTo ErrHandle
    If IntFlag = 1 Then '10 �豸
        strTemp = " And A.����id = [1] And B.���� = '10' "
    Else '11 ����
        strTemp = " And A.����id = [1] And B.���� = '11' "
    End If
    
    strSQL = " Select Count(A.����id) As ���� " & _
             "   From ��������˵�� A, �������ʷ��� B " & _
             "  Where A.�������� = B.����" & strTemp
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strInput))
    
    If Not rsTemp.EOF Then
        int���� = IIF(IsNull(rsTemp!����), 0, rsTemp!����)
        If int���� > 0 Then
            DeptSQLIsValid = False
        End If
    End If
    rsTemp.Close
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Function Save��Ա() As Boolean
    Dim lng��Աid  As Long
    Dim str����ID As String, str�������� As String
    Dim str��Ա���� As String
    Dim i As Integer, int���� As Integer
    Dim nod As Node
    Dim lst As ListItem
    Dim strרҵ����ְ�� As String, blnTran As Boolean
    Dim strվ�� As String
    Dim strSQL As String
    Dim curDate As Date
    Dim str���� As String
    
    On Error GoTo ErrHandle
    
    If Trim(txtEdit(text����).Text) = "" Then
        txtEdit(text����).Text = txtEdit(Text����).Text
    End If
        
    '�����в�������һ������ѡ�е�Ϊ1
    For i = 1 To lvw����.ListItems.Count
        str����ID = str����ID & Mid(lvw����.ListItems(i).Key, 2) & ":"
        If lvw����.ListItems(i).SubItems(1) = "��" Then
            str����ID = str����ID & "1;"
            
            str�������� = Mid(lvw����.ListItems(i).Text, InStr(lvw����.ListItems(i).Text, "��") + 1)
        Else
            str����ID = str����ID & "0;"
        End If
    Next
    
    '������ѡ�еĹ�����������һ����
    For i = 0 To lst����(code��Ա����).ListCount - 1
        If lst����(code��Ա����).Selected(i) = True Then
            str��Ա���� = str��Ա���� & lst����(code��Ա����).List(i) & ";"
        End If
    Next
    
    strרҵ����ְ�� = cbo����ְ��.Text
    If strרҵ����ְ�� <> "" Then
        strרҵ����ְ�� = "'" & Mid(strרҵ����ְ��, InStr(1, strרҵ����ְ��, ".") + 1) & "'"
    Else
        strרҵ����ְ�� = "NULL"
    End If
    
    gcnOracle.BeginTrans: blnTran = True
    
    If cmbStationNo.Text = "" Then
        strվ�� = "Null"
    Else
        strվ�� = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
    End If
    If txt����.Text <> "" Then
        str���� = "'" & Mid(txt����.Text, 1, InStr(1, txt����.Text, ".") - 1) & "'"
    Else
        str���� = "Null"
    End If
    
    '��ʽ����
    If mstrID = "" Then       '����һ����¼
        lng��Աid = Sys.NextId("��Ա��")
        gstrSQL = "zl_��Ա��_����(" & lng��Աid & _
            ",'" & txtEdit(Text���).Text & "','" & txtEdit(Text����).Text & "','" & txtEdit(text����).Text & "','" & _
            txtEdit(Text���֤��).Text & "'," & IIF(txtDate(Date��������).Text = "", "null", "to_date('" & txtDate(Date��������).Text & "','yyyy-MM-dd')") & "," & _
            GetTextFromCombo(cmb����(code�Ա�), True, ".") & "," & GetTextFromCombo(cmb����(code����), True, ".") & "," & _
            IIF(txtDate(Date�μӹ���).Text = "", "null", "to_date('" & txtDate(Date�μӹ���).Text & "','yyyy-MM-dd')") & ",'" & _
            txtEdit(Text�绰).Text & "','" & txtEdit(Text�����ʼ�).Text & "'," & _
            str���� & "," & GetTextFromList(lst����(codeִҵ��Χ)) & "," & _
            GetTextFromCombo(cmb����(code����ְ��), True, ".") & "," & strרҵ����ְ�� & "," & _
            GetTextFromCombo(cmb����(codeƸ�μ���ְ��), False, ".") & "," & GetTextFromCombo(cmb����(codeѧ��), True, ".") & "," & _
            GetTextFromCombo(cmb����(code��ѧרҵ), False, ".") & ",'" & txt��(Number��ѧʱ��).Text & "'," & _
            GetTextFromList(lst����(code��ѧ����)) & "," & GetTextFromList(lst����(code������ѵ)) & "," & _
            GetTextFromList(lst����(code���п���)) & ",'" & txtEdit(Text���˼��).Text & "','" & _
            str����ID & "','" & str��Ա���� & "','" & txtEdit(text����).Text & "'," & IIF(cmbStationNo.Visible = False Or cmbStationNo.Text = "", "Null", strվ��) & _
            ",'" & txtEdit(Textǩ��).Text & "','" & txtEdit(textִҵ֤��).Text & "','" & Me.txt�ʸ�֤����.Text & _
            IIF(Me.dtpִҵʱ��.value = Null, "',null", "',to_date('" & Me.dtpִҵʱ��.value & "','yyyy-MM-dd')") & _
            "," & Me.chk����Ȩ��־.value & _
            ",'" & Me.cboSS.Text & "','" & txtEdit(Text�ƶ��绰).Text & "'," & _
            IIF(txtEdit(Text˳��).Text = "", "Null", txtEdit(Text˳��).Text) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Else
        '�޸�
        lng��Աid = Val(mstrID)
        gstrSQL = "zl_��Ա��_�޸�(" & lng��Աid & _
            ",'" & txtEdit(Text���).Text & "','" & txtEdit(Text����).Text & "','" & txtEdit(text����).Text & "','" & _
            txtEdit(Text���֤��).Text & "'," & IIF(txtDate(Date��������).Text = "", "null", "to_date('" & txtDate(Date��������).Text & "','yyyy-MM-dd')") & "," & _
            GetTextFromCombo(cmb����(code�Ա�), True, ".") & "," & GetTextFromCombo(cmb����(code����), True, ".") & "," & _
            IIF(txtDate(Date�μӹ���).Text = "", "null", "to_date('" & txtDate(Date�μӹ���).Text & "','yyyy-MM-dd')") & ",'" & _
            txtEdit(Text�绰).Text & "','" & txtEdit(Text�����ʼ�).Text & "'," & _
            str���� & "," & GetTextFromList(lst����(codeִҵ��Χ)) & "," & _
            GetTextFromCombo(cmb����(code����ְ��), True, ".") & "," & strרҵ����ְ�� & "," & _
            GetTextFromCombo(cmb����(codeƸ�μ���ְ��), False, ".") & "," & GetTextFromCombo(cmb����(codeѧ��), True, ".") & "," & _
            GetTextFromCombo(cmb����(code��ѧרҵ), False, ".") & ",'" & txt��(Number��ѧʱ��).Text & "'," & _
            GetTextFromList(lst����(code��ѧ����)) & "," & GetTextFromList(lst����(code������ѵ)) & "," & _
            GetTextFromList(lst����(code���п���)) & ",'" & txtEdit(Text���˼��).Text & "','" & _
            str����ID & "','" & str��Ա���� & "','" & txtEdit(text����).Text & "'," & IIF(cmbStationNo.Visible = False Or cmbStationNo.Text = "", "Null", strվ��) & _
            ",'" & txtEdit(Textǩ��).Text & "','" & txtEdit(textִҵ֤��).Text & "','" & Me.txt�ʸ�֤����.Text & _
            IIF(Me.dtpִҵʱ��.value = Null, "',null", "',to_date('" & Me.dtpִҵʱ��.value & "','yyyy-MM-dd')") & _
            "," & Me.chk����Ȩ��־.value & _
            ",'" & Me.cboSS.Text & "','" & txtEdit(Text�ƶ��绰).Text & "'," & _
            IIF(txtEdit(Text˳��).Text = "", "Null", txtEdit(Text˳��).Text) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    '������Ա��Ƭ
    If mbln��Ƭ���� = True Then
        'ֻ�з����˸��Ĳ���Ҫ����
        'zlDatabase.ExecuteProcedure "delete from ��Ա��Ƭ where ��ԱID=" & lng��Աid, Me.Caption
        Call zlDatabase.ExecuteProcedure("zl_��Ա��Ƭ_Delete(" & lng��Աid & ")", Me.Caption)
        If mbln��Ƭ = True Then
            '����
            If Sys.SaveLob(100, 16, lng��Աid, img��Ƭ.Tag) = False Then
                gcnOracle.RollbackTrans
                MsgBox "��Ƭ����ʧ�ܡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    If mblnǩ��ͼ���� Then
        If Saveǩ��ͼƬ(lng��Աid, True) = False Then
            gcnOracle.RollbackTrans
            MsgBox "��Ƭ���ʧ�ܡ�", vbInformation, gstrSysName
            Exit Function
        End If
        If mblnǩ��ͼ Then
            If Saveǩ��ͼƬ(lng��Աid, False) = False Then
                gcnOracle.RollbackTrans
                MsgBox "��Ƭ����ʧ�ܡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    'סԺ����ҩ����Ȩ
    If mbln����ҩ�� = True Then
        If mstrKssZY <> cmbKss(0).Text Then
            If mstrID <> "" Or cmbKss(0).Text <> "" Then
                curDate = Sys.Currentdate
                strSQL = "Zl_��Ա����ҩ��Ȩ��_Update('" & _
                         lng��Աid & "'," & _
                         cmbKss(0).ListIndex & ",'" & _
                         gstrUserName & "'," & _
                         "to_date('" & curDate & "','YYYY-MM-DD HH24:MI:SS'), 1) "
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
        End If
    End If
    '���￹��ҩ����Ȩ
    If mbln����ҩ�� = True Then
        If mstrKssMZ <> cmbKss(1).Text Then
            If mstrID <> "" Or cmbKss(1).Text <> "" Then
                curDate = Sys.Currentdate
                strSQL = "Zl_��Ա����ҩ��Ȩ��_Update('" & _
                         lng��Աid & "'," & _
                         cmbKss(1).ListIndex & ",'" & _
                         gstrUserName & "'," & _
                         "to_date('" & curDate & "','YYYY-MM-DD HH24:MI:SS'), 2) "
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
        End If
    End If
    
    '����RIS�ӿڣ��޸���Ա��Ϣ����Ϊ��Ҫ��ͬ���û���Ϣ��������������Ա������׼�棬���ò�������������Ϊ����顱�Ĳ�����Ա���ӿڲ�����Ч��ǰ����
    If Int(glngSys / 100) = 1 And mblnPACSInterface = True And mstrID <> "" Then
        If IsCheckDeptPres(lng��Աid) Then
            If Not gobjRIS Is Nothing Then
                If gobjRIS.HISBasicDictTable(RISBaseItemType.Personnel, RISBaseItemOper.Modify, lng��Աid) <> 1 Then
                    gcnOracle.RollbackTrans
                    
                    '����ʱ��ʾ�ӿڴ�����Ϣ
                    If gobjRIS.LastErrorInfo <> "" Then
                        MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                    Else
                        MsgBox "����RIS�ӿڴ��󣬲��ܼ�����ǰ����������ϵͳ����Ա��ϵ", vbInformation, gstrSysName
                    End If
                    
                    Exit Function
                End If
            Else
                gcnOracle.RollbackTrans
                
                '�ӿڲ�����Чʱ��ֹ����ʾ
                MsgBox "RIS�ӿڴ���ʧ�ܣ����ܼ�����ǰ�����������ǽӿ��ļ���װ��ע�᲻����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                
                Exit Function
            End If
        End If
    End If
    
    gcnOracle.CommitTrans: blnTran = False
    
    frmPresManage.FillList frmPresManage.tvwMain_S.SelectedItem.Key
    Save��Ա = True
    Exit Function
ErrHandle:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function IsCheckDeptPres(ByVal lngPres As Long) As Boolean
    '�Ƿ��������Ա
    Dim rsData  As ADODB.Recordset
    
    gstrSQL = "Select 1 From ������Ա A, ��������˵�� B Where a.����id = b.����id And �������� = '���' And a.��Աid = [1] "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "IsCheckDeptPres", lngPres)
    
    IsCheckDeptPres = Not rsData.EOF
End Function
Public Function �༭��Ա(Optional strID As String = "", Optional ByVal str����ID As String = "") As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim varTemp As Variant
    Dim strTempFile As String
    Dim strSQL As String
    Dim strTemp As String
    Dim j As Integer
    Dim rs��Ա���� As New ADODB.Recordset
    Dim blnKind As Boolean
    
    rsTemp.CursorLocation = adUseClient
   
    img��Ƭ.ToolTipText = "�����С��" & img��Ƭ.Width & "��" & img��Ƭ.Height
    Call InitEnv
   
    Call IniStationNo
    
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    
    mbln��Ƭ = False:    mbln��Ƭ���� = False
    mblnǩ��ͼ = False:  mblnǩ��ͼ���� = False
    On Error GoTo ErrHandle
    mstrID = strID
    If strID <> "" Then
        Dim i As Integer, varValue As Variant
        gstrSQL = "Select  a.ID, a.���,b.���� as ִҵ������, a.����, a.����, a.���֤��, a.��������, a.�Ա�, a.����, a.��������, a.�칫�ҵ绰,a.�ƶ��绰, a.�����ʼ�,b.���� as ִҵ���, a.ִҵ��Χ, a.����ְ��, a.רҵ����ְ��,a.Ƹ�μ���ְ��, a.ѧ��, a.��ѧרҵ, a.��ѧʱ��, a.��ѧ����," & _
                          " a.������ѵ , a.���п���, a.���˼��, a.����, a.վ��, a.ǩ��,a.ִҵ֤��, a.�ʸ�֤���, a.ִҵ��ʼ����, a.����Ȩ��־,a.�����ȼ�,a.˳�� " & _
                   " From ��Ա�� a,ִҵ��� b" & _
                   " Where a.ID = [1] and a.ִҵ���=b.����(+) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
                
        txtEdit(Text���).Text = rsTemp("���")
        txtEdit(Text����).Text = rsTemp("����")
        txtEdit(text����).Text = IIF(IsNull(rsTemp("����")), "", rsTemp("����"))
        txtEdit(text����).Text = IIF(IsNull(rsTemp("����")), "", rsTemp("����"))
        txtEdit(Textǩ��).Text = IIF(IsNull(rsTemp("ǩ��")), "", rsTemp("ǩ��"))
        txtEdit(Text˳��).Text = IIF(IsNull(rsTemp("˳��")), "", rsTemp("˳��"))
        
        txtEdit(Text���֤��).Text = IIF(IsNull(rsTemp("���֤��")), "", rsTemp("���֤��"))
        txtEdit(Text�绰).Text = IIF(IsNull(rsTemp("�칫�ҵ绰")), "", rsTemp("�칫�ҵ绰"))
        txtEdit(Text�ƶ��绰).Text = IIF(IsNull(rsTemp("�ƶ��绰")), "", rsTemp("�ƶ��绰"))
        txtEdit(Text�����ʼ�).Text = IIF(IsNull(rsTemp("�����ʼ�")), "", rsTemp("�����ʼ�"))
        txtEdit(Text���˼��).Text = IIF(IsNull(rsTemp("���˼��")), "", rsTemp("���˼��"))
        
        txt��(Number��ѧʱ��).Text = IIF(IsNull(rsTemp("��ѧʱ��")), "", rsTemp("��ѧʱ��"))
        
        SetComboByText cmb����(code�Ա�), IIF(IsNull(rsTemp("�Ա�")), "", rsTemp("�Ա�")), True, "."
        SetComboByText cmb����(code����), IIF(IsNull(rsTemp("����")), "", rsTemp("����")), True, "."
        SetComboByText cmb����(codeѧ��), IIF(IsNull(rsTemp("ѧ��")), "", rsTemp("ѧ��")), True, "."
        
        
        'SetComboByText cmb����(codeרҵ����ְ��), IIF(IsNull(rsTemp("רҵ����ְ��")), "", rsTemp("רҵ����ְ��")), True, "."
        
        SetComboByText cmb����(code����ְ��), IIF(IsNull(rsTemp("����ְ��")), "", rsTemp("����ְ��")), True, "."
        SetComboByText cmb����(code��ѧרҵ), IIF(IsNull(rsTemp("��ѧרҵ")), "", rsTemp("��ѧרҵ")), False, "."
        
        txt����.Text = IIF(IsNull(rsTemp("ִҵ���")), "", rsTemp!ִҵ������ & "." & rsTemp("ִҵ���"))
        SetComboByText cmb����(codeƸ�μ���ְ��), IIF(IsNull(rsTemp("Ƹ�μ���ְ��")), "", rsTemp("Ƹ�μ���ְ��")), False, "."
        
        SetListByText lst����(codeִҵ��Χ), IIF(IsNull(rsTemp("ִҵ��Χ")), "", rsTemp("ִҵ��Χ"))
        SetListByText lst����(code��ѧ����), IIF(IsNull(rsTemp("��ѧ����")), "", rsTemp("��ѧ����"))
        SetListByText lst����(code������ѵ), IIF(IsNull(rsTemp("������ѵ")), "", rsTemp("������ѵ"))
        SetListByText lst����(code���п���), IIF(IsNull(rsTemp("���п���")), "", rsTemp("���п���"))
        
        txtDate(Date��������).Text = Format(rsTemp("��������"), "yyyy-MM-dd")
        txtDate(Date�μӹ���).Text = Format(rsTemp("��������"), "yyyy-MM-dd")
        
        txtEdit(textִҵ֤��).Text = IIF(IsNull(rsTemp!ִҵ֤��), "", rsTemp!ִҵ֤��)
        txt�ʸ�֤����.Text = IIF(IsNull(rsTemp!�ʸ�֤���), "", rsTemp!�ʸ�֤���)
        chk����Ȩ��־.value = IIF(IsNull(rsTemp!����Ȩ��־), 0, rsTemp!����Ȩ��־)
        dtpִҵʱ��.value = IIF(IsNull(rsTemp!ִҵ��ʼ����), Null, rsTemp!ִҵ��ʼ����)
        If NVL(rsTemp!�����ȼ�) <> "" Then
            cboSS.Text = NVL(rsTemp!�����ȼ�)
        Else
            cboSS.ListIndex = 0
        End If
        
        
        SetStationNo (IIF(IsNull(rsTemp("վ��")), "", rsTemp("վ��")))
        
        strTempFile = Sys.ReadLob(100, 15, Val(strID))
        If strTempFile <> "" Then
            picSign.Picture = LoadPicture(strTempFile)
            picǩ��ͼƬ.PaintPicture picSign.Picture, 0, 0, picǩ��ͼƬ.ScaleX(picǩ��ͼƬ.Width, vbTwips, vbPixels), picǩ��ͼƬ.ScaleY(picǩ��ͼƬ.Height, vbTwips, vbPixels)
            Kill strTempFile
        End If
        
        strTempFile = Trim(NVL(rsTemp!רҵ����ְ��))
        If strTempFile <> "" Then
            gstrSQL = "Select ���� From רҵ����ְ�� where ���� =[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTempFile)
            If rsTemp.EOF Then
            Else
                cbo����ְ��.SelectItemID = NVL(rsTemp!����)
                cbo����ְ��.Text = strTempFile
            End If
        End If
        
        '�������б�
'        If rsTemp.State = adStateOpen Then rsTemp.Close
        gstrSQL = "select C.����ID,b.���� as ����,b.���� as ���ű���,c.ȱʡ" & _
                    "  from ���ű� b,������Ա C " & _
                    " where C.����ID=B.ID and C.��Աid=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
        lvw����.ListItems.Clear
        Do Until rsTemp.EOF
            lvw����.ListItems.Add , "C" & rsTemp("����ID"), "��" & rsTemp("���ű���") & "��" & rsTemp("����")
            If rsTemp("ȱʡ") = 1 Then
                lvw����.ListItems("C" & rsTemp("����ID")).SubItems(1) = "��"
            End If
            rsTemp.MoveNext
        Loop
        
        '����ͼƬ
        strTempFile = Sys.ReadLobV2("��Ա��Ƭ", "��Ƭ", "��ԱID=[1]", "", Val(strID))
        img��Ƭ.Picture = LoadPicture(strTempFile)
        mbln��Ƭ = True
        lblͼƬ˵�� = GetPictureInfo(img��Ƭ.Picture)
        'ɾ������ʱ�ļ�
        If lblͼƬ˵�� <> "����Ƭ" Then
            Kill strTempFile
        End If
        
        'סԺ����ҩ��Ȩ��
        strSQL = "Select Max(����) ���� from ��Ա����ҩ��Ȩ�� where ��ԱID=[1] And ��¼״̬=1 and ���� = 1 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
        
        Call cbo.SetIndex(cmbKss(0).hwnd, Val(rsTemp!���� & ""))
        '���￹��ҩ��Ȩ��
        strSQL = "Select Max(����) ���� from ��Ա����ҩ��Ȩ�� where ��ԱID=[1] And ��¼״̬=1 and ���� = 2 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
        Call cbo.SetIndex(cmbKss(1).hwnd, Val(rsTemp!���� & ""))
        
    Else
        txtEdit(Text���).Text = Sys.MaxCode("��Ա��", "���", 6)
        
        '����ȱʡ�Ĳ��ű�
        gstrSQL = "select a.ID,a.���� ,a.����  from ���ű� A  where A.ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(str����ID))
                
        lvw����.ListItems.Clear
        lvw����.ListItems.Add , "C" & rsTemp("ID"), "��" & rsTemp("����") & "��" & rsTemp("����")
        lvw����.ListItems("C" & rsTemp("ID")).SubItems(1) = "��"
    End If
    
    '��¼��ʼ�Ŀ�����ֵ
    mstrKssZY = cmbKss(0).Text
    mstrKssMZ = cmbKss(1).Text
    
    If Not lvw����.SelectedItem Is Nothing Then
        If InStr(frmPresManage.mstrPrivs, "���в���") = 0 Then
            If Val(mstrID) = glngUserId Then
                cmdRemove.Enabled = False
                cmdAdd.Enabled = False
            Else
                If CheckDeptPermission(1, Mid(lvw����.SelectedItem.Key, 2)) = False Then
                    cmdRemove.Enabled = False
                Else
                    cmdRemove.Enabled = lvw����.SelectedItem.SubItems(1) = ""
                End If
            End If
        Else
            cmdRemove.Enabled = lvw����.SelectedItem.SubItems(1) = ""
        End If
    End If
    
    '�г�����Ա������
    If rsTemp.State = 1 Then rsTemp.Close
    If strID = "" Then
         gstrSQL = "select ����,null as ��Ա���� from ��Ա���ʷ��� order by ����"
    Else
         gstrSQL = "select A.����,B.��Ա���� from ��Ա���ʷ��� A,��Ա����˵�� B where A.����=B.��Ա����(+) and b.��ԱID(+)=[1] order by decode(��Ա����,null,1,0),A.����"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
    Dim lst As ListItem
    Do Until rsTemp.EOF
        lst����(code��Ա����).AddItem rsTemp("����")
        If Not IsNull(rsTemp("��Ա����")) Then lst����(code��Ա����).Selected(lst����(code��Ա����).NewIndex) = True
        rsTemp.MoveNext
    Loop
    
    For j = 0 To lst����(code��Ա����).ListCount - 1
        If lst����(code��Ա����).Selected(j) = True And (lst����(code��Ա����).List(j) = "ҽ��" Or lst����(code��Ա����).List(j) = "��ʿ") Then
            strTemp = lst����(code��Ա����).List(j)
        End If
    Next
    
    '����Ȩ���ж��Ƿ�����޸�
    If strID <> "" And InStr(frmPresManage.mstrPrivs, ";�޸�ʱ���޶���Ա����;") = 0 Then
        gstrSQL = "Select ��Ա���� From ��Ա����˵�� Where ��Աid = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ա���ʲ�ѯ", glngUserId)
        If rsTemp.RecordCount <> 0 Then
            For j = 0 To lst����(code��Ա����).ListCount - 1
                rsTemp.MoveFirst
                If lst����(code��Ա����).Selected(j) = True Then
                    Do While Not rsTemp.EOF
                        If lst����(code��Ա����).List(j) = rsTemp!��Ա���� Then
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
                fraҳ(0).Enabled = False
                fraҳ(1).Enabled = False
                fraҳ(2).Enabled = False
            Else
                For j = 1 To txtEdit.UBound
                    txtEdit(j).Enabled = False
                Next
                txt����.Enabled = False
                txt�ʸ�֤����.Enabled = False
                cmb����(0).Enabled = False
                cmb����(2).Enabled = False
                cmb����(4).Enabled = False
                dtpִҵʱ��.Enabled = False
                cbo����ְ��.Enabled = False
                cmdSelect.Enabled = False
                lst����(8).Enabled = False
                lst����(7).Enabled = False
                chk����Ȩ��־.Enabled = False
                cmbStationNo.Enabled = False
                cmbKss(0).Enabled = False
                cmbKss(1).Enabled = False
                cboSS.Enabled = False
                
                fraҳ(1).Enabled = False
                fraҳ(2).Enabled = False
            End If
        End If
    End If
    
    gstrSQL = " Select ���� as ID,decode( substr(����,1,2),����,null ,substr(����,1,2)) �ϼ�ID,����,����,���� From רҵ����ְ�� order by ����"
    zlDatabase.OpenRecordset rs��Ա����, gstrSQL, Me.Caption
    With rs��Ա����
        If .EOF Then
            MsgBox "רҵ����ְ��δ��װ,����ϵͳ����Ա��", vbInformation, gstrSysName
            Exit Function
        End If
        
        If cbo����ְ��.FullCboData(rs��Ա����, "", "����,����,����", "����|1000,����|2000,����|800", "", strTemp) = False Then
            MsgBox "���ݼ�������,��鿴רҵ����ְ�Ƿ���ȷ��", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    Call ��ʾ��ͼƬ
    '��ʼ�����
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
    Dim strID As String, str���� As String, str���� As String
    
    If InStr(frmPresManage.mstrPrivs, "���в���") = 0 Then
        blnRe = frmTreeSel.ShowTreePrivs(IIF(mstrID = "", glngUserId, mstrID), strID, str����, str����)
    Else
        gstrSQL = "select id,�ϼ�id,����,����,Upper(����) as ���� from ���ű� where ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')  start with �ϼ�id is null connect by prior id =�ϼ�id "
        blnRe = frmTreeSel.ShowTree(gstrSQL, strID, str����, str����, "", "��Ա��", "���в���", False)
    End If
    DoEvents
    If blnRe Then
        For lngCount = 1 To lvw����.ListItems.Count
            If Mid(lvw����.ListItems(lngCount).Key, 2) = strID Then
                MsgBox "��" & str���� & "���Ѿ��Ǹ���Ա�����������ˡ�", vbExclamation, gstrSysName
                Exit Sub
            End If
        Next
        lvw����.ListItems.Add , "C" & strID, "��" & str���� & "��" & str����
        lvw����.Refresh
        mblnChange = True
    End If
    
    If CheckOrder = True Then
        txtEdit(Text˳��).SetFocus
        Exit Sub
    End If
    
    Call lvw����_ItemClick(lvw����.SelectedItem)
End Sub

Private Sub cmdRemove_Click()
    If lvw����.SelectedItem Is Nothing Then Exit Sub
    
    lvw����.ListItems.Remove lvw����.SelectedItem.Key
    lvw����.ListItems(1).Selected = True
    
    Call lvw����_ItemClick(lvw����.SelectedItem)
End Sub

Private Sub cmdSelect_Click()
    With tvwִҵ���
        .Top = txt����.Top + txt����.Height
        .Left = txt����.Left
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub cmd����_Click()
    If txt������չ.Visible = False Then
        txt������չ.Visible = True
        txt������չ.SetFocus
    End If
End Sub

Private Sub cmdǩ��_Click(Index As Integer)
Dim sZoom As Single, lDesWidth As Long, lDesHeight As Long
    If Index = 1 Then '-���ͼƬ
        Set picǩ��ͼƬ.Picture = Nothing
        picǩ��ͼƬ.Tag = ""
        picǩ��ͼƬ.Cls
        mblnǩ��ͼ���� = True:   mblnǩ��ͼ = False
    Else
        With cdl��Ƭ
            .CancelError = True
            .Filter = "ͼƬ�ļ�(*.bmp,*.gif,*.jpg)|*.bmp;*.gif;*.jpg"
            
            On Error Resume Next
            .ShowOpen
            If Err <> 0 Then
                'ûѡ���ļ�
                Err.Clear
            Else
                picǩ��ͼƬ.Cls
                Set picǩ��ͼƬ.Picture = Nothing
                picǩ��ͼƬ.Picture = LoadPicture(.FileName)
                '�߶Ȳ�����50����,�����ݺ��,���ͼƬ��׺��.PIC
                If picǩ��ͼƬ.ScaleY(picǩ��ͼƬ.Picture.Height, vbHimetric, vbPixels) <= 50 Then
                    lDesWidth = picǩ��ͼƬ.ScaleX(picǩ��ͼƬ.Picture.Width, vbHimetric, vbPixels)
                    lDesHeight = picǩ��ͼƬ.ScaleY(picǩ��ͼƬ.Picture.Height, vbHimetric, vbPixels)
                Else
                    sZoom = picǩ��ͼƬ.ScaleY(picǩ��ͼƬ.Picture.Height, vbHimetric, vbPixels) / picǩ��ͼƬ.ScaleX(picǩ��ͼƬ.Picture.Width, vbHimetric, vbPixels)
                    lDesHeight = 50: lDesWidth = 50 / sZoom
                End If
                picǩ��ͼƬ.PaintPicture picǩ��ͼƬ.Picture, 0, 0, lDesWidth, lDesHeight
                picSign.Cls: Set picSign.Picture = Nothing
                picSign.Width = picSign.ScaleX(lDesWidth, vbPixels, vbTwips) + 45: picSign.Height = picSign.ScaleY(lDesHeight, vbPixels, vbTwips) + 45
                picSign.PaintPicture picǩ��ͼƬ.Picture, 0, 0, lDesWidth, lDesHeight
                SavePicture picSign.Image, Mid(.FileName, 1, Len(.FileName) - 3) & "PIC"
                If Err <> 0 Then
                    MsgBox "ͼƬ�ļ���Ч�����ļ������ڡ�", vbInformation, ""
                    Err.Clear
                    Exit Sub
                End If
                picǩ��ͼƬ.Tag = Mid(.FileName, 1, Len(.FileName) - 3) & "PIC"
                mblnǩ��ͼ���� = True:   mblnǩ��ͼ = True
            End If
        End With
    End If
End Sub

Private Sub img��Ƭ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
        msngStartY = Y
    End If
End Sub

Private Sub img��Ƭ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngLeft As Single
    Dim sngTop As Single
    
    '����״̬������
    If img��Ƭ.Stretch = True Then Exit Sub
    If Button = 1 Then
        '����������ܵ�
        sngLeft = img��Ƭ.Left + X - msngStartX
        sngTop = img��Ƭ.Top + Y - msngStartY
        
        '���ÿ��ܵ���߾�
        If img��Ƭ.Width < pic����.ScaleWidth Or sngLeft > pic����.ScaleLeft Then
            sngLeft = pic����.ScaleLeft
        Else
            If sngLeft + img��Ƭ.Width < pic����.ScaleWidth Then
                sngLeft = pic����.ScaleWidth - img��Ƭ.Width
            End If
        End If
        '���ÿ��ܵĶ��߾�
        If img��Ƭ.Height < pic����.ScaleHeight Or sngTop > pic����.ScaleTop Then
            sngTop = pic����.ScaleTop
        Else
            If sngTop + img��Ƭ.Height < pic����.ScaleHeight Then
                sngTop = pic����.ScaleHeight - img��Ƭ.Height
            End If
        End If
        img��Ƭ.Left = sngLeft
        img��Ƭ.Top = sngTop
    End If
End Sub

Private Sub cmd��Ƭ_Click(Index As Integer)
    Select Case Index
        Case 0 '�ļ�
            With cdl��Ƭ
                .CancelError = True
                .Filter = "ͼƬ�ļ�(*.bmp,*.gif,*.jpg)|*.bmp;*.gif;*.jpg"
                
                On Error Resume Next
                .ShowOpen
                If Err <> 0 Then
                    'ûѡ���ļ�
                    Err.Clear
                Else
                    img��Ƭ.Picture = LoadPicture(.FileName)
                    img��Ƭ.Left = pic����.ScaleLeft
                    img��Ƭ.Top = pic����.ScaleTop
                    
                    If Err <> 0 Then
                        MsgBox "ͼƬ�ļ���Ч�����ļ������ڡ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    lblͼƬ˵�� = GetPictureInfo(img��Ƭ.Picture)
                    img��Ƭ.Tag = .FileName
                    mbln��Ƭ = True
                    mbln��Ƭ���� = True
                End If
            End With
        Case 1 '���
            mbln��Ƭ = False
            mbln��Ƭ���� = True
            Call ��ʾ��ͼƬ
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
            If mblnClickְ�� = True Then
                mblnClickְ�� = False
            Else
                Unload Me
                Exit Sub
            End If
    End Select
    
    If KeyCode = vbKeyReturn Then
        If ActiveControl Is lvw���� Then
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
    
    tvwִҵ���.Visible = False
    mblnLoad = True
    For j = 0 To 1
        cmbKss(j).Enabled = False
        For i = 0 To lst����(code��Ա����).ListCount - 1
            If lst����(code��Ա����).List(i) = "ҽ��" Then
                If lst����(code��Ա����).Selected(i) Then
                    cmbKss(j).Enabled = True
                    Exit For
                End If
            End If
        Next
        If Mid(cbo����ְ��.Text, InStr(cbo����ְ��.Text, ".") + 1) = "ҽʦ" Or Mid(cbo����ְ��.Text, InStr(cbo����ְ��.Text, ".") + 1) = "����ҽʦ" Or Mid(cbo����ְ��.Text, InStr(cbo����ְ��.Text, ".") + 1) = "������ҽʦ" Or Mid(cbo����ְ��.Text, InStr(cbo����ְ��.Text, ".") + 1) = "����ҽʦ" Then
            cmbKss(j).Enabled = True
        End If
    Next
    
    strPrivs = ";" & zl9ComLib.GetPrivFunc(glngSys, 1002) & ";"
    blnSign = (Val(zlDatabase.GetPara("����ǩ����֤����", glngSys)) = 0)
    blnModify = InStr(strPrivs, ";�޸ĵ���ǩ��ͼƬ;") > 0
    If blnSign Then
        cmdǩ��(0).Visible = True
        cmdǩ��(1).Visible = True
    Else
        If blnModify Then
            cmdǩ��(0).Visible = True
            cmdǩ��(1).Visible = True
        Else
            cmdǩ��(0).Visible = False
            cmdǩ��(1).Visible = False
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
        fraҳ(intFra).Left = tabMain.ClientLeft
        fraҳ(intFra).Top = tabMain.ClientTop
        fraҳ(intFra).Height = tabMain.ClientHeight
        fraҳ(intFra).Width = tabMain.ClientWidth
        fraҳ(intFra).Visible = False
    Next
    
    With txt������չ
        .Visible = False
        .Top = lblEdit(13).Top + lblEdit(13).Height + 50
        .Left = lblEdit(13).Left
        .Width = txtEdit(7).Left - lblEdit(13).Left + txtEdit(7).Width + 50
        .Height = 2000
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
    If Cancel <> 1 Then Set mcol���� = Nothing
    mstr���� = ""
    Set mrs���� = Nothing
    
End Sub

Private Sub lst����_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub lst����_ItemCheck(Index As Integer, Item As Integer)
    If Index = code��Ա���� Then
        If lst����(code��Ա����).List(Item) = "ҽ��" Then
            If lst����(code��Ա����).Selected(Item) Then
                cmbKss(0).Enabled = True
                cmbKss(1).Enabled = True
                cboSS.Enabled = True
            Else
                If Mid(cbo����ְ��.Text, InStr(cbo����ְ��.Text, ".") + 1) = "ҽʦ" Or Mid(cbo����ְ��.Text, InStr(cbo����ְ��.Text, ".") + 1) = "����ҽʦ" Or Mid(cbo����ְ��.Text, InStr(cbo����ְ��.Text, ".") + 1) = "������ҽʦ" Or Mid(cbo����ְ��.Text, InStr(cbo����ְ��.Text, ".") + 1) = "����ҽʦ" Then
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

Private Sub lvw����_DblClick()
    Dim lst As ListItem
    
    If lvw����.SelectedItem Is Nothing Then Exit Sub
    If InStr(frmPresManage.mstrPrivs, "���в���") = 0 And CheckDeptPermission(1, Mid(lvw����.SelectedItem.Key, 2)) = False Then
        cmdRemove.Enabled = False
        Exit Sub
    End If
    For Each lst In lvw����.ListItems
        If lst Is lvw����.SelectedItem Then
            lvw����.SelectedItem.SubItems(1) = "��"
        Else
            lst.SubItems(1) = ""
        End If
    Next
    cmdRemove.Enabled = lvw����.SelectedItem.SubItems(1) = ""
End Sub

Private Sub lvw����_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If InStr(frmPresManage.mstrPrivs, "���в���") = 0 Then
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

Private Sub lvw����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        lvw����_DblClick
    End If
End Sub

Private Sub pic���_Paint()
    Dim r As RECT
    
    With r
        .Left = 0
        .Right = pic���.ScaleWidth
        .Top = 0
        .Bottom = pic���.ScaleHeight
    End With
    DrawEdge pic���.hdc, r, BDR_RAISEDINNER, BF_RECT
End Sub

Private Sub tabMain_Click()
    Dim lngIndex As Long
    
    fraҳ(0).Visible = False
    fraҳ(1).Visible = False
    fraҳ(2).Visible = False
    
    lngIndex = Val(tabMain.SelectedItem.Index - 1)
    fraҳ(lngIndex).Visible = True
    fraҳ(lngIndex).ZOrder
    Select Case lngIndex
        Case 0
            If txtEdit(Text����).Enabled = True Then
                txtEdit(Text����).SetFocus
            End If
        Case 1
            If cmb����(code����).Enabled = True Then
                cmb����(code����).SetFocus
            End If
        Case 2
            If txt��(Number��ѧʱ��).Enabled = True Then
                txt��(Number��ѧʱ��).SetFocus
            End If
    End Select
End Sub

Private Sub tvwִҵ���_LostFocus()
    tvwִҵ���.Visible = False
End Sub

Private Sub tvwִҵ���_NodeClick(ByVal Node As MSComctlLib.Node)
    With tvwִҵ���
        If InStr(1, Node.Key, "C") > 0 Then
            txt����.Text = Node.Text
            lblEdit(12).Tag = Node.Key
            .Visible = False
        End If
    End With
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
    Dim strDate As String
    
    strDate = zlCommFun.AddDate(txtDate(Index).Text)
    If Not IsDate(strDate) And strDate <> "" Then
        MsgBox "�밴���¸�ʽ�������ڣ�2000-01-01��", vbInformation, gstrSysName
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
    If Index = Text���� Then
        txtEdit(text����).Text = zlStr.GetCodeByVB(txtEdit(Text����).Text)
    ElseIf Index = text���� Then
        txt������չ.Text = txtEdit(text����).Text
    End If
End Sub
Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
        Case Text����, Text���˼��, text����, Textǩ��
            OS.OpenIme True
        Case Else
            OS.OpenIme False
    End Select
End Sub

'
Private Sub cmb����_Click(Index As Integer)
    mblnChange = True
    
    Dim str���� As String
    Dim lngCount As Long
    
'    If Index = codeִҵ��� Then
'        tvwִҵ���.Visible = True
'        tvwִҵ���.Top = cmb����(7).Top
''        If cmb����(codeִҵ���).Text = "" Then
''            lblִҵ����.Caption = ""
''        Else
''            str���� = Mid(cmb����(codeִҵ���), 1, InStr(cmb����(codeִҵ���), ".") - 1)
''            lblִҵ����.Caption = mcol����("K" & str����)
''        End If
''
''        lst����(codeִҵ��Χ).Enabled = (lblִҵ����.Caption = "ִҵҽʦ" Or lblִҵ����.Caption = "ִҵ����ҽʦ")
''        If lst����(codeִҵ��Χ).Enabled = False Then
''            For lngCount = 0 To lst����(codeִҵ��Χ).ListCount - 1
''                lst����(codeִҵ��Χ).Selected(lngCount) = False
''            Next
''        End If
'    End If
End Sub

Private Sub cmb����_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngIdx As Long
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset
    
'    If SendMessage(cmb����(Index).hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then OS.PressKey vbKeyF4
    If cmb����(Index).Locked = True Then Exit Sub
    
    On Error GoTo ErrHandle
    If Index = code���� Then
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
            If Chr(KeyAscii) <> mstr���� Then
               gstrSQL = "select ����,����,���� from ���� where ���� like [1]"
               Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����ѯ", UCase(Chr(KeyAscii)) & "%")
               If rsTemp.RecordCount > 0 Then
                  cmb����(code����).Text = rsTemp!���� & "." & rsTemp!����
                  mstr���� = Chr(KeyAscii)
                  If Not rsTemp.EOF Then
                        Set mrs���� = rsTemp
                        mrs����.MoveNext
                  End If
               End If
            ElseIf Chr(KeyAscii) = mstr���� And Not mrs����.EOF Then '��ͬ���Ҽ��ϻ�û�е����
                cmb����(code����).Text = mrs����!���� & "." & mrs����!����
                If Not mrs����.EOF Then
                    mrs����.MoveNext
                End If
            End If
        End If
    Else
        lngIdx = cbo.MatchIndex(cmb����(Index).hwnd, KeyAscii)
        If lngIdx <> -2 Then cmb����(Index).ListIndex = lngIdx
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmb����_GotFocus(Index As Integer)
    OS.OpenIme False
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = Text���˼�� Then
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
        End If
    End If
    
    If Index = Text���� Then
        If KeyAscii = vbKeyReturn Then
            txtEdit(text����).Text = txtEdit(Text����).Text
            txtEdit(Textǩ��).Text = txtEdit(Text����).Text
        End If
    End If
    
    If Index = Text���� Or Index = text���� Or Index = Textǩ�� Or Index = textִҵ֤�� Then
        If InStr("';/", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    ElseIf Index = Text��� Or Index = text���� Then
        If InStr("';", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
    
    If Index = text���� Or Index = Text��� Then
        If InStr(1, "0123456789qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM_", Chr(KeyAscii)) = 0 And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    End If
    
    If Index = Text�ƶ��绰 Then
        If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    End If
    
    If Index = Text˳�� Then
        If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 And KeyAscii <> 13 Then KeyAscii = 0
    End If
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    If Index = Text˳�� Then
        If CheckOrder = True Then
            Cancel = True
        End If
    End If
End Sub

Private Function CheckOrder() As Boolean
    Dim rsTemp As Recordset
    Dim intOrder As Integer
    Dim i As Integer
    Dim str����ID As String
    
    On Error GoTo ErrHandle
    
    If Val(txtEdit(Text˳��).Text) = 0 Then Exit Function
    CheckOrder = False
    
    For i = 1 To lvw����.ListItems.Count
        str����ID = str����ID & "," & Mid(lvw����.ListItems(i).Key, 2)
    Next
    
    If mstrID = "" Then '����
        gstrSQL = "Select 1 From ��Ա�� A, ������Ա B" & vbNewLine & _
                "Where a.Id = b.��Աid" & vbNewLine & _
                "  And b.����id In (Select /*+cardinality(a,10)*/ Column_Value" & vbNewLine & _
                "            From Table(f_Num2list([1])))" & vbNewLine & _
                "  And ˳�� = [2] And Rownum < 2"
    Else
        gstrSQL = "Select 1 From ��Ա�� A, ������Ա B" & vbNewLine & _
                "Where a.Id = b.��Աid" & vbNewLine & _
                "  And b.����id In (Select /*+cardinality(a,10)*/ Column_Value" & vbNewLine & _
                "            From Table(f_Num2list([1])))" & vbNewLine & _
                "  And ˳�� = [2] And ID <> [3] And Rownum < 2"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ��Ա˳��", Mid(str����ID, 2), Val(txtEdit(Text˳��).Text), Val(mstrID))
    
    If Not rsTemp.EOF Then
        gstrSQL = "Select Max(Nvl(a.˳��, 0)) As ���˳�� From ��Ա�� A, ������Ա B" & vbNewLine & _
                "Where a.Id = b.��Աid" & vbNewLine & _
                "  And b.����id In (Select /*+cardinality(a,10)*/ Column_Value" & vbNewLine & _
                "            From Table(f_Num2list([1])))"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ���˳��", Mid(str����ID, 2))
        
        MsgBox "����������˳��Ϊ��" & Val(txtEdit(Text˳��).Text) & "������Ա�Ѵ��ڣ������˳��Ϊ��" & rsTemp!���˳�� & "��" & "��������������Ա˳��", vbInformation, gstrSysName
        CheckOrder = True
    End If
    
    rsTemp.Close
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Txt����_GotFocus()
    txt����.SelStart = 0
    txt����.SelLength = Len(txt����.Text)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    Else
        txt����.Text = ""
    End If
End Sub

Private Sub txt������չ_Change()
    txtEdit(text����).Text = txt������չ.Text
End Sub

Private Sub txt������չ_LostFocus()
    txt������չ.Visible = False
End Sub


Private Sub txt��_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt��(Index)
    OS.OpenIme False
End Sub

Private Sub txt��_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
'�������ڣ������人��
    OS.OpenIme False
    zlControl.TxtSelAll txtDate(Index)
End Sub

Private Sub InitEnv()
    Dim rsTemp As New ADODB.Recordset
    Dim strPrivs As String
    Dim strTemp As String
    Dim i As Integer
    
    Err = 0: On Error GoTo ErrHand:
    img��Ƭ.Stretch = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & frmPresManage.Name, "��Ƭ�Զ�����", 1)) = 1)
    img��Ƭ.Left = pic����.ScaleLeft
    img��Ƭ.Top = pic����.ScaleTop
    If img��Ƭ.Stretch = True Then
        '����Ҫ����λ��
        img��Ƭ.MousePointer = vbArrow
        img��Ƭ.Width = pic����.ScaleWidth
        img��Ƭ.Height = pic����.ScaleHeight
    Else
        img��Ƭ.MousePointer = vbSizeAll
    End If
    
    lbl˵��.Caption = "    ˵������Ա���������ڶ�����ţ���ȱʡ��������ֻ����һ����˫����ʹ�ÿո����ʹ��ָ�����ų�Ϊȱʡ���š�"
    
    LoadComboFromSQL "select ����,����,ȱʡ��־ from �Ա� order by ����", cmb����(code�Ա�)
    LoadComboFromSQL "select ����,����,ȱʡ��־ from ���� order by ����", cmb����(code����)
    LoadComboFromSQL "select ����,����,ȱʡ��־ from ѧ�� order by ����", cmb����(codeѧ��)
    
   ' LoadComboFromSQL "select ����,����,0 as ȱʡ��־ from רҵ����ְ�� where �Ƿ�ѡ��=1 order by ����", cmb����(codeרҵ����ְ��)
    
    LoadComboFromArray Array("1.����", "2.����", "3.�м�", "4.����/ʦ��", "5.Ա/ʿ", "9.��Ƹ"), cmb����(codeƸ�μ���ְ��): cmb����(codeƸ�μ���ְ��).ListIndex = -1
    LoadComboFromArray Array("11.ҽ��(��ҽ)", "12.��ҽ", "13.��ǻ", "14.����", "15.��������", "16.ҩѧ", "17.����" _
                              , "21.����", "22.��Ϣ/�����", "23.����", "24.ͳ��", "25.���", "26.����", "99.����"), cmb����(code��ѧרҵ): cmb����(code��ѧרҵ).ListIndex = -1
    
    LoadComboFromArray Array("11.�ڿ�רҵ", "12.���רҵ", "13.������רҵ", "14.����רҵ", "15.�۶����ʺ��רҵ", "16.Ƥ�������Բ�רҵ", "17.��������רҵ", "18.ְҵ��רҵ", _
                             "19.ҽѧӰ��ͷ�������רҵ", "20.ҽѧ���顢����רҵ", "21.ȫ��ҽѧרҵ", "22.����ҽѧרҵ", "23.����ҽѧרҵ", "24.Ԥ������רҵ", "25.����ҽѧ�����ҽѧרҵ", "26.�ƻ�������������רҵ", _
                             "31.��ǻ��רҵ", "41.�����������רҵ", "51.��ҽרҵ", "52.����ҽ���רҵ", "53.��ҽרҵ", "54.��ҽרҵ", "55.άҽרҵ", "56.��ҽרҵ"), lst����(codeִҵ��Χ)
                             
    LoadComboFromArray Array("1.����������֯��ѧ��", "2.����ҽѧ��ѧ��", "3.�������д���", "4.������", "5.ʡ��ԺУ˫�߽���", "6.��λ/�Էѹ���", "7.�Է�", "9.����"), lst����(code��ѧ����)
    LoadComboFromArray Array("1.סԺҽʦ�淶����ѵ�Ѻϸ�", "2.���ڽ���סԺҽʦ�淶����ѵ", "3.���ܼ���ҽѧ����>=25ѧ��", _
                             "4.���ܼ���ҽѧ����<25ѧ��", "5.������λ��ѵ", "6.���ް�������"), lst����(code������ѵ)
    LoadComboFromArray Array("1.��Ȼ��ѧ����", "2.���ҿƼ����ؼƻ�", "3.863�ƻ�", "4.973�ƻ�", _
                             "5.�������ҿƼ��ƻ�", "6.�������Ƽ�ר��", "7.ʡ���Ƽ��ƻ�", "9.����"), lst����(code���п���)
    
    gstrSQL = "select distinct ���� from ִҵ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Set mcol���� = Nothing
'    cmb����(codeִҵ���).Clear
    tvwִҵ���.Nodes.Clear
    tvwִҵ���.Nodes.Add , , "Root", "���з���", "Root", "Root"
    tvwִҵ���.Nodes("Root").Sorted = True
    Do Until rsTemp.EOF
        With tvwִҵ���
            .Nodes.Add "Root", tvwChild, "K" & rsTemp!����, rsTemp!����, "Root"
'            cmb����(codeִҵ���).AddItem rsTemp("����") & "." & rsTemp("����")
'            mcol����.Add CStr(IIF(IsNull(rsTemp("����")), "", rsTemp("����"))), "K" & rsTemp("����")
            rsTemp.MoveNext
        End With
    Loop
    gstrSQL = "select ����,����,���� from ִҵ��� order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯִ�����")
    Do Until rsTemp.EOF
        With tvwִҵ���
            .Nodes.Add "K" & rsTemp!����, tvwChild, "C" & rsTemp!����, rsTemp!���� & "." & rsTemp!����, "Nature"
            rsTemp.MoveNext
        End With
    Loop
    tvwִҵ���.Nodes.Item("Root").Expanded = True
    
    '����ְ��
    gstrSQL = "Select ����,���� From ����ְ�� Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    cmb����(code����ְ��).Clear
    Do Until rsTemp.EOF
        cmb����(code����ְ��).AddItem rsTemp("����") & "." & rsTemp("����")
        rsTemp.MoveNext
    Loop
        
    '���˺�:2007/06/05:��Ҫ�ǽ�Combox�ؼ���Ϊ���Զ���Ŀؼ�,��Ҫ�ǿ�������ı���\���������
    
    'LoadComboFromSQL "select ����,����,0 as ȱʡ��־ from רҵ����ְ�� where �Ƿ�ѡ��=1 order by ����", cmb����(codeרҵ����ְ��)
    
'    For i = 0 To lst����(code��Ա����).ListCount - 1
'        If lst����(code��Ա����).Selected(i) = True And (lst����(code��Ա����).List(i) = "ҽ��" Or lst����(code��Ա����).List(i) = "��ʿ") Then
'            strTemp = lst����(code��Ա����).List(i)
'        End If
'    Next
'
'    gstrSQL = " Select ���� as ID,decode( substr(����,1,2),����,null ,substr(����,1,2)) �ϼ�ID,����,����,���� From רҵ����ְ�� order by ����"
'    zlDatabase.OpenRecordset rstemp, gstrSQL, Me.Caption
'    With rstemp
'        If .EOF Then
'            MsgBox "רҵ����ְ��δ��װ,����ϵͳ����Ա��", vbInformation, gstrSysName
'            Exit Sub
'        End If
'
'        If cbo����ְ��.FullCboData(rstemp, "", "����,����,����", "����|1000,����|2000,����|800", "", strTemp) = False Then
'            MsgBox "���ݼ�������,��鿴רҵ����ְ�Ƿ���ȷ��", vbInformation, gstrSysName
'            Exit Sub
'        End If
'    End With
    
    'סԺ�����￹��ҩ����Ȩ
    For i = 0 To 1
        cmbKss(i).Clear
        cmbKss(i).AddItem ""
        cmbKss(i).AddItem "������ʹ��"
        cmbKss(i).AddItem "����ʹ��"
        cmbKss(i).AddItem "����ʹ��"
    Next
    If Int(glngSys / 100) = 1 Then
        strPrivs = ";" & zl9ComLib.GetPrivFunc(glngSys, 1024) & ";"
        If InStr(strPrivs, ";������ʹ��;") > 0 And InStr(strPrivs, ";����ʹ��;") > 0 And InStr(strPrivs, ";����ʹ��;") > 0 Then
            cmbKss(0).Visible = True
            cmbKss(1).Visible = True
            lblEdit(28).Visible = True
            lblEdit(30).Visible = True
            mbln����ҩ�� = True
        Else
            lblEdit(29).Top = lblEdit(16).Top
            cboSS.Top = cbo����ְ��.Top
        End If
        
        lblEdit(29).Visible = True
        cboSS.Visible = True
    End If
    
    gstrSQL = "Select ���� From ��������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ��������")
    With cboSS
        .Clear
        .Enabled = False
        .AddItem ""
        Do While Not rsTemp.EOF
            .AddItem rsTemp!����
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
'�������Ĺ����Ǵ����ݿ��ж����б�ֵ��װ����������
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
    
    '����������
    If IsArray(cmbTemp) Then
        For intCount = LBound(cmbTemp) To UBound(cmbTemp)
            cmbTemp(intCount).Clear
            Do Until rsTemp.EOF
                If IsNull(rsTemp("����")) Then
                    cmbTemp(intCount).AddItem rsTemp.AbsolutePosition & "." & rsTemp("����")
                Else
                    cmbTemp(intCount).AddItem rsTemp("����") & "." & rsTemp("����")
                End If
                If blnID = True Then cmbTemp(intCount).ItemData(cmbTemp(intCount).NewIndex) = rsTemp("ID")
                If rsTemp("ȱʡ��־") = 1 Then
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
            If IsNull(rsTemp("����")) Then
                cmbTemp.AddItem rsTemp.AbsolutePosition & "." & rsTemp("����")
            Else
                cmbTemp.AddItem rsTemp("����") & "." & rsTemp("����")
            End If
            If blnID = True Then cmbTemp.ItemData(cmbTemp.NewIndex) = rsTemp("ID")
            If rsTemp("ȱʡ��־") = 1 Then
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
'�������Ĺ����������ж����б�ֵװ����������
    
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

Private Sub ��ʾ��ͼƬ()
'��ͼƬ������ʾ��ͼƬ��Ϣ
    If mbln��Ƭ = False Then
        img��Ƭ.Picture = Nothing
        img��Ƭ.Tag = ""
        lblͼƬ˵�� = "����Ƭ"
    End If
End Sub
Private Function Saveǩ��ͼƬ(ByVal lng��Աid As Long, ByVal blnClear As Boolean) As Boolean
Dim rsTemp As New ADODB.Recordset, blnOk As Boolean
    
    On Error GoTo ErrHandle

    If blnClear Then
        gstrSQL = "Update ��Ա�� Set ǩ��ͼƬ = Null Where ID = " & lng��Աid
        gcnOracle.Execute gstrSQL
        blnOk = True
    Else
        blnOk = Sys.SaveLob(100, 15, lng��Աid, picǩ��ͼƬ.Tag)
        Kill picǩ��ͼƬ.Tag
    End If
    
    Saveǩ��ͼƬ = blnOk
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt�ʸ�֤����_Change()
    If Len(txt�ʸ�֤����.Text) > 0 Then
        lblEdit(26).Caption = "�ʸ�֤����(&N)" & "   ��" & Len(txt�ʸ�֤����.Text) & "λ"
    Else
        lblEdit(26).Caption = "�ʸ�֤����(&N)"
    End If
End Sub

Private Sub txt�ʸ�֤����_KeyPress(KeyAscii As Integer)
    If InStr(" ';`/.,\][`-=~!@#$%^&*()_+{}:|<>?", Chr(KeyAscii)) > 0 Or KeyAscii = 34 Then KeyAscii = 0
End Sub

Private Sub CheckWorkNature()
'���ܣ�����Ƿ���ҽ���Ĺ������ʣ����豸����ҩ��Ȩ��

    Dim i As Integer
    Dim blnDuty As Boolean
    Dim strDuty As String
    
    If cmbKss(0).Visible = False Then Exit Sub
    
    strDuty = Mid(cbo����ְ��.Text, InStr(cbo����ְ��.Text, ".") + 1)
    blnDuty = strDuty = "ҽʦ" Or strDuty = "����ҽʦ" Or strDuty = "������ҽʦ" Or strDuty = "����ҽʦ"
    
    For i = 0 To lst����(code��Ա����).ListCount - 1
        If lst����(code��Ա����).List(i) = "ҽ��" Then
            cmbKss(0).Enabled = lst����(code��Ա����).Selected(i) Or blnDuty
            cmbKss(1).Enabled = lst����(code��Ա����).Selected(i) Or blnDuty
            Exit For
        End If
    Next
End Sub

