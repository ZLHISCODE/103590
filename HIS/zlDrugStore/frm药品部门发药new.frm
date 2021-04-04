VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Begin VB.Form frm部门发药管理New 
   Caption         =   "药品部门发药"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13440
   DrawStyle       =   1  'Dash
   Icon            =   "frm药品部门发药new.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frm药品部门发药new.frx":030A
   ScaleHeight     =   9015
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.Timer TimerReturn 
      Interval        =   10000
      Left            =   7920
      Top             =   240
   End
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   8640
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraColorStateSend 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2520
      TabIndex        =   72
      Top             =   5640
      Visible         =   0   'False
      Width           =   6705
      Begin VB.PictureBox picColorStateSend 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1200
         Picture         =   "frm药品部门发药new.frx":09F4
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   79
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picColorStateSend 
         Appearance      =   0  'Flat
         BackColor       =   &H00D7D7FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   5080
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   78
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox picColorStateSend 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   4080
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   77
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox picColorStateSend 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBDBDB&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3240
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   76
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox picColorStateSend 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFDDDD&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   2400
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   75
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox picColorStateSend 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   0
         Picture         =   "frm药品部门发药new.frx":7246
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   74
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picColorStateSend 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   6000
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   73
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblColorStateSend 
         AutoSize        =   -1  'True
         Caption         =   "抗菌药物"
         Height          =   180
         Index           =   4
         Left            =   1460
         TabIndex        =   86
         Top             =   30
         Width           =   720
      End
      Begin VB.Label lblColorStateSend 
         AutoSize        =   -1  'True
         Caption         =   "缺药"
         Height          =   180
         Index           =   3
         Left            =   5500
         TabIndex        =   85
         Top             =   30
         Width           =   360
      End
      Begin VB.Label lblColorStateSend 
         AutoSize        =   -1  'True
         Caption         =   "不处理"
         Height          =   180
         Index           =   2
         Left            =   4440
         TabIndex        =   84
         Top             =   30
         Width           =   540
      End
      Begin VB.Label lblColorStateSend 
         AutoSize        =   -1  'True
         Caption         =   "拒发"
         Height          =   180
         Index           =   1
         Left            =   3600
         TabIndex        =   83
         Top             =   30
         Width           =   360
      End
      Begin VB.Label lblColorStateSend 
         AutoSize        =   -1  'True
         Caption         =   "发药"
         Height          =   180
         Index           =   0
         Left            =   2790
         TabIndex        =   82
         Top             =   30
         Width           =   360
      End
      Begin VB.Label lblColorStateSend 
         AutoSize        =   -1  'True
         Caption         =   "高危药品"
         Height          =   180
         Index           =   5
         Left            =   260
         TabIndex        =   81
         Top             =   30
         Width           =   720
      End
      Begin VB.Label lblColorStateSend 
         AutoSize        =   -1  'True
         Caption         =   "新"
         Height          =   180
         Index           =   6
         Left            =   6420
         TabIndex        =   80
         Top             =   30
         Width           =   180
      End
   End
   Begin VB.PictureBox picCondition 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8535
      Left            =   120
      ScaleHeight     =   8535
      ScaleWidth      =   3615
      TabIndex        =   13
      Top             =   0
      Width           =   3615
      Begin VB.PictureBox picConOther 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3612
         Left            =   30
         ScaleHeight     =   3615
         ScaleWidth      =   3375
         TabIndex        =   47
         Top             =   3600
         Width           =   3375
         Begin VB.CheckBox chkWithNotAudited 
            BackColor       =   &H00FFFFFF&
            Caption         =   "包含销帐申请单据"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   0
            TabIndex        =   87
            Top             =   240
            Width           =   2172
         End
         Begin VB.CommandButton cmd药品剂型 
            Height          =   250
            Left            =   2985
            Picture         =   "frm药品部门发药new.frx":77D0
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   840
            Width           =   270
         End
         Begin VB.CommandButton cmd给药途径 
            Height          =   250
            Left            =   2985
            Picture         =   "frm药品部门发药new.frx":830A
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   480
            Width           =   270
         End
         Begin VB.TextBox txt药品剂型 
            Height          =   300
            Left            =   840
            TabIndex        =   65
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox txt给药途径 
            Height          =   300
            Left            =   840
            TabIndex        =   64
            Top             =   480
            Width           =   2415
         End
         Begin VB.Frame fraLineH2 
            Height          =   50
            Left            =   0
            TabIndex        =   63
            Top             =   120
            Width           =   3525
         End
         Begin VB.OptionButton opt范围 
            BackColor       =   &H00FFFFFF&
            Caption         =   "退药请求"
            Height          =   225
            Index           =   2
            Left            =   840
            TabIndex        =   62
            Top             =   2160
            Width           =   1125
         End
         Begin VB.OptionButton opt范围 
            BackColor       =   &H00FFFFFF&
            Caption         =   "发药请求"
            Height          =   225
            Index           =   1
            Left            =   2040
            TabIndex        =   61
            Top             =   1920
            Width           =   1125
         End
         Begin VB.OptionButton opt范围 
            BackColor       =   &H00FFFFFF&
            Caption         =   "所有请求"
            Height          =   225
            Index           =   0
            Left            =   840
            TabIndex        =   60
            Top             =   1920
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.ComboBox Cbo医嘱类型 
            Height          =   276
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   2400
            Width           =   2415
         End
         Begin VB.CheckBox chkType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "婴儿药品"
            Height          =   180
            Index           =   1
            Left            =   2160
            TabIndex        =   58
            Top             =   2760
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "病人药品"
            Height          =   180
            Index           =   0
            Left            =   840
            TabIndex        =   57
            Top             =   2760
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkDanger 
            BackColor       =   &H00FFFFFF&
            Caption         =   "仅提取高危药品"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   0
            TabIndex        =   56
            Top             =   3060
            Width           =   1695
         End
         Begin VB.CheckBox chkDangerType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "A类"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   55
            Top             =   3348
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox chkDangerType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "B类"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   1
            Left            =   960
            TabIndex        =   54
            Top             =   3348
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox chkDangerType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "C类"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   2
            Left            =   1680
            TabIndex        =   53
            Top             =   3348
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox chkToxicology 
            BackColor       =   &H00FFFFFF&
            Caption         =   "麻醉药"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   52
            Top             =   1440
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkToxicology 
            BackColor       =   &H00FFFFFF&
            Caption         =   "毒性药"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   1
            Left            =   1440
            TabIndex        =   51
            Top             =   1440
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkToxicology 
            BackColor       =   &H00FFFFFF&
            Caption         =   "精神I类"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   50
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkToxicology 
            BackColor       =   &H00FFFFFF&
            Caption         =   "精神II类"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   3
            Left            =   1440
            TabIndex        =   49
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkToxicologyType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "仅提取的毒理分类"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   0
            TabIndex        =   48
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lbl药品剂型 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "药品剂型"
            Height          =   180
            Left            =   0
            TabIndex        =   69
            Top             =   900
            Width           =   720
         End
         Begin VB.Label lbl给药途径 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "给药途径"
            Height          =   180
            Left            =   0
            TabIndex        =   68
            Top             =   540
            Width           =   720
         End
         Begin VB.Label lbl处理条件 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "处理范围"
            Height          =   180
            Left            =   0
            TabIndex        =   67
            Top             =   2040
            Width           =   720
         End
         Begin VB.Label Lbl医嘱类型 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "医嘱类型"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   66
            Top             =   2460
            Width           =   720
         End
      End
      Begin VB.PictureBox picDeptList 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   0
         ScaleHeight     =   1335
         ScaleWidth      =   3375
         TabIndex        =   38
         Top             =   6960
         Width           =   3375
         Begin VB.Frame fraLineH3 
            Height          =   50
            Left            =   0
            TabIndex        =   44
            Top             =   120
            Width           =   3525
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "刷新清单"
            Height          =   375
            Left            =   2040
            Picture         =   "frm药品部门发药new.frx":8E44
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdRefreshDept 
            Caption         =   "刷新科室"
            Height          =   375
            Left            =   1320
            TabIndex        =   42
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdListSel 
            Height          =   255
            Left            =   50
            Picture         =   "frm药品部门发药new.frx":93CE
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   300
            Width           =   255
         End
         Begin VB.CheckBox chkAll 
            BackColor       =   &H00FFFFFF&
            Caption         =   "全选"
            Enabled         =   0   'False
            Height          =   180
            Index           =   1
            Left            =   360
            TabIndex        =   40
            Top             =   337
            Width           =   735
         End
         Begin VB.CheckBox chkAll 
            BackColor       =   &H00FFFFFF&
            Caption         =   "全选"
            Enabled         =   0   'False
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   39
            Top             =   337
            Width           =   735
         End
         Begin MSComctlLib.TreeView tvwList 
            Height          =   735
            Index           =   0
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1296
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   476
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            ImageList       =   "imgTvw"
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.TreeView tvwList 
            Height          =   735
            Index           =   1
            Left            =   120
            TabIndex        =   46
            Top             =   720
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1296
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   476
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            ImageList       =   "imgTvw"
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox picConMain 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   30
         ScaleHeight     =   3615
         ScaleWidth      =   3375
         TabIndex        =   14
         Top             =   0
         Width           =   3375
         Begin VB.Frame fraLineH1 
            Height          =   50
            Left            =   0
            TabIndex        =   30
            Top             =   480
            Width           =   3405
         End
         Begin VB.ComboBox cbo发药药房 
            ForeColor       =   &H00FF0000&
            Height          =   276
            Left            =   840
            TabIndex        =   29
            Text            =   "cbo发药药房"
            Top             =   120
            Width           =   2415
         End
         Begin VB.ComboBox cbo时间范围 
            Height          =   300
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtInput 
            Height          =   300
            Left            =   840
            TabIndex        =   27
            Top             =   1680
            Width           =   2415
         End
         Begin VB.CheckBox chkSend 
            BackColor       =   &H00FFFFFF&
            Caption         =   "离院带药"
            Height          =   180
            Index           =   1
            Left            =   2160
            TabIndex        =   26
            Top             =   2040
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkSend 
            BackColor       =   &H00FFFFFF&
            Caption         =   "自取药"
            Height          =   180
            Index           =   2
            Left            =   840
            TabIndex        =   25
            Top             =   2280
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkSend 
            BackColor       =   &H00FFFFFF&
            Caption         =   "院内用药"
            Height          =   180
            Index           =   0
            Left            =   840
            TabIndex        =   24
            Top             =   2040
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.PictureBox picSendType 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   240
            ScaleHeight     =   255
            ScaleWidth      =   2895
            TabIndex        =   22
            Top             =   2880
            Width           =   2895
            Begin VB.CheckBox chkSendType 
               BackColor       =   &H00FFFFFF&
               Caption         =   "自定义发药类型，动态增加"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   23
               Top             =   0
               Width           =   2535
            End
         End
         Begin VB.PictureBox picShowOther 
            BackColor       =   &H00FFEDDD&
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   0
            MouseIcon       =   "frm药品部门发药new.frx":9F08
            ScaleHeight     =   270
            ScaleWidth      =   2655
            TabIndex        =   19
            Tag             =   "0"
            Top             =   3240
            Width           =   2655
            Begin VB.PictureBox picUpOrDown 
               BackColor       =   &H00FFEDDD&
               BorderStyle     =   0  'None
               Height          =   270
               Left            =   2400
               Picture         =   "frm药品部门发药new.frx":A212
               ScaleHeight     =   270
               ScaleWidth      =   270
               TabIndex        =   20
               Top             =   0
               Width           =   270
            End
            Begin VB.Label lblComment 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEDDD&
               Caption         =   "显示其它条件"
               ForeColor       =   &H00FF0000&
               Height          =   180
               Left            =   0
               TabIndex        =   21
               Top             =   45
               Width           =   1080
            End
         End
         Begin VB.PictureBox picShowSendType 
            BackColor       =   &H00FFEDDD&
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   0
            MouseIcon       =   "frm药品部门发药new.frx":A554
            ScaleHeight     =   270
            ScaleWidth      =   2655
            TabIndex        =   16
            Tag             =   "0"
            Top             =   2520
            Width           =   2655
            Begin VB.PictureBox picUpOrDown1 
               BackColor       =   &H00FFEDDD&
               BorderStyle     =   0  'None
               Height          =   270
               Left            =   2400
               Picture         =   "frm药品部门发药new.frx":A85E
               ScaleHeight     =   270
               ScaleWidth      =   270
               TabIndex        =   17
               Top             =   0
               Width           =   270
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEDDD&
               Caption         =   "显示其它发药类型"
               ForeColor       =   &H00FF0000&
               Height          =   180
               Left            =   0
               TabIndex        =   18
               Top             =   45
               Width           =   1440
            End
         End
         Begin VB.CommandButton cmdIC 
            Caption         =   "读卡"
            Height          =   300
            Left            =   2640
            TabIndex        =   15
            Top             =   1680
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSComCtl2.DTPicker Dtp结束时间 
            Height          =   315
            Left            =   840
            TabIndex        =   31
            Top             =   1320
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   123011075
            CurrentDate     =   39998
         End
         Begin MSComCtl2.DTPicker Dtp开始时间 
            Height          =   300
            Left            =   840
            TabIndex        =   32
            Top             =   960
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   123011075
            CurrentDate     =   39998
         End
         Begin zlIDKind.IDKindNew IDKNType 
            Height          =   300
            Left            =   0
            TabIndex        =   88
            Top             =   1680
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   529
            ShowSortName    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   9
            FontName        =   "宋体"
            IDKind          =   -1
            ShowPropertySet =   -1  'True
            AllowAutoICCard =   -1  'True
            AllowAutoIDCard =   -1  'True
            BackColor       =   16777215
         End
         Begin VB.Label lbl发药药房 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "发药药房"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   0
            TabIndex        =   37
            Top             =   180
            Width           =   720
         End
         Begin VB.Label lbl时间范围 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "时间范围"
            Height          =   180
            Left            =   0
            TabIndex        =   36
            Top             =   660
            Width           =   720
         End
         Begin VB.Label lblTimeEnd 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "结束时间"
            Height          =   180
            Left            =   0
            TabIndex        =   35
            Top             =   1387
            Width           =   720
         End
         Begin VB.Label lblTimeBegin 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "开始时间"
            Height          =   180
            Left            =   0
            TabIndex        =   34
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label lbl发药类型 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "发药类型"
            Height          =   180
            Left            =   0
            TabIndex        =   33
            Top             =   2160
            Width           =   720
         End
      End
   End
   Begin VB.Frame fraColorStateReturn 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   4440
      Width           =   2880
      Begin VB.PictureBox picColorStateReturn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFDDDD&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   9
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox picColorStateReturn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFDDDD&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   920
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   8
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox picColorStateReturn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFDDDD&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1800
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblColorStateReturn 
         AutoSize        =   -1  'True
         Caption         =   "退药"
         Height          =   180
         Index           =   0
         Left            =   380
         TabIndex        =   12
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColorStateReturn 
         AutoSize        =   -1  'True
         Caption         =   "原始"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   11
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColorStateReturn 
         AutoSize        =   -1  'True
         Caption         =   "已发药"
         Height          =   180
         Index           =   2
         Left            =   2200
         TabIndex        =   10
         Top             =   37
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList imgPacker 
      Left            =   5520
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm药品部门发药new.frx":ABA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm药品部门发药new.frx":B13A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm药品部门发药new.frx":B6D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerAuto 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3840
      Top             =   240
   End
   Begin VB.PictureBox picDetail 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   4080
      ScaleHeight     =   1935
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   960
      Width           =   3015
      Begin VB.Frame fraLineV1 
         Height          =   2085
         Left            =   120
         TabIndex        =   1
         Top             =   -120
         Width           =   45
      End
      Begin XtremeSuiteControls.TabControl tbcDetail 
         Height          =   975
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   1720
         _StockProps     =   64
         Enabled         =   -1  'True
      End
   End
   Begin MSComctlLib.ImageList imgTvw 
      Left            =   6240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm药品部门发药new.frx":BC6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm药品部门发药new.frx":C208
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm药品部门发药new.frx":C7A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm药品部门发药new.frx":CD3C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLvwSel 
      Left            =   6840
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm药品部门发药new.frx":D2D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm药品部门发药new.frx":D5F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm药品部门发药new.frx":D90A
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm药品部门发药new.frx":DC5C
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lvw给药途径 
      Height          =   345
      Left            =   4320
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgLvwSel"
      SmallIcons      =   "imgLvwSel"
      ColHdrIcons     =   "imgLvwSel"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView Lvw药品剂型 
      Height          =   345
      Left            =   4320
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgLvwSel"
      SmallIcons      =   "imgLvwSel"
      ColHdrIcons     =   "imgLvwSel"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   8655
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "frm药品部门发药new.frx":DFAE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12091
            Key             =   "HINT"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "未处理的销帐数据0条   "
            TextSave        =   "未处理的销帐数据0条   "
            Key             =   "CHARGEOFF"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "分包机"
            TextSave        =   "分包机"
            Key             =   "PACKER"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   5040
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frm药品部门发药new.frx":E842
      Left            =   4440
      Top             =   360
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frm部门发药管理New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''弹出菜单
Private Const mconMenu_TypePopup = 3000                  '给药途径分类
Private mTypeCount As Integer                            '给药途径分类的数量

Private Const mconMenu_SortPopup = 6000                  '排序方式
Private Const mconMenu_SortPopup_ByName = 6001           '科室列表，病人排序方式：按姓名
Private Const mconMenu_SortPopup_ByBedNo = 6002          '按床位号

Public mblnEnter As Boolean                              '是否能进入

Private mblnStartPacker As Boolean                       '是否启用药品分包机接口
Private mblnPackerConnect As Boolean                     '分包机接口数据库是否连接
Private mlng药房ID As Long
Private mstr药房编码 As String

Private mstrCardType As String   '银行卡类别，格式：短名|全名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密);…
Private mintCardCount As Integer  '卡数量

Private mblnFreshDeptList As Boolean
Private mblnStart As Boolean
Private mblnIs配置中心 As Boolean
Private mstr病人id As String

Private mstrDeptNode As String      '所选药房对应的站点
Private mRsDept As Recordset
Private mblnCheck As Boolean        '检查同组发药是否需要检查
Private rstemp As Recordset
Private mrsApplyforcredit As Recordset      '用于记录存在销帐申请的单据
Private mblnCustomCheck As Boolean      '是否开启自定义审核功能
Private mstrCustomCheckName As String   '自定义审核功能的名称

Private mclsComLib As Object
Private mobjDrugMAC As Object       '发药接口部件
Private mobjPlugIn As Object             '外挂接口对象
Private mobjCISJOB As Object  '电子病案查阅对象

Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库

Private mintCostDigit As Integer            '成本价小数位数
Private mintPriceDigit As Integer           '售价小数位数
Private mintNumberDigit As Integer          '数量小数位数
Private mintMoneyDigit As Integer           '金额小数位数

'消息相关对象变量
Private WithEvents mobjMipModule As zl9ComLib.clsMipModule
Attribute mobjMipModule.VB_VarHelpID = -1
Private mrsReceiveMsg As ADODB.Recordset    '已收到的消息记录集
Private mdateBegin As Date

Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

''''常量
'列表类型
Private Enum mDeptType
    发药 = 0
    退药 = 1
End Enum

Private Enum mListType
    发药 = 0
    汇总 = 1
    缺药 = 2
    拒发 = 3
    退药 = 4
    销账 = 5
End Enum

'时间范围
Private Enum mTimeRange
    当天 = 0
    两天内 = 1
    三天内 = 2
    指定时间范围 = 3
End Enum

'录入信息类型
Private Enum mInputType
    住院号 = 1
    姓名 = 2
    床号 = 3
    NO = 4
    病人ID = 5
    领药号 = 6
    发药号 = 7
    领药部门 = 8
    IC卡 = 9
End Enum

'执行状态
Private Enum mState
    缺药 = 0
    发药 = 1
    拒发 = 2
    不处理 = 3
    拒发_恢复 = 4
    拒发_不处理 = 5
    退药 = 6
    退药_原始记录 = 7
    退药_发药记录 = 8
    退药_退药记录 = 9
    转出记录 = 10
End Enum

'给药途径、药品剂型选择
Private Enum mSel
    给药途径 = 0
    药品剂型 = 1
End Enum

'发药列表颜色
Private Enum mSendListColor
    SendState = 0
    RejectState = 1
    UnProcessState = 2
    ShortageState = 3
End Enum

'退药列表颜色
Private Enum mReturnListColor
    ReturnState = 0
    OriginalState = 1
    SendedState = 2
End Enum

'''变量

'默认的窗体大小
Private Const mcstlngWinNormalWidth As Long = 13275
Private Const mcstlngWinNormalHeight As Long = 8500

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private mlngMyWindow As Long

Private mdate上次刷新时间 As Date                       '记录上次刷新时系统时间

Private mstrPrivs As String                                 '权限串
Private mlngMode As Long                                    '模块号

Private mcur汇总发药号 As Currency

'查询提示
Private Type TYPE_FindWar
    blnNoAsk_Dept_Send As Boolean                      '查询时间过长时提示：是否下次不在提示，发药时
    blnNoAsk_Dept_Sended As Boolean                    '查询时间过长时提示：是否下次不在提示，退药时
    blnNoAsk_Rec As Boolean                            '查询明细记录过多时提示：是否下次不再提示
    blnProc_Dept_Send As Boolean                       '查询科室列表，是否继续，发药时
    blnProc_Dept_Sended As Boolean                     '查询科室列表，是否继续，退药时
    blnProc_Rec As Boolean                             '查询明细记录，是否继续
End Type
Private mFindWar As TYPE_FindWar

Private mfrmDetail As New frm部门发药清单

Private mblnExistOtherSendType As Boolean                   '是否有自定义的发药类型
Private mblnCard As Boolean                                 '是否刷就诊卡
Private mobjSquareCard As Object             '一卡通接口

Public BlnSetPara As Boolean                                '参数设置窗体是否确定后退出
Public BlnRefresh As Boolean                                '其他窗体是否处理了数据,是则刷新
Private mblnInput As Boolean                                '是否是通过录入病人信息方式来查找数据

Private mstr科室ID串 As String                              '领药部门ID
Private mstr科室名串 As String                              '领药部门名称

'常用数据集
Private mrsDeptList As ADODB.Recordset                      '根据部门列表实际勾选的情况用勾选的部门、NO等作为主要条件用于提取明细数据
Private mrsSendData As ADODB.Recordset                      '待发药品记录集
Private mrsReturnData As ADODB.Recordset                    '退药药品记录集
Private mrsChargeOff As New ADODB.Recordset                   '用于显示销帐申请记录
Private mrsChargeOffMain As New ADODB.Recordset               '用于销账
Private mrs给药途径 As ADODB.Recordset
Private mrs用品剂型 As ADODB.Recordset
Private mrsPASS As ADODB.Recordset                          'PASS用数据集

'医保接口
Private gclsInsure As New clsInsure
Private Type TYPE_MedicarePAR
    负数记帐 As Boolean
    记帐上传 As Boolean
    记帐完成后上传 As Boolean
    记帐作废上传 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

'权限
Private Type Type_Privs
    bln所有药房 As Boolean
    bln发药 As Boolean
    bln退药 As Boolean
    bln退其它药房的处方 As Boolean
    bln发退结帐处方 As Boolean
    bln发退出院病人处方 As Boolean
    bln缺药申领 As Boolean
    bln拒发 As Boolean
    bln合理用药监测 As Boolean
    bln医生查询 As Boolean
    bln参数设置 As Boolean
    bln退药销帐 As Boolean
    bln修改留存数量 As Boolean
    bln停止发药 As Boolean
    bln恢复发药 As Boolean
    bln查看已发药清单 As Boolean
    bln过滤时间 As Boolean
    bln电子病案查阅 As Boolean
End Type
Private mPrives As Type_Privs


'列表查询条件
Private Type Type_Condition
    '主要条件
    lng药房id As Long
    str开始时间 As String
    str结束时间 As String
    int发药类型 As Integer
    str其它发药类型 As String
        
    '录入信息，单选条件
    str住院号 As String
    str姓名 As String
    str床号 As String
    strNo As String
    lng病人ID As Long
    str就诊卡 As String
    str领药号 As String
    cur发药号 As Currency
    lng领药部门ID As Long
    strIC卡 As String
    
    '简要条件
    str给药途径 As String
    str药品剂型 As String
    int处理范围 As Integer
    int医嘱类型 As Integer
    int病人类型 As Integer
    
    
    '其它条件
    int操作模式 As Integer
    str记账人 As String
    bln显示退药待发单据 As Boolean
    bln显示所有退药单据 As Boolean
    bln显示领药退药人 As Boolean
    int显示退药树表模式 As Integer          '0-按科室、病人、NO组织；1-按发药号、科室、病人、NO组织
End Type
Private mcondition As Type_Condition

'使用到的参数（来自系统参数表或其它参数表或本机注册表）
Private Type Type_Params
    '参数表中的系统参数
    bln允许未审核处方发药 As Boolean
    bln门诊医嘱先作废后退药 As Boolean
    int金额保留位数 As Integer
    bln审核划价单 As Boolean
    int效期显示方式 As Integer          '0-以失效期显示，1-以有效期显示
    int药品名称显示 As Integer          '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
    bln启用审方     As Boolean          '是否启用处方审查系统
    
    '参数表中的其它参数
    intDays As Integer
    bln领药人签名 As Boolean
    bln缺药检查 As Boolean
    bln退药人签名 As Boolean
    int自动刷新未发药清单 As Integer
    bln药品储备 As Boolean
    bln汇总发药 As Boolean
    int操作模式 As Integer
    int医嘱类型 As Integer
    bln汇总显示 As Boolean
    str记帐人 As String
    str毒理分类 As String
    str价值分类 As String
    str高危分类 As String
    str高危发放 As String
    lng药房id As Long
    int自动打印 As Integer
    bln待发单据 As Boolean
    bln过程单据 As Boolean
    int查询发药天数 As Integer
    int查询退药天数 As Integer
    lng最大记录数 As Long
    bln审核出院销账申请 As Boolean
    int退药清单打印 As Integer
    intCheck As Integer
    int发药时审核医嘱 As Integer
    bln检查存储库房 As Integer
    bln检查销帐申请 As Integer
    
    '库存检查
    IntCheckStock As Integer
    
    '库房单位
    strUnit As String
    
    '启用合理用药PASS
    blnStarPass As Boolean
    
    '配置中心
    bln配制中心 As Boolean
    
    '科室来源
    strSourceDep As String
    
    '注册表参数
    int药品名称编码显示 As String       '0－药品编码与名称；1－药品编码；2－药品名称
    intFont As Integer                  '字体号
    StrFindStyle As String              '输入匹配
    int输入模式索引 As Integer
    int病人排序 As Integer                  '科室列表中，病人排序方式：1-按姓名；2-按床位
    blnOnlyShowDept As Boolean              '部门列表是否仅显示部门名称
    intShowDept As Integer                  '0-显示所有科室;1-显示临床科室;2-显示医技科室;3-显示病人病区
    blnShowReject As Boolean                '提取拒发药品：0-不提取拒发药品；1-提取拒发药品
    intAdviceType As Integer                '医嘱类型：0-包含所有单据;1-仅含长期医嘱;2-仅含临时医嘱;3-普通记帐单据;4-包含所有医嘱
    blnSort As Boolean                      '医嘱列表科室按医嘱最后发送时间排序
    int页签 As Integer                      '窗体退出时当前页签
    bln保持上一次页签 As Boolean
    
    int输入模式 As Integer
    
    '注册表参数：包装机相关
    int暂停传送 As Integer              '暂停发药时向包装机传送数据:0-传送;1-暂停传送
    str包装机单据 As String             '传送数据的类型，格式“00”，第1位表示“长嘱”，第2位表示“临嘱”；0－表示不包含；1－表示包含
    str包装机剂型 As String             '传送数据的剂型：剂型名称串，用“;”分隔，如果是“所有”则表示所有剂型
End Type
Private mParams As Type_Params

Private Sub DrugStoreWork_CustomCheck()
    '用于用户自定义的医嘱审核，调用zlPlugIn接口，传入发药数据，返回审核未通过数据，更新界面
    Dim rsSendData As ADODB.Recordset
    Dim str收发ids As String
    Dim str返回收发ids As String
    Dim strReserve As String
    
    '取发药数据集
    Set rsSendData = mfrmDetail.GetSendRecord
    
    If rsSendData Is Nothing Then Exit Sub
    
    If rsSendData.RecordCount = 0 Then Exit Sub
    
    With rsSendData
        .Filter = "执行状态=" & mState.发药
        
        If .RecordCount = 0 Then Exit Sub
        
        Do While Not .EOF
            str收发ids = IIf(str收发ids = "", "", str收发ids & ",") & !收发ID
            
            .MoveNext
        Loop
                    
        If Not mobjPlugIn Is Nothing And str收发ids <> "" Then
            On Error Resume Next
            mobjPlugIn.DrugSendCustomCheck str收发ids, str返回收发ids, strReserve
            
            err.Clear: On Error GoTo 0
        End If
        
        If str返回收发ids <> "" Then
            mfrmDetail.SetSendBillStateByCustom str返回收发ids
        End If
    End With

End Sub
Private Function CheckDangerDrug(ByVal rsData As ADODB.Recordset) As Boolean
    '检查高危药品：如果高危药品需要单独发放时，则检查是否存在普通药品
    Dim bln普通药品 As Boolean
    Dim bln高危药品 As Boolean
    Dim lng药品id As Long
    
    If mParams.str高危发放 = "" Then
        CheckDangerDrug = True
        Exit Function
    End If
    
    With rsData
        .Filter = "执行状态=" & mState.发药
        .Sort = "药品ID Asc"
        
        Do While Not .EOF
            If lng药品id <> !药品ID Then
                If !高危药品 = 0 Then
                    bln普通药品 = True
                ElseIf InStr(1, mParams.str高危发放, !高危药品) > 0 Then
                    bln高危药品 = True
                End If
                
                If bln普通药品 And bln高危药品 Then
                    MsgBox "提示：高危药品不能和普通药品汇总发药！", vbInformation, gstrSysName
                    CheckDangerDrug = False
                    Exit Function
                End If
                    
                lng药品id = !药品ID
            End If
            .MoveNext
        Loop
    End With
    
    CheckDangerDrug = True
End Function
Private Sub DrugStoreWork_PrintBill()
    '打印发药单据
    Dim intAllFormat As Integer
    
    If GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\zl9Report\LocalSet\ZL1_BILL_1342", "AllFormat") <> "" Then
        intAllFormat = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\zl9Report\LocalSet\ZL1_BILL_1342", "AllFormat", 0))
    Else
        intAllFormat = Val(GetSetting("ZLSOFT", "私有模块\zl9Report\LocalSet\ZL1_BILL_1342", "AllFormat", 0))
    End If
    
    If mParams.int自动打印 = 2 Then
        If MsgBox("是否打印发药清单？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            If intAllFormat = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                    "发药库房=" & mcondition.lng药房id, _
                    "发药号=" & mcur汇总发药号, _
                    "领药部门=" & mstr科室名串 & "|" & " IN (" & mstr科室ID串 & ")", _
                    "包装系数=" & IIf(mParams.strUnit = "门诊单位", "S.门诊包装", "S.住院包装"), _
                    "PrintEmpty=0", 2)
            Else
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                    "发药库房=" & mcondition.lng药房id, _
                    "发药号=" & mcur汇总发药号, _
                    "领药部门=" & mstr科室名串 & "|" & " IN (" & mstr科室ID串 & ")", _
                    "包装系数=" & IIf(mParams.strUnit = "门诊单位", "S.门诊包装", "S.住院包装"), _
                    "ReportFormat=" & mfrmDetail.Get当前发药单格式, "PrintEmpty=0", 2)
            End If
        End If
    ElseIf mParams.int自动打印 = 1 Then
        If intAllFormat = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                "发药库房=" & mcondition.lng药房id, _
                "发药号=" & mcur汇总发药号, _
                "领药部门=" & mstr科室名串 & "|" & " IN (" & mstr科室ID串 & ")", _
                "包装系数=" & IIf(mParams.strUnit = "门诊单位", "S.门诊包装", "S.住院包装"), _
                "PrintEmpty=0", 2)
        Else
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                "发药库房=" & mcondition.lng药房id, _
                "发药号=" & mcur汇总发药号, _
                "领药部门=" & mstr科室名串 & "|" & " IN (" & mstr科室ID串 & ")", _
                "包装系数=" & IIf(mParams.strUnit = "门诊单位", "S.门诊包装", "S.住院包装"), _
                "ReportFormat=" & mfrmDetail.Get当前发药单格式, "PrintEmpty=0", 2)
        End If
    End If
End Sub

Private Sub DrugStoreWork_SendToPacker(ByVal rsData As ADODB.Recordset)
Dim str收发ids As String, strMessage As String
    Dim arr收发ids As Variant
    Dim lng当前部门id As Long
    Dim n As Integer
    
    On Error GoTo errHandle
    
    If mblnPackerConnect = True And Not mobjDrugMAC Is Nothing Then
        If TypeName(mobjDrugMAC) = "clsDrugMachine" Then
            '新接口
            
            With rsData
                .Filter = "执行状态=" & mState.发药 & " And 发药数量>0"
                .Sort = "领药部门ID,病人id"
        
                If .EOF Then Exit Sub
                            
                arr收发ids = Array()
                
                '按领药部门分批上传
                Do While Not .EOF
                    If lng当前部门id <> !领药部门ID Then
                        If str收发ids <> "" Then
                            ReDim Preserve arr收发ids(UBound(arr收发ids) + 1)
                            arr收发ids(UBound(arr收发ids)) = str收发ids
                        End If
                        
                        lng当前部门id = !领药部门ID
                        str收发ids = "2|" & !收发ID
                    Else
                        str收发ids = str收发ids & ";" & !收发ID
                    End If
                    
                    .MoveNext
                    
                    If .EOF And str收发ids <> "" Then
                        '后面没有记录时加入到数组
                        ReDim Preserve arr收发ids(UBound(arr收发ids) + 1)
                        arr收发ids(UBound(arr收发ids)) = str收发ids
                    End If
                Loop
                
                For n = 0 To UBound(arr收发ids)
                    mobjDrugMAC.Operation gstrDbUser, Val("21-明细上传"), CStr(arr收发ids(n)), strMessage
                Next
            End With
        Else
            '兼容老接口
            Call PackerTransDetail_DYEY(rsData)
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub PackerTransDetail_DLSY(ByVal rsData As ADODB.Recordset)
    '自动包药机数据传输：大连三院专用
    '直接传到HIS端的中间表
    
''流水号
'PrescriptionNo
''序号
'Seqno
''小组标志
'Group_No
''机器号
'MachineNo
''处理状态
'ProcFlg
''病人ID
'PatientID
''病人姓名
'PatientName
''病人性别
'Sex
''门诊住院标志
'IOFlg
''病区编码
'WardCd
''病区名称
'WardName
''床位号
'BedNo
''处方日期
'PrescriptionDate
''首次用药日期
'TakeDate
''开始时间
'TakeTime
''结束时间
'LastTime
''紧急类别
'Presc_Class
''药品编码
'Drugcd
''药品名称
'DrugName
''摆药单位
'DispensedUnit
''摆药天数
'Dispense_days
''用法
'Freq_desc
''服用时间
'Freq_desc_Detail
''填单时间
'MakeRecTime

End Sub
Private Sub PackerTransDetail_DYEY(ByVal rsData As ADODB.Recordset)
    '自动包药机数据传输：大医二院专用
    '调用接口函数传到中间数据库
    Dim str部门 As String
    Dim rsTmp As ADODB.Recordset
    Dim lng当前部门 As Long
    Dim str明细 As String
    Dim strReturn As String
    Dim str分包设备编号 As String
    Dim strTmp As String
    Dim strFilter As String
    Dim strDetail As String
    
    On Error GoTo errHandle
    
    If mblnStartPacker = False Or mblnPackerConnect = False Then Exit Sub
    If mParams.int暂停传送 = 1 Then Exit Sub
    If Val(Mid(mParams.str包装机单据, 1, 1)) = 0 And Val(Mid(mParams.str包装机单据, 2, 1)) = 0 Then Exit Sub
    If mParams.str包装机剂型 = "" Then Exit Sub
    
    str分包设备编号 = "1"
    
    If mlng药房ID <> mcondition.lng药房id Or mstr药房编码 = "" Then
        mlng药房ID = mcondition.lng药房id
        gstrSQL = "select 编码 from 部门表 where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "取药房编码", mlng药房ID)
        mstr药房编码 = rsTmp!编码
    End If
    
    With rsData
        If Val(Mid(mParams.str包装机单据, 1, 1)) = 1 Then
            strFilter = "执行状态=" & mState.发药 & " And 类型='长嘱' And 发药数量>0"
        End If
        If Val(Mid(mParams.str包装机单据, 2, 1)) = 1 Then
            If strFilter <> "" Then
                strFilter = "(" & strFilter & ")"
                strFilter = strFilter & " Or (执行状态=" & mState.发药 & " And 类型='临嘱' And 发药数量>0) "
            Else
                strFilter = "执行状态=" & mState.发药 & " And 类型='临嘱' And 发药数量>0 "
            End If
        End If
        
        .Filter = strFilter
        .Sort = "领药部门ID"
        
        If .EOF Then Exit Sub
        
        lng当前部门 = !领药部门ID
        str部门 = !领药部门编码 & ";" & mstr药房编码 & ";" & str分包设备编号
        Do While Not .EOF
            If lng当前部门 <> !领药部门ID Then
                '当前部门不一样时，传递数据，并返回没有传递成功的收发ID
                If str明细 <> "" Then
                    strReturn = IIf(strReturn = "", "", strReturn & ";") & mobjDrugMAC.TranDrugPacker(str部门 & "|" & str明细)
                End If
                
                '重新指定当前部门
                lng当前部门 = !领药部门ID
                str部门 = !领药部门编码 & ";" & mstr药房编码 & ";" & str分包设备编号
                str明细 = GetMediPackerDetail(!收发ID, mParams.str包装机剂型, !类型)
            Else
                strDetail = GetMediPackerDetail(!收发ID, mParams.str包装机剂型, !类型)
                If strDetail <> "" Then
                    str明细 = IIf(str明细 = "", "", str明细 & "|") & strDetail
                End If
            End If
            
            .MoveNext
            
            If .EOF And str明细 <> "" Then
                '后面没有记录时，传递数据，并返回没有传递成功的收发ID
                strReturn = IIf(strReturn = "", "", strReturn & ";") & mobjDrugMAC.TranDrugPacker(str部门 & "|" & str明细)
            End If
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub ExecuteWriteOffByMessage(ByVal objMsgBar As CommandBarControl)
    '通过消息调用冲销界面
    '由用户点击的消息项目来传递冲销需要的关键信息
    '传递格式为：申请时间,病人id|申请时间,病人id
    '如果是某个具体消息，那么取菜单中保存的信息来传递；如果是批量执行的，那么根据记录集来传递消息
    Dim strMsg As String
    
    If Not objMsgBar Is Nothing Then
        If objMsgBar.Parameter <> "" Then
            strMsg = objMsgBar.Parameter
        Else
            With mrsReceiveMsg
                If mrsReceiveMsg.RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        strMsg = IIf(strMsg = "", "", strMsg & "|") & !申请时间 & "," & !病人ID
                        .MoveNext
                    Loop
                End If
            End With
        End If
       
        '调用销账审核窗口
        Call ShowWindow_ReVerify(strMsg)
    End If
End Sub

Private Sub SetMessageBar(ByVal rsMsg As ADODB.Recordset)
    '设置消息菜单
    '先删除子菜单，再根据记录集中的数据增加子菜单
    '有数据时父菜单显示有多少消息；如果没有任何消息，则隐藏父菜单
    '添加子菜单时：如果消息超过5条，则只显示5条具体消息
    '如果消息超过1条时，另外添加一个子菜单“全部审核”
    'strDelMsg：不为空时，删除消息记录集中对应的项目
    Dim cbrControlMain As CommandBarPopup
    Dim cbrControlPopup As CommandBarControl
    Dim intCount As Integer
    Dim intTemp As Integer
    Dim blnTemp As Boolean
                
    If mobjMipModule Is Nothing Then Exit Sub
    
    If rsMsg Is Nothing Then Exit Sub
    
    Set cbrControlMain = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_File_Message)
    cbrControlMain.Visible = True
    If Not cbrControlMain Is Nothing Then
        Set cbrControlMain = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_File_Message)
        cbrControlMain.Visible = True
        If rsMsg.RecordCount > 0 Then rsMsg.MoveFirst
        If rsMsg.RecordCount = 0 Then
            cbrControlMain.Visible = False
        Else
            cbrControlMain.Caption = "↓消息提醒" & "(" & rsMsg.RecordCount & ")"
            
            For Each cbrControlPopup In cbrControlMain.CommandBar.Controls
                If Not rsMsg.EOF And intTemp <= 5 Then
                    cbrControlPopup.Caption = Format(rsMsg!申请时间, "mm-dd hh:mm") & " " & rsMsg!姓名 & " " & rsMsg!住院号
                    cbrControlPopup.Parameter = rsMsg!申请时间 & "," & rsMsg!病人ID
                    cbrControlPopup.Visible = True
                    rsMsg.MoveNext
                Else
                    If intTemp < cbrControlMain.CommandBar.Controls.count Then
                        cbrControlPopup.Visible = False
                    Else
                        blnTemp = True
                    End If
                End If
                
                intTemp = intTemp + 1
            Next
                
            For intCount = intTemp + 1 To rsMsg.RecordCount
                If intCount <= 5 Then
                    Set cbrControlPopup = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_File_Message + intCount, Format(rsMsg!申请时间, "mm-dd hh:mm") & " " & rsMsg!姓名 & " " & rsMsg!住院号)
                    cbrControlPopup.Parameter = rsMsg!申请时间 & "," & rsMsg!病人ID
                Else
                    Exit For
                End If
                rsMsg.MoveNext
            Next
            If intCount > 2 And (blnTemp = True Or intTemp < 6) Then
                Set cbrControlPopup = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_File_Message + intCount, "全部审核")
            End If
        End If
    End If
End Sub

Private Sub cbo发药药房_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str工作性质 As String
    
    str工作性质 = "L,M,N"
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cbo发药药房.ListCount = 0 Then Exit Sub
    
    If cbo发药药房.ListIndex >= 0 Then
        If Val(cbo发药药房.Tag) = cbo发药药房.ItemData(cbo发药药房.ListIndex) Then
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cbo发药药房, Trim(cbo发药药房.Text), str工作性质, IIf(IsInString(mstrPrivs, "所有药房", ";"), False, True), "2,3") = False Then
        Exit Sub
    End If
    If cbo发药药房.ListIndex >= 0 Then
        cbo发药药房.Tag = cbo发药药房.ItemData(cbo发药药房.ListIndex)
    End If
End Sub

Private Sub cbo发药药房_KeyPress(KeyAscii As Integer)
    '屏蔽输入单引号
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cbo发药药房_Validate(Cancel As Boolean)
    If cbo发药药房.ListCount > 0 Then
        If cbo发药药房.ListIndex = -1 Then
            MsgBox "请选择一个药库或者药房！", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub chkDanger_Click()
    chkDangerType(0).Enabled = (chkDanger.Value = 1)
    chkDangerType(1).Enabled = chkDangerType(0).Enabled
    chkDangerType(2).Enabled = chkDangerType(0).Enabled
    
    
End Sub

Private Sub chkDangerType_Click(index As Integer)
    Dim objChk As CheckBox
    Dim blnAllUnCheck As Boolean
    
    If mblnStart = False Then Exit Sub
    
    blnAllUnCheck = True
    
    For Each objChk In chkDangerType
        If objChk.Value = 1 Then
            blnAllUnCheck = False
        End If
    Next
    
    If blnAllUnCheck = True Then
        chkDangerType(index).Value = 1
    End If
End Sub


Private Sub chkToxicologyType_Click()
    Me.chkToxicology(0).Enabled = (Me.chkToxicologyType.Value = 1)
    Me.chkToxicology(1).Enabled = Me.chkToxicology(0).Enabled
    Me.chkToxicology(2).Enabled = Me.chkToxicology(0).Enabled
    Me.chkToxicology(3).Enabled = Me.chkToxicology(0).Enabled
End Sub

Private Sub InitIDKindNew()
    Dim strTemp As String
    
    strTemp = "住|住院号|0;姓|姓名|0;床|床号|0;单|单据号|0;病|病人ID|0;领|领药号|0;发|发药号|0;部|领药部门|0;IC|IC卡|1"
    Me.IDKNType.IDKindStr = strTemp
    Call IDKNType.zlInit(Me, glngSys, mlngMode, gcnOracle, gstrDbUser, mobjSquareCard, strTemp, txtInput)
'    IDKNType.SetAutoReadCard True
    Me.IDKNType.IDKind = mParams.int输入模式索引
End Sub

Private Sub IDKNType_ItemClick(index As Integer, objCard As zlIDKind.Card)
    mParams.int输入模式索引 = index
    mParams.int输入模式 = Get输入模式(IDKNType.GetCurCard.名称)
    
    If objCard.卡号密文规则 <> "" Then
        txtInput.PasswordChar = "*"
    Else
        txtInput.PasswordChar = ""
    End If
    
    txtInput.Text = ""
    
    Call picConMain_Resize
End Sub

Private Function Get输入模式(ByVal str类型 As String) As Integer
    '从IDKind中返回当前程序内部所定义的类型
    Dim i As Integer
    Dim str类型串 As String
    
    'str类型串与传入的IDKindStr类型名称、顺序一致
    str类型串 = "住院号,姓名,床号,单据号,病人ID,领药号,发药号,领药部门,IC卡"
    
    For i = 0 To UBound(Split(str类型串, ","))
        If Split(str类型串, ",")(i) = str类型 Then
            Get输入模式 = i + 1
            Exit For
        End If
    Next
    
    '当IDKindf返回的类型不是IDKindStr传入的类型，则结合一卡通的序号赋值一个大于IDKindStr类型个数的有效数字
    If Get输入模式 = 0 Then
        For i = 0 To UBound(Split(mstrCardType, ";"))
            If Split(Split(mstrCardType, ";")(i), "|")(1) = str类型 Then
                Get输入模式 = i + 10
                Exit For
            End If
        Next
    End If
    
End Function

Private Sub IDKNType_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtInput.Text = objPatiInfor.卡号
    If txtInput.Text <> "" Then Call txtInput_KeyPress(vbKeyReturn)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strNo As String)
    If Not txtInput.Locked And txtInput.Text = "" And Me.ActiveControl Is txtInput And strNo <> "" Then
        txtInput.Text = strNo
        
        If txtInput.Text = "" Then
            Call mobjICCard.SetEnabled(False)
        Else
            If mParams.int输入模式 <> mInputType.IC卡 Then
                mParams.int输入模式 = mInputType.IC卡

                DoEvents
            End If
            
            Call txtInput_KeyPress(vbKeyReturn)
        End If
    End If
End Sub
Private Function DrugStoreWork_CheckSend(ByVal rsData As ADODB.Recordset) As Boolean
    '发药检查
    Dim rsGroupCheck As ADODB.Recordset
    Dim strCheckMsg As String
    
    On Error GoTo errHandle
    
    '优先检查高危药品
    If CheckDangerDrug(rsData) = False Then Exit Function
    
    '检查存储库房
    If CheckDrugStock(rsData) = False Then Exit Function
    
    '检查是否存在销帐申请未审核的单据
    If CheckNotAudited(rsData) = False Then Exit Function
    
    '检查处方是否已结帐、检查该病人是否已出院，并对权限进行检查
    If Not CheckCorrelation(0, rsData) Then Exit Function
    
    '库存数量检查
    If CheckShortage(rsData, True, strCheckMsg) = False Then
        '库存检查
        If mParams.IntCheckStock = 2 Then
            '库存不足禁止发药
            MsgBox "以下药品实际库存数量不足，不能发药！" & vbCrLf & strCheckMsg, vbInformation, gstrSysName
            
            If mParams.bln缺药检查 Then
                Call mfrmDetail.RefreshList(mListType.发药, mrsSendData, mrsChargeOff)
            End If
            Exit Function
        ElseIf mParams.IntCheckStock = 1 Then
            '库存不足，提醒
            If MsgBox("以下药品实际库存数量不足，是否继续发药？" & vbCrLf & strCheckMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If mParams.bln缺药检查 Then
                    Call mfrmDetail.RefreshList(mListType.发药, mrsSendData, mrsChargeOff)
                End If
                Exit Function
            End If
        End If
    End If
    
    '单据状态及分组检查
    Set rsGroupCheck = rsData.Clone
    With rsData
        .Filter = "执行状态=" & mState.发药
        .Sort = "收发ID"
        Do While Not .EOF
            '检查单据状态
            If DeptSendWork_CheckBill(1, !收发ID, mParams.bln允许未审核处方发药) > 0 Then Exit Function
            
            '检查分组状态
            If Not mblnCheck Then
                If CheckGroupSend(rsGroupCheck, Val(!相关id), !NO) = False Then Exit Function
            End If
            
            .MoveNext
        Loop
    End With
    
    '零差价管理
    If CheckPriceAdjustByRecord(rsData, mcondition.lng药房id) = False Then
        Exit Function
    End If
    
    DrugStoreWork_CheckSend = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckPriceAdjustByRecord(ByVal rsData As ADODB.Recordset, ByVal lng库房ID As Long) As Boolean
    '按记录集检查零差价
    Dim rsItem As ADODB.Recordset
    Dim strArr收发id As Variant
    Dim str收发ID串 As String
    Dim strTmp As String
    Dim strMsg As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    '如果没开启全局的零差价管理，则不进行后续检查，返回true
    If Val(zlDatabase.GetPara(275, 100, , 0)) = 0 Then CheckPriceAdjustByRecord = True: Exit Function
    
    '发药状态且有零差价管理属性的药品才要进行检查
    rsData.Filter = "零差价管理=1 And 执行状态=" & mState.发药
    rsData.Sort = "药品ID,批次"
    If rsData.EOF Then CheckPriceAdjustByRecord = True: Exit Function
    
    Do While Not rsData.EOF
        If strTmp <> rsData!药品ID & "," & rsData!批次 Then
            strTmp = rsData!药品ID & "," & rsData!批次
            
            str收发ID串 = IIf(str收发ID串 = "", "", str收发ID串 & "|") & strTmp
        End If
        
        rsData.MoveNext
    Loop
    
    strArr收发id = GetArrayByStr(str收发ID串, 3950, "|")
    
    For i = 0 To UBound(strArr收发id)
        strMsg = CheckPriceAdjustBatch(lng库房ID, CStr(strArr收发id(i)))
        If strMsg <> "" Then
            MsgBox "以下药品已启用了零差价管理，但在药房中售价和成本价不一致，不能发药，请先进行调价再发药！" & _
                 vbCrLf & strMsg, vbInformation, gstrSysName
        End If
    Next
    
    CheckPriceAdjustByRecord = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub AutoRefresh()
    '自动刷新只针对未发药品清单
    Dim dateCurr As Date
        
    '如果窗口最小化时退出
    If Me.WindowState = 1 Then Exit Sub
    
    '如果活动窗口不是当前窗口时退出
    If mlngMyWindow = 0 Then
        mlngMyWindow = GetActiveWindow()
    Else
        If mlngMyWindow <> GetActiveWindow() Then Exit Sub
    End If
    
    '如果不是未发药界面或者自动刷新参数为0时退出
    If tbcDetail.Selected.index <> mListType.发药 Or mParams.int自动刷新未发药清单 = 0 Then Exit Sub
    
    '根据当前时间与上次刷新时间间隔来控制是否刷新
    dateCurr = Sys.Currentdate
    If DateDiff("s", mdate上次刷新时间, dateCurr) < mParams.int自动刷新未发药清单 * 60 Then Exit Sub
    
    TimerAuto.Enabled = False
    
    '刷新数据
    cmdRefresh_Click

'    MsgBox "Ok！" & "[" & Format(dateCurr, "yyyy-mm-dd hh:mm:ss") & "]" & "[" & Format(mdate上次刷新时间, "yyyy-mm-dd hh:mm:ss") & "]"
'    mdate上次刷新时间 = Sys.Currentdate
    
    DoEvents
    TimerAuto.Enabled = True
End Sub

Private Sub BillPrint_Restore()
    '功能：打印退药通知单
    Dim StrDate As String
    
    StrDate = Format(mfrmDetail.GetReturnDate, "yyyy-MM-dd HH:mm:ss")
    
    If StrDate = "" Then Exit Sub
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, _
        "退药时间=" & StrDate, _
        "包装系数=" & IIf(mParams.strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), _
        "发药库房=" & mcondition.lng药房id, _
        2)
End Sub


Private Sub BillPrint_Total()
    Dim rsTmp As ADODB.Recordset
    Dim str药房 As String, str科室 As String
    Dim str发药 As String
    Dim str领药部门 As String
    Dim str领药部门ID As String
    Dim var发药号 As Variant
    Dim intAllFormat As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select 编码,名称 From 部门表 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[读取当前药房的名称]", mcondition.lng药房id)

    If Not rsTmp.RecordCount <= 0 Then str药房 = "(" & rsTmp!编码 & ")" & rsTmp!名称
    
    If GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\zl9Report\LocalSet\ZL1_BILL_1342", "AllFormat") <> "" Then
        intAllFormat = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\zl9Report\LocalSet\ZL1_BILL_1342", "AllFormat", 0))
    Else
        intAllFormat = Val(GetSetting("ZLSOFT", "私有模块\zl9Report\LocalSet\ZL1_BILL_1342", "AllFormat", 0))
    End If
    
    If tbcDetail.Selected.index = mListType.退药 Then
        str发药 = mfrmDetail.GetSendedInfo
                
        If str发药 <> "" Then
            str领药部门 = Split(str发药, "|")(0)
            str领药部门ID = Split(str发药, "|")(1)
            var发药号 = Split(str发药, "|")(2)
        End If
        
        If intAllFormat = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                "发药库房=" & str药房 & "|" & mcondition.lng药房id, _
                "发药号=" & var发药号, _
                "领药部门=" & str领药部门 & "|" & " IN (" & str领药部门ID & ")")
        Else
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                "发药库房=" & str药房 & "|" & mcondition.lng药房id, _
                "发药号=" & var发药号, _
                "领药部门=" & str领药部门 & "|" & " IN (" & str领药部门ID & ")", "ReportFormat=" & mfrmDetail.Get当前发药单格式)
        End If
    Else
        If intAllFormat = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                "发药库房=" & str药房 & "|" & mcondition.lng药房id, _
                "领药部门=" & mstr科室名串 & "|" & " IN (" & mstr科室ID串 & ")", _
                "包装系数=" & IIf(mParams.strUnit = "门诊单位", "S.门诊包装", "S.住院包装"))
        Else
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1342", Me, _
                "发药库房=" & str药房 & "|" & mcondition.lng药房id, _
                "领药部门=" & mstr科室名串 & "|" & " IN (" & mstr科室ID串 & ")", _
                "包装系数=" & IIf(mParams.strUnit = "门诊单位", "S.门诊包装", "S.住院包装"), "ReportFormat=" & mfrmDetail.Get当前发药单格式)
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub BillPrint_Wait()
    Dim rsTmp As New ADODB.Recordset
    Dim str显示 As String, str绑定 As String
    Dim str药房 As String, i As Long
    Dim n As Integer

    '库房条件
    On Error GoTo errHandle
    gstrSQL = "Select 名称 From 部门表 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[读取当前药房的名称]", mcondition.lng药房id)

    str药房 = rsTmp!名称 & "|" & mcondition.lng药房id

    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1342_1", Me, _
        "住院药局=" & str药房, "住院科室=" & mstr科室名串 & "|" & " IN (" & mstr科室ID串 & ")", _
        "开始时间=" & mcondition.str开始时间, "结束时间=" & mcondition.str结束时间, 1)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckShortage(ByRef rsSendData As ADODB.Recordset, ByVal blnSendCheck As Boolean, Optional ByRef strMsg As String) As Boolean
    '缺药检查
    '1、blnSendCheck=False：自动缺药检查时取数据的库存数量来和汇总发药数量比较
    '2、blnSendCheck=True：发药检查时取当前数据库的库存数量来和汇总发药数量比较
    
    Dim rsData As ADODB.Recordset
    Dim dblSum As Double
    Dim dblStock As Double
    Dim str当前药品 As String
    Dim blnTmp As Boolean       '是否有新的缺药
    Dim str品名 As String
    Dim intCount As Integer
    
    blnTmp = True
    
    If mParams.IntCheckStock = 0 Then
        CheckShortage = True
        Exit Function
    End If
    
    rsSendData.Filter = "执行状态=" & mState.发药
    rsSendData.Sort = "药品ID,批次,NO"

    With rsSendData
        Do While Not .EOF
            If str当前药品 <> !药品ID & ";" & !批次 Then
                If blnSendCheck = True Then
                    dblSum = MediWork_GetMediRealAmount(mcondition.lng药房id, Val(!药品ID), Val(!批次))
                Else
                    dblSum = zlStr.NVL(!库存数量, 0)
                End If
                
                str当前药品 = !药品ID & ";" & !批次
            End If
            
            dblSum = dblSum - !发药数量
                
            If dblSum < 0 Then
                If str品名 <> !品名 Then
                    str品名 = !品名
                    
                    intCount = intCount + 1
                    If intCount < 6 Then
                        strMsg = IIf(strMsg = "", str品名, strMsg & vbCrLf & str品名)
                    End If
                End If
                
                If !执行状态 <> mState.缺药 Then
                    If mParams.bln缺药检查 Then
                        !执行状态 = mState.缺药
                        !状态 = "缺药"
                        .Update
                    End If
                    blnTmp = False
                End If
            End If
            
            .MoveNext
        Loop
        
        rsSendData.Filter = ""
        rsSendData.Sort = ""
    End With
    
    If strMsg <> "" Then
        If intCount > 5 Then strMsg = strMsg & vbCrLf & "其他还有" & intCount - 5 & "个缺药药品......"
    End If
    
    CheckShortage = blnTmp
End Function

Private Function CheckNotAudited(ByRef rsData As ADODB.Recordset) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim bln销帐申请 As Boolean
    Dim bln允许发送 As Boolean
    
    On Error GoTo errHandle
    
    If mParams.bln检查销帐申请 = False Then CheckNotAudited = True: Exit Function
    
    Call InitApplyforcredit
    
    CheckNotAudited = True
    bln销帐申请 = True
    
    gstrSQL = "Select 数量 As 销帐申请数量 From 病人费用销帐 Where 费用id = [1] And 状态 = 0"
    
    With rsData
        .Filter = "执行状态=" & mState.发药
        .Sort = "药品ID Asc"
        
        Do While Not .EOF
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在销帐申请未审核的单据", rsData!费用ID)
            
            If rsTmp.RecordCount > 0 Then
                bln销帐申请 = False
                
                With mrsApplyforcredit
                    .AddNew
                    
                    !执行状态 = rsData!执行状态
                    !费用ID = rsData!费用ID
                    !收发ID = rsData!收发ID
                    !NO = rsData!NO
                    !药品名称 = rsData!药品名称
                    !批号 = rsData!批号
                    !数量 = rsData!数量
                    !销帐申请数量 = zlStr.FormatEx(rsTmp!销帐申请数量 / rsData!包装, mintNumberDigit) & rsData!单位
                    !姓名 = rsData!姓名
                    !性别 = rsData!性别
                    !年龄 = rsData!年龄
                    !领药部门 = rsData!领药部门
                    !床号 = rsData!床号
                    !病人科室 = rsData!科室
                End With

            End If
            
            .MoveNext
        Loop
    End With
    
    '对含有销帐申请的单据进行处理
    If bln销帐申请 = False Then
        Call frm部门发药销帐申请清单.ShowCard(Me, mrsApplyforcredit, bln允许发送)
        
        '由子窗体返回用户是否继续执行操作，若【取消】则禁止继续发送
        CheckNotAudited = bln允许发送
        If CheckNotAudited = False Then Exit Function
        
        '修正取消发送的单据的执行状态
        mrsApplyforcredit.Filter = "执行状态 = 3"
        If mrsApplyforcredit.RecordCount > 0 Then
            Do While Not mrsApplyforcredit.EOF
                rsData.Filter = "收发ID = " & mrsApplyforcredit!收发ID
                rsData!执行状态 = 3
                rsData.Update
                mrsApplyforcredit.MoveNext
            Loop
        End If
        
        rsData.Filter = ""
    End If
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckDrugStock(ByVal rsData As ADODB.Recordset) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Integer
    Dim lng药品id As Long
    
    If mParams.bln检查存储库房 = False Then CheckDrugStock = True: Exit Function
    
    CheckDrugStock = True
    With rsData
        .Filter = "执行状态=" & mState.发药
        .Sort = "药品ID Asc"
        
        Do While Not .EOF
            If lng药品id <> !药品ID Then
                If MediWork_CheckStorageStock(mcondition.lng药房id, Val(!药品ID)) = False Then
                    MsgBox !品名 & "未设置存储库房，不能发药！", vbInformation, gstrSysName
                    CheckDrugStock = False
                    Exit Function
                End If
                    
                lng药品id = !药品ID
            End If
            .MoveNext
        Loop
    End With
    
    CheckDrugStock = True
End Function
Private Sub ClearData(ByVal intType As Integer)
    '清除界面数据：清除科室列表数据，清除明细数据
    
    ClearTreeView IIf(intType = mListType.发药, 0, 1)
    ClearDetailList intType
End Sub

Private Sub ClearDetailList(ByVal intType As Integer)
    '清空明细列表
    If intType <> mListType.发药 Then
        mfrmDetail.ClearList mListType.退药
    Else
        mfrmDetail.ClearList mListType.发药
    End If
End Sub

Private Sub ClearTreeView(ByVal intType As Integer)
    tvwList(intType).Nodes.Clear
    tvwList(intType).Tag = 1
    chkAll(intType).Value = 0
End Sub

Private Function DrugStoreWork_SendProc(ByVal rsData As ADODB.Recordset, ByVal StrCurDate As String) As Boolean
   '处理发药数据
    Dim lng病人ID As Long
    Dim strID批次串 As String         '格式：收发ID,批次|收发ID,批次...
    Dim strID串 As String             '格式：收发ID,收发ID...
    Dim blnBeginTrans As Boolean
    Dim str领药人 As String
    Dim str配药人 As String
    Dim str核查人 As String
    Dim blnUpdate As Boolean
    Dim str签名记录 As String
    Dim strsql As String
    Dim arrSql As Variant
    Dim lngRow As Long
    Dim strFilter As String
    Dim blnIsCommit As Boolean        '是否有数据提交
    Dim strInputID As String
    Dim rsSign As ADODB.Recordset     '用于处理电子签名
    Dim strReserve As String
    Dim blnzlPlugInResult As Boolean    '用于zlPlugIn接口返回结果
    Dim blnCheck As Boolean           '用于优化电子签名的重复检查数据。False-需要重复；True-不重复
    
    arrSql = Array()
    
    '领药人签名
    TimerAuto.Enabled = False
    str领药人 = ""
    If mParams.bln领药人签名 = True Then
        str领药人 = zlDatabase.UserIdentify(Me, "领药人签名", glngSys, 1342, "")
        If str领药人 = "" Then
            TimerAuto.Enabled = True
            Exit Function
        End If
    End If
    TimerAuto.Enabled = True
    
    '取配药人
    str配药人 = mfrmDetail.Get当前配药人
    
    '取审核人
    str核查人 = mfrmDetail.Get当前核查人
       
    On Error GoTo errHandle
    
    '克隆发药数据集用于电子签名，不能在循环中直接用发药数据集
    Set rsSign = rsData.Clone
    
    '发药（按病人ID批量发药）
    With rsData
        .Filter = "执行状态=" & mState.发药
        
        '必须按病人ID，药品ID排序
        .Sort = "病人ID Asc ,药品ID Asc"
        
        Do While Not .EOF
            If lng病人ID = 0 Then
                lng病人ID = !病人ID
            End If
                
            '病人ID相同时候
            If lng病人ID = !病人ID Then
                strID批次串 = IIf(strID批次串 = "", !收发ID & "," & zlStr.NVL(!批次, 0), strID批次串 & "|" & !收发ID & "," & zlStr.NVL(!批次, 0))
                strID串 = IIf(strID串 = "", !收发ID, strID串 & "," & !收发ID)
                If InStr(1, strInputID, !医嘱id & ",1|") < 1 And NVL(!医嘱id, 0) <> 0 And Not (!类别 = "E" And !执行分类 = 1 And mblnIs配置中心) Then
                    strInputID = strInputID & !医嘱id & ",1|"
                End If
            Else
                '如果病人ID不同则提交事务
                gstrSQL = "Zl_药品收发记录_批量发药("
                '收发ID，批次串
                gstrSQL = gstrSQL & "'" & strID批次串 & "'"
                '库房ID
                gstrSQL = gstrSQL & "," & mcondition.lng药房id
                '审核人
                gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                '审核日期
                gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
                '发药方式
                gstrSQL = gstrSQL & ",3"
                '领药人
                gstrSQL = gstrSQL & ",'" & str领药人 & "'"
                '汇总发药号
                gstrSQL = gstrSQL & "," & mcur汇总发药号
                '金额保留位数
                gstrSQL = gstrSQL & "," & mParams.int金额保留位数
                '配药人
                gstrSQL = gstrSQL & ",'" & str配药人 & "'"
                '审核处方人
                gstrSQL = gstrSQL & ",'" & str核查人 & "'"
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
                                    
                If mParams.bln审核划价单 = True Then
                    gstrSQL = "Zl_住院记帐记录_发药审核("
                    '收发ID串
                    gstrSQL = gstrSQL & "'" & strID串 & "'"
                    '操作员编号
                    gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
                    '操作员姓名
                    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                    '审核时间
                    gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
                
                '发药则表示医嘱审核通过
                If strInputID <> "" And mParams.int发药时审核医嘱 = 1 Then
                    gstrSQL = "Zl_输液配药记录_审核("
                    '医嘱ID
                    gstrSQL = gstrSQL & "'" & strInputID & "'"
                    gstrSQL = gstrSQL & ")"
                
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
                
                '签名失败，发药事务退出
                If gblnESign部门发药 = True And gblnESignUserStoped = False Then
                    mstr病人id = IIf(mstr病人id = "", "病人ID <>" & lng病人ID, mstr病人id & " And 病人ID <>" & lng病人ID)
                    gstrSQL = Signature(rsSign, StrCurDate, str配药人, lng病人ID, blnCheck)
                    If gstrSQL = "" Then Exit Function
                    
                    blnCheck = True
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
                
                '调用发药前的外挂接口
                err.Clear
                If Not mobjPlugIn Is Nothing Then
                    On Error Resume Next
                    If mobjPlugIn.DrugBeforeSendBySumID(mcondition.lng药房id, strID串, strReserve) = False Then
                        If err.Number <> 0 Then
                            err.Clear: On Error GoTo 0
                        Else
                            Exit Function
                        End If
                    End If
                    err.Clear: On Error GoTo 0
                End If
                
                On Error GoTo errHandle
                
                gcnOracle.BeginTrans
                blnBeginTrans = True
                
                For lngRow = 0 To UBound(arrSql)
                    Call zlDatabase.ExecuteProcedure(CStr(arrSql(lngRow)), Me.Caption & "-电子签名")
                Next
                
                gcnOracle.CommitTrans
                
                blnIsCommit = True
                blnBeginTrans = False
                blnUpdate = True
                strFilter = IIf(strFilter = "", "(病人id=" & lng病人ID & " and 执行状态=1)", strFilter & " or (病人id=" & lng病人ID & " and 执行状态=1)")
                lng病人ID = !病人ID
                arrSql = Array()
                strID批次串 = !收发ID & "," & zlStr.NVL(!批次, 0)
                strID串 = !收发ID
                If NVL(!医嘱id, 0) <> 0 And Not (!类别 = "E" And !执行分类 = 1 And mblnIs配置中心) Then
                    strInputID = !医嘱id & ",1|"
                End If
            End If
            
            .MoveNext
            
            '如果后面没有记录并且传入字符串不为空，则提交事务
            If .EOF And strID批次串 <> "" Then
                gstrSQL = "Zl_药品收发记录_批量发药("
                '收发ID，批次串
                gstrSQL = gstrSQL & "'" & strID批次串 & "'"
                '库房ID
                gstrSQL = gstrSQL & "," & mcondition.lng药房id
                '审核人
                gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                '审核日期
                gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
                '发药方式
                gstrSQL = gstrSQL & ",3"
                '领药人
                gstrSQL = gstrSQL & ",'" & str领药人 & "'"
                '汇总发药号
                gstrSQL = gstrSQL & "," & mcur汇总发药号
                '金额保留位数
                gstrSQL = gstrSQL & "," & mParams.int金额保留位数
                '配药人
                gstrSQL = gstrSQL & ",'" & str配药人 & "'"
                '审核处方人
                gstrSQL = gstrSQL & ",'" & str核查人 & "'"
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
                                   
                If mParams.bln审核划价单 = True Then
                    gstrSQL = "Zl_住院记帐记录_发药审核("
                    '收发ID串
                    gstrSQL = gstrSQL & "'" & strID串 & "'"
                    '操作员编号
                    gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
                    '操作员姓名
                    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                    '审核时间
                    gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
                    gstrSQL = gstrSQL & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                    
                End If
                
                '发药则表示医嘱审核通过
                If strInputID <> "" And mParams.int发药时审核医嘱 = 1 Then
                    gstrSQL = "Zl_输液配药记录_审核("
                    '医嘱ID
                    gstrSQL = gstrSQL & "'" & strInputID & "'"
                    gstrSQL = gstrSQL & ")"
                
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
                
                '签名失败，发药事务退出
                If gblnESign部门发药 = True And gblnESignUserStoped = False Then
                    mstr病人id = IIf(mstr病人id = "", "病人ID <>" & lng病人ID, mstr病人id & " And 病人ID <>" & lng病人ID)
                    gstrSQL = Signature(rsSign, StrCurDate, str配药人, lng病人ID, blnCheck)
                    If gstrSQL = "" Then Exit Function
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
                
                '调用发药前的外挂接口
                err.Clear
                If Not mobjPlugIn Is Nothing Then
                    On Error Resume Next
                    If mobjPlugIn.DrugBeforeSendBySumID(mcondition.lng药房id, strID串, strReserve) = False Then
                        If err.Number <> 0 Then
                            err.Clear: On Error GoTo 0
                        Else
                            Exit Function
                        End If
                    End If
                    err.Clear: On Error GoTo 0
                End If
                
                On Error GoTo errHandle
                
                gcnOracle.BeginTrans
                blnBeginTrans = True
                
                For lngRow = 0 To UBound(arrSql)
                    Call zlDatabase.ExecuteProcedure(CStr(arrSql(lngRow)), Me.Caption & "-电子签名")
                Next
                gcnOracle.CommitTrans
                blnIsCommit = True
                blnBeginTrans = False
            End If
        Loop
    End With
    
    '调用发药后的外挂接口
    If Not mobjPlugIn Is Nothing Then
        On Error Resume Next
        mobjPlugIn.DrugSendBySumID mcondition.lng药房id, mcur汇总发药号, strReserve
        err.Clear: On Error GoTo 0
    End If
    
    DrugStoreWork_SendProc = True
    Exit Function
errHandle:
    '如果已开启事务，并且未提交，则出错时回滚事务
    If blnBeginTrans Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    '已提交过数据，打印提交的汇总数据
    If blnIsCommit = True Then
        Call DrugStoreWork_PrintBill
    End If
End Function


Private Function Signature(ByVal rsData As Recordset, ByVal StrCurDate As String, ByVal str配药人 As String, ByVal lng病人ID As Long, Optional blnCheck As Boolean) As String
    Dim str签名记录 As String
    Dim strsql As String
    Dim rstemp As Recordset
    Dim lng签名id As Long
    Dim str收发ID As String
    Dim lngRow As Long
    Dim strTemp As String
    Dim arrSql As Variant
    
    On Error GoTo errHandle
    
    arrSql = Array()
    rsData.Filter = "病人id=" & lng病人ID
    '进行签名处理
    If gblnESign部门发药 = True And gblnESignUserStoped = False Then
        If rsData.RecordCount > 0 Then
            If GetSignatureRecoredGather(EsignTache.send, rsData, mcondition.lng药房id, str配药人, gstrUserName, StrCurDate, str签名记录, blnCheck) = False Then
                Exit Function
            End If
            
            If str签名记录 <> "" Then
                lng签名id = Sys.NextId("药品签名记录")
                
                str收发ID = Mid(Mid(str签名记录, 1, Len(str签名记录) - 1), InStrRev(Mid(str签名记录, 1, Len(str签名记录) - 1), "'") + 1)
                str签名记录 = Mid(Mid(str签名记录, 1, Len(str签名记录) - 1), 1, InStrRev(Mid(str签名记录, 1, Len(str签名记录) - 1), "'") - 1)
                
                strsql = "Zl_药品签名记录_Insert(" & str签名记录 & "'" & str收发ID & "'," & lng签名id & ")"
                
                Signature = strsql
                rsData.Filter = mstr病人id
            Else
                MsgBox "对发药人电子签名失败！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function











Private Function DrugStoreWork_StayProc(ByVal StrCurDate As String) As Boolean
    '处理留存数据
    Dim rsData As ADODB.Recordset
    Dim Str期间 As String
    Dim arrSql As Variant
    Dim lngRow As Long
    Dim int留存方式 As Integer
    
    '取留存数据集
    Set rsData = mfrmDetail.GetStayRecord
    rsData.Sort = "药品id"
    arrSql = Array()
    
    int留存方式 = Val(zlDatabase.GetPara("按月留存领用", glngSys, 模块号.药品领用))
    With rsData
        Str期间 = Format(StrCurDate, IIf(int留存方式 = 0, "yyyy", "yyyymm"))
        .Filter = ""
        
        Do While Not .EOF
            gstrSQL = "ZL_药品留存记录_INSERT("
            '期间
            gstrSQL = gstrSQL & "'" & Str期间 & "'"
            '汇总发药号
            gstrSQL = gstrSQL & "," & mcur汇总发药号
            '库房ID
            gstrSQL = gstrSQL & "," & mcondition.lng药房id
            '药品ID
            gstrSQL = gstrSQL & "," & !药品ID
            '批次
            gstrSQL = gstrSQL & "," & !批次
            '留存数量
            gstrSQL = gstrSQL & "," & !留存数量
            '零售价
            gstrSQL = gstrSQL & "," & !单价
            '填制人
            gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
            '填制日期
            gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
            '领药部门ID
            gstrSQL = gstrSQL & "," & !领药部门ID
            gstrSQL = gstrSQL & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
                
            .MoveNext
        Loop
        
    End With
    
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    For lngRow = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(lngRow)), Me.Caption & "-保存留存")
    Next
    gcnOracle.CommitTrans
    
    DrugStoreWork_StayProc = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DrugStoreWork_CancelVerifyProc(ByVal StrCurDate As String) As Boolean
    '处理销帐数据
    Dim i As Integer
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant
    Dim int审核标志 As Integer
    Dim bln是否有退药 As Boolean
    Dim str序号数量 As String
    Dim lngPre费用id As Long
    Dim str药品id As String
    Dim strPreNo As String
    Dim lngPre费用序号 As Long
    Dim dblSum As Double
    Dim rsData As ADODB.Recordset
    Dim arrSql As Variant
    Dim blnBeginTrans As Boolean
    Dim strWriteOffInfo As String
    Dim strReturnInfo As String
    Dim strReserve As String
    
    arrSql = Array()
    
    '前提条件是汇总销帐记录一并发药
    If mParams.bln汇总发药 = True Then
        '取销账数据集
        Set rsData = mfrmDetail.GetChargeOffRecord
    
        If rsData.State <> 0 Then
            rsData.Filter = "执行标志=1"
            rsData.Sort = "药品id,No,费用id,收发id"
            If rsData.RecordCount > 0 Then
                With rsData
                    '初始化医保部件
                    gclsInsure.InitOracle gcnOracle
                    Do While Not .EOF
                        If !审核标志 = 1 And !销帐数量 <> 0 Then
                            If IsOutPatient(mstrPrivs, !单据, !NO, 2, 2) = False Then Exit Function
                            If IsReceiptBalance_Charge(1, mstrPrivs, !单据, !NO, !费用序号, 2, 2) = False Then Exit Function
                        End If
                
                        If !审核标志 <> 0 Then
                            If lngPre费用id <> !费用ID Then
                                '费用销帐记录处理
                                gstrSQL = "zl_病人费用销帐_Audit("
                                '费用ID
                                gstrSQL = gstrSQL & !费用ID
                                '申请时间
                                gstrSQL = gstrSQL & ",To_Date('" & !申请时间 & "','YYYY-MM-DD HH24:MI:SS')"
                                '审核人
                                gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                                '审核时间
                                gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
                                '审核标志
                                gstrSQL = gstrSQL & "," & !审核标志
                                gstrSQL = gstrSQL & ")"

                                ReDim Preserve arrSql(UBound(arrSql) + 1)
                                arrSql(UBound(arrSql)) = gstrSQL
                
                                lngPre费用id = !费用ID
                                
                                '记录当前销账审核的记录的申请时间和病人ID，用于更新销账消息菜单
                                If strWriteOffInfo = "" Then
                                    strWriteOffInfo = Format(!申请时间, "yyyy-mm-dd hh:mm:ss") & "," & !病人ID
                                ElseIf InStr(strWriteOffInfo & "|", Format(!申请时间, "yyyy-mm-dd hh:mm:ss") & "," & !病人ID & "|") = 0 Then
                                    strWriteOffInfo = strWriteOffInfo & "|" & Format(!申请时间, "yyyy-mm-dd hh:mm:ss") & "," & !病人ID
                                End If
                                
                            End If
                        End If
                        
                        '退药处理
                        If !审核标志 = 1 And !销帐数量 <> 0 Then
                            gstrSQL = "zl_药品收发记录_部门退药("
                            '收发ID
                            gstrSQL = gstrSQL & !收发ID
                            '审核人
                            gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                            '审核时间
                            gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
                            '批号
                            gstrSQL = gstrSQL & "," & IIf(IsNull(!批号), "NULL", IIf(Mid(!批号, 1, 1) = "(", "NULL", "'" & Mid(!批号, 1, 8) & "'"))
                            '效期
                            gstrSQL = gstrSQL & "," & IIf(IsNull(!效期), "NULL", IIf(!效期 = "", "NULL", "To_Date('" & Format(!效期, "yyyy-MM-dd") & "','yyyy-MM-dd')"))
                            '产地
                            gstrSQL = gstrSQL & "," & IIf(IsNull(!产地), "NULL", "'" & !产地 & "'")
                            '退药数量
                            gstrSQL = gstrSQL & "," & !销帐数量
                            '退药库房
                            gstrSQL = gstrSQL & ",NULL"
                            '退药人
                            gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                            '金额保留位数
                            gstrSQL = gstrSQL & "," & mParams.int金额保留位数
                            '门诊
                            gstrSQL = gstrSQL & ",2"
                            '汇总发药号
                            gstrSQL = gstrSQL & "," & mcur汇总发药号
                            gstrSQL = gstrSQL & ")"

                            ReDim Preserve arrSql(UBound(arrSql) + 1)
                            arrSql(UBound(arrSql)) = gstrSQL
                
                            bln是否有退药 = True
                            
                            If InStr("," & str药品id & ",", "," & !药品ID & ",") = 0 Then
                                str药品id = IIf(str药品id = "", "", str药品id & ",") & !药品ID
                            End If
                            
                            strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & Val(!收发ID) & "," & Val(!销帐数量)
                            
                            '销帐处理
                            strPreNo = !NO
                            lngPre费用序号 = !费用序号
                            dblSum = dblSum + !销帐数量
                            
                            .MoveNext
                            If .EOF Then
                                .MovePrevious
                                str序号数量 = !费用序号 & ":" & dblSum
                
                                gstrSQL = "ZL_住院记帐记录_Delete("
                                'NO
                                gstrSQL = gstrSQL & "'" & !NO & "'"
                                '序号，数量串
                                gstrSQL = gstrSQL & ",'" & str序号数量 & "'"
                                '操作员编号
                                gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
                                '操作员姓名
                                gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                                '记录性质
                                gstrSQL = gstrSQL & "," & !记录性质
                                '操作状态
                                gstrSQL = gstrSQL & ",1"
                                gstrSQL = gstrSQL & ")"

                                ReDim Preserve arrSql(UBound(arrSql) + 1)
                                arrSql(UBound(arrSql)) = gstrSQL
                
                                '医保处理
                                If Not IsNull(!险类) And InStr(1, strMCNO, !NO) = 0 Then
                                    MCPAR.记帐作废上传 = gclsInsure.GetCapability(support记帐作废上传, , Val(!险类))
                                    MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, , Val(!险类))
                                    strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !险类 & _
                                            "," & IIf(MCPAR.记帐作废上传, "1", "0") & "," & IIf(MCPAR.记帐完成后上传, "1", "0")
                                End If
                                .MoveNext
                            Else
                                If strPreNo <> !NO Or (strPreNo = !NO And lngPre费用序号 <> !费用序号) Then
                                    .MovePrevious
                                    str序号数量 = !费用序号 & ":" & dblSum
                                    
                                    gstrSQL = "ZL_住院记帐记录_Delete("
                                    'NO
                                    gstrSQL = gstrSQL & "'" & !NO & "'"
                                    '序号，数量串
                                    gstrSQL = gstrSQL & ",'" & str序号数量 & "'"
                                    '操作员编号
                                    gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
                                    '操作员姓名
                                    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                                    '记录性质
                                    gstrSQL = gstrSQL & "," & !记录性质
                                    '操作状态
                                    gstrSQL = gstrSQL & ",1"
                                    gstrSQL = gstrSQL & ")"

                                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                                    arrSql(UBound(arrSql)) = gstrSQL
                    
                                    '医保处理
                                    If Not IsNull(!险类) And InStr(1, strMCNO, !NO) = 0 Then
                                        MCPAR.记帐作废上传 = gclsInsure.GetCapability(support记帐作废上传, , Val(!险类))
                                        MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, , Val(!险类))
                                        strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !险类 & _
                                                "," & IIf(MCPAR.记帐作废上传, "1", "0") & "," & IIf(MCPAR.记帐完成后上传, "1", "0")
                                    End If
                                    
                                    dblSum = 0
                                    .MoveNext
                                End If
                            End If
                            .MovePrevious
                        End If
                        .MoveNext
                    Loop
                End With
                
                '集中处理退药销账事务
                gcnOracle.BeginTrans
                blnBeginTrans = True
                
                For i = 0 To UBound(arrSql)
                    Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "DrugStoreWork_CancelVerifyProc")
                Next
            
                '医保，记帐作废上传，作废时上传
                If strMCNO <> "" Then
                    arrMCRec = Split(strMCNO, "|")
                    For i = 0 To UBound(arrMCRec)
                        arrMCPar = Split(arrMCRec(i), ",")
                        If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                            If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                                gcnOracle.RollbackTrans
                                GoTo errHandle
                            End If
                        End If
                    Next
                End If
                
                gcnOracle.CommitTrans
                blnBeginTrans = False
                
                '医保，记帐作废上传，完成后上传
                If strMCNO <> "" Then
                    For i = 0 To UBound(arrMCRec)
                        arrMCPar = Split(arrMCRec(i), ",")
                        If arrMCPar(2) = 1 And arrMCPar(3) = 1 Then
                            If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                                MsgBox "单据""" & CStr(arrMCPar(0)) & """的销帐数据向医保传送失败，该单据已销帐。", vbInformation, gstrSysName
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End If
    
    '删除消息记录集中已经审核过的消息记录
    If strWriteOffInfo <> "" And Not mobjMipModule Is Nothing Then
        If Not mrsReceiveMsg Is Nothing Then
            If mrsReceiveMsg.RecordCount > 0 Then
                With mrsReceiveMsg
                    .MoveFirst
                    Do While Not .EOF
                        If InStr(strWriteOffInfo & "|", !申请时间 & "," & !病人ID & "|") > 0 Then
                            .Delete adAffectCurrent
                        End If
                        
                        .MoveNext
                    Loop
                End With
                '设置消息菜单
                Call SetMessageBar(mrsReceiveMsg)
            End If
        End If
    End If
    
    '调用退药后的外挂接口
    If Not mobjPlugIn Is Nothing And bln是否有退药 Then
        On Error Resume Next
        mobjPlugIn.DrugReturnByID mlng药房ID, strReturnInfo, CDate(StrCurDate), strReserve
        err.Clear: On Error GoTo 0
    End If
        
    DrugStoreWork_CancelVerifyProc = True
    Exit Function
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FindRow()
    Dim strFind As String
    
    If tbcDetail.Selected.index <> mListType.发药 And tbcDetail.Selected.index <> mListType.退药 Then Exit Sub
    
    TimerAuto.Enabled = False
    strFind = Frm部门发药定位.ShowMe(mcondition.lng药房id, Me, mstrPrivs)
    
    If strFind <> "" Then
        mfrmDetail.FindRecord tbcDetail.Selected.index, strFind
    End If
    
    TimerAuto.Enabled = True
End Sub

Private Sub FindRowNext()
    If tbcDetail.Selected.index <> mListType.发药 And tbcDetail.Selected.index <> mListType.退药 Then Exit Sub
    
    mfrmDetail.FindRecord tbcDetail.Selected.index
End Sub

Private Sub GetCondition()
    '更新条件
    Dim dteTime As Date
    Dim n As Integer
    
    dteTime = Sys.Currentdate
    
    With mcondition
        '药房ID
        .lng药房id = cbo发药药房.ItemData(cbo发药药房.ListIndex)
        
        '时间范围
        Select Case cbo时间范围.ListIndex
            Case mTimeRange.当天
                .str开始时间 = Format(dteTime, "yyyy-mm-dd") & " 00:00:00"
                .str结束时间 = Format(dteTime, "yyyy-mm-dd") & " 23:59:59"
            Case mTimeRange.两天内
                .str开始时间 = Format(DateAdd("d", -1, dteTime), "yyyy-mm-dd") & " 00:00:00"
                .str结束时间 = Format(dteTime, "yyyy-mm-dd") & " 23:59:59"
            Case mTimeRange.三天内
                .str开始时间 = Format(DateAdd("d", -2, dteTime), "yyyy-mm-dd") & " 00:00:00"
                .str结束时间 = Format(dteTime, "yyyy-mm-dd") & " 23:59:59"
            Case mTimeRange.指定时间范围
                .str开始时间 = Format(Dtp开始时间.Value, "yyyy-mm-dd hh:mm:ss")
                .str结束时间 = Format(Dtp结束时间.Value, "yyyy-mm-dd hh:mm:ss")
            Case Else
                .str开始时间 = Format(dteTime, "yyyy-mm-dd") & " 00:00:00"
                .str结束时间 = Format(dteTime, "yyyy-mm-dd") & " 23:59:59"
        End Select
        
        '录入信息
        .str住院号 = ""
        .str姓名 = ""
        .str床号 = ""
        .strNo = ""
        .lng病人ID = -1
        .str就诊卡 = ""
        .str领药号 = ""
        .cur发药号 = 0
        .lng领药部门ID = -1
        .strIC卡 = ""
        
        If Trim(txtInput.Text) <> "" Then
            Select Case mParams.int输入模式
                Case mInputType.住院号
                    If InStr(txtInput.Text, "-") > 0 Then
                        .str住院号 = Mid(Trim(txtInput.Text), 1, InStr(txtInput.Text, "-") - 1)
                    Else
                        .str住院号 = Trim(txtInput.Text)
                    End If
                Case mInputType.姓名
'                    If mblnCard = True Then
'                        .lng病人ID = Val(txtInput.Tag)
'                    Else
'                        .str姓名 = Trim(txtInput.Text)
'                    End If
                    .lng病人ID = Val(txtInput.Tag)
                Case mInputType.床号
                    '由于床号不唯一，转为用病人ID来查询
                    .lng病人ID = Val(txtInput.Tag)
                Case mInputType.NO
                    If InStr(txtInput.Text, "-") > 0 Then
                        .strNo = Mid(Trim(txtInput.Text), 1, InStr(txtInput.Text, "-") - 1)
                    Else
                        .strNo = Trim(txtInput.Text)
                    End If
                Case mInputType.病人ID
                    If InStr(txtInput.Text, "-") > 0 Then
                        .lng病人ID = Mid(Trim(txtInput.Text), 1, InStr(txtInput.Text, "-") - 1)
                    Else
                        .lng病人ID = Val(Trim(txtInput.Text))
                    End If
                Case mInputType.领药号
                    .str领药号 = Trim(txtInput.Text)
                Case mInputType.发药号
                    .cur发药号 = Val(Trim(txtInput.Text))
                Case mInputType.领药部门
                    .lng领药部门ID = Val(txtInput.Tag)
                Case mInputType.IC卡
                    .lng病人ID = Val(txtInput.Tag)
                Case Else
                    '其他的消费卡，返回病人ID
                    .lng病人ID = Val(txtInput.Tag)
            End Select
        End If
        
        '发药类型
        '0-所有,1-不含离院带药,2-仅含离院带药,3-不含自取药,4-仅含自取药,5-院内用药(不包括离院带药和自取药),6-离院带药和自取药
        If chkSend(0).Value = 1 And chkSend(1).Value = 1 And chkSend(2).Value = 1 Then
            .int发药类型 = 0
        ElseIf chkSend(0).Value = 1 And chkSend(2).Value = 1 Then
            .int发药类型 = 1
        ElseIf chkSend(0).Value = 1 And chkSend(1).Value = 1 Then
            .int发药类型 = 3
        ElseIf chkSend(1).Value = 1 And chkSend(2).Value = 1 Then
            .int发药类型 = 6
        ElseIf chkSend(0).Value = 1 Then
            .int发药类型 = 5
        ElseIf chkSend(1).Value = 1 Then
            .int发药类型 = 2
        ElseIf chkSend(2).Value = 1 Then
            .int发药类型 = 4
        End If
        
        '自定义发药类型
        .str其它发药类型 = ""
        If mblnExistOtherSendType = True Then
            For n = 0 To chkSendType.UBound
                If chkSendType(n).Value = 1 Then
                    .str其它发药类型 = IIf(.str其它发药类型 = "", "", .str其它发药类型 & ",") & chkSendType(n).Caption
                End If
            Next
        End If
        
        '给药途径
        If Trim(txt给药途径.Text) = "" Or InStr(Trim(txt给药途径.Text), "所有给药途径") > 0 Then
            .str给药途径 = ""
        Else
            .str给药途径 = Trim(txt给药途径.Text)
        End If
        
        '药品剂型
        If Trim(txt药品剂型.Text) = "" Or InStr(Trim(txt药品剂型.Text), "所有药品剂型") > 0 Then
            .str药品剂型 = ""
        Else
            .str药品剂型 = Trim(txt药品剂型.Text)
        End If
        
        '处理范围
        If Me.opt范围(1).Value = True Then
            .int处理范围 = 1
        ElseIf Me.opt范围(2).Value = True Then
            .int处理范围 = 2
        Else
            .int处理范围 = 0
        End If
        
        '医嘱类型
        .int医嘱类型 = 0
        If Cbo医嘱类型.ListIndex <> -1 Then .int医嘱类型 = Cbo医嘱类型.ListIndex
                
        '病人类型
        If chkType(0).Value = 1 And chkType(1).Value = 1 Then
            .int病人类型 = 2
        ElseIf chkType(1).Value = 1 Then
            .int病人类型 = 1
        Else
            .int病人类型 = 0
        End If
        
        '操作模式
        .int操作模式 = mParams.int操作模式
        
        '记账人
        .str记账人 = mParams.str记帐人
        
        '退药待发
        .bln显示退药待发单据 = mParams.bln待发单据
        
        '所有过程单据
        .bln显示所有退药单据 = mParams.bln过程单据
        
        '领药/退药人
        .bln显示领药退药人 = False
    End With
End Sub

Private Sub GetPrivs()
    '权限
    mPrives.bln所有药房 = IsInString(mstrPrivs, "所有药房", ";")
    mPrives.bln发药 = IsInString(mstrPrivs, "发药", ";")
    mPrives.bln退药 = IsInString(mstrPrivs, "退药", ";")
    mPrives.bln退其它药房的处方 = IsInString(mstrPrivs, "退其它药房的处方", ";")
    mPrives.bln发退结帐处方 = IsInString(mstrPrivs, "发退结帐处方", ";")
    mPrives.bln发退出院病人处方 = IsInString(mstrPrivs, "发退出院病人处方", ";")
    mPrives.bln缺药申领 = IsInString(mstrPrivs, "缺药申领", ";")
    mPrives.bln拒发 = IsInString(mstrPrivs, "拒发", ";")
    mPrives.bln合理用药监测 = IsInString(mstrPrivs, "合理用药监测", ";")
    mPrives.bln医生查询 = IsInString(mstrPrivs, "医生查询", ";")
    mPrives.bln参数设置 = IsInString(mstrPrivs, "参数设置", ";")
    mPrives.bln退药销帐 = IsInString(mstrPrivs, "退药销帐", ";")
    mPrives.bln修改留存数量 = IsInString(mstrPrivs, "修改留存数量", ";")
    mPrives.bln停止发药 = IsInString(mstrPrivs, "停止发药", ";")
    mPrives.bln恢复发药 = IsInString(mstrPrivs, "恢复发药", ";")
    mPrives.bln查看已发药清单 = IsInString(mstrPrivs, "查看已发药清单", ";")
    mPrives.bln过滤时间 = IsInString(mstrPrivs, "过滤时间", ";")
    mPrives.bln电子病案查阅 = IsInString(mstrPrivs, "电子病案查阅", ";")
End Sub

Private Sub GetDeptListRecord(ByVal rsData As ADODB.Recordset)
    Set mrsDeptList = New ADODB.Recordset
    
    With mrsDeptList
        If .State = 1 Then .Close
        
        .Fields.Append "科室ID", adDouble, 18, adFldIsNullable                  '领药部门ID
        .Fields.Append "科室名称", adLongVarChar, 50, adFldIsNullable                  '领药部门名称
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable                  '药品收发记录NO号
        .Fields.Append "收发ID", adDouble, 18, adFldIsNullable                  '药品收发记录ID
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable                  '药品收发记录药品ID
        .Fields.Append "执行状态", adDouble, 1, adFldIsNullable
        .Fields.Append "病人id", adDouble, 18, adFldIsNullable
        .Fields.Append "病人姓名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "领药号", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        If mParams.bln启用审方 And tbcDetail.Selected.index = mListType.发药 Then
            rsData.Filter = "(审查结果=1 and 审查id<>0) or 审查id=0"
        Else
            rsData.Filter = ""
        End If
        
        If rsData.RecordCount <> 0 Then rsData.MoveFirst
        Do While Not rsData.EOF
            .AddNew
            !科室ID = rsData!Id
            !科室名称 = rsData!科室名称
            !NO = rsData!NO
            !收发ID = rsData!收发ID
            !药品ID = rsData!药品ID
            !病人ID = rsData!病人ID
            !病人姓名 = rsData!姓名
            If Me.tbcDetail.Selected.index = mListType.发药 Or Me.tbcDetail.Selected.index = mListType.汇总 Or Me.tbcDetail.Selected.index = mListType.拒发 Then
                !领药号 = rsData!领药号
            End If
            
            !执行状态 = 0
            
            .Update
            
            rsData.MoveNext
        Loop
    End With
End Sub
Private Sub GetSendDeptTreeView(ByRef rsData As ADODB.Recordset)
    '刷新待发药部门树表
    Dim objNode As Node
    Dim objItem As listItem
    Dim lng当前科室 As Long
    Dim str当前科室 As String
    Dim str当前领药号 As String
    Dim str当前病人Key As String
    Dim lng当前病人ID As Long
    Dim str当前病人姓名 As String
    Dim str当前NO As String
    Dim int科室药品数 As Integer
    Dim lng当前药品 As Long
    Dim strType As String
    Dim arr科室 As Variant
    Dim i As Integer
    Dim j As Integer
    Dim count As Integer

    If rsData.EOF Then
        Set objNode = tvwList(mDeptType.发药).Nodes.Add(, , "_无数据", "未找到满足条件的记录")
        tvwList(mDeptType.发药).Checkboxes = False
        tvwList(mDeptType.发药).Tag = "0"
        chkAll(mDeptType.发药).Enabled = False

        mfrmDetail.ClearList mListType.发药
        Exit Sub
    End If
    
    '根据数据集的结果组织树表，按科室（包含药品种类）、领药号（如无领药号则没有这级）、病人、单据号显示各级
    
    chkAll(mDeptType.发药).Enabled = True
    tvwList(mDeptType.发药).Checkboxes = True
    arr科室 = Array()
    With tvwList(mDeptType.发药)
        If Not rsData.EOF Then
            '记录所有科室名称
            rsData.Sort = "科室名称,ID"
            
            If mParams.blnSort Then
                Set rstemp = New Recordset
                With rstemp
                    If .State = 1 Then .Close
                    .Fields.Append "ID", adDouble, 18, adFldIsNullable
                    .Fields.Append "科室名称", adLongVarChar, 40, adFldIsNullable
                    .Fields.Append "发送时间", adLongVarChar, 40, adFldIsNullable
                    
                    .CursorLocation = adUseClient
                    .CursorType = adOpenStatic
                    .LockType = adLockOptimistic
                    .Open
                End With
            End If
            
            Do While Not rsData.EOF
                If lng当前科室 <> rsData!Id Then
                    lng当前科室 = rsData!Id
                    If mParams.blnSort Then
                        rstemp.AddNew
                        rstemp!Id = rsData!Id
                        rstemp!科室名称 = rsData!科室名称
                        rstemp!发送时间 = Format(rsData!填制日期, "yyyy-mm-dd hh:mm:ss")
                    Else
                        ReDim Preserve arr科室(UBound(arr科室) + 1)
                        arr科室(UBound(arr科室)) = lng当前科室 & "|" & rsData!科室名称
                    End If
                End If
                rsData.MoveNext
            Loop
            
            If mParams.blnSort Then
                rstemp.Sort = "发送时间"
                rstemp.MoveFirst
                Do While Not rstemp.EOF
                    ReDim Preserve arr科室(UBound(arr科室) + 1)
                    arr科室(UBound(arr科室)) = rstemp!Id & "|" & rstemp!科室名称
                    rstemp.MoveNext
                Loop
            End If
            
            
            '按科室组织数据树表
            For i = 0 To UBound(arr科室)
                If mParams.bln启用审方 Then
                    rsData.Filter = "(审查结果=1 and 审查id<>0 and ID= '" & Split(arr科室(i), "|")(0) & "') or (审查id=0 and ID=' " & Split(arr科室(i), "|")(0) & "')"
                Else
                    rsData.Filter = "ID= '" & Split(arr科室(i), "|")(0) & "' "
                End If
                
                '计算当前科室药品种类
                rsData.Sort = "药品ID"
                lng当前药品 = 0
                int科室药品数 = 0
                Do While Not rsData.EOF
                    If lng当前药品 <> rsData!药品ID Then
                        int科室药品数 = int科室药品数 + 1
                        lng当前药品 = rsData!药品ID
                    End If
                    rsData.MoveNext
                Loop
                
                Set objNode = .Nodes.Add(, , "D_" & Split(arr科室(i), "|")(0), Split(arr科室(i), "|")(1) & "（" & int科室药品数 & "种药品待发）", 1)
                objNode.Expanded = False
                
                If mParams.blnOnlyShowDept = False Then
                    '先过滤有领药号的记录
                    If mParams.bln启用审方 Then
                        rsData.Filter = "(审查结果=1 and 审查id<>0 And 领药号<>0 and ID= '" & Split(arr科室(i), "|")(0) & "') or (审查id=0 And 领药号<>0 and ID=' " & Split(arr科室(i), "|")(0) & "')"
                    Else
                        rsData.Filter = "ID= '" & Split(arr科室(i), "|")(0) & "' And 领药号<>0"
                    End If

                    If mParams.int病人排序 = 1 Then
                        rsData.Sort = "领药号,姓名,病人ID,婴儿费,NO"
                    Else
                        rsData.Sort = "领药号,床号,姓名,病人ID,婴儿费,NO"
                    End If
                    
                    str当前领药号 = ""
                    str当前病人Key = ""
                    str当前NO = ""
                    str当前病人姓名 = ""
                    Do While Not rsData.EOF
                        If str当前领药号 <> rsData!领药号 Then
                            str当前领药号 = rsData!领药号
                            str当前病人Key = ""        '不同领药号可能存在相同的病人，当领药号不同时，初始病人信息要清空
                            
                            Set objNode = .Nodes.Add("D_" & Split(arr科室(i), "|")(0), 4, "R_" & Split(arr科室(i), "|")(0) & str当前领药号, str当前领药号, 2)
                            objNode.Expanded = False
                            objNode.Tag = str当前领药号 & "|" & Split(arr科室(i), "|")(0)
                        End If
                        
                        If str当前病人Key & lng当前病人ID <> rsData!姓名 & "(" & IIf(IsNull(rsData!床号), "", rsData!床号 & "床 ") & rsData!性别 & " " & rsData!年龄 & ")" & rsData!病人ID Then
                            If IIf(IsNull(rsData!床号), "", rsData!床号) <> "" Then
                                str当前病人Key = rsData!姓名 & "(" & rsData!床号 & "床 " & rsData!性别 & " " & rsData!年龄 & ")"
                            Else
                                str当前病人Key = rsData!姓名 & "(" & rsData!性别 & " " & rsData!年龄 & ")"
                            End If
                            lng当前病人ID = rsData!病人ID
                            
                            Set objNode = .Nodes.Add("R_" & Split(arr科室(i), "|")(0) & str当前领药号, 4, "P_" & Split(arr科室(i), "|")(0) & str当前领药号 & str当前病人Key & rsData!病人ID, str当前病人Key, 3)
                            objNode.ForeColor = IIf(IsNull(rsData!颜色), vbBlack, rsData!颜色)
                            objNode.Tag = rsData!病人ID & "|R" & str当前领药号 & "|" & rsData!姓名
                            objNode.Expanded = False
                        End If
                        
                        If ((str当前NO <> rsData!NO) Or (str当前NO = rsData!NO And str当前病人姓名 <> rsData!姓名)) Then
                            str当前NO = rsData!NO
                            str当前病人姓名 = rsData!姓名  '用于区分母亲和婴儿为同一张单据NO的情况

                            strType = IIf(zlStr.NVL(rsData!医嘱序号, 0) = 0, IIf(rsData!门诊标志 = 1 Or rsData!门诊标志 = 4, "门诊记帐单", IIf(rsData!单据 = 9, "住院记帐单", "住院记帐表")), IIf(IsNull(rsData!扣率) = True, "住院记帐单", IIf(rsData!扣率 Like "0*", "长嘱", IIf(rsData!扣率 Like "1*", "临嘱", "记帐表"))))
                            strType = strType & " " & Format(rsData!填制日期, "mm-dd hh:mm:ss")
                            
                            Set objNode = .Nodes.Add("P_" & Split(arr科室(i), "|")(0) & str当前领药号 & str当前病人Key & rsData!病人ID, 4, "N" & str当前病人Key & rsData!病人ID & "_" & str当前NO & "_" & rsData!姓名, str当前NO & "(" & strType & ")", 4)
                            objNode.Expanded = False
                            objNode.Tag = rsData!NO & "|" & rsData!病人ID & "|" & rsData!姓名
                            If rsData!拒发 = 1 Then
                                objNode.ForeColor = vbRed
                                objNode.Text = objNode.Text & "(已拒发)"
                            End If
                        End If
                        
                        rsData.MoveNext
                    Loop
                    
                    '处理无领药号的记录（无领药号就没有领药号这级）
                    If mParams.bln启用审方 Then
                        rsData.Filter = "(审查结果=1 and 审查id<>0 And 领药号=0 and ID= '" & Split(arr科室(i), "|")(0) & "') or (审查id=0 And 领药号=0 and ID=' " & Split(arr科室(i), "|")(0) & "')"
                    Else
                        rsData.Filter = "ID= '" & Split(arr科室(i), "|")(0) & "' And 领药号=0"
                    End If
                    If mParams.int病人排序 = 1 Then
                        rsData.Sort = "姓名,病人ID,婴儿费,NO"
                    Else
                        rsData.Sort = "床号,姓名,病人ID,婴儿费,NO"
                    End If

                    str当前病人Key = ""
                    str当前NO = ""
                    str当前病人姓名 = ""
                    Do While Not rsData.EOF
                        If str当前病人Key & lng当前病人ID <> rsData!姓名 & "(" & IIf(IsNull(rsData!床号), "", rsData!床号 & "床 ") & rsData!性别 & " " & rsData!年龄 & ")" & rsData!病人ID Then
                            If IIf(IsNull(rsData!床号), "", rsData!床号) <> "" Then
                                str当前病人Key = rsData!姓名 & "(" & rsData!床号 & "床 " & rsData!性别 & " " & rsData!年龄 & ")"
                            Else
                                str当前病人Key = rsData!姓名 & "(" & rsData!性别 & " " & rsData!年龄 & ")"
                            End If
                            lng当前病人ID = rsData!病人ID
                            
                            Set objNode = .Nodes.Add("D_" & Split(arr科室(i), "|")(0), 4, "P_" & Split(arr科室(i), "|")(0) & str当前病人Key & rsData!病人ID, str当前病人Key, 3)
                            objNode.ForeColor = IIf(IsNull(rsData!颜色), vbBlack, rsData!颜色)
                            objNode.Tag = rsData!病人ID & "|" & Split(arr科室(i), "|")(0) & "|" & rsData!姓名
                            objNode.Expanded = False
                            
                        End If
                        
                        If ((str当前NO <> rsData!NO) Or (str当前NO = rsData!NO And str当前病人姓名 <> rsData!姓名)) Then
                            str当前NO = rsData!NO
                            str当前病人姓名 = rsData!姓名  '用于区分母亲和婴儿为同一张单据NO的情况
                            
                            strType = IIf(zlStr.NVL(rsData!医嘱序号, 0) = 0, IIf(rsData!门诊标志 = 1 Or rsData!门诊标志 = 4, "门诊记帐单", IIf(rsData!单据 = 9, "住院记帐单", "住院记帐表")), IIf(IsNull(rsData!扣率) = True, "住院记帐单", IIf(rsData!扣率 Like "0*", "长嘱", IIf(rsData!扣率 Like "1*", "临嘱", "记帐表"))))
                            strType = strType & " " & Format(rsData!填制日期, "mm-dd hh:mm:ss")
                            Set objNode = .Nodes.Add("P_" & Split(arr科室(i), "|")(0) & str当前病人Key & rsData!病人ID, 4, "N" & str当前病人Key & rsData!病人ID & "_" & str当前NO & "_" & rsData!姓名, str当前NO & "(" & strType & ")", 4)
                            objNode.Expanded = False
                            objNode.Tag = rsData!NO & "|" & rsData!病人ID & "|" & rsData!姓名
                            If rsData!拒发 = 1 Then
                                objNode.ForeColor = vbRed
                                objNode.Text = objNode.Text & "(已拒发)"
                            End If
                        End If
                        
                        rsData.MoveNext
                    Loop
                End If
            Next
        End If
    End With
End Sub
Private Sub GetReturnDeptTreeView(ByRef rsData As ADODB.Recordset)
    '刷新退药部门树表
    Dim objNode As Node
    Dim objItem As listItem
    Dim lng科室ID As Long
    Dim str当前病人Key As String
    Dim lng当前病人ID As Long
    Dim str当前病人姓名 As String
    Dim str当前NO As String
    Dim strType As String
    
    Dim arr科室 As Variant
    Dim i As Integer
    Dim j As Integer
    Dim count As Integer

    If rsData.EOF Then
        Set objNode = tvwList(mDeptType.退药).Nodes.Add(, , "_无数据", "未找到满足条件的记录")
        tvwList(mDeptType.退药).Checkboxes = False
        tvwList(mDeptType.退药).Tag = "0"
        chkAll(mDeptType.退药).Enabled = False

        mfrmDetail.ClearList mListType.退药
        Exit Sub
    End If
    
    '根据数据集的结果组织树表，两种方式：
    '1、按发药号、科室（包含药品种类）、病人、单据号显示各级；
    '2、按科室（包含药品种类）、病人、单据号显示各级；
    
    chkAll(mDeptType.退药).Enabled = True
    tvwList(mDeptType.退药).Checkboxes = True
    arr科室 = Array()
    With tvwList(mDeptType.退药)
        If Not rsData.EOF Then
            '记录所有科室名称
            rsData.Sort = "科室名称,ID"
            Do While Not rsData.EOF
                If lng科室ID <> rsData!Id Then
                    ReDim Preserve arr科室(UBound(arr科室) + 1)
                    lng科室ID = rsData!Id
                    arr科室(UBound(arr科室)) = rsData!Id & "|" & rsData!科室名称
                End If
                rsData.MoveNext
            Loop
    
            '按科室组织数据树表
            For i = 0 To UBound(arr科室)
                rsData.Filter = "ID= '" & Split(arr科室(i), "|")(0) & "' "
                
                '计算当前科室药品种类
                rsData.Sort = "药品ID"
                
                Set objNode = .Nodes.Add(, , "D_" & Split(arr科室(i), "|")(0), Split(arr科室(i), "|")(1), 1)
                objNode.Expanded = False
                
                If mParams.blnOnlyShowDept = False Then
                    rsData.Filter = "ID= '" & Split(arr科室(i), "|")(0) & "'"
                    If mParams.int病人排序 = 1 Then
                        rsData.Sort = "姓名,病人ID,婴儿费,NO"
                    Else
                        rsData.Sort = "床号,姓名,病人ID,婴儿费,NO"
                    End If
                    str当前病人Key = ""
                    str当前NO = ""
                    str当前病人姓名 = ""
                    Do While Not rsData.EOF
                        If str当前病人Key & lng当前病人ID <> rsData!姓名 & "(" & IIf(IsNull(rsData!床号), "", rsData!床号 & "床 ") & rsData!性别 & " " & rsData!年龄 & ")" & rsData!病人ID Then
                            If IIf(IsNull(rsData!床号), "", rsData!床号) <> "" Then
                                str当前病人Key = rsData!姓名 & "(" & rsData!床号 & "床 " & rsData!性别 & " " & rsData!年龄 & ")"
                            Else
                                str当前病人Key = rsData!姓名 & "(" & rsData!性别 & " " & rsData!年龄 & ")"
                            End If
                            lng当前病人ID = rsData!病人ID
                            
                            Set objNode = .Nodes.Add("D_" & Split(arr科室(i), "|")(0), 4, "P_" & Split(arr科室(i), "|")(0) & str当前病人Key & rsData!病人ID, str当前病人Key, 3)
                            objNode.ForeColor = IIf(IsNull(rsData!颜色), vbBlack, rsData!颜色)
                            objNode.Tag = rsData!病人ID & "|D" & Split(arr科室(i), "|")(0) & "|" & rsData!姓名
                            objNode.Expanded = False
                        End If
                        
                        If (str当前NO <> rsData!NO) Or (str当前NO = rsData!NO And str当前病人姓名 <> rsData!姓名) Then
                            str当前NO = rsData!NO
                            str当前病人姓名 = rsData!姓名  '用于区分母亲和婴儿为同一张单据NO的情况

                            strType = IIf(zlStr.NVL(rsData!医嘱序号, 0) = 0, IIf(rsData!门诊标志 = 1 Or rsData!门诊标志 = 4, "门诊记帐单", IIf(rsData!单据 = 9, "住院记帐单", "住院记帐表")), IIf(IsNull(rsData!扣率) = True, "住院记帐单", IIf(rsData!扣率 Like "0*", "长嘱", IIf(rsData!扣率 Like "1*", "临嘱", "记帐表"))))
                            strType = strType & " " & Format(rsData!填制日期, "mm-dd hh:mm:ss")
                            
                            Set objNode = .Nodes.Add("P_" & Split(arr科室(i), "|")(0) & str当前病人Key & rsData!病人ID, 4, "N" & str当前病人Key & rsData!病人ID & "_" & str当前NO & Split(arr科室(i), "|")(0) & "_" & rsData!姓名, str当前NO & "(" & strType & ")", 4)
                            objNode.Tag = rsData!NO & "|" & rsData!病人ID & "|" & rsData!姓名
                            objNode.Expanded = False
                        End If
                        
                        rsData.MoveNext
                    Loop
                End If
            Next
        End If
    End With
    
End Sub
Private Function GetDrugFormat() As Integer
    Dim strSave As String
    Dim arrColumn
    
    '取得药品名称的格式方式
    strSave = zlDatabase.GetPara("列设置", glngSys, 1342)
    If strSave = "" Then strSave = "0|药品名称,0|其它名,0|英文名,0|科室,0|开单医生,0|状态,0|类型,0|NO,0|记帐员,0|床号,0|姓名,0|住院号,0|规格,0|产地,0|批号,0|付,0|数量,0|已退数,0|准退数,0|退药数,0|单价,0|金额,0|单量,0|频次,0|用法,0|记帐时间,0|说明,0|操作员,0|发药时间,0|领/退药人,0|库房货位"
    arrColumn = Split(strSave, ",")
    GetDrugFormat = Val(Split(arrColumn(0), "|")(0))
End Function

Private Sub ReturnSelected给药途径(ByVal intType As Integer)
    'intType:0-双击给药途径列表时；1-给药途径列表中按回车时
    Dim n As Integer
    
    With Lvw给药途径
        If .SelectedItem Is Nothing Then Exit Sub
        Me.txt给药途径.Tag = ""
        Me.txt给药途径.Text = ""
        
        '如果选择了全选，则不用取所有给药途径了
        If .ListItems(1).Checked Then
            Me.txt给药途径.Tag = ""
            Me.txt给药途径.Text = "所有给药途径"
            .Visible = False
            Exit Sub
        End If
        For n = 1 To .ListItems.count
            If .ListItems(n).Checked Then
                Me.txt给药途径.Tag = IIf(Me.txt给药途径.Tag = "", Mid(.ListItems(n).Key, 2), Me.txt给药途径.Tag & "," & Mid(.ListItems(n).Key, 2))
                Me.txt给药途径.Text = IIf(Me.txt给药途径.Text = "", .ListItems(n).Text, Me.txt给药途径.Text & "," & .ListItems(n).Text)
            End If
        Next
        
        If intType = 0 Then
            '如果当前双击的给药途径未被选上，将当前双击的给药途径也加入到编辑框中
            If .SelectedItem.Checked = False Then
                .SelectedItem.Checked = True
                Me.txt给药途径.Tag = IIf(Me.txt给药途径.Tag = "", Mid(.SelectedItem.Key, 2), Me.txt给药途径.Tag & "," & Mid(.SelectedItem.Key, 2))
                Me.txt给药途径.Text = IIf(Me.txt给药途径.Text = "", .SelectedItem.Text, Me.txt给药途径.Text & "," & .SelectedItem.Text)
            End If
            
            '如果选择了全选，则不用取所有给药途径了
            If .ListItems(1).Checked Then
                Me.txt给药途径.Tag = ""
                Me.txt给药途径.Text = "所有给药途径"
                .Visible = False
                Exit Sub
            End If
        End If
        
        .Visible = False
    End With
End Sub

Private Sub ReturnSelected剂型(ByVal intType As Integer)
    'intType:0-双击剂型列表时；1-剂型列表中按回车时
    Dim n As Integer
    
    With Lvw药品剂型
        If .SelectedItem Is Nothing Then Exit Sub
        Me.txt药品剂型.Text = ""
        
        '如果选择了全选，则不用取所有给药途径了
        If .ListItems(1).Checked Then
             Me.txt药品剂型.Text = "所有药品剂型"
            .Visible = False
            Exit Sub
        End If
        
        For n = 1 To .ListItems.count
            If .ListItems(n).Checked Then
                Me.txt药品剂型.Text = IIf(Me.txt药品剂型.Text = "", Mid(.ListItems(n).Text, InStr(1, .ListItems(n).Text, "-") + 1), Me.txt药品剂型.Text & "," & Mid(.ListItems(n).Text, InStr(1, .ListItems(n).Text, "-") + 1))
            End If
        Next
        
        If intType = 0 Then
            '如果当前双击的给药途径未被选上，将当前双击的给药途径也加入到编辑框中
            If .SelectedItem.Checked = False Then
                .SelectedItem.Checked = True
                Me.txt药品剂型.Text = IIf(Me.txt药品剂型.Text = "", Mid(.SelectedItem.Text, InStr(1, .SelectedItem.Text, "-") + 1), Me.txt药品剂型.Text & "," & Mid(.SelectedItem.Text, InStr(1, .SelectedItem.Text, "-") + 1))
            End If
            
            If .ListItems(1).Checked Then
                 Me.txt药品剂型.Text = "所有药品剂型"
                .Visible = False
                Exit Sub
            End If
        End If
        
        .Visible = False
    End With
End Sub

Private Sub InitSendRec()
    Set mrsSendData = New ADODB.Recordset
    With mrsSendData
        If .State = 1 Then .Close
        
        '该记录对应的单据信息
        .Fields.Append "收发ID", adDouble, 18, adFldIsNullable              '药品收发ID
        .Fields.Append "序号", adDouble, 18, adFldIsNullable                '药品收发序号
        .Fields.Append "记录状态", adDouble, 2, adFldIsNullable             '药品收发记录的记录状态
        .Fields.Append "类型", adLongVarChar, 20, adFldIsNullable           '记账单、记账表、医嘱（临嘱、长嘱）
        .Fields.Append "扣率", adDouble, 2, adFldIsNullable                  '发药类型列，显示院内用药，离院带药，自取药等内容
        .Fields.Append "已收费", adDouble, 2, adFldIsNullable               '是否已收费：1－已收费
        .Fields.Append "单据", adDouble, 18, adFldIsNullable                '药品收发单据类型：8－门诊收费单；9－住院记账单；10－住院记账表
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable              '药品收发NO号
        .Fields.Append "记帐员", adLongVarChar, 20, adFldIsNullable         '住院费用记录中的操作员
        .Fields.Append "说明", adLongVarChar, 40, adFldIsNullable           '药品收发记录中摘要
        .Fields.Append "记帐时间", adLongVarChar, 20, adFldIsNullable       '住院费用记录中的登记时间
        .Fields.Append "配药人", adLongVarChar, 20, adFldIsNullable         '药品收发记录中配药人
        .Fields.Append "审核人", adLongVarChar, 20, adFldIsNullable         '住院费用记录中的操作员
        .Fields.Append "主页ID", adDouble, 18, adFldIsNullable              '住院费用记录中的主页ID
        .Fields.Append "费用序号", adDouble, 18, adFldIsNullable
    
        .Fields.Append "审查结果", adDouble, 18, adFldIsNullable            '来自“病人医嘱记录”的“审查结果”，用于合理用药（PASS）
        .Fields.Append "医嘱id", adDouble, 18, adFldIsNullable              '“病人医嘱记录”的ID或“住院费用记录”的“医嘱序号”
        .Fields.Append "相关id", adDouble, 18, adFldIsNullable              '来自“病人医嘱记录”的“相关ID”，用于分组
        .Fields.Append "抗生素", adDouble, 1, adFldIsNullable
        .Fields.Append "用药目的", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "用药理由", adLongVarChar, 1000, adFldIsNullable
        .Fields.Append "用药次数", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "皮试结果", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "开嘱时间", adLongVarChar, 20, adFldIsNullable
        
        .Fields.Append "科室ID", adDouble, 18, adFldIsNullable              '来自“住院费用记录”的“病人科室ID”
        .Fields.Append "领药部门编码", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "领药部门", adLongVarChar, 50, adFldIsNullable       '来时“药品收发记录”的“对方部门ID”对应的部门
        .Fields.Append "领药部门ID", adDouble, 18, adFldIsNullable          '来时“药品收发记录”的“对方部门ID”
        .Fields.Append "领药号", adLongVarChar, 20, adFldIsNullable
        
        .Fields.Append "医生嘱托", adLongVarChar, 40, adFldIsNullable
        
        '病人信息
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "性别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "住院号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "科室", adLongVarChar, 50, adFldIsNullable           '病人科室
        .Fields.Append "床号", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "开单医生", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "开单部门id", adDouble, 18, adFldIsNullable
        .Fields.Append "年龄", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "病人类型", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "颜色", adDouble, 18, adFldIsNullable
        
        '药品信息
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "药名ID", adDouble, 18, adFldIsNullable
        .Fields.Append "品名", adLongVarChar, 50, adFldIsNullable           '药品名称：0－药品编码与名称；1－药品编码；2－药品名称
        .Fields.Append "其它名", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "英文名", adLongVarChar, 80, adFldIsNullable         '来自“诊疗项目别名”，可考虑优化
        .Fields.Append "配方名称", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "原产地", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "效期", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "分批", adDouble, 2, adFldIsNullable                 '分批属性，来自“药品规格”
        .Fields.Append "付", adLongVarChar, 50, adFldIsNullable             '中药付数，来自“药品收发记录”
        .Fields.Append "数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "单价", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "金额", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "单量", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "原始单量", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "频次", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "实际数量", adDouble, 18, adFldIsNullable            '最小单位的数量，判断库存用
        .Fields.Append "单量单位", adLongVarChar, 20, adFldIsNullable       '来自“诊疗项目目录”的“计算单位”，用于合理用药（PASS）
        .Fields.Append "用法", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "药品编码和名称", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "药品编码", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "药品名称", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "包装", adDouble, 18, adFldIsNullable
        .Fields.Append "发药数量", adDouble, 18, adFldIsNullable            '收发记录中的实际数量，用来比较库存
        .Fields.Append "高危药品", adDouble, 2, adFldIsNullable
        .Fields.Append "是否皮试", adDouble, 2, adFldIsNullable
        .Fields.Append "禁忌药品说明", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "药师审核标志", adDouble, 18, adFldIsNullable
        .Fields.Append "执行分类", adDouble, 18, adFldIsNullable
        
        .Fields.Append "类别", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "零差价管理", adDouble, 1, adFldIsNullable
        
        '药品毒麻、价值信息（综合过滤条件“药品剂型”可考虑优化）
        .Fields.Append "毒理分类", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "价值分类", adLongVarChar, 10, adFldIsNullable
        
        '来自“药品储备限额”和“药品库存”，根据参数“药品储备”可考虑优化
        .Fields.Append "库房货位", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "库存下限", adDouble, 18, adFldIsNullable
        
        .Fields.Append "留存数量", adDouble, 18, adFldIsNullable            '记录科室对该药品的计划留存数量，可根据参数优化
        
        .Fields.Append "退药人", adLongVarChar, 20, adFldIsNullable         '单独关联的一段SQL，可以根据条件优化
        
        .Fields.Append "位置", adDouble, 18, adFldIsNullable                '用于定位
        .Fields.Append "状态", adLongVarChar, 10, adFldIsNullable           '状态：发药、拒发、不处理
        .Fields.Append "执行状态", adDouble, 1, adFldIsNullable             '状态的内部标识：0－缺药；1－发药；2－拒发；3－不处理
        
        .Fields.Append "库存数量", adDouble, 18, adFldIsNullable
        
        '用于电子签名
        .Fields.Append "入出类别id", adDouble, 18, adFldIsNullable
        .Fields.Append "入出系数", adDouble, 18, adFldIsNullable
        .Fields.Append "填制人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "填制日期", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "配药日期", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set mrsChargeOff = New ADODB.Recordset
    With mrsChargeOff
        If .State = 1 Then .Close
        .Fields.Append "领药部门", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "领药部门ID", adDouble, 18, adFldIsNullable
        .Fields.Append "单据", adDouble, 18, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "申请时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "收发序号", adDouble, 18, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "原产地", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "效期", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "准退数量", adDouble, 18, adFldIsNullable
        .Fields.Append "销帐数量", adDouble, 18, adFldIsNullable
        .Fields.Append "包装", adDouble, 18, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "收发ID", adDouble, 18, adFldIsNullable
        .Fields.Append "主页ID", adDouble, 18, adFldIsNullable
        .Fields.Append "费用序号", adDouble, 18, adFldIsNullable
        .Fields.Append "险类", adDouble, 18, adFldIsNullable
        .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
        .Fields.Append "记录性质", adDouble, 18, adFldIsNullable
        .Fields.Append "审核标志", adDouble, 18, adFldIsNullable
        .Fields.Append "药品名称", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "执行标志", adDouble, 2, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "医生嘱托", adLongVarChar, 40, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
        
    Set mrsChargeOffMain = New ADODB.Recordset
    With mrsChargeOffMain
        If .State = 1 Then .Close
        .Fields.Append "领药部门ID", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "申请时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "准退数量", adDouble, 18, adFldIsNullable
        .Fields.Append "销帐数量", adDouble, 18, adFldIsNullable
        .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub LoadCustomSet()
    Dim str发药类型 As String
    Dim intSendType As Integer
    Dim n As Integer
   
    mParams.blnShowReject = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "显示拒发", "0")) = 1)
    mParams.blnSort = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "按发送时间排序", "0")) = 1)
    mParams.blnOnlyShowDept = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "科室列表", "0")) = 1)
    mParams.intShowDept = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "显示科室", "0"))
    mParams.int病人排序 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "病人排序", "1"))
    mParams.int输入模式索引 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "输入模式", 0))
    If mParams.int输入模式索引 < 0 Then
        mParams.int输入模式索引 = 0
    End If
    
    mParams.intAdviceType = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理", "医嘱类型", "0"))
    
    If mParams.intAdviceType >= 0 And mParams.intAdviceType <= Cbo医嘱类型.ListCount - 1 Then
        Cbo医嘱类型.ListIndex = mParams.intAdviceType
    End If
    
    '其他发药类型
    str发药类型 = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理", "发药类型", "")
    If mblnExistOtherSendType = True And str发药类型 <> "" Then
        For n = 0 To chkSendType.UBound
            If InStr(1, "," & str发药类型 & ",", "," & chkSendType(n).Caption & ",") > 0 Then
                chkSendType(n).Value = 1
            End If
        Next
        picShowSendType_Click
    End If
    
    '发药类型
    intSendType = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理", "原发药类型", "0"))
    
    chkSend(0).Value = 0
    chkSend(1).Value = 0
    chkSend(2).Value = 0
    
    If intSendType < 0 Or intSendType > 6 Then
        intSendType = 0
    End If
    
    If intSendType = 0 Then
        chkSend(0).Value = 1
        chkSend(1).Value = 1
        chkSend(2).Value = 1
    ElseIf intSendType = 1 Then
        chkSend(0).Value = 1
        chkSend(2).Value = 1
    ElseIf intSendType = 3 Then
        chkSend(0).Value = 1
        chkSend(1).Value = 1
    ElseIf intSendType = 6 Then
        chkSend(1).Value = 1
        chkSend(2).Value = 1
    ElseIf intSendType = 5 Then
        chkSend(0).Value = 1
    ElseIf intSendType = 2 Then
        chkSend(1).Value = 1
    ElseIf intSendType = 4 Then
        chkSend(2).Value = 1
    End If
End Sub
Private Sub RefreshChargeOffDetail()
    '更新冲销申请明细
    Dim strSubUnit As String
    Dim rstemp As ADODB.Recordset
    Dim strCon As String
    Dim strTmpCon As String
    Dim str申请时间 As String
    Dim lng领药部门ID As Long
    
    Dim lng费用id As Long
    Dim dbl准退数量 As Double
    
    Dim str科室ID串 As String
    Dim lng科室 As Long
    Dim str药品ID串 As String
    Dim lng药品id As Long
    Dim strSql药名 As String
    
    '要有相应权限和参数时才能进行销账处理
    If mPrives.bln退药销帐 = False Or mParams.bln汇总发药 = False Then Exit Sub
    
    With mrsDeptList
        .Filter = ""
        .Sort = "科室ID,药品ID"
        Do While Not .EOF
            If !执行状态 = 1 Then
                If lng科室 <> !科室ID Then
                    lng科室 = !科室ID
                    str科室ID串 = str科室ID串 & IIf(str科室ID串 = "", "", ",") & !科室ID
                End If
                
                If lng药品id <> !药品ID Then
                    lng药品id = !药品ID
                    str药品ID串 = str药品ID串 & IIf(str药品ID串 = "", "", ",") & !药品ID
                End If
            End If
            
            .MoveNext
        Loop
    End With
    If str科室ID串 = "" Then Exit Sub
        
    '单位，包装换算
    Select Case mParams.strUnit
    Case "售价单位"
        strSubUnit = "X.计算单位 单位,1 包装,C.实际数量 As 准退数量,A.数量 As 销帐数量"
    Case "门诊单位"
        strSubUnit = "D.门诊单位 单位,D.门诊包装 包装,C.实际数量 As 准退数量,A.数量 As 销帐数量"
    Case "住院单位"
        strSubUnit = "D.住院单位 单位,D.住院包装 包装,C.实际数量 As 准退数量,A.数量 As 销帐数量"
    Case "药库单位"
        strSubUnit = "D.药库单位 单位,D.药库包装 包装,C.实际数量 As 准退数量,A.数量 As 销帐数量"
    End Select
    
    If mcondition.strNo <> "" Then
    ElseIf mcondition.str住院号 <> "" Then
        strCon = strCon & " And B.标识号=[4] "
    ElseIf mcondition.str姓名 <> "" Then
        strCon = strCon & " And B.姓名 Like [5] "
    ElseIf mcondition.lng病人ID <> -1 Then
        strCon = strCon & " And B.病人ID=[6] "
    ElseIf mcondition.str床号 <> "" Then
        strCon = strCon & " And B.床号 = [7] "
    End If
    
    If mParams.int药品名称显示 = 1 Then
        strSql药名 = "'['||X.编码||']'|| Nvl(K.名称,X.名称) As 药品名称,"
    Else
        strSql药名 = "'['||X.编码||']'|| X.名称 As 药品名称,"
    End If
    
    gstrSQL = "Select Distinct " & strSql药名 & "K.名称 As 商品名," & _
        " C.ID As 收发ID, C.药品ID, C.单据, C.NO, C.序号 As 收发序号, C.产地, C.批号, C.批次,C.效期, F.险类, P.名称 As 开单科室,E.名称 As 领药部门,E.Id As 领药部门Id, " & _
        " A.费用id, B.序号 As 费用序号, B.记录性质, B.主页ID, B.病人id, A.申请时间, " & strSubUnit & " " & _
        " From 病人费用销帐 A, 住院费用记录 B," & _
        " (Select A.ID, A.单据, A.NO, A.序号, A.药品id, A.产地, A.批号,A.批次, A.效期, A.费用id, B.实际数量 " & _
            " From 药品收发记录 A, " & _
            " (Select C.单据, C.NO, C.序号,C.批次, C.药品id, Sum(Nvl(C.付数, 1) * C.实际数量) As 实际数量 " & _
            " From 药品收发记录 C, 病人费用销帐 A, 住院费用记录 B " & _
            " Where A.申请类别=1 And A.费用id = B.ID And B.NO = C.NO And B.ID = C.费用id And A.状态 = 0 " & _
            " And C.单据 In (9, 10) And C.审核日期 Is Not Null And C.库房id = [1] And Instr([3], ',' || A.收费细目id || ',') > 0 " & strTmpCon

    '排除已在输液配置中心管理中产生的单据
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From 输液配药内容 Y Where Y.收发id = C.ID) "
    
    gstrSQL = gstrSQL & " Group By C.单据, C.NO, C.序号,C.批次, C.药品id " & _
            " Having Sum(Nvl(C.付数, 1) * C.实际数量) > 0) B" & _
            " Where A.NO = B.NO And A.单据 = B.单据 And A.药品id + 0 = B.药品id And A.序号 = B.序号 And A.审核人 Is Not Null " & _
            " And (A.记录状态 = 1 Or Mod(A.记录状态, 3) = 0))C, " & _
        " 药品规格 D, 收费项目目录 X, 收费项目别名 K, 部门表 P, 病案主页 F, 部门表 E " & _
        " Where A.申请类别=1 And A.费用id = B.ID And B.NO = C.NO And B.ID = C.费用id And B.开单部门id = P.ID And B.收费细目id = D.药品id And B.收费细目id = X.ID And B.病人id = F.病人id And B.主页id = F.主页id  And A.申请部门id = E.ID " & strCon & _
        " And X.Id = K.收费细目ID(+) AND K.性质(+)=3  And B.执行部门id = [1] And Instr([2], ',' || A.申请部门id || ',') > 0 And A.审核人 Is Null And A.状态 = 0 "
    
    If mParams.bln审核出院销账申请 = False Then
        gstrSQL = gstrSQL & " And F.出院日期 Is Null "
    End If
        
    gstrSQL = gstrSQL & " Order By A.申请时间, C.单据, C.NO, C.序号 Desc "
    
    On Error GoTo errHandle
    
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取批次明细", _
        mcondition.lng药房id, _
        "," & str科室ID串 & ",", _
        "," & str药品ID串 & ",", _
        mcondition.str住院号, _
        mcondition.str姓名, _
        mcondition.lng病人ID, _
        mcondition.str床号)
    
    If rstemp.EOF Then
        Exit Sub
    End If
    
    Do While Not rstemp.EOF
        With mrsChargeOff
            .AddNew
            !药品名称 = rstemp!药品名称
            !领药部门 = rstemp!领药部门
            !领药部门ID = rstemp!领药部门ID
            !单据 = rstemp!单据
            !NO = rstemp!NO
            !药品ID = rstemp!药品ID
            !申请时间 = Format(rstemp!申请时间, "yyyy-mm-dd hh:mm:ss")
            !病人ID = rstemp!病人ID
            !收发序号 = rstemp!收发序号
            !产地 = rstemp!产地
            !批号 = rstemp!批号
            !效期 = rstemp!效期
            !批次 = NVL(rstemp!批次, 0)
            
            If gtype_UserSysParms.P149_效期显示方式 = 1 And zlStr.NVL(!效期) <> "" Then
                '换算为有效期
                !效期 = Format(DateAdd("D", -1, !效期), "yyyy-mm-dd")
            End If
            
            !准退数量 = rstemp!准退数量
            !销帐数量 = rstemp!销帐数量
            !包装 = rstemp!包装
            !单位 = rstemp!单位
            !收发ID = rstemp!收发ID
            !主页id = IIf(IsNull(rstemp!主页id), 0, rstemp!主页id)
            !费用序号 = rstemp!费用序号
            !险类 = rstemp!险类
            !费用ID = rstemp!费用ID
            !记录性质 = rstemp!记录性质
            !审核标志 = 0
            !执行标志 = 0
            
            .Update
        End With
        
        With mrsChargeOffMain
'            dbl准退数量 = dbl准退数量 + rstemp!准退数量
            If lng领药部门ID <> rstemp!领药部门ID Or str申请时间 <> Format(rstemp!申请时间, "yyyy-mm-dd hh:mm:ss") Or lng费用id <> rstemp!费用ID Then
                .AddNew
                !领药部门ID = rstemp!领药部门ID
                !药品ID = rstemp!药品ID
                !申请时间 = Format(rstemp!申请时间, "yyyy-mm-dd hh:mm:ss")
                !费用ID = rstemp!费用ID
                !准退数量 = rstemp!准退数量
                !销帐数量 = rstemp!销帐数量
                
                .Update
                
                dbl准退数量 = 0
            Else
                !准退数量 = !准退数量 + rstemp!准退数量
                .Update
            End If
            lng领药部门ID = rstemp!领药部门ID
            str申请时间 = Format(rstemp!申请时间, "yyyy-mm-dd hh:mm:ss")
            lng费用id = rstemp!费用ID
        End With
        
        rstemp.MoveNext
    Loop
    
    '只处理发药清单对应的药品（按领药部门ID，药品ID为准）
    mrsChargeOff.MoveFirst
    Do While Not mrsChargeOff.EOF
        mrsSendData.Filter = "执行状态=" & mState.发药
        mrsSendData.Sort = "领药部门id,药品id"
        Do While Not mrsSendData.EOF
            If mrsChargeOff!领药部门ID = mrsSendData!领药部门ID And mrsChargeOff!药品ID = mrsSendData!药品ID Then
                mrsChargeOff!审核标志 = 1
                mrsChargeOff.Update
            End If
            mrsSendData.MoveNext
        Loop
        mrsChargeOff.MoveNext
    Loop
        
    Call AutoExpendQuantity
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AutoExpendQuantity()
    '考虑到同一费用ID对应多个收发ID的情况，需要将销帐数量分解到多个收发记录上
    '分解的原则是按序号大的优先分配（已按序号降序排序）
    Dim n As Integer
    Dim dbl准退数量 As Double
    Dim dbl剩余数量 As Double
    Dim int收发序号 As Integer
    Dim lng费用id As Long
    Dim lng药品id As Long
    Dim str申请时间 As String
    
    With mrsChargeOff
        If .RecordCount > 0 Then .MoveFirst
        For n = 1 To .RecordCount
            dbl准退数量 = !准退数量

            If lng费用id = !费用ID And lng药品id = !药品ID And str申请时间 = !申请时间 Then

            Else
                dbl剩余数量 = !销帐数量
            End If

            If dbl剩余数量 >= dbl准退数量 Then
                dbl剩余数量 = dbl剩余数量 - dbl准退数量
                !销帐数量 = dbl准退数量
            Else
                !销帐数量 = dbl剩余数量
                dbl剩余数量 = 0
            End If

            lng费用id = !费用ID
            lng药品id = !药品ID
            str申请时间 = !申请时间

            .Update
            .MoveNext
        Next
    End With
    
    '销帐数量大于了准退数量，则标志为拒绝审核
    With mrsChargeOffMain
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            mrsChargeOff.Filter = "药品ID=" & !药品ID & _
                " And 费用ID=" & !费用ID & _
                " And 申请时间='" & !申请时间 & "'"
            If mrsChargeOff.RecordCount > 0 Then
                If !准退数量 < !销帐数量 Then
                    Do While Not mrsChargeOff.EOF
                        mrsChargeOff!审核标志 = 2
                        mrsChargeOff.Update
                        mrsChargeOff.MoveNext
                    Loop
                End If
            End If
            .MoveNext
        Loop
    End With
End Sub

Private Sub InitReturnRec()
    '已发处方记录集
    Set mrsReturnData = New ADODB.Recordset
    With mrsReturnData
        If .State = 1 Then .Close
        
        .Fields.Append "收发ID", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adDouble, 18, adFldIsNullable
        .Fields.Append "科室", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "类型", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "扣率", adDouble, 2, adFldIsNullable                  '发药类型列，显示院内用药，离院带药，自取药等内容
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "单据", adDouble, 18, adFldIsNullable
        
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "主页ID", adDouble, 18, adFldIsNullable
        .Fields.Append "床号", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "性别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "住院号", adLongVarChar, 20, adFldIsNullable
        
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "品名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "其它名", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "英文名", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "毒理分类", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "效期", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "原产地", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "分批", adDouble, 2, adFldIsNullable
        .Fields.Append "付", adDouble, 18, adFldIsNullable
        .Fields.Append "数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "已退数", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "准退数", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "退药数", adDouble, 18, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "单价", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "金额", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "单量", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "频次", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "用法", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "说明", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "实际数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "库房货位", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "包装", adDouble, 18, adFldIsNullable
        .Fields.Append "高危药品", adDouble, 2, adFldIsNullable
        
        .Fields.Append "操作员", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "发药时间", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "费用序号", adDouble, 18, adFldIsNullable
       
        .Fields.Append "审查结果", adDouble, 18, adFldIsNullable
        .Fields.Append "医嘱id", adDouble, 18, adFldIsNullable
        .Fields.Append "领药人", adLongVarChar, 20, adFldIsNullable
        
        .Fields.Append "相关id", adDouble, 18, adFldIsNullable
        .Fields.Append "单量单位", adLongVarChar, 20, adFldIsNullable
        
        .Fields.Append "药品编码和名称", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "药品编码", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "药品名称", adLongVarChar, 50, adFldIsNullable
        
        .Fields.Append "发药号", adDouble, 18, adFldIsNullable
        
        .Fields.Append "医生嘱托", adLongVarChar, 40, adFldIsNullable
        
        .Fields.Append "状态", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "执行状态", adDouble, 1, adFldIsNullable
        
        .Fields.Append "转出", adDouble, 1, adFldIsNullable
        .Fields.Append "领药部门ID", adDouble, 18, adFldIsNullable
        .Fields.Append "发送时间", adLongVarChar, 40, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub InitApplyforcredit()
    '存在销帐申请的记录集
    Set mrsApplyforcredit = New ADODB.Recordset
    With mrsApplyforcredit
        If .State = 1 Then .Close
        
        .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
        .Fields.Append "收发ID", adDouble, 18, adFldIsNullable              '药品收发ID
        .Fields.Append "执行状态", adDouble, 1, adFldIsNullable             '状态的内部标识：0－缺药；1－发药；2－拒发；3－不处理
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "药品名称", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "销帐申请数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "性别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "年龄", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "领药部门", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "床号", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "病人科室", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub InitMsgRec()
    '消息接收记录集
    Set mrsReceiveMsg = New ADODB.Recordset
    With mrsReceiveMsg
        If .State = 1 Then .Close
        .Fields.Append "科室", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "住院号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "申请时间", adLongVarChar, 40, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub
Private Function LoadSendRecord(ByVal rsData As ADODB.Recordset) As Boolean
    '装载发药数据集
    Dim intState As Integer
    Dim strState As String
    
    On Error GoTo errHandle
    
    With rsData
        Do While Not .EOF
            mrsSendData.AddNew
            
            mrsSendData!收发ID = !收发ID
            mrsSendData!序号 = !序号
            mrsSendData!记录状态 = !记录状态
            mrsSendData!类型 = IIf(zlStr.NVL(!医嘱序号, 0) = 0, IIf(!门诊标志 = 1 Or !门诊标志 = 4, "门诊记帐单", IIf(!单据 = 9, "住院记帐单", "住院记帐表")), IIf(IsNull(!扣率) = True, "住院记帐单", IIf(!扣率 Like "0*", "长嘱", IIf(!扣率 Like "1*", "临嘱", "记帐表"))))
            mrsSendData!扣率 = IIf(IsNull(!扣率), 0, !扣率)
            mrsSendData!已收费 = !已收费
            mrsSendData!单据 = !单据
            mrsSendData!NO = !NO
            mrsSendData!记帐员 = IIf(IsNull(!操作员姓名), "", !操作员姓名)
            mrsSendData!说明 = IIf(IsNull(!说明), "", !说明)
            mrsSendData!记帐时间 = IIf(IsNull(!登记时间), "", Format(!登记时间, "yyyy-MM-dd HH:mm:ss"))
            mrsSendData!配药人 = IIf(IsNull(!配药人), "", !配药人)
            mrsSendData!审核人 = IIf(IsNull(!审核人), "", !审核人)
            mrsSendData!主页id = !主页id
            
            mrsSendData!费用序号 = !费用序号
            mrsSendData!审查结果 = !审查结果
            mrsSendData!医嘱id = !医嘱id
            mrsSendData!相关id = IIf(IsNull(!相关id), 0, !相关id)
            mrsSendData!抗生素 = !抗生素
            mrsSendData!用药目的 = zlStr.NVL(!用药目的)
            mrsSendData!用药理由 = !用药理由
            mrsSendData!用药次数 = IIf(Val(zlStr.NVL(!发送数次)) < 0, 0, Val(zlStr.NVL(!发送数次)))
            mrsSendData!皮试结果 = !皮试结果
            mrsSendData!开嘱时间 = !开嘱时间
            
            mrsSendData!科室ID = !科室ID
            mrsSendData!领药部门编码 = !领药部门编码
            mrsSendData!领药部门 = !领药部门
            mrsSendData!领药部门ID = !领药部门ID
            mrsSendData!领药号 = IIf(IsNull(!领药号), "", !领药号)
            
            mrsSendData!病人ID = !病人ID
            mrsSendData!姓名 = !姓名
            mrsSendData!性别 = IIf(IsNull(!性别), "", !性别)
            mrsSendData!住院号 = zlStr.NVL(!标识号)
            mrsSendData!科室 = !科室
            mrsSendData!床号 = !床号
            mrsSendData!开单医生 = !开单医生
            mrsSendData!开单部门id = !开单部门id
            mrsSendData!年龄 = !年龄
            mrsSendData!病人类型 = !病人类型
            mrsSendData!颜色 = IIf(IsNull(!颜色), vbBlack, !颜色)
            
            mrsSendData!药品ID = !药品ID
            mrsSendData!药名ID = !药名ID
            mrsSendData!品名 = !品名
            mrsSendData!其它名 = IIf(IsNull(!其它名), "", !其它名)
            mrsSendData!英文名 = IIf(IsNull(!英文名), "", !英文名)
            mrsSendData!配方名称 = IIf(IsNull(!配方名称), "", !配方名称)
            mrsSendData!规格 = !规格
            mrsSendData!产地 = IIf(IsNull(!产地), "", !产地)
            mrsSendData!原产地 = IIf(IsNull(!原产地), "", !原产地)
            mrsSendData!是否皮试 = !是否皮试
            
            mrsSendData!批次 = IIf(IsNull(!批次), 0, !批次)
            mrsSendData!批号 = IIf(IsNull(!批号), "", !批号)
            mrsSendData!效期 = IIf(IsNull(!效期), "", !效期)
            mrsSendData!分批 = IIf(IsNull(!分批), 0, !分批)
            mrsSendData!付 = IIf(IsNull(!付), 1, !付)
            mrsSendData!数量 = zlStr.FormatEx(IIf(IsNull(!数量), 1, !数量) / !包装, mintNumberDigit) & !单位
            mrsSendData!单价 = zlStr.FormatEx(!单价 * !包装, 5)
            
            mrsSendData!单位 = !单位
            mrsSendData!包装 = !包装
            
'            mrsSendData!金额 = Format(!金额, "#####0.00;-#####0.00; ;")
            mrsSendData!金额 = !金额
            mrsSendData!单量 = IIf(IsNull(!单量), "", zlStr.FormatEx(!单量, mintNumberDigit) & zlStr.NVL(!计算单位) & "(" & zlStr.FormatEx(!单量 / !剂量系数 / !包装, mintNumberDigit) & !单位 & ")")
            mrsSendData!原始单量 = IIf(IsNull(!单量), "", zlStr.FormatEx(!单量, mintNumberDigit) & zlStr.NVL(!计算单位))
            mrsSendData!频次 = IIf(IsNull(!频次), "", !频次)
            mrsSendData!实际数量 = zlStr.FormatEx(Val(IIf(IsNull(!数量), 1, !数量)) * (Val(IIf(IsNull(!付), 1, !付))) / !包装, 5)
            mrsSendData!单量单位 = zlStr.NVL(!计算单位)
            mrsSendData!用法 = IIf(IsNull(!用法), "", !用法)
            
            mrsSendData!发药数量 = IIf(IsNull(!数量), 1, !数量)
            
            mrsSendData!高危药品 = IIf(IsNull(!高危药品), 0, !高危药品)
            
            mrsSendData!毒理分类 = IIf(IsNull(!毒理分类), "", !毒理分类)
            mrsSendData!价值分类 = IIf(IsNull(!价值分类), "", !价值分类)
            
            mrsSendData!库房货位 = IIf(IsNull(!库房货位), "", !库房货位)
            mrsSendData!库存下限 = !库存下限
            
            mrsSendData!留存数量 = zlStr.FormatEx(IIf(IsNull(!留存数量), 0, !留存数量) / !包装, 5)
            
            mrsSendData!退药人 = IIf(IsNull(!退药人), "", !退药人)
            
            mrsSendData!医生嘱托 = IIf(IsNull(!医生嘱托), "", !医生嘱托)
            
            mrsSendData!药品编码和名称 = !品名
            mrsSendData!药品编码 = !药品编码
            mrsSendData!药品名称 = !药品名称
            
            mrsSendData!库存数量 = !库存数量
            
            mrsSendData!入出类别id = !入出类别id
            mrsSendData!入出系数 = !入出系数
            mrsSendData!填制人 = IIf(IsNull(!填制人), "", !填制人)
            mrsSendData!填制日期 = IIf(IsNull(!填制日期), "", Format(!填制日期, "yyyy-MM-dd HH:mm:ss"))
            mrsSendData!配药日期 = IIf(IsNull(!配药日期), "", Format(!配药日期, "yyyy-MM-dd HH:mm:ss"))
            mrsSendData!费用ID = !费用ID
            mrsSendData!禁忌药品说明 = zlStr.NVL(!禁忌药品说明)
            mrsSendData!药师审核标志 = NVL(!药师审核标志, 0)
            mrsSendData!执行分类 = NVL(!执行分类, 0)
            mrsSendData!类别 = NVL(!类别, 0)
            mrsSendData!零差价管理 = NVL(!零差价管理, 0)
            
            mrsSendData!位置 = .AbsolutePosition
            
            '检查是否允许发药
            intState = mState.发药
            If !已收费 = 0 Then intState = mState.不处理
            If Not IsNull(!说明) Then
                intState = IIf(!说明 = "拒发", mState.拒发_不处理, intState)
            End If
            If mParams.bln允许未审核处方发药 = False Then
                If IsNull(!审核人) Then
                    intState = mState.不处理
                Else
                    If Trim(!审核人) = "" Then intState = mState.不处理
                End If
            ElseIf intState = mState.不处理 Then
                intState = mState.发药
            End If
            
            '检查毒理分类，价值分类，高危分类
            If intState <> mState.不处理 Then
                If mParams.str毒理分类 <> "" And !毒理分类 <> "" Then
                    If InStr("," & mParams.str毒理分类 & ",", "," & !毒理分类 & ",") > 0 Then
                        intState = mState.不处理
                    End If
                End If
                If mParams.str价值分类 <> "" And !价值分类 <> "" Then
                    If InStr("," & mParams.str价值分类 & ",", "," & !价值分类 & ",") > 0 Then
                        intState = mState.不处理
                    End If
                End If
                If mParams.str高危分类 <> "" And !高危药品 <> "" Then
                    If InStr("," & mParams.str高危分类 & ",", "," & !高危药品 & ",") > 0 Then
                        intState = mState.不处理
                    End If
                End If
            End If
            
'            If !记录状态 > 1 Then
'                intState = mState.不处理
'            End If
            
            mrsSendData!执行状态 = intState
            
            Select Case intState
                Case mState.缺药
                    strState = "缺药"
                Case mState.发药
                    strState = "发药"
                Case mState.拒发
                    strState = "拒发"
                Case mState.不处理, mState.拒发_不处理
                    strState = "不处理"
            End Select
            mrsSendData!状态 = strState
            
            mrsSendData.Update
            
            .MoveNext
        Loop
        
        '缺药检查
        If mParams.bln缺药检查 = True Then
            Call CheckShortage(mrsSendData, False)
        End If
    End With
    
    LoadSendRecord = True
    Exit Function
errHandle:
    MsgBox "产生内部记录集时，发生不可预知的错误！", vbInformation, gstrSysName
    Call InitSendRec
    Exit Function
End Function

Private Function LoadReturnRecord(ByVal rsData As ADODB.Recordset) As Boolean
    Dim dblSumSended As Double '已发数量
    
    On Error GoTo errHandle
    
    With rsData
        Do While Not .EOF
            mrsReturnData.AddNew
            mrsReturnData!收发ID = !收发ID
            mrsReturnData!药品ID = !药品ID
            mrsReturnData!科室 = !科室
            mrsReturnData!领药部门ID = !领药部门ID
            mrsReturnData!类型 = IIf(zlStr.NVL(!医嘱序号, 0) = 0, IIf(!门诊标志 = 1 Or !门诊标志 = 4, "门诊记帐单", IIf(!单据 = 9, "住院记帐单", "住院记帐表")), IIf(IsNull(!扣率) = True, "住院记帐单", IIf(!扣率 Like "0*", "长嘱", IIf(!扣率 Like "1*", "临嘱", "记帐表"))))
            mrsReturnData!扣率 = IIf(IsNull(!扣率), 0, !扣率)
            mrsReturnData!NO = !NO
            mrsReturnData!单据 = !单据
            mrsReturnData!序号 = !序号
            mrsReturnData!费用序号 = !费用序号
            mrsReturnData!病人ID = !病人ID
            mrsReturnData!主页id = !主页id
            mrsReturnData!床号 = !床号
            mrsReturnData!姓名 = IIf(IsNull(!姓名), "", !姓名)
            mrsReturnData!性别 = IIf(IsNull(!性别), "", !性别)
            mrsReturnData!住院号 = zlStr.NVL(!标识号)
            mrsReturnData!品名 = !品名
            mrsReturnData!其它名 = !其它名
            mrsReturnData!英文名 = !英文名
            mrsReturnData!规格 = IIf(IsNull(!规格), "", !规格)
            mrsReturnData!产地 = IIf(IsNull(!产地), "", !产地)
            mrsReturnData!原产地 = IIf(IsNull(!产地), "", !产地)
            mrsReturnData!毒理分类 = zlStr.NVL(!毒理分类)
            mrsReturnData!分批 = IIf(IsNull(!分批), 0, !分批)
            mrsReturnData!批次 = IIf(IsNull(!批次), 0, !批次)
            mrsReturnData!批号 = IIf(IsNull(!批号), "", !批号)
            mrsReturnData!效期 = IIf(IsNull(!效期), "", !效期)
            mrsReturnData!付 = IIf(IsNull(!付), 1, !付)
            mrsReturnData!数量 = zlStr.FormatEx(IIf(IsNull(!数量), 1, !数量) / !包装, mintNumberDigit) & !单位
'            If !可操作 <> 1 Then
'                mrsReturnData!已退数 = zlStr.FormatEx(IIf(IsNull(!已退数量), 1, !已退数量) / !包装, 5)
'                mrsReturnData!准退数 = zlStr.FormatEx(IIf(IsNull(!准退数), 1, !准退数) / !包装, 5)
'                mrsReturnData!退药数 = zlStr.FormatEx(IIf(IsNull(!准退数), 1, !准退数) / !包装, 5)
'            Else
                dblSumSended = GetSumSended(!单据, !NO, !药品ID, !序号)
                mrsReturnData!已退数 = zlStr.FormatEx((Val(IIf(IsNull(!数量), 1, !数量)) * (Val(IIf(IsNull(!付), 1, !付))) - dblSumSended) / !包装, mintNumberDigit)
                mrsReturnData!准退数 = zlStr.FormatEx(dblSumSended / !包装, mintNumberDigit)
                mrsReturnData!退药数 = zlStr.FormatEx(dblSumSended / !包装, mintNumberDigit)
'            End If
            mrsReturnData!包装 = !包装
            mrsReturnData!单位 = !单位
            mrsReturnData!单价 = zlStr.FormatEx(!单价 * !包装, 5)
            mrsReturnData!金额 = !金额
            mrsReturnData!单量 = IIf(IsNull(!单量), "", zlStr.FormatEx(!单量, mintNumberDigit) & zlStr.NVL(!计算单位) & "(" & zlStr.FormatEx(!单量 / !剂量系数 / !包装, mintNumberDigit) & !单位 & ")")
            mrsReturnData!单量单位 = zlStr.NVL(!计算单位)
            mrsReturnData!频次 = IIf(IsNull(!频次), "", !频次)
            mrsReturnData!用法 = IIf(IsNull(!用法), "", !用法)
            mrsReturnData!说明 = IIf(IsNull(!说明), "", !说明)
            mrsReturnData!操作员 = IIf(IsNull(!审核人), "", !审核人)
            mrsReturnData!发药时间 = IIf(IsNull(!发药时间), "", !发药时间)
                        
            mrsReturnData!医生嘱托 = IIf(IsNull(!医生嘱托), "", !医生嘱托)
                        
            mrsReturnData!高危药品 = IIf(IsNull(!高危药品), 0, !高危药品)
            
            mrsReturnData!审查结果 = !审查结果
            mrsReturnData!医嘱id = !医嘱id
            mrsReturnData!领药人 = !领药人
            mrsReturnData!实际数量 = dblSumSended
            mrsReturnData!库房货位 = IIf(IsNull(!库房货位), "", !库房货位)
            mrsReturnData!转出 = Val(!转出)
            
            mrsReturnData!药品编码和名称 = !品名
            mrsReturnData!药品编码 = !药品编码
            mrsReturnData!药品名称 = !药品名称
            mrsReturnData!发送时间 = !发送时间
            
            mrsReturnData!发药号 = !发药号
            
            mrsReturnData!相关id = IIf(IsNull(!相关id), 0, !相关id)
            
            If Val(!转出) = -1 Then
                mrsReturnData!执行状态 = mState.转出记录
            ElseIf Val(!可操作) = 1 Then
                mrsReturnData!执行状态 = mState.退药_原始记录
            ElseIf Val(!可操作) = 2 Then
                mrsReturnData!执行状态 = mState.退药_发药记录
            ElseIf Val(!可操作) = 3 Then
                mrsReturnData!执行状态 = mState.退药_退药记录
            End If
            
            mrsReturnData!状态 = "不处理"
            
            mrsReturnData.Update
            
            .MoveNext
        Loop
    End With
    
    LoadReturnRecord = True
    Exit Function
errHandle:
    MsgBox "产生内部记录集时，发生不可预知的错误！", vbInformation, gstrSysName
    Call InitReturnRec
    Exit Function
End Function

Private Function GetSumSended(ByVal Int单据 As Integer, ByVal strNo As String, ByVal lng药品id As Long, ByVal int序号 As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strsql As String
    
    On Error GoTo errHandle
    strsql = "Select Sum(Nvl(付数, 1) * 实际数量) 已发数量 From 药品收发记录 Where 单据 = [1] And NO = [2] And 药品ID+0 = [3] And 序号 = [4] And 审核日期 Is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, "计算已发数量", Int单据, strNo, lng药品id, int序号)
    
    If Not rsTmp.EOF Then
        GetSumSended = rsTmp!已发数量
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub Load取自定义发药类型()
    '提取发药类型，并动态增加发药类型选择框
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    
    Set rsData = DeptSendWork_Get自定义发药类型
    
    With rsData
        mblnExistOtherSendType = Not .EOF
        picShowSendType.Visible = mblnExistOtherSendType
        picSendType.Visible = mblnExistOtherSendType
        
        If mblnExistOtherSendType = False Then
            Exit Sub
        Else
            chkSendType(0).Caption = rsData!名称
            chkSendType(0).Width = 150 + LenB(chkSendType(0).Caption) * 128
            If rsData.RecordCount > 1 Then
                rsData.MoveNext
                For i = 2 To rsData.RecordCount
                    Load chkSendType(i - 1)
                    chkSendType(i - 1).Visible = True
                    chkSendType(i - 1).Caption = rsData!名称
                    chkSendType(i - 1).Width = 150 + LenB(chkSendType(i - 1).Caption) * 128
                    rsData.MoveNext
                Next
            End If
        End If
    End With
End Sub

Private Sub Load时间范围()
    Dim dteTime As Date
    
    
    
    With cbo时间范围
        .Enabled = mPrives.bln过滤时间
        .Clear
        .AddItem "0-当天"
        .AddItem "1-两天内"
        .AddItem "2-三天内"
        .AddItem "3-指定时间范围"
        
        .ListIndex = IIf(mParams.intDays < 4, mParams.intDays, 3)
        
        If .ListIndex <> Val(.Tag) Then
            If (Val(.Tag) = 3 And .ListIndex < 3) Or (Val(.Tag) < 3 And .ListIndex = 3) Then
                Call picConMain_Resize
                Call picCondition_Resize
            End If
            .Tag = .ListIndex
        End If
    End With
    
    Dtp开始时间.Enabled = mPrives.bln过滤时间
    Dtp结束时间.Enabled = mPrives.bln过滤时间
    
    dteTime = Sys.Currentdate
    Dtp开始时间.Value = Format(DateAdd("d", -1 * mParams.intDays, dteTime), "yyyy-MM-dd 00:00:00")
    Dtp结束时间.Value = Format(dteTime, "yyyy-MM-dd") & " 23:59:59"
    mdateBegin = dteTime
End Sub
Private Sub RefreshData()
    Dim intType As Integer
    
    '刷新数据：刷新科室列表，并默认全部勾选，再刷新明细清单
    DoEvents
    cmdRefreshDept_Click
    
    DoEvents
    intType = IIf(tbcDetail.Selected.index = 0, 0, 1)
    chkAll(intType).Value = 1
    chkAll_Click intType
    
    DoEvents
    cmdRefresh_Click
End Sub

Private Sub SaveCustomSet()
    Dim str发药类型 As String
    Dim intSendType As Integer
    Dim n As Integer
    
    '发药类型
    '0-所有,1-不含离院带药,2-仅含离院带药,3-不含自取药,4-仅含自取药,5-院内用药(不包括离院带药和自取药),6-离院带药和自取药
    If chkSend(0).Value = 1 And chkSend(1).Value = 1 And chkSend(2).Value = 1 Then
        intSendType = 0
    ElseIf chkSend(0).Value = 1 And chkSend(2).Value = 1 Then
        intSendType = 1
    ElseIf chkSend(0).Value = 1 And chkSend(1).Value = 1 Then
        intSendType = 3
    ElseIf chkSend(1).Value = 1 And chkSend(2).Value = 1 Then
        intSendType = 6
    ElseIf chkSend(0).Value = 1 Then
        intSendType = 5
    ElseIf chkSend(1).Value = 1 Then
        intSendType = 2
    ElseIf chkSend(2).Value = 1 Then
        intSendType = 4
    End If

    '其他发药类型
    If mblnExistOtherSendType = True Then
        For n = 0 To chkSendType.UBound
            If chkSendType(n).Value = 1 Then
                str发药类型 = IIf(str发药类型 = "", "", str发药类型 & ",") & chkSendType(n).Caption
            End If
        Next
    End If
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "显示拒发", IIf(mParams.blnShowReject, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "按发送时间排序", IIf(mParams.blnSort, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "科室列表", IIf(mParams.blnOnlyShowDept, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "显示科室", mParams.intShowDept
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "病人排序", mParams.int病人排序
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "输入模式", mParams.int输入模式索引
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理", "医嘱类型", mParams.intAdviceType
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理", "发药类型", str发药类型
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理", "原发药类型", intSendType
End Sub
Private Sub SetColorState()
    '发药列表的颜色状态
    picColorStateSend(mSendListColor.SendState).BackColor = mListColor.State_Send
    picColorStateSend(mSendListColor.RejectState).BackColor = mListColor.State_Reject
    picColorStateSend(mSendListColor.UnProcessState).BackColor = mListColor.State_UnProcess
    picColorStateSend(mSendListColor.ShortageState).BackColor = mListColor.State_Shortage
    
    '退药列表的颜色状态
    picColorStateReturn(mReturnListColor.ReturnState).BackColor = mListColor.Return_Returned
    picColorStateReturn(mReturnListColor.OriginalState).BackColor = mListColor.Return_Original
    picColorStateReturn(mReturnListColor.SendedState).BackColor = mListColor.Return_Sended
End Sub

Private Sub SetCommandBar(ByVal intType As Integer)
    '1、根据系统参数、权限等改变菜单状态
    '2、根据当前页面和当前选择的明细记录，改变菜单状态
    
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl

    Select Case intType
        Case mListType.发药
            '发药
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = True
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = True
            End If
                
            '验证签名
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            If gblnESign部门发药 = True Then
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = True
                If Not cbrControl Is Nothing Then cbrControl.Visible = True
                
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            Else
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = False
                If Not cbrControl Is Nothing Then cbrControl.Visible = False
            End If
            
            '拒发恢复
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '退药
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '全选
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_SelAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            
            '全清
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_ClsAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            
            '自定义审核
            If mblnCustomCheck = True Then
                Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_CustomCheck, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Enabled = True
                End If
            End If
        Case mListType.汇总
            '发药
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = True
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = True
            End If
            
            '验证签名
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            If gblnESign部门发药 = True Then
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = True
                If Not cbrControl Is Nothing Then cbrControl.Visible = True
                
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            Else
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = False
                If Not cbrControl Is Nothing Then cbrControl.Visible = False
            End If
            
            '拒发
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '拒发恢复
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '退药
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '全选
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_SelAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            
            '全清
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_ClsAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            
            '自定义审核
            If mblnCustomCheck = True Then
                Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_CustomCheck, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Enabled = False
                End If
            End If
        Case mListType.拒发
            '发药
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
           '验证签名
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            If gblnESign部门发药 = True Then
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = True
                If Not cbrControl Is Nothing Then cbrControl.Visible = True
                
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            Else
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = False
                If Not cbrControl Is Nothing Then cbrControl.Visible = False
            End If
            
            '拒发恢复
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '退药
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '全选
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_SelAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            
            '全清
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_ClsAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            
            '自定义审核
            If mblnCustomCheck = True Then
                Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_CustomCheck, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Enabled = False
                End If
            End If
        Case mListType.缺药
            '发药
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '验证签名
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            If gblnESign部门发药 = True Then
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = True
                If Not cbrControl Is Nothing Then cbrControl.Visible = True
                
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            Else
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = False
                If Not cbrControl Is Nothing Then cbrControl.Visible = False
            End If
            
            '拒发
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '拒发恢复
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '退药
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '全选
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_SelAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            
            '全清
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_ClsAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            
            '自定义审核
            If mblnCustomCheck = True Then
                Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_CustomCheck, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Enabled = False
                End If
            End If
        Case mListType.退药
            '发药
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '验证签名
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            If gblnESign部门发药 = True Then
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = True
                If Not cbrControl Is Nothing Then cbrControl.Visible = True
                
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            Else
                If Not cbrMenu Is Nothing Then cbrMenu.Visible = False
                If Not cbrControl Is Nothing Then cbrControl.Visible = False
            End If
            
            '拒发
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '拒发恢复
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = False
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = False
            End If
            
            '退药
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = True
            End If
            Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Return, , True)
            If Not cbrControl Is Nothing Then
                cbrControl.Enabled = True
            End If
            
            '全选
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_SelAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = True
            End If
            
            '全清
            Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_ClsAll, , True)
            If Not cbrMenu Is Nothing Then
                cbrMenu.Enabled = True
            End If
            
            '自定义审核
            If mblnCustomCheck = True Then
                Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_CustomCheck, , True)
                If Not cbrControl Is Nothing Then
                    cbrControl.Enabled = False
                End If
            End If
    End Select
End Sub

Private Sub RefreshSendDept()
    '刷新待发药部门列表
    Dim rsData As ADODB.Recordset
    Dim strTmpSql As String
    Dim strDanger As String
    Dim strToxicology As String
   
    '''select
    gstrSQL = "Select " & IIf(mParams.strSourceDep = "", "", "/*+rule*/") & "  distinct H.ID, H.名称 As 科室名称, Nvl(A.领药号, 0) As 领药号, Decode(Nvl(c.婴儿费,0), 0, Nvl(b.姓名, c.姓名), z.婴儿姓名) 姓名, C.病人ID, Decode(Nvl(c.婴儿费,0), 0, Nvl(p.性别, c.性别), z.婴儿性别) 性别, Decode(Nvl(c.婴儿费,0), 0, p.年龄, Ceil(Sysdate - z.出生时间) || '天') 年龄, S.单据, S.NO, S.药品id, " & _
        " Decode(Nvl(C.医嘱序号, 0), 0, 0, 1) 医嘱序号, C.门诊标志, Nvl(S.扣率, 0) 扣率, S.ID As 收发id, S.填制日期 填制日期, 0 As 拒发, Nvl(B.当前床号,'') As 床号,W.颜色,c.婴儿费" & IIf(mParams.bln启用审方, ", nvl(q.审查结果,0) 审查结果,nvl(q.id,0) 审查id", "")
    
    '''from
    gstrSQL = gstrSQL & " From 住院费用记录 C, 药品收发记录 S, 病人信息 B, 药品规格 D, 药品特性 T, 未发药品记录 A,病案主页 P,部门表 H,病人类型 W, 病人新生儿记录 Z " & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([17]) As zlTools.t_NumList)) E ")
    gstrSQL = gstrSQL & IIf(mblnIs配置中心 And mParams.intCheck = 1, ",病人医嘱记录 Q", "")
    
    gstrSQL = gstrSQL & IIf(mParams.bln启用审方, ",处方审查记录 Q,处方审查明细 K ", "")
    
    '''where
    gstrSQL = gstrSQL & " Where A.对方部门id = H.ID" & IIf(mParams.strSourceDep = "", "", " And A.对方部门id=E.Column_Value ") & _
        " And C.病人id = B.病人id And C.病人id=P.病人ID And C.主页id=P.主页id And A.单据 = S.单据 And A.NO = S.NO And Nvl(A.库房id,[1]) = Nvl(S.库房id,[1]) And S.费用id = C.ID And c.病人id = z.病人id(+) And c.婴儿费 = z.序号(+) And C.主页id=Z.主页id(+) " & _
        IIf(mblnIs配置中心 And mParams.intCheck = 1, "And Q.id(+)=C.医嘱序号 And (q.id is null or (q.id is not null and q.药师审核标志=1)) ", "") & _
        " And Nvl(A.库房id,[1]) = Nvl(C.执行部门id,[1]) And S.药品id = D.药品id And D.药名id = T.药名id And P.病人类型=W.名称(+) " & _
        " And (H.撤档时间 Is Null Or H.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
        " And A.填制日期 Between [2] And [3] And S.审核日期 Is Null "
        
        gstrSQL = gstrSQL & IIf(mParams.bln启用审方, " and c.医嘱序号=k.医嘱id(+) and Q.id(+)=K.审方id and K.最后提交(+)=1", "")
    
    '站点控制
    If mstrDeptNode <> "" Then
        gstrSQL = gstrSQL & " And (H.站点 = [16] Or H.站点 Is Null) "
    End If
    
    '当前药房
    gstrSQL = gstrSQL & " And Nvl(A.库房id,[1]) + 0 = [1] "
    
    '录入信息
    If mcondition.str住院号 <> "" Then
        gstrSQL = gstrSQL & " And P.住院号 = [4] "
    ElseIf mcondition.str床号 <> "" Then
        '由于床号不唯一，转为通过病人ID来查询
        gstrSQL = gstrSQL & " And C.病人ID = [9] "
    ElseIf mcondition.str就诊卡 <> "" Then
        gstrSQL = gstrSQL & " And B.就诊卡号 = [6] "
    ElseIf mcondition.str姓名 <> "" Then
        gstrSQL = gstrSQL & " And P.姓名 = [7] "
    ElseIf mcondition.strNo <> "" Then
        gstrSQL = gstrSQL & " And A.NO = [8] "
    ElseIf mcondition.lng病人ID <> -1 Or (mParams.int输入模式 = mInputType.IC卡 And Me.txtInput.Text <> "") Then
        gstrSQL = gstrSQL & " And C.病人ID = [9] "
    ElseIf mcondition.str领药号 <> "" Then
        gstrSQL = gstrSQL & " And A.领药号 = [10] "
    ElseIf mcondition.lng领药部门ID <> -1 Then
        gstrSQL = gstrSQL & " And A.对方部门id + 0 = [11] "
    End If
    
    '操作模式:0-所有,1-记帐单,2-记帐表
    If mcondition.int操作模式 = 0 Then
        gstrSQL = gstrSQL & " And A.单据 IN(9,10)"
    ElseIf mcondition.int操作模式 = 1 Then
        gstrSQL = gstrSQL & " And A.单据=9"
    ElseIf mcondition.int操作模式 = 2 Then
        gstrSQL = gstrSQL & " And A.单据=10"
    End If
    
    '记账人
    If mcondition.str记账人 <> "所有记帐人" Then
        gstrSQL = gstrSQL & " And S.填制人 = [12] "
    End If
    
    '医嘱类型:0-所有,1-长嘱,2-临嘱,3-普通
    '用单量是否填写区分是否医嘱产生的药品单据
    If mcondition.int医嘱类型 = 0 Then
    ElseIf mcondition.int医嘱类型 = 1 Then
        gstrSQL = gstrSQL & " And S.扣率 Is Not Null And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '0_' And Nvl(C.医嘱序号,0) + 0 >0 "
    ElseIf mcondition.int医嘱类型 = 2 Then
        gstrSQL = gstrSQL & " And S.扣率 Is Not Null And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '1_' And Nvl(C.医嘱序号,0) + 0 >0 "
    ElseIf mcondition.int医嘱类型 = 3 Then
        gstrSQL = gstrSQL & " And (Nvl(C.医嘱序号,0) + 0 =0 Or S.扣率 Is Null) "
    ElseIf mcondition.int医嘱类型 = 4 Then
        gstrSQL = gstrSQL & " And S.扣率 Is Not Null And (Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '0_' Or Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '1_') And Nvl(C.医嘱序号,0) + 0 > 0 "
    End If
    
    '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药,3-不含自取药,4-仅含自取药,5-院内用药(不包括离院带药和自取药),6-离院带药和自取药
    If mcondition.int发药类型 = 0 Then
    ElseIf mcondition.int发药类型 = 1 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3'"
    ElseIf mcondition.int发药类型 = 2 Then
        gstrSQL = gstrSQL & " And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3'"
    ElseIf mcondition.int发药类型 = 3 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf mcondition.int发药类型 = 4 Then
        gstrSQL = gstrSQL & " And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf mcondition.int发药类型 = 5 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3' And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf mcondition.int发药类型 = 6 Then
        gstrSQL = gstrSQL & " And (Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3' Or Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4')"
    End If
    
    '处理范围
    Select Case mcondition.int处理范围
    Case 1
        gstrSQL = gstrSQL & " And S.实际数量>=0"
    Case 2
        gstrSQL = gstrSQL & " And S.实际数量<0"
    End Select
    
    '病人类型：病人或婴儿
    If mcondition.int病人类型 = 0 Then
        gstrSQL = gstrSQL & " And Nvl(C.婴儿费, 0) = 0 "
    ElseIf mcondition.int病人类型 = 1 Then
        gstrSQL = gstrSQL & " And Nvl(C.婴儿费, 0) > 0 "
    End If
    
'    '是否显示待发单据
'    If mcondition.bln显示退药待发单据 = False Then
'        gstrSQL = gstrSQL & " And S.记录状态 = 1"
'    Else
'        gstrSQL = gstrSQL & " And Mod(S.记录状态, 3) = 1"
'    End If
    
    gstrSQL = gstrSQL & " And Mod(S.记录状态, 3) = 1"
    
    '给药途径
    If mcondition.str给药途径 <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [13] || ',',',' || S.用法 || ',') > 0 "
    End If
    
    '药品剂型
    If mcondition.str药品剂型 <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [14] || ',',',' || T.药品剂型 || ',') > 0 "
    End If
    
    '其它发药类型
    If mcondition.str其它发药类型 <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [15] || ',',',' || D.发药类型 || ',') > 0 "
    End If
    
    '科室类型
    If Trim(txtInput.Text) = "" Then
        If mParams.intShowDept = 1 Then
            gstrSQL = gstrSQL & " And H.id In (Select 部门id From 部门性质说明 Where 工作性质 = '临床' And 服务对象 In (2, 3)) "
        ElseIf mParams.intShowDept = 2 Then
            gstrSQL = gstrSQL & " And H.id In (Select 部门ID From 部门性质说明 Where 工作性质 In ('检查','检验','治疗','手术','营养') And 服务对象 IN(2,3)) "
        ElseIf mParams.intShowDept = 3 Then
            gstrSQL = gstrSQL & " And H.id In (Select 部门ID From 部门性质说明 Where 工作性质='护理' And 服务对象 IN(2,3)) "
        End If
    End If

    '排除已在输液配置中心管理中产生的单据
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From 输液配药内容 Y Where Y.收发id = S.ID) "
    
    '排除对未发药品的销帐记录
    If chkWithNotAudited.Value = 0 Then
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From 病人费用销帐 X " & _
            " Where X.申请类别 = 0 And X.状态+0 = 0 And X.收费细目id+0 = S.药品id And X.费用id = S.费用id)"
    End If
    
    '高危药品
    If chkDanger.Value = 1 Then
        If chkDangerType(0).Value = 1 Then strDanger = IIf(strDanger = "", 1, strDanger & "," & 1)
        If chkDangerType(1).Value = 1 Then strDanger = IIf(strDanger = "", 2, strDanger & "," & 2)
        If chkDangerType(2).Value = 1 Then strDanger = IIf(strDanger = "", 3, strDanger & "," & 3)
    End If
    
    '毒理分类
    If Me.chkToxicologyType.Value = 1 Then
        If Me.chkToxicology(0).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(0).Caption, strToxicology & "," & Me.chkToxicology(0).Caption)
        If Me.chkToxicology(1).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(1).Caption, strToxicology & "," & Me.chkToxicology(1).Caption)
        If Me.chkToxicology(2).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(2).Caption, strToxicology & "," & Me.chkToxicology(2).Caption)
        If Me.chkToxicology(3).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(3).Caption, strToxicology & "," & Me.chkToxicology(3).Caption)
    End If
    
    If strDanger <> "" Then gstrSQL = gstrSQL & " And Instr(',' || [18] || ',' , ',' || Nvl(D.高危药品,0) || ',') > 0 "
    
    If strToxicology <> "" Then gstrSQL = gstrSQL & " And Instr(',' || [19] || ',' , ',' || T.毒理分类 || ',') > 0 "
    
    If mParams.blnShowReject = True Then
        '合并拒发记录
        strTmpSql = " (Select A.单据, A.NO, A.病人id, A.主页id, A.姓名, Nvl(B.优先级, 0) 优先级, A.对方部门id, A.库房id, A.发药窗口, A.填制日期, A.已收费, Null As 配药人," & _
                " 0 As 打印状态, 0 As 未发数, A.产品合格证 As 领药号 " & _
                " From (Select B.单据, B.NO, A.病人id, A.主页id, A.姓名, Decode(A.记录状态, 0, 0, 1) 已收费, B.对方部门id, B.库房id, " & _
                " B.发药窗口 , B.填制日期, C.身份, B.产品合格证 " & _
                " From 住院费用记录 A, 药品收发记录 B, 病人信息 C " & _
                 IIf(mblnIs配置中心 And mParams.intCheck = 1, ",病人医嘱记录 Q", "") & _
                " Where A.ID = B.费用id + 0 And B.单据 In (9, 10) And B.审核日期 Is Null And B.摘要 = '拒发' And " & _
                IIf(mblnIs配置中心 And mParams.intCheck = 1, " Q.id(+)=A.医嘱序号 And (q.id is null or (q.id is not null and q.药师审核标志=1)) And ", "") & _
                " Nvl(B.库房id,[1]) + 0 = [1] And B.填制日期 Between [2] And [3] And A.病人id = C.病人id(+)) A, 身份 B " & _
                " Where B.名称(+) = A.身份) "
                
        strTmpSql = Replace(gstrSQL, "未发药品记录", strTmpSql)
        strTmpSql = Replace(strTmpSql, "0 As 拒发", "1 As 拒发")
        
        gstrSQL = gstrSQL & " Union All " & strTmpSql
    End If
    
    '''order by
    gstrSQL = gstrSQL & " Order By 科室名称,填制日期 desc, ID, 领药号, 姓名, NO "
    
    On Error GoTo errHandle
    
    Me.MousePointer = 11
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "提取待发药科室汇总", _
        mcondition.lng药房id, _
        CDate(mcondition.str开始时间), _
        CDate(mcondition.str结束时间), _
        mcondition.str住院号, _
        mcondition.str床号, _
        mcondition.str就诊卡, _
        mcondition.str姓名, _
        mcondition.strNo, _
        mcondition.lng病人ID, _
        mcondition.str领药号, _
        mcondition.lng领药部门ID, _
        mcondition.str记账人, _
        mcondition.str给药途径, _
        mcondition.str药品剂型, _
        mcondition.str其它发药类型, _
        mstrDeptNode, _
        mParams.strSourceDep, _
        strDanger, _
        strToxicology)
      
    If mParams.bln启用审方 Then
        rsData.Filter = "(审查结果=1 and 审查id<>0) or 审查id=0"
    End If
    '更新部门树表
    Call GetSendDeptTreeView(rsData)
    
    '更新部门树表对应的收发记录数据集
    Call GetDeptListRecord(rsData)
    
    Me.MousePointer = 0
    Exit Sub
errHandle:
    Me.MousePointer = 0
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub RefreshSendDetail()
    '刷新待发药明细列表
    Dim rsData As ADODB.Recordset
    Dim strSql退药人 As String
    Dim str收发ID串 As String
    Dim lng当前科室 As Long
    Dim str当前NO As String
    Dim strSqlTmp As String
    Dim strSqlUnion As String
    Dim i As Integer
    Dim strArr收发id As Variant
    Dim ArrTmp As Variant
    Dim intCount As Integer
    Dim strTmp As String
    
    If Val(tvwList(mDeptType.发药).Tag) = 0 Then Exit Sub
    On Error GoTo errHandle
    '根据部门列表实际勾选情况，按科室ID、NO、收发ID等组织条件
    If mrsDeptList Is Nothing Then Exit Sub
    mrsDeptList.Filter = ""
    mrsDeptList.Sort = "科室ID,NO,收发ID"
    mstr科室ID串 = ""
    mstr科室名串 = ""
    With mrsDeptList
        Do While Not .EOF
            If !执行状态 = 1 Then
                If lng当前科室 <> !科室ID Then
                    mstr科室ID串 = mstr科室ID串 & IIf(mstr科室ID串 = "", "", ",") & !科室ID
                    mstr科室名串 = mstr科室名串 & IIf(mstr科室名串 = "", "", ",") & !科室名称
                    lng当前科室 = !科室ID
                End If
                
                If InStr(1, "," & str收发ID串 & ",", "," & !收发ID & ",") = 0 Then
                    str收发ID串 = str收发ID串 & IIf(str收发ID串 = "", "", ",") & !收发ID
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    If str收发ID串 = "" Then Exit Sub
    
    '分解收发ID串
    '收发ID串大于4K时分成小于4K的串（绑定变量时，最大变量长度为4K字符）
    strArr收发id = Array()
    ArrTmp = Split(str收发ID串 & ",", ",")
    intCount = UBound(ArrTmp)
    
    '查询提示
    If WarRecoredCount(intCount) = False Then Exit Sub

    If Len(str收发ID串) >= 4000 Then
        For i = 0 To intCount
            If ArrTmp(i) <> "" Then
                If Len(IIf(strTmp = "", "", strTmp & ",") & ArrTmp(i)) >= 4000 Then
                    ReDim Preserve strArr收发id(UBound(strArr收发id) + 1)
                    strArr收发id(UBound(strArr收发id)) = strTmp
                    strTmp = ArrTmp(i)
                Else
                    strTmp = IIf(strTmp = "", "", strTmp & ",") & ArrTmp(i)
                End If
            End If
                   
            If i = intCount Then
                ReDim Preserve strArr收发id(UBound(strArr收发id) + 1)
                strArr收发id(UBound(strArr收发id)) = strTmp
            End If
        Next
    Else
        ReDim Preserve strArr收发id(UBound(strArr收发id) + 1)
        strArr收发id(UBound(strArr收发id)) = str收发ID串
    End If
    
    '''select
    gstrSQL = "SELECT /*+rule*/ Distinct A.*, Nvl(C.留存数量,0) As 留存数量 " & IIf(mcondition.bln显示领药退药人 = True, ", B.退药人", ",'' As 退药人") & " FROM ("
    
    strSqlTmp = "SELECT DISTINCT S.ID As 收发ID,to_char(s.效期,'yyyy-mm-dd') 效期,S.记录状态,S.药品ID,S.费用id,NVL(N.已收费,0) 已收费,P.名称 科室,S.配药人,C.开单人 开单医生,C.开单部门id,C.操作员姓名 审核人,S.单据,S.扣率, " & _
             " S.NO,S.序号,C.病人ID,Nvl(C.主页ID,0) As 主页ID,Nvl(C.床号,'(未安排)') As 床号,Decode(Nvl(c.婴儿费,0), 0, Nvl(Q.姓名, C.姓名), U.婴儿姓名) 姓名,Decode(Nvl(c.婴儿费,0), 0, Nvl(Q.性别, C.性别), U.婴儿性别) 性别,C.门诊标志,C.标识号,C.操作员姓名,S.付数 付,S.实际数量 数量," & _
             " NVL(D.药房分批,0) 分批,Nvl(D.高危药品,0) As 高危药品,X.规格,T.毒理分类,T.价值分类,Nvl(T.抗生素,0) 抗生素,C.登记时间,H.编码 As 领药部门编码,H.名称 As 领药部门,H.Id As 领药部门Id," & _
             " S.零售价 单价,S.零售金额 金额,S.单量,S.频次,S.用法,S.摘要 说明,DECODE(S.批号,NULL,'',S.批号)||DECODE(S.批次,NULL,'',0,'','('||S.批次||')') 批号,NVL(S.批次,0) 批次, Ceil((s.实际数量 * d.剂量系数) / Nvl(s.单量, 1)) As 发送数次," & _
             " C.医嘱序号,I.计算单位,NVL(S.产地,NVL(X.产地,'')) 产地,S.原产地,nvl(M.审查结果,-1) 审查结果,M.皮试结果,Nvl(M.开嘱时间,C.登记时间) As 开嘱时间,decode(m.用药目的,1,'预防',2,'治疗',3,'预防和治疗','') 用药目的,m.用药理由,D.药名ID,nvl(C.医嘱序号,-1) 医嘱id," & IIf(mParams.bln药品储备 = True, "L.", "'' ") & "库房货位," & _
             " M.相关ID,M.药师审核标志,M.禁忌药品说明,C.病人科室ID As 科室ID,C.序号 费用序号," & IIf(mParams.bln药品储备 = True, "Decode(Sign(Nvl(K.库存数量, 0) - Nvl(L.下限, 0)), -1, 0, 1) ", "0 ") & " 库存下限, Z.名称 As 英文名, Decode(Nvl(c.婴儿费,0), 0, Q.年龄, Ceil(Sysdate - U.出生时间) || '天') 年龄,Q.病人类型,W.颜色,N.领药号, " & _
             IIf(mParams.int药品名称显示 = 0 Or mParams.int药品名称显示 = 2, "NVL(E.名称,'')", "Decode(E.名称,Null,'',X.名称)") & " As 其它名, " & _
             "'['||X.编码||']'||" & IIf(mParams.int药品名称显示 = 1, "NVL(E.名称,X.名称)", "X.名称") & " As 品名,nvl(K.名称,'') 配方名称," & _
             "X.编码" & " As 药品编码," & IIf(mParams.int药品名称显示 = 1, "NVL(E.名称,X.名称)", "X.名称") & " As 药品名称,s.入出类别id,s.入出系数,s.填制人,s.填制日期,s.配药日期,Nvl(t.是否皮试,0) As 是否皮试,F.执行分类,F.类别,D.剂量系数, m.医生嘱托, nvl(d.是否零差价管理,0) as 零差价管理 "
    
'    '测试分组（相关ID设为1）
'    strSqlTmp = "SELECT DISTINCT S.ID As 收发ID,S.记录状态,S.药品ID,NVL(N.已收费,0) 已收费,P.名称 科室,S.配药人,C.开单人 开单医生,C.操作员姓名 审核人,S.单据,S.扣率," & _
'             " S.NO,S.序号,C.病人ID,C.床号,C.姓名,C.门诊标志,C.标识号,C.操作员姓名,S.付数 付,S.实际数量 数量," & _
'             " NVL(D.药房分批,0) 分批,X.规格,T.毒理分类,T.价值分类,C.登记时间,H.名称 As 领药部门,H.Id As 领药部门Id," & _
'             " S.零售价 单价,S.零售金额 金额,S.单量,S.频次,S.用法,S.摘要 说明,DECODE(S.批号,NULL,'',S.批号)||DECODE(S.批次,NULL,'',0,'','('||S.批次||')') 批号,NVL(S.批次,0) 批次," & _
'             " C.医嘱序号,I.计算单位,NVL(S.产地,NVL(X.产地,'')) 产地,nvl(M.审查结果,-1) 审查结果,nvl(C.医嘱序号,-1) 医嘱id," & IIf(mParams.bln药品储备 = True, "L.", "'' ") & "库房货位," & _
'             " 1 相关ID,C.病人科室ID As 科室ID,C.序号 费用序号," & IIf(mParams.bln药品储备 = True, "Decode(Sign(Nvl(K.库存数量, 0) - Nvl(L.下限, 0)), -1, 0, 1) ", "0 ") & " 库存下限, Z.名称 As 英文名, R.年龄, N.领药号, " & _
'             IIf(mParams.int药品名称显示 = 0 Or mParams.int药品名称显示 = 2, "NVL(E.名称,'')", "Decode(E.名称,Null,'',X.名称)") & " As 其它名, " & _
'             "'['||X.编码||']'||" & IIf(mParams.int药品名称显示 = 1, "NVL(E.名称,X.名称)", "X.名称") & " As 品名," & _
'             "X.编码" & " As 药品编码," & IIf(mParams.int药品名称显示 = 1, "NVL(E.名称,X.名称)", "X.名称") & " As 药品名称,s.入出类别id,s.入出系数,s.填制人,s.填制日期,s.配药日期"
           
    '单位设置
    Select Case mParams.strUnit
    Case "售价单位"
        strSqlTmp = strSqlTmp & ",X.计算单位 单位,1 包装 "
    Case "门诊单位"
        strSqlTmp = strSqlTmp & ",D.门诊单位 单位,D.门诊包装 包装 "
    Case "住院单位"
        strSqlTmp = strSqlTmp & ",D.住院单位 单位,D.住院包装 包装 "
    Case "药库单位"
        strSqlTmp = strSqlTmp & ",D.药库单位 单位,D.药库包装 包装 "
    End Select
    
    '缺药检查
    If mParams.bln缺药检查 = True Then
        strSqlTmp = strSqlTmp & " ,A.实际数量 As 库存数量 "
    Else
        strSqlTmp = strSqlTmp & " ,0 As 库存数量 "
    End If
    
    '''from
    strSqlTmp = strSqlTmp & " FROM 药品收发记录 S,住院费用记录 C,病人医嘱记录 M,病人医嘱记录 G,未发药品记录 N,收费项目别名 E,收费项目目录 X,诊疗项目目录 I,诊疗项目目录 K,诊疗项目目录 F," & _
             " 药品规格 D,药品特性 T," & IIf(mParams.bln药品储备 = True, "药品储备限额 L,", "") & "诊疗项目别名 Z,部门表 P,部门表 H,病人信息 R,病案主页 Q,病人类型 W,病人新生儿记录 U "
             
            
    '用收发ID是最忧的，尽量用收发ID作为条件
    strSqlTmp = strSqlTmp & " ,Table(Cast(f_Num2List([15]) As zlTools.t_NumList)) G "
    
    If mParams.bln药品储备 = True Then
        strSqlTmp = strSqlTmp & ",(Select 库房id, 药品id, Nvl(Sum(实际数量), 0) 库存数量 From 药品库存 Where 性质 = 1 And 库房id = [1] Group By 库房id, 药品id) K "
    End If
    
    If mParams.bln缺药检查 = True Then
        strSqlTmp = strSqlTmp & ",(Select 库房id, 药品id, 实际数量, Nvl(批次, 0) 批次 From 药品库存 Where 性质 = 1 And 库房id = [1]) A "
    End If
             
    strSqlTmp = strSqlTmp & " WHERE S.NO=N.NO AND S.单据=N.单据 AND NVL(S.库房ID,[1])+0=NVL(N.库房ID,[1]) AND S.费用ID=C.ID And S.药品ID=D.药品ID And c.病人id = u.病人id(+) And c.婴儿费 = u.序号(+) And C.主页id=U.主页id(+) " & _
            " And C.病人id = R.病人id And C.病人id=Q.病人id And C.主页id=Q.主页id And Q.病人类型=W.名称(+) " & _
            " AND S.对方部门ID+0=H.ID AND S.审核人 IS NULL AND NVL(S.库房ID,[1])+0=[1] " & _
            " AND C.病人科室ID=P.id And d.药品ID=X.ID and D.药名ID=T.药名ID AND D.药名ID=I.ID and C.医嘱序号=M.ID(+) and M.相关id=G.id(+) and G.配方id=K.id(+) and G.诊疗项目id=F.id(+) " & _
            " And D.药名id = Z.诊疗项目id(+) And Z.性质(+) = 2 " & IIf(mParams.bln药品储备 = True, " And S.药品ID=L.药品ID(+) And Nvl(S.库房ID,[1])=L.库房ID(+) ", "") & _
            " AND D.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
            " And nvl(S.发药方式,-999)<>-1 " & _
            " And S.单据 In(9,10)  And N.填制日期 Between [2] And [3] "
    
    strSqlTmp = strSqlTmp & " And S.ID= G.Column_Value "
    
    If mParams.bln药品储备 = True Then
        strSqlTmp = strSqlTmp & " And Nvl(S.库房id, [1]) + 0 = K.库房id(+) And S.药品id = K.药品id(+) "
    End If
    
    If mParams.bln缺药检查 = True Then
        strSqlTmp = strSqlTmp & " And Nvl(S.库房id, [1]) + 0 = A.库房id(+) And S.药品id = A.药品id(+) And Nvl(S.批次, 0) = A.批次(+) "
    End If
    
    '录入信息
    If mcondition.str住院号 <> "" Then
        strSqlTmp = strSqlTmp & " And Q.住院号 = [8] "
    ElseIf mcondition.str床号 <> "" Then
        strSqlTmp = strSqlTmp & " And R.当前床号 = [9] "
    ElseIf mcondition.str就诊卡 <> "" Then
        strSqlTmp = strSqlTmp & " And R.就诊卡号 = [10] "
    ElseIf mcondition.str姓名 <> "" Then
        strSqlTmp = strSqlTmp & " And N.姓名 = [11] "
    ElseIf mcondition.strNo <> "" Then
        strSqlTmp = strSqlTmp & " And N.NO = [12] "
    ElseIf mcondition.lng病人ID <> -1 Then
        strSqlTmp = strSqlTmp & " And N.病人ID = [13] "
    ElseIf mcondition.str领药号 <> "" Then
        strSqlTmp = strSqlTmp & " And N.领药号 = [14] "
    End If
    
    '操作模式:0-所有,1-记帐单,2-记帐表
    If mcondition.int操作模式 = 1 Then
        strSqlTmp = strSqlTmp & " And S.单据=9"
    ElseIf mcondition.int操作模式 = 2 Then
        strSqlTmp = strSqlTmp & " And S.单据=10"
    End If
    
    '记账人
    If mcondition.str记账人 <> "所有记帐人" Then
        strSqlTmp = strSqlTmp & " And S.填制人 = [7] "
    End If
    
    '医嘱类型:0-所有,1-长嘱,2-临嘱,3-普通
    '用单量是否填写区分是否医嘱产生的药品单据
    If mcondition.int医嘱类型 = 0 Then
    ElseIf mcondition.int医嘱类型 = 1 Then
        strSqlTmp = strSqlTmp & " And S.扣率 Is Not Null And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '0_' And Nvl(C.医嘱序号,0) + 0 >0 "
    ElseIf mcondition.int医嘱类型 = 2 Then
        strSqlTmp = strSqlTmp & " And S.扣率 Is Not Null And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '1_' And Nvl(C.医嘱序号,0) + 0 >0 "
    ElseIf mcondition.int医嘱类型 = 3 Then
        strSqlTmp = strSqlTmp & " And (Nvl(C.医嘱序号,0) + 0 =0 Or S.扣率 Is Null) "
    ElseIf mcondition.int医嘱类型 = 4 Then
        strSqlTmp = strSqlTmp & " And S.扣率 Is Not Null And (Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '0_' Or Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '1_') And Nvl(C.医嘱序号,0) + 0 > 0 "
    End If
    
    '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药,3-不含自取药,4-仅含自取药,5-院内用药(不包括离院带药和自取药),6-离院带药和自取药
    If mcondition.int发药类型 = 0 Then
    ElseIf mcondition.int发药类型 = 1 Then
        strSqlTmp = strSqlTmp & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3'"
    ElseIf mcondition.int发药类型 = 2 Then
        strSqlTmp = strSqlTmp & " And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3'"
    ElseIf mcondition.int发药类型 = 3 Then
        strSqlTmp = strSqlTmp & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf mcondition.int发药类型 = 4 Then
        strSqlTmp = strSqlTmp & " And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf mcondition.int发药类型 = 5 Then
        strSqlTmp = strSqlTmp & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3' And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf mcondition.int发药类型 = 6 Then
        strSqlTmp = strSqlTmp & " And (Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3' Or Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4')"
    End If
    
    '处理范围
    Select Case mcondition.int处理范围
    Case 1
        strSqlTmp = strSqlTmp & " And S.实际数量>=0"
    Case 2
        strSqlTmp = strSqlTmp & " And S.实际数量<0"
    End Select
    
    '病人类型：病人或婴儿
    If mcondition.int病人类型 = 0 Then
        strSqlTmp = strSqlTmp & " And Nvl(C.婴儿费, 0) = 0 "
    ElseIf mcondition.int病人类型 = 1 Then
        strSqlTmp = strSqlTmp & " And Nvl(C.婴儿费, 0) > 0 "
    End If
    
    '给药途径
    If mcondition.str给药途径 <> "" Then
        strSqlTmp = strSqlTmp & " And Instr(',' || [4] || ',',',' || S.用法 || ',') > 0 "
    End If
    
    '药品剂型
    If mcondition.str药品剂型 <> "" Then
        strSqlTmp = strSqlTmp & " And Instr(',' || [5] || ',',',' || T.药品剂型 || ',') > 0 "
    End If
    
    '其它发药类型
    If mcondition.str其它发药类型 <> "" Then
        strSqlTmp = strSqlTmp & " And Instr(',' || [6] || ',',',' || D.发药类型 || ',') > 0 "
    End If
    
    '科室类型
    If Trim(txtInput.Text) = "" Then
        If mParams.intShowDept = 1 Then
            strSqlTmp = strSqlTmp & " And H.id In (Select 部门id From 部门性质说明 Where 工作性质 = '临床' And 服务对象 In (2, 3)) "
        ElseIf mParams.intShowDept = 2 Then
            strSqlTmp = strSqlTmp & " And H.id In (Select 部门ID From 部门性质说明 Where 工作性质 In ('检查','检验','治疗','手术','营养') And 服务对象 IN(2,3)) "
        ElseIf mParams.intShowDept = 3 Then
            strSqlTmp = strSqlTmp & " And H.id In (Select 部门ID From 部门性质说明 Where 工作性质='护理' And 服务对象 IN(2,3)) "
        End If
    End If
    
    '排除已在输液配置中心管理中产生的单据
    strSqlTmp = strSqlTmp & " And Not Exists (Select 1 From 输液配药内容 Y Where Y.收发id = S.ID) "
    
    '合并拒发记录
    strSqlUnion = " (Select A.单据, A.NO, A.病人id, A.主页id, A.姓名, Nvl(B.优先级, 0) 优先级, A.对方部门id, A.库房id, A.发药窗口, A.填制日期, A.已收费, Null As 配药人," & _
            " 0 As 打印状态, 0 As 未发数, A.产品合格证 As 领药号 " & _
            " From (Select B.单据, B.NO, A.病人id, Nvl(A.主页ID,0) As 主页ID,A.姓名, Decode(A.记录状态, 0, 0, 1) 已收费, B.对方部门id, B.库房id, " & _
            " B.发药窗口 , B.填制日期, C.身份, B.产品合格证 " & _
            " From 住院费用记录 A, 药品收发记录 B, 病人信息 C " & _
            " Where A.ID = B.费用id + 0 And B.审核日期 Is Null And B.摘要 = '拒发' And " & _
            " Nvl(B.库房id,[1]) = [1] And B.填制日期 Between [2] And [3] And A.病人id = C.病人id(+)) A, 身份 B " & _
            " Where B.名称(+) = A.身份) "
            
    strSqlTmp = strSqlTmp & " Union All " & Replace(strSqlTmp, "未发药品记录", strSqlUnion)
    
    gstrSQL = gstrSQL & strSqlTmp & ") A "
    
    gstrSQL = gstrSQL & ", (Select 药品id,库房id,部门id,留存数量 From 药品留存计划  Where 状态=0) C "
    
    '求最后一次退药的退药人
    If mcondition.bln显示领药退药人 = True Then
        strSql退药人 = ",(Select a.单据 ,a.No,a.序号,a.领用人 退药人 From 药品收发记录 a," & _
                " (Select s.单据,s.No,s.序号, Max(s.记录状态) 记录状态 " & _
                " From 药品收发记录 s, 未发药品记录 n " & _
                " Where s.No = n.No And s.单据 = n.单据 And Nvl(s.库房id, [1]) + 0 = Nvl(n.库房id, [1]) And " & _
                " Nvl(s.库房id, [1]) + 0 = [1] " & _
                " And Nvl(s.发药方式, -999) <> -1 And " & _
                " Mod(s.记录状态, 3) = 2 And s.单据 In (9, 10) " & _
                " Group By s.单据,s.No,s.序号) b " & _
                " Where a.单据=b.单据 And a.No=b.No And a.序号=b.序号 And a.记录状态=b.记录状态) B "
        gstrSQL = gstrSQL & strSql退药人
    End If
    
    gstrSQL = gstrSQL & " Where A.领药部门id = C.部门id(+) And C.库房id(+) = [1] And A.药品id = C.药品id(+) "
    
    If mcondition.bln显示领药退药人 = True Then
        gstrSQL = gstrSQL & " And A.单据 = B.单据(+) And A.No = B.No(+) And A.序号 = B.序号(+) "
    End If
    
    '排除对未发药品的销帐记录
    If chkWithNotAudited.Value = 0 Then
        gstrSQL = gstrSQL & " And Not Exists (Select 1 From 病人费用销帐 X " & _
            " Where X.申请类别 = 0 And X.状态+0 = 0 And X.收费细目id+0 = A.药品id And X.费用id = A.费用id) "
    End If
    
    gstrSQL = gstrSQL & "  Order By a.科室,a.No,a.费用序号 "
    
    On Error GoTo errHandle
    
    Me.MousePointer = 11
    Call AviShow(Me)
    
    Call InitSendRec
    
    '根据收发ID串的数组数目，循环执行
    For i = 0 To UBound(strArr收发id)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "提取单据信息", _
            mcondition.lng药房id, CDate(mcondition.str开始时间), CDate(mcondition.str结束时间), mcondition.str给药途径, mcondition.str药品剂型, _
            mcondition.str其它发药类型, mcondition.str记账人, mcondition.str住院号, mcondition.str床号, mcondition.str就诊卡, _
            mcondition.str姓名, mcondition.strNo, mcondition.lng病人ID, mcondition.str领药号, _
            CStr(strArr收发id(i)))
            
        If Not rsData.EOF Then
            '装载发药数据集
            If LoadSendRecord(rsData) = False Then
                Me.MousePointer = 0
                Exit Sub
            End If
        End If
    Next
    
    If mrsSendData.RecordCount > 0 Then
        '装载销账数据集
        Call RefreshChargeOffDetail
        '给子窗体传递数据集
        Call mfrmDetail.RefreshList(mListType.发药, mrsSendData, mrsChargeOff)
    End If
    
    Me.MousePointer = 0
    Call AviShow(Me, False)
    Exit Sub
errHandle:
    Me.MousePointer = 0
    Call AviShow(Me, False)
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub RefreshReturnDetail()
    '刷新退药明细列表
    Dim rsData As ADODB.Recordset
    Dim strSql退药人 As String
    Dim strSqlSelect As String
    Dim i As Integer
    Dim strArr收发id As Variant
    Dim ArrTmp As Variant
    Dim intCount As Integer
    Dim strTmp As String
    Dim str收发ID串 As String
    
    If Val(tvwList(mDeptType.退药).Tag) = 0 Then Exit Sub
    
    '根据部门列表已勾选的情况组织主要的条件
    If mrsDeptList Is Nothing Then Exit Sub
    mrsDeptList.Filter = ""
    With mrsDeptList
        Do While Not .EOF
            If !执行状态 = 1 Then
                If InStr(1, "," & str收发ID串 & ",", "," & !收发ID & ",") = 0 Then
                    str收发ID串 = str收发ID串 & IIf(str收发ID串 = "", "", ",") & !收发ID
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    If str收发ID串 = "" Then Exit Sub
    
    '分解收发ID串
    '收发ID串大于4K时分成小于4K的串（绑定变量时，最大变量长度为4K字符）
    strArr收发id = Array()
    ArrTmp = Split(str收发ID串 & ",", ",")
    intCount = UBound(ArrTmp)
    
    '查询提示
    If WarRecoredCount(intCount) = False Then Exit Sub
    
    If Len(str收发ID串) >= 4000 Then
        For i = 0 To intCount
            If ArrTmp(i) <> "" Then
                If Len(IIf(strTmp = "", "", strTmp & ",") & ArrTmp(i)) >= 4000 Then
                    ReDim Preserve strArr收发id(UBound(strArr收发id) + 1)
                    strArr收发id(UBound(strArr收发id)) = strTmp
                    strTmp = ArrTmp(i)
                Else
                    strTmp = IIf(strTmp = "", "", strTmp & ",") & ArrTmp(i)
                End If
            End If
                   
            If i = intCount Then
                ReDim Preserve strArr收发id(UBound(strArr收发id) + 1)
                strArr收发id(UBound(strArr收发id)) = strTmp
            End If
        Next
    Else
        ReDim Preserve strArr收发id(UBound(strArr收发id) + 1)
        strArr收发id(UBound(strArr收发id)) = str收发ID串
    End If
    
    
    '单位设置
    Select Case mParams.strUnit
    Case "售价单位"
        strSqlSelect = "X.计算单位 单位,1 包装,"
    Case "门诊单位"
        strSqlSelect = "D.门诊单位 单位,D.门诊包装 包装,"
    Case "住院单位"
        strSqlSelect = "D.住院单位 单位,D.住院包装 包装,"
    Case "药库单位"
        strSqlSelect = "D.药库单位 单位,D.药库包装 包装,"
    End Select
        
    strSqlSelect = strSqlSelect & IIf(mParams.int药品名称显示 = 0 Or mParams.int药品名称显示 = 2, "NVL(A.名称,'')", "Decode(A.名称,Null,'',X.名称)") & " As 其它名, " & _
             "'['||X.编码||']'||" & IIf(mParams.int药品名称显示 = 1, "NVL(A.名称,X.名称)", "X.名称") & " As 品名," & _
             "X.编码" & " As 药品编码," & IIf(mParams.int药品名称显示 = 1, "NVL(A.名称,X.名称)", "X.名称") & " As 药品名称,"

    gstrSQL = " SELECT /*+rule*/ DISTINCT S.ID As 收发ID,S.单据,S.药品ID,S.NO,S.序号,S.扣率,H.ID As 领药部门ID,P.名称 科室,C.门诊标志,C.标识号,C.病人ID,Nvl(C.主页ID,0) As 主页ID,C.床号,Decode(Nvl(c.婴儿费,0), 0, Nvl(W.姓名, C.姓名), U.婴儿姓名) 姓名,Decode(Nvl(c.婴儿费,0), 0, Nvl(W.性别, C.性别), U.婴儿性别) 性别," & _
             " NVL(D.药房分批,0) 分批,Nvl(D.高危药品,0) As 高危药品,X.规格,T.毒理分类,TO_CHAR(Q.发送时间,'YYYY-MM-DD HH24:MI:SS') 发送时间," & _
             strSqlSelect & _
             " S.付数 付,S.实际数量 数量,S.已退数量,S.已发数量 准退数,DECODE(S.批号,NULL,'',S.批号)||DECODE(S.批次,NULL,'',0,'','('||S.批次||')') 批号,NVL(S.批次,0) 批次,to_char(S.效期,'yyyy-mm-dd') 效期," & _
             " S.零售价 单价,S.零售金额 金额,S.单量,S.频次,S.用法,S.摘要 说明,TO_CHAR(S.审核日期,'YYYY-MM-DD HH24:MI:SS') 发药时间,S.审核人,S.审核日期,可操作,C.医嘱序号,I.计算单位," & _
             " NVL(S.产地,NVL(X.产地,'')) 产地,S.原产地,nvl(M.审查结果,-1) 审查结果,M.禁忌药品说明,nvl(C.医嘱序号,-1) 医嘱id,S.领药人," & IIf(mParams.bln药品储备 = True, "L.", "'' ") & "库房货位, " & _
             " M.相关ID,c.序号 费用序号,Z.名称 As 英文名,0 As 转出, S.发药号,D.剂量系数,m.医生嘱托 " & _
             " FROM "
    gstrSQL = gstrSQL & _
             "          (SELECT A.ID,A.NO,A.单据,A.序号,A.药品ID,A.费用ID,A.批次,A.批号,A.效期,NVL(A.扣率,0) 扣率," & _
             "              NVL(A.付数,1) 付数,A.实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态," & _
             "              A.零售价 , A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID,1 可操作,A.产地,A.原产地," & _
             "              decode(nvl(A.领用人,''),'','',Decode(A.记录状态,1,'(领)'||A.领用人," & _
             "              decode(Mod(A.记录状态,3),0,'(领)'||A.领用人,1,'(领)'||A.领用人,2,'(退)'||A.领用人))) 领药人,Nvl(A.汇总发药号, 0) 发药号,A.填制人 " & _
             "          FROM 药品收发记录 A," & _
             "          (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
             "          FROM 药品收发记录 A,Table(Cast(f_Num2List([15]) As zlTools.t_NumList)) G " & _
             "          WHERE A.ID= G.Column_Value And A.审核人 IS NOT NULL" & _
             "          AND A.库房ID+0=[1] AND A.审核日期 BETWEEN [2] AND [3] " & _
             "          GROUP BY A.NO,A.单据,A.药品ID,A.序号) B" & _
             "          WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号 And A.审核人 IS NOT NULL AND (A.记录状态=1 OR MOD(A.记录状态,3)=0) "
    gstrSQL = gstrSQL & _
             "          UNION" & _
             "          SELECT A.ID,A.NO,A.单据,A.序号,A.药品ID,A.费用ID,A.批次,A.批号,A.效期,NVL(A.扣率,0)," & _
             "          NVL(A.付数,1) 付数,A.实际数量,0 已退数,0 已发数量,A.记录状态," & _
             "          A.零售价 , A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID," & _
             "          DECODE(A.记录状态,1,1,DECODE(MOD(A.记录状态,3),0,1,MOD(A.记录状态,3)+1)) 可操作,A.产地,A.原产地," & _
             "          decode(nvl(A.领用人,''),'','',Decode(A.记录状态,1,'(领)'||A.领用人," & _
             "          decode(Mod(A.记录状态,3),0,'(领)'||A.领用人,1,'(领)'||A.领用人,2,'(退)'||A.领用人))) 领药人,Nvl(A.汇总发药号, 0) 发药号,A.填制人 " & _
             "          FROM 药品收发记录 A,Table(Cast(f_Num2List([15]) As zlTools.t_NumList)) G " & _
             "          WHERE A.ID= G.Column_Value And A.审核人 IS NOT NULL AND NOT (记录状态=1 OR MOD(记录状态,3)=0)" & _
             "          AND A.库房ID+0=[1] AND A.审核日期 BETWEEN [2] AND [3] " & _
             "          ) S,"
    gstrSQL = gstrSQL & "" & _
             "      住院费用记录 C,部门表 P,药品规格 D,收费项目目录 X,收费项目别名 A,药品特性 T,诊疗项目目录 I,病人医嘱记录 M,病人医嘱发送 Q,病人信息 R,病案主页 W," & IIf(mParams.bln药品储备 = True, "药品储备限额 L,", "") & "诊疗项目别名 Z,部门表 H, 病人新生儿记录 U "
     
    '''where
    gstrSQL = gstrSQL & " WHERE S.药品ID=D.药品ID And C.病人id = R.病人id And C.病人id=W.病人id And C.主页id=W.主页id AND D.药名ID=T.药名ID AND d.药品ID=x.ID AND C.病人科室ID+0=P.ID AND D.药名ID=I.ID and C.医嘱序号=M.ID(+)  and C.医嘱序号=Q.医嘱id(+) And c.No = q.No(+) And c.病人id = u.病人id(+) And c.婴儿费 = u.序号(+) And C.主页id=U.主页id(+) " & _
             " And D.药名id = Z.诊疗项目id(+) And Z.性质(+) = 2 " & IIf(mParams.bln药品储备 = True, " And S.药品ID=L.药品ID(+) And R.病人id=w.病人id And Nvl(S.库房ID,[1])=L.库房ID(+) ", "") & _
             " AND D.药品ID=A.收费细目ID(+) AND A.性质(+)=3 " & _
             " AND S.费用ID=C.ID And S.单据 IN(9,10) " & _
             " AND S.审核人 IS NOT NULL And s.对方部门id + 0 = h.Id "
    
    '录入信息
    If mcondition.str住院号 <> "" Then
        gstrSQL = gstrSQL & " And W.住院号 = [8] "
    ElseIf mcondition.str床号 <> "" Then
        gstrSQL = gstrSQL & " And R.当前床号 = [9] "
    ElseIf mcondition.str就诊卡 <> "" Then
        gstrSQL = gstrSQL & " And R.就诊卡号 = [10] "
    ElseIf mcondition.str姓名 <> "" Then
        gstrSQL = gstrSQL & " And C.姓名 = [11] "
    ElseIf mcondition.strNo <> "" Then
        gstrSQL = gstrSQL & " And C.NO = [12] "
    ElseIf mcondition.lng病人ID <> -1 Then
        gstrSQL = gstrSQL & " And C.病人ID = [13] "
    ElseIf mcondition.cur发药号 <> 0 Then
        gstrSQL = gstrSQL & " And S.发药号 = [14] "
    End If
    
    '操作模式:0-所有,1-记帐单,2-记帐表
    If mcondition.int操作模式 = 1 Then
        gstrSQL = gstrSQL & " And S.单据=9"
    ElseIf mcondition.int操作模式 = 2 Then
        gstrSQL = gstrSQL & " And S.单据=10"
    End If
    
    '记账人
    If mcondition.str记账人 <> "所有记帐人" Then
        gstrSQL = gstrSQL & " And S.填制人 = [7] "
    End If
    
    '医嘱类型:0-所有,1-长嘱,2-临嘱,3-普通
    '用单量是否填写区分是否医嘱产生的药品单据
    If mcondition.int医嘱类型 = 0 Then
    ElseIf mcondition.int医嘱类型 = 1 Then
        gstrSQL = gstrSQL & " And S.扣率 Is Not Null And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '0_' And Nvl(C.医嘱序号,0) + 0 >0 "
    ElseIf mcondition.int医嘱类型 = 2 Then
        gstrSQL = gstrSQL & " And S.扣率 Is Not Null And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '1_' And Nvl(C.医嘱序号,0) + 0 >0 "
    ElseIf mcondition.int医嘱类型 = 3 Then
        gstrSQL = gstrSQL & " And (Nvl(C.医嘱序号,0) + 0 =0 Or S.扣率 Is Null) "
    ElseIf mcondition.int医嘱类型 = 4 Then
        gstrSQL = gstrSQL & " And S.扣率 Is Not Null And (Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '0_' Or Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '1_') And Nvl(C.医嘱序号,0) + 0 > 0 "
    End If
    
    '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药,3-不含自取药,4-仅含自取药,5-院内用药(不包括离院带药和自取药),6-离院带药和自取药
    If mcondition.int发药类型 = 0 Then
    ElseIf mcondition.int发药类型 = 1 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3'"
    ElseIf mcondition.int发药类型 = 2 Then
        gstrSQL = gstrSQL & " And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3'"
    ElseIf mcondition.int发药类型 = 3 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf mcondition.int发药类型 = 4 Then
        gstrSQL = gstrSQL & " And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf mcondition.int发药类型 = 5 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3' And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf mcondition.int发药类型 = 6 Then
        gstrSQL = gstrSQL & " And (Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3' Or Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4')"
    End If
    
    '病人类型：病人或婴儿
    If mcondition.int病人类型 = 0 Then
        gstrSQL = gstrSQL & " And Nvl(C.婴儿费, 0) = 0 "
    ElseIf mcondition.int病人类型 = 1 Then
        gstrSQL = gstrSQL & " And Nvl(C.婴儿费, 0) > 0 "
    End If
    
    '给药途径
    If mcondition.str给药途径 <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [4] || ',',',' || S.用法 || ',') > 0 "
    End If
    
    '药品剂型
    If mcondition.str药品剂型 <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [5] || ',',',' || T.药品剂型 || ',') > 0 "
    End If
    
    '其它发药类型
    If mcondition.str其它发药类型 <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [6] || ',',',' || D.发药类型 || ',') > 0 "
    End If
    
    Dim blnMoved As Boolean
    Dim strsql As String
    '判断是否存在部分数据已转出
    blnMoved = Sys.IsMovedByDate(mcondition.str开始时间)
    If blnMoved Then
        'SQL按记录序号汇总，因任何一笔明细要么在线，要么后备，因此，以UNION方式处理
        strsql = gstrSQL
        strsql = Replace(strsql, "药品收发记录", "H药品收发记录")
        strsql = Replace(strsql, "住院费用记录", "H住院费用记录")
        strsql = Replace(strsql, "0 As 转出", "1 As 转出")
        
        gstrSQL = gstrSQL & " UNION ALL " & strsql
    End If
    
    gstrSQL = gstrSQL & " Order By No,单据,审核日期"
    
    On Error GoTo errHandle
    
    Me.MousePointer = 11
    Call AviShow(Me)
    Call InitReturnRec
    
    '根据收发ID串的数组数目，循环执行
    For i = 0 To UBound(strArr收发id)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mcondition.lng药房id, _
            CDate(mcondition.str开始时间), _
            CDate(mcondition.str结束时间), _
            mcondition.str给药途径, _
            mcondition.str药品剂型, _
            mcondition.str其它发药类型, _
            mcondition.str记账人, _
            mcondition.str住院号, _
            mcondition.str床号, _
            mcondition.str就诊卡, _
            mcondition.str姓名, _
            mcondition.strNo, _
            mcondition.lng病人ID, _
            mcondition.cur发药号, _
            CStr(strArr收发id(i)))
        
        If Not rsData.EOF Then
            If LoadReturnRecord(rsData) = False Then
                Me.MousePointer = 0
                Exit Sub
            End If
        End If
    Next
    
    If mrsReturnData.RecordCount > 0 Then
        Call mfrmDetail.RefreshList(mListType.退药, mrsReturnData)
    End If
    
    Me.MousePointer = 0
    Call AviShow(Me, False)
    Exit Sub
errHandle:
    Me.MousePointer = 0
    Call AviShow(Me, False)
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

'
Private Sub GetParams()
    '取该模块用到的参数信息
    Dim int金额 As Integer
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "select 精度 from 药品卫材精度 where 性质=0 and 类别 = 1 And 内容 = 4 And 单位 = 5"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询金额精度")
    If rstemp.RecordCount = 0 Then
        int金额 = 2
    Else
        int金额 = rstemp!精度
    End If
    With mParams
        '参数表中的系统参数
        .bln允许未审核处方发药 = (gtype_UserSysParms.P6_未审核记帐处方发药 = 1)
        .bln门诊医嘱先作废后退药 = (gtype_UserSysParms.P68_门诊药嘱先作废后退药 = 1)
        .int金额保留位数 = int金额
        .bln审核划价单 = True
        .int效期显示方式 = gtype_UserSysParms.P149_效期显示方式
        .int药品名称显示 = gint药品名称显示
        .bln启用审方 = (gtype_UserSysParms.P240_药房处方审查 = 2 Or gtype_UserSysParms.P240_药房处方审查 = 3)
    
        '参数设置中的参数
        '基础参数
        .lng药房id = Val(zlDatabase.GetPara("发药药房", glngSys, 1342))
        .int操作模式 = Val(zlDatabase.GetPara("操作模式", glngSys, 1342))
        .intDays = Val(zlDatabase.GetPara("查询天数", glngSys, 1342)) - 1
        .int自动刷新未发药清单 = Val(zlDatabase.GetPara("自动刷新未发药清单", glngSys, 1342))
        .str记帐人 = zlDatabase.GetPara("记帐人", glngSys, 1342, "所有记帐人")
        .bln汇总发药 = (Val(zlDatabase.GetPara("发药时汇总退药销帐记录", glngSys, 1342, 0)) = 1)
        .bln汇总显示 = (Val(zlDatabase.GetPara("按科室汇总显示汇总清单", glngSys, 1342)) = 1)
        .bln领药人签名 = (Val(zlDatabase.GetPara("领药人签名", glngSys, 1342)) = 1)
        .bln退药人签名 = (Val(zlDatabase.GetPara("退药人签名", glngSys, 1342)) = 1)
        .bln审核出院销账申请 = (Val(zlDatabase.GetPara("审核出院病人的销账申请", glngSys, 1342, 0)) = 1)
        
        '辅助参数
        .bln缺药检查 = (Val(zlDatabase.GetPara("缺药检查", glngSys, 1342, 1)) = 1)
        .int自动打印 = Val(zlDatabase.GetPara("自动打印", glngSys, 1342))
        .bln药品储备 = (Val(zlDatabase.GetPara("库房货位及库存限量提示", glngSys, 1342, 0)) = 1)
        .str毒理分类 = zlDatabase.GetPara("毒理分类", glngSys, 1342)
        .str价值分类 = zlDatabase.GetPara("价值分类", glngSys, 1342)
        .str高危分类 = zlDatabase.GetPara("高危分类", glngSys, 1342, "")
        .str高危发放 = zlDatabase.GetPara("高危药品发放", glngSys, 1342, "")
        .int退药清单打印 = Val(zlDatabase.GetPara("打印退药清单", glngSys, 1342))
        .intCheck = Val(zlDatabase.GetPara("审核该药房的所有数据", glngSys, 1345))
        
        .int医嘱类型 = Val(zlDatabase.GetPara("医嘱类型", glngSys, 1342))
        
        .bln待发单据 = (Val(GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\Frm部门发药管理", "显示退药待发单据", 1)) = 1)
        
        .int药品名称编码显示 = GetDrugFormat
        .int发药时审核医嘱 = Val(zlDatabase.GetPara("发药时审核医嘱", glngSys, 1342))
        .bln检查存储库房 = (Val(zlDatabase.GetPara("发药时检查存储库房", glngSys, 1342)) = 1)
        .bln检查销帐申请 = (Val(zlDatabase.GetPara("发药时检查销帐申请数据", glngSys, 1342)) = 1)
                
        '查询提示
        .int查询发药天数 = Val(zlDatabase.GetPara("查询发药天数", glngSys, 1342, 7))
        .int查询退药天数 = Val(zlDatabase.GetPara("查询退药天数", glngSys, 1342, 3))
        .lng最大记录数 = Val(zlDatabase.GetPara("查询明细记录数", glngSys, 1342, 3000))
        .int页签 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品部门发药管理", "当前页签", 0))
        .bln保持上一次页签 = (GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品部门发药管理", "保持上一次窗体关闭时的页签", 0) = 1)
        
        '库存检查
        .IntCheckStock = MediWork_GetCheckStockRule(.lng药房id)
    
        '库房单位
        .strUnit = GetSpecUnit(.lng药房id, gint住院药房)
        
        '是否启用PASS
        .blnStarPass = gintPass <> 0 And mPrives.bln合理用药监测 = True
        
        '配置中心
        .bln配制中心 = CheckIsCenter(.lng药房id)
        
        '参数设置：来源科室
        .strSourceDep = zlDatabase.GetPara("来源科室", glngSys, 1342)
        
        '注册表参数
        .intFont = Val(zlDatabase.GetPara("字体", glngSys, 1342))
        .StrFindStyle = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
        
        '注册表参数：包装机相关
        .int暂停传送 = Val(GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "部门发药管理\包装机设置", "暂停传送", "0"))
        .int暂停传送 = IIf(.int暂停传送 = 1, 1, 0)
        .str包装机单据 = GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "部门发药管理\包装机设置", "单据类型", "11")
        .str包装机剂型 = GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "部门发药管理\包装机设置", "选择剂型", "所有")
        
        mblnIs配置中心 = Is配置中心(.lng药房id)
        Call GetDrugDigit(.lng药房id, "药品部门发药", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshReturnDept()
    '刷新退药部门列表
    Dim rsData As ADODB.Recordset
    Dim strDanger As String
    Dim strToxicology As String
    
    '''select
    gstrSQL = "Select" & IIf(mParams.strSourceDep = "", "", "/*+rule*/") & "  H.ID, H.名称 As 科室名称, S.汇总发药号 As 发药号, Decode(Nvl(c.婴儿费,0), 0, Nvl(b.姓名, c.姓名), z.婴儿姓名) 姓名, B.病人ID, Decode(Nvl(c.婴儿费,0), 0, Nvl(p.性别, c.性别), z.婴儿性别) 性别, Decode(Nvl(c.婴儿费,0), 0, p.年龄, Ceil(Sysdate - z.出生时间) || '天') 年龄, S.单据, S.NO, S.药品id, " & _
        " Decode(Nvl(C.医嘱序号, 0), 0, 0, 1) 医嘱序号, C.门诊标志, Nvl(S.扣率, 0) 扣率, S.ID As 收发id, S.填制日期, Nvl(B.当前床号,'') As 床号,W.颜色,c.婴儿费 "
    
    '''from
    gstrSQL = gstrSQL & " From 药品收发记录 S, 住院费用记录 C, 病人信息 B, 药品规格 D, 药品特性 T, 病案主页 P, 部门表 H,病人类型 W, 病人新生儿记录 Z " & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([17]) As zlTools.t_NumList)) E ")
    
    '''where
    gstrSQL = gstrSQL & " Where S.对方部门id = H.ID" & IIf(mParams.strSourceDep = "", "", " And S.对方部门id=E.Column_Value ") & _
        " And C.病人id = B.病人id And C.病人id=P.病人id And C.主页id=P.主页id And C.NO = S.NO And S.费用id = C.ID And c.病人id = z.病人id(+) And c.婴儿费 = z.序号(+) And C.主页id=Z.主页id(+) " & _
        " And S.库房id = C.执行部门id And S.药品id = D.药品id And D.药名id = T.药名id And P.病人类型=W.名称(+) " & _
        " And (H.撤档时间 Is Null Or H.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
        " And S.审核日期 Between [2] And [3] And S.审核人 IS NOT NULL "
    
    '站点控制
    If mstrDeptNode <> "" Then
        gstrSQL = gstrSQL & " And (H.站点 = [16] Or H.站点 Is Null) "
    End If
    
    '当前药房
    gstrSQL = gstrSQL & " And S.库房id + 0 = [1] "
    
    '录入信息
    If mcondition.str住院号 <> "" Then
        gstrSQL = gstrSQL & " And P.住院号 = [4] "
    ElseIf mcondition.str床号 <> "" Then
        '由于床号不唯一，转为通过病人ID来查询
        gstrSQL = gstrSQL & " And B.病人ID+0 = [9] "
    ElseIf mcondition.str就诊卡 <> "" Then
        gstrSQL = gstrSQL & " And B.就诊卡号 = [6] "
    ElseIf mcondition.str姓名 <> "" Then
        gstrSQL = gstrSQL & " And P.姓名 = [7] "
    ElseIf mcondition.strNo <> "" Then
        gstrSQL = gstrSQL & " And S.NO = [8] "
    ElseIf mcondition.lng病人ID <> -1 Then
        gstrSQL = gstrSQL & " And B.病人ID+0 = [9] "
    ElseIf mcondition.cur发药号 <> 0 Then
        gstrSQL = gstrSQL & " And S.汇总发药号 = [10] "
    ElseIf mcondition.lng领药部门ID <> -1 Then
        gstrSQL = gstrSQL & " And S.对方部门id + 0 = [11] "
    End If
    
    '操作模式:0-所有,1-记帐单,2-记帐表
    If mcondition.int操作模式 = 0 Then
        gstrSQL = gstrSQL & " And S.单据 IN(9,10)"
    ElseIf mcondition.int操作模式 = 1 Then
        gstrSQL = gstrSQL & " And S.单据=9"
    ElseIf mcondition.int操作模式 = 2 Then
        gstrSQL = gstrSQL & " And S.单据=10"
    End If
    
    '记账人
    If mcondition.str记账人 <> "所有记帐人" Then
        gstrSQL = gstrSQL & " And S.填制人 = [12] "
    End If
    
    '医嘱类型:0-所有,1-长嘱,2-临嘱,3-普通
    '用单量是否填写区分是否医嘱产生的药品单据
    If mcondition.int医嘱类型 = 0 Then
    ElseIf mcondition.int医嘱类型 = 1 Then
        gstrSQL = gstrSQL & " And S.扣率 Is Not Null And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '0_' And Nvl(C.医嘱序号,0) + 0 >0 "
    ElseIf mcondition.int医嘱类型 = 2 Then
        gstrSQL = gstrSQL & " And S.扣率 Is Not Null And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '1_' And Nvl(C.医嘱序号,0) + 0 >0 "
    ElseIf mcondition.int医嘱类型 = 3 Then
        gstrSQL = gstrSQL & " And (Nvl(C.医嘱序号,0) + 0 =0 Or S.扣率 Is Null) "
    ElseIf mcondition.int医嘱类型 = 4 Then
        gstrSQL = gstrSQL & " And S.扣率 Is Not Null And (Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '0_' Or Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '1_') And Nvl(C.医嘱序号,0) + 0 > 0 "
    End If
    
    '离院带药:'0-所有,1-不含离院带药,2-仅含离院带药,3-不含自取药,4-仅含自取药,5-院内用药(不包括离院带药和自取药),6-离院带药和自取药
    If mcondition.int发药类型 = 0 Then
    ElseIf mcondition.int发药类型 = 1 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3'"
    ElseIf mcondition.int发药类型 = 2 Then
        gstrSQL = gstrSQL & " And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3'"
    ElseIf mcondition.int发药类型 = 3 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf mcondition.int发药类型 = 4 Then
        gstrSQL = gstrSQL & " And Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf mcondition.int发药类型 = 5 Then
        gstrSQL = gstrSQL & " And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3' And Not Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4'"
    ElseIf mcondition.int发药类型 = 6 Then
        gstrSQL = gstrSQL & " And (Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_3' Or Ltrim(To_Char(Nvl(S.扣率,0),'00')) Like '_4')"
    End If
    
    '病人类型：病人或婴儿
    If mcondition.int病人类型 = 0 Then
        gstrSQL = gstrSQL & " And Nvl(C.婴儿费, 0) = 0 "
    ElseIf mcondition.int病人类型 = 1 Then
        gstrSQL = gstrSQL & " And Nvl(C.婴儿费, 0) > 0 "
    End If
    
    '给药途径
    If mcondition.str给药途径 <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [13] || ',',',' || S.用法 || ',') > 0 "
    End If
    
    '药品剂型
    If mcondition.str药品剂型 <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [14] || ',',',' || T.药品剂型 || ',') > 0 "
    End If
    
    '其它发药类型
    If mcondition.str其它发药类型 <> "" Then
        gstrSQL = gstrSQL & " And Instr(',' || [15] || ',',',' || D.发药类型 || ',') > 0 "
    End If
    
    '科室类型
    If Trim(txtInput.Text) = "" Then
        If mParams.intShowDept = 1 Then
            gstrSQL = gstrSQL & " And H.id In (Select 部门id From 部门性质说明 Where 工作性质 = '临床' And 服务对象 In (2, 3)) "
        ElseIf mParams.intShowDept = 2 Then
            gstrSQL = gstrSQL & " And H.id In (Select 部门ID From 部门性质说明 Where 工作性质 In ('检查','检验','治疗','手术','营养') And 服务对象 IN(2,3)) "
        ElseIf mParams.intShowDept = 3 Then
            gstrSQL = gstrSQL & " And H.id In (Select 部门ID From 部门性质说明 Where 工作性质='护理' And 服务对象 IN(2,3)) "
        End If
    End If
    
    '排除已在输液配置中心管理中产生的单据
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From 输液配药内容 Y,药品收发记录 Z Where  Y.收发id=z.id and  Z.NO= S.NO) "
    
    '高危药品
    If chkDanger.Value = 1 Then
        If chkDangerType(0).Value = 1 Then strDanger = IIf(strDanger = "", 1, strDanger & "," & 1)
        If chkDangerType(1).Value = 1 Then strDanger = IIf(strDanger = "", 2, strDanger & "," & 2)
        If chkDangerType(2).Value = 1 Then strDanger = IIf(strDanger = "", 3, strDanger & "," & 3)
    End If
    If strDanger <> "" Then gstrSQL = gstrSQL & " And Instr(',' || [18] || ',' , ',' || Nvl(D.高危药品,0) || ',') > 0 "
    
    '毒理分类
    If Me.chkToxicologyType.Value = 1 Then
        If Me.chkToxicology(0).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(0).Caption, strToxicology & "," & Me.chkToxicology(0).Caption)
        If Me.chkToxicology(1).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(1).Caption, strToxicology & "," & Me.chkToxicology(1).Caption)
        If Me.chkToxicology(2).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(2).Caption, strToxicology & "," & Me.chkToxicology(2).Caption)
        If Me.chkToxicology(3).Value = 1 Then strToxicology = IIf(strToxicology = "", Me.chkToxicology(3).Caption, strToxicology & "," & Me.chkToxicology(3).Caption)
    End If
    
    If strToxicology <> "" Then gstrSQL = gstrSQL & " And Instr(',' || [19] || ',' , ',' || T.毒理分类 || ',') > 0 "
    
    '''order by
    gstrSQL = gstrSQL & " Order By H.名称, 发药号, B.姓名, S.NO "
    
    On Error GoTo errHandle
    
    Me.MousePointer = 11
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "提取待发药科室汇总", _
        mcondition.lng药房id, _
        CDate(mcondition.str开始时间), _
        CDate(mcondition.str结束时间), _
        mcondition.str住院号, _
        mcondition.str床号, _
        mcondition.str就诊卡, _
        mcondition.str姓名, _
        mcondition.strNo, _
        mcondition.lng病人ID, _
        mcondition.cur发药号, _
        mcondition.lng领药部门ID, _
        mcondition.str记账人, _
        mcondition.str给药途径, _
        mcondition.str药品剂型, _
        mcondition.str其它发药类型, _
        mstrDeptNode, _
        mParams.strSourceDep, _
        strDanger, _
        strToxicology)
    
    '更新部门树表
    Call GetReturnDeptTreeView(rsData)
    
    '更新部门树表对应的收发记录数据集
    Call GetDeptListRecord(rsData)
    
    Me.MousePointer = 0
    Exit Sub
errHandle:
    Me.MousePointer = 0
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckAdvice(ByVal rsData As ADODB.Recordset) As Boolean
    '先检查是否允许退药（医嘱）
    Dim rsTmp As ADODB.Recordset
    
    CheckAdvice = False
    On Error GoTo errHandle
    If mParams.bln门诊医嘱先作废后退药 = True Then
        CheckAdvice = True
        Exit Function
    End If
    
    With rsData
        .Filter = "执行状态=" & mState.退药
        
        Do While Not .EOF
            gstrSQL = "select 扣率 From 药品收发记录 Where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是临嘱]", CLng(!收发ID))
            
            If (rsTmp!扣率 Like "1*") Then       '临嘱
                gstrSQL = "Select Nvl(医嘱序号,0) 医嘱序号,Nvl(门诊标志,1) 门诊标志 From 住院费用记录 Where ID=(Select 费用ID From 药品收发记录 Where ID=[1])"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是医嘱]", CLng(!收发ID))
                
                If Not rsTmp.EOF Then
                    If (rsTmp!门诊标志 = 1 Or rsTmp!门诊标志 = 4) And rsTmp!医嘱序号 <> 0 Then
                        gstrSQL = "Select decode(医嘱状态,4,1,0) 作废 From 病人医嘱记录 Where ID=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断该医嘱是否作废]", CLng(rsTmp!医嘱序号))
                        
                        If rsTmp!作废 = 0 Then
                            MsgBox "[" & " & !NO & " & "]中的药品[" & !品名 & "]对应的医嘱还未作废，不能退药！", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    CheckAdvice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitComandBars()
    '初始化菜单：加载全部菜单，工具栏，弹出菜单等
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    Dim intCount As Integer
    Dim strCardName As String
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = frmPublic.imgPublic.Icons
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsMain.ActiveMenuBar.Title = "菜单"
    Me.cbsMain.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.Id = mconMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Preview, "预览(&V)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Print, "打印(&P)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Excel, "输出到&Excel…")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Dept_BillPrint, "单据打印(&B)")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Dept_BillPrintTotal, "打印汇总清单(&C)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Dept_BillPrintRestore, "打印退药通知单(&R)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Dept_BillPrintWait, "打印药品摆药单(&W)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Parameter, "参数设置(&T)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "退出(&X)")
        cbrControlMain.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.Id = mconMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Verify, "发药(&V)")
        cbrControlMain.Visible = mPrives.bln发药
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Reject, "拒发确认(&H)")
        cbrControlMain.Visible = mPrives.bln拒发
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, "拒发恢复(&H)")
        cbrControlMain.Visible = mPrives.bln拒发
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Return, "退药(&R)")
        cbrControlMain.Visible = mPrives.bln退药
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_EMR, "病案查询(&Z)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_ReturnOther, "退其它药房的处方(&T)")
        cbrControlMain.Visible = mPrives.bln退其它药房的处方
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_VerifySign, "验证签名(&S)")
        cbrControlMain.Visible = gblnESign部门发药
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_ReVerify, "药品退药销账(&B)")
        cbrControlMain.Visible = mPrives.bln退药销帐
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_StopFlag, "停止发药标记(&S)")
        cbrControlMain.Visible = (mPrives.bln停止发药 = True Or mPrives.bln恢复发药 = True)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Packer, "分包机接口设置(&P)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Hot_IC, "读IC卡(&I)")
        cbrControlMain.Visible = False
        
        '扩展接口
        Call zlPlugIn_SetMenu(glngSys, glngModul, mobjPlugIn, cbrMenuBar.CommandBar.Controls, mconMenu_Edit_PlugIn)
    End With
    
'    '自动化发药设置菜单
'    If Not gobjPackerMZ Is Nothing Then
'        Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_AutoSend, "药房自动化设置(&V)", -1, False)
'        cbrMenuBar.Id = mconMenu_AutoSend
'    End If
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.Id = mconMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_View_ToolBar, "工具栏(&T)")
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Size, "大图标(&B)", -1, False)
        cbrControl.Checked = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_StatusBar, "状态栏(&S)")
        cbrControlMain.Checked = True
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_View_FontSize, "字体(&F)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_1, "小字体(&S)", -1, False)
        If mParams.intFont = 0 Then cbrControl.Checked = True
        cbrControl.Parameter = 0
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_2, "中字体(&M)", -1, False)
        If mParams.intFont = 1 Then cbrControl.Checked = True
        cbrControl.Parameter = 1
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_3, "大字体(&B)", -1, False)
        If mParams.intFont = 2 Then cbrControl.Checked = True
        cbrControl.Parameter = 2
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Find, "查找(&L)")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_FindNext, "查找下一条(&N)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_SelAll, "全选(&A)")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_ClsAll, "全清(&C)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Refresh, "刷新(&R)")
        cbrControlMain.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.Id = mconMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_Help, "帮助主题(&H)")
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_Help_Web, "&WEB上的中联")
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Home, "中联主页(&H)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Forum, "中联论坛(&F)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_About, "关于(&A)…")
        cbrControlMain.BeginGroup = True
    End With
    
    '主菜单右侧的消息提示下拉菜单，接收到消息时动态增加
    With cbsMain.ActiveMenuBar.Controls
        Set cbrMenuBar = .Add(xtpControlPopup, mconMenu_File_Message, "↓消息提醒")
        cbrMenuBar.Id = mconMenu_File_Message
        cbrMenuBar.Flags = xtpFlagRightAlign
        cbrMenuBar.Visible = mPrives.bln退药销帐
    End With
        
    '快键绑定
    With Me.cbsMain.KeyBindings
'        .Add FCONTROL, Asc("S"), mconMenu_Edit_Save
'        .Add FCONTROL, Asc("Z"), mconMenu_Edit_Untread
'        .Add FCONTROL, Asc("M"), mconMenu_Edit_Modify
'        .Add FSHIFT, VK_DELETE, mconMenu_Edit_Delete
        .Add FCONTROL, VK_F4, mconMenu_Edit_Dept_Hot_IC
        .Add 0, VK_F12, mconMenu_File_Parameter
        .Add 0, VK_F5, mconMenu_View_Refresh
        .Add 0, VK_F1, mconMenu_Help_Help
        .Add FSHIFT, 65, mconMenu_Edit_Dept_Verify
    End With

    '设置不常用菜单
    With Me.cbsMain.Options
        .AddHiddenCommand mconMenu_File_PrintSet
        .AddHiddenCommand mconMenu_File_Excel
    End With
    
    '设置部门列表项目弹出菜单
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ListPopup, "项目(&I)", -1, False)
    cbrMenuBar.Id = mconMenu_ListPopup
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_ShowReject, "包含拒发药品(&R)")
        cbrControlMain.Checked = mParams.blnShowReject
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_OnlyShowDept, "仅显示科室(&0)")
        cbrControlMain.Checked = mParams.blnOnlyShowDept
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_ShowOther, "显示详细信息(&1)")
        cbrControlMain.Checked = Not mParams.blnOnlyShowDept
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_ShowAll, "显示所有科室(&A)")
        cbrControlMain.Checked = (mParams.intShowDept = 0)
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_ShowClin, "显示临床科室(&C)")
        cbrControlMain.Checked = (mParams.intShowDept = 1)
    
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_ShowTech, "显示医技科室(&T)")
        cbrControlMain.Checked = (mParams.intShowDept = 2)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_ShowArea, "显示病人病区(&B)")
        cbrControlMain.Checked = (mParams.intShowDept = 3)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_List_Sort, "科室按医嘱发送时间排序(&D)")
        cbrControlMain.Checked = mParams.blnSort
        cbrControlMain.BeginGroup = True
    End With
    
    '设置给药途径分类弹出菜单
    Set rsData = DeptSendWork_Get给药途径分类
    
    If rsData.RecordCount > 0 Then
        Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_TypePopup, "分类(&T)", -1, False)
        cbrMenuBar.Id = mconMenu_TypePopup
        cbrMenuBar.Visible = False
        
        mTypeCount = rsData.RecordCount
        With cbrMenuBar.CommandBar.Controls
            For i = 1 To rsData.RecordCount
                Set cbrControlMain = .Add(xtpControlButton, mconMenu_TypePopup + i, rsData!分类)
                rsData.MoveNext
            Next
        End With
    End If
    
    '设置部门列表中病人排序菜单
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_SortPopup, "病人排序(&P)", -1, False)
    cbrMenuBar.Id = mconMenu_SortPopup
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_SortPopup_ByName, "按姓名排序(&0)")
        cbrControlMain.Checked = (mParams.int病人排序 = 1)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_SortPopup_ByBedNo, "按床位排序(&1)")
        cbrControlMain.Checked = (mParams.int病人排序 = 2)
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Preview, "预览")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Print, "打印")
        
        If mblnCustomCheck = True Then
            Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_CustomCheck, IIf(mstrCustomCheckName = "", "审核", mstrCustomCheckName))
            cbrControlMain.BeginGroup = True
        End If
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Verify, "发药")
        cbrControlMain.Visible = mPrives.bln发药
        cbrControlMain.BeginGroup = Not mblnCustomCheck
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Reject, "拒发")
        cbrControlMain.Visible = mPrives.bln拒发
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, "恢复")
        cbrControlMain.Visible = mPrives.bln拒发
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_Return, "退药")
        cbrControlMain.Visible = mPrives.bln退药
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_VerifySign, "验证签名")
        cbrControlMain.Visible = gblnESign部门发药

        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_ReVerify, "销帐")
        cbrControlMain.Visible = mPrives.bln退药销帐
        cbrControlMain.BeginGroup = True
        
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_EMR, "病案查询")
'        cbrControlMain.BeginGroup = True
        
        '电子病案查阅
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Dept_MedicalRecord, "电子病案查阅")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = mPrives.bln电子病案查阅
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Refresh, "刷新")
        cbrControlMain.BeginGroup = True
        
        '外挂接口
        Call zlPlugIn_SetToolbar(glngSys, glngModul, mobjPlugIn, cbrToolBar.Controls, mconMenu_Edit_PlugIn)

        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "退出")
        cbrControlMain.BeginGroup = True
    End With
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
End Sub


Private Sub InitPanes()
    '初始化分栏控件
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
'    Me.dkpMain.Options.DefaultPaneOptions = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
    
    Dim objPaneCon As Pane
    Dim objPaneList As Pane
    Dim objPaneDetail As Pane
    
    Set objPaneCon = Me.dkpMain.CreatePane(mconPane_Dept_Condition, 225, 100, DockLeftOf, Nothing)
    objPaneCon.Title = "过滤条件"
    objPaneCon.Options = PaneNoCloseable Or PaneNoFloatable
'    objPaneCon.MaxTrackSize.SetSize 290, 500
    
'    Set objPaneList = Me.dkpMain.CreatePane(mconPane_SelDept, 290, 250, DockBottomOf, objPaneCon)
'    objPaneList.Title = "待发科室"
'    objPaneList.Options = PaneNoCloseable Or PaneNoFloatable
End Sub
Private Sub InitTabControl()
    '初始化分页控件
    With Me.tbcDetail
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem(0, "未发药品清单(&0)", mfrmDetail.hWnd, 0).Tag = "未发药品清单_"
        .InsertItem(1, "汇总清单(&1)", mfrmDetail.hWnd, 0).Tag = "汇总清单_"
        .InsertItem(2, "缺药清单(&2)", mfrmDetail.hWnd, 0).Tag = "缺药清单_"
        .InsertItem(3, "拒发药清单(&3)", mfrmDetail.hWnd, 0).Tag = "拒发药清单_"
        
        If mPrives.bln查看已发药清单 = True Then
            .InsertItem(4, "已发药清单(&4)", mfrmDetail.hWnd, 0).Tag = "已发药清单_"
        End If
        
        .Item(1).Selected = True
        If mParams.int页签 <> 0 And mParams.bln保持上一次页签 Then
            .Item(mParams.int页签).Selected = True
        Else
            .Item(0).Selected = True
        End If
        
    End With
    
End Sub


Private Sub Load给药途径()
    Dim rsData As ADODB.Recordset
    
    Set rsData = DeptSendWork_Get给药途径
    
    With Lvw给药途径
        .ListItems.Clear
        .ListItems.Add , "_" & .ListItems.count + 1, "所有给药途径", 1, 1
        .ListItems(.ListItems.count).Checked = True
        Do While Not rsData.EOF
            .ListItems.Add , "_" & .ListItems.count + 1, rsData!用法, 1, 1
            .ListItems(.ListItems.count).Checked = True
            .ListItems(.ListItems.count).Tag = rsData!分类
            rsData.MoveNext
        Loop
    End With
End Sub
Private Function Load发药药房() As Boolean
    Dim rsData As ADODB.Recordset
    Dim strMsg As String
    Dim intIndex As Integer
    
    Set rsData = DeptSendWork_GetDrugstore(mstrPrivs, glngUserId, gstrNodeNo)
    
    If rsData.EOF Then
        If IsInString(mstrPrivs, "所有药房", ";") Then
            strMsg = "请初始化药房（部门管理）"
        Else
            strMsg = "你不是药房工作人员，不能操作本模块！"
        End If
        
        MsgBox strMsg, vbInformation, gstrSysName
        Load发药药房 = False
        Exit Function
    Else
        rsData.Filter = "id=" & mParams.lng药房id
        If rsData.EOF Then
            Call ResetParams(True)
        End If
        
        rsData.Filter = ""
        With cbo发药药房
            .Clear
            
            Do While Not rsData.EOF
                .AddItem rsData!名称
                .ItemData(.NewIndex) = rsData!Id
                
                If rsData!Id = mParams.lng药房id Then intIndex = .NewIndex
                
                rsData.MoveNext
            Loop
            
            .ListIndex = intIndex
            
            .Tag = .ItemData(intIndex)
        End With
        Load发药药房 = True
    End If
End Function

Private Sub Load药品剂型(ByVal lng药房id As Long)
    Dim rsData As ADODB.Recordset
    Dim bln中药库房 As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From 部门性质说明 " & _
         " Where 工作性质 Like '中药%' And 部门ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查部门性质]", Val(cbo发药药房.ItemData(cbo发药药房.ListIndex)))
    
    If Not rsData.EOF Then bln中药库房 = True
    
    Set rsData = DeptSendWork_Get剂型(lng药房id)
    
    With Lvw药品剂型
        .ListItems.Clear
        .ListItems.Add , "_" & .ListItems.count + 1, "所有药品剂型", 1, 1
        .ListItems(.ListItems.count).Checked = True
        Do While Not rsData.EOF
            .ListItems.Add , "_" & .ListItems.count + 1, rsData!剂型, 1, 1
            .ListItems(.ListItems.count).Checked = True
            rsData.MoveNext
        Loop
        If bln中药库房 Then
           .ListItems.Add , "_" & .ListItems.count + 1, "0-方剂", 1, 1
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Load医嘱类型()
    '设置医嘱类型
    With Cbo医嘱类型
        .Clear
        .AddItem "0-包含所有单据"
        .AddItem "1-仅含长期医嘱"
        .AddItem "2-仅含临时医嘱"
        .AddItem "3-普通记帐单据"
        .AddItem "4-包含所有医嘱"
'        .ListIndex = Lng医嘱类型
    End With
End Sub
Private Sub ResetParams(Optional ByVal blnNext As Boolean)
    Dim intFixedCol As Integer
    Dim dateCurDate As Date
    Dim i As Integer
    
    BlnSetPara = False
    With Frm部门发药参数设置
        .strPrivs = mstrPrivs
        .blnStartPacker = (TypeName(mobjDrugMAC) = "clsDrugPacker" And mblnStartPacker)
        .Show 1, Me
    End With
    
    If BlnSetPara Then
        '重新取参数
        Call GetParams
        If blnNext = True Then Exit Sub
        
        '重设药房
        If Val(cbo发药药房.Tag) <> mParams.lng药房id Then
            For i = 0 To cbo发药药房.ListCount - 1
                If Val(cbo发药药房.ItemData(i)) = mParams.lng药房id Then
                    cbo发药药房.Tag = cbo发药药房.ItemData(i)
                    cbo发药药房.ListIndex = i
                    Exit For
                End If
            Next
            
            ClearDetailList IIf(tbcDetail.Selected.index = 0, mListType.发药, mListType.退药)
            
            mstrDeptNode = GetDeptStationNode(mParams.lng药房id)
        End If
        
        mfrmDetail.Load核查人 (mParams.lng药房id)
        
        Call Load时间范围
        
        Call SetPacker
        
        '更新子窗口的参数
        mfrmDetail.SetParams
        
        '重新刷新明细
        Call cmdRefreshDept_Click
        Call cmdRefresh_Click
    End If
End Sub

Private Sub SetListItemCheck(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    '列表显示方式，科室显示方式，是否包含拒发
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_ListPopup)
    If Not cbrMenuBar Is Nothing Then
        For Each cbrControl In cbrMenuBar.CommandBar.Controls
            If cbrControl.Id = mconMenu_List_ShowReject And Control.Id = mconMenu_List_ShowReject Then
                cbrControl.Checked = Not cbrControl.Checked
                mParams.blnShowReject = cbrControl.Checked
            ElseIf (cbrControl.Id > mconMenu_ListPopup And cbrControl.Id <= mconMenu_List_ShowOther) _
                And (Control.Id > mconMenu_ListPopup And Control.Id <= mconMenu_List_ShowOther) Then
                cbrControl.Checked = (cbrControl.Id = Control.Id)
                If cbrControl.Id = mconMenu_List_OnlyShowDept Then
                    mParams.blnOnlyShowDept = cbrControl.Checked
                End If
            ElseIf (cbrControl.Id >= mconMenu_List_ShowAll And cbrControl.Id <= mconMenu_List_ShowArea) _
                And (Control.Id >= mconMenu_List_ShowAll And Control.Id <= mconMenu_List_ShowArea) Then
                cbrControl.Checked = (cbrControl.Id = Control.Id)
                mParams.intShowDept = Control.Id - mconMenu_List_ShowAll
            ElseIf cbrControl.Id = mconMenu_List_Sort And Control.Id = mconMenu_List_Sort Then
                cbrControl.Checked = Not cbrControl.Checked
                mParams.blnSort = cbrControl.Checked
            End If
        Next
    End If
End Sub

Private Sub SetPacker()
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Packer, , True)

    If mblnStartPacker = False Then
        If Not cbrMenu Is Nothing Then
            cbrMenu.Visible = False
        End If
        
        '未启动时不显示包装机图标
        Me.stbThis.Panels("PACKER").Visible = False
    Else
        If Not cbrMenu Is Nothing Then
            cbrMenu.Visible = True
        End If
        
        '根据连接状态显示不同的包装机图标
        If mblnPackerConnect = True Then
            If mParams.int暂停传送 = 0 Then
                '正常传送时
                Me.stbThis.Panels("PACKER").Picture = imgPacker.ListImages(1).Picture
            Else
                '暂停传送时
                Me.stbThis.Panels("PACKER").Picture = imgPacker.ListImages(3).Picture
            End If
            
            Me.stbThis.Panels("PACKER").Enabled = True
        Else
            '未连接状态
            Me.stbThis.Panels("PACKER").Picture = imgPacker.ListImages(2).Picture
            Me.stbThis.Panels("PACKER").Enabled = False
        End If
    End If
End Sub

Private Sub ShowMedicalRecord(ByVal intType As Integer)
    '【功能】:查阅当前病人的电子病案
    
    '目前只支持对[发药]列表、[退药]列表的查阅
    If Not (intType = mListType.发药 Or intType = mListType.退药) Then Exit Sub
    
    With mfrmDetail.vsfList(intType)
        '判断当前行是否有效
        If .Row < 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("病人ID")) = "" Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("主页ID")) = "" Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("病人ID")) = .TextMatrix(.Row, .ColIndex("NO")) Then Exit Sub
        
        '调用电子病案查阅接口
        If Not mobjCISJOB Is Nothing Then
            On Error Resume Next
            Call mobjCISJOB.ShowArchive(Me, Val(.TextMatrix(.Row, .ColIndex("病人ID"))), Val(.TextMatrix(.Row, .ColIndex("主页ID"))))
            err.Clear: On Error GoTo 0
        End If
    End With
End Sub

Private Sub ShowOtherConditon()
    picShowOther.Tag = Abs(Val(picShowOther.Tag) - 1)
    picUpOrDown.Picture = imgLvwSel.ListImages(Val(picShowOther.Tag) + 3).Picture
    Call picCondition_Resize
End Sub

Private Sub ShowWindow_ReturnOther()
    TimerAuto.Enabled = False
    
    If Not frm批量退药.ShowEditor(Me, mcondition.lng药房id, False, mParams.int金额保留位数, mstrPrivs) Then
        TimerAuto.Enabled = True
        Exit Sub
    End If
    
    DoEvents
    
    TimerAuto.Enabled = True
End Sub

Private Sub ShowWindow_ReVerify(ByVal strWriteOffMsg As String)
    Dim strWriteOffInfo As String   '销账审核界面返回的上次操作进行审核过的信息：申请时间,病人id|申请时间,病人id...
    
    TimerAuto.Enabled = False
    
    BlnRefresh = False
    
    strWriteOffInfo = Frm药品销账.ShowForm(Me, mcondition.lng药房id, mParams.strUnit, _
        mParams.int金额保留位数, mstrCardType, mParams.int退药清单打印, strWriteOffMsg, _
        mobjSquareCard, mobjPlugIn)
    
    If BlnRefresh Then
        '删除消息记录集中已经审核过的消息记录
        If strWriteOffInfo <> "" Then
            If Not mrsReceiveMsg Is Nothing Then
                If mrsReceiveMsg.RecordCount > 0 Then
                    With mrsReceiveMsg
                        .MoveFirst
                        Do While Not .EOF
                            If InStr(strWriteOffInfo & "|", Format(!申请时间, "yyyy-mm-dd hh:mm:ss") & "," & !病人ID & "|") > 0 Then
                                .Delete adAffectCurrent
                            End If
                            
                            .MoveNext
                        Loop
                    End With
                    '设置消息菜单
                    Call SetMessageBar(mrsReceiveMsg)
                End If
            End If
        End If
        
        cmdRefresh_Click
    End If
    
    DoEvents
    
    TimerAuto.Enabled = True
End Sub

Private Sub ShowWindow_StopFlag()
    Dim frmFlag As New Frm不再发药处方标志
    
    TimerAuto.Enabled = False
    BlnRefresh = True
    
    frmFlag.In_库存检查 = mParams.IntCheckStock
    frmFlag.gstrParentName = "Frm部门发药管理New"
    frmFlag.ShowMe Me, Val(cbo发药药房.ItemData(cbo发药药房.ListIndex))
    
    If BlnRefresh Then
        cmdRefresh_Click
    End If
    
    DoEvents
    TimerAuto.Enabled = True
End Sub



Private Sub UpdateDeptListRecord(ByVal intType As Integer, ByVal Node As Object)
    '根据部门树表的勾选情况更新数据集
    Dim i As Integer

    If mrsDeptList Is Nothing Then Exit Sub
    If mrsDeptList.State = 0 Then Exit Sub
    

    If Mid(Node.Key, 1, 1) = "N" Then
        mrsDeptList.Filter = "NO='" & Split(Node.Tag, "|")(0) & "' and 病人id=" & Val(Split(Node.Tag, "|")(1)) & " and 病人姓名='" & Split(Node.Tag, "|")(2) & "'"
    ElseIf Mid(Node.Key, 1, 1) = "P" Then
        If InStr(1, Node.Tag, "R") > 1 Then
            mrsDeptList.Filter = "病人id=" & Val(Split(Node.Tag, "|")(0)) & " and 领药号='" & Mid(Split(Node.Tag, "|")(1), 2) & "' and 病人姓名='" & Split(Node.Tag, "|")(2) & "'"
        ElseIf InStr(1, Node.Tag, "D") > 1 Then
            mrsDeptList.Filter = "病人id=" & Val(Split(Node.Tag, "|")(0)) & " and 科室ID=" & Val(Mid(Split(Node.Tag, "|")(1), 2)) & " and 病人姓名='" & Split(Node.Tag, "|")(2) & "'"
        Else
            mrsDeptList.Filter = "病人id=" & Val(Split(Node.Tag, "|")(0)) & " and 领药号=0 and 科室ID=" & Val(Split(Node.Tag, "|")(1))
        End If
    ElseIf Mid(Node.Key, 1, 1) = "D" Then
        mrsDeptList.Filter = "科室ID=" & Mid(Node.Key, 3) & ""
    ElseIf Mid(Node.Key, 1, 1) = "R" Then
        mrsDeptList.Filter = "领药号='" & Split(Node.Tag, "|")(0) & "' and 科室ID=" & Val(Split(Node.Tag, "|")(1))
    End If
    
    Do While Not mrsDeptList.EOF
        mrsDeptList!执行状态 = IIf(Node.Checked = True, 1, 0)
        mrsDeptList.Update
        
        mrsDeptList.MoveNext
    Loop
    mrsDeptList.Filter = ""
End Sub

Private Function WarRecoredCount(ByVal lngCount As Long) As Boolean
    Dim intProc As Integer
    
    If mFindWar.blnNoAsk_Rec = True Then
        WarRecoredCount = mFindWar.blnProc_Rec
        Exit Function
    End If
    
    intProc = vbYes
    
    '查询记录数过多时警告
    If mFindWar.blnNoAsk_Rec = False Then
        If lngCount > mParams.lng最大记录数 Then
            intProc = frmMsgBox.ShowMsgBox("查询可能需要很长时间，是否继续？", Me)
            mFindWar.blnNoAsk_Rec = (intProc = vbIgnore Or intProc = vbCancel)
            mFindWar.blnProc_Rec = (intProc = vbYes Or intProc = vbIgnore)
        End If
    End If
    
    WarRecoredCount = mFindWar.blnProc_Rec
End Function

Private Function WarTimeArea() As Boolean
    Dim intDateDiff As Integer
    Dim intProc As Integer
    
    '查询时间间隔
    intDateDiff = DateDiff("d", CDate(mcondition.str开始时间), CDate(mcondition.str结束时间))
    
    '小于查询时间间隔的允许继续操作
    If tbcDetail.Selected.index = mListType.退药 Then
        If intDateDiff <= mParams.int查询退药天数 Then
            WarTimeArea = True
            Exit Function
        End If
    Else
        If intDateDiff <= mParams.int查询发药天数 Then
            WarTimeArea = True
            Exit Function
        End If
    End If
    
    '大于查询时间时，如果上次选择的是不再提示，则按上次选择继续操作
    If tbcDetail.Selected.index = mListType.退药 Then
        If mFindWar.blnNoAsk_Dept_Sended = True Then
            WarTimeArea = mFindWar.blnProc_Dept_Sended
            Exit Function
        End If
    Else
        If mFindWar.blnNoAsk_Dept_Send = True Then
            WarTimeArea = mFindWar.blnProc_Dept_Send
            Exit Function
        End If
    End If
    
    '显示提示
    If tbcDetail.Selected.index = mListType.退药 Then
        If intDateDiff > mParams.int查询退药天数 Then
            intProc = frmMsgBox.ShowMsgBox("查询可能需要很长时间，是否继续？", Me)
                
            mFindWar.blnNoAsk_Dept_Sended = (intProc = vbIgnore Or intProc = vbCancel)
            mFindWar.blnProc_Dept_Sended = (intProc = vbYes Or intProc = vbIgnore)
        End If
        WarTimeArea = mFindWar.blnProc_Dept_Sended
    Else
        If intDateDiff > mParams.int查询发药天数 Then
            intProc = frmMsgBox.ShowMsgBox("查询可能需要很长时间，是否继续？", Me)
                
            mFindWar.blnNoAsk_Dept_Send = (intProc = vbIgnore Or intProc = vbCancel)
            mFindWar.blnProc_Dept_Send = (intProc = vbYes Or intProc = vbIgnore)
        End If
        WarTimeArea = mFindWar.blnProc_Dept_Send
    End If
End Function

Private Sub Cbo发药药房_Click()
    If mblnStart = False Then Exit Sub
    
    With cbo发药药房
        If Val(.Tag) <> Val(.ItemData(.ListIndex)) Then
            Call Load药品剂型(Val(.ItemData(.ListIndex)))
            .Tag = Val(.ItemData(.ListIndex))
            
            mcondition.lng药房id = Val(.Tag)
            
            mfrmDetail.Load核查人 (mcondition.lng药房id)
            
            If Not gobjESign Is Nothing Then
                gblnESign部门发药 = EsignIsOpen(mcondition.lng药房id)
            End If
            
            mstrDeptNode = GetDeptStationNode(Val(.Tag))
            
            zlDatabase.SetPara "发药药房", mcondition.lng药房id, glngSys, 1342
            mblnIs配置中心 = Is配置中心(Val(.Tag))
            
            '库存检查
            mParams.IntCheckStock = MediWork_GetCheckStockRule(Val(.Tag))
                
            '库房单位
            mParams.strUnit = GetSpecUnit(Val(.Tag), gint住院药房)
            
            Call GetDrugDigit(mParams.lng药房id, "药品部门发药", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
            
            '更新子窗口的参数
            mfrmDetail.SetParams
            
            '清除列表
            ClearTreeView IIf(tbcDetail.Selected.index = mListType.退药, 1, 0)
            
            Select Case tbcDetail.Selected.index
                Case mListType.发药, mListType.汇总, mListType.拒发
                    ClearDetailList mListType.发药
                Case mListType.退药
                    ClearDetailList mListType.退药
            End Select
            
            Call SetCommandBar(tbcDetail.Selected.index)
        End If
    End With
End Sub
Private Function Is配置中心(ByVal lng药房id As Long)
    'Is配置中心
    Dim rsSQL As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errHandle
    strTmp = "select 部门id from 部门性质说明 where 部门id=[1] and 工作性质='配制中心'"
    Set rsSQL = zlDatabase.OpenSQLRecord(strTmp, "Is配置中心", lng药房id)
    Is配置中心 = Not (rsSQL.EOF)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbo时间范围_Click()
    With cbo时间范围
        If .ListIndex <> Val(.Tag) Then
            If (Val(.Tag) = 3 And .ListIndex < 3) Or (Val(.Tag) < 3 And .ListIndex = 3) Then
                Call picConMain_Resize
                Call picCondition_Resize
            End If
            .Tag = .ListIndex
        End If
    End With
End Sub


Private Sub Cbo医嘱类型_Click()
    mParams.intAdviceType = Cbo医嘱类型.ListIndex
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    Dim strReturn As String
    
    Select Case Control.Id
        '''''文件
        Case mconMenu_File_PrintSet     '打印设置
            zlPrintSet
        Case mconMenu_File_Preview      '打印预览
            zlSubPrint 2
        Case mconMenu_File_Print        '打印
            zlSubPrint 1
        Case mconMenu_File_Excel        '输出到Excel
            zlSubPrint 3
        Case mconMenu_File_Dept_BillPrintTotal              '打印汇总清单
            Call BillPrint_Total
        Case mconMenu_File_Dept_BillPrintRestore            '打印退药通知单
            Call BillPrint_Restore
        Case mconMenu_File_Dept_BillPrintWait               '打印药品摆药单
            Call BillPrint_Wait
        Case mconMenu_File_Parameter                    '参数设置
            ResetParams
        Case mconMenu_File_Exit                         '退出
            Unload Me
        
        ''''编辑
        Case mconMenu_Edit_Dept_Verify                    '发药
            If Me.tbcDetail.Selected.index = mListType.发药 Or Me.tbcDetail.Selected.index = mListType.汇总 Then
            
                Call DrugStoreWork_Send
            End If
        Case mconMenu_Edit_Dept_Reject                    '拒发确认
            Call DrugStoreWork_Reject
        Case mconMenu_Edit_Dept_RejectRestore             '拒发恢复
            Call DrugStoreWork_RejectRestore
        Case mconMenu_Edit_Dept_Return                    '退药
            Call DrugStoreWork_Return
        
        Case mconMenu_Edit_Dept_ReturnOther               '退其它药房处方
            ShowWindow_ReturnOther
        Case mconMenu_Edit_Dept_ReVerify                  '药品退药销账
            Call ShowWindow_ReVerify("")
        Case mconMenu_Edit_Dept_StopFlag                  '停止发药标记
            ShowWindow_StopFlag
        Case mconMenu_Edit_Dept_VerifySign                    '验证签名
            If gblnESign部门发药 = True Then mfrmDetail.VerifySign
        Case mconMenu_Edit_PlugIn + 1 To mconMenu_Edit_PlugIn + 99 '外挂发药业务功能调用
            DrugSendDeptNormal Control.Parameter
        Case mconMenu_Edit_Dept_CustomCheck                 '自定义审核功能
            Call DrugStoreWork_CustomCheck
        Case mconMenu_Edit_Dept_MedicalRecord               '电子病案查阅
            Call ShowMedicalRecord(tbcDetail.Selected.index)
        
        ''''查看
        Case mconMenu_View_ToolBar_Button               '标准按钮
            Control.Checked = Not Control.Checked
            Me.cbsMain(2).Visible = Control.Checked
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_ToolBar_Text                 '文本标签
            Control.Checked = Not Control.Checked
            For Each cbrControl In Me.cbsMain(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_ToolBar_Size                 '大图标
            Control.Checked = Not Control.Checked
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_StatusBar                    '状态栏
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_FontSize_1, mconMenu_View_FontSize_2, mconMenu_View_FontSize_3                   '字号设置
            mParams.intFont = Val(Control.Parameter)
            Call SetFontSize
            Call zlDatabase.SetPara("字体", mParams.intFont, glngSys, 1342)
        Case mconMenu_View_Find                         '查找
            FindRow
        Case mconMenu_View_FindNext                     '查找下一个
            FindRowNext
        Case mconMenu_View_SelAll                       '全选
            If Not mfrmDetail Is Nothing Then mfrmDetail.SetAllReturn
        Case mconMenu_View_ClsAll                       '全清
            If Not mfrmDetail Is Nothing Then mfrmDetail.SetAllNotReturn
        Case mconMenu_View_Refresh                      '刷新
            cmdRefresh_Click
        
        ''''帮助
        Case mconMenu_Help_Help                         '帮助
            Call ShowHelp(App.ProductName, Me.hWnd, "Frm部门发药管理")
        Case mconMenu_Help_Web                          'WEB上的中联
        Case mconMenu_Help_Web_Home                     '中联主页
            Call zlHomePage(Me.hWnd)
        Case mconMenu_Help_Web_Forum                    '中联论坛
            Call zlWebForum(Me.hWnd)
        Case mconMenu_Help_Web_Mail                     '发送反馈
            Call zlMailTo(Me.hWnd)
        Case mconMenu_Help_About                        '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case mconMenu_List_ShowReject, mconMenu_List_OnlyShowDept, mconMenu_List_ShowOther, mconMenu_List_ShowAll, mconMenu_List_ShowClin, mconMenu_List_ShowTech, mconMenu_List_ShowArea, mconMenu_List_Sort
            '列表和科室显示方式
            Call SetListItemCheck(Control)
        Case mconMenu_Edit_Dept_Packer
            If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
                strReturn = mobjDrugMAC.DrugPackerSet(gcnOracle, mblnPackerConnect)
                mblnPackerConnect = (Left(strReturn, 1) = 1)
                
                '重新设置图标状态
                Call SetPacker
            End If
        Case mconMenu_File_Exit                      '退出
            Unload Me
        
        ''''特殊热键
        Case mconMenu_Edit_Dept_Hot_IC
            If mParams.int输入模式 = mInputType.IC卡 Then
                Call cmdIC_Click
            End If
        Case Else
            If Control.Id > 401 And Control.Id < 499 Then
                '执行自定义报表
                Call BillPrint_Custom(Control)
            End If
            
'            '药房自动发药接口菜单
'            If Control.Id > mconMenu_AutoSend And Control.Id < mconMenu_AutoSend + 10 Then
'                gobjPackerZY.SetInterface Control.Id - mconMenu_AutoSend - 1, mParams.lng药房ID
'            End If
    End Select
    
    ''''药品给药途径分类
    If Control.Id > mconMenu_TypePopup And Control.Id < mconMenu_TypePopup + mTypeCount + 1 Then
        Dim i As Integer
        Dim objPopup As CommandBarControl
        Dim strType As String
        
        Control.Checked = Not Control.Checked
        
        For i = 1 To mTypeCount
            Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_TypePopup + i, , True)
            If Not objPopup Is Nothing Then
                If objPopup.Checked = True Then
                    strType = strType & ";" & objPopup.Caption & ";"
                End If
            End If
        Next
        
        With Lvw给药途径
            For i = 1 To .ListItems.count
                If InStr(1, strType, ";" & .ListItems(i).Tag & ";") > 0 Then
                    .ListItems(i).Checked = True
                Else
                    .ListItems(i).Checked = False
                End If
            Next
        End With
    End If
    
    '病人排序弹出菜单
    If Control.Id > mconMenu_SortPopup And Control.Id < mconMenu_SortPopup + 10 Then
        Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_SortPopup)
        If Not objPopup Is Nothing Then
            For Each cbrControl In objPopup.CommandBar.Controls
                cbrControl.Checked = False
            Next
        End If
        
        Control.Checked = True
        If mParams.int病人排序 <> Control.Id - mconMenu_SortPopup Then
            mParams.int病人排序 = Control.Id - mconMenu_SortPopup
            cmdRefreshDept_Click
        End If
    End If
    
    '消息提醒菜单
    If Control.Id > mconMenu_File_Message And Control.Id < mconMenu_File_Message + 10000 Then
        Call ExecuteWriteOffByMessage(Control)
    End If
End Sub

Private Sub DrugSendDeptNormal(ByVal strFunName As String)
    Dim str当前处方 As String, Int单据 As Integer, strNo As String
    
    If Not mobjPlugIn Is Nothing Then
        str当前处方 = mfrmDetail.GetRecordInfo
        
        If str当前处方 <> "" Then
            Int单据 = Val(Split(str当前处方, "|")(0))
            strNo = Split(str当前处方, "|")(1)
        End If
        
        On Error Resume Next
        Call mobjPlugIn.DrugSendWorkNormal(glngModul, strFunName, mParams.lng药房id, strNo, Int单据)
        err.Clear: On Error GoTo 0
    End If
    
End Sub

Private Sub BillPrint_Custom(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '打印自定义报表
    '默认参数：药品=药品id，药房=药房id，病人ID=病人id，住院号=住院号，NO=处方NO，单据类型=药品收发记录.单据
    
    Dim str当前处方 As String
    Dim Int单据 As Integer, strNo As String
    Dim lng药品id As Long
    Dim strName As String
    
    str当前处方 = mfrmDetail.GetRecordInfo
    
    If str当前处方 <> "" Then
        Int单据 = Val(Split(str当前处方, "|")(0))
        strNo = Split(str当前处方, "|")(1)
        lng药品id = Val(Split(str当前处方, "|")(2))
    End If
    
    strName = Split(Control.Parameter, ",")(1)
    
    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strName, Me, _
        "药品=" & IIf(lng药品id = 0, "", lng药品id), _
        "药房=" & IIf(mcondition.lng药房id = 0, "", mcondition.lng药房id), _
        "病人ID=" & IIf(mcondition.lng病人ID = 0 Or mcondition.lng病人ID = -1, "", mcondition.lng病人ID), _
        "住院号=" & mcondition.str住院号, _
        "NO=" & strNo, _
        "单据类型=" & IIf(Int单据 = 0, "", Int单据))
End Sub
Private Sub zlSubPrint(ByVal bytMode As Byte)
    'bytMode：1-打印；2-预览；3-输出到Excel
    Dim ObjThis As Object
    Dim objPrint As New zlPrint1Grd
    Dim ObjAppRow As New zlTabAppRow
    Dim strTitle As String
    
    '取打印列表对象
    Set ObjThis = mfrmDetail.GetPrintObject(True)
    
    If ObjThis Is Nothing Then
        mfrmDetail.GetPrintObject False
        Exit Sub
    End If
    
    Select Case tbcDetail.Selected.index
        Case mListType.发药
            strTitle = "药品发药清单"
        Case mListType.汇总
            strTitle = "药品汇总发药清单"
        Case mListType.拒发
            strTitle = "药品拒发清单"
        Case mListType.缺药
            strTitle = "药品缺药清单"
        Case mListType.退药
            strTitle = "药品退药清单"
    End Select
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "打印人:" & gstrUserName
    ObjAppRow.Add "打印日期:" & Format(Sys.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add ObjAppRow
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "开始时间:" & Format(Dtp开始时间.Value, "yyyy-MM-dd HH:mm:ss")
    ObjAppRow.Add "结束时间:" & Format(Dtp结束时间.Value, "yyyy-MM-dd HH:mm:ss")
    objPrint.UnderAppRows.Add ObjAppRow
    
    objPrint.Title.Text = strTitle
    Set objPrint.Body = ObjThis
    
    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrView1Grd objPrint, 1
        Case 2
            zlPrintOrView1Grd objPrint, 2
        Case 3
            zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    
    mfrmDetail.GetPrintObject False
End Sub
Private Sub SetFontSize()
    Dim intFont As Integer
    Dim stdfnt As StdFont
    
    Select Case mParams.intFont
        Case 0
            intFont = 9
        Case 1
            intFont = 11
        Case 2
            intFont = 15
        Case Else
            intFont = 9
    End Select
    
    mfrmDetail.SetFontSize intFont
    
    If Not tbcDetail.PaintManager.Font Is Nothing Then
        With tbcDetail
            Set stdfnt = .PaintManager.Font
            stdfnt.Size = intFont
             Set .PaintManager.Font = stdfnt
              .PaintManager.Layout = xtpTabLayoutAutoSize
        End With
    End If
    Me.FontSize = intFont
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    Me.picDetail.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
        
    With fraColorStateSend
        .ZOrder 0
        .Top = stbThis.Top + 90
        .Left = stbThis.Panels("HINT").Left + stbThis.Panels("HINT").Width - .Width - 50
    End With
    
    With fraColorStateReturn
        .ZOrder 0
        .Top = fraColorStateSend.Top
        .Left = stbThis.Panels("HINT").Left + stbThis.Panels("HINT").Width - .Width - 50
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case mconMenu_View_StatusBar '状态栏
            Control.Checked = Me.stbThis.Visible
        Case mconMenu_View_FontSize_1, mconMenu_View_FontSize_2, mconMenu_View_FontSize_3       '字体
            Control.Checked = Val(Control.Parameter) = mParams.intFont
        Case mconMenu_Edit_Dept_MedicalRecord
            If Not (tbcDetail.Selected.index = mListType.发药 Or tbcDetail.Selected.index = mListType.退药) Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
    End Select
End Sub

Private Sub chkAll_Click(index As Integer)
    Dim i As Long
    
    If chkAll(index).Value = 2 Then Exit Sub
    
    mrsDeptList.Filter = ""
    Do While Not mrsDeptList.EOF
        mrsDeptList!执行状态 = chkAll(index).Value
        mrsDeptList.Update
        
        mrsDeptList.MoveNext
    Loop
    
    With tvwList(index)
        For i = 1 To .Nodes.count
            If .Nodes(i).Parent Is Nothing Then
                .Nodes(i).Checked = (chkAll(index).Value = 1)
                TvwCheckNode .Nodes(i), .Nodes(i).Checked
            End If
        Next
    End With
End Sub


Private Sub chkSend_Click(index As Integer)
    Dim objChk As CheckBox
    Dim blnAllUnCheck As Boolean
    
    If mblnStart = False Then Exit Sub
    
    blnAllUnCheck = True
    
    For Each objChk In chkSend
        If objChk.Value = 1 Then
            blnAllUnCheck = False
        End If
    Next
    
    If blnAllUnCheck = True Then
        chkSend(index).Value = 1
    End If
End Sub

Private Sub cmdIC_Click()
    Dim strOutXML As String
    
    If mParams.int输入模式 = mInputType.IC卡 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtInput.Text = mobjICCard.Read_Card()
            If txtInput.Text <> "" Then Call txtInput_KeyPress(vbKeyReturn)
        End If
    Else
        If Not mobjSquareCard Is Nothing Then
            Call mobjSquareCard.zlReadCard(Me, mlngMode, Val(Split(txtInput.Tag, "|")(gCardFormat.卡类别ID)), True, "", txtInput.Text, strOutXML)
            If txtInput.Text <> "" Then Call txtInput_KeyPress(vbKeyReturn)
        End If
    End If
End Sub

Private Sub cmdListSel_Click()
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_ListPopup)
    If Not objPopup Is Nothing Then
        For Each cbrControl In objPopup.CommandBar.Controls
            If Trim(txtInput.Text) = "" Then
                If cbrControl.Id >= mconMenu_List_ShowAll And cbrControl.Id <= mconMenu_List_ShowArea Then
                    cbrControl.Visible = True
                End If
            Else
                If cbrControl.Id >= mconMenu_List_ShowAll And cbrControl.Id <= mconMenu_List_ShowArea Then
                    cbrControl.Visible = False
                End If
            End If
            
            If cbrControl.Id = mconMenu_List_ShowReject Then
                cbrControl.Visible = Not (tbcDetail.Selected.index = mListType.退药)
            End If
        Next
        
        objPopup.CommandBar.ShowPopup
    End If
End Sub


Private Sub cmdRefresh_Click()
    If Val(tvwList(IIf(tbcDetail.Selected.index = 4, mDeptType.退药, mDeptType.发药)).Tag) = 0 Then Exit Sub
    
    GetCondition
    
    mdate上次刷新时间 = Sys.Currentdate
    
    Select Case tbcDetail.Selected.index
        Case mListType.发药, mListType.汇总, mListType.拒发
            ClearDetailList mListType.发药
            
            Call RefreshSendDetail
        Case mListType.退药
            ClearDetailList mListType.退药
            
            Call RefreshReturnDetail
    End Select
End Sub

Private Sub cmdRefreshDept_Click()
    Dim blnExecute As Boolean
    
    If mblnFreshDeptList = True Then Exit Sub
    
    mblnFreshDeptList = True
    
    '刷新时先清除列表
    ClearTreeView IIf(tbcDetail.Selected.index = mListType.退药, 1, 0)
    
    Call GetCondition
    
    If mblnInput = True Then
        blnExecute = True
    ElseIf WarTimeArea = True Then
        blnExecute = True
    End If
    
    If blnExecute Then
        Call AviShow(Me)
    
        Select Case tbcDetail.Selected.index
            Case mListType.发药, mListType.汇总
                Call RefreshSendDept
            Case mListType.退药
                Call RefreshReturnDept
        End Select
        
        Call AviShow(Me, False)
    End If
    
    mblnFreshDeptList = False
End Sub
Private Sub cmd给药途径_Click()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With Lvw给药途径
        .Visible = True
        
        .Top = picCondition.Top + picConOther.Top + txt给药途径.Top + txt给药途径.Height + lngTop
        .Left = picCondition.Left + picConOther.Left + txt给药途径.Left
        .Width = txt给药途径.Width * 3
        .Height = picDeptList.Height + picConOther.Height - txt给药途径.Top - txt给药途径.Height - 50
        
        .SetFocus
        .ZOrder 0
    End With
End Sub

Private Sub cmd药品剂型_Click()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With Lvw药品剂型
        .Visible = True
        
        .Top = picCondition.Top + picConOther.Top + txt药品剂型.Top + txt药品剂型.Height + lngTop
        .Left = picCondition.Left + picConOther.Left + txt药品剂型.Left
        .Width = txt药品剂型.Width * 2
        .Height = picDeptList.Height + picConOther.Height - txt药品剂型.Top - txt药品剂型.Height - 50
        
        .SetFocus
        .ZOrder 0
    End With
End Sub


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1
            Item.Handle = picCondition.hWnd
        Case 2
            Item.Handle = picDeptList.hWnd
        Case 3
'            Item.Handle = tbcDetail.hWnd
            
    End Select
End Sub

Private Sub Form_Activate()
    Call picConMain_Resize
    Call picCondition_Resize
    
    TimerAuto.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Lvw给药途径.Visible = True Then
        If KeyCode = 102 Or KeyCode = 65 Then
            If Shift = vbCtrlMask Then   'Ctrl+A
                Call SelectAllCheck(Lvw给药途径)
            End If
        End If
        
        If KeyCode = 102 Or KeyCode = 82 Then
            If Shift = vbCtrlMask Then   'Ctrl+R
                Call UnSelectAllCheck(Lvw给药途径)
            End If
        End If
    End If
    
    If Lvw药品剂型.Visible = True Then
        If KeyCode = 102 Or KeyCode = 65 Then
            If Shift = vbCtrlMask Then   'Ctrl+A
                Call SelectAllCheck(Lvw药品剂型)
            End If
        End If
        
        If KeyCode = 102 Or KeyCode = 82 Then
            If Shift = vbCtrlMask Then   'Ctrl+R
                Call UnSelectAllCheck(Lvw药品剂型)
            End If
        End If
    End If
    
    '查找
    If tbcDetail.Selected.index = mListType.发药 Or tbcDetail.Selected.index = mListType.退药 Then
        If KeyCode = vbKeyF3 Then
            FindRowNext
        End If
    End If
    
    'Ctrl+F4  读IC卡
    If KeyCode = vbKeyF4 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then
            If mParams.int输入模式 = mInputType.IC卡 Then
                Call cmdIC_Click
            End If
        End If
    End If
End Sub

Private Sub UnSelectAllCheck(ByVal UserListView As ListView)
    Dim n As Integer
    
    For n = 1 To UserListView.ListItems.count
        UserListView.ListItems(n).Checked = False
    Next
End Sub
Private Sub SelectAllCheck(ByVal UserListView As ListView)
    Dim n As Integer
    
    For n = 1 To UserListView.ListItems.count
        UserListView.ListItems(n).Checked = True
    Next
End Sub
Private Sub Form_Load()
    Dim strStart As String
    Dim strPrivs As String
    Dim strMessage As String
    
    mblnStart = False
    mblnEnter = False
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    
    Me.Width = mcstlngWinNormalWidth
    Me.Height = mcstlngWinNormalHeight
    
    On Error Resume Next
    
    'IC卡接口
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hWnd)
    Set mobjICCard.gcnOracle = gcnOracle
    
    '一卡通接口
    mstrCardType = zlfuncCard_Ini(mobjSquareCard, Me, mlngMode)
    
    '初始化界面显示
    mParams.blnShowReject = False
    mParams.int病人排序 = 1
    
    '初始化查询提醒参数
    With mFindWar
        .blnNoAsk_Dept_Send = False
        .blnNoAsk_Dept_Sended = False
        .blnProc_Dept_Send = True
        .blnProc_Dept_Sended = True
        .blnNoAsk_Rec = False
        .blnProc_Rec = True
    End With
    
    '取权限
    Call GetPrivs
    
    '取参数
    Call GetParams
    
    Call SetFontSize
    
    '初始化数据
    mcondition.lng药房id = mParams.lng药房id
    mstrDeptNode = GetDeptStationNode(mParams.lng药房id)
   
    If Load发药药房 = False Then Exit Sub
    '是否进入
    mblnEnter = True
    
    Call Load时间范围
    Call Load取自定义发药类型
    
    Call Load医嘱类型
    
    Call Load给药途径
    Call Load药品剂型(Val(cbo发药药房.ItemData(cbo发药药房.ListIndex)))
    
    Call SetColorState
    
    '------------------------------------------------------------------
    '药品分包机接口
    mblnStartPacker = False
    mblnPackerConnect = False
    
    Set mclsComLib = New zl9ComLib.clsComLib
    
    On Error Resume Next
    
    If Val(zlDatabase.GetPara("启用药品自动化设备接口", glngSys, Val("9010-药品自动化设备接口"))) = 1 Then
        Set mobjDrugMAC = Nothing
        '优先新接口
        Set mobjDrugMAC = CreateObject("zlDrugMachine.clsDrugMachine")
        If err.Number <> 0 Then
            '其次旧接口
            Set mobjDrugMAC = CreateObject("zlDrugPacker.clsDrugPacker")
        End If
    Else
        Set mobjDrugMAC = CreateObject("zlDrugPacker.clsDrugPacker")
    End If
    
    err.Clear: On Error GoTo 0
    
    If TypeName(mobjDrugMAC) = "clsDrugMachine" Then
        '新接口
        ''获取接口的权限
        strPrivs = ";" & zl9ComLib.GetPrivFunc(glngSys, Val("9010-药品自动化设备接口")) & ";"
        If strPrivs Like "*;基本;*" Then
            
            mblnPackerConnect = mobjDrugMAC.Init(1, mclsComLib, strMessage)
        Else
            mblnPackerConnect = False
        End If
    ElseIf TypeName(mobjDrugMAC) = "clsDrugPacker" Then
        '旧接口
        
        '如果存在注册表并且为0表示未启用住院药房接口
        strStart = GetSetting("ZLSOFT", "公共模块\自动发药机", "启用住院药房")
        If Not mobjDrugMAC Is Nothing And strStart <> "0" Then
            mblnStartPacker = True
            If mobjDrugMAC.DBConnect Then
                mblnPackerConnect = True
            Else
                mblnPackerConnect = False
                MsgBox "药品分包机接口数据库未能正常连接，不能传递数据！" & vbCrLf & "提示：你可以在菜单中选择手动重新设置连接。", vbInformation, gstrSysName
            End If
        End If
        
        '药房自动发药机接口相关菜单和状态栏设置
        Call SetPacker
    Else
        mblnPackerConnect = False
    End If
    
    '外挂接口
    Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
    
    '创建电子病案查阅对象
    If mobjCISJOB Is Nothing Then
        On Error Resume Next
        Set mobjCISJOB = CreateObject("zl9CISJob.clsCISJob")
        
        If Not mobjCISJOB Is Nothing Then
            Call mobjCISJOB.InitCISJob(gcnOracle, Me, glngSys, mstrPrivs, gobjBrower.mobjEmr)
        End If
        err.Clear: On Error GoTo 0
    End If
    
    '是否开启自定义审核功能
    If Not mobjPlugIn Is Nothing Then
        On Error Resume Next
        mblnCustomCheck = mobjPlugIn.DrugSendCustomCheckSet(mstrCustomCheckName)
        
        err.Clear: On Error GoTo 0
    End If
    
    '------------------------------------------------------------------
        
    '放到InitComandBars前面，否则有些按钮个性化设置无效
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        '恢复个性化参数
        LoadCustomSet
    End If
    
     '电子签名接口控制
    gblnESign部门发药 = EsignIsOpen(mParams.lng药房id)
    gblnESignUserStoped = False
    If gblnESign部门发药 = True Then
        On Error Resume Next
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        err.Clear: On Error GoTo 0
        If Not gobjESign Is Nothing Then
            If Not gobjESign.Initialize(gcnOracle, glngSys) Then
                Set gobjESign = Nothing
                gblnESign部门发药 = False
            Else
                gblnESign部门发药 = True
                gblnESignUserStoped = gobjESign.CertificateStoped(gstrUserName)
            End If
        Else
            gblnESign部门发药 = False
        End If
    End If
    
    Call Cbo发药药房_Click
    
    '初始化菜单，方格，页面等界面布局
    Call InitComandBars
    Call InitPanes
    Call InitTabControl
    Call InitIDKindNew
    
    '添加自定义报表
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrprivs)
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        '恢复窗口
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
    End If
    
    Call RestoreWinState(Me, App.ProductName)
    Me.picColorStateSend(6).BackColor = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\1345", "未审核医嘱颜色", 33023)
    
    '初始化消息对象
    If mPrives.bln退药销帐 Then
        err = 0
        On Error Resume Next
        Set mobjMipModule = New zl9ComLib.clsMipModule
        Call mobjMipModule.InitMessage(glngSys, mlngMode, mstrPrivs)
        Call AddMipModule(mobjMipModule)
        
        If Not mobjMipModule Is Nothing Then
            Call InitMsgRec
        End If
    End If
    mblnStart = True
End Sub

Private Sub SetSendTypePosition()
    '调整发药类型选择框位置
    Dim n As Integer
    Dim dbl最大宽度 As Double
    Dim dblTmp As Double
    Dim dblSumTmp As Double
    Dim int行数 As Integer
    Dim dblCheckControlH As Double
    Const cst间隔宽度 = 50
    Const cst行距 = 50
    
    picSendType.Visible = mblnExistOtherSendType
    picShowSendType.Visible = mblnExistOtherSendType
    
    If picShowSendType.Visible = False Then Exit Sub
    
    If chkSendType.UBound > 0 Then
        dbl最大宽度 = picSendType.Width - 100
        dblCheckControlH = chkSendType(0).Height
        picSendType.Height = chkSendType(0).Height + 75
        
        int行数 = 0
        dblSumTmp = chkSendType(0).Width + cst间隔宽度
        For n = 1 To chkSendType.UBound
            dblTmp = chkSendType(n).Width + dblSumTmp
            
            If dblTmp <= dbl最大宽度 Then
                chkSendType(n).Top = chkSendType(n - 1).Top
                chkSendType(n).Left = chkSendType(n - 1).Left + chkSendType(n - 1).Width + cst间隔宽度
                dblSumTmp = dblSumTmp + chkSendType(n).Width + cst间隔宽度
            Else
                '换新行，并调整其他控件位置
                int行数 = int行数 + 1
                chkSendType(n).Left = chkSendType(0).Left
                chkSendType(n).Top = chkSendType(0).Top + (dblCheckControlH + cst行距) * int行数
                dblSumTmp = chkSendType(n).Width + cst间隔宽度

                picSendType.Height = chkSendType(n).Top + chkSendType(n).Height + 50
            End If
        Next
    End If
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < mcstlngWinNormalWidth Then Me.Width = mcstlngWinNormalWidth
    If Me.Height < mcstlngWinNormalHeight Then Me.Height = mcstlngWinNormalHeight
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngMyWindow = 0
    mblnFreshDeptList = False
    
    '卸载IC卡刷卡接口
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    
    '卸载药品自动发药机接口
    Set mobjDrugMAC = Nothing
    Set mclsComLib = Nothing
    
    '卸载一卡通接口
    mstrCardType = ""
    Call zlfuncCard_Unload(mobjSquareCard)
    
    '卸载电子病案查阅接口
    Set mobjCISJOB = Nothing
    
    '卸载引用的窗口
    If Not mfrmDetail Is Nothing Then
        Unload mfrmDetail
        Set mfrmDetail = Nothing
    End If
    
    '保存窗口及参数
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
        
        Call SaveWinState(Me, App.ProductName)
    
        '保存个性化设置
        SaveCustomSet
    End If
    
    If mParams.bln保持上一次页签 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品部门发药管理", "当前页签", tbcDetail.Selected.index)
    End If
    
    '卸载消息对象
    If Not mobjMipModule Is Nothing Then
        Call mobjMipModule.CloseMessage
        Call DelMipModule(mobjMipModule)
        Set mobjMipModule = Nothing
    End If

    '卸载外挂接口
    Call zlPlugIn_Unload(mobjPlugIn)
End Sub

Private Sub lblComment_Click()
    ShowOtherConditon
End Sub

Private Sub Lvw给药途径_DblClick()
    ReturnSelected给药途径 0
End Sub

Private Sub Lvw给药途径_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With Lvw给药途径
        For n = 1 To .ListItems.count
            .ListItems(n).Selected = False
        Next
        Item.Selected = True
        If Item.Text = "所有给药途径" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub

Private Sub Lvw给药途径_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ReturnSelected给药途径 1
    End If
End Sub

Private Sub Lvw给药途径_LostFocus()
    Lvw给药途径.Visible = False
End Sub


Private Sub Lvw给药途径_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    
    If Button = 2 Then
        Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_TypePopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub


Private Sub Lvw药品剂型_DblClick()
    ReturnSelected剂型 0
End Sub

Private Sub Lvw药品剂型_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With Lvw药品剂型
        For n = 1 To .ListItems.count
            .ListItems(n).Selected = False
        Next
        Item.Selected = True
        If Item.Text = "所有药品剂型" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub

Private Sub Lvw药品剂型_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ReturnSelected剂型 1
    End If
End Sub

Private Sub Lvw药品剂型_LostFocus()
    Lvw药品剂型.Visible = False
End Sub

Private Sub mobjMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    '接收到消息要验证消息的有效性，如药房等
    '更新消息数据集，并动态增加消息项目下拉菜单
    '消息下拉菜单最多显示5条，当超过5条时，增加一条显示“全部销账”
    Dim objXML As New zl9ComLib.clsXML
    Dim rsMsg As ADODB.Recordset
    Dim blnValid As Boolean
    Dim str科室 As String
    Dim str病人id As String
    Dim str姓名 As String
    Dim str住院号 As String
    Dim str申请时间 As String
    Dim strsql As String
    Dim rstemp As Recordset
    Dim i As Integer
    
'    'ZLHIS_CHARGE_001
'    patient_info 病人信息
'    patient_id 病人id
'    patient_name 姓名
'    identity_card 身份证号
'    in_number 住院号
'    out_number 门诊号
'    cancel_reqeust 销帐申请
'    cancel_charge
'       charge_id 费用id
'       request_kind 申请类别
'       request_time 申请时间
'       request_person 申请人员
'       cancel_item_id 销帐项目id
'       cancel_item_title 销帐项目
'       calcel_num 销帐数量
'       audit_dept_id 审核部门id
'       audit_dept_title 审核部门


    '消息对象为空时退出
    
    If mobjMipModule Is Nothing Then Exit Sub
    
    '消息服务连接失败时不接收消息
    If mobjMipModule.IsConnect = False Then Exit Sub
    
    If objXML Is Nothing Then Exit Sub
    '打开XML文件
    objXML.OpenXMLDocument strMsgContent
    
    '检查消息是否有效，主要是检查药房
    If objXML.GetMultiNodeRecord("cancel_charge", rsMsg) = False Then Exit Sub
    If rsMsg Is Nothing Then Exit Sub
    If rsMsg.RecordCount = 0 Then Exit Sub
    
    blnValid = False
    Do While Not rsMsg.EOF
        If rsMsg("node_name").Value = "audit_dept_id" Then
            If Val(rsMsg("node_value").Value) = mcondition.lng药房id Then
                blnValid = True
                Exit Do
            End If
        End If
        rsMsg.MoveNext
    Loop
    If blnValid = False Then Exit Sub
    
    '如果是有效消息则加入消息数据集
'    str科室 = ""
'    If objXML.GetSingleNodeValue("patient_id", str病人id, xsString) = False Then Exit Sub
'    If objXML.GetSingleNodeValue("patient_name", str姓名, xsString) = False Then Exit Sub
'    If objXML.GetSingleNodeValue("in_number", str住院号, xsString) = False Then Exit Sub
'    If objXML.GetSingleNodeValue("request_time", str申请时间, xsString) = False Then Exit Sub
    
    Call mobjMipModule.ShowMessage(strMsgItemIdentity, "有新的销账申请，请操作员注意在消息列表中查看和处理！", "消息提醒")
    
    '消息有效则从数据库读取消息
    strsql = "select distinct A.申请时间,B.病人id,B.姓名,C.住院号 from 病人费用销帐 A,住院费用记录 B,病案主页 C where A.费用ID=B.ID And B.病人ID=C.病人ID And B.主页id=C.主页id And A.审核部门ID=[1] And A.申请时间>[2] and A.审核人 is null and A.状态=0"
    Set rstemp = zlDatabase.OpenSQLRecord(strsql, "", mcondition.lng药房id, mdateBegin)
    
    Call InitMsgRec
    With mrsReceiveMsg
        For i = 1 To rstemp.RecordCount
            .AddNew
            !科室 = ""
            !病人ID = Val(rstemp!病人ID)
            !姓名 = zlStr.NVL(rstemp!姓名, "")
            !住院号 = zlStr.NVL(rstemp!住院号, "")
            !申请时间 = Format(rstemp!申请时间, "yyyy-MM-dd HH:mm:ss")
            .Update
            
            rstemp.MoveNext
        Next
    End With
    
    '设置消息菜单
    
    Call SetMessageBar(mrsReceiveMsg)
End Sub

Private Sub picColorStateSend_Click(index As Integer)
    On Error GoTo errHandle
    
    If index = 6 Then
        cmdialog.CancelError = True
        cmdialog.ShowColor
        picColorStateSend(6).BackColor = cmdialog.Color
        SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\1345", "未审核医嘱颜色", Val(picColorStateSend(6).BackColor)
        Exit Sub
    End If
errHandle:
End Sub

Private Sub picCondition_Resize()
    On Error Resume Next
    
    With picConMain
        .Width = picCondition.Width
    End With
    
    With picConOther
        If Val(picShowOther.Tag) = 1 Then
            .Visible = True
            .Left = picConMain.Left
            .Top = picConMain.Top + picConMain.Height - 60
            .Width = picCondition.Width
        Else
            .Visible = False
        End If
    End With
    
    With picDeptList
        If Val(picShowOther.Tag) = 1 Then
            .Top = picConOther.Top + picConOther.Height
        Else
            .Top = picConMain.Top + picConMain.Height
        End If
        
        .Left = picConMain.Left
        .Width = picCondition.Width
        .Height = picCondition.Height - .Top - 50
    End With
End Sub

Private Sub picConMain_Resize()
    On Error Resume Next
    
    With cbo发药药房
        .Width = picConMain.Width - .Left - 50
    End With
    
    With fraLineH1
        .Width = picConMain.Width + 150
    End With
    
    With cbo时间范围
        .Left = cbo发药药房.Left
        .Width = cbo发药药房.Width
    End With
    
    If cbo时间范围.ListIndex <> 3 Then
        lblTimeBegin.Visible = False
        Dtp开始时间.Visible = False
        lblTimeEnd.Visible = False
        Dtp结束时间.Visible = False
        
        With txtInput
            .Top = cbo时间范围.Top + cbo时间范围.Height + 60
            .Width = cbo时间范围.Width
        End With
        
        With IDKNType
            .Top = txtInput.Top
        End With
    Else
        lblTimeBegin.Visible = True
        Dtp开始时间.Visible = True
        lblTimeEnd.Visible = True
        Dtp结束时间.Visible = True
        
        With lblTimeBegin
            .Top = lbl时间范围.Top + lbl时间范围.Height + 180
        End With
        
        With Dtp开始时间
            .Top = cbo时间范围.Top + cbo时间范围.Height + 60
            .Width = cbo发药药房.Width
        End With
        
        With lblTimeEnd
            .Top = lblTimeBegin.Top + lblTimeBegin.Height + 180
        End With
        
        With Dtp结束时间
            .Top = Dtp开始时间.Top + Dtp开始时间.Height + 60
            .Width = cbo发药药房.Width
        End With
        
        With txtInput
            .Top = Dtp结束时间.Top + Dtp结束时间.Height + 60
            .Width = cbo发药药房.Width
        End With
        
        With IDKNType
            .Top = txtInput.Top
        End With
    End If
    
    With cmdIC
        .Visible = (IDKNType.GetCurCard.名称 = "IC卡")
        .Top = txtInput.Top
        .Left = picConMain.Width - .Width - 80
        
        If IDKNType.GetCurCard.名称 = "IC卡" Then
            txtInput.Width = .Left - txtInput.Left - 50
        Else
            txtInput.Width = cbo发药药房.Width
        End If
    End With
    
    With chkSend(0)
        .Top = txtInput.Top + txtInput.Height + 60
    End With
    
    With chkSend(1)
        .Top = chkSend(0).Top
    End With
    
    If picConMain.Width > chkSend(1).Left + chkSend(1).Width + chkSend(2).Width + 200 Then
        chkSend(2).Top = chkSend(1).Top
        chkSend(2).Left = chkSend(1).Left + chkSend(1).Width + 100
        lbl发药类型.Top = chkSend(0).Top
    Else
        chkSend(2).Top = chkSend(0).Top + chkSend(0).Height + 50
        chkSend(2).Left = chkSend(0).Left
        lbl发药类型.Top = chkSend(0).Top + 100
    End If
    
    '自定义发药类型的位置
    Call SetSendTypePosition
    If picShowSendType.Visible = True Then
        picShowSendType.Top = chkSend(2).Top + chkSend(2).Height + 100
        picShowSendType.Width = picConMain.Width - 50
        picSendType.Left = picShowSendType.Left + 240
        picSendType.Top = picShowSendType.Top + picShowSendType.Height + 50
        picSendType.Width = picConMain.Width - picSendType.Left
        
        If Val(picShowSendType.Tag) = 1 Then
            picSendType.Visible = True
            picShowOther.Top = picSendType.Top + picSendType.Height + 50
        Else
            picSendType.Visible = False
            picShowOther.Top = picShowSendType.Top + picShowSendType.Height + 50
        End If
    Else
        picShowOther.Top = chkSend(2).Top + chkSend(2).Height + 50
    End If
    
    With picShowOther
        .Left = lbl发药药房.Left
        .Width = picConMain.Width - 50
    End With
    
    With picConMain
        .Height = picShowOther.Top + picShowOther.Height
    End With
    
    With Lvw给药途径
        .Top = picConOther.Top + txt给药途径.Top + txt给药途径.Height
        .Left = picConOther.Left + txt给药途径.Left
        .Width = txt给药途径.Width
        .Height = picDeptList.Height + picConOther.Height - txt给药途径.Top - txt给药途径.Height - 50
    End With
    
    With Lvw药品剂型
        .Top = picConOther.Top + txt药品剂型.Top + txt药品剂型.Height
        .Left = picConOther.Left + txt药品剂型.Left
        .Width = txt药品剂型.Width
        .Height = picDeptList.Height + picConOther.Height - txt药品剂型.Top - txt药品剂型.Height - 50
    End With
End Sub



Private Sub picConOther_Resize()
    On Error Resume Next
    
    With fraLineH2
        .Width = picConOther.Width + 150
    End With
    
    With Cbo医嘱类型
        .Width = picConOther.Width - .Left - 50
    End With
    
    With cmd给药途径
        .Left = picConOther.Width - .Width - 50
        If .Left < txt给药途径.Left + 100 Then .Left = txt给药途径.Left + 100
    End With
    
    With txt给药途径
        .Width = cmd给药途径.Left - .Left + cmd给药途径.Width
    End With
    
    With cmd药品剂型
        .Left = picConOther.Width - .Width - 50
        If .Left < txt药品剂型.Left + 100 Then .Left = txt药品剂型.Left + 100
    End With
    
    With txt药品剂型
        .Width = cmd药品剂型.Left - .Left + cmd药品剂型.Width
    End With
    
    With picConOther
        .Height = chkDangerType(0).Top + chkDangerType(0).Height
    End With
End Sub
Private Sub picDeptList_Resize()
    On Error Resume Next
    
    With fraLineH3
        .Width = picDeptList.Width + 150
    End With
    
    With cmdRefresh
        .Left = picDeptList.Width - .Width - 100
    End With
    
    With cmdRefreshDept
        .Left = cmdRefresh.Left - .Width - 50
    End With
    
    With tvwList(mDeptType.发药)
        .Top = cmdRefreshDept.Top + cmdRefreshDept.Height + 50
        .Left = 0
        .Width = picDeptList.Width - 100
        .Height = picDeptList.Height - .Top - 50
    End With
    
    With tvwList(mDeptType.退药)
        .Top = tvwList(mDeptType.发药).Top
        .Left = 0
        .Width = tvwList(mDeptType.发药).Width
        .Height = tvwList(mDeptType.发药).Height
    End With
End Sub
Private Sub picDetail_Resize()
    On Error Resume Next
    
    With fraLineV1
'        .Top = 0
        .Left = 0
        .Height = picDetail.Height + 100
    End With
    
    With tbcDetail
        .Top = 0
        .Left = fraLineV1.Left + 50
        .Width = picDetail.Width - fraLineV1.Width
        .Height = picDetail.Height - 50
    End With
End Sub

Private Sub picShowOther_Click()
    ShowOtherConditon
End Sub


Private Sub picShowOther_Resize()
    With picUpOrDown
        .Left = picShowOther.Width - .Width
        .Top = 0
    End With
End Sub


Private Sub picShowSendType_Click()
    picShowSendType.Tag = Abs(Val(picShowSendType.Tag) - 1)
    picUpOrDown1.Picture = imgLvwSel.ListImages(Val(picShowSendType.Tag) + 3).Picture
    
    picSendType.Visible = (Val(picShowSendType.Tag) = 1)
    Call picConMain_Resize
    Call picCondition_Resize
End Sub

Private Sub picShowSendType_Resize()
    With picUpOrDown1
        .Left = picShowSendType.Width - .Width
        .Top = 0
    End With
End Sub


Private Sub picUpOrDown_Click()
    ShowOtherConditon
End Sub

Private Sub picUpOrDown1_Click()
    picShowSendType_Click
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.index = 3 Then
        Call ShowWindow_ReVerify("")
    End If
End Sub
Private Sub tbcDetail_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    '切换未发药（包括未发药清单、汇总清单、缺药清单、拒发药清单）和已发药清单（退药清单）
    Dim cbrControl As CommandBarControl
    
    Call mfrmDetail.ShowList(Item.index, Val(cbo发药药房.ItemData(cbo发药药房.ListIndex)))
    Call SetCommandBar(Item.index)
    
    Select Case Item.index
        Case mListType.发药, mListType.汇总, mListType.拒发, mListType.缺药
            Me.dkpMain.FindPane(mconPane_Dept_Condition).Title = "过滤条件(发药模式)"
            
            tvwList(mDeptType.发药).Visible = True
            chkAll(mDeptType.发药).Visible = True
            
            tvwList(mDeptType.退药).Visible = False
            chkAll(mDeptType.退药).Visible = False
            
            chkWithNotAudited.Enabled = True
        Case mListType.退药
            Me.dkpMain.FindPane(mconPane_Dept_Condition).Title = "过滤条件(退药模式)"
            
            tvwList(mDeptType.退药).Visible = True
            chkAll(mDeptType.退药).Visible = True
            
            tvwList(mDeptType.发药).Visible = False
            chkAll(mDeptType.发药).Visible = False
            
            chkWithNotAudited.Enabled = False
    End Select
    
    fraColorStateSend.Visible = (Item.index = mListType.发药)
    fraColorStateReturn.Visible = (Item.index = mListType.退药)

    txtInput.Text = ""
End Sub

Private Sub TimerReturn_Timer()
    Dim strsql As String
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    strsql = "select count(费用id) 数量 from (Select distinct A.费用id,A.申请时间 " & vbNewLine & _
        "From 病人费用销帐 A, 药品收发记录 B" & vbNewLine & _
        "Where A.费用id = B.费用id And Not Exists" & vbNewLine & _
        " (Select 1 From 输液配药内容 C Where C.收发id = B.ID) And 审核部门id = [1] And 申请时间 Between Trunc(Sysdate) And" & vbNewLine & _
        "      Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And 审核时间 Is Null and (B.记录状态=1 or mod(B.记录状态,3)=0))"

    Set rstemp = zlDatabase.OpenSQLRecord(strsql, "", mParams.lng药房id)
    
    Me.stbThis.Panels("CHARGEOFF").Text = "未处理的销帐数据" & rstemp!数量 & "条"
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub DrugStoreWork_Send()
    '药房工作：发药
    Dim rsSendData As ADODB.Recordset
    Dim StrCurDate As String
    
    On Error GoTo errHandle
    
    mblnCheck = False
    
    '取发药数据集
    Set rsSendData = mfrmDetail.GetSendRecord
    
    If rsSendData Is Nothing Then Exit Sub
    
    If rsSendData.RecordCount = 0 Then Exit Sub
    
    If MsgBox("你确定要发药吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '启用电子签名时检查用户是否注册
    If gblnESign部门发药 = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Sub
        End If
    End If
    
    '发药检查
    If DrugStoreWork_CheckSend(rsSendData) = False Then Exit Sub
    
    '执行预调价
    Call setNOtExcetePrice
    
    '取系统时间
    StrCurDate = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    '取汇总发药号
    mcur汇总发药号 = Val(zlDatabase.GetNextNo(20))
    
    '发药处理
    If DrugStoreWork_SendProc(rsSendData, StrCurDate) = False Then Exit Sub
        
    '留存处理
    If DrugStoreWork_StayProc(StrCurDate) = False Then Exit Sub
    
    '销帐处理
    If DrugStoreWork_CancelVerifyProc(StrCurDate) = False Then Exit Sub
    
    '向药品分包机传递数据
    Call DrugStoreWork_SendToPacker(rsSendData)

    '打印汇总单据
    Call DrugStoreWork_PrintBill
    
    '发药后更新部门列表和明细界面
    cmdRefreshDept_Click
    
    mfrmDetail.AfterSendRefresh
    
    If mcur汇总发药号 > 0 Then
        stbThis.Panels("HINT").Text = "上次发药号：" & mcur汇总发药号 & ""
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckGroupSend(ByVal rsGroupRec As ADODB.Recordset, ByVal lng相关ID As Long, ByVal strNo As String) As Boolean
    '检查同组药品是否能够发送
    '前提是药房具有配制中心属性
    '同组药品，只有当所有都是发药状态（其它包括缺药、拒发、不处理）才能发药
    Dim i As Integer
    
    '默认是允许发
    CheckGroupSend = True
    
    '不是配制中心则无该规则
    If mParams.bln配制中心 = False Then Exit Function
    
    '无分组的不管
    If lng相关ID = 0 Then Exit Function
    
    '根据传入的NO，相关ID号判断是否该组药品都能发药
    With rsGroupRec
        .Filter = "NO='" & strNo & "'" & " And 相关ID = " & lng相关ID
        
        If .EOF Then Exit Function
        
        Do While Not .EOF
            '只要存在执行状态不为1，就不能发药；如果高危药品发药方式选择了高危药品种类，那么可以不包括高危药品
            If !执行状态 <> 1 And InStr(1, mParams.str高危发放, !高危药品) = 0 And Not mblnCheck Then
                If MsgBox("同组药品的发药状态不一致，是否继续发药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mblnCheck = True
                    CheckGroupSend = True
                Else
                    mblnCheck = False
                    CheckGroupSend = False
                End If
                Exit Function
            End If
            
            .MoveNext
        Loop
    End With
End Function

Private Function CheckCorrelation(ByVal intType As Integer, ByVal rsSendData As ADODB.Recordset) As Boolean
    'intType:0-发药;1-退药
    '检查处方是否已结帐、检查该病人是否已出院，并对权限进行检查
    Dim strNo As String, lng单据 As Long, str序号 As String, lng病人ID As Long, lng主页ID As Long, str姓名 As String
    
    With rsSendData
        .Filter = "执行状态=" & IIf(intType = 0, mState.发药, mState.退药)
        
        Do While Not .EOF
            strNo = !NO & !单据
            lng单据 = !单据
            strNo = !NO
            lng病人ID = !病人ID
            lng主页ID = !主页id
            str姓名 = !姓名
            str序号 = zlStr.NVL(!费用序号)
            If Not IsReceiptBalance_Charge(intType, mstrPrivs, lng单据, strNo, str序号, 2, 2) Then Exit Function
            If Not IsOutPatient(mstrPrivs, lng单据, strNo, 2, 2, lng病人ID, lng主页ID, 0, str姓名) Then Exit Function
            .MoveNext
        Loop
    End With
    
    CheckCorrelation = True
End Function
Private Sub DrugStoreWork_Reject()
    '药房工作：拒发药
    Dim rsSendData As ADODB.Recordset
    Dim blnBeginTrans As Boolean
    Dim arrSql As Variant
    Dim lngRow As Long
    
    On Error GoTo errHandle
    
    '取发药数据集
    Set rsSendData = mfrmDetail.GetSendRecord
    arrSql = Array()
    
    With rsSendData
        .Filter = "执行状态=" & mState.拒发
        .Sort = "药品ID Asc"
        
        If .EOF Then Exit Sub
        
        Do While Not .EOF
            '检查单据状态
            If DeptSendWork_CheckBill(0, !收发ID, mParams.bln允许未审核处方发药) > 0 Then Exit Sub
            
            .MoveNext
        Loop
        
        .MoveFirst
        
        
        
        Do While Not .EOF
            gstrSQL = "zl_药品收发记录_部门拒发(" & !收发ID & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            .MoveNext
        Loop
    End With
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    For lngRow = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(lngRow)), Me.Caption & "-设置拒发药品")
    Next
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    mfrmDetail.AfterRejectRefresh
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DrugStoreWork_RejectRestore()
    '药房工作：拒发药恢复
    Dim rsSendData As ADODB.Recordset
    Dim blnBeginTrans As Boolean
    Dim lngRow As Long
    Dim arrSql As Variant
    
    On Error GoTo errHandle
    
    '取发药数据集
    Set rsSendData = mfrmDetail.GetSendRecord
    arrSql = Array()
    
    With rsSendData
        .Filter = "执行状态=" & mState.拒发_恢复
        .Sort = "药品ID Asc"

        Do While Not .EOF
            gstrSQL = "zl_药品收发记录_部门恢复(" & !收发ID & ")"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            .MoveNext
        Loop
        
        gcnOracle.BeginTrans
        blnBeginTrans = True
        For lngRow = 0 To UBound(arrSql)
            Call zlDatabase.ExecuteProcedure(CStr(arrSql(lngRow)), Me.Caption & "-恢复拒发药品")
        Next
        gcnOracle.CommitTrans
        blnBeginTrans = False
    End With
    
    mfrmDetail.AfterRejectRestoreRefresh
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub DrugStoreWork_Return()
    '药房工作：退药
    Dim rsReturnData As ADODB.Recordset
    Dim blnBeginTrans As Boolean
    Dim str退药人 As String
    Dim dbl退药数 As Double
    Dim str价格失效提示 As String
    Dim str药品id As String
    Dim StrDate As String
    Dim bln是否有退药 As Boolean
    Dim blnIsReturn As Boolean
    Dim arrSql As Variant
    Dim i As Integer
    Dim str签名记录 As String
    Dim strsql As String
    Dim rstemp As Recordset
    Dim Int退药 As Integer
    Dim strReturnInfo As String
    Dim strReserve As String
    Dim blnCheck As Boolean           '用于优化电子签名的重复检查数据。False-需要重复；True-不重复
    
    On Error GoTo errHandle
    
    '启用电子签名时检查用户是否注册
    If gblnESign部门发药 = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Sub
        End If
    End If
    
    arrSql = Array()
    
    '取发药数据集
    Set rsReturnData = mfrmDetail.GetReturnRecord
    
    If rsReturnData Is Nothing Then Exit Sub
    
    If MsgBox("你确定要退药吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
     '检查USB-KEY
    If gblnESign部门发药 = True And gblnESignUserStoped = False Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            MsgBox "请检查用于电子签名的KEY盘是否插好！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '检查处方是否已结帐、检查该病人是否已出院，并对权限进行检查
    If Not CheckCorrelation(1, rsReturnData) Then Exit Sub
    
    '检查医嘱
    If CheckAdvice(rsReturnData) = False Then Exit Sub
    
    '退药人签名
    str退药人 = ""
    If mParams.bln退药人签名 = True Then
        str退药人 = zlDatabase.UserIdentify(Me, "退药人签名", glngSys, 1342, "退药")
        If str退药人 = "" Then
            Exit Sub
        End If
    End If
    
    '如果原来不分批而现在分批
    '如果批号或效期为空，则提取供用户输入（在子窗口中完成校验和输入）
    
    '单据状态检查
    With rsReturnData
        .Filter = "执行状态=" & mState.退药
        .Sort = "收发ID"
        Do While Not .EOF
            '检查单据状态
            If DeptSendWork_CheckBill(2, !收发ID, mParams.bln允许未审核处方发药) > 0 Then Exit Sub
            
            .MoveNext
        Loop
    End With
    
    '执行预调价
    Call setNOtExcetePrice
    
    '退药
    With rsReturnData
        .Filter = "执行状态=" & mState.退药
        .Sort = "药品ID Asc"
        
        StrDate = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
        
        Do While Not .EOF
            If Val(!退药数) = Val(!准退数) Then
                dbl退药数 = Val(!实际数量)
            Else
                dbl退药数 = Val(!退药数) * Val(!包装)
            End If
            
            If dbl退药数 <> 0 Then
                blnIsReturn = False
                If CheckPrice(!收发ID, str价格失效提示) = False Then
                    If MsgBox("药品[" & !品名 & "(" & !规格 & ")]" & str价格失效提示, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        blnIsReturn = True
                    End If
                Else
                    blnIsReturn = True
                End If
                
                If blnIsReturn = True Then
                    gstrSQL = "zl_药品收发记录_部门退药("
                    '收发ID
                    gstrSQL = gstrSQL & !收发ID
                    '审核人
                    gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                    '审核时间
                    gstrSQL = gstrSQL & ",To_Date('" & StrDate & "','yyyy-MM-dd hh24:mi:ss')"
                    '批号
                    gstrSQL = gstrSQL & "," & IIf(IsNull(!批号), "NULL", IIf(Mid(!批号, 1, 1) = "(", "NULL", "'" & Mid(!批号, 1, 8) & "'"))
                    '效期
                    gstrSQL = gstrSQL & "," & IIf(IsNull(!效期), "NULL", IIf(!效期 = "", "NULL", "To_Date('" & Format(!效期, "yyyy-MM-dd") & "','yyyy-MM-dd')"))
                    '产地
                    gstrSQL = gstrSQL & "," & IIf(IsNull(!产地), "NULL", "'" & !产地 & "'")
                    '退药数
                    gstrSQL = gstrSQL & "," & dbl退药数
                    '退药库房
                    gstrSQL = gstrSQL & ",NULL"
                    '退药人
                    gstrSQL = gstrSQL & ",'" & str退药人 & "'"
                    '金额保留位数
                    gstrSQL = gstrSQL & "," & mParams.int金额保留位数
                    '门诊
                    gstrSQL = gstrSQL & ",2"
                    '汇总发药号
                    gstrSQL = gstrSQL & ",Null"
                    gstrSQL = gstrSQL & ")"
    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                    
                    bln是否有退药 = True
                    
                    If InStr("," & str药品id & ",", "," & !药品ID & ",") = 0 Then
                        str药品id = IIf(str药品id = "", "", str药品id & ",") & !药品ID
                    End If
                    
                    strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & Val(!收发ID) & "," & dbl退药数
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    '提示停用药品
    If str药品id <> "" Then
        Int退药 = 1
        Call CheckStopMedi(str药品id, Int退药)
        If Int退药 = 2 Then Exit Sub
    End If
    
    '集中处理退药事务
    gcnOracle.BeginTrans
    blnBeginTrans = True
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption & "-药品退药")
    Next
    
    '进行签名处理
    If UBound(arrSql) >= 0 And gblnESign部门发药 = True And gblnESignUserStoped = False Then
        With rsReturnData
            .Filter = "执行状态=" & mState.退药
            
            '必须按病人ID，药品ID排序
            .Sort = "单据 Asc ,NO Asc"
            Do While Not .EOF
                str签名记录 = ""
                strsql = "Select id From 药品收发记录 Where mod(记录状态,3)=2 and no=[1] And 单据=[2] And 库房id=[3] and 审核日期=[4]"
                Set rstemp = zlDatabase.OpenSQLRecord(strsql, "", !NO, !单据, mcondition.lng药房id, CDate(StrDate))
                
                If GetSignatureRecored(EsignTache.returnStep, !单据, !NO, mcondition.lng药房id, str签名记录, rstemp!Id, , , , blnCheck) = False Then
                    If blnBeginTrans = True Then gcnOracle.RollbackTrans
                    Exit Sub
                End If
                
                blnCheck = True
                
                If str签名记录 <> "" Then
                    strsql = "Zl_药品签名记录_Insert(" & str签名记录 & ")"
                    
                    Call zlDatabase.ExecuteProcedure(strsql, Me.Caption & "-电子签名")
                Else
                    gcnOracle.RollbackTrans
                    MsgBox "对退药人电子签名失败！", vbInformation, gstrSysName
                    Exit Sub
                End If
                .MoveNext
            Loop
        End With
    End If
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    '打印报表
    If bln是否有退药 = True Then
        If mParams.int退药清单打印 = 2 Then
            If MsgBox("你需要打印退药清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "退药时间=" & StrDate, "包装系数=" & IIf(mParams.strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "发药库房=" & mcondition.lng药房id, 2)
            End If
        ElseIf mParams.int退药清单打印 = 1 Then
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "退药时间=" & StrDate, "包装系数=" & IIf(mParams.strUnit = "门诊单位", "C.门诊包装", "C.住院包装"), "发药库房=" & mcondition.lng药房id, 2)
        End If
        
    Else
        MsgBox "本次没有退药。"
        Exit Sub
    End If
    
    '调用退药后的外挂接口
    If Not mobjPlugIn Is Nothing And bln是否有退药 Then
        On Error Resume Next
        mobjPlugIn.DrugReturnByID mcondition.lng药房id, strReturnInfo, CDate(StrDate), strReserve
        
        err.Clear: On Error GoTo 0
    End If
    
    '发药后更新部门列表和明细界面
    cmdRefreshDept_Click
    
    mfrmDetail.AfterReturnRefresh
    
    Exit Sub
errHandle:
    '如果已开启事务，并且未提交，则出错时回滚事务
    If blnBeginTrans Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub tvwList_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    
    If Button = 2 Then
        If Mid(tvwList(index).SelectedItem.Key, 1, 2) = "P_" Then
            Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_SortPopup)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
End Sub

Private Sub tvwList_NodeCheck(index As Integer, ByVal Node As MSComctlLib.Node)
    Dim i As Long
    Dim blnAllChecked As Boolean
    Dim blnAllUnChecked As Boolean
     
    Call TvwCheckNode(Node, Node.Checked, True)
    Call TvwSetParentNode(tvwList(index), Node, Node.Checked)
    
    blnAllChecked = True
    blnAllUnChecked = True
    
    With tvwList(index)
        For i = 1 To .Nodes.count
            If .Nodes(i).Checked = True Then
                blnAllUnChecked = False
            Else
                blnAllChecked = False
            End If
        Next
    End With
    
    If blnAllChecked = True Then
        chkAll(index).Value = 1
    ElseIf blnAllUnChecked = True Then
        chkAll(index).Value = 0
    Else
        chkAll(index).Value = 2
    End If
    
    Call UpdateDeptListRecord(index, Node)
End Sub


Private Sub txtInput_Change()
    If txtInput.Text <> "" And Len(txtInput.Text) = 18 And Not mobjSquareCard Is Nothing And IDKNType.GetCurCard.名称 = "二代身份证" Then
        Call TxtInput_Validate(False)
    End If

    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtInput.Text = "" And Me.ActiveControl Is txtInput)
End Sub

Private Sub txtInput_GotFocus()
    Call SelAll(txtInput)

    If Not mobjICCard Is Nothing And txtInput.Text = "" Then
        Call mobjICCard.SetEnabled(True)
    End If
End Sub


Private Sub TxtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
         Call TxtInput_Validate(True)
    End If
End Sub


Private Sub txtInput_KeyPress(KeyAscii As Integer)
    mblnCard = False
    
    If mParams.int输入模式 = mInputType.住院号 Or mParams.int输入模式 = mInputType.病人ID Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyEscape Or KeyAscii = vbKeyBack Then Exit Sub
        KeyAscii = 0
    ElseIf mParams.int输入模式 = mInputType.姓名 Then
        '姓名类别
        mblnCard = zlCommFun.InputIsCard(txtInput, KeyAscii, False)
    End If
    
    If mParams.int输入模式 > 9 Then
        '其他的是消费卡
        If InStr(":：;；?？''||" & Chr(22) & Chr(32), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        
        If Len(txtInput.Text) = txtInput.MaxLength - 1 And KeyAscii <> 8 Then
            txtInput.Text = txtInput.Text & Chr(KeyAscii)
            txtInput.SelStart = Len(txtInput.Text)
            KeyAscii = 0
        End If
        
'        mblnCard = (KeyAscii <> 8 And Len(txtInput.Text) = txtInput.MaxLength - 1 And txtInput.SelLength <> Len(txtInput.Text))
    End If
End Sub

Private Sub txtInput_LostFocus()
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
End Sub

Private Sub TxtInput_Validate(Cancel As Boolean)
    Dim strDeptInfo As String
    Dim strInput As String
    Dim strInput缓存 As String

    '取病人名称，病人当前病区，并提取处方记录
    '当取到病人信息后，返回输入框格式：输入信息-病人姓名
    '输入科室信息，返回科室ID，科室名称
    
    mblnInput = False
    
    If InStr(Trim(txtInput.Text), "-") > 0 Then
        '取“-”前面的输入信息
        strInput = Mid(Trim(txtInput.Text), 1, InStr(Trim(txtInput.Text), "-") - 1)
    Else
        strInput = Trim(txtInput.Text)
    End If
    
    If strInput = "" Then Exit Sub
    
    strInput缓存 = strInput
    
    If mParams.int输入模式 = mInputType.NO Then
        If IsNumeric(strInput) Then
            strInput = GetFullNO(strInput, 14)
        End If
    End If
    
    strDeptInfo = GetPatiInfo(mParams.int输入模式, strInput)
    
    txtInput.Tag = ""
    If strDeptInfo <> "" Then
        Select Case mParams.int输入模式
        Case mInputType.姓名
'            If mblnCard = True Then
'                txtInput.Text = UCase(strInput)
'                txtInput.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
'            Else
'                txtInput.Text = Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
'                txtInput.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
'            End If
            txtInput.Text = Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
            txtInput.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
            txtInput.PasswordChar = ""
'        Case mInputType.就诊卡
'            txtInput.PasswordChar = ""
'            txtInput.MaxLength = 0
'            txtInput.Text = Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
'            txtInput.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
        Case mInputType.领药号, mInputType.发药号
            txtInput.Text = strDeptInfo
        Case mInputType.领药部门
            '返回部门ID，部门名称
            txtInput.Text = Split(strDeptInfo, ",")(1)
            txtInput.Tag = Split(strDeptInfo, ",")(0)
        Case mInputType.IC卡
            txtInput.Text = Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
            txtInput.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
        Case mInputType.床号
            '床号实际通过病人ID来查询，界面显示床号-病人信息，Tag记录病人ID
            txtInput.Text = strInput & "-" & Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
            txtInput.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
        Case Else
            If mParams.int输入模式 > 9 Then
                '其他消费卡，返回病人ID
                txtInput.Text = Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
                txtInput.Tag = Mid(Split(strDeptInfo, "|")(1), 1, InStr(Split(strDeptInfo, "|")(1), ",") - 1)
                If IDKNType.GetCurCard.名称 = "就诊卡" Then
                    txtInput.PasswordChar = ""
                End If
            Else
                txtInput.Text = strInput & "-" & Mid(Split(strDeptInfo, "|")(1), InStr(Split(strDeptInfo, "|")(1), ",") + 1)
            End If
        End Select
    Else
        txtInput.Tag = 0
    End If
    
    If mParams.int输入模式 <> mInputType.领药部门 Then
        mblnInput = True
    End If
        
    '刷新部门列表
    DoEvents
    cmdRefreshDept_Click
    
    '自动设置为全选，并提取明细记录
    If chkAll(IIf(tbcDetail.Selected.index <> mListType.退药, 0, 1)).Enabled = True Then
        chkAll(IIf(tbcDetail.Selected.index <> mListType.退药, 0, 1)).Value = 1
        Call chkAll_Click(IIf(tbcDetail.Selected.index <> mListType.退药, 0, 1))

        DoEvents
        Call cmdRefresh_Click
    End If
    
    tbcDetail.SetFocus
    
    mblnInput = False
    
End Sub

Private Function GetPatiInfo(ByVal intType As Integer, ByVal strInfo As String) As String
    'intType：mInputType的项目值
    '返回病人信息：当前病区（ID和部门名称），病人信息（ID和姓名）
    '格式：13,一病区|1,张三
    Dim rstemp As ADODB.Recordset
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim lngH As Long
    Dim blnCancel As Boolean
    Dim lng病人ID As Long
    
    On Error GoTo errHandle
    If intType = mInputType.住院号 Then
        If Not IsNumeric(strInfo) Then Exit Function
        
        gstrSQL = "Select Nvl(A.当前病区id, A.出院科室id) As 科室ID, C.编码 || '-' || C.名称 As 部门名称, B.病人id, B.姓名 As 病人姓名 " & _
            " From 病案主页 A, 病人信息 B, 部门表 C, 病案主页 P " & _
            " Where A.病人id = B.病人id And A.主页id = B.主页id And B.病人id = P.病人id And Nvl(A.当前病区id, A.出院科室id) = C.ID(+) And P.住院号 = [1]"
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人信息", strInfo)
    ElseIf intType = mInputType.病人ID Then
        If Not IsNumeric(strInfo) Then Exit Function
        
        gstrSQL = "Select Nvl(A.当前病区id, A.出院科室id) As 科室ID, C.编码 || '-' || C.名称 As 部门名称, B.病人id, B.姓名 As 病人姓名 " & _
            " From 病案主页 A, 病人信息 B, 部门表 C " & _
            " Where A.病人id = B.病人id And A.主页id = B.主页id And Nvl(A.当前病区id, A.出院科室id) = C.ID(+) And A.病人id = [1]"
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人信息", Val(strInfo))
    ElseIf intType = mInputType.NO Then
        gstrSQL = "Select Distinct Nvl(A.病人病区id, 病人科室id) As 科室ID, B.编码 || '-' || B.名称 As 部门名称, A.病人id, A.姓名 As 病人姓名 " & _
            " From 住院费用记录 A, 部门表 B " & _
            " Where Nvl(A.病人病区id, 病人科室id) = B.ID(+) And A.NO = [1] And A.门诊标志=2 And A.执行部门id = [2] "
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人信息", strInfo, mcondition.lng药房id)
    ElseIf intType = mInputType.床号 Then
        '床号可能不唯一，返回列表供选择
        gstrSQL = "Select Rownum As ID, B.姓名 As 病人姓名, C.编码 || '-' || C.名称 As 部门名称, Nvl(A.当前病区id, A.出院科室id) As 科室ID, B.病人id " & _
            " From 病案主页 A, 病人信息 B, 部门表 C " & _
            " Where A.病人id = B.病人id And A.主页id = B.主页id And Nvl(A.当前病区id, A.出院科室id) = C.ID(+) And B.当前床号 = [1]"
            
        vRect = zlControl.GetControlRect(txtInput.hWnd)
        lngH = txtInput.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
        
        Set rstemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "取病人信息", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strInfo)
        If blnCancel = True Then Exit Function
    ElseIf intType = mInputType.姓名 Then
        If mblnCard = True Then
            gstrSQL = "Select /*+rule*/ Nvl(A.当前病区id, A.出院科室id) As 科室ID, C.编码 || '-' || C.名称 As 部门名称, B.病人id, B.姓名 As 病人姓名 " & _
                " From 病案主页 A, 病人信息 B, 部门表 C " & _
                " Where A.病人id = B.病人id And A.主页id = B.主页id And Nvl(A.当前病区id, A.出院科室id) = C.ID(+) And B.就诊卡号 = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人信息", strInfo)
        Else
            '病人名称可能会有重复，返回列表供选择
            gstrSQL = "Select /*+rule*/ Rownum As ID, 病人姓名, 科室ID, 部门名称, 病人id" & _
                " From (Select Distinct B.姓名 As 病人姓名, B.病人id, Nvl(A.当前病区id, A.出院科室id) As 科室ID, C.编码 || '-' || C.名称 As 部门名称 " & _
                " From 病案主页 A, 病人信息 B, 部门表 C " & _
                " Where A.病人id = B.病人id And A.主页id = B.主页id And Nvl(A.当前病区id, A.出院科室id) = C.ID(+) And B.姓名 Like [1])"
            
            vRect = zlControl.GetControlRect(txtInput.hWnd)
            lngH = txtInput.Height
            sngX = vRect.Left - 15
            sngY = vRect.Top
            
            Set rstemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "取病人信息", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strInfo & "%")
            If blnCancel = True Then Exit Function
        End If
    ElseIf intType = mInputType.领药部门 Then
        gstrSQL = " Select ID,编码,名称 From 部门表 " & _
             " Where ID in (Select 部门ID From 部门性质说明 Where 工作性质 In ('临床','检查','检验','治疗','手术','营养','护理') And 服务对象 IN(2,3))" & _
             " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) And (编码 Like [1] Or 简码 Like [1] Or 名称 Like [2])" & _
             " Order By 编码"
        
        vRect = zlControl.GetControlRect(txtInput.hWnd)
        lngH = txtInput.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
        
        Set rstemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "取部门信息", False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, UCase(strInfo) & "%", IIf(gstrMatchMethod = "0", "%", "") & UCase(strInfo) & "%")
        If blnCancel = True Then Exit Function
        
        If rstemp Is Nothing Then Exit Function
        If rstemp.EOF Then Exit Function
        
        GetPatiInfo = rstemp!Id & "," & "[" & rstemp!编码 & "]" & rstemp!名称
        Exit Function
    ElseIf intType = mInputType.IC卡 Then
        If Not mobjSquareCard Is Nothing Then
            '通过卡ID和卡号查找病人ID
            Call mobjSquareCard.zlGetPatiID("IC卡", UCase(txtInput.Text), False, lng病人ID)
        End If
        
        If lng病人ID > 0 Then
            gstrSQL = "Select Nvl(A.当前病区id, A.出院科室id) As 科室ID, C.编码 || '-' || C.名称 As 部门名称, B.病人id, B.姓名 As 病人姓名 " & _
                " From 病案主页 A, 病人信息 B, 部门表 C " & _
                " Where A.病人id = B.病人id And A.主页id = B.主页id And Nvl(A.当前病区id, A.出院科室id) = C.ID(+) And B.病人id = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人信息", lng病人ID)
        End If
    ElseIf intType > 9 Then
        '消费卡
        If Not mobjSquareCard Is Nothing Then
            '通过卡ID和卡号查找病人ID
            Call mobjSquareCard.zlGetPatiID(Split(Split(mstrCardType, ";")(intType - 10), "|")(3), txtInput.Text, True, lng病人ID)
        End If
        
        If lng病人ID > 0 Then
            gstrSQL = "Select Nvl(A.当前病区id, A.出院科室id) As 科室ID, C.编码 || '-' || C.名称 As 部门名称, B.病人id, B.姓名 As 病人姓名 " & _
                " From 病案主页 A, 病人信息 B, 部门表 C " & _
                " Where A.病人id(+) = B.病人id And A.主页id(+) = B.主页id And Nvl(A.当前病区id, A.出院科室id) = C.ID(+) And b.病人id = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人信息", lng病人ID)
        End If
    Else
        GetPatiInfo = strInfo
        Exit Function
    End If
    
    If rstemp Is Nothing Then Exit Function
    If rstemp.EOF Then Exit Function
    
    GetPatiInfo = rstemp!科室ID & "," & rstemp!部门名称 & "|" & rstemp!病人ID & "," & rstemp!病人姓名
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub LoadDept()
    On Error GoTo errHandle
    gstrSQL = "select A.id,A.名称 from 部门表 A" & IIf(mParams.strSourceDep = "", "", ",Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) B ") & " where A.id=B.Column_Value"
    Set mRsDept = zlDatabase.OpenSQLRecord(gstrSQL, "LoadDept", mParams.strSourceDep)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub










