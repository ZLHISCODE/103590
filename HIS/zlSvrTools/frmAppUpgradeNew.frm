VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAppUpgradeNew 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "系统升迁"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   Icon            =   "frmAppUpgradeNew.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer tmrRefresh 
      Interval        =   2000
      Left            =   720
      Top             =   6600
   End
   Begin VB.Frame fraStep 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   5412
      Index           =   0
      Left            =   0
      TabIndex        =   60
      Top             =   840
      Visible         =   0   'False
      Width           =   11412
      Begin VB.CheckBox chkStat 
         Caption         =   "收集统计信息"
         Height          =   312
         Left            =   840
         TabIndex        =   79
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin VB.Frame fraErrOption 
         BorderStyle     =   0  'None
         Height          =   252
         Left            =   2040
         TabIndex        =   73
         Top             =   4920
         Width           =   3252
         Begin VB.OptionButton optErrOption 
            Caption         =   "忽略所有错误"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   75
            Top             =   0
            Width           =   1452
         End
         Begin VB.OptionButton optErrOption 
            Caption         =   "忽略次要错误"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   0
            Value           =   -1  'True
            Width           =   1452
         End
      End
      Begin VB.Frame fraLogType 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   67
         Top             =   4379
         Width           =   3015
         Begin VB.OptionButton optLogType 
            Caption         =   "全部记录"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   69
            Top             =   60
            Width           =   1095
         End
         Begin VB.OptionButton optLogType 
            Caption         =   "仅记录出错SQL"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   68
            Top             =   60
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.TextBox txtCpu 
         Alignment       =   2  'Center
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   33
         Text            =   "4"
         Top             =   1840
         Width           =   350
      End
      Begin VB.TextBox txtToolsUser 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "ZLTOOLS"
         Top             =   60
         Width           =   1725
      End
      Begin VB.Frame fraImpRpt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2280
         TabIndex        =   65
         Top             =   1463
         Width           =   2655
         Begin VB.OptionButton optRpt 
            Caption         =   "只导入数据源"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   27
            Top             =   60
            Width           =   1455
         End
         Begin VB.OptionButton optRpt 
            Caption         =   "整体导入"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   60
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox txtLogLong 
         Alignment       =   2  'Center
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6405
         MaxLength       =   3
         TabIndex        =   54
         Text            =   "3"
         Top             =   4416
         Width           =   405
      End
      Begin VB.CheckBox chkLogLong 
         Caption         =   "记录执行超过     分钟的SQL语句"
         Height          =   255
         Left            =   5040
         TabIndex        =   53
         Top             =   4440
         Value           =   1  'Checked
         Width           =   3060
      End
      Begin VB.CheckBox ckhIdxOnLine 
         Caption         =   "创建索引采用在线模式"
         Height          =   180
         Left            =   5040
         TabIndex        =   36
         Top             =   1900
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.Frame fraStasCollect 
         Caption         =   "               "
         Height          =   1785
         Left            =   840
         TabIndex        =   61
         Top             =   2205
         Width           =   10092
         Begin VB.Frame fraStatType 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   1560
            TabIndex        =   76
            Top             =   0
            Width           =   3612
            Begin VB.OptionButton optStatType 
               Caption         =   "只记录待收集对象"
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   78
               Top             =   0
               Width           =   1812
            End
            Begin VB.OptionButton optStatType 
               Caption         =   "升级中直接收集"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   77
               Top             =   0
               Value           =   -1  'True
               Width           =   1812
            End
         End
         Begin VB.TextBox txtCpu 
            Alignment       =   2  'Center
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1170
            MaxLength       =   2
            TabIndex        =   40
            Text            =   "4"
            Top             =   600
            Width           =   350
         End
         Begin VB.TextBox txtEsper 
            Alignment       =   2  'Center
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3120
            MaxLength       =   3
            TabIndex        =   45
            Top             =   1200
            Width           =   345
         End
         Begin VB.CheckBox chkAutoPer 
            Caption         =   "采用自动百分比"
            Height          =   255
            Left            =   4080
            TabIndex        =   48
            Top             =   1236
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin MSComCtl2.UpDown udCpu 
            Height          =   300
            Index           =   1
            Left            =   1476
            TabIndex        =   41
            Top             =   600
            Width           =   252
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   4
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCpu(1)"
            BuddyDispid     =   196616
            BuddyIndex      =   1
            OrigLeft        =   3435
            OrigTop         =   3600
            OrigRight       =   3690
            OrigBottom      =   3915
            Max             =   6
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udEsper 
            Height          =   300
            Left            =   3420
            TabIndex        =   46
            Top             =   1200
            Width           =   252
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   4
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtEsper"
            BuddyDispid     =   196626
            OrigLeft        =   3435
            OrigTop         =   3600
            OrigRight       =   3690
            OrigBottom      =   3915
            Max             =   100
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lblCpuWarn 
            AutoSize        =   -1  'True
            Caption         =   "未超过4个CPU，不能并行！"
            ForeColor       =   &H002222B2&
            Height          =   180
            Index           =   1
            Left            =   1800
            TabIndex        =   42
            Top             =   660
            Visible         =   0   'False
            Width           =   2160
         End
         Begin VB.Label lblBigTable 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "大表"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   3060
            TabIndex        =   38
            ToolTipText     =   "ZLBakTables与ZlBigTables中定义的表"
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lblStaCpuName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "并行度="
            Height          =   180
            Left            =   480
            TabIndex        =   39
            Top             =   660
            Width           =   636
         End
         Begin VB.Label lblEsperName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "采样百分比： ESTIMATE_PERCENT="
            Height          =   180
            Left            =   360
            TabIndex        =   44
            Top             =   1260
            Width           =   2700
         End
         Begin VB.Label lblEsperPer 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   180
            Left            =   3720
            TabIndex        =   47
            Top             =   1260
            Width           =   96
         End
         Begin VB.Label lblStaCpuNote 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "如果当前服务器有其他应用程序运行，请根据实际情况预留CPU资源给其他应用程序"
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   612
            TabIndex        =   43
            Top             =   960
            Width           =   6780
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblEsperNote 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "建议使用自动百分比，如果已证实自动百分比不准确，可以根据数据量和硬件配置指定百分比"
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   612
            TabIndex        =   49
            Top             =   1560
            Width           =   7980
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblStaInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "仅对本次升级中表结构发生变化的    进行统计信息收集"
            Height          =   180
            Left            =   360
            TabIndex        =   37
            Top             =   360
            Width           =   4740
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CheckBox chkRpt 
         Caption         =   "导入报表"
         Height          =   180
         Left            =   840
         TabIndex        =   25
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "执行可选过程"
         Height          =   180
         Left            =   840
         TabIndex        =   22
         Top             =   1275
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox txtDBAPwd 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5190
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   390
         Width           =   1725
      End
      Begin VB.TextBox txtDBAUser 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2400
         TabIndex        =   11
         Text            =   "System"
         Top             =   390
         Width           =   1725
      End
      Begin VB.TextBox txtHisPwd 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5190
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   720
         Width           =   1725
      End
      Begin VB.CheckBox chkHisAll 
         Caption         =   "全部升级"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   743
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtToolsPwd 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5190
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   60
         Width           =   1725
      End
      Begin MSComCtl2.UpDown udCpu 
         Height          =   300
         Index           =   0
         Left            =   3420
         TabIndex        =   34
         Top             =   1845
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   4
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCpu(0)"
         BuddyDispid     =   196616
         BuddyIndex      =   0
         OrigLeft        =   3435
         OrigTop         =   3600
         OrigRight       =   3690
         OrigBottom      =   3915
         Max             =   6
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkParallel 
         Caption         =   "采用"
         Height          =   180
         Left            =   840
         TabIndex        =   30
         Top             =   1900
         Value           =   1  'Checked
         Width           =   660
      End
      Begin VB.Label lblErrOption 
         AutoSize        =   -1  'True
         Caption         =   "错误处理方式"
         Height          =   180
         Left            =   840
         TabIndex        =   72
         Top             =   4956
         Width           =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   720
         X2              =   720
         Y1              =   4800
         Y2              =   5285
      End
      Begin VB.Label lblAtuoOption 
         AutoSize        =   -1  'True
         Caption         =   "其他"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   71
         Top             =   4866
         Width           =   216
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLogType 
         AutoSize        =   -1  'True
         Caption         =   "日志记录方式"
         Height          =   180
         Left            =   840
         TabIndex        =   70
         Top             =   4476
         Width           =   1080
      End
      Begin VB.Label lblCpuWarn 
         AutoSize        =   -1  'True
         Caption         =   "未超过4个CPU，不能并行！"
         ForeColor       =   &H002222B2&
         Height          =   180
         Index           =   0
         Left            =   3720
         TabIndex        =   35
         Top             =   1900
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.Label lblParallelNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "并行DDL"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   1500
         TabIndex        =   31
         ToolTipText     =   "并行DDL只对索引、约束的创建有效，可以大幅缩短执行时间。"
         Top             =   1900
         Width           =   630
      End
      Begin VB.Label lblParallel 
         AutoSize        =   -1  'True
         Caption         =   "并行度="
         Height          =   180
         Left            =   2400
         TabIndex        =   32
         Top             =   1900
         Width           =   630
      End
      Begin VB.Label lblLogOption 
         AutoSize        =   -1  'True
         Caption         =   "日志"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   50
         Top             =   4224
         Width           =   216
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblUpOption 
         AutoSize        =   -1  'True
         Caption         =   "升级选项"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   480
         TabIndex        =   21
         Top             =   2275
         Width           =   210
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblValiUser 
         AutoSize        =   -1  'True
         Caption         =   $"frmAppUpgradeNew.frx":058A
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   480
         TabIndex        =   3
         Top             =   180
         Width           =   210
         WordWrap        =   -1  'True
      End
      Begin VB.Line linLogOption 
         BorderColor     =   &H8000000C&
         X1              =   720
         X2              =   720
         Y1              =   4152
         Y2              =   4637
      End
      Begin VB.Line linUpOption 
         BorderColor     =   &H8000000C&
         X1              =   720
         X2              =   720
         Y1              =   1275
         Y2              =   3990
      End
      Begin VB.Line linValiUser 
         BorderColor     =   &H8000000C&
         X1              =   720
         X2              =   720
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Label lblLogModi 
         AutoSize        =   -1  'True
         Caption         =   "修改…"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   6960
         TabIndex        =   52
         Top             =   4152
         Width           =   540
      End
      Begin VB.Label lblLog 
         AutoSize        =   -1  'True
         Caption         =   "升迁日志文件：C:\APPSOFT\Log\安装升迁\150930_00010304062124_1645.log"
         Height          =   180
         Left            =   840
         TabIndex        =   51
         Top             =   4152
         Width           =   6120
      End
      Begin VB.Label lblRptSel 
         AutoSize        =   -1  'True
         Caption         =   "修改…"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   5280
         TabIndex        =   28
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label lblOptSel 
         AutoSize        =   -1  'True
         Caption         =   "修改…"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   5280
         TabIndex        =   23
         Top             =   1275
         Width           =   540
      End
      Begin VB.Label lblOptTotal 
         AutoSize        =   -1  'True
         Caption         =   "总数：8，执行：4"
         ForeColor       =   &H00808080&
         Height          =   180
         Left            =   6000
         TabIndex        =   24
         Top             =   1275
         Width           =   1440
      End
      Begin VB.Label lblRptTotal 
         AutoSize        =   -1  'True
         Caption         =   "总数：8，整体导入：4，只导入数据源：2"
         ForeColor       =   &H00808080&
         Height          =   180
         Left            =   6000
         TabIndex        =   29
         Top             =   1560
         Width           =   3330
      End
      Begin VB.Label lblDBAPwd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密  码"
         Height          =   180
         Left            =   4560
         TabIndex        =   12
         Top             =   450
         Width           =   540
      End
      Begin VB.Label lblDBAUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户名"
         Height          =   180
         Left            =   1800
         TabIndex        =   10
         Top             =   450
         Width           =   540
      End
      Begin VB.Label lblDBA 
         AutoSize        =   -1  'True
         Caption         =   "DBA用户"
         Height          =   180
         Left            =   840
         TabIndex        =   9
         Top             =   450
         Width           =   630
      End
      Begin VB.Label lblToolsPwd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密  码"
         Height          =   180
         Left            =   4560
         TabIndex        =   7
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lblToolsUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户名"
         Height          =   180
         Left            =   1800
         TabIndex        =   5
         Top             =   120
         Width           =   1050
      End
      Begin VB.Label lblTools 
         AutoSize        =   -1  'True
         Caption         =   "管理工具"
         Height          =   180
         Left            =   840
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblHisPwd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密  码"
         Height          =   180
         Left            =   4560
         TabIndex        =   17
         Top             =   780
         Width           =   540
      End
      Begin VB.Label lblHis 
         AutoSize        =   -1  'True
         Caption         =   "历史库"
         Height          =   180
         Left            =   840
         TabIndex        =   14
         Top             =   780
         Width           =   540
      End
      Begin VB.Label lblHisSel 
         AutoSize        =   -1  'True
         Caption         =   "修改…"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   7080
         TabIndex        =   20
         Top             =   780
         Width           =   540
      End
      Begin VB.Label lblHisTotal 
         AutoSize        =   -1  'True
         Caption         =   "总数：8，选择：2"
         ForeColor       =   &H00808080&
         Height          =   180
         Left            =   2880
         TabIndex        =   16
         Top             =   780
         Width           =   1440
      End
      Begin VB.Label lblHisWarn 
         AutoSize        =   -1  'True
         Caption         =   "3个历史库未通过验证！"
         ForeColor       =   &H002222B2&
         Height          =   180
         Left            =   7080
         TabIndex        =   19
         Top             =   780
         Width           =   1890
      End
   End
   Begin VB.Frame fraStep 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5412
      Index           =   1
      Left            =   0
      TabIndex        =   62
      Top             =   840
      Visible         =   0   'False
      Width           =   11412
      Begin VB.TextBox txtSQL 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   5016
         Left            =   3120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
         Top             =   360
         Width           =   8172
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPlan 
         Height          =   5412
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   3060
         _cx             =   5397
         _cy             =   9546
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   16777215
         GridColorFixed  =   16777215
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAppUpgradeNew.frx":059C
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
         OutlineBar      =   5
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
         Begin MSComctlLib.ImageList imgPlan 
            Left            =   2160
            Top             =   0
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
                  Picture         =   "frmAppUpgradeNew.frx":05C6
                  Key             =   "Finish"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAppUpgradeNew.frx":0B60
                  Key             =   "Doing"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAppUpgradeNew.frx":10FA
                  Key             =   "Wait"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件:"
         Height          =   180
         Left            =   3120
         TabIndex        =   66
         Top             =   60
         Width           =   450
      End
   End
   Begin MSComDlg.CommonDialog cdgPub 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picStepInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11400
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      Begin MSComctlLib.ImageList imgStep 
         Left            =   555
         Top             =   135
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradeNew.frx":1694
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppUpgradeNew.frx":31E6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "………………"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1365
         TabIndex        =   1
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "………………………………………………………………………………………………………………………………………………"
         Height          =   360
         Left            =   1365
         TabIndex        =   2
         Top             =   390
         Width           =   8790
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   480
         Top             =   60
         Width           =   720
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   13000
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   13000
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "开始升迁(&N)"
      Height          =   350
      Left            =   8652
      TabIndex        =   55
      Top             =   6456
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&C)"
      Height          =   350
      Left            =   10176
      TabIndex        =   56
      Top             =   6456
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   59
      Top             =   6900
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAppUpgradeNew.frx":4D38
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16536
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "14:17"
            Key             =   "STANUM"
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
   Begin MSComctlLib.ProgressBar prgThis 
      Height          =   252
      Left            =   4320
      TabIndex        =   58
      Top             =   6505
      Visible         =   0   'False
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblPer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "##%"
      Height          =   180
      Left            =   7680
      TabIndex        =   57
      Top             =   6541
      Visible         =   0   'False
      Width           =   276
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   13000
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   13000
      Y1              =   6360
      Y2              =   6360
   End
End
Attribute VB_Name = "frmAppUpgradeNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'====================================================================
'==变量
'====================================================================
Private mintStep As Integer '当前页面
Private Const STEP_INFO = _
    "系统升迁设置|升迁时所需用户验证，升迁操作选择，升迁使用参数设置，以及日志记录等。" & _
    "||系统升迁处理|正在进行升迁，请注意当前显示的进度信息；如果出现错误，请仔细查看错误信息，并对错误进行分析之后再采取相应的措施。"
Private Enum IDX_STEP
    SI_升迁设置 = 0
    SI_系统升迁 = 1
End Enum

Private Enum ErrType
    ET_忽略次要错误 = 0
    ET_忽略所有错误 = 1
End Enum

Private Enum StatType
    ST_升级收集 = 0
    ST_仅记录 = 1
End Enum

'流程步骤
Private Const FS_升迁检查 = "UPCHCEK"
Private Const FS_工具升迁 = "TOOLSUP"
Private Const FS_应用系统升迁 = "APPUP"
Private Const FS_历史库升迁 = "HISTORYUP"
Private Const FS_公共同义词 = "PUBSYNONYM"
Private Const FS_编译无效对象 = "COMPILE"
Private Const FS_重整序列 = "ADJUSTSEQ"
Private Const FS_报表升级 = "REPORTUP"
Private Const FS_角色授权 = "ROLEGRANT"
Private Const FS_统计信息收集 = "GATHERSTAT"

'--入口参数
'应用系统升迁参数
Private mrsSysInfo As ADODB.Recordset '各个系统状态
Private mrsSysFiles As ADODB.Recordset '各个系统的升迁文件
Private mblnExecBef As Boolean '是否提前升级
'--返回参数
Private mblnOk As Boolean '是否升级完成后退出
Private mstrRunModule As String '升级后跳转的模块
'--变量
Private mrsHistorySpace As ADODB.Recordset '各个系统历史库信息
Private mrsOptionalProc As ADODB.Recordset '各个系统以及历史库的可选过程
Private mrsReport As ADODB.Recordset '各个系统的报表
Private mblnFinal As Boolean '是否有系统升迁到最终版本
Private mblnHaveST As Boolean '标准版是否在本次升级中
Private mstrSysCodes As String '本次升级的系统编号的字符串，以逗号分割
Private mstrChangeTables As String '本次升级过程中结构发生的变化的表，以逗号分割
Private mclsRunScript As New clsRunScript '脚本运行对象
Private mintDDLParallel As Integer '并行度
Private mblnInstallPLJson As Boolean    '存在安装PLJSON的任务
Private mblnJSONRemain    As Boolean   '存在JSOn安装残留
Private mstrToolsFloder    As String  'TOOLS目录

'====================================================================
'==公共接口
'====================================================================
Public Function ShowMe(frmParent As Object, ByVal rsSysInfo As ADODB.Recordset, ByVal rsSysFiles As ADODB.Recordset, Optional ByVal blnExecBef As Boolean, Optional ByRef strRunModule As String) As Boolean
 '功能：公共入口
 '    :strRunModule=完成升级后跳转的模块
 '返回：是否升级完成后退出
    Set mrsSysInfo = rsSysInfo
    Set mrsSysFiles = rsSysFiles
    mblnExecBef = blnExecBef
    mintStep = -1
    mstrRunModule = ""
    Me.Show 1, frmParent
    strRunModule = mstrRunModule
    ShowMe = mblnOk
End Function

Public Function HistoryUp(frmParent As Object, objStep As Object, ByVal lngSys As Long, ByVal strBakDB As String, ByVal strIntFile As String, ByVal strUserName As String, ByVal strPassword As String, ByVal strServer As String, ByVal strMaxVer As String, ByVal strDbLink As String) As Boolean
 '功能：历史库单独升级接口
 '参数：objStep=显示步骤的对象
 '          lngSys=系统编号
 '          strIntFile=该系统的安装配置文件
 '          strBAKDB=历史库名
 '          strUserName=历史库用户名称
 '          strPassWord=历史库用户密码
 '          strServer=历史库服务器
 '          strMaxVer=历史库目标版本
 '          strBakSpaceName=历史表空间名
 '          strDBLInk=DBLink名称
 '返回：是否升级成功
 '该公共过程仅使用当前窗体的两个对象,mrsSysFiles,与mclsRunScript
    Dim rsTmp As ADODB.Recordset
    Dim cnHistory As ADODB.Connection
    Dim rsUpFiles As ADODB.Recordset
    Dim strSteps  As String, arrStep As Variant, i As Long
    Dim strCurMax As String
    Dim strSQL As String
    
    On Error GoTo errh
    If strIntFile = "" Then
        MsgBox "无效的安装配置文件!", vbInformation, App.Title
        Exit Function
    End If
    '重新实例化，清除使用痕迹
    Set mclsRunScript = New clsRunScript
    If strServer = "" Then strServer = gstrServer
    If strDbLink <> "" Then
        strSQL = "Select Owner, Db_Link, Username, Host" & vbNewLine & _
                    "From All_Db_Links" & vbNewLine & _
                    "Where Owner =[1] And Username =[2] And Db_Link||'.' Like [3]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取DBLink服务器", gstrUserName, UCase(strUserName), UCase(strDbLink) & ".%")
        If Not rsTmp.EOF Then strServer = rsTmp!Host & ""
    End If
    
    '设置参数类参数
    Call mclsRunScript.InitGlobalPara(frmParent, lngSys, False, GetLogPath(LT_历史库升迁, lngSys, strUserName))
    mclsRunScript.Server = strServer
    Set cnHistory = gobjRegister.GetConnection(strServer, strUserName, strPassword, False, MSODBC, "", False)
    If cnHistory.State = adStateOpen Then
        Set rsTmp = ReadHisUpgrade(cnHistory, strUserName, False, lngSys, strDbLink <> "")
        If rsTmp Is Nothing Then
            MsgBox "获取该历史库版本信息失败！该历史库无法升级。", vbInformation, App.Title
            Exit Function
        End If
        If rsTmp.RecordCount = 0 Then
            MsgBox "获取该历史库版本信息失败！该历史库无法升级。", vbInformation, App.Title
            Exit Function
        End If
    Else
        Exit Function
    End If
    
    Call SetSQLTrace(strServer, strUserName, cnHistory)
    
    '最后一个参数应该传strBakDB，但是传了strBakUser，尽管两者有区别，但是不影响脚本获取
    Set mrsSysFiles = GetUpgradeFiles(rsUpFiles, rsTmp!系统编号, rsTmp!当前版本, strIntFile, rsTmp!中止信息, rsTmp!提前中止信息, strMaxVer, , strBakDB)
    mrsSysFiles.Filter = "": mrsSysFiles.Sort = "FullSPVer"
    Do While Not mrsSysFiles.EOF
        If InStr(strSteps & ",", "," & mrsSysFiles!SPVer & ",") = 0 Then
            strSteps = strSteps & "," & mrsSysFiles!SPVer
            strCurMax = mrsSysFiles!SPVer
        End If
        mrsSysFiles.MoveNext
    Loop
    If strCurMax <> strMaxVer Then '没有脚本，或目标版本没有脚本，都添加一个版本的流程
        strSteps = strSteps & "," & strMaxVer
    End If
    strSteps = strSteps & "," & "历史库结构修正"
    strSteps = Mid(strSteps, 2)
    arrStep = Split(strSteps, ",")
    For i = LBound(arrStep) To UBound(arrStep)
        objStep.Text = IIf(i = UBound(arrStep), "", "升迁到") & arrStep(i)
        objStep.ToolTipText = IIf(i = UBound(arrStep), "", "升迁到") & arrStep(i)
        If i = UBound(arrStep) Then '历史库结构修正
            Call RepairHisDB(cnHistory, lngSys, strUserName, strServer, strBakDB, strDbLink, , True)
        Else '升迁
            Call RunScriptByVersion(lngSys, arrStep(i), i = LBound(arrStep), , , True, cnHistory, strBakDB, True)
        End If
    Next
    Call mclsRunScript.CloseLog
    Set mclsRunScript = Nothing
    Set mrsSysFiles = Nothing
    HistoryUp = True
    Exit Function
errh:
    Call mclsRunScript.CloseLog
    Set mclsRunScript = Nothing
    Set mrsSysFiles = Nothing
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Function

Public Function ToolsInstallUp(frmParent As Object, objStep As Object, ByVal lngSys As Long, ByVal strInstallFile As String, ByVal strLogFile As String) As Boolean
'功能：系统安装中管理工具版本较低时的升迁接口
'参数：
'       frmParent=父窗体
'       objStep=显示步骤的对象
'       lngSys=将要安装的应用系统的序号
'       strInstallFile   应用系统安装脚本的完整位置
'       strLogFile=系统安装日志
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim strToolsVer As String, strMaxToolsVer As String, strCurMax As String
    Dim cnTools As New ADODB.Connection, rsINI As ADODB.Recordset
    Dim strPath As String
    Dim objSys As New Scripting.FileSystemObject
    Dim strBeforeInfo As String, strNormalInfo As String
    Dim strSteps As String, arrStep As Variant, i As Long

    On Error GoTo errh
    '1、检查安装配置文件
    If Not CheckInitFile(lngSys, strInstallFile, , rsINI) Then Exit Function
    rsINI.Filter = "项目='管理工具版本号'"
    If Not rsINI.EOF Then strMaxToolsVer = rsINI!内容 & ""
    '2、判断管理工具的版本
    Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Ver")
    If rsTmp.EOF Then
        '如果没有，就进行版本检查，主要是以前没有版本控制
        strToolsVer = JudgeOldToolsVer
        '并且更新数据库
        Call OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Update_Ver", strToolsVer)
    Else
        '产生一个12位的数字
        strToolsVer = rsTmp("内容") & ""
    End If
    '3、比较版本，是否需要升级
    If VerFull(strToolsVer) >= VerFull(strMaxToolsVer) Then
        '满足要求，不需要升级
        ToolsInstallUp = True
        Exit Function
    End If
    '4、获取升级脚本目录
    On Error Resume Next
    strPath = objSys.GetParentFolderName(objSys.GetParentFolderName(objSys.GetParentFolderName(strInstallFile))) & "\Tools\ZLSERVER.SQL"
    If err.Number <> 0 Then err.Clear
    If gobjFSO.FileExists(strPath) Then
        mstrToolsFloder = gobjFSO.GetParentFolderName(strPath)
    End If
    On Error GoTo errh
    If Not objSys.FileExists(strPath) Then
        MsgBox "打开管理脚本存放目录（[安装目录]\Tools）错误。", vbInformation, gstrSysName
        Exit Function
    End If
    '获取管理工具上次升迁与提前升迁的中止信息
    '检查ZLUPGRADE表及其字段”提前执行“
    If CheckAndAdjustMustTable("ZLUPGRADE", "提前执行", False) Then
        '获取所有系统上次升迁以及上次提前升迁信息
        strSQL = "Select  提前执行, 中止语句, 升迁结果, 结果版本" & vbNewLine & _
                        "From (Select 提前执行, 升迁时间, 中止语句, 升迁结果, 结果版本, Max(升迁时间) Over(Partition By Decode(提前执行, Null, -1, 0)) 当前时间" & vbNewLine & _
                        "       From Zlupgrade Where 系统 is Null) a" & vbNewLine & _
                        "Where A.升迁时间 = A.当前时间 "
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取上次升迁信息")
        '系统上次执行升迁信息
        rsTmp.Filter = "提前执行=Null"
        If Not rsTmp.EOF Then
            strNormalInfo = FormatUpgradeBreak(0, rsTmp!结果版本 & "", rsTmp!中止语句 & "")
        Else
            strNormalInfo = FormatUpgradeBreak(0, strToolsVer)
        End If
        '系统上次提前执行升迁信息
        rsTmp.Filter = "提前执行<>Null"
        If Not rsTmp.EOF Then
            strBeforeInfo = FormatUpgradeBreak(0, rsTmp!结果版本 & "", rsTmp!中止语句 & "")
        Else
            strBeforeInfo = FormatUpgradeBreak(0, strToolsVer)
        End If
    Else
        strBeforeInfo = FormatUpgradeBreak(0, strToolsVer)
        strNormalInfo = FormatUpgradeBreak(0, strToolsVer)
    End If
    '获取升迁脚本
    Set mrsSysFiles = GetUpgradeFiles(Nothing, 0, strToolsVer, strPath, strNormalInfo, strBeforeInfo, strMaxToolsVer, strCurMax, , True)
    If VerFull(strCurMax) < VerFull(strMaxToolsVer) Then
        '脚本支持到的版本小于要求升级到的版本，则不能升级
        MsgBox "缺少管理工具" & strMaxToolsVer & "版本的升迁脚本！", vbInformation, gstrSysName
        Exit Function
    End If
    '6、连接zltools
    Set gcnTools = GetConnection("ZLTOOLS")
    If gcnTools Is Nothing Then
        MsgBox "无法以ZLTOOLS用户连接!", vbInformation, gstrSysName
        Exit Function
    End If
    '7、创建脚本解析执行类
    '重新实例化，清除使用痕迹
    Set mclsRunScript = New clsRunScript
    '设置参数类参数
    Call mclsRunScript.InitGlobalPara(frmParent, 0, False, strLogFile)
    mclsRunScript.Server = gstrServer
    mclsRunScript.WriteLog "管理工具版本较低，无法支持该版本应用系统安装。"
    mclsRunScript.WriteLog "管理工具自动升级：" & strToolsVer & "->" & strMaxToolsVer
    Set gcnSystem = gcnOracle '系统安装才调用管理工具单独升级，此时gcnOracle为DBA连接
    'PLJSON安装
    If IsCanInstallPLJson(mstrToolsFloder, mblnJSONRemain) Then
        Call InstallPLJSON(gcnSystem, mstrToolsFloder, mclsRunScript, mblnJSONRemain)
    End If
    mrsSysFiles.Filter = "": mrsSysFiles.Sort = "FullSPVer"
    Do While Not mrsSysFiles.EOF
        If InStr(strSteps & ",", "," & mrsSysFiles!SPVer & ",") = 0 Then
            strSteps = strSteps & "," & mrsSysFiles!SPVer
            strCurMax = mrsSysFiles!SPVer
        End If
        mrsSysFiles.MoveNext
    Loop
    strSteps = strSteps & "," & "对象授权修正"
    strSteps = Mid(strSteps, 2)
    arrStep = Split(strSteps, ",")
    mclsRunScript.SysNo = 0
    For i = LBound(arrStep) To UBound(arrStep)
        objStep.Text = IIf(i = UBound(arrStep), "", "管理工具升迁到") & arrStep(i)
        objStep.ToolTipText = IIf(i = UBound(arrStep), "", "管理工具升迁到") & arrStep(i)
        If i = UBound(arrStep) Then '对象授权修正
            gcnOracle.Execute "Update zlUpGrade Set 提前执行=0 Where 提前执行 = 1 And 系统 is Null "
            Call ReGrantForTools(gcnTools, , True)
        Else '升迁
            If Not RunScriptByVersion(0, arrStep(i), i = LBound(arrStep), strToolsVer, strMaxToolsVer, , , , True) Then
                MsgBox "管理工具自动升级失败，请查看日志，做相应处理！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Next
    mclsRunScript.WriteLog "管理工具自动升级成功！"
    Call mclsRunScript.CloseLog
    Set mclsRunScript = Nothing
    Set mrsSysFiles = Nothing
    ToolsInstallUp = True
    Exit Function
errh:
    mclsRunScript.WriteLog "管理工具自动升级失败！"
    Call mclsRunScript.CloseLog
    Set mclsRunScript = Nothing
    Set mrsSysFiles = Nothing
    If MsgBox("出现下列错误，是否继续？" & vbCrLf & "    " & err.Description, vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        Resume
    End If
End Function

'====================================================================
'==控件事件
'====================================================================
Private Sub chkHisAll_Click()
    Call RecUpdate(mrsHistorySpace, "", "升级", IIf(chkHisAll.value = 0, 0, 1))
    Call RecUpdate(mrsHistorySpace, "升级=0 And 当前=1", "升级", 1) '当前历史库必须升级
    '重新读取可选脚本
    Call ReadOptionalProc(True)
    '刷新历史库汇总信息
    Call RefreshTotalInfo(0)
End Sub

Private Sub chkLogLong_Click()
    Call SetCtrlEnabled(chkLogLong.value = 1, txtLogLong)
End Sub

Private Sub chkOpt_Click()
    Call SetCtrlEnabled(chkOpt.value = 1, lblOptSel, lblOptTotal)
    Call RecUpdate(mrsOptionalProc, "", "执行", IIf(chkOpt.value = 1, 1, 0))
    Call RefreshTotalInfo(2)
    lblOptSel.Visible = (chkOpt.value = 1): lblOptTotal.Visible = (chkOpt.value = 1)
End Sub

Private Sub chkParallel_Click()
    Call SetCtrlEnabled(chkParallel.value = 1, lblParallel, txtCpu(0), udCpu(0))
    lblCpuWarn(0).Visible = chkParallel.value = 1 And lblCpuWarn(0).Tag <> ""
End Sub

Private Sub chkRpt_Click()
    Call SetCtrlEnabled(chkRpt.value = 1, optRpt(0), optRpt(1), lblRptSel, lblRptTotal)
    Call RecUpdate(mrsReport, "", "覆盖类型", IIf(chkRpt.value = 1, IIf(optRpt(0).value, "!默认覆盖类型", 2), 0))
    Call RefreshTotalInfo(1)
    lblRptSel.Visible = (chkRpt.value = 1): lblRptTotal.Visible = (chkRpt.value = 1)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    Call StepSwitch(1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Dim strSysCodes As String, i As Long
    
    Dim blnHaveDBA As Boolean
    Dim blnHaveApp As Boolean '是否有应用系统需要升级
    
    Call ApplyOEM(stbThis)
    If Not mblnExecBef Then ShowFlash ("正在收集升级所需要数据资源，请稍候！")
    mrsSysInfo.Filter = "系统编号<>0 And 升级=1"
    blnHaveApp = mrsSysInfo.RecordCount <> 0
    '//////////////////////////////////////////////////////////////////////
    '///////////////           界面数据初始化////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    '修正ZLupgrade的目标版本字段，防止目标版本是特殊SP导致的数据更新出错。
    Call AdjustZLupgrade
    '读取历史库
    Call ReadHistorySpace
    '读取报表
    Call ReadImpReports
    '读取可选过程
    Call ReadOptionalProc
    '汇总信息刷新
    Call RefreshTotalInfo
    '是否存在PLJSON安装任务
    If Not mblnExecBef Then
        mrsSysInfo.Filter = "系统编号=0"
        On Error Resume Next
        mstrToolsFloder = gobjFSO.GetParentFolderName(mrsSysInfo!配置文件 & "")
        If err.Number <> 0 Then err.Clear
        If mstrToolsFloder <> "" Then
            mblnInstallPLJson = IsCanInstallPLJson(mstrToolsFloder, mblnJSONRemain)
        End If
    End If
    '提前执行才有在线创建索引
    ckhIdxOnLine.Visible = mblnExecBef: ckhIdxOnLine.value = IIf(mblnExecBef And blnHaveApp Or Not blnHaveApp, 1, 0)
    ckhIdxOnLine.Enabled = blnHaveApp
    '提前执行默认不收集统计信息
    chkStat.Enabled = blnHaveApp
    chkStat.value = IIf(mblnExecBef And blnHaveApp Or Not blnHaveApp, 0, 1)
    '设置并行度
    Call SetCpuCount
    chkParallel.value = IIf(blnHaveApp, 0, chkParallel.value)
    chkParallel.Enabled = chkParallel.Enabled And blnHaveApp
    '日志路径获取
    mrsSysInfo.Filter = "升级=1": mrsSysInfo.Sort = "Sort"
    For i = 0 To mrsSysInfo.RecordCount - 1
        strSysCodes = strSysCodes & "," & mrsSysInfo!系统编号
        mrsSysInfo.MoveNext
    Next
    lblLog.Tag = GetLogPath(IIf(mblnExecBef, LT_提前升迁, LT_常规升迁), Mid(strSysCodes, 2))  '保存默认路径
    '以前注册表中存在日志路径，则将该路径做为初始路径,以前UpgradeLogDir+编号的就不再使用
    If gobjFile.FolderExists(GetSetting("ZLSOFT", "公共模块", "UpgradeLogDir", "")) Then
        '尽管文件不存在，仍然可以用gobjFile.GetFileName来获取文件名，只要不是打开
        lblLogModi.Tag = GetSetting("ZLSOFT", "公共模块", "UpgradeLogDir", "") & "\" & gobjFile.GetFileName(lblLog.Tag)
    Else
        lblLogModi.Tag = lblLog.Tag
    End If
    lblLog.Caption = "升迁日志文件：" & lblLogModi.Tag
    lblLog.ToolTipText = lblLogModi.Tag
    If lblLog.Width >= 8000 Then
        lblLog.Width = 8000 '防止丢失修改标签
    End If
    Call SetCtrlPosOnLine(False, 0, lblLog, 60, lblLogModi)
    '//////////////////////////////////////////////////////////////////////
    '/////////////// 用户验证相关控件默认值////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    'ZLTOOLS
    mrsSysInfo.Filter = "升级=1 And 系统编号=0"
    Call SetCtrlEnabled(Not mrsSysInfo.EOF, lblToolsUser, lblToolsPwd, txtToolsPwd)
    txtToolsPwd.BackColor = IIf(txtToolsPwd.Enabled, &H80000005, &H8000000F)
    If Not GetConnection("ZLTOOLS", False) Is Nothing Then
        txtToolsPwd.Text = gstrToolsPwd
    End If
    'DBA用户
    mrsSysFiles.Filter = " FileType=" & FT_DBA
    If Not mrsSysFiles.EOF Then lblDBA.Tag = 1 '标记存在DBA脚本
    Call SetCtrlEnabled(lblDBA.Tag <> "" Or chkStat.value = 1 And optStatType(ST_升级收集).value Or mblnInstallPLJson, lblDBAUser, txtDBAUser, lblDBAPwd, txtDBAPwd)
    txtDBAUser.Text = IIf(gstrSysUser = "", "System", gstrSysUser)
    If Not GetConnection("DBA", False) Is Nothing Then
        txtDBAPwd.Text = gstrSysPwd
    End If
    txtDBAUser.BackColor = IIf(txtDBAUser.Enabled, &H80000005, &H8000000F)
    txtDBAPwd.BackColor = IIf(txtDBAPwd.Enabled, &H80000005, &H8000000F)
    '//////////////////////////////////////////////////////////////////////
    '///////////////直接调用控件事件来刷新界面////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    '统计信息收集相关控件设置
    Call chkStat_Click
    '并行DDL相关控件可用性设置
    Call chkParallel_Click
    '界面数据展示
    Call cmdNext_Click
    '查看是否存在最终版本
    If Not mblnExecBef Then
        mblnFinal = True
        mrsSysInfo.Filter = "升级=1 And 系统编号<>0 And 目标版本<>Null"
        mrsSysInfo.Sort = "系统编号"
        Do While Not mrsSysInfo.EOF
            '存在一个系统不能升迁到最终版本，即不进行角色授权
            If mrsSysInfo!目标版本 & "" <> mrsSysInfo!最终版本 & "" Then
                mblnFinal = False: Exit Do
            End If
            mrsSysInfo.MoveNext
        Loop
    Else
        mblnFinal = False
    End If
    If Not mblnExecBef Then ShowFlash ("")
End Sub

Private Sub chkAutoPer_Click()
    txtEsper.Text = IIf(chkAutoPer.value = 1, "", 10)
    Call SetCtrlEnabled(chkAutoPer.value = 0, txtEsper, udEsper)
End Sub

Private Sub chkStat_Click()
    lblCpuWarn(1).Visible = chkStat.value = 1 And lblCpuWarn(1).Tag <> ""
    Call SetCtrlEnabled(chkStat.value = 1, fraStatType, optStatType(0), optStatType(1))
    Call SetCtrlEnabled(chkStat.value = 1, lblStaInfo, lblEsperName, chkAutoPer, lblEsperNote)
    Call SetCtrlEnabled(chkStat.value = 1 And chkAutoPer.value = 0, txtEsper, udEsper)
    Call SetCtrlEnabled(chkStat.value = 1 And lblStaCpuName.Tag = "", lblStaCpuName, lblStaCpuNote, txtCpu(1), udCpu(1))
    'DBA用户设置
    Call SetCtrlEnabled(lblDBA.Tag <> "" Or chkStat.value = 1 And optStatType(ST_升级收集).value Or mblnInstallPLJson, lblDBAUser, txtDBAUser, lblDBAPwd, txtDBAPwd)
    txtDBAUser.BackColor = IIf(txtDBAUser.Enabled, &H80000005, &H8000000F)
    txtDBAPwd.BackColor = IIf(txtDBAPwd.Enabled, &H80000005, &H8000000F)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOk Then
        If Not cmdCancel.Enabled Then
            Cancel = 1: Exit Sub
        ElseIf mintStep < SI_系统升迁 Then
            If MsgBox("要退出系统升迁向导吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    Set mrsSysInfo = Nothing
    Set mrsSysFiles = Nothing
    Set mrsHistorySpace = Nothing
    Set mrsOptionalProc = Nothing
    Set mrsReport = Nothing
    Set mclsRunScript = Nothing
End Sub

Private Sub lblHisSel_Click()
    '重新读取历史库升迁文件
    If frmAppUpgradeSel.ShowMe(Me, AST_His, mrsHistorySpace, mrsSysFiles, mblnExecBef) Then
    End If
    '重新读取可选过程,历史库可能也有存储过程
    Call ReadOptionalProc(True)
    '刷新历史库汇总信息
    Call RefreshTotalInfo(0)
End Sub

Private Sub lblLogModi_Click()
    With cdgPub
        .DialogTitle = "确定升迁日志文件"
        .Filter = "升迁日志文件(*.log)|*.log"
        .flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
        .InitDir = gobjFile.GetParentFolderName(lblLogModi.Tag)
        .FileName = gobjFile.GetFileName(lblLogModi.Tag)
        .CancelError = True
        On Error GoTo errh
        .ShowSave
        On Error GoTo 0
        lblLogModi.Tag = .FileName
        SaveSetting "ZLSOFT", "公共模块", "UpgradeLogDir", gobjFile.GetParentFolderName(.FileName)
        lblLog.Caption = "升迁日志文件：" & lblLogModi.Tag
        lblLog.ToolTipText = lblLogModi.Tag
        lblLog.Refresh
        If lblLog.Width >= 8000 Then
            lblLog.Width = 8000
        End If
        Call SetCtrlPosOnLine(False, 0, lblLog, 60, lblLogModi)
    End With
errh:
End Sub

Private Sub lblOptSel_Click()
    If frmAppUpgradeSel.ShowMe(Me, AST_OptProc, mrsOptionalProc) Then
    End If
    Call RefreshTotalInfo(2)
End Sub

Private Sub lblRptSel_Click()
    If frmAppUpgradeSel.ShowMe(Me, AST_Report, mrsReport) Then
    End If
    Call RefreshTotalInfo(1)
End Sub

Private Sub optErrOption_Click(Index As Integer)
    If Index = ET_忽略次要错误 Then
        optErrOption(ET_忽略所有错误).ForeColor = &H80000012
    Else
        optErrOption(ET_忽略所有错误).ForeColor = &H80000012
        MsgBox "忽略所有错误可能会造成一些错误不能得到有效处理！", vbInformation, gstrSysName
    End If
End Sub

Private Sub optRpt_Click(Index As Integer)
    Call RecUpdate(mrsReport, "", "覆盖类型", Index + 1)
    Call RefreshTotalInfo(1)
End Sub

Private Sub tmrRefresh_Timer()
    Me.Refresh
End Sub

Private Sub txtCpu_GotFocus(Index As Integer)
    Call SelAll(txtCpu(Index))
End Sub

Private Sub txtCpu_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtCpu_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtCpu(Index).Text) < udCpu(Index).Min Then
        udCpu(Index).value = udCpu(Index).Min
    ElseIf Val(txtCpu(Index).Text) > udCpu(Index).Max Then
        udCpu(Index).value = udCpu(Index).Max
    End If
End Sub

Private Sub txtDBAPwd_GotFocus()
    Call SelAll(txtDBAPwd)
End Sub

Private Sub txtDBAPwd_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim cnTmp As New ADODB.Connection
    Dim strErr As String
    
    If txtDBAPwd.Text <> "" And txtDBAUser.Text <> "" Then
        If UCase(txtDBAUser.Text) = UCase(gstrSysUser) And UCase(gstrSysPwd) <> UCase(txtDBAPwd.Text) Then
            MsgBox "DBA用户密码错误！", vbInformation, gstrSysName
             Cancel = True: Exit Sub
        End If
        If gcnSystem Is Nothing Then
            Set gcnSystem = New ADODB.Connection
        ElseIf gcnSystem.State = 1 Then
            gcnSystem.Close
        End If
        err.Clear
       
        Set cnTmp = gobjRegister.GetConnection(gstrServer, txtDBAUser.Text, txtDBAPwd.Text, False, MSODBC, strErr, False)
        If cnTmp.State = adStateClosed Then
            MsgBox strErr, vbInformation, "验证失败"
            Call SelAll(txtDBAPwd)
            Cancel = True: Exit Sub
        End If
        
        '检查是否DBA
        strSQL = "Select 1 From Session_Roles Where Role = 'DBA'"
        Set rsTmp = gclsBase.OpenSQLRecord(cnTmp, strSQL, "是否DBA")
        If rsTmp.EOF Then
            MsgBox "该用户不具有DBA身份！", vbExclamation, gstrSysName
            txtDBAUser.SetFocus: Exit Sub
        End If
        
        '暂时不设置SetSQLTrace，执行前再设置
        Set gcnSystem = cnTmp
        gstrSysUser = txtDBAUser.Text
        gstrSysPwd = txtDBAPwd.Text
    End If
End Sub

Private Sub txtDBAUser_GotFocus()
    Call SelAll(txtDBAUser)
End Sub

Private Sub txtDBAUser_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim cnTmp As New ADODB.Connection
    Dim strErr As String
    
    If UCase(txtDBAUser.Text) = UCase(gstrSysUser) And gstrSysUser <> "" Then
        txtDBAPwd.Text = gstrSysPwd
    Else
        txtDBAPwd.Text = ""
    End If
    If txtDBAPwd.Text <> "" And txtDBAUser.Text <> "" Then
        If UCase(txtDBAUser.Text) = UCase(gstrSysUser) And UCase(gstrSysPwd) <> UCase(txtDBAPwd.Text) Then
            MsgBox "DBA用户密码错误！", vbInformation, gstrSysName
             Cancel = True: Exit Sub
        End If
        If gcnSystem Is Nothing Then
            Set gcnSystem = New ADODB.Connection
        ElseIf gcnSystem.State = 1 Then
            gcnSystem.Close
        End If
        err.Clear
        
        Set cnTmp = gobjRegister.GetConnection(gstrServer, txtDBAUser.Text, txtDBAPwd.Text, False, MSODBC, "", False)
        If cnTmp.State = adStateClosed Then
            Cancel = True: Exit Sub
        End If
        On Error GoTo 0
        '检查是否DBA
        strSQL = "Select 1 From Session_Roles Where Role = 'DBA'"
        Set rsTmp = gclsBase.OpenSQLRecord(cnTmp, strSQL, "是否DBA")
        If rsTmp.EOF Then
            MsgBox "该用户不具有DBA身份！", vbExclamation, gstrSysName
            txtDBAUser.SetFocus: Exit Sub
        End If
        
        '暂时不设置SetSQLTrace，执行前再设置
        Set gcnSystem = cnTmp
        gstrSysUser = txtDBAUser.Text
        gstrSysPwd = txtDBAPwd.Text
    End If
End Sub

Private Sub txtEsper_GotFocus()
    Call SelAll(txtEsper)
End Sub

Private Sub txtEsper_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtEsper_Validate(Cancel As Boolean)
    If Val(txtEsper.Text) < udEsper.Min Then
        udEsper.value = udEsper.Min
    ElseIf Val(txtEsper.Text) > udEsper.Max Then
        udEsper.value = udEsper.Max
    End If
End Sub

Private Sub txtHisPwd_GotFocus()
    Call SelAll(txtHisPwd)
End Sub

Private Sub txtHisPwd_Validate(Cancel As Boolean)
    Dim cnTmp As ADODB.Connection
    Dim rsTmp As ADODB.Recordset
    Dim cllBakDB As New Collection, Item As Variant, arrTmp As Variant
    Dim strMaxVer As String, strFilter As String, strTmp As String
    Dim strBakName As String
    
    If txtHisPwd.Text <> "" And txtHisPwd.Tag <> Trim(txtHisPwd.Text) Then
        mrsHistorySpace.Filter = "验证=0"
        mrsHistorySpace.Sort = "名称,所有者,服务器"
        ShowFlash ("正在验证历史库，并获取历史库升迁脚本，请稍候！")
        DoEvents
        On Error Resume Next
        Do While Not mrsHistorySpace.EOF
            strTmp = mrsHistorySpace!所有者 & ";" & mrsHistorySpace!服务器 & ";" & mrsHistorySpace!DB连接
            cllBakDB.Add strTmp, strTmp
            If err.Number <> 0 Then err.Clear
            mrsHistorySpace.MoveNext
        Loop
        On Error GoTo errh
        For Each Item In cllBakDB
            arrTmp = Split(Item, ";")
            
            Set cnTmp = gobjRegister.GetConnection(arrTmp(1), arrTmp(0), txtHisPwd.Text, False, MSODBC, "", False)
            If cnTmp.State = adStateOpen Then
                 '暂时不设置SetSQLTrace，执行前再设置
                
                Set rsTmp = ReadHisUpgrade(cnTmp, arrTmp(0), , , arrTmp(2) <> "")
                Call RecUpdate(mrsHistorySpace, "所有者='" & arrTmp(0) & "' And 服务器='" & arrTmp(1) & "' And 验证=0", "验证", 1)
                rsTmp.Sort = ""
                If rsTmp.EOF Then
                    Call RecUpdate(mrsHistorySpace, "所有者='" & arrTmp(0) & "' And 服务器='" & arrTmp(1) & "'", "密码", txtHisPwd.Text, "可升级", 0, "可提前升级", 0, "检查结果", "历史表空间数据结构缺失导致无法升级！")
                Else
                    Do While Not rsTmp.EOF
                        mrsHistorySpace.Filter = "系统编号=" & rsTmp!系统编号 & " And 所有者='" & arrTmp(0) & "' And 服务器='" & arrTmp(1) & "'"
                        Do While Not mrsHistorySpace.EOF
                            If mrsHistorySpace!验证 = 1 Then mrsHistorySpace.Update "验证", 2
                            strBakName = UCase(mrsHistorySpace!名称 & "")
                            mrsHistorySpace.Update Array("密码", "当前版本", "中止信息", "提前中止信息"), Array(txtHisPwd.Text, rsTmp!当前版本, rsTmp!中止信息, rsTmp!提前中止信息)
                            '判断能否升迁
                            If Not IsVerSion(rsTmp!当前版本 & "") Then
                                mrsHistorySpace.Update Array("可升级", "检查结果", "可提前升级"), Array(0, "历史数据空间的版本不可识别。请检查！", 0)
                            ElseIf VerFull(rsTmp!当前版本 & "") >= VerFull(mrsHistorySpace!目标版本 & "") Then '标识为无需升级
                                mrsHistorySpace.Update Array("可升级", "检查结果", "可提前升级"), Array(0, "历史数据空间的版本高于本次升迁目标版本，不能升迁！", 0)
                            Else
                                Set mrsSysFiles = GetUpgradeFiles(mrsSysFiles, rsTmp!系统编号, rsTmp!当前版本, mrsHistorySpace!配置文件, rsTmp!中止信息, rsTmp!提前中止信息, mrsHistorySpace!目标版本, , strBakName)
                                '获取提前执行的目标版本
                                If mblnExecBef Then
                                    strFilter = "所有者='" & strBakName & "' And FileType=" & FT_Before
                                    mrsSysFiles.Filter = strFilter: mrsSysFiles.Sort = "FullSPVer Desc": strMaxVer = ""
                                    If Not mrsSysFiles.EOF Then
                                        strMaxVer = mrsSysFiles!SPVer
                                        mrsSysFiles.Filter = strFilter & " And 配置版本>'" & VerFull(rsTmp!当前版本 & "") & "'": mrsSysFiles.Sort = "FullSPVer"
                                        If Not mrsSysFiles.EOF Then
                                            mrsSysFiles.Filter = strFilter & " And FullSPVer<'" & mrsSysFiles!FullSPVer & "'": mrsSysFiles.Sort = "FullSPVer Desc"
                                            If Not mrsSysFiles.EOF Then
                                                strMaxVer = mrsSysFiles!SPVer
                                            Else
                                                strMaxVer = ""
                                                mrsHistorySpace.Update Array("可提前升级", "提前检查结果"), Array(0, "没有可执行的提前升级脚本，不能提前升迁！")
                                            End If
                                        End If
                                    Else
                                        mrsHistorySpace.Update Array("可提前升级", "提前检查结果"), Array(0, "没有提前升级脚本，不能提前升迁！")
                                    End If
                                    mrsHistorySpace.Update "提前目标版本", strMaxVer
                                    '删除非提前执行脚本
                                    Call RecDelete(mrsSysFiles, "所有者='" & strBakName & "' And FileType<>" & FT_Before)
                                    '删除大于提前目标版本的提前升级脚本
                                    Call RecDelete(mrsSysFiles, strFilter & " And FullSPVer>'" & VerFull(strMaxVer) & "'")
                                End If
                            End If
                            mrsHistorySpace.MoveNext
                        Loop
                        rsTmp.MoveNext
                    Loop
                End If
                '标记未在历史空间中注册
                Call RecUpdate(mrsHistorySpace, "验证=1", "可升级", 0, "可提前升级", 0, "检查结果", "该系统的历史空间未在ZLBakInfo中注册！")
            End If
        Next
        txtHisPwd.Tag = Trim(txtHisPwd.Text)
        '重新读取可选脚本
        Call ReadOptionalProc(True)
        '刷新历史库汇总信息
        Call RefreshTotalInfo(0)
        ShowFlash ("")
        Me.Refresh
    End If
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Private Sub txtLogLong_GotFocus()
    Call SelAll(txtLogLong)
End Sub

Private Sub txtLogLong_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) <= 0 Then KeyAscii = 0
End Sub

Private Sub txtLogLong_Validate(Cancel As Boolean)
    If Val(txtLogLong.Text) < 1 Then txtLogLong.Text = 1
End Sub

Private Sub txtToolsPwd_GotFocus()
    Call SelAll(txtToolsPwd)
End Sub

Private Sub txtToolsPwd_Validate(Cancel As Boolean)
    Dim strErr As String
    
    If txtToolsPwd.Text <> "" Then
        If gstrToolsPwd <> "" And UCase(gstrToolsPwd) <> UCase(Trim(txtToolsPwd.Text)) Then
             MsgBox "管理工具密码错误！", vbInformation, gstrSysName
             Cancel = True: Exit Sub
        End If
        err.Clear: On Error Resume Next
        If gcnTools Is Nothing Then
            Set gcnTools = New ADODB.Connection
        ElseIf gcnTools.State = 1 Then
            gcnTools.Close
        End If
                
        Set gcnTools = gobjRegister.GetConnection(gstrServer, "zltools", txtToolsPwd.Text, False, MSODBC, "", False)
        If gcnTools.State = adStateClosed Then
            MsgBox "连接管理工具用户时出现错误：" & vbCrLf & vbCrLf & strErr, vbCritical, gstrSysName
            Cancel = True: Exit Sub
        End If
        Call SetSQLTrace(gstrServer, "zltools", gcnTools)
        gstrToolsPwd = txtToolsPwd.Text '赋值
    End If
End Sub

Private Sub udCpu_Change(Index As Integer)
    Call SelAll(txtCpu(Index))
End Sub

Private Sub udEsper_Change()
    Call SelAll(txtEsper)
End Sub

'====================================================================
'==方法
'====================================================================
Private Sub ReadImpReports()
'获取选择升级系统的可导入报表
    Dim strIniPath As String
    Dim blnDo As Boolean, blnAdd As Boolean
    Dim rsINI As ADODB.Recordset
    Dim arrTmp As Variant
    Dim lngID As Long
    Dim strVer As String
    
    On Error GoTo errh
    Set mrsReport = CopyNewRec(Nothing, True, , Array("ID", adInteger, Empty, Empty, "系统编号", adInteger, Empty, Empty, "系统名称", adVarChar, 50, Empty, "SPVer", adVarChar, 30, Empty, "FULLSPVer", adVarChar, 30, Empty, "编号", adVarChar, 20, Empty, "名称", adVarChar, 30, Empty, _
                                                                                        "FilePath", adVarChar, 1000, Empty, "FileName", adVarChar, 200, Empty, "覆盖类型", adInteger, Empty, Empty, "默认覆盖类型", adInteger, Empty, Empty))
    If mblnExecBef Then Exit Sub '提前升迁，只初始化记录集即可
    mrsSysInfo.Filter = "升级=1"
    mrsSysInfo.Sort = "系统编号"
    Do While Not mrsSysInfo.EOF
        strIniPath = gobjFile.GetParentFolderName(gobjFile.GetParentFolderName(mrsSysInfo!配置文件)) & "\导出报表"
        blnDo = gobjFile.FileExists(strIniPath & "\zlReport.ini")
        If blnDo Then
            Set rsINI = ReadINIToRec(strIniPath & "\zlReport.ini")
            blnDo = rsINI.RecordCount > 0
        End If
        If blnDo Then
            Do While Not rsINI.EOF
                blnAdd = IsVerSion(rsINI!项目 & "")
                If blnAdd Then
                    strVer = rsINI!项目 & ""
                    blnAdd = VerFull(rsINI!项目 & "") > VerFull(mrsSysInfo!系统版本号)
                    If blnAdd Then
                        blnAdd = VerFull(rsINI!项目 & "") <= VerFull(mrsSysInfo!目标版本)
                    End If
                    If blnAdd Then
                        arrTmp = Split(rsINI!内容, "|")
                        blnAdd = gobjFile.FileExists(strIniPath & "\" & arrTmp(2))
                    End If
                End If
                If blnAdd Then
                    mrsReport.Filter = "编号='" & UCase(arrTmp(0)) & "'"
                    blnAdd = mrsReport.EOF
                    If blnAdd Then
                        mrsReport.AddNew Array("ID", "系统编号", "系统名称", "SPVer", "编号", "名称", "FilePath", "FileName", "覆盖类型", "默认覆盖类型"), _
                                                        Array(Identity(lngID), mrsSysInfo!系统编号, mrsSysInfo!系统名称, strVer, UCase(Trim(arrTmp(0))), UCase(Trim(arrTmp(1))), strIniPath & "\" & arrTmp(2), arrTmp(2), IIf(Val(arrTmp(3)) = 0, 1, 2), IIf(Val(arrTmp(3)) = 0, 1, 2))
                    Else
                        mrsReport.Update Array("覆盖类型", "默认覆盖类型", "SPVer"), Array(IIf(Val(arrTmp(3)) = 0, 1, 2), IIf(Val(arrTmp(3)) = 0, 1, 2), strVer)
                    End If
                End If
                rsINI.MoveNext
            Loop
        End If
        mrsSysInfo.MoveNext
    Loop
    Exit Sub
errh:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub ReadHistorySpace()
    Dim rsSpaces As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strServer As String
    Dim lngID As Long
    
    On Error GoTo errh
    '必要结构检查
    If Not CheckAndAdjustMustTable("Zlbakspaces", , True) Then
        Exit Sub
    End If
    If Not CheckAndAdjustMustTable("ZLBAKTABLES", , True) Then
        Exit Sub
    End If
    strSQL = "Select 系统, 编号, 名称, 所有者, Db连接, 当前 From Zltools.Zlbakspaces"
    Set rsSpaces = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    '升级：=1，选择升级；=0，不选择升级，-1，选择升级但是改变了服务器名,该状态是中间状态
    '可升级：=1,可以常规升级，=0,不能进行常规升级
    '可提前升级：=1,可以提前升级，=0,不能进行提前升级
    '验证：=0,该历史库未通过验证，=1，该历史库用户通过验证，但是历史空间未注册该历史库，=2，验证成功
    '注意历史库的主键为：系统编号,名称
    Set mrsHistorySpace = CopyNewRec(Nothing, True, , Array("ID", adInteger, Empty, Empty, "系统编号", adInteger, Empty, Empty, "系统名称", adVarChar, 50, Empty, "系统版本", adVarChar, 20, Empty, "配置文件", adVarChar, 2000, Empty, _
                                                                                                "编号", adInteger, Empty, Empty, "名称", adVarChar, 30, Empty, "所有者", adVarChar, 50, Empty, _
                                                                                                "当前", adInteger, Empty, Empty, "DB连接", adVarChar, 200, Empty, "密码", adVarChar, 100, Empty, _
                                                                                                "服务器", adVarChar, 500, Empty, "升级", adInteger, Empty, Empty, "当前版本", adVarChar, 20, Empty, _
                                                                                                "目标版本", adVarChar, 20, Empty, "中止信息", adVarChar, 2000, Empty, "可升级", adInteger, 1, 0, "检查结果", adVarChar, 2000, Empty, _
                                                                                                "提前目标版本", adVarChar, 20, Empty, "提前中止信息", adVarChar, 2000, Empty, "可提前升级", adInteger, 1, 0, "提前检查结果", adVarChar, 2000, Empty, _
                                                                                                "验证", adInteger, Empty, Empty))
    mrsSysInfo.Filter = "升级=1"
    mrsSysInfo.Sort = "系统编号"
    Do While Not mrsSysInfo.EOF
        rsSpaces.Filter = "系统=" & mrsSysInfo!系统编号
        rsSpaces.Sort = "当前,编号"
        Do While Not rsSpaces.EOF
            strServer = gstrServer
            If rsSpaces!DB连接 & "" <> "" Then
                strSQL = "Select Owner, Db_Link, Username, Host" & vbNewLine & _
                            "From All_Db_Links" & vbNewLine & _
                            "Where Owner =[1] And Username =[2] And Db_Link||'.' Like [3]"
                Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, gstrUserName, UCase(rsSpaces!所有者 & ""), UCase(rsSpaces!DB连接 & "") & ".%")
                If Not rsTmp.EOF Then strServer = rsTmp!Host & ""
            End If
            mrsHistorySpace.AddNew Array("ID", "系统编号", "系统名称", "系统版本", "目标版本", "配置文件", "编号", "名称", "当前", "所有者", "DB连接", "密码", "服务器", "升级", "可升级", "可提前升级", "验证"), _
                                                Array(Identity(lngID), mrsSysInfo!系统编号, mrsSysInfo!系统名称, mrsSysInfo!系统版本号, mrsSysInfo!目标版本, mrsSysInfo!配置文件, rsSpaces!编号, rsSpaces!名称, Val(rsSpaces!当前 & ""), Trim(UCase(rsSpaces!所有者 & "")), rsSpaces!DB连接, Null, UCase(strServer), 1, 1, 1, 0)
            rsSpaces.MoveNext
        Loop
        mrsSysInfo.MoveNext
    Loop
    Exit Sub
errh:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub ReadOptionalProc(Optional ByVal blnReadHis As Boolean)
'功能：读取可选过程
'参数：blnReadHis=是读取历史库的可选存储过程
    Dim arrTmp As Variant, strTmp As String
    Dim strName As String, strTip As String
    Dim lngID As Long, i As Long
    Dim strSysName As String, lngSysNo As Long
    
    On Error GoTo errh
    If mrsOptionalProc Is Nothing Or Not blnReadHis Then
        Set mrsOptionalProc = CopyNewRec(Nothing, True, , Array("ID", adInteger, Empty, Empty, "系统编号", adInteger, Empty, Empty, "系统名称", adVarChar, 50, Empty, "执行者", adVarChar, 100, Empty, "历史库", adInteger, Empty, Empty, "SPVer", adVarChar, 30, Empty, _
                                                                                                    "名称", adVarChar, 100, Empty, "FilePath", adVarChar, 2000, Empty, "注释", adLongVarChar, 2000, Empty, "执行", adInteger, Empty, Empty))
        If mblnExecBef Then Exit Sub '提前升迁，只初始化记录集即可
        mrsSysInfo.Filter = "升级=1"
        mrsSysInfo.Sort = "系统编号"
        Do While Not mrsSysInfo.EOF
            '当前系统的非历史库的可选脚本的过滤
            mrsSysFiles.Filter = "SysType<>" & ST_History & " And 系统编号=" & mrsSysInfo!系统编号 & " And FullSPVer<='" & VerFull(mrsSysInfo!目标版本) & "' And FileType=" & FT_Optional
            mrsSysFiles.Sort = "FullSPVer"
            Do While Not mrsSysFiles.EOF
                strTmp = mclsRunScript.CollectProcs(mrsSysFiles!FilePath)
                arrTmp = Split(strTmp, "?")
                For i = LBound(arrTmp) To UBound(arrTmp)
                    strName = Left(arrTmp(i), InStr(arrTmp(i), "|") - 1)
                    strTip = Mid(arrTmp(i), InStr(arrTmp(i), "|") + 1)
                    mrsOptionalProc.AddNew Array("ID", "系统编号", "系统名称", "执行者", "历史库", "SPVer", "名称", "FilePath", "注释", "执行"), _
                                                            Array(Identity(lngID), mrsSysInfo!系统编号, mrsSysInfo!系统名称, IIf(mrsSysInfo!系统编号 = 0, "ZLTOOLS", gstrUserName), 0, mrsSysFiles!SPVer, strName, mrsSysFiles!FilePath, RemoveMark(strTip), 1)
                Next
                mrsSysFiles.MoveNext
            Loop
            mrsSysInfo.MoveNext
        Loop
    ElseIf blnReadHis Then
        If mblnExecBef Then
             '清空服务器改变标志
            Call RecUpdate(mrsHistorySpace, "升级=-1", "升级", 1)
            Exit Sub '提前升迁，只初始化记录集即可
        End If
        '删除不能升迁的历史库、不选择升迁、以及改变服务器重新验证的历史库的升迁脚本
        mrsHistorySpace.Filter = "升级=0  OR 可升级=0 OR 验证<>2 OR 升级=-1 "
        Do While Not mrsHistorySpace.EOF '删除取消勾选的历史库的可选过程
            Call RecDelete(mrsOptionalProc, "系统编号=" & mrsHistorySpace!系统编号 & " And 执行者='" & UCase(mrsHistorySpace!名称 & "") & "'") '先删除历史库的可选存储过程
            mrsHistorySpace.MoveNext
        Loop
        '清空服务器改变标志
        Call RecUpdate(mrsHistorySpace, "升级=-1", "升级", 1)
        mrsOptionalProc.Filter = ""
        lngID = mrsOptionalProc.RecordCount
        mrsHistorySpace.Filter = "升级=1 And 可升级=1 And 验证=2" '增加勾选升级的历史库的可选过程
        Do While Not mrsHistorySpace.EOF
            mrsOptionalProc.Filter = "系统编号=" & mrsHistorySpace!系统编号 & " And 历史库=1 And 执行者='" & mrsHistorySpace!名称 & "'"
            If mrsOptionalProc.EOF Then '该历史库没有可选存储过程记录，则重新收集
                mrsSysFiles.Filter = "系统编号=" & mrsHistorySpace!系统编号 & " And SysType=" & ST_History & " And FileType=" & FT_Optional
                mrsSysFiles.Sort = "FullSPVer"
                Do While Not mrsSysFiles.EOF
                    strTmp = mclsRunScript.CollectProcs(mrsSysFiles!FilePath)
                    arrTmp = Split(strTmp, "?")
                    For i = LBound(arrTmp) To UBound(arrTmp)
                        strName = Left(arrTmp(i), InStr(arrTmp(i), "|") - 1)
                        strTip = Mid(arrTmp(i), InStr(arrTmp(i), "|") + 1)
                        mrsOptionalProc.AddNew Array("ID", "系统编号", "系统名称", "执行者", "历史库", "SPVer", "名称", "FilePath", "注释", "执行"), _
                                                                Array(Identity(lngID), mrsHistorySpace!系统编号, mrsHistorySpace!系统名称, mrsSysFiles!所有者, 1, mrsSysFiles!SPVer, strName, mrsSysFiles!FilePath, RemoveMark(strTip))
                    Next
                    mrsSysFiles.MoveNext
                Loop
            End If
            mrsHistorySpace.MoveNext
        Loop
        Call RefreshTotalInfo(2) '刷新可选过程汇总信息
    End If
    Exit Sub
errh:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshTotalInfo(Optional ByVal intRefreshType As Integer = -1)
'功能：刷新汇总信息
'参数：intRefreshType=刷新类型，-1：所有的汇总信息刷新, 0:刷新历史库, 1:刷新导入报表，2：刷新可选过程
    '历史库汇总信息刷新
    If intRefreshType = -1 Or intRefreshType = 0 Then
        mrsHistorySpace.Filter = ""
        If intRefreshType = -1 Then
            If mrsHistorySpace.RecordCount = 0 Then
                lblHisWarn.Visible = False: lblHisTotal.Visible = False: lblHisSel.Visible = False
                chkHisAll.value = 0
            End If
            Call SetCtrlEnabled(mrsHistorySpace.RecordCount <> 0, chkHisAll, lblHisPwd, txtHisPwd)
        End If
        lblHisTotal.Caption = "总数：" & mrsHistorySpace.RecordCount & "，选择："
        mrsHistorySpace.Filter = "升级=1"
        lblHisTotal.Caption = lblHisTotal.Caption & mrsHistorySpace.RecordCount
        mrsHistorySpace.Filter = "升级=1 And 验证<>2"
        lblHisWarn.Caption = mrsHistorySpace.RecordCount & "个历史库未通过验证！"
        lblHisWarn.Visible = mrsHistorySpace.RecordCount <> 0
        If lblHisWarn.Visible Then
            Call SetCtrlPosOnLine(False, 0, txtHisPwd, 60, lblHisWarn, 60, lblHisSel)
        Else
            Call SetCtrlPosOnLine(False, 0, txtHisPwd, 60, lblHisSel)
        End If
        Call RecToLog(mrsHistorySpace, "系统编号,编号", IIf(intRefreshType = -1, "原始历史库记录集", "历史库记录集刷新"))
    End If
    '导入报表汇总信息刷新
    If intRefreshType = -1 Or intRefreshType = 1 Then
        mrsReport.Filter = ""
        If intRefreshType = -1 Then
            If mrsReport.RecordCount = 0 Then
                lblRptSel.Visible = False: lblRptTotal.Visible = False
                chkRpt.value = 0: chkRpt.Enabled = False
            End If
            '导入报表相关控件可用性设置
            Call chkRpt_Click
        End If
        
        lblRptTotal.Caption = "总数：" & mrsReport.RecordCount & "，整体导入："
        mrsReport.Filter = "覆盖类型=1"
        lblRptTotal.Caption = lblRptTotal.Caption & mrsReport.RecordCount & "，只导入数据源："
        mrsReport.Filter = "覆盖类型=2"
        lblRptTotal.Caption = lblRptTotal.Caption & mrsReport.RecordCount
        Call RecToLog(mrsReport, "系统编号,编号", IIf(intRefreshType = -1, "原始导入报表记录集", "导入报表记录集刷新"))
    End If
    '可选过程汇总信息刷新
    If intRefreshType = -1 Or intRefreshType = 2 Then
        mrsOptionalProc.Filter = ""
        If intRefreshType = -1 Then
            If mrsOptionalProc.RecordCount = 0 Then
                lblOptSel.Visible = False: lblOptTotal.Visible = False
                chkOpt.value = 0: chkOpt.Enabled = False
            End If
            Call chkOpt_Click
        End If
        lblOptTotal.Caption = "总数：" & mrsOptionalProc.RecordCount & "，执行："
        mrsOptionalProc.Filter = "执行=1"
        lblOptTotal.Caption = lblOptTotal.Caption & mrsOptionalProc.RecordCount
        Call RecToLog(mrsOptionalProc, "系统编号,ID", IIf(intRefreshType = -1, "原始可选过程录集", "可选过程录集刷新"))
    End If
End Sub

Private Sub StepSwitch(ByVal intWay As Integer)
    Dim strPre As String, arrTmp As Variant
    Dim strOptProcs As String
    
    On Error GoTo errh
    If intWay = 1 Then
        If Not StepValidate(mintStep) Then Exit Sub
    End If
    If mintStep = SI_升迁设置 Then
        If MsgBox("系统升迁工作事关重大，请确认已经做好了各项准备工作。" & vbCrLf & vbCrLf & "要开始进行系统升迁吗？", _
                vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    mintStep = mintStep + intWay
    If mintStep = SI_系统升迁 Then
        '删除不需要升级的历史库脚本
        mrsHistorySpace.Filter = "升级=0"
        Do While Not mrsHistorySpace.EOF
            Call RecDelete(mrsSysFiles, "系统编号=" & mrsHistorySpace!系统编号 & " And 所有者='" & UCase(mrsHistorySpace!名称 & "") & "' And SysType=" & ST_History)
            mrsHistorySpace.MoveNext
        Loop
        '删除标准历史库脚本记录集
        Call RecDelete(mrsSysFiles, "所有者=Null And SysType=" & ST_History)
       '标记需要执行的可选过程
        If Not mblnExecBef Then
            mrsOptionalProc.Filter = "执行=1"
            mrsOptionalProc.Sort = "系统编号,SPVer,执行者,历史库"
            Do While Not mrsOptionalProc.EOF
                If strPre <> mrsOptionalProc!执行者 & "|" & mrsOptionalProc!SPVer & "|" & mrsOptionalProc!系统编号 & "|" & mrsOptionalProc!历史库 Then
                    If strPre <> "" Then
                        arrTmp = Split(strPre, "|")
                        Call RecUpdate(mrsSysFiles, "系统编号=" & arrTmp(2) & " And SPVer='" & arrTmp(1) & "' And FileType=" & FT_Optional & IIf(arrTmp(3) = 1, " And SysType=" & ST_History & " And 所有者='" & arrTmp(0) & "'", " And SysType<>" & ST_History), "Optional", IIf(strOptProcs = "", Null, Mid(strOptProcs, 2)))
                    End If
                    strPre = mrsOptionalProc!执行者 & "|" & mrsOptionalProc!SPVer & "|" & mrsOptionalProc!系统编号 & "|" & mrsOptionalProc!历史库
                    strOptProcs = ""
                End If
                strOptProcs = strOptProcs & "," & mrsOptionalProc!名称
                mrsOptionalProc.MoveNext
            Loop
            If strPre <> "" Then
                arrTmp = Split(strPre, "|")
                Call RecUpdate(mrsSysFiles, "系统编号=" & arrTmp(2) & " And SPVer='" & arrTmp(1) & "' And FileType=" & FT_Optional & IIf(arrTmp(3) = 1, " And SysType=" & ST_History & " And 所有者='" & arrTmp(0) & "'", " And SysType<>" & ST_History), "Optional", IIf(strOptProcs = "", Null, Mid(strOptProcs, 2)))
            End If
            '删除没有执行的可选脚本
            Call RecDelete(mrsSysFiles, "FileType=" & FT_Optional & " And Optional=Null")
        End If
    End If
    Call StepDisplay(mintStep)
    If mintStep = SI_系统升迁 Then
        '重新实例化，清除使用痕迹
        Set mclsRunScript = New clsRunScript
        '设置参数类参数
        Call mclsRunScript.InitGlobalPara(Me, 0, optErrOption(ET_忽略所有错误).value, _
                                                            lblLogModi.Tag, IIf(chkLogLong.value = 0, 0, Val(txtLogLong.Text)), chkStat.value = 1, mblnExecBef And ckhIdxOnLine.value = 1, optLogType(1).value)
        '初始化用户密码信息，加密块可能用到
        Call mclsRunScript.InitUserList(gstrUserName, gstrPassword, txtToolsPwd.Text, txtDBAUser.Text, txtDBAPwd.Text)
        mclsRunScript.Server = gstrServer
        '升迁日志记录升迁设置，以及升迁内容
        Call LogSetInfo
        Call UpgradeExecute
        On Error Resume Next
        Unload Me
    End If
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Private Sub LogSetInfo()
'功能：记录日志信息
    Dim strLog As String, strTmp As String
    Dim lngLen As Long
    Dim vsnRoot As VSFlexNode, vsTmp As VSFlexNode
    Dim i As Long
    
    On Error GoTo errh
    '升迁日志记录升迁设置，以及升迁内容
    lngLen = 16
    mclsRunScript.WriteSection "升迁概要信息"
    mclsRunScript.WriteLog "服务器时间：" & Format(CurrentDate, "yyyy-MM-dd HH:mm:ss") & String(4, " ") & "，本机时间：" & Format(Now, "yyyy-MM-dd HH:mm:ss")
    mclsRunScript.WriteLog "说明：为了减少与数据库服务器的交互，以下将使用本机时间作为记录日志的时间"
    mrsSysInfo.Filter = "系统编号=0" '管理工具
    mclsRunScript.WriteLog "Oracle Version：" & GetOracleVersion(False)
    mclsRunScript.WriteLog "服    务    器：" & gstrServer
    strTmp = IIf(mblnExecBef, mrsSysInfo!提前目标版本, mrsSysInfo!目标版本)
    mclsRunScript.WriteLog "管  理  工  具：" & mrsSysInfo!系统版本号 & IIf(strTmp <> "", "-->" & IIf(mblnExecBef, mrsSysInfo!提前目标版本, mrsSysInfo!目标版本), "")
    mrsSysInfo.Filter = "系统编号<>0 and 升级=1"
    mrsSysInfo.Sort = "Sort,系统编号"
    Do While Not mrsSysInfo.EOF
        strTmp = IIf(mblnExecBef, mrsSysInfo!提前目标版本, mrsSysInfo!目标版本)
        mclsRunScript.WriteLog mrsSysInfo!系统编号 & "-" & mrsSysInfo!系统名称 & "：" & mrsSysInfo!系统版本号 & IIf(strTmp <> "", "-->" & IIf(mblnExecBef, mrsSysInfo!提前目标版本, mrsSysInfo!目标版本), "")
        mrsSysInfo.MoveNext
    Loop
    mclsRunScript.WriteSection "升迁设置"
    '参数设置日志
    mclsRunScript.WriteLog "升迁参数"
    If chkParallel.value = 0 Or chkParallel.Enabled = False Then
        mintDDLParallel = 0
        mclsRunScript.WriteLog "  不采用并行DDL"
    Else
        mintDDLParallel = Val(txtCpu(0).Text)
        mclsRunScript.WriteLog "  采用并行DDL 并行度=" & Val(txtCpu(0).Text)
    End If
    If Not ckhIdxOnLine.Visible Or ckhIdxOnLine.value = 0 Then
        mclsRunScript.WriteLog "  不采用在线模式创建索引"
    Else
        mclsRunScript.WriteLog "  采用在线模式创建索引"
    End If
    If chkStat.value = 0 Then
        mclsRunScript.WriteLog "  未启用统计信息收集"
    Else
        mclsRunScript.WriteLog "  启用统计信息收集 方式：" & IIf(optStatType(0).value, "升级中直接收集", "只记录待收集对象")
        mclsRunScript.WriteLog "    统计信息收集并行度=" & Val(txtCpu(1).Text)
        mclsRunScript.WriteLog "    统计信息收集采样百分比" & IIf(chkAutoPer.value = 0, "=" & Val(txtEsper.Text), "  采用自动百分比")
    End If
    mclsRunScript.WriteLog "  日志记录方式采取" & IIf(optLogType(1).value, "只记录出错SQL", "全部记录")
    If chkLogLong.value = 0 Then
        mclsRunScript.WriteLog "  日志不记录长时执行SQL"
    Else
        mclsRunScript.WriteLog "  日志记录执行超过" & Val(txtLogLong.Text) & "分钟的SQL语句"
    End If
    mclsRunScript.WriteLog "  错误处理方式采取" & IIf(optErrOption(ET_忽略次要错误).value, "忽略次要错误", "忽略所有错误")
    '历史库选择日志
    mrsHistorySpace.Filter = ""
    mrsHistorySpace.Sort = "系统编号,当前,编号"
    If mrsHistorySpace.RecordCount <> 0 Then
        mclsRunScript.WriteLog String(80, "-")
        mclsRunScript.WriteLog "历史空间升迁"
        Do While Not mrsHistorySpace.EOF
            strLog = "    " & Lpad(mrsHistorySpace!系统编号, 4) & "-" & RPAD(mrsHistorySpace!系统名称, 16)
            strLog = strLog & "  " & RPAD(mrsHistorySpace!名称, 14) & "  " & RPAD(IIf(mrsHistorySpace!当前 = 1, "当前", "非当前"), 5)
            strLog = strLog & "  " & IIf(mrsHistorySpace!升级 = 1, "升级", "不升级")
            mclsRunScript.WriteLog strLog
            mrsHistorySpace.MoveNext
        Loop
    End If
    '可选过程日志
    mrsOptionalProc.Filter = ""
    mrsOptionalProc.Sort = "系统编号,历史库,ID"
    If mrsOptionalProc.RecordCount <> 0 Then
        mclsRunScript.WriteLog String(80, "-")
        mclsRunScript.WriteLog "执行可选过程"
        Do While Not mrsOptionalProc.EOF
            strLog = "    " & Lpad(mrsOptionalProc!系统编号, 4) & "-" & RPAD(mrsOptionalProc!系统名称, 16)
            strLog = strLog & "  " & RPAD(mrsOptionalProc!名称, 32) & "  " & RPAD(mrsOptionalProc!执行者, lngLen - 2)
            strLog = strLog & "  " & RPAD(IIf(mrsOptionalProc!历史库 = 1, "历史库", "非历史库"), 6) & "  " & RPAD(IIf(mrsOptionalProc!执行 = 1, "执行", "不执行"), 6)
            strLog = strLog & "  " & mrsOptionalProc!FilePath
            mclsRunScript.WriteLog strLog
            mrsOptionalProc.MoveNext
        Loop
    End If
    '导入报表日志
    mrsReport.Filter = ""
    mrsReport.Sort = "系统编号,ID"
    If mrsReport.RecordCount <> 0 Then
        mclsRunScript.WriteLog String(80, "-")
        mclsRunScript.WriteLog "导入报表"
        Do While Not mrsReport.EOF
            strLog = "    " & Lpad(mrsReport!系统编号, 4) & "-" & RPAD(mrsReport!系统名称, lngLen)
            strLog = strLog & "  " & RPAD(mrsReport!编号, 20) & "  " & RPAD(mrsReport!名称, 30)
            strLog = strLog & "  " & RPAD(Decode(mrsReport!覆盖类型, 0, "不导入", 1, "整体导入", 2, "数据源导入"), 10)
            strLog = strLog & "  " & mrsReport!FilePath
            mclsRunScript.WriteLog strLog
            mrsReport.MoveNext
        Loop
    End If
    mclsRunScript.WriteSection "升迁流程"
    For i = vsPlan.FixedRows + 1 To vsPlan.Rows - IIf(mblnExecBef, 1, 2)
        Set vsTmp = vsPlan.GetNode(i)
        mclsRunScript.WriteLog vsTmp.Text
        vsTmp.Expanded = False
    Next
    
    mclsRunScript.WriteLog
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Private Function StepValidate(ByVal intStep As IDX_STEP) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim cnTmp As New ADODB.Connection
    Dim strMsg As String
    Dim strErr As String
    Dim blnDo As Boolean
    
    On Error GoTo errh
    If intStep = SI_升迁设置 Then
        If txtToolsPwd.Enabled And txtToolsPwd.Text = "" Then
            MsgBox "请输入管理工具用户的密码。", vbInformation, gstrSysName
            txtToolsPwd.SetFocus: Exit Function
        End If
        If txtDBAUser.Enabled And txtDBAUser.Text = "" Then
            MsgBox "请输入具有DBA身份的用户名。", vbInformation, gstrSysName
            txtDBAUser.SetFocus: Exit Function
        End If
        If txtDBAPwd.Enabled And txtDBAPwd.Text = "" Then
            MsgBox "请输入DBA用户的密码。", vbInformation, gstrSysName
            txtDBAPwd.SetFocus: Exit Function
        End If
        If txtToolsPwd.Enabled Then
            '管理工具密码验证
            If gstrToolsPwd <> "" And UCase(gstrToolsPwd) <> UCase(Trim(txtToolsPwd.Text)) Then
                 MsgBox "管理工具密码错误！", vbInformation, gstrSysName
                 Exit Function
            End If
            err.Clear
            
            If gcnTools Is Nothing Then
                blnDo = True
            ElseIf gcnTools.State = adStateClosed Then
                blnDo = True
            End If
            
            If blnDo Then
                Set gcnTools = gobjRegister.GetConnection(gstrServer, "zltools", txtToolsPwd.Text, False, MSODBC, "", False)
                If gcnTools.State = adStateClosed Then
                    MsgBox "连接管理工具用户时出现错误：" & vbCrLf & vbCrLf & strErr, vbInformation, gstrSysName
                    Exit Function
                End If
                Call SetSQLTrace(gstrServer, "zltools", gcnTools)
                gstrToolsPwd = txtToolsPwd.Text '赋值
            End If
        End If
        If txtDBAPwd.Enabled Then
            'DBA用户密码验证
            If UCase(txtDBAUser.Text) = UCase(gstrSysUser) And UCase(gstrSysPwd) <> UCase(txtDBAPwd.Text) And gstrSysPwd <> "" Then
                MsgBox "DBA用户密码错误！", vbInformation, gstrSysName
                Exit Function
            End If
            If gcnSystem Is Nothing Then
                blnDo = True
            ElseIf gcnSystem.State = adStateClosed Then
                blnDo = True
            End If
            
            If blnDo Then
                Set cnTmp = gobjRegister.GetConnection(gstrServer, txtDBAUser.Text, txtDBAPwd.Text, False, MSODBC, "", False)
                If cnTmp.State = adStateClosed Then
                    MsgBox "连接DBA用户时出现错误.", vbInformation, gstrSysName
                    Exit Function
                End If
                On Error GoTo 0
                '检查是否DBA
                strSQL = "Select 1 From Session_Roles Where Role = 'DBA'"
                Set rsTmp = gclsBase.OpenSQLRecord(cnTmp, strSQL, "是否DBA")
                If rsTmp.EOF Then
                    MsgBox "该用户不具有DBA身份！", vbExclamation, gstrSysName
                    txtDBAUser.SetFocus: Exit Function
                End If
                
                Call SetSQLTrace(gstrServer, txtDBAUser.Text, cnTmp)
                Set gcnSystem = cnTmp
                gstrSysUser = txtDBAUser.Text
                gstrSysPwd = txtDBAPwd.Text
            Else
                Call SetSQLTrace(gstrServer, gstrSysUser, gcnSystem)
            End If
        End If
        '必须输入升迁日志
        If lblLog.Caption = "升迁日志文件：" Then
            MsgBox "请确定升迁日志文件的存放位置和名字。", vbInformation, gstrSysName
            Exit Function
        End If
        '当前历史库必须升级，没有注册的历史库则不检查，检查没有验证密码或者验证密码，没有选择升级的历史库
        Call RecUpdate(mrsHistorySpace, "当前=1 And 升级=0  And 验证<>1", "升级", 1)
        Call RecUpdate(mrsHistorySpace, "验证=1", "升级", 0)
        mrsHistorySpace.Filter = "当前=1 And 验证=0 And 升级=1"
        mrsHistorySpace.Sort = "系统编号,ID": strMsg = ""
        Do While Not mrsHistorySpace.EOF
            strMsg = strMsg & vbNewLine & "【" & mrsHistorySpace!系统名称 & "】的表空间-" & mrsHistorySpace!名称
            mrsHistorySpace.MoveNext
        Loop
        If strMsg <> "" Then
            MsgBox "以下系统的当前历史表空间必须升级：" & strMsg & "！请进行验证。", vbInformation, gstrSysName
            '重新读取可选脚本
            Call ReadOptionalProc(True)
            '刷新历史库汇总信息
            Call RefreshTotalInfo(0)
            Exit Function
        End If
        mrsHistorySpace.Filter = "升级=1 And 验证=0"
        mrsHistorySpace.Sort = "系统编号,ID": strMsg = ""
        Do While Not mrsHistorySpace.EOF
            strMsg = strMsg & vbNewLine & "【" & mrsHistorySpace!系统名称 & "】的表空间-" & mrsHistorySpace!名称 & "，"
            mrsHistorySpace.MoveNext
        Loop
        If strMsg <> "" Then
            If MsgBox("以下历史表空间未通过验证不能升级：" & strMsg & vbNewLine & "是否继续？", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                '重新读取可选脚本
                Call ReadOptionalProc(True)
                '刷新历史库汇总信息
                Call RefreshTotalInfo(0)
                Exit Function
            End If
            '将没有通过验证的历史库取消升级
            Call RecUpdate(mrsHistorySpace, "升级=1 And 验证<>2 ", "升级", 0)
        End If
        '将通过验证且不能升级的历史库取消升级
        Call RecUpdate(mrsHistorySpace, "升级=1 And 验证=2 " & IIf(mblnExecBef, "  And 可提前升级=0", " And 可升级=0"), "升级", 0)
    End If
    StepValidate = True
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub StepDisplay(ByVal intStep As IDX_STEP)
    Dim i As Integer
    Dim arrTmp As Variant
    Dim strTmp As String, strMaxVer As String
    Dim vsnRoot As VSFlexNode, vsnTop As VSFlexNode, vsnSecd As VSFlexNode, vsnChild As VSFlexNode
    Dim vsnAPP As VSFlexNode, vsnHis As VSFlexNode, vsnRpt As VSFlexNode, vsnCompile As VSFlexNode
    Dim vsnCHCEK As VSFlexNode, vsnTools As VSFlexNode, vsnAdjustSeq As VSFlexNode
    
    mblnHaveST = False
    arrTmp = Split(Split(STEP_INFO, "||")(intStep), "|")
    For i = 0 To fraStep.UBound
        fraStep(i).Visible = i = intStep
    Next
    cmdCancel.Enabled = intStep < SI_系统升迁
    If intStep = SI_系统升迁 Then
        Call SetSQLState(True, True)
        With vsPlan
            '注意：关键字各个单次以^分割不用下划线，主要是由于，历史库以及用户等，可以输入下划线
            .Rows = .FixedRows: .Rows = .FixedRows + 1: .IsSubtotal(.Rows - 1) = True
            '添加一个根节点，方便添加子节点
            Set vsnRoot = .GetNode(.Rows - 1): vsnRoot.Text = "系统升迁": vsnRoot.Key = "Main": Set vsnRoot.Image = imgPlan.ListImages("Doing").Picture: vsnRoot.Expanded = True
             .Rows = .Rows + 1: .IsSubtotal(.Rows - 1) = True
            If Not mblnExecBef Then
                Set vsnTop = .GetNode(.Rows - 1): vsnTop.Text = "客户端站点部件升级": vsnTop.Key = "Client": Set vsnTop.Image = imgPlan.ListImages("Wait").Picture: vsnTop.Expanded = True
                Set vsnCHCEK = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "升迁检查", FS_升迁检查, imgPlan.ListImages("Wait").Picture)
                
            End If
            If txtToolsPwd.Enabled Then
                Set vsnTools = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "管理工具" & IIf(mblnExecBef, "提前", "") & "升迁", FS_工具升迁, imgPlan.ListImages("Wait").Picture)
                 If Not mblnExecBef Then
                    Call vsnCHCEK.AddNode(flexNTLastChild, GetCode(vsnCHCEK.Text) & "." & (vsnCHCEK.Children + 1) & "管理工具", vsnCHCEK.Key & "^TOOLS", imgPlan.ListImages("Wait").Picture)
                End If
                'PLJSON安装流程,提前升级没有该流程
                If mblnInstallPLJson Then
                    Call vsnTools.AddNode(flexNTLastChild, "PLJSON安装", vsnTools.Key & "^PLJSON", imgPlan.ListImages("Wait").Picture)
                End If
                mrsSysFiles.Filter = "系统编号=0": mrsSysFiles.Sort = "FullSPVer"
                strTmp = ""
                Do While Not mrsSysFiles.EOF
                    If InStr(strTmp & ",", "," & mrsSysFiles!SPVer & ",") = 0 Then
                        strTmp = strTmp & "," & mrsSysFiles!SPVer
                        '添加管理工具升迁到某一个版本
                        Call vsnTools.AddNode(flexNTLastChild, mrsSysFiles!SPVer, vsnTools.Key & "^" & mrsSysFiles!SPVer, imgPlan.ListImages("Wait").Picture)
                    End If
                    mrsSysFiles.MoveNext
                Loop
                Call vsnTools.AddNode(flexNTLastChild, "对象授权修正", vsnTools.Key & "^PUBGRANT", imgPlan.ListImages("Wait").Picture)
            End If
            '系统升迁处理
            mrsSysInfo.Filter = "系统编号<>0 And 升级=1": mrsSysInfo.Sort = "Sort"
            If Not mrsSysInfo.EOF Then
                Set vsnAPP = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "应用系统" & IIf(mblnExecBef, "提前", "") & "升迁", FS_应用系统升迁, imgPlan.ListImages("Wait").Picture)
                mrsHistorySpace.Filter = IIf(mblnExecBef, "升级=1", "")
                If mblnExecBef And Not mrsHistorySpace.EOF Or Not mblnExecBef Then
                    Set vsnHis = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "历史表空间" & IIf(mblnExecBef, "提前", "") & "升迁", FS_历史库升迁, imgPlan.ListImages("Wait").Picture)
                End If
                Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & ".公共同义词创建", FS_公共同义词, imgPlan.ListImages("Wait").Picture)
                If Not mblnExecBef Then
                    Set vsnCompile = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "编译无效对象", FS_编译无效对象, imgPlan.ListImages("Wait").Picture)
                    Set vsnAdjustSeq = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "重新调整序列", FS_重整序列, imgPlan.ListImages("Wait").Picture)
                End If
                mrsReport.Filter = "覆盖类型<>0"
                If Not mrsReport.EOF Then
                    Set vsnRpt = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "报表导入升级", FS_报表升级, imgPlan.ListImages("Wait").Picture)
                End If
                If mblnFinal Then Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "角色重新授权", FS_角色授权, imgPlan.ListImages("Wait").Picture)
            ElseIf Not mblnExecBef Then
               Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & ".公共同义词创建", FS_公共同义词, imgPlan.ListImages("Wait").Picture)
                Set vsnCompile = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "编译无效对象", FS_编译无效对象, imgPlan.ListImages("Wait").Picture)
                Set vsnAdjustSeq = vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "重新调整序列", FS_重整序列, imgPlan.ListImages("Wait").Picture)
                If mblnFinal Then Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "角色重新授权", FS_角色授权, imgPlan.ListImages("Wait").Picture)
            Else
                Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & ".公共同义词创建", FS_公共同义词, imgPlan.ListImages("Wait").Picture)
            End If
            If chkStat.value = 1 Then
                Call vsnRoot.AddNode(flexNTLastChild, (vsnRoot.Children + 1) & "." & "统计信息收集", FS_统计信息收集, imgPlan.ListImages("Wait").Picture)
            End If
            
            mstrSysCodes = ""
            Do While Not mrsSysInfo.EOF
                If mrsSysInfo!系统编号 \ 100 = 1 Then mblnHaveST = True
                mstrSysCodes = mstrSysCodes & IIf(mstrSysCodes = "", "", ",") & mrsSysInfo!系统编号
                '升迁检查流程增加
                 If Not mblnExecBef Then Call vsnCHCEK.AddNode(flexNTLastChild, GetCode(vsnCHCEK.Text) & "." & (vsnCHCEK.Children + 1) & mrsSysInfo!系统名称, vsnCHCEK.Key & "^" & mrsSysInfo!系统编号, imgPlan.ListImages("Wait").Picture)
                '应用系统升迁流程增加
                mrsSysFiles.Filter = "系统编号=" & mrsSysInfo!系统编号 & " And SysType<>" & ST_History: mrsSysFiles.Sort = "FullSPVer"
               Set vsnTop = vsnAPP.AddNode(flexNTLastChild, GetCode(vsnAPP.Text) & "." & (vsnAPP.Children + 1) & "." & mrsSysInfo!系统名称, vsnAPP.Key & "^" & mrsSysInfo!系统编号, imgPlan.ListImages("Wait").Picture)
                strTmp = ""
                Do While Not mrsSysFiles.EOF
                    If InStr(strTmp & ",", "," & mrsSysFiles!SPVer & ",") = 0 Then
                        strTmp = strTmp & "," & mrsSysFiles!SPVer
                        Call vsnTop.AddNode(flexNTLastChild, mrsSysFiles!SPVer, vsnTop.Key & "^" & mrsSysFiles!SPVer, imgPlan.ListImages("Wait").Picture)
                    End If
                    mrsSysFiles.MoveNext
                Loop
                Call vsnTop.AddNode(flexNTLastChild, "H表访问权限修正", vsnTop.Key & "^HTABLEREPAIR", imgPlan.ListImages("Wait").Picture)
                '历史库升迁流程增加
                mrsHistorySpace.Filter = "升级=1 And 系统编号=" & mrsSysInfo!系统编号: mrsHistorySpace.Sort = "当前 Desc,编号"
                If Not mrsHistorySpace.EOF Then
                    '添加历史库所属系统
                    Set vsnTop = vsnHis.AddNode(flexNTLastChild, GetCode(vsnHis.Text) & "." & (vsnHis.Children + 1) & "." & mrsSysInfo!系统名称, vsnHis.Key & "^" & mrsSysInfo!系统编号, imgPlan.ListImages("Wait").Picture)
                    Do While Not mrsHistorySpace.EOF
                        '添加某个系统历史库
                        Set vsnSecd = vsnTop.AddNode(flexNTLastChild, mrsHistorySpace!名称, vsnTop.Key & "^" & mrsHistorySpace!名称, imgPlan.ListImages("Wait").Picture)
                        mrsSysFiles.Filter = "所有者='" & UCase(mrsHistorySpace!名称 & "") & "' And 系统编号=" & mrsSysInfo!系统编号 & " And SysType=" & ST_History: mrsSysFiles.Sort = "FullSPVer"
                        strTmp = "": strMaxVer = ""
                        '添加某个系统历史库升迁流程
                        Do While Not mrsSysFiles.EOF
                            If InStr(strTmp & ",", "," & mrsSysFiles!SPVer & ",") = 0 Then
                                strTmp = strTmp & "," & mrsSysFiles!SPVer
                                Call vsnSecd.AddNode(flexNTLastChild, mrsSysFiles!SPVer, vsnSecd.Key & "^" & mrsSysFiles!SPVer, imgPlan.ListImages("Wait").Picture)
                                strMaxVer = mrsSysFiles!SPVer & ""
                            End If
                            mrsSysFiles.MoveNext
                        Loop
                        If strMaxVer = "" Then strMaxVer = mrsHistorySpace!当前版本
                        '非提前执行，如果脚本不支持到目标版本，则则自动修正目标版本
                        If VerFull(strMaxVer) < VerFull(mrsHistorySpace!目标版本) And Not mblnExecBef Then
                            Call vsnSecd.AddNode(flexNTLastChild, mrsHistorySpace!目标版本, vsnSecd.Key & "^" & mrsHistorySpace!目标版本, imgPlan.ListImages("Wait").Picture)
                        End If
                        Call vsnSecd.AddNode(flexNTLastChild, "历史库结构修正", vsnSecd.Key & "^HISREPAIR", imgPlan.ListImages("Wait").Picture)
                        mrsHistorySpace.MoveNext
                    Loop
                ElseIf Not mblnExecBef Then   '没有历史库，则需要验证
                    Set vsnTop = vsnHis.AddNode(flexNTLastChild, GetCode(vsnHis.Text) & "." & (vsnHis.Children + 1) & "." & mrsSysInfo!系统名称 & "历史库检查", vsnHis.Key & "^" & mrsSysInfo!系统编号, imgPlan.ListImages("Wait").Picture)
                End If
                '报表导入添加
                mrsReport.Filter = "覆盖类型<>0 And 系统编号=" & mrsSysInfo!系统编号
                If Not mrsReport.EOF Then
                    Call vsnRpt.AddNode(flexNTLastChild, GetCode(vsnRpt.Text) & "." & (vsnRpt.Children + 1) & "." & mrsSysInfo!系统名称, vsnRpt.Key & "^" & mrsSysInfo!系统编号, imgPlan.ListImages("Wait").Picture)
                End If
                mrsSysInfo.MoveNext
            Loop
            '编译无效对象流程增加
            If Not vsnCompile Is Nothing Then
                If Not vsnTools Is Nothing Then
                    Call vsnCompile.AddNode(flexNTLastChild, GetCode(vsnCompile.Text) & "." & (vsnCompile.Children + 1) & ".管理工具", vsnCompile.Key & "^TOOLS", imgPlan.ListImages("Wait").Picture)
                End If
                If Not vsnAPP Is Nothing Then
                    Call vsnCompile.AddNode(flexNTLastChild, GetCode(vsnCompile.Text) & "." & (vsnCompile.Children + 1) & ".应用系统", vsnCompile.Key & "^APP", imgPlan.ListImages("Wait").Picture)
                End If
            End If
            '重整序列流程增加
            If Not vsnAdjustSeq Is Nothing Then
                If Not vsnTools Is Nothing Then
                    Call vsnAdjustSeq.AddNode(flexNTLastChild, GetCode(vsnAdjustSeq.Text) & "." & (vsnAdjustSeq.Children + 1) & ".管理工具", vsnAdjustSeq.Key & "^TOOLS", imgPlan.ListImages("Wait").Picture)
                End If
                If Not vsnAPP Is Nothing Then
                    Call vsnAdjustSeq.AddNode(flexNTLastChild, GetCode(vsnAdjustSeq.Text) & "." & (vsnAdjustSeq.Children + 1) & ".应用系统", vsnAdjustSeq.Key & "^APP", imgPlan.ListImages("Wait").Picture)
                End If
            End If
        End With
        txtSQL.SetFocus: Me.Refresh
    End If
    Set imgInfo.Picture = imgStep.ListImages(intStep + 1).Picture
    lblStep.Caption = arrTmp(0)
    lblInfo.Caption = arrTmp(1)
    cmdNext.Enabled = intStep + 1 <= fraStep.UBound
    cmdNext.Visible = cmdNext.Enabled
End Sub

Private Sub UpgradeExecute()
'功能：根据向导的设置，进行系统升迁
    Dim vsnStep As VSFlexNode, vsnTmp As VSFlexNode
    Dim arrTmp As Variant
    Dim strMsg As String
    Dim i As Long
    Dim blnFirstUp As Boolean
    Dim datStart As Date, lngSec As Long
    Dim cnTmp As ADODB.Connection, lngCount As Long
    Dim strPreVer As String
    
    tmrRefresh.Enabled = True
    On Error GoTo errh
    mstrChangeTables = ""
    Call UpdateSysFiles '记录本次升迁系统的配置文件
    For i = vsPlan.FixedRows To vsPlan.Rows - 2
        Call vsPlan.ShowCell(i, 0)
        Set vsnStep = vsPlan.GetNode(i)
        If vsnStep.Children = 0 Then  '可以执行的步骤
            arrTmp = Split(vsnStep.Key, "^")
            If UBound(arrTmp) = 0 Then
                Call SetSQLState(False) '关闭SQL
                mclsRunScript.WriteSection vsnStep.Text, IIf(i = vsPlan.FixedRows, "=", "-")
            Else
                mclsRunScript.WriteLog "[" & vsnStep.Text & "]"
            End If
            datStart = Now
            Call SetStepStateImg(vsnStep)  '开始执行
            Select Case arrTmp(0)
                Case FS_升迁检查
                    If Not UpgradeCheck(Val(arrTmp(1))) Then GoTo AbortLine
                Case FS_工具升迁
                    If arrTmp(1) = "PUBGRANT" Then
                        If Not mblnExecBef Then '将提前执行修改为0 ,表明提前执行已经不处于中间状态1
                            gcnOracle.Execute "Update zlUpGrade Set 提前执行=0 Where 提前执行 = 1 And 系统 is Null "
                        End If
                        mclsRunScript.SysNo = 0
                        Call ReGrantForTools(gcnTools, , True)
                    ElseIf arrTmp(1) = "PLJSON" Then
                        Call InstallPLJSON(gcnSystem, mstrToolsFloder, mclsRunScript, mblnJSONRemain)
                    Else
                        If blnFirstUp Then mrsSysInfo.Filter = "系统编号=0": mclsRunScript.SysNo = 0: strPreVer = ""
                        Call SetSQLState(True, True)
                        If Not RunScriptByVersion(0, arrTmp(1), blnFirstUp, IIf(strPreVer = "", mrsSysInfo!系统版本号, strPreVer), IIf(mblnExecBef, mrsSysInfo!提前目标版本, mrsSysInfo!目标版本)) Then Exit Sub
                        strPreVer = arrTmp(1)
                    End If
                Case FS_应用系统升迁
                    If arrTmp(2) = "HTABLEREPAIR" Then
                        If Not mblnExecBef Then '将提前执行修改为0 ,表明提前执行已经不处于中间状态1
                            gcnOracle.Execute "Update zlUpGrade Set 提前执行=0 Where 提前执行 = 1 And 系统 =" & Val(arrTmp(1))
                        End If
                        mclsRunScript.SysNo = Val(arrTmp(1))
                        Call HTablePrivsRepair(Val(arrTmp(1)))
                    Else
                        If blnFirstUp Then mrsSysInfo.Filter = "系统编号=" & arrTmp(1): mclsRunScript.SysNo = Val(arrTmp(1)): strPreVer = ""
                        Call SetSQLState(True, True)
                        If Not RunScriptByVersion(Val(arrTmp(1)), arrTmp(2), blnFirstUp, IIf(strPreVer = "", mrsSysInfo!系统版本号, strPreVer), IIf(mblnExecBef, mrsSysInfo!提前目标版本, mrsSysInfo!目标版本)) Then Exit Sub
                        strPreVer = arrTmp(1)
                    End If
                Case FS_历史库升迁
                    If UBound(arrTmp) = 3 Then '历史库升迁修正
                        If blnFirstUp Then
                            mrsHistorySpace.Filter = "系统编号=" & arrTmp(1) & " And 名称='" & arrTmp(2) & "'"
                            mclsRunScript.SysNo = Val(arrTmp(1))
                            Set cnTmp = gobjRegister.GetConnection(mrsHistorySpace!服务器, mrsHistorySpace!所有者, mrsHistorySpace!密码, False, MSODBC, "", False)
                            If cnTmp.State = adStateClosed Then
                               Set cnTmp = Nothing
                            Else
                               Call SetSQLTrace(mrsHistorySpace!服务器, mrsHistorySpace!所有者, cnTmp)
                            End If
                            strPreVer = ""
                        End If
                        If Not cnTmp Is Nothing Then
                            If arrTmp(3) = "HISREPAIR" Then
                                If Not mblnExecBef Then '将提前执行修改为0 ,表明提前执行已经不处于中间状态1
                                    cnTmp.Execute "Update zlbakinfo Set 中止语句=NULL,提前中止语句=NULL,提前执行=0  Where 系统=" & Val(arrTmp(1))
                                End If
                                Call RepairHisDB(cnTmp, Val(arrTmp(1)), mrsHistorySpace!所有者, mrsHistorySpace!服务器, mrsHistorySpace!名称, mrsHistorySpace!DB连接 & "", mrsHistorySpace!当前 = 1)
                            Else
                                Call RunScriptByVersion(Val(arrTmp(1)), arrTmp(3), blnFirstUp, IIf(strPreVer = "", mrsHistorySpace!当前版本, strPreVer), IIf(mblnExecBef, mrsHistorySpace!提前目标版本, mrsHistorySpace!目标版本), True, cnTmp, arrTmp(2))
                                strPreVer = arrTmp(1)
                            End If
                        End If
                    ElseIf UBound(arrTmp) = 1 Then '没有历史库
                        lngCount = 0
                        If CheckHavHistory(Val(arrTmp(1))) Then
ReDo:
                            lngCount = lngCount + 1
                            MsgBox "由于该系统存在历史数据空间表，但未设置相应的历史数据空间，你必需创建该空间!", vbInformation + vbDefaultButton1, gstrSysName
                            If frmHistorySpaceSet.ShowInstall(Me, gcnOracle, gstrUserName, gstrPassword, Val(arrTmp(1)), 0, 0, , True) = False Then
                                If lngCount < 3 Then
                                    GoTo ReDo:
                                Else
                                    MsgBox "由于你未创历史数据空间,因此,可能系统运行不正常,请随后在[数据管理-->数据转移]中处理!", vbInformation + vbDefaultButton1, gstrSysName
                                End If
                            End If
                        End If
                    End If
                Case FS_公共同义词
                    '为升级新增的对象创建公共同义词('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION')
                    gcnOracle.Execute "Zl_Createpubsynonyms", , adCmdStoredProc
                Case FS_编译无效对象
                    Call ReCompileObjects(IIf(arrTmp(1) = "TOOLS", gcnTools, gcnOracle))
                Case FS_重整序列
                    Call ReAdjustSequence(IIf(arrTmp(1) = "TOOLS", gcnTools, gcnOracle))
                Case FS_报表升级
                    Call ImportReports(Val(arrTmp(1)))
                Case FS_角色授权
                    Call GrantToRole
                Case FS_统计信息收集
                    Call GatherStatistics
            End Select
            
            mclsRunScript.WriteLog
            lngSec = DateDiff("s", datStart, Now)
            mclsRunScript.WriteLog "[" & vsnStep.Text & "]：从" & _
                                        Format(datStart, "HH:mm:ss") & "到" & Format(Now, "HH:mm:ss") & _
                                        "，共耗时" & IIf(lngSec > 60, (lngSec \ 60) & "分钟" & (lngSec Mod 60) & "秒", lngSec & "秒")

            mclsRunScript.WriteLog
            
            If blnFirstUp Then blnFirstUp = False
            Call SetStepStateImg(vsnStep, True)  '开始执行
        Else
            Call SetSQLState(False)
            blnFirstUp = True
            mclsRunScript.WriteSection vsnStep.Text, IIf(i = vsPlan.FixedRows, "=", "-")
            vsnStep.Expanded = True
        End If
        Me.Refresh
    Next
    Call UpgradeFinish(True)
    mblnOk = True
    If Not vsnStep Is Nothing Then Call SetStepStateImg(vsnStep, True)  '开始执行
    '部件升级
    If Not mblnExecBef Then
        Set vsnStep = vsPlan.GetNode(vsPlan.Rows - 1)
        Call SetStepStateImg(vsnStep)  '开始执行
        Call SetStepStateImg(vsnStep, True)  '开始执行
        If MsgBox("数据升迁完成后，需要对客户端站点进行部件升级," & vbCrLf & "需要对站点部件升级进行配置吗?", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
            Unload Me
            mstrRunModule = "0307"
        End If
        
'        '延迟执行的脚本
'        blnNormal = True
'        mrsSQLSys.Filter = "SysType=" & ST_App & " And FileType=" & FT_DefUp
'        mrsSQLSys.Sort = "FullSPVer"
'        mclsRunScript.ConnectType = 0: mclsRunScript.IsGather = False
'        If mrsSQLSys.RecordCount > 0 Then
'            blnNormal = True
'            If Mid(mrsSQLSys!SPVer, 1, 5) = "10.25" Then
'                MsgBox "数据升迁完成后，接下来将运行延迟脚本，在此期间系统可正常使用，建议你运行报表调整工具(ZLRPTSQLAdjust)调整数据源中涉及[病人费用记录]表的SQL语句。", vbInformation, gstrSysName
'            Else
'                MsgBox "数据升迁完成后，接下来将运行延迟脚本，在此期间系统可正常使用。", vbInformation, gstrSysName
'            End If
'            Set mclsRunScript.Connection = gcnOracle
'            Do While Not mrsSQLSys.EOF
'                Call RunSQLScript(mrsSQLSys!FilePath, , , False)
'                mrsSQLSys.MoveNext
'            Loop
'        End If
'
'        mrsSQLSys.Filter = "SysType=" & ST_AppHis & " And FileType=" & FT_DefUp
'        mrsSQLSys.Sort = "UserServer,UserName,FullSPVer"
'        If mrsSQLSys.RecordCount > 0 Then
'            Do While Not mrsSQLSys.EOF
'                If strPreBakUserName <> mrsSQLSys!UserName Or Not blnConn Then
'                    strPreBakUserName = mrsSQLSys!UserName
'                    blnConn = True '是否打开连接成功
'                    If OpenHistoryConnect(Nvl(mrsSQLSys!UserName), Nvl(mrsSQLSys!UserPass), Nvl(mrsSQLSys!UserServer), True) = False Then
'                        '一般此种情况不存在.因为在之间已经检查，这里保证连接是当前历史库的连接
'                        blnConn = False
'                    End If
'                End If
'                If blnConn Then
'                    Set mclsRunScript.Connection = mcnHistory
'                    Call RunSQLScript(mrsSQLSys!FilePath, , , False)
'                End If
'                mrsSQLSys.MoveNext
'            Loop
'        End If
'        blnNormal = False
'        On Error GoTo 0
    End If
    Exit Sub
errh:
    tmrRefresh.Enabled = False
    If 0 = 1 Then
        Resume
    End If
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, App.Title
    Else
        MsgBox err.Description, vbInformation, App.Title
    End If
    Exit Sub
AbortLine:
    tmrRefresh.Enabled = False
    cmdCancel.Enabled = True
    Call UpgradeFinish(False)
End Sub

Private Function GetUpType() As Boolean
'----------------------------------------------------------------------------------------
'功能:获取升级方式信息
'----------------------------------------------------------------------------------------
    On Error GoTo errh
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    strSQL = " Select 项目,内容 From zlregInfo where 项目= '升级类型'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)

    If rsTemp.EOF = False Then
        strTemp = Nvl(rsTemp!内容, "0")
        If strTemp = "1" Then
            GetUpType = True
        Else
            GetUpType = False
        End If
    Else
        GetUpType = False
    End If
    Exit Function
errh:
    If err Then
        MsgBox "初始化升级方式出错,错误信息如下:" & vbCrLf & "错误号:" & err.Number & vbCrLf & "错误描述:" & err.Description, vbInformation, gstrSysName
    End If
End Function

Private Sub SetStepStateImg(ByVal vsnCurrent As VSFlexNode, Optional ByVal blnDone As Boolean)
'功能：设置步骤的状态图片
'参数：vsnCurrent=当前节点
'          blnDone=是否该步骤已经完成
    Dim vsnTmp As VSFlexNode, vsnParent As VSFlexNode
    Dim strImg As String
    strImg = IIf(blnDone, "Finish", "Doing")
    DoEvents
    If Not blnDone Then
        Set vsnTmp = vsnCurrent
        Do While Not vsnTmp Is Nothing
            Set vsnTmp.Image = imgPlan.ListImages(strImg).Picture
            vsnTmp.Expanded = True
            Set vsnTmp = vsnTmp.GetNode(flexNTParent)
        Loop
    Else
        Set vsnTmp = vsnCurrent.GetNode(flexNTNextSibling)
        Set vsnCurrent.Image = imgPlan.ListImages(strImg).Picture
        vsnCurrent.Expanded = False
        Set vsnParent = vsnCurrent
        Do While vsnParent.GetNode(flexNTNextSibling) Is Nothing '本机最后一个节点完成
            Set vsnParent = vsnParent.GetNode(flexNTParent)
            If vsnParent Is Nothing Then Exit Do
            Set vsnParent.Image = imgPlan.ListImages(strImg).Picture
            vsnParent.Expanded = False
        Loop
    End If
    vsPlan.Refresh
    
End Sub

Private Function UpgradeCheck(ByVal lngSys As Long) As Boolean
'功能：对系统进行对象检查
'参数：lngSys=系统号
'          strMsg=错误信息
    Dim cnTmp As ADODB.Connection
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strCheckFile As String, strName As String
    Dim strResult As String
    
    On Error GoTo errh
    mrsSysInfo.Filter = "系统编号=" & lngSys
    Call SetSQLState(False)
    If lngSys = 0 Then
        Set cnTmp = GetConnection("ZLTOOLS")
        strName = "zlUpgradeCheck"
        strCheckFile = gobjFile.GetParentFolderName(mrsSysInfo!配置文件) & "\" & strName & ".sql"
    Else
        Set cnTmp = gcnOldOra
        strName = "zl" & lngSys \ 100 & "_UpgradeCheck"
        strCheckFile = gobjFile.GetParentFolderName(gobjFile.GetParentFolderName(mrsSysInfo!配置文件)) & "\升级脚本\" & strName & ".sql"
    End If
    '创建检查函数
    mclsRunScript.IsUseLog = False
    lblFile.Caption = strCheckFile
    If Not mclsRunScript.ExecuteFile(strCheckFile, , , IIf(lngSys = 0, 1, 0), cnTmp) Then
        mclsRunScript.IsUseLog = True
        GoTo AbortLine
    End If
    mclsRunScript.IsUseLog = True
    err.Clear: On Error Resume Next
    strSQL = "Select " & strName & "('" & mrsSysInfo!系统版本号 & "', '" & mrsSysInfo!目标版本 & "') As Info From Dual"
    Set rsTmp = gclsBase.OpenSQLRecord(IIf(lngSys = 0, cnTmp, gcnOracle), strSQL, App.Title)
    If err.Number <> 0 Then '检查出错
        strResult = err.Description: err.Clear
        mclsRunScript.WriteLog "检查结果：" & strResult
        MsgBox strResult, vbExclamation, gstrSysName: GoTo AbortLine
    Else
        strResult = rsTmp!Info & ""
        If strResult <> "" Then
            mclsRunScript.WriteLog "检查结果：" & strResult
            MsgBox strResult, vbExclamation, gstrSysName: GoTo AbortLine
        Else
            mclsRunScript.WriteLog "检查结果：通过"
        End If
    End If
    UpgradeCheck = True
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
    Exit Function
AbortLine:
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function RunScriptByVersion(ByVal lngSys As Long, ByVal strVersion As String, Optional ByVal blnFirstUpdate As Boolean, Optional ByVal strOldVer As String, _
                                            Optional ByVal strAimVer As String, Optional blnHistory As Boolean, Optional ByVal cnTmp As ADODB.Connection, Optional ByVal strBakDB As String, _
                                            Optional ByVal blnUpInterface As Boolean) As Boolean
'功能：执行脚本文件并更新系统版本
'参数：lngSys=系统号
'         strVersion=当前到达的版本
'         blnFirstUpdate=是否第一次升迁版本更新
'         strOldVer=原始版本，blnFirstUpdate=True需传
'         strAimVer=目标版本，blnFirstUpdate=True需传
'         blnHistory=是否历史库版本更新
'         cnTmp=连接，历史库版本更新需要
'         blnUpInterface=是否升迁接口调用，升迁接口调用不能访问当前窗体空间对象以及属性，
'                                   当前有历史库单独升级与管理工具单独升级接口
    Dim strLogSQL As String, strVerSQL As String
    Dim blnAbort As Boolean, strAbortInfo As String
    Dim strResult As String
    
    On Error GoTo errh
    With mrsSysFiles
        .Filter = "系统编号=" & lngSys & " And SPVer='" & strVersion & "'" & IIf(blnHistory, " And  SysType=" & ST_History & " And 所有者='" & UCase(strBakDB) & "'", " And SysType<>" & ST_History)
        .Sort = "FileType"
        If .EOF And Not blnUpInterface Then Call SetSQLState(False)
        Do While Not .EOF
            If !FileType = FT_DBA Then
                Set mclsRunScript.Connection = gcnSystem: mclsRunScript.ConnectType = 2
            Else
                If lngSys = 0 Then
                    Set mclsRunScript.Connection = gcnTools: mclsRunScript.ConnectType = 1
                ElseIf Not blnHistory Then
                    Set mclsRunScript.Connection = gcnOldOra: mclsRunScript.ConnectType = 0
                Else
                    Set mclsRunScript.Connection = cnTmp: mclsRunScript.ConnectType = 0
                End If
            End If
            If Not RunSQLScript(!FilePath, Val(!AbortLine & ""), !Optional & "", blnHistory Or lngSys = 0, blnUpInterface) Then
                If Not blnHistory Then
                    If blnFirstUpdate Then '第一次更新版本,须在Zlupgrade中增加一条新记录
                        strLogSQL = "Insert Into Zlupgrade" & vbNewLine & _
                                    "  (系统, 原始版本, 目标版本, 升迁时间, 升迁结果, 结果版本, 中止语句, 提前执行)" & vbNewLine & _
                                    "  Select " & IIf(lngSys = 0, "Null", lngSys) & ", '" & strOldVer & "', '" & strAimVer & "', Sysdate, 1, '" & IIf(!FileType <= FT_Standard, strOldVer, strVersion) & "','" & Replace(mclsRunScript.AbortInfo, "'", "''") & "', " & IIf(mblnExecBef, 1, "Null") & " From Dual"
                    Else
                        strLogSQL = "Update Zlupgrade a" & vbNewLine & _
                                        "Set 结果版本 =" & IIf(!FileType <= FT_Standard, "结果版本", "'" & strVersion & "'") & " , 升迁结果=1 ,中止语句='" & Replace(mclsRunScript.AbortInfo, "'", "''") & "'" & vbNewLine & _
                                        "Where 系统 " & IIf(lngSys = 0, " Is Null", "=" & lngSys) & " And 升迁时间 = (Select Max(升迁时间) From Zlupgrade Where 系统 " & IIf(lngSys = 0, " Is Null", "=" & lngSys) & " And " & IIf(mblnExecBef, " Not ", "") & "  提前执行 Is Null)"
                    End If
                    '日志更新
                    gcnOracle.Execute strLogSQL
                Else
                    If Not mblnExecBef Then
                        '正式升级，清空相关提前执行信息
                        cnTmp.Execute "Update zlbakinfo Set 版本号=" & IIf(!FileType <= FT_Standard, "版本号", "'" & strVersion & "'") & " ,中止语句='" & Replace(mclsRunScript.AbortInfo, "'", "''") & "' Where 系统=" & lngSys
                    Else
                        '提前执行，保留提前执行版本，记录提前执行信息
                        cnTmp.Execute "Update zlbakinfo Set 提前中止语句='" & Replace(mclsRunScript.AbortInfo, "'", "''") & "' ,提前执行=1 Where 系统=" & lngSys
                    End If
                End If
                GoTo AbortLine
            End If
            .MoveNext
        Loop
    End With
    
    If Not blnHistory Then
        If blnFirstUpdate Then '第一次更新版本,须在Zlupgrade中增加一条新记录
            strLogSQL = "Insert Into Zlupgrade" & vbNewLine & _
                        "  (系统, 原始版本, 目标版本, 升迁时间, 升迁结果, 结果版本, 中止语句, 提前执行)" & vbNewLine & _
                        "  Select " & IIf(lngSys = 0, "Null", lngSys) & ", '" & strOldVer & "', '" & strAimVer & "', Sysdate, 0, '" & strVersion & "', Null, " & IIf(mblnExecBef, 1, "Null") & " From Dual"
        Else
            strLogSQL = "Update Zlupgrade a" & vbNewLine & _
                            "Set 结果版本 = '" & strVersion & "'" & vbNewLine & _
                            "Where 系统 " & IIf(lngSys = 0, " Is Null", "=" & lngSys) & " And 升迁时间 = (Select Max(升迁时间) From Zlupgrade Where 系统 " & IIf(lngSys = 0, " Is Null", "=" & lngSys) & " And " & IIf(mblnExecBef, " Not ", "") & "  提前执行 Is Null)"
        End If
        If Not mblnExecBef Then '提前执行不处理版本
            '系统版本更新
            If lngSys = 0 Then
                strVerSQL = "zlTools.B_Public.Update_Ver"
                '更新管理工具版本号:zlRegInfo
                '这里用ZLHIS的新连接处理,因为gcnTools是用的老连接用于执行脚本
                Call OpenCursor(gcnOracle, strVerSQL, strVersion)
            Else
                strVerSQL = "Update Zlsystems Set 版本号 = '" & strVersion & "' Where 编号 = " & lngSys
                gcnOracle.Execute strVerSQL
            End If
        End If
        '日志更新
        gcnOracle.Execute strLogSQL
    Else
        If Not mblnExecBef Then
            '正式升级，清空相关提前执行信息
            cnTmp.Execute "Update zlbakinfo Set 版本号='" & strVersion & "' ,中止语句=Null,提前中止语句=NULL,提前执行=0 Where 系统=" & lngSys
        Else
            '提前执行，保留提前执行版本，记录提前执行信息
            cnTmp.Execute "Update zlbakinfo Set 提前中止语句='" & strVersion & "' ,提前执行=1 Where 系统=" & lngSys
        End If
    End If
    RunScriptByVersion = True
    If Not blnUpInterface Then Call SetSQLState(False)
    Exit Function
AbortLine: '正常捕获到的中止跳转
    If mclsRunScript.Connection.State = adStateClosed Then
        If MsgBox("升级过程中连接意外中断，是否重试？", vbRetryCancel + vbInformation, App.Title) = vbRetry Then
            Resume
        End If
    End If
    If blnUpInterface Then Exit Function
    Call SetSessionParallel(mclsRunScript.Connection)
    Call SetSessionParallel(gcnOldOra)
    Call UpgradeFinish(False)
    cmdCancel.Enabled = True '不然不能Form_Unload
    MsgBox "系统升迁失败，请检查升迁日志文件并进行相应处理之后重新进行升迁。", vbExclamation, gstrSysName
    Exit Function
errh:
    If mclsRunScript.Connection.State = adStateClosed Then
        If MsgBox("升级过程中连接意外中断，是否重试？", vbRetryCancel + vbInformation, App.Title) = vbRetry Then
            Resume
        End If
    End If
    If MsgBox("升级过程中发生意外错误：" & vbNewLine & err.Description & vbNewLine & "是否重试？", vbRetryCancel + vbInformation, App.Title) = vbRetry Then
        Resume
    End If
End Function

Private Sub HTablePrivsRepair(ByVal lngSys As Long)
'功能：H表权限修正
    Dim objSQL As New clsSQLInfo
    Dim datStart As Date, lngSec As Long
    
    datStart = Now
    Call SetSQLState(False)
    On Error Resume Next
    objSQL.SQL = "Insert Into zlProgPrivs" & vbNewLine & _
            "  (系统, 序号, 功能, 对象, 所有者, 权限)" & vbNewLine & _
            "  Select 系统, 序号, 功能, 'H' || 对象, User, 'SELECT'" & vbNewLine & _
            "  From zlProgPrivs" & vbNewLine & _
            "  Where (Upper(所有者), Upper(对象)) In (Select User, 表名 From zlBakTables Where 系统 = " & lngSys & ") And Upper(权限) = 'SELECT' And" & vbNewLine & _
            "        系统 = " & lngSys & vbNewLine & _
            "  Minus" & vbNewLine & _
            "  Select 系统, 序号, 功能, 对象, User, 权限" & vbNewLine & _
            "  From zlProgPrivs" & vbNewLine & _
            "  Where 系统 = " & lngSys & "  And Upper(权限) = 'SELECT' And 对象 Like 'H%'"
    gcnOracle.Execute objSQL.SQL
    
    If err.Number <> 0 Then
        mclsRunScript.ErrCount = mclsRunScript.ErrCount + 1
        mclsRunScript.WriteLog "出 错 的 SQL：" & GetLogSQL(objSQL)
        mclsRunScript.WriteLog "错误(已忽略)：" & err.Description
        err.Clear
    End If
End Sub

Private Sub UpgradeFinish(ByVal blnSuccess As Boolean)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngLen  As Long
    
    Call SetSQLState(False)
    strSQL = "Select 编号, 版本号" & vbNewLine & _
                    "From Zlsystems " & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select 0, 内容 From Zlreginfo Where 项目 = '版本号'"
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    mclsRunScript.WriteSection "升迁系统" & IIf(blnSuccess, "成功！", "失败！")
    mclsRunScript.WriteLog "服务器时间：" & Format(CurrentDate, "yyyy-MM-dd HH:mm:ss") & String(4, " ") & "，本机时间：" & Format(Now, "yyyy-MM-dd HH:mm:ss")
    mrsSysInfo.Filter = "升级=1"
    mrsSysInfo.Sort = "Sort,系统编号"
    Do While Not mrsSysInfo.EOF
        rsTmp.Filter = "编号=" & mrsSysInfo!系统编号
        mclsRunScript.WriteLog IIf(mrsSysInfo!系统编号 <> 0, mrsSysInfo!系统编号 & "-", "") & mrsSysInfo!系统名称 & "：" & mrsSysInfo!系统版本号 & "-->" & rsTmp!版本号
        mrsSysInfo.MoveNext
    Loop
    mclsRunScript.WriteLog
    mclsRunScript.WriteLog "总共发生的错误次数：" & mclsRunScript.ErrCount
    If mclsRunScript.AbortInfo <> "" Then
        mclsRunScript.WriteLog "中止文件名称：" & Split(mclsRunScript.AbortInfo, "|")(0)
        mclsRunScript.WriteLog "中止文件行号：" & Split(mclsRunScript.AbortInfo, "|")(1)
    End If
    Call mclsRunScript.CloseLog
    If lblLog.Tag <> lblLogModi.Tag Then
        Call gobjFile.CopyFile(lblLogModi.Tag, lblLog.Tag, True)
    End If
End Sub

Private Function RunSQLScript(ByVal strFile As String, Optional ByVal lngAbort As Long, Optional strExecProcs As String, Optional ByVal blnHistory As Boolean, Optional ByVal blnUpInterface As Boolean) As Boolean
'功能：执行SQL脚本
'      strFile=SQL脚本名
'      lngAbort=中断号
'      strExecProcs=执行文件时，为选择的可选过程。
'      blnHistory=是否是历史库脚本
'      blnUpInterface=是否升迁接口调用，升迁接口调用不能访问当前窗体空间对象以及属性，
'                                   当前有历史库单独升级与管理工具单独升级接口
'返回：RunSQLScript=文件是否执行成功
    Dim strTmp As String
    Dim strTmpPath As String
    Dim strCaption As String, blnClose As Boolean

    With mclsRunScript
        .Procedures = strExecProcs
        .ProcMode = 0
        .GatherTables = ""
        If Not blnUpInterface Then
            Call SetSQLState(True, True)
            If ActualLen(strFile) <= 50 Then
                strCaption = "文件:" & strFile
            Else
                strTmpPath = gobjFile.GetParentFolderName(strFile)
                strTmp = gobjFile.GetFileName(strFile)
                strTmpPath = ActualStr(strTmpPath, 50 - ActualLen(strTmp) - 3) & "..."
                strCaption = "文件:" & strTmpPath & "\" & strTmp
            End If
        End If
        '执行存储过程，说明脚本是可选脚本，可选脚本中是存储过程，执行时不能从中断行号执行。
        If strExecProcs <> "" Then .Abort = 0: .ProcMode = 2
        If .OpenFile(strFile, lngAbort) Then
            Call SetSessionParallel(.Connection, True)
            Do While Not .EOF
                If Not blnUpInterface Then
                    lblFile.Caption = strCaption & "," & .Line
                    prgThis.value = .Line / .LinesCount * 100
                    lblPer.Caption = Format(prgThis.value / 100, "0%")
                    Me.txtSQL.Text = IIf(.SQLInfo.Tip <> "", .SQLInfo.Tip & vbCrLf, "") & .SQLInfo.SQL
                End If
                '数据结构修正结束或者遇到DLL才需要关闭并行
                If .SQLInfo.LobDDL And .SectionNumber < 2 Or .SectionNumber > 1 Then
                    Call SetSessionParallel(.Connection, False)
                    blnClose = True
                End If
                If .ExecuteSQL = False Then
                    Call SetSessionParallel(.Connection, False)
                    blnClose = True
                    Exit Function
                End If
                If .SQLInfo.LobDDL And .SectionNumber < 2 Then
                    Call SetSessionParallel(.Connection, True)
                    blnClose = False
                End If
                If Not blnUpInterface Then Call .CollectTables
                Call .ReadNextSQL
            Loop
            '可能没有SQL导致并行没有关闭，此处关闭
            If Not blnClose Then
                Call SetSessionParallel(.Connection, False)
            End If
            RunSQLScript = True
        Else
            RunSQLScript = False
        End If
        If Not blnHistory And Not blnUpInterface Then
            mstrChangeTables = mstrChangeTables & IIf(mstrChangeTables = "", "", ",") & .GatherTables
        End If
    End With
End Function

Private Sub UpdateSysFiles()
'功能：更新ZLSysFiles表
    On Error GoTo errh
    If mstrSysCodes = "" Then Exit Sub
    gcnOracle.Execute "Delete From zlSysFiles Where 系统 IN (" & mstrSysCodes & ")  And 操作 In(1,2)"
    mrsSysInfo.Filter = "系统编号<>0 And 升级=1"
    Do While Not mrsSysInfo.EOF
        gcnOracle.Execute "Insert Into zlSysFiles(系统,操作,文件名,日期,操作人) Values(" & _
                mrsSysInfo!系统编号 & ",1,'" & Replace(ActualStr(mrsSysInfo!配置文件, 100), "'", "''") & "',Sysdate,User)"
        mrsSysInfo.MoveNext
    Loop
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Private Sub ReCompileObjects(cnThis As ADODB.Connection)
'功能：编译指定连接所有者的无效对象
'参数：cnThis=所有者连接,本函数可针对不同所有者调用
    Dim rsObjects As New ADODB.Recordset
    Dim rsDepends As New ADODB.Recordset
    Dim arrObjects As Variant, strCompile As String
    Dim strSQL As String, i As Long
    Dim strUser As String
    Dim arrTmp As Variant
    
    lblFile.Caption = "正在编译无效对象 ...": txtSQL.Text = ""
    prgThis.value = 0: lblPer.Caption = "0%"
    
    On Error GoTo errHandle
    strSQL = _
        "Select User, Object_Name, Object_Type" & vbNewLine & _
        "From User_Objects" & vbNewLine & _
        "Where Object_Type In" & vbNewLine & _
        "      ('PROCEDURE', 'FUNCTION', 'VIEW', 'MATERIALIZED VIEW', 'TRIGGER', 'PACKAGE', 'PACKAGE BODY', 'TYPE', 'TYPE BODY') And" & vbNewLine & _
        "      Object_Name Not Like 'BIN$%' And Status = 'INVALID'" & vbNewLine & _
        "Order By Object_Type, Object_Name"
    rsObjects.CursorLocation = adUseClient
    rsObjects.Open strSQL, cnThis, adOpenKeyset
    If Not rsObjects.EOF Then
        strUser = rsObjects!User
        strSQL = _
            "Select Name, Type, Referenced_Name, Referenced_Type" & vbNewLine & _
            "From User_Dependencies" & vbNewLine & _
            "Where Referenced_Owner = User And Type In ('PROCEDURE', 'FUNCTION', 'VIEW', 'MATERIALIZED VIEW', 'TRIGGER', 'PACKAGE'," & vbNewLine & _
            "       'PACKAGE BODY', 'TYPE', 'TYPE BODY') And" & vbNewLine & _
            "      Referenced_Type In" & vbNewLine & _
            "      ('PROCEDURE', 'FUNCTION', 'VIEW', 'MATERIALIZED VIEW', 'TRIGGER', 'PACKAGE', 'PACKAGE BODY', 'TYPE', 'TYPE BODY') And" & vbNewLine & _
            "      Not(Name=Referenced_Name And Type=Referenced_Type) And" & vbNewLine & _
            "      Name Not Like 'BIN$%' And Referenced_Name Not Like 'BIN$%'"
        rsDepends.CursorLocation = adUseClient
        rsDepends.Open strSQL, cnThis, adOpenKeyset

        ReDim arrObjects(rsObjects.RecordCount - 1) As String
        For i = 1 To rsObjects.RecordCount
            arrObjects(i - 1) = rsObjects!Object_Name & "," & rsObjects!Object_Type
            rsObjects.MoveNext
        Next

        '编译无效对象
        For i = 0 To UBound(arrObjects)
            arrTmp = Split(arrObjects(i), ",")
            lblFile.Caption = "正在编译无效对象 " & i + 1 & "/" & (UBound(arrObjects) + 1) & " ..."
            prgThis.value = (i + 1) / (UBound(arrObjects) + 1) * 100
            lblPer.Caption = Format(prgThis.value / 100, "0%")
            Call ComplieObject(cnThis, arrTmp(0), arrTmp(1), rsObjects, rsDepends, strCompile)
        Next
        mclsRunScript.WriteLog RPAD("共编译了 " & strUser & " 的 " & UBound(arrObjects) + 1 & " 个无效对象", 33)
    End If
    Exit Sub
errHandle: '函数内部的其他未知异常
    If MsgBox(err.Description, vbRetryCancel + vbCritical, gstrSysName) = vbRetry Then Resume
End Sub

Private Sub ComplieObject(cnThis As ADODB.Connection, ByVal strName As String, ByVal strType As String, _
    rsObjects As ADODB.Recordset, rsDepends As ADODB.Recordset, strCompile As String)
'功能：编译指定的无效对象
'参数：strCompile=已经编译的对象名串
'说明：ReCompileObjects的子函数
    Dim arrObjRef As Variant, strErrInfor As String
    Dim strSQL As String, i As Long

    If InStr(strCompile & ",", "," & strName & ",") > 0 Then Exit Sub

    '递归编译当前对象所引用的对象
    rsDepends.Filter = "Name='" & strName & "' And Type='" & strType & "'" '不加类型可能引起递归溢出(同名BODY)
    If Not rsDepends.EOF Then
        ReDim arrObjRef(rsDepends.RecordCount - 1) As String
        For i = 1 To rsDepends.RecordCount
            arrObjRef(i - 1) = rsDepends!Referenced_Name & "," & rsDepends!Referenced_Type
            rsDepends.MoveNext
        Next
        For i = 0 To UBound(arrObjRef)
            rsObjects.Filter = "Object_Name='" & Split(arrObjRef(i), ",")(0) & "' And Object_Type='" & Split(arrObjRef(i), ",")(1) & "'"
            If Not rsObjects.EOF Then '引用对象也是无效对象时
                Call ComplieObject(cnThis, Split(arrObjRef(i), ",")(0), Split(arrObjRef(i), ",")(1), rsObjects, rsDepends, strCompile)
            End If
        Next
    End If

    '编译当前对象
    Select Case strType
    Case "PROCEDURE"
        strSQL = "ALTER PROCEDURE " & strName & " COMPILE"
    Case "FUNCTION"
        strSQL = "ALTER FUNCTION " & strName & " COMPILE"
    Case "VIEW"
        strSQL = "ALTER VIEW " & strName & " COMPILE"
    Case "MATERIALIZED VIEW"
        strSQL = "ALTER MATERIALIZED VIEW " & strName & " COMPILE"
    Case "TRIGGER"
        strSQL = "ALTER TRIGGER " & strName & " COMPILE"
    Case "PACKAGE"
        strSQL = "ALTER PACKAGE " & strName & " COMPILE"
    Case "PACKAGE BODY"
        strSQL = "ALTER PACKAGE " & strName & " COMPILE BODY"
    Case "TYPE"
        strSQL = "ALTER TYPE " & strName & " COMPILE"
    Case "TYPE BODY"
        strSQL = "ALTER TYPE " & strName & " COMPILE BODY"
    End Select
    If strSQL <> "" Then
        txtSQL.Text = txtSQL.Text & strSQL & vbCrLf
        txtSQL.SelStart = Len(txtSQL.Text): DoEvents
    
        strErrInfor = ""
        err.Clear: On Error Resume Next
        cnThis.Execute strSQL
        If cnThis.Errors.Count > 0 Then
            '特殊情况(未出错):Err.Number=0,NativeError=0
            '[Microsoft][ODBC driver for Oracle]创建的过程或软件包带有编译错误
            '没有更多的结果。
            If Not (cnThis.Errors(0).NativeError = 0 And cnThis.Errors.Count = 1) Then
                If cnThis.Errors(0).NativeError <> 0 Then
                    strErrInfor = strName & ":" & cnThis.Errors(0).Description
                Else
                    strErrInfor = strName & ":其他编译错误"
                End If
            End If
        End If
        If strErrInfor <> "" Then
            mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "：" & strSQL & "，错误：" & strErrInfor
        End If
        err.Clear: On Error GoTo 0
        strCompile = strCompile & "," & strName
    End If
End Sub

Private Sub ReAdjustSequence(ByVal cnThis As ADODB.Connection)
'功能：重新调整序列
'参数：cnThis=所有者连接,本函数可针对不同所有者调用
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strUser As String, K As Long
    On Error GoTo errHandle

    lblFile.Caption = "正在重新调整序列 ..."
    txtSQL.Text = "": txtSQL.Enabled = False: txtSQL.BackColor = Me.BackColor
    prgThis.value = 0: lblPer.Caption = "0%"

    Set rsTemp = GetSequence("", cnThis, True)
    For i = 1 To rsTemp.RecordCount
        K = K + 1
        If strUser = "" Then strUser = rsTemp!User

        prgThis.value = i / rsTemp.RecordCount * 100
        lblPer.Caption = Format(prgThis.value / 100, "0%")
        lblFile.Caption = "正在重新调整序列:" & rsTemp!Sequence_Name & " ..."
        DoEvents
        Call AdjustNameSequece(rsTemp!User & "." & rsTemp!Table_Name, cnThis, rsTemp!Column_Name)
        rsTemp.MoveNext
    Next

    If mblnHaveST And Not cnThis Is gcnTools Then
        K = K + 1
        If strUser = "" Then strUser = gstrUserName
        lblFile.Caption = "正在重新调整序列:结帐ID ...": DoEvents
        Call Adjust结帐ID(cnThis)
    End If

    mclsRunScript.WriteLog RPAD("共对 " & strUser & " 的 " & K & " 个序列进行了重新整理", 33)
    txtSQL.Enabled = True: txtSQL.BackColor = &H80000005
    Exit Sub
errHandle: '函数内部的其他未知异常
    If MsgBox(err.Description, vbRetryCancel + vbCritical, gstrSysName) = vbRetry Then Resume
End Sub


Private Sub ImportReports(ByVal lngSys As Long)
'功能：导入报表
'说明：出错不中止升迁
    Dim i As Long, lngCount As Long, lngAll As Long
    Dim datStart As Date, lngSec As Long
    
    datStart = Now
    mrsReport.Filter = "系统编号=" & lngSys & " And 覆盖类型<>0"
    lngAll = mrsReport.RecordCount
    mrsReport.Sort = "ID"
    lblFile.Caption = "正在导入报表 ...": txtSQL.Text = ""
    prgThis.value = 0: lblPer.Caption = "0%"
    If gobjReport Is Nothing Then
        err.Clear: On Error Resume Next
        Set gobjReport = CreateObject("zl9Report.clsReport")
        err.Clear: On Error GoTo 0
    End If
    If gobjReport Is Nothing Then
        txtSQL.Text = "报表部件创建失败,不能对报表进行导入!"
        mclsRunScript.ErrCount = mclsRunScript.ErrCount + 1
        mclsRunScript.WriteLog String(4, " ") & txtSQL.Text: Sleep 2000: Exit Sub
    End If
    lngCount = 0
    
    For i = 1 To mrsReport.RecordCount
        prgThis.value = i / (mrsReport.RecordCount) * 100
        lblPer.Caption = Format(prgThis.value / 100, "0%")
        lblFile.Caption = "正在导入报表 " & i & "/" & mrsReport.RecordCount & " ..."
        txtSQL.Text = txtSQL.Text & "导入:" & mrsReport!编号 & "/" & mrsReport!名称
        txtSQL.SelStart = Len(txtSQL.Text): DoEvents
        If gobjFile.FileExists(mrsReport!FilePath) Then
            '###
            If gobjReport.ReportImport(mrsReport!FilePath, gcnOracle, mrsReport!编号, mrsReport!覆盖类型 = 2) Then
                txtSQL.Text = txtSQL.Text & ",成功!"
                mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "，" & RPAD(mrsReport!FilePath, 70) & String(4, " ") & IIf(mrsReport!覆盖类型 = 2, ",导入数据源成功", "整体导入成功")
            Else
                lngCount = lngCount + 1
                txtSQL.Text = txtSQL.Text & ",失败!"
                mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "，" & RPAD(mrsReport!FilePath, 70) & String(4, " ") & IIf(mrsReport!覆盖类型 = 2, ",导入数据源失败", "整体导入失败")
            End If
        Else
            lngCount = lngCount + 1
            txtSQL.Text = txtSQL.Text & ",文件不存在!"
            mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "，" & "丢失文件:" & RPAD(mrsReport!FilePath, 65) & String(4, " ") & IIf(mrsReport!覆盖类型 = 2, ",导入数据源", "整体导入")
        End If
        txtSQL.Text = txtSQL.Text & vbCrLf
        txtSQL.SelStart = Len(txtSQL.Text): DoEvents
        mrsReport.MoveNext
    Next
    lngSec = DateDiff("s", datStart, Now)
    mclsRunScript.WriteLog RPAD("共" & (lngAll - lngCount) & " 张报表导入成功," & lngCount & "张报表导入失败", 33)
    mclsRunScript.ErrCount = mclsRunScript.ErrCount + lngCount
End Sub

Private Sub GrantToRole()
    Dim lngCount As Long
    Dim strOwners(0 To 0) As String

    On Error Resume Next
    strOwners(0) = gstrUserName
    '授予权限表中填写的权限
    Call SetSQLState(True)
    lblFile.Caption = "正在对角色重新授权 ..."
    Call ReGrantToRole(gcnOracle, "", True, strOwners, prgThis, lblPer, lngCount)
    mclsRunScript.WriteLog RPAD("共对 " & lngCount & " 个角色进行了重新授权", 33)
    txtSQL.Enabled = True: txtSQL.BackColor = &H80000005
End Sub

Private Sub GatherStatistics()
'功能：搜集统计信息（仅历史库升级时，只搜集历史库，否则历史库与在线库均搜集）
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim rsBakTable As ADODB.Recordset
    Dim lngCount As Long, i As Long, lngCur As Long
    Dim strUser As String, strOtherPara As String
    Dim cnDBA As ADODB.Connection
    Dim datStart As Date, datStartTmp As Date, lngSec As Long, lngSecTmp As Long
    Dim lngErr As Long
    
    SetSQLState (True)
    lblFile.Caption = "正在对大表进行统计信息收集 ..."
    datStart = Now
    On Error Resume Next
    strSQL = "Select Distinct A.表名" & vbNewLine & _
                    "From (Select 表名" & vbNewLine & _
                    "       From Zlbigtables" & vbNewLine & _
                    "       Where 系统 in(" & mstrSysCodes & ")" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select 表名 From zlBakTables Where 系统 in(" & mstrSysCodes & ")) A"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    If err.Number <> 0 Then
        err.Clear
        strSQL = "Select Distinct 表名" & vbNewLine & _
                "From zlBakTables" & vbNewLine & _
                "Where 系统 in(" & mstrSysCodes & ")" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select Column_Value From Table(F_Str2list('病人信息,病案主页,病人信息从表,病案主页从表,就诊登记记录,医保病人档案,医保病人关联表'))"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    End If
    If err.Number = 0 Then
        On Error GoTo errh
        Set rsBakTable = CopyNewRec(rsTmp, , , Array("收集", adInteger, Empty, Empty))
        mstrChangeTables = "," & UCase(mstrChangeTables) & ","
        mstrChangeTables = Replace(Replace(Replace(mstrChangeTables, vbNewLine, ""), ",,", ","), ",,", ",")
        
        '标记需要收集的表
        rsBakTable.Filter = ""
        For i = 1 To rsBakTable.RecordCount
            If mstrChangeTables = "," Then Exit For
            If mstrChangeTables Like "*," & UCase(rsBakTable!表名) & ",*" Then
                If ",病人信息,病案主页,病人信息从表,病案主页从表,就诊登记记录,医保病人档案,医保病人关联表," Like "*," & rsBakTable!表名 & ",*" Then
                    rsBakTable.Update "收集", 2
                Else
                    rsBakTable.Update "收集", 1
                End If
            Else
                rsBakTable.Update "收集", 0
            End If
            mstrChangeTables = Replace(Replace(mstrChangeTables, "," & UCase(rsBakTable!表名) & ",", ","), ",,", ",")
            rsBakTable.MoveNext
        Next
        
        mrsHistorySpace.Filter = "升级=1 And 当前=1 And Db连接=Null"
        rsBakTable.Filter = "收集=1"
        lngCount = rsBakTable.RecordCount * mrsHistorySpace.RecordCount
        rsBakTable.Filter = "收集<>0"
        lngCount = lngCount + rsBakTable.RecordCount
        
        'i=0 标识在线库统计信息收集，历史库收集表与在线库相同
        strOtherPara = ",cascade => True" & IIf(chkAutoPer.value = 1, "", ",estimate_percent => " & Val(txtEsper.Text)) & _
                        ",method_opt => 'for all columns size skewonly'" & _
                        IIf(Val(txtCpu(1).Text) = 0, "", ",degree => " & txtCpu(1).Text) & ",no_invalidate => false)"
        Set cnDBA = GetConnection("DBA")
        
        For i = 0 To mrsHistorySpace.RecordCount
            If i = 0 Then
                mclsRunScript.WriteLog "收集统计信息的参数：" & Mid(strOtherPara, 2), , 1
                strUser = gstrUserName
                rsBakTable.Filter = "收集<>0"
            Else
                strUser = mrsHistorySpace!所有者
                If i = 1 Then rsBakTable.Filter = "收集=1"
            End If
            If rsBakTable.RecordCount <> 0 Then rsBakTable.MoveFirst
            DoEvents
            Do While Not rsBakTable.EOF
                lngCur = lngCur + 1
                prgThis.value = lngCur / lngCount * 100
                lblPer.Caption = Format(prgThis.value / 100, "0%")
                lblFile.Caption = "正在对表:" & strUser & "." & rsBakTable!表名 & "进行统计信息搜集 ..."
                datStartTmp = Now
                Me.Refresh
                
                strSQL = "dbms_stats.gather_table_stats(ownname => '" & strUser & "',tabname =>'" & rsBakTable!表名 & "'" & strOtherPara
                If optStatType(0).value Then '直接升级过程中收集
                    '调用包时指定参数名，仅ODBC连接方式支持
                    '用connection对象，excute方法的Options参数值为这几个都可以：adCmdUnknown 'adCmdStoredProc 'adExecuteNoRecords
                    '用Command对象，必须指定CommandType = adCmdStoredProc
                    On Error Resume Next
                    cnDBA.Execute strSQL, , adExecuteNoRecords
                    If err.Number = 0 Then
                        lngSecTmp = DateDiff("s", datStartTmp, Now)
                        mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "，" & RPAD(strUser & "." & rsBakTable!表名, 50) & "耗时：" & IIf(lngSecTmp > 60, (lngSecTmp \ 60) & "分钟" & (lngSecTmp Mod 60) & "秒", lngSecTmp & "秒")
                    Else
                        mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "，" & RPAD(strUser & "." & rsBakTable!表名 & String(8, " ") & "收集失败", 50) & "错误：" & err.Description & String(8, " ") & "SQL:" & strSQL
                        err.Clear: lngErr = lngErr + 1
                    End If
                Else '仅记录收集表
                    mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "，" & strSQL
                End If
                rsBakTable.MoveNext
            Loop
            If i <> 0 Then mrsHistorySpace.MoveNext
        Next
        lngSec = DateDiff("s", datStart, Now)
        mclsRunScript.WriteLog "共对 " & lngCount & " 个表进行了统计信息收集", , 1
    Else
        mclsRunScript.WriteLog "由于未查询到定义的大表，因此没有对任何表进行统计信息收集"
    End If
    mclsRunScript.ErrCount = mclsRunScript.ErrCount + lngErr
    SetSQLState
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Private Sub RepairHisDB(ByVal cnHistory As ADODB.Connection, ByVal lngSys As Long, ByVal strBakUser As String, ByVal strBakServer As String, _
    ByVal strBakSpaceName As String, ByVal strDbLink As String, Optional ByVal blnCurDB As Boolean, Optional ByVal blnAloneUpHistory As Boolean)
'功能：修正历史库结构问题
'参数：blnAloneUpHistory-True:单独升级历史库,false:系统升级界面中升级历史库
    Dim datStartTmp As Date, lngSecTmp As Long
    Dim rsRepairSQL As ADODB.Recordset, lngCount As Long, i As Long
    Dim comTmp As New ADODB.Command
    
    On Error GoTo errh
    If Not blnAloneUpHistory Then
        Call SetSQLState(True, True)
        lblFile.Caption = "正在检查历史库结构问题 ..."
    End If
    datStartTmp = Now
    
    '搜集历史库修正SQL
    Call frmHistorySpaceRepair.ShowRepair(Me, lngSys, True, strBakUser, strBakSpaceName, blnCurDB, rsRepairSQL, cnHistory, strDbLink)
    lngSecTmp = DateDiff("s", datStartTmp, Now)
    If Not rsRepairSQL Is Nothing Then
        mclsRunScript.WriteLog RPAD("历史库结构检查发现" & rsRepairSQL.RecordCount & "个问题", 30) & ",耗时" & IIf(lngSecTmp > 60, (lngSecTmp \ 60) & "分钟" & (lngSecTmp Mod 60) & "秒", lngSecTmp & "秒")
        rsRepairSQL.Sort = "ExecOrder,FixType,ExecIndex"
        lngCount = rsRepairSQL.RecordCount: datStartTmp = Now
        If lngCount <> 0 And Not blnAloneUpHistory Then lblFile.Caption = "正在修正" & strBakUser & "的结构问题 ..."
        Call SetSessionParallel(cnHistory, True)
        Call SetSessionParallel(gcnOracle, True)
        On Error Resume Next
        For i = 1 To rsRepairSQL.RecordCount
            If Not blnAloneUpHistory Then
                prgThis.value = i / lngCount * 100
                lblPer.Caption = Format(prgThis.value / 100, "0%")
                Me.Refresh
            End If
            If rsRepairSQL!ExecDB = 1 Then
                Set comTmp.ActiveConnection = gcnOracle
            Else
                Set comTmp.ActiveConnection = cnHistory
            End If
            comTmp.CommandText = rsRepairSQL!SQL
            DoEvents
            comTmp.Execute
            If err.Number <> 0 Then
                mclsRunScript.ErrCount = mclsRunScript.ErrCount + 1
                mclsRunScript.WriteLog Format(Now, "HH:mm:ss") & "，" & IIf(rsRepairSQL!ExecDB = 0, "历史库：" & strBakUser & "，", "在线库，") & rsRepairSQL!SQL
                mclsRunScript.WriteLog "错误（已忽略）：" & err.Description
                err.Clear
            End If
            rsRepairSQL.MoveNext
        Next
        Call SetSessionParallel(cnHistory, False)
        Call SetSessionParallel(gcnOracle, False)
    End If
    If strDbLink = "" Then
         If Not blnAloneUpHistory Then lblFile.Caption = "正在修正" & strBakUser & "的访问权限问题 ..."
        '需要重新授权,向所有者:刘兴宏20071202
        Call GrantBakToUser(cnHistory, gstrUserName)
    End If
    If blnCurDB Then
         If Not blnAloneUpHistory Then
            lblFile.Caption = "正在修正" & strBakUser & "的历史数据空间视图 ..."
            lblPer.Caption = ""
        End If
        Call CreateAppView(gstrUserName, strBakUser, lngSys, IIf(strDbLink = "", "", "@" & strDbLink), IIf(blnAloneUpHistory, Nothing, prgThis), mclsRunScript)
    End If
     If Not blnAloneUpHistory Then Me.Refresh
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Private Sub SetSessionParallel(ByRef cnInput As ADODB.Connection, Optional ByVal blnEnabled As Boolean)
'启用或禁用DDL
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errh
    If mintDDLParallel <= 1 Then Exit Sub
    If blnEnabled Then
        strSQL = "Alter Session FORCE PARALLEL DDL PARALLEL " & mintDDLParallel
        cnInput.Execute strSQL
    Else
        strSQL = "ALTER Session DISABLE PARALLEL DDL "
        cnInput.Execute strSQL
        strSQL = "Select 'alter index ' || Index_Name || ' noparallel' SQL" & vbNewLine & _
                    "From User_Indexes" & vbNewLine & _
                    "Where Degree Not In ('0', '1') and index_type='NORMAL' " & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select 'alter table ' || Table_Name || ' noparallel' SQL" & vbNewLine & _
                    "From User_Tables" & vbNewLine & _
                    "Where Degree != ('         1')"
        Set rsTmp = gclsBase.OpenSQLRecord(cnInput, strSQL, App.Title)
        On Error Resume Next
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                cnInput.Execute rsTmp!SQL, , adCmdText
                If err.Number <> 0 Then
                    mclsRunScript.WriteLog "取消并行出错：" & rsTmp!SQL
                    If cnInput.Errors.Count > 0 Then
                        mclsRunScript.WriteLog "错误（已忽略）：" & cnInput.Errors(0).Description
                    Else
                        mclsRunScript.WriteLog "错误（已忽略）：" & err.Description
                    End If
                    err.Clear
                End If
                rsTmp.MoveNext
            Loop
        End If
    End If
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, App.Title
End Sub

Private Function GetCode(ByVal strCaption As String) As String
'功能：获取流程的编码
    Dim arrTmp As Variant, i As Long
    Dim strCode As String
    
    arrTmp = Split(strCaption, ".")
    For i = LBound(arrTmp) To UBound(arrTmp)
        If i <> UBound(arrTmp) Then
            strCode = strCode & "." & arrTmp(i)
        End If
    Next
    GetCode = Mid(strCode, 2)
End Function

Private Sub SetCpuCount()
'功能：设置统计信息收集以及并行DDL的并行度
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Integer, intDefault As Integer, intMax As Integer, intMin As Integer
    
     '最大并行为CPU数，防止过高，实际为CPU个数*单个CPU上并行进程
'    Dim intPerParallel As Integer
'    strSQL = "Select Nvl(Max(Value),0) Parallel From V$parameter Where Name = 'parallel_threads_per_cpu'"
'    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取单个CPU并行数")
'    intPerParallel = Val(rsTmp!Parallel, "")
'    intPerParallel = IIf(intPerParallel < 1, 1, intPerParallel) '防御性编程，不了解实际ORacle这个参数情况
    strSQL = "Select Nvl(Max(Value),0) CPU From V$parameter Where Name = 'cpu_count'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取可用CUP数")
    
    intMin = 1
    If rsTmp!cpu <= 4 Then
        chkParallel.value = 0: chkParallel.Enabled = False: lblStaCpuName.Tag = "Cpu<=4"
        intDefault = 1
        intMax = IIf(rsTmp!cpu = 0, 1, rsTmp!cpu)
        
        For i = 0 To 1
            lblCpuWarn(i).Visible = True: lblCpuWarn(i).Tag = "显示警告"
        Next
        Call SetCtrlPosOnLine(False, 0, lblCpuWarn(0), 60, ckhIdxOnLine)
    ElseIf rsTmp!cpu <= 8 Then
        intDefault = 4
        intMax = rsTmp!cpu
    ElseIf rsTmp!cpu <= 12 Then
        intDefault = 8
        intMax = rsTmp!cpu
    Else
        intDefault = 12
        intMax = rsTmp!cpu
    End If
    For i = 0 To 1
        txtCpu(i).Text = intDefault
        udCpu(i).Max = intMax '最大并行只为CPU数，防止过高，实际为CPU个数*单个CPU上并行进程
'        udCpu(i).Max = intMax * intPerParallel '最高并行度为CPU个数*单个CPU上并行进程
        udCpu(i).Min = intMin
    Next
End Sub

Private Sub SetSQLState(Optional ByVal blnStart As Boolean, Optional ByVal blnSQLEnable As Boolean)
    lblFile.Caption = "": txtSQL.Text = ""
    prgThis.value = 0: lblPer.Caption = "0%"
    lblPer.Visible = blnStart
    lblFile.Visible = blnStart
    prgThis.Visible = blnStart
    lblPer.Visible = blnStart
    txtSQL.Enabled = blnSQLEnable
    txtSQL.BackColor = IIf(blnSQLEnable, &H80000005, Me.BackColor)
End Sub

Private Sub vsPlan_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Function JudgeOldToolsVer() As String
'功能：判断管理工具的版本
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 编号 from zlSvrTools where 编号='0502'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取zlSvrTools")
    If rsTmp.EOF = True Then
        '那是最早的，版本为9.0.0
        JudgeOldToolsVer = "9.0.0"
        Exit Function
    End If
    
    strSQL = "SELECT CONSTRAINT_NAME FROM All_Constraints C WHERE C.CONSTRAINT_NAME='ZLOPTIONS_PK' AND C.OWNER='ZLTOOLS' AND C.TABLE_NAME='ZLOPTIONS'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "判别ZLOPTIONS_PK")
    If rsTmp.EOF = True Then
        '如果不存在ZLOPTIONS_PK约束，说明没有执行第二个升级脚本，版本为9.1.0
        JudgeOldToolsVer = "9.1.0"
        Exit Function
    End If
    strSQL = "SELECT CONSTRAINT_NAME FROM All_Constraints C WHERE C.CONSTRAINT_NAME='ZLXLSVERIFY_FK_报表号' AND C.OWNER='ZLTOOLS' AND C.TABLE_NAME='ZLXLSVERIFY'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "判别ZLXLSVERIFY_FK_报表号")
    If Not rsTmp.EOF Then
        '如果存在ZLXLSVERIFY_FK_报表号约束，说明没有执行第三个升级脚本，版本为9.2.0
        JudgeOldToolsVer = "9.2.0"
        Exit Function
    End If
    JudgeOldToolsVer = "9.3.0"
End Function

Private Sub AdjustZLupgrade()
'修正ZLupgrade的目标版本
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error Resume Next
    strSQL = "Select a.Owner" & vbNewLine & _
        "From All_Tab_Columns a" & vbNewLine & _
        "Where a.Table_Name = 'ZLUPGRADE' And a.Column_Name = '目标版本' And a.Data_Length < 20"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "判别ZLUPGRADE目标版本长度")
    If Not rsTmp.EOF Then
        gcnOracle.Execute "alter table " & rsTmp!Owner & ".ZLUPGRADE modify 目标版本 varchar2(20)", , adCmdText
    End If
    If err.Number <> 0 Then err.Clear
End Sub

