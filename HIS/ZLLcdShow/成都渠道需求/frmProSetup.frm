VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmProSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6495
   Icon            =   "frmProSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   6495
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4080
      TabIndex        =   120
      Top             =   6720
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5275
      TabIndex        =   121
      Top             =   6720
      Width           =   1100
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基础设置(&1)"
      TabPicture(0)   =   "frmProSetup.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra显示模式"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frm背景图片"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmRect"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra数据刷新"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "呼叫区域设置(&2)"
      TabPicture(1)   =   "frmProSetup.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frm呼叫区域"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fra药房"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "排队区域设置(&3)"
      TabPicture(2)   =   "frmProSetup.frx":688A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra待发药"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "提示区域设置(&4)"
      TabPicture(3)   =   "frmProSetup.frx":68A6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frm提示"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "fra显示时间"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.Frame fra数据刷新 
         Caption         =   " 数据刷新 "
         Height          =   1215
         Left            =   120
         TabIndex        =   116
         Top             =   4200
         Width           =   6015
         Begin VB.TextBox txt呼叫显示时间 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1440
            TabIndex        =   122
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txt数据轮询时间 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1440
            TabIndex        =   118
            Top             =   315
            Width           =   1215
         End
         Begin VB.Label lbl数据轮询时间 
            AutoSize        =   -1  'True
            Caption         =   "呼叫显示时间"
            Height          =   180
            Index           =   3
            Left            =   240
            TabIndex        =   124
            Top             =   765
            Width           =   1080
         End
         Begin VB.Label lbl数据轮询时间 
            AutoSize        =   -1  'True
            Caption         =   "秒(范围：1-60)"
            Height          =   180
            Index           =   2
            Left            =   2760
            TabIndex        =   123
            Top             =   765
            Width           =   1260
         End
         Begin VB.Label lbl数据轮询时间 
            AutoSize        =   -1  'True
            Caption         =   "秒(范围：1-60)"
            Height          =   180
            Index           =   1
            Left            =   2760
            TabIndex        =   119
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label lbl数据轮询时间 
            AutoSize        =   -1  'True
            Caption         =   "数据轮询时间"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   117
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   96
         Top             =   3300
         Width           =   6015
         Begin VB.TextBox txt已过号_行数 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   4200
            TabIndex        =   114
            Top             =   2235
            Width           =   975
         End
         Begin VB.TextBox txt已过号_行高 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2280
            TabIndex        =   112
            Top             =   2235
            Width           =   975
         End
         Begin VB.TextBox txt已过号_列宽 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   720
            TabIndex        =   110
            Top             =   2235
            Width           =   855
         End
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   10
            Left            =   240
            TabIndex        =   108
            Top             =   1800
            Width           =   5595
         End
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   9
            Left            =   240
            TabIndex        =   103
            Top             =   960
            Width           =   5595
         End
         Begin VB.CommandButton cmd已过号颜色 
            Caption         =   "字体颜色"
            Height          =   350
            Left            =   4200
            TabIndex        =   107
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmd已过号字体 
            Caption         =   "字体设置"
            Height          =   350
            Left            =   240
            TabIndex        =   105
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txt已过号_顶 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2880
            TabIndex        =   102
            Top             =   555
            Width           =   1695
         End
         Begin VB.TextBox txt已过号_左 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   100
            Top             =   555
            Width           =   1695
         End
         Begin VB.CheckBox chk显示已过号 
            Caption         =   "显示已过号"
            Height          =   180
            Left            =   240
            TabIndex        =   97
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "行数："
            Height          =   180
            Index           =   34
            Left            =   3720
            TabIndex        =   115
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "行高："
            Height          =   180
            Index           =   33
            Left            =   1800
            TabIndex        =   113
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "列宽："
            Height          =   180
            Index           =   32
            Left            =   240
            TabIndex        =   111
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "表格"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   31
            Left            =   240
            TabIndex        =   109
            Top             =   1920
            Width           =   360
         End
         Begin VB.Shape shp已过号颜色 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   5280
            Top             =   1350
            Width           =   375
         End
         Begin VB.Label lbl已过号字体 
            AutoSize        =   -1  'True
            Caption         =   "微软雅黑;12"
            Height          =   180
            Left            =   1350
            TabIndex        =   106
            Top             =   1410
            Width           =   990
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "字体"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   30
            Left            =   240
            TabIndex        =   104
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "顶："
            Height          =   180
            Index           =   29
            Left            =   2520
            TabIndex        =   101
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "位置"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   28
            Left            =   240
            TabIndex        =   99
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "左："
            Height          =   180
            Index           =   27
            Left            =   240
            TabIndex        =   98
            Top             =   600
            Width           =   360
         End
      End
      Begin VB.Frame fra待发药 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   27
         Top             =   420
         Width           =   6015
         Begin VB.TextBox txt待发药_行数 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   4200
            TabIndex        =   45
            Top             =   2235
            Width           =   975
         End
         Begin VB.TextBox txt待发药_行高 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2280
            TabIndex        =   43
            Top             =   2235
            Width           =   975
         End
         Begin VB.TextBox txt待发药_列宽 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   720
            TabIndex        =   41
            Top             =   2235
            Width           =   855
         End
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   7
            Left            =   240
            TabIndex        =   39
            Top             =   1800
            Width           =   4755
         End
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   6
            Left            =   240
            TabIndex        =   34
            Top             =   960
            Width           =   5595
         End
         Begin VB.CommandButton cmd待发药颜色 
            Caption         =   "字体颜色"
            Height          =   350
            Left            =   4200
            TabIndex        =   38
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmd待发药字体 
            Caption         =   "字体设置"
            Height          =   350
            Left            =   240
            TabIndex        =   36
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txt待发药_顶 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2880
            TabIndex        =   33
            Top             =   555
            Width           =   1695
         End
         Begin VB.TextBox txt待发药_左 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   31
            Top             =   555
            Width           =   1695
         End
         Begin VB.CheckBox chk显示待发药 
            Caption         =   "显示待发药"
            Height          =   180
            Left            =   240
            TabIndex        =   28
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "行数："
            Height          =   180
            Index           =   26
            Left            =   3720
            TabIndex        =   46
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "行高："
            Height          =   180
            Index           =   25
            Left            =   1800
            TabIndex        =   44
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "列宽："
            Height          =   180
            Index           =   24
            Left            =   240
            TabIndex        =   42
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "表格"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   23
            Left            =   240
            TabIndex        =   40
            Top             =   1920
            Width           =   360
         End
         Begin VB.Shape shp待发药颜色 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   5280
            Top             =   1350
            Width           =   375
         End
         Begin VB.Label lbl待发药字体 
            AutoSize        =   -1  'True
            Caption         =   "微软雅黑;12"
            Height          =   180
            Left            =   1350
            TabIndex        =   37
            Top             =   1410
            Width           =   990
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "字体"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   22
            Left            =   240
            TabIndex        =   35
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "顶："
            Height          =   180
            Index           =   21
            Left            =   2520
            TabIndex        =   32
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "位置"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   20
            Left            =   240
            TabIndex        =   30
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "左："
            Height          =   180
            Index           =   19
            Left            =   240
            TabIndex        =   29
            Top             =   600
            Width           =   360
         End
      End
      Begin VB.Frame fra显示时间 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   84
         Top             =   3180
         Width           =   6015
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   5
            Left            =   240
            TabIndex        =   91
            Top             =   960
            Width           =   5595
         End
         Begin VB.CommandButton cmd时间颜色 
            Caption         =   "字体颜色"
            Height          =   350
            Left            =   4200
            TabIndex        =   95
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmd时间字体 
            Caption         =   "字体设置"
            Height          =   350
            Left            =   240
            TabIndex        =   93
            Top             =   1320
            Width           =   975
         End
         Begin VB.CheckBox chk显示时间 
            Caption         =   "显示时间"
            Height          =   180
            Left            =   240
            TabIndex        =   85
            Top             =   0
            Width           =   1095
         End
         Begin VB.TextBox txt时间_左 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   88
            Top             =   555
            Width           =   1695
         End
         Begin VB.TextBox txt时间_顶 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2880
            TabIndex        =   90
            Top             =   555
            Width           =   1695
         End
         Begin VB.Shape shp时间颜色 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   5280
            Top             =   1350
            Width           =   375
         End
         Begin VB.Label lbl时间字体 
            AutoSize        =   -1  'True
            Caption         =   "微软雅黑;12"
            Height          =   180
            Left            =   1320
            TabIndex        =   94
            Top             =   1410
            Width           =   990
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "字体"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   17
            Left            =   240
            TabIndex        =   92
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "顶："
            Height          =   180
            Index           =   16
            Left            =   2520
            TabIndex        =   89
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "位置"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   15
            Left            =   240
            TabIndex        =   87
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "左："
            Height          =   180
            Index           =   14
            Left            =   240
            TabIndex        =   86
            Top             =   600
            Width           =   360
         End
      End
      Begin VB.Frame frm提示 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   1
         Top             =   420
         Width           =   6015
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   8
            Left            =   240
            TabIndex        =   13
            Top             =   1800
            Width           =   4755
         End
         Begin VB.TextBox txt提示_内容 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   240
            TabIndex        =   15
            Top             =   2160
            Width           =   4695
         End
         Begin VB.TextBox txt提示_顶 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2880
            TabIndex        =   8
            Top             =   555
            Width           =   1695
         End
         Begin VB.TextBox txt提示_左 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   7
            Top             =   555
            Width           =   1695
         End
         Begin VB.CheckBox chk显示提示 
            Caption         =   "显示提示"
            Height          =   180
            Left            =   240
            TabIndex        =   2
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmd提示字体 
            Caption         =   "字体设置"
            Height          =   350
            Left            =   240
            TabIndex        =   10
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmd提示颜色 
            Caption         =   "字体颜色"
            Height          =   350
            Left            =   4200
            TabIndex        =   12
            Top             =   1320
            Width           =   975
         End
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   4
            Left            =   240
            TabIndex        =   9
            Top             =   960
            Width           =   5595
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "内容"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   18
            Left            =   240
            TabIndex        =   14
            Top             =   1920
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "左："
            Height          =   180
            Index           =   13
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "位置"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   12
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "顶："
            Height          =   180
            Index           =   11
            Left            =   2520
            TabIndex        =   4
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "字体"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   10
            Left            =   240
            TabIndex        =   3
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lbl提示字体 
            AutoSize        =   -1  'True
            Caption         =   "微软雅黑;12"
            Height          =   180
            Left            =   1350
            TabIndex        =   11
            Top             =   1410
            Width           =   990
         End
         Begin VB.Shape shp提示颜色 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   5280
            Top             =   1350
            Width           =   375
         End
      End
      Begin VB.Frame fra药房 
         Caption         =   " 药房 "
         Height          =   1935
         Left            =   -74880
         TabIndex        =   16
         Top             =   420
         Width           =   6015
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   0
            Left            =   240
            TabIndex        =   22
            Top             =   960
            Width           =   5595
         End
         Begin VB.CommandButton cmd药房颜色 
            Caption         =   "字体颜色"
            Height          =   350
            Left            =   4200
            TabIndex        =   26
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmd药房字体 
            Caption         =   "字体设置"
            Height          =   350
            Left            =   240
            TabIndex        =   24
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txt药房_顶 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2880
            TabIndex        =   21
            Top             =   555
            Width           =   1695
         End
         Begin VB.TextBox txt药房_左 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   19
            Top             =   555
            Width           =   1695
         End
         Begin VB.Shape shp药房颜色 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   5280
            Top             =   1350
            Width           =   375
         End
         Begin VB.Label lbl药房字体 
            AutoSize        =   -1  'True
            Caption         =   "微软雅黑;12"
            Height          =   180
            Left            =   1350
            TabIndex        =   25
            Top             =   1410
            Width           =   990
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "字体"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   6
            Left            =   240
            TabIndex        =   23
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "顶："
            Height          =   180
            Index           =   5
            Left            =   2520
            TabIndex        =   20
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "位置"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   4
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "左："
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   360
         End
      End
      Begin VB.Frame frm呼叫区域 
         Caption         =   " 呼叫"
         Height          =   3915
         Left            =   -74880
         TabIndex        =   54
         Top             =   2460
         Width           =   6015
         Begin VB.Frame fraLine 
            Height          =   840
            Index           =   3
            Left            =   240
            TabIndex        =   70
            Top             =   2880
            Width           =   5595
            Begin VB.CheckBox chk呼叫窗口单独设置 
               Caption         =   "窗口单独设置"
               Height          =   180
               Left            =   240
               TabIndex        =   71
               Top             =   0
               Width           =   1455
            End
            Begin VB.CommandButton cmd呼叫颜色_窗口 
               Caption         =   "字体颜色"
               Height          =   350
               Left            =   3960
               TabIndex        =   74
               Top             =   360
               Width           =   975
            End
            Begin VB.CommandButton cmd呼叫字体_窗口 
               Caption         =   "字体设置"
               Height          =   350
               Left            =   240
               TabIndex        =   72
               Top             =   360
               Width           =   975
            End
            Begin VB.Shape shp呼叫颜色_窗口 
               BackColor       =   &H00FFFFFF&
               FillStyle       =   0  'Solid
               Height          =   300
               Left            =   5040
               Top             =   390
               Width           =   375
            End
            Begin VB.Label lbl呼叫字体_窗口 
               AutoSize        =   -1  'True
               Caption         =   "微软雅黑;12"
               Height          =   180
               Left            =   1350
               TabIndex        =   73
               Top             =   450
               Width           =   990
            End
         End
         Begin VB.Frame fraLine 
            Height          =   960
            Index           =   2
            Left            =   240
            TabIndex        =   65
            Top             =   1800
            Width           =   5595
            Begin VB.CheckBox chk呼叫姓名单独设置 
               Caption         =   "姓名单独设置"
               Height          =   180
               Left            =   240
               TabIndex        =   66
               Top             =   0
               Width           =   1455
            End
            Begin VB.CommandButton cmd呼叫颜色_姓名 
               Caption         =   "字体颜色"
               Height          =   350
               Left            =   3960
               TabIndex        =   69
               Top             =   360
               Width           =   975
            End
            Begin VB.CommandButton cmd呼叫字体_姓名 
               Caption         =   "字体设置"
               Height          =   350
               Left            =   240
               TabIndex        =   67
               Top             =   360
               Width           =   975
            End
            Begin VB.Shape shp呼叫颜色_姓名 
               BackColor       =   &H00FFFFFF&
               FillStyle       =   0  'Solid
               Height          =   300
               Left            =   5040
               Top             =   390
               Width           =   375
            End
            Begin VB.Label lbl呼叫字体_姓名 
               AutoSize        =   -1  'True
               Caption         =   "微软雅黑;12"
               Height          =   180
               Left            =   1350
               TabIndex        =   68
               Top             =   450
               Width           =   990
            End
         End
         Begin VB.Frame fraLine 
            Height          =   35
            Index           =   1
            Left            =   240
            TabIndex        =   60
            Top             =   960
            Width           =   5595
         End
         Begin VB.TextBox txt呼叫_顶 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2880
            TabIndex        =   59
            Top             =   555
            Width           =   1695
         End
         Begin VB.TextBox txt呼叫_左 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   57
            Top             =   555
            Width           =   1695
         End
         Begin VB.CommandButton cmd呼叫颜色_通用 
            Caption         =   "字体颜色"
            Height          =   350
            Left            =   4200
            TabIndex        =   64
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmd呼叫字体_通用 
            Caption         =   "字体设置"
            Height          =   350
            Left            =   240
            TabIndex        =   62
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "字体"
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   9
            Left            =   240
            TabIndex        =   61
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "顶："
            Height          =   180
            Index           =   8
            Left            =   2520
            TabIndex        =   58
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "左："
            Height          =   180
            Index           =   7
            Left            =   240
            TabIndex        =   56
            Top             =   600
            Width           =   360
         End
         Begin VB.Shape shp呼叫颜色_通用 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   5280
            Top             =   1350
            Width           =   375
         End
         Begin VB.Label lbl呼叫字体_通用 
            AutoSize        =   -1  'True
            Caption         =   "微软雅黑;12"
            Height          =   180
            Left            =   1350
            TabIndex        =   63
            Top             =   1410
            Width           =   990
         End
         Begin VB.Label lbl呼叫_位置 
            AutoSize        =   -1  'True
            Caption         =   "位置"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   240
            TabIndex        =   55
            Top             =   240
            Width           =   360
         End
      End
      Begin VB.Frame frmRect 
         Caption         =   " 液晶屏位置（分辨率为单位）"
         Height          =   1150
         Left            =   120
         TabIndex        =   75
         Top             =   2940
         Width           =   6015
         Begin VB.TextBox txt液晶屏_左 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   76
            Top             =   310
            Width           =   1935
         End
         Begin VB.TextBox txt液晶屏_顶 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   3720
            TabIndex        =   79
            Top             =   310
            Width           =   1935
         End
         Begin VB.TextBox txt液晶屏_宽度 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   600
            TabIndex        =   80
            Top             =   710
            Width           =   1935
         End
         Begin VB.TextBox txt液晶屏_高度 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   3720
            TabIndex        =   82
            Top             =   710
            Width           =   1935
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "顶："
            Height          =   180
            Index           =   3
            Left            =   3360
            TabIndex        =   78
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "左："
            Height          =   180
            Index           =   0
            Left            =   285
            TabIndex        =   77
            Top             =   345
            Width           =   360
         End
         Begin VB.Label lbl标签 
            AutoSize        =   -1  'True
            Caption         =   "宽度："
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   81
            Top             =   750
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "高度："
            Height          =   180
            Index           =   3
            Left            =   3240
            TabIndex        =   83
            Top             =   750
            Width           =   540
         End
      End
      Begin VB.Frame frm背景图片 
         Caption         =   " 背景图片 "
         Height          =   735
         Left            =   120
         TabIndex        =   50
         Top             =   2100
         Width           =   6015
         Begin VB.TextBox txt图片位置 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   300
            Width           =   4695
         End
         Begin VB.CommandButton cmd图片位置 
            Caption         =   "…"
            Height          =   270
            Left            =   5520
            TabIndex        =   53
            TabStop         =   0   'False
            Tag             =   "分类"
            ToolTipText     =   "按*打开选择器"
            Top             =   315
            Width           =   270
         End
         Begin VB.Label lbl背景图片 
            AutoSize        =   -1  'True
            Caption         =   "位置"
            Height          =   180
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   360
         End
      End
      Begin VB.Frame fra显示模式 
         Height          =   1575
         Left            =   120
         TabIndex        =   47
         Top             =   420
         Width           =   6015
         Begin VB.CheckBox chk多窗口模式 
            Caption         =   "多窗口模式"
            Height          =   180
            Left            =   240
            TabIndex        =   48
            Top             =   0
            Width           =   1215
         End
         Begin VB.ListBox lst发药窗口 
            Appearance      =   0  'Flat
            Columns         =   3
            ForeColor       =   &H80000012&
            Height          =   1080
            IMEMode         =   3  'DISABLE
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   49
            Top             =   360
            Width           =   5520
         End
      End
   End
   Begin MSComDlg.CommonDialog cdl交互 
      Left            =   120
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmProSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrWins As String      '传入的窗口串
Private mstrReg As String       '本地注册表路径

Public Sub ShowMe(ByVal frmParent As Form, ByVal strWins As String)
    mstrWins = strWins
    
    Me.Show 1, frmParent
End Sub

Private Sub chk多窗口模式_Click()
    lst发药窗口.Enabled = chk多窗口模式.Value
End Sub

Private Sub chk呼叫窗口单独设置_Click()
    cmd呼叫字体_窗口.Enabled = chk呼叫窗口单独设置.Value
    cmd呼叫颜色_窗口.Enabled = chk呼叫窗口单独设置.Value
End Sub

Private Sub chk呼叫姓名单独设置_Click()
    cmd呼叫字体_姓名.Enabled = chk呼叫姓名单独设置.Value
    cmd呼叫颜色_姓名.Enabled = chk呼叫姓名单独设置.Value
End Sub

Private Sub chk显示待发药_Click()
    txt待发药_左.Enabled = chk显示待发药.Value
    txt待发药_顶.Enabled = chk显示待发药.Value
    cmd待发药字体.Enabled = chk显示待发药.Value
    cmd待发药颜色.Enabled = chk显示待发药.Value
    txt待发药_列宽.Enabled = chk显示待发药.Value
    txt待发药_行高.Enabled = chk显示待发药.Value
    txt待发药_行数.Enabled = chk显示待发药.Value
End Sub

Private Sub chk显示时间_Click()
    txt时间_左.Enabled = chk显示时间.Value
    txt时间_顶.Enabled = chk显示时间.Value
    cmd时间字体.Enabled = chk显示时间.Value
    cmd时间颜色.Enabled = chk显示时间.Value
End Sub

Private Sub chk显示提示_Click()
    txt提示_左.Enabled = chk显示提示.Value
    txt提示_顶.Enabled = chk显示提示.Value
    cmd提示字体.Enabled = chk显示提示.Value
    cmd提示颜色.Enabled = chk显示提示.Value
    txt提示_内容.Enabled = chk显示提示.Value
End Sub

Private Sub chk显示已过号_Click()
    txt已过号_左.Enabled = chk显示已过号.Value
    txt已过号_顶.Enabled = chk显示已过号.Value
    cmd已过号字体.Enabled = chk显示已过号.Value
    cmd已过号颜色.Enabled = chk显示已过号.Value
    txt已过号_列宽.Enabled = chk显示已过号.Value
    txt已过号_行高.Enabled = chk显示已过号.Value
    txt已过号_行数.Enabled = chk显示已过号.Value
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    '功能：保存相关设置到注册表
    Dim strWin As String
    Dim i As Integer
    
    SaveSetting "ZLSOFT", mstrReg, "窗口模式", chk多窗口模式.Value
    
    For i = 0 To Me.lst发药窗口.ListCount - 1
        If lst发药窗口.Selected(i) Then
            strWin = strWin & IIf(strWin = "", "", ",") & lst发药窗口.List(i)
        End If
    Next
    SaveSetting "ZLSOFT", mstrReg, "多窗口", strWin
    
    SaveSetting "ZLSOFT", mstrReg, "图片位置", txt图片位置.Text
    
    SaveSetting "ZLSOFT", mstrReg, "液晶屏_左", txt液晶屏_左.Text
    SaveSetting "ZLSOFT", mstrReg, "液晶屏_顶", txt液晶屏_顶.Text
    SaveSetting "ZLSOFT", mstrReg, "液晶屏_宽度", txt液晶屏_宽度.Text
    SaveSetting "ZLSOFT", mstrReg, "液晶屏_高度", txt液晶屏_高度.Text
    
    SaveSetting "ZLSOFT", mstrReg, "数据轮询时间", txt数据轮询时间.Text
    SaveSetting "ZLSOFT", mstrReg, "呼叫显示时间", txt呼叫显示时间.Text
    
    SaveSetting "ZLSOFT", mstrReg, "药房_左", txt药房_左.Text
    SaveSetting "ZLSOFT", mstrReg, "药房_顶", txt药房_顶.Text
    SaveSetting "ZLSOFT", mstrReg, "药房颜色", shp药房颜色.FillColor
    
    SaveSetting "ZLSOFT", mstrReg, "呼叫_左", txt呼叫_左.Text
    SaveSetting "ZLSOFT", mstrReg, "呼叫_顶", txt呼叫_顶.Text
    SaveSetting "ZLSOFT", mstrReg, "呼叫姓名单独设置", chk呼叫姓名单独设置.Value
    SaveSetting "ZLSOFT", mstrReg, "呼叫窗口单独设置", chk呼叫窗口单独设置.Value
    SaveSetting "ZLSOFT", mstrReg, "呼叫颜色_通用", shp呼叫颜色_通用.FillColor
    SaveSetting "ZLSOFT", mstrReg, "呼叫颜色_姓名", shp呼叫颜色_姓名.FillColor
    SaveSetting "ZLSOFT", mstrReg, "呼叫颜色_窗口", shp呼叫颜色_窗口.FillColor
    
    SaveSetting "ZLSOFT", mstrReg, "显示待发药", chk显示待发药.Value
    SaveSetting "ZLSOFT", mstrReg, "待发药_左", txt待发药_左.Text
    SaveSetting "ZLSOFT", mstrReg, "待发药_顶", txt待发药_顶.Text
    SaveSetting "ZLSOFT", mstrReg, "待发药_列宽", txt待发药_列宽.Text
    SaveSetting "ZLSOFT", mstrReg, "待发药_行高", txt待发药_行高.Text
    SaveSetting "ZLSOFT", mstrReg, "待发药_行数", txt待发药_行数.Text
    SaveSetting "ZLSOFT", mstrReg, "待发药颜色", shp待发药颜色.FillColor
    
    SaveSetting "ZLSOFT", mstrReg, "显示已过号", chk显示已过号.Value
    SaveSetting "ZLSOFT", mstrReg, "已过号_左", txt已过号_左.Text
    SaveSetting "ZLSOFT", mstrReg, "已过号_顶", txt已过号_顶.Text
    SaveSetting "ZLSOFT", mstrReg, "已过号_列宽", txt已过号_列宽.Text
    SaveSetting "ZLSOFT", mstrReg, "已过号_行高", txt已过号_行高.Text
    SaveSetting "ZLSOFT", mstrReg, "已过号_行数", txt已过号_行数.Text
    SaveSetting "ZLSOFT", mstrReg, "已过号颜色", shp已过号颜色.FillColor
    
    SaveSetting "ZLSOFT", mstrReg, "显示提示", chk显示提示.Value
    SaveSetting "ZLSOFT", mstrReg, "提示_左", txt提示_左.Text
    SaveSetting "ZLSOFT", mstrReg, "提示_顶", txt提示_顶.Text
    SaveSetting "ZLSOFT", mstrReg, "提示_内容", txt提示_内容.Text
    SaveSetting "ZLSOFT", mstrReg, "提示颜色", shp提示颜色.FillColor
    
    SaveSetting "ZLSOFT", mstrReg, "显示时间", chk显示时间.Value
    SaveSetting "ZLSOFT", mstrReg, "时间_左", txt时间_左.Text
    SaveSetting "ZLSOFT", mstrReg, "时间_顶", txt时间_顶.Text
    SaveSetting "ZLSOFT", mstrReg, "时间颜色", shp时间颜色.FillColor
    
    Unload Me
End Sub

Private Sub cmd待发药颜色_Click()
    cdl交互.Color = shp药房颜色.FillColor
    cdl交互.ShowColor
    shp待发药颜色.FillColor = cdl交互.Color
End Sub

Private Sub cmd呼叫颜色_窗口_Click()
    cdl交互.Color = shp药房颜色.FillColor
    cdl交互.ShowColor
    shp呼叫颜色_窗口.FillColor = cdl交互.Color
End Sub

Private Sub cmd呼叫颜色_通用_Click()
    cdl交互.Color = shp药房颜色.FillColor
    cdl交互.ShowColor
    shp呼叫颜色_通用.FillColor = cdl交互.Color
End Sub

Private Sub cmd呼叫颜色_姓名_Click()
    cdl交互.Color = shp药房颜色.FillColor
    cdl交互.ShowColor
    shp呼叫颜色_姓名.FillColor = cdl交互.Color
End Sub

Private Sub cmd呼叫字体_窗口_Click()
    On Error GoTo errHandle
    
    cdl交互.Flags = cdlCFBoth
    cdl交互.CancelError = False  '把点取消当作错误处理
    
    cdl交互.FontName = GetSetting("ZLSOFT", mstrReg, "呼叫字体_窗口", "微软雅黑")
    cdl交互.FontBold = GetSetting("ZLSOFT", mstrReg, "呼叫窗口字体_粗体", "False")
    cdl交互.FontItalic = GetSetting("ZLSOFT", mstrReg, "呼叫窗口字体_斜体", "False")
    cdl交互.FontSize = GetSetting("ZLSOFT", mstrReg, "呼叫窗口字体_字号", "12")
    
    cdl交互.ShowFont

    '设置字体
    SaveSetting "ZLSOFT", mstrReg, "呼叫字体_窗口", cdl交互.FontName
    SaveSetting "ZLSOFT", mstrReg, "呼叫窗口字体_粗体", cdl交互.FontBold
    SaveSetting "ZLSOFT", mstrReg, "呼叫窗口字体_斜体", cdl交互.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "呼叫窗口字体_字号", cdl交互.FontSize
    
    lbl呼叫字体_窗口.Caption = cdl交互.FontName & "," & IIf(cdl交互.FontBold, "粗体,", "") & IIf(cdl交互.FontItalic, "斜体,", "") & cdl交互.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmd呼叫字体_通用_Click()
    On Error GoTo errHandle
    
    cdl交互.Flags = cdlCFBoth
    cdl交互.CancelError = False  '把点取消当作错误处理
    
    cdl交互.FontName = GetSetting("ZLSOFT", mstrReg, "呼叫字体_通用", "微软雅黑")
    cdl交互.FontBold = GetSetting("ZLSOFT", mstrReg, "呼叫通用字体_粗体", "False")
    cdl交互.FontItalic = GetSetting("ZLSOFT", mstrReg, "呼叫通用字体_斜体", "False")
    cdl交互.FontSize = GetSetting("ZLSOFT", mstrReg, "呼叫通用字体_字号", "12")
    
    cdl交互.ShowFont

    '设置字体
    SaveSetting "ZLSOFT", mstrReg, "呼叫字体_通用", cdl交互.FontName
    SaveSetting "ZLSOFT", mstrReg, "呼叫通用字体_粗体", cdl交互.FontBold
    SaveSetting "ZLSOFT", mstrReg, "呼叫通用字体_斜体", cdl交互.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "呼叫通用字体_字号", cdl交互.FontSize
    
    lbl呼叫字体_通用.Caption = cdl交互.FontName & "," & IIf(cdl交互.FontBold, "粗体,", "") & IIf(cdl交互.FontItalic, "斜体,", "") & cdl交互.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmd呼叫字体_姓名_Click()
    On Error GoTo errHandle
    
    cdl交互.Flags = cdlCFBoth
    cdl交互.CancelError = False  '把点取消当作错误处理
    
    cdl交互.FontName = GetSetting("ZLSOFT", mstrReg, "呼叫字体_姓名", "微软雅黑")
    cdl交互.FontBold = GetSetting("ZLSOFT", mstrReg, "呼叫姓名字体_粗体", "False")
    cdl交互.FontItalic = GetSetting("ZLSOFT", mstrReg, "呼叫姓名字体_斜体", "False")
    cdl交互.FontSize = GetSetting("ZLSOFT", mstrReg, "呼叫姓名字体_字号", "12")
    
    cdl交互.ShowFont

    '设置字体
    SaveSetting "ZLSOFT", mstrReg, "呼叫字体_姓名", cdl交互.FontName
    SaveSetting "ZLSOFT", mstrReg, "呼叫姓名字体_粗体", cdl交互.FontBold
    SaveSetting "ZLSOFT", mstrReg, "呼叫姓名字体_斜体", cdl交互.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "呼叫姓名字体_字号", cdl交互.FontSize
    
    lbl呼叫字体_姓名.Caption = cdl交互.FontName & "," & IIf(cdl交互.FontBold, "粗体,", "") & IIf(cdl交互.FontItalic, "斜体,", "") & cdl交互.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmd时间颜色_Click()
    cdl交互.Color = shp药房颜色.FillColor
    cdl交互.ShowColor
    shp时间颜色.FillColor = cdl交互.Color
End Sub

Private Sub cmd时间字体_Click()
    On Error GoTo errHandle
    
    cdl交互.Flags = cdlCFBoth
    cdl交互.CancelError = False  '把点取消当作错误处理
    
    cdl交互.FontName = GetSetting("ZLSOFT", mstrReg, "时间字体", "微软雅黑")
    cdl交互.FontBold = GetSetting("ZLSOFT", mstrReg, "时间粗体", "False")
    cdl交互.FontItalic = GetSetting("ZLSOFT", mstrReg, "时间斜体", "False")
    cdl交互.FontSize = GetSetting("ZLSOFT", mstrReg, "时间字号", "12")
    
    cdl交互.ShowFont

    '设置字体
    SaveSetting "ZLSOFT", mstrReg, "时间字体", cdl交互.FontName
    SaveSetting "ZLSOFT", mstrReg, "时间粗体", cdl交互.FontBold
    SaveSetting "ZLSOFT", mstrReg, "时间斜体", cdl交互.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "时间字号", cdl交互.FontSize
    
    lbl时间字体.Caption = cdl交互.FontName & "," & IIf(cdl交互.FontBold, "粗体,", "") & IIf(cdl交互.FontItalic, "斜体,", "") & cdl交互.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmd提示颜色_Click()
    cdl交互.Color = shp药房颜色.FillColor
    cdl交互.ShowColor
    shp提示颜色.FillColor = cdl交互.Color
End Sub

Private Sub cmd提示字体_Click()
    On Error GoTo errHandle
    
    cdl交互.Flags = cdlCFBoth
    cdl交互.CancelError = False  '把点取消当作错误处理
    
    cdl交互.FontName = GetSetting("ZLSOFT", mstrReg, "提示字体", "微软雅黑")
    cdl交互.FontBold = GetSetting("ZLSOFT", mstrReg, "提示粗体", "False")
    cdl交互.FontItalic = GetSetting("ZLSOFT", mstrReg, "提示斜体", "False")
    cdl交互.FontSize = GetSetting("ZLSOFT", mstrReg, "提示字号", "12")
    
    cdl交互.ShowFont

    '设置字体
    SaveSetting "ZLSOFT", mstrReg, "提示字体", cdl交互.FontName
    SaveSetting "ZLSOFT", mstrReg, "提示粗体", cdl交互.FontBold
    SaveSetting "ZLSOFT", mstrReg, "提示斜体", cdl交互.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "提示字号", cdl交互.FontSize
    
    lbl提示字体.Caption = cdl交互.FontName & "," & IIf(cdl交互.FontBold, "粗体,", "") & IIf(cdl交互.FontItalic, "斜体,", "") & cdl交互.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmd图片位置_Click()
    With cdl交互
        .CancelError = True
        .Filter = "Pictures (*.jpg)|*.jpg"
        
        On Error Resume Next
        .ShowOpen
        
        If err <> 0 Then
            '没选中文件
            err.Clear
        Else
            txt图片位置.Text = .FileName
        End If
    End With
End Sub

Private Sub cmd药房颜色_Click()
    cdl交互.Color = shp药房颜色.FillColor
    cdl交互.ShowColor
    shp药房颜色.FillColor = cdl交互.Color
End Sub

Private Sub cmd药房字体_Click()
    On Error GoTo errHandle
    
    cdl交互.Flags = cdlCFBoth
    cdl交互.CancelError = False  '把点取消当作错误处理
    
    cdl交互.FontName = GetSetting("ZLSOFT", mstrReg, "药房字体", "微软雅黑")
    cdl交互.FontBold = GetSetting("ZLSOFT", mstrReg, "药房粗体", "False")
    cdl交互.FontItalic = GetSetting("ZLSOFT", mstrReg, "药房斜体", "False")
    cdl交互.FontSize = GetSetting("ZLSOFT", mstrReg, "药房字号", "12")
    
    cdl交互.ShowFont

    '设置字体
    SaveSetting "ZLSOFT", mstrReg, "药房字体", cdl交互.FontName
    SaveSetting "ZLSOFT", mstrReg, "药房粗体", cdl交互.FontBold
    SaveSetting "ZLSOFT", mstrReg, "药房斜体", cdl交互.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "药房字号", cdl交互.FontSize
    
    lbl药房字体.Caption = cdl交互.FontName & "," & IIf(cdl交互.FontBold, "粗体,", "") & IIf(cdl交互.FontItalic, "斜体,", "") & cdl交互.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmd待发药字体_Click()
    On Error GoTo errHandle
    
    cdl交互.Flags = cdlCFBoth
    cdl交互.CancelError = False  '把点取消当作错误处理
    
    cdl交互.FontName = GetSetting("ZLSOFT", mstrReg, "待发药字体", "微软雅黑")
    cdl交互.FontBold = GetSetting("ZLSOFT", mstrReg, "待发药粗体", "False")
    cdl交互.FontItalic = GetSetting("ZLSOFT", mstrReg, "待发药斜体", "False")
    cdl交互.FontSize = GetSetting("ZLSOFT", mstrReg, "待发药字号", "12")
    
    cdl交互.ShowFont

    '设置字体
    SaveSetting "ZLSOFT", mstrReg, "待发药字体", cdl交互.FontName
    SaveSetting "ZLSOFT", mstrReg, "待发药粗体", cdl交互.FontBold
    SaveSetting "ZLSOFT", mstrReg, "待发药斜体", cdl交互.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "待发药字号", cdl交互.FontSize
    
    lbl待发药字体.Caption = cdl交互.FontName & "," & IIf(cdl交互.FontBold, "粗体,", "") & IIf(cdl交互.FontItalic, "斜体,", "") & cdl交互.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmd已过号颜色_Click()
    cdl交互.Color = shp药房颜色.FillColor
    cdl交互.ShowColor
    shp已过号颜色.FillColor = cdl交互.Color
End Sub

Private Sub cmd已过号字体_Click()
    On Error GoTo errHandle
    
    cdl交互.Flags = cdlCFBoth
    cdl交互.CancelError = False  '把点取消当作错误处理
    
    cdl交互.FontName = GetSetting("ZLSOFT", mstrReg, "已过号字体", "微软雅黑")
    cdl交互.FontBold = GetSetting("ZLSOFT", mstrReg, "已过号粗体", "False")
    cdl交互.FontItalic = GetSetting("ZLSOFT", mstrReg, "已过号斜体", "False")
    cdl交互.FontSize = GetSetting("ZLSOFT", mstrReg, "已过号字号", "12")
    
    cdl交互.ShowFont

    '设置字体
    SaveSetting "ZLSOFT", mstrReg, "已过号字体", cdl交互.FontName
    SaveSetting "ZLSOFT", mstrReg, "已过号粗体", cdl交互.FontBold
    SaveSetting "ZLSOFT", mstrReg, "已过号斜体", cdl交互.FontItalic
    SaveSetting "ZLSOFT", mstrReg, "已过号字号", cdl交互.FontSize
    
    lbl已过号字体.Caption = cdl交互.FontName & "," & IIf(cdl交互.FontBold, "粗体,", "") & IIf(cdl交互.FontItalic, "斜体,", "") & cdl交互.FontSize
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub Form_Load()
    '路径初始化
    mstrReg = "公共模块\药房排队叫号\液晶电视Pro"
    
    '初始化发药窗口
    Call LoadWins
    
    '恢复本地设置
    Call LoadLocalSettings
End Sub

Private Sub LoadWins()
    '功能：初始化发药窗口
    Dim i As Integer
    
    For i = 0 To UBound(Split(mstrWins, ","))
        Me.lst发药窗口.AddItem Split(mstrWins, ",")(i)
    Next
End Sub

Private Sub LoadLocalSettings()
    '功能：恢复本地设置
    Dim strWin As String
    Dim i As Integer
    
    '恢复窗口模式
    chk多窗口模式.Value = IIf(Val(GetSetting("ZLSOFT", mstrReg, "窗口模式", "0")) = 1, 1, 0)
    
    '恢复选中发药窗口
    strWin = GetSetting("ZLSOFT", mstrReg, "多窗口", "")
    
    If strWin <> "" Then
        For i = 0 To Me.lst发药窗口.ListCount - 1
            If InStr(1, strWin, lst发药窗口.List(i)) > 0 Then
                lst发药窗口.Selected(i) = True
            End If
        Next
    End If
    
    lst发药窗口.Enabled = (Val(GetSetting("ZLSOFT", mstrReg, "窗口模式", "0")) = 1)
    
    '恢复背景图片
    txt图片位置.Text = GetSetting("ZLSOFT", mstrReg, "图片位置", "")
    
    '恢复液晶屏位置
    txt液晶屏_左.Text = Val(GetSetting("ZLSOFT", mstrReg, "液晶屏_左", "0"))
    txt液晶屏_顶.Text = Val(GetSetting("ZLSOFT", mstrReg, "液晶屏_顶", "0"))
    txt液晶屏_宽度.Text = Val(GetSetting("ZLSOFT", mstrReg, "液晶屏_宽度", "1024"))
    txt液晶屏_高度.Text = Val(GetSetting("ZLSOFT", mstrReg, "液晶屏_高度", "768"))
    
    '恢复数据刷新
    txt数据轮询时间.Text = Val(GetSetting("ZLSOFT", mstrReg, "数据轮询时间", "1"))
    If Val(txt数据轮询时间.Text) < 1 Then
        txt数据轮询时间.Text = 1
    ElseIf Val(txt数据轮询时间.Text) > 60 Then
        txt数据轮询时间.Text = 60
    End If
    
    '恢复显示刷新
    txt呼叫显示时间.Text = Val(GetSetting("ZLSOFT", mstrReg, "呼叫显示时间", "1"))
    If Val(txt呼叫显示时间.Text) < 1 Then
        txt呼叫显示时间.Text = 1
    ElseIf Val(txt呼叫显示时间.Text) > 60 Then
        txt呼叫显示时间.Text = 60
    End If
    
    '恢复药房设置
    txt药房_左.Text = Val(GetSetting("ZLSOFT", mstrReg, "药房_左", "0"))
    txt药房_顶.Text = Val(GetSetting("ZLSOFT", mstrReg, "药房_顶", "0"))
    shp药房颜色.FillColor = GetSetting("ZLSOFT", mstrReg, "药房颜色", vbBlack)
    
    lbl药房字体.Caption = GetSetting("ZLSOFT", mstrReg, "药房字体", "微软雅黑")
    lbl药房字体.Caption = lbl药房字体.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "药房粗体", "False"), ";粗体", "")
    lbl药房字体.Caption = lbl药房字体.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "药房斜体", "False"), ";斜体", "")
    lbl药房字体.Caption = lbl药房字体.Caption & IIf(lbl药房字体.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "药房字号", "12")
    
    '恢复呼叫设置
    txt呼叫_左.Text = Val(GetSetting("ZLSOFT", mstrReg, "呼叫_左", "0"))
    txt呼叫_顶.Text = Val(GetSetting("ZLSOFT", mstrReg, "呼叫_顶", "0"))
    shp呼叫颜色_通用.FillColor = GetSetting("ZLSOFT", mstrReg, "呼叫颜色_通用", vbBlack)
    
    lbl呼叫字体_通用.Caption = GetSetting("ZLSOFT", mstrReg, "呼叫字体_通用", "微软雅黑")
    lbl呼叫字体_通用.Caption = lbl呼叫字体_通用.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "呼叫通用字体_粗体", "False"), ";粗体", "")
    lbl呼叫字体_通用.Caption = lbl呼叫字体_通用.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "呼叫通用字体_斜体", "False"), ";斜体", "")
    lbl呼叫字体_通用.Caption = lbl呼叫字体_通用.Caption & IIf(lbl呼叫字体_通用.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "呼叫通用字体_字号", "12")
    
    chk呼叫姓名单独设置.Value = IIf(Val(GetSetting("ZLSOFT", mstrReg, "呼叫姓名单独设置", "0")) = 1, 1, 0)
    shp呼叫颜色_姓名.FillColor = GetSetting("ZLSOFT", mstrReg, "呼叫颜色_姓名", vbBlack)
    
    lbl呼叫字体_姓名.Caption = GetSetting("ZLSOFT", mstrReg, "呼叫字体_姓名", "微软雅黑")
    lbl呼叫字体_姓名.Caption = lbl呼叫字体_姓名.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "呼叫姓名字体_粗体", "False"), ";粗体", "")
    lbl呼叫字体_姓名.Caption = lbl呼叫字体_姓名.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "呼叫姓名字体_斜体", "False"), ";斜体", "")
    lbl呼叫字体_姓名.Caption = lbl呼叫字体_姓名.Caption & IIf(lbl呼叫字体_姓名.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "呼叫姓名字体_字号", "12")
    
    cmd呼叫字体_姓名.Enabled = (Val(GetSetting("ZLSOFT", mstrReg, "呼叫姓名单独设置", "0")) = 1)
    cmd呼叫颜色_姓名.Enabled = cmd呼叫字体_姓名.Enabled
    
    chk呼叫窗口单独设置.Value = IIf(Val(GetSetting("ZLSOFT", mstrReg, "呼叫窗口单独设置", "0")) = 1, 1, 0)
    shp呼叫颜色_窗口.FillColor = GetSetting("ZLSOFT", mstrReg, "呼叫颜色_窗口", vbBlack)
    
    lbl呼叫字体_窗口.Caption = GetSetting("ZLSOFT", mstrReg, "呼叫字体_窗口", "微软雅黑")
    lbl呼叫字体_窗口.Caption = lbl呼叫字体_窗口.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "呼叫窗口字体_粗体", "False"), ";粗体", "")
    lbl呼叫字体_窗口.Caption = lbl呼叫字体_窗口.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "呼叫窗口字体_斜体", "False"), ";斜体", "")
    lbl呼叫字体_窗口.Caption = lbl呼叫字体_窗口.Caption & IIf(lbl呼叫字体_窗口.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "呼叫窗口字体_字号", "12")
    
    cmd呼叫字体_窗口.Enabled = (Val(GetSetting("ZLSOFT", mstrReg, "呼叫窗口单独设置", "0")) = 1)
    cmd呼叫颜色_窗口.Enabled = cmd呼叫字体_窗口.Enabled
    
    '恢复待发药设置
    chk显示待发药.Value = IIf(Val(GetSetting("ZLSOFT", mstrReg, "显示待发药", "1")) = 1, 1, 0)
    txt待发药_左.Text = Val(GetSetting("ZLSOFT", mstrReg, "待发药_左", "0"))
    txt待发药_顶.Text = Val(GetSetting("ZLSOFT", mstrReg, "待发药_顶", "0"))
    txt待发药_列宽.Text = Val(GetSetting("ZLSOFT", mstrReg, "待发药_列宽", "800"))
    txt待发药_行高.Text = Val(GetSetting("ZLSOFT", mstrReg, "待发药_行高", "350"))
    txt待发药_行数.Text = Val(GetSetting("ZLSOFT", mstrReg, "待发药_行数", "5"))
    shp待发药颜色.FillColor = GetSetting("ZLSOFT", mstrReg, "待发药颜色", vbBlack)
    
    lbl待发药字体.Caption = GetSetting("ZLSOFT", mstrReg, "待发药字体", "微软雅黑")
    lbl待发药字体.Caption = lbl待发药字体.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "待发药粗体", "False"), ";粗体", "")
    lbl待发药字体.Caption = lbl待发药字体.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "待发药斜体", "False"), ";斜体", "")
    lbl待发药字体.Caption = lbl待发药字体.Caption & IIf(lbl待发药字体.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "待发药字号", "12")
    
    txt待发药_左.Enabled = (Val(GetSetting("ZLSOFT", mstrReg, "显示待发药", "1")) = 1)
    txt待发药_顶.Enabled = txt待发药_左.Enabled
    cmd待发药字体.Enabled = txt待发药_左.Enabled
    cmd待发药颜色.Enabled = txt待发药_左.Enabled
    txt待发药_列宽.Enabled = txt待发药_左.Enabled
    txt待发药_行高.Enabled = txt待发药_左.Enabled
    txt待发药_行数.Enabled = txt待发药_左.Enabled
    
    '恢复已过号设置
    chk显示已过号.Value = IIf(Val(GetSetting("ZLSOFT", mstrReg, "显示已过号", "1")) = 1, 1, 0)
    txt已过号_左.Text = Val(GetSetting("ZLSOFT", mstrReg, "已过号_左", "0"))
    txt已过号_顶.Text = Val(GetSetting("ZLSOFT", mstrReg, "已过号_顶", "0"))
    txt已过号_列宽.Text = Val(GetSetting("ZLSOFT", mstrReg, "已过号_列宽", "800"))
    txt已过号_行高.Text = Val(GetSetting("ZLSOFT", mstrReg, "已过号_行高", "350"))
    txt已过号_行数.Text = Val(GetSetting("ZLSOFT", mstrReg, "已过号_行数", "5"))
    shp已过号颜色.FillColor = GetSetting("ZLSOFT", mstrReg, "已过号颜色", vbBlack)
    
    lbl已过号字体.Caption = GetSetting("ZLSOFT", mstrReg, "已过号字体", "微软雅黑")
    lbl已过号字体.Caption = lbl已过号字体.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "已过号粗体", "False"), ";粗体", "")
    lbl已过号字体.Caption = lbl已过号字体.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "已过号斜体", "False"), ";斜体", "")
    lbl已过号字体.Caption = lbl已过号字体.Caption & IIf(lbl已过号字体.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "已过号字号", "12")
    
    txt已过号_左.Enabled = (Val(GetSetting("ZLSOFT", mstrReg, "显示已过号", "1")) = 1)
    txt已过号_顶.Enabled = txt已过号_左.Enabled
    cmd已过号字体.Enabled = txt已过号_左.Enabled
    cmd已过号颜色.Enabled = txt已过号_左.Enabled
    txt已过号_列宽.Enabled = txt已过号_左.Enabled
    txt已过号_行高.Enabled = txt已过号_左.Enabled
    txt已过号_行数.Enabled = txt已过号_左.Enabled
    
    '恢复提示设置
    chk显示提示.Value = IIf(Val(GetSetting("ZLSOFT", mstrReg, "显示提示", "1")) = 1, 1, 0)
    txt提示_左.Text = Val(GetSetting("ZLSOFT", mstrReg, "提示_左", "0"))
    txt提示_顶.Text = Val(GetSetting("ZLSOFT", mstrReg, "提示_顶", "0"))
    txt提示_内容.Text = GetSetting("ZLSOFT", mstrReg, "提示_内容", "")
    shp提示颜色.FillColor = GetSetting("ZLSOFT", mstrReg, "提示颜色", vbBlack)
    
    lbl提示字体.Caption = GetSetting("ZLSOFT", mstrReg, "提示字体", "微软雅黑")
    lbl提示字体.Caption = lbl提示字体.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "提示粗体", "False"), ";粗体", "")
    lbl提示字体.Caption = lbl提示字体.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "提示斜体", "False"), ";斜体", "")
    lbl提示字体.Caption = lbl提示字体.Caption & IIf(lbl提示字体.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "提示字号", "12")
    
    txt提示_左.Enabled = (Val(GetSetting("ZLSOFT", mstrReg, "显示提示", "1")) = 1)
    txt提示_顶.Enabled = txt提示_左.Enabled
    cmd提示字体.Enabled = txt提示_左.Enabled
    cmd提示颜色.Enabled = txt提示_左.Enabled
    txt提示_内容.Enabled = txt提示_左.Enabled
    
    '恢复时间设置
    chk显示时间.Value = IIf(Val(GetSetting("ZLSOFT", mstrReg, "显示时间", "1")) = 1, 1, 0)
    txt时间_左.Text = Val(GetSetting("ZLSOFT", mstrReg, "时间_左", "0"))
    txt时间_顶.Text = Val(GetSetting("ZLSOFT", mstrReg, "时间_顶", "0"))
    shp时间颜色.FillColor = GetSetting("ZLSOFT", mstrReg, "时间颜色", vbBlack)
    
    lbl时间字体.Caption = GetSetting("ZLSOFT", mstrReg, "时间字体", "微软雅黑")
    lbl时间字体.Caption = lbl时间字体.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "时间粗体", "False"), ";粗体", "")
    lbl时间字体.Caption = lbl时间字体.Caption & IIf(GetSetting("ZLSOFT", mstrReg, "时间斜体", "False"), ";斜体", "")
    lbl时间字体.Caption = lbl时间字体.Caption & IIf(lbl时间字体.Caption = "", "", ";") & GetSetting("ZLSOFT", mstrReg, "时间字号", "12")
    
    txt时间_左.Enabled = (Val(GetSetting("ZLSOFT", mstrReg, "显示时间", "1")) = 1)
    txt时间_顶.Enabled = txt提示_左.Enabled
    cmd时间字体.Enabled = txt提示_左.Enabled
    cmd时间颜色.Enabled = txt提示_左.Enabled
    
End Sub

Private Sub txt待发药_顶_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt待发药_顶
End Sub

Private Sub txt待发药_顶_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt待发药_行高_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt待发药_行高
End Sub

Private Sub txt待发药_行高_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt待发药_行数_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt待发药_行数
End Sub

Private Sub txt待发药_行数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt待发药_列宽_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt待发药_列宽
End Sub

Private Sub txt待发药_列宽_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt待发药_左_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt待发药_左
End Sub

Private Sub txt待发药_左_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt呼叫_顶_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt呼叫_顶
End Sub

Private Sub txt呼叫_顶_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt呼叫_左_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt呼叫_左
End Sub

Private Sub txt呼叫_左_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt呼叫显示时间_Change()
    If Val(txt呼叫显示时间.Text) < 1 Then
        txt呼叫显示时间.Text = 1
    ElseIf Val(txt呼叫显示时间.Text) > 60 Then
        txt呼叫显示时间.Text = 60
    End If
End Sub

Private Sub txt呼叫显示时间_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt呼叫显示时间
End Sub

Private Sub txt呼叫显示时间_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt时间_顶_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt时间_顶
End Sub

Private Sub txt时间_顶_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt时间_左_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt时间_左
End Sub

Private Sub txt时间_左_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt数据轮询时间_Change()
    If Val(txt数据轮询时间.Text) < 1 Then
        txt数据轮询时间.Text = 1
    ElseIf Val(txt数据轮询时间.Text) > 60 Then
        txt数据轮询时间.Text = 60
    End If
End Sub

Private Sub txt数据轮询时间_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt数据轮询时间
End Sub

Private Sub txt数据轮询时间_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt提示_顶_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt提示_顶
End Sub

Private Sub txt提示_顶_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt提示_内容_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt提示_内容
End Sub

Private Sub txt提示_左_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt提示_左
End Sub

Private Sub txt提示_左_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt药房_顶_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt药房_顶
End Sub

Private Sub txt药房_顶_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt药房_左_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt药房_左
End Sub

Private Sub txt药房_左_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt液晶屏_顶_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt液晶屏_顶
End Sub

Private Sub txt液晶屏_顶_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt液晶屏_高度_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt液晶屏_高度
End Sub

Private Sub txt液晶屏_高度_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt液晶屏_宽度_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt液晶屏_宽度
End Sub

Private Sub txt液晶屏_宽度_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt液晶屏_左_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt液晶屏_左
End Sub

Private Sub txt液晶屏_左_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt已过号_顶_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt已过号_顶
End Sub

Private Sub txt已过号_顶_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt已过号_行高_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt已过号_行高
End Sub

Private Sub txt已过号_行高_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt已过号_行数_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt已过号_行数
End Sub

Private Sub txt已过号_行数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt已过号_列宽_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt已过号_列宽
End Sub

Private Sub txt已过号_列宽_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt已过号_左_GotFocus()
    gobjComLib.zlControl.TxtSelAll txt已过号_左
End Sub

Private Sub txt已过号_左_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

