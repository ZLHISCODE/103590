VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~3.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "CO70B6~1.OCX"
Begin VB.Form frmLabMB 
   Caption         =   "酶标仪"
   ClientHeight    =   7800
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   14550
   Icon            =   "frmLabMB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   14550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSScriptControlCtl.ScriptControl Calc 
      Left            =   2520
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.PictureBox PicList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      FillColor       =   &H00FDD6C6&
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   90
      ScaleHeight     =   3195
      ScaleWidth      =   3195
      TabIndex        =   66
      Top             =   2160
      Width           =   3195
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   2205
         Left            =   60
         TabIndex        =   67
         Top             =   360
         Width           =   1875
         _Version        =   589884
         _ExtentX        =   3307
         _ExtentY        =   3889
         _StockProps     =   0
         BorderStyle     =   2
         AutoColumnSizing=   0   'False
      End
      Begin VB.OptionButton opt过滤 
         BackColor       =   &H00FDD6C6&
         Caption         =   "本年"
         Height          =   255
         Index           =   3
         Left            =   2430
         TabIndex        =   73
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton opt过滤 
         BackColor       =   &H00FDD6C6&
         Caption         =   "本月"
         Height          =   255
         Index           =   2
         Left            =   1640
         TabIndex        =   70
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton opt过滤 
         BackColor       =   &H00FDD6C6&
         Caption         =   "本周"
         Height          =   255
         Index           =   1
         Left            =   850
         TabIndex        =   69
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton opt过滤 
         BackColor       =   &H00FDD6C6&
         Caption         =   "今天"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   68
         Top             =   60
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.ComboBox cbo检验仪器 
      Height          =   300
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   1590
      Width           =   1935
   End
   Begin VB.PictureBox PicMain 
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   7215
      Left            =   3330
      ScaleHeight     =   7215
      ScaleWidth      =   10755
      TabIndex        =   0
      Top             =   30
      Width           =   10755
      Begin VB.Frame fra微孔板 
         Caption         =   "微孔板"
         Height          =   3945
         Left            =   0
         TabIndex        =   45
         Top             =   3150
         Width           =   10485
         Begin VB.TextBox txt最小阴性对照 
            Height          =   285
            Left            =   6870
            TabIndex        =   78
            Top             =   180
            Width           =   585
         End
         Begin VB.OptionButton opt孔选择 
            Caption         =   "质控(QC)"
            Height          =   180
            Index           =   4
            Left            =   4410
            TabIndex        =   77
            Top             =   240
            Width           =   1035
         End
         Begin VB.OptionButton opt孔选择 
            Caption         =   "阳性(PC)"
            Height          =   180
            Index           =   3
            Left            =   3330
            TabIndex        =   59
            Top             =   240
            Width           =   1035
         End
         Begin VB.OptionButton opt孔选择 
            Caption         =   "阴性(NC)"
            Height          =   180
            Index           =   2
            Left            =   2250
            TabIndex        =   58
            Top             =   240
            Width           =   1035
         End
         Begin VB.OptionButton opt孔选择 
            Caption         =   "空白(BC)"
            Height          =   180
            Index           =   1
            Left            =   1170
            TabIndex        =   57
            Top             =   240
            Width           =   1035
         End
         Begin VB.OptionButton opt孔选择 
            Caption         =   "普通(S)"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.Frame fra对照 
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            Height          =   375
            Left            =   60
            TabIndex        =   47
            Top             =   3480
            Width           =   10215
            Begin VB.TextBox txt存放位置 
               Height          =   300
               Left            =   8610
               TabIndex        =   72
               Top             =   90
               Width           =   1305
            End
            Begin VB.TextBox txtCutOff 
               Height          =   300
               Left            =   6570
               Locked          =   -1  'True
               TabIndex        =   51
               Top             =   90
               Width           =   1005
            End
            Begin VB.TextBox txt阳性对照 
               Height          =   300
               Left            =   4830
               Locked          =   -1  'True
               TabIndex        =   50
               Top             =   90
               Width           =   1005
            End
            Begin VB.TextBox txt阴性对照 
               Height          =   300
               Left            =   2850
               Locked          =   -1  'True
               TabIndex        =   49
               Top             =   90
               Width           =   1005
            End
            Begin VB.TextBox txt空白对照 
               Height          =   300
               Left            =   840
               Locked          =   -1  'True
               TabIndex        =   48
               Top             =   90
               Width           =   975
            End
            Begin VB.Label lbl存放位置 
               AutoSize        =   -1  'True
               Caption         =   "存放位置"
               Height          =   180
               Left            =   7800
               TabIndex        =   71
               Top             =   150
               Width           =   720
            End
            Begin VB.Label lblCutOff 
               AutoSize        =   -1  'True
               Caption         =   "CutOff"
               Height          =   180
               Left            =   5970
               TabIndex        =   55
               Top             =   150
               Width           =   540
            End
            Begin VB.Label lbl阳性对照 
               AutoSize        =   -1  'True
               Caption         =   "阳性对照"
               Height          =   180
               Left            =   4050
               TabIndex        =   54
               Top             =   150
               Width           =   720
            End
            Begin VB.Label lbl阴性对照 
               AutoSize        =   -1  'True
               Caption         =   "阴性对照"
               Height          =   180
               Left            =   2070
               TabIndex        =   53
               Top             =   150
               Width           =   720
            End
            Begin VB.Label lbl空白对照 
               AutoSize        =   -1  'True
               Caption         =   "空白对照"
               Height          =   180
               Left            =   60
               TabIndex        =   52
               Top             =   150
               Width           =   720
            End
         End
         Begin VB.CommandButton cmd重置 
            Caption         =   "重置"
            Height          =   285
            Left            =   8940
            TabIndex        =   46
            Top             =   180
            Width           =   1155
         End
         Begin VSFlex8Ctl.VSFlexGrid vsList 
            Height          =   2835
            Left            =   120
            TabIndex        =   60
            Top             =   510
            Width           =   9855
            _cx             =   17383
            _cy             =   5001
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
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
         Begin VB.CheckBox chk阴性对照 
            Caption         =   "阴性对照小于       时按设定值计算"
            Height          =   180
            Left            =   5520
            TabIndex        =   76
            Top             =   240
            Width           =   4065
         End
      End
      Begin VB.Frame fra模板 
         Caption         =   "模板"
         Height          =   615
         Left            =   0
         TabIndex        =   37
         Top             =   2520
         Width           =   10485
         Begin VB.OptionButton opt方向 
            Caption         =   "纵向"
            Height          =   255
            Index           =   1
            Left            =   6120
            TabIndex        =   75
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton opt方向 
            Caption         =   "横向"
            Height          =   255
            Index           =   0
            Left            =   5400
            TabIndex        =   74
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.CommandButton cmd删除模板 
            Caption         =   "删除模板"
            Height          =   285
            Left            =   3930
            TabIndex        =   42
            Top             =   210
            Width           =   1155
         End
         Begin VB.CommandButton cmd保存模板 
            Caption         =   "保存模板"
            Height          =   285
            Left            =   2670
            TabIndex        =   41
            Top             =   210
            Width           =   1155
         End
         Begin VB.ComboBox cbo选择模板 
            Height          =   300
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   210
            Width           =   1605
         End
         Begin VB.TextBox txt开始标本号 
            Height          =   300
            Left            =   7920
            TabIndex        =   39
            ToolTipText     =   " "
            Top             =   210
            Width           =   2025
         End
         Begin VB.CommandButton cmd确定 
            Caption         =   "OK"
            Height          =   285
            Left            =   9960
            TabIndex        =   38
            Top             =   210
            Width           =   375
         End
         Begin VB.Label lbl选择模板 
            AutoSize        =   -1  'True
            Caption         =   "选择模板"
            Height          =   180
            Left            =   150
            TabIndex        =   44
            Top             =   270
            Width           =   720
         End
         Begin VB.Label lbl开始标本号 
            AutoSize        =   -1  'True
            Caption         =   "开始标本号"
            Height          =   180
            Left            =   6990
            TabIndex        =   43
            Top             =   270
            Width           =   900
         End
      End
      Begin VB.Frame fra测量参数 
         Caption         =   "测量参数"
         Height          =   2505
         Left            =   1380
         TabIndex        =   6
         Top             =   0
         Width           =   9105
         Begin VB.TextBox txt试剂批号 
            Height          =   300
            Left            =   7200
            TabIndex        =   80
            Top             =   624
            Width           =   1305
         End
         Begin VB.TextBox txtCutOff公式 
            Height          =   300
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   2040
            Width           =   1635
         End
         Begin VB.TextBox txt振板时间 
            Height          =   300
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   990
            Width           =   1935
         End
         Begin VB.ComboBox cbo波长 
            Height          =   300
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   630
            Width           =   1935
         End
         Begin VB.ComboBox cbo参考波长 
            Height          =   300
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   630
            Width           =   1935
         End
         Begin VB.ComboBox cbo振板频率 
            Height          =   300
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   990
            Width           =   1935
         End
         Begin VB.ComboBox cbo进板方式 
            Height          =   300
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1350
            Width           =   1935
         End
         Begin VB.ComboBox cbo空白形式 
            Height          =   300
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1350
            Width           =   1935
         End
         Begin VB.TextBox txt试剂效期 
            Height          =   300
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   978
            Width           =   1635
         End
         Begin VB.TextBox txt试剂厂商 
            Height          =   300
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1332
            Width           =   1635
         End
         Begin VB.TextBox txt测试方法 
            Height          =   300
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   1686
            Width           =   1635
         End
         Begin VB.TextBox txt测试板号 
            Height          =   300
            Left            =   7200
            TabIndex        =   12
            Top             =   270
            Width           =   1635
         End
         Begin VB.OptionButton opt单板多项 
            Caption         =   "单板多项"
            Height          =   180
            Left            =   4980
            TabIndex        =   11
            Top             =   1770
            Width           =   1065
         End
         Begin VB.OptionButton opt单板单项 
            Caption         =   "单板单项"
            Height          =   180
            Left            =   3870
            TabIndex        =   10
            Top             =   1770
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.TextBox txt弱阳性公式 
            Height          =   300
            Left            =   4080
            TabIndex        =   9
            Top             =   2040
            Width           =   1935
         End
         Begin VB.TextBox txt阳性公式 
            Height          =   300
            Left            =   1050
            TabIndex        =   8
            Top             =   2070
            Width           =   1935
         End
         Begin VB.ComboBox cbo测试项目 
            Height          =   300
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1710
            Width           =   2805
         End
         Begin MSComCtl2.DTPicker dtp测试时间 
            Height          =   300
            Left            =   1050
            TabIndex        =   21
            Top             =   270
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   123142147
            CurrentDate     =   39497
         End
         Begin VB.CommandButton cmdSl 
            Height          =   300
            Left            =   8520
            Picture         =   "frmLabMB.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   615
            Width           =   300
         End
         Begin VB.Label lblCutOff公式 
            AutoSize        =   -1  'True
            Caption         =   "CutOff公式"
            Height          =   180
            Left            =   6210
            TabIndex        =   63
            Top             =   2130
            Width           =   900
         End
         Begin VB.Label lbl测试时间 
            AutoSize        =   -1  'True
            Caption         =   "测试时间"
            Height          =   180
            Left            =   240
            TabIndex        =   36
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lbl波长 
            AutoSize        =   -1  'True
            Caption         =   "波    长"
            Height          =   180
            Left            =   240
            TabIndex        =   35
            Top             =   690
            Width           =   720
         End
         Begin VB.Label lbl参考波长 
            AutoSize        =   -1  'True
            Caption         =   "参考波长"
            Height          =   180
            Left            =   3300
            TabIndex        =   34
            Top             =   690
            Width           =   720
         End
         Begin VB.Label lbl振板频率 
            AutoSize        =   -1  'True
            Caption         =   "振板频率"
            Height          =   180
            Left            =   240
            TabIndex        =   33
            Top             =   1050
            Width           =   720
         End
         Begin VB.Label lbl振板时间 
            AutoSize        =   -1  'True
            Caption         =   "振板时间"
            Height          =   180
            Left            =   3300
            TabIndex        =   32
            Top             =   1050
            Width           =   720
         End
         Begin VB.Label lbl进板方式 
            AutoSize        =   -1  'True
            Caption         =   "进板方式"
            Height          =   180
            Left            =   240
            TabIndex        =   31
            Top             =   1410
            Width           =   720
         End
         Begin VB.Label lbl空白形式 
            AutoSize        =   -1  'True
            Caption         =   "空白形式"
            Height          =   180
            Left            =   3300
            TabIndex        =   30
            Top             =   1410
            Width           =   720
         End
         Begin VB.Label lbl阳性公式 
            AutoSize        =   -1  'True
            Caption         =   "阳性公式"
            Height          =   180
            Left            =   240
            TabIndex        =   29
            Top             =   2130
            Width           =   720
         End
         Begin VB.Label lbl弱阳性公式 
            AutoSize        =   -1  'True
            Caption         =   "弱阳性公式"
            Height          =   180
            Left            =   3120
            TabIndex        =   28
            Top             =   2130
            Width           =   900
         End
         Begin VB.Label lbl试剂厂商 
            AutoSize        =   -1  'True
            Caption         =   "试剂厂商"
            Height          =   180
            Left            =   6390
            TabIndex        =   27
            Top             =   1410
            Width           =   720
         End
         Begin VB.Label lbl测试方法 
            AutoSize        =   -1  'True
            Caption         =   "测试方法"
            Height          =   180
            Left            =   6390
            TabIndex        =   26
            Top             =   1770
            Width           =   720
         End
         Begin VB.Label lbl测试板号 
            AutoSize        =   -1  'True
            Caption         =   "测试板号"
            Height          =   180
            Left            =   6390
            TabIndex        =   25
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lbl试剂批号 
            AutoSize        =   -1  'True
            Caption         =   "试剂批号"
            Height          =   180
            Left            =   6390
            TabIndex        =   24
            Top             =   690
            Width           =   720
         End
         Begin VB.Label lbl试剂效期 
            AutoSize        =   -1  'True
            Caption         =   "试剂效期"
            Height          =   180
            Left            =   6390
            TabIndex        =   23
            Top             =   1050
            Width           =   720
         End
         Begin VB.Label lbl测试项目 
            AutoSize        =   -1  'True
            Caption         =   "测试项目"
            Height          =   180
            Left            =   240
            TabIndex        =   22
            Top             =   1770
            Width           =   720
         End
      End
      Begin VB.Frame fra显示方式 
         Caption         =   "显示方式"
         Height          =   2505
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1305
         Begin VB.OptionButton opt显示 
            Caption         =   "定性值"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   5
            Top             =   1470
            Width           =   1125
         End
         Begin VB.OptionButton opt显示 
            Caption         =   "OD值"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   4
            Top             =   1070
            Width           =   1125
         End
         Begin VB.OptionButton opt显示 
            Caption         =   "原始OD值"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   3
            Top             =   670
            Width           =   1125
         End
         Begin VB.OptionButton opt显示 
            Caption         =   "样本编号"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   270
            Value           =   -1  'True
            Width           =   1125
         End
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   720
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   1305
      _cx             =   2302
      _cy             =   1270
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
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
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   79
      Top             =   7440
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabMB.frx":D0A4
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20585
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   1800
      Top             =   210
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmLabMB.frx":D938
      Left            =   1530
      Top             =   1020
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLabMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conPane_List = 201
Const conPane_Base = 202
Const conFontColor_BC = vbCyan
Const conFontColor_NC = vbBlue
Const conFontColor_PC = vbRed
Const conFontColor_QC = vbGreen
Const conFontColor_BK = vbBlack
Const conFontColor_YR = vbMagenta
Const conFontColor_YL = &H7B55DF

Private mlngEditWidth As Long       '为适应大字体情况下窗体变大.先读入窗体大小.
Private Enum mCol
    ID = 0: 板号: 测试时间: 试剂批号: 试剂效期: 试剂厂商: 测试方法: 波长: 参考波长: 振板频率: 振板时间: 进板方式: 空白形式: OD减空白: 单板单项: 测试项目: 阳性公式: 弱阳性公式: CutOff公式: 测试结果: 存放位置: 试剂记录
End Enum
Private mEditState As Integer                           '编辑状态: 0=浏览 1=增加 2=修改
Private mTestData(3, 1 To 8, 1 To 12) As String         ' 一维：（0=编号;1=原始OD:2=OD;3=定性) 二维三维:(微孔板坐标)
Private mTestItem(2, 1 To 8) As String                  '每一行的公式(0=阳性公式;1=弱阳性公式;2=Cutoff公式
Private mTestReagent(1 To 8) As String                  '第一行的试剂ID
Private mintEditState As Integer
Private mblnShowStop As Boolean
Private mlngKey As Long                         '当前记录的ID
Private mbln_Init As Boolean                    '仪器是否初始化成功
Private mblnModify As Boolean                   '是否正在修改
Private mblnRefresh As Boolean                  '是否刷新批号
Private mstr公式 As String
Private mlngMachine As Long                     '仪器ID
Private mrsCalc As adodb.Recordset              '记录计算公式
Private mblnMBSelect As Boolean                 '是否选择了模板

Private Sub cbo测试项目_Click()
    Dim rsTmp As New adodb.Recordset
    Dim intRow As Integer
    Dim intLoop As Integer
    On Error GoTo errH
    If Me.Visible = False Then Exit Sub
    If Me.cbo测试项目.ListIndex = -1 Then Exit Sub
    
    
    '取计算项目
    mrsCalc.filter = "诊治项目ID=" & Val(Me.cbo测试项目.ItemData(Me.cbo测试项目.ListIndex))
    If mrsCalc.EOF = False Then
        Me.txt阳性公式.Text = mrsCalc("阳性公式") & ""
        Me.txt弱阳性公式.Text = mrsCalc("弱阳性公式") & ""
        Me.txtCutOff公式.Text = mrsCalc("Cutoff公式") & ""
    End If
    
    
    
    
'    Me.txt阳性公式.Text = mTestItem(0, Me.vsList.Row)
'    Me.txt弱阳性公式.Text = mTestItem(1, Me.vsList.Row)
'    Me.txtCutOff公式.Text = mTestItem(2, Me.vsList.Row)
'
'
'    If rsTmp.EOF = True Then Exit Sub
'    If Me.vsList.Row = 0 Then Me.vsList.Select 1, 1
'    If mTestItem(0, Me.vsList.Row) = "" Or mblnModify = True Then mTestItem(0, Me.vsList.Row) = Nvl(rsTmp("阳性公式"))
'    If mTestItem(1, Me.vsList.Row) = "" Or mblnModify = True Then mTestItem(1, Me.vsList.Row) = Nvl(rsTmp("弱阳性公式"))
'    If mTestItem(2, Me.vsList.Row) = "" Or mblnModify = True Then mTestItem(2, Me.vsList.Row) = Nvl(rsTmp("CutOff公式"))
'    vsList.TextMatrix(Me.vsList.Row, 13) = Me.cbo测试项目.ItemData(Me.cbo测试项目.ListIndex)
    With Me.vsList
        If .Cols >= 13 And .Rows > 0 Then
            If Me.opt单板单项 Then
                For intLoop = 1 To Me.vsList.Rows - 1

                    .TextMatrix(intLoop, 13) = Me.cbo测试项目.ItemData(Me.cbo测试项目.ListIndex)
                    '阳性公式
                    If mTestItem(0, intLoop) = "" Or mblnModify = True Then mTestItem(0, intLoop) = Me.txt阳性公式.Text
                    If mTestItem(1, intLoop) = "" Or mblnModify = True Then mTestItem(1, intLoop) = Me.txt弱阳性公式.Text
                    If mTestItem(2, intLoop) = "" Or mblnModify = True Then mTestItem(2, intLoop) = Me.txtCutOff公式.Text
                Next
            End If
        End If
    End With
'
'    Me.txt阳性公式.Text = mTestItem(0, Me.vsList.Row)
'    Me.txt弱阳性公式.Text = mTestItem(1, Me.vsList.Row)
'    Me.txtCutOff公式.Text = mTestItem(2, Me.vsList.Row)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo测试项目_GotFocus()
    mblnModify = True
End Sub

Private Sub cbo测试项目_LostFocus()
    mblnModify = False
End Sub

Private Sub cbo检验仪器_Click()
    Dim rsTmp As New adodb.Recordset
    Dim aItem() As String
    Dim intLoop As Integer
    
    On Error GoTo errH
    
    If Me.cbo检验仪器.ListCount = 0 Then Exit Sub
    
    If mbln_Init Then frmLabMBControl.MB_Stop  '停止已初始的仪器控制
    
    If cbo检验仪器.ListIndex >= 0 Then
        mlngMachine = cbo检验仪器.ItemData(cbo检验仪器.ListIndex)
    End If
    
    gstrSql = "select 波长,振板频率,振板时间,进板方式,空白形式 from 检验仪器 where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.cbo检验仪器.ItemData(Me.cbo检验仪器.ListIndex))
    If rsTmp.EOF = True Then Exit Sub
    
'    If mEditState = 0 Then Exit Sub
    
            
    With Me.cbo波长
        .Clear
        Me.cbo参考波长.Clear
        Me.cbo参考波长.AddItem ""
        Me.cbo参考波长.ItemData(Me.cbo参考波长.NewIndex) = 0
        aItem = Split(Nvl(rsTmp("波长")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
            Me.cbo参考波长.AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    With Me.cbo振板频率
        .Clear
        aItem = Split(Nvl(rsTmp("振板频率")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    Me.txt振板时间.Text = Nvl(rsTmp("振板时间"))
    
    With Me.cbo进板方式
        .Clear
        aItem = Split(Nvl(rsTmp("进板方式")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    With Me.cbo空白形式
        .Clear
        aItem = Split(Nvl(rsTmp("空白形式")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    Call RefreshList
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo选择模板_Click()
    Dim rsTmp As New adodb.Recordset
    Dim intLoop As Integer
    Dim intRow As Integer, intCol As Integer
    Dim aResult() As String
    Dim aItem() As String
    Dim blnOne As Boolean
    Dim lngItemID As Long
    
    If Me.cbo选择模板.ItemData(Me.cbo选择模板.ListIndex) <= 0 Then Exit Sub
    
    gstrSql = "select id,编号,名称,项目,内容 from 检验酶标模板 where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.cbo选择模板.ItemData(Me.cbo选择模板.ListIndex))
    
    If rsTmp.EOF = True Then
        MsgBox "没有找到模板记录!", vbInformation
        Exit Sub
    End If
            
   aItem = Split(rsTmp("项目"), ";")
   aResult = Split(rsTmp("内容"), "|")
   
   intLoop = Val(Me.txt开始标本号)
   
   If Me.opt方向(0).Value = True Then
        '横向
        For intRow = 1 To 8
            If Me.opt单板多项.Value = True Then intLoop = Val(Me.txt开始标本号)
            For intCol = 1 To 12
                With Me.vsList
                    If intLoop = 0 Then
                        .TextMatrix(intRow, intCol) = Split(aResult(intRow - 1), ";")(intCol - 1)
                    Else
                        If Trim(.TextMatrix(intRow, intCol)) <> "" Then
                            If IsNumeric(Split(aResult(intRow - 1), ";")(intCol - 1)) = True Then
                                .TextMatrix(intRow, intCol) = intLoop
                                intLoop = intLoop + 1
                            Else
                                .TextMatrix(intRow, intCol) = Split(aResult(intRow - 1), ";")(intCol - 1)
                            End If
                        End If
                    End If
                    mTestData(0, intRow, intCol) = .TextMatrix(intRow, intCol)
                End With
            Next
        Next
    Else
        '纵向
        For intCol = 1 To 12
            For intRow = 1 To 8
                With Me.vsList
                    If intLoop = 0 Then
                        .TextMatrix(intRow, intCol) = Split(aResult(intRow - 1), ";")(intCol - 1)
                    Else
                        If Trim(.TextMatrix(intRow, intCol)) <> "" Then
                            If IsNumeric(Split(aResult(intRow - 1), ";")(intCol - 1)) = True Then
                                .TextMatrix(intRow, intCol) = intLoop
                                intLoop = intLoop + 1
                            Else
                                .TextMatrix(intRow, intCol) = Split(aResult(intRow - 1), ";")(intCol - 1)
                            End If
                        End If
                    End If
                    mTestData(0, intRow, intCol) = .TextMatrix(intRow, intCol)
                End With
            Next
        Next
    End If
    
    
    For intRow = 1 To 8
        With Me.vsList
            .TextMatrix(intRow, 13) = aItem(intRow - 1)
            
            
            If lngItemID = 0 And .TextMatrix(intRow, 13) <> "" Then
                lngItemID = .TextMatrix(intRow, 13)
            End If
            If intRow > 1 Then
                If .TextMatrix(intRow, 13) <> aItem(intRow - 2) Then
                    blnOne = True
                End If
            End If
            
        End With
    Next
    If blnOne = True Then
        Me.opt单板多项.Value = True
    Else
        Me.opt单板单项.Value = True
    End If
    
    Me.txt开始标本号.Text = ""
    
    For intLoop = 0 To 3
        If Me.opt显示(intLoop).Value = True Then
            Call opt显示_Click(intLoop)
        End If
    Next
    
    Erase mTestItem
    For intRow = 1 To 8
        For intLoop = 0 To Me.cbo测试项目.ListCount - 1
            If Val(aItem(intRow - 1)) = Me.cbo测试项目.ItemData(intLoop) Then
                Me.vsList.Row = intRow
                Me.cbo测试项目.ListIndex = intLoop
                If mTestItem(0, intRow) = "" Then mTestItem(0, intRow) = Me.txt阳性公式
                If mTestItem(1, intRow) = "" Then mTestItem(1, intRow) = Me.txt弱阳性公式
                If mTestItem(2, intRow) = "" Then mTestItem(2, intRow) = Me.txtCutOff公式
                Exit For
            End If
        Next
    Next
    
    mblnMBSelect = True
'    With Me.vsList
'        For intRow = 1 To 8
'            If Val(.TextMatrix(intRow, 13)) <> 0 Then
'                With cbo测试项目
'                    For intLoop = 0 To .ListCount - 1
'                        If .ItemData(intLoop) = Val(vsList.TextMatrix(intRow, 13)) Then
'                            .ListIndex = intLoop
'                        End If
'                    Next
'                End With
'            End If
'        Next
'    End With
'
'    If lngItemID <> 0 Then
'        With cbo测试项目
'            For intRow = 0 To .ListCount - 1
'                If .ItemData(intRow) = lngItemID Then
'                    .ListIndex = intRow
'                End If
'            Next
'        End With
'    End If
'    Me.cbo选择模板.ListIndex = 0
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl                 '文本标签
    Dim strFilter As String                             '过滤字串
    Dim rsTmp As New adodb.Recordset
    
    Select Case Control.ID
        Case conMenu_File_PrintSet                                                          '打印设置
            zlPrintSet
        Case conMenu_File_Preview                                                           '预览
            Call zlRptPrint(0)
        Case conMenu_File_Print                                                             '打印
            Call zlRptPrint(1)
        Case conMenu_File_Excel                                                             '输出到Excel
            Call zlRptPrint(2)
        Case conMenu_File_Parameter                                                         '参数设置
            frmLabMBSetup.Show vbModal, Me
        Case conMenu_Edit_Save                                                              '保存
            Call SaveData
        Case conMenu_Edit_Untread                                                           '取消
            Call InitItem: mEditState = 0: RefreshItem (mlngKey)
        Case conMenu_File_Exit                                                              '退出
            Unload Me
        '----------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem                                                           '新增
            Call AddNew
        Case conMenu_Edit_Modify                                                            '修改
            mEditState = 2
        Case conMenu_Edit_Delete                                                            '删除
            Call DelData
        Case conMenu_Edit_Leave_Post                                                        '计算
            Call CalcData
        Case conMenu_Edit_Send                                                              '测量
            Call MBcontrol
        Case conMenu_LIS_MB_Connect                                                               '选定仪器(连接仪器)
            If Me.cbo检验仪器.ListIndex >= 0 Then
                mbln_Init = frmLabMBControl.MB_Start(Me, Me.cbo检验仪器.ItemData(Me.cbo检验仪器.ListIndex))
            Else
                MsgBox "请选择一个酶标仪", vbInformation, Me.Caption
            End If
        Case conMenu_LIS_MB_Disconnect                                                               '取消选定(断开仪器连接)
            Call frmLabMBControl.MB_Stop
            mbln_Init = False
        Case conMenu_Edit_QCRes                                                             '试剂管理
            frmLabMBReagent.Show vbModal, Me

        Case conMenu_Edit_Adjust                                                            '批量调整OD
            mstr公式 = frmLabMBcalc.ShowMe(Me)
            Call CalcData
            
        '-----------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Button                                                    '标准按钮
            Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text                                                      '文本标签
            For Each cbrControl In Me.cbsThis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size                                                      '大图标
            Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
            Me.cbsThis.RecalcLayout
        Case conMenu_View_StatusBar                                                         '状态栏
            
        Case conMenu_View_Find                                                              '查找
            strFilter = frmLabMBFilter.ShowMe(Me)
            If strFilter <> "" Then Call RefreshList(2, strFilter)
        Case conMenu_View_Refresh                                                           '刷新
            Call RefreshList
        '-----------------------------------------------------------------------------------------------------
        Case conMenu_Help_Help                                                              '帮助主题
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web                                                               'WEB
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Home                                                          '主页
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail                                                          '发送返馈
            Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About                                                             '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height

End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_PrintSet                                                          '打印设置
        Case conMenu_File_Preview                                                           '预览
        Case conMenu_File_Print                                                             '打印
        Case conMenu_File_Excel                                                             '输出到Excel
        Case conMenu_Edit_Save                                                              '保存
            Control.Enabled = (mEditState > 0)
        Case conMenu_Edit_Untread                                                           '取消
            Control.Enabled = (mEditState > 0)
        '----------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem                                                           '新增
            Control.Enabled = (mEditState = 0)
        Case conMenu_Edit_Modify                                                            '修改
            Control.Enabled = (mEditState = 0 And Me.rptList.Records.Count > 0)
        Case conMenu_Edit_Delete                                                            '删除
            Control.Enabled = (mEditState = 0 And Me.rptList.Records.Count > 0)
        Case conMenu_Edit_Leave_Post                                                        '计算
            Control.Enabled = (mEditState > 0)
        Case conMenu_Edit_Send                                                              '发送
            Control.Enabled = (mEditState > 0) And mbln_Init
        Case conMenu_LIS_MB_Connect                                                               '连接
            
            Control.Enabled = Not mbln_Init
        
        Case conMenu_LIS_MB_Disconnect                                                               '断开
            Control.Enabled = mbln_Init
        
        Case conMenu_Edit_Adjust                                                            '批量调整
            Control.Enabled = (mEditState > 0)
        '-----------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Button                                                    '标准按钮
            Control.Checked = Me.cbsThis(2).Visible
        Case conMenu_View_ToolBar_Text                                                      '文本标签
            Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size                                                      '大图标
            Control.Checked = Me.cbsThis.Options.LargeIcons
        Case conMenu_View_StatusBar                                                         '状态栏
        Case conMenu_View_Find                                                              '查找
            
        Case conMenu_View_Refresh                                                           '刷新
        '-----------------------------------------------------------------------------------------------------
        Case conMenu_Help_Help                                                              '帮助主题
        Case conMenu_Help_Web                                                               'WEB
        Case conMenu_Help_Web_Home                                                          '主页
        Case conMenu_Help_Web_Mail                                                          '发送返馈
        Case conMenu_Help_About                                                             '关于
    End Select
    If mEditState = 0 Then
        Me.fra测量参数.Enabled = False
        Me.fra模板.Enabled = False
'        Me.fra微孔板.Enabled = False
        Me.vsList.Editable = flexEDNone
        PicList.Enabled = True
    Else
        Me.fra测量参数.Enabled = True
        Me.fra模板.Enabled = True
        'Me.fra微孔板.Enabled = True
        Me.vsList.Editable = flexEDKbdMouse
        PicList.Enabled = False
    End If
End Sub

Private Sub cmdSl_Click()
    Me.txt试剂批号 = ""
    Call SelectBatch
End Sub

Private Sub cmd保存模板_Click()
    Call SaveTemplet
End Sub

Private Sub cmd确定_Click()
    Call subWriteNumber
End Sub

Private Sub cmd删除模板_Click()
    If Me.cbo选择模板.Text = "" Then Exit Sub
    If MsgBox("是否确定要删除<" & Me.cbo选择模板 & ">模板?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        '删除
        On Error GoTo errH
        gstrSql = "Zl_检验酶标模板_Delete(" & Me.cbo选择模板.ItemData(Me.cbo选择模板.ListIndex) & ")"
        zlDatabase.ExecuteProcedure gstrSql, Me.Caption
        RefreshTemplet
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd重置_Click()
    Dim intY As Integer, intX As Integer
    Erase mTestData
    Erase mTestItem
    Erase mTestReagent
    With Me.vsList
        For intY = 1 To .Rows - 1
            For intX = 1 To .Cols - 1
                .TextMatrix(intY, intX) = ""
            Next
        Next
    End With
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = Me.PicList.hWnd
    Case conPane_Base
        Item.Handle = Me.PicMain.hWnd
    End Select
End Sub

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
    Me.cbsThis.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    Top = lngTop
    Bottom = Me.ScaleHeight - lngBottom
End Sub

Private Sub dkpMan_Resize()
    Me.cbsThis.RecalcLayout
End Sub

Private Sub Form_Load()
    Dim intLoop As Integer
    Dim intX As Integer, intY As Integer
    Dim rsTmp As New adodb.Recordset
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim strName As String
    Dim lngMachine As Long
    Dim rptCol As ReportColumn
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
'    mstrPrivs = gstrPrivs
    
    mlngEditWidth = Me.PicMain.Width
    
    mintEditState = 0: mblnShowStop = False
    Me.cbsThis.EnableCustomization False
    
'    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
   '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
'    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&O)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Leave_Post, "记算(&C)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_MB_Connect, "连接仪器(&N)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "测量(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_MB_Disconnect, "断开仪器(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_QCRes, "试剂管理(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "批量OD调整(&O)"): cbrControl.BeginGroup = True
        
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
'        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Find, "查找(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "检验仪器")
    cbrControl.Flags = xtpFlagRightAlign

    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_Report_DrugQuery, "检验仪器")
    cbrCustom.ShortcutText = "检验仪器"
    cbrCustom.Handle = Me.cbo检验仪器.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    cbrCustom.Style = xtpButtonIconAndCaption
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add FCONTROL, Asc("B"), conMenu_Edit_Compend
        .Add FCONTROL, Asc("E"), conMenu_Edit_ApplyTo
        .Add FCONTROL, Asc("G"), conMenu_Edit_Test
        .Add FCONTROL, Asc("F"), conMenu_View_Find
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_Edit_Pause
        .AddHiddenCommand conMenu_Edit_Reuse
        .AddHiddenCommand conMenu_View_Refresh
        .AddHiddenCommand conMenu_View_Option
    End With
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Leave_Post, "记算"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_MB_Connect, "连接")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "测量")
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_MB_Disconnect, "断开")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Find, "查找")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '设置词句显示停靠窗格
    Dim panSub1 As Pane, panSub2 As Pane, panSub3 As Pane

    
    Set panSub1 = dkpMan.CreatePane(conPane_List, 300, 580, DockLeftOf, Nothing)
    panSub1.Title = "测试板列表"
    panSub1.Options = PaneNoCaption

    Set panSub2 = dkpMan.CreatePane(conPane_Base, 550, 200, DockRightOf, Nothing)
    panSub2.Title = "控制界面"
    panSub2.Options = PaneNoCaption

    panSub1.Select
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.板号, "板号", 80, True): .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.测试时间, "测试时间", 85, True)
        Set rptCol = .Columns.Add(mCol.试剂批号, "试剂批号", 85, True)
        Set rptCol = .Columns.Add(mCol.试剂效期, "试剂效期", 85, True)
        Set rptCol = .Columns.Add(mCol.试剂厂商, "试剂厂商", 85, True)
        Set rptCol = .Columns.Add(mCol.测试方法, "测试方法", 85, True)
        Set rptCol = .Columns.Add(mCol.波长, "波长", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.参考波长, "参考波长", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.振板频率, "振板频率", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.振板时间, "振板时间", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.进板方式, "进板方式", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.空白形式, "空白形式", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.OD减空白, "OD减空白", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.单板单项, "单板单项", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.测试项目, "测试项目", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.阳性公式, "阳生公式", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.弱阳性公式, "弱阳性公式", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.CutOff公式, "CutOff公式", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.测试结果, "测试结果", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.存放位置, "存放位置", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.试剂记录, "试剂记录", 85, False): rptCol.Visible = False
        
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    
    '-----------------------------------------------------
    '界面恢复
'    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    With Me.vsList
        .Rows = 9
        .Cols = 14
        .FixedRows = 1
        .FixedCols = 1 '
        For intLoop = 1 To .Cols - 2
            .TextMatrix(0, intLoop) = intLoop
        Next
        .TextMatrix(0, 13) = "项目"
        
        For intLoop = 1 To .Rows - 1
            .TextMatrix(intLoop, 0) = Chr(intLoop + 64)
        Next
       
       .Select 0, 0, 8, 13

      .FillStyle = flexFillRepeat

      .CellAlignment = flexAlignCenterCenter

      'return .FillStyle to its default (if needed)

      .FillStyle = flexFillSingle
      .Select 0, 0, 0, 0
     
      .Cell(flexcpBackColor, 1, 13, 8, 13) = RGB(200, 200, 200)
    End With
    
    Call InitRecordSet(mrsCalc)
    
'    Me.chk阴性对照.Value = Mid(GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "阴性对照", "0,"), 1, 1)
'    Me.txt最小阴性对照.Text = Mid(GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "阴性对照", "0,"), 3)
    If zlDatabase.GetPara("frmLabMB_阴性对照", 100, 1208, "") = "" Then
        Me.chk阴性对照.Value = 0
        Me.txt最小阴性对照.Text = ""
    Else
        Me.chk阴性对照.Value = Mid(zlDatabase.GetPara("frmLabMB_阴性对照", 100, 1208, "0,"), 1, 1)
        Me.txt最小阴性对照.Text = Mid(zlDatabase.GetPara("frmLabMB_阴性对照", 100, 1208, "0,"), 3)
    End If
    '读入仪器
'    lngMachine = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "仪器ID", 0)
'    lngMachine = zlDatabase.GetPara("frmLabMB_仪器ID", 100, 1208, 0)
    
    gstrSql = "select id,编码,名称 from 检验仪器 where  微生物 = 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With Me.cbo检验仪器
        .Clear
        Do Until rsTmp.EOF
            .AddItem rsTmp("编码") & "-" & rsTmp("名称")
            .ItemData(.NewIndex) = rsTmp("ID")
            If rsTmp("ID") = mlngMachine Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
        If .ListCount > 0 Then
            If .ListIndex < 0 Then .ListIndex = 0
        End If
    End With
    
    '读入检验项目
    If Me.cbo检验仪器.ListCount > 0 Then
        gstrSql = "select id,中文名,英文名 from  诊治所见项目 a , 检验项目 b,检验仪器项目 c  where a.id = b.诊治项目id and 项目类别 = 4 " & _
                   " And c.仪器id = [1] And a.id = c.项目id "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.cbo检验仪器.ItemData(Me.cbo检验仪器.ListIndex)))
        With Me.cbo测试项目
            .Clear
            Do Until rsTmp.EOF
                .AddItem Nvl(rsTmp("中文名")) & "(" & Nvl(rsTmp("英文名")) & ")"
                .ItemData(.NewIndex) = rsTmp("id")
                strName = strName & "|#" & Nvl(rsTmp("ID")) & ";" & Nvl(rsTmp("中文名"))
                rsTmp.MoveNext
            Loop
            With Me.vsList
                .ColComboList(13) = strName
            End With
        End With
        RefreshList
    End If
    mbln_Init = False
    Call RefreshTemplet
    
    Call cbo检验仪器_Click
    
    Call RestoreWinState(Me, App.ProductName)                   '界面恢复
End Sub

Private Sub Form_Resize()
    Dim panBase As Pane
    Dim intLoop As Integer
    
    If Me.WindowState = vbMinimized Then Exit Sub
    Set panBase = Me.dkpMan.FindPane(conPane_Base)
    panBase.MinTrackSize.SetSize mlngEditWidth / Screen.TwipsPerPixelX, 265
'    panBase.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 265
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters

'    panBase.MinTrackSize.SetSize 0, 0
'    panBase.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 265

    
    
    
End Sub

Private Sub fraMain_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mEditState = 0
    mlngKey = 0
    Erase mTestData
    Erase mTestItem
    Erase mTestReagent
'    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "仪器ID", cbo检验仪器.ItemData(cbo检验仪器.NewIndex))
'    SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "阴性对照", Me.chk阴性对照.Value & "," & Me.txt最小阴性对照
    
    zlDatabase.SetPara "frmLabMB_仪器ID", cbo检验仪器.ItemData(cbo检验仪器.NewIndex), 100, 1208
'    zlDatabase.SetPara "frmLabMB_阴性对照", Me.chk阴性对照.Value & "," & Me.txt最小阴性对照, 100, 1208
    mblnMBSelect = False
End Sub


Private Sub opt过滤_Click(Index As Integer)
    Call RefreshList(1)
End Sub

Private Sub opt显示_Click(Index As Integer)
    Dim intRow As Integer, intCol As Integer
    For intRow = 1 To 8
        For intCol = 1 To 12
            With Me.vsList
                .TextMatrix(intRow, intCol) = mTestData(Index, intRow, intCol)
                                
                If InStr(mTestData(0, intRow, intCol), "BC") > 0 Then
                    '空白
                    .Cell(flexcpFontBold, intRow, intCol) = True
                    .Cell(flexcpForeColor, intRow, intCol) = conFontColor_BC
                ElseIf InStr(mTestData(0, intRow, intCol), "NC") > 0 Then
                    '阴性
                    .Cell(flexcpFontBold, intRow, intCol) = True
                    .Cell(flexcpForeColor, intRow, intCol) = conFontColor_NC
                ElseIf InStr(mTestData(0, intRow, intCol), "PC") > 0 Then
                    '阳性
                    .Cell(flexcpFontBold, intRow, intCol) = True
                    .Cell(flexcpForeColor, intRow, intCol) = conFontColor_PC
                ElseIf InStr(mTestData(0, intRow, intCol), "QC") > 0 Then
                    '质控
                    .Cell(flexcpFontBold, intRow, intCol) = True
                    .Cell(flexcpForeColor, intRow, intCol) = conFontColor_QC
                Else
                    If InStr(mTestData(3, intRow, intCol), "+") > 0 Then
                        .Cell(flexcpFontBold, intRow, intCol) = True
                        .Cell(flexcpForeColor, intRow, intCol) = conFontColor_YR
                    ElseIf InStr(mTestData(3, intRow, intCol), "±") > 0 Then
                        .Cell(flexcpFontBold, intRow, intCol) = True
                        .Cell(flexcpForeColor, intRow, intCol) = conFontColor_YL
                    Else
                        .Cell(flexcpFontBold, intRow, intCol) = False
                        .Cell(flexcpForeColor, intRow, intCol) = conFontColor_BK
                    End If
                End If
                
            End With
        Next
    Next
                
End Sub

Private Sub picList_Resize()
    With Me.rptList
        .Left = 10
        .Width = Me.PicList.ScaleWidth - 20
        .Height = Me.PicList.ScaleHeight - .Top - 20
        .Top = Me.opt过滤(0).Top + Me.opt过滤(1).Height + 10
    End With
End Sub

Private Sub picMain_Resize()
    Dim intLoop As Integer
    
    On Error Resume Next

    With Me.fra测量参数
        .Width = Me.PicMain.ScaleWidth - .Left - 50
    End With
    
    With Me.fra模板
        .Width = Me.PicMain.ScaleWidth - .Left - 50
    End With
    
    With Me.fra微孔板
        .Width = Me.PicMain.ScaleWidth - .Left - 50
        .Height = Me.PicMain.ScaleHeight - .Top - 50
    End With
    
    With Me.fra对照
        .Top = Me.fra微孔板.Height - .Height - 50
    End With
    
    With Me.vsList
        .Width = Me.fra微孔板.Width - .Left - 50
        .Height = Me.fra对照.Top - .Top
    End With
    
    With Me.vsList
        
        For intLoop = 0 To .Rows - 1
            .RowHeight(intLoop) = (.Height - 9 * 15 - 300) / 8
        Next
        
        For intLoop = 0 To .Cols - 1
            .ColWidth(intLoop) = (.Width - 14 * 15 - 300 - 2000) / 12
        Next
        
        .ColWidth(0) = 300
        .RowHeight(0) = 300
        .ColWidth(13) = 2000
    End With
    
    With Me.txt测试板号
        .Width = Me.fra测量参数.Width - .Left - 100
    End With
    
    With Me.txt试剂批号
        .Width = Me.txt测试板号.Width - Me.cmdSl.Width
        Me.cmdSl.Left = .Left + .Width
        Me.cmdSl.Top = .Top
    End With
    
    With Me.txt试剂效期
        .Width = Me.txt测试板号.Width
    End With
    
    With Me.txt试剂厂商
        .Width = Me.txt测试板号.Width
    End With
    
    With Me.txt测试方法
        .Width = Me.txt测试板号.Width
    End With
    
    With Me.txtCutOff公式
        .Width = Me.txt测试板号.Width
    End With
End Sub

Private Sub rptList_SelectionChanged()
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    mlngKey = Me.rptList.FocusedRow.Record(mCol.ID).Value
    Erase mTestData
    Erase mTestItem
    RefreshItem mlngKey
    
End Sub

Private Sub txtCutOff公式_KeyPress(KeyAscii As Integer)
    Dim intRow  As Integer
    On Error Resume Next
    If KeyAscii = 13 Then
        mrsCalc.filter = "诊治项目ID=" & Me.cbo测试项目.ItemData(Me.cbo测试项目.ListIndex)
        If mrsCalc.EOF = False Then
            mrsCalc("CutOff公式") = Me.txtCutOff公式
            mrsCalc.Update
        End If
        mTestItem(2, Me.vsList.Row) = Me.txtCutOff公式
        With Me.vsList
            For intRow = 1 To 8
                If .TextMatrix(intRow, 13) = Me.cbo测试项目.ItemData(Me.cbo测试项目.ListIndex) Then
                    mTestItem(2, intRow) = Me.txtCutOff公式
                End If
            Next
        End With
        MsgBox "修改公式成功!", vbInformation, Me.Caption
    End If
End Sub

Private Sub txt开始标本号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call subWriteNumber
    End If
End Sub

Private Sub txt弱阳性公式_KeyPress(KeyAscii As Integer)
    Dim intRow  As Integer
    On Error Resume Next
    If KeyAscii = 13 Then
        mTestItem(1, Me.vsList.Row) = Me.txt弱阳性公式
        mrsCalc.filter = "诊治项目ID=" & Me.cbo测试项目.ItemData(Me.cbo测试项目.ListIndex)
        If mrsCalc.EOF = False Then
            mrsCalc("弱阳性公式") = Me.txt弱阳性公式
            mrsCalc.Update
        End If
        With Me.vsList
            For intRow = 1 To 8
                If .TextMatrix(intRow, 13) = Me.cbo测试项目.ItemData(Me.cbo测试项目.ListIndex) Then
                    mTestItem(1, intRow) = Me.txt弱阳性公式
                End If
            Next
        End With
        MsgBox "修改公式成功!", vbInformation, Me.Caption
    End If
End Sub

Private Sub txt试剂批号_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call SelectBatch
    End If
End Sub

Private Sub txt试剂批号_Validate(Cancel As Boolean)
    If txt试剂批号.Text <> "" Then
        Call SelectBatch
    Else
        txt试剂批号.Tag = ""
        txt试剂效期 = ""
        txt试剂厂商 = ""
        txt测试方法 = ""
    End If
End Sub

Private Sub txt阳性公式_KeyPress(KeyAscii As Integer)
    Dim intRow  As Integer
    On Error Resume Next
    If KeyAscii = 13 Then
        mrsCalc.filter = "诊治项目ID=" & Me.cbo测试项目.ItemData(Me.cbo测试项目.ListIndex)
        If mrsCalc.EOF = False Then
            mrsCalc("阳性公式") = Me.txt阳性公式
            mrsCalc.Update
        End If
        mTestItem(0, Me.vsList.Row) = Me.txt阳性公式
        With Me.vsList
            For intRow = 1 To 8
                If .TextMatrix(intRow, 13) = Me.cbo测试项目.ItemData(Me.cbo测试项目.ListIndex) Then
                    mTestItem(0, intRow) = Me.txt阳性公式
                End If
            Next
        End With
        MsgBox "修改公式成功!", vbInformation, Me.Caption
    End If
End Sub

Private Sub vsList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim intLoop As Integer
    Dim intRow As Integer
    If mEditState = 0 Then Exit Sub
    If Col = 13 And Row > 0 Then
        Erase mTestItem
        With Me.vsList
            For intLoop = 1 To Me.vsList.Rows - 1
                If Me.opt单板单项 = True Then
                    .TextMatrix(intLoop, 13) = .TextMatrix(Row, Col)
                End If
                Me.vsList.Row = intLoop
                For intRow = 0 To Me.cbo测试项目.ListCount - 1
                    If Val(.TextMatrix(intLoop, 13)) = Me.cbo测试项目.ItemData(intRow) Then
                        Me.cbo测试项目.ListIndex = intRow
                        Call cbo测试项目_Click
                        mTestItem(0, intLoop) = Me.txt阳性公式
                        mTestItem(1, intLoop) = Me.txt弱阳性公式
                        mTestItem(2, intLoop) = Me.txtCutOff公式
                    End If
                Next
            Next
        End With
    End If
    
    For intLoop = 0 To Me.opt显示.UBound
        If opt显示(intLoop).Value = True Then
            Exit For
        End If
    Next
    If Row > 0 And Col > 0 And Col < 13 Then
        mTestData(intLoop, Row, Col) = Me.vsList.TextMatrix(Row, Col)
    End If
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call CalcData
End Sub

Private Sub vsList_Click()
    Dim lngKey As Long
    Dim intLoop As Integer
    With Me.vsList
        If .Row > 0 Then
            lngKey = Val(.TextMatrix(.Row, 13))
        End If
    End With
    If lngKey <> 0 Then
        With Me.cbo测试项目
            For intLoop = 0 To .ListCount - 1
                If .ItemData(intLoop) = lngKey Then
                    .ListIndex = intLoop
                    Me.txt阳性公式.Text = mTestItem(0, Me.vsList.Row)
                    Me.txt弱阳性公式.Text = mTestItem(1, Me.vsList.Row)
                    Me.txtCutOff公式.Text = mTestItem(2, Me.vsList.Row)
                    CalcData .ItemData(intLoop)
                End If
            Next
        End With
    End If
    On Error Resume Next
    If Me.vsList.Row > 0 And Me.vsList.Row < 9 Then
        If mTestReagent(Me.vsList.Row) <> "" Then
            mblnRefresh = True
            txt试剂批号.Text = Split(mTestReagent(Me.vsList.Row), ";")(0)
            txt试剂效期.Text = Split(mTestReagent(Me.vsList.Row), ";")(1)
            txt试剂厂商.Text = Split(mTestReagent(Me.vsList.Row), ";")(2)
            txt测试方法.Text = Split(mTestReagent(Me.vsList.Row), ";")(3)
            mblnRefresh = False
        Else
            txt试剂批号 = ""
        End If
    End If
    mblnMBSelect = False
End Sub

Private Sub vsList_DblClick()
    Dim strMaxNumber As String
    If mEditState = 0 Then Exit Sub
    With Me.vsList
        If .Row > 0 And .Col > 0 And .Col < 13 And Me.opt显示(0).Value = True Then
            strMaxNumber = GetMaxNumber
            .TextMatrix(.Row, .Col) = strMaxNumber
            If InStr(strMaxNumber, "BC") > 0 Then
                '空白
                .Cell(flexcpFontBold, .Row, .Col) = True
                .Cell(flexcpForeColor, .Row, .Col) = conFontColor_BC
            ElseIf InStr(strMaxNumber, "NC") > 0 Then
                '阴性
                .Cell(flexcpFontBold, .Row, .Col) = True
                .Cell(flexcpForeColor, .Row, .Col) = conFontColor_NC
            ElseIf InStr(strMaxNumber, "PC") > 0 Then
                '阳性
                .Cell(flexcpFontBold, .Row, .Col) = True
                .Cell(flexcpForeColor, .Row, .Col) = conFontColor_PC
            ElseIf InStr(strMaxNumber, "QC") > 0 Then
                '质控
                .Cell(flexcpFontBold, .Row, .Col) = True
                .Cell(flexcpForeColor, .Row, .Col) = conFontColor_QC
            End If
            
        End If
    End With
End Sub

Private Sub vsList_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim intX As Integer, intY As Integer
    If KeyCode = vbKeyDelete Then
        If MsgBox("是否确定要删除选中的编号?", vbYesNo + vbDefaultButton2 + vbQuestion, Me.Caption) = vbNo Then Exit Sub
        With vsList
            For intY = .Row To .RowSel
                For intX = .Col To .ColSel
                    .TextMatrix(intY, intX) = ""
                    mTestData(0, intY, intX) = ""
                    mTestData(1, intY, intX) = ""
                    mTestData(2, intY, intX) = ""
                    mTestData(3, intY, intX) = ""
                Next
            Next
        End With
    End If
End Sub

Private Sub vsList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mEditState = 0 Then Exit Sub
    Call subWriteNumber
End Sub

Private Sub subWriteNumber()
    '功能           按一规则写入标本编号
    Dim intX As Integer, intY As Integer
    Dim intLoop As Integer
    
    If Val(Me.txt开始标本号.Text) = 0 Then Exit Sub
    If Me.opt显示(0).Value = False Then Exit Sub
    
    '有模板选择时直接使用模板来生成
    If mblnMBSelect = True Then
        If Me.cbo选择模板.Text <> "" Then cbo选择模板_Click: Exit Sub
    End If
    
    intLoop = Val(Me.txt开始标本号.Text) - 1
    With Me.vsList
        If Me.opt方向(0).Value = True Then
            For intY = .Row To .RowSel
                If Me.opt单板多项.Value = True Then intLoop = Val(Me.txt开始标本号.Text) - 1
                For intX = .Col To .ColSel
                    If intY > 0 And intX > 0 And intX < 13 Then
                        intLoop = intLoop + 1
                        .TextMatrix(intY, intX) = intLoop
                        mTestData(0, intY, intX) = intLoop
                        .Cell(flexcpData, intY, intX) = intLoop
                        .Cell(flexcpFontBold, intY, intX) = False
                        .Cell(flexcpForeColor, intY, intX) = vbBlack
                    End If
                Next
            Next
        Else
            For intX = .Col To .ColSel
                If Me.opt单板多项.Value = True Then intLoop = Val(Me.txt开始标本号.Text) - 1
                For intY = .Row To .RowSel
                    If intY > 0 And intX > 0 And intX < 13 Then
                        intLoop = intLoop + 1
                        .TextMatrix(intY, intX) = intLoop
                        mTestData(0, intY, intX) = intLoop
                        .Cell(flexcpData, intY, intX) = intLoop
                        .Cell(flexcpFontBold, intY, intX) = False
                        .Cell(flexcpForeColor, intY, intX) = vbBlack
                    End If
                Next
            Next
        End If
    End With
    Me.txt开始标本号 = ""
End Sub
Private Function GetMaxNumber() As String
    '功能取得列表中的最大编号
    Dim intLoop As Integer
    Dim intY  As Integer, intX As Integer
    Dim intMax As Integer
    Dim strTmp As String
    Dim strType As String
    
    '得到孔类型
    For intLoop = 0 To Me.opt孔选择.UBound
        If Me.opt孔选择(intLoop).Value = True Then
            Exit For
        End If
    Next
    
    Select Case intLoop
        Case 0  '普通
            strType = "S"
        Case 1  '空白
            strType = "BC"
        Case 2  '阴性
            strType = "NC"
        Case 3  '阳性
            strType = "PC"
        Case 4  '质控
            strType = "QC"
    End Select
    
    With Me.vsList
        For intY = 1 To .Rows - 1
            For intX = 1 To .Cols - 2
                If strType = "S" Then
                    strTmp = Val(.TextMatrix(intY, intX))
                    If Val(strTmp) <> 0 Then
                        If CInt(strTmp) >= intMax Then intMax = CInt(strTmp) + 1
                    End If
                Else
                    If InStr(.TextMatrix(intY, intX), strType) > 0 Then
                        strTmp = Val(Trim(Replace(.TextMatrix(intY, intX), strType, "")))
                        If CInt(strTmp) >= intMax Then intMax = CInt(strTmp) + 1
                    End If
                End If
            Next
        Next
    End With
    
    GetMaxNumber = Replace(strType, "S", "") & IIf(intMax = 0, 1, intMax)
End Function
Private Sub SaveTemplet()
    ''''''''''''''''''''''''''''''''
    '功能   保存到模板
    ''''''''''''''''''''''''''''''''
    Dim intY As Integer, intX As Integer
    Dim intRow As Integer, intCol As Integer
    Dim strNumber As String
    Dim intLoop As Integer
    Dim strResult As String
    
    '标本号编号检查
    For intRow = 1 To 8
        For intCol = 1 To 12
            With Me.vsList
                strNumber = .TextMatrix(intRow, intCol)
                If Trim(strNumber) <> "" Then
                    If IsNumeric(strNumber) = True Then
                        If Len(strNumber) > 4 Then
                            MsgBox "标本编号最大只能为<9999>，请修改！", vbInformation
                            .Select intRow, intCol
                            Exit Sub
                        End If
                    Else
                        If InStr(strNumber, "BC") = 0 And InStr(strNumber, "NC") = 0 And InStr(strNumber, "PC") = 0 And InStr(strNumber, "QC") = 0 Then
                            MsgBox "你输入了不正确的编号<" & strNumber & ">请修正!", vbInformation
                            .Select intRow, intCol
                            Exit Sub
                        End If
                    End If
                    For intY = intRow To 8
                        For intX = intCol + 1 To 12
                            If strNumber = .TextMatrix(intY, intX) Then
                                MsgBox "发现相同的编号，请修改！", vbInformation
                                .Select intY, intX
                                Exit Sub
                            End If
                        Next
                    Next
                    intLoop = intLoop + 1
                End If
            End With
        Next
    Next
    If intLoop = 0 Then
        MsgBox "没有选择编号不能保存!", vbInformation
        Exit Sub
    End If
    
    '组织保存数据
    
    '项目
    For intRow = 1 To 8
        With Me.vsList
            strNumber = .TextMatrix(intRow, 13)
            If Trim(strNumber) <> "" Then
                intLoop = intLoop + 1
            End If
            If intRow = 1 Then
                strResult = strResult & strNumber
            Else
                strResult = strResult & ";" & strNumber
            End If
            
        End With
    Next
    
    If intLoop < 8 And Me.opt单板多项.Value = True Then
        MsgBox "你选择了单版多项目，但还有项目没有选择！", vbInformation
        Me.vsList.Select 1, 13
        Exit Sub
    End If
    
    If strResult = ";;;;;;;" Then
        MsgBox "没有选择检验项目，请选择检查项目!", vbInformation
        Me.vsList.Select 1, 13
        Exit Sub
    End If
    
    '结果
    For intRow = 1 To 8
        strResult = strResult & "|"
        For intCol = 1 To 12
            With Me.vsList
                strNumber = .TextMatrix(intRow, intCol)
                If intCol = 1 Then
                    strResult = strResult & strNumber
                Else
                    strResult = strResult & ";" & strNumber
                End If
            End With
        Next
    Next
    
    frmLabMBTemplet.ShowMe Me, strResult
    RefreshTemplet
End Sub

Private Sub RefreshTemplet()
    '功能   刷新当前模板
    Dim rsTmp As New adodb.Recordset
    '写入酶标模板数据
    gstrSql = "select id,编号,名称 from 检验酶标模板 order by 编号"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.cbo选择模板.Clear
    Me.cbo选择模板.AddItem ""
    Me.cbo选择模板.ItemData(Me.cbo选择模板.NewIndex) = 0
    Do Until rsTmp.EOF
        With Me.cbo选择模板
            .AddItem rsTmp("编号") & "-" & rsTmp("名称")
            .ItemData(.NewIndex) = rsTmp("ID")
        End With
        rsTmp.MoveNext
    Loop
End Sub

Private Sub MBcontrol()
    '功能 打开酶标控制界面
    Dim strControl As String
    Dim strResult As String
    Dim intRow As Integer, intCol As Integer
    Dim aRow() As String, aCol() As String
    
    If Me.cbo检验仪器.Text = "" Then
        MsgBox "请选择一个仪器!", vbInformation, Me.Caption
        Me.cbo检验仪器.SetFocus
        Exit Sub
    End If
    
    If Me.cbo波长.Text = "" Then
        MsgBox "请选择波长!    ", vbInformation, Me.Caption
        Me.cbo波长.SetFocus
        Exit Sub
    End If

    If Me.cbo振板频率.Text = "" Then
        MsgBox "请选择振板频率!", vbInformation, Me.Caption
        Me.cbo振板频率.SetFocus
        Exit Sub
    End If

    If Me.txt振板时间.Text = "" Then
        MsgBox "请选择振板时间!", vbInformation, Me.Caption
        Me.cbo振板频率.SetFocus
        Exit Sub
    End If

    If Me.cbo进板方式.Text = "" Then
        MsgBox "请选择进板方式!", vbInformation, Me.Caption
        Me.cbo振板频率.SetFocus
        Exit Sub
    End If

    If Me.cbo空白形式.Text = "" Then
        MsgBox "请选择空白形式!", vbInformation, Me.Caption
        Me.cbo振板频率.SetFocus
        Exit Sub
    End If
    
    Me.opt显示(0).Value = True
    For intRow = 1 To 8
        For intCol = 1 To 12
            mTestData(0, intRow, intCol) = Me.vsList.TextMatrix(intRow, intCol)
        Next
    Next
    strControl = Me.cbo波长.Text & ";" & Me.cbo振板频率.Text & ";" & Me.txt振板时间 & _
                 ";" & Me.cbo进板方式.Text & ";" & Me.cbo空白形式 & ";" & Me.cbo参考波长.Text
                 
    frmLabMBControl.ShowMe Me, strControl, strResult
    
    If strResult = "" Then MsgBox "没有采集到数据，请重新测量!": Exit Sub
    
    aRow = Split(strResult, "|")
    For intRow = 1 To 8
        aCol = Split(aRow(intRow - 1), ";")
        For intCol = 1 To 12
            With Me.vsList
                If Trim(.TextMatrix(intRow, intCol)) <> "" Then
                    .Cell(flexcpData, intRow, intCol, intRow, intCol) = aCol(intCol - 1)
                    .TextMatrix(intRow, intCol) = Format(aCol(intCol - 1), "##0.000#")
                    mTestData(1, intRow, intCol) = Format(aCol(intCol - 1), "##0.000#")
                End If
            End With
        Next
    Next
    Me.opt显示(1).Value = True
    Me.cbo选择模板.ListIndex = 0
    '测量完成就计算
    Call CalcData
End Sub
Private Sub AddNew()
    '功能           增加一个新版
    Dim rsTmp As New adodb.Recordset
    
    '没有选择仪器时退出
    If Me.cbo检验仪器.Text = "" Then MsgBox "请先选择仪器!", vbInformation: Me.cbo检验仪器.SetFocus: Exit Sub
'    ReDim mTestData(0 To 3, 1 To 8, 1 To 12)
    Erase mTestData
    Erase mTestItem
    Erase mTestReagent
    
    mEditState = 1
    
    Me.opt显示(0).Value = True
    Call InitItem
    gstrSql = "select count(*) +1  from 检验酶标记录 where 测试时间 between to_date(to_char(sysdate,'yyyy-MM-dd') || ' 00:00:00','yyyy-MM-dd HH24:mi:ss')" & vbNewLine & _
                "              and to_date(to_char(sysdate,'yyyy-MM-dd') || ' 23:59:59','yyyy-MM-dd HH24:mi:ss') and 仪器id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngMachine)
    Me.txt测试板号.Text = Format(Now, "yyyymmdd") & "_" & rsTmp(0)
    Me.cbo选择模板.ListIndex = 0
    mblnMBSelect = False
End Sub
Private Sub SaveData()
    '功能           保存当前测量数据
    
    Dim intRow As Integer, intCol As Integer
    Dim strNumber As String             '用于检查的数据
    Dim strItem As String               '项目数据
    Dim strData As String               '结果数据
    Dim strCalcOne As String            '阳性公式
    Dim strCalcTwo As String            '弱阳性公式
    Dim strCalcThree As String          'CutOff公式
    Dim str试剂 As String               '试剂
    Dim lngKey As Long
    
    Dim rsTmp As New adodb.Recordset
    
    On Error GoTo errH
    
    '检查数据是否完整
    For intRow = 1 To 8
        strData = strData & "|"
        For intCol = 1 To 13
            If intCol < 13 Then
                strNumber = strNumber & ";" & mTestData(0, intRow, intCol)
                strData = strData & ";" & mTestData(0, intRow, intCol) & "^" & mTestData(1, intRow, intCol)
            Else
                '有编码没有项目时
'                If Len(strNumber) > 1 And Me.vsList.Cell(flexcpText, intRow, intCol) = "" Then
'                    MsgBox "请选择一个检验项目或取消当前行的编码!", vbInformation, Me.Caption
'                    Me.vsList.Select intRow, intCol
'                    Exit Sub
'                End If
                
                '单版多项目时有项目没有编码时
                If Len(strNumber) = 1 And Me.vsList.Cell(flexcpText, intRow, intCol) <> "" Then
                    MsgBox "请选择当前行的编码!", vbInformation, Me.Caption
                    Me.vsList.Select intRow, 1
                    Exit Sub
                End If
                strItem = strItem & ";" & Me.vsList.Cell(flexcpText, intRow, intCol)
                
                mrsCalc.filter = "诊治项目id=" & Val(Me.vsList.TextMatrix(intRow, intCol))
                
                '阳性公式
                If mTestItem(0, intRow) <> "" Then
                    strCalcOne = strCalcOne & ";" & mTestItem(0, intRow)
                Else
                    If mrsCalc.EOF = False Then
                        strCalcOne = strCalcOne & ";" & Nvl(mrsCalc("阳性公式"))
                    Else
                        strCalcOne = strCalcOne & ";"
                    End If
                End If
                
                '弱阳性公式
                If mTestItem(1, intRow) <> "" Then
                    strCalcTwo = strCalcTwo & ";" & mTestItem(1, intRow)
                Else
                    If mrsCalc.EOF = False Then
                        strCalcTwo = strCalcTwo & ";" & Nvl(mrsCalc("弱阳性公式"))
                    Else
                        strCalcTwo = strCalcTwo & ";"
                    End If
                End If
                
                '阳性公式
                If mTestItem(2, intRow) <> "" Then
                    strCalcThree = strCalcThree & ";" & mTestItem(2, intRow)
                Else
                    If mrsCalc.EOF = False Then
                        strCalcThree = strCalcThree & ";" & Nvl(mrsCalc("CutOff公式"))
                    Else
                        strCalcThree = strCalcThree & ";"
                    End If
                End If
            End If
        Next
    Next
    
    If mEditState = 1 Then
        lngKey = zlDatabase.GetNextId("检验酶标记录")
    Else
        lngKey = mlngKey
    End If
    
    str试剂 = Join(mTestReagent, "|")
    If Replace(str试剂, "|", "") = "" Then
        str试剂 = ""
    End If
    
    '开始保存
    gstrSql = "Zl_检验酶标记录_Insert(" & lngKey & ",'" & Me.txt测试板号 & "'," & _
                "to_date('" & Me.dtp测试时间.Value & "','yyyy-MM-dd HH24:MI:ss')" & ",'" & Me.cbo波长.Text & "','" & _
                Me.cbo参考波长.Text & "','" & Me.cbo振板频率.Text & "','" & Me.txt振板时间.Text & "','" & _
                Me.cbo进板方式.Text & "','" & Me.cbo空白形式.Text & "','" & txt试剂批号.Tag & "'," & _
                IIf(Me.txt试剂效期 <> "", "to_date('" & Me.txt试剂效期.Text & "','yyyy-MM-dd HH:MI:ss')", "NULL") & _
                ",'" & Me.txt试剂厂商.Text & "','" & _
                Me.txt测试方法.Text & "'," & IIf(Me.opt单板单项.Value = True, 1, 0) & ",'" & Me.txt存放位置.Text & "','" & Mid(strItem, 2) & "','" & _
                strCalcOne & "','" & strCalcTwo & "','" & strCalcThree & "','" & Mid(strData, 2) & "','" & str试剂 & "'," & _
                Me.cbo检验仪器.ItemData(Me.cbo检验仪器.ListIndex) & ")"
    
    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
    
    '保存当前仪器设置到注册
    On Error Resume Next
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "波长", cbo波长.Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "参考波长", cbo参考波长.Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "振板频率", cbo振板频率.Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "振板时间", txt振板时间.Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "进板方式", cbo进板方式.Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "空白形式", cbo空白形式.Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "项目ID", Me.cbo测试项目.ItemData(Me.cbo测试项目.NewIndex))
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "空白形式", cbo空白形式.Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & App.ProductName, "横向纵向", opt方向(0).Value)
    zlDatabase.SetPara "frmLabMB_波长", cbo波长.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_参考波长", cbo参考波长.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_振板频率", cbo振板频率.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_振板时间", txt振板时间.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_进板方式", cbo进板方式.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_空白形式", cbo空白形式.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_项目ID", Me.cbo测试项目.ItemData(Me.cbo测试项目.NewIndex), 100, 1208
    On Error GoTo 0
    mEditState = 0
    mlngKey = lngKey
    Call SendData               '发送到技师工作站
    Call RefreshList
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub RefreshList(Optional intType As Integer = 1, Optional strFilter As String)
    '功能               刷新左边列表
    '参数               intType 刷新参数(用于区分快速过滤和过滤）
    '                       1=快速过滤
    '                       2=过滤(第二个中的过滤条件进行过滤) 格式:"板号;试剂批号:是否使用时间查询,开始时间,结束时间"
    '                       3=通过ID查找ID间使用","进行分隔"
    '                   Strfilter 过滤字串
    
    Dim rsTmp As New adodb.Recordset
    Dim Record As ReportRecord
    Dim intLoop As Integer
    Dim strBeginDate As String, strEndDate As String
    Dim strDate As String
    Dim aItem() As String
    Dim aRow() As String
    Dim strWhere As String
    Dim strReagentNo As String          '试剂批号
    Dim strReagentDate As String        '试剂效期
    Dim strReagentManufacturer          '厂商
    Dim strReagentMeans                 '方法

    
    On Error GoTo errH
    
    gstrSql = "select ID,板号,测试时间,波长,参考波长,振板频率,振板时间,进板方式,空白形式,试剂批号,试剂效期,试剂厂商,测试方法,是否发送," & vbNewLine & _
              "       存放位置,测试项目,阳性公式,弱阳性公式,CutOff公式,测试结果,试剂记录 from 检验酶标记录 a "
                  
    strBeginDate = Format(GetDateTime("今  天", 1), "yyyy-MM-dd 00:00:00")
    strEndDate = Format(GetDateTime("今  天", 2), "yyyy-MM-dd 23:59:59")
    If intType = 1 Then
        '快速过滤
        gstrSql = gstrSql & " Where 测试时间 between [1] and [2] "
        If Me.opt过滤(0).Value = True Then
            '今天
            strBeginDate = Format(GetDateTime("今  天", 1), "yyyy-MM-dd 00:00:00")
            strEndDate = Format(GetDateTime("今  天", 2), "yyyy-MM-dd 23:59:59")
        ElseIf Me.opt过滤(1).Value = True Then
            '本周
            strBeginDate = Format(GetDateTime("本  周", 1), "yyyy-MM-dd 00:00:00")
            strEndDate = Format(GetDateTime("本  周", 2), "yyyy-MM-dd 23:59:59")
        ElseIf Me.opt过滤(2).Value = True Then
            '本月
            strBeginDate = Format(GetDateTime("本  月", 1), "yyyy-MM-dd 00:00:00")
            strEndDate = Format(GetDateTime("本  月", 2), "yyyy-MM-dd 23:59:59")
        ElseIf Me.opt过滤(3).Value = True Then
            '本年
            strBeginDate = Format(GetDateTime("本  年", 1), "yyyy-MM-dd 00:00:00")
            strEndDate = Format(GetDateTime("本  年", 2), "yyyy-MM-dd 23:59:59")
        End If
        gstrSql = gstrSql & " And 仪器ID = [3] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(strBeginDate), CDate(strEndDate), _
        Me.cbo检验仪器.ItemData(Me.cbo检验仪器.ListIndex))
    ElseIf intType = 2 Then
        '过滤
        aItem = Split(strFilter, ";")
        
        If aItem(0) <> "" Then
            '板号
            strWhere = " where 板号 = [3] "
        End If
                
        If aItem(1) <> "" Then
            '试剂批号
            If strWhere = "" Then
                strWhere = " Where 试剂批号 = [4] "
            Else
                strWhere = strWhere & " And 试剂批号 = [4] "
            End If
        End If
        
        If Split(aItem(2), ",")(0) = 1 Then
            '试剂批号
            If strWhere = "" Then
                strWhere = " Where 测试时间 between [1] and [2] "
            Else
                strWhere = strWhere & " And 测试时间 between [1] and [2] "
            End If
            strBeginDate = Split(aItem(2), ",")(1)
            strEndDate = Split(aItem(2), ",")(2)
        End If
        
        gstrSql = gstrSql & strWhere
        
        
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(strBeginDate), CDate(strEndDate), aItem(0), aItem(1), _
                    Me.cbo检验仪器.ItemData(Me.cbo检验仪器.ListIndex))
    ElseIf intType = 3 Then
        If strFilter <> "" Then
            gstrSql = gstrSql & " , (Select * From Table(Cast(f_str2list([1]) As zltools.t_strlist))) H where a.id = h.Column_Value"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(strBeginDate), CDate(strEndDate), aItem(0), aItem(1))
        End If
    End If
    
    Me.rptList.Records.DeleteAll
    Do Until rsTmp.EOF
        Set Record = Me.rptList.Records.Add
        
        For intLoop = 0 To Me.rptList.Columns.Count
            Record.AddItem ""
        Next
        
        Record.Item(mCol.ID).Value = Nvl(rsTmp("ID"))
        Record.Item(mCol.板号).Value = Nvl(rsTmp("板号"))
        Record.Item(mCol.测试时间).Value = Nvl(rsTmp("测试时间"))
'        Record.Item(mCol.试剂批号).Value = Nvl(rsTmp("试剂批号"))
'        Record.Item(mCol.试剂效期).Value = Nvl(rsTmp("试剂效期"))
'        Record.Item(mCol.试剂厂商).Value = Nvl(rsTmp("试剂厂商"))
'        Record.Item(mCol.测试方法).Value = Nvl(rsTmp("测试方法"))
        Record.Item(mCol.波长).Value = Nvl(rsTmp("波长"))
        Record.Item(mCol.参考波长).Value = Nvl(rsTmp("参考波长"))
        Record.Item(mCol.振板频率).Value = Nvl(rsTmp("振板频率"))
        Record.Item(mCol.振板时间).Value = Nvl(rsTmp("振板时间"))
        Record.Item(mCol.进板方式).Value = Nvl(rsTmp("进板方式"))
        Record.Item(mCol.空白形式).Value = Nvl(rsTmp("空白形式"))
        Record.Item(mCol.存放位置).Value = Nvl(rsTmp("存放位置"))
        Record.Item(mCol.测试项目).Value = Nvl(rsTmp("测试项目"))
        Record.Item(mCol.阳性公式).Value = Nvl(rsTmp("阳性公式"))
        Record.Item(mCol.弱阳性公式).Value = Nvl(rsTmp("弱阳性公式"))
        Record.Item(mCol.CutOff公式).Value = Nvl(rsTmp("CutOff公式"))
        Record.Item(mCol.测试结果).Value = Nvl(rsTmp("测试结果"))
        Record.Item(mCol.试剂记录).Value = Nvl(rsTmp("试剂记录"))
        
        If Replace(Record.Item(mCol.试剂记录).Value, "|", "") <> "" Then
            strReagentNo = "": strReagentDate = "": strReagentManufacturer = "": strReagentMeans = ""
            '写入试剂记录
            aRow = Split(Record.Item(mCol.试剂记录).Value, "|")
            
            For intLoop = 0 To UBound(aRow)
                aItem = Split(aRow(intLoop), ";")
                If UBound(aItem) >= 3 Then
                    '试剂
                    If InStr(strReagentNo, "<" & aItem(0) & ">") <= 0 Then
                        strReagentNo = strReagentNo & "<" & aItem(0) & ">"
                    End If
                    '效期
                    If InStr(strReagentDate, "<" & aItem(1) & ">") <= 0 Then
                        strReagentDate = strReagentDate & "<" & aItem(1) & ">"
                    End If
                    '厂商
                    If InStr(strReagentManufacturer, "<" & aItem(2) & ">") <= 0 Then
                        strReagentManufacturer = strReagentManufacturer & "<" & aItem(2) & ">"
                    End If
                    '方法
                    If InStr(strReagentMeans, "<" & aItem(3) & ">") <= 0 Then
                        strReagentMeans = strReagentMeans & "<" & aItem(3) & ">"
                    End If
                End If
            Next
            Record.Item(mCol.试剂批号).Value = strReagentNo
            Record.Item(mCol.试剂效期).Value = strReagentDate
            Record.Item(mCol.试剂厂商).Value = strReagentManufacturer
            Record.Item(mCol.测试方法).Value = strReagentMeans
        End If
        rsTmp.MoveNext
    Loop
    Me.rptList.Populate
    If mlngKey = 0 Then
        If Me.rptList.Rows.Count > 0 Then
            Call RefreshItem(Me.rptList.Rows(0).Record(mCol.ID).Value)
        End If
    Else
        Call RefreshItem(mlngKey)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub RefreshItem(lngKey As Long)
    '功能           刷新当前项目
    Dim rsTmp As New adodb.Recordset
    Dim lngLoop As Long, intLoop As Integer
    Dim aItem() As String
    Dim aRow() As String, aCol() As String
    Dim intRow As Integer, intCol As Integer
    Dim aRule() As String
    
    On Error GoTo errH
    
    Call InitItem
    Erase mTestItem
    gstrSql = "select 诊治项目ID,阳性公式,弱阳性公式,CutOff公式 from 检验项目 where 项目类别 = 4"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
    With Me.rptList
        For lngLoop = 0 To .Rows.Count - 1
            If .Rows(lngLoop).Record(mCol.ID).Value = lngKey Then
                '找到ID后写入
                On Error Resume Next
                dtp测试时间 = .Rows(lngLoop).Record(mCol.测试时间).Value
                cbo波长.Text = .Rows(lngLoop).Record(mCol.波长).Value
                cbo参考波长.Text = .Rows(lngLoop).Record(mCol.参考波长).Value
                cbo振板频率.Text = .Rows(lngLoop).Record(mCol.振板频率).Value
                txt振板时间.Text = .Rows(lngLoop).Record(mCol.振板时间).Value
                cbo进板方式.Text = .Rows(lngLoop).Record(mCol.进板方式).Value
                cbo空白形式.Text = .Rows(lngLoop).Record(mCol.空白形式).Value
                txt测试板号.Text = .Rows(lngLoop).Record(mCol.板号).Value
'                cbo试剂批号.Text = .Rows(lngLoop).Record(mCol.试剂批号).Value
'                txt试剂效期.Text = .Rows(lngLoop).Record(mCol.试剂效期).Value
'                txt试剂厂商.Text = .Rows(lngLoop).Record(mCol.试剂厂商).Value
'                txt测试方法.Text = .Rows(lngLoop).Record(mCol.测试方法).Value
                txtCutOff公式.Text = .Rows(lngLoop).Record(mCol.CutOff公式).Value
                txt存放位置.Text = .Rows(lngLoop).Record(mCol.存放位置).Value
                
                
                
                aItem = Split(.Rows(lngLoop).Record(mCol.测试项目).Value, ";")
                For intLoop = 0 To UBound(aItem)
                    Me.vsList.TextMatrix(intLoop + 1, 13) = aItem(intLoop)
                    If intLoop > 0 Then
                        If Me.vsList.TextMatrix(intLoop, 13) <> Me.vsList.TextMatrix(intLoop + 1, 13) Then
                            '判断是否是一版多项目
                            Me.opt单板多项.Value = True
                        End If
                    End If
                Next
                
                '取检验项目中的公式
                aItem = Split(.Rows(lngLoop).Record(mCol.测试项目).Value, ";")
                
                
                For intRow = 1 To 8
                    '阳性公式
                    aRule = Split(Mid(.Rows(lngLoop).Record(mCol.阳性公式).Value, 2), ";")
                    mTestItem(0, intRow) = aRule(intRow - 1)
                    
                    '弱阳性公式
                    aRule = Split(Mid(.Rows(lngLoop).Record(mCol.弱阳性公式).Value, 2), ";")
                    mTestItem(1, intRow) = aRule(intRow - 1)
                    
                    'CutOff公式
                    aRule = Split(Mid(.Rows(lngLoop).Record(mCol.CutOff公式).Value, 2), ";")
                    mTestItem(2, intRow) = aRule(intRow - 1)
                    
                    rsTmp.filter = "诊治项目id=" & aItem(intRow - 1)
                    If rsTmp.EOF = False Then
                        For intLoop = 0 To Me.cbo测试项目.ListCount - 1
                            If Me.cbo测试项目.ItemData(intLoop) = rsTmp("诊治项目ID") Then
                                Me.vsList.Row = intRow
                                Me.cbo测试项目.ListIndex = intLoop
                            End If
                        Next
                    End If
                Next
                    

                On Error GoTo 0
'                If .Rows(lngLoop).Record(mCol.阳性公式).Value <> "" Then
'                    txt阳性公式.Text = Split(.Rows(lngLoop).Record(mCol.阳性公式).Value, ";")(1)
'                End If
'                If .Rows(lngLoop).Record(mCol.弱阳性公式).Value <> "" Then
'                    txt弱阳性公式.Text = Split(.Rows(lngLoop).Record(mCol.弱阳性公式).Value, ";")(1)
'                End If
'                If .Rows(lngLoop).Record(mCol.CutOff公式).Value <> "" Then
'                    txtCutOff公式.Text = Split(.Rows(lngLoop).Record(mCol.CutOff公式).Value, ";")(1)
'                End If
                
                
                aRow = Split(.Rows(lngLoop).Record(mCol.测试结果).Value, "|")
                
                For intRow = 1 To 8
                    aCol = Split(aRow(intRow - 1), ";")
                    For intCol = 1 To 12
                        With Me.vsList
                            mTestData(0, intRow, intCol) = Split(aCol(intCol), "^")(0)
                            mTestData(1, intRow, intCol) = Split(aCol(intCol), "^")(1)
                        End With
                    Next
                Next
                For intLoop = 0 To opt显示.Count - 1
                    If opt显示(intLoop).Value = True Then
                        Call opt显示_Click(intLoop)
                    End If
                Next
                Erase mTestReagent
                aItem = Split(.Rows(lngLoop).Record(mCol.试剂记录).Value, "|")
                For intLoop = 0 To UBound(aItem)
                    mTestReagent(intLoop + 1) = aItem(intLoop)
                Next
            End If
        Next
    End With
            
    Call CalcData
    Me.vsList.Row = 1
    Call vsList_Click
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DelData()
    '功能           删除数据
    
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    
    If MsgBox("是否确定要删除板号为<" & Me.rptList.FocusedRow.Record(mCol.板号).Value & ">的结果!", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo errH
    gstrSql = "Zl_检验酶标记录_Delete(" & Me.rptList.FocusedRow.Record(mCol.ID).Value & ")"
    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
    
    RefreshList
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
    
Private Sub CalcData(Optional OneCalc As Long)
    '功能           记算空白(BC)、阴性(NC)、阳性(PC)、质控(QC)
    Dim intLoop As Integer, lngLoop As Long
    Dim strItem As String
    Dim intRow As Integer, intCol As Integer
    Dim strBC As String, strNC As String, strPC As String, strQC As String
    Dim aBC() As String, aNC() As String, aPC() As String, aQC() As String
    Dim dblBC As Double, dblNC As Double, dblPC As Double, dblQC As Double
    Dim str阳性 As String, str弱阳性 As String, strCutOff As String
    Dim rsTmp As New adodb.Recordset
    Dim aItem() As String
    Dim bln减空白对照 As Boolean
    Dim bln小于阴性对照 As Boolean
    Dim str阴性对照 As String
    Dim intCount As Integer
    
    On Error GoTo errH
    
    '批量调整OD值
    If mstr公式 <> "" Then
        For intRow = 1 To 8
            For intCol = 1 To 12
                If mTestData(1, intRow, intCol) <> "" Then
                    mTestData(1, intRow, intCol) = Calc.Eval(Replace(UCase(mstr公式), "R", mTestData(1, intRow, intCol)))
                    mTestData(1, intRow, intCol) = Format(mTestData(1, intRow, intCol), "##0.000#")
                End If
            Next
        Next
    End If
    mstr公式 = ""
    
    
'    bln减空白对照 = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\frmLabMB", "减空白对照", "0")
    bln减空白对照 = zlDatabase.GetPara("frmLabMB_减空白对照", 100, 1208, "0")
    bln小于阴性对照 = chk阴性对照.Value
    str阴性对照 = Trim(txt最小阴性对照)
    
'    gstrSql = "select 诊治项目ID,阳性公式,弱阳性公式,CutOff公式 from 检验项目 where 项目类别 = 4"
'    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
    '找出已做项目，去掉重复的项目
    For intLoop = 1 To 8
        With Me.vsList
            If InStr(strItem & ";", ";" & .TextMatrix(intLoop, 13) & ";") <= 0 Then
                strItem = strItem & ";" & .TextMatrix(intLoop, 13)
            End If
        End With
    Next
    
    If OneCalc <> 0 Then strItem = ";" & OneCalc
    
    '开始记算
    aItem = Split(Mid(strItem, 2), ";")
    For lngLoop = 0 To UBound(aItem)
        strBC = "": strNC = "": strPC = "": strQC = ""
        dblBC = 0: dblNC = 0: dblPC = 0: dblQC = 0
        For intRow = 1 To 8
            For intCol = 1 To 12
                If Me.vsList.TextMatrix(intRow, 13) = aItem(lngLoop) Then
                    If InStr(mTestData(0, intRow, intCol), "BC") Then
                        strBC = strBC & ";" & mTestData(1, intRow, intCol)
                    ElseIf InStr(mTestData(0, intRow, intCol), "NC") Then
                        strNC = strNC & ";" & mTestData(1, intRow, intCol)
                    ElseIf InStr(mTestData(0, intRow, intCol), "PC") Then
                        strPC = strPC & ";" & mTestData(1, intRow, intCol)
                    ElseIf InStr(mTestData(0, intRow, intCol), "QC") Then
                        strQC = strQC & ";" & mTestData(1, intRow, intCol)
                    End If
                End If
            Next
        Next
        'BC
        If Trim(strBC) <> "" Then
            aBC = Split(Mid(strBC, 2), ";")
            For intLoop = 0 To UBound(aBC)
                dblBC = dblBC + Val(aBC(intLoop))
            Next
            If dblBC = 0 Then
                Me.txt空白对照.Text = 0
            Else
                Me.txt空白对照.Text = dblBC / Val((UBound(aBC)) + 1)
            End If
            Me.txt空白对照.Text = Format(Me.txt空白对照.Text, "##0.000#")
        Else
            Me.txt空白对照.Text = Format(0, "##0.000#")
        End If
        'NC
        If Trim(strNC) <> "" Then
            aNC = Split(Mid(strNC, 2), ";")
            For intLoop = 0 To UBound(aNC)
                dblNC = dblNC + Val(aNC(intLoop))
            Next
            If dblNC = 0 Then
                Me.txt阴性对照.Text = 0
            Else
                Me.txt阴性对照.Text = dblNC / Val((UBound(aNC)) + 1)
            End If
            Me.txt阴性对照.Text = Format(Me.txt阴性对照.Text - IIf(bln减空白对照, Me.txt空白对照.Text, 0), "##0.000#")
            
        Else
            Me.txt阴性对照.Text = Format(0, "##0.000#")
        End If
        If bln小于阴性对照 = True And Val(Me.txt阴性对照.Text) <= Val(str阴性对照) And Val(str阴性对照) <> 0 Then
            Me.txt阴性对照.Text = Format(Val(str阴性对照), "##0.000#")
        End If
        'PC
        If Trim(strPC) <> "" Then
            aPC = Split(Mid(strPC, 2), ";")
            For intLoop = 0 To UBound(aPC)
                dblPC = dblPC + Val(aPC(intLoop))
            Next
            If dblPC = 0 Then
                Me.txt阳性对照.Text = 0
            Else
                Me.txt阳性对照.Text = dblPC / Val((UBound(aPC)) + 1)
            End If
            Me.txt阳性对照.Text = Format(Me.txt阳性对照 - IIf(bln减空白对照, Me.txt空白对照.Text, 0), "##0.000#")
        Else
            Me.txt阳性对照.Text = Format(0, "##0.000#")
        End If
        'QC
        If Trim(strQC) <> "" Then
            aQC = Split(Mid(strQC, 2), ";")
            For intLoop = 0 To UBound(aQC)
                dblQC = dblQC + Val(aQC(intLoop))
            Next
            If dblQC = 0 Then
                dblQC = 0
            Else
                dblQC = dblQC / Val((UBound(aQC)) + 1)
            End If
            dblQC = Format(dblQC - IIf(bln减空白对照, Me.txt空白对照.Text, 0), "##0.000#")
        Else
            dblQC = Format(0, "##0.000#")
        End If
        
        For intRow = 1 To 8
            If Me.vsList.TextMatrix(intRow, 13) = aItem(lngLoop) Then
                strCutOff = mTestItem(2, intRow)
            End If
        Next
        If strCutOff <> "" Then
            strCutOff = Replace(strCutOff, "BC", Me.txt空白对照.Text)
            strCutOff = Replace(strCutOff, "NC", Me.txt阴性对照.Text)
            strCutOff = Replace(strCutOff, "PC", Me.txt阳性对照.Text)
            strCutOff = Replace(strCutOff, "QC", dblQC)
            strCutOff = Calc.Eval(strCutOff)
            Me.txtCutOff.Text = Format(strCutOff, "##0.000#")
        End If
        
        '计算阳性和弱阳性计算
        For intRow = 1 To 8
            For intCol = 1 To 12
                If Me.vsList.TextMatrix(intRow, 13) = aItem(lngLoop) Then
                    If IsNumeric(mTestData(0, intRow, intCol)) = True Then
                        '只处理普通标本
                        With Me.vsList
                            '阳性
'                            str阳性 = mTestItem(0, intRow)
                            mrsCalc.filter = "诊治项目ID=" & Val(Me.vsList.TextMatrix(intRow, 13))
                            If mrsCalc.EOF = False Then
                                str阳性 = mrsCalc("阳性公式") & ""
                            Else
                                str阳性 = ""
                            End If
                            If Trim(str阳性) <> "" Then
                                str阳性 = Replace(str阳性, "NC", Me.txt阴性对照.Text)
                                str阳性 = Replace(str阳性, "PC", Me.txt阳性对照.Text)
                                str阳性 = Replace(str阳性, "BC", Me.txt空白对照.Text)
                                str阳性 = Replace(str阳性, "QC", dblQC)
                                str阳性 = Replace(str阳性, "OD", Val(mTestData(1, intRow, intCol)) - IIf(bln减空白对照, Me.txt空白对照.Text, 0))
                                If mTestData(1, intRow, intCol) <> "" Then
                                    mTestData(3, intRow, intCol) = IIf(Calc.Eval(str阳性), "阳性(+)", "阴性(-)")
                                End If
        '                        .TextMatrix(intRow, intCol) = IIf(Calc.Eval(str阳性), "阳性(+)", "阴性(-)")
                            End If
                            '弱阳性
                            mrsCalc.filter = "诊治项目ID=" & Val(Me.vsList.TextMatrix(intRow, 13))
                            If mrsCalc.EOF = False Then
                                str弱阳性 = mrsCalc("弱阳性公式") & ""
                            Else
                                str弱阳性 = ""
                            End If
                                
                            If str弱阳性 <> "" And mTestData(3, intRow, intCol) <> "阳性(+)" Then
                                str弱阳性 = Replace(str弱阳性, "NC", Me.txt阴性对照.Text)
                                str弱阳性 = Replace(str弱阳性, "PC", Me.txt阳性对照.Text)
                                str弱阳性 = Replace(str弱阳性, "BC", Me.txt空白对照.Text)
                                str弱阳性 = Replace(str弱阳性, "QC", dblQC)
                                str弱阳性 = Replace(str弱阳性, "OD", Val(mTestData(1, intRow, intCol)) - IIf(bln减空白对照, Me.txt空白对照.Text, 0))
                                If mTestData(1, intRow, intCol) <> "" Then
                                    mTestData(3, intRow, intCol) = IIf(Calc.Eval(str弱阳性), "弱阳性(±)", "阴性(-)")
                                    
                                End If
        '                        .TextMatrix(intRow, intCol) = IIf(Calc.Eval(str阳性), "弱阳性(+-)", "阴性(-)")
                            End If
                        End With
                        '计算减去空白的OD
                        If Me.txt空白对照.Text <> "" And mTestData(1, intRow, intCol) <> "" Then
                            mTestData(2, intRow, intCol) = Format(mTestData(1, intRow, intCol) - IIf(bln减空白对照, Me.txt空白对照.Text, 0), "##0.000#")
                        End If
                    Else
                        If mTestData(1, intRow, intCol) <> "" Then
'                            If InStr(mTestData(0, intRow, intCol), "BC") = 0 Then
                                mTestData(2, intRow, intCol) = Format(mTestData(1, intRow, intCol) - IIf(bln减空白对照, Me.txt空白对照.Text, 0), "##0.000#")
                                mTestData(3, intRow, intCol) = Format(mTestData(1, intRow, intCol) - IIf(bln减空白对照, Me.txt空白对照.Text, 0), "##0.000#")
'                            Else
'                                mTestData(2, intRow, intCol) = Format(mTestData(1, intRow, intCol), "##0.000#")
'                                mTestData(3, intRow, intCol) = Format(mTestData(1, intRow, intCol), "##0.000#")
'                            End If
                        End If
                    End If
                End If
            Next
        Next
    Next
    
    For intLoop = 0 To Me.opt显示.Count - 1
        If Me.opt显示(intLoop).Value = True Then
            Call opt显示_Click(intLoop)
            Exit For
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitItem()
    '功能   清空当前编写界面内的内容
    Dim lngKey As Long
    Dim intLoop As Integer
    Dim intRow As Integer, intCol As Integer
    Dim rsTmp As New adodb.Recordset
    Dim aItem() As String
    Dim int排列 As Integer
    
    Me.dtp测试时间.Value = Now
    On Error GoTo errH
    If Me.cbo检验仪器.ListIndex = -1 Then Exit Sub
    
    gstrSql = "select 波长,振板频率,振板时间,进板方式,空白形式 from 检验仪器 where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.cbo检验仪器.ItemData(Me.cbo检验仪器.ListIndex))
    If rsTmp.EOF = True Then Exit Sub
    
    Me.txt空白对照.Text = ""
    Me.txt阴性对照.Text = ""
    Me.txt阳性对照.Text = ""
    Me.txtCutOff.Text = ""
    Me.txt阳性公式.Text = ""
    Me.txt弱阳性公式.Text = ""
    Me.txtCutOff公式.Text = ""
    Me.opt单板单项.Value = True
    
    With Me.cbo波长
        .Clear
        Me.cbo参考波长.Clear
        Me.cbo参考波长.AddItem ""
        Me.cbo参考波长.ItemData(Me.cbo参考波长.NewIndex) = 0
        aItem = Split(Nvl(rsTmp("波长")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
            Me.cbo参考波长.AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    With Me.cbo振板频率
        .Clear
        aItem = Split(Nvl(rsTmp("振板频率")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    Me.txt振板时间.Text = Nvl(rsTmp("振板时间"))
    
    With Me.cbo进板方式
        .Clear
        aItem = Split(Nvl(rsTmp("进板方式")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    With Me.cbo空白形式
        .Clear
        aItem = Split(Nvl(rsTmp("空白形式")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    '调出上次的使用参数
    On Error Resume Next
    cbo波长.Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "波长", "")
    cbo参考波长.Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "参考波长", "")
    cbo振板频率.Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "振板频率", "")
    'txt振板时间.Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "振板时间", Me.txt振板时间.Text)
    If txt振板时间.Text = "" Then
        Me.txt振板时间.Text = Nvl(rsTmp("振板时间"))
    End If
    cbo空白形式.Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "空白形式", "")
    cbo进板方式.Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "进板方式", "")
    cbo空白形式.Text = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "空白形式", "")
    int排列 = Val(GetSetting("ZLSOFT", "私有模块\" & App.ProductName, "横向纵向", 1))
    If int排列 = 1 Then
        Me.opt方向(0).Value = True
    Else
        Me.opt方向(1).Value = True
    End If
    lngKey = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\" & Me.Name, "项目ID", 0)
    
    
    
    cbo波长.Text = zlDatabase.GetPara("frmLabMB_波长", 100, 1208, "")
    cbo参考波长.Text = zlDatabase.GetPara("frmLabMB_参考波长", 100, 1208, "")
    cbo振板频率.Text = zlDatabase.GetPara("frmLabMB_振板频率", 100, 1208, "")
    txt振板时间.Text = zlDatabase.GetPara("frmLabMB_振板时间", 100, 1208, Me.txt振板时间.Text)
    If txt振板时间.Text = "" Then
        Me.txt振板时间.Text = Nvl(rsTmp("振板时间"))
    End If
    cbo进板方式.Text = zlDatabase.GetPara("frmLabMB_进板方式", 100, 1208, "")
    cbo空白形式.Text = zlDatabase.GetPara("frmLabMB_空白形式", 100, 1208, "")
    lngKey = zlDatabase.GetPara("frmLabMB_项目ID", 100, 1208, 0)
    Me.chk阴性对照.Value = Mid(zlDatabase.GetPara("frmLabMB_阴性对照", 100, 1208, "0,"), 1, 1)
    Me.txt最小阴性对照.Text = Mid(zlDatabase.GetPara("frmLabMB_阴性对照", 100, 1208, "0,"), 3)
'    On Error GoTo 0
'    If lngKey <> 0 Then
'        For intLoop = 0 To Me.cbo测试项目.ListCount - 1
'            If Me.cbo测试项目.ItemData(intLoop) = lngKey Then
'                Me.cbo测试项目.ListIndex = intLoop
'                Call cbo测试项目_Click
'                Exit For
'            End If
'        Next
'    End If
    
    For intRow = 1 To 8
        For intCol = 1 To 13
            With Me.vsList
                .TextMatrix(intRow, intCol) = ""
            End With
        Next
    Next
    txt试剂批号.Tag = ""
    txt试剂批号.Text = ""
    txt试剂效期.Text = ""
    txt试剂厂商.Text = ""
    txt测试方法.Text = ""
    Erase mTestReagent
    Erase mTestItem
    For intLoop = 1 To 8
        If Me.vsList.TextMatrix(intLoop, 13) <> "" Then
            Me.vsList.Row = intLoop
            For intRow = 0 To Me.cbo测试项目.ListCount - 1
                If Me.cbo测试项目.ItemData(intRow) = Me.vsList.TextMatrix(intLoop, 13) Then
                    Call cbo测试项目_Click
                    
                End If
            Next
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 0 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
'    If Me.rptList.Records.Count = 0 Then Exit Sub
'
'    '-------------------------------------------------
'    '复制数据表格
'    If zlReportToVSFlexGrid(Me.vfgList, Me.rptList) = False Then Exit Sub
'
'    '-------------------------------------------------
'    '调用打印部件处理
'    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
'
'    Set objPrint.Body = Me.vfgList
'    objPrint.Title.Text = "检验项目清单"
'    Set objAppRow = New zlTabAppRow
'    Call objAppRow.Add("")
'    Call objAppRow.Add("打印时间:" & Now())
'    Call objPrint.BelowAppRows.Add(objAppRow)
'
'    If bytMode = 1 Then
'        bytMode = zlPrintAsk(objPrint)
'        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
'    Else
'        zlPrintOrView1Grd objPrint, bytMode
'    End If
    Dim intRow As Integer, intCol As Integer
    Dim strSQL1 As String, strSQL2 As String, strSQL3 As String
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For intRow = 1 To 8
        strSQL1 = "Zl_检验酶标板打印_Insert('OD','" & Chr(65 + intRow - 1)
        strSQL2 = "Zl_检验酶标板打印_Insert('定性','" & Chr(65 + intRow - 1)
        strSQL3 = "Zl_检验酶标板打印_Insert('编号','" & Chr(65 + intRow - 1)
        For intCol = 1 To 12
            strSQL1 = strSQL1 & "','" & mTestData(1, intRow, intCol)
            strSQL2 = strSQL2 & "','" & mTestData(3, intRow, intCol)
            strSQL3 = strSQL3 & "','" & mTestData(3, intRow, intCol)
        Next
        strSQL1 = strSQL1 & "')"
        strSQL2 = strSQL2 & "')"
        strSQL3 = strSQL3 & "')"
        zlDatabase.ExecuteProcedure strSQL1, Me.Caption
        zlDatabase.ExecuteProcedure strSQL2, Me.Caption
        zlDatabase.ExecuteProcedure strSQL3, Me.Caption
    Next
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1208_7", Me, "酶标板ID=" & mlngKey, IIf(bytMode, 2, 1))
    
    gcnOracle.CommitTrans
    Exit Sub
errH:
    gcnOracle.RollbackTrans
End Sub
Private Sub SendData()
    '功能               发送酶标数据到技师工作站
    
    Dim rsTmp As New adodb.Recordset
    Dim intRow As Integer, intCol As Integer
    Dim strDate As String
    Dim lngMachine As Long
    Dim lngID As Long
    Dim lngDept As Long
    Dim strSampleType As String
    Dim strSex As String
    Dim strBirth As String
    Dim blnAuditing As Boolean
    Dim lngItemID As Long
    Dim str质控  As String
    Dim lngQCID As Long, i As Integer
    Dim strQCList() As String '保存需要计算的内容
    Dim blnBegin As Boolean
    Dim astrSQL() As String
    
    
    On Error GoTo errH
        
    If Me.cbo检验仪器.ListIndex = -1 Then Exit Sub
    lngMachine = Me.cbo检验仪器.ItemData(Me.cbo检验仪器.ListIndex)
    
    gstrSql = "select 使用小组ID from 检验仪器 where id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMachine)
    If rsTmp.EOF = True Then Call MsgBox("请在检验仪器里选择一个使用小组!", vbInformation): Exit Sub
    lngDept = rsTmp("使用小组ID")
    
    strDate = zlDatabase.Currentdate
    '处理为事务
    ReDim strQCList(0) As String
    ReDim astrSQL(0)

    blnBegin = True
    
    For intRow = 1 To 8
        gstrSql = "select 检验标本 from 检验报告项目 a , 诊疗项目目录 b where a.诊疗项目id = b.id and a.报告项目id = [1] and b.组合项目 = 0 "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Val(Me.vsList.TextMatrix(intRow, 13))))
        If rsTmp.EOF = False Then strSampleType = Nvl(rsTmp("检验标本"))
        '计算当前行的CutOFF值,从文本框里取
        Call CalcData(Val(Me.vsList.TextMatrix(intRow, 13)))
        For intCol = 1 To 12
            If (IsNumeric(mTestData(0, intRow, intCol)) = True Or UCase(Trim(mTestData(0, intRow, intCol))) Like "QC*") And IsNumeric(mTestData(1, intRow, intCol)) = True Then
                gstrSql = "Select a.*,Decode(c.性别,Null,0,'男',1,'女',2) As 性别,to_char(c.出生日期,'yyyy-mm-dd') As 出生日期 From 检验标本记录 a,病人医嘱记录 b,病人信息 c " & _
                        " Where a.医嘱id=b.id(+) And b.病人id=c.病人id(+)" & _
                        " And a.核收时间 Between [1] And [2]" & _
                        " And a.仪器ID=[3] And a.标本序号=[4] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "查询标本记录", CDate(Format(strDate, "yyyy-MM-dd") & " 00:00:00"), _
                        CDate(Format(strDate, "yyyy-MM-dd") & " 23:59:59"), lngMachine, mTestData(0, intRow, intCol))
                If rsTmp.EOF = True Then
                    strSex = 0: strBirth = ""
                    lngID = zlDatabase.GetNextId("检验标本记录")
                    str质控 = "0"
                    If UCase(Trim(mTestData(0, intRow, intCol))) Like "QC*" Then
                        str质控 = "1"
                    End If
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_INSERT(" & lngID & ",NULL,'" & _
                        mTestData(0, intRow, intCol) & "',NULL,NULL," & lngMachine & ",NULL," & _
                        "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),NULL," & _
                        "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & strSampleType & "'," & _
                        "Null,To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & UserInfo.姓名 & "','" & str质控 & "'," & lngDept & ",0,0)"
                Else
                    strSex = Nvl(rsTmp("性别"), 0)
                    strBirth = Nvl(rsTmp("出生日期"))
                    strSampleType = Nvl(rsTmp("标本类型"))
                    lngID = rsTmp("ID")
                    blnAuditing = Not IsNull(rsTmp("初审人"))
                    If blnAuditing = False Then
                        blnAuditing = Not IsNull(rsTmp("审核人"))
                    End If
                End If
                
                '只保存没有审核的标本
                If Not blnAuditing Then
'                    strItemRecords = Mid(strItemRecords, 2)
                        Dim strValue As String
                        If Val(Me.txtCutOff.Text) <> 0 Then
                            strValue = Format(Abs(Val(mTestData(2, intRow, intCol)) / Val(Me.txtCutOff.Text)), "##0.000#")
                        Else
                            strValue = "0"
                        End If
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        astrSQL(UBound(astrSQL)) = "ZL_检验普通结果_BATCHUPDATE(" & lngID & "," & _
                            lngMachine & ",'" & strSampleType & "'," & strSex & "," & _
                            IIf(strBirth = "", "Null", "To_Date('" & strBirth & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                            Me.vsList.TextMatrix(intRow, 13) & "^" & mTestData(3, intRow, intCol) & "^" & _
                            Format(Abs(Val(mTestData(2, intRow, intCol))), "##0.000#") & "^" & _
                            Format(Abs(Val(Me.txtCutOff.Text)), "##0.000#") & _
                            "^" & strValue & _
                           "',0," & mlngKey & ")"
                           ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                           astrSQL(UBound(astrSQL)) = "Zl_重新计算结果_Cale(" & lngID & ")"
                End If
                
                
                If lngID > 0 And UCase(Trim(mTestData(0, intRow, intCol))) Like "QC*" Then
                    lngQCID = SendQC(lngID, Trim(mTestData(0, intRow, intCol)))
                    '自动计算
                    If lngQCID > 0 Then
                        If strQCList(UBound(strQCList)) <> "" Then ReDim Preserve strQCList(UBound(strQCList) + 1)
                        strQCList(UBound(strQCList)) = Format(CDate(strDate), "yyyy-MM-dd") & "," & CStr(lngQCID)
                    End If
                End If
            End If
        Next
    Next
'    gcnOracle.BeginTrans
'    blnBegin = True
    For i = LBound(astrSQL) To UBound(astrSQL)
        If astrSQL(i) <> "" Then
            zlDatabase.ExecuteProcedure astrSQL(i), "发送到技师站"
        End If
    Next
'    gcnOracle.CommitTrans
    
'    gcnOracle.BeginTrans
'    blnBegin = True
    For i = LBound(strQCList) To UBound(strQCList)
        If InStr(strQCList(i), ",") > 0 Then
            Call AutoQCCompute(CDate(Split(strQCList(i), ",")(0)), Split(strQCList(i), ",")(1))
        End If
    Next
'    gcnOracle.CommitTrans
    
    Exit Sub
errH:
'    If blnBegin Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub


Private Function SendQC(ByVal lngID As Long, ByVal strSampleID As String) As Long
    '保存为质控标本
    
    Dim date当前日期 As Date, lngQCID As Long, str标本号 As String
    Dim var标本号 As Variant, iCoutn As Integer, lngDeviceID As Long
    Dim rsTmp As adodb.Recordset
    On Error GoTo errH
    lngQCID = 0
    date当前日期 = zlDatabase.Currentdate
    lngDeviceID = Me.cbo检验仪器.ItemData(Me.cbo检验仪器.ListIndex)
    
    gstrSql = "Select ID,标本号 From 检验质控品 Where [2] between 开始日期 and 结束日期 And 仪器id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "取质控品数据", lngDeviceID, date当前日期)
    
    Do Until rsTmp.EOF Or lngQCID <> 0
        str标本号 = "" & rsTmp.Fields("标本号")
        If InStr(str标本号, ",") > 0 Then
            var标本号 = Split(str标本号, ",")
            For iCoutn = 0 To UBound(var标本号)
                If var标本号(iCoutn) Like "*-*" Then
                    If strSampleID >= Val(Split(var标本号(iCoutn), "-")(0)) And strSampleID <= Val(Split(var标本号(iCoutn), "-")(1)) Then
                        lngQCID = rsTmp.Fields("ID")
                    End If
                Else
                    If var标本号(iCoutn) = strSampleID Then
                        lngQCID = rsTmp.Fields("ID")
                    End If
                End If
            Next
        ElseIf str标本号 Like "*-*" Then
            If strSampleID >= Val(Split(str标本号, "-")(0)) And strSampleID <= Val(Split(str标本号, "-")(1)) Then
                lngQCID = rsTmp.Fields("ID")
            End If
        Else
            If strSampleID = str标本号 Then
                lngQCID = rsTmp.Fields("ID")
            End If
        End If
        
        rsTmp.MoveNext
    Loop
    
    If lngQCID > 0 Then
        gstrSql = "ZL_检验质控记录_EDIT(1," & lngID & "," & lngQCID & ")"
        zlDatabase.ExecuteProcedure gstrSql, "保存为质控品"
        
        SendQC = lngQCID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub AutoQCCompute(ByVal date日期 As Date, ByVal str质控品 As String)

    '自动计算质控标本
    ' date日期 :质控计算日期
    ' str质控品 :质控品
    Dim rsTemp As adodb.Recordset, rsTmp As adodb.Recordset, strReturn As String
    Dim lngDeviceID As Long
    lngDeviceID = mlngMachine
    On Error GoTo errH
    gstrSql = "Select Distinct B.项目id, C.编码, C.中文名, C.英文名" & vbNewLine & _
              " From 检验质控品 A, 检验质控品项目 B, 诊治所见项目 C" & vbNewLine & _
              " Where A.ID = B.质控品id And B.项目id = C.ID And A.仪器id = [1] "
        
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "LisComm自动计算", lngDeviceID)
    Do Until rsTmp.EOF
        '计算一段时间
            gstrSql = "Select Zl_检验质控记录_Compute(" & lngDeviceID & ", " & rsTmp("项目ID") & ", To_Date('" & Format(date日期, "yyyy-mm-dd") & "','yyyy-mm-dd'), '" & str质控品 & "') From Dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "LisComm自动计算")

            If rsTemp.RecordCount <= 0 Then strReturn = strReturn & Format(date日期, "yyyy-mm-dd") & " " & Nvl(rsTmp("中文名")) & "(" & Nvl(rsTmp("英文名")) & ")  计算过程调用错误！" & vbCrLf
            If InStr(rsTemp.Fields(0).Value, "出现失控！") > 0 Then
                strReturn = strReturn & Format(date日期, "yyyy-mm-dd") & " " & Nvl(rsTmp("中文名")) & "(" & Nvl(rsTmp("英文名")) & ")" & rsTemp.Fields(0).Value & vbCrLf

            ElseIf InStr(rsTemp.Fields(0).Value, "计算完成！") <= 0 Then
                If InStr(rsTemp.Fields(0).Value, "按规则未发现警告和失控！") <= 0 Then
                strReturn = strReturn & Format(date日期, "yyyy-mm-dd") & " " & Nvl(rsTmp("中文名")) & "(" & Nvl(rsTmp("英文名")) & ")" & rsTemp.Fields(0).Value & vbCrLf
                End If
            End If
        rsTmp.MoveNext
    Loop
    If Trim(strReturn) <> "" Then
        MsgBox "数据已保存，质控计算时发现失控或警告！", vbInformation, "保存数据"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub ShowMe(objfrm As Object, lngMachine As Long)
    '打开窗体
    mlngMachine = lngMachine
    Me.Show , objfrm
End Sub
Private Sub InitRecordSet(rsNumber As adodb.Recordset)
    '初始化记录集(用于记录计算项目)
    Dim rsTmp As New adodb.Recordset
    
    Set rsNumber = New adodb.Recordset
    rsNumber.Fields.Append "诊治项目ID", adBigInt
    rsNumber.Fields.Append "阳性公式", adVarChar, 50
    rsNumber.Fields.Append "弱阳性公式", adVarChar, 50
    rsNumber.Fields.Append "CutOff公式", adVarChar, 50
    rsNumber.CursorLocation = adUseClient
    rsNumber.LockType = adLockOptimistic
    rsNumber.CursorType = adOpenStatic
    rsNumber.Open
    
    
    gstrSql = "select 诊治项目ID,阳性公式,弱阳性公式,CutOff公式 from 检验项目 where 项目类别 = 4"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
    Do Until rsTmp.EOF
        rsNumber.AddNew
        rsNumber("诊治项目ID") = rsTmp("诊治项目ID")
        rsNumber("阳性公式") = rsTmp("阳性公式") & ""
        rsNumber("弱阳性公式") = rsTmp("弱阳性公式") & ""
        rsNumber("CutOff公式") = rsTmp("CutOff公式") & ""
        rsNumber.Update
        rsTmp.MoveNext
    Loop
    
End Sub

Private Sub SelectBatch()
    '试剂批号选择功能
    Dim strReturn As String
    Dim lngItemID As Long
    Dim intLoop As Integer
    
    If Me.cbo测试项目.ListIndex > -1 Then
        lngItemID = Me.cbo测试项目.ItemData(Me.cbo测试项目.ListIndex)
    End If
    
    strReturn = GetBatchNo(Me.txt试剂批号, Me.txt试剂批号.Text, lngItemID)
    If UBound(Split(strReturn, "|")) = 4 Then
        lngItemID = Val(Split(strReturn, "|")(0))
        txt试剂批号.Tag = Split(strReturn, "|")(1) '免得被误操作改了
        txt试剂批号 = Split(strReturn, "|")(1)
        txt试剂效期 = Split(strReturn, "|")(2)
        txt试剂厂商 = Split(strReturn, "|")(3)
        txt测试方法 = Split(strReturn, "|")(4)
        
        If lngItemID = 0 Then
            If Me.cbo测试项目.ListIndex > -1 Then
                lngItemID = Me.cbo测试项目.ItemData(Me.cbo测试项目.ListIndex)
            End If
        Else
            For intLoop = 0 To Me.cbo测试项目.ListCount - 1
                If Val(Me.cbo测试项目.ItemData(intLoop)) = lngItemID Then
                    Me.cbo测试项目.ListIndex = intLoop
                    Exit For
                End If
            Next
        End If
        For intLoop = 1 To 8
            If Val(Me.vsList.TextMatrix(intLoop, 13)) = lngItemID Then
                mTestReagent(intLoop) = txt试剂批号.Tag & ";" & txt试剂效期 & ";" & txt试剂厂商 & ";" & txt测试方法
            End If
        Next
    Else
        txt试剂批号.Text = ""
        txt试剂批号.Tag = ""
        txt试剂效期.Text = ""
        txt试剂厂商.Text = ""
        txt测试方法.Text = ""
        If Me.cbo测试项目.ListIndex > -1 Then
            lngItemID = Val(Me.cbo测试项目.ItemData(Me.cbo测试项目.ListIndex))
        End If
        For intLoop = 1 To 8
            If Val(Me.vsList.TextMatrix(intLoop, 13)) = 0 Then
                mTestReagent(intLoop) = ""
            End If
        Next
    
    End If
End Sub
Private Function GetBatchNo(ByRef objTxt As TextBox, ByVal strInput As String, ByVal lngItemID As Long) As String
    '试剂批号选择器
    Dim rsTmp As adodb.Recordset, strsql As String
    Dim objPoint As POINTAPI
    Dim sglX As Single, sglY As Single
    Dim strKey As String '查找关键字
    On Error GoTo hErr
    
    strKey = DelInvalidChar(strInput) & "%"
    If lngItemID = 0 Then
    strsql = "Select Rownum As ID, a.* From (Select a.试剂批号, a.试剂效期, a.试剂厂商, a.测试方法, b.名称 As 测试项目, c.报告项目id" & vbNewLine & _
            "From 检验酶标试剂 A, 诊疗项目目录 B, 检验报告项目 C" & vbNewLine & _
            "Where a.测试项目id = b.Id(+) And a.测试项目id = c.诊疗项目id(+) And b.组合项目(+) = 0 And a.试剂效期 > Sysdate " & vbNewLine & _
            " And (A.试剂批号 Like [1] Or A.试剂厂商 Like [2] Or A.测试方法 Like [2] ) " & vbNewLine & _
            ") A"
    Else
    strsql = "Select Rownum As ID, a.* From (Select a.试剂批号, a.试剂效期, a.试剂厂商, a.测试方法, b.名称 As 测试项目, c.报告项目id" & vbNewLine & _
            "From 检验酶标试剂 A, 诊疗项目目录 B, 检验报告项目 C" & vbNewLine & _
            "Where a.测试项目id = b.Id And a.测试项目id = c.诊疗项目id And b.组合项目 = 0 And a.试剂效期 > Sysdate " & vbNewLine & _
            " And  C.报告项目ID = [3]  And (A. 试剂批号 Like [1] Or A.试剂厂商 Like [2] Or A.测试方法 Like [2] ) " & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.试剂批号, a.试剂效期, a.试剂厂商, a.测试方法, Null As 测试项目, Null As 报告项目id" & vbNewLine & _
            "From 检验酶标试剂 A" & vbNewLine & _
            "Where a.试剂效期 > Sysdate And 测试项目id Is Null" & vbNewLine & _
            "  And (A. 试剂批号 Like [1] Or A.试剂厂商 Like [2] Or A.测试方法 Like [2] ) " & vbNewLine & _
            "Order By 试剂效期 Desc) A"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, "取试剂批号", strKey, "%" & strKey, lngItemID)
    If rsTmp.EOF Then
        GetBatchNo = strInput
    Else
        If rsTmp.RecordCount = 1 Then
            GetBatchNo = "" & rsTmp!报告项目ID & "|" & rsTmp!试剂批号 & "|" & rsTmp!试剂效期 & "|" & rsTmp!试剂厂商 & "|" & rsTmp!测试方法
        Else
            Call ClientToScreen(objTxt.hWnd, objPoint)
            sglX = objPoint.x * 15 - 30
            sglY = objPoint.y * 15 + objTxt.Height
            If frmSelectList.ShowSelect(Me, rsTmp, "试剂批号,800,0,0;试剂效期,800,0,0;试剂厂商,1500,0,0;测试方法,2500,0,0;测试项目,5500,0,0", sglX, sglY, objTxt.Width, 2000, Me.Name & "\酶标试剂批号选择", "请选择试剂批号") Then
                GetBatchNo = "" & rsTmp!报告项目ID & "|" & rsTmp!试剂批号 & "|" & rsTmp!试剂效期 & "|" & rsTmp!试剂厂商 & "|" & rsTmp!测试方法
            Else
                GetBatchNo = strInput
            End If
        End If
    End If
    Exit Function
hErr:
    MsgBox Err.Description
End Function

