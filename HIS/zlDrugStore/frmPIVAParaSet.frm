VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPIVAParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "静脉输液配置中心参数设置"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11310
   Icon            =   "frmPIVAParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picPRI 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   3120
      ScaleHeight     =   2055
      ScaleWidth      =   2535
      TabIndex        =   5
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmdYes 
         Height          =   360
         Left            =   720
         Picture         =   "frmPIVAParaSet.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   810
      End
      Begin VB.CommandButton cmdNO 
         Height          =   360
         Left            =   1560
         Picture         =   "frmPIVAParaSet.frx":6DDC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1560
         Width           =   810
      End
      Begin MSComctlLib.ListView lvwPRI 
         Height          =   1305
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "双击或按回车键确认"
         Top             =   120
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2302
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgLvwSel"
         SmallIcons      =   "imgLvwSel"
         ColHdrIcons     =   "imgLvwSel"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   6135
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   9
      TabHeight       =   520
      OLEDropMode     =   1
      TabCaption(0)   =   "基础设置(&0)"
      TabPicture(0)   =   "frmPIVAParaSet.frx":6F26
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra输液单控制"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkOpen"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra显示控制"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "工作批次(&1)"
      TabPicture(1)   =   "frmPIVAParaSet.frx":6F42
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "vsfBatch"
      Tab(1).Control(2)=   "cmdDel"
      Tab(1).Control(3)=   "cmdAdd"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "打印控制(&2)"
      TabPicture(2)   =   "frmPIVAParaSet.frx":6F5E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "lblNum"
      Tab(2).Control(2)=   "lbl票据"
      Tab(2).Control(3)=   "cboNum"
      Tab(2).Control(4)=   "cmd打印设置"
      Tab(2).Control(5)=   "cbo票据设置"
      Tab(2).Control(6)=   "fra瓶签打印方式"
      Tab(2).Control(7)=   "fra其他报表打印方式"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "优先级设置(&3)"
      TabPicture(3)   =   "frmPIVAParaSet.frx":6F7A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblpritip"
      Tab(3).Control(1)=   "vsfPri"
      Tab(3).Control(2)=   "vsfDept"
      Tab(3).Control(3)=   "chkAll"
      Tab(3).Control(4)=   "cmdIN"
      Tab(3).Control(5)=   "cmdDelPri"
      Tab(3).Control(6)=   "cmdAddPri"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "容量设置(&4)"
      TabPicture(4)   =   "frmPIVAParaSet.frx":6F96
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblvoltip"
      Tab(4).Control(1)=   "vsfVolume"
      Tab(4).Control(2)=   "cmdVolAdd"
      Tab(4).Control(3)=   "cmdVolDel"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "常用药品设置(&5)"
      TabPicture(5)   =   "frmPIVAParaSet.frx":6FB2
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblMedi"
      Tab(5).Control(1)=   "vsfPrint"
      Tab(5).Control(2)=   "chkByMedi"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "不配置药品设置(&6)"
      TabPicture(6)   =   "frmPIVAParaSet.frx":6FCE
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblNoMedi"
      Tab(6).Control(1)=   "vsfNoMedi"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "显示来源病区(&7)"
      TabPicture(7)   =   "frmPIVAParaSet.frx":6FEA
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Lvw来源病区"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      Begin VB.Frame fra显示控制 
         Caption         =   " 显示控制 "
         Height          =   855
         Left            =   120
         TabIndex        =   69
         Top             =   3720
         Width           =   8775
         Begin VB.ComboBox cbo药品名称显示方式 
            Height          =   300
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   300
            Width           =   2295
         End
         Begin VB.Label lal药品名称显示方式 
            AutoSize        =   -1  'True
            Caption         =   "药品名称显示方式"
            Height          =   180
            Left            =   360
            TabIndex        =   70
            Top             =   360
            Width           =   1440
         End
      End
      Begin VB.CheckBox chkByMedi 
         Caption         =   "是否根据设置的常用药品进行药品过滤操作"
         Height          =   255
         Left            =   -74880
         TabIndex        =   65
         Top             =   360
         Width           =   3855
      End
      Begin VB.Frame fra其他报表打印方式 
         Caption         =   " 其他报表打印方式"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   55
         Top             =   2280
         Width           =   6255
         Begin VB.ComboBox cbo摆药单 
            Height          =   300
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   660
            Width           =   2415
         End
         Begin VB.ComboBox cbo发送单 
            Height          =   300
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1035
            Width           =   2415
         End
         Begin VB.ComboBox cboSum 
            Height          =   300
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   300
            Width           =   2415
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "发送确认后"
            Height          =   180
            Left            =   120
            TabIndex        =   64
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "摆药确认后"
            Height          =   180
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "摆药汇总清单"
            Height          =   180
            Left            =   3720
            TabIndex        =   62
            Top             =   720
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "汇总发送清单"
            Height          =   180
            Left            =   3720
            TabIndex        =   61
            Top             =   1095
            Width           =   1080
         End
         Begin VB.Label lblSum 
            AutoSize        =   -1  'True
            Caption         =   "汇总报表"
            Height          =   180
            Left            =   3720
            TabIndex        =   60
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblSumPrint 
            AutoSize        =   -1  'True
            Caption         =   "打印标签后"
            Height          =   180
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame fra瓶签打印方式 
         Caption         =   " 瓶签打印方式"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   49
         Top             =   480
         Width           =   6255
         Begin VB.CheckBox chkManPrint 
            Caption         =   "允许手工控制打印瓶签（可进行补打）"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   1080
            Width           =   3375
         End
         Begin VB.ComboBox cbo标签打印 
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox chkPrintLabelStep 
            Caption         =   "摆药后"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   52
            Top             =   300
            Width           =   855
         End
         Begin VB.CheckBox chkPrintLabelStep 
            Caption         =   "配药后"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   660
            Width           =   855
         End
         Begin VB.ComboBox cbo标签打印 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.ComboBox cbo票据设置 
         Height          =   300
         Left            =   -73500
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   4305
         Width           =   2415
      End
      Begin VB.CommandButton cmd打印设置 
         Caption         =   "打印设置(&P)"
         Height          =   345
         Left            =   -70980
         TabIndex        =   44
         Top             =   4275
         Width           =   1155
      End
      Begin VB.ComboBox cboNum 
         Height          =   300
         Left            =   -73500
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         Caption         =   " 卡片控制 "
         Height          =   615
         Left            =   120
         TabIndex        =   39
         Top             =   2880
         Width           =   8800
         Begin VB.ComboBox cbo数量 
            Height          =   300
            ItemData        =   "frmPIVAParaSet.frx":7006
            Left            =   1080
            List            =   "frmPIVAParaSet.frx":7013
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbl数量2 
            AutoSize        =   -1  'True
            Caption         =   "张卡片"
            Height          =   180
            Left            =   2040
            TabIndex        =   42
            Top             =   300
            Width           =   540
         End
         Begin VB.Label lbl数量1 
            Caption         =   "单行显示"
            Height          =   195
            Left            =   240
            TabIndex        =   41
            Top             =   300
            Width           =   735
         End
      End
      Begin VB.CheckBox chkOpen 
         Caption         =   "启用接收时间段控制"
         Height          =   180
         Left            =   360
         TabIndex        =   27
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Frame fra输液单控制 
         Height          =   1215
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   8800
         Begin VB.CheckBox chk当日医嘱 
            Caption         =   "接收当日及以前的医嘱"
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox txtDeff 
            Enabled         =   0   'False
            Height          =   270
            Left            =   3000
            TabIndex        =   30
            Text            =   "0"
            Top             =   795
            Width           =   375
         End
         Begin MSComCtl2.UpDown updDeff 
            Height          =   270
            Left            =   3480
            TabIndex        =   29
            Top             =   795
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   315
            Left            =   960
            TabIndex        =   32
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   84934658
            CurrentDate     =   36985
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   315
            Left            =   3240
            TabIndex        =   33
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   84934658
            CurrentDate     =   36985
         End
         Begin VB.Label lbl时间控制 
            AutoSize        =   -1  'True
            Caption         =   "医嘱发送不在该时间段输液医嘱将不再产生输液单。"
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   4560
            TabIndex        =   38
            Top             =   360
            Width           =   4140
         End
         Begin VB.Label lbl当日医嘱 
            AutoSize        =   -1  'True
            Caption         =   "勾选时配置中心将接收满足时间差条件的当日执行的医嘱。"
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   3840
            TabIndex        =   37
            Top             =   840
            Width           =   4680
         End
         Begin VB.Label lblBegin 
            Caption         =   "开始时间"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblEnd 
            Caption         =   "结束时间"
            Height          =   255
            Left            =   2400
            TabIndex        =   35
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblDeff 
            Caption         =   "小时差"
            Height          =   255
            Left            =   2355
            TabIndex        =   34
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdVolDel 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   -66600
         TabIndex        =   20
         Top             =   1560
         Width           =   1100
      End
      Begin VB.CommandButton cmdVolAdd 
         Caption         =   "新增(&A)"
         Height          =   350
         Left            =   -66600
         TabIndex        =   19
         Top             =   1080
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddPri 
         Caption         =   "新增(&A)"
         Height          =   350
         Left            =   -66600
         TabIndex        =   15
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CommandButton cmdDelPri 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   -66600
         TabIndex        =   14
         Top             =   2400
         Width           =   1100
      End
      Begin VB.CommandButton cmdIN 
         Caption         =   "保存(&S)"
         Height          =   350
         Left            =   -66600
         TabIndex        =   13
         Top             =   1800
         Width           =   1100
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "应用于所有科室的优先级规则"
         Height          =   250
         Left            =   -74880
         TabIndex        =   12
         Top             =   720
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.Frame Frame2 
         Caption         =   " 配置中心库房选择 "
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   8800
         Begin VB.CheckBox chkCheck 
            Caption         =   "审核该药房的所有医嘱"
            Height          =   255
            Left            =   4080
            TabIndex        =   67
            Top             =   240
            Width           =   3855
         End
         Begin VB.ComboBox CboStore 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Width           =   2280
         End
         Begin VB.Label lblStore 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "配置中心"
            Height          =   180
            Left            =   360
            TabIndex        =   11
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新增(&A)"
         Height          =   350
         Left            =   -66480
         TabIndex        =   4
         Top             =   960
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   -66480
         TabIndex        =   3
         Top             =   1560
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDept 
         Height          =   4905
         Left            =   -74880
         TabIndex        =   16
         Top             =   1080
         Width           =   2400
         _cx             =   4233
         _cy             =   8652
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":7020
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
      Begin VSFlex8Ctl.VSFlexGrid vsfPri 
         Height          =   4905
         Left            =   -72360
         TabIndex        =   17
         Top             =   1080
         Width           =   5505
         _cx             =   9710
         _cy             =   8652
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
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
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":70B6
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
      Begin VSFlex8Ctl.VSFlexGrid vsfVolume 
         Height          =   5145
         Left            =   -74880
         TabIndex        =   21
         Top             =   840
         Width           =   7995
         _cx             =   14102
         _cy             =   9075
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":716F
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
      Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
         Height          =   5025
         Left            =   -74880
         TabIndex        =   23
         Top             =   960
         Width           =   7995
         _cx             =   14102
         _cy             =   8864
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":7213
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
      Begin VSFlex8Ctl.VSFlexGrid vsfNoMedi 
         Height          =   5145
         Left            =   -74880
         TabIndex        =   25
         Top             =   840
         Width           =   8000
         _cx             =   14111
         _cy             =   9075
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":727C
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
      Begin VSFlex8Ctl.VSFlexGrid vsfBatch 
         Height          =   5025
         Left            =   -74880
         TabIndex        =   66
         Top             =   960
         Width           =   8160
         _cx             =   14393
         _cy             =   8864
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
         BackColorSel    =   16711680
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   9
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":72E5
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
      Begin MSComctlLib.ListView Lvw来源病区 
         Height          =   5445
         Left            =   -74880
         TabIndex        =   72
         Top             =   480
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   9604
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
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "设置配置中心工作批次(0批次作为特殊批次存在，不按给药时间范围划定)"
         Height          =   180
         Left            =   -74760
         TabIndex        =   68
         Top             =   600
         Width           =   5850
      End
      Begin VB.Label lbl票据 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "票据和报表"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74640
         TabIndex        =   48
         Top             =   4365
         Width           =   900
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         Caption         =   "瓶签打印份数"
         Height          =   180
         Left            =   -74640
         TabIndex        =   47
         Top             =   4860
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "仅限于摆药或配药后打印"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   -70980
         TabIndex        =   46
         Top             =   4860
         Width           =   1980
      End
      Begin VB.Label lblNoMedi 
         AutoSize        =   -1  'True
         Caption         =   "设置配置中心不进行配置的药品"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   -74880
         TabIndex        =   26
         Top             =   480
         Width           =   2520
      End
      Begin VB.Label lblMedi 
         AutoSize        =   -1  'True
         Caption         =   "设置常用药品，在输液单界面可以按药品进行过滤和排序"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   -74880
         TabIndex        =   24
         Top             =   720
         Width           =   4500
      End
      Begin VB.Label lblvoltip 
         AutoSize        =   -1  'True
         Caption         =   "设置某个科室单个病人某个批次可以配药的容量"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   3780
      End
      Begin VB.Label lblpritip 
         AutoSize        =   -1  'True
         Caption         =   "可以设置同个批次中同组药品的优先级"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   3060
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   10170
      TabIndex        =   1
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8760
      TabIndex        =   0
      Top             =   6360
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgLvwSel 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmPIVAParaSet.frx":7479
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAParaSet.frx":7793
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAParaSet.frx":7AAD
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAParaSet.frx":7DFF
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   5880
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPIVAParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrPrivs As String                              '权限串
Public mlng库房id As Long
Private mblnSetPara As Boolean
Private mRsDept As Recordset
Private mRsPC As Recordset
Private mRsType As Recordset
Private mintRow As Integer
Private mintCol As Integer
Private mblnPri As Boolean
Private mblnEdit As Boolean     '是否编辑优先级
Private mstrSourceDep As String '来源病区

Private Sub LoadStore()
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    
    On Error GoTo errHandle
    gstrSQL = "Select distinct B.id,B.名称 From 部门性质说明 A,部门表 B" & _
    " Where A.部门ID=B.ID And A.工作性质='配制中心' And B.Id In (Select 部门id From 部门人员 Where 人员id = [1])"
    
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取配置中心的部门", glngUserId)
    
    With Me.CboStore
        Do While Not rstemp.EOF
            .AddItem rstemp!名称
            .ItemData(.NewIndex) = rstemp!Id
            rstemp.MoveNext
        Loop
        If rstemp.RecordCount > 0 Then .ListIndex = 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load来源病区()
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "Select 编码 || '-' || 名称 科室, Id " & _
            " From 部门表 " & _
            " Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And Id In (Select 部门id From 部门性质说明 Where 工作性质 = '护理' And 服务对象 In (2,3)) And " & _
            " (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By 编码 || '-' || 名称 "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Load来源病区")
    
    '初始化病区
    With rsData
        If .EOF Then
            MsgBox "没有设置该类部门！（部门管理）", vbInformation, gstrSysName
            Exit Sub
        End If
        Lvw来源病区.ListItems.Clear
        Do While Not .EOF
            Lvw来源病区.ListItems.Add , "_" & !Id, !科室, 1, 1
            If mstrSourceDep <> "" Then
                If InStr("," & mstrSourceDep & ",", "," & CStr(!Id) & ",") > 0 Then
                    Lvw来源病区.ListItems("_" & !Id).Checked = True
                End If
            End If
            .MoveNext
        Loop
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadParams()
    Dim int摆药单 As Integer
    Dim int发药单 As Integer
    Dim int批次设置 As Integer
    Dim int上次批次 As Integer
    Dim int打包设置 As Integer
    Dim intBarCode As Integer
    Dim strAutoPrint As String
    Dim intManPrint As Integer
    Dim str截止时间 As String
    Dim int医嘱类型 As Integer
    Dim str输液给药途径 As String
    Dim str来源科室 As String
    Dim rsData As ADODB.Recordset
    Dim str当日医嘱 As String
    Dim intCount As Integer
    Dim intOpen As Integer
    Dim lng部门ID As Long
    Dim IntLocate As Integer
    Dim dateNow As Date
    Dim intNum As Integer
    Dim int配药后打包 As Integer
    Dim i As Integer
    Dim int汇总 As Integer
    Dim intTPN As Integer
    Dim intSpecial As Integer
    
    On Error GoTo errHandle
    '基础
    int摆药单 = Val(zlDatabase.GetPara("摆药后打印", glngSys, 1345, 0, Array(Label3, cbo摆药单, Label5), mblnSetPara))
    int发药单 = Val(zlDatabase.GetPara("发送后打印", glngSys, 1345, 0, Array(Label4, cbo发送单, Label6), mblnSetPara))
    
    strAutoPrint = zlDatabase.GetPara("瓶签自动打印", glngSys, 1345, "00|00", Array(chkPrintLabelStep(0), chkPrintLabelStep(1)), mblnSetPara)
    intManPrint = Val(zlDatabase.GetPara("瓶签手工打印", glngSys, 1345, "0", Array(chkManPrint), mblnSetPara))
    intCount = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\输液卡片", "卡片数量", 3))
    
    int汇总 = Val(zlDatabase.GetPara("打印标签后是否打印汇总报表", glngSys, 1345, 0, Array(lblSumPrint, cboSum, lblSum), mblnSetPara))
    
    Me.cbo药品名称显示方式.ListIndex = Val(zlDatabase.GetPara("药品名称显示方式", glngSys, 1345, 0, Array(lal药品名称显示方式, cbo药品名称显示方式), mblnSetPara))
    
    '辅助控制
    str截止时间 = zlDatabase.GetPara("工作截止时间", glngSys, 1345, "", Array(lblBegin, dtpBegin, lblEnd, dtpEnd), mblnSetPara)
    str当日医嘱 = zlDatabase.GetPara("不接收当日及以前医嘱", glngSys, 1345, 0, Array(chk当日医嘱, txtDeff, updDeff, lblDeff), mblnSetPara)
    
    intOpen = Val(zlDatabase.GetPara("启用接收时间控制", glngSys, 1345, 0, Array(chkOpen), mblnSetPara))
    lng部门ID = Val(zlDatabase.GetPara("配置中心", glngSys, 1345, 0, Array(CboStore, lblStore), mblnSetPara))
    intNum = Val(zlDatabase.GetPara("瓶签打印份数", glngSys, 1345, 1, Array(lblNum, cboNum), mblnSetPara))
    Me.chkByMedi.Value = Val(zlDatabase.GetPara("是否按设置的常用药品进行药品过滤操作", glngSys, 1345, 0, Array(chkByMedi), mblnSetPara))
    Me.chkCheck.Value = Val(zlDatabase.GetPara("审核该药房的所有数据", glngSys, 1345, 0, Array(chkCheck), mblnSetPara))

    '显示来源病区
    mstrSourceDep = zlDatabase.GetPara("显示来源病区", glngSys, 1345, "")

    If lng部门ID <> 0 Then                                  '定位药房
        '不存在该药房则提示
        For IntLocate = 0 To Me.CboStore.ListCount - 1
            If Me.CboStore.ItemData(IntLocate) = lng部门ID Then
                Me.CboStore.ListIndex = IntLocate
                Exit For
            End If
        Next
        If IntLocate > (CboStore.ListCount - 1) Then
            MsgBox "请重新设置配置中心（原来设置的配置中心已失效）！", vbInformation, gstrSysName
            If CboStore.ListCount >= 1 Then CboStore.ListIndex = 0
        End If
    Else
        MsgBox "请设置配置中心！", vbInformation, gstrSysName
    End If
    
    Me.chkOpen.Value = intOpen
    
    If InStr(1, str截止时间, "|") > 0 Then
        Me.dtpBegin.Value = Mid(str截止时间, 1, InStr(1, str截止时间, "|") - 1)
        Me.dtpEnd.Value = Mid(str截止时间, InStr(1, str截止时间, "|") + 1)
    End If
    
    Me.chk当日医嘱.Value = Mid(str当日医嘱, 1, 1)
    If InStr(1, str当日医嘱, "|") > 1 Then
        Me.txtDeff.Text = Mid(str当日医嘱, 3)
    Else
        Me.txtDeff.Text = 0
    End If
    
    ''基础设置
    If int摆药单 >= 0 And int摆药单 <= cbo摆药单.ListCount - 1 Then
        cbo摆药单.ListIndex = int摆药单
    End If
    
    If int汇总 >= 0 And int汇总 <= cboSum.ListCount - 1 Then
        cboSum.ListIndex = int汇总
    End If
    
    If int发药单 >= 0 And int发药单 <= cbo摆药单.ListCount - 1 Then
        cbo发送单.ListIndex = int发药单
    End If
    
    If InStr(1, strAutoPrint, "|") = 0 Or Len(strAutoPrint) <> 5 Then
        strAutoPrint = "00|00"
    End If
    
    If Mid(strAutoPrint, 1, 1) = 1 Then
        chkPrintLabelStep(0).Value = 1
        If Val(Mid(strAutoPrint, 2, 1)) = 1 Then
            cbo标签打印(0).ListIndex = 1
        Else
            cbo标签打印(0).ListIndex = 0
        End If
    End If
    
    If Mid(strAutoPrint, 4, 1) = 1 Then
        chkPrintLabelStep(1).Value = 1
        If Val(Mid(strAutoPrint, 5, 1)) = 1 Then
            cbo标签打印(1).ListIndex = 1
        Else
            cbo标签打印(1).ListIndex = 0
        End If
    End If
    
    cbo标签打印(0).Enabled = chkPrintLabelStep(0).Enabled And (chkPrintLabelStep(0).Value = 1)
    cbo标签打印(1).Enabled = chkPrintLabelStep(1).Enabled And (chkPrintLabelStep(1).Value = 1)
    
    vsfVolume.Enabled = mblnSetPara
    vsfPrint.Enabled = mblnSetPara
    vsfNoMedi.Enabled = mblnSetPara
    vsfPri.Enabled = mblnSetPara
    cmdAddPri.Enabled = mblnSetPara
    cmdIN.Enabled = mblnSetPara
    cmdDelPri.Enabled = mblnSetPara
    cmdVolAdd.Enabled = mblnSetPara
    cmdVolDel.Enabled = mblnSetPara
    
    
    If intManPrint < 0 Or intManPrint > 1 Then
        chkManPrint.Value = 0
    Else
        chkManPrint.Value = intManPrint
    End If
    
    If chkManPrint.Value = 1 Then
        cboSum.Enabled = True
    Else
        cboSum.Enabled = False
    End If
    
    '卡片张数
    Me.cbo数量.Text = IIf(intCount = 0, 3, intCount)
    
    Me.cboNum.Text = IIf(intNum = 0, 3, intNum)

    '常用药品打印设置
    gstrSQL = "select 药品id,名称 from 输液优先打印药品"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Load药品")
    
    Me.vsfPrint.rows = rsData.RecordCount + 2
    For i = 1 To rsData.RecordCount
        Me.vsfPrint.TextMatrix(i, vsfPrint.ColIndex("药品id")) = rsData!药品ID
        Me.vsfPrint.TextMatrix(i, vsfPrint.ColIndex("药品名称与编码")) = rsData!名称
       
       rsData.MoveNext
    Next
    
    
    '输液不配置药品
    gstrSQL = "select 药品id,名称 from 输液不配置药品"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Load药品")
    
    Me.vsfNoMedi.rows = rsData.RecordCount + 2
    For i = 1 To rsData.RecordCount
        Me.vsfNoMedi.TextMatrix(i, vsfNoMedi.ColIndex("药品id")) = rsData!药品ID
        Me.vsfNoMedi.TextMatrix(i, vsfNoMedi.ColIndex("药品名称与编码")) = rsData!名称
       
       rsData.MoveNext
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CboStore_Click()
    Call LoadBatchSet
    Call loadVolume
End Sub

Private Sub chkAll_Click()
    If mblnEdit Then
        If MsgBox("请保存设置的优先级，切换科室后所作的优先级设置将失效，是否切换？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
           If Me.chkAll.Value = 0 Then
                Me.vsfPri.Left = Me.vsfDept.Width + Me.vsfDept.Left + 100
                Me.vsfPri.Width = Me.vsfPri.Width - Me.vsfDept.Width - 100
                Me.vsfDept.Visible = True
                Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("科室id"))))
            Else
                Me.vsfPri.Width = Me.vsfPri.Width + Me.vsfDept.Width + 100
                Me.vsfPri.Left = Me.vsfDept.Left
                Me.vsfDept.Visible = False
                
                Call LoadVsfPRI(0)
            End If
            mblnEdit = False
            
        End If
    Else
        If Me.chkAll.Value = 0 Then
            Me.vsfPri.Left = Me.vsfDept.Width + Me.vsfDept.Left + 100
            Me.vsfPri.Width = Me.vsfPri.Width - Me.vsfDept.Width - 100
            Me.vsfDept.Visible = True
            Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("科室id"))))
        Else
            Me.vsfPri.Width = Me.vsfPri.Width + Me.vsfDept.Width + 100
            Me.vsfPri.Left = Me.vsfDept.Left
            Me.vsfDept.Visible = False
            
            Call LoadVsfPRI(0)
        End If
    End If
End Sub

Private Sub chkManPrint_Click()
    If chkManPrint.Value = 1 Then
        cboSum.Enabled = True
    Else
        cboSum.Enabled = False
    End If
End Sub

Private Sub chkOpen_Click()
    Me.dtpBegin.Enabled = (Me.chkOpen.Value = 1)
    Me.dtpEnd.Enabled = (Me.chkOpen.Value = 1)
    Me.chk当日医嘱.Enabled = (Me.chkOpen.Value = 1)
    Me.updDeff.Enabled = (Me.chkOpen.Value = 1)
End Sub

Private Sub chkPrintLabelStep_Click(index As Integer)
    cbo标签打印(index).Enabled = (chkPrintLabelStep(index).Value = 1)
End Sub

Private Sub cmdAdd_Click()
    With vsfBatch
        If .rows > 2 Then
            If Trim(.TextMatrix(.rows - 1, .ColIndex("配置时间开始"))) = "" Or _
                Trim(.TextMatrix(.rows - 1, .ColIndex("配置时间结束"))) = "" Or _
                Trim(.TextMatrix(.rows - 1, .ColIndex("给药时间开始"))) = "" Or _
                Trim(.TextMatrix(.rows - 1, .ColIndex("给药时间结束"))) = "" Then
                Exit Sub
            End If
        End If
        
        .rows = .rows + 1
        
        If .rows > 3 Then
            .TextMatrix(.rows - 1, .ColIndex("批次")) = Mid(.TextMatrix(.rows - 2, .ColIndex("批次")), 1, Len(.TextMatrix(.rows - 2, .ColIndex("批次"))) - 1) + 1 & "#"
        Else
            .TextMatrix(.rows - 1, .ColIndex("批次")) = "0#"
        End If
        .TextMatrix(.rows - 1, .ColIndex("启用")) = "√"
    End With
End Sub

Private Sub cmdAddPri_Click()
    If Me.vsfPri.TextMatrix(Me.vsfPri.rows - 1, Me.vsfPri.ColIndex("配药类型")) <> "" And Me.vsfPri.TextMatrix(Me.vsfPri.rows - 1, Me.vsfPri.ColIndex("频次")) <> "" Then
        Me.vsfPri.rows = Me.vsfPri.rows + 1
        Me.vsfPri.RowHeight(Me.vsfPri.rows - 1) = 250
        Me.vsfPri.TextMatrix(Me.vsfPri.rows - 1, vsfPri.ColIndex("序号")) = Me.vsfPri.rows - 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    Dim lngRow As Long
    Dim lngCur As Long
    
    With vsfBatch
        If .Row > 1 Then
            If MsgBox("是否删除批次(" & .TextMatrix(.Row, .ColIndex("批次")) & ")？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            .Redraw = flexRDNone
            
            lngCur = .Row
            .RemoveItem .Row
            
            '重设批次号
            For lngRow = lngCur To .rows - 1
                .TextMatrix(lngRow, .ColIndex("批次")) = Mid(.TextMatrix(lngRow, .ColIndex("批次")), 1, Len(.TextMatrix(lngRow, .ColIndex("批次"))) - 1) - 1 & "#"
            Next
            
            .Redraw = flexRDDirect
        End If
    End With
End Sub

Private Sub cmdDelPri_Click()
    Dim i As Integer
    Dim intRow As Integer
    
    If Me.vsfPri.Row = 0 Then Exit Sub
    intRow = Me.vsfPri.Row
    Me.vsfPri.RemoveItem Me.vsfPri.Row
    
    '调整序号
    For i = intRow To Me.vsfPri.rows - 1
        Me.vsfPri.TextMatrix(i, Me.vsfPri.ColIndex("序号")) = i
    Next
    
    mblnEdit = True
End Sub

Private Sub cmdIN_Click()
    Dim intCount As Integer
    Dim lngRow As Long
    
    If mblnSetPara Then
         '保存优先级设置
        With vsfPri
            intCount = 1
            
            If .rows = 1 Then
                gstrSQL = "Zl_输液药品优先级_Save("
                '科室id
                gstrSQL = gstrSQL & "'" & IIf(chkAll.Value = 1, 0, vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("科室id"))) & "'"
                gstrSQL = gstrSQL & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "保存优先级")
            End If
            
            For lngRow = 1 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("配药类型")) <> "" And .TextMatrix(lngRow, .ColIndex("频次")) <> "" Then
                    
                    gstrSQL = "Zl_输液药品优先级_Save("
                    '科室id
                    gstrSQL = gstrSQL & "'" & IIf(chkAll.Value = 1, 0, vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("科室id"))) & "',"
                    '科室名称
                    gstrSQL = gstrSQL & "'" & IIf(chkAll.Value = 1, "所有科室", vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("科室名称"))) & "',"
                    '配药类型
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("配药类型")) & "',"
                    '频次
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("频次")) & "',"
                    '有效
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("有效"))) & ","
                    '优先级
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("序号")))
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存优先级")
                    intCount = intCount + 1
                End If
            Next
        End With
    End If
    
    mblnEdit = False
End Sub

Private Sub cmdNo_Click()
    picPRI.Visible = False
    CmdOK.Enabled = True
    CmdCancel.Enabled = True
End Sub

Private Sub cmdOk_Click()
    Dim strInput As String
    Dim lngRow As Long
    Dim strPrintLabel As String
    Dim intCount As Integer
    Dim i As Integer
    
    On Error GoTo errHandle
    
    '瓶签打印方式
    If chkPrintLabelStep(0).Value = 0 Then
        strPrintLabel = "00"
    Else
        strPrintLabel = "1" & cbo标签打印(0).ListIndex
    End If
    strPrintLabel = strPrintLabel & "|"
    If chkPrintLabelStep(1).Value = 0 Then
        strPrintLabel = strPrintLabel & "00"
    Else
        strPrintLabel = strPrintLabel & "1" & cbo标签打印(1).ListIndex
    End If

    '显示来源病区
    mstrSourceDep = ""
    With Me.Lvw来源病区
        For i = 1 To .ListItems.count
            If .ListItems(i).Checked Then
                If mstrSourceDep = "" Then
                    mstrSourceDep = Mid(.ListItems(i).Key, 2)
                Else
                    mstrSourceDep = mstrSourceDep & "," & Mid(.ListItems(i).Key, 2)
                End If
            End If
        Next
    End With

    '保存私有参数
    '基础设置
    zlDatabase.SetPara "摆药后打印", cbo摆药单.ListIndex, glngSys, 1345
    zlDatabase.SetPara "发送后打印", cbo发送单.ListIndex, glngSys, 1345
    zlDatabase.SetPara "瓶签自动打印", strPrintLabel, glngSys, 1345
    zlDatabase.SetPara "瓶签手工打印", chkManPrint.Value, glngSys, 1345
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\输液卡片", "卡片数量", Me.cbo数量.Text
    zlDatabase.SetPara "打印标签后是否打印汇总报表", cboSum.ListIndex, glngSys, 1345
    zlDatabase.SetPara "药品名称显示方式", cbo药品名称显示方式.ListIndex, glngSys, 1345
    
    '辅助控制
    zlDatabase.SetPara "工作截止时间", Format(dtpBegin.Value, "hh:mm:ss") & "|" & Format(Me.dtpEnd.Value, "hh:mm:ss"), glngSys, 1345
    zlDatabase.SetPara "不接收当日及以前医嘱", chk当日医嘱.Value & "|" & Me.txtDeff.Text, glngSys, 1345
    zlDatabase.SetPara "启用接收时间控制", chkOpen.Value, glngSys, 1345
    zlDatabase.SetPara "配置中心", Me.CboStore.ItemData(Me.CboStore.ListIndex), glngSys, 1345
    zlDatabase.SetPara "瓶签打印份数", Me.cboNum.Text, glngSys, 1345
    zlDatabase.SetPara "是否按设置的常用药品进行药品过滤操作", chkByMedi.Value, glngSys, 1345
    zlDatabase.SetPara "审核该药房的所有数据", chkCheck.Value, glngSys, 1345
    
    If zlStr.IsHavePrivs(mstrPrivs, "设置工作批次") Then
        With vsfBatch
            For lngRow = 2 To .rows - 1
                If IsDate(.TextMatrix(lngRow, .ColIndex("配置时间开始"))) And _
                    IsDate(.TextMatrix(lngRow, .ColIndex("配置时间结束"))) And _
                    IsDate(.TextMatrix(lngRow, .ColIndex("给药时间开始"))) And _
                    IsDate(.TextMatrix(lngRow, .ColIndex("给药时间结束"))) Then
                    
                    strInput = IIf(strInput = "", "", strInput & "|") & _
                        Mid(.TextMatrix(lngRow, .ColIndex("批次")), 1, Len(.TextMatrix(lngRow, .ColIndex("批次"))) - 1) & "," & _
                        .TextMatrix(lngRow, .ColIndex("配置时间开始")) & "-" & .TextMatrix(lngRow, .ColIndex("配置时间结束")) & "," & _
                        .TextMatrix(lngRow, .ColIndex("给药时间开始")) & "-" & .TextMatrix(lngRow, .ColIndex("给药时间结束")) & "," & _
                        IIf(.TextMatrix(lngRow, .ColIndex("打包")) = "", 0, 1) & "," & _
                        IIf(.TextMatrix(lngRow, .ColIndex("启用")) = "", 0, 1) & "," & _
                        .Cell(flexcpBackColor, lngRow, .ColIndex("颜色")) & "," & _
                        IIf(Trim(.TextMatrix(lngRow, .ColIndex("药品类型"))) = "", Null, .TextMatrix(lngRow, .ColIndex("药品类型")))
                End If
            Next
        End With
        
        '如果strInput为空表示删除整个工作批次
        gstrSQL = "Zl_配药工作批次_Save("
        '批次信息
        gstrSQL = gstrSQL & "'" & strInput & "',"
        gstrSQL = gstrSQL & Me.CboStore.ItemData(Me.CboStore.ListIndex)
        gstrSQL = gstrSQL & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存配药工作批次")
    End If
    
    '显示来源病区
    zlDatabase.SetPara "显示来源病区", mstrSourceDep, glngSys, 1345

    If mblnSetPara Then
        '保存容量设置
        With Me.vsfVolume
            For lngRow = 0 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("科室名称")) <> "" And .TextMatrix(lngRow, .ColIndex("容量")) <> "" Then
                    
                    gstrSQL = "Zl_科室容量设置_Save("
                    '科室id
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("科室id")) & "',"
                    '科室名称
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("科室名称")) & "',"
                    '批次
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("配药批次")) & "',"
                    '容量
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("容量"))) & ","
                    '优先级
                    gstrSQL = gstrSQL & lngRow & ","
                    '配置中心ID
                    gstrSQL = gstrSQL & Me.CboStore.ItemData(Me.CboStore.ListIndex)
                    gstrSQL = gstrSQL & ")"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存容量")
                End If
            Next
        End With
        
        '保存常用药品
        With Me.vsfPrint
            For i = 1 To .rows - 1
                If (.TextMatrix(i, .ColIndex("药品id")) <> "" And .TextMatrix(i, .ColIndex("药品名称与编码")) <> "") Or i = 1 Then
                    gstrSQL = "Zl_输液优先打印药品_打印设置("
                    '药品id
                    gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("药品id"))) & ","
                    '药品名称
                    gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("药品名称与编码")) & "',"
                    gstrSQL = gstrSQL & i & ")"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存常用药品")
                End If
            Next
        End With
        
        '保存不接受药品
        With Me.vsfNoMedi
            For i = 1 To .rows - 1
                If (.TextMatrix(i, .ColIndex("药品id")) <> "" And .TextMatrix(i, .ColIndex("药品名称与编码")) <> "") Or i = 1 Then
                    gstrSQL = "Zl_输液不配置药品_设置("
                    '药品id
                    gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("药品id"))) & ","
                    '药品名称
                    gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("药品名称与编码")) & "',"
                    gstrSQL = gstrSQL & i & ")"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存不接受药品")
                End If
            Next
        End With
    End If
    
    frmPIVAMain.mblnParamsRefresh = True
    
    Unload Me
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadVsfPRI(ByVal str科室id As String)
    Dim rstemp As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select 科室id,科室名称,配药类型,频次,有效,优先级 from 输液药品优先级 where (科室id=[1] or 科室id='0') order by 优先级"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取优先级数据", str科室id)
    
    i = 1
    rstemp.Filter = "科室id='" & str科室id & "'"
    If rstemp.EOF Then rstemp.Filter = ""
    With Me.vsfPri
        .RowHeight(0) = 250
        
        If rstemp.RecordCount = 0 Then
            .rows = 1
            .rows = 2
            .TextMatrix(1, .ColIndex("序号")) = 1
        Else
            .rows = rstemp.RecordCount + 1
        End If
       
        Do While Not rstemp.EOF
            .RowHeight(i) = 250
            .TextMatrix(i, .ColIndex("序号")) = rstemp!优先级
            .TextMatrix(i, .ColIndex("配药类型")) = rstemp!配药类型
            .TextMatrix(i, .ColIndex("频次")) = rstemp!频次
            .TextMatrix(i, .ColIndex("有效")) = rstemp!有效
            i = i + 1
            rstemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdVolAdd_Click()
    If Me.vsfVolume.TextMatrix(Me.vsfVolume.rows - 1, Me.vsfVolume.ColIndex("科室名称")) <> "" And Me.vsfVolume.TextMatrix(Me.vsfVolume.rows - 1, Me.vsfVolume.ColIndex("容量")) <> "" Then
        Me.vsfVolume.rows = Me.vsfVolume.rows + 1
        Me.vsfVolume.RowHeight(Me.vsfVolume.rows - 1) = 250
    End If
End Sub

Private Sub cmdVolDel_Click()
    If Me.vsfVolume.Row = 0 Then Exit Sub
    Me.vsfVolume.RemoveItem Me.vsfVolume.Row
End Sub

Private Sub cmdYes_Click()
    Dim strIDS As String
    Dim strReturn As String
    
    strReturn = ReturnSelectedPri(1, strIDS)
    
    If mblnPri Then
        With Me.vsfPri
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("科室名称") Then
                .TextMatrix(mintRow, .ColIndex("科室id")) = strIDS
            End If
        End With
    Else
        With Me.vsfVolume
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("科室名称") Then
                .TextMatrix(mintRow, .ColIndex("科室id")) = strIDS
            End If
        End With
    End If
    
End Sub

Private Sub cmd打印设置_Click()
    Dim strBill As String
    Select Case cbo票据设置.ListIndex
    Case 0
        '输液瓶标签
        strBill = "ZL1_BILL_1345_1"
    Case 1
        '摆药药品汇总清单
        strBill = "ZL1_INSIDE_1345_1"
    Case 2
        '发送药品汇总清单
        strBill = "ZL1_INSIDE_1345_2"
    Case 3
        '退药销帐清单
        strBill = "ZL1_BILL_1345_2"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_Load()
    mblnSetPara = zlStr.IsHavePrivs(mstrPrivs, "参数设置")
    
    With cbo标签打印(0)
        .Clear
        .AddItem "0-提示是否打印"
        .AddItem "1-自动打印"
        .ListIndex = 0
    End With
    
    With cbo标签打印(1)
        .Clear
        .AddItem "0-提示是否打印"
        .AddItem "1-自动打印"
        .ListIndex = 0
    End With
    
    With cbo摆药单
        .Clear
        .AddItem "0-提示是否打印"
        .AddItem "1-自动打印"
        .AddItem "2-不打印"
    End With
    
    With cbo发送单
        .Clear
        .AddItem "0-提示是否打印"
        .AddItem "1-自动打印"
        .AddItem "2-不打印"
    End With
    
    With cboSum
        .Clear
        .AddItem "0-提示是否打印"
        .AddItem "1-自动打印"
        .AddItem "2-不打印"
    End With
    
    With cbo票据设置
        .Clear
        .AddItem "1-输液瓶标签"
        .AddItem "2-摆药药品汇总清单"
        .AddItem "3-发送药品汇总清单"
        .AddItem "4-退药销帐清单"

        .ListIndex = 0
    End With
    
    With cboNum
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        
        .ListIndex = 0
    End With
        
    With vsfBatch
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(.ColIndex("批次")) = True
        .MergeCol(.ColIndex("颜色")) = True
        .MergeCol(.ColIndex("配置时间开始")) = True
        .MergeCol(.ColIndex("配置时间结束")) = True
        .MergeCol(.ColIndex("给药时间开始")) = True
        .MergeCol(.ColIndex("给药时间结束")) = True
        .MergeCol(.ColIndex("打包")) = True
        .MergeCol(.ColIndex("启用")) = True
        .MergeCol(.ColIndex("药品类型")) = True
        .MergeCells = flexMergeFixedOnly
    End With
    
    With cbo药品名称显示方式
        .Clear
        .AddItem "药品编码+药品名称", 0
        .AddItem "药品名称", 1
        .AddItem "药品编码", 2
    End With
    
    Call LoadStore
        
    '提取参数
    Call LoadBatchSet
    Call LoadParams
    Call LoadPRI
    
    Call loadVolume
    Call LoadDept
    Call Load来源病区
    
    Call chkAll_Click
    
    Call chkOpen_Click
End Sub
Private Sub LoadBatchSet()
    '提取配药中心工作批次
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 批次,颜色, 配药时间, 给药时间, 打包, 启用,药品类型 From 配药工作批次 where 配置中心ID=[1] Order By 批次"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取配药中心工作批次", Me.CboStore.ItemData(Me.CboStore.ListIndex))
    
    With vsfBatch
        .rows = 2
        .ColComboList(.ColIndex("药品类型")) = "   |肿瘤药|营养药|抗生素"
        Do While Not rsTmp.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("批次")) = rsTmp!批次 & "#"
            .TextMatrix(.rows - 1, .ColIndex("配置时间开始")) = Mid(rsTmp!配药时间, 1, InStr(rsTmp!配药时间, "-") - 1)
            .TextMatrix(.rows - 1, .ColIndex("配置时间结束")) = Mid(rsTmp!配药时间, InStr(rsTmp!配药时间, "-") + 1)
            .TextMatrix(.rows - 1, .ColIndex("给药时间开始")) = Mid(rsTmp!给药时间, 1, InStr(rsTmp!给药时间, "-") - 1)
            .TextMatrix(.rows - 1, .ColIndex("给药时间结束")) = Mid(rsTmp!给药时间, InStr(rsTmp!给药时间, "-") + 1)
            .TextMatrix(.rows - 1, .ColIndex("打包")) = IIf(rsTmp!打包 = 0, "", "√")
            .TextMatrix(.rows - 1, .ColIndex("启用")) = IIf(rsTmp!启用 = 0, IIf(rsTmp!批次 = 0, "√", ""), "√")
            .TextMatrix(.rows - 1, .ColIndex("药品类型")) = NVL(rsTmp!药品类型)
            
            If .TextMatrix(.rows - 1, .ColIndex("启用")) = "" Then
                .Cell(flexcpBackColor, .rows - 1, 0, .rows - 1, .Cols - 1) = &HE0E0E0
            Else
                .Cell(flexcpBackColor, .rows - 1, 0, .rows - 1, .Cols - 1) = &H80000005
            End If
            
            .Cell(flexcpBackColor, .rows - 1, .ColIndex("颜色"), .rows - 1, .ColIndex("颜色")) = IIf(rsTmp!批次 = 0, &H80000005, rsTmp!颜色)
            rsTmp.MoveNext
        Loop
        
        vsfBatch.Enabled = IsHavePrivs(mstrPrivs, "设置工作批次")
        If vsfBatch.Enabled = False Then
            Label2.Caption = Label2.Caption & "(无权限进行修改)"
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnEdit = False
End Sub

Private Sub lvwPRI_DblClick()
    Dim strIDS As String
    Dim strReturn As String
    
    strReturn = ReturnSelectedPri(0, strIDS)
    
    If mblnPri Then
        With Me.vsfPri
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("科室名称") Then
                .TextMatrix(mintRow, .ColIndex("科室id")) = strIDS
            End If
        End With
    Else
        With Me.vsfVolume
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("科室名称") Then
                .TextMatrix(mintRow, .ColIndex("科室id")) = strIDS
            End If
        End With
    End If
End Sub

Private Sub lvwPRI_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With lvwPRI
        For n = 1 To .ListItems.count
            .ListItems(n).Selected = False
        Next
        
        Item.Selected = True
        If Mid(Item.Text, 1, 2) = "所有" Then
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

Private Sub lvwPRI_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strIDS As String
    Dim strReturn As String

    If KeyCode = vbKeyReturn Then
        strReturn = ReturnSelectedPri(1, strIDS)
        
        If mblnPri Then
            With Me.vsfPri
                .TextMatrix(mintRow, mintCol) = strReturn
                If mintCol = .ColIndex("科室名称") Then
                    .TextMatrix(mintRow, .ColIndex("科室id")) = strIDS
                End If
            End With
        Else
            With Me.vsfVolume
                .TextMatrix(mintRow, mintCol) = strReturn
                If mintCol = .ColIndex("科室名称") Then
                    .TextMatrix(mintRow, .ColIndex("科室id")) = strIDS
                End If
            End With
        End If
    End If
End Sub






Private Function ReturnSelectedPri(ByVal intType As Integer, ByRef strIDS As String) As String
    'intType:0-双击列表时；1-列表中按回车时
    Dim n As Integer
    Dim strReturn As String
    
    With lvwPRI
        If .SelectedItem Is Nothing Then Exit Function
        
        strReturn = .SelectedItem.Text
        strIDS = Mid(.SelectedItem.Key, 2)
        
'        '如果选择了全选，则不用取所有选项了
'        If .ListItems(1).Checked Then
'            strReturn = .ListItems(1).Text
'            ReturnSelectedPri = strReturn
'            picPRI.Visible = False
'            Exit Function
'        End If
'
'        For n = 1 To .ListItems.Count
'            If .ListItems(n).Checked Then
'                strReturn = IIf(strReturn = "", .ListItems(n).Text, strReturn & "," & .ListItems(n).Text)
'                strIDS = IIf(strIDS = "", Mid(.ListItems(n).Key, 2), strIDS & "," & Mid(.ListItems(n).Key, 2))
'            End If
'        Next
'
'        If intType = 0 Then
'            '如果当前双击的选项未被选上，将当前双击的选项也加入到编辑框中
'            If .SelectedItem.Checked = False Then
'                .SelectedItem.Checked = True
'                strReturn = IIf(strReturn = "", .SelectedItem.Text, strReturn & "," & .SelectedItem.Text)
'                strIDS = IIf(strIDS = "", Mid(.ListItems(n).Key, 2), strIDS & "," & Mid(.ListItems(n).Key, 2))
'            End If
'
'            If .ListItems(1).Checked Then
'                strReturn = .ListItems(1).Text
'                ReturnSelectedPri = strReturn
'                Exit Function
'            End If
'        End If
        
        picPRI.Visible = False
        
        CmdOK.Enabled = True
        CmdCancel.Enabled = True
        ReturnSelectedPri = strReturn
        mblnEdit = True
    End With
End Function

Private Sub picPRI_Resize()
    On Error Resume Next
    
    With lvwPRI
        .Top = 0
        .Left = 0
        .Width = picPRI.Width
        .Height = picPRI.Height - 200 - cmdNO.Height
    End With
    
    With cmdNO
        .Top = picPRI.Height - .Height - 50
        .Left = picPRI.Width - .Width - 50
    End With
    
    With cmdYes
        .Top = cmdNO.Top
        .Left = cmdNO.Left - .Width - 100
    End With
End Sub



Private Sub sstMain_Click(PreviousTab As Integer)
    Dim i As Integer
    
    If PreviousTab = 5 Then
        Me.vsfVolume.Row = Me.vsfVolume.rows - 1
        Me.vsfVolume.Col = Me.vsfVolume.ColIndex("科室名称")
    End If
End Sub







Private Sub LoadPRI()

    Set mRsDept = DeptSendWork_Get科室名称
    
    Set mRsType = DeptSendWork_Get配药类型
    
    Set mRsPC = DeptSendWork_Get频次
    
End Sub


Private Sub updDeff_DownClick()
    If Me.txtDeff.Text <> "0" Then
        Me.txtDeff.Text = Me.txtDeff.Text - 1
    End If
End Sub

Private Sub updDeff_UpClick()
    If Me.txtDeff.Text <> "24" Then
        Me.txtDeff.Text = Me.txtDeff.Text + 1
    End If
End Sub

Private Sub vsfBatch_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfBatch
        Select Case Col
            Case .ColIndex("配置时间开始"), .ColIndex("配置时间结束"), .ColIndex("给药时间开始"), .ColIndex("给药时间结束")
                If .TextMatrix(Row, Col) = "" Then Exit Sub
                
                If IsDate(.TextMatrix(Row, Col)) = False Then
                    MsgBox "请录入时间格式，比如12:59或者9:20等。", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = ""
                    Exit Sub
                End If
                
                If Col = .ColIndex("配置时间开始") And .TextMatrix(Row, .ColIndex("配置时间结束")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("配置时间开始"))) >= CDate(.TextMatrix(Row, .ColIndex("配置时间结束"))) Then
                        MsgBox "开始时间必须小于结束时间，请重新设置。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
                
                If Col = .ColIndex("配置时间结束") And .TextMatrix(Row, .ColIndex("配置时间开始")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("配置时间结束"))) <= CDate(.TextMatrix(Row, .ColIndex("配置时间开始"))) Then
                        MsgBox "结束时间必须大于开始时间，请重新设置。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
                
                If Col = .ColIndex("给药时间开始") And .TextMatrix(Row, .ColIndex("给药时间结束")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("给药时间开始"))) >= CDate(.TextMatrix(Row, .ColIndex("给药时间结束"))) Then
                        MsgBox "开始时间必须小于结束时间，请重新设置。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
                
                If Col = .ColIndex("给药时间结束") And .TextMatrix(Row, .ColIndex("给药时间开始")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("给药时间结束"))) <= CDate(.TextMatrix(Row, .ColIndex("给药时间开始"))) Then
                        MsgBox "结束时间必须大于开始时间，请重新设置。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfBatch_DblClick()
    With vsfBatch
        If .Row < 2 Then Exit Sub
        If (.Col <> .ColIndex("打包") And .Col <> .ColIndex("启用")) And .Col <> .ColIndex("颜色") Then Exit Sub
        If (.MouseRow <> .Row Or .MouseCol <> .Col) And .Col <> .ColIndex("颜色") Then Exit Sub
        
        If .Col <> .ColIndex("颜色") Then
            If .TextMatrix(.Row, .Col) = "√" Then
                If .TextMatrix(.Row, .ColIndex("批次")) = "0#" And .Col = .ColIndex("启用") Then
                    MsgBox "0批次作为特殊批次，无法设置为【不启用】状态！"
                Else
                    .TextMatrix(.Row, .Col) = ""
                End If
            Else
                .TextMatrix(.Row, .Col) = "√"
            End If
            
            If .Col = .ColIndex("启用") Then
                If .TextMatrix(.Row, .Col) = "" Then
                    .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = &HE0E0E0
                Else
                    .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = &H80000005
                End If
            End If
        
        Else
            On Error GoTo errHandle
            cmdialog.CancelError = True
            cmdialog.ShowColor
            .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = cmdialog.Color
            
errHandle:
        End If
    End With
End Sub


Private Sub vsfBatch_EnterCell()
    With vsfBatch
        If .Row < 2 Then Exit Sub
        .Editable = flexEDNone
        
        If .Col = .ColIndex("配置时间开始") Or .Col = .ColIndex("配置时间结束") Or .Col = .ColIndex("给药时间开始") Or .Col = .ColIndex("给药时间结束") Or .Col = .ColIndex("药品类型") Then
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub


Private Sub vsfBatch_KeyPress(KeyAscii As Integer)
    With vsfBatch
        If KeyAscii = 13 Then
            If .Col < .Cols - 1 Then
                .Col = .Col + 1
            Else
                If .Row < .rows - 1 Then
                    .Row = .Row + 1
                    .Col = .ColIndex("配置时间开始")
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfBatch_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    With vsfBatch
        Select Case Col
            Case .ColIndex("配置时间开始"), .ColIndex("配置时间结束"), .ColIndex("给药时间开始"), .ColIndex("给药时间结束")
                If InStr("1234567890:" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                ElseIf KeyAscii = Asc(":") Then
                    If InStr(.EditText, ":") <> 0 Then
                        KeyAscii = 0
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfDept_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row <> 1 Then Cancel = True
End Sub

Private Sub vsfDept_EnterCell()
    If mblnEdit Then
        If MsgBox("请保存设置的优先级，切换科室后所作的优先级设置将失效，是否切换？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("科室id"))))
            mblnEdit = False
            
        End If
    Else
        If Me.vsfDept.Row > 1 Then
            Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("科室id"))))
        End If
    End If
    
End Sub

Private Sub vsfDept_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    
    With Me.vsfDept
        If KeyAscii <> 13 Or .TextMatrix(1, .ColIndex("科室名称")) = "" Or .Row <> 1 Then Exit Sub
        
        For intRow = 2 To .rows - 1
            If .TextMatrix(intRow, .ColIndex("简码")) = UCase(.TextMatrix(1, .ColIndex("科室名称"))) Then
                .Row = intRow
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub vsfNoMedi_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim i As Integer
    Dim strKey As String
    Dim StrCode As String
    
    If KeyCode = 13 Then
        vRect = zlControl.GetControlRect(vsfNoMedi.hWnd)
        dblLeft = vRect.Left + vsfNoMedi.CellLeft
        dblTop = vRect.Top + vsfNoMedi.CellTop + vsfNoMedi.CellHeight + 3200
        
        With vsfNoMedi
            If Col = .ColIndex("药品名称与编码") Then
                strKey = Trim(.EditText)
                If strKey = "" Then Exit Sub
                
                If IsNumeric(strKey) Then
                    '纯数字
                    StrCode = " d.编码 like [1] "
                ElseIf zlCommFun.IsCharAlpha(strKey) Then
                    '纯字母
                    StrCode = " n.简码 Like [1] "
                ElseIf zlCommFun.IsCharChinese(strKey) Then
                    '纯汉字
                    StrCode = " d.名称 like [1] "
                Else
                    StrCode = " (n.简码 Like [1] Or d.编码 Like [1] Or n.名称 Like [1]) "
                End If
                                
                gstrSQL = "Select Distinct d.Id ,'【' || d.编码 || '】' || d.名称 || '(' || d.规格 || ')' As 通用名" & vbNewLine & _
                    " From 药品规格 T, 收费项目目录 D, 收费项目别名 N" & vbNewLine & _
                    " Where t.药品id = d.Id And t.药品id = n.收费细目id And D.类别 In ('5', '6') And" & StrCode & vbNewLine & _
                    " And (d.撤档时间 Is Null Or To_Char(d.撤档时间, 'yyyy-MM-dd') = '3000-01-01')" & vbNewLine & _
                    " Order By '【' || d.编码 || '】' || d.名称 || '(' || d.规格 || ')'"
                Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "药品名称与编码", False, "", "", False, False, _
                True, dblLeft, dblTop, .Height, blnCancel, False, True, IIf(gstrMatchMethod = 0, "", "%") & UCase(.EditText) & "%")
    
                If rsRecord Is Nothing Then
                    .EditText = ""
                    Exit Sub
                Else
                    For i = 1 To .rows - 1
                        If rsRecord!Id = Val(.TextMatrix(i, .ColIndex("药品ID"))) Then
                            MsgBox rsRecord!通用名 & "已经录入，请重新选择！", vbInformation + vbOKOnly, gstrSysName
                            .EditText = ""
                            Exit Sub
                        End If
                    Next
                    
                    .TextMatrix(.Row, .ColIndex("药品ID")) = rsRecord!Id
                    .TextMatrix(.Row, .ColIndex("药品名称与编码")) = rsRecord!通用名
                    .EditText = rsRecord!通用名
                    If .Row = .rows - 1 Then
                        .rows = .rows + 1
                        .Row = .rows - 1
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfNoMedi_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsfPRI_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    mblnPri = True
    mintRow = Row
    mintCol = Col
    With Me.picPRI
        .Visible = True
    
        .Height = vsfPri.Height
        .Top = sstMain.Top + vsfPri.Top
        .Left = sstMain.Left + vsfPri.Left
        .Width = vsfPri.Width
    End With
            
    Select Case Col
        Case vsfPri.ColIndex("科室名称")
            With Me.lvwPRI
                .ListItems.Clear
                .ListItems.Add , "_" & 0, "所有科室", 1, 1
                mRsDept.MoveFirst
                Do While Not mRsDept.EOF
                    .ListItems.Add , "_" & mRsDept!Id, mRsDept!名称, 1, 1
                    mRsDept.MoveNext
                Loop
                .ListItems.Add , "_00", "其他科室", 1, 1
            End With
        Case vsfPri.ColIndex("配药类型")
            With Me.lvwPRI
                .ListItems.Clear
                If mRsType.RecordCount > 0 Then mRsType.MoveFirst
                Do While Not mRsType.EOF
                    .ListItems.Add , "_" & mRsType!编码, mRsType!名称, 1, 1
                    mRsType.MoveNext
                Loop
                 .ListItems.Add , "_00", "其他类型", 1, 1
            End With
        Case vsfPri.ColIndex("频次")
            With Me.lvwPRI
                .ListItems.Clear
                .ListItems.Add , "_" & 0, "所有频次", 1, 1
                mRsPC.MoveFirst
                Do While Not mRsPC.EOF
                    .ListItems.Add , "_" & mRsPC!编码, mRsPC!名称 & "(" & mRsPC!英文名称 & ")", 1, 1
                    mRsPC.MoveNext
                Loop
                .ListItems.Add , "_00", "其他频次", 1, 1
            End With
    End Select
End Sub

Private Sub vsfPrint_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If Me.vsfPrint.rows = 2 Then
            Me.vsfPrint.TextMatrix(vsfPrint.Row, vsfPrint.ColIndex("药品id")) = ""
            Me.vsfPrint.TextMatrix(vsfPrint.Row, vsfPrint.ColIndex("药品名称与编码")) = ""
        Else
            Me.vsfPrint.RemoveItem vsfPrint.Row
        End If
        
    End If
End Sub

Private Sub vsfPrint_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim i As Integer
    Dim strKey As String
    Dim StrCode As String
    
    If KeyCode = 13 Then
        vRect = zlControl.GetControlRect(vsfPrint.hWnd)
        dblLeft = vRect.Left + vsfPrint.CellLeft
        dblTop = vRect.Top + vsfPrint.CellTop + vsfPrint.CellHeight + 3200
        
        With vsfPrint
            If Col = .ColIndex("药品名称与编码") Then
                strKey = Trim(.EditText)
                If strKey = "" Then Exit Sub
                
                If IsNumeric(strKey) Then
                    '纯数字
                    StrCode = " d.编码 like [1] "
                ElseIf zlCommFun.IsCharAlpha(strKey) Then
                    '纯字母
                    StrCode = " n.简码 Like [1] "
                ElseIf zlCommFun.IsCharChinese(strKey) Then
                    '纯汉字
                    StrCode = " d.名称 like [1] "
                Else
                    StrCode = " (n.简码 Like [1] Or d.编码 Like [1] Or n.名称 Like [1]) "
                End If
                                
                gstrSQL = "Select Distinct d.Id ,'【' || d.编码 || '】' || d.名称 || '(' || d.规格 || ')' As 通用名" & vbNewLine & _
                    " From 药品规格 T, 收费项目目录 D, 收费项目别名 N" & vbNewLine & _
                    " Where t.药品id = d.Id And t.药品id = n.收费细目id And D.类别 In ('5', '6') And" & StrCode & vbNewLine & _
                    " And (d.撤档时间 Is Null Or To_Char(d.撤档时间, 'yyyy-MM-dd') = '3000-01-01')" & vbNewLine & _
                    " Order By '【' || d.编码 || '】' || d.名称 || '(' || d.规格 || ')'"
                Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "药品名称与编码", False, "", "", False, False, _
                True, dblLeft, dblTop, .Height, blnCancel, False, True, IIf(gstrMatchMethod = 0, "", "%") & UCase(.EditText) & "%")
    
                If rsRecord Is Nothing Then
                    .EditText = ""
                    Exit Sub
                Else
                    For i = 1 To .rows - 1
                        If rsRecord!Id = Val(.TextMatrix(i, .ColIndex("药品ID"))) Then
                            MsgBox rsRecord!通用名 & "已经录入，请重新选择！", vbInformation + vbOKOnly, gstrSysName
                            .EditText = ""
                            Exit Sub
                        End If
                    Next
                    
                    .TextMatrix(.Row, .ColIndex("药品ID")) = rsRecord!Id
                    .TextMatrix(.Row, .ColIndex("药品名称与编码")) = rsRecord!通用名
                    .EditText = rsRecord!通用名
                    If .Row = .rows - 1 Then
                        .rows = .rows + 1
                        .Row = .rows - 1
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfPrint_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub vsfVolume_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = Me.vsfVolume.ColIndex("容量") Then
        If Not IsNumeric(vsfVolume.TextMatrix(Row, Col)) Then
            MsgBox "容量请录入数字！", vbInformation + vbOKOnly, gstrSysName
            vsfVolume.Col = vsfVolume.ColIndex("容量")
        End If
    End If
End Sub

Private Sub vsfVolume_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str批次 As String
    Dim i As Integer
    
    If Col <> vsfVolume.ColIndex("配药批次") Then Exit Sub
    With Me.vsfBatch
        If .rows > 2 Then
            For i = 2 To .rows - 1
                If .TextMatrix(i, .ColIndex("批次")) <> "" And .TextMatrix(i, .ColIndex("启用")) <> "" Then
                    str批次 = IIf(str批次 = "", "", str批次 & "|") & .TextMatrix(i, .ColIndex("批次"))
                End If
            Next
        End If
        If str批次 <> "" Then Me.vsfVolume.ColComboList(vsfVolume.ColIndex("配药批次")) = str批次
    End With
End Sub

Private Sub vsfVolume_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'    mblnPri = False
'    mintRow = Row
'
'    mintCol = Col
''    With Me.picPRI
'        .Visible = True
'        .Height = vsfVolume.Height
'        .Top = sstMain.Top + vsfPri.Top
'        .Left = sstMain.Left + vsfVolume.Left
'        .Width = vsfVolume.Width
'    End With
'
'    With vsfVolume
'        If Col = .ColIndex("科室名称") Then
'            With Me.lvwPRI
'                .ListItems.Clear
'                .ListItems.Add , "_" & 0, "所有科室", 1, 1
'                mRsDept.MoveFirst
'                Do While Not mRsDept.EOF
'                    .ListItems.Add , "_" & mRsDept!Id, mRsDept!名称, 1, 1
'                    mRsDept.MoveNext
'                Loop
'                .ListItems.Add , "_00", "其他科室", 1, 1
'            End With
'        End If
'    End With

    mblnPri = False
    mintRow = vsfVolume.Row
    mintCol = vsfVolume.Col

    With Me.lvwPRI
        .ListItems.Clear
        .ListItems.Add , "_" & 0, "所有科室", 1, 1
        mRsDept.MoveFirst
        Do While Not mRsDept.EOF
            If vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col) <> "" Then
                If mRsDept!简码 = UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col)) Or mRsDept!五笔简码 = UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col)) Or vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col) = mRsDept!名称 Then
                    vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col) = mRsDept!名称
                    vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.ColIndex("科室id")) = mRsDept!Id
                    Exit Sub

                ElseIf InStr(1, mRsDept!五笔简码, UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col))) > 0 Or InStr(1, mRsDept!简码, UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col))) > 0 Or InStr(1, mRsDept!名称, vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col)) > 0 Then
                    .ListItems.Add , "_" & mRsDept!Id, mRsDept!名称, 1, 1
                End If
            Else
                .ListItems.Add , "_" & mRsDept!Id, mRsDept!名称, 1, 1
            End If
            mRsDept.MoveNext
        Loop
        
        If .ListItems.count = 1 Then
            .ListItems.Clear
            MsgBox "你输入的简码没有与之匹配的科室，请重新录入！"
            Exit Sub
        End If
        
        .ListItems.Add , "_00", "其他科室", 1, 1
    End With
    

    With Me.picPRI
        .Visible = True
        .Height = vsfVolume.Height
        .Top = sstMain.Top + vsfPri.Top
        .Left = sstMain.Left + vsfVolume.Left
        .Width = vsfVolume.Width
    End With
End Sub

Private Sub loadVolume()
    Dim rstemp As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select 科室id,科室名称,容量,配药批次 from 科室容量设置 where 配置中心ID=[1] order by 科室id"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取科室容量数据", Me.CboStore.ItemData(Me.CboStore.ListIndex))
    
    i = 1
    With Me.vsfVolume
        .RowHeight(0) = 250
        .rows = 1
        .rows = IIf(rstemp.RecordCount = 0, 1, rstemp.RecordCount) + 1
        Do While Not rstemp.EOF
            .RowHeight(i) = 250
            .TextMatrix(i, .ColIndex("科室id")) = rstemp!科室ID
            .TextMatrix(i, .ColIndex("科室名称")) = rstemp!科室名称
            .TextMatrix(i, .ColIndex("配药批次")) = zlStr.NVL(rstemp!配药批次)
            .TextMatrix(i, .ColIndex("容量")) = rstemp!容量
            i = i + 1
            rstemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfVolume_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> 13 Then Exit Sub
'
'    With Me.vsfVolume
'        If .Row = .rows - 1 Then
'            If .Col = .Cols - 1 Then
'                Exit Sub
'            Else
'                .Col = .Col + 1
'            End If
'        Else
'            If .Col = .Cols - 1 Then
'                .Row = .Row + 1
'                .Col = .ColIndex("科室名称")
'            Else
'                .Col = .Col + 1
'            End If
'        End If
'    End With
End Sub

Private Sub vsfVolume_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    
    With Me.vsfVolume
        If .Row = .rows - 1 Then
            If .Col = .Cols - 1 Then
                Exit Sub
            Else
                .Col = .Col + 1
            End If
        Else
            If .Col = .Cols - 1 Then
                .Row = .Row + 1
                .Col = .ColIndex("科室名称")
            Else
                .Col = .Col + 1
            End If
        End If
    End With
End Sub

Private Sub vsfVolume_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    With vsfVolume
        If Col = .ColIndex("容量") Then
            If InStr("1234567890-." & Chr(8), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End With

End Sub

Private Sub LoadDept()
    Dim i As Integer
    
    i = 1
    vsfDept.rows = mRsDept.RecordCount + 2
    Do While Not mRsDept.EOF
        With Me.vsfDept
            i = i + 1
            .TextMatrix(i, .ColIndex("序号")) = i - 1
            .TextMatrix(i, .ColIndex("科室id")) = mRsDept!Id
            .TextMatrix(i, .ColIndex("科室名称")) = mRsDept!名称
            .TextMatrix(i, .ColIndex("简码")) = mRsDept!简码
        End With
        mRsDept.MoveNext
    Loop
End Sub

Private Sub vsfPrint_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> Me.vsfPrint.ColIndex("药品名称与编码") Then Cancel = True
End Sub

Private Sub vsfNoMedi_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> Me.vsfNoMedi.ColIndex("药品名称与编码") Then Cancel = True
End Sub

Private Sub vsfNoMedi_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If KeyCode = 46 Then
        If Me.vsfNoMedi.rows = 2 Then
            Me.vsfNoMedi.TextMatrix(vsfNoMedi.Row, vsfNoMedi.ColIndex("药品id")) = ""
            Me.vsfNoMedi.TextMatrix(vsfNoMedi.Row, vsfNoMedi.ColIndex("药品名称与编码")) = ""
        Else
            Me.vsfNoMedi.RemoveItem vsfNoMedi.Row
        End If
    End If
End Sub
