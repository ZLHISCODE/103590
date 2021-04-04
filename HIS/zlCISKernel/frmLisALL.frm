VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLisALL 
   Caption         =   "检验报告查阅"
   ClientHeight    =   10455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18135
   Icon            =   "frmLisALL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   18135
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   6480
      ScaleHeight     =   1695
      ScaleWidth      =   2295
      TabIndex        =   50
      Top             =   240
      Width           =   2295
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   735
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   615
         _Version        =   589884
         _ExtentX        =   1085
         _ExtentY        =   1296
         _StockProps     =   0
         AutoColumnSizing=   0   'False
      End
   End
   Begin MSComDlg.CommonDialog cdgPrint 
      Left            =   4545
      Top             =   915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picWB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   12810
      ScaleHeight     =   1455
      ScaleWidth      =   2760
      TabIndex        =   48
      Top             =   195
      Width           =   2760
      Begin SHDocVwCtl.WebBrowser webSub 
         Height          =   1320
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   2010
         ExtentX         =   3545
         ExtentY         =   2328
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   915
      Left            =   525
      TabIndex        =   0
      Top             =   720
      Width           =   1110
      _Version        =   589884
      _ExtentX        =   1958
      _ExtentY        =   1614
      _StockProps     =   64
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7590
      Left            =   645
      ScaleHeight     =   7590
      ScaleWidth      =   15240
      TabIndex        =   1
      Top             =   1785
      Width           =   15240
      Begin VB.Frame fraLR 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2745
         Left            =   5760
         MousePointer    =   9  'Size W E
         TabIndex        =   6
         Top             =   1320
         Width           =   45
      End
      Begin VB.Frame fraRpt 
         Height          =   5115
         Left            =   4050
         TabIndex        =   5
         Top             =   1785
         Width           =   7650
         Begin XtremeSuiteControls.TabControl tbcArchive 
            Height          =   1290
            Left            =   915
            TabIndex        =   47
            Top             =   570
            Width           =   2055
            _Version        =   589884
            _ExtentX        =   3625
            _ExtentY        =   2275
            _StockProps     =   64
         End
      End
      Begin VB.Frame fraPList 
         Height          =   5250
         Left            =   510
         TabIndex        =   4
         Top             =   1305
         Width           =   5235
         Begin XtremeSuiteControls.TabControl tbcHistory 
            Height          =   1395
            Left            =   345
            TabIndex        =   46
            Top             =   1365
            Width           =   855
            _Version        =   589884
            _ExtentX        =   1508
            _ExtentY        =   2461
            _StockProps     =   64
         End
      End
      Begin VB.PictureBox picBar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   480
         ScaleHeight     =   510
         ScaleWidth      =   1425
         TabIndex        =   3
         Top             =   105
         Width           =   1425
         Begin XtremeCommandBars.CommandBars cbsMain 
            Left            =   0
            Top             =   0
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
         End
      End
      Begin MSComctlLib.StatusBar stbThis 
         Bindings        =   "frmLisALL.frx":6852
         Height          =   360
         Left            =   1290
         TabIndex        =   2
         Top             =   7080
         Width           =   12450
         _ExtentX        =   21960
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Bevel           =   2
               Object.Width           =   2355
               MinWidth        =   882
               Picture         =   "frmLisALL.frx":6866
               Text            =   "中联软件"
               TextSave        =   "中联软件"
               Key             =   "ZLFLAG"
               Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   19526
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
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   3990
         ScaleHeight     =   975
         ScaleWidth      =   7695
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   495
         Width           =   7695
         Begin VB.Frame fraInfo 
            Caption         =   " 基本就诊信息 "
            Height          =   840
            Left            =   90
            TabIndex        =   8
            Top             =   60
            Width           =   7500
            Begin VB.Frame fraIn 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   450
               Left            =   195
               TabIndex        =   27
               Top             =   255
               Visible         =   0   'False
               Width           =   7170
               Begin VB.Label lbl类型zy 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   4770
                  TabIndex        =   45
                  Top             =   0
                  Width           =   1080
               End
               Begin VB.Label lbl类型zy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "类型:"
                  Height          =   180
                  Index           =   0
                  Left            =   4305
                  TabIndex        =   44
                  Top             =   0
                  Width           =   450
               End
               Begin VB.Label lbl住院号zy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "住院号:"
                  Height          =   180
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   43
                  Top             =   0
                  Width           =   630
               End
               Begin VB.Label lbl姓名zy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "姓名:"
                  Height          =   180
                  Index           =   0
                  Left            =   0
                  TabIndex        =   42
                  Top             =   0
                  Width           =   450
               End
               Begin VB.Label lbl付款zy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "付款:"
                  Height          =   180
                  Index           =   0
                  Left            =   0
                  TabIndex        =   41
                  Top             =   255
                  Width           =   450
               End
               Begin VB.Label lbl床号zy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "床号:"
                  Height          =   180
                  Index           =   0
                  Left            =   3150
                  TabIndex        =   40
                  Top             =   0
                  Width           =   450
               End
               Begin VB.Label lbl医保号zy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "医保号:"
                  Height          =   180
                  Index           =   0
                  Left            =   5940
                  TabIndex        =   39
                  Top             =   0
                  Width           =   630
               End
               Begin VB.Label lbl入院zy 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "入院:"
                  Height          =   180
                  Index           =   0
                  Left            =   4305
                  TabIndex        =   38
                  Top             =   255
                  Width           =   450
               End
               Begin VB.Label lbl病况zy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "病况:"
                  Height          =   180
                  Index           =   0
                  Left            =   3150
                  TabIndex        =   37
                  Top             =   255
                  Width           =   450
               End
               Begin VB.Label lbl护理zy 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "护  理:"
                  Height          =   180
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   36
                  Top             =   255
                  Width           =   630
               End
               Begin VB.Label lbl护理zy 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   2190
                  TabIndex        =   35
                  Top             =   255
                  Width           =   900
               End
               Begin VB.Label lbl病况zy 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H000000FF&
                  Height          =   180
                  Index           =   1
                  Left            =   3585
                  TabIndex        =   34
                  Top             =   255
                  Width           =   675
               End
               Begin VB.Label lbl入院zy 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   4770
                  TabIndex        =   33
                  Top             =   255
                  Width           =   90
               End
               Begin VB.Label lbl医保号zy 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00008000&
                  Height          =   180
                  Index           =   1
                  Left            =   6600
                  TabIndex        =   32
                  Top             =   0
                  Width           =   90
               End
               Begin VB.Label lbl床号zy 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   3585
                  TabIndex        =   31
                  Top             =   0
                  Width           =   675
               End
               Begin VB.Label lbl付款zy 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   435
                  TabIndex        =   30
                  Top             =   255
                  Width           =   1080
               End
               Begin VB.Label lbl姓名zy 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   435
                  TabIndex        =   29
                  Top             =   0
                  Width           =   1080
               End
               Begin VB.Label lbl住院号zy 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   2190
                  TabIndex        =   28
                  Top             =   0
                  Width           =   900
               End
            End
            Begin VB.Frame fraOut 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   450
               Left            =   195
               TabIndex        =   9
               Top             =   240
               Visible         =   0   'False
               Width           =   7170
               Begin VB.Label lbl急 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "急"
                  BeginProperty Font 
                     Name            =   "黑体"
                     Size            =   21.75
                     Charset         =   134
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   435
                  Left            =   6705
                  TabIndex        =   26
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   435
               End
               Begin VB.Label lbl挂号单mz 
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   3870
                  TabIndex        =   25
                  Top             =   0
                  Width           =   1065
               End
               Begin VB.Label lbl挂号单mz 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "挂号单:"
                  Height          =   180
                  Index           =   0
                  Left            =   3255
                  TabIndex        =   24
                  Top             =   0
                  Width           =   630
               End
               Begin VB.Label lbl医生mz 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   2385
                  TabIndex        =   23
                  Top             =   0
                  Width           =   780
               End
               Begin VB.Label lbl医生mz 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "医生:"
                  Height          =   180
                  Index           =   0
                  Left            =   1935
                  TabIndex        =   22
                  Top             =   0
                  Width           =   450
               End
               Begin VB.Label lbl社区号mz 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00008000&
                  Height          =   180
                  Index           =   1
                  Left            =   5655
                  TabIndex        =   21
                  Top             =   255
                  Width           =   90
               End
               Begin VB.Label lbl社区号mz 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "社区号:"
                  Height          =   180
                  Index           =   0
                  Left            =   5025
                  TabIndex        =   20
                  Top             =   255
                  Width           =   630
               End
               Begin VB.Label lbl门诊号mz 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "门诊号:"
                  Height          =   180
                  Index           =   0
                  Left            =   3240
                  TabIndex        =   19
                  Top             =   255
                  Width           =   630
               End
               Begin VB.Label lbl姓名mz 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "姓名:"
                  Height          =   180
                  Index           =   0
                  Left            =   0
                  TabIndex        =   18
                  Top             =   0
                  Width           =   450
               End
               Begin VB.Label lbl费别mz 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "费别:"
                  Height          =   180
                  Index           =   0
                  Left            =   1935
                  TabIndex        =   17
                  Top             =   255
                  Width           =   450
               End
               Begin VB.Label lbl医保号mz 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "医保号:"
                  Height          =   180
                  Index           =   0
                  Left            =   5025
                  TabIndex        =   16
                  Top             =   0
                  Width           =   630
               End
               Begin VB.Label lbl医保号mz 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00008000&
                  Height          =   180
                  Index           =   1
                  Left            =   5655
                  TabIndex        =   15
                  Top             =   0
                  Width           =   90
               End
               Begin VB.Label lbl费别mz 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   2385
                  TabIndex        =   14
                  Top             =   255
                  Width           =   765
               End
               Begin VB.Label lbl姓名mz 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   450
                  TabIndex        =   13
                  Top             =   0
                  Width           =   1425
               End
               Begin VB.Label lbl门诊号mz 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   3870
                  TabIndex        =   12
                  Top             =   255
                  Width           =   1095
               End
               Begin VB.Label lbl付款mz 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "付款:"
                  Height          =   180
                  Index           =   0
                  Left            =   0
                  TabIndex        =   11
                  Top             =   255
                  Width           =   450
               End
               Begin VB.Label lbl付款mz 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   450
                  TabIndex        =   10
                  Top             =   255
                  Width           =   1455
               End
            End
         End
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5040
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":70FA
            Key             =   "未执行"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":7694
            Key             =   "已执行"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":7C2E
            Key             =   "拒绝执行"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":81C8
            Key             =   "正在执行"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":8762
            Key             =   "已报到"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":8CFC
            Key             =   "CheckCol"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":9296
            Key             =   "Path"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":9830
            Key             =   "已核对"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":9DCA
            Key             =   "printer"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2880
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":1062C
            Key             =   "住院"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":16E8E
            Key             =   "object_report"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":17428
            Key             =   "object_case"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":179C2
            Key             =   "object_tend"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":17F5C
            Key             =   "object_first"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":184F6
            Key             =   "object_advice"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":18A90
            Key             =   "object_file"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":1902A
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":1F88C
            Key             =   "Path"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLisALL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PATIREPORT_COLUMN
    COL_选择
    COL_图标
    COL_项目 '报告名称
    COL_审核时间
    COL_标本
    COL_申请时间
    COL_申请人
    COL_采集时间
    
    '隐藏列
    COL_医嘱ID
    COL_报告ID
    COL_类型
    COL_打印次数
    COL_文档标题
    COL_禁止打印
End Enum

Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fmask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private mblnMoved As Boolean
Private mlng病人ID As Long
Private mlng就诊ID As Long
Private mlngPreIndex As Long
Private mrsData As ADODB.Recordset
Private mlngPreDept As Long
Private mstr挂号单 As String
Private mlngModel As Long '1252-p门诊医嘱下达，1253-p住院医嘱下达
Private mstrMainPrivs As String
Private mlng科室ID  As Long
Private mlng病区ID  As Long
Private mfrmLisCom As Object 'LIS检验接口返回的窗体，区分新版和老版
Private mstrFilesTemp As String '本次生成的临时文件目录串用<STAB>分割
Private mstrCurFile As String '当前网页控件占用的文件目录
Private mobjPopup As CommandBarPopup
Private mlngPages As Long
Private Const M_S As String = "<SPL>" '字符串分割符
Private mcolCpt As Collection '集合，标题
Private mstrCpage As String '当前选的中这次就诊的标题
Private mlng医嘱ID As Long
Private mlng报告ID As Long
Private mlng类型 As Long '0-PDF文件
Private mstrPrePati As String
Private mfrmPDF As Object '包含了PDF控件对象的窗体
Private mblnPDF As Boolean
Private mclsPDF As Object
Private mbln禁止打印 As Boolean

Public Function ShowMe(ByVal frmParent As Object, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal lng科室id As Long, ByVal lng病区ID As Long, ByVal lngModel As Long, ByVal strPrivs As String) As Boolean
'功能：公共接口
'参数：
'      frmParent     父窗体对
'      lng病人ID     病人id
'      lng就诊ID     门诊病人对应挂号id,住院病人为主页id
'      lng科室ID     病人所在科室，挂号科室/病人当前科室
'      lng病区ID     病人所在病区，门诊病人传0
'      lngModel      模块号，1252-p门诊医嘱下达，1253-p住院医嘱下达
'      strPrivs      权限

    mblnMoved = False
    mlng病人ID = lng病人ID
    mlng就诊ID = lng就诊ID
    mlngModel = lngModel
    mlng科室ID = lng科室id
    mlng病区ID = lng病区ID
    mstrMainPrivs = strPrivs
    mlngPreDept = -1
    Me.Show , frmParent
End Function

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim arrpara As Variant
    
    arrpara = Split(mobjPopup.Parameter, ",")
    mobjPopup.Visible = mlngPages > 0
    
    Select Case Control.ID
    Case conMenu_View_Forward, conMenu_View_Backward   '上一页,下一页
        Control.Visible = mlngPages > 1
        If Control.Visible Then
            If Val(arrpara(0)) = 1 And Control.ID = conMenu_View_Forward Then
                Control.Enabled = False
            ElseIf Val(arrpara(0)) = mlngPages And Control.ID = conMenu_View_Backward Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        End If
    Case conMenu_File_Print
        Control.Enabled = Not mbln禁止打印
    End Select
End Sub

Private Sub Form_Load()
    Dim intIdx As Integer
    Dim strErr As String
    Dim objTab As Object
    Dim strFile As String
    
    '初始化菜单
    Call InitBar
    strFile = mstrCurFile
    Call CreatePDFobj
    Call InitObjLis(mlngModel)
    If Not gobjLIS Is Nothing Then
        If mlngModel = p门诊医嘱下达 Then
            Call gobjLIS.PatientSampleBrowse(Me, mlng病人ID, mstrMainPrivs, mlng科室ID, 0, 1, 0, strErr, False, mfrmLisCom)
        Else
            Call gobjLIS.PatientSampleBrowse(Me, mlng病人ID, mstrMainPrivs, mlng科室ID, mlng病区ID, 2, mlng就诊ID, strErr, False, mfrmLisCom)
        End If
        If strErr <> "" Then
            MsgBox strErr, vbInformation, Me.Caption
        End If
    Else
        Call frmLisView.ShowMe(mlng病人ID, mlngModel, Me, False, mfrmLisCom)
    End If
         
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(intIdx, "普通检验结果", mfrmLisCom.hwnd, 0).Tag = "普通": intIdx = intIdx + 1
        .InsertItem(intIdx, "微生物/细胞学报告", picMain.hwnd, 0).Tag = "微生物": intIdx = intIdx + 1
        .Item(0).Selected = True '新建时就自动选中了这个,不会再激活事件
    End With
    
    '信息展示区域
    With tbcArchive
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .Color = xtpTabColorOffice2003
            .Layout = xtpTabLayoutAutoSize
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        Set objTab = .InsertItem(intIdx, "报告", picWB.hwnd, 0): objTab.Tag = objTab.Caption
        If mblnPDF Then
            intIdx = intIdx + 1
            Set objTab = .InsertItem(intIdx, "报告", mfrmPDF.hwnd, 0): objTab.Tag = objTab.Caption
                objTab.Visible = False
        Else
            .Item(0).Selected = True
        End If
    End With
    
    '就诊历史列表
    '-----------------------------------------------------
    With tbcHistory
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
            .DisableLunaColors = False
            .BoldSelected = True
            .HotTracking = True
            .ShowIcons = True
        End With
        .SetImageList ils16
    End With
    Call InitReportColumn
    '历次就诊信息
    Call InitBasicData
    mstrCurFile = strFile
    Call DeleteLISTempFile(0)
    mstrCurFile = ""
    Call RestoreWinState(Me, App.ProductName)
    stbThis.Visible = True
End Sub

Private Sub InitBar()
'功能：菜单初始化
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    
    '工具栏----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
     
    '工具栏定义:包括公共部份
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False

    With objBar.Controls
    
        Set mobjPopup = .Add(xtpControlPopup, conMenu_Edit_NewItem, "页面")
        mobjPopup.IconId = conMenu_Edit_Modify
        mobjPopup.Style = xtpButtonIconAndCaption
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Forward, "上一页", -1, False)
        objControl.IconId = conMenu_View_Forward
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Backward, "下一页", -1, False)
        objControl.IconId = conMenu_View_Backward
         
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印机设置")
            objControl.IconId = conMenu_File_Parameter
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
'        Set objControl = .Add(xtpControlButton, conMenu_Tool_Reference, "参考")
'            objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngPage As Long
    Dim strCaption As String
    Dim arrpara As Variant
    Dim i As Long
    
    Select Case Control.ID
        Case conMenu_View_Jump
            arrpara = Split(Control.Parameter, M_S)
            lngPage = Val(arrpara(0))
            strCaption = arrpara(1)
        Case conMenu_View_Forward, conMenu_View_Backward   '上一页,下一页
            arrpara = Split(mobjPopup.Parameter, M_S)
            strCaption = arrpara(1)
            If Control.ID = conMenu_View_Forward Then
                lngPage = Val(arrpara(0)) - 1
            ElseIf Control.ID = conMenu_View_Backward Then
                lngPage = Val(arrpara(0)) + 1
            End If
        Case conMenu_File_Print '打印
            Call FuncPrint
        Case conMenu_Tool_Reference
            Call FunReference
        Case conMenu_File_PrintSet
            cdgPrint.ShowPrinter
    End Select
    
    Select Case Control.ID
    Case conMenu_View_Jump, conMenu_View_Forward, conMenu_View_Backward
        mstrCpage = mcolCpt("" & lngPage)
        mobjPopup.Caption = mstrCpage
        mobjPopup.Parameter = lngPage & M_S & strCaption
        mobjPopup.SetFocus
        For i = 1 To mobjPopup.CommandBar.Controls.Count
            If lngPage = i Then
                mobjPopup.CommandBar.Controls(i).Checked = True
            Else
                mobjPopup.CommandBar.Controls(i).Checked = False
            End If
        Next
        Call LoadFaceData(mlngPages - lngPage + 1)
        cbsMain.RecalcLayout
    End Select
End Sub

Private Sub InitBasicData()
'功能：初始化一些基本数据，如下拉列表加载等
    Dim strSQL As String
    Dim objTab As TabControlItem
    Dim strTmp As String
    Dim str病人IDs As String
    Dim rsTmp As ADODB.Recordset
    Dim str身份证号 As String
    Dim strTemp As String
    Dim n As Long, p As Long
    Dim strThis As String
    Dim strSQLPati As String
    Dim varPar(0 To 10) As String
    Dim objControl As CommandBarControl
    Dim objCloc As CommandBarControl '用于定位的那一次就诊
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    If Not mobjPopup Is Nothing Then mobjPopup.CommandBar.Controls.DeleteAll
    
    mlngPreIndex = -1
    strSQL = "select a.身份证号 from 病人信息 a where a.病人id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    strTmp = rsTmp!身份证号 & ""
    If strTmp <> "" Then
        '验证身份证号的合法性
        If InitObjPublicPatient Then
            If gobjPublicPatient.CheckPatiIdcard(strTmp) Then
                str身份证号 = strTmp
            End If
        End If
    End If
         
    '通过身份证号查关联
    If str身份证号 <> "" Then
        strSQL = "select a.病人id from 病人信息 a where a.病人id<>[1] and a.身份证号=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, str身份证号)
        Do While Not rsTmp.EOF
            str病人IDs = str病人IDs & "," & rsTmp!病人ID
            rsTmp.MoveNext
        Loop
        str病人IDs = Mid(str病人IDs, 2)
    End If
    
    '通过关联表查关联
'    strTmp = GetPatiRelate(mlng病人ID, str身份证号)
'    If strTmp <> "" Then
'        If str病人IDs <> "" Then
'            str病人IDs = str病人IDs & "," & strTmp
'        Else
'            str病人IDs = strTmp
'        End If
'    End If
    
    If str病人IDs = "" Then
        strSQL = " Select 病人id,ID as 就诊ID,NO,发生时间 as 开始时间,Null as 结束时间,执行部门ID as 科室ID,0 as 数据转出,-1 as 病人性质,null as 就诊号 From 病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
            " Union ALL" & _
            " Select 病人id,ID as 就诊ID,NO,发生时间 as 开始时间,Null as 结束时间,执行部门ID as 科室ID,1 as 数据转出,-1 as 病人性质,null as 就诊号 From H病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
            " Union ALL" & _
            " Select 病人id,主页ID as 就诊ID,Null,入院日期 as 开始时间,出院日期,出院科室ID,数据转出,NVL(病人性质,0) as 病人性质,null as 就诊号 From 病案主页 Where 病人ID=[1] And Nvl(主页ID,0)<>0"
        strSQL = "Select Rownum As 序号,a.病人ID,A.就诊ID,A.NO,A.开始时间,A.结束时间,B.名称 as 科室,A.数据转出 ,A.病人性质,a.就诊号 From (" & strSQL & ") A,部门表 B Where A.科室ID=B.ID Order by 开始时间 Desc"
        Set mrsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    Else
        str病人IDs = mlng病人ID & "," & str病人IDs
        
        '大于4000长度的拆分
        strTemp = "Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X"
        n = 0
        Do While True
            If Len(str病人IDs) < 4000 Then
                p = Len(str病人IDs) + 1
            Else
                p = InStrRev(Mid(str病人IDs, 1, 4000), ",")
            End If
            strThis = Mid(str病人IDs, 1, p - 1)
            
            If n > 10 Then
                strSQLPati = strSQLPati & vbNewLine & " Union All " & Replace(strTemp, "[1]", "'" & strThis & "'")
            Else
                varPar(n) = strThis
                strSQLPati = IIF(strSQLPati = "", "", strSQLPati & vbNewLine & " Union All ") & Replace(strTemp, "[1]", "[" & (n + 1) & "]")
            End If
            
            n = n + 1
            str病人IDs = Mid(str病人IDs, p + 1)
            If str病人IDs = "" Then Exit Do
        Loop
        strTmp = " 病人ID In (" & strSQLPati & ")"
        
        
        strSQL = " Select 病人id,ID as 就诊ID,NO,发生时间 as 开始时间,Null as 结束时间,执行部门ID as 科室ID,0 as 数据转出,-1 as 病人性质,null as 就诊号 From 病人挂号记录 Where " & strTmp & " And 记录性质=1 And 记录状态=1 and NO is not null" & _
            " Union ALL" & _
            " Select 病人id,ID as 就诊ID,NO,发生时间 as 开始时间,Null as 结束时间,执行部门ID as 科室ID,1 as 数据转出,-1 as 病人性质,null as 就诊号 From H病人挂号记录 Where " & strTmp & " And 记录性质=1 And 记录状态=1 and NO is not null" & _
            " Union ALL" & _
            " Select 病人id,主页ID as 就诊ID,Null,入院日期 as 开始时间,出院日期,出院科室ID,数据转出,NVL(病人性质,0) as 病人性质,住院号 as 就诊号 From 病案主页 Where " & strTmp & " And Nvl(主页ID,0)<>0"
        strSQL = "Select Rownum As 序号,a.病人ID,A.就诊ID,A.NO,A.开始时间,A.结束时间,B.名称 as 科室,A.数据转出 ,A.病人性质,a.就诊号 From (" & strSQL & ") A,部门表 B Where A.科室ID=B.ID  Order by 开始时间 Desc"
        Set mrsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9), varPar(10))
    End If
    
    Set mcolCpt = New Collection
    n = 0
    mlngPages = mrsData.RecordCount
    Do While Not mrsData.EOF
        
        strTmp = IIF(IsNull(mrsData!NO), "第" & mrsData!就诊ID & "次" & IIF(mrsData!病人性质 = 1, "门诊留观", IIF(mrsData!病人性质 = 2, "住院留观", "住院")), "门诊就诊") & ":" & mrsData!科室 & "," & Format(mrsData!开始时间, "yyyy-MM-dd HH:mm") & _
            IIF(Not IsNull(mrsData!结束时间), "～" & Format(mrsData!结束时间, "yyyy-MM-dd HH:mm"), "")
            
        If mrsData.AbsolutePosition = 1 Then
            Set objTab = tbcHistory.InsertItem(tbcHistory.ItemCount, strTmp, picList.hwnd, IIF(IsNull(mrsData!NO), 0, 1))
        End If
        n = n + 1
        Set objControl = mobjPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, strTmp, -1, False)
        objControl.Parameter = n & M_S & Val(mrsData!序号) & M_S & strTmp
        mcolCpt.Add strTmp, n & ""
        If mlng病人ID = Val(mrsData!病人ID & "") And mlng就诊ID = Val(mrsData!就诊ID & "") Then
            Set objCloc = objControl
        End If
        mrsData.MoveNext
    Loop
    
    If objCloc Is Nothing Then
        objControl.Execute
    Else
        objCloc.Execute
    End If
        
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tbcSub.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    mlng医嘱ID = 0
    Set mrsData = Nothing
    Unload mfrmLisCom
    Set mfrmLisCom = Nothing
    Call DeleteLISTempFile(1)
    If Not mfrmPDF Is Nothing Then
        Unload mfrmPDF
        Set mfrmPDF = Nothing
    End If
    Set mclsPDF = Nothing
    mblnPDF = False
    mbln禁止打印 = False
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    
    
    '--界面布局，菜单，左侧，分割线，右上，中间
    
    picBar.Move 0, 0, picMain.ScaleWidth, 500
    stbThis.Width = picMain.ScaleWidth
    stbThis.Left = 0
    stbThis.Top = picMain.ScaleHeight - stbThis.Height
    
    fraPList.Top = picBar.Height
    fraPList.Left = 0
    fraPList.Height = picMain.Height - picBar.Height - stbThis.Height
 
    fraLR.Top = fraPList.Top
    fraLR.Height = fraPList.Height
    fraLR.Left = fraPList.Width
    
    picInfo.Top = fraPList.Top
    picInfo.Left = fraLR.Left + 45
    picInfo.Width = picMain.ScaleWidth - picInfo.Left
    picInfo.Height = 840
    
    fraRpt.Top = picInfo.Top + picInfo.Height
    fraRpt.Left = picInfo.Left
    fraRpt.Width = picInfo.Width
    fraRpt.Height = fraLR.Height - picInfo.Height
    
    
    '--左侧病人列表就诊列表容器内部
    tbcHistory.Move 0, 140, fraPList.Width - 50, fraPList.Height - 150
    
    '中间区域
    tbcArchive.Move 0, 0, fraRpt.Width, fraRpt.Height
    
    '--病人基本信息的背景色
    picInfo.BackColor = fraLR.BackColor
    fraInfo.BackColor = picInfo.BackColor
    fraIn.BackColor = picInfo.BackColor
    fraOut.BackColor = picInfo.BackColor
    
    '--picInfo容器内部控件位置设置
    fraInfo.Left = 0
    fraInfo.Top = 0
    fraInfo.Width = picInfo.Width - fraInfo.Left * 3
    fraIn.Width = fraInfo.Width - fraIn.Left * 2
    fraOut.Width = fraIn.Width
    lbl急.Left = fraOut.Width - lbl急.Width - 60
   
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    If Button = 1 Then
        If fraLR.Left + X < 1000 Or fraLR.Left + X > picMain.Width - 3000 Then Exit Sub
        fraPList.Width = fraPList.Width + X
        Call picMain_Resize
    End If
End Sub
 
Private Function ShowOutPatiInfo() As Boolean
'功能：选择门诊病人某次历史就诊记录时，读取相关的病人信息
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If mlng病人ID <> 0 Then
        strSQL = "Select B.Id,B.NO,B.门诊号,B.姓名,B.性别,B.年龄,A.医疗付款方式," & _
            " A.费别,A.险类,A.医保号,B.急诊,B.发生时间,B.执行人,B.执行状态,B.执行时间," & _
            " B.执行部门ID as 科室ID,B.诊室,B.社区,D.社区号,C.名称 as 科室" & _
            " From 病人信息 A,病人挂号记录 B,部门表 C,病人社区信息 D" & _
            " Where A.病人ID=B.病人ID And B.ID=[1] And B.执行部门ID=C.ID" & _
            " And B.病人ID=D.病人ID(+) And B.社区=D.社区(+) And B.记录性质=1 And B.记录状态=1"
        If mblnMoved Then
            strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng就诊ID)
        With rsTmp
            '保险病人姓名红色显示
            lbl姓名mz(1).Caption = NVL(!姓名)
            If Not IsNull(!险类) Then
                lbl姓名mz(1).ForeColor = vbRed
            Else
                lbl姓名mz(1).ForeColor = lbl门诊号mz(1).ForeColor
            End If
            lbl医生mz(1).Caption = NVL(!执行人)
            lbl挂号单mz(1).Caption = !NO
            lbl门诊号mz(1).Caption = NVL(!门诊号)
            lbl付款mz(1).Caption = NVL(!医疗付款方式)
            lbl费别mz(1).Caption = NVL(!费别)
            lbl医保号mz(1).Caption = NVL(!医保号)
            lbl社区号mz(1).Caption = NVL(!社区号)
            lbl急.Visible = NVL(!急诊, 0) <> 0
        End With
    Else
        fraOut.Visible = True
        lbl姓名mz(1).Caption = ""
        lbl医生mz(1).Caption = ""
        lbl挂号单mz(1).Caption = ""
        lbl门诊号mz(1).Caption = ""
        lbl付款mz(1).Caption = ""
        lbl费别mz(1).Caption = ""
        lbl医保号mz(1).Caption = ""
        lbl社区号mz(1).Caption = ""
    End If
    ShowOutPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowInPatiInfo() As Boolean
'功能：选择某次住院记录时，读取相关的病人信息
'返回：blnMoved=本次住院病案是否转出了
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If mlng病人ID <> 0 Then
        strSQL = "Select NVL(B.姓名,A.姓名) 姓名, NVL(B.性别,A.性别) 性别, NVL(B.年龄,A.年龄) 年龄,B.住院号,B.出院病床,B.医疗付款方式," & _
            " D.信息值 as 医保号,B.险类,B.当前病况,C.名称 as 护理等级,B.入院日期," & _
            " B.出院日期,B.病人类型,B.状态,B.出院科室ID,B.当前病区ID,A.住院次数" & _
            " From 病人信息 A,病案主页 B,收费项目目录 C,病案主页从表 D" & _
            " Where A.病人ID=B.病人ID And A.病人ID=[1] And B.主页ID=[2] And B.护理等级ID=C.ID(+)" & _
            " And B.病人ID=D.病人ID(+) And B.主页ID=D.主页ID(+) And D.信息名(+)='医保号'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID)
        
        With rsTmp
            '保险病人颜色特殊显示
            lbl姓名zy(1).Caption = NVL(!姓名)
            lbl姓名zy(1).ForeColor = zlDatabase.GetPatiColor(NVL(!病人类型))   'GetPatiColor(NVL(!病人类型))
            
            lbl住院号zy(1).Caption = NVL(!住院号)
            lbl床号zy(1).Caption = NVL(!出院病床)
            lbl医保号zy(1).Caption = NVL(!医保号)
            lbl护理zy(1).Caption = NVL(!护理等级)
            lbl付款zy(1).Caption = NVL(!医疗付款方式)
            
            '危重病人病况红色显示
            lbl病况zy(1).Caption = NVL(!当前病况)
            If NVL(!当前病况) = "危" Or NVL(!当前病况) = "重" Or NVL(!当前病况) = "急" Then
                lbl病况zy(1).ForeColor = vbRed
            Else
                lbl病况zy(1).ForeColor = lbl住院号zy(1).ForeColor
            End If
            
            lbl入院zy(1).Caption = Format(!入院日期, "yyyy-MM-dd HH:mm")
            If Not IsNull(!出院日期) Then
                lbl入院zy(1).Caption = lbl入院zy(1).Caption & "～" & Format(!出院日期, "yyyy-MM-dd HH:mm")
            End If
            lbl类型zy(1).Caption = NVL(!病人类型)
        End With
    Else
        '保险病人颜色特殊显示
        fraIn.Visible = True
        lbl姓名zy(1).Caption = ""
        lbl住院号zy(1).Caption = ""
        lbl床号zy(1).Caption = ""
        lbl医保号zy(1).Caption = ""
        lbl护理zy(1).Caption = ""
        lbl付款zy(1).Caption = ""
        lbl病况zy(1).Caption = ""
        lbl入院zy(1).Caption = ""
        lbl类型zy(1).Caption = ""
    End If
    ShowInPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub picWB_Resize()
    On Error Resume Next
    webSub.Move 0, 0, picWB.Width, picWB.Height
End Sub
 
Private Sub WebShow(ByVal strKey As String, ByVal strCpt As String)
'功能：控件展示文件要区分是否加载了PDF控件的情况
'参数: strKey 固定格式,报告ID;医嘱ID;类型
'       strCpt 标题
    Dim strUrl As String '文件路径
    
    If strKey = "" Then
        Call webSub.Navigate("about:blank")
        webSub.Visible = False
        mstrCurFile = ""
        
        If mblnPDF Then
            tbcArchive(1).Visible = False
            tbcArchive(1).Selected = False
        End If
        
        tbcArchive(0).Visible = True
        tbcArchive(0).Selected = True
        tbcArchive(0).Caption = strCpt
    Else
        strUrl = GetLisRptFile(strKey)
        If mblnPDF And mlng类型 = 0 Then
            Call mclsPDF.LoadFile(strUrl)
            tbcArchive(0).Visible = False
            tbcArchive(0).Selected = False
            tbcArchive(1).Visible = True
            tbcArchive(1).Selected = True
            tbcArchive(1).Caption = strCpt
	    mfrmPDF.Enabled = Not mbln禁止打印
        Else
            If strUrl <> "" Then
                webSub.Navigate strUrl
                mstrCurFile = strUrl
            End If
            webSub.Visible = True
            
            If mblnPDF Then
                tbcArchive(1).Visible = False
                tbcArchive(1).Selected = False
            End If
            
            tbcArchive(0).Visible = True
            tbcArchive(0).Selected = True
            tbcArchive(0).Caption = strCpt
	    picWB.Enabled = Not mbln禁止打印
        End If
    End If
End Sub
 
Private Function GetLisRptFile(ByVal strTag As String) As String
'功能：打开LIS报告文件查看，获取临时文件路径
'参数: strTag 入参数,固定格式,报告ID;医嘱ID;类型
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim lng报告ID As Long
    Dim lng医嘱ID As Long
    Dim lng类型 As String
    Dim varTmp As Variant
    Dim strSuffix As String '文件后缀名
    
    Screen.MousePointer = 11
    varTmp = Split(strTag, ";")
    lng报告ID = Val(varTmp(0))
    lng医嘱ID = Val(varTmp(1))
    lng类型 = varTmp(2)
    If lng类型 = 0 Then
        strSuffix = "pdf"
    ElseIf lng类型 = 1 Then
        strSuffix = "html"
    Else
        strSuffix = "xps"
    End If
    strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\tmpReport_" & lng报告ID & "." & strSuffix
    If InStr(mstrFilesTemp, strFile) = 0 Then
        mstrFilesTemp = mstrFilesTemp & "<STAB>" & strFile
    End If
    If Not objFile.FileExists(strFile) Then
        strFile = Sys.ReadLob(glngSys, 22, lng报告ID, strFile)
        If Not objFile.FileExists(strFile) Then
            MsgBox "文件内容读取失败！", vbInformation, gstrSysName:
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    GetLisRptFile = strFile
    Screen.MousePointer = 0
End Function

Private Sub FuncPrint()
'功能：打印当前PDF文件，或者多个文件,可能是多个文件包含了xps和pdf文件
'      多个文件时合并后再打印
    Dim i As Long
    Dim lng报告ID As Long
    Dim lng医嘱ID As Long
    Dim rsPati As ADODB.Recordset
    Dim lngPage As Long
    
    lngPage = mlngPreDept
    Set rsPati = New ADODB.Recordset
    rsPati.Fields.Append "医嘱ID", adBigInt
    rsPati.Fields.Append "报告ID", adBigInt
    rsPati.Fields.Append "类型", adBigInt
    rsPati.Fields.Append "项目", adVarChar, 4000
    rsPati.CursorLocation = adUseClient
    rsPati.LockType = adLockOptimistic
    rsPati.CursorType = adOpenStatic
    rsPati.Open
    
    For i = 0 To rptPati.Rows.Count - 1
        If Not rptPati.Rows(i).GroupRow Then
            If rptPati.Rows(i).Record(COL_选择).Checked Then
                With rptPati.Rows(i).Record
                    If Val(.Item(COL_禁止打印).value) = 0 Then
                        rsPati.AddNew
                        rsPati!医嘱ID = Val(.Item(COL_医嘱ID).value)
                        rsPati!报告ID = Val(.Item(COL_报告ID).value)
                        rsPati!类型 = Val(.Item(COL_类型).value)
                        rsPati!项目 = .Item(COL_项目).value
                        rsPati.Update
                    End If
                End With
            End If
        End If
    Next
    
    '当没有勾选的时候才打印当前选中的这一行
    If rptPati.SelectedRows.Count > 0 And rsPati.RecordCount = 0 Then
        With rptPati.SelectedRows(0)
            lng医嘱ID = Val(.Record(COL_医嘱ID).value)
            lng报告ID = Val(.Record(COL_报告ID).value)
            rsPati.Filter = "医嘱ID=" & lng医嘱ID & " and 报告ID=" & lng报告ID
            If rsPati.EOF Then
                If Val(.Item(COL_禁止打印).value) = 0 Then
                    rsPati.AddNew
                    rsPati!医嘱ID = lng医嘱ID
                    rsPati!报告ID = lng报告ID
                    rsPati!类型 = Val(.Record(COL_类型).value)
                    rsPati!项目 = .Record(COL_项目).value
                    rsPati.Update
                End If
            End If
        End With
    End If
    
    rsPati.Filter = 0
    If rsPati.RecordCount = 0 Then
        Exit Sub
    End If
    
    '无PDF控件时
    If Not mblnPDF Then
        Call PrintNoPdf(rsPati)
    Else
        '有无PDF控件时,区分xps和PDF
        rsPati.Filter = "类型=2"
        If Not rsPati.EOF Then
            Call PrintNoPdf(rsPati)
        End If
        rsPati.Filter = "类型=0"
        If Not rsPati.EOF Then
            Call PrintPdf(rsPati)
        End If
    End If
    Call LoadFaceData(lngPage)
End Sub

Private Sub PrintPdf(ByRef rsPati As ADODB.Recordset)
'功能:打印能应用于PDF控件的文件
    Dim i As Long
    Dim strTag As String
    Dim lng报告ID As Long
    Dim strFileSource As String
    Dim strSQL As String
    Dim lngCnt As Long
    
   On Error GoTo errH

    lngCnt = rsPati.RecordCount
    
    For i = 1 To rsPati.RecordCount
        lng报告ID = rsPati!报告ID
        strTag = rsPati!报告ID & ";" & rsPati!医嘱ID & ";" & rsPati!类型
        strFileSource = GetLisRptFile(strTag)

        Call mclsPDF.LoadFile(strFileSource)
        Call mclsPDF.PrintFile(0)
        
        If lngCnt <> 1 Then
            '超过多份时才弹出等待
            Call mclsPDF.WaitTime(0, strFileSource, rsPati!项目 & "")
        End If
        
        strSQL = "Zl_医嘱报告内容_Print(" & lng报告ID & ",0)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        rsPati.MoveNext
    Next
    mlngPreDept = -1
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub PrintNoPdf(ByRef rsPati As ADODB.Recordset)
'功能:打印不能应用于PDF控件的文件
    Dim i As Long
    Dim strTag As String
    Dim lng报告ID As Long
    Dim strFileSource As String
    
    For i = 1 To rsPati.RecordCount
        strTag = rsPati!报告ID & ";" & rsPati!医嘱ID & ";" & rsPati!类型
        strFileSource = GetLisRptFile(strTag)
        lng报告ID = Split(strTag, ";")(0)
        Call FunFastPrint(strFileSource, lng报告ID)
        rsPati.MoveNext
    Next
    mlngPreDept = -1
End Sub

Private Sub FunFastPrint(ByVal strFile As String, ByVal lngRptID As Long)
'功能：API调用快速打印PDF文件
'参数：strFile 文件路径
    Dim RetVal As Long
    Dim strSQL As String
    Dim ShExInfo As SHELLEXECUTEINFO
    
    On Error GoTo errH
    With ShExInfo
        .cbSize = Len(ShExInfo)
        .fmask = &H40
        .hwnd = 0
        .lpVerb = "print"
        .lpFile = strFile
        .lpParameters = ""
        .lpDirectory = vbNullChar
        .nShow = 0
    End With
    RetVal = ShellExecuteEx(ShExInfo)
    If RetVal = 0 Then
        Exit Sub
    End If
    strSQL = "Zl_医嘱报告内容_Print(" & lngRptID & ",0)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function DeleteLISTempFile(ByVal intType As Integer) As Boolean
'功能：删除临时文件
'参数：intType 0-打开窗体时调用，1-关闭窗口时调用
    Dim objFile As New FileSystemObject
    Dim i As Long
    Dim varFiles As Variant
    Dim strTmp As String
    
    On Error GoTo errH
    If intType = 1 Then
        If mstrFilesTemp = "" Then Exit Function
        Call webSub.Navigate("about:blank")
        varFiles = Split(mstrFilesTemp, "<STAB>")
        For i = 0 To UBound(varFiles)
            strTmp = varFiles(i)
            If strTmp <> "" And strTmp <> mstrCurFile Then
                If objFile.FileExists(strTmp) Then
                    objFile.DeleteFile strTmp, True
                End If
            End If
        Next
        mstrFilesTemp = ""
        strTmp = mstrCurFile
        If strTmp <> "" Then
            On Error Resume Next
            If objFile.FileExists(strTmp) Then
                objFile.DeleteFile strTmp, True
            End If
        End If
    Else
        strTmp = mstrCurFile
        If strTmp <> "" Then
            On Error Resume Next
            If objFile.FileExists(strTmp) Then
                objFile.DeleteFile strTmp, True
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPatiRelate(ByVal lng病人ID As Long, ByVal strIDNo As String) As String
'功能：获取指定人与之关的联的病人ID串，串中不包含当前传入的病人ID
'参数：lng病人ID，病人id;strIDNo 身份证号
'说明：当传入的身份证号不为空字符串时则返回所有病人，当不为空则排除身份证为strIDNo的病人，要求传的身份证号必须是合法的身份证号即通过ZLHIS身份证号验证
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If strIDNo = "" Then
        strSQL = "select b.病人id from 病人身份关联 a,病人身份关联 b where a.关联id=b.关联id and a.病人id=[1] and b.病人id+0<>[1]"
    Else
        strSQL = "select b.病人id from 病人身份关联 a,病人身份关联 b,病人信息 c where a.关联id=b.关联id and b.病人id=c.病人id and a.病人id=[1] and b.病人id+0<>[1] and nvl(c.身份证号,'-')<>[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPatiRelate", lng病人ID, strIDNo)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & rsTmp!病人ID
        rsTmp.MoveNext
    Loop
    
    GetPatiRelate = Mid(strSQL, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadFaceData(ByVal lngPage As Long)

    If mlngPreDept = lngPage Then Exit Sub
    
    mlngPreDept = lngPage
    
    mrsData.Filter = "序号=" & mlngPreDept
    
    mlng就诊ID = mrsData!就诊ID
    mlng病人ID = mrsData!病人ID

    If Not mrsData.EOF Then
        mstr挂号单 = NVL(mrsData!NO, "")
        mblnMoved = Val(NVL(mrsData!数据转出, "")) = 1
    End If
    '显示基本信息
    If mstr挂号单 <> "" Then
        Call ShowOutPatiInfo
    Else
        Call ShowInPatiInfo
    End If
    
    fraOut.Visible = mstr挂号单 <> ""
    fraIn.Visible = mstr挂号单 = ""

    '显示档案目录
    Me.tbcHistory(0).Caption = mstrCpage
    Call LoadPatients
    Call Form_Resize
End Sub

Private Sub picList_Resize()
    On Error Resume Next
    rptPati.Move 0, 0, picList.Width, picList.Height
End Sub

Private Sub InitReportColumn()
'功能:初始化表格
    Dim objCol As ReportColumn

    With rptPati
        Set objCol = .Columns.Add(COL_选择, "", 18, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft
        objCol.Editable = True
        objCol.Icon = 5
         
        Set objCol = .Columns.Add(COL_图标, "", 18, False)
            objCol.Sortable = False
            objCol.Alignment = xtpAlignmentCenter
            objCol.AllowDrag = False
        Set objCol = .Columns.Add(COL_项目, "项目", 120, True)
        Set objCol = .Columns.Add(COL_审核时间, "审核时间", 106, True)
        Set objCol = .Columns.Add(COL_标本, "标本", 60, True)
        Set objCol = .Columns.Add(COL_申请时间, "申请时间", 106, True)
        Set objCol = .Columns.Add(COL_申请人, "申请人", 55, True)
        Set objCol = .Columns.Add(COL_采集时间, "采集时间", 106, True)
        
        '隐藏列
        Set objCol = .Columns.Add(COL_医嘱ID, "医嘱ID", 0, False)
        Set objCol = .Columns.Add(COL_报告ID, "报告ID", 0, False)
        Set objCol = .Columns.Add(COL_类型, "类型", 0, False)
        Set objCol = .Columns.Add(COL_打印次数, "打印次数", 0, False)
        Set objCol = .Columns.Add(COL_文档标题, "文档标题", 0, False)
        Set objCol = .Columns.Add(COL_禁止打印, "禁止打印", 0, False)
        
        For Each objCol In .Columns
            If objCol.Index <> COL_选择 Then objCol.Editable = False
            If objCol.Width = 0 Then objCol.Visible = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的病人..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
End Sub

Private Sub LoadPatients()
'功能：加载病人列表
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    Screen.MousePointer = 11
        
    On Error GoTo errH
     
    '三方LIS报告
    If mstr挂号单 = "" Then
        strSQL = "select b.id as 报告ID,b.报告名 as 项目, To_Char(b.创建时间,'YYYY-MM-DD HH24:MI') as 审核时间,b.报告名||','||To_Char(A.开嘱时间,'YYYY-MM-DD HH24:MI') as 文档标题,c.医嘱ID,b.类型,b.打印次数,b.是否禁止打印," & _
            " To_Char(A.开始执行时间,'YYYY-MM-DD HH24:MI') as 申请时间,a.标本部位 as 标本,a.开嘱医生 as 申请人,To_Char(d.采样时间,'YYYY-MM-DD HH24:MI') as 采集时间" & _
            " from 病人医嘱记录 a, 医嘱报告内容 b,病人医嘱报告 c,病人医嘱发送 d" & _
            " where b.id=c.报告id and a.id=c.医嘱id and c.报告id is not null and a.id=d.医嘱id(+) and b.类型 in (0,2) and a.病人id=[1] and a.主页id=[2]" & _
            " order by b.创建时间 desc"
    Else
        strSQL = "select b.id as 报告ID,b.报告名 as 项目,To_Char(b.创建时间,'YYYY-MM-DD HH24:MI') as 审核时间,b.报告名||','||To_Char(A.开嘱时间,'YYYY-MM-DD HH24:MI') as 文档标题,c.医嘱ID,b.类型,b.打印次数,b.是否禁止打印," & _
            " To_Char(A.开始执行时间,'YYYY-MM-DD HH24:MI') as 申请时间,a.标本部位 as 标本,a.开嘱医生 as 申请人,To_Char(d.采样时间,'YYYY-MM-DD HH24:MI') as 采集时间" & _
            " from 病人医嘱记录 a, 医嘱报告内容 b,病人医嘱报告 c,病人医嘱发送 d" & _
            " where b.id=c.报告id and a.id=c.医嘱id and c.报告id is not null and a.id=d.医嘱id(+) and b.类型 in (0,2) and a.挂号单=[3]" & _
            " order by b.创建时间 desc"
    End If
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
        strSQL = Replace(strSQL, "医嘱报告内容", "H医嘱报告内容")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
    End If
 
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng就诊ID, mstr挂号单)
    mstrPrePati = ""
    Call WebShow("", "报告")
    rptPati.Records.DeleteAll
    Do Until rsTmp.EOF
        Set objRecord = Me.rptPati.Records.Add()
            objRecord.Tag = CStr("_" & rsTmp!医嘱ID & "_" & rsTmp!报告ID) '用于定位唯一行
        '选择列
        Set objItem = objRecord.AddItem("")
        objItem.HasCheckbox = True
        
        '图标
        Set objItem = objRecord.AddItem("")
        If Val(rsTmp!打印次数 & "") > 0 Then
            objItem.Icon = 8
        End If
        
        '项目
        objRecord.AddItem rsTmp!项目 & ""
        
        '审核时间
        objRecord.AddItem rsTmp!审核时间 & ""
        
        '标本
        objRecord.AddItem rsTmp!标本 & ""
        
        '申请时间
        objRecord.AddItem rsTmp!申请时间 & ""
        
        '申请人
        objRecord.AddItem rsTmp!申请人 & ""
        
        '采集时间
        objRecord.AddItem rsTmp!采集时间 & ""
        
        '医嘱ID
        objRecord.AddItem rsTmp!医嘱ID & ""
        
        '报告ID
        objRecord.AddItem rsTmp!报告ID & ""
        
        '类型
        objRecord.AddItem rsTmp!类型 & ""
        
        '打印次数
        objRecord.AddItem rsTmp!打印次数 & ""
        
        '文档标题
        objRecord.AddItem rsTmp!文档标题 & ""
        
        '禁止打印
        objRecord.AddItem rsTmp!是否禁止打印 & ""
        
        rsTmp.MoveNext
    Loop
    rptPati.Populate
   
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
  
    Dim objColumn As ReportColumn
        
    If Button = 1 Then
        If rptPati.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptPati.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = COL_选择 Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        Call SelectALLPati(True)
                    Else
                        objColumn.Caption = ""
                        Call SelectALLPati(False)
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub SelectALLPati(ByVal blnSelect As Boolean)
'功能:设置全选或者全清
    Dim i As Long
    
    If rptPati.Columns(COL_选择).Visible And rptPati.SelectedRows.Count > 0 Then
        '先清除所有记录的选择状态
        For i = 0 To rptPati.Records.Count - 1
            rptPati.Records(i)(COL_选择).Checked = False
        Next
        
        For i = 0 To rptPati.Rows.Count - 1
            rptPati.Rows(i).Record(COL_选择).Checked = blnSelect
        Next
        rptPati.Redraw
    End If
End Sub

Private Sub rptPati_SelectionChanged()
'功能:切换行
    Dim strCurPati As String
    Dim strKey As String
    If rptPati.SelectedRows.Count <= 0 Then Exit Sub
    With rptPati.SelectedRows(0)
        If Not .GroupRow Then strCurPati = .Record.Tag
        If strCurPati = mstrPrePati Then Exit Sub
        mstrPrePati = strCurPati
        If Not .GroupRow Then
            mlng医嘱ID = .Record(COL_医嘱ID).value
            mlng报告ID = .Record(COL_报告ID).value
            mlng类型 = Val(.Record(COL_类型).value)
            mbln禁止打印 = Val(.Record(COL_禁止打印).value) = 1
        End If
        strKey = mlng报告ID & ";" & mlng医嘱ID & ";" & mlng类型
        Call zlControl.FormLock(Me.hwnd)
        Call WebShow(strKey, .Record(COL_文档标题).value)
        Call zlControl.FormLock(0)
    End With
End Sub

Private Sub CreatePDFobj()
'功能:创建对象
    On Error Resume Next
    mblnPDF = False
    Set mclsPDF = CreateObject("zlPDFViewer.clsPDFViewer")
    If Not mclsPDF Is Nothing Then
        Set mfrmPDF = mclsPDF.GetFrm
        If Not mfrmPDF Is Nothing Then
            mblnPDF = True
        End If
    End If
End Sub

Private Sub FunReference()
'功能：调用诊疗参考
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strIDs As String
    Dim i As Long
    
    On Error GoTo errH
    
    If 0 <> mlng医嘱ID Then
        strSQL = "select a.诊疗项目ID from 病人医嘱记录 a where a.id=[1] or a.相关id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID)
        
        For i = 1 To rsTmp.RecordCount
            strIDs = strIDs & "," & rsTmp!诊疗项目ID
            rsTmp.MoveNext
        Next
    End If
    
'    Call frmClinicHelp.ShowMe(0, Me, 0, True, Mid(strIDs, 2))
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
