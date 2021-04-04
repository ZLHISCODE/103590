VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUnitSubjectSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "病区标记设置"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9615
   Icon            =   "frmUnitSubjectSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9615
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraUnit 
      Height          =   1815
      Left            =   1800
      TabIndex        =   19
      Top             =   840
      Width           =   3615
      Begin VB.TextBox txtDays 
         Height          =   300
         Left            =   960
         MaxLength       =   3
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   960
         MaxLength       =   10
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "有效天数"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "天"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   1920
         TabIndex        =   24
         Top             =   780
         Width           =   180
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0表示永久有效"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   23
         Top             =   780
         Width           =   1170
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "标记名称"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame fraUd 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4935
      Begin XtremeReportControl.ReportControl UnitReportControl 
         Height          =   2415
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   4260
         _StockProps     =   0
      End
   End
   Begin VB.Frame fraLine 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   5280
      TabIndex        =   17
      Top             =   960
      Width           =   100
   End
   Begin VB.Frame fraInfo 
      Height          =   4575
      Left            =   5520
      TabIndex        =   3
      Top             =   960
      Width           =   3975
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   1080
         ScaleHeight     =   2265
         ScaleWidth      =   2625
         TabIndex        =   11
         Top             =   1560
         Visible         =   0   'False
         Width           =   2655
         Begin VB.PictureBox pic标记 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1335
            Left            =   360
            ScaleHeight     =   1335
            ScaleWidth      =   1335
            TabIndex        =   12
            Top             =   120
            Width           =   1335
            Begin VB.PictureBox picIcon 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Index           =   0
               Left            =   120
               ScaleHeight     =   615
               ScaleWidth      =   615
               TabIndex        =   16
               Top             =   120
               Width           =   615
               Begin VB.Image imgICon 
                  Height          =   360
                  Index           =   0
                  Left            =   120
                  Picture         =   "frmUnitSubjectSet.frx":08CA
                  Top             =   0
                  Width           =   360
               End
               Begin VB.Label lblInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   0
                  Left            =   120
                  TabIndex        =   13
                  Top             =   480
                  Width           =   90
               End
               Begin VB.Label lblSelect 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   360
                  Index           =   0
                  Left            =   120
                  TabIndex        =   15
                  Top             =   120
                  Width           =   300
               End
            End
         End
         Begin VB.VScrollBar HScr 
            Height          =   2295
            LargeChange     =   50
            Left            =   2400
            Max             =   100
            SmallChange     =   100
            TabIndex        =   14
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdImage 
         Appearance      =   0  'Flat
         Caption         =   "&P"
         Height          =   300
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(F4)"
         Top             =   720
         Width           =   270
      End
      Begin VB.ComboBox cbo标记 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   1905
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1200
         Width           =   1935
      End
      Begin MSComctlLib.ImageCombo imaCustom 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   720
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "标记说明"
         Height          =   180
         Index           =   8
         Left            =   240
         TabIndex        =   9
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "个性标记"
         Height          =   180
         Index           =   9
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "标记图形"
         Height          =   180
         Index           =   7
         Left            =   240
         TabIndex        =   6
         Top             =   780
         Width           =   720
      End
   End
   Begin VB.ComboBox cboUnit 
      Height          =   300
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1905
   End
   Begin MSComctlLib.ImageList Img标记 
      Index           =   999
      Left            =   3840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   43
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":0FCC
            Key             =   "监护仪"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":131E
            Key             =   "等待审查"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":1670
            Key             =   "拒绝审查"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":19C2
            Key             =   "正在抽查"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":1D14
            Key             =   "正在审查"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":2066
            Key             =   "抽查反馈"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":23B8
            Key             =   "审查反馈"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":270A
            Key             =   "抽查整改"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":2A5C
            Key             =   "审查整改"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":2DAE
            Key             =   "未导入"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":3100
            Key             =   "执行中"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":3452
            Key             =   "不符合"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":37A4
            Key             =   "正常结束"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":3AF6
            Key             =   "变异结束"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":3E48
            Key             =   "预转科"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":419A
            Key             =   "预出院"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":44EC
            Key             =   "刀"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":483E
            Key             =   "男孩"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":4B90
            Key             =   "女孩"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":4EE2
            Key             =   "男人"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":5234
            Key             =   "女人"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":5586
            Key             =   "药"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":58D8
            Key             =   "针"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":5C2A
            Key             =   "盾牌"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":5F7C
            Key             =   "铅笔"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":62CE
            Key             =   "曲别针"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":6620
            Key             =   "体温计"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":6972
            Key             =   "准备"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":6CC4
            Key             =   "停止"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":7016
            Key             =   "正确"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":7368
            Key             =   "PDA"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":76BA
            Key             =   "灯泡"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":7A0C
            Key             =   "提醒"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":7D5E
            Key             =   "红旗"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":80B0
            Key             =   "禁止"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":8402
            Key             =   "手机"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":8754
            Key             =   "刷子"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":8AA6
            Key             =   "锁"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":8DF8
            Key             =   "确认"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":914A
            Key             =   "疑问"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":949C
            Key             =   "五角星"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":97EE
            Key             =   "胸花"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":9B40
            Key             =   "病床"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Img标记 
      Index           =   1
      Left            =   3120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   43
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":9E92
            Key             =   "监护仪"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":A5A4
            Key             =   "等待审查"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":ACB6
            Key             =   "拒绝审查"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":B3C8
            Key             =   "正在抽查"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":BADA
            Key             =   "正在审查"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":C1EC
            Key             =   "抽查反馈"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":C8FE
            Key             =   "审查反馈"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":D010
            Key             =   "抽查整改"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":D722
            Key             =   "审查整改"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":DE34
            Key             =   "未导入"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":E546
            Key             =   "执行中"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":EC58
            Key             =   "不符合"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":F36A
            Key             =   "正常结束"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":FA7C
            Key             =   "变异结束"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":1018E
            Key             =   "预转科"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":108A0
            Key             =   "预出院"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":10FB2
            Key             =   "刀"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":116C4
            Key             =   "男孩"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":11DD6
            Key             =   "女孩"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":124E8
            Key             =   "男人"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":12BFA
            Key             =   "女人"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":1330C
            Key             =   "药"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":13A1E
            Key             =   "针"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":14130
            Key             =   "盾牌"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":14842
            Key             =   "铅笔"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":14F54
            Key             =   "曲别针"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":15666
            Key             =   "体温计"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":15D78
            Key             =   "准备"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":1648A
            Key             =   "停止"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":16B9C
            Key             =   "正确"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":172AE
            Key             =   "PDA"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":179C0
            Key             =   "灯泡"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":180D2
            Key             =   "提醒"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":187E4
            Key             =   "红旗"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":18EF6
            Key             =   "禁止"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":19608
            Key             =   "手机"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":19D1A
            Key             =   "刷子"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":1A42C
            Key             =   "锁"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":1AB3E
            Key             =   "确认"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":1B250
            Key             =   "疑问"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":1B962
            Key             =   "五角星"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":1C074
            Key             =   "胸花"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnitSubjectSet.frx":1C786
            Key             =   "病床"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
      Height          =   420
      Left            =   240
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
      _cx             =   1508
      _cy             =   741
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      TabIndex        =   26
      Top             =   6030
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmUnitSubjectSet.frx":1CE98
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14526
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
   Begin XtremeCommandBars.ImageManager ImgMain 
      Left            =   7080
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmUnitSubjectSet.frx":1D72A
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   1680
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmUnitSubjectSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const COL_NULL = 0
Const COL_标注 = 1
Const COL_说明 = 2
Const COL_主题序号 = 3
Const COL_有效天数 = 4
Const COL_原始主题 = 5
Const COL_原始标记 = 6
Const COL_主题说明 = 7
  
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private mRect As RECT

Private Type TYPE_UNIT
    病区ID  As Long
    主题序号 As Long
    标记序号 As Long
    说明 As String
    图形索引 As Long
    有效天数 As Long
    原始主题 As Long
    原始标记 As Long
End Type

Private mUnit As TYPE_UNIT

Const Enable_Color = &HE0E0E0
Const UnEnable_Color = &H80000005

Private mblnChange As Boolean '记录标记内容变动
Private mstrSubject As String '标记分类名称
Private mlngDay As Long '标记分类天数
Private mlngCount As Long  '存放标记分类数目

Public mstrPrivs As String
Private mstrUnits As String
Private m病区ID As Long
Private mstr病区名称 As String

Private mcbrToolBars As CommandBar  '工具栏
Private mcbrMenuBars As CommandBarControl
Const mlngImgIndex As Long = 16 '定义图片索引从第几个开始显示

Private mblnOK As Boolean
Private mrsData As New ADODB.Recordset

Public Function ShowMe(ByVal frmParent As Form, ByVal lngUnitID As Long, ByVal strPrivs As String) As Boolean
    m病区ID = lngUnitID
    mstrPrivs = strPrivs
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cboUnit_Click()
    If cboUnit.ListCount > 0 And m病区ID <> Val(cboUnit.ItemData(cboUnit.ListIndex)) Then
        m病区ID = Val(cboUnit.ItemData(cboUnit.ListIndex))
        mstr病区名称 = cboUnit.Text
    
        Call RefreshData
    End If
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Call Cbo.MatchIndex(cboUnit.hwnd, KeyAscii)
End Sub

Private Sub cbo标记_Click()
'-------------------------------------------------
'功能:根据选择主题序号改变标记内容位置
'-------------------------------------------------
    Dim strTag As String
    Dim lngPreID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngRowIndex As Long, lngRow As Long, lngOldID As Long
    Dim strFileds As String, strValues As String
    Dim str标记 As String, strCaption As String
    Dim intDay As Integer
    
    If UnitReportControl.Records.Count = 0 Then Exit Sub
    If cbo标记.ListIndex = -1 Or fraInfo.Tag = "新增" Or mblnChange = False Then Exit Sub
    If UnitReportControl.FocusedRow.GroupRow And UnitReportControl.FocusedRow.Childs.Count <> 0 Then Exit Sub
    If mrsData Is Nothing Then Exit Sub
    
    strFileds = "主题序号," & adDouble & ",18|标记序号," & adDouble & ",18|说明," & adLongVarChar & ",100|图形索引," & _
        adDouble & ",18|有效天数," & adDouble & ",18,|原始主题序号," & adDouble & ",18|原始标记序号," & adDouble & ",18"
    Call Record_Init(rsTemp, strFileds)
    'A.主题序号,A.标记序号,A.说明,A.图形索引,A.有效天数,A.主题序号 原始主题序号,A.标记序号 原始标记序号
    strFileds = "主题序号|标记序号|说明|图形索引|有效天数|原始主题序号|原始标记序号"
    
    lngRowIndex = UnitReportControl.FocusedRow.Index
    
    str标记 = ""
    mrsData.Filter = ""
    For lngRow = 0 To UnitReportControl.Rows.Count - 1
        If Not UnitReportControl.Rows(lngRow).GroupRow Then
            lngOldID = Val(Split(UnitReportControl.Rows(lngRow).Record(COL_主题序号).Record.Tag, "-")(0))
            mrsData.Filter = "主题序号=" & lngOldID & " and 标记序号=0"
            If mrsData.RecordCount > 0 Then
                strCaption = Nvl(mrsData!说明)
                intDay = Val(Nvl(mrsData!有效天数))
            End If
            
            If UnitReportControl.Rows(lngRow).Index = lngRowIndex Then
                mUnit.主题序号 = Val(cbo标记.ItemData(cbo标记.ListIndex))
                lngPreID = AgainComputePreId(Val(cbo标记.ItemData(cbo标记.ListIndex))) '获取标记序号
                mUnit.标记序号 = lngPreID
                
                mrsData.Filter = "主题序号=" & mUnit.主题序号 & " and 标记序号=0"
                If mrsData.RecordCount > 0 Then mUnit.有效天数 = Val(Nvl(mrsData!有效天数))
                str标记 = mUnit.主题序号 & "-" & mUnit.标记序号 & "-" & m病区ID & "-" & mUnit.有效天数
            Else
                mUnit.主题序号 = Val(Split(UnitReportControl.Rows(lngRow).Record(COL_主题序号).Record.Tag, "-")(0))
                mUnit.标记序号 = Val(Split(UnitReportControl.Rows(lngRow).Record(COL_主题序号).Record.Tag, "-")(1))
                mUnit.有效天数 = intDay ' Val(Nvl(UnitReportControl.Rows(lngRow).Record(COL_有效天数).Value, 0))
            End If
                        
            mUnit.说明 = Nvl(UnitReportControl.Rows(lngRow).Record(COL_说明).Value)
            mUnit.图形索引 = Val(Nvl(UnitReportControl.Rows(lngRow).Record(COL_标注).Icon, 0))
            mUnit.原始主题 = Nvl(UnitReportControl.Rows(lngRow).Record(COL_原始主题).Value, 0)
            mUnit.原始标记 = Nvl(UnitReportControl.Rows(lngRow).Record(COL_原始标记).Value, 0)
            
            '检查主题序号是否存在 不存在就添加
            rsTemp.Filter = "主题序号=" & lngOldID & " and 标记序号=0"
            If rsTemp.RecordCount = 0 Then
                strValues = lngOldID & "|" & 0 & "|" & strCaption & "|0|" & _
                    intDay & "|" & mUnit.原始主题 & "|" & mUnit.原始标记
                Call Rec.AddNew(rsTemp, strFileds, strValues)
            End If
            If Val(Split(UnitReportControl.Rows(lngRow).Record(COL_主题序号).Record.Tag, "-")(1)) <> 0 Then
                strValues = mUnit.主题序号 & "|" & mUnit.标记序号 & "|" & mUnit.说明 & "|" & mUnit.图形索引 & "|" & _
                    mUnit.有效天数 & "|" & mUnit.原始主题 & "|" & mUnit.原始标记
                Call Rec.AddNew(rsTemp, strFileds, strValues)
            End If
        End If
    Next lngRow

    rsTemp.Filter = 0
    rsTemp.Sort = "主题序号,标记序号"
    'Call OutputRsData(rsTemp)
    Call RefreshData(0, str标记, rsTemp)
    mblnChange = True
'    With UnitReportControl.FocusedRow.Record(COL_主题序号)
'        .GroupCaption = "分组：" & cbo标记.ItemData(cbo标记.ListIndex) & "-" & cbo标记.Text
'        strTag = .Record.Tag
'        lngPreID = AgainComputePreId(Val(cbo标记.ItemData(cbo标记.ListIndex))) '获取标记序号
'        .Record.Tag = cbo标记.ItemData(cbo标记.ListIndex) & "-" & lngPreID & "-" & Split(strTag, "-")(2)
'    End With
'
'    UnitReportControl.Populate

End Sub

Private Sub cbo标记_KeyPress(KeyAscii As Integer)
    Call Cbo.MatchIndex(cbo标记.hwnd, KeyAscii)
End Sub

Private Sub cbsMain_Resize()
    Call ResizeState
End Sub

Private Sub cmdImage_Click()
'功能显示现有图片信息
    Dim i As Integer, j As Integer
    Dim lngCurXCount As Long
    Dim lngH As Integer, lngW As Integer '记录picture的高度和宽度
    Dim lngX1 As Long 'pictrue之间的间隔
    Dim lngX As Long, lngY As Long  '设定image的顶部和左侧边距
    Dim lngIndex As Long
    Dim vRect As RECT
    Dim vRect1 As RECT
    
    
    lngIndex = 0
    lngY = 60
    lngX = 60

    imgIcon(lngIndex).Top = lngY
    imgIcon(lngIndex).Left = lngX
    
    lblSelect(lngIndex).Top = lngY / 2
    lblSelect(lngIndex).Left = lngX / 2
    lblSelect(lngIndex).Width = imgIcon(lngIndex).Width + lngX
    lblSelect(lngIndex).Height = imgIcon(lngIndex).Height + lngY
    
    lblInfo(lngIndex).FontSize = 8
    lblInfo(lngIndex).Top = lngY + imgIcon(lngIndex).Width + lngY / 2
    lblInfo(lngIndex).Caption = Img标记(1).ListImages(mlngImgIndex + 1).Key
    
    picIcon(lngIndex).Top = 0
    picIcon(lngIndex).Left = 0
    picIcon(lngIndex).Height = imgIcon(lngIndex).Height + lngY + lngY / 2 + lblInfo(lngIndex).Height + 10
    picIcon(lngIndex).Width = imgIcon(lngIndex).Width + imgIcon(lngIndex).Left * 2 + lngX / 2
    
    lngH = picIcon(lngIndex).Height
    lngW = picIcon(lngIndex).Width
    
    lblInfo(lngIndex).Left = (lngW - lblInfo(lngIndex).Width) / 2
    
    '获取计算picback的位置的宽度
    vRect = zlControl.GetControlRect(imaCustom.hwnd)
    vRect1 = zlControl.GetControlRect(fraInfo.hwnd)
    picBack.Top = vRect.Bottom - vRect1.Top
    picBack.Left = vRect.Left - vRect1.Left
    picBack.Width = vRect1.Right - vRect.Left - 10
    
    pic标记.Width = picBack.ScaleWidth - HScr.Width
    
    '计算每行可存放的图片数量
    lngCurXCount = (pic标记.Width - HScr.Width) \ lngW
    '重新计算位置
    lngX1 = (pic标记.Width - HScr.Width - (lngW * lngCurXCount)) / (lngCurXCount + 1)
    picIcon(lngIndex).Left = lngX1
    
    imgIcon(lngIndex).Picture = Img标记(1).ListImages(mlngImgIndex + 1).Picture
    
    HScr.Top = 0
    HScr.Min = 0
    HScr.Left = pic标记.Width
    HScr.Value = 0
    HScr.Height = picBack.ScaleHeight
    
    picBack.Visible = True
    pic标记.Visible = True
    pic标记.Top = 0
    pic标记.Left = 0
    pic标记.SetFocus
    
    i = 1
    For j = mlngImgIndex + 1 To Img标记(1).ListImages.Count - 1
        Load picIcon(i)
        picIcon(i).Visible = True
        
        If i < lngCurXCount Then
            picIcon(i).Top = 0
            picIcon(i).Left = lngW * i + (i + 1) * lngX1
        Else
            picIcon(i).Top = lngH * ((i \ lngCurXCount))
            picIcon(i).Left = lngW * (i Mod lngCurXCount) + ((i Mod lngCurXCount) + 1) * lngX1
        End If
        
        picIcon(i).Width = picIcon(lngIndex).Width
        picIcon(i).Height = picIcon(lngIndex).Height
        
        '加载图片信息
        Load imgIcon(i)
        imgIcon(i).Visible = True
        Set imgIcon(i).Container = picIcon(i)
        imgIcon(i).Picture = Img标记(1).ListImages(j + 1).Picture
        imgIcon(i).Top = imgIcon(lngIndex).Top
        imgIcon(i).Left = imgIcon(lngIndex).Left

        '加载选择控件
        Load lblSelect(i)
        lblSelect(i).Visible = True
        Set lblSelect(i).Container = picIcon(i)
        lblSelect(i).Top = lblSelect(lngIndex).Top
        lblSelect(i).Left = lblSelect(lngIndex).Left
        lblSelect(i).Width = lblSelect(lngIndex).Width
        lblSelect(i).Height = lblSelect(lngIndex).Height
        
        '加载图片说明
        Load lblInfo(i)
        lblInfo(i).Visible = True
        Set lblInfo(i).Container = picIcon(i)
        lblInfo(i).FontSize = lblInfo(lngIndex).FontSize
        lblInfo(i).Top = lblInfo(lngIndex).Top
        lblInfo(i).Caption = Img标记(1).ListImages(j + 1).Key
        lblInfo(i).Left = (lngW - lblInfo(i).Width) / 2
        
        i = i + 1
    Next j
    
    pic标记.Height = picIcon(i - 1).Top + picIcon(i - 1).Height
    pic标记.Refresh
    
    If pic标记.ScaleHeight - picBack.ScaleHeight <= 0 Then
        HScr.Max = 0
        HScr.Min = 0
    Else
        HScr.Max = pic标记.ScaleHeight - picBack.ScaleHeight
    End If
    cmdImage.Enabled = False
End Sub

Private Function GetMarkCount() As Long
    '获取标记项目总数
    Dim lngRow As Long
    Dim lngCount As Long
    
    For lngRow = 0 To UnitReportControl.Rows.Count - 1
        '标记序号=0的为标记主题分类，不进行统计
        If Not UnitReportControl.Rows(lngRow).GroupRow And UnitReportControl.Rows(lngRow).Childs.Count = 0 Then
            If Val(Split(UnitReportControl.Rows(lngRow).Record(COL_主题序号).Record.Tag, "-")(1)) <> 0 Then
                lngCount = lngCount + 1
            End If
        End If
    Next lngRow
    
    GetMarkCount = lngCount
End Function

Private Sub RefreshStateInfo()
'------------------------------------------------------------------------------------------------------------------
'功能：刷新状态栏显示信息
'-----------------------------------------------------------------------------------------------------------------
    stbThis.Panels(2).Text = "共有 " & GetMarkCount & " 个标记内容！"
End Sub

Private Sub UnLoadImage()
'功能:卸载pic标注上的所有控件
    Dim i As Integer
    For i = picIcon.Count - 1 To 1 Step -1
        Unload imgIcon(i)
        Unload lblInfo(i)
        Unload lblSelect(i)
        Unload picIcon(i)
    Next i
    picBack.Visible = False
    cmdImage.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 39 Then KeyCode = 0
    
    If KeyCode = 27 Then
        Call UnLoadImage
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ZLCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_Load()
    gblnOK = False
    Call RestoreWinState(Me, App.ProductName)
    '加载菜单工具栏
    Call InitCommandBar
    '提取病区信息
    Call InitUnits
    '加载主题标致信息
    Call InitReportControl
    '读取数据
    Call RefreshData
End Sub

Private Sub AddImage()
'------------------------------------
'功能:加载所有图片信息到ImageCombo
'------------------------------------
    Dim objNewItem As ComboItem
    Dim i As Long
 
    imaCustom.ImageList = Img标记(999)
    For i = 1 To Img标记(999).ListImages.Count - mlngImgIndex
        Set objNewItem = imaCustom.ComboItems.Add(i, "A" & i, Img标记(999).ListImages(mlngImgIndex + i).Key, mlngImgIndex + i)
    Next i
    
End Sub

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If UnitReportControl.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(vsfPrint, UnitReportControl) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    Set objPrint.Body = vsfPrint
    
    objPrint.Title.Text = "病区标记内容清单"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub InitCommandBar()
'功能:初始化菜单栏
    Dim cbrTools As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    
    On Error GoTo ErrHand
    
    Set cbsMain.Icons = ZLCommFun.GetPubIcons
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .ShowTextBelowIcons = False
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .UseSharedImageList = False '显示图形
    End With
    
        '菜单定义
    cbsMain.ActiveMenuBar.Title = "菜单栏"
    cbsMain.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set mcbrMenuBars = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBars.ID = conMenu_FilePopup
    With mcbrMenuBars.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "取消(&Z)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        cbrControl.BeginGroup = True
    End With

    Set mcbrMenuBars = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBars.ID = conMenu_EditPopup
    With mcbrMenuBars.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewParent, "新增分类(&I)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "修改分类(&U) ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_DeleteParent, "删除分类(&E)")
    
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
    End With

    Set mcbrMenuBars = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBars.ID = conMenu_ViewPopup
    With mcbrMenuBars.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBars = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBars.ID = conMenu_HelpPopup
    With mcbrMenuBars.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."): cbrControl.BeginGroup = True
    End With
    
     '快键绑定
    With cbsMain.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Reuse
        .Add FSHIFT, VK_INSERT, conMenu_Edit_NewParent
        .Add FSHIFT, VK_DELETE, conMenu_Edit_DeleteParent
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '--添加工具栏
    Set mcbrToolBars = cbsMain.Add("工具栏", xtpBarTop)
    mcbrToolBars.EnableDocking xtpFlagStretched
    With mcbrToolBars.Controls
        Set cbrTools = .Add(xtpControlPopup, conMenu_Edit_FileMan, "分类", -1, False)
        cbrTools.IconId = conMenu_Edit_FileMan
        cbrTools.ToolTipText = "标记分类"
        cbrTools.BeginGroup = True
        
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_NewParent, "新增"
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_ModifyParent, "修改"
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_DeleteParent, "删除"
        
        Set cbrTools = .Add(xtpControlPopup, conMenu_Edit_Leave_Add, "标记", -1, False)
        cbrTools.IconId = conMenu_Edit_NewItem
        cbrTools.ToolTipText = "标记内容"
        
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_NewItem, "新增"
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_Modify, "修改"
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_Delete, "删除"
        

        Set cbrTools = .Add(xtpControlButton, conMenu_Edit_Save, "保存")
        cbrTools.ToolTipText = "保存"
        cbrTools.BeginGroup = True
        
        Set cbrTools = .Add(xtpControlButton, conMenu_Edit_Reuse, "取消")
        cbrTools.ToolTipText = "取消"

        Set cbrTools = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        cbrTools.ToolTipText = "帮助"
        cbrTools.BeginGroup = True
        Set cbrTools = .Add(xtpControlButton, conMenu_File_Exit, "退出")

    End With
    
    For Each cbrControl In mcbrToolBars.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '工具栏右侧病区下拉框选择
    With mcbrToolBars.Controls
        Set objControl = .Add(xtpControlLabel, conMenu_View_Find, "病区")
        objControl.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "病区")
        objCustom.Handle = Me.cboUnit.hwnd
        objCustom.Flags = xtpFlagRightAlign
        objControl.IconId = conMenu_View_Find
    End With
    
    '加载图片信息
    Call AddImage
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitReportControl()
'功能:初始化ReportControl

    Dim Column As ReportColumn
    
    With UnitReportControl
   
    Set Column = .Columns.Add(COL_NULL, " ", 10, False)
    Column.Editable = False: Column.Groupable = False: Column.Sortable = False: Column.Alignment = xtpAlignmentCenter
    Set Column = .Columns.Add(COL_标注, "标注", 50, True)
    Column.Editable = False: Column.Groupable = False: Column.AllowDrag = False
    
    Set Column = .Columns.Add(COL_说明, "说明", 190, True)
    Column.AllowDrag = False: Column.Editable = False: Column.Groupable = False
    Set Column = .Columns.Add(COL_主题序号, "主题序号", 0, False)
    Column.Visible = False: Column.Editable = False: Column.Groupable = True
    Set Column = .Columns.Add(COL_有效天数, "有效天数", 60, True)
    Column.AllowDrag = False: Column.Editable = False: Column.Groupable = False
    Set Column = .Columns.Add(COL_原始主题, "原始主题", 0, False)
    Column.Visible = False: Column.Editable = False: Column.Groupable = False
    Set Column = .Columns.Add(COL_原始标记, "原始标记", 0, False)
    Column.Visible = False: Column.Editable = False: Column.Groupable = False
    Set Column = .Columns.Add(COL_主题说明, "主题说明", 0, False)
    Column.Visible = False: Column.Editable = False: Column.Groupable = False
    
    With .PaintManager
        .ColumnStyle = xtpColumnFlat
        .MaxPreviewLines = 1
        .GroupForeColor = &HC00000
        .GridLineColor = RGB(225, 225, 225)
        .VerticalGridStyle = xtpGridSolid
        .ShadeGroupHeadings = False
        .NoItemsText = "没有可显示的标记分类和标记内容信息..."
    End With
    
    .AllowColumnResize = False
    .ShowItemsInGroups = False '是否按排序自分理处分组
    .PreviewMode = True
    .MultipleSelection = False '会引发SelectionChanged事件
    .SetImageList Me.Img标记(999)
        
    .GroupsOrder.Add .Columns(COL_主题序号)
    .GroupsOrder(0).SortAscending = True
    .GroupsOrder(0).Groupable = True
    
    '分组之后可能失去记录集中的顺序,因此强行加入排序列
    .SortOrder.Add .Columns(COL_说明)
    .SortOrder(0).SortAscending = True
    .SortOrder.Add .Columns(COL_主题序号)
    .SortOrder(1).SortAscending = True
    End With
End Sub

Private Function RefreshData(Optional lngPreIdx As Long, Optional str标记 As String = "", Optional ByVal rsTemp As ADODB.Recordset) As Boolean
'-------------------------------------------------------------
'功能:提取病区个性化设置
'参数:lngPreIdx 选择行索引,str标记 选择行信息（用来快速定位）
'说明 lngPreIdx=-1时不进行病区标记分类检查
'-------------------------------------------------------------
    Dim strUnit As String, strInfo As String, strDay As String, strOldUnit As String
    Dim lngImgIndex As Long
    Dim blnDouble As Boolean
    Dim lngIndex As Long '存放当前序号
    Dim blnRead As Boolean
    Dim strSQL As String
    'Dim rsTemp As New ADODB.Recordset
    Dim strSubject As String '存放标记分类的信息
    Dim objRow As ReportRow, i As Long
    Dim strFileds As String, strValues As String
    
    mblnChange = False
    Screen.MousePointer = 11
    On Error GoTo ErrHand
    
    mlngCount = CheckUnitSubject(m病区ID)
    
    If rsTemp Is Nothing Then blnRead = True
    If blnRead = False Then
        If rsTemp.State = adStateClosed Then blnRead = True
    End If
    If blnRead = True Then
        
        strFileds = "主题序号," & adDouble & ",18|标记序号," & adDouble & ",18|说明," & adLongVarChar & ",100|图形索引," & _
            adDouble & ",18|有效天数," & adDouble & ",18,|原始主题序号," & adDouble & ",18|原始标记序号," & adDouble & ",18"
        Call Record_Init(mrsData, strFileds)
        strFileds = "主题序号|标记序号|说明|图形索引|有效天数|原始主题序号|原始标记序号"
         '提取病区信息
        strSQL = _
            " SELECT A.主题序号,A.标记序号,A.说明,A.图形索引,A.有效天数,A.主题序号 原始主题序号,A.标记序号 原始标记序号" & vbNewLine & _
            " FROM 病区标记内容 A,病区标记内容 B" & vbNewLine & _
            " WHERE  A.病区ID=B.病区ID And A.主题序号=B.主题序号 And B.标记序号=0  And A.病区ID=[1]" & vbNewLine & _
            " ORDER BY A.主题序号,A.标记序号"
                
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取病区主题信息", m病区ID)
    End If
    
    UnitReportControl.Records.DeleteAll
    
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    With rsTemp
        Do While Not .EOF
            If Nvl(!标记序号) = 0 Then
                If strSubject <> "" Then
                    strUnit = strSubject
                    strInfo = "此分类下没有可显示的标记内容信息..."
                    lngImgIndex = 0
                    AddRecord strUnit, lngImgIndex, strInfo, mlngDay, strOldUnit
                    strSubject = ""
                End If
                mstrSubject = Nvl(!说明, "个性标注" & Nvl(!主题序号))
                mlngDay = Val(Nvl(!有效天数, 0))
                strSubject = Nvl(!主题序号) & "-" & Nvl(!标记序号) & "-" & m病区ID
                strOldUnit = Nvl(!原始主题序号) & "-" & Nvl(!原始标记序号) & "-" & m病区ID
            Else
                strUnit = Nvl(!主题序号) & "-" & Nvl(!标记序号) & "-" & m病区ID
                strOldUnit = Nvl(!原始主题序号) & "-" & Nvl(!原始标记序号) & "-" & m病区ID
                strInfo = Nvl(!说明)
                strDay = Nvl(!有效天数, 0)
                lngImgIndex = Nvl(!图形索引, 0)
                AddRecord strUnit, lngImgIndex, strInfo, mlngDay, strOldUnit
                strSubject = ""
            End If
            If blnRead = True Then
                strValues = Val(Nvl(!主题序号)) & "|" & Val(Nvl(!标记序号)) & "|" & Nvl(!说明) & "|" & Val(Nvl(!图形索引)) & "|" & _
                   Val(Nvl(!有效天数)) & "|" & Val(Nvl(!原始主题序号)) & "|" & Val(Nvl(!原始标记序号))
                Call Rec.AddNew(mrsData, strFileds, strValues)
            End If
        .MoveNext
        Loop
    End With
    
    If strSubject <> "" Then
        strUnit = strSubject
        strInfo = "此分类下没有可显示的标记内容信息..."
        lngImgIndex = 0
        AddRecord strUnit, lngImgIndex, strInfo, mlngDay, strOldUnit
        strSubject = ""
    End If
    
    UnitReportControl.Populate
    
    If UnitReportControl.Rows.Count <> 0 Then
        Call UnitRefresh(lngPreIdx, str标记)
    Else
        Call SetFraResize(True)
        txtName.Enabled = False
        txtName.Text = ""
        txtDays.Enabled = False
        txtDays.Text = ""
        txtName.BackColor = Enable_Color
        txtDays.BackColor = Enable_Color
    End If
    
    Call RefreshStateInfo
    
    '检查是否设置病区标记分类(-1不进行提示)
    If lngPreIdx <> -1 Then
        If mlngCount = 0 Then
            'MsgBox "病区【" & Split(mstr病区名称, "-")(1) & "】还未设置病区标记分类,请添加.", vbInformation, gstrSysName
        End If
    End If
    
    Screen.MousePointer = 0
    RefreshData = True
    
    Exit Function
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
        Call SaveErrLog
    End If
End Function


Private Function UnitRefresh(Optional lngPreIdx As Long, Optional str标记 As String = "") As Boolean
'-----------------------------------------------
'功能:标记项目新增，修改后定位到选择的记录
'参数:lngreIdx 上次选择列的索引
'     str标记 上次选择列的内容 格式:主题序号-标记序号-病区ID
'-----------------------------------------------
    Dim objRow As ReportRow, i As Long, j As Long
    Dim blnRetrun As Boolean, blnChild As Boolean
    Dim arrCode() As String
    Dim lngRow As Long, lngGroup As Long
    
    If lngPreIdx < 0 Then lngPreIdx = 0
    
    If str标记 <> "" Then
        
        str标记 = str标记 & String(3 - UBound(Split(str标记, "-")), "-")
        arrCode = Split(str标记, "-")
        blnChild = Val(arrCode(1)) <> 0
        
        If blnChild = True Then
            If GetMarkCount = 0 Then blnChild = False
        End If
        
        If blnChild = True Then
            '先快速定位
            If lngPreIdx <= UnitReportControl.Rows.Count - 1 Then
                If Not UnitReportControl.Rows(lngPreIdx).GroupRow And UnitReportControl.Rows(lngPreIdx).Childs.Count = 0 Then
                    If UnitReportControl.Rows(lngPreIdx).Record(COL_主题序号).Record.Tag = str标记 Then
                        Set objRow = UnitReportControl.Rows(lngPreIdx)
                    End If
                End If
            End If
            '再进行查找
            If objRow Is Nothing Then
                For i = 0 To UnitReportControl.Rows.Count - 1
                    If Not UnitReportControl.Rows(i).GroupRow And UnitReportControl.Rows(i).Childs.Count = 0 Then
                        If UnitReportControl.Rows(i).Record(COL_主题序号).Record.Tag = str标记 Then
                            Set objRow = UnitReportControl.Rows(i): Exit For
                        End If
                    End If
                Next
            End If
        Else
            For i = 0 To UnitReportControl.Rows.Count - 1
                   If UnitReportControl.Rows(i).GroupRow And UnitReportControl.Rows(i).Childs.Count > 0 Then
                        If Split(UnitReportControl.Rows(i).Childs(0).Record(COL_主题序号).Record.Tag, "-")(0) = arrCode(0) And arrCode(1) = 0 Then
                            Set objRow = UnitReportControl.Rows(i): Exit For
                        End If
                   End If
            Next i
        End If
    End If
    
    '取第一个非分组行
    If objRow Is Nothing Then
        For i = 0 To UnitReportControl.Rows.Count - 1
            If blnChild Then
                If Not UnitReportControl.Rows(i).GroupRow And UnitReportControl.Rows(i).Childs.Count = 0 Then
                    If Val(Split(UnitReportControl.Rows(i).Record(COL_主题序号).Record.Tag, "-")(1)) <> 0 Then
                        Set objRow = UnitReportControl.Rows(i): Exit For
                    End If
                End If
            Else
                Set objRow = UnitReportControl.Rows(i)
                If objRow.GroupRow And objRow.Childs.Count > 0 Then
                    For j = 0 To objRow.Childs.Count - 1
                        If Val(Split(objRow.Childs(j).Record(COL_主题序号).Record.Tag, "-")(1)) <> 0 Then
                            Set objRow = UnitReportControl.Rows(i + 1)
                            Exit For
                        End If
                    Next j
                End If
                Exit For
            End If
        Next
    End If
    
    If Not objRow Is Nothing Then
        blnRetrun = True
        If Not objRow.GroupRow Then
            If Val(Split(objRow.Record(COL_主题序号).Record.Tag, "-")(1)) = 0 Then
                Set objRow = UnitReportControl.Rows(objRow.Index - 1)
            End If
        End If
        Set UnitReportControl.FocusedRow = objRow '该行选中且显示在可见区域,并引发SelectionChanged事件
        UnitReportControl.FocusedRow.Selected = True
        
    End If
    
    UnitRefresh = blnRetrun
End Function

Private Function AddRecord(ByVal strUnit As String, ByVal lngImgIndex As Long, ByVal strInfo As String, ByVal lngDay As Long, _
    Optional ByVal strUnitOld As String = "") As ReportRecord
'-------------------------------------------------------------------------------------------
'功能：向ReportRecord添加病区标记记录
'------------------------------------------------------------------------------------------
    Dim blnParent As Boolean
    Dim Record As ReportRecord
    Set Record = UnitReportControl.Records.Add()
    
    If strUnitOld = "" Then strUnitOld = strUnit
    Dim Item As ReportRecordItem
   
    blnParent = Val(Split(strUnit, "-")(1)) = 0
    
    Set Item = Record.AddItem("")
    If blnParent Then Item.BackColor = RGB(255, 255, 255)
    
    Set Item = Record.AddItem("")
    If lngImgIndex >= mlngImgIndex And lngImgIndex <= Img标记(999).ListImages.Count - 1 Then
        Item.Icon = Img标记(999).ListImages(lngImgIndex).Index
    End If
    If blnParent Then Item.BackColor = RGB(255, 255, 255)
    
    Set Item = Record.AddItem(strInfo)
    If blnParent Then Item.BackColor = RGB(255, 255, 255)
    
    Set Item = Record.AddItem(Val(Split(strUnit, "-")(0)))
    Item.GroupCaption = "分组：" & Val(Split(strUnit, "-")(0)) & "-" & mstrSubject
    '主题序号 & "-" & 标记序号 & "-" & 病区Id & "-" & "有效天数"
    Item.Record.Tag = strUnit & "-" & lngDay
    
    Set Item = Record.AddItem(IIf(blnParent, "", lngDay)) '有效天数
    If blnParent Then Item.BackColor = RGB(255, 255, 255)
    Record.AddItem CInt(Split(strUnitOld, "-")(0))  '记录原始主题序号
    Record.AddItem CInt(Split(strUnitOld, "-")(1)) '记录原始标记序号
    Record.AddItem mstrSubject
    
    Set AddRecord = Record
End Function

Private Function InitUnits() As Boolean
'功能：初始化住院护理病区
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim blnTrue As Boolean
    On Error GoTo errH
    mstrUnits = GetUser病区IDs
    '包含门观察室
    If InStr(mstrPrivs, "全院病人") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.编码"
    Else
        '求有权病区：直接所在病区+所在科室所属病区
        strSQL = _
            " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.编码,C.名称,Nvl(B.缺省,0) as 缺省" & _
            " From 病区科室对应 A,部门人员 B,部门表 C" & _
            " Where A.病区ID=C.ID And B.部门ID=A.科室ID And B.人员ID=[1]" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSQL & ") Group by ID,编码,名称 Order by 编码"
    End If

    cboUnit.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            
            If m病区ID = rsTmp!ID Then
                Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                If cboUnit.ListIndex <> -1 Then blnTrue = True
            End If
            
            If Not blnTrue Then
                If InStr(mstrPrivs, "全院病人") > 0 Then
                    If rsTmp!ID = UserInfo.部门ID Then  '直接所属优先
                        Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                    End If
                    If InStr("," & mstrUnits & ",", "," & rsTmp!ID & ",") > 0 And cboUnit.ListIndex = -1 Then
                        Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                    End If
                Else '所属缺省病区包含的可能有多个
                    If rsTmp!缺省 = 1 And cboUnit.ListIndex = -1 Then
                        Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                    End If
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then
        Call Cbo.SetIndex(cboUnit.hwnd, 0)
    End If
    
    If cboUnit.ListIndex <> -1 Then
        m病区ID = cboUnit.ItemData(cboUnit.ListIndex)
        mstr病区名称 = cboUnit.Text
    End If
    
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    Call ResizeState
End Sub

Private Sub SetControlEnable(Optional blnEnable As Boolean = False)
'------------------------------------------------------------------
'功能:设置是否可以编辑
'------------------------------------------------------------------
        Dim blnNone As Boolean
        Dim i As Integer
        cbo标记.Enabled = blnEnable
       
        cbo标记.BackColor = IIf(blnEnable = False, Enable_Color, UnEnable_Color)
        
        blnNone = IIf(fraInfo.Tag = "新增", True, False)
        
        If blnNone = False Then
            If UnitReportControl.SelectedRows.Count > 0 Then
                If Not UnitReportControl.SelectedRows(0).GroupRow And UnitReportControl.SelectedRows(0).Childs.Count = 0 Then
                    blnNone = False
                Else
                    blnNone = True
                End If
            Else
                blnNone = True
            End If
        End If
        
        If UnitReportControl.Records.Count = 0 Then
            cbo标记.ListIndex = -1
        Else
            If UnitReportControl.SelectedRows.Count > 0 Then
                If Not UnitReportControl.SelectedRows(0).GroupRow And UnitReportControl.SelectedRows(0).Childs.Count = 0 Then
                    cbo标记.ListIndex = SetCboIndex(cbo标记, Val(Split(UnitReportControl.SelectedRows(0).Record(COL_主题序号).Record.Tag, "-")(0)))
                Else
                    cbo标记.ListIndex = SetCboIndex(cbo标记, Val(Split(UnitReportControl.SelectedRows(0).Childs(0).Record(COL_主题序号).Record.Tag, "-")(0)))
                End If
            End If
        End If
        
        If blnNone = True Then lblSet(9).Tag = "": cbo标记.Tag = ""
        txtInfo.Enabled = blnEnable
        txtInfo.BackColor = IIf(blnEnable = False, Enable_Color, UnEnable_Color)
        If blnNone Then txtInfo.Text = "": lblSet(8).Tag = "":: txtInfo.Tag = ""
        imaCustom.Enabled = blnEnable
        imaCustom.Locked = True
        imaCustom.BackColor = IIf(blnEnable = False, Enable_Color, UnEnable_Color)
        If blnNone Then imaCustom.Text = "": lblSet(7).Tag = "": imaCustom.Tag = ""
        
        cmdImage.Enabled = blnEnable
        
        If blnEnable = True And fraInfo.Visible = True Then cbo标记.SetFocus
End Sub

Private Sub ResizeState()
'功能:设置窗体所有控件位置
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Dim blnGourp As Boolean
    Dim objRow As ReportRow
    Dim i As Integer
    
    If Me.WindowState = 1 Then Exit Sub
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If lngTop = 0 Then lngTop = 600
    
    mRect.Top = lngTop
    mRect.Left = lngLeft
    mRect.Right = lngRight
    mRect.Bottom = lngBottom
    
    fraUd.Top = lngTop
    fraUd.Left = 0
    fraUd.Width = ScaleWidth * 0.6
    fraUd.Height = lngBottom - lngTop
    
    UnitReportControl.Move 0, 100, fraUd.Width - 50, fraUd.Height - 150
    
    fraLine.Width = 50
    fraLine.Top = lngTop
    fraLine.Left = ScaleWidth * 0.6
    fraLine.Height = lngBottom - lngTop

    If InStr(1, ",新增,修改,", "," & fraInfo.Tag & ",") = 0 And InStr(1, ",新增,修改,", "," & fraUnit.Tag & ",") = 0 Then
        blnGourp = False
        If UnitReportControl.Rows.Count > 0 Then
            If GetMarkCount > 0 Then
                For i = 0 To UnitReportControl.Rows.Count - 1
                    If UnitReportControl.Rows(i).Selected = True Then
                        Set objRow = UnitReportControl.Rows(i)
                    End If
                Next i
                
                If Not objRow Is Nothing Then
                    If objRow.GroupRow Then
                        blnGourp = True
                    Else
                        blnGourp = False
                    End If
                Else
                    blnGourp = False
                End If
            Else
                blnGourp = True
            End If
        Else
            blnGourp = True
        End If
    ElseIf InStr(1, ",新增,修改,", "," & fraInfo.Tag & ",") = 0 Then
        blnGourp = True
    Else
        blnGourp = False
    End If
    
    Call SetFraResize(blnGourp)
End Sub

Private Sub SetFraResize(Optional blnGroup As Boolean = False)
    If blnGroup = True Then
        fraInfo.Visible = False
        fraInfo.Enabled = False
        fraUnit.Visible = True
        fraUnit.Enabled = True
        fraUnit.Top = mRect.Top
        fraUnit.Width = ScaleWidth * 0.4 - fraLine.Width
        fraUnit.Height = mRect.Bottom - mRect.Top
        fraUnit.Left = ScaleWidth * 0.6 + fraLine.Width
    Else
        fraUnit.Visible = False
        fraUnit.Enabled = False
        fraInfo.Visible = True
        fraInfo.Enabled = True
        fraInfo.Top = mRect.Top
        fraInfo.Width = ScaleWidth * 0.4 - fraLine.Width
        fraInfo.Height = mRect.Bottom - mRect.Top
        fraInfo.Left = ScaleWidth * 0.6 + fraLine.Width
    End If
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub


Private Sub Form_Unload(Cancel As Integer)
    mstrSubject = ""
    mlngDay = 0
    Call UnLoadImage
    mblnOK = (fraUd.Tag = "1")
    If Not (mrsData Is Nothing) Then Set mrsData = Nothing
'    If mblnChange = True Then
'        If MsgBox("病区【" & Split(mstr病区名称, "-")(1) & "】标记内容已经发生改变，你确定要退出吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
'    End If
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub HScr_Change()
    pic标记.Top = HScr.Top - HScr.Value
    If picBack.Visible = True Then picBack.SetFocus
End Sub

Private Sub HScr_Scroll()
    pic标记.Top = HScr.Top - HScr.Value
End Sub

Private Sub imaCustom_Click()
     Call showIcon(imaCustom.SelectedItem.Index - 1)
End Sub

Private Sub imaCustom_KeyPress(KeyAscii As Integer)
        Dim i As Integer
    If KeyAscii <> vbKeyReturn Then
        Call Cbo.MatchIndex(imaCustom.hwnd, KeyAscii)
    Else
    '由于敲回车后ImageCombo图形丢失，此处重新显示图标
    If KeyAscii = vbKeyReturn Then
        If imaCustom.Text <> "" Then
             For i = 1 To Img标记(999).ListImages.Count - mlngImgIndex
                If imaCustom.Text = Img标记(999).ListImages(mlngImgIndex + i).Key Then
                    imaCustom.ComboItems(i).Selected = True
                End If
            Next i
        End If
    End If
    End If
End Sub


Private Sub imgIcon_DblClick(Index As Integer)
    Call showIcon(Index)
End Sub

Private Sub showIcon(ByVal Index As Integer)
'功能:展示用户选择的图标
    If Index < 0 Then Exit Sub
    imaCustom.ComboItems(Index + 1).Selected = True
    Call UnLoadImage
    
    If fraInfo.Tag = "修改" Then
        With UnitReportControl.FocusedRow.Record(COL_标注)
            .Icon = Index + mlngImgIndex
        End With
        
        UnitReportControl.Populate
    End If
    
    If (txtInfo.Text = "" Or txtInfo.Tag <> "改变") And IIf(fraInfo.Tag = "修改", lblSet(8).Tag = "", True) Then txtInfo.Text = imaCustom.ComboItems(Index + 1).Text
    
End Sub

Private Sub ShowSelect(ByVal Index As Integer)
'功能:选中图标
    Dim i As Integer
    lblSelect(Index).BackColor = &H8000000D
    lblInfo(Index).BackColor = &H8000000D
    For i = 0 To Img标记(1).ListImages.Count - mlngImgIndex - 1
        If i <> Index Then
            lblSelect(i).BackColor = &H8000000E
            lblInfo(i).BackColor = &H8000000E
        End If
    Next i
End Sub

Private Sub imgIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowSelect(Index)
End Sub

Private Function AgainComputePreId(ByVal lngPreVId As Long, Optional bln新增 As Boolean = False) As Long
'--------------------------------------
'功能:计算算标记序号
'参数：lngPreVId：主题序号
'--------------------------------------
    Dim lngTmp As Long
    Dim blnTrue As Boolean
    Dim i As Integer
    For i = 0 To UnitReportControl.Records.Count - 1
        If lngPreVId = Val(Split(UnitReportControl.Records(i).Item(COL_主题序号).Record.Tag, "-")(0)) Then
            If lngTmp < Val(Split(UnitReportControl.Records(i).Item(COL_主题序号).Record.Tag, "-")(0)) Then
                lngTmp = Val(Split(UnitReportControl.Records(i).Item(COL_主题序号).Record.Tag, "-")(0))
            End If
        End If
    Next i
    
    If bln新增 = True Then
        '新增的记录直接加一
        lngTmp = lngTmp + 1
    Else
        '个性标记改变时如果和以前不同就序号直接加一，如果回复到以前则检测以前序号是否被使用，使用的话重新获取新的序号
        If Val(Split(UnitReportControl.FocusedRow.Record(COL_主题序号).Record.Tag, "-")(0)) = lngPreVId Then
            '检查原始序号是否被新增记录使用
            For i = 0 To UnitReportControl.Records.Count - 1
                If lngPreVId = Val(Split(UnitReportControl.Records(i).Item(COL_主题序号).Record.Tag, "-")(0)) Then
                    If UnitReportControl.FocusedRow.Record(COL_原始标记).Value = Val(Split(UnitReportControl.Records(i).Item(COL_主题序号).Record.Tag, "-")(1)) Then
                        blnTrue = True
                    End If
                End If
            Next i
            
            If blnTrue = True Then
                lngTmp = UnitReportControl.FocusedRow.Record(COL_原始标记).Value
            Else
                lngTmp = lngTmp + 1
            End If
        Else
            lngTmp = lngTmp + 1
        End If
    End If

    AgainComputePreId = lngTmp
    
End Function


Private Function SaveData() As Boolean
'------------------------------------------------------------------
'功能：病区标记数据保存
'------------------------------------------------------------------
    Dim lngRowIndex As Long '选择列的索引
    Dim i As Integer
    Dim Record As ReportRecord
    Dim strTemp As String, strSQL As String
    Dim blnTran As Boolean
    Dim strSQLAdd() As String
    Dim StrSQLMod() As String
    Dim strTmp1 As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    ReDim Preserve strSQLAdd(0 To 0)
    ReDim Preserve StrSQLMod(0 To 0)
    lngRowIndex = 0
    
    If InStr(1, ",新增,修改,", "," & fraInfo.Tag & ",") <> 0 Then
        If imaCustom.Text = "" Then
            MsgBox "标记图形不能为空,请选择标记图形后在进行保存操作.", vbInformation, gstrSysName
            imaCustom.SetFocus
            Exit Function
        End If
    End If
    
    If InStr(1, ",新增,修改,", "," & fraUnit.Tag & ",") <> 0 Then
        If Trim(txtName.Text) = "" Then
            MsgBox "标记名称不能为空,请检查.", vbInformation, gstrSysName
            txtName.SetFocus
            Exit Function
        End If
        
        If Not ZLCommFun.StrIsValid(txtDays.Text, 3, txtDays.hwnd, "有效天数") Then Exit Function
    End If
    
    '修改
    If fraInfo.Tag = "修改" Then
        If UnitReportControl.FocusedRow Is Nothing Then Exit Function
        
        lngRowIndex = UnitReportControl.FocusedRow.Index
        mUnit.病区ID = m病区ID
        mUnit.主题序号 = Val(Split(UnitReportControl.Rows(lngRowIndex).Record(COL_主题序号).Record.Tag, "-")(0))
        mUnit.标记序号 = Val(Split(UnitReportControl.Rows(lngRowIndex).Record(COL_主题序号).Record.Tag, "-")(1))
        mUnit.说明 = Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_说明).Value)
        mUnit.说明 = Trim(txtInfo.Text)
        mUnit.图形索引 = Val(Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_标注).Icon, 0))
        mUnit.有效天数 = Val(Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_有效天数).Value, 0))
        mUnit.原始主题 = Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_原始主题).Value, 0)
        mUnit.原始标记 = Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_原始标记).Value, 0)
        
        mrsData.Filter = "主题序号=" & Val(mUnit.主题序号) & " and 标记序号=0"
        If mrsData.RecordCount > 0 Then
            mUnit.有效天数 = Val(Nvl(mrsData!有效天数))
        End If
        
        '修改后数据无任何变化,不进行数据写入操作
        If CheckChange Then
            If mUnit.主题序号 <> mUnit.原始主题 Then '主题序号发生改变
                StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_病区标记内容_Delete(" & mUnit.病区ID & "," & mUnit.原始主题 & "," & mUnit.原始标记 & ")"
                
                mUnit.标记序号 = GetMaxPreID(mUnit.病区ID, mUnit.主题序号)
                
                strTmp1 = mUnit.主题序号 & "-" & mUnit.标记序号
                
                StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_病区标记内容_Insert(" & mUnit.病区ID & "," & mUnit.主题序号 & "," & _
                mUnit.标记序号 & ",'" & mUnit.说明 & "'," & mUnit.图形索引 & "," & mUnit.有效天数 & ")"
            Else
                strTmp1 = mUnit.主题序号 & "-" & mUnit.原始标记
                
                StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_病区标记内容_Update(" & mUnit.病区ID & "," & mUnit.主题序号 & "," & _
                    mUnit.原始标记 & ",'" & mUnit.说明 & "'," & mUnit.图形索引 & "," & mUnit.有效天数 & ")"
            End If
            
            If UBound(StrSQLMod) > 1 Then
                gcnOracle.BeginTrans
                blnTran = True
                For i = 0 To UBound(StrSQLMod)
                    If StrSQLMod(i) <> "" Then Call zlDatabase.ExecuteProcedure(StrSQLMod(i), Me.Caption)
                Next i
                gcnOracle.CommitTrans
            Else
                For i = 0 To UBound(StrSQLMod)
                    If StrSQLMod(i) <> "" Then Call zlDatabase.ExecuteProcedure(StrSQLMod(i), Me.Caption)
                Next i
            End If
            
            fraUd.Tag = "1"
        Else
            strTmp1 = mUnit.主题序号 & "-" & mUnit.原始标记
        End If
        strTemp = strTmp1 & "-" & mUnit.病区ID & "-" & Val(mUnit.有效天数)
    End If
    
    '新增
    If fraInfo.Tag = "新增" Then
        If cbo标记.ListIndex = -1 Then Exit Function
        mUnit.病区ID = m病区ID
        mUnit.主题序号 = cbo标记.ItemData(cbo标记.ListIndex)
        mUnit.标记序号 = GetMaxPreID(mUnit.病区ID, mUnit.主题序号)
        mUnit.说明 = txtInfo.Text
        mUnit.图形索引 = imaCustom.SelectedItem.Index - 1 + mlngImgIndex
        mUnit.有效天数 = 0
        
        For i = 0 To UnitReportControl.Rows.Count - 1
            If Not UnitReportControl.Rows(i).GroupRow And UnitReportControl.Rows(i).Childs.Count = 0 Then
                If Val(Split(UnitReportControl.Rows(i).Record(COL_主题序号).Record.Tag, "-")(0)) = cbo标记.ItemData(cbo标记.ListIndex) Then
                    mUnit.有效天数 = Val(Split(UnitReportControl.Rows(i).Record(COL_有效天数).Record.Tag, "-")(3))
                    Exit For
                End If
            End If
        Next i
        
        mrsData.Filter = "主题序号=" & Val(mUnit.主题序号) & " and 标记序号=0"
        If mrsData.RecordCount > 0 Then
            mUnit.有效天数 = Val(Nvl(mrsData!有效天数))
        End If
        
        strTmp1 = mUnit.主题序号 & "-" & mUnit.标记序号
        
        strSQLAdd(ReDimArray(strSQLAdd)) = "Zl_病区标记内容_Insert(" & mUnit.病区ID & "," & mUnit.主题序号 & "," & _
            mUnit.标记序号 & ",'" & mUnit.说明 & "'," & mUnit.图形索引 & "," & mUnit.有效天数 & ")"
            
        For i = 0 To UBound(strSQLAdd)
            If strSQLAdd(i) <> "" Then Call zlDatabase.ExecuteProcedure(strSQLAdd(i), Me.Caption)
        Next i
        
        strTemp = strTmp1 & "-" & mUnit.病区ID & "-" & Val(mUnit.有效天数)
        
        mstrSubject = cbo标记.Text
        Set Record = AddRecord(mUnit.主题序号 & "-" & mUnit.标记序号 & "-" & mUnit.病区ID, mUnit.图形索引, mUnit.说明, Val(mUnit.有效天数))
        fraUd.Tag = "1"
        UnitReportControl.Populate
    End If
                
    '新增主题名称
    If fraUnit.Tag = "新增" Then
        mUnit.主题序号 = GetSubjectId(cboUnit.ItemData(cboUnit.ListIndex))
        If mUnit.主题序号 = 0 Then Exit Function
        
        strSQLAdd(ReDimArray(strSQLAdd)) = "Zl_病区标记内容_Insert(" & cboUnit.ItemData(cboUnit.ListIndex) & "," & mUnit.主题序号 & "," & _
            0 & ",'" & Replace(Trim(txtName.Text), "'", "") & "'," & 0 & "," & Val(txtDays.Text) & ")"
        
        For i = 0 To UBound(strSQLAdd)
            If strSQLAdd(i) <> "" Then Call zlDatabase.ExecuteProcedure(strSQLAdd(i), Me.Caption)
        Next i
        
        strTemp = mUnit.主题序号 & "-0-" & mUnit.病区ID & "-" & Val(txtDays.Text)
        
        fraUd.Tag = "1"
    End If
    
    '修改主题名称
    If fraUnit.Tag = "修改" Then
        If UnitReportControl.Rows(UnitReportControl.Tag) Is Nothing Then Exit Function
        
        mUnit.主题序号 = Val(Split(UnitReportControl.Rows(UnitReportControl.Tag).Childs(0).Record(COL_主题序号).Record.Tag, "-")(0))
        mUnit.病区ID = cboUnit.ItemData(cboUnit.ListIndex)
        
        '标记分类发生变化则进行修改操作
        If CheckChange Then
            If GetSubjectId(cboUnit.ItemData(cboUnit.ListIndex), mUnit.主题序号) = 0 Then Exit Function
                         
            strSQL = "select 标记序号,说明,图形索引,有效天数 from 病区标记内容 where 病区ID=[1] and  主题序号=[2] and 标记序号<>0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病区标记内容", mUnit.病区ID, mUnit.主题序号)
            
            StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_病区标记内容_Update(" & mUnit.病区ID & "," & mUnit.主题序号 & "," & _
                0 & ",'" & Replace(Trim(txtName.Text), "'", "") & "'," & 0 & "," & Val(txtDays.Text) & ")"
                
            '检查子分类的天数是否和分类相同，不同则进行修改
            With rsTmp
                Do While Not .EOF
                    If Nvl(!有效天数, 0) <> Val(txtDays.Text) Then
                        StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_病区标记内容_Update(" & mUnit.病区ID & "," & mUnit.主题序号 & "," & _
                            Nvl(!标记序号, 0) & ",'" & Replace(Nvl(!说明), "'", "") & "'," & Nvl(!图形索引, 0) & "," & Val(txtDays.Text) & ")"
                    End If
                .MoveNext
                Loop
            End With
            
            If UBound(StrSQLMod) > 1 Then
                gcnOracle.BeginTrans
                blnTran = True
                For i = 0 To UBound(StrSQLMod)
                    If StrSQLMod(i) <> "" Then Call zlDatabase.ExecuteProcedure(StrSQLMod(i), Me.Caption)
                Next i
                gcnOracle.CommitTrans
            Else
                For i = 0 To UBound(StrSQLMod)
                    If StrSQLMod(i) <> "" Then Call zlDatabase.ExecuteProcedure(StrSQLMod(i), Me.Caption)
                Next i
            End If
            fraUd.Tag = "1"
        End If
        strTemp = mUnit.主题序号 & "-0-" & mUnit.病区ID & "-" & Val(txtDays.Text)
    End If
    
    mblnChange = False
    
    fraInfo.Tag = ""
    fraUnit.Tag = ""
    UnitReportControl.Tag = ""
    '定位相应的列上
    Call RefreshData(lngRowIndex, strTemp)
    fraUd.Enabled = True
    UnitReportControl.SetFocus
    
    SaveData = True
    Exit Function
ErrHand:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRowIndex As Long '选择列的索引
    Dim i As Integer
    Dim Record As ReportRecord
    Dim strTemp As String, strSQL As String
    Dim blnTran As Boolean
    Dim cbrControl As CommandBarControl
    Dim strTmp1 As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHand

    
    Select Case Control.ID
        Case conMenu_File_PrintSet
            
            Call zlPrintSet
                    
        Case conMenu_File_Preview
            
            Call zlRptPrint(2)
        
        Case conMenu_File_Print
        
            Call zlRptPrint(1)
        
        Case conMenu_File_Excel
        
            Call zlRptPrint(3)
    
        Case conMenu_View_ToolBar_Button
        
            cbsMain(2).Visible = Not cbsMain(2).Visible
            cbsMain.RecalcLayout
        
        Case conMenu_View_ToolBar_Text
        
            For Each cbrControl In cbsMain(2).Controls
                If cbrControl.Type <> xtpControlLabel Then
                    cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
            cbsMain.RecalcLayout
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsMain.RecalcLayout
            
        Case conMenu_Edit_NewItem     '*新增
            fraInfo.Tag = "新增"
            fraUnit.Tag = ""
            Call SetFraResize
            Call SetControlEnable(True)
            mblnChange = True
        Case conMenu_Edit_Modify      '*修改(&M)
            
            fraInfo.Tag = "修改"
            fraUnit.Tag = ""
            Call SetControlEnable(True)
            mblnChange = True
            
        Case conMenu_Edit_Delete      '*删除(&D)
        
            If MsgBox("你确定要删除病区【" & Split(mstr病区名称, "-")(1) & "】内容【" & UnitReportControl.FocusedRow.Record(COL_说明).Value & "】的标记信息吗?", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            strTemp = UnitReportControl.FocusedRow.Record(COL_主题序号).Record.Tag
            
            mUnit.病区ID = CInt(Split(strTemp, "-")(2))
            mUnit.主题序号 = CInt(Split(strTemp, "-")(0))
            mUnit.标记序号 = CInt(Split(strTemp, "-")(1))
            
            '检查改主题内容该病区是否正在使用
            If CheckUseUnit(mUnit.病区ID, mUnit.主题序号, mUnit.标记序号) = True Then Exit Sub
            
            strSQL = "Zl_病区标记内容_Delete(" & mUnit.病区ID & "," & mUnit.主题序号 & "," & mUnit.标记序号 & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            '定位到下一列
            lngRowIndex = UnitReportControl.FocusedRow.Index
            
            Call UnitReportControl.Records.RemoveAt(UnitReportControl.FocusedRow.Record.Index)
            UnitReportControl.Populate
            
            If UnitReportControl.Records.Count > 0 Then
                lngRowIndex = IIf(UnitReportControl.Rows.Count - 1 > lngRowIndex, lngRowIndex, UnitReportControl.Rows.Count - 1)
                
                If UnitReportControl.Rows(lngRowIndex).GroupRow And UnitReportControl.Rows(lngRowIndex).Childs.Count <> 0 Then
                    lngRowIndex = lngRowIndex - 1
                End If
                
                If UnitReportControl.Rows(lngRowIndex).GroupRow Then
                    strTemp = UnitReportControl.Rows(lngRowIndex).Childs.Record(COL_主题序号).Record.Tag
                Else
                    strTemp = UnitReportControl.Rows(lngRowIndex).Record(COL_主题序号).Record.Tag
                End If
            End If
            Call RefreshData(lngRowIndex, strTemp)
            mblnChange = False
            fraUd.Tag = "1"
            fraUd.Enabled = True
            UnitReportControl.SetFocus
            gblnOK = True
        Case conMenu_Edit_NewParent '*新增分类
            fraInfo.Tag = ""
            fraUnit.Tag = "新增"
            Call SetFraResize(True)
            txtName.Enabled = True
            txtName.Text = ""
            txtDays.Enabled = True
            txtDays.Text = ""
            txtName.BackColor = UnEnable_Color
            txtDays.BackColor = UnEnable_Color
            txtName.SetFocus
            UnitReportControl.Tag = ""
            mblnChange = True
            
        Case conMenu_Edit_ModifyParent ' "修改分类(&U)"
            fraInfo.Tag = ""
            fraUnit.Tag = "修改"
            txtName.Enabled = True
            txtDays.Enabled = True
            txtName.BackColor = UnEnable_Color
            txtDays.BackColor = UnEnable_Color
            txtName.SetFocus
            UnitReportControl.Tag = UnitReportControl.FocusedRow.Index
            mblnChange = True

        Case conMenu_Edit_DeleteParent '"删除分类(&E)"
            If UnitReportControl.FocusedRow Is Nothing Then Exit Sub
            
            If MsgBox("你确定要删除病区【" & Split(mstr病区名称, "-")(1) & "】标记分类【" & UnitReportControl.FocusedRow.Childs(0).Record(COL_主题序号).GroupCaption & "】的信息吗?", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            strTemp = UnitReportControl.FocusedRow.Childs(0).Record(COL_主题序号).Record.Tag
            
            mUnit.病区ID = CInt(Split(strTemp, "-")(2))
            mUnit.主题序号 = CInt(Split(strTemp, "-")(0))
            mUnit.标记序号 = 0
            
            '检查改主题内容该病区是否正在使用
            If CheckUseUnit(mUnit.病区ID, mUnit.主题序号, mUnit.标记序号) = True Then Exit Sub
            
            strSQL = "Zl_病区标记内容_Delete(" & mUnit.病区ID & "," & mUnit.主题序号 & "," & mUnit.标记序号 & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            Call RefreshData(-1)
            
            mblnChange = False
            fraUd.Tag = "1"
            fraUd.Enabled = True
            UnitReportControl.SetFocus
            
        Case conMenu_Edit_Save     '*保存
            Call SaveData
            gblnOK = True
        Case conMenu_Edit_Reuse    '*取消
                        
            '记录现在选中的标注
            If UnitReportControl.SelectedRows.Count > 0 Then
                If Not UnitReportControl.SelectedRows(0) Is Nothing Then
                    If Not UnitReportControl.SelectedRows(0).GroupRow And UnitReportControl.SelectedRows(0).Childs.Count = 0 Then
                        lngRowIndex = UnitReportControl.SelectedRows(0).Index '用于快速重新定位
                        strTemp = UnitReportControl.SelectedRows(0).Record(COL_主题序号).Record.Tag
                    Else
                        lngRowIndex = UnitReportControl.SelectedRows(0).Index '用于快速重新定位
                        strTmp1 = UnitReportControl.SelectedRows(0).Childs(0).Record(COL_主题序号).Record.Tag
                        strTemp = Split(strTmp1, "-")(0) & "-0-" & Split(strTmp1, "-")(2) & "-" & Split(strTmp1, "-")(3)
                    End If
                End If
            Else
                If UnitReportControl.Tag <> "" Then
                    If Not UnitReportControl.Rows(UnitReportControl.Tag) Is Nothing Then
                        If Not UnitReportControl.Rows(UnitReportControl.Tag).GroupRow And UnitReportControl.Rows(UnitReportControl.Tag).Childs.Count = 0 Then
                            lngRowIndex = UnitReportControl.Rows(UnitReportControl.Tag).Index
                            strTemp = UnitReportControl.Rows(UnitReportControl.Tag).Record(COL_主题序号).Record.Tag
                        Else
                            lngRowIndex = UnitReportControl.Rows(UnitReportControl.Tag).Index
                            strTmp1 = UnitReportControl.Rows(UnitReportControl.Tag).Childs(0).Record(COL_主题序号).Record.Tag
                            strTemp = Split(strTmp1, "-")(0) & "-0-" & Split(strTmp1, "-")(2) & "-" & Split(strTmp1, "-")(3)
                        End If
                    End If
                End If
            End If
            
            fraInfo.Tag = ""
            fraUnit.Tag = ""
            Call RefreshData(lngRowIndex, strTemp)
            mblnChange = False
            fraUd.Enabled = True
            UnitReportControl.SetFocus
            
        Case conMenu_View_Refresh  '刷新
            '记录现在选中的标注
            If UnitReportControl.SelectedRows.Count > 0 Then
                If Not UnitReportControl.SelectedRows(0) Is Nothing Then
                    If Not UnitReportControl.SelectedRows(0).GroupRow And UnitReportControl.SelectedRows(0).Childs.Count = 0 Then
                        lngRowIndex = UnitReportControl.SelectedRows(0).Index '用于快速重新定位
                        strTemp = UnitReportControl.SelectedRows(0).Record(COL_主题序号).Record.Tag
                    Else
                        lngRowIndex = UnitReportControl.SelectedRows(0).Index '用于快速重新定位
                        strTmp1 = UnitReportControl.SelectedRows(0).Childs(0).Record(COL_主题序号).Record.Tag
                        strTemp = Split(strTmp1, "-")(0) & "-0-" & Split(strTmp1, "-")(2) & "-" & Split(strTmp1, "-")(3)
                    End If
                End If
            End If
            
            fraInfo.Tag = ""
            fraUnit.Tag = ""
            Call RefreshData(lngRowIndex, strTemp)
            mblnChange = False
            fraUd.Enabled = True
            UnitReportControl.SetFocus
            
        Case conMenu_Help_About
            
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
            
        Case conMenu_Help_Web_Home
            
            Call zlHomePage(Me.hwnd)
            
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.hwnd)

        Case conMenu_Help_Web_Mail '发送Email
            
            Call zlMailTo(Me.hwnd)
            
        Case conMenu_Help_Help        '*帮助主题(&H)
             Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
             
        Case conMenu_File_Exit        '*退出(&X)
            Unload Me
    End Select
    
    Call RefreshStateInfo
    
    cbsMain.RecalcLayout
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveData
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            Control.Enabled = (UnitReportControl.Records.Count > 0)
        Case conMenu_Edit_NewItem   '*新增(&A)
            If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange
                End If
            Else
                Control.Enabled = mlngCount > 0
            End If
        Case conMenu_Edit_Modify      '*修改(&M)
            If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then Control.Enabled = Not UnitReportControl.FocusedRow.GroupRow
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange And Val(Split(UnitReportControl.FocusedRow.Record(COL_主题序号).Record.Tag, "-")(1)) <> 0
                End If
            Else
                Control.Enabled = False
            End If
        Case conMenu_Edit_Delete      '*删除(&D)
            If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then Control.Enabled = Not UnitReportControl.FocusedRow.GroupRow
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange And Val(Split(UnitReportControl.FocusedRow.Record(COL_主题序号).Record.Tag, "-")(1)) <> 0
                End If
            Else
                Control.Enabled = False
            End If
        
        Case conMenu_Edit_NewParent '*新增分类
            Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
            If Control.Enabled = True Then
                Control.Enabled = Not mblnChange And mlngCount < 2 And UnitReportControl.FocusedRow.GroupRow
            Else
                If UnitReportControl.Rows.Count > 0 Then
                    Control.Enabled = Not mblnChange And mlngCount < 2
                Else
                    Control.Enabled = True And Not mblnChange
                End If
            End If
             
        Case conMenu_Edit_ModifyParent ' "修改分类(&U)"
             If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange And UnitReportControl.FocusedRow.GroupRow
                End If
             Else
                Control.Enabled = False
             End If
        Case conMenu_Edit_DeleteParent '"删除分类(&E)"
             If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange And UnitReportControl.FocusedRow.GroupRow
                End If
             Else
                Control.Enabled = False
             End If
        Case conMenu_Edit_Save     '*保存
            Control.Enabled = mblnChange
        Case conMenu_Edit_Reuse     '*取消
            Control.Enabled = mblnChange
        Case conMenu_View_Refresh '*刷新
            Control.Enabled = Not mblnChange
        Case conMenu_View_ToolBar_Button
            Control.Checked = Me.cbsMain(2).Visible
        Case conMenu_View_ToolBar_Text
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_StatusBar
            Control.Checked = Me.stbThis.Visible
    End Select
    
    cboUnit.Enabled = Not mblnChange
    fraUd.Enabled = Not mblnChange
    
End Sub

Private Sub lblSelect_DblClick(Index As Integer)
    Call showIcon(Index)
End Sub

Private Sub lblSelect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowSelect(Index)
End Sub

Private Sub picBack_LostFocus()
    'Call UnLoadImage
End Sub

Private Sub pic标记_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Call UnLoadImage
    End If
End Sub


Private Sub txtDays_KeyPress(KeyAscii As Integer)
     If KeyAscii > 45 And KeyAscii < 58 Then
        If KeyAscii = 46 Then
            If Len(txtDays.Text) = 0 Then
                KeyAscii = 0
            Else
                If InStr(1, txtDays.Text, ".") <> 0 Then
                    KeyAscii = 0
                End If
            End If
        End If
    Else
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtInfo_GotFocus()
    txtInfo.SelStart = Len(txtInfo.Text)
    Call zlControl.TxtSelAll(txtInfo)
End Sub


Private Sub txtInfo_Change()
    If mblnChange = False Then Exit Sub
    
    If fraInfo.Tag = "修改" Then
        With UnitReportControl.FocusedRow.Record(COL_说明)
            .Value = txtInfo.Text
        End With
        UnitReportControl.Populate
    End If
    
    '判定操作员是否手工录入修改了标注说明
    If lblSet(8).Tag <> "" And lblSet(8).Tag <> Trim(txtInfo.Text) And Trim(txtInfo.Text) <> cmdImage.Tag Then
        txtInfo.Tag = "改变"
    End If
    
    If imaCustom.ComboItems.Count > 0 Then cmdImage.Tag = imaCustom.Text
End Sub

Private Sub txtInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txtInfo.Text) <> "" Then
            txtInfo.Tag = "改变"
        End If
    Else
        If Chr(KeyCode) = "'" Then KeyCode = 0
    End If
End Sub


Private Sub txtName_Change()
    Dim i As Integer
    Dim lngPreIdx As Long
    Dim strTemp As String, str标记 As String
    If mblnChange = False Then Exit Sub
    
    If fraUnit.Tag = "修改" And UnitReportControl.Tag <> "" Then
        If UnitReportControl.Rows(UnitReportControl.Tag) Is Nothing Then Exit Sub
        With UnitReportControl.Rows(UnitReportControl.Tag)
            lngPreIdx = .Index
            strTemp = .Childs(0).Record(COL_主题序号).Record.Tag
            str标记 = Split(strTemp, "-")(0) & "-0-" & Split(strTemp, "-")(2) & "-" & Split(strTemp, "-")(3)
            
            For i = 0 To .Childs.Count - 1
                .Childs(i).Record(COL_主题序号).GroupCaption = "分组：" & Split(strTemp, "-")(0) & "-" & Replace(txtName.Text, "'", "")
            Next i
        End With
        UnitReportControl.Populate
    End If
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = Len(txtName.Text)
    Call zlControl.TxtSelAll(txtName)
End Sub

Private Sub txtDays_GotFocus()
    txtDays.SelStart = Len(txtDays.Text)
    Call zlControl.TxtSelAll(txtDays)
End Sub

Private Sub txtDays_Change()
    Dim i As Integer
    If mblnChange = False Then Exit Sub
    '更改分类天数时，子分类同步更新
    If fraUnit.Tag = "修改" And UnitReportControl.Tag <> "" Then
        If UnitReportControl.Rows(UnitReportControl.Tag) Is Nothing Then Exit Sub
        With UnitReportControl.Rows(UnitReportControl.Tag)
            For i = 0 To .Childs.Count - 1
                If Val(Split(.Childs(i).Record(COL_主题序号).Record.Tag, "-")(1)) = 0 Then
                    .Childs(i).Record(COL_有效天数).Value = ""

                Else
                    .Childs(i).Record(COL_有效天数).Value = IIf(txtDays.Text = "", 0, txtDays.Text)
                End If
            Next i
        End With
        UnitReportControl.Populate
    End If
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chr(KeyCode) = "'" Then KeyCode = 0
End Sub

Private Sub UnitReportControl_ColumnClick(ByVal Column As XtremeReportControl.IReportColumn)
    Call Arrange(Column.Index)
End Sub

Public Sub Arrange(Column As Long)
    UnitReportControl.SortOrder.DeleteAll
    UnitReportControl.SortOrder.Add UnitReportControl.Columns.Find(Column)
    UnitReportControl.SortOrder(0).SortAscending = Not UnitReportControl.SortOrder(0).SortAscending
    UnitReportControl.Populate
End Sub


Private Sub UnitReportControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
         If Not (UnitReportControl.FocusedRow Is Nothing) Then
            If Not UnitReportControl.FocusedRow.GroupRow And UnitReportControl.FocusedRow.Childs.Count = 0 Then
              Call UnitReportControl_RowDblClick(UnitReportControl.FocusedRow, UnitReportControl.FocusedRow.Record.Item(COL_主题序号))
            End If
        End If
    End If
End Sub

Private Sub UnitReportControl_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
'功能:弹出邮件菜单
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    If Button <> 2 Then Exit Sub
    
    If cbsMain.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub UnitReportControl_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Not (Row Is Nothing) Then
        If Not Row.GroupRow And Row.Childs.Count = 0 And Val(Split(Row.Record(COL_主题序号).Record.Tag, "-")(1)) <> 0 Then
            Call cbsMain_Execute(cbsMain.FindControl(, conMenu_Edit_Modify, True, True))
        Else
            Call cbsMain_Execute(cbsMain.FindControl(, conMenu_Edit_ModifyParent, True, True))
        End If
    End If
End Sub


Private Sub UnitReportControl_SelectionChanged()
'-------------------------------------------------
'功能:根据ReportControl的选择列，提取对应的病区主题信息
'
'--------------------------------------------------
    Dim i As Integer
    txtInfo.Text = "": txtInfo.Tag = "": lblSet(7).Tag = "": lblSet(8).Tag = "": imaCustom.Text = "": imaCustom.Tag = ""
    lblSet(9).Tag = "": cbo标记.Tag = "": lblSet(1).Tag = "": txtName.Text = "": lblSet(4).Tag = "": txtDays.Text = ""
    
    On Error GoTo ErrHand
        With UnitReportControl.FocusedRow
            If Not UnitReportControl.FocusedRow Is Nothing Then
                If Not .GroupRow And .Childs.Count = 0 Then
                    If Val(Split(.Record(COL_主题序号).Record.Tag, "-")(1)) <> 0 Then
                        cbo标记.ListIndex = SetCboIndex(cbo标记, Val(Split(.Record(COL_主题序号).Record.Tag, "-")(0)))
                        lblSet(9).Tag = cbo标记.ListIndex
                        lblSet(8).Tag = .Record(COL_说明).Value
                        txtInfo.Text = .Record(COL_说明).Value
                        lblSet(7).Tag = IIf(Val(.Record(COL_标注).Icon) <= 0, "0", Val(.Record(COL_标注).Icon))
                        If lblSet(7).Tag >= mlngImgIndex Then
                            imaCustom.ComboItems(Val(lblSet(7).Tag) - mlngImgIndex + 1).Selected = True
                        End If
                        Call SetControlEnable(fraInfo.Tag <> "")
                        Call SetFraResize
                    Else
                        UnitReportControl.FocusedRow = UnitReportControl.Rows(UnitReportControl.FocusedRow.Index - 1)
                    End If
                Else
                    lblSet(1).Tag = Split(.Childs(0).Record(COL_主题序号).GroupCaption, "-")(1)
                    txtName.Text = lblSet(1).Tag
                    lblSet(4).Tag = Val(.Childs(0).Record(COL_有效天数).Value)
                    txtDays.Text = lblSet(4).Tag
                    
                    txtName.Enabled = fraUnit.Tag <> ""
                    txtDays.Enabled = fraUnit.Tag <> ""
                    
                    txtName.BackColor = IIf(fraUnit.Tag <> "", UnEnable_Color, Enable_Color)
                    txtDays.BackColor = IIf(fraUnit.Tag <> "", UnEnable_Color, Enable_Color)
                    
                    Call SetFraResize(True)
                End If
            End If
        End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SetCboIndex(ByVal objCbo As Object, ByVal intItemData As Integer) As Integer
'------------------------------------------------------------------------
'功能:根据itemdata的值获取cbo的Index
'------------------------------------------------------------------------
    Dim i As Integer
    Dim intIndex As Integer
    
    intIndex = -1
    
    For i = 0 To objCbo.ListCount - 1
        If Val(objCbo.ItemData(i)) = intItemData Then
           intIndex = i
           Exit For
        End If
    Next i
    
    SetCboIndex = intIndex
End Function

Private Function GetCboText(ByVal objCbo As Object, ByVal intItemData As Integer) As String
'------------------------------------------------------------------------
'功能:根据itemdata的值获取cbo的Index
'------------------------------------------------------------------------
    Dim i As Integer
    Dim strText As String
    
    strText = ""
    
    For i = 0 To objCbo.ListCount - 1
        If Val(objCbo.ItemData(i)) = intItemData Then
           strText = objCbo.Text
           Exit For
        End If
    Next i
    
    GetCboText = strText
End Function

Private Function CheckChange() As Boolean
'-----------------------------------------------------
'功能:修改时检查内容是否发生变化
'-----------------------------------------------------
    Dim blnChage As Boolean
    If fraInfo.Tag = "修改" Then
        If lblSet(9).Tag <> cbo标记.ListIndex Or lblSet(8).Tag <> txtInfo.Text Or _
            lblSet(7).Tag <> imaCustom.SelectedItem.Index - 1 + mlngImgIndex Then
            blnChage = True
        End If
    ElseIf fraUnit.Tag = "修改" Then
        If lblSet(1).Tag <> txtName.Text Or lblSet(4).Tag <> txtDays.Text Then
            blnChage = True
        End If
    End If
    CheckChange = blnChage
End Function

Private Function CheckUseUnit(ByVal lngUnitID As Long, ByVal lngSubjectID As Long, ByVal lngTracerID As Long) As Boolean
'----------------------------------------------------------
'功能：检查改标记内容是否正在使用
'参数：lngUnitId 病区ID，lngSubjectID 主题序号 ，lngTracerID 标记序号
'----------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim blnTrue As Boolean
    Dim strSQL
    On Error GoTo ErrHand
    
    If lngTracerID <> 0 Then
        strSQL = "select 1 From 病区标记记录" & _
            "   where  病区Id=[1] and 主题序号=[2] and 标记序号=[3] and rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病区标记记录", lngUnitID, lngSubjectID, lngTracerID)
        If Not rsTmp.EOF Then
            blnTrue = True
            MsgBox "该标记内容目前改病区正在使用,请取消使用后在删除.", vbInformation, gstrSysName
        End If
    Else
        strSQL = _
            " SELECT 1" & vbNewLine & _
            " FROM 病区标记内容 A,病区标记记录 B" & vbNewLine & _
            " WHERE  A.病区ID=B.病区ID And A.主题序号=B.主题序号 And  A.病区ID=[1]  And A.主题序号=[2]  " & vbNewLine & _
            " and rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病区标记记录", lngUnitID, lngSubjectID)
        If Not rsTmp.EOF Then
            blnTrue = True
            MsgBox "该标记分类下的标记内容目前改病区正在使用,请取消使用后在删除.", vbInformation, gstrSysName
        End If
    End If
    CheckUseUnit = blnTrue
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetMaxPreID(ByVal lng病区ID As Long, ByVal lngPreVId As Long) As Long
'--------------------------------------------------------------------
'功能:提取某病区某主题下的最大标记序号
'参数:lng病区ID：病区ID ； lngPreVID ：主题序号
'--------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngPreID As Long
    On Error GoTo ErrHand
    strSQL = _
        " select MAX(标记序号) PreID " & _
        " From 病区标记内容" & _
        " Where 病区ID=[1] and 主题序号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病区ID, lngPreVId)
    If rsTemp.EOF Then
        lngPreID = 1
    Else
        lngPreID = Val(Nvl(rsTemp!PreID, 0)) + 1
    End If
    
    GetMaxPreID = lngPreID
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetSubjectId(ByVal lng病区ID As Long, Optional lng主题序号 As Long = 0) As Long
'------------------------------------------------------------------------
'功能:新增标注分类时，提取某病区标记主题的主题序号.主题序号（1,2）
'     修改分类时检查名称是否与其他分类名称重复。lng主题序号=0 新增，lng主题序号<>0 为修改的主题序号
'------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim strInfo As String
    Dim lngSubjectID As Long
    
    On Error GoTo ErrHand:
    If lng主题序号 = 0 Then
        strSQL = _
            " select 主题序号,说明 from 病区标记内容" & _
            " where 病区Id=[1] and 主题序号=[2] and 标记序号=0" & _
            " union all" & _
            " select 主题序号,说明 from 病区标记内容" & _
            " where 病区Id=[1] and 主题序号=[3] and 标记序号=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病区标记内容", lng病区ID, 1, 2)
    Else
        strSQL = _
               " select 主题序号,说明 from 病区标记内容" & _
               " where 病区Id=[1] and 主题序号=[2] and 标记序号=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病区标记内容", lng病区ID, 2 \ lng主题序号)
    End If
    
    With rsTmp
        Do While Not .EOF
            strTmp = strTmp & "," & !主题序号
            strInfo = strInfo & "," & !说明
            .MoveNext
        Loop
    End With
    strTmp = Mid(strTmp, 2)
    strInfo = Mid(strInfo, 2)
    
    '检查标记分类名称是否重复
    If InStr(1, "," & strInfo & ",", "," & Trim(txtName.Text) & ",") <> 0 Then
        MsgBox "此标记名称已经存在,请重新填写标记名称.", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Function
    Else
        lngSubjectID = 1
    End If
    
    If lng主题序号 = 0 Then
        If InStr(1, strTmp, "1") = 0 Then
            lngSubjectID = 1
        ElseIf InStr(1, strTmp, "2") = 0 Then
            lngSubjectID = 2
        Else
            lngSubjectID = 1
        End If
    End If
    GetSubjectId = lngSubjectID
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckUnitSubject(ByVal lng病区ID As Long) As Long
'---------------------------------------------------
'功能:检查是否存在标注主题名称,不存在提示操作员进行设置
'---------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrHand

    strSQL = _
        " select 主题序号,说明 from 病区标记内容" & _
        " where 病区Id=[1] and 主题序号=[2] and 标记序号=0" & _
        " union all" & _
        " select 主题序号,说明 from 病区标记内容" & _
        " where 病区Id=[1] and 主题序号=[3] and 标记序号=0"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病区标记内容", lng病区ID, 1, 2)
    
    cbo标记.Clear
    With rsTmp
        Do While Not .EOF
            cbo标记.AddItem Nvl(!说明, "个性标注" & Nvl(!主题序号))
            cbo标记.ItemData(cbo标记.NewIndex) = Val(Nvl(!主题序号))
            If cbo标记.ListIndex = -1 Then
                Call Cbo.SetIndex(cbo标记.hwnd, cbo标记.NewIndex)
            End If
        .MoveNext
        Loop
    End With
                
    CheckUnitSubject = rsTmp.RecordCount

    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'################################################################################################################
'## 功能：  将数据从一个XtremeReportControl控件复制到VSFlexGrid，以便进行打印
'################################################################################################################
Private Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '将全部组强制展开,复制数据表格
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    Dim strGroupCaption As String
    
    Dim lngCol As Long, lngRow As Long
    
    On Error GoTo ErrHand:
    For Each rptRow In rptList.Rows
        If rptRow.GroupRow Then rptRow.Expanded = True
    Next
    
    With vfgList
        .Clear
        .Rows = rptList.Records.Count + 1
        .Cols = 0: .Cols = rptList.Columns.Count
        .FixedCols = rptList.GroupsOrder.Count
        
        '标题行复制
        .Row = 0
        lngCol = 0
        For Each rptCol In rptList.GroupsOrder
            .TextMatrix(0, lngCol) = rptCol.Caption
            .ColData(lngCol) = rptCol.ItemIndex
            Select Case rptCol.Alignment
            Case xtpAlignmentLeft: .FixedAlignment(lngCol) = flexAlignLeftCenter
            Case xtpAlignmentCenter: .FixedAlignment(lngCol) = flexAlignCenterCenter
            Case xtpAlignmentRight:  .FixedAlignment(lngCol) = flexAlignRightCenter
            End Select
            .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .FixedAlignment(lngCol)
            .ColWidth(lngCol) = 100 * 15
            .MergeCol(lngCol) = True
            lngCol = lngCol + 1
        Next
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .TextMatrix(0, lngCol) = rptCol.Caption
                If rptCol.Caption = "标注" Then rptCol.Width = 10
                .ColData(lngCol) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .ColAlignment(lngCol) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .ColAlignment(lngCol) = flexAlignCenterCenter
                Case xtpAlignmentRight: .ColAlignment(lngCol) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
                .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .ColAlignment(lngCol)
                If rptCol.Width < 20 Then
                    .ColWidth(lngCol) = 0
                Else
                    .ColWidth(lngCol) = rptCol.Width * 15
                End If
                lngCol = lngCol + 1
            End If
        Next
        vfgList.Cols = lngCol
        
        '数据行复制
        lngRow = 0
        For Each rptRow In rptList.Rows
            If rptRow.GroupRow = False Then
                lngRow = lngRow + 1
                For lngCol = 0 To .Cols - 1
                    If rptRow.Record(.ColData(lngCol)).GroupCaption <> "" Then
                        strGroupCaption = Split(rptRow.Record(.ColData(lngCol)).GroupCaption, "：")(1)
                    Else
                        strGroupCaption = rptRow.Record(.ColData(lngCol)).GroupCaption
                    End If
                    .TextMatrix(lngRow, lngCol) = IIf(.TextMatrix(0, lngCol) = "主题序号", strGroupCaption, rptRow.Record(.ColData(lngCol)).Value)
                    If rptRow.Record(.ColData(lngCol)).Icon > 0 Then
                        '.CellPicture = Img标记(999).ListImages(rptRow.Record(.ColData(lngCol)).Icon).Picture
                    End If
                Next
            End If
        Next
    End With
    zlReportToVSFlexGrid = True
    Exit Function

ErrHand:
    zlReportToVSFlexGrid = False
End Function


Private Function ReDimArray(ByRef strArray() As String) As Long
    '----------------------------------------------------------------------
    '功能：重新定义数组
    '----------------------------------------------------------------------
    Dim lngCount As Long
    Dim strTmp As String
    
    On Error GoTo InitHand
    strTmp = strArray(0)
    lngCount = UBound(strArray) + 1
    GoTo OkHand
InitHand:
    lngCount = 1
OkHand:
    ReDim Preserve strArray(0 To lngCount)
    ReDimArray = lngCount
End Function


