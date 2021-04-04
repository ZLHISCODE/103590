VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPatholArchivesStation 
   Caption         =   "病理归档工作站"
   ClientHeight    =   8895
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   14760
   Icon            =   "frmPatholArchivesStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   14760
   StartUpPosition =   3  '窗口缺省
   Begin zl9PacsControl.ucSplitter ucSplitter1 
      Height          =   7620
      Left            =   4455
      TabIndex        =   29
      Top             =   840
      Width           =   100
      _ExtentX        =   185
      _ExtentY        =   13441
      BackColor       =   -2147483633
      SplitWidth      =   100
      SplitLevel      =   3
      SyncParentHeight=   0   'False
      AllowPaintOtherSpliter=   -1  'True
      Control1Name    =   "Picture1"
      Control2Name    =   "Picture2"
   End
   Begin MSComDlg.CommonDialog diaFont 
      Left            =   4320
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgMenus 
      Left            =   4680
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":179A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":1AEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":1E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":21C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":2512
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":2864
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":2BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":2F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":325A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":35AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":38FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":3C50
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":3FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":42F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":4646
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":4998
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":4CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":503C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":538E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":56E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":5A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":5D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":60D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":6428
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":677A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":6ACC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":6E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":7170
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":74C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":819C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":8E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":9B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":A82A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":B504
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":C1DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":CEB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":DB92
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholArchivesStation.frx":E86C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrTools 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   1402
      ButtonWidth     =   1455
      ButtonHeight    =   1349
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "标签预览"
            Key             =   "tbn_LabView"
            Object.Tag             =   "标签预览"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "标签打印"
            Key             =   "tbn_LabPrint"
            Object.Tag             =   "标签打印"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增档案"
            Key             =   "tbn_NewArchives"
            Object.Tag             =   "新增档案"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "删除档案"
            Key             =   "tbn_DelArchives"
            Object.Tag             =   "删除档案"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "更新档案"
            Key             =   "tbn_UpdateArchives"
            Object.Tag             =   "更新档案"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查询档案"
            Key             =   "tbn_QueryArchives"
            Object.Tag             =   "查询档案"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "档案归档"
            Key             =   "tbn_EnterArchives"
            Object.Tag             =   "档案归档"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "撤销归档"
            Key             =   "tbn_CancelArchives"
            Object.Tag             =   "撤销归档"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "帮助"
            Key             =   "tbn_Help"
            Object.Tag             =   "帮助"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "tbn_Exit"
            Object.Tag             =   "退出"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   7620
      Left            =   4555
      ScaleHeight     =   7620
      ScaleWidth      =   10200
      TabIndex        =   1
      Top             =   840
      Width           =   10205
      Begin VB.TextBox txtNumberInf 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "当前材料数量：0   在档数量：0   已借数量：0   遗失数量：0   "
         Top             =   90
         Width           =   5415
      End
      Begin VB.PictureBox picTag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame framArchivesDetail 
         Height          =   6735
         Left            =   1320
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CommandButton cmdFilter 
            Caption         =   "过 滤(&L)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton cmdPreview 
            Caption         =   "预 览(&W)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "删 除(&D)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   7320
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "打 印(&P)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton cmdRead 
            Caption         =   "读取档案内容(&R)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   6240
            Width           =   1695
         End
         Begin zl9PACSWork.ucFlexGrid ufgArchivesDetail 
            Height          =   5895
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   10398
            GridRows        =   201
            BackColor       =   12648447
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
      End
      Begin VB.Frame framEnterArchives 
         Height          =   7095
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   9735
         Begin VB.CommandButton cmdEnterArchives 
            Caption         =   "材料入档(&I)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   5760
            TabIndex        =   42
            Top             =   6360
            Width           =   1335
         End
         Begin VB.CheckBox chkTeShu 
            Caption         =   "特检材料"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4680
            TabIndex        =   25
            Top             =   6360
            Width           =   1215
         End
         Begin VB.CheckBox chkSlices 
            Caption         =   "切片材料"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   24
            Top             =   6360
            Width           =   1335
         End
         Begin VB.CheckBox chkWaxStone 
            Caption         =   "蜡块材料"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   23
            Top             =   6360
            Width           =   1215
         End
         Begin VB.CheckBox chkNotEnter 
            Caption         =   "尚未入档"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   22
            Top             =   6360
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkComplete 
            Caption         =   "检查完成"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   6360
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.ComboBox cbxRequestDetail 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmPatholArchivesStation.frx":F546
            Left            =   6120
            List            =   "frmPatholArchivesStation.frx":F548
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   840
            Width           =   1455
         End
         Begin VB.ComboBox cbxRequestType 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmPatholArchivesStation.frx":F54A
            Left            =   3600
            List            =   "frmPatholArchivesStation.frx":F54C
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   840
            Width           =   1455
         End
         Begin VB.ComboBox cbxStudyType 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmPatholArchivesStation.frx":F54E
            Left            =   1080
            List            =   "frmPatholArchivesStation.frx":F550
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   840
            Width           =   1455
         End
         Begin VB.Frame framQuery 
            Height          =   735
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   9735
            Begin VB.ComboBox cbxQueryType 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmPatholArchivesStation.frx":F552
               Left            =   120
               List            =   "frmPatholArchivesStation.frx":F55C
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton cmdQuery 
               Caption         =   "查询(&Q)"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   7680
               TabIndex        =   14
               Top             =   180
               Width           =   975
            End
            Begin VB.TextBox txtEndPatholNum 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6600
               TabIndex        =   13
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox txtStartPatholNum 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   5400
               TabIndex        =   11
               Top             =   240
               Width           =   975
            End
            Begin MSComCtl2.DTPicker dtpStartDate 
               Height          =   330
               Left            =   1320
               TabIndex        =   7
               Top             =   240
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "yyyy-MM-dd 00:00:00"
               Format          =   114032643
               CurrentDate     =   40884
            End
            Begin MSComCtl2.DTPicker dtpEndDate 
               Height          =   330
               Left            =   3015
               TabIndex        =   9
               Top             =   240
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "yyyy-MM-dd 23:59:59"
               Format          =   114032643
               CurrentDate     =   40884
            End
            Begin VB.Label Label3 
               Caption         =   "到"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6390
               TabIndex        =   12
               Top             =   300
               Width           =   255
            End
            Begin VB.Label Label2 
               Caption         =   "病理号："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4680
               TabIndex        =   10
               Top             =   300
               Width           =   735
            End
            Begin VB.Label labTo 
               Caption         =   "到"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2805
               TabIndex        =   8
               Top             =   300
               Width           =   255
            End
         End
         Begin zl9PACSWork.ucFlexGrid ufgMaterialQuery 
            Height          =   4935
            Left            =   120
            TabIndex        =   5
            Top             =   1320
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   8705
            GridRows        =   201
            BackColor       =   12648447
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
         Begin VB.Line lineSplit2 
            BorderColor     =   &H00C0C0C0&
            X1              =   2400
            X2              =   2400
            Y1              =   6360
            Y2              =   6600
         End
         Begin VB.Label labRequestDetail 
            Caption         =   "检查细目："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   19
            Top             =   885
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "检查过程："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   17
            Top             =   885
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "检查类型："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   880
            Width           =   975
         End
      End
      Begin XtremeSuiteControls.TabControl tabFilter 
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   4125
         _Version        =   589884
         _ExtentX        =   7276
         _ExtentY        =   661
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   7620
      Left            =   0
      ScaleHeight     =   7620
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   840
      Width           =   4455
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   7335
         Left            =   120
         ScaleHeight     =   7335
         ScaleWidth      =   4335
         TabIndex        =   31
         Top             =   120
         Width           =   4335
         Begin zl9PacsControl.ucSplitter ucSplitter2 
            Height          =   100
            Left            =   0
            TabIndex        =   32
            Top             =   3930
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   185
            BackColor       =   -2147483633
            MousePointer    =   7
            SplitWidth      =   100
            SplitType       =   0
            SplitLevel      =   3
            Control1Name    =   "ufgArchives"
            Control2Name    =   "rtbDetail"
         End
         Begin RichTextLib.RichTextBox rtbDetail 
            Height          =   3305
            Left            =   0
            TabIndex        =   33
            Top             =   4030
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   5821
            _Version        =   393217
            BackColor       =   16761024
            BorderStyle     =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmPatholArchivesStation.frx":F574
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin zl9PACSWork.ucFlexGrid ufgArchives 
            Height          =   3930
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   6932
            GridRows        =   201
            BackColor       =   12648447
            IsEnterNextCell =   0   'False
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   28
      Top             =   8535
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   1764
            Picture         =   "frmPatholArchivesStation.frx":F611
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "未归档档案数："
            TextSave        =   "未归档档案数："
            Key             =   "sb_NoEnter"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "未入档蜡块数："
            TextSave        =   "未入档蜡块数："
            Key             =   "sb_NoEnterWaxStone"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3176
            MinWidth        =   3176
            Text            =   "未入档切片数："
            TextSave        =   "未入档切片数："
            Key             =   "sb_NoEnterSlices"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "未入档特检数："
            TextSave        =   "未入档特检数："
            Key             =   "sb_NoEnterSpeEx"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6802
            MinWidth        =   2
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Menu mnu_File 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnu_ParameterConfig 
         Caption         =   "参数设置(&M)"
      End
      Begin VB.Menu mnu_PrintConfig 
         Caption         =   "打印设置(&C)"
      End
      Begin VB.Menu mnu_Split10 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ArchivesClassCfg 
         Caption         =   "档案分类设置(&A)"
      End
      Begin VB.Menu mnu_Split4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ListPreview 
         Caption         =   "预 览(&V)"
      End
      Begin VB.Menu mnu_ListPrint 
         Caption         =   "打 印(&P)"
      End
      Begin VB.Menu mnu_Split5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ExportExcel 
         Caption         =   "输出到Excel(&E)"
      End
      Begin VB.Menu mnu_Split3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "退 出(&Q)"
      End
   End
   Begin VB.Menu mnu_Edit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnu_NewArchives 
         Caption         =   "新增档案(&N)"
      End
      Begin VB.Menu mnu_DelArchives 
         Caption         =   "删除档案(&D)"
      End
      Begin VB.Menu mnu_UpdateArchives 
         Caption         =   "更新档案(&U)"
      End
      Begin VB.Menu mnu_Split2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_EnterArchives 
         Caption         =   "档案归档(&T)"
      End
      Begin VB.Menu mnu_CancelArchives 
         Caption         =   "撤销归档(&R)"
      End
      Begin VB.Menu mnu_Split6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_QueryArchives 
         Caption         =   "查询档案(&Q)"
      End
      Begin VB.Menu mnu_Split9 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_LabPreview 
         Caption         =   "标签预览(&V)"
      End
      Begin VB.Menu mnu_LabPrint 
         Caption         =   "标签打印(&P)"
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnu_ToolsBar 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnu_StandardBut 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_WordLabel 
            Caption         =   "文本标签(&L)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu_StateBar 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_Split7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Font 
         Caption         =   "字 体(&F)"
      End
   End
   Begin VB.Menu mnu_Tools 
      Caption         =   "工具(&T)"
      Visible         =   0   'False
      Begin VB.Menu mnu_Zoom 
         Caption         =   "放大镜(&Z)"
      End
      Begin VB.Menu mnu_Calc 
         Caption         =   "计算器(&C)"
      End
   End
   Begin VB.Menu mnu_MainHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnu_Help 
         Caption         =   "帮助主题(&H)"
      End
      Begin VB.Menu mnu_WebZL 
         Caption         =   "WEB上的中联(&W)"
         Begin VB.Menu mnu_MainPage 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnu_BBS 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnu_Return 
            Caption         =   "发送反馈(&K)"
         End
      End
      Begin VB.Menu mnu_Split8 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_About 
         Caption         =   "关于...(&A)"
      End
   End
End
Attribute VB_Name = "frmPatholArchivesStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const DebugState = False

Private Const ArchivesState_NoEnter As String = "未归档"
Private Const ArchivesState_Enter As String = "已归档"

'为菜单设置相应的图形
Private Const MF_BITMAP = &H400&


'档案材料类型枚举
Private Enum TArchivesMaterialType
    amtTable = 0
    amtMaterial = 1
    amdReport = 2
End Enum


Private mstrPrivs As String
Private mcurMaterialType As TArchivesMaterialType
Private mlngCurArchivesId As Long
Private mblnMoved As Boolean

Private mlngDefaultQueryDays As Long
Private mstrLabelReportName As String
Private mblnIsFormLoaded As Boolean



Dim WithEvents zlReport As zl9Report.clsReport
Attribute zlReport.VB_VarHelpID = -1



Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long '取得窗口的菜单句柄,hwnd是窗口的句柄
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal npos As Long) As Long '取得子菜单句柄，nPos是菜单的位置
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal npos As Long, ByVal wFlags As Long, ByVal hBitUnchecked As Long, ByVal hBitChecked As Long) As Long







Private Sub InitMenuIcoConfig()
'初始化菜单图标显示
On Error Resume Next
    Dim hMenu As Long
    Dim hSubMenu As Long
    Dim hSubSubMenu As Long
    
    hMenu = GetMenu(Me.hWnd)
    
    '设置第一项菜单(文件)
    hSubMenu = GetSubMenu(hMenu, 0) '取得第一项菜单的子菜单句柄
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(28).Picture, imgMenus.ListImages(28).Picture) '参数设置
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(3).Picture, imgMenus.ListImages(3).Picture) '打印设置
    Call SetMenuItemBitmaps(hSubMenu, 3, MF_BITMAP, imgMenus.ListImages(6).Picture, imgMenus.ListImages(6).Picture) '档案分类设置
    Call SetMenuItemBitmaps(hSubMenu, 5, MF_BITMAP, imgMenus.ListImages(18).Picture, imgMenus.ListImages(18).Picture) '打印预览
    Call SetMenuItemBitmaps(hSubMenu, 6, MF_BITMAP, imgMenus.ListImages(19).Picture, imgMenus.ListImages(19).Picture) '打印
    Call SetMenuItemBitmaps(hSubMenu, 8, MF_BITMAP, imgMenus.ListImages(4).Picture, imgMenus.ListImages(4).Picture) '导出Excel
    Call SetMenuItemBitmaps(hSubMenu, 10, MF_BITMAP, imgMenus.ListImages(5).Picture, imgMenus.ListImages(5).Picture) '退出
    

    '设置第二项菜单（编辑）
    hSubMenu = GetSubMenu(hMenu, 1) '取得第二项菜单的子菜单句柄
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(7).Picture, imgMenus.ListImages(7).Picture) '新增档案
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(8).Picture, imgMenus.ListImages(8).Picture) '删除档案
    Call SetMenuItemBitmaps(hSubMenu, 2, MF_BITMAP, imgMenus.ListImages(9).Picture, imgMenus.ListImages(9).Picture) '更新档案
    Call SetMenuItemBitmaps(hSubMenu, 4, MF_BITMAP, imgMenus.ListImages(10).Picture, imgMenus.ListImages(10).Picture) '档案归档
    Call SetMenuItemBitmaps(hSubMenu, 5, MF_BITMAP, imgMenus.ListImages(11).Picture, imgMenus.ListImages(11).Picture) '撤销归档
    Call SetMenuItemBitmaps(hSubMenu, 7, MF_BITMAP, imgMenus.ListImages(12).Picture, imgMenus.ListImages(12).Picture) '查询档案
    Call SetMenuItemBitmaps(hSubMenu, 9, MF_BITMAP, imgMenus.ListImages(1).Picture, imgMenus.ListImages(1).Picture) '打印预览
    Call SetMenuItemBitmaps(hSubMenu, 10, MF_BITMAP, imgMenus.ListImages(2).Picture, imgMenus.ListImages(2).Picture) '打印
    
    
    '设置第二项菜单（查看）
    hSubMenu = GetSubMenu(hMenu, 2) '取得第二项菜单的子菜单句柄
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(27).Picture, imgMenus.ListImages(27).Picture) '工具栏
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(22).Picture, imgMenus.ListImages(21).Picture) '状态栏
    Call SetMenuItemBitmaps(hSubMenu, 3, MF_BITMAP, imgMenus.ListImages(23).Picture, imgMenus.ListImages(23).Picture) '字体
    
        hSubSubMenu = GetSubMenu(hSubMenu, 0)
    
        Call SetMenuItemBitmaps(hSubSubMenu, 0, MF_BITMAP, imgMenus.ListImages(26).Picture, imgMenus.ListImages(20).Picture) '标准按钮
        Call SetMenuItemBitmaps(hSubSubMenu, 1, MF_BITMAP, imgMenus.ListImages(25).Picture, imgMenus.ListImages(24).Picture) '文本标签
    
    
    
    '设置第五项菜单（帮助）
    hSubMenu = GetSubMenu(hMenu, 3) '取得第五项菜单的子菜单句柄
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(13).Picture, imgMenus.ListImages(13).Picture) '帮助主题
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(14).Picture, imgMenus.ListImages(14).Picture) 'web中联
    Call SetMenuItemBitmaps(hSubMenu, 3, MF_BITMAP, imgMenus.ListImages(15).Picture, imgMenus.ListImages(15).Picture) '关
    
        hSubSubMenu = GetSubMenu(hSubMenu, 1)
    
        Call SetMenuItemBitmaps(hSubSubMenu, 0, MF_BITMAP, imgMenus.ListImages(14).Picture, imgMenus.ListImages(13).Picture) '帮助主题
        Call SetMenuItemBitmaps(hSubSubMenu, 1, MF_BITMAP, imgMenus.ListImages(16).Picture, imgMenus.ListImages(16).Picture) '中联论坛
        Call SetMenuItemBitmaps(hSubSubMenu, 2, MF_BITMAP, imgMenus.ListImages(17).Picture, imgMenus.ListImages(17).Picture) '发送反馈
    
    err.Clear

End Sub


Private Sub RefreshStateInf(ByVal blnIsRefreshArchives As Boolean, ByVal blnIsRefreshMaterial As Boolean)
'刷新状态信息，如未归档档案数量，未归档蜡块数量等...
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    If blnIsRefreshArchives Then
        '刷新档案数量
        strSql = "select /*+ Rule*/ count(1) as 返回值 from 病理档案信息 where 档案状态=0"
                
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        
        If rsData.RecordCount > 0 Then
            stbThis.Panels(2).Text = "未归档档案数：" & Nvl(rsData!返回值)
        End If
    End If
    
    If blnIsRefreshMaterial Then
        '刷新蜡块数量
        strSql = "select /*+ Rule*/ count(1) as 返回值 from 病理取材信息 a, 病人医嘱发送 b, 病理检查信息 c " & _
                " Where a.病理医嘱id = c.病理医嘱id And b.医嘱ID = c.医嘱ID And b.执行过程 = 6 And a.归档状态 = 0 and a.确认状态=1 " & _
                " and a.取材时间 between sysdate - 365 and sysdate "
                
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        
        If rsData.RecordCount > 0 Then
            stbThis.Panels(3).Text = "未入档蜡块数：" & Nvl(rsData!返回值)
        End If
    
    
    
        '刷新制片数量
        strSql = "select /*+ Rule*/ count(1) as 返回值 from 病理制片信息 a, 病人医嘱发送 b, 病理检查信息 c " & _
                " Where a.病理医嘱id = c.病理医嘱id And b.医嘱ID = c.医嘱ID And b.执行过程 = 6 And a.归档状态 = 0 and a.当前状态=2 " & _
                " and a.制片时间 between sysdate - 365 and sysdate "
                
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        
        If rsData.RecordCount > 0 Then
            stbThis.Panels(4).Text = "未入档制片数：" & Nvl(rsData!返回值)
        End If
        
        
        '刷新特检数量
        strSql = "select /*+ Rule*/ count(1) as 返回值 from 病理特检信息 a, 病人医嘱发送 b, 病理检查信息 c " & _
                " Where a.病理医嘱id = c.病理医嘱id And b.医嘱ID = c.医嘱ID And b.执行过程 = 6 And a.归档状态 = 0 and a.当前状态=2 " & _
                " and a.完成时间 between sysdate - 365 and sysdate "
                
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        
        If rsData.RecordCount > 0 Then
            stbThis.Panels(5).Text = "未入档特检数：" & Nvl(rsData!返回值)
        End If
    End If
    
End Sub


Private Sub QueryMaterialData()
'查询材料数据
    Dim strSql As String
    Dim strPatholNumQuery As String
    Dim strFilterDate As String
    Dim strRequestFrom As String
    Dim lngCurArchivesId As Long
    
    lngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))


    strRequestFrom = ""
    strFilterDate = ""
    
    If cbxQueryType.Text = "报到时间" Then
        strFilterDate = " and a.报到时间 between [1] and [2] "
    Else
        strFilterDate = " and c.申请ID=r.申请ID and r.申请时间 between [1] and [2]"
        strRequestFrom = " ,病理申请信息 r "
    End If


    strPatholNumQuery = ""
    If Trim(txtStartPatholNum.Text) <> "" And Trim(txtEndPatholNum.Text) <> "" Then
        strPatholNumQuery = " and (REGEXP_SUBSTR(upper(a.病理号), '[[:alpha:]]+') >=REGEXP_SUBSTR(upper([3]),'[[:alpha:]]+') and to_number(REGEXP_SUBSTR(upper(a.病理号), '[[:digit:]]+')) >=to_number(REGEXP_SUBSTR(upper([3]),  '[[:digit:]]+'))) "
        strPatholNumQuery = strPatholNumQuery & " and  (REGEXP_SUBSTR(upper(a.病理号), '[[:alpha:]]+') <=REGEXP_SUBSTR(upper([4]),'[[:alpha:]]+') and to_number(REGEXP_SUBSTR(upper(a.病理号),  '[[:digit:]]+')) <=to_number(REGEXP_SUBSTR(upper([4]), '[[:digit:]]+'))) "
    ElseIf Trim(txtStartPatholNum.Text) <> "" Then
        strPatholNumQuery = " and upper(a.病理号)=upper([3]) "
    ElseIf Trim(txtStartPatholNum.Text) <> "" Then
        strPatholNumQuery = " and upper(a.病理号) =upper([4]) "
    End If
    
    
    If mcurMaterialType <> amtMaterial Then
        '先查询包含申请的检查信息，在单独查询检查信息（便于对界面中的几种状态过滤）
        strSql = " select distinct a.病理医嘱ID, '' as 材料类别, 0 as 序号, 4 as 档案来源,  '4-' || a.病理医嘱ID as 来源ID,a.病理号, b.姓名,b.性别,b.年龄, " & _
                " b.医嘱内容 as 检查项目, a.检查类型, null as 材块号, null as 标本名称, null as 取材位置, null as 材料明细, null as 借阅状态, " & _
                " null as 数量, decode((select 档案ID from 病理归档信息 where 病理医嘱ID=a.病理医嘱ID and 档案ID=[5]),[5],'已存在','未归档') as 存放状态,  a.报到时间, 1 as 是否申请, " & _
                " decode(r.申请类型, 0, '免疫',1,'特染',2,'分子',3,'再制片',4,'补取材','') as 申请类型, null as 申请细目,c.执行过程  " & _
                " from 病理检查信息 a, 病人医嘱记录 b, 病人医嘱发送 c, 病理申请信息 r " & _
                " where a.医嘱ID=b.id and a.医嘱ID=c.医嘱ID and a.病理医嘱id=r.病理医嘱ID  " & _
                IIf(cbxQueryType.Text = "报到时间", " and a.报到时间 between [1] and [2]", " and r.申请时间 between [1] and [2]") & strPatholNumQuery & _
                " Union All  " & _
                " select distinct a.病理医嘱ID, '' as 材料类别, 0 as 序号, 4 as 档案来源,  '4-' || a.病理医嘱ID as 来源ID,a.病理号, b.姓名,b.性别,b.年龄,  " & _
                " b.医嘱内容 as 检查项目, a.检查类型, null as 材块号, null as 标本名称, null as 取材位置, null as 材料明细, null as 借阅状态, " & _
                " null as 数量, decode((select 档案ID from 病理归档信息 where 病理医嘱ID=a.病理医嘱ID and 档案ID=[5]), [5],'已存在','未归档') as 存放状态,  a.报到时间, 0 as 是否申请, null as 申请类型, null as 申请细目,c.执行过程  " & _
                " from 病理检查信息 a, 病人医嘱记录 b, 病人医嘱发送 c " & IIf(cbxQueryType.Text = "报到时间", "", " , 病理申请信息 r") & _
                " where a.医嘱ID=b.id and a.医嘱ID=c.医嘱ID  " & _
                IIf(cbxQueryType.Text = "报到时间", " and a.报到时间 between [1] and [2] ", " and a.病理医嘱id=r.病理医嘱ID and r.申请时间 between [1] and [2] ") & _
                strPatholNumQuery
    Else
        strFilterDate = strFilterDate & strPatholNumQuery
        
        strSql = "select  1 as 档案来源, '1-' || c.材块ID as 来源ID, a.病理医嘱id, a.病理号, b.姓名, b.性别, b.年龄,b.医嘱内容 as 检查项目, e.执行过程, " & _
                " a.检查类型, c.序号,c.取材位置,c.标本名称,to_number(c.蜡块数) as 数量, a.报到时间, '蜡块' as 材料类别, '' 申请细目," & _
                " case when c.申请ID is null then '常规取材' else '补取材' end as 申请类型, case when c.申请ID is null then '常规取材' else '补取材' end as 材料明细, " & _
                " case when d.存放状态 is null then '未归档' else decode(d.存放状态, 0, '存档中', 1, '部分遗失', '已遗失') end as 存放状态 , k.档案名称, k.详细地址, " & _
                " case when k.id is null then '' else '房间:' || k.所属房间 || ' 柜号:' || k.所属柜号 || ' 抽屉:' || k.所属抽屉 end as 存放位置 " & _
                " from  病理检查信息 a, 病人医嘱记录 b, 病理取材信息 c, 病理档案信息 k, 病理归档信息 d, 病人医嘱发送 e " & strRequestFrom & _
                " where a.医嘱id=b.id and a.病理医嘱id=c.病理医嘱id and b.ID=e.医嘱ID and k.id(+) = d.档案ID and c.材块id=d.材块id(+) and c.确认状态=1 and c.蜡块数>0 and a.检查类型<>3 " & strFilterDate & _
                " Union All select 2 as 档案来源, '2-' || c.ID as 来源ID, a.病理医嘱id, a.病理号, b.姓名, b.性别, b.年龄,b.医嘱内容 as 检查项目, f.执行过程, " & _
                " a.检查类型, d.序号,d.取材位置,d.标本名称,to_number(c.制片数) as 数量, a.报到时间, '切片' as 材料类别, " & _
                " decode(c.制片方式,0,'正常',1,'重切',2,'深切',3,'连切',4,'白片',5,'重染',6,'薄片','其他') 申请细目,case when c.申请ID is null then '常规制片' else '再制片' end as 申请类型, " & _
                " decode(c.制片方式,0,'正常',1,'重切',2,'深切',3,'连切',4,'白片',5,'重染',6,'薄片','其他') as 材料明细, " & _
                " case when e.存放状态 is null then '未归档' else decode(e.存放状态, 0, '存档中', 1, '部分遗失', '已遗失') end as 存放状态, k.档案名称, k.详细地址, " & _
                " case when k.id is null then '' else '房间:' || k.所属房间 || ' 柜号:' || k.所属柜号 || ' 抽屉:' || k.所属抽屉 end as 存放位置  " & _
                " from  病理检查信息 a, 病人医嘱记录 b, 病理制片信息 c, 病理取材信息 d, 病理档案信息 k, 病理归档信息 e, 病人医嘱发送 f " & strRequestFrom & _
                " where a.医嘱id=b.id and a.病理医嘱id = d.病理医嘱id and b.ID=f.医嘱ID and d.材块id=c.材块id and k.id(+) = e.档案ID and c.id=e.制片id(+) and c.当前状态=2 " & strFilterDate & _
                " Union All select 3 as 档案来源, '3-' || c.ID as 来源ID, a.病理医嘱id, a.病理号, b.姓名, b.性别, b.年龄,b.医嘱内容 as 检查项目, g.执行过程, " & _
                " a.检查类型, d.序号,d.取材位置,d.标本名称,1 as 数量, a.报到时间, decode(c.特检类型,0, '免疫',1,'特染',2,'分子') as 材料类别, " & _
                " decode(特检细目,1,'鉴别',2,'多耐药',3,'荧光',4,'普通') as 申请细目, decode(c.特检类型,0, '免疫',1,'特染',2,'分子') as 申请类型, " & _
                " decode(c.特检细目,0,decode(c.特检类型,0, '免疫',1,'特染',2,'分子'),1,'鉴别',2,'多耐药',3,'荧光',4,'普通') || '(' || f.抗体名称 || decode(c.制作类型,-1,'-补',0,'','-重' || c.制作类型) || ')' as 材料明细, " & _
                " case when e.存放状态 is null then '未归档' else decode(e.存放状态, 0, '存档中', 1, '部分遗失', '已遗失') end as 存放状态, k.档案名称, k.详细地址, " & _
                " case when k.id is null then '' else '房间:' || k.所属房间 || ' 柜号:' || k.所属柜号 || ' 抽屉:' || k.所属抽屉 end as 存放位置 " & _
                " from  病理检查信息 a, 病人医嘱记录 b, 病理特检信息 c, 病理取材信息 d, 病理档案信息 k, 病理归档信息 e, 病理抗体信息 f,病人医嘱发送 g  " & strRequestFrom & _
                " where a.医嘱id=b.id and a.病理医嘱id = d.病理医嘱id  and b.ID=g.医嘱ID and d.材块id=c.材块id and k.id(+) = e.档案ID and c.id=e.特检id(+) and c.抗体ID=f.抗体id and c.当前状态=2 " & strFilterDate
    End If
    
'    If mblnMoved Then
'        strSql = strSql & " union all " & GetMovedDataSql(strSql)
'    End If
    
    strSql = "select /*+ Rule*/ * from (" & strSql & ")  res  order by 材料类别,病理号,材料明细,序号"
    
    Set ufgMaterialQuery.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                            CDate(Format(dtpStartDate.value, "yyyy-mm-dd 00:00:00")), _
                                            CDate(Format(dtpEndDate.value, "yyyy-mm-dd 23:59:59")), _
                                            txtStartPatholNum.Text, _
                                            txtEndPatholNum.Text, _
                                            lngCurArchivesId _
                                            )
    
    Call FilterMaterialData
    
End Sub


Private Sub FilterMaterialData()
    Dim strFilter As String
    Dim strState As String
    Dim strIsRequest As String
    
    If ufgMaterialQuery.AdoData Is Nothing Then Exit Sub
    
    strFilter = ""
    strIsRequest = " 是否申请=0"
    
    If cbxStudyType.Text <> "" Then
        strFilter = strFilter & " 检查类型=" & cbxStudyType.ItemData(cbxStudyType.ListIndex)
    End If
    
    If cbxRequestType.Text <> "" Then
        If (cbxRequestType.Text = "常规取材" Or cbxRequestType.Text = "常规制片") And mcurMaterialType <> amtMaterial Then
        Else
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & " 申请类型='" & cbxRequestType.Text & "'"
            strIsRequest = ""
        End If
    End If
    
    
    If cbxRequestDetail.Text <> "" Then
        If strFilter <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & " 申请细目='" & cbxRequestDetail.Text & "'"
        strIsRequest = ""
    End If
    
    If Not (chkComplete.value = 0) Then
        If strFilter <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & " 执行过程=6"
    End If
    
    If Not (chkNotEnter.value = 0) Then
        If strFilter <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & " 存放状态='未归档'"
    End If
    
    
    strState = ""
    
    '当材料类型不为检查材料（蜡块，切片）时，则不会执行如下过滤条件
    If Not (chkWaxStone.value = 0) Then
        strState = strState & "(" & IIf(strFilter = "", "", strFilter & " and ") & " 材料类别='蜡块')"
    End If

    If Not (chkSlices.value = 0) Then
        If strState <> "" Then strState = strState & " or "
        strState = strState & "(" & IIf(strFilter = "", "", strFilter & " and ") & " 材料类别='切片')"
    End If

    If Not (chkTeShu.value = 0) Then
        If strState <> "" Then strState = strState & " or "
        strState = strState & "(" & IIf(strFilter = "", "", strFilter & " and ") & " 材料类别='免疫')"

        If strState <> "" Then strState = strState & " or "
        strState = strState & "(" & IIf(strFilter = "", "", strFilter & " and ") & " 材料类别='分子')"

        If strState <> "" Then strState = strState & " or "
        strState = strState & "(" & IIf(strFilter = "", "", strFilter & " and ") & " 材料类别='特染')"
    End If
    
    '当过滤的材料类型不是检查材料时，且没有使用申请类型和申请细目的过滤条件，则需要过滤出“是否申请”为0的数据显示
    If strIsRequest <> "" And mcurMaterialType <> amtMaterial Then
        strFilter = IIf(strFilter <> "", strFilter & " and " & strIsRequest, strIsRequest)
    End If
        
    ufgMaterialQuery.AdoData.Filter = IIf(strState = "", strFilter, strState)
    
    Call ufgMaterialQuery.RefreshData
End Sub


Private Sub cbxQueryType_KeyPress(KeyAscii As Integer)
'回车执行查询
On Error GoTo errHandle

    If KeyAscii = 13 Then
         '调用查询方法
         Call QueryMaterialData
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub dtpEndDate_Change()
'回车执行查询
On Error GoTo errHandle

    '调用查询方法
    Call QueryMaterialData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dtpStartDate_Change()
'回车执行查询
On Error GoTo errHandle

    '调用查询方法
    Call QueryMaterialData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub txtEndPatholNum_KeyPress(KeyAscii As Integer)
'回车执行查询
On Error GoTo errHandle

    If KeyAscii = 13 Then
         '调用查询方法
         Call QueryMaterialData
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtStartPatholNum_KeyPress(KeyAscii As Integer)
'回车执行查询
On Error GoTo errHandle

    If KeyAscii = 13 Then
         '调用查询方法
         Call QueryMaterialData
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cbxRequestDetail_Click()
On Error GoTo errHandle
    If Not cbxRequestDetail.Visible Then Exit Sub
    
    Call FilterMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbxRequestType_Click()
On Error GoTo errHandle
    If Not cbxRequestType.Visible Then Exit Sub
    
    Call FilterMaterialData
    
    labRequestDetail.Enabled = True
    cbxRequestDetail.Enabled = True
    Select Case cbxRequestType.Text
        Case "常规取材", "补取材", "常规制片", "特染"
            cbxRequestDetail.ListIndex = 0
            
            labRequestDetail.Enabled = False
            cbxRequestDetail.Enabled = False
        Case "再制片"
            Call cbxRequestDetail.Clear
            
            Call cbxRequestDetail.AddItem("")
            Call cbxRequestDetail.AddItem("重切")
            Call cbxRequestDetail.AddItem("深切")
            Call cbxRequestDetail.AddItem("连切")
            Call cbxRequestDetail.AddItem("白片")
            Call cbxRequestDetail.AddItem("重染")
            Call cbxRequestDetail.AddItem("薄片")
        Case "免疫"
            Call cbxRequestDetail.Clear
            
            Call cbxRequestDetail.AddItem("")
            Call cbxRequestDetail.AddItem("鉴别")
            Call cbxRequestDetail.AddItem("多耐药")
        Case "分子"
            Call cbxRequestDetail.Clear
            
            Call cbxRequestDetail.AddItem("")
            Call cbxRequestDetail.AddItem("荧光")
            Call cbxRequestDetail.AddItem("普通")
        Case Else
            Call ConfigRequestDetail
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbxStudyType_Click()
On Error GoTo errHandle
    If Not cbxStudyType.Visible Then Exit Sub
    
    Call FilterMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkComplete_Click()
On Error GoTo errHandle
    If Not chkComplete.Visible Then Exit Sub
    
    Call FilterMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkNotEnter_Click()
On Error GoTo errHandle
    If Not chkNotEnter.Visible Then Exit Sub
    
    Call FilterMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkSlices_Click()
On Error GoTo errHandle
    If Not chkSlices.Visible Then Exit Sub
    
    Call FilterMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkTeShu_Click()
On Error GoTo errHandle
    If Not chkTeShu.Visible Then Exit Sub
    
    Call FilterMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkWaxStone_Click()
On Error GoTo errHandle
    If Not chkWaxStone.Visible Then Exit Sub
    
    Call FilterMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function MaterailEnterArchives(ByVal lngArchivesId As Long) As String
'档案材料入档
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim dtServicesTime As Date
    Dim bFind As Boolean
    Dim strLog As String
    Dim strFormId As String
    
    strSql = "select ZL_病理档案_材料入档([1],[2],[3],[4]) as 返回值 from dual"
    
    dtServicesTime = zlDatabase.Currentdate
    
    strLog = ""
    For i = 1 To ufgMaterialQuery.GridRows - 1
        If ufgMaterialQuery.GetRowCheck(i) Then
        
            '先判断检查是否完成，只有已完成的检查才能进行入档操作
            If Val(ufgMaterialQuery.Text(i, gstrPatholCol_执行过程)) = 6 Then
                '如果材料类型为检查材料，则需要判断材料是否已经入档，已入档的材料不能再次入档
                If mcurMaterialType = amtMaterial And ufgMaterialQuery.Text(i, gstrPatholCol_存放状态) <> "未归档" Then
                    If strLog <> "" Then strLog = strLog & vbCrLf
                    strLog = strLog & "病理号为 [ " & ufgMaterialQuery.Text(i, gstrPatholCol_病理号) & _
                            " ] 材块号为 [ " & ufgMaterialQuery.Text(i, gstrPatholCol_材块号) & "] 的" & _
                            ufgMaterialQuery.Text(i, gstrPatholCol_材料明细) & ufgMaterialQuery.Text(i, gstrPatholCol_材料类别) & "已入档，不能再次入档。"
                ElseIf mcurMaterialType <> amtMaterial And ufgMaterialQuery.Text(i, gstrPatholCol_存放状态) <> "未归档" Then
                    If strLog <> "" Then strLog = strLog & vbCrLf
                    strLog = strLog & "病理号为 [ " & ufgMaterialQuery.Text(i, gstrPatholCol_病理号) & _
                            "] 的检查在该档案中已经存在，不能再次入档。"
                Else
                
                    strFormId = ufgMaterialQuery.Text(i, gstrPatholCol_来源ID)
                    strFormId = Mid(strFormId, InStr(strFormId, "-") + 1, 18)
            
                    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                        lngArchivesId, _
                                                        ufgMaterialQuery.Text(i, gstrPatholCol_病理医嘱ID), _
                                                        Val(ufgMaterialQuery.Text(i, gstrPatholCol_档案来源)), _
                                                        Val(strFormId))
    
    
                    If rsData.RecordCount <= 0 Then
                        Call err.Raise(0, "ExecuteArchivesFile", "未成功获取入档后的入档ID,处理失败。")
                        Exit Function
                    End If
                
                    If mcurMaterialType = amtMaterial Then
                        Call ufgMaterialQuery.SyncText(i, gstrPatholCol_存放状态, "存档中", True)
                    Else
                        Call ufgMaterialQuery.SyncText(i, gstrPatholCol_存放状态, "已存在", True)
                    End If
                End If
            Else
                If strLog <> "" Then strLog = strLog & vbCrLf
                strLog = strLog & "病理号为 [ " & ufgMaterialQuery.Text(i, gstrPatholCol_病理号) & " ]的检查尚未执行完成，不能进行入档操作。"
            End If
                                                               
        End If
    Next i
    
    MaterailEnterArchives = strLog
End Function


Private Function AllowDelArchivesMatierial(ByVal lngRow As Long) As String
'判断档案中的材料是否允许被移除
    AllowDelArchivesMatierial = ""
    
    If ufgArchives.Text(lngRow, gstrPatholCol_档案状态) = ArchivesState_Enter Then
        AllowDelArchivesMatierial = "档案已归档，不能从档案中移除材料。"
        Exit Function
    End If
    
End Function


Private Sub cmdDel_Click()
'删除档案材料
On Error GoTo errHandle
    Dim strInf As String
    
    If Not ufgArchives.IsSelectionRow Then
        Exit Sub
    End If
    
    '判断该档案是否允许移除（已遗失和已借阅的材料不能进行移除操作。）
    strInf = AllowDelArchivesMatierial(ufgArchives.SelectionRow)
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not ufgArchivesDetail.IsCheckedRow Then
        Call MsgBoxD(Me, "请选择需要移除的档案材料。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    If MsgBoxD(Me, "确认要从档案中移除所选择的材料吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    strInf = Execute_ClearArchivesMaterial
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
    End If
    
    
    Call RefreshStateInf(False, True)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cmdEnterArchives_Click()
'材料入档
On Error GoTo errHandle
    Dim strLog As String
    
    If mlngCurArchivesId <= 0 Then
        Call MsgBoxD(Me, "请选择所属的档案记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not ufgMaterialQuery.IsCheckedRow Then
        Call MsgBoxD(Me, "请选择需要进行入档的数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strLog = MaterailEnterArchives(Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID)))
    
    If strLog <> "" Then
        Call MsgBoxD(Me, strLog, vbOKOnly, Me.Caption)
    Else
        Call MsgBoxD(Me, "已完成入档操作。", vbOKOnly, Me.Caption)
    End If
    
    Call RefreshStateInf(False, True)

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub PrintArchives(ByVal lngArchivesId As Long, ByVal strReportName As String, Optional ByVal blnIsPrint As Boolean = True)
'打印档案内容
    Dim i As Long
    Dim j As Long
    Dim strValue(7) As String
    
    j = 0
    strValue(0) = "0": strValue(1) = "0": strValue(2) = "0": strValue(3) = "0": strValue(4) = "0": strValue(5) = "0": strValue(6) = "0": strValue(7) = "0"
    
    If mcurMaterialType = amdReport Then
        For i = 1 To ufgArchivesDetail.GridRows - 1
            If ufgArchivesDetail.GetRowCheck(i) Then
                If blnIsPrint Then
                    Call zlReport.ReportOpen(gcnOracle, 100, strReportName, Me, _
                        "档案ID=" & lngArchivesId, "归档ID=" & ufgArchivesDetail.KeyValue(i), 2)
                Else
                    '如果是预览，则只预览第一条选中的数据
                    Call zlReport.ReportOpen(gcnOracle, 100, strReportName, Me, _
                        "档案ID=" & lngArchivesId, "归档ID=" & ufgArchivesDetail.KeyValue(i), 1)
                        
                    Exit Sub
                End If
            End If
        Next i

    Else
        For i = 1 To ufgArchivesDetail.GridRows - 1
            If ufgArchivesDetail.GetRowCheck(i) Then
                If zlCommFun.ActualLen(strValue(j)) > 3000 Then
                    j = j + 1
                    strValue(j) = ""
                End If
    
                If strValue(j) <> "" Then strValue(j) = strValue(j) & ","
    
                strValue(j) = strValue(j) & ufgArchivesDetail.KeyValue(i)
            End If
        Next i
        
        Call zlReport.ReportOpen(gcnOracle, 100, strReportName, Me, _
            "档案ID=" & lngArchivesId, "归档ID1=" & strValue(0), "归档ID2=" & strValue(1), "归档ID3=" & strValue(2), "归档ID4=" & strValue(3), "归档ID5=" & strValue(4), "归档ID6=" & strValue(5), "归档ID7=" & strValue(6), "归档ID8=" & strValue(7), _
            IIf(blnIsPrint, 2, 1)) '1：预览，2：打印
    End If
    

End Sub


Private Sub cmdFilter_Click()
On Error GoTo errHandle
    Dim strFilter As String
    
    If ufgArchivesDetail.AdoData Is Nothing Then
        Call cmdRead_Click
    End If
    
    Call frmPatholArchivesLocate.ShowFilterWindow(Me)
    
    If Not frmPatholArchivesLocate.blnOk Then Exit Sub
        

    strFilter = " 报到时间>='" & Format(frmPatholArchivesLocate.dtpStart.value, "yyyy-mm-dd 00:00:00") & "' and 报到时间 <= '" & Format(frmPatholArchivesLocate.dtpEnd.value, "yyyy-mm-dd 23:59:59") & "'"


    If frmPatholArchivesLocate.txtName.Text <> "" Then
        strFilter = " 姓名 like '" & frmPatholArchivesLocate.txtName.Text & "*'"
    End If
    
    
    If frmPatholArchivesLocate.txtPatholNum.Text <> "" Then
        strFilter = " 病理号='" & frmPatholArchivesLocate.txtPatholNum.Text & "' or 病理号='" & UCase(frmPatholArchivesLocate.txtPatholNum.Text) & "'"
    End If
        

    
    ufgArchivesDetail.AdoData.Filter = strFilter
    
    Call ufgArchivesDetail.RefreshData
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdPreview_Click()
'打印档案内容
On Error GoTo errHandle
    If Not ufgArchives.IsSelectionRow Then Exit Sub
    
    If mlngCurArchivesId <= 0 Then
        mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    End If
    
    If ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_报表名称) = "" Then
        Call MsgBoxD(Me, "档案分类未设置对应报表，请在档案分类配置中设置对应的报表名称。", vbOKOnly, Me.Caption)
'        If MsgBoxD(Me, "档案分类未设置对应报表，请在档案分类配置中设置对应的报表名称。是否立即设置？", vbYesNo, Me.Caption) = vbNo Then
'            Exit Sub
'        Else
'            Call mnu_ArchivesClassCfg_Click
'        End If
    End If
    
    If Not ufgArchivesDetail.IsCheckedRow Then
        Call MsgBoxD(Me, "请选择需要预览的档案材料。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call PrintArchives(mlngCurArchivesId, ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_报表名称), False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdPrint_Click()
'打印档案内容
On Error GoTo errHandle
    If Not ufgArchives.IsSelectionRow Then Exit Sub
    
    If mlngCurArchivesId <= 0 Then
        mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    End If
    
    If ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_报表名称) = "" Then
        Call MsgBoxD(Me, "档案分类未设置对应报表，请在档案分类配置中设置对应的报表名称。", vbOKOnly, Me.Caption)
'        If MsgBoxD(Me, "档案分类未设置对应报表，请在档案分类配置中设置对应的报表名称。是否立即设置？", vbYesNo, Me.Caption) = vbNo Then
'            Exit Sub
'        Else
'            Call mnu_ArchivesClassCfg_Click
'
'        End If
    End If
    
    If Not ufgArchivesDetail.IsCheckedRow Then
        Call MsgBoxD(Me, "请选择需要打印的档案材料。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call PrintArchives(mlngCurArchivesId, ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_报表名称), True)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdQuery_Click()
'查询归档数据
On Error GoTo errHandle
    Call QueryMaterialData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdRead_Click()
On Error GoTo errHandle

    If Not ufgArchives.IsSelectionRow Then Exit Sub
    
    If mlngCurArchivesId <= 0 Then
        mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    End If
    
    Call LoadArchivesDetail(mlngCurArchivesId)
    
    If ufgArchivesDetail.AdoData.RecordCount <= 0 Then
        Call MsgBoxD(Me, "该档案尚不存在明细数据。", vbOKOnly, Me.Caption)
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadParameterConfig()
'载入相关参数配置
    mlngDefaultQueryDays = zlDatabase.GetPara("档案默认查询天数", glngSys, G_LNG_PATHOLARCHIVES_NUM, "30")
    mstrLabelReportName = zlDatabase.GetPara("档案标签报表名称", glngSys, G_LNG_PATHOLARCHIVES_NUM, "")
End Sub


Private Sub ConfigPopedomFace()
'更加权限配置界面，如果不具备权限时，则隐藏对应功能按钮
    Dim i As Long
    
    mnu_ParameterConfig.Visible = CheckPopedom(mstrPrivs, "参数设置")
    mnu_ArchivesClassCfg.Visible = CheckPopedom(mstrPrivs, "参数设置")
    
    mnu_CancelArchives.Visible = CheckPopedom(mstrPrivs, "撤销归档")
    
    For i = 1 To tbrTools.Buttons.Count
        If UCase(tbrTools.Buttons(i).Key) = UCase("tbn_CancelArchives") Then
            tbrTools.Buttons(i).Visible = CheckPopedom(mstrPrivs, "撤销归档")
        End If
    Next i
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Dim curDate As Date
    
'    #If DebugState = True Then
'        Call InitDebugObject(1295, Me, "zlhis", "HIS")
'    #End If
    mblnIsFormLoaded = False
    
    Call RestoreWinState(Me, App.ProductName)
    
'    Call InitCommandBars
    
    Call InitFace
    Call InitMenuIcoConfig
    
    Call InitArchivesFileList
    
    Call LoadParameterConfig
    
    Call ConfigStudyType
    Call ConfigRequestType
    Call ConfigRequestDetail
    
    Call SwitchArchivesFace(amtTable)
    
    curDate = zlDatabase.Currentdate
    
    cbxQueryType.ListIndex = 0
    dtpStartDate.value = curDate
    dtpEndDate.value = curDate
    mlngCurArchivesId = -1
    
    mstrPrivs = gstrPrivs
    
    Call ConfigPopedomFace
    
    Set zlReport = New zl9Report.clsReport
    
    
    Call QueryArchivesData(CDate(Format(curDate - mlngDefaultQueryDays, "yyyy-mm-dd 00:00:00")), CDate(Format(curDate, "yyyy-mm-dd 23:59:59")))
    
    Call RefreshStateInf(True, True)
    mblnIsFormLoaded = True
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


'Private Sub InitCommandBars()
'    '功能创建工具条
'    Dim cbrControl As CommandBarControl
'    Dim cbrPopControl As CommandBarControl
'    Dim cbrMenuBar As CommandBarPopup
'    Dim cbrToolBar As CommandBar
'    Dim cbrCustom As CommandBarControlCustom
'    Dim str3DFuncs() As String
'
'    Dim rsCollection As ADODB.Recordset
'    Dim rsViewShare As ADODB.Recordset
'    Dim rsShareCount As ADODB.Recordset
'    Dim rsTemp As ADODB.Recordset
'
'    Dim i As Integer
'    Dim i3DFunc As Integer
'
'    '-----------------------------------------------------
'    CommandBarsGlobalSettings.App = App
'    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
'    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
'
'    Me.cbrQuery.VisualTheme = xtpThemeOffice2003
'
'    Set Me.cbrQuery.Icons = zlCommFun.GetPubIcons
'    With Me.cbrQuery.Options
'        .ShowExpandButtonAlways = False
'        .ToolBarAccelTips = True
'        .AlwaysShowFullMenus = False
'        .IconsWithShadow = True '放在VisualTheme后有效
'        .UseDisabledIcons = True
'        .LargeIcons = True
'        .SetIconSize True, 24, 24
'    End With
'    Me.cbrQuery.EnableCustomization False
'    Me.cbrQuery.ActiveMenuBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
'
'
'
'    Set cbrCustom = cbrQuery.ActiveMenuBar.Controls.Add(xtpControlCustom, 1, "查询时间")
'        cbrCustom.Handle = cbxQueryType.hWnd
'        cbrCustom.flags = xtpFlagAlignLeft
'        cbrCustom.Style = xtpButtonIconAndCaption
'        cbrCustom.Category = "Main"
'
'    Set cbrCustom = cbrQuery.ActiveMenuBar.Controls.Add(xtpControlCustom, 2, "开始时间")
'        cbrCustom.Handle = dtpStartDate.hWnd
'        cbrCustom.flags = xtpFlagAlignLeft
'        cbrCustom.Style = xtpButtonIconAndCaption
'        cbrCustom.Category = "Main"
'
'    Call cbrQuery.ActiveMenuBar.Controls.Add(xtpControlLabel, 3, "到")
'
'    Set cbrCustom = cbrQuery.ActiveMenuBar.Controls.Add(xtpControlCustom, 4, "结束时间")
'        cbrCustom.Handle = dtpEndDate.hWnd
'        cbrCustom.flags = xtpFlagAlignLeft
'        cbrCustom.Style = xtpButtonIconAndCaption
'        cbrCustom.Category = "Main"
'
'
'
'    Call cbrQuery.ActiveMenuBar.Controls.Add(xtpControlLabel, 5, "病理号：")
'
'    Set cbrCustom = cbrQuery.ActiveMenuBar.Controls.Add(xtpControlCustom, 6, "开始病理号")
'        cbrCustom.Handle = txtStartPatholNum.hWnd
'        cbrCustom.flags = xtpFlagAlignLeft
'        cbrCustom.Style = xtpButtonIconAndCaption
'        cbrCustom.Category = "Main"
'
'    Call cbrQuery.ActiveMenuBar.Controls.Add(xtpControlLabel, 7, "到")
'
'    Set cbrCustom = cbrQuery.ActiveMenuBar.Controls.Add(xtpControlCustom, 8, "结束病理号")
'        cbrCustom.Handle = txtEndPatholNum.hWnd
'        cbrCustom.flags = xtpFlagAlignLeft
'        cbrCustom.Style = xtpButtonIconAndCaption
'        cbrCustom.Category = "Main"
'
'
'
'    Set cbrCustom = cbrQuery.ActiveMenuBar.Controls.Add(xtpControlCustom, 9, "查询按钮")
'        cbrCustom.Handle = cmdQuery.hWnd
'        cbrCustom.flags = xtpFlagAlignLeft
'        cbrCustom.Style = xtpButtonIconAndCaption
'        cbrCustom.Category = "Main"
'
'
''    Set cbrToolBar = Me.cbrQuery.Add("工具栏", xtpBarTop)
''    cbrToolBar.ShowTextBelowIcons = True
'
''    cbrToolBar.EnableDocking xtpFlagStretched '+ xtpFlagHideWrap
''    With cbrToolBar.Controls
''        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "查 询")
''            cbrControl.IconId = 814
''            cbrControl.ToolTipText = "查 询"
'
''        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印"): cbrControl.IconId = 103: cbrControl.ToolTipText = "报告打印"
''        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "登记"): cbrControl.BeginGroup = True: cbrControl.IconId = 211
''        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "报到"): cbrControl.IconId = 744
''
''        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
''        cbrControl.BeginGroup = True
''
''    End With
'
'End Sub


Private Sub ConfigStudyType()
'配置检查类型
    Call cbxStudyType.Clear

    Call cbxStudyType.AddItem("")
    cbxStudyType.ItemData(cbxStudyType.ListCount - 1) = -1
        
    Call cbxStudyType.AddItem("常规")
    cbxStudyType.ItemData(cbxStudyType.ListCount - 1) = 0
    
    Call cbxStudyType.AddItem("冰冻")
    cbxStudyType.ItemData(cbxStudyType.ListCount - 1) = 1
    
    Call cbxStudyType.AddItem("细胞")
    cbxStudyType.ItemData(cbxStudyType.ListCount - 1) = 2
    
    Call cbxStudyType.AddItem("会诊")
    cbxStudyType.ItemData(cbxStudyType.ListCount - 1) = 3
    
    Call cbxStudyType.AddItem("尸检")
    cbxStudyType.ItemData(cbxStudyType.ListCount - 1) = 4
    
    Call cbxStudyType.AddItem("快速石蜡")
    cbxStudyType.ItemData(cbxStudyType.ListCount - 1) = 5
    
    cbxStudyType.ListIndex = 0
End Sub


Private Sub ConfigRequestType()
'配置申请类型
    Call cbxRequestType.Clear
    
    Call cbxRequestType.AddItem("")
    
    Call cbxRequestType.AddItem("常规取材")
    Call cbxRequestType.AddItem("补取材")
    Call cbxRequestType.AddItem("常规制片")
    Call cbxRequestType.AddItem("再制片")
    Call cbxRequestType.AddItem("免疫")
    Call cbxRequestType.AddItem("特染")
    Call cbxRequestType.AddItem("分子")
    
    
    cbxRequestType.ListIndex = 0
End Sub


Private Sub ConfigRequestDetail()
'配置申请细目
    Call cbxRequestDetail.Clear
    
    Call cbxRequestDetail.AddItem("")
    
    Call cbxRequestDetail.AddItem("鉴别")
    Call cbxRequestDetail.AddItem("多耐药")
    
    Call cbxRequestDetail.AddItem("荧光")
    Call cbxRequestDetail.AddItem("普通")
    
    Call cbxRequestDetail.AddItem("重切")
    Call cbxRequestDetail.AddItem("深切")
    Call cbxRequestDetail.AddItem("连切")
    Call cbxRequestDetail.AddItem("白片")
    Call cbxRequestDetail.AddItem("重染")
    Call cbxRequestDetail.AddItem("薄片")
    
    cbxRequestDetail.ListIndex = 0
End Sub


Private Sub SwitchArchivesFace(ByVal amtMaterialType As TArchivesMaterialType)
'根据材料类型切换档案材料界面
'    If mcurMaterialType = amtMaterialType Then Exit Sub
    
    mcurMaterialType = amtMaterialType
    
    Call InitArchivesQueryList(amtMaterialType)
    Call InitArchivesDetailList(amtMaterialType)
    
    txtNumberInf.Visible = IIf(amtMaterialType = amtMaterial, True, False) And tabFilter.Selected.Index = 1
    
    labRequestDetail.Visible = IIf(amtMaterialType = amtMaterial, True, False)
    cbxRequestDetail.Visible = IIf(amtMaterialType = amtMaterial, True, False)
    lineSplit2.Visible = IIf(amtMaterialType = amtMaterial, True, False)
'    chkNotEnter.Visible = IIf(amtMaterialType = amtMaterial, True, False)
    chkWaxStone.Visible = IIf(amtMaterialType = amtMaterial, True, False)
    chkSlices.Visible = IIf(amtMaterialType = amtMaterial, True, False)
    chkTeShu.Visible = IIf(amtMaterialType = amtMaterial, True, False)
    
    
    cbxRequestType.ListIndex = 0
    cbxRequestDetail.ListIndex = 0
    chkNotEnter.value = 1 'IIf(mcurMaterialType = amtMaterial, 1, 0)
    chkComplete.value = 1
    chkWaxStone.value = 0
    chkSlices.value = 0
    chkTeShu.value = 0
    
    Call Picture2_Resize
End Sub


Private Sub InitArchivesFileList()
'初始化档案列表
    Dim strTemp As String
    

    
    ufgArchives.IsKeepRows = False
    ufgArchives.IsCopyMode = True
    
    strTemp = zlDatabase.GetPara("档案列表配置", glngSys, G_LNG_PATHOLARCHIVES_NUM, "")

    If strTemp = "" Then
        ufgArchives.ColNames = gstrArchivesManageCols
    Else
        ufgArchives.ColNames = strTemp
    End If
        '设置行数
    ufgArchives.GridRows = glngStandardRowCount
    '设置行高
    ufgArchives.RowHeightMin = glngStandardRowHeight
    ufgArchives.DefaultColNames = gstrArchivesManageCols
    ufgArchives.ColConvertFormat = gstrArchivesManageConvertFormat
End Sub


Private Sub InitArchivesQueryList(ByVal amtMaterialType As TArchivesMaterialType)
'初始化检查明显列表
'lngMaterialType:材料类型  0-文字材料，1-检查材料

    Dim strTemp As String

    
    strTemp = zlDatabase.GetPara(IIf(mcurMaterialType = amtMaterial, "档案材料查询列表配置", "档案纸质查询列表配置"), glngSys, G_LNG_PATHOLARCHIVES_NUM, "")

    ufgMaterialQuery.IsKeepRows = False
    If amtMaterialType <> amtMaterial Then
        ufgMaterialQuery.IsCopyMode = True
        ufgMaterialQuery.ColNames = IIf(strTemp <> "", strTemp, gstrArchivesWordCols)
        ufgMaterialQuery.DefaultColNames = gstrArchivesWordCols
        ufgMaterialQuery.ColConvertFormat = gstrArchivesWordConvertFormat
    Else
        ufgMaterialQuery.ColNames = IIf(strTemp <> "", strTemp, gstrArchivesMaterialCols)
        ufgMaterialQuery.DefaultColNames = gstrArchivesMaterialCols
        ufgMaterialQuery.ColConvertFormat = gstrArchivesMaterialConvertFormat
    End If
        
    '设置行数
    ufgMaterialQuery.GridRows = glngStandardRowCount
    '设置行高
    ufgMaterialQuery.RowHeightMin = glngStandardRowHeight
    Set ufgMaterialQuery.AdoData = Nothing
    Call ufgMaterialQuery.RefreshData
    
End Sub



Private Sub InitArchivesDetailList(ByVal amtMaterialType As TArchivesMaterialType)
'初始化检查明显列表
'lngMaterialType:材料类型  0-文字材料，1-检查材料
    Dim strTemp As String
    

    
    strTemp = zlDatabase.GetPara(IIf(mcurMaterialType = amtMaterial, "档案材料明细列表配置", "档案纸质明细列表配置"), glngSys, G_LNG_PATHOLARCHIVES_NUM, "")

    ufgArchivesDetail.IsKeepRows = False
    If amtMaterialType <> amtMaterial Then
        ufgArchivesDetail.ColNames = IIf(strTemp <> "", strTemp, gstrArchivesWordCols)
        ufgArchivesDetail.DefaultColNames = gstrArchivesWordCols
        ufgArchivesDetail.ColConvertFormat = gstrArchivesWordConvertFormat
    Else
        ufgArchivesDetail.ColNames = IIf(strTemp <> "", strTemp, gstrArchivesMaterialDetailCols)
        ufgArchivesDetail.DefaultColNames = gstrArchivesMaterialDetailCols
        ufgArchivesDetail.ColConvertFormat = gstrArchivesMaterialConvertFormat
    End If
        '设置行数
    ufgArchivesDetail.GridRows = glngStandardRowCount
    '设置行高
    ufgArchivesDetail.RowHeightMin = glngStandardRowHeight
    Set ufgArchivesDetail.AdoData = Nothing
    Call ufgArchivesDetail.RefreshData
End Sub


Private Sub QueryArchivesData(ByVal dtStartDate As Date, ByVal dtEndDate As Date, Optional ByVal lngArchivesClassId As Long, _
    Optional ByVal strArchivsName As String, Optional ByVal strArchivesCode As String)
'查询指定时间范围内的数据到档案
    Dim strSql As String
    
    
    mblnMoved = MovedByDate(dtStartDate)
    
    strSql = "select a.ID, a.档案名称, a.档案编号, a.检查范围, " & _
                " a.开始日期, a.结束日期, a.档案说明, a.档案状态, a.创建人, a.创建日期, b.分类名称 as 档案分类, B.材料类型,B.报表名称," & _
                " a.所属房间, a.所属柜号, a.所属抽屉, a.详细地址,a.归档时间 " & _
                " from 病理档案信息 a, 病理档案分类 b " & _
                " where a.分类ID=b.id  and a.创建日期 between [1] and [2] " & _
                IIf(lngArchivesClassId <= 0, "", " and a.分类ID=[3]") & _
                IIf(strArchivsName = "", "", " and upper(a.档案名称)=upper([4])") & _
                IIf(strArchivesCode = "", "", " and upper(a.档案编号)=upper([5])") & _
                " order by a.创建日期,a.档案名称 "
                   
    Set ufgArchives.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CDate(Format(dtStartDate, "yyyy-mm-dd 00:00:00")), _
                            CDate(Format(dtEndDate, "yyyy-mm-dd 23:59:59")), lngArchivesClassId, strArchivsName, strArchivesCode)
    
    Call ufgArchives.RefreshData
    
    Call ufgArchives.LocateRow(1)
    
    '读取附加说明信息
    If ufgArchives.IsSelectionRow Then
        Call ReadArchivesInf(ufgArchives.SelectionRow)
    End If
End Sub


Private Sub InitFace()
    With tabFilter
        .RemoveAll
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .RemoveAll
        

        .InsertItem 0, "资料入档", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "资料入档"
        
        
        .InsertItem 1, "档案明细", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).Tag = "档案明细"
        
        .Item(0).Selected = True
    End With
    
    framEnterArchives.Visible = True
End Sub



Private Sub AdjustLayOut()
    
    Picture1.Top = IIf(tbrTools.Visible, tbrTools.Top + tbrTools.Height, 0)
    Picture1.Height = Me.ScaleHeight - IIf(tbrTools.Visible, tbrTools.Height, 0) - IIf(stbThis.Visible, stbThis.Height, 120)
    
    Call ucSplitter1.RePaint
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Call AdjustLayOut
err.Clear
End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    Set zlReport = Nothing
err.Clear
End Sub

Private Sub mnu_About_Click()
'关于
On Error GoTo errHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_ArchivesClassCfg_Click()
On Error GoTo errHandle
    If Not CheckPopedom(mstrPrivs, "参数设置") Then
        Call MsgBoxD(Me, "不具备执行该操作的权限。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Dim frmArchivesClass As New frmPatholArchivesClass
    On Error GoTo errFree
        Call frmArchivesClass.Show(1, Me)
errFree:
        Call Unload(frmArchivesClass)
        Set frmArchivesClass = Nothing
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_BBS_Click()
'中联论坛
On Error GoTo errHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_CancelArchives_Click()
'撤销归档
On Error GoTo errHandle
    Call Execute_CancelEnterArchives
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_DelArchives_Click()
'删除档案
On Error GoTo errHandle
    Call Execute_DelArchives
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_EnterArchives_Click()
'档案归档
On Error GoTo errHandle
    Call Execute_EnterArchives
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Exit_Click()
'退出
On Error GoTo errHandle
    Call Execute_Exit
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub mnu_ExportExcel_Click()
'导处Excel
On Error GoTo errHandle
    Call MenuPrint(3)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Public Sub MenuPrint(intOutMode As Byte)
    '---------------------------------------------------
    '功能：    根据屏幕打印预览, 0弹出操作选择对话框，1预览，2打印，3导出Excel
    '参数：    输出方式
    '返回：
    '---------------------------------------------------
    Dim objPrint As New zlPrint1Grd

    Set objPrint.Body = ufgArchives.DataGrid
    
    objPrint.Title = "病理档案清单"

    If intOutMode = 0 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrView1Grd objPrint, 1
        Case 2
            zlPrintOrView1Grd objPrint, 2
        Case 3
            zlPrintOrView1Grd objPrint, 3
        Case Else
        End Select
    Else
        zlPrintOrView1Grd objPrint, intOutMode
    End If

End Sub



Private Sub mnu_Font_Click()
'字体
On Error GoTo errHandle
    
    diaFont.flags = 1
    diaFont.FontBold = ufgArchives.DataGrid.Font.Bold
    diaFont.FontName = ufgArchives.DataGrid.Font.Name
    diaFont.FontSize = ufgArchives.DataGrid.Font.Size
    diaFont.FontStrikethru = ufgArchives.DataGrid.Font.Strikethrough
    diaFont.FontUnderline = ufgArchives.DataGrid.Font.Underline
    
    diaFont.ShowFont
    
    '档案列表
    ufgArchives.DataGrid.Font.Bold = diaFont.FontBold
    ufgArchives.DataGrid.Font.Name = diaFont.FontName
    ufgArchives.DataGrid.Font.Size = diaFont.FontSize
    ufgArchives.DataGrid.Font.Strikethrough = diaFont.FontStrikethru
    ufgArchives.DataGrid.Font.Underline = diaFont.FontUnderline
    
    
    Call ufgArchives.DataGrid.Refresh
    
    ufgArchives.DataGrid.AutoSizeMode = flexAutoSizeColWidth
    Call ufgArchives.DataGrid.AutoSize(0, ufgArchives.DataGrid.Rows - 1)
    
    ufgArchives.DataGrid.AutoSizeMode = flexAutoSizeRowHeight
    Call ufgArchives.DataGrid.AutoSize(0, ufgArchives.DataGrid.Rows - 1)
    
    
    '查询列表
    ufgMaterialQuery.DataGrid.Font.Bold = diaFont.FontBold
    ufgMaterialQuery.DataGrid.Font.Name = diaFont.FontName
    ufgMaterialQuery.DataGrid.Font.Size = diaFont.FontSize
    ufgMaterialQuery.DataGrid.Font.Strikethrough = diaFont.FontStrikethru
    ufgMaterialQuery.DataGrid.Font.Underline = diaFont.FontUnderline
    
    Call ufgMaterialQuery.DataGrid.Refresh
    
    ufgMaterialQuery.DataGrid.AutoSizeMode = flexAutoSizeColWidth
    Call ufgMaterialQuery.DataGrid.AutoSize(0, ufgMaterialQuery.DataGrid.Rows - 1)
    
    ufgMaterialQuery.DataGrid.AutoSizeMode = flexAutoSizeRowHeight
    Call ufgMaterialQuery.DataGrid.AutoSize(0, ufgMaterialQuery.DataGrid.Rows - 1)
    
    
    '明细列表
    ufgArchivesDetail.DataGrid.Font.Bold = diaFont.FontBold
    ufgArchivesDetail.DataGrid.Font.Name = diaFont.FontName
    ufgArchivesDetail.DataGrid.Font.Size = diaFont.FontSize
    ufgArchivesDetail.DataGrid.Font.Strikethrough = diaFont.FontStrikethru
    ufgArchivesDetail.DataGrid.Font.Underline = diaFont.FontUnderline
    
    Call ufgArchivesDetail.DataGrid.Refresh
    
    ufgArchivesDetail.DataGrid.AutoSizeMode = flexAutoSizeColWidth
    Call ufgArchivesDetail.DataGrid.AutoSize(0, ufgArchivesDetail.DataGrid.Rows - 1)
    
    ufgArchivesDetail.DataGrid.AutoSizeMode = flexAutoSizeRowHeight
    Call ufgArchivesDetail.DataGrid.AutoSize(0, ufgArchivesDetail.DataGrid.Rows - 1)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Help_Click()
'帮助
On Error GoTo errHandle
    Call Execute_Help
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_LabPreview_Click()
'标签预览
On Error GoTo errHandle
    Call Execute_PrintArchivesLabel(False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_LabPrint_Click()
'标签打印
On Error GoTo errHandle
    Call Execute_PrintArchivesLabel(True)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub mnu_ListPreview_Click()
'预览数据列表
On Error GoTo errHandle
    Call MenuPrint(0)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_ListPrint_Click()
'打印数据列表
On Error GoTo errHandle
    Call MenuPrint(1)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_MainPage_Click()
'中联主页
On Error GoTo errHandle
    Call zlHomePage(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_NewArchives_Click()
'新增档案
On Error GoTo errHandle
    Call Execute_NewArchives
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_ParameterConfig_Click()
'参数配置
On Error GoTo errHandle
    If Not CheckPopedom(mstrPrivs, "参数设置") Then
        Call MsgBoxD(Me, "不具备执行该操作的权限。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call Execute_ParameterConfig
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_PrintConfig_Click()
'打印配置
On Error GoTo errHandle
    Call zlPrintSet
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_QueryArchives_Click()
'查询档案
On Error GoTo errHandle
    Call Execute_QueryArchives
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Return_Click()
'发送反馈
On Error GoTo errHandle
    Call zlMailTo(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_StandardBut_Click()
On Error GoTo errHandle
    Dim intCount As Long
    Me.mnu_StandardBut.Checked = Not Me.mnu_StandardBut.Checked
    Me.tbrTools.Visible = Me.mnu_StandardBut.Checked
    
    If Me.mnu_WordLabel.Checked Then
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = Me.tbrTools.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = ""
        Next
    End If

    Me.tbrTools.Refresh
    
    Form_Resize
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_StateBar_Click()
On Error GoTo errHandle
    Me.mnu_StateBar.Checked = Not Me.mnu_StateBar.Checked
    Me.stbThis.Visible = Me.mnu_StateBar.Checked
    
    Call Form_Resize
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub mnu_WordLabel_Click()
On Error GoTo errHandle
    Dim intCount As Long
    
    Me.mnu_WordLabel.Checked = Not Me.mnu_WordLabel.Checked

    If Me.mnu_WordLabel.Checked Then
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = Me.tbrTools.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = ""
        Next
    End If
    
    Me.tbrTools.Refresh
    
    Call Form_Resize
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Picture1_Resize()
On Error Resume Next
    Picture3.Left = 120
    Picture3.Top = 120
    Picture3.Width = Picture1.ScaleWidth - 120
    Picture3.Height = Picture1.ScaleHeight - 120
    
    Call ucSplitter2.RePaint
err.Clear
End Sub


Private Sub Picture2_Resize()
On Error Resume Next
    tabFilter.Left = 0
    tabFilter.Top = 0
    tabFilter.Width = Picture2.ScaleWidth
    
    framEnterArchives.Left = 0
    framEnterArchives.Top = tabFilter.Height
    framEnterArchives.Width = Picture2.ScaleWidth
    framEnterArchives.Height = Picture2.ScaleHeight - tabFilter.Height
    
    framQuery.Width = framEnterArchives.Width
    
    
    
    
    ufgMaterialQuery.Left = 120
    ufgMaterialQuery.Top = cbxStudyType.Top + cbxStudyType.Height + 120
    ufgMaterialQuery.Width = framEnterArchives.Width - 240
    ufgMaterialQuery.Height = framEnterArchives.Height - cbxStudyType.Top - chkComplete.Height - 840   '- cmdEnterArchives.Height
    
    
'    If cbxRequestDetail.Visible Then
'        chkComplete.Left = cbxRequestDetail.Left + cbxRequestDetail.Width + 120
'        chkComplete.Top = cbxRequestDetail.Top + 50
'    Else
'        chkComplete.Left = cbxRequestType.Left + cbxRequestType.Width + 120
'        chkComplete.Top = cbxRequestType.Top + 50
'    End If

    cmdEnterArchives.Left = 120 'framQuery.Width - cmdEnterArchives.Width - 120
    cmdEnterArchives.Top = ufgMaterialQuery.Top + ufgMaterialQuery.Height + 120

    If chkTeShu.Visible Then
        chkComplete.Left = ufgMaterialQuery.Width - chkComplete.Width * 5 - 600 '120
    Else
        chkComplete.Left = ufgMaterialQuery.Width - chkComplete.Width * 2 - 240
    End If
    
    chkComplete.Top = cmdEnterArchives.Top + 50
    
    chkNotEnter.Left = chkComplete.Left + chkComplete.Width + 120
    chkNotEnter.Top = chkComplete.Top
    
    lineSplit2.X1 = chkNotEnter.Left + chkNotEnter.Width + 120
    lineSplit2.X2 = lineSplit2.X1
    lineSplit2.Y1 = chkNotEnter.Top
    lineSplit2.Y2 = lineSplit2.Y1 + chkNotEnter.Height
    
    chkWaxStone.Left = lineSplit2.X1 + 120
    chkWaxStone.Top = chkComplete.Top
    
    chkSlices.Left = chkWaxStone.Left + chkWaxStone.Width + 120
    chkSlices.Top = chkComplete.Top
    
    chkTeShu.Left = chkSlices.Left + chkSlices.Width + 120
    chkTeShu.Top = chkComplete.Top
    
    
    
    '================================================================================
    
    framArchivesDetail.Left = 0
    framArchivesDetail.Top = tabFilter.Height
    framArchivesDetail.Width = Picture2.ScaleWidth
    framArchivesDetail.Height = Picture2.ScaleHeight - tabFilter.Height
    
    ufgArchivesDetail.Left = 120
    ufgArchivesDetail.Top = 240
    ufgArchivesDetail.Width = framArchivesDetail.Width - 240
    ufgArchivesDetail.Height = framArchivesDetail.Height - cmdDel.Height - 480
        
    cmdDel.Left = framArchivesDetail.Width - cmdDel.Width - 120
    cmdDel.Top = ufgArchivesDetail.Top + ufgArchivesDetail.Height + 120
    
    cmdPrint.Left = cmdDel.Left - cmdPrint.Width - 120
    cmdPrint.Top = cmdDel.Top
    
    cmdPreview.Left = cmdPrint.Left - cmdPreview.Width - 120
    cmdPreview.Top = cmdDel.Top
    
    cmdRead.Left = 120
    cmdRead.Top = cmdDel.Top
    
    cmdFilter.Left = cmdRead.Left + cmdRead.Width + 120
    cmdFilter.Top = cmdDel.Top
err.Clear
End Sub



Private Sub tabFilter_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errHandle
    
    framEnterArchives.Visible = IIf(Item.Index = 0, True, False)
    framArchivesDetail.Visible = IIf(Item.Index = 0, False, True)
    txtNumberInf.Visible = IIf(Item.Index = 0, False, True) And mcurMaterialType = amtMaterial
    
'    If Item.Index = 1 Then
'        If Not ufgArchives.IsSelectRow Then Exit Sub
'        If mlngCurArchivesId <= 0 Then mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectRowIndex, gstrArchivesManage_ID))
'
'        Call LoadArchivesDetail(mlngCurArchivesId)
'    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function AllowUpdateArchivesFile(ByVal lngDelRow As Long) As String
'判断是否允许更新档案
    AllowUpdateArchivesFile = ""
    
    If mblnMoved Then
        AllowUpdateArchivesFile = "数据已被转移，不能进行更新。"
        Exit Function
    End If
    
    If ufgArchives.Text(lngDelRow, gstrPatholCol_档案状态) <> "未归档" Then
        AllowUpdateArchivesFile = "档案已归档，不能进行更新。"
        Exit Function
    End If
End Function


Private Function AllowDelArchivesFile(ByVal lngDelRow As Long) As String
'判断是否允许删除档案
    AllowDelArchivesFile = ""
    
    If mblnMoved Then
        AllowDelArchivesFile = "数据已被转移，不能进行删除。"
        Exit Function
    End If
    
    If ufgArchives.Text(lngDelRow, gstrPatholCol_档案状态) <> "未归档" Then
        AllowDelArchivesFile = "档案已归档，不能进行删除。"
        Exit Function
    End If
    
'    If ufgStudy.ShowDataRows > 0 Then
'        AllowDelArchivesFile = "档案中包含检查数据，不能进行删除。"
'        Exit Function
'    End If
End Function


Private Sub DelArchivesFileData(lngArchivesId As Long)
'删除档案记录
    Dim strSql As String
    
    strSql = "Zl_病理档案_删除文件档案(" & lngArchivesId & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
End Sub


Private Sub tbrTools_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo errHandle
    
    Select Case UCase(Button.Key)
        Case UCase("tbn_LabView")   '预览档案标签
            Call Execute_PrintArchivesLabel(False)
            
        Case UCase("tbn_LabPrint")  '打印档案标签
            Call Execute_PrintArchivesLabel(True)
            
        Case UCase("tbn_NewArchives")   '新增档案
            Call Execute_NewArchives
    
        Case UCase("tbn_DelArchives")   '删除档案
            Call Execute_DelArchives
            
        Case UCase("tbn_UpdateArchives")    '更新档案
            Call Execute_UpdateArchives
                
        Case UCase("tbn_QueryArchives")     '查询档案
            Call Execute_QueryArchives
            
        Case UCase("tbn_EnterArchives")     '档案归档
            Call Execute_EnterArchives
            
        Case UCase("tbn_CancelArchives")    '撤销归档
            Call Execute_CancelEnterArchives
            
        Case UCase("tbn_Help")  '帮助
            Call Execute_Help
            
        Case UCase("tbn_Exit")  '推出档案管理模块
            Call Unload(Me)
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Execute_Exit()
'退出
    Call Unload(Me)
End Sub


Private Sub Execute_Help()
'帮助
    Shell "hh.exe  zl9start.chm", vbNormalFocus
End Sub


Private Sub Execute_QueryArchives()
'查询档案
On Error GoTo errHandle
    Dim strSql As String
    
    Call frmPatholArchivesQuery.ShowArchivesQueryWindow(mlngDefaultQueryDays, Me)
    
    If frmPatholArchivesQuery.mblnIsOk Then
        Call QueryArchivesData(frmPatholArchivesQuery.dtStartDate, frmPatholArchivesQuery.dtEndDate, _
            frmPatholArchivesQuery.lngArchivesClassId, frmPatholArchivesQuery.strArchivesName, frmPatholArchivesQuery.strArchivesCode)
    End If

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Execute_NewArchives()
'新增档案
    If Not frmPatholArchivesFileNew.ShowAddArchivesFileWindow(ufgArchives, Me) Then Exit Sub
    
    '读取档案附加显示信息
    If ufgArchives.IsSelectionRow Then
        Call ReadArchivesInf(ufgArchives.SelectionRow)
    End If
    
    Call RefreshStateInf(True, False)
End Sub


Private Sub Execute_DelArchives()
'删除档案
    Dim strInf As String
    
    '需要判断档案是否已经封存，且档案中不包含检查
    If Not ufgArchives.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要删除的档案记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mlngCurArchivesId <= 0 Then mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    
    strInf = AllowDelArchivesFile(ufgArchives.SelectionRow)
    
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "确认要删除选择的档案记录吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    
    Call DelArchivesFileData(mlngCurArchivesId)
    Call ufgArchives.DelRow(ufgArchives.SelectionRow, False, True)
    
    '读取档案附加显示信息
    If ufgArchives.SelectionRow Then
        Call ReadArchivesInf(ufgArchives.SelectionRow)
    Else
        '...
    End If
    
    Call RefreshStateInf(True, False)
End Sub


Private Sub Execute_UpdateArchives()
'更新档案
    Dim strInf As String
    
    If Not ufgArchives.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要更新的档案记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strInf = AllowUpdateArchivesFile(ufgArchives.SelectionRow)
    
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call frmPatholArchivesFileNew.ShowUpdateArchivesFileWindow(ufgArchives, Me)
    
    '读取档案附加显示信息
    If ufgArchives.IsSelectionRow Then
        Call ReadArchivesInf(ufgArchives.SelectionRow)
    End If
End Sub


Private Function ShowPlaceSureWindow(ByVal lngArchivesIndex As Long, ByRef strRoom As String, _
                                ByRef strBox As String, ByRef strDrawer As String) As Boolean
    Dim frmPlaceDialog As frmPatholArchivesPlaceDialog
    
    strRoom = ""
    strBox = ""
    strDrawer = ""
    
    On Error GoTo errFree:
        Set frmPlaceDialog = New frmPatholArchivesPlaceDialog
        
        Call frmPlaceDialog.ShowPlaceDialog(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_所属房间), _
                                        ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_所属柜号), _
                                        ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_所属抽屉), _
                                        Me)
        If frmPlaceDialog.IsOk Then
            strRoom = frmPlaceDialog.Room
            strBox = frmPlaceDialog.Box
            strDrawer = frmPlaceDialog.Drawer
        End If
        
        ShowPlaceSureWindow = frmPlaceDialog.IsOk
errFree:
    Call Unload(frmPlaceDialog)
    Set frmPlaceDialog = Nothing
End Function

Private Sub Execute_EnterArchives()
'执行档案归档操作
    Dim strRoom As String
    Dim strBox As String
    Dim strDrawer As String
    Dim curDate As Date
    
    If Not ufgArchives.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要归档的档案记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mblnMoved Then
        Call MsgBoxD(Me, "数据已被转移，不能执行该操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_档案状态) = ArchivesState_Enter Then
        Call MsgBoxD(Me, "档案已归档，不能进行归档处理。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not ShowPlaceSureWindow(ufgArchives.SelectionRow, strRoom, strBox, strDrawer) Then Exit Sub
    
    If strDrawer = "" And strBox = "" And strRoom = "" Then
        Call MsgBoxD(Me, "未选择档案存放位置，不能进行归档。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
        
    If mlngCurArchivesId <= 0 Then mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    
    '更新档案存放位置
    Call zlDatabase.ExecuteProcedure("ZL_病理档案_位置更新(" & mlngCurArchivesId & _
                                        ",'" & strRoom & "','" & strBox & "','" & strDrawer & "')", Me.Caption)
    '更新档案状态
    curDate = zlDatabase.Currentdate
    
    Call zlDatabase.ExecuteProcedure("Zl_病理档案_文件档案归档(" & mlngCurArchivesId & ",1," & To_Date(Format(curDate, "yyyy-mm-dd")) & ")", Me.Caption)
    
    ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_档案状态) = "已归档"
    ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_所属房间) = strRoom
    ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_所属柜号) = strBox
    ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_所属抽屉) = strDrawer
    ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_归档时间) = Format(curDate, "yyyy-mm-dd")

    Call ReadArchivesInf(ufgArchives.SelectionRow)
    
    Call ConfigArchivesModifyState(True)
    
    Call RefreshStateInf(True, False)
End Sub


Private Sub Execute_ParameterConfig()
'参数配置
    Dim frmParameter As frmPatholArchivesParameter
    
    Set frmParameter = New frmPatholArchivesParameter
On Error GoTo errFree
    Call frmParameter.ShowParameterWindow(mlngDefaultQueryDays, mstrLabelReportName, Me)
    
    mlngDefaultQueryDays = frmParameter.lngDefaultQueryDays
    mstrLabelReportName = frmParameter.strLabelReportName
errFree:
    Call Unload(frmParameter)
    Set frmParameter = Nothing
    
End Sub

Private Sub Execute_PrintArchivesLabel(ByVal blnIsAtOncePrint As Boolean)
'预览打印档案标签
On Error GoTo errHandle
    If Not ufgArchives.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要打印的档案记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mblnMoved Then
        Call MsgBoxD(Me, "数据已被转移，不能执行该操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Trim(mstrLabelReportName) = "" Then
        Call MsgBoxD(Me, "尚未配置标签对应的报表名称，请到“参数设置”中进行配置。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mlngCurArchivesId <= 0 Then
        mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    End If
        
    Call zlReport.ReportOpen(gcnOracle, 100, mstrLabelReportName, Me, "档案ID=" & mlngCurArchivesId, IIf(blnIsAtOncePrint, 2, 1)) '1：预览，2：打印
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Execute_CancelEnterArchives()
'执行档案撤销归档操作

    If Not ufgArchives.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要撤销归档的档案记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If mblnMoved Then
        Call MsgBoxD(Me, "数据已被转移，不能执行该操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_档案状态) = ArchivesState_NoEnter Then
        Call MsgBoxD(Me, "档案未归档，不能进行撤销处理。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "确认要对该档案进行撤销归档的操作吗？撤销归档后，档案相关信息将允许被修改。", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    If mlngCurArchivesId <= 0 Then mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    
    Call zlDatabase.ExecuteProcedure("Zl_病理档案_文件档案归档(" & mlngCurArchivesId & ",0,null)", Me.Caption)
    
    ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_档案状态) = "未归档"
    ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_归档时间) = ""
    
    Call ReadArchivesInf(ufgArchives.SelectionRow)
    
    Call ConfigArchivesModifyState(False)
    
    Call RefreshStateInf(True, False)
End Sub


Private Sub LoadArchivesDetail(ByVal lngArchivesId As Long)
    Dim strSql As String
    
    If lngArchivesId <= 0 Then Exit Sub
    
    If mcurMaterialType <> amtMaterial Then
        strSql = "select /*+ Rule*/ * from (" & _
                " select a.id as 来源ID, a.病理医嘱ID, 4 as 档案来源, b.病理号,c.姓名,c.性别,c.年龄,c.医嘱内容 as 检查项目, b.检查类型, " & _
                " decode(a.存放状态, 0, '存档中', 1, '部分遗失', '已遗失') as 存放状态, decode(a.借阅状态, 0, '未借出', 1, '部分借出', '已借出') as 借阅状态,a.借阅状态 as 排序,b.报到时间,null as 执行过程 " & _
                " from 病理归档信息 a,  病理检查信息 b, 病人医嘱记录 c " & _
                " Where a.资料来源 = 4 And a.病理医嘱id = b.病理医嘱id And b.医嘱ID = c.ID and 档案ID=[1]" & _
                " )order by 排序"
    Else
        strSql = "select /*+ Rule*/ * from (" & _
                " select a.id as 来源ID, a.病理医嘱ID, 1 as 档案来源, c.病理号,d.姓名,d.性别,d.年龄,d.医嘱内容 as 检查项目, c.检查类型, b.序号, b.标本名称, b.取材位置, '蜡块' as 材料类别, " & _
                " case when b.申请ID is null then '常规取材' else '补取材' end as 材料明细, b.蜡块数 as 数量, " & _
                " decode(a.存放状态, 0, '存档中', 1, '部分遗失', '已遗失') as 存放状态, decode(a.借阅状态, 0, '未借出', 1, '部分借出', '已借出') as 借阅状态,a.借阅状态 as 排序,c.报到时间, null as 执行过程 " & _
                " from 病理归档信息 a, 病理取材信息 b, 病理检查信息 c, 病人医嘱记录 d " & _
                " Where a.资料来源 = 1 And a.材块ID = b.材块ID And b.病理医嘱id = c.病理医嘱id And c.医嘱ID = d.ID and 档案ID=[1] " & _
            " Union All " & _
                " select a.id as 来源ID, a.病理医嘱ID, 2 as 档案来源, d.病理号,e.姓名,e.性别,e.年龄,e.医嘱内容 as 检查项目, d.检查类型, c.序号, c.标本名称, c.取材位置, '切片' as 材料类别, " & _
                " decode(b.制片方式,0,'正常',1,'重切',2,'深切',3,'连切',4,'白片',5,'重染',6,'薄片','其他') as 材料明细, b.制片数 as 数量, " & _
                " decode(a.存放状态, 0, '存档中', 1, '部分遗失', '已遗失') as 存放状态, decode(a.借阅状态, 0, '未借出', 1, '部分借出', '已借出') as 借阅状态,a.借阅状态 as 排序,d.报到时间, null as 执行过程 " & _
                " from 病理归档信息 a, 病理制片信息 b, 病理取材信息 c, 病理检查信息 d, 病人医嘱记录 e " & _
                " Where a.资料来源 = 2 And a.制片id = b.ID And b.材块ID = c.材块ID And c.病理医嘱id = d.病理医嘱id And d.医嘱ID = e.ID and 档案ID=[1] " & _
            " Union All " & _
                " select a.id as 来源ID, a.病理医嘱ID, 3 as 档案来源, d.病理号,e.姓名,e.性别,e.年龄,e.医嘱内容 as 检查项目, d.检查类型, c.序号, c.标本名称, c.取材位置, " & _
                " decode(b.特检类型,0, '免疫',1,'特染',2,'分子') as 材料类别, " & _
                " decode(b.特检细目,0,decode(b.特检类型,0, '免疫',1,'特染',2,'分子'),1,'鉴别',2,'多耐药',3,'荧光',4,'普通') || '(' || f.抗体名称 || decode(b.制作类型,-1,'-补',0,'','-重' || b.制作类型) || ')' as 材料明细, 1 as 数量, " & _
                " decode(a.存放状态, 0, '存档中', 1, '部分遗失', '已遗失') as 存放状态, decode(a.借阅状态, 0, '未借出', 1, '部分借出', '已借出') as 借阅状态,a.借阅状态 as 排序,d.报到时间, null as 执行过程 " & _
                " from 病理归档信息 a, 病理特检信息 b, 病理取材信息 c, 病理检查信息 d, 病人医嘱记录 e, 病理抗体信息 f " & _
                " Where a.资料来源 = 3 And a.特检id = b.ID And b.材块ID = c.材块ID And c.病理医嘱id = d.病理医嘱id And d.医嘱ID = e.ID And b.抗体id = f.抗体id  and 档案ID=[1]" & _
                " )order by 排序"
    End If
    
'    If mblnMoved Then
'        strSql = strSql & " Union all " & GetMovedDataSql(strSql)
'    End If
    
    Set ufgArchivesDetail.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngArchivesId)
    Call ufgArchivesDetail.RefreshData
End Sub


Private Sub ReadArchivesInf(ByVal lngArchivesRowIndex As Long)
'读取档案信息
    Dim strInf As String
    If lngArchivesRowIndex <= 0 Then Exit Sub
    
    strInf = "档案名称：" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_档案名称) & vbCrLf
    strInf = strInf & "档案编号：" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_档案编号) & vbCrLf
    strInf = strInf & "档案分类：" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_档案分类) & vbCrLf
    strInf = strInf & "检查范围：" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_检查范围) & vbCrLf
    strInf = strInf & "存放位置：[房间:" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_所属房间) & "  柜号:" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_所属柜号) & "  抽屉:" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_所属抽屉) & "]" & vbCrLf
    strInf = strInf & "详细地址：" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_详细地址) & vbCrLf
    strInf = strInf & "档案说明：" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_档案说明) & vbCrLf
    strInf = strInf & "档案状态：" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_档案状态) & vbCrLf
    
    strInf = strInf & "开始日期：" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_开始日期) & vbCrLf
    strInf = strInf & "结束日期：" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_结束日期) & vbCrLf
    strInf = strInf & "创 建 人：" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_创建人) & vbCrLf
    strInf = strInf & "创建日期：" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_创建日期) & vbCrLf
    strInf = strInf & "归档时间：" & ufgArchives.Text(lngArchivesRowIndex, gstrPatholCol_归档时间)
    
    rtbDetail.Text = strInf
End Sub



Private Function Execute_ClearArchivesMaterial() As String
'清除档案所包含的材料信息
    Dim i As Integer
    Dim strLog As String
    Dim blnAllowDel As Boolean
    
    strLog = ""
    For i = ufgArchivesDetail.GridRows - 1 To 1 Step -1
        If ufgArchivesDetail.GetRowCheck(i) Then
            blnAllowDel = True
            
            If ufgArchivesDetail.Text(i, gstrPatholCol_存放状态) <> "存档中" Then
                If strLog <> "" Then strLog = strLog & vbCrLf
                strLog = strLog & "病理号为 [ " & ufgArchivesDetail.Text(i, gstrPatholCol_病理号) & _
                                " ] 材块号为 [ " & ufgArchivesDetail.Text(i, gstrPatholCol_材块号) & "] 的" & _
                                ufgArchivesDetail.Text(i, gstrPatholCol_材料明细) & ufgArchivesDetail.Text(i, gstrPatholCol_材料类别) & "已发生遗失，不能从该档案中移除。"
                                
                blnAllowDel = False
            End If
            
            If ufgArchivesDetail.Text(i, gstrPatholCol_借阅状态) <> "未借出" And blnAllowDel Then
                If strLog <> "" Then strLog = strLog & vbCrLf
                strLog = strLog & "病理号为 [ " & ufgArchivesDetail.Text(i, gstrPatholCol_病理号) & _
                                " ] 材块号为 [ " & ufgArchivesDetail.Text(i, gstrPatholCol_材块号) & "] 的" & _
                                ufgArchivesDetail.Text(i, gstrPatholCol_材料明细) & ufgArchivesDetail.Text(i, gstrPatholCol_材料类别) & "已被借阅，不能从该档案中移除。"
                                
                blnAllowDel = False
            End If
        
            If blnAllowDel Then
                Call zlDatabase.ExecuteProcedure("ZL_病理档案_撤销入档(" & ufgArchivesDetail.Text(i, gstrPatholCol_来源ID) & ")", Me.Caption)
                
                '数据删除成功后移除界面中的数据
                Call ufgArchivesDetail.RemoveRow(i)
            End If
        End If
    Next i
    
    Execute_ClearArchivesMaterial = strLog
End Function


Private Sub ConfigArchivesModifyState(ByVal blnIsEnterArchives As Boolean)
'配置档案修改状态
'blnIsEnterArchives：是否归档(true：已归档, false：未归档)
    Dim i As Long

    For i = 1 To tbrTools.Buttons.Count
        Select Case UCase(tbrTools.Buttons(i).Key)
            Case UCase("tbn_DelArchives"), UCase("tbn_UpdateArchives"), UCase("tbn_EnterArchives")
                tbrTools.Buttons(i).Enabled = Not blnIsEnterArchives
        End Select
    Next i
'
'    tabFilter.Item(0).Enabled = Not blnIsEnterArchives
    
'    If blnIsEnterArchives Then
'        tabFilter.Item(1).Selected = blnIsEnterArchives
'    Else
'        tabFilter.Item(0).Selected = Not blnIsEnterArchives
'    End If
    
    mnu_DelArchives.Enabled = Not blnIsEnterArchives
    mnu_UpdateArchives.Enabled = Not blnIsEnterArchives
    mnu_EnterArchives.Enabled = Not blnIsEnterArchives
    
    cmdDel.Enabled = Not blnIsEnterArchives
    cmdEnterArchives.Enabled = Not blnIsEnterArchives

End Sub

Private Sub ConfigArchivesPrintState(ByVal blnIsValidReport As Boolean)
'配置报表打印按钮状态
'blnIsValidReport:是否有效报表（0：无效，1：有效）
    cmdPreview.Enabled = blnIsValidReport
    cmdPrint.Enabled = blnIsValidReport
End Sub



Private Sub ufgArchives_OnColFormartChange()
On Error GoTo errHandle
    zlDatabase.SetPara "档案列表配置", ufgArchives.GetColsString(ufgArchives), glngSys, G_LNG_PATHOLARCHIVES_NUM
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgArchives_OnColsNameReSet()
On Error GoTo errHandle
    Dim curDate As Date
    
    curDate = zlDatabase.Currentdate
    Call QueryArchivesData(CDate(Format(curDate - mlngDefaultQueryDays, "yyyy-mm-dd 00:00:00")), CDate(Format(curDate, "yyyy-mm-dd 23:59:59")))
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgArchives_OnSelChange()
On Error GoTo errHandle
    If ufgArchives.SelectionRow <= 0 Then
        mlngCurArchivesId = -1
        Exit Sub
    End If
    
    If ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID) = "" Then Exit Sub
    
    '单击时，如果档案ID相同，则不做任何处理
    If mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID)) Then Exit Sub
    
    mlngCurArchivesId = Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_ID))
    
    Call SwitchArchivesFace(Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_材料类型)))
        
'    If tabFilter.Selected.Index = 1 Then Call LoadArchivesDetail(mlngCurArchivesId)

    Call ReadArchivesInf(ufgArchives.SelectionRow)
    
    Call ConfigArchivesModifyState(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_档案状态) = "已归档")
    
    Call ConfigArchivesPrintState(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_报表名称) <> "")
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgArchivesDetail_OnColFormartChange()
'保存材料明细列表的配置
On Error GoTo errHandle
    If mblnIsFormLoaded Then
        zlDatabase.SetPara IIf(mcurMaterialType <> amtMaterial, "档案纸质明细列表配置", "档案材料明细列表配置"), ufgArchivesDetail.GetColsString(ufgArchivesDetail), glngSys, G_LNG_PATHOLARCHIVES_NUM
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgArchivesDetail_OnColsNameReSet()
On Error GoTo errHandle

    If ufgArchivesDetail.DataGrid.Rows > 1 Then Call LoadArchivesDetail(mlngCurArchivesId)

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgArchivesDetail_OnNewRow(ByVal Row As Long)
    '判断材料类型是非文字材料才进行 借阅状态的判断
    If Val(ufgArchives.Text(ufgArchives.SelectionRow, gstrPatholCol_材料类型)) = 1 Then
        If Nvl(ufgArchivesDetail.Text(Row, "借阅状态")) <> "未借出" Then
            Call ufgArchivesDetail.DisableCheck(Row, ufgArchivesDetail.GetColIndexWithRowCheck)
        End If
    End If
End Sub

Private Sub ufgArchivesDetail_OnSelChange()
On Error GoTo errHandle
    If Not ufgArchivesDetail.IsSelectionRow Then Exit Sub
    
    Call LoadMaterialDetialNumber(ufgArchivesDetail.Text(ufgArchivesDetail.SelectionRow, gstrPatholCol_来源ID))
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub LoadMaterialDetialNumber(ByVal lngMaterialArchivesId As Long)
'载入材料明细数量
'lngMaterialArchivesId:材料归档ID

    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    '只有档案类型为材料类型时，才读取档案数量信息
    If mcurMaterialType <> amtMaterial Then Exit Sub
    
    If Not txtNumberInf.Visible Then txtNumberInf.Visible = True
    
    strSql = "select a.ID, zl_病理材料_获取数量(a.ID) as 存档数量,  nvl(b.遗失数量, 0) as 遗失数量, nvl(c.已借数量, 0) as 已借数量  from 病理归档信息 a, " & _
             " (select nvl(sum(遗失数量),0) as 遗失数量, 归档ID from 病理遗失信息 where 归档ID=[1] group by 归档ID) b, " & _
             " (select (nvl(sum(借阅数量), 0) - nvl(sum(归还数量), 0)) as 已借数量, 归档ID " & _
             " From 病理借阅关联 where  归还状态=0  and 归档ID=[1] group by 归档ID) c " & _
             " where a.id =b.归档ID(+) and a.id=c.归档ID(+) and a.id = [1]"
             
'    If mblnMoved Then
'        strSql = "select sum(存档数量) as 存档数量, sum(遗失数量) as 遗失数量, sum(已借数量) as 已借数量 from (" & _
'                    strSql & " Union all" & GetMovedDataSql(strSql) & ") group by id"
'    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngMaterialArchivesId)
    
    txtNumberInf.Text = "当前材料数量：0   在档数量：0   已借数量：0   遗失数量：0"
    If rsData.RecordCount <= 0 Then Exit Sub
    
    txtNumberInf.Text = "当前材料数量：" & Nvl(rsData!存档数量) & _
                        "   在档数量：" & Val(Nvl(rsData!存档数量)) - Val(Nvl(rsData!遗失数量)) - Val(Nvl(rsData!已借数量)) & _
                        "   已借数量：" & Nvl(rsData!已借数量) & _
                        "   遗失数量：" & Nvl(rsData!遗失数量)
End Sub




Private Sub ufgMaterialQuery_OnColFormartChange()
'保存材料查询列表的配置
On Error GoTo errHandle
    If mblnIsFormLoaded Then
        zlDatabase.SetPara IIf(mcurMaterialType <> amtMaterial, "档案纸质查询列表配置", "档案材料查询列表配置"), ufgMaterialQuery.GetColsString(ufgMaterialQuery), glngSys, G_LNG_PATHOLARCHIVES_NUM
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

