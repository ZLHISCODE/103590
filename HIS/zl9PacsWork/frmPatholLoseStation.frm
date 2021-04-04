VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPatholLoseStation 
   Caption         =   "病理材料遗失工作站"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   -450
   ClientWidth     =   13530
   Icon            =   "frmPatholLoseStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   13530
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog diaFont 
      Left            =   6360
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   13335
      TabIndex        =   6
      Top             =   1680
      Width           =   13335
      Begin VB.TextBox txtStudyDate 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtStudyType 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtStudyItem 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   120
         Width           =   3375
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "检查日期："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   11160
         TabIndex        =   18
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "检查类型："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9240
         TabIndex        =   17
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "检查项目："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "年 龄："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "性 别："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "姓 名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2040
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":1042
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":1D1C
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":29F6
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":36D0
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":43AA
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":5084
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":5D5E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":6A38
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":7712
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":83EC
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":90C6
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgMenus 
      Left            =   2880
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":A118
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":A46A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":A7BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":AB3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":AE90
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":B1E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":B534
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":B886
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":BBD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":BF2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":C27C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":C5CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":C920
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":CC72
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":CFC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":D316
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":D668
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":D9BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":DD0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":E05E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":E3B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":E702
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":EA54
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":EDA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":F0F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":F44A
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":F79C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":FAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatholLoseStation.frx":FE40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame framQuery 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   13335
      Begin VB.OptionButton Option2 
         Caption         =   "遗失日期查询"
         Height          =   375
         Left            =   100
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "病理号查询"
         Height          =   375
         Left            =   4680
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   3240
         TabIndex        =   24
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   56623107
         CurrentDate     =   40928
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   300
         Left            =   1560
         TabIndex        =   22
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   56623107
         CurrentDate     =   40928
      End
      Begin VB.CheckBox chkMaterial 
         Caption         =   "特检材料"
         Height          =   180
         Index           =   2
         Left            =   10920
         TabIndex        =   21
         Top             =   330
         Width           =   1095
      End
      Begin VB.CheckBox chkMaterial 
         Caption         =   "切片材料"
         Height          =   180
         Index           =   1
         Left            =   9840
         TabIndex        =   20
         Top             =   330
         Width           =   1095
      End
      Begin VB.CheckBox chkMaterial 
         Caption         =   "蜡块材料"
         Height          =   180
         Index           =   0
         Left            =   8760
         TabIndex        =   19
         Top             =   330
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查 询(&Q)"
         Height          =   400
         Left            =   7440
         TabIndex        =   5
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox txtPatholNo 
         Height          =   300
         Left            =   6000
         TabIndex        =   4
         Top             =   280
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "到"
         Height          =   255
         Left            =   2985
         TabIndex        =   23
         Top             =   330
         Width           =   255
      End
   End
   Begin MSComctlLib.Toolbar tbrTools 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   1402
      ButtonWidth     =   1455
      ButtonHeight    =   1349
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "预览"
            Key             =   "tbn_PreviewList"
            Object.Tag             =   "预览"
            ImageKey        =   "IMG1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "打印"
            Key             =   "tbn_PrintList"
            Object.Tag             =   "打印"
            ImageKey        =   "IMG2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "材料遗失"
            Key             =   "tbn_NewLose"
            Object.Tag             =   "材料遗失"
            ImageKey        =   "IMG4"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "材料找回"
            Key             =   "tbn_FindLose"
            Object.Tag             =   "材料找回"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "帮助"
            Key             =   "tbn_Help"
            Object.Tag             =   "帮助"
            ImageKey        =   "IMG9"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出"
            Key             =   "tbn_Exit"
            Object.Tag             =   "退出"
            ImageKey        =   "IMG10"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin zl9PACSWork.ucFlexGrid ufgLose 
      Height          =   6570
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   11589
      DefaultCols     =   ""
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   7575
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   1764
            Picture         =   "frmPatholLoseStation.frx":10192
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "遗失材料数量："
            TextSave        =   "遗失材料数量："
            Key             =   "sb_NoEnter"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   14237
            MinWidth        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   18
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
      Begin VB.Menu mnu_PrintConfig 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnu_Split1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Preview 
         Caption         =   "预览(&V)"
      End
      Begin VB.Menu mnu_Print 
         Caption         =   "打印(&P)"
      End
      Begin VB.Menu mnu_Split2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ExportExcel 
         Caption         =   "输出到Excel(&E)"
      End
      Begin VB.Menu mnu_Split5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "退出(&Q)"
      End
   End
   Begin VB.Menu mnu_Edit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnu_MaterialLose 
         Caption         =   "材料遗失(&E)"
      End
      Begin VB.Menu mnu_MaterialFind 
         Caption         =   "材料找回(&F)"
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnu_ToolsBar 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnu_StandardButton 
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
      Begin VB.Menu mnu_Split3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Font 
         Caption         =   "字体(&F)"
      End
   End
   Begin VB.Menu mnu_Tools 
      Caption         =   "工具(&T)"
      Visible         =   0   'False
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnu_HelpMain 
         Caption         =   "帮助主题(&H)"
      End
      Begin VB.Menu mnu_WebZl 
         Caption         =   "WEB上的中联(&W)"
         Begin VB.Menu mnu_HomePage 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnu_BBS 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnu_back 
            Caption         =   "发送反馈(&K)"
         End
      End
      Begin VB.Menu mnu_Split4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_About 
         Caption         =   "关于...(&A)"
      End
   End
End
Attribute VB_Name = "frmPatholLoseStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



#Const DebugState = False


'为菜单设置相应的图形
Private Const MF_BITMAP = &H400&

Private Enum TQueryWay
    qwLoseDate = 0
    qwPatholNum = 1
End Enum


Dim WithEvents zlReport As zl9Report.clsReport
Attribute zlReport.VB_VarHelpID = -1


Private mstrPrivs As String
Private mblnMoved As Boolean

Private mqwQueryWay As TQueryWay
Private mstrCurSelectPatholNum As String

Private Sub InitMenuIcoConfig()
'初始化菜单图标显示
On Error Resume Next
    Dim hMenu As Long
    Dim hSubMenu As Long
    Dim hSubSubMenu As Long
    
    hMenu = GetMenu(Me.hWnd)
    
    '设置第一项菜单(文件)
    hSubMenu = GetSubMenu(hMenu, 0) '取得第一项菜单的子菜单句柄
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(3).Picture, imgMenus.ListImages(3).Picture) '打印设置
    Call SetMenuItemBitmaps(hSubMenu, 2, MF_BITMAP, imgMenus.ListImages(18).Picture, imgMenus.ListImages(18).Picture) '打印预览
    Call SetMenuItemBitmaps(hSubMenu, 3, MF_BITMAP, imgMenus.ListImages(19).Picture, imgMenus.ListImages(19).Picture) '打印
    Call SetMenuItemBitmaps(hSubMenu, 5, MF_BITMAP, imgMenus.ListImages(4).Picture, imgMenus.ListImages(4).Picture) '导出Excel
    Call SetMenuItemBitmaps(hSubMenu, 7, MF_BITMAP, imgMenus.ListImages(5).Picture, imgMenus.ListImages(5).Picture) '退出
    

    '设置第二项菜单（编辑）
    hSubMenu = GetSubMenu(hMenu, 1) '取得第二项菜单的子菜单句柄
    
    Call SetMenuItemBitmaps(hSubMenu, 0, MF_BITMAP, imgMenus.ListImages(7).Picture, imgMenus.ListImages(7).Picture) '材料遗失
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BITMAP, imgMenus.ListImages(10).Picture, imgMenus.ListImages(10).Picture) '材料找回
    
    
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


Private Sub ConfigPopedomFace()
'更加权限配置界面，如果不具备权限时，则隐藏对应功能按钮
    Dim i As Long
    
    mnu_MaterialFind.Visible = CheckPopedom(mstrPrivs, "材料找回")
    
    mnu_MaterialLose.Visible = CheckPopedom(mstrPrivs, "材料遗失")
    
    For i = 1 To tbrTools.Buttons.Count
        If UCase(tbrTools.Buttons(i).Key) = UCase("tbn_NewLose") Then
            tbrTools.Buttons(i).Visible = CheckPopedom(mstrPrivs, "材料遗失")
            
        End If
    Next i
End Sub




Private Sub Form_Load()
On Error GoTo ErrHandle
    Dim curDate As Date
'    #If DebugState = True Then
'        Call InitDebugObject(1294, Me, "zlhis", "HIS")
'    #End If
    
    Call RestoreWinState(Me, App.ProductName)
    
    mstrPrivs = gstrPrivs
    
    Call InitMenuIcoConfig
    
    Call InitLoseList
    
    Call RefreshStateInf
    
    curDate = zlDatabase.Currentdate
    
    dtpStart.value = Format(DateAdd("m", -1, curDate), "yyyy-mm-dd 00:00:00")
    dtpEnd.value = Format(curDate, "yyyy-mm-dd 23:59:59")
    
    mqwQueryWay = qwPatholNum
    mstrCurSelectPatholNum = ""
Exit Sub
ErrHandle:
If ErrCenter() = 1 Then Resume
End Sub



Private Sub RefreshStateInf()
'刷新材料遗失数量
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select sum(遗失数量) as 返回值 from 病理遗失信息"
            
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If rsData.RecordCount > 0 Then
        stbThis.Panels(2).Text = "遗失材料数量：" & Nvl(rsData!返回值)
    End If
End Sub


Private Sub chkMaterial_Click(Index As Integer)
'过滤不同类别的材料
On Error GoTo ErrHandle
    Dim strFilter As String
    
    strFilter = ""
    If chkMaterial(0).value <> 0 Then
        strFilter = " 材料类别='蜡块'"
    End If
    
    If chkMaterial(1).value <> 0 Then
        If strFilter <> "" Then strFilter = strFilter & " or "
        strFilter = strFilter & " 材料类别='切片'"
    End If
    
    If chkMaterial(2).value <> 0 Then
        If strFilter <> "" Then strFilter = strFilter & " or "
        strFilter = strFilter & " 材料类别='免疫' or 材料类别='分子' or 材料类别='特染' "
    End If
    
    If ufgLose.AdoData Is Nothing Then Exit Sub
    
    ufgLose.AdoData.Filter = strFilter
    
    Call ufgLose.RefreshData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub Form_Resize()
On Error Resume Next
    framQuery.Top = IIf(tbrTools.Visible, tbrTools.Top + tbrTools.Height, 0)
    framQuery.Left = 120
    framQuery.Width = Me.ScaleWidth - 240
    
    picInfo.Top = framQuery.Top + framQuery.Height
    picInfo.Left = 120
    picInfo.Width = Me.ScaleWidth - 240
    
    ufgLose.Top = picInfo.Top + picInfo.Height
    ufgLose.Left = 120
    ufgLose.Width = Me.ScaleWidth - 240
    ufgLose.Height = Me.ScaleHeight - framQuery.Height - picInfo.Height - IIf(tbrTools.Visible, tbrTools.Height, 0) - IIf(stbThis.Visible, stbThis.Height, 120)
err.Clear
End Sub


Private Sub cmdQuery_Click()
'查询材料
On Error GoTo ErrHandle
    mblnMoved = MovedByDate(dtpStart.value)
    
    If txtPatholNo.Enabled Then
        mqwQueryWay = qwPatholNum
        
        Call QueryStudyInf(txtPatholNo.Text)
        
        Call ufgLose.ClearListData
        
        If txtPatholNo.Text = "" Then Exit Sub
    Else
        mqwQueryWay = qwLoseDate
        
        Call QueryStudyInf("")
    End If
    
    Call QueryPatholMaterialData(txtPatholNo.Text, Format(dtpStart.value, "yyyy-mm-dd 23:59:59"), Format(dtpEnd.value, "yyyy-mm-dd 23:59:59"))
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub DTPEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdQuery.SetFocus
    End If
End Sub

Private Sub dtpStart_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        dtpEnd.SetFocus
    End If
    
End Sub

Private Sub QueryStudyInf(ByVal strPatholNum As String)
'查询检查基本信息
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    If Trim(strPatholNum) = "" Then Exit Sub
    
    strSQL = "select b.姓名, b.性别,b.年龄,b.医嘱内容,decode(a.检查类型,0,'常规',1,'冰冻',2,'细胞',3,'会诊',4,'尸检','快速石蜡') as 检查类型,a.报到时间 " & _
            " from 病人医嘱记录 b , 病理检查信息 a where a.医嘱ID=b.Id and a.病理号=[1]"
            
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatholNum)
    If rsData.RecordCount <= 0 Then Exit Sub
    
    txtName.Text = Nvl(rsData!姓名)
    txtSex.Text = Nvl(rsData!性别)
    txtAge.Text = Nvl(rsData!年龄)
    txtStudyItem.Text = Nvl(rsData!医嘱内容)
    txtStudyType.Text = Nvl(rsData!检查类型)
    txtStudyDate.Text = Format(Nvl(rsData!报到时间), "yyyy-mm-dd")
End Sub


'
'Private Sub QueryPatholMaterialDataByDate(ByVal dtStartDate As Date, ByVal dtEndDate As Date)
''查询病理材料
'    Dim strSql As String
'    Dim strLinkTable As String
'
'
'    Call ufgLose.ClearDataList
'
'    '统计遗失的材料数量（统计借阅数量时，只能统计未归还的借阅数量，部分规划或已遗失的材料将进行遗失处理，并体现到遗失数量中）
'    strLinkTable = " (select nvl(sum(遗失数量),0) as 遗失数量, 归档ID " & _
'                    " from 病理遗失信息 Where 遗失日期 between [1] and [2] group by 归档ID ) x, " & _
'                    " (select (nvl(sum(借阅数量), 0) - nvl(sum(归还数量), 0)) as 已借数量, a.归档ID " & _
'                    " from 病理借阅关联 a where a.归还状态=0 and  a.归档ID in(select 归档ID from 病理遗失信息 where 遗失日期 between [1] and [2]) " & _
'                    " group by a.归档ID" & ") y"
'
'
'
'    strSql = "select distinct d.检查类型, d.病理号, a.id, c.序号, c.标本名称, c.取材位置, '蜡块' as 材料类别, " & _
'            " case when c.申请ID is null then '常规取材' else '补取材' end as 材料明细, " & _
'            " (c.蜡块数 - nvl(x.遗失数量,0)  - nvl(y.已借数量, 0) ) as 在档数量, nvl(x.遗失数量,0) as 遗失数量, nvl(y.已借数量,0) as 已借数量, a.存放状态,a.借阅状态, " & _
'            " f.档案名称, '房间:' || f.所属房间 || ' 柜号:' || f.所属柜号 || ' 抽屉:' || f.所属抽屉 as 存放位置, f.详细地址 " & _
'            " from 病理档案信息 f, 病理检查信息 d, 病理取材信息 c, 病理归档信息 a, 病理遗失信息 g," & strLinkTable & _
'            " where f.id=a.档案id and d.病理医嘱id=c.病理医嘱id and c.材块id=a.材块id and a.id = x.归档ID(+) and a.id=y.归档id(+) and a.ID=g.归档ID and g.遗失日期 between [1] and [2] " & _
'        " Union All " & _
'            " select distinct d.检查类型, d.病理号, a.id, c.序号, c.标本名称, c.取材位置, '切片' as 材料类别, " & _
'            " decode(b.制片方式,0,'正常',1,'重切',2,'深切',3,'连切',4,'白片',5,'重染',6,'薄片','其他') as 材料明细, " & _
'            " (b.制片数 - nvl(x.遗失数量,0)  - nvl(y.已借数量, 0)) as 在档数量, nvl(x.遗失数量,0) as 遗失数量, nvl(y.已借数量,0) as 已借数量, a.存放状态,a.借阅状态, " & _
'            " e.档案名称, '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置, e.详细地址 " & _
'            " from 病理档案信息 e, 病理检查信息 d, 病理取材信息 c, 病理制片信息 b, 病理归档信息 a, 病理遗失信息 f," & strLinkTable & _
'            " where e.id = a.档案id and  d.病理医嘱id=c.病理医嘱id and c.材块id=b.材块id and b.id=a.制片id and a.id = x.归档ID(+) and a.id=y.归档id(+) and a.ID=f.归档ID and f.遗失日期 between [1] and [2] " & _
'        " Union All " & _
'            " select distinct d.检查类型, d.病理号,  a.id, c.序号, c.标本名称, c.取材位置, decode(b.特检类型,0, '免疫',1,'特染',2,'分子') as 材料类别, " & _
'            " decode(b.特检细目,0,decode(b.特检类型,0, '免疫',1,'特染',2,'分子'),1,'鉴别',2,'多耐药',3,'荧光',4,'普通') || '(' || f.抗体名称 || decode(b.制作类型,-1,'-补',0,'','-重' || b.制作类型) || ')' as 材料明细, " & _
'            " (1 - nvl(x.遗失数量,0)  - nvl(y.已借数量, 0)) as 在档数量, nvl(x.遗失数量,0) as 遗失数量, nvl(y.已借数量,0) as 已借数量, a.存放状态,a.借阅状态, " & _
'            " e.档案名称, '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置, e.详细地址 " & _
'            " from 病理档案信息 e, 病理抗体信息 f, 病理检查信息 d, 病理取材信息 c, 病理特检信息 b, 病理归档信息 a, 病理遗失信息 g, " & strLinkTable & _
'            " where e.id = a.档案id and f.抗体ID=b.抗体ID and d.病理医嘱id=c.病理医嘱id and c.材块id=b.材块id and b.id=a.特检id and  a.id = x.归档ID(+) and a.id=y.归档id(+) " & _
'            " and a.ID=g.归档ID and  g.遗失日期 between [1] and [2] "
'
'    If mblnMoved Then
'        strSql = strSql & " Union All " & GetMovedDataSql(strSql)
'    End If
'
'    strSql = "select /*+RULE*/ * from ( " & strSql & ") order by 材料类别, 病理号,序号,材料明细,存放状态"
'
'    Set ufgLose.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate)
'
'    Call ufgLose.RefreshData
'
'
'    If ufgLose.AdoData.RecordCount <= 0 Then
'        Call MsgBoxD(Me, "未查询到相关数据。", vbOKOnly, Me.Caption)
'        Exit Sub
'    End If
'End Sub



'Private Sub QueryPatholMaterialDataByPatholNum(ByVal strPatholNum As String)
''查询病理材料
'    Dim strSql As String
'    Dim strLinkTable As String
'
'
'    Call ufgLose.ClearDataList
'
'    If Trim(txtPatholNo.Text) = "" Then Exit Sub
'
'    '统计遗失的材料数量（统计借阅数量时，只能统计未归还的借阅数量，部分规划或已遗失的材料将进行遗失处理，并体现到遗失数量中）
'    strLinkTable = " (select nvl(sum(遗失数量),0) as 遗失数量, 归档ID " & _
'                    " from 病理遗失信息 a, 病理归档信息 b, 病理检查信息 d Where a.归档ID = b.ID And b.病理医嘱id = d.病理医嘱id " & _
'                    " and d.病理号=[1] group by 归档ID ) x, " & _
'                    " (select (nvl(sum(借阅数量), 0) - nvl(sum(归还数量), 0)) as 已借数量, 归档ID " & _
'                    " from 病理借阅关联 a, 病理归档信息 b, 病理检查信息 d where a.归档ID = b.ID And b.病理医嘱id = d.病理医嘱id " & _
'                    " and a.归还状态=0  and d.病理号=[1] group by 归档ID" & ") y"
'
'
'
'    strSql = "select * from (select d.检查类型, d.病理号, a.id, c.序号, c.标本名称, c.取材位置, '蜡块' as 材料类别, " & _
'            " case when c.申请ID is null then '常规取材' else '补取材' end as 材料明细, " & _
'            " (c.蜡块数 - nvl(x.遗失数量,0)  - nvl(y.已借数量, 0) ) as 在档数量, nvl(x.遗失数量,0) as 遗失数量, nvl(y.已借数量,0) as 已借数量, a.存放状态,a.借阅状态, " & _
'            " f.档案名称, '房间:' || f.所属房间 || ' 柜号:' || f.所属柜号 || ' 抽屉:' || f.所属抽屉 as 存放位置, f.详细地址 " & _
'            " from 病理归档信息 a, 病理取材信息 c, 病理检查信息 d, 病理档案信息 f, " & strLinkTable & _
'            " where a.材块id=c.材块id and c.病理医嘱id=d.病理医嘱id and a.id = x.归档ID(+) and a.id=y.归档id(+) and a.档案id=f.id  and d.病理号=[1] " & _
'        " Union All " & _
'            " select d.检查类型, d.病理号, a.id, c.序号, c.标本名称, c.取材位置, '切片' as 材料类别, " & _
'            " decode(b.制片方式,0,'正常',1,'重切',2,'深切',3,'连切',4,'白片',5,'重染',6,'薄片','其他') as 材料明细, " & _
'            " (b.制片数 - nvl(x.遗失数量,0)  - nvl(y.已借数量, 0)) as 在档数量, nvl(x.遗失数量,0) as 遗失数量, nvl(y.已借数量,0) as 已借数量, a.存放状态,a.借阅状态, " & _
'            " e.档案名称, '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置, e.详细地址 " & _
'            " from 病理归档信息 a, 病理制片信息 b, 病理取材信息 c, 病理检查信息 d, 病理档案信息 e, " & strLinkTable & _
'            " where a.制片id=b.id and b.材块id=c.材块id and c.病理医嘱id=d.病理医嘱id and a.id = x.归档ID(+) and a.id=y.归档id(+) and a.档案id=e.id  and d.病理号=[1] " & _
'        " Union All " & _
'            " select d.检查类型, d.病理号, a.id, c.序号, c.标本名称, c.取材位置, decode(b.特检类型,0, '免疫',1,'特染',2,'分子') as 材料类别, " & _
'            " decode(b.特检细目,0,decode(b.特检类型,0, '免疫',1,'特染',2,'分子'),1,'鉴别',2,'多耐药',3,'荧光',4,'普通') || '(' || f.抗体名称 || decode(b.制作类型,-1,'-补',0,'','-重' || b.制作类型) || ')' as 材料明细, " & _
'            " (1 - nvl(x.遗失数量,0)  - nvl(y.已借数量, 0)) as 在档数量, nvl(x.遗失数量,0) as 遗失数量, nvl(y.已借数量,0) as 已借数量, a.存放状态,a.借阅状态, " & _
'            " e.档案名称, '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置, e.详细地址 " & _
'            " from 病理归档信息 a, 病理特检信息 b, 病理取材信息 c, 病理检查信息 d, 病理档案信息 e, 病理抗体信息 f, " & strLinkTable & _
'            " where a.特检id=b.id and b.材块id=c.材块id and c.病理医嘱id=d.病理医嘱id and a.id = x.归档ID(+) and a.id=y.归档id(+) " & _
'            " and a.档案id=e.id and b.抗体ID=f.抗体ID and d.病理号=[1] " & _
'        ") order by 材料类别, 序号,材料明细,存放状态"
'
'
'    Set ufgLose.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatholNum)
'
'    Call ufgLose.RefreshData
'
'
'    If ufgLose.AdoData.RecordCount <= 0 Then
'        Call MsgBoxD(Me, "未查询到相关数据。", vbOKOnly, Me.Caption)
'        Exit Sub
'    End If
'End Sub


Private Sub QueryPatholMaterialData(ByVal strPatholNum As String, ByVal dtStartDate As Date, ByVal dtEndDate As Date)
'查询病理材料
    Dim strSQL As String
    Dim strFilter As String
    
    Dim strSqlMaterial As String
    Dim strSqlSlices As String
    Dim strSqlSpecial As String
    Dim strSqlMaterialCount As String
    Dim strSqlLoseCount As String
    
    
    Call ufgLose.ClearListData
    
    If strPatholNum <> "" Then
        strFilter = " and b.病理号=[1] "
        
        strSqlMaterialCount = " select 归档ID, sum(nvl(借阅数量,0)) - sum(nvl(归还数量,0)) as 已借数量 " & _
                                " from 病理检查信息 h, 病理归档信息 i, 病理借阅关联 j " & _
                                " where h.病理医嘱id=i.病理医嘱id and i.id = j.归档id and h.病理号=[1] and j.归还状态=0 " & _
                                " and not exists(select 1 from 病理遗失信息 where 借阅ID=j.借阅id and 归档ID=j.归档id) " & _
                                " group by 归档ID"
                                
        strSqlLoseCount = " select 归档ID, sum(nvl(遗失数量,0)) as 总遗失数量 " & _
                            " from 病理检查信息 h, 病理归档信息 i, 病理遗失信息 j " & _
                            " where h.病理医嘱id=i.病理医嘱id and i.id = j.归档id and  h.病理号=[1] " & _
                            " group by 归档ID"
    Else
        strFilter = " and e.归档ID in(select 归档ID from 病理遗失信息　where 遗失日期 between [2] and [3]) "
        
        strSqlMaterialCount = " select j.归档ID, sum(nvl(j.借阅数量,0)) - sum(nvl(j.归还数量,0)) as 已借数量 " & _
                                " from 病理检查信息 h, 病理归档信息 i, 病理借阅关联 j " & _
                                " Where h.病理医嘱id = i.病理医嘱id And i.ID = j.归档id and j.归还状态=0 " & _
                                " and j.归档id in(select 归档id from 病理遗失信息 where 借阅id is null " & _
                                " and  遗失日期 between [2] and [3]) " & _
                                " group by j.归档ID "
                                
        strSqlLoseCount = " select j.归档ID, sum(nvl(j.遗失数量,0)) as 总遗失数量 " & _
                            " from 病理检查信息 h, 病理归档信息 i, 病理遗失信息 j " & _
                            " Where h.病理医嘱id = i.病理医嘱id And i.ID = j.归档id " & _
                            " and j.归档ID in (select 归档ID  from 病理遗失信息 where 遗失日期 between [2] and [3]) " & _
                            " group by j.归档ID  "
    End If
        
    '统计遗失的材料数量（统计借阅数量时，只能统计未归还的借阅数量，部分规划或已遗失的材料将进行遗失处理，并体现到遗失数量中）
    
    
    strSqlMaterial = " select c.id, b.检查类型, '蜡块' as 材料类别, c.存放状态, b.病理号,a.序号,a.标本名称,a.取材位置,a.蜡块数 as 数量, e.遗失数量, " & _
                    " decode( e.借阅id, null,'内部遗失', '借阅遗失') as 遗失原因,d.档案名称, '房间:' || d.所属房间 || ' 柜号:' || d.所属柜号 || ' 抽屉:' || d.所属抽屉 as 存放位置, " & _
                    " case when a.申请ID is null then '常规取材' else '补取材' end as 材料明细 " & _
                    " from 病理取材信息 a, 病理检查信息 b, 病理归档信息 c, 病理档案信息 d, 病理遗失信息 e " & _
                    " where a.病理医嘱id=b.病理医嘱id and a.材块id=c.材块id and c.档案id=d.id and c.id=e.归档id " & strFilter & _
                    IIf(strPatholNum = "", "", _
                    " Union All select c.id, b.检查类型, '蜡块' as 材料类别,  c.存放状态, b.病理号,a.序号,a.标本名称,a.取材位置, a.蜡块数 as 数量, 0 as 遗失数量, '无遗失' as 遗失原因,d.档案名称, " & _
                    " '房间:' || d.所属房间 || ' 柜号:' || d.所属柜号 || ' 抽屉:' || d.所属抽屉 as 存放位置, case when a.申请ID is null then '常规取材' else '补取材' end as 材料明细 " & _
                    " from 病理取材信息 a, 病理检查信息 b,  病理归档信息 c, 病理档案信息 d " & _
                    " Where a.病理医嘱id = b.病理医嘱id And a.材块id = c.材块id And c.档案id = d.ID " & _
                    " and not exists(select 1 from 病理遗失信息 where 归档ID=c.id) " & strFilter)
    
    strSqlSlices = " select c.id,  b.检查类型, '切片' as 材料类别,c.存放状态, b.病理号,a.序号,a.标本名称,a.取材位置, x.制片数 as 数量, e.遗失数量, " & _
                    " decode( e.借阅id, null,'内部遗失', '借阅遗失') as 遗失原因,d.档案名称, '房间:' || d.所属房间 || ' 柜号:' || d.所属柜号 || ' 抽屉:' || d.所属抽屉 as 存放位置, " & _
                    " decode(x.制片方式,0,'正常',1,'重切',2,'深切',3,'连切',4,'白片',5,'重染',6,'薄片','其他') as 材料明细 " & _
                    " from 病理制片信息 x, 病理取材信息 a, 病理检查信息 b, 病理归档信息 c, 病理档案信息 d, 病理遗失信息 e " & _
                    " where x.材块id=a.材块id and a.病理医嘱id=b.病理医嘱id and x.id=c.制片id and c.档案id=d.id and c.id=e.归档id " & strFilter & _
                    IIf(strPatholNum = "", "", _
                    " Union All select c.id, b.检查类型, '切片' as 材料类别, c.存放状态, b.病理号,a.序号,a.标本名称,a.取材位置,x.制片数 as 数量, 0 as 遗失数量, '无遗失' as 遗失原因,d.档案名称, " & _
                    " '房间:' || d.所属房间 || ' 柜号:' || d.所属柜号 || ' 抽屉:' || d.所属抽屉 as 存放位置, decode(x.制片方式,0,'正常',1,'重切',2,'深切',3,'连切',4,'白片',5,'重染',6,'薄片','其他') as 材料明细 " & _
                    " from 病理制片信息 x, 病理取材信息 a, 病理检查信息 b, 病理归档信息 c, 病理档案信息 d " & _
                    " Where X.材块id = a.材块id And a.病理医嘱id = b.病理医嘱id And X.ID = c.制片id And c.档案id = d.ID " & _
                    " and not exists(select 1 from 病理遗失信息 where 归档ID=c.id) " & strFilter)


    strSqlSpecial = " select c.id,  b.检查类型, decode(x.特检类型,0, '免疫',1,'特染',2,'分子') as 材料类别,c.存放状态, b.病理号,a.序号,a.标本名称,a.取材位置, 1 as 数量, e.遗失数量, " & _
                    " decode( e.借阅id, null,'内部遗失', '借阅遗失') as 遗失原因,d.档案名称,'房间:' || d.所属房间 || ' 柜号:' || d.所属柜号 || ' 抽屉:' || d.所属抽屉 as 存放位置, " & _
                    " decode(x.特检细目,0,decode(x.特检类型,0, '免疫',1,'特染',2,'分子'),1,'鉴别',2,'多耐药',3,'荧光',4,'普通') || '(' || f.抗体名称 " & _
                    " || decode(x.制作类型,-1,'-补',0,'','-重' || x.制作类型) || ')' as 材料明细 " & _
                    " from 病理特检信息 x, 病理取材信息 a, 病理检查信息 b, 病理归档信息 c, 病理档案信息 d, 病理遗失信息 e, 病理抗体信息 f " & _
                    " where x.材块id=a.材块id and a.病理医嘱id=b.病理医嘱id and x.id=c.特检id and c.档案id=d.id and c.id=e.归档id and x.抗体id=f.抗体id " & strFilter & _
                    IIf(strPatholNum = "", "", _
                    " Union All select c.id, b.检查类型, decode(x.特检类型,0, '免疫',1,'特染',2,'分子') as 材料类别, c.存放状态, b.病理号,a.序号,a.标本名称,a.取材位置,1 as 数量,0 as 遗失数量, '无遗失' as 遗失原因,d.档案名称, " & _
                    " '房间:' || d.所属房间 || ' 柜号:' || d.所属柜号 || ' 抽屉:' || d.所属抽屉 as 存放位置, " & _
                    " decode(x.特检细目,0,decode(x.特检类型,0, '免疫',1,'特染',2,'分子'),1,'鉴别',2,'多耐药',3,'荧光',4,'普通') || '(' || f.抗体名称 " & _
                    " || decode(x.制作类型,-1,'-补',0,'','-重' || x.制作类型) || ')' as 材料明细 " & _
                    " from 病理特检信息 x, 病理取材信息 a, 病理检查信息 b, 病理归档信息 c, 病理档案信息 d, 病理抗体信息 f " & _
                    " Where X.材块id = a.材块id And a.病理医嘱id = b.病理医嘱id And X.ID = c.特检id And c.档案id = d.ID And X.抗体id = f.抗体id " & _
                    " and not exists(select 1 from 病理遗失信息 where 归档ID=c.id) " & strFilter)
    
    strSQL = "select id,检查类型,材料类别,存放状态,病理号,序号, 标本名称,取材位置,材料明细, decode(sum(nvl(遗失数量,0)), 0, '无遗失', 遗失原因) as 遗失原因,档案名称,存放位置, " & _
                " (nvl(数量,0) - nvl(已借数量, 0) - nvl(总遗失数量, 0)) as 在档数量, sum(nvl(遗失数量,0)) as 遗失数量 " & _
                " from( " & strSqlMaterial & " union all " & strSqlSlices & " union all " & strSqlSpecial & ") u, (" & _
                strSqlMaterialCount & ")v, (" & strSqlLoseCount & ") w" & _
                " where u.id =v.归档ID(+) and u.id = w.归档ID(+) " & _
                " group by id,检查类型,材料类别,存放状态,病理号,序号, 标本名称,取材位置, 遗失原因,档案名称,存放位置,材料明细,数量,已借数量,总遗失数量 " & _
                IIf(strPatholNum = "", " having sum(nvl(遗失数量,0)) > 0 ", "") & _
                " order by 材料类别,检查类型,病理号,序号, 遗失原因 "
    
    
    Set ufgLose.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatholNum, dtStartDate, dtEndDate)
                                                    
    Call ufgLose.RefreshData
                                                          

    If ufgLose.AdoData.RecordCount <= 0 Then
        mstrCurSelectPatholNum = ""
        txtName.Text = ""
        txtSex.Text = ""
        txtAge.Text = ""
        txtStudyItem.Text = ""
        txtStudyType.Text = ""
        txtStudyDate.Text = ""
        Call MsgBoxD(Me, "未查询到相关数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
End Sub


Private Sub InitLoseList()
'初始化借阅列表
    Dim strTemp As String
    
        '设置行数
    ufgLose.GridRows = glngStandardRowCount
    '设置行高
    ufgLose.RowHeightMin = glngStandardRowHeight
    
    strTemp = zlDatabase.GetPara("遗失列表配置", glngSys, G_LNG_PATHOLLOSE_NUM, "")
    
    ufgLose.IsCopyMode = True
    ufgLose.IsKeepRows = False
    ufgLose.DefaultColNames = gstrMaterialLoseCols
    ufgLose.ColNames = IIf(strTemp <> "", strTemp, gstrMaterialLoseCols)
    '禁止右键弹出列表配置窗口
    ufgLose.IsEjectConfig = False
    ufgLose.ColConvertFormat = gstrMaterialLoseConvertFormat
                                 
End Sub


Private Sub Execute_MaterialLose()
'材料遗失处理
    Dim frmLose As frmPatholLoseEnreg
    
    If Not ufgLose.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要进行遗失处理的材料记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgLose.Text(ufgLose.SelectionRow, gstrPatholCol_存放状态) = "已遗失" Then
        Call MsgBoxD(Me, "该材料已遗失，不能进行遗失处理。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
On Error GoTo errFree
        Set frmLose = New frmPatholLoseEnreg
        
        Call frmLose.ShowLoseWindow(ufgLose, Me)
        
        If frmLose.blnIsOk Then
            Call RefreshStateInf
        End If
        
errFree:
    Call Unload(frmLose)
    Set frmLose = Nothing
End Sub


Private Sub Execute_MaterialFind()
'材料找回处理
    Dim frmFind As frmPatholLoseEnreg
    
    If Not ufgLose.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要进行遗失处理的材料记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgLose.Text(ufgLose.SelectionRow, gstrPatholCol_存放状态) = "存档中" Then
        Call MsgBoxD(Me, "该材料处于存档中，不能进行找回处理。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgLose.Text(ufgLose.SelectionRow, gstrPatholCol_遗失原因) = "借阅遗失" Then
        Call MsgBoxD(Me, "因借阅产生的遗失，只能通过借阅归还找回遗失的材料。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
On Error GoTo errFree
        Set frmFind = New frmPatholLoseEnreg
        
        Call frmFind.ShowFindWindow(ufgLose, Me)
        
        If frmFind.blnIsOk Then
            Call RefreshStateInf
        End If
        
errFree:
    Call Unload(frmFind)
    Set frmFind = Nothing
End Sub

Private Sub Execute_Help()
'帮助
    Shell "hh.exe  zl9start.chm", vbNormalFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
err.Clear
End Sub

Private Sub mnu_About_Click()
'关于
On Error GoTo ErrHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_back_Click()
'发送反馈
On Error GoTo ErrHandle
    Call zlMailTo(Me.hWnd)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_BBS_Click()
'中联论坛
On Error GoTo ErrHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Exit_Click()
'退出
On Error GoTo ErrHandle
    Call Unload(Me)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_ExportExcel_Click()
'导处Excel
On Error GoTo ErrHandle
    Call MenuPrint(3)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Font_Click()
'字体
On Error GoTo ErrHandle
    
    diaFont.flags = 1
    diaFont.FontBold = ufgLose.DataGrid.Font.Bold
    diaFont.FontName = ufgLose.DataGrid.Font.Name
    diaFont.FontSize = ufgLose.DataGrid.Font.Size
    diaFont.FontStrikethru = ufgLose.DataGrid.Font.Strikethrough
    diaFont.FontUnderline = ufgLose.DataGrid.Font.Underline
    
    diaFont.ShowFont
    
    '档案列表
    ufgLose.DataGrid.Font.Bold = diaFont.FontBold
    ufgLose.DataGrid.Font.Name = diaFont.FontName
    ufgLose.DataGrid.Font.Size = diaFont.FontSize
    ufgLose.DataGrid.Font.Strikethrough = diaFont.FontStrikethru
    ufgLose.DataGrid.Font.Underline = diaFont.FontUnderline
    
    
    Call ufgLose.DataGrid.Refresh
    
    ufgLose.DataGrid.AutoSizeMode = flexAutoSizeColWidth
    Call ufgLose.DataGrid.AutoSize(0, ufgLose.DataGrid.Rows - 1)
    
    ufgLose.DataGrid.AutoSizeMode = flexAutoSizeRowHeight
    Call ufgLose.DataGrid.AutoSize(0, ufgLose.DataGrid.Rows - 1)
    
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_HelpMain_Click()
'帮助
On Error GoTo ErrHandle
    Call Execute_Help
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_HomePage_Click()
'中联主页
On Error GoTo ErrHandle
    Call zlHomePage(Me.hWnd)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_MaterialFind_Click()
'材料找回
On Error GoTo ErrHandle
    Call Execute_MaterialFind
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_MaterialLose_Click()
'材料遗失
On Error GoTo ErrHandle
    Call Execute_MaterialLose
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Preview_Click()
'预览数据列表
On Error GoTo ErrHandle
    Call MenuPrint(0)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_Print_Click()
'预览数据列表
On Error GoTo ErrHandle
    Call MenuPrint(1)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_PrintConfig_Click()
'打印配置
On Error GoTo ErrHandle
    Call zlPrintSet
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_StandardButton_Click()
On Error GoTo ErrHandle
    Dim intCount As Long
    Me.mnu_StandardButton.Checked = Not Me.mnu_StandardButton.Checked
    Me.tbrTools.Visible = Me.mnu_StandardButton.Checked
    
    If Me.mnu_WordLabel.Checked Then
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = Me.tbrTools.Buttons(intCount).tag
        Next
    Else
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = ""
        Next
    End If

    Me.tbrTools.Refresh
    
    Form_Resize
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_StateBar_Click()
On Error GoTo ErrHandle
    Me.mnu_StateBar.Checked = Not Me.mnu_StateBar.Checked
    Me.stbThis.Visible = Me.mnu_StateBar.Checked
    
    Call Form_Resize
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnu_WordLabel_Click()
On Error GoTo ErrHandle
    Dim intCount As Long
    
    Me.mnu_WordLabel.Checked = Not Me.mnu_WordLabel.Checked

    If Me.mnu_WordLabel.Checked Then
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = Me.tbrTools.Buttons(intCount).tag
        Next
    Else
        For intCount = 1 To Me.tbrTools.Buttons.Count
            Me.tbrTools.Buttons(intCount).Caption = ""
        Next
    End If
    
    Me.tbrTools.Refresh
    
    Call Form_Resize
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Option1_Click()
On Error Resume Next
    txtPatholNo.Enabled = True
    txtPatholNo.BackColor = &H80000005
    
    dtpStart.Enabled = False
    dtpEnd.Enabled = False
    
    err.Clear
End Sub

Private Sub Option2_Click()
On Error Resume Next
    txtPatholNo.Enabled = False
    txtPatholNo.Text = ""
    txtPatholNo.BackColor = &H8000000F
    
    dtpStart.Enabled = True
    dtpEnd.Enabled = True
    
    err.Clear
End Sub

Private Sub tbrTools_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrHandle
    
    Select Case UCase(Button.Key)
        Case UCase("tbn_PreviewList")   '预览
            Call MenuPrint(0)
            
        Case UCase("tbn_PreviewPrint")  '打印
            Call MenuPrint(1)
            
        Case UCase("tbn_NewLose")   '材料遗失
            Call Execute_MaterialLose
    
        Case UCase("tbn_FindLose")   '材料找回
            Call Execute_MaterialFind
                        
        Case UCase("tbn_Help")  '帮助
            Call Execute_Help
            
        Case UCase("tbn_Exit")  '推出档案管理模块
            Call Unload(Me)
    End Select
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Public Sub MenuPrint(intOutMode As Byte)
    '---------------------------------------------------
    '功能：    根据屏幕打印预览, 0弹出操作选择对话框，1预览，2打印，3导出Excel
    '参数：    输出方式
    '返回：
    '---------------------------------------------------
    Dim objPrint As New zlPrint1Grd

    Set objPrint.Body = ufgLose.DataGrid
    
    objPrint.Title = "病理材料状态清单"

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


Private Sub txtPatholNo_KeyPress(KeyAscii As Integer)
 '回车执行查询
 On Error GoTo ErrHandle
    If KeyAscii = 13 Then
        mblnMoved = MovedByDate(dtpStart.value)
        mqwQueryWay = qwPatholNum
        
        Call QueryStudyInf(txtPatholNo.Text)
        Call ufgLose.ClearListData
        
        If txtPatholNo.Text = "" Then Exit Sub
        
        Call QueryPatholMaterialData(txtPatholNo.Text, Format(dtpStart.value, "yyyy-mm-dd 23:59:59"), Format(dtpEnd.value, "yyyy-mm-dd 23:59:59"))
    End If

Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgLose_OnColFormartChange()
On Error GoTo ErrHandle
    zlDatabase.SetPara "遗失列表配置", ufgLose.GetColsString(ufgLose), glngSys, G_LNG_PATHOLLOSE_NUM
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgLose_OnSelChange()
On Error GoTo ErrHandle
    '查询病人信息
    
    If mqwQueryWay = qwPatholNum Then Exit Sub
    If Not ufgLose.IsSelectionRow Then Exit Sub
    If ufgLose.Text(ufgLose.SelectionRow, gstrPatholCol_病理号) = "" Then Exit Sub
    If mstrCurSelectPatholNum = ufgLose.Text(ufgLose.SelectionRow, gstrPatholCol_病理号) Then Exit Sub
    
    mstrCurSelectPatholNum = ufgLose.Text(ufgLose.SelectionRow, gstrPatholCol_病理号)
    
    Call QueryStudyInf(mstrCurSelectPatholNum)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub
