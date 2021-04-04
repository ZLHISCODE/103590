VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMidMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "数据库优化工具"
   ClientHeight    =   10605
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   16080
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgNormal 
      Left            =   12960
      Top             =   2400
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
            Picture         =   "mdiMain.frx":08CA
            Key             =   "会话解锁"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":11A4
            Key             =   "空间管理"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1A7E
            Key             =   "外键索引"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2358
            Key             =   "数据库性能"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2C32
            Key             =   "SQL性能"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":350C
            Key             =   "会话解锁_hot"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3DE6
            Key             =   "空间管理_hot"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":46C0
            Key             =   "外键索引_hot"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4F9A
            Key             =   "数据库性能_hot"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5874
            Key             =   "SQL性能_hot"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgHot 
      Left            =   12360
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":614E
            Key             =   "会话解锁"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6A28
            Key             =   "空间管理"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":7302
            Key             =   "外键索引"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":7BDC
            Key             =   "数据库性能"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":84B6
            Key             =   "SQL性能"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tblMenu 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   1508
      ButtonWidth     =   1640
      ButtonHeight    =   1455
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgNormal"
      HotImageList    =   "imgHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "数据库性能"
            Key             =   "_0601"
            Object.Tag             =   "功能说明：数据库运行状况的查看，快速获取AWE、ASH、ADDM报告。"
            ImageKey        =   "数据库性能"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SQL性能"
            Key             =   "_0602"
            Object.Tag             =   "功能说明：低性能SQL的快速筛查，执行计划调整和自动优化，对象统计信息查看，SQL相关的执行信息和关联报表查询，优化器相关参数查看。"
            ImageKey        =   "SQL性能"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "外键索引"
            Key             =   "_0605"
            Object.Tag             =   "功能说明：外键字段对应的索引缺失情况检查，索引补建或外键删除。"
            ImageKey        =   "外键索引"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "会话解锁"
            Key             =   "_0604"
            Object.Tag             =   "功能说明：并发阻塞情况的查询，会话查杀。"
            ImageKey        =   "会话解锁"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "空间管理"
            Key             =   "_0606"
            Object.Tag             =   "功能说明：表和索引在数据文件中的分布情况查看，表和索引的数据重整和收缩，数据文件和临时文件、UNDO文件的收缩。"
            ImageKey        =   "空间管理"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox pctTip 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   6360
         ScaleHeight     =   855
         ScaleWidth      =   12000
         TabIndex        =   1
         Top             =   0
         Width           =   12000
         Begin VB.Label lblTip 
            AutoSize        =   -1  'True
            Caption         =   "功能说明：数据库运行状况的查看，快速获取AWR、ASH、ADDM报告。"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   0
            TabIndex        =   2
            Top             =   600
            Width           =   9600
         End
      End
   End
End
Attribute VB_Name = "frmMidMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MDIForm_Load()
    tblMenu.Buttons(1).Image = "数据库性能_hot"
End Sub

Private Sub MDIForm_resize()
     frmParent.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub tblMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim btnTmp As MSComctlLib.Button
    
    frmParent.ShowForm Mid(Button.Key, 2)
    lblTip.Caption = Button.Tag
    
    For Each btnTmp In tblMenu.Buttons
        btnTmp.Image = btnTmp.Caption
    Next
    
    Button.Image = Button.Caption & "_hot"
    
End Sub

Public Sub SetToolBarEnable(ByVal blnEnable As Boolean)
    tblMenu.Enabled = blnEnable
End Sub


