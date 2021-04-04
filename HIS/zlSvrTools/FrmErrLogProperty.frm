VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmErrLogProperty 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "错误日志属性"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   4155
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   7329
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "属性(&P)"
      TabPicture(0)   =   "FrmErrLogProperty.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Txt错误序号"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lbl错误序号"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Txt用户名"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lbl用户名"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Txt错误类型"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lbl错误类型"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Lbl错误信息"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Txt工作站"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Lbl工作站"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Txt进入时间"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Lbl进入时间"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Txt会话号"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Lbl会话号"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Txt错误信息"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      Begin VB.TextBox Txt错误信息 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   1020
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   2310
         Width           =   2970
      End
      Begin VB.Label Lbl会话号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "会话号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   14
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Txt会话号 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   13
         Top             =   450
         Width           =   3000
      End
      Begin VB.Label Lbl进入时间 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "进入时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   12
         Top             =   1710
         Width           =   720
      End
      Begin VB.Label Txt进入时间 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   11
         Top             =   1710
         Width           =   3000
      End
      Begin VB.Label Lbl工作站 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "工作站"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   10
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Txt工作站 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   9
         Top             =   1080
         Width           =   3000
      End
      Begin VB.Label Lbl错误信息 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "错误信息"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   8
         Top             =   2310
         Width           =   720
      End
      Begin VB.Label Lbl错误类型 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "错误类型"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   7
         Top             =   780
         Width           =   720
      End
      Begin VB.Label Txt错误类型 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   6
         Top             =   780
         Width           =   3000
      End
      Begin VB.Label Lbl用户名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "用户名"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   5
         Top             =   1410
         Width           =   540
      End
      Begin VB.Label Txt用户名 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   4
         Top             =   1410
         Width           =   3000
      End
      Begin VB.Label Lbl错误序号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "错误序号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   3
         Top             =   2010
         Width           =   720
      End
      Begin VB.Label Txt错误序号 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   2
         Top             =   2010
         Width           =   3000
      End
   End
   Begin VB.CommandButton Cmd退出 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2760
      TabIndex        =   0
      Top             =   4290
      Width           =   1100
   End
End
Attribute VB_Name = "FrmErrLogProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd退出_Click()
    Unload Me
End Sub
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               