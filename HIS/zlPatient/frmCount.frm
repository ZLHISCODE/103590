VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "编码长度设置"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "frmCount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "长度"
      Height          =   1755
      Left            =   180
      TabIndex        =   2
      Top             =   150
      Width           =   2715
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   510
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   900
         Width           =   765
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1276
         TabIndex        =   3
         Top             =   900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text1"
         BuddyDispid     =   196610
         OrigLeft        =   1530
         OrigTop         =   900
         OrigRight       =   1770
         OrigBottom      =   1215
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "   请输入你要想要得到的编码长度。"
         Height          =   675
         Left            =   420
         TabIndex        =   5
         Top             =   330
         Width           =   1845
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3300
      TabIndex        =   1
      Top             =   270
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3300
      TabIndex        =   0
      Top             =   720
      Width           =   1100
   End
End
Attribute VB_Name = "frmCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    mblnOK = False
    Me.Hide
End Sub



Private Sub cmdOK_Click()
    mblnOK = True
    Me.Hide
End Sub

Public Function GetLength(ByVal intValue As Integer, ByVal intMax As Integer) As Integer
    UpDown1.Min = intValue
    UpDown1.Max = intMax
    UpDown1.Value = intValue
    Me.Show vbModal
    GetLength = IIf(mblnOK, UpDown1.Value, 0)
    Unload Me
End Function
