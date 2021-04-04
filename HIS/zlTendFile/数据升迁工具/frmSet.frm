VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4860
   Icon            =   "frmSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3540
      TabIndex        =   8
      Top             =   1890
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3540
      TabIndex        =   7
      Top             =   1380
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "运行时间设定"
      Height          =   1605
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3075
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   300
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   92078083
         CurrentDate     =   40540
      End
      Begin MSComCtl2.DTPicker dtp开始时间1 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   690
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   92078083
         CurrentDate     =   40540.0833333333
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   92078083
         CurrentDate     =   40540.1666666667
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结束处理时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   675
         TabIndex        =   5
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "打印解析开始时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   315
         TabIndex        =   3
         Top             =   750
         Width           =   1440
      End
      Begin VB.Label lbl开始时间 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "数据升迁开始时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   315
         TabIndex        =   1
         Top             =   360
         Width           =   1440
      End
   End
   Begin VB.Label Label3 
      Caption         =   "    数据升迁工具根据指定的开始时间进行数据升迁工作至打印解析开始时间时停止，然后进行打印解析工作至结束处理时间为止。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   4335
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\护理数据升迁", "开始时间", Format(Me.dtp开始时间.Value, "HH:mm")
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\护理数据升迁", "开始时间1", Format(Me.dtp开始时间1.Value, "HH:mm")
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\护理数据升迁", "结束时间", Format(Me.dtp结束时间.Value, "HH:mm")
    
    Unload Me
End Sub

Private Sub Form_Load()
    dtp开始时间.Value = GetSetting("ZLSOFT", "私有模块\ZLHIS\护理数据升迁", "开始时间", "00:00")
    dtp开始时间1.Value = GetSetting("ZLSOFT", "私有模块\ZLHIS\护理数据升迁", "开始时间1", "02:00")
    dtp结束时间.Value = GetSetting("ZLSOFT", "私有模块\ZLHIS\护理数据升迁", "结束时间", "04:00")
End Sub
