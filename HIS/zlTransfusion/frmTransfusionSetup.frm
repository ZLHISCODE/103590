VERSION 5.00
Begin VB.Form frmTransfusionSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmTransfusionSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5351.164
   ScaleMode       =   0  'User
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkTimeCall 
      Caption         =   "启用移动呼叫功能"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.CheckBox chk接单穿刺 
      Caption         =   "接单后直接进入穿刺状态"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3045
   End
   Begin VB.CheckBox chkAutoReady 
      Caption         =   "通过查找功能找到病人后自动接单"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Frame frmCardSet 
      Caption         =   "设备配置"
      Height          =   675
      Left            =   270
      TabIndex        =   8
      Top             =   2040
      Width           =   4470
      Begin VB.CommandButton cmdCardSet 
         Caption         =   "配置(&P)"
         Height          =   350
         Left            =   2985
         TabIndex        =   9
         Top             =   210
         Width           =   1100
      End
   End
   Begin VB.Frame fra 
      Caption         =   "请选择本工作站显示的单据类型"
      Height          =   660
      Left            =   270
      TabIndex        =   3
      Top             =   1200
      Width           =   4485
      Begin VB.CheckBox chkType 
         Caption         =   "治疗"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkType 
         Caption         =   "输液"
         Height          =   195
         Index           =   1
         Left            =   1275
         TabIndex        =   5
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkType 
         Caption         =   "注射"
         Height          =   195
         Index           =   2
         Left            =   2355
         TabIndex        =   6
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkType 
         Caption         =   "皮试"
         Height          =   195
         Index           =   3
         Left            =   3435
         TabIndex        =   7
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   210
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2955
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2490
      TabIndex        =   11
      Top             =   2955
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   12
      Top             =   2955
      Width           =   1100
   End
End
Attribute VB_Name = "frmTransfusionSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'2015-09-15：屏蔽“门诊输液自动接单”参数

Public mstrPrivs As String
Public mlng科室ID As Long 'IN:当前执行科室ID
Public mblnOk As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCardSet_Click()
    Call zlCommFun.DeviceSetup(Me, glngSys, glngModul)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOk_Click()
    Dim strPar As String, i As Long
    Dim strType As String
    Dim blnModify As Boolean
    
    '执行间范围
    blnModify = False
    If InStr(mstrPrivs, "参数设置") > 0 Then blnModify = True
    
    '接单后直接进入穿刺状态
    Call zlDatabase.SetPara("接单直接穿刺", chk接单穿刺.Value, glngSys, 1264)
    
    '移动呼叫
    Call zlDatabase.SetPara("移动呼叫", chkTimeCall.Value, glngSys, 1264)
    
    '2008-11-12
    strType = ""
    For i = 0 To chkType.Count - 1
        strType = strType & "," & chkType(i).Value
    Next
    Call zlDatabase.SetPara("显示单据种类", Mid(strType, 2), glngSys, 1264, blnModify)
    
    '2012-05-14 10.30 sp 添加
    Call zlDatabase.SetPara("门诊输液自动接单", chkAutoReady.Value, glngSys, 1264, blnModify)
    
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then Call cmdHelp_Click
End Sub

Private Sub Form_Load()
    Dim strType As String, i As Integer
    Dim intType As Integer '本机参数类型
    Dim blnModify As Boolean
    
    mblnOk = False
    blnModify = InStr(";" & mstrPrivs & ";", ";参数设置;") > 0
    
    cmdCardSet.Enabled = blnModify
    
    '接单后直接进入穿刺状态
    chk接单穿刺.Value = Val(zlDatabase.GetPara("接单直接穿刺", glngSys, 1264, ""))
    
    '移动定时呼叫
    chkTimeCall.Value = Val(zlDatabase.GetPara("移动呼叫", glngSys, 1264))
        
    '2008-11-12
    'strType = zlDatabase.GetPara("显示单据种类", glngSys, 1264, "1,1,1,1", Array(Me.chkType(0), Me.chkType(1), Me.chkType(2), Me.chkType(3)), blnModify, intType)
    strType = zlDatabase.GetPara("显示单据种类", glngSys, 1264, "1,1,1,1")
    For i = 0 To chkType.Count - 1
        chkType(i).Value = Val(Split(strType, ",")(i))
    Next
    '2012-05-14
    chkAutoReady.Value = Val(zlDatabase.GetPara("门诊输液自动接单", glngSys, 1264, "", Array(chkAutoReady), blnModify))
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng科室ID = 0
    mstrPrivs = ""
End Sub

