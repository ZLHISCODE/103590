VERSION 5.00
Begin VB.Form frmParam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IC卡参数设置"
   ClientHeight    =   1845
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4845
   Icon            =   "frmParam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3510
      TabIndex        =   5
      Top             =   645
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3510
      TabIndex        =   4
      Top             =   180
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   150
      TabIndex        =   6
      Top             =   105
      Width           =   3225
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   1920
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   285
         Width           =   1920
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "波特率(&B)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   780
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "端口(&P)"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   0
         Top             =   360
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
'变量、API、函数申明

Private mblnStartUp As Boolean
Private mblnOK As Boolean

'######################################################################################################################
'自定义过程、函数

Public Function ShowMe(Optional ByVal frmParent As Form) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：显示参数窗体
    '参数：frmParent        参数窗体的上级窗体对象
    '返回：更改返回True；否则返回False
    '------------------------------------------------------------------------------------------------------------------
    
    mblnStartUp = True
    mblnOK = False
    
    If frmParent Is Nothing Then
        frmParam.Show 1
    Else
        frmParam.Show 1, frmParent
    End If
    ShowMe = mblnOK
    
End Function

'######################################################################################################################
'对象事件处理

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    SaveSetting "ZLSOFT", "IC卡参数", "端口", cbo(0).Text
    SaveSetting "ZLSOFT", "IC卡参数", "波特率", cbo(1).Text
    
    mblnOK = True
    Unload Me
    
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    On Error Resume Next
    
    cbo(0).Text = GetSetting(AppName:="ZLSOFT", Section:="IC卡参数", key:="端口", Default:="COM1")
    cbo(1).Text = GetSetting(AppName:="ZLSOFT", Section:="IC卡参数", key:="波特率", Default:="9600")
    
End Sub

Private Sub Form_Load()
    With cbo(0)
        .AddItem "COM1"
        .AddItem "COM2"
        .AddItem "COM3"
        .AddItem "COM4"
        .ListIndex = 0
    End With
    
    With cbo(1)
        .AddItem "110"
        .AddItem "300"
        .AddItem "600"
        .AddItem "1200"
        .AddItem "2400"
        .AddItem "4800"
        .AddItem "9600"
        .AddItem "14400"
        .AddItem "19200"
        .AddItem "28800"
        .AddItem "38400"
        .AddItem "56000"
        .AddItem "57600"
        .AddItem "115200"
        .AddItem "128000"
        .AddItem "256000"
        .ListIndex = 6
    End With
End Sub
