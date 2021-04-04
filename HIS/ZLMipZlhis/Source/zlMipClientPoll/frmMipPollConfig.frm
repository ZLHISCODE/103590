VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMipPollConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "配置"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6510
   Icon            =   "frmMipPollConfig.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6510
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame3 
      Caption         =   "连接消息服务平台"
      Height          =   1590
      Left            =   90
      TabIndex        =   12
      Top             =   2085
      Width           =   6330
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   10
         Left            =   4440
         TabIndex        =   20
         Top             =   1110
         Width           =   1740
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   9
         Left            =   1890
         TabIndex        =   19
         Top             =   1110
         Width           =   1695
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   4440
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   750
         Width           =   1740
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   7
         Left            =   1890
         TabIndex        =   17
         Top             =   750
         Width           =   1695
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "设置连接消息服务平台所需要的相关参数"
         Height          =   180
         Index           =   11
         Left            =   1065
         TabIndex        =   21
         Top             =   375
         Width           =   3240
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   1
         Left            =   225
         Picture         =   "frmMipPollConfig.frx":6852
         Top             =   315
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "连接端口"
         Height          =   180
         Index           =   10
         Left            =   3660
         TabIndex        =   16
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "连接地址"
         Height          =   180
         Index           =   9
         Left            =   1080
         TabIndex        =   15
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "连接密码"
         Height          =   180
         Index           =   8
         Left            =   3660
         TabIndex        =   14
         Top             =   795
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "连接用户"
         Height          =   180
         Index           =   7
         Left            =   1080
         TabIndex        =   13
         Top             =   795
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "时间间隔"
      Height          =   1920
      Left            =   90
      TabIndex        =   2
      Top             =   60
      Width           =   6330
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   6
         Left            =   2745
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1485
         Width           =   750
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   5
         Left            =   2910
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1110
         Width           =   660
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   4
         Left            =   2910
         MaxLength       =   2
         TabIndex        =   7
         Top             =   735
         Width           =   630
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Index           =   5
         Left            =   3570
         TabIndex        =   5
         Top             =   1110
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txt(5)"
         BuddyDispid     =   196610
         BuddyIndex      =   5
         OrigLeft        =   2625
         OrigTop         =   795
         OrigRight       =   2880
         OrigBottom      =   1065
         Max             =   60
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Index           =   4
         Left            =   3540
         TabIndex        =   6
         Top             =   735
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txt(4)"
         BuddyDispid     =   196610
         BuddyIndex      =   4
         OrigLeft        =   2280
         OrigTop         =   345
         OrigRight       =   2535
         OrigBottom      =   615
         Max             =   60
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Index           =   6
         Left            =   3495
         TabIndex        =   11
         Top             =   1485
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt(6)"
         BuddyDispid     =   196610
         BuddyIndex      =   6
         OrigLeft        =   2625
         OrigTop         =   795
         OrigRight       =   2880
         OrigBottom      =   1065
         Max             =   30
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "设置连接产品数据库所需要的相关参数"
         Height          =   180
         Index           =   13
         Left            =   1035
         TabIndex        =   22
         Top             =   390
         Width           =   3060
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   2
         Left            =   165
         Picture         =   "frmMipPollConfig.frx":81D4
         Top             =   300
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "轮询服务启动后每隔            分钟检查一次"
         Height          =   180
         Index           =   6
         Left            =   1080
         TabIndex        =   9
         Top             =   1545
         Width           =   3780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "等待消息收到回执超过           秒判断为发送消息失败"
         Height          =   180
         Index           =   5
         Left            =   1080
         TabIndex        =   4
         Top             =   1170
         Width           =   4590
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "连接消息服务平台超过           秒判断为连接失败"
         Height          =   180
         Index           =   4
         Left            =   1080
         TabIndex        =   3
         Top             =   810
         Width           =   4230
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   5295
      TabIndex        =   1
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   345
      Left            =   4125
      TabIndex        =   0
      Top             =   3840
      Width           =   1100
   End
End
Attribute VB_Name = "frmMipPollConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'变量定义
Private mblnDataChanged As Boolean
Private mstrTitle As String
Private mclsMipServiceData As clsMipServiceData

'######################################################################################################################
'接口方法

Public Function ShowConfigDialog(ByVal frmParent As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsCondition As ADODB.Recordset
    Dim rs As zlDataSQLite.SQLiteRecordset
    Dim strPara As String
    Dim varPara As Variant
    
    
    Set mclsMipServiceData = New clsMipServiceData
    
    txt(4).Text = "5"
    txt(5).Text = "5"
    txt(6).Text = "5"
    
    strPara = ""
    If mclsMipServiceData.OpenFile(App.Path & "\Data\zlMspPollService.db") = True Then
        '取参数
        Set rsCondition = zlCommFun.CreateCondition
'        Call zlCommFun.SetCondition(rsCondition, "参数编号", "1")
'        rs = mclsMipServiceData.GetPara("Filter", rsCondition)
'        If rs.DataSet.BOF = False Then
'            strPara = zlCommFun.NVL(rs.DataSet("Content").Value)
'            If strPara <> "" Then
'                varPara = Split(strPara, ";")
'                txt(0).Text = varPara(0)
'                txt(1).Text = varPara(1)
'                txt(2).Text = varPara(2)
'                txt(3).Text = varPara(3)
'            End If
'        End If
        
        '取连接目标超时
        Call zlCommFun.SetCondition(rsCondition, "参数编号", "2")
        rs = mclsMipServiceData.GetPara("Filter", rsCondition)
        If rs.DataSet.BOF = False Then
            strPara = zlCommFun.NVL(rs.DataSet("Content").Value)
            txt(4).Text = Val(strPara)
        End If
        
        '取等待收到回执消息超时时间
        Call zlCommFun.SetCondition(rsCondition, "参数编号", "3")
        rs = mclsMipServiceData.GetPara("Filter", rsCondition)
        If rs.DataSet.BOF = False Then
            strPara = zlCommFun.NVL(rs.DataSet("Content").Value)
            txt(5).Text = Val(strPara)
        End If
        
        '发送服务启动间隔
        Call zlCommFun.SetCondition(rsCondition, "参数编号", "4")
        rs = mclsMipServiceData.GetPara("Filter", rsCondition)
        If rs.DataSet.BOF = False Then
            strPara = zlCommFun.NVL(rs.DataSet("Content").Value)
            txt(6).Text = Val(strPara)
        End If
        
        '消息服务平台参数
        Call zlCommFun.SetCondition(rsCondition, "参数编号", "5")
        rs = mclsMipServiceData.GetPara("Filter", rsCondition)
        If rs.DataSet.BOF = False Then
            strPara = zlCommFun.NVL(rs.DataSet("Content").Value)
            If strPara <> "" Then
                varPara = Split(strPara, ";")
                txt(7).Text = varPara(0)
                txt(8).Text = varPara(1)
                txt(9).Text = varPara(2)
                txt(10).Text = Val(varPara(3))
            End If
        End If
        
    End If
    
    mclsMipServiceData.CloseFile
        
    mblnDataChanged = False
    
    Me.Show 1, frmParent
        
    ShowConfigDialog = mblnDataChanged
    
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    Dim blnRet As Boolean
    Dim strPara As String
    
'    strPara = txt(0).Text & ";" & txt(1).Text & ";" & txt(2).Text & ";" & txt(3).Text
    
    If mclsMipServiceData.OpenFile(App.Path & "\Data\zlMspPollService.db") = True Then
        
'        blnRet = mclsMipServiceData.EditPara("1", strPara)
        blnRet = mclsMipServiceData.EditPara("2", Val(txt(4).Text))
        If blnRet Then blnRet = mclsMipServiceData.EditPara("3", Val(txt(5).Text))
        If blnRet Then blnRet = mclsMipServiceData.EditPara("4", Val(txt(6).Text))
        
        strPara = txt(7).Text & ";" & txt(8).Text & ";" & txt(9).Text & ";" & txt(10).Text
        If blnRet Then blnRet = mclsMipServiceData.EditPara("5", strPara)
        
        If blnRet = True Then
            mclsMipServiceData.CloseFile
            mblnDataChanged = True
            Unload Me
            Exit Sub
        End If
    End If
    mclsMipServiceData.CloseFile
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsMipServiceData = Nothing
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not zlCommFun.StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub
