VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNoticeSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "数据变动通知参数设置"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNoticeSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraJob 
      Caption         =   "客户端存活状态检查"
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Top             =   4800
      Width           =   5175
      Begin VB.TextBox txtJob 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   5
         Top             =   705
         Width           =   495
      End
      Begin VB.CheckBox chkJob 
         Caption         =   "开启自动作业，维护异常关闭的客户端信息"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblJob 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "分钟执行一次"
         Height          =   195
         Index           =   1
         Left            =   3600
         TabIndex        =   24
         Top             =   750
         Width           =   1080
      End
      Begin VB.Label lblJob 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "自动作业执行频率：每"
         ForeColor       =   &H80000011&
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   23
         Top             =   750
         Width           =   1800
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&O)"
      Height          =   345
      Left            =   3000
      TabIndex        =   13
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   4200
      TabIndex        =   12
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame fraClient 
      Caption         =   "客户端设置"
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   5175
      Begin VB.TextBox txtCheck 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   33
         Top             =   1035
         Width           =   495
      End
      Begin VB.TextBox txtPortEnd 
         Height          =   285
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   3
         Tag             =   "21"
         Text            =   "0"
         Top             =   360
         Width           =   750
      End
      Begin VB.TextBox txtPortStart 
         Height          =   285
         Left            =   840
         MaxLength       =   5
         TabIndex        =   2
         Tag             =   "21"
         Text            =   "0"
         Top             =   360
         Width           =   1005
      End
      Begin MSComCtl2.UpDown udPortStart 
         Height          =   285
         Left            =   1845
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   20001
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtPortStart"
         BuddyDispid     =   196618
         OrigLeft        =   2040
         OrigTop         =   360
         OrigRight       =   2295
         OrigBottom      =   630
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udPortEnd 
         Height          =   285
         Left            =   3390
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   20010
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtPortEnd"
         BuddyDispid     =   196617
         OrigLeft        =   3646
         OrigTop         =   360
         OrigRight       =   3901
         OrigBottom      =   645
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCheckTip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(对于实时性要求较高的环境，可适当加快检查频率)"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   840
         TabIndex        =   34
         Top             =   1320
         Width           =   4080
      End
      Begin VB.Label lblCheck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "秒执行一次"
         Height          =   195
         Index           =   1
         Left            =   3840
         TabIndex        =   32
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblCheck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "服务器存活状态检查频率：每"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   31
         Top             =   1080
         Width           =   2340
      End
      Begin VB.Label lblClientTip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(用于设置导航台接收消息时使用的端口范围)"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   840
         TabIndex        =   20
         Top             =   720
         Width           =   3540
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-->"
         Height          =   195
         Left            =   2280
         TabIndex        =   19
         Top             =   405
         Width           =   240
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "端口号"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   405
         Width           =   540
      End
   End
   Begin VB.Frame fraSer 
      Caption         =   "服务器设置"
      Height          =   2775
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtInterval 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3045
         MaxLength       =   3
         TabIndex        =   29
         Top             =   2115
         Width           =   495
      End
      Begin VB.CheckBox chkLog 
         Caption         =   "记录数据变动通知发送日志"
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   1800
         Width           =   2775
      End
      Begin VB.CheckBox chkLogin 
         Caption         =   "网络断开、恢复后自动重新连接服务器"
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox txtIp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   8
         Tag             =   "IP地址"
         Top             =   525
         Width           =   315
      End
      Begin VB.TextBox txtIp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   870
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "IP地址"
         Top             =   525
         Width           =   435
      End
      Begin VB.TextBox txtIp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "IP地址"
         Top             =   525
         Width           =   435
      End
      Begin VB.TextBox txtIp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   7
         Tag             =   "IP地址"
         Top             =   525
         Width           =   315
      End
      Begin VB.TextBox txtSerPort 
         Height          =   285
         Left            =   840
         MaxLength       =   5
         TabIndex        =   1
         Tag             =   "21"
         Top             =   915
         Width           =   1005
      End
      Begin MSComCtl2.UpDown udSerPort 
         Height          =   285
         Left            =   1845
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   915
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   9999
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtSerPort"
         BuddyDispid     =   196629
         OrigLeft        =   2040
         OrigTop         =   795
         OrigRight       =   2295
         OrigBottom      =   1065
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtIpSet 
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "         ．         ．         ．         "
         Top             =   480
         Width           =   2235
      End
      Begin VB.Label lblInterval1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(对于业务量较大的环境，可适当加快发送频率)"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   840
         TabIndex        =   30
         Top             =   2400
         Width           =   3720
      End
      Begin VB.Label lblInterval 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数据变动通知发送频率：每             毫秒执行一次"
         Height          =   195
         Left            =   840
         TabIndex        =   28
         Top             =   2160
         Width           =   3825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP"
         Height          =   195
         Left            =   630
         TabIndex        =   21
         Top             =   525
         Width           =   150
      End
      Begin VB.Label lblSerTip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(用于向客户端发送消息)"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2280
         TabIndex        =   16
         Top             =   960
         Width           =   1920
      End
      Begin VB.Label lblSer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "端口号"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmNoticeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnLogin As Boolean
Private mlngPortS As Long
Private mlngPortE As Long
Private mblnJob As Boolean
Private mintInterval As Integer
Private mlngCheckInterval As Long


Public Sub ShowEdit()
    Dim i As Integer, rsTmp As ADODB.Recordset
    
    '1.加载服务器端IP
    For i = 0 To 3
        txtIp(i).Text = Split(gstrIp, ".")(i)
    Next
    txtSerPort.Text = glngPort
    
    '断网重连设置
    chkLogin.Value = IIf(Val(zlDatabase.GetPara("网络断网自动重连")) = 1, 1, 0)
    mblnLogin = chkLogin.Value = 1
    
    chkLog.Value = gintLog
    txtInterval.Text = gintInterval
    
    '2.加载客户端设置
    Set rsTmp = GetClientPort
    If rsTmp Is Nothing Then
        txtPortStart.Text = 0
        txtPortEnd.Text = 0
    Else
        If rsTmp!参数值 & "" = 0 Then   '参数值为0, 说明已经禁用了消息接收
            txtPortStart.Text = 0
            txtPortEnd.Text = 0
        Else
            txtPortStart.Text = Split(rsTmp!参数值 & "", "-")(0)
            txtPortEnd.Text = Split(rsTmp!参数值 & "", "-")(1)
        End If
    End If
    mlngPortS = Val(txtPortStart.Text)
    mlngPortE = Val(txtPortEnd.Text)
    
    mlngCheckInterval = GetCheckInterval
    txtCheck.Text = mlngCheckInterval
    
    '3.加载自动作业设置
    Set rsTmp = GetJobs
    If rsTmp Is Nothing Or rsTmp.RecordCount = 0 Then
        chkJob.Value = 0: mblnJob = False
    Else
        If rsTmp!作业号 = 0 Then
            chkJob.Value = 0: mblnJob = False
        Else
            chkJob.Value = 1: mblnJob = True
            txtJob.Text = Val(rsTmp!间隔时间 & ""): mintInterval = Val(txtJob.Text)
        End If
    End If
    
    Call UpdateClientSet
    Call UpdateJobSet
    Me.Show 1
End Sub


Private Sub UpdateClientSet()
    '同步导航台端口设置
    With txtPortStart
        udPortStart.Enabled = .Enabled
        txtPortEnd.Enabled = .Enabled
        udPortEnd.Enabled = .Enabled
        
        If .Enabled Then
            lblClient.ForeColor = vbBlack
            lblRange.ForeColor = vbBlack
        Else
            lblClient.ForeColor = vbGrayText
            lblRange.ForeColor = vbGrayText
        End If
    End With
End Sub

Private Sub UpdateJobSet()
    '同步自动作业设置
    txtJob.Enabled = chkJob.Value = 1
    If txtJob.Enabled Then
        lblJob(0).ForeColor = vbBlack
        lblJob(1).ForeColor = vbBlack
    Else
        lblJob(0).ForeColor = vbGrayText
        lblJob(1).ForeColor = vbGrayText
    End If
End Sub

Private Sub chkClient_Click()
    UpdateClientSet
End Sub

Private Sub chkJob_Click()
    UpdateJobSet
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim blnChanged As Boolean


    '检查数据合法性
    If Val(txtSerPort.Text) = 0 Then
        MsgBox "服务器端口号不能为0或空值，请输入正确的值", , "提示"
        Exit Sub
    End If
    
    If Val(txtInterval.Text) = 0 Then
        MsgBox "数据变动通知发送频率不能为0或空值，请输入正确的值", , "提示"
        Exit Sub
    End If
    
    If chkJob.Value = 1 And Val(txtJob.Text) = 0 Then
        MsgBox "自动作业周期不能为0或空值，请输入正确的值", , "提示"
        Exit Sub
    End If
    
    '保存服务器信息
    If txtSerPort.Text <> glngPort Or chkLog.Value <> gintLog Or txtInterval.Text <> gintInterval Then
        blnChanged = ChangeServerSet2DB(Val(txtSerPort.Text), chkLog.Value, Val(txtInterval.Text))
    End If
    
    If (chkLogin.Value = 1) <> mblnLogin Then
        Call zlDatabase.SetPara("网络断网自动重连", chkLogin.Value)
    End If
    
    '保存客户端配置
    If mlngPortS <> txtPortStart.Text Or mlngPortE <> txtPortEnd.Text Or mlngCheckInterval <> txtCheck.Text Then
       blnChanged = ChangeClientSet2DB(txtPortStart.Text, txtPortEnd.Text, txtCheck.Text)
    End If
    
    '保存自动作业设置
    If (chkJob.Value = 1) <> mblnJob Or Val(txtJob.Text) <> mintInterval Then
        If GetZltools Then '由于自动作业归属于zltools用户,这里需要获取zltools连接
            If mblnJob = False Then
                '之前没有开启自动作业,需要重新开启
                blnChanged = ChangeJobSet2DB(1, Val(txtJob.Text))
            Else
                '之前已经开启
                If chkJob.Value = 1 Then
                    '修改频率
                    blnChanged = ChangeJobSet2DB(2, Val(txtJob.Text))
                Else
                    '取消自动任务
                    blnChanged = ChangeJobSet2DB(3, Val(txtJob.Text))
                End If
            End If
        End If
    End If
    If blnChanged Then
        MsgBox "参数修改成功，重新启动后生效。", , "提示"
    End If
    Unload Me
End Sub

Private Sub txtCheck_KeyPress(KeyAscii As Integer)
    OnlyIntCK KeyAscii
End Sub

Private Sub txtJob_GotFocus()
    txtJob.SelStart = Len(txtJob.Text)
End Sub

Private Sub txtJob_KeyPress(KeyAscii As Integer)
    OnlyIntCK KeyAscii
End Sub

Private Sub txtPortEnd_GotFocus()
    txtPortEnd.SelStart = Len(txtPortEnd.Text)
End Sub

Private Sub txtPortEnd_KeyPress(KeyAscii As Integer)
    OnlyIntCK KeyAscii
End Sub

Private Sub txtPortStart_GotFocus()
    txtPortStart.SelStart = Len(txtPortStart.Text)
End Sub

Private Sub txtPortStart_KeyPress(KeyAscii As Integer)
    OnlyIntCK KeyAscii
End Sub

Private Sub txtSerPort_GotFocus()
    txtSerPort.SelStart = Len(txtSerPort.Text)
End Sub

Private Sub txtSerPort_KeyPress(KeyAscii As Integer)
    OnlyIntCK KeyAscii
End Sub
