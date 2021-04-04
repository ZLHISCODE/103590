VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAutoArchive 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "自动策略设置"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame frmAutoArchive 
      Caption         =   "时间规则设置"
      Height          =   2415
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   3975
      Begin VB.TextBox txtDay 
         Height          =   300
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "1"
         Top             =   360
         Width           =   1275
      End
      Begin VB.OptionButton optDay 
         Caption         =   "每天"
         Height          =   300
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "每月"
         Height          =   300
         Left            =   360
         TabIndex        =   8
         Top             =   890
         Width           =   1095
      End
      Begin VB.TextBox txtMonth 
         Height          =   300
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "1"
         Top             =   870
         Width           =   1455
      End
      Begin VB.ComboBox cobTimeArchiveStyle 
         Height          =   315
         ItemData        =   "frmAutoArchive.frx":0000
         Left            =   1440
         List            =   "frmAutoArchive.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpTime 
         Height          =   300
         Left            =   1440
         TabIndex        =   7
         Top             =   1350
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16646146
         CurrentDate     =   38226
      End
      Begin MSComCtl2.UpDown udMonth 
         Height          =   300
         Left            =   2880
         TabIndex        =   9
         Top             =   870
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udDay 
         Height          =   300
         Left            =   660
         TabIndex        =   15
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "间隔                 天"
         Height          =   180
         Left            =   1440
         TabIndex        =   16
         Top             =   420
         Width           =   2070
      End
      Begin VB.Label Label2 
         Caption         =   "号                            天"
         Height          =   195
         Left            =   3330
         TabIndex        =   12
         Top             =   930
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "归档时间"
         Height          =   180
         Left            =   360
         TabIndex        =   11
         Top             =   1395
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "归档方式"
         Height          =   180
         Left            =   360
         TabIndex        =   10
         Top             =   1860
         Width           =   720
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
      Height          =   350
      Left            =   3600
      TabIndex        =   3
      Top             =   3480
      Width           =   1100
   End
   Begin VB.CommandButton Command2 
      Caption         =   "应用"
      Height          =   350
      Left            =   1920
      TabIndex        =   2
      Top             =   3480
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   350
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   1100
   End
   Begin VB.CheckBox chkAutoArchive 
      Caption         =   "启用自动归档(将自动归档所有记录)"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   4425
   End
End
Attribute VB_Name = "frmAutoArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkAutoArchive_Click()
    bAutoArchive = IIf(Me.chkAutoArchive.Value = 1, True, False)
    If bAutoArchive = False Then
        Me.frmAutoArchive.Enabled = False
    Else
        Me.frmAutoArchive.Enabled = True
    End If
End Sub


Private Sub Command1_Click()
    subApply
    Unload Me
End Sub

Private Sub Command2_Click()
    subApply
End Sub
Private Sub subApply()
    '将现有策略保存到临时变量
    If Me.chkAutoArchive = 1 Then       '设置了自动归档策略
        bAutoArchive = True
        '处理时间归档策略
        If Me.optDay.Value = True Then
            strTimePolicy = "time,day," & Me.txtDay.Text & "," & Me.dtpTime.Hour & ":" & _
                  Me.dtpTime.Minute & ":" & Me.dtpTime.Second & "," & _
                  Me.cobTimeArchiveStyle.ListIndex & ",1"
        Else
            strTimePolicy = "time,month," & Me.txtMonth.Text & "," & Me.dtpTime.Hour & ":" & _
                  Me.dtpTime.Minute & ":" & Me.dtpTime.Second & "," & _
                  Me.cobTimeArchiveStyle.ListIndex & ",1"
        End If
    Else                        '设置为没有自动归档策略
        bAutoArchive = False
        strTimePolicy = "time,N/A"
    End If
    '将临时变量内容保存到注册表
    SaveSetting "ZLSOFT", "公共模块\归档管理", "时间归档策略", strTimePolicy
    SaveSetting "ZLSOFT", "公共模块\归档管理", "使用自动归档", CStr(bAutoArchive)
End Sub
Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strTempPolicy() As String   '暂存被解析出来的自动归档策略
    Dim strTempTime() As String     '暂存被解析出来的自动归档策略中的归档时间
    
    If bAutoArchive = True Then    '有自动归档策略
        Me.chkAutoArchive.Value = 1
        '解析时间策略
        strTempPolicy = Split(strTimePolicy, ",")
        If UCase(strTempPolicy(1)) = "DAY" Then
            Me.optDay.Value = True
            Me.txtDay.Text = strTempPolicy(2)
        ElseIf UCase(strTempPolicy(1)) = "MONTH" Then
            Me.optMonth.Value = True
            Me.txtMonth.Text = strTempPolicy(2)
        End If
        strTempTime = Split(strTempPolicy(3), ":")
        Me.dtpTime.Hour = strTempTime(0)
        Me.dtpTime.Minute = strTempTime(1)
        Me.dtpTime.Second = strTempTime(2)
        If strTempPolicy(4) = "1" And strTempPolicy(5) = "1" Then
            Me.cobTimeArchiveStyle.ListIndex = 1        '删除且归档
        ElseIf strTempPolicy(4) = "0" Then
            Me.cobTimeArchiveStyle.ListIndex = 0        '只归档
        End If
    Else    '没有自动归档策略
        Me.chkAutoArchive.Value = 0
    End If
End Sub

Private Sub udDay_DownClick()
    Me.txtDay.Text = Val(Me.txtDay.Text) - 1
    If Val(Me.txtDay.Text) < 1 Then Me.txtDay.Text = 31
End Sub

Private Sub udDay_UpClick()
    Me.txtDay.Text = Val(Me.txtDay.Text) + 1
    If Val(Me.txtDay.Text) > 31 Then Me.txtDay.Text = 1
End Sub

Private Sub udMonth_DownClick()
    Me.txtMonth.Text = Val(Me.txtMonth.Text) - 1
    If Val(Me.txtMonth.Text) < 1 Then Me.txtMonth.Text = 31
End Sub

Private Sub udMonth_UpClick()
    Me.txtMonth.Text = Val(Me.txtMonth.Text) + 1
    If Val(Me.txtMonth.Text) > 31 Then Me.txtMonth.Text = 1
End Sub
