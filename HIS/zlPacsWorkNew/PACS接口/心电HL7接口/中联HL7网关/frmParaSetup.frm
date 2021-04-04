VERSION 5.00
Begin VB.Form frmParaSetup 
   Caption         =   "参数设置"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7365
   Icon            =   "frmParaSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   7365
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Caption         =   "消息接收方式"
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7095
      Begin VB.OptionButton optInputDataType 
         Caption         =   "文件"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   2655
      End
      Begin VB.Frame frmFileInput 
         Height          =   1935
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   6735
         Begin VB.TextBox txtFileBackupDir 
            Height          =   375
            Left            =   1320
            TabIndex        =   14
            Top             =   1350
            Width           =   5200
         End
         Begin VB.TextBox txtFileSuffix 
            Height          =   375
            Left            =   1320
            TabIndex        =   12
            Top             =   825
            Width           =   5200
         End
         Begin VB.TextBox txtFileDir 
            Height          =   375
            Left            =   1320
            TabIndex        =   10
            Top             =   300
            Width           =   5200
         End
         Begin VB.Label Label4 
            Caption         =   "备份目录："
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1410
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "文件后缀"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   885
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "文件目录："
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.OptionButton optInputDataType 
         Caption         =   "socket（默认）"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   370
      Left            =   6120
      TabIndex        =   2
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   370
      Left            =   4440
      TabIndex        =   1
      Top             =   4440
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "基础设置"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   7095
      Begin VB.TextBox txtTimeOut 
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "超时：                    秒"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmParaSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '保存参数
    gintTimeOutMax = Val(txtTimeOut.Text)
    SaveSetting "ZLSOFT", gstrRegPath, "超时", gintTimeOutMax
    
    If optInputDataType(0).Value = True Then
        gintInputDataType = 0
    Else
        gintInputDataType = 1
    End If
    SaveSetting "ZLSOFT", gstrRegPath, "接收消息方式", gintInputDataType
    
    gstrFileDir = txtFileDir.Text
    SaveSetting "ZLSOFT", gstrRegPath, "文件消息目录", gstrFileDir
    
    gstrFileSuffix = txtFileSuffix.Text
    SaveSetting "ZLSOFT", gstrRegPath, "文件消息后缀", gstrFileSuffix
    
    gstrFileBackupDir = txtFileBackupDir.Text
    SaveSetting "ZLSOFT", gstrRegPath, "文件消息备份目录", gstrFileBackupDir
    

    Unload Me
End Sub

Private Sub Form_Load()
    
    '从注册表读取超时设置
    txtTimeOut.Text = gintTimeOutMax
    If gintInputDataType = 1 Then
        optInputDataType(1).Value = True
        frmFileInput.Enabled = True
    Else
        optInputDataType(0).Value = True
        frmFileInput.Enabled = False
    End If
    
    txtFileDir.Text = gstrFileDir
    txtFileSuffix.Text = gstrFileSuffix
    txtFileBackupDir.Text = gstrFileBackupDir
        
End Sub

Public Sub zlSohwMe(frmParent As Form)
    Me.Show 1, frmParent
End Sub

Private Sub optInputDataType_Click(Index As Integer)
    If Index = 0 Then
        frmFileInput.Enabled = False
    Else
        frmFileInput.Enabled = True
    End If
End Sub
