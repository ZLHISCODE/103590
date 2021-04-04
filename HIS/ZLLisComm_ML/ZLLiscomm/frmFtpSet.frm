VERSION 5.00
Begin VB.Form frmFtpSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FTP设置"
   ClientHeight    =   4755
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5460
   Icon            =   "frmFtpSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picReadme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   90
      ScaleHeight     =   1485
      ScaleWidth      =   5220
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   5250
      Begin VB.Image imgNote 
         Height          =   240
         Index           =   2
         Left            =   75
         Picture         =   "frmFtpSet.frx":000C
         Top             =   75
         Width           =   240
      End
      Begin VB.Label lblFTP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmFtpSet.frx":0596
         ForeColor       =   &H00FF0000&
         Height          =   1080
         Left            =   390
         TabIndex        =   11
         Top             =   240
         Width           =   4695
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox txtFtpPath 
      Height          =   300
      Left            =   1590
      TabIndex        =   5
      Top             =   3330
      Width           =   3000
   End
   Begin VB.TextBox txtDevAdress 
      Height          =   300
      Left            =   1605
      TabIndex        =   4
      Top             =   2865
      Width           =   3000
   End
   Begin VB.TextBox txtPassWord 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1605
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
      Width           =   3000
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   1605
      TabIndex        =   2
      Top             =   1920
      Width           =   3000
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   3195
      TabIndex        =   1
      Top             =   3975
      Width           =   1100
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   1230
      TabIndex        =   0
      Top             =   3975
      Width           =   1100
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FTP路径"
      Height          =   180
      Left            =   840
      TabIndex        =   9
      Top             =   3405
      Width           =   630
   End
   Begin VB.Label lblIP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FTP地址"
      Height          =   180
      Left            =   840
      TabIndex        =   8
      Top             =   2925
      Width           =   630
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FTP密码"
      Height          =   180
      Left            =   840
      TabIndex        =   7
      Top             =   2445
      Width           =   630
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FTP用户"
      Height          =   180
      Left            =   840
      TabIndex        =   6
      Top             =   1965
      Width           =   630
   End
End
Attribute VB_Name = "frmFtpSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strFtp As String
    strFtp = gobjDatabase.GetPara("FTP设置", glngSys, 1208, "")
    If UBound(Split(strFtp, ";")) >= 3 Then
        txtUser = Split(strFtp, ";")(0)
        txtPassWord = Split(strFtp, ";")(1)
        txtDevAdress = Split(strFtp, ";")(2)
        txtFtpPath = Split(strFtp, ";")(3)
    End If
End Sub

Private Sub OKButton_Click()
    Dim strFTPSet As String, strOpenFTP As String
    If Trim(txtUser) = "" And Trim(txtPassWord) = "" And Trim(txtDevAdress) = "" And Trim(txtFtpPath) = "" Then
        '清空设置
        Call gobjDatabase.SetPara("FTP设置", "", glngSys, 1208, InStr(";" & gstrPrivs & ";", ";通讯参数设置;") > 0)
        Unload Me
    Else
        strOpenFTP = TestFTP(txtUser, txtPassWord, txtDevAdress, txtFtpPath)
        If strOpenFTP = "" Then
            '保存参数
            strFTPSet = txtUser & ";" & txtPassWord & ";" & txtDevAdress & ";" & txtFtpPath
            Call gobjDatabase.SetPara("FTP设置", strFTPSet, glngSys, 1208, InStr(";" & gstrPrivs & ";", ";通讯参数设置;") > 0)
            Unload Me
        Else
            MsgBox strOpenFTP, vbInformation, Me.Caption
        End If
    End If
End Sub


