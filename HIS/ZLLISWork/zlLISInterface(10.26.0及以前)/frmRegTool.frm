VERSION 5.00
Begin VB.Form frmRegTool 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "第三方LIS接口授权码生成工具"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   435
      Left            =   960
      TabIndex        =   3
      Top             =   1215
      Width           =   4320
   End
   Begin VB.CommandButton cmdGetReg 
      Caption         =   "产生授权文件"
      Height          =   480
      Left            =   3495
      TabIndex        =   2
      Top             =   1905
      Width           =   1620
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Text            =   "3000-01-01"
      Top             =   405
      Width           =   4290
   End
   Begin VB.TextBox txtUnti 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Text            =   "宝信医院信息测试系统"
      Top             =   780
      Width           =   4290
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "授权码"
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   6
      Top             =   1185
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "授权用户"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Top             =   840
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "截止日期"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Top             =   435
      Width           =   720
   End
End
Attribute VB_Name = "frmRegTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetReg_Click()
    If Trim(txtUnti.Text) = "" Then
        MsgBox "授权用户不能为空!"
        Exit Sub
    End If
    If Trim(txtDate.Text) = "" Then
        MsgBox "授权截止日期不能为空!"
        Exit Sub
    End If
    If Not IsDate(txtDate.Text) Or Not (txtDate.Text Like "####-##-##") Then
        MsgBox "授权截止日期格式不正确，请按yyyy-MM-dd格式输入!"
        Exit Sub
    End If
    Call GetRegInfo
End Sub
Private Function GetRegInfo() As Boolean

    Dim strLine As String, str日期 As String, date日期 As Date
    Dim strUnti As String, strCode As String, strKey As String
    Dim lngFile  As Long
    
    strKey = "陈东"
    strUnti = txtUnti.Text
    str日期 = txtDate.Text
    strCode = sha1(EncodeBase64String(strUnti & "|" & str日期 & "|" & strKey))
    txtCode = strCode
    lngFile = FreeFile
    Open App.Path & "\RegFile.ini" For Output Access Write As lngFile
    
    Print #lngFile, "授权截止日期=" & str日期 & vbNewLine & "授权码=" & strCode
    Close lngFile
    MsgBox "已生成授权文件到:" & App.Path & "\RegFile.ini"
End Function

