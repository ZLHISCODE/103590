VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "用户登录"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4140
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4140
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   2640
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Text            =   "MBFBJ"
      Top             =   1118
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "YMJTZYY2015AQA"
      Top             =   638
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "zlhis"
      Top             =   158
      Width           =   2415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "数据库"
      Height          =   210
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "密  码"
      Height          =   210
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   210
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error GoTo errHandle
    
    With gcnOracle
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & Text3.Text & "", Text1.Text, TranPasswd(Text2.Text)
    End With
    
    frmMain.Show
    
    Unload Me
    
    Exit Sub
errHandle:
    MsgBox "执行SQL错误！", vbInformation, ""
    Err = 0
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew
End Function



