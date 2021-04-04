VERSION 5.00
Begin VB.Form frmToolsPwd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "工具所有者密码"
   ClientHeight    =   1935
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -180
      TabIndex        =   3
      Top             =   1335
      Width           =   4965
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1785
      TabIndex        =   1
      Top             =   1470
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2985
      TabIndex        =   2
      Top             =   1470
      Width           =   1100
   End
   Begin VB.TextBox txtPwd 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1785
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "zltools"
      Top             =   765
      Width           =   2055
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   120
      Picture         =   "frmToolsPwd.frx":0000
      Top             =   135
      Width           =   720
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Left            =   1290
      TabIndex        =   5
      Top             =   825
      Width           =   360
   End
   Begin VB.Label lblNote 
      Caption         =   "    请输入正确的工具所有者密码，以便建立连接。"
      Height          =   375
      Left            =   930
      TabIndex        =   4
      Top             =   135
      Width           =   3240
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmToolsPwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrUser As String
Private mstrPass As String

Public Function GetToolsPwd(Optional ByVal strUser As String = "ZlTOOLS") As String
    Dim cnTemp As New ADODB.Connection

    mstrUser = strUser
    mstrPass = ""
    
    On Error Resume Next
    
    With cnTemp
        If mstrUser = "ZLTOOLS" Then
            .Provider = "OraOLEDB.Oracle"
            .Open Trim(gstrServer), "ZLTOOLS", "ZLTOOLS"
            '.Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrServer, "ZLTOOLS", "ZLTOOLS"
            If Err = 0 Then
                .Close
                GetToolsPwd = "ZLTOOLS"
                Unload Me
                Exit Function
            End If
            If .State = adStateOpen Then .Close
            Err.Clear
        Else
            Me.Caption = Replace(Me.Caption, "工具所有者", UCase(strUser))
            lblNote.Caption = Replace(lblNote.Caption, "工具所有者", UCase(strUser))
            txtPwd.Text = ""
        End If
        
        Me.Show 1
        GetToolsPwd = mstrPass
    End With
End Function

Private Sub cmdCancel_Click()
    mstrPass = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim cnTemp As New ADODB.Connection

    Err.Clear: On Error Resume Next
    With cnTemp
        .Provider = "OraOLEDB.Oracle"
        .Open Trim(gstrServer), mstrUser, Trim(txtPwd.Text)
        If Err <> 0 Then
            MsgBox "密码输入错误，不能继续。" & vbNewLine & Err.Description, vbExclamation, gstrSysName
            If .State = adStateOpen Then .Close
            Err.Clear
            txtPwd.SetFocus
            Exit Sub
        End If
        .Close
    End With
        
    mstrPass = Trim(Me.txtPwd.Text)
    Unload Me
End Sub

Private Sub txtPwd_GotFocus()
    Me.txtPwd.SelStart = 0: Me.txtPwd.SelLength = 100
End Sub
