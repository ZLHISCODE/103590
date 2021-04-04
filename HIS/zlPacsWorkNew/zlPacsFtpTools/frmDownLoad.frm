VERSION 5.00
Begin VB.Form frmDownLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "文件下载"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5880
   Icon            =   "frmDownLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5880
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   420
      Left            =   4200
      TabIndex        =   5
      Top             =   1680
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下载"
      Height          =   420
      Left            =   3000
      TabIndex        =   4
      Top             =   1680
      Width           =   1100
   End
   Begin VB.TextBox txtLocal 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtFtp 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "/"
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label lblHint 
      AutoSize        =   -1  'True
      Caption         =   "（Ftp文件路径指FTP中要下载的文件的FTP路径，文件名需带上后缀）"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   5490
   End
   Begin VB.Label lblLocal 
      AutoSize        =   -1  'True
      Caption         =   "本地存储路径"
      Height          =   180
      Left            =   150
      TabIndex        =   2
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label lblFtp 
      AutoSize        =   -1  'True
      Caption         =   "Ftp文件路径"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   990
   End
End
Attribute VB_Name = "frmDownLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mstrFileName As String

Public Event DoDownLoad(ByVal strLocal As String, ByVal strFile As String)

Private Sub cmdCancel_Click()
    On Error GoTo errHandle

    Unload Me
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdLocal_Click()
    On Error GoTo errHandle
    
    dlgDown.FileName = mstrFileName
    dlgDown.ShowSave
    
    txtLocal.Text = dlgDown.FileName
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandle
    Dim i As Long
    
    If Len(Trim(txtFtp.Text)) = 0 Then
        MsgBox "请先选择需要下载的FTP文件", vbInformation, Me.Caption
        Exit Sub
    End If

    If Len(Trim(txtLocal.Text)) = 0 Then
        MsgBox "请先选择需要FTP文件保存的本地路径", vbInformation, Me.Caption
        Exit Sub
    End If
    
 
    RaiseEvent DoDownLoad(Trim(txtLocal.Text), txtFtp.Text)
 
    
    MsgBox "下载测试完成。"
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdUp_Click()
    On Error GoTo errHandle
    
    dlgDown.ShowOpen
    
    txtFtp.Text = dlgDown.FileName
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Load()
    txtLocal.Text = App.Path & "\Download\"
End Sub

Private Sub txtFtp_Validate(Cancel As Boolean)
    Dim arrFile() As String

    On Error GoTo errHandle
    
    If Len(Trim(txtFtp.Text)) > 0 Then
        arrFile = Split(Trim(txtFtp.Text), "/")
        
        mstrFileName = IIf(InStr(arrFile(UBound(arrFile)), ".") > 0, arrFile(UBound(arrFile)), "")
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub
