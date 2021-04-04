VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUpLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "文件上传"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5940
   Icon            =   "frmUpLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5940
   StartUpPosition =   1  '所有者中心
   Begin MSComDlg.CommonDialog dlgUp 
      Left            =   480
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   420
      Left            =   4200
      TabIndex        =   6
      Top             =   1680
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "上传"
      Height          =   420
      Left            =   2880
      TabIndex        =   5
      Top             =   1680
      Width           =   1100
   End
   Begin VB.CommandButton cmdLocal 
      Caption         =   "…"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtFtp 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Text            =   "/"
      Top             =   1200
      Width           =   3975
   End
   Begin VB.ListBox lstLocal 
      Height          =   780
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label lblFtp 
      AutoSize        =   -1  'True
      Caption         =   "Ftp相对路径"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   990
   End
   Begin VB.Label lblLocal 
      AutoSize        =   -1  'True
      Caption         =   "文件路径"
      Height          =   180
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmUpLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event DoUpLoad(ByVal strFtpRoad As String, arrFiles() As String)


Private Sub cmdCancel_Click()
    On Error GoTo errHandle
    
    Unload Me
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdLocal_Click()
    On Error GoTo errHandle
    
    dlgUp.ShowOpen
    
    If Len(dlgUp.FileName) > 0 Then
        If Not CheckRepeat(dlgUp.FileName) Then
            lstLocal.AddItem dlgUp.FileName
        Else
            MsgBox "文件已选择！", vbInformation, Me.Caption
        End If
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Function CheckRepeat(strFile As String) As Boolean
    Dim i As Long
    
    CheckRepeat = False
    
    For i = 0 To lstLocal.ListCount - 1
        If strFile = lstLocal.List(i) Then
            CheckRepeat = True
            Exit Function
        End If
    Next
End Function



Private Sub cmdOK_Click()
    Dim arrFiles() As String
    Dim i As Long
    
    On Error GoTo errHandle
    
    If lstLocal.ListCount > 0 Then
        ReDim arrFiles(0)
        For i = 0 To lstLocal.ListCount - 1
            ReDim Preserve arrFiles(UBound(arrFiles) + 1)
            
            arrFiles(UBound(arrFiles)) = lstLocal.List(i)
        Next
        
        RaiseEvent DoUpLoad(Trim(txtFtp.Text), arrFiles())
    Else
        MsgBox "请选择需要上传的文件。", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Unload Me
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub lstLocal_Click()
    Dim i As Long

    On Error GoTo errHandle
    
    For i = 0 To lstLocal.ListCount - 1
        If lstLocal.Selected(i) = True Then
            lstLocal.ToolTipText = lstLocal.List(i)
            Exit Sub
        End If
    Next
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

