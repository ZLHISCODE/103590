VERSION 5.00
Begin VB.Form frmMsgBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "frmMsgBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4275
      TabIndex        =   2
      Top             =   975
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CheckBox chkSkip 
      Caption         =   "不再提示(&S)"
      Height          =   195
      Left            =   375
      TabIndex        =   3
      Top             =   1050
      Width           =   1290
   End
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "否(&N)"
      Height          =   350
      Left            =   3075
      TabIndex        =   1
      Top             =   975
      Width           =   1100
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "是(&Y)"
      Height          =   350
      Left            =   1875
      TabIndex        =   0
      Top             =   975
      Width           =   1100
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   270
      Picture         =   "frmMsgBox.frx":000C
      Top             =   195
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMsgBox.frx":08D6
      Height          =   360
      Left            =   1005
      TabIndex        =   4
      Top             =   285
      Width           =   2700
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   270
      Picture         =   "frmMsgBox.frx":0918
      Top             =   195
      Width           =   480
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrInfo As String
Private mintType As Integer
Private mblnNoAsk As Boolean
Private mvResult As VbMsgBoxResult

Private Sub cmdYes_Click()
    If mintType = 0 Then
        mvResult = IIf(chkSkip.Value = 1, vbIgnore, vbYes)
    ElseIf mintType = 1 Then
        mvResult = vbYes
    End If
    Unload Me
End Sub

Private Sub cmdNo_Click()
    If mintType = 0 Then
        mvResult = IIf(chkSkip.Value = 1, vbCancel, vbNo)
    ElseIf mintType = 1 Then
        mvResult = vbNo
    End If
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    If mintType = 1 Then
        mvResult = vbCancel
    End If
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '点击窗体关闭按钮:同按下NO或Cancel处理
    If UnloadMode = vbFormControlMenu Then
        If mintType = 0 Then
            mvResult = IIf(chkSkip.Value = 1, vbCancel, vbNo)
        ElseIf mintType = 1 Then
            mvResult = vbCancel
        End If
    End If
End Sub

Private Sub Form_Activate()
    If cmdCancel.Visible Then
        cmdCancel.SetFocus
    ElseIf cmdNo.Visible Then
        cmdNo.SetFocus
    End If
    Beep
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyS And mintType = 0 Then
        chkSkip.Value = IIf(chkSkip.Value = 1, 0, 1)
    ElseIf KeyCode = vbKeyY And mintType = 0 Then
        Call cmdYes_Click
    ElseIf KeyCode = vbKeyN And mintType = 0 Then
        Call cmdNo_Click
    ElseIf KeyCode = vbKeyAdd And mintType = 1 Then
        Call cmdYes_Click
    ElseIf KeyCode = vbKeySubtract And mintType = 1 Then
        Call cmdNo_Click
    ElseIf KeyCode = vbKeyC And mintType = 1 Then
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    Caption = gstrSysName
    
    lblInfo.Caption = mstrInfo
    
    If lblInfo.Left + lblInfo.Width + 500 > 4500 Then
        Me.Width = lblInfo.Left + lblInfo.Width + 500
    Else
        Me.Width = 4500
    End If
    
    If lblInfo.Top + lblInfo.Height + 1150 > 1800 Then
        Me.Height = lblInfo.Top + lblInfo.Height + 1150
    Else
        Me.Height = 1800
    End If
    
    If mintType = 0 Then
        cmdYes.Visible = True
        cmdNo.Visible = Not mblnNoAsk
        cmdCancel.Visible = False
        cmdNo.Cancel = True
        
        cmdYes.Caption = IIf(mblnNoAsk, "确定(&O)", "是(&Y)")
        cmdNo.Caption = "否(&N)"
        
        imgIcon(0).Visible = Not mblnNoAsk
        imgIcon(1).Visible = mblnNoAsk
        
        chkSkip.Visible = True
    ElseIf mintType = 1 Then
        cmdYes.Visible = True
        cmdNo.Visible = True
        cmdCancel.Visible = True
        cmdCancel.Cancel = True
        
        cmdYes.Caption = "阳性(+)"
        cmdNo.Caption = "阴性(-)"
        
        imgIcon(0).Visible = True
        imgIcon(1).Visible = False
        
        chkSkip.Visible = False
    End If
    
    cmdYes.Top = Me.ScaleHeight - cmdYes.Height - 100
    cmdNo.Top = cmdYes.Top
    cmdCancel.Top = cmdYes.Top
    chkSkip.Top = cmdYes.Top + (cmdYes.Height - chkSkip.Height) / 2
    
    If mintType = 0 Then
        If mblnNoAsk Then
            cmdYes.Left = (Me.ScaleWidth - cmdYes.Width) / 2 + 200
            chkSkip.Left = cmdYes.Left - chkSkip.Width - 200
        Else
            cmdNo.Left = Me.ScaleWidth - cmdNo.Width - 200
            cmdYes.Left = cmdNo.Left - cmdYes.Width - 100
            chkSkip.Left = cmdYes.Left - chkSkip.Width - 200
        End If
    ElseIf mintType = 1 Then
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 200
        cmdNo.Left = cmdCancel.Left - cmdNo.Width - 100
        cmdYes.Left = cmdNo.Left - cmdYes.Width - 15
    End If
End Sub
