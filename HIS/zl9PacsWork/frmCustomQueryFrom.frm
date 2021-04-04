VERSION 5.00
Begin VB.Form frmCustomQueryFrom 
   Caption         =   "编辑数据"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomQueryFrom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7455
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdInsertPar 
      Caption         =   "插入参数(&I)"
      Height          =   375
      Left            =   135
      TabIndex        =   3
      Top             =   5235
      Width           =   1095
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   375
      Left            =   5085
      TabIndex        =   2
      Top             =   5220
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   5220
      Width           =   1095
   End
   Begin VB.TextBox txtSql 
      Height          =   5040
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   105
      Width           =   7215
   End
End
Attribute VB_Name = "frmCustomQueryFrom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mblnIsOk As Boolean
Public mobjInputList As ucFlexGrid


Public Function ShowSqlFromWindow(ByVal strSqlFrom As String, objInputList As Object, owner As Object) As String
    ShowSqlFromWindow = strSqlFrom
    
    Set Me.mobjInputList = objInputList
    
    Me.mblnIsOk = False
    Me.txtSql.Text = strSqlFrom
    Me.txtSql.Locked = Me.mobjInputList.ReadOnly
    Me.cmdSure.Enabled = Not Me.mobjInputList.ReadOnly
    Me.cmdInsertPar.Enabled = Not Me.mobjInputList.ReadOnly
    
    If Me.txtSql.Locked Then
        Me.txtSql.BackColor = &H8000000F
    Else
        Me.txtSql.BackColor = &H80000005
    End If
    
    
    
    Call Me.Show(1, owner)
    
    If Me.mblnIsOk Then
        ShowSqlFromWindow = Me.txtSql.Text
    End If
End Function

Private Sub cmdCancel_Click()
On Error GoTo errHandle
    mblnIsOk = False
    
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdInsertPar_Click()
'插入参数
On Error GoTo errHandle
    Dim strPar As String
    Dim frmPar As New frmCustomInsertPar
    
    strPar = frmPar.ShowParameterWindow(mobjInputList, False, Me)
    If strPar <> "" Then
        txtSql.SelText = strPar
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdSure_Click()
On Error GoTo errHandle
    mblnIsOk = True
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    mblnIsOk = False
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
On Error Resume Next
    txtSql.Width = Me.ScaleWidth - txtSql.Left * 2
    txtSql.Height = Me.ScaleHeight - txtSql.Top * 2 - cmdSure.Height - 120
    
    cmdInsertPar.Top = txtSql.Top + txtSql.Height + 60
    
    cmdCancel.Left = txtSql.Width - cmdCancel.Width
    cmdCancel.Top = cmdInsertPar.Top
    
    cmdSure.Left = cmdCancel.Left - 60 - cmdSure.Width
    cmdSure.Top = cmdInsertPar.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errHandle
    Call SaveWinState(Me, App.ProductName)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
