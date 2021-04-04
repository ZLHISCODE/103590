VERSION 5.00
Begin VB.Form frmDiagRefFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找"
   ClientHeight    =   2280
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6045
   ControlBox      =   0   'False
   Icon            =   "frmDiagRefFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkUpper 
      Caption         =   "区分大小写(&U)"
      Height          =   240
      Left            =   105
      TabIndex        =   4
      Top             =   1875
      Width           =   1650
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "全部替换(&A)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   2
      Left            =   4380
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "替换(&R)..."
      Enabled         =   0   'False
      Height          =   350
      Index           =   1
      Left            =   4380
      TabIndex        =   7
      Top             =   960
      Width           =   1530
   End
   Begin VB.TextBox txtObject 
      Height          =   300
      Left            =   1155
      TabIndex        =   3
      Top             =   645
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4380
      TabIndex        =   6
      Top             =   465
      Width           =   1530
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找下一行(&N)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   0
      Left            =   4380
      TabIndex        =   5
      Top             =   105
      Width           =   1530
   End
   Begin VB.ComboBox cboSource 
      Height          =   300
      Left            =   1155
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2970
   End
   Begin VB.Label lblObject 
      AutoSize        =   -1  'True
      Caption         =   "替换为(&P)"
      Height          =   180
      Left            =   105
      TabIndex        =   2
      Top             =   705
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "查找内容(&F)"
      Height          =   180
      Left            =   105
      TabIndex        =   0
      Top             =   195
      Width           =   990
   End
End
Attribute VB_Name = "frmDiagRefFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFind As String
Dim lngCount As Long
Dim lngCurRow As Long

Public frmParent As Form

Private Sub cboSource_GotFocus()
    Call zlCommFun.OpenIme(True)
    Me.cboSource.SelStart = 0: Me.cboSource.SelLength = 100
End Sub

Private Sub cboSource_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Me.cmdFind(0).Enabled = True
    Me.cmdFind(1).Enabled = True
    Me.cmdFind(2).Enabled = True
End Sub

Private Sub chkUpper_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Dim strTemp As String, lngResult As Long
        
    If Index = 1 And Me.cmdFind(1).Caption = "替换(&R)..." Then
        Me.Tag = "替换"
        Me.lblObject.Visible = True
        Me.txtObject.Visible = True
        Me.cmdFind(2).Visible = True
        Me.cmdFind(1).Caption = "替换(&R)"
        Me.txtObject.SetFocus
        Exit Sub
    End If
    
    If Me.cboSource.Text = "" Then
        MsgBox "请输入查找的内容", vbExclamation, gstrSysName
        Me.cboSource.SetFocus: Exit Sub
    End If
    strTemp = ""
    For lngCount = 0 To Me.cboSource.ListCount
        strTemp = strTemp & ";" & Me.cboSource.List(lngCount)
    Next
    If InStr(1, strTemp, ";" & Me.cboSource.Text) = 0 Then
        Me.cboSource.AddItem Me.cboSource.Text, 0
    End If

    With Me.frmParent
        If Index <> 2 Then
            If strFind <> Me.cboSource.Text Then
                strFind = Me.cboSource.Text
                lngCurRow = .hgdRefer.FixedRows
            End If
            If Index = 1 And lngCurRow > .hgdRefer.FixedRows Then
                Call .zlWordReplace(lngCurRow - 1, Me.cboSource.Text, Me.txtObject.Text)
            End If
            For lngCount = lngCurRow To .hgdRefer.Rows - 1
                lngResult = .zlWordSelect(lngCount, Me.cboSource.Text)
                If lngResult <> 0 Then lngCurRow = lngCount + 1: Exit Sub
            Next
            MsgBox "已经查找到最后一行！", vbExclamation, gstrSysName
            lngCurRow = .hgdRefer.FixedRows
        Else
            For lngCount = .hgdRefer.FixedRows To .hgdRefer.Rows - 1
                Call .zlWordReplace(lngCount, Me.cboSource.Text, Me.txtObject.Text)
            Next
        End If
    End With
End Sub



Private Sub Form_Activate()
    If Me.Tag = "替换" Then
        Me.lblObject.Visible = True
        Me.txtObject.Visible = True
        Me.cmdFind(2).Visible = True
        Me.cmdFind(1).Caption = "替换(&R)"
    Else
        Me.lblObject.Visible = False
        Me.txtObject.Visible = False
        Me.cmdFind(2).Visible = False
        Me.cmdFind(1).Caption = "替换(&R)..."
    End If
    Me.cboSource.SetFocus
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    strFind = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub txtObject_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub
