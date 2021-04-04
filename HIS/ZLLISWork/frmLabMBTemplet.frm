VERSION 5.00
Begin VB.Form frmLabMBTemplet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "保存模板"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3870
      TabIndex        =   6
      Top             =   1380
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2310
      TabIndex        =   5
      Top             =   1380
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   105
      Left            =   -30
      TabIndex        =   4
      Top             =   1140
      Width           =   5385
   End
   Begin VB.TextBox txt名称 
      Height          =   315
      Left            =   1050
      TabIndex        =   3
      Top             =   690
      Width           =   2865
   End
   Begin VB.TextBox txt编号 
      Height          =   315
      Left            =   1050
      TabIndex        =   1
      Top             =   180
      Width           =   1485
   End
   Begin VB.Label lbl名称 
      AutoSize        =   -1  'True
      Caption         =   "名称:"
      Height          =   180
      Left            =   450
      TabIndex        =   2
      Top             =   757
      Width           =   450
   End
   Begin VB.Label lbl编号 
      AutoSize        =   -1  'True
      Caption         =   "编号:"
      Height          =   180
      Left            =   450
      TabIndex        =   0
      Top             =   247
      Width           =   450
   End
End
Attribute VB_Name = "frmLabMBTemplet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrResult As String
Public Sub ShowMe(Objfrm As Object, strResult As String)
    mstrResult = strResult
    Me.Show vbModal, Objfrm
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '保存模板
    
    '检查数据是否正确
    If Len(Trim(Me.txt编号)) < 1 Then
        MsgBox "请输入编号!", vbInformation
        Me.txt编号.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(Me.txt编号) = False Then
        MsgBox "编号必须为数据，请修改!", vbInformation
        Me.txt编号.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Me.txt名称)) < 1 Then
        MsgBox "请输入名称!", vbInformation
        Me.txt名称.SetFocus
        Exit Sub
    End If
    
    On Error GoTo errH
   
    gstrSql = "Zl_检验酶标模板_Insert(" & zlDatabase.GetNextId("检验酶标模板") & "," & Val(Me.txt编号) & ",'" & _
              Me.txt名称 & "','" & Split(mstrResult, "|")(0) & "','" & Mid(mstrResult, InStr(mstrResult, "|") + 1) & "')"
    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
    
    MsgBox "保存完成!", vbInformation
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    gstrSql = "select nvl(max(编号),0)+ 1  as 编号 from 检验酶标模板"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.txt编号.Text = rsTmp(0)
End Sub

Private Sub txt编号_GotFocus()
    Me.txt编号.SelStart = 0: Me.txt编号.SelLength = Len(Me.txt编号.Text)
End Sub

Private Sub txt编号_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
    Else
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = Len(Me.txt名称.Text)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
