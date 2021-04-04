VERSION 5.00
Begin VB.Form frmCureFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找"
   ClientHeight    =   2385
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4210
      TabIndex        =   3
      Top             =   1935
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找下一条(&N)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2640
      TabIndex        =   2
      Top             =   1935
      Width           =   1530
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   30
      TabIndex        =   4
      Top             =   1785
      Width           =   5565
   End
   Begin VB.ComboBox cboSource 
      Height          =   300
      Left            =   1860
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   930
      Width           =   3435
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "frmCureFind.frx":0000
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblComment 
      Caption         =   "    输入希望查找的参考目录的编码、名称、别名或者其简码。如存在多条，可依序查找下一条，直到找到你希望查找的项目。"
      Height          =   525
      Left            =   810
      TabIndex        =   6
      Top             =   135
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "(共查找到10条，当前为第1条)"
      Height          =   180
      Left            =   810
      TabIndex        =   5
      Top             =   1455
      Width           =   2430
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "查找内容(&F)"
      Height          =   180
      Left            =   810
      TabIndex        =   0
      Top             =   1005
      Width           =   990
   End
End
Attribute VB_Name = "frmCureFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFind As New ADODB.Recordset
Dim strFind As String
Dim intCount As Integer

Private Sub cboSource_GotFocus()
    Me.cboSource.SelStart = 0: Me.cboSource.SelLength = 100
End Sub

Private Sub cboSource_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If Trim(Me.cboSource.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFind_Click()
    Dim lng分类ID As Long, lng目录ID As Long
    Dim strTemp As String
    
    If Trim(Me.cboSource.Text) = "" Then
        MsgBox "请输入查找的内容", vbExclamation, gstrSysName
        Me.cboSource.SetFocus: Exit Sub
    End If
    strTemp = ""
    For intCount = 0 To Me.cboSource.ListCount
        strTemp = strTemp & ";" & Me.cboSource.List(intCount)
    Next
    If InStr(1, strTemp, ";" & Trim(Me.cboSource.Text)) = 0 Then
        Me.cboSource.AddItem Trim(Me.cboSource.Text), 0
    End If
    gstrSql = "select distinct L.分类ID,N.参考目录ID" & _
            " from 诊疗参考目录 L,诊疗参考别名 N" & _
            " where L.类型=[1] and L.ID=N.参考目录ID " & _
            "       and (L.编码 like [2] " & _
            "            or N.名称 like [3] " & _
            "            or N.简码 like [3])"
    Err = 0: On Error GoTo ErrHand

    If strFind <> gstrSql Or rsFind.State <> adStateOpen Then
        Set rsFind = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, Trim(Me.cboSource.Text) & "%", gstrMatch & Trim(Me.cboSource.Text) & "%")
        
        If rsFind.EOF Then
            MsgBox "不存在查找的参考！", vbExclamation, gstrSysName
            rsFind.Close: Me.cmdFind.Enabled = False: Me.lblnote.Caption = ""
            Me.cboSource.SetFocus: Exit Sub
        End If
        strFind = gstrSql
    Else
        rsFind.MoveNext
        If rsFind.EOF Then
            MsgBox "已查找到最后一条参考！", vbExclamation, gstrSysName
            rsFind.Close: Me.cboSource.Text = "": Me.cmdFind.Enabled = False: Me.lblnote.Caption = ""
            Me.cboSource.SetFocus: Exit Sub
        End If
    End If
    Me.lblnote.Caption = "(共查找到" & rsFind.RecordCount & "条，当前为第" & rsFind.AbsolutePosition & "条)"
    lng分类ID = rsFind!分类ID
    lng目录ID = rsFind!参考目录ID

    Call frmCureRefers.zlLocateItem(lng分类ID, lng目录ID)
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub Form_Activate()
    Me.Tag = Val(frmCureRefers.tvwClass.Tag) + 1
    Me.Caption = "参考查找..."
    Me.cboSource.SetFocus
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    strFind = ""
    Me.lblNote.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
