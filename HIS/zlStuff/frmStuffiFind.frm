VERSION 5.00
Begin VB.Form frmStuffFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找"
   ClientHeight    =   2385
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5595
   ControlBox      =   0   'False
   Icon            =   "frmStuffiFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
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
      Left            =   1905
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   930
      Width           =   3405
   End
   Begin VB.Label lblComment 
      Caption         =   "    输入希望查找的材料编码、名称或其简码。如存在多条，可依序查找下一条，直到找到你希望查找的材料。"
      Height          =   525
      Left            =   855
      TabIndex        =   6
      Top             =   135
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "(共查找到10条，当前为第1条)"
      Height          =   180
      Left            =   855
      TabIndex        =   5
      Top             =   1455
      Width           =   2430
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "frmStuffiFind.frx":058A
      Top             =   225
      Width           =   480
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "查找内容(&F)"
      Height          =   180
      Left            =   855
      TabIndex        =   0
      Top             =   1005
      Width           =   990
   End
End
Attribute VB_Name = "frmStuffFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFind As New ADODB.Recordset
Dim strFind As String
Dim intCount As Integer
Private mbln显示停用材料 As Boolean

Private Sub cboSource_Click()
    If Trim(Me.cboSource.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub cboSource_GotFocus()
    Me.cboSource.SelStart = 0: Me.cboSource.SelLength = 100
End Sub

Private Sub cboSource_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboSource_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(Me.cboSource.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub


Private Sub cmdFind_Click()
    Dim lng分类id As Long, lng材料ID As Long
    Dim lng诊疗ID As Long
    Dim strTemp As String
    Dim strMach As String
    Dim strSerach As String
    
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
    
    strMach = Trim(Me.cboSource.Text)
    
   
    strSerach = " And (C.编码 Like [1] OR N.名称 Like [1] OR N.简码 LIKE upper([1]))"
    
    If IsNumeric(strMach) Then                          '如果是数字,则只取编码
        If Mid(gSystem_Para.Para_输入方式, 1, 1) = "1" Then strSerach = " And (C.编码 Like [1])"
        strMach = "" & GetMatchingSting(UCase(strMach)) & ""
    ElseIf zlStr.IsCharAlpha(strMach) Then          '输入全是字母时只匹配简码
        If Mid(gSystem_Para.Para_输入方式, 2, 1) = "1" Then strSerach = " And N.简码 Like [1] "
        strMach = "" & GetMatchingSting(UCase(strMach)) & ""
    ElseIf zlStr.IsCharChinese(strMach) Then
        strSerach = " And N.名称 Like [1] "
        strMach = "" & GetMatchingSting(strMach) & ""
    Else
        strMach = "" & GetMatchingSting(strMach) & ""
    End If
    
    
    gstrSQL = "" & _
        "   Select distinct I.分类ID,B.材料ID,B.诊疗ID" & _
        "   From 诊疗项目目录 I,收费项目别名 N,材料特性 B,收费项目目录 C" & _
        "   Where   I.类别='4' And I.id=b.诊疗id and b.材料ID=N.收费细目id and b.材料id=C.id " & strSerach
    If Not mbln显示停用材料 Then gstrSQL = gstrSQL & " And (C.撤档时间 Is NULL Or C.撤档时间 >=to_date('3000-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'))"
    
    
    err = 0: On Error GoTo ErrHand
    
    If strFind <> IIf(mbln显示停用材料, 1, 0) & ":" & strMach Or rsFind.State <> adStateOpen Then
        Set rsFind = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strMach)
        If rsFind.EOF Then
            MsgBox "不存在查找的卫材！", vbExclamation, gstrSysName
            rsFind.Close: Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
            Me.cboSource.SetFocus: Exit Sub
        End If
        strFind = IIf(mbln显示停用材料, 1, 0) & ":" & strMach
    Else
        rsFind.MoveNext
        If rsFind.EOF Then
            MsgBox "已查找到最后一条卫材！", vbExclamation, gstrSysName
            rsFind.Close: Me.cboSource.Text = "": Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
            Me.cboSource.SetFocus: Exit Sub
        End If
    End If

    Me.lblNote.Caption = "(共查找到" & rsFind.RecordCount & "条，当前为第" & rsFind.AbsolutePosition & "条)"
    lng分类id = Val(rsFind!分类id)
    lng材料ID = Val(rsFind!材料ID)
    lng诊疗ID = Val(zlStr.Nvl(rsFind!诊疗id))
    Call frmStuffMgr.zlLocateItem(lng分类id, lng诊疗ID, lng材料ID)
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Form_Activate()
    Me.Tag = 4: Me.Caption = "卫生材料查找..."
    Me.cboSource.SetFocus
End Sub

Private Sub Form_Load()
    strFind = ""
    Me.lblNote.Caption = ""
End Sub

Private Sub optMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
End Sub

Public Sub ShowMe(ByVal frmParent As Object, ByVal bln停用 As Boolean)
    mbln显示停用材料 = bln停用
    Me.Show , frmParent
End Sub
