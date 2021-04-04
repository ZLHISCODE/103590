VERSION 5.00
Begin VB.Form frmMediFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找"
   ClientHeight    =   2490
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5595
   ControlBox      =   0   'False
   Icon            =   "frmMediFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk规格 
      Alignment       =   1  'Right Justify
      Caption         =   "查找规格药品"
      Height          =   210
      Left            =   810
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4210
      TabIndex        =   6
      Top             =   2085
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找下一条(&N)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2640
      TabIndex        =   5
      Top             =   2085
      Width           =   1530
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   30
      TabIndex        =   4
      Top             =   1935
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
      Caption         =   "    输入希望查找的药品编码、名称、别名或者其简码。如存在多条，可依序查找下一条，直到找到你希望查找的药品。"
      Height          =   525
      Left            =   855
      TabIndex        =   7
      Top             =   135
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "(共查找到10条，当前为第1条)"
      Height          =   180
      Left            =   855
      TabIndex        =   3
      Top             =   1635
      Width           =   2430
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "frmMediFind.frx":058A
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
Attribute VB_Name = "frmMediFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFind As New ADODB.Recordset
Dim strFind As String
Dim intCount As Integer
Private mbln显示停用药品 As Boolean
Private mblnSelfMedi As Boolean  '自管药 true-自管药 false-不是自管药

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
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboSource_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(Me.cboSource.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub chk规格_Click()
    If Trim(Me.cboSource.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub chk规格_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub CmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFind_Click()
    Dim lng分类id As Long, lng药名id As Long, lng药品ID As Long
    Dim strTemp As String
    
    If Trim(Me.cboSource.Text) = "" Then
        MsgBox "请输入查找的内容", vbExclamation, gstrSysName
        On Error Resume Next
        Me.Show
        Me.cboSource.SetFocus: Exit Sub
    End If
    strTemp = ""
    For intCount = 0 To Me.cboSource.ListCount
         strTemp = strTemp & ";" & Me.cboSource.List(intCount)
    Next
    If InStr(1, strTemp, ";" & Trim(Me.cboSource.Text)) = 0 Then
        Me.cboSource.AddItem Trim(Me.cboSource.Text), 0
    End If
    
    If Me.chk规格.Value = 0 Then
        If mblnSelfMedi = True Then '自管药
            gstrSql = "Select Distinct i.分类id, i.Id As 药名id, 0 As 药品id" & vbNewLine & _
                    "From 诊疗项目目录 I, 诊疗项目别名 N, 药品特性 A" & vbNewLine & _
                    "Where i.Id = n.诊疗项目id And i.Id = a.药名id And a.临床自管药 = 1 And i.类别 = [1] And" & vbNewLine & _
                    "      (i.编码 Like [2] Or n.名称 Like [2] Or n.简码 Like [2])"

        Else
            gstrSql = "SELECT DISTINCT I.分类ID,I.ID AS 药名ID,0 AS 药品ID" & _
                    " FROM 诊疗项目目录 I,诊疗项目别名 N" & _
                    " WHERE I.ID=N.诊疗项目ID " & _
                    " AND I.类别=[1] " & _
                    " AND (I.编码 LIKE [2] " & _
                    "     OR N.名称 LIKE [3] " & _
                    "     OR N.简码 LIKE [3])"
        End If
    Else
        If mblnSelfMedi = True Then '自管药
            gstrSql = "Select Distinct i.分类id, i.Id As 药名id, d.Id As 药品id" & vbNewLine & _
                    "From 诊疗项目目录 I, 药品规格 T, 药品特性 A, 收费项目目录 D, 收费项目别名 N" & vbNewLine & _
                    "Where i.Id = t.药名id And i.Id = a.药名id And t.药品id = d.Id And t.药品id = n.收费细目id And i.类别 = [1] And" & vbNewLine & _
                    "      (d.编码 Like [2] Or n.名称 Like[2] Or n.简码 Like[2]) And a.临床自管药 = 1"

        Else
            gstrSql = "SELECT DISTINCT I.分类ID,I.ID AS 药名ID,D.ID AS 药品ID  " & _
                     " FROM 诊疗项目目录 I,药品规格 T,收费项目目录 D,收费项目别名 N " & _
                     " WHERE I.ID=T.药名ID And T.药品ID=D.ID AND T.药品ID=N.收费细目ID  " & _
                     " AND I.类别=[1] " & _
                     " AND (D.编码 LIKE [2] " & _
                     "     OR N.名称 LIKE [3] " & _
                     "     OR N.简码 LIKE [3])"
        End If
    End If
    If Not mbln显示停用药品 Then
        gstrSql = gstrSql & " And (I.撤档时间 Is NULL Or to_Char(I.撤档时间,'yyyy-MM-dd')='3000-01-01')"
        If Me.chk规格.Value = 1 Then gstrSql = gstrSql & " And (D.撤档时间 Is NULL Or to_Char(D.撤档时间,'yyyy-MM-dd')='3000-01-01')"
    End If
    
    Err = 0: On Error GoTo errHand
 
    If strFind <> chk规格.Value & ";" & Trim(Me.cboSource.Text) Or rsFind.State <> adStateOpen Then
        Set rsFind = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, gstrMatch & Trim(Me.cboSource.Text) & "%", gstrMatch & Trim(Me.cboSource.Text) & "%")
        
        If rsFind.EOF Then
            MsgBox "不存在查找的药品！", vbExclamation, gstrSysName
            On Error Resume Next
            Me.Show
            rsFind.Close: Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
            Me.cboSource.SetFocus
            Exit Sub
        End If
        strFind = chk规格.Value & ";" & Trim(Me.cboSource.Text)
    Else
        rsFind.MoveNext
        If rsFind.EOF Then
            MsgBox "已查找到最后一条药品！", vbExclamation, gstrSysName
            On Error Resume Next
            Me.Show
            rsFind.Close: Me.cboSource.Text = "": Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
            Me.cboSource.SetFocus
            Exit Sub
        End If
    End If
    Me.lblNote.Caption = "(共查找到" & rsFind.RecordCount & "条，当前为第" & rsFind.AbsolutePosition & "条)"
    lng分类id = rsFind!分类ID
    lng药名id = rsFind!药名ID
    lng药品ID = rsFind!药品id
           
    
    Me.Hide
    Call frmMediLists.zlLocateItem(lng分类id, lng药名id, lng药品ID)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub Form_Activate()
    Select Case Val(frmMediLists.tvwClass.Tag)
    Case 0
        Me.Tag = 5: Me.Caption = "西成药查找..."
    Case 1
        Me.Tag = 6: Me.Caption = "中成药查找..."
    Case 2
        Me.Tag = 7: Me.Caption = "中草药查找..."
    End Select
    Me.cboSource.SetFocus
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    strFind = ""
    Me.lblNote.Caption = ""
End Sub

Private Sub optMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Public Sub ShowMe(ByVal frmParent As Object, ByVal bln停用 As Boolean, ByVal blnSelMedi As Boolean)
    mbln显示停用药品 = bln停用
    mblnSelfMedi = blnSelMedi
    Me.Show , frmParent
End Sub

Public Sub FindNext()
    Call cmdFind_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
