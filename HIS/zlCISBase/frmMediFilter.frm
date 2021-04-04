VERSION 5.00
Begin VB.Form frmMediFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   1905
      TabIndex        =   0
      Top             =   1278
      Width           =   3525
   End
   Begin VB.TextBox txtSim 
      Height          =   300
      Left            =   1905
      TabIndex        =   2
      Top             =   2055
      Width           =   3525
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1905
      TabIndex        =   1
      Top             =   1671
      Width           =   3525
   End
   Begin VB.CheckBox chk中草药 
      Caption         =   "中草药"
      Height          =   210
      Left            =   4365
      TabIndex        =   12
      Top             =   945
      Width           =   1035
   End
   Begin VB.CheckBox chk中成药 
      Caption         =   "中成药"
      Height          =   210
      Left            =   3142
      TabIndex        =   11
      Top             =   945
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.CheckBox chk西成药 
      Caption         =   "西成药"
      Height          =   210
      Left            =   1920
      TabIndex        =   10
      Top             =   945
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.CommandButton cmdFilter 
      Cancel          =   -1  'True
      Caption         =   "过滤(&F)"
      Height          =   350
      Left            =   3300
      TabIndex        =   4
      Top             =   2850
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4560
      TabIndex        =   5
      Top             =   2850
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   60
      TabIndex        =   9
      Top             =   2700
      Width           =   5760
   End
   Begin VB.CheckBox chk规格 
      Alignment       =   1  'Right Justify
      Caption         =   "过滤规格药品"
      Height          =   210
      Left            =   540
      TabIndex        =   3
      Top             =   2445
      Width           =   1545
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   135
      TabIndex        =   7
      Top             =   630
      Width           =   5760
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "输入药品编码"
      Height          =   180
      Left            =   555
      TabIndex        =   15
      Top             =   1335
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "输入药品简码"
      Height          =   180
      Left            =   555
      TabIndex        =   14
      Top             =   2115
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "输入药品名称"
      Height          =   180
      Left            =   555
      TabIndex        =   13
      Top             =   1725
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "选择药品材质"
      Height          =   180
      Left            =   585
      TabIndex        =   8
      Top             =   945
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   345
      Picture         =   "frmMediFilter.frx":0000
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblComment 
      Caption         =   "    输入希望过滤的药品的材质及药品编码、名称、别名或者其简码。如存在多条，则返回多条过滤结果。"
      Height          =   435
      Left            =   1080
      TabIndex        =   6
      Top             =   105
      Width           =   4500
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMediFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbln显示停用药品 As Boolean
Private mblnSelfMedi As Boolean  '自管药 true-自管药 false-不是自管药

Public Sub ShowMe(ByVal frmParent As Object, ByVal bln停用 As Boolean, ByVal blnSelfMedi As Boolean)
    mbln显示停用药品 = bln停用
    mblnSelfMedi = blnSelfMedi
    Me.Show , frmParent
End Sub

Private Sub chk规格_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub chk西成药_Click()
    If chk西成药.Value = 1 Then
        chk中草药.Value = 0
    ElseIf chk中成药.Value = 0 Then
        chk中草药.Value = 1
    End If
End Sub

Private Sub chk中草药_Click()
    If chk中草药.Value = 1 Then
        chk西成药.Value = 0
        chk中成药.Value = 0
    Else
        chk西成药.Value = 1
        chk中成药.Value = 1
    End If
End Sub

Private Sub chk中成药_Click()
    If chk中成药.Value = 1 Then
        chk中草药.Value = 0
    ElseIf chk西成药.Value = 0 Then
        chk中草药.Value = 1
    End If
End Sub


Private Sub CmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFilter_Click()
    Dim rs As New ADODB.Recordset
    Dim strKind As String
    Dim strKind1 As String
    Dim strResult As String
    Dim n As Long
    Dim strCondition As String
    
    If Len(Trim(Me.txtCode.Text)) > 0 Then
        If Me.chk规格.Value = 0 Then
            strCondition = " AND I.编码 LIKE [1] "
        Else
            strCondition = " AND D.编码 LIKE [1] "
        End If
    End If
    If Len(Trim(Me.txtName.Text)) > 0 Then
        strCondition = " AND N.名称 LIKE [2] "
    End If
    If Len(Trim(Me.txtSim.Text)) > 0 Then
        strCondition = " AND N.简码 LIKE [3] "
    End If
    
    If Len(strCondition) = 0 Then
        MsgBox "请输入药品信息", vbExclamation, gstrSysName
        On Error Resume Next
        Me.Show
        Me.txtCode.SetFocus: Exit Sub
    End If
    
    If chk中草药.Value = 1 Then
        strKind = "7"
        strKind1 = ",7,"
    ElseIf chk西成药.Value = 1 And chk中成药.Value = 1 Then
        strKind = "5,6"
        strKind1 = ",5,6,"
    ElseIf chk西成药.Value = 1 Then
        strKind = "5"
        strKind1 = ",5,"
    ElseIf chk中成药.Value = 1 Then
        strKind = "6"
        strKind1 = ",6,"
    End If
    
    If Me.chk规格.Value = 0 Then
        If mblnSelfMedi = True Then
            gstrSql = "Select Distinct i.分类id, i.Id As 药名id, 0 As 药品id" & vbNewLine & _
                    "From 诊疗项目目录 I, 诊疗项目别名 N, 药品特性 A" & vbNewLine & _
                    "Where i.Id = n.诊疗项目id And i.Id = a.药名id And Instr([4], ',' || i.类别 || ',') > 0 And a.临床自管药 = 1 " & strCondition

        Else
            gstrSql = "SELECT DISTINCT I.分类ID,I.ID AS 药名ID,0 AS 药品ID" & _
                    " FROM 诊疗项目目录 I,诊疗项目别名 N" & _
                    " WHERE I.ID=N.诊疗项目ID AND Instr([4], ','||I.类别||',') > 0 " & strCondition
        End If

    Else
        If mblnSelfMedi = True Then
            gstrSql = "SELECT DISTINCT I.分类ID,I.ID AS 药名ID,D.ID AS 药品ID  " & _
                     " FROM 诊疗项目目录 I,药品规格 T,收费项目目录 D,收费项目别名 N,药品特性 A " & _
                     " WHERE I.ID=T.药名ID And T.药品ID=D.ID AND T.药品ID=N.收费细目ID AND i.ID=A.药名id and a.临床自管药=1 AND Instr([4], ','||I.类别||',') > 0 " & strCondition
        Else
            gstrSql = "SELECT DISTINCT I.分类ID,I.ID AS 药名ID,D.ID AS 药品ID  " & _
                     " FROM 诊疗项目目录 I,药品规格 T,收费项目目录 D,收费项目别名 N " & _
                     " WHERE I.ID=T.药名ID And T.药品ID=D.ID AND T.药品ID=N.收费细目ID AND Instr([4], ','||I.类别||',') > 0 " & strCondition
        End If
                 
    End If
    If Not mbln显示停用药品 Then
        gstrSql = gstrSql & " And (I.撤档时间 Is NULL Or to_Char(I.撤档时间,'yyyy-MM-dd')='3000-01-01')"
        If Me.chk规格.Value = 1 Then gstrSql = gstrSql & " And (D.撤档时间 Is NULL Or to_Char(D.撤档时间,'yyyy-MM-dd')='3000-01-01')"
    End If

    Err = 0: On Error GoTo errHand
    
    Set rs = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Trim(Me.txtCode.Text) & "%", gstrMatch & Trim(Me.txtName.Text) & "%", gstrMatch & Trim(Me.txtSim.Text) & "%", strKind1)
    
    With rs
        If .EOF Then
            MsgBox "没有找到药品信息！", vbExclamation, gstrSysName
            On Error Resume Next
            Me.Show
            Me.txtCode.SetFocus
            Exit Sub
        Else
            For n = 1 To .RecordCount
                If n = 1 Then
                    strResult = Val(!药名ID)
                Else
                    strResult = strResult & "," & Val(!药名ID)
                End If
                .MoveNext
            Next
        End If
    End With
    
    Me.Hide
    Call frmMediLists.zlGetFilter(strKind, strResult)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    If mblnSelfMedi = True Then
        chk中草药.Visible = False
    Else
        chk中草药.Visible = True
    End If
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub txtAlia_Change()
    
End Sub

Private Sub txtAlia_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub

Private Sub txtAlia_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtcode_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub


Private Sub txtcode_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub





Private Sub txtName_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtSim_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub

Private Sub txtSim_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


