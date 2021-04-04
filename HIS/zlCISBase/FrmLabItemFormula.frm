VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLabItemFormula 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "公式"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9195
   Icon            =   "FrmLabItemFormula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9195
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.TreeView tvwItem 
      Height          =   4290
      Left            =   105
      TabIndex        =   6
      Top             =   360
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   7567
      _Version        =   393217
      Indentation     =   459
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "取消(&E)"
      Height          =   350
      Left            =   7815
      TabIndex        =   3
      Top             =   4275
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6315
      TabIndex        =   2
      Top             =   4275
      Width           =   1100
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "验证(&C)"
      Height          =   350
      Left            =   3555
      TabIndex        =   1
      Top             =   4275
      Width           =   1100
   End
   Begin VB.TextBox txtFormula 
      Height          =   3825
      Left            =   3285
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   5745
   End
   Begin VB.Label lblFormula 
      Caption         =   "例如：([SD]+[CV])/100"
      Height          =   210
      Left            =   3360
      TabIndex        =   5
      Top             =   105
      Width           =   4080
   End
   Begin VB.Label lblItem 
      Caption         =   "项目"
      Height          =   210
      Left            =   180
      TabIndex        =   4
      Top             =   105
      Width           =   585
   End
End
Attribute VB_Name = "FrmLabItemFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFormula As String '传入的公式
Private mlngID As Long '检验项目ID
Private mstrItem As String '项目，用于检查

Private Sub cmdCheck_Click()
    If CheckFormula(txtFormula) Then
        cmdOk.Enabled = True
    Else
        MsgBox "公式错误！", vbExclamation, Me.Caption
    End If
    
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If CheckFormula(txtFormula) Then
        mstrFormula = txtFormula
        Unload Me
    Else
         MsgBox "公式错误！", vbExclamation, Me.Caption
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim itemTmp As ListItem
    Dim rsGroup As ADODB.Recordset
    
    On Error GoTo ErrHandle
    txtFormula = mstrFormula
    tvwItem.Nodes.Clear
    strSQL = "Select 编码||'-'||名称 as 显示名称 ,名称 ,编码 From 诊疗检验类型"
    Set rsGroup = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    mstrItem = ","
    Do Until rsGroup.EOF
        tvwItem.Nodes.Add , , "" & rsGroup.Fields("名称"), "" & rsGroup.Fields("显示名称")
        strSQL = "Select distinct A.诊治项目id, A.缩写, B.中文名 " & vbNewLine & _
                "From 诊疗项目目录 D, 检验报告项目 C, 诊治所见项目 B, 检验项目 A" & vbNewLine & _
                "Where C.诊疗项目id = D.ID And C.报告项目id = A.诊治项目id And A.诊治项目id = B.ID And D.操作类型 = [1] And A.结果类型 = 1"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, "" & rsGroup.Fields("名称"))
        Do Until rsTmp.EOF
            '可以引用自身，被引用的项目可以是计算项目。
            mstrItem = mstrItem & "[" & IIf("" & rsTmp.Fields("缩写") = "", rsTmp.Fields("诊治项目ID"), rsTmp.Fields("缩写")) & "]" & ","
            tvwItem.Nodes.Add "" & rsGroup.Fields("名称"), tvwChild, "K" & rsGroup.Fields("编码") & "_" & rsTmp.Fields("诊治项目ID"), "[" & IIf("" & rsTmp.Fields("缩写") = "", rsTmp.Fields("诊治项目ID"), rsTmp.Fields("缩写")) & "]" & rsTmp.Fields("中文名")
            rsTmp.MoveNext
        Loop
        rsGroup.MoveNext
    Loop
    cmdOk.Enabled = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function DefFormula(ByVal lngID As Long, ByVal strFormula As String, ByVal frmMain As Form) As String
    'lngID :当前操作的检验项目 ID
    'strFormula :原来的公式
    'frmMain: 调用窗体
    mlngID = lngID
    mstrFormula = strFormula
    
    Me.Show vbModal, frmMain
    DefFormula = mstrFormula
End Function

Private Function CheckFormula(ByVal strFormula As String) As Boolean
    '
    Dim strTmp As String, strLine As String, i As Integer
    Dim dblValues As Double, strItem As String, lngLength As Long
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHandle
    strLine = strFormula
    strTmp = ""
    Do While strLine Like "*[[]*[]]*"
        strTmp = strTmp & Mid(strLine, 1, InStr(strLine, "[") - 1) & "(" & i & "+ 1)"
        lngLength = InStr(strLine, "]") - InStr(strLine, "[")
        strItem = Mid(strLine, InStr(strLine, "["), lngLength + 1)
        If InStr(mstrItem, "," & strItem & ",") <= 0 Then
            Exit Function
        End If
        strLine = Mid(strLine, InStr(strLine, "]") + 1)
        i = i + 1
    Loop
    strTmp = strTmp & strLine

    Set rsTmp = zldatabase.OpenSQLRecord("Select " & strTmp & " as 计算结果 From Dual", Me.Caption)
    If Not rsTmp.EOF Then
        dblValues = rsTmp.Fields("计算结果")
        CheckFormula = True
    End If
    
    Exit Function
ErrHandle:
    CheckFormula = False
End Function

Private Sub tvwItem_DblClick()
    If InStr(tvwItem.SelectedItem.Text, "]") > 0 Then
        txtFormula.SelText = Mid(tvwItem.SelectedItem.Text, 1, InStr(tvwItem.SelectedItem.Text, "]"))
    End If
End Sub

Private Sub txtFormula_Change()
    cmdOk.Enabled = False
End Sub
