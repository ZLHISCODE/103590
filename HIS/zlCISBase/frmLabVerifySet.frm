VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabVerifySet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "规则"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10545
   Icon            =   "frmLabVerifySet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.TreeView tvwItem 
      Height          =   4995
      Left            =   75
      TabIndex        =   5
      Top             =   90
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   8811
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
      Left            =   9165
      TabIndex        =   3
      Top             =   3765
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7665
      TabIndex        =   2
      Top             =   3765
      Width           =   1100
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "验证(&C)"
      Height          =   350
      Left            =   3555
      TabIndex        =   1
      Top             =   3765
      Width           =   1100
   End
   Begin VB.TextBox txtFormula 
      Height          =   3600
      Left            =   3210
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   7200
   End
   Begin VB.Label lbl说明 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLabVerifySet.frx":000C
      ForeColor       =   &H00008000&
      Height          =   555
      Left            =   3255
      TabIndex        =   6
      Top             =   4260
      Width           =   7170
   End
   Begin VB.Label lblFormula 
      Caption         =   "例如：[白细胞]>2 AND [红细胞]<10) OR ([红细胞平均体积] >4 AND [红细胞压积] <20) "
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   3255
      TabIndex        =   4
      Top             =   4860
      Width           =   7170
   End
End
Attribute VB_Name = "frmLabVerifySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFormula As String '传入的规则
Private mlngID As Long        '检验项目ID
Private mstrItem As String    '项目，用于检查
Private mlng仪器ID As Long

'----------------------------------------------------
'-- 以下是本窗体控件过程
'----------------------------------------------------
Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim itemTmp As ListItem
    Dim rsGroup As ADODB.Recordset
    Dim strItem As String
    
    txtFormula = mstrFormula
    tvwItem.Nodes.Clear
    
    On Error GoTo errHand
    If mlngID = 0 And mlng仪器ID = 0 Then
        strSQL = "Select 编码||'-'||名称 as 显示名称 ,名称 ,编码 From 诊疗检验类型 where 名称 IN (" & vbNewLine & _
                        "Select D.操作类型 From 检验项目 A, 诊治所见项目 B, 诊疗项目目录 D, 检验报告项目 C" & vbNewLine & _
                        "Where A.诊治项目id = B.ID And B.ID = C.报告项目id And C.诊疗项目id = D.ID And D.类别 = 'C'  And" & vbNewLine & _
                        "      Nvl(D.组合项目, 0) = 0 )"
    Else
'        strSQL = "Select 编码 || '-' || 名称 As 显示名称, 名称, 编码" & vbNewLine & _
'                "From 诊疗检验类型" & vbNewLine & _
'                "Where 名称 In (" & vbNewLine & _
'                "" & vbNewLine & _
'                "             Select A.操作类型" & vbNewLine & _
'                "             From 诊疗项目组合 B, 诊疗项目目录 A" & vbNewLine & _
'                "             Where A.ID = B.诊疗项目id And A.类别 = 'C' And Nvl(A.组合项目, 0) = 0 And Nvl(A.单独应用, 0) = 1 And B.诊疗组合id = [1]" & vbNewLine & _
'                "             Union" & vbNewLine & _
'                "" & vbNewLine & _
'                "             Select A.操作类型" & vbNewLine & _
'                "             From 诊疗项目目录 A" & vbNewLine & _
'                "             Where Nvl(单独应用, 0) = 1 And Nvl(A.组合项目, 0) = 0 And A.类别 = 'C' And A.ID = [1]" & vbNewLine & _
'                "             Union" & vbNewLine & _
'                "" & vbNewLine & _
'                "             Select D.操作类型" & vbNewLine & _
'                "             From 诊疗项目目录 D, 检验报告项目 C, 诊治所见项目 B, 检验仪器项目 A" & vbNewLine & _
'                "             Where C.诊疗项目id = D.ID And B.ID = C.报告项目id And A.项目id = B.ID And D.类别 = 'C' And Nvl(D.组合项目, 0) = 0 And" & vbNewLine & _
'                "                   Nvl(D.单独应用, 0) = 1 And A.仪器id = [2])
        strSQL = "Select 编码 || '-' || 名称 As 显示名称, 名称, 编码" & vbNewLine & _
                "From 诊疗检验类型" & vbNewLine & _
                "Where 名称 In (Select 操作类型" & vbNewLine & _
                "             From 诊疗项目目录" & vbNewLine & _
                "             Where ID = [1]" & vbNewLine & _
                "             Union" & vbNewLine & _
                "             Select D.操作类型" & vbNewLine & _
                "             From 诊疗项目目录 D, 检验报告项目 C, 诊治所见项目 B, 检验仪器项目 A" & vbNewLine & _
                "             Where C.诊疗项目id = D.ID And B.ID = C.报告项目id And A.项目id = B.ID And D.类别 = 'C' And Nvl(D.组合项目, 0) = 0 And" & vbNewLine & _
                "                    A.仪器id = [2])"


    End If
    Set rsGroup = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID, mlng仪器ID)
    mstrItem = ","
    Do Until rsGroup.EOF
        tvwItem.Nodes.Add , , "" & rsGroup.Fields("名称"), "" & rsGroup.Fields("显示名称")
        If mlngID = 0 And mlng仪器ID = 0 Then
            
            strSQL = "Select Distinct A.诊治项目id, A.缩写, B.中文名,d.编码 " & vbNewLine & _
                    "From 检验项目 A, 诊治所见项目 B, 诊疗项目目录 D, 检验报告项目 C" & vbNewLine & _
                    "Where A.诊治项目id = B.ID And B.ID = C.报告项目id And C.诊疗项目id = D.ID And D.类别 = 'C'  And" & vbNewLine & _
                    "      Nvl(D.组合项目, 0) = 0 And D.操作类型 = [1]"

        Else
'            strSQL = "Select E.诊治项目id, E.缩写, D.中文名" & vbNewLine & _
'                    "From 检验项目 E, 诊治所见项目 D, 检验报告项目 C, 诊疗项目组合 B, 诊疗项目目录 A" & vbNewLine & _
'                    "Where A.ID = C.诊疗项目id And C.报告项目id = D.ID And D.ID = E.诊治项目id And A.ID = B.诊疗项目id And A.类别 = 'C' And Nvl(A.组合项目, 0) = 0 And" & vbNewLine & _
'                    "      Nvl(A.单独应用, 0) = 1 And B.诊疗组合id = [2] And A.操作类型 = [1]" & vbNewLine & _
'                    "Union" & vbNewLine & _
'                    "" & vbNewLine & _
'                    "Select E.诊治项目id, E.缩写, B.中文名" & vbNewLine & _
'                    "From 检验项目 E, 检验报告项目 C, 诊治所见项目 B, 诊疗项目目录 A" & vbNewLine & _
'                    "Where E.诊治项目id = B.ID And A.ID = C.诊疗项目id And C.报告项目id = E.诊治项目id And Nvl(A.单独应用, 0) = 1 And Nvl(A.组合项目, 0) = 0 And" & vbNewLine & _
'                    "      A.类别 = 'C' And A.ID = [2] And A.操作类型 = [1]" & vbNewLine & _
'                    "Union" & vbNewLine & _
'                    "" & vbNewLine & _
'                    "Select E.诊治项目id, E.缩写, B.中文名" & vbNewLine & _
'                    "From 检验项目 E, 诊疗项目目录 D, 检验报告项目 C, 诊治所见项目 B, 检验仪器项目 A" & vbNewLine & _
'                    "Where E.诊治项目id = C.报告项目id And C.诊疗项目id = D.ID And B.ID = C.报告项目id And A.项目id = B.ID And D.类别 = 'C' And" & vbNewLine & _
'                    "      Nvl(D.组合项目, 0) = 0 And Nvl(D.单独应用, 0) = 1 And A.仪器id = [3] And D.操作类型 = [1]"
            strSQL = "Select E.诊治项目id, E.缩写, D.中文名,d.编码 " & vbNewLine & _
                    "From 检验项目 E, 诊治所见项目 D, 检验报告项目 C" & vbNewLine & _
                    "Where C.报告项目id = D.ID And D.ID = E.诊治项目id And C.诊疗项目id = [2]" & vbNewLine & _
                    "Union" & vbNewLine & _
                    "" & vbNewLine & _
                    "Select E.诊治项目id, E.缩写, B.中文名,d.编码 " & vbNewLine & _
                    "From 检验项目 E, 诊疗项目目录 D, 检验报告项目 C, 诊治所见项目 B, 检验仪器项目 A" & vbNewLine & _
                    "Where E.诊治项目id = C.报告项目id And C.诊疗项目id = D.ID And B.ID = C.报告项目id And A.项目id = B.ID And D.类别 = 'C' And" & vbNewLine & _
                    "      Nvl(D.组合项目, 0) = 0  And A.仪器id = [3] And D.操作类型 = [1]"

        End If
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, "" & rsGroup.Fields("名称"), mlngID, mlng仪器ID)
        Do Until rsTmp.EOF
            mstrItem = mstrItem & "[" & IIf("" & rsTmp.Fields("缩写") = "", rsTmp.Fields("诊治项目ID"), rsTmp.Fields("编码") & "_" & rsTmp.Fields("缩写")) & "]" & ","
            tvwItem.Nodes.Add "" & rsGroup.Fields("名称"), tvwChild, "K" & rsGroup.Fields("编码") & "_" & rsTmp.Fields("诊治项目ID"), _
            "[" & IIf("" & rsTmp.Fields("缩写") = "", rsTmp.Fields("诊治项目ID"), rsTmp.Fields("编码") & "_" & rsTmp.Fields("缩写")) & "]" & rsTmp.Fields("中文名")
            rsTmp.MoveNext
        Loop
        rsGroup.MoveNext
    Loop
    cmdOk.Enabled = False
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tvwItem_DblClick()
    If InStr(tvwItem.SelectedItem.Text, "]") > 0 Then
        txtFormula.SelText = Mid(tvwItem.SelectedItem.Text, 1, InStr(tvwItem.SelectedItem.Text, "]"))
    End If
End Sub

Private Sub txtFormula_Change()
    If Trim(txtFormula.Text) <> "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub cmdCheck_Click()
    If Trim(txtFormula.Text) = "" Then cmdOk.Enabled = True: Exit Sub
    If CheckRule(txtFormula, mstrItem) Then
        cmdOk.Enabled = True
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtFormula.Text) = "" Then mstrFormula = "": Unload Me: Exit Sub
    If CheckRule(txtFormula, mstrItem) Then
        mstrFormula = txtFormula
        Unload Me
    Else
         MsgBox "规则设置错误！", vbExclamation, Me.Caption
    End If
End Sub

'-----------------------------------------------------------------
'-- 以下是 自定义过程
'-----------------------------------------------------------------

Public Function DefFormula(ByVal lngID As Long, ByVal lng仪器ID As Long, ByVal strFormula As String, ByVal frmMain As Form) As String
    '功能：调用入口
    'lngID :当前操作的检验项目 ID
    'strFormula :原来的公式
    'frmMain: 调用窗体
    mlngID = lngID: mlng仪器ID = lng仪器ID
    mstrFormula = strFormula
    
    Me.Show vbModal, frmMain
    DefFormula = mstrFormula
End Function


