VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabVerifyEspecial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "特殊规则"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   10350
   Icon            =   "frmLabVerifyEspecial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraRule5 
      Caption         =   "与上次结果标记比较"
      Height          =   3750
      Left            =   3555
      TabIndex        =   37
      Top             =   2730
      Width           =   6630
      Begin VB.TextBox txtLastTag 
         Height          =   2300
         Left            =   3015
         TabIndex        =   39
         Top             =   510
         Width           =   3300
      End
      Begin MSComctlLib.TreeView tvwLastTag 
         Height          =   2295
         Left            =   135
         TabIndex        =   38
         Top             =   510
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   4048
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   459
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmLabVerifyEspecial.frx":000C
         ForeColor       =   &H0000C000&
         Height          =   720
         Left            =   225
         TabIndex        =   42
         Top             =   2850
         Width           =   5400
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "公式编辑"
         Height          =   210
         Left            =   3015
         TabIndex        =   41
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目列表"
         Height          =   195
         Left            =   150
         TabIndex        =   40
         Top             =   300
         Width           =   885
      End
   End
   Begin MSComctlLib.TreeView tvwItem 
      Height          =   6795
      Left            =   165
      TabIndex        =   0
      Top             =   195
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   11986
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   617
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "验证(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   21
      Top             =   6615
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7155
      TabIndex        =   20
      Top             =   6600
      Width           =   1100
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "取消(&E)"
      Height          =   350
      Left            =   8655
      TabIndex        =   19
      Top             =   6600
      Width           =   1100
   End
   Begin VB.TextBox txt公式 
      Height          =   2000
      Left            =   3555
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   195
      Width           =   6645
   End
   Begin VB.OptionButton optOr 
      Caption         =   "OR"
      Height          =   270
      Left            =   8220
      TabIndex        =   3
      Top             =   2250
      Value           =   -1  'True
      Width           =   675
   End
   Begin VB.OptionButton optAnd 
      Caption         =   "AND"
      Height          =   270
      Left            =   7500
      TabIndex        =   2
      Top             =   2250
      Width           =   705
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "加入(&A)"
      Height          =   350
      Left            =   8955
      TabIndex        =   1
      Top             =   2220
      Width           =   1100
   End
   Begin VB.Frame FraRule2 
      Caption         =   "与上次结果比较"
      Height          =   3750
      Left            =   3570
      TabIndex        =   4
      Top             =   2730
      Width           =   6630
      Begin VB.TextBox txtLast 
         Height          =   2300
         Left            =   3120
         TabIndex        =   5
         Top             =   495
         Width           =   3300
      End
      Begin MSComctlLib.TreeView tvwLast 
         Height          =   2295
         Left            =   210
         TabIndex        =   25
         Top             =   495
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   4048
         _Version        =   393217
         Indentation     =   459
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label9 
         Caption         =   "项目列表"
         Height          =   195
         Left            =   225
         TabIndex        =   27
         Top             =   285
         Width           =   885
      End
      Begin VB.Label Label8 
         Caption         =   "公式编辑"
         Height          =   210
         Left            =   3150
         TabIndex        =   26
         Top             =   285
         Width           =   885
      End
      Begin VB.Label Label7 
         Caption         =   $"frmLabVerifyEspecial.frx":00EF
         ForeColor       =   &H0000C000&
         Height          =   765
         Left            =   420
         TabIndex        =   6
         Top             =   2850
         Width           =   5970
      End
   End
   Begin VB.Frame FraRule1 
      Caption         =   "结果为X的超过N个"
      Height          =   3750
      Left            =   3540
      TabIndex        =   11
      Top             =   2730
      Width           =   6630
      Begin VB.ComboBox cbo项目个数 
         Height          =   300
         ItemData        =   "frmLabVerifyEspecial.frx":01CB
         Left            =   3345
         List            =   "frmLabVerifyEspecial.frx":01E1
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   585
         Width           =   810
      End
      Begin VB.ComboBox cbo检验结果 
         Height          =   300
         ItemData        =   "frmLabVerifyEspecial.frx":0213
         Left            =   195
         List            =   "frmLabVerifyEspecial.frx":022C
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3300
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txt项目个数 
         Height          =   285
         Left            =   4140
         TabIndex        =   13
         Top             =   585
         Width           =   1080
      End
      Begin VB.TextBox txt检验结果 
         Height          =   285
         Left            =   1410
         TabIndex        =   12
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label12 
         Caption         =   "的时候提示。"
         Height          =   210
         Left            =   5265
         TabIndex        =   24
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "的项目个数"
         Height          =   210
         Left            =   2400
         TabIndex        =   16
         Top             =   630
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "检验结果等于"
         Height          =   210
         Left            =   270
         TabIndex        =   15
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmLabVerifyEspecial.frx":0264
         ForeColor       =   &H0000C000&
         Height          =   1065
         Left            =   1485
         TabIndex        =   14
         Top             =   1770
         Width           =   5025
      End
   End
   Begin VB.Frame fraRule4 
      Caption         =   "漏项、多项检查"
      Height          =   3750
      Left            =   3555
      TabIndex        =   33
      Top             =   2715
      Width           =   6630
      Begin VB.ComboBox cbo检查方式 
         Height          =   300
         ItemData        =   "frmLabVerifyEspecial.frx":02CE
         Left            =   1545
         List            =   "frmLabVerifyEspecial.frx":02DB
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   480
         Width           =   1725
      End
      Begin VB.Label Label13 
         Caption         =   "检查方式"
         Height          =   225
         Left            =   660
         TabIndex        =   36
         Top             =   525
         Width           =   810
      End
      Begin VB.Label Label16 
         Caption         =   "例如：仪器传回的结果缺少RBC时，则提示。"
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   330
         TabIndex        =   35
         Top             =   2235
         Width           =   5040
      End
   End
   Begin VB.Frame FraRule3 
      Caption         =   "除几个项目外，结果为X"
      Height          =   3750
      Left            =   3555
      TabIndex        =   7
      Top             =   2730
      Width           =   6630
      Begin VB.ComboBox cboNot符号 
         Height          =   300
         ItemData        =   "frmLabVerifyEspecial.frx":0301
         Left            =   2865
         List            =   "frmLabVerifyEspecial.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2625
         Width           =   810
      End
      Begin VB.TextBox txtNot值 
         Height          =   285
         Left            =   3735
         TabIndex        =   30
         Top             =   2625
         Width           =   1080
      End
      Begin VB.TextBox txtNot项目 
         Height          =   1770
         Left            =   2865
         TabIndex        =   8
         Top             =   495
         Width           =   3570
      End
      Begin MSComctlLib.TreeView tvwNot 
         Height          =   2445
         Left            =   210
         TabIndex        =   28
         Top             =   495
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   4313
         _Version        =   393217
         Indentation     =   459
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label5 
         Caption         =   "除以上项目外，其他项目的结果中，如果有"
         Height          =   210
         Left            =   2865
         TabIndex        =   31
         Top             =   2370
         Width           =   3540
      End
      Begin VB.Label Label11 
         Caption         =   "项目列表"
         Height          =   195
         Left            =   225
         TabIndex        =   29
         Top             =   255
         Width           =   885
      End
      Begin VB.Label Label6 
         Caption         =   "的项目，则提示。"
         Height          =   180
         Left            =   4905
         TabIndex        =   10
         Top             =   2685
         Width           =   1470
      End
      Begin VB.Label Label4 
         Caption         =   "例如：除BEact,Beecf外,其他项目的结果有负数的，则提示。"
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   345
         TabIndex        =   9
         Top             =   3150
         Width           =   5040
      End
   End
   Begin VB.Label Label10 
      Caption         =   "新加入规则与原规则的关系"
      Height          =   195
      Left            =   5220
      TabIndex        =   18
      Top             =   2280
      Width           =   2205
   End
End
Attribute VB_Name = "frmLabVerifyEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngID As Long
Private mlng仪器ID As Long
Private mstrFormula As String
Private mStrItem As String

Private Sub cmdAdd_Click()
    Dim strAndOr As String
    
    If InStr(txt公式.Text, "{") > 0 Then strAndOr = IIf(optAnd, " AND ", " OR ")
    If FraRule1.Visible Then
        txt公式.Text = txt公式.Text & strAndOr & GenFormula("A", mStrItem, txt检验结果, Gen符号(cbo项目个数), txt项目个数)
    ElseIf FraRule2.Visible Then
        txt公式.Text = txt公式.Text & strAndOr & GenFormula("B", mStrItem, txtLast)
    ElseIf FraRule3.Visible Then
        txt公式.Text = txt公式.Text & strAndOr & GenFormula("C", Replace(mStrItem, "上次.", ""), txtNot项目, Gen符号(cboNot符号), txtNot值)
    ElseIf fraRule4.Visible Then
        txt公式.Text = txt公式.Text & strAndOr & GenFormula("D", mStrItem, cbo检查方式)
    ElseIf fraRule5.Visible Then
        txt公式.Text = txt公式.Text & strAndOr & GenFormula("E", mStrItem, txtLastTag)
    End If
End Sub

Private Sub cmdCheck_Click()
    If Trim(Me.txt公式.Text) = "" Then cmdOk.Enabled = True: Exit Sub
    If CheckEspecial(txt公式, mStrItem) Then
        cmdOk.Enabled = True
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(Me.txt公式.Text) = "" Then mstrFormula = "": Unload Me: Exit Sub
    If CheckEspecial(txt公式, mStrItem) Then
        mstrFormula = txt公式
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim itemTmp As ListItem
    Dim rsGroup As ADODB.Recordset
    Dim strItem As String
    
    On Error GoTo ErrHandle
    txt公式 = mstrFormula
    tvwLast.Nodes.Clear
    tvwNot.Nodes.Clear
    tvwLastTag.Nodes.Clear
    If mlngID = 0 And mlng仪器ID = 0 Then
        strSQL = "Select 编码||'-'||名称 as 显示名称 ,名称 ,编码 From 诊疗检验类型 where 名称 IN (" & vbNewLine & _
                        "Select D.操作类型 From 检验项目 A, 诊治所见项目 B, 诊疗项目目录 D, 检验报告项目 C" & vbNewLine & _
                        "Where A.诊治项目id = B.ID And B.ID = C.报告项目id And C.诊疗项目id = D.ID And D.类别 = 'C'  And" & vbNewLine & _
                        "      Nvl(D.组合项目, 0) = 0 )"
    Else
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
    Set rsGroup = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID, mlng仪器ID)
    mStrItem = ","
    Do Until rsGroup.EOF
        tvwLast.Nodes.Add , , "" & rsGroup.Fields("名称"), "" & rsGroup.Fields("显示名称")
        tvwNot.Nodes.Add , , "" & rsGroup.Fields("名称"), "" & rsGroup.Fields("显示名称")
        tvwLastTag.Nodes.Add , , "" & rsGroup.Fields("名称"), "" & rsGroup.Fields("显示名称")
        If mlngID = 0 And mlng仪器ID = 0 Then
            strSQL = "Select Distinct A.诊治项目id, A.缩写, B.中文名,b.编码 " & vbNewLine & _
                    "From 检验项目 A, 诊治所见项目 B, 诊疗项目目录 D, 检验报告项目 C" & vbNewLine & _
                    "Where A.诊治项目id = B.ID And B.ID = C.报告项目id And C.诊疗项目id = D.ID And D.类别 = 'C'  And" & vbNewLine & _
                    "      Nvl(D.组合项目, 0) = 0 And D.操作类型 = [1]"
        Else
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "" & rsGroup.Fields("名称"), mlngID, mlng仪器ID)
        Do Until rsTmp.EOF
            mStrItem = mStrItem & "[" & IIf("" & rsTmp.Fields("缩写") = "", rsTmp.Fields("诊治项目ID"), rsTmp.Fields("编码") & "_" & rsTmp.Fields("缩写")) & "]" & ","
            mStrItem = mStrItem & "[上次." & IIf("" & rsTmp.Fields("缩写") = "", rsTmp.Fields("诊治项目ID"), rsTmp.Fields("编码") & "_" & rsTmp.Fields("缩写")) & "]" & ","
            mStrItem = mStrItem & "[标记." & IIf("" & rsTmp.Fields("缩写") = "", rsTmp.Fields("诊治项目ID"), rsTmp.Fields("编码") & "_" & rsTmp.Fields("缩写")) & "]" & ","
            
            '-- B类规则的可选项
            tvwLast.Nodes.Add "" & rsGroup.Fields("名称"), tvwChild, "K" & rsGroup.Fields("编码") & "_" & rsTmp.Fields("诊治项目ID"), "[" & _
                IIf("" & rsTmp.Fields("缩写") = "", rsTmp.Fields("诊治项目ID"), rsTmp.Fields("编码") & "_" & rsTmp.Fields("缩写")) & "]" & rsTmp.Fields("中文名")
            tvwLast.Nodes.Add "" & rsGroup.Fields("名称"), tvwChild, "KL" & rsGroup.Fields("编码") & "_" & rsTmp.Fields("诊治项目ID"), "[上次." & _
                IIf("" & rsTmp.Fields("缩写") = "", rsTmp.Fields("诊治项目ID"), rsTmp.Fields("编码") & "_" & rsTmp.Fields("缩写")) & "]" & rsTmp.Fields("中文名")
            '-- C类规则的可选项
            tvwNot.Nodes.Add "" & rsGroup.Fields("名称"), tvwChild, "K" & rsGroup.Fields("编码") & "_" & rsTmp.Fields("诊治项目ID"), "[" & _
                IIf("" & rsTmp.Fields("缩写") = "", rsTmp.Fields("诊治项目ID"), rsTmp.Fields("编码") & "_" & rsTmp.Fields("缩写")) & "]" & rsTmp.Fields("中文名")
            '-- E类规则的可选项
            tvwLastTag.Nodes.Add "" & rsGroup.Fields("名称"), tvwChild, "K" & rsGroup.Fields("编码") & "_" & rsTmp.Fields("诊治项目ID"), "[标记." & _
                 IIf("" & rsTmp.Fields("缩写") = "", rsTmp.Fields("诊治项目ID"), rsTmp.Fields("编码") & "_" & rsTmp.Fields("缩写")) & "]" & rsTmp.Fields("中文名")
            tvwLastTag.Nodes.Add "" & rsGroup.Fields("名称"), tvwChild, "KL" & rsGroup.Fields("编码") & "_" & rsTmp.Fields("诊治项目ID"), "[上次." & _
                 IIf("" & rsTmp.Fields("缩写") = "", rsTmp.Fields("诊治项目ID"), rsTmp.Fields("编码") & "_" & rsTmp.Fields("缩写")) & "]" & rsTmp.Fields("中文名")
            rsTmp.MoveNext
        Loop
        rsGroup.MoveNext
    Loop
    
    '----
    Dim nodX As Node
    tvwItem.Nodes.Clear
    tvwItem.LabelEdit = tvwManual
    Set nodX = tvwItem.Nodes.Add(, , "R1", "{A:X|N}　结果为X的有N个")
    Set nodX = tvwItem.Nodes.Add(, , "R2", "{B:P} 与上次结果比较")
    Set nodX = tvwItem.Nodes.Add(, , "R3", "{C:not N|X} 除N项外,结果为X")
    Set nodX = tvwItem.Nodes.Add(, , "R4", "{D:X} 漏项，多项检查")
    Set nodX = tvwItem.Nodes.Add(, , "R5", "{E:P} 与上次结果标志比较")
    cbo检验结果.ListIndex = 0: cbo项目个数.ListIndex = 0: cboNot符号.ListIndex = 0: cbo检查方式.ListIndex = 2
    Call tvwItem_NodeClick(tvwItem.Nodes("R1"))
    tvwItem.Nodes("R1").Selected = True
    
    cmdOk.Enabled = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tvwItem_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = "R1" Then
        FraRule1.Visible = True
        FraRule2.Visible = False
        FraRule3.Visible = False
        fraRule4.Visible = False
        fraRule5.Visible = False
    ElseIf Node.Key = "R2" Then
        FraRule1.Visible = False
        FraRule2.Visible = True
        FraRule3.Visible = False
        fraRule4.Visible = False
        fraRule5.Visible = False
    ElseIf Node.Key = "R3" Then
        FraRule1.Visible = False
        FraRule2.Visible = False
        FraRule3.Visible = True
        fraRule4.Visible = False
        fraRule5.Visible = False
    ElseIf Node.Key = "R4" Then
        FraRule1.Visible = False
        FraRule2.Visible = False
        FraRule3.Visible = False
        fraRule4.Visible = True
        fraRule5.Visible = False
    Else
        FraRule1.Visible = False
        FraRule2.Visible = False
        FraRule3.Visible = False
        fraRule4.Visible = False
        fraRule5.Visible = True
    End If
End Sub

Private Sub tvwLast_DblClick()
    If InStr(tvwLast.SelectedItem.Text, "]") > 0 Then
        txtLast.SelText = Mid(tvwLast.SelectedItem.Text, 1, InStr(tvwLast.SelectedItem.Text, "]"))
    End If
End Sub

Private Sub tvwLastTag_DblClick()
    If InStr(tvwLastTag.SelectedItem.Text, "]") > 0 Then
        txtLastTag.SelText = Mid(tvwLastTag.SelectedItem.Text, 1, InStr(tvwLastTag.SelectedItem.Text, "]"))
    End If
End Sub

Private Sub tvwNot_DblClick()
    If InStr(tvwNot.SelectedItem.Text, "]") > 0 Then
        txtNot项目 = IIf(txtNot项目 = "", "", txtNot项目 & ",") & Mid(tvwNot.SelectedItem.Text, 1, InStr(tvwNot.SelectedItem.Text, "]"))
    End If
End Sub

'-----------------------------
'-- 以下是自定义过程
'------------------------------
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


Private Function Gen符号(ByVal strIn As String) As String
    Select Case strIn
    Case "等于": Gen符号 = "="
    Case "大于": Gen符号 = ">"
    Case "小于": Gen符号 = "<"
    Case "大于等于": Gen符号 = ">="
    Case "小于等于": Gen符号 = "<="
    Case "不等于": Gen符号 = "<>"
    Case "包含": Gen符号 = " Like "
    End Select
End Function





Private Sub txt公式_Change()
    If Trim(txt公式.Text) = "" Then
        Me.cmdOk.Enabled = True
    End If
End Sub
