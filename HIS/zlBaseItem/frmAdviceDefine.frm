VERSION 5.00
Begin VB.Form frmAdviceDefine 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医嘱内容定义"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "frmAdviceDefine.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdAdd 
      Height          =   270
      Left            =   5175
      Picture         =   "frmAdviceDefine.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "加入字段(ALT+A)"
      Top             =   2325
      Width           =   270
   End
   Begin VB.ComboBox cbo字段 
      Height          =   300
      Left            =   1020
      TabIndex        =   5
      Top             =   2310
      Width           =   4125
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "检查(&K)"
      Height          =   350
      Left            =   2070
      TabIndex        =   7
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4365
      TabIndex        =   9
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3270
      TabIndex        =   8
      Top             =   2865
      Width           =   1100
   End
   Begin VB.TextBox txtAdvice 
      Height          =   1125
      Left            =   1020
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1155
      Width           =   4440
   End
   Begin VB.ComboBox cbo类别 
      Height          =   300
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   825
      Width           =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   -75
      X2              =   5985
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   -165
      X2              =   5895
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   6060
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   -90
      X2              =   5970
      Y1              =   2745
      Y2              =   2745
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAdviceDefine.frx":00D6
      Height          =   645
      Left            =   345
      TabIndex        =   10
      Top             =   75
      Width           =   5040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "诊疗类别"
      Height          =   180
      Left            =   225
      TabIndex        =   0
      Top             =   885
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "医嘱内容"
      Height          =   180
      Left            =   225
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "字段项目"
      Height          =   180
      Left            =   225
      TabIndex        =   4
      Top             =   2370
      Width           =   1455
   End
End
Attribute VB_Name = "frmAdviceDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOk As Boolean
Private mblnChange As Boolean
Private mintIndex As Integer
Private mrsField As ADODB.Recordset
Private mrsAdvice As ADODB.Recordset

Public Function ShowMe(frmParent As Object, rsAdvice As ADODB.Recordset) As Boolean
    Set mrsAdvice = Rec.CopyNew(rsAdvice)
    Me.Show 1, frmParent
    If mblnOk Then
        Set rsAdvice = Rec.CopyNew(mrsAdvice)
    End If
    ShowMe = mblnOk
End Function

Private Sub cbo字段_GotFocus()
    Call zlControl.TxtSelAll(cbo字段)
End Sub

Private Sub cbo字段_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
    If cbo字段.Text = "" Then Exit Sub
    txtAdvice.SelText = cbo字段.Text
    cbo字段.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    Dim strMsg As String
    
    If Trim(txtAdvice.Text) = "" Then
        MsgBox "没有内容。", vbInformation, gstrSysName
    Else
        strMsg = CheckAdvice(txtAdvice.Text)
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
        Else
            MsgBox "医嘱内容书写正确。", vbInformation, gstrSysName
        End If
    End If
    txtAdvice.SetFocus
End Sub

Private Function CheckAdvice(ByVal strText As String) As String
'功能：检查医嘱内容是否正确
'返回：错误信息
'      strPreview=预览医嘱内容效果
    Dim intLeft As Integer, intRight As Integer
    Dim strTmp As String, strPar As String
    Dim strMsg As String, i As Long
    Dim objVBA As Object, strEval As String
    Dim objScript As New clsScript
    
    If Trim(strText) = "" Then Exit Function
    If zlCommFun.ActualLen(strText) > txtAdvice.MaxLength Then
        strMsg = "医嘱定义内容太长，只允许 " & txtAdvice.MaxLength & " 个字符或 " & txtAdvice.MaxLength \ 2 & " 个汉字。"
        GoTo EndLine
    End If
        
    '检查配对情况
    For i = 1 To Len(strText)
        If Mid(strText, i, 1) = "[" Then
            intLeft = intLeft + 1
        ElseIf Mid(strText, i, 1) = "]" Then
            intRight = intRight + 1
            If intLeft <> intRight Then
                strMsg = """[""与""]""括号不配对。"
                GoTo EndLine
            End If
        End If
    Next
    If intLeft = 0 And intRight = 0 Then Exit Function
    If intLeft <> intRight Then
        strMsg = """[""与""]""括号不配对。"
        GoTo EndLine
    End If
    
    '检查字段名称
    strTmp = strText
    Do While InStr(strTmp, "[") > 0
        strTmp = Mid(strTmp, InStr(strTmp, "[") + 1)
        strPar = Trim(Left(strTmp, InStr(strTmp, "]") - 1))
                        
        If strPar = "" Then
            strMsg = """[]""括号之中没有书写字段名。"
            GoTo EndLine
        End If
        
        For i = 0 To cbo字段.ListCount - 1
            If cbo字段.List(i) = "[" & strPar & "]" Then Exit For
        Next
        If i > cbo字段.ListCount - 1 Then
            strMsg = "使用了不存在的""[" & strPar & "]""字段。"
            GoTo EndLine
        End If
    Loop
    
    '执行测试
    On Error Resume Next
    Set objVBA = CreateObject("ScriptControl")
    If objVBA Is Nothing Then
        strMsg = "Microsoft Script Control 未正确安装(msscript.ocx)，不能执行检查。请重新安装客户端程序。"
        GoTo EndLine
    End If
    Err.Clear: On Error GoTo 0
    objVBA.Language = "VBScript"
    objVBA.AddObject "clsScript", objScript, True
    strEval = Replace(strText, "[", """")
    strEval = Replace(strEval, "]", """")
    On Error Resume Next
    Call objVBA.Eval(strEval)
    If objVBA.Error.Number <> 0 Then
        strMsg = objVBA.Error.Description
        objVBA.Error.Clear
    End If
EndLine:
    CheckAdvice = strMsg
End Function

Private Sub cmdOK_Click()
    If Not UpdateAdvice Then
        txtAdvice.SetFocus: Exit Sub
    End If
    mrsAdvice.Filter = 0
    mblnChange = False
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbAltMask Then
        Call cmdAdd_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    
    mblnOk = False
    
    '初始化不同类别可用的字段内容
    On Error GoTo ErrHandle
    Set mrsField = New ADODB.Recordset
    mrsField.Fields.Append "类别", adVarChar, 4
    mrsField.Fields.Append "字段", adVarChar, 2000
    mrsField.CursorLocation = adUseClient
    mrsField.LockType = adLockBatchOptimistic
    mrsField.CursorType = adOpenStatic
    mrsField.Open
'    Set mrsField.ActionConnection = Nothing
    mrsField.AddNew: mrsField!类别 = "公共": mrsField!字段 = "[开始时间],[医生嘱托]" '公共的字段项目
    mrsField.AddNew: mrsField!类别 = "其他": mrsField!字段 = "[诊疗项目],[单量],[总量],[中文频率],[英文频率],[执行时间]"
    mrsField.AddNew: mrsField!类别 = "4": mrsField!字段 = "[卫生材料],[规格],[产地]"
    mrsField.AddNew: mrsField!类别 = "5": mrsField!字段 = "[输入名],[通用名],[商品名],[英文名],[规格],[产地],[单量],[总量],[中文频率],[英文频率],[执行时间],[给药途径]"
    mrsField.AddNew: mrsField!类别 = "6": mrsField!字段 = "[输入名],[通用名],[商品名],[英文名],[规格],[产地],[单量],[总量],[中文频率],[英文频率],[执行时间],[给药途径]"
    mrsField.AddNew: mrsField!类别 = "8": mrsField!字段 = "[付数],[配方组成],[中文频率],[英文频率],[执行时间],[用法],[煎法]"
    mrsField.AddNew: mrsField!类别 = "C": mrsField!字段 = "[检验项目],[检验标本],[采集方法]"
    mrsField.AddNew: mrsField!类别 = "D": mrsField!字段 = "[检查项目],[检查部位]"
    mrsField.AddNew: mrsField!类别 = "F": mrsField!字段 = "[手术时间],[主要手术],[附加手术],[麻醉方法]"
    mrsField.AddNew: mrsField!类别 = "K": mrsField!字段 = "[输血时间],[输血项目],[输血途径],[血型],[RH],[执行分类]"
    mrsField.UpdateBatch
    
    '可用诊疗类别，不包含:
    '7-中草药:不能单独下医嘱
    '9-成套:非单个诊疗项目
    'G-麻醉:不能单独下医嘱
    gstrSQL = "Select 编码,名称 From 诊疗项目类别 Where 编码 Not IN('7','9','G') Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Do While Not rsTmp.EOF
        cbo类别.AddItem rsTmp!编码 & "-" & rsTmp!名称
        rsTmp.MoveNext
    Loop
    cbo类别.ListIndex = 0
    
    mintIndex = cbo类别.ListIndex
    mblnChange = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function UpdateAdvice() As Boolean
    Dim strMsg As String
    
    strMsg = CheckAdvice(txtAdvice.Text)
    If strMsg <> "" Then
        Call zlControl.CboSetIndex(cbo类别.hwnd, mintIndex)
        MsgBox strMsg, vbInformation, gstrSysName
        txtAdvice.SetFocus: Exit Function
    End If
    mrsAdvice.Filter = "诊疗类别='" & Left(cbo类别.List(mintIndex), 1) & "'"
    If mrsAdvice.EOF Then
        If Trim(txtAdvice.Text) <> "" Then '原本没内容的情况下
            mrsAdvice.AddNew
            mrsAdvice!诊疗类别 = Left(cbo类别.List(mintIndex), 1)
            mrsAdvice!医嘱内容 = txtAdvice.Text
            mrsAdvice.Update
            mblnChange = True
        End If
    Else
        If Trim(txtAdvice.Text) = "" Then '原本有内容的情况下未设置
            Call zlControl.CboSetIndex(cbo类别.hwnd, mintIndex)
            MsgBox "当前类别的医嘱内容没有设置。", vbInformation, gstrSysName
            txtAdvice.SetFocus: Exit Function
        ElseIf mrsAdvice!医嘱内容 <> txtAdvice.Text Then
            mrsAdvice!医嘱内容 = txtAdvice.Text
            mrsAdvice.Update
            mblnChange = True
        End If
    End If
    txtAdvice.Tag = ""
    UpdateAdvice = True
End Function

Private Sub cbo类别_Click()
    Dim arrField As Variant, i As Long
    
    '1.检查并更新当前类别的医嘱内容
    '------------------------------
    If Visible And txtAdvice.Tag = "1" Then
        If Not UpdateAdvice Then Exit Sub
    End If
    '2.显示新切换到的类别的医嘱内容
    '------------------------------
    mintIndex = cbo类别.ListIndex
    
    '显示可用字段列表
    cbo字段.Clear
    
    mrsField.Filter = "类别='公共'"
    Do While Not mrsField.EOF
        arrField = Split(mrsField!字段, ",")
        For i = 0 To UBound(arrField)
            cbo字段.AddItem arrField(i)
        Next
        mrsField.MoveNext
    Loop
    
    mrsField.Filter = "类别='" & Left(cbo类别.Text, 1) & "'"
    If mrsField.EOF Then
        mrsField.Filter = "类别='其他'"
    End If
    arrField = Split(mrsField!字段, ",")
    For i = 0 To UBound(arrField)
        cbo字段.AddItem arrField(i)
    Next
    
    lblPrompt.Caption = "请选择下面的可用字段项，使用与VBScript兼容的表达式对医嘱内容进行组合；字段项请使用方括符""[]""括起表示，所有字段项取值皆为字符串。"
    '显示当前设置的医嘱内容
    mrsAdvice.Filter = "诊疗类别='" & Left(cbo类别.Text, 1) & "'"
    If Not mrsAdvice.EOF Then
        txtAdvice.Text = mrsAdvice!医嘱内容
        If mrsAdvice!诊疗类别 = "D" Then
            lblPrompt.Caption = lblPrompt.Caption & "对于病理类别，[检查部位]指""标本+材料""。"
        End If
    Else
        txtAdvice.Text = ""
    End If
    txtAdvice.Tag = ""
           
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Or txtAdvice.Tag = "1" Then
        If MsgBox("如果退出将会丢失你所改变的内容，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
    End If
End Sub

Private Sub txtAdvice_Change()
    txtAdvice.Tag = "1"
End Sub
