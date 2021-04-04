VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFormula 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "表列数据设置"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "frmFormula.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkAutoRowHeight 
      Caption         =   "该列的高度随内容自动调整"
      Height          =   195
      Left            =   840
      TabIndex        =   8
      Top             =   3900
      Width           =   2500
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "应用至所有列"
      Height          =   195
      Left            =   3540
      TabIndex        =   9
      Top             =   3900
      Width           =   1380
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "列特性设置"
      Height          =   350
      Left            =   240
      TabIndex        =   22
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CheckBox chkVisible 
      Caption         =   "该列没有任何数据时不显示本列"
      Height          =   195
      Left            =   840
      TabIndex        =   10
      Top             =   4635
      Width           =   2880
   End
   Begin VB.CommandButton cmdRelation 
      Height          =   255
      Left            =   4620
      Picture         =   "frmFormula.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2775
      Width           =   270
   End
   Begin VB.CheckBox chkAutoFont 
      Caption         =   "该列单元格的内容过多时自动缩小字体进行打印"
      Height          =   195
      Left            =   840
      TabIndex        =   7
      Top             =   3645
      Width           =   4080
   End
   Begin VB.CheckBox chk换页 
      Caption         =   "该列数据值发生变化时自动对表格进行换页处理"
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   3390
      Width           =   4080
   End
   Begin VB.CheckBox chkMerge 
      Caption         =   "相邻行数据相同时列自动合并(当前一列合并时)"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   3135
      Width           =   4080
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3000
      Left            =   480
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   5292
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      PathSeparator   =   "."
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.TextBox txtFormat 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   840
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2355
      Width           =   3795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4065
      TabIndex        =   13
      Top             =   5160
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2835
      TabIndex        =   12
      Top             =   5160
      Width           =   1100
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "校验(&V)"
      Height          =   350
      Left            =   1590
      TabIndex        =   11
      Top             =   5160
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   -105
      TabIndex        =   15
      Top             =   5010
      Width           =   5910
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   315
      Left            =   4635
      Picture         =   "frmFormula.frx":0258
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "F2"
      Top             =   1965
      Width           =   330
   End
   Begin VB.CommandButton cmdShow 
      Height          =   255
      Left            =   4350
      Picture         =   "frmFormula.frx":0432
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1995
      Width           =   270
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1965
      Width           =   3810
   End
   Begin VB.TextBox txtFormula 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   885
      Width           =   5160
   End
   Begin MSScriptControlCtl.ScriptControl srt 
      Left            =   2385
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.TextBox txtRelation 
      Height          =   300
      Left            =   840
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2760
      Width           =   3765
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "请避免跳跃式设置列，这样会引起中间未设置的列出现换行情况。 如：列1、列3、列5。"
      Height          =   435
      Left            =   1110
      TabIndex        =   23
      Top             =   4140
      Width           =   3825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "关联报表"
      Height          =   180
      Left            =   60
      TabIndex        =   21
      Top             =   2820
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "格式串"
      Height          =   180
      Left            =   240
      TabIndex        =   18
      Top             =   2415
      Width           =   540
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFormula.frx":0540
      Height          =   735
      Left            =   705
      TabIndex        =   17
      Top             =   90
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   105
      Picture         =   "frmFormula.frx":05D7
      Top             =   165
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "数据项"
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   2025
      Width           =   540
   End
End
Attribute VB_Name = "frmFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frmParent As Object
Public strInit As String
Public strFormat As String

Public mblnMerge As Boolean
Public mblnPreMerge As Boolean

Public mbln换页 As Boolean
Public mblnCan换页 As Boolean

Public mblnAutoRowHeight As Boolean
Public mblnAutoFont As Boolean
Public mblnVisible As Boolean
Public mobjRelations As RPTRelations  '出/关联报表项目
Public mobjColProtertys As RPTColProtertys '出/列特性设置

Public objReport As Report
Public intCol As Integer, intCur As Integer
Private blnDo As Boolean
Private objScript As clsScript
Private mblnBinary As Boolean

Private Sub chkAutoRowHeight_Click()
    If Me.Visible Then
        If objReport.Items("_" & intCur).自调 = False Then
            If chkAutoRowHeight.Value <> Val(chkAutoRowHeight.Tag) Then
                chkAutoRowHeight.Value = Val(chkAutoRowHeight.Tag)
                If MsgBox("报表设计主界面的表格“换行”属性未开启，它是本设置的前提条件，需要开启吗？" _
                    , vbInformation + vbDefaultButton1 + vbYesNo, Me.Caption) = vbYes Then
                    objReport.Items("_" & intCur).自调 = True
                    chkAutoRowHeight.Value = 1
                    chkAutoRowHeight.Tag = chkAutoRowHeight.Value
                End If
                Exit Sub
            End If
        End If
    End If
    If chkAutoRowHeight.Value = 1 Then
        chkAutoFont.Value = 0
    End If
End Sub

Private Sub chkAutoFont_Click()
    If chkAutoFont.Value = 1 Then
        chkAutoRowHeight.Value = 0
    End If
End Sub

Private Sub cmdAdd_Click()
    If txtItem.Text <> "" Then
        txtFormula.SelText = "[" & txtItem.Text & "]"
        txtFormula.SetFocus
    Else
        Call PlayWarn
    End If
End Sub

Private Sub cmdAdd_GotFocus()
    tvw.Visible = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancel_GotFocus()
    tvw.Visible = False
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim objRelationID As RelatID
    Dim objItem As RPTItem
    
    Call cmdVerify_Click
    If Not cmdOK.Enabled Then Exit Sub
    '如果切换了报表，没有设置参数，则清空参数对照
'    If txtRelation.Tag <> mobjRelations.Item(1).关联报表ID Then
'        For i = mobjRelations.count To 1 Step -1
'            mobjRelations.Remove i
'        Next
'    End If
    '如果对照了关联报表，却一个参数都没有对照，则增加一条参数名为NULL的表示
'    If Val(txtRelation.Tag) <> 0 Then
'        If mobjRelations.count = 0 Then
'            mobjRelations.Add Val(txtRelation.Tag), "NULL", "", txtRelation.Text
'        End If
'    End If

    '“自适应行高”应用到所有列
    If chkAll.Value = 1 Then
        For Each objRelationID In objReport.Items("_" & intCur).SubIDs
            Set objItem = objReport.Items("_" & objRelationID.id)
            If Not objItem Is Nothing Then
                If objItem.类型 = Val("6-任意表的列") Then
                    objItem.自适应行高 = chkAutoRowHeight.Value = 1
                End If
            End If
        Next
    End If
    
    gblnOK = True
    Hide
End Sub

Private Sub cmdOK_GotFocus()
    tvw.Visible = False
End Sub

Private Sub cmdRelation_Click()
    Dim strInfo As String
    Dim lngReportID As Long
    Dim strReportID As String
    Dim X As Long, Y As Long, k As Long
    
    X = InStr(1, txtFormula.Text, "]")
    Y = InStr(1, txtFormula.Text, ".")
    k = InStr(1, txtFormula.Text, "[")
    If X > k And X > Y And X <> 0 And k <> 0 Then
        strReportID = FindReport("", txtRelation.hwnd, strInfo, Val(txtRelation.Tag), objReport, mobjRelations, 1, Me)
        If strReportID <> "" Then
            strReportID = Split(Split(strReportID, "(")(1), ")")(0)
            If strReportID <> "" Then
                txtRelation.Text = strInfo
                txtRelation.Tag = strReportID
                txtRelation.Locked = True
            End If
        Else
            '判断是取消还是清除
            If mobjRelations.count > 0 Then
                txtRelation.SetFocus
            Else
                txtRelation.Text = ""
                txtRelation.Tag = ""
                txtRelation.Locked = False
            End If
        End If
    Else
        MsgBox "当前列必须先绑定一个数据源，例如：[数据源.字段],绑定后再设置关联报表。", vbInformation, Me.Caption
    End If
End Sub

Private Sub cmdSetup_Click()
    Dim objSet As New frmColSetup
    Dim X As Long, Y As Long, k As Long
            
    If txtFormula.Text Like "*[[]*.*[]]*" Then
        X = InStr(1, txtFormula.Text, "]")
        Y = InStr(1, txtFormula.Text, ".")
        k = InStr(1, txtFormula.Text, "[")
        If X > k And X > Y And X <> 0 And k <> 0 Then
            Call objSet.ShowMe(Me, mobjColProtertys, 1 _
                                , Mid(txtFormula.Text, k + 1, Y - k - 1) _
                                , Mid(txtFormula.Text, Y + 1, X - Y - 1))
        Else
            MsgBox "当前列必须先绑定一个数据源，例如：[数据源.字段]", vbInformation, Me.Caption
        End If
    Else
        MsgBox "当前列必须先绑定一个数据源，例如：[数据源.字段]", vbInformation, Me.Caption
    End If
End Sub

Private Sub cmdShow_Click()
    SetParent tvw.hwnd, 0
    tvw.Top = Top + txtItem.Top + txtItem.Height + 350
    tvw.Left = Left + txtItem.Left + 60
    tvw.ZOrder
    tvw.Visible = Not tvw.Visible
    txtItem.SetFocus
End Sub

Private Sub cmdVerify_Click()
    Dim tmpID As RelatID, tmpItem As RPTItem
    Dim strErr As String
        
    '格式化表达式
    txtFormula.Text = Replace(txtFormula.Text, " ", "")
    txtFormula.Text = Replace(UCase(txtFormula.Text), "MOD", " Mod ")
    txtFormula.Text = Replace(UCase(txtFormula.Text), "AND", " And ")
    txtFormula.Text = Replace(UCase(txtFormula.Text), "OR", " Or ")
    'txtFormula.Text = Replace(UCase(txtFormula.Text), "F OR MAT", "Format")
    txtFormula.Text = Replace(UCase(txtFormula.Text), "XOR", " Xor ")
    txtFormula.Text = Replace(UCase(txtFormula.Text), "NOT", " Not ")
    txtFormula.Text = Replace(UCase(txtFormula.Text), "LIKE", " Like ")
    
    If txtFormula.Text <> "" Then
        If LenB(StrConv(txtFormula.Text, vbFromUnicode)) > 255 Then
            MsgBox "该表格列的计算公式过长,不能超过255个字符！", vbInformation, App.Title
            cmdOK.Enabled = False
            txtFormula.SetFocus: Exit Sub
        End If
        If LenB(StrConv(txtFormat.Text, vbFromUnicode)) > 50 Then
            MsgBox "该表格列的格式串过长,不能超过50个字符！", vbInformation, App.Title
            cmdOK.Enabled = False
            txtFormat.SetFocus: Exit Sub
        End If
        
        '递归调用检查
        '自已调自已
        'If InStr(Replace(txtFormula.Text, "@", "@"), "[" & intCol & "]") > 0 And Not cmdOK.Enabled Then
        If InStr(Replace(txtFormula.Text, "@", ""), "[" & intCol & "]") > 0 And Not cmdOK.Enabled Then
            If MsgBox("计算公式中引用了该列本身，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
                cmdOK.Enabled = False
                txtFormula.SetFocus: Exit Sub
            End If
        End If
        
        '互相调用
        For Each tmpID In objReport.Items("_" & intCur).SubIDs
            Set tmpItem = objReport.Items("_" & tmpID.id)
            If InStr(Replace(tmpItem.内容, "@", ""), "[" & intCol & "]") > 0 _
                And InStr(Replace(txtFormula.Text, "@", ""), "[" & tmpItem.序号 & "]") > 0 Then
                If intCol <> tmpItem.序号 Then
                    If MsgBox("该列的计算公式与第 " & tmpItem.序号 & " 列循环引用，这可能造成数据无法正确计算，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
                        cmdOK.Enabled = False
                        txtFormula.SetFocus: Exit Sub
                    End If
                End If
            End If
        Next
        
        '语法检测
        strErr = CheckFormula(txtFormula.Text)
        If strErr <> "" Then
            MsgBox strErr, vbInformation, App.Title
            cmdOK.Enabled = False
            txtFormula.SetFocus: Exit Sub
        End If
        
        '图片字段禁止特定选项
        If mblnBinary Then
            If chkMerge.Value = 1 Then
                MsgBox "图片数据列不能进行自动合并。", vbInformation, App.Title
                If chkMerge.Enabled And chkMerge.Visible Then chkMerge.SetFocus
                cmdOK.Enabled = False: Exit Sub
            End If
            If chk换页.Value = 1 Then
                MsgBox "图片数据列不能进行自动换页。", vbInformation, App.Title
                If chk换页.Enabled And chk换页.Visible Then chk换页.SetFocus
                cmdOK.Enabled = False: Exit Sub
            End If
            If chkAutoFont.Value = 1 Then
                MsgBox "图片数据列不能自动缩小字体。", vbInformation, App.Title
                If chkAutoFont.Enabled And chkAutoFont.Visible Then chkAutoFont.SetFocus
                cmdOK.Enabled = False: Exit Sub
            End If
        End If
    End If
        
    strInit = txtFormula.Text
    cmdOK.Enabled = True
    cmdOK.SetFocus
End Sub

Private Sub cmdVerify_GotFocus()
    tvw.Visible = False
End Sub

Private Sub Form_Activate()
    chkAutoRowHeight.Tag = CStr(chkAutoRowHeight.Value)
End Sub

Private Sub Form_Click()
    tvw.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        cmdAdd_Click
    ElseIf KeyCode = vbKeyA And Shift = 2 Then
        SelAll txtFormula
    ElseIf KeyCode = vbKeyEscape Then
        If tvw.Visible Then
            tvw.Visible = False
            txtFormula.SetFocus
        Else
            cmdCancel_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '同时也不准在公式框人按回车
    If KeyAscii = 13 And ActiveControl.name <> "txtItem" Then KeyAscii = 0: SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    Dim strRelation As String
    Dim strRelationID As String
    Dim i As Integer
    Set objScript = New clsScript
    srt.AddObject "clsScript", objScript, True
    
    blnDo = False
    txtFormula.Text = strInit
    txtFormat.Text = strFormat
    
    chk换页.Value = IIF(mbln换页, 1, 0)
    If Not mblnCan换页 Then
        chk换页.Value = 0
        chk换页.Enabled = False
    End If
    
    chkMerge.Value = IIF(mblnMerge, 1, 0)
    If Not mblnPreMerge Then
        chkMerge.Value = 0
        chkMerge.Enabled = False
    End If
    
    chkAutoRowHeight.Value = IIF(mblnAutoRowHeight, 1, 0)
    chkAutoFont.Value = IIF(mblnAutoFont, 1, 0)
    For i = 1 To mobjRelations.count
        If InStr(strRelation, "," & mobjRelations.Item(i).关联报表名称) = 0 Then
            strRelation = strRelation & "," & mobjRelations.Item(i).关联报表名称
        End If
        If InStr(strRelationID, "," & mobjRelations.Item(i).关联报表ID) = 0 Then
            strRelationID = strRelationID & "," & mobjRelations.Item(i).关联报表ID
        End If
    Next
    txtRelation.Text = Mid(strRelation, 2)
    txtRelation.Tag = Mid(strRelation, 2)
    If txtRelation.Tag <> "" Then txtRelation.Locked = True
    chkVisible.Value = IIF(mblnVisible, 1, 0)
    
    gblnOK = False
    Call CopyTree(frmParent.tvwSQL, tvw, True)
    blnDo = True
    
    If CheckFormula(txtFormula.Text) <> "" Then cmdOK.Enabled = False
End Sub

Private Sub Form_Paint()
    tvw.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objReport = Nothing
End Sub


Private Sub tvw_DblClick()
    tvw.Visible = False
    Call cmdAdd_Click
    txtFormula.SetFocus
End Sub

Private Sub tvw_LostFocus()
    tvw.Visible = False
    txtFormula.SetFocus
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key <> "Root" And Node.Children = 0 Then
        txtItem.Text = LevelText(Node)
    Else
        txtItem.Text = ""
    End If
    txtItem.SetFocus
End Sub

Private Sub txtFormat_KeyPress(KeyAscii As Integer)
    If InStr("!^|@'`~""", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtFormula_Change()
    If Not blnDo Then Exit Sub
    If txtFormula.Text = strInit Or txtFormula.Text = "" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Sub txtFormula_GotFocus()
    tvw.Visible = False
End Sub

Private Sub txtFormula_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = 2 Then SelAll txtFormula
End Sub

Private Sub txtFormula_KeyPress(KeyAscii As Integer)
    If InStr("~`!$%^}{?'|", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: VBA.Beep
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        cmdShow_Click
    ElseIf KeyCode = vbKeyUp Then
        If tvw.SelectedItem.Index - 1 >= 1 Then
            tvw.Nodes(tvw.SelectedItem.Index - 1).Selected = True
            tvw.SelectedItem.EnsureVisible
            Call tvw_NodeClick(tvw.SelectedItem)
            txtItem.SelStart = 0: txtItem.SelLength = 0
        End If
    ElseIf KeyCode = vbKeyDown Then
        If tvw.SelectedItem.Index + 1 <= tvw.Nodes.count Then
            tvw.Nodes(tvw.SelectedItem.Index + 1).Selected = True
            tvw.SelectedItem.EnsureVisible
            Call tvw_NodeClick(tvw.SelectedItem)
            txtItem.SelStart = 0: txtItem.SelLength = 0
        End If
    End If
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If tvw.Visible Then tvw_DblClick
    End If
End Sub

Private Sub txtItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tvw.Visible = False
End Sub

Private Function CheckFormula(ByVal strFormula As String) As String
'功能：检查表列公式书写是否正确,对其它列的引用变量视为数字型
    Dim strCheck As String, strLeft As String, strRight As String
    Dim strVar As String
    Dim X As Integer, Y As Integer, i As Integer
    
    mblnBinary = False
    strCheck = strFormula
    
    For i = 1 To Len(strCheck)
        If Mid(strCheck, i, 1) = "[" Then X = X + 1
        If Mid(strCheck, i, 1) = "]" Then Y = Y + 1
    Next
    If X <> Y Then
        CheckFormula = "数据源项目括号[]不配对！"
        Exit Function
    End If
    
    Call Randomize(timer)
    
    Do While InStr(strCheck, "[") > 0
        strLeft = Left(strCheck, InStr(strCheck, "[") - 1)
        strRight = Mid(strCheck, InStr(strCheck, "]") + 1)
        strVar = Mid(strCheck, InStr(strCheck, "[") + 1, InStr(strCheck, "]") - InStr(strCheck, "[") - 1)
        If IsNumeric(strVar) Or (Left(strVar, 1) = "@" And IsNumeric(Mid(strVar, 2))) Then
            strCheck = strLeft & Rnd * 55 & strRight
        ElseIf InStr(strVar, ".") > 0 Then
            If InStr(strVar, "@") > 0 Then
                CheckFormula = "字段数据不支持对上一行的引用！"
                Exit Function
            End If
            Select Case GetNodeType(strVar, tvw)
                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                    strCheck = strLeft & """字符串""" & strRight
                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                    strCheck = strLeft & Rnd * 55 & strRight
                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                    strCheck = strLeft & "cDate(""2000年08月04日"")" & strRight
                Case adBinary, adVarBinary, adLongVarBinary
                    mblnBinary = True
                    strCheck = strLeft & "图片字段" & strRight
                Case Else
                    CheckFormula = "不能确定数据[" & strVar & "]的来源！"
                    Exit Function
            End Select
        Else
            strCheck = strLeft & strRight
        End If
    Loop
    
    If mblnBinary Then
        If strCheck <> "图片字段" Then
            CheckFormula = "图片字段不能进行运算。"
        End If
    Else
        Err.Clear
        On Error Resume Next
        Call srt.Eval(strCheck)
        If srt.Error.Number <> 0 Then
            CheckFormula = "语法错误,详细信息为：         " & _
                vbCrLf & vbCrLf & srt.Error.Description & vbCrLf & vbCrLf & "请仔细检查！"
            srt.Error.Clear
        End If
    End If
End Function

Private Sub txtformat_GotFocus()
    SelAll txtFormat
End Sub

Private Sub txtRelation_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInfo As String
    Dim lngReportID As Long
    Dim strReportID As String
    Dim X As Long, Y As Long, k As Long
    
    If KeyCode = vbKeyReturn Then
        X = InStr(1, txtFormula.Text, "]")
        Y = InStr(1, txtFormula.Text, ".")
        k = InStr(1, txtFormula.Text, "[")
        If X > k And X > Y And X <> 0 And k <> 0 Then
            strReportID = FindReport(txtRelation.Text, txtRelation.hwnd, strInfo, Val(txtRelation.Tag), objReport, mobjRelations, 1, Me)
            If strReportID <> "" Then
                strReportID = Split(Split(strReportID, "(")(1), ")")(0)
                If strReportID <> "" Then
                    txtRelation.Text = strInfo
                    txtRelation.Tag = strReportID
                    txtRelation.Locked = True
                End If
            Else
                txtRelation.SetFocus
            End If
        Else
            MsgBox "当前列必须先绑定一个数据源，例如：[数据源.字段],绑定后再设置关联报表。", vbInformation, Me.Caption
        End If
    End If
End Sub

Private Sub txtRelation_LostFocus()
    If txtRelation.Text = "" Then txtRelation.Tag = ""
End Sub

