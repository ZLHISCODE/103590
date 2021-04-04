VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmBillSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保存模板"
   ClientHeight    =   5820
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5595
   ControlBox      =   0   'False
   Icon            =   "frmBillSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   2190
      TabIndex        =   8
      Top             =   5355
      Width           =   1100
   End
   Begin VB.CheckBox chkLocal 
      Caption         =   "保存为本人私有模板"
      Height          =   315
      Left            =   255
      TabIndex        =   7
      Top             =   5370
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.TextBox txt项目名称 
      Height          =   300
      Left            =   1560
      MaxLength       =   60
      TabIndex        =   3
      Top             =   4800
      Width           =   3915
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4380
      TabIndex        =   5
      Top             =   5350
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存(&O)"
      Height          =   350
      Left            =   3285
      TabIndex        =   4
      Top             =   5350
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   0
      TabIndex        =   6
      Top             =   5160
      Width           =   5565
   End
   Begin MSComctlLib.TreeView tvwElement 
      Height          =   4275
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   7541
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "iLstItem"
      Appearance      =   1
   End
   Begin VB.Label lbl项目名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "模板名称(&N)："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   4860
      Width           =   1170
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "请选择模板保存的所在分类目录(&D)："
      ForeColor       =   &H80000007&
      Height          =   180
      Left            =   285
      TabIndex        =   0
      Top             =   135
      Width           =   5100
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBillSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ifOK As Boolean
Private mlngElementID As Long, mstrContent As String

Public Function ShowMe(objParent As Object, ByVal ElementID As Long, ByVal Content As String) As Boolean
    mlngElementID = ElementID: mstrContent = Content
    ifOK = False
    
    If Not ShowTemplate(mlngElementID) Then ShowMe = False: Exit Function
    
    Me.Show vbModal, objParent
    ShowMe = ifOK
End Function

Private Function ShowTemplate(ByVal lngElementID As Long) As Boolean
'显示可用于当前元素的模板树
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim objCurrNode As MSComctlLib.Node
    
    Err = 0: On Error GoTo ErrHand
    ShowTemplate = False
    strSQL = "Select Distinct 0 As 末级,上级ID,ID,名称,'' As 内容,编码 From 病历模板分类" & _
        " Start With ID In" & _
        " (Select A.模板分类ID From 病历模板应用 A,病历模板分类 B Where A.模板分类ID=B.ID And 病历元素ID=[1] And " & _
        "(B.所属人员 Is Null Or B.所属人员='" & UserInfo.姓名 & "'))" & _
        " Connect By Prior 上级ID=ID" & _
        " Union All" & _
        " Select 1,a.分类ID,a.ID,a.名称,a.内容,a.编码 From 病历模板内容 a,病历模板应用 b,病历模板分类 c" & _
        " Where a.分类id=b.模板分类id And b.模板分类ID=c.ID And b.病历元素id=[1] And (c.所属人员 Is Null Or c.所属人员='" & UserInfo.姓名 & "') Order By 末级,编码"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngElementID)
    If rsTmp.EOF Then
        MsgBox "没有模板分类适用该报告文本，请先进入模板管理设置分类！", vbInformation, gstrSysName
        Exit Function
    End If
    
    tvwElement.Nodes.Clear
    Do While Not rsTmp.EOF
        With tvwElement
            If IsNull(rsTmp("上级ID")) Then
                Set objCurrNode = .Nodes.Add(, , IIf(rsTmp("末级") = 0, "C", "T") & rsTmp("ID"), rsTmp("名称"), IIf(rsTmp("末级") = 0, "Close", "Template"), IIf(rsTmp("末级") = 0, "Open", "Template"))
                objCurrNode.Expanded = True
            Else
                Set objCurrNode = .Nodes.Add("C" & rsTmp("上级ID"), tvwChild, IIf(rsTmp("末级") = 0, "C", "T") & rsTmp("ID"), rsTmp("名称"), IIf(rsTmp("末级") = 0, "Close", "Template"), IIf(rsTmp("末级") = 0, "Open", "Template"))
            End If
            objCurrNode.Tag = NVL(rsTmp("编码"))
        End With
        
        rsTmp.MoveNext
    Loop
    If tvwElement.Nodes.Count > 0 Then tvwElement.Nodes(1).Expanded = True
    ShowTemplate = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function NextCode(ByVal strClass As String) As String
    '功能获取指定分类下的编码
    Dim strTemp As String, rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Err = 0: On Error GoTo ErrHand
    strTemp = strClass
    strSQL = "select nvl(max(编码),'0000000000') as 编码" & _
            " From 病历模板内容" & _
            " Where 编码 like [1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, strTemp & "%")
    With rsTmp
        Err = 0: On Error Resume Next
        NextCode = strTemp & Right(String(10, "0") & Val(Mid(!编码, Len(strTemp) + 1)) + 1, _
            IIf(Len(!编码) - Len(strTemp) < Len(CStr(Val(Mid(!编码, Len(strTemp) + 1)) + 1)), _
                Len(CStr(Val(Mid(!编码, Len(strTemp) + 1)) + 1)), Len(!编码) - Len(strTemp)))
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    On Error GoTo ErrHandle
    Dim strKey As String
    Dim intIndex As Long
    
    If Me.tvwElement.SelectedItem Is Nothing Then Exit Sub
    
    On Error GoTo ErrHandle
    With tvwElement
        If Mid(.SelectedItem.Key, 1, 1) = "C" Then
            If MsgBox("真的删除该分类“" & .SelectedItem.Text & "”吗", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            gstrSql = "zl_病历模板分类_Delete(" & Mid(.SelectedItem.Key, 2) & ")"
            Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
        Else
            If MsgBox("你确认要删除模板：" & .SelectedItem.Text & "吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                gstrSql = "zl_病历模板内容_DELETE(" & Mid(.SelectedItem.Key, 2) & ")"
                Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
            End If
        End If
        Call .Nodes.Remove(.SelectedItem.Key)
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim strFormula As String
    Dim strErrorMsg As String, iErrorPos As Integer
    Dim strMid As Variant
    Dim i As Integer, lngVItemID0 As Long, strItemCode As String
    Dim lngItemID As Long
    
    If tvwElement.SelectedItem Is Nothing Then
        MsgBox "请选择一个模板分类！", vbInformation, gstrSysName
        tvwElement.SetFocus
        Exit Sub
    End If
    '一般特性检查
    If Trim(Me.txt项目名称.Text) = "" Then
        MsgBox "请输入模板名称！", vbInformation, gstrSysName
        Me.txt项目名称.SetFocus: Exit Sub
    End If
    
    '数据保存
    lngItemID = zlDatabase.GetNextId("病历模板内容")
    strItemCode = NextCode(tvwElement.SelectedItem.Tag)
    gstrSql = Mid(tvwElement.SelectedItem.Key, 2) & "," & lngItemID & ",'" & strItemCode & "'"
    gstrSql = gstrSql & ",'" & Replace(Trim(Me.txt项目名称.Text), "'", "''") & "','" & _
        zlCommFun.SpellCode(Me.txt项目名称.Text) & "','" & Replace(mstrContent, "'", "''") & "'"
'    If chkLocal.Value = 1 Then gstrSql = gstrSql & ",'" & UserInfo.姓名 & "'"
    gstrSql = "zl_病历模板内容_Insert(" & gstrSql & ")"
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    ifOK = True
    
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call cmdCancel_Click
End Sub

Private Sub Form_Load()
    chkLocal.Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "保存私有模板", "1"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "保存私有模板", chkLocal.Value)
End Sub

Private Sub tvwElement_Click()
    With tvwElement
        If .SelectedItem Is Nothing Then Exit Sub
        
        If Mid(.SelectedItem.Key, 1, 1) = "C" Then
            Me.cmdOK.Enabled = True
        Else
            Me.cmdOK.Enabled = False
        End If
    End With
End Sub

Private Sub tvwElement_DblClick()
    With tvwElement
        If .SelectedItem Is Nothing Then Exit Sub
        
        Call zlCommFun.PressKey(vbKeyTab)
    End With
End Sub

Private Sub txt项目名称_GotFocus()
    Me.txt项目名称.SelStart = 0: Me.txt项目名称.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt项目名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    If InStr(" ~!@#$%^&*_+|=`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt项目名称_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub
