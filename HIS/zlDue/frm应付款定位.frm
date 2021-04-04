VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm应付款定位 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "应付款定位条件"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   Icon            =   "frm应付款定位.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ListView lvwDept 
      Height          =   1290
      Left            =   435
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   2275
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1650
      MaxLength       =   100
      TabIndex        =   2
      Top             =   450
      Width           =   2355
   End
   Begin VB.Frame fraTemp 
      Height          =   75
      Left            =   -300
      TabIndex        =   9
      Top             =   1800
      Width           =   5505
   End
   Begin VB.CommandButton cmd上级 
      Caption         =   "…"
      Enabled         =   0   'False
      Height          =   240
      Left            =   3720
      TabIndex        =   6
      Top             =   1380
      Width           =   255
   End
   Begin VB.OptionButton opt定位 
      Caption         =   "按单据号定位(&N)"
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1725
   End
   Begin VB.OptionButton opt定位 
      Caption         =   "按药品供应商定位(&S)"
      Height          =   285
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   1020
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3090
      TabIndex        =   8
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1950
      TabIndex        =   7
      Top             =   2040
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   1650
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1350
      Width           =   2355
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "入库单据号(&M)"
      Height          =   180
      Index           =   1
      Left            =   450
      TabIndex        =   1
      Top             =   540
      Width           =   1170
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "供应商(&U)"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   810
      TabIndex        =   4
      Top             =   1440
      Width           =   810
   End
End
Attribute VB_Name = "frm应付款定位"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mstr单据号 As String
Dim mstr供应商ID As String
Dim msngDownX As Single
Dim msngDownY As Single
Private mstrPrivs As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    For lngIndex = txtEdit.LBound To txtEdit.UBound
        If txtEdit(lngIndex).Enabled = True Then
            If StrIsValid(txtEdit(lngIndex).Text, txtEdit(lngIndex).MaxLength) = False Then
                txtEdit(lngIndex).SetFocus
                Exit Sub
            End If
            
            Select Case lngIndex
                Case 0
                    mstr供应商ID = txtEdit(lngIndex).Tag
                Case 1
                    mstr单据号 = UCase(Trim(txtEdit(lngIndex).Text))
            End Select
        End If
    Next
    
    If mstr单据号 = "" And mstr供应商ID = "" Then
        MsgBox "请输入定位条件。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd上级_Click()
    Dim rs供应商 As New ADODB.Recordset
    Dim str权限 As String
    str权限 = " and (末级<>1 or ( 末级=1 " & zl_获取站点限制 & "  and " & Get分类权限(gstrPrivs) & "))"
        
    gstrSQL = "" & _
        "   Select id,上级ID,末级,编码,简码,名称 " & _
        "   From 供应商 " & _
        "   Where nvl(撤档时间,to_date('3000-01-01','yyyy-MM-dd'))=to_date('3000-01-01','yyyy-MM-dd') " & str权限 & _
        "   start with 上级ID is null connect by prior ID =上级ID " & _
        "   order by level,ID"
    On Error GoTo errHandle
    zlDatabase.OpenRecordset rs供应商, gstrSQL, Me.Caption
    
    txtEdit(0).SetFocus
    If rs供应商.EOF Then
        rs供应商.Close
        Exit Sub
    End If
    With frm供应商选择
        Me.Tag = .SelDept(mstrPrivs)
        If Me.Tag <> "" Then
            txtEdit(0).Tag = Left(Me.Tag, InStr(Me.Tag, ",") - 1)
            txtEdit(0).Text = Mid(Me.Tag, InStr(Me.Tag, ",") + 1)
        End If
    End With
    Unload frm供应商选择
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnOK = False
End Sub

Public Function Get定位条件(ByVal strPrivs As String, str单据号 As String, str供应商ID As String) As Boolean
    mstrPrivs = strPrivs
    frm应付款定位.Show vbModal, frm清单管理
    
    Get定位条件 = mblnOK
    If mblnOK = True Then
        str单据号 = mstr单据号
        str供应商ID = mstr供应商ID
    End If
End Function

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Private Sub lvwDept_DblClick()
    If lvwDept.HitTest(msngDownX, msngDownY) Is Nothing Then Exit Sub
    txtEdit(0).Tag = Mid(lvwDept.SelectedItem.Key, 2)
    txtEdit(0).Text = lvwDept.SelectedItem.SubItems(1)
    cmdOK.SetFocus
    lvwDept.Visible = False
End Sub

Private Sub lvwDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Not (lvwDept.SelectedItem Is Nothing) Then
        txtEdit(0).Tag = Mid(lvwDept.SelectedItem.Key, 2)
        txtEdit(0).Text = lvwDept.SelectedItem.SubItems(1)
        cmdOK.SetFocus
        lvwDept.Visible = False
    ElseIf KeyCode = 27 Then
        txtEdit(0).SetFocus
        lvwDept.Visible = False
    End If
End Sub

Private Sub lvwDept_LostFocus()
    lvwDept.Visible = False
End Sub

Private Sub lvwDept_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownX = X
    msngDownY = Y
End Sub

Private Sub opt定位_Click(Index As Integer)
    txtEdit(0).Enabled = opt定位(0).Value
    lbl(0).Enabled = opt定位(0).Value
    cmd上级.Enabled = opt定位(0).Value
    
    txtEdit(1).Enabled = opt定位(1).Value
    lbl(1).Enabled = opt定位(1).Value
    
    txtEdit(Index).SetFocus
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rstTemp As New ADODB.Recordset, strSQL As String, ltmDept As ListItem
    Dim str权限 As String, strKey As String
    
    str权限 = " and " & Get分类权限(gstrPrivs)
    On Error GoTo errHandle
    If Index = 0 And KeyAscii = 13 Then
        strKey = GetMatchingSting(txtEdit(0).Text, False)
        'by lesfeng 2009-12-2 性能优化
        If IsNumeric(txtEdit(0).Text) Then
            strSQL = "" & _
                "   Select id,上级ID,末级,编码,简码,名称 From 供应商 " & _
                "   Where 末级=1 " & zl_获取站点限制 & "  " & _
                "        And 编码 Like [1] " & str权限
        Else
            strSQL = "Select id,上级ID,末级,编码,简码,名称 From 供应商 Where 末级=1 " & zl_获取站点限制 & " And (简码 Like [1] Or 名称 Like [1]) " & str权限
        End If
        Set rstTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strKey)
        
        If rstTemp.EOF Then
            MsgBox "指定的供应商不存在，请重新输入。", vbInformation, Me.Caption
            txtEdit(0).SetFocus
        ElseIf rstTemp.RecordCount > 1 Then
            lvwDept.ListItems.Clear
            While Not rstTemp.EOF
                Set ltmDept = lvwDept.ListItems.Add(, "D" & rstTemp!ID, rstTemp!编码)
                ltmDept.ListSubItems.Add , , rstTemp!名称
                rstTemp.MoveNext
            Wend
            Set lvwDept.SelectedItem = lvwDept.ListItems(1)
            lvwDept.ColumnHeaders(1).Width = 1000
            lvwDept.ColumnHeaders(2).Width = lvwDept.Width - 1300
            lvwDept.Visible = True
            lvwDept.SetFocus
        Else
            txtEdit(0).Tag = rstTemp!ID
            txtEdit(0).Text = rstTemp!名称
            cmdOK.SetFocus
        End If
    End If
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    Dim intYear  As Integer, strYear As String
    If IsNumeric(txtEdit(Index)) And txtEdit(Index).Text <> "" And Index = 1 Then
        If Len(txtEdit(1).Text) < 8 And Len(txtEdit(1)) > 0 Then
            txtEdit(1).Text = UCase(LTrim(txtEdit(1).Text))
            intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            txtEdit(1).Text = strYear & String(7 - Len(txtEdit(1).Text), "0") & txtEdit(1).Text
        End If
    End If
    txtEdit(1).Text = UCase(txtEdit(1).Text)
End Sub
