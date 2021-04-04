VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresAdjust 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "人员部门调整"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   Icon            =   "frmPresAdjust.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDel 
      Caption         =   "<"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   ">"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5160
      TabIndex        =   7
      Top             =   4680
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6480
      TabIndex        =   8
      Top             =   4680
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvwSelect 
      Height          =   3735
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   6588
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.PictureBox picPerson 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7800
      Begin VB.TextBox txtLocate 
         Height          =   320
         Left            =   6480
         TabIndex        =   11
         ToolTipText     =   "查找下一个F3或回车，定位输入框F4"
         Top             =   110
         Width           =   1095
      End
      Begin VB.Label lblLocate 
         BackColor       =   &H80000005&
         Caption         =   "查找(&F)"
         Height          =   255
         Left            =   5760
         TabIndex        =   10
         Top             =   150
         Width           =   735
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "姓名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   180
         Width           =   585
      End
   End
   Begin MSComctlLib.ListView lvwChoose 
      Height          =   3765
      Left            =   4200
      TabIndex        =   9
      ToolTipText     =   "双击选项确定缺省，且只能有一个缺省。"
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   6641
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "所属部门"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "缺省"
         Object.Width           =   970
      EndProperty
   End
   Begin VB.Label lblChoose 
      AutoSize        =   -1  'True
      Caption         =   "所属部门(&H)"
      Height          =   180
      Left            =   4200
      TabIndex        =   5
      Top             =   600
      Width           =   990
   End
   Begin VB.Label lblDeptSelect 
      AutoSize        =   -1  'True
      Caption         =   "可选部门(&S)"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   990
   End
End
Attribute VB_Name = "frmPresAdjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPersonID As Long        '人员ID
Private mstrPrivs As String         '权限
Private mDataChange As Boolean      '是否有修改

Private Sub cmdAdd_Click()
    If tvwSelect.SelectedItem Is Nothing Then Exit Sub
    If FindKey(tvwSelect.SelectedItem.Key, 0) = False And tvwSelect.SelectedItem.ForeColor <> &H8000000C Then
        lvwChoose.ListItems.Add , tvwSelect.SelectedItem.Key, tvwSelect.SelectedItem.Text
        tvwSelect.SelectedItem.Bold = True
    End If
    tvwSelect.SetFocus
End Sub

Public Sub EntryPort(ByVal lngPersonID As Long, ByVal strPrivs As String)
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    mlngPersonID = lngPersonID
    mstrPrivs = strPrivs
    gstrSQL = "select 姓名 from 人员表 where id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Caption & "-人员姓名", mlngPersonID)
    If rsTmp.RecordCount = 1 Then
        lblName.Caption = "姓名： " & rsTmp!姓名
    Else
        lblName.Caption = "姓名："
    End If
    Call InitChooseLvw: Call lvwChoose_Click
    Call InitSelectTvw: Call tvwSelect_Click
    Exit Sub
errHandle:
    If errCenter() = 1 Then Resume
    Call saveErrLog
End Sub

Private Sub InitSelectTvw()
'初始化
    Dim rsDeptID As ADODB.Recordset
    Dim nodTmp As Node
    
    On Error GoTo errHandle
    If InStr(mstrPrivs, "所有部门") = 0 Then
        gstrSQL = "Select Max(Level) as 层,A.ID,A.上级ID,A.名称,'【'||A.编码||'】' 编码,Upper(a.简码) as 简码 " & _
                  "From 部门表 A Start With ID IN(Select 部门ID From 部门人员 Where 人员ID=[1]) Connect by Prior 上级ID=ID " & _
                  "Group by A.ID,A.上级ID,A.名称,A.编码,a.简码 " & _
                  "Order by A.编码,层 Desc"
        Set rsDeptID = zlDatabase.OpenSQLRecord(gstrSQL, Caption, glngUserId)
        With tvwSelect
            .Sorted = True
            .Nodes.Clear
            Do While Not rsDeptID.EOF
                If IIF(IsNull(rsDeptID!上级ID), 0, rsDeptID!上级ID) = 0 Then
                    If tvwSelect.Nodes.Count > 0 Then
                        If FindKey("K" & rsDeptID!ID, 1) = False Then
                            Set nodTmp = .Nodes.Add(, , "K" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称)
                        Else
                            Set nodTmp = .Nodes("K" & rsDeptID!ID)
                        End If
                    Else
                        Set nodTmp = .Nodes.Add(, , "K" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称)
                    End If
                Else
                    If FindKey("K" & rsDeptID!ID, 1) = False Then
                        Set nodTmp = .Nodes.Add("K" & rsDeptID!上级ID, tvwChild, "K" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称)
                    Else
                        Set nodTmp = .Nodes("K" & rsDeptID!ID)
                    End If
                End If
                nodTmp.Tag = rsDeptID!简码
                nodTmp.ForeColor = &H8000000C
                rsDeptID.MoveNext
            Loop
            rsDeptID.Close
        End With
        '生成子结点
        gstrSQL = "Select ID,上级ID,'【'||编码||'】' 编码,名称,Upper(a.简码) as 简码 " & _
                  "From 部门表 A " & _
                  "Start With ID IN(Select 部门ID From 部门人员 Where 人员ID=[1]) Connect by Prior ID=上级ID"
        Set rsDeptID = zlDatabase.OpenSQLRecord(gstrSQL, Caption, glngUserId)
        With tvwSelect
            Do While Not rsDeptID.EOF
                If IIF(IsNull(rsDeptID!上级ID), 0, rsDeptID!上级ID) = 0 Then
                    If tvwSelect.Nodes.Count > 0 Then
                        If FindKey("K" & rsDeptID!ID, 1) = False Then
                            Set nodTmp = .Nodes.Add(, , "K" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称)
                        Else
                            Set nodTmp = .Nodes("K" & rsDeptID!ID)
                        End If
                    Else
                        Set nodTmp = .Nodes.Add(, , "K" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称)
                    End If
                Else
                    If FindKey("K" & rsDeptID!ID, 1) = False Then
                        Set nodTmp = .Nodes.Add("K" & rsDeptID!上级ID, tvwChild, "K" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称)
                    Else
                        Set nodTmp = .Nodes("K" & rsDeptID!ID)
                    End If
                End If
                nodTmp.Tag = rsDeptID!简码
                nodTmp.ForeColor = vbBlack
                rsDeptID.MoveNext
            Loop
            rsDeptID.Close
        End With
        
        For Each nodTmp In tvwSelect.Nodes
            If FindKey(nodTmp.Key, 0) Then
                nodTmp.Bold = True
            End If
        Next
    Else
        gstrSQL = "select ID,上级ID,'【'||编码||'】' 编码,名称,Upper(a.简码) as 简码 from 部门表 A " & _
                  "where 撤档时间=to_date('3000-1-1','yyyy-mm-dd') and substr(编码,1)<>'-' " & _
                  "start with 上级id is null connect by prior id=上级id"
        Set rsDeptID = zlDatabase.OpenSQLRecord(gstrSQL, Caption)
        With tvwSelect
            .Sorted = True
            .Nodes.Clear
            Do While Not rsDeptID.EOF
                If IIF(IsNull(rsDeptID!上级ID), 0, rsDeptID!上级ID) = 0 Then
                    Set nodTmp = .Nodes.Add(, , "K" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称)
                Else
                    Set nodTmp = .Nodes.Add("K" & rsDeptID!上级ID, tvwChild, "K" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称)
                End If
                If FindKey(nodTmp.Key, 0) Then
                    nodTmp.Bold = True
                End If
                nodTmp.Tag = rsDeptID!简码
                rsDeptID.MoveNext
            Loop
            rsDeptID.Close
        End With
    End If
    If tvwSelect.Nodes.Count > 0 Then tvwSelect.Nodes(1).Selected = True
    Exit Sub
errHandle:
    If errCenter() = 1 Then Resume
    Call saveErrLog
End Sub

Private Sub InitChooseLvw()
'初始化
    Dim rsDeptID As ADODB.Recordset
    Dim lstTmp As ListItem
    
    On Error GoTo errHandle
    gstrSQL = "select a.部门id,'【'||b.编码||'】' 编码,b.名称,a.缺省 from 部门人员 a, 部门表 b where a.部门id=b.id and a.人员id=[1]"
    Set rsDeptID = zlDatabase.OpenSQLRecord(gstrSQL, Caption, mlngPersonID)
    
    With lvwChoose
        .Sorted = True
        .ListItems.Clear
        Do While Not rsDeptID.EOF
            Set lstTmp = .ListItems.Add(, "K" & rsDeptID!部门id, rsDeptID!编码 & rsDeptID!名称)
            If rsDeptID!缺省 = 1 Then
                lstTmp.SubItems(1) = "√"
            End If
            rsDeptID.MoveNext
        Loop
        .ListItems(1).Selected = True
    End With
    Exit Sub
errHandle:
    If errCenter() = 1 Then Resume
    Call saveErrLog
End Sub

Private Function FindKey(ByVal strKey As String, ByVal bytType As Byte) As Boolean
    If bytType = 0 Then
        Dim lstTmp As ListItem
        For Each lstTmp In lvwChoose.ListItems
            If lstTmp.Key = strKey Then
                FindKey = True
                Exit Function
            End If
        Next
    Else
        Dim nodTmp As Node
        For Each nodTmp In tvwSelect.Nodes
            If nodTmp.Key = strKey Then
                FindKey = True
                Exit Function
            End If
        Next
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    If lvwChoose.SelectedItem Is Nothing Then Exit Sub
    Dim i As Long
    Dim strKey As String
    Dim nodTmp As Node
    
    If lvwChoose.SelectedItem Is Nothing Then Exit Sub
    If lvwChoose.SelectedItem.SubItems(1) = "√" Then Exit Sub
    
    i = lvwChoose.SelectedItem.Index
    strKey = lvwChoose.SelectedItem.Key
    lvwChoose.ListItems.Remove i
    For Each nodTmp In tvwSelect.Nodes
        If nodTmp.Key = strKey Then
            nodTmp.Bold = False
            Exit For
        End If
    Next
    If lvwChoose.ListItems.Count > 0 Then
        If lvwChoose.ListItems.Count > i - 1 Then
            lvwChoose.ListItems(i).Selected = True
        Else
            lvwChoose.ListItems(i - 1).Selected = True
        End If
        '判断是否有该部门的权限
        If InStr(mstrPrivs, "所有部门") = 0 And CheckDeptPermission(mlngPersonID, Mid(lvwChoose.SelectedItem.Key, 2)) = False Then
            cmdDel.Enabled = False
        Else
            cmdDel.Enabled = lvwChoose.SelectedItem.SubItems(1) = ""
        End If
    End If
    lvwChoose.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim str部门ID As String ', str科室名称 As String
    
    '把所有部门做成一个串，选中的为1
    With lvwChoose
        For i = 1 To .ListItems.Count
            str部门ID = str部门ID & Mid(.ListItems(i).Key, 2) & ":"
            If .ListItems(i).SubItems(1) = "√" Then
                str部门ID = str部门ID & "1;"
'                str科室名称 = Mid(.ListItems(i).Text, InStr(.ListItems(i).Text, "】") + 1)
            Else
                str部门ID = str部门ID & "0;"
            End If
        Next
    End With
    gstrSQL = "zl_部门人员_update(" & mlngPersonID & ",'" & str部门ID & "')"
    On Error GoTo errHandle
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Unload Me
    Exit Sub
errHandle:
    If errCenter() = 1 Then Resume
    Call saveErrLog
End Sub

Private Sub lvwChoose_Click()
    If lvwChoose.SelectedItem Is Nothing Then Exit Sub
    If InStr(mstrPrivs, "所有部门") = 0 And CheckDeptPermission(mlngPersonID, Mid(lvwChoose.SelectedItem.Key, 2)) = False Then
        cmdDel.Enabled = False
    Else
        cmdDel.Enabled = lvwChoose.SelectedItem.SubItems(1) = ""
    End If
End Sub

Private Sub lvwChoose_DblClick()
    Dim lstTmp As ListItem
    If lvwChoose.SelectedItem Is Nothing Then Exit Sub
    If InStr(mstrPrivs, "所有部门") = 0 And CheckDeptPermission(mlngPersonID, Mid(lvwChoose.SelectedItem.Key, 2)) = False Then
        cmdDel.Enabled = False
        Exit Sub
    End If
    For Each lstTmp In lvwChoose.ListItems
        If lstTmp = lvwChoose.SelectedItem Then
            lvwChoose.SelectedItem.SubItems(1) = "√"
        Else
            lstTmp.SubItems(1) = ""
        End If
    Next
    cmdDel.Enabled = lvwChoose.SelectedItem.SubItems(1) = ""
End Sub

Private Sub lvwChoose_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then Call lvwChoose_DblClick
End Sub

Private Sub tvwSelect_Click()
    If tvwSelect.SelectedItem Is Nothing Then Exit Sub
    If tvwSelect.SelectedItem.ForeColor = &H8000000C Or FindKey(tvwSelect.SelectedItem.Key, 0) Then
        cmdAdd.Enabled = False
    Else
        cmdAdd.Enabled = True
    End If
End Sub

Private Sub tvwSelect_DblClick()
    If tvwSelect.SelectedItem Is Nothing Then Exit Sub
    Call cmdAdd_Click
End Sub


Private Sub tvwSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        tvwSelect_DblClick
    End If
End Sub

Private Sub txtLocate_GotFocus()
    zlControl.TxtSelAll txtLocate
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    Dim i As Long, lngStart As Long
    
    If KeyAscii = vbKeyReturn Then
        If txtLocate.Tag <> txtLocate.Text Then
            lblLocate.Tag = ""
            txtLocate.Tag = txtLocate.Text
        End If
        
        lngStart = Val("" & lblLocate.Tag) + 1
        If lngStart >= tvwSelect.Nodes.Count Then lngStart = 1
    
        For i = lngStart To tvwSelect.Nodes.Count
            If tvwSelect.Nodes(i).Text Like "*" & txtLocate.Text & "*" Or tvwSelect.Nodes(i).Tag Like "*" & UCase(txtLocate.Text) & "*" Then
                Call tvwSelect.Nodes(i).EnsureVisible
                tvwSelect.Nodes(i).Selected = True
                lblLocate.Tag = i
                tvwSelect.SetFocus
                Exit For
            End If
        Next
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txtLocate_KeyPress (vbKeyReturn)
    ElseIf KeyCode = vbKeyF4 Then
        If txtLocate.Visible And txtLocate.Enabled Then txtLocate.SetFocus
    End If
End Sub
