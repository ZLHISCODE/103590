VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMedicalTeamMember 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医疗小组成员"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   Icon            =   "frmMedicalTeamMember.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraView 
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   4335
      Begin VB.TextBox txtMember 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cboTeam 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtTeam 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblMember 
         AutoSize        =   -1  'True
         Caption         =   "小组成员(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   990
      End
      Begin VB.Label lblTrans 
         AutoSize        =   -1  'True
         Caption         =   "转入小组(&T)"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   1250
         Width           =   990
      End
      Begin VB.Label lblFromTeam 
         AutoSize        =   -1  'True
         Caption         =   "所属小组(&F)"
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   770
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3480
      TabIndex        =   1
      Top             =   4950
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2280
      TabIndex        =   0
      Top             =   4950
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvwList 
      Height          =   2655
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   4683
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   476
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "imgTvw"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList Img小图标 
      Left            =   3120
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalTeamMember.frx":000C
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalTeamMember.frx":0326
            Key             =   "User"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalTeamMember.frx":0640
            Key             =   "Role"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMedicalTeamMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytStatus As Byte
Private mlngDeptID As Long, mlngTeamID As Long, mlngMemberID As Long
Public mblnOK As Boolean
Public mstrPrivs As String

Property Get Status() As Byte
'状态 1-添加成员; 2-转小组
    Status = mbytStatus
End Property
Property Let Status(ByVal bytStatus As Byte)
    Caption = "医疗小组成员"
    '界面控制
    If bytStatus = 1 Then
        Caption = Caption & "-添加"
    Else
        Caption = Caption & "-转小组"
    End If
    
    '刷新rpcView控件
    RefreshViewRPC bytStatus
    mbytStatus = bytStatus
End Property

Public Sub ShowMe(ByVal frmVal As Form, ByVal bytStatus As Byte, ByVal lngDeptID As Long, _
ByVal lngTeamID As Long, Optional ByVal lngMemberID As Long)
    Dim rsTmp As ADODB.Recordset
    Dim nodTmp As Node, nodParent As Node
    mlngDeptID = lngDeptID
    mlngTeamID = lngTeamID
    mlngMemberID = lngMemberID
    Status = bytStatus
    
    On Error GoTo ErrHandle
    If Status = 1 Then
        gstrSQL = "select a.ID,a.名称 from 部门表 a, 临床医疗小组 b where a.id=b.科室id and b.id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngTeamID)
        If rsTmp.RecordCount = 1 Then
            '科室定位
            For Each nodTmp In tvwList.Nodes
                If Val(Mid(nodTmp.Key, 2)) = rsTmp!ID Then
                    Set nodParent = nodTmp
                    nodTmp.Expanded = True
                    nodTmp.Selected = True
                    Do Until nodParent Is Nothing
                        Set nodParent = nodParent.Parent
                        If Not nodParent Is Nothing Then
                            Set nodTmp = nodParent
                            nodTmp.Expanded = True
                        End If
                    Loop
                    Exit For
                End If
            Next
        End If
    End If
    Show vbModal, frmVal
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Status = 1 Then
        If SaveMember() Then
            Unload Me
        Else
            MsgBox "未勾选人员！", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        If Me.cboTeam.ListIndex < 0 Then
            MsgBox "未选择医疗小组！", vbInformation, gstrSysName
            Exit Sub
        End If
        If TransMember() Then
            Unload Me
'        Else
'            MsgBox "未选择人员！", vbInformation, gstrSysName
'            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    With tvwList
        .ImageList = Me.Img小图标
        .LabelEdit = tvwManual
    End With
    mblnOK = False
End Sub

Private Sub Form_Resize()
    If Status = 1 Then
        With tvwList
            .Visible = True
            .Top = 0
            .Left = 0
            .Width = Me.ScaleWidth
            .Height = Me.ScaleHeight - Me.cmdOK.Height - 200
        End With
        fraView.Visible = False
    Else
        With fraView
            .Visible = True
            .Top = 10
            .Left = 100
            .Width = Me.ScaleWidth - 200
        End With
        cmdOK.Top = fraView.Top + fraView.Height + 100
        cmdCancel.Top = cmdOK.Top
        Top = Top + (Height - (cmdOK.Top + cmdOK.Height + 600)) / 2
        Height = cmdOK.Top + cmdOK.Height + 600
        tvwList.Visible = False
    End If
End Sub

Private Sub RefreshViewRPC(ByVal bytStatus As Byte)
    Dim i As Long, lngChars As Long
    Dim objNode As Node
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If bytStatus = 1 Then
        '添加成员
        gstrSQL = "select id,编码,名称,上级id From 部门表 where 编码<>'-'" & _
                  "  and (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) " & _
                  "  and id=[1] " & _
                  "start with 上级id is null connect by prior id=上级id "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDeptID)
        With tvwList
            .Nodes.Clear
            Do While Not rsTmp.EOF
'                If IsNull(rsTmp("上级id")) Then
                    Set objNode = .Nodes.Add(, , "D" & rsTmp("id"), "【" & rsTmp("编码") & "】" & rsTmp("名称"), "Dept", "Dept")
'                Else
'                    Set objNode = .Nodes.Add("D" & rsTmp("上级id"), tvwChild, "D" & rsTmp("id").Value, "【" & rsTmp("编码") & "】" & rsTmp("名称"), "Dept", "Dept")
'                End If
                rsTmp.MoveNext
            Loop
            rsTmp.Close
'            If InStr(mstrPrivs, "所有科室") = 0 Then
                gstrSQL = "Select a.Id, a.编号, a.姓名, b.部门id, (select 1 from 医疗小组人员 where 小组id=[1] and 人员id=a.id) 小组成员" & vbNewLine & _
                          "From 人员表 A, 部门表 C, 部门人员 B,部门性质说明 D,人员性质说明 E " & vbNewLine & _
                          "Where (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And a.Id = b.人员id And b.部门ID=[2] " & vbNewLine & _
                          " And b.部门id = c.Id And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null)" & vbNewLine & _
                          " And a.ID=e.人员ID and e.人员性质='医生' and c.ID=d.部门ID and d.工作性质='临床' and d.服务对象 in (2,3) " & vbNewLine & _
                          "Order by a.编号 "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID, mlngDeptID)
'            Else
'                gstrSQL = "Select a.Id, a.编号, a.姓名, b.部门id, (select 1 from 医疗小组人员 where 小组id=[1] and 人员id=a.id) 小组成员" & vbNewLine & _
'                          "From 人员表 A, 部门表 C, 部门人员 B,部门性质说明 D,人员性质说明 E " & vbNewLine & _
'                          "Where (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And a.Id = b.人员id And b.缺省 = 1 And b.部门id = c.Id And" & vbNewLine & _
'                          "  (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null)" & vbNewLine & _
'                          " and a.ID=e.人员ID and e.人员性质='医生' and c.ID=d.部门ID and d.工作性质='临床' and d.服务对象 in (2,3) " & vbNewLine & _
'                          "Order by a.编号 "
'                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID)
'            End If
'            gstrSQL = "Select a.Id, a.编号, a.姓名, b.部门id, (Select 1 From 医疗小组人员 Where 小组id = [1] And 人员id = a.Id) 小组成员 " & vbNewLine & _
'                      "From 人员表 A, 部门表 C, 部门人员 B, 部门性质说明 D, 人员性质说明 E, 部门人员 F " & vbNewLine & _
'                      "Where (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And a.Id = f.人员id And b.部门id = c.Id And " & vbNewLine & _
'                      "      b.部门id = f.部门id And b.人员id = [2] And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) And " & vbNewLine & _
'                      "      a.Id = e.人员id And e.人员性质 = '医生' And c.Id = d.部门id And d.工作性质 = '临床' And d.服务对象 In (2, 3) " & vbNewLine & _
'                      "Order By a.编号 "
'            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID, glngUserId)
            Do Until rsTmp.EOF
                Set objNode = .Nodes.Add("D" & rsTmp("部门id"), 4, "P" & rsTmp("id"), "【" & rsTmp("编号") & "】" & rsTmp("姓名"), "User", "User")
                objNode.ForeColor = RGB(0, 0, 255)
                If rsTmp!小组成员 = 1 Then
                    objNode.Checked = True
                End If
                rsTmp.MoveNext
            Loop
            rsTmp.Close
        End With
    Else
        '转小组
        gstrSQL = "select 编号,姓名 from 人员表 where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngMemberID)
        If rsTmp.RecordCount > 0 Then
            txtMember.Text = "【" & rsTmp!编号 & "】" & rsTmp!姓名
        End If
        rsTmp.Close
        gstrSQL = "Select 名称 From 临床医疗小组 Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID)
        If rsTmp.RecordCount > 0 Then
            txtTeam.Text = IIF(IsNull(rsTmp!名称), "", rsTmp!名称)
            txtTeam.Tag = mlngTeamID
            rsTmp.Close
            cboTeam.Clear
            If InStr(mstrPrivs, "所有科室") = 0 Then
                gstrSQL = "Select a.ID, a.名称, a.科室ID, c.名称 科室 From 临床医疗小组 a, 部门人员 b, 部门表 c " & _
                          "Where a.ID <> [1] and a.科室ID=b.部门ID and b.人员ID=[2] and b.部门id=c.id and substr(a.名称,1,1)<>'-' " & _
                          "  and not a.ID in (select 小组id from 医疗小组人员 where 人员id=[3]) " & _
                          "  and a.撤档时间=to_date('3000-1-1', 'YYYY-MM-DD') order by a.科室ID,a.名称"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID, glngUserId, mlngMemberID)
            Else
                gstrSQL = "Select a.ID, a.名称, a.科室ID, b.名称 科室 From 临床医疗小组 a, 部门表 b " & _
                          "Where a.ID <> [1] and substr(a.名称,1,1)<>'-' " & _
                          "  and not a.ID in (select 小组id from 医疗小组人员 where 人员id=[2]) " & _
                          "  and a.撤档时间=to_date('3000-1-1', 'YYYY-MM-DD') and a.科室id=b.id order by a.科室ID,a.名称"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID, mlngMemberID)
            End If
            For i = 0 To rsTmp.RecordCount - 1
                cboTeam.AddItem IIF(IsNull(rsTmp!科室), "", rsTmp!科室) & " | " & IIF(IsNull(rsTmp!名称), "", rsTmp!名称)
                cboTeam.ItemData(i) = rsTmp!ID
                If rsTmp!科室ID = mlngDeptID And cboTeam.ListIndex < 0 Then cboTeam.ListIndex = i
                If Len(rsTmp!名称 & rsTmp!科室) > lngChars Then lngChars = Len(rsTmp!名称 & rsTmp!科室)
                rsTmp.MoveNext
            Next
            zlControl.CboSetWidth cboTeam.hwnd, (lngChars + 2) * 15 * 6.5
            If cboTeam.ListIndex < 0 And cboTeam.ListCount > 0 Then cboTeam.ListIndex = 0
        End If
        rsTmp.Close
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tvwList_Click()
    Dim objNode As Node
    Set objNode = tvwList.SelectedItem
    If Left(objNode.Key, 1) = "D" Then objNode.Checked = False
End Sub

Private Sub tvwList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If Left(tvwList.SelectedItem.Key, 1) = "D" Then
            tvwList.SelectedItem.Checked = False
            KeyCode = 0
        End If
    End If
End Sub

Private Sub tvwList_NodeCheck(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
End Sub

Private Sub tvwList_NodeClick(ByVal Node As MSComctlLib.Node)
    If Left(Node.Key, 1) = "D" Then
        Node.Checked = False
    End If
End Sub

Private Function SaveMember() As Boolean
    Dim objNode As Node
    Dim strMemberIDs As String
    For Each objNode In tvwList.Nodes
        If objNode.Checked Then
            strMemberIDs = strMemberIDs & Mid(objNode.Key, 2, 20) & ";"
        End If
    Next
    If strMemberIDs = "" Then Exit Function
    
    On Error GoTo ErrHandle
    gstrSQL = "ZL_医疗小组人员_INSERT(" & mlngTeamID & ",'" & Left(strMemberIDs, Len(strMemberIDs) - 1) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    SaveMember = True
    mblnOK = True
    Exit Function
    
ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function TransMember() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strMess As String
    Dim i As Long
    On Error GoTo ErrHandle
    '如果住院医师下有病人就提示
'    gstrSQL = "Select a.病人id, a.住院号, a.出院病床, b.姓名" & vbNewLine & _
'              "From 病案主页 a, 病人信息 b " & vbNewLine & _
'              "Where a.住院医师 = (Select 姓名" & vbNewLine & _
'              "              From 人员表" & vbNewLine & _
'              "              Where ID = [2] And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)) And" & vbNewLine & _
'              "      a.医疗小组id = [1] and a.病人id=b.病人id and b.在院=1 "
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtTeam.Tag, mlngMemberID)
'    With rsTmp
'        For i = 1 To .RecordCount
'            strMess = strMess & "姓名：" & !姓名 & "；" & vbTab & _
'                      "住院号：" & IIF(IsNull(!住院号), "", !住院号) & "；" & vbTab & _
'                      "床号：" & IIF(IsNull(!出院病床), "", !出院病床) & vbTab & vbNewLine
'            .MoveNext
'        Next
'    End With
    strMess = MedicalTeamPatients(Val(txtTeam.Tag), mlngMemberID)
    If strMess <> "" Then
        If MsgBox("该医生当前有以下在院病人，" & vbNewLine & vbNewLine & strMess & vbNewLine & "确定以上病人的医疗小组也将一并转入吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    '判断所属小组是否存在
    gstrSQL = "select count(*) rec from 医疗小组人员 where 小组ID=[1] and 人员ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngTeamID, mlngMemberID)
    If rsTmp!rec = 1 Then
        '转小组
        gstrSQL = "Zl_医疗小组人员_Update("
        gstrSQL = gstrSQL & txtTeam.Tag & ","                            '来自小组ID
        gstrSQL = gstrSQL & mlngMemberID & ","                           '人员ID
        gstrSQL = gstrSQL & cboTeam.ItemData(cboTeam.ListIndex) & ",'"   '转入小组ID
        gstrSQL = gstrSQL & gstrUserCode & "','"                         '操作员编号
        gstrSQL = gstrSQL & gstrUserName & "')"                          '操作员姓名
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        TransMember = True
        mblnOK = True
    Else
        TransMember = True
        mblnOK = True
        MsgBox "该医生已经被其他用户移除！", vbInformation, gstrSysName
    End If
    Exit Function
    
ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Function

