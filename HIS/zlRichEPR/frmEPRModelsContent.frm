VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRModelsContent 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picContent 
      Appearance      =   0  'Flat
      BackColor       =   &H00E7CFBA&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4290
      Left            =   0
      ScaleHeight     =   4290
      ScaleWidth      =   5700
      TabIndex        =   1
      Top             =   2865
      Width           =   5700
      Begin VB.CheckBox chklevel 
         BackColor       =   &H00E7CFBA&
         Caption         =   "全院通用"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   525
         Width           =   1035
      End
      Begin VB.CheckBox chklevel 
         BackColor       =   &H00E7CFBA&
         Caption         =   "科室通用"
         Height          =   225
         Index           =   1
         Left            =   1275
         TabIndex        =   8
         Top             =   525
         Width           =   1035
      End
      Begin VB.CheckBox chklevel 
         BackColor       =   &H00E7CFBA&
         Caption         =   "个人使用"
         Height          =   225
         Index           =   2
         Left            =   2370
         TabIndex        =   7
         Top             =   525
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CommandButton cmdContent 
         Caption         =   "∧ 添加到范文包中(&A)"
         Height          =   350
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   30
         Width           =   2055
      End
      Begin VB.TextBox txtSeek 
         Height          =   270
         Left            =   4305
         TabIndex        =   4
         ToolTipText     =   "输入后回车，以名称查找；或输入简码定位。"
         Top             =   495
         Width           =   1170
      End
      Begin VB.CommandButton cmdContent 
         Caption         =   "∨ 从范文包中删除(&D)"
         Height          =   350
         Index           =   1
         Left            =   3420
         TabIndex        =   3
         Top             =   45
         Width           =   2055
      End
      Begin MSComctlLib.ListView lvwModel 
         Height          =   3420
         Left            =   0
         TabIndex        =   2
         Top             =   825
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   6033
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E7CFBA&
         Caption         =   "名称过滤"
         Height          =   165
         Left            =   3525
         TabIndex        =   5
         Top             =   540
         Width           =   780
      End
   End
   Begin MSComctlLib.ListView lvwModelContent 
      Height          =   2805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   4948
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmEPRModelsContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlngModelsID As Long
Private Sub Initlvw()
    With lvwModelContent.ColumnHeaders
        .Clear
        .Add , "_ID", "", 300
        .Add , "_编号", "编号", 800
        .Add , "_名称", "名称", 2000
        .Add , "_种类", "种类", 1600
        .Add , "_通用级", "通用级", 1000
        .Add , "_说明", "说明", 1800
        .Add , "_简码", "简码", 600
        .Add , "_科室ID", "科室ID", 0
        .Add , "_人员ID", "人员ID", 0
        .Add , "_科室", "科室", 800
        .Add , "_人员", "人员", 800
        .Add , "_文件ID", "文件ID", 0
    End With
    
    With lvwModel.ColumnHeaders
        .Clear
        .Add , "_ID", "", 300
        .Add , "_编号", "编号", 800
        .Add , "_名称", "名称", 2000
        .Add , "_种类", "种类", 1600
        .Add , "_通用级", "通用级", 1000
        .Add , "_说明", "说明", 1800
        .Add , "_简码", "简码", 600
        .Add , "_科室ID", "科室ID", 0
        .Add , "_人员ID", "人员ID", 0
        .Add , "_科室", "科室", 800
        .Add , "_人员", "人员", 800
    End With
End Sub
Private Sub cmdContent_Click(Index As Integer)
Dim arrSQL() As Variant, blnTran As Boolean, l As Integer, strTypes As String
    On Error GoTo ErrHandle
    arrSQL = Array()
    If Index = 0 Then '增加
        For l = 1 To lvwModelContent.ListItems.Count '取出已选病历文件类别
            strTypes = strTypes & lvwModelContent.ListItems(l).SubItems(3) & "|"
        Next
    
        For l = 1 To lvwModel.ListItems.Count
            If lvwModel.ListItems(l).Checked Then
                If InStr(strTypes, lvwModel.ListItems(l).SubItems(3) & "|") > 0 Then '同一种病历只能加入一份,比如出现多个"入院记录"
                    MsgBox "        同一种病历只能加入一份，请检查：" & vbCrLf & "是否选中了相同种类或已经加入了选中种类的病历文件！", vbInformation, gstrSysName: Exit Sub
                End If
                strTypes = strTypes & lvwModel.ListItems(l).SubItems(3) & "|"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_病历范文包组成_Update(1," & mlngModelsID & "," & lvwModel.ListItems(l).Tag & ")"
            End If
        Next
    Else              '删除
        For l = 1 To lvwModelContent.ListItems.Count
            If lvwModelContent.ListItems(l).Checked Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_病历范文包组成_Update(0," & mlngModelsID & "," & lvwModelContent.ListItems(l).Tag & ")"
            End If
        Next
    End If
    
    gcnOracle.BeginTrans '--------------------------写入数据
    blnTran = True
    For l = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(l)), "写入中心病种数据")
    Next
    gcnOracle.CommitTrans: blnTran = False
    
    Call RefreshContent
    Call RefreshModel
    Exit Sub
ErrHandle:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub Form_Load()
    Initlvw
End Sub
Public Sub zlRefresh(ByVal lngModelsID As Long, ByVal strPrivs As String, ByVal bytType As Byte)
'lngModelsID－－－当前范文包ID bytType-0 查询 bytType-1更改范文包组成

    mstrPrivs = strPrivs: mlngModelsID = lngModelsID
    If picContent.Enabled Then bytType = 1
    If InStr(mstrPrivs, "病历范文包管理") > 0 And bytType = 1 Then
        picContent.Enabled = True
        Call Form_Resize
        If InStr(mstrPrivs, "个人病历范文") <= 0 Then chklevel(2).Enabled = False: chklevel(2).Value = False
        If InStr(mstrPrivs, "科室病历范文") <= 0 Then chklevel(1).Enabled = False: chklevel(1).Value = False
        If InStr(mstrPrivs, "全院病历范文") <= 0 Then chklevel(0).Enabled = False: chklevel(0).Value = False
        Call RefreshModel
    Else
        picContent.Enabled = False
        Call Form_Resize
    End If
    Call RefreshContent
End Sub
Private Sub RefreshContent()
Dim rsTemp As ADODB.Recordset, objItem As ListItem
    On Error GoTo ErrHandle
    gstrSQL = "select /*+ rule*/ A.ID,A.编号,A.名称,A.简码,A.说明,A.通用级,A.科室ID,A.人员ID ,C.名称 类别,D.名称 科室,E.姓名,A.文件ID" & _
                " from 病历范文目录 A,病历范文包组成 B ,病历文件列表 C,部门表 D,人员表 E" & _
                " where B.范文包ID=[1] AND A.ID=B.范文ID And nvl(A.性质,0)=0 AND A.文件ID=C.ID AND C.种类=2 AND A.科室ID=D.ID AND A.人员ID=E.ID" & _
                " Order by C.种类,C.编号,A.通用级,A.编号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngModelsID)
    lvwModelContent.ListItems.Clear
    With rsTemp
        Do Until .EOF
            Set objItem = lvwModelContent.ListItems.Add(, "_" & !ID, "")
                objItem.Tag = !ID
                objItem.SubItems(1) = !编号
                objItem.SubItems(2) = !名称
                objItem.SubItems(3) = NVL(!类别)
                objItem.SubItems(4) = Decode(NVL(!通用级, 0), 0, "全院通用", 1, "科室通用", 2, "个人使用")
                objItem.SubItems(5) = NVL(!说明)
                objItem.SubItems(6) = NVL(!简码)
                objItem.SubItems(7) = NVL(!科室ID, 0)
                objItem.SubItems(8) = NVL(!人员ID, 0)
                objItem.SubItems(9) = NVL(!科室)
                objItem.SubItems(10) = NVL(!姓名)
                objItem.SubItems(11) = NVL(!文件ID, 0)
                objItem.Checked = True
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub RefreshModel()
Dim rsTemp As ADODB.Recordset, objItem As ListItem, lngID As Long, i As Integer, debarID As String
    On Error GoTo ErrHandle
    If lvwModel.ListItems.Count > 0 Then '记下当前所选
        lngID = lvwModel.SelectedItem.Tag
    End If
    
'    For i = 1 To lvwModelContent.ListItems.Count
'        debarID = debarID & "," & lvwModelContent.ListItems(i).Tag
'    Next
'    If debarID <> "" Then debarID = Mid(debarID, 2)
    
    gstrSQL = ""
    If chklevel(0).Value = vbChecked Then gstrSQL = "A.通用级=0" '全院通用
    If chklevel(1).Value = vbChecked Then gstrSQL = gstrSQL & IIf(gstrSQL = "", "", " or ") & "(A.通用级=1 and A.科室ID=[1])" '科室通用
    If chklevel(2).Value = vbChecked Then gstrSQL = gstrSQL & IIf(gstrSQL = "", "", " or ") & "(A.通用级=2 and A.人员ID=[2])" '个人使用
    If chklevel(0).Value = vbChecked And chklevel(1).Value = vbChecked And chklevel(2).Value = vbChecked Then gstrSQL = "" '全选
        
    If gstrSQL = "" Then '全选时跟据权限加条件
        If chklevel(0).Enabled Then gstrSQL = "A.通用级=0"
        If chklevel(1).Enabled Then gstrSQL = gstrSQL & IIf(gstrSQL = "", "", " OR ") & "(A.通用级=1 and A.科室ID=[1])"
        If chklevel(2).Enabled Then gstrSQL = gstrSQL & IIf(gstrSQL = "", "", " OR ") & "(A.通用级=2 and A.人员ID=[2])"
    End If
    
    gstrSQL = "select /*+ rule*/ A.ID,A.编号,A.名称,A.简码,A.说明,A.通用级,A.科室ID,A.人员ID,B.名称 类别,C.名称 科室,D.姓名 " & _
                " from 病历范文目录 A,病历文件列表 B ,部门表 C,人员表 D" & _
                " where A.文件ID=B.ID AND B.种类=2 and A.科室ID=C.ID and A.人员ID=D.ID AND nvl(A.性质,0)=0" & IIf(gstrSQL = "", "", " and (" & gstrSQL & ")")
    If Trim(txtSeek.Text) <> "" Then
        gstrSQL = gstrSQL & " And " & zlCommFun.GetLike("A", "名称", Trim(txtSeek))
    End If
'    If debarID <> "" Then
'        gstrSQL = gstrSQL & " And A.文件ID not in(Select Distinct 文件ID from 病历范文目录 where ID IN (" & debarID & "))"
'    End If
    gstrSQL = gstrSQL & " Order by A.通用级,A.简码,B.种类,A.编号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngDeptId, glngUserId)
    lvwModel.ListItems.Clear
    With rsTemp
        Do Until .EOF
            Set objItem = lvwModel.ListItems.Add(, "_" & !ID, "")
                objItem.Tag = !ID
                objItem.SubItems(1) = !编号
                objItem.SubItems(2) = !名称
                objItem.SubItems(3) = NVL(!类别)
                objItem.SubItems(4) = Decode(NVL(!通用级, 0), 0, "全院通用", 1, "科室通用", 2, "个人使用")
                objItem.SubItems(5) = NVL(!说明)
                objItem.SubItems(6) = NVL(!简码)
                objItem.SubItems(7) = NVL(!科室ID, 0)
                objItem.SubItems(8) = NVL(!人员ID, 0)
                objItem.SubItems(9) = NVL(!科室)
                objItem.SubItems(10) = NVL(!姓名)
            If !ID = lngID Then
                objItem.Selected = True
            End If
            .MoveNext
        Loop
    End With
    If lvwModel.ListItems.Count > 0 Then
        If lvwModel.SelectedItem Is Nothing Then
            lvwModel.ListItems(1).Selected = True
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    lvwModelContent.Width = Me.Width
    picContent.Width = Me.Width
    lvwModel.Width = Me.Width
    If picContent.Enabled Then
        picContent.Visible = True
        lvwModelContent.Height = 2800
        picContent.Height = Me.ScaleHeight - lvwModelContent.Height
        lvwModel.Height = picContent.Height - (chklevel(0).Top + chklevel(0).Height)
        picContent.Top = lvwModelContent.Height + lvwModelContent.Top
    Else
        picContent.Visible = False
        lvwModelContent.Top = 0
        lvwModelContent.Height = Me.ScaleHeight
    End If
    Err = 0: Err.Clear
End Sub

Private Sub chklevel_Click(Index As Integer)
Dim i As Integer, blnOnly As Boolean
    For i = 0 To chklevel.UBound
        If chklevel(i).Enabled Then
            If chklevel(i).Value = vbChecked Then
                blnOnly = True: Exit For '只要有被选中即退出
            End If
        End If
    Next
    
    If blnOnly = False Then chklevel(Index).Value = vbChecked '保证始终有一个被选中
    Call RefreshModel
End Sub

Private Sub lvwModel_Click()
Dim i As Integer
    For i = 1 To lvwModel.ListItems.Count
        If lvwModel.ListItems(i).Checked = True Then cmdContent(0).Enabled = True: Exit Sub
    Next
    cmdContent(0).Enabled = False
End Sub

Private Sub lvwModelContent_Click()
Dim i As Integer
    For i = 1 To lvwModelContent.ListItems.Count
        If lvwModelContent.ListItems(i).Checked = True Then cmdContent(1).Enabled = True: Exit Sub
    Next
    cmdContent(1).Enabled = False
End Sub


Private Sub txtSeek_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call RefreshModel
    ElseIf InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then '简码定位
        Dim i As Integer, strtmp As String
        If txtSeek.SelLength > 0 Then
            strtmp = ""
        Else
            strtmp = txtSeek.Text
        End If
        For i = 1 To lvwModel.ListItems.Count
            If UCase(lvwModel.ListItems(i).SubItems(6)) Like UCase(Trim(strtmp)) & UCase(Chr(KeyAscii)) & "*" Then
                lvwModel.SelectedItem.Selected = False: lvwModel.ListItems(i).Selected = True: Exit Sub
            End If
        Next
    End If
End Sub
