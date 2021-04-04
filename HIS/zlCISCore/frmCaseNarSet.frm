VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmCaseNarSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "麻醉项目设置"
   ClientHeight    =   4440
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6900
   Icon            =   "frmCaseNarSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(&D)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3330
      TabIndex        =   18
      Top             =   3930
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "新项目(&N)"
      Height          =   350
      Left            =   2205
      TabIndex        =   17
      Top             =   3930
      Visible         =   0   'False
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3525
      Left            =   30
      TabIndex        =   1
      Top             =   225
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   6218
      SortKey         =   3
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "项目名称"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "单位"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "字符"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "记录类型"
         Object.Width           =   1852
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6585
      Top             =   5070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraData 
      Caption         =   "项目数据"
      Height          =   2265
      Left            =   4005
      TabIndex        =   23
      Top             =   135
      Width           =   2835
      Begin MSComCtl2.UpDown udMin 
         Height          =   300
         Left            =   2326
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1380
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   10
         BuddyControl    =   "txtMin"
         BuddyDispid     =   196615
         OrigLeft        =   3330
         OrigTop         =   1350
         OrigRight       =   3570
         OrigBottom      =   1560
         Increment       =   20
         Max             =   300
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMax 
         Height          =   300
         Left            =   2326
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1800
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   10
         BuddyControl    =   "txtMax"
         BuddyDispid     =   196614
         OrigLeft        =   5565
         OrigTop         =   480
         OrigRight       =   5805
         OrigBottom      =   690
         Increment       =   20
         Max             =   300
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cboType 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "frmCaseNarSet.frx":000C
         Left            =   1230
         List            =   "frmCaseNarSet.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1005
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   3
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtMax 
         Height          =   300
         Left            =   1230
         MaxLength       =   12
         TabIndex        =   11
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtMin 
         Height          =   300
         Left            =   1230
         MaxLength       =   12
         TabIndex        =   9
         Top             =   1380
         Width           =   1095
      End
      Begin VB.TextBox txtUnit 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1230
         MaxLength       =   6
         TabIndex        =   5
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "记录类型(&T)"
         Height          =   240
         Left            =   165
         TabIndex        =   6
         Top             =   1050
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "项目名称(&M)"
         Height          =   240
         Left            =   165
         TabIndex        =   2
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "最大值(&A)"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "最小值(&I)"
         Height          =   195
         Left            =   165
         TabIndex        =   8
         Top             =   1455
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "单位(&U)"
         Height          =   240
         Left            =   165
         TabIndex        =   4
         Top             =   690
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   5730
      TabIndex        =   19
      Top             =   3915
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4500
      TabIndex        =   16
      Top             =   3930
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   45
      TabIndex        =   20
      Top             =   3915
      Width           =   1100
   End
   Begin VB.Frame fraDisplay 
      Caption         =   "显示效果"
      Height          =   1155
      Left            =   4005
      TabIndex        =   21
      Top             =   2595
      Width           =   2850
      Begin VB.ComboBox cboChar 
         Height          =   300
         Left            =   1245
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "…"
         Height          =   270
         Left            =   2370
         TabIndex        =   15
         Top             =   720
         Width           =   270
      End
      Begin VB.Label lblColor 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1245
         TabIndex        =   22
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "记录色(&L)"
         Height          =   210
         Left            =   390
         TabIndex        =   14
         Top             =   735
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "记录符(&R)"
         Height          =   225
         Left            =   390
         TabIndex        =   12
         Top             =   345
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1230
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseNarSet.frx":0010
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseNarSet.frx":0468
            Key             =   "NewItem"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "麻醉项目(&G)"
      Height          =   180
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1005
   End
End
Attribute VB_Name = "frmCaseNarSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnOK As Boolean
Private strSQL As String
Private rsTmp As New ADODB.Recordset

Private svrItem As MSComctlLib.ListItem
'标志：第2个列标志为：0以前    1新增
'      第3个列标志为：0无任何变化  1更新   2删除

Private Sub ItemDel(objLvw As ListView)
'标志为删除
Dim i As Long
Dim lngMaxIndex As Long
Dim objList As ListItem

With objLvw
    If .SelectedItem Is Nothing Then Exit Sub
    lngMaxIndex = .SelectedItem.Index
    '设置删除颜色
    For i = 1 To .ColumnHeaders.Count
        If i = 1 Then
            .SelectedItem.ForeColor = RGB(255, 0, 0)
        Else
            .SelectedItem.ListSubItems(i - 1).ForeColor = RGB(255, 0, 0)
        End If
    Next
    If lvw.SelectedItem.ListSubItems(2).Tag = 0 Then
        Me.cmdDelete.Caption = "反删除(&D)"
    Else
        Me.cmdDelete.Caption = "删除(&D)"
    End If
    '设置删除标志
    If .SelectedItem.ListSubItems(2).Tag = 0 Then
        .SelectedItem.ListSubItems(3).Tag = 2
    Else
        .ListItems.Remove lngMaxIndex
    End If
    '设置下一选择项
    On Error Resume Next
    Err.Clear
    Set objList = objLvw.ListItems(lngMaxIndex)
    If Not (objList Is Nothing) Then
        objList.Selected = True
        objList.EnsureVisible
        lvw_ItemClick lvw.SelectedItem
    ElseIf Err <> 0 Then
        Set objList = objLvw.ListItems(lngMaxIndex - 1)
        If Err <> 0 Or Not (objList Is Nothing) Then
            objList.Selected = True
            objList.EnsureVisible
            lvw_ItemClick lvw.SelectedItem
        Else
            Err.Clear
        End If
    End If
End With
End Sub

Private Sub UNItemDel(objLvw As ListView)
'取消删除标志
Dim i As Long
Dim lngMaxIndex As Long
Dim objList As ListItem

With objLvw
    If .SelectedItem Is Nothing Then Exit Sub
    lngMaxIndex = .SelectedItem.Index
    '取消删除标志
    For i = 1 To .ColumnHeaders.Count
        If i = 1 Then
            .SelectedItem.ForeColor = 0
        Else
            .SelectedItem.ListSubItems(i - 1).ForeColor = 0
        End If
    Next
    Me.cmdDelete.Caption = "删除(&D)"
    .SelectedItem.ListSubItems(3).Tag = 1
End With
End Sub

Private Sub WriteTag(ByVal bytTag As Byte)
'写入标志
'标志：第2个列标志为：0以前    1新增
'      第3个列标志为：0无任何变化  1更新   2删除
    If lvw.SelectedItem Is Nothing Then Exit Sub
    lvw.SelectedItem.ListSubItems(3).Tag = bytTag
End Sub

Private Sub cboChar_Change()
    cmdOK.Enabled = True
    If Not (svrItem Is Nothing) Then WriteBack svrItem
End Sub

Private Sub cboChar_Click()
    cmdOK.Enabled = True
    If Not (svrItem Is Nothing) Then WriteBack svrItem
End Sub

Private Sub cboChar_GotFocus()
    zlControl.TxtSelAll cboChar
End Sub

Private Sub cboChar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If Check是否包含(UCase(Chr(KeyAscii)), "'") = True Then
        KeyAscii = 0
    End If
    If InStr(UCase(Chr(KeyAscii)), ";") > 0 Then
        KeyAscii = 0
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub cboChar_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(cboChar.Text, 2)
    If Cancel = False Then
        If InStr(cboChar.Text, ";") > 0 Then
            MsgBox "包含有非法字符！", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub cboType_Click()
    If Not (lvw.SelectedItem Is Nothing) Then CustomEnabled cboType.Text
    cmdOK.Enabled = True
    If Not (svrItem Is Nothing) Then WriteBack svrItem
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    If cmdOK.Enabled = True Then
        If MsgBox("你确认就这样不保存就退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdColor_Click()
    
    With dlg
        .Color = lblColor.BackColor
        .ShowColor
        If lblColor.BackColor <> .Color Then
            lblColor.BackColor = .Color
            cmdOK.Enabled = True
            If Not (svrItem Is Nothing) Then WriteBack svrItem
        End If
    End With
End Sub

Private Sub cmdDelete_Click()
'标志：第2个列标志为：0以前    1新增
'      第3个列标志为：0无任何变化  1更新   2删除
    If lvw.SelectedItem Is Nothing Then Exit Sub
    If lvw.SelectedItem.ListSubItems(3).Tag = 2 Then
        UNItemDel lvw
    Else
        ItemDel lvw
    End If
    If lvw.SelectedItem Is Nothing Then
        cmdDelete.Enabled = False
    Else
        cmdDelete.Enabled = True
        lvw_ItemClick lvw.SelectedItem
    End If
    cmdOK.Enabled = True
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdNew_Click()
    Dim itmx As ListItem
    Dim strType As String
On Error GoTo ErrHandle
    
    If lvw.ListItems.Count >= 12 Then
        MsgBox "曲线项目、标注项目及固定项目总数不能超过12", vbInformation, gstrSysName
        Exit Sub
    End If
    strType = "曲线项目"
    If Not (lvw.SelectedItem Is Nothing) Then
        If lvw.SelectedItem.SubItems(3) <> "固定项目" Then strType = lvw.SelectedItem.SubItems(3)
    End If
    
    
    Set itmx = lvw.ListItems.Add(, , "新项目", "NewItem", "NewItem")
    itmx.SubItems(1) = ""
    itmx.SubItems(2) = ""
    itmx.SubItems(3) = strType
    itmx.SubItems(4) = ""
    itmx.SubItems(5) = ""
    itmx.ListSubItems(1).Tag = "10;300;0"
    '标志：第2个列标志为：0以前    1新增
    '      第3个列标志为：0无任何变化  1更新   2删除
    itmx.ListSubItems(2).Tag = 1
    itmx.ListSubItems(3).Tag = 0
    itmx.Selected = True
    itmx.EnsureVisible
    lvw_ItemClick itmx
    '标志：第2个列标志为：0以前    1新增
    '      第3个列标志为：0无任何变化  1更新   2删除
    itmx.ListSubItems(2).Tag = 1
    itmx.ListSubItems(3).Tag = 0
    cmdDelete.Enabled = True
    cmdOK.Enabled = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, j As Long
    Dim itmx As ListItem
    Dim lng序号 As Long
    Dim strErr As String
    Dim strTmp As String
    
    On Error GoTo ErrHand
    
    strTmp = ""
    For i = 1 To lvw.ListItems.Count
        Set itmx = lvw.ListItems(i)
        If Trim(itmx.Text) = "说明" Or Trim(itmx.Text) = "输氧" Or Trim(itmx.Text) = "麻醉剂" Or Trim(itmx.Text) = "尿液" Or Trim(itmx.Text) = "用药" Or Trim(itmx.Text) = "输液" Then
            strTmp = "第 " & i & " 行的项目名称与系统名称重复请重新命名！"
            lvw.ListItems(i).Selected = True
            lvw.ListItems(i).EnsureVisible
            lvw_ItemClick lvw.ListItems(i)
            If Me.txtName.Enabled And Me.txtName.Visible Then Me.txtName.SetFocus
            Exit For
        End If
        If Trim(itmx.Text) = "" Then
            strTmp = "第 " & i & " 行的项目名称不能为空！"
            lvw.ListItems(i).Selected = True
            lvw.ListItems(i).EnsureVisible
            lvw_ItemClick lvw.ListItems(i)
            If Me.txtName.Enabled And Me.txtName.Visible Then Me.txtName.SetFocus
            Exit For
        End If
        For j = 1 To lvw.ListItems.Count
            If Trim(lvw.ListItems(i).Text) = Trim(lvw.ListItems(j).Text) And i <> j Then
                lvw.ListItems(j).Selected = True
                lvw.ListItems(j).EnsureVisible
                lvw_ItemClick lvw.ListItems(j)
                MsgBox "项目【" & lvw.ListItems(j).Text & "】有重复！", vbOKOnly + vbInformation, gstrSysName
                If Me.txtName.Enabled And Me.txtName.Visible Then Me.txtName.SetFocus
                Exit Sub
            End If
        Next
        If Trim(itmx.SubItems(2)) = "" Then
            strTmp = "项目【" & itmx.Text & "】的记录符不能为空！"
            lvw.ListItems(i).Selected = True
            lvw.ListItems(i).EnsureVisible
            lvw_ItemClick lvw.ListItems(i)
            If Me.cboChar.Enabled And Me.cboChar.Visible Then Me.cboChar.SetFocus
            Exit For
        End If
                        
        If itmx.SubItems(3) = "曲线项目" And Val(Split(itmx.ListSubItems(1).Tag, ";")(1)) <= (Val(Split(itmx.ListSubItems(1).Tag, ";")(0))) Then
            strTmp = "第　" & i & "　行中曲线项目最大值必须大于最小值！"
            lvw.ListItems(i).Selected = True
            lvw.ListItems(i).EnsureVisible
            lvw_ItemClick lvw.ListItems(i)
            If Me.txtMax.Enabled And Me.txtMax.Visible Then Me.txtMax.SetFocus
            Exit For
        End If
    Next
    If strTmp <> "" Then
        MsgBox strTmp, vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    
    For i = 1 To lvw.ListItems.Count
        Set itmx = lvw.ListItems(i)
        If itmx.SubItems(3) = "曲线项目" Or itmx.SubItems(3) = "标注项目" Then
'            If itmx.ListSubItems(2).Tag = 1 Then    '新项目
                '项目名称===单位===字符===记录类型===序号
                '名称_IN、单位_IN、记录法_IN、记录符_IN、记录色_IN、最大值_IN、最小值_IN
                'itmx.ListSubItems(1).Tag = 最小值;最大值;颜色
'                gstrSql = "ZL_麻醉记录项目_INSERT('" & Trim(itmx.Text) & "','" & _
'                        Trim(itmx.SubItems(1)) & "'," & IIf(itmx.SubItems(3) = "曲线项目", 1, 2) & ",'" & _
'                        Trim(itmx.SubItems(2)) & "'," & Split(itmx.ListSubItems(1).Tag, ";")(2) & "," & _
'                        Split(itmx.ListSubItems(1).Tag, ";")(1) & "," & Split(itmx.ListSubItems(1).Tag, ";")(0) & ")"
'
'                Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
'            Else
            If itmx.ListSubItems(2).Tag = 0 Then    '旧项目
                If itmx.ListSubItems(3).Tag = 1 Then
                    '控件号_IN       体麻项记录法.序号%TYPE,
                    '名称_IN         体麻项记录法.记录名%TYPE,
                    '记录符_IN       体麻项记录法.记录符%TYPE,
                    '记录色_IN       体麻项记录法.记录色%TYPE,
                    '最大值_IN       体麻项记录法.最大值%TYPE,
                    '最小值_IN       体麻项记录法.最小值%TYPE
                    '项目名称,1500,0,1;单位,900,0,2;字符,600,0,2;记录类型,1200,0,2
                    'itmx.ListSubItems(1).Tag = 最小值;最大值;颜色
                    gstrSql = "ZL_麻醉记录项目_UPDATE(" & itmx.Tag & ",'" & Trim(itmx.Text) & "','" & _
                            Trim(itmx.SubItems(2)) & "'," & Split(itmx.ListSubItems(1).Tag, ";")(2) & "," & _
                            Split(itmx.ListSubItems(1).Tag, ";")(1) & "," & Split(itmx.ListSubItems(1).Tag, ";")(0) & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
'                ElseIf itmx.ListSubItems(3).Tag = 2 Then
'                     '项目ID_IN、控件号_IN
'                    gstrSql = "ZL_麻醉记录项目_DELETE(" & itmx.Tag & "," & Trim(itmx.SubItems(4)) & ")"
'                    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
                End If
            End If
'            End If
        End If
    Next
    
    gcnOracle.CommitTrans
    blnOK = True
    cmdOK.Enabled = False
    Unload Me
    Exit Sub
ErrHand:
    strErr = Err.Description
    If InStr(1, strErr, "[ZLSOFT]") > 0 Then
        On Error Resume Next
        strErr = Split(strErr, "[ZLSOFT]")(1)
    End If
    gcnOracle.RollbackTrans
    Call SaveErrLog
    If strErr <> "" Then MsgBox strErr, vbExclamation, gstrSysName
End Sub

Private Sub Form_Load()
    blnOK = False
    Set svrItem = Nothing
    zlControl.LvwSelectColumns lvw, "项目名称,1500,0,1;单位,900,0,2;字符,600,0,2;记录类型,1200,0,2;序号,0,0,2;空,0,0,2", True
    zlControl.LvwFlatColumnHeader lvw
    cboType.AddItem "曲线项目"
    cboType.AddItem "标注项目"
    lvw.Sorted = False
    
    Init
    If lvw.ListItems.Count > 0 Then
        lvw.ListItems(1).Selected = True
        cmdDelete.Enabled = True
    End If
    lvw_ItemClick lvw.SelectedItem
    cmdOK.Enabled = False
End Sub

Private Sub Init()
    Dim i As Long
    Dim itmx As ListItem
    Dim rsTmp As New ADODB.Recordset
On Error GoTo ErrHandle
    
    With rsTmp
        lvw.ListItems.Clear
        '求出固定项目
        strSQL = _
            "SELECT nvl(a.id,0) 项目id,b.序号 项目号, " & vbCrLf & _
            "    b.记录名 项目名,nvl(a.单位,'') 单位, " & vbCrLf & _
            "    b.最大值,b.最小值,b.记录符,b.记录色 " & vbCrLf & _
            " FROM 诊治所见项目 a,体麻项记录法 b " & vbCrLf & _
            " WHERE b.项目id=a.id(+) AND b.类型=2 AND b.记录法 is null and b.记录名 in ('麻  醉','手  术')" & vbCrLf & _
            " ORDER BY b.序号"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If .BOF = False Then
            While Not .EOF
                Set itmx = lvw.ListItems.Add(, "K" & rsTmp!项目号, rsTmp!项目名, 1, 1)
                itmx.Tag = rsTmp!项目号
                itmx.SubItems(1) = IIf(IsNull(rsTmp!单位), "", rsTmp!单位)
                itmx.SubItems(2) = IIf(IsNull(rsTmp!记录符), "", rsTmp!记录符)
                itmx.SubItems(3) = "固定项目"
                itmx.SubItems(4) = CStr(rsTmp!项目号)
                itmx.SubItems(5) = ""
                itmx.ListSubItems(1).Tag = IIf(IsNull(rsTmp!最小值), "10", rsTmp!最小值) & ";" & IIf(IsNull(rsTmp!最大值), "300", rsTmp!最大值) & ";" & IIf(IsNull(rsTmp!记录色), 0, rsTmp!记录色)
                '标志：第2个列标志为：0以前    1新增
                '      第3个列标志为：0无任何变化  1更新   2删除
                itmx.ListSubItems(2).Tag = 0
                itmx.ListSubItems(3).Tag = 0
                .MoveNext
            Wend
        End If
        '曲线数据
        strSQL = _
            "SELECT nvl(a.id,0) 项目id,b.序号 项目号, " & vbCrLf & _
            "    b.记录名 项目名,nvl(a.单位,'') 单位, " & vbCrLf & _
            "    b.最大值,b.最小值,b.记录符,b.记录色 " & vbCrLf & _
            " FROM 诊治所见项目 a,体麻项记录法 b " & vbCrLf & _
            " WHERE b.项目id=a.id(+) AND b.类型=2 and  b.记录法 is  null and not b.记录名 in ('麻  醉','手  术')" & vbCrLf & _
            " ORDER BY b.序号"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If .BOF = False Then
            While Not .EOF
                Set itmx = lvw.ListItems.Add(, "K" & rsTmp!项目号, rsTmp!项目名, 1, 1)
                itmx.Tag = rsTmp!项目号
                itmx.SubItems(1) = IIf(IsNull(rsTmp!单位), "", rsTmp!单位)
                itmx.SubItems(2) = IIf(IsNull(rsTmp!记录符), "", rsTmp!记录符)
                itmx.SubItems(3) = "曲线项目"
                itmx.SubItems(4) = CStr(rsTmp!项目号)
                itmx.SubItems(5) = ""
                itmx.ListSubItems(1).Tag = IIf(IsNull(rsTmp!最小值), "10", rsTmp!最小值) & ";" & IIf(IsNull(rsTmp!最大值), "300", rsTmp!最大值) & ";" & IIf(IsNull(rsTmp!记录色), 0, rsTmp!记录色)
                '标志：第2个列标志为：0以前    1新增
                '      第3个列标志为：0无任何变化  1更新   2删除
                itmx.ListSubItems(2).Tag = 0
                itmx.ListSubItems(3).Tag = 0
                .MoveNext
            Wend
        End If
        
        '没有非曲线项目
'        '非曲线数据
''        strSQL = _
''            "SELECT a.id 项目id,a.小数 项目号," & vbCrLf & _
''            "   a.中文名 项目名,a.单位," & vbCrLf & _
''            "    b.最大值,b.最小值,b.记录符,b.记录色" & vbCrLf & _
''            "FROM 诊治所见项目 a,体麻项记录法 b" & vbCrLf & _
''            "WHERE b.项目id=a.id AND b.类型=2 AND b.记录法=2" & vbCrLf & _
''            "ORDER BY a.小数"
'        strSQL = _
'            "SELECT nvl(a.id,0) 项目id,b.序号 项目号, " & vbCrLf & _
'            "    b.记录名 项目名,nvl(a.单位,'') 单位, " & vbCrLf & _
'            "    b.最大值,b.最小值,b.记录符,b.记录色 " & vbCrLf & _
'            " FROM 诊治所见项目 a,体麻项记录法 b " & vbCrLf & _
'            " WHERE b.项目id=a.id(+) AND b.类型=2 and  b.记录法 is  null and not b.记录名 in ('麻  醉','手  术')" & vbCrLf & _
'            " ORDER BY b.序号"
'        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
'        If .BOF = False Then
'            While Not .EOF
'                Set itmx = lvw.ListItems.Add(, "K" & rsTmp!项目号, rsTmp!项目名, 1, 1)
'                itmx.Tag = rsTmp!项目ID
'                itmx.SubItems(1) = IIf(IsNull(rsTmp!单位), "", rsTmp!单位)
'                itmx.SubItems(2) = IIf(IsNull(rsTmp!记录符), "", rsTmp!记录符)
'                itmx.SubItems(3) = "标注项目"
'                itmx.SubItems(4) = CStr(rsTmp!项目号)
'                itmx.SubItems(5) = ""
'                itmx.ListSubItems(1).Tag = IIf(IsNull(rsTmp!最小值), "10", rsTmp!最小值) & ";" & IIf(IsNull(rsTmp!最大值), "300", rsTmp!最大值) & ";" & IIf(IsNull(rsTmp!记录色), 0, rsTmp!记录色)
'                '标志：第2个列标志为：0以前    1新增
'                '      第3个列标志为：0无任何变化  1更新   2删除
'                itmx.ListSubItems(2).Tag = 0
'                itmx.ListSubItems(3).Tag = 0
'                .MoveNext
'            Wend
'        End If
        
        For i = 65 To 90
            cboChar.AddItem Chr(i)
        Next
        cboChar.AddItem "⊕"
        cboChar.AddItem "☉"
        cboChar.AddItem "+"
        cboChar.AddItem "*"
        cboChar.AddItem "ο"
        cboChar.Text = "A"
        If cboType.ListCount > 0 Then cboType.ListIndex = 0
        Me.txtMax.Text = "20"
        Me.txtMin.Text = "0"
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RefreshItem Item
    If Item Is Nothing Then
        CustomEnabled "标注项目"
    Else
        CustomEnabled Item.SubItems(3)
    End If
End Sub

Private Sub RefreshItem(ByVal Item As MSComctlLib.ListItem)
    Dim svrOK As Boolean
On Error GoTo ErrHandle
    
    If Item Is Nothing Then Exit Sub
    Set svrItem = Nothing
    svrOK = cmdOK.Enabled
    cboType.Clear
    txtName.Text = Item.Text
    txtUnit.Text = Item.SubItems(1)
    cboChar.Text = Item.SubItems(2)
    If Item.SubItems(3) = "固定项目" Then
        cboType.AddItem "固定项目"
        Me.cmdDelete.Caption = "删除(&D)"
    Else
        cboType.AddItem "曲线项目"
        cboType.AddItem "标注项目"
        If Item.ListSubItems(3).Tag = 2 Then
            Me.cmdDelete.Caption = "反删除(&D)"
        Else
            Me.cmdDelete.Caption = "删除(&D)"
        End If
    End If
    cboType.Text = Item.SubItems(3)
    
    txtMin.Text = Split(Item.ListSubItems(1).Tag, ";")(0)
    txtMax.Text = Split(Item.ListSubItems(1).Tag, ";")(1)
    udMin.Value = IIf(Val(txtMin.Text) < 0, 0, Val(txtMin.Text))
    udMax.Value = IIf(Val(txtMax.Text) < 0, 0, Val(txtMax.Text))
    lblColor.BackColor = Val(Split(Item.ListSubItems(1).Tag, ";")(2))
    Set svrItem = Item
    cmdOK.Enabled = svrOK
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub WriteBack(ByVal Item As MSComctlLib.ListItem)
    Item.Text = txtName.Text
    Item.SubItems(1) = txtUnit.Text
    Item.SubItems(2) = cboChar.Text
    Item.SubItems(3) = cboType.Text
    Item.ListSubItems(1).Tag = txtMin.Text & ";" & txtMax.Text & ";" & lblColor.BackColor
    Call WriteTag(1)
End Sub

Private Sub CustomEnabled(ByVal strFlag As String)
    If lvw.SelectedItem Is Nothing Then
        fraData.Enabled = False
            '项目名称
            Me.Label7.Enabled = False
            '单位
            Me.Label1.Enabled = False
            '记录类型
            Me.Label8.Enabled = False
            '最大值
            Me.Label5.Enabled = False
            '最小值
            Me.Label4.Enabled = False
        fraDisplay.Enabled = False
            '记录符
            Me.Label2.Enabled = False
            '记录色
            Me.Label3.Enabled = False
        txtName.Enabled = False
        txtUnit.Enabled = False
        txtMin.Enabled = False
        txtMax.Enabled = False
        cboType.Enabled = False
        udMax.Enabled = False
        udMin.Enabled = False
        cmdColor.Enabled = False
        cboChar.Enabled = False
        cmdDelete.Enabled = False
        Exit Sub
    End If
    fraData.Enabled = True
        '项目名称
        Me.Label7.Enabled = True
        '单位
        Me.Label1.Enabled = False
        '记录类型
        Me.Label8.Enabled = False
        '最大值
        Me.Label5.Enabled = True
        '最小值
        Me.Label4.Enabled = True
    fraDisplay.Enabled = True
        '记录符
        Me.Label2.Enabled = True
        '记录色
        Me.Label3.Enabled = True
    txtName.Enabled = True
    txtUnit.Enabled = False
    txtMin.Enabled = True
    txtMax.Enabled = True
    '记录类型
    cboType.Enabled = False
    udMax.Enabled = True
    udMin.Enabled = True
    cmdColor.Enabled = True
    cboChar.Enabled = True
    cmdDelete.Enabled = False
    
    If lvw.SelectedItem.ListSubItems(3).Tag = 2 Then
        fraData.Enabled = False
            '项目名称
            Me.Label7.Enabled = False
            '单位
            Me.Label1.Enabled = False
            '记录类型
            Me.Label8.Enabled = False
            '最大值
            Me.Label5.Enabled = False
            '最小值
            Me.Label4.Enabled = False
        fraDisplay.Enabled = False
            '记录符
            Me.Label2.Enabled = False
            '记录色
            Me.Label3.Enabled = False
        txtName.Enabled = False
        txtUnit.Enabled = False
        txtMin.Enabled = False
        txtMax.Enabled = False
        cboType.Enabled = False
        udMax.Enabled = False
        udMin.Enabled = False
        cmdColor.Enabled = False
        cboChar.Enabled = False
    Else
        If strFlag = "固定项目" Then
            fraData.Enabled = False
                '项目名称
                Me.Label7.Enabled = False
                '单位
                Me.Label1.Enabled = False
                '记录类型
                Me.Label8.Enabled = False
                '最大值
                Me.Label5.Enabled = False
                '最小值
                Me.Label4.Enabled = False
                txtName.Enabled = False
                txtUnit.Enabled = False
                txtMin.Enabled = False
                txtMax.Enabled = False
                cboType.Enabled = False
                udMax.Enabled = False
                udMin.Enabled = False
                cmdDelete.Enabled = False
        ElseIf strFlag = "标注项目" Then
            txtUnit.Enabled = False
            txtMin.Enabled = False
            txtMax.Enabled = False
            udMax.Enabled = False
            udMin.Enabled = False
            '最大值
            Me.Label5.Enabled = False
            '最小值
            Me.Label4.Enabled = False
        End If
    End If
End Sub

Private Sub txtMax_Change()
    cmdOK.Enabled = True
    If Not (svrItem Is Nothing) Then WriteBack svrItem
    If Val(txtMax.Text) > 300 Then txtMax.Text = "300": txtMax.SelStart = Len(txtMax.Text)
    udMin.Max = Val(txtMax.Text)
End Sub

Private Sub txtMax_GotFocus()
    zlControl.TxtSelAll txtMax
End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If Check是否包含(UCase(Chr(KeyAscii)), "正整数") = True Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMax_Validate(Cancel As Boolean)
On Error GoTo ErrHandle
    If IsNumeric(txtMax.Text) = False Then
        MsgBox "请输入正确数字！", vbInformation, gstrSysName
        Cancel = True
    End If
    If Val(txtMin.Text) >= Val(txtMax.Text) Then
        MsgBox "输入值无效，最小值应小于最大值！", vbInformation, gstrSysName
        Cancel = True
    End If
    If Val(txtMax.Text) > 300 Or Val(txtMax.Text) < 20 Then
        MsgBox "输入值无效，只能输入20到300之间的数！", vbInformation, gstrSysName
        Cancel = True
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtMin_Change()
    cmdOK.Enabled = True
    If Not (svrItem Is Nothing) Then WriteBack svrItem
    If Val(txtMin.Text) > 300 Then txtMin.Text = "300": txtMin.SelStart = Len(txtMin.Text)
    udMax.Min = Format(txtMin.Text)
End Sub

Private Sub txtMin_GotFocus()
    zlControl.TxtSelAll txtMin
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If Check是否包含(UCase(Chr(KeyAscii)), "正整数") = True Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMin_Validate(Cancel As Boolean)
    If IsNumeric(txtMin.Text) = False Then
        MsgBox "请输入正确数字！", vbInformation, gstrSysName
        Cancel = True
    End If
    If Val(txtMin.Text) >= Val(txtMax.Text) Then
        MsgBox "输入值无效，最小值应小于最大值！", vbInformation, gstrSysName
        Cancel = True
    End If
    If Val(txtMin.Text) > 300 Then
        MsgBox "输入值无效，只能输入0到300之间的数！", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub txtName_Change()
    cmdOK.Enabled = True
    If Not (svrItem Is Nothing) Then WriteBack svrItem
End Sub

Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
    zlCommFun.OpenIme True
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If Check是否包含(UCase(Chr(KeyAscii)), "'") = True Then
        KeyAscii = 0
    End If
    If Check是否包含(UCase(Chr(KeyAscii)), ";") = True Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txtName.Text, txtName.MaxLength)
    If Cancel = False Then
        If InStr(txtName.Text, ";") > 0 Then
            MsgBox "包含有非法字符！", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub txtUnit_Change()
    cmdOK.Enabled = True
    If Not (svrItem Is Nothing) Then WriteBack svrItem
End Sub

Private Sub txtUnit_GotFocus()
    zlControl.TxtSelAll txtUnit
End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If Check是否包含(UCase(Chr(KeyAscii)), "'") = True Then
        KeyAscii = 0
    End If
    If Check是否包含(UCase(Chr(KeyAscii)), ";") = True Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtUnit_Validate(Cancel As Boolean)
On Error GoTo ErrHandle
    Cancel = Not StrIsValid(txtUnit.Text, txtUnit.MaxLength)
    If Cancel = False Then
        If InStr(txtUnit.Text, ";") > 0 Then
            MsgBox "包含有非法字符！", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

