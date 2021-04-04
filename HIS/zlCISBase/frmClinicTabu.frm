VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmClinicTabu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "诊疗排斥关系"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   Icon            =   "frmClinicTabu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraLine 
      Height          =   60
      Index           =   0
      Left            =   -90
      TabIndex        =   18
      Top             =   720
      Width           =   9345
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Index           =   1
      Left            =   -240
      TabIndex        =   17
      Top             =   4560
      Width           =   9345
   End
   Begin VB.OptionButton optType 
      Caption         =   "停止互斥长嘱(&3)"
      Height          =   210
      Index           =   2
      Left            =   5745
      TabIndex        =   9
      Top             =   1425
      Width           =   1740
   End
   Begin VB.OptionButton optType 
      Caption         =   "禁止(&2)"
      Height          =   210
      Index           =   1
      Left            =   4687
      TabIndex        =   8
      Top             =   1425
      Width           =   990
   End
   Begin VB.OptionButton optType 
      Caption         =   "提醒(&1)"
      Height          =   210
      Index           =   0
      Left            =   3630
      TabIndex        =   7
      Top             =   1425
      Width           =   990
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增(&A)"
      Height          =   350
      Left            =   225
      Picture         =   "frmClinicTabu.frx":058A
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   975
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   1365
      Picture         =   "frmClinicTabu.frx":06D4
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   975
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   3630
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1020
      Width           =   3750
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   3720
      TabIndex        =   12
      Top             =   1715
      Width           =   1095
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复(&R)"
      Height          =   350
      Left            =   4920
      TabIndex        =   13
      Top             =   1715
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwGroup 
      Height          =   3090
      Left            =   105
      TabIndex        =   1
      Top             =   1380
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   5450
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2595
      Left            =   240
      TabIndex        =   16
      Top             =   5280
      Visible         =   0   'False
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   4890
      Top             =   5280
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
            Picture         =   "frmClinicTabu.frx":081E
            Key             =   "Group"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicTabu.frx":0DB8
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      Picture         =   "frmClinicTabu.frx":1352
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1155
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      Height          =   350
      Left            =   6120
      TabIndex        =   14
      Top             =   4680
      Width           =   1155
   End
   Begin ZL9BillEdit.BillEdit msfTabu 
      Height          =   2265
      Left            =   2670
      TabIndex        =   11
      Top             =   2205
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3995
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "排斥类型"
      Height          =   180
      Left            =   2670
      TabIndex        =   6
      Top             =   1425
      Width           =   720
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "组名称(&N)"
      Height          =   180
      Left            =   2670
      TabIndex        =   4
      Top             =   1080
      Width           =   810
   End
   Begin VB.Label lblTabu 
      AutoSize        =   -1  'True
      Caption         =   "排斥项(&T)"
      Height          =   180
      Left            =   2670
      TabIndex        =   10
      Top             =   1800
      Width           =   810
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    部分诊疗项目之间，存在互相排斥的关系；根据项目的应用特性，恰当地设置定义这些排斥的组合，能使医嘱在下达和执行过程中更加方便。"
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   870
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   225
      Picture         =   "frmClinicTabu.frx":149C
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmClinicTabu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------
'说明：
'   1、编辑状态：由Me.tag存放，是否为编辑，由上级程序通过ShowMe传入
'---------------------------------------------------
Private strInputed As String
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim strTemp As String, aryTemp() As String
Dim intCount As Integer, intFence As Integer

Public Sub ShowMe(ByVal frmParent As Object, ByVal bln编辑 As Boolean)
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
    Me.Tag = IIf(bln编辑, "编辑", "查阅")
    If Me.Tag = "查阅" Then
        Me.cmdAdd.Enabled = False: Me.cmdDel.Enabled = False
        Me.txtName.Enabled = False
        Me.optType(0).Enabled = False: Me.optType(1).Enabled = False: Me.optType(2).Enabled = False
        Me.msfTabu.Active = False
        Me.cmdSave.Enabled = False: Me.cmdRestore.Enabled = False
    End If
    Me.lvwGroup.ListItems.Clear
    With Me.lvwGroup.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2000
    End With
    Call zlGroupRef
    Me.Show 1, frmParent
End Sub


Private Sub cmdAdd_Click()
    Me.lblName.Tag = "": Me.txtName.Text = ""
    Me.optType(0).Value = False: Me.optType(1).Value = False: Me.optType(2).Value = False
    Me.msfTabu.ClearBill

    Me.cmdDel.Enabled = True
    Me.txtName.Enabled = True
    Me.optType(0).Enabled = True: Me.optType(1).Enabled = True: Me.optType(2).Enabled = True
    Me.msfTabu.Active = True
    Me.cmdSave.Enabled = True: Me.cmdRestore.Enabled = True
    
    Me.txtName.SetFocus

End Sub

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdDel_Click()
    If MsgBox("真的删除排斥组“" & Trim(Me.txtName.Text) & "”吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gstrSql = "zl_诊疗互斥项目_Save(" & Val(Me.lblName.Tag) & ",'" & Trim(Me.txtName.Text) & "',0,null)"
    Err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    Call zlGroupRef
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdRestore_Click()
    Call zlGroupRef(Val(Me.lblName.Tag))
End Sub

Private Sub cmdSave_Click()
    If Trim(Me.txtName.Text) = "" Then
        MsgBox "名称必须输入", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > Me.txtName.MaxLength Then
        MsgBox "名称超过" & Me.txtName.MaxLength & "的长度限制", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    '-----------------------------------------
    strTemp = ""
    With Me.msfTabu
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" And .RowData(intCount) <> 0 Then
                If InStr(1, strTemp & ";", ";" & .RowData(intCount) & ";") > 0 Then
                    MsgBox intCount & "行项目“" & Trim(.TextMatrix(intCount, 1)) & "”重复！", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
                strTemp = strTemp & ";" & .RowData(intCount)
            End If
        Next
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    If UBound(Split(strTemp, ";")) < 1 Then
        MsgBox "至少两个以上项目才能组成诊疗排斥关系！", vbExclamation, gstrSysName
        Me.msfTabu.SetFocus
        Exit Sub
    End If
    
    gstrSql = "zl_诊疗互斥项目_Save(" & Val(Me.lblName.Tag) & ",'" & Trim(Me.txtName.Text) & "'"
    If Me.optType(0).Value Then
        gstrSql = gstrSql & ",1,'" & strTemp & "')"
    ElseIf Me.optType(1).Value Then
        gstrSql = gstrSql & ",2,'" & strTemp & "')"
    Else
        gstrSql = gstrSql & ",3,'" & strTemp & "')"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    MsgBox "相关排斥组保存成功！", vbExclamation, gstrSysName
    
    Call zlGroupRef(Val(Me.lblName.Tag))
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwItems.Visible Then
        Me.lvwItems.Visible = False
        Me.msfTabu.SetFocus
    Else
        cmdClose_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With Me.msfTabu
        .Active = False
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 2
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "诊疗项目"
        .ColData(0) = 5: .ColData(1) = 1
        .ColWidth(0) = 250: .ColWidth(1) = 3500
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2500
        .Add , "编码", "编码", 1000
        .Add , "计算单位", "剂量单位", 900
    End With
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
    End With
End Sub

Private Sub lvwGroup_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.lblName.Tag = Mid(Item.Key, 2): Me.txtName.Text = Item.Text
    Me.optType(Val(Item.Tag) - 1).Value = True
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称" & _
            " from 诊疗项目目录 I,诊疗互斥项目 R" & _
            " where I.ID=R.项目ID and R.组编号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Mid(Item.Key, 2))
        
    With rsTemp
        Me.msfTabu.ClearBill
        Do While Not .EOF
            If Me.msfTabu.Rows - 1 < .AbsolutePosition Then Me.msfTabu.Rows = Me.msfTabu.Rows + 1
            Me.msfTabu.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.msfTabu.RowData(.AbsolutePosition) = !ID
            Me.msfTabu.TextMatrix(.AbsolutePosition, 1) = "[" & !编码 & "]" & !名称
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        Me.msfTabu.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
        Me.msfTabu.TextMatrix(Me.msfTabu.Row, 1) = Me.msfTabu.Text
        Me.msfTabu.RowData(Me.msfTabu.Row) = Mid(.SelectedItem.Key, 2)
        Me.msfTabu.SetFocus
        Call zlCommFun.PressKey(vbKeyReturn)
    End With
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub msfTabu_AfterAddRow(Row As Long)
    With Me.msfTabu
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfTabu_AfterDeleteRow()
    With Me.msfTabu
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfTabu_cboClick(ListIndex As Long)
    Me.msfTabu.TextMatrix(Me.msfTabu.Row, 2) = Me.msfTabu.CboText
End Sub

Private Sub msfTabu_CommandClick()
    Err = 0: On Error GoTo ErrHand
    gstrSql = "select I.ID,I.编码,I.名称,I.计算单位" & _
            " from 诊疗项目目录 I" & _
            " where I.类别>='A'" & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "msfTabu_CommandClick")
'        Call SQLTest
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "在没有建立用于长嘱的诊疗项目，无法设置排斥关系！", vbExclamation, gstrSysName: Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "Item": objItem.SmallIcon = "Item"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("计算单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.msfTabu.Name
        .Left = Me.msfTabu.Left + 300
        .Top = Me.msfTabu.Top + Me.msfTabu.RowHeight(0)
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfTabu_EnterCell(Row As Long, Col As Long)
    strInputed = Me.msfTabu.TextMatrix(Row, Col)
End Sub

Private Sub msfTabu_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msfTabu_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msfTabu
        If .Active = False Then Exit Sub
        If .Col <> 1 Then Exit Sub
        If .TxtVisible = False Then
            If .TextMatrix(.Row, 1) = "" Then Exit Sub
            strTemp = UCase(Trim(.TextMatrix(.Row, 1)))
        Else
            If Trim(.Text) = "" Then Exit Sub
            strTemp = UCase(Trim(.Text))
        End If
    End With
    If strTemp = strInputed Then Exit Sub
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.编码,I.名称,I.计算单位" & _
            " from 诊疗项目目录 I,诊疗项目别名 N" & _
            " where I.ID=N.诊疗项目ID and I.类别>='A'" & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [1] or N.名称 like [2] or N.简码 like [2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "未找到项目，请重新输入！", vbExclamation, gstrSysName: Cancel = True: Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msfTabu.Text = "[" & !编码 & "]" & !名称
            Me.msfTabu.TextMatrix(Me.msfTabu.Row, 1) = Me.msfTabu.Text
            Me.msfTabu.RowData(Me.msfTabu.Row) = !ID
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "Item": objItem.SmallIcon = "Item"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("计算单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.msfTabu.Name
        .Left = Me.msfTabu.Left + 300
        .Top = Me.msfTabu.Top + Me.msfTabu.RowHeight(0)
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtName_GotFocus()
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
    Call OS.OpenIme(True)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtName_LostFocus()
    Call OS.OpenIme(False)
End Sub

Private Sub zlGroupRef(Optional lngGrpNo As Long)
    '--------------------------------------------------------
    '功能：刷新显示排斥组
    '入参：lngGrdNo-指定要选中的分组
    '--------------------------------------------------------
    Err = 0: On Error GoTo ErrHand
    gstrSql = "select distinct 组编号,组名称,类型" & _
            " from 诊疗互斥项目" & _
            " where 组名称 is not null and 组名称 not like '...%'"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.Title, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "zlGroupRef")
'        Call SQLTest
    With rsTemp
        Me.lvwGroup.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwGroup.ListItems.Add(, "_" & !组编号, !组名称, "Group", "Group"): objItem.Tag = !类型
            .MoveNext
        Loop
    End With
    
    Err = 0: On Error Resume Next
    If Me.lvwGroup.ListItems.Count > 0 Then
        If lngGrpNo <> 0 Then Me.lvwGroup.ListItems("_" & lngGrpNo).Selected = True
        If Me.lvwGroup.SelectedItem Is Nothing Then Me.lvwGroup.ListItems(0).Selected = True
        Me.lvwGroup.SelectedItem.EnsureVisible
    End If
    
    Err = 0: On Error GoTo 0
    If Me.lvwGroup.ListItems.Count > 0 Then
        Call lvwGroup_ItemClick(Me.lvwGroup.SelectedItem)
        If Me.Tag <> "查阅" Then
            Me.cmdDel.Enabled = True
            Me.txtName.Enabled = True
            Me.optType(0).Enabled = True: Me.optType(1).Enabled = True: Me.optType(2).Enabled = True
            Me.msfTabu.Active = True
            Me.cmdSave.Enabled = True: Me.cmdRestore.Enabled = True
        End If
    Else
        Me.lblName.Tag = "": Me.txtName.Text = ""
        Me.optType(0).Value = False: Me.optType(1).Value = False: Me.optType(2).Value = False
        Me.msfTabu.ClearBill
        
        Me.cmdDel.Enabled = False
        Me.txtName.Enabled = False
        Me.optType(0).Enabled = False: Me.optType(1).Enabled = False: Me.optType(2).Enabled = False
        Me.msfTabu.Active = False
        Me.cmdSave.Enabled = False: Me.cmdRestore.Enabled = False
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

