VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRFileApplyTo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "适用范围"
   ClientHeight    =   5640
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4995
   Icon            =   "frmEPRFileApplyTo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pic产科 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2940
      ScaleHeight     =   315
      ScaleWidth      =   1815
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1890
      Width           =   1815
      Begin VB.ComboBox cbo分娩时机 
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "适用场合"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   30
         TabIndex        =   8
         Top             =   60
         Width           =   720
      End
   End
   Begin MSComctlLib.ListView lvwBakup 
      Height          =   2475
      Left            =   -510
      TabIndex        =   15
      Tag             =   "10"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   4366
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3615
      TabIndex        =   14
      Top             =   5175
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2445
      TabIndex        =   13
      Top             =   5175
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -60
      TabIndex        =   12
      Top             =   5025
      Width           =   5115
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "仅显示选择部门(&L)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2910
      TabIndex        =   11
      Top             =   4785
      Width           =   1830
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -60
      TabIndex        =   10
      Top             =   585
      Width           =   5115
   End
   Begin VB.OptionButton optApply 
      Caption         =   "适用于以下部门(&2)"
      Height          =   195
      Index           =   2
      Left            =   570
      TabIndex        =   4
      Top             =   1935
      Width           =   1950
   End
   Begin VB.OptionButton optApply 
      Caption         =   "全院通用病历(&1)"
      Height          =   195
      Index           =   1
      Left            =   570
      TabIndex        =   3
      Top             =   1635
      Width           =   1950
   End
   Begin VB.OptionButton optApply 
      Caption         =   "暂不使用(&0)"
      Height          =   195
      Index           =   0
      Left            =   570
      TabIndex        =   2
      Top             =   1350
      Value           =   -1  'True
      Width           =   1950
   End
   Begin MSComctlLib.ListView lvwApply 
      Height          =   2475
      Left            =   570
      TabIndex        =   6
      Tag             =   "10"
      Top             =   2235
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   4366
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "全清(&E)"
      Height          =   350
      Index           =   1
      Left            =   1650
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4710
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "全选(&A)"
      Height          =   350
      Index           =   0
      Left            =   570
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4710
      Width           =   1100
   End
   Begin VB.Label lblApply 
      AutoSize        =   -1  'True
      Caption         =   "使用范围(&S)"
      Height          =   180
      Left            =   255
      TabIndex        =   5
      Top             =   1050
      Width           =   990
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "文件名称:   001-入院记录"
      Height          =   180
      Left            =   255
      TabIndex        =   1
      Top             =   750
      Width           =   2160
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   225
      Picture         =   "frmEPRFileApplyTo.frx":058A
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "可以根据医学专业的不同要求，指定该文件适用于部分部门或全院通用。"
      Height          =   360
      Left            =   780
      TabIndex        =   0
      Top             =   120
      Width           =   3960
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEPRFileApplyTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mintDef As Integer          '保留
Private mstrCode As String          '病历文件编号
Private mintKind As Integer       '病历种类
Private mlngFileID As Long        '病历文件ID
Private mblnOK As Boolean
Dim objItem As ListItem


Public Function ShowMe(ByVal frmParent As Object, ByVal lngFileId As Long) As Boolean
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
    mblnOK = False
    mlngFileID = lngFileId
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select 种类, 编号, 名称, 通用,保留 From 病历文件列表 Where ID = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        If .RecordCount = 0 Then MsgBox "文件丢失(可能被其他用户删除)！", vbInformation, gstrSysName: Exit Function
        mintDef = !保留
        mintKind = !种类
        mstrCode = !编号
        Me.lblFile.Caption = "文件名称:   " & !编号 & "-" & !名称
        Me.optApply(IIf(IsNull(!通用), 0, !通用)).Value = True
    End With
    
    '---------------------------------------------------
    '可选部门与已选部门列表
    With Me.lvwBakup.ColumnHeaders
        .Clear
        .Add , "_编码", "编码", 900
        .Add , "_名称", "名称", 2000
        .Add , "_简码", "简码", 800
    End With
    With Me.lvwApply.ColumnHeaders
        .Clear
        .Add , "_编码", "编码", 900
        .Add , "_名称", "名称", 2000
        .Add , "_简码", "简码", 800
    End With
    With Me.lvwApply
        .SortKey = .ColumnHeaders("_编码").Index - 1
        .SortOrder = lvwAscending
    End With
    With Me.cbo分娩时机
        .Clear
        .AddItem "未设定"
        .AddItem "分娩前"
        .AddItem "分娩后"
        .ListIndex = 0
    End With

    Select Case mintKind
    Case 1
        gstrSQL = "Select d.Id, d.编码, d.名称, d.简码, Decode(s.科室id, Null, 0, 1) As 选择" & _
                " From 部门表 d, 部门性质说明 m, (Select 科室id From 病历应用科室 Where 文件id = [1]) s" & _
                " Where d.Id = m.部门id And d.Id = s.科室id(+) And m.工作性质 = '临床' And m.服务对象 In (1, 3)"
    Case 2
        gstrSQL = "Select d.Id, d.编码, d.名称, d.简码, Decode(s.科室id, Null, 0, 1) As 选择" & _
                " From 部门表 d, 部门性质说明 m, (Select 科室id From 病历应用科室 Where 文件id = [1]) s" & _
                " Where d.Id = m.部门id And d.Id = s.科室id(+) And m.工作性质 = '临床' And m.服务对象 In (2, 3)"
    Case 3, 4
        gstrSQL = "Select d.Id, d.编码, d.名称, d.简码, Decode(s.科室id, Null, 0, 1) As 选择" & _
                " From 部门表 d, 部门性质说明 m, (Select 科室id From 病历应用科室 Where 文件id = [1]) s" & _
                " Where d.Id = m.部门id And d.Id = s.科室id(+) And m.工作性质 = '护理' And m.服务对象 In (2, 3)"
    Case 5, 6
        gstrSQL = "Select d.Id, d.编码, d.名称, d.简码, Decode(s.科室id, Null, 0, 1) As 选择" & _
                " From 部门表 d, 部门性质说明 m, (Select 科室id From 病历应用科室 Where 文件id = [1]) s" & _
                " Where d.Id = m.部门id And d.Id = s.科室id(+) And m.工作性质 = '临床'"
    Case Else
        Unload Me: ShowMe = False: Exit Function
    End Select
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        Do While Not .EOF
            Set objItem = Me.lvwBakup.ListItems.Add(, "_" & !ID, !编码)
            objItem.SubItems(Me.lvwBakup.ColumnHeaders("_名称").Index - 1) = !名称
            objItem.SubItems(Me.lvwBakup.ColumnHeaders("_简码").Index - 1) = "" & !简码
            If !选择 = 1 Then objItem.Checked = True
            
            Set objItem = Me.lvwApply.ListItems.Add(, "_" & !ID, !编码)
            objItem.SubItems(Me.lvwApply.ColumnHeaders("_名称").Index - 1) = !名称
            objItem.SubItems(Me.lvwBakup.ColumnHeaders("_简码").Index - 1) = "" & !简码
            If !选择 = 1 Then objItem.Checked = True
            .MoveNext
        Loop
    End With
    
    If mintKind = 3 Then
        '检查是否为产科,并进行相应设置
        Call SetObstetric
        '如果是产科,则提取并恢复分娩时机(8)
        Dim str格式 As String
        str格式 = ";;;;;;;;;"
        gstrSQL = "Select 格式 From 病历页面格式 Where 种类=[1] And 编号=[2]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取病历页面格式", mintKind, mstrCode)
        If NVL(rsTemp!格式) <> "" Then
            str格式 = rsTemp!格式
        End If
        
        If UBound(Split(str格式, ";")) >= 8 Then Me.cbo分娩时机.ListIndex = Val(Split(str格式, ";")(8))
    End If
    
    Me.Show vbModal, frmParent
    ShowMe = mblnOK
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub chkSelect_Click()
Dim objAdd As ListItem
Dim objItem As ListItem

    Me.lvwApply.ListItems.Clear
    If Me.chkSelect.Value Then
        For Each objItem In Me.lvwBakup.ListItems
            If objItem.Checked Then
                Set objAdd = Me.lvwApply.ListItems.Add(, objItem.Key, objItem.Text)
                objAdd.SubItems(Me.lvwApply.ColumnHeaders("_名称").Index - 1) = objItem.SubItems(Me.lvwApply.ColumnHeaders("_名称").Index - 1)
                objAdd.SubItems(Me.lvwApply.ColumnHeaders("_简码").Index - 1) = objItem.SubItems(Me.lvwApply.ColumnHeaders("_简码").Index - 1)
                objAdd.Checked = objItem.Checked
            End If
        Next
    Else
        For Each objItem In Me.lvwBakup.ListItems
            Set objAdd = Me.lvwApply.ListItems.Add(, objItem.Key, objItem.Text)
            objAdd.SubItems(Me.lvwApply.ColumnHeaders("_名称").Index - 1) = objItem.SubItems(Me.lvwApply.ColumnHeaders("_名称").Index - 1)
            objAdd.SubItems(Me.lvwApply.ColumnHeaders("_简码").Index - 1) = objItem.SubItems(Me.lvwApply.ColumnHeaders("_简码").Index - 1)
            objAdd.Checked = objItem.Checked
        Next
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim arr格式
    Dim strSelected As String
    Dim intStart As Integer, intEnd As Integer
    Dim str病历 As String, str产科 As String, str格式 As String
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    '提取选择的科室清单
    strSelected = ""
    For Each objItem In Me.lvwApply.ListItems
        If objItem.Checked Then strSelected = strSelected & ";" & Mid(objItem.Key, 2)
    Next
    If strSelected <> "" Then strSelected = Mid(strSelected, 2)
    
    If Me.optApply(0).Value Then
        str病历 = "Zl_病历文件列表_Applyto(" & mlngFileID & ",0,Null)"
    ElseIf Me.optApply(1).Value Then
        str病历 = "Zl_病历文件列表_Applyto(" & mlngFileID & ",1,Null)"
    Else
        If strSelected = "" Then MsgBox "没有选择科室！", vbInformation, gstrSysName: Me.lvwApply.SetFocus: Exit Sub
        str病历 = "Zl_病历文件列表_Applyto(" & mlngFileID & ",2,'" & strSelected & "')"
    End If
    
    '保存产科护理记录单的分娩时机
    If mintKind = 3 And mintDef <> -1 Then
        str格式 = ";;;;;;;"
        gstrSQL = "Select 格式 From 病历页面格式 Where 种类=[1] And 编号=[2]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取病历页面格式", mintKind, mstrCode)
        If NVL(rsTemp!格式) <> "" Then
            str格式 = rsTemp!格式
        End If
        
        '前8位不动,拼接上产科分娩时机
        intEnd = 7
        arr格式 = Split(str格式, ";")
        str格式 = ""
        For intStart = 0 To intEnd
            str格式 = str格式 & ";" & arr格式(intStart)
        Next
        str格式 = Mid(str格式, 2) & ";" & IIf(pic产科.Visible, Me.cbo分娩时机.ListIndex, 0)
        
        str产科 = "Zl_病历页面格式_Format(" & mintKind & ",'" & mstrCode & "','" & str格式 & "')"
    End If
    
    Err = 0: On Error GoTo errHand
    If str产科 <> "" Then
        gcnOracle.BeginTrans
        blnTrans = True
    End If
    Call zldatabase.ExecuteProcedure(str病历, Me.Caption)
    If str产科 <> "" Then
        Call zldatabase.ExecuteProcedure(str产科, Me.Caption)
        gcnOracle.CommitTrans
        blnTrans = False
    End If
    mblnOK = True
    Unload Me
    Exit Sub

errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim objItem As ListItem
    For Each objItem In Me.lvwBakup.ListItems
        objItem.Checked = IIf(Index = 0, True, False)
    Next
    Call chkSelect_Click
    Call SetObstetric
    Me.lvwApply.SetFocus
End Sub

Private Sub lvwApply_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwApply.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwApply.SortOrder = IIf(Me.lvwApply.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwApply.SortKey = ColumnHeader.Index - 1
        Me.lvwApply.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwApply_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Me.lvwBakup.ListItems(Item.Key).Checked = Item.Checked
    Call SetObstetric
End Sub

Private Sub lvwApply_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call SetObstetric
End Sub

Private Sub optApply_Click(Index As Integer)
    Me.lvwApply.Enabled = Me.optApply(2).Value
    Me.chkSelect.Enabled = Me.optApply(2).Value
    Me.cmdSelect(0).Enabled = Me.optApply(2).Value
    Me.cmdSelect(1).Enabled = Me.optApply(2).Value
    Call SetObstetric
End Sub

Private Sub optApply_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Function IsObstetric() As Boolean
    Dim strSelected As String
    Dim intStart As Integer, intEnd As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '所选科室列表是否都为产科
    
    If mintKind <> 3 Then Exit Function     '只有护理文件才进入此流程
    If mintDef = -1 Then Exit Function
    If Not optApply(2).Value Then Exit Function
    
    '提取选择的科室
    strSelected = ""
    intEnd = Me.lvwApply.ListItems.Count
    For intStart = 1 To intEnd
        If lvwApply.ListItems(intStart).Checked Then
            strSelected = strSelected & "," & Mid(lvwApply.ListItems(intStart).Key, 2)
        End If
    Next
    If strSelected = "" Then Exit Function
    strSelected = Mid(strSelected, 2)
    
    '检查是否都具备产科的属性
    gstrSQL = "" & _
              " SELECT ID FROM 部门表 WHERE ID IN (Select Column_Value From Table(ZLTOOLS.f_Num2list([1])))" & vbNewLine & _
              " MINUS" & vbNewLine & _
              " SELECT 部门ID FROM 部门性质说明 WHERE 工作性质='产科'"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "检查是否都具备产科的属性", strSelected)
    IsObstetric = (rsTemp.RecordCount = 0)  '没有非产科部门
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetObstetric()
    Dim blnVisible As Boolean
    '在选择时判断,如果所选科室都具有产科属性则允许设置产科分娩时机
    
    blnVisible = IsObstetric
    pic产科.Visible = blnVisible
End Sub


