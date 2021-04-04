VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendItemDept 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "适用科室"
   ClientHeight    =   5625
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5025
   Icon            =   "frmTendItemDept.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSelect 
      Caption         =   "全选(&A)"
      Height          =   350
      Index           =   0
      Left            =   585
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "全清(&E)"
      Height          =   350
      Index           =   1
      Left            =   1665
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4665
      Width           =   1100
   End
   Begin VB.OptionButton optApply 
      Caption         =   "暂不使用(&0)"
      Height          =   195
      Index           =   0
      Left            =   585
      TabIndex        =   8
      Top             =   1305
      Value           =   -1  'True
      Width           =   1950
   End
   Begin VB.OptionButton optApply 
      Caption         =   "全院通用项目(&1)"
      Height          =   195
      Index           =   1
      Left            =   585
      TabIndex        =   7
      Top             =   1590
      Width           =   1950
   End
   Begin VB.OptionButton optApply 
      Caption         =   "适用于以下部门(&2)"
      Height          =   195
      Index           =   2
      Left            =   585
      TabIndex        =   6
      Top             =   1890
      Width           =   1950
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -45
      TabIndex        =   5
      Top             =   540
      Width           =   5115
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "仅显示选择部门(&L)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2925
      TabIndex        =   4
      Top             =   4740
      Width           =   1830
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -45
      TabIndex        =   3
      Top             =   5085
      Width           =   5115
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2460
      TabIndex        =   2
      Top             =   5190
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   1
      Top             =   5190
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwBakup 
      Height          =   2475
      Left            =   -840
      TabIndex        =   0
      Tag             =   "10"
      Top             =   2175
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
   Begin MSComctlLib.ListView lvwApply 
      Height          =   2475
      Left            =   585
      TabIndex        =   9
      Tag             =   "10"
      Top             =   2190
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
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "可以根据医学专业的不同要求，指定该项目适用于部分部门或全院通用。"
      Height          =   360
      Left            =   795
      TabIndex        =   14
      Top             =   75
      Width           =   3960
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   240
      Picture         =   "frmTendItemDept.frx":000C
      Top             =   15
      Width           =   480
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "项目名称:   体温"
      Height          =   180
      Left            =   270
      TabIndex        =   13
      Top             =   705
      Width           =   1440
   End
   Begin VB.Label lblApply 
      AutoSize        =   -1  'True
      Caption         =   "使用范围(&S)"
      Height          =   180
      Left            =   270
      TabIndex        =   12
      Top             =   1005
      Width           =   990
   End
End
Attribute VB_Name = "frmTendItemDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mintKind As Integer       '病历种类
Private mlngFileId As Long        '病历文件ID
Private mblnOK As Boolean

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim lngCount As Long

Public Function ShowMe(ByVal frmParent As Object, ByVal lngFileId As Long) As Boolean
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
    mlngFileId = lngFileId
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select 项目名称, 适用科室 From 护理记录项目 Where 项目序号 = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId)
    With rsTemp
        If .RecordCount = 0 Then MsgBox "护理记录项目丢失(可能被其他用户删除)！", vbInformation, gstrSysName: Exit Function
        lblFile.Caption = "护理记录项目:   " & !项目名称
        optApply(IIf(IsNull(!适用科室), 0, !适用科室)).Value = True
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '可选部门与已选部门列表
    With Me.lvwBakup.ColumnHeaders
        .Add , "_编码", "编码", 900
        .Add , "_名称", "名称", 2000
        .Add , "_简码", "简码", 800
    End With
    With Me.lvwApply.ColumnHeaders
        .Add , "_编码", "编码", 900
        .Add , "_名称", "名称", 2000
        .Add , "_简码", "简码", 800
    End With
    With Me.lvwApply
        .SortKey = .ColumnHeaders("_编码").Index - 1
        .SortOrder = lvwAscending
    End With

    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.Id, d.编码, d.名称, d.简码, Decode(s.科室id, Null, 0, 1) As 选择" & _
            " From 部门表 d, 部门性质说明 m, (Select 科室id From 护理适用科室 Where 项目序号 = [1]) s" & _
            " Where d.Id = m.部门id And d.Id = s.科室id(+) And m.工作性质 = '临床' And m.服务对象 In (2, 3)"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId)
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
    
    '------------------------------------------------------------------------------------------------------------------
    Me.Show vbModal, frmParent
    
    ShowMe = mblnOK: Unload Me
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub chkSelect_Click()
    Dim objAdd As ListItem

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
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim strSelected As String
    strSelected = ""
    For Each objItem In Me.lvwApply.ListItems
        If objItem.Checked Then strSelected = strSelected & ";" & Mid(objItem.Key, 2)
    Next
    If strSelected <> "" Then strSelected = Mid(strSelected, 2)
    
    If Me.optApply(0).Value Then
        gstrSQL = "Zl_护理适用科室_Apply(" & mlngFileId & ",0,Null)"
    ElseIf Me.optApply(1).Value Then
        gstrSQL = "Zl_护理适用科室_Apply(" & mlngFileId & ",1,Null)"
    Else
        If strSelected = "" Then MsgBox "没有选择科室！", vbInformation, gstrSysName: Me.lvwApply.SetFocus: Exit Sub
        gstrSQL = "Zl_护理适用科室_Apply(" & mlngFileId & ",2,'" & strSelected & "')"
    End If
    
    Err = 0: On Error GoTo errHand
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mblnOK = True: Me.Hide: Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    For Each objItem In Me.lvwBakup.ListItems
        objItem.Checked = IIf(Index = 0, True, False)
    Next
    Call chkSelect_Click
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
End Sub

Private Sub optApply_Click(Index As Integer)
    Me.lvwApply.Enabled = Me.optApply(2).Value
    Me.chkSelect.Enabled = Me.optApply(2).Value
    Me.cmdSelect(0).Enabled = Me.optApply(2).Value
    Me.cmdSelect(1).Enabled = Me.optApply(2).Value
End Sub

Private Sub optApply_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub


