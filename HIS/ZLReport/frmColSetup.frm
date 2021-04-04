VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "列特性设置"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   Icon            =   "frmColSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboAlign 
      Height          =   300
      ItemData        =   "frmColSetup.frx":08CA
      Left            =   3240
      List            =   "frmColSetup.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CheckBox chkRowIS 
      Caption         =   "整行应用"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   3255
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwIf 
      Height          =   3975
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483628
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "条件名称"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "条件字段"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "条件关系"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "条件值"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "字体颜色"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "背景颜色"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "是否加粗"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "是否整行应用"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "对齐"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.TextBox txtValue 
      Height          =   300
      Left            =   3240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1453
      Width           =   2295
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "字体加粗"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtBackColor 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3315
      TabIndex        =   22
      Top             =   2325
      Width           =   255
   End
   Begin VB.CommandButton cmdBackColor 
      Height          =   255
      Left            =   5250
      Picture         =   "frmColSetup.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2295
      Width           =   270
   End
   Begin VB.TextBox txtBackColor1 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3240
      TabIndex        =   19
      Top             =   2272
      Width           =   2295
   End
   Begin VB.TextBox txtForeColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3315
      TabIndex        =   17
      Top             =   1918
      Width           =   255
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增(&A)"
      Height          =   350
      Left            =   240
      TabIndex        =   16
      Top             =   4350
      Width           =   1100
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   1470
      TabIndex        =   15
      Top             =   4350
      Width           =   1100
   End
   Begin VB.CommandButton cmdField 
      Height          =   255
      Left            =   5250
      Picture         =   "frmColSetup.frx":09DC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   652
      Width           =   270
   End
   Begin VB.TextBox txtField 
      Height          =   300
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   629
      Width           =   2295
   End
   Begin VB.CommandButton cmdForeColor 
      Height          =   255
      Left            =   5250
      Picture         =   "frmColSetup.frx":0AEA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1888
      Width           =   270
   End
   Begin VB.ComboBox cboRelation 
      Height          =   300
      ItemData        =   "frmColSetup.frx":0BF8
      Left            =   3240
      List            =   "frmColSetup.frx":0C05
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1041
      Width           =   2295
   End
   Begin VB.TextBox txtForeColor1 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3240
      TabIndex        =   12
      Top             =   1865
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   0
      TabIndex        =   10
      Top             =   4200
      Width           =   6000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3420
      TabIndex        =   7
      Top             =   4350
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&E)"
      Height          =   350
      Left            =   4650
      TabIndex        =   9
      Top             =   4350
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3240
      TabIndex        =   0
      Top             =   217
      Width           =   2295
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3000
      Left            =   720
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5000
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
   Begin MSComDlg.CommonDialog cdg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "对齐方式"
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   2685
      Width           =   735
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "*该设置会影响报表的执行效率，请谨慎使用！"
      Height          =   420
      Left            =   2400
      TabIndex        =   28
      Top             =   3600
      Width           =   3150
   End
   Begin VB.Label Label6 
      Caption         =   "条件值"
      Height          =   255
      Left            =   2580
      TabIndex        =   24
      Top             =   1485
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "背景颜色"
      Height          =   255
      Left            =   2400
      TabIndex        =   18
      Top             =   2310
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "条件字段"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   660
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "条件关系"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   1065
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "字体颜色"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   1890
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "条件名称"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmColSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjColProtertys As New RPTColProtertys
Private mblnOK As Boolean
Private mlngType As Long '0-汇总表项，1-任意表项
Private mfrmParent As Object
Private mstrData As String  '当前表格元素对应的数据源名称
Private mstrColData As String '当前表格元素对应的列名称
Private mstrSummaryFile As String

Private Enum enum_Col
    col_条件字段 = 1
    col_条件关系 = 2
    col_条件值 = 3
    col_字体颜色 = 4
    col_背景颜色 = 5
    col_是否加粗 = 6
    col_是否整行应用 = 7
    col_对齐 = 8
End Enum

Public Function ShowMe(objParent As Object, objColProtertys As RPTColProtertys, ByVal lngType As Long, ByVal strData As String, ByVal strColData As String, _
                       Optional ByVal strSummaryFile As String) As Boolean
    Set mobjColProtertys = objColProtertys
    mlngType = lngType
    Set mfrmParent = objParent
    mstrData = strData
    mstrColData = strColData
    mstrSummaryFile = strSummaryFile
    
    Me.Show 1, objParent
    
    Set objColProtertys = mobjColProtertys
    ShowMe = mblnOK
End Function

Private Sub cboAlign_Click()
    If Me.Visible = False Then Exit Sub
    If Not lvwIf.SelectedItem Is Nothing Then
        lvwIf.SelectedItem.SubItems(col_对齐) = CStr(cboAlign.ListIndex)
    End If
End Sub

Private Sub cboRelation_Click()
    If Not lvwIf.SelectedItem Is Nothing Then
        lvwIf.SelectedItem.SubItems(col_条件关系) = cboRelation.Text
    End If
End Sub

Private Sub cboRelation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub chkBold_Click()
    If Not lvwIf.SelectedItem Is Nothing Then
        lvwIf.SelectedItem.SubItems(col_是否加粗) = chkBold.Value
    End If
End Sub

Private Sub chkBold_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub chkRowIS_Click()
    If Not lvwIf.SelectedItem Is Nothing Then
        lvwIf.SelectedItem.SubItems(col_是否整行应用) = chkRowIS.Value
    End If
End Sub

Private Sub chkRowIS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub cmdAdd_Click()
    Dim strName As String
    Dim i As Long, j As Long
    Dim lvwItem As ListItem
    
    '获取不重复的名称
    For i = 1 To 1000
        For j = 1 To lvwIf.ListItems.count
            If lvwIf.ListItems(j).Text = "样式条件" & i Then
                Exit For
            End If
        Next
        If j > lvwIf.ListItems.count Then
            strName = "样式条件" & i
            Exit For
        End If
    Next
    '加入列表
    Set lvwItem = lvwIf.ListItems.Add(, , strName)
    lvwItem.SubItems(col_背景颜色) = &H80000005
    lvwItem.Selected = True
    lvwIf_ItemClick lvwItem
End Sub

Private Sub cmdBackColor_Click()
    On Error Resume Next
    
    cdg.CancelError = True
    cdg.Flags = &H1 Or &H2
    cdg.Color = txtBackColor.BackColor
    cdg.ShowColor
    If Err.Number = 0 Then
        txtBackColor.BackColor = cdg.Color
        If Not lvwIf.SelectedItem Is Nothing Then
            lvwIf.SelectedItem.SubItems(col_背景颜色) = txtBackColor.BackColor
        End If
    Else
        Err.Clear
    End If
End Sub

Private Sub cmdBackColor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call PressKey(vbKeyTab)
    If KeyCode = vbKeySpace Then cmdForeColor_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim lngindex As Long
    If Not lvwIf.SelectedItem Is Nothing Then
        lngindex = lvwIf.SelectedItem.Index
        If lvwIf.SelectedItem.Index - 1 > 0 Then
            lvwIf.ListItems(lvwIf.SelectedItem.Index - 1).Selected = True
            lvwIf_ItemClick lvwIf.SelectedItem
        ElseIf lvwIf.SelectedItem.Index + 1 < lvwIf.ListItems.count Then
            lvwIf.ListItems(lvwIf.SelectedItem.Index + 1).Selected = True
            lvwIf_ItemClick lvwIf.SelectedItem
        End If
        lvwIf.ListItems.Remove lngindex
    End If
End Sub

Private Sub cmdField_Click()
    SetParent tvw.hwnd, 0
    tvw.Top = Top + txtField.Top + txtField.Height + 350
    tvw.Left = Left + txtField.Left + 60
    tvw.ZOrder
    tvw.Visible = Not tvw.Visible
    tvw.Tag = 0
End Sub

Private Sub cmdField_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then cmdField_Click
End Sub

Private Sub cmdForeColor_Click()
    On Error Resume Next
    
    cdg.CancelError = True
    cdg.Flags = &H1 Or &H2
    cdg.Color = txtForeColor.BackColor
    cdg.ShowColor
    If Err.Number = 0 Then
        txtForeColor.BackColor = cdg.Color
        If Not lvwIf.SelectedItem Is Nothing Then
            lvwIf.SelectedItem.SubItems(col_字体颜色) = txtForeColor.BackColor
        End If
    Else
        Err.Clear
    End If
End Sub

Private Sub cmdForeColor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call PressKey(vbKeyTab)
    If KeyCode = vbKeySpace Then cmdForeColor_Click
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    For i = 1 To lvwIf.ListItems.count
        With lvwIf.ListItems(i)
            '条件名称不能超过25个汉字
            If TLen(.Text) > 50 Then
                .Selected = True
                lvwIf_ItemClick lvwIf.ListItems(i)
                MsgBox "条件名称不能超过25个汉字的长度。", vbInformation, App.Title
                txtName.SetFocus
                SelAll txtName
                
                Exit Sub
            End If
            If Trim(.Text) = "" Then
                .Selected = True
                lvwIf_ItemClick lvwIf.ListItems(i)
                MsgBox "请输入条件名称。", vbInformation, App.Title
                txtName.SetFocus
                SelAll txtName
                Exit Sub
            End If
            If chkRowIS.Value = 1 Then
                If .SubItems(col_条件字段) = "" Then
                    .Selected = True
                    lvwIf_ItemClick lvwIf.ListItems(i)
                    MsgBox "请选择条件字段。", vbInformation, App.Title
                    txtField.SetFocus
                    Exit Sub
                End If
                If .SubItems(col_条件关系) = "" Then
                    .Selected = True
                    lvwIf_ItemClick lvwIf.ListItems(i)
                    MsgBox "请选择条件关系。", vbInformation, App.Title
                    cboRelation.SetFocus
                    Exit Sub
                End If
            End If
        End With
    Next
    
    Set mobjColProtertys = New RPTColProtertys
    With lvwIf.ListItems
        If .count = 1 Then
            i = 1
            If Val(.Item(i).SubItems(col_字体颜色)) <> vbBlack Or Val(.Item(i).SubItems(col_背景颜色)) <> vbWhite _
                Or Val(.Item(i).SubItems(col_是否加粗)) = 1 Or Val(.Item(i).SubItems(col_对齐)) > Val("0-自动") Then
                GoSub proAdd
            End If
        Else
            For i = 1 To .count
                GoSub proAdd
            Next
        End If
    End With
    
    mblnOK = True
    Unload Me
    Exit Sub
    
proAdd:
    With lvwIf.ListItems
        mobjColProtertys.Add .Item(i).Text, Nvl(.Item(i).SubItems(col_条件字段), "") _
            , Nvl(.Item(i).SubItems(col_条件关系), ""), Nvl(.Item(i).SubItems(col_条件值), "") _
            , Val(.Item(i).SubItems(col_字体颜色)), Val(.Item(i).SubItems(col_背景颜色)) _
            , Val(.Item(i).SubItems(col_是否加粗)) = 1, Val(.Item(i).SubItems(col_是否整行应用)) = 1 _
            , Val(.Item(i).SubItems(col_对齐)), "_" & .Item(i).Text
    End With
    Return
End Sub

'Private Sub cmdValue_Click()
'    SetParent tvw.hwnd, 0
'    tvw.Top = Top + txtValue.Top + txtValue.Height + 350
'    tvw.Left = Left + txtValue.Left + 60
'    tvw.ZOrder
'    tvw.Visible = Not tvw.Visible
'    tvw.Tag = 1
'End Sub
'
'Private Sub cmdValue_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then Call PressKey(vbKeyTab)
'    If KeyCode = vbKeySpace Then cmdField_Click
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If tvw.Visible = True Then tvw.Visible = False: Exit Sub
        Unload Me
    End If
End Sub

Private Sub tvw_LostFocus()
    tvw.Visible = False
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key <> "Root" And Node.Children = 0 Then
        IIF(tvw.Tag = 1, txtValue, txtField).Text = Node.Parent.Text & "." & LevelText(Node)
    Else
        IIF(tvw.Tag = 1, txtValue, txtField).Text = ""
    End If
    tvw.Visible = False

    IIF(tvw.Tag = 1, txtValue, txtField).SetFocus

End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim lvwItem As ListItem
    
    With cboRelation
        .Clear
        .AddItem ""
        .AddItem "等于"
        .AddItem "大于"
        .AddItem "小于"
        .AddItem "不等于"
        .AddItem "大于等于"
        .AddItem "小于等于"
        .AddItem "左匹配"
        .AddItem "双向匹配"
    End With
    
    With cboAlign
        .Clear
        .AddItem "自动（缺省）"
        .AddItem "居左"
        .AddItem "居中"
        .AddItem "居右"
    End With

    If mlngType = 2 Then
        Call CopySubTree(mfrmParent.tvwSQL)
    Else
        Call CopySubTree(mfrmParent.tvw)
    End If

    With mobjColProtertys
        For i = 1 To mobjColProtertys.count
            Set lvwItem = lvwIf.ListItems.Add(, , mobjColProtertys.Item(i).条件名称)
            lvwItem.SubItems(col_条件字段) = mobjColProtertys.Item(i).条件字段
            lvwItem.SubItems(col_条件关系) = mobjColProtertys.Item(i).条件关系
            lvwItem.SubItems(col_条件值) = mobjColProtertys.Item(i).条件值
            lvwItem.SubItems(col_字体颜色) = mobjColProtertys.Item(i).字体颜色
            lvwItem.SubItems(col_背景颜色) = mobjColProtertys.Item(i).背景颜色
            lvwItem.SubItems(col_是否加粗) = IIF(mobjColProtertys.Item(i).是否加粗, 1, 0)
            lvwItem.SubItems(col_是否整行应用) = IIF(mobjColProtertys.Item(i).是否整行应用, 1, 0)
            lvwItem.SubItems(col_对齐) = Val(mobjColProtertys.Item(i).对齐)
        Next
        If lvwIf.ListItems.count > 0 Then lvwIf.ListItems.Item(1).Selected = True: lvwIf_ItemClick lvwIf.ListItems.Item(1)
    End With
    
    '自动产生一个条件名称
    If mobjColProtertys.count = 0 Then
        Set lvwItem = lvwIf.ListItems.Add(, , "[" & mstrColData & "]" & "1")
        lvwItem.SubItems(col_背景颜色) = vbWhite
        lvwItem.Selected = True
        lvwIf_ItemClick lvwItem
        txtField.Text = mstrData & "." & mstrColData
        cboAlign.ListIndex = 0
        lvwIf.ColumnHeaders(1).Width = lvwIf.Width - 120
    End If
End Sub

Private Sub lvwIf_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtName.Text = Item.Text
    txtField.Text = Item.SubItems(col_条件字段)
    CboSetText cboRelation, Item.SubItems(col_条件关系), False
    txtValue.Text = Item.SubItems(col_条件值)
    txtForeColor.BackColor = Val(Item.SubItems(col_字体颜色))
    txtBackColor.BackColor = Val(Item.SubItems(col_背景颜色))
    chkBold.Value = Val(Item.SubItems(col_是否加粗))
    chkRowIS.Value = Val(Item.SubItems(col_是否整行应用))
    If Val(Item.SubItems(col_对齐)) >= 0 Then
        cboAlign.ListIndex = Val(Item.SubItems(col_对齐))
    Else
        cboAlign.ListIndex = 0
    End If
End Sub

Private Sub CopySubTree(objtvw As Object)
    Dim objNode As Object, tmpNode As Object
    Dim objPar As RPTPar
    Dim objData As RPTData
    Dim strTmp As String
    
    For Each objNode In objtvw.Nodes
        If mdlPublic.GetStdNodeText(objNode.Text) = mstrData And objNode.Children <> 0 And objNode.Key <> "Root" Then Exit For
    Next
    
    tvw.Nodes.Clear
    Set tvw.ImageList = objtvw.ImageList
    
    Set tmpNode = tvw.Nodes.Add(, , objNode.Key, objNode.Text, objNode.Image, objNode.SelectedImage)
    tmpNode.Expanded = True
    tmpNode.Selected = True
    
    Set objNode = objNode.Child
    Do While Not objNode Is Nothing
        If mlngType = 1 Or InStr(mstrSummaryFile, objNode.Text) > 0 Then
            If Not IsType(Val(objNode.Tag), adLongVarBinary) Then
                Set tmpNode = tvw.Nodes.Add(objNode.Parent.Key, 4, objNode.Key, objNode.Text, objNode.Image, objNode.SelectedImage)
                tmpNode.Tag = objNode.Tag
            End If
        End If
        Set objNode = objNode.Next
    Loop
End Sub

Private Sub txtField_Change()
    If Not lvwIf.SelectedItem Is Nothing Then
        lvwIf.SelectedItem.SubItems(col_条件字段) = txtField.Text
    End If
End Sub

Private Sub txtField_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub txtName_GotFocus()
    SelAll txtName
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    If Not lvwIf.SelectedItem Is Nothing Then
        lvwIf.SelectedItem.Text = txtName.Text
    End If
End Sub

Private Sub txtValue_Change()
    If Not lvwIf.SelectedItem Is Nothing Then
        lvwIf.SelectedItem.SubItems(col_条件值) = txtValue.Text
    End If
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub


