VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFixValue 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "固定值列表编辑"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "frmFixValue.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdAdd 
      Caption         =   "增加(&A)"
      Height          =   350
      Left            =   5115
      TabIndex        =   6
      Top             =   1035
      Width           =   1100
   End
   Begin VB.CommandButton cmdModi 
      Caption         =   "修改(&M)"
      Height          =   350
      Left            =   6465
      TabIndex        =   7
      Top             =   1035
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   7815
      TabIndex        =   8
      Top             =   1035
      Width           =   1100
   End
   Begin VB.TextBox txtDisp 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5940
      TabIndex        =   4
      Top             =   195
      Width           =   2925
   End
   Begin VB.TextBox txtBand 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5940
      TabIndex        =   5
      Top             =   600
      Width           =   2925
   End
   Begin VB.Frame Fra选择参数模式 
      Caption         =   "参数样式"
      Height          =   1275
      Left            =   5025
      TabIndex        =   9
      Top             =   1740
      Width           =   3855
      Begin VB.OptionButton Opt 
         Caption         =   "下拉框(&L)"
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   10
         Top             =   300
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton Opt 
         Caption         =   "单选框(&S)"
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   11
         Top             =   630
         Width           =   2055
      End
      Begin VB.OptionButton Opt 
         Caption         =   "单个复选框(&F)"
         Height          =   195
         Index           =   2
         Left            =   780
         TabIndex        =   12
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7455
      TabIndex        =   14
      Top             =   3210
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6150
      TabIndex        =   13
      Top             =   3210
      Width           =   1100
   End
   Begin VB.Frame fraValue 
      Caption         =   "参数:住院部门"
      ForeColor       =   &H00C00000&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   4665
      Begin VB.CommandButton cmdDown 
         Caption         =   "↓"
         Height          =   435
         Left            =   4200
         TabIndex        =   3
         ToolTipText     =   "向下移"
         Top             =   2880
         Width           =   345
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "↑"
         Height          =   435
         Left            =   4200
         TabIndex        =   2
         ToolTipText     =   "向上移"
         Top             =   2310
         Width           =   345
      End
      Begin MSComctlLib.ListView lvwValue 
         Height          =   3225
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   5689
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "显示内容"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "绑定值"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "缺省"
            Object.Width           =   1058
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   270
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFixValue.frx":0ECA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "显示内容(&V)"
      Height          =   180
      Left            =   4905
      TabIndex        =   16
      Top             =   255
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "绑定值(&B)"
      Height          =   180
      Left            =   5085
      TabIndex        =   15
      Top             =   660
      Width           =   810
   End
End
Attribute VB_Name = "frmFixValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbytSelType As Byte  '选择模式：0-下拉框;1-单选框;2-复选框
Public mstrValue As String '入/出：可选固定值列表串(显示,值|...)
Public mbytDataType As Byte    '入：参数数据类型
Public mstrParName As String '入：参数名称

Private Sub cmdAdd_Click()
    Dim i As Integer, intLen As Integer
    Dim blnExist As Boolean, objItem As ListItem
    
    '类型检查
    Select Case mbytDataType
        Case 1
            If Not IsNumeric(txtBand.Text) Then
                MsgBox "该参数是数字类型，只能绑定数字值！", vbInformation, App.Title
                txtBand.SetFocus: Exit Sub
            End If
        Case 2
            If Not IsDate(txtBand.Text) Then
                MsgBox "该参数是日期类型，只能绑定日期值！", vbInformation, App.Title
                txtBand.SetFocus: Exit Sub
            End If
    End Select
    
    '长度、重复检查
    For i = 1 To lvwValue.ListItems.Count
        If lvwValue.ListItems(i).Text = txtDisp.Text Then blnExist = True: Exit For
        intLen = intLen + TLen(lvwValue.ListItems(i).Text) + TLen(lvwValue.ListItems(i).SubItems(1)) + 2
    Next
    If blnExist Then
        MsgBox "该选择值显示内容已经存在，请修改！", vbInformation, App.Title
        txtDisp.SetFocus: Exit Sub
    End If
    If intLen + TLen(txtDisp.Text) + TLen(txtBand.Text) + 4 > 4000 Then
        MsgBox "加入的可选值条数过多，请减少选择值的内容长度后再试！", vbInformation, App.Title
        txtDisp.SetFocus: Exit Sub
    End If
    
    Set objItem = lvwValue.ListItems.Add(, , txtDisp.Text, , 1)
    objItem.SubItems(1) = txtBand.Text
    
    objItem.Selected = True
    objItem.EnsureVisible
    
    Call SetFunState
    
    txtDisp.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    Dim intIdx As Integer
    
    intIdx = lvwValue.SelectedItem.Index
    
    lvwValue.ListItems.Remove lvwValue.SelectedItem.Index
    
    If lvwValue.ListItems.Count = 0 Then
        txtDisp.Text = ""
        txtBand.Text = ""
    ElseIf intIdx <= lvwValue.ListItems.Count Then
        lvwValue.ListItems(intIdx).Selected = True
        lvwValue.SelectedItem.EnsureVisible
        Call lvwValue_ItemClick(lvwValue.SelectedItem)
    Else
        lvwValue.ListItems(lvwValue.ListItems.Count).Selected = True
        lvwValue.SelectedItem.EnsureVisible
        Call lvwValue_ItemClick(lvwValue.SelectedItem)
    End If
    
    Call SetFunState
    
    lvwValue.SetFocus
End Sub

Private Sub cmdDown_Click()
    Dim strTmp As String
    
    strTmp = lvwValue.ListItems(lvwValue.SelectedItem.Index + 1).Text
    lvwValue.ListItems(lvwValue.SelectedItem.Index + 1).Text = lvwValue.SelectedItem.Text
    lvwValue.SelectedItem.Text = strTmp
    
    strTmp = lvwValue.ListItems(lvwValue.SelectedItem.Index + 1).SubItems(1)
    lvwValue.ListItems(lvwValue.SelectedItem.Index + 1).SubItems(1) = lvwValue.SelectedItem.SubItems(1)
    lvwValue.SelectedItem.SubItems(1) = strTmp
    
    strTmp = lvwValue.ListItems(lvwValue.SelectedItem.Index + 1).SubItems(2)
    lvwValue.ListItems(lvwValue.SelectedItem.Index + 1).SubItems(2) = lvwValue.SelectedItem.SubItems(2)
    lvwValue.SelectedItem.SubItems(2) = strTmp
    
    lvwValue.ListItems(lvwValue.SelectedItem.Index + 1).Selected = True
    lvwValue.SelectedItem.EnsureVisible
    
    Call lvwValue_ItemClick(lvwValue.SelectedItem)
    Call SetFunState
    
    lvwValue.SetFocus
End Sub

Private Sub cmdModi_Click()
    Dim i As Integer, intLen As Integer
    Dim blnExist As Boolean
    
    '类型检查
    Select Case mbytDataType
        Case 1
            If Not IsNumeric(txtBand.Text) Then
                MsgBox "该参数是数字类型，只能绑定数字值！", vbInformation, App.Title
                txtBand.SetFocus: Exit Sub
            End If
        Case 2
            If Not IsDate(txtBand.Text) Then
                MsgBox "该参数是日期类型，只能绑定日期值！", vbInformation, App.Title
                txtBand.SetFocus: Exit Sub
            End If
    End Select
    
    '长度、重复检查
    For i = 1 To lvwValue.ListItems.Count
        If i <> lvwValue.SelectedItem.Index Then
            If lvwValue.ListItems(i).Text = txtDisp.Text Then blnExist = True: Exit For
            intLen = intLen + TLen(lvwValue.ListItems(i).Text) + TLen(lvwValue.ListItems(i).SubItems(1)) + 2
        End If
    Next
    If blnExist Then
        MsgBox "该选择值显示内容已经存在，请修改！", vbInformation, App.Title
        txtDisp.SetFocus: Exit Sub
    End If
    If intLen + TLen(txtDisp.Text) + TLen(txtBand.Text) + 4 > 4000 Then
        MsgBox "加入的可选值条数过多，请减少选择值的内容长度后再试！", vbInformation, App.Title
        txtDisp.SetFocus: Exit Sub
    End If
    
    lvwValue.SelectedItem.Text = txtDisp.Text
    lvwValue.SelectedItem.SubItems(1) = txtBand.Text
        
    lvwValue.SelectedItem.EnsureVisible
    
    Call SetFunState
    
    lvwValue.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    
    If lvwValue.ListItems.Count = 0 Then
        MsgBox "没有设置可选择的固定值！", vbInformation, App.Title
        lvwValue.SetFocus: Exit Sub
    End If
    
    '获取新值列表
    '类型及长度在新增及修改时已检查
    mstrValue = ""
    For i = 1 To lvwValue.ListItems.Count
        mstrValue = mstrValue & "|" & lvwValue.ListItems(i).SubItems(2) & lvwValue.ListItems(i).Text & "," & lvwValue.ListItems(i).SubItems(1)
    Next
    mstrValue = Mid(mstrValue, 2)

    If InStr(mstrValue, "√") = 0 Then
        MsgBox "必须设置一个缺省的固定值！", vbInformation, App.Title
        mstrValue = "": lvwValue.SetFocus: Exit Sub
    End If
    If Opt(1).Value And lvwValue.ListItems.Count > 12 Then
        MsgBox "选择模式为单选框的参数值最多12个！", vbInformation, App.Title
        Exit Sub
    End If
    
    gblnOK = True
    Hide
End Sub

Private Sub cmdUp_Click()
    Dim strTmp As String
    
    strTmp = lvwValue.ListItems(lvwValue.SelectedItem.Index - 1).Text
    lvwValue.ListItems(lvwValue.SelectedItem.Index - 1).Text = lvwValue.SelectedItem.Text
    lvwValue.SelectedItem.Text = strTmp
    
    strTmp = lvwValue.ListItems(lvwValue.SelectedItem.Index - 1).SubItems(1)
    lvwValue.ListItems(lvwValue.SelectedItem.Index - 1).SubItems(1) = lvwValue.SelectedItem.SubItems(1)
    lvwValue.SelectedItem.SubItems(1) = strTmp
    
    strTmp = lvwValue.ListItems(lvwValue.SelectedItem.Index - 1).SubItems(2)
    lvwValue.ListItems(lvwValue.SelectedItem.Index - 1).SubItems(2) = lvwValue.SelectedItem.SubItems(2)
    lvwValue.SelectedItem.SubItems(2) = strTmp
    
    lvwValue.ListItems(lvwValue.SelectedItem.Index - 1).Selected = True
    lvwValue.SelectedItem.EnsureVisible
    
    Call lvwValue_ItemClick(lvwValue.SelectedItem)
    Call SetFunState
    
    lvwValue.SetFocus
End Sub

Private Sub Form_Activate()
    If lvwValue.ListItems.Count = 0 Then txtDisp.SetFocus
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim objItem As ListItem
    
    gblnOK = False
    
    For i = 0 To UBound(Split(mstrValue, "|"))
        If Left(Split(Split(mstrValue, "|")(i), ",")(0), 1) = "√" Then
            Set objItem = lvwValue.ListItems.Add(, , Mid(Split(Split(mstrValue, "|")(i), ",")(0), 2), , 1)
            objItem.SubItems(2) = "√"
        Else
            Set objItem = lvwValue.ListItems.Add(, , Split(Split(mstrValue, "|")(i), ",")(0), , 1)
        End If
        objItem.SubItems(1) = Split(Split(mstrValue, "|")(i), ",")(1)
    Next
    
    fraValue.Caption = "参数：" & IIf(mstrParName = "", "未定义", mstrParName)
    Opt(mbytSelType).Value = True
    If Not lvwValue.SelectedItem Is Nothing Then Call lvwValue_ItemClick(lvwValue.SelectedItem)
    Call SetFunState
End Sub

Private Sub lvwValue_DblClick()
    Dim i As Integer
    If lvwValue.SelectedItem Is Nothing Then Exit Sub
    
    For i = 1 To lvwValue.ListItems.Count
        If i <> lvwValue.SelectedItem.Index Then
            lvwValue.ListItems(i).SubItems(2) = ""
        Else
            If lvwValue.ListItems(i).SubItems(2) = "" Then
                lvwValue.ListItems(i).SubItems(2) = "√"
            Else
                lvwValue.ListItems(i).SubItems(2) = ""
            End If
        End If
    Next
End Sub

Private Sub lvwValue_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtDisp.Text = Item.Text
    txtBand.Text = Item.SubItems(1)
End Sub

Private Sub lvwValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And cmdDel.Enabled Then cmdDel_Click
End Sub

Private Sub lvwValue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then Call lvwValue_DblClick
End Sub

Private Sub Opt_Click(Index As Integer)
    mbytSelType = Index
End Sub

Private Sub txtBand_Change()
    Call SetFunState
End Sub

Private Sub txtBand_GotFocus()
    SelAll txtBand
End Sub

Private Sub txtBand_KeyPress(KeyAscii As Integer)
    If InStr("&~`!@#$^"",|√", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtDisp_Change()
    Call SetFunState
End Sub

Private Sub txtDisp_GotFocus()
    SelAll txtDisp
End Sub

Private Sub txtDisp_KeyPress(KeyAscii As Integer)
    If InStr("&~`!@#$^"",|√", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub SetFunState()
'功能：根据当前界面状态，设置功能键状态
    Opt(2).Enabled = (lvwValue.ListItems.Count = 2)
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    cmdAdd.Enabled = False
    cmdModi.Enabled = False
    cmdDel.Enabled = False
    
    If Not lvwValue.SelectedItem Is Nothing Then
        cmdDel.Enabled = True
        If lvwValue.ListItems.Count > 1 Then
            If lvwValue.SelectedItem.Index > 1 Then cmdUp.Enabled = True
            If lvwValue.SelectedItem.Index < lvwValue.ListItems.Count Then cmdDown.Enabled = True
        End If
        If Len(Trim(txtDisp.Text)) <> 0 And Len(Trim(txtBand.Text)) <> 0 Then
            If txtDisp.Text <> lvwValue.SelectedItem.Text Then
                cmdAdd.Enabled = True
            End If
            If txtDisp.Text <> lvwValue.SelectedItem.Text Or txtBand.Text <> lvwValue.SelectedItem.SubItems(1) Then
                cmdModi.Enabled = True
            End If
        End If
    Else
        If Len(Trim(txtDisp.Text)) <> 0 And Len(Trim(txtBand.Text)) <> 0 Then
            cmdAdd.Enabled = True
        End If
    End If
End Sub
