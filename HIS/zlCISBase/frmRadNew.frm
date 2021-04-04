VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRadNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "影像项目增加"
   ClientHeight    =   6360
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8400
   Icon            =   "frmRadNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8400
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   405
      Left            =   6840
      TabIndex        =   20
      Top             =   3840
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   714
      ButtonWidth     =   1349
      ButtonHeight    =   609
      TextAlignment   =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "全选"
            Key             =   "全选"
            Object.ToolTipText     =   "选择所有显示项目"
            Object.Tag             =   "全选"
            ImageKey        =   "SelectAll"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "全清"
            Key             =   "全清"
            Object.ToolTipText     =   "清除所有选择标志"
            Object.Tag             =   "全清"
            ImageKey        =   "ClearAll"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "影像检查补充信息"
      Height          =   1575
      Left            =   2760
      TabIndex        =   19
      Top             =   4200
      Width           =   5655
      Begin VB.ComboBox cbo类别 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cbo病检 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   555
         Width           =   2055
      End
      Begin VB.ComboBox cbo胶片 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   870
         Width           =   2055
      End
      Begin VB.TextBox txt准备 
         Height          =   300
         Left            =   1050
         MaxLength       =   100
         TabIndex        =   14
         Top             =   1185
         Width           =   4230
      End
      Begin VB.TextBox txt图象 
         Height          =   300
         Left            =   3855
         MaxLength       =   2
         TabIndex        =   12
         Top             =   870
         Width           =   780
      End
      Begin VB.Label lbl类别 
         AutoSize        =   -1  'True
         Caption         =   "影像类别"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl病检 
         AutoSize        =   -1  'True
         Caption         =   "可行病检"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl胶片 
         AutoSize        =   -1  'True
         Caption         =   "可发胶片"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   930
         Width           =   720
      End
      Begin VB.Label lbl准备 
         AutoSize        =   -1  'True
         Caption         =   "检查准备"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   1245
         Width           =   720
      End
      Begin VB.Label lbl图象 
         AutoSize        =   -1  'True
         Caption         =   "报告最大图象数目"
         Height          =   180
         Left            =   3840
         TabIndex        =   11
         Top             =   630
         Width           =   1440
      End
   End
   Begin VB.CheckBox chkOnly 
      Caption         =   "只显示检查项目(&C)"
      Height          =   255
      Left            =   2910
      TabIndex        =   4
      Top             =   3915
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      Height          =   350
      Left            =   7110
      TabIndex        =   18
      Top             =   5940
      Width           =   1100
   End
   Begin VB.CommandButton cmd帮助 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   195
      Picture         =   "frmRadNew.frx":058A
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5940
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   5970
      TabIndex        =   16
      Top             =   5940
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -45
      TabIndex        =   15
      Top             =   5820
      Width           =   8535
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -30
      TabIndex        =   1
      Top             =   510
      Width           =   8535
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   60
      Top             =   4785
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadNew.frx":06D4
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadNew.frx":0C6E
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadNew.frx":1208
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadNew.frx":1422
            Key             =   "ClearAll"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   5190
      Left            =   0
      TabIndex        =   2
      Tag             =   "1000"
      Top             =   585
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   9155
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   3255
      Left            =   2760
      TabIndex        =   3
      Top             =   570
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   5741
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
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   60
      Picture         =   "frmRadNew.frx":163C
      Top             =   15
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    影像检查项目只能从已经建立的在用检查类诊疗项目中选择增加，然后补充必要的影像检查信息，从而保证和临床应用的一致性。"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   630
      TabIndex        =   0
      Top             =   90
      Width           =   7650
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "弹出"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuSele 
         Caption         =   "全部选中(&A)"
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuPopuSele 
         Caption         =   "全部取消(&R)"
         Index           =   1
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "frmRadNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem

Dim strTemp As String, aryTemp() As String
Dim intCount As Integer

Private Sub cbo病检_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo胶片_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkOnly_Click()
    If Me.Tag = "Loading" Then Exit Sub
    LoadClass
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strDescribe As String
    
    strDescribe = "'" & Split(Me.cbo类别.Text, "-")(0) & "'"
    strDescribe = strDescribe & "," & Left(Me.cbo病检.Text, 1)
    strDescribe = strDescribe & "," & Left(Me.cbo胶片.Text, 1)
    strDescribe = strDescribe & ",'" & Trim(Me.txt准备.Text) & "'"
    strDescribe = strDescribe & "," & Val(Me.txt图象.Text)
    
    For Each objItem In Me.lvwItem.ListItems
        If objItem.Checked Then
            gstrSql = "zl_影像检查项目_Insert(" & Mid(objItem.Key, 2) & "," & strDescribe & ")"
            Err = 0: On Error Resume Next
            Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
            If Err <> 0 Then
                Call SaveErrLog
            End If
        End If
    Next
    
    MsgBox "本次设置的影像检查项目保存完毕！", vbExclamation, gstrSysName
    Call tvwClass_NodeClick(Me.tvwClass.SelectedItem)
    Call frmRadLists.zlRefItems
End Sub

Private Sub cmd帮助_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Load()
    Me.lvwItem.ColumnHeaders.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_编码", "编码", 1000
        .Add , "_名称", "名称", 2500
        .Add , "_部位", "部位", 1000
        .Add , "_计算单位", "单位", 600
    End With
    With Me.lvwItem
        .ColumnHeaders("_编码").Position = 1
        .SortKey = .ColumnHeaders("_编码").Index - 1: .SortOrder = lvwAscending
    End With
    
    '---------------------------------
    '装入基本数据
    gstrSql = "Select * From 影像检查类别 Order By 排列"
    Err = 0: On Error GoTo ErrHand
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    Me.cbo类别.Clear
    With rsTemp
        Do While Not .EOF
            Me.cbo类别.AddItem !编码 & "-" & !名称
            If !编码 = Mid(frmRadLists.lvwKind.SelectedItem.Key, 2) Then
                Me.cbo类别.ListIndex = Me.cbo类别.NewIndex
            End If
            .MoveNext
        Loop
    End With
        
    aryTemp = Split("0-不可能;1-必须;2-选择进行", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo病检.AddItem aryTemp(intCount)
    Next
    Me.cbo病检.ListIndex = 0
    
    aryTemp = Split("0-不可能;1-必须;2-选择发放", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo胶片.AddItem aryTemp(intCount)
    Next
    Me.cbo胶片.ListIndex = 0
    
    Me.Tag = "Loading"
    chkOnly.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "影像只选择检查项目", 1))
    Me.Tag = ""
    LoadClass
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadClass()
    Dim strCurrKey As String
    '----------------------------------
'    gstrSql = "Select Distinct ID, 编码, 名称, 上级ID" & _
'            " From 诊疗分类目录" & _
'            " Where 类型 = '5'" & _
'            " Start With id In (Select Distinct 分类id From 诊疗项目目录 Where 类别 = 'D' Or" & _
'            " (类别='E' And 操作类型='5') Or 类别='Z')" & _
'            " Connect By Prior 上级ID = ID" & _
'            " Order By 编码"
    gstrSql = "Select Distinct ID, 编码, 名称, 上级ID" & _
            " From 诊疗分类目录" & _
            " Where 类型 = '5'" & _
            " Start With id In (Select Distinct 分类id From 诊疗项目目录" & _
            IIf(chkOnly.Value = 1, " Where 类别 = 'D')", ")") & _
            " Connect By Prior 上级ID = ID" & _
            " Order By 编码"
    
    Err = 0: On Error GoTo ErrHand
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "LoadClass")
'        Call SQLTest
    With rsTemp
        If Not tvwClass.SelectedItem Is Nothing Then strCurrKey = tvwClass.SelectedItem.Key
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            End If
            objNode.Sorted = True
            objNode.ExpandedImage = "expend"
            If strCurrKey = objNode.Key Then objNode.Selected = True
            .MoveNext
        Loop
    End With
    Err = 0: On Error GoTo 0
    If Me.tvwClass.Nodes.count > 0 Then
        If tvwClass.SelectedItem Is Nothing Then Me.tvwClass.Nodes(1).Selected = True
        tvwClass.SelectedItem.EnsureVisible
        Call tvwClass_NodeClick(Me.tvwClass.SelectedItem)
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "影像只选择检查项目", chkOnly.Value)
End Sub

Private Sub lvwItem_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu Me.mnuPopu, 2
    End If
End Sub

Private Sub mnuPopuSele_Click(Index As Integer)
    For Each objItem In Me.lvwItem.ListItems
        objItem.Checked = (Index = 0)
    Next
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
     Select Case Button.Key
        Case "全选"
            SelectAll True
        Case "全清"
            SelectAll False
     End Select
End Sub

Private Sub SelectAll(ByVal blnSelect As Boolean)
    Dim i As Integer
    With lvwItem
        For i = 1 To .ListItems.count
            .ListItems(i).Checked = blnSelect
        Next
    End With
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    
    gstrSql = "Select I.ID,I.编码, I.名称,I.标本部位, I.计算单位" & _
            "   From 诊疗项目目录 I" & _
            " Where " & IIf(chkOnly.Value = 1, "类别 = 'D' And ", "") & "分类id In " & _
            "       (Select id From 诊疗分类目录 Start With id = [1] Connect By Prior id = 上级ID)" & _
            "       And (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            " Minus" & _
            " Select I.ID,I.编码, I.名称,I.标本部位, I.计算单位" & _
            "   From 诊疗项目目录 I, 影像检查项目 R" & _
            "  Where I.ID = R.诊疗项目id"
    
    Err = 0: On Error GoTo ErrHand
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Mid(Node.Key, 2))
        
    With rsTemp
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !编码)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_名称").Index - 1) = !名称
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_部位").Index - 1) = IIf(IsNull(!标本部位), "", !标本部位)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_计算单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
            .MoveNext
        Loop
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt图象_GotFocus()
    Me.txt图象.SelStart = 0: Me.txt图象.SelLength = 100
End Sub

Private Sub txt图象_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt准备_GotFocus()
    Me.txt准备.SelStart = 0: Me.txt准备.SelLength = Me.txt准备.MaxLength
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt准备_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt准备_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub
