VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditInOutClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "编辑入出类别"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   Icon            =   "frmEditInOutClass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList ImgLvwSmall 
      Left            =   7650
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -105
      TabIndex        =   13
      Top             =   4095
      Width           =   8595
   End
   Begin MSComctlLib.ImageList ImgLvw单据Small 
      Left            =   4620
      Top             =   930
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7110
      TabIndex        =   6
      Top             =   4260
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5850
      TabIndex        =   5
      Top             =   4260
      Width           =   1100
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   180
      TabIndex        =   7
      Top             =   4260
      Width           =   1100
   End
   Begin MSComctlLib.ListView Lvw单据分类列表 
      Height          =   3060
      Left            =   120
      TabIndex        =   3
      Top             =   900
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   5398
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "单据名称"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "说明"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.ComboBox Cbo性质 
      Height          =   300
      Left            =   6015
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   150
      Width           =   2175
   End
   Begin VB.TextBox Txt名称 
      Height          =   300
      Left            =   2370
      MaxLength       =   20
      TabIndex        =   1
      Top             =   180
      Width           =   2145
   End
   Begin VB.TextBox Txt编码 
      Height          =   300
      Left            =   585
      MaxLength       =   2
      TabIndex        =   0
      Top             =   180
      Width           =   645
   End
   Begin MSComctlLib.ListView Lvw入出类别 
      Height          =   1755
      Left            =   5400
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   3096
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
         Text            =   "名称"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "类别"
         Object.Width           =   1058
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "可使用该类别的单据："
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   675
      Width           =   1800
   End
   Begin VB.Label Lbl说明 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "备注(该单据已包含的类别)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5400
      TabIndex        =   11
      Top             =   960
      Width           =   2160
   End
   Begin VB.Label Lbl性质 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "性质"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5565
      TabIndex        =   10
      Top             =   210
      Width           =   360
   End
   Begin VB.Label Lbl名称 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1920
      TabIndex        =   9
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Lbl编码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "编码"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   8
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "frmEditInOutClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IntEditState As Integer             '1-新增、2-修改
Private BlnModifySuccess As Boolean         '是否编辑成功
Private BlnStartUp As Boolean
Private strSQL As String                    'Sql语句
Private RecClass As New ADODB.Recordset     '药品单据分类
Private BlnRunTime As Boolean               '是否正在动态装入
'----修改时传入
Private Lng类别ID As Long                   '类别ID
Private strCode As String                   '编码
Private strName As String                   '名称
Private StrInOut As Integer                 '入出系数
Private mstrKey As String                   '用来记录不符合要求的key值

Public Property Get EditState() As Integer
    EditState = IntEditState
End Property

Public Property Let EditState(ByVal vNewValue As Integer)
    IntEditState = vNewValue
End Property

Public Property Get 类别ID() As Long
    类别ID = Lng类别ID
End Property

Public Property Let 类别ID(ByVal vNewValue As Long)
    Lng类别ID = vNewValue
End Property

Public Property Get 编码() As String
    编码 = strCode
End Property

Public Property Let 编码(ByVal vNewValue As String)
    strCode = vNewValue
End Property

Public Property Get 名称() As String
    名称 = strName
End Property

Public Property Let 名称(ByVal vNewValue As String)
    strName = vNewValue
End Property

Public Property Get 系数() As String
    系数 = StrInOut
End Property

Public Property Let 系数(ByVal vNewValue As String)
    StrInOut = vNewValue
End Property

Private Sub Cbo性质_Click()
    '检查用户选择的内容是否正确（防止当用户在入库性质时选择了在出库性质时不能选择的单据，然后用户改为出库保存而产生误数据）
    Dim ItemSelect As ListItem
    
    mstrKey = ""
    DependOnCheck
    LoadInLvw
    '对用户所选择的单据再次选择，可清除非法选择
    For Each ItemSelect In Lvw单据分类列表.ListItems
        ItemSelect.Selected = True
        ItemSelect.Checked = False
        ItemSelect.Ghosted = Not CheckItemCheck
    Next
    Call RemoveList(mstrKey)
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    '--合法性检测--
    If Trim(Txt编码) = "" Then
        MsgBox "编码不能为空！", vbInformation, gstrSysName
        Txt编码.SetFocus
        Exit Sub
    End If
    If Trim(Txt名称) = "" Then
        MsgBox "名称不能为空！", vbInformation, gstrSysName
        Txt名称.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(Txt编码) Then
        MsgBox "编码中含有非法字符！", vbInformation, gstrSysName
        Txt编码.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Txt名称, vbFromUnicode)) > 20 Then
        MsgBox "名称超长！（最多20个字符或10个汉字）", vbInformation, gstrSysName
        Txt名称.SetFocus
        Exit Sub
    End If
    Txt编码 = Trim(Txt编码)
    If Len(Txt编码) <> 3 Then Txt编码 = String(3 - Len(Txt编码), "0") & Txt编码
    
    '--保存--
    Dim ItemThis As ListItem
    On Error GoTo ErrHand
    
    gcnOracle.BeginTrans
    Select Case EditState
    Case 1
        
        '--插入药品入出类别--
        Lng类别ID = zlDatabase.GetNextId("药品入出类别")
        gstrSQL = "zl_药品入出类别_insert (" & Lng类别ID & ",'" & Txt编码 & "','" & Txt名称 & "'," & Me.Cbo性质.ItemData(Me.Cbo性质.ListIndex) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-插入药品入出类别")
    Case 2
        '--修改药品入出类别--
        gstrSQL = "zl_药品入出类别_update (" & Lng类别ID & ",'" & Txt编码 & "','" & Txt名称 & "'," & Me.Cbo性质.ItemData(Me.Cbo性质.ListIndex) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-修改药品入出类别")
        '--删除药品单据性质--
        gstrSQL = "zl_药品单据性质_delete (" & Lng类别ID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-删除药品单据性质")
    End Select
        
    '--依次插入药品单据性质--
    For Each ItemThis In Lvw单据分类列表.ListItems
        With ItemThis
            If .Checked Then
                gstrSQL = "zl_药品单据性质_insert (" & Lng类别ID & "," & Mid(.Key, 3) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-依次插入药品单据性质")
            End If
        End With
    Next
    gcnOracle.CommitTrans
    
    BlnModifySuccess = True  '增加成功
    Call frmMedInOutClass.EditReturn(BlnModifySuccess)
    '--设置为新增状态--
    If EditState = 1 Then
        ClearConsForAddNew
    Else
        Unload Me
    End If
    Exit Sub
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Call frmMedInOutClass.EditReturn(BlnModifySuccess)
    
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub GetDefineSize()
    '功能：得到数据库的表字段的长度
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
    Dim lng编码长度 As Long
    Dim lng名称长度 As Long
    
    gstrSQL = "Select 编码,名称 From 药品入出类别 Where ID = 0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    lng编码长度 = rsTmp.Fields("编码").DefinedSize
    lng名称长度 = rsTmp.Fields("名称").DefinedSize
    
    Txt编码.MaxLength = lng编码长度
    Txt名称.MaxLength = lng名称长度
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strTemp As String
    
    BlnStartUp = False
    BlnModifySuccess = False
    
    Call GetDefineSize '获得字段长度
    
    If DependOnCheck = False Then Exit Sub
    If LoadInIcon = False Then Exit Sub
    LoadInLvw
    
    With Me.Cbo性质
        .Clear
        .AddItem "入库"
        .ItemData(.NewIndex) = 1
        .AddItem "出库"
        .ItemData(.NewIndex) = -1
        .ListIndex = 0
    End With
    
    If EditState = 1 Then
        Me.Txt编码 = GetMaxCode()
        Me.Cbo性质.ListIndex = IIF(系数 = 1, 0, 1)
    Else
        SetSelect
    End If
    BlnStartUp = True
End Sub

Private Function LoadInIcon() As Boolean
    '--为各控件装载图标--
    On Error Resume Next
    Err = 0
    LoadInIcon = False
    
    '--列表Lvw所属单据--
    With ImgLvw单据Small
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add , , LoadResPicture("BILL1", vbResIcon)
    End With
    With Lvw单据分类列表
        Set .SmallIcons = ImgLvw单据Small
    End With
    
    '--列表Lvw所属单据--
    With ImgLvwSmall
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add , , LoadResPicture("CLASS", vbResIcon)
    End With
    With Lvw入出类别
        Set .SmallIcons = ImgLvwSmall
    End With
    
    If Err <> 0 Then
        MsgBox "相关资源文件丢失，请与软件开发商联系！", vbInformation, gstrSysName
        Exit Function
    End If
    LoadInIcon = True
End Function

Private Function DependOnCheck() As Boolean
    DependOnCheck = False
    '--依赖数据检测--
    On Error GoTo errHandle
'        If .State = 1 Then .Close
    strSQL = "Select 编码,名称,性质,说明 From 药品单据分类 Order by 编码"
        
'        Call SQLTest(App.Title, Me.Caption, strSQL)
    Set RecClass = zlDatabase.OpenSQLRecord(strSQL, "DependOnCheck")
'        Call SQLTest
    With RecClass
        If .EOF Then
            MsgBox "药品单据分类数据不全，请与系统管理员联系！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    DependOnCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadInLvw()
    '--将所有单据分类填入--
    
    Dim ItemThis As ListItem
    
    Lvw单据分类列表.ListItems.Clear
    With RecClass
        Do While Not .EOF
            Set ItemThis = Lvw单据分类列表.ListItems.Add(, "K_" & !编码, !名称, , 1)
            ItemThis.SubItems(1) = IIF(IsNull(!说明), "", !说明)
            ItemThis.Tag = IIF(IsNull(!性质), 1, !性质)
            
            .MoveNext
        Loop
    End With
    
    With Lvw单据分类列表
        .ListItems(1).Selected = True
        .SelectedItem.Selected = True
    End With
    Lvw单据分类列表_ItemClick Lvw单据分类列表.SelectedItem
End Function

Private Sub Form_Unload(Cancel As Integer)
    mstrKey = ""
    If BlnRunTime Then
        MsgBox "正在动态装入数据，请稍候...", vbInformation, gstrSysName
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub Lvw单据分类列表_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Lvw单据分类列表
        .Sorted = False
        
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIF(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

Private Sub Lvw单据分类列表_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    '未完--检测是否允许设置为指定单据的入出类别--
    If Item.Ghosted Then Item.Checked = False
End Sub

Private Sub Lvw单据分类列表_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '--显示所选择的单据分类已包含的药品入出类别--
    Call 装入各单据已包含的入出类别
End Sub

Private Sub Lvw单据分类列表_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '--使当前鼠标所在的Item项为选中--
    
    Dim ItemThis As ListItem
    If Button <> 1 And Button <> 2 Then Exit Sub
    On Error Resume Next
    Err = 0
    
    With Lvw单据分类列表
        Set ItemThis = .HitTest(X, Y)
        If Err <> 0 Then Exit Sub
        
        ItemThis.Selected = True
        .SelectedItem.Selected = True
    End With
    Lvw单据分类列表_ItemClick Lvw单据分类列表.SelectedItem
End Sub

Private Function GetMaxCode() As String
    '--获取最大的编码--
    Dim RecGetMaxCode As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    
'        If .State = 1 Then .Close
    strSQL = "Select Max(编码) 编码 From 药品入出类别"
        
'        Call SQLTest(App.Title, Me.Caption, strSQL)
    Set RecGetMaxCode = zlDatabase.OpenSQLRecord(strSQL, "GetMaxCode")
'        Call SQLTest
    With RecGetMaxCode
        If .EOF Then
            GetMaxCode = "01"
        Else
            If IsNull(!编码) Then
                GetMaxCode = "01"
            Else
                GetMaxCode = CInt(!编码) + 1
                If Len(GetMaxCode) > 2 Then
                    GetMaxCode = "01"
                Else
                    GetMaxCode = String(2 - Len(GetMaxCode), "0") & GetMaxCode
                End If
            End If
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetSelect()
    '--装入数据--
    '--并根据药品单据性质设置Lvw单据分类列表各项的选中状态--
    
    Dim RecSetSelect As New ADODB.Recordset
    
    Me.Txt编码 = 编码
    Me.Txt名称 = 名称
    
    If 系数 = -1 Then Me.Cbo性质.ListIndex = 1
    
    On Error GoTo errHandle
    strSQL = "Select 单据 From 药品单据性质 Where 类别ID=[1] "
    Set RecSetSelect = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Lng类别ID)

    With RecSetSelect
        If .EOF Then Exit Sub
        Do While Not .EOF
            With RecClass
                .MoveFirst
                .Find "编码=" & RecSetSelect!单据
                If Not .EOF Then Lvw单据分类列表.ListItems("K_" & RecSetSelect!单据).Checked = True
            End With
            
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ClearConsForAddNew()
    '清除相关控件内容，为新增下一个做准备
    Dim ItemThis As ListItem
    Me.Txt编码 = GetMaxCode()
    Me.Txt名称 = ""
    
    For Each ItemThis In Lvw单据分类列表.ListItems
        ItemThis.Checked = False
    Next
    Call Cbo性质_Click
    Me.Txt编码.SetFocus
End Sub

Private Function CheckItemCheck() As Boolean
    '--检测是否允许设置为指定单据的入出类别--
    Dim RecCheck As New ADODB.Recordset
    Dim IntBillStyle As Integer
    
    CheckItemCheck = False
    
    On Error GoTo errHandle
    IntBillStyle = Lvw单据分类列表.SelectedItem.Tag
    If RecCheck.State = 1 Then RecCheck.Close
    
    Select Case IntBillStyle
    Case "1", "2"   '只允许一种入库/只允许一种出库
        If Me.Cbo性质.ItemData(Me.Cbo性质.ListIndex) = IIF(IntBillStyle = 1, -1, 1) Then
            mstrKey = mstrKey & "|" & Lvw单据分类列表.SelectedItem.Key
            Exit Function  '只允许一种入库时，当前是出库则退出；反之，亦然
        End If
        strSQL = "Select 类别ID From 药品单据性质 Where 单据=[1]"
        Set RecCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(Lvw单据分类列表.SelectedItem.Key, 3)))

        With RecCheck
            If Not .EOF Then
                If Not IsNull(!类别ID) Then
                    If EditState = 1 Then
                        mstrKey = mstrKey & "|" & Lvw单据分类列表.SelectedItem.Key
                        Exit Function
                    End If
                    If !类别ID <> 类别ID Then
                        mstrKey = mstrKey & "|" & Lvw单据分类列表.SelectedItem.Key
                        Exit Function
                    End If
                End If
            End If
        End With
        
    Case "3"    '只允许一种入库及出库
        strSQL = " Select ID,系数 From 药品入出类别 Where ID IN " & _
                 " (Select 类别ID From 药品单据性质 Where 单据=[1])"
        Set RecCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(Lvw单据分类列表.SelectedItem.Key, 3)))
        
        With RecCheck
            .Find "系数=" & Me.Cbo性质.ItemData(Me.Cbo性质.ListIndex)
            If Not .EOF Then
                If Not IsNull(!ID) Then
                    If EditState = 1 Then
                        mstrKey = mstrKey & "|" & Lvw单据分类列表.SelectedItem.Key
                        Exit Function
                    End If
                    If !ID <> 类别ID Then
                        mstrKey = mstrKey & "|" & Lvw单据分类列表.SelectedItem.Key
                        Exit Function
                    End If
                End If
            End If
        End With
        
    Case "4", "5"   '允许多种入库/允许多种出库
        If Me.Cbo性质.ItemData(Me.Cbo性质.ListIndex) = IIF(IntBillStyle = 4, -1, 1) Then
            mstrKey = mstrKey & "|" & Lvw单据分类列表.SelectedItem.Key
            Exit Function  '只允许一种入库时，当前是出库则退出；反之，亦然
        End If
    End Select
    
    CheckItemCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckData()
    '检查用户选择的内容是否正确（防止当用户在入库性质时选择了在出库性质时不能选择的单据，然后用户改为出库保存而产生误数据）
    Dim ItemSelect As ListItem
    
    '对用户所选择的单据再次选择，可清除非法选择
    For Each ItemSelect In Lvw单据分类列表.ListItems
        If ItemSelect.Checked Then
            ItemSelect.Selected = True
            Call CheckItemCheck
        End If
    Next
    Call RemoveList(mstrKey)
End Function

Private Sub Lvw入出类别_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Lvw入出类别
        .Sorted = False
        
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIF(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

Private Sub Lvw入出类别_GotFocus()
    OS.OpenImeByName
End Sub

Private Sub Txt编码_GotFocus()
    zlControl.TxtSelAll Txt编码
End Sub

Private Sub Txt名称_GotFocus()
    zlControl.TxtSelAll Txt名称
    If GetSetting("ZLSOFT", "私有全局\" & gstrDbUser, "输入法", "") <> "" Then
        OS.OpenImeByName GetSetting("ZLSOFT", "私有全局\" & gstrDbUser, "输入法", "")
    End If
End Sub

Private Sub 装入各单据已包含的入出类别()
    Dim RecLoad As New ADODB.Recordset
    Dim ItemThis As ListItem
    Dim strBegin As String, StrMiddle As String, strEnd As String
    Dim StrLoad As String, str单据 As String
    Dim StrShow As String '显示入出类别
    Dim IntStyle As Integer '入出系数
    
    On Error GoTo errHandle
    strBegin = " select '['||编码||']'||名称 入出类别,nvl(系数,1) 系数 From 药品入出类别" & _
               " Where ID IN (select 类别ID from 药品单据性质 where 单据=[1] "
    strEnd = " ) Order by 系数 Desc"
    
    Lvw入出类别.ListItems.Clear
    With Lvw单据分类列表
        If .ListItems.Count = 0 Then Exit Sub
        If .SelectedItem Is Nothing Then Exit Sub
    
        Set ItemThis = .SelectedItem
    End With
    
    StrShow = ""
    str单据 = Mid(ItemThis.Key, 3)
        
    StrLoad = strBegin & strEnd
    Set RecLoad = zlDatabase.OpenSQLRecord(StrLoad, Me.Caption, Val(str单据))
    
    With RecLoad
        Do While Not .EOF
            Set ItemThis = Lvw入出类别.ListItems.Add(, , !入出类别, , 1)
            ItemThis.SubItems(1) = IIF(!系数 = 1, "入库", "出库")
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Txt名称_LostFocus()
    If GetSetting("ZLSOFT", "私有全局\" & gstrDbUser, "输入法", "") <> "" Then OS.OpenImeByName
End Sub

Private Sub RemoveList(ByVal strKey As String)
    '移除不符合条件的列表
    Dim i As Integer
    
    If strKey <> "" Then
        For i = 1 To UBound(Split(strKey, "|"))
            Lvw单据分类列表.ListItems.Remove (Split(strKey, "|")(i))
        Next
    End If
End Sub
