VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeItemFind 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "收费项目查找"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList ils16 
      Left            =   1980
      Top             =   3870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeItemFind.frx":0000
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeItemFind.frx":0458
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeItemFind.frx":08AC
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeItemFind.frx":0D00
            Key             =   "ItemR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeItemFind.frx":1B52
            Key             =   "ItemRNo"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra内容 
      Caption         =   "查找内容"
      Height          =   765
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   7365
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   2835
         MaxLength       =   40
         TabIndex        =   4
         Tag             =   "A.标识主码"
         Top             =   300
         Width           =   1020
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   765
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "A.编码"
         Top             =   300
         Width           =   1020
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   4530
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "B.名称"
         Top             =   300
         Width           =   1020
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   6210
         MaxLength       =   12
         TabIndex        =   8
         Tag             =   "B.简码"
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "标识主码(&P)"
         Height          =   180
         Index           =   0
         Left            =   1845
         TabIndex        =   3
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "编码(&U)"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   2
         Left            =   3885
         TabIndex        =   5
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "简码(&S)"
         Height          =   180
         Index           =   3
         Left            =   5595
         TabIndex        =   7
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame fra高级 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   -60
      TabIndex        =   24
      Top             =   900
      Width           =   7635
      Begin VB.CheckBox chkCase 
         Caption         =   "区分大小写(&E)"
         Height          =   255
         Left            =   2940
         TabIndex        =   17
         Top             =   1170
         Width           =   1485
      End
      Begin VB.CheckBox chkStop 
         Caption         =   "包括已停用项目(&T)"
         Enabled         =   0   'False
         Height          =   225
         Left            =   4740
         TabIndex        =   18
         Top             =   1170
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Frame fra范围 
         Caption         =   "查找范围"
         Height          =   1035
         Left            =   2820
         TabIndex        =   13
         Top             =   30
         Width           =   4755
         Begin VB.OptionButton optScope 
            Caption         =   "当前分类下的所有项目(&2)"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   15
            Top             =   660
            Width           =   2385
         End
         Begin VB.OptionButton optScope 
            Caption         =   "所有类别(&1)"
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   14
            Top             =   285
            Width           =   1335
         End
         Begin VB.OptionButton optScope 
            Caption         =   "当前分类下的直属项目(&3)"
            Height          =   195
            Index           =   3
            Left            =   2280
            TabIndex        =   16
            Top             =   315
            Width           =   2385
         End
      End
      Begin VB.Frame fra方式 
         Caption         =   "匹配方式"
         Height          =   1305
         Left            =   210
         TabIndex        =   9
         Top             =   30
         Width           =   2295
         Begin VB.OptionButton optMode 
            Caption         =   "包含输入内容(&C)"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   12
            Top             =   990
            Value           =   -1  'True
            Width           =   1845
         End
         Begin VB.OptionButton optMode 
            Caption         =   "以输入内容开头(&B)"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   11
            Top             =   660
            Width           =   1845
         End
         Begin VB.OptionButton optMode 
            Caption         =   "完全相同(&A)"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   10
            Top             =   330
            Width           =   1845
         End
      End
   End
   Begin VB.CommandButton cmdLocate 
      Caption         =   "定位(&L)"
      Height          =   350
      Left            =   7770
      TabIndex        =   21
      Top             =   680
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   7770
      TabIndex        =   22
      Top             =   1680
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   7770
      TabIndex        =   23
      Top             =   2640
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   3735
      Left            =   120
      TabIndex        =   20
      Top             =   2430
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "_名称"
         Object.Tag             =   "名称"
         Text            =   "名称"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "_编码"
         Object.Tag             =   "编码"
         Text            =   "编码"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "_标识主码"
         Object.Tag             =   "标识主码"
         Text            =   "标识主码"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "_标识子码"
         Object.Tag             =   "标识子码"
         Text            =   "标识子码"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "_简码"
         Object.Tag             =   "简码"
         Text            =   "简码"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "_分类"
         Text            =   "分类"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   7770
      TabIndex        =   19
      Top             =   180
      Width           =   1100
   End
   Begin VB.Label lbl数量 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   7770
      TabIndex        =   26
      Top             =   4860
      Width           =   1100
   End
   Begin VB.Label lbl 
      Caption         =   "查找结果："
      Height          =   180
      Left            =   7770
      TabIndex        =   25
      Top             =   4560
      Width           =   900
   End
End
Attribute VB_Name = "frmChargeItemFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintColumn As Integer
Dim mblnItem As Boolean

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdLocate_Click()
    Dim strKey As String
    Dim strClass As String
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    On Error Resume Next
    With lvwMain.SelectedItem
        strClass = Mid(.Tag, 2, 1)
        If .SubItems(3) <> "未分类" Then
            strKey = "R" & .ListSubItems(1).Tag
            frmChargeManage.tvwMainItem.Nodes(strKey).Selected = True
            frmChargeManage.tvwMainItem.Nodes(strKey).EnsureVisible
            frmChargeManage.tvwMainItem_NodeClick frmChargeManage.tvwMainItem.SelectedItem
            Err.Clear
            frmChargeManage.lvwMain_S.ListItems(.Tag).Selected = True
            frmChargeManage.lvwMain_S.ListItems(.Tag).EnsureVisible
            If Err.Number = 35601 Then
                MsgBox "你找到的这条记录可能已被删除或停用，请刷新列表。", vbInformation, gstrSysName
                Err.Clear
                Exit Sub
            End If
            frmChargeManage.lvwMain_S_ItemClick frmChargeManage.lvwMain_S.SelectedItem
        Else
            frmChargeManage.tvwMainItem.Nodes("Root").Selected = True
            frmChargeManage.tvwMainItem.Nodes(strKey).EnsureVisible
            frmChargeManage.tvwMainItem_NodeClick frmChargeManage.tvwMainItem.SelectedItem
            Err.Clear
            frmChargeManage.lvwMain_S.ListItems(.Tag).Selected = True
            frmChargeManage.lvwMain_S.ListItems(.Tag).EnsureVisible
            If Err.Number = 35601 Then
                MsgBox "你找到的这条记录可能已被删除或停用，请刷新列表。", vbInformation, gstrSysName
                Err.Clear
                Exit Sub
            End If
            frmChargeManage.lvwMain_S_ItemClick frmChargeManage.lvwMain_S.SelectedItem
        End If
    End With
    Err.Clear
End Sub

Private Sub cmdFind_Click()
    On Error GoTo ErrHandle
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    Dim strTable As String
    Dim strWhere As String
    Dim str分类ID As String
    Dim i As Long
    Dim str编码 As String
    Dim str标识主码 As String
    Dim str名称 As String
    Dim str简码 As String
    
    str编码 = IIF(chkCase.Value = 0, UCase(txtEdit(0).Text), txtEdit(0).Text)
    str标识主码 = IIF(chkCase.Value = 0, UCase(txtEdit(1).Text), txtEdit(1).Text)
    str名称 = IIF(chkCase.Value = 0, UCase(txtEdit(2).Text), txtEdit(2).Text)
    str简码 = IIF(chkCase.Value = 0, UCase(txtEdit(3).Text), txtEdit(3).Text)
    
    For i = 0 To 3
        If zlCommFun.StrIsValid(txtEdit(i).Text) = False Then
            txtEdit(i).SetFocus
            Exit Sub
        End If
    Next
    With frmChargeManage.tvwMainItem
        If .SelectedItem Is Nothing Then Exit Sub
        If .SelectedItem.Key = "Root" Then
            str分类ID = ""
        Else
            If .SelectedItem.Key <> "Root" Then
                str分类ID = .SelectedItem.Tag
                If str分类ID = "0" Then
                    str分类ID = ""
                End If
            Else
                str分类ID = ""
            End If
        End If
    End With
    
    If chkStop.Value = 0 Then
        strWhere = " and (撤档时间 = to_date('3000-01-01','YYYY-MM-DD') or 撤档时间 is null) "
    End If
    '查找范围
    If optScope(0).Value = True Then
        strTable = "select ID,类别,分类ID,名称,编码,标识主码,标识子码,撤档时间 from 收费项目目录 where 类别<>'5' and 类别<>'6' and 类别<>'7'" & strWhere
    ElseIf optScope(2).Value = True Then
        If str分类ID = "" Then
            strTable = "select ID,类别,分类ID,名称,编码,标识主码,标识子码,撤档时间 from 收费项目目录  where 类别<>'5' and 类别<>'6' and 类别<>'7' " & strWhere & vbCrLf & _
            " and (分类id IN (SELECT id FROM 收费分类目录  START WITH 上级id is null  CONNECT BY PRIOR id=上级id) OR 分类id  is null ) "   ' start with 类别='" & str类别 & "'and 上级ID is null connect by prior ID=上级ID"
        Else
            strTable = "select ID,类别,分类ID,名称,编码,标识主码,标识子码,撤档时间 from 收费项目目录   where  类别<>'5' and 类别<>'6' and 类别<>'7' " & strWhere & vbCrLf & _
            " and (分类id IN (SELECT id FROM 收费分类目录  START WITH 上级ID=[1] CONNECT BY PRIOR id=上级id) OR 分类id=[1] ) "
        End If
    Else
        If str分类ID = "" Then
            strTable = "select ID,类别,分类ID,名称,编码,标识主码,标识子码,撤档时间 from 收费项目目录 " & _
            "where  类别<>'5' and 类别<>'6' and 类别<>'7'and 分类ID is null " & strWhere
        Else
            strTable = "select ID,类别,分类ID,名称,编码,标识主码,标识子码,撤档时间 from 收费项目目录 " & _
            "where   类别<>'5' and 类别<>'6' and 类别<>'7'and 分类ID=[1] " & strWhere
        End If
    End If
    '比较方式
    strWhere = ""
    If optmode(0).Value = True Then
        For i = 0 To 3
            If txtEdit(i).Text <> "" Then
                strWhere = strWhere & " and " & IIF(chkCase.Value = 0, "Upper(", "") & txtEdit(i).Tag & IIF(chkCase.Value = 0, ")", "") & "=[" & i + 2 & "] "
            End If
        Next
    ElseIf optmode(1).Value = True Then
        For i = 0 To 3
            If txtEdit(i).Text <> "" Then
                strWhere = strWhere & " and " & IIF(chkCase.Value = 0, "Upper(", "") & txtEdit(i).Tag & IIF(chkCase.Value = 0, ")", "") & " like [" & i + 2 & "] "
            End If
        Next
        str编码 = str编码 & "%"
        str标识主码 = str标识主码 & "%"
        str名称 = str名称 & "%"
        str简码 = str简码 & "%"
    Else
        For i = 0 To 3
            If txtEdit(i).Text <> "" Then
                strWhere = strWhere & " and " & IIF(chkCase.Value = 0, "Upper(", "") & txtEdit(i).Tag & IIF(chkCase.Value = 0, ")", "") & " like [" & i + 2 & "] "
            End If
        Next
        str编码 = "%" & str编码 & "%"
        str标识主码 = "%" & str标识主码 & "%"
        str名称 = "%" & str名称 & "%"
        str简码 = "%" & str简码 & "%"
    End If
    
    '得到SQL语句
    gstrSQL = "select distinct A.ID,A.类别,B.名称,A.编码,A.标识主码,A.标识子码,B.简码,C.名称 as 分类,A.分类ID,A.撤档时间 from (" & _
        strTable & ") A,(Select A.收费细目id, A.名称, A.简码 || '/' || B.简码 As 简码" & _
        " From 收费项目别名 A, 收费项目别名 B " & _
        " Where A.收费细目id = B.收费细目id And A.码类 = 1 And B.码类 = 2) B,收费分类目录 C where A.分类id=c.id(+) And  A.ID=B.收费细目ID and C.名称 is not NULL " & strWhere
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str分类ID, str编码, str标识主码, str名称, str简码)
    
    Me.MousePointer = 11
    zlControl.FormLock lvwMain.hwnd
    With lvwMain.ListItems
        .Clear
        i = 1
        Do Until rsTemp.EOF
            '得出正确的图标
            strWhere = "Item"
            If Not CDate(IIF(IsNull(rsTemp("撤档时间")), CDate("3000/1/1"), rsTemp("撤档时间"))) = CDate("3000/1/1") Then
                strWhere = strWhere & "No"
            End If
            '添加节点
            Set lst = .Add(, "C" & i, rsTemp("名称"), strWhere, strWhere)
            If InStr(strWhere, "No") > 0 Then lst.ForeColor = RGB(255, 0, 0)
            
            Dim lngCol  As Long
            Dim varValue As Variant
            '根据ListView的列名从数据库取数
            For lngCol = 2 To lvwMain.ColumnHeaders.Count
                varValue = rsTemp(lvwMain.ColumnHeaders(lngCol).Text).Value
                lst.SubItems(lngCol - 1) = IIF(IsNull(varValue), "", varValue)
                
                lst.ListSubItems(1).Tag = IIF(IsNull(rsTemp("分类ID")), "", rsTemp("分类ID"))
                lst.Tag = "C" & rsTemp("类别") & rsTemp("id")
                If InStr(strWhere, "No") > 0 Then lst.ListSubItems(lngCol - 1).ForeColor = RGB(255, 0, 0)
            Next
            rsTemp.MoveNext
            i = i + 1
        Loop
        If .Count > 0 Then .Item(1).Selected = True
        lbl数量.Caption = "共" & .Count & "条"
    End With
    
    zlControl.FormLock 0
    Me.MousePointer = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlControl.FormLock 0
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
    Dim intSel As Integer
    
    RestoreWinState Me, App.ProductName
    
    intSel = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "匹配方式", 2)
    If intSel > 2 Or intSel < 0 Then intSel = 2
    optmode(intSel).Value = True
    
    intSel = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "查找范围", 0)
    If intSel > 3 Or intSel < 0 Then intSel = 0
    optScope(intSel).Value = True
    
    intSel = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "区分大小写", 0)
    If intSel > 1 Or intSel < 0 Then intSel = 0
    chkCase.Value = intSel
    
    chkStop.Value = IIF(frmChargeManage.mnuViewShowStop.Checked, 1, 0)
    fra内容.Caption = fra内容.Caption & IIF(chkStop.Value, "(包含停用项目)", "(不含停用项目)")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intSel As Integer
    
    For intSel = 0 To 2
        If optmode(intSel).Value = True Then Exit For
    Next
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "匹配方式", intSel)
    For intSel = 0 To 3
        If intSel <> 1 Then
            If optScope(intSel).Value = True Then Exit For
        End If
    Next
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "查找范围", intSel)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "区分大小写", chkCase.Value)
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim sngLeft As Single
    
    lvwMain.Top = IIF(fra高级.Visible = True, fra高级.Top + fra高级.Height, fra高级.Top)
    lvwMain.Height = Me.ScaleHeight - lvwMain.Top - 120
    
    sngLeft = ScaleWidth - cmdFind.Width - 200
    If sngLeft >= 7770 Then
        cmdFind.Left = sngLeft
    Else
        sngLeft = 7770
        cmdFind.Left = sngLeft
    End If
    cmdLocate.Left = cmdFind.Left
    cmdExit.Left = cmdFind.Left
    cmdHelp.Left = cmdFind.Left
    lbl.Left = cmdFind.Left
    lbl数量.Left = cmdFind.Left
    lvwMain.Width = cmdFind.Left - lvwMain.Left - 245
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwMain.SortOrder = IIF(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub chkCase_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub chkStop_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub lvwMain_DblClick()
    If mblnItem = True Then Call cmdLocate_Click
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub optMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub optScope_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub
