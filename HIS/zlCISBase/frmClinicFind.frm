VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicFind 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "诊疗项目查找"
   ClientHeight    =   6240
   ClientLeft      =   2835
   ClientTop       =   3825
   ClientWidth     =   8925
   Icon            =   "frmClinicFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   7710
      TabIndex        =   23
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   7710
      TabIndex        =   21
      Top             =   2670
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   7710
      TabIndex        =   20
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton cmdLocate 
      Caption         =   "定位(&L)"
      Height          =   350
      Left            =   7710
      TabIndex        =   19
      Top             =   705
      Width           =   1100
   End
   Begin VB.Frame fra高级 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   -120
      TabIndex        =   7
      Top             =   930
      Width           =   7635
      Begin VB.Frame fra方式 
         Caption         =   "匹配方式"
         Height          =   1305
         Left            =   210
         TabIndex        =   15
         Top             =   30
         Width           =   2295
         Begin VB.OptionButton optMode 
            Caption         =   "完全相同(&A)"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   18
            Top             =   330
            Width           =   1845
         End
         Begin VB.OptionButton optMode 
            Caption         =   "以输入内容开头(&B)"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   17
            Top             =   660
            Width           =   1845
         End
         Begin VB.OptionButton optMode 
            Caption         =   "包含输入内容(&C)"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   16
            Top             =   990
            Value           =   -1  'True
            Width           =   1845
         End
      End
      Begin VB.Frame fra范围 
         Caption         =   "查找范围"
         Height          =   1035
         Left            =   2600
         TabIndex        =   11
         Top             =   30
         Width           =   4995
         Begin VB.OptionButton optScope 
            Caption         =   "当前分类下的直属项目(&3)"
            Height          =   195
            Index           =   3
            Left            =   2280
            TabIndex        =   14
            Top             =   315
            Width           =   2385
         End
         Begin VB.OptionButton optScope 
            Caption         =   "所有类别(&1)"
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   13
            Top             =   285
            Width           =   1335
         End
         Begin VB.OptionButton optScope 
            Caption         =   "当前分类下的所有项目(&2)"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   12
            Top             =   660
            Width           =   2385
         End
      End
      Begin VB.CheckBox chkStop 
         Caption         =   "包括已停用项目(&T)"
         Enabled         =   0   'False
         Height          =   225
         Left            =   4200
         TabIndex        =   10
         Top             =   1170
         Width           =   1845
      End
      Begin VB.CheckBox chkCase 
         Caption         =   "区分大小写(&E)"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   1170
         Width           =   1485
      End
      Begin VB.CheckBox chkAlias 
         Caption         =   "按别名查找(&I)"
         Height          =   180
         Left            =   6120
         TabIndex        =   8
         Top             =   1170
         Width           =   1575
      End
   End
   Begin VB.Frame fra内容 
      Caption         =   "查找内容"
      Height          =   765
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   7365
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   5610
         MaxLength       =   12
         TabIndex        =   3
         Top             =   300
         Width           =   1620
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   3210
         MaxLength       =   40
         TabIndex        =   2
         Top             =   300
         Width           =   1620
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   765
         MaxLength       =   10
         TabIndex        =   1
         Top             =   300
         Width           =   1620
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "简码(&S)"
         Height          =   180
         Index           =   3
         Left            =   4985
         TabIndex        =   6
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   2
         Left            =   2565
         TabIndex        =   5
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "编码(&U)"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   630
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1920
      Top             =   3900
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
            Picture         =   "frmClinicFind.frx":058A
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicFind.frx":09E2
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicFind.frx":0E36
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicFind.frx":128A
            Key             =   "ItemR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicFind.frx":20DC
            Key             =   "ItemRNo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   3735
      Left            =   60
      TabIndex        =   22
      Top             =   2460
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "_名称"
         Object.Tag             =   "名称"
         Text            =   "名称"
         Object.Width           =   4763
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "_编码"
         Object.Tag             =   "编码"
         Text            =   "编码"
         Object.Width           =   1942
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "_简码"
         Object.Tag             =   "简码"
         Text            =   "简码"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "_分类"
         Text            =   "分类"
         Object.Width           =   2648
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "查找结果："
      Height          =   180
      Left            =   7710
      TabIndex        =   25
      Top             =   4590
      Width           =   900
   End
   Begin VB.Label lbl数量 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   7710
      TabIndex        =   24
      Top             =   4890
      Width           =   1095
   End
End
Attribute VB_Name = "frmClinicFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintColumn As Integer
Dim mblnItem As Boolean

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdLocate_Click()
    Dim strkey As String
    Dim strItemKey As String
    
    If LvwMain.SelectedItem Is Nothing Then Exit Sub
    On Error Resume Next
    With LvwMain.SelectedItem
        strkey = "_" & .ListSubItems(1).Tag
        strItemKey = "_" & .Tag
        If .SubItems(3) <> "未分类" Then
            frmClinicLists.tvwClass.Nodes(strkey).Selected = True
            frmClinicLists.tvwClass.Nodes(strkey).EnsureVisible
            frmClinicLists.tvwClass_NodeClick frmClinicLists.tvwClass.SelectedItem
            Err.Clear
            frmClinicLists.lvwItems.ListItems(strItemKey).Selected = True
            frmClinicLists.lvwItems.ListItems(strItemKey).EnsureVisible
            If Err.Number = 35601 Then
                MsgBox "你找到的这条记录可能已被删除或停用，请刷新列表。", vbInformation, gstrSysName
                Err.Clear
                Exit Sub
            End If
            frmClinicLists.lvwItems_ItemClick frmClinicLists.lvwItems.SelectedItem
        Else
            frmClinicLists.tvwClass.Nodes("Root").Selected = True
            frmClinicLists.tvwClass.Nodes(strkey).EnsureVisible
            frmClinicLists.tvwClass_NodeClick frmClinicLists.tvwClass.SelectedItem
            Err.Clear
            frmClinicLists.lvwItems.ListItems(strItemKey).Selected = True
            frmClinicLists.lvwItems.ListItems(strItemKey).EnsureVisible
            If Err.Number = 35601 Then
                MsgBox "你找到的这条记录可能已被删除或停用，请刷新列表。", vbInformation, gstrSysName
                Err.Clear
                Exit Sub
            End If
            frmClinicLists.lvwItems_ItemClick frmClinicLists.lvwItems.SelectedItem
        End If
    End With
    Err.Clear
End Sub

Private Sub cmdFind_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    Dim strTable As String
    Dim strWhere As String
    Dim strWhereAlias As String
    Dim strAlaisSql As String
    Dim str分类ID As String
    Dim i As Long
    Dim str编码 As String
    Dim str名称 As String
    Dim str简码 As String
    
    On Error GoTo errHandle
    
    str编码 = IIf(chkCase.Value = 0, UCase(txtEdit(0).Text), txtEdit(0).Text)
    str名称 = IIf(chkCase.Value = 0, UCase(txtEdit(1).Text), txtEdit(1).Text)
    str简码 = IIf(chkCase.Value = 0, UCase(txtEdit(2).Text), txtEdit(2).Text)
    
    For i = 0 To 2
        If zlCommFun.StrIsValid(txtEdit(i).Text) = False Then
            txtEdit(i).SetFocus
            Exit Sub
        End If
    Next
    With frmClinicLists.tvwClass
        If .SelectedItem Is Nothing Then Exit Sub
        If .SelectedItem.Key = "Root" Then
            str分类ID = ""
        Else
            If .SelectedItem.Key <> "Root" Then
                str分类ID = Val(Mid(.SelectedItem.Key, 2))
                If str分类ID = "0" Then
                    str分类ID = ""
                End If
            Else
                str分类ID = ""
            End If
        End If
    End With
    
    strWhere = " And 类别" & Me.Tag
    If chkStop.Value = 0 Then
        strWhere = strWhere & " and (撤档时间 = to_date('3000-01-01','YYYY-MM-DD') or 撤档时间 is null) "
    End If
    '查找范围
    If optScope(0).Value = True Then
        strTable = "select ID,类别,分类ID,名称,编码,撤档时间 from 诊疗项目目录 where 类别 Not In('4','5','6','7') " & strWhere
    ElseIf optScope(2).Value = True Then
        If str分类ID = "" Then
            strTable = "select ID,类别,分类ID,名称,编码,撤档时间 from 诊疗项目目录  where 类别 Not In('4','5','6','7') " & strWhere & vbCrLf & _
            " and (分类id IN (SELECT id FROM 诊疗分类目录  START WITH 上级id is null  CONNECT BY PRIOR id=上级id) OR 分类id  is null ) "   ' start with 类别='" & str类别 & "'and 上级ID is null connect by prior ID=上级ID"
        Else
            strTable = "select ID,类别,分类ID,名称,编码,撤档时间 from 诊疗项目目录   where  类别 Not In('4','5','6','7') " & strWhere & vbCrLf & _
            " and (分类id IN (SELECT id FROM 诊疗分类目录  START WITH 上级ID=[1] CONNECT BY PRIOR id=上级id) OR 分类id=[1] ) "
        End If
    Else
        If str分类ID = "" Then
            strTable = "select ID,类别,分类ID,名称,编码,撤档时间 from 诊疗项目目录 " & _
            "where  类别 Not In('4','5','6','7') and 分类ID is null " & strWhere
        Else
            strTable = "select ID,类别,分类ID,名称,编码,撤档时间 from 诊疗项目目录 " & _
            "where  类别 Not In('4','5','6','7') and 分类ID=[1] " & strWhere
        End If
    End If
    '比较方式
    strWhere = ""
    If optMode(0).Value = True Then
        If Trim(txtEdit(0).Text) <> "" Then
            strWhere = strWhere & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "A.编码" & IIf(chkCase.Value = 0, ")", "") & "=[2] "
        End If
        
        If Trim(txtEdit(1).Text) <> "" Then
            strWhereAlias = strWhereAlias & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "A.名称" & IIf(chkCase.Value = 0, ")", "") & "=[3] "
            strWhere = strWhere & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "Nvl(B.名称, A.名称)" & IIf(chkCase.Value = 0, ")", "") & "=[3] "
        End If
        
        If Trim(txtEdit(2).Text) <> "" Then
            strWhereAlias = strWhereAlias & " And (" & IIf(chkCase.Value = 0, "Upper(", "") & "A.简码" & IIf(chkCase.Value = 0, ")", "") & "=[4] " & _
                " Or " & IIf(chkCase.Value = 0, "Upper(", "") & "B.简码" & IIf(chkCase.Value = 0, ")", "") & "=[4] " & ") "
            strWhere = strWhere & " And (" & IIf(chkCase.Value = 0, "Upper(", "") & "B.拼音码" & IIf(chkCase.Value = 0, ")", "") & "=[4] " & _
                " Or " & IIf(chkCase.Value = 0, "Upper(", "") & "B.五笔码" & IIf(chkCase.Value = 0, ")", "") & "=[4] " & ") "
        End If
    ElseIf optMode(1).Value = True Then
        If Trim(txtEdit(0).Text) <> "" Then
            strWhere = strWhere & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "A.编码" & IIf(chkCase.Value = 0, ")", "") & " Like [2] "
        End If
        
        If Trim(txtEdit(1).Text) <> "" Then
            strWhereAlias = strWhereAlias & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "A.名称" & IIf(chkCase.Value = 0, ")", "") & " Like [3] "
            strWhere = strWhere & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "Nvl(B.名称, A.名称)" & IIf(chkCase.Value = 0, ")", "") & " Like [3] "
        End If
        
        If Trim(txtEdit(2).Text) <> "" Then
            strWhereAlias = strWhereAlias & " And (" & IIf(chkCase.Value = 0, "Upper(", "") & "A.简码" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & _
                " Or " & IIf(chkCase.Value = 0, "Upper(", "") & "B.简码" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & ") "
            strWhere = strWhere & " And (" & IIf(chkCase.Value = 0, "Upper(", "") & "B.拼音码" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & _
                " Or " & IIf(chkCase.Value = 0, "Upper(", "") & "B.五笔码" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & ") "
        End If
        
        str编码 = str编码 & "%"
        str名称 = str名称 & "%"
        str简码 = str简码 & "%"
    Else
        If Trim(txtEdit(0).Text) <> "" Then
            strWhere = strWhere & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "A.编码" & IIf(chkCase.Value = 0, ")", "") & " Like [2] "
        End If
        
        If Trim(txtEdit(1).Text) <> "" Then
            strWhereAlias = strWhereAlias & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "A.名称" & IIf(chkCase.Value = 0, ")", "") & " Like [3] "
            strWhere = strWhere & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "Nvl(B.名称, A.名称)" & IIf(chkCase.Value = 0, ")", "") & " Like [3] "
        End If
        
        If Trim(txtEdit(2).Text) <> "" Then
            strWhereAlias = strWhereAlias & " And (" & IIf(chkCase.Value = 0, "Upper(", "") & "A.简码" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & _
                " Or " & IIf(chkCase.Value = 0, "Upper(", "") & "B.简码" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & ") "
            strWhere = strWhere & " And (" & IIf(chkCase.Value = 0, "Upper(", "") & "B.拼音码" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & _
                " Or " & IIf(chkCase.Value = 0, "Upper(", "") & "B.五笔码" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & ") "
        End If
        
        str编码 = "%" & str编码 & "%"
        str名称 = "%" & str名称 & "%"
        str简码 = "%" & str简码 & "%"
    End If
    
    '得到SQL语句
    strAlaisSql = " (Select Distinct A.诊疗项目id, A.名称, A.简码 As 拼音码, B.简码 As 五笔码, A.简码 || '/' || B.简码 As 简码" & _
        " From 诊疗项目别名 A, 诊疗项目别名 B " & _
        " Where A.诊疗项目id = B.诊疗项目id And A.码类 = 1 And B.码类 = 2 " & strWhereAlias
    If chkAlias.Value = 1 Then
        '查找别名
        strAlaisSql = strAlaisSql & " And A.性质 = 9 And B.性质 = 9 ) B "
    Else
        '查找通用名
        strAlaisSql = strAlaisSql & " And A.性质 = 1 And B.性质 = 1 ) B "
    End If
    gstrSql = "select distinct A.ID,A.类别,Nvl(B.名称, A.名称) As 名称,A.编码,B.简码,C.名称 as 分类,A.分类ID,A.撤档时间 from (" & _
        strTable & ") A," & strAlaisSql & ",诊疗分类目录 C where A.分类id=c.id(+) And  A.ID=B.诊疗项目id(+) and C.名称 is not NULL And C.类型 In (4,5,6) " & strWhere
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(str分类ID), str编码, str名称, str简码)
    
    Me.MousePointer = 11
    zlControl.FormLock LvwMain.hWnd
    With LvwMain.ListItems
        .Clear
        i = 1
        Do Until rsTemp.EOF
            '得出正确的图标
            strWhere = "Item"
            If Not CDate(IIf(IsNull(rsTemp("撤档时间")), CDate("3000/1/1"), rsTemp("撤档时间"))) = CDate("3000/1/1") Then
                strWhere = strWhere & "No"
            End If
            '添加节点
            Set lst = .Add(, "C" & i, rsTemp("名称"), strWhere, strWhere)
            If InStr(strWhere, "No") > 0 Then lst.ForeColor = RGB(255, 0, 0)
            
            Dim lngCol  As Long
            Dim varValue As Variant
            '根据ListView的列名从数据库取数
            For lngCol = 2 To LvwMain.ColumnHeaders.Count
                varValue = rsTemp(LvwMain.ColumnHeaders(lngCol).Text).Value
                lst.SubItems(lngCol - 1) = IIf(IsNull(varValue), "", varValue)
                
                lst.ListSubItems(1).Tag = IIf(IsNull(rsTemp("分类ID")), "", rsTemp("分类ID"))
                lst.Tag = rsTemp("id")
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
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlControl.FormLock 0
    Me.MousePointer = 0
End Sub

Private Sub Form_Activate()
    Select Case Val(frmClinicLists.tvwClass.Tag)
    Case 0
        Me.Tag = ">='A'": Me.Caption = "诊疗项目查找..."
    Case 1
        Me.Tag = "='8'": Me.Caption = "中药配方查找..."
    Case 2
        Me.Tag = "='9'": Me.Caption = "成套方案查找..."
    End Select
End Sub

Private Sub Form_Load()
    Dim intSel As Integer
    
    RestoreWinState Me, App.ProductName
    
    intSel = Val(zlDatabase.GetPara("匹配方式", glngSys, 1054, 2))
    
    If intSel > 2 Or intSel < 0 Then intSel = 2
    optMode(intSel).Value = True
    
    intSel = Val(zlDatabase.GetPara("查找范围", glngSys, 1054, 0))
    If intSel > 3 Or intSel < 0 Then intSel = 0
    optScope(intSel).Value = True
    
    intSel = Val(zlDatabase.GetPara("区分大小写", glngSys, 1054, 0))
    If intSel > 1 Or intSel < 0 Then intSel = 0
    chkCase.Value = intSel
    
    intSel = Val(zlDatabase.GetPara("查找别名", glngSys, 1054, 0))
    If intSel > 1 Or intSel < 0 Then intSel = 0
    chkAlias.Value = intSel
    
    chkStop.Value = IIf(frmClinicLists.mnuViewStoped.Checked, 1, 0)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim intSel As Integer
    
    For intSel = 0 To 2
        If optMode(intSel).Value = True Then Exit For
    Next
    Call zlDatabase.SetPara("匹配方式", intSel, glngSys, 1054)
    For intSel = 0 To 3
        If intSel <> 1 Then
            If optScope(intSel).Value = True Then Exit For
        End If
    Next
    Call zlDatabase.SetPara("查找范围", intSel, glngSys, 1054)
    Call zlDatabase.SetPara("区分大小写", chkCase.Value, glngSys, 1054)
    Call zlDatabase.SetPara("查找别名", chkCase.Value, glngSys, 1054)
        
    SaveWinState Me, App.ProductName
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Dim sngLeft As Single
    
    LvwMain.Top = IIf(fra高级.Visible = True, fra高级.Top + fra高级.Height, fra高级.Top)
    LvwMain.Height = Me.ScaleHeight - LvwMain.Top - 120
    
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
    LvwMain.Width = cmdFind.Left - LvwMain.Left - 245
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        LvwMain.SortOrder = IIf(LvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        LvwMain.SortKey = mintColumn
        LvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub chkCase_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkStop_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub lvwMain_DblClick()
    If mblnItem = True Then Call cmdLocate_Click
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnItem = False
End Sub

Private Sub optMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optScope_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub


