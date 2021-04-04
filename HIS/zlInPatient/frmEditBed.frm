VERSION 5.00
Begin VB.Form frmEditBed 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病床"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   Icon            =   "frmEditBed.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   390
      TabIndex        =   8
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3735
      TabIndex        =   7
      Top             =   2055
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2550
      TabIndex        =   6
      Top             =   2055
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   4830
      Begin VB.TextBox txt床号 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   1000
         MaxLength       =   5
         TabIndex        =   0
         Top             =   280
         Width           =   1095
      End
      Begin VB.TextBox txt编制 
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cbo编制 
         Height          =   300
         Left            =   3345
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   645
         Width           =   1290
      End
      Begin VB.ComboBox cbo等级 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   3660
      End
      Begin VB.TextBox txt房间号 
         Height          =   300
         Left            =   3345
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1290
      End
      Begin VB.ComboBox cbo性别 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   645
         Width           =   1470
      End
      Begin VB.ComboBox cbo科室 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1065
         Width           =   3660
      End
      Begin VB.Label lbl编制 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床位编制"
         Height          =   180
         Left            =   2550
         TabIndex        =   15
         Top             =   705
         Width           =   720
      End
      Begin VB.Label lbl等级 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床位等级"
         Height          =   180
         Left            =   195
         TabIndex        =   14
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label lbl房间号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "房间号"
         Height          =   180
         Left            =   2730
         TabIndex        =   13
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "适用性别"
         Height          =   180
         Left            =   195
         TabIndex        =   12
         Top             =   705
         Width           =   720
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "所属科室"
         Height          =   180
         Left            =   195
         TabIndex        =   11
         Top             =   1125
         Width           =   720
      End
      Begin VB.Label lbl床号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   180
         Left            =   555
         TabIndex        =   10
         Top             =   300
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmEditBed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mblnModi As Boolean '窗体编辑状态(缺省新增)
Public mlngUnit As Long '当前病区ID
Public mlvwBeds As ListView
Public mobjSta As StatusBar
Public mblnChange As Boolean
Private mrs编制 As New ADODB.Recordset

Private Sub cbo编制_Click()
    Dim strTemp As String
    
    mblnChange = True

    strTemp = Split(cbo编制.Text, "-")(0)
    
    mrs编制.Filter = "编码=" & strTemp
    
    If Not mrs编制.EOF Then
        txt编制.Text = mrs编制!符号 & ""
    End If

    If mblnModi = False Then
        txt床号.Text = NextBedNo(mlngUnit, NeedName(cbo编制.Text), mrs编制!符号 & "")
    End If
End Sub

Private Sub cbo等级_Click()
    mblnChange = True
End Sub

Private Sub cbo科室_Click()
    mblnChange = True
End Sub

Private Sub cbo性别_Click()
    mblnChange = True
End Sub

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo性别.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = MatchIndex(cbo性别.hwnd, KeyAscii)
        If lngIdx <> -2 Then cbo性别.ListIndex = lngIdx
    ElseIf cbo性别.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo科室.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = MatchIndex(cbo科室.hwnd, KeyAscii)
        If lngIdx <> -2 Then cbo科室.ListIndex = lngIdx
    ElseIf cbo科室.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo编制_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo编制.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = MatchIndex(cbo编制.hwnd, KeyAscii)
        If lngIdx <> -2 Then cbo编制.ListIndex = lngIdx
    ElseIf cbo编制.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo等级_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo等级.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = MatchIndex(cbo等级.hwnd, KeyAscii)
        If lngIdx <> -2 Then cbo等级.ListIndex = lngIdx
    ElseIf cbo等级.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String, strSQL As String
    Dim objItem As ListItem
    Dim str床号 As String, lngDept As Long
    
    If mblnModi = False Then
        If Not IsNumeric(txt床号.Text) Then
            MsgBox "床号必须输入！", vbInformation, gstrSysName
            txt床号.SetFocus: Exit Sub
        End If
    End If
    
    If InStr(txt房间号.Text, "'") > 0 Then
        MsgBox "房间号中包含非法字符,请检查！", vbInformation, gstrSysName
        txt房间号.SetFocus: Exit Sub
    End If
    
    If LenB(StrConv(txt房间号.Text, vbFromUnicode)) > 10 Then
        MsgBox "房间号的长度不能大于10！", vbInformation, gstrSysName
        txt房间号.SetFocus: Exit Sub
    End If
    If cbo科室.ListIndex = -1 Then
        MsgBox "必须确定该病床所在科室！", vbInformation, gstrSysName
        cbo科室.SetFocus: Exit Sub
    End If
    If cbo性别.ListIndex = -1 Then
        MsgBox "必须确定该病床的性别分类！", vbInformation, gstrSysName
        cbo性别.SetFocus: Exit Sub
    End If
    If cbo等级.ListIndex = -1 Then
        MsgBox "必须确定该病床的等级！", vbInformation, gstrSysName
        cbo等级.SetFocus: Exit Sub
    End If
    If cbo编制.ListIndex = -1 Then
        MsgBox "必须确定该病床的编制类型！", vbInformation, gstrSysName
        cbo编制.SetFocus: Exit Sub
    End If
    
    mblnChange = False
    
    If mblnModi = False Then
        str床号 = txt编制 & txt床号.Text
    Else
        str床号 = txt床号.Text
    End If
    lngDept = cbo科室.ItemData(cbo科室.ListIndex)

    If mblnModi Then
        strSQL = "zl_床位状况记录_INSERT('" & Mid(mlvwBeds.SelectedItem.Key, 2) & "'," & mlngUnit & "," & _
            IIf(lngDept = 0, "NULL", lngDept) & "," & _
            "'" & txt房间号.Text & "'," & _
            IIf(cbo性别.ListIndex = -1, "NULL,", "'" & NeedName(cbo性别.Text) & "',") & _
            IIf(cbo编制.ListIndex = -1, "NULL,", "'" & NeedName(cbo编制.Text) & "',") & _
            IIf(cbo等级.ListIndex = -1, "NULL", cbo等级.ItemData(cbo等级.ListIndex)) & ",0)"
        On Error GoTo errH
            zldatabase.ExecuteProcedure strSQL, Me.Caption
        On Error GoTo 0
        
        Set objItem = mlvwBeds.SelectedItem
        objItem.SubItems(mlvwBeds.ColumnHeaders("_科室").Index - 1) = NeedName(cbo科室.Text)
        objItem.SubItems(mlvwBeds.ColumnHeaders("_房间号").Index - 1) = txt房间号.Text
        objItem.SubItems(mlvwBeds.ColumnHeaders("_性别分类").Index - 1) = NeedName(cbo性别.Text)
        objItem.SubItems(mlvwBeds.ColumnHeaders("_等级").Index - 1) = NeedName(cbo等级.Text)
        objItem.SubItems(mlvwBeds.ColumnHeaders("_床位编制").Index - 1) = NeedName(cbo编制.Text)
        If cbo性别.ListIndex = 0 Then
            objItem.Icon = "M_Empty"
            objItem.SmallIcon = "M_Empty"
        ElseIf cbo性别.ListIndex = 1 Then
            objItem.Icon = "F_Empty"
            objItem.SmallIcon = "F_Empty"
        Else
            objItem.Icon = "Empty"
            objItem.SmallIcon = "Empty"
        End If
        objItem.Tag = lngDept
        objItem.ListSubItems(1).Tag = ""
        If lngDept = 0 Then objItem.ListSubItems(1).Tag = 1 '共用病床
        
        Call SetBedIcon(mlvwBeds, objItem)
        
        objItem.EnsureVisible
        
        With objItem
            mobjSta.Panels(2) = "床号[" & Trim(.Text) & "]" & _
                " 状态:" & .SubItems(mlvwBeds.ColumnHeaders("_状态").Index - 1) & _
                " 性别分类:" & .SubItems(mlvwBeds.ColumnHeaders("_性别分类").Index - 1) & _
                " 科室:" & .SubItems(mlvwBeds.ColumnHeaders("_科室").Index - 1) & _
                " 等级:" & .SubItems(mlvwBeds.ColumnHeaders("_等级").Index - 1)
        End With
        gblnOK = True
        Unload Me
    Else
        strTmp = isRepeat(mlngUnit, "'" & str床号 & "'")
        If strTmp <> "" Then
            MsgBox "当前输入的床号已经存在！", vbInformation, gstrSysName
            txt床号.SetFocus: Exit Sub
        End If
        
        strSQL = "zl_床位状况记录_INSERT('" & str床号 & "'," & mlngUnit & "," & _
            IIf(lngDept = 0, "NULL", lngDept) & "," & _
            "'" & txt房间号.Text & "'," & _
            IIf(cbo性别.ListIndex = -1, "NULL,", "'" & NeedName(cbo性别.Text) & "',") & _
            IIf(cbo编制.ListIndex = -1, "NULL,", "'" & NeedName(cbo编制.Text) & "',") & _
            IIf(cbo等级.ListIndex = -1, "NULL", cbo等级.ItemData(cbo等级.ListIndex)) & ",1)"
        
        On Error GoTo errH
        zldatabase.ExecuteProcedure strSQL, Me.Caption
        On Error GoTo 0
        
        If cbo性别.ListIndex = 0 Then
            Set objItem = mlvwBeds.ListItems.Add(, "_" & str床号, str床号, "M_Empty", "M_Empty")
        ElseIf cbo性别.ListIndex = 1 Then
            Set objItem = mlvwBeds.ListItems.Add(, "_" & str床号, str床号, "F_Empty", "F_Empty")
        Else
            Set objItem = mlvwBeds.ListItems.Add(, "_" & str床号, str床号, "Empty", "Empty")
        End If
        
        objItem.SubItems(mlvwBeds.ColumnHeaders("_科室").Index - 1) = NeedName(cbo科室.Text)
        objItem.SubItems(mlvwBeds.ColumnHeaders("_房间号").Index - 1) = txt房间号.Text
        objItem.SubItems(mlvwBeds.ColumnHeaders("_状态").Index - 1) = "空床"
        objItem.SubItems(mlvwBeds.ColumnHeaders("_性别分类").Index - 1) = NeedName(cbo性别.Text)
        objItem.SubItems(mlvwBeds.ColumnHeaders("_等级").Index - 1) = NeedName(cbo等级.Text)
        objItem.SubItems(mlvwBeds.ColumnHeaders("_床位编制").Index - 1) = NeedName(cbo编制.Text)
        objItem.Tag = lngDept
        If lngDept = 0 Then objItem.ListSubItems(1).Tag = 1 '共用病床
        
        Call SetBedIcon(mlvwBeds, objItem)
        
        objItem.Selected = True
        objItem.EnsureVisible
        With objItem
            mobjSta.Panels(2) = "床号[" & Trim(.Text) & "]" & _
                " 状态:" & .SubItems(mlvwBeds.ColumnHeaders("_状态").Index - 1) & _
                " 性别分类:" & .SubItems(mlvwBeds.ColumnHeaders("_性别分类").Index - 1) & _
                " 科室:" & .SubItems(mlvwBeds.ColumnHeaders("_科室").Index - 1) & _
                " 等级:" & .SubItems(mlvwBeds.ColumnHeaders("_等级").Index - 1)
        End With
        
        Call frmManageBed.SetBedNOLen
        Call frmManageBed.SetMenuState
        
        txt床号.Text = NextBedNo(mlngUnit, NeedName(cbo编制.Text), txt编制.Text)
        
        gblnOK = True
        
        txt床号.SetFocus
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    mblnChange = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    Dim str编制 As String
    
    gblnOK = False
    
    If Not InitData Then Unload Me: Exit Sub
    
    If mblnModi Then
        txt床号.Enabled = False
        
        With mlvwBeds.SelectedItem

            
            cbo科室.ListIndex = FindCboIndex(cbo科室, Val(.Tag))
            If cbo科室.ListIndex = -1 Then
                If .SubItems(mlvwBeds.ColumnHeaders("_科室").Index - 1) <> "" Then
                    cbo科室.ListIndex = 0
                End If
            End If
            
            cbo性别.ListIndex = GetCboIndex(cbo性别, .SubItems(mlvwBeds.ColumnHeaders("_性别分类").Index - 1))
            cbo等级.ListIndex = GetCboIndex(cbo等级, .SubItems(mlvwBeds.ColumnHeaders("_等级").Index - 1))
            cbo编制.ListIndex = GetCboIndex(cbo编制, .SubItems(mlvwBeds.ColumnHeaders("_床位编制").Index - 1))
            txt床号.MaxLength = 10
            txt床号.Text = Mid(.Key, 2)
            txt床号.Width = TextWidth(txt床号.Text)
            txt房间号.Text = .SubItems(mlvwBeds.ColumnHeaders("_房间号").Index - 1)
            txt编制.Text = ""
            
            '因为影响床位增减记录,禁止调整
            cbo编制.Enabled = False
        End With
        Me.Caption = "调整病床"
    Else
        Me.Caption = "新增病床"
        txt床号.MaxLength = 5
        If cbo编制.Text <> "" Then str编制 = Split(cbo编制.Text, "-")(1)
        txt床号.Text = NextBedNo(mlngUnit, str编制, txt编制.Text)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnModi And mblnChange And Visible Then
        If MsgBox("你修改了的内容尚未保存,确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    mblnModi = False
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strTmp As String
    
    '性别分类
    cbo性别.AddItem "1-男床"
    cbo性别.AddItem "2-女床"
    cbo性别.AddItem "3-不限床"
    If Not mblnModi Then cbo性别.ListIndex = 2
    
    '确定病区的服务对象
    strSQL = "Select 服务对象 From 部门性质说明 Where 工作性质='护理' And 部门ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUnit)
    
    cbo科室.Clear
    
    If rsTmp!服务对象 = 1 Then
        '门诊观察室设置对应的门诊临床科室
        strTmp = "1,3"
    ElseIf rsTmp!服务对象 = 2 Then
        strTmp = "2,3"
    ElseIf rsTmp!服务对象 = 3 Then
        strTmp = "1,2,3"
    End If
    Set rsTmp = GetDeptOrUnit(0, mlngUnit, strTmp)
    
    If Not rsTmp.EOF Then
        cbo科室.AddItem "<共用病床>" '共用病床
        For i = 1 To rsTmp.RecordCount
            cbo科室.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Next
        If Not mblnModi And cbo科室.ListIndex = -1 Then cbo科室.ListIndex = 1
    Else
        MsgBox "未初始化临床科室或没有设置病区科室对应信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '床位等级
    strSQL = "Select ID as 序号,编码,名称 From 收费项目目录 Where 类别='J' And (撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL) Order by 编码"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo等级.Clear
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo等级.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cbo等级.ItemData(i - 1) = rsTmp!序号
            rsTmp.MoveNext
        Next
        If Not mblnModi Then cbo等级.ListIndex = 0
    Else
        MsgBox "没有初始化床位等级信息,请先到床位等级设置中处理！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '床位编制
    strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省,符号 From  床位编制分类 Order by 编码"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo编制.Clear
    Set mrs编制 = rsTmp.Clone
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo编制.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then cbo编制.ListIndex = cbo编制.NewIndex
            rsTmp.MoveNext
        Next
    Else
        MsgBox "没有初始化床位编制信息,请到字典管理中初始化床位编制分类！", vbInformation, gstrSysName
        Exit Function
    End If
    
    InitData = True
End Function

Private Sub txt编制_Change()

    txt床号.Left = txt编制.Left + TextWidth(txt编制.Text) + 60
    txt床号.Width = txt编制.Left + txt编制.Width - txt床号.Left - 60
End Sub

Private Sub txt床号_GotFocus()
    SelAll txt床号
End Sub

Private Sub txt床号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt床号.Text = "" Then
            Call Beep: Exit Sub
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub txt房间号_Change()
    mblnChange = True
End Sub

Private Sub txt房间号_GotFocus()
    SelAll txt房间号
End Sub

Private Sub txt房间号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub SetBedIcon(objLvw As Object, objItem As ListItem)
    If objItem.SubItems(objLvw.ColumnHeaders("_床位编制").Index - 1) = "加床" Then
        objItem.Icon = "加床_" & objItem.Icon
        objItem.SmallIcon = "加床_" & objItem.SmallIcon
    ElseIf objItem.SubItems(objLvw.ColumnHeaders("_床位编制").Index - 1) = "非编" Then
        objItem.Icon = "非编_" & objItem.Icon
        objItem.SmallIcon = "非编_" & objItem.SmallIcon
    End If
    
    If Val(objItem.ListSubItems(1).Tag) <> 0 Then
        objItem.Icon = "共用_" & objItem.Icon
        objItem.SmallIcon = "共用_" & objItem.SmallIcon
    End If
End Sub

