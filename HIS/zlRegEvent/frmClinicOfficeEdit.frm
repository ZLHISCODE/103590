VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicOfficeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "门诊诊室设置"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6780
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClinicOfficeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra基本信息 
      Caption         =   "基本信息"
      Height          =   2085
      Left            =   90
      TabIndex        =   14
      Top             =   90
      Width           =   5085
      Begin VB.TextBox txtEdit 
         Height          =   350
         Index           =   0
         Left            =   660
         TabIndex        =   12
         Top             =   2310
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox chk是否 
         Caption         =   "Check1"
         Height          =   210
         Index           =   0
         Left            =   210
         TabIndex        =   13
         Top             =   2760
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txt编码 
         Height          =   350
         Left            =   660
         MaxLength       =   3
         TabIndex        =   0
         Top             =   330
         Width           =   1245
      End
      Begin VB.TextBox txt位置 
         Height          =   350
         Left            =   660
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1650
         Width           =   4335
      End
      Begin VB.TextBox txt名称 
         Height          =   350
         Left            =   660
         MaxLength       =   20
         TabIndex        =   1
         Top             =   765
         Width           =   4335
      End
      Begin VB.TextBox txt简码 
         Height          =   350
         Left            =   660
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1215
         Width           =   1245
      End
      Begin VB.ComboBox cboStationNo 
         Height          =   330
         Left            =   2790
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1225
         Width           =   2205
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "简码"
         Height          =   210
         Index           =   0
         Left            =   210
         TabIndex        =   21
         Top             =   2340
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lbl位置 
         AutoSize        =   -1  'True
         Caption         =   "位置"
         Height          =   210
         Left            =   210
         TabIndex        =   19
         Top             =   1720
         Width           =   420
      End
      Begin VB.Label lbl编码 
         AutoSize        =   -1  'True
         Caption         =   "编码"
         Height          =   210
         Left            =   210
         TabIndex        =   15
         Top             =   400
         Width           =   420
      End
      Begin VB.Label lbl名称 
         AutoSize        =   -1  'True
         Caption         =   "名称"
         Height          =   210
         Left            =   210
         TabIndex        =   16
         Top             =   835
         Width           =   420
      End
      Begin VB.Label lbl简码 
         AutoSize        =   -1  'True
         Caption         =   "简码"
         Height          =   210
         Left            =   210
         TabIndex        =   17
         Top             =   1285
         Width           =   420
      End
      Begin VB.Label lblStationNo 
         AutoSize        =   -1  'True
         Caption         =   "站点"
         Height          =   210
         Left            =   2340
         TabIndex        =   18
         Top             =   1285
         Width           =   420
      End
   End
   Begin VB.Frame fraDept 
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   90
      TabIndex        =   22
      Top             =   2250
      Width           =   5115
      Begin VB.TextBox txtSelect 
         Height          =   350
         Left            =   870
         TabIndex        =   5
         Top             =   30
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "…"
         Height          =   345
         Left            =   2820
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   30
         Width           =   345
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "移除(&D)"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4170
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   30
         Width           =   915
      End
      Begin MSComctlLib.ListView lvwDept 
         Height          =   2145
         Left            =   30
         TabIndex        =   8
         Top             =   405
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   3784
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "适用科室"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label lbl适用科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "适用科室"
         Height          =   210
         Left            =   0
         TabIndex        =   23
         Top             =   105
         Width           =   840
      End
   End
   Begin VB.Frame frmSplit 
      Height          =   5205
      Left            =   5220
      TabIndex        =   20
      Top             =   -150
      Width           =   30
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   360
      Left            =   5460
      TabIndex        =   9
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   5460
      TabIndex        =   10
      Top             =   690
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   360
      Left            =   5460
      TabIndex        =   11
      Top             =   4290
      Width           =   1100
   End
End
Attribute VB_Name = "frmClinicOfficeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytFun As G_Enum_Fun '0-查看,1-添加,2-调整,3-删除
Private mlngID As Long '门诊诊室ID
Private mrs科室 As ADODB.Recordset

Private mblnOK As Boolean
Private mstrAddNewItem As String

Public Function ShowMe(frmParent As Form, ByVal bytFun As G_Enum_Fun, _
    Optional ByVal lngID As Long, Optional ByRef strAddNewItem As String) As Boolean
    '程序入口
    '入参：
    '   frmParent - 父窗口
    '   bytFun - 操作类型, 0-查看，1-新增，2-修改，3-删除
    '出参：
    '   strAddNewItem:新增诊室名称
    mbytFun = bytFun: mlngID = lngID
    mstrAddNewItem = ""
    
    Err = 0: On Error Resume Next
    mblnOK = False
    Me.Show 1, frmParent
    If mblnOK Then strAddNewItem = mstrAddNewItem
    ShowMe = mblnOK
End Function

Private Sub cboStationNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk是否_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdAdd_Click()
    Call SelectDept(True)
End Sub

Private Sub SelectDept(ByVal blnButton As Boolean, Optional strLike As String)
    '弹出选择器，选择使用科室
    Dim strSQL As String, rsResult As ADODB.Recordset
    Dim strID As String, str名称 As String
    Dim i As Integer, vRect As RECT
    Dim blnCancel As Boolean, strIDs As String
    Dim objItem As ListItem
    
    Err = 0: On Error GoTo ErrHandler
    For i = 1 To lvwDept.ListItems.Count
        strIDs = strIDs & "," & Val(Mid(lvwDept.ListItems(i).Key, 2))
    Next
    If strIDs <> "" Then strIDs = Mid(strIDs, 2)
    
    strSQL = "Select a.ID, a.编码, a.名称, Upper(a.简码) as 简码" & vbNewLine & _
            " From 部门表 A,部门性质说明 B" & vbNewLine & _
            " Where a.ID=b.部门ID " & vbNewLine & _
            "       And (b.服务对象=1 Or b.服务对象=3) And b.工作性质 = '临床'" & vbNewLine & _
            "       And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & vbNewLine
    If blnButton = False Then
        '模糊查找
        strSQL = strSQL & _
            "       And (a.编码 Like [1] Or a.名称 Like [1] Or Upper(a.简码) Like Upper([1]))" & vbNewLine
    End If
    If strIDs <> "" Then
        '排除已选择科室
        strSQL = strSQL & _
            "       And a.ID Not In(Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)))" & vbNewLine
    End If
    strSQL = strSQL & " Order By a.名称"
    vRect = zlControl.GetControlRect(txtSelect.Hwnd)
    Set rsResult = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "科室", False, "", "", False, False, IIf(blnButton = False, True, False), _
        vRect.Left, vRect.Top, txtSelect.Height, blnCancel, True, False, strLike & "%", strIDs)
    If blnCancel Then Exit Sub
    If rsResult Is Nothing Then Exit Sub
    If rsResult.EOF Then Exit Sub
    
    Do While Not rsResult.EOF
        strID = Nvl(rsResult!id): str名称 = Nvl(rsResult!名称)
        For i = 1 To lvwDept.ListItems.Count
            If Mid(lvwDept.ListItems(i).Key, 2) = strID Then Exit Sub
        Next
        Set objItem = lvwDept.ListItems.Add(, "K" & strID, str名称)
        rsResult.MoveNext
    Loop
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdRemove_Click()
    Err = 0: On Error GoTo ErrHandler
    If lvwDept.SelectedItem Is Nothing Then Exit Sub
    
    lvwDept.ListItems.Remove lvwDept.SelectedItem.Key
    If lvwDept.ListItems.Count > 0 Then
        lvwDept.ListItems(1).Selected = True
    End If
    
    If lvwDept.SelectedItem Is Nothing Then cmdRemove.Enabled = False: Exit Sub
    Call lvwDept_ItemClick(lvwDept.SelectedItem)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Activate()
    If Me.ActiveControl Is txt编码 And txt编码.Text <> "" Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo ErrHandler
    
    Me.Caption = Choose(mbytFun + 1, "查看", "新增", "修改", "删除") & "门诊诊室"
    If InitFaceEx() = False Then Unload Me: Exit Sub
    If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
        If InitData() = False Then Unload Me: Exit Sub
    End If
    If mbytFun = Fun_Add Then
        txt编码.Text = GetMaxLocalCode("门诊诊室")
        Exit Sub
    End If
    
    Select Case mbytFun
    Case Fun_View
        cmdCancel.Visible = False
        cmdOk.Left = cmdCancel.Left
        Call SetEnabled(Me.Controls, False)
    Case Fun_Update
        txt编码.Enabled = False
    End Select
    If LoadData(mlngID) = False Then Unload Me: Exit Sub
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitData() As Boolean
    Dim i As Long, strSQL As String, rsTemp As ADODB.Recordset
    Dim intRow As Integer, intCol As Integer
    
    Err = 0: On Error GoTo ErrHandler
    '加载站点数据
    strSQL = "Select 编号, 名称 From Zlnodelist"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboStationNo.Clear
    cboStationNo.AddItem ""
    Do While Not rsTemp.EOF
        cboStationNo.AddItem Nvl(rsTemp!编号) & "-" & Nvl(rsTemp!名称)
        If gstrNodeNo = Nvl(rsTemp!编号) Then cboStationNo.ListIndex = cboStationNo.NewIndex
        rsTemp.MoveNext
    Loop
    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadData(ByVal lngID As Long) As Boolean
    '加载数据
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objItem As Field, Index As Integer, i As Integer
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select b.编号 As 站点编号, a.*" & vbNewLine & _
            " From 门诊诊室 A,Zlnodelist B" & vbNewLine & _
            " Where a.站点=b.名称(+) And ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    If rsTemp.EOF Then Exit Function
    
    txt编码.Text = Nvl(rsTemp!编码)
    txt名称.Text = Nvl(rsTemp!名称)
    txt简码.Text = Nvl(rsTemp!简码)
    txt位置.Text = Nvl(rsTemp!位置)
    zlControl.CboSetText cboStationNo, Nvl(rsTemp!站点), False
    If cboStationNo.ListIndex = -1 Then
        cboStationNo.AddItem Nvl(rsTemp!站点编号) & "-" & Nvl(rsTemp!站点)
        cboStationNo.ListIndex = cboStationNo.NewIndex
    End If
    
    '加载扩展字段值
    For Each objItem In rsTemp.Fields
        If InStr(",站点编号,ID,编码,名称,简码,位置,缺省标志,站点,", "," & UCase(objItem.Name) & ",") = 0 Then
            Index = -1
            If objItem.Name Like "是否*" Or (objItem.Type = adNumeric And objItem.Precision = 1) Then
                '字段名含“是否”、Numeric类型，宽度1B，用CheckBox表现
                For i = 1 To chk是否.UBound
                    If chk是否(i).Caption = objItem.Name Then Index = i: Exit For
                Next
                If Index > 0 Then
                    chk是否(Index).Value = IIf(Val(Nvl(objItem.Value)) = 0, vbUnchecked, vbChecked)
                End If
            Else
                For i = 1 To lblEdit.UBound
                    If lblEdit(i).Caption = objItem.Name Then Index = i: Exit For
                Next
                If Index > 0 Then
                    If Val(lblEdit(Index).Tag) = 2 Then '日期型
                        txtEdit(Index).Text = Format(Nvl(objItem.Value), "yyyy-mm-dd")
                    Else
                        txtEdit(Index).Text = Nvl(objItem.Value)
                    End If
                End If
            End If
        End If
    Next
    
    '适用科室
    lvwDept.ListItems.Clear
    strSQL = "Select b.Id, b.名称" & vbNewLine & _
            " From 门诊诊室适用科室 A, 部门表 B" & vbNewLine & _
            " Where a.科室id = b.Id And a.诊室id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    If rsTemp.EOF Then LoadData = True: Exit Function
    
    Do Until rsTemp.EOF
        lvwDept.ListItems.Add , "K" & Nvl(rsTemp!id), Nvl(rsTemp!名称)
        rsTemp.MoveNext
    Loop
        
    LoadData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_View Then Unload Me: Exit Sub
    
    cmdOk.Enabled = False
    If IsValied() = False Then cmdOk.Enabled = True: Exit Sub
    If SaveData() = False Then cmdOk.Enabled = True: Exit Sub
    
    mblnOK = True
    mstrAddNewItem = Trim(txt名称.Text)
    If mbytFun = Fun_Add Then
        Call ClearFaceInfor
        cmdOk.Enabled = True
        Exit Sub
    End If
    Unload Me
    Exit Sub
ErrHandler:
    cmdOk.Enabled = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearFaceInfor()
    '功能:清除界面信息，以便重新输入数据
    Dim i As Integer
    
    On Error GoTo errHandle
    txt编码.Text = GetMaxLocalCode("门诊诊室")
    txt名称.Text = ""
    txt简码.Text = ""
    txt位置.Text = ""
    txtSelect.Text = ""
    
    For i = 1 To txtEdit.UBound
        txtEdit(i).Text = ""
    Next
    
    For i = 1 To chk是否.UBound
        chk是否(i).Value = vbUnchecked
    Next
    
    lvwDept.ListItems.Clear
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String, i As Long
    Dim strTemp As String, str适用科室 As String
    
    Err = 0: On Error GoTo ErrHandler
    
    For i = 1 To lvwDept.ListItems.Count
        strTemp = Val(Mid(lvwDept.ListItems(i).Key, 2))
        str适用科室 = str适用科室 & ";" & strTemp
    Next
    If str适用科室 <> "" Then str适用科室 = Mid(str适用科室, 2)
    
    Select Case mbytFun
    Case Fun_Add
        'Zl_门诊诊室_Modify(
        strSQL = "Zl_门诊诊室_Modify("
        '操作类型_In Number,--0-新增，1-修改
        strSQL = strSQL & "" & 0 & ","
        'Id_In       门诊诊室.Id%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '编码_In     门诊诊室.编码%Type := Null,
        strSQL = strSQL & "'" & Trim(txt编码.Text) & "',"
        '名称_In     门诊诊室.名称%Type := Null,
        strSQL = strSQL & "'" & Trim(txt名称.Text) & "',"
        '简码_In     门诊诊室.简码%Type := Null,
        strSQL = strSQL & "'" & Trim(txt简码.Text) & "',"
        '位置_In     门诊诊室.位置%Type := Null,
        strSQL = strSQL & "'" & Trim(txt位置.Text) & "',"
        '站点_In     门诊诊室.站点%Type := Null,
        strSQL = strSQL & "'" & NeedCode(cboStationNo.Text) & "',"
        '适用科室_In Varchar2:=Null--格式：科室1;科室2;科室3;...
        strSQL = strSQL & "'" & str适用科室 & "',"
        '扩展_In Varchar2:=Null--用户扩展字段值，格式：字段名1=值1,字段名2=值2,...
        strSQL = strSQL & "'" & ExpandSaveStr() & "')"
    Case Fun_Update
        'Zl_门诊诊室_Modify(
        strSQL = "Zl_门诊诊室_Modify("
        '操作类型_In Number,--0-新增，1-修改
        strSQL = strSQL & "" & 1 & ","
        'Id_In       门诊诊室.Id%Type,
        strSQL = strSQL & "" & mlngID & ","
        '编码_In     门诊诊室.编码%Type := Null,
        strSQL = strSQL & "'" & Trim(txt编码.Text) & "',"
        '名称_In     门诊诊室.名称%Type := Null,
        strSQL = strSQL & "'" & Trim(txt名称.Text) & "',"
        '简码_In     门诊诊室.简码%Type := Null,
        strSQL = strSQL & "'" & Trim(txt简码.Text) & "',"
        '位置_In     门诊诊室.位置%Type := Null,
        strSQL = strSQL & "'" & Trim(txt位置.Text) & "',"
        '站点_In     门诊诊室.站点%Type := Null,
        strSQL = strSQL & "'" & NeedCode(cboStationNo.Text) & "',"
        '适用科室_In Varchar2:=Null--格式：科室1;科室2;科室3;...
        strSQL = strSQL & "'" & str适用科室 & "',"
        '扩展_In Varchar2:=Null--用户扩展字段值，格式：字段名1=值1,字段名2=值2,...
        strSQL = strSQL & "'" & ExpandSaveStr() & "')"
    End Select
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValied() As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    If zlControl.FormCheckInput(Me) = False Then Exit Function
    If zlControl.TxtCheckInput(txt编码, "编码", 3, False) = False Then Exit Function
    If zlControl.TxtCheckInput(txt名称, "名称", 20, False) = False Then Exit Function
    If zlControl.TxtCheckInput(txt简码, "简码", 6, False) = False Then Exit Function
    If zlControl.TxtCheckInput(txt位置, "位置", 40, True) = False Then Exit Function
    
    If IsValidEx() = False Then Exit Function
    
    If mbytFun = Fun_Add Then
        strSQL = "Select 1 From 门诊诊室 Where 名称 = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(txt名称.Text))
        If Not rsTemp Is Nothing Then
            If Not rsTemp.EOF Then
                MsgBox Trim(txt名称.Text) & " 已存在！", vbInformation, gstrSysName
                If txt名称.Visible And txt名称.Enabled Then txt名称.SetFocus
                zlControl.TxtSelAll txt名称
                Exit Function
            End If
        End If
    ElseIf mbytFun = Fun_Update Then
        strSQL = "Select 1 From 门诊诊室 Where 名称 = [1] And ID <> [2] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(txt名称.Text), mlngID)
        If Not rsTemp Is Nothing Then
            If Not rsTemp.EOF Then
                MsgBox Trim(txt名称.Text) & " 已存在！", vbInformation, gstrSysName
                If txt名称.Visible And txt名称.Enabled Then txt名称.SetFocus
                zlControl.TxtSelAll txt名称
                Exit Function
            End If
        End If
    End If
    IsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not mrs科室 Is Nothing Then Set mrs科室 = Nothing
End Sub

Private Sub lvwDept_GotFocus()
    cmdRemove.Enabled = Not lvwDept.SelectedItem Is Nothing
    If lvwDept.ListItems.Count = 0 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub lvwDept_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwDept.SelectedItem Is Nothing Then Exit Sub
    cmdRemove.Enabled = True
End Sub

Private Sub lvwDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtSelect_GotFocus()
    zlControl.TxtSelAll txtSelect
End Sub

Private Sub txtSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtSelect.Text) = "" Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        Call SelectDept(False, Trim(txtSelect.Text))
        zlControl.TxtSelAll txtSelect
    End If
End Sub

Private Sub txt编码_GotFocus()
    zlControl.TxtSelAll txt编码
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt简码_GotFocus()
    zlControl.TxtSelAll txt简码
End Sub

Private Sub txt简码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt名称_Change()
    txt简码.Text = zlCommFun.SpellCode(txt名称.Text)
End Sub

Private Sub txt名称_GotFocus()
    zlControl.TxtSelAll txt名称
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(txt名称.Text) = "" Then
            MsgBox "名称不能为空！", vbInformation, gstrSysName
            txt名称.SetFocus: Exit Sub
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt位置_GotFocus()
    zlControl.TxtSelAll txt位置
End Sub

Private Sub txt位置_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Function InitFaceEx() As Boolean
    '初始化界面，动态加载用户扩展字段,113315
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim Index As Integer, i As Integer
    Dim objItem As Field
    Dim intTabIndex As Integer
    Dim sngAddHeight As Single
    Dim sngTop As Single, sngSplit As Single
    
    On Error GoTo errHandle
    strSQL = "Select * From 门诊诊室 Where 1 = 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "读取门诊诊室表结构")
    
    For Each objItem In rsTemp.Fields
        If InStr(",ID,编码,名称,简码,位置,缺省标志,站点,", "," & UCase(objItem.Name) & ",") = 0 Then
            If objItem.Name Like "是否*" Or (objItem.Type = adNumeric And objItem.Precision = 1) Then
                '字段名含“是否”、Numeric类型，宽度1B，用CheckBox表现
                Index = chk是否.Count
                
                Load chk是否(Index): Set chk是否(Index).Container = fra基本信息
                chk是否(Index).Visible = True
                chk是否(Index).Caption = objItem.Name
                
                chk是否(Index).Width = 300 + Me.TextWidth(chk是否(Index).Caption)
                If chk是否(Index).Left + chk是否(Index).Width - 100 > fra基本信息.Width Then
                    chk是否(Index).Width = fra基本信息.Width - chk是否(Index).Left - 100
                End If
            Else
                Index = lblEdit.Count
                
                Load lblEdit(Index): Set lblEdit(Index).Container = fra基本信息
                Load txtEdit(Index): Set txtEdit(Index).Container = fra基本信息
                lblEdit(Index).Visible = True
                txtEdit(Index).Visible = True
                lblEdit(Index).Caption = objItem.Name
                
                '字段类型,为1表示数字型,2表示日期
                If objItem.Type = adNumeric Then
                    lblEdit(Index).Tag = 1
                    txtEdit(Index).MaxLength = objItem.Precision
                ElseIf objItem.Type = adDate Or objItem.Type = adDBTimeStamp _
                    Or objItem.Type = adDBDate Or objItem.Type = adDBTime Then
                    lblEdit(Index).Tag = 2
                    txtEdit(Index).MaxLength = 10
                Else
                    lblEdit(Index).Tag = 3
                    txtEdit(Index).MaxLength = objItem.DefinedSize
                End If
                
                txtEdit(Index).Left = lblEdit(Index).Left + lblEdit(Index).Width + 30
                txtEdit(Index).Width = fra基本信息.Width - txtEdit(Index).Left - 90
            End If
        End If
    Next
    
    sngTop = txt位置.Top + txt位置.Height
    intTabIndex = txt位置.TabIndex + 1
    sngAddHeight = 0
    sngSplit = 85
    
    For i = 1 To lblEdit.UBound
        txtEdit(i).Top = sngTop + sngSplit
        lblEdit(i).Top = txtEdit(i).Top + 70
        txtEdit(i).TabIndex = intTabIndex '设置Tab顺序
        
        sngTop = txtEdit(i).Top + txtEdit(i).Height
        intTabIndex = intTabIndex + 1
        sngAddHeight = sngAddHeight + txtEdit(i).Height + sngSplit
    Next
    
    For i = 1 To chk是否.UBound
        chk是否(i).Top = sngTop + sngSplit
        chk是否(i).TabIndex = intTabIndex '设置Tab顺序
        
        sngTop = chk是否(i).Top + chk是否(i).Height
        intTabIndex = intTabIndex + 1
        sngAddHeight = sngAddHeight + chk是否(i).Height + sngSplit
    Next
    
    fra基本信息.Height = fra基本信息.Height + sngAddHeight
    fraDept.Top = fraDept.Top + sngAddHeight
    frmSplit.Height = frmSplit.Height + sngAddHeight
    
    Me.Height = Me.Height + sngAddHeight
    cmdHelp.Top = Me.ScaleHeight - cmdHelp.Height - 200
    
    InitFaceEx = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValidEx() As Boolean
    '分析用户扩展字段所输入的内容是否有效,113315
    Dim i As Integer
    Dim strTemp As String
    
    On Error GoTo errHandle
    For i = 1 To lblEdit.UBound
        strTemp = Trim(txtEdit(i).Text)
        If zlCommFun.StrIsValid(strTemp, txtEdit(i).MaxLength, txtEdit(i).Hwnd) = False Then
            zlControl.TxtSelAll txtEdit(i)
            Exit Function
        End If
        
        Select Case Val(lblEdit(i).Tag)
        Case 1 '数字型字段
            If strTemp <> "" And Not IsNumeric(strTemp) Then
                MsgBox lblEdit(i).Caption & "应该输入数字。", vbExclamation, gstrSysName
                zlControl.TxtSelAll txtEdit(i)
                If txtEdit(i).Visible And txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
        Case 2 '日期型字段
            strTemp = zlCommFun.AddDate(strTemp)
            If strTemp <> "" Then
                If Not IsDate(strTemp) Then
                    MsgBox lblEdit(i).Caption & "不是有效的日期格式(yyyy-mm-dd)或(yyyymmdd)。", vbExclamation, gstrSysName
                    zlControl.TxtSelAll txtEdit(i)
                    If txtEdit(i).Visible And txtEdit(i).Enabled Then txtEdit(i).SetFocus
                    Exit Function
                End If
                
                Err = 0: On Error Resume Next
                strTemp = Format(strTemp, "yyyy-mm-dd")
                If Err <> 0 Or Not IsDate(strTemp) Then
                    MsgBox lblEdit(i).Caption & "不是有效的日期格式(yyyy-mm-dd)或(yyyymmdd)。", vbExclamation, gstrSysName
                    zlControl.TxtSelAll txtEdit(i)
                    If txtEdit(i).Visible And txtEdit(i).Enabled Then txtEdit(i).SetFocus
                    Exit Function
                End If
                
                txtEdit(i).Text = strTemp
            End If
        End Select
    Next
    IsValidEx = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExpandSaveStr() As String
    '获取扩展字段保存字符串,113315
    Dim strSQL As String
    Dim i As Integer
    
    On Error GoTo errHandle
    For i = 1 To lblEdit.UBound
        strSQL = strSQL & "," & lblEdit(i).Caption & "="
        Select Case Val(lblEdit(i).Tag)
        Case 1 '数值型
            strSQL = strSQL & Val(txtEdit(i).Text)
        Case 2   '日期型
            strSQL = strSQL & "To_Date('" & Format(Trim(txtEdit(i).Text), "yyyy-mm-dd") & "','yyyy-mm-dd')"
        Case Else
            strSQL = strSQL & "'" & Trim(txtEdit(i).Text) & "'"
        End Select
    Next
    
    For i = 1 To chk是否.UBound
        strSQL = strSQL & "," & chk是否(i).Caption & "=" & IIf(chk是否(i).Value = 1, "1", "0")
    Next
    If strSQL <> "" Then strSQL = Mid(strSQL, 2)
    
    ExpandSaveStr = Replace(strSQL, "'", "''")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
