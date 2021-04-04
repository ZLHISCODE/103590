VERSION 5.00
Begin VB.Form frmBedEdit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraEdit 
      Height          =   4185
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4830
      Begin VB.ComboBox cbo等级 
         Height          =   315
         Left            =   975
         TabIndex        =   7
         Text            =   "cbo等级"
         Top             =   3240
         Width           =   3660
      End
      Begin VB.TextBox txt顺序号 
         Height          =   300
         Left            =   960
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cbo科室 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2730
         Width           =   3660
      End
      Begin VB.CheckBox chkContAdd 
         Caption         =   "连续增加"
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txt床号 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   1000
         MaxLength       =   10
         TabIndex        =   0
         Top             =   280
         Width           =   1095
      End
      Begin VB.ComboBox cbo编制 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2220
         Width           =   1455
      End
      Begin VB.TextBox txt房间号 
         Height          =   300
         Left            =   975
         MaxLength       =   10
         TabIndex        =   2
         Top             =   735
         Width           =   1455
      End
      Begin VB.ComboBox cbo性别 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1710
         Width           =   1455
      End
      Begin VB.TextBox txt编制 
         Enabled         =   0   'False
         Height          =   300
         Left            =   975
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox CboLevel 
         Height          =   315
         Left            =   1905
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3255
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.Label lbl顺序号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "顺序号"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   1275
         Width           =   540
      End
      Begin VB.Label lbl等级 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床位等级"
         Height          =   180
         Left            =   195
         TabIndex        =   13
         Top             =   3300
         Width           =   720
      End
      Begin VB.Label lbl编制 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床位编制"
         Height          =   180
         Left            =   195
         TabIndex        =   14
         Top             =   2295
         Width           =   720
      End
      Begin VB.Label lbl房间号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "房间号"
         Height          =   180
         Left            =   375
         TabIndex        =   12
         Top             =   805
         Width           =   540
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "适用性别"
         Height          =   180
         Left            =   195
         TabIndex        =   11
         Top             =   1785
         Width           =   720
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "所属科室"
         Height          =   180
         Left            =   195
         TabIndex        =   10
         Top             =   2805
         Width           =   720
      End
      Begin VB.Label lbl床号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   180
         Left            =   555
         TabIndex        =   9
         Top             =   300
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmBedEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mblnAdd As Boolean '窗体编辑状态
Public mlngUnit As Long '当前病区ID
Public mobjSta As StatusBar
Public mblnChange As Boolean
Public mintCancle As Integer

Private mrs编制 As New ADODB.Recordset
Private mrsBedLevel As New ADODB.Recordset
Private mrptRecord As ReportRecord

Private Sub cbo编制_Click()
    Dim strTemp As String
    If cbo编制.Text = "" Then Exit Sub
    
    If cbo编制.ListIndex <> Val(cbo编制.Tag) Then
        cbo编制.Tag = cbo编制.ListIndex
        mblnChange = True
    End If

    strTemp = Split(cbo编制.Text, "-")(0)
    
    mrs编制.Filter = "编码='" & strTemp & "'"
    
    If Not mrs编制.EOF Then
        txt编制.Text = mrs编制!符号 & ""
    End If

    If mblnAdd = True Then
        txt床号.Text = NextBedNo(mlngUnit, zlCommFun.GetNeedName(cbo编制.Text), mrs编制!符号 & "")
    End If
End Sub

Private Sub cbo等级_Click()
    If cbo等级.ListIndex <> Val(cbo等级.Tag) Then
        cbo等级.Tag = cbo等级.ListIndex
        mblnChange = True
    End If
End Sub

Private Sub cbo等级_GotFocus()
    zlControl.TxtSelAll cbo等级
End Sub

Private Sub cbo等级_Validate(Cancel As Boolean)
    If isCheckBedLevelExists(cbo等级.Text, True, False) = False Then
        cbo等级.Text = ""
        cbo等级.ListIndex = -1
    End If
End Sub

Private Sub cbo科室_Click()
    If cbo科室.ListIndex <> Val(cbo科室.Tag) Then
        cbo科室.Tag = cbo科室.ListIndex
        mblnChange = True
    End If
End Sub

Private Sub cbo性别_Click()
    If cbo性别.ListIndex <> Val(cbo性别.Tag) Then
        cbo性别.Tag = cbo性别.ListIndex
        mblnChange = True
    End If
End Sub

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo性别.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo性别.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo性别.ListIndex = lngIdx
    ElseIf cbo性别.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo科室.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo科室.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo科室.ListIndex = lngIdx
    ElseIf cbo科室.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo编制_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo编制.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo编制.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo编制.ListIndex = lngIdx
    ElseIf cbo编制.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo等级_KeyPress(KeyAscii As Integer)
    '69273:刘鹏飞,2014-01-03,快速定位床位等级
    Dim lngIdx As Long
    Dim i As Long, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim rsTemp As ADODB.Recordset
    
    If KeyAscii <> 13 Then
'        If SendMessage(cbo等级.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
'        lngIdx = MatchIndex(cbo等级.hWnd, KeyAscii)
'        If lngIdx <> -2 Then cbo等级.ListIndex = lngIdx
    Else
        If cbo等级.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = UCase(cbo等级.Text)
        If cbo等级.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> cbo等级.List(cbo等级.ListIndex) Then Call cbo.SetIndex(cbo等级.hWnd, -1)
        End If
        If strText = "" Then
            cbo等级.ListIndex = -1
        ElseIf cbo等级.ListIndex = -1 Then
            strFilter = ""
            '先复制记录集
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrsBedLevel)
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrsBedLevel.Filter = strFilter: iCount = 0
            With mrsBedLevel
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrsBedLevel.EOF
                    Select Case intInputType
                    Case 0  '输入的是全数字
                        '如果输入的数字,需要检查:
                        '1.编号输入值相等,主要输入如:12 匹配000012这种情况
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                        
                        '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                        If Nvl(!编码) = strText Then strResult = Nvl(!名称): iCount = 0: Exit Do
                        
                        '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                        If Val(Nvl(!编码)) = Val(strText) Then
                            If iCount = 0 Then strResult = Nvl(!名称)
                            iCount = iCount + 1
                        End If
                        
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                         If Val(Nvl(!编码)) Like strText & "*" Then
                            If isCheckBedLevelExists(Nvl(!名称)) Then Call zlDatabase.zlInsertCurrRowData(mrsBedLevel, rsTemp)
                         End If
                    Case 1  '输入的是全字母
                        '规则:
                        ' 1.输入的简码相等,则直接定位
                        ' 2.根据参数来匹配相同数据
                        
                        '1.输入的简码相等,则直接定位
                        If Trim(Nvl(!简码)) = strText Then
                            If iCount = 0 Then strResult = Nvl(!名称)   '可能存在多个相同的多个
                            iCount = iCount + 1
                        End If
                        
                        '2.根据参数来匹配相同数据
                        If Trim(Nvl(!简码)) Like strCompents Then
                            If isCheckBedLevelExists(Nvl(!名称)) Then Call zlDatabase.zlInsertCurrRowData(mrsBedLevel, rsTemp)
                        End If
                    Case Else  ' 2-其他
                        '规则:可能存在汉字等情况,或编号类似于N001简码可能有ZYK01这种情况
                        '1.编码\简码相等,直接定位
                        '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                        
                        '1.编码\简码相等,直接定位
                        If Trim(!编码) = strText Or Trim(!简码) = strText Or Trim(!名称) = strText Then
                            If iCount = 0 Then strResult = Nvl(!名称)   '可能存在多个相同的多个
                            iCount = iCount + 1
                        End If
                        
                        '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                        If Trim(!编码) Like strText & "*" Or Trim(Nvl(!简码)) Like strCompents Or Trim(Nvl(!名称)) Like strCompents Then
                            If isCheckBedLevelExists(Nvl(!名称)) Then Call zlDatabase.zlInsertCurrRowData(mrsBedLevel, rsTemp)
                        End If
                    End Select
                    mrsBedLevel.MoveNext
                Loop
            End With
            If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!名称)
            '直接定位
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheckBedLevelExists(strResult, True) Then cbo等级.SetFocus:  zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '需要检查是否有多条满足条件的记录
            If rsTemp.RecordCount <> 0 Then
                '先按某种方式进行排序
                rsTemp.Sort = "简码,编码"
                '弹出选择器
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, 1130, cbo等级, rsTemp, True, "", "", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '进行定位
                            If isCheckBedLevelExists(Nvl(rsReturn!名称), True) Then
                                cbo等级.SetFocus
                                zlCommFun.PressKey vbKeyTab
                                Exit Sub
                            End If
                        End If
                    End If
                Else
                    cbo等级.SetFocus
                    Exit Sub
                End If
            Else
                '未找到
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: cbo等级.ListIndex = -1: zlControl.TxtSelAll cbo等级: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
        End If
        
        If cbo等级.ListIndex = -1 Then
            cbo等级.Text = ""
            Exit Sub
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Function isCheckBedLevelExists(ByVal str名称 As String, Optional blnLocateItem As Boolean = False, Optional ByVal blnLevel As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查名称是否在床位等级下拉列表中
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If blnLevel = True Then
        For i = 0 To CboLevel.ListCount - 1
            If CboLevel.List(i) = str名称 Then
                If blnLocateItem Then cbo等级.ListIndex = i
                isCheckBedLevelExists = True
                Exit Function
            End If
        Next
    Else
        For i = 0 To cbo等级.ListCount - 1
            If cbo等级.List(i) = str名称 Then
                If blnLocateItem Then cbo等级.ListIndex = i
                isCheckBedLevelExists = True
                Exit Function
            End If
        Next
    End If
End Function

Private Sub Form_Activate()
    On Error Resume Next
    Me.SetFocus
    If Err <> 0 Then Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnChange = False
    mintCancle = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mintCancle = Cancel
    If mblnAdd = False And mblnChange And Visible Then
        If MsgBox("你修改了的内容尚未保存,确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            mintCancle = 1: Cancel = 1: Exit Sub
        End If
    End If
    mblnAdd = False
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strTmp As String
    
    '性别分类
    cbo性别.Clear
    cbo性别.AddItem "1-男床"
    cbo性别.AddItem "2-女床"
    cbo性别.AddItem "3-不限床"
    If mblnAdd Then cbo性别.ListIndex = 2
    
    '确定病区的服务对象
    strSQL = "Select 服务对象 From 部门性质说明 Where 工作性质='护理' And 部门ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUnit)
    
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
        If mblnAdd And cbo科室.ListIndex = -1 Then cbo科室.ListIndex = 1
    Else
        MsgBox "未初始化临床科室或没有设置病区科室对应信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '69273:刘鹏飞,2014-01-03,提供床位登记的快速查找
    '床位等级
    strSQL = "Select ID,编码,名称,zlspellcode(名称,20) 简码 From 收费项目目录 Where 类别='J' And (撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL) Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo等级.Clear
    CboLevel.Clear: CboLevel.Visible = False
    Set mrsBedLevel = rsTmp.Clone
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo等级.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cbo等级.ItemData(i - 1) = rsTmp!ID
            CboLevel.AddItem rsTmp!名称
            CboLevel.ItemData(i - 1) = rsTmp!ID
            rsTmp.MoveNext
        Next
        If mblnAdd Then cbo等级.ListIndex = 0
    Else
        MsgBox "没有初始化床位等级信息,请先到床位等级设置中处理！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '床位编制
    strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省,符号 From  床位编制分类 Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
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
    txt床号.width = txt编制.Left + txt编制.width - txt床号.Left - 60
End Sub

Private Sub txt床号_GotFocus()
    zlControl.TxtSelAll txt床号
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
    zlControl.TxtSelAll txt房间号
End Sub

Private Sub txt房间号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Public Function zlEditStart(ByVal blnAdd As Boolean, ByVal lngUnitID As Long, Optional ByVal rptRecord As ReportRecord) As Boolean
    '功能：开始项目编辑
    '参数： blnAdd-是否增加，否则为修改
    '       lngItemId-增加的参照项目，或者指定编辑的项目
    
    Dim i As Integer
    Dim strTmp As String
    Dim rsTemp As New ADODB.Recordset, rsLength As New ADODB.Recordset
    Dim str编制 As String
    
    gblnOK = False
    
    mblnAdd = blnAdd
    mlngUnit = lngUnitID
    Set mrptRecord = rptRecord
    
    cbo科室.Tag = "-1"
    cbo性别.Tag = "-1"
    cbo等级.Tag = "-1"
    cbo编制.Tag = "-1"
    
    If chkContAdd.Value Then
        txt床号.Text = NextBedNo(mlngUnit, zlCommFun.GetNeedName(cbo编制.Text), txt编制.Text)
    Else
        If Not InitData Then Exit Function
        
        chkContAdd.Visible = blnAdd
        chkContAdd.Value = IIf(blnAdd, 1, 0)
        If blnAdd Then
            Me.Caption = "新增病床"
            txt床号.MaxLength = 10
            If cbo编制.Text <> "" Then str编制 = Split(cbo编制.Text, "-")(1)
            txt床号.Text = NextBedNo(mlngUnit, str编制, txt编制.Text)
        Else
            txt床号.Enabled = False
            
            With mrptRecord
    
                
                cbo科室.ListIndex = cbo.FindIndex(cbo科室, Val(.Item(mCol.科室ID).Value))
                If cbo科室.ListIndex = -1 Then
                    If .Item(mCol.科室ID).Value <> "" Then
                        cbo科室.ListIndex = 0
                    End If
                End If
                
                cbo性别.ListIndex = cbo.FindIndex(cbo性别, .Item(mCol.性别分类).Value, True)
                cbo等级.ListIndex = cbo.FindIndex(cbo等级, .Item(mCol.等级).Value, True)
                If cbo等级.ListIndex = -1 Then isCheckBedLevelExists .Item(mCol.等级).Value, True
                cbo编制.ListIndex = cbo.FindIndex(cbo编制, .Item(mCol.床位编制).Value, True)
                txt床号.MaxLength = 10
                txt床号.Text = .Item(mCol.床号).Value
                txt床号.width = TextWidth(txt床号.Text)
                txt房间号.Text = .Item(mCol.房间号).Value
                txt顺序号.Text = .Item(mCol.顺序号).Value
                txt编制.Text = ""
                
                
                '因为影响床位增减记录,禁止调整
                cbo编制.Enabled = False
            End With
            Me.Caption = "调整病床"
        End If
    End If
    
    mblnChange = False
    zlEditStart = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Dim objControl As Control
    For Each objControl In Me.Controls
        If objControl.Enabled = False Then objControl.Enabled = True
    Next
    Me.chkContAdd.Value = 0
    mblnChange = False
End Sub

Public Function zlEditSave() As String
    '功能：保存正在进行的编辑,并返回正在编辑床位号,保存失败返回空
    Dim strTmp As String, strSQL As String
    Dim objItem As ListItem
    Dim str床号 As String, lngDept As Long

    If mblnAdd = True Then
        If Not IsNumeric(txt床号.Text) Then
            MsgBox "床号必须输入！", vbInformation, gstrSysName
            txt床号.SetFocus: Exit Function
        End If
    End If

    If InStr(txt房间号.Text, "'") > 0 Then
        MsgBox "房间号中包含非法字符,请检查！", vbInformation, gstrSysName
        txt房间号.SetFocus: Exit Function
    End If

    If LenB(StrConv(txt房间号.Text, vbFromUnicode)) > 10 Then
        MsgBox "房间号的长度不能大于10！", vbInformation, gstrSysName
        txt房间号.SetFocus: Exit Function
    End If
    
    If InStr(Trim(txt顺序号.Text), ".") <> 0 Then
        If LenB(StrConv(txt顺序号.Text, vbFromUnicode)) > 10 Then
            MsgBox "顺序号的长度包含小数点在内不能大于10位！", vbInformation, gstrSysName
            txt顺序号.SetFocus: Exit Function
        End If
    Else
        If Len(Trim(txt顺序号.Text)) > 9 Then
            MsgBox "顺序号的长度除去小数点不能大于9位！", vbInformation, gstrSysName
            txt顺序号.SetFocus: Exit Function
        End If
    End If
    
    If InStr(Trim(txt顺序号.Text), ".") <> 0 Then
        If Len(Mid(Trim(txt顺序号.Text), InStr(Trim(txt顺序号.Text), ".") + 1)) > 1 Then
            MsgBox "顺序号只能有一位小数！", vbInformation, gstrSysName
            txt顺序号.SetFocus: Exit Function
        End If
    End If
    
    If cbo科室.ListIndex = -1 Then
        MsgBox "必须确定该病床所在科室！", vbInformation, gstrSysName
        cbo科室.SetFocus: Exit Function
    End If
    If cbo性别.ListIndex = -1 Then
        MsgBox "必须确定该病床的性别分类！", vbInformation, gstrSysName
        cbo性别.SetFocus: Exit Function
    End If
    If cbo等级.ListIndex = -1 Then
        MsgBox "必须确定该病床的等级！", vbInformation, gstrSysName
        cbo等级.SetFocus: Exit Function
    End If
    If cbo编制.ListIndex = -1 Then
        MsgBox "必须确定该病床的编制类型！", vbInformation, gstrSysName
        cbo编制.SetFocus: Exit Function
    End If

    mblnChange = False

    If mblnAdd = True Then
        str床号 = txt编制 & txt床号.Text
    Else
        str床号 = txt床号.Text
    End If
    lngDept = cbo科室.ItemData(cbo科室.ListIndex)

    If mblnAdd Then
        str床号 = Trim(str床号)
        strTmp = isRepeat(mlngUnit, "'" & str床号 & "'")
        If strTmp <> "" Then
            MsgBox "当前输入的床号已经存在！", vbInformation, gstrSysName
            txt床号.SetFocus: Exit Function
        End If

        gstrSQL = "zl_床位状况记录_INSERT('" & Trim(str床号) & "'," & mlngUnit & "," & _
            IIf(lngDept = 0, "NULL", lngDept) & "," & _
            "'" & txt房间号.Text & "'," & _
            IIf(cbo性别.ListIndex = -1, "NULL,", "'" & zlCommFun.GetNeedName(cbo性别.Text) & "',") & _
            IIf(cbo编制.ListIndex = -1, "NULL,", "'" & zlCommFun.GetNeedName(cbo编制.Text) & "',") & _
            IIf(cbo等级.ListIndex = -1, "NULL", cbo等级.ItemData(cbo等级.ListIndex)) & ",1" & ",'" & txt顺序号.Text & "')"
        
    Else
        str床号 = Trim(mrptRecord.Item(mCol.床号).Value)
        gstrSQL = "zl_床位状况记录_INSERT('" & Trim(str床号) & "'," & mlngUnit & "," & _
             IIf(lngDept = 0, "NULL", lngDept) & "," & _
             "'" & txt房间号.Text & "'," & _
             IIf(cbo性别.ListIndex = -1, "NULL,", "'" & zlCommFun.GetNeedName(cbo性别.Text) & "',") & _
             IIf(cbo编制.ListIndex = -1, "NULL,", "'" & zlCommFun.GetNeedName(cbo编制.Text) & "',") & _
             IIf(cbo等级.ListIndex = -1, "NULL", cbo等级.ItemData(cbo等级.ListIndex)) & ",0" & ",'" & txt顺序号.Text & "')"
    End If
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    zlEditSave = str床号
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt顺序号_Change()
    mblnChange = True
End Sub

Private Sub txt顺序号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    End If
End Sub

