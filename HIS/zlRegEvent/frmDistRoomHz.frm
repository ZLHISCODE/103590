VERSION 5.00
Begin VB.Form frmDistRoomHz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "回诊病人签到"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5490
   Icon            =   "frmDistRoomHz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5490
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2595
      TabIndex        =   8
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3810
      TabIndex        =   7
      Top             =   2865
      Width           =   1100
   End
   Begin VB.ComboBox cbo科室 
      Height          =   300
      ItemData        =   "frmDistRoomHz.frx":058A
      Left            =   2055
      List            =   "frmDistRoomHz.frx":058C
      TabIndex        =   6
      Top             =   1125
      Width           =   2025
   End
   Begin VB.ComboBox cbo诊室 
      Height          =   300
      Left            =   2055
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2010
      Width           =   2025
   End
   Begin VB.ComboBox cbo医生 
      Height          =   300
      Left            =   2055
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   3
      Top             =   2700
      Width           =   6900
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   5490
      TabIndex        =   0
      Top             =   0
      Width           =   5490
      Begin VB.Image Image1 
         Height          =   720
         Left            =   4350
         Picture         =   "frmDistRoomHz.frx":058E
         Top             =   45
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "回诊信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   2
         Top             =   135
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请指定病人所要回诊到的目标科室等信息。"
         Height          =   180
         Left            =   600
         TabIndex        =   1
         Top             =   390
         Width           =   3420
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -45
         X2              =   5500
         Y1              =   765
         Y2              =   765
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "回诊科室"
      Height          =   180
      Left            =   1275
      TabIndex        =   11
      Top             =   1185
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "回诊诊室"
      Height          =   180
      Left            =   1275
      TabIndex        =   10
      Top             =   2070
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "回诊医生"
      Height          =   180
      Left            =   1275
      TabIndex        =   9
      Top             =   1620
      Width           =   720
   End
End
Attribute VB_Name = "frmDistRoomHz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrNo As String
Private mlng科室ID As Long
Private mstr诊室 As String
Private mstr医生 As String
Private mlng医生ID As Long
Private mstr原诊室 As String
Private mlng原科室ID As Long
Private mstr原医生 As String
Private mlng挂号ID As Long
Private mrsDept As ADODB.Recordset
Private mlngPreDept As Long
Private mstrLike As String
Private mblnOk As Boolean
Private mlngModule As Long, mstrPrivs As String
Public Function ShowMe(frmParent As Object, ByVal lngModule As Long, strPrivs As String, ByVal strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:回诊签到
    '入参:strNO=要回诊的挂号单
    '出参:
    '返回:回诊成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-16 14:59:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrNo = strNO: mlngModule = lngModule: mstrPrivs = strPrivs: mblnOk = False
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Sub cbo科室_Click()
    If cbo科室.ListIndex <> -1 Then
        If mlngPreDept <> cbo科室.ItemData(cbo科室.ListIndex) Then
            mlngPreDept = cbo科室.ItemData(cbo科室.ListIndex)
            '读取该科室医生、诊室
            Call LoadDoctor
            Call LoadRoom
        End If
    Else
        mlngPreDept = 0
    End If
End Sub
Private Sub cbo科室_GotFocus()
    Call zlControl.TxtSelAll(cbo科室)
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    If cbo科室.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If mrsDept Is Nothing Then Exit Sub
    If zlSelectDept(Me, mlngModule, cbo科室, mrsDept, cbo科室.Text, True, "部门选择") = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo科室_Validate(Cancel As Boolean)
    Dim lngID As Long
    If cbo科室.ListIndex >= 0 Then Exit Sub
    lngID = mlngPreDept
   zlControl.CboLocate cbo科室, lngID, True
   If cbo科室.ListIndex < 0 And cbo科室.ListCount <> 0 Then cbo科室.ListIndex = 0
End Sub

Private Sub cbo医生_Click()
    Call LoadRoom
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Function Valied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的合法性
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-16 15:25:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnYes As Boolean
    On Error GoTo errHandle
    If cbo科室.ListIndex = -1 Then
        MsgBox "请确定要回诊的科室。", vbInformation, gstrSysName
        cbo科室.SetFocus: Exit Function
    End If
    If cbo诊室.ListIndex = -1 Then
        MsgBox "请确定要回诊的诊室。", vbInformation, gstrSysName
        cbo诊室.SetFocus: Exit Function
    End If
    If cbo科室.ItemData(cbo科室.ListIndex) <> mlng原科室ID And blnYes = False Then
        If MsgBox("注意:" & vbCrLf & "  你选择的科室与回诊的科室不一致,你是否要调整病人回诊科室?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            cbo科室.SetFocus: Exit Function
        End If
        blnYes = True
                
    End If
    If cbo诊室.Text <> mstr原诊室 And blnYes = False Then
        If MsgBox("注意:" & vbCrLf & "  你选择的诊室与回诊的诊室不一致,你是否要调整病人的回诊诊室?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            cbo诊室.SetFocus: Exit Function
        End If
        blnYes = True
    End If
    If NeedName(cbo医生.Text) <> mstr原医生 And blnYes = False Then
        If MsgBox("注意:" & vbCrLf & "  你选择的医生与回诊的医生不一致,你是否要调整病人的回诊医生?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        cbo医生.SetFocus: Exit Function
        End If
        blnYes = True
    End If
    Valied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOk_Click()
    Dim strSQL As String
    If Valied = False Then Exit Sub
    '返回数据
    mlng科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    mstr诊室 = cbo诊室.Text
    mstr医生 = NeedName(cbo医生.Text)
    If cbo医生.ListIndex <> -1 Then
        mlng医生ID = cbo医生.ItemData(cbo医生.ListIndex)
    End If
    'Zl_病人挂号记录_回诊
    strSQL = "Zl_病人挂号记录_回诊("
    '  Id_In         病人挂号记录.ID%Type,
    strSQL = strSQL & "" & mlng挂号ID & ","
    '  新执行科室_In 病人挂号记录.执行部门id%Type,
    strSQL = strSQL & "" & mlng科室ID & ","
    '  新诊室_In     病人挂号记录.诊室%Type,
    strSQL = strSQL & "'" & mstr诊室 & "',"
    '  新医生_In     病人挂号记录.执行人%Type,
    strSQL = strSQL & "'" & mstr医生 & "',"
    '  需回诊_In Integer:=0
    strSQL = strSQL & "0,"
    '预约方式
    strSQL = strSQL & "'" & zl_Get预约方式ByID(mlng挂号ID) & "')" '问题号:48350
    zlDatabase.ExecuteProcedure strSQL, Me.Caption '问题号:53508
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not Me.ActiveControl Is cbo科室 Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    mlng科室ID = 0
    mstr诊室 = ""
    mstr医生 = ""
    mlng医生ID = 0
    mblnOk = False
    mlngPreDept = 0
    mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "") '输入匹配方式
    
    On Error GoTo errH
    
    '原挂号相关信息
    strSQL = "Select ID, 执行部门ID,诊室,执行人 From 病人挂号记录 Where NO=[1] and 记录性质=1 and 记录状态=1 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNo)
    mstr原诊室 = Nvl(rsTmp!诊室)
    mlng原科室ID = rsTmp!执行部门id
    mstr原医生 = Nvl(rsTmp!执行人)
    mlng挂号ID = Val(Nvl(rsTmp!id))
    '读取门诊科室:缺省为本科室
    strSQL = "" & _
    " Select Distinct B.ID,B.编码,B.名称,B.简码,Decode(B.ID,[1],1,0) as 缺省" & _
    " From 部门表 B,部门性质说明 C" & _
    " Where B.ID=C.部门ID And C.服务对象 In(1,3) And C.工作性质='临床'" & _
    "       And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is Null)" & _
    "       And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & _
    " Order by B.编码"
    Set mrsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng原科室ID)
    Do While Not mrsDept.EOF
        cbo科室.AddItem mrsDept!编码 & "-" & mrsDept!名称
        cbo科室.ItemData(cbo科室.NewIndex) = mrsDept!id
        If Val(Nvl(mrsDept!缺省)) = 1 Then
            cbo科室.ListIndex = cbo科室.NewIndex '主动激活Click
            mlngPreDept = mrsDept!id
        End If
        mrsDept.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDoctor()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
                
    cbo医生.Clear
    If cbo科室.ListIndex = -1 Then Exit Sub
    
    strSQL = "" & _
    " Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
    " From 人员表 A,部门人员 B,人员性质说明 C" & _
    " Where A.ID=B.人员ID And A.ID=C.人员ID" & _
    "       And C.人员性质='医生' And B.部门ID=[1]" & _
    "       And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
    "       And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
    " Order by A.简码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbo科室.ItemData(cbo科室.ListIndex))
    
    cbo医生.AddItem ""
    Call zlControl.CboSetIndex(cbo医生.Hwnd, 0)
    Do While Not rsTmp.EOF
        cbo医生.AddItem Nvl(rsTmp!简码) & "-" & Nvl(rsTmp!姓名)
        cbo医生.ItemData(cbo医生.NewIndex) = rsTmp!id
        If Nvl(rsTmp!姓名) = mstr原医生 Then
            cbo医生.ListIndex = cbo医生.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadRoom()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, bln临床安排 As Boolean
    
    On Error GoTo errH
    
    cbo诊室.Clear
    If cbo科室.ListIndex = -1 Then Exit Sub
    
    bln临床安排 = False
    If gbytRegistMode = 1 Then
        If Sys.Currentdate >= gdatRegistTime Then bln临床安排 = True
    End If
    
    If bln临床安排 = False Then
        strSQL = _
            "Select Distinct 门诊诊室 As 名称" & vbNewLine & _
            "From 挂号安排诊室 A, 挂号安排 B" & vbNewLine & _
            "Where a.号表id = b.Id And b.科室id = [1] And Nvl(b.医生姓名,Nvl([2],'-')) = Nvl([2],'-')" & vbNewLine & _
            "Order By 名称"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbo科室.ItemData(cbo科室.ListIndex), NeedName(cbo医生.Text))
    Else
        strSQL = _
            " Select Distinct c.名称" & vbNewLine & _
            " From 临床出诊诊室记录 A, 临床出诊记录 B, 门诊诊室 C" & vbNewLine & _
            " Where a.记录id = b.Id And a.诊室id = c.Id And b.科室id+0 = [1]" & vbNewLine & _
            "       And Nvl(b.医生姓名,Nvl([2],'-')) = Nvl([2],'-')" & vbNewLine & _
            "       And b.出诊日期 Between Trunc(Sysdate) - 1 And Trunc(Sysdate)" & vbNewLine & _
            " Order By 名称"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbo科室.ItemData(cbo科室.ListIndex), NeedName(cbo医生.Text))
        If rsTmp.RecordCount = 0 Then
            '重新从科室适用诊室中读取诊室,121589
            Set rsTmp = GetDoctorRooms(cbo科室.ItemData(cbo科室.ListIndex))
        End If
    End If
    
    cbo诊室.AddItem ""
    Call zlControl.CboSetIndex(cbo诊室.Hwnd, 0)
    Do While Not rsTmp.EOF
        cbo诊室.AddItem rsTmp!名称
        If cbo科室.ItemData(cbo科室.ListIndex) = mlng原科室ID And rsTmp!名称 = mstr原诊室 Then
            cbo诊室.ListIndex = cbo诊室.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot优先级 As Boolean = False, Optional str所有部门 As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:部门选择器
    '入参:cboDept-指定的部门部件
    '     rsDept-指定的部门
    '     strSearch-要搜索的串
    '     blnNot优先级-是否存在优先级字段
    '     str所有部门-所有部门名称
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 10:20:11
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim strIDs As String, str简码 As String
    
    '先复制记录集
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str所有部门 <> "" Then
        str简码 = zlCommFun.SpellCode(str所有部门)
        If intInputType = 1 Then
            If Trim(str简码) Like strCompents Then
                rsTemp.AddNew
                rsTemp!id = -1
                rsTemp!编码 = "-"
                rsTemp!名称 = str所有部门
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str简码) Like strCompents Or UCase(str所有部门) Like strCompents Then
                rsTemp.AddNew
                rsTemp!id = -1
                rsTemp!编码 = "-"
                rsTemp!名称 = str所有部门
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编码) = strSearch Then lngDeptID = Nvl(!id): iCount = 0: Exit Do
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编码)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!id))
                    iCount = iCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Nvl(!编码) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!id)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!id)) & ","
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!id))   '可能存在多个相同简码
                    iCount = iCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!id)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!id)) & ","
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有LXH01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编码) = strSearch Or Trim(!简码) = strSearch Or UCase(Trim(!名称)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!id))   '可能存在多个相同的多个
                    iCount = iCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If UCase(Trim(!编码)) Like strSearch & "*" Or Trim(Nvl(!简码)) Like strCompents Or UCase(Trim(Nvl(!名称))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!id)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!id)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then lngDeptID = Nvl(rsTemp!id)
        
    '刘兴洪:直接定位
    If lngDeptID <> 0 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '需要检查是否有多条满足条件的记录
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编码"
    Case 1 '输入全拼音
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "简码"
    Case Else
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编码"
    End Select
    
    '弹出选择器
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "缺省," & IIf(blnNot优先级, "", ",优先级") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!id))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.CboLocate cboDept, lngDeptID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlSelectDept = True
    Exit Function
GoNotSel:
    '未找到
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboDept
End Function


