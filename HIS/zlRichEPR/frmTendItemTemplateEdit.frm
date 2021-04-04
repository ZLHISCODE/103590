VERSION 5.00
Begin VB.Form frmTendItemTemplateEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "护理项目模板编辑"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6840
   Icon            =   "frmTendItemTemplateEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5835
      Left            =   5340
      TabIndex        =   16
      Top             =   -270
      Width           =   45
   End
   Begin VB.ComboBox cbo科室 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5550
      TabIndex        =   15
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5550
      TabIndex        =   14
      Top             =   300
      Width           =   1100
   End
   Begin VB.ComboBox cbo适用护理等级 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   570
      Width           =   2655
   End
   Begin VB.TextBox txt模板名称 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   180
      Width           =   2655
   End
   Begin VB.PictureBox picCloumn 
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   60
      ScaleHeight     =   3405
      ScaleWidth      =   5205
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1260
      Width           =   5205
      Begin VB.CommandButton cmdMove 
         Caption         =   "下移(&D)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2130
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2310
         Width           =   975
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "上移(&U)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   2130
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2010
         Width           =   975
      End
      Begin VB.ListBox lstColumnItems 
         Height          =   2760
         Left            =   240
         TabIndex        =   8
         Top             =   465
         Width           =   1770
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "选用(&S)"
         Height          =   300
         Index           =   0
         Left            =   2130
         TabIndex        =   12
         Top             =   885
         Width           =   975
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "删除(&E)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2130
         TabIndex        =   13
         Top             =   1185
         Width           =   975
      End
      Begin VB.ListBox lstColumnUsed 
         Height          =   2760
         Left            =   3240
         TabIndex        =   9
         Top             =   450
         Width           =   1770
      End
      Begin VB.Label lblColumnItems 
         AutoSize        =   -1  'True
         Caption         =   "可选护理记录项目:"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   180
         Width           =   1530
      End
   End
   Begin VB.Label lbl科室 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "科室"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   630
      TabIndex        =   4
      Top             =   1020
      Width           =   360
   End
   Begin VB.Label lbl适用护理等级 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "护理等级"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   2
      Top             =   630
      Width           =   720
   End
   Begin VB.Label lbl模板名称 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "模板名称"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmTendItemTemplateEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Private mblnSel As String
Private mblnStart As Boolean
Private mlng科室ID As Long
Private mint护理等级 As Integer
Private mstr模板名称 As String
Private mblnEdit As Boolean

Public Function ShowEditor(ByVal objParent As Object, ByVal lng科室id As Long, ByVal str模板名称 As String, ByVal int护理等级 As Integer) As String
    On Error Resume Next
    mblnSel = ""
    mblnEdit = False
    mlng科室ID = lng科室id
    mstr模板名称 = str模板名称
    mint护理等级 = int护理等级
    Me.Show 1, objParent
    ShowEditor = mblnSel
End Function

Private Sub cbo科室_Click()
    Call cbo适用护理等级_Click
End Sub

Private Sub cbo适用护理等级_Click()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    If mblnStart = False Then Exit Sub
    
    gstrSQL = " Select A.项目序号,A.项目名称 From 护理记录项目 A" & _
              " Where A.应用方式<>0 " & IIf(cbo适用护理等级.ItemData(cbo适用护理等级.ListIndex) = -1, "", " And A.护理等级>=[2]") & _
              " And (A.适用科室=1 Or (A.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=A.项目序号 And D.科室id=[1])))" & _
              " MINUS " & _
              " Select B.项目序号,B.项目名称 From 护理项目模板 A,护理记录项目 B " & _
              " Where A.项目序号=B.项目序号 And A.科室ID =[1] And A.护理等级=[2]" & _
              " Order by 项目序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.cbo科室.ItemData(Me.cbo科室.ListIndex), Me.cbo适用护理等级.ItemData(Me.cbo适用护理等级.ListIndex))
    
    With rsTemp
        Me.lstColumnItems.Clear
        Do While Not .EOF
            Me.lstColumnItems.AddItem !项目名称
            Me.lstColumnItems.ItemData(Me.lstColumnItems.NewIndex) = !项目序号
            .MoveNext
        Loop
    End With
    
    '提取已选择的项目清单
    gstrSQL = " Select B.项目序号,B.项目名称 From 护理项目模板 A,护理记录项目 B " & _
              " Where A.项目序号=B.项目序号 And A.科室ID =[1] And A.护理等级=[2]" & _
              " Order by A.排列序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.cbo科室.ItemData(Me.cbo科室.ListIndex), Me.cbo适用护理等级.ItemData(Me.cbo适用护理等级.ListIndex))
    
    With rsTemp
        Me.lstColumnUsed.Clear
        Do While Not .EOF
            Me.lstColumnUsed.AddItem !项目名称
            Me.lstColumnUsed.ItemData(Me.lstColumnUsed.NewIndex) = !项目序号
            .MoveNext
        Loop
    End With

    cmdColumn(0).Enabled = (lstColumnItems.ListCount <> 0)
    cmdColumn(1).Enabled = (lstColumnUsed.ListCount <> 0)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdColumn_Click(Index As Integer)
    Dim intIndex As Integer
    Dim objlst As ListBox
    If Index = 0 Then
        If Me.lstColumnItems.ListIndex < 0 Then Exit Sub
        intIndex = Me.lstColumnItems.ListIndex
        Me.lstColumnUsed.AddItem Me.lstColumnItems.Text
        Me.lstColumnUsed.ItemData(Me.lstColumnUsed.NewIndex) = Me.lstColumnItems.ItemData(Me.lstColumnItems.ListIndex)
        Me.lstColumnItems.RemoveItem Me.lstColumnItems.ListIndex
        Set objlst = lstColumnItems
    Else
        If Me.lstColumnUsed.ListIndex < 0 Then Exit Sub
        intIndex = Me.lstColumnUsed.ListIndex
        Me.lstColumnItems.AddItem Me.lstColumnUsed.Text
        Me.lstColumnItems.ItemData(Me.lstColumnItems.NewIndex) = Me.lstColumnUsed.ItemData(Me.lstColumnUsed.ListIndex)
        Me.lstColumnUsed.RemoveItem Me.lstColumnUsed.ListIndex
        Set objlst = lstColumnUsed
    End If
    If objlst.ListCount >= intIndex + 1 Then
        objlst.ListIndex = intIndex
    Else
        objlst.ListIndex = objlst.ListCount - 1
    End If
    
    cmdColumn(0).Enabled = (lstColumnItems.ListCount <> 0)
    cmdColumn(1).Enabled = (lstColumnUsed.ListCount <> 0)
    mblnEdit = True
    
    Call SetMoveState
End Sub

Private Sub cmdMove_Click(Index As Integer)
    Dim arrData
    Dim strCopy As String
    Dim lngDo As Long, lngMax As Long
    Dim lngSelIndex As Long, lngTarIndex As Long
    
    '当前索引
    lngSelIndex = lstColumnUsed.ListIndex
    '目标索引
    lngTarIndex = lngSelIndex + IIf(Index = 0, -1, 1)
    lngMax = lstColumnUsed.ListCount - 1
    For lngDo = 0 To lngMax
        If lngDo = lngTarIndex Then
            strCopy = strCopy & "|" & lstColumnUsed.List(lngSelIndex) & "," & lstColumnUsed.ItemData(lngSelIndex)
        ElseIf lngDo = lngSelIndex Then
            strCopy = strCopy & "|" & lstColumnUsed.List(lngTarIndex) & "," & lstColumnUsed.ItemData(lngTarIndex)
        Else
            strCopy = strCopy & "|" & lstColumnUsed.List(lngDo) & "," & lstColumnUsed.ItemData(lngDo)
        End If
    Next
    strCopy = Mid(strCopy, 2)
    Debug.Print strCopy
    
    lstColumnUsed.Clear
    arrData = Split(strCopy, "|")
    For lngDo = 0 To lngMax
        lstColumnUsed.AddItem Split(arrData(lngDo), ",")(0)
        lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = Val(Split(arrData(lngDo), ",")(1))
    Next
    lstColumnUsed.ListIndex = lngTarIndex
    Call SetMoveState
End Sub

Private Sub cmdOK_Click()
    Dim blnTrans As Boolean
    Dim intRow As Integer, intCount As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If Me.lstColumnUsed.ListCount = 0 Then
        MsgBox "请选择项目！", vbInformation, gstrSysName
        Exit Sub
    End If
    If Trim(txt模板名称.Text) = "" Then
        MsgBox "请录入模板名称！", vbInformation, gstrSysName
        txt模板名称.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(txt模板名称.Text, vbFromUnicode)) > 50 Then
        MsgBox "模板名称超长，最多允许25个汉字或50个字符！", vbInformation, gstrSysName
        txt模板名称.SetFocus
        Exit Sub
    End If
    
    '如果是新增则检查是否存在该模板
    If mint护理等级 = 9 Then
        gstrSQL = " Select 1 From 护理项目模板 Where 护理等级=[1] And 科室ID=[2] And Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.cbo适用护理等级.ItemData(Me.cbo适用护理等级.ListIndex), Me.cbo科室.ItemData(Me.cbo科室.ListIndex))
        If rsTemp.RecordCount <> 0 Then
            If MsgBox("已存在该护理项目模板,点“是”则更新！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    '准备保存
    intCount = Me.lstColumnUsed.ListCount
    gcnOracle.BeginTrans
    blnTrans = True
    
    Call zlDatabase.ExecuteProcedure("zl_护理项目模板_Delete(" & Me.cbo科室.ItemData(Me.cbo科室.ListIndex) & "," & Me.cbo适用护理等级.ItemData(Me.cbo适用护理等级.ListIndex) & ")", "删除当前模板")
    For intRow = 1 To intCount
        Debug.Print "zl_护理项目模板_Insert(" & cbo科室.ItemData(Me.cbo科室.ListIndex) & ",'" & Me.txt模板名称.Text & "'," & Me.cbo适用护理等级.ItemData(Me.cbo适用护理等级.ListIndex) & "," & Me.lstColumnUsed.ItemData(intRow - 1) & "," & intRow & ")"
        Call zlDatabase.ExecuteProcedure("zl_护理项目模板_Insert(" & cbo科室.ItemData(Me.cbo科室.ListIndex) & ",'" & Me.txt模板名称.Text & "'," & Me.cbo适用护理等级.ItemData(Me.cbo适用护理等级.ListIndex) & "," & Me.lstColumnUsed.ItemData(intRow - 1) & "," & intRow & ")", "产生模板数据")
    Next
    
    gcnOracle.CommitTrans
    blnTrans = False
    mblnSel = Me.cbo适用护理等级.ItemData(Me.cbo适用护理等级.ListIndex) & "|" & Me.cbo科室.ItemData(Me.cbo科室.ListIndex)
    
    mblnEdit = False
    Unload Me
    Exit Sub
errHand:
    If blnTrans Then
        gcnOracle.RollbackTrans
        blnTrans = False
    End If
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        mblnEdit = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '装入缺省数据
    mblnStart = False
    
    With Me.cbo适用护理等级
        .Clear
        .AddItem "特级护理录入模板"
        .ItemData(.NewIndex) = 0
        .AddItem "一级护理录入模板"
        .ItemData(.NewIndex) = 1
        .AddItem "二级护理录入模板"
        .ItemData(.NewIndex) = 2
        .AddItem "三级护理录入模板"
        .ItemData(.NewIndex) = 3
        .AddItem "不限/批量录入模板"
        .ItemData(.NewIndex) = -1
        .ListIndex = 0
    End With
    
    '提取临床科室(考虑到可能护士属于病区,需查找出对应的科室,都允许该护士处理;再考虑到临床科室的人直接来设置模板,两种模式都支持)
    If InStr(1, mstrPrivs, "编辑其它科室模板") <> 0 Then
        gstrSQL = " Select B.ID,B.名称 " & _
                  " From 部门性质说明 A,部门表 B" & _
                  " Where A.工作性质='临床' And A.服务对象 IN (2,3) And A.部门ID=B.ID" & _
                  " Order by B.编码"
    Else
        gstrSQL = " Select B.ID,B.编码,B.名称 " & _
                  " From 部门性质说明 A,部门表 B,部门人员 C" & _
                  " Where A.工作性质='临床' And A.服务对象 IN (2,3) And A.部门ID=B.ID" & _
                  " And B.ID=C.部门ID And C.人员ID=[1]" & _
                  " UNION " & _
                  " Select B.ID,B.编码,B.名称 " & _
                  " From 部门性质说明 A,部门表 B,病区科室对应 C" & _
                  " Where A.工作性质='临床' And A.服务对象 IN (2,3) And A.部门ID=B.ID And B.ID=C.科室ID And C.病区ID=[2]"
        gstrSQL = " Select Distinct ID,编码,名称 From (" & gstrSQL & ") Order by 编码"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId, glngDeptId)
    With rsTemp
        Me.cbo科室.Clear
        Do While Not .EOF
            Me.cbo科室.AddItem !名称
            Me.cbo科室.ItemData(Me.cbo科室.NewIndex) = !ID
            If !ID = mlng科室ID Then Me.cbo科室.ListIndex = .AbsolutePosition - 1
            .MoveNext
        Loop
        If .RecordCount = 0 Then
            MsgBox "你不属于任何一个临床科室！", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
        
        If Me.cbo科室.ListIndex = -1 Then Me.cbo科室.ListIndex = 0
    End With
    mblnStart = True
    
    Me.txt模板名称.Text = mstr模板名称
    If mint护理等级 = 9 Then
        Me.cbo适用护理等级.Enabled = True
    Else
        If mint护理等级 = -1 Then
            Me.cbo适用护理等级.ListIndex = 4
        Else
            Me.cbo适用护理等级.ListIndex = mint护理等级
        End If
        Me.cbo适用护理等级.Enabled = False
        Me.cbo科室.Enabled = False
    End If
    Call cbo适用护理等级_Click
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnEdit Then
        If MsgBox("还未保存数据，是否退出？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
End Sub

Private Sub lstColumnItems_DblClick()
    If lstColumnItems.ListCount = 0 Then Exit Sub
    Call cmdColumn_Click(0)
End Sub

Private Sub lstColumnUsed_Click()
    Call SetMoveState
End Sub

Private Sub lstColumnUsed_DblClick()
    If lstColumnUsed.ListCount = 0 Then Exit Sub
    Call cmdColumn_Click(1)
End Sub

Private Sub txt模板名称_GotFocus()
    Call zlControl.TxtSelAll(txt模板名称)
End Sub

Private Sub SetMoveState()
    cmdMove(0).Enabled = False
    cmdMove(1).Enabled = False
    
    If lstColumnUsed.ListIndex < 0 Then Exit Sub
    If lstColumnUsed.SelCount < 0 Then Exit Sub
    cmdMove(0).Enabled = (lstColumnUsed.ListIndex > 0)
    cmdMove(1).Enabled = (lstColumnUsed.ListIndex < lstColumnUsed.ListCount - 1)
End Sub
