VERSION 5.00
Begin VB.Form frmNoticeBoardSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "公告栏设置"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   Icon            =   "frmNoticeBoardSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picUnit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   75
      ScaleHeight     =   1770
      ScaleWidth      =   2115
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3630
      Visible         =   0   'False
      Width           =   2145
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   990
         Picture         =   "frmNoticeBoardSet.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "确认"
         Top             =   1425
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterCancel 
         Height          =   315
         Left            =   1530
         Picture         =   "frmNoticeBoardSet.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "取消"
         Top             =   1425
         Width           =   450
      End
      Begin VB.ListBox lstUnit 
         Appearance      =   0  'Flat
         Height          =   1290
         Left            =   -15
         Style           =   1  'Checkbox
         TabIndex        =   23
         Top             =   15
         Width           =   2145
      End
   End
   Begin VB.CommandButton cmdSynchro 
      Caption         =   "同步(&C)"
      Height          =   350
      Left            =   75
      TabIndex        =   21
      Top             =   5475
      Width           =   1100
   End
   Begin VB.Frame fraSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   30
      TabIndex        =   20
      Top             =   5340
      Width           =   6645
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   5445
      TabIndex        =   10
      Top             =   5475
      Width           =   1100
   End
   Begin VB.Frame fraShape 
      Caption         =   "要素设置"
      Height          =   4800
      Left            =   4245
      TabIndex        =   15
      Top             =   420
      Width           =   2325
      Begin VB.ComboBox cboName 
         Height          =   300
         Left            =   600
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   285
         Width           =   1575
      End
      Begin VB.TextBox txtRow 
         Height          =   300
         Left            =   600
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1005
         Width           =   1575
      End
      Begin VB.ComboBox cboPosition 
         Height          =   300
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1365
         Width           =   1575
      End
      Begin VB.CheckBox chkHide 
         Caption         =   "无数据时隐藏该项"
         Height          =   225
         Left            =   180
         TabIndex        =   5
         Top             =   1755
         Width           =   1905
      End
      Begin VB.CommandButton cmdboundItem 
         Caption         =   "绑定诊疗项目"
         Height          =   345
         Left            =   180
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2025
      End
      Begin VB.TextBox txtItem 
         Appearance      =   0  'Flat
         Height          =   1920
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2415
         Width           =   2025
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除"
         Height          =   350
         Left            =   1275
         TabIndex        =   9
         Top             =   4380
         Width           =   945
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新增"
         Height          =   350
         Left            =   165
         TabIndex        =   8
         Top             =   4380
         Width           =   945
      End
      Begin VB.TextBox txtCName 
         Height          =   300
         Left            =   600
         MaxLength       =   20
         TabIndex        =   2
         Top             =   645
         Width           =   1575
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   195
         TabIndex        =   19
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblCName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "别名"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   195
         TabIndex        =   18
         Top             =   705
         Width           =   360
      End
      Begin VB.Label lblRow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "行号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   195
         TabIndex        =   17
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label lblPosition 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "位置"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   195
         TabIndex        =   16
         Top             =   1425
         Width           =   360
      End
   End
   Begin VB.PictureBox picBak 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   60
      Picture         =   "frmNoticeBoardSet.frx":0B20
      ScaleHeight     =   4665
      ScaleWidth      =   4065
      TabIndex        =   12
      Top             =   510
      Width           =   4095
      Begin VB.Timer tmrFresh 
         Enabled         =   0   'False
         Interval        =   60
         Left            =   3495
         Top             =   225
      End
      Begin VB.Label lblElementName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "要素名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblElementCT 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "要素内容"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   0
         Left            =   840
         TabIndex        =   13
         Top             =   45
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   0
         Left            =   -30
         Top             =   -30
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   1
         Left            =   330
         Top             =   -30
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   2
         Left            =   720
         Top             =   -30
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   3
         Left            =   720
         Top             =   120
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   4
         Left            =   720
         Top             =   270
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   5
         Left            =   330
         Top             =   270
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   6
         Left            =   -30
         Top             =   270
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   7
         Left            =   -30
         Top             =   120
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.ComboBox cboUnit 
      Height          =   300
      Left            =   510
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   3645
   End
   Begin VB.Label lbl病区 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "病区"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   75
      TabIndex        =   11
      Top             =   135
      Width           =   360
   End
End
Attribute VB_Name = "frmNoticeBoardSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mblnClick As Boolean
Private mstrPrivs As String
Private mrsBoard As New ADODB.Recordset
Private mrsUnit As New ADODB.Recordset
Private mlngUnitID As Long

Public Function ShowMe(ByVal frmParent As Object, ByVal strPrivs As String, Optional ByVal lngUnitID As Long = 0) As Boolean
    mstrPrivs = strPrivs
    mlngUnitID = lngUnitID
    
    mblnOK = False
    If frmParent Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmParent
    End If
    ShowMe = mblnOK
End Function

Private Sub cboName_Change()
    If mblnClick Then Exit Sub
End Sub

Private Sub cboName_Click()
    cmdboundItem.Enabled = True
    If cboName.ListIndex > -1 And cboName.ListIndex < cboName.ListCount - 1 Then
        cmdboundItem.Enabled = False
    End If
End Sub

Private Sub cboName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cboName_Validate(False)
    ElseIf KeyAscii = vbKeySpace Then
        KeyAscii = 0
    End If
End Sub

Private Sub cboName_Validate(Cancel As Boolean)
    If Not Cbo.Locate(cboName, cboName.Text) And cboName.ListIndex <> -1 Then cboName.ListIndex = -1
    cmdboundItem.Enabled = True
    If cboName.ListIndex > -1 And cboName.ListIndex < cboName.ListCount - 1 Then
        cmdboundItem.Enabled = False
    End If
End Sub

Private Sub cboUnit_Click()
    picUnit.Visible = False
    If cboUnit.ListIndex = -1 Then Exit Sub
    tmrFresh.Enabled = True
End Sub

Private Sub cmdAdd_Click()
    Dim lngID As Long
    Dim intPos As Integer
    Dim intCount As Integer
    Dim blnTrans As Boolean
    Dim strIDs As String, strItems As String
    Dim strSQL As String
    
    If Trim(cboName.Text) = "" Then
        MsgBox "名称不能为空！", vbInformation, gstrSysName
        cboName.SetFocus
        Exit Sub
    End If
    If CheckLen(cboName, 20, "名称") = False Then Exit Sub
    
    If Trim(txtCName.Text) = "" Then
        MsgBox "别名不能为空！", vbInformation, gstrSysName
        txtCName.SetFocus
        Exit Sub
    End If
    If CheckLen(txtCName, 20, "别名") = False Then Exit Sub
    
    If Trim(txtRow.Text) = "" Then
        MsgBox "行号不能为空！", vbInformation, gstrSysName
        txtRow.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtRow.Text) Then
        MsgBox "行号中不能含有非法字符！", vbInformation, gstrSysName
        txtRow.SetFocus
        Exit Sub
    End If
    If Val(txtRow.Text) < 0 Or Val(txtRow.Text) > 13 Then
        MsgBox "行号不能小于零或大于13！", vbInformation, gstrSysName
        txtRow.SetFocus
        Exit Sub
    End If
    
    '项目名称不能重复
    If Val(fraShape.Tag) = 0 Then '新增
        mrsBoard.Filter = "名称='" & cboName.Text & "'"
    Else
        mrsBoard.Filter = "名称='" & cboName.Text & "' And id<>" & Val(fraShape.Tag)
    End If
    If mrsBoard.RecordCount > 0 Then
        MsgBox "要素名称[" & cboName.Text & "]已经存在，请检查！", vbInformation, gstrSysName
        cboName.SetFocus
        Exit Sub
    End If
    '位置不能重复
    mrsBoard.Filter = "行号=" & Val(txtRow.Text) & " And 位置=" & Me.cboPosition.ListIndex + 1 & _
        IIf(Val(fraShape.Tag) = 0, "", " And id<>" & Val(fraShape.Tag))
    
    If mrsBoard.RecordCount > 0 Then
        MsgBox "要素名称[" & cboName.Text & "]与[" & mrsBoard!名称 & "]的行号和位置重复，请检查！", vbInformation, gstrSysName
        txtRow.SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    blnTrans = True
    
    If txtItem.Tag <> "" Then
        strIDs = txtItem.Tag
    End If
    lngID = Val(fraShape.Tag)
    
    If lngID = 0 Then lngID = zlDatabase.GetNextId("病区公告栏样式")
    strSQL = "Zl_病区公告栏样式_Insert(" & lngID & "," & Me.cboUnit.ItemData(Me.cboUnit.ListIndex) & "," & _
        "'" & Me.cboName.Text & "','" & Me.txtCName.Text & "'," & Me.txtRow.Text & "," & Me.cboPosition.ListIndex + 1 & "," & _
        IIf(Me.cboName.ListIndex = -1, 0, 1) & "," & chkHide.Value & ")"
        
    Call zlDatabase.ExecuteProcedure(strSQL, "病区公告栏")
    '保存格式:<ITEMLIST><ITEM><XH/><MC/></ITEM></ITEMLIST>
    intCount = 0
    If strIDs <> "" Then
        Do While strIDs <> ""
            If Len(strIDs) > 3800 Then
                '向左搜寻逗号
                intPos = GetSplit(Mid(strIDs, 1, 3800))
                strItems = Mid(strIDs, 1, intPos)
                strIDs = Mid(strIDs, intPos + 1)
            Else
                strItems = strIDs
                strIDs = ""
            End If
            
            strSQL = "Zl_病区公告栏样式_Updateitem(" & lngID & ",'" & strItems & "'," & IIf(intCount = 0, "1", "0") & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "病区公告栏")
            intCount = intCount + 1
        Loop
    Else
        strItems = ""
        strSQL = "Zl_病区公告栏样式_Updateitem(" & lngID & ",'" & strItems & "'," & IIf(intCount = 0, "1", "0") & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "病区公告栏")
    End If
    gcnOracle.CommitTrans
    blnTrans = False
    
    mblnOK = True
    tmrFresh.Enabled = True
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdboundItem_Click()
    Dim strIDs As String, strNames As String
    strIDs = txtItem.Tag
    strNames = txtItem.Text
    If frmClinicSelect.ShowMe(Me, strIDs, strNames) = True Then
        txtItem.Tag = strIDs
        txtItem.Text = strNames
    End If
End Sub

Private Sub cmdDel_Click()
    On Error GoTo ErrHand
    
    If Val(fraShape.Tag) = 0 Then Exit Sub
    If MsgBox("你确定要删除要素：" & cboName.Text & "？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call zlDatabase.ExecuteProcedure("ZL_病区公告栏样式_DELETEITEM(" & Val(fraShape.Tag) & ")", "ZL_公告栏样式_DELETEITEM")
    
    tmrFresh.Enabled = True
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFilterCancel_Click()
    picUnit.Visible = False
End Sub

Private Sub cmdFilterOK_Click()
    Dim i As Integer, lngID As Long
    Dim blnTrans As Boolean
    Dim arrSQL, strSQL As String
    
    On Error GoTo ErrHand
    If cboUnit.ListIndex = -1 Then Exit Sub
    
    If lstUnit.SelCount = 0 Then
        MsgBox "请至少选择一个病区!", vbInformation, gstrSysName
        lstUnit.SetFocus
    End If
    
    arrSQL = Array()
    For i = 1 To lstUnit.ListCount - 1
        If lstUnit.Selected(i) = True Then
            lngID = Val(lstUnit.ItemData(i))
            strSQL = "Zl_病区公告栏样式_Build(" & lngID & "," & Val(cboUnit.ItemData(cboUnit.ListIndex)) & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        End If
    Next
    
    '更新数据
    gcnOracle.BeginTrans
    blnTrans = True
    For i = 0 To UBound(arrSQL)
        If CStr(arrSQL(i)) <> "" Then
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "病区公告栏样式同步")
        End If
    Next i
    gcnOracle.CommitTrans
    blnTrans = False
    picUnit.Visible = False
    MsgBox "已完成病区公告栏样式的同步!", vbInformation, gstrSysName
    Exit Sub
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSynchro_Click()
    If mrsUnit Is Nothing Then Exit Sub
    If mrsUnit.RecordCount < 1 Then Exit Sub
    If cboUnit.ListIndex = -1 Then Exit Sub
    
    With mrsUnit
        lstUnit.Clear
        lstUnit.AddItem "全部"
        .Filter = ""
        Do While Not .EOF
            If !ID <> cboUnit.ItemData(cboUnit.ListIndex) Then
                lstUnit.AddItem !编码 & "-" & !名称
                lstUnit.ItemData(lstUnit.NewIndex) = !ID
            End If
        .MoveNext
        Loop
    End With
    
    lstUnit.ListIndex = 0
    picUnit.Top = cmdSynchro.Top - picUnit.Height - 30
    picUnit.Left = cmdSynchro.Left
    picUnit.Visible = True
    picUnit.ZOrder
    lstUnit.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ZLCommFun.PressKey vbKeyTab
    ElseIf KeyCode = vbKeyEscape Then
        cmdAdd.Caption = "新增"
        Call SetShape
        fraShape.Tag = ""
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    tmrFresh.Enabled = False
    Call InitUnits
    With cboName
        .Clear
        .AddItem "新入院列表"
        .ItemData(.NewIndex) = 1
        .AddItem "新转入列表"
        .ItemData(.NewIndex) = 2
        .AddItem "一级护理列表"
        .ItemData(.NewIndex) = 3
        .AddItem "特级护理列表"
        .ItemData(.NewIndex) = 4
        .AddItem "病危列表"
        .ItemData(.NewIndex) = 5
        .AddItem "预出院列表"
        .ItemData(.NewIndex) = 6
        .AddItem "手术列表"
        .ItemData(.NewIndex) = 7
        .AddItem "发烧列表"
        .ItemData(.NewIndex) = 8
        .AddItem "过敏史列表"
        .ItemData(.NewIndex) = 9
        .AddItem "测血压列表"
        .ItemData(.NewIndex) = 10
    End With
    
    cboPosition.Clear
    cboPosition.AddItem "左"
    cboPosition.AddItem "右"
    cboPosition.ListIndex = 0
    
    Call SetEnable
    tmrFresh.Enabled = True
End Sub

Private Sub SetEnable()
    fraShape.Enabled = (cboUnit.ListCount > 0 And InStr(1, mstrPrivs, ";病区公告栏设置;") > 0)
    cmdAdd.Enabled = fraShape.Enabled
    cmdDel.Enabled = cmdAdd.Enabled
    cmdSynchro.Enabled = cmdAdd.Enabled
End Sub

Private Function InitUnits() As Boolean
'功能：初始化住院护理病区
    Dim strSQL As String, strUnits As String, i As Long

    On Error GoTo errH
    strUnits = GetUserUnits
    
    '包含门观察室
    If InStr(mstrPrivs, "全院病人") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by A.编码"
    Else
        '求有权病区：直接所在病区+所在科室所属病区
        strSQL = _
            " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.编码,C.名称,Nvl(B.缺省,0) as 缺省" & _
            " From 病区科室对应 A,部门人员 B,部门表 C" & _
            " Where A.病区ID=C.ID And B.部门ID=A.科室ID And B.人员ID=[1]" & _
            " And Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=A.科室ID)" & _
            " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=A.科室ID)" & _
            " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSQL & ") Group by ID,编码,名称 Order by 编码"
    End If

    cboUnit.Clear
    Set mrsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    If Not mrsUnit.EOF Then
        For i = 1 To mrsUnit.RecordCount
            cboUnit.AddItem mrsUnit!编码 & "-" & mrsUnit!名称
            cboUnit.ItemData(cboUnit.NewIndex) = mrsUnit!ID
            If InStr(mstrPrivs, "全院病人") > 0 Then
                If mlngUnitID > 0 And mlngUnitID = mrsUnit!ID Then
                    Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
                If mrsUnit!ID = UserInfo.部门ID And cboUnit.ListIndex = -1 Then '直接所属优先
                    Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
                If InStr("," & strUnits & ",", "," & mrsUnit!ID & ",") > 0 And cboUnit.ListIndex = -1 Then
                    Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
            Else '所属缺省病区包含的可能有多个
                If mrsUnit!缺省 = 1 And cboUnit.ListIndex = -1 Then
                    Call Cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
            End If
            mrsUnit.MoveNext
        Next
    End If
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then
        Call Cbo.SetIndex(cboUnit.hwnd, 0)
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetUserUnits() As String
'功能：获取操作员所属的病区(直接属于病区或所在科室所属的病区),可能有多个
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, blnNew As Boolean
        
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    If blnNew Then
        strSQL = _
            "Select Distinct 病区ID From (" & _
            " Select A.部门ID as 病区ID" & _
            " From 部门性质说明 A,部门人员 B" & _
            " Where A.部门ID=B.部门ID And B.人员ID=[1]" & _
            " And A.服务对象 in(1,2,3) And A.工作性质='护理'" & _
            " Union" & _
            " Select A.病区ID From 病区科室对应 A,部门人员 B" & _
            " Where A.科室ID=B.部门ID And B.人员ID=[1])"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
    ElseIf rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
    End If
    For i = 1 To rsTmp.RecordCount
        GetUserUnits = GetUserUnits & "," & rsTmp!病区ID
        rsTmp.MoveNext
    Next
    
    GetUserUnits = Mid(GetUserUnits, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetItemNameOrID(ByVal lngID As Long, Optional bytType As Byte = 1) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHand:
    
    If bytType = 0 Then
        strSQL = "" & _
        " SELECT a.XH 项目" & _
        " FROM 病区公告栏样式 p," & _
        " XMLTable('/ITEMLIST/ITEM/XH' PASSING p.诊疗项目" & _
        " COLUMNS XH VARCHAR2(256) PATH '/XH') a" & _
        " Where p.id = [1]"
    Else
        strSQL = "" & _
            " SELECT a.MC 项目" & _
            " FROM 病区公告栏样式 p," & _
            " XMLTable('/ITEMLIST/ITEM/MC' PASSING p.诊疗项目" & _
            " COLUMNS MC VARCHAR2(256) PATH '/MC') a" & _
            " Where p.id = [1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取项目内容", lngID)
    With rsTemp
        Do While Not rsTemp.EOF
            GetItemNameOrID = GetItemNameOrID & "," & rsTemp!项目
            rsTemp.MoveNext
        Loop
    End With
    
    If GetItemNameOrID <> "" Then GetItemNameOrID = Mid(GetItemNameOrID, 2)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mrsUnit = Nothing
    Set mrsBoard = Nothing
End Sub

Private Sub lblElementName_Click(Index As Integer)
    Dim intDo As Integer, intCount As Integer
    
    mblnClick = True
    intCount = lblElementName.Count - 1
    For intDo = 1 To intCount
        lblElementName(intDo).BackStyle = 0
    Next
    fraShape.Tag = lblElementName(Index).Tag
    cmdAdd.Caption = "修改"
    lblElementName(Index).BackStyle = 1
    Call SetShape(Index)
    
    '定位该要素，显示相应的属性
    mrsBoard.Filter = "ID=" & Val(lblElementName(Index).Tag)
    If mrsBoard.RecordCount = 0 Then Exit Sub
    
    If Not Cbo.Locate(cboName, mrsBoard!名称) Then cboName.Text = mrsBoard!名称
    txtCName.Text = IIf(IsNull(mrsBoard!别名), "", mrsBoard!别名)
    txtRow.Text = mrsBoard!行号
    cboPosition.ListIndex = mrsBoard!位置 - 1
    chkHide.Value = mrsBoard!是否隐藏
    
    txtItem.Text = GetItemNameOrID(Val(lblElementName(Index).Tag))
    txtItem.Tag = GetItemNameOrID(Val(lblElementName(Index).Tag), 0)
    If cboName.Enabled And cboName.Visible Then cboName.SetFocus
    Call cboName_Validate(False)
    mblnClick = False
End Sub

Private Sub lstUnit_ItemCheck(Item As Integer)
    Dim i As Integer
    
    If Item = 0 Then
        For i = 1 To lstUnit.ListCount - 1
            lstUnit.Selected(i) = lstUnit.Selected(0)
        Next
    ElseIf Not lstUnit.Selected(Item) Then
        lstUnit.Selected(0) = False
    ElseIf lstUnit.SelCount = lstUnit.ListCount - 1 Then
        lstUnit.Selected(0) = True
    End If
End Sub

Private Sub lstUnit_LostFocus()
    If Not Me.ActiveControl Is cmdFilterOK _
        And Not Me.ActiveControl Is cmdFilterCancel _
        And Not Me.ActiveControl Is lstUnit _
        And Not Me.ActiveControl Is picUnit Then picUnit.Visible = False
End Sub


Private Sub picUnit_GotFocus()
    If lstUnit.Visible And lstUnit.Enabled Then lstUnit.SetFocus
End Sub

Private Sub picUnit_Resize()
    On Error Resume Next
    
    lstUnit.Left = -15
    lstUnit.Top = -15
    lstUnit.Width = picUnit.Width
    
    cmdFilterCancel.Left = picUnit.ScaleWidth - cmdFilterCancel.Width - 100
    cmdFilterOK.Left = cmdFilterCancel.Left - cmdFilterOK.Width - 60
    
    cmdFilterOK.Top = lstUnit.Height + (picUnit.ScaleHeight - lstUnit.Height - cmdFilterOK.Height) / 2
    cmdFilterCancel.Top = cmdFilterOK.Top
End Sub

Private Sub tmrFresh_Timer()
    tmrFresh.Enabled = False
    Call SetShape
    Call RefreshBoard
End Sub

Private Sub SetShape(Optional ByVal intIndex As Integer = 0)
    Dim blnShow As Boolean
    blnShow = (intIndex > 0)
    
    If blnShow Then
        shpCircle(0).Left = lblElementName(intIndex).Left - shpCircle(0).Width
        shpCircle(0).Top = lblElementName(intIndex).Top - shpCircle(0).Height
        shpCircle(1).Left = lblElementName(intIndex).Left + (lblElementName(intIndex).Width - shpCircle(0).Width) / 2
        shpCircle(1).Top = shpCircle(0).Top
        shpCircle(2).Left = lblElementName(intIndex).Left + lblElementName(intIndex).Width
        shpCircle(2).Top = shpCircle(0).Top
        shpCircle(3).Left = shpCircle(2).Left
        shpCircle(3).Top = lblElementName(intIndex).Top + (lblElementName(intIndex).Height - shpCircle(3).Height) / 2
        shpCircle(4).Left = shpCircle(2).Left
        shpCircle(4).Top = lblElementName(intIndex).Top + lblElementName(intIndex).Height
        shpCircle(5).Left = shpCircle(1).Left
        shpCircle(5).Top = shpCircle(4).Top
        shpCircle(6).Left = shpCircle(0).Left
        shpCircle(6).Top = shpCircle(4).Top
        shpCircle(7).Left = shpCircle(0).Left
        shpCircle(7).Top = shpCircle(3).Top
    End If
    
    shpCircle(0).Visible = blnShow
    shpCircle(1).Visible = blnShow
    shpCircle(2).Visible = blnShow
    shpCircle(3).Visible = blnShow
    shpCircle(4).Visible = blnShow
    shpCircle(5).Visible = blnShow
    shpCircle(6).Visible = blnShow
    shpCircle(7).Visible = blnShow
End Sub


Private Sub RefreshBoard()
    Dim lng病区ID As Long
    Dim intDel As Integer, intCount As Integer
    Dim strSQL As String
    Dim arrLeft, arrRight
    Dim i As Integer, j As Integer
    
    On Error GoTo ErrHand
    '刷新公告栏
    
    '先删除所有控件
    intCount = lblElementName.Count - 1
    For intDel = 1 To intCount
        Unload lblElementName(intDel)
        Unload lblElementCT(intDel)
    Next
    '赋初值
    cboName.Text = ""
    txtCName.Text = ""
    txtRow.Text = ""
    cboPosition.ListIndex = 0
    chkHide.Value = 0
    txtItem.Text = ""
    txtItem.Tag = ""
    cmdboundItem.Enabled = True
    '提取数据
    fraShape.Tag = "": cmdAdd.Caption = "新增"
    lng病区ID = Me.cboUnit.ItemData(Me.cboUnit.ListIndex)
    strSQL = " Select ID,名称,别名,行号,位置,是否固定,是否隐藏,内容" & _
              " From 病区公告栏样式 " & _
              " Where 病区ID=[1] " & _
              " Order by 行号,位置"
    Set mrsBoard = zlDatabase.OpenSQLRecord(strSQL, "提取病区公告", lng病区ID)
    
    '依次加载控件
    arrLeft = Array()
    arrRight = Array()
    With mrsBoard
        Do While Not .EOF
            Load lblElementName(.AbsolutePosition)
            lblElementName(.AbsolutePosition).Tag = !ID
            lblElementName(.AbsolutePosition).Caption = !别名
            lblElementName(.AbsolutePosition).Top = lblElementName(0).Top + (!行号 - 1) * 360
            lblElementName(.AbsolutePosition).Left = IIf(!位置 = 1, 60, picBak.Width - lblElementName(.AbsolutePosition).Width - 1000)
            lblElementName(.AbsolutePosition).Visible = True
            
            Load lblElementCT(.AbsolutePosition)
            lblElementCT(.AbsolutePosition).Caption = IIf(IsNull(!内容), "", !内容)
            lblElementCT(.AbsolutePosition).Top = lblElementCT(0).Top + (!行号 - 1) * 360
            lblElementCT(.AbsolutePosition).Left = lblElementName(.AbsolutePosition).Left + lblElementName(.AbsolutePosition).Width + 60
            lblElementCT(.AbsolutePosition).AutoSize = False
            lblElementCT(.AbsolutePosition).WordWrap = False
            lblElementCT(.AbsolutePosition).Height = 240
            lblElementCT(.AbsolutePosition).Visible = False
            If !位置 = 1 Then
                ReDim Preserve arrLeft(UBound(arrLeft) + 1)
                arrLeft(UBound(arrLeft)) = .AbsolutePosition & "," & Val(!行号)
            Else
                ReDim Preserve arrRight(UBound(arrRight) + 1)
                arrRight(UBound(arrRight)) = .AbsolutePosition & "," & Val(!行号)
            End If
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    '重新整理要素位置
    For i = 0 To UBound(arrLeft)
        For j = 0 To UBound(arrRight)
            If Split(arrLeft(i), ",")(1) = Split(arrRight(j), ",")(1) Then
                lblElementCT(Val(Split(arrLeft(i), ",")(0))).Width = lblElementName(Val(Split(arrRight(j), ",")(0))).Left - lblElementCT(Val(Split(arrLeft(i), ",")(0))).Left - 60
                Exit For
            End If
        Next j
    Next i
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetSplit(ByVal strInput As String) As Integer
    Dim intPos As Integer
    '向左搜寻逗号,返回逗号的位置
    intPos = 3800
    Do While True
        If Mid(strInput, intPos, 1) = "," Then
            intPos = intPos - 1
            GetSplit = intPos
            Exit Function
        End If
        intPos = intPos - 1
    Loop
End Function

Private Sub txtCName_GotFocus()
    Call zlControl.TxtSelAll(txtCName)
End Sub

Private Sub txtRow_GotFocus()
    Call zlControl.TxtSelAll(txtRow)
End Sub

Private Sub txtRow_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        If InStr(1, ",0,1,2,3,4,5,6,7,8,9," & Chr(8) & ",", "," & Chr(KeyAscii) & ",") = 0 Then KeyAscii = 0
    End If
End Sub

Public Function CheckLen(txt As Object, intLen As Integer, strName As String) As Boolean
'功能：检查工本框的真实长度是否在指定限制长度内
    If LenB(StrConv(txt.Text, vbFromUnicode)) > intLen Then
        MsgBox strName & "只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！", vbInformation, gstrSysName
        txt.SetFocus: Exit Function
    End If
    CheckLen = True
End Function
