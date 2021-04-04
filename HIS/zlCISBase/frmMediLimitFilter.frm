VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMediLimitFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "过滤条件"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "frmMediLimitFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   6630
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "药品分类(&0)"
      TabPicture(0)   =   "frmMediLimitFilter.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tvw分类"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "药品剂型(&1)"
      TabPicture(1)   =   "frmMediLimitFilter.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Chk剂型"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Lvw剂型"
      Tab(1).ControlCount=   2
      Begin VB.CheckBox Chk剂型 
         Appearance      =   0  'Flat
         Caption         =   "全选"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -74880
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   675
      End
      Begin MSComctlLib.TreeView tvw分类 
         Height          =   6255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   11033
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   476
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin MSComctlLib.ListView Lvw剂型 
         Height          =   5925
         Left            =   -74880
         TabIndex        =   5
         Top             =   720
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   10451
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5160
      TabIndex        =   1
      Top             =   7200
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   0
      Top             =   7200
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLimitFilter.frx":0342
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMediLimitFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mfrmMain As Form
Dim mstr剂型 As String
Dim mstr分类 As String
Dim mstr分类ID As String
Dim mlng库房ID As Long
Dim mstr类别 As String
Dim mblnSelect As Boolean

Private Sub Get药品剂型(ByVal lng库房ID As Long)
    Dim blnEXIST As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim bln中药库房 As Boolean
    
    '提取该库房现有剂型，供用户选择
    On Error GoTo errHandle
    bln中药库房 = False
    gstrSql = "Select 1 From 部门性质说明 " & _
             " Where 工作性质 Like '中药%' And 部门ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption & "[检查部门性质]", lng库房ID)

    If Not rsTemp.EOF Then bln中药库房 = True
    
    gstrSql = "Select Distinct J.编码,J.名称 " & _
             " From 诊疗执行科室 A,药品特性 B,药品剂型 J " & _
             " Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称" & _
             " And A.执行科室ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption & "[提取该库房现在剂型]", lng库房ID)
    Lvw剂型.ListItems.Clear
    
    With rsTemp
        Do While Not .EOF
            If blnEXIST = False Then
                blnEXIST = (!名称 = "方剂")
            End If
            Lvw剂型.ListItems.Add , "K" & !编码, !名称, , 1
            .MoveNext
        Loop
        If bln中药库房 And blnEXIST = False Then
            Lvw剂型.ListItems.Add , "KK1", "方剂", , 1
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Chk剂型_Click()
    If Chk剂型.Value = 2 Then Exit Sub
    Call SetSelect(Lvw剂型, Chk剂型.Value)
End Sub


Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim intItem As Integer, intItems As Integer
    Dim blnAllSelect As Boolean
    
    '取得剂型（不选剂型则不提取药品，需要手工录入）
    mstr剂型 = ""
    
    If Chk剂型.Value = 1 Then
        mstr剂型 = ""
    Else
        intItems = Me.Lvw剂型.ListItems.count
        blnAllSelect = True
        For intItem = 1 To intItems
            If Lvw剂型.ListItems(intItem).Checked Then
                mstr剂型 = mstr剂型 & "," & Lvw剂型.ListItems(intItem).Text
            Else
                blnAllSelect = False
            End If
        Next
    
        If mstr剂型 <> "" Then mstr剂型 = Mid(mstr剂型, 2)
        If blnAllSelect = True Then mstr剂型 = ""
    End If

    '取得药品分类（不选分类表示所有分类）
    mstr分类ID = ""
    mstr分类 = ""
    For intItem = 1 To tvw分类.Nodes.count
        If tvw分类.Nodes(intItem).Key = "Root" And tvw分类.Nodes(intItem).Checked = True Then
            mstr分类 = "所有"
            mstr分类ID = ""
            Exit For
        ElseIf tvw分类.Nodes(intItem).Key <> "Root" And _
            tvw分类.Nodes(intItem).Key <> "_中成药" And _
            tvw分类.Nodes(intItem).Key <> "_中草药" And _
            tvw分类.Nodes(intItem).Key <> "_西成药" And _
            tvw分类.Nodes(intItem).Checked Then
            mstr分类 = mstr分类 & "," & tvw分类.Nodes(intItem).Text
            mstr分类ID = mstr分类ID & "," & Mid(tvw分类.Nodes(intItem).Key, 2)
        End If
    Next

    If mstr分类ID <> "" Then mstr分类ID = Mid(mstr分类ID, 2)
     
    If mstr分类 <> "" Then mstr分类 = Mid(mstr分类, 2)
    
    mblnSelect = True
    
    Unload Me
End Sub

Private Sub Form_Load()
    Get药品分类 mlng库房ID
    Get药品剂型 mlng库房ID
End Sub
Private Sub tvw分类_NodeCheck(ByVal node As MSComctlLib.node)
    CheckNode node, node.Checked
    SetParentNode node, node.Checked
End Sub

Private Sub SetParentNode(ByVal node As MSComctlLib.node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not node.Parent Is Nothing Then
        If blnCheck = True Then
            '看是否他的兄弟接点是否也全是TRUE，如是，则置其父节点也为TRUE，否则，不管
            intIdx = node.FirstSibling.Index
            Do While intIdx <> node.LastSibling.Index
                If tvw分类.Nodes(intIdx).Checked = False Then
                    node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = tvw分类.Nodes(intIdx).Next.Index
            Loop
            If intIdx = node.LastSibling.Index Then
                If tvw分类.Nodes(intIdx).Checked = True Then
                    node.Parent.Checked = True
                End If
            End If
        Else
            node.Parent.Checked = False
        End If
        
        Set node = node.Parent
        If Not node Is Nothing Then
            SetParentNode node, blnCheck
        End If
    End If
End Sub

Private Function CheckNode(ByVal node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If node.Children > 0 Then
        Set node = node.Child
        Do While Not node Is Nothing
            node.Checked = blnCheck
            If node.Children > 0 Then
                CheckNode node, blnCheck
            End If
            Set node = node.Next
        Loop
    Else
        node.Checked = blnCheck
    End If
End Function

Public Function GetCondition(FrmMain As Form, ByVal lng库房ID As Long, ByVal str类别 As String, ByRef str分类 As String, ByRef str分类ID As String, ByRef str剂型 As String) As Boolean
    mstr剂型 = ""
    mstr分类 = ""
    mstr分类ID = ""
    mblnSelect = False
    
    mstr类别 = str类别
    mlng库房ID = lng库房ID
    
    Set mfrmMain = FrmMain
    Me.Show vbModal, FrmMain
    GetCondition = mblnSelect
    
    str剂型 = mstr剂型
    str分类 = mstr分类
    str分类ID = mstr分类ID
End Function

Private Sub Get药品分类(ByVal lng库房ID As Long)
    Dim rsData As ADODB.Recordset
    Dim strCon As String
    Dim objNode As node
    Dim str部门性质 As String
    Dim int类型 As Integer
    
'    str部门性质 = Get部门性质(mlng库房ID)
'
'    If InStr(1, str部门性质, "西药") > 0 And InStr(1, str部门性质, "成药") > 0 And InStr(1, str部门性质, "中药") > 0 Then
'        strCon = "类型 In (1, 2, 3)"
'    ElseIf InStr(1, str部门性质, "西药") > 0 And InStr(1, str部门性质, "成药") > 0 Then
'        strCon = "类型 In (1, 2)"
'    ElseIf InStr(1, str部门性质, "成药") > 0 And InStr(1, str部门性质, "中药") > 0 Then
'        strCon = "类型 In (2, 3)"
'    ElseIf InStr(1, str部门性质, "西药") > 0 And InStr(1, str部门性质, "中药") > 0 Then
'        strCon = "类型 In (1, 3)"
'    ElseIf InStr(1, str部门性质, "西药") > 0 Then
'        strCon = "类型 =1 "
'    ElseIf InStr(1, str部门性质, "成药") > 0 Then
'        strCon = "类型 =2 "
'    ElseIf InStr(1, str部门性质, "中药") > 0 Then
'        strCon = "类型 =3 "
'    ElseIf InStr(1, str部门性质, "制剂室") > 0 Then
'        strCon = "类型 In (1, 2, 3) "
'    End If
    
    If mstr类别 = "5" Then
        int类型 = 1
    ElseIf mstr类别 = "6" Then
        int类型 = 2
    ElseIf mstr类别 = "7" Then
        int类型 = 3
    End If
    
    On Error GoTo errHandle
    gstrSql = "Select Level as 层,ID,上级ID,名称,DECODE(类型,1,'西成药',2,'中成药','中草药') As 材质 " & _
        " From 诊疗分类目录 " & _
        " Where 类型=[1] " & _
        " Start With 上级id Is Null " & _
        " Connect By Prior ID = 上级id"
    Set rsData = zldatabase.OpenSQLRecord(gstrSql, "取药品分类", int类型)
    
    tvw分类.Nodes.Clear
    Set objNode = tvw分类.Nodes.Add(, , "Root", "所有分类", 1)
'    If InStr(1, str部门性质, "西药") > 0 Or InStr(1, str部门性质, "制剂室") > 0 Then Set objNode = tvw分类.Nodes.Add("Root", 4, "_西成药", "西成药", 1)
'    If InStr(1, str部门性质, "成药") > 0 Or InStr(1, str部门性质, "制剂室") > 0 Then Set objNode = tvw分类.Nodes.Add("Root", 4, "_中成药", "中成药", 1)
'    If InStr(1, str部门性质, "中药") > 0 Or InStr(1, str部门性质, "制剂室") > 0 Then Set objNode = tvw分类.Nodes.Add("Root", 4, "_中草药", "中草药", 1)
    
    If int类型 = 1 Then Set objNode = tvw分类.Nodes.Add("Root", 4, "_西成药", "西成药", 1)
    If int类型 = 2 Then Set objNode = tvw分类.Nodes.Add("Root", 4, "_中成药", "中成药", 1)
    If int类型 = 3 Then Set objNode = tvw分类.Nodes.Add("Root", 4, "_中草药", "中草药", 1)
    
    Do While Not rsData.EOF
        If rsData!层 = 1 Then
            Set objNode = tvw分类.Nodes.Add("_" & rsData!材质, 4, "_" & rsData!ID, rsData!名称, 1)
        Else
            Set objNode = tvw分类.Nodes.Add("_" & rsData!上级ID, 4, "_" & rsData!ID, rsData!名称, 1)
        End If
        rsData.MoveNext
    Loop

    tvw分类.Nodes("Root").Selected = True
    tvw分类.Nodes("Root").Expanded = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get部门性质(ByVal lng部门ID As Long) As String
    Dim rsData As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errHandle
    gstrSql = "Select 工作性质 From 部门性质说明 Where 部门id = [1]"
    Set rsData = zldatabase.OpenSQLRecord(gstrSql, "取部门工作性质", lng部门ID)
    
    With rsData
        Do While Not .EOF
            strTmp = IIf(strTmp = "", "", strTmp & ";") & !工作性质
            .MoveNext
        Loop
    End With
    
    Get部门性质 = strTmp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


