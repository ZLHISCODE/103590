VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCheckCondition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "盘点条件设置"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab sst 
      Height          =   5895
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本(&1)"
      TabPicture(0)   =   "frmCheckCondition.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lbl分类"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDate"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lbl盘点方式"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl库房"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tvw分类"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dtpDate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkNoNum"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkNum"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Cbo盘点方式"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cbo库房"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "库房货位(&2)"
      TabPicture(1)   =   "frmCheckCondition.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "tvw货位"
      Tab(1).Control(2)=   "chk货位"
      Tab(1).ControlCount=   3
      Begin VB.CheckBox chk货位 
         Caption         =   "仅显示当前库房已分配的货位"
         Height          =   255
         Left            =   -73080
         TabIndex        =   14
         Top             =   480
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.ComboBox cbo库房 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   3045
      End
      Begin VB.ComboBox Cbo盘点方式 
         Height          =   300
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   4185
         Width           =   3045
      End
      Begin VB.CheckBox chkNum 
         Caption         =   "盘无库存材料"
         Height          =   255
         Left            =   930
         TabIndex        =   5
         Top             =   4965
         Width           =   1935
      End
      Begin VB.CheckBox chkNoNum 
         Caption         =   "仅盘无数量，但有库存金额或差价的材料"
         Height          =   255
         Left            =   930
         TabIndex        =   4
         Top             =   5250
         Width           =   3585
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   930
         TabIndex        =   6
         Top             =   4575
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   121700355
         CurrentDate     =   36901
      End
      Begin MSComctlLib.TreeView tvw分类 
         Height          =   3000
         Left            =   240
         TabIndex        =   9
         Top             =   1125
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   5292
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvw货位 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   15
         Top             =   750
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   8493
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "材料库房货位(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74760
         TabIndex        =   16
         Top             =   510
         Width           =   1350
      End
      Begin VB.Label lbl库房 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库房(&S)"
         Height          =   180
         Left            =   285
         TabIndex        =   13
         Top             =   540
         Width           =   630
      End
      Begin VB.Label Lbl盘点方式 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "方式(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   270
         TabIndex        =   12
         Top             =   4245
         Width           =   630
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间(&T)"
         Height          =   180
         Left            =   270
         TabIndex        =   11
         Top             =   4635
         Width           =   630
      End
      Begin VB.Label Lbl分类 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "分类(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   285
         TabIndex        =   10
         Top             =   870
         Width           =   630
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   5055
      TabIndex        =   2
      Top             =   4980
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5040
      TabIndex        =   1
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5040
      TabIndex        =   0
      Top             =   525
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
            Picture         =   "frmCheckCondition.frx":0038
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
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
            Picture         =   "frmCheckCondition.frx":0F12
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCheckCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean
Private mblnBootUp As Boolean
Private mstr分类ID  As String
Private mlng库房id As Long
Private mlng盘点方式 As Integer
Private mstr盘点时间 As String
Private mint盘无库存材料 As Integer
Private mbln盘点零数量且有金额 As Boolean
Private mfrmMain As Form
Private Const mlngModule = 1719
Private mstr库房货位  As String
Public Function GetCondition(frmMain As Form, ByRef str分类ID As String, ByRef lng库房ID As Long, _
        ByRef 盘点方式 As Integer, ByRef str盘点时间, ByRef int盘无库存材料 As Integer, _
        ByRef bln盘点零数量且有金额 As Boolean, ByRef str库房货位 As String) As Boolean
    
    mstr分类ID = ""
    mlng库房id = 0
    mlng盘点方式 = 0
    mstr盘点时间 = ""
    mint盘无库存材料 = 0
    mblnSelect = False
    mbln盘点零数量且有金额 = False
    mstr库房货位 = "所有"
    
    Set mfrmMain = frmMain
    Me.Show vbModal, frmMain
    GetCondition = mblnSelect
    
    str分类ID = mstr分类ID
    lng库房ID = mlng库房id
    盘点方式 = mlng盘点方式
    str盘点时间 = mstr盘点时间
    int盘无库存材料 = mint盘无库存材料
    bln盘点零数量且有金额 = mbln盘点零数量且有金额
    str库房货位 = mstr库房货位
End Function

Private Sub chkNoNum_Click()
    chkNum.Enabled = chkNoNum.Value <> 1
    If chkNum.Enabled = False Then
        chkNum.Value = False
    End If
End Sub
Private Sub chkNum_Click()
    chkNoNum.Enabled = chkNum.Value <> 1
    If chkNoNum.Enabled = False Then
        chkNoNum.Value = 0
    End If
End Sub

 
Private Sub chk货位_Click()
    Load库房货位
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim intItem As Integer, intItems As Integer
    Dim i As Long
    
    mstr分类ID = ""
    If tvw分类.Nodes("Root").Checked Then
        '所有卫材
        mstr分类ID = "所有卫生材料"
    Else
        For i = 1 To tvw分类.Nodes.Count
            If tvw分类.Nodes(i).Key <> "Root" And _
                tvw分类.Nodes(i).Checked Then
                mstr分类ID = mstr分类ID & "," & Mid(tvw分类.Nodes(i).Key, 2)
            End If
        Next
        If mstr分类ID <> "" Then
                mstr分类ID = Mid(mstr分类ID, 2)
        End If
    End If
    
    '取得库房货位（不选库房表示所有库房）
    mstr库房货位 = ""
    For intItem = 1 To tvw货位.Nodes.Count
        If tvw货位.Nodes(intItem).Key <> "Root" Then
            If tvw货位.Nodes(intItem).Checked Then
                mstr库房货位 = mstr库房货位 & "," & tvw货位.Nodes(intItem).Text
            End If
        End If
        
'        If tvw货位.Nodes(intItem).Key = "Root" And tvw货位.Nodes(intItem).Checked = True Then
'            mstr库房货位 = ""
'            Exit For
'        ElseIf tvw货位.Nodes(intItem).Checked Then
'            mstr库房货位 = mstr库房货位 & "," & tvw货位.Nodes(intItem).Text
'        End If
    Next
    
    If mstr库房货位 <> "" Then
        mstr库房货位 = Mid(mstr库房货位, 2)
    End If
    
    mlng库房id = cbo库房.ItemData(cbo库房.ListIndex)
    mlng盘点方式 = Cbo盘点方式.ItemData(Cbo盘点方式.ListIndex)
    mstr盘点时间 = Format(dtpDate.Value, "yyyy-MM-dd hh:mm:ss")
    mblnSelect = True
    mint盘无库存材料 = chkNum.Value
    mbln盘点零数量且有金额 = (chkNoNum.Value = 1)
    
    frmCheckCard.txtStock.Caption = cbo库房.Text
    frmCheckCard.txtStock.Tag = mlng库房id
    frmCheckCard.txtCheckDate = mstr盘点时间
    frmCheckCard.CmdSave.Enabled = False
    frmCheckCard.CmdCancel.Enabled = False
    
    Unload Me
End Sub

Private Sub Command1_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub Form_Activate()
    If mblnBootUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strSelectStock As String
    
    On Error GoTo ErrHandle
    strSelectStock = IIf(Val(zlDatabase.GetPara("是否选择库房", glngSys, mlngModule, "0")) = 1, 1, 0)
    '卫材材质权限控制
    
    dtpDate.Value = Format(sys.Currentdate, dtpDate.CustomFormat)
    dtpDate.MaxDate = dtpDate.Value
    
    mblnBootUp = False

    With Cbo盘点方式
        .Clear
        .AddItem "每日"
        .ItemData(.NewIndex) = 1
        .AddItem "每周"
        .ItemData(.NewIndex) = 2
        .AddItem "每月"
        .ItemData(.NewIndex) = 3
        .AddItem "每季度"
        .ItemData(.NewIndex) = 4
        .AddItem "忽略盘点方式"
        .ItemData(.NewIndex) = 5
        .ListIndex = 0
    End With
    
    With mfrmMain.cboStock
        cbo库房.Clear
        For i = 0 To .ListCount - 1
            cbo库房.AddItem .List(i)
            cbo库房.ItemData(cbo库房.NewIndex) = .ItemData(i)
        Next
        cbo库房.ListIndex = .ListIndex
    End With
        
    If InStr(1, gstrPrivs, "所有库房") <> 0 Then
        If strSelectStock = "0" Then
            cbo库房.Enabled = False
        Else
            cbo库房.Enabled = True
        End If
    Else
        cbo库房.Enabled = False
    End If
    
    With rsTemp
        gstrSQL = "Select 编码,名称 From 诊疗分类目录 where 类型=7 order by 编码 "
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "卫材分类")
        
        If .EOF Then
            MsgBox "卫材分类不完整！", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
'用途
    gstrSQL = "" & _
        "   Select Level as 层,ID,上级ID,名称 From 诊疗分类目录 where 类型=7" & _
        "   Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        "   Order by 层"
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.RecordCount = 0 Then
        MsgBox "卫材分类不完整！", vbInformation, gstrSysName
        Exit Sub
    End If

    Dim objNode As Node
    Set objNode = tvw分类.Nodes.Add(, , "Root", "所有卫材分类", "Item")
    
    Do While Not rsTemp.EOF
        If rsTemp!层 = 1 Then
            Set objNode = tvw分类.Nodes.Add("Root", 4, "_" & rsTemp!Id, rsTemp!名称, "Item")
        Else
            Set objNode = tvw分类.Nodes.Add("_" & rsTemp!上级ID, 4, "_" & rsTemp!Id, rsTemp!名称, "Item")
        End If
        rsTemp.MoveNext
    Loop
    tvw分类.Nodes("Root").Selected = True
    tvw分类.Nodes("Root").Expanded = True
    mblnBootUp = True
    
    '库房货位
    Load库房货位
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckCount() As Integer
    Dim i As Integer
    For i = 1 To tvw分类.Nodes.Count
        If tvw分类.Nodes(i).Checked Then CheckCount = CheckCount + 1
    Next
End Function

Private Sub Load库房货位()
    Dim rsTemp As ADODB.Recordset
    Dim objNode As Node
    
    On Error GoTo ErrHandle
    '提取所有材料库房货位
    If chk货位.Value = 1 Then
        gstrSQL = "Select Distinct B.编码, B.名称" & _
            " From 材料储备限额 A, 材料库房货位 B " & _
            " Where A.库房货位 = B.名称 And A.库房id = [1] " & _
            " Order By B.编码"
    Else
        gstrSQL = "Select 编码,名称 From 材料库房货位 Order By 编码 "
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取所有材料库房货位", Val(cbo库房.ItemData(cbo库房.ListIndex)))
    
    tvw货位.Nodes.Clear
    Set objNode = tvw货位.Nodes.Add(, , "Root", "所有库房", 1)
    Do While Not rsTemp.EOF
        Set objNode = tvw货位.Nodes.Add("Root", 4, "_" & rsTemp!编码, rsTemp!名称, 1)

        rsTemp.MoveNext
    Loop
    tvw货位.Nodes("Root").Selected = True
    tvw货位.Nodes("Root").Expanded = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub tvw分类_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode tvw分类, Node, Node.Checked
End Sub

Private Sub tvw货位_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode tvw货位, Node, Node.Checked
End Sub


Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function

Private Sub SetParentNode(ByVal objMyTreeView As TreeView, ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '看是否他的兄弟接点是否也全是TRUE，如是，则置其父节点也为TRUE，否则，不管
            intIdx = Node.FirstSibling.Index
            Do While intIdx <> Node.LastSibling.Index
                If objMyTreeView.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = objMyTreeView.Nodes(intIdx).Next.Index
            Loop
            If intIdx = Node.LastSibling.Index Then
                If objMyTreeView.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode objMyTreeView, Node, blnCheck
        End If
    End If
End Sub
