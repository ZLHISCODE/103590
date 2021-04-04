VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSelVis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择所见项"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   ControlBox      =   0   'False
   Icon            =   "frmSelVis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraList 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton cmdTab 
         Caption         =   "标记图"
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComctlLib.TreeView tvwItem 
         Height          =   1995
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   3519
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "iLsTree"
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ListView lvwSubItem 
      Height          =   2295
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4048
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      SmallIcons      =   "iLsTree"
      ColHdrIcons     =   "iLsTree"
      ForeColor       =   -2147483641
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "编码"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "英文名"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "替换域"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "类型"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "长度"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "小数"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "单位"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "表示法"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "性别域"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "数值域"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "正常域"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "初始值"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "文字表述"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "空值文字"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "临床意义"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList iLsTree32 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelVis.frx":000C
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelVis.frx":08E6
            Key             =   "Attr"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelVis.frx":0C00
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iLsTree 
      Left            =   45
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelVis.frx":0D5A
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelVis.frx":0EB4
            Key             =   "Attr"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelVis.frx":11CE
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgY 
      Height          =   5115
      Left            =   2040
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "frmSelVis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsComLib As New zl9ComLib.clsComLib
Private clsDatabase As New zl9ComLib.clsDatabase

Private iCurrTab As Integer
Public ItemID As String

Private Sub cmdTab_Click(Index As Integer)
    If iCurrTab = Index Then Me.tvwItem(Index).SetFocus: Exit Sub
    iCurrTab = Index
        
    Form_Resize
    If tvwItem(iCurrTab).Nodes.Count > 0 Then
        Set tvwItem(iCurrTab).SelectedItem = tvwItem(iCurrTab).Nodes(1)
        tvwItem(iCurrTab).SetFocus
        tvwItem_NodeClick iCurrTab, tvwItem(iCurrTab).SelectedItem
    Else
        lvwSubItem.ListItems.Clear
    End If
End Sub

Private Sub Form_Activate()
    ItemID = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Me.Hide
End Sub

Private Sub Form_Load()
    CreateItemTree
    
    On Error Resume Next
    iCurrTab = 1
    Set tvwItem(1).SelectedItem = tvwItem(1).Nodes(1)
    tvwItem_NodeClick iCurrTab, tvwItem(iCurrTab).SelectedItem
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With lvwSubItem
        .Left = imgY.Left + imgY.Width: .Top = 0
        .Width = Me.ScaleWidth - .Left: .Height = Me.ScaleHeight - .Top
        .Refresh
    End With
    
    '   显示选项卡
    ShowList imgY.Left
End Sub

Private Sub imgY_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Not Button = vbLeftButton Then Exit Sub
            
    imgY.Left = imgY.Left + x
    If imgY.Left < 1000 Then imgY.Left = 1000
    If imgY.Left > 3000 Then imgY.Left = 3000
    
    Form_Resize
End Sub

'创建所见项分类及其项目的TreeView
Private Sub CreateItemTree()
    Dim rsItem As New ADODB.Recordset
    Dim sCurID As String
    Dim iStackPoint As Integer '堆栈指针
    Dim aStack() As String '堆栈
    Dim TmpNode As Node
    Dim i As Integer, AttrID As String
    
    '从诊治所见性质中提取
    clsDatabase.OpenRecordset rsItem, "Select * From 诊治所见性质 Order By 编码", ""
    Do While Not rsItem.EOF
        Load cmdTab(cmdTab.Count)
        With cmdTab(cmdTab.Count - 1)
            .Caption = rsItem("名称") '+ IIf(rsItem("固定") = 1, "（只读）", "")
            .Tag = rsItem("固定") & "-" & rsItem("编码")
            .ZOrder 0
            .Visible = True
        End With
        Load tvwItem(tvwItem.Count)
        tvwItem(tvwItem.Count - 1).Visible = True
        
        rsItem.MoveNext
    Loop
    
    For i = 1 To cmdTab.Count - 1
        AttrID = Mid(cmdTab(i).Tag, InStr(cmdTab(i).Tag, "-") + 1)
    
        ReDim aStack(0)
        aStack(0) = ""
        iStackPoint = 0
        
        Do While iStackPoint > -1
            sCurID = aStack(iStackPoint)
            '添加下级所见项分类
            gstrSql = "Select * From 诊治所见分类 Where 上级ID" + IIf(sCurID = "", " is null ", "=[1] ") + " And 性质=[2]"
            Set rsItem = zldatabase.OpenSQLRecord(gstrSql, "查询所见项目分类", sCurID, AttrID)
                        
            '该分类的下级已处理，将其从堆栈中弹出
            iStackPoint = iStackPoint - 1
            
            Do While Not rsItem.EOF
                If sCurID = "" Then
                    Set TmpNode = tvwItem(i).Nodes.Add(, , "Key" & rsItem("ID"), rsItem("名称"), "Class")
                Else
                    Set TmpNode = tvwItem(i).Nodes.Add("Key" + sCurID, tvwChild, "Key" & rsItem("ID"), rsItem("名称"), "Class")
                End If
                TmpNode.Tag = rsItem("性质") & "||" & rsItem("编码") & "||" & rsItem("名称") & "||" & rsItem("简码")
                
                '将新分类压入堆栈
                iStackPoint = iStackPoint + 1
                ReDim Preserve aStack(iStackPoint)
                aStack(iStackPoint) = rsItem("ID")
                
                rsItem.MoveNext
            Loop
        Loop
    Next
End Sub


Private Sub ShowSubItem(ByVal NodeID As String, ByVal AttributeID As String)
    Dim rsItem As New ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim sSQL As String
    lvwSubItem.ListItems.Clear
    '添加下级所见项目
    sSQL = "Select ID,编码,中文名,nvl(英文名,' '),nvl(替换域,1),nvl(类型,0)," + _
       "nvl(长度,10),nvl(小数,0),nvl(单位,' '),nvl(表示法,0),nvl(性别域,0)," + _
       "nvl(数值域,' '),nvl(正常域,' '),nvl(初始值,' '),nvl(文字表述,1),nvl(空值文字,' '),nvl(临床意义,' ') " + _
       "From 诊治所见项目 Where " + IIf(NodeID = "", "性质=[1] And 分类ID is null ", "分类ID=[2] ")
    Set rsItem = zldatabase.OpenSQLRecord(sSQL, "查询所见项目", AttributeID, NodeID)
        
    Do While Not rsItem.EOF
        Set tmpItem = lvwSubItem.ListItems.Add(, "Item" & rsItem(0), rsItem(2), "Item", "Item")
        tmpItem.SubItems(1) = rsItem(0)
        tmpItem.SubItems(3) = rsItem(3)
        tmpItem.SubItems(4) = rsItem(4)
        tmpItem.SubItems(5) = rsItem(5)
        tmpItem.SubItems(6) = rsItem(6)
        tmpItem.SubItems(7) = rsItem(7)
        tmpItem.SubItems(8) = rsItem(8)
        tmpItem.SubItems(9) = rsItem(9)
        tmpItem.SubItems(10) = rsItem(10)
        tmpItem.SubItems(11) = rsItem(11)
        tmpItem.SubItems(12) = rsItem(12)
        tmpItem.SubItems(13) = rsItem(13)
        tmpItem.SubItems(14) = rsItem(14)
        tmpItem.SubItems(15) = rsItem(15)
        tmpItem.SubItems(16) = rsItem(16)
        
        rsItem.MoveNext
    Loop
End Sub

Private Sub lvwSubItem_DblClick()
    If Not lvwSubItem.SelectedItem Is Nothing Then
        ItemID = lvwSubItem.SelectedItem.SubItems(1)
        Me.Hide
    End If
End Sub

Private Sub tvwItem_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
    If Node Is Nothing Then Exit Sub
    If Node.Key Like "Key_*" Then
        ShowSubItem "", Mid(Node.Key, 5)
    Else
        ShowSubItem Mid(Node.Key, 4), ""
    End If
End Sub

Private Sub ShowList(ByVal Width As Long, Optional ByVal Top As Long = -1)
    Dim i As Integer
    With fraList
        .Left = 0: .Top = 0
        .Width = Width
        .Height = Me.ScaleHeight - .Top
        .Visible = True
    End With
    For i = 1 To tvwItem.Count - 1
        tvwItem(i).Visible = IIf(i = iCurrTab, True, False)
        With cmdTab(i)
            If i <= iCurrTab Then
                .Top = (i - 1) * (cmdTab(0).Height - 15)
            Else
                .Top = fraList.Height - (tvwItem.Count - i) * (cmdTab(0).Height - 15)
            End If
            
            .Width = fraList.Width
            .Left = 0
            
            .Visible = True
        End With
    Next
    
    With tvwItem(iCurrTab)
        .Left = 0
        .Top = cmdTab(iCurrTab).Top + cmdTab(iCurrTab).Height
        .Width = fraList.Width
        .Height = fraList.Height - (tvwItem.Count - iCurrTab - 1) * (cmdTab(0).Height - 15) - .Top
    End With
End Sub


