VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAuditItemTypeSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "病案审查项目分类选择"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   2
      Left            =   90
      ScaleHeight     =   3495
      ScaleWidth      =   4110
      TabIndex        =   2
      Top             =   105
      Width           =   4110
      Begin MSComctlLib.TreeView tvwAuditType 
         Height          =   3465
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   6112
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "ils16"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1815
      TabIndex        =   1
      Top             =   3930
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3105
      TabIndex        =   0
      Top             =   3930
      Width           =   1100
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   5865
      Y1              =   3795
      Y2              =   3795
   End
   Begin VB.Line Line1 
      X1              =   -15
      X2              =   5865
      Y1              =   3780
      Y2              =   3780
   End
End
Attribute VB_Name = "frmAuditItemTypeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintTypeID          As Integer
Private mblnCancel          As Boolean  '确定 or 取消
Private mlngLeft            As Long
Private mlngTop             As Long
Private mlngProjectID       As Long '方案ID
Private mstrProjectName      As String '方案名称

Public Property Let lngLeft(ByVal vlngLeft As Long)
    mlngLeft = vlngLeft
End Property

Public Property Let lngTop(ByVal vlngTop As Long)
    mlngTop = vlngTop
End Property

Public Property Get blnCancel() As Boolean
    blnCancel = mblnCancel
End Property

Public Property Let blnCancel(ByVal vNewValue As Boolean)
    mblnCancel = vNewValue
End Property

Public Property Get intTypeID() As Integer
    intTypeID = mintTypeID
End Property

Public Property Let intTypeID(ByVal vNewValue As Integer)
    mintTypeID = vNewValue
End Property


Public Property Get lngProjectID() As Long
    lngProjectID = mlngProjectID
End Property

Public Property Let lngProjectID(ByVal vNewValue As Long)
    mlngProjectID = vNewValue
End Property

Public Property Get strProjectName() As String
    strProjectName = mstrProjectName
End Property

Public Property Let strProjectName(ByVal vNewValue As String)
    mstrProjectName = vNewValue
End Property

Private Sub CmdCancel_Click()
On Error GoTo ErrH
    
    blnCancel = True
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub CmdOK_Click()
    Dim nTmpNode As Node
    On Error GoTo ErrH
    
    blnCancel = False
    If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
        mintTypeID = -1
        mlngProjectID = Replace(tvwAuditType.SelectedItem.Key, "Root", "")
        strProjectName = tvwAuditType.SelectedItem.Text
    Else
        mintTypeID = Val(Mid(tvwAuditType.SelectedItem.Key, 2))
        Set nTmpNode = tvwAuditType.SelectedItem
        While Not nTmpNode.Parent Is Nothing
            Set nTmpNode = nTmpNode.Parent
        Wend
        
        lngProjectID = Replace(nTmpNode.Key, "Root", "")
        strProjectName = nTmpNode.Text
    End If
    
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    blnCancel = True
    Me.Top = mlngTop
    Me.Left = mlngLeft
    Call InitTreeView
End Sub

'==============================================================================
'=功能： 病案审查分类
'==============================================================================
Private Sub InitTreeView()
    Dim rsTree      As ADODB.Recordset
    Dim nod         As Node
    Dim i           As Long
    Dim FirstKey    As String
    Dim v           As Variant
    Dim intStartid As Integer

    On Error GoTo ErrH

    'Tree的初始化
    Set tvwAuditType.ImageList = GetImageList(16)
    tvwAuditType.Nodes.Clear

    '添加根节点
    gstrSQL = "Select ID,名称,启用时间 From 病案审查方案"
    Set rsTree = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name)
    Do Until rsTree.EOF
        If zlCommFun.NVL(rsTree!启用时间) <> "" Then
            intStartid = rsTree!ID
        End If
        Set nod = tvwAuditType.Nodes.Add(, , "Root" & rsTree!ID, zlCommFun.NVL(rsTree!名称, "默认方案"), 20, 20)
        nod.Expanded = True
            
        rsTree.MoveNext
    Loop
    
'    Set nod = tvwAuditType.Nodes.Add(, , "Root", "分类", 20, 20)
'    nod.Expanded = True
    gstrSQL = "SELECT /*+ rule */ id,上级ID,方案ID,编码,名称 FROM 病案审查分类 START WITH 上级ID is NULL CONNECT BY PRIOR ID = 上级ID"
    Set rsTree = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name)
    rsTree.Sort = "编码"
    i = 1
    Do Until rsTree.EOF
        '添加子节点
        Set nod = tvwAuditType.Nodes.Add(IIf("" & rsTree("上级ID") = "", "Root" & rsTree("方案ID"), "A" & rsTree("上级ID")), tvwChild, "A" & rsTree("ID"), "【" + "" & rsTree("编码") + "】" + "" & rsTree("名称"), 23, 24)
        If i = 1 Then FirstKey = nod.Key
        If FirstKey = nod.Key Then i = 2
        If FirstKey = "" And i = 1 Then FirstKey = nod.Key: i = 2
        rsTree.MoveNext
    Loop
    FirstKey = "A" & mintTypeID
    For Each v In tvwAuditType.Nodes
        If v.Key = FirstKey Then
            '设置选中
            v.Selected = True
            v.EnsureVisible
        End If
    Next
    If tvwAuditType.SelectedItem Is Nothing Then
        tvwAuditType.Nodes("Root" & intStartid).Selected = True
        tvwAuditType.Nodes("Root" & intStartid).Bold = True
        tvwAuditType.Nodes("Root" & intStartid).Tag = 1
    End If
    DoEvents
'    tvwAuditType_NodeClick tvwAuditType.SelectedItem
    
'    gstrSQL = "SELECT /*+ rule */ id,上级ID,编码,名称 FROM 病案审查分类 START WITH 上级ID is NULL CONNECT BY PRIOR ID = 上级ID"
'    Set rsTree = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
'    i = 1
'    Do Until rsTree.EOF
'        '添加子节点
'        Set nod = tvwAuditType.Nodes.Add(IIf("" & rsTree("上级ID") = "", "Root", "A" & rsTree("上级ID")), tvwChild, "A" & rsTree("ID"), "[" + rsTree("编码") + "]" + rsTree("名称"), 23, 24)
'        If i = 1 Then FirstKey = nod.Key
'        If FirstKey = nod.Key Then i = 2
'        If FirstKey = "" And i = 1 Then FirstKey = nod.Key: i = 2
'        rsTree.MoveNext
'    Loop
    FirstKey = "A" & mintTypeID
    For Each v In tvwAuditType.Nodes
        If v.Key = FirstKey Then
            '设置选中
            v.Selected = True
            v.EnsureVisible
        End If
    Next
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tvwAuditType_DblClick()
    Call CmdOK_Click
End Sub

