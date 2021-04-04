VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClassSel 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "疾病编码分类"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4890
   Icon            =   "frmClassSel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3690
      TabIndex        =   1
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3690
      TabIndex        =   0
      Top             =   660
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3315
      Left            =   120
      TabIndex        =   2
      Top             =   450
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5847
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3570
      Top             =   2700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClassSel.frx":0442
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClassSel.frx":0896
            Key             =   "Root"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmClassSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrID As String
Dim mstr上级ID As String
Dim mstr上级名称 As String
Dim mstr编码范围 As String
Dim mblnRoot As Boolean '允许选择根结点

Dim mblnNode As Boolean
Dim mblnSecceed As Boolean

Private Sub cmdCancel_Click()
    mblnSecceed = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim nod As Node
    Dim i As Integer
    Dim str编码 As String
    
    Set nod = tvw.SelectedItem
    
    '判断是否本级
    Do Until nod.Key = "Root"
        If mstrID = Mid(nod.Key, 2) Then
'            MsgBox "此节点不符合要求，请另选。", vbExclamation, gstrSysName
            Exit Sub
        End If
        Set nod = nod.Parent
    Loop
    mblnSecceed = True
    With tvw.SelectedItem
        If .Key = "Root" Then
            If mblnRoot = False Then Exit Sub
            mstr上级ID = ""
            mstr上级名称 = "无"
            mstr编码范围 = ""
        Else
            mstr上级ID = Mid(.Key, 2)
            mstr上级名称 = .Text
            mstr编码范围 = .Tag
        End If
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tvw.Top = 100
    tvw.Left = 100
    tvw.Height = ScaleHeight - 200
    If Me.ScaleWidth > 3000 Then
        cmdOK.Left = ScaleWidth - cmdOK.Width - 200
        cmdCancel.Left = cmdOK.Left
        tvw.Width = cmdOK.Left - tvw.Left - 200
    End If
End Sub

Public Function ShowTree(str上级ID As String, str上级名称 As String, str编码范围 As String _
    , ByVal str编码类别 As String, ByVal strID As String, Optional blnRoot As Boolean = True) As Boolean
'功能:根据SQL语句显示所有项目,并选出某个末级项目
'参数:str上级ID     返回所选的项目的上级ID
'     str上级名称   返回所选的项目的上级名称
'     strID         所选的项目的ID，用以判断是否在其下级中
'     blnRoot       是否允许选择根节点
'返回:有所选择返回True,否则返回False.
    
    Dim rsTemp As New ADODB.Recordset
    Dim nodTemp As Node
    
    mblnRoot = blnRoot
    mstrID = strID
    
    On Error GoTo ErrHandle
    gstrSQL = "select level,ID,上级ID,序号,名称,编码范围 from 疾病编码分类 where 类别=[1] " & _
        " start with 上级ID is null connect by prior id=上级ID order by level,序号"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, str编码类别)
        
    tvw.Nodes.Clear
    tvw.Nodes.Add , , "Root", "编码类别", "Root", "Root"
    Do Until rsTemp.EOF
        If IsNull(rsTemp("上级ID")) Then
            Set nodTemp = tvw.Nodes.Add("Root", tvwChild, "K" & rsTemp("ID"), "【" & rsTemp("序号") & "】" & Trim(rsTemp("名称")), "Write", "Write")
        Else
            Set nodTemp = tvw.Nodes.Add("K" & rsTemp("上级ID"), tvwChild, "K" & rsTemp("ID"), "【" & rsTemp("序号") & "】" & Trim(rsTemp("名称")), "Write", "Write")
        End If
        nodTemp.Tag = IIF(IsNull(rsTemp("编码范围")), "", rsTemp("编码范围"))
        rsTemp.MoveNext
        
    Loop
    If str上级ID = "" Then
        tvw.Nodes("Root").Selected = True
        tvw.Nodes("Root").Expanded = True
    Else
        tvw.Nodes("K" & str上级ID).Selected = True
        tvw.Nodes("K" & str上级ID).EnsureVisible
    End If
    
    Me.Show vbModal
    ShowTree = mblnSecceed
    '成功了才返回值
    If mblnSecceed = True Then
        str上级ID = mstr上级ID
        str上级名称 = mstr上级名称
        str编码范围 = mstr编码范围
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub tvw_DblClick()
    If mblnNode Then cmdOK_Click
End Sub

Private Sub tvw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnNode = False
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnNode = True
End Sub
