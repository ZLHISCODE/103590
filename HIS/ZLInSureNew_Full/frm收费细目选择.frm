VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm收费细目选择 
   Caption         =   "收费细目"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   Icon            =   "frm收费细目选择.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5130
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3690
      TabIndex        =   2
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3690
      TabIndex        =   1
      Top             =   750
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3315
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5847
      _Version        =   393217
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
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
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm收费细目选择.frx":0442
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm收费细目选择.frx":0894
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm收费细目选择.frx":0CE6
            Key             =   "Write"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm收费细目选择"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrID As String
Dim mstr编码 As String
Dim mstr名称 As String
Dim mblnSecceed As Boolean
Dim mblnNode As Boolean

Private Sub cmdCancel_Click()
    mblnSecceed = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If tvw.SelectedItem.Image <> "Item" Then Exit Sub
    mblnSecceed = True
    With tvw.SelectedItem
        mstrID = Mid(.Key, 2)
        mstr编码 = Mid(.Text, 2, InStr(.Text, "】") - 2)
        mstr名称 = Mid(.Text, InStr(.Text, "】") + 1)
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
'        cmdHelp.Left = cmdOK.Left
        tvw.Width = cmdOK.Left - tvw.Left - 200
    End If
End Sub

Public Function ShowTree(strID As String, str编码 As String, str名称 As String) As Boolean
'功能:显示所有收费细目,并得出选择
'参数:strID     返回所选的收费细目的ID
'     str名称   返回所选的收费细目的名称
'返回:有所选择返回True,否则返回False.

    Dim rsTree As New ADODB.Recordset
    Dim strTemp As String
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    mstrID = strID
    
    strSQL = "select 编码,类别 from 收费类别"
    Call OpenRecordset(rsTree, Me.Caption, strSQL)
    
    tvw.Nodes.Clear
    Do Until rsTree.EOF
        tvw.Nodes.Add , , "C" & rsTree("编码"), "【" & rsTree("编码") & "】" & rsTree("类别"), "Root", "Root"
        tvw.Nodes.Add "C" & rsTree("编码"), tvwChild, "K" & rsTree("编码"), "临时"
        rsTree.MoveNext
    Loop
    Me.Show vbModal
    ShowTree = mblnSecceed
    '成功了才返回值
    If mblnSecceed = True Then
        strID = mstrID
        str编码 = mstr编码
        str名称 = mstr名称
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub tvw_DblClick()
    If Not mblnNode Then Exit Sub
    cmdOK_Click
End Sub

Private Sub tvw_Expand(ByVal Node As MSComctlLib.Node)
    Dim rsTree As New ADODB.Recordset
    Dim strTemp As String
    Dim strSQL As String
    
    If Node.Image = "Root" And Left(Node.Child.Key, 1) = "K" Then
    '只对未设置下级的的根节点处理
        
        '删除临时节点
        tvw.Nodes.Remove Node.Child.Key
        
        '再增加新的下级
        rsTree.CursorLocation = adUseClient
        strSQL = "select ID,上级ID,编码,名称,末级 from 收费细目  " & _
            " where 撤档时间 is null or 撤档时间 =to_date('3000-01-01','YYYY-MM-DD') " & _
            " start with 上级ID is null and 类别='" & Mid(Node.Key, 2, 1) & "' connect by prior ID =上级ID"
        Call OpenRecordset(rsTree, Me.Caption, strSQL)
        
        Do Until rsTree.EOF
            strTemp = IIf(rsTree("末级") = 1, "Item", "Write")
            If IsNull(rsTree("上级id")) Then
                tvw.Nodes.Add Node.Key, tvwChild, "_" & rsTree("id"), "【" & rsTree("编码") & "】" & rsTree("名称"), strTemp, strTemp
            Else
                tvw.Nodes.Add "_" & rsTree("上级id"), tvwChild, "_" & rsTree("id"), "【" & rsTree("编码") & "】" & rsTree("名称"), strTemp, strTemp
            End If
            rsTree.MoveNext
        Loop
    End If
End Sub

Private Sub tvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdOK_Click
End Sub

Private Sub tvw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnNode = False
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnNode = True
End Sub
