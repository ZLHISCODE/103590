VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTreeLeafSel 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4890
   Icon            =   "frmTreeLeafSel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3600
      TabIndex        =   2
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3600
      TabIndex        =   1
      Top             =   150
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
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3630
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeLeafSel.frx":0442
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeLeafSel.frx":0896
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeLeafSel.frx":0CEA
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeLeafSel.frx":113E
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeLeafSel.frx":1458
            Key             =   "Man"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTreeLeafSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrID As String
Dim mstr名称 As String
Dim mblnSecceed As Boolean
Dim mblnNode As Boolean
Dim mstrCaption As String

Private Sub cmdCancel_Click()
    mblnSecceed = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If tvw.SelectedItem.Image <> "Item" And tvw.SelectedItem.Image <> "Man" Then Exit Sub
    mblnSecceed = True
    With tvw.SelectedItem
        mstrID = Mid(.Key, 2)
        mstr名称 = .Text
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = mstrCaption
    RestoreWinState Me
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

Public Function ShowTree(ByVal strSQL As String, strID As String, str名称 As String, _
                ByVal strCaption As String, Optional ByVal bln人员 As Boolean = False) As Boolean
'功能:根据SQL语句显示所有项目,并选出某个末级项目
'参数:strSql    SQL语句
'     strID     返回所选的项目的ID
'     str名称   返回所选的项目的名称
'     strCaption   窗口的标题
'返回:有所选择返回True,否则返回False.
    Dim rsTree As New ADODB.Recordset
    Dim strTemp As String, strPre As String
    Dim nod As Node
    
    On Error GoTo errHandle
    mstrID = strID
    mstrCaption = strCaption
'    Call SQLTest(App.ProductName, Me.Caption, strSQL)
    Set rsTree = zlDatabase.OpenSQLRecord(strSQL, "ShowTree")
'    Call SQLTest
    If rsTree.RecordCount = 0 Then
        MsgBox "选择器没找到合适项目。", vbExclamation, gstrSysName
        ShowTree = False
        Exit Function
    End If
    tvw.Nodes.Clear
    Do Until rsTree.EOF
        If bln人员 = True Then
            strTemp = IIF(rsTree("末级") = 1, "Man", "Dept")
        Else
            strTemp = IIF(rsTree("末级") = 1, "Item", "Write")
        End If
        
        If IsNull(rsTree("上级id")) Then
            Set nod = tvw.Nodes.Add(, , "C" & rsTree("id"), rsTree("名称"), strTemp, strTemp)
        Else
            strPre = IIF(rsTree("末级") = 1, "K", "C")
            tvw.Nodes.Add "C" & rsTree("上级id"), tvwChild, strPre & rsTree("id"), rsTree("名称"), strTemp, strTemp
        End If
        nod.Sorted = True
        rsTree.MoveNext
    Loop
    
    If strID <> "" And strID <> "0" Then
        '可能该节点已经被删除了
        On Error Resume Next
        tvw.Nodes("K" & strID).Selected = True
        tvw.Nodes("K" & strID).EnsureVisible
    End If
    Me.Show vbModal
    ShowTree = mblnSecceed
    '成功了才返回值
    If mblnSecceed = True Then
        strID = mstrID
        str名称 = mstr名称
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me
End Sub

Private Sub tvw_DblClick()
    If mblnNode Then cmdOK_Click
End Sub

Private Sub tvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnNode = False
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnNode = True
End Sub
