VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm树型选择 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4890
   Icon            =   "frm树型选择.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
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
      Indentation     =   494
      LabelEdit       =   1
      Sorted          =   -1  'True
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
            Picture         =   "frm树型选择.frx":0442
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm树型选择.frx":0896
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm树型选择.frx":0CEA
            Key             =   "End"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm树型选择"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Dim mstrID As String
Dim mstr上级ID As String
Dim mstr上级名称 As String
Dim mstr上级编码 As String
Dim mstr原编码 As String
Dim mblnSecceed As Boolean
Dim mblnNode As Boolean
Dim mblnRoot As Boolean '允许选择根结点
Dim mblnSel末级 As Boolean

Dim mstrCaption As String

Private Sub cmdCancel_Click()
    mblnSecceed = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim nod As Node
    Dim i As Integer
    Dim str编码 As String
    
    
    Set nod = tvw.SelectedItem
    
    If mblnSel末级 And nod.Tag <> "1" Then Exit Sub
    
    If mstr原编码 <> "" Then
        If nod.Key = "Root" Then
            str编码 = ""
        Else
            str编码 = Mid(nod.Text, 2, InStr(nod.Text, "】") - 2)
        End If
        'mstrID为空表示新增，这时不考虑上下级关系
        If mstr原编码 = Mid(str编码, 1, Len(mstr原编码)) And mstrID <> "" Then Exit Sub
    End If
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
            mstr上级编码 = ""
        Else
            i = InStr(.Text, "】")
            mstr上级ID = Mid(.Key, 2)
            mstr上级名称 = Mid(.Text, i + 1)
            mstr上级编码 = Mid(.Text, 2, i - 2)
        End If
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = mstrCaption
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

Public Function ShowTree(ByVal strSQL As String, str上级ID As String, str上级名称 As String, str上级编码 As String, ByVal strID As String, ByVal strCaption As String, _
    ByVal strRoot As String, Optional blnRoot As Boolean = True, Optional str原编码 As String, Optional blnSel末级 As Boolean = False) As Boolean
'功能:根据SQL语句显示所有项目,并选出某个末级项目
'参数:strSql        SQL语句
'     str上级ID     返回所选的项目的上级ID
'     str上级名称   返回所选的项目的上级名称
'     str上级编码   返回所选的项目的上级编码
'     strID         返回所选的项目的ID
'     strRoot       树根的标题
'     strICO        图标资源的名称
'     strCaption    窗口的标题
'返回:有所选择返回True,否则返回False.
    
    Dim rs树型 As New ADODB.Recordset
    
    mblnRoot = blnRoot
    mstrCaption = strCaption
    
    Set rs树型 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    mblnSel末级 = blnSel末级
    
    tvw.Nodes.Clear
    tvw.Nodes.Add , , "Root", strRoot, "Root", "Root"
    tvw.Nodes("Root").Sorted = True
    Do Until rs树型.EOF
        
        If IsNull(rs树型("上级id")) Then
            tvw.Nodes.Add "Root", tvwChild, "C" & rs树型("id"), "【" & rs树型("编码") & "】" & rs树型("名称"), "Write", "Write"
        Else
            tvw.Nodes.Add "C" & rs树型("上级id"), tvwChild, "C" & rs树型("id"), "【" & rs树型("编码") & "】" & rs树型("名称"), "Write", "Write"
        End If
        tvw.Nodes("C" & rs树型("id")).Sorted = True
        If blnSel末级 Then
            tvw.Nodes("C" & rs树型("id")).Tag = Val(IIf(IsNull(rs树型("末级")), "0", rs树型("末级")))
            If tvw.Nodes("C" & rs树型("id")).Tag = "1" Then
                tvw.Nodes("C" & rs树型("id")).Image = "End"
                tvw.Nodes("C" & rs树型("id")).SelectedImage = "End"
            End If
        End If
        rs树型.MoveNext
    Loop
    If str上级ID = "" Then
        tvw.Nodes("Root").Selected = True
        tvw.Nodes("Root").Expanded = True
    Else
        tvw.Nodes("C" & str上级ID).Selected = True
        tvw.Nodes("C" & str上级ID).EnsureVisible
    End If
    
    mstrID = strID
    mstr原编码 = str原编码
    Me.Show vbModal
    ShowTree = mblnSecceed
    '成功了才返回值
    If mblnSecceed = True Then
        str上级ID = mstr上级ID
        str上级名称 = mstr上级名称
        str上级编码 = mstr上级编码
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
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
