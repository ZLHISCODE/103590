VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm树型选择 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "树型选择"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4890
   ControlBox      =   0   'False
   Icon            =   "frm树型选择.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3690
      TabIndex        =   2
      Top             =   30
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3690
      TabIndex        =   1
      Top             =   450
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   4305
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   7594
      _Version        =   393217
      HideSelection   =   0   'False
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm树型选择.frx":0442
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm树型选择.frx":0896
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm树型选择.frx":0CEA
            Key             =   "Read"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm树型选择.frx":113E
            Key             =   "Root"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm树型选择"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrID As String
Dim mstr上级ID As String
Dim mstr上级名称 As String
Dim mstr上级编码 As String
Dim mblnSecceed As Boolean
Dim mblnNode As Boolean

Private Sub cmdCancel_Click()
    mblnSecceed = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim nod As Node
    Dim intTemp As Integer
    
    Set nod = tvw.SelectedItem
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
            mstr上级ID = ""
            mstr上级名称 = "无"
            mstr上级编码 = ""
        Else
            intTemp = InStr(.Text, "】")
            mstr上级ID = Mid(.Key, 2)
            mstr上级名称 = Mid(.Text, intTemp + 1)
            mstr上级编码 = Mid(.Text, 2, intTemp - 2)
        End If
    End With
    Unload Me
End Sub

Public Function ShowTree(ByVal strSql As String, str上级ID As String, str上级名称 As String, str上级编码 As String, strID As String, ByVal strCaption As String, ByVal strRoot As String) As Boolean
    Dim rs树型 As New ADODB.Recordset
    
    mstrID = strID
    
    On Error GoTo errHandle
    zlDatabase.OpenRecordset rs树型, strSql, Me.Caption
    
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
        rs树型.MoveNext
    Loop
    If str上级ID = "" Then
        tvw.Nodes("Root").Selected = True
        tvw.Nodes("Root").Expanded = True
    Else
        tvw.Nodes("C" & str上级ID).Selected = True
        tvw.Nodes("C" & str上级ID).EnsureVisible
    End If
    Me.Caption = strCaption
    Me.Show vbModal
    ShowTree = mblnSecceed
    '成功了才返回值
    If mblnSecceed = True Then
        str上级ID = mstr上级ID
        str上级名称 = mstr上级名称
        str上级编码 = mstr上级编码
    End If
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub tvw_DblClick()
    If mblnNode Then cmdOK_Click
End Sub

Private Sub tvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnNode = False
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnNode = True
End Sub
