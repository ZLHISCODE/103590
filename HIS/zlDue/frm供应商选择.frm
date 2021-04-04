VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm供应商选择 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "供应商选择"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4155
      TabIndex        =   2
      Top             =   60
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4155
      TabIndex        =   1
      Top             =   435
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvwList 
      Height          =   4410
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7779
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   4380
      Top             =   1695
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
            Picture         =   "frm供应商选择.frx":0000
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商选择.frx":0458
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商选择.frx":08B0
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商选择.frx":0D04
            Key             =   "No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm供应商选择.frx":115C
            Key             =   "Write"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm供应商选择"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mselStr As String
Private mstrPrivs As String
Private msngDownX As Single, msngDownY As Single

Public Function SelDept(ByVal strPrivs As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:供应商选择
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-18 14:46:41
    '-----------------------------------------------------------------------------------------------------------
    Dim rstTemp As New ADODB.Recordset, strSql As String
    mstrPrivs = strPrivs
    
    tvwList.Nodes.Clear
    tvwList.Nodes.Add , , "Root", "所有供应商", 1
    Set tvwList.SelectedItem = tvwList.Nodes("Root")
    tvwList.SelectedItem.Expanded = True
    tvwList.SelectedItem.Sorted = True
    Dim str权限 As String
    
    str权限 = " and (末级<>1 or (末级=1 " & zl_获取站点限制() & "  and " & Get分类权限(mstrPrivs) & ")) "
    strSql = "" & _
        "   Select ID,上级ID,编码,名称,末级 " & _
        "   From 供应商 " & _
        "   Where (撤档时间 is null or  撤档时间=TO_DATE('3000-1-1','yyyy-MM-dd'))" & str权限 & _
        "   start with 上级ID is null connect by prior ID =上级ID"
    On Error GoTo errHandle
    zlDatabase.OpenRecordset rstTemp, strSql, Me.Caption
    
    While Not rstTemp.EOF
        If IsNull(rstTemp!上级ID) Then
            tvwList.Nodes.Add "Root", tvwChild, "P" & rstTemp!ID, "[" & rstTemp!编码 & "]" & rstTemp!名称, IIf(rstTemp!末级 <> 1, 5, 2)
        Else
            tvwList.Nodes.Add "P" & rstTemp!上级ID, tvwChild, "P" & rstTemp!ID, "[" & rstTemp!编码 & "]" & rstTemp!名称, IIf(rstTemp!末级 <> 1, 5, 2)
        End If
        rstTemp.MoveNext
    Wend
    Me.Show vbModal
    SelDept = mselStr
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdCancel_Click()
    mselStr = ""
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    If tvwList.SelectedItem.Image = 2 Then
        mselStr = Mid(tvwList.SelectedItem.Key, 2) & "," & tvwList.SelectedItem.Text
        Me.Hide
    End If
End Sub

Private Sub tvwList_DblClick()
    If tvwList.HitTest(msngDownX, msngDownY) Is Nothing Then Exit Sub
    If tvwList.SelectedItem.Image = 2 Then
        mselStr = Mid(tvwList.SelectedItem.Key, 2) & "," & Mid(tvwList.SelectedItem.Text, InStr(tvwList.SelectedItem.Text, "】") + 1)
        Me.Hide
    End If
End Sub

Private Sub tvwList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownX = X
    msngDownY = Y
End Sub

Private Sub tvwList_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Image = 2 Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub
