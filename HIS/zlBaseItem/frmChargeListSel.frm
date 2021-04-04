VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeListSel 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "收费细目"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4890
   Icon            =   "frmChargeListSel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
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
            Picture         =   "frmChargeListSel.frx":0442
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeListSel.frx":0894
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeListSel.frx":0CE6
            Key             =   "Write"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmChargeListSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrID As String
Dim mstr名称 As String
Dim mblnSecceed As Boolean
Dim mblnNode As Boolean
Dim mstrWhere As String

Private Sub cmdCancel_Click()
    mblnSecceed = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle
    If tvw.SelectedItem.Image <> "Item" Then Exit Sub
    If Trim(mstrWhere) <> "" Then
        If tvw.SelectedItem.Tag <> "1" Then
            MsgBox "请选择蓝色的变价项目！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    mblnSecceed = True
    With tvw.SelectedItem
        mstrID = Mid(.Key, 2)
        mstr名称 = Mid(.Text, InStr(.Text, "】") + 1)
    End With
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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

Public Function ShowTree(strID As String, str名称 As String, blnAllDrug As Boolean, Optional strWhere As String = "") As Boolean
'功能:显示所有收费细目,并得出选择
'参数:strID     返回所选的收费细目的ID
'     str名称   返回所选的收费细目的名称
'     blnAllDrug 是否包含药品和卫材
'返回:有所选择返回True,否则返回False.
On Error GoTo errHandle

    Dim rsTree As New ADODB.Recordset
    Dim strTemp As String
    Dim strSQL As String
    
    strSQL = "select 编码,类别 名称 from 收费类别"
    If Not blnAllDrug Then
        strSQL = strSQL & " where 编码 Not In('4','5','6','7') "
    End If
    mstrID = strID
    mstrWhere = strWhere
    
    Call zldatabase.OpenRecordset(rsTree, strSQL, Me.Caption)
    tvw.Nodes.Clear
    Do Until rsTree.EOF
        tvw.Nodes.Add , , "C" & rsTree("编码"), "【" & rsTree("编码") & "】" & rsTree("名称"), "Root", "Root"
        tvw.Nodes.Add "C" & rsTree("编码"), tvwChild, "K" & rsTree("编码"), "临时"
        rsTree.MoveNext
    Loop
    Me.Show vbModal
    ShowTree = mblnSecceed
    '成功了才返回值
    If mblnSecceed = True Then
        strID = mstrID
        str名称 = mstr名称
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub tvw_DblClick()
    If Not mblnNode Then Exit Sub
    cmdOK_Click
End Sub

Private Sub tvw_Expand(ByVal Node As MSComctlLib.Node)
On Error GoTo errHandle
    Dim rsTree As New ADODB.Recordset
    Dim strTemp As String
    Dim ObjItem As Node
    
    If Node.Image = "Root" And Left(Node.Child.Key, 1) = "K" Then
    '只对未设置下级的的根节点处理
        
        '删除临时节点
        tvw.Nodes.Remove Node.Child.Key
        
        '再增加新的下级
        rsTree.CursorLocation = adUseClient
        gstrSQL = "select ID,上级ID,编码,名称,末级,是否变价 from 收费细目  " & _
            " where (撤档时间 is null or 撤档时间 =to_date('3000-01-01','YYYY-MM-DD')) and 是否变价 <> 1 " & _
            " start with 上级ID is null and 类别=[1] connect by prior ID =上级ID"
        Set rsTree = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Mid(Node.Key, 2, 1))
        
        Do Until rsTree.EOF
            strTemp = IIF(rsTree("末级") = 1, "Item", "Write")
            If IsNull(rsTree("上级id")) Then
                Set ObjItem = tvw.Nodes.Add(Node.Key, tvwChild, "_" & rsTree("id"), "【" & rsTree("编码") & "】" & rsTree("名称"), strTemp, strTemp)
                ObjItem.Tag = rsTree("是否变价")
                If Trim(mstrWhere) <> "" Then
                    If rsTree("是否变价") = 1 Then
                        ObjItem.ForeColor = RGB(0, 0, 255)
                    Else
                        ObjItem.ForeColor = RGB(0, 0, 0)
                    End If
                End If
            Else
                Set ObjItem = tvw.Nodes.Add("_" & rsTree("上级id"), tvwChild, "_" & rsTree("id"), "【" & rsTree("编码") & "】" & rsTree("名称"), strTemp, strTemp)
                ObjItem.Tag = rsTree("是否变价")
                If Trim(mstrWhere) <> "" Then
                    If rsTree("是否变价") = 1 Then
                        ObjItem.ForeColor = RGB(0, 0, 255)
                    Else
                        ObjItem.ForeColor = RGB(0, 0, 0)
                    End If
                End If
            End If
            rsTree.MoveNext
        Loop
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
