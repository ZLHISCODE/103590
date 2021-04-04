VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTreeSel 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5700
   Icon            =   "frmTreeSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picOpt 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5685
      Left            =   4365
      ScaleHeight     =   5685
      ScaleWidth      =   1335
      TabIndex        =   1
      Top             =   0
      Width           =   1335
      Begin VB.CommandButton cmdNext 
         Caption         =   "下一个(&N)"
         Height          =   350
         Left            =   30
         TabIndex        =   6
         Top             =   1920
         Width           =   1100
      End
      Begin VB.TextBox txtLocate 
         Height          =   320
         Left            =   30
         TabIndex        =   5
         ToolTipText     =   "查找下一个F3或回车，定位输入框F4"
         Top             =   1470
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   0
         TabIndex        =   3
         Top             =   690
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label lblLocate 
         Caption         =   "查找(&F)"
         Height          =   255
         Left            =   30
         TabIndex        =   4
         Top             =   1230
         Width           =   1095
      End
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeSel.frx":0442
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeSel.frx":0896
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeSel.frx":0CEA
            Key             =   "Book"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeSel.frx":0E44
            Key             =   "BookOpen"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeSel.frx":0F9E
            Key             =   "bm"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeSel.frx":1538
            Key             =   "item"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTreeSel"
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
Private mIntStart As Integer               '记录查询的开始位置
Private mIntEnd As Integer                  '记录最后位置

Dim mstrCaption As String

Dim mblnCheckChild As Boolean

Private Sub cmdCancel_Click()
    mblnSecceed = False
    Unload Me
End Sub

Private Function FullChild(ByVal NodeSour As Node, ByVal NodeFind As Node) As Boolean
'检查要查找的那个对象是不是它的子对象
'根据指定的对象递归查找子对象
Dim i As Long
Dim objNode As Node
Dim blnReturn As Boolean
    
    If NodeSour.Key = NodeFind.Key Then
        FullChild = True
        Exit Function
    Else
        If Not NodeSour.Child Is Nothing Then
            i = NodeSour.Child.FirstSibling.Index
            Set objNode = NodeSour.Child
            While i <= NodeSour.Child.LastSibling.Index
                If objNode.Key = NodeFind.Key Then
                    FullChild = True
                    Exit Function
                Else
                    blnReturn = FullChild(objNode, NodeFind)
                    If blnReturn = True Then
                        FullChild = True
                        Exit Function
                    End If
                    Set objNode = objNode.Next
                    If Not objNode Is Nothing Then
                        i = objNode.Index
                    Else
                        Exit Function
                    End If
                End If
            Wend
        End If
    End If
End Function


Private Sub cmdNext_Click()
    Call txtLocate_KeyPress(vbKeyReturn)
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle
    Dim nod As Node
    Dim i As Integer
    Dim str编码 As String
    
    
    Set nod = tvw.SelectedItem
    
    If tvw.SelectedItem.Key = "Root" Then
        tvw.SelectedItem.Expanded = True
        tvw.SelectedItem.EnsureVisible
        If mblnRoot = False Then
            MsgBox "请选择子分类。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If Not (Trim(mstr上级ID) = "" Or Trim(mstr上级ID) = "0") Then
        If FullChild(tvw.Nodes.Item("C" & mstr上级ID), tvw.SelectedItem) And mblnCheckChild = True Then
            If IsNumeric(mstrID) Then      '只有不是新增的才检查
                If CLng(mstrID) > 0 Then
                    MsgBox "此节点不符合要求，请另选。", vbExclamation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
    End If
    mblnSecceed = True
    With tvw.SelectedItem
        If .Key = "Root" Then
            If mblnRoot = False Then
                MsgBox "请选择子分类。", vbInformation, gstrSysName
                Exit Sub
            End If
            mstr上级ID = ""
            mstr上级名称 = "无"
            mstr上级编码 = ""
        ElseIf .ForeColor = &H8000000C Then
            MsgBox "无该部门的权限！", vbInformation, gstrSysName
            mstr上级ID = ""
            mstr上级名称 = "无"
            mstr上级编码 = ""
            Exit Sub
        Else
            i = InStr(.Text, "】")
            mstr上级ID = Mid(.Key, 2)
            mstr上级名称 = Mid(.Text, i + 1)
            mstr上级编码 = Mid(.Text, 2, i - 2)
        End If
    End With
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txtLocate_KeyPress (vbKeyReturn)
    ElseIf KeyCode = vbKeyF4 Then
        If txtLocate.Visible And txtLocate.Enabled Then txtLocate.SetFocus
    End If
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
    tvw.Width = picOpt.Left - tvw.Left - 200
    
End Sub

Public Function ShowTree(ByVal strSQL As String, str上级ID As String, str上级名称 As String, str上级编码 As String, _
                        strID As String, ByVal strCaption As String, _
                        ByVal strRoot As String, Optional blnRoot As Boolean = True, Optional str原编码 As String, _
                        Optional ByVal IconIndex As Long = 0, _
                        Optional ByVal SelectIconIndex As Long = 0, _
                        Optional ByVal ExpIconIndex As Long = 0, _
                        Optional ByVal blnChild As Boolean = True) As Boolean
'功能:根据SQL语句显示所有项目,并选出某个末级项目
'参数:strSql        SQL语句
'     str上级ID     返回所选的项目的上级ID
'     str上级名称   返回所选的项目的上级名称
'     str上级编码   返回所选的项目的上级编码
'     strID         返回所选的项目的ID
'     strRoot       树根的标题
'     strICO        图标资源的名称
'     strCaption    窗口的标题
'     IconIndex     图标索引
'     SelectIconIndex   选择项的图标索引
'     ExpIconIndex  扩展图标索引
'     blnChild      检查子结点
'返回:有所选择返回True,否则返回False.
On Error GoTo errHandle
    Dim rs树型 As New ADODB.Recordset
    Dim objNode As Node, bln简码 As Boolean, i As Long
    
    
    mblnRoot = blnRoot
    mstrCaption = strCaption
    mblnCheckChild = blnChild
    
    Call zlDatabase.OpenRecordset(rs树型, strSQL, Me.Caption)
'    For i = 0 To rs树型.Fields.Count - 1
'        If rs树型.Fields(i).Name = "简码" Then
'            bln简码 = True
'            Exit For
'        End If
'    Next
    
    tvw.Nodes.Clear
    tvw.Nodes.Add , , "Root", strRoot, "Root", "Root"
    tvw.Nodes("Root").Sorted = True
    Do Until rs树型.EOF
        
        If IsNull(rs树型("上级id")) Then
            Set objNode = tvw.Nodes.Add("Root", tvwChild, "C" & rs树型("id"), "【" & rs树型("编码") & "】" & rs树型("名称"), IIF(IconIndex > 0 And IconIndex < 7, IconIndex, "Write"), IIF(SelectIconIndex > 0 And SelectIconIndex < 7, SelectIconIndex, "Write"))
        Else
            Set objNode = tvw.Nodes.Add("C" & rs树型("上级id"), tvwChild, "C" & rs树型("id"), "【" & rs树型("编码") & "】" & rs树型("名称"), IIF(IconIndex > 0 And IconIndex < 7, IconIndex, "Write"), IIF(SelectIconIndex > 0 And SelectIconIndex < 7, SelectIconIndex, "Write"))
        End If
        objNode.Tag = Nvl(rs树型!编码)
'        If bln简码 Then objNode.Tag = rs树型!简码
        If SelectIconIndex > 0 And SelectIconIndex < 7 Then
            objNode.ExpandedImage = SelectIconIndex
        End If
        objNode.Sorted = True
        rs树型.MoveNext
    Loop
    If str上级ID = "0" Then str上级ID = ""
    If str上级ID = "" Then
        tvw.Nodes("Root").Selected = True
        tvw.Nodes("Root").Expanded = True
    Else
        tvw.Nodes("C" & str上级ID).Selected = True
        tvw.Nodes("C" & str上级ID).EnsureVisible
    End If
    
    mstrID = strID
    mstr上级ID = str上级ID
    mstr原编码 = str原编码
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
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    mIntEnd = 0
    mIntStart = 0
End Sub

Private Sub tvw_DblClick()
    If mblnNode Then cmdOK_Click
End Sub

Private Sub tvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mblnNode Then cmdOK_Click
    End If
End Sub

Private Sub tvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnNode = False
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnNode = True
    mIntStart = tvw.SelectedItem.Index
End Sub

Public Function ShowTreePrivs(ByVal lngOperationID As Long, str上级ID As String, str上级名称 As String, str上级编码 As String) As Boolean
'功能:装入所属部门到tvwMain_S
    Dim nodTmp As Node
    Dim rsDeptID As ADODB.Recordset
    Dim strTemp As String
    strTemp = "Write"
    
    On Error GoTo errHandle
    gstrSQL = "Select Max(Level) as 层,A.ID,A.上级ID,A.名称,'【'||A.编码||'】' 编码,Upper(a.简码) as 简码 " & _
              "From 部门表 A Start With ID IN(Select 部门ID From 部门人员 Where 人员ID=[1]) Connect by Prior 上级ID=ID " & _
              "Group by A.ID,A.上级ID,A.名称,A.编码,a.简码 " & _
              "Order by A.编码,层 Desc"
    Set rsDeptID = zlDatabase.OpenSQLRecord(gstrSQL, Caption, glngUserId)
    With tvw
        .LineStyle = tvwRootLines
        .Sorted = True
        .Nodes.Clear
        Do While Not rsDeptID.EOF
            If IIF(IsNull(rsDeptID!上级ID), 0, rsDeptID!上级ID) = 0 Then
                If .Nodes.Count > 0 Then
                    If FindKey("C" & rsDeptID!ID) = False Then
                        Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称, strTemp, strTemp)
                    Else
                        Set nodTmp = .Nodes("C" & rsDeptID!ID)
                    End If
                Else
                    Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称, strTemp, strTemp)
                End If
            Else
                If FindKey("C" & rsDeptID!ID) = False Then
                    Set nodTmp = .Nodes.Add("C" & rsDeptID!上级ID, tvwChild, "C" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称, strTemp, strTemp)
                Else
                    Set nodTmp = .Nodes("C" & rsDeptID!ID)
                End If
            End If
            nodTmp.Tag = rsDeptID!简码
            nodTmp.ForeColor = &H8000000C
            rsDeptID.MoveNext
        Loop
        rsDeptID.Close
    End With
    '生成子结点
    gstrSQL = "Select ID,上级ID,'【'||编码||'】' 编码,名称,Upper(简码) as 简码 " & _
              "From 部门表 A " & _
              "Start With ID IN(Select 部门ID From 部门人员 Where 人员ID=[1]) Connect by Prior ID=上级ID"
    Set rsDeptID = zlDatabase.OpenSQLRecord(gstrSQL, Caption, glngUserId)
    With tvw
        Do While Not rsDeptID.EOF
            If IIF(IsNull(rsDeptID!上级ID), 0, rsDeptID!上级ID) = 0 Then
                If .Nodes.Count > 0 Then
                    If FindKey("C" & rsDeptID!ID) = False Then
                        Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称, strTemp, strTemp)
                    Else
                        Set nodTmp = .Nodes("C" & rsDeptID!ID)
                    End If
                Else
                    Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称, strTemp, strTemp)
                End If
            Else
                If FindKey("C" & rsDeptID!ID) = False Then
                    Set nodTmp = .Nodes.Add("C" & rsDeptID!上级ID, tvwChild, "C" & rsDeptID!ID, rsDeptID!编码 & rsDeptID!名称, strTemp, strTemp)
                Else
                    Set nodTmp = .Nodes("C" & rsDeptID!ID)
                End If
            End If
            nodTmp.Tag = rsDeptID!简码
            nodTmp.ForeColor = vbBlack
            rsDeptID.MoveNext
        Loop
        rsDeptID.Close
    
        If .Nodes.Count > 0 Then .Nodes(1).Selected = True
    
    End With
    Me.Show vbModal
    ShowTreePrivs = mblnSecceed
    If mblnSecceed = True Then
        str上级ID = mstr上级ID
        str上级名称 = mstr上级名称
        str上级编码 = mstr上级编码
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FindKey(ByVal strKey As String) As Boolean
    Dim nodTmp As Node
    For Each nodTmp In tvw.Nodes
        If nodTmp.Key = strKey Then
            FindKey = True
            Exit Function
        End If
    Next
End Function

Private Sub txtLocate_GotFocus()
    zlControl.TxtSelAll txtLocate
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    Dim i As Long, lngStart As Long
    
    If KeyAscii = vbKeyReturn Then
        If txtLocate.Tag <> txtLocate.Text Then
            lblLocate.Tag = ""
            txtLocate.Tag = txtLocate.Text
        End If
        
        lngStart = Val("" & lblLocate.Tag) + 1

        If lngStart >= tvw.Nodes.Count Then lngStart = 1
        
        For i = tvw.Nodes.Count To 1 Step -1    '查找最后一个的位置
            If UCase(tvw.Nodes(i).Text) Like "*" & UCase(txtLocate.Text) & "*" Or UCase(tvw.Nodes(i).Tag) Like "*" & UCase(txtLocate.Text) & "*" Then
                mIntEnd = i
                Exit For
            End If
        Next
        
        If lngStart - 1 = mIntEnd Then  '如果是最后一个则查询第一个
            lngStart = 1
        End If
        If mIntStart < Val(lblLocate.Tag) And mIntStart <> 0 Then '重新选择了查询的位置
            lngStart = mIntStart + 1
            mIntStart = 0
        End If
        For i = lngStart To tvw.Nodes.Count
            If tvw.Nodes(i).Text Like "*" & txtLocate.Text & "*" Or tvw.Nodes(i).Tag Like "*" & UCase(txtLocate.Text) & "*" Then
                Call tvw.Nodes(i).EnsureVisible
                tvw.Nodes(i).Selected = True
                lblLocate.Tag = i
                tvw.SetFocus
                Exit For
            End If
            If i = tvw.Nodes.Count Then
                MsgBox "没有查询到你所输入的信息，请重新输入！", vbInformation, gstrSysName
                txtLocate.Text = ""
                txtLocate.SetFocus
            End If
        Next
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub
