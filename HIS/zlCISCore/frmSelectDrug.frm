VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSelectDrug 
   Caption         =   "药品目录"
   ClientHeight    =   5610
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   8970
   Icon            =   "frmSelectDrug.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   8970
   StartUpPosition =   1  '所有者中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf 
      Height          =   2940
      Left            =   2745
      TabIndex        =   1
      Top             =   135
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   5186
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483628
      GridColor       =   -2147483632
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5400
      TabIndex        =   4
      Top             =   5130
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4200
      TabIndex        =   3
      Top             =   5130
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   2
      Top             =   5130
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   2805
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   4948
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1620
      Top             =   3960
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
            Picture         =   "frmSelectDrug.frx":27A2
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectDrug.frx":2BF4
            Key             =   "book"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectDrug.frx":2D4E
            Key             =   "bookopen"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectDrug.frx":2EA8
            Key             =   "Write"
         EndProperty
      EndProperty
   End
   Begin VB.Image picX 
      Height          =   2925
      Left            =   2505
      MousePointer    =   9  'Size W E
      Top             =   15
      Width           =   45
   End
End
Attribute VB_Name = "frmSelectDrug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Public blnOK As Boolean
Public strReturn As String
Private strSaveKey As String
Private blnFirst As Boolean
Private v_SaveColor As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    If msf.Row < 1 Then Exit Sub
    If Val(msf.TextMatrix(msf.Row, 0)) <= 0 Then Exit Sub
    
    strReturn = ""
    For i = 0 To msf.Cols - 1
        strReturn = strReturn & ";" & msf.TextMatrix(msf.Row, i)
    Next
    If Len(strReturn) > 0 Then strReturn = Mid(strReturn, 2)
    blnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If blnFirst = False Then Exit Sub
    
    '加入初始化设置
    DrawGrid
    DoEvents
    
    RefreshTree
    If tvw.Nodes.Count > 0 Then
        tvw.Nodes(1).Selected = True
        tvw.Nodes(1).Expanded = True
        tvw_NodeClick tvw.Nodes(1)
    End If
    blnFirst = False
End Sub

Private Sub Form_Load()
    blnOK = False
    blnFirst = True
    strReturn = ""
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With tvw
        .Left = 0
        .Top = 0
        .Width = picX.Left
        .Height = Me.ScaleHeight - 450
    End With
    
    With msf
        .Left = picX.Left + picX.Width
        .Top = tvw.Top
        .Width = Me.ScaleWidth - .Left
        .Height = tvw.Height
    End With
        
    With picX
        .Top = msf.Top
        .Height = msf.Height
    End With
    
    cmdHelp.Left = 60
    cmdHelp.Top = tvw.Top + tvw.Height + 90
    cmdOK.Left = msf.Left + msf.Width - 2 * cmdOK.Width - 60 * 2
    cmdOK.Top = cmdHelp.Top
    cmdCancel.Left = cmdOK.Left + cmdOK.Width + 60
    cmdCancel.Top = cmdOK.Top
    
End Sub

Public Function ShowSelectDrug(frmMain As Form, strReturn As String) As Boolean
    ShowSelectDrug = False
    Me.tvw.Nodes.Clear
    
    With frmSelectDrug
        .Show 1, frmMain
        ShowSelectDrug = .blnOK
        strReturn = .strReturn
    End With
End Function

Private Sub msf_DblClick()
    cmdOK_Click
End Sub

Private Sub msf_EnterCell()
    v_SaveColor = msf.CellForeColor
    SelectRow msf
End Sub

Private Sub msf_GotFocus()
    msf_LeaveCell
    msf_EnterCell
End Sub

Private Sub msf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then msf_DblClick
End Sub

Private Sub msf_LeaveCell()
    UnSelectRow msf, v_SaveColor
End Sub

Private Sub DrawGrid()
    msf.Cols = 6
    SetColumnText msf, 0, Array("ID", "药品编码", "名称", "规格", "计算单位", "类别")
    SetColumnWidth msf, Array(0, 1200, 3000, 1200, 900, 1200)
End Sub

Private Sub RefreshTree()
    Dim nodx As Node
    Dim i As Long
    Dim rsTmp As New ADODB.Recordset
    
'    gstrSQL = "select decode(名称,'西成药',5,'中成药',6,7) AS 类型, 编码,名称 from 药品材质分类"
'    gstrSql = "select 编码,名称 from 药品材质分类 order by 编码"
    gstrSql = "SELECT ID,编码,名称 FROM 收费分类目录 WHERE 上级ID IS NULL OR 上级ID=0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    tvw.Nodes.Clear
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            Set nodx = tvw.Nodes.Add(, , "R_" & rsTmp!ID, rsTmp!名称, "book", "bookopen")
            nodx.ExpandedImage = "bookopen"
            nodx.Tag = rsTmp!ID
            If i = 1 Then
                RefreshTreeLoop nodx.Key
            End If
            rsTmp.MoveNext
        Next
    End If
End Sub

Private Sub RefreshTreeLoop(ByVal strKey As String)
    Dim nodx As Node
    Dim i As Long
    Dim rsTmp As New ADODB.Recordset
    '由于诊疗项目目录里无法提取规格所以没有诊疗项目目录里提取
'    If Mid(strKey, 1, 1) = "R" Then
'        gstrSQL = "SELECT ID,编码,名称 FROM 诊疗分类目录 WHERE 类型='" & Mid(strKey, 2) & "' AND (上级ID IS NULL OR 上级ID=0)"
'    Else
'        gstrSQL = "SELECT ID,编码,名称 FROM 诊疗分类目录 WHERE 上级ID=" & Val(Mid(strKey, 2))
'    End If
    
    gstrSql = "SELECT ID,编码,名称 FROM 收费分类目录 WHERE 上级ID=" & Val(Mid(strKey, InStrRev(strKey, "_") + 1))
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        For i = 0 To rsTmp.RecordCount - 1
            Set nodx = tvw.Nodes.Add(strKey, tvwChild, "K_" & rsTmp!ID, rsTmp!名称, "book", "bookopen")
            nodx.ExpandedImage = "bookopen"
            nodx.Tag = rsTmp!ID
            RefreshTreeLoop nodx.Key
            rsTmp.MoveNext
        Next
    End If
End Sub

Private Sub RefreshList(ByVal strKey As String)
    Dim rs As New ADODB.Recordset
    
    msf.Rows = 2
    ClearSpecRowCol msf, 1, Array()
    '由于诊疗项目目录里无法提取规格所以没有诊疗项目目录里提取
'    gstrSQL = "SELECT A.药品ID,A.编码,B.通用名称,A.规格,B.剂量单位 FROM 药品目录 A,药品信息 B WHERE A.药名ID=B.药名ID AND B.用途分类ID=" & Val(strKey)
    gstrSql = "SELECT A.ID,A.编码,A.名称,A.规格,A.计算单位 剂量单位,decode(A.类别,'5','西成药','6','中成药','中草药') 类别 FROM 收费项目目录 A,收费项目别名 B " & vbCrLf & _
                "WHERE A.类别 IN ('5','6','7') AND  (A.撤档时间 IS NULL OR A.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD')) " & vbCrLf & _
                "AND A.Id =B.收费细目id AND A.分类id=" & Val(strKey)
    Call zlDatabase.OpenRecordset(rs, gstrSql, Me.Caption)
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Set msf.DataSource = rs
    End If
    DrawGrid
End Sub

Private Sub msf_LostFocus()
    msf_LeaveCell
    v_SaveColor = msf.CellForeColor
    SelectRow msf, RGB(192, 192, 192), 0
End Sub

Private Sub picX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    On Error Resume Next
    
    picX.Left = picX.Left + X
    If picX.Left < 1000 Then picX.Left = 1000
    If Me.Width - picX.Left - picX.Width < 1000 Then picX.Left = Me.Width - picX.Width - 1000
    Form_Resize
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
'    If strSaveKey <> Node.Tag Then
'        strSaveKey = Node.Tag
        RefreshList CStr(Node.Tag)
        msf_LostFocus
'    End If
End Sub
