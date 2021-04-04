VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmManSelect 
   Caption         =   "人员选择"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6645
   Icon            =   "frmManSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3660
      TabIndex        =   5
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5130
      TabIndex        =   4
      Top             =   4470
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -600
      TabIndex        =   3
      Top             =   4170
      Width           =   7770
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3030
      Top             =   2220
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
            Picture         =   "frmManSelect.frx":06EA
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSelect.frx":0D36
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSelect.frx":1052
            Key             =   "WriteNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSelect.frx":1372
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSelect.frx":190C
            Key             =   "Man"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2595
      Left            =   3120
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2595
      ScaleWidth      =   30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1620
      Width           =   30
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2955
      Left            =   3390
      TabIndex        =   1
      Top             =   570
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvwMain 
      Height          =   3345
      Left            =   330
      TabIndex        =   0
      Top             =   330
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   5900
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmManSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrReturn As String

Dim msngStartX As Single    '移动前鼠标的位置
Dim mblnItem As Boolean  '为真表示单击到ListView某一项上
Dim mintColumn As Integer
Dim mblnLoad As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not lvwMain.SelectedItem Is Nothing Then
        With lvwMain.SelectedItem
            mstrReturn = Mid(.Key, 2) & ";" & .SubItems(1) & ";" & .Text
        End With
    Else
        mstrReturn = ""
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        FillTree
        tvwMain_NodeClick tvwMain.Nodes("Root")
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    mblnLoad = True
    mstrReturn = ""
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    cmdCancel.Top = ScaleHeight - cmdCancel.Height - 200
    cmdOK.Top = cmdCancel.Top
    Frame1.Top = cmdOK.Top - 300
    
    
    cmdCancel.Left = ScaleWidth - cmdCancel.Width - 300
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    Frame1.Width = ScaleWidth + 800
    
    sngTop = Me.ScaleTop
    sngBottom = Frame1.Top
    
    
    tvwMain.Top = sngTop
    tvwMain.Height = IIf(sngBottom - tvwMain.Top > 0, sngBottom - tvwMain.Top, 0)
    tvwMain.Left = 0
    
    picSplit.Top = sngTop
    picSplit.Height = IIf(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = tvwMain.Left + tvwMain.Width
    
    lvwMain.Left = picSplit.Left + picSplit.Width
    lvwMain.Top = sngTop
    lvwMain.Height = IIf(sngBottom - lvwMain.Top > 0, sngBottom - lvwMain.Top, 0)
    If Me.ScaleWidth - lvwMain.Left > 0 Then lvwMain.Width = Me.ScaleWidth - lvwMain.Left
    Me.Refresh
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwMain.SortOrder = IIf(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub


Private Sub lvwMain_DblClick()
    If mblnItem = True Then cmdOK_Click
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub lvwMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub picsplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub picsplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplit.Left + X - msngStartX
        If sngTemp > 500 And Me.ScaleWidth - (sngTemp + picSplit.Width) > 500 Then
            picSplit.Left = sngTemp
            tvwMain.Width = picSplit.Left - tvwMain.Left
            lvwMain.Left = picSplit.Left + picSplit.Width
            lvwMain.Width = Me.ScaleWidth - lvwMain.Left
        End If
        tvwMain.SetFocus
    End If
End Sub

Private Sub tvwMain_NodeClick(ByVal Node As MSComctlLib.Node)
    FillList Node.Key
End Sub


Private Sub FillTree()
'功能:装入所有部门到tvwMain
'参数:
    Dim strTemp As String
    Dim rs部门 As New ADODB.Recordset
    
    rs部门.CursorLocation = adUseClient
    rs部门.CursorType = adOpenKeyset
    rs部门.LockType = adLockReadOnly
    
    strTemp = " where (撤档时间 = to_date('3000-01-01','YYYY-MM-DD') or 撤档时间 is null ) "
    rs部门.Open "select id,上级id,编码 ,名称,撤档时间  from 部门表 " & strTemp & " start with 上级id is null connect by prior id =上级id ", gcnOracle
    tvwMain.Nodes.Clear
    tvwMain.Nodes.Add , , "Root", "所有部门", "Root", "Root"
    tvwMain.Nodes("Root").Sorted = True
    Do Until rs部门.EOF
        If CDate(IIf(IsNull(rs部门("撤档时间")), CDate("3000/1/1"), rs部门("撤档时间"))) = CDate("3000/1/1") Then
            strTemp = "Write"
        Else
            strTemp = "WriteNo"
        End If
        If IsNull(rs部门("上级id")) Then
            tvwMain.Nodes.Add "Root", tvwChild, "C" & rs部门("id"), "【" & rs部门("编码") & "】" & rs部门("名称"), strTemp, strTemp
        Else
            tvwMain.Nodes.Add "C" & rs部门("上级id"), tvwChild, "C" & rs部门("id"), "【" & rs部门("编码") & "】" & rs部门("名称"), strTemp, strTemp
        End If
        tvwMain.Nodes("C" & rs部门("id")).Sorted = True
        rs部门.MoveNext
    Loop
    tvwMain.Nodes("Root").Selected = True
    tvwMain.Nodes("Root").Expanded = True
End Sub

Public Sub FillList(ByVal str部门ID As String)
'功能:装入对应部门的人员到lvwMain
'参数:str部门ID 部门的标识

    Dim rs人员 As New ADODB.Recordset
    Dim fld As Field, lst As ListItem
    Dim strSQL As String
    
    Dim strIcon As String
    Set rs人员 = New ADODB.Recordset
    rs人员.CursorLocation = adUseClient
    rs人员.CursorType = adOpenKeyset
    rs人员.LockType = adLockReadOnly
    
    If str部门ID = "Root" Then
        strSQL = "" & _
            "   Select distinct a.ID,C.部门ID,a.姓名,a.编号,a.简码,A.性别 ,b.名称 as 部门 " & _
            "   from 人员表 a,部门表 b,部门人员 C " & _
            "   where (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) " & _
            "        And  A.id=c.人员ID and C.部门id = b.id And nvl(C.缺省,0)=1"
        rs人员.Open strSQL, gcnOracle
    Else
        strSQL = "" & _
            "   Select A.ID,A.姓名,A.编号,A.简码,A.性别  " & _
            "   From 人员表 A,部门人员 C " & _
            "   Where (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null )  " & _
            "        and C.人员id=A.id   And C.部门ID = " & Mid(str部门ID, 2)
        
        rs人员.Open strSQL, gcnOracle
    End If
    '重新设置列头
    lvwMain.ColumnHeaders.Clear
    For Each fld In rs人员.Fields
       If InStr(fld.Name, "ID") = 0 Then lvwMain.ColumnHeaders.Add , fld.Name, fld.Name
    Next
    
    '取得数据
    lvwMain.ListItems.Clear
    Do Until rs人员.EOF
        strIcon = IIf(rs人员("性别") = "女", "Woman", "Man")
        Set lst = lvwMain.ListItems.Add(, "C" & rs人员("ID"), rs人员("姓名"), strIcon, strIcon)
        For Each fld In rs人员.Fields
            If InStr(fld.Name, "ID") = 0 And fld.Name <> "姓名" Then
                lst.SubItems(lvwMain.ColumnHeaders(fld.Name).Index - 1) = IIf(IsNull(fld.Value), "", fld.Value)
            End If
        Next
        If str部门ID = "Root" Then
            lst.ListSubItems(1).Tag = rs人员("部门ID")
        End If
        rs人员.MoveNext
    Loop
End Sub
