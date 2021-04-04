VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm病种选择_重庆银海版 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病种选择"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   Icon            =   "frm病种选择_重庆银海版.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt查找 
      Height          =   300
      Left            =   915
      TabIndex        =   1
      Top             =   3900
      Width           =   1800
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   2145
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3210
      ScaleWidth      =   45
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   540
      Width           =   45
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   7350
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4260
      Width           =   7350
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4845
         TabIndex        =   7
         Top             =   105
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6075
         TabIndex        =   8
         Top             =   105
         Width           =   1100
      End
      Begin VB.TextBox txt并发症 
         Height          =   300
         Left            =   915
         TabIndex        =   6
         Top             =   150
         Width           =   3720
      End
      Begin VB.Label lbl并发症 
         AutoSize        =   -1  'True
         Caption         =   "并发症(&B)"
         Height          =   180
         Left            =   60
         TabIndex        =   5
         Top             =   210
         Width           =   810
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   7350
      TabIndex        =   9
      Top             =   0
      Width           =   7350
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择一个项目,然后点击确定"
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   120
         Width           =   2430
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3240
      Left            =   2205
      TabIndex        =   4
      Top             =   555
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   5715
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "编码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "名称"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "简码"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   3240
      Left            =   15
      TabIndex        =   3
      Top             =   540
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   5715
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   3615
      Top             =   765
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm病种选择_重庆银海版.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl查找 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "查找(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   3960
      Width           =   630
   End
End
Attribute VB_Name = "frm病种选择_重庆银海版"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint就诊类型 As Integer
Private mstr病种编码 As String
Private mstr并发症 As String
Private mblnOK As Boolean
'返回操作员选择的病种编码和并发症

Public Function ShowSelect(ByVal frmParent As Object, ByVal int就诊类型 As Integer, _
    str病种编码 As String, str并发症 As String) As Boolean
    mblnOK = False
    mint就诊类型 = int就诊类型
    mstr病种编码 = str病种编码
    mstr并发症 = str并发症
    Me.Show 1, frmParent
    
    str病种编码 = mstr病种编码
    str并发症 = mstr并发症
    ShowSelect = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If lvw.ListItems.Count = 0 Then Exit Sub
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    mstr病种编码 = lvw.SelectedItem.Text
    mstr并发症 = txt并发症.Text
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If lvw.Visible Then
        lvw.SetFocus
    Else
        tvw_s.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And cmdOK.Enabled Then cmdOK_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strTvw_Key As String, strLvw_Key As String, blnSelect As Boolean
    Dim rsTemp As New ADODB.Recordset
    '强制加入五大类
    If mint就诊类型 = 13 Then        '门诊特殊病
        Call tvw_s.Nodes.Add(, , "K1", "特殊病", 1, 1)
    ElseIf mint就诊类型 = 14 Then    '门诊急诊
        Call tvw_s.Nodes.Add(, , "K2", "急诊病", 1, 1)
    Else
        Call tvw_s.Nodes.Add(, , "K0", "普通病", 1, 1)
        Call tvw_s.Nodes.Add(, , "K3", "恶性肿瘤", 1, 1)
        Call tvw_s.Nodes.Add(, , "K4", "精神病", 1, 1)
    End If
    
    '如果病种编码不为空,提取该病种相关信息
    If mstr病种编码 <> "" Then
        gstrSQL = "Select 类别,ID From 保险病种 Where 险类=[1] And 编码=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病种信息,定位到历史病种选择", TYPE_重庆银海版, mstr病种编码)
        If Not rsTemp.EOF Then
            strTvw_Key = "K" & rsTemp!类别
            strLvw_Key = "K" & rsTemp!ID
            
            '定位
            On Error GoTo errHand
            tvw_s.Nodes(strTvw_Key).Selected = True
            tvw_s.SelectedItem.Selected = True
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
            If lvw.ListItems.Count <> 0 Then
                lvw.ListItems(strLvw_Key).Selected = True
                lvw.SelectedItem.Selected = True
                lvw.SelectedItem.EnsureVisible
            End If
        End If
    End If
            
errHand:
    If mstr病种编码 = "" Or blnSelect = False Then
        tvw_s.Nodes(1).Selected = True
        tvw_s.SelectedItem.Selected = True
        Call tvw_s_NodeClick(tvw_s.SelectedItem)
        If lvw.ListItems.Count <> 0 Then
            lvw.ListItems(1).Selected = True
            lvw.SelectedItem.Selected = True
            lvw.SelectedItem.EnsureVisible
        End If
    End If
    txt并发症.Text = mstr并发症
    
    If tvw_s.Nodes.Count = 1 Then tvw_s.Visible = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    If tvw_s.Visible Then
        tvw_s.Top = picInfo.Height
        tvw_s.Left = 0
        tvw_s.Width = pic.Left
        tvw_s.Height = Me.ScaleHeight - picInfo.Height - picCmd.Height - txt查找.Height - 100
        
        lvw.Top = tvw_s.Top
        lvw.Left = tvw_s.Width + pic.Width
        lvw.Width = Me.ScaleWidth - tvw_s.Width - pic.Width
        lvw.Height = tvw_s.Height
    Else
        pic.Visible = False
        lvw.Top = picInfo.Height
        lvw.Left = 0
        lvw.Width = Me.ScaleWidth
        lvw.Height = Me.ScaleHeight - picInfo.Height - picCmd.Height - txt查找.Height - 100
    End If
    
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.3
    cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me)
End Sub

Private Sub lvw_DblClick()
    If cmdOK.Enabled And Not lvw.SelectedItem Is Nothing Then cmdOK_Click
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tvw_s.Width + x < 1000 Or lvw.Width - x < 1000 Then Exit Sub
        pic.Left = pic.Left + x
        tvw_s.Width = tvw_s.Width + x
        lvw.Left = lvw.Left + x
        lvw.Width = lvw.Width - x
        Me.Refresh
    End If
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select ID,编码,名称,简码 From 保险病种 Where 险类=[1] And 类别=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取该大类下所有病种", TYPE_重庆银海版, CStr(Mid(tvw_s.SelectedItem.Key, 2)))
    
    lvw.ListItems.Clear
    
    With rsTemp
        Do While Not .EOF
            lvw.ListItems.Add , "K" & !ID, !编码, , 1
            lvw.ListItems("K" & !ID).SubItems(1) = !名称
            lvw.ListItems("K" & !ID).SubItems(2) = Nvl(!简码)
            .MoveNext
        Loop
        
        If lvw.ListItems.Count <> 0 Then
            lvw.ListItems(1).Selected = True
            lvw.SelectedItem.Selected = True
            lvw.SelectedItem.EnsureVisible
            Call zlControl.LvwSetColWidth(lvw)
        End If
    End With
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvw.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
    lvw.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub

Private Sub txt查找_Change()
    Dim lst As ListItem, lngIndex As Long, lngSubItems As Long
    Dim strFind As String
    
    strFind = UCase(Trim(txt查找.Text))
    If strFind = "" Then Exit Sub
    If lvw.ListItems.Count = 0 Then Exit Sub
    
    Set lst = lvw.FindItem(strFind, lvwText, , lvwPartial)
    If Not lst Is Nothing Then
        lst.Selected = True
        lst.EnsureVisible
    Else
        '非文本不能做到部分匹配
        lngSubItems = lvw.ColumnHeaders.Count - 1
        For Each lst In lvw.ListItems
            For lngIndex = 1 To lngSubItems
                If lst.SubItems(lngIndex) Like strFind & "*" Then
                    lst.Selected = True
                    lst.EnsureVisible
                    Exit Sub
                End If
            Next
        Next
    End If
End Sub
