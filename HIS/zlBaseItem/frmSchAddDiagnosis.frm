VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSchAddDiagnosis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "新增诊疗项目"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5970
   Icon            =   "frmSchAddDiagnosis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ListBox lstAttentions 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "frmSchAddDiagnosis.frx":058A
      Left            =   3120
      List            =   "frmSchAddDiagnosis.frx":058C
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭"
      Height          =   350
      Left            =   4560
      TabIndex        =   7
      Top             =   5640
      Width           =   1100
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增"
      Height          =   350
      Left            =   3360
      TabIndex        =   6
      Top             =   5640
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "选择 …"
      Height          =   1095
      Left            =   5040
      TabIndex        =   5
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtNode 
      Height          =   1095
      Left            =   1080
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4320
      Width           =   3975
   End
   Begin VB.TextBox txtTime 
      Height          =   270
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "15"
      Top             =   3915
      Width           =   615
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2595
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgKind 
      Left            =   0
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchAddDiagnosis.frx":058E
            Key             =   "kind"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchAddDiagnosis.frx":0B28
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   720
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
            Picture         =   "frmSchAddDiagnosis.frx":10C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchAddDiagnosis.frx":165C
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchAddDiagnosis.frx":1BF6
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchAddDiagnosis.frx":1E10
            Key             =   "ClearAll"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   390
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   688
      ButtonWidth     =   1349
      ButtonHeight    =   582
      TextAlignment   =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "全选"
            Key             =   "全选"
            Object.ToolTipText     =   "选择所有显示项目"
            Object.Tag             =   "全选"
            ImageKey        =   "SelectAll"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "全清"
            Key             =   "全清"
            Object.ToolTipText     =   "清除所有选择标志"
            Object.Tag             =   "全清"
            ImageKey        =   "ClearAll"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "诊疗项目"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "注意事项"
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   4320
      Width           =   720
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "检查时长         分钟"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   3960
      Width           =   1890
   End
   Begin VB.Label lblHint 
      AutoSize        =   -1  'True
      Caption         =   "以下两个参数在新增时是批量设置，可以在新增完成后再逐个修改。"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   5400
   End
End
Attribute VB_Name = "frmSchAddDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrResult As String         '返回字符串，格式：项目1<->项目2<->项目3<->..<*>检查时长<*>注意事项
Private mstrItem As String           '已添加的诊疗项目
Private mstrCurItem As String        '已选中的诊疗项目
Private mstrOldItem As String
Private mstrModality As String       '影像类别

Public Event OnAddDiagnosis(ByVal strResult As String, ByVal strItem As String)

Private Sub cmdAdd_Click()
    Dim i As Long
    
    On Error GoTo errHandle
    
    If Val(txtTime.Text) = 0 Then
        MsgBox "检查时长需要大于0，请重新输入。", vbInformation, "检查预约提示"
        txtTime.SetFocus
        Exit Sub
    End If
    
    mstrResult = ""
    
    mstrResult = mstrCurItem
    
    If Len(mstrResult) > 0 Then
        mstrResult = mstrResult & "<*>" & Val(txtTime.Text)
        mstrResult = mstrResult & "<*>" & Trim(txtNode.Text)
    End If
    
    RaiseEvent OnAddDiagnosis(mstrResult, mstrItem)
    
    '去掉已经勾选的项目
    For i = 1 To lvwItem.ListItems.Count
        If lvwItem.ListItems(i).Checked = True Then
            mstrOldItem = mstrOldItem & "<" & Mid(lvwItem.ListItems(i).Key, 2) & ">"
        End If
    Next i
    mstrCurItem = ""
    
    Call RefreshData
    
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, "提示"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo errHandle
        lstAttentions.Visible = Not lstAttentions.Visible
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, "提示"
    err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    
    Call RefreshData
    
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, "提示"
    err.Clear
End Sub

Private Sub RefreshData()
    Dim strSql As String
    Dim rsTemp As Recordset
    Dim objItem As ListItem
    
    Me.lvwItem.ColumnHeaders.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_编码", "编码", 1400
        .Add , "_名称", "名称", 2200
        .Add , "_部位", "部位", 1800
    End With
    With Me.lvwItem
        .SortKey = .ColumnHeaders("_编码").Index - 1: .SortOrder = lvwAscending
    End With
    
    Me.lvwItem.ColumnHeaders("_编码").Position = 1
    
    Call zlRefItems
    On Error GoTo errH
    strSql = "select distinct 注意事项 from 影像预约项目"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "查询预约项目注意事项")
    
    Me.lstAttentions.Clear
    
    If rsTemp.RecordCount > 0 Then
        Do While Not rsTemp.EOF
            If Len(NVL(rsTemp!注意事项)) > 0 Then
                lstAttentions.AddItem NVL(rsTemp!注意事项)
            End If
            rsTemp.MoveNext
        Loop
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    lstAttentions.Move txtNode.Left, txtNode.Top - lstAttentions.Height - 10
End Sub

Private Sub lstAttentions_DblClick()
    On Error GoTo errHandle
    
    txtNode.Text = lstAttentions.List(lstAttentions.ListIndex)
    
    lstAttentions.Visible = False
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, "提示"
    err.Clear
End Sub

Private Sub lvwItem_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo errHandle
    
    If Item.Checked Then
        mstrCurItem = mstrCurItem & IIF(Len(mstrCurItem) > 0, "<->", "") & "<" & Item.Text & "|" & Replace(Item.Key, "_", "") & "|" & Item.Tag & ">"
        mstrItem = mstrItem & "<" & Replace(Item.Key, "_", "") & ">"
    Else
        mstrCurItem = Replace(mstrCurItem, "<" & Item.Text & "|" & Replace(Item.Key, "_", "") & "|" & Item.Tag & ">", "")
        mstrItem = Replace(mstrItem, "<" & Replace(Item.Key, "_", "") & ">", "")
    End If
    
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, "提示"
    err.Clear
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo errHandle
    
     Select Case Button.Key
        Case "全选"
            SelectAll True
        Case "全清"
            SelectAll False
     End Select
     
     Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, "提示"
    err.Clear
End Sub

Private Sub SelectAll(ByVal blnSelect As Boolean)
    Dim i As Integer
    With lvwItem
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = blnSelect
            If blnSelect Then
                mstrCurItem = mstrCurItem & IIF(Len(mstrCurItem) > 0, "<->", "") & "<" & .ListItems(i).Text & "|" & Replace(.ListItems(i).Key, "_", "") & "|" & .ListItems(i).Tag & ">"
                mstrItem = mstrItem & "<" & Replace(.ListItems(i).Key, "_", "") & ">"
            Else
                mstrCurItem = Replace(mstrCurItem, "<" & .ListItems(i).Text & "|" & Replace(.ListItems(i).Key, "_", "") & "|" & .ListItems(i).Tag & ">", "")
                mstrItem = Replace(mstrItem, "<" & Replace(.ListItems(i).Key, "_", "") & ">", "")
            End If
        Next
    End With
End Sub

Private Sub txtNode_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandle
    
    If lstAttentions.Visible Then
        lstAttentions.Visible = False
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, "提示"
    err.Clear
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub zlRefItems(Optional lngItemId As Long)
'-------------------------------------------------
'功能:刷新当前的项目列表
'-------------------------------------------------
    Dim strSql As String
    Dim rsTemp As Recordset
    Dim objItem As ListItem
    
    On Error GoTo errH
    strSql = "Select I.ID,I.编码, I.名称,I.标本部位,R.影像类别" & _
            "  From 诊疗项目目录 I, 影像检查项目 R" & _
            " Where I.ID = R.诊疗项目id And R.影像类别=[1] "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "刷新项目列表", mstrModality)
    
    
    With rsTemp
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            If InStr(mstrOldItem, "<" & !ID & ">") = 0 Then
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !编码)
                objItem.Tag = !影像类别
    '            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !名称, "item", "item")
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_名称").Index - 1) = !名称
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_部位").Index - 1) = IIF(IsNull(!标本部位), "", !标本部位)
                
                If InStr(mstrCurItem, "<" & !编码 & "|" & !ID & "|" & !影像类别 & ">") > 0 Then objItem.Checked = True
            End If
            .MoveNext
        Loop
    End With
    If Me.lvwItem.ListItems.Count > 0 Then
        If lngItemId > 0 Then
            Me.lvwItem.ListItems("_" & lngItemId).Selected = True
        End If
        If Me.lvwItem.SelectedItem Is Nothing Then Me.lvwItem.ListItems(1).Selected = True
        Me.lvwItem.SelectedItem.EnsureVisible
    Else
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Public Sub zlShowMe(ByVal strItem As String, ByVal strModality As String, ByVal ower As Object)
    mstrResult = ""
    mstrCurItem = ""
    mstrItem = strItem
    mstrOldItem = strItem
    mstrModality = strModality
    
    Me.Show 1, ower
End Sub
