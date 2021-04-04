VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "列设置"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmColSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6975
      TabIndex        =   13
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6975
      TabIndex        =   14
      Top             =   660
      Width           =   1100
   End
   Begin VB.Frame fra 
      Caption         =   "列选择"
      Height          =   5520
      Left            =   75
      TabIndex        =   15
      Top             =   75
      Width           =   6660
      Begin VB.CommandButton cmdOper 
         Caption         =   ">>"
         Height          =   345
         Index           =   3
         Left            =   3135
         TabIndex        =   5
         Top             =   1905
         Width           =   390
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   ">"
         Height          =   345
         Index           =   2
         Left            =   3135
         TabIndex        =   4
         Top             =   1560
         Width           =   390
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "<<"
         Height          =   345
         Index           =   1
         Left            =   3135
         TabIndex        =   3
         Top             =   1170
         Width           =   390
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "<"
         Height          =   345
         Index           =   0
         Left            =   3135
         TabIndex        =   2
         Top             =   810
         Width           =   390
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "↑"
         Height          =   345
         Index           =   4
         Left            =   3135
         TabIndex        =   6
         Top             =   2535
         Width           =   390
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "↓"
         Height          =   345
         Index           =   5
         Left            =   3135
         TabIndex        =   7
         Top             =   2880
         Width           =   390
      End
      Begin MSComctlLib.ListView lvwSelCol 
         Height          =   3930
         Left            =   90
         TabIndex        =   1
         Top             =   525
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   6932
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         Icons           =   "ilt"
         SmallIcons      =   "ilt"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "列名"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView lvwUnSelCol 
         Height          =   3930
         Left            =   3615
         TabIndex        =   9
         Top             =   525
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   6932
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         Icons           =   "ilt"
         SmallIcons      =   "ilt"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "列名"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Image img 
         Height          =   480
         Left            =   105
         Picture         =   "frmColSet.frx":000C
         Top             =   4605
         Width           =   480
      End
      Begin VB.Label lblNote 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "以下是您选择的列清单"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   0
         Top             =   255
         Width           =   2715
      End
      Begin VB.Label lblNote 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "以下是未选择的列清单"
         Height          =   195
         Index           =   3
         Left            =   3735
         TabIndex        =   8
         Top             =   285
         Width           =   2535
      End
      Begin VB.Label lblNote 
         Caption         =   "    你除了可以显示或隐藏某些非必须列外，还可以根据需要调整列的顺序"
         Height          =   225
         Index           =   2
         Left            =   315
         TabIndex        =   12
         Top             =   5085
         Width           =   6285
      End
      Begin VB.Label lblNote 
         Caption         =   "    黑色项目为可选项目，显示的是一些辅助信息，可以根据需要显示或隐藏"
         Height          =   225
         Index           =   1
         Left            =   315
         TabIndex        =   11
         Top             =   4815
         Width           =   6285
      End
      Begin VB.Label lblNote 
         Caption         =   "    蓝色项目为必须输入项目，不允许隐藏"
         Height          =   225
         Index           =   0
         Left            =   315
         TabIndex        =   10
         Top             =   4545
         Width           =   4935
      End
   End
   Begin MSComctlLib.ImageList ilt 
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
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColSet.frx":08D6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmColSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mstrColumns As String
Private mstrSelColumn As String
Private mstrUnSelColumn As String
Private mlngKey As Long

Enum OperState
    SelItem_
    SelItems_
    ReMove_
    ReMoves_
    Up_
    Down_
End Enum

Public Function ShowMe(ByVal FrmParent As Object, ByVal strColumns As String, strSelColumn As String) As Boolean
    '--------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示列设置界面
    '入参:frmMain-父窗口
    '     strColumns-所有列数,格式为(C1,0|C2,1|....)   0-代表必选项,1-代表可选项
    '     strSelColumn-初始选择的列,如果为空,表示选择了所有列.格式为:C1|C2...
    '出参:strSelColumn-初选择的列。格式为:C1|C2...\
    '返回:设置成功,返回true,否则返回false
    '--------------------------------------------------------------------------------------------------------------------------------------------------

    Dim arrColumn
    Dim intCol As Integer, intCols As Integer
    mblnOk = False
    mstrColumns = strColumns
    mstrSelColumn = strSelColumn
    
    If mstrSelColumn = "" Then
        '如果没有已选择列,则将所有列视为已选择列,转换为已选择列所需的格式
        arrColumn = Split(mstrColumns, "|")
        intCols = UBound(arrColumn)
        For intCol = 0 To intCols
            mstrSelColumn = mstrSelColumn & "|" & Split(arrColumn(intCol), ",")(0)
        Next
        mstrSelColumn = Mid(mstrSelColumn, 2)
    End If
    
    Me.Show 1, FrmParent
    strSelColumn = mstrSelColumn & "||" & mstrUnSelColumn
    ShowMe = mblnOk
End Function

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim intCol As Integer, intCols As Integer
    Dim strReturn As String
    With lvwSelCol
        '组合成选择列
        intCols = .ListItems.Count
        For intCol = 1 To intCols
            strReturn = strReturn & "|" & .ListItems(intCol).Text
        Next
        strReturn = Mid(strReturn, 2)
    End With
    mstrSelColumn = strReturn
    strReturn = ""
    With lvwUnSelCol
        '组合成未选择列串
        intCols = .ListItems.Count
        For intCol = 1 To intCols
            strReturn = strReturn & "|" & .ListItems(intCol).Text
        Next
        strReturn = Mid(strReturn, 2)
    End With
    mstrUnSelColumn = strReturn
    mblnOk = True
    
    Unload Me
End Sub

Private Sub cmdOper_Click(Index As Integer)
    Dim intSelect As Integer
    Dim strText As String
    Dim lvwItem As ListItem
    If Val(cmdOper(Index).Tag) = 0 Then Exit Sub
    
    Select Case Index
    Case SelItem_
        '选择当前未选部分的某一项内容
        If lvwUnSelCol.SelectedItem Is Nothing Then Exit Sub
        
        intSelect = lvwUnSelCol.SelectedItem.Index
        lvwSelCol.ListItems.Add , lvwUnSelCol.SelectedItem.Key, lvwUnSelCol.SelectedItem.Text, , 1
        mlngKey = mlngKey + 1
        
        lvwUnSelCol.ListItems.Remove lvwUnSelCol.SelectedItem.Key
        If lvwUnSelCol.ListItems.Count <> 0 Then
            If lvwUnSelCol.ListItems.Count > intSelect Then
                lvwUnSelCol.ListItems(intSelect).Selected = True
            Else
                lvwUnSelCol.ListItems(lvwUnSelCol.ListItems.Count).Selected = True
            End If
            lvwUnSelCol.SelectedItem.Selected = True
            lvwUnSelCol.SelectedItem.EnsureVisible
        End If
    Case SelItems_
        '选择所有未选部分
        For Each lvwItem In lvwUnSelCol.ListItems
            lvwSelCol.ListItems.Add , lvwItem.Key, lvwItem.Text, , 1
        Next
        If lvwSelCol.SelectedItem Is Nothing And lvwSelCol.ListItems.Count <> 0 Then
            lvwSelCol.ListItems(0).Selected = True
             lvwUnSelCol.SelectedItem.EnsureVisible
        End If
        lvwUnSelCol.ListItems.Clear
    Case ReMove_
        '移出已经选择列的当前项
        If lvwSelCol.SelectedItem Is Nothing Then Exit Sub
        If IsFixedCol(lvwSelCol.SelectedItem) Then Exit Sub
        
        intSelect = lvwSelCol.SelectedItem.Index
        lvwUnSelCol.ListItems.Add , "K" & mlngKey, lvwSelCol.SelectedItem.Text, , 1
        mlngKey = mlngKey + 1
        
        lvwSelCol.ListItems.Remove lvwSelCol.SelectedItem.Key
        If lvwSelCol.ListItems.Count <> 0 Then
            If lvwSelCol.ListItems.Count > intSelect Then
                lvwSelCol.ListItems(intSelect).Selected = True
            Else
                lvwSelCol.ListItems(lvwSelCol.ListItems.Count).Selected = True
            End If
            lvwSelCol.SelectedItem.Selected = True
            lvwSelCol.SelectedItem.EnsureVisible
        End If
    Case ReMoves_
        '选择所有已选部分,但不包含固定数据部分
        Dim cllKey As New Collection
        Dim i As Integer
        For Each lvwItem In lvwSelCol.ListItems
            If IsFixedCol(lvwItem, False) Then
            Else
                lvwUnSelCol.ListItems.Add , lvwItem.Key, lvwItem.Text, , 1
                cllKey.Add lvwItem.Key, lvwItem.Key
            End If
        Next
        For i = 1 To cllKey.Count
                lvwSelCol.ListItems.Remove cllKey(i)
        Next
        If lvwSelCol.SelectedItem Is Nothing And lvwSelCol.ListItems.Count <> 0 Then
            lvwSelCol.ListItems(0).Selected = True
             lvwUnSelCol.SelectedItem.EnsureVisible
        End If
    Case Up_, Down_
        Dim strTag As String
        If lvwSelCol.SelectedItem Is Nothing Then Exit Sub
        
        strText = lvwSelCol.SelectedItem.Text
        intSelect = lvwSelCol.SelectedItem.Index
        strTag = lvwSelCol.SelectedItem.Tag
        If Index = Up_ Then
            If intSelect - 1 < 1 Then Exit Sub
            lvwSelCol.SelectedItem.Text = lvwSelCol.ListItems(intSelect - 1).Text
            lvwSelCol.SelectedItem.Tag = lvwSelCol.ListItems(intSelect - 1).Tag
            lvwSelCol.ListItems(intSelect - 1).Text = strText
            lvwSelCol.ListItems(intSelect - 1).Selected = True
            lvwSelCol.ListItems(intSelect - 1).Tag = strTag
            
        Else
            If intSelect + 1 > lvwSelCol.ListItems.Count Then Exit Sub
            lvwSelCol.SelectedItem.Text = lvwSelCol.ListItems(intSelect + 1).Text
            lvwSelCol.SelectedItem.Tag = lvwSelCol.ListItems(intSelect + 1).Tag
            lvwSelCol.ListItems(intSelect + 1).Text = strText
            lvwSelCol.ListItems(intSelect + 1).Selected = True
            lvwSelCol.ListItems(intSelect + 1).Tag = strTag
        End If
        lvwSelCol.SelectedItem.Selected = True
        lvwSelCol.SelectedItem.EnsureVisible
        Call lvwSelCol_ItemClick(lvwSelCol.SelectedItem)
    End Select
    Call SetOper
End Sub

Private Sub Form_Load()
    Dim arrColumn
    Dim strSelected As String
    Dim intCol As Integer, intCols As Integer
    '显示已选择的列
    arrColumn = Split(mstrSelColumn, "|")
    intCols = UBound(arrColumn)
    For intCol = 0 To intCols
        lvwSelCol.ListItems.Add , "K" & mlngKey, arrColumn(intCol), 1, 1
        If IsFixedCol(arrColumn(intCol), False) Then
            lvwSelCol.ListItems("K" & mlngKey).ForeColor = &H8000000D
        End If
        mlngKey = mlngKey + 1
    Next
    lvwSelCol.ListItems(1).Selected = True
    lvwSelCol.SelectedItem.Selected = True
    lvwSelCol.SelectedItem.EnsureVisible
    
    '显示所有列(将未选择的列显示出来)
    strSelected = "|" & mstrSelColumn & "|"
    arrColumn = Split(mstrColumns, "|")
    intCols = UBound(arrColumn)
    For intCol = 0 To intCols
        If InStr(1, strSelected, "|" & Split(arrColumn(intCol), ",")(0) & "|") = 0 Then
            lvwUnSelCol.ListItems.Add , "K" & mlngKey, Split(arrColumn(intCol), ",")(0), , 1
            mlngKey = mlngKey + 1
        End If
    Next
    If lvwUnSelCol.ListItems.Count <> 0 Then
        lvwUnSelCol.ListItems(1).Selected = True
        lvwUnSelCol.SelectedItem.Selected = True
        lvwUnSelCol.SelectedItem.EnsureVisible
    End If
    
    Call SetOper
End Sub

Private Sub SetOper()
     Dim blnAdd As Boolean, blnDel As Boolean
    Dim blnSelData As Boolean, blnUnSelData As Boolean
    Dim blnUp As Boolean, blnDown As Boolean
    Dim lngCol As Long, lngCols As Long
    blnUp = False: blnDown = False
    
    blnSelData = (lvwUnSelCol.ListItems.Count <> 0)
    blnUnSelData = (lvwSelCol.ListItems.Count <> 0)
    blnAdd = Not lvwUnSelCol.SelectedItem Is Nothing
    blnDel = Not lvwSelCol.SelectedItem Is Nothing
    
    If blnDel Then
        If lvwSelCol.SelectedItem.Index > 1 Then blnUp = True
        If lvwSelCol.SelectedItem.Index < lvwSelCol.ListItems.Count Then blnDown = True
    End If
    
    cmdOper(SelItem_).Tag = IIf(blnAdd, 1, 0)
    cmdOper(SelItem_).Enabled = blnAdd
    cmdOper(SelItems_).Tag = IIf(blnSelData, 1, 0)
    cmdOper(SelItems_).Enabled = blnSelData
    
    cmdOper(ReMove_).Tag = IIf(blnDel, 1, 0)
    cmdOper(ReMove_).Enabled = IIf(blnDel, True, False)
    cmdOper(ReMoves_).Tag = IIf(blnUnSelData, 1, 0)
    cmdOper(ReMoves_).Enabled = blnUnSelData
    
    cmdOper(Up_).Tag = IIf(blnUp, 1, 0)
    cmdOper(Up_).Enabled = blnUp
    cmdOper(Down_).Tag = IIf(blnDown, 1, 0)
    cmdOper(Down_).Enabled = blnDown
    
    lngCols = lvwSelCol.ListItems.Count
    For lngCol = 1 To lngCols
        If IsFixedCol(lvwSelCol.ListItems(lngCol), False) Then
            lvwSelCol.ListItems(lngCol).ForeColor = &H8000000D
        Else
            lvwSelCol.ListItems(lngCol).ForeColor = 0
        End If
    Next
End Sub

Private Function IsFixedCol(ByVal strColumn As String, Optional ByVal ShowMsg As Boolean = True) As Boolean
    Dim arrColumn
    Dim intCol As Integer, intCols As Integer
    '检查是否是固定列,也就是必须选择的列
    IsFixedCol = True
    arrColumn = Split(mstrColumns, "|")
    intCols = UBound(arrColumn)
    For intCol = 0 To intCols
        If strColumn = Split(arrColumn(intCol), ",")(0) Then
            If Val(Split(arrColumn(intCol), ",")(1)) = 0 Then
                If ShowMsg Then MsgBox "这列是必须输入项，不允许隐藏！", vbInformation, gstrSysName
                Exit Function
            End If
            IsFixedCol = False
            Exit For
        End If
    Next
End Function

Private Sub lvwSelCol_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call SetOper
End Sub


Private Sub lvwSelCol_DblClick()
    If lvwSelCol.SelectedItem Is Nothing Then Exit Sub
    Call cmdOper_Click(ReMove_)
End Sub

Private Sub lvwUnSelCol_DblClick()
    If lvwUnSelCol.SelectedItem Is Nothing Then Exit Sub
    Call cmdOper_Click(SelItem_)
End Sub

