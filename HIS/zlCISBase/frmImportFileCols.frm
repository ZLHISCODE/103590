VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmImportFileCols 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "列设置"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   Icon            =   "frmImportFileCols.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6165
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存(&O)"
      Height          =   350
      Left            =   3600
      TabIndex        =   1
      Top             =   4620
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&C)"
      Height          =   350
      Left            =   4875
      TabIndex        =   0
      Top             =   4620
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   105
      Top             =   4590
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
            Picture         =   "frmImportFileCols.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   405
      Left            =   960
      TabIndex        =   2
      Top             =   4605
      Width           =   1170
      _Version        =   589884
      _ExtentX        =   2064
      _ExtentY        =   714
      _StockProps     =   64
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4185
      Left            =   60
      ScaleHeight     =   4155
      ScaleWidth      =   5985
      TabIndex        =   3
      Top             =   375
      Width           =   6015
      Begin VB.CommandButton cmdOper 
         Caption         =   "↓"
         Height          =   345
         Index           =   3
         Left            =   2565
         TabIndex        =   7
         Top             =   2910
         Width           =   825
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "↑"
         Height          =   345
         Index           =   2
         Left            =   2565
         TabIndex        =   6
         Top             =   2400
         Width           =   825
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   ">>"
         Height          =   345
         Index           =   1
         Left            =   2565
         TabIndex        =   5
         Top             =   1455
         Width           =   825
      End
      Begin VB.CommandButton cmdOper 
         Caption         =   "<<"
         Height          =   345
         Index           =   0
         Left            =   2565
         TabIndex        =   4
         Top             =   960
         Width           =   825
      End
      Begin MSComctlLib.ListView lvwColumns_UnSelect 
         Height          =   3615
         Left            =   3540
         TabIndex        =   8
         Top             =   465
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   6376
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         SmallIcons      =   "ImageList1"
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
      Begin MSComctlLib.ListView lvwColumns_Select 
         Height          =   3630
         Left            =   45
         TabIndex        =   9
         Top             =   480
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   6403
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         SmallIcons      =   "ImageList1"
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
      Begin VB.Label lblNote 
         BackColor       =   &H00FFFFFF&
         Caption         =   "隐藏的列"
         Height          =   270
         Index           =   4
         Left            =   3540
         TabIndex        =   11
         Top             =   105
         Width           =   2715
      End
      Begin VB.Label lblNote 
         BackColor       =   &H00FFFFFF&
         Caption         =   "显示的列，蓝色表示固定显示的列，不能隐藏"
         Height          =   450
         Index           =   3
         Left            =   90
         TabIndex        =   10
         Top             =   75
         Width           =   2370
      End
   End
End
Attribute VB_Name = "frmImportFileCols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrMediColumns     As String  '
Private mstrColumn_Select   As String  '明细显示列
Private mstrColumn_UnSelect As String  '明细隐藏列
Private mstrColumn_Must     As String  '明细必显示列
Private mstrType_Select     As String  '分类显示列
Private mstrType_Unselect   As String  '分类隐藏列
Private mstrType_Must       As String  '分类必显示列
Private mlngKey             As Long
Private mblnOK              As Boolean
'列的移动方式
Enum OperState
    Add_ = 0
    Del_
    Up_
    Down_
End Enum

Public Function ShowMe(ByVal frmParent As Object, ByRef strMediColumns As String) As Boolean
    Dim arrMediColumn As Variant  '明细
    Dim arrTypeColumn As Variant  '分类
    Dim intCol        As Integer
    Dim intNum        As Integer
    
    '设置列格式
    'strMediColumns:A1,0,0|A2,1,1|A3,1,0|...    '0-必选;1-可选，0-显示;1-隐藏
    mblnOK = False
    mstrMediColumns = strMediColumns
    mstrColumn_Select = ""
    mstrColumn_UnSelect = ""
    mstrColumn_Must = ""
    mstrType_Select = ""
    mstrType_Unselect = ""
    mstrType_Must = ""
    
    If mstrMediColumns <> "" Then
        arrTypeColumn = Split(Split(mstrMediColumns, "||")(0) & "|", "|")
        arrMediColumn = Split(Split(mstrMediColumns, "||")(1), "|")
        '分类
        Do While arrTypeColumn(intNum) <> ""
            If Split(arrTypeColumn(intNum), ",")(2) = 0 Then
                '显示列
                mstrType_Select = mstrType_Select & arrTypeColumn(intNum) & "|"
                If Split(arrTypeColumn(intNum), ",")(1) = 0 Then
                    '必显示列
                    mstrType_Must = mstrType_Must & arrTypeColumn(intNum) & "|"
                End If
            Else
                '隐藏列
                mstrType_Unselect = mstrType_Unselect & arrTypeColumn(intNum) & "|"
            End If
            intNum = intNum + 1
        Loop
        '明细
        Do While arrMediColumn(intCol) <> ""
            If Split(arrMediColumn(intCol), ",")(2) = 0 Then
                '显示列
                mstrColumn_Select = mstrColumn_Select & arrMediColumn(intCol) & "|"
                If Split(arrMediColumn(intCol), ",")(1) = 0 Then
                    '必显示列
                    mstrColumn_Must = mstrColumn_Must & arrMediColumn(intCol) & "|"
                End If
            Else
                '隐藏列
                mstrColumn_UnSelect = mstrColumn_UnSelect & arrMediColumn(intCol) & "|"
            End If
            intCol = intCol + 1
        Loop
    End If
    
    Me.Show 1, frmParent
    strMediColumns = mstrType_Select & mstrType_Unselect & "|" & mstrColumn_Select & mstrColumn_UnSelect
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call SaveColumn
    mblnOK = True
    mlngKey = 0
    Unload Me
End Sub

Private Sub SaveColumn()
    '保存最新的列串
    Dim intCol    As Integer
    Dim intCols   As Integer
    Dim strReturn As String
    
    If TabControl.Selected.Caption = "分类" Then
        With lvwColumns_Select
            '组合成显示列串
            intCols = .ListItems.Count
            For intCol = 1 To intCols
                strReturn = strReturn & .ListItems(intCol).Text & "," & IIf(.ListItems(intCol).ForeColor = &H8000000D, 0, 1) & "," & 0 & "|"
            Next
            mstrType_Select = strReturn
        End With
        
        strReturn = ""
        With lvwColumns_UnSelect
            '组合成隐藏列串
            intCols = .ListItems.Count
            For intCol = 1 To intCols
                strReturn = strReturn & .ListItems(intCol).Text & "," & 1 & "," & 1 & "|"
            Next
            mstrType_Unselect = strReturn
        End With
    Else
        With lvwColumns_Select
            '组合成显示列串
            intCols = .ListItems.Count
            For intCol = 1 To intCols
                strReturn = strReturn & .ListItems(intCol).Text & "," & IIf(.ListItems(intCol).ForeColor = &H8000000D, 0, 1) & "," & 0 & "|"
            Next
            mstrColumn_Select = strReturn
        End With
        
        strReturn = ""
        With lvwColumns_UnSelect
            '组合成隐藏列串
            intCols = .ListItems.Count
            For intCol = 1 To intCols
                strReturn = strReturn & .ListItems(intCol).Text & "," & 1 & "," & 1 & "|"
            Next
            mstrColumn_UnSelect = strReturn
        End With
    End If
End Sub

Private Sub cmdOper_Click(Index As Integer)
    '移动列位置
    Dim intSelect As Integer
    Dim strText   As String
    
    If Val(cmdOper(Index).Tag) = 0 Then Exit Sub
    
    Select Case Index
    Case Add_
        If lvwColumns_UnSelect.SelectedItem Is Nothing Then Exit Sub
        intSelect = lvwColumns_UnSelect.SelectedItem.Index
        lvwColumns_Select.ListItems.Add , "K" & mlngKey, lvwColumns_UnSelect.SelectedItem.Text, , 1
        mlngKey = mlngKey + 1
        
        lvwColumns_UnSelect.ListItems.Remove lvwColumns_UnSelect.SelectedItem.Key
        If lvwColumns_UnSelect.ListItems.Count <> 0 Then
            If lvwColumns_UnSelect.ListItems.Count > intSelect Then
                lvwColumns_UnSelect.ListItems(intSelect).Selected = True
            Else
                lvwColumns_UnSelect.ListItems(lvwColumns_UnSelect.ListItems.Count).Selected = True
            End If
            lvwColumns_UnSelect.SelectedItem.Selected = True
            lvwColumns_UnSelect.SelectedItem.EnsureVisible
        End If
    Case Del_
        If lvwColumns_Select.SelectedItem Is Nothing Then Exit Sub
        If IsFixedCol(lvwColumns_Select.SelectedItem) Then Exit Sub
        intSelect = lvwColumns_Select.SelectedItem.Index
        lvwColumns_UnSelect.ListItems.Add , "K" & mlngKey, lvwColumns_Select.SelectedItem.Text, , 1
        mlngKey = mlngKey + 1
        
        lvwColumns_Select.ListItems.Remove lvwColumns_Select.SelectedItem.Key
        If lvwColumns_Select.ListItems.Count <> 0 Then
            If lvwColumns_Select.ListItems.Count > intSelect Then
                lvwColumns_Select.ListItems(intSelect).Selected = True
            Else
                lvwColumns_Select.ListItems(lvwColumns_Select.ListItems.Count).Selected = True
            End If
            lvwColumns_Select.SelectedItem.Selected = True
            lvwColumns_Select.SelectedItem.EnsureVisible
        End If
    Case Up_, Down_
        If lvwColumns_Select.SelectedItem Is Nothing Then Exit Sub
        strText = lvwColumns_Select.SelectedItem.Text
        intSelect = lvwColumns_Select.SelectedItem.Index
        If Index = Up_ Then
            If Index = 1 Then Exit Sub
            lvwColumns_Select.SelectedItem.Text = lvwColumns_Select.ListItems(intSelect - 1).Text
            lvwColumns_Select.ListItems(intSelect - 1).Text = strText
            lvwColumns_Select.ListItems(intSelect - 1).Selected = True
        Else
            If Index > lvwColumns_Select.ListItems.Count Then Exit Sub
            lvwColumns_Select.SelectedItem.Text = lvwColumns_Select.ListItems(intSelect + 1).Text
            lvwColumns_Select.ListItems(intSelect + 1).Text = strText
            lvwColumns_Select.ListItems(intSelect + 1).Selected = True
        End If
        lvwColumns_Select.SelectedItem.Selected = True
        lvwColumns_Select.SelectedItem.EnsureVisible
        Call lvwColumns_Select_ItemClick(lvwColumns_Select.SelectedItem)
    End Select
    Call SetOper
    Call SaveColumn
End Sub

Private Sub Form_Load()
    Call InitTabControl
    Call SetOper
End Sub

Private Sub SetOper()
    '属性判断
    Dim blnAdd As Boolean, blnDel As Boolean
    Dim blnUp  As Boolean, blnDown As Boolean
    Dim lngCol As Long, lngCols As Long
    
    blnAdd = (lvwColumns_UnSelect.ListItems.Count <> 0)
    blnDel = (lvwColumns_Select.ListItems.Count <> 0)
    blnUp = False: blnDown = False
    If Not lvwColumns_Select.SelectedItem Is Nothing Then
        If lvwColumns_Select.SelectedItem.Index > 1 Then blnUp = True
        If lvwColumns_Select.SelectedItem.Index < lvwColumns_Select.ListItems.Count Then blnDown = True
    End If
    
    cmdOper(Add_).Tag = IIf(blnAdd, 1, 0)
    cmdOper(Add_).Enabled = IIf(blnAdd, True, False)
    cmdOper(Del_).Tag = IIf(blnDel, 1, 0)
    cmdOper(Del_).Enabled = IIf(blnDel, True, False)
    cmdOper(Up_).Tag = IIf(blnUp, 1, 0)
    cmdOper(Up_).Enabled = IIf(blnUp, True, False)
    cmdOper(Down_).Tag = IIf(blnDown, 1, 0)
    cmdOper(Down_).Enabled = IIf(blnDown, True, False)
    
    lngCols = lvwColumns_Select.ListItems.Count
    For lngCol = 1 To lngCols
        If IsFixedCol(lvwColumns_Select.ListItems(lngCol), False) Then
            lvwColumns_Select.ListItems(lngCol).ForeColor = &H8000000D
        Else
            lvwColumns_Select.ListItems(lngCol).ForeColor = 0
        End If
    Next
End Sub

Private Function IsFixedCol(ByVal strColumn As String, Optional ByVal ShowMsg As Boolean = True) As Boolean
    '检查是否是固定列
    Dim arrColumn As Variant
    Dim intCol    As Integer
    Dim strMust   As String
    
    If TabControl.Selected.Caption = "分类" Then
        strMust = mstrType_Must
    Else
        strMust = mstrColumn_Must
    End If
    arrColumn = Split(strMust, "|")
    Do While arrColumn(intCol) <> ""
        If strColumn = Split(arrColumn(intCol), ",")(0) Then
            If Val(Split(arrColumn(intCol), ",")(1)) = 0 Then
                If ShowMsg Then MsgBox "这列是必须输入项，不允许隐藏！", vbInformation, gstrSysName
                IsFixedCol = True
                Exit Function
            End If
            IsFixedCol = False
            Exit Do
        End If
        intCol = intCol + 1
    Loop
End Function

Private Sub Form_Resize()
    TabControl.Move 0, 0, ScaleWidth, ScaleHeight - 500
End Sub

Private Sub lvwColumns_Select_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call SetOper
End Sub

Private Function InitTabControl()
    '初始化分页控件
    With TabControl
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPageSelected
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
        End With
        
        .InsertItem 1, "分类", pic.hWnd, 101
        .InsertItem 2, "明细", pic.hWnd, 102
        .Item(0).Selected = True
    End With
End Function

Private Sub TabControl_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.Index
        Case 0
            Call AddColumn(mstrType_Select, mstrType_Unselect)
        Case 1
            Call AddColumn(mstrColumn_Select, mstrColumn_UnSelect)
    End Select
End Sub

Private Sub AddColumn(ByVal strSelect As String, ByVal strUnselect As String)
    '添加数据到列表
    Dim arrColumn As Variant
    Dim intCol    As Integer
    Dim intCols   As Integer
    
    lvwColumns_Select.ListItems.Clear
    lvwColumns_UnSelect.ListItems.Clear
    
    '显示的列
    If strSelect <> "" Then
    arrColumn = Split(strSelect, "|")
    Do While arrColumn(intCol) <> ""
        lvwColumns_Select.ListItems.Add , "K" & mlngKey, Split(arrColumn(intCol), ",")(0), , 1
        If Split(arrColumn(intCol), ",")(1) = 0 Then
            lvwColumns_Select.ListItems("K" & mlngKey).ForeColor = &H8000000D
        End If
        mlngKey = mlngKey + 1
        intCol = intCol + 1
    Loop
    lvwColumns_Select.ListItems(1).Selected = True
    lvwColumns_Select.SelectedItem.Selected = True
    lvwColumns_Select.SelectedItem.EnsureVisible
    End If
    
    '隐藏的列
    If strUnselect <> "" Then
        arrColumn = Split(strUnselect, "|")
        Do While arrColumn(intCols) <> ""
            lvwColumns_UnSelect.ListItems.Add , "K" & mlngKey, Split(arrColumn(intCols), ",")(0), , 1
            mlngKey = mlngKey + 1
            intCols = intCols + 1
        Loop
        lvwColumns_UnSelect.ListItems(1).Selected = True
        lvwColumns_UnSelect.SelectedItem.Selected = True
        lvwColumns_UnSelect.SelectedItem.EnsureVisible
    End If
    
    Call SetOper
End Sub
