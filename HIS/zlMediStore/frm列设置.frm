VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm列设置 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "列设置"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   Icon            =   "frm列设置.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOper 
      Caption         =   "下移"
      Height          =   345
      Index           =   3
      Left            =   2730
      TabIndex        =   7
      Top             =   3450
      Width           =   825
   End
   Begin VB.CommandButton cmdOper 
      Caption         =   "上移"
      Height          =   345
      Index           =   2
      Left            =   2730
      TabIndex        =   6
      Top             =   2910
      Width           =   825
   End
   Begin VB.CommandButton cmdOper 
      Caption         =   "移出"
      Height          =   345
      Index           =   1
      Left            =   2730
      TabIndex        =   5
      Top             =   2190
      Width           =   825
   End
   Begin VB.CommandButton cmdOper 
      Caption         =   "移入"
      Height          =   345
      Index           =   0
      Left            =   2730
      TabIndex        =   4
      Top             =   1680
      Width           =   825
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   1230
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
            Picture         =   "frm列设置.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6570
      TabIndex        =   3
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6570
      TabIndex        =   2
      Top             =   1260
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwColumns_All 
      Height          =   4035
      Left            =   30
      TabIndex        =   0
      Top             =   1080
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   7117
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
   Begin MSComctlLib.ListView lvwColumns_Selected 
      Height          =   4035
      Left            =   3660
      TabIndex        =   1
      Top             =   1080
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   7117
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
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "以下是您选择的列清单"
      Height          =   195
      Index           =   4
      Left            =   3690
      TabIndex        =   12
      Top             =   900
      Width           =   2715
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "以下是未选择的列清单"
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   11
      Top             =   900
      Width           =   2535
   End
   Begin VB.Label lblNote 
      Caption         =   "    你除了可以显示或隐藏某些非必须列外，还可以根据需要调整列的顺序"
      Height          =   225
      Index           =   2
      Left            =   1020
      TabIndex        =   10
      Top             =   600
      Width           =   6285
   End
   Begin VB.Label lblNote 
      Caption         =   "    黑色项目为可选项目，显示的是一些辅助信息，可以根据需要显示或隐藏"
      Height          =   225
      Index           =   1
      Left            =   1020
      TabIndex        =   9
      Top             =   330
      Width           =   6285
   End
   Begin VB.Label lblNote 
      Caption         =   "    蓝色项目为必须输入项目，不允许隐藏"
      Height          =   225
      Index           =   0
      Left            =   1020
      TabIndex        =   8
      Top             =   60
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   420
      Picture         =   "frm列设置.frx":1D16
      Top             =   210
      Width           =   480
   End
End
Attribute VB_Name = "frm列设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mstrColumns As String
Private mstrColumn_Select As String
Private mstrColumn_UnSelect As String
Private mlngKey As Long

Enum OperState
    Add_
    Del_
    Up_
    Down_
End Enum

Public Function ShowME(ByVal frmParent As Object, ByVal strColumns As String, strColumn_Select As String) As Boolean
    Dim arrColumn
    Dim intCol As Integer, intCols As Integer
    '设置列格式
    'strColumns:A1,0|A2,1|A3,1      '0-必选;1-可选
    'strColumn_Select:A1|A2         '已选择的列,因为不可能为空,如果为空,表示选择了所有列
    mblnOK = False
    mstrColumns = strColumns
    mstrColumn_Select = strColumn_Select
    
    If mstrColumn_Select = "" Then
        '如果没有已选择列,则将所有列视为已选择列,转换为已选择列所需的格式
        arrColumn = Split(mstrColumns, "|")
        intCols = UBound(arrColumn)
        For intCol = 0 To intCols
            mstrColumn_Select = mstrColumn_Select & "|" & Split(arrColumn(intCol), ",")(0)
        Next
        mstrColumn_Select = Mid(mstrColumn_Select, 2)
    End If
    
    Me.Show 1, frmParent
    strColumn_Select = mstrColumn_Select & "||" & mstrColumn_UnSelect
    ShowME = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intCol As Integer, intCols As Integer
    Dim strReturn As String
    With lvwColumns_Selected
        '组合成选择列
        intCols = .ListItems.count
        For intCol = 1 To intCols
            strReturn = strReturn & "|" & .ListItems(intCol).Text
        Next
        strReturn = Mid(strReturn, 2)
    End With
    mstrColumn_Select = strReturn
    strReturn = ""
    With lvwColumns_All
        '组合成未选择列串
        intCols = .ListItems.count
        For intCol = 1 To intCols
            strReturn = strReturn & "|" & .ListItems(intCol).Text
        Next
        strReturn = Mid(strReturn, 2)
    End With
    mstrColumn_UnSelect = strReturn
    mblnOK = True
    
    Unload Me
End Sub

Private Sub cmdOper_Click(Index As Integer)
    Dim intSelect As Integer
    Dim strText As String
    If Val(cmdOper(Index).Tag) = 0 Then Exit Sub
    
    Select Case Index
    Case Add_
        If lvwColumns_All.SelectedItem Is Nothing Then Exit Sub
        intSelect = lvwColumns_All.SelectedItem.Index
        lvwColumns_Selected.ListItems.Add , "K" & mlngKey, lvwColumns_All.SelectedItem.Text, , 1
        mlngKey = mlngKey + 1
        
        lvwColumns_All.ListItems.Remove lvwColumns_All.SelectedItem.Key
        If lvwColumns_All.ListItems.count <> 0 Then
            If lvwColumns_All.ListItems.count > intSelect Then
                lvwColumns_All.ListItems(intSelect).Selected = True
            Else
                lvwColumns_All.ListItems(lvwColumns_All.ListItems.count).Selected = True
            End If
            lvwColumns_All.SelectedItem.Selected = True
            lvwColumns_All.SelectedItem.EnsureVisible
        End If
    Case Del_
        If lvwColumns_Selected.SelectedItem Is Nothing Then Exit Sub
        If IsFixedCol(lvwColumns_Selected.SelectedItem) Then Exit Sub
        intSelect = lvwColumns_Selected.SelectedItem.Index
        lvwColumns_All.ListItems.Add , "K" & mlngKey, lvwColumns_Selected.SelectedItem.Text, , 1
        mlngKey = mlngKey + 1
        
        lvwColumns_Selected.ListItems.Remove lvwColumns_Selected.SelectedItem.Key
        If lvwColumns_Selected.ListItems.count <> 0 Then
            If lvwColumns_Selected.ListItems.count > intSelect Then
                lvwColumns_Selected.ListItems(intSelect).Selected = True
            Else
                lvwColumns_Selected.ListItems(lvwColumns_Selected.ListItems.count).Selected = True
            End If
            lvwColumns_Selected.SelectedItem.Selected = True
            lvwColumns_Selected.SelectedItem.EnsureVisible
        End If
    Case Up_, Down_
        If lvwColumns_Selected.SelectedItem Is Nothing Then Exit Sub
        strText = lvwColumns_Selected.SelectedItem.Text
        intSelect = lvwColumns_Selected.SelectedItem.Index
        If Index = Up_ Then
            If Index = 1 Then Exit Sub
            lvwColumns_Selected.SelectedItem.Text = lvwColumns_Selected.ListItems(intSelect - 1).Text
            lvwColumns_Selected.ListItems(intSelect - 1).Text = strText
            lvwColumns_Selected.ListItems(intSelect - 1).Selected = True
        Else
            If Index > lvwColumns_Selected.ListItems.count Then Exit Sub
            lvwColumns_Selected.SelectedItem.Text = lvwColumns_Selected.ListItems(intSelect + 1).Text
            lvwColumns_Selected.ListItems(intSelect + 1).Text = strText
            lvwColumns_Selected.ListItems(intSelect + 1).Selected = True
        End If
        lvwColumns_Selected.SelectedItem.Selected = True
        lvwColumns_Selected.SelectedItem.EnsureVisible
        Call lvwColumns_Selected_ItemClick(lvwColumns_Selected.SelectedItem)
    End Select
    Call SetOper
End Sub

Private Sub Form_Load()
    Dim arrColumn
    Dim strSelected As String
    Dim intCol As Integer, intCols As Integer
    '显示已选择的列
    arrColumn = Split(mstrColumn_Select, "|")
    intCols = UBound(arrColumn)
    For intCol = 0 To intCols
        lvwColumns_Selected.ListItems.Add , "K" & mlngKey, arrColumn(intCol), , 1
        If IsFixedCol(arrColumn(intCol), False) Then
            lvwColumns_Selected.ListItems("K" & mlngKey).ForeColor = &H8000000D
        End If
        mlngKey = mlngKey + 1
    Next
    lvwColumns_Selected.ListItems(1).Selected = True
    lvwColumns_Selected.SelectedItem.Selected = True
    lvwColumns_Selected.SelectedItem.EnsureVisible
    
    '显示所有列(将未选择的列显示出来)
    strSelected = "|" & mstrColumn_Select & "|"
    arrColumn = Split(mstrColumns, "|")
    intCols = UBound(arrColumn)
    For intCol = 0 To intCols
        If InStr(1, strSelected, "|" & Split(arrColumn(intCol), ",")(0) & "|") = 0 Then
            lvwColumns_All.ListItems.Add , "K" & mlngKey, Split(arrColumn(intCol), ",")(0), , 1
            mlngKey = mlngKey + 1
        End If
    Next
    If lvwColumns_All.ListItems.count <> 0 Then
        lvwColumns_All.ListItems(1).Selected = True
        lvwColumns_All.SelectedItem.Selected = True
        lvwColumns_All.SelectedItem.EnsureVisible
    End If
    
    Call SetOper
End Sub

Private Sub SetOper()
    Dim blnAdd As Boolean, blnDel As Boolean
    Dim blnUp As Boolean, blnDown As Boolean
    Dim lngCol As Long, lngCols As Long
    blnAdd = (lvwColumns_All.ListItems.count <> 0)
    blnDel = (lvwColumns_Selected.ListItems.count <> 0)
    blnUp = False: blnDown = False
    If Not lvwColumns_Selected.SelectedItem Is Nothing Then
        If lvwColumns_Selected.SelectedItem.Index > 1 Then blnUp = True
        If lvwColumns_Selected.SelectedItem.Index < lvwColumns_Selected.ListItems.count Then blnDown = True
    End If
    
    cmdOper(Add_).Tag = IIf(blnAdd, 1, 0)
    cmdOper(Add_).Enabled = IIf(blnAdd, True, False)
    cmdOper(Del_).Tag = IIf(blnDel, 1, 0)
    cmdOper(Del_).Enabled = IIf(blnDel, True, False)
    cmdOper(Up_).Tag = IIf(blnUp, 1, 0)
    cmdOper(Up_).Enabled = IIf(blnUp, True, False)
    cmdOper(Down_).Tag = IIf(blnDown, 1, 0)
    cmdOper(Down_).Enabled = IIf(blnDown, True, False)
    
    lngCols = lvwColumns_Selected.ListItems.count
    For lngCol = 1 To lngCols
        If IsFixedCol(lvwColumns_Selected.ListItems(lngCol), False) Then
            lvwColumns_Selected.ListItems(lngCol).ForeColor = &H8000000D
        Else
            lvwColumns_Selected.ListItems(lngCol).ForeColor = 0
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

Private Sub lvwColumns_Selected_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call SetOper
End Sub
