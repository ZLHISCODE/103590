VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmSelCurr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmSelCurr.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra1 
      Height          =   2745
      Left            =   5880
      TabIndex        =   3
      Top             =   1545
      Width           =   1260
      Begin VB.OptionButton Opt 
         Caption         =   "完全匹配"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   2115
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton Opt 
         Caption         =   "模式匹配"
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   2400
         Width           =   1035
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找(&F)"
         Height          =   350
         Left            =   120
         TabIndex        =   8
         Top             =   1605
         Width           =   1035
      End
      Begin VB.TextBox txtValues 
         Height          =   270
         Left            =   105
         TabIndex        =   7
         Top             =   1185
         Width           =   1035
      End
      Begin VB.TextBox txtCol 
         Height          =   270
         Left            =   105
         TabIndex        =   5
         Top             =   555
         Width           =   1035
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查找内容"
         Height          =   180
         Left            =   150
         TabIndex        =   6
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lblCol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查找列名"
         Height          =   180
         Left            =   150
         TabIndex        =   4
         Top             =   315
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5850
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   810
      Width           =   1260
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5850
      TabIndex        =   1
      Top             =   255
      Width           =   1260
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   6600
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   11642
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelCurr.frx":08CA
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelCurr.frx":0D1C
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelCurr.frx":1036
            Key             =   "ry"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelCurr.frx":15D0
            Key             =   "dq"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelCurr.frx":1B6A
            Key             =   "bm"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelCurr.frx":2104
            Key             =   "item"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelCurr.frx":225E
            Key             =   "book"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelCurr.frx":23B8
            Key             =   "bookopen"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSelCurr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintReturnCol As Integer    '返回哪一列
Private mstrReturn As String    '返回字符串
Private mstrColHead As String
Private mRs As ADODB.Recordset
Private mblnSorted As Boolean
Private mstrDefItem As String
Private mlngDefIndex As Long
Private mIconIndex As Long
Private mlngWidth As Long
Private mlngHeight As Long

Public Function ShowCurrSel(frmParent As Object, rs As ADODB.Recordset, _
    ByVal strColHead As String, _
    Optional ByVal intReturnCol As Integer = 0, _
    Optional ByVal strCaption As String = "选择项目", _
    Optional ByVal blnSorted As Boolean = True, _
    Optional ByVal strDefItem As String = "", _
    Optional ByVal lngDefIndex As Long = -1, _
    Optional ByVal IconIndex As Long = 1, _
    Optional ByVal lngWidth As Long = 0, _
    Optional ByVal lngHeight As Long = 0) As String
    '功能:外部调用本选择器之用
    
    '先检查
    ShowCurrSel = ""
    mintReturnCol = intReturnCol
    mstrColHead = strColHead
    mblnSorted = blnSorted
    mstrDefItem = strDefItem
    mlngDefIndex = lngDefIndex
    mIconIndex = IconIndex
    
    mlngWidth = lngWidth
    mlngHeight = lngHeight
    Set mRs = rs
    
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount < 1 Then
        Exit Function
    End If
    If InStr(strColHead, ",") < 1 Then Exit Function
    If intReturnCol > -1 Then
        If InStr(strColHead, ";") > 0 Then
            If intReturnCol > UBound(Split(strColHead, ";")) Then Exit Function
        Else
            If intReturnCol > 0 Then intReturnCol = 0
        End If
    End If
    
    Me.Caption = strCaption
    '显示窗体
    Me.Show 1, frmParent
    
    ShowCurrSel = mstrReturn
    
    
End Function

Private Sub cmdCancel_Click()
    mstrReturn = ""
    Unload Me
End Sub

Private Sub cmdFind_Click()
    '进行查找
    Dim i As Long
    Dim lngCol As Long
    Dim ObjItem As ListItem
    Dim lngSel As Long
    
    If Trim(Me.txtCol.Text) = "" Then Exit Sub
    If Trim(Me.txtValues.Text) = "" Then Exit Sub
    If lvw.ListItems.Count < 1 Then Exit Sub
    If lvw.SelectedItem Is Nothing Then
ReDo:   lngSel = 1
    Else
        lngSel = lvw.SelectedItem.Index + 1
    End If
    
    For lngCol = 1 To lvw.ColumnHeaders.Count
        If Trim(lvw.ColumnHeaders(lngCol).Text) = Trim(Me.txtCol.Text) Then
            For i = lngSel To lvw.ListItems.Count
                Set ObjItem = lvw.ListItems(i)
                If lngCol = 1 Then
                    If Opt(0).Value Then
                        If Trim(ObjItem.Text) Like Trim(Me.txtValues.Text) Then
                            ObjItem.Selected = True
                            ObjItem.EnsureVisible
                            Exit Sub
                        End If
                    Else
                        If Trim(ObjItem.Text) Like Trim(Me.txtValues.Text) & "*" Then
                            ObjItem.Selected = True
                            ObjItem.EnsureVisible
                            Exit Sub
                        End If
                    End If
                Else
                    If Opt(0).Value Then
                        If Trim(ObjItem.SubItems(lngCol - 1)) Like Trim(Me.txtValues.Text) Then
                            ObjItem.Selected = True
                            ObjItem.EnsureVisible
                            Exit Sub
                        End If
                    Else
                        If Trim(ObjItem.SubItems(lngCol - 1)) Like Trim(Me.txtValues.Text) & "*" Then
                            ObjItem.Selected = True
                            ObjItem.EnsureVisible
                            Exit Sub
                        End If
                    End If
                End If
                If i = lvw.ListItems.Count And i > 10 Then
                    If MsgBox("已经查询到末尾是否重新开始？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                        GoTo ReDo
                    Else
                        Exit Sub
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    If Not lvw.SelectedItem Is Nothing Then
        If mintReturnCol > -1 Then
            If mintReturnCol = 0 Then
                mstrReturn = lvw.SelectedItem.Text
            Else
                mstrReturn = lvw.SelectedItem.SubItems(mintReturnCol)
            End If
        Else
            '如果返回的列索引为 < 0 的表示返所有列并以 ; 分隔
            For i = 0 To lvw.ColumnHeaders.Count - 1
                If i = 0 Then
                    mstrReturn = lvw.SelectedItem.Text & ";"
                Else
                    mstrReturn = mstrReturn & lvw.SelectedItem.SubItems(i) & ";"
                End If
            Next
            mstrReturn = Left(mstrReturn, Len(mstrReturn) - 1)
        End If
    Else
        MsgBox "请选择项目！", vbExclamation, gstrSysName
        lvw.SetFocus
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
    
    Dim i As Long, j As Long
    Dim ObjItem As ListItem
    Dim ObjDistItem As ListItem
    
    Screen.MousePointer = vbHourglass
    '设置选择器
    zlControl.LvwSelectColumns Me.lvw, mstrColHead, True
    If mblnSorted = False Then
        zlControl.LvwFlatColumnHeader Me.lvw
    End If
    
    Me.lvw.Sorted = False
    Me.lvw.ListItems.Clear
    mRs.MoveFirst
    For i = 1 To mRs.RecordCount
        Set ObjItem = lvw.ListItems.Add(, , zlCommFun.NVL(mRs(0).Value), mIconIndex, mIconIndex)
        For j = 1 To mRs.Fields.Count - 1
            ObjItem.SubItems(j) = zlCommFun.NVL(mRs(j).Value)
            If j = mlngDefIndex And ObjItem.SubItems(j) = mstrDefItem Then
                Set ObjDistItem = ObjItem
            End If
        Next
        mRs.MoveNext
    Next
    
    If mlngWidth <> 0 Then
        If mlngWidth < 1000 Then mlngWidth = 1000
        Me.Width = mlngWidth
    End If
    If mlngHeight <> 0 Then
        If mlngHeight < 1000 Then mlngHeight = 1000
        Me.Height = mlngHeight
    End If
    Me.lvw.Sorted = True
    If Not ObjDistItem Is Nothing Then
        ObjDistItem.Selected = True
        ObjDistItem.EnsureVisible
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    Dim i As Long
    Me.cmdOk.Left = Me.ScaleWidth - (Me.cmdCancel.Width + Me.lvw.Left * 1)
    Me.cmdCancel.Left = Me.cmdOk.Left
    Me.fra1.Left = Me.cmdOk.Left
    Me.fra1.Width = Me.cmdOk.Width
    
    i = Me.ScaleWidth - (Me.lvw.Left * 3.5 + Me.cmdOk.Width)
    Me.lvw.Width = IIf(i > 0, i, Screen.TwipsPerPixelX)
    
    i = Me.ScaleHeight - (Me.lvw.Top * 2)
    Me.lvw.Height = IIf(i > 0, i, Screen.TwipsPerPixelY)
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    '进行排序
    Me.txtCol.Text = ColumnHeader.Text
    Me.txtValues.Text = ""
    If mblnSorted = False Then Exit Sub
    Me.lvw.Sorted = True
    If Me.lvw.SortKey = ColumnHeader.Index - 1 Then
        If Me.lvw.SortOrder <> lvwDescending Then
            Me.lvw.SortOrder = lvwDescending
        Else
            Me.lvw.SortOrder = lvwAscending
        End If
    Else
        Me.lvw.SortKey = ColumnHeader.Index - 1
        Me.lvw.SortOrder = lvwAscending
    End If
    
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lngCol As Long
    Dim i As Long
    
    For lngCol = 1 To lvw.ColumnHeaders.Count
        If Trim(lvw.ColumnHeaders(lngCol).Text) = Trim(Me.txtCol.Text) Then
            If lngCol = 1 Then
                Me.txtValues.Text = Item.Text
            Else
                Me.txtValues.Text = Item.SubItems(lngCol - 1)
            End If
        End If
    Next
End Sub

Private Sub lvw_DblClick()
    cmdOK_Click
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOK_Click
    End If
End Sub

Private Sub txtValues_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub
