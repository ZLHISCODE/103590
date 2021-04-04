VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelCur 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "选择器"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4530
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3345
      TabIndex        =   1
      Top             =   255
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3345
      TabIndex        =   2
      Top             =   810
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3990
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   7038
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
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelCur.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelCur.frx":0452
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelCur.frx":05AC
            Key             =   "Write"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSelCur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrReturn As String    '返回字符串
Private mstrColHead As String
Private mRs As ADODB.Recordset
Private mblnSorted As Boolean
Private mstrDefItem As String
Private mlngDefIndex As Long

Private mblnMulti As Boolean '是否可多选

Public Function ShowCurrSel(frmParent As Object, rs As ADODB.Recordset, _
                        ByVal strColHead As String, Optional ByVal strCaption As String = "选择项目", _
                        Optional ByVal blnSorted As Boolean = True, Optional ByVal strDefItem As String = "", _
                        Optional ByVal lngDefIndex As Long = -1, Optional ByVal frmWidth As Long = 0, Optional ByVal blnMulti As Boolean = False) As String
'功能:外部调用本选择器之用
'参数:  rs              显示的数据
'       strColHead      列头字符串
'       strCaption      选择器标题名称
'       blnSorted       是否允许用户进行排序
'       strDefItem      用来指定由lngDefIndex指定列的字符串的那一行为默认选定的行
'       lngDefIndex     用来确定默认行的比较列索引
'       frmwidth        用来指定选择器的宽度    如果缺省则为0
On Error GoTo errHandle
Dim i As Long
    
    '先检查
    ShowCurrSel = ""
    mstrColHead = strColHead
    mblnSorted = blnSorted
    mstrDefItem = strDefItem
    mlngDefIndex = lngDefIndex
    mblnMulti = blnMulti
    mstrReturn = ""
    Set mRs = rs
    
    If rs Is Nothing Then
        Exit Function
    End If
    If rs.RecordCount < 1 Then
        Exit Function
    End If
    If InStr(strColHead, ",") < 1 Then Exit Function
    If frmWidth > 0 Then Me.Width = frmWidth
    Me.Caption = strCaption
    '显示窗体
    Me.Show 1, frmParent
    
    ShowCurrSel = mstrReturn
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    mstrReturn = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle
    Dim strTmp As String
    Dim i As Long, intLoop As Integer
    If mblnMulti Then
        '多选
        strTmp = ""
        For intLoop = 1 To lvw.ListItems.Count
            
            If lvw.ListItems(intLoop).Checked Then
                strTmp = ""
                For i = 1 To Me.lvw.ColumnHeaders.Count - 1
                    strTmp = strTmp & Me.lvw.ListItems(intLoop).SubItems(i) & ","
                Next
                strTmp = Me.lvw.ListItems(intLoop).Text & "," & strTmp
                strTmp = Left(strTmp, Len(strTmp) - 1)
                mstrReturn = mstrReturn & "|" & strTmp
            End If
        Next
        If mstrReturn <> "" Then mstrReturn = Mid(mstrReturn, 2)
    ElseIf Not lvw.SelectedItem Is Nothing Then
        strTmp = ""
        For i = 1 To Me.lvw.ColumnHeaders.Count - 1
            strTmp = strTmp & Me.lvw.SelectedItem.SubItems(i) & ","
        Next
        strTmp = Me.lvw.SelectedItem.Text & "," & strTmp
        strTmp = Left(strTmp, Len(strTmp) - 1)
        mstrReturn = strTmp
    Else
        MsgBox "请选择项目！", vbExclamation, gstrSysName
        lvw.SetFocus
        Exit Sub
    End If
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
Dim i As Long, j As Long
Dim lngRecCount As Long
Dim objItem As ListItem
    
    '设置选择器
    zlControl.LvwSelectColumns Me.lvw, mstrColHead, True
    If mblnSorted = False Then
        zlControl.LvwFlatColumnHeader Me.lvw
    End If
    
    Me.lvw.Sorted = False
    Me.lvw.ListItems.Clear
    '在填充数据之前设置CheckBox样式
    If mblnMulti Then
        lvw.Checkboxes = True
    End If
    
    
    mRs.MoveFirst
    lngRecCount = mRs.RecordCount
    If lngRecCount > 2000 Then lngRecCount = 2000
    For i = 1 To lngRecCount
        Set objItem = lvw.ListItems.Add(, , Nvl(mRs(0).Value), "Root", "Root")
        For j = 1 To mRs.Fields.Count - 1
            objItem.SubItems(j) = Nvl(mRs(j).Value)
            If j = mlngDefIndex And objItem.SubItems(j) = mstrDefItem Then
                objItem.Selected = True
                objItem.EnsureVisible
            End If
        Next
        mRs.MoveNext
    Next
    Me.lvw.Sorted = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Me.cmdOK.Left = Me.ScaleWidth - (Me.lvw.Left * 2 + Me.cmdOK.Width - Screen.TwipsPerPixelX * 4)
    Me.cmdCancel.Left = Me.cmdOK.Left
    Me.lvw.Width = Me.cmdOK.Left - Me.lvw.Left * 2
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    '进行排序
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

Private Sub lvw_DblClick()
    cmdOK_Click
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOK_Click
    End If
End Sub

