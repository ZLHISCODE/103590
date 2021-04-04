VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ucSplitPageNew 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   ScaleHeight     =   300
   ScaleWidth      =   5775
   Begin MSComCtl2.FlatScrollBar FScroll 
      Height          =   300
      Left            =   3000
      TabIndex        =   4
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   529
      _Version        =   393216
      Arrows          =   65536
      Max             =   0
      Orientation     =   1572865
   End
   Begin VB.ComboBox cbxPage 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox txtPageRecord 
      Height          =   300
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   0
      ToolTipText     =   "每页数量"
      Top             =   15
      Width           =   300
   End
   Begin MSComCtl2.UpDown udPageRecord 
      Height          =   300
      Left            =   1981
      TabIndex        =   6
      Top             =   15
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtPageRecord"
      BuddyDispid     =   196610
      OrigLeft        =   5320
      OrigRight       =   5575
      OrigBottom      =   285
      Max             =   99
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label labTotal 
      AutoSize        =   -1  'True
      Caption         =   "共："
      Height          =   180
      Left            =   2280
      TabIndex        =   5
      Top             =   60
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "           页"
      Height          =   255
      Left            =   -75
      TabIndex        =   3
      Top             =   60
      Width           =   1260
   End
   Begin VB.Label labPageCount 
      Caption         =   "每页："
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   60
      Width           =   615
   End
End
Attribute VB_Name = "ucSplitPageNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mlngPageCount As Long   '页数量
Private mlngRecordCount As Long '记录总数
Private mlngPageIndex As Long
Private mblnIsConfiging As Boolean
Private mlngPageRecord As Long
Private mlngItemIndex As Long
Private mblnImageChaged As Boolean

Private Const CON_STR_HINT_TITLE = "提示"

'页改变事件
Public Event OnBeforeImageChange(ByRef Cancel As Boolean)
Public Event OnPageChange(ByVal lngPageIndex As Long, ByVal lngPageCount As Long)
Public Event OnItemChange(ByVal lngPageIndex As Long, ByVal lngPageRecord As Long)
Public Event OnPageRecordChange(ByVal lngPageRecord As Long)


'
Property Get AutoRedrawStyle() As Boolean
    AutoRedrawStyle = AutoRedraw
End Property

Property Let AutoRedrawStyle(value As Boolean)
    AutoRedraw = value
End Property


'获取页数
Property Get PageCount() As Long
    PageCount = mlngPageCount
End Property


'设置记录数量
Property Get RecordCount() As Long
    RecordCount = mlngRecordCount
End Property

Property Let RecordCount(value As Long)
    mlngRecordCount = value
    
    If Val(txtPageRecord.Text) <= 0 Then Exit Property
    
    If value <= 0 Then
        mlngPageCount = 0
    Else
        mlngPageCount = IIf(value Mod Val(txtPageRecord.Text) > 0, Fix(value / Val(txtPageRecord.Text)) + 1, value / Val(txtPageRecord.Text))
    End If
    
    Call ConfigPageCount(mlngPageCount)
End Property



'设置每页的记录数
Property Get PageRecord() As Long
    PageRecord = Val(txtPageRecord.Text)
End Property


Property Let PageRecord(value As Long)
    If value <= 0 Then
        txtPageRecord.Text = 6
    Else
        txtPageRecord.Text = value
    End If
    
    If mlngRecordCount <= 0 Then
        mlngPageCount = 0
    Else
        mlngPageCount = IIf(mlngRecordCount Mod Val(txtPageRecord.Text) > 0, Fix(mlngRecordCount / Val(txtPageRecord.Text)) + 1, mlngRecordCount / Val(txtPageRecord.Text))
    End If
    
    Call ConfigPageCount(mlngPageCount)
End Property


'获取当前页索引
Property Get PageIndex() As Long
    PageIndex = cbxPage.ListIndex
End Property

Property Let PageIndex(value As Long)
    
    mlngPageIndex = value
    cbxPage.ListIndex = value
End Property

Property Get PageNumber() As Long
'获取页号码
    
    PageNumber = GetPageNum(True)
End Property

Property Get ItemIndex() As Long
'获取索引

    ItemIndex = mlngItemIndex
End Property

Property Let ItemIndex(value As Long)
    mlngItemIndex = value
    
    If value > FScroll.Max Then
        FScroll.Max = value
    End If
    FScroll.value = value
    
    mlngPageIndex = cbxPage.ListIndex
End Property


Private Function GetPageNum(blnType As Boolean) As Long
'blnType=true:获取显示页号 blnType=false:获取显示总页数
    If cbxPage.ListCount > 0 Then
        GetPageNum = Val(Split(cbxPage.Text & "/", "/")(IIf(blnType, 0, 1)))
    Else
        GetPageNum = 0
    End If
End Function



Private Sub cbxPage_Change()
    
    mlngPageIndex = cbxPage.ListIndex
End Sub

Private Sub cbxPage_Click()
    Dim blnCancel As Boolean
    
On Error Resume Next
    
    If mlngPageIndex = cbxPage.ListIndex Then Exit Sub
    RaiseEvent OnBeforeImageChange(blnCancel)
    
    If blnCancel Then
        cbxPage.ListIndex = mlngPageIndex
        Exit Sub
    End If
    
    If mblnIsConfiging Then Exit Sub
    If GetPageNum(False) <= 0 Then Exit Sub
    
    If GetPageNum(True) <= mlngPageCount Then
        FScroll.value = (GetPageNum(True) - 1) * Val(txtPageRecord.Text) + 1
    End If
    
    mlngPageIndex = cbxPage.ListIndex
    mblnImageChaged = True
    RaiseEvent OnPageChange(cbxPage.ListIndex + 1, Val(txtPageRecord.Text))
End Sub


Private Sub ConfigPageCount(ByVal lngPageCount As Long)
'配置页数量显示
On Error GoTo errHandle
    Dim i As Long
    Dim lngOldIndex As Long
    
    mblnIsConfiging = True
    
    labTotal.Caption = "共：" & mlngRecordCount
    
    lngOldIndex = cbxPage.ListIndex
    
    cbxPage.Clear
    
    If lngPageCount <= 0 Then
        cbxPage.AddItem "0/0"
        cbxPage.ListIndex = 0
        
        Exit Sub
    End If
    
    FScroll.Min = 1
    FScroll.Max = mlngRecordCount
    For i = 0 To lngPageCount - 1
        cbxPage.AddItem (i + 1) & "/" & lngPageCount
    Next i
    
    If lngOldIndex >= 0 And lngOldIndex < cbxPage.ListCount Then
        cbxPage.ListIndex = lngOldIndex
    ElseIf lngOldIndex - 1 >= 0 And lngOldIndex - 1 < cbxPage.ListCount Then
        cbxPage.ListIndex = lngOldIndex - 1
    Else
        cbxPage.ListIndex = 0
    End If
    
    mblnIsConfiging = False
    
errHandle:
End Sub


'Private Sub cmdPage_Click(Index As Integer)
'On Error Resume Next
'
'    Select Case Index
'        Case 0
'            If cbxPage.ListIndex = 0 Then Exit Sub
'
'            cbxPage.ListIndex = 0
'
'        Case 1
'            If cbxPage.ListIndex = 0 Then   '如果已经是第一页，则移到最后一页
'                cbxPage.ListIndex = cbxPage.ListCount - 1
'                Exit Sub
'            End If
'
'            If cbxPage.ListIndex > 0 Then cbxPage.ListIndex = cbxPage.ListIndex - 1
'
'        Case 2
'            If cbxPage.ListIndex = cbxPage.ListCount - 1 Then   '如果是最后一页，则移到第一页
'                cbxPage.ListIndex = 0
'                Exit Sub
'            End If
'
'            If cbxPage.ListIndex < cbxPage.ListCount - 1 Then cbxPage.ListIndex = cbxPage.ListIndex + 1
'
'        Case 3
'            If cbxPage.ListIndex = cbxPage.ListCount - 1 Then Exit Sub
'
'            cbxPage.ListIndex = cbxPage.ListCount - 1
'
'    End Select
'
'    RaiseEvent OnPageChange(cbxPage.ListIndex + 1, Val(txtPageRecord.Text))
'
'    Err.Clear
'End Sub

Public Sub MovePage(ByVal lngPage As Long)
'移动到指定页面
    
    If lngPage <= 0 Or lngPage > mlngPageCount Then Exit Sub
    
    mblnIsConfiging = True
    cbxPage.ListIndex = lngPage - 1
    mlngPageIndex = cbxPage.ListIndex
    mblnIsConfiging = False
    
    RaiseEvent OnPageChange(lngPage, Val(txtPageRecord.Text))
End Sub

Public Sub FirstPage()
'第一页
On Error Resume Next
    If mlngPageCount = 0 Then Exit Sub

    mblnIsConfiging = True
    cbxPage.ListIndex = 0
    mblnIsConfiging = False

    RaiseEvent OnPageChange(cbxPage.ListIndex + 1, Val(txtPageRecord.Text))

    err.Clear
End Sub

Public Sub LastPage()
'上页
    If mlngPageCount = 0 Then Exit Sub

    If cbxPage.ListIndex = 0 Then   '如果已经是第一页，则移到最后一页
        cbxPage.ListIndex = cbxPage.ListCount - 1
        Exit Sub
    End If

    If cbxPage.ListIndex > 0 Then cbxPage.ListIndex = cbxPage.ListIndex - 1
    
    RaiseEvent OnPageChange(cbxPage.ListIndex + 1, Val(txtPageRecord.Text))
End Sub

Public Sub NextPage()
'下页
    If mlngPageCount = 0 Then Exit Sub

    If cbxPage.ListIndex = cbxPage.ListCount - 1 Then   '如果是最后一页，则移到第一页
        cbxPage.ListIndex = 0
        Exit Sub
    End If

    If cbxPage.ListIndex < cbxPage.ListCount - 1 Then cbxPage.ListIndex = cbxPage.ListIndex + 1
    
    RaiseEvent OnPageChange(cbxPage.ListIndex + 1, Val(txtPageRecord.Text))
End Sub

Public Sub EndPage()
'最后页
On Error Resume Next
    If mlngPageCount = 0 Then Exit Sub

    mblnIsConfiging = True
    cbxPage.ListIndex = cbxPage.ListCount - 1
    mblnIsConfiging = False

    RaiseEvent OnPageChange(cbxPage.ListIndex + 1, Val(txtPageRecord.Text))

    err.Clear
End Sub

Public Sub RedrawSelf()

    FScroll.Refresh
    
    Label1.Refresh
'    labTotal.Refresh
    labPageCount.Refresh
    
    cbxPage.Refresh
    txtPageRecord.Refresh
    
End Sub

Private Sub FScroll_Change()
    If mlngItemIndex = FScroll.value Then Exit Sub
    
    If FScroll.value <= mlngRecordCount Then
        MoveItem FScroll.value
    End If
    
    mlngItemIndex = FScroll.value
End Sub

Public Sub MoveItem(ByVal lngObjIndex As Long)
    Dim lngMovePageNum As Long
    Dim lngMovePageIndex As Long
    Dim blnCancel As Boolean
    
    If mlngPageRecord = 0 Then Exit Sub
    If lngObjIndex = 0 Then Exit Sub
    
    lngMovePageIndex = lngObjIndex Mod mlngPageRecord
    lngMovePageNum = lngObjIndex \ mlngPageRecord
    
    If lngMovePageIndex > 0 Then
        lngMovePageNum = lngMovePageNum + 1
    Else
        lngMovePageIndex = mlngPageRecord
    End If
    
    If lngMovePageNum <> GetPageNum(True) Then
        RaiseEvent OnBeforeImageChange(blnCancel)
        FScroll.value = mlngItemIndex
        If blnCancel Then Exit Sub
        
        MovePage lngMovePageNum
    End If
    
    RaiseEvent OnItemChange(lngMovePageIndex, lngMovePageNum)
End Sub


Private Sub txtPageRecord_Change()
'改变每页显示记录数量
    Dim blnCancel As Boolean

On Error GoTo errHandle
    
    If Val(txtPageRecord.Text) = mlngPageRecord Then Exit Sub
    
    RaiseEvent OnBeforeImageChange(blnCancel)
    
    If blnCancel Then
        txtPageRecord.Text = mlngPageRecord
        Exit Sub
    End If
    
    mlngPageRecord = Val(txtPageRecord.Text)
    
    If Val(txtPageRecord.Text) <= 0 Then Exit Sub
    If mlngRecordCount <= 0 Then Exit Sub
    
    If mlngRecordCount <= 0 Then
        mlngPageCount = 0
    Else
        mlngPageCount = IIf(mlngRecordCount Mod Val(txtPageRecord.Text) > 0, Fix(mlngRecordCount / Val(txtPageRecord.Text)) + 1, mlngRecordCount / Val(txtPageRecord.Text))
    End If
    
    Call ConfigPageCount(mlngPageCount)
    
    RaiseEvent OnPageRecordChange(Val(txtPageRecord.Text))
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    MsgboxEx hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE
End Sub

Private Sub txtPageRecord_GotFocus()
    On Error GoTo errHandle

    mlngPageRecord = Val(txtPageRecord.Text)
    
    Exit Sub
errHandle:
    MsgboxEx hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE
End Sub

Private Sub txtPageRecord_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandle

    If InStr("0123456789", Chr(KeyAscii)) = 0 And Chr(KeyAscii) <> vbBack Then
        KeyAscii = 0
    End If
    
    If Val(txtPageRecord.Text & Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
    
    Exit Sub
errHandle:
    MsgboxEx hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE
End Sub

Private Sub txtPageRecord_LostFocus()
    On Error GoTo errHandle

    If Len(txtPageRecord.Text) = 0 Then
        txtPageRecord.Text = mlngPageRecord
    End If
    
    Exit Sub
errHandle:
    MsgboxEx hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE
End Sub


Private Sub UserControl_Initialize()
'初始化组件
    mblnIsConfiging = False
    mlngPageCount = 0
    mlngPageRecord = 9
    txtPageRecord.Text = 9
    mlngItemIndex = 0
    
    Call RefrshLayout
    Call ConfigPageCount(mlngPageCount)
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'属性读取
    mlngPageCount = PropBag.ReadProperty("PageCount", 0)
    mlngPageRecord = PropBag.ReadProperty("PageRecord", 9)
    
    AutoRedraw = PropBag.ReadProperty("AutoRedrawStyle", True)
    
    txtPageRecord.Text = mlngPageRecord
    Call ConfigPageCount(mlngPageCount)
End Sub

    
Private Sub UserControl_Resize()
On Error Resume Next
'    Width = 5955
'    Height = 330
    
    Call RefrshLayout
    err.Clear
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'写入属性
    Call PropBag.WriteProperty("PageCount", mlngPageCount)
    Call PropBag.WriteProperty("PageRecord", mlngPageRecord)
    Call PropBag.WriteProperty("AutoRedrawStyle", AutoRedraw, True)
End Sub

Private Sub RefrshLayout()
    FScroll.Top = 0
    FScroll.Height = 300
    FScroll.Left = 3000
    FScroll.Width = Width - FScroll.Left
    cbxPage.Top = 0
    Label1.Top = 60
    labPageCount.Top = 60
    labTotal.Top = 60
    txtPageRecord.Top = 0
    udPageRecord.Top = 0
End Sub
