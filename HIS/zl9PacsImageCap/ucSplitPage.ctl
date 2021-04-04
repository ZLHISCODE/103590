VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ucSplitPage 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5955
   ScaleHeight     =   300
   ScaleWidth      =   5955
   ToolboxBitmap   =   "ucSplitPage.ctx":0000
   Begin VB.TextBox txtPageRecord 
      Height          =   285
      Left            =   5430
      MaxLength       =   2
      TabIndex        =   9
      Text            =   "9"
      ToolTipText     =   "每页数量"
      Top             =   0
      Width           =   265
   End
   Begin MSComCtl2.UpDown udPageRecord 
      Height          =   285
      Left            =   5700
      TabIndex        =   8
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtPageRecord"
      BuddyDispid     =   196609
      OrigLeft        =   5320
      OrigRight       =   5575
      OrigBottom      =   285
      Max             =   99
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   ">|"
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   5
      ToolTipText     =   "尾页"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   ">>"
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   4
      ToolTipText     =   "下一页"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "<<"
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "上一页"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "|<"
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "首页"
      Top             =   0
      Width           =   375
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
      Left            =   1750
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Label labPageCount 
      Caption         =   "每页："
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   60
      Width           =   615
   End
   Begin VB.Label labTotal 
      Caption         =   "总数：0           "
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "第          页"
      Height          =   255
      Left            =   1480
      TabIndex        =   6
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label labPage 
      Caption         =   "共 1 页"
      Height          =   255
      Left            =   2830
      TabIndex        =   3
      Top             =   60
      Width           =   855
   End
End
Attribute VB_Name = "ucSplitPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngPageCount As Long   '页数量
Private mlngRecordCount As Long '记录总数
Private mblnIsConfiging As Boolean


'页改变事件
Public Event OnPageChange(ByVal lngPageIndex As Long, ByVal lngPageCount As Long)


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


Property Get PageNumber() As Long
'获取页号码
    If cbxPage.ListCount > 0 Then
        PageNumber = Val(cbxPage.Text)
    Else
        PageNumber = 0
    End If
End Property


Private Sub cbxPage_Click()
On Error Resume Next
    If mblnIsConfiging Then Exit Sub
    If Val(cbxPage.Text) <= 0 Then Exit Sub
    
    RaiseEvent OnPageChange(cbxPage.ListIndex + 1, Val(txtPageRecord.Text))
End Sub


Private Sub ConfigPageCount(ByVal lngPageCount As Long)
'配置页数量显示
On Error GoTo errHandle
    Dim i As Long
    Dim lngOldIndex As Long
    
    mblnIsConfiging = True
    
    labPage.Caption = "共 " & lngPageCount & " 页"
    labTotal.Caption = "总数：" & mlngRecordCount
    
    lngOldIndex = cbxPage.ListIndex
    
    cbxPage.Clear
    
    If lngPageCount <= 0 Then
        cbxPage.AddItem "0"
        cbxPage.ListIndex = 0
        
        Exit Sub
    End If
    
    For i = 0 To lngPageCount - 1
        cbxPage.AddItem i + 1
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


Private Sub cmdPage_Click(Index As Integer)
On Error Resume Next

    Select Case Index
        Case 0
            If cbxPage.ListIndex = 0 Then Exit Sub
            
            cbxPage.ListIndex = 0
            
        Case 1
            If cbxPage.ListIndex = 0 Then Exit Sub
            
            If cbxPage.ListIndex > 0 Then cbxPage.ListIndex = cbxPage.ListIndex - 1
            
        Case 2
            If cbxPage.ListIndex = cbxPage.ListCount - 1 Then Exit Sub
            
            If cbxPage.ListIndex < cbxPage.ListCount - 1 Then cbxPage.ListIndex = cbxPage.ListIndex + 1
            
        Case 3
            If cbxPage.ListIndex = cbxPage.ListCount - 1 Then Exit Sub
            
            cbxPage.ListIndex = cbxPage.ListCount - 1
            
    End Select
    
    RaiseEvent OnPageChange(cbxPage.ListIndex + 1, Val(txtPageRecord.Text))
    
    err.Clear
End Sub

Public Sub MovePage(ByVal lngPage As Long)
'移动到指定页面
    If lngPage <= 0 Or lngPage > mlngPageCount Then Exit Sub
    
    mblnIsConfiging = True
    cbxPage.ListIndex = lngPage - 1
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
    
    Call cmdPage_Click(1)
End Sub

Public Sub NextPage()
'下页
    If mlngPageCount = 0 Then Exit Sub
    
    Call cmdPage_Click(2)
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

    cmdPage.Item(0).Refresh
    cmdPage.Item(1).Refresh
    cmdPage.Item(2).Refresh
    cmdPage.Item(3).Refresh
    
    Label1.Refresh
    labPage.Refresh
    labTotal.Refresh
    labPageCount.Refresh
    
    cbxPage.Refresh
    txtPageRecord.Refresh
    
End Sub


Private Sub txtPageRecord_Change()
'改变每页显示记录数量
On Error GoTo errHandle
    
    If Val(txtPageRecord.Text) <= 0 Then Exit Sub
    If mlngRecordCount <= 0 Then Exit Sub
    
    If mlngRecordCount <= 0 Then
        mlngPageCount = 0
    Else
        mlngPageCount = IIf(mlngRecordCount Mod Val(txtPageRecord.Text) > 0, Fix(mlngRecordCount / Val(txtPageRecord.Text)) + 1, mlngRecordCount / Val(txtPageRecord.Text))
    End If
    
    Call ConfigPageCount(mlngPageCount)
    
    RaiseEvent OnPageChange(cbxPage.ListIndex + 1, Val(txtPageRecord.Text))
Exit Sub
errHandle:
    'If ErrCenter() = 1 Then Resume
    MsgboxEx hWnd, err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub UserControl_Initialize()
'初始化组件
    mblnIsConfiging = False
    mlngPageCount = 0
    txtPageRecord.Text = 9
    
    Call ConfigPageCount(mlngPageCount)
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'属性读取
    mlngPageCount = PropBag.ReadProperty("PageCount", 0)
    txtPageRecord.Text = PropBag.ReadProperty("PageRecord", 9)
    AutoRedraw = PropBag.ReadProperty("AutoRedrawStyle", True)
    
    Call ConfigPageCount(mlngPageCount)
End Sub

    
Private Sub UserControl_Resize()
On Error Resume Next
    Width = 5955
    Height = 330

    err.Clear
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'写入属性
    Call PropBag.WriteProperty("PageCount", mlngPageCount)
    Call PropBag.WriteProperty("PageRecord", Val(txtPageRecord.Text))
    Call PropBag.WriteProperty("AutoRedrawStyle", AutoRedraw, True)
End Sub
