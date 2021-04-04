VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl WordInputModule 
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   5520
   ScaleWidth      =   4800
   ToolboxBitmap   =   "WordInputModule.ctx":0000
   Begin RichTextLib.RichTextBox txtWordContext 
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   3720
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1931
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"WordInputModule.ctx":0312
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView trvWord 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5953
      _Version        =   393217
      Indentation     =   353
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   1680
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
            Picture         =   "WordInputModule.ctx":03AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WordInputModule.ctx":0AA9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpWordModule 
      Left            =   3960
      Top             =   480
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "WordInputModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'词库模板类型
Private mstrModuleName As String
Private mlngDepartId As Long
'Private mlngWordHeight As Long


Public Event OnWordDbClickEvent(ByVal strWord As String)




'模板名称
Property Get ModuleName() As String
    ModuleName = mstrModuleName
End Property


Property Let ModuleName(value As String)
    mstrModuleName = value
End Property


'当前科室ID
Property Get CurDepartId() As Long
    CurDepartId = mlngDepartId
End Property


Property Let CurDepartId(value As Long)
    mlngDepartId = value
End Property




Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property



'Property Get WordHeight() As Long
'    WordHeight = mlngWordHeight
'End Property
'
'Property Let WordHeight(value As Long)
'    If mlngWordHeight = value Then Exit Property
'
'    mlngWordHeight = value
'
'    Call InitFace
'End Property




Private Sub InitFace()
'初始化界面布局
    Dim Pane1 As Pane, Pane2 As Pane
    Dim lngWordHeight As Long
    
    With dkpWordModule
        .CloseAll
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    lngWordHeight = Round(Height / 3 * 2)
    
    Set Pane1 = dkpWordModule.CreatePane(1, 0, lngWordHeight, DockTopOf, Nothing)
    Pane1.Title = "词句模板"
    Pane1.Handle = trvWord.hWnd
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Pane1.MinTrackSize.Width = 50
    Pane1.MinTrackSize.Height = 50
    
    Set Pane2 = dkpWordModule.CreatePane(2, 0, Height - lngWordHeight, DockBottomOf, Pane1)
    Pane2.Title = "词句内容"
    Pane2.Handle = txtWordContext.hWnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Pane2.MinTrackSize.Width = 50
    Pane2.MinTrackSize.Height = 50
End Sub



Public Sub LoadWordModel()
'载入词库模板
    Dim strSql As String
    Dim rsWordClass As ADODB.Recordset
    Dim rsWordContext As ADODB.Recordset
    
    
    Dim objNode As Node
    Dim objPnode As Node
    Dim objWord As Node
    
    strSql = "select distinct id, 上级id, 编码, 名称, 范围 from 病历词句分类 a" & _
                " start with id in ( " & _
                " select distinct id from  病历词句分类 where 名称='" & mstrModuleName & "')" & _
                " connect by prior id=上级ID order by 编码 "
                
    
    Set rsWordClass = zlDatabase.OpenSQLRecord(strSql, "词库模板")
    
    '清除树结构数据
    Call trvWord.Nodes.Clear
    
    If rsWordClass.RecordCount <= 0 Then Exit Sub
    
    
'select a.分类ID, a.名称, b.词句ID, b.内容文本
'                 from  病历词句示范 a, 病历词句组成 b , 病历词句分类 c
'                 where a.分类ID=c.id and a.id=b.词句id and c.名称='巨检描述'
'                 order by 名称
                 
                 
    strSql = "select a.id,a.分类ID, a.名称, b.词句ID, b.内容文本 " & _
                 " from  病历词句示范 a, 病历词句组成 b " & _
                 " Where a.ID = b.词句id and b.内容性质=0 and b.排列次序=1 And ((a.通用级 = 2 And a.人员ID = " & UserInfo.ID & ") Or a.通用级 = 0 or a.通用级 is null Or (a.通用级 = 1 And a.科室ID = " & mlngDepartId & ")) " & _
                 " and a.分类id in (select distinct id from 病历词句分类 " & _
                                                         " start with id in (select distinct id from  病历词句分类 where 名称='" & mstrModuleName & "') " & _
                                                         " connect by prior id=上级ID) " & _
                                                         " order by 名称 "
                
    Set rsWordContext = zlDatabase.OpenSQLRecord(strSql, "模板词句")
    
    Do While Not rsWordClass.EOF
        
        Set objNode = Nothing
        
        On Error Resume Next
        Set objNode = trvWord.Nodes("T-" & rsWordClass("ID").value)
        If zlCommFun.Nvl(rsWordClass("上级id").value, 0) <> 0 Then
            Set objPnode = trvWord.Nodes("T-" & rsWordClass("上级id").value)
        Else
            Set objPnode = Nothing
        End If
        
        On Error GoTo errHandle
        
        If objNode Is Nothing Then
            If objPnode Is Nothing Then
                Set objNode = trvWord.Nodes.Add(, , "T-" & rsWordClass("ID").value, rsWordClass("名称").value, 2)
            Else
                Set objNode = trvWord.Nodes.Add("T-" & Nvl(rsWordClass("上级id").value, 0), tvwChild, "T-" & rsWordClass("ID").value, rsWordClass("名称").value, 2)
            End If
            
            If Not objNode.Parent Is Nothing Then
                objNode.Expanded = False
            Else
                objNode.Expanded = True
            End If
            
            '读取词句
            rsWordContext.Filter = "分类ID='" & rsWordClass("ID").value & "'"
            If rsWordContext.RecordCount > 0 Then
                rsWordContext.MoveFirst
                While Not rsWordContext.EOF
                    Set objWord = trvWord.Nodes.Add("T-" & Nvl(rsWordClass("ID").value, 0), tvwChild, "W-" & rsWordContext("词句ID").value, rsWordContext("名称").value, 1)
                    
                    objWord.Tag = Nvl(rsWordContext("内容文本").value, "")
                    rsWordContext.MoveNext
                Wend
            End If

            
        End If
        rsWordClass.MoveNext
    Loop

    
    Exit Sub
errHandle:
    If err.Number <> 35602 Then '35602表示键值重复
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End If
End Sub


Private Sub trvWord_DblClick()
    If Not (trvWord.SelectedItem Is Nothing) Then
        RaiseEvent OnWordDbClickEvent(CStr(trvWord.SelectedItem.Tag))
    End If
End Sub

Private Sub trvWord_Expand(ByVal Node As MSComctlLib.Node)
    trvWord.SelectedItem = Node
End Sub

Private Sub trvWord_NodeClick(ByVal Node As MSComctlLib.Node)
    txtWordContext.Text = CStr(Node.Tag)
End Sub








Private Sub UserControl_Initialize()
'    mlngWordHeight = Round(Height / 3) * 2

    Call InitFace
    
End Sub

Private Sub UserControl_InitProperties()
    mstrModuleName = ""
    mlngDepartId = ""
End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    mstrModuleName = PropBag.ReadProperty("ModuleName", "")
    mlngDepartId = PropBag.ReadProperty("CurDepartId", "")
'    WordHeight = PropBag.ReadProperty("WordHeight", Round(Height / 3) * 2)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next

    Call PropBag.WriteProperty("ModuleName", mstrModuleName, "")
    Call PropBag.WriteProperty("CurDepartId", mlngDepartId, "")
'    Call PropBag.WriteProperty("WordHeight", mlngWordHeight, Round(Height / 3) * 2)
End Sub

