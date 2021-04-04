VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompends 
   BorderStyle     =   0  'None
   Caption         =   "frmStucture"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   284
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.TreeView Tree 
      Height          =   1920
      Left            =   555
      TabIndex        =   0
      Top             =   285
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   3387
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   0
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "imlTreeIcons"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTreeIcons 
      Left            =   30
      Top             =   2445
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65382
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompends.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompends.frx":0116
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   90
      Top             =   675
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00808080&
      Height          =   375
      Left            =   450
      Top             =   180
      Width           =   330
   End
End
Attribute VB_Name = "frmCompends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmParent As frmMain    '系统主窗体
Private blnButton As Integer    '按下鼠标左右键的记录
Public Event NodeSelected(lngCompendID As Long)

Public Sub SetParent(Parent As Object)
    Set frmParent = Parent
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '菜单执行事件
    Dim lKey As Long
    Select Case Control.ID
    Case ID_EDIT_REFCOMPEND
        frmParent.Document.Compends.UpdateOrdersFromText frmParent.Editor1
        frmParent.Document.Compends.FillTree Tree
    Case ID_EDIT_ADDCOMPEND
        Dim f_Add As New frmInsCompend
        f_Add.ShowMe frmParent, frmParent.Editor1, frmParent.Document.Compends
    Case ID_EDIT_MODCOMPEND
        If Not Me.Tree.SelectedItem Is Nothing Then
             lKey = Me.Tree.SelectedItem.Tag
             If lKey > 0 Then
                If frmParent.Document.Compends("K" & lKey).预制提纲ID <> 0 And frmParent.Document.EditType <> cprET_病历文件定义 Then
                    MsgBox "不允许编辑保留提纲！", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                Dim f_Mod As New frmInsCompend
                f_Mod.ShowMe frmParent, frmParent.Editor1, frmParent.Document.Compends, frmParent.Document.Compends("K" & lKey)
             End If
        End If
    Case ID_EDIT_DELCOMPEND
        If Not Me.Tree.SelectedItem Is Nothing Then
             lKey = Me.Tree.SelectedItem.Tag
             If lKey > 0 Then
                frmParent.DeleteOutline lKey
             End If
        End If
    Case ID_EDIT_COMPENDWORD
        If Me.Tree.SelectedItem Is Nothing Then Exit Sub
        lKey = frmParent.Document.Compends("K" & Me.Tree.SelectedItem.Tag).ID
        If lKey = 0 Then MsgBox "需要先保存提纲才能关联词句示范！", vbInformation, gstrSysName: Exit Sub
        If frmCompendWord.ShowMe(frmParent, lKey, frmParent.Document.EPRFileInfo.种类) = True Then
            blnButton = vbLeftButton
            Call Tree_NodeClick(Me.Tree.SelectedItem)
        End If
    End Select
End Sub

Private Sub CommandBars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    Select Case Control.ID
    Case ID_EDIT_ADDCOMPEND, ID_EDIT_REFCOMPEND
        Control.Enabled = (frmParent.Editor1.ViewMode = cprNormal) And (frmParent.Editor1.AuditMode = False)
    Case ID_EDIT_DELCOMPEND
        If Tree.SelectedItem Is Nothing Then
            Control.Enabled = False
        Else
            Control.Enabled = (frmParent.Editor1.ViewMode = cprNormal) And (frmParent.Editor1.AuditMode = False)
        End If
    Case ID_EDIT_MODCOMPEND
        If Tree.SelectedItem Is Nothing Then
            Control.Enabled = False
        Else
            Control.Enabled = (frmParent.Editor1.ViewMode = cprNormal) And (frmParent.Editor1.AuditMode = False)
        End If
    Case ID_EDIT_COMPENDWORD
        Control.Enabled = Not (Tree.SelectedItem Is Nothing)
    End Select
    If frmParent.Document Is Nothing Then
        Control.Enabled = False
        Control.Visible = False
    Else
        If frmParent.Document.EditType <> cprET_病历文件定义 And Control.ID <> ID_EDIT_REFCOMPEND Then '只能在定义时修改提纲
            Control.Enabled = False
            Control.Visible = False
        End If
    End If
End Sub

Private Sub Form_Load()
    CommandBars.Icons = zlCommFun.GetPubIcons
    CommandBars.ActiveMenuBar.Visible = False
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Tree.Move 2, 2, Me.ScaleWidth - 4, Me.ScaleHeight - 4
    shpBorder.Move 1, 1, Tree.Width + 2, Tree.Height + 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set frmParent = Nothing
    imlTreeIcons.ListImages.Clear
End Sub

Private Sub Tree_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then KeyCode = 0
End Sub

Private Sub Tree_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim Node As MSComctlLib.Node
    If Me.Tree.Tag <> "NodeUnchecked" Then Exit Sub
    For Each Node In Me.Tree.Nodes
        If Node.Children > 0 Then Node.Checked = False
        If frmParent.Document.Compends("K" & Node.Tag).ID = 0 Then Node.Checked = False
        If frmParent.Document.Compends("K" & Node.Tag).定义提纲ID = 0 Then Node.Checked = False
    Next
    Me.Tree.Tag = ""
End Sub

Private Sub Tree_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnButton = Button
    Call Tree_KeyUp(vbKeySpace, 0)
    If Button <> vbRightButton Then Exit Sub
    
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    Set Popup = CommandBars.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_REFCOMPEND, "刷新提纲(&R)")
        Set Control = .Add(xtpControlButton, ID_EDIT_ADDCOMPEND, "新增提纲(&A)"): Control.BeginGroup = True
        Set Control = .Add(xtpControlButton, ID_EDIT_DELCOMPEND, "删除提纲(&D)")
        Set Control = .Add(xtpControlButton, ID_EDIT_MODCOMPEND, "修改提纲(&M)")
        Set Control = .Add(xtpControlButton, ID_EDIT_COMPENDWORD, "关联词句示范(&S)")
        Control.BeginGroup = True: Control.STYLE = xtpButtonCaption
    End With
    '定位到提纲末尾
    Dim lKey As Long, lS As Long, lE As Long
    If Not Me.Tree.SelectedItem Is Nothing Then
         lKey = Me.Tree.SelectedItem.Tag
         If lKey > 0 Then
            frmParent.Editor1.Tag = "TreeMenu"
            frmParent.Document.Compends("K" & lKey).GetPosition frmParent.Editor1, lS, lE
            If frmParent.Editor1.Range(lE - 2, lE) = vbCrLf And frmParent.Editor1.Range(lE - 2, lE).Font.Protected = False Then lE = lE - 2
            frmParent.Editor1.Range(lE, lE).Selected
            frmParent.Editor1.Tag = ""
         End If
    End If
    Popup.ShowPopup
End Sub

Private Sub Tree_NodeCheck(ByVal Node As MSComctlLib.Node)
    If Node.Children > 0 Then Me.Tree.Tag = "NodeUnchecked": Exit Sub
    If frmParent.Document.Compends("K" & Node.Tag).ID = 0 Then Me.Tree.Tag = "NodeUnchecked": Exit Sub
    If frmParent.Document.Compends("K" & Node.Tag).定义提纲ID = 0 Then Me.Tree.Tag = "NodeUnchecked": Exit Sub
End Sub

Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)
    '动态同步文本提纲
    If blnButton = vbRightButton Then
        Exit Sub
    Else
        If frmParent.Editor1.ViewMode = cprNormal Then frmParent.Document.Compends("K" & Node.Tag).GotoStartPos frmParent.Editor1
    End If
    '同步词句示范的显示
    RaiseEvent NodeSelected(frmParent.Document.Compends("K" & Node.Tag).预制提纲ID)
End Sub
