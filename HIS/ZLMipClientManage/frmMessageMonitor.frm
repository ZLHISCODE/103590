VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmMessageMonitor 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4245
      Index           =   1
      Left            =   420
      ScaleHeight     =   4245
      ScaleWidth      =   2370
      TabIndex        =   0
      Top             =   1020
      Width           =   2370
      Begin XtremeSuiteControls.TaskPanel tpl 
         Height          =   4770
         Left            =   345
         TabIndex        =   1
         Top             =   495
         Width           =   3210
         _Version        =   589884
         _ExtentX        =   5662
         _ExtentY        =   8414
         _StockProps     =   64
         Behaviour       =   1
         ItemLayout      =   2
         HotTrackStyle   =   3
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   375
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMessageMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'变量定义
Private mstrBusiness As String
Private mlngModualCode As Long
Private mstrSQL As String
Private mblnStartUp As Boolean
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mstrCurrentGroup As String
Private WithEvents mfrmSendLog As frmSendLog
Attribute mfrmSendLog.VB_VarHelpID = -1
Private WithEvents mfrmReceiveLog As frmReceiveLog
Attribute mfrmReceiveLog.VB_VarHelpID = -1

Public Event AfterClose(ByVal lngModual As Long)
Public Event AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)

'######################################################################################################################
'接口方法
Public Function ShowForm()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Call Form_Activate
End Function

'######################################################################################################################
Private Sub InitTaskPanel()
    
    Dim tplGroup As TaskPanelGroup
    Dim tplItem As TaskPanelGroupItem
    
    With tpl
        .SetIconSize 24, 24
        Call .Icons.AddIcons(zlCommFun.GetPubIcons)
        .VisualTheme = xtpTaskPanelThemeNativeWinXP
        .Behaviour = xtpTaskPanelBehaviourToolbox
        .ItemLayout = xtpTaskItemLayoutImagesWithTextBelow
        
        .SetMargins 5, 5, 5, 5, 5
        .SetItemInnerMargins 0, 5, 0, 5
        .SelectItemOnFocus = True
        
        Set tplGroup = .Groups.Add(0, "分组")
        tplGroup.Expandable = False
        tplGroup.CaptionVisible = False
        
        Set tplItem = tplGroup.Items.Add(1, "发送消息", xtpTaskItemTypeLink, enumIcon.Message_Send + 1)
        tplItem.Tag = "G01"
        tplItem.Tooltip = "系统固定配置的消息"
        
        tplItem.Selected = True
        mstrCurrentGroup = tplItem.Tag
        
        Set tplItem = tplGroup.Items.Add(2, "接收消息", xtpTaskItemTypeLink, enumIcon.Message_Receive + 1)
        tplItem.Tag = "G02"
        tplItem.Tooltip = "用户自己定义的消息"
        
        .Reposition
    
    End With
    
    Exit Sub

errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    Call zlComLib.SaveErrLog
End Sub

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "分组"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 300, 100, DockRightOf, objPane)
    objPane.Title = "内容"
    objPane.Options = PaneNoCaption
        
'    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Select Case Pane
    Case dkpMain.Panes(1)
        Select Case Action
        Case PaneActionPinned, PaneActionPinning, PaneActionExpanded, PaneActionExpanding, PaneActionCollapsed, PaneActionCollapsing
            Cancel = False
        Case Else
            Cancel = True
        End Select
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(1).hWnd
    Case 2
        Select Case mstrCurrentGroup
        Case "G01"
            If mfrmSendLog Is Nothing Then Set mfrmSendLog = New frmSendLog
            Item.Handle = mfrmSendLog.hWnd
        Case "G02"
            If mfrmReceiveLog Is Nothing Then Set mfrmReceiveLog = New frmReceiveLog
            Item.Handle = mfrmReceiveLog.hWnd
        End Select
    End Select
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    DoEvents
    mblnStartUp = False
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mlngModualCode = 1004
    
    Call InitDockPannel
    Call InitTaskPanel
    
    Call zlComLib.RestoreWinState(Me, App.ProductName)
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call zlCommFun.SetPaneRange(dkpMain, 1, 100, 15, 100, Me.ScaleHeight)
    Call zlCommFun.SetPaneRange(dkpMain, 2, 15, 200, Me.ScaleWidth, 500)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    If Not (mfrmSendLog Is Nothing) Then
        Unload mfrmSendLog
        Set mfrmSendLog = Nothing
    End If
    
    If Not (mfrmReceiveLog Is Nothing) Then
        Unload mfrmReceiveLog
        Set mfrmReceiveLog = Nothing
    End If
    
End Sub

Private Sub mfrmReceiveLog_AfterClose()
    Unload Me
    RaiseEvent AfterClose(mlngModualCode)
End Sub

Private Sub mfrmSendLog_AfterClose()
    Unload Me
    RaiseEvent AfterClose(mlngModualCode)
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 1
        tpl.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    End Select
End Sub

Private Sub tpl_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    mstrCurrentGroup = Item.Tag

    Select Case mstrCurrentGroup
    Case "G01"
        If mfrmSendLog Is Nothing Then Set mfrmSendLog = New frmSendLog
        dkpMain.Panes(2).Handle = mfrmSendLog.hWnd
    Case "G02"
        If mfrmReceiveLog Is Nothing Then Set mfrmReceiveLog = New frmReceiveLog
        dkpMain.Panes(2).Handle = mfrmReceiveLog.hWnd
    End Select

End Sub


