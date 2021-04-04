VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmRunControl 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4245
      Index           =   1
      Left            =   45
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   6795
      Top             =   420
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
            Picture         =   "frmRunControl.frx":0000
            Key             =   "file"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunControl.frx":015A
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunControl.frx":69BC
            Key             =   "folder_open"
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Left            =   2655
      Top             =   555
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmRunControl.frx":D21E
   End
End
Attribute VB_Name = "frmRunControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'变量定义

Private Enum Command
    初始控件
    读注册表
    初始数据
    
    刷新目录数据
    刷新项目数据
    
    新增目录
    修改目录
    删除目录
    刷新指定目录
        
    新增信息
    修改信息
    删除信息
    内容配置
    刷新指定信息
    移除指定信息
End Enum

Private mstrBusiness As String
Private mlngModualCode As Long
Private mstrSQL As String
Private mblnStartUp As Boolean
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mstrCurrentGroup As String

Private WithEvents mfrmStation As frmStation

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
Private Function ExecuteCommand(ByVal enmCommand As Command, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim intRow As Integer
    Dim varTmp As Variant

    On Error GoTo errHand
    
    
'    Set mrsCondition = zlCommFun.CreateCondition
    
    Select Case enmCommand
    '------------------------------------------------------------------------------------------------------------------
    Case Command.初始控件

        Call InitDockPannel
        Call InitTaskPanel
    End Select
    
    
    GoTo EndHand

    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    Call zlComLib.SaveErrLog
    
    '------------------------------------------------------------------------------------------------------------------
EndHand:
End Function

Private Sub InitTaskPanel()
    
    Dim tplGroup As TaskPanelGroup
    Dim tplItem As TaskPanelGroupItem
    
    With tpl
        .SetIconSize 24, 24
        Call .Icons.AddIcons(imgMain.Icons)
        .VisualTheme = xtpTaskPanelThemeNativeWinXP
        .Behaviour = xtpTaskPanelBehaviourToolbox
        .ItemLayout = xtpTaskItemLayoutImagesWithTextBelow
        
        .SetMargins 5, 5, 5, 5, 5
        .SetItemInnerMargins 0, 5, 0, 5
        .SelectItemOnFocus = True
        
                
        Set tplGroup = .Groups.Add(0, "分组")
        tplGroup.Expandable = False
        tplGroup.CaptionVisible = False
        
        Set tplItem = tplGroup.Items.Add(1, "用户配置", xtpTaskItemTypeLink, 2621)
        tplItem.Tag = "G01"
        tplItem.Tooltip = "配置产品工作站对应的消息用户"
        
        mstrCurrentGroup = tplItem.Tag
        tplItem.Selected = True
        
        
'        Set tplItem = tplGroup.Items.Add(2, "在线模块", xtpTaskItemTypeLink, 2622)
'        tplItem.Tag = "G02"
'        tplItem.Tooltip = "用户自己定义的消息"
'
'        Set tplItem = tplGroup.Items.Add(3, "在线用户", xtpTaskItemTypeLink, 2622)
'        tplItem.Tag = "G03"
'        tplItem.Tooltip = "用户自己定义的消息"
        
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
    objPane.Title = "SQL"
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
            If mfrmStation Is Nothing Then Set mfrmStation = New frmStation
            Item.Handle = mfrmStation.hWnd
            Call mfrmStation.Execute
        Case "G02"
            
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
    mlngModualCode = 1003
    
    Call ExecuteCommand(Command.初始控件)
    Call ExecuteCommand(Command.读注册表)

    Call zlComLib.RestoreWinState(Me, App.ProductName)
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Call zlDataBase.ShowReportMenu(Me, ParamInfo.系统号, ParamInfo.模块号, UserInfo.模块权限)
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call zlCommFun.SetPaneRange(dkpMain, 1, 100, 15, 100, Me.ScaleHeight)
    Call zlCommFun.SetPaneRange(dkpMain, 2, 15, 200, Me.ScaleWidth, 500)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not (mfrmStation Is Nothing) Then
        Unload mfrmStation
        Set mfrmStation = Nothing
    End If
    
End Sub


Private Sub mfrmStation_AfterClose()
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
        If mfrmStation Is Nothing Then Set mfrmStation = New frmStation
        dkpMain.Panes(2).Handle = mfrmStation.hWnd
        Call mfrmStation.Execute
    Case "G02"
        
    End Select

End Sub
