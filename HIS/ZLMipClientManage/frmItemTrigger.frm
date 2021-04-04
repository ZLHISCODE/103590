VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmItemTrigger 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2160
      Index           =   0
      Left            =   75
      ScaleHeight     =   2160
      ScaleWidth      =   5670
      TabIndex        =   0
      Top             =   705
      Width           =   5670
      Begin RichTextLib.RichTextBox txtSQL 
         Height          =   1590
         Left            =   225
         TabIndex        =   1
         Top             =   255
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2805
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   1
         Appearance      =   0
         TextRTF         =   $"frmItemTrigger.frx":0000
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmItemTrigger.frx":009D
      Left            =   690
      Top             =   180
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmItemTrigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngModualCode As Long
Private mstrPrivs As String
Private mstrSQL As String

Private mblnStartUp As Boolean
Private mlngTmp As Long
Private mblnShowAll As Boolean
Private mblnShowStop As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mblnDataChanged As Boolean
Private mblnNew As Boolean

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
            
    InitGrid = True
    
End Function

Public Function RefreshData(ByVal strDataKey As String) As Boolean

    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim rsTmp As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "id", strDataKey)

    Set rsTmp = gclsBusiness.ItemRead("id", rsCondition)
    If rsTmp.BOF = False Then
        txtSQL.Text = zlCommFun.NVL(rsTmp("trigger_condition").Value)
    End If
        
    RefreshData = True
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 300, DockTopOf, Nothing)
    objPane.Title = "SQL"
    objPane.Options = PaneNoCaption
        
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objFindKey As CommandBarControl
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call zlCommFun.CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeWhidbey
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份
    
    
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
'    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Adjust, "配置", , , xtpButtonIconAndCaption)
'    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 0, "    ", , , xtpButtonIconAndCaption)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, conMenu_View_LocationItem, "满足如下条件,SQL执行有记录或无SQL表示成立,无记录表示不成立", , , xtpButtonIconAndCaption)
    
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    End Select
End Sub

Private Sub Form_Load()
    Call InitDockPannel
    Call InitCommandBar
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        
    Call zlCommFun.SetPaneRange(dkpMain, 2, 15, 75, Me.ScaleWidth, 135)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjFindKey = Nothing
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        txtSQL.Move 0, 15, picPane(Index).Width, picPane(Index).Height - 15
    End Select
End Sub


Private Sub txtSQL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtSQL.Locked Then
        glngTXTProc = GetWindowLong(txtSQL.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtSQL.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtSQL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtSQL.Locked Then
        Call SetWindowLong(txtSQL.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub


