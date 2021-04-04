VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmTest 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmTest.frx":0000
      Left            =   15
      Top             =   690
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjMipModule As Object


Private mfrmTestSend As frmTestSend
Private mfrmTestReceive As frmTestReceive

Private mblnStartup As Boolean

Public Function Initialize() As Boolean
    
    Initialize = True
    
End Function

Public Sub ActiveForm()
    
    '启动消息服务平台客户端收发服务
    '------------------------------------------------------------------------------------------------------------------
    If gobjComLib.ConnectMip(Me.hWnd) = True Then
        Set mobjMipModule = CreateObject("zl9ComLib.clsMipModule")
        Call mobjMipModule.InitMessage(0, 0, "")
        Call gobjComLib.AddMipModule(mobjMipModule)
    Else
        MsgBox Err.Description
    End If
        
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case 1
        Set mfrmTestSend = New frmTestSend
        Item.Handle = mfrmTestSend.hWnd
    Case 2
        Set mfrmTestReceive = New frmTestReceive
        Item.Handle = mfrmTestReceive.hWnd
    End Select
End Sub

'Private Sub Form_Activate()
'    If mblnStartup = False Then Exit Sub
'    mblnStartup = False
'
'    '启动消息服务平台客户端收发服务
'    '------------------------------------------------------------------------------------------------------------------
'    If gobjComLib.ConnectMip(Me.hWnd) = True Then
'        Set mobjMipModule = CreateObject("zl9ComLib.clsMipModule")
'        Call mobjMipModule.InitMessage(0, 0, "")
'        Call gobjComLib.AddMipModule(mobjMipModule)
'    End If
'End Sub

Private Sub Form_Load()
    
    mblnStartup = True
    
    Call InitCommandBar
    Call InitDockPannel
    
End Sub

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objExtendedBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim strList As String
    Dim strListName() As String
    Dim i As Long
    Dim blnChck As Boolean
    Dim strTmp As String
    
    '初始设置
    '------------------------------------------------------------------------------------------------------------------
    Call CommandBarInit(cbsMain)

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
        
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 500, 100, DockTopOf, Nothing)
    objPane.Title = "发送"
    objPane.Options = PaneNoCaption

    Set objPane = dkpMain.CreatePane(2, 500, 100, DockBottomOf, objPane)
    objPane.Title = "接收"
    objPane.Options = PaneNoCaption


    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMain, 1, 15, 350, Me.ScaleWidth, 350)
        
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    If Not (mfrmTestSend Is Nothing) Then
        Unload mfrmTestSend
        Set mfrmTestSend = Nothing
    End If
    
    If Not (mfrmTestReceive Is Nothing) Then
        Unload mfrmTestReceive
        Set mfrmTestReceive = Nothing
    End If
    
    If Not (mobjMipModule Is Nothing) Then
        mobjMipModule.CloseMessage
        Call gobjComLib.DelMipModule(mobjMipModule)
        Set mobjMipModule = Nothing
    End If
    
    Call gobjComLib.DisConnectMip
    
End Sub

