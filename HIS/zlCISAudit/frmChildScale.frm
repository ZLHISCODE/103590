VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmChildScale 
   Caption         =   "病案查看"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13080
   Icon            =   "frmChildScale.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   13080
   StartUpPosition =   2  '屏幕中心
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChildScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mfrmMain  As Object
Private mblnInist As Boolean '是否初始化,特殊处理

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errH
    Call RestoreWinState(Me, App.ProductName)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
       
    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
       Dim objPane As Pane
       Set objPane = dkpMain.CreatePane(1, 100, 200, DockLeftOf, Nothing): objPane.Title = "病案显示": objPane.Options = PaneNoCaption
       Call DockPannelInit(dkpMain)
    End Select
End Function
     
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = mfrmMain.hWnd
    End Select
End Sub

Public Function zlInitData(ByVal frmMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If mblnInist = False Then
        Set mfrmMain = frmMain
        zlInitData = ExecuteCommand("初始控件")
        mblnInist = True
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    Cancel = 1
    On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    Exit Sub
'    Unload mfrmMain
'    Set mfrmMain = Nothing
End Sub
