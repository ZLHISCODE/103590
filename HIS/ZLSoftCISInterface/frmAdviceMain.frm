VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdviceMain 
   BorderStyle     =   0  'None
   Caption         =   "中联软件"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Icon            =   "frmAdviceMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   360
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAdviceMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjForm As Object '医嘱信息窗体
Private mobjAdvice As Object '医嘱类,zlPublicAdvice.clsDockInAdvices
Private madvice As Object 'zlPublicAdvice.clsPublicAdvice
Private mclsReport As Object

Private Sub Func功能()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objB As Object

    On Error GoTo errH
  
    If madvice Is Nothing Then
        Set madvice = CreateObject("zlPublicAdvice.clsPublicAdvice")
    End If
    Set mclsReport = CreateObject("zl9Report.clsReport")
    mclsReport.InitOracle gcnOracle
    Call madvice.InitCommon(gcnOracle, 100)
 
    If mobjAdvice Is Nothing Then
        Set mobjAdvice = CreateObject("zlPublicAdvice.clsDockInAdvices")
        Set mobjForm = mobjAdvice.zlGetForm
        CommandBarsGlobalSettings.App = App
        CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
        CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
        cbsMain.VisualTheme = xtpThemeOffice2003
        With Me.cbsMain.Options
            .ShowExpandButtonAlways = False
            .ToolBarAccelTips = True
            .AlwaysShowFullMenus = False
            .IconsWithShadow = True
            .UseDisabledIcons = True
            .LargeIcons = True
            .SetIconSize True, 24, 24
            .SetIconSize False, 16, 16
        End With
        cbsMain.EnableCustomization False
        cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 1, "文件(&F)", -1, False).Id = 1 '固有
        cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 2, "医嘱(&A)", -1, False).Id = 2 '固有
        cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 8, "工具(&T)", -1, False).Id = 8 '固有
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 7, "查看(&V)", -1, False) '固有
        Call objMenu.CommandBar.Controls.Add(xtpControlButton, 791, "刷新(&R)")
        objMenu.Id = 7
        objMenu.CommandBar.Controls.Add(xtpControlButton, 702, "状态栏(&S)").Id = 702 '固有
        Set objBar = cbsMain.Add("工具栏", xtpBarTop)
        mobjAdvice.zlDefCommandBars Me, cbsMain, 0, False, Nothing, False
        mobjAdvice.zlRefresh glng病人ID, glng主页ID, glng病区ID, glng科室ID, 0
        Set objB = cbsMain.FindControl(, glngFunID, , True)
        mobjAdvice.zlExecuteCommandBars objB
    End If
    Exit Sub
errH:
    MsgBox err.Description, "中联软件"
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    mobjAdvice.zlExecuteCommandBars Control
End Sub

Private Sub Form_Activate()
    Call Func功能
End Sub

Public Function zlCloseMe()
    Unload Me
End Function

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    madvice.CloseWindows
    Set mobjForm = Nothing
    Set mobjAdvice = Nothing
    Set madvice = Nothing
    Set mclsReport = Nothing
End Sub
