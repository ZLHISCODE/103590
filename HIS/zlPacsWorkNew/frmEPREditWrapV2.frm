VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmEPREditWrapV2 
   BorderStyle     =   0  'None
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmEPREditWrapV2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   480
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmEPREditWrapV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private WithEvents mobjReport As zlRichEPR.cDockReport
Attribute mobjReport.VB_VarHelpID = -1
  

Private mlngModuleNo As Long
Private mlngDeptID As Long

Private mobjStudyInfo As clsStudyInfo

Private mObjNotify As IEventNotify
Private mobjCaller As frmReportV2
 
 
Public Function InitEprEditor( _
    objCaller As frmReportV2, objNotify As IEventNotify, _
    ByVal lngModuleNo As Long, ByVal lngDeptId As Long) As Boolean
'初始化
    InitEprEditor = False
    
    Me.Hide
    
    mlngModuleNo = lngModuleNo
    mlngDeptID = lngDeptId
    
    Set mobjCaller = objCaller
    Set mObjNotify = objNotify
    
    If gobjRichEPR Is Nothing Then
        Set gobjRichEPR = New zlRichEPR.cRichEPR
        Call gobjRichEPR.InitRichEPR(gcnOracle, Me, 100, False)
    End If
    
    If gobjReport Is Nothing Then
        Set gobjReport = New zlRichEPR.cDockReport
    End If
    
    If gobjReport Is Nothing Then
        MsgBox "病历部件[zlRichEPR]创建失败，不能使用病历方式进行编辑。", vbOKOnly, "提示"
        Unload Me
        Exit Function
    End If
   
    Set mobjReport = gobjReport
    
    InitEprEditor = True
End Function


Public Sub OpenEprEditor(objStudyInfo As clsStudyInfo, ByVal blnIsAuditing As Boolean)
'打开epr编辑器
    Dim cbrControl As CommandBarControl
    
    Set mobjStudyInfo = objStudyInfo
    
    If blnIsAuditing Then
        Set cbrControl = cbrMain.FindControl(, conMenu_Edit_Audit, False)
        If cbrControl Is Nothing Then
            Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_Edit_Audit, "病历修订"): cbrControl.ID = conMenu_Edit_Audit
        End If
    Else
        Set cbrControl = cbrMain.FindControl(, conMenu_Edit_Modify, False)
        If cbrControl Is Nothing Then
            Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "病历编辑"): cbrControl.ID = conMenu_Edit_Modify
        End If
    End If
    
    mobjReport.zlRefresh 0, 0, , , mobjStudyInfo.blnCanPrint, mlngModuleNo
    
    Call mobjReport.zlRefresh(mobjStudyInfo.lngAdviceId, mlngDeptID, , , mobjStudyInfo.blnCanPrint, mlngModuleNo)     ' True, blnIsMoved, True,
    Call mobjReport.zlExecuteCommandBars(cbrControl)
End Sub

Public Sub ExecuteMenu(objStudyInfo As clsStudyInfo, ByVal lngControlID As Long)
'执行菜单
    Dim cbrControl As CommandBarControl
 
    Set mobjStudyInfo = objStudyInfo
    
    Set cbrControl = cbrMain.FindControl(, lngControlID, False)
    If cbrControl Is Nothing Then
        Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlButton, lngControlID, lngControlID): cbrControl.ID = lngControlID
    End If
    
    mobjReport.zlRefresh 0, 0, , , , mlngModuleNo
    
    Call mobjReport.zlRefresh(mobjStudyInfo.lngAdviceId, mlngDeptID, , , , mlngModuleNo)     ' True, blnIsMoved, True,
    Call mobjReport.zlExecuteCommandBars(cbrControl)
End Sub
 

Private Sub Form_Unload(Cancel As Integer)
    Set mObjNotify = Nothing
    Set mobjReport = Nothing
    
    Set mobjStudyInfo = Nothing
End Sub

Private Sub SyncReport()
'同步报告内容
    If mobjCaller Is Nothing Then Exit Sub
    
    If mobjCaller.AdviceId = mobjStudyInfo.lngAdviceId Then
       Call mobjCaller.zlRefresh(mobjStudyInfo)
    End If
End Sub

Private Sub mobjReport_AfterClosed(ByVal lngOrderID As Long)
On Error GoTo errhandle
    If mobjCaller Is Nothing Then
        Call mObjNotify.Broadcast(BM_REPORT_EVENT_CLOSEEPR, 1, lngOrderID)
    End If
    
    Unload Me
Exit Sub
errhandle:
    MsgBoxD Me, err.Description, vbOKOnly, "提示"
End Sub

Private Sub mobjReport_AfterDeleted(ByVal lngOrderID As Long)
'发送报告删除消息
    Call mObjNotify.SendRequest(WM_LIST_SYNCROW, 1, lngOrderID)
End Sub

Private Sub mobjReport_AfterPrinted(ByVal lngOrderID As Long)
'病历编辑器打印后事件

    '发送打印通知
    Call mObjNotify.Broadcast(BM_REPORT_EVENT_PRINT, 1, lngOrderID, 0, -1) '-1表示病历编辑器
End Sub

Private Sub mobjReport_AfterSaved(ByVal lngOrderID As Long, ByVal lngSaveType As Long)
'lngSaveType：0-普通保存，1-诊断签名，2-审核签名
'注：回退操作会触发此保存事件
On Error GoTo errhandle
    Call SyncReport
    
    Select Case lngSaveType
        Case 0
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_SAVE, 1, lngOrderID, 0, -1) '-1表示病历编辑器
            
        Case 1
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_SIGN, 1, lngOrderID, 0, -1)
            
        Case 2
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_AUDIT, 1, lngOrderID, 0, -1)
            
    End Select
Exit Sub
errhandle:
    
End Sub
