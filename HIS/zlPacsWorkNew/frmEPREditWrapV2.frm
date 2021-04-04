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
   StartUpPosition =   3  '����ȱʡ
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
'��ʼ��
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
        MsgBox "��������[zlRichEPR]����ʧ�ܣ�����ʹ�ò�����ʽ���б༭��", vbOKOnly, "��ʾ"
        Unload Me
        Exit Function
    End If
   
    Set mobjReport = gobjReport
    
    InitEprEditor = True
End Function


Public Sub OpenEprEditor(objStudyInfo As clsStudyInfo, ByVal blnIsAuditing As Boolean)
'��epr�༭��
    Dim cbrControl As CommandBarControl
    
    Set mobjStudyInfo = objStudyInfo
    
    If blnIsAuditing Then
        Set cbrControl = cbrMain.FindControl(, conMenu_Edit_Audit, False)
        If cbrControl Is Nothing Then
            Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_Edit_Audit, "�����޶�"): cbrControl.ID = conMenu_Edit_Audit
        End If
    Else
        Set cbrControl = cbrMain.FindControl(, conMenu_Edit_Modify, False)
        If cbrControl Is Nothing Then
            Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "�����༭"): cbrControl.ID = conMenu_Edit_Modify
        End If
    End If
    
    mobjReport.zlRefresh 0, 0, , , mobjStudyInfo.blnCanPrint, mlngModuleNo
    
    Call mobjReport.zlRefresh(mobjStudyInfo.lngAdviceId, mlngDeptID, , , mobjStudyInfo.blnCanPrint, mlngModuleNo)     ' True, blnIsMoved, True,
    Call mobjReport.zlExecuteCommandBars(cbrControl)
End Sub

Public Sub ExecuteMenu(objStudyInfo As clsStudyInfo, ByVal lngControlID As Long)
'ִ�в˵�
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
'ͬ����������
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
    MsgBoxD Me, err.Description, vbOKOnly, "��ʾ"
End Sub

Private Sub mobjReport_AfterDeleted(ByVal lngOrderID As Long)
'���ͱ���ɾ����Ϣ
    Call mObjNotify.SendRequest(WM_LIST_SYNCROW, 1, lngOrderID)
End Sub

Private Sub mobjReport_AfterPrinted(ByVal lngOrderID As Long)
'�����༭����ӡ���¼�

    '���ʹ�ӡ֪ͨ
    Call mObjNotify.Broadcast(BM_REPORT_EVENT_PRINT, 1, lngOrderID, 0, -1) '-1��ʾ�����༭��
End Sub

Private Sub mobjReport_AfterSaved(ByVal lngOrderID As Long, ByVal lngSaveType As Long)
'lngSaveType��0-��ͨ���棬1-���ǩ����2-���ǩ��
'ע�����˲����ᴥ���˱����¼�
On Error GoTo errhandle
    Call SyncReport
    
    Select Case lngSaveType
        Case 0
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_SAVE, 1, lngOrderID, 0, -1) '-1��ʾ�����༭��
            
        Case 1
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_SIGN, 1, lngOrderID, 0, -1)
            
        Case 2
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_AUDIT, 1, lngOrderID, 0, -1)
            
    End Select
Exit Sub
errhandle:
    
End Sub
