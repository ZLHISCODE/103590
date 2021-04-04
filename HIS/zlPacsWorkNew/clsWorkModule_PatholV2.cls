VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkModule_PatholV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IWorkMenuV2


Private Const M_STR_HINT_NoSelectData As String = "��Ч�ļ�����ݣ���ѡ����Ҫִ�еļ���¼��"
Private Const M_STR_MODULE_MENU_TAG As String = "����"

Public Enum TPatholModuleType
    pmtAll = -1
    pmtSpecimen = 0  '�걾����ģ��
    pmtMaterial = 1  '�걾ȡ��ģ��
    pmtSlices = 2    '������Ƭģ��
    pmtSpeExam = 3   '�����ؼ�ģ��
    pmtProRep = 4    '��������ģ��
End Enum


Private mlngModule  As Long
Private mstrPrivs  As String
Private mlngCurDeptId As Long
Private mObjNotify As IEventNotify

Private mobjStudyInfo As clsStudyInfo
Private mblnReadOnly As Boolean

Private mfrmPatholSpecimen As frmPatholSpecimen              '�걾����
Attribute mfrmPatholSpecimen.VB_VarHelpID = -1
Private mfrmPatholMaterial As frmPatholMaterials             'ȡ��
Attribute mfrmPatholMaterial.VB_VarHelpID = -1
Private mfrmPatholSlices As frmPatholSlices                  '��Ƭ
Attribute mfrmPatholSlices.VB_VarHelpID = -1
Private mfrmPatholSpeExam As frmPatholSpecialExamined        '�ؼ�
Attribute mfrmPatholSpeExam.VB_VarHelpID = -1
Private mfrmPatholProRep As frmPatholProcedureRep                       '���̱���

Private aryRegModule() As Object                                        '���浱ǰ������ģ��

    
Private mfrmConRequest As frmPatholConsultation                           '��������
Private mfrmPatholDecalinTask As frmPatholDecalcification            '�Ѹ�����

Private mObjActiveMenuBar As CommandBars

'��ǰ����ID
Property Get DeptId() As Long
    DeptId = mlngCurDeptId
End Property

'��ȡ�˵��ӿڶ���
Property Get zlMenu() As IWorkMenuV2
    Set zlMenu = Me
End Property

'��ǰҽ��ID
Property Get AdviceId() As Long
    AdviceId = mobjStudyInfo.lngAdviceId
End Property


Private Sub Class_Initialize()
    Set mfrmPatholDecalinTask = New frmPatholDecalcification
    Set mfrmConRequest = New frmPatholConsultation
    
    '����ʹ��uboundʱ���������Ĭ�ϳ�ʼ��һ��������Ԫ��
    ReDim Preserve aryRegModule(0)
    Set aryRegModule(0) = Nothing
End Sub

Public Sub Free()
On Error GoTo errhandle
    Dim i As Long
    
    If SafeArrayGetDim(aryRegModule) <> 0 Then
        For i = LBound(aryRegModule) To UBound(aryRegModule)
            If Not aryRegModule(i) Is Nothing Then
                Unload aryRegModule(i)
                Set aryRegModule(i) = Nothing
            End If
        Next i
    End If
    
    Set mfrmPatholSpecimen = Nothing
    Set mfrmPatholMaterial = Nothing
    Set mfrmPatholSlices = Nothing
    Set mfrmPatholSpeExam = Nothing
    Set mfrmPatholProRep = Nothing
    
    Unload mfrmConRequest
    Set mfrmConRequest = Nothing
    
    Unload mfrmPatholDecalinTask
    Set mfrmPatholDecalinTask = Nothing
    
errhandle:
End Sub


Public Function GetModule(ByVal strModuleName As String, Optional ByVal blnAutoCreate As Boolean = True) As Object
'��ȡָ��ģ�����
    Set GetModule = Nothing
    
    Select Case strModuleName
        Case C_TAB_NAME_�걾����
            If mfrmPatholSpecimen Is Nothing And blnAutoCreate Then
                Set mfrmPatholSpecimen = New frmPatholSpecimen

                Call mfrmPatholSpecimen.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mObjNotify.Owner)
                
'                FormSetCaption mfrmPatholSpecimen, False, False
                Call SetWindowStyle(mfrmPatholSpecimen.hwnd, False)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholSpecimen
            End If
            
            Set GetModule = mfrmPatholSpecimen
            
        Case C_TAB_NAME_����ȡ��
            If mfrmPatholMaterial Is Nothing And blnAutoCreate Then
                Set mfrmPatholMaterial = New frmPatholMaterials
                
                Call mfrmPatholMaterial.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mObjNotify.Owner)
                
                Call SetWindowStyle(mfrmPatholMaterial.hwnd, False)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholMaterial
            End If
            
            Set GetModule = mfrmPatholMaterial
               
        Case C_TAB_NAME_������Ƭ
            If mfrmPatholSlices Is Nothing And blnAutoCreate Then
                Set mfrmPatholSlices = New frmPatholSlices

                Call mfrmPatholSlices.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mObjNotify.Owner)
                
                Call SetWindowStyle(mfrmPatholSlices.hwnd, False)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholSlices
            End If
            
            Set GetModule = mfrmPatholSlices
            
        Case C_TAB_NAME_�����ؼ�
            If mfrmPatholSpeExam Is Nothing And blnAutoCreate Then
                Set mfrmPatholSpeExam = New frmPatholSpecialExamined

                Call mfrmPatholSpeExam.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mObjNotify.Owner)
                
                Call SetWindowStyle(mfrmPatholSpeExam.hwnd, False)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholSpeExam
            End If
            
            Set GetModule = mfrmPatholSpeExam
            
        Case C_TAB_NAME_���̱���
            If mfrmPatholProRep Is Nothing And blnAutoCreate Then
                Set mfrmPatholProRep = New frmPatholProcedureRep

                Call mfrmPatholProRep.zlInitModule(mlngModule, mstrPrivs, mlngCurDeptId, mObjNotify.Owner)
                
                Call SetWindowStyle(mfrmPatholProRep.hwnd, False)
                
                ReDim Preserve aryRegModule(UBound(aryRegModule) + 1)
                Set aryRegModule(UBound(aryRegModule)) = mfrmPatholProRep
            End If
            
            Set GetModule = mfrmPatholProRep
    End Select
End Function


Private Sub Class_Terminate()
    Set mfrmPatholSpecimen = Nothing
    Set mfrmPatholMaterial = Nothing
    Set mfrmPatholSlices = Nothing
    Set mfrmPatholSpeExam = Nothing
    Set mfrmPatholProRep = Nothing
    
    ReDim aryRegModule(0)
    Set aryRegModule(0) = Nothing
    
    Set mfrmConRequest = Nothing
    Set mfrmPatholDecalinTask = Nothing
    Set mObjActiveMenuBar = Nothing
    Set mObjNotify = Nothing
    Set mobjStudyInfo = Nothing
End Sub

'�ӿ�ʵ�ֲ���*********************************************************************************

Public Function IWorkMenuV2_zlBaseMenuID() As Long

End Function

Public Function IWorkMenuV2_zlExecuteCmd(ByVal lngCmdType As Long)
'ִ�в˵�����

End Function

Public Function IWorkMenuV2_zlIsModuleMenu(ByVal strModuleName As String, objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'�жϲ˵��Ƿ����ڸ�ģ��˵�
    Dim blnIsModuleMenu As Boolean
    Dim i As Long
    Dim objIWork As IWorkMenuV2
    
    blnIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
    
'    If blnIsModuleMenu = False Then
'        Select Case strModuleName
'            Case "�걾����"
'                If Not mfrmPatholSpecimen Is Nothing Then blnIsModuleMenu = mfrmPatholSpecimen.zlMenu.zlIsModuleMenu(objControlMenu)
'            Case "����ȡ��"
'                If Not mfrmPatholMaterial Is Nothing Then blnIsModuleMenu = mfrmPatholMaterial.zlMenu.zlIsModuleMenu(objControlMenu)
'            Case "������Ƭ"
'                If Not mfrmPatholSlices Is Nothing Then blnIsModuleMenu = mfrmPatholSlices.zlMenu.zlIsModuleMenu(objControlMenu)
'            Case "�����ؼ�"
'                If Not mfrmPatholSpeExam Is Nothing Then blnIsModuleMenu = mfrmPatholSpeExam.zlMenu.zlIsModuleMenu(objControlMenu)
'            Case "���̱���"
'                If Not mfrmPatholProRep Is Nothing Then blnIsModuleMenu = mfrmPatholProRep.zlMenu.zlIsModuleMenu(objControlMenu)
'        End Select
'    End If
    
    IWorkMenuV2_zlIsModuleMenu = blnIsModuleMenu
End Function


Public Sub IWorkMenuV2_zlCreateMenu(ByVal strModuleName As String, objMenuBar As Object)
'����Ӱ���¼��Ӧ�Ĳ˵�
    Set mObjActiveMenuBar = objMenuBar
    
    If Len(strModuleName) <= 0 Then
        Call CreatePublicPatholMenu(objMenuBar)
    Else
        Call CreatePrivatePatholMenu(strModuleName, objMenuBar)
    End If
End Sub

Private Sub ClearPrivaePatholMenu(ByVal strModuleName As String, objMenuBar As Object)
    Dim objControlMenu As CommandBarControl
    
    '����֮ǰ�����Ĳ˵�
    If strModuleName = "�걾����" Then
        Set objControlMenu = objMenuBar.FindControl(, conMenu_PatholSpecimen)
        If Not objControlMenu Is Nothing Then objControlMenu.Delete
    End If
    
    If strModuleName = "����ȡ��" Then
        Set objControlMenu = objMenuBar.FindControl(, conMenu_PatholMaterial)
        If Not objControlMenu Is Nothing Then objControlMenu.Delete
    End If
    
    If strModuleName = "������Ƭ" Then
        Set objControlMenu = objMenuBar.FindControl(, conMenu_PatholSlices)
        If Not objControlMenu Is Nothing Then objControlMenu.Delete
    End If
    
    If strModuleName = "�����ؼ�" Then
        Set objControlMenu = objMenuBar.FindControl(, conMenu_PatholSpeExam)
        If Not objControlMenu Is Nothing Then objControlMenu.Delete
    End If
    
    If strModuleName = "���̱���" Then
        Set objControlMenu = objMenuBar.FindControl(, conMenu_PatholProRep)
        If Not objControlMenu Is Nothing Then objControlMenu.Delete
    End If
End Sub

Private Sub CreatePrivatePatholMenu(ByVal strModuleName As String, objMenuBar As Object)
'������ģ���Ӧ�Ĳ����˵�
    Call ClearPrivaePatholMenu(strModuleName, objMenuBar)
    
    Select Case strModuleName
        Case "�걾����"
            If Not mfrmPatholSpecimen Is Nothing Then Call mfrmPatholSpecimen.zlMenu.zlCreateMenu(objMenuBar)
        Case "����ȡ��"
            If Not mfrmPatholMaterial Is Nothing Then Call mfrmPatholMaterial.zlMenu.zlCreateMenu(objMenuBar)
        Case "������Ƭ"
            If Not mfrmPatholSlices Is Nothing Then Call mfrmPatholSlices.zlMenu.zlCreateMenu(objMenuBar)
        Case "�����ؼ�"
            If Not mfrmPatholSpeExam Is Nothing Then Call mfrmPatholSpeExam.zlMenu.zlCreateMenu(objMenuBar)
        Case "���̱���"
            If Not mfrmPatholProRep Is Nothing Then Call mfrmPatholProRep.zlMenu.zlCreateMenu(objMenuBar)
    End Select
End Sub


Private Sub CreatePublicPatholMenu(objMenuBar As Object)
'�������������˵�
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim objIWork As IWorkMenuV2
    
    Dim str3DFuncs() As String
    Dim i As Long
    Dim lng3DFunc As Long
    
    
    If Not HasMenu(objMenuBar, conMenu_PatholManage) Then
        Set cbrMenuBar = objMenuBar.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PatholManage, "����", 3, False)
        cbrMenuBar.ID = conMenu_PatholManage
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_Antibody_Manage, "�������", "", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_MealManage, "�ײ�ά��", "", 0, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_Decalin_Task, "�Ѹ�����", "", 0, True)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_Request, "��������", "", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_ReportDelay, "�ӳٵǼ�", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_ConRequest, "��������", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_ConFeedback, "���ﷴ��", "", 0, False)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_BatSlicesAccept, "������Ƭ", "", G_INT_ICONID_SLICES, True)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_BatSpeExamAccept, "�����ؼ�", "", G_INT_ICONID_SPEEXAM, True)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_Quality_Manage, "��������", "", 0, True)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PatholSlices_Quality, "��Ƭ����", "", 3097, True)
            
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_NumConfig, "��������", "", 9023, True)
        End With
    End If
End Sub

Private Sub CreatePublicPatholTools(objToolBar As Object)
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrLogOut As CommandBarControl
    Dim lngIndex As Long
        
  
    If HasMenu(objToolBar, conMenu_Pathol_BatProcess) Then Exit Sub
    
    Set cbrLogOut = objToolBar.FindControl(, conMenu_Manage_InQueue)
    
    lngIndex = 4
    If Not cbrLogOut Is Nothing Then lngIndex = cbrLogOut.Index
    
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Pathol_BatProcess, "��������", "����������Ƭ���ܡ���Ƭȷ�ϡ��ؼ���ܵȲ���", G_INT_ICONID_BATPROCESS, True, lngIndex + 1)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Pathol_BatSlicesAccept, "��Ƭ��������", "��Ƭ��������", G_INT_ICONID_SLICES, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Pathol_BatSpeExamAccept, "�ؼ���������", "�ؼ���������", G_INT_ICONID_SPEEXAM, False)
    
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_Pathol_Request, "��������", "��������", 0, False, lngIndex + 2)
     
End Sub

Public Sub IWorkMenuV2_zlCreateToolBar(ByVal strModuleName As String, objToolBar As Object)
'����������
    If Len(strModuleName) <= 0 Then Call CreatePublicPatholTools(objToolBar)
End Sub



Public Sub IWorkMenuV2_zlClearMenu(ByVal strModuleName As String)
'����������Ĳ˵�
    Dim cbrControl As CommandBarControl
    
    If mObjActiveMenuBar Is Nothing Then Exit Sub
    
    If Len(strModuleName) <= 0 Then
        Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_PatholManage)
        If Not cbrControl Is Nothing Then Call cbrControl.Delete
    Else
        Call ClearPrivaePatholMenu(strModuleName, mObjActiveMenuBar)
    End If
End Sub


Public Sub IWorkMenuV2_zlClearToolBar(ByVal strModuleName As String)
'��������Ĺ�����
    Dim cbrControl As CommandBarControl
    
    If mObjActiveMenuBar Is Nothing Then Exit Sub
    
    If Len(strModuleName) <= 0 Then
        Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Pathol_BatProcess)
        If Not cbrControl Is Nothing Then Call cbrControl.Delete
        
        Set cbrControl = mObjActiveMenuBar.FindControl(, conMenu_Pathol_Request)
        If Not cbrControl Is Nothing Then Call cbrControl.Delete
    End If
End Sub

Public Sub IWorkMenuV2_zlExecuteMenu(ByVal strModuleName As String, ByVal lngMenuId As Long)
'���ݲ˵�IDִ�ж�Ӧ����
    Dim objCbrControl As XtremeCommandBars.CommandBarControl
    Dim objIWork As IWorkMenuV2
    
    Select Case lngMenuId
        Case conMenu_Pathol_NumConfig     '������������
            Call Menu_Manage_����������
            
        Case conMenu_Pathol_Antibody_Manage    '�������
            Call Menu_Manage_�������
            
        Case conMenu_Pathol_MealManage        '�ײ�ά��
            Call Menu_Manage_�ײ�ά��
            
        Case conMenu_Pathol_Request     '��������
            Call Menu_Manage_��������
            
        Case conMenu_Pathol_ReportDelay       '�ӳٵǼ�
            Call Menu_Manage_�ӳٵǼ�
        
        Case conMenu_Pathol_ConRequest, conMenu_Pathol_ConFeedback       '�������뷴��
            Call Menu_Manage_�������뷴��(lngMenuId)
            
        Case conMenu_Pathol_Decalin_Task       '�Ѹ�����
            Call Menu_Manage_�Ѹ��������
            
        Case conMenu_Pathol_BatSlicesAccept   '��Ƭ��������
            Call Menu_Manage_��Ƭ��������
            
        Case conMenu_Pathol_BatSpeExamAccept  '�ؼ���������
            Call Menu_Manage_�ؼ���������
            
        Case conMenu_Pathol_Quality_Manage '������������
            Call Menu_Manage_������������
            
         Case conMenu_PatholSlices_Quality  '������Ƭ����
            Call Menu_Manage_������Ƭ����
    End Select
End Sub


Public Sub IWorkMenuV2_zlUpdateMenu(ByVal strModuleName As String, Control As XtremeCommandBars.ICommandBarControl)
'���²˵�
    Dim objIWork As IWorkMenuV2
    
    Select Case Control.ID
       Case conMenu_PatholManage
            Control.Enabled = True
        
       Case conMenu_Pathol_Antibody_Manage
            If Not (CheckPopedom(mstrPrivs, "�������") Or CheckPopedom(mstrPrivs, "���巴��")) Then
                Control.Enabled = False
            Else
            
            End If
        Case conMenu_Pathol_MealManage
            If Not CheckPopedom(mstrPrivs, "�ײ�ά��") Then
                Control.Enabled = False
            Else
            
            End If
        Case conMenu_Pathol_Request
            If Not (CheckPopedom(mstrPrivs, "�ؼ�����") Or CheckPopedom(mstrPrivs, "��Ƭ����") Or CheckPopedom(mstrPrivs, "��ȡ����")) Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_ReportDelay
            If Not CheckPopedom(mstrPrivs, "�����ӳ�") Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_ConRequest
            If Not CheckPopedom(mstrPrivs, "��������") Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_ConFeedback
            If Not CheckPopedom(mstrPrivs, "���ﷴ��") Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_Decalin_Task
            If Not CheckPopedom(mstrPrivs, "����ȡ��") Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_BatSlicesAccept ', conMenu_Pathol_BatSlicesSure
            If Not CheckPopedom(mstrPrivs, "������Ƭ") Then
                Control.Enabled = False
            Else
'                control.Enabled = Not mblnReadOnly
            End If
            
        Case conMenu_Pathol_BatSpeExamAccept
            If Not (CheckPopedom(mstrPrivs, "�����黯") Or CheckPopedom(mstrPrivs, "���Ӳ���") Or CheckPopedom(mstrPrivs, "����Ⱦɫ")) Then
                Control.Enabled = False
            Else
'                control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_Pathol_Quality_Manage
            If Not CheckPopedom(mstrPrivs, "��������") Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnReadOnly
            End If
        Case conMenu_PatholSlices_Quality
            '������߱���������Ȩ�ޣ�����Բ鿴��Ƭ����������
'            If Not CheckPopedom(mstrPrivs, "��������") Then
'                control.Enabled = False
'            End If
            Control.Enabled = Not mblnReadOnly
        Case conMenu_Pathol_BatProcess
            If Not (CheckPopedom(mstrPrivs, "������Ƭ") Or CheckPopedom(mstrPrivs, "�����黯") Or CheckPopedom(mstrPrivs, "���Ӳ���") Or CheckPopedom(mstrPrivs, "����Ⱦɫ")) Then
                Control.Enabled = False
            Else
'                control.Enabled = Not mblnReadOnly
            End If
    End Select
End Sub


Public Sub IWorkMenuV2_zlPopupMenu(ByVal strModuleName As String, objPopup As XtremeCommandBars.ICommandBar)
'�����Ҽ��˵�
    Exit Sub
End Sub

Public Sub IWorkMenuV2_zlRefreshSubMenu(ByVal strModuleName As String, objMenuBar As Object)
'ˢ�µ������Ӳ˵�
    Exit Sub
End Sub
'*********************************************************************************************


Public Sub zlInitModule(objNotify As IEventNotify, ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long)
'��ʼ��ģ�����
    Dim i As Long
    
    Set mObjNotify = objNotify
    
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngCurDeptId = lngDepartId
     

    For i = LBound(aryRegModule) To UBound(aryRegModule)
        If Not aryRegModule(i) Is Nothing Then
            Call aryRegModule(i).zlInitModule(lngModule, strPrivs, lngDepartId, mObjNotify.Owner)
        End If
    Next i
End Sub


Public Sub zlRefresh(objStudyInfo As clsStudyInfo, ByVal strSelModuleName As String, _
    Optional ByVal blnIsForceRefresh As Boolean = False, Optional ByVal blnIsHistory As Boolean = False)
'����ҽ����Ϣ
    
    If Not mobjStudyInfo Is Nothing And Not objStudyInfo Is Nothing Then
        If mobjStudyInfo.IsEquals(objStudyInfo) And blnIsForceRefresh = False Then Exit Sub
    End If
    
    Set mobjStudyInfo = objStudyInfo

    mblnReadOnly = False
    
    Select Case strSelModuleName
        Case C_TAB_NAME_�걾����
            Call mfrmPatholSpecimen.zlUpdateAdviceInf(mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, mobjStudyInfo.intStep, mobjStudyInfo.blnMoved)
            Call mfrmPatholSpecimen.zlRefreshFace(blnIsForceRefresh)
            
        Case C_TAB_NAME_����ȡ��
            Call mfrmPatholMaterial.zlUpdateAdviceInf(mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, mobjStudyInfo.intStep, mobjStudyInfo.blnMoved)
            Call mfrmPatholMaterial.zlRefreshFace(blnIsForceRefresh)
            
        Case C_TAB_NAME_������Ƭ
            Call mfrmPatholSlices.zlUpdateAdviceInf(mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, mobjStudyInfo.intStep, mobjStudyInfo.blnMoved)
            Call mfrmPatholSlices.zlRefreshFace(blnIsForceRefresh)
            
        Case C_TAB_NAME_�����ؼ�
            Call mfrmPatholSpeExam.zlUpdateAdviceInf(mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, mobjStudyInfo.intStep, mobjStudyInfo.blnMoved)
            Call mfrmPatholSpeExam.zlRefreshFace(blnIsForceRefresh)
            
        Case C_TAB_NAME_���̱���
            Call mfrmPatholProRep.zlUpdateAdviceInf(mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, mobjStudyInfo.intStep, mobjStudyInfo.blnMoved)
            Call mfrmPatholProRep.zlRefreshFace(blnIsForceRefresh)
            
    End Select
    
    '���Ϊ��ɣ��Ǽǣ��ܾ�����ˣ����ܶԼ����д��� 'Or lngStudyState = 5
    If mobjStudyInfo.lngAdviceId <= 0 Or mobjStudyInfo.intStep = 6 Or mobjStudyInfo.intStep = 0 Or mobjStudyInfo.intStep = 1 Or mobjStudyInfo.intStep = -2 Then
        mblnReadOnly = True
    End If
End Sub



Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, _
    Optional blnStartGroup As Boolean = False, Optional lngIndex As Long) As CommandBarControl
'������ģ���ڵĲ˵�
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If
    
    CreateModuleMenu.ID = lngID '������ﲻָ��id�����ܽ���Щ�˵����ӵ��Ҽ��˵���
    
    If lngIconId <> 0 Then CreateModuleMenu.iconid = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = M_STR_MODULE_MENU_TAG
End Function


Private Sub Menu_Manage_����������()
Dim frmPatholNum As New frmPatholNumConfig
On Error GoTo errFree
    Call frmPatholNum.Show(1, mObjNotify.Owner)
    
errFree:
    Call Unload(frmPatholNum)
    Set frmPatholNum = Nothing
End Sub

Public Function GetWindowCaption() As String
    GetWindowCaption = "�������ϵͳ"
End Function

Private Sub Menu_Manage_�������()
'�������
If Not (CheckPopedom(mstrPrivs, "�������") Or CheckPopedom(mstrPrivs, "���巴��")) Then
    Call MsgBoxD(mObjNotify.Owner, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, GetWindowCaption)
    Exit Sub
End If


Dim frmAntibody As New frmPatholAntibody
On Error GoTo errFree
    Call frmAntibody.ShowAntibodyManageWind(mstrPrivs, mObjNotify.Owner)
errFree:
    Call Unload(frmAntibody)
    Set frmAntibody = Nothing
End Sub



Private Sub Menu_Manage_�ײ�ά��()
'�ײ�ά��

If Not CheckPopedom(mstrPrivs, "�ײ�ά��") Then
    Call MsgBoxD(mObjNotify.Owner, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, GetWindowCaption)
    Exit Sub
End If

Dim frmMeal As New frmPatholMeal
On Error GoTo errFree
    Call frmMeal.ShowMealWindow(mstrPrivs, mObjNotify.Owner)
errFree:
    Call Unload(frmMeal)
    Set frmMeal = Nothing
End Sub


Private Sub Menu_Manage_��������()
'��������
On Error GoTo errFree
    Dim frmRequest As frmPatholRequisition

    If Not (CheckPopedom(mstrPrivs, "�ؼ�����") Or CheckPopedom(mstrPrivs, "��Ƭ����") Or CheckPopedom(mstrPrivs, "��ȡ����")) Then
        Call MsgBoxD(mObjNotify.Owner, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If
    
    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD mObjNotify.Owner, M_STR_HINT_NoSelectData, vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    Set frmRequest = New frmPatholRequisition
    
    Call frmRequest.zlRefresh(mobjStudyInfo.lngAdviceId, mblnReadOnly, mstrPrivs, mobjStudyInfo.blnMoved, mlngCurDeptId, mObjNotify.Owner)

    If frmRequest.blnIsUpdate Then
        Call SendMsgToMainWindow(Me, wetPatholRequest, mobjStudyInfo.lngAdviceId)
    End If
errFree:
    Call Unload(frmRequest)
    Set frmRequest = Nothing
End Sub


Private Sub Menu_Manage_�ӳٵǼ�()
'�ӳٵǼ�
On Error GoTo errFree
    Dim frmDelay As frmPatholReportDelay

    If Not CheckPopedom(mstrPrivs, "�����ӳ�") Then
        Call MsgBoxD(mObjNotify.Owner, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If

    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    Set frmDelay = frmPatholReportDelay

    Call frmDelay.zlRefresh(mobjStudyInfo.lngAdviceId, mblnReadOnly, mstrPrivs, mobjStudyInfo.blnMoved, mlngCurDeptId, mObjNotify.Owner)
errFree:
    Call Unload(frmDelay)
    Set frmDelay = Nothing
End Sub


Private Sub Menu_Manage_������������()
'������������
On Error GoTo errFree
    Dim frmQulity As frmPatholQuality

    If Not CheckPopedom(mstrPrivs, "��������") Then
        Call MsgBoxD(mObjNotify.Owner, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If

    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    Set frmQulity = New frmPatholQuality

    Call frmQulity.zlRefresh(mobjStudyInfo.lngAdviceId, mblnReadOnly, mstrPrivs, mobjStudyInfo.blnMoved, mlngCurDeptId, mObjNotify.Owner)
    
    If frmQulity.blnIsOk Then
        Call SendMsgToMainWindow(Me, wetPatholQuality, mobjStudyInfo.lngAdviceId, frmQulity.strQuality)
    End If
errFree:
    Call Unload(frmQulity)
    Set frmQulity = Nothing
End Sub

Private Sub Menu_Manage_������Ƭ����()
'��ʾ��Ƭ����
On Error GoTo errhandle
     Dim frmSlicesQuality As frmPatholSlices_Quality
    
     If Not CheckPopedom(mstrPrivs, "��������") Then
        Call MsgBoxD(mObjNotify.Owner, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, GetWindowCaption)
        Exit Sub
     End If
     
     Set frmSlicesQuality = New frmPatholSlices_Quality
     
     Call frmSlicesQuality.ShowSlideEvaluateWindow(mobjStudyInfo.lngAdviceId, mobjStudyInfo.intStep, mstrPrivs, mObjNotify.Owner)
     
Exit Sub
errhandle:
    Call Unload(frmSlicesQuality)
    Set frmSlicesQuality = Nothing
End Sub


Private Sub Menu_Manage_�������뷴��(ByVal lngMenuId As Long)
'�������뷴��
On Error GoTo errFree

    If Not (CheckPopedom(mstrPrivs, "��������") Or CheckPopedom(mstrPrivs, "���ﷴ��")) Then
        Call MsgBoxD(mObjNotify.Owner, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If

    If mobjStudyInfo.lngAdviceId = 0 Then
        MsgBoxD mObjNotify.Owner, M_STR_HINT_NoSelectData, vbInformation, GetWindowCaption
        Exit Sub
    End If

    Call mfrmConRequest.zlRefresh(mobjStudyInfo.lngAdviceId, mblnReadOnly, mstrPrivs, mobjStudyInfo.blnMoved, mlngCurDeptId, IIf(lngMenuId = conMenu_Pathol_ConFeedback, True, False), mObjNotify.Owner)

    Exit Sub
errFree:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Manage_�Ѹ��������()
'�Ѹ��������

If Not CheckPopedom(mstrPrivs, "����ȡ��") Then
    Call MsgBoxD(mObjNotify.Owner, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, GetWindowCaption)
    Exit Sub
End If

Call mfrmPatholDecalinTask.ShowDecalinTaskWind(mstrPrivs, mobjStudyInfo.blnMoved, mlngModule, mObjNotify.Owner)

End Sub



Private Sub Menu_Manage_��Ƭ��������()
'��Ƭ��������
On Error GoTo errFree
    Dim frmBatSlices As frmPatholSlices_WorkPrint
  
    If Not CheckPopedom(mstrPrivs, "������Ƭ") Then
        Call MsgBoxD(mObjNotify.Owner, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If
    
    Set frmBatSlices = New frmPatholSlices_WorkPrint
    
    Call frmBatSlices.ShowWorkPrint(Nothing, 0, mObjNotify.Owner)
        
        
    If frmBatSlices.blnIsOk Then
        Call SendMsgToMainWindow(Me, wetPatholBatSlices, mobjStudyInfo.lngAdviceId)
    End If
errFree:
    Call Unload(frmBatSlices)
    Set frmBatSlices = Nothing
End Sub



Private Sub Menu_Manage_�ؼ���������()
'�ؼ���������
On Error GoTo errFree
    Dim frmBatSpeExam As frmPatholSpecialExamined_WorkPrint
    
    If Not (CheckPopedom(mstrPrivs, "�����黯") Or CheckPopedom(mstrPrivs, "����Ⱦɫ") Or CheckPopedom(mstrPrivs, "���Ӳ���")) Then
        Call MsgBoxD(mObjNotify.Owner, "���߱�ִ�иò�����Ȩ�ޡ�", vbOKOnly, GetWindowCaption)
        Exit Sub
    End If

    Set frmBatSpeExam = frmPatholSpecialExamined_WorkPrint
                
    If CheckPopedom(mstrPrivs, "�����黯") Then
        Call frmBatSpeExam.ShowWorkPrint(Nothing, 0, 0, mstrPrivs, mObjNotify.Owner)
    ElseIf CheckPopedom(mstrPrivs, "����Ⱦɫ") Then
        Call frmBatSpeExam.ShowWorkPrint(Nothing, 0, 1, mstrPrivs, mObjNotify.Owner)
    ElseIf CheckPopedom(mstrPrivs, "���Ӳ���") Then
        Call frmBatSpeExam.ShowWorkPrint(Nothing, 0, 2, mstrPrivs, mObjNotify.Owner)
    End If
    
    If frmBatSpeExam.blnIsOk Then
        Call SendMsgToMainWindow(Me, wetPatholBatSpeExm, mobjStudyInfo.lngAdviceId)
    End If
errFree:
    Call Unload(frmBatSpeExam)
    Set frmBatSpeExam = Nothing
End Sub
 