VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRichReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const M_STR_MODULE_MENU_TAG = ""

Public Event BeforeEdit(ByVal lngOrderID As Long, ByVal strDocID As String)
Public Event AfterOpen(ByVal lngOrderID As Long, ByVal strDocID As String)
Public Event AfterSaved(ByVal lngOrderID As Long, ByVal strDocID As String, ByVal lngSaveType As Long)
Public Event AfterClosed(ByVal lngOrderID As Long, ByVal strDocID As String)
Public Event AfterPrinted(ByVal lngOrderID As Long, ByVal strDocID As String)
Public Event AfterDeleted(ByVal lngOrderID As Long, ByVal strDocID As String)
Public Event DataRefresh(ByVal strDocInfo As String)
Public Event DocPluginAction(ByVal actionType As Long, ByVal data As String, ByVal tag As String)

Private WithEvents mobjInterfaceEvent As ZLPacsVBCommon.clsEvent
Attribute mobjInterfaceEvent.VB_VarHelpID = -1
Private mobjPacsInterface As Object
Attribute mobjPacsInterface.VB_VarHelpID = -1
Private mobjOwner As Object
Private mstrPrivs As String
Private mlngModule As Long
Private mlngDeptID As Long
Private mlngAdviceID As Long
Private mstrTmpReportId As String
Private mlngPatientId As Long
Private mlngStudyState As Long
Private mblnCanPrint As Boolean     'ƽ������˺���ܴ�ӡ
Private mblnDirectPrint As Boolean  '�������˺���ɫͨ�����˿�ֱ�Ӵ�ӡ

Private mblnVisible As Boolean

Private mblnRefreshState As Boolean
Private mobjPacsCore As zl9PacsCore.clsViewer
Private mObjActiveMenuBar As CommandBars
Private mlngReportFormHandle As Long

Private mobjFtp As New clsFtp
Private mobjFtpInf As TFtpDeviceInf
Private mstrSubDir As String
Private mstrDestMainDir As String
Private mblnIsNoAskPrint As Boolean             '��Ĭ��ӡ,True--��Ĭ��ӡ,False--�Ǿ�Ĭ��ӡ
Private mintCompletePrint As Integer
Private mblnIsShowLastState As Boolean
Private mblnIsFirstImgUpLoadFiald As Boolean      '�����һ��ͼ���ϴ�ʧ�ܣ����Ժ���ʧ�ܣ�������ͼ���������ϴ�
Private mblnIsFirstImgDownLoadFiald As Boolean      '�����һ��ͼ������ʧ�ܣ����Ժ���ʧ�ܣ�������ͼ������������

Property Get AdviceId()
'��ȡ��ǰ�����ҽ��Id
    AdviceId = mlngAdviceID
End Property

'��Ĭ��ӡ
Public Property Get IsNoAskPrint() As Boolean
    IsNoAskPrint = mblnIsNoAskPrint
End Property

Public Property Let IsNoAskPrint(ByVal vNewValue As Boolean)
    mblnIsNoAskPrint = vNewValue
    Call SetIsNoAskPrint(vNewValue)
End Property


'���ñ����ͼ��������
Property Get PacsCore() As zl9PacsCore.clsViewer
    Set PacsCore = mobjPacsCore
End Property

Property Set PacsCore(objPacsCore As zl9PacsCore.clsViewer)
    Set mobjPacsCore = objPacsCore
End Property


Public Sub NotificationRefresh()
'֪ͨˢ��
    mblnRefreshState = False
End Sub


Public Sub zlCreateMenu(ByVal objActiveMenuBar As Object)
'����ģ��˵�����ť
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar

On Error GoTo ErrorHand
    Set mObjActiveMenuBar = objActiveMenuBar

    Set cbrMenuBar = objActiveMenuBar.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����", 3, False)
    cbrMenuBar.ID = conMenu_EditPopup
    cbrMenuBar.Category = ""
    
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_NoAskPrint, "ʹ�þ�Ĭ��ӡ", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PacsReport_FinalShowMode, "����״̬��ʾ", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Preview, "Ԥ��", "", 102, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Print, "��ӡ", "", 103, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_BatPrint, "������ӡ", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PacsReport_Open, "��д", "", 3002, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PacsReport_ClearWritingState, "���״̬", "", 21903, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Edit_Delete, "ɾ��", "", 0, False)
    End With
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Public Sub zlCreateToolBar(ByVal objToolBar As Object)
'������������ť
    Dim cbrControl As CommandBarControl
    Dim cbrLogOut As CommandBarControl
    Dim lngIndex As Long

On Error GoTo ErrorHand
    Set cbrLogOut = objToolBar.FindControl(, conMenu_Manage_InQueue, , True)
    
    lngIndex = 4
    If Not cbrLogOut Is Nothing Then lngIndex = cbrLogOut.Index

    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_File_Preview, "Ԥ��", "����Ԥ��", 102, True, lngIndex + 1)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_File_Print, "��ӡ", "�����ӡ", 103, False, lngIndex + 2)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_PacsReport_Open, "��д", "", 2607, False, lngIndex + 3) 'IconId=3002
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Public Sub zlExecuteMenu(ByVal lngMenuId As Long, Optional ByVal strDocID As String = 0)
'ִ�б༭���˵�
    Dim objControl As XtremeCommandBars.ICommandBarControl
        
On Error GoTo ErrorHand
    Set objControl = mObjActiveMenuBar.FindControl(, lngMenuId, , True)
    If objControl Is Nothing Then Exit Sub
    
    Select Case lngMenuId
        Case conMenu_PacsReport_FinalShowMode
            objControl.Checked = Not objControl.Checked
            mblnIsShowLastState = objControl.Checked
            zlDatabase.SetPara "����״̬��ʾ", IIf(mblnIsShowLastState, 1, 0), glngSys, mlngModule
            SetReportFinalShowMode mblnIsShowLastState
            
        Case conMenu_File_Preview                   'Ԥ��
            Call OpenFormForReportPrint(strDocID, True)
            
        Case conMenu_File_Print                     '��ӡ
            Call OpenFormForReportPrint(strDocID, False)
            
        Case conMenu_File_BatPrint                  '������ӡ
            Call OpenFormForReportPrint(strDocID, False)
            
        Case conMenu_PacsReport_Open                '��д
            Call OpenFormForEditReport
        
        Case conMenu_PacsReport_ClearWritingState   '���״̬
            Call ClearWritingState
                
        Case conMenu_Edit_Delete                    'ɾ��
            Call DeleteReport
        
    End Select
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Private Sub ClearWritingState()
'���״̬
    Dim strDocID As String
    
On Error GoTo ErrorHand
    
    If mobjPacsInterface Is Nothing Then Exit Sub
    
    strDocID = GetCurrDocId(mlngAdviceID)
    
    If Trim(strDocID) = "" Then Exit Sub
    
    Call mobjPacsInterface.EditLockedInfo(strDocID, "")
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Public Sub zlUpdateMenu(ByVal objControl As XtremeCommandBars.ICommandBarControl)
    
'���±༭���˵�
On Error GoTo ErrorHand
    If objControl Is Nothing Then Exit Sub
    If mobjPacsInterface Is Nothing Then Exit Sub
    
    Select Case objControl.ID
        Case conMenu_PacsReport_FinalShowMode
            objControl.Checked = mblnIsShowLastState
            
        Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_BatPrint     '��ӡ����,Ԥ������,������ӡ
            objControl.Visible = CheckPopedom(mstrPrivs, "PACS�����ӡ")
             
            If objControl.ID <> conMenu_File_BatPrint Then
                If Not mobjPacsInterface Is Nothing Then
                    objControl.Enabled = mobjPacsInterface.CanPrint("")
                Else
                    objControl.Enabled = False
                End If
            End If
            
        Case conMenu_PacsReport_Open                '��д
            Call InitEditButtonCaption(objControl)
            
        Case conMenu_Edit_Delete                    'ɾ��
            objControl.Visible = (GetCurrDocId(mlngAdviceID) <> "" And (CheckPopedom(mstrPrivs, "PACS������д") Or CheckPopedom(mstrPrivs, "PACS����ɾ��")))
            If objControl.Visible Then objControl.Visible = (CheckPopedom(mstrPrivs, "PACS���˱���") Or GetDocCreator(mlngAdviceID) = UserInfo.����)
        Case conMenu_File_Open                      '����
            
        Case conMenu_Tool_Search                    '���������
            
    End Select
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub InitEditButtonCaption(ByVal objControl As XtremeCommandBars.ICommandBarControl)
'���ݱ���״̬��ʾ��д��ť����
    Dim intReportState As Integer
On Error GoTo ErrorHand
    '1.��û��ǩ��������£���д��ť��ʾ����Ϊ����д��
    '2.���ǩ�������û�б������Ȩ�ޣ���д��ť����ʾ���޶�������������Ȩ�ޣ���д��ť����ʾ����ˡ�
    '3.���ǩ������д��ť����ʾ����ˡ�
    '4.����ǩ������д��ť��ʾ�����ġ�
    '5.���ش�������ϲ��ذ����ǩ����������˲��ذ����ǩ������
    If mlngAdviceID <= 0 Then
        objControl.Enabled = False
        objControl.Caption = "��д"
        objControl.ToolTipText = objControl.Caption
        Exit Sub
    End If
    
    intReportState = GetCurrReportState(mlngAdviceID)
    
    If intReportState = 1 Then                      '�༭��
        objControl.Caption = "��д"
        objControl.Enabled = CheckPopedom(mstrPrivs, "PACS������д") Or CheckPopedom(mstrPrivs, "PACS�����޶�") Or CheckPopedom(mstrPrivs, "PACS��������")
        
    ElseIf intReportState = 2 Or intReportState = 5 Then                  '����ϣ���ϲ���
        If CheckPopedom(mstrPrivs, "PACS�����޶�") Then '��PACS�����޶���Ȩ�ޣ��ڱ����ĵ��༭���е� ��� Ȩ��ʹ��
            objControl.Caption = "���"
        Else
            objControl.Caption = "�޶�"
        End If
        
    ElseIf intReportState = 3 Or intReportState = 6 Then                  '����ˣ���˲���
        objControl.Caption = "���"
        objControl.Enabled = CheckPopedom(mstrPrivs, "PACS��������")
        
    ElseIf intReportState = 4 Then                  '������
        objControl.Caption = "����"
        objControl.Enabled = True
    End If
    
    objControl.ToolTipText = objControl.Caption
    Exit Sub
ErrorHand:
    objControl.Caption = "��д"
    objControl.ToolTipText = objControl.Caption
End Sub

Public Function zlInit(ByVal lngSysNo As Long, ByVal lngDeptID As Long, ByVal strCode As String, ByVal strPrivs As String, objOwner As Object) As Boolean
'��ʼ����������
    Dim strUser As String, strPwd As String, strServer As String
    Dim lngIndex As Long
    Dim blnResult As Boolean
    
On Error GoTo ErrorHand
    Set mobjOwner = objOwner
    mstrPrivs = strPrivs
    mlngDeptID = lngDeptID
    mlngModule = strCode
    
    strUser = UCase(gcnOracle.Properties(23))
    strPwd = gcnOracle.Properties(24)
    strServer = gcnOracle.Properties(8) 'UCase(Split(gcnOracle.Properties(8), "=")(2))
    
    lngIndex = InStr(strServer, "(DESCRIPTION=")
    If lngIndex >= 1 Then
        strServer = Mid(strServer, lngIndex, 255)
    Else
        strServer = UCase(Split(gcnOracle.Properties(8), "=")(2))
    End If
    
    mblnIsShowLastState = zlDatabase.GetPara("����״̬��ʾ", glngSys, mlngModule, 0) = "1"
    mblnIsNoAskPrint = zlDatabase.GetPara("NoAsk", glngSys, 1070, 0) = "1"
    mblnCanPrint = GetDeptPara(mlngDeptID, "ƽ������˲��ܴ򱨸�") = "1"             'ƽ����Ҫ��˲��ܴ�ӡ =true
    mintCompletePrint = Val(GetDeptPara(mlngDeptID, "�����ֱ�Ӵ�ӡ", 0))
    
    Set mobjInterfaceEvent = New ZLPacsVBCommon.clsEvent
    Set mobjPacsInterface = DynamicCreate("ZLSoft.ZLPacs.Interface.PacsInterface", "PACS���ܱ���༭���ӿڲ���zlPacsInterfaceC")
    
    If mobjPacsInterface Is Nothing Or mobjInterfaceEvent Is Nothing Then Exit Function
    If InitRegister = False Then Exit Function
    
    BUGEX "zlInit InitRegister " & gstrUserPswd
    
    blnResult = mobjPacsInterface.Init(mobjInterfaceEvent, lngSysNo, strCode, objOwner.Caption, gstrUserName, gstrUserPswd, gstrServerName, lngDeptID)
    
    '���ñ༭�����ԣ���ͬ�汾���Է��������в��죬�����Ҫ���Ե����쳣
    Call SetAttribute
    
    If blnResult Then
        Call SetReportFinalShowMode(mblnIsShowLastState)
        Call SetIsNoAskPrint(mblnIsNoAskPrint)
    End If
    
    zlInit = blnResult
    Exit Function
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Sub SetAttribute()
On Error Resume Next
    If mobjPacsInterface Is Nothing Then Exit Sub
    
    Call mobjPacsInterface.SetOwner(mobjOwner.hWnd)
err.Clear
End Sub

Public Sub zlUpdateAdviceInf(ByVal lngAdviceID As Long, ByVal lngPatientID As Long, ByVal lngStudyState As Long)
'ͬ��ҽ����Ϣ
    mlngPatientId = lngPatientID
    mlngAdviceID = lngAdviceID
    mlngStudyState = lngStudyState
    
End Sub

Public Sub zlEditReport()
'��д����
    Call OpenFormForEditReport
End Sub


Public Sub zlRefreshFace(ByVal blnCanPrint As Boolean, ByVal blnForceRefresh As Boolean)
'ˢ��
    mblnDirectPrint = blnCanPrint
    
    Call RefreshDocReportForm(blnForceRefresh)
    
    BUGEX "zlRefreshFaceRich mblnDirectPrint = " & mblnDirectPrint
    BUGEX "zlRefreshFaceRich mblnCanPrint = " & mblnCanPrint
End Sub


Public Function zlGetForm() As Long
'��ȡ���洰�ھ��
    If mobjPacsInterface Is Nothing Then Exit Function
    
    If mlngReportFormHandle <= 0 Then mlngReportFormHandle = mobjPacsInterface.GetReportFormHandle(mlngAdviceID, mlngPatientId, mlngStudyState, mblnDirectPrint)
    zlGetForm = mlngReportFormHandle
    mblnVisible = True
End Function


Public Sub CloseForms()
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.CloseForms
    mblnVisible = False
End Sub

Private Sub OpenFormForReportPrint(ByVal strDocID As String, ByVal blnPreview As Boolean)
'Ԥ������ӡ
    Dim strResultInput As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo ErrorHand
    If Not blnPreview Then
        '��ӡǰ¼����
        If Val(GetDeptPara(mlngDeptID, "��Ͻ����ʾ����", 0)) = 2 Then
            strResultInput = ""
            If GetDeptPara(mlngDeptID, "��Ӱ�����Ϊ����", 0) = "1" Then '��Ӱ�����Ϊ����  -����ʾ�Զ����
                Call EvaluatResult(mlngAdviceID, "0")
            End If
                
            strSql = "Select B.Σ��״̬, A.�������, B.Ӱ������, A.��������, B.������� " & _
                     "From Ӱ�񱨸��¼ A, Ӱ�����¼ B " & _
                     "Where A.Id=[1] and A.ҽ��Id = B.ҽ��Id"
                     
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�������", strDocID)
        
            If IsNull(rsTemp!Σ��״̬) And Val(GetDeptPara(mlngDeptID, "Σ������ж�", 0)) <> 0 Then strResultInput = "Σ��״̬|"
            If IsNull(rsTemp!�������) And Not GetDeptPara(mlngDeptID, "���Խ��������", 0) = "1" Then strResultInput = strResultInput & "�������|"
            If IsNull(rsTemp!Ӱ������) And NVL(GetDeptPara(mlngDeptID, "Ӱ�������ȼ�", "��,��")) <> "" And Val(GetDeptPara(mlngDeptID, "Ӱ�������ж�", "0")) <> 0 And CheckPopedom(mstrPrivs, "Ӱ���ʿ�") Then strResultInput = strResultInput & "Ӱ������|"
            If IsNull(rsTemp!��������) And NVL(GetDeptPara(mlngDeptID, "���������ȼ�", "��,��")) <> "" And Val(GetDeptPara(mlngDeptID, "���������ж�", 0)) <> 0 And CheckPopedom(mstrPrivs, "�����ʿ�") Then strResultInput = strResultInput & "��������|"
            If IsNull(rsTemp!�������) And Val(GetDeptPara(mlngDeptID, "��������ж�", 0)) <> 0 Then strResultInput = strResultInput & "�������|"
            
            If strResultInput <> "" Then Call PromptResultRich(mlngAdviceID, strDocID, mlngModule, Me, mlngDeptID, strResultInput)
        End If
    End If
    
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.PrintReport(mobjOwner.hWnd, strDocID, blnPreview)
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Private Function OpenFormForEditReport() As String
'��д--���������洰��ģʽ
'����:������Ϣ
    If mobjPacsInterface Is Nothing Then Exit Function
    Call ShowErrMessage(mobjPacsInterface.OpenFormForEditReport(mobjOwner.hWnd, mlngAdviceID, mlngPatientId, mlngStudyState, mblnDirectPrint))
End Function


Private Function DeleteReport() As String
'ɾ������
'����:������Ϣ
    Dim strDocID As String
    
    If mobjPacsInterface Is Nothing Then Exit Function
    
    strDocID = GetCurrDocId(mlngAdviceID)
    
    If Trim(strDocID) = "" Then Exit Function
    
    Call ShowErrMessage(mobjPacsInterface.DeleteReport(mobjOwner.hWnd, strDocID))
End Function


Private Function RefreshDocReportForm(ByVal isForce As Boolean) As String
'ˢ��Ƕ��ʽ���洰��
    If mobjPacsInterface Is Nothing Then Exit Function
    If mblnRefreshState = False Then
        isForce = True
        mblnRefreshState = True
    End If
    Call ShowErrMessage(mobjPacsInterface.RefreshReportForm(mlngAdviceID, mlngPatientId, mlngDeptID, mlngStudyState, mblnDirectPrint, isForce))
End Function

Private Sub ShowErrMessage(ByVal strErrMessage As String)
    If strErrMessage = "" Then Exit Sub
    
    MsgboxEx mobjOwner, strErrMessage, vbExclamation, gstrSysName
End Sub

Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'������ģ���ڵĲ˵�
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If
    
    CreateModuleMenu.ID = lngID '������ﲻָ��Id�����ܽ���Щ�˵����ӵ��Ҽ��˵���
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = "" 'M_STR_MODULE_MENU_TAG
End Function


Private Sub Class_Terminate()
    If Not mobjInterfaceEvent Is Nothing Then Set mobjInterfaceEvent = Nothing
    If Not mobjPacsInterface Is Nothing Then Set mobjPacsInterface = Nothing
    
    mblnVisible = False
End Sub

Public Sub CreatePacsInterface()
    If mobjPacsInterface Is Nothing Then Set mobjPacsInterface = DynamicCreate("ZLSoft.ZLPacs.Interface.PacsInterface", "PACS���ܱ���༭���ӿڲ���zlPacsInterfaceC")
End Sub

'�¼�
Private Sub mobjInterfaceEvent_OnDataRefresh(ByVal strDocID As String)
    RaiseEvent DataRefresh(strDocID)
End Sub

Private Sub mobjInterfaceEvent_OnDocPluginAction(ByVal strPluginID As String, ByVal strPluginName As String, ByVal lngActionType As Long, ByVal strData As String, ByVal strTag As String)
    RaiseEvent DocPluginAction(lngActionType, strData, strTag)
End Sub

Private Sub mobjInterfaceEvent_OnImageLoadFromFTP(ByVal lngOrderID As Long, ByVal strDocID As String, ByVal strImageName As String, ByVal strImagePath As String, blnResult As Boolean)
'��FTP�����ر���ͼ���ص�������
    Dim objFSO As New Scripting.FileSystemObject
    Dim i As Integer
    Dim strImages() As String
    Dim strLoadFailureImgs As String
    Dim strLocalFileName As String
    Dim strVirtualPath As String
    Dim blnLoadResult As Boolean
    
On Error GoTo ErrorHand
    mblnIsFirstImgDownLoadFiald = False
    
    If lngOrderID <= 0 Or strImageName = "" Or strImagePath = "" Then
        blnResult = True
        Exit Sub
    End If
    
    strImages = Split(strImageName, "���")
    
    '��������·��
    If Not objFSO.FolderExists(strImagePath) Then Call MkLocalDir(strImagePath)
    
    '��ʼ��FTP��Ϣ
    If Not InitFtpInf(strDocID) Then Exit Sub
    
    For i = 0 To UBound(strImages)
        strVirtualPath = Replace(mobjFtpInf.strFtpDir & mstrSubDir, "\", "/")
        strLocalFileName = Replace(strImagePath & "\" & strImages(i), "/", "\")
        
        '������ز��������FTP������
        If Not objFSO.FileExists(strLocalFileName) Then
            '����FTP
            If Not ConnFtp Then Exit Sub
            
            '�Ӵ洢�豸�ϻ�ȡͼ��
            If mblnIsFirstImgDownLoadFiald Then
                blnLoadResult = mobjFtp.FuncDownloadFile(strVirtualPath, strLocalFileName, strImages(i)) = 1
            Else
                blnLoadResult = FuncDownloadFile(mobjFtp, strVirtualPath, strLocalFileName, strImages(i))
            End If
            
            If Not blnLoadResult Then strLoadFailureImgs = strLoadFailureImgs & "��" & strImages(i)
        End If
    Next
    
    If Mid(strLoadFailureImgs, 2) <> "" Then
        If UBound(Split(Mid(strLoadFailureImgs, 2), "��")) = UBound(strImages) Then
            MsgBox "ͼ������ʧ�ܣ������Ƿ������ϲ�����ͼ������粻�ȶ���ɣ�" & vbCrLf & "ͼ����Ϊ��" & Mid(strLoadFailureImgs, 2), vbExclamation, gstrSysName
            blnResult = False
        Else
            MsgBox "����ͼ������ʧ�ܣ������Ƿ������ϲ�����ͼ������粻�ȶ���ɣ�" & vbCrLf & "ͼ����Ϊ��" & Mid(strLoadFailureImgs, 2), vbExclamation, gstrSysName
            blnResult = True
        End If
    Else
        blnResult = True
    End If
    
    Exit Sub
ErrorHand:
    blnResult = False
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub mobjInterfaceEvent_OnImageSaveToFTP(ByVal lngOrderID As Long, ByVal strDocID As String, ByVal strImageName As String, ByVal strImageBase64 As String, ByVal strFTPDeviceNo As String, ByVal strRptCreateDate As String, blnResult As Boolean)
'�������еı���ͼ���浽FTP
    Dim objFSO As New Scripting.FileSystemObject
    Dim strImages() As String
    Dim strImagesBase64() As String
    Dim strNeedUpImgs As String
    Dim strUpFailureImgs As String
    Dim i As Integer
    Dim blnLoadResult As Boolean
    
On Error GoTo ErrorHand
    mblnIsFirstImgUpLoadFiald = False
    
    If lngOrderID <= 0 Or strImageName = "" Then
        blnResult = True
        Exit Sub
    End If
    
    strImages = Split(strImageName, "���")
    strImagesBase64 = Split(strImageBase64, "���")
    
    '��ʼ��FTP��Ϣ
    If Not InitFtpInf(strDocID, strFTPDeviceNo, Format(strRptCreateDate, "YYYYMMDD")) Then Exit Sub
    
    '��������·��
    If Not objFSO.FolderExists(mstrDestMainDir & mstrSubDir) Then Call MkLocalDir(mstrDestMainDir & mstrSubDir)
    
    '����FTP
    If Not ConnFtp Then Exit Sub
    
    For i = 0 To UBound(strImages)
        If strImagesBase64(i) = "-1" Then
            'Ϊ��˵������ͼ�Ѿ�ɾ��,���ɾ��FTP�������϶�Ӧ��ͼ��
            Call mobjFtp.FuncDelFile(mobjFtpInf.strFtpDir & mstrSubDir, strImages(i))
            
            '������ش�����ɾ��
            If objFSO.FileExists(mstrDestMainDir & mstrSubDir & "\" & strImages(i)) Then Call objFSO.DeleteFile(mstrDestMainDir & mstrSubDir & "\" & strImages(i), True)
        Else
            strNeedUpImgs = strNeedUpImgs & "��" & strImages(i) '��¼��Ҫ�ϴ���ͼ������
            
            '��ͼ�������ļ�
            If WritTextFile(mstrDestMainDir & mstrSubDir & "\" & strImages(i), strImagesBase64(i)) Then
                '��ͼ���ļ��ϴ���FTP�洢�豸
                mobjFtp.FuncFtpMkDir "/", mobjFtpInf.strFtpDir & mstrSubDir
                
                If mblnIsFirstImgUpLoadFiald Then
                    blnLoadResult = mobjFtp.FuncUploadFile(mobjFtpInf.strFtpDir & mstrSubDir, mstrDestMainDir & mstrSubDir & "\" & strImages(i), strImages(i))
                Else
                    blnLoadResult = FuncUploadFile(mobjFtp, mobjFtpInf.strFtpDir & mstrSubDir, mstrDestMainDir & mstrSubDir & "\" & strImages(i), strImages(i))
                End If
                
                If Not blnLoadResult Then strUpFailureImgs = strUpFailureImgs & "��" & strImages(i)
            Else
                strUpFailureImgs = strUpFailureImgs & "��" & strImages(i)
            End If
        End If
    Next
    
    '�Ͽ�FTP����
    mobjFtp.FuncFtpDisConnect
    
    If Mid(strUpFailureImgs, 2) <> "" Then
        If UBound(Split(Mid(strUpFailureImgs, 2), "��")) = UBound(Split(Mid(strNeedUpImgs, 2), "��")) Then
            MsgBox "ͼ���ϴ���FTP������ʧ�ܣ����������粻�ȶ���ɣ�" & vbCrLf & "ͼ����Ϊ��" & Mid(strUpFailureImgs, 2), vbExclamation, gstrSysName
            blnResult = False
        Else
            MsgBox "����ͼ���ϴ���FTP������ʧ�ܣ����������粻�ȶ���ɣ�" & vbCrLf & "ͼ����Ϊ��" & Mid(strUpFailureImgs, 2), vbExclamation, gstrSysName
            blnResult = True
        End If
    Else
        blnResult = True
    End If
    
    Exit Sub
ErrorHand:
    blnResult = False
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub mobjInterfaceEvent_OnReportAction(ByVal lngReportEvtAction As ZLPacsVBCommon.TReportEvtActionType, ByVal lngOrderID As Long, ByVal strDocID As String, ByVal lngReportState As Long)
    On Error GoTo ErrorHand
    
    Select Case lngReportEvtAction
        Case TReportEvtActionType.ratAfterClosed
            RaiseEvent AfterClosed(lngOrderID, strDocID)
            Call RefreshDocReportForm(True)
            
        Case TReportEvtActionType.ratAfterDeleted
            RaiseEvent AfterDeleted(lngOrderID, strDocID)
        
        Case TReportEvtActionType.ratAfterOpen
            RaiseEvent AfterOpen(lngOrderID, strDocID)
        
        Case TReportEvtActionType.ratAfterPrinted
            RaiseEvent AfterPrinted(lngOrderID, strDocID)
        
        Case TReportEvtActionType.ratAfterReportStateChanged
            '�����ֱ�Ӵ�ӡ
            BUGEX "mobjInterfaceEvent_OnReportAction mintCompletePrint = " & mintCompletePrint & "  reportState = " & lngReportState
            If mintCompletePrint = 1 And lngReportState = 4 Then
                Call OpenFormForReportPrint(strDocID, False)
            End If
            
            '��ϲ��ػ�����˲��غ󣬷��𱣴��¼����޸ļ��״̬
            If lngReportState = 5 Or lngReportState = 6 Then
                'reportState0-��ͨ���棬1-���ǩ����2-��ˣ�����ǩ��
                RaiseEvent AfterSaved(lngOrderID, strDocID, lngReportState)
            End If
        Case TReportEvtActionType.ratAfterSaved
            'reportState0-��ͨ���棬1-���ǩ����2-��ˣ�����ǩ��
            RaiseEvent AfterSaved(lngOrderID, strDocID, lngReportState)
        
        Case TReportEvtActionType.ratBeforeEdit
            RaiseEvent BeforeEdit(lngOrderID, strDocID)
            
    End Select
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Function FuncDownloadFile(objFtp As clsFtp, strVirtualPath As String, strLocalFileName As String, strRemoteFileName As String) As Boolean
    If objFtp.FuncDownloadFile(strVirtualPath, strLocalFileName, strRemoteFileName) <> 0 Then
        If MsgBox("ͼ������ʧ�ܣ������Ƿ������ϲ�����ͼ������粻�ȶ���ɣ�", vbRetryCancel, gstrSysName) = vbRetry Then
            Call FuncDownloadFile(objFtp, strVirtualPath, strLocalFileName, strRemoteFileName)
        Else
            mblnIsFirstImgDownLoadFiald = True
            FuncDownloadFile = False
            Exit Function
        End If
    End If
    
    FuncDownloadFile = True
End Function

Private Function FuncUploadFile(objFtp As clsFtp, strVirtualPath As String, strLocalFileName As String, strRemoteFileName As String) As Boolean
    If objFtp.FuncUploadFile(strVirtualPath, strLocalFileName, strRemoteFileName) <> 0 Then
        If MsgBox("ͼ���ϴ���FTP������ʧ�ܣ����������粻�ȶ���ɣ�", vbRetryCancel, gstrSysName) = vbRetry Then
            Call FuncUploadFile(objFtp, strVirtualPath, strLocalFileName, strRemoteFileName)
        Else
            mblnIsFirstImgUpLoadFiald = True
            FuncUploadFile = False
            Exit Function
        End If
    End If
    
    FuncUploadFile = True
End Function

Private Function ConnFtp(Optional ByVal blnRefresh As Boolean = False) As Boolean
    If mobjFtp.hConnection = 0 Or blnRefresh Then
        '����FTP�洢�豸
        If mobjFtp.FuncFtpConnect(mobjFtpInf.strFtpIp, mobjFtpInf.strFTPUser, mobjFtpInf.strFTPPwd) = 0 Then
            MsgBox "FTP��������ʧ�ܣ����������FTP����." + vbCrLf + "FTP:" + mobjFtpInf.strFtpIp + "  USER:" + mobjFtpInf.strFTPUser, vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    
    ConnFtp = True
End Function

Private Function InitFtpInf(ByVal strDocID As String, Optional ByVal strDeviceNO As String, Optional ByVal strCreateDate As String) As Boolean
'��ȡ����ͼ���FTP��Ϣ
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    If mstrTmpReportId = strDocID Then
        InitFtpInf = True
        Exit Function
    End If
    
    mstrTmpReportId = strDocID
    
    If strDeviceNO = "" Then
        strSql = "Select 'ReportImages/' || to_Char(b.����ʱ��,'YYYYMMDD') || '/' || b.id || '/' As URL," & _
                 "a.�豸��, a.FTP�û���, a.FTP����, a.IP��ַ,'/'||a.FtpĿ¼||'/' As Root " & _
                 "From Ӱ���豸Ŀ¼ a, Ӱ�񱨸��¼ b where a.�豸�� = b.�豸�� And b.id = [1]"
                 
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡFTP��Ϣ", strDocID)
    Else
        strSql = "Select 'ReportImages/'||[1]||'/'||[2]||'/' As URL," & _
                 "�豸��, FTP�û���, FTP����, IP��ַ,'/'||FtpĿ¼||'/' As Root From Ӱ���豸Ŀ¼ where �豸�� = [3]"
                 
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡFTP��Ϣ", strCreateDate, strDocID, strDeviceNO)
    End If
    
    If rsTmp.RecordCount <= 0 Then Exit Function
    
    mobjFtpInf.strFtpDir = NVL(rsTmp("Root"))
    mobjFtpInf.strFtpIp = NVL(rsTmp("IP��ַ"))
    mobjFtpInf.strFTPPwd = NVL(rsTmp("FTP����"))
    mobjFtpInf.strFTPUser = NVL(rsTmp("FTP�û���"))
    mobjFtpInf.strDeviceId = Trim(NVL(rsTmp("�豸��")))
    mstrSubDir = NVL(rsTmp("URL"))
    mstrDestMainDir = IIf(Len(App.Path) > 3, App.Path & "\TmpReportImage\", App.Path & "TmpReportImage\")
    
    InitFtpInf = True
End Function

Public Function GetCurrDocId(ByVal lngOrderID As Long) As String
'��ȡ��ǰѡ�еı����Id
On Error Resume Next
    If mobjPacsInterface Is Nothing Then Exit Function
    GetCurrDocId = mobjPacsInterface.GetCurrDocId(lngOrderID)
End Function

Public Sub EvaluatReportQuality(ByVal strDocID As String, ByVal strResult As String, _
     Optional ByVal lngOrderID As Long = 0)
'--------------------------------------------------------
'��  �ܣ�������������
'������ strDocID -- ����ID
'       strResult -- ������������
'       lngOrderID -- ҽ��ID����ѡ���������strDocID=""����ͨ��ҽ��ID��ȡ����ID
'���أ���
'-------------------------------------------------------
    
    If mobjPacsInterface Is Nothing Then Exit Sub
    
    If Trim(strDocID) = "" And lngOrderID <> 0 Then
        strDocID = GetCurrDocId(lngOrderID)
    End If
    
    If Trim(strDocID) = "" Then Exit Sub
    Call mobjPacsInterface.EvaluatReportQuality(strDocID, strResult)
End Sub

Public Sub EvaluatResult(ByVal strDocID As String, ByVal strResult As String)
'����������
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.EvaluatResult(strDocID, strResult)
End Sub

Public Sub ReportRelease(ByVal lngOrderID As Long, ByVal intType As Integer)
'���淢�Ż����
'intType:0--������գ�1--���淢��
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.ReportRelease(mobjOwner.hWnd, lngOrderID, intType)
End Sub

Public Function GetReportReleaseState(ByVal lngOrderID As Long) As Integer
'���ر��汨�淢�����
'0-ȫ����δ���ţ�1-���ַ���(��ǰѡ�еı���δ����)��2-���ַ���(��ǰѡ�еı����ѷ���)��3-ȫ�����ѷ���
    If mobjPacsInterface Is Nothing Then Exit Function
    GetReportReleaseState = mobjPacsInterface.GetReportReleaseState(lngOrderID, mlngPatientId)
End Function

Public Function GetCurrReportState(ByVal lngOrderID As Long) As Integer
'���ص�ǰ����ı���״̬
    Dim strDocID As String
    
    If mobjPacsInterface Is Nothing Then Exit Function
    
    strDocID = Me.GetCurrDocId(lngOrderID)
    
    If Trim(strDocID) = "" Then Exit Function
    
    GetCurrReportState = mobjPacsInterface.GetCurrReportState(strDocID)
End Function

Public Sub UpdateImageVideoState(ByVal lngEventType As Integer, ByVal lngAdviceID As Long, ByVal other As Variant)
'ˢ��ǰ̨�ͺ�̨�ɼ���ͼ��
    
    If mobjPacsInterface Is Nothing Then Exit Sub
    
    Select Case lngEventType
        Case TVideoEventType.vetUpdateImg, TVideoEventType.vetCaptureFirstImg, TVideoEventType.vetDelAllImg, _
             TVideoEventType.vetAfterUpdateImg, TVideoEventType.vetExportImage, TVideoEventType.vetImportImage, TVideoEventType.vetImgDeled
            '����Id���գ�ҽ��IdΪ0��ʾˢ�º�̨�ɼ���ͼ��
            Call mobjPacsInterface.ExecutePlugInCmd(lngAdviceID, "", "ͼ��ˢ��", lngAdviceID)
        Case TVideoEventType.vetLockStudy
            Call mobjPacsInterface.ExecutePlugInCmd(0, "", "�������", "��" & other & "����Ƶ�ɼ�")
        Case TVideoEventType.vetUnLockStudy
            Call mobjPacsInterface.ExecutePlugInCmd(0, "", "�������", "��Ƶ�ɼ�")
    End Select
End Sub

Private Sub SetReportFinalShowMode(ByVal blnFinalShowMode As Boolean)
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.SetReportFinalShowMode(blnFinalShowMode)
End Sub

Private Sub SetIsNoAskPrint(ByVal blnNoAskPrint As Boolean)
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.SetIsNoAskPrint(blnNoAskPrint)
End Sub

'��ȡ��ǰ����Ĵ�����
Public Function GetDocCreator(ByVal lngOrderID As Long)
    Dim strDocID As String
    
    If mobjPacsInterface Is Nothing Then Exit Function
    
    strDocID = Me.GetCurrDocId(lngOrderID)
    
    If Trim(strDocID) = "" Then Exit Function
    
    GetDocCreator = mobjPacsInterface.GetDocCreator(strDocID)
End Function

'����ִ��
Public Sub SetDocCreator(ByVal strCreatorName)
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.SetRecordDoctor(strCreatorName)
End Sub

Public Property Get Visible() As Boolean
    Visible = mblnVisible
End Property

Public Sub RefreshVideo()
    If mobjPacsInterface Is Nothing Then Exit Sub
    
    Call mobjPacsInterface.ExecutePlugInCmd(0, "", "��Ƶˢ��", "")
End Sub

Public Sub AddXWReportImgToRichRPT(ByVal lngOrderID As Long, strDcmImgPath As String)
    '���ļ�·�����������ĵ��༭��������ͼ����У��ڲ�����ڶ�ͼ����в���
    If Trim(strDcmImgPath) = "" Then Exit Sub
    
    Call mobjPacsInterface.ExecutePlugInCmd(lngOrderID, "", "XWͼ������", strDcmImgPath)
End Sub