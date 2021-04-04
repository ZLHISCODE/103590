VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRichEPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'################################################################################################################
'##ģ �� ����czlRichEPR.cls
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��8��11��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ����ϵͳ�����
'################################################################################################################

Option Explicit

'################################################################################################################
'## ���ܣ�  ������ʼ��
'##
'## ˵����  ���й�������Ӧ�ڱ�������ģ������ɶ���
'################################################################################################################
Public Sub InitRichEPR(cnMain As ADODB.Connection, frmParent As Object, ByVal lngSys As Long, ByVal blnShowInTaskBar As Boolean)
    If gfrmPublic Is Nothing Then Set gfrmPublic = New frmPublic
    
    Set gfrmParent = frmParent
    Set gcnOracle = cnMain
    'Call InitCommon(gcnOracle)
        
    glngSys = lngSys
    gstrDbOwner = GetDbOwner(glngSys)
    gblnShowInTaskBar = blnShowInTaskBar
    gstrProductName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("��Ʒ����"), Default:="����")
    gstrAviPath = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gstrMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
    gstrPrivsEpr = ";" & GetPrivFunc(glngSys, 1070) & ";"
    
    Call GetUserInfo
    Call InitPreDefinedKeys
    
    On Error Resume Next
    If Not frmParent Is Nothing Then
        Set gobjEmr = frmParent.mobjEMR  '�²�����Ҫ����
        If Err.Number <> 0 Then
            Set gobjEmr = Nothing
            Err.Clear
        End If
    End If
      
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, glngSys, 1070)
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub

'################################################################################################################
'## ���ܣ�  �ⲿֱ��ģ�����
'## ������  lngSys      :ϵͳ���
'##         lngModul    :ģ����
'##         frmParent   :������
'################################################################################################################
Public Sub zlExecuteCode(ByVal lngSys As Long, ByVal lngModul As Long, ByVal frmParent As Variant, Optional strPrivs As String)
    glngSys = lngSys: glngModul = lngModul
    gstrDbOwner = GetDbOwner(glngSys)
    
    If strPrivs = "" Then
        gstrPrivs = GetPrivFunc(glngSys, glngModul)
    Else
        gstrPrivs = strPrivs
    End If
    
    Select Case lngModul
    Case 1071 '�������ͼ�ι���
        If IsObject(frmParent) Then
            frmMarkMapMan.Show , frmParent
        Else
            'ZLBH�ںϵ���
            zlCommFun.ShowChildWindow frmMarkMapMan.hWnd, frmParent
        End If
    Case 1072 '������¼��Ŀ����
        If IsObject(frmParent) Then
            frmTendItemMan.Show , frmParent
        Else
            'ZLBH�ںϵ���
            zlCommFun.ShowChildWindow frmTendItemMan.hWnd, frmParent
        End If
    Case 1073 '����Ԥ����ٹ���
        If IsObject(frmParent) Then
            frmSentenceMan.Show , frmParent
        Else
            'ZLBH�ںϵ���
            zlCommFun.ShowChildWindow frmSentenceMan.hWnd, frmParent
        End If
    Case 1074 '�����ļ�����
        If IsObject(frmParent) Then
            frmEPRFileMan.Show , frmParent
        Else
            'ZLBH�ںϵ���
            zlCommFun.ShowChildWindow frmEPRFileMan.hWnd, frmParent
        End If
    Case 1075 '�������Ĺ���
        If IsObject(frmParent) Then
            frmEPRModelMan.Show , frmParent
        Else
            'ZLBH�ںϵ���
            zlCommFun.ShowChildWindow frmEPRModelMan.hWnd, frmParent
        End If
    Case 1076   '���Ƶ�������
        If IsObject(frmParent) Then
            frmEPRBillMan.Show , frmParent
        Else
            'ZLBH�ںϵ���
            zlCommFun.ShowChildWindow frmEPRBillMan.hWnd, frmParent
        End If
    Case 1950   '�ƶ���ʿվ��������
        If IsObject(frmParent) Then
            frmTendItemTransfusion.Show , frmParent
        Else
            'ZLBH�ںϵ���
            zlCommFun.ShowChildWindow frmTendItemTransfusion.hWnd, frmParent
        End If
    Case 1273 '���˲�������
        If IsObject(frmParent) Then
            Dim objForm As Object
            Set objForm = frmParent 'ֱ�Ӹ�ֵ������
            frmEPRSearchMan.ShowSearchClinic objForm, glngDeptId
        Else
            'ZLBH�ںϵ���
            frmEPRSearchMan.ShowSearchClinic Nothing, glngDeptId
        End If
    Case 1274 '�����걨����
        If IsObject(frmParent) Then
            frmDiseaseReportMan.Show , frmParent
        Else
            'ZLBH�ںϵ���
            zlCommFun.ShowChildWindow frmDiseaseReportMan.hWnd, frmParent
        End If
    End Select
End Sub

'################################################################################################################
'## ���ܣ�  �ر����д���
'################################################################################################################
Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    If Not gfrmPublic Is Nothing Then Unload gfrmPublic
    Set gfrmPublic = Nothing
    CloseWindows = (Forms.Count = 0)
End Function

Public Sub PrintOrPreviewDoc(frmParent As Object, eDocType As EPRDocTypeEnum, lngRecId As Long, _
    Optional blnPrint As Boolean, Optional blnFixLastVersion As Boolean = False)
'��ʽ��ӡ��Ԥ���Ѿ������ָ���ĵ�
Dim rsTemp As ADODB.Recordset
    gstrSQL = "Select �༭��ʽ From ���Ӳ�����¼ where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, App.Title, lngRecId)
    If rsTemp.EOF Then Exit Sub
    
    If rsTemp!�༭��ʽ = 0 Then
        Dim frmP As frmPrintPreview, blnOrigMode As Boolean, strAsk As String
        If lngRecId = 0 Then Exit Sub
        
        If blnFixLastVersion Then
            blnOrigMode = False
        Else
            strAsk = "���԰�����ʽ��ԭʼ��ʽ" & IIf(blnPrint, "��ӡ", "Ԥ��") & "��"
            strAsk = strAsk & vbCrLf & "    ���ո�ʽ���������޸ĺۼ�������ʽ"
            strAsk = strAsk & vbCrLf & "    ԭʼ��ʽ�������޸ĺۼ��Ĳݸ��ʽ"
            strAsk = strAsk & vbCrLf & "�������ո�ʽ��ģʽ" & IIf(blnPrint, "��ӡ", "Ԥ��") & "��"
            
            Select Case MsgBox(strAsk, vbYesNoCancel + vbQuestion, gstrSysName)
            Case vbYes
                blnOrigMode = False
            Case vbNo
                blnOrigMode = True
            Case Else
                Exit Sub
            End Select
        End If
        
        Set frmP = New frmPrintPreview
        frmP.DoMultiDocPreview frmParent, eDocType, , , , , lngRecId, blnPrint, blnOrigMode
        Unload frmP
        Set frmP = Nothing
    Else
        Dim objTableEpr As New cTableEPR
        objTableEpr.InitTableEPR gcnOracle, glngSys, gstrDbOwner
        objTableEpr.InitOpenEPR frmParent, TabEm_�޸�, TabET_���������, lngRecId, False, 2
        objTableEpr.zlPrintDoc frmParent, Not blnPrint
        Set objTableEpr = Nothing
    End If
End Sub

Public Sub ViewDocument(frmParent As Object, lngRecId As Long, Optional blnCanPrint As Boolean)
'�򿪲����ĵ�����
Dim rsTemp As ADODB.Recordset
    gstrSQL = "Select �༭��ʽ From ���Ӳ�����¼ where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, App.Title, lngRecId)
    If rsTemp.EOF Then Exit Sub
    
    If rsTemp!�༭��ʽ = 0 Then
        Dim fViewDoc As New frmEPRView
        fViewDoc.ShowMe frmParent, lngRecId, , blnCanPrint
    Else
        Dim objTableEpr As New cTableEPR
        objTableEpr.InitTableEPR gcnOracle, glngSys, gstrDbOwner
        objTableEpr.InitOpenEPR frmParent, TabEm_�޸�, TabET_���������, lngRecId, True, 2
        Set objTableEpr = Nothing
    End If
End Sub


Public Function GetDocumentText(lngRecId As Long) As String
'���ܣ����ݲ��˲���id��ȡ�����ļ����ı����ݡ�
'������lngRecId-���Ӳ����ļ�ID
    Dim frmContent As frmDockEPRContent, strTmp As String
    
    strTmp = Clipboard.GetText
    
    Set frmContent = New frmDockEPRContent
    Call frmContent.zlRefresh(lngRecId, "", False)
    frmContent.edtThis.SelectAll
    frmContent.edtThis.Copy
    GetDocumentText = Clipboard.GetText
        
    Clipboard.SetText strTmp
End Function
Public Sub EditDocument(ByVal lngModul As Long, ByVal frmParent As Object, ByVal lngDeptId As Long, ByVal strFileID As String)
'lngModul �ĸ�ģ�����
'frmParent ������
'lngDeptID ��������,���ǲ��˿���,Ҳ�ǲ���Ա����,�����ڷ�ȱʡ������д�Ǳ��ƻ����¼
'strFileID �ļ�ID,Ԥ����չ��string
Dim rsTemp As New ADODB.Recordset, strEprPrivs As String, lngFileID As Long
Dim objDoc As New cEPRDocument, objTabEpr As Object
    On Error GoTo errHand
    lngFileID = CLng(strFileID)
    gstrSQL = "Select ����id, ��ҳid, ������Դ, Ӥ��, ��������, ����id, �ļ�id, ������, ������, �鵵��, ���ʱ��, Nvl(�༭��ʽ, 0) �༭��ʽ" & vbNewLine & _
                "From ���Ӳ�����¼" & vbNewLine & _
                "Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", lngFileID)
    If rsTemp.EOF Then Exit Sub
            
    Select Case rsTemp!��������
        Case cpr��������
            strEprPrivs = GetPrivFunc(glngSys, 1255)
        Case cprסԺ����, cpr֪���ļ�, cpr�������, cpr���Ʊ���
            strEprPrivs = GetPrivFunc(glngSys, 1251)
        Case cpr���ﲡ��
            strEprPrivs = GetPrivFunc(glngSys, 1250)
    End Select
    
    If InStr(strEprPrivs, IIf(rsTemp!�������� = cpr��������, "����������д", "������д")) = 0 Then ViewDocument frmParent, lngFileID: Exit Sub
    If NVL(rsTemp!���ʱ��) = "" And lngDeptId <> rsTemp!����ID And rsTemp!������ <> gstrUserName Then ViewDocument frmParent, lngFileID: Exit Sub '�Ǳ��Ʋ����������޸�(��д��,���Һϲ�,��Ա���Ҵ������Ӧ�����޸�)
    If NVL(rsTemp!���ʱ��) = "" And InStr(1, strEprPrivs, IIf(rsTemp!�������� = cpr��������, "���˻�������", "���˲���")) = 0 And rsTemp!������ <> gstrUserName Then ViewDocument frmParent, lngFileID: Exit Sub 'δǩ���������˲���Ȩ�ޣ��Ǳ��˴����������޸�
    If NVL(rsTemp!���ʱ��) <> "" And InStr(1, strEprPrivs, IIf(rsTemp!�������� = cpr��������, "������������", "��������")) = 0 Then ViewDocument frmParent, lngFileID: Exit Sub  '��ǩ���������˲���Ȩ�ޣ�������˱༭�������޸�
'    If NVL(rsTemp!�鵵��, "") <> "" Then ViewDocument frmParent, lngFileID: Exit Sub '�ѹ鵵�������༭
    
    If rsTemp!�༭��ʽ = 0 Then 'ȫ��ʽ�༭
        objDoc.InitEPRDoc cprEM_�޸�, IIf(rsTemp!���ʱ�� = "", cprET_�������༭, cprET_���������), lngFileID, NVL(rsTemp!������Դ, 2), rsTemp!����ID, rsTemp!��ҳID, NVL(rsTemp!Ӥ��, 0), lngDeptId
        objDoc.ShowEPREditor frmParent
    Else '������
        Set objTabEpr = DynamicCreate("zlTableEPR.cTableEPR", "������")
        objTabEpr.InitTableEPR gcnOracle, glngSys, gstrDbOwner
        objTabEpr.InitOpenEPR frmParent, cprEM_�޸�, IIf(rsTemp!���ʱ�� = "", cprET_�������༭, cprET_���������), lngFileID, True, 0, NVL(rsTemp!������Դ, 2), rsTemp!����ID, rsTemp!��ҳID, NVL(rsTemp!Ӥ��, 0), lngDeptId, , strEprPrivs, False, False
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Unload frmPublic
    If Not gfrmPublic Is Nothing Then Unload gfrmPublic
    Set gfrmPublic = Nothing
    Set gobjFSO = Nothing
    Set gcnOracle = Nothing
    Set gfrmParent = Nothing
    Set gobjPacsCore = Nothing
    Set gobjESign = Nothing
    Set gobjEmr = Nothing
End Sub