VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�̳д����¼�
Public Event Activate()
Public Event RefreshPrompt(ByVal strInfo As String, ByVal blnImportant As Boolean)

Private WithEvents frmInTend As frmDockInTendMain
Attribute frmInTend.VB_VarHelpID = -1
Private WithEvents frmInTendFile As frmDockInTend_File
Attribute frmInTendFile.VB_VarHelpID = -1

Private Sub Class_Initialize()
    mlngInitClsCount = mlngInitClsCount + 1
    Set frmInTend = New frmDockInTendMain
    Set frmInTendFile = New frmDockInTend_File
End Sub

Private Sub Class_Terminate()
    mlngInitClsCount = mlngInitClsCount - 1
    If Not frmInTend Is Nothing Then
        Unload frmInTend
        Set frmInTend = Nothing
    End If
    If Not frmInTendFile Is Nothing Then
        Unload frmInTendFile
        Set frmInTendFile = Nothing
    End If
    
    If mlngInitClsCount <= 0 Then
        If Not gobjBodyEditor Is Nothing Then Set gobjBodyEditor = Nothing
        If Not gobjPartogram Is Nothing Then Set gobjPartogram = Nothing
    End If
    If Not mclsUnzip Is Nothing Then Set mclsUnzip = Nothing
    If Not mclsZip Is Nothing Then Set mclsZip = Nothing
End Sub

Private Sub frmInTend_Activate()
    RaiseEvent Activate
End Sub

Private Sub frmInTend_RefreshPrompt(ByVal strInfo As String, ByVal blnImportant As Boolean)
    RaiseEvent RefreshPrompt(strInfo, blnImportant)
End Sub

Public Function zlGetForm() As Object
    Set zlGetForm = frmInTend
End Function

Public Function zlGetfrmInTendFile() As Object
    Set zlGetfrmInTendFile = frmInTendFile
End Function

Public Sub zlDefCommandBars(ByVal CommandBars As Object, Optional ByVal blnChildToolBar As Boolean = False)
    Call frmInTend.zlDefCommandBars(CommandBars, blnChildToolBar)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInTend.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInTend.zlExecuteCommandBars(Control)
End Sub

Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, ByVal bnEdit As Boolean, _
    Optional ByVal blnDoctorStation As Boolean, Optional ByVal lngUnitId As Long, Optional ByVal intState As Integer, Optional ByVal intCurveReSize As Integer = 0) As Long
    '����:����ˢ��ָ�����˵Ĳ������ݣ�����������ṩ�༭����
    '����:  lngPatiId-����id;
    '       lngPageId-��ҳid;
    '       lngDeptId-��ǰ�������ţ�ע�ⲻ�ǲ��˵�ǰ���ڲ������Ժ������
    '       lngUnitID��lngDeptId-��intState=��5-���ת�Ʋ��ˡ�ʱΪ����ԭ������ԭ����
    '       blnEdit-�Ƿ������༭��ͨ����ǰ�������Ų��ǲ������ڲ������Ժ��������Ӧ�ò������༭��ҽ��վֱ������Ϊ�������༭
    '       intState-��ʿվ��ҽ��վ����Ĳ���״̬TYPE_PATI_State
    '       intCurveReSize ���µ��������Ƿ�Ϊ����ģʽ 0����ģʽ 1ԭʼ��С
    '56965:������,2012-12-24
    glng����ID = IIf(blnDoctorStation = True, lngUnitId, lngDeptID)
    zlRefresh = frmInTend.zlRefresh(lngPatiID, lngPageId, glng����ID, bnEdit, , blnDoctorStation, , intCurveReSize)
End Function

Public Function GetFormOperation() As String
'��¼����ѡ����Ϣ����Ϊ����վ���л�ҳ��ʱ���ͷ��˶��󣬻�����ʱ���³�ʼ��ˢ�µġ�
    GetFormOperation = frmInTend.GetFormOperation
End Function

Public Sub RestoreFormOperation(ByVal strValue As String)
'�ָ�����ѡ����Ϣ������վ��ˢ��֮ǰ����
    Call frmInTend.RestoreFormOperation(strValue)
End Sub

Public Function zlRefreshTendFile(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer, ByVal lngDeptID As Long, ByVal blnEdit As Boolean, _
    Optional ByVal blnDoctorStation As Boolean, Optional ByVal intSEL As Integer, Optional ByVal lngKey As Long, Optional ByVal intCurveReSize As Integer = 0) As Long
    'bytSel:0-���µ�;1-��¼��;2 ����ͼ
    Call frmInTendFile.zlRefresh(lngPatiID, lngPageId, intBaby, lngDeptID, blnEdit, blnDoctorStation, intSEL, lngKey, intCurveReSize)
End Function

Public Function zlPrintTendFile(ByVal bytKind As Byte, Optional ByVal bytMode As Byte = 2, Optional ByVal strPrintDevice As String = "") As Long
    '��ɼ�¼�������µ�������ͼ��Ԥ����ӡ
    '����:bytkind  1-���µ� 2-��¼��  3-����ͼ
    '     bytMode 1 Ԥ�� 2��ӡ 3�����Excel
    '     strPrintDevice ��ӡ�����ƣ���Ϊ��ʱ��ӡ�����������ݣ�Ϊ��ʱ����ѯ�ʴ���
    Call frmInTendFile.zlPrintTendFile(bytKind, bytMode, strPrintDevice)
End Function

Public Sub zlPrintDocument(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer, ByVal lngDeptID As Long, ByVal lngFileID As Long, ByVal bytKind As Byte, ByVal strPrintDeviceName As String, Optional ByVal blnPrint As Boolean = True)
    '����:��ɼ�¼�������µ�������ͼ���(��ֱ��ʹ��)
    '����:lngPatiID-����ID
    '     lngPageId-��ҳID
    '     intBaby-Ӥ�����
    '     lngDeptID-����ID
    '     lngFileID-�ļ�ID
    '     bytKind:1-���µ�:2-��¼�� 3-����ͼ
    '     strPrintDeviceName:��ӡ������
    '     blnPrint:TRUE ��ӡ��FalseԤ��
    Call frmInTendFile.zlPrintDocument(lngPatiID, lngPageId, intBaby, lngDeptID, lngFileID, bytKind, strPrintDeviceName, blnPrint)
End Sub

Public Sub zlLocateData(ByVal intType As Integer)
'����:��λ��������¼��Ż����������ݽ���
'����:intType=0-������¼����,1-������������
    Call frmInTend.zlLocateData(intType)
End Sub

Public Sub TendFileMutilEditor(ByVal objParent As Object, ByVal lngDeptID As Long, ByVal strPrivs As String, Optional ByVal bytSize As Byte = 0)
    Dim frmTendFileMutil As New frmTendFileMutilEditor
    
    Call frmTendFileMutil.ShowMe(objParent, lngDeptID, strPrivs, bytSize)
End Sub

Public Sub InitTendFile(cnMain As ADODB.Connection, ByVal lngSys As Long)
    Set gcnOracle = cnMain
    glngSys = lngSys
    gstrDbOwner = GetDbOwner(glngSys)
    gstrProductName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("��Ʒ����"), Default:="����")
    gstrAviPath = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    
    Call GetUserInfo
End Sub

'Public Function zlTendPara(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
'    '��ɼ�¼����������
'    '����:frmMain  ���ô���
'    '     strPrivs Ȩ��
'    zlTendPara = frmTendPara.ShowPara(frmMain, strPrivs)
'End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '���������Ƿ��ṩ�ļ���Ϣ�б�
    '����:������
    '����:2012-06-19 15:16
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If frmInTend Is Nothing Then Exit Sub
    Call frmInTend.SetFontSize(bytSize)
End Sub

Public Sub UnloadClass()
'�°���Ӳ���ʹ�ã����°���Ӳ����޷���ж�ض���ʱ����Class_Terminate
    Call Class_Terminate
End Sub