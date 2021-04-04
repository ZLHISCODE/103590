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

Private Sub Class_Initialize()
    Set frmInTend = New frmDockInTendMain
End Sub

Private Sub Class_Terminate()
    Unload frmInTend
    Set frmInTend = Nothing
    
    Set mclsUnzip = Nothing
    Set mclsZip = Nothing
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
    Optional ByVal blnDoctorStation As Boolean, Optional ByVal lngUnitID As Long, Optional ByVal intState As Integer) As Long
    '����:����ˢ��ָ�����˵Ĳ������ݣ�����������ṩ�༭����
    '����:  lngPatiId-����id;
    '       lngPageId-��ҳid;
    '       lngDeptId-��ǰ�������ţ�ע�ⲻ�ǲ��˵�ǰ���ڲ������Ժ������
    '       lngUnitID��lngDeptId-��intState=��5-���ת�Ʋ��ˡ�ʱΪ����ԭ������ԭ����
    '       blnEdit-�Ƿ������༭��ͨ����ǰ�������Ų��ǲ������ڲ������Ժ��������Ӧ�ò������༭��ҽ��վֱ������Ϊ�������༭
    '       intState-��ʿվ��ҽ��վ����Ĳ���״̬TYPE_PATI_State
    zlRefresh = frmInTend.zlRefresh(lngPatiID, lngPageId, lngDeptID, bnEdit, , blnDoctorStation)
End Function

Public Sub zlLocateData(ByVal intType As Integer)
'����:��λ��������¼��Ż����������ݽ���
'����:intType=0-������¼����,1-������������
    Call frmInTend.zlLocateData(intType)
End Sub

Public Sub TendFileMutilEditor(ByVal objParent As Object, ByVal lngDeptID As Long, ByVal strPrivs As String)
    Dim frmTendFileMutil As New frmTendFileMutilEditor
    
    Call frmTendFileMutil.ShowMe(objParent, lngDeptID, strPrivs)
End Sub

Public Sub InitTendFile(cnMain As ADODB.Connection, ByVal lngSys As Long)
    Set gcnOracle = cnMain
    Call InitCommon(gcnOracle)
        
    glngSys = lngSys
    gstrDbOwner = GetDbOwner(glngSys)
    gstrProductName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("��Ʒ����"), Default:="����")
    gstrAviPath = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    
    Call GetUserInfo
End Sub
