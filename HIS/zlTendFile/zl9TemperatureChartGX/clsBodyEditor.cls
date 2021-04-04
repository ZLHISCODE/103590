VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBodyEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mfrmCaseTendBody As frmCaseTendBody
Attribute mfrmCaseTendBody.VB_VarHelpID = -1

Private mfrmPrintSet As Object
Public Event zlAfterPrint()

Public Function GetTendBody() As Object
    Set GetTendBody = frmCaseTendBody
End Function

Public Function GetNewTendBody() As Object
    Set GetNewTendBody = New frmCaseTendBody
End Function

Public Sub InitBodyEditor(ByVal lngSys As Long, ByVal cnOracle As ADODB.Connection)
    glngSys = lngSys
    Set gcnOracle = cnOracle
    
    gstrProductName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("��Ʒ����"), Default:="����")
    gstrAviPath = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gstrMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
    
    Call GetUserInfo
End Sub

Public Function GetCaseTendBodyPara() As Object
    Set GetCaseTendBodyPara = frmCaseTendBodyPara
End Function

Private Sub Class_Initialize()
    Set mfrmCaseTendBody = New frmCaseTendBody
    Set mfrmPrintSet = New frmPrintSet
End Sub

Private Sub Class_Terminate()
    If Not (mfrmCaseTendBody Is Nothing) Then Unload mfrmCaseTendBody
    If Not (mfrmPrintSet Is Nothing) Then Unload mfrmPrintSet
    Set mfrmCaseTendBody = Nothing
    Set mfrmPrintSet = Nothing
    
    If Not (mrsTabTime Is Nothing) Then Set mrsTabTime = Nothing
    If Not (mrsCollect Is Nothing) Then Set mrsCollect = Nothing
    If Not (mrsWave Is Nothing) Then Set mrsWave = Nothing
    If Not (mrsCurInfo Is Nothing) Then Set mrsCurInfo = Nothing
End Sub

Private Sub mfrmCaseTendBody_AfterPrint()
    RaiseEvent zlAfterPrint
End Sub

Public Sub zlPrintSet(ByVal frmParent As Object, Optional ByVal bytMode As Byte = 1)
'-----------------------------------------------------------------------------------------------------------------------------
'����:���µ���ӡ����
'-----------------------------------------------------------------------------------------------------------------------------
    Call mfrmPrintSet.ShowMe(frmParent, bytMode)
End Sub

Public Sub BodyMutilEditor(ByVal objParent As Object, ByVal lngDeptID As Long, ByVal strPrivs As String, Optional ByVal bytSize As Byte = 0)
'�������µ�����¼�봰��
    Dim frmBodyFileMutil As New frmTendBodyMutilEditor

    Call frmBodyFileMutil.ShowMe(objParent, lngDeptID, strPrivs, bytSize)
End Sub

Public Function zlCurvePrint(ByVal intPrintRange As Integer, ByVal blnPrint As Boolean, Optional lngBeginY As Long, _
    Optional ByVal intPageNo As Integer = -1, Optional ByVal strPrintDevice As String, Optional strPage As String, Optional strParam As String = "") As Boolean
    '---------------------------------------------------------------------------
    '����:���µ���ӡ
    'strPage intPrintRange=1 �����ṩ��ӡ��ʼҳ��ҳ����ҳ�� ��ʽΪ 0;1 ��ʾ��ʼҳ��1 ����ҳ�� 2
    'strParam �´������ʱҪ���� �ļ�ID;����ID;��ҳID;����ID;Ӥ��
    '---------------------------------------------------------------------------
    zlCurvePrint = mfrmCaseTendBody.BodyEdit.PrintState(intPrintRange, blnPrint, lngBeginY, intPageNo, strPrintDevice, strPage, strParam)
End Function