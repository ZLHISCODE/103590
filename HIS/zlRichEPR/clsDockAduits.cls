VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockAduits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
'�̳д����¼�
Private mfrmCaseTendBody As Object
Attribute mfrmCaseTendBody.VB_VarHelpID = -1
Private WithEvents mfrmTendFileOpen As frmTendFileOpen
Attribute mfrmTendFileOpen.VB_VarHelpID = -1
Private WithEvents mfrmDockAduitEPR As frmDockAduitEPR
Attribute mfrmDockAduitEPR.VB_VarHelpID = -1
Private mfrmEPRAuditMonitor As frmEPRAuditMonitor
Private WithEvents mfrmEPRAuditTime As frmEPRAuditTime
Attribute mfrmEPRAuditTime.VB_VarHelpID = -1

Private WithEvents mfrmCaseTendEditForSinglePerson As frmCaseTendEditForSinglePerson
Attribute mfrmCaseTendEditForSinglePerson.VB_VarHelpID = -1
Private mblnChange As Boolean
Private mlngCert As Long
Private mstrǩ������ As String              'ǩ������Ϊ��ʿ��˵������ǩ
Private mfrmMain As Object

Public Event ShowItemInfo(ByVal strInfo As String)
Public Event AfterDocumentChanged(ByVal lngEPRKey As Long)
Public Event AfterEprPrint(ByVal lngRecordId As Long)
Public Event AfterTendPrint(ByVal lngFileID As Long)
'######################################################################################################################

Private Sub Class_initialize()
    If Not CreateBodyEditor Then Exit Sub
    Set mfrmCaseTendBody = gobjBodyEditor.GetNewTendBody
    Set mfrmTendFileOpen = New frmTendFileOpen
    Set mfrmDockAduitEPR = New frmDockAduitEPR
    Set mfrmCaseTendEditForSinglePerson = New frmCaseTendEditForSinglePerson
    Call mfrmCaseTendEditForSinglePerson.SetEditable(True)
    
    Set mfrmEPRAuditMonitor = New frmEPRAuditMonitor
    Set mfrmEPRAuditTime = New frmEPRAuditTime
    
    Call mfrmCaseTendBody.zlInit
End Sub

Private Sub Class_Terminate()
    If Not (mfrmCaseTendBody Is Nothing) Then Unload mfrmCaseTendBody
    If Not (mfrmTendFileOpen Is Nothing) Then Unload mfrmTendFileOpen
    If Not (mfrmCaseTendEditForSinglePerson Is Nothing) Then Unload mfrmCaseTendEditForSinglePerson
    If Not (mfrmDockAduitEPR Is Nothing) Then Unload mfrmDockAduitEPR
    If Not (mfrmEPRAuditMonitor Is Nothing) Then Unload mfrmEPRAuditMonitor
    If Not (mfrmEPRAuditTime Is Nothing) Then Unload mfrmEPRAuditTime
    
    Set mfrmCaseTendBody = Nothing
    Set mfrmTendFileOpen = Nothing
    Set mfrmCaseTendEditForSinglePerson = Nothing
    Set mfrmDockAduitEPR = Nothing
    Set mfrmEPRAuditMonitor = Nothing
    Set mfrmEPRAuditTime = Nothing
    Set mfrmMain = Nothing
End Sub

'######################################################################################################################

Public Function zlGetFormEPR() As Object
    Set zlGetFormEPR = mfrmDockAduitEPR
End Function

Public Function zlGetFormTendFile() As Object
    Set zlGetFormTendFile = mfrmTendFileOpen
End Function

Public Function zlGetFormTendBody() As Object
    Set zlGetFormTendBody = mfrmCaseTendBody
End Function

Public Function zlGetFormTendEdit() As Object
    Set zlGetFormTendEdit = mfrmCaseTendEditForSinglePerson
End Function

Public Function zlGetFormAuditMonitor() As Object
    Set zlGetFormAuditMonitor = mfrmEPRAuditMonitor
End Function

Public Function zlGetFormAuditTimeLimit() As Object
    Set zlGetFormAuditTimeLimit = mfrmEPRAuditTime
End Function

Public Function zlDataChange() As Boolean
    '��ȡ�����Ƿ��޸�
    zlDataChange = mblnChange
End Function

Public Function zlIsCert() As Boolean
    '��ȡǩ�����
    zlIsCert = (mlngCert <> 0)
End Function

Public Function zlIsInspect() As Boolean
    '��ȡ��ǩ���
    zlIsInspect = (mlngCert <> 0) And (mstrǩ������ = "��ʿ��")
End Function

Public Function zlIsPigeonhole() As Boolean
    '��ȡ�鵵���
    zlIsPigeonhole = mfrmCaseTendEditForSinglePerson.IsPigeonhole
End Function

Public Sub zlInitMonitor(ByVal frmMain As Object)
    Call zlGetFormAuditMonitor.zlInitData(frmMain)
End Sub

Public Sub zlInitTime(ByVal frmMain As Object)
    Call mfrmEPRAuditTime.zlInitData(frmMain)
End Sub

Public Sub zlClearMonitor()
    Call zlGetFormAuditMonitor.zlClearData
End Sub

Public Sub zlClearTime()
    Call mfrmEPRAuditTime.zlClearData
End Sub

Public Function zlRefreshMonitor(ByVal lngEPRid As Long) As Boolean
    Call zlGetFormAuditMonitor.zlRefreshData(lngEPRid)
End Function

Public Function zlRefreshTime(ByVal lngPatientKey As Long, ByVal lngPatientPageKey As Long, ByVal intKind As Integer) As Boolean
    Call mfrmEPRAuditTime.zlRefreshData(lngPatientKey, lngPatientPageKey, intKind)
End Function

Public Function zlRefresh(ByVal bytKind As Byte, ByVal lngKey As Long, Optional ByVal lngPatiID As Long, _
Optional ByVal lngPageId As Long, Optional ByVal lngDeptId As Long, Optional ByVal strPeriod As String, _
Optional ByVal byt�������� As Byte = 3, Optional ByVal intBaby As Integer = 0, Optional ByVal blnDataMoved As Boolean) As Long
    '����:  ��ʾָ���ĵ��Ӳ���
    '����:  lngKey-���Ӳ�����ID;
    
    Select Case bytKind
    Case 3      '������¼�ļ�
        Call mfrmTendFileOpen.ShowMe(Nothing, lngKey, lngPatiID, lngPageId, lngDeptId, intBaby, strPeriod, True, byt��������, blnDataMoved)
    Case Else
        Call mfrmDockAduitEPR.zlRefresh(mfrmMain, lngKey, blnDataMoved)
    End Select

End Function

Public Sub zlOpenEPRDocument(ByVal lngKey As Long, ByVal frmMain As Object)
    '����:  ��ʾָ���ĵ��Ӳ���
    '����:  lngKey-���Ӳ�����ID;
    
    Dim frm As New frmEPRView
    frm.ShowMe frmMain, lngKey
    
End Sub


Public Function ParentForm(ByVal frmMain As Object) As Boolean

    Set mfrmMain = frmMain
    
End Function

Public Function zlRefreshTendBody(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, Optional ByVal intBaby As Integer = 0, Optional ByVal blnDataMove As Boolean) As Long
    '����:  ��ʾ���µ�
        
    Call mfrmCaseTendBody.zlRefresh(lngPatiID & ";" & lngPageId & ";" & lngDeptId & ";0;0;" & intBaby)
    
End Function

Public Sub SignMarker()
    Call mfrmCaseTendEditForSinglePerson.SignMarker
End Sub

Public Function CancelMe() As Boolean
    CancelMe = mfrmCaseTendEditForSinglePerson.CancelMe()
End Function

Public Function zlRefreshTendEdit(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, _
    Optional ByVal byt�������� As Byte = 3, Optional ByVal intBaby As Integer = 0, Optional ByVal strPrivs As String, _
    Optional ByVal blnCancel As Boolean = False, Optional ByVal blnEditable As Boolean = True) As Long
    '����:  ��ʾ�����˿���¼�뵥
    Call mfrmCaseTendEditForSinglePerson.ShowMe(mfrmMain, lngPatiID, lngPageId, lngDeptId, intBaby, byt��������, strPrivs, blnCancel, blnEditable)
    
End Function

Public Function zlPrintDocument(ByVal bytKind As Byte, Optional ByVal bytMode As Byte = 2, Optional ByVal lngKey As Long, Optional ByVal strPrintDeviceName As String, Optional ByVal blnDataMove As Boolean) As Long
'����:��ӡ��������'���µ�=1��������¼��=2������=3�����Ʊ���=4
'����:bytKind
'    bytMode:1-Ԥ�� 2-��ӡ 
'    lngKey:�ؼ�ID
'    strPrintDeviceName ��bytMode=1��2ʱ��ʾ��ӡ���� 

    Dim strSQL As String
    
    If bytKind = 1 Then
        '���µ�(����ֵ:1-�ɹ�;2-��ӡ)
        zlPrintDocument = mfrmCaseTendBody.zlPrintBody(bytMode, strPrintDeviceName)
    ElseIf bytKind = 2 Then
        '������¼��
        Call mfrmTendFileOpen.zlPrintTend(bytMode, strPrintDeviceName)
    ElseIf bytKind = 3 Then
        Call mfrmDockAduitEPR.zlPrintDocument(cprסԺ����, bytMode, lngKey, strPrintDeviceName, blnDataMove)
    Else
        Call mfrmDockAduitEPR.zlPrintDocument(cpr���Ʊ���, bytMode, lngKey, strPrintDeviceName, blnDataMove)
    End If
    
End Function

Private Sub mfrmCaseTendEditForSinglePerson_AfterDataChanged()
    mblnChange = True
End Sub

Private Sub mfrmCaseTendEditForSinglePerson_AfterRefresh()
    mblnChange = False
End Sub

Private Sub mfrmCaseTendEditForSinglePerson_AfterRowColChange(ByVal strInfo As String)
    RaiseEvent ShowItemInfo(strInfo)
End Sub

Private Sub mfrmCaseTendEditForSinglePerson_AfterSelChange(ByVal lngCert As Long, ByVal strCertLevel As String)
    mlngCert = lngCert
    mstrǩ������ = strCertLevel
End Sub

Private Sub mfrmDockAduitEPR_PrintEpr(ByVal lngRecordId As Long)
    RaiseEvent AfterEprPrint(lngRecordId)
End Sub

Private Sub mfrmEPRAuditTime_AfterDocumentChanged(ByVal lngEPRKey As Long)
    RaiseEvent AfterDocumentChanged(lngEPRKey)
End Sub

Private Sub mfrmTendFileOpen_zlAfterPrint(ByVal lngFileID As Long)
    RaiseEvent AfterTendPrint(lngFileID)
End Sub

Public Sub SetFontSize(ByVal bytKind As Byte, ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytKind��0-����¼�룬1-���µ���2-��¼��;bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-18 15:16
    '����:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Select Case bytKind
        Case 0 ' ����¼��
            Call mfrmCaseTendEditForSinglePerson.SetFontSize(bytSize)
        Case 1 '���µ�
            'Ŀǰֻ�б�׼���µ���
            On Error Resume Next
            Call mfrmCaseTendBody.SetFontSize(bytSize)
            If Err <> 0 Then Err.Clear
        Case 2 '��¼��
            Call mfrmTendFileOpen.SetFontSize(bytSize)
    End Select
End Sub