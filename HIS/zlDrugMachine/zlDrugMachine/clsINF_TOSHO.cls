VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINF_TOSHO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'˵�����������ӿڣ�TOSHO����
'���ߣ�YuZhiYong
'���䣺yzyong@zlsoft.cn
'------------------------------------------------------------------------------

Implements clsINF

Private mobjBase As New clsINF
Private mcnThird As ADODB.Connection

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Set mobjBase = Nothing
    Set mcnThird = Nothing
End Sub

Private Property Let clsINF_Action(ByVal RHS As Boolean)
'˵���������ӿ�Ϊ�״̬
    mobjBase.Action = RHS
    
    On Error GoTo hErr
    
    If RHS Then
        If mcnThird Is Nothing Then
            mobjBase.Action = False
            mobjBase.mobjLog.Add "����TOSHO������ʧ��", 1
            Exit Property
        End If
        
        mobjBase.mobjLog.Add "��ʼ����TOSHO������", 1
        mcnThird.Open mobjBase.URL
        mobjBase.mobjLog.Add "�ɹ�����TOSHO������", 1
    Else
        Set mcnThird = Nothing
    End If
    
    mobjBase.mobjLog.Save
    Exit Property
    
hErr:
    mobjBase.Action = False
    mobjBase.mobjLog.Add Err.Number & "��" & Err.Description, 1
    mobjBase.mobjLog.Add "����TOSHO�������쳣", 1
    mobjBase.mobjLog.Save
End Property

Private Property Get clsINF_Action() As Boolean
    clsINF_Action = mobjBase.Action
End Property

Private Function clsINF_FilterData(ByVal strCode As String, ByVal intType As Integer, ByVal rsData As ADODB.Recordset, rsFinish As ADODB.Recordset) As Boolean
    clsINF_FilterData = mobjBase.FilterData(strCode, intType, rsData, rsFinish)
End Function

Private Function clsINF_GetData(ByVal intType As Integer, ByVal strData As String, blnFinish As Boolean) As ADODB.Recordset
    Set clsINF_GetData = mobjBase.GetData(intType, strData, blnFinish)
    
    '�����ʹ�û���ļ�¼�������뽫ǰ��������Σ����������д�µĴ���
    'TODO��
End Function

Private Property Let clsINF_KEY(ByVal RHS As String)
    mobjBase.KEY = RHS
End Property

Private Property Get clsINF_KEY() As String
    clsINF_KEY = mobjBase.KEY
End Property

Private Property Set clsINF_mobjComLib(ByVal RHS As Object)
    Set mobjBase.mobjComLib = RHS
End Property

Private Property Get clsINF_mobjComLib() As Object
    Set clsINF_mobjComLib = mobjBase.mobjComLib
End Property

Private Property Set clsINF_mobjLog(ByVal RHS As clsLog)
    Set mobjBase.mobjLog = RHS
End Property

Private Property Get clsINF_mobjLog() As clsLog
    Set clsINF_mobjLog = mobjBase.mobjLog
End Property

Private Property Let clsINF_mstrNOs(ByVal RHS As String)
    mobjBase.mstrNOs = RHS
End Property

Private Property Get clsINF_mstrNOs() As String
    clsINF_mstrNOs = mobjBase.mstrNOs
End Property

Private Function clsINF_SendData(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, _
    ByVal strCode As String, ByVal rsData As ADODB.Recordset, strMessage As String, _
    Optional ByRef strWinInfo As String) As Boolean
'���ܡ����������ء�˵������μ�clsINF�����SendData����

    Dim strSQL As String
    Dim colProc As New Collection

    On Error GoTo hErr
    
    If rsData.State <> adStateOpen Then Exit Function
    
    mobjBase.mobjLog.Add "��ʼ��������", 1
    
    Select Case intType
    Case 1 To 3
        rsData.Filter = "ͣ��=0"
    Case 4 To 5
        rsData.Filter = ""
    Case 21
        rsData.Filter = "�Ƿ���=1"
    Case Else
        GoTo makExit
    End Select
    
    If rsData.RecordCount <= 0 Then
makExit:
        mobjBase.mobjLog.Add "�����ݿɴ���", 1
        mobjBase.mobjLog.Save
        clsINF_SendData = True
        Exit Function
    End If
    
    With rsData
        Do While .EOF = False
            '�ô洢���̵����ƹ̶�
            Select Case intType
            Case Val("3-ҩƷĿ¼")
                strSQL = "proHISDrug("
                strSQL = strSQL & "'" & mdlDrugMachine.NVL(!ҩƷid) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!ͨ����) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!���) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!ҩƷ����) & "'"
                strSQL = strSQL & "," & mdlDrugMachine.NVL(!����ϵ��)
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!������λ) & "'"
                strSQL = strSQL & "," & mdlDrugMachine.NVL(!סԺ��װ)
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!סԺ��λ) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!�����̱���) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!������) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!ͨ��������) & "'"
                strSQL = strSQL & ")"
                
            Case Val("21-��ҩ")
                strSQL = "proHISDosage("
                strSQL = strSQL & "'" & mdlDrugMachine.NVL(!����id) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!����) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!�ⷿid) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!סԺ��) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!����id) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!��������) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!����ҽ��) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!����) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!�÷�) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!ҩƷid) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!ͨ����) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!ҩƷ���) & "'"
                strSQL = strSQL & "," & mdlDrugMachine.NVL(!����ϵ��)
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!������λ) & "'"
                strSQL = strSQL & "," & mdlDrugMachine.NVL(!סԺ��λ����)
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!ִ��ʱ��) & "'"
                strSQL = strSQL & ",'" & IIf(Val(mdlDrugMachine.NVL(!���ȼ�)) = 1, "1", "0") & "'"
                strSQL = strSQL & "," & mdlDrugMachine.NVL(!ҽ��id)
                strSQL = strSQL & ")"
                
            Case Else
                strSQL = ""
            End Select
            
            If strSQL <> "" Then
                mdlDrugMachine.AddArray colProc, strSQL
            End If
            
            .MoveNext
        Loop
    End With
    
    '�����������洢����
    On Error GoTo makFail
    Call mdlDrugMachine.ExecuteProcedureBeach(colProc, "", mcnThird, mobjBase.mobjLog)
    On Error GoTo hErr
    
    '������ݴ��ͳɹ�
    Call mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsData, mobjBase, True)
    
    mobjBase.mobjLog.Add "��ɴ�������", 1
    mobjBase.mobjLog.Save
    
    'ִ�л���Ĵ���
    clsINF_SendData = mobjBase.SendData(intAppType, intType, intIO, strCode, rsData, strMessage)
    Exit Function
    
makFail:
    '������ݴ���ʧ��
    Call mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsData, mobjBase, False)
    
hErr:
    mobjBase.mobjLog.Add Err.Number & ":" & Err.Description, 1
    mobjBase.mobjLog.Add "��������ʧ��", 1
    mobjBase.mobjLog.Save
End Function

Private Property Let clsINF_Token(ByVal RHS As String)
    '
End Property

Private Property Get clsINF_Token() As String
    '
End Property

Private Property Let clsINF_URL(ByVal RHS As String)
    mobjBase.URL = RHS
    
    Set mcnThird = New ADODB.Connection
End Property

Private Property Get clsINF_URL() As String
    clsINF_URL = mobjBase.URL
End Property

Private Property Let clsINF_UserName(ByVal RHS As String)
    mobjBase.UserName = RHS
End Property

Private Property Get clsINF_UserName() As String
    clsINF_UserName = mobjBase.UserName
End Property