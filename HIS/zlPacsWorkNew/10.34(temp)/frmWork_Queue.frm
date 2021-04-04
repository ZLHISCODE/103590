VERSION 5.00
Object = "{6856A5DD-B624-47EE-85F4-F9812BFD363A}#1.0#0"; "UcQueueManage.ocx"
Begin VB.Form frmWork_Queue 
   BorderStyle     =   0  'None
   Caption         =   "�ŶӽкŹ���"
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9165
   Icon            =   "frmWork_Queue.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin UcQueueManage.UcQueue ucPacsQueue 
      Height          =   5085
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   8969
      Interval        =   30000
      ValidDays       =   0
   End
End
Attribute VB_Name = "frmWork_Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const M_LNG_PACS_BUSINESS_TYPE As Long = 1                  'pacsҵ�����Ͷ���
Private Const M_STR_NOT_ALLOT_TECHNIC As String = "δ�������"      'δ����������ƶ���


Private mrsPacsQueueGroupConfig As ADODB.Recordset
Private mrsPacsQueueTechnicConfig As ADODB.Recordset


Private mlngCurDeptId As Long                       '��ǰ����ID
Private mstrQueryTechnicQueueNames  As String       'pacs�ŶӽкŲ�ѯ��������

Private mstrQueueCols As String
Private mstrCalledCols As String

Private mstrCurTechnicRoomName As String            '��ǰִ�м�����
Private mstrCurTechnicGroupName As String           '��ǰִ�м���������

Private mlngModule As Long


'����¼�
Public Event OnCompleted(ByVal lngAdviceID As Long)
'�����¼�
Public Event OnDiagnose(ByVal lngAdviceID As Long)
'�����¼�
Public Event OnCalled(ByVal lngAdviceID As Long)

'�Ŷӽкŵ�ѡ��ı��¼�
Public Event OnSelChange(ByVal lngAdviceID As Long)




Property Get Queue() As clsQueueOperation
'���в�������
    Set Queue = ucPacsQueue.QueueOper
End Property






Public Sub zlInitPacsQueueCfg(ByVal lngModule As Long, ByVal lngCurDeptId As Long)
'��ʼ��pacs�ŶӽкŶ�������

    
    mlngModule = lngModule
    mlngCurDeptId = lngCurDeptId

    '��ȡ�ŶӽкŲ�������
    Call ReadQueueParameters(lngCurDeptId)
    
    
    ucPacsQueue.GroupField = "��������"
    
    ucPacsQueue.FindWayEx = "�����,סԺ��,�����,ҽ����"
    ucPacsQueue.DisplayQueueFields = mstrQueueCols
    ucPacsQueue.DisplayCallFields = mstrCalledCols
    
    
    Call ucPacsQueue.InitQueue(gcnOracle, _
                                M_LNG_PACS_BUSINESS_TYPE, _
                                Me, _
                                App.ProductName, _
                                ",���,˳��,ֱ��,�㲥,����,���,����,����,��ͣ,����,�ָ�,���,ˢ��,����,�޸�,����,")
                                                                
    
End Sub


Public Sub zlRefreshQueueData(ByVal strTechnics As String)
'ˢ���Ŷ�����
    
    '������Ҫ��ȡ��ִ�м����ݣ���ָ�����ŶӶ������ݣ�
    mstrQueryTechnicQueueNames = strTechnics & "," & mstrCurTechnicGroupName & "," & M_STR_NOT_ALLOT_TECHNIC
    
    ucPacsQueue.QueryQueueNames = mstrQueryTechnicQueueNames
    
    Call ucPacsQueue.RefreshQueueData
End Sub


Private Sub ReadQueueParameters(ByVal lngCurDeptId As Long)
'��ȡ�ŶӽкŲ���
    '��ȡ��ǰִ�м�����
    mstrCurTechnicRoomName = zlDatabase.GetPara("����ִ�м�����", glngSys, mlngModule, "")
    
    mstrQueueCols = GetDeptPara(lngCurDeptId, "�ŶӶ�����Ϣ����", "")
    mstrCalledCols = GetDeptPara(lngCurDeptId, "���ж�����Ϣ����", "")
    
    mstrCurTechnicGroupName = GetTechnicRoomGrounName(NeedNo(mstrCurTechnicRoomName))   '��ȡ��ǰִ�м����
End Sub


Private Sub ReadQueueRuleConfig()
'��ȡ�Ŷӹ�������
    Dim strSql As String
    
    strSql = "select id,����ID,����,����ǰ׺,��ǰ��� from Ӱ��ִ�з���"
    Set mrsPacsQueueGroupConfig = zlDatabase.OpenSQLRecord(strSql, "��ѯ�Ŷӷ�����Ϣ")
    
    strSql = "select ����ID,ִ�м�,����,��ǰ����,����豸,����ǰ׺,����ID,��ǰ��� from ҽ��ִ�з���"
    Set mrsPacsQueueTechnicConfig = zlDatabase.OpenSQLRecord(strSql, "��ѯִ�м���Ϣ")
End Sub


Public Function zlGetStudyGroupName(ByVal lngAdviceID As Long) As String
'��ȡ�����Ŀ��������
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    zlGetStudyGroupName = ""
    strSql = "select ���� from Ӱ��ִ�з��� " & _
            " where id=(select ����ID " & _
                    " from Ӱ������Ŀ a, ����ҽ����¼ b " & _
                    " where a.������Ŀid = b.������Ŀid and b.id=[1] and b.���ID is null)"
    
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ������", lngAdviceID)
    If rsData.RecordCount <= 0 Then Exit Function
    
    zlGetStudyGroupName = Nvl(rsData!����)
End Function

Public Function zlGetGroupCodeNo(ByVal strGroupName As String) As String
'��ѯ������ŶӺ�����
    mrsPacsQueueGroupConfig.Filter = "����='" & strGroupName & "'"
    
    zlGetGroupCodeNo = ""
    
    If mrsPacsQueueGroupConfig.RecordCount <= 0 Then
        mrsPacsQueueGroupConfig.Filter = ""
        Exit Function
    End If
    
    zlGetGroupCodeNo = Nvl(mrsPacsQueueGroupConfig!����ǰ׺)
    mrsPacsQueueGroupConfig.Filter = ""
End Function

Public Function zlGetTechnicRoomCodeNo(ByVal strTechnicRoom As String, ByVal lngDeptID As Long) As String
'��ѯִ�м���ŶӺ�����
    mrsPacsQueueTechnicConfig.Filter = "ִ�м�='" & strTechnicRoom & "' and ����ID=" & lngDeptID
    
    zlGetTechnicRoomCodeNo = ""
    
    If mrsPacsQueueTechnicConfig.RecordCount <= 0 Then
        mrsPacsQueueGroupConfig.Filter = ""
        Exit Function
    End If
    
    zlGetTechnicRoomCodeNo = Nvl(mrsPacsQueueTechnicConfig!����ǰ׺)
    mrsPacsQueueTechnicConfig.Filter = ""
End Function


Private Function GetTechnicRoomGrounName(ByVal strTechnicRoom As String) As String
'��ȡִ�м������
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetTechnicRoomGrounName = ""
    strSql = "select ���� from Ӱ��ִ�з��� a, ҽ��ִ�з��� b where a.id=b.����ID and b.����Id=[1] and b.ִ�м�=[2]"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯҽ������", mlngCurDeptId, strTechnicRoom)
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetTechnicRoomGrounName = Nvl(rsData!����)
End Function


Public Function zlInPacsQueue(ByVal lngAdviceID As Long, _
                                ByVal strName As String, _
                                ByVal strQueueName As String, _
                                ByVal strTarget As String, _
                                ByVal strNoTag As String) As Boolean
'����pacs�ŶӶ���
On Error GoTo ErrHandle
    Dim lngQueueId As Long
    
    zlInPacsQueue = False
    
    '�����������
    lngQueueId = ucPacsQueue.QueueOper.InsertQueue(strQueueName, , lngAdviceID, strName, strTarget, , "�Ŷӱ��='" & strNoTag & "'")
    If lngQueueId <= 0 Then Exit Function
    
    '��ʼ�Ŷ�
    Call ucPacsQueue.QueueOper.StartQueue(lngQueueId)
    
    Call ucPacsQueue.RefreshQueueData
    
    zlInPacsQueue = True
Exit Function
ErrHandle:
    zlInPacsQueue = False
End Function


Public Sub zlExecuteCommandbar(control As CommandBarControl)
'ִ�в˵��¼�
    Call ucPacsQueue.zlExecuteCommandBars(control)
End Sub


Private Sub Form_Load()
'    'Debug Code...
'        Call InitDebugObject(1290, Me, "zlhis", "HIS")
'        Call InitPacsQueueCfg("���Զ���,050204-��������", "��������", "�ŶӺ���,��������,ҽ������", "�ŶӺ���,��������")
'    'Debug End

    Call ReadQueueRuleConfig
End Sub

Private Sub Form_Resize()
On Error Resume Next
    ucPacsQueue.Left = 0
    ucPacsQueue.Top = 0
    ucPacsQueue.Width = Me.ScaleWidth
    ucPacsQueue.Height = Me.ScaleHeight
err.Clear
End Sub

Private Function GetQueryCol(ByVal strCols As String) As String
'����߱��Ĳ�ѯ���ֶΣ�ID,��������,ҵ��ID,��������,�Ŷ�״̬,�ŶӺ���,�Ŷ�ʱ��,ҵ������,�Ŷ����
    Dim strResult As String
    
    strResult = UCase(strCols)
    
    strResult = Replace(strResult, "ID,", "")
    strResult = Replace(strResult, "��������,", "")
    strResult = Replace(strResult, "ҵ��ID,", "")
    strResult = Replace(strResult, "��������,", "")
    strResult = Replace(strResult, "�Ŷ�״̬,", "")
    strResult = Replace(strResult, "�ŶӺ���,", "")
    strResult = Replace(strResult, "�Ŷ�ʱ��,", "")
    strResult = Replace(strResult, "ҵ������,", "")
    strResult = Replace(strResult, "�Ŷ����,", "")
    
    strResult = "a.ID,a.��������,a.ҵ��ID,a.��������,a.�Ŷ�״̬, a.�Ŷӱ�� || a.�ŶӺ��� as �ŶӺ���,a.�Ŷ�ʱ��,a.ҵ������,a.�Ŷ���� " & IIf(strResult = "", "", "," & strResult)
    
    GetQueryCol = strResult
End Function


Private Function GetAdviceId(ByVal lngQueueId As Long) As Long
'��ȡ�ŶӽкŶ�Ӧ��ҽ��ID
    GetAdviceId = Val(Nvl(ucPacsQueue.QueueOper.GetQueueInf(lngQueueId, "ҵ��ID")!ҵ��ID))
End Function


Private Sub UcPacsQueue_OnCallPreBefore(ByVal lngQueueId As Long, ByVal lngCallWay As UcQueueManage.TCallWay, strCallContext As String, blnCancel As Boolean)
'Pacs�Ŷӽкź����¼�����
    Dim lngAdviceID As Long
    
    
    '����δ����ִ�м�ļ�飬����Ҫ���뵱ǰ���ڵ�ִ�м䵽�Ŷӽкŵ�������
    If lngCallWay = cwOrder Or lngCallWay = cwSpecify Or lngCallWay = cwWaitRoom Then
        lngAdviceID = GetAdviceId(lngQueueId)
        
        Call ucPacsQueue.QueueOper.WriteTarget(lngQueueId, mstrCurTechnicRoomName)
        RaiseEvent OnCalled(lngAdviceID)
    End If
End Sub

Private Sub UcPacsQueue_OnCmdBarUpdate(objComandBarControl As Object)
'���ν��ﰴť
'    If objComandBarControl.ID = TMenuId.mi���� Then
'        objComandBarControl.Visible = False
'    End If
End Sub

Private Sub UcPacsQueue_OnQueryCallData(rsData As ADODB.Recordset, blnUseCustom As Boolean)
'��ѯpacs�Ŷ��Ѻ�������
'�����漰����ѯpacs�����ص�������Ϣ�������Ҫʹ�ø��¼������Զ����ѯ
    Dim strSql As String
    Dim strCurQueryQueueNames As String
    
    blnUseCustom = True
    
    strCurQueryQueueNames = Replace(mstrQueryTechnicQueueNames, ",", "','")
    
    strSql = "select " & GetQueryCol(mstrCalledCols) & " from �ŶӽкŶ��� a, ����ҽ����¼ b where a.ҵ��ID=b.Id and b.���ID is null and a.ҵ������=1 and a.�Ŷ�״̬ in(1,7) " & IIf(strCurQueryQueueNames = "", "", "and �������� in ('" & strCurQueryQueueNames & "') ")
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯPACS�Ѻ��ж���")
End Sub

Private Sub UcPacsQueue_OnQueryQueueData(ByVal lngQueueLoadModle As UcQueueManage.TQueueSelState, rsData As ADODB.Recordset, blnUseCustom As Boolean)
'��ѯpacs�ŶӶ�������
'�����漰����ѯpacs�����ص�������Ϣ�������Ҫʹ�ø��¼������Զ����ѯ
    Dim strSql As String
    Dim strCurQueryQueueNames As String
    
    blnUseCustom = True
    
    strCurQueryQueueNames = Replace(mstrQueryTechnicQueueNames, ",", "','")
    
    strSql = "select " & GetQueryCol(mstrQueueCols) & " from �ŶӽкŶ��� a, ����ҽ����¼ b where a.ҵ��ID=b.Id  and b.���ID is null and a.ҵ������=1 and a.�Ŷ�״̬=[1] " & IIf(strCurQueryQueueNames = "", "", "and �������� in ('" & strCurQueryQueueNames & "') ")
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯPACS�ŶӶ���", lngQueueLoadModle)
End Sub

Private Sub UcPacsQueue_OnSelectionChanged(ByVal lngListType As UcQueueManage.TQueueFromType, ByVal lngQueueId As Long, objQueueList As Object, objReportRow As Object)
'�Ŷӽк�ѡ���иı��¼�
    Dim lngAdviceID As Long
    Dim lngColIndex As Long
    
    If objReportRow Is Nothing Then Exit Sub
    If objReportRow.Record Is Nothing Then Exit Sub
    
    lngColIndex = ucPacsQueue.GetColumnIndex(lngListType, "ҵ��ID")
    
    lngAdviceID = Val(objReportRow.Record(lngColIndex).value)
    
    RaiseEvent OnSelChange(lngAdviceID)
End Sub

Private Sub UcPacsQueue_OnWorkBefore(ByVal lngQueueId As Long, ByVal lngOperationType As UcQueueManage.TOperationType, blnCancel As Boolean)
'������н������������Ҫ���¼��ġ�ִ�м䡱����
    Dim lngAdviceID As Long
    
    '���������Ҫ���������¼�
    If lngOperationType = otComplete Then
        lngAdviceID = GetAdviceId(lngQueueId)
        RaiseEvent OnCompleted(lngAdviceID)
        
    ElseIf lngOperationType = otDiagnose Then
        lngAdviceID = GetAdviceId(lngQueueId)
        RaiseEvent OnDiagnose(lngAdviceID)
        
    End If
End Sub
