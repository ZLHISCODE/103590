VERSION 5.00
Object = "*\A..\queueOper\zlQueueOper.vbp"
Begin VB.Form frmWork_Queue 
   BorderStyle     =   0  'None
   Caption         =   "�ŶӽкŹ���"
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11595
   Icon            =   "frmWork_Queue.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin zlQueueOper.UcQueue ucPacsQueue 
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


Private Const M_LNG_PACS_BUSINESS_IMG_TYPE As Long = 1                  'pacsӰ��ҽ��ҵ�����Ͷ���
Private Const M_LNG_PACS_BUSINESS_CAP_TYPE As Long = 1                  'pacs��Ƶ�ɼ�ҵ�����Ͷ���

Private Const M_STR_NOT_ALLOT_TECHNIC As String = "���Ҷ���"      'δ����������ƶ���
Private Const M_STR_FINDWAY_EX As String = "�����,סԺ��,���￨,ҽ����"


Private mrsPacsQueueGroupConfig As ADODB.Recordset
Private mrsPacsQueueTechnicConfig As ADODB.Recordset


Private mlngCurDeptId As Long                       '��ǰ����ID
Private mstrCurDeptName As String                   '��ǰ��������
Private mstrQueryTechnicQueueNames  As String       'pacs�ŶӽкŲ�ѯ��������
Private mlngQueueNoWay As Long                      '�ŶӺ������ɷ�ʽ
Private mlngValidDays As Long
Private mstrReportNum As String
Private mlngPrintWay As Long
Private mblnUseQueueMsg As Boolean

Private mstrQueueCols As String
Private mstrCalledCols As String


Private mstrCurTechnicRoomName As String            '��ǰִ�м�����
Private mstrCurTechnicDevice As String              '��Ӧ��ǰִ�м��豸
Private mstrCurTechnicGroupName As String           '��ǰִ�м���������
Private mstrTurnPage As String                      '�������תҳ��

Private mlngModule As Long
Private mstrPrivs As String

Public Event OnQueueQuick(blnOpenQuick As Boolean)
Public Event OnCallAboutLock(ByVal lngType As Long, ByRef strLockedName As String, ByVal blnLockPara As Boolean)
'104686��أ����к�������飬
'lngType����  1:�ж��Ƿ������˲��������Ƿ��Ѿ��б������ļ��        2:���²���
'strLockedName   ��="" ������û��Ӱ�죬����˵���Ѿ����ò������ҷ���֮ǰ�����ļ�黼������
'blnLockPara   ���ڸ���PacsMain�еĲ���

'�����¼�
Public Event OnResotre(ByVal lngAdviceID As Long, ByVal strExeRoom As String)
'����¼�
Public Event OnCompleted(ByVal lngAdviceID As Long, ByVal strExeRoom As String)
'�����¼�
Public Event OnDiagnose(ByVal lngAdviceID As Long, ByVal strExeRoom As String, ByVal strTurnPage As String)
'�����¼��� ���к���Ҫ���ļ�����ң�ֻ���ڽ�������ʱ�Ž��и���
Public Event OnCalled(ByVal lngAdviceID As Long, ByVal strRoom As String, ByVal TCallWay As zlQueueOper.TCallWay)

'�Ŷӽкŵ�ѡ��ı��¼�
Public Event OnSelChange(ByVal lngAdviceID As Long)
'����������ʾ�ı��¼�
Public Event OnGroupHint(ByVal strHint As String)



Public Sub zlInitPacsQueueCfg(ByVal lngModule As Long, _
                            ByVal lngCurDeptId As Long, _
                            ByVal strCurDeptName As String, _
                            ByVal strPrivs As String)
'��ʼ��pacs�ŶӽкŶ�������
    Dim lngCurWorkType As Long
    Dim strQueuePrivs As String
    
    mlngModule = lngModule
    mlngCurDeptId = lngCurDeptId
    mstrCurDeptName = strCurDeptName
    mstrPrivs = strPrivs
    
    strQueuePrivs = ";" & GetPrivFunc(glngSys, 1160) & ";"
    
    lngCurWorkType = IIf(mlngModule = 1290, M_LNG_PACS_BUSINESS_IMG_TYPE, M_LNG_PACS_BUSINESS_CAP_TYPE)
    
    '��ȡ�ŶӽкŲ�������
    Call ReadQueueParameters(lngCurDeptId)
    
    
    ucPacsQueue.ValidDays = mlngValidDays
    ucPacsQueue.ReportNum = mstrReportNum
    ucPacsQueue.GroupField = "��������"
    ucPacsQueue.IsReleationQueueTag = True
    
    ucPacsQueue.FindWayEx = M_STR_FINDWAY_EX
    
    '��Ҫʹ����ҵ���йصĲ�ѯʱ����Ҫ��DataFields���Խ�������
    ucPacsQueue.DataFields = "ID,ҵ������,��������,����ID,����ID,ҵ��ID,�Ŷ����,�ŶӺ���,����,��������,�Ա�,����,�����Ŀ,ҽ������,�Ŷ�״̬,�Ŷ�ʱ��,����ҽ��,����ʱ��,��ע"
    ucPacsQueue.DisplayQueueFields = mstrQueueCols '& ",�Ŷ����"
    ucPacsQueue.DisplayCallFields = mstrCalledCols '& ",�Ŷ����"
    
    ucPacsQueue.CalledTarget = mstrCurTechnicRoomName       '���ú�������Ŀ�ĵ�
    
    
    If mblnUseQueueMsg = True Then
        '�����Ŷ���Ϣ����
        Call ucPacsQueue.UseMsgCenter(glngSys, lngModule)
    End If
    
    Call ucPacsQueue.InitQueue(gcnOracle, _
                                lngCurWorkType, _
                                Me, _
                                App.ProductName, _
                                UserInfo.����, _
                                strQueuePrivs)
                                                                
    '����Ѿ����ڵ��Ŷӽк�ҵ��
    Call ucPacsQueue.QueueOper.CustomClearData("����ID=" & lngCurDeptId)
    
    'Ӧ�ú������ã���������������
    Call ucPacsQueue.ApplyVoiceConfig
End Sub


Public Sub zlRefreshQueueData(ByVal strTechnics As String)
'ˢ���Ŷ�����
    Dim i As Integer
    Dim strTmp As String
    Dim strTechnicGroupNames As String
    
    '������Ҫ��ȡ��ִ�м����ݣ���ָ�����ŶӶ������ݣ�
    mstrQueryTechnicQueueNames = ""
    
    If strTechnics <> "" Then
        '0-��Ĭ�Ϲ�����飬1-���������÷���
        If mlngQueueNoWay = 1 Then
            '��ȡ����ѡ���ִ�м��Ӧ�ķ���,�����:80403
            If UBound(Split(strTechnics, ",")) > 0 Then
                For i = 0 To UBound(Split(strTechnics, ","))
                    strTmp = GetTechnicRoomGrounName(mlngCurDeptId, Split(Split(strTechnics, ",")(i), "-")(1))
                    If strTmp <> "" Then strTechnicGroupNames = strTechnicGroupNames & "," & strTmp
                Next
                
                strTechnicGroupNames = Mid(strTechnicGroupNames, 2)
            Else
                strTmp = GetTechnicRoomGrounName(mlngCurDeptId, Split(strTechnics, "-")(1))
                If strTmp <> "" Then strTechnicGroupNames = strTmp
            End If
            
            mstrQueryTechnicQueueNames = strTechnics & "," & strTechnicGroupNames
        Else
            mstrQueryTechnicQueueNames = strTechnics & "," & mstrCurDeptName & "-" & M_STR_NOT_ALLOT_TECHNIC
        End If
    End If
    
    ucPacsQueue.QueryQueueNames = mstrQueryTechnicQueueNames
    
    If mlngQueueNoWay = 0 Then
        ucPacsQueue.LastFixedQueue = M_STR_NOT_ALLOT_TECHNIC
    Else
        ucPacsQueue.LastFixedQueue = mstrCurTechnicGroupName
    End If
    
    Call ucPacsQueue.RefreshQueueData
End Sub


Private Sub ReadQueueParameters(ByVal lngCurDeptId As Long)
'��ȡ�ŶӽкŲ���
    Dim strDeptId As String
    Dim strRoomName As String
    
    '��ȡ��ǰִ�м�����
    strDeptId = Val(zlDatabase.GetPara("����ִ�м����", glngSys, mlngModule, ""))
    strRoomName = zlDatabase.GetPara("����ִ�м�����", glngSys, mlngModule, "")
    mstrTurnPage = zlDatabase.GetPara("�������תҳ��", glngSys, mlngModule, "")
    mlngValidDays = Val(GetDeptPara(lngCurDeptId, "�Ŷ����ݱ�������", 1))
    mstrReportNum = GetDeptPara(lngCurDeptId, "�Ŷӵ�������", "")
    mlngPrintWay = Val(GetDeptPara(lngCurDeptId, "�Ŷӵ���ӡ��ʽ", 0))
    mblnUseQueueMsg = Val(GetDeptPara(lngCurDeptId, "�����Ŷ���Ϣ����", 1))
    
    mstrCurTechnicRoomName = Trim(zlStr.NeedCode(strRoomName))
    mstrCurTechnicDevice = Trim(zlStr.NeedName(strRoomName))
    
    mstrQueueCols = zlDatabase.GetPara("�ŶӶ�����Ϣ����", glngSys, mlngModule, "�ŶӺ���,��������") 'GetDeptPara(lngCurDeptId, "�ŶӶ�����Ϣ����", "")
    mstrCalledCols = zlDatabase.GetPara("���ж�����Ϣ����", glngSys, mlngModule, "�ŶӺ���,��������") 'GetDeptPara(lngCurDeptId, "���ж�����Ϣ����", "")
    
    mlngQueueNoWay = Val(GetDeptPara(lngCurDeptId, "�Ŷӽкű������", 0))
    
    mstrCurTechnicGroupName = GetTechnicRoomGrounName(strDeptId, mstrCurTechnicRoomName)   '��ȡ��ǰִ�м����
End Sub

Private Sub ReadQueueRuleConfig()
'��ȡ�Ŷӹ�������
    Dim strSql As String
    
    strSql = "select id,����ID,����,����ǰ׺ from Ӱ��ִ�з���"
    Set mrsPacsQueueGroupConfig = zlDatabase.OpenSQLRecord(strSql, "��ѯ�Ŷӷ�����Ϣ")
    
    strSql = "select ����ID,ִ�м�,����,��ǰ����,����豸,����ǰ׺,����ID from ҽ��ִ�з���"
    Set mrsPacsQueueTechnicConfig = zlDatabase.OpenSQLRecord(strSql, "��ѯִ�м���Ϣ")
End Sub


Public Sub zlGetInQueueInf(ByVal lngAdviceID As Long, ByVal lngExecuteDeptId As Long, _
    ByRef strQueueName As String, ByRef strCodeNo As String)
'��ȡ��������Ϣ
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strQueueName = ""
    strCodeNo = ""
    
    If mlngQueueNoWay = 0 Then
        '�������Ŷ�
        strSql = "select ���� from ���ű� where id=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ��������", lngExecuteDeptId)
        
        If rsData.RecordCount <= 0 Then Exit Sub
                
        strQueueName = Nvl(rsData!����) & "-" & M_STR_NOT_ALLOT_TECHNIC
        strCodeNo = ""
    Else
        '�������Ŷ�
        strSql = "select a.����,a.����ǰ׺,b.���� from Ӱ��ִ�з��� a, ���ű� b " & _
                " where a.����Id=b.Id and a.id=(select a.����ID " & _
                        " from Ӱ�������� a, ����ҽ����¼ b " & _
                        " where a.������Ŀid = b.������Ŀid and a.����ID=[1] and b.id=[2] and b.���ID is null)"
                        
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ������", lngExecuteDeptId, lngAdviceID)
        
        If rsData.RecordCount <= 0 Then Exit Sub
        
        strQueueName = Nvl(rsData!����) & "-" & Nvl(rsData!����)
        strCodeNo = Nvl(rsData!����ǰ׺)
    End If
End Sub

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


Private Function GetTechnicRoomGrounName(ByVal lngDeptID As Long, ByVal strTechnicRoom As String) As String
'��ȡִ�м������
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetTechnicRoomGrounName = ""
    strSql = "select c.����,���� from Ӱ��ִ�з��� a, ҽ��ִ�з��� b, ���ű� c where a.id=b.����ID  and b.����ID=c.Id and b.����Id=[1] and b.ִ�м�=[2]"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯҽ������", lngDeptID, strTechnicRoom)
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetTechnicRoomGrounName = Nvl(rsData!����) & "-" & Nvl(rsData!����)
End Function

Public Function zlInQueue(ByVal lngAdviceID As Long, _
                                ByVal strName As String, _
                                ByVal lngDeptID As Long, _
                                ByVal strQueueName As String, _
                                ByVal strTarget As String, _
                                ByVal strNoTag As String) As Boolean
        
    Dim rsData As ADODB.Recordset
    Dim lngTimePoint As Long
    Dim lngTimeInterval As Long
On Error GoTo errHandle
    
    zlInQueue = False

    Set rsData = ucPacsQueue.QueueOper.FindQueueInf(lngAdviceID)
    
    If rsData.RecordCount > 0 Then  '�����Ŷ�����
        lngTimePoint = Val(Format(time, "h"))
        If lngTimePoint <= 4 Then
            lngTimeInterval = DateDiff("s", Nvl(rsData!�Ŷ�ʱ��), Format(zlDatabase.Currentdate - 1, "YYYY-MM-DD 20:00:00"))
        Else
            lngTimeInterval = DateDiff("s", Nvl(rsData!�Ŷ�ʱ��), Format(zlDatabase.Currentdate, "YYYY-MM-DD 00:00:00"))
        End If
        
        If lngTimeInterval > 0 Then
            '���ǽ�����ǰ�����ݣ���ֱ�Ӹ����Ŷ�
            Call zlUpdatePacsQueue(lngAdviceID, strName, lngDeptID, strQueueName, strTarget, strNoTag)
        Else
            If MsgBoxD(Me, "�˲��������ŶӽкŶ����У��Ƿ������Ŷӣ�", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                Call zlUpdatePacsQueue(lngAdviceID, strName, lngDeptID, strQueueName, strTarget, strNoTag)
            End If
        End If
    Else
        Call zlInPacsQueue(lngAdviceID, strName, lngDeptID, strQueueName, strTarget, strNoTag)
    End If
    
    zlInQueue = True
Exit Function
errHandle:
    zlInQueue = False
End Function

Public Function zlInPacsQueue(ByVal lngAdviceID As Long, _
                                ByVal strName As String, _
                                ByVal lngDeptID As Long, _
                                ByVal strQueueName As String, _
                                ByVal strTarget As String, _
                                ByVal strNoTag As String) As Boolean
'����pacs�ŶӶ���

On Error GoTo errHandle
    Dim lngQueueId As Long
    Dim strExpandData As String
    Dim strNewQueueNo As String
    
    zlInPacsQueue = False
    
    strExpandData = "����Id=" & lngDeptID & ",�Ŷӱ��='" & strNoTag & "'"
    '�����������
    lngQueueId = ucPacsQueue.QueueOper.InsertQueue(strQueueName, , lngAdviceID, strName, strTarget, , strExpandData)
    If lngQueueId <= 0 Then Exit Function
    
    '��ʼ�Ŷ�
    Call ucPacsQueue.QueueOper.LineQueue(lngQueueId, strNewQueueNo)
    
    'ˢ���б�������ʾ
    Call ucPacsQueue.RefreshQueueRowState(lngQueueId, TQueueState.qsQueueing)
    
    Call AutoPrintQueueInf(lngQueueId)
    
    zlInPacsQueue = True
Exit Function
errHandle:
    zlInPacsQueue = False
End Function

Private Sub AutoPrintQueueInf(ByVal lngQueueId As Long)
'�Զ���ӡ������Ϣ
On Error GoTo errHandle
    If mlngPrintWay = 1 Then
        '�Զ���ӡ
        Call ucPacsQueue.QueueOper.PrintQueueNo(lngQueueId, True, Me)
    ElseIf mlngPrintWay = 2 Then
        '��ʾ��ӡ
        If MsgBoxD(Me, "�Ƿ��ӡ��ǰ�ź���Ϣ��", vbYesNo, gstrSysName) = vbYes Then
            Call ucPacsQueue.QueueOper.PrintQueueNo(lngQueueId, True, Me)
        End If
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function zlUpdatePacsQueue(ByVal lngAdviceID As Long, _
                                ByVal strPatientName As String, _
                                ByVal lngDeptID As Long, _
                                Optional ByVal strQueueName As String = "", _
                                Optional ByVal strTarget As String = " ", _
                                Optional ByVal strNoTag As String = " ") As Boolean
'�����ŶӶ�����Ϣ
    Dim lngQueueId As Long
    Dim strExpandData As String
    
    zlUpdatePacsQueue = False
    
    If strQueueName = "" Then Exit Function
    
    lngQueueId = ucPacsQueue.QueueOper.FindQueueId(lngAdviceID)

    
    If strPatientName <> "" Then
        Call ucPacsQueue.QueueOper.DeleteQueue(lngQueueId)
        zlUpdatePacsQueue = zlInPacsQueue(lngAdviceID, strPatientName, lngDeptID, strQueueName, strTarget, strNoTag)
    Else
    
        strExpandData = ""
        If strPatientName <> "" Then
            strExpandData = strExpandData & "��������=''" & strPatientName & "''"
        End If
    
        If strTarget <> " " Then
            If strExpandData <> "" Then strExpandData = strExpandData & ","
            strExpandData = strExpandData & "����=''" & strTarget & "''"
        End If
        
        If strNoTag <> " " Then
            If strExpandData <> "" Then strExpandData = strExpandData & ","
            strExpandData = strExpandData & "�Ŷӱ��=''" & strNoTag & "''"
        End If
    
        Call ucPacsQueue.QueueOper.UpdateQueue(lngQueueId, strExpandData)
        Call ucPacsQueue.RefreshQueueData
    
        zlUpdatePacsQueue = True
    End If
End Function


Public Function zlCancelPacsQueue(ByVal lngAdviceID As Long) As Boolean
'����pacs�Ŷ�
    Dim lngQueueId As Long
    
    zlCancelPacsQueue = False
    lngQueueId = ucPacsQueue.QueueOper.FindQueueId(lngAdviceID)
    
    'ִ������ɾ������
    Call ucPacsQueue.QueueOper.DeleteQueue(lngQueueId)
    
    zlCancelPacsQueue = True
    
    'ˢ���б�������ʾ
    Call ucPacsQueue.RefreshQueueRowState(lngQueueId, TQueueState.qsAbstain)
End Function


Public Function zlCompletePacsQueue(ByVal lngAdviceID As Long) As Boolean
'���pacs�Ŷ�
    Dim lngQueueId As Long
    
    lngQueueId = ucPacsQueue.QueueOper.FindQueueId(lngAdviceID)
    
    'ִ������ŶӲ���
    zlCompletePacsQueue = ucPacsQueue.QueueOper.CompleteQueue(lngQueueId)
    
    'ˢ���б�������ʾ
    Call ucPacsQueue.RefreshQueueRowState(lngQueueId, TQueueState.qsComplete)
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
    
    If Me.ScaleWidth < 12900 Then
        ucPacsQueue.IsIconLarge = False
        ucPacsQueue.IsShowToolText = IIf(Me.ScaleWidth < 8000, False, True)
    Else
        ucPacsQueue.IsIconLarge = True
        ucPacsQueue.IsShowToolText = True
    End If
err.Clear
End Sub

Private Function GetRoom(ByVal lngQueueId As Long) As String
'��ȡ�ŶӽкŶ�Ӧ�����ң�ִ�м䣩
    GetRoom = Nvl(ucPacsQueue.QueueOper.GetQueueInf(lngQueueId, "����")!����)
End Function


Private Function GetAdviceId(ByVal lngQueueId As Long) As Long
'��ȡ�ŶӽкŶ�Ӧ��ҽ��ID
    GetAdviceId = Val(Nvl(ucPacsQueue.QueueOper.GetQueueInf(lngQueueId, "ҵ��ID")!ҵ��ID))
End Function


Private Sub ucPacsQueue_OnCallPreAfter(ByVal lngQueueId As Long, ByVal lngCallWay As zlQueueOper.TCallWay)
'���к󴥷��¼�
    Dim lngAdviceID As Long
    Dim strRoom As String
    
    lngAdviceID = GetAdviceId(lngQueueId)
    strRoom = GetRoom(lngQueueId)
    
    RaiseEvent OnCalled(lngAdviceID, strRoom, lngCallWay)

End Sub

Private Sub UcPacsQueue_OnCallPreBefore(ByVal lngQueueId As Long, ByVal lngCallWay As zlQueueOper.TCallWay, strCallContext As String, blnCancel As Boolean)
'Pacs�Ŷӽкź����¼�����
    Dim strOldTechnicRoomName As String
    Dim lngResult As Long
    Dim lngRowIndex As Long
    Dim strSql As String
    Dim lngAdviceID As Long
    Dim strName As String
    Dim blnTmp As Boolean
        
    '����δ����ִ�м�ļ�飬����Ҫ���뵱ǰ���ڵ�ִ�м䵽�Ŷӽкŵ�������
    If lngCallWay = cwOrder Or lngCallWay = cwSpecify Or lngCallWay = cwWaitRoom Then
        '�Ѿ��������ļ�飬����Ȼ�����������ڵ�
        RaiseEvent OnCallAboutLock(1, strName, blnTmp)
                
        '�жϵ�ǰ�����Ƿ��Ѿ���ִ�м䣬����Ѿ����䵫�뵱ǰִ�м䲻ͬ������Ҫ��������
        strOldTechnicRoomName = Trim(Nvl(ucPacsQueue.QueueOper.GetQueueInf(lngQueueId, "����")!����))
        
        If strOldTechnicRoomName <> "" And strOldTechnicRoomName <> mstrCurTechnicRoomName Then
            lngResult = MsgBoxD(Me, "��ǰ����ѱ����䵽 ��" & strOldTechnicRoomName & "�� ִ�м䣬�Ƿ���Ҫ���ĵ���ִ�м�ִ�У�" & vbCrLf & _
                                    "ѡ���ǡ���ʾ���ĵ���ִ�м����У�" & vbCrLf & _
                                    "ѡ�񡰷񡱱�ʾ������ִ�м�ֱ�Ӻ��У�" & vbCrLf & _
                                    "ѡ��ȡ������ʾ�����к��У�", vbYesNoCancel, "��ʾ")
            
            If lngResult = vbCancel Then
                blnCancel = True
                Exit Sub
            End If
        End If
          
        '��������Ŀ�ĵ�
        If lngResult = vbYes Or strOldTechnicRoomName = "" Then
            Call ucPacsQueue.QueueOper.WriteTarget(lngQueueId)
            '��Ҫͬ������ҽ�����͵�ִ�м�
            
            lngAdviceID = GetAdviceId(lngQueueId)
            
            strSql = "zl_Ӱ����_����ִ�м�(" & lngAdviceID & ",'" & mstrCurTechnicRoomName & "','" & mstrCurTechnicDevice & "')"
            
            Call zlDatabase.ExecuteProcedure(strSql, "���¼��ִ�м�")
        
            '�����Ŷ��б��ϵ�������ʾ
            lngRowIndex = ucPacsQueue.GetRowIndex(qftWaitQueue, "ID", lngQueueId)
            If lngRowIndex >= 0 Then
                Call ucPacsQueue.SetListValue(qftWaitQueue, lngRowIndex, "����", mstrCurTechnicRoomName)
                Call ucPacsQueue.Populate(qftWaitQueue)
            End If
        End If
    End If
End Sub

Private Sub UcPacsQueue_OnCmdBarUpdate(objComandBarControl As Object)
'���ν��ﰴť
'    If objComandBarControl.ID = TMenuId.mi���� Then
'        objComandBarControl.Visible = False
'    End If
End Sub


Private Sub ucPacsQueue_OnConfigEvent(blnUseCustom As Boolean)
'Pacs�Ŷӽк������¼�
On Error GoTo errHandle
    Dim objCfgWindow As frmWork_QueueCfg
    Dim blnLock As Boolean
    Dim blnQuick As Boolean
    Dim strTmp As String
    
    blnUseCustom = True
    
    Set objCfgWindow = New frmWork_QueueCfg
    
    If objCfgWindow.ShowQueueConfig(ucPacsQueue, mlngModule, mstrPrivs, Me, blnLock, blnQuick) Then
        '���¶�ȡ��Ӧ������
        RaiseEvent OnCallAboutLock(2, strTmp, blnLock)
        RaiseEvent OnQueueQuick(blnQuick)
        Call zlInitPacsQueueCfg(mlngModule, mlngCurDeptId, mstrCurDeptName, mstrPrivs)
        Call ucPacsQueue.RefreshQueueData
    End If
    
Exit Sub
errHandle:
    Unload objCfgWindow
    Set objCfgWindow = Nothing
    
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucPacsQueue_OnFindData(ByVal strFindWay As String, ByVal strFindValue As String, txtFind As Object, rsData As ADODB.Recordset, blnUseCustom As Boolean)
'�Զ������
    Dim strSql As String
    Dim strCurQueryQueueNames As String
    Dim strQueryCols As String
    Dim blnQueryProject As Boolean
    Dim strTempCols As String '��Щ�е���������Ϊnumber���ͣ��������val����
    
    blnUseCustom = True
    strCurQueryQueueNames = Replace(mstrQueryTechnicQueueNames, ",", "','")
    
    If strFindWay = "�ŶӺ�" Or strFindWay = "�ŶӺ���" Then strFindWay = "�Ŷӱ��||a.�ŶӺ���"
    If strFindWay = "����" Then strFindWay = "��������"
    strTempCols = "�����,סԺ��,ID,ҵ������,����ID,����ID,ҵ��ID,�Ŷ�״̬"
        
    '"ID,ҵ������,��������,����ID,����ID,ҵ��ID,�Ŷ����,�ŶӺ���,����,��������,�Ա�,����,�����Ŀ,ҽ������,�Ŷ�״̬,�Ŷ�ʱ��,����ҽ��,����ʱ��,��ע"
    
    '��ȡ��Ҫ�����ݿ��в�ѯ���ֶ�
    strQueryCols = ucPacsQueue.GetValidCols("a.ID,a.ҵ������,a.��������,a.����ID,a.����ID,a.ҵ��ID,a.�ŶӺ���,a.�Ŷӱ��,a.�Ŷ����,a.����," & _
                                            "a.��������,b.�Ա�,b.����,c.���� as �����Ŀ,b.ҽ������,a.�Ŷ�״̬," & _
                                            "a.�Ŷ�ʱ��,a.����ҽ��,a.����ʱ��,a.��ע", "a")
    
    blnQueryProject = IIf(InStr(strQueryCols, "�����Ŀ") > 0, True, False)
    
    strSql = "select " & strQueryCols & _
            " from �ŶӽкŶ��� a, ����ҽ����¼ b,������Ϣ d " & IIf(blnQueryProject, ", ������ĿĿ¼ c ", "") & _
            " where a.ҵ��ID=b.Id and b.����ID=d.����ID " & _
                    IIf(blnQueryProject, " and b.������ĿID=c.ID and c.���='D'", "") & _
            "       and b.���ID is null and a.ҵ������=1 " & _
            "       and a.����ID=[1] " & IIf(strCurQueryQueueNames = "", "", "and �������� in ('" & strCurQueryQueueNames & "') ") & _
            IIf(InStr(M_STR_FINDWAY_EX, strFindWay) > 0, " and upper(d.", " and upper(a.") & IIf(strFindWay = "���￨", "���￨��", strFindWay) & ")=upper([2]) " & _
            IIf(ucPacsQueue.QueueOper.CustomOrder = "", "", " order by " & ucPacsQueue.QueueOper.CustomOrder)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯPACS�ŶӶ���", mlngCurDeptId, IIf(InStr(strTempCols, strFindWay) > 0, Val(strFindValue), Trim(strFindValue)))
End Sub

Private Sub ucPacsQueue_OnLocateData(ByVal strLocateWay As String, ByVal strLocateValue As String, txtFind As Object, lngQueueId As Long, blnUseCustom As Boolean)
'�Ŷ����ݶ�λ�¼�
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strTempCols As String '��Щ�е���������Ϊnumber���ͣ��������val����
    
    blnUseCustom = True
    
    If strLocateWay = "�ŶӺ�" Or strLocateWay = "�ŶӺ���" Then strLocateWay = "�Ŷӱ��||a.�ŶӺ���"
    If strLocateWay = "����" Then strLocateWay = "��������"
    strTempCols = "�����,סԺ��,ID,ҵ������,����ID,����ID,ҵ��ID,�Ŷ�״̬"
    
    strSql = "select a.ID from �ŶӽкŶ��� a, ����ҽ����¼ b, ������Ϣ d" & _
            " where a.ҵ��ID=b.ID and b.����ID=d.����ID and b.���ID is null and upper(" & _
            IIf(InStr(M_STR_FINDWAY_EX, strLocateWay) > 0, " d.", " a.") & IIf(strLocateWay = "���￨", "���￨��", strLocateWay) & ")=upper([1])"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��λ�Ŷ�����", IIf(InStr(strTempCols, strLocateWay) > 0, Val(strLocateValue), Trim(strLocateValue)))
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    lngQueueId = Val(Nvl(rsData!ID))
End Sub

Private Sub ucPacsQueue_OnGroupHint(ByVal strHintContext As String)
On Error Resume Next
    RaiseEvent OnGroupHint(strHintContext)
err.Clear
End Sub

Private Sub ucPacsQueue_OnModifyBefore(ByVal lngListType As zlQueueOper.TQueueFromType, ByVal lngQueueId As Long, objInputCfg As Dictionary, blnCancel As Boolean, blnUseCustom As Boolean)
    '��ѯ��ǰ���ҵ�������Ϣ
    Dim strRooms As String
    
    If mrsPacsQueueTechnicConfig Is Nothing Then Exit Sub
    
    mrsPacsQueueTechnicConfig.Filter = "����ID=" & mlngCurDeptId
    If mrsPacsQueueTechnicConfig.RecordCount <= 0 Then Exit Sub
    
    While Not mrsPacsQueueTechnicConfig.EOF
        If strRooms <> "" Then strRooms = strRooms & ","
        strRooms = strRooms & Nvl(mrsPacsQueueTechnicConfig!ִ�м�)
        
        Call mrsPacsQueueTechnicConfig.MoveNext
    Wend
    
    mrsPacsQueueTechnicConfig.Filter = ""
    
    objInputCfg.Item("����") = objInputCfg.Item("����") & ":" & strRooms
err.Clear
End Sub

Private Sub ucPacsQueue_OnModifyAfter(ByVal lngQueueId As Long, objUpdateValue As Dictionary)
    objUpdateValue.Item("�ŶӺ���") = objUpdateValue.Item("�Ŷӱ��") & objUpdateValue.Item("�ŶӺ���")
End Sub

Private Sub UcPacsQueue_OnQueryQueueData(rsData As ADODB.Recordset, blnUseCustom As Boolean)
'��ѯpacs�ŶӶ�������
'�����漰����ѯpacs�����ص�������Ϣ�������Ҫʹ�ø��¼������Զ����ѯ
'��ѯ������Ŷ����
    Dim strSql As String
    Dim strCurQueryQueueNames As String
    Dim lngTimePoint As Long
    Dim strStartTime As String
    Dim strEndTime As String
    Dim strQueryCols As String
    Dim blnQueryProject As Boolean
    Dim dtNow As Date
    blnUseCustom = True
    
    strCurQueryQueueNames = Replace(mstrQueryTechnicQueueNames, ",", "','")
    dtNow = zlDatabase.Currentdate
    
    lngTimePoint = Val(Format(time, "h"))
    If lngTimePoint <= 4 Then
        strStartTime = zlStr.To_Date(Format(dtNow - 1, "yy-mm-dd 20:00:00"))
        strEndTime = zlStr.To_Date(Format(dtNow, "yy-mm-dd 08:00:00"))
    Else
        strStartTime = zlStr.To_Date(Format(dtNow, "yy-mm-dd 00:00:00"))
        strEndTime = zlStr.To_Date(Format(dtNow, "yy-mm-dd 23:59:59"))
    End If
    
    '"ID,ҵ������,��������,����ID,����ID,ҵ��ID,�Ŷ����,�ŶӺ���,����,��������,�Ա�,����,�����Ŀ,ҽ������,�Ŷ�״̬,�Ŷ�ʱ��,����ҽ��,����ʱ��,��ע"
    
    '��ȡ��Ҫ�����ݿ��в�ѯ���ֶ�
    strQueryCols = ucPacsQueue.GetValidCols("a.ID,a.ҵ������,a.��������,a.����ID,a.����ID,a.ҵ��ID,a.�Ŷӱ��,a.�ŶӺ���,a.�Ŷ����,a.����," & _
                                            "a.��������,b.�Ա�,b.����,c.���� as �����Ŀ,b.ҽ������,a.�Ŷ�״̬," & _
                                            "a.�Ŷ�ʱ��,a.����ҽ��,a.����ʱ��,a.��ע", "a")
    
    'strQueryCols = Replace(strQueryCols, "A.�ŶӺ���", "A.�Ŷӱ�� || A.�ŶӺ��� as �ŶӺ���")
    
    blnQueryProject = IIf(InStr(strQueryCols, "�����Ŀ") > 0, True, False)
    
    strSql = "select " & strQueryCols & _
            " from �ŶӽкŶ��� a, ����ҽ����¼ b" & IIf(blnQueryProject, ", ������ĿĿ¼ c ", "") & _
            " where a.ҵ��ID=b.Id " & _
                    IIf(blnQueryProject, " and b.������ĿID=c.ID and c.���='D'", "") & " and b.���ID is null and a.ҵ������=1 and a.�Ŷ�ʱ�� between " & strStartTime & " and " & strEndTime & _
            "       and a.����ID=[1] " & IIf(strCurQueryQueueNames = "", "", "and �������� in ('" & strCurQueryQueueNames & "') ") & IIf(ucPacsQueue.QueueOper.CustomOrder = "", "", " order by " & ucPacsQueue.QueueOper.CustomOrder)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯPACS�ŶӶ���", mlngCurDeptId)
End Sub

Private Sub UcPacsQueue_OnSelectionChanged(ByVal lngListType As zlQueueOper.TQueueFromType, ByVal lngQueueId As Long, objQueueList As Object, objReportRow As Object)
'�Ŷӽк�ѡ���иı��¼�
    Dim lngAdviceID As Long
    Dim lngColIndex As Long
    
    If objReportRow Is Nothing Then Exit Sub
    If objReportRow.Record Is Nothing Then Exit Sub
    
    lngColIndex = ucPacsQueue.GetColumnIndex(lngListType, "ҵ��ID")
    
    lngAdviceID = Val(objReportRow.Record(lngColIndex).value)
    
    RaiseEvent OnSelChange(lngAdviceID)
End Sub

Private Sub ucPacsQueue_OnWorkAfter(ByVal lngQueueId As Long, ByVal strCurQueueName As String, ByVal lngOperationType As zlQueueOper.TOperationType)
'������н������������Ҫ���¼��ġ�ִ�м䡱����
    Dim lngAdviceID As Long
    Dim strSql As String
    Dim strRoom As String
    Dim lngRowIndex As Long
    Dim strCodeTag As String
    Dim rsData As ADODB.Recordset
    
    If lngOperationType = otComplete Then
        lngAdviceID = GetAdviceId(lngQueueId)
        
        '���ʱ����Ҫ�������յ�ִ�м�
        strSql = "zl_Ӱ����_����ִ�м�(" & lngAdviceID & ",'" & mstrCurTechnicRoomName & "','" & mstrCurTechnicDevice & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        RaiseEvent OnCompleted(lngAdviceID, mstrCurTechnicRoomName)
        
    ElseIf lngOperationType = otDiagnose Then
        lngAdviceID = GetAdviceId(lngQueueId)
        
        '����ʱ�����µ�ǰ����ִ�м�
        strSql = "zl_Ӱ����_����ִ�м�(" & lngAdviceID & ",'" & mstrCurTechnicRoomName & "','" & mstrCurTechnicDevice & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                        
        RaiseEvent OnDiagnose(lngAdviceID, mstrCurTechnicRoomName, mstrTurnPage)
    ElseIf lngOperationType = otRestore Then
        strRoom = ""
        
        lngRowIndex = ucPacsQueue.GetRowIndex(qftWaitQueue, "ID", lngQueueId)
            
        '��������ţ���Ҫ�ж��Ƿ�������ŶӶ��У�������н����˵���������Ҫ�����һ�ִ�м���ж�Ӧ�ĸ���
        If strCurQueueName <> mstrCurTechnicGroupName And strCurQueueName <> mstrCurDeptName & "-" & M_STR_NOT_ALLOT_TECHNIC Then
            '��ȡ��ǰ���ж�Ӧ����������
            strRoom = Replace(strCurQueueName, mstrCurDeptName & "-", "")
            
            '�����Ŷ�����
            Call ucPacsQueue.QueueOper.WriteTarget(lngQueueId, strRoom)
        End If
        
        '����ҽ��ִ�м�
        lngAdviceID = GetAdviceId(lngQueueId)
        
        strSql = "zl_Ӱ����_����ִ�м�(" & lngAdviceID & ",'" & strRoom & "','" & mstrCurTechnicDevice & "')"
        Call zlDatabase.ExecuteProcedure(strSql, "���¼��ִ�м�")

        'ˢ�½����ŶӺ��뼰�Ŷ����ҵ���ʾ
        If lngRowIndex >= 0 Then
            Call ucPacsQueue.SetListValue(qftWaitQueue, lngRowIndex, "����", strRoom)
            Call ucPacsQueue.Populate(qftWaitQueue)
        End If
        
        RaiseEvent OnResotre(lngAdviceID, strRoom)
    End If
End Sub


Private Sub ucPacsQueue_OnCreateQueueNo(ByVal lngQueueId As Long, ByVal strQueueName As String, strQueueNo As String)
'�ŶӺ��������¼�
    Dim strRoom As String
    Dim strCodeTag As String
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    If strQueueNo = "" Then Exit Sub
    
    strCodeTag = ""
    If strQueueName <> mstrCurTechnicGroupName And strQueueName <> mstrCurDeptName & "-" & M_STR_NOT_ALLOT_TECHNIC Then
        '��ȡִ�м�ǰ׺
        strRoom = Replace(strQueueName, mstrCurDeptName & "-", "")
        strCodeTag = zlGetTechnicRoomCodeNo(strRoom, mlngCurDeptId)
    Else
        
        '����ǰ������Ŷӣ���û���Ŷӱ��
        If mlngQueueNoWay = 1 Then
            '��ȡ������Ŷӱ��
            strSql = "select a.����,a.����ǰ׺ from Ӱ��ִ�з��� a " & _
                    " where a.����ID=[1] and a.����=[2]"
            Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ���ŷ���ǰ׺", mlngCurDeptId, Replace(mstrCurTechnicGroupName, mstrCurDeptName & "-", ""))
                    
            If rsData.RecordCount > 0 Then
                strCodeTag = Nvl(rsData!����ǰ׺)
            End If
        End If
    End If
    
    '��Ҫ�����Ŷӱ��
    Call ucPacsQueue.QueueOper.UpdateQueue(lngQueueId, "�Ŷӱ��=''" & strCodeTag & "''")
        
    strQueueNo = strCodeTag & strQueueNo
End Sub

Public Sub CloseQueueQuick()
    If Not ucPacsQueue Is Nothing Then
        ucPacsQueue.CloseQueueQuick
    End If
End Sub

Public Sub OpenQueueQuick(ByVal strTechnics As String, objOwer As Object)
    Call zlRefreshQueueData(strTechnics)
    
    If Not ucPacsQueue Is Nothing Then
        ucPacsQueue.OpenQueueQuick objOwer
    End If
End Sub
