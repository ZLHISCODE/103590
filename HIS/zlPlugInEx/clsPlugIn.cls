VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'---------------------------------------------------------------------------
'���ܣ���չ�ӿڲ��ģ��
'���Ҫ������ʹ�� ��ʵ�� GetUserName ����
'---------------------------------------------------------------------------
Public gblnMec As Boolean '�Ƿ�����չ������鲡����ҳ
Public gblnLoadMec As Boolean '�Ƿ�����չ�������ز�����ҳ����

Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int���� As Integer = -1)
'���ܣ��ӿڳ�ʼ��
'������cnOracle=�����
'      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
    Set gcnOracle = cnOracle
    
    gblnMec = True '�����Զ����鲡����ҳ����
    gblnLoadMec = True '�����Զ�����ز�����ҳ����
End Sub

Public Sub Terminate(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int���� As Integer = -1)
'���ܣ��ӿ���ֹ
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
     
End Sub

Public Function GetFuncNames(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int���� As Integer = -1, Optional strReserve As String) As String
'���ܣ���ȡ��ҹ����б����ýӿ�һ���ڳ�ʼ��֮�����һ�Σ�����ʱ����Ƶ�ȿ�����ͬ�����������ͬ
'������lngSys,lngModual=��ǰ���ýӿڵ��ϼ�ϵͳ�ż�ģ��ţ���������ʱ lngModual=1070����int���ϲ���
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      strReserve=��������,������չʹ��
'      ҽ��վ�еĵ���ʱ,����Ϊ��ǰִ�е�ǰ��ҽ��ID
'���أ���ҹ����б�����ʽΪ"������1,������2,������3,..."
'      ���ĳ��������Ҫ�ڲ����л�(��ˢ��)ʱ�Զ����ã����ڹ�����ǰ�����"Auto:"ǰ׺�����ֹ���ֻ������һ����

'����˵�������������Ա����ģ�����ʱֻ����lngSys��lngModual���������е�lngModual��0
'          �Զ����ù�������Ա����ģ�����ʱ�����Զ�����
    Dim strFuncNames As String
    
    Dim str�˵��� As String '�˵����з��õĹ��ܣ��˵�����Ӧ�ð������й������� str�˵��� = str������ �� str������
    Dim str������ As String
    Dim str������ As String
    
    Select Case lngModual
    Case mҽ������ģ��
        '��ҽ������վ������ҳǩ
        strFuncNames = "����Ex"
        
        str�˵��� = "����Ex"
        str������ = "|:InTool:����Ex"
        str������ = "InTool:����Ex"
        strReserve = "<menubar>" & str�˵��� & "</menubar><toolbar>" & str������ & "</toolbar><popbar>" & str������ & "</popbar>"
    Case m����ҽ��ģ��
        If int���� = 0 Then
            '������ҽ������վ��ҽ��ҳǩ
            strFuncNames = "����1,����2,Auto:����3"
            
            str�˵��� = "����Ex"
            str������ = "|:InTool:����Ex"
            str������ = "InTool:����Ex"
            strReserve = "<menubar>" & str�˵��� & "</menubar><toolbar>" & str������ & "</toolbar><popbar>" & str������ & "</popbar>"
        Else
            '��ҽ����������ҽ��ҳǩ
            strFuncNames = "����1,����2,Auto:����3"
            
            str�˵��� = "����Ex"
            str������ = "|:InTool:����Ex"
            str������ = "InTool:����Ex"
            strReserve = "<menubar>" & str�˵��� & "</menubar><toolbar>" & str������ & "</toolbar><popbar>" & str������ & "</popbar>"
        End If
    Case mסԺҽ��ģ��
        If int���� = 0 Then
            '��סԺҽ������վ��ҽ��ҳǩ
            strFuncNames = "����Ex"
            
            str�˵��� = "����Ex"
            str������ = "|:InTool:����Ex"
            str������ = "InTool:����Ex"
            strReserve = "<menubar>" & str�˵��� & "</menubar><toolbar>" & str������ & "</toolbar><popbar>" & str������ & "</popbar>"
        ElseIf int���� = 1 Then
            '���ϰ滤ʿ����վ��ҽ��ҳǩ/�°滤ʿվ�������񴰿ڣ�ҽ��ҳǩ
            strFuncNames = "����Ex"
            
            str�˵��� = "����Ex"
            str������ = "|:InTool:����Ex"
            str������ = "InTool:����Ex"
            strReserve = "<menubar>" & str�˵��� & "</menubar><toolbar>" & str������ & "</toolbar><popbar>" & str������ & "</popbar>"
        Else
            '��ҽ������վ��ҽ��ҳǩ
            strFuncNames = "����Ex"
            
            str�˵��� = "����Ex"
            str������ = "|:InTool:����Ex"
            str������ = "InTool:����Ex"
            strReserve = "<menubar>" & str�˵��� & "</menubar><toolbar>" & str������ & "</toolbar><popbar>" & str������ & "</popbar>"
        End If
    Case m�ٴ�·��ģ��
        If int���� = 0 Then
            '��סԺҽ������վ���ٴ�·��ҳǩ
            strFuncNames = "����Ex"
            
            str�˵��� = "����Ex"
            str������ = "|:InTool:����Ex"
            str������ = "InTool:����Ex"
            strReserve = "<menubar>" & str�˵��� & "</menubar><toolbar>" & str������ & "</toolbar><popbar>" & str������ & "</popbar>"
        Else
            '����ʿ����վ���ٴ�·��ҳǩ/�°滤ʿվ�������񴰿ڣ��ٴ�·��ҳǩ
            strFuncNames = "����Ex"
            str�˵��� = "����Ex"
            str������ = "|:InTool:����Ex"
            str������ = "InTool:����Ex"
            strReserve = "<menubar>" & str�˵��� & "</menubar><toolbar>" & str������ & "</toolbar><popbar>" & str������ & "</popbar>"
        End If
    Case m�°滤ʿվ
        '���°滤ʿ����վ������
        strFuncNames = "����Ex"
        
        str�˵��� = "����Ex"
        str������ = "|:InTool:����Ex"
        str������ = "InTool:����Ex"
        strReserve = "<menubar>" & str�˵��� & "</menubar><toolbar>" & str������ & "</toolbar><popbar>" & str������ & "</popbar>"
    Case m����ģ��
        '��������дҳ��
        strFuncNames = "����Ex"
    Case m��Ա����ģ��
        '����Ա����ҳ��
        strFuncNames = "����Ex"
    Case m������Ĺ���
        '��������Ĺ���
        strFuncNames = "����Ex1,����Ex2,InTool:����Ex3"
    Case m����ܼ�Ǽ�
        '������ܼ�Ǽ�
        strFuncNames = "����Ex1,����Ex2,InTool:����Ex3"
    Case m���ֿ�ִ��
        '�����ֿ�ִ��
        strFuncNames = "����Ex1,����Ex2,InTool:����Ex3"
    Case m������Ǽ�
        '��������Ǽ�
        strFuncNames = "����Ex1,����Ex2,InTool:����Ex3"
    End Select
    
    GetFuncNames = strFuncNames
End Function

Public Sub ExecuteFunc(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strFuncName As String, ByVal lngPatiID As Long, ByVal varRecId As Variant, _
    ByVal varKeyId As Variant, Optional ByVal strReserve As String, Optional ByVal int���� As Integer = -1)
'���ܣ�ִ��ĳ����ҹ���
'������lngSys,lngModual=��ǰ���ýӿڵ��ϼ�ϵͳ�ż�ģ��ţ���������ʱ lngModual��1070����int���ϲ���
'      strFuncName=Ҫִ�еĹ��������ù�����֮ǰ��GetFuncNames��������
'      lngPatiID=��ǰ����ID,��Ա��������ʱ���������ԱID��
'      varRecId=���ֻ����ַ����������ﲡ�ˣ�Ϊ��ǰ�Һŵ��Ż��߹Һ�ID����סԺ���ˣ�Ϊ��ǰסԺ��ҳID
'      varKeyId=���ֻ����ַ�������ǰ�Ĺؼ�ҵ������Ψһ��ʶID����ҽ��ID
'           ��������ʱ varKeyId=���Ӳ�����¼.id���������༭δ����ǰ��ֵ=0,�������޸ġ���ʱִ����ֵ��
'           �ٴ�·�� varKeyId ����·��id
'      strReserve=��������,������չʹ�� ���ǲ�������ʱ strReserve�� �����ļ��б�.id
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)

'����˵�������������Ա����ģ�����ʱֻ����lngSys��lngModual��strFuncName��lngPatiID��varRecId��varKeyId��
'          lngPatiID������Աid�����lngPatiID��0��ʾ������δѡ����Ա�����е�varRecId��varKeyId��ֵ���Ǵ���0

    'ע�⣺�ڲ��÷������󣬹��ܵ�ִ��ʱ��Ӧ���ȶԹ����������жϡ�
    '�ȼ�鹦���Ƿ��ڱ������zlPlugIn���У�Ȼ����ȥִ����չ���
    
    Select Case lngModual
        Case m����ҽ��ģ��
            If int���� = 0 Then
                MsgBox strFuncName & "ִ�гɹ���" '����ҽ��վҽ��ҳ��
            ElseIf int���� = 2 Then
                MsgBox strFuncName & "ִ�гɹ���" 'ҽ��վ����ҽ��ҳ��
            End If
        Case mסԺҽ��ģ��
            If int���� = 0 Then
                MsgBox strFuncName & "ִ�гɹ���" 'סԺҽ��վҽ��ҳ��
            ElseIf int���� = 2 Then
                MsgBox strFuncName & "ִ�гɹ���" 'ҽ��վסԺҽ��ҳ��
            End If
        Case mסԺ��ʿվģ��
            MsgBox strFuncName & "ִ�гɹ���" 'סԺ��ʿվҽ��ҳ��
        Case m�ٴ�·��ģ��
            If int���� = 0 Then
                MsgBox strFuncName & "ִ�гɹ���" 'סԺҽ��վ�ٴ�·��ҳ��
            ElseIf int���� = 1 Then
                MsgBox strFuncName & "ִ�гɹ���" 'סԺ��ʿվ�ٴ�·��ҳ��
            End If
        Case m����ģ��
            MsgBox strFuncName & "ִ�гɹ���" '������дҳ��
        Case m��Ա����ģ��
            MsgBox strFuncName & "ִ�гɹ���" '��Ա����ҳ��
        Case mҽ������ģ��
            MsgBox strFuncName & "ִ�гɹ���" 'ҽ��վҽ������ҳ��
        Case m������Ĺ���
            MsgBox strFuncName & "ִ�гɹ���" '������Ĺ���
        Case m����ܼ�Ǽ�
            MsgBox strFuncName & "ִ�гɹ���" '����ܼ�Ǽ�
        Case m���ֿ�ִ��
            MsgBox strFuncName & "ִ�гɹ���" '���ֿ�ִ��
        Case m������Ǽ�
            MsgBox strFuncName & "ִ�гɹ���" '������Ǽ�
    End Select
End Sub
 
Public Function GetFormCaption(ByVal lngSys As Long, ByVal lngModual As Long) As String
'���ܣ���סԺҽ������վ�����ӿ�Ƭ�����ƣ����ŷָ
        Exit Function '��ʹ�ô˹������޸Ľӿ��ڲ�ʵ��
    Select Case lngModual
    Case mסԺҽ������վ
        GetFormCaption = "Test3"
    Case mסԺ��ʿ����վ
        GetFormCaption = "Test3"
    Case m����ҽ������վ
        GetFormCaption = "Test3"
    Case mҽ������վ
        GetFormCaption = "Test3"
    End Select
End Function

Public Function GetForm(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strName As String) As Object
'���ܣ���������ֵ���ض��󡣹���վ�������ʱ����
        Exit Function '��ʹ�ô˹������޸Ľӿ��ڲ�ʵ��
    Select Case lngModual
    Case mסԺҽ������վ
        Select Case strName
        Case "Test3"
            Set GetForm = New frmDockEx
        End Select
    Case mסԺ��ʿ����վ
        Select Case strName
        Case "Test3"
            Set GetForm = New frmDockEx
        End Select
    Case m����ҽ������վ
        Select Case strName
        Case "Test3"
            Set GetForm = New frmDockEx
        End Select
    Case mҽ������վ
        Select Case strName
        Case "Test3"
            Set GetForm = New frmDockEx
        End Select
    End Select
 
End Function

Public Sub RefreshForm(ByVal lngSys As Long, ByVal lngModual As Long, ByRef objForm As Object, ByVal strName As String, ByVal lngPatiID As Long, ByVal str�Һŵ� As String, ByVal lng��ҳID As Long, _
    Optional ByVal blnMoved As Boolean, Optional ByVal lng�������ID As Long, Optional ByVal lngҽ������ID As Long, Optional ByVal lng����ID As Long, Optional ByVal lng����id As Long, _
    Optional ByVal lngǰ��ID As Long, Optional ByVal int״̬ As Integer, Optional ByVal intִ��״̬ As Integer, Optional ByVal lng·��״̬ As Long = -1)
'���ܣ�ˢ���ڲ�����
'������strName ��ǰ��������
'      lngPatiID ����id,str�Һŵ� �Һŵ������ﲡ�˲��з���Ϊ�գ�lng��ҳID סԺ���˲��з���Ϊ0
'      blnMoved=�ò��˵������Ƿ���ת��
'      lng�������ID=�����ǰҽ��վ�ǻ��ﲡ�ˣ���Ϊ�������ID�������ҽ��վ���ã���Ϊҽ������ID
'      lngҽ������ID=ҽ��վ����ID
'      lng����ID��lng����ID=����5-���ת�Ʋ��ˡ�ʱΪ����ԭ������ԭ����
'      lngǰ��ID=����ҽ��վ����ʱ����
'      int״̬=0-��Ժ����,1-Ԥ��Ժ����,2-��Ժ����,3-�����ﲡ��(��Ժ),4-�ѻ��ﲡ��,5-���ת�Ʋ���(��Ժ),6-ת�ƻ�ת��������ס����(������Ժ����ס)
'      intִ��״̬=����ҽ��վ����ʱ����,ҽ��ִ����Ŀ��ִ��״̬
'      lng·��״̬=-1:δ����,0-�����ϵ���������1-ִ���У�2-����������3-�������
    Exit Sub '��ʹ�ô˹������޸Ľӿ��ڲ�ʵ��
    If objForm Is Nothing Then Exit Sub
    Call objForm.RefreshInSide
       
End Sub

Public Function GetButtomName(ByVal lngSys As Long, ByVal lngModual As Long, ByRef objForm As Object, ByVal strName As String) As String
'���ܣ����ӵ��˵��͹������ϵİ�ť���ƣ��˵��ϼ�Ϊ ��չ����������ֱ�ӼӰ�ť
'������strName ��ǰ��������
        Exit Function '��ʹ�ô˹������޸Ľӿ��ڲ�ʵ��
    If objForm Is Nothing Then Exit Function
    GetButtomName = objForm.GetInSideFunc
 
End Function

Public Function ExeButtomClick(ByVal lngSys As Long, ByVal lngModual As Long, ByRef objForm As Object, ByVal strName As String, ByVal strFuncName As String, _
    ByVal lngPatiID As Long, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String, Optional ByVal lngҽ��ID As Long, Optional ByVal strReserve As String) As Boolean
'���ܣ�ִ�д����ڲ��Ĺ���
'������strName ��ǰ�������ƣ�strFuncName ������(�����й�������)
'      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      lngPatientID=����ID,lng��ҳID,str�Һŵ�
'      lngҽ��ID=ֻ��ҽ��վ������ʱ�ᴫ�룬 ����û��ֵ
    Exit Function '��ʹ�ô˹������޸Ľӿ��ڲ�ʵ��
    If objForm Is Nothing Then Exit Function
    ExeButtomClick = objForm.ExecuteFunc(strFuncName)
End Function

Public Function GetUserName() As String
'���ܣ���ȡ��λ����
'˵�����û�����ʱ��ʾ������
'��  GetUserName = "������������ҽԺ��Ϣϵͳ"

    GetUserName = ""
End Function
 
 
'*****************************************************************
'                   ����˵����ע��
'*****************************************************************
'-------�Ӵ��п�ʼ���µĽӿ���ʹ����չ����ʱ����������------------
'1.���½ӿڽ�֧�ֻ�����չ�����ӿ�ֻ����չһ�Σ�
'      ���� QueryPatiInfo ����չΪһ����������zlPlugInQP����zlPlugInQP��ʵ����QueryPatiInfo��������ʱ�Ͳ���������һ����չ�������������չ�ˡ�
'      ԭ��a)�ӿڱ����������������Ӱ�죬����ֵ���������Ͳ����ȣ�
'            b)����ӿڶ����չ�󣬷���ֵ�����ݲ��ܵ��ӣ�
'            c)������ڵ�����չ�������ʱ���Ǳ����Ե��ã����õ��Ⱥ�˳���Ӱ�����ս����
'               �磺��Fun1����������չ����(����A��B)ʵ�֣����A��B����ͬʱ���ڣ���A.Fun1��B.Fun1���ᱻ���ã���ִ�еĽӿڵķ���ֵ�Ḳ��ǰһ������A��B��AʧЧ����֮BʧЧ��
'                   ������ڵ�����չ�����ӿڵ�˳��Ŀǰ�ǲ���Ԥ���ж��ģ����Ծ�Ҫ��ֻ�ܱ���չһ�Ρ�
'2.��չ����ʵ�ֽӿ������Ĺ�ϵ����һ�����Ϲ�ϵ��˵����
'  ALL={�ӿ�1,�ӿ�2,....};���нӿ�
'  EX1={�ӿ�1,�ӿ�3};��չ����EX1ʵ�ֵĽӿ�;
'  EX2={�ӿ�5,�ӿ�6,�ӿ�8,�ӿ�9};��չ����EX2ʵ�ֵĽӿ�;
'  .
'  .
'  .
'  EXn={�ӿ�n,.....};��չ����EX1ʵ�ֵĽӿ�;
'  EXn�� ALL �� EX1+EX2+...+EXn �� ALL �� EXn��EXn-1 = �գ�
'3.���ڲ�ʵ�ֵĽӿ���ɾ�������ӿ�������

'����˵���翴��������ѯ���������ܲ��з�����
 
Public Sub LogInAfter()
'���ܣ���½����̨֮��

End Sub

Public Sub LogOutBefore()
'���ܣ��رպ���̨����ע����ǰ
    
End Sub

Public Function QueryPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, _
    ByVal str���� As String, ByRef rsInfo As ADODB.Recordset) As Boolean
'˵���������ģ��
    QueryPatiInfo = True
End Function

Public Function CommitPatiInfo(ByVal str���� As String, ByRef rsInfo As ADODB.Recordset) As Boolean
'˵���������ģ��
    CommitPatiInfo = True
End Function
 
Public Function ClinicReceive(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long) As Boolean
'˵���������ģ��
    ClinicReceive = True
End Function

Public Function ClinicFinish(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long) As Boolean
'˵���������ģ��
    ClinicFinish = True
End Function

Public Function DiagnosisEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lngԭ���id As Long, Optional ByVal int���� As Integer = -1) As Boolean
'˵���������ģ��
    DiagnosisEnter = True
End Function

Public Function DiagnosisOtherEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lngԭ���id As Long, Optional ByVal int���� As Integer = -1) As Boolean
'˵���������ģ��
    DiagnosisOtherEnter = True
End Function

Public Function DiagnosisDeleted(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional ByVal int���� As Integer = -1) As Boolean
'˵���������ģ��
    DiagnosisDeleted = True
End Function

Public Function AdviceEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional ByVal int���� As Integer = -1) As Boolean
'˵���������ģ��
    AdviceEnter = True
End Function

Public Function AdviceEditAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional ByVal int���� As Integer = -1) As Boolean
'˵���������ģ��
    AdviceEditAfter = True
End Function

Public Function AdviceSave(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional int���� As Integer = -1) As Boolean
'˵���������ģ��
    AdviceSave = True
End Function

Public Function AdviceSend(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngSendID As Long) As Boolean
'˵���������ģ��
    AdviceSend = True
End Function

Public Function AdviceSendEnd(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strSendIDs As String) As Boolean
'˵���������ģ��
    AdviceSendEnd = True
End Function

Public Function AdviceRowChange(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional int���� As Integer = -1) As Boolean
'˵���������ģ��
    AdviceRowChange = True
End Function

Public Function AdviceDeletBefor(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int���� As Integer = -1) As Boolean
'˵���������ģ��
    AdviceDeletBefor = True
End Function

Public Function AdviceDeleted(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int���� As Integer = -1) As Boolean
'˵���������ģ��
    AdviceDeleted = True
End Function

Public Function AdviceRevoked(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int���� As Integer = -1) As Boolean
'˵���������ģ��
    AdviceRevoked = True
End Function

Public Function AdviceRevokedBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, ByVal int���� As Integer, ByRef strMsg As String, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    AdviceRevokedBefore = True
End Function
 
Public Function AdviceRollSend(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, ByVal int���� As Integer, ByRef strMsg As String, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    AdviceRollSend = True
End Function
 
Public Function AdviceRollSendBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, ByVal int���� As Integer, ByRef strMsg As String, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    AdviceRollSendBefore = True
End Function

Public Function AddEMRBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant) As Boolean
'˵���������ģ��
    AddEMRBefore = True
End Function

Public Function BeforeEMRDocSave(ByVal lngSource As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngFileId As Long, ByVal strTitle As String, ByVal strContent As String, ParamArray arrInput() As Variant) As Boolean
'���ܣ����Ӳ�������ǰ����
'������lngSource 1 ���� 2 סԺ
'      lngPatientID=����ID,lngPageId=����ID/
'      lngFileId=�����ļ��б�.id
'      arrInput�����飺��չ����������չʹ�ã�ͨ���±��һֱ����ʱ���ܻᱨ���ִ���9���±�Խ��/13�����Ͳ�ƥ�䣺����ʾ����û�д���
'           ���Է��ʲ���ǰӦ���жϣ�If TypeName(arrInput(i)) <> "Error" Then ��������������ʹ������������������Ϊ�������û�д���
'           arrInput Ŀǰδ�����κβ���
    BeforeEMRDocSave = True: Exit Function
End Function

Public Function SignEMRBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant) As Boolean
'˵���������ģ��
    SignEMRBefore = True
End Function

Public Sub SignEMRAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant)
'˵���������ģ��
End Sub

Public Function PathImportBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng·��ID As Long, Optional ByVal lng�汾��ID As Long, Optional ByVal int���� As Integer = -1, _
    Optional ByVal lngDiagnosisType As Long, Optional ByVal lng����ID As Long, Optional ByVal lng���ID As Long, _
    Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    PathImportBefore = True
End Function

Public Function PathImportAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng·��ID As Long, Optional ByVal lng�汾��ID As Long, Optional ByVal int���� As Integer = -1, _
    Optional ByVal lngDiagnosisType As Long, Optional ByVal lng����ID As Long, Optional ByVal lng���ID As Long, _
    Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    PathImportAfter = True
End Function

Public Function PathEvaluateBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng·����¼ID As Long, Optional ByVal lng�׶�ID As Long, Optional ByVal int���� As Integer = -1, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
   PathEvaluateBefore = True
End Function

Public Function PathEvaluateAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng·����¼ID As Long, Optional ByVal lng�׶�ID As Long, Optional ByVal int���� As Integer = -1, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
   PathEvaluateAfter = True
End Function

Public Function DiagnosisEnterIn(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lngԭ���id As Long, Optional ByVal int���� As Integer = -1, _
    Optional ByVal int������� As Integer, Optional ByVal int��ϴ��� As Integer, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    DiagnosisEnterIn = True
End Function

Public Function GetFormHwnd(Optional ByVal strReserve As String) As Long
'˵���������ģ��
    GetFormHwnd = 0
End Function

Public Function PatiInfoSaveBefore(ByVal lng����ID As Long, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    PatiInfoSaveBefore = True
End Function

Public Function PatiInfoSaveAfter(ByVal lng����ID As Long, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    PatiInfoSaveAfter = True
End Function

'--ҵ�����۽ӿ�
Public Function PatiPrePayAfter(ByVal lng����ID As Long, ByVal intԤ������ As Integer, ByVal lngԤ��ID As Long, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    PatiPrePayAfter = True
End Function

Public Function OutPatiRegisterAfter(ByVal lng����ID As Long, ByVal lng�Һ�ID As Long, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    OutPatiRegisterAfter = True
End Function
 
Public Function OutPatiExseAfter(ByVal lng����ID As Long, ByVal lng����ID As Long, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    OutPatiExseAfter = True
End Function
  
Public Function OutPatiInvoicePrintAfter(ByVal lng����ID As Long, ByVal str����IDs As String, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    OutPatiInvoicePrintAfter = True
End Function
 
Public Function OutPatiMedicineAfter(ByVal lng����ID As Long, ByVal strNO As String, ByVal int���� As Integer, ByVal lngҩ��ID As Long, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    OutPatiMedicineAfter = True
End Function
 
Public Function DrugStuffWorkNoramal(ByVal lngModual As Long, ByVal strFuncName As String, ByVal lng�ⷿID As Long, ByVal strNO As String, ByVal int���� As Integer, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    DrugStuffWorkNoramal = True
End Function

Public Function InPatiCheckInAfter(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    InPatiCheckInAfter = True
End Function
 
Public Function InPatiCheckInBranchAfter(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    InPatiCheckInBranchAfter = True
End Function
 
Public Function InPatiCheckOutAfter(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    InPatiCheckOutAfter = True
End Function

Public Function InPatiCashierAfter(ByVal lng����ID As Long, ByVal lng����ID As Long, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    InPatiCashierAfter = True
End Function

Public Sub AdviceBeforeSend(ByVal strEndDate As String, ByRef rsData As ADODB.Recordset, ByRef strAdviceIDs As String, ByRef strMsg As String)
'˵���������ģ��
End Sub

Public Function AdviceCheckSendFee(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal cur�ϼ� As Currency, Optional ByVal int���� As Integer, Optional strReserve As String) As Boolean
'˵���������ģ��
    AdviceCheckSendFee = True
End Function


Public Function PatiFeeAuditingAndCancelCheck(ByVal lngModual As Long, _
    ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal blnCancel As Boolean = False, _
    Optional strExpendXML As String) As Boolean
'˵���������ģ��
    PatiFeeAuditingAndCancelCheck = True
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Ѫ����ؽӿ�
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function BloodSendAfter(ByVal lngAdviceID As Long, ByVal strBloodIDs As String, Optional strReserve As String) As Boolean
'˵���������ģ��
    BloodSendAfter = True
End Function

Public Function BloodSendReverseAfter(ByVal lngAdviceID As Long, ByVal strBloodIDs As String, Optional strReserve As String) As Boolean
'˵���������ģ��
    BloodSendReverseAfter = True
End Function

Public Function BloodScrapAfter(ByVal strBloodIDs As String, Optional ByVal intDelete As Integer = 0) As Boolean
'˵���������ģ��
    BloodScrapAfter = True
End Function

Public Function BloodStoreImport(ByRef lng��ⷽʽ As Long, ByRef rsData As ADODB.Recordset)
'˵���������ģ��
    BloodStoreImport = True
End Function

Public Function EMPI_QueryPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal rsIn As ADODB.Recordset, ByRef rsOut As ADODB.Recordset, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    EMPI_QueryPatiInfo = False
End Function

Public Function EMPI_AddPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngClinicID As Long, ByRef strMsg As String, Optional ByRef strReserve As String) As Long
'˵���������ģ��
    EMPI_AddPatiInfo = 1
End Function

Public Function EMPI_ModifyPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngClinicID As Long, ByRef strMsg As String, Optional ByRef strReserve As String) As Long
'˵���������ģ��
    EMPI_ModifyPatiInfo = 1
End Function

Public Function PatiValiedCheck(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngType As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal strPatiInforXML As String, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    PatiValiedCheck = True
End Function

Public Function PatiRegTriageCheck(lngSys As Long, ByVal lngModule As Long, bytFun As Byte, ByVal lng�Һ�ID As Long, _
    Optional ByVal strҽ�� As String, Optional ByVal str���� As String, Optional str�º��� As String = "", Optional lng�³����¼ID As Long = 0) As Boolean
'˵���������ģ��
    PatiRegTriageCheck = True
End Function

Public Function CheckChargeItem(lngSys As Long, ByVal lngModule As Long, ByVal intType As Integer, ByVal intMode As Integer, _
    ByRef rsDetail As ADODB.Recordset, Optional strExpend As String = "") As Boolean
'˵���������ģ��
    CheckChargeItem = True
End Function

Public Function CheckMecInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByRef colCtl As Collection, ByRef strMsg As String, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    CheckMecInfo = True
End Function

Public Function LoadMecInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatiType As Long, ByRef colCtl As Collection) As Boolean
'˵���������ģ��
    LoadMecInfo = True
End Function

Public Function AdviceMakeFee(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strCurFee As String, ByVal rsNowFee As ADODB.Recordset, ByRef dblCount As Double, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
        dblCount = -1
        AdviceMakeFee = True
End Function

Public Function AdviceLoadApplyCustom(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatiID As Long, ByVal lng����ID As String, ByVal lng�ļ�ID As Long, ByRef strXML As String, Optional ByVal lngҽ��ID As Long, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    AdviceLoadApplyCustom = True
End Function

Public Function AdviceSaveApplyCustom(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatiID As Long, ByVal lng����ID As String, ByVal lng�ļ�ID As Long, ByRef strXML As String, ByRef objWeb As Object, Optional ByVal lngҽ��ID As Long, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    AdviceSaveApplyCustom = True

End Function

Public Function AdviceCheckApplyCustom(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatiID As Long, ByVal lng����ID As String, ByVal lng�ļ�ID As Long, ByVal strXML As String, ByRef objWeb As Object, Optional ByVal lngҽ��ID As Long, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    AdviceCheckApplyCustom = True
End Function

Public Function CheckPathInItem(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, ByRef lng·����ĿID As Long, ByRef str���� As String, Optional ByVal int���� As Integer = -1, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    CheckPathInItem = True
End Function

Public Function InPatiCancelOutBefore(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    InPatiCancelOutBefore = True
End Function

Public Function InPatiCheckInCancel(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngType As Long, Optional ByVal strReserve As String) As Boolean
'˵���������ģ��
    InPatiCheckInCancel = True
End Function

Public Function TransfusionCall(ByVal lngSys As Long, ByVal lngModual As Long, ByVal str�Һŵ� As String, ByVal lng��ˮ�� As Long, ByRef strContent As String, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    TransfusionCall = True
End Function

Public Function TransfusionShowPatiList(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lng����id As Long, ByVal BeginTime As Date, ByVal EndTime As Date, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    TransfusionShowPatiList = True
End Function

Public Function ShowClinicHelp(lngOwnHwnd As Long, ByVal int���� As Integer, ByVal int������Դ As Integer, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal str��ĿIDs As String, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    ShowClinicHelp = False
End Function

Public Function SetPassDiag(ByVal lngPatientID As Long, ByVal lng��ҳID As Long, ByRef rsDiag As ADODB.Recordset) As Boolean
'˵���������ģ��
    SetPassDiag = True
End Function

Public Function SetPassExams(ByVal lngPatientID As Long, ByVal lng��ҳID As Long, ByVal lng�Һ�ID As Long, ByRef strExamsInfo As String) As Boolean
'˵���������ģ��
    SetPassExams = True
End Function

Public Function BeforeVitalSignsSave(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng�ļ�ID As Long, ByVal strDataXML As String) As Boolean
'˵���������ģ��
    BeforeVitalSignsSave = True
End Function

Public Sub AfterSelectorReady(ByVal intType As Integer, ByVal strTitle As String, ByVal objList As Object, _
    Optional ByVal frmMain As Object, Optional ByVal strExt As String)
'˵���������ģ��
End Sub


Public Function AdviceEditBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int���� As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal str�Һŵ� As String, ByVal lngǰ��ID As Long, ByVal intӤ�� As Integer, ByVal lng�������ID As Long, ByVal lngΣ��ֵID As Long, ByVal lng����ҽ��ID As Long, Optional ByRef strReserve As String) As Boolean
'˵���������ģ��
    AdviceEditBefore = True
End Function

Public Function PeisDelItemCheck(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngTaskId As Long, _
    ByVal lngListId As Long, ByRef strMsg As String, Optional ByRef strReserve As String) As Boolean
'���ܣ������Ŀҳǩ��ɾ����ʱ����
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���(1252=����ҽ���´�,1253=סԺҽ���´�)
'      lngPatientID=����ID,lngTaskId=����id
'      lngListID=�嵥id
'      strMsg   ��ʾ��Ϣ���ӿ��ڲ���֯���ظ������������ʾ���մ�����ʾ
'      strReserve   ��չ����
'���أ�
    
    strMsg = "" '����Ҫ���Ʋ���ɾ����ʱ������strMsgΪ��ʾ����Ϣ
    PeisDelItemCheck = True
End Function

Public Function PeisAbortSignInCheck(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngTaskId As Long, _
    ByVal lngListId As Long, ByVal strNO As String, ByVal lngRecType As Long) As Boolean
'���ܣ�������Ĺ�����ȡ������ʱ���ã����ڼ�����ķ���״̬�Ǵӳ���У�黹�ǽӿ�У�鷵��
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���(1252=����ҽ���´�,1253=סԺҽ���´�)
'      lngPatientID=����ID,lngTaskId=����id
'      lngListID=�嵥id
'      strNo   ���õ��ݺ�
'      lngRecType   ���ü�¼����
'���أ�true-���ķ��ϼ��ͨ�� false-���ķ��ϼ�鲻ͨ��
 
    PeisAbortSignInCheck = True
End Function
