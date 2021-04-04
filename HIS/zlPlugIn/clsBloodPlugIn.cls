VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBloodPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int���� As Integer = -1)
'���ܣ��ӿڳ�ʼ��
'������cnOracle=�����
'      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      int����  ���ó���:-1
'      �������� lngModual=1070����int���ϲ���
    If Not gblnInited Then
        Set gcnOracle = cnOracle
        Call HandlePlugIn(1, lngSys, lngModual, cnOracle, int����)
        gblnInited = True
    End If
End Sub

Public Function GetFormCaption(ByVal lngSys As Long, ByVal lngModual As Long) As String
'���ܣ��ڿ�����Ѫ���������ӿ�Ƭ�����ƣ����ŷָ
    Dim strTmp As String, strThis As String
    'Exit Function '��ʹ�ô˹������޸Ľӿ��ڲ�ʵ��
    strTmp = GetFormCaptionEx(lngSys, lngModual)
    
    Select Case lngModual
    Case m������Ѫ����
        strThis = "Test1,Test2"
    End Select
    GetFormCaption = strThis & IIf(strTmp = "", "", "," & strTmp)
End Function

Public Function GetForm(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strName As String) As Object
'���ܣ���������ֵ���ض��󡣹���վ�������ʱ����
        'Exit Function '��ʹ�ô˹������޸Ľӿ��ڲ�ʵ��
    Select Case lngModual
    Case m������Ѫ����
        Select Case strName
        Case "Test1"
            Set GetForm = New frmDockEx
        Case "Test2"
            Set GetForm = New frmDockEx
        Case Else
            Set GetForm = GetFormEx(lngSys, lngModual, strName)
        End Select
    End Select
End Function

Public Function GetFuncNames(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int���� As Integer = -1, Optional strReserve As String) As String
'���ܣ���ȡ��ҹ����б����ýӿ�һ���ڳ�ʼ��֮�����һ�Σ�����ʱ����Ƶ�ȿ�����ͬ�����������ͬ
'������lngSys,lngModual=��ǰ���ýӿڵ��ϼ�ϵͳ�ż�ģ��ţ���������ʱ lngModual=1070����int���ϲ���
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS),3-����Һŵ���,4-�����շѵ���,5-����ҽ���´������չ,6-��ҳ�����չ
'      strReserve=��������,������չʹ��

'���أ���ҹ����б�����ʽΪ"������1,������2,������3,..."
'       "Auto"   �ؼ���֧���°滤ʿվ-�������񴰿ڣ�����ҽ��(�����סԺ)Dockҳǩ���桢�ٴ�·��Dockҳǩ���棬ҽ������Dockҳǩ���棬���л����˺���Զ�ִ�иù��ܣ��ù���ֻ������һ��
'       "InTool" ��ʾ��ǰ������ܣ��Ƿ���ֱ����ʾ������������ʽ�����˵���ʾ��
'˵�������������Ա����ģ�����ʱֻ����lngSys��lngModual���������е�lngModual��0

'strReserve ��չ�����ڿؼ�������ZLHIS����վ�в˵�����ʾ��ʽ��������Դ˲�����ֵ������������ǰ����ֵ��ʽ
'Ӧ�ó���:
'   1.������Ѫ����
'�������л�����Ӧҳǩ����Ч:
'      a.���渴��Dockҳǩ���棬[�˵�����ҽ�������߰�ť��;���������Ҽ������� ҽ��  �˵�����]��
'      b.��Ѫ��¼Dockҳǩ���棬[�˵�����·�������� ��ť�£����������Ҽ������� ·�� �˵�����]��
'��ʽΪ: xml�ṹ "<menubar>...</menubar><toolbar>...</toolbar><popbar>...</popbar><hotkey>...</hotkey>"
'menubar   ����������˵���"����"��չʾ��ʽ��������ݶ��ŷָ"����1,����2,����3,����4"��
'          ÿ������ǰ����ǰ׺ ��|:Auto:InTool:����1��
'          "|"���ߣ���ʾ��ǰλ���в˵��ķָ���ǲ˵�֮ǰ�ķָ��ߣ�
'          "Auto"   �ؼ���֧���°滤ʿվ-�������񴰿ڣ�����ҽ��(�����סԺ)Dockҳǩ���桢�ٴ�·��Dockҳǩ���棬ҽ������Dockҳǩ���棬���л����˺���Զ�ִ�иù��ܣ��ù���ֻ������һ��
'          "InTool" ��ʾ��ǰ������ܣ��Ƿ���ֱ����ʾ������������ʽ�����˵���ʾ��
'toolbar  ���������湤������չʾ��ʽ��������ݶ��ŷָ"����1,����2,����3,����4"��
'         ÿ������ǰ����ǰ׺ ��|:InTool:����1��
'         "|"���ߣ���ʾ��ǰλ���в˵��ķָ���ǲ˵�֮ǰ�ķָ��ߣ�
'         "InTool" ��ʾ��ǰ������ܣ��Ƿ���ֱ����ʾ������������ʽ�����˵���ʾ��
'popbar   �Ҽ������˵�չʾ��ʽ��������ݶ��ŷָ"����1,����2,����3,����4"��
'         ÿ������ǰ����ǰ׺ ��|:InTool:����1��
'         "|"���ߣ���ʾ��ǰλ���в˵��ķָ���ǲ˵�֮ǰ�ķָ��ߣ�
'         "InTool" ��ʾ��ǰ������ܣ��Ƿ���ֱ����ʾ������������ʽ�����˵���ʾ��
'hotkey   �����п�ݽ���������|KeyCode|Shift ��ϳɳɣ�����"��ݼ�����1|116|0,��ݼ�����2|123|0,��ݼ�����3|49|4
'         ԭ���������� �������Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)�¼����д���
'         �ù��ܽ�����ҽ������վҽ���´����-ҽ���༭ ҳǩ��Ч��


    Dim strTmp As String, strFunc As String
    Dim strReserveOther As String
    
    Dim strFuncNames As String '���й������ƴ������ŷָ�ɺ�һ�� Auto�ؼ��֣��ӿ���Ӧ�öԴ˱�����ֵ���Լ�����ǰ�ϰ汾��Ʒ
    
    Dim str�˵��� As String '�˵����з��õĹ��ܣ��˵�����Ӧ�ð������й������� str�˵��� = str������ �� str������
    Dim str������ As String
    Dim str������ As String
    Dim str��ݼ� As String '�̶���ʽ  F5--��ݼ�����1;F12--��ݼ�����2,Alt+1 --��ݼ�����3  ��Ӧ�Ĵ� ��"��ݼ�����1|116|0,��ݼ�����2|123|0,��ݼ�����3|49|4"
    
'==================================================================================================================================================
'***************************************************����������Ϊʾ������ɽ��е���*****************************************************************
    
    Select Case lngModual
    Case m������Ѫ����_���渴��
        '��������Ѫ����_���渴�����
        strFuncNames = "����1,����2,����3,����4"
        
        str�˵��� = "����1,����2,����3,����4"
        str������ = "InTool:����1,|:InTool:����3"
        str������ = "InTool:����1"
        strReserve = "<menubar>" & str�˵��� & "</menubar><toolbar>" & str������ & "</toolbar><popbar>" & str������ & "</popbar>"
    Case m������Ѫ����_��Ѫ��¼
        '��������Ѫ����_��Ѫ��¼����
        strFuncNames = "����1,����2,����3,����4"
        
        str�˵��� = "����1,����2,����3,����4"
        str������ = "InTool:����1,|:InTool:����3"
        str������ = "InTool:����1"
        strReserve = "<menubar>" & str�˵��� & "</menubar><toolbar>" & str������ & "</toolbar><popbar>" & str������ & "</popbar>"
    Case m���ҷ�Ѫ����
        '�����ҷ�Ѫ��������
        strFuncNames = "����1,����2,����3,����4"
        
        str�˵��� = "����1,����2,����3,����4"
        str������ = "InTool:����1,|:InTool:����3"
        str������ = "InTool:����1"
        strReserve = "<menubar>" & str�˵��� & "</menubar><toolbar>" & str������ & "</toolbar><popbar>" & str������ & "</popbar>"
    Case mѪҺ��Ѫ��Ӧ
        '����Ѫ��Ӧ������
        strFuncNames = "����1,����2,����3,����4"
        
        str�˵��� = "����1,����2,����3,����4"
        str������ = "InTool:����1,|:InTool:����3"
        str������ = "InTool:����1"
        strReserve = "<menubar>" & str�˵��� & "</menubar><toolbar>" & str������ & "</toolbar><popbar>" & str������ & "</popbar>"
    End Select
    
'***************************************************����������Ϊʾ������ɽ��е���*****************************************************************
'==================================================================================================================================================

    strTmp = ""
    strReserveOther = ""
    Call HandlePlugIn(2, lngSys, lngModual, , int����, strReserveOther, strTmp)
    If strReserveOther <> "" Then strReserve = strReserve & strReserveOther
    GetFuncNames = strFuncNames & IIf(strTmp = "", "", "," & strTmp)
End Function

Public Sub ExecuteFunc(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strFuncName As String, ByVal lngPatiID As Long, ByVal varRecId As Variant, _
    ByVal varKeyId As Variant, Optional ByVal strReserve As String, Optional ByVal int���� As Integer = -1, Optional ByVal strXML As String = "")
'���ܣ�ִ��ĳ����ҹ���
'������lngSys,lngModual=��ǰ���ýӿڵ��ϼ�ϵͳ�ż�ģ��ţ���������ʱ lngModual��1070����int���ϲ���
'      strFuncName=Ҫִ�еĹ��������ù�����֮ǰ��GetFuncNames��������
'      lngPatiID=��ǰ����ID
'      varRecId=���ֻ����ַ����������ﲡ�ˣ�Ϊ��ǰ�Һŵ��Ż��߹Һ�ID����סԺ���ˣ�Ϊ��ǰסԺ��ҳID
'      varKeyId=���ֻ����ַ�������ǰ�Ĺؼ�ҵ������Ψһ��ʶID���������Ѫ����Ϊ�����ҷ�Ѫ����Ϊ����Ѫ��ӦΪ
'      strReserve=��������,������չʹ��
'      int����  ���ó���:0-,1-������Ѫ����,2-���ҷ�Ѫ����,3-��Ѫ��Ӧ����

'      strXmlΪXML��ʽ���ַ���,����<ROOT><XXX>XXX</XXX></ROOT>
'       �������סԺҽ���������ʱ��strXml����Ľڵ㣺������Ŀ����
'       ����Ϊ5��6ʱ��strXml����Ľڵ㣺���ID������ID

'����˵�������������Ա����ģ�����ʱֻ����lngSys��lngModual��strFuncName��lngPatiID��varRecId��varKeyId��
'          lngPatiID������Աid�����lngPatiID��0��ʾ������δѡ����Ա�����е�varRecId��varKeyId��ֵ���Ǵ���0

    'ע�⣺�ڲ��÷������󣬹��ܵ�ִ��ʱ��Ӧ���ȶԹ����������жϡ�
    '�ȼ�鹦���Ƿ��ڱ������zlPlugIn���У�Ȼ����ȥִ����չ���
    
    Select Case lngModual
        Case m������Ѫ����_���渴��
            MsgBox strFuncName & "ִ�гɹ���" '������Ѫ����ҳ��
        Case m������Ѫ����_��Ѫ��¼
            MsgBox strFuncName & "ִ�гɹ���" '������Ѫ����ҳ��
        Case m���ҷ�Ѫ����
            MsgBox strFuncName & "ִ�гɹ���" '���ҷ�Ѫ����ҳ��
        Case mѪҺ��Ѫ��Ӧ
            'lngPatiID��varRecIdΪ0��varKeyId:��Ѫ��Ӧ��¼.�շ�ID
            MsgBox strFuncName & "ִ�гɹ���" 'ѪҺ��Ѫ��Ӧҳ��
    End Select
    Call HandlePlugIn(3, lngSys, lngModual, , int����, strReserve, strFuncName, lngPatiID, varRecId, varKeyId)
End Sub

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

Public Function GetButtomName(ByVal lngSys As Long, ByVal lngModual As Long, ByRef objForm As Object, ByVal strName As String) As String
'���ܣ����ӵ��˵��͹������ϵİ�ť���ƣ��˵��ϼ�Ϊ ��չ����������ֱ�ӼӰ�ť
'������strName ��ǰ��������
        'Exit Function '��ʹ�ô˹������޸Ľӿ��ڲ�ʵ��
    If objForm Is Nothing Then Exit Function
    GetButtomName = objForm.GetInSideFunc
 
End Function

Public Sub Terminate(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int���� As Integer = -1)
'���ܣ��ӿ���ֹ
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      int����=-1
    Call HandlePlugIn(4, lngSys, lngModual, , int����)
End Sub

Public Function AfterBloodPrepareRecept(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng����ID As Long) As Boolean
'���ܣ�������Ѫ�����һ��
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      intMode��1-�������룬2-ȡ����������
'      lngPatientID=����ID,lngClinicID=����ID���Һż�¼ID��/��ҳID(סԺ����ʱ)
'      lng����ID=ҽ��id(����ҽ��״̬.ҽ��id)

    AfterBloodPrepareRecept = True
'----------��չ�����ӿڱ�������------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AfterBloodPrepareRecept = gcolPlugIn(i).AfterBloodPrepareRecept(lngSys, lngModual, intMode, lngPatientID, lngClinicID, lng����ID)
            Err.Clear
        Next i
    End If
End Function

Public Function AfterBloodVerify(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lng�䷢ID As Long, _
    ByVal strIDs As String, Optional ByVal strTime As String = "") As Boolean
'���ܣ���Ѫ��ɺ����һ��
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      intMode��1-��Ѫ��ˣ�2-ȡ����Ѫ���
'      lng�䷢ID = �䷢ID
'      strIDs���շ�ID�����ԡ������ָ�
'      strTime�����ʱ�䣬���ܲ����루Ϊ�գ�

    AfterBloodVerify = True
'----------��չ�����ӿڱ�������------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AfterBloodVerify = gcolPlugIn(i).AfterBloodVerify(lngSys, lngModual, intMode, strIDs, strTime)
            Err.Clear
        Next i
    End If
End Function

Public Function AfterSendBlood(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lngPatientID As Long, ByVal lng�䷢ID As Long, _
        ByVal strIDs As String) As Boolean
'���ܣ���Ѫ����Ѫ��ɺ����һ��
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      intMode��1-��Ѫ��2-��Ѫ
'      lngPatientID = ����ID
'      lng�䷢ID = �䷢ID
'      strIDs = �շ�ID�����ԡ������ָ�

    AfterSendBlood = True
'----------��չ�����ӿڱ�������------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AfterSendBlood = gcolPlugIn(i).AfterSendBlood(lngSys, lngModual, intMode, lngPatientID, lng�䷢ID, strIDs)
            Err.Clear
        Next i
    End If
End Function

Public Function AfterBloodPurchase(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal strNO As String) As Boolean
'���ܣ�ѪҺ��Ӧ��Ᵽ��ʱ����
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      intMode��1-��Ӧ�����ˣ�2-��Ӧ������
'      strNO-���ݺ�

    AfterBloodPurchase = True
'----------��չ�����ӿڱ�������------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AfterBloodPurchase = gcolPlugIn(i).AfterBloodPurchase(lngSys, lngModual, intMode, strNO)
            Err.Clear
        Next i
    End If
End Function

Public Function AfterBloodReaction(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lng�շ�id As String) As Boolean
'���ܣ���Ѫ��Ӧ��¼��������
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      intMode��1-ҽ���ύ����Ѫ�Ʊ���ʱ��2-ҽ�����˻�����Ѫ��ɾ��ʱ
'      lng�շ�id-��Ѫ��Ӧ��¼.�շ�ID

    AfterBloodReaction = True
'----------��չ�����ӿڱ�������------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AfterBloodReaction = gcolPlugIn(i).AfterBloodReaction(lngSys, lngModual, intMode, lng�շ�id)
            Err.Clear
        Next i
    End If
End Function

Public Function AfterBloodScrap(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal strNO As String) As Boolean
'����:ѪҺ������˻���������(ѪҺ���ϳ���)
'������strBloodIDs=���η���ѪҺ��ѪҺ�շ�ID��������շ�ID֮����","�ָ�磺1001,1002
'          intMode=0 �����ϵǼ�;=1 ���ϵǼ�ɾ��
'˵�����ɸ���strBloodIDs��"ѪҺ�շ�ID�л�ȡ���������Ϣ"
AfterBloodScrap = True
'----------��չ�����ӿڱ�������------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AfterBloodScrap = gcolPlugIn(i).AfterBloodScrap(lngSys, lngModual, intMode, strNO)
            Err.Clear
        Next i
    End If
End Function

Public Function BloodStoreImport(ByRef lng��ⷽʽ As Long, ByRef rsData As ADODB.Recordset)
'���ܣ�ѪҺ����������룬��Ѫ��ϵͳ���е���
'������rsData ���ݼ� �����ڴ˽ӿڽ������ӻ����
'      lng��ⷽʽ '1-�⹺���;3 ���������
'˵�������������÷�ʽ���룬�ڽӿ��н��и�ֵ�����غ��ɵ��÷���ֱ��ʹ����������������¼�� rsData �밴�����ֶ���֯��¼������
'������ʽ��������Ҫ�����ĵĹ���HIS�е�ѪҺ�����ж��չ���(����ѪҺĿ¼�����ṩ����չ���������)������ʱ�������ĵĹ������ȡ��HIS�ж��յĹ��ID��Ȼ���ٸ��ݹ��ID���Ҽ�¼��������ֶ���Ϣ
'�ֶΣ�
'  ID           Number(18),--ѪҺ���ID
' �д�         Number(5), --��ţ���1��ʼ�����ռ�¼������
'  ѪҺ����  Varchar2(200), --Ʒ������
'  ���         Varchar2(100), --�շ���ĿĿ¼�еĹ��
'  ABO         Varchar2(20),
'  RH(D)       Varchar2(1),
'  RH(C)       Varchar2(2),
'  RH (c)      Varchar2(2),
'  RH(E)       Varchar2(2),
'  RH (e)      Varchar2(2),
'  ��λ         Varchar2(20), --�շ���ĿĿ¼�еļ��㵥λ
'  ����ʱ��     Date, --��Ӧ�����д
'  Ч��         Date, --��Ӧ�����д
'  ��ɫ         Varchar2(20),
'  ���         Varchar2(20), --��Ӧ�����д
'  ��©         Number(1), --0 �� 1��
'  ����         Number(16, 5), --������һ�㶼��1
'  �ۼ�         Number(16, 5), -- �շѼ�Ŀ�е��ּ�
'  �ɹ���      Number(16, 5), --ѪҺ����еĳɱ���
'  ��۽��     Number(16, 5),--�ۼ�-�ɹ���
'  Ѫ�����     Varchar2(30), --Ѫ����Ψһ��ʶ
'  ����ABO      Varchar2(20),
'  ����RH(D)    Varchar2(1),
'  ����RH(C)    Varchar2(2),
'  ����RH (c)   Varchar2(2),
'  ����RH(E)    Varchar2(2),
'  ����RH (e)   Varchar2(2),
'  ��Ѫ��       Varchar2(20), --��Ӧ�����д
'  ��Ѫ��       Varchar2(20),
'  ��Ѫʱ��     Date,
'  ��Ѫ��       Varchar2(20), --��Ӧ�����д
'  ��Ѫʱ��     Date, --��Ӧ�����д
'  ��Ѫ�����¶� Number(16, 5), --��Ӧ�����д
'  ������       Varchar2(20)  --��Ӧ�����д
    Dim strFileds As String, strValues As String
    '��¼����ʼ��
    Set rsData = New ADODB.Recordset
    strFileds = "ID," & adLongVarBinary & ",18|�д�," & adLongVarBinary & ",5|���," & adLongVarBinary & ",5|ѪҺ����," & adLongVarChar & ",200|���," & adLongVarChar & ",100|" & _
            "ABO," & adLongVarChar & ",20|RH(D)," & adVarChar & ",1|RH(C)," & adVarChar & ",2|RH (c)," & adVarChar & ",5|RH(E) ," & adVarChar & ",5|" & _
            "RH (e)," & adVarChar & ",5|��λ," & adLongVarChar & ",20|����ʱ��," & adDate & ",20|Ч��," & adDate & ",20|��ɫ," & adLongVarChar & ",20|" & _
            "���," & adLongVarChar & ",20|��©," & adNumeric & ",1|����," & adDouble & ",16|�ۼ�," & adDouble & ",16|�ɹ���," & adDouble & ",16|" & _
            "��۽��," & adDouble & ",16|Ѫ�����," & adLongVarChar & ",30|����ABO," & adLongVarChar & ",20|����RH(D)," & adVarChar & ",1|����RH(C)," & adVarChar & ",2|" & _
            "����RH (c," & adVarChar & ",2|����RH(E)," & adVarChar & ",2|����RH (e)," & adVarChar & ",2|��Ѫ��," & adLongVarChar & ",20|��Ѫ��," & adLongVarChar & ",20|" & _
            "��Ѫʱ��," & adDate & ",20|��Ѫ��," & adLongVarChar & ",20|��Ѫʱ��," & adDate & ",20|��Ѫ�����¶�," & adDouble & ",16|������," & adLongVarChar & ",20"
    Call Record_Init(rsData, strFileds)
    '����Record_Update��Record_Add����������ӻ���²���
    BloodStoreImport = True
    
'----------��չ�����ӿڱ�������------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            BloodStoreImport = gcolPlugIn(i).BloodStoreImport(lng��ⷽʽ, rsData)
            Err.Clear
        Next i
    End If
End Function