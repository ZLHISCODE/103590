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

'���ܣ���չ�ӿڲ��ģ��
'���Ҫ������ʹ�� ��ʵ�� GetUserName ����
Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int���� As Integer = -1)
'���ܣ��ӿڳ�ʼ��
'������cnOracle=�����
'      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      int����  ���ó���:-1
'      �������� lngModual=1070����int���ϲ���
    Set gcnOracle = cnOracle
End Sub

Public Function GetFormCaption(ByVal lngSys As Long, ByVal lngModual As Long) As String
'���ܣ��ڿ�����Ѫ���������ӿ�Ƭ�����ƣ����ŷָ
    Dim strTmp As String, strThis As String
    
    Exit Function '��ʹ�ô˹������޸Ľӿ��ڲ�ʵ��
    
    'strTmp = GetFormCaptionEx(lngSys, lngModual)
    
    Select Case lngModual
    Case m������Ѫ����
        strThis = "Test1,Test2"
    End Select
    GetFormCaption = strThis & IIf(strTmp = "", "", "," & strTmp)
End Function

Public Function GetForm(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strName As String) As Object
'���ܣ���������ֵ���ض��󡣹���վ�������ʱ����
    
    Exit Function '��ʹ�ô˹������޸Ľӿ��ڲ�ʵ��
    
    Select Case lngModual
    Case m������Ѫ����
        Select Case strName
        Case "Test1"
            Set GetForm = New frmDockEx
        Case "Test2"
            Set GetForm = New frmDockEx
        Case Else
            'Set GetForm = GetFormEx(lngSys, lngModual, strName)
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
    Dim strFuncNames As String '���й������ƴ������ŷָ�ɺ�һ�� Auto�ؼ��֣��ӿ���Ӧ�öԴ˱�����ֵ���Լ�����ǰ�ϰ汾��Ʒ
    
    Dim str�˵��� As String '�˵����з��õĹ��ܣ��˵�����Ӧ�ð������й������� str�˵��� = str������ �� str������
    Dim str������ As String
    Dim str������ As String
    Dim str��ݼ� As String '�̶���ʽ  F5--��ݼ�����1;F12--��ݼ�����2,Alt+1 --��ݼ�����3  ��Ӧ�Ĵ� ��"��ݼ�����1|116|0,��ݼ�����2|123|0,��ݼ�����3|49|4"
    
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
End Sub

Public Function GetUserName() As String
'���ܣ���ȡ��λ����
'˵�����û�����ʱ��ʾ������
'��  GetUserName = "������������ҽԺ��Ϣϵͳ"

    GetUserName = ""
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

End Sub

Public Function AfterBloodPrepareRecept(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng����ID As Long) As Boolean
'���ܣ�������Ѫ�����һ��
'������lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      intMode��1-�������룬2-ȡ����������
'      lngPatientID=����ID,lngClinicID=����ID���Һż�¼ID��/��ҳID(סԺ����ʱ)
'      lng����ID=ҽ��id(����ҽ��״̬.ҽ��id)

    AfterBloodPrepareRecept = True
End Function

Public Function AfterBloodVerify(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lngPatientID As Long, ByVal lng�䷢ID As Long, _
    ByVal strIDs As String, Optional ByVal strTime As String = "") As Boolean
'˵���������ģ��
    AfterBloodVerify = True
End Function

Public Function AfterSendBlood(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lngPatientID As Long, ByVal lng�䷢ID As Long, _
        ByVal strIDs As String) As Boolean
'˵���������ģ��
    AfterSendBlood = True
End Function

Public Function AfterBloodPurchase(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal strNO As String) As Boolean
'˵���������ģ��
    AfterBloodPurchase = True
End Function

Public Function AfterBloodReaction(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lng�շ�id As String) As Boolean
'˵���������ģ��
    AfterBloodReaction = True
End Function

Public Function AfterBloodScrap(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal strNO As String) As Boolean
'˵���������ģ��
    AfterBloodScrap = True
End Function

Public Function BloodStoreImport(ByRef lng��ⷽʽ As Long, ByRef rsData As ADODB.Recordset)
'˵���������ģ��
    BloodStoreImport = True
End Function