Attribute VB_Name = "mdl������"
Option Explicit
'�ӿ�������
Public Declare Function CZ_FirstRow Lib "HG_interface.dll" Alias "firstrow" (ByVal pint As Long) As Long
Public Declare Function CZ_NextRow Lib "HG_interface.dll" Alias "nextrow" (ByVal pint As Long) As Long
Public Declare Function CZ_PrevRow Lib "HG_interface.dll" Alias "prevrow" (ByVal pint As Long) As Long
Public Declare Function CZ_LastRow Lib "HG_interface.dll" Alias "lastrow" (ByVal pint As Long) As Long
Public Declare Function CZ_Run Lib "HG_interface.dll" Alias "run" (ByVal pint As Long) As Long
Public Declare Function CZ_NewInterface Lib "HG_interface.dll" Alias "newinterface" () As Long
Public Declare Function CZ_Start Lib "HG_interface.dll" Alias "start" (ByVal pint As Long, ByVal ID As Long) As Long
Public Declare Function CZ_Init Lib "HG_interface.dll" Alias "init" (ByVal pint As Long, ByVal addr As String, ByVal Port As Long, ByVal servlet As String) As Long
Public Declare Function CZ_SetDebug Lib "HG_interface.dll" Alias "setdebug" (ByVal pint As Long, ByVal flag As Integer, ByVal in_direct As String) As Long
Public Declare Function CZ_DataPut Lib "HG_interface.dll" Alias "put" (ByVal pint As Long, ByVal Row As Long, ByVal pname As String, ByVal pvalue As String) As Long
Public Declare Function CZ_GetRowCount Lib "HG_interface.dll" Alias "getrowcount" (ByVal pint As Long) As Long
Public Declare Function CZ_SetRecordset Lib "HG_interface.dll" Alias "setresultset" (ByVal pint As Long, ByVal result_name As String) As Long
Public Declare Function CZ_GetRecordset Lib "HG_interface.dll" Alias "getresultnamebyindex" (ByVal pint As Long, ByVal intIndex As Integer, ByVal result_name As String) As Long
Public Declare Function CZ_GetByName Lib "HG_interface.dll" Alias "getbyname" (ByVal pint As Long, ByVal pname As String, ByVal pvalue As String) As Long
Public Declare Function CZ_GetByIndex Lib "HG_interface.dll" Alias "getbyindex" (ByVal pint As Long, ByVal pindex As Long, ByVal pvalue As String) As Long
Public Declare Function CZ_GetMessage Lib "HG_interface.dll" Alias "getmessage" (ByVal pint As Long, ByVal msg As String) As Long
Public Declare Function CZ_GetException Lib "HG_interface.dll" Alias "getexception" (ByVal pint As Long, ByVal msg As String) As Long
Public Declare Function CZ_SetICCommport Lib "HG_interface.dll" Alias "set_ic_commport" (ByVal pint As Long, ByVal iport As Integer) As Long
Public Declare Sub CZ_DestoryInterface Lib "HG_interface.dll" Alias "destoryinterface" (ByVal pint As Long)

'ȫ�ֱ�����
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����
Enum ҵ�����_������                            '���ڵ��ӿ�ʱ����ҵ�����
    ��ͨ���� = 11
    ��ͨסԺ = 12
    ��ͥ���� = 14
    ����涨�� = 13
    
    ���Ｑ�� = 15
    �����ؼ� = 16
    �������� = 21
    ����סԺ = 22
    �������� = 41
    ����סԺ = 42
End Enum
Enum ��¼��_������
    MoveFirst = 1
    MoveLast = 2
    MovePrev = 3
    MoveNext = 4
End Enum
Enum Debug_������
    Normal = 0                                  '����ģʽ
    Record = 1                                  '����ģʽ
End Enum
Enum Function_������
    ��¼���� = 0
    '--��Ŀƥ��(1001)
    ��Ŀƥ��_ȡ��Ŀ��Ϣ = 100102
    ��Ŀƥ��_ȡƥ����Ŀ��Ϣ = 100103
    ��Ŀƥ��_ɾ��ƥ����Ϣ = 100104
    ��Ŀƥ��_����ƥ����Ϣ = 100105
    ��Ŀƥ��_��Ŀƥ�� = 100106
    'Modified By ���� ��������ɳ ԭ�����ӹ���
    ��Ŀƥ��_ȡ������Ŀƥ����Ϣ = 120507
    '--��ͨ����ҵ�����¼��(���ķ�) (1101)
    ��ͨ����_�����֤ = 110101                  'Ҫ���ظ����ʻ����
    ��ͨ����_�շ� = 110104                      '��������ٱ���
    ����_����ϸ = 110111
    '--��ͨסԺ��Ժ�Ǽ�(1201)
    ��ͨסԺ_�����֤ = 120101                  '�����ظ����ʻ�����Ԥ����ʱ����
    ��ͨסԺ_��Ժ�Ǽ� = 120104
    '--��ͨסԺȡ���Ǽ�(1211)
    ��ͨסԺ_ȡ����Ժ = 121104
    '--סԺ����(1205)
    סԺ����_�ϴ���ϸ = 120502                  '��Ӧ�ӿڵı�����ϸ
    סԺ����_����ϸ = 120503
    '--סԺ����(1202)
    סԺ����_Ԥ���� = 120206
    סԺ����_��ʽ���� = 120214
    סԺ����_������� = 120218                  '����סԺ�ڼ����ȫ������
    ��Ժ�Ǽ� = 120204
    '--ȡ����Ժ(1212)
    ȡ����Ժ = 121202
    '--סԺ��Ϣ�޸�
    סԺ��Ϣ_�޸� = 120302
    '--����涨��ҵ�����¼��(���ķ�) (1305)
    ����涨��_�����֤ = 130501
    ����涨��_�շ� = 130504
    '--��������
    ����_���� = 200900
    ����_�޸����� = 200910
    ����_������Ϣ = 200001
    ����_������� = 200004                      '������Ļ����ţ����ػ����������ʻ���ҽ������...��
    ����_������У�� = 200009
    ����_���ŵ��ݳ��� = 112010
    ����_��ȡ��Ʊ��Ϣ = 200040
    '--���㵥
    ���㵥_סԺ = 200030
    ���㵥_���� = 200031
    ���㵥_����涨�� = 200032
    '--������ܱ����ʵ���
    ������ܱ�_סԺ = 200035
    ������ܱ�_���� = 200037
    ������ܱ�_����涨�� = 200038
    '--תԺ����
    תԺ����_������Ϣ = 121002
    תԺ����_У����Ϣ = 121003
    תԺ����_����תԺ���� = 121005
    תԺ����_��ѯ�����Ϣ = 121006
    ��ȡҽԺ��Ϣ = 121007
End Enum

Private Type ComInfo_������
    ҽԺ���� As String
    ����Ա���� As String
    ҵ������ As String
    ���˱�� As String
    ҵ�����к� As String
    �ʻ���� As Currency
    �ܷ��� As Currency
    �������� As String                      '���������֤�󷵻صļ�������
End Type
Public gCominfo_������ As ComInfo_������

Private Const סԺ�����ۼ� = 0
Private Const ��Ժ���ұ�� = 1
Private Const ��Ժ�������� = 2
Private Const ��Ժ������� = 3
Private Const ��Ժ�������� = 4
Private Const ��Ժ������� = 5
Private Const ��λ���� = 6
Private Const סԺ�� = 7
Private Const ��Ժ��� = 8
Private Const ��Ժ��� = 9

Private Const mintTest As Integer = 0           '���㣬����Ԥ���㣨סԺԤ����������Ƿֿ��ģ�����ʹ�ã�
Private Const mintICCard As Integer = 1         'ʹ��IC��
Private Const strDebug_Path As String = "C:\Log" '���������Ϣ���ļ���,ҽ����ʼ���Ĵ���

Public glngInterface_������ As Long             '���Ӵ�
Public glngReturn_������ As Long                '�ӿڷ���ֵ
Public gstrErrInfo_������ As String             '������Ϣ��������
Public gstrField_������ As String
Public gstrValue_������ As String
Private mintICPort As Integer                   'IC�豸�Ķ˿ں�
Public mint���õ���_���� As Integer             '1-��������;2-��������
Private mstrAddress As String, mstrPort As String, mstrServlet As String

Public Sub ErrInformation()
    If glngReturn_������ < 0 Then
        glngReturn_������ = CZ_GetMessage(glngInterface_������, gstrErrInfo_������)
        MsgBox gstrErrInfo_������, vbInformation, gstrSysName
    End If
End Sub

Public Function ҽ����ʼ��_������() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim objFileSys As FileSystemObject
    On Error GoTo errHand
    
    If glngInterface_������ = 0 Then
        '����ҽ���ӿڵĳ�ʼ������
        gstrErrInfo_������ = Space(1000)
        If Not GetServerInfo Then Exit Function
        glngInterface_������ = CZ_NewInterface()
        glngReturn_������ = CZ_Init(glngInterface_������, mstrAddress, mstrPort, mstrServlet)
        Call ErrInformation
        If glngReturn_������ = -1 Then
            '����ʧ��
            Call ҽ����ֹ_������
            Exit Function
        End If
        '����IC�豸�Ķ˿ں�
        Call CZ_SetICCommport(glngInterface_������, mintICPort)
        '���õ���Ŀ¼
        Set objFileSys = New FileSystemObject
        If Not objFileSys.FolderExists(strDebug_Path) Then
            objFileSys.CreateFolder (strDebug_Path)
        End If
        Call CZ_SetDebug(glngInterface_������, Debug_������.Record, strDebug_Path)
        '��¼����(ʧ����Ͽ����Ӳ��˳�)
        If Not frm��¼����.LoginCenter(TYPE_������) Then
            Call ҽ����ֹ_������
            Exit Function
        End If
        gCominfo_������.����Ա���� = Right(UserInfo.���, 5)
        'ȡҽԺ����
        gstrSQL = "Select ҽԺ���� From ������� Where ���=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽԺ����", TYPE_������)
        gCominfo_������.ҽԺ���� = Nvl(rsTemp!ҽԺ����)
        
        'ȡ���õ���
        mint���õ���_���� = 0
        gstrSQL = "Select ����ֵ From ���ղ��� Where ������='���õ���' And ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���õ���", TYPE_������)
        If Not rsTemp.EOF Then
            mint���õ���_���� = Nvl(rsTemp!����ֵ, 0)
        End If
    End If
    
    ҽ����ʼ��_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ҽ����ֹ_������() As Boolean
    If glngInterface_������ = 0 Then
        ҽ����ֹ_������ = True
        Exit Function
    End If
    
    Call CZ_DestoryInterface(glngInterface_������)
    glngInterface_������ = 0
    
    ҽ����ֹ_������ = True
End Function

Public Function ҽ������_����() As Boolean
    ҽ������_���� = frmSet����.ShowME
End Function

Private Function GetServerInfo() As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '��ȡIC�豸�Ķ˿ں�
    mintICPort = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "IC�豸�˿�", 1)
    
    '��ȡ��������ַ���˿ڼ��������('��������ַ','�������˿ں�','��������ڳ���')
    gstrSQL = " Select ������,����ֵ From ���ղ���" & _
              " Where ����=[1] And ������ Like '������%'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������ַ���˿ڼ��������", TYPE_������)
    
    With rsTemp
        Do While Not .EOF
            Select Case !������
            Case "��������ַ"
                mstrAddress = Nvl(!����ֵ)
            Case "�������˿ں�"
                mstrPort = Nvl(!����ֵ)
            Case "��������ڳ���"
                mstrServlet = Nvl(!����ֵ)
            End Select
            .MoveNext
        Loop
    End With
    
    GetServerInfo = Not (mstrAddress = "" Or mstrPort = "" Or mstrServlet = "")
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ���ýӿ�_׼��_������(ByVal lng���� As Long) As Boolean
    On Error GoTo errHand
    
    glngReturn_������ = CZ_Start(glngInterface_������, lng����)
    Call ErrInformation
    
    'Modified By ���� ��ɳ    <=0
    If glngReturn_������ < 0 Then Exit Function
    
    ���ýӿ�_׼��_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ���ýӿ�_ִ��_������() As Boolean
    On Error GoTo errHand
    
    glngReturn_������ = CZ_Run(glngInterface_������)
    Call ErrInformation
    'Modified By ���� ��ɳ    <=0
    If glngReturn_������ < 0 Then Exit Function
    
    ���ýӿ�_ִ��_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ���ýӿ�_ָ����¼��_������(ByVal strRecordName As String) As Boolean
    On Error GoTo errHand
    
    glngReturn_������ = CZ_SetRecordset(glngInterface_������, strRecordName)
    Call ErrInformation
    'Modified By ���� ��ɳ    <=0
    If glngReturn_������ < 0 Then Exit Function
    
    ���ýӿ�_ָ����¼��_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ���ýӿ�_��¼��_������() As Boolean
    Dim lngRecord As Long
    On Error GoTo errHand
    
    lngRecord = CZ_GetRowCount(glngInterface_������)
    ���ýӿ�_��¼��_������ = (lngRecord > 0)
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ���ýӿ�_�ƶ���¼��_������(ByVal intType As ��¼��_������) As Boolean
    On Error Resume Next
    
    Err = 0
    Select Case intType
    Case ��¼��_������.MoveFirst
        glngReturn_������ = CZ_FirstRow(glngInterface_������)
    Case ��¼��_������.MovePrev
        glngReturn_������ = CZ_PrevRow(glngInterface_������)
    Case ��¼��_������.MoveNext
        glngReturn_������ = CZ_NextRow(glngInterface_������)
    Case ��¼��_������.MoveLast
        glngReturn_������ = CZ_LastRow(glngInterface_������)
    End Select
    
    If Err <> 0 Then Exit Function
    ���ýӿ�_�ƶ���¼��_������ = Not (glngReturn_������ < 0)
    Exit Function
End Function

Public Function ���ýӿ�_��ȡ����_������(ByVal strField As String, strValue As String) As Boolean
    On Error GoTo errHand
    
    strValue = Space(1000)
    Call DebugTool("ȡ�ֶΣ�" & strField)
    glngReturn_������ = CZ_GetByName(glngInterface_������, strField, strValue)
    Call ErrInformation
    If glngReturn_������ <= 0 Then Exit Function
    'Modified By ���� ��������ɳ ԭ�򣺼���ȥ��0�ַ���Replace���
    strValue = Trim(Replace(strValue, Chr(0), ""))
    
    ���ýӿ�_��ȡ����_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ���ýӿ�_д��ڲ���_������(ByVal lngRow As Long) As Boolean
    Dim intField As Integer, intCOUNT As Integer
    Dim arrField, arrData
    Dim blnErr As Boolean
    On Error GoTo errHand
    
    arrField = Split(gstrField_������, "||")
    'Modified By ���� ��������ɳ ԭ��arrDataд����arrField
    arrData = Split(gstrValue_������, "||")
    intCOUNT = UBound(arrField)
    For intField = 0 To intCOUNT
        glngReturn_������ = CZ_DataPut(glngInterface_������, lngRow, arrField(intField), arrData(intField))
        If Not blnErr Then
            blnErr = (glngReturn_������ <= 0)
        End If
    Next
    
    'Modified By ���� ��������ɳ ԭ��Exit FunctionӦ�÷������
    ���ýӿ�_д��ڲ���_������ = Not blnErr
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �ҺŽ���_������(ByVal lng����ID As Long) As Boolean
    Dim str���㷽ʽ As String
    Dim arr���㷽ʽ, intStart As Integer, intTotal As Integer, lng������Ŀ As Long
    Dim cur�����ʻ� As Currency, cur�����ܶ� As Currency '�ϴ������ܶ�
    Dim rsTemp As New ADODB.Recordset
    '�ȵ�����Ԥ����,��Ҫ��ȡ�����ʻ�֧����,�ٵ��������
    '���ղ����б�����������ʻ�֧����������ĿID���ҺŽ���ʱ�жϣ����δ���ã�����ȫ�Ը��������ϴ���������ϴ���һ����ϸ
    
    On Error GoTo errHand
    
    '����ǳ������������ڹҺ�ȫ���ֽ���㣬����Ҫ���ϴ��Һ���ϸ
    If mint���õ���_���� = 1 Then '��������
        �ҺŽ���_������ = True
        Exit Function
    End If
    
    '����ȡ���ղ���
    Call DebugTool("��ȡ������Ŀ")
    gstrSQL = "Select ����ֵ Value From ���ղ��� Where ����=[1] And ������='�����ʻ�֧��(�Һ�)'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ɸ����ʻ�֧����������ĿID", TYPE_������)
    If Not rsTemp.EOF Then
        lng������Ŀ = Nvl(rsTemp!Value, 0)
        Call DebugTool("������ĿID:" & lng������Ŀ)
    End If
    
    '��ȡ���ν���ķ�����ϸ,��������������ȡԤ����
    gstrSQL = "Select Rownum ��ʶ��,A.ID,A.����ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.�Ǽ�ʱ��,A.������," & _
            "   A.�շ�ϸĿID,A.������ĿID,A.����*A.���� as ����,A.���㵥λ,Round(A.���ʽ��/(A.����*A.����),4) as ����,A.ʵ�ս��," & _
            "   A.�շ����,B.���� as ��Ŀ����,B.���� as ��Ŀ����,B.���,D.��Ŀ���� ҽ������," & _
            "   C.���� ��������,E.���� �ܵ�����" & _
            " From (Select * From ������ü�¼ Where ����ID=" & lng����ID & ") A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D,���ű� E " & _
            " Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID(+) And A.ִ�в���ID=E.ID(+) And A.�շ�ϸĿID=D.�շ�ϸĿID And D.����=[1]" & _
            " Order by A.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ν��ʷ�����ϸ", TYPE_������)
    With rsTemp
        Do While Not .EOF
            cur�����ܶ� = cur�����ܶ� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    If Not �����������_������(rsTemp, str���㷽ʽ, True, lng������Ŀ) Then Exit Function
    
    '�ֽ���㷽ʽ,��ȡ�����ʻ�֧�������֧���"������ʽ;���;�Ƿ������޸�|...."��
    arr���㷽ʽ = Split(str���㷽ʽ, "|")
    intTotal = UBound(arr���㷽ʽ)
    For intStart = 0 To intTotal
        If Split(arr���㷽ʽ(intStart), ";")(0) = "�����ʻ�" Then
            cur�����ʻ� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
            Exit For
        End If
    Next
    
    Call DebugTool("�ϴ������ܶ�:" & cur�����ܶ� & "|���ظ����ʻ�֧����:" & cur�����ʻ�)
    '�ٵ���������㣨strSelfNO��Ϊû��ʹ�ã����Դ��գ�
    If Not �������_������(lng����ID, cur�����ʻ�, "", True, lng������Ŀ) Then Exit Function
    
    '�޸Ĳ���Ԥ����¼�������ֽ���㷽ʽ������Ӧ�ļ�¼��
    If lng������Ŀ <> 0 Then
        '����Ԥ����¼
        For intStart = 0 To intTotal
            If Split(arr���㷽ʽ(intStart), ";")(0) <> "�ֽ�" Then
                gstrSQL = " insert into ����Ԥ����¼(ID,��¼����,NO,��¼״̬,����ID,��ҳID,����ID,�ɿλ," & _
                         " ��λ������,��λ�ʺ�,ժҪ,���,���㷽ʽ,�������,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID) " & _
                         " select ����Ԥ����¼_ID.nextval ID,��¼����,NO,��¼״̬,����ID,��ҳID,����ID, " & _
                         " �ɿλ,��λ������,��λ�ʺ�,ժҪ,���,'" & Split(arr���㷽ʽ(intStart), ";")(0) & "',�������,�տ�ʱ��,����Ա���, " & _
                         " ����Ա����," & Val(Split(arr���㷽ʽ(intStart), ";")(1)) & ",����ID " & _
                         " from ����Ԥ����¼" & _
                         " Where ����ID=" & lng����ID & " And ���㷽ʽ='�ֽ�' And not (��¼����=1 or ��¼����=11)"
                         cur�����ܶ� = cur�����ܶ� - Val(Split(arr���㷽ʽ(intStart), ";")(1))
                gcnOracle.Execute gstrSQL
            End If
        Next
        
        '�����ֽ�֧����
        If cur�����ܶ� <> 0 Then
            '�޸��ֽ�֧����
            gstrSQL = " Update ����Ԥ����¼ Set ��Ԥ��= " & cur�����ܶ� & _
                      " Where ����ID=" & lng����ID & " And ���㷽ʽ='�ֽ�' And not (��¼����=1 or ��¼����=11)"
        Else
            '���ֽ�֧���ɾ����Ԥ����¼
            gstrSQL = " Delete ����Ԥ����¼ " & _
                      " Where ����ID=" & lng����ID & " And ���㷽ʽ='�ֽ�' And not (��¼����=1 or ��¼����=11)"
        End If
        gcnOracle.Execute gstrSQL
    End If
    
    Call frm������Ϣ.ShowME(lng����ID)
    �ҺŽ���_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �Һų���_������(ByVal lng����ID As Long) As Boolean
    Dim cur�����ʻ� As Currency, lng����ID As Long, lng��¼ID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    'ȡ��������¼�Ľ���ID������ID
    gstrSQL = "select distinct A.����ID,A.����ID from ������ü�¼ A,������ü�¼ B where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���²����Ľ���ID", lng����ID)
    lng��¼ID = rsTemp!����ID
    lng����ID = rsTemp!����ID
    
    '��ȡԭʼ��¼�ĸ����ʻ�֧�����Ϊ����������ʱδʹ�ò����������ʻ��������Բ�ȡ��
    gstrSQL = "Select Nvl(A.��Ԥ��,0) �����ʻ� " & _
        " From ����Ԥ����¼ A,�����ʻ� B " & _
        " Where A.����ID=B.����ID And B.����=[2]" & _
        " And A.���㷽ʽ in ('�����ʻ�') And A.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡԭʼ���ݵĸ����ʻ�֧����", lng��¼ID, TYPE_������)
    cur�����ʻ� = 0
    If Not rsTemp.EOF Then
        cur�����ʻ� = rsTemp!�����ʻ�
    End If
    
    If Not ����������_������(lng����ID, cur�����ʻ�, lng����ID) Then Exit Function
    �Һų���_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �����������_������(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String, _
Optional ByVal bln�Һ� As Boolean = False, Optional ByVal lng������ĿID As Long = 0) As Boolean
    Dim lng���� As Long, str���˱�� As String
    Dim str�������� As String, str������ As String, str֧����� As String
    Dim str����ʱ�� As String, str�������� As String
    Dim strҽ����� As String, strҽ������ As String
    Dim str��� As String, str���� As String, str���� As String, str��Ŀ���� As String
    Dim str���Һ� As String, str�������� As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsPhysic As New ADODB.Recordset
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�������,����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    '�ҺŽ���Ҳʹ�ñ����̣�������ֽ�֧������Ŀ����Ҫ���ض���ҽԺ�����ҽ������ 110100001 �Һŷ�
    On Error GoTo errHand
    
    '--��IC��
    '����IC���е���Ϣ
    If Not ���ýӿ�_׼��_������(Function_������.����_����) Then Exit Function
    If Not ���ýӿ�_ִ��_������ Then Exit Function
    'ȡ���صļ�¼��
    'If Not ���ýӿ�_ָ����¼��_������("ICInfo") Then Exit Sub
    If Not ���ýӿ�_��ȡ����_������("indi_id", str���˱��) Then Exit Function
    
    'У����˱���Ƿ���ȷ
    gstrSQL = "Select ���� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "У��ò����Ƿ�ʹ���Լ��Ŀ����н���", TYPE_������, CLng(rs��ϸ!����ID))
    If rsTemp!���� <> str���˱�� Then
        MsgBox "�ò��˲���ʹ���Լ���ҽ���������������ֹ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str�������� = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    lng���� = IIf(gCominfo_������.ҵ������ = ҵ�����_������.����涨��, _
                    Function_������.����涨��_�շ�, Function_������.��ͨ����_�շ�)
    If Not ���ýӿ�_׼��_������(lng����) Then Exit Function
    
    'ȡҽ������
    Call DebugTool("ȡҽ������")
    strҽ����� = "": strҽ������ = ""
    If bln�Һ� Then
        Call DebugTool("�Һ�-ȡҽ������")
        gstrSQL = "select ҽ������ from �ҺŰ��� where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����", CStr(rs��ϸ!���㵥λ))
        strҽ������ = Nvl(rsTemp!ҽ������)
    Else
        Call DebugTool("����-ȡҽ������")
        strҽ������ = Nvl(rs��ϸ!������)
    End If
    
    Call DebugTool("ȡҽ�����")
    If Trim(strҽ������) <> "" Then
        gstrSQL = "Select ��� From ��Ա�� Where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����", strҽ������)
        strҽ����� = rsTemp!���
    End If
    
    If strҽ������ <> "" And mint���õ���_���� = 2 Then
        Call DebugTool("ȡ���ұ��������")
        gstrSQL = " Select ����,���� From ���ű� " & _
                  " Where ID in " & _
                  "     (Select ����ID From ������Ա " & _
                  "     Where ��ԱID=" & _
                  "         (Select ID From ��Ա�� Where ����=[1]))"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���ұ��������", strҽ������)
        If rsTemp.RecordCount <> 0 Then
            str���Һ� = rsTemp!����
            str�������� = rsTemp!����
        End If
    End If
    
    Call DebugTool("��ʼ�ϴ���ϸ")
    With rs��ϸ
        'д��ڲ���
'        1   hospital_id    ҽ�ƻ�������   20   ��
'        2   indi_id        ���˱��        8   ��
'        3   busi_type      ҵ������        2   ��  "11"������
'        4   ic_flag        �ÿ���־        1   ��  "0"����ʹ��IC����"1"��ʹ��IC��
'        5   reg_staff      �Ǽ���Ա����    5   ��
'        6   reg_man        �Ǽ�������      10  ��
'        7   begin_date     ����ʱ��            ��  ��ʽ��YYYY-MM-DD HH:MI:SS(24Сʱ)
'        8   in_disease     �Ǽ����        20  ��  ��������
'        9   calcSaveFlag   ���㱣���־    1   ��  "0"�����㣻"1"���շ�
'        10  accMoney       �����ʻ�֧�����18  ��
'        11  recipe_no      ������          20  ��
'--------2004-01-12����--------
'        12  doctor_no      ����ҽ�����    12
'        13  doctor_name    ����ҽ������    10
'------------------------------
'        14  note           ��ע            100 ��
'������������������������Ҫ------------------------------
'        15  in_dept        ���Һ�          10
'        16  in_dept_name   ��������        20
        gstrField_������ = "hospital_id||indi_id||busi_type||ic_flag||reg_staff||" & _
                    "reg_man||begin_date||in_disease||calcSaveFlag||accMoney||recipe_no||doctor_no||doctor_name||note"
        If mint���õ���_���� = 2 Then gstrField_������ = gstrField_������ & "||in_dept||in_dept_name"
        gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.���˱�� & "||" & _
                    gCominfo_������.ҵ������ & "||" & mintICCard & "||" & gCominfo_������.����Ա���� & "||" & _
                    gstrUserName & "||" & str����ʱ�� & "||" & _
                    gCominfo_������.�������� & "||" & mintTest & "||0||||" & strҽ����� & "||" & strҽ������ & "||"
        If mint���õ���_���� = 2 Then gstrValue_������ = gstrValue_������ & "||" & str���Һ� & "||" & str��������
        If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
        
        '���ö�Ӧ�ļ�¼����׼������ϸ
        If Not ���ýӿ�_ָ����¼��_������("FeeInfo") Then Exit Function
        'д������ϸ
        gCominfo_������.�ܷ��� = 0
        Do While Not .EOF
'            1   medi_item_type ��ĿҩƷ����        1   ��  "0"��������Ŀ��"1"����ҩ��"2"���г�ҩ��"3"���в�ҩ
'            2   his_item_code  ҽԺҩƷ��Ŀ����    20  ��
'            3   his_item_name  ҽԺҩƷ��Ŀ����    50  ��
'            4   model          ����                30  ��
'            5   factory        ����                50  ��
'            6   standard       ���                30  ��
'            7   fee_date       ���÷���ʱ��            ��  ��ʽ��YYYY-MM-DD
'            8   unit           ������λ            10  ��
'            9   price          ����                12  ��
'            10  dosage         ����                12  ��
'            11  money          ���                12  ��
'            12  opp_serial_fee ��Ӧ�������к�      12  ��
            
            '�����ҩƷ����ȥȡ����
            str���� = ""
            str��Ŀ���� = ""
            If !�շ���� = 5 Or !�շ���� = 6 Or !�շ���� = 7 Then
                gstrSQL = " Select C.��ʶ��,C.����,B.���� ���� From ҩƷ��Ϣ A,ҩƷ���� B,ҩƷĿ¼ C " & _
                          " Where A.����=B.���� And A.ҩ��ID=C.ҩ��ID And C.ҩƷID =[1]"
                Set rsPhysic = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����", CLng(!�շ�ϸĿID))
                str���� = Nvl(rsPhysic!����)
                str��Ŀ���� = Nvl(rsPhysic!��ʶ��)
                If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsPhysic!����)
            End If
            
            'ȡ���Ͳ���
            gstrSQL = "Select ����,����,���,��ʶ���� From �շ�ϸĿ Where ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ�ϸĿ�����Ϣ", CLng(!�շ�ϸĿID))
            str���� = ""
            str��� = Nvl(rsTemp!���)
            If InStr(1, str���, "��") <> 0 Then
                str���� = ToVarchar(Trim(Split(str���, "��")(1)), 50)
                str��� = ToVarchar(Trim(Split(str���, "��")(0)), 30)
            Else
                str��� = ToVarchar(Trim(str���), 30)
            End If
            
            '����ǳ�ҩƷ����Ŀ��ȡ�������ȡ��ʶ����
            If Not (!�շ���� = 5 Or !�շ���� = 6 Or !�շ���� = 7) Then
                str��Ŀ���� = Nvl(rsTemp!��ʶ����)
                If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsTemp!����)
            End If
            
            '����ǹҺţ���������ĿID<>����������ĿID���򴫹̶���ҽԺ�����ҽ������
            If bln�Һ� Then
                If rs��ϸ!������Ŀid <> lng������ĿID Then
                    str��Ŀ���� = "110100001"
                End If
            End If
            
            gstrField_������ = "medi_item_type||his_item_code||his_item_name||model||factory||" & _
                        "standard||fee_date||unit||price||dosage||money||opp_serial_fee||hos_serial"
            gstrValue_������ = IIf(!�շ���� = 5, "1", IIf(!�շ���� = 6, "2", IIf(!�շ���� = 7, "3", "0"))) & "||" & _
                        str��Ŀ���� & "||" & Nvl(rsTemp!����) & "||" & str���� & "||" & str���� & "||" & str��� & "||" & _
                        str�������� & "||" & Nvl(!���㵥λ) & "||" & !���� & "||" & !���� & "||" & !ʵ�ս�� & "||||"
            
            If .AbsolutePosition <> 1 Then
                If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Function
            End If
            If Not ���ýӿ�_д��ڲ���_������(.AbsolutePosition) Then Exit Function
            
            gCominfo_������.�ܷ��� = gCominfo_������.�ܷ��� + !ʵ�ս��
            .MoveNext
        Loop
    End With
    
    If Not ���ýӿ�_ִ��_������() Then Exit Function
    
    '��ȡ�����ĸ������֧����
    If Not ���ýӿ�_ָ����¼��_������("BizInfo") Then Exit Function
'    1   fund_id;    �������    3
'    2   fund_name   ��������    30
'    3   real_pay    ֧�����    12
'    4   serial_no   ҵ�����к�  12
    If ���ýӿ�_��¼��_������ Then
        Do While True
            Call ���ýӿ�_��ȡ����_������("fund_id", str������)
            Call ���ýӿ�_��ȡ����_������("fund_name", str��������)
            Call ���ýӿ�_��ȡ����_������("real_pay", str֧�����)
            If str������ <> 999 And Val(str֧�����) <> 0 Then
                str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ = "", "", "|") & _
                            IIf(str������ = "003", "�����ʻ�", str��������) & ";" & str֧����� & ";0"
            End If
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
    End If
    
    If str���㷽ʽ = "" Then str���㷽ʽ = "�����ʻ�;0;0"
    �����������_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �������_������(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String, _
Optional ByVal bln�Һ� As Boolean = False, Optional ByVal lng������ĿID As Long = 0) As Boolean
    Dim lng���� As Long, lng����ID As Long
    Dim str�������� As String, str������ As String, str֧����� As String
    Dim strҵ�����к� As String, str����ʱ�� As String, str�������� As String
    Dim curͳ����� As Currency, cur�ֽ� As Currency
    Dim strҽ����� As String, strҽ������ As String, str��Ŀ���� As String
    Dim str��� As String, str���� As String, str���� As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsPhysic As New ADODB.Recordset
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    'ע�⣺�ӿڹ涨��������ϸ�������ϴ���סԺ��ϸ��Ԥ����ʱ�ϴ���������ڽ��㣬����ʹ��Ȧ��ӿڣ����������Ǯ���������ڣ������ӿ��ڽ��
    '���������Ҫͨ��������������ȡ����Ȧ�����ǽӿڷ��أ���Ҫ�޸�
    On Error GoTo errHand
    str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str�������� = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    lng���� = IIf(gCominfo_������.ҵ������ = ҵ�����_������.����涨��, _
                    Function_������.����涨��_�շ�, Function_������.��ͨ����_�շ�)
    If Not ���ýӿ�_׼��_������(lng����) Then Exit Function
    
    'ȡҽ�����������
    Call DebugTool("ȡҽ�����������")
    strҽ����� = "": strҽ������ = ""
    
    If bln�Һ� Then
        gstrSQL = "select ҽ������ ������ from �ҺŰ��� where ����=(Select ���㵥λ From ������ü�¼ Where ����ID=[1] And Rownum<2)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����", lng����ID)
    Else
        gstrSQL = "Select ������ From ������ü�¼ Where ����ID=[1] And Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҽ��", lng����ID)
    End If
    strҽ������ = Nvl(rsTemp!������)

    If Trim(strҽ������) <> "" Then
        gstrSQL = "Select ���,���� From ��Ա�� Where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����", strҽ������)
        strҽ����� = rsTemp!���
    End If
    
    '�ϴ�������ϸ��¼
    Call DebugTool("�ϴ�������ϸ��¼")
    gstrSQL = "Select Rownum ��ʶ��,A.ID,A.����ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.�Ǽ�ʱ��,A.������ as ҽ��," & _
            "   A.�շ�ϸĿID,A.������ĿID,A.���㵥λ,A.����*A.���� as ����,Round(A.���ʽ��/(A.����*A.����),4) as ����,A.ʵ�ս�� ���," & _
            "   A.�շ����,B.��ʶ����,B.���� as ����,B.���� as ��Ŀ����,B.���,D.��Ŀ���� ҽ������," & _
            "   C.���� ��������,E.���� �ܵ�����" & _
            " From (Select * From ������ü�¼ Where ����ID=[2] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0) A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D,���ű� E " & _
            " Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID(+) And A.ִ�в���ID=E.ID(+) And A.�շ�ϸĿID=D.�շ�ϸĿID And D.����=[1]" & _
            " Order by A.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ν��ʷ�����ϸ", TYPE_������, lng����ID)
    lng����ID = rsTemp!����ID
    With rsTemp
        'д��ڲ���
        Call DebugTool("�ϴ�������ϸ��¼-����ͷ")
        gstrField_������ = "hospital_id||indi_id||busi_type||ic_flag||reg_staff||" & _
                    "reg_man||begin_date||in_disease||calcSaveFlag||accMoney||recipe_no||doctor_no||doctor_name||note"
        gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.���˱�� & "||" & _
                    gCominfo_������.ҵ������ & "||" & mintICCard & "||" & gCominfo_������.����Ա���� & "||" & _
                    gstrUserName & "||" & str����ʱ�� & "||" & _
                    gCominfo_������.�������� & "||1||" & cur�����ʻ� & "||" & !NO & "||" & strҽ����� & "||" & strҽ������ & "||"
        If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
        
        '���ö�Ӧ�ļ�¼����׼������ϸ
        If Not ���ýӿ�_ָ����¼��_������("FeeInfo") Then Exit Function
        
        Call DebugTool("�ϴ�������ϸ��¼-������")
        Do While Not .EOF
            
            '�����ҩƷ����ȥȡ����
            Call DebugTool("�ϴ�������ϸ��¼-������-ȡҩƷ���ͣ�ҽ������")
            str���� = ""
            str��Ŀ���� = ""
            If !�շ���� = 5 Or !�շ���� = 6 Or !�շ���� = 7 Then
                gstrSQL = " Select B.���� ����,C.����,C.��ʶ�� From ҩƷ��Ϣ A,ҩƷ���� B,ҩƷĿ¼ C " & _
                          " Where A.����=B.���� And A.ҩ��ID=C.ҩ��ID And C.ҩƷID =[1]"
                Set rsPhysic = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����", CLng(!�շ�ϸĿID))
                str���� = Nvl(rsPhysic!����)
                str��Ŀ���� = Nvl(rsPhysic!��ʶ��)
                If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsPhysic!����)
            Else
                str��Ŀ���� = Nvl(!��ʶ����)
                If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(!����)
            End If
            
            Call DebugTool("�ϴ�������ϸ��¼-������-ȡ��񡢲���")
            str���� = ""
            str��� = Nvl(!���)
            If InStr(1, str���, "��") <> 0 Then
                str���� = ToVarchar(Trim(Split(str���, "��")(1)), 50)
                str��� = ToVarchar(Trim(Split(str���, "��")(0)), 30)
            Else
                str��� = ToVarchar(Trim(str���), 30)
            End If
            
            '����ǹҺţ���������ĿID<>����������ĿID���򴫹̶���ҽԺ�����ҽ������
            Call DebugTool("�ϴ�������ϸ��¼-������-����ǹҺŻ�Ҫ����ȡ�ض�����Ŀ����")
            If bln�Һ� Then
                If rsTemp!������Ŀid <> lng������ĿID Then
                    str��Ŀ���� = "110100001"
                End If
            End If
            
            gstrField_������ = "medi_item_type||his_item_code||his_item_name||model||factory||" & _
                        "standard||fee_date||unit||price||dosage||money||opp_serial_fee||hos_serial"
            gstrValue_������ = IIf(!�շ���� = 5, "1", IIf(!�շ���� = 6, "2", IIf(!�շ���� = 7, "3", "0"))) & "||" & _
                        str��Ŀ���� & "||" & Nvl(!��Ŀ����) & "||" & str���� & "||" & str���� & "||" & str��� & "||" & _
                        str�������� & "||" & Nvl(!���㵥λ) & "||" & !���� & "||" & !���� & "||" & !��� & "||||" & !ID
            
            If .AbsolutePosition <> 1 Then
                Call DebugTool("�ϴ�������ϸ��¼-������-MOVENEXT")
                If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Function
            End If
            Call DebugTool("�ϴ�������ϸ��¼-������-д��ڲ���")
            If Not ���ýӿ�_д��ڲ���_������(.AbsolutePosition) Then Exit Function
            '�����ϴ���־����Ϊ��ϸ������ȷ�ϴ��󣬲��ܱ�֤�������ȷ��
            'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
            Call DebugTool("�ϴ�������ϸ��¼-������-���ϴ���־")
            gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsTemp("NO") & "'," & rsTemp("���") & "," & rsTemp("��¼����") & "," & rsTemp("��¼״̬") & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
            
            .MoveNext
        Loop
    End With
    
    If Not ���ýӿ�_ִ��_������() Then Exit Function
    
    '��ȡ�����ĸ������֧����
    Call DebugTool("ȡ������")
    If Not ���ýӿ�_ָ����¼��_������("BizInfo") Then Exit Function
'    1   fund_id;    �������    3
'    2   fund_name   ��������    30
'    3   real_pay    ֧�����    12
'    4   serial_no   ҵ�����к�  12
    
    If ���ýӿ�_��¼��_������ Then
        curͳ����� = 0
        cur�ֽ� = 0
        
        Do While True
            Call ���ýӿ�_��ȡ����_������("fund_id", str������)
            Call ���ýӿ�_��ȡ����_������("fund_name", str��������)
            Call ���ýӿ�_��ȡ����_������("real_pay", str֧�����)
            Call ���ýӿ�_��ȡ����_������("serial_no", strҵ�����к�)
            
            Select Case str������
            Case "003"
                cur�����ʻ� = Val(str֧�����)
            Case Is >= "900"
                cur�ֽ� = cur�ֽ� + Val(str֧�����)
            Case Else
                curͳ����� = curͳ����� + Val(str֧�����)
            End Select
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
    End If

    '��д�����¼
    '�ʻ��ۼ�����=����Ա��������;�ʻ��ۼ�֧��=����Ա���˲�������
    '�ۼƽ���ͳ��=���˲������;�ۼ�ͳ�ﱨ��=ҽ�Ʋ�����
    Call DebugTool("д���ս����¼")
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gCominfo_������.�ܷ��� & "," & cur�ֽ� & "," & 0 & "," & curͳ����� & "," & curͳ����� & ",0," & _
        0 & "," & cur�����ʻ� & ",'" & strҵ�����к� & "',null,null," & gCominfo_������.ҵ������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���������շ�����")
    
    gCominfo_������.ҵ�����к� = strҵ�����к�
    �������_������ = True
    
    '20031228:���:�������ID
    Call DebugTool("ȡ��Ʊ��Ϣ")
    Call GetBalance(lng����ID, lng����ID, strҵ�����к�, gCominfo_������.ҽԺ����)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ����������_������(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long, Optional ByVal bln�˷� As Boolean = True) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim strҵ�����к� As String, strҵ������ As String, str���˱�� As String
    Dim lng��¼ID As Long                               '������¼�Ľ���ID
    Dim int���� As Integer, int״̬ As Integer, str���ݺ� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '--��IC��
    '����IC���е���Ϣ
    If Not ���ýӿ�_׼��_������(Function_������.����_����) Then Exit Function
    If Not ���ýӿ�_ִ��_������ Then Exit Function
    'ȡ���صļ�¼��
    'If Not ���ýӿ�_ָ����¼��_������("ICInfo") Then Exit Sub
    If Not ���ýӿ�_��ȡ����_������("indi_id", str���˱��) Then Exit Function
    
    'У����˱���Ƿ���ȷ
    gstrSQL = "Select ���� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "У��ò����Ƿ�ʹ���Լ��Ŀ����н���", TYPE_������, lng����ID)
    If rsTemp!���� <> str���˱�� Then
        Err.Raise 9000, gstrSysName, "�ò��˲���ʹ���Լ���ҽ���������������ֹ��"
        Exit Function
    End If
    
    Call ��ȡ���˻�����Ϣ(lng����ID)
    'ȡ��������¼�Ľ���ID
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���²����Ľ���ID", lng����ID)
    lng��¼ID = rsTemp!����ID
    
    'ȡԭ�����¼����ϸ
    gstrSQL = "Select * From ���ս����¼ Where ����=[1] And ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����¼", TYPE_������, lng����ID)
    strҵ�����к� = rsTemp!֧��˳���
    strҵ������ = rsTemp!��ע
    
    '�������������¼
    With rsTemp
        gstrSQL = "zl_���ս����¼_insert(" & !���� & "," & lng��¼ID & "," & TYPE_������ & "," & !����ID & "," & _
            !��� & "," & -1 * Nvl(!�ʻ��ۼ�����, 0) & "," & -1 * Nvl(!�ʻ��ۼ�֧��, 0) & "," & -1 * Nvl(!�ۼƽ���ͳ��, 0) & "," & -1 * Nvl(!�ۼ�ͳ�ﱨ��, 0) & ",NULL,0,0,0," & _
            -1 * Nvl(!�������ý��, 0) & "," & -1 * Nvl(!ȫ�Ը����, 0) & "," & -1 * Nvl(!�����Ը����, 0) & "," & -1 * Nvl(!����ͳ����, 0) & "," & -1 * Nvl(!ͳ�ﱨ�����, 0) & ",0," & _
            0 & "," & -1 * Nvl(!�����ʻ�֧��, 0) & ",'" & strҵ�����к� & "',null,null," & strҵ������ & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�������������¼")
    End With
    
    If bln�˷� Then
        Call DebugTool("�����˷�")
        '�������ŵ��ݳ����ӿ�
        If Not ���ýӿ�_׼��_������(Function_������.����_���ŵ��ݳ���) Then Exit Function
        'д��ڲ���
        gstrField_������ = "hospital_id||serial_no||indi_id||busi_type||ic_flag||staff_no||staff_man"
        gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & strҵ�����к� & "||" & gCominfo_������.���˱�� & "||" & gCominfo_������.ҵ������ & "||" & mintICCard & "||" & Right(UserInfo.���, 5) & "||" & gstrUserName
        If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
        'ִ�нӿ�
        If Not ���ýӿ�_ִ��_������() Then Exit Function
    Else
        '���øķѹ���ʵ�����ŵ��ݳ�������Ϊ�ӿڲ�����ֱ�����м�ĳ��ҵ�񣬶��ķ��޴�����
        gCominfo_������.���˱�� = str���˱��
        gCominfo_������.ҵ������ = strҵ������
        gCominfo_������.ҵ�����к� = strҵ�����к�
        
        Call DebugTool("����ķѡ���ҵ�����ͣ�" & gCominfo_������.ҵ������ & "��ҵ�����кţ�" & gCominfo_������.ҵ�����к�)
        gstrSQL = "Select ��¼����,NO From ������ü�¼ Where ����ID=[1] And Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡԭ���������Ϣ", lng����ID)
        int���� = rsTemp!��¼����
        int״̬ = 2
        str���ݺ� = rsTemp!NO
        If Not ����ķ�_������(int����, int״̬, str���ݺ�) Then Exit Function
    End If
    
    ����������_������ = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ��Ժ�Ǽ�_������(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    Dim str���ֱ��� As String, str�������� As String
    Dim str��Ժ����ʱ�� As String
    Dim arrPatient
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
'   ��Ժ�ӿ���ڲ���
'    1   hospital_id        ҽ�ƻ�������    20  ��
'    2   indi_id            ���˱��    8   ��
'    3   busi_type          ҵ������    2   ��  "12"��סԺ
'    4   reg_staff          �Ǽ���Ա����    5   ��
'    5   reg_man            �Ǽ�������  10  ��
'    6   reg_flag           ��Ժ��ʽ    1   ��  "0"����ͨסԺ�Ǽ�
'    7   rela_hospital_id   ����ҽ�ƻ�������    20  ��
'    8   rela_serial_no     ����ҵ�����к�  12  ��
'    9   begin_date         ��Ժʱ��        ��  ��ʽ��YYYY-MM-DD
'    10  biz_times          ����סԺ����    2   ��
'    11  in_dept            ��Ժ���ұ��    3   ��
'    12  in_dept_name       ��Ժ��������    20  ��
'    13  in_area            ��Ժ�������    3   ��
'    14  in_area_name       ��Ժ��������    20  ��
'    15  in_bed             ��Ժ�������    10  ��
'    16  bed_type           ��λ����    1   ��  "0"����ͨ��λ��"1"�����ȣ�"2"�����ۣ�"3"���߸�
'    17  patient_id         סԺ��  20  ��
'    18  foregift           Ԥ�����ܶ�  10  ��
'    19  foregift_remain    Ԥ�������  10  ��  ����Ԥ�����ܶ�
'    20  in_disease         ��Ժ���    20  ��  ��������
'    21  ic_flag            �ÿ���־    1   ��  "0"����ʹ��IC����"1"��ʹ��IC��
'    22  note               ����֢      ȡ��Ժ��ϣ��������ڱ����ʻ��Ĳ���֢��
'   ���ؼ�¼��
'    1."ResultSet"������סԺ��ҵ�����кţ������������ݣ�
'    ���    �ֶ�    �ֶ�˵��    ��󳤶�    ��ע
'    1   serialno    ҵ�����к�  12
    
    On Error GoTo errHand
    Call ��ȡ���˻�����Ϣ(lng����ID, False)
    arrPatient = Split(��ȡ���������Ϣ(lng����ID, lng��ҳID), "||")
    
    '��ȡ������Ժ����
    gstrSQL = "Select ��Ժ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ����", lng����ID, lng��ҳID)
    str��Ժ����ʱ�� = Format(rsTemp!��Ժ����, "yyyy-MM-dd")
    
    '��ȡ��Ժ���ֱ���
    gstrSQL = "Select ����,���� From ���ղ��� Where ID=(Select ����ID From �����ʻ� Where ����ID=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ID", lng����ID)
    str���ֱ��� = rsTemp!����: str�������� = rsTemp!����
    
    gstrField_������ = "hospital_id||indi_id||busi_type||reg_staff||reg_man||reg_flag||" & _
                       "rela_hospital_id||rela_serial_no||begin_date||biz_times||in_dept||" & _
                       "in_dept_name||in_area||in_area_name||in_bed||bed_type||patient_id||" & _
                       "foregift||foregift_remain||in_disease||ic_flag||note"
    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.���˱�� & "||" & _
                       gCominfo_������.ҵ������ & "||" & gCominfo_������.����Ա���� & "||" & _
                       gstrUserName & "||0||||||" & str��Ժ����ʱ�� & "||" & _
                       arrPatient(סԺ�����ۼ�) & "||" & arrPatient(��Ժ���ұ��) & "||" & _
                       arrPatient(��Ժ��������) & "||" & arrPatient(��Ժ�������) & "||" & _
                       arrPatient(��Ժ��������) & "||" & arrPatient(��Ժ�������) & "||" & _
                       arrPatient(��λ����) & "||" & arrPatient(סԺ��) & "||" & _
                       "||||" & str���ֱ��� & "||" & mintICCard & "||" & arrPatient(��Ժ���)
    If Not ���ýӿ�_׼��_������(Function_������.��ͨסԺ_��Ժ�Ǽ�) Then Exit Function
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
    If Not ���ýӿ�_ִ��_������ Then Exit Function
    If Not ���ýӿ�_ָ����¼��_������("ResultSet") Then Exit Function
    If Not ���ýӿ�_��ȡ����_������("serialno", gCominfo_������.ҵ�����к�) Then Exit Function
    
    '���¸����ʻ��е���Ϣ
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'˳���','''" & gCominfo_������.ҵ�����к� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժҵ�����к�")
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    
    On Error Resume Next
    '���²���֢��Ϣ
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'����֢','''" & arrPatient(��Ժ���) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���֢��Ϣ")
    
    ��Ժ�Ǽ�_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽǳ���_������(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
                'ȡ��Ժ�Ǽ���֤�����ص�˳���
    On Error GoTo errHand
    
    If ����δ�����(lng����ID, lng��ҳID) Then
        MsgBox "��ҽ�����˴���δ����ã��������������Ժ�Ǽǣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��ȡ���˻�����Ϣ
    Call ��ȡ���˻�����Ϣ(lng����ID, False)
    
'    1   hospital_id    ҽ�ƻ�������    20  ��
'    2   serial_no      ҵ�����к�      12  ��
'    3   indi_id        ���˱��        8   ��
'    4   staff_no       ����Ա����      5   ��
'    5   staff_name     ����Ա����      10  ��
    gstrField_������ = "hospital_id||serial_no||indi_id||staff_no||staff_name"
    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.ҵ�����к� & "||" & _
                       gCominfo_������.���˱�� & "||" & gCominfo_������.����Ա���� & "||" & gstrUserName
    If Not ���ýӿ�_׼��_������(Function_������.��ͨסԺ_ȡ����Ժ) Then Exit Function
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
    If Not ���ýӿ�_ִ��_������ Then Exit Function
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_������(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim int��Ժ��� As Integer
    Dim str�������� As String, str��Ժ���� As String
    Dim str��Ժ���� As String, str��Ժ���� As String, str����֢ As String, str��Ժ��ʽ As String
    Dim blnҽ����Ժ As Boolean, bln���� As Boolean
    Dim arrPatient
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ��
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
                'ȡ��Ժ�Ǽ���֤�����ص�˳���
    
    'Modified By ���� ��������ɳ ԭ�����סԺ������޸ģ��˴������ѽ�����õĲ��˰���ҽ����Ժ
    blnҽ����Ժ = False
    Call ��ȡ���˻�����Ϣ(lng����ID, False)
    arrPatient = Split(��ȡ���������Ϣ(lng����ID, lng��ҳID), "||")
    
    '��ȡ���˵ĳ�Ժ���
    gstrSQL = "Select decode(��Ժ��ʽ,'תԺ',3,0) ��Ժ���,��Ժ��ʽ,��Ժ����,��Ժ���� From ������ҳ " & _
            " Where ����ID = [1] And ��ҳID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ��ʽ", lng����ID, lng��ҳID)
    int��Ժ��� = rsTemp!��Ժ���
    str��Ժ��ʽ = rsTemp!��Ժ��ʽ
    str�������� = Format(rsTemp!��Ժ����, "yyyy-MM-dd")
    If Not IsNull(rsTemp!��Ժ����) Then
        str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy-MM-dd")
    Else
        str��Ժ���� = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    End If

    '���û�ѡ����Ժ���ּ���Ժ���֣��Ա������Ժ��Ϣ
    If Not frm����ѡ��_����.ShowSelect(TYPE_������, lng����ID, lng��ҳID, str��Ժ����, str��Ժ����, str����֢) Then Exit Function
    
    If ����δ�����(lng����ID, lng��ҳID) Then  'HIS��ԺҪ������Ժ��Ϣ����Ժ����
        '�ȸ��²�����Ժ���
    '    1   hospital_id    ҽ�ƻ�������20  ��
    '    2   serial_no      ҵ�����к�  12  ��
    '    3   busi_type      ҵ������    2   ��  "12"��סԺ
    '    4   staff_no       ����Ա����  5   ��
    '    5   staff_name     ����Ա����  10  ��
    '    6   begin_date     ����ʱ��        ��  ��ʽ��YYYY-MM-DD
    '    7   in_dept        ��Ժ���ұ��3   ��
    '    8   in_dept_name   ��Ժ��������20  ��
    '    9   in_area        ��Ժ�������3   ��
    '    10  in_area_name   ��Ժ��������20  ��
    '    11  in_bed         ��Ժ�������10  ��
    '    12  bed_type       ��λ����    1   ��  "0"����ͨ��λ��"1"�����ȣ�"2"�����ۣ�"3"���߸�
    '    13  patient_id     סԺ��      20  ��
    '    14  old_patient_id ԭסԺ��    20  ��
    '    15  in_disease     ��Ժ���    20  ��  ��������
    '    16  note           ��ע        100 ��
        'ȡ����֢��Ϣ
        gstrField_������ = "hospital_id||serial_no||busi_type||staff_no||staff_name||begin_date||" & _
                        "in_dept||in_dept_name||in_area||in_area_name||in_bed||bed_type||patient_id||old_patient_id||in_disease||note||fin_disease"
        gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.ҵ�����к� & "||" & _
                        gCominfo_������.ҵ������ & "||" & gCominfo_������.����Ա���� & "||" & _
                        gstrUserName & "||" & str�������� & "||" & arrPatient(��Ժ���ұ��) & "||" & _
                        arrPatient(��Ժ��������) & "||" & arrPatient(��Ժ�������) & "||" & _
                        arrPatient(��Ժ��������) & "||" & arrPatient(��Ժ�������) & "||" & _
                        arrPatient(��λ����) & "||" & arrPatient(סԺ��) & "||" & _
                        arrPatient(סԺ��) & "||" & str��Ժ���� & "||" & str����֢ & "||" & str��Ժ����
        If Not ���ýӿ�_׼��_������(Function_������.סԺ��Ϣ_�޸�) Then Exit Function
        If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
        If Not ���ýӿ�_ִ��_������ Then Exit Function
    Else
        blnҽ����Ժ = True
        '��Ժ�ӿ���ڲ���
    '    1   hospital_id    ҽ�ƻ�������20  ��
    '    2   serial_no      ҵ�����к�  12  ��
    '    3   indi_id        ���˱��    8   ��
    '    4   busi_type      ҵ������    2   ��  "12"��סԺ
    '    5   fin_disease    ��Ժ����    20      ��������
    '    6   end_date       ��Ժ����            ��ʽ��YYYY-MM-DD
    '    7   fin_info       ��Ժ����    10
    '    8   end_staff      �ս��˹���  5   ��
    '    9   end_man        �ս�������  10  ��
    '    10  end_flag       �սᴦ��    1   ��  "0"�������ս᣻"3"��סԺתԺ
    '    11  begin_date     ����ʱ��        ��  ��ʽ��YYYY-MM-DD
    '    12  ic_flag        �ÿ���־    1   ��  "0"����ʹ��IC����"1"��ʹ��IC��
        '�жϸò����Ƿ�������û�н�����Ĳ��˷���Ϊ�㣬˵����Ҫ���þ���Ǽǳ���
        bln���� = False
        gstrSQL = "Select 1 From סԺ���ü�¼ Where ����ID=[1] And ��ҳID=[2] And Nvl(����ID,0)<>0 and Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�õ��þ���Ǽǳ���", lng����ID, lng��ҳID)
        If Not rsTemp.EOF Then
            bln���� = True
        End If
        If Not bln���� Then
            gstrField_������ = "hospital_id||serial_no||indi_id||staff_no||staff_name"
            gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.ҵ�����к� & "||" & _
                               gCominfo_������.���˱�� & "||" & gCominfo_������.����Ա���� & "||" & gstrUserName
            If Not ���ýӿ�_׼��_������(Function_������.��ͨסԺ_ȡ����Ժ) Then Exit Function
            If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
            If Not ���ýӿ�_ִ��_������ Then Exit Function
        Else
            gstrField_������ = "hospital_id||serial_no||indi_id||busi_type||fin_disease||end_date||" & _
                            "fin_info||end_staff||end_man||end_flag||begin_date||ic_flag"
            gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.ҵ�����к� & "||" & _
                            gCominfo_������.���˱�� & "||" & gCominfo_������.ҵ������ & "||" & _
                            str��Ժ���� & "||" & str��Ժ���� & "||" & str��Ժ��ʽ & "||" & _
                            gCominfo_������.����Ա���� & "||" & gstrUserName & "||" & int��Ժ��� & "||" & _
                            str�������� & "||" & mintICCard
            If Not ���ýӿ�_׼��_������(Function_������.��Ժ�Ǽ�) Then Exit Function
            If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
            If Not ���ýӿ�_ִ��_������ Then Exit Function
        End If
    End If
    
    '����HIS��Ժ
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
    
    MsgBox IIf(blnҽ����Ժ, "ҽ����Ժ", "HIS��Ժ") & "����ɹ���", vbInformation, gstrSysName
    ��Ժ�Ǽ�_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽǳ���_������(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim int��Ժ��� As Integer, strסԺ�� As String
    Dim rsTemp As New ADODB.Recordset
'    1   hospital_id    ҽ�ƻ�������    20  ��
'    2   busi_type      ҵ������        2   ��  "12"��סԺ
'    3   indi_id        ���˱��        8   ��
'    4   end_flag       �սᴦ���־    1   ��  "0"�������ս᣻"3"��תԺסԺ
'    5   serial_no      ҵ�����к�      12  ��
'    6   patient_id     סԺ��          20  ��
    On Error GoTo errHand
        
    '���ڳ�Ժ����ܽ���������ҵ������ȡ���һ��סԺ�����ҵ�����ͺ�ҵ�����к�
    gstrSQL = " Select ֧��˳���,��ע From ���ս����¼ " & _
              " Where ����=2 And ����=" & TYPE_������ & " And ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡסԺ�����Ϣ", lng����ID, lng��ҳID)
    If Not rsTemp.EOF Then
        gCominfo_������.ҵ�����к� = rsTemp!֧��˳���
        gCominfo_������.ҵ������ = rsTemp!��ע
    Else
        Call ��ȡ���˻�����Ϣ(lng����ID, False)
    End If
    
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        '��ȡ���˵ĳ�Ժ���
        gstrSQL = "Select decode(A.��Ժ��ʽ,'תԺ',3,0) ��Ժ��ʽ,B.סԺ�� " & _
                " From ������ҳ A,������Ϣ B " & _
                " Where A.����ID=B.����ID And A.����ID = [1] And A.��ҳID = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ��ʽ", lng����ID, lng��ҳID)
        int��Ժ��� = rsTemp!��Ժ��ʽ
        strסԺ�� = Nvl(rsTemp!סԺ��)
        
        '������Ժ�ӿ���ڲ���
        gstrField_������ = "hospital_id||busi_type||indi_id||end_flag||serial_no||patient_id"
        gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.ҵ������ & "||" & _
                        gCominfo_������.���˱�� & "||" & int��Ժ��� & "||" & _
                        gCominfo_������.ҵ�����к� & "||" & strסԺ��
        If Not ���ýӿ�_׼��_������(Function_������.ȡ����Ժ) Then Exit Function
        If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
        If Not ���ýӿ�_ִ��_������ Then Exit Function
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function סԺ�������_������(rsExse As Recordset, ByVal lng����ID As Long) As String
    Dim bln����ͷ As Boolean, blnUp As Boolean, blnMoveNext As Boolean, blnTrans As Boolean
    Dim curYB�ܶ� As Currency
    Dim str֧����� As String, str�������� As String, str������ As String
    Dim str�������� As String, strҽ����� As String, arrҽ�����
    Dim strҽ����� As String, strҽ������ As String, str��Ŀ���� As String
    Dim str��� As String, str���� As String, str���� As String
    Dim lngRecord As Long, lngԭʼ����ID As Long, lng������ As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rsDetail As New ADODB.Recordset
    Dim rsPhysic As New ADODB.Recordset
    Dim gcn�ϴ� As New ADODB.Connection
    Dim i As Long
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    '�ӿڷ��صı������ȥ����סԺ�ڼ�����������Ļ��ܽ��󣬲��Ǳ��ε�ʵ�ʱ�����
    'rsExse��¼���е��ֶ��嵥
    '��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID,
    '�շ����,�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������
    '���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��,�Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    On Error GoTo errHand
    
    'ȡҽ�����������
    Call DebugTool("ȡҽ������")
    strҽ����� = "": strҽ������ = ""
    If Not IsNull(rsExse!ҽ��) Then
        gstrSQL = "Select ��� From ��Ա�� Where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����", CStr(rsExse!ҽ��))
        strҽ����� = rsTemp!���
        strҽ������ = rsExse!ҽ��
    End If
    
    '�´�һ�����������ϴ�������ϸ�������ظ��ϴ�
    Set gcn�ϴ� = GetNewConnection
    Call DebugTool("��ȡ���˻�����Ϣ")
    Call ��ȡ���˻�����Ϣ(lng����ID, False)
    gCominfo_������.�ܷ��� = 0
    '��δ�ϴ��ķ�����ϸ�ϴ�
'    1   hospital_id    ҽ�ƻ�������    20  ��
'    2   indi_id        ���˱��        8   ��
'    3   busi_type      ҵ������        2   ��  "12"��סԺ
'    4   serial_no      ҵ�����к�      12  ��
'    5   ordinal_no     �ڲ�����        3   ��
'    6   input_staff    ¼���˹���      5   ��
'    7   input_man      ¼��������      10  ��
'    8   recipe_no      ������          20  ��
'--------2004-01-12����--------
'    9   doctor_no      ����ҽ�����    12
'    10  doctor_name    ����ҽ������    10
'------------------------------
'    ���ݼ�����������������ϸ��Ϣ��������Ϊ��"FeeInfo"��������������:
'    ���    ���    ���˵��    ��󳤶�    �Ƿ��Ϊ��  ��ע
'    1   medi_item_type ��ĿҩƷ����    1   ��  "0"��������Ŀ��"1"����ҩ��"2"���г�ҩ��"3"���в�ҩ
'    2   his_item_code  ҽԺҩƷ��Ŀ����20  ��
'    3   his_item_name  ҽԺҩƷ��Ŀ����50  ��
'    4   model          ����            30  ��
'    5   factory        ����            50  ��
'    6   standard       ���            30  ��
'    7   fee_date       ���÷���ʱ��        ��  ��ʽ��YYYY-MM-DD���ķ�ʱ���Բ���¼
'    8   unit           ������λ        10  ��
'    9   price          ����            12  ��
'    10  dosage         ����            12  ��
'    11  money          ���            12  ��
'    12  usage_flag     ��ҩ��־    1   ��  "0"����ͨ��"1"����Ժ��ҩ��"2"������
'    13  usage_days     ��Ժ��ҩ����    3   ��
'    14  opp_serial_fee ��Ӧ�������к�  12  ��  �˷�ʱʹ��
    
    'סԺԤ����ʱ����Ҫ�ϴ��Ķ��Ǵ�λ�ѵ��Զ���������Ĵ�����ȫ��������һ�Ŵ����Ԥ�����ϴ�����ϸ�Է�û���棩
    Call DebugTool("�ϴ���ϸ")
    For i = 1 To 2
        With rsExse
            '20031231:���:��λ�ѵĳ������ݲ����µ���,��Ӧ�ú��ϴ���������,��Ȼȡ����Ҫ������opp_serial_fee
            If i = 1 Then
                .Filter = "���>=0"
            ElseIf i = 2 Then
                .Filter = "���<0"
            End If
            
            'todo :��ȡ��Ӧ��ҽ���������кţ����ڸ���¼���޷��ϴ�����Ҫ����������Ǵ�λ�Ѳ����ĸ����ݣ��ӿ�ֻ�ܴ�������ĵ��ݣ�
            strҽ����� = ""
            Do While Not .EOF
                If Nvl(!���, 0) < 0 And Nvl(!�Ƿ��ϴ�, 0) = 0 Then
                    '20040105:���:��ֱ�ӳ����˷Ѻ����븺���˷ѷֿ�����
                    If !��¼״̬ = 1 Then
                        'Ϊ�µ������С��0,�����븺���˷�,����ȷ��Ӧ��ԭʼ���ü�¼
                        If !�շ���� = 5 Or !�շ���� = 6 Or !�շ���� = 7 Then
                            gstrSQL = "Select Decode(trim(��ʶ��),NULL,����,'',����,��ʶ��) ���� From ҩƷĿ¼ Where ҩƷID=[1]"
                        Else
                            gstrSQL = "Select Decode(trim(��ʶ����),NULL,����,'',����,��ʶ����) ���� From �շ�ϸĿ Where ID=[2]"
                        End If
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ŀ�ı�����ʶ��", CLng(!�շ�ϸĿID))
                        strҽ����� = strҽ����� & IIf(strҽ����� = "", "", "|") & GetInsureSerial2(rsTemp!����, Nvl(!���, 0), IIf(strҽ����� = "", True, False))
                    Else
                        gstrSQL = " Select ID From סԺ���ü�¼" & _
                                  " Where ��¼����=[1] And ��¼״̬=3 And NO=[2] And ���=[3]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡԭʼ��¼�ķ���ID", CLng(!��¼����), CStr(!NO), CLng(!���))
                        lngԭʼ����ID = rsTemp!ID
                        strҽ����� = strҽ����� & IIf(strҽ����� = "", "", "|") & GetInsureSerial(lngԭʼ����ID, IIf(strҽ����� = "", True, False))
                    End If
                End If
                .MoveNext
            Loop
            
            '�ϴ�������ϸ
            lngRecord = 1: lng������ = 1
            
            If strҽ����� <> "" Then arrҽ����� = Split(strҽ�����, "|")
            If .RecordCount <> 0 Then .MoveFirst
            
            bln����ͷ = False: blnUp = False '20031231:���
            gcn�ϴ�.BeginTrans: blnTrans = True '20031231:���:ȷ��HIS��Է��ϴ�����һ��
            
            Do While Not .EOF
                gCominfo_������.�ܷ��� = gCominfo_������.�ܷ��� + !���
                
                '������ͷ����ֻ֤��һ�Σ���Ϊû���Ķ����Զ���������ĵ��ݣ�
                If Nvl(!�Ƿ��ϴ�, 0) = 0 Then
                    blnUp = True
                    If Not bln����ͷ Then
                        Call DebugTool("�ϴ���ϸ-����ͷ")
                        gstrField_������ = "hospital_id||indi_id||busi_type||serial_no||ordinal_no||input_staff||input_man||recipe_no||doctor_no||doctor_name"
                        gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.���˱�� & "||" & _
                                        gCominfo_������.ҵ������ & "||" & gCominfo_������.ҵ�����к� & "||||" & _
                                        gCominfo_������.����Ա���� & "||" & gstrUserName & "||||" & strҽ����� & "||" & strҽ������
                        If Not ���ýӿ�_׼��_������(Function_������.סԺ����_�ϴ���ϸ) Then
                            If blnTrans Then gcn�ϴ�.RollbackTrans
                            Exit Function
                        End If
                        If Not ���ýӿ�_д��ڲ���_������(1) Then
                            If blnTrans Then gcn�ϴ�.RollbackTrans
                            Exit Function
                        End If
                        bln����ͷ = True
                        blnMoveNext = False
                        
                        'ָ����¼��
                        If Not ���ýӿ�_ָ����¼��_������("FeeInfo") Then
                            If blnTrans Then gcn�ϴ�.RollbackTrans
                            Exit Function
                        End If
                    End If
                    
                    '��������ϸ
                    Call DebugTool("�ϴ���ϸ-������ϸ")
                    gstrSQL = "Select ��ʶ����,����,����,��� From �շ�ϸĿ Where ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ�ϸĿ�����Ϣ", CLng(!�շ�ϸĿID))
                    str��� = Nvl(rsTemp!���)
                    str���� = ""
                    If InStr(1, str���, "��") <> 0 Then
                        str���� = ToVarchar(Trim(Split(str���, "��")(1)), 50)
                        str��� = ToVarchar(Trim(Split(str���, "��")(0)), 30)
                    Else
                        str��� = ToVarchar(Trim(str���), 30)
                    End If
            
                    '�����ҩƷ����ȥȡ����
                    Call DebugTool("�ϴ���ϸ-������ϸ��ȡ��ʶ�룩")
                    str���� = ""
                    str��Ŀ���� = ""
                    If !�շ���� = 5 Or !�շ���� = 6 Or !�շ���� = 7 Then
                        gstrSQL = " Select C.����,C.��ʶ��,B.���� ���� From ҩƷ��Ϣ A,ҩƷ���� B,ҩƷĿ¼ C " & _
                                  " Where A.����=B.���� And A.ҩ��ID=C.ҩ��ID And C.ҩƷID =[1]"
                        Set rsPhysic = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����", CLng(!�շ�ϸĿID))
                        str���� = Nvl(rsPhysic!����)
                        str��Ŀ���� = Nvl(rsPhysic!��ʶ��)
                        If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsPhysic!����)
                    Else
                        str��Ŀ���� = Nvl(rsTemp!��ʶ����)
                        If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsTemp!����)
                    End If
                    
                    gstrSQL = "Select ID From סԺ���ü�¼ Where NO=[1] And ��¼����=[2] And ��¼״̬=[3] And ���=[4]"
                    Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ID", CStr(!NO), CLng(!��¼����), CLng(!��¼״̬), CLng(!���))
                    
                    Call DebugTool("�ϴ���ϸ-������ϸ��ȡ�Ǽ�ʱ�䣩")
                    str�������� = Format(!����ʱ��, "yyyy-MM-dd")
                    gstrField_������ = "medi_item_type||his_item_code||his_item_name||model||factory||standard||" & _
                                "fee_date||unit||price||dosage||money||usage_flag||usage_days||opp_serial_fee||hos_serial"
                    If Nvl(!���, 0) >= 0 Then
                        gstrValue_������ = IIf(!�շ���� = 5, "1", IIf(!�շ���� = 6, "2", IIf(!�շ���� = 7, "3", "0"))) & "||" & _
                                    str��Ŀ���� & "||" & Nvl(rsTemp!����) & "||" & str���� & "||" & str���� & "||" & str��� & "||" & _
                                    str�������� & "||" & Nvl(!���㵥λ) & "||" & !�۸� & "||" & !���� & "||" & Nvl(!���, 0) & "||0||||||" & rsDetail!ID
                    Else
                        '20040105:���:�˷�ʱ,��Hos_Serial��ΪHIS����˷Ѽ�¼��ID,�����Ǳ��˷Ѽ�¼��ID,��Ϊ���������޶�Ӧԭʼ��¼
                        gstrValue_������ = IIf(!�շ���� = 5, "1", IIf(!�շ���� = 6, "2", IIf(!�շ���� = 7, "3", "0"))) & "||" & _
                                    str��Ŀ���� & "||" & Nvl(rsTemp!����) & "||" & str���� & "||" & str���� & "||" & str��� & "||" & _
                                    str�������� & "||" & Nvl(!���㵥λ) & "||" & !�۸� & "||" & !���� & "||" & Nvl(!���, 0) & "||0||||" & _
                                    arrҽ�����(lng������ - 1) & "||" & rsDetail!ID
                    End If
                    Call DebugTool("�ϴ���ϸ-������ϸ����ɣ�")
                    
                    If bln����ͷ And blnMoveNext Then
                        If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then
                            If blnTrans Then gcn�ϴ�.RollbackTrans
                            Exit Function
                        End If
                    End If
                    If Not ���ýӿ�_д��ڲ���_������(lngRecord) Then
                        If blnTrans Then gcn�ϴ�.RollbackTrans
                        Exit Function
                    End If
                    
                    lngRecord = lngRecord + 1
                    If !��� < 0 Then lng������ = lng������ + 1
                    blnMoveNext = True
                    
                    '�����ϴ���־����Ϊ��ϸ������ȷ�ϴ��󣬲��ܱ�֤�������ȷ��
                    'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                    gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsExse("NO") & "'," & rsExse("���") & "," & rsExse("��¼����") & "," & rsExse("��¼״̬") & ")"
                    gcn�ϴ�.Execute gstrSQL, , adCmdStoredProc
                End If
                .MoveNext
            Loop
            If blnUp Then
                If Not ���ýӿ�_ִ��_������() Then
                    If blnTrans Then gcn�ϴ�.RollbackTrans
                    Exit Function
                End If
            End If
            gcn�ϴ�.CommitTrans: blnTrans = False
        End With
    Next
    
    
    '----------------------------------------------------------------------------------------------------------------
'    Ԥ������ڲ���
'    1   hospital_id     ҽ�ƻ�������    20  ��
'    2   serial_no   ҵ�����к�  12  ��
    Call DebugTool("��Ԥ����ӿ�")
    gstrField_������ = "hospital_id||serial_no"
    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.ҵ�����к�
    If Not ���ýӿ�_׼��_������(Function_������.סԺ����_Ԥ����) Then Exit Function
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
    If Not ���ýӿ�_ִ��_������ Then Exit Function
'    ���ؼ�¼����Ϣ
'    1."CalcResultInfo"�����˵ı���סԺδ�շ�����Ϣ�������������ݣ�
'    ���    �ֶ�    �ֶ�˵��    ��󳤶�    ��ע
'    1   fund_id        �������    3
'    2   fund_name      ��������    30
'    3   real_pay       ֧�����    12  ��λ��Ԫ
'    2."lastbalance"�������ʻ������Ϣ�������������ݣ�
'    ���    �ֶ�    �ֶ�˵��    ��󳤶�    ��ע
'    1   Last_balance   �����ʻ����18  ��λ��Ԫ
    If Not ���ýӿ�_ָ����¼��_������("CalcResultInfo") Then Exit Function
    curYB�ܶ� = 0
    
    Call DebugTool("׼��ȡԤ������")
    If ���ýӿ�_��¼��_������ Then
        Do While True
            Call ���ýӿ�_��ȡ����_������("fund_id", str������)
            Call ���ýӿ�_��ȡ����_������("fund_name", str��������)
            Call ���ýӿ�_��ȡ����_������("real_pay", str֧�����)
            curYB�ܶ� = curYB�ܶ� + Val(str֧�����)
            If str������ <> 999 And Val(str֧�����) <> 0 Then
                סԺ�������_������ = סԺ�������_������ & IIf(סԺ�������_������ = "", "", "|") & _
                            IIf(str������ = "003", "�����ʻ�", str��������) & ";" & str֧����� & ";0" ' & IIf(str������ = "003", "1", "0")
            End If
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
    End If
    
    Call DebugTool("�жϷ����Ƿ����")
    If Format(gCominfo_������.�ܷ���, "#####0.00") <> Format(curYB�ܶ�, "#####0.00") Then
        If MsgBox("HISϵͳ���ܷ�����ҽ�����ص��ܷ��ò�����Ҫ����������" & vbCrLf & _
        "    HIS�ܷ��ã�" & Format(gCominfo_������.�ܷ���, "#####0.00") & Space(5) & "ҽ���ܷ��ã�" & Format(curYB�ܶ�, "#####0.00"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            סԺ�������_������ = ""
            Exit Function
        End If
    End If
    If סԺ�������_������ = "" Then סԺ�������_������ = "�����ʻ�;0;0"
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcn�ϴ�.RollbackTrans
End Function

Public Function סԺ����_����(lng����ID As Long, ByVal lng����ID As Long) As Boolean
    Dim lng��ҳID As Long, lngRecord As Long
    Dim bln����ͷ As Boolean, blnUp As Boolean, blnMoveNext As Boolean
    Dim str�������� As String, str������ As String, str֧����� As String, str�������� As String, str���˱�� As String
    Dim cur�ʻ�֧�� As Currency
    Dim curͳ����� As Currency, cur�ֽ� As Currency
    Dim rsTemp As New ADODB.Recordset
    Dim rsExse As New ADODB.Recordset
    '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
        '������㣨���ص����ݼ�ȥ���ν������ݣ��͵��ڱ��ε���ʵ�������ݣ�
    On Error GoTo errHand
    
    '--��IC��
    '����IC���е���Ϣ
    If Not ���ýӿ�_׼��_������(Function_������.����_����) Then Exit Function
    If Not ���ýӿ�_ִ��_������ Then Exit Function
    'ȡ���صļ�¼��
    'If Not ���ýӿ�_ָ����¼��_������("ICInfo") Then Exit Sub
    If Not ���ýӿ�_��ȡ����_������("indi_id", str���˱��) Then Exit Function
    
    'У����˱���Ƿ���ȷ
    gstrSQL = "Select ���� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "У��ò����Ƿ�ʹ���Լ��Ŀ����н���", TYPE_������, lng����ID)
    If rsTemp!���� <> str���˱�� Then
        Err.Raise 9000, gstrSysName, "�ò��˲���ʹ���Լ���ҽ���������������ֹ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Modified By ���� ��������ɳ ԭ����ΪסԺ��������ǳ�������סԺ���н����¼�����HISδ��Ժ��ҽ�����ˣ���������г�Ժ����
    If Not ҽ�������Ѿ���Ժ(lng����ID) Then
        MsgBox "ҽ�����˱����Ժ����ܰ���ҽ����Ժ���㣡", vbInformation, gstrSysName
        Exit Function
    End If
    Call ��ȡ���˻�����Ϣ(lng����ID, False)
    
    'Ԥ����ʱ���ϴ���ϸ���˴�����Ҫ�ٴ��ϴ�
    '��ȡ��ҳID
    gstrSQL = "Select סԺ���� ��ҳID From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҳID", lng����ID)
    lng��ҳID = rsTemp!��ҳID
    
    '��ȡ�ʻ�֧����
    gstrSQL = "Select Nvl(A.��Ԥ��,0) �����ʻ� " & _
        " From ����Ԥ����¼ A,�����ʻ� B " & _
        " Where A.����ID=B.����ID And B.����=" & TYPE_������ & _
        " And A.���㷽ʽ='�����ʻ�' And A.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ʻ�֧����", lng����ID)
    cur�ʻ�֧�� = 0
    If Not rsTemp.EOF Then
        cur�ʻ�֧�� = Nvl(rsTemp!�����ʻ�, 0)
    End If
    
'    ��ʽ����
'    ���    ���    ���˵��    ��󳤶�    �Ƿ��Ϊ��  ��ע
'    1   hospital_id    ҽ�ƻ�������    20  ��
'    2   serial_no      ҵ�����к�      12  ��
'    3   ic_flag        �ÿ���־        1   ��  "0"����ʹ��IC����"1"��ʹ��IC��
'    4   debit_money    ������ĸ����ʻ�֧�����    12  ��  ��λ��Ԫ
'    5   indi_id        ���˱��        8   ��
    gstrField_������ = "hospital_id||serial_no||ic_flag||debit_money||indi_id"
    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.ҵ�����к� & "||" & _
                    mintICCard & "||" & cur�ʻ�֧�� & "||" & gCominfo_������.���˱��
    If Not ���ýӿ�_׼��_������(Function_������.סԺ����_��ʽ����) Then Exit Function
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
    If Not ���ýӿ�_ִ��_������() Then Exit Function
    
    'ָ����¼��
    If Not ���ýӿ�_ָ����¼��_������("payment") Then Exit Function
'    1   fund_id;    �������    3
'    2   fund_name   ��������    30
'    3   real_pay    ֧�����    12
'    4   serial_no   ҵ�����к�  12
    If ���ýӿ�_��¼��_������ Then
        Do While True
            Call ���ýӿ�_��ȡ����_������("fund_id", str������)
            Call ���ýӿ�_��ȡ����_������("fund_name", str��������)
            Call ���ýӿ�_��ȡ����_������("real_pay", str֧�����)
            
            Select Case str������
            Case "003"
                cur�ʻ�֧�� = Val(str֧�����)
            Case Is >= "900"
                cur�ֽ� = cur�ֽ� + Val(str֧�����)
            Case Else
                curͳ����� = curͳ����� + Val(str֧�����)
            End Select
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
    End If
    
    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʼ�¼�����ϴ���־")
    
    '��д���ս����¼
    '�ʻ��ۼ�����=����Ա��������;�ʻ��ۼ�֧��=����Ա���˲�������
    '�ۼƽ���ͳ��=���˲������;�ۼ�ͳ�ﱨ��=ҽ�Ʋ�����
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & lng��ҳID & "," & 0 & "," & 0 & "," & 0 & "," & _
        cur�ʻ�֧�� + cur�ֽ� + curͳ����� & "," & cur�ֽ� & "," & 0 & "," & curͳ����� & "," & curͳ����� & ",0," & _
        0 & "," & cur�ʻ�֧�� & ",'" & gCominfo_������.ҵ�����к� & "',null,null," & gCominfo_������.ҵ������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ��������")
    
    סԺ����_���� = True
    'ͬʱ�����Ժ�Ǽ�
    Call ��Ժ�Ǽ�_������(lng����ID, lng��ҳID)
    
    '20031228:���:�������ID
    Call GetBalance(lng����ID, lng����ID, gCominfo_������.ҵ�����к�, gCominfo_������.ҽԺ����)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function סԺ�������_������(lng����ID As Long) As Boolean
    Dim lng��¼ID As Long, lng����ID As Long, lng��ҳID As Long
    Dim strҵ�����к� As String, strҵ������ As String
    Dim rsTemp As New ADODB.Recordset
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '      4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    On Error GoTo errHand
    Call ��ȡ���˻�����Ϣ(lng����ID, False)                '��ȡҵ�����к�
    
    'ȡ������¼�Ľ���ID
    gstrSQL = "select distinct A.ID,A.����ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���²����Ľ���ID", lng����ID)
    lng��¼ID = rsTemp!ID
    lng����ID = rsTemp!����ID
    
    '��ȡ��ҳID
    gstrSQL = "Select סԺ���� ��ҳID From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҳID", lng����ID)
    lng��ҳID = rsTemp!��ҳID
    
    'ȡԭ�����¼����ϸ
    gstrSQL = "Select * From ���ս����¼ Where ����=" & TYPE_������ & " And ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����¼", lng����ID)
    strҵ�����к� = rsTemp!֧��˳���
    strҵ������ = rsTemp!��ע
    
    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng��¼ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʼ�¼�����ϴ���־")
    
    'todo ��ע�⣬�ǻ��˱���סԺ�ڼ����н��㵥�ݣ��˹����账��
    '�������������¼
    With rsTemp
        gstrSQL = "zl_���ս����¼_insert(" & !���� & "," & lng��¼ID & "," & TYPE_������ & "," & !����ID & "," & _
            !��� & "," & -1 * Nvl(!�ʻ��ۼ�����, 0) & "," & -1 * Nvl(!�ʻ��ۼ�֧��, 0) & "," & -1 * Nvl(!�ۼƽ���ͳ��, 0) & "," & -1 * Nvl(!�ۼ�ͳ�ﱨ��, 0) & "," & lng��ҳID & ",0,0,0," & _
            -1 * Nvl(!�������ý��, 0) & "," & -1 * Nvl(!ȫ�Ը����, 0) & "," & -1 * Nvl(!�����Ը����, 0) & "," & -1 * Nvl(!����ͳ����, 0) & "," & -1 * Nvl(!ͳ�ﱨ�����, 0) & ",0," & _
            0 & "," & -1 * Nvl(!�����ʻ�֧��, 0) & ",'" & strҵ�����к� & "',null,null," & strҵ������ & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�������������¼")
    End With
    
    '׼������סԺ��������ӿ�
'    ���    ���    ���˵��    ��󳤶�    �Ƿ��Ϊ��  ��ע
'    1   hospital_id     ҽ�ƻ�������    20  ��
'    2   serial_no   ҵ�����к�  12  ��
'    3   ic_flag �ÿ���־    1   ��  "0"����ʹ��IC����"1"��ʹ��IC��
    gstrField_������ = "hospital_id||serial_no||ic_flag"
    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & strҵ�����к� & "||" & mintICCard
    If Not ���ýӿ�_׼��_������(Function_������.סԺ����_�������) Then Exit Function
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
    If Not ���ýӿ�_ִ��_������ Then Exit Function
    
    סԺ�������_������ = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ��ݱ�ʶ_������(Optional bytType As Byte, Optional lng����ID As Long) As String
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-���1-סԺ
    '���أ��ջ���Ϣ��
    'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
    '      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
    '      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    ��ݱ�ʶ_������ = frmIdentify����.GetPatient(bytType, lng����ID)
End Function

Public Function �������_������(ByVal strSelfNo As String, Optional ByVal bln���� As Boolean = True) As Currency
    '����: ���ظ����ʻ���������ýӿں���ʧ�ܣ����������֤ʱ�����
    '����: �Ƿ����
    '����: ���ظ����ʻ����
    Dim lng����ID As Long
    Dim str���˱�� As String, str�α��˵�λ As String, strData As String
    Dim cur�ʻ���� As Currency
    Dim rsAccount As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = " Select A.����,B.������λ,Nvl(A.�ʻ����,0) �ʻ���� From �����ʻ� A,������Ϣ B " & _
              " Where A.����ID=B.����ID And A.����=" & TYPE_������
    If bln���� Then
        gstrSQL = gstrSQL & " And A.ҽ����=[2]"
    Else
        gstrSQL = gstrSQL & " And A.����ID=[1]"
    End If
    Set rsAccount = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ż���λ����", CLng(strSelfNo), strSelfNo)
    str���˱�� = rsAccount!����
    If InStr(1, rsAccount!������λ, "]") <> 0 Then str�α��˵�λ = Mid(rsAccount!������λ, 2, InStr(1, rsAccount!������λ, "]") - 2)
    cur�ʻ���� = rsAccount!�ʻ����
    �������_������ = cur�ʻ����
    If InStr(1, rsAccount!������λ, "]") = 0 Then
        Call DebugTool("�޵�λ���룬�������ݿ����ʻ����")
        Exit Function
    End If
    
    '--��ȡ�����ʻ�������סԺû�з��أ�ǿ��ȡһ�Σ�
    If Not ���ýӿ�_׼��_������(Function_������.����_�������) Then Exit Function
    
    'д��ڲ���
'    1   fund_id    ������    3   ��
'    2   indi_id    ���˱��    8   ��
'    3   corp_ID    ��λ���    3
    'Modified By ���� ��������ɳ ԭ����Ҫ�ഫһ��������corp_id��
    gstrField_������ = "fund_id||indi_id||corp_id"
    gstrValue_������ = "003||" & str���˱�� & "||" & str�α��˵�λ
    Call ���ýӿ�_д��ڲ���_������(1)
    If Not ���ýӿ�_ִ��_������ Then Exit Function
    If Not ���ýӿ�_ָ����¼��_������("PersonAccount") Then Exit Function
    Call ���ýӿ�_��ȡ����_������("last_balance", strData)
    cur�ʻ���� = Val(strData)
    
    '���¸����ʻ�������
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'�ʻ����','" & cur�ʻ���� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʻ����")
    �������_������ = cur�ʻ����
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ��ȡ���˻�����Ϣ(ByVal lng����ID As Long, Optional bln���� As Boolean = True) As Boolean
    Dim str���˱�� As String, strҵ�����к� As String, str�������� As String
    Dim rsTemp As New ADODB.Recordset
    '���ز��˵ĸ��˱�š�ҵ�����к�,��������
    gstrSQL = " Select ����,˳���,�Ҷȼ�,��ǰ״̬,ҵ������ From �����ʻ�" & _
              " Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˻�����Ϣ", TYPE_������, lng����ID)
    If rsTemp.EOF Then Exit Function
    
    gCominfo_������.���˱�� = rsTemp!����
    gCominfo_������.ҵ�����к� = Nvl(rsTemp!˳���)
    If bln���� Then
        gCominfo_������.ҵ������ = rsTemp!�Ҷȼ�
    Else
        gCominfo_������.ҵ������ = rsTemp!ҵ������
    End If
    
    Call �������_������(lng����ID, False)
    ��ȡ���˻�����Ϣ = True
End Function

Private Function ��ȡ���������Ϣ(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
    Dim lngסԺ���� As Long
    Dim str��Ժ���ұ�� As String, str��Ժ�������� As String, str��Ժ������� As String, str��Ժ�������� As String, str��Ժ������� As String, str��λ���� As String
    Dim strסԺ�� As String, str��Ժ��� As String, str��Ժ��� As String
    Dim rsTemp As New ADODB.Recordset
    '��ȡ���������Ϣ������סԺ��������Ժ���ұ�š���Ժ�������ơ���Ժ������š���Ժ�������ơ���Ժ������š���λ���͡�סԺ�š���Ժ��ϡ���Ժ��ϣ�
'    ��λ����
'    "0"����ͨ��λ
'    "1"������
'    "2"������
'    "3"���߸�
    
    gstrSQL = " Select nvl(סԺ�����ۼ�,1) סԺ���� From �ʻ������Ϣ " & _
              " Where ����ID=[1] And ���=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���������Ϣ", lng����ID, CStr(Format(zlDatabase.Currentdate, "yyyy")))
    If Not rsTemp.EOF Then lngסԺ���� = rsTemp!סԺ����
    If lngסԺ���� = 0 Then lngסԺ���� = 1
    
    '��ȡ��Ժ�����Ϣ
    gstrSQL = "select C.���� ��Ժ���ұ��,C.���� ��Ժ��������,B.���� ��Ժ�������,B.���� ��Ժ��������, " & _
             " A.��Ժ���� ��Ժ�������,F.��λ����,E.סԺ�� סԺ�� " & _
             " from ������ҳ A,���ű� B,���ű� C,������Ϣ E, " & _
             " (Select D.���� ��λ����,F.����,F.����ID,F.����ID  From ��λ�ȼ� D ,��λ״����¼ F Where F.�ȼ�ID=D.���) F " & _
             " Where A.��Ժ����ID=B.ID(+) And A.��Ժ����ID=C.ID(+) And A.����ID=E.����ID ANd A.����ID=[1] And A.��ҳID=[2]" & _
             " And A.��Ժ����=F.����(+) And F.����ID(+)=A.��Ժ����ID And F.����ID(+)=A.��Ժ����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ�����Ϣ", lng����ID, lng��ҳID)
    If Not rsTemp.EOF Then
        str��Ժ���ұ�� = ToVarchar(Nvl(rsTemp!��Ժ���ұ��), 3)
        str��Ժ�������� = ToVarchar(Nvl(rsTemp!��Ժ��������), 20)
        str��Ժ������� = ToVarchar(Nvl(rsTemp!��Ժ�������), 3)
        str��Ժ�������� = ToVarchar(Nvl(rsTemp!��Ժ��������), 20)
        str��Ժ������� = ToVarchar(Nvl(rsTemp!��Ժ�������), 10)
        str��λ���� = Nvl(rsTemp!��λ����)
        Select Case str��λ����
        Case "����"
            str��λ���� = 1
        Case "����"
            str��λ���� = 2
        Case "�߸�"
            str��λ���� = 3
        Case Else
            str��λ���� = 0
        End Select
        strסԺ�� = Nvl(rsTemp!סԺ��)
    End If
    
    '��ȡ���Ժ��ϣ��������룩
    str��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, True, False, True)
    str��Ժ��� = Split(str��Ժ���, "|")(1)
    str��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, False, True)
    str��Ժ��� = Split(str��Ժ���, "|")(1)
    ��ȡ���������Ϣ = lngסԺ���� & "||" & str��Ժ���ұ�� & "||" & str��Ժ�������� & "||" & _
                    str��Ժ������� & "||" & str��Ժ�������� & "||" & str��Ժ������� & "||" & _
                    str��λ���� & "||" & strסԺ�� & "||" & str��Ժ��� & "||" & str��Ժ���
End Function

Public Function �����Ǽ�_������(ByVal int���� As Integer, ByVal int״̬ As Integer, ByVal str���ݺ� As String) As Boolean
    Dim lng����ID As Long, strNO As String, str�������� As String
    Dim blnUp As Boolean, blnMoveNext As Boolean, blnInsure As Boolean, blnFirstAID As Boolean, bln�ϴ��ɹ� As Boolean
    Dim rsExse As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsInsure As New ADODB.Recordset
    Dim rsPhysic As New ADODB.Recordset
    Dim strҽ����� As String, strҽ������ As String, str��Ŀ���� As String
    Dim str��� As String, str���� As String, str���� As String
    Dim strҽ����� As String, arrҽ����� As Variant, lng������ As Long, lngRecords As Long
    '----------ֻҪ��һ�����˵ĵ����ϴ��ɹ����򷵻��������ݱ��棬δ�ϴ��ɹ�����סԺ������㴦���ϴ�----------
    
    '�ϴ�������ϸ
    On Error GoTo errHand
    Call DebugTool("��������ϴ�")
    gstrSQL = " Select A.ID,A.����ID,A.��ҳID,A.NO,A.���,A.��¼����,A.��¼״̬,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') �Ǽ�ʱ��,A.�շ����," & _
              " A.������ ҽ��,B.���� ��������,A.�շ�ϸĿID,A.���㵥λ,C.��Ŀ���� ҽ����Ŀ���� ,A.ʵ�ս�� ���,A.����*Nvl(A.����,1) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�" & _
              " From סԺ���ü�¼ A,���ű� B,(Select * From ����֧����Ŀ Where ����=" & TYPE_������ & ") C " & _
              " Where A.��¼����=[1] And A.��¼״̬=[2] And A.NO=[3]" & _
              " And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0" & _
              " Order by A.NO,A.����ID"
    Set rsExse = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", int����, int״̬, str���ݺ�)
    
    '�ȼ���Ƿ����δ�������ϸ
    With rsExse
        Do While Not .EOF
            If IsNull(!ҽ����Ŀ����) Then
                MsgBox "��" & !��� & "�еļ�¼δ����ҽ����Ŀ���룡", vbInformation, gstrSysName
                Exit Function
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    '��鱾�����Ƿ��Ǽ�����ҩ�����ڽ�����������Ƿ񼱾���ҩ����
    '��Ϊ����ʱ�ǳ����Զ�ѡ��������ϸ���г��������Գ�����¼���ع���ҩ��־
    blnFirstAID = False
    If FirstAid Then
        blnFirstAID = (MsgBox("���ŵ����Ǽ�����ҩ�𣿣�����ǣ��򱾵���������ϸ������Ϊ������ҩ�ϴ���ҽ�����ģ�", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
    End If
    
    blnUp = False
    With rsExse
        '��ȡ��Ӧ��ҽ���������к�:20031231:���
        '--------------------------------------------------------------------------------------------
        strҽ����� = ""
        Do While Not .EOF
            If Nvl(!���, 0) < 0 Then
                If lng����ID <> !����ID Then
                    '��鱾���Ƿ���ҽ�������Ժ
                    gstrSQL = "Select Count(*) Records From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.����ID=[1] And A.��ҳID=B.סԺ���� And A.����=[2]"
                    Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�ҽ������", CLng(!����ID), TYPE_������)
                    blnInsure = (rsInsure!Records = 1)
                    If blnInsure Then
                        blnInsure = ��ȡ���˻�����Ϣ(!����ID, False)
                        If blnInsure Then lng����ID = !����ID
                    End If
                End If
                
                If blnInsure Then
                    '200312:���:��Ϊ���µ������븺��,���û��ԭʼ���ü�¼,���⴦��
                    '�˷�ʱ,��Hos_Serial��ΪHIS����˷Ѽ�¼��ID,�����Ǳ��˷Ѽ�¼��ID
                    If !�շ���� = 5 Or !�շ���� = 6 Or !�շ���� = 7 Then
                        gstrSQL = "Select Decode(trim(��ʶ��),NULL,����,'',����,��ʶ��) ���� From ҩƷĿ¼ Where ҩƷID=[1]"
                    Else
                        gstrSQL = "Select Decode(trim(��ʶ����),NULL,����,'',����,��ʶ����) ���� From �շ�ϸĿ Where ID=[1]"
                    End If
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡԭʼ��¼�ķ���ID", CLng(!�շ�ϸĿID))
                    strҽ����� = strҽ����� & IIf(strҽ����� = "", "", "|") & GetInsureSerial2(rsTemp!����, Nvl(!���, 0), IIf(strҽ����� = "", True, False))
                End If
            End If
            .MoveNext
        Loop
        
        '���ϴ�����
        '------------------------------------------------------------------------------------------
        '20031231:���
        lng������ = 1: lng����ID = 0
        If strҽ����� <> "" Then
            arrҽ����� = Split(strҽ�����, "|")
        End If
        If .RecordCount <> 0 Then .MoveFirst
        
        Do While Not .EOF
            '������ͷ
            If lng����ID <> !����ID Or strNO <> !NO Then
                Call DebugTool("������ͷ")
                If lng����ID <> 0 And blnInsure Then
                    '�ϴ�
                    blnUp = False
                    blnMoveNext = False
                    If Not ���ýӿ�_ִ��_������() Then
                        �����Ǽ�_������ = bln�ϴ��ɹ�
                        Exit Function
                    End If
                    bln�ϴ��ɹ� = True
                End If
                If lng����ID <> !����ID Then
                    '��鱾���Ƿ���ҽ�������Ժ
                    gstrSQL = "Select Count(*) Records From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.����ID=[1] And A.��ҳID=B.סԺ���� And A.����=[2]"
                    Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�ҽ������", CLng(!����ID), TYPE_������)
                    blnInsure = (rsInsure!Records = 1)
                    If blnInsure Then
                        blnInsure = ��ȡ���˻�����Ϣ(!����ID, False)
                        If blnInsure Then lng����ID = !����ID
                    End If
                End If
                If blnInsure Then
                    strNO = !NO
    
                    'ȡҽ�����������
                    strҽ����� = "": strҽ������ = Nvl(rsExse!ҽ��)
                    If strҽ������ <> "" Then
                        gstrSQL = "Select ���,���� From ��Ա�� Where ����=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����", strҽ������)
                        strҽ����� = rsTemp!���
                    End If
                    
                    'д��ڲ���
                    gstrField_������ = "hospital_id||indi_id||busi_type||serial_no||ordinal_no||input_staff||input_man||recipe_no||doctor_no||doctor_name"
                    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.���˱�� & "||" & _
                                    gCominfo_������.ҵ������ & "||" & gCominfo_������.ҵ�����к� & "||||" & _
                                    gCominfo_������.����Ա���� & "||" & gstrUserName & "||" & strNO & "||" & strҽ����� & "||" & strҽ������
                    If Not ���ýӿ�_׼��_������(Function_������.סԺ����_�ϴ���ϸ) Then
                        �����Ǽ�_������ = bln�ϴ��ɹ�
                        Exit Function
                    End If
                    If Not ���ýӿ�_д��ڲ���_������(1) Then
                        �����Ǽ�_������ = bln�ϴ��ɹ�
                        Exit Function
                    End If
                    blnUp = True
                    
                    'ָ����¼��
                    If Not ���ýӿ�_ָ����¼��_������("FeeInfo") Then
                        �����Ǽ�_������ = bln�ϴ��ɹ�
                        Exit Function
                    End If
                    lngRecords = 1
                End If
            End If
            
            '��������ϸ
            If blnInsure Then
                Call DebugTool("��������ϸ")
                gstrSQL = "Select ��ʶ����,����,����,��� From �շ�ϸĿ Where ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ�ϸĿ�����Ϣ", CLng(!�շ�ϸĿID))
                str���� = ""
                str��� = Nvl(rsTemp!���)
                If InStr(1, str���, "��") <> 0 Then
                    str���� = ToVarchar(Trim(Split(str���, "��")(1)), 50)
                    str��� = ToVarchar(Trim(Split(str���, "��")(0)), 30)
                Else
                    str��� = ToVarchar(Trim(str���), 30)
                End If
                
                str���� = ""
                str��Ŀ���� = ""
                str�������� = Format(!�Ǽ�ʱ��, "yyyy-MM-dd")
                
                Call DebugTool("ȡ��ʶ��")
                If !�շ���� = 5 Or !�շ���� = 6 Or !�շ���� = 7 Then
                    gstrSQL = " Select C.����,C.��ʶ��,B.���� ���� From ҩƷ��Ϣ A,ҩƷ���� B,ҩƷĿ¼ C " & _
                              " Where A.����=B.���� And A.ҩ��ID=C.ҩ��ID And C.ҩƷID = [1]"
                    Set rsPhysic = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����", CLng(!�շ�ϸĿID))
                    Call DebugTool("����:" & Nvl(rsPhysic!����) & "|��ʶ��:" & Nvl(rsPhysic!��ʶ��) & "|����:" & Nvl(rsPhysic!����))
                    str���� = Nvl(rsPhysic!����)
                    str��Ŀ���� = Nvl(rsPhysic!��ʶ��)
                    If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsPhysic!����)
                Else
                    str��Ŀ���� = Nvl(rsTemp!��ʶ����)
                    If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsTemp!����)
                End If
                
                Call DebugTool("��ʶ���ȡ�ɹ�")
                gstrField_������ = "medi_item_type||his_item_code||his_item_name||model||factory||standard||" & _
                            "fee_date||unit||price||dosage||money||usage_flag||usage_days||opp_serial_fee||hos_serial"
                
                
                If Nvl(!���, 0) < 0 Then
                    '20031231:���:��������Ҫ�ϴ�opp_serial_fee
                    Call DebugTool("�ϴ�������")
                    gstrValue_������ = IIf(!�շ���� = 5, "1", IIf(!�շ���� = 6, "2", IIf(!�շ���� = 7, "3", "0"))) & "||" & _
                                str��Ŀ���� & "||" & Nvl(rsTemp!����) & "||" & str���� & "||" & str���� & "||" & str��� & "||" & _
                                str�������� & "||" & Nvl(!���㵥λ) & "||" & Format(!��� / !����, "#####0.0000;-#####0.0000;0;") & "||" & !���� & "||" & !��� & "||0||||" & _
                                arrҽ�����(lng������ - 1) & "||" & !ID
                Else
                    Call DebugTool("�ϴ�������")
                    gstrValue_������ = IIf(!�շ���� = 5, "1", IIf(!�շ���� = 6, "2", IIf(!�շ���� = 7, "3", "0"))) & "||" & _
                                str��Ŀ���� & "||" & Nvl(rsTemp!����) & "||" & str���� & "||" & str���� & "||" & str��� & "||" & _
                                str�������� & "||" & Nvl(!���㵥λ) & "||" & Format(!��� / !����, "#####0.0000;-#####0.0000;0;") & "||" & !���� & "||" & !��� & "||" & IIf(blnFirstAID, "2", "0") & "||||||" & !ID
                End If
                
                If blnMoveNext Then
                    Call DebugTool("�ƶ���¼��")
                    Call ���ýӿ�_�ƶ���¼��_������(MoveNext)
                End If
                Call DebugTool("д��ڲ���")
                If Not ���ýӿ�_д��ڲ���_������(lngRecords) Then
                    �����Ǽ�_������ = bln�ϴ��ɹ�
                    Exit Function
                End If
                lngRecords = lngRecords + 1
                '�����ϴ���־����Ϊ��ϸ������ȷ�ϴ��󣬲��ܱ�֤�������ȷ��
                'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                Call DebugTool("���ϴ���־")
                gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsExse("NO") & "'," & rsExse("���") & "," & rsExse("��¼����") & "," & rsExse("��¼״̬") & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
                
                blnMoveNext = True
                lng������ = lng������ + 1
            End If
            .MoveNext
        Loop
        If blnUp And blnInsure Then
            If Not ���ýӿ�_ִ��_������() Then
                �����Ǽ�_������ = bln�ϴ��ɹ�
                Exit Function
            End If
            bln�ϴ��ɹ� = True
        End If
    End With
    
    Call DebugTool("�ϴ��ɹ�")
    �����Ǽ�_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    �����Ǽ�_������ = bln�ϴ��ɹ�
End Function

Public Function ��������_������(ByVal int���� As Integer, ByVal int״̬ As Integer, ByVal str���ݺ� As String) As Boolean
    '��Ҫ�ҵ�ԭʼ�Ǳʼ�¼�ķ���ID�����ܹ���������
    Dim lng����ID As Long, strNO As String, str�������� As String, intԭʼ��¼״̬ As Integer
    Dim str���� As String, str���� As String, lngԭʼ����ID As Long, int��� As Integer, lng������ As Long
    Dim blnUp As Boolean, blnMoveNext As Boolean, blnInsure As Boolean, bln�ϴ��ɹ� As Boolean
    Dim strҽ����� As String, arrҽ�����
    Dim strҽ����� As String, strҽ������ As String, str��Ŀ���� As String
    Dim str��� As String, str���� As String, str���� As String
    
    Dim rsExse As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsInsure As New ADODB.Recordset
    Dim rsPhysic As New ADODB.Recordset
    Dim cnUpData As New ADODB.Connection
    '�ϴ�������ϸ
    On Error GoTo errHand
    
    intԭʼ��¼״̬ = 3
    gstrSQL = " Select A.ID,A.����ID,A.NO,A.���,A.��¼����,A.��¼״̬,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') �Ǽ�ʱ��,A.�շ����," & _
              " A.������ ҽ��,B.���� ��������,A.�շ�ϸĿID,A.���㵥λ,C.��Ŀ���� ҽ����Ŀ���� ,A.ʵ�ս�� ���,A.����*Nvl(A.����,1) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�" & _
              " From סԺ���ü�¼ A,���ű� B,(Select * From ����֧����Ŀ Where ����=[4]) C " & _
              " Where A.��¼����=[1] And A.��¼״̬=[2] And A.NO=[3]" & _
              " And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0" & _
              " Order by A.NO,A.����ID"
    Set rsExse = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", int����, int״̬, str���ݺ�, TYPE_������)
    
    With rsExse
        '��ȡ��Ӧ��ҽ���������к�
        strҽ����� = ""
        Do While Not .EOF
            If lng����ID <> !����ID Then
                '��鱾���Ƿ���ҽ�������Ժ
                gstrSQL = "Select Count(*) Records From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.����ID=[1] And A.��ҳID=B.סԺ���� And A.����=[2]"
                Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�ҽ������", CLng(!����ID), TYPE_������)
                 blnInsure = (rsInsure!Records = 1)
                If blnInsure Then
                    blnInsure = ��ȡ���˻�����Ϣ(!����ID, False)
                    If blnInsure Then lng����ID = !����ID
                End If
            End If
            
            If blnInsure Then
                gstrSQL = " Select ID From סԺ���ü�¼" & _
                          " Where ��¼����=[1] And ��¼״̬=[2] And NO=[3] And ���=[4]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡԭʼ��¼�ķ���ID", int����, intԭʼ��¼״̬, str���ݺ�, CStr(!���))
                lngԭʼ����ID = rsTemp!ID
                
                '20040105:���:
                '�˷�ʱ,��Hos_Serial��ΪHIS����˷Ѽ�¼��ID,�����Ǳ��˷Ѽ�¼��ID,��Ϊ��������ʱ�޶�Ӧԭʼ��¼
                strҽ����� = strҽ����� & IIf(strҽ����� = "", "", "|") & GetInsureSerial(lngԭʼ����ID, IIf(strҽ����� = "", True, False))
            End If
            .MoveNext
        Loop
        
        '���ϴ�����
        blnUp = False
        lng����ID = 0: strNO = ""
        arrҽ����� = Split(strҽ�����, "|")
        lng������ = 1
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '������ͷ
            If lng����ID <> !����ID Or strNO <> !NO Then
                If lng����ID <> 0 And blnInsure Then
                    '�ϴ�
                    blnUp = False
                    blnMoveNext = False
                    If Not ���ýӿ�_ִ��_������() Then
                        ��������_������ = bln�ϴ��ɹ�
                        Exit Function
                    End If
                    bln�ϴ��ɹ� = True
                End If
                strNO = !NO
                If lng����ID <> !����ID Then
                    '��鱾���Ƿ���ҽ�������Ժ
                    gstrSQL = "Select Count(*) Records From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.����ID=[1] And A.��ҳID=B.סԺ���� And A.����=[2]"
                    Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�ҽ������", CLng(!����ID), TYPE_������)
                    blnInsure = (rsInsure!Records = 1)
                    If blnInsure Then
                        blnInsure = ��ȡ���˻�����Ϣ(!����ID, False)
                        If blnInsure Then lng����ID = !����ID
                    End If
                End If
                
                If blnInsure Then
    
                    'ȡҽ�����������
                    strҽ����� = "": strҽ������ = Nvl(rsExse!ҽ��)
                    If strҽ������ <> "" Then
                        gstrSQL = "Select ���,���� From ��Ա�� Where ����=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����", strҽ������)
                        strҽ����� = rsTemp!���
                    End If
                    
                    'д��ڲ���
                    gstrField_������ = "hospital_id||indi_id||busi_type||serial_no||ordinal_no||input_staff||input_man||recipe_no||doctor_no||doctor_name"
                    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.���˱�� & "||" & _
                                    gCominfo_������.ҵ������ & "||" & gCominfo_������.ҵ�����к� & "||||" & _
                                    gCominfo_������.����Ա���� & "||" & gstrUserName & "||" & strNO & "||" & strҽ����� & "||" & strҽ������
                    If Not ���ýӿ�_׼��_������(Function_������.סԺ����_�ϴ���ϸ) Then
                        ��������_������ = bln�ϴ��ɹ�
                        Exit Function
                    End If
                    If Not ���ýӿ�_д��ڲ���_������(1) Then
                        ��������_������ = bln�ϴ��ɹ�
                        Exit Function
                    End If
                    blnUp = True
                    
                    'ָ����¼��
                    If Not ���ýӿ�_ָ����¼��_������("FeeInfo") Then
                        ��������_������ = bln�ϴ��ɹ�
                        Exit Function
                    End If
                End If
            End If
            
            If blnInsure Then
                '��ȡ�����Ϣ
                int��� = !���
                
                gstrSQL = "Select ��ʶ����,����,����,��� From �շ�ϸĿ Where ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ�ϸĿ�����Ϣ", CLng(!�շ�ϸĿID))
                str���� = rsTemp!����
                str���� = rsTemp!����
                str��� = Nvl(rsTemp!���)
                str���� = ""
                If InStr(1, str���, "��") <> 0 Then
                    str���� = ToVarchar(Trim(Split(str���, "��")(1)), 50)
                    str��� = ToVarchar(Trim(Split(str���, "��")(0)), 30)
                Else
                    str��� = ToVarchar(Trim(str���), 30)
                End If
                
                '��������ϸ
                str���� = ""
                str��Ŀ���� = ""
                str�������� = "" '�ķ�ʱ��������Ϊ��
                
                If !�շ���� = 5 Or !�շ���� = 6 Or !�շ���� = 7 Then
                    gstrSQL = " Select C.����,C.��ʶ��,B.���� ���� From ҩƷ��Ϣ A,ҩƷ���� B,ҩƷĿ¼ C " & _
                              " Where A.����=B.���� And A.ҩ��ID=C.ҩ��ID And C.ҩƷID =[1]"
                    Set rsPhysic = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����", CLng(!�շ�ϸĿID))
                    str���� = Nvl(rsPhysic!����)
                    str��Ŀ���� = Nvl(rsPhysic!��ʶ��)
                    If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsPhysic!����)
                Else
                    str��Ŀ���� = Nvl(rsTemp!��ʶ����)
                    If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsTemp!����)
                End If
                
                gstrField_������ = "medi_item_type||his_item_code||his_item_name||model||factory||standard||" & _
                            "fee_date||unit||price||dosage||money||usage_flag||usage_days||opp_serial_fee||hos_serial"
                gstrValue_������ = IIf(!�շ���� = 5, "1", IIf(!�շ���� = 6, "2", IIf(!�շ���� = 7, "3", "0"))) & "||" & _
                            str��Ŀ���� & "||" & str���� & "||" & str���� & "||" & str���� & "||" & str��� & "||" & _
                            str�������� & "||" & Nvl(!���㵥λ) & "||" & Format(!��� / !����, "#####0.0000;-#####0.0000;0;") & "||" & !���� & "||" & !��� & "||0||||" & _
                            arrҽ�����(lng������ - 1) & "||" & !ID
                
                If blnMoveNext Then Call ���ýӿ�_�ƶ���¼��_������(MoveNext)
                If Not ���ýӿ�_д��ڲ���_������(.AbsolutePosition) Then
                    ��������_������ = bln�ϴ��ɹ�
                    Exit Function
                End If
                'Call ErrInformation
                '�����ϴ���־����Ϊ��ϸ������ȷ�ϴ��󣬲��ܱ�֤�������ȷ��
                'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsExse("NO") & "'," & rsExse("���") & "," & rsExse("��¼����") & "," & rsExse("��¼״̬") & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
                
                lng������ = lng������ + 1
                blnMoveNext = True
            End If
            .MoveNext
        Loop
        If blnUp And blnInsure Then
            If Not ���ýӿ�_ִ��_������() Then
                ��������_������ = bln�ϴ��ɹ�
                Exit Function
            End If
            bln�ϴ��ɹ� = True
        End If
    End With
    
    ��������_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    ��������_������ = bln�ϴ��ɹ�
End Function

Public Function ����ķ�_������(ByVal int���� As Integer, ByVal int״̬ As Integer, ByVal str���ݺ� As String) As Boolean
    '��Ҫ�ҵ�ԭʼ�Ǳʼ�¼�ķ���ID�����ܹ���������
    Dim lng����ID As Long, strNO As String, str�������� As String, intԭʼ��¼״̬ As Integer
    Dim str���� As String, str���� As String, lngԭʼ����ID As Long, int��� As Integer, lng������ As Long
    Dim blnUp As Boolean, blnMoveNext As Boolean, blnInsure As Boolean
    Dim strҽ����� As String, arrҽ�����
    Dim strҽ����� As String, strҽ������ As String, str��Ŀ���� As String
    Dim str��� As String, str���� As String, str���� As String
    Dim rsExse As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsInsure As New ADODB.Recordset
    Dim rsPhysic As New ADODB.Recordset
    '�ϴ�������ϸ
    On Error GoTo errHand
    
    intԭʼ��¼״̬ = 3
    gstrSQL = " Select A.ID,A.����ID,A.NO,A.���,A.��¼����,A.��¼״̬,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') �Ǽ�ʱ��,A.�շ����," & _
              " A.������ ҽ��,B.���� ��������,A.�շ�ϸĿID,A.���㵥λ,C.��Ŀ���� ҽ����Ŀ���� ,A.ʵ�ս�� ���,A.����*Nvl(A.����,1) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�" & _
              " From ������ü�¼ A,���ű� B,(Select * From ����֧����Ŀ Where ����=[4]) C " & _
              " Where A.��¼����=[1] And A.��¼״̬=[2] And A.NO=[3]" & _
              " And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0" & _
              " Order by A.NO,A.����ID"
    Set rsExse = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", int����, int״̬, str���ݺ�, TYPE_������)
    Call DebugTool("ȡ�ó�����¼����" & rsExse.RecordCount)
    
    With rsExse
        '��ȡ��Ӧ��ҽ���������к�
        strҽ����� = ""
        Do While Not .EOF
            gstrSQL = " Select ID From ������ü�¼" & _
                      " Where ��¼����=[1] And ��¼״̬=[2] And NO=[3] And ���=[4]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡԭʼ��¼�ķ���ID", int����, intԭʼ��¼״̬, str���ݺ�, CLng(!���))
            lngԭʼ����ID = rsTemp!ID
            
            '20040105:���:
            '�˷�ʱ,��Hos_Serial��ΪHIS����˷Ѽ�¼��ID,�����Ǳ��˷Ѽ�¼��ID,��Ϊ��������ʱ�޶�Ӧԭʼ��¼
            strҽ����� = strҽ����� & IIf(strҽ����� = "", "", "|") & GetInsureSerial_OutExse(lngԭʼ����ID, IIf(strҽ����� = "", True, False))
            .MoveNext
        Loop
        Call DebugTool("�˷����кţ�" & strҽ�����)
        
        '���ϴ�����
        blnUp = False
        lng����ID = 0: strNO = ""
        arrҽ����� = Split(strҽ�����, "|")
        lng������ = 1
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '������ͷ
            If lng����ID <> !����ID Then
                lng����ID = !����ID
                'ȡҽ�����������
                strҽ����� = "": strҽ������ = Nvl(rsExse!ҽ��)
                If strҽ������ <> "" Then
                    gstrSQL = "Select ���,���� From ��Ա�� Where ����=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����", strҽ������)
                    strҽ����� = rsTemp!���
                End If
                
                'д��ڲ���
'                    1   hospital_id    ҽ�ƻ�������    20  ��
'                    2   indi_id        ���˱��    8   ��
'                    3   busi_type      ҵ������    2   ��  "11"������
'                    4   serial_no      ҵ�����к�  12  ��
'                    5   ic_flag        �ÿ���־    1   ��  "0"����ʹ��IC��                    "1"��ʹ��IC��
'                    6   Reg_staff      �Ǽ���Ա����    5   ��
'                    7   Reg_man        �Ǽ�������  10  ��
'                    8   begin_date     ����ʱ��        ��  ��ʽ��YYYY-MM-DD HH:MI:SS(24Сʱ)
'                    9   calcSaveFlag   ���㱣���־    1   ��  "0"������                    "1"���շ�
'                    10  accMoney       �����ʻ�֧�����    18  ��
'                    11  recipe_no      ������  20  ��
                gstrField_������ = "hospital_id||indi_id||busi_type||serial_no||ic_flag||Reg_staff||Reg_man||begin_date||calcSaveFlag||accMoney||recipe_no"
                gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.���˱�� & "||" & _
                                gCominfo_������.ҵ������ & "||" & gCominfo_������.ҵ�����к� & "||1||" & _
                                gCominfo_������.����Ա���� & "||" & gstrUserName & "||" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "||" & _
                                "1||0||" & str���ݺ�
                If Not ���ýӿ�_׼��_������(IIf(gCominfo_������.ҵ������ = ҵ�����_������.����涨��, ����涨��_�շ�, ��ͨ����_�շ�)) Then Exit Function
                If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
                Call DebugTool("�����ӿ�")
                
                'ָ����¼��
                If Not ���ýӿ�_ָ����¼��_������("FeeInfo") Then Exit Function
            End If
            
            '��ȡ�����Ϣ
            int��� = !���
            
            gstrSQL = "Select ��ʶ����,����,����,��� From �շ�ϸĿ Where ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ�ϸĿ�����Ϣ", CLng(!�շ�ϸĿID))
            str���� = rsTemp!����
            str���� = rsTemp!����
            str��� = Nvl(rsTemp!���)
            str���� = ""
            If InStr(1, str���, "��") <> 0 Then
                str���� = ToVarchar(Trim(Split(str���, "��")(1)), 50)
                str��� = ToVarchar(Trim(Split(str���, "��")(0)), 30)
            Else
                str��� = ToVarchar(Trim(str���), 30)
            End If
            
            '��������ϸ
            str���� = ""
            str��Ŀ���� = ""
            str�������� = "" '�ķ�ʱ��������Ϊ��
            
            If !�շ���� = 5 Or !�շ���� = 6 Or !�շ���� = 7 Then
                gstrSQL = " Select C.����,C.��ʶ��,B.���� ���� From ҩƷ��Ϣ A,ҩƷ���� B,ҩƷĿ¼ C " & _
                          " Where A.����=B.���� And A.ҩ��ID=C.ҩ��ID And C.ҩƷID =[1]"
                Set rsPhysic = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����", CLng(!�շ�ϸĿID))
                str���� = Nvl(rsPhysic!����)
                str��Ŀ���� = Nvl(rsPhysic!��ʶ��)
                If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsPhysic!����)
            Else
                str��Ŀ���� = Nvl(rsTemp!��ʶ����)
                If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsTemp!����)
            End If
            
            Call DebugTool("׼��������ϸ")
            gstrField_������ = "medi_item_type||his_item_code||his_item_name||model||factory||" & _
                        "standard||fee_date||unit||price||dosage||money||opp_serial_fee||hos_serial"
            gstrValue_������ = IIf(!�շ���� = 5, "1", IIf(!�շ���� = 6, "2", IIf(!�շ���� = 7, "3", "0"))) & "||" & _
                        str��Ŀ���� & "||" & str���� & "||" & str���� & "||" & str���� & "||" & str��� & "||" & _
                        str�������� & "||" & Nvl(!���㵥λ) & "||" & Format(!��� / !����, "#####0.0000;-#####0.0000;0;") & "||" & !���� & "||" & !��� & "||" & arrҽ�����(lng������ - 1) & "||" & !ID
            
            If blnMoveNext Then Call ���ýӿ�_�ƶ���¼��_������(MoveNext)
            If Not ���ýӿ�_д��ڲ���_������(.AbsolutePosition) Then Exit Function
            
            Call DebugTool("׼�����ϴ���־")
            'Call ErrInformation
            '�����ϴ���־����Ϊ��ϸ������ȷ�ϴ��󣬲��ܱ�֤�������ȷ��
            'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
            gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsExse("NO") & "'," & rsExse("���") & "," & rsExse("��¼����") & "," & rsExse("��¼״̬") & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
            
            lng������ = lng������ + 1
            blnMoveNext = True
            .MoveNext
        Loop
        If Not ���ýӿ�_ִ��_������() Then Exit Function
    End With
    
    Call DebugTool("�Ըķѵķ�ʽ��ɳ���ԭʼ����ҵ��Ĺ���")
    ����ķ�_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function GetInsureSerial(ByVal lng�������к� As Long, Optional ByVal blnִ�� As Boolean = False) As Long
    Dim str���к� As String
    '����ҽ���������кţ���ҽ���Լ��ķ������к�
'    1   hospital_id    ҽ�ƻ�������   20  ��
'    2   serial_no      ҵ�����к�     12  ��
'    3   calc_flag      �����־       1
    If blnִ�� Then
        gstrField_������ = "hospital_id||serial_no||calc_flag"
        gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.ҵ�����к� & "||1"
        If Not ���ýӿ�_׼��_������(Function_������.סԺ����_����ϸ) Then Exit Function
        If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
        If Not ���ýӿ�_ִ��_������ Then Exit Function
        If Not ���ýӿ�_ָ����¼��_������("calc_fee_info") Then Exit Function
    End If
    
    If ���ýӿ�_��¼��_������ Then
        Call ���ýӿ�_�ƶ���¼��_������(MoveFirst)
        Do While True
            Call ���ýӿ�_��ȡ����_������("hos_serial", str���к�)
            If Val(str���к�) = lng�������к� Then
                Call ���ýӿ�_��ȡ����_������("serial_fee", str���к�)
                GetInsureSerial = Val(str���к�)
                Exit Function
            End If
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Function
        Loop
    End If
End Function

Private Function GetInsureSerial_OutExse(ByVal lng�������к� As Long, Optional ByVal blnִ�� As Boolean = False) As Long
    Dim str���к� As String
    '����ҽ���������кţ���ҽ���Լ��ķ������к�
'    1   hospital_id    ҽ�ƻ�������   20  ��
'    2   serial_no      ҵ�����к�     12  ��
'    3   calc_flag      �����־       1
    If blnִ�� Then
        gstrField_������ = "hospital_id||serial_no"
        gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.ҵ�����к�
        If Not ���ýӿ�_׼��_������(Function_������.����_����ϸ) Then Exit Function
        If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
        If Not ���ýӿ�_ִ��_������ Then Exit Function
        If Not ���ýӿ�_ָ����¼��_������("FeeInfo") Then Exit Function
    End If
    
    If ���ýӿ�_��¼��_������ Then
        Call ���ýӿ�_�ƶ���¼��_������(MoveFirst)
        Do While True
            Call ���ýӿ�_��ȡ����_������("hos_serial", str���к�)
            Call DebugTool("ҽ���б����HIS����ID��" & Val(str���к�) & "�����˵�ԭʼ��¼�ķ���ID��" & lng�������к�)
            If Val(str���к�) = lng�������к� Then
                Call ���ýӿ�_��ȡ����_������("serial_fee", str���к�)
                Call DebugTool("�ҵ��ˣ�" & Val(str���к�))
                GetInsureSerial_OutExse = Val(str���к�)
                Exit Function
            End If
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then
                Call DebugTool("û���ҵ�")
                Exit Function
            End If
        Loop
    End If
End Function

Private Function GetInsureSerial2(ByVal str���� As String, cur��� As Currency, Optional ByVal blnִ�� As Boolean = False) As Long
'���ܣ�����ҽ���˷������кţ�������ֱ�����븺�����������
'������str����=HIS���շ�ϸĿ����
'      cur���=��ǰ�˷ѽ��
    Dim str���к� As String, str��Ӧ���к� As String, curMoney As Currency
    Dim arr�˷Ѽ�() As Variant
    Dim rsԭʼ�� As New ADODB.Recordset
    Dim strFields As String, strValues As String
    Dim strTemp As String, i As Long, j As Long
    
'    1   hospital_id    ҽ�ƻ�������   20  ��
'    2   serial_no      ҵ�����к�     12  ��
'    3   calc_flag      �����־       1
    If blnִ�� Then
        gstrField_������ = "hospital_id||serial_no||calc_flag"
        gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.ҵ�����к� & "||1"
        If Not ���ýӿ�_׼��_������(Function_������.סԺ����_����ϸ) Then Exit Function
        If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
        If Not ���ýӿ�_ִ��_������ Then Exit Function
        If Not ���ýӿ�_ָ����¼��_������("calc_fee_info") Then Exit Function
    End If
    
    If ���ýӿ�_��¼��_������ Then
        '��ʼ��ԭʼ��¼��
        strFields = "���к�" & "," & adDouble & "," & "20" & "|" & _
                    "���" & "," & adDouble & "," & "20"
        Call Record_Init(rsԭʼ��, strFields)
        
        arr�˷Ѽ� = Array()
        Call ���ýӿ�_�ƶ���¼��_������(MoveFirst)
        Do While True
            '20040105:���:�����·���,�����ϸ��ӦҪ�˷ѵ�ԭʼ��¼
            '��ͬһ��Ŀ(������ͬ),����㹻�˵�ԭʼ��¼
            Call ���ýӿ�_��ȡ����_������("his_item_code", strTemp)
            If strTemp = str���� Then
                Call ���ýӿ�_��ȡ����_������("serial_fee", strTemp)
                str���к� = strTemp
                Call ���ýӿ�_��ȡ����_������("opp_serial_fee", strTemp)
                str��Ӧ���к� = strTemp
                Call ���ýӿ�_��ȡ����_������("money", strTemp)
                curMoney = Val(strTemp)
                                
                '����㹻�˵��Ҳ����˷Ѽ�¼�ſ���Ϊ���˵�
                If curMoney >= Abs(cur���) And Val(str��Ӧ���к�) = 0 Then
                    strFields = "���к�|���"
                    strValues = Val(str���к�) & "|" & curMoney
                    Call Record_Add(rsԭʼ��, strFields, strValues)
                End If
                
                '��Ӧ���кŲ�Ϊ�յĲ����˷ѵļ�¼
                If Val(str��Ӧ���к�) <> 0 Then
                    ReDim Preserve arr�˷Ѽ�(UBound(arr�˷Ѽ�) + 1)
                    arr�˷Ѽ�(UBound(arr�˷Ѽ�)) = str��Ӧ���к�
                End If
                
            End If
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
        
        'ֻ����δ�˹��ѵļ�¼�����˷ѣ�����һ�ʽ����ӽ��ļ�¼��Ϊ�˷�ԭʼ��¼��
        With rsԭʼ��
            If .RecordCount <> 0 Then .Sort = "��� asc"
            Do While Not .EOF
                For j = 0 To UBound(arr�˷Ѽ�)
                    If Val(arr�˷Ѽ�(j)) = !���к� Then
                        Exit For
                    End If
                Next
                If j > UBound(arr�˷Ѽ�) Then
                    'û���κ�һ���˷Ѽ�¼�Ķ�Ӧ���кź�ԭʼ��¼�е�ǰ��¼�����к���ͬ
                    GetInsureSerial2 = !���к�
                    Exit Function
                End If
                .MoveNext
            Loop
        End With
    End If
End Function

Public Function GetBalance(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal strҵ�����к� As String, ByVal strҽԺ���� As String) As Boolean
'���ܣ�����ҵ�����кŻ�ȡ��������������ʱ���У���Ʊ��ӡ
'    1   hospital_id ҽ�ƻ�������               20  ��
'    2   serial_no   ҵ�����к�                 12  ��
'    3   fee_flag    �Ƿ�ȡ�ô�ҵ��ķ�����ϸ   1   ��  0����ȡ��ϸ��1��ȡ��ϸ
    
    Dim int���� As Integer, cur��� As Currency, blnDelete As Boolean, strPara As String
    Dim str���� As String, str���� As String, int��� As Integer '20031228:���
    Dim cur�Էѷ��� As Currency, cur�Ը����� As Currency
        
    int���� = 1
    blnDelete = True
    gstrField_������ = "hospital_id||serial_no||fee_flag"
    gstrValue_������ = strҽԺ���� & "||" & strҵ�����к� & "||0"
    If Not ���ýӿ�_׼��_������(Function_������.����_��ȡ��Ʊ��Ϣ) Then Exit Function
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
    If Not ���ýӿ�_ִ��_������ Then Exit Function
    'InvoiceInfo,����ҵ��Ļ���֧�����
    If Not ���ýӿ�_ָ����¼��_������("InvoiceInfo") Then Exit Function
    If ���ýӿ�_��¼��_������ Then
        Call ���ýӿ�_�ƶ���¼��_������(MoveFirst)
        Do While True
            '20031228:���
            '�������
            If Not ���ýӿ�_��ȡ����_������("Fund_id", strPara) Then Exit Function
            str���� = Trim(strPara)
            
            '��������
            If Not ���ýӿ�_��ȡ����_������("Fund_name", strPara) Then Exit Function
            str���� = Trim(strPara)
            
            '������
            If Not ���ýӿ�_��ȡ����_������("real_pay", strPara) Then Exit Function
            cur��� = Val(strPara)
            
            '�ϼƱ��:0-���ϼƵ�,1-�ϼƵ�
            If Not ���ýӿ�_��ȡ����_������("sum_flag", strPara) Then Exit Function
            int��� = Val(strPara)
            
            '����
            gstrSQL = "ZL_��Ʊ��Ϣ_INSERT(" & lng����ID & "," & lng����ID & "," & int���� & "," & _
                "'" & str���� & "','" & str���� & "'," & cur��� & "," & int��� & "," & _
                IIf(blnDelete, "1", "0") & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "���뷢Ʊ��Ϣ")
            blnDelete = False
            
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do '20031228:���,��ΪExit Do
        Loop
    End If
    
    '20031228:���:����ѯ��ʱ����,��������
    'FeeSquareInfo,���ý��������Ϣ
    int���� = 0
    If Not ���ýӿ�_ָ����¼��_������("FeeSquareInfo") Then Exit Function
    If ���ýӿ�_��¼��_������ Then
        Call ���ýӿ�_�ƶ���¼��_������(MoveFirst)
        Do While True
            Call ���ýӿ�_��ȡ����_������("stat_name", str����)
            Call ���ýӿ�_��ȡ����_������("zfy", strPara)
            cur��� = Val(strPara)
            Call ���ýӿ�_��ȡ����_������("qzf", strPara)
            cur�Էѷ��� = Val(strPara)
            Call ���ýӿ�_��ȡ����_������("blzf", strPara)
            cur�Ը����� = Val(strPara)
            gstrSQL = "ZL_��Ʊ��Ϣ_INSERT(" & lng����ID & "," & lng����ID & "," & int���� & "," & _
                "'" & cur�Էѷ��� & "|" & cur�Ը����� & "','" & str���� & "'," & cur��� & ",NULL," & IIf(blnDelete, "1", "0") & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "���뷢Ʊ��Ϣ")

            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do '20031228:���,��ΪExit Do
        Loop
    End If
    
    GetBalance = True
End Function

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function ��ȡ���㵥_������(ByVal lng����ID As Long, ByVal lng����ID As Long, intҵ������ As Integer, strҵ�����к� As String) As Boolean
    Dim intType As Integer
    Dim strData As String
    Dim strTemp As String
    Dim bln���� As Boolean, blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    '��ȡ���˵Ľ��㵥�����������涨����סԺ�������ͣ�
    On Error GoTo errHand
    
    If Not ҽ����ʼ��_������ Then Exit Function
    
    gstrSQL = " Select ����,֧��˳��� ҵ�����к�,��ע ҵ������ From ���ս����¼ " & _
              " Where ����=[1] And ����ID=[2] And ��¼ID=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ò��˱���ҵ���ҵ�����ͼ�ҵ�����к�", TYPE_������, lng����ID, lng����ID)
    If rsTemp.EOF Then
        MsgBox "��ȡ�ò��˱���ҵ������ʱ�������󣡣�δ�ҵ��������ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    strҵ�����к� = Trim(rsTemp!ҵ�����к�)
    intҵ������ = Val(rsTemp!ҵ������)
    bln���� = (rsTemp!���� = 1)
    
    If strҵ�����к� = "" Then
        MsgBox "ҵ�����к�Ϊ�գ��޷���ȡ�ò��˱��ν��׵���ϸ���ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '1   hospital_id    ҽ�ƻ�������    20  ��
    '2   serial_no      ҵ�����к�      12  ��
    gstrField_������ = "hospital_id||serial_no"
    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & strҵ�����к�
    
    If bln���� Then
        If intҵ������ = ҵ�����_������.����涨�� Then
            intType = 2
            If Not ���ýӿ�_׼��_������(Function_������.���㵥_����涨��) Then Exit Function
        Else
            intType = 1
            If Not ���ýӿ�_׼��_������(Function_������.���㵥_����) Then Exit Function
        End If
    Else
        intType = 3
        If Not ���ýӿ�_׼��_������(Function_������.���㵥_סԺ) Then Exit Function
    End If
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
    If Not ���ýӿ�_ִ��_������ Then Exit Function
    
    Call DebugTool("�������ҵ������")
    gcnOracle.BeginTrans
    'ɾ��������ǰ�Ľ��㵥
    gstrSQL = "Delete From ���㵥_������Ϣ Where ����<Sysdate-1"
    gcnOracle.Execute gstrSQL
    'ɾ����������ͬҵ�����кŵ�����
    gstrSQL = "Delete From ���㵥_������Ϣ Where ҵ�����к�='" & strҵ�����к� & "'"
    gcnOracle.Execute gstrSQL
    
    blnTrans = True
    'ȡ���˻�����Ϣ��ֻ��һ����
    Call DebugTool("ȡ���˻�����Ϣ")
    If Not ���ýӿ�_ָ����¼��_������("Info") Then GoTo ExitSub
    gstrSQL = " Insert Into ���㵥_������Ϣ" & _
              "(ҵ�����к�,����,�Ա�,����,ҽ����,���֤��,��Ա���,�Ƿ����ܹ���Ա����,����Ա�ȼ�����,סԺ��," & _
              " ҽ�ƻ���,ҽ�ƻ����ȼ�,��λ����,�ٴ����,��Ժ����,��Ժ����,סԺ����,��������,��������,����)" & _
              " Values ('" & strҵ�����к� & "'"
    If Not GetPatientSQL(gstrSQL, intType) Then GoTo ExitSub
    Call DebugTool("��ִ�е�SQL��" & gstrSQL)
    gcnOracle.Execute gstrSQL
    
    'ȡ��ʷҵ����Ϣ��ֻ��һ����
    Call DebugTool("ȡ��ʷҵ����Ϣ")
    If Not ���ýӿ�_ָ����¼��_������("His") Then GoTo ExitSub
    If intҵ������ <> 2 Then
        gstrSQL = " Insert Into ���㵥_��ʷ������Ϣ" & _
                  "(ҵ�����к�,סԺ��־,��������ۼ�,��������ۼ�,�����ܷ���,�����ܷ���,����,�����ʻ�,ҽ������,�������,����Ա����,�������)" & _
                  " Values ('" & strҵ�����к� & "'," & IIf(bln����, 1, 3)
    Else
        gstrSQL = " Insert Into ���㵥_��ʷ������Ϣ_����" & _
                  "(ҵ�����к�,�걨����,�ܷ���,����,ҽ������,ȫ�Ը�,�����Ը�,�����Ը�,�������,����Ա����,��������ҵ�����)" & _
                  " Values ('" & strҵ�����к� & "'"
    End If
    If Not GetHistorySQL(gstrSQL, intType) Then GoTo ExitSub
    Call DebugTool("��ִ�е�SQL��" & gstrSQL)
    gcnOracle.Execute gstrSQL
    
    'ȡ�ֶη�����Ϣ��������
    Call DebugTool("ȡ�ֶη�����Ϣ")
    If Not ���ýӿ�_ָ����¼��_������("Seg") Then GoTo ExitSub
    gstrSQL = " Insert Into ���㵥_�������" & _
              "(ҵ�����к�,֧������,�ֽ�,�����ʻ�,ҽ������,����ҽ��,����Ա����)" & _
              " Values ('" & strҵ�����к� & "'"
    If ���ýӿ�_��¼��_������ Then
        Do While True
            strTemp = gstrSQL
            Call ���ýӿ�_��ȡ����_������("policy_type", strData)
            Call CombinateString(strTemp, strData)
            Call ���ýӿ�_��ȡ����_������("cash_pay", strData)
            Call CombinateString(strTemp, strData)
            Call ���ýӿ�_��ȡ����_������("acct_pay", strData)
            Call CombinateString(strTemp, strData)
            Call ���ýӿ�_��ȡ����_������("found_pay", strData)
            Call CombinateString(strTemp, strData)
            Call ���ýӿ�_��ȡ����_������("additional_pay", strData)
            Call CombinateString(strTemp, strData)
            Call ���ýӿ�_��ȡ����_������("official_pay", strData)
            Call CombinateString(strTemp, strData)
            strTemp = strTemp & ")"
            Call DebugTool("��ִ�е�SQL��" & strTemp)
            gcnOracle.Execute strTemp
            
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
    End If
    
    'ȡ���������Ϣ��������
    Call DebugTool("ȡ���������Ϣ")
    If Not ���ýӿ�_ָ����¼��_������("Fee") Then GoTo ExitSub
    gstrSQL = " Insert Into ���㵥_�������" & _
              "(ҵ�����к�,�շ���Ŀ����,�շ���Ŀ����,�ܷ���,ȫ�Է�,�����Ը�)" & _
              " Values ('" & strҵ�����к� & "'"
    If ���ýӿ�_��¼��_������ Then
        Do While True
            strTemp = gstrSQL
            Call ���ýӿ�_��ȡ����_������("stat_type", strData)
            Call CombinateString(strTemp, strData)
            Call ���ýӿ�_��ȡ����_������("stat_name", strData)
            Call CombinateString(strTemp, strData)
            Call ���ýӿ�_��ȡ����_������("zfy", strData)
            Call CombinateString(strTemp, strData)
            Call ���ýӿ�_��ȡ����_������("qzf", strData)
            Call CombinateString(strTemp, strData)
            Call ���ýӿ�_��ȡ����_������("blzf", strData)
            Call CombinateString(strTemp, strData)
            strTemp = strTemp & ")"
            Call DebugTool("��ִ�е�SQL��" & strTemp)
            gcnOracle.Execute strTemp
            
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
    End If
    
    gcnOracle.CommitTrans
    ��ȡ���㵥_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
ExitSub:
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Public Function ��ȡ�����_������() As Boolean
    '��ȡ������������涨�������סԺ�����
    Dim strStart As String, strEnd As String
    Dim blnTrans As Boolean
    On Error GoTo errHand
    If Not ҽ����ʼ��_������ Then Exit Function
    
    If Not frm���ڷ�Χ_����.Show_ME(strStart, strEnd) Then Exit Function
    strStart = Format(strStart, "yyyy-MM-dd")
    strEnd = Format(strEnd, "yyyy-MM-dd")
'    ���    ���    ���˵��    ��󳤶�    �Ƿ��Ϊ��  ��ע
'    1   hospital_id    ҽԺ���        20  ��
'    2   startdate      ������ʼ����        ��  ��ʽ:YYYY-MM-DD
'    3   enddate        ������ֹ����        ��  ��ʽ:YYYY-MM-DD
    
    gcnOracle.BeginTrans
    blnTrans = True
    Call DebugTool("��ʼ�����������ر�����")
    'ɾ�����н����һ��һ����ȡһ�Σ�
    gstrSQL = "Delete �����_���ڷ�Χ"
    gcnOracle.Execute gstrSQL
    gstrSQL = "Delete סԺ�����_��ϸ�嵥"
    gcnOracle.Execute gstrSQL
    gstrSQL = "Delete ��������_��ϸ�嵥"
    gcnOracle.Execute gstrSQL
    gstrSQL = "Delete �涨�������_��ϸ�嵥"
    gcnOracle.Execute gstrSQL
    
    gstrSQL = " Insert Into �����_���ڷ�Χ" & _
              " (��ʼ����,��������)" & _
              " Values ('" & strStart & "','" & strEnd & "')"
    Call DebugTool("��ִ�е�SQL��" & gstrSQL)
    gcnOracle.Execute gstrSQL
    
    '��ȡסԺ������ܱ�
    Call DebugTool("ȡסԺ������ܱ�")
    If Not סԺ�����(strStart, strEnd) Then GoTo ExitSub
    
    'ȡ���������ܱ�
    Call DebugTool("ȡ���������ܱ�")
    If Not ��������(strStart, strEnd) Then GoTo ExitSub
    
    'ȡ���ֲ�������ܱ�
    Call DebugTool("ȡ����涨��������ܱ�")
    If Not ���ֲ������(strStart, strEnd) Then GoTo ExitSub
    
    gcnOracle.CommitTrans
    ��ȡ�����_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
ExitSub:
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Private Function סԺ�����(ByVal strStart As String, ByVal strEnd As String) As Boolean
    Dim intColumn As Integer, intColumns As Integer
    Dim strData As String, strTemp As String
    Const strColumn As String = "name|corp_name|pers_name|official|begin_date|end_date|total_declare" & _
    "|total|allself|partself|saccount|fund_offi|fund_bs|fund_add|cash_sta|sacc_sta|offi_sta|cash_bs|sacc_bs" & _
    "|offi_bs|cash_add|sacc_add|offi_add"
    Dim arrColumn
    On Error GoTo errHand
'    1   name           ����            10
'    2   corp_name      ��λ����        50
'    3   pers_name      ��Ա���        20
'    4   official       �Ƿ���Ա      1   "��"-��,""-��
'    5   begin_date     ��Ժ����        ��ʽ:YYYY-MM-DD
'    6   end_date       ��Ժ����        ��ʽ:YYYY-MM-DD
'    7   total_declare  ���ν���ͳ���������    18  ��λ��Ԫ
'    8   total          �ܷ���          18  ��λ��Ԫ
'    9   allself        �ֽ�֧���Է�    18  ��λ��Ԫ
'    10  partself       �ֽ�֧���Ը�    18  ��λ��Ԫ
'    11  saccount       ����Ӧ�����ø����ʻ�    18  ��λ��Ԫ
'    12  fund_offi      ����Ӧ�����ù���Ա����  18  ��λ��Ԫ
'    13  fund_bs        ����Ӧ������ͳ�����    18  ��λ��Ԫ
'    14  fund_add       ���䱣��Ӧ��    18  ��λ��Ԫ
'    15  cash_sta       �����ֽ�      18  ��λ��Ԫ
'    16  sacc_sta       ���߸����ʻ�  18  ��λ��Ԫ
'    17  offi_sta       ���߹���Ա����18  ��λ��Ԫ
'    18  cash_bs        ����ҽ�Ʊ����ֽ�18  ��λ��Ԫ
'    19  sacc_bs        ����ҽ�Ʊ��ո����ʻ�    18  ��λ��Ԫ
'    20  offi_bs        ����ҽ�Ʊ��չ���Ա����  18  ��λ��Ԫ
'    21  cash_add       ����ҽ�Ʊ����ֽ�        18  ��λ��Ԫ
'    22  sacc_add       ����ҽ�Ʊ��ո����ʻ�    18  ��λ��Ԫ
'    23  offi_add       ����ҽ�Ʊ��չ���Ա����  18  ��λ��Ԫ
    gstrField_������ = "hospital_id||startdate||enddate"
    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & strStart & "||" & strEnd
    If Not ���ýӿ�_׼��_������(Function_������.������ܱ�_סԺ) Then Exit Function
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
    If Not ���ýӿ�_ִ��_������ Then Exit Function
    
    arrColumn = Split(strColumn, "|")
    intColumns = UBound(arrColumn)
    
    If Not ���ýӿ�_ָ����¼��_������("His") Then Exit Function
    gstrSQL = " Insert Into סԺ�����_��ϸ�嵥" & _
              "(����,��λ����,��Ա���,����Ա,��Ժ����,��Ժ����,����ͳ��,�ܷ���,���Է�,���Ը�,�����ʻ�," & _
              " ����Ա����,ͳ�����,���䱣��,����,���߸����ʻ�,���߹���Ա����,����ҽ�Ʊ����ֽ�," & _
              " ����ҽ�Ʊ��ո����ʻ�,����ҽ�Ʊ��չ���Ա����,����ҽ�Ʊ����ֽ�,����ҽ�Ʊ��ո����ʻ�,����ҽ�Ʊ��չ���Ա����)" & _
              " Values ("
    If ���ýӿ�_��¼��_������ Then
        Do While True
            strTemp = gstrSQL
            For intColumn = 0 To intColumns
                Call ���ýӿ�_��ȡ����_������(arrColumn(intColumn), strData)
                Call CombinateString(strTemp, strData)
            Next
            strTemp = strTemp & ")"
            Call DebugTool("��ִ�е�SQL��" & strTemp)
            gcnOracle.Execute strTemp
            
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
    End If
    
    סԺ����� = True
errHand:
End Function

Private Function ��������(ByVal strStart As String, ByVal strEnd As String) As Boolean
    Dim intColumn As Integer, intColumns As Integer
    Dim strData As String, strTemp As String
    Const strColumn As String = "name|corp_name|pers_name|official|end_date|total|allself|partself|saccount|fund_offi|fund_bs"
    Dim arrColumn
    On Error GoTo errHand
'    1   name           ����            10
'    2   corp_name      ��λ����        50
'    3   pers_name      ��Ա���        20
'    4   official       �Ƿ���Ա      1   "��"-��,""-��
'    5   end_date       ��Ժ����        ��ʽ:YYYY-MM-DD
'    6   total          �ܷ���          18  ��λ��Ԫ
'    7   allself        �ֽ�֧���Է�    18  ��λ��Ԫ
'    8   partself       �ֽ�֧���Ը�    18  ��λ��Ԫ
'    9   saccount       ����Ӧ�����ø����ʻ�    18  ��λ��Ԫ
'    10  fund_offi      ����Ӧ�����ù���Ա����  18  ��λ��Ԫ
'    11  fund_bs        ����Ӧ������ͳ�����    18  ��λ��Ԫ
    gstrField_������ = "hospital_id||startdate||enddate"
    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & strStart & "||" & strEnd
    If Not ���ýӿ�_׼��_������(Function_������.������ܱ�_����) Then Exit Function
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
    If Not ���ýӿ�_ִ��_������ Then Exit Function
    
    arrColumn = Split(strColumn, "|")
    intColumns = UBound(arrColumn)
    
    If Not ���ýӿ�_ָ����¼��_������("His") Then Exit Function
    gstrSQL = " Insert Into ��������_��ϸ�嵥" & _
              "(����,��λ����,��Ա���,����Ա,��Ժ����,�ܷ���,���Է�,���Ը�,�����ʻ�,����Ա����,ͳ�����)" & _
              " Values ("
    If ���ýӿ�_��¼��_������ Then
        Do While True
            strTemp = gstrSQL
            For intColumn = 0 To intColumns
                Call ���ýӿ�_��ȡ����_������(arrColumn(intColumn), strData)
                Call CombinateString(strTemp, strData)
            Next
            strTemp = strTemp & ")"
            Call DebugTool("��ִ�е�SQL��" & strTemp)
            gcnOracle.Execute strTemp
            
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
    End If
    
    �������� = True
errHand:
End Function

Private Function ���ֲ������(ByVal strStart As String, ByVal strEnd As String) As Boolean
    Dim intColumn As Integer, intColumns As Integer
    Dim strData As String, strTemp As String
    Const strColumn As String = "name|corp_name|pers_name|official|end_date|total_declare" & _
    "|total|allself|partself|saccount|fund_offi|fund_bs|fund_add|cash_sta|sacc_sta|offi_sta|cash_bs|sacc_bs" & _
    "|offi_bs|cash_add|sacc_add|offi_add"
    Dim arrColumn
    On Error GoTo errHand
'    1   name           ����            10
'    2   corp_name      ��λ����        50
'    3   pers_name      ��Ա���        20
'    4   official       �Ƿ���Ա      1   "��"-��,""-��
'    5   end_date       ��Ժ����        ��ʽ:YYYY-MM-DD
'    6   total_declare  ���ν���ͳ���������    18  ��λ��Ԫ
'    7   total          �ܷ���          18  ��λ��Ԫ
'    8   allself        �ֽ�֧���Է�    18  ��λ��Ԫ
'    9   partself       �ֽ�֧���Ը�    18  ��λ��Ԫ
'    10  saccount       ����Ӧ�����ø����ʻ�    18  ��λ��Ԫ
'    11  fund_offi      ����Ӧ�����ù���Ա����  18  ��λ��Ԫ
'    12  fund_bs        ����Ӧ������ͳ�����    18  ��λ��Ԫ
'    13  fund_add       ���䱣��Ӧ��    18  ��λ��Ԫ
'    14  cash_sta       �����ֽ�      18  ��λ��Ԫ
'    15  sacc_sta       ���߸����ʻ�  18  ��λ��Ԫ
'    16  offi_sta       ���߹���Ա����18  ��λ��Ԫ
'    17  cash_bs        ����ҽ�Ʊ����ֽ�18  ��λ��Ԫ
'    18  sacc_bs        ����ҽ�Ʊ��ո����ʻ�    18  ��λ��Ԫ
'    19  offi_bs        ����ҽ�Ʊ��չ���Ա����  18  ��λ��Ԫ
'    20  cash_add       ����ҽ�Ʊ����ֽ�        18  ��λ��Ԫ
'    21  sacc_add       ����ҽ�Ʊ��ո����ʻ�    18  ��λ��Ԫ
'    22  offi_add       ����ҽ�Ʊ��չ���Ա����  18  ��λ��Ԫ
    gstrField_������ = "hospital_id||startdate||enddate"
    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & strStart & "||" & strEnd
    If Not ���ýӿ�_׼��_������(Function_������.������ܱ�_����涨��) Then Exit Function
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
    If Not ���ýӿ�_ִ��_������ Then Exit Function
    
    arrColumn = Split(strColumn, "|")
    intColumns = UBound(arrColumn)
    
    If Not ���ýӿ�_ָ����¼��_������("His") Then Exit Function
    gstrSQL = " Insert Into �涨�������_��ϸ�嵥" & _
              "(����,��λ����,��Ա���,����Ա,��Ժ����,����ͳ��,�ܷ���,���Է�,���Ը�,�����ʻ�," & _
              " ����Ա����,ͳ�����,���䱣��,����,���߸����ʻ�,���߹���Ա����,����ҽ�Ʊ����ֽ�," & _
              " ����ҽ�Ʊ��ո����ʻ�,����ҽ�Ʊ��չ���Ա����,����ҽ�Ʊ����ֽ�,����ҽ�Ʊ��ո����ʻ�,����ҽ�Ʊ��չ���Ա����)" & _
              " Values ("
    If ���ýӿ�_��¼��_������ Then
        Do While True
            strTemp = gstrSQL
            For intColumn = 0 To intColumns
                Call ���ýӿ�_��ȡ����_������(arrColumn(intColumn), strData)
                Call CombinateString(strTemp, strData)
            Next
            strTemp = strTemp & ")"
            Call DebugTool("��ִ�е�SQL��" & strTemp)
            gcnOracle.Execute strTemp
            
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
    End If
    
    ���ֲ������ = True
errHand:
End Function

Private Function GetPatientSQL(strSQL As String, ByVal intType As Integer) As Boolean
    Const str���� As String = ",name,sex,age,insr_code,idcard,pers_name,official_name,hospital_name,corp_name,disease,begin_date,grade_name,official,"
    Const str����涨�� As String = ",name,sex,age,patient_id,idcard,pers_name,official_name,hospital_name,corp_name,disease,begin_date,grade_name,official,"
    Const strסԺ As String = ",name,sex,age,insr_code,idcard,pers_name,official_name,official,patient_id,hospital_name,grade_name,corp_name,disease,begin_date,end_date,days,in_area_name,in_dept_name,in_bed,"
    Dim strCompare As String
    On Error GoTo errHand
    'intType��1-����;2-����涨��;3-סԺ
    
    Select Case intType
    Case 1
        strCompare = str����
    Case 2
        strCompare = str����涨��
    Case 3
        strCompare = strסԺ
    End Select
    
    Call CombinateString(gstrSQL, GetSQL(strCompare, "name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "sex"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "age"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "insr_code"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "idcard"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "pers_name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "official_name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "official"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "patient_id"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "hospital_name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "grade_name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "corp_name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "disease"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "begin_date"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "end_date"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "days"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "in_area_name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "in_dept_name"))
    Call CombinateString(gstrSQL, GetSQL(strCompare, "in_bed"))
    gstrSQL = gstrSQL & ")"
    GetPatientSQL = True
errHand:
End Function

Private Function GetSQL(ByVal strCompareSQL As String, ByVal strColumn As String) As String
    Dim strData As String
    If InStr(1, strCompareSQL, "," & strColumn & ",") <> 0 Then
        If Not ���ýӿ�_��ȡ����_������(strColumn, strData) Then Exit Function
        GetSQL = strData
    End If
End Function

Private Function GetHistorySQL(strSQL As String, ByVal intType As Integer) As Boolean
    Const str���� As String = ",declare_fee,inhosp_count,total_fee,cur_total_fee,start_pay,self_pay,fund_pay,additional_pay,official_pay,biz_times,"
    Const str����涨�� As String = ",declare_fee,total_fee,start_pay,fund_pay,all_self_pay,self_pay,percent_pay,additional_pay,official_pay,biz_times,"
    Const strסԺ As String = ",declare_fee,inhosp_count,total_fee,cur_total_fee,start_pay,self_pay,fund_pay,additional_pay,official_pay,biz_times,"
    Dim strCompare As String
    On Error GoTo errHand
    'intType��1-����;2-����涨��;3-סԺ
    
    Select Case intType
    Case 1
        strCompare = str����
    Case 2
        strCompare = str����涨��
    Case 3
        strCompare = strסԺ
    End Select
    
    If intType = 2 Then
        Call CombinateString(gstrSQL, GetSQL(strCompare, "declare_fee"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "inhosp_count"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "total_fee"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "cur_total_fee"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "start_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "self_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "fund_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "additional_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "official_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "biz_times"))
    Else
        Call CombinateString(gstrSQL, GetSQL(strCompare, "declare_fee"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "total_fee"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "start_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "fund_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "all_self_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "self_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "percent_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "additional_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "official_pay"))
        Call CombinateString(gstrSQL, GetSQL(strCompare, "biz_times"))
    End If
    gstrSQL = gstrSQL & ")"
    GetHistorySQL = True
errHand:
End Function

Private Sub CombinateString(strSQL As String, ByVal strData As String)
    '����ַ���
    'intType��1-��ͨ;2-�ַ���
    strSQL = strSQL & "," & "'" & strData & "'"
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Public Sub ���²���_������(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    Dim arrPatient
    Dim str�������� As String, str��Ժ���� As String, str��Ժ���� As String, str����֢ As String
    Dim rsTemp As New ADODB.Recordset
    
    '������Ժ���˵Ĳ�����Ϣ
    Call ��ȡ���˻�����Ϣ(lng����ID, False)
    arrPatient = Split(��ȡ���������Ϣ(lng����ID, lng��ҳID), "||")
        
    '��ȡ���˵ĳ�Ժ���
    gstrSQL = "Select decode(��Ժ��ʽ,'תԺ',3,0) ��Ժ��ʽ,��Ժ���� From ������ҳ " & _
            " Where ����ID = [1] And ��ҳID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ��ʽ", lng����ID, lng��ҳID)
    str�������� = Format(rsTemp!��Ժ����, "yyyy-MM-dd")
    
    If Not frm����ѡ��_����.ShowSelect(TYPE_������, lng����ID, lng��ҳID, str��Ժ����, str��Ժ����, str����֢) Then Exit Sub
    
    '���²�����Ժ�������Ժ����
    '    1   hospital_id    ҽ�ƻ�������20  ��
    '    2   serial_no      ҵ�����к�  12  ��
    '    3   busi_type      ҵ������    2   ��  "12"��סԺ
    '    4   staff_no       ����Ա����  5   ��
    '    5   staff_name     ����Ա����  10  ��
    '    6   begin_date     ����ʱ��        ��  ��ʽ��YYYY-MM-DD
    '    7   in_dept        ��Ժ���ұ��3   ��
    '    8   in_dept_name   ��Ժ��������20  ��
    '    9   in_area        ��Ժ�������3   ��
    '    10  in_area_name   ��Ժ��������20  ��
    '    11  in_bed         ��Ժ�������10  ��
    '    12  bed_type       ��λ����    1   ��  "0"����ͨ��λ��"1"�����ȣ�"2"�����ۣ�"3"���߸�
    '    13  patient_id     סԺ��      20  ��
    '    14  old_patient_id ԭסԺ��    20  ��
    '    15  in_disease     ��Ժ���    20  ��  ��������
    '    16  note           ��ע        100 ��
    '    17  fin_disease    ��Ժ����
    gstrField_������ = "hospital_id||serial_no||busi_type||staff_no||staff_name||begin_date||" & _
                    "in_dept||in_dept_name||in_area||in_area_name||in_bed||bed_type||patient_id||old_patient_id||in_disease||note||fin_disease"
    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & gCominfo_������.ҵ�����к� & "||" & _
                    gCominfo_������.ҵ������ & "||" & gCominfo_������.����Ա���� & "||" & _
                    gstrUserName & "||" & str�������� & "||" & arrPatient(��Ժ���ұ��) & "||" & _
                    arrPatient(��Ժ��������) & "||" & arrPatient(��Ժ�������) & "||" & _
                    arrPatient(��Ժ��������) & "||" & arrPatient(��Ժ�������) & "||" & _
                    arrPatient(��λ����) & "||" & arrPatient(סԺ��) & "||" & _
                    arrPatient(סԺ��) & "||" & str��Ժ���� & "||" & str����֢ & "||" & str��Ժ����
    If Not ���ýӿ�_׼��_������(Function_������.סԺ��Ϣ_�޸�) Then Exit Sub
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Sub
    If Not ���ýӿ�_ִ��_������ Then Exit Sub
End Sub

Private Function FirstAid() As Boolean
    '��鱾����վ�Ƿ��Ǽ�����ҩר�û�
    FirstAid = (GetSetting("ZLSOFT", "����ҽ�����߰�", "������ҩר�û�", 0) = 1)
End Function
