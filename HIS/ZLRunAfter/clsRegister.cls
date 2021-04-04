VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==================================================================================================
'��д           lshuo
'����           2018/12/25
'ģ��           clsRegister
'˵��
'==================================================================================================
Private Const mstrCurModule     As String = "clsRegister"           '��ǰģ������
'����������ZL9Rigister����
Private mcnOracle           As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Private mstrUserName        As String               '�û���
Private mstrPassword        As String               '����
Private mstrServer          As String               '������
Private mobjRegister        As Object               '10.35.10���ϵ�ZLRigster�����������ڸ����ڹ�������ֻ��һ��ʵ������˷������ڣ�����ֲ�������������뽫�ö������ģ��

Public Enum enuProvider
    MSODBC = 0
    OraOLEDB = 1
    OriginalConnection = 9
End Enum


Private Sub Class_Initialize()
    On Error Resume Next
    Set mobjRegister = CreateObject("zlRegisterAlone.clsRegister")
    Err.Clear
    If mobjRegister Is Nothing Then
        Set mobjRegister = CreateObject("zlRegister.clsRegister")
        Err.Clear: On Error GoTo 0
        '��ǰ����������10.35.10���¡�
    End If
End Sub

Private Sub Class_Terminate()
    Set mobjRegister = Nothing
    Set mcnOracle = Nothing
End Sub


Public Function GetConnection(ByVal strServer As String, ByVal strUsername As String, ByVal strPassword As String, ByVal blnTransPassword As Boolean, _
     Optional ByVal bytProvider As enuProvider = MSODBC, Optional ByRef strError As String = "���뷵�ش�����Ϣ", Optional ByVal blnSaveAccount As Boolean = True) As ADODB.Connection
'���ܣ� ��ָ�������ݿ⣬��������ʵ������ADO���Ӷ���(�����10.35.10��ǰ�����룬���µ�ת�������������)
'������ strServer       :�������������߿���ֱ��ָ��IP:Port/SID
'       strUserName     :�û���
'       strPassword     :����
'       blnTransPassword:�Ƿ��������ת��
'       bytProvider     :�����ݿ����ӵ����ַ�ʽ,0-msODBC��ʽ,1-OraOLEDB��ʽ
'       strError        :����ʧ�ܺ����ָ���˴˲������򷵻ش�����Ϣ��δָ��ʱֱ�ӵ�����ʾ��Ϣ��
'       blnSaveAccount  :�����û��������롢����������ȫ�ֱ�����һ�㣬���ڵ�¼����ʱ���棬���ӿ�ReGetConnection��GetUserName��GetServerName��GetPassword��LoginValidateʹ�ã�
'���أ� ���ݿ�򿪳ɹ������Ӷ����״̬���Է���adStateOpen(1),ʧ���򷵻�AdStateClosed(0)
    If mobjRegister Is Nothing Then
        Dim cnOracle As ADODB.Connection
        Dim strPWD_New As String, strPWD_Old As String
        Dim blnPrompt As Boolean
        
        blnPrompt = (strError = "���뷵�ش�����Ϣ")
        strError = ""
        
        If blnTransPassword Then
            strPWD_Old = TransPassword(strPassword)
        Else
            strPWD_Old = strPassword
        End If
        Set cnOracle = OraDBOpen(strServer, strUsername, strPWD_Old, bytProvider, strError, blnSaveAccount)
        
        If blnPrompt And strError <> "" Then
            MsgBox strError, vbExclamation, "��ʾ"
        End If
        
        mstrPassword = strPWD_Old
        Set GetConnection = cnOracle
    Else
        If blnSaveAccount Then
            mstrServer = strServer
            mstrUserName = strUsername
            If blnTransPassword Then
                mstrPassword = TransPassword(strPassword)
            Else
                mstrPassword = strPassword
            End If
        End If
        Set GetConnection = mobjRegister.GetConnection(strServer, strUsername, strPassword, blnTransPassword, bytProvider, strError, blnSaveAccount)
    End If
End Function


Public Function ReGetConnection(ByVal bytProvider As enuProvider, ByRef strError As String) As ADODB.Connection
'���ܣ����ص�¼����̨ʱ�����Ӷ��󣬻��߸���֮ǰ�򿪵����ݿ����Ӷ������»�ȡһ��OLEDB��MSODBC��ʽ�򿪵����Ӷ���
'������bytProvider  :�����ݿ����ӵ����ַ�ʽ,0-msODBC��ʽ,1-OraOLEDB��ʽ,9-��¼����̨ʱ�����Ӷ���
'      strError     :���ش�����ʧ�ܺ�Ĵ�����Ϣ
'���أ� ���ݿ�򿪳ɹ������Ӷ����״̬���Է���adStateOpen(1),ʧ���򷵻�AdStateClosed(0)
    If mobjRegister Is Nothing Then
        If bytProvider = OriginalConnection Then
            Set ReGetConnection = mcnOracle
        Else
            Set ReGetConnection = OraDBOpen(mstrServer, mstrUserName, mstrPassword, bytProvider, strError, False)
        End If
    Else
        Set ReGetConnection = mobjRegister.ReGetConnection(bytProvider, strError)
    End If
End Function

Public Function GetUserName(Optional ByVal blnUcase As Boolean = True) As String
'���ܣ���ȡ��ǰ��¼���û���
'������blnUcase     :�Ƿ�ת��Ϊ��д��ĸ
    If mobjRegister Is Nothing Then
        GetUserName = IIf(blnUcase, UCase(mstrUserName), mstrUserName)
    Else
        GetUserName = mobjRegister.GetUserName(blnUcase)
    End If
End Function

Public Function GetServerName(Optional ByVal blnUcase As Boolean = True) As String
'���ܣ���ȡ��ǰ��¼�ķ�������
'������blnUcase     :�Ƿ�ת��Ϊ��д��ĸ
    If mobjRegister Is Nothing Then
        GetServerName = IIf(blnUcase, UCase(mstrServer), mstrServer)
    Else
        GetServerName = mobjRegister.GetServerName(blnUcase)
    End If
End Function


Public Function GetPassword(Optional lngInstance As Long) As String
'���ܣ���ȡ��ǰ��¼(����GetConnection֮��)�û����ܺ������
'������lngInstance-��ǰӦ�ó���ʵ���ľ����û�д����������������̶�ʱ����Ҫ���룩
'˵����������ָ���ĳ���ʹ��ڷ��ʱ��ӿ�(δ����������£�������ʾ���֣����磺δ��Ȩ�ĳ���δ��Ȩ�Ĵ���)
    If mobjRegister Is Nothing Then
        GetPassword = mstrPassword
    Else
        GetPassword = mobjRegister.GetPassword(lngInstance)
    End If
    If GetPassword Like "δ��Ȩ�ĳ���:*" And IsDesinMode Then
        GetPassword = mstrPassword
    End If
End Function

'======================================================================
'˽�з���
'======================================================================
Private Function TransPassword(ByVal strPassword) As String
'���ܣ�ת���û�����,������û�����дת��
'������
    Dim intBit As Integer, intLen As Integer
    Dim strPWDNew As String, strBit As String
    
    strPassword = UCase(Trim(strPassword))

     If Len(strPassword) = 0 Then TransPassword = "": Exit Function
     
     strPWDNew = ""
     For intBit = 1 To Len(strPassword)
         strBit = Mid(strPassword, intBit, 1)
         Select Case (intBit Mod 3)
         Case 1
             strPWDNew = strPWDNew & _
                 Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                    strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                    strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                    strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
         Case 2
             strPWDNew = strPWDNew & _
                 Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                    strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                    strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                    strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
         Case 0
             strPWDNew = strPWDNew & _
                 Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                    strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                    strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                    strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
         End Select
     Next
    
    TransPassword = strPWDNew
End Function

Private Function OraDBOpen(ByVal strServer As String, ByVal strUsername As String, ByVal strPassword As String, _
    ByVal bytProvider As enuProvider, ByRef strError As String, ByVal blnSaveAccount As Boolean) As ADODB.Connection
'���ܣ� ��ָ�������ݿ⣬������ADO���Ӷ���
'������ strServer���������������߿���ֱ��ָ��IP:Port/SID
'       strUserName���û���
'       strUserPwd������
'       bytProvider�������ݿ����ӵ����ַ�ʽ,0-msODBC��ʽ,1-OraOLEDB��ʽ
'       blnSaveAccount�������û��������롢����������ȫ�ֱ�����һ�㣬��¼ʱ�����򱣴棬��֤�û�����ʱ�����棩
'���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    Dim strPersist_Security_Info As String
    Dim arrTmp As Variant, strIp As String, strPort As String, strSID As String
    
    On Error Resume Next

    Set OraDBOpen = New ADODB.Connection
        
    With OraDBOpen
        If InStr(strServer, "/") > 0 Then
            arrTmp = Split(strServer, "/")
            strSID = arrTmp(1)
            If InStr(arrTmp(0), ":") > 0 Then
                arrTmp = Split(arrTmp(0), ":")
                strIp = arrTmp(0)
                strPort = arrTmp(1)
            Else
                strIp = arrTmp(0)
                strPort = "1521"
            End If
            strServer = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strIp & ")(PORT=" & strPort & "))(CONNECT_DATA=(SERVICE_NAME=" & strSID & ")))"
            
            '�������ּ���ADDRESS_LIST��д������ODBC�£�ֻ֧��SID����֧��SERVICE_NAME;OLEDB�����ֶ�֧��
            'If bytProvider = enuProvider.MSODBC Then
            'strServer = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strIP & ")(PORT=" & strPort & ")))(CONNECT_DATA=(SID=" & strSID & ")))"
        End If
        
        '��Persist Security InfoΪfalseʱ�����Ӷ����ConnectionString�����У����������룬MSDataShape��ʽ��������������������,���ԣ���ģ������洢���Ա�����һ�����ӷ�ʽʱʹ��
        If blnSaveAccount Then
            mstrServer = strServer
            mstrUserName = strUsername
            mstrPassword = strPassword
        End If
        
        strPersist_Security_Info = ";Persist Security Info=False" '��������ߴӷ��ص����Ӷ����л���û����룬��ָ�������ԵĻ���ȱʡ��false
        
        'ȱʡΪadUseServer�������ָ�����䣬������OLEDB�򿪵����ӣ�����Command����Execute�������ص�Recordset�����ActiveConnection = Nothing�ᱨ��:�����ʱ����������(MSODBC��ʽ�򿪵����Ӳ��ᱨ��)
        .CursorLocation = adUseClient
        
        If bytProvider = enuProvider.MSODBC Then
            .Provider = "MSDataShape"
            .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer & strPersist_Security_Info, strUsername, strPassword
        Else
            .Provider = "OraOLEDB.Oracle"
            .Open "PLSQLRSet=1;Data Source=" & strServer & strPersist_Security_Info, strUsername, strPassword
            'DistribTX=1,�����ֲ�����(ȱʡ);DistribTx=0:���ηֲ�����oracle8.1.7�汾��BUG������10.35.10֮ǰ�Ĺ������ߵ�¼ʱ�ǽ��õġ�
            'PLSQLRSet=1 ���ڲ��������α�����Ĵ洢���̣�Ҳ��д��Extended Properties=PLSQLRSet=1
        End If
    End With
    
    If Err = 0 Then
        strError = ""
    Else
        strError = Err.Description
        On Error GoTo 0
        
        If InStr(strError, "�Զ�������") > 0 Then
            If bytProvider = enuProvider.MSODBC Then
                strError = "msoracl32.dll"
            Else
                strError = "OraOLEDB.dll"
            End If
            strError = "�޷��������Ӷ����������ݷ��ʲ���(" & strError & ")�Ƿ�������װ��ע�ᡣ"
        ElseIf InStr(strError, "ORA-12505") > 0 Then
            strError = "ORA-12505,��������ǰ�޷�ʶ���������������������� SID,��������������õ�ʵ�����ơ�"
            
        ElseIf InStr(strError, "ORA-12170") > 0 Then
            strError = "ORA-12170,���ӳ�ʱ��������������Ƿ���ȷ�������Ƿ�ɷ��ʣ��Լ��Ƿ񱻷���������ǽ��ֹ��"
            
        ElseIf InStr(strError, "ORA-12154") > 0 Then
            strError = "ORA-12154,�޷���������������" & vbCrLf & "���鱾����Oracle�����ļ�(tnsnames.ora)���Ƿ���ڵ�ǰʹ�õķ�������"
            
        ElseIf InStr(strError, "ORA-12541") > 0 Then
            strError = "ORA-12541,�޷����ӷ�����������������ϵ�Oracle�����������Ƿ�������"
            
        ElseIf InStr(strError, "ORA-01033") > 0 Then
            strError = "ORA-01033,ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�"
            
        ElseIf InStr(strError, "ORA-01034") > 0 Then
            strError = "ORA-01034,ORACLE�����ã��������ݿ�ʵ���Ƿ�������"
            
        ElseIf InStr(strError, "ORA-02391") > 0 Then
            strError = "ORA-02391,�û�" & strUsername & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ������������¼��)��"
            
        ElseIf InStr(strError, "ORA-01017") > 0 Then
            strError = "ORA-01017,��Ч���û��������룬��¼���ܾ���"
        
        ElseIf InStr(strError, "ORA-28000") > 0 Then
            strError = "ORA-28000,���û��Ѿ������ã���������¼��"
        End If
    End If
End Function

Private Function UpdatePassword(ByVal cnOracle As ADODB.Connection, ByVal strUsername As String, ByVal strPassword As String, ByVal blnTransPassword As Boolean) As String
'���ܣ��޸��û�����
'���أ��޸ĳ�����Ĵ�������
    Dim strPWD_New As String
    
    If blnTransPassword Then
        strPWD_New = TransPassword(strPassword)
    Else
        strPWD_New = strPassword
    End If
    
    '���������븴�Ӷ�Ҫ��ʱ�������ж��ŵ������ַ���������Ҫ��˫����
    'Υ�����밲ȫ���Ե���������
    On Error Resume Next
    cnOracle.Execute "Alter user " & strUsername & " identified by """ & strPWD_New & """"
    
    If Err.Number <> 0 Then
       UpdatePassword = Err.Description
       Err.Clear
    End If
End Function

Private Sub GetConnectionInfo(ByVal strConect As String, ByRef strServerName As String, ByRef strUsername As String, ByRef strPassword As String)
'���ܣ� ����MSODBC���Ӷ����е�ORACLE���е� ���������û���������
'���أ� �ɹ�ʧ�ܣ�����True��ʧ�ܣ�����False

    Dim i As Integer
    Dim strTemp As String
    If strConect = "" Then Exit Sub
            
    strServerName = ""
    strUsername = ""
    strPassword = ""
    strConect = Replace(strConect, """", "")
    
    If InStr(strConect, "ODBC") > 0 Then
        'Provider=MSDataShape.1;Extended Properties="Driver={Microsoft ODBC for Oracle};Server=DYYY";Persist Security Info=True;User ID=zlhis;Password=his;Data Provider=MSDASQL"
        'Provider=MSDataShape.1;Persist Security Info=False;User ID=ZLHIS;Data Provider=MSDASQL;
        '��ȡ strServerName(SecurityΪfalseʱ���޷����)
        i = InStrRev(strConect, "Server=", -1)
        If i > 0 Then
            strTemp = Right(strConect, Len(strConect) - i - 6)
            i = InStr(1, strTemp, ";")
            If i > 0 Then
                strServerName = Left(strTemp, i - 1)
            End If
        End If
    Else
        'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source="DYYY";Extended Properties="PLSQLRSet=1"
        'Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=ZLHIS;Data Source="DYYY"
        i = InStrRev(strConect, "Data Source=", -1)
        If i > 0 Then
            strTemp = Right(strConect, Len(strConect) - i - 11)
            i = InStr(1, strTemp, ";")
            If i > 0 Then
                strServerName = Left(strTemp, i - 1)
            Else    'SecurityΪfalseʱ��û��;��
                strServerName = strTemp
            End If
        End If
    End If
    
    '��ȡ strUserName
    i = InStrRev(strConect, "User ID=", -1)
    If i > 0 Then
        strTemp = Right(strConect, Len(strConect) - i - 7)
        i = InStr(1, strTemp, ";")
        If i > 0 Then
            strUsername = Left(strTemp, i - 1)
        End If
    End If
    
    '��ȡ strPassword
    i = InStrRev(strConect, "Password=", -1)
    If i > 0 Then
        strTemp = Right(strConect, Len(strConect) - i - 8)
        i = InStr(1, strTemp, ";")
        If i > 0 Then
            strPassword = Left(strTemp, i - 1)
        End If
    End If
End Sub

Public Function OpenSQLRecord(ByVal cnOracle As ADODB.Connection, ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    Dim arrPars() As Variant
    arrPars = arrInput
    Set OpenSQLRecord = OpenSQLRecordByArray(cnOracle, strSQL, strTitle, arrPars)
End Function

Public Function OpenSQLRecordByArray(ByVal cnOracle As ADODB.Connection, ByVal strSQL As String, ByVal strTitle As String, arrInput() As Variant, Optional intLobOprate As Integer = 0) As ADODB.Recordset
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'               ��Ϊʹ�ð󶨱���,�Դ�"'"���ַ�����,����Ҫʹ��"''"��ʽ��
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'      intLobOprate=0:��ͨSQL,1:LOB���Ͷ�ȡSQL,2:LOB����SQL
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        If lngRight = 0 Then Exit Do
        '������������"[����]����"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop
    
    If UBound(arrInput) + 1 < intMax Then
        Err.Raise 9527, strTitle, "SQL���󶨱�����ȫ��������Դ��" & strTitle
    End If

    '�滻Ϊ"?"����
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next
    
    '�����µĲ���
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '����
            '���ַ�ʽ������һЩIN�Ӿ��Union���
            '��ʾͬһ�������Ķ��ֵ,�����Ų�������������Ĳ����Ž���,��Ҫ��֤�����ֵ��������
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '�ַ�
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '�ò������������õ��ڼ���ֵ��
        End Select
    Next
    
    'ִ�з��ؼ�¼��
    'If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = cnOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
    'End If

    cmdData.CommandText = strSQL
    If intLobOprate > 0 Then '����LOB,��ȡLOBҲҪʹ�øò��������������Լ10�����
        Set OpenSQLRecordByArray = New ADODB.Recordset
        OpenSQLRecordByArray.Open cmdData, , adOpenStatic, adLockOptimistic
    Else
        Set OpenSQLRecordByArray = cmdData.Execute
    End If
    If intLobOprate < 2 Then
        Set OpenSQLRecordByArray.ActiveConnection = Nothing
    End If
End Function