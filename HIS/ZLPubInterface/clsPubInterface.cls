VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPubInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum enuProvider
    MSODBC = 0
    OraOLEDB = 1
End Enum
'======================================================================================================================
'�ӿ�           GetConnection           ��֤�����ӿڵ�APPID����Ȩ�룬�ɹ����򷵻ص�����ΪOpen״̬������ΪClosed״̬
'����ֵ         Connection����
'����б�:
'������         ����                    ˵��
'Key            String                  �����ӿ���Ȩ�������
'Server         String                  �����ӿ�Ҫ���ӵķ����������߿���ֱ��ָ��IP:Port/SID
'bytProvider    Byte��Optional��        �����ݿ����ӵ����ַ�ʽ,0-msODBC��ʽ,1-OraOLEDB��ʽ��ȱʡΪ0-msODBC��ʽ��(C#�汾û�иò���)
'strErrNote     String(Out,Optional)    ���ش�����Ϣ
'======================================================================================================================
Public Function GetConnection(ByVal strKey As String, ByVal strServer As String, Optional ByVal bytProvider As Byte = MSODBC, Optional ByRef strErrNote As String) As ADODB.Connection
    Dim conZLUA     As ADODB.Connection
    Dim conReturn   As New ADODB.Connection
    Dim strIntPWD   As String
    
    strErrNote = ""
    On Error GoTo errh
    '�������Ӵ���
    Set conZLUA = OraDBOpen(strServer, "ZLUA", Sm4DecryptEcb("ZLSV2:" & G_UA_PWD, GetGeneralAccountKey(G_UA_KEY)), MSODBC, strErrNote)
    If conZLUA.State = adStateClosed Then
        Set GetConnection = conReturn
        Exit Function
    End If
    'У����Ȩ��
    If Not CheckGrantKey(conZLUA, strKey, strErrNote) Then
        Set GetConnection = conReturn
        Exit Function
    End If
    '��ȡZLInterface����
    strIntPWD = GetZLInterfacePWD(conZLUA, strErrNote)
    If strErrNote <> "" Then
        Set GetConnection = conReturn
        Exit Function
    End If
    '����ZLInterface����
    Set conReturn = OraDBOpen(strServer, "ZLINTERFACE", strIntPWD, bytProvider, strErrNote)
    Set GetConnection = conReturn
    Exit Function
errh:
    Set GetConnection = conReturn
    strErrNote = "(" & Err.Number & ")" & Err.Description
    Err.Clear
End Function
'======================================================================================================================
'�ӿ�           GetUnitName             ��֤�����ӿڵ���Ȩ�룬�ɹ����ص�λ���ƣ����򷵻ؿմ�
'����ֵ         String
'����б�:
'������         ����                    ˵��
'Key            String                  �����ӿ���Ȩ�������
'varServerOrConnection   String         �����ӿ�Ҫ���ӵķ�����������ֱ��ָ��IP:Port/SID�������Ѿ��򿪵����Ӷ���
'strErrNote     String(Out,Optional)    ���ش�����Ϣ
'======================================================================================================================
Public Function GetUnitName(ByVal strKey As String, ByVal varServerOrConnection As Variant, Optional ByRef strErrNote As String) As String
    Dim connTmp     As ADODB.Connection
    Dim strSQL      As String, rstmp    As ADODB.Recordset
    
    On Error GoTo errh
    strErrNote = ""
    '���ݴ���Ĳ������ͳ�ʼ������
    If TypeName(varServerOrConnection) = "String" Then
        Set connTmp = GetConnection(strKey, varServerOrConnection, MSODBC, strErrNote)
        If connTmp.State = adStateClosed Then
            Exit Function
        End If
    ElseIf TypeName(varServerOrConnection) = "Connection" Then
        If varServerOrConnection Is Nothing Then
            strErrNote = "���������δʵ������"
            Exit Function
        ElseIf varServerOrConnection.State = adStateClosed Then
            strErrNote = "�����������δ�򿪡�"
            Exit Function
        Else
            Set connTmp = varServerOrConnection
            If Not CheckGrantKey(connTmp, strKey, strErrNote) Then
                strErrNote = "�ӿ���Ȩ����Ч��"
                Exit Function
            End If
        End If
    Else
        strErrNote = "����Ĳ���String��ADODB.Connection���͡�"
        Exit Function
    End If
    strSQL = "Select Max(����) ���� From zlRegInfo A Where a.��Ŀ  = [1]"
    Set rstmp = OpenSQLRecord(connTmp, strSQL, "GetUnitName", "��λ����")
    If IsNull(rstmp!����) Then
        strErrNote = "ȱʧ��λ������Ϣ����¼�������������߻��Զ��ؽ�����"
    Else
        GetUnitName = rstmp!���� & ""
    End If
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    strErrNote = "(" & Err.Number & ")" & Err.Description
    Err.Clear
End Function
'======================================================================================================================
'�ӿ�           GetPassword             ��֤�����ӿڵ���Ȩ�룬�ɹ������û��ĵ�¼���룬���򷵻ؿմ�
'����ֵ         String
'����б�:
'������         ����                    ˵��
'Key            String                  �����ӿ���Ȩ�������
'varServerOrConnection   String         �����ӿ�Ҫ���ӵķ�����������ֱ��ָ��IP:Port/SID�������Ѿ��򿪵����Ӷ���
'strUserName    String                  �û�����
'lngSys         Long                    ϵͳ��ţ�-1ȱʡ�ӱ�׼���ȡ��û�ж����������ϵͳ��ȡ������ϵͳ��űȽ�С��
'strErrNote     String(Out,Optional)    ���ش�����Ϣ
'======================================================================================================================
Public Function GetPassword(ByVal strKey As String, ByVal varServerOrConnection As Variant, ByVal strUserName As String, Optional ByVal lngSys As Long = -1, Optional ByRef strErrNote As String) As String
    Dim connTmp     As ADODB.Connection
    Dim strSQL      As String, rstmp    As ADODB.Recordset
    Dim strOwner    As String
    
    On Error GoTo errh
    strErrNote = ""
    '���ݴ���Ĳ������ͳ�ʼ������
    If TypeName(varServerOrConnection) = "String" Then
        '�������Ӵ���
        Set connTmp = OraDBOpen(varServerOrConnection, "ZLUA", Sm4DecryptEcb("ZLSV2:" & G_UA_PWD, GetGeneralAccountKey(G_UA_KEY)), MSODBC, strErrNote)
        If connTmp.State <> adStateOpen Then
            Exit Function
        End If
    ElseIf TypeName(varServerOrConnection) = "Connection" Then
        If varServerOrConnection Is Nothing Then
            strErrNote = "���������δʵ������"
            Exit Function
        ElseIf varServerOrConnection.State <> adStateOpen Then
            strErrNote = "�����������δ�򿪡�"
            Exit Function
        Else
            Set connTmp = varServerOrConnection
        End If
    Else
        strErrNote = "����Ĳ���String��ADODB.Connection���͡�"
        Exit Function
    End If
    If Not CheckGrantKey(connTmp, strKey, strErrNote) Then
        Exit Function
    End If
    If lngSys = -1 Then
        strSQL = "Select a.������ From zlSystems A Where a.��� In (100, 300, 400, 600, 2500) Order By a.���"
    Else
        strSQL = "Select a.������ From zlSystems A Where a.��� =[1]"
    End If
    Set rstmp = OpenSQLRecord(connTmp, strSQL, "��ȡ�ϻ���Ա��������", lngSys)
    If Not rstmp.EOF Then
        strOwner = rstmp!������ & ""
        strSQL = "Select Max(��¼����) ��¼���� From " & strOwner & ".�ϻ���Ա�� A Where a.�û���  = [1]"
        Set rstmp = OpenSQLRecord(connTmp, strSQL, "GetPassword", UCase(strUserName))
        If Not IsNull(rstmp!��¼����) Then
            GetPassword = Sm4DecryptEcb(rstmp!��¼���� & "", GetGeneralAccountKey(G_PASSWORD_KEY))
        End If
    Else
        strErrNote = IIf(lngSys = -1, "��ǰ���ݿⲻ�����ϻ���Ա��", "��ǰϵͳ�������ϻ���Ա��")
    End If
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
    strErrNote = "(" & Err.Number & ")" & Err.Description
    Err.Clear
End Function