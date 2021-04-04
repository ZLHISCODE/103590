Attribute VB_Name = "mdlPublicOneCard"
Option Explicit
'--------------------------------------------------------------------------------------------------
'--ϵͳ
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public glngModul As Long, glngSys As Long
Public gstrAviPath As String, gstrVersion As String
Public gstrMatchMethod As String
Public gstrProductName As String
Public gstrComputerName As String
Public gstrHelpPath As String
Public gstrDBUser As String   '��ǰ���ݿ��û�
Public gstrUnitName As String '�û���λ����
Public gcnOracle As ADODB.Connection
Public gstrNodeNo As String
Public gblnAutoGetOracleConnect As Boolean   '�Ƿ��Զ���ȡOracle����

Public Type Ty_UserInfor
    id As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    �������� As String
    
End Type
Public UserInfo As Ty_UserInfor
Public glngInstanceCount As Long    'ʵ����

'-----------------------------------------------------------------------------------------------------
'С����ʽ����
Public Enum gС������
    g_���� = 0
    g_�ɱ���
    g_�ۼ�
    g_���
    g_�ۿ���
End Enum
Private Type m_С��λ
    ����С�� As Integer
    �ɱ���С�� As Integer
    ���ۼ�С�� As Integer
    ���С�� As Integer
    �ۿ��� As Integer
End Type

Public g_С��λ�� As m_С��λ
Public Type g_FmtString
    FM_���� As String
    FM_�ɱ��� As String
    FM_���ۼ� As String
    FM_��� As String
    FM_�ۿ��� As String
End Type
Public gVbFmtString As g_FmtString
Public gOraFmtString As g_FmtString
'-----------------------------------------------------------------------------------------------------
'��ɫ�������
Public Type Ty_Color
     lngGridColorSel As OLE_COLOR     'ѡ����ɫ
     lngGridColorLost As OLE_COLOR   '�뿪��ɫ
End Type
Public gSysColor As Ty_Color
'-----------------------------------------------------------------------------------------------------
'��������(zl9ComLib)
Public gobjComLib As Object
Public gobjCommFun As Object
Public gobjDatabase As Object
Public gobjControl As Object
Public gobjOneDataBase As clsDataBase      'һ��ͨ�������Ӷ���
Public gobjOneDataObject As clsOneCardDataObject   'һ��ͨ���ݶ���
'------------------------------------------------------------------------------------------------------------------------------------
'Api����.
'��������(ComputerName)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


'------------------------------------------------------------------------------------------------------------------------------------
'����
Private Type Ty_TestDebug
    blndebug As Boolean
    objSquareCard As clsCard
    bytType  As Byte  '1-�����������,2-��ȡ����
    strStartNo As String    '��ʼ����
    bln�������� As Boolean
End Type
Public gTy_TestBug As Ty_TestDebug
Public gbln�Զ���ȡ As Boolean '��ǰ�Ƿ�Ϊ��Ƶ��

Public Sub ��ʼС��λ��()
    '------------------------------------------------------------------------------------------------------
    '����:��ʼС��λ��
    '���:
    '����:
    '����:7
    '�޸���:���˺�
    '�޸�ʱ��:2007/3/6
    '------------------------------------------------------------------------------------------------------
    With g_С��λ��
        .�ɱ���С�� = 7
        .���ۼ�С�� = 7
        .���С�� = 2
        .����С�� = 3
        .�ۿ��� = 2
    End With
    With gVbFmtString
        .FM_�ɱ��� = GetFmtString(g_�ɱ���, False)
        .FM_��� = GetFmtString(g_���, False)
        .FM_���ۼ� = GetFmtString(g_�ۼ�, False)
        .FM_���� = GetFmtString(g_����, False)
        .FM_�ۿ��� = GetFmtString(g_�ۿ���, False)
    End With
    With gOraFmtString
        .FM_�ɱ��� = GetFmtString(g_�ɱ���, True)
        .FM_��� = GetFmtString(g_���, True)
        .FM_���ۼ� = GetFmtString(g_�ۼ�, True)
        .FM_���� = GetFmtString(g_����, True)
        .FM_�ۿ��� = GetFmtString(g_�ۿ���, True)
    End With
End Sub

Public Function GetFmtString(ByVal С������ As gС������, Optional blnOracle As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------
    '����:����ָ����С����ʽ��
    '���: lngС��λ��-С��λ��
    '     blnOracle-������oracle�ĸ�ʽ������Vb�ĸ�ʽ��
    '����:
    '����:����ָ���ĸ�ʽ��
    '�޸���:���˺�
    '�޸�ʱ��:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim strFmt As String
    Dim intλ�� As Integer
    Select Case С������
    Case g_����
         intλ�� = g_С��λ��.����С��
    Case g_���
         intλ�� = g_С��λ��.���С��
    Case g_�ɱ���
         intλ�� = g_С��λ��.�ɱ���С��
    Case g_�ۼ�
         intλ�� = g_С��λ��.���ۼ�С��
    Case Else
        intλ�� = 0
    End Select
    If blnOracle Then
       GetFmtString = "'999999999990." & String(intλ��, "9") & "'"
    Else
       GetFmtString = "#0." & String(intλ��, "0") & ";-#0." & String(intλ��, "0") & "; ;"
    End If
End Function

Public Function zlCheckTableIsExsit(ByVal strTableName As String, Optional cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ƿ����
    '���:strTableName-����
    '����:�ɴ淵��true,���򷵻�False
    '����:���˺�
    '����:2018-12-04 10:48:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objDatabase As clsDataBase
    
    On Error GoTo errHandle
    If zlGetOneDataBase(cnOracle, objDatabase) = False Then Exit Function
    strSQL = "Select 1 From All_tables where table_name=[1]"
    Set rsTemp = objDatabase.OpenSQLRecord(strSQL, "�����Ƿ����", strTableName)
    zlCheckTableIsExsit = Not rsTemp.EOF
    Set objDatabase = Nothing
    Exit Function
errHandle:
    If objDatabase.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetOneDataBase(ByRef cnOracle As ADODB.Connection, ByRef objDataBase_Out As Object, Optional ByVal blnIsObjRegisterAlone As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡһ��ͨ���Ӷ���
    '���:cnOracle-���ݿ�����
    '����:objDataBase_Out-�������ݲ�������(�ӿڷ���trueʱ����)
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-12-03 13:55:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gobjOneDataBase Is Nothing Then Set objDataBase_Out = gobjOneDataBase: zlGetOneDataBase = True: Exit Function
    
    On Error GoTo errHandle
    Set gobjOneDataBase = New clsDataBase
    gobjOneDataBase.InitCommon cnOracle, blnIsObjRegisterAlone
    Set objDataBase_Out = gobjOneDataBase
    zlGetOneDataBase = True
    Exit Function
errHandle:
    Exit Function
End Function
Public Function zlGetOneCardDataObject(ByRef cnOracle As ADODB.Connection, ByRef objOneDataObject_Out As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡһ��ͨ���ݷ��ʶ���
    '���:
    '����:objOneDataObject_Out-����һ��ͨ���ݷ��ʶ���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-12-04 14:10:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
   On Error GoTo errHandle
    If Not gobjOneDataObject Is Nothing Then Set objOneDataObject_Out = gobjOneDataObject: zlGetOneCardDataObject = True: Exit Function
    
    Set gobjOneDataObject = New clsOneCardDataObject
    gobjOneDataObject.InitCommon cnOracle
    Set objOneDataObject_Out = gobjOneDataObject
    zlGetOneCardDataObject = True
    Exit Function
errHandle:
    Exit Function
End Function


Public Sub zlInitPublicVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2018-12-03 13:03:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrAviPath = GetSetting("ZLSOFT", "ע����Ϣ", "gstrAviPath", "")
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrHelpPath = gstrAviPath & "\help"
    gstrComputerName = zlGetComputerName
    With gSysColor
        .lngGridColorLost = &HE0E0E0   '�뿪��ɫ
        .lngGridColorSel = &HFFEBD7       'ѡ����ɫ
    End With
    Call ��ʼС��λ��
    
    'ȡվ��
    If gobjComLib Is Nothing Then zlInitCommLib
    If Not gobjComLib Is Nothing And gstrNodeNo = "" Then
        gstrNodeNo = gobjComLib.gstrNodeNo
    End If
    
End Sub
Public Function zlGetComputerName() As String
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ��������
    '������
    '˵����
    '------------------------------------------------------------------------------------------------------------------
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    strComputer = strComputer
    zlGetComputerName = Trim(Replace(strComputer, Chr(0), ""))
End Function

Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ��Ϣ��
    '������strMsgInfor-��ʾ��Ϣ
    '     blnYesNo-�Ƿ��ṩYES��NO��ť
    '���أ�blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub
Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    'clsCommFun���ڸú���
    '���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


Public Function zlCloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    zlCloseWindows = Forms.count = 0
End Function

Public Function zlReleaseResources() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ͷ���Դ
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-02-13 10:30:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'ʵ����Ϊ0ʱ���ŷ���Դ
    If glngInstanceCount > 0 Then Exit Function
    Call zlCloseWindows '�ͷŴ�����Դ
    Set gobjComLib = Nothing: Set gobjCommFun = Nothing: Set gobjDatabase = Nothing
    Set gobjControl = Nothing: Set gobjOneDataBase = Nothing: Set gobjLog = Nothing
    zlReleaseResources = True
End Function

Public Sub zlInitCommLib()
   '��ʼ����������
    If Not gobjComLib Is Nothing Then Exit Sub

    Err = 0: On Error Resume Next
    Set gobjComLib = GetObject("", "zl9Comlib.clsComlib")
    Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    Err = 0: On Error GoTo 0
 End Sub
 
 Public Function zlStringEncode(ByVal strPutString As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ַ�������
    '���:strPutString-��Ҫ���ܵĴ�
    '����:
    '����:���ܴ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If strPutString = "" Then Exit Function
    zlStringEncode = Md5_String_Calc(strPutString)
End Function
Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
    '       ʵ�����ݴ洢����
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function


Public Function SubB(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'����:��ȡָ���ִ���ֵ,�ִ��п��԰�������
 '���:strInfor-ԭ��
 '         lngStart-ֱʼλ��
'         lngLen-����
'����:�Ӵ�
    Err = 0: On Error GoTo errH:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
errH:
    Err.Clear
    SubB = ""
End Function
