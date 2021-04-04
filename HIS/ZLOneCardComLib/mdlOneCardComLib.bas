Attribute VB_Name = "mdlOneCardComLib"
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
Public glngInstanceCount As Long    'ʵ����

'-----------------------------------------------------------------------------------------------------
'���漰����
Public grsҽ�ƿ���� As ADODB.Recordset
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

Private Function GetCardNODencodeRule(ByVal lng�����ID As Long, _
    Optional bln���ѿ� As Boolean = False, Optional cnOracle As ADODB.Connection) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����ID�Ĺ���
    '���:lng�����ID-�����ID
    '        bln���ѿ�-�Ƿ����ѿ�
    '����:�����Ŀ��ű������
    '����:���˺�
    '����:2011-06-22 11:01:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    If bln���ѿ� Then
        Set rsTemp = zlGet���ѿ��ӿ�()
        rsTemp.Filter = "ID=" & lng�����ID
        If rsTemp.EOF Then GoTo GoEnd:
        GetCardNODencodeRule = NVL(rsTemp!�Ƿ�����)
        GoTo GoEnd:
    End If
    Set rsTemp = zlGetҽ�ƿ����()
    rsTemp.Filter = "ID=" & lng�����ID
    If rsTemp.EOF Then GoTo GoEnd:
    GetCardNODencodeRule = NVL(rsTemp!��������)
GoEnd:
    rsTemp.Filter = 0
End Function

Public Function GetCardNODencode(ByVal strCardNo As String, _
    Optional lng�����ID As Long = 0, _
    Optional strRule As String = "", Optional bln���ѿ� As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������
    '���:lng�����ID-�����ID�����ѿ����,�������,����ҽ�ƿ��������ѿ����е�"��������"���Ƿ����Ľ��м���
    '       strRule-����:2-4��ʾ��2λ��4λ��*����,����-��,���ʾ�����λ��ʾΪ*
    '       strCardNo-����
    '����:
    '����:��**�Ŀ���,�������,���ؿ�
    '����:���˺�
    '����:2011-06-21 14:21:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPass As Variant
    Dim strCardPassText As String, i As Long, J As Long
    If bln���ѿ� Then
        If Val(strRule) = 1 Then GetCardNODencode = String(Len(strCardNo), "*"): Exit Function
        If lng�����ID = 0 Then GetCardNODencode = strCardNo: Exit Function
        If Val(GetCardNODencodeRule(lng�����ID, True)) = 1 Then
            GetCardNODencode = String(Len(strCardNo), "*"): Exit Function
        Else
            GetCardNODencode = strCardNo: Exit Function
        End If
    End If
    If lng�����ID <> 0 And strRule = "" Then
        strCardPassText = GetCardNODencodeRule(lng�����ID)
    Else
        'ȡ�Ź���
        strCardPassText = strRule
    End If
    If strCardPassText = "" Then
       GetCardNODencode = strCardNo
    End If
    varPass = Split(strCardPassText & "-", "-")
    If Val(varPass(0)) = 0 Or Val(varPass(1)) = 0 Then
        '���λ��ʾ*
        i = IIf(Val(varPass(0)) = 0, Val(varPass(1)), Val(varPass(0)))
        If i = 0 Then GetCardNODencode = strCardNo: Exit Function
        J = Len(strCardNo) - i: J = IIf(J < 0, 0, J)
        GetCardNODencode = Mid(strCardNo, 1, J) & String(i, "*")
        Exit Function
    End If
    i = Val(varPass(0)): J = Val(varPass(1))
    If i > Len(strCardNo) Then GetCardNODencode = strCardNo: Exit Function
    If J > Len(strCardNo) Then J = Len(strCardNo)
    If J < i Then J = i
   GetCardNODencode = Mid(strCardNo, 1, i - 1) & String(J - i + 1, "*") & Mid(strCardNo, J + 1)
End Function
Public Function GetAvailabilityWriteCardType() As String
       '---------------------------------------------------------------------------------------------------------------------------------------------
        '����:����д�����
        '����:����д�����,����ö��ŷ���
        '����:����д������ID,��:123,232,...
        '����:���˺�
        '����:2013-06-07 10:40:59
        '˵��:
        '---------------------------------------------------------------------------------------------------------------------------------------------
        Dim rsTemp As ADODB.Recordset, strWriteCardIDs As String
        Dim intAutoRead As Integer, intAutoSplitTime As Integer, blnStartCardType As Boolean, str���� As String
        On Error GoTo errHandle
        
        Set rsTemp = zlGetҽ�ƿ����
        rsTemp.Filter = "�Ƿ�д��=1 And �Ƿ�����=1"
        If rsTemp.EOF Then GetAvailabilityWriteCardType = "": Exit Function
        strWriteCardIDs = ""
         With rsTemp
            '���ƿ�(�����ѿ�)
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                ' "����ȫ��\SquareCard\" & mlngCardNo, "�Զ���ȡ"
                intAutoRead = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard\ҽ�ƿ�\" & NVL(!����), "�Զ���ȡ", "0"))
                intAutoSplitTime = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard\ҽ�ƿ�\" & NVL(!����), "�Զ���ȡ���", "300"))
                If Val(NVL(rsTemp!�Ƿ�����)) = 1 Then   '���ƿ�,������
                    blnStartCardType = True
                Else
                    blnStartCardType = Val(GetSetting("ZLSOFT", "����ģ��\zlSquareCard\ҽ�ƿ�\" & NVL(!����), "����", "0")) = 1
                End If
                If blnStartCardType Then
                    strWriteCardIDs = strWriteCardIDs & "," & Val(NVL(rsTemp!id))
                End If
                .MoveNext
            Loop
         End With
         Set rsTemp = Nothing
        If strWriteCardIDs <> "" Then strWriteCardIDs = Mid(strWriteCardIDs, 2)
        GetAvailabilityWriteCardType = strWriteCardIDs
        Exit Function
errHandle:
        If gobjComLib.ErrCenter() = 1 Then
            Resume
        End If
End Function


Public Function GetCardFromCardtypeID(ByVal lngCardTypeID As Long, ByVal bln���ѿ� As Boolean, ByRef objCard As Card) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ�ƿ�����������(�˶���
    '���:lngCardTypeID-�����ID
    '       bln���ѿ�-�Ƿ����ѿ�
    '����:objCard-���ؿ�����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-25 10:28:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, str���� As String
    Dim int�Զ���ȡ As Integer, int�Զ���ȡ��� As Integer, bln���� As Boolean
    Dim objDatabase As New clsDataBase, str�������� As String
    
    On Error GoTo errHandle
    Set objCard = New Card
    If Not bln���ѿ� Then
        Set rsTemp = zlGetҽ�ƿ����
        rsTemp.Filter = "id=" & lngCardTypeID
        If rsTemp.EOF Then rsTemp.Filter = 0: Exit Function
        If Val(NVL(rsTemp!�Ƿ�����)) = 1 Then
            ' "����ȫ��\SquareCard\" & mlngCardNo, "�Զ���ȡ"
            int�Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard\ҽ�ƿ�\" & NVL(rsTemp!����), "�Զ���ȡ", "0"))
            int�Զ���ȡ��� = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard\ҽ�ƿ�\" & NVL(rsTemp!����), "�Զ���ȡ���", "300"))
            If Val(NVL(rsTemp!�Ƿ�����)) = 1 Then   '���ƿ�,������
                bln���� = True
            Else
                '�����:54098
                If (NVL(rsTemp!����) Like "*���֤*" Or NVL(rsTemp!����) Like "*IC��*") And Val(NVL(rsTemp!�Ƿ�̶�)) = 1 And NVL(rsTemp!����) = "" Then
                    bln���� = True
                Else
                    bln���� = Val(GetSetting("ZLSOFT", "����ģ��\zlSquareCard\ҽ�ƿ�\" & NVL(rsTemp!����), "����", "0")) = 1
                End If
            End If
        Else
            bln���� = False
        End If
        str���� = Trim(NVL(rsTemp!����))
        'ID,����,����,����,ǰ׺�ı�,���ų���,ȱʡ��־,�Ƿ�̶�,�Ƿ��ϸ����,�Ƿ�ˢ��,�Ƿ�����,�Ƿ�����ʻ�,�Ƿ�ȫ��,����,��ע,�ض���Ŀ,���㷽ʽ,�Ƿ�����
        Set objCard = New Card
        With objCard
            .�ӿ���� = NVL(rsTemp!id)
            .�ӿڱ��� = NVL(rsTemp!����)
            .���� = NVL(rsTemp!����)
            .���� = NVL(rsTemp!����)
            .ǰ׺�ı� = NVL(rsTemp!ǰ׺�ı�)
            .���ų��� = Val(NVL(rsTemp!���ų���)) + Val(NVL(rsTemp!�豸�Ƿ����ûس�))
            .ȱʡ��־ = Val(NVL(rsTemp!ȱʡ��־)) = 1
            .ϵͳ = Val(NVL(rsTemp!�Ƿ�̶�)) = 1
            .�Ƿ��ϸ���� = Val(NVL(rsTemp!�Ƿ��ϸ����)) = 1
            .�Ƿ��Զ���ȡ = int�Զ���ȡ
            .�Զ���ȡ��� = int�Զ���ȡ���
            .���ƿ� = Val(NVL(rsTemp!�Ƿ�����)) = 1
            .�Ƿ�����ʻ� = Val(NVL(rsTemp!�Ƿ�����ʻ�)) = 1
            .�Ƿ�ȫ�� = Val(NVL(rsTemp!�Ƿ�ȫ��)) = 1
            .�����ظ�ʹ�� = Val(NVL(rsTemp!�Ƿ��ظ�ʹ��)) = 1
            .���㷽ʽ = NVL(rsTemp!���㷽ʽ)
            .�ӿڳ����� = NVL(rsTemp!����)
            .�ض���Ŀ = NVL(rsTemp!�ض���Ŀ)
            .���� = bln����
            .��ע = NVL(rsTemp!��ע)
            .�������Ĺ��� = NVL(rsTemp!��������)
            .�Ƿ����� = Val(NVL(rsTemp!�Ƿ�����)) = 1
            .���볤�� = Val(NVL(rsTemp!���볤��))
            .���볤������ = Val(NVL(rsTemp!���볤������))
            .������� = Val(NVL(rsTemp!�������))
            .������������ = Val(NVL(rsTemp!������������))
            .�Ƿ�ȱʡ���� = Val(NVL(rsTemp!�Ƿ�ȱʡ����)) = 1
            .�Ƿ��ƿ� = Val(NVL(rsTemp!�Ƿ��ƿ�)) = 1   '56615
            .�Ƿ񷢿� = Val(NVL(rsTemp!�Ƿ񷢿�)) = 1 Or .���ƿ�
            .�Ƿ�д�� = Val(NVL(rsTemp!�Ƿ�д��)) = 1
            .�������� = Val(NVL(rsTemp!��������))
            .�Ƿ�ת�ʼ����� = Val(NVL(rsTemp!�Ƿ�ת�ʼ�����)) = 1
            str�������� = NVL(rsTemp!��������, "1000")
            .�Ƿ�ˢ�� = Mid(str��������, 1, 1) = 1
            .�Ƿ�ɨ�� = Mid(str��������, 2, 1) = 1
            .�Ƿ�Ӵ�ʽ���� = Mid(str��������, 3, 1) = 1
            .�Ƿ�ǽӴ�ʽ���� = Mid(str��������, 4, 1) = 1
            .�Ƿ�ֿ����� = Val(NVL(rsTemp!�Ƿ�ֿ�����)) = 1
            .�Ƿ��˿��鿨 = Val(NVL(rsTemp!�Ƿ��˿��鿨)) = 1
            .�Ƿ�֤�� = Val(NVL(rsTemp!�Ƿ�֤��)) = 1
            .�豸�Ƿ����ûس� = Val(NVL(rsTemp!�豸�Ƿ����ûس�)) = 1
            .�Ƿ�ȱʡ���� = Val(NVL(rsTemp!�Ƿ�ȱʡ����)) = 1
            .�Ƿ�������� = Val(NVL(rsTemp!�Ƿ��������)) = 1
            .�Ƿ�֧��ɨ�븶 = Val(NVL(rsTemp!�Ƿ�֧��ɨ�븶)) = 1
        End With
        rsTemp.Filter = 0:
       GetCardFromCardtypeID = True
        Exit Function
    End If
    
    
    Set rsTemp = zlGet���ѿ��ӿ�
    rsTemp.Filter = "ID=" & lngCardTypeID
    If rsTemp.EOF Then Set rsTemp.Filter = 0: Exit Function
    
    With rsTemp
        '���ƿ�(�����ѿ�)
        If .RecordCount <> 0 Then .MoveFirst
        bln���� = Val(NVL(!����)) = 1
        If bln���� Then
            ' "����ȫ��\SquareCard\" & mlngCardNo, "�Զ���ȡ"
            int�Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard\" & NVL(!���), "�Զ���ȡ", "0"))
            int�Զ���ȡ��� = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard\" & NVL(!���), "�Զ���ȡ���", "300"))
            bln���� = Val(GetSetting("ZLSOFT", "����ģ��\zlSquareCard\" & NVL(!���), "����", "0")) = 1
        End If
        '���,����,���㷽ʽ,nvl(���ƿ�,0)  as ���ƿ�,ǰ׺�ı�,���ų���,����,ϵͳ,�Ƿ�����
        str���� = Trim(NVL(rsTemp!����))
        Set objCard = New Card
        With objCard
            .�ӿ���� = NVL(rsTemp!���)
            .�ӿڱ��� = NVL(rsTemp!���)
            .���� = Left(NVL(rsTemp!����), 1)   'Ĭ��ȡ��һ��
            .���� = NVL(rsTemp!����)
            .ǰ׺�ı� = NVL(rsTemp!ǰ׺�ı�)
            .���ų��� = Val(NVL(rsTemp!���ų���))
            .ϵͳ = Val(NVL(rsTemp!ϵͳ)) = 1
            .�Ƿ��ϸ���� = False
            .�Ƿ��Զ���ȡ = int�Զ���ȡ
            .�Զ���ȡ��� = int�Զ���ȡ���
            .���ƿ� = Val(NVL(rsTemp!���ƿ�)) = 1
            .�Ƿ�����ʻ� = True 'Not (Val(Nvl(rsTemp!���ƿ�)) = 1)
            .�Ƿ�ȫ�� = Val(NVL(rsTemp!�Ƿ�ȫ��)) = 1
            .���㷽ʽ = NVL(rsTemp!���㷽ʽ)
            .�ӿڳ����� = NVL(rsTemp!����)
            .�ض���Ŀ = ""
            .���� = bln����
            .�����ظ�ʹ�� = True
            .��ע = ""
            .�������Ĺ��� = NVL(rsTemp!�Ƿ�����)
            .���ѿ� = True
            .�Ƿ����� = Val(NVL(rsTemp!�Ƿ�����)) = 1
            .���볤�� = Val(NVL(rsTemp!���볤��))
            .���볤������ = Val(NVL(rsTemp!���볤������))
            .������� = Val(NVL(rsTemp!�������))
            .������������ = Val(NVL(rsTemp!������������))
            .�Ƿ�ȱʡ���� = Val(NVL(rsTemp!�Ƿ�ȱʡ����)) = 1
            .�Ƿ��ƿ� = Val(NVL(rsTemp!�Ƿ��ƿ�)) = 1   '56615
            .�Ƿ񷢿� = Val(NVL(rsTemp!�Ƿ񷢿�)) = 1 Or .���ƿ�
            .�Ƿ�д�� = Val(NVL(rsTemp!�Ƿ�д��)) = 1
            .�������� = Val(NVL(rsTemp!��������))
            str�������� = NVL(rsTemp!��������, "1000")
            .�Ƿ�ˢ�� = Mid(str��������, 1, 1) = 1
            .�Ƿ�ɨ�� = Mid(str��������, 2, 1) = 1
            .�Ƿ�Ӵ�ʽ���� = Mid(str��������, 3, 1) = 1
            .�Ƿ�ǽӴ�ʽ���� = Mid(str��������, 4, 1) = 1
        End With
    End With
    Set rsTemp.Filter = 0:
    GetCardFromCardtypeID = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
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
    Set gobjComLib = Nothing: Set gobjCommFun = Nothing
    Set gobjDatabase = Nothing: Set gobjControl = Nothing
    Set gobjLog = Nothing: Set gobjOneDataBase = Nothing
    Set grs���ѿ��ӿ� = Nothing: Set grsҽ�ƿ���� = Nothing
    Set gcnOracle = Nothing
    Set gobjOneDataObject = Nothing
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
    Err = 0: On Error GoTo ErrH:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
ErrH:
    Err.Clear
    SubB = ""
End Function
Public Function zlGetCardTypeRecStru(ByRef rsCardType As ADODB.Recordset) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����ṹ
     '����:rsCardType-���صļ�¼���ṹ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-24 18:10:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsCardType = New ADODB.Recordset
    With rsCardType
        If .State = 1 Then .Close
        'adBigInt
        .fields.Append "ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .fields.Append "����", adLongVarChar, 200, adFldIsNullable
        .fields.Append "����", adLongVarChar, 50, adFldIsNullable
        
        .fields.Append "ǰ׺�ı�", adLongVarChar, 30, adFldIsNullable
        .fields.Append "���ų���", adSmallInt, 20, adFldIsNullable
        .fields.Append "ȱʡ��־", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ�̶�", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ��ϸ����", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ�����", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ�����ʻ�", adSmallInt, , adFldIsNullable
        
        .fields.Append "�Ƿ�����", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ�ȱʡ����", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ�ȫ��", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ��ظ�ʹ��", adSmallInt, , adFldIsNullable
        .fields.Append "��������", adSmallInt, , adFldIsNullable
        .fields.Append "���볤��", adSmallInt, , adFldIsNullable
        .fields.Append "���볤������", adSmallInt, , adFldIsNullable
        .fields.Append "�������", adSmallInt, , adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "��ע", adLongVarChar, 300, adFldIsNullable
        .fields.Append "�ض���Ŀ", adLongVarChar, 100, adFldIsNullable
        .fields.Append "���㷽ʽ", adLongVarChar, 50, adFldIsNullable
        .fields.Append "�Ƿ�����", adSmallInt, , adFldIsNullable
        .fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        
        .fields.Append "������������", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ�ȱʡ����", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ�ģ������", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ��ƿ�", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ�д��", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ񷢿�", adSmallInt, , adFldIsNullable
        .fields.Append "��������", adSmallInt, , adFldIsNullable
        
        
        .fields.Append "��������", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ�֤��", adSmallInt, , adFldIsNullable
        
        
        .fields.Append "�Ƿ�ת�ʼ�����", adSmallInt, , adFldIsNullable
        .fields.Append "��������", adLongVarChar, 20, adFldIsNullable
        .fields.Append "�Ƿ�ֿ�����", adSmallInt, , adFldIsNullable
        .fields.Append "���͵��ýӿ�", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ��˿��鿨", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ��������", adSmallInt, , adFldIsNullable
        .fields.Append "ȱʡ��Чʱ��", adLongVarChar, 50, adFldIsNullable
        .fields.Append "����ʶ�����", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ�֧��ɨ�븶", adSmallInt, , adFldIsNullable
        .fields.Append "����", adSmallInt, , adFldIsNullable
        .fields.Append "���̿��Ʒ�ʽ", adSmallInt, , adFldIsNullable
        .fields.Append "�Ƿ����ûس�", adSmallInt, , adFldIsNullable
    
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    zlGetCardTypeRecStru = True
End Function

Public Function zlGetQueryPatiInforStru(ByRef rsPati As ADODB.Recordset, _
    Optional ByVal bytPatiInfoShowType As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ѯ�Ĳ�����Ϣ���ݼ�
    '��Σ�
    '       bytPatiInfoShowType-����ѡ��������ʾ��ʽ��0-������Ϣ��1-������Ϣ
    '����:rsPati-���صļ�¼���ṹ
    '����:�ɹ�����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsPati = New ADODB.Recordset
    With rsPati
        .fields.Append "ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "��Ժ", adLongVarChar, 4, adFldIsNullable
        .fields.Append "����ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "�Ա�", adLongVarChar, 20, adFldIsNullable
        .fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .fields.Append "���֤��", adLongVarChar, 30, adFldIsNullable
        .fields.Append "IC����", adLongVarChar, 50, adFldIsNullable
        .fields.Append "�����", adLongVarChar, 18, adFldIsNullable
        .fields.Append "סԺ��", adLongVarChar, 18, adFldIsNullable
        .fields.Append "�ֻ���", adLongVarChar, 40, adFldIsNullable
        
        If bytPatiInfoShowType = 0 Then
            .fields.Append "��������", adLongVarChar, 30, adFldIsNullable
            .fields.Append "�����ص�", adLongVarChar, 200, adFldIsNullable
            .fields.Append "�ѱ�", adLongVarChar, 50, adFldIsNullable
            .fields.Append "ҽ�Ƹ��ʽ", adLongVarChar, 100, adFldIsNullable
            .fields.Append "����", adLongVarChar, 30, adFldIsNullable
            .fields.Append "��ͥ��ַ", adLongVarChar, 200, adFldIsNullable
            .fields.Append "��ͥ�绰", adLongVarChar, 50, adFldIsNullable
            .fields.Append "��ϵ������", adLongVarChar, 100, adFldIsNullable
            .fields.Append "��ϵ�˹�ϵ", adLongVarChar, 50, adFldIsNullable
            .fields.Append "��ϵ�˵绰", adLongVarChar, 100, adFldIsNullable
            .fields.Append "����Ԥ�����", adDouble, , adFldIsNullable
            .fields.Append "סԺԤ�����", adDouble, , adFldIsNullable
            .fields.Append "����", adLongVarChar, 200, adFldIsNullable
            .fields.Append "����ID", adLongVarChar, 100, adFldIsNullable
         End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    zlGetQueryPatiInforStru = True
End Function

Public Function GetOneCardTypes(ByRef rsTypes_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���»�ȡҽ�ƿ�������ݼ�
    '������true,���򷵻�False
    '����:���˺�
    '����:2019-11-21 15:08:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cllData As Collection, cllTemp As Variant
 
    On Error GoTo errHandle
    

    If zlGetCardTypeRecStru(rsTypes_Out) = False Then Exit Function
    If zl_PatiSvr_GetCardTypes(cllData) = False Then Exit Function
    
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then Exit Function
    
    'output
    '    cardtype_id N   1   ID
    '    cardtype_code   C   1   ����
    '    cardtype_name   C   1   ����
    '    cardtype_stname C   1   ����
    '    prefix_text C   1   ǰ׺�ı�
    '    cardno_len  N   1   ���ų���
    '    default    N   1   ȱʡ��־
    '    fixed N   1   �Ƿ�̶�:1-��ϵͳ�̶�;0-����ϵͳ�̶�
    '    strict   N   1   �Ƿ��ϸ����:1-���ϸ����;0-�����ϸ����
    '    self_make N   1   �Ƿ�����:1-�ǵ�;0-����
    '    exist_account  N   1   �Ƿ�����ʻ�:1-�����ʻ�;0-�������˻�
    '    allow_return_cash    N   1   �Ƿ�����:1-����;0-������
    '    must_all_return   N   1   �Ƿ�ȫ��:1-����ȫ��;0-��������
    '    component   C   1   ����
    '    memo    C   1   ��ע
    '    spec_item   C   1   �ض���Ŀ
    '    blnc_mode   C   1   ���㷽ʽ
    '    blnc_nature N   1   ��������
    '    cardno_pwdtxt   C   1   ��������:���Ŵӵڼ�λ���ڼ�λ��ʾ����,��ʽΪ:S-N:S��ʾ�ӵڼ�λ��ʼ,���ڼ�λ����.����:3-10,��ʾ��3λ��10λ������*��ʾ:12********3323��Ҫ����Ӧ��ͬ����ҽ�ƿ�
    '    allow_repeat_use N   1   �Ƿ��ظ�ʹ��:1-����;0-������
    '    enabled    N   1   �Ƿ�����:1-������;0-δ����
    '    pwd_len N   1   ���볤��
    '    pwd_len_limit   N   1   ���볤������:0-��������;1-�̶����볤��;-n��ʾ�����������ö��λ��������,�����ܳ������볤��
    '    pwd_rule    N   1   �������:��-���ֺ��ַ����;1-��Ϊ�������
    '    allow_vaguefind    N   1   �Ƿ�ģ������:1-֧��ģ������;0-��֧��
    '    pwd_require    N   1   ������������:0-������;1-������,����;2-�������ֹ;ȱʡΪ������
    '    default_pwd  N   1   �Ƿ�ȱʡ����:1-�����֤��N(�����볤��Ϊ׼)λ��Ϊȱʡ����;0-��ȱʡ����
    '    allow_makecard N   1   �Ƿ��ƿ�:1-��;0-��
    '    allow_sendcard N   1   �Ƿ񷢿�:1-��;0-��
    '    allow_writecard    N   1   �Ƿ�д��:1-��;0-��
    '    insurance_type  N   1   ����
    '    sendcard_nature N   1   ��������:0-������;1-ͬһ����ֻ�ܷ�һ�ſ�;2-ͬһ�����������ſ���������ʾ;ȱʡΪ0
    '    allow_transfer N   1   �Ƿ�ת�ʼ�����:1-֧��ת�ʼ�����;0-��֧��
    '    readcard_nature C   1   ��������,ҽ�ƿ�������ʽ����һλΪ:�Ƿ�ˢ��;�ڶ�λΪ�Ƿ�ɨ��;����λ�Ƿ�Ӵ�ʽ����;����λ�Ƿ�ǽӴ�ʽ����������ˢ����'1000'
    '    keyboard_mode   N   1   ���̿��Ʒ�ʽ:��0-��ֹʹ�������;1-ʹ����������� ,2-ʹ���ַ������
    '    advsend_buildqrcode N   1   �Ƿ�ҽ�����͵����������ɽӿ�:1-���͵������ɶ�ά��ӿ�;0-������
    '    holding_pay   N   1   �Ƿ�ֿ�����:1-��;0-��
    '    cert_cardtype    N   1   �Ƿ�֤�����͵�ҽ�ƿ�:0-���ǣ�1-��
    '    verfycard    N   1   �Ƿ��˿��鿨
    '    sendcard_sign   N   1   ��������:0��NULL-����ʱ�����ű���ﵽ���ų���;1-����ʱ��������С�ڵ��ڿ��ų���,����ʱ��С�ڿ��ų���ʱ������ʾ����Ա;2-����ʱ��������С�ڵ��ڿ��ų���,С��ʱ����ʾ����Ա��
    '    enterkey_enabled N   1   �豸�Ƿ����ûس�:ҽ�ƿ���Ӧ��ˢ���豸�Ƿ������˻س�����������˻س����򿨺ų���Ĭ������һλ�����λس�
    '    def_return_cash N   1   �Ƿ�ȱʡ����:��������ʱ,Ĭ���Ƿ�����
    '    balalone N   1   �Ƿ��������:1-��������;0-�Ƕ�������
    '    discern_rule    N   1   ����ʶ�����:1-ȫ��ת��Ϊ��д;0-�����ִ�Сд
    '    def_valid_time  C   1   ȱʡ��Чʱ��:NULLʱ����ʾ������;�ǿ�ʱ����ʽΪ:ʱ��+��λ(�죬��),���磺3��,3��
    '    scanpay  N   1   �Ƿ�֧��ɨ�븶:�Ƿ�֧��ɨ�븶,֧��ʱ������á�zlReadQRCode������

    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        With rsTypes_Out
            .AddNew
                !id = cllTemp("_cardtype_id")
                !���� = cllTemp("_cardtype_code")
                !���� = cllTemp("_cardtype_name")
                !���� = cllTemp("_cardtype_stname")
                
                !ǰ׺�ı� = cllTemp("_prefix_text")
                !���ų��� = cllTemp("_cardno_len")
                !ȱʡ��־ = cllTemp("_default")
                !�Ƿ�̶� = cllTemp("_fixed")
                
                !�Ƿ��ϸ���� = cllTemp("_strict")
                !�Ƿ����� = cllTemp("_self_make")
                !�Ƿ�����ʻ� = cllTemp("_exist_account")
                !�Ƿ����� = cllTemp("_allow_return_cash")
                !�Ƿ�ȱʡ���� = cllTemp("_def_return_cash")
                !�Ƿ�ȫ�� = cllTemp("_must_all_return")
                !���� = cllTemp("_component")
                !��ע = cllTemp("_memo")
                
                !�ض���Ŀ = cllTemp("_spec_item")
                !���㷽ʽ = cllTemp("_blnc_mode")
                !�������� = cllTemp("_cardno_pwdtxt")
                  
                !�Ƿ��ظ�ʹ�� = cllTemp("_allow_repeat_use")
                !�Ƿ����� = cllTemp("_enabled")
                
                !���볤�� = cllTemp("_pwd_len")
                !���볤������ = cllTemp("_pwd_len_limit")
                !������� = cllTemp("_pwd_rule")
                !������������ = cllTemp("_pwd_require")
                
                !�Ƿ�ģ������ = cllTemp("_allow_vaguefind")
                !�Ƿ�ȱʡ���� = cllTemp("_default_pwd")
                
                
                
                !�Ƿ��ƿ� = cllTemp("_allow_makecard")
                !�Ƿ񷢿� = cllTemp("_allow_sendcard")
                !�Ƿ�д�� = cllTemp("_allow_writecard")
                !�������� = cllTemp("_sendcard_sign")
                
                !�������� = cllTemp("_blnc_nature")
                !���� = cllTemp("_insurance_type")
                !�������� = cllTemp("_sendcard_nature")
                !�Ƿ�ת�ʼ����� = cllTemp("_allow_transfer")
                !�������� = cllTemp("_readcard_nature")
                !���̿��Ʒ�ʽ = cllTemp("_keyboard_mode")
                
                !�Ƿ�ֿ����� = cllTemp("_holding_pay")
                !�Ƿ�֤�� = cllTemp("_cert_cardtype")
                
                !���͵��ýӿ� = cllTemp("_advsend_buildqrcode")
                !�Ƿ��˿��鿨 = cllTemp("_verfycard")
                                
                !�Ƿ�������� = cllTemp("_balalone")
                !ȱʡ��Чʱ�� = cllTemp("_def_valid_time")
                !����ʶ����� = cllTemp("_discern_rule")
                !�Ƿ�֧��ɨ�븶 = cllTemp("_scanpay")
                !�Ƿ����ûس� = cllTemp("_enterkey_enabled")
                
            .Update
       End With
    Next
    If rsTypes_Out.RecordCount <> 0 Then rsTypes_Out.MoveFirst
    GetOneCardTypes = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetҽ�ƿ����() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ�ƿ����
    '����:����ҽ�ƿ����ļ�¼��
    '����:���˺�
    '����:2011-05-23 17:25:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cllData As Collection, cllTemp As Variant
    
    On Error GoTo errHandle
    
    If Not grsҽ�ƿ���� Is Nothing Then
        If grsҽ�ƿ����.State = 1 Then
            grsҽ�ƿ����.Filter = 0
            If grsҽ�ƿ����.RecordCount <> 0 Then grsҽ�ƿ����.MoveFirst
            Set zlGetҽ�ƿ���� = grsҽ�ƿ����
            Exit Function
        End If
    End If
    If GetOneCardTypes(grsҽ�ƿ����) = False Then Set grsҽ�ƿ���� = Nothing: Exit Function
    Set zlGetҽ�ƿ���� = grsҽ�ƿ����
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Set grsҽ�ƿ���� = Nothing
End Function
Public Function GetPatiSurplusFromPatiID(ByVal lng����ID As Long, ByRef dbl����Ԥ�����_out As Double, ByRef dblסԺԤ�����_Out As Double, _
    ByRef dbl����������_Out As Double, ByRef dblסԺ�������_Out As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID��ȡ���˷��ü�Ԥ�����
    '���:lng����ID-����ID
    '
    '����:dbl����Ԥ�����_out
    '     dblסԺԤ�����_Out
    '     dbl����������_Out
    '     dblסԺ�������_Out
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-05 20:57:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllData As Collection, cllTemp As Collection
    
    dbl����Ԥ�����_out = 0
    dblסԺԤ�����_Out = 0
    
    dbl����������_Out = 0
    dblסԺ�������_Out = 0
    If zl_ExseSvr_GetPatiSurplusInfo(lng����ID, cllData) = False Then Exit Function
    Set cllTemp = zlGetNodeObjectFromCollect(cllData, "_" & lng����ID)
    If cllTemp Is Nothing Then GetPatiSurplusFromPatiID = True: Exit Function
    dbl����Ԥ�����_out = Val(NVL(mdlPubJson.zlGetNodeValueFromCollect(cllTemp, "_outdpst_surplus", "C")))
    dblסԺԤ�����_Out = Val(NVL(mdlPubJson.zlGetNodeValueFromCollect(cllTemp, "_indpst_surplus", "C")))
    
    dbl����������_Out = Val(NVL(mdlPubJson.zlGetNodeValueFromCollect(cllTemp, "_outfee_surplus", "C")))
    dblסԺ�������_Out = Val(NVL(mdlPubJson.zlGetNodeValueFromCollect(cllTemp, "_infee_surplus", "C")))
    GetPatiSurplusFromPatiID = True
End Function
