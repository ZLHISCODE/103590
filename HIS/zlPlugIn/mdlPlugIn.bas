Attribute VB_Name = "mdlPlugIn"
Option Explicit

Public gcnOracle As ADODB.Connection
Public gblnInited As Boolean
Public gcolPlugIn As Collection '��չ�������ü��Ϸ�ʽ�ݴ���չ�������ʵ��


Public Enum Enum_Modue 'ģ���
    m����ҽ��ģ�� = 1252
    mסԺҽ��ģ�� = 1253
    mסԺ��ʿվģ�� = 1254
    m�ٴ�·��ģ�� = 1256
    m����ģ�� = 1070
    m��Ա����ģ�� = 1002
    mҽ������ģ�� = 1257
    
    m����ҽ������վ = 1260
    mסԺҽ������վ = 1261
    mסԺ��ʿ����վ = 1262
    mҽ������վ = 1263
    
    m����ҽ���´� = 1252
    mסԺҽ���´� = 1253
    m�°滤ʿվ = 1265
    
    'Ѫ�����ģ��
    m������Ѫ���� = 1935
    '������Ѫ����ҳǩû��ģ��ţ�������ű�ʾ��Ӧҳǩ
    m������Ѫ����_���渴�� = 193501
    m������Ѫ����_��Ѫ��¼ = 193502
    mѪҺĿ¼���� = 1900
    mѪҺ��Ѫ��Ӧ = 1938
    m���ҷ�Ѫ���� = 1936
    mѪҺ��Ӧ��� = 1915
    mѪҺ���ϳ��� = 1922
    
    'LIS���ģ��
    m�ٴ�ʵ���ҹ��� = 2500
    
    m������Ϣ���� = 1101
    m���Ӳ�����ӡ = 1566
    
    'ҩƷ����ģ��
    m�����⹺������ = 1712
    mҩƷ�⹺������ = 1300
    mҩƷ�ƻ����� = 1330
    mҩƷ������ҩ = 1341
    mҩƷ���ŷ�ҩ = 1342
    m��Һ�������� = 1345
    
    m���˹Һ�ģ�� = 1111
    m�����շ�ģ�� = 1121
    m���˽��ʴ��� = 1137
    
    m������Ĺ��� = 2121
    m����ܼ�Ǽ� = 2125
    m���ֿ�ִ�� = 2122
    m������Ǽ� = 2123
End Enum

'���µ����д���������֧�ֶ���ͬʱ�ҽ�
'CkeckUseable ������������ʹ�ã���д��չ���ʱֱ�Ӵ���ǰ ��λ���Ƽ��ɡ�ʹ�ø÷���ʱҪ���� zl9ComLib.dll
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ע���ؼ��ְ�ȫѡ��...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Const ERROR_NO_MORE_ITEMS = 259&
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_LOCAL_MACHINE = &H80000002

'֧�ֻ��ֵĳ���
Public Const WM_MOUSEWHEEL = &H20A
Public Const GWL_WNDPROC = -4

' ע�����������...
Public Enum ValueType
    REG_SZ = 1                         ' �ַ���ֵ
    REG_EXPAND_SZ = 2                  ' �������ַ���ֵ
    REG_BINARY = 3                     ' ������ֵ
    REG_DWORD = 4                      ' DWORDֵ
    REG_MULTI_SZ = 7                   ' ���ַ���ֵ
End Enum

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private marrName As Variant
Public glngPreHWnd As Long '����֧�������ֹ���
Public gobjMec As Object '��ҳ��������


'��¼����ر���
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Private Sub GetPathNames()
'���ܣ���ȡע���CLSID�¼�Ŀ¼

    Dim hKey As Long, Cnt As Long, sName As String, sData As String, Ret As Long, RetData As Long
    Const BUFFER_SIZE As Long = 255
    marrName = Array()
    Ret = BUFFER_SIZE
    If RegOpenKey(HKEY_CLASSES_ROOT, "CLSID", hKey) = 0 Then
        sName = Space(BUFFER_SIZE)
        While RegEnumKeyEx(hKey, Cnt, sName, Ret, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve marrName(UBound(marrName) + 1)
            marrName(UBound(marrName)) = "CLSID\" & Left$(sName, Ret)
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            Ret = BUFFER_SIZE
        Wend
        RegCloseKey hKey
    End If
    Cnt = 0
End Sub

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, ValueName As String, Optional ValueType As Long) As String
'���ܣ�����Ѵ��ڵ�ע���ؼ��ֵ�ֵ
'������ValueName="" �򷵻� KeyName ���Ĭ��ֵ
'      ���ָ����ע���ؼ��ֲ�����, �򷵻ؿմ�
'      KeyRoot--������, KeyName--��������, ValueName--ֵ������, ValueType--ֵ������
    Dim i As Integer
    Dim hKey As Long
    Dim TempValue As String                             ' ע���ؼ��ֵ���ʱֵ
    Dim Value As String                                 ' ע���ؼ��ֵ�ֵ
    Dim ValueSize As Long                               ' ע���ؼ��ֵ�ֵ��ʵ�ʳ���
    TempValue = Space(1024)                             ' �洢ע���ؼ��ֵ���ʱֵ�Ļ�����
    ValueSize = 1024                                    ' ����ע���ؼ��ֵ�ֵ��Ĭ�ϳ���
    
    ' ��һ���Ѵ��ڵ�ע���ؼ���...
    RegOpenKeyEx KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey
    
    ' ����Ѵ򿪵�ע���ؼ��ֵ�ֵ...
    RegQueryValueEx hKey, ValueName, 0, ValueType, ByVal TempValue, ValueSize
    
    ' ����ע���ؼ��ֵĵ�ֵ...
    Select Case ValueType                                                        ' ͨ���жϹؼ��ֵ�����, ���д���
        Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ
            TempValue = Left$(TempValue, ValueSize - 1)                          ' ȥ��TempValueβ���ո�
            Value = TempValue
        Case REG_DWORD
            ReDim dValue(3) As Byte
            RegQueryValueEx hKey, ValueName, 0, REG_DWORD, dValue(0), ValueSize
            For i = 3 To 0 Step -1
                Value = Value + String(2 - Len(Hex(dValue(i))), "0") + Hex(dValue(i))   ' ���ɳ���Ϊ8��ʮ�������ַ���
            Next i
            If CDbl("&H" & Value) < 0 Then                                              ' ��ʮ�����Ƶ� Value ת��Ϊʮ����
                Value = 2 ^ 32 + CDbl("&H" & Value)
            Else
                Value = CDbl("&H" & Value)
            End If
        Case REG_BINARY
            If ValueSize > 0 Then
                ReDim bValue(ValueSize - 1) As Byte                                     ' �洢 REG_BINARY ֵ����ʱ����
                RegQueryValueEx hKey, ValueName, 0, REG_BINARY, bValue(0), ValueSize
                For i = 0 To ValueSize - 1
                    Value = Value + String(2 - Len(Hex(bValue(i))), "0") + Hex(bValue(i)) + " "  ' ������ת�����ַ���
                Next i
            End If
    End Select
    
    ' �ر�ע���ؼ���...
    RegCloseKey hKey
    GetKeyValue = Trim(Value)                                                    ' ���غ���ֵ
End Function

Private Function GetAllPlugIns() As String
'���ܣ���ȡ��չ����Ĳ������ƣ����Ÿ
    Dim strTmp As String
    Dim strName As String
    Dim strResult As String
    Dim i As Integer
    
    Call GetPathNames
    
    For i = 1 To UBound(marrName)
        strResult = GetKeyValue(HKEY_CLASSES_ROOT, CStr(marrName(i)), strTmp, REG_SZ)
        '��ZLPLUGIN��ͷ
        If UCase(Left(strResult, 8)) = "ZLPLUGIN" Then
            If InStr(strResult, ".") > 0 Then
                If Len(Split(strResult, ".")(0)) > 8 And InStr(strName, Split(strResult, ".")(0)) = 0 Then
                    strName = IIf(strName = "", "", strName & ",") & Split(strResult, ".")(0)
                End If
            End If
        End If
    Next
    GetAllPlugIns = strName
End Function

Public Function HandlePlugIn(ByVal bytType As Byte, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal cnOracle As ADODB.Connection, _
        Optional ByVal int���� As Integer = -1, Optional strReserve As String, Optional strFuncName As String, Optional ByVal lngPatiID As Long, _
        Optional ByVal varRecId As Variant, Optional ByVal varKeyId As Variant)
'���ܣ���չ�������֧����ش���
'������bytType �������� 1=��ʼ����2=��ȡ��������3=ִ�й��ܣ�4=��ֹ����bytType=2ʱ strFunName��Ϊ����
'      cnOracle=�����
'      lngSys,lngModual=��ǰ���ýӿڵ��ϼ�ϵͳ�ż�ģ���
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      strReserve=��������,������չʹ��
'      strFunName ���κ���� ��bytType=2ʱ���Σ���bytType=3ʱ���
'      lngPatiID=��ǰ����ID
'      varRecId=���ֻ����ַ����������ﲡ�ˣ�Ϊ��ǰ�Һŵ��Ż��߹Һ�ID����סԺ���ˣ�Ϊ��ǰסԺ��ҳID
'      varKeyId=���ֻ����ַ�������ǰ�Ĺؼ�ҵ������Ψһ��ʶID����ҽ��ID
    Dim strTmp As String
    Dim strFuncNameTmp As String
    Dim strUserName As String
    Dim objTmp As Object
    Dim varArr As Variant
    Dim i As Integer
    Dim strTmpReserve As String
    Dim strReserveOther As String
    
    On Error Resume Next
    
    If bytType = 1 Then
        strTmp = GetAllPlugIns
        If strTmp = "" Then Exit Function
        varArr = Split(strTmp, ",")
        Set gcolPlugIn = New Collection
        For i = 0 To UBound(varArr)
            Set objTmp = CreateObject(varArr(i) & ".clsPlugIn")
            If Not objTmp Is Nothing Then
                Call objTmp.Initialize(cnOracle, lngSys, lngModual, int����)
                '����ʹ�����ƣ��û�����ʱ��ʾ������
                strUserName = objTmp.GetUserName 'ҽԺ�û�--��λ����
                
                If strUserName <> "" Then
                    If CkeckUseable(strUserName) Then
                        gcolPlugIn.Add objTmp, "_" & varArr(i)
                    End If
                Else
                    gcolPlugIn.Add objTmp, "_" & varArr(i)
                End If
            End If
            Set objTmp = Nothing
        Next i
    End If
    
    If gcolPlugIn Is Nothing Then Exit Function
    
    If bytType = 2 Then
        For i = 1 To gcolPlugIn.Count
            Set objTmp = gcolPlugIn.Item(i)
            strTmp = ""
            strTmpReserve = ""
            strTmp = objTmp.GetFuncNames(lngSys, lngModual, int����, strTmpReserve)
            strFuncNameTmp = IIf(strFuncNameTmp = "", "", strFuncNameTmp & ",") & strTmp
            strReserveOther = IIf(strReserveOther = "", "", strReserveOther & ",") & strTmpReserve
        Next i
        strFuncName = strFuncNameTmp
        strReserve = strReserveOther
    ElseIf bytType = 3 Then
        For i = 1 To gcolPlugIn.Count
            Set objTmp = gcolPlugIn.Item(i)
            Call objTmp.ExecuteFunc(lngSys, lngModual, strFuncName, lngPatiID, varRecId, varKeyId, strReserve, int����)
        Next i
    ElseIf bytType = 4 Then
        For i = 1 To gcolPlugIn.Count
            Set objTmp = gcolPlugIn.Item(i)
            Call objTmp.Terminate(lngSys, lngModual, int����)
        Next i
    End If
    Err.Clear: On Error GoTo 0
End Function

Public Function GetFormCaptionEx(ByVal lngSys As Long, ByVal lngModual As Long) As String
'��ȡ��չ�����еĿ�Ƭ���ƣ�Ҫ��ÿ����չ�����Լ�������֮�俨Ƭ��������
    Dim i As Integer
    Dim objTmp As Object
    Dim strTmp As String
    Dim strCaption As String
    
    If gcolPlugIn Is Nothing Then Exit Function
    On Error Resume Next
    For i = 1 To gcolPlugIn.Count
        Set objTmp = gcolPlugIn.Item(i)
        strTmp = ""
        strTmp = objTmp.GetFormCaption(lngSys, lngModual)
        strCaption = IIf(strCaption = "", "", strCaption & ",") & strTmp
    Next i
    GetFormCaptionEx = strCaption
    Err.Clear: On Error GoTo 0
End Function

Public Function GetFormEx(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strName As String) As Object
'��ȡ��չ�����еĿ�Ƭ������Ϊ��Ƭ���Ʋ���������ÿ�ε���ֻ�᷵��һ�����󣬻��߲�����
    Dim i As Integer
    Dim objTmp As Object
    Dim objForm As Object
    Dim strTmp As String
    Dim strCaption As String
    
    If gcolPlugIn Is Nothing Then Exit Function
    On Error Resume Next
    For i = 1 To gcolPlugIn.Count
        Set objTmp = gcolPlugIn.Item(i)
        Set objForm = objTmp.GetForm(lngSys, lngModual, strName)
        If Not objForm Is Nothing Then Exit For
    Next i
    Err.Clear: On Error GoTo 0
    
    Set GetFormEx = objForm
End Function

Private Function CkeckUseable(ByVal str��λ���� As String) As Boolean
'���ܣ���չ���ʹ������ʾ������
'������ʹ�õ�λ��ȫ��
    Dim strTmp As String
    
    strTmp = zlRegInfo("��λ����", , 0)
    If strTmp = "" Then Exit Function
    If InStr("," & strTmp & ",", "," & str��λ���� & ",") > 0 Then CkeckUseable = True
End Function



'----------------------------------------------------------------------------------------------------------------------------
'��¼���������
'----------------------------------------------------------------------------------------------------------------------------
Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
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

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '���¼�¼,���������,������
    'strPrimary:�ֶ���|ֵ
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'strPrimary = "RecordID|5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '��λ��ָ����¼
    'strPrimary:����,ֵ
    'blnDelete=True,��ü�¼������"ɾ��"�ֶ�
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !ɾ�� = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
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
                Case adVarChar
                    lngLength = madLongVarCharDefault
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

Public Sub OutputRsData(ByVal rsObj As ADODB.Recordset)
    Dim intCol As Integer, intCols As Integer
    Dim strValues As String
    With rsObj
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strValues = ""
            intCols = .Fields.Count - 1
            For intCol = 0 To intCols
                strValues = strValues & "," & .Fields(intCol).Name & ":" & .Fields(intCol).Value
            Next
            Debug.Print strValues
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub
'----------------------------------------------------------------------------------------------------------------------------

Public Function MecFlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'֧����Ҹ�ҳ������ֵĹ�����35���ϰ汾����
    On Error GoTo errH
    If Not gobjMec Is Nothing Then
        Call gobjMec.PlugWndProc(wMsg, wParam, lParam, 0)
    End If
    MecFlexScroll = CallWindowProc(glngPreHWnd, hwnd, wMsg, wParam, lParam)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFlexCol(ByVal objFlex As Object, ByVal strCaption As String) As Long
'���ܣ�������ͷ��ʾ���ֻ�ȡvsFlexGrid���к�
'���أ��޶�Ӧ����ʱ������-1
    Dim i As Long
    
    GetFlexCol = -1
    
    For i = 1 To objFlex.Cols - 1
        If UCase(objFlex.TextMatrix(0, i)) = UCase(strCaption) Then
            GetFlexCol = i: Exit Function
        End If
    Next
End Function
