Attribute VB_Name = "mdlBaseCode"
Option Explicit 'Ҫ���������

Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'�л���ָ�������뷨��
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'��ȡĳ�����뷨������
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'�ж�ĳ�����뷨�Ƿ��������뷨
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'��ȡָ�����뷨����Layout,����Ϊ0ʱ��ʾ��ǰ���뷨��
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
'��ȡ��ǰ���뷨����Layout��
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
'�������뷨Layout���������뷨�л������뷨�л�˳�����ǰͷ(������������Ч),flags����=KLF_REORDER
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
Public Const KLF_REORDER = &H8

Public Function SystemImes() As Variant
'���ܣ���ϵͳ�������뷨���Ʒ��ص�һ���ַ���������
'���أ�����������������뷨,�򷵻ؿմ�
    Dim arrIme(99) As Long, arrName() As String
    Dim lngLen As Long, strName As String * 255
    Dim lngCount As Long, i As Integer, j As Integer

    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then 'Ϊ1��ʾ�������뷨
            ReDim Preserve arrName(j)
            lngLen = ImmGetDescription(arrIme(i), strName, Len(strName))
            arrName(j) = Mid(strName, 1, InStr(1, strName, Chr(0)) - 1)
            j = j + 1
        End If
    Next
    SystemImes = IIf(j > 0, arrName, vbNullString)
End Function

Public Function OpenIme(Optional strIme As String) As Boolean
'����:�����ƴ��������뷨,��ָ������ʱ�ر��������뷨��֧�ֲ������ơ�
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    
    If strIme = "���Զ�����" Then OpenIme = True: Exit Function
       
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            ImmGetDescription arrIme(lngCount), strName, Len(strName)
            If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                Exit Function
            End If
        ElseIf strIme = "" Then
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function

Public Function ChooseIME(cmbIME As Object) As Boolean
    Dim varIME As Variant
    Dim i As Integer
    
    varIME = SystemImes
    If Not IsArray(varIME) Then
        MsgBox "�㻹û��װ�κκ������뷨������ʹ�ñ����ܡ�" & vbCrLf & _
               "���뷨�İ�װ���ڿ����������ɡ�", vbInformation, gstrSysName
        Exit Function
    End If
    cmbIME.Clear
    cmbIME.AddItem "���Զ�����"
    For i = LBound(varIME) To UBound(varIME)
        cmbIME.AddItem varIME(i)
        If gstrIme = varIME(i) Then cmbIME.Text = gstrIme
    Next
    If cmbIME.ListIndex < 0 Then cmbIME.ListIndex = 0
    ChooseIME = True
End Function

Public Function GetMax(ByVal strTable As String, ByVal strField As String) As String
 '���ܣ���ȡָ����ı�����������ֵ
'������strTable  ����;
'      strField  �ֶ���;
'      intLength �ֶγ���
'���أ��ɹ����� �¼�������; ���߷��� 0
    Dim rsTemp As New ADODB.Recordset
    Dim varTemp As Variant, strSQL As String
    
    Err = 0
    On Error GoTo errHand
    With rsTemp
        strSQL = "SELECT MAX(" & strField & ") FROM " & strTable
        Call zldatabase.OpenRecordset(rsTemp, strSQL, "mdlBaseCode")
        varTemp = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
        .Close
    End With
    If IsNumeric(varTemp) Then
        GetMax = CStr(Val(varTemp) + 1)
    Else
        GetMax = Mid(varTemp, 1, Len(varTemp) - 1) & Chr(Asc(Right(varTemp, 1)) + 1)
    End If
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0
End Function

Public Function GetDownCodeLength(ByVal strID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '������������ȡָ����ı����������󳤶�
    '�������������ID������
    '����������ɹ����� �¼�������; ���߷��� 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If strID = "" Then
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " start with �ϼ�ID is null " & strWhere & " connect by prior id=�ϼ�id"
    Else
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " start with ID=" & strID & strWhere & " connect by prior id=�ϼ�id"
    End If
    
    Call zldatabase.OpenRecordset(rsTemp, strSQL, "mdlBaseCode")
    If rsTemp.RecordCount = 0 Then
        GetDownCodeLength = 0
    Else
        GetDownCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetDownCodeLength = 0
End Function

Public Function GetLocalCodeLength(ByVal str�ϼ�ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '������������ȡָ����ı����������󳤶�
    '����������ϼ�ID������
    '����������ɹ����� ������; ���߷��� 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " where �ϼ�ID is null" & strWhere
    Else
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " where �ϼ�ID=" & str�ϼ�ID & strWhere
    End If

    Call zldatabase.OpenRecordset(rsTemp, strSQL, "mdlBaseCode")
    If rsTemp.RecordCount = 0 Then
        GetLocalCodeLength = 0
    Else
        GetLocalCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetLocalCodeLength = 0
End Function

Public Function GetParentCode(ByVal str�ϼ�ID As String, ByVal strTableName As String) As String
    '������������ȡ�ϼ�����
    '����������ϼ�ID,����
    '����������ɹ����� �ϼ�����; ���߷��� ��
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        GetParentCode = ""
        Exit Function
    Else
        strSQL = "select ���� from " & strTableName & " where ID=[1]"
    End If

    'by lesfeng 2010-03-08 �����Ż�
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "mdlBaseCode", Val(str�ϼ�ID))
    If rsTemp.RecordCount = 0 Then
        GetParentCode = ""
    Else
        GetParentCode = rsTemp.Fields("����").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetParentCode = ""
End Function

Public Function GetMaxLocalCode(ByVal str�ϼ�ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As String
    '��������������ָ������ϼ�ID ��ȡ������������
    '����������ϼ�ID,����
    '����������ɹ����� ������; ���߷��� ��
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim intCode As Integer, strCode As String, strAllCode As String
    Dim intLength   As Integer
    Err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        strSQL = "select max(to_number(����))+1 as MaxCode from " & strTableName & " where �ϼ�ID is null" & strWhere
    Else
        strSQL = "select nvl(max(to_number(����)),0)+1 as MaxCode from " & strTableName & " where �ϼ�ID=" & str�ϼ�ID & strWhere
    End If
    intCode = GetLocalCodeLength(str�ϼ�ID, strTableName)

    Call zldatabase.OpenRecordset(rsTemp, strSQL, "mdlBasecode")
    If rsTemp.EOF Then
        GetMaxLocalCode = ""
        Exit Function
    End If
    intLength = intCode - Len(IIf(IsNull(rsTemp.Fields("MaxCode").Value), 0, rsTemp.Fields("MaxCode").Value))
    strAllCode = String(IIf(intLength < 0, 0, intLength), "0") & rsTemp.Fields("MaxCode").Value
    'strCode = Mid(strAllCode, Len(GetParentCode(str�ϼ�ID, strTableName)) + 1)
    'GetMaxLocalCode = String(intCode - Len(strAllCode), "0") & strCode
    GetMaxLocalCode = Mid(strAllCode, Len(GetParentCode(str�ϼ�ID, strTableName)) + 1)
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetMaxLocalCode = ""
End Function

Public Sub �ı����(nodParent As Node, int��ȥ���� As Integer, str�������� As String)
'����:�ı������б���ڵ�ı����б����ֵ
'����:nodParent         Ҫ�ı�������ʼ�ڵ�
'     int��ȥ����       ��������ȥ����
'     str��������       ��������������
    Dim nod As Node
    '�����¼�ҲҪ�ı����
    If nodParent.Children > 0 Then
        Set nod = nodParent.Child
        Do While Not (nod Is Nothing)
            nod.Text = "��" & str�������� & Mid(nod.Text, int��ȥ���� + 2)
            �ı���� nod, int��ȥ����, str��������
            Set nod = nod.Next
        Loop
    End If
End Sub
