Attribute VB_Name = "mdlCommEdit"
Option Explicit
Public gcnOracle As New ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public glngModul As Long

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼

Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public gstr��λ���� As String
Public gstrSQL As String
Public glngSys As Long

Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Public Sub GetUserInfo()
    '����:�õ��û�����Ϣ
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo ErrHand
    
    Set rsTemp = zlDatabase.GetUserInfo
    
    With rsTemp
        If .RecordCount <> 0 Then
            glngUserId = .Fields("ID").Value                '��ǰ�û�id
            gstrUserCode = .Fields("���").Value            '��ǰ�û�����
            gstrUserName = .Fields("����").Value            '��ǰ�û�����
            gstrUserAbbr = IIF(IsNull(.Fields("����").Value), "", .Fields("����").Value)          '��ǰ�û�����
            glngDeptId = .Fields("����id").Value            '��ǰ�û�����id
            gstrDeptCode = .Fields("������").Value        '��ǰ�û�
            gstrDeptName = .Fields("������").Value        '��ǰ�û�
        Else
            glngUserId = 0
            gstrUserCode = ""
            gstrUserName = ""
            gstrUserAbbr = ""
            glngDeptId = 0
            gstrDeptCode = ""
            gstrDeptName = ""
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0
End Sub

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
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " start with �ϼ�ID=" & strID & strWhere & " connect by prior id=�ϼ�id"
    End If
'    Call SQLTest(App.ProductName, "������󳤶�", strSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetDownCodeLength")
'    Call SQLTest
    
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
'    Call SQLTest(App.ProductName, "������󳤶�", strSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetLocalCodeLength")
'    Call SQLTest
    
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
        strSQL = "select ���� from " & strTableName & " where ID=" & str�ϼ�ID
    End If
'    Call SQLTest(App.ProductName, "�ϼ�����", strSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetParentCode")
'    Call SQLTest
    
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
        strSQL = "select nvl(max(to_number(����)),0)+1 as MaxCode from " & strTableName & " where �ϼ�ID is null" & strWhere
        
        '����ǲ��ű���Ҫ�ų�"��ɾ������"�����ID
        If strTableName = "���ű�" Then
            strSQL = strSQL & " And ���� <> '-'"
        End If
    Else
        strSQL = "select nvl(max(to_number(����)),0)+1 as MaxCode from " & strTableName & " where �ϼ�ID=" & str�ϼ�ID & strWhere
    End If
    intCode = GetLocalCodeLength(str�ϼ�ID, strTableName, strWhere)
'    Call SQLTest(App.ProductName, "����������", strSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetMaxLocalCode")
'    Call SQLTest
    
    If rsTemp.EOF Then
        GetMaxLocalCode = ""
        Exit Function
    End If
    intLength = intCode - Len(IIF(IsNull(rsTemp.Fields("MaxCode").Value), 0, rsTemp.Fields("MaxCode").Value))
    strAllCode = String(IIF(intLength < 0, 0, intLength), "0") & rsTemp.Fields("MaxCode").Value
    'strCode = Mid(strAllCode, Len(GetParentCode(str�ϼ�ID, strTableName)) + 1)
    'GetMaxLocalCode = String(intCode - Len(strAllCode), "0") & strCode
    GetMaxLocalCode = Mid(strAllCode, Len(GetParentCode(str�ϼ�ID, strTableName)) + 1)
    If GetMaxLocalCode = "" Then GetMaxLocalCode = "1"
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetMaxLocalCode = ""
End Function

Public Function Where����ʱ��(Optional strAlias As String) As String
    If strAlias = "" Then
        Where����ʱ�� = " (����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or ����ʱ�� is null) "
    Else
        Where����ʱ�� = " (" & strAlias & ".����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or " & strAlias & ".����ʱ�� is null) "
    End If
End Function

Public Function TruncateDate(ByVal datFull As Date) As Date
'ȥ�������е�ʱ���֡���
    TruncateDate = CDate(Format(datFull, "yyyy-MM-dd"))
End Function

Public Function GetTextFromList(lstTemp As ListBox) As String
'������lstTemp  ׼����ȡ���ݵ�ListBox�ؼ�
    Dim lngCount As Long
    Dim lngPos As Long
    Dim strTemp As String
    
    For lngCount = 0 To lstTemp.ListCount - 1
        If lstTemp.Selected(lngCount) = True Then
            lngPos = InStr(lstTemp.List(lngCount), ".")
            If lngPos = 0 Then
                strTemp = strTemp & lstTemp.List(lngCount) & ","
            Else
                strTemp = strTemp & Mid(lstTemp.List(lngCount), 1, lngPos - 1) & ","
            End If
        End If
    Next
    If strTemp <> "" Then
        'ȥ�����һ��,����
        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
    End If
    GetTextFromList = "'" & strTemp & "'"
End Function

Public Sub SetListByText(lstTemp As ListBox, ByVal strText As String)
'������lstTemp  ׼�����õ�ListBox�ؼ�
    Dim lngCount As Long, lngIndex As Long, lngPos As Long
    Dim strTemp As String, varTemp As Variant
    Dim blnMatch As Boolean
    
    varTemp = Split(strText, ",")
    For lngCount = 0 To lstTemp.ListCount - 1
        blnMatch = False
        'ȡ���õ�ֵ
        lngPos = InStr(lstTemp.List(lngCount), ".")
        If lngPos = 0 Then
            strTemp = lstTemp.List(lngCount)
        Else
            strTemp = Mid(lstTemp.List(lngCount), 1, lngPos - 1)
        End If
        For lngIndex = LBound(varTemp) To UBound(varTemp)
            If strTemp = varTemp(lngIndex) Then
                '�Ѿ��ҵ���ͬ��
                blnMatch = True
                Exit For
            End If
        Next
        lstTemp.Selected(lngCount) = blnMatch
    Next
End Sub


Public Sub ResetSelect(lvw As ListView, ByVal strKey As String)
'���ܣ���������ListView��ѡ����
'������strKey   ˢ��ǰ��ѡ����
    Dim lst As ListItem
    
    If lvw.ListItems.Count > 0 Then
        On Error Resume Next
        Set lst = lvw.ListItems(strKey)
        If Err <> 0 Then
            'û��ѡ�У�Ҳ������Ѿ���ɾ��
            Err.Clear
            Set lst = lvw.ListItems(1)
        End If
        
        '����ѡ����
        lst.Selected = True
        lst.EnsureVisible
    End If
End Sub

Public Sub RemoveSelect(lvw As ListView)
'���ܣ�ɾ����ǰѡ����
    Dim lngIndex  As Long
    
    With lvw
        If .SelectedItem Is Nothing Then Exit Sub
        
        lngIndex = .SelectedItem.Index
        .ListItems.Remove lngIndex
        
        If .ListItems.Count > 0 Then
            '��������б��������һ��ѡ��
            lngIndex = IIF(.ListItems.Count > lngIndex, lngIndex, .ListItems.Count)
            .ListItems(lngIndex).Selected = True
            .ListItems(lngIndex).EnsureVisible
        End If
    End With

End Sub

