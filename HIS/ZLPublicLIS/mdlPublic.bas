Attribute VB_Name = "mdlpublic"
Option Explicit

Public Function GetLisSample() As ADODB.Recordset
    '��ȡ����걾
    Dim strSql  As String
    If gblnNewLis Then
        strSql = "select A.����, B.˳�� from ����걾���� A, �걾˳�� B where A.���� = B.����(+) order by B.˳��"
    Else
        strSql = "select A.����, B.˳�� from ���Ƽ���걾 A, �걾˳�� B where A.���� = B.����(+) order by B.˳��"
    End If
    Set GetLisSample = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
End Function

Public Function GetLisType() As ADODB.Recordset
    '��ȡ�������
    Dim strSql  As String
    If gblnNewLis Then
        strSql = "select A.���� As ����, B.˳�� from (select distinct ���� from ���������Ŀ) A, ���˳�� B where A.���� = B.����(+) order by B.˳��"
    Else
        strSql = "select A.����, B.˳�� from ���Ƽ������� A, ���˳�� B where A.���� = B.����(+) order by B.˳��"
    End If
    Set GetLisType = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
End Function

Public Function GetLisName() As ADODB.Recordset
    '��ȡ��������
    Dim strSql  As String
    If gblnNewLis Then
        strSql = "select A.����, B.˳��" & vbNewLine & _
            "  from ���������Ŀ A, ��Ŀ˳�� B" & vbNewLine & _
            " where A.���� = B.����(+)" & vbNewLine & _
            "   And (A.ͣ������ is null or A.ͣ������ > sysdate)" & vbNewLine & _
            " order by B.˳��"
            Set GetLisName = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName)
    Else
        strSql = "select Decode(Instr(A.����, '('),0,A.����,substr(A.����, 1, Instr(A.����, '(') - 1)) As ����, B.˳��" & vbNewLine & _
            "  from ������ĿĿ¼ A, ��Ŀ˳�� B" & vbNewLine & _
            " where A.���� = B.����(+)" & vbNewLine & _
            "   And A.��� = 'C'" & vbNewLine & _
            "   And A.����Ӧ�� = 1" & vbNewLine & _
            "   And (A.����ʱ�� is null or A.����ʱ�� > sysdate)" & vbNewLine & _
            " order by B.˳��"
         Set GetLisName = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
    End If
   
    
End Function

Public Function JustType() As Integer
    Dim strSql  As String
    Dim rsType  As ADODB.Recordset
    Dim rsSamp  As ADODB.Recordset
    Dim rsName  As ADODB.Recordset

    strSql = "select ˳��,count(*) As ���� from ���˳�� group by ˳��  Order By ˳��"
    Set rsType = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
    strSql = "select ˳��,count(*) As ���� from �걾˳�� group by ˳��  Order By ˳�� "
    Set rsSamp = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
    strSql = "select ˳��,count(*) As ���� from ��Ŀ˳�� group by ˳��  Order By ˳��"
    Set rsName = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
    If rsType.RecordCount > 1 Then
        JustType = 0
    ElseIf rsSamp.RecordCount > 1 Then
        JustType = 1
    ElseIf rsType.RecordCount > 1 And rsSamp.RecordCount > 1 Then
        JustType = 2
    ElseIf rsName.RecordCount > 1 Then
        JustType = 3
    Else
        JustType = 0
    End If

'    If JustType < 0 Then JustType = 0
End Function

Public Function CopyRecordStruct(ByVal rsFrom As ADODB.Recordset, Optional ByVal blnRowID As Boolean = False, Optional ByVal blnNotOpen As Boolean = False) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Dim lngLoop As Long
    Dim rs As ADODB.Recordset

    On Error GoTo errHand

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.CursorType = adOpenStatic

    For lngLoop = 0 To rsFrom.Fields.count - 1

        Select Case rsFrom.Fields(lngLoop).type
        Case 135            'Oracle��Date��
            rs.Fields.Append rsFrom.Fields(lngLoop).Name, adVarChar, 100, rsFrom.Fields(lngLoop).Attributes
        Case Else
            rs.Fields.Append rsFrom.Fields(lngLoop).Name, adVarChar, rsFrom.Fields(lngLoop).DefinedSize + 100
        End Select

    Next
    If blnRowID Then
        rs.Fields.Append "�к�", adVarChar, 30
    End If

    If blnNotOpen = False Then rs.Open

    Set CopyRecordStruct = rs

    Exit Function
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CopyRecordData(ByVal rsFrom As ADODB.Recordset, ByRef rsTo As ADODB.Recordset, Optional blnAll As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strTmp As String
    Dim lngLoop As Long

    On Error GoTo errHand

    If blnAll Then
        If rsFrom.RecordCount > 0 Then rsFrom.MoveFirst
    End If

    Do While Not rsFrom.EOF
        rsTo.AddNew
        For lngLoop = 0 To rsFrom.Fields.count - 1

            On Error Resume Next
            strTmp = ""
            strTmp = rsTo.Fields(rsFrom.Fields(lngLoop).Name).Name
            On Error GoTo errHand

            If UCase(strTmp) = UCase(rsFrom.Fields(lngLoop).Name) Then
                rsTo.Fields(strTmp).Value = Trim(Nvl(rsFrom.Fields(lngLoop).Value))
            End If

        Next
        If blnAll = False Then Exit Do
        rsFrom.MoveNext
    Loop

    CopyRecordData = True

    Exit Function

errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function



