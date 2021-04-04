Attribute VB_Name = "mdlPublic"
Option Explicit
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Public Const tomAutoColor As Long = -9999997
Public gstrFields As String
Public gstrValues As String


Public gstrProductName As String            '��Ʒ��ƣ����磺����
Public gstrSysName As String                'ϵͳ���ƣ����磺�������
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public gstrMatch As String                  '���ݱ��ز�����ƥ��ģʽ��ȷ������ƥ�����

Public gstrDbOwner As String                '��ǰ���ݿ������ߣ���ͬģ����ܲ�һ����
Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public glngSys As Long
Public gstrSQL As String
Public gcnOracle As New ADODB.Connection
Public gobjTendEditor As Object

Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'################################################################################################################
'##  �õ��û�����Ϣ
'################################################################################################################
Public Sub GetUserInfo()
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    On Error GoTo Errhand
    strSQL = "select u.�û���, P.*,D.���� as ���ű���,D.���� as ��������,M.����ID" & _
                " from �ϻ���Ա�� U,��Ա�� P,���ű� D,������Ա M " & _
                " Where U.��Աid = P.id And P.ID=M.��ԱID and  M.ȱʡ=1 and M.����id = D.id and U.�û���=user And (p.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or p.����ʱ�� Is Null) "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetUserInfo")
    With rsTemp
        If .RecordCount <> 0 Then
            gstrDBUser = .Fields("�û���").Value
            glngUserId = .Fields("ID").Value                '��ǰ�û�id
            gstrUserCode = .Fields("���").Value            '��ǰ�û�����
            gstrUserName = .Fields("����").Value            '��ǰ�û�����
            gstrUserAbbr = IIf(IsNull(.Fields("����").Value), "", .Fields("����").Value)          '��ǰ�û�����
            glngDeptId = .Fields("����id").Value            '��ǰ�û�����id
            gstrDeptCode = .Fields("���ű���").Value        '��ǰ�û�
            gstrDeptName = .Fields("��������").Value        '��ǰ�û�
        Else
            gstrDBUser = ""
            glngUserId = 0
            gstrUserCode = ""
            gstrUserName = ""
            gstrUserAbbr = ""
            glngDeptId = 0
            gstrDeptCode = ""
            gstrDeptName = ""
        End If
        .Close
    End With
   
   
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0
End Sub

'################################################################################################################
'## ���ܣ�  ��ָ����LOB�ֶθ���Ϊ��ʱ�ļ�
'##
'## ������  Action      :�������ͣ����������ǲ����ĸ���
'##         KeyWord     :ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'##         strFile     :�û�ָ����ŵ��ļ�������ָ��ʱ��ȡ��ǰ·�������ļ���
'##
'## ���أ�  ������ݵ��ļ�����ʧ���򷵻��㳤��""
'##
'## ˵����  Actionȡֵ˵����
'##         0-�������ͼ�Σ�1-�����ļ���ʽ��2-�����ļ�ͼ�Σ�3-�������ĸ�ʽ��4-��������ͼ�Σ�5-���Ӳ�����ʽ��6-���Ӳ���ͼ�Σ�
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, _
                           ByVal KeyWord As String, _
                           Optional ByRef strFile As String, _
                           Optional ByVal blnMoved As Boolean) As String
    
    Const conChunkSize As Integer = 10240

    Dim lngFileNum     As Long, lngCount As Long, lngBound As Long

    Dim aryChunk()     As Byte, strText As String

    Dim rsLob          As New ADODB.Recordset
    
    Err = 0: On Error GoTo Errhand
    
    lngFileNum = FreeFile

    If strFile = "" Then
        lngCount = 0

        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"

            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop

    End If

    Open strFile For Binary As lngFileNum
    
    gstrSQL = "Select Zl_Lob_Read(" & Action & ",'" & KeyWord & "'," & "[1]) as Ƭ�� From Dual"
    lngCount = 0

    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlBlobRead", lngCount)

        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte

        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop

    Close lngFileNum

    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile

    Exit Function

Errhand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub UniteCellCol(ByVal objCell As Object, _
                        ByVal intCOl As Integer, _
                        ByVal intRow As Integer, _
                        Optional startCol As Integer = 1)

    '���ܣ��ϲ���Ԫ�����.
    '����: intcol Ҫ�ϲ�������  introw�ڼ���  startCol ��ʼ��
    On Error GoTo Errhand

    Dim strText As String
    Dim i As Integer, j As Integer

    With objCell
        .MergeRow(intRow) = True
        strText = " " & String(intRow, " ")
        
        For i = startCol To .Cols - 2
        
            j = i - objCell.FixedCols + 1
            If j < 0 Then j = 1
            
            If j Mod intCOl <> 0 Then
                .MergeCol(i) = True
                .Row = intRow
                .Col = i
                .Text = strText
                .CellAlignment = 4
                .Row = intRow
                .Col = i + 1
                .Text = strText
                .CellAlignment = 4
            Else
                strText = String(j / intCOl + 1, " ") & String(intRow, " ")
            End If
        Next i
    End With

    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------------------
'�����ǻ������������
'---------------------------------------------------------------------------------
Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, _
                      ByVal strFields As String, _
                      ByVal strValues As String)

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

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, _
                         ByVal strFields As String, _
                         ByVal strValues As String, _
                         ByVal strPrimary As String, _
                         Optional ByVal blnDelete As Boolean = False)

    Dim arrFields, arrValues, intField As Integer

    '���¼�¼,���������,������
    'strPrimary:�ֶ���,ֵ
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'strPrimary = "RecordID,5188"
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

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, _
                              ByVal strPrimary As String, _
                              Optional ByVal blnDelete As Boolean = False) As Boolean

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

Public Sub OutputRsData(ByVal rsObj As ADODB.Recordset, _
                         Optional ByVal blnMod_Add As Boolean = False)

    Dim strOutput As String

    Dim intCOl    As Integer, intCols As Integer

    With rsObj

        If .RecordCount <> 0 Then .MoveFirst

        Do While Not .EOF
            strOutput = ""
            intCols = .Fields.Count

            For intCOl = 1 To intCols

                If Not blnMod_Add Then
                    strOutput = strOutput & "," & .Fields(intCOl - 1).Name & ":" & .Fields(intCOl - 1).Value
                Else
                    strOutput = strOutput & "|" & .Fields(intCOl - 1).Value
                End If

            Next

            Debug.Print Mid(strOutput, 2)
            
            .MoveNext
        Loop

        If .RecordCount <> 0 Then .MoveFirst
    End With

End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function LPAD(ByVal strText As String, ByVal intCount As Integer, ByVal strPAD As String) As String
'���ܣ���ͬOracle��LPAD����
    If LenB(StrConv(strText, vbFromUnicode)) < intCount Then
        LPAD = String(intCount - LenB(StrConv(strText, vbFromUnicode)), strPAD) & strText
    Else
        LPAD = strText
    End If
End Function



