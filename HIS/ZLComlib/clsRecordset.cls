VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'##############################################################################
'#���ܣ���¼��������صĴ������жϵȹ�������
'�������򣺺�����������Ϊͨ���׶���Ӣ�ļ��ɣ����ü�Rec�ؼ���
'�ϼ��ࣺclsDatabase
'##############################################################################
Public Function InsertCurRow(ByVal rsSource As ADODB.Recordset, ByRef rsNewData As ADODB.Recordset, Optional strPKFields As String = "ID", Optional blnPkIsNumder As Boolean = True) As Boolean
'����:����ǰ��¼������,���뵽Ŀ���¼����
'���:rsSouce-ԭ��¼��
'       strPKFields-�������ĸ�(�������������ֵͬ�����ܲ���.),ֻ����һ���ֶ�
'       blnPkIsNumder-��������
'����:rsNewData-�������ݺ���¼�¼��
 '����:����ɹ�,����true,���򷵻�False
    InsertCurRow = gobjComLib.zlDatabase.zlInsertCurrRowData(rsSource, rsNewData, strPKFields, blnPkIsNumder)
End Function

Public Function CopyStructure(ByVal rsSource As ADODB.Recordset) As ADODB.Recordset
'����:    ���Ƽ�¼�������ݽṹ
'������ rsSource-ԭ��¼��
'���أ�������ԭ��¼����ͬ�ṹ�ı��ؼ�¼��
    Set CopyStructure = gobjComLib.zlDatabase.zlCopyDataStructure(rsSource)
End Function

Public Function Update(ByRef rsInput As Recordset, ByVal strFilter As String, ParamArray arrInput() As Variant) As Boolean
'���ܣ�����ָ�������ļ�¼���ļ�¼
'������rsInput=��¼��
'      strFilter=����
'      arrInput=������ֶ����Լ�ֵ����ʽ���ֶ���1,ֵ1, �ֶ���2,ֵ2,....
'���أ��Ƿ�ɹ�
'      rsInput=�������º�ļ�¼��
'˵����arrInput���ֶ�ֵ�����ü�¼���е������ֶ������¸��ֶΣ���ʱ��ʽΪ��!�ֶ���
    Dim strFiledName As String, strFileValue As String
    Dim blnFiled As Boolean, i As Long

    On Error GoTo errH
    With rsInput
        .Filter = strFilter
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            For i = LBound(arrInput) To UBound(arrInput) Step 2
                strFiledName = arrInput(i)
                If IsNull(arrInput(i + 1)) Then
                    rsInput(strFiledName).value = Null
                Else
                    If arrInput(i + 1) Like "!?*" Then
                        blnFiled = True
                        On Error Resume Next
                        strFileValue = rsInput(Mid(arrInput(i + 1), 2)).value & ""
                        If Err.Number <> 0 Then Err.Clear: blnFiled = False
                        On Error GoTo errH
                    End If
                    If Not blnFiled Then
                        rsInput(strFiledName).value = arrInput(i + 1)
                    Else
                        rsInput(strFiledName).value = rsInput(Mid(arrInput(i + 1), 2)).value
                    End If
                End If
                blnFiled = False
                Call rsInput.Update
            Next
            .MoveNext
        Loop
    End With
    Update = True
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function ToString(ByVal rsObj As ADODB.Recordset) As String
'���ܣ�����¼��ת�����ַ������Ǽ�¼��GetString�����ļ�ʹ�ã��������Լ�¼��
    ToString = rsObj.GetString(, , "[Col]", vbNewLine, "NULL")
    If rsObj.RecordCount <> 0 Then rsObj.MoveFirst
End Function

Public Function ToSQL(ByVal rsTemp As ADODB.Recordset, Optional ByVal strHiddenCol As String) As String
'���ܣ�����¼��ת��SQL��
'����:
'    strHiddenCol -����ʾ���� ����֮���÷ֺŷָ�
    Dim strSQL As String
    Dim i As Long
    
    With rsTemp
        strSQL = ""
        Do While Not .EOF
            strSQL = strSQL & " Union ALL Select "
            For i = 0 To .Fields.count - 1
                If InStr(";" & strHiddenCol & ";", ";" & .Fields(i).Name & ";") = 0 Then '���ε�������
                    If Not IsNull(.Fields(i).value) Then
                        If IsType(.Fields(i).type, adVarChar) Then
                            strSQL = strSQL & "'" & Replace(Replace(.Fields(i).value, "[", "("), "]", ")") & "'"
                        ElseIf IsType(.Fields(i).type, adDBTimeStamp) Then    '������
                            strSQL = strSQL & "to_date('" & .Fields(i).value & "','" & "YYYY-MM-DD')"
                        Else
                            strSQL = strSQL & .Fields(i).value 'û��������
                        End If
                    Else
                        If IsType(.Fields(i).type, adBigInt) Or IsType(.Fields(i).type, adSmallInt) Or IsType(.Fields(i).type, adSingle) Then
                            strSQL = strSQL & "-Null"
                        Else
                            strSQL = strSQL & "Null"
                        End If
                    End If
                    strSQL = strSQL & " As " & .Fields(i).Name & ","
                End If
            Next
            strSQL = Left(strSQL, Len(strSQL) - 1) & " From Dual"
            .MoveNext
        Loop
        .Filter = ""
        strSQL = Mid(strSQL, 12)
    End With
    
    ToSQL = strSQL
End Function

Public Function ToComboList(ByVal rsInput As ADODB.Recordset, ByVal strFormat As String, ParamArray arrInput() As Variant) As String
'���ܣ�����¼����ָ���ֶΣ���װ���ض���ʽ�ַ��������ַ�����������VsFlexGrid�������б�
'������rsInput=����Դ
'           strFormat=��ʽ�ַ�����[0]��[1]...�Ƚ���ռλ
'           arrInput=ռλ����Ӧ���ֶ���
'���أ����������б�ֵ
    Dim i As Long
    Dim strTmp As String, strReturn As String
    
    Do While Not rsInput.EOF
        strTmp = strFormat
        For i = LBound(arrInput) To UBound(arrInput)
              strTmp = Replace(strTmp, "[" & i & "]", rsInput.Fields(arrInput(i)).value & "")
        Next
        strReturn = strReturn & strTmp
        rsInput.MoveNext
    Loop
    ToComboList = strReturn
End Function

Public Function FilterNew(ByVal rsInput As ADODB.Recordset, Optional ByVal strFilter As String, Optional ByVal strFiledsName As String) As ADODB.Recordset
'���ܣ�����ָ�����������ļ�¼��������ü�¼�����µļ�¼������ԭ���ļ�¼��û�й�ϵ
'������rsInput=Ŀ���¼��
'      strFilter=�������������ι�������������ļ�¼���Ѿ����ˣ��ڹ��˺�Ĺ�������
'      strFiledsName=��Ҫ���Ƶ��ֶ���
'���أ����˺�ļ�¼���ĸ���
'˵�������ܹ�������̫���Ҹ��ӣ������Ҫ���ι��ˣ��ú�����Ҫ����������
    Dim rsClone As ADODB.Recordset
    
    On Error GoTo errH
    'û�й�������ֱ�Ӹ���
    If strFilter = "" Then
        Set FilterNew = CopyNew(rsInput, , strFiledsName)
        Exit Function
    End If
    Set rsClone = CopyNew(rsInput)
    rsClone.Filter = strFilter
    Set FilterNew = CopyNew(rsClone, , strFiledsName)
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function ExitsField(ByVal rsData As ADODB.Recordset, ByVal strFieldName As String) As Boolean
    '******************************************************************************************************************
    '���ܣ��жϼ�¼���Ƿ����ָ�����Ƶ��ֶ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strTmp As String
    
    On Error Resume Next
    strTmp = ""
    strTmp = rsData.Fields(strFieldName).Name
    ExitsField = (strTmp = strFieldName)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Public Function Distinct(ByVal rsSource As ADODB.Recordset, Optional ByVal strDisFieldsName As String, Optional ByVal strFieldsName As String) As ADODB.Recordset
'���ܣ���¼��ȥ�ظ�
'������rsSource=Ҫȥ�ظ��ļ�¼��
'strDisFieldsName=ȥ�ظ����ֶ�,Ϊ�գ���������ֶ�ȥ��
'strFieldsName=���ؽ�����ֶΣ�Ϊ�գ��򷵻�ȥ�ظ����ֶ�
'���أ�������ļ�¼��
    Dim rsReturn As ADODB.Recordset
    Dim arrFilds As Variant, arrValues As Variant
    Dim i As Long, j As Long
    Dim strTmp As String, strOldRow As String
    
    On Error GoTo errH
    '��ȡĬ���ֶ���
    If strDisFieldsName = "" Then
        For i = 0 To rsSource.Fields.count - 1
            strTmp = strTmp & "," & rsSource.Fields(i).Name
        Next
        strTmp = Mid(strTmp, 2)
        If strDisFieldsName = "" Then strDisFieldsName = strTmp
    End If
    If strFieldsName = "" Then strFieldsName = strDisFieldsName
    
    Set rsReturn = CopyNew(rsSource, , strFieldsName)
    If rsSource.RecordCount = 0 Then Set Distinct = rsReturn: Exit Function
    
    rsReturn.Sort = strDisFieldsName '�����Զ�������ƶ�����ͷ
    Do While Not rsReturn.EOF
        strTmp = rsReturn.GetString(, 1, "[ColumnSpliter]", , "[NULLEXP]") '�Զ��ƶ����
        rsReturn.MovePrevious
        If strTmp = strOldRow Then  'ɾ���ظ���
            Call rsReturn.Delete: Call rsReturn.Update
        Else
            strOldRow = strTmp
        End If
        rsReturn.MoveNext
    Loop
    rsReturn.Sort = strDisFieldsName
    Set Distinct = rsReturn
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function Delete(ByRef rsInput As ADODB.Recordset, Optional ByVal strFilter As String) As Boolean
'���ܣ�ɾ��ָ�������ļ�¼���ļ�¼
'������rsInput=��¼��
'      strFilter=����
'���أ��Ƿ�ɹ�
'      rsInput=����ɾ����ļ�¼��
    On Error GoTo errH
    rsInput.Filter = strFilter
    If rsInput.RecordCount > 0 Then
        rsInput.MoveFirst
        Do While Not rsInput.EOF
            Call rsInput.Delete: Call rsInput.Update
            rsInput.MoveNext
        Loop
    End If
    Delete = True
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function Compare(ByVal rsSouce As ADODB.Recordset, ByVal rsCompare As ADODB.Recordset) As Boolean
'�Ƚ�������¼���Ƿ���ͬ
'˵��������¼���ṹ��ͬ����False,�ṹ��ͬ���ݲ�ͬ����False,�����Լ��ṹ��ȫ��ͬ����True
'      �ṹ���ǱȽ��ֶ������Լ�˳�򣬾������Ͳ��Ƚϣ���ʱ�ò�������Ҫʱ������չ��
    Dim strTmp1 As String, strTmp2 As String
    Dim i As Long
    
    On Error GoTo errH
    If rsSouce.RecordCount <> rsCompare.RecordCount Then Compare = False: Exit Function
    If rsSouce.Fields.count <> rsCompare.Fields.count Then Compare = False: Exit Function
    
    For i = 0 To rsSouce.Fields.count - 1
        strTmp1 = strTmp1 & "," & rsSouce.Fields(i).Name
    Next
    
    For i = 0 To rsCompare.Fields.count - 1
        strTmp2 = strTmp2 & "," & rsCompare.Fields(i).Name
    Next
    
    If strTmp1 <> strTmp2 Then Compare = False: Exit Function
    strTmp1 = Mid(strTmp1, 2)
    rsCompare.Sort = strTmp1: rsSouce.Sort = strTmp1
    strTmp1 = rsCompare.GetString(, , "[ColumnSplit]", "[RowSplit]", "NULL")
    strTmp2 = rsSouce.GetString(, , "[ColumnSplit]", "[RowSplit]", "NULL")
    Compare = (strTmp1 = strTmp2)
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function Append(ByRef rsSource As ADODB.Recordset, ByVal rsAppend As ADODB.Recordset, ParamArray arrInput() As Variant) As Boolean
'���ܣ���ָ����¼�����������ӵ���һ����¼����
'������rsSource=Ŀ���¼��
'      rsAppend=���ݼ�¼��
'      arrInput=�ֶζ�Ӧ���򣬸ò�������ʱ��Ĭ������¼���ṹ��ͬ����ʽ��arrInput(0):[��¼��1].�ֶ�1,�ֶ�2...��arrInput(1)��[��¼��2].�ֶ�1,�ֶ�2...
'���أ��Ƿ�ɹ�
'      rsSource=�������ݺ�ļ�¼��
    Dim arrSource As Variant, arrAppend As Variant
    Dim i As Long, arrValues() As Variant
    Dim strTmp As String
    
    If rsAppend Is Nothing Then Append = True: Exit Function
    If rsAppend.RecordCount = 0 Then Append = True: Exit Function
    If rsSource Is Nothing Then Exit Function
    On Error GoTo errH
    If LBound(arrInput) = 2 Then
        arrSource = Split(arrInput(LBound(arrInput)), ",")
        arrAppend = Split(arrInput(UBound(arrInput)), ",")
        If UBound(arrSource) <> UBound(arrAppend) Then Exit Function
        ReDim arrValues(UBound(arrAppend)): rsAppend.MoveFirst
        Do While Not rsAppend.EOF
            For i = LBound(arrAppend) To UBound(arrAppend)
                arrValues(i) = rsAppend(arrAppend(i)).value
            Next
            rsSource.AddNew arrSource, arrValues
            Erase arrValues
            rsAppend.MoveNext
        Loop
    ElseIf LBound(arrInput) = 0 Then
        Do While Not rsAppend.EOF
            rsSource.AddNew
            For i = 0 To rsSource.Fields.count - 1
                rsSource.Fields(i).value = rsAppend.Fields(i).value
            Next
            rsSource.Update
            rsAppend.MoveNext
        Loop
    End If
    
    Append = True
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Sub AddNew(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    '���Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    Dim arrFields, arrValues, intField As Integer
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
            .Fields(arrFields(intField)).value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function CopyNew(ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean, Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
'���Ƽ�¼��
'������strFields=��Ҫ���Ƶļ�¼�����ֶε���˳����ֶ�����ɵ��ַ���
'          �磺1 ����1,3 ����2,7 ����3...��ʾ���Ƽ�¼���ĵ�1,3,7..�ֶ���ɼ�¼��������
'              ID ����1,���� ����2,....��ʾ���Ƽ�¼����ID,����...�ֶ���ɼ�¼������
'              ����*Ϊ�µļ�¼��������
'              �������ͻ�����׳���������ͬ�����⣬��ע��
'           arrAppFields=׷�ӵ��ֶ���Ϣ������,����,����,Ĭ��ֵ,û��Ĭ��ֵ��Empty,û��ָ�����ȴ�Empty
'      blnOnlyStructure=�Ƿ�ֻ���ƽṹ
'�ڳ����У��������漰���໥���ݼ�¼������ʹ��ADO��Clone���Ʋ����ļ�¼����������һ����¼�������ݷ����仯��ʱ�����и�������������ͬ�ı仯��ͨ��ָ�޸Ļ�ɾ����������������ϣ����Щ��¼���໥�䱣�ֶ���
    Set CopyNew = gobjComLib.zlDatabase.CopyNewRec(rsSource, blnOnlyStructure, strFields, arrAppFields)
End Function

Public Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'���ܣ��ж�ĳ��ADO�ֶ����������Ƿ���ָ���ֶ�������ͬһ��(������,����,�ַ�,������)
    Dim intA As Integer, intB As Integer
    
    Select Case varBase
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intA = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intA = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intA = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intA = -4
        Case Else
            intA = varBase
    End Select
    Select Case varType
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intB = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intB = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intB = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intB = -4
        Case Else
            intB = varType
    End Select
    IsType = intA = intB
End Function

Public Function DownloadPicture(rsTable As ADODB.Recordset, strField As String, Optional strFile As String) As String
'���ܣ���ָ���ļ�¼��ͼ���ֶ����ص���������Ϊͼ����ʱ�ļ�
'������
'       rsTable   ͼ�δ洢��¼��
'       strField  ͼ���ֶ�
'       strFile   �û�������ļ�������ѡ�
'���أ�
    DownloadPicture = gobjComLib.zlDatabase.ReadPicture(rsTable, strField, strFile)
End Function

Public Function UploadPicture(strFile As String, rsTable As ADODB.Recordset, strField As String) As Boolean
'���ܣ���ָ����ͼ���ļ��ϴ���ָ���ļ�¼���ֶ���
'������
'       strFile��ͼ���ļ�����
'       rsTable��ͼ�δ洢��¼��
'       strField��ͼ���ֶ�
'���أ�
    UploadPicture = gobjComLib.zlDatabase.SavePicture(strFile, rsTable, strField)
End Function