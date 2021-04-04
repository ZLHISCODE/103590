VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCSVFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'ʹ��˵����
'1����ȡCSV
'    Set objCsv = New clsCSVFile
'    If objCsv.LoadFile("C:\Test.csv") Then
'        For i = 0 To objCsv.Rows - 1
'            For j = 0 To objCsv.Cols - 1
'                Debug.Print objCsv.Text(i, j)
'            Next
'        Next
'    End If
'    objCsv.Text(2, 2)=3 '���ڶ��е������޸�Ϊ3
'    objCsv.ReSaveFile("C:\Test1.csv")'���Ѿ���ȡ�����޸ĺ��CSV�ļ����±���
'2������CSV
'    Set objCsv = New clsCSVFile
'    If Not objCsv.SaveRec("C:\Test.csv", rsTmp) Then
'        Debug.Print "ERROR"
'    End If

'    Set objCsv = New clsCSVFile
'    If Not objCsv.SaveStart("C:\Test.csv") Then
'        Debug.Print objCsv.ErrorDescription
'    End If
'    Call objCsv.SaveRow("1", "2", "3")  '�����һ�м�¼��1,2,3
'    Call objCsv.SaveValues("3", "4")    '����ڶ��м�¼��ǰ���У�3,4
'    Call objCsv.SaveRow                 '��ǵڶ��������ɣ��ڶ��е������ÿ�ֵ��䣬�ڶ�������Ϊ3,4,
'    Call objCsv.SaveRow("5", "6", "7", "8") '��������м�¼����������������һ���������Զ��ضϣ�ֻ�������5,6,7
'    If Not objCsv.SaveEnd() Then
'        Debug.Print objCsv.ErrorDescription
'    End If
'3��ͬһ�����󽻴���������ͼ��ػᵼ�����ݴ����������
'    If Not objCsv.SaveStart("C:\Test.csv") Then
'        Debug.Print objCsv.ErrorDescription
'    End If
'    If objCsv.LoadFile("C:\Test.csv") Then
'        For i = 0 To objCsv.Rows - 1
'            For j = 0 To objCsv.Cols - 1
'                Debug.Print objCsv.Text(i, j)
'            Next
'        Next
'    End If
'    Call objCsv.SaveRow("1", "2", "3")  '�����һ�м�¼��1,2,3
'��Ҫʹ��һ�������������Ͷ�ȡ���뱣֤��ȡ��ɺ󱣴桢�򱣴���ɺ��ٶ�ȡ��������ʹ�����������Ķ���
'����ֵ
Private mstrError       As String       'CSV���ػ��߱���Ĵ�����Ϣ
Private mlngRows        As Long         'CSV���ػ��߱���ʱ������
Private mlngCols        As Long         'CSV���ػ��߱���ʱ������

Private mlngCurCol      As Long         '��ǰ��¼�Ѿ��������
Private mobjFSO         As New FileSystemObject
'SaveStart��SaveValues��SaveRow��SaveEndʹ�õ�ģ�����
Private mblnFirstRow    As Boolean      '�Ƿ��ǵ�һ��
Private mobjSaveFile    As TextStream   'д�ļ����ļ���

Private mclsBuffer      As New clsStringBulider
Private mobjTextLex     As TextStream
Private marrValues()    As String
Private mstrContent     As String
Private Enum StateType  '����״̬
    NewFieldStart = 0   '���ֶο�ʼ
    NonQuotesField      '�������ֶ�
    QuotesField         '�����ֶ�
    FieldSeparator      '�ֶηָ�
    QuoteInQuotesField  '�����ֶ��е�����
    RowSeparator        '�зָ����ַ�1���س�
    LexError            '�﷨����
End Enum
'======================================================================================================================
'����           ErrorDescription       CSV���ػ��߱���Ĵ�����Ϣ
'����           String
'���ʿ���       (R)
'======================================================================================================================
Public Property Get ErrorDescription() As String
    ErrorDescription = mstrError
End Property
'======================================================================================================================
'����           Rows                    CSV���ػ��߱���ʱ������
'����           Long
'���ʿ���       (R)
'======================================================================================================================
Public Property Get Rows() As Long
    Rows = mlngRows
End Property
'======================================================================================================================
'����           Cols                    CSV���ػ��߱���ʱ������
'����           Long
'���ʿ���       (R)
'======================================================================================================================
Public Property Get Cols() As Long
    Cols = mlngCols
End Property
'======================================================================================================================
'����           Text                    ��ȡCSVĳһ��ĳһ�е�ֵ
'����           String
'���ʿ���       (R/W)
'======================================================================================================================
Public Property Get Text(Row As Long, Col As Long) As String
    If Row < mlngRows And Col < mlngCols And mlngCols > 0 Then
        Text = marrValues(Col + Row * mlngCols)
    End If
End Property

Public Property Let Text(Row As Long, Col As Long, ByVal strValues As String)
    If Row < mlngRows And Col < mlngCols And mlngCols > 0 Then
        marrValues(Col + Row * mlngCols) = strValues
    End If
End Property
'======================================================================================================================
'����           LoadFile                ���ض�ȡCSV�ļ�
'����ֵ         Boolean                 �Ƿ���سɹ�
'����б�:
'������         ����                    ˵��
'strFileName    String                  CSV�ļ�·��
'blnFile        Boolean                 ������Ƿ����ļ����������Ǵ��������ַ���
'======================================================================================================================
Public Function LoadFile(ByVal strFileName As String, Optional ByVal blnFile As Boolean = True) As Boolean
    On Error GoTo ErrH
    If Not mobjSaveFile Is Nothing Then
        Call SaveEnd
    End If
    ReDim Preserve marrValues(16)
    mlngRows = 0
    mlngCols = 0
    mblnFirstRow = True
    If blnFile Then
        Set mobjTextLex = mobjFSO.OpenTextFile(strFileName, ForReading)
    Else
        Set mobjTextLex = Nothing
    End If
    mstrContent = strFileName
    LoadFile = LexCSVContens
    If blnFile Then mobjTextLex.Close
    Set mobjTextLex = Nothing
    Exit Function
ErrH:
    mstrError = err.Description
End Function
'======================================================================================================================
'����           ReSaveFile              ���±����CSV�ļ��ж�ȡ���ļ�
'����ֵ         Boolean                 �Ƿ񱣴�ɹ�
'����б�:
'������         ����                    ˵��
'strFileName    String                  CSV�ļ�·��
'======================================================================================================================
Public Function ReSaveFile(ByVal strFileName As String) As Boolean
    Dim i As Long, j As Long
    If Not mobjSaveFile Is Nothing Then
        Call SaveEnd
    End If
    If Not SaveStart(strFileName) Then
        Exit Function
    End If
    For i = 0 To Me.Rows - 1
        For j = 0 To Me.Cols - 1
            If Not SaveOneValue(Me.Text(i, j)) Then
                Exit Function
            End If
        Next
        If Not SaveRow() Then
            Exit Function
        End If
    Next
    If Not SaveEnd() Then
        Exit Function
    End If
    ReSaveFile = True
End Function
'======================================================================================================================
'����           SaveRec                 ����¼������ΪCSV�ļ�
'����ֵ         Boolean                 �Ƿ񱣴�ɹ�
'����б�:
'������         ����                    ˵��
'strFileName    String                  CSV�ļ�·��
'rsSource       ADODB.Recordset         Ҫ����ļ�¼��
'˵������¼���ӵ�ǰλ�ÿ�ʼ���棬��Ҫ����ȫ����¼�������Filter����MoveFirst
'======================================================================================================================
Public Function SaveRec(ByVal strFileName As String, ByVal rsSource As ADODB.Recordset) As Boolean
    Dim i       As Long
    
    If Not SaveStart(strFileName) Then
        Exit Function
    End If
    If Not rsSource Is Nothing Then
        Do While Not rsSource.EOF
            For i = 0 To rsSource.Fields.Count - 1
                If Not SaveOneValue(rsSource.Fields(i).value & "") Then
                    Exit Function
                End If
            Next
            If Not SaveRow() Then
                Exit Function
            End If
            rsSource.MoveNext
        Loop
    End If
    If Not SaveEnd() Then
        Exit Function
    End If
    SaveRec = True
End Function

'======================================================================================================================
'����           SaveStart            ���б���CSV�ļ�����ʼ������SaveStart��SaveRow��SaveEnd������ʹ��
'����ֵ         Boolean                 �Ƿ񴴽�CSV�ļ��ɹ�
'����б�:
'������         ����                    ˵��
'strFileName    String                  CSV�ļ�·��
'======================================================================================================================
Public Function SaveStart(ByVal strFileName As String, Optional ByVal blnOwverWrite As Boolean = True) As Boolean
    On Error GoTo ErrH
     If Not mobjSaveFile Is Nothing Then
        Call SaveEnd
    End If
    mlngRows = 0
    mlngCols = 0
    mblnFirstRow = True
    Set mobjSaveFile = mobjFSO.OpenTextFile(strFileName, IIf(blnOwverWrite, ForWriting, ForAppending), True)
    SaveStart = True
    Exit Function
ErrH:
    mstrError = err.Description
    err.Clear
End Function
'======================================================================================================================
'����           CSVSaveValue            ���浥������ֵ��CSV�ļ�
'����ֵ         Boolean                 �Ƿ񱣴�ɹ�
'����б�:
'������         ����                    ˵��
'arrInput()     Variant                 Ҫ�����ֵ�����߱���Ķ��ֵ������
'˵�����ú�����SaveRow���������ڣ�SaveRow��Ϊ�����������һ�м�¼��CSVSaveValue��Ϊ�������һ�м�¼�еĲ���ֵ����Ҫ��ϵ���SaveRow�����������м�¼
'======================================================================================================================
Public Function SaveValues(ParamArray arrInput() As Variant) As Boolean
    Dim i       As Long, j  As Long

    For i = LBound(arrInput) To UBound(arrInput)
        If (varType(arrInput(i)) And vbArray) = vbArray Then
            For j = LBound(arrInput(i)) To UBound(arrInput(i))
                If Not SaveOneValue(CStr(arrInput(i)(j))) Then
                    Exit Function
                End If
            Next
        Else
            If Not SaveOneValue(CStr(arrInput(i))) Then
                Exit Function
            End If
        End If
    Next
    SaveValues = True
End Function

'======================================================================================================================
'����           SaveOneValue         ���浥��ֵ��CSV�ļ�
'����ֵ         Boolean                 �Ƿ񱣴�ɹ�
'����б�:
'������         ����                    ˵��
'strInput       String                  Ҫ�����ֵ�����߱���Ķ��ֵ������
'======================================================================================================================
Private Function SaveOneValue(strInput As String) As Boolean
    
    If mblnFirstRow Then
        mlngCols = mlngCols + 1
    ElseIf mlngCurCol >= mlngCols Then
        SaveOneValue = True
        mstrError = "(����)�У�" & mlngRows & "�������������" & mlngCols
        Exit Function
    End If
    mlngCurCol = mlngCurCol + 1
    If mlngCurCol > 1 Then
        mobjSaveFile.Write "," & ValueAdjust(strInput)
    Else
        mobjSaveFile.Write ValueAdjust(strInput)
    End If
    SaveOneValue = True
End Function
'======================================================================================================================
'����           SaveRow              ���浥�е�CSV�ļ�
'����ֵ         Boolean                 �Ƿ񱣴�ɹ�
'����б�:
'������         ����                    ˵��
'arrInput()     Variant                  Ҫ����ĵ��е�ÿ�У���ֻ��һ����������������Ϊ�������ͣ��򽫸������Ԫ����Ϊÿ�У�����Ԫ�ؾ���Ϊ�ַ�������
'˵����CSV��ÿ������������ͬ���Ե�һ�ε��øú�����������Ϊȱʡ���������㲹�룬���˽ضϡ�
'      ��û�д����κβ�������Ĭ�����һ�����з�
'======================================================================================================================
Public Function SaveRow(ParamArray arrInput() As Variant) As Boolean
    Dim i       As Long, j  As Long

    For i = LBound(arrInput) To UBound(arrInput)
        If (varType(arrInput(i)) And vbArray) = vbArray Then
            For j = LBound(arrInput(i)) To UBound(arrInput(i))
                If Not SaveOneValue(CStr(arrInput(i)(j))) Then
                    Exit Function
                End If
            Next
        Else
            If Not SaveOneValue(CStr(arrInput(i))) Then
                Exit Function
            End If
        End If
    Next
    If mblnFirstRow Then
        mblnFirstRow = False
    ElseIf mlngCurCol < mlngCols Then
        For i = mlngCurCol + 1 To mlngCols
            If Not SaveOneValue("") Then
                Exit Function
            End If
        Next
        mstrError = "(����)�У�" & mlngRows & "����������㣺" & mlngCols
    End If
    mobjSaveFile.WriteLine
    mlngRows = mlngRows + 1
    mlngCurCol = 0
    SaveRow = True
End Function
'======================================================================================================================
'����           SaveEnd              ���CSV�ļ������
'����ֵ         Boolean                 �Ƿ񱣴�ɹ�
'����б�:
'������         ����                    ˵��
'======================================================================================================================
Public Function SaveEnd() As Boolean
    mobjSaveFile.Close
    Set mobjSaveFile = Nothing
    SaveEnd = True
End Function

'======================================================================================================================
'����           ValueAdjust             ����Ϊ�����CSV��ʽ
'����ֵ         String                  CSV��ʽ�ַ�
'����б�:
'������         ����                    ˵��
'strInput       String                  ��Ҫת��ΪCSV��ֵ
'======================================================================================================================
Private Function ValueAdjust(ByRef strInput As String) As String
    If InStr(strInput, """") > 0 Then
        ValueAdjust = """" & Replace(strInput, """", """""") & """"
    Else
        If InStr(strInput, ",") > 0 Then
            ValueAdjust = """" & strInput & """"
        ElseIf InStr(strInput, vbCr) > 0 Then
            ValueAdjust = """" & strInput & """"
        ElseIf InStr(strInput, vbLf) > 0 Then
            ValueAdjust = """" & strInput & """"
        Else
            ValueAdjust = strInput
        End If
    End If
End Function

'======================================================================================================================
'����           LexCSVContens           CSV�﷨������,ʹ������״̬��
'����ֵ         String                  �﷨����������Ľ��
'����б�:
'������         ����                    ˵��
'======================================================================================================================
Private Function LexCSVContens() As Boolean
    Dim strLine     As String
    Dim i           As Long
    Dim strOneChar  As String
    Dim stCur       As StateType
    
    strLine = ReadLine & vbCrLf
    mclsBuffer.Clear
    i = 1
    Do
        strOneChar = Mid(strLine, i, 1)
        Select Case stCur
            Case NewFieldStart
                Select Case strOneChar
                    Case """"
                        stCur = QuotesField
                    Case ","
                        Call ReadOneValue("")
                        stCur = FieldSeparator
                    Case vbCr, vbLf
                        mstrError = "���ڿ���"
                        stCur = LexError
                    Case Else
                        mclsBuffer.Append strOneChar
                        stCur = NonQuotesField
                End Select
            Case NonQuotesField
                Select Case strOneChar
                    Case ","
                        Call ReadOneValue(mclsBuffer.ToString())
                        mclsBuffer.Clear
                        stCur = FieldSeparator
                    Case vbCr
                        Call ReadOneValue(mclsBuffer.ToString())
                        stCur = RowSeparator
                    Case Else
                        mclsBuffer.Append strOneChar
                End Select
            Case QuotesField
                Select Case strOneChar
                    Case """"
                        stCur = QuoteInQuotesField
                    Case Else
                        mclsBuffer.Append strOneChar
                End Select
            Case FieldSeparator
                Select Case strOneChar
                    Case ","
                        Call ReadOneValue("")
                    Case """"
                        mclsBuffer.Clear
                        stCur = QuotesField
                    Case vbCr
                        Call ReadOneValue("")
                        stCur = RowSeparator
                    Case Else
                        mclsBuffer.Append strOneChar
                        stCur = NonQuotesField
                End Select
            Case QuoteInQuotesField
                Select Case strOneChar
                    Case ","
                        Call ReadOneValue(mclsBuffer.ToString())
                        mclsBuffer.Clear
                        stCur = FieldSeparator
                    Case vbCr
                        Call ReadOneValue(mclsBuffer.ToString())
                        stCur = RowSeparator
                    Case """" 'ת���ַ�
                        mclsBuffer.Append strOneChar
                        stCur = QuotesField
                    Case Else
                        mstrError = "�﷨���� ת���ַ�""�������ת��������ֶν�β����û�н����ֶηָ���"
                        stCur = LexError
                End Select
            Case RowSeparator
                Select Case strOneChar
                    Case vbLf
                        '��ȡ���һ��
                        mclsBuffer.Clear
                        stCur = NewFieldStart
                        If mblnFirstRow Then mblnFirstRow = False
                        mlngRows = mlngRows + 1
                        mlngCurCol = 0
                    Case Else
                        mstrError = "�﷨�����зָ����˻س�vbCr����δʹ�ûس�����vbCrLf"
                        stCur = LexError
                End Select
        End Select
        If stCur = LexError Then
            Exit Do
        End If
        If i = Len(strLine) Then
            If stCur <> NewFieldStart Then
                If AtEndOfStream Then
                    mstrError = "�﷨�������һ�м�¼�ֶ��а����س�vbCr����vbLf,����δʹ��˫����������"
                    stCur = LexError
                    Exit Do
                End If
                strLine = strLine & ReadLine & vbCrLf
            Else
                i = 0
                If Not AtEndOfStream Then
                    strLine = ReadLine & vbCrLf
                Else '������������ļ�
                    Exit Do
                End If
            End If
        End If
        i = i + 1
    Loop While (i <= Len(strLine))
    LexCSVContens = stCur <> LexError
    Exit Function
End Function
'======================================================================================================================
'����           ReadOneValue            ��ȡһ���ַ�������
'����ֵ         Boolean                 �Ƿ񱣴�ɹ�
'����б�:
'������         ����                    ˵��
'strInput       String                  Ҫ�����ֵ
'======================================================================================================================
Private Sub ReadOneValue(ByRef strValue As String)
    If mlngCurCol + mlngRows * mlngCols > UBound(marrValues) Then
        ReDim Preserve marrValues(UBound(marrValues) + 16 * (mlngCols + 1))
    End If
    If mblnFirstRow Then
        marrValues(mlngCurCol) = strValue
        mlngCols = mlngCols + 1
        mlngCurCol = mlngCurCol + 1
    ElseIf mlngCurCol < mlngCols Then
        marrValues(mlngCurCol + mlngRows * mlngCols) = strValue
        mlngCurCol = mlngCurCol + 1
    Else
        mstrError = "(����)�У�" & mlngRows & "��ȡ����������" & mlngCols
    End If
End Sub

Private Function ReadLine() As String
    If mobjTextLex Is Nothing Then
        ReadLine = mstrContent
    Else
        ReadLine = mobjTextLex.ReadLine
    End If
End Function

Private Function AtEndOfStream() As Boolean
    If mobjTextLex Is Nothing Then
        AtEndOfStream = True
    Else
        AtEndOfStream = mobjTextLex.AtEndOfStream
    End If
End Function

Private Sub Class_Initialize()
    ReDim Preserve marrValues(16)
End Sub

Private Sub Class_Terminate()
    If Not mobjSaveFile Is Nothing Then
        Call SaveEnd
    End If
    If Not mobjTextLex Is Nothing Then
        mobjTextLex.Close
        Set mobjTextLex = Nothing
    End If
    Set mclsBuffer = Nothing
    Set mobjFSO = Nothing
    Erase marrValues
End Sub

