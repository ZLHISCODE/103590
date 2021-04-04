VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_U2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:U2
'ҽԺ:������ҽԺ
'����:����ǿ
'ʱ��:2009/6/17   �޸�ʱ�� 2013/12/19
'�޸��ˣ�������
'����:


Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'��׼���������ļ���ȡ������
'   strFile������·�����ļ���
'   lngSampleNO���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
'   dtStart����ʼʱ�䡣���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
'   dtEnd������ʱ�䡣ֻ�е�ѡȡ����걾��lngSampleNO=-1��ʱ���ò�����Ч�������ָ���ò���ֵ����ʱ�䷶ΧΪ>=dtStart��
'   ���صļ���������(������������밴���±�׼��֯���)
'   Ԫ��֮����|�ָ�
'   ��0��Ԫ�أ�����ʱ��
'   ��1��Ԫ�أ��������
'   ��2��Ԫ�أ�������
'   ��3��Ԫ�أ��걾
'   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
'   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
'       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������

    Dim aResults() As String
    Dim cnsql As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsItem As ADODB.Recordset

    Dim strResult As String         '���
    Dim strItem As String           '��Ŀ��ͨ����)
    Dim st���� As String, en���� As String, str���� As String
    Dim strSer As String, strDB As String, strPWD As String, strUserID As String
    Dim strValues As String, varValues As Variant
    Dim strItemValue As String, i As Integer
    Dim str�걾�� As String
    Dim str��� As String
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults

    On Error GoTo errH
    strSer = Val(ReadIni("U2", "SERVER", App.Path & "\U2.ini"))
    If strSer = "" Then strSer = "127.0.0.1": Call WriteIni("U2", "SERVER", "127.0.0.1", App.Path & "\U2.ini")

    strDB = Val(ReadIni("U2", "DATABASE", App.Path & "\U2.ini"))
    If strDB = "" Then strDB = "urine_20130516": Call WriteIni("U2", "DATABASE", "urine_20130516", App.Path & "\U2.ini")

    strUserID = Val(ReadIni("U2", "USERID", App.Path & "\U2.ini"))
    If strUserID = "" Then strUserID = "sa": Call WriteIni("U2", "USERID", "sa", App.Path & "\U2.ini")

    strPWD = Val(ReadIni("U2", "PWD", App.Path & "\U2.ini"))
    If strPWD = "" Then strPWD = "saa": Call WriteIni("U2", "PWD", "saa", App.Path & "\U2.ini")


    With cnsql
        Set cnsql = New ADODB.Connection
        If cnsql.State = adStateOpen Then cnsql.Close

        cnsql.ConnectionTimeout = 25
        cnsql.Provider = "sqloledb"
        cnsql.Properties("data source").Value = strSer    '������������
        cnsql.Properties("initial catalog").Value = strDB    '����
        'conn.Properties("integrated security").Value = "SSPI" '��½����
        cnsql.Properties("user id").Value = strUserID  '�û���
        cnsql.Properties("password").Value = strPWD     '����
        cnsql.Open
        '    st���� = Format(dtStart, "yyyy-mm-dd")
        '    en���� = Format(dtEnd, "yyyy-mm-dd") & " 23:59:59"
        strSQL = "select smpno,daterpt,testvalues_text from urinerpts where daterpt between ' " & dtStart & "' and '" & dtEnd & "'" & _
                    IIf(lngSampleNO = -1, " ", " and smpno = " & lngSampleNO)
        Set rsTmp = cnsql.Execute(strSQL)
        Do Until rsTmp.EOF
            strItemValue = ""
            varValues = Split(rsTmp!testvalues_text, Chr(10))
            str�걾�� = Val(rsTmp!smpno)
            str���� = Format(rsTmp!daterpt, "YYYY-MM-DD")
            
            For i = LBound(varValues) To UBound(varValues)
                strItem = Fun_���(varValues(i), "code")
                Select Case strItem
                Case "PRO", "BLD", "WBC", "NIT", "GLU", "KET", "URO", "BIL", "VC", "MA", "CR", "Ca"    'ȥ����ֵ����
                    str��� = Fun_���(varValues(i), "value")
                    If InStr(str���, "(") > 0 Then
                        str��� = Mid(str���, 1, InStr(str���, "(") - 1)
                    End If
                    If Fun_���(varValues(i), "pcode") = "micro" And strItem = "WBC" Then
                        strItem = strItem & 1
                    End If
                    If str��� <> "" Then
                        strItemValue = strItemValue & "|" & strItem & "|" & str���
                    End If
                Case Else
                    If Fun_���(varValues(i), "pcode") <> "root" And strItem <> "" Then
                        If strItem = "micro" Then
                            MsgBox ""
                        End If
                        strItemValue = strItemValue & "|" & strItem & "|" & Fun_���(varValues(i), "value")
                    End If
                End Select
            Next
            
            If strItemValue <> "" Then
                If aResults(UBound(aResults)) <> "" Then
                    ReDim Preserve aResults(UBound(aResults) + 1)
                End If
                strResult = str���� & "|" & str�걾�� & "| |��Һ|0" & strItemValue
                aResults(UBound(aResults)) = strResult
            End If
            
            rsTmp.MoveNext
        Loop
        ResultFromFile = aResults
    End With
    Exit Function
errH:
    WriteLog "U2", "ResultFromFile", Err.Number & "    " & Err.Description
End Function



Public Function Fun_���(ByVal StrInput As String, ByVal strKey As String) As String
    Dim str��� As String
    Dim strVar As Variant
    Dim i As Long
    Fun_��� = ""
    On Error GoTo errH
    If StrInput = "" Then Exit Function
    StrInput = Replace(Replace(Replace(StrInput, Chr(10), ""), Chr(13), ""), "<p", "p")
    strVar = Split(StrInput, ";")
    For i = LBound(strVar) To UBound(strVar)
        If Split(Replace(strVar(i), vbCrLf, ""), "=")(0) = strKey Then
            Fun_��� = Trim(Split(strVar(i), "=")(1))
            Exit Function
        End If
    Next i
    Exit Function
errH:
    WriteLog "U2", "Fun_���", Err.Number & "    " & Err.Description
End Function







