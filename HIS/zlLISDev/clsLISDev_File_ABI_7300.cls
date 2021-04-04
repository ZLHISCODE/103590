VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_ABI_7300"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ABI7300
'����:����ABI
'����:����
'�û�:����Ժ
Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

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
    
    Dim aResults() As String, objFile As New Scripting.FileSystemObject
    Dim strChenel As String, str����Ա As String, str���� As String
    Dim strFilePath As String
    Dim strFilename As String
    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str�걾�� As String, str������ As String, blnReturn As Boolean
    Dim varStr As Variant
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strData As String           '��������
    Dim intCount As Integer         '����
    Dim blnBegin As Boolean         '�Ƿ�ʼ����
    Dim blnItem As Boolean          '�Ƿ�ʼһ����Ŀ
    Dim intLoop As Integer
    Dim aItem() As String           '��Ŀ
    Dim strDate As String           'ʱ��
    Dim strItem As String           '��Ŀ
    Dim strPath As String
    Dim str��� As String
    
    
    On Error GoTo errHandle
    'û���ҵ��ļ�ʱ�˳�
    'If Dir(strFile) = "" Then Exit Function
    strPath = Mid(strFile, 1, InStrRev(strFile, "\"))
    strFile = Dir(strPath)
    If strFile <> "" Then
        ReDim Preserve aResults(0)
    End If
    ReDim Preserve aResults(0)
    Do While strFile <> ""
        
        Open strPath & strFile For Input As #1  ' ���ļ���
        Do While Not EOF(1) ' ѭ�����ļ�β��
            Line Input #1, strLine ' ����һ�����ݲ����丳��ĳ������
            aItem = Split(strLine, ",")
            If UBound(aItem) >= 0 Then
                'ʱ��
                If Mid(aItem(0), 1, 9) = "Run Date:" Then
                    strDate = DateValue(aItem(1) & "," & aItem(2)) & " " & Right(aItem(2), 8)
                End If
                '���
                If UBound(aItem) = 10 Then
                    If aItem(1) <> "" Then
                        str�걾�� = Val(GetNumber(Mid(aItem(1), 4)))
                        If str�걾�� <> 0 Then
                            strItem = Trim(Mid(aItem(1), 1, 3))
                            '���� ��ѧ��������С��λ��
                            str��� = IIf(aItem(6) = "", " ", aItem(6))
                            If InStr(UCase(str���), "E") > 0 Then
                                str��� = Format(Val(Mid(str���, 1, InStr(UCase(str���), "E") - 1)), "0.00") & Mid(str���, InStr(UCase(str���), "E"))
                            End If
                            strData = strDate & "|" & str�걾�� & "| |Ѫ��|0|" & strItem & "|" & UCase(str���)
                            Debug.Print strData
                            aResults(UBound(aResults)) = strData
                            ReDim Preserve aResults(UBound(aResults) + 1)
                        End If
                    End If
                End If
            End If
        Loop
        Close #1 ' �ر��ļ���
        strFile = Dir
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    Resume
    WriteLog "clsLISDev_File_ABI_7300.ResultFromFile", strFile, Err.Number & " " & Err.Description
    Close #1
End Function


Private Function TrunMonth(ByVal strMonth As String) As String
    'ת�����ĵ��·�Ϊ����
    Select Case UCase(strMonth)
        Case "һ��": TrunMonth = "01"
        Case "����": TrunMonth = "02"
        Case "����": TrunMonth = "03"
        Case "����": TrunMonth = "04"
        Case "����": TrunMonth = "05"
        Case "����": TrunMonth = "06"
        Case "����": TrunMonth = "07"
        Case "����": TrunMonth = "08"
        Case "����": TrunMonth = "09"
        Case "ʮ��": TrunMonth = "10"
        Case "ʮһ": TrunMonth = "11"
        Case "ʮ����": TrunMonth = "12"
        
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select
End Function

Private Function GetNumber(strValue As String) As String
    Dim intLoop As Integer
    For intLoop = 1 To Len(strValue)
        If IsNumeric(Mid(strValue, intLoop, 1)) = True Then
            GetNumber = GetNumber & Mid(strValue, intLoop, 1)
        End If
    Next
End Function





