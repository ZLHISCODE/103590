VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_BHX100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:BH-2100,5100
'����:����
'����:�¶�
'�û�:�ܺ�ȫ
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
    
    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str�걾�� As String, str������ As String, blnReturn As Boolean
    Dim lngRow As Long, varLine As Variant, str��ʽ As String
    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    str��ʽ = "5100"
    If InStr(strFile, "2100") > 0 Then str��ʽ = "2100"
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        
        If str��ʽ = "2100" Then
            varLine = Split(strLine, ",")
            If UBound(varLine) = 18 Then

                str�걾�� = Val(varLine(0))
                str���� = Replace(Trim(varLine(15)), """", "")
                If Val(str�걾��) > 0 And str���� <> "" Then
                    '"Ǧ����� ǦŨ�� ϡ�ͱ��� ѪǦ����
                    str���� = Format(CDate(str����), "yyyy-MM-dd")
                    If Not (str���� Like "####-##-##" And IsDate(str����)) Then
                        str���� = Format(Now, "yyyy-MM-dd")
                    End If
                    
                    str������ = "|" & varLine(1) & "|" & Format(Val(varLine(2)), "0.000") & "|" & varLine(3) & "|" & Format(Val(varLine(4)), "0.000")
                    
                    If str������ <> "|" & varLine(1) & "|0.000|" & varLine(3) & "|0.000" Then
                        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| | |0" & Replace(str������, """", "")
                    End If
                End If
            End If
        Else
           ' strLine = StrConv(strLine, vbNarrow)
            varLine = Split(strLine, Chr(&H9))
            
            If UBound(varLine) = 10 Then
                str�걾�� = Val(varLine(9))
                str���� = Mid(varLine(8), 1, 10)
                
                If Val(str�걾��) > 0 And str���� <> "" Then
                    'cu  zn  ca  mg  fe
                    
                    If Not (str���� Like "####-##-##" And IsDate(str����)) Then
                        str���� = Format(Now, "yyyy-MM-dd")
                    End If
                    
                    str������ = "|cu|" & Format(Val(varLine(3)), "0.000") & "|zn|" & Format(Val(varLine(4)), "0.000") & "|ca|" & Format(Val(varLine(5)), "0.000") & "|mg|" & Format(Val(varLine(6)), "0.000") & "|fe|" & Format(Val(varLine(7)), "0.000")
                    
                    If str������ <> "|cu|0.000|zn|0.000|ca|0.000|mg|0.000|fe|0.000" Then
                        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| | |0" & str������
                    End If
                End If
            End If
        End If
        lngRow = lngRow + 1
        DoEvents
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_BHX100.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function



