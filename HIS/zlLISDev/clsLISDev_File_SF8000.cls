VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_SF8000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:��Ѫ������ SF-8000
'����:
'����:�¶�
'�û�:�����ο�ҽԺ
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
    
    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean, strImg As String, iField As Integer, i As Integer
    
    Dim cnErs As New ADODB.Connection
    Dim strFileErs As String, rsErs As ADODB.Recordset
    Dim strImgErs As String, str���� As String
    Dim strChar As String
    Dim intMaxValue As Integer, strFormat As String
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    strFileErs = strFile
    If Dir(App.Path & "\SF8000.ini") <> "" Then
        strFormat = ReadIni("SF8000", "��ʽ", App.Path & "\SF8000.ini", "")
    Else
        strFormat = ""
        Call WriteIni("SF8000", "��ʽ", "", App.Path & "\SF8000.ini")
    End If
    If Trim(Dir(strFileErs)) <> "" Then
        With cnErs
        If .State <> adStateClosed Then .Close
        .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then strFileErs = ""
        End With
    Else
        strFileErs = ""
    End If
        
    ReDim aResults(0)
    
    If strFileErs <> "" Then
        If strFormat = "�����ֶ�" Then
            strSQL = "Select �걾�� as ˳���,�������� as ����,��Ŀ����+��λ as ��Ŀ,���Խ�� as ��� From TB_TEST_RESULT " & _
                   " Where �������� >= #" & Format(dtStart, "yyyy-MM-dd") & "#" & _
                     " And �������� <= #" & Format(dtEnd, "yyyy-MM-dd") & "#" & _
                    IIf(lngSampleNO = -1, "", " And �걾�� =" & lngSampleNO)
        Else
            strSQL = "Select TEST_ID as ˳���,RESULT_DATE as ����,RESULT_NAME+RESULT_UNIT as ��Ŀ,RESULT_VALUE as ��� From TB_TEST_RESULT " & _
                   " Where RESULT_DATE >= #" & Format(dtStart, "yyyy-MM-dd") & "#" & _
                     " And RESULT_DATE <= #" & Format(dtEnd, "yyyy-MM-dd") & "#" & _
                    IIf(lngSampleNO = -1, "", " And TEST_ID =" & lngSampleNO)
        End If
        Set rsErs = cnErs.Execute(strSQL)
        Do Until rsErs.EOF
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            
            str���� = Format("" & rsErs("����"), "yyyy-MM-dd")
            aResults(UBound(aResults)) = str����   '����ʱ��
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val("" & rsErs("˳���"))   '�������
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| |Ѫ��|0"  '������,�걾,�ʿ�
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Replace(Trim("" & rsErs!��Ŀ), "g/L", "") & "|" & Val("" & rsErs!���)

            rsErs.MoveNext
        Loop
    End If
        

    ResultFromFile = aResults
End Function



