VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AFS2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:�ɲ� CRP���
'����:�ɲ� AFS2000
'����:������
'�û�:��̨����ҽҽԺ
'����:2015-12-11

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'��׼���������ļ���ȡ������
'   strFile������·�����ļ���
'   lngSampleNO���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
'   dtStart����ʼʱ�䡣���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
'   dtEnd������ʱ�䡣ֻ�е�ѡȡ����걾��lngSampleNO=-1��ʱ���ò�����Ч�������ָ���ò���ֵ����ʱ�䷶ΧΪ>=dtStart��
'   ���صļ���������(������������밴���±�׼��֯���)
'   Ԫ��֮����||�ָ�
'   ��0��Ԫ�أ�����ʱ��
'   ��1��Ԫ�أ��������
'   ��2��Ԫ�أ�������
'   ��3��Ԫ�أ��걾
'   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
'   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
'       �磺��5i��Ԫ��Ϊ������Ŀ����5i + 1��Ԫ��Ϊ������

'΢����ӿڴ˴��� ������|���, �����صĽ����ʽΪ
'      ҩ�����^ҩ������^���ֵ

'΢������ʾ��:
'    ҩ�����ȡֵΪ:R,S,I ,ESBL,BLAC,R*��R-��ҩ��S-���У�I-�н飬
'��������       ESBL �¡�������ø,BLAC δ֪��R���������ȷ
'              ҩ������ȡֵΪ:1,2,3��1-MIC��2-DISK ��3-K-B


    
    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset, yp As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String, strlog As String, i As Integer
    Dim blnUnion As Boolean
    Dim strDns As String, strUser As String, strPassWord As String
    
    On Error GoTo errHand
    
    strUser = "admin": strPassWord = "afs2000"
    
    Err = 0
    
    'д��־
    strlog = "�Ѿ����ò���"
    Call writelislog(strlog)
    ReDim Preserve aResults(0)
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";uid=" & strUser & ";pwd=" & strPassWord & ";Driver={Microsoft Access Driver (*.mdb)}"
'        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        
        If .State = adStateClosed Then Exit Function

        'д��־
        strlog = "�Ѿ����ļ�"
        Call writelislog(strlog)
        lngSampleNO = -1
        If .State = adStateClosed Then Exit Function
        strSQL = "": blnUnion = False
        strSQL = "SELECT TDate+' '+ TTime As ��������,SampleID As �걾��,OutName As ͨ����,Result As ������ from SamplesHis A where " & vbNewLine & _
                IIf(lngSampleNO = -1, "  Format(TDate,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                " And Format(TDate,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'", _
                " SampleID='" & Format(lngSampleNO, "00") & "'") & _
                " Union all SELECT TDate+' '+ TTime As ��������,SampleID As �걾��,OutName As ͨ����,Result As ������ from Samples A where " & vbNewLine & _
                IIf(lngSampleNO = -1, "  Format(TDate,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                " And Format(TDate,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'", _
                " SampleID='" & Format(lngSampleNO, "00") & "'") & _
                " order by ��������,�걾�� desc"

        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        Do While Not rsTmp.EOF
            If strBh <> rsTmp("�걾��") Then  '�±걾
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(rsTmp("��������"), "yyyy-mm-dd")  '����ʱ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val(rsTmp("�걾��")) '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    " " '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    "ѪҺ" '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    "0" '�ʿ�
                    
                strBh = rsTmp("�걾��")
            End If
            '���Խ��[^OD^CutOff^sco]
            aResults(lngResultCount) = aResults(lngResultCount) & _
                 IIf(rsTmp("������") = "", "", "|" & UCase(rsTmp("ͨ����")) & "|" & rsTmp("������"))
            rsTmp.MoveNext
        Loop
    End With
   'д��־
    strlog = "���һ�����ݣ�" & aResults(lngResultCount)
    Call writelislog(strlog)
    
    ResultFromFile = aResults
    
    Exit Function
errHand:
    WriteLog "ResultFromFile", "", Err.Number & " " & Err.Description
End Function

Public Function writelislog(strlog As String)
    Const strFile As String = "c:\lislog_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim strFilename As String
    strFilename = strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFilename) Then Call objFile.CreateTextFile(strFilename)
    Set objStream = objFile.OpenTextFile(strFilename, ForAppending)
    objStream.WriteLine (strlog)
    objStream.Close
    Set objStream = Nothing

End Function


