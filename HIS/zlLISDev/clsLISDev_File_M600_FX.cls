VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_M600_FX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:��Ѫ M600
'�û�:������ҽԺ

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
    Dim lngResultCount As Long, strBH As String
    Dim blnUnion As Boolean, strimg As String, iField As Integer, i As Integer
    
    Dim cnErs As New ADODB.Connection
    Dim strFileErs As String, rsErs As ADODB.Recordset
    Dim strImgErs As String, str���� As String, str��ˮ�� As String
    
    On Error GoTo Errhand
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        'ȡ���������ı걾
        strSQL = "select distinct a.SampleNo as ��ˮ��,SerialNo as �걾��,format(InputTime,'yyyy-MM-dd') as ���� from ToDayDetail a, ToDayMain b " & _
        "where a.SampleNo=b.SampleNo and Format(InputTime,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "' And Format(InputTime,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "' " & _
        IIf(lngSampleNO = -1, "", " And INSTID_ID= '" & lngSampleNO & "'")
        
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        
        lngResultCount = -1
        strBH = ""
        str��ˮ�� = ""
        Do While Not rsTmp.EOF
            If str��ˮ�� <> rsTmp("��ˮ��") Then '�±걾
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(rsTmp("����"), "yyyy-mm-dd")  '��������
                
                str���� = Format(rsTmp("����"), "yyyy-mm-dd")
                
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val("" & rsTmp("�걾��"))   '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      " "  '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      "ѪҺ"  '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|0" '& _
                   ' Val(rsTmp("�ʿ�")) '�ʿ�
                    
                str��ˮ�� = rsTmp("��ˮ��")
            End If
          '��ʼ��ѯ���
            strSQL = "select Alias as ��Ŀ,Second as ��� " & _
                      " from ToDayDetail a, ToDayMain b where a.SampleNo=b.SampleNo and a.SampleNo= '" & str��ˮ�� & "'" & _
                      " union all " & _
                      "select Unit2 as ��Ŀ,fResult2 as ��� " & _
                      " from ToDayDetail a, ToDayMain b where a.SampleNo=b.SampleNo and a.SampleNo= '" & str��ˮ�� & "' "
            Set rsErs = cnAccess.Execute(strSQL)
                Do While Not rsErs.EOF
                    If Trim(rsErs("��Ŀ")) <> "" Then
                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                             rsErs("��Ŀ") & "|" & Trim(Val(IIf(IsNull(rsErs("���")), "0", rsErs("���"))))
                    End If
                    rsErs.MoveNext
                Loop
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    
    Exit Function
Errhand:
    WriteLog "M600_FX", Err.Number & "��" & Err.Description, ""
End Function













