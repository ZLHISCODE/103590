VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_SA7000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:SA_7000
'����:
'����:������
'�û�:Ѯ����ҽԺ
'ʱ��:20151217
 
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
   Dim strFilePath As String, strFilename As String, k As Integer, int�ʿ��� As Integer, strRQ As String
   Dim strNeedFile As String  'ȷ��������Ҫ��ȡ�ı걾������
   Dim FILE��ʼ���� As Date, FILE�������� As Date, strTemp As String
   
   Dim objStream As textStream, lngResultCount As Long
   Dim strLine As String, varItem As Variant, varTmp As Variant
   Dim str�걾�� As Long, blnReturn As Boolean
   Dim lngRow As Long, varLine As Variant
   Dim tempstr As String, NowFile As String, NewFile As String, str��Ŀ As String, str������ As String

  


   '��������±����
    ResultFromFile = aResults
    ReDim aResults(0)
    
    NewFile = strFile
    Set objStream = objFile.OpenTextFile(NewFile, ForReading, False, -2)
    lngRow = 0
    str���� = ""
    str��Ŀ = ""
    str������ = ""
    str�걾�� = 0
            
  ''���ݿ��ʽ����
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim str As String
    
    On Error GoTo errHand
    
    cn.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & strFile '�������ݿ�
    If lngSampleNO = -1 Then
          rs.Open "select * from TestData where ���� between Cdate('" & Format(dtStart, "yyyy-MM-dd") & "') and  Cdate('" & Format(dtEnd, "yyyy-MM-dd") & "')", cn, adOpenKeyset, adLockOptimistic  '�������ݱ�
    Else
          rs.Open "select * from TestData where ���� = Cdate('" & Format(dtStart, "yyyy-MM-dd") & "') and TestDataID=" & lngSampleNO, cn, adOpenKeyset, adLockOptimistic  '�������ݱ�"
    End If
    If rs.RecordCount > 0 Then
         rs.MoveFirst
    End If

           Do While rs.EOF = False
           
               str���� = Format(rs.Fields("����") & " " & rs.Fields("ʱ��"), "yyyy-mm-dd  hh:mm:ss")
               str�걾�� = Val(rs.Fields("TestDataID"))

'                ��ǰ���
              str������ = "|Ѫ��ճ��|" + Trim(rs.Fields("Ѫ��ճ��") & "") + "|Ѫ��|" + Trim(rs.Fields("Ѫ��") & "") + "|ѹ��|" + Trim(Format(rs.Fields("ѹ��"), "0.00") & "") + "|ȫѪ�������ָ��|" + _
                             Trim(rs.Fields("ȫѪ�������ָ��") & "") + "|ȫѪ�������ָ��|" + Trim(rs.Fields("ȫѪ�������ָ��") & "") + "|Ѫ������Kֵ|" + Trim(rs.Fields("Ѫ������Kֵ") & "") + _
                             "|��ϸ���ۼ�ָ��|" + Trim(rs.Fields("��ϸ���ۼ�ָ��") & "") + "|ȫѪ���л�ԭճ��|" + Trim(rs.Fields("ȫѪ���л�ԭճ��") & "") + "|ȫѪ���л�ԭճ��|" + _
                             Trim(rs.Fields("ȫѪ���л�ԭճ��") & "") + "|��ϸ������ָ��|" + Trim(rs.Fields("��ϸ������ָ��") & "") + "|��ϸ������ָ��TK|" + Trim(Format(rs.Fields("��ϸ������ָ��TK"), "0.00") & "")

                 If str������ <> "" And str�걾�� <> 0 And str���� <> "" Then
                   If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                   aResults(UBound(aResults)) = str���� & "|" & Val(str�걾��) & "| |" & "Ѫ��|0" & str������
                   Debug.Print aResults(UBound(aResults))
                   str������ = ""
                   str���� = ""
                   str�걾�� = 0
                End If

               rs.MoveNext
           Loop

        cn.Close

    ResultFromFile = aResults
    
    Exit Function
errHand:
    WriteLog "ResultFromFile", "", Err.Number & " " & Err.Description
End Function



