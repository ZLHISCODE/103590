VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_TRFIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:���ݷỪ
'����:
'����:������
'�û�:
'ʱ��:20151223
Public objFso As New FileSystemObject

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
   Dim FILE��ʼ���� As String, FILE�������� As String, strTemp As String
   
   Dim objStream As textStream, lngResultCount As Long
   Dim strLine As String, varItem As Variant, varTmp As Variant
   Dim str�걾�� As Long, blnReturn As Boolean
   Dim lngRow As Long, varLine As Variant
   Dim tempstr As String, NowFile As String, NewFile As String, str������ As String

  
   '��������±����
    ResultFromFile = aResults
    ReDim aResults(0)
    
    NewFile = strFile
    Set objStream = objFile.OpenTextFile(NewFile, ForReading, False, -2)
    lngRow = 0
    str���� = ""
    str������ = ""
    str�걾�� = 0


    FILE��ʼ���� = Year(dtStart) & IIf(Len(Month(dtStart)) = 1, 0 & Month(dtStart), _
                   Month(dtStart)) & IIf(Len(Day(dtStart)) = 1, 0 & Day(dtStart), Day(dtStart))
    FILE�������� = Year(dtEnd) & IIf(Len(Month(dtEnd)) = 1, 0 & Month(dtEnd), _
                   Month(dtEnd)) & IIf(Len(Day(dtEnd)) = 1, 0 & Day(dtEnd), Day(dtEnd))

              
  ''���ݿ��ʽ����
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim str As String
    
    On Error GoTo errHand
    
    cn.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & strFile '�������ݿ�    D:\Uranus\ Modules\�Ừʱ��ֱ�\Data.dat
    If lngSampleNO = -1 Then
         rs.Open "select a.A17 as A17 ,a.A2 as A2 ,b.C4 as  C4,b.C6 as C6 from h_Patient a inner join h_Result b on a.A1=b.C2 where len(b.C4)<>0   and len(b.C6)<>0   and  a.A17  between '" & FILE��ʼ���� & "' and '" & FILE�������� & "'", cn, adOpenKeyset, adLockOptimistic '�������ݱ�"
    Else
         rs.Open "select a.A17 as A17 ,a.A2 as A2 ,b.C4 as  C4,b.C6 as C6 from h_Patient a inner join h_Result b on a.A1=b.C2 where len(b.C4)<>0   and  len(b.C6)<>0   and b.C0= " & lngSampleNO & " and a.A17 = '" & FILE��ʼ���� & "' ", cn, adOpenKeyset, adLockOptimistic '�������ݱ�"
    End If
    If rs.RecordCount > 0 Then
         rs.MoveFirst
    End If
    
            Do While rs.EOF = False
             
               str���� = rs.Fields("A17")
               str���� = Mid(str����, 1, 4) + "-" + Mid(str����, 5, 2) + "-" + Mid(str����, 7, 2)
'               str���� = Format(str����, "YYYY-MM-DD")
               str�걾�� = Val(rs.Fields("A2"))
               '��Ⱦ8��  Trim(rs.Fields("C6") & " ")  IIf(Len(rs.Fields("C6")) = 0, " ", rs.Fields("C6")) Trim(rs.Fields("C4") & " ")
               str������ = "|" + rs.Fields("C4") + "|" + rs.Fields("C6")

             If str������ <> "" And str�걾�� > 0 And str���� <> "" Then
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


