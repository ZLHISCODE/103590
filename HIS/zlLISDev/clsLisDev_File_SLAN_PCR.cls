VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_File_SLAN_PCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:SLANӫ�ⶨ��PCR���ϵͳ
'����:��ʯ
'����:�ܺ�ȫ
'�û�:����������ҽԺ
'ʱ��:2012-03-11
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
'       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������("^" & ODֵ & "^" & COֵ & "^" & S/Co)
    
    Dim aResults() As String, objFile As New Scripting.FileSystemObject
    Dim str�걾 As String, str����Ա As String, str���� As String, strDate As String
    Dim strFilePath As String, strNewFile As String
    
    Dim xlApp As Object            'EXCEL�ļ�
    Dim xlBook As Object
    Dim xlSheet As Object
    
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str������ As String, str������ As String, blnReturn As Boolean
    Dim lngRow As Long, varLine As Variant
    Dim strFileImage As String, strImage As String
    
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    Set xlApp = CreateObject("Excel.Application")
    
    '2012-03-11 ZHQ �豸�ṩ�ļ�ΪXLS�ļ���ʽ
    strNewFile = strFilePath & "\slan.xls"
    If Dir(strNewFile) = "" Then
        Set xlBook = xlApp.Workbooks.Open(strFile)
    Else
        Set xlBook = xlApp.Workbooks.Open(strNewFile)
    End If
    Set xlSheet = xlBook.Worksheets(1)
    
    '�׺�   ͨ��    �������    �����Ŀ(4)     ����    ѭ����ֵ    �������(7) ƽ��Ũ��    ��������    ������ֵ    ��λ    CVֵ(%) �걾Ψһ��ʶ(13)    ����    �Ա�    ����
    'B3     1       ���Զ���    HBV-DNA         No Ct   <1000       ��������    1000        Copies/ml
    'B4     1       ���Զ���    HBV-DNA         28.65   1.04E+05    ��������    1000        Copies/ml
    'B5     1       ��Ʒ        HBV-DNA         24.62   1.50E+06    ��������    1000        Copies/ml                               120229HBV01         ������  ��      51
    'B6     1       ��Ʒ        HBV-DNA         24.96   1.20E+06    ��������    1000        Copies/ml
    'B7     1       ��Ʒ        HBV-DNA         21.14   1.51E+07    ��������    1000        Copies/ml                               120229HBV02         �׷���  Ů      17
    'B8     1       ��Ʒ        HBV-DNA         No Ct   <1000       ��������    1000        Copies/ml
    
    str������ = ""
    str������ = ""
    lngRow = 2
    Do While True
        strImage = ""
        If xlSheet.cells(lngRow, 4) <> "" And xlSheet.cells(lngRow, 7) <> "" And xlSheet.cells(lngRow, 13) <> "" Then
            If aResults(UBound(aResults)) <> "" Then
                ReDim Preserve aResults(UBound(aResults) + 1)
            End If
            str���� = Mid(xlSheet.cells(lngRow, 13), 1, 6)
            str���� = Mid(Format(date, "YYYY"), 1, 2) & str����
            str���� = Mid(str����, 1, 4) & "-" & Mid(str����, 5, 2) & "-" & Mid(str����, 7, 2)
            str������ = xlSheet.cells(lngRow, 13)
            If str������ <> "" Then
                str������ = "|" & xlSheet.cells(lngRow, 4) & "|" & GetENum(Trim(xlSheet.cells(lngRow, 7)))
                aResults(UBound(aResults)) = str���� & "|" & str������ & "| |ѪҺ|0" & str������
                
            End If
        ElseIf xlSheet.cells(lngRow, 4) = "" And xlSheet.cells(lngRow, 7) = "" And xlSheet.cells(lngRow, 13) = "" Then
            Exit Do
        End If
        
        lngRow = lngRow + 1
        DoEvents
    Loop
    xlBook.Close
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing

'    If str������ <> Mid(aResults(UBound(aResults)), 12, 3) And 1 = 1 Then
'        ReDim Preserve aResults(UBound(aResults) + 1)
'        aResults(UBound(aResults)) = str���� & "|" & str������ & "|" & str����Ա & "|" & str�걾 & "|0" & str������
'    End If
    ResultFromFile = aResults

End Function


Public Function GetENum(ByVal nNumber As Long) As String
'���ܣ����ݴ���Ĵ����ַ��ش���ӦС����10��n����
'������nNumber���磺15120000,Ҫ�󷵻أ�1.512E+07
'���أ���ֵ������
    
    Dim dbl10 As Double
    Dim i As Integer, j As Integer
    
    If nNumber >= 1 Then
        i = Len(CStr(nNumber))
        
        If i > 1 Then
            dbl10 = Round(nNumber / (10 ^ (i - 1)), 2)
            
            GetENum = dbl10 & "E+" & Right("00" & CStr(i - 1), 2)
        Else
            GetENum = nNumber
        End If
    Else
        'С���������
        j = 1
        Do While True
            dbl10 = nNumber * 10 ^ j
            If dbl10 >= 1 Then
                dbl10 = Round(dbl10, 2)
                GetENum = dbl10 & "E-" & Right("00" & CStr(j), 2)
            End If
            j = j + 1
        Loop
    End If
End Function