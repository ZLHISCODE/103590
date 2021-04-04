VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_XI921CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:XI-921CT
'����:���ڿ���
'����:�µ���
'�û�:�Ĵ�ʡ��ҵҽԺ
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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    Dim i As Integer
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����

    Dim iStrPoint As Integer, intSpot As Integer
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String
    Dim iPos As Long, strTmp As String, dblValue As Double, strItemValue As String
    Dim strLog As String
    Dim strLine As String
    Dim intStart As Integer, intEnd As Integer
    Dim aData() As String, aItem() As String
    Dim intLoop As Integer
    Dim strItem As String
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": strTmp = "": strData = ""
    
    strLog = Format(Now, "yyyy-MM-dd HH:mm:ss") & " ����ԭʼ���ݣ�" & strSource
    Call writelislog(strLog)
    
    If Len(strSource) < 62 Then strReserved = strSource: Exit Sub
    
    If InStr(strSource, Chr(&HD)) = 0 Then
      strSource = strSource & Chr(&HD) & Chr(&HA)
    End If
   
    strLog = Format(Now, "yyyy-MM-dd HH:mm:ss") & " ����ԭʼ���ݣ�" & strSource
    Call writelislog(strLog)
    
    aData = Split(strSource, Chr(&HD) & Chr(&HA))
    For intLoop = 0 To UBound(aData)
        strTmp = aData(intLoop)
        strTmp = Replace(strTmp, Chr(&H20) & Chr(&H20), Chr(&H20)) 'Replace(strTmp, "  ", " ")
        aItem = Split(Trim(strTmp), Chr(&H20))
        For i = 0 To UBound(aItem)
            If Val(aItem(0)) <> 0 Then
               Select Case i
                 Case 0: strData = strData & "||" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "|" & Val(aItem(i)) & "| |Ѫ��|0"
                 Case 3: strData = strData & "|K|" & Val(aItem(i))
                 Case 4: strData = strData & "|Na|" & Val(aItem(i))
                 Case 5: strData = strData & "|Cl|" & Val(aItem(i))
                 Case 6: strData = strData & "|iCa|" & Val(aItem(i)) & "|TCa|" & Val(Val(aItem(i)) * 1.95)
                 Case 7: strData = strData & "|pH|" & Val(aItem(i))
               End Select
            End If
        Next
    Next
    strResult = Mid(strData, 3)
    strLog = "���������ݣ�" & strResult
    Call writelislog(strLog)
    strCmd = ""
End Sub

Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_XI921CT_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As TextStream
    Dim strFileName As String
    strFileName = strFile & Format(Date, "yyyy-mm-dd") & ".txt"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function