VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_SWIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:΢����
'����:�Ϻ��¼��пƼ���չ���޹�˾ ��ϵ��:(��� 13795478025)
'����:����
'�û�:�����ж�Ժ
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
    Dim varStr As Variant
    Dim lngloop As Long
    Dim strTmp As String
    Dim strData As String           '��������
    Dim intCount As Integer         '����
    
    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    Open strFile For Input As #1 ' ���ļ���
    Do While Not EOF(1) ' ѭ�����ļ�β��
        Line Input #1, strLine ' ����һ�����ݲ����丳��ĳ������
        varStr = Split(strLine, Chr(9))
        
        '�����ڷ�Χ�ھ��˳�
        If CDate(Trim(varStr(39))) < dtStart Or CDate(Trim(varStr(39))) > dtEnd Then
            Close #1
            Exit Function
        End If
        
        '�滻A=1 B=2 C=3 D=4
        strTmp = Replace(Replace(Replace(Replace(varStr(7), "A", 1), "B", 2), "C", 3), "D", 4)
        strData = Trim(varStr(39)) & "|" & Trim(varStr(6) & strTmp) & "^0^| |" & varStr(9) & "|0"
        
        For lngloop = 40 To UBound(varStr)
            If Trim(varStr(lngloop)) <> "" Then
                intCount = intCount + 1
                strData = strData & "|" & Trim(varStr(lngloop))
                If intCount = 2 Then
                    strData = strData & "^^"
                    intCount = 0
                End If
            End If
        Next
        
        strData = Replace(strData, "RESIST", "R")
        strData = Replace(strData, "SUSC", "S")
        strData = Replace(strData, "INTER", "I")
        '�������͵ģ���ʱ������
        strData = Replace(strData, "NOINTP", "")
        strData = Replace(strData, "NOMIC", "")
        aResults(UBound(aResults)) = strData
        ReDim Preserve aResults(UBound(aResults) + 1)
    Loop
    Close #1 ' �ر��ļ���

    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_LightCycleII.ResultFromFile", strFile, Err.Number & " " & Err.Description
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

