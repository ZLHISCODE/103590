VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_MP_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:΢��Ԫ��
'����:??
'����:����
'�û�:�Ͼ���������ҽԺ
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
    Dim strFileName As String
    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str�걾�� As String, str������ As String, blnReturn As Boolean
    Dim varStr As Variant
    Dim lngloop As Long
    Dim strTmp As String
    Dim strData As String           '��������
    Dim intCount As Integer         '����
    Dim blnBegin As Boolean         '�Ƿ�ʼ����
    Dim blnItem As Boolean          '�Ƿ�ʼһ����Ŀ
    Dim intLoop As Integer
    
    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = Mid(strFile, 1, InStrRev(strFile, "\"))
    strFileName = Dir(strFilePath & "*.bdt")
    
    Do While strFileName <> ""
        strData = ""
        blnBegin = False
        Open strFilePath & strFileName For Input As #1 ' ���ļ���
        Do While Not EOF(1) ' ѭ�����ļ�β��
            Line Input #1, strLine ' ����һ�����ݲ����丳��ĳ������
            strData = strData & strLine & vbCrLf
        Loop
        Close #1 ' �ر��ļ���
        varStr = Split(strData, vbCrLf)
        
        '�����ڷ�Χ�ھ��˳�
        If CDate(Trim(varStr(8))) >= dtStart And CDate(Trim(varStr(8))) <= dtEnd Then
            strData = Format(varStr(8), "yyyy-mm-dd") & "|" & Mid(varStr(12), Len(varStr(12)) - 3) & "| |ȫѪ|0"
            
            For intLoop = 13 To UBound(varStr)
                If varStr(intLoop) = "(null)" Then
                    blnBegin = True
                End If
                If varStr(intLoop) = "ѪǦ" And blnBegin = True Then
                    strData = strData & "|PB" & "|" & Val(Trim(varStr(intLoop + 1)))
                    blnItem = True
                End If
                
                If varStr(intLoop) = "п(Zn)" And blnBegin = True Then
                    strData = strData & "|Zn" & "|" & Val(Trim(varStr(intLoop + 1)))
                    blnItem = True
                End If
                
                If varStr(intLoop) = "��(Fe)" And blnBegin = True Then
                    strData = strData & "|Fe" & "|" & Val(Trim(varStr(intLoop + 1)))
                    blnItem = True
                End If
                
                If varStr(intLoop) = "��(Ca)" And blnBegin = True Then
                    strData = strData & "|Ca" & "|" & Val(Trim(varStr(intLoop + 1)))
                    blnItem = True
                End If
                
                If varStr(intLoop) = "þ(Mg)" And blnBegin = True Then
                    strData = strData & "|Mg" & "|" & Val(Trim(varStr(intLoop + 1)))
                    blnItem = True
                End If
                
                If varStr(intLoop) = "ͭ(Cu)" And blnBegin = True Then
                    strData = strData & "|Cu" & "|" & Val(varStr(intLoop + 1))
                    blnItem = True
                End If
            Next
            aResults(UBound(aResults)) = strData
            ReDim Preserve aResults(UBound(aResults) + 1)
        End If
        
        strFileName = Dir
    Loop

    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_MP_2.ResultFromFile", strFile, Err.Number & " " & Err.Description
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



