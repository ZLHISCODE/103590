VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Evolis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:EVOLISȫ�Զ�ø�����ϵͳ
'����:��������
'����:����
'�û�:��ú
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
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strData As String           '��������
    Dim intCount As Integer         '����
    Dim blnBegin As Boolean         '�Ƿ�ʼ����
    Dim blnItem As Boolean          '�Ƿ�ʼһ����Ŀ
    Dim intloop As Integer
    Dim aitem() As String           '��Ŀ
    Dim strDate As String           'ʱ��
    Dim curCoutOff As Currency, curOD As Currency, curSCO As Currency
    On Error GoTo errHandle
    
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    '
    lngLoop = DateDiff("d", dtStart, dtEnd)
    ReDim Preserve aResults(0)
    
    For intCount = 0 To lngLoop
        str���� = Format(DateAdd("d", intCount, dtStart), "yyMMdd")
        strFile = Dir(strFilePath & "\*" & str���� & "*.TXT")
        Do While strFile <> ""
            
            Open strFilePath & "\" & strFile For Input As #1 ' ���ļ���
            Do While Not EOF(1) ' ѭ�����ļ�β��
                Line Input #1, strLine ' ����һ�����ݲ����丳��ĳ������
        
                strLine = Replace(strLine, """", "")
                strLine = Replace(strLine, """", "")
                aitem = Split(strLine, ",")
                
                If UBound(aitem) >= 0 Then
                    If aitem(0) = "ʱ��:" Then
                        strDate = aitem(1)
                    End If
                    
                    If aitem(0) = "����:" Then
                        strDate = aitem(1) & " " & strDate
                    End If
                    If aitem(0) = "Cut-off Value" Then
                        curCoutOff = Val(Trim(aitem(2)))
                    End If
                    If Val(aitem(0)) <> 0 And UBound(aitem) > 5 Then
                        '�������
                        If aitem(0) <> "" Then
                            curOD = Val(aitem(UBound(aitem) - 1))
                            If curCoutOff <> 0 Then
                                curSCO = curOD / curCoutOff
                            Else
                                curSCO = 0
                            End If
                            strData = strDate & "|" & aitem(0) & "| |Ѫ��|0|" & aitem(1) & "|" & aitem(UBound(aitem)) & "^" & curOD & "^" & curCoutOff & "^" & curSCO
                            Debug.Print strData
                            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                            aResults(UBound(aResults)) = strData
                        End If
                    End If
                End If
            Loop
            Close #1 ' �ر��ļ���
            
            strFile = Dir
        Loop
    Next
    'û���ҵ��ļ�ʱ�˳�
    
    
    ResultFromFile = aResults
    Exit Function
errHandle:
    'Resume
    WriteLog "clsLISDev_File_Evolis.ResultFromFile", strFile, Err.Number & " " & Err.Description
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







