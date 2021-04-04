VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Coda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:CODAȫ�Զ�ø�����ϵͳ
'����:��������
'����:�¶�
'�û�:������ҽԺ
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
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
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
    Dim strChenel As String, strDate As String   'ͨ��,����Ա,����
    Dim strFilePath As String   '�ļ�·��
    Dim strLine As String
    Dim strSampleID As String, strValue As String   '�걾�ţ�������
    Dim lngLoop As Long
    Dim strData As String           'һ������
    Dim intCount As Integer         '����
     
    
    On Error GoTo errHandle
    
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    '
    lngLoop = DateDiff("d", dtStart, dtEnd)
    ReDim Preserve aResults(0)
    
    For intCount = 0 To lngLoop
        strDate = Format(DateAdd("d", intCount, dtStart), "ddMMyy")
        strFile = Dir(strFilePath & "\*" & strDate & "*.ast")
        Do While strFile <> ""
            
            Open strFilePath & "\" & strFile For Input As #1 ' ���ļ���
            Do While Not EOF(1) ' ѭ�����ļ�β��
                Line Input #1, strLine ' ����һ�����ݲ����丳��ĳ������
                If strLine Like "P|*|*|*|*" Then
                    strSampleID = Trim(Split(strLine, "|")(3))
                ElseIf strLine Like "O|*|*|*|*" Then
                    strDate = Split(strLine, "|")(22)
                    If strDate Like "##############" Then
                        strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2)
                    End If
                ElseIf strLine Like "R|*|*|*" Then
                    strChenel = Trim(Split(strLine, "|")(2))
                    strValue = Trim(Split(strLine, "|")(6))
                    
                    If IsDate(strDate) And strSampleID <> "" And strChenel <> "" And strValue <> "" Then
                        If strSampleID Like "##########" Then
                            strSampleID = Mid(strSampleID, 7)
                            strData = strDate & "|" & strSampleID & "| |Bold|0|"
                        Else
                            strData = strDate & "|" & strSampleID & "| |Bold|1|"
                        End If
                        If strChenel <> "" And strValue <> "" Then
                            strData = strData & strChenel & "|" & strValue
                            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                            aResults(UBound(aResults)) = strData
                        End If
                        strDate = ""
                        strSampleID = ""
                        
                    End If
                End If
            Loop
            Close #1 ' �ر��ļ���
            
            '�����Ѷ�ȡ���ļ���Ȼ��ɾ��
            If Not objFile.FolderExists(strFilePath & "\Bak") Then objFile.CreateFolder strFilePath & "\Bak"
            objFile.CopyFile strFilePath & "\" & strFile, strFilePath & "\Bak\" & strFile
            objFile.DeleteFile strFilePath & "\" & strFile
            
            strFile = Dir
        Loop
    Next
    'û���ҵ��ļ�ʱ�˳�
    
    
    ResultFromFile = aResults
    Exit Function
errHandle:
    'Resume
    WriteLog "CODA", strFile, Err.Number & " " & Err.Description
    Close #1
End Function

