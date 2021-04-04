VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_EStar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ȫ�Զ�ø����
'����:��ʿ������(HAMILTON)
'����:�¶�
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
    
    Dim aResults() As String
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream, strLine As String
    Dim lngResultCount As Long
    Dim str���� As String, str������Ŀ As String, str���Խ�� As String, str������� As String
    Dim blnOk As Boolean '�Ƿ񰴱걾����ȡ����
    Dim strFilePath As String
    Dim strCO As String 'cut-offֵ��config.iniȡֵ

    
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    
    If Dir(strFile) = "" Then Exit Function
    strFile = UCase(strFile)
    
    str������Ŀ = Replace(Replace(strFile, strFilePath, ""), ".TXT", "")
    If Left(str������Ŀ, 1) = "\" Then str������Ŀ = Mid(str������Ŀ, 2)
    '��config��ȡcut-offֵ

    Call AnalyConf(UCase(str������Ŀ), strCO)
    
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    lngResultCount = -1
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If InStr(strLine, "���ʱ��:") > 0 Then
            str���� = Mid(strLine, InStr(strLine, ":") + 1, 18)
            If IsDate(str����) Then
                str���� = Format(str����, "yyyy-MM-dd HH:mm:ss")
            Else
                str���� = Format(Now, "yyyy-MM-dd HH:mm:ss")
            End If
        End If
        If Val(Mid(strLine, 3, 3)) > 0 Then
            blnOk = False
            If lngSampleNO <> -1 Then
                If lngSampleNO = Val(Mid(strLine, 3, 3)) Then
                    blnOk = True
                End If
            Else
                blnOk = True
            End If
            
            If CDate(str����) >= dtStart And CDate(str����) <= dtEnd Then
                blnOk = True
            End If
            If blnOk Then
            
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                str���Խ�� = Trim(Mid(strLine, 20, 2))
                str���Խ�� = Switch(str���Խ�� = "+", "����(+)", str���Խ�� = "-", "����", str���Խ�� = "+-", "������(+-)")
                
                If InStr(Mid(strLine, 50), ".") <= 0 Then
                    str������� = Val(Mid(strLine, 30))
                Else
                    str������� = Val(Mid(strLine, 50))
                End If
                If Left(str�������, 1) = "." Then str������� = "0" & str�������
                aResults(lngResultCount) = str���� & "|" & Val(Mid(strLine, 3, 3)) & "| |ѪҺ|0|" & str������Ŀ & "|" & str���Խ�� & "^" & str������� & "^" & strCO & "^0"
            End If
            
        End If
    Loop
    
    ResultFromFile = aResults
End Function

Public Sub AnalyConf(strItem As String, strCO As String)
    Dim arrLine
    Dim strLine As String
    Dim strFile As String
    Dim blnOpen As Boolean
    Dim objFilesys As New FileSystemObject
    Dim objStream As TextStream
    On Error GoTo ErrHand
    Dim intloop As Long
    
    
    '�������ļ��ж�ȡ�������ļ��е�����cotf
    strFile = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Config.ini"
    If objFilesys.FileExists(strFile) Then
        Set objStream = objFilesys.OpenTextFile(strFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            If strLine = "" Then Exit Do
            arrLine = Split(strLine, "=")
            If arrLine(0) = strItem Then
               strCO = arrLine(1)
               Exit Sub
            End If
           
        Loop
        objStream.Close
        blnOpen = False
    End If
    
   
    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub