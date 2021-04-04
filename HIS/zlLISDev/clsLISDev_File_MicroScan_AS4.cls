VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_MicroScan_AS4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:Microscan AS-4 ΢���������
'����:��������(DADE BEHRING INC)
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
    
    Dim aResults() As String
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim lngResultCount As Long
    Dim objStream As Scripting.TextStream, strLine As String
    
    Dim str�걾�� As String, strҩ����Ŀ As String
    Dim str���� As String, strϸ�� As String, str���ж� As String
    If Dir(strFile) = "" Then Exit Function
    Set objStream = objFile.OpenTextFile(strFile, ForReading)

    lngResultCount = -1
    ReDim aResults(0)
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If strLine Like "R,*" Then
            
            ''������ϸ�����ϵ������ �ȷ�����һ��ϸ�����
            If Val(str�걾��) > 0 And strϸ�� <> "" And str���� <> "" And strҩ����Ŀ <> "" Then
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
'                aResults(lngResultCount) = str���� & "|" & Val(str�걾��) & "| |" & strϸ�� & "|0" & strҩ����Ŀ
                '����Ժ�޸�ȡ�걾�ŵĺ�4λ
                aResults(lngResultCount) = str���� & "|" & Val((str�걾��)) & "| |" & strϸ�� & "|0" & strҩ����Ŀ
                str���� = "": str�걾�� = "": strϸ�� = "": strҩ����Ŀ = ""
            End If
            
            'һ���걾��ʼ,ϸ�����
            
            str�걾�� = Trim(Replace(Split(strLine, ",")(3), """", ""))
            str�걾�� = Mid(str�걾��, 1, Len(str�걾��) - 1)
            str���� = Replace(Split(strLine, ",")(6), """", "")
            str���� = Mid(str����, 1, 4) & "-" & Mid(str����, 5, 2) & "-" & Mid(str����, 7)
            'by����ǿ,��ԭ����ϸ�������޸�Ϊϸ�����.ԭ����ȡ(12)
            strϸ�� = Replace(Split(strLine, ",")(11), """", "")
            
            If str���� >= Format(dtStart, "yyyy-MM-dd") And str���� <= Format(dtEnd, "yyyy-MM-dd") Then
                
                If lngSampleNO <> -1 Then 'Ҫָ���걾��
                    If str�걾�� <> lngSampleNO Then
                        '�걾�Ų�ͬ��������
                        str�걾�� = ""
                        str���� = ""
                        strϸ�� = ""
                    End If
                End If
            Else
                '���ڲ�����������
                str�걾�� = ""
                str���� = ""
                strϸ�� = ""
            End If
            
        ElseIf strLine Like "M,*" Then
            'ҩ�����
            If str�걾�� <> "" Then
            '�޸�ϸ��������,����r*,blac,esbl ��������ҩby����ǿ.
                str���ж� = Replace(Split(strLine, ",")(7), """", "")
                Select Case str���ж�
                       Case "R*": str���ж� = "R"
                       Case "BLAC": str���ж� = "R"
                       Case "ESBL": str���ж� = "R"
                End Select
                strҩ����Ŀ = strҩ����Ŀ & "|" & Replace(Split(strLine, ",")(2), """", "") & "|" & str���ж� & "^1" & "^" & Replace(Split(strLine, ",")(4), """", "")
'                strҩ����Ŀ = strҩ����Ŀ & "|" & Replace(Split(strLine, ",")(2), """", "") & "|" & Replace(Split(strLine, ",")(7), """", "") & "^1" & "^" & Replace(Split(strLine, ",")(4), """", "")
            
            End If
        ElseIf strLine Like "L,*" Then
            'һ���걾����
            If Val(str�걾��) > 0 And strϸ�� <> "" And str���� <> "" And strҩ����Ŀ <> "" Then
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                aResults(lngResultCount) = str���� & "|" & Val(str�걾��) & "| |" & strϸ�� & "|0" & strҩ����Ŀ
                str���� = "": str�걾�� = "": strϸ�� = "": strҩ����Ŀ = ""
            End If
        End If
    Loop
    ResultFromFile = aResults
    
End Function
