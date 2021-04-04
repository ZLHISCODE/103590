VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_EH_2050B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:��Һ���� ȫ�Զ������������
'����:����
'����:������
'�û�:˷��ƽ³������ҽԺ
'����:2013-06-14
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
    Dim strFilePath As String, intCount As Long, lngLoop As Long
    Dim StrTime As String
    Dim objStream As TextStream
    Dim strLine As String, varTmp As Variant
    Dim str�걾�� As String, str������ As String, blnReturn As Boolean
    Dim lngRow As Long, varLine As Variant, strItems As String, strImages As String
    Dim intStep As Integer, strItem As String, strImage As String
    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    '
    lngLoop = DateDiff("d", dtStart, dtEnd)

    For intCount = 0 To lngLoop
        str���� = Format(DateAdd("d", intCount, dtStart), "yyyyMMdd")
        strFile = Dir(strFilePath & "\*" & str���� & "*.TXT")
        Do While strFile <> ""

            Open strFilePath & "\" & strFile For Input As #1    ' ���ļ���
            Do While Not EOF(1)    ' ѭ�����ļ�β��
                Line Input #1, strLine    ' ����һ�����ݲ����丳��ĳ������
                strLine = Replace(strLine, vbTab, "")
                Select Case intStep
                Case 1
                    If InStr(strLine, "Sediment Data Start") > 0 Then
                        intStep = 2
                    Else
                        strItem = Trim(Split(strLine, ",")(2))
                        str������ = Trim(Split(strLine, ",")(5))
                        If str������ <> "" Then
                            str������ = Trim(Mid(str������, 1, Len(str������) - 1))
                        End If
                    End If
                    If strItem <> "" And str������ <> "" Then
                        strItems = strItems & "|" & strItem & "|" & str������
                        strItem = ""
                        str������ = ""
                    End If
                Case 2
                    If InStr(strLine, "Image filename Start") > 0 Then
                        intStep = 3
                    Else
                        strItem = Trim(Split(strLine, ",")(2))
                        str������ = Trim(Split(strLine, ",")(3))
                    End If
                    If strItem <> "" And str������ <> "" Then
                        strItems = strItems & "|" & strItem & "|" & str������
                        strItem = ""
                        str������ = ""
                    End If
                Case 3
                    If InStr(strLine, "--End--") > 0 Then
                        intStep = 4
                    Else
                        If strImage = "" Then
                            intCount = 1
                            strImage = Trim(Split(strLine, ",")(2))
                            strImages = intCount & ";110;0;" & Trim(Mid(strImage, 1, Len(strImage) - 1))
                        Else
                            intCount = intCount + 1
                            strImage = Trim(Split(strLine, ",")(2))
                            strImage = intCount & ";110;0;" & Trim(Mid(strImage, 1, Len(strImage) - 1))
                            strImages = strImages & "^" & strImage
                        End If
                    End If
                End Select

                If InStr(strLine, "ItemName=UI_ReportDate") > 0 Then
                    StrTime = Trim(Split(strLine, ",")(3))
                    StrTime = Mid(StrTime, 1, Len(StrTime) - 1)
                    StrTime = Format(StrTime, "yyyy-mm-dd HH:MM:SS")
                ElseIf InStr(strLine, "ItemName=SampleID") > 0 Then
                    str�걾�� = Trim(Split(strLine, ",")(3))
                    str�걾�� = Trim(Mid(str�걾��, 1, Len(str�걾��) - 1))
                ElseIf InStr(strLine, "Dry Chemical Data Start") > 0 Then
                    intStep = 1
                End If
                
                If intStep = 4 Then
                    If aResults(UBound(aResults)) <> "" Then
                        ReDim Preserve aResults(UBound(aResults) + 1)
                    End If
                    aResults(UBound(aResults)) = StrTime & "|" & str�걾�� & "|ѪҺ| |0" & strItems & vbCrLf & strImages
                End If
                
                DoEvents
            Loop
            Close #1    ' �ر��ļ���
            strFile = Dir
        Loop

    Next

    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "EH-2050B Plus", strFile, Err.Number & " " & Err.Description
End Function








