VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AB7300REALPCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:AB7300REALPCR
'����:���
'�û�:����������ҽԺ

Private Type T�걾
    T�걾�� As String
    T��Ŀ�� As Integer
    T��� As String
    T����() As String
    T��Ŀ() As String
End Type

Private M�걾() As T�걾

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
    Dim strChenel As String, str����Ա As String, str���� As String
    Dim i As Integer, str��� As String, j As Integer, k As Integer
    Dim str�걾�� As String, str��� As String, blnReturn As Boolean
    Dim lngResultCount As Long, str��ʽ1 As String, str��ʽ2 As String
    Dim lngRtCount As Long, strFilePath As String
    Dim myArray() As String, strFilename As String
    Dim MyColor(2) As Long
    Dim objStream As textStream, objFile As New Scripting.FileSystemObject
    Dim strLine As String
    Dim objForm   As New frmGraph
    Dim varline As Variant
    
    

    On Error GoTo errHandle
    '��������±����
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    str���� = Format(Now, "yyyy-MM-dd hh:mm:ss")

    MyColor(0) = vbRed
    MyColor(1) = vbYellow
    MyColor(2) = vbBlue

    str��ʽ1 = Format(Now, "yyyymmdd")
    str��ʽ2 = Format(Now, "yyyy-mm-dd")


    'ͼ������
    Set objStream = objFile.OpenTextFile(strFilePath & "\" & str��ʽ1 & ".csv", ForReading)
    lngResultCount = 0
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If strLine <> "" Then
            varline = Split(strLine, ",")
            If varline(0) <> "Well" Then
                    ReDim Preserve myArray(lngResultCount)
                    myArray(lngResultCount) = varline(0)
                For i = 3 To UBound(varline)
                    myArray(lngResultCount) = myArray(lngResultCount) & "," & Round(varline(i) / 1000, 0)
                Next i
                lngResultCount = lngResultCount + 1
            End If
        End If
    Loop

    '�걾����
    Set objStream = objFile.OpenTextFile(strFilePath & "\" & str��ʽ2 & ".csv", ForReading)
    lngRtCount = 0
    Do Until objStream.AtEndOfStream
            strLine = objStream.ReadLine
            If j > 29 Then
            varline = Split(strLine, ",")
            If varline(1) <> "" Then
            If CInt(Val(Mid(varline(1), 1, 3))) > 0 Then
                
                str��� = (Asc(Mid(varline(0), 1, 1)) - 65) * 12 + Mid(varline(0), 2)
                str�걾�� = CInt(Val(Mid(varline(1), 1, 3)))
                str��� = Mid(varline(1), 4) & "|" & varline(4)
                For i = 0 To lngRtCount - 1
                    If M�걾(i).T�걾�� = str�걾�� Then
                        M�걾(i).T�걾�� = str�걾��
                        M�걾(i).T��� = M�걾(i).T��� & str���
                        M�걾(i).T��Ŀ�� = Val(M�걾(i).T��Ŀ��) + 1
                        '�걾ͼ������
                        For k = 0 To lngResultCount - 1
                            If Trim(str���) = Trim(Split(myArray(k), ",")(0)) Then
                                ReDim Preserve M�걾(i).T��Ŀ(M�걾(i).T��Ŀ��)
                                ReDim Preserve M�걾(i).T����(M�걾(i).T��Ŀ��)
                                M�걾(i).T��Ŀ(M�걾(i).T��Ŀ��) = Mid(varline(1), 4)
                                M�걾(i).T����(M�걾(i).T��Ŀ��) = Mid(myArray(k), InStr(1, myArray(k), ",") + 1)
                            End If
                        Next k
                        Exit For
                    End If
                Next i
                
                If i >= lngRtCount Then
                    ReDim Preserve M�걾(lngRtCount)
                    M�걾(lngRtCount).T�걾�� = str�걾��
                    M�걾(lngRtCount).T��� = str���
                    M�걾(lngRtCount).T��Ŀ�� = 0
                    '�걾ͼ������
                    For k = 0 To lngResultCount - 1
                        If Trim(str���) = Trim(Split(myArray(k), ",")(0)) Then
                            ReDim Preserve M�걾(lngRtCount).T����(M�걾(lngRtCount).T��Ŀ��)
                            ReDim Preserve M�걾(lngRtCount).T��Ŀ(M�걾(lngRtCount).T��Ŀ��)
                            M�걾(lngRtCount).T��Ŀ(M�걾(lngRtCount).T��Ŀ��) = Mid(varline(1), 4)
                            M�걾(lngRtCount).T����(M�걾(lngRtCount).T��Ŀ��) = Mid(myArray(k), InStr(1, myArray(k), ",") + 1)
                        End If
                    Next k
                    lngRtCount = lngRtCount + 1
                End If
            End If
            End If
        End If
        j = j + 1
    Loop
    
    For i = 0 To lngRtCount - 1
        ReDim Preserve aResults(i)
        aResults(i) = str����
        aResults(i) = aResults(i) & "|" & M�걾(i).T�걾��                '�������
        aResults(i) = aResults(i) & "|" & " "                             '������
        aResults(i) = aResults(i) & "|" & "ѪҺ"                          '�걾
        aResults(i) = aResults(i) & "|0|"                                  '�ʿ�
        aResults(i) = aResults(i) & M�걾(i).T���                        '��Ŀ�ͽ��
        '�Ȼ���
        objForm.Draw
        '�ٻ�ͼ
        objForm.DrawGraph M�걾(i).T����, MyColor, M�걾(i).T��Ŀ
        strFilename = strFilePath & "\" & M�걾(i).T�걾�� & ".bmp"
        '����ͼ��
        SavePicture objForm.Picture1.Image, strFilename
        aResults(i) = aResults(i) & vbCrLf & " ;103;1;" & strFilename
        '���ͼƬ�ؼ��е�ͼ��
        objForm.Picture1.Cls
        'frmGraphӫ��.Draw
    Next i
    '�ͷ���Դռ��
    Unload objForm
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_AB7300REALPCR.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function







