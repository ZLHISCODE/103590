VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_blotray886"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:�׶��Զ�����ӡ����
'����:�׶�
'�û�:��������ҽԺ
'���ߣ�������

'�ļ���Ϊ*.txt,�ļ�����ù̶���ͨ�������ļ�����ȡ�����
'�׶��Զ�����ӡ��������
'id , Hospital, HospitalDepartment, TestName, PatientNumber, Bed, PatientName, Sex, Age, PatientAddress, PatientTelephone, SourceTxt, SampleCode, SampleType, DataCollectd, DataReceived, DataTested, Doctor, Checker, ClinicalSymptom, SenderDepartment, SenderDoctor, zuhenumber, pic, T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16, Name1, Name2, Name3, Name4, Name5, Name6, Name7, Name8, Name9, Name10, Name11, Name12, Name13, Name14, Name15, Name16, U1, U2, U3, U4, U5, U6, U7, U8, U9, U10, U11, U12, U13, U14, U15, U16
'10,����������ҽԺ,JYK,,,,,,,,,,1,,,2016-05-17,2016-05-17,,,,,,�ۺ���14��,,,,,0,0,1,0,0,0,0,0,0,0,0,0,0,�����ʿ�,�����ʿ�,��lgE,�ݳ���/�۳���,�ݳ�,����/����/��ͩ/����/����,�బ/����/���,����ù/��֦��ù/����ù/������ù/��ù/ëù,èëƤм/��ëƤм,����/����,ţ��,����/����/����,Ϻ/з/�ȱ�,ţ��/����,â��/����/ƻ��/����/��ݮ,����/���Ĺ�/����/���,,,70.31,0.31,0.27,0.48,0.03,0.07,0.17,0.14,0.09,0.06,0.08,0.14,0.12,0.02
'11,����������ҽԺ,JYK,,,,,,,,,,2,,,2016-05-17,2016-05-17,,,,,,�ۺ���14��,,,,,0,0,0,0,2,1,0,0,0,0,0,0,0,�����ʿ�,�����ʿ�,��lgE,�ݳ���/�۳���,�ݳ�,����/����/��ͩ/����/����,�బ/����/���,����ù/��֦��ù/����ù/������ù/��ù/ëù,èëƤм/��ëƤм,����/����,ţ��,����/����/����,Ϻ/з/�ȱ�,ţ��/����,â��/����/ƻ��/����/��ݮ,����/���Ĺ�/����/���,,,18.29,0.16,0.13,0.17,0.02,0.99,0.39,0.03,0.13,0.04,0.01,0.23,0.07,0.04
'12,����������ҽԺ,JYK,,,,,,,,,,3,,,2016-05-17,2016-05-17,,,,,,�ۺ���14��,,,,,0,0,0,0,0,1,0,0,0,0,0,0,0,�����ʿ�,�����ʿ�,��lgE,�ݳ���/�۳���,�ݳ�,����/����/��ͩ/����/����,�బ/����/���,����ù/��֦��ù/����ù/������ù/��ù/ëù,èëƤм/��ëƤм,����/����,ţ��,����/����/����,Ϻ/з/�ȱ�,ţ��/����,â��/����/ƻ��/����/��ݮ,����/���Ĺ�/����/���,,,101.66,0.31,0.08,0.12,0.03,0.04,0.38,0.09,0.23,0.14,0.14,0.26,0.02,0.12
'13,����������ҽԺ,JYK,,,,,,,,,,4,,,2016-05-17,2016-05-17,,,,,,�ۺ���14��,,,,,0,0,0,0,0,0,0,0,0,0,0,0,0,�����ʿ�,�����ʿ�,��lgE,�ݳ���/�۳���,�ݳ�,����/����/��ͩ/����/����,�బ/����/���,����ù/��֦��ù/����ù/������ù/��ù/ëù,èëƤм/��ëƤм,����/����,ţ��,����/����/����,Ϻ/з/�ȱ�,ţ��/����,â��/����/ƻ��/����/��ݮ,����/���Ĺ�/����/���,,,74.33,0.12,0.07,0.06,0.06,0.07,0.02,0.12,0.12,0.06,0.06,0.02,0.08,0.04
'14,����������ҽԺ,JYK,,,,,,,,,,1,,,2016-05-17,2016-05-17,,,,,,�ۺ���14��,,,,,0,0,1,0,0,0,0,0,0,0,0,0,0,�����ʿ�,�����ʿ�,��lgE,�ݳ���/�۳���,�ݳ�,����/����/��ͩ/����/����,�బ/����/���,����ù/��֦��ù/����ù/������ù/��ù/ëù,èëƤм/��ëƤм,����/����,ţ��,����/����/����,Ϻ/з/�ȱ�,ţ��/����,â��/����/ƻ��/����/��ݮ,����/���Ĺ�/����/���,,,73.80,0.33,0.30,0.50,0.12,0.12,0.19,0.15,0.13,0.06,0.12,0.16,0.02,0.06
'15,����������ҽԺ,JYK,,,,,,,,,,2,,,2016-05-17,2016-05-17,,,,,,�ۺ���14��,,,,,0,0,0,0,2,1,0,0,0,0,0,0,0,�����ʿ�,�����ʿ�,��lgE,�ݳ���/�۳���,�ݳ�,����/����/��ͩ/����/����,�బ/����/���,����ù/��֦��ù/����ù/������ù/��ù/ëù,èëƤм/��ëƤм,����/����,ţ��,����/����/����,Ϻ/з/�ȱ�,ţ��/����,â��/����/ƻ��/����/��ݮ,����/���Ĺ�/����/���,,,16.30,0.14,0.02,0.15,0.11,0.93,0.39,0.06,0.09,0.02,0.03,0.22,0.08,0.12
'16,����������ҽԺ,JYK,,,,,,,,,,3,,,2016-05-17,2016-05-17,,,,,,�ۺ���14��,,,,,1,0,0,0,0,1,0,0,0,0,0,0,0,�����ʿ�,�����ʿ�,��lgE,�ݳ���/�۳���,�ݳ�,����/����/��ͩ/����/����,�బ/����/���,����ù/��֦��ù/����ù/������ù/��ù/ëù,èëƤм/��ëƤм,����/����,ţ��,����/����/����,Ϻ/з/�ȱ�,ţ��/����,â��/����/ƻ��/����/��ݮ,����/���Ĺ�/����/���,,,103.20,0.36,0.03,0.02,0.05,0.05,0.41,0.08,0.27,0.14,0.15,0.29,0.10,0.04
'17,����������ҽԺ,JYK,,,,,,,,,,4,,,2016-05-17,2016-05-17,,,,,,�ۺ���14��,,,,,0,0,0,0,0,0,0,0,0,0,0,0,0,�����ʿ�,�����ʿ�,��lgE,�ݳ���/�۳���,�ݳ�,����/����/��ͩ/����/����,�బ/����/���,����ù/��֦��ù/����ù/������ù/��ù/ëù,èëƤм/��ëƤм,����/����,ţ��,����/����/����,Ϻ/з/�ȱ�,ţ��/����,â��/����/ƻ��/����/��ݮ,����/���Ĺ�/����/���,,,74.71,0.07,0.07,0.04,0.13,0.03,0.05,0.02,0.08,0.03,0.06,0.11,0.16,0.04

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
    Dim strChenel As String, str����Ա As String, str���� As String
    Dim strFilePath As String, strTmp As String

    Dim objStream As textStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str�걾�� As String, str������ As String, blnReturn As Boolean
    Dim lngRow As Long, varLine As Variant

    Dim intCol_U3 As Integer
    Dim intCol_U4 As Integer
    Dim intCol_U5 As Integer
    Dim intCol_U6 As Integer
    Dim intCol_U7 As Integer
    Dim intCol_U8 As Integer
    Dim intCol_U9 As Integer
    Dim intCol_U10 As Integer
    Dim intCol_U11 As Integer
    Dim intCol_U12 As Integer
    Dim intCol_U13 As Integer
    Dim intCol_U14 As Integer
    Dim intCol_U15 As Integer
    Dim intCol_U16 As Integer

    '���ڣ��걾��
    Dim intCol_No As Integer
    Dim intDateS As Integer, intDateLen As Integer
    Dim intCol_Date As Integer

    On Error GoTo errHandle
        
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    '---��ȡ����
    intCol_No = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "�걾����", "12"))
    intCol_Date = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "������", "16"))
    intDateS = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "���ڿ�ʼλ��", "1"))
    intDateLen = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "���ڳ���", "10"))
    intCol_U3 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "U3", "58"))
    intCol_U4 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "U4", "59"))
    intCol_U5 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "U5", "60"))
    intCol_U6 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "U6", "61"))
    intCol_U7 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "U7", "62"))
    intCol_U8 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "U8", "63"))
    intCol_U9 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "U9", "64"))
    intCol_U10 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "U10", "65"))
    intCol_U11 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "U11", "66"))
    intCol_U12 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "U12", "67"))
    intCol_U13 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "U13", "68"))
    intCol_U14 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "U14", "69"))
    intCol_U15 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "U15", "70"))
    intCol_U16 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "����", "U16", "71"))
    
    '-----
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    Set objStream = objFile.OpenTextFile(strFilePath & "\all.csv", ForReading)
    
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If InStr(strLine, ",") > 0 Then
            varLine = Split(strLine, ",")
            If lngRow = 0 Then
                str���� = Format(Now, "yyyy-MM-dd")
            Else
                str������ = ""
                str���� = Trim(varLine(intCol_Date))
                If Len(str����) > intDateS + intDateLen Then
                    str���� = Mid(str����, intDateS, intDateLen)
                End If
                If Not IsDate(str����) Then str���� = Format(Now, "yyyy-MM-dd")
        
                str�걾�� = Trim(varLine(intCol_No))
                If Val(str�걾��) > 0 Then
                   
                    strTmp = Trim$(varLine(intCol_U3))
                    If strTmp <> "" Then str������ = str������ & "|U3|" & strTmp
                    
                       strTmp = Trim$(varLine(intCol_U4))
                    If strTmp <> "" Then str������ = str������ & "|U4|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U5))
                    If strTmp <> "" Then str������ = str������ & "|U5|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U6))
                    If strTmp <> "" Then str������ = str������ & "|U6|" & strTmp
                       
                    strTmp = Trim$(varLine(intCol_U7))
                    If strTmp <> "" Then str������ = str������ & "|U7|" & strTmp
                       
                    strTmp = Trim$(varLine(intCol_U8))
                    If strTmp <> "" Then str������ = str������ & "|U8|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U9))
                    If strTmp <> "" Then str������ = str������ & "|U9|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U10))
                    If strTmp <> "" Then str������ = str������ & "|U10|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U11))
                    If strTmp <> "" Then str������ = str������ & "|U11|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U12))
                    If strTmp <> "" Then str������ = str������ & "|U12|" & strTmp
                       
                    strTmp = Trim$(varLine(intCol_U13))
                    If strTmp <> "" Then str������ = str������ & "|U13|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U14))
                    If strTmp <> "" Then str������ = str������ & "|U14|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U15))
                    If strTmp <> "" Then str������ = str������ & "|U15|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U16))
                    If strTmp <> "" Then str������ = str������ & "|U16|" & strTmp
                    
                    If str������ <> "" Then
                        If aResults(UBound(aResults)) <> "" Then
                            ReDim Preserve aResults(UBound(aResults) + 1)
                        End If
                        aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| | |0" & str������
                    End If
               End If
            End If
        End If
        lngRow = lngRow + 1
        DoEvents
    Loop
    ResultFromFile = aResults
    
    Exit Function
errHandle:
'     WriteLog " BH5100," & Erl() & Err.Description
End Function

Private Function ChangeValue(ByVal strInValue As String) As String
    'ת�����ֵ������Ϊ
    'ֵ>0�������Ч����ԭֵ���أ�������<=0,��ʾΪ'δ���'
    
    If Trim(strInValue) = "" Then
        ChangeValue = ""
    ElseIf Val(strInValue) > 0 Then
        ChangeValue = Trim(strInValue)
    Else
        ChangeValue = "δ���"
    End If
    
End Function


