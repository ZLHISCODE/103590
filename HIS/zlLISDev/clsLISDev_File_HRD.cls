VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_HRD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:Ѫ��������
'����:���������
'����:������
'�û�:��Ȩ������ҽԺ
'ʱ��:2013-04-7
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
                               ByVal dtStart As Date, Optional dtEnd As Date = CDate("3000-12-31")) As String()
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
    ' Dim DBConn As New ADODB.Connection
    '��SQL���ݿ�
    ' DBConn.Open "Provider   =   sqloledb;Data   Source= '�������� ';
    'Initial   Catalog= '���ݿ����� ';User   Id= '�û��� ';Password= '���� ' "
    '�粻�뽫�����ַ���д�ڳ�����Ļ���������ͨ����INI�ļ�������������������:
    '���鼼ʦ����վָ���ļ�ʱֻ��ָ�����ini�ļ�����
    '[LX2000]
    'Database Name = ExamDir
    'SERVER Name = (local)
    'USER Name = sa
    'PASSWORD=

    Dim aResults() As String
    Dim ConnSQL As New ADODB.Connection                                 '���ݿ�����
    Dim rsSQL As ADODB.Recordset                                        '���صļ�¼��
    Dim strSQL As String                                                'ִ�е�SQL���
    Dim strDate As String                                               '�ҳ���ʱ��
    Dim strNo As String                                                 '�걾��
    Dim strItems As String                                              'ͨ���뼰���
    Dim lngResultCount As Long, strid As String
    Dim strLog As String                                                '��¼����־
    
    Dim strIniFile As String, strImageFileName As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream
    Dim strBH As String, intloop As Integer                                                 '�ж��Ƿ����µı걾��
    Dim strPicPath As String, strPicFix As String                       'ͼƬ�ļ�
    Dim intDebug As Integer

    Dim DataBaseName As String
    Dim serverName As String
    Dim userName As String
    Dim passWord As String
    Dim objFile As New Scripting.FileSystemObject                       '�ж��Ƿ���ͼ���ļ�
    Dim strImgBmp As String, strTmp As String                           '����ͼ������
    Dim strYear As String, strYearMonth As String, strYearMonthDay As String
    Dim strImg As String
    Dim varItem As Variant, strItem As String
    On Err GoTo errH
    strIniFile = App.Path & "\HRD.ini"

    '------------------------------------��ȡ�����ļ�--------------------
    If Not objFileSystem.FileExists(strFile) Then
        Set objStream = objFileSystem.CreateTextFile(strIniFile)
        objStream.WriteLine "[Debug]"
        objStream.WriteLine "Text=0"    '1������ģʽ�������־ 0=����ģʽ���������־

        objStream.WriteLine "[Conn]"
        objStream.WriteLine "DatabaseName= ExamDir"
        objStream.WriteLine "serverName=(local)"
        objStream.WriteLine "userName=sa"
        objStream.WriteLine "passWord="

'        objStream.WriteLine "[Options]"
'        objStream.WriteLine "PicPath=D:\Program Files\US-2020A\US-2020A\pic"
'        objStream.WriteLine "PicFix=U100N"
'        objStream.Close
'        Set objStream = Nothing
'        strPicPath = "D:\Program Files\US-2020A\US-2020A\pic"
        intDebug = 0
        DataBaseName = "ExamDir"  '"th_us2020(1.5)"
        serverName = "(local)" ' "YF_WZT\SQLEXPRESS"
        userName = "sa"
        passWord = ""

    Else
'        strPicPath = ReadIni("Options", "PicPath", strIniFile)
'        strPicFix = ReadIni("Options", "PicFix", strIniFile)
'        If (Len(strPicPath) > 0) Then
'            If Mid(strPicPath, Len(strPicPath)) <> "\" Then strPicPath = strPicPath + "\"
'        End If
        intDebug = ReadIni("Debug", "Text", strIniFile, 0)
        DataBaseName = ReadIni("Conn", "DatabaseName", strIniFile, "ExamDir")
        serverName = ReadIni("Conn", "serverName", strIniFile, "(local)")  '"YF_WZT\SQLEXPRESS")
        userName = ReadIni("Conn", "userName", strIniFile, "sa")
        passWord = ReadIni("Conn", "passWord", strIniFile, "")
    End If
    '----------------------------------END-------------------------------
    With ConnSQL
        .ConnectionString = "Provider=SQLOLEDB;Data Source=" & serverName & ";user id=" & userName & ";password=" & passWord & ";initial catalog=" & DataBaseName
        .Open
        If .State = adStateClosed Then
            If intDebug = 1 Then
                strLog = "����SQL SERVER���ݿ�ʧ��"
                Call writelislog(strLog)
            End If
            Exit Function
        End If

        If intDebug = 1 Then
            strLog = "����SQL SERVER���ݿ�ɹ�"
            Call writelislog(strLog)
        End If
        strSQL = "select a.person_id,a.person_exam,a.id0,a.id1,a.id2,a.id3,a.id4,a.id5,a.id6,a.id7,a.id8,a.id9,a.id10," & vbNewLine & _
                 "a.id11,a.id12,a.id13,a.id14,a.id15,a.id16,a.id17,a.id18,a.id19,a.id20,a.id21,a.id22," & vbNewLine & _
                 "a.id0_val,a.id1_val,a.id2_val,a.id3_val,a.id4_val,a.id5_val,a.id6_val,a.id7_val,a.id8_val,a.id9_val,a.id10_val,a.id11_val," & vbNewLine & _
                 "a.id12_val,a.id13_val,a.id14_val,a.id15_val,a.id16_val,a.id17_val,a.id18_val,a.id18_val,a.id19_val,a.id30_val,a.id21_val,a.id22_val," & vbNewLine & _
                 "a.check_date from persons as a where " & vbNewLine & _
                 "a.check_date>= '" & Format(dtStart, "YYYY-MM-DD") & "' " & _
                 " and a.check_date<= '" & Format(dtEnd, "YYYY-MM-DD 23:59:59") & "'" & _
                 IIf(lngSampleNo = -1, " ", "a.person_exam =" & lngSampleNo & "and") & _
                 "order by a.person_id"
                    
        Set rsSQL = ConnSQL.Execute(strSQL)

        If rsSQL.EOF Then
            .Close
            ReDim Preserve aResults(0)
            aResults(0) = ""
            ResultFromFile = aResults
            Exit Function
        End If

        rsSQL.MoveFirst
        lngResultCount = -1
        strBH = "***"
        Do While Not rsSQL.EOF
            If strBH <> "" & rsSQL("person_exam") Then  '������
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                aResults(lngResultCount) = Format("" & rsSQL("check_date"), "YYYY-MM-DD hh:mm:ss")
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                           Val("" & rsSQL("person_exam"))                                  '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                         " "  '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                           "ѪҺ"  '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|0"    '& _

'                If strItems <> "" And strTmp <> "" Then
'                    aResults(lngResultCount - 1) = aResults(lngResultCount - 1) & vbCrLf & Mid(strTmp, 2)
'                    strTmp = ""
'                End If
'                If strPicPath <> "" Then
'                  '��ȡ����ѡ��ͼƬ
'                    strImgBmp = ""
'                    strTmp = ""
'                    strYear = Format(rsSQL("sjrq"), "yyyy")
'                    strYearMonth = Val(Format(rsSQL("sjrq"), "MM"))
'                    strYearMonthDay = Val(Format(rsSQL("sjrq"), "dd"))
'                    strImgBmp = strPicPath & strYear & "\" & strYearMonth & "\" & strYearMonthDay & "\pic\"
'                    strImageFileName = Dir(strImgBmp & "\*.jpg", vbHidden + vbDirectory + vbNormal + vbReadOnly + vbSystem + vbVolume)
'                    Do While strImageFileName <> ""
'                        strid = Mid(strImageFileName, 1, InStr(strImageFileName, "_") - 1)
'                        If strid = rsSQL("id") Then
'                            strTmp = strTmp & "^" & strImageFileName & ";113;0;" & strImgBmp & strImageFileName
'                            If intDebug = 1 Then
'                                strLog = "�ҵ�ѡ��ͼƬ: " & strImgBmp & strImageFileName
'                                Call writelislog(strLog)
'                            End If
'                        End If
'                          strImg = rsSql("ID") & "_00" & intloop & ".jpg"
'
'                          If Dir(strImgBmp & strImg, vbHidden) Then
'                              strTmp = strTmp & "^" & CStr(intloop) & ";113;0;" & strImgBmp & strImg
'                              If intDebug = 1 Then
'                                  strLog = "�ҵ�ѡ��ͼƬ: " & strImgBmp & strImg
'                                  Call writelislog(strLog)
'                              End If
'                          End If
'                        strImageFileName = Dir
'                    Loop
                    
                    
                If intDebug = 1 Then
                    strLog = "��ǰ����: " & aResults(lngResultCount)
                    Call writelislog(strLog)
                End If
                
                strBH = rsSQL("person_exam")
            
            End If
                
            If IsNumeric(strBH) Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                        rsSQL("id0") & "|" & IIf(InStr(rsSQL("id0_val"), ".") = 1, Format(rsSQL("id0_val"), "0.00"), rsSQL("id0_val")) & "|" & rsSQL("id1") & "|" & IIf(InStr(rsSQL("id1_val"), ".") = 1, Format(rsSQL("id1_val"), "0.00"), rsSQL("id1_val")) & "|" & rsSQL("id2") & "|" & IIf(InStr(rsSQL("id2_val"), ".") = 1, Format(rsSQL("id2_val"), "0.00"), rsSQL("id2_val")) & "|" & rsSQL("id3") & "|" & IIf(InStr(rsSQL("id3_val"), ".") = 1, Format(rsSQL("id3_val"), "0.00"), rsSQL("id3_val")) & _
                        "|" & rsSQL("id4") & "|" & IIf(InStr(rsSQL("id4_val"), ".") = 1, Format(rsSQL("id4_val"), "0.00"), rsSQL("id4_val")) & "|" & rsSQL("id5") & "|" & IIf(InStr(rsSQL("id5_val"), ".") = 1, Format(rsSQL("id5_val"), "0.00"), rsSQL("id5_val")) & "|" & rsSQL("id6") & "|" & IIf(InStr(rsSQL("id6_val"), ".") = 1, Format(rsSQL("id6_val"), "0.00"), rsSQL("id6_val")) & "|" & rsSQL("id7") & "|" & IIf(InStr(rsSQL("id7_val"), ".") = 1, Format(rsSQL("id7_val"), "0.00"), rsSQL("id7_val")) & _
                        "|" & rsSQL("id8") & "|" & IIf(InStr(rsSQL("id8_val"), ".") = 1, Format(rsSQL("id8_val"), "0.00"), rsSQL("id8_val")) & "|" & rsSQL("id9") & "|" & IIf(InStr(rsSQL("id9_val"), ".") = 1, Format(rsSQL("id9_val"), "0.00"), rsSQL("id9_val")) & "|" & rsSQL("id10") & "|" & IIf(InStr(rsSQL("id10_val"), ".") = 1, Format(rsSQL("id10_val"), "0.00"), rsSQL("id10_val")) & "|" & rsSQL("id11") & "|" & IIf(InStr(rsSQL("id11_val"), ".") = 1, Format(rsSQL("id11_val"), "0.00"), rsSQL("id11_val")) & _
                        "|" & rsSQL("id12") & "|" & IIf(InStr(rsSQL("id12_val"), ".") = 1, Format(rsSQL("id12_val"), "0.00"), rsSQL("id12_val")) & "|" & rsSQL("id13") & "|" & IIf(InStr(rsSQL("id13_val"), ".") = 1, Format(rsSQL("id13_val"), "0.00"), rsSQL("id13_val")) & "|" & rsSQL("id14") & "|" & IIf(InStr(rsSQL("id14_val"), ".") = 1, Format(rsSQL("id14_val"), "0.00"), rsSQL("id14_val")) & "|" & rsSQL("id15") & "|" & IIf(InStr(rsSQL("id15_val"), ".") = 1, Format(rsSQL("id15_val"), "0.00"), rsSQL("id15_val")) & _
                        "|" & rsSQL("id16") & "|" & IIf(InStr(rsSQL("id16_val"), ".") = 1, Format(rsSQL("id16_val"), "0.00"), rsSQL("id16_val")) & "|" & rsSQL("id17") & "|" & IIf(InStr(rsSQL("id17_val"), ".") = 1, Format(rsSQL("id17_val"), "0.00"), rsSQL("id17_val")) & "|" & rsSQL("id8") & "|" & IIf(InStr(rsSQL("id18_val"), ".") = 1, Format(rsSQL("id18_val"), "0.00"), rsSQL("id18_val")) & "|" & rsSQL("id19") & "|" & IIf(InStr(rsSQL("id19_val"), ".") = 1, Format(rsSQL("id19_val"), "0.00"), rsSQL("id19_val")) & _
                        "|" & rsSQL("id20") & "|" & IIf(InStr(rsSQL("id30_val"), ".") = 1, Format(rsSQL("id30_val"), "0.00"), rsSQL("id30_val")) & "|" & rsSQL("id21") & "|" & IIf(InStr(rsSQL("id21_val"), ".") = 1, Format(rsSQL("id21_val"), "0.00"), rsSQL("id21_val")) & "|" & rsSQL("id22") & "|" & IIf(InStr(rsSQL("id22_val"), ".") = 1, Format(rsSQL("id22_val"), "0.00"), rsSQL("id22_val"))
            End If
            If intDebug = 1 Then
                strLog = "��ǰ����: " & aResults(lngResultCount)
                Call writelislog(strLog)
            End If
            rsSQL.MoveNext
        Loop
    End With
    If rsSQL.EOF Then
        ConnSQL.Close
    End If
    ResultFromFile = aResults
    Exit Function
errH:
    If intDebug = 1 Then
        Call writelislog("�����(" & Err.Number & "):(" & Err.Description & ")")
    End If
End Function


Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_hrd"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As TextStream
    Dim strFileName As String
    strFileName = strFile & Format(date, "yyyy-mm-dd") & ".txt"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & strLog)
    objStream.Close
    Set objStream = Nothing
End Function
 

