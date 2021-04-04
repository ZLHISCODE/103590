VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_US_2020A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:US-2020A �������������
'����:�����캣ҽ���豸���޹�˾
'����:������
'�û�:�ӽ�����ְ��ҽԺ
'ʱ��:2012-10-18
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
    Dim rsSql As ADODB.Recordset                                        '���صļ�¼��
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
    strIniFile = App.Path & "\US_2020A.ini"

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

        objStream.WriteLine "[Options]"
        objStream.WriteLine "PicPath=D:\Program Files\US-2020A\US-2020A\pic"
        objStream.WriteLine "PicFix=U100N"
        objStream.Close
        Set objStream = Nothing
        strPicPath = "D:\Program Files\US-2020A\US-2020A\pic"
        intDebug = 0
        DataBaseName = "ExamDir"  '"th_us2020(1.5)"
        serverName = "(local)" ' "YF_WZT\SQLEXPRESS"
        userName = "sa"
        passWord = ""

    Else
        strPicPath = ReadIni("Options", "PicPath", strIniFile)
        strPicFix = ReadIni("Options", "PicFix", strIniFile)
        If (Len(strPicPath) > 0) Then
            If Mid(strPicPath, Len(strPicPath)) <> "\" Then strPicPath = strPicPath + "\"
        End If
        intDebug = ReadIni("Debug", "Text", strIniFile, 0)
        DataBaseName = ReadIni("Conn", "DatabaseName", strIniFile, "ExamDir")    ' "th_us2020(1.5)")
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
    
        strSQL = "select a.id, b.sxh as sxh, a.ifprint,a.caption as caption ,a.val as val ,b.name as name ,b.sjsj as sjsj " & _
                    ",b.sjrq as sjrq from dbo.jz_item as a,dbo.jzxx as b where " & vbNewLine & _
                    " b.rq>= '" & Format(dtStart, "YYYY-MM-DD") & "' " & vbNewLine & _
                    " and b.sjrq<= '" & Format(dtEnd, "YYYY-MM-DD 23:59:59") & "' and  " & _
                    IIf(lngSampleNO = -1, " ", "b.sxh =" & lngSampleNO & "and") & _
                    " a.id= b.id and a.val is not null order by a.id  "
                    
        Set rsSql = ConnSQL.Execute(strSQL)

        If rsSql.EOF Then
            .Close
            ReDim Preserve aResults(0)
            aResults(0) = ""
            ResultFromFile = aResults
            Exit Function
        End If

        rsSql.MoveFirst
        lngResultCount = -1
        strBH = "***"
        Do While Not rsSql.EOF
            If strBH <> rsSql("ID") Then    '������
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                aResults(lngResultCount) = Format(rsSql("sjrq") & " " & rsSql("sjsj"), "YYYY-MM-DD hh:mm:ss")
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                           Val(rsSql("sxh"))                                    '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                         " "  '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                           "��Һ"  '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|0"    '& _

                If strItems <> "" And strTmp <> "" Then
                    aResults(lngResultCount - 1) = aResults(lngResultCount - 1) & vbCrLf & Mid(strTmp, 2)
                    strTmp = ""
                End If
                If strPicPath <> "" Then
                  '��ȡ����ѡ��ͼƬ
                    strImgBmp = ""
                    strTmp = ""
                    strYear = Format(rsSql("sjrq"), "yyyy")
                    strYearMonth = Val(Format(rsSql("sjrq"), "MM"))
                    strYearMonthDay = Val(Format(rsSql("sjrq"), "dd"))
                    strImgBmp = strPicPath & strYear & "\" & strYearMonth & "\" & strYearMonthDay & "\pic\"
                    strImageFileName = Dir(strImgBmp & "\*.jpg", vbHidden + vbDirectory + vbNormal + vbReadOnly + vbSystem + vbVolume)
                    Do While strImageFileName <> ""
                        strid = Mid(strImageFileName, 1, InStr(strImageFileName, "_") - 1)
                        If strid = rsSql("id") Then
                            strTmp = strTmp & "^" & strImageFileName & ";113;0;" & strImgBmp & strImageFileName
                            If intDebug = 1 Then
                                strLog = "�ҵ�ѡ��ͼƬ: " & strImgBmp & strImageFileName
                                Call writelislog(strLog)
                            End If
                        End If
'                          strImg = rsSql("ID") & "_00" & intloop & ".jpg"
'
'                          If Dir(strImgBmp & strImg, vbHidden) Then
'                              strTmp = strTmp & "^" & CStr(intloop) & ";113;0;" & strImgBmp & strImg
'                              If intDebug = 1 Then
'                                  strLog = "�ҵ�ѡ��ͼƬ: " & strImgBmp & strImg
'                                  Call writelislog(strLog)
'                              End If
'                          End If
                        strImageFileName = Dir
                    Loop
                    
                    
                    If intDebug = 1 Then
                        strLog = "��ǰ����: " & aResults(lngResultCount)
                        Call writelislog(strLog)
                    End If
                End If
                
                strBH = rsSql("id")
            
            End If

            If Trim(rsSql("caption")) <> "˳���" Then
                strItems = rsSql("caption")
                strItems = ReplaceItem(strItems, rsSql)                                         '������ͨ�����滻��Ӣ��
                
                If strItems <> "" Or IsNull(strItems) = False Then
                    aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                        Trim(strItems) & "|" & IIf(IsNull(rsSql("val")), " ", Split(rsSql("val"), " ")(0))
                End If
            End If
            If intDebug = 1 Then
                strLog = "��ǰ����: " & aResults(lngResultCount)
                Call writelislog(strLog)
            End If
            rsSql.MoveNext
        Loop
        If strTmp <> "" Then aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & Mid(strTmp, 2)
    End With
    If rsSql.EOF Then
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
    Const strFile As String = "C:\LisLog_US_2020A_"
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
 
 '������ͨ�����滻��Ӣ��
Private Function ReplaceItem(ByVal strItems As String, ByVal rsSql As Recordset) As String
        strItems = Replace(strItems, "������", "GLU")
        strItems = Replace(strItems, "������", "PRO")
        strItems = Replace(strItems, "������", "BIL")
        strItems = Replace(strItems, "��ԭ", "URO")
        strItems = Replace(strItems, "PHֵ", "PH")
        strItems = Replace(strItems, "����", "SG")
        strItems = Replace(strItems, "��Ѫ", "BLD")

        strItems = Replace(strItems, "ͪ��", "KET")
        strItems = Replace(strItems, "��������", "NIT")
        If InStr(rsSql("val"), "+") > 0 Or InStr(rsSql("val"), "-") > 0 Then
            strItems = Replace(strItems, "��ϸ��", "LUE")
        Else
            strItems = Replace(strItems, "��ϸ��", "BXB")
        End If
        strItems = Replace(strItems, "��ϸ��", "HXB")
        strItems = Replace(strItems, "��Ƥϸ��", "SPXB")

        strItems = Replace(strItems, "СԲ��Ƥ", "XYSP")
        strItems = Replace(strItems, "��ϸ������", "HXBGX")
        strItems = Replace(strItems, "��ϸ������", "BXBGX")
        strItems = Replace(strItems, "��Ƥϸ������", "SPXBGX")
        strItems = Replace(strItems, "��������", "CYGX")
        strItems = Replace(strItems, "֬������", "ZFGX")
        strItems = Replace(strItems, "��������", "KLGX")

        strItems = Replace(strItems, "��Ϲ���", "HHGX")
        strItems = Replace(strItems, "���ι���", "KXGX")
        strItems = Replace(strItems, "����ᾧ", "LSJJ")
        strItems = Replace(strItems, "����ƽᾧ", "CSGJJ")
        strItems = Replace(strItems, "������ᾧ", "LASJJ")
        strItems = Replace(strItems, "�����ؽᾧ", "DHSJJ")
        strItems = Replace(strItems, "���̴��ᾧ", "DGCJJ")

        strItems = Replace(strItems, "�ǰ��ᾧ", "HAJJ")
        strItems = Replace(strItems, "����ϸ��1", "QTXB1")
        strItems = Replace(strItems, "����ϸ��2", "QTXB2")
        strItems = Replace(strItems, "����ϸ��3", "QTXB3")
        strItems = Replace(strItems, "����ϸ��4", "QTXB4")

        strItems = Replace(strItems, "����", "GX")
        strItems = Replace(strItems, "����", "QT")
        strItems = Replace(strItems, "ù��", "MJ")
        strItems = Replace(strItems, "����", "JZ")
        If rsSql("ifprint") = True Then
            strItems = Replace(strItems, "�ᾧ", "JJZ")
        Else
            strItems = Replace(strItems, "�ᾧ", "JJ")
        End If
        strItems = Replace(strItems, "ϸ��", "XJ")
        ReplaceItem = strItems
End Function





