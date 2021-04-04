VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImageDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mblnIsLoadReportImage As Boolean '�Ǹ��ݱ���ͼ���ֶμ��صı���ͼ
Private mslQueryLevel As tQueryLevel      'ͼ����ʾ����
Private mtyFileLoadType As FileLoadType

'�ļ����ط�ʽ
Property Get ImgLoadType() As FileLoadType
    ImgLoadType = mtyFileLoadType
End Property

Property Let ImgLoadType(value As FileLoadType)
    mtyFileLoadType = value
End Property


'ͼ����ʾ����
Property Get QueryLevel() As tQueryLevel
    QueryLevel = mslQueryLevel
End Property

Property Let QueryLevel(value As tQueryLevel)
    mslQueryLevel = value
End Property

Public Function DownloadImages(ByRef arrImages() As String, strQueryValue As String, ByVal lngStartIndex As Long, ByVal lngEndIndex As Long, _
                blnOnlyLoadReportImage As Boolean, blnMoved As Boolean, _
                Optional blnQueryTmpRecord As Boolean) As Boolean
'lngAdviceId����ѯ����ֵ
'blnOnlyLoadReportImage��ΪTrueʱ���� ����ͼ�� �ֶ��еı���ͼ,��֮�������б���ͼ
'lngStartIndex����һ��ͼ������
'lngEndIndex�����һ��ͼ������
'blnMoved��ͼ���Ƿ�ת��
'slQueryLevel��ͼ�����ؼ���
'blnQueryTmpRecord���Ƿ���ʱͼ��
    Dim rsData As Recordset

    If strQueryValue = "0" Then Exit Function
    If QueryLevel = slAdvice And Not IsNumeric(Val(strQueryValue)) Then Exit Function

    If blnOnlyLoadReportImage Then
        '���� Ӱ�����¼.����ͼ�� �ֶ��е�ֵ���أ����Ϊ�գ� ���������б���ͼ��
        Set rsData = GetImageRptData(strQueryValue, lngStartIndex, lngEndIndex, blnMoved)
        
        mblnIsLoadReportImage = rsData.RecordCount > 0
        
        If rsData.RecordCount <= 0 Then
            Set rsData = GetImageViewData(mslQueryLevel, strQueryValue, lngStartIndex, lngEndIndex, blnQueryTmpRecord, blnMoved, blnOnlyLoadReportImage)
        End If
    Else
        Set rsData = GetImageViewData(mslQueryLevel, strQueryValue, lngStartIndex, lngEndIndex, blnQueryTmpRecord, blnMoved, blnOnlyLoadReportImage)
    End If
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    If mtyFileLoadType = Normal Then
        DownloadImages = LoadViewImageToFaceWithNormal(rsData, arrImages)     'ʹ��ԭʼģʽ����
    Else
        DownloadImages = LoadViewImageToFaceWithService(rsData, arrImages)   'ʹ��ZLPacsServerCenter����,��̨����
    End If

End Function

Private Function GetImageRptData(ByVal lngOrderID As Long, ByVal lngStartIndex As Long, ByVal lngEndIndex As Long, blnMoved As Boolean) As ADODB.Recordset
'���ݱ���ͼ�� �ֶλ�ȡ���ͼ��
    Dim strSQL As String

    strSQL = "Select rownum As ˳���,a.ҽ��id,a.����,a.�Ա�,a.����, rownum As ͼ���,Replace(Trim(D.Column_Value),'.jpg','') as ͼ��UID, A.���UID, " & _
            "'' As ����UID, 0 as ��̬ͼ,'' as ��������,'' as �ɼ�ʱ��, '' as ¼�Ƴ���, '' as ����ͼ," & _
            "B.FTP�û��� As User1,B.FTP���� As Pwd1,B.IP��ַ As Host1,'/'||B.FtpĿ¼||'/' As Root1, " & _
            "B.����Ŀ¼ as ����Ŀ¼1,B.����Ŀ¼�û��� as ����Ŀ¼�û���1,B.����Ŀ¼���� as ����Ŀ¼����1, " & _
            "Decode(A.��������,Null,'',to_Char(A.��������,'YYYYMMDD')||'/') ||A.���UID||'/'||Replace(Trim(D.Column_Value),'.jpg','') As URL,B.�豸�� as �豸��1, B.�豸�� as �豸��1, " & _
            "C.FTP�û��� As User2,C.FTP���� As Pwd2,C.IP��ַ As Host2,'/'||C.FtpĿ¼||'/' As Root2, " & _
            "C.����Ŀ¼ as ����Ŀ¼2,C.����Ŀ¼�û��� as ����Ŀ¼�û���2,C.����Ŀ¼���� as ����Ŀ¼����2,C.�豸�� as �豸��2, C.�豸�� as �豸��2 " & _
            "From Ӱ�����¼ A, Ӱ���豸Ŀ¼ B, Ӱ���豸Ŀ¼ C, Table(Cast(f_Str2list(A.����ͼ��,';') As zlTools.t_Strlist)) D " & _
            "Where A.λ��һ = B.�豸��(+) And A.λ�ö� = C.�豸��(+) And A.ҽ��id = [1]"

    If blnMoved = True Then strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")


    strSQL = "select * from (" & strSQL & " order by ����UID, ͼ���) where ˳���>=" & lngStartIndex & " and ˳���<=" & lngEndIndex

    Set GetImageRptData = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ͼ��", lngOrderID)
End Function

Private Function GetImageViewData(ByVal slQueryLevel As tQueryLevel, ByVal strSearchValue As String, _
    ByVal lngStartIndex As Long, ByVal lngEndIndex As Long, ByVal blnTmpRecord As Boolean, blnMoved As Boolean, blnOnlyLoadReportImage As Boolean) As ADODB.Recordset
'��ȡԤ��ͼ������

    Dim strSQL As String

    strSQL = "Select rownum as ˳���,[2] ҽ��id,c.����,c.�Ա�,c.����, A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1," & _
            "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1,D.����Ŀ¼ as ����Ŀ¼1,D.����Ŀ¼�û��� as ����Ŀ¼�û���1,D.����Ŀ¼���� as ����Ŀ¼����1," & _
            "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/') " & _
            "||C.���UID||'/'||A.ͼ��UID As URL,d.�豸�� as �豸��1, D.�豸�� As �豸��1," & _
            "E.FTP�û��� As User2,E.FTP���� As Pwd2,A.����ͼ," & _
            "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2,E.����Ŀ¼ as ����Ŀ¼2,E.����Ŀ¼�û��� as ����Ŀ¼�û���2,E.����Ŀ¼���� as ����Ŀ¼����2," & _
            "E.�豸�� as �豸��2, E.�豸�� As �豸��2, A.ͼ��UID,C.���UID,B.����UID,A.��̬ͼ,A.��������,A.�ɼ�ʱ��, A.¼�Ƴ��� " & _
            "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
            "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+)" & IIf(blnOnlyLoadReportImage, " And nvl(A.��̬ͼ,0) = 0 ", "")

    If blnTmpRecord Then
        strSQL = Replace(strSQL, "Ӱ����", "Ӱ����ʱ")
    Else
        If blnMoved Then
            strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
            strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
        End If
    End If

    Select Case slQueryLevel
        Case slAdvice
            strSQL = "select * from (" & strSQL & " and C.ҽ��ID=[1])"
        Case slStudy
            strSQL = "select * from (" & strSQL & " and C.���UID=[1])"
        Case slSeries
            strSQL = "select * from (" & strSQL & " and B.����UID=[1])"
        Case slImage
            strSQL = "select * from (" & strSQL & " and A.ͼ��UID=[1])"
    End Select

    strSQL = "select * from (" & strSQL & " order by ����UID, ͼ���) where ˳���>=" & lngStartIndex & " and ˳���<=" & lngEndIndex

    Set GetImageViewData = zlDatabase.OpenSQLRecord(strSQL, "��ѯͼ����Ϣ", strSearchValue, IIf(blnTmpRecord, "-1", strSearchValue))
End Function


Private Function LoadViewImageToFaceWithNormal(rsCurImageData As ADODB.Recordset, ByRef arrImages() As String) As Boolean
'����Ԥ��ͼ�񵽽���
    Dim strTmpFile As String
    Dim strCachePath As String

    Dim curImage As DicomImage

    Dim objFile As New Scripting.FileSystemObject

    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp

    Dim dcmTag As clsImageTagInf
    Dim strImgInstanceUid As String
    Dim strCurInstanceUids As String
    Dim blnIsAddImage As Boolean
    Dim objImgInfo As Object
    Dim strItem As String

BUGEX "LoadViewImageToFaceWithNormal 1"

    blnIsAddImage = False
    LoadViewImageToFaceWithNormal = False

BUGEX "LoadViewImageToFaceWithNormal 2"

    '��������ͼ�񻺴�Ŀ¼
    strCachePath = GetCacheDir
    MkLocalDir strCachePath & objFile.GetParentFolderName(NVL(rsCurImageData("URL")))
    
    ReDim arrImages(0)
BUGEX "LoadViewImageToFaceWithNormal 3"
    Do While Not rsCurImageData.EOF
        'ѭ������ͼ��DicomViewer��
        strImgInstanceUid = Trim(NVL(rsCurImageData!ͼ��UID))

        If InStr(strCurInstanceUids, strImgInstanceUid) <= 0 And strImgInstanceUid <> "" Then

            blnIsAddImage = True

            '��������Ƶ����ʾ�ļ������Ϊ����Ƶ�ļ�ʱ���ù��̽����ӷ�������ֱ�����������ļ�
            If NVL(rsCurImageData!��̬ͼ, imgTag) = VIDEOTAG Then
                strTmpFile = GetResourceDir & "Avi.bmp"
            ElseIf NVL(rsCurImageData!��̬ͼ, imgTag) = AUDIOTAG Then
                strTmpFile = GetResourceDir & "wav.bmp"
            Else
                strTmpFile = strCachePath & NVL(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")
            End If

            strTmpFile = Replace(Trim(strTmpFile), "/", "\")

            If Dir(strTmpFile) = vbNullString Then
                '���ػ���ͼ�񲻴��ڣ����ȡFTPͼ��
                '����FTP����
                If NVL(rsCurImageData("�豸��1")) <> vbNullString And Inet1.hConnection = 0 Then
                    If Inet1.FuncFtpConnect(NVL(rsCurImageData("Host1")), NVL(rsCurImageData("User1")), NVL(rsCurImageData("Pwd1"))) = 0 Then
                        If NVL(rsCurImageData("�豸��2")) <> vbNullString Then
                            If Inet2.FuncFtpConnect(NVL(rsCurImageData("Host2")), NVL(rsCurImageData("User2")), NVL(rsCurImageData("Pwd2"))) = 0 Then
                                MsgBox "FTP�����������ӣ������������á�", vbOKOnly, CON_STR_HINT_TITLE
                                Exit Function
                            End If
                        Else
                            MsgBox "FTP�����������ӣ������������á�", vbOKOnly, CON_STR_HINT_TITLE
                            Exit Function
                        End If
                    End If
                End If

                If Inet1.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsCurImageData("Root1")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")) <> 0 Then
                    '���豸��1��ȡͼ��ʧ�ܣ�����豸��2��ȡͼ��
                    If NVL(rsCurImageData("�豸��2")) <> vbNullString Then
                        If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect NVL(rsCurImageData("Host2")), NVL(rsCurImageData("User2")), NVL(rsCurImageData("Pwd2"))
                        Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsCurImageData("Root2")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", ""))
                    End If
                End If
            End If
            
            
            
            strItem = strTmpFile & "|" & NVL(rsCurImageData!��̬ͼ, imgTag) & "|" & NVL(rsCurImageData!����ͼ) & "|" & _
                NVL(rsCurImageData("��������"), "") & "|" & NVL(rsCurImageData("�ɼ�ʱ��")) & "|" & strCachePath & NVL(rsCurImageData("URL")) & ".avi" & "|" & _
                Val(NVL(rsCurImageData("¼�Ƴ���"), "0")) & "|" & NVL(rsCurImageData("ͼ��UID")) & "|" & NVL(rsCurImageData("����UID")) & "|" & NVL(rsCurImageData("���UID"))
            ReDim Preserve arrImages(UBound(arrImages) + 1)
            
            arrImages(UBound(arrImages)) = strItem
        End If
        rsCurImageData.MoveNext
    Loop
    
    LoadViewImageToFaceWithNormal = IIf(Trim(strCurInstanceUids) <> "" And blnIsAddImage = True, True, False)
    
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect

BUGEX "LoadViewImageToFaceWithNormal End"
End Function


Private Function LoadViewImageToFaceWithService(rsCurImageData As ADODB.Recordset, ByRef arrImages() As String) As Boolean
'ͨ��ZLPacsServerCenter�������Ԥ��ͼ�񵽽���
'����Ԥ��ͼ�񵽽���
    Dim strTmpFile As String
    Dim strCachePath As String
    
    Dim curImage As DicomImage
    
    Dim objFile As New Scripting.FileSystemObject
    
    Dim dcmTag As clsImageTagInf
    Dim strImgInstanceUid As String
    Dim strCurInstanceUids As String
    Dim blnIsAddImage As Boolean
    Dim fileMsg As TransferFileMsg
    Dim blnIsSendOk As Boolean
    Dim strItem As String
    
    blnIsAddImage = False
    
    LoadViewImageToFaceWithService = False
        
    '��������ͼ�񻺴�Ŀ¼
    MkLocalDir GetResourceDir
    strCachePath = GetCacheDir
    MkLocalDir strCachePath & objFile.GetParentFolderName(NVL(rsCurImageData("URL")))
    
    ReDim arrImages(0)
    Do While Not rsCurImageData.EOF
        'ѭ������ͼ��DicomViewer��
        strImgInstanceUid = Trim(NVL(rsCurImageData!ͼ��UID))
        
        If InStr(strCurInstanceUids, strImgInstanceUid) <= 0 And strImgInstanceUid <> "" Then
            blnIsAddImage = True
            
            '��������Ƶ����ʾ�ļ������Ϊ����Ƶ�ļ�ʱ���ù��̽����ӷ�������ֱ�����������ļ�
            If NVL(rsCurImageData!��̬ͼ, imgTag) = VIDEOTAG Then
                strTmpFile = GetResourceDir & "Avi.bmp"
            ElseIf NVL(rsCurImageData!��̬ͼ, imgTag) = AUDIOTAG Then
                strTmpFile = GetResourceDir & "wav.bmp"
            Else
                strTmpFile = strCachePath & NVL(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")
            End If
            
            strTmpFile = Trim(strTmpFile)
            
            blnIsSendOk = True
            
            If Dir(strTmpFile) = vbNullString Then
                '���ػ���ͼ�񲻴��ڣ����ļ����ݷ�����������ʹ�÷����̨����
                With fileMsg
                    fileMsg.strAdviceId = Val(NVL(rsCurImageData("ҽ��ID")))
                    fileMsg.strName = NVL(rsCurImageData("����"))
                    fileMsg.strSex = NVL(rsCurImageData("�Ա�"))
                    fileMsg.strAge = NVL(rsCurImageData("����"))
                    
                    fileMsg.ftpInfo.strDeviceId = NVL(rsCurImageData("�豸��1"))
                    fileMsg.ftpInfo.strFtpDir = NVL(rsCurImageData("Root1"))
                    fileMsg.ftpInfo.strFTPIP = NVL(rsCurImageData("Host1"))
                    fileMsg.ftpInfo.strFTPPwd = NVL(rsCurImageData("Pwd1"))
                    fileMsg.ftpInfo.strFTPUser = NVL(rsCurImageData("User1"))
                    fileMsg.ftpInfo.strSDDir = NVL(rsCurImageData("����Ŀ¼1"))
                    fileMsg.ftpInfo.strSDPswd = NVL(rsCurImageData("����Ŀ¼����1"))
                    fileMsg.ftpInfo.strSDUser = NVL(rsCurImageData("����Ŀ¼�û���1"))
                    
                    fileMsg.bakFtpInfo.strDeviceId = NVL(rsCurImageData("�豸��2"))
                    fileMsg.bakFtpInfo.strFtpDir = NVL(rsCurImageData("Root2"))
                    fileMsg.bakFtpInfo.strFTPIP = NVL(rsCurImageData("Host2"))
                    fileMsg.bakFtpInfo.strFTPPwd = NVL(rsCurImageData("Pwd2"))
                    fileMsg.bakFtpInfo.strFTPUser = NVL(rsCurImageData("User2"))
                    fileMsg.bakFtpInfo.strSDDir = NVL(rsCurImageData("����Ŀ¼2"))
                    fileMsg.bakFtpInfo.strSDPswd = NVL(rsCurImageData("����Ŀ¼����2"))
                    fileMsg.bakFtpInfo.strSDUser = NVL(rsCurImageData("����Ŀ¼�û���2"))
                    
                    fileMsg.strLocalDir = strTmpFile
                    fileMsg.strFileName = NVL(rsCurImageData("ͼ��UID")) & IIf(mblnIsLoadReportImage, ".jpg", "")
                    fileMsg.strSubDir = NVL(rsCurImageData("URL"))
                    fileMsg.strMediaType = NVL(rsCurImageData!��̬ͼ, imgTag)
                End With
                
                If Not SendDataToservice("����ͼ", LoadCommand.COMMAND_RPTIMG_DOWNLOAD, "ͼ������", fileMsg) Then
                    blnIsSendOk = False
                End If
            End If
            
            strItem = strTmpFile & "|" & NVL(rsCurImageData!��̬ͼ, imgTag) & "|" & NVL(rsCurImageData!����ͼ) & "|" & _
                NVL(rsCurImageData("��������"), "") & "|" & NVL(rsCurImageData("�ɼ�ʱ��")) & "|" & strCachePath & NVL(rsCurImageData("URL")) & ".avi" & "|" & _
                Val(NVL(rsCurImageData("¼�Ƴ���"), "0")) & "|" & NVL(rsCurImageData("ͼ��UID")) & "|" & NVL(rsCurImageData("����UID")) & "|" & NVL(rsCurImageData("���UID"))
            ReDim Preserve arrImages(UBound(arrImages) + 1)
            arrImages(UBound(arrImages)) = strItem
        End If
        
        rsCurImageData.MoveNext
    Loop
    
    LoadViewImageToFaceWithService = IIf(Trim(strCurInstanceUids) <> "" And blnIsAddImage = True, True, False)
End Function

Private Function SendDataToservice(ByVal strDataTag As String, ByVal intCommandIdentify As Integer, ByVal strDataFrom As String, fileMsg As TransferFileMsg)
    Dim objServiceHelper As New clsServiceHelper
    
    SendDataToservice = objServiceHelper.SendDataToservice(strDataTag, intCommandIdentify, strDataFrom, fileMsg)
    
    Set objServiceHelper = Nothing
End Function

Public Function GetRpRecordCount(ByVal slQueryLevel As tQueryLevel, ByVal strSearchValue As String, ByVal blnTmpRecord As Boolean, blnMoved As Boolean) As Long
'��ȡͼ������
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    strSQL = "Select Count(B.Column_Value) ����ֵ From Ӱ�����¼ A, Table(Cast(f_Str2list(Replace(A.����ͼ��,';',',')) As zlTools.t_Strlist)) B Where ҽ��ID = [1]"
    '�����ѯ��ʱ��¼������Ҫ����ѯ���滻Ϊ��ʱ�洢���ݵı�
    If blnTmpRecord Then
        strSQL = Replace(strSQL, "Ӱ����", "Ӱ����ʱ")
    Else
        If blnMoved Then
            strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        End If
    End If
        
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯͼ������", strSearchValue)
    If rsData.RecordCount > 0 Then lngRecordCount = NVL(rsData!����ֵ)
    
    If lngRecordCount <= 0 Then
        Select Case slQueryLevel
            Case slAdvice
                strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c where a.����UID=b.����UID and b.���UID=c.���UID and nvl(a.��̬ͼ,0)=0 and c.ҽ��ID=[1]"
            Case slStudy
                strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b where a.����UID=b.����UID and nvl(a.��̬ͼ,0)=0 and b.���UID=[1]"
            Case slSeries
                strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ��  where nvl(��̬ͼ,0)=0 and ����UID=[1]"
            Case slImage
                strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ��  where nvl(��̬ͼ,0)=0 and ͼ��UID=[1]"
        End Select
        
        '�����ѯ��ʱ��¼������Ҫ����ѯ���滻Ϊ��ʱ�洢���ݵı�
        If blnTmpRecord Then
            strSQL = Replace(strSQL, "Ӱ����", "Ӱ����ʱ")
        Else
            If blnMoved Then
                strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
                strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
            End If
        End If
    
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯͼ������", strSearchValue)
        
        If rsData.RecordCount > 0 Then
            lngRecordCount = NVL(rsData!����ֵ)
        Else
            lngRecordCount = 0
        End If
    End If
    
    GetRpRecordCount = lngRecordCount
End Function

Public Function GetRecordCount(ByVal slQueryLevel As tQueryLevel, ByVal strSearchValue As String, ByVal blnTmpRecord As Boolean, ByVal blnMoved As Boolean) As Long
'��ȡͼ������
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    Select Case slQueryLevel
        Case slAdvice
            strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c where a.����UID=b.����UID and b.���UID=c.���UID and c.ҽ��ID=[1]"
        Case slStudy
            strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b where a.����UID=b.����UID and b.���UID=[1]"
        Case slSeries
            strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ��  where  ����UID=[1]"
        Case slImage
            strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ��  where  ͼ��UID=[1]"
    End Select
    
    '�����ѯ��ʱ��¼������Ҫ����ѯ���滻Ϊ��ʱ�洢���ݵı�
    If blnTmpRecord Then
        strSQL = Replace(strSQL, "Ӱ����", "Ӱ����ʱ")
    Else
        If blnMoved Then
            strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        End If
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯͼ������", strSearchValue)
    
    If rsData.RecordCount > 0 Then
        lngRecordCount = NVL(rsData!����ֵ)
    Else
        lngRecordCount = 0
    End If
    
    GetRecordCount = lngRecordCount
End Function

Public Function GetImageIdex(ByVal slQueryLevel As tQueryLevel, ByVal strSearchValue As String, ByVal blnTmpRecord As Boolean, ByVal blnMoved As Boolean, strUid As String) As Long
'����ͼ��UID��ȡ����
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    Select Case slQueryLevel
        Case slAdvice
            strSQL = "select rownum as ˳���,a.ͼ��UID,a.����UID,a.ͼ��� from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c where a.����UID=b.����UID and b.���UID=c.���UID and c.ҽ��ID=[1]"
        Case slStudy
            strSQL = "select rownum as ˳���,a.ͼ��UID,a.����UID,a.ͼ��� from Ӱ����ͼ�� a, Ӱ�������� b where a.����UID=b.����UID and b.���UID=[1]"
        Case slSeries
            strSQL = "select rownum as ˳���,a.ͼ��UID,a.����UID,a.ͼ��� from Ӱ����ͼ��  where  ����UID=[1]"
        Case slImage
            strSQL = "select rownum as ˳���,a.ͼ��UID,a.����UID,a.ͼ��� from Ӱ����ͼ��  where  ͼ��UID=[1]"
    End Select
    strSQL = "select * from (" & strSQL & " order by ����UID, ͼ���) where ͼ��UID = [2]"
    '�����ѯ��ʱ��¼������Ҫ����ѯ���滻Ϊ��ʱ�洢���ݵı�
    If blnTmpRecord Then
        strSQL = Replace(strSQL, "Ӱ����", "Ӱ����ʱ")
    Else
        If blnMoved Then
            strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        End If
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯͼ������", strSearchValue, strUid)
    
    If rsData.RecordCount > 0 Then
        lngRecordCount = NVL(rsData!˳���)
    Else
        lngRecordCount = 0
    End If
    
    GetImageIdex = lngRecordCount
End Function
