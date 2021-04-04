VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrawLisImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function DrawImg(ByVal strType As String, ByVal strData As String, ByVal strFilename As String, _
                        Optional ByVal intSaveType As Integer) As Boolean
    '�ⲿ���ýӿ�һ
    '���ܣ����ݴ����ͼ�����ݣ���ͼ
    DrawImg = DrawImgAndSaveFile(strType, strData, strFilename, intSaveType)
End Function

Public Function GetSampleImages(ByVal lngSampleId As Long, ByVal blnClearOldFile As Boolean, Optional intSaveImgType As Integer = 0) As String
    '��ȡһ���걾������ͼƬ��
    '
    'lngsampleID : �걾ID
    'strPath     : ͼƬ����·��,���������Զ�����
    'blnClearOldFile :�Ƿ���ջ����ڱ��ص�ͼ���ļ�,True��ÿ�ζ������ݿ���ļ����浽����;False-��һ�ε���ʱ�����ݿ��ͼ�β���ͼƬ��֮��ֱ��ʹ��
    'strErr      : ��������ֵΪ�մ�ʱ�����ص���ʾ��Ϣ
    'intImgType  : ���ص�ͼƬ�ļ���ʽ��0��cht(Ĭ��),1-jgp,2-png

        Dim rsTmp       As New ADODB.Recordset
        Dim DrawIndex   As Integer
        Dim strFileType As String
        Dim strTime     As Date
        Dim strReturn   As String
        Dim strFilename As String
        Dim objFso      As New FileSystemObject
        Dim strSql      As String
        
104     strReturn = ""
    
106     If intSaveImgType = 1 Then
108         strFileType = ".jpg"
110     ElseIf intSaveImgType = 2 Then
112         strFileType = ".png"
        Else
114         strFileType = ".cht"
        End If

118     If Trim(gstrFilePath) = "" Then
120         Call ErrLog("GetSampleImages", CStr(Erl()) & "��", "��ָ������ͼƬ��·����", "")
            Exit Function
'122     ElseIf Not objFso.FolderExists(gstrFilePath) Then
'124         Call ErrLog("GetSampleImages", CStr(Erl()) & "��", gstrFilePath & "�����ڣ�", "")
'            Exit Function
        End If
        If Not objFso.FolderExists(gstrFilePath) Then
            objFso.CreateFolder (gstrFilePath)
        End If
126     If Right(gstrFilePath, 1) = "\" Then gstrFilePath = Left(gstrFilePath, Len(gstrFilePath) - 1)
        '---------------------------------------------------
        On Error GoTo ErrH
128     strTime = Now
130     If gblnNewLis Then
132         strSql = "select ID ,�걾ID,ͼ������ from ���鱨��ͼ�� where �걾ID = [1] order by ID"
            Set rsTmp = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName, lngSampleId)
        Else
134         strSql = "select ID ,�걾ID,ͼ������ from ����ͼ���� where �걾ID = [1] order by ID"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "GetSampleImage", lngSampleId)
        End If
    
136
138     If rsTmp.EOF Then Call ErrLog("GetSampleImages", CStr(Erl()) & "��", "��ͼ�����ݣ�", "")
        
    
140     Do Until rsTmp.EOF
            If Not gblnNewLis Then
                strFilename = rsTmp("ID") & strFileType
                'strFilename = rsTmp("�걾ID") & IIf(Trim$("" & rsTmp("ͼ������")) <> "", "_" & Trim$("" & rsTmp("ͼ������")), "") & strFileType
            Else
                strFilename = rsTmp("�걾ID") & IIf(Trim$("" & rsTmp("ͼ������")) <> "", "_" & Trim$("" & rsTmp("ͼ������")), "") & strFileType
            End If
142         If blnClearOldFile = True Then
                'ǿ����������ڱ��ص�ͼƬ
144             If objFso.FileExists(gstrFilePath & "\" & strFilename) Then
                    Call objFso.DeleteFile(gstrFilePath & "\" & strFilename, True)
                    Call ErrLog("ǿ��������ػ����ͼƬ", gstrFilePath, strFilename, "")
                End If
            End If
        
146         If Not objFso.FileExists(gstrFilePath & "\" & strFilename) Then
        
148             If LoadImageData(gstrFilePath, rsTmp("ID"), 1, strFilename) Then
150                 If objFso.FileExists(gstrFilePath & "\" & strFilename) Then strReturn = strReturn & "#" & strFilename
                End If
            
152             DrawIndex = DrawIndex + 1
            Else
                strReturn = strReturn & "#" & strFilename
            End If
154         rsTmp.MoveNext
        Loop
156     If strReturn <> "" Then strReturn = Mid(strReturn, 2)
158     GetSampleImages = strReturn
        Call ErrLog("����ͼ���ļ���", "", strReturn, "")
        'Debug.Print "ID=" & lngKeyID & ",��ʱ:" & DateDiff("s", strTime, Now)
        Set objFso = Nothing
        Exit Function
ErrH:
160    GetSampleImages = ""
162
164    Call ErrLog("GetSampleImages", CStr(Erl()) & "��", err.Description, "")
End Function

Public Function GetImage(ByVal lngImgID As Long, ByVal blnClearOldFile As Boolean, Optional intSaveImgType As Integer = 0) As String
        '��ȡָ����һ��ͼƬ��
        '
        'lngImgID    : ͼƬID
        'strPath     : ͼƬ����·��,���������Զ�����
        'blnClearOldFile :�Ƿ���ջ����ڱ��ص�ͼ���ļ�,True��ÿ�ζ������ݿ���ļ����浽����;False-��һ�ε���ʱ�����ݿ��ͼ�β���ͼƬ��֮��ֱ��ʹ��
        'strErr      : ��������ֵΪ�մ�ʱ�����ص���ʾ��Ϣ
        'intImgType  : ���ص�ͼƬ�ļ���ʽ��0��cht(Ĭ��),1-jgp,2-png
        
        Dim strFileType As String
        Dim strReturn   As String
        Dim rsTmp       As ADODB.Recordset
        Dim strFilename As String
        Dim objFso      As New FileSystemObject
        Dim strSql      As String
        
104     If intSaveImgType = 1 Then
106         strFileType = ".jpg"
108     ElseIf intSaveImgType = 2 Then
110         strFileType = ".png"
        Else
112         strFileType = ".cht"
        End If

116     If Trim(gstrFilePath) = "" Then
118         Call ErrLog("GetImage", CStr(Erl()) & "��", "��ָ������ͼƬ��·����", "")
            Exit Function
120     ElseIf Not objFso.FolderExists(gstrFilePath) Then
122         Call ErrLog("GetImage", CStr(Erl()) & "��", gstrFilePath & "�����ڣ�", "")
            Exit Function
        End If
124     If Right(gstrFilePath, 1) = "\" Then gstrFilePath = Left(gstrFilePath, Len(gstrFilePath) - 1)
        On Error GoTo ErrH
    
126     If gblnNewLis Then
128         strSql = "select ID ,�걾ID,ͼ������ from ���鱨��ͼ�� where ID = [1]"
            Set rsTmp = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName, lngImgID)
        Else
130         strSql = "select ID ,�걾ID,ͼ������ from ����ͼ���� where ID = [1]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "GetImage", lngImgID)
        End If
    
132
134     If rsTmp.EOF Then Call ErrLog("GetImage", CStr(Erl()) & "��", "��ͼ�����ݣ�", "")
        If Not gblnNewLis Then
            strFilename = rsTmp("ID") & strFileType
        Else
            strFilename = rsTmp("ID") & IIf(Trim$("" & rsTmp("ͼ������")) <> "", "_" & Trim$("" & rsTmp("ͼ������")), "") & strFileType
        End If
136     If blnClearOldFile = True Then
            'ǿ����������ڱ��ص�ͼƬ
138         If objFso.FileExists(gstrFilePath & "\" & strFilename) Then objFso.DeleteFile gstrFilePath & "\" & strFilename, True
            Call ErrLog("ǿ��������ػ����ͼƬ", gstrFilePath, strFilename, "")
        End If
    
140     If Not objFso.FileExists(gstrFilePath & "\" & strFilename) Then
142         If LoadImageData(gstrFilePath, rsTmp("ID"), intSaveImgType, strFilename) Then
144             If objFso.FileExists(gstrFilePath & "\" & strFilename) Then strReturn = strReturn & "," & strFilename
            Else
                Call ErrLog("GetImage", CStr(Erl()) & "��", "����ͼ���ļ�ʧ�ܣ�", "")
            End If
        Else
            strReturn = strReturn & "," & strFilename
        End If
146     If strReturn <> "" Then strReturn = Mid(strReturn, 2)
148     GetImage = strReturn
        Call ErrLog("����ͼ���ļ���", "", strReturn, "")
        Set objFso = Nothing
        Exit Function
ErrH:
150     GetImage = ""
152
154     Call ErrLog("GetImage", CStr(Erl()) & "��", err.Description, "")
End Function

Public Function LocalImageSaveToCht(ByVal strTitle As String, ByVal strImageData As String, ByVal strSaveFileName As String) As Boolean
    '������ͼ���ļ�����Cht����ָ��Ŀ¼
    Dim strPath     As String
    Dim strChtName  As String
    Dim objFso      As New FileSystemObject
    
    strPath = App.Path & "\LisImage"
    If Not objFso.FolderExists(strPath) Then objFso.CreateFolder (strPath)
    strChtName = strPath & "\" & strSaveFileName
    If objFso.FileExists(strChtName) Then objFso.DeleteFile strChtName, True
    
    LocalImageSaveToCht = DrawImgAndSaveFile(strTitle, strImageData, strChtName, 0)
    Set objFso = Nothing
End Function




