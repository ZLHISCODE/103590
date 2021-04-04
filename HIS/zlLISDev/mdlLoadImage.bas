Attribute VB_Name = "mdlLoadImage"
Option Explicit
'��ȡͼƬ���ݹ���ģ��

Public gcnOracle As New Connection                                  '��������
Public grsParas As ADODB.Recordset                                  'ϵͳ��������
Public grsUserParas As ADODB.Recordset                              'ϵͳ��������
Public gComLib As Object                                            '��������
Public gblnInit As Boolean                                          '�Ƿ��ʼ��
Public gstrSql  As String
Public glngSys As Long                                              'ϵͳ��
Public gstrComputerName As String

Public Function DrawImgAndSaveFile(ByVal strType As String, ByVal strData As String, ByVal strFileName As String, Optional ByVal intSaveType As Integer) As Boolean
    '���ݴ���Ĳ�����ͼ��������Ϊָ���ļ�
    Dim frmDraw As Form
    Set frmDraw = New frmChart
    frmDraw.Hide
    DrawImgAndSaveFile = frmDraw.DrawImg(strType, strData, strFileName, intSaveType)
    Unload frmDraw
    Set frmDraw = Nothing
End Function

Public Function FunFtpSet(blnFtp As Boolean, intVer As Integer, strFtpPara As String, strFtpUser As String, strFtpPass As String, strFtpIP As String, strFtpDir As String)
    On Error GoTo errHandle
    blnFtp = False
    If strFtpPara = "" Then
        If intVer = 0 Then
            strFtpPara = GetPara("FTP����", glngSys, 1208, "")
        Else
            strFtpPara = GetPara("FTP����", glngSys, 2500, "")
        End If
    End If
    If UBound(Split(strFtpPara, ";")) >= 3 Then
        strFtpUser = Split(strFtpPara, ";")(0)
        strFtpPass = Split(strFtpPara, ";")(1)
        strFtpIP = Split(strFtpPara, ";")(2)
        strFtpDir = Split(strFtpPara, ";")(3)
        If TestFTP(strFtpUser, strFtpPass, strFtpIP, strFtpDir) = "" Then
            blnFtp = True
        End If
    End If
    Exit Function
errHandle:
    WriteLog "FunFtpSet", CStr(Erl()) & "�� ", Err.Description
End Function

Public Function LoadImageDataTwo(ByVal strPath As String, ByVal lngID As Long, Optional ByVal intSaveType As Integer, Optional ByVal intVer As Integer, _
                                 Optional ByVal strFileName As String) As Boolean
        '�����ݿ��ȡһ��ͼ�����ݣ����ƺ󱣴浽ָ����·���¡�
        '��Σ�
        '   strPath ·��
        '   lngID   ����ͼ������ID
        '   intSaveType :ֻ���ͼƬ���ͣ�0-cht(Ĭ��) 1-jpg,2-png
        '   intVer      :�汾
        '--����еĻ�, ɾ��ԭ������ʱͼ���ļ�
        
        Dim rsTmp As New ADODB.Recordset, rsImage As New ADODB.Recordset
        Dim rsItem As New ADODB.Recordset
        Dim strImageType As String
        Dim strImageData As String
        Dim DrawIndex As Integer
        Dim intLoop As Integer
        Dim lngStart As Long
        Dim strTmp As String
        Dim strSQL  As String
    
        Dim blnPic As Boolean '�Ƿ�ͼƬ��ʽ
        Dim lngFileNum As Long, lngCount As Long, lngBound As Long
        Dim aryChunk() As Byte, strFile As String
        Dim intLayOut As Integer
        Dim objPic As New frmChartPic
        Dim killFile As String
    
        Dim strDownOk As String, strFtpPath   As String, strLocalFile As String
        Dim objStream As textStream
        Dim strFileType As String
        
        
        On Error GoTo errHandle
100     If intSaveType = 1 Then
102         strFileType = ".jpg"
104     ElseIf intSaveType = 2 Then
106         strFileType = ".png"
        Else
108         strFileType = ".cht"
        End If
        
        If strFileName = "" Then strFileName = lngID & strFileType
110     If Dir(strPath & "\" & strFileName) <> "" Then
112         LoadImageDataTwo = True
            Exit Function
        End If
            
138     lngCount = 0
140     strFile = ""

        If intVer = 0 Then
           strSQL = "select �걾id,ͼ������ from ����ͼ���� where id = [1] "
        Else
           strSQL = "select �걾id,ͼ������ from ���鱨��ͼ�� where id = [1] "
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "zlLISDev.LoadImageData", lngID)

160     If rsTmp.EOF = True Then Exit Function
        
162     Do Until rsTmp.EOF
            strImageType = Trim("" & rsTmp("ͼ������"))
            '- ͼ��������ݿ��У���ԭ���ķ�ʽ����
            If intVer = 0 Then
                gstrSql = "select Zl_FUN_Get����ͼ��(" & rsTmp("�걾id") & ",'" & Nvl(rsTmp("ͼ������")) & "',0) from dual "
            Else
                gstrSql = "select Zl_FUN_Get����ͼ��(" & rsTmp("�걾id") & ",'" & Nvl(rsTmp("ͼ������")) & "',0) from dual "
            End If
            Set rsImage = OpenSQLRecord(gstrSql, "LoadImgData")
            strTmp = Nvl(rsImage(0))
            strImageData = strImageData & Replace(Replace(Trim(strTmp), vbCr, ""), vbLf, "")
        
            If strImageData <> "" Then
                 intLoop = 0
                
                If Val(Mid(strImageData, 1, 3)) >= 100 And Val(Mid(strImageData, 1, 3)) <= 227 And Mid(strImageData, 4, 1) = ";" Then
                
                    blnPic = True
                    If Mid(strImageData, 1, 3) >= 100 And Mid(strImageData, 1, 3) <= 107 Then
                        strFile = App.Path & "\zlLisPic" & lngID & ".bmp"
                    ElseIf Mid(strImageData, 1, 3) >= 110 And Mid(strImageData, 1, 3) <= 117 Then
                         strFile = App.Path & "\zlLisPic" & lngID & ".jpg"
                    ElseIf Mid(strImageData, 1, 3) >= 120 And Mid(strImageData, 1, 3) <= 127 Then
                         strFile = App.Path & "\zlLisPic" & lngID & ".gif"
                    ElseIf Mid(strImageData, 1, 3) >= 200 And Mid(strImageData, 1, 3) <= 227 Then
                        If gobjFSO.FolderExists(App.Path & "\ZLLIS_ZIP") = False Then
                            gobjFSO.CreateFolder App.Path & "\ZLLIS_ZIP"
                        End If
                        If gobjFSO.FolderExists(App.Path & "\ZLLIS_ZIP\" & lngID) = False Then
                            gobjFSO.CreateFolder App.Path & "\ZLLIS_ZIP\" & lngID
                        End If
                        strFile = App.Path & "\ZLLIS_ZIP\" & lngID & "\ZLISPIC.ZIP"
                        End If
                    
                    
                        intLayOut = Val(Mid(strImageData, 1, 3))
                        strImageData = Mid(strImageData, 5)
                        lngFileNum = FreeFile
                        lngCount = 0
    
                    If Dir(strFile) <> "" Then Kill strFile
                    Open strFile For Binary As lngFileNum
                    ReDim aryChunk(Len(strImageData) / 2 - 1) As Byte
                    For lngBound = LBound(aryChunk) To UBound(aryChunk)
                        aryChunk(lngBound) = CByte("&H" & Mid(strImageData, lngBound * 2 + 1, 2))
                    Next
                    
                    Put lngFileNum, , aryChunk()
                    
                End If
                    '-------����ΪͼƬ�ļ�
                Do While strTmp <> ""
                    intLoop = intLoop + 1
                    If intVer = 0 Then
                        gstrSql = "select Zl_FUN_Get����ͼ��(" & rsTmp("�걾id") & ",'" & Nvl(rsTmp("ͼ������")) & "'," & intLoop & ") from dual "
                    Else
                        gstrSql = "select Zl_FUN_Get����ͼ��(" & rsTmp("�걾id") & ",'" & Nvl(rsTmp("ͼ������")) & "'," & intLoop & ") from dual "
                    End If
                    Set rsImage = OpenSQLRecord(gstrSql, "LoadImgData")
                    
                    strTmp = Nvl(rsImage(0))
    
                    If blnPic Then
                            '
                        If strTmp <> "" Then
                            ReDim aryChunk(Len(strTmp) / 2 - 1) As Byte
                            For lngBound = LBound(aryChunk) To UBound(aryChunk)
                                aryChunk(lngBound) = CByte("&H" & Mid(strTmp, lngBound * 2 + 1, 2))
                            Next
                            
                            Put lngFileNum, , aryChunk()
                        End If
                    Else
                        'ͼ������
                        strImageData = strImageData & Replace(Replace(Trim(strTmp), vbCr, ""), vbLf, "")
                    End If
                Loop
                
                If blnPic Then
                    strImageData = intLayOut & ";" & strFile
                    Close lngFileNum
                End If
            End If
        
304         If Len(strImageData) <> 0 Then
                '��ͼ������ͼ���ļ�
306             LoadImageDataTwo = DrawImgAndSaveFile(strImageType, strImageData, strPath & "\" & strFileName, intSaveType)
                
            End If
308         intLoop = 0
310         Do Until intLoop > 100
312             intLoop = intLoop + 1
314             If gobjFSO.FileExists(strFile) Then
316                 WriteLog "LoadImageData", "��" & intLoop & "��ɾ��ԭʼ�ļ�" & strFile, ""
318                 Call gobjFSO.DeleteFile(strFile)
                Else
320                 If strFile <> "" Then WriteLog "LoadImageData", "ԭʼ�ļ�" & strFile & "��ɾ��!", ""
                    Exit Do
                End If
            Loop
'322         intLoop = 0
'324         Do Until intLoop > 100
'326             intLoop = intLoop + 1
'328             If gobjFSO.FileExists(strLocalFile) Then
'330                 WriteLog "LoadImageData", "��" & intLoop & "��ɾ��FTP���ص�ԭʼ�ļ�" & strLocalFile, ""
'332                 Call gobjFSO.DeleteFile(strLocalFile)
'                Else
'334                 If strLocalFile <> "" Then WriteLog "LoadImageData", "FTP���ص�ԭʼ�ļ�" & strLocalFile & "��ɾ��!", ""
'                    Exit Do
'                End If
'            Loop
336         strTmp = "": strImageData = ""
338         rsTmp.MoveNext
        Loop
        Exit Function
errHandle:
340     WriteLog "LoadImagedata", CStr(Erl()) & "�� ", Err.Description
End Function

Public Function LoadImageData(ByVal strPath As String, ByVal lngID As Long, Optional ByVal intSaveType As Integer, Optional ByVal intVer As Integer, Optional ByVal strFileName As String) As Boolean
        '�����ݿ��ȡһ��ͼ�����ݣ����ƺ󱣴浽ָ����·���¡�
        '��Σ�
        '   strPath ·��
        '   lngID   ����ͼ������ID
        '   intSaveType :ֻ���ͼƬ���ͣ�0-cht(Ĭ��) 1-jpg,2-png
        '   intVer      :�汾
        '--����еĻ�, ɾ��ԭ������ʱͼ���ļ�
        
        Dim rsTmp As New ADODB.Recordset, rsImage As New ADODB.Recordset
        Dim rsItem As New ADODB.Recordset
        Dim strImageType As String
        Dim strImageData As String
        Dim DrawIndex As Integer
        Dim intLoop As Integer
        Dim lngStart As Long
        Dim strTmp As String
        Dim strSQL  As String
    
        Dim blnPic As Boolean '�Ƿ�ͼƬ��ʽ
        Dim lngFileNum As Long, lngCount As Long, lngBound As Long
        Dim aryChunk() As Byte, strFile As String
        Dim intLayOut As Integer
        Dim objPic As New frmChartPic
        Dim killFile As String
    
        Dim blnFtp As Boolean       'FTP�Ƿ����
        Static strFtpPara As String       '����FTP����
        Dim strFtpUser As String, strFtpPass As String, strFtpIP As String, strFtpDir As String
        Dim strDownOk As String, strFtpPath   As String, strLocalFile As String
        Dim objStream As textStream
        Dim strFileType As String
        
        
        On Error GoTo errHandle
100     If intSaveType = 1 Then
102         strFileType = ".jpg"
104     ElseIf intSaveType = 2 Then
106         strFileType = ".png"
        Else
108         strFileType = ".cht"
        End If
        
        If strFileName = "" Then strFileName = lngID & strFileType
110     If Dir(strPath & "\" & strFileName) <> "" Then
112         LoadImageData = True
            Exit Function
        End If
    
        'FTP���Ӽ�飬��Ч����԰�FTP��ʽȡͼƬ
114     blnFtp = False
116     If strFtpPara = "" Then
118         If intVer = 0 Then
120             strFtpPara = GetPara("FTP����", glngSys, 1208, "")
            Else
122             strFtpPara = GetPara("FTP����", glngSys, 2500, "")
            End If
        End If
124     If UBound(Split(strFtpPara, ";")) >= 3 Then
126        strFtpUser = Split(strFtpPara, ";")(0)
128        strFtpPass = Split(strFtpPara, ";")(1)
130        strFtpIP = Split(strFtpPara, ";")(2)
132        strFtpDir = Split(strFtpPara, ";")(3)
134        If TestFTP(strFtpUser, strFtpPass, strFtpIP, strFtpDir) = "" Then
136             blnFtp = True
           End If
        End If
        
138     lngCount = 0
140     strFile = ""
        
142     If blnFtp Then
144          If intVer = 0 Then
146             strSQL = "select �걾id,ͼ������,ͼ��λ�� from ����ͼ���� where id = [1]"
             Else
148             strSQL = "select �걾id,ͼ������,ͼ��λ�� from ���鱨��ͼ�� where id = [1]"
             End If
150          Set rsTmp = OpenSQLRecord(strSQL, "zlLISDev.LoadImageData", lngID)
        Else
152          If intVer = 0 Then
154             strSQL = "select �걾id,ͼ������ from ����ͼ���� where id = [1] "
             Else
156             strSQL = "select �걾id,ͼ������ from ���鱨��ͼ�� where id = [1] "
             End If
158          Set rsTmp = OpenSQLRecord(strSQL, "zlLISDev.LoadImageData", lngID)
             
        End If
160     If rsTmp.EOF = True Then Exit Function

    
        
162     Do Until rsTmp.EOF
164         strImageType = Trim("" & rsTmp("ͼ������"))
166         strFtpPath = ""
168         If blnFtp Then strFtpPath = Trim("" & rsTmp!ͼ��λ��)
170         If InStr(strFtpPath, ";") <= 0 Or Not blnFtp Then
                '- ͼ��������ݿ��У���ԭ���ķ�ʽ����
                If intVer = 0 Then
172                 gstrSql = "select Zl_FUN_Get����ͼ��(" & rsTmp("�걾id") & ",'" & Nvl(rsTmp("ͼ������")) & "',0) from dual "
                Else
                    gstrSql = "select Zl_FUN_Get����ͼ��(" & rsTmp("�걾id") & ",'" & Nvl(rsTmp("ͼ������")) & "',0) from dual "
                End If
174             Set rsImage = OpenSQLRecord(gstrSql, "LoadImgData")
176             strTmp = Nvl(rsImage(0))
178             strImageData = strImageData & Replace(Replace(Trim(strTmp), vbCr, ""), vbLf, "")
            
180             If strImageData <> "" Then
182                 intLoop = 0
                
184                 If Val(Mid(strImageData, 1, 3)) >= 100 And Val(Mid(strImageData, 1, 3)) <= 227 And Mid(strImageData, 4, 1) = ";" Then
                
186                     blnPic = True
188                     If Mid(strImageData, 1, 3) >= 100 And Mid(strImageData, 1, 3) <= 107 Then
190                         strFile = App.Path & "\zlLisPic" & lngID & ".bmp"
192                     ElseIf Mid(strImageData, 1, 3) >= 110 And Mid(strImageData, 1, 3) <= 117 Then
194                         strFile = App.Path & "\zlLisPic" & lngID & ".jpg"
196                     ElseIf Mid(strImageData, 1, 3) >= 120 And Mid(strImageData, 1, 3) <= 127 Then
198                         strFile = App.Path & "\zlLisPic" & lngID & ".gif"
200                     ElseIf Mid(strImageData, 1, 3) >= 200 And Mid(strImageData, 1, 3) <= 227 Then
202                         If gobjFSO.FolderExists(App.Path & "\ZLLIS_ZIP") = False Then
204                             gobjFSO.CreateFolder App.Path & "\ZLLIS_ZIP"
                            End If
206                         If gobjFSO.FolderExists(App.Path & "\ZLLIS_ZIP\" & lngID) = False Then
208                             gobjFSO.CreateFolder App.Path & "\ZLLIS_ZIP\" & lngID
                            End If
210                         strFile = App.Path & "\ZLLIS_ZIP\" & lngID & "\ZLISPIC.ZIP"
                        End If
                    
                    
212                     intLayOut = Val(Mid(strImageData, 1, 3))
214                     strImageData = Mid(strImageData, 5)
216                     lngFileNum = FreeFile
218                     lngCount = 0
    
220                     If Dir(strFile) <> "" Then Kill strFile
222                     Open strFile For Binary As lngFileNum
224                     ReDim aryChunk(Len(strImageData) / 2 - 1) As Byte
226                     For lngBound = LBound(aryChunk) To UBound(aryChunk)
228                         aryChunk(lngBound) = CByte("&H" & Mid(strImageData, lngBound * 2 + 1, 2))
                        Next
                    
230                     Put lngFileNum, , aryChunk()
                    
                    End If
                    '-------����ΪͼƬ�ļ�
232                 Do While strTmp <> ""
234                     intLoop = intLoop + 1
                        If intVer = 0 Then
236                         gstrSql = "select Zl_FUN_Get����ͼ��(" & rsTmp("�걾id") & ",'" & Nvl(rsTmp("ͼ������")) & "'," & intLoop & ") from dual "
                        Else
                            gstrSql = "select Zl_FUN_Get����ͼ��(" & rsTmp("�걾id") & ",'" & Nvl(rsTmp("ͼ������")) & "'," & intLoop & ") from dual "
                        End If
238                     Set rsImage = OpenSQLRecord(gstrSql, "LoadImgData")
                    
240                     strTmp = Nvl(rsImage(0))
    
242                     If blnPic Then
                            '
244                         If strTmp <> "" Then
246                             ReDim aryChunk(Len(strTmp) / 2 - 1) As Byte
248                             For lngBound = LBound(aryChunk) To UBound(aryChunk)
250                                 aryChunk(lngBound) = CByte("&H" & Mid(strTmp, lngBound * 2 + 1, 2))
                                Next
                            
252                             Put lngFileNum, , aryChunk()
                            End If
                        Else
                            'ͼ������
254                         strImageData = strImageData & Replace(Replace(Trim(strTmp), vbCr, ""), vbLf, "")
                        End If
                    Loop
                
256                 If blnPic Then
258                     strImageData = intLayOut & ";" & strFile
260                     Close lngFileNum
                    End If
                End If
            Else
                'ͼ�����FTP�У���FTP��ȡ����
                'ͼ��λ�õ����ݸ�ʽΪ��ͼ���ʽ;FTP�ļ�·��
            
262             intLayOut = Val(Split(strFtpPath, ";")(0))
264             strFtpPath = Trim(Split(strFtpPath, ";")(1))
266             strImageData = ""
268             If intLayOut >= 100 And intLayOut <= 227 Then
                    ' ͼƬ�ļ���ֱ�����ص�����
270                 strLocalFile = strPath & "\zlLisPic" & Split(strFtpPath, "/")(UBound(Split(strFtpPath, "/")))
272                 If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
274                 strDownOk = DownFile(strFtpUser, strFtpPass, strFtpIP, strFtpPath, strLocalFile)
276                 If strDownOk = "" Then
278                     strImageData = intLayOut & ";" & strLocalFile
                    End If
                Else
                    ' ͼ�����ݣ���Ҫ�����ص��ı��ļ��ж�ȡ����
280                 strLocalFile = strPath & "\" & lngID & "_" & strImageType & ".txt"
282                 If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
284                 strDownOk = DownFile(strFtpUser, strFtpPass, strFtpIP, strFtpPath, strLocalFile)
286                 If strDownOk = "" Then
288                     Set objStream = gobjFSO.OpenTextFile(strLocalFile, ForReading)
290                     Do Until objStream.AtEndOfLine
292                         strImageData = strImageData & objStream.ReadLine
                        Loop
294                     objStream.Close
296                     Set objStream = Nothing
298                     strImageData = Replace(Replace(Trim(strImageData), vbCr, ""), vbLf, "")
300                     strImageData = intLayOut & ";" & strImageData
                    End If
302                 If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
                End If
            End If
        
304         If Len(strImageData) <> 0 Then
                '��ͼ������ͼ���ļ�
306             LoadImageData = DrawImgAndSaveFile(strImageType, strImageData, strPath & "\" & strFileName, intSaveType)
            End If
308         intLoop = 0
310         Do Until intLoop > 100
312             intLoop = intLoop + 1
314             If gobjFSO.FileExists(strFile) Then
316                 WriteLog "LoadImageData", "��" & intLoop & "��ɾ��ԭʼ�ļ�" & strFile, ""
318                 Call gobjFSO.DeleteFile(strFile)
                Else
320                 If strFile <> "" Then WriteLog "LoadImageData", "ԭʼ�ļ�" & strFile & "��ɾ��!", ""
                    Exit Do
                End If
            Loop
322         intLoop = 0
324         Do Until intLoop > 100
326             intLoop = intLoop + 1
328             If gobjFSO.FileExists(strLocalFile) Then
330                 WriteLog "LoadImageData", "��" & intLoop & "��ɾ��FTP���ص�ԭʼ�ļ�" & strLocalFile, ""
332                 Call gobjFSO.DeleteFile(strLocalFile)
                Else
334                 If strLocalFile <> "" Then WriteLog "LoadImageData", "FTP���ص�ԭʼ�ļ�" & strLocalFile & "��ɾ��!", ""
                    Exit Do
                End If
            Loop
336         strTmp = "": strImageData = ""
338         rsTmp.MoveNext
        Loop
        Exit Function
errHandle:
340     WriteLog "LoadImagedata", CStr(Erl()) & "�� ", Err.Description
End Function


Public Function CheckGif(ByVal strFile As String) As Boolean
    '���GIF�ļ������Ƿ�����
    'GIF��ͷ��00 3B����
    Dim intFileNo As Integer, lngFileSize As Long, arrEnd(2) As Byte, arrTitle(3) As Byte
    Dim lngCount As Long
    On Error GoTo hErr
100 If Dir(strFile) <> "" Then
102     intFileNo = FreeFile
104     Open strFile For Binary Access Read As intFileNo
106     lngFileSize = LOF(intFileNo)
108     If lngFileSize > 0 Then
110         Get intFileNo, , arrTitle
112         Seek intFileNo, lngFileSize - 1
114         Get intFileNo, , arrEnd
        End If
116     Close intFileNo
        
118     If UCase(Chr(arrTitle(0)) & Chr(arrTitle(1)) & Chr(arrTitle(2))) = "GIF" And arrEnd(0) = 0 And arrEnd(1) = 59 Then
120         CheckGif = True
        End If
        '�ж��Ƿ��Լ�����ͼƬ����ġ�ʹ�ÿؼ�����ͼƬ��
        '����gif��ʽͼƬ����ֻ�Ǻ�׺����Ϊgif��ʽ��ʵ�ʱ����ͼƬ����bmpͼƬ��
        If UCase(Chr(arrTitle(0)) & Chr(arrTitle(1))) = "BM" Then
            CheckGif = True
        End If
    End If
    Exit Function
hErr:
122     WriteLog "CheckGif", CStr(Erl()) & "�� ", Err.Description
    
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'               ��Ϊʹ�ð󶨱���,�Դ�"'"���ַ�����,����Ҫʹ��"''"��ʽ��
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strlog As String, varValue As Variant
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        If lngRight = 0 Then Exit Do
        '������������"[����]����"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '�滻Ϊ"?"����
    strlog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            strlog = Replace(strlog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strlog = Replace(strlog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strlog = Replace(strlog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '���ԭ�в���:��Ȼ�����ظ�ִ��
'    cmdData.CommandText = "" '��Ϊ����ʱ�����������
'    Do While cmdData.Parameters.Count > 0
'        cmdData.Parameters.Delete 0
'    Loop
    
    '�����µĲ���
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '����
            '���ַ�ʽ������һЩIN�Ӿ��Union���
            '��ʾͬһ�������Ķ��ֵ,�����Ų�������������Ĳ����Ž���,��Ҫ��֤�����ֵ��������
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strlog = Replace(strlog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '�ַ�
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strlog = Replace(strlog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strlog = Replace(strlog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '�ò������������õ��ڼ���ֵ��
        End Select
    Next

    'ִ�з��ؼ�¼��
    'If cmdData.ActiveConnection Is Nothing Then
    
     Set cmdData.ActiveConnection = gcnOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
    'End If
    cmdData.CommandText = strSQL
    
'    Call gobjComLib.SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Set OpenSQLRecord.ActiveConnection = Nothing
'    Call gobjComLib.SQLTest
End Function

Public Sub ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'���ܣ�ִ�й������,���Զ��Թ��̲������а󶨱�������
'������strSQL=�������,���ܴ�����,����"������(����1,����2,...)"��
'˵�������¼���������̲�����ʹ�ð󶨱���,�����ϵĵ��÷�����
'  1.���������Ǳ��ʽ,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1,100.12*0.15,...)"
'  2.�м�û�д�����ȷ�Ŀ�ѡ����,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1, , ,����3,...)"
'  3.��Ϊ�ù������Զ�����,����һ��ʹ�ð󶨱���,�Դ�"'"���ַ�����,��Ҫʹ��"''"��ʽ��
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    
    If Right(Trim(strSQL), 1) = ")" Then
        '���ԭ�в���:��Ȼ�����ظ�ִ��
'        cmdData.CommandText = "" '��Ϊ����ʱ�����������
'        Do While cmdData.Parameters.Count > 0
'            cmdData.Parameters.Delete 0
'        Loop
        
        'ִ�еĹ�����
        strTemp = Trim(strSQL)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        'ִ�й��̲���
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For i = 1 To Len(strTemp)
            '�Ƿ����ַ����ڣ��Լ����ʽ��������
            If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                strPar = Trim(strPar)
                With cmdData
                    If IsNumeric(strPar) Then '����
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarNumeric, adParamInput, 30, strPar)
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '�ַ���
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        
                        'Oracle���ӷ�����:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '˫"''"�İ󶨱�������
                        If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                        
                        '���Ӳ�������LOBʱ������ð󶨱���ת��ΪRAWʱ����2000���ַ�Ҫ��adLongVarChar
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
                        If intMax <= 2000 Then
                            intMax = IIf(intMax <= 200, 200, 2000)
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarChar, adParamInput, intMax, strPar)
                        Else
                            If intMax < 4000 Then intMax = 4000
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adLongVarChar, adParamInput, intMax, strPar)
                        End If
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '����
                        strPar = Split(strPar, "(")(1)
                        strPar = Trim(Split(strPar, ",")(0))
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        If strPar = "" Then
                            'NULLֵ�������ִ���ɼ�����������
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(strPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adDBTimeStamp, adParamInput, , CDate(strPar))
                        End If
                    ElseIf UCase(strPar) = "SYSDATE" Then '����
                        If datCur = CDate(0) Then datCur = Currentdate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then 'NULLֵ�����ַ�����ɼ�����������
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarChar, adParamInput, 200, Null)
                    ElseIf strPar = "" Then '��ѡ��������NULL������ܸı���ȱʡֵ:��˿�ѡ��������д���м�
                        GoTo NoneVarLine
                    Else '�������������ӵı��ʽ���޷�����
                        GoTo NoneVarLine
                    End If
                End With
                
                strPar = ""
            Else
                strPar = strPar & Mid(strTemp, i, 1)
            End If
        Next
        
        '����Ա���ù���ʱ��д����
        If blnStr Or intBra <> 0 Then
            Err.Raise -2147483645, , "���� Oracle ����""" & strProc & """ʱ�����Ż�������д��ƥ�䡣ԭʼ������£�" & vbCrLf & vbCrLf & strSQL
            Exit Sub
        End If
        
        '����?��
        strTemp = ""
        For i = 1 To cmdData.Parameters.count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        
        
        'ִ�й���
        'If cmdData.ActiveConnection Is Nothing Then
            Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
            cmdData.CommandType = adCmdText
        'End If
        cmdData.CommandText = strProc
        
'        Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
        Call cmdData.Execute
'        Call gobjComLib.SQLTest
    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:
'    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
    
    '˵����Ϊ�˼��������ӷ�ʽ
    '1.��������adCmdStoredProc��ʽ��8i����������
    '2.�����������ʹ��{},��ʹ����û�в���ҲҪ��()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    
    gcnOracle.Execute strSQL, , adCmdText
'    Call gobjComLib.SQLTest
End Sub
Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional inttype As Integer) As String
'���ܣ���ȡָ���Ĳ���ֵ
'������varPara=�����Ż�������������ֻ��ַ����ʹ�������
'      lngSys=ʹ�øò�����ϵͳ��ţ���100
'      lngModual=ʹ�øò�����ģ��ţ���1230
'      strDefault=�����ݿ���û�иò���ʱʹ�õ�ȱʡֵ(ע�ⲻ��Ϊ��ʱ)
'      blnNotCache=�Ƿ񲻴ӻ����ж�ȡ
'      arrControl=�ؼ����飬��Array(Me.Text1, Me.CheckBox1)�����ں����ڲ��Զ������Ӧ�ؼ�����ʾ��ɫ���Ƿ��ֹ���á�
'      blnSetup=����ģ���Ƿ��в�������Ȩ��
'      intType=���ز��������ز�������
'���أ�����ֵ���ַ�����ʽ
    Dim strSQL As String, i As Integer
    Dim blnNew As Boolean, blnEnabled As Boolean
    Dim strDBUser As String
    
    
    strDBUser = GetUserDB()
    On Error GoTo errH
    
    inttype = 0
    
    '��һ�μ��ز�������
    If grsParas Is Nothing Then
        blnNew = True
    ElseIf grsParas.State = 0 Then
        blnNew = True
    End If
    If blnNew Then
        strSQL = "Select ID,Nvl(ϵͳ,0) as ϵͳ,Nvl(ģ��,0) as ģ��,Nvl(˽��,0) as ˽��,Nvl(����,0) as ����,Nvl(��Ȩ,0) as ��Ȩ,������,������," & _
            " Nvl(����ֵ,ȱʡֵ) as ����ֵ,[1] as �û���,[2] as ������ From zlParameters"
        Set grsParas = New ADODB.Recordset
        Set grsParas = OpenSQLRecord(strSQL, "GetPara", strDBUser, gstrComputerName)
        
        strSQL = _
            " Select ����ID,Nvl(�û���,'NullUser') as �û���,Nvl(������,'NullMachine') as ������,����ֵ From zlUserParas Where �û���=[1]" & _
            " Union" & _
            " Select ����ID,Nvl(�û���,'NullUser') as �û���,Nvl(������,'NullMachine') as ������,����ֵ From zlUserParas Where ������=[2]"
        Set grsUserParas = New ADODB.Recordset
        Set grsUserParas = OpenSQLRecord(strSQL, "GetPara", strDBUser, gstrComputerName)
    End If
    
    'ʹ�û���
    If TypeName(varPara) = "String" Then
        grsParas.Filter = "������='" & CStr(varPara) & "' And ģ��=" & lngModual & " And ϵͳ=" & lngSys
    Else
        grsParas.Filter = "������=" & Val(varPara) & " And ģ��=" & lngModual & " And ϵͳ=" & lngSys
    End If
    If Not grsParas.EOF Then
        '��ȡ����ֵ
        If grsParas!˽�� = 1 Or grsParas!���� = 1 Then
            grsUserParas.Filter = "����ID=" & grsParas!id & _
                IIf(grsParas!˽�� = 1, " And �û���='" & grsParas!�û��� & "'", " And �û���='NullUser'") & _
                IIf(grsParas!���� = 1, " And ������='" & grsParas!������ & "'", " And ������='NullMachine'")
            If Not grsUserParas.EOF Then
                GetPara = Nvl(grsUserParas!����ֵ, strDefault)
            Else
                GetPara = Nvl(grsParas!����ֵ, strDefault)
            End If
        Else
            GetPara = Nvl(grsParas!����ֵ, strDefault)
        End If
        
        '���ز������ͣ�1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
        If grsParas!ϵͳ <> 0 And grsParas!ģ�� = 0 And grsParas!˽�� = 0 And grsParas!���� = 0 Then
            inttype = 1
        ElseIf grsParas!ģ�� = 0 And grsParas!˽�� = 1 And grsParas!���� = 0 Then
            inttype = 2
        ElseIf grsParas!ϵͳ <> 0 And grsParas!ģ�� <> 0 And grsParas!˽�� = 0 And grsParas!���� = 0 Then
            inttype = 3
        ElseIf grsParas!ϵͳ <> 0 And grsParas!ģ�� <> 0 And grsParas!˽�� = 1 And grsParas!���� = 0 Then
            inttype = 4
        ElseIf grsParas!ϵͳ <> 0 And grsParas!ģ�� <> 0 And grsParas!˽�� = 0 And grsParas!���� = 1 Then
            inttype = IIf(grsParas!��Ȩ = 1, 15, 5)
        ElseIf grsParas!ϵͳ <> 0 And grsParas!ģ�� <> 0 And grsParas!˽�� = 1 And grsParas!���� = 1 Then
            inttype = 6
        End If
        
        '�����Ӧ�Ŀؼ���ɫ���ɿ�״̬
        If IsArray(arrControl) And (inttype = 3 Or (inttype Mod 10) = 5) Then
            blnEnabled = Not ((inttype = 3 Or (inttype Mod 10) = 5 And grsParas!��Ȩ = 1) And Not blnSetup)
            For i = 0 To UBound(arrControl)
                Select Case TypeName(arrControl(i))
                Case "Label"
                    arrControl(i).ForeColor = vbBlue
                Case "TextBox", "MaskEdBox", "CheckBox", "OptionButton", "ComboBox", "ListBox", "Frame", "PictureBox", "ListView"
                    arrControl(i).ForeColor = vbBlue
                    If Not blnEnabled Then arrControl(i).Enabled = False
                Case "CommandButton", "DTPicker"
                    If Not blnEnabled Then arrControl(i).Enabled = False
                Case "MSHFlexGrid"
                    arrControl(i).ForeColor = vbBlue
                    arrControl(i).ForeColorFixed = vbBlue
                    If Not blnEnabled Then arrControl(i).Enabled = False
                Case "VSFlexGrid"
                    arrControl(i).ForeColor = vbBlue
                    arrControl(i).ForeColorFixed = vbBlue
                    If Not blnEnabled Then arrControl(i).Editable = 0
                Case Else
                    On Error Resume Next
                    arrControl(i).ForeColor = vbBlue
                    If Not blnEnabled Then arrControl(i).Enabled = False
                    Err.Clear: On Error GoTo errH
                End Select
            Next
        End If
    Else
        GetPara = strDefault
    End If
    
    Exit Function
errH:
'    If gobjComLib.ErrCenter() = 1 Then
'        Resume
'    End If
End Function

Public Function GetUserDB() As String
    Dim strTmp As String
    Dim strConnStr As String
    strConnStr = gcnOracle.ConnectionString
    strTmp = Mid(strConnStr, InStr(strConnStr, "User ID="))
    strTmp = Mid(strTmp, 9, InStr(strTmp, ";") - 9)
    GetUserDB = UCase(strTmp)
End Function

Public Function Currentdate() As Date
    '-------------------------------------------------------------
    '���ܣ���ȡ�������ϵ�ǰ����
    '������
    '���أ�����Oracle���ڸ�ʽ�����⣬����
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errH
    With rsTemp
        .CursorLocation = adUseClient
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    Currentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
errH:
'    If gobjComLib.ErrCenter() = 1 Then Resume
    Currentdate = 0
    Err = 0
End Function


Public Sub SaveLog(ByVal StrInput As String)
    '------------------------------------------------------
    '--  ����:���ݵ��Ա�־,д��־����ǰĿ¼
    '------------------------------------------------------

    '���±������ڼ�¼���ýӿڵ����
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As textStream
    Dim objFileSystem As New FileSystemObject

    '���ж��Ƿ���ڸ��ļ����������򴴽�������=0��ֱ���˳���������������������Ϣ��
    strFileName = App.Path & "\LisDev_" & Format(date, "yyyyMMdd") & ".LOG"

    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (strDate & ":" & StrInput)
    objStream.Close
    Set objStream = Nothing
End Sub
