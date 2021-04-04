Attribute VB_Name = "mdlLisImage"
Option Explicit
'��ȡͼƬ���ݹ���ģ��

Public Function DrawImgAndSaveFile(ByVal strType As String, ByVal strData As String, ByVal strFilename As String, Optional ByVal intSaveType As Integer) As Boolean
    '���ݴ���Ĳ�����ͼ��������Ϊָ���ļ�
    Dim frmDraw As Form
    Set frmDraw = New frmChart
    frmDraw.Hide
    DrawImgAndSaveFile = frmDraw.DrawImg(strType, strData, strFilename, intSaveType)
    Set frmDraw = Nothing
End Function

Public Function LoadImageData(ByVal strPath As String, ByVal lngͼ��ID As Long, Optional ByVal intSaveType As Integer, Optional ByVal strFilename As String) As Boolean
        '�����ݿ��ȡһ��ͼ�����ݣ����ƺ󱣴浽ָ����·���¡�
        '��Σ�
        '   strPath ·��
        '   lngͼ��ID   ����ͼ������ID
        '   intSaveType :ֻ���ͼƬ���ͣ�0-cht(Ĭ��) 1-jpg,2-png
        
        Dim rsTmp           As ADODB.Recordset
        Dim rsImage         As ADODB.Recordset
        Dim strImageType    As String
        Dim strImageData    As String
        Dim intLoop         As Integer
        Dim strTmp          As String
        Dim strSql          As String
    
        Dim blnPic          As Boolean '�Ƿ�ͼƬ��ʽ
        Dim lngFileNum      As Long
        Dim lngCount        As Long
        Dim lngBound        As Long
        Dim aryChunk()      As Byte
        Dim strFile         As String
        Dim intLayOut       As Integer
    
        Dim blnFtp          As Boolean       'FTP�Ƿ����
        Static strFtpPara   As String       '����FTP����
        Dim strFtpUser      As String
        Dim strFtpPass      As String
        Dim strFtpIP        As String
        Dim strFtpDir       As String
        Dim strDownOk       As String
        Dim strFtpPath      As String
        Dim strLocalFile    As String
        Dim objStream       As TextStream
        Dim strFileType     As String
        Dim objFso          As New FileSystemObject
        
        
On Error GoTo ErrH:
100     If intSaveType = 1 Then
102         strFileType = ".jpg"
104     ElseIf intSaveType = 2 Then
106         strFileType = ".png"
        Else
108         strFileType = ".cht"
        End If
        
        If strFilename = "" Then strFilename = lngͼ��ID & strFileType
110     If Dir(strPath & "\" & strFilename) <> "" Then
112         LoadImageData = True
            Exit Function
        End If
    
        'FTP���Ӽ�飬��Ч����԰�FTP��ʽȡͼƬ
114     blnFtp = False
116     If strFtpPara = "" Then
122         strFtpPara = gobjDatabase.GetPara("FTP����", glngSys, glngModual, "")
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
144         If Not gblnNewLis Then
146             strSql = "select �걾ID,ͼ������,ͼ��λ�� from ����ͼ���� where ID = [1]"
            Else
148             strSql = "select �걾ID,ͼ������,ͼ��λ�� from ���鱨��ͼ�� where ID = [1]"
            End If
150         Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lngͼ��ID)
        Else
152         If Not gblnNewLis Then
154             strSql = "select �걾ID,ͼ������ from ����ͼ���� where ID = [1] "
            Else
156             strSql = "select �걾ID,ͼ������ from ���鱨��ͼ�� where ID = [1] "
            End If
158         Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, gstrSysName, lngͼ��ID)
        End If
160     If rsTmp.EOF = True Then Exit Function

    
        
162     Do Until rsTmp.EOF
164         strImageType = Trim("" & rsTmp("ͼ������"))
166         strFtpPath = ""
168         If blnFtp Then strFtpPath = Trim("" & rsTmp!ͼ��λ��)
170         If InStr(strFtpPath, ";") <= 0 Or Not blnFtp Then
                '- ͼ��������ݿ��У���ԭ���ķ�ʽ����
                If Not gblnNewLis Then
172                 strSql = "select Zl_FUN_Get����ͼ��(" & rsTmp("�걾ID") & ",'" & gobjCommFun.Nvl(rsTmp("ͼ������")) & "',0) from dual "
                Else
                    strSql = "select Zl_FUN_Get����ͼ��(" & rsTmp("�걾ID") & ",'" & gobjCommFun.Nvl(rsTmp("ͼ������")) & "',0) from dual "
                End If
174             Set rsImage = gobjDatabase.OpenSQLRecord(strSql, "LoadImgData")
176             strTmp = gobjCommFun.Nvl(rsImage(0))
178             strImageData = strImageData & Replace(Replace(Trim(strTmp), vbCr, ""), vbLf, "")
            
180             If strImageData <> "" Then
182                 intLoop = 0
                
184                 If Val(Mid(strImageData, 1, 3)) >= 100 And Val(Mid(strImageData, 1, 3)) <= 227 And Mid(strImageData, 4, 1) = ";" Then
                
186                     blnPic = True
188                     If Mid(strImageData, 1, 3) >= 100 And Mid(strImageData, 1, 3) <= 107 Then
190                         strFile = strPath & "\zlLisPic" & lngͼ��ID & ".bmp"
192                     ElseIf Mid(strImageData, 1, 3) >= 110 And Mid(strImageData, 1, 3) <= 117 Then
194                         strFile = strPath & "\zlLisPic" & lngͼ��ID & ".jpg"
196                     ElseIf Mid(strImageData, 1, 3) >= 120 And Mid(strImageData, 1, 3) <= 127 Then
198                         strFile = strPath & "\zlLisPic" & lngͼ��ID & ".gif"
200                     ElseIf Mid(strImageData, 1, 3) >= 200 And Mid(strImageData, 1, 3) <= 227 Then
202                         If objFso.FolderExists(strPath & "\ZLLIS_ZIP") = False Then
204                             Call objFso.CreateFolder(strPath & "\ZLLIS_ZIP")
                            End If
206                         If objFso.FolderExists(strPath & "\ZLLIS_ZIP\" & lngͼ��ID) = False Then
208                             Call objFso.CreateFolder(strPath & "\ZLLIS_ZIP\" & lngͼ��ID)
                            End If
210                         strFile = App.Path & "\ZLLIS_ZIP\" & lngͼ��ID & "\ZLISPIC.ZIP"
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
                        If Not gblnNewLis Then
236                         strSql = "select Zl_FUN_Get����ͼ��(" & rsTmp("�걾ID") & ",'" & gobjCommFun.Nvl(rsTmp("ͼ������")) & "'," & intLoop & ") from dual "
                        Else
                            strSql = "select Zl_FUN_Get����ͼ��(" & rsTmp("�걾ID") & ",'" & gobjCommFun.Nvl(rsTmp("ͼ������")) & "'," & intLoop & ") from dual "
                        End If
238                     Set rsImage = gobjDatabase.OpenSQLRecord(strSql, "LoadImgData")
                    
240                     strTmp = gobjCommFun.Nvl(rsImage(0))
    
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
270                 strLocalFile = strPath & "\" & Split(strFtpPath, "/")(UBound(Split(strFtpPath, "/")))
272                 If objFso.FileExists(strLocalFile) Then objFso.DeleteFile strLocalFile
274                 strDownOk = DownFile(strFtpUser, strFtpPass, strFtpIP, strFtpPath, strLocalFile)
276                 If strDownOk = "" Then
278                     strImageData = intLayOut & ";" & strLocalFile
                    End If
                Else
                    ' ͼ�����ݣ���Ҫ�����ص��ı��ļ��ж�ȡ����
280                 strLocalFile = strPath & "\" & lngͼ��ID & "_" & strImageType & ".txt"
282                 If objFso.FileExists(strLocalFile) Then objFso.DeleteFile strLocalFile
284                 strDownOk = DownFile(strFtpUser, strFtpPass, strFtpIP, strFtpPath, strLocalFile)
286                 If strDownOk = "" Then
288                     Set objStream = objFso.OpenTextFile(strLocalFile, ForReading)
290                     Do Until objStream.AtEndOfLine
292                         strImageData = strImageData & objStream.ReadLine
                        Loop
294                     objStream.Close
296                     Set objStream = Nothing
298                     strImageData = Replace(Replace(Trim(strImageData), vbCr, ""), vbLf, "")
300                     strImageData = intLayOut & ";" & strImageData
                    End If
302                 If objFso.FileExists(strLocalFile) Then objFso.DeleteFile strLocalFile
                End If
            End If
        
304         If Len(strImageData) <> 0 Then
                '��ͼ������ͼ���ļ�
306             LoadImageData = DrawImgAndSaveFile(strImageType, strImageData, strPath & "\" & strFilename, intSaveType)
                
            End If
308         intLoop = 0
310         Do Until intLoop > 100
312             intLoop = intLoop + 1
314             If objFso.FileExists(strFile) Then
316                 WriteLog "LoadImageData", "��" & intLoop & "��ɾ��ԭʼ�ļ�" & strFile, ""
318                 Call objFso.DeleteFile(strFile)
                Else
320                 If strFile <> "" Then WriteLog "LoadImageData", "ԭʼ�ļ�" & strFile & "��ɾ��!", ""
                    Exit Do
                End If
            Loop
322         intLoop = 0
324         Do Until intLoop > 100
326             intLoop = intLoop + 1
328             If objFso.FileExists(strLocalFile) Then
330                 WriteLog "LoadImageData", "��" & intLoop & "��ɾ��FTP���ص�ԭʼ�ļ�" & strLocalFile, ""
332                 Call objFso.DeleteFile(strLocalFile)
                Else
334                 If strLocalFile <> "" Then WriteLog "LoadImageData", "FTP���ص�ԭʼ�ļ�" & strLocalFile & "��ɾ��!", ""
                    Exit Do
                End If
            Loop
336         strTmp = "": strImageData = ""
338         rsTmp.MoveNext
        Loop
        Set objStream = Nothing
        Set objFso = Nothing
        Exit Function
ErrH:
    Set objFso = Nothing
340 WriteLog "LoadImagedata", CStr(Erl()) & "�� ", err.Description
End Function

Public Function CheckGif(ByVal strFile As String) As Boolean
    '���GIF�ļ������Ƿ�����
    'GIF��ͷ��00 3B����
    Dim intFileNo   As Integer
    Dim lngFileSize As Long
    Dim arrEnd(2)   As Byte
    Dim arrTitle(3) As Byte
    Dim lngCount    As Long
    
On Error GoTo ErrH
'100 If Dir(strFile) <> "" Then
'102     intFileNo = FreeFile
'104     Open strFile For Binary Access Read As intFileNo
'106     lngFileSize = LOF(intFileNo)
'108     If lngFileSize > 0 Then
'110         Get intFileNo, , arrTitle
'112         Seek intFileNo, lngFileSize - 1
'114         Get intFileNo, , arrEnd
'        End If
'116     Close intFileNo
'
'118     If UCase(Chr(arrTitle(0)) & Chr(arrTitle(1)) & Chr(arrTitle(2))) = "GIF" And arrEnd(0) = 0 And arrEnd(1) = 59 Then
'120         CheckGif = True
'        End If
'    End If
    CheckGif = True
    Exit Function
ErrH:
122 WriteLog "CheckGif", CStr(Erl()) & "�� ", err.Description
    
End Function
