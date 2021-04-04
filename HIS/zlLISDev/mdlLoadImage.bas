Attribute VB_Name = "mdlLoadImage"
Option Explicit
'获取图片数据功能模块

Public gcnOracle As New Connection                                  '公共连接
Public grsParas As ADODB.Recordset                                  '系统参数表缓存
Public grsUserParas As ADODB.Recordset                              '系统参数表缓存
Public gComLib As Object                                            '公共部件
Public gblnInit As Boolean                                          '是否初始化
Public gstrSql  As String
Public glngSys As Long                                              '系统号
Public gstrComputerName As String

Public Function DrawImgAndSaveFile(ByVal strType As String, ByVal strData As String, ByVal strFileName As String, Optional ByVal intSaveType As Integer) As Boolean
    '根据传入的参数绘图，并保存为指定文件
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
            strFtpPara = GetPara("FTP设置", glngSys, 1208, "")
        Else
            strFtpPara = GetPara("FTP设置", glngSys, 2500, "")
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
    WriteLog "FunFtpSet", CStr(Erl()) & "行 ", Err.Description
End Function

Public Function LoadImageDataTwo(ByVal strPath As String, ByVal lngID As Long, Optional ByVal intSaveType As Integer, Optional ByVal intVer As Integer, _
                                 Optional ByVal strFileName As String) As Boolean
        '从数据库读取一个图形数据，绘制后保存到指定的路径下。
        '入参：
        '   strPath 路径
        '   lngID   检验图像结果的ID
        '   intSaveType :只存的图片类型，0-cht(默认) 1-jpg,2-png
        '   intVer      :版本
        '--如果有的话, 删除原来的临时图形文件
        
        Dim rsTmp As New ADODB.Recordset, rsImage As New ADODB.Recordset
        Dim rsItem As New ADODB.Recordset
        Dim strImageType As String
        Dim strImageData As String
        Dim DrawIndex As Integer
        Dim intLoop As Integer
        Dim lngStart As Long
        Dim strTmp As String
        Dim strSQL  As String
    
        Dim blnPic As Boolean '是否图片格式
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
           strSQL = "select 标本id,图像类型 from 检验图像结果 where id = [1] "
        Else
           strSQL = "select 标本id,图像类型 from 检验报告图像 where id = [1] "
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "zlLISDev.LoadImageData", lngID)

160     If rsTmp.EOF = True Then Exit Function
        
162     Do Until rsTmp.EOF
            strImageType = Trim("" & rsTmp("图像类型"))
            '- 图像存在数据库中，按原来的方式处理
            If intVer = 0 Then
                gstrSql = "select Zl_FUN_Get检验图像(" & rsTmp("标本id") & ",'" & Nvl(rsTmp("图像类型")) & "',0) from dual "
            Else
                gstrSql = "select Zl_FUN_Get报告图像(" & rsTmp("标本id") & ",'" & Nvl(rsTmp("图像类型")) & "',0) from dual "
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
                    '-------保存为图片文件
                Do While strTmp <> ""
                    intLoop = intLoop + 1
                    If intVer = 0 Then
                        gstrSql = "select Zl_FUN_Get检验图像(" & rsTmp("标本id") & ",'" & Nvl(rsTmp("图像类型")) & "'," & intLoop & ") from dual "
                    Else
                        gstrSql = "select Zl_FUN_Get报告图像(" & rsTmp("标本id") & ",'" & Nvl(rsTmp("图像类型")) & "'," & intLoop & ") from dual "
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
                        '图形数据
                        strImageData = strImageData & Replace(Replace(Trim(strTmp), vbCr, ""), vbLf, "")
                    End If
                Loop
                
                If blnPic Then
                    strImageData = intLayOut & ";" & strFile
                    Close lngFileNum
                End If
            End If
        
304         If Len(strImageData) <> 0 Then
                '画图并产生图形文件
306             LoadImageDataTwo = DrawImgAndSaveFile(strImageType, strImageData, strPath & "\" & strFileName, intSaveType)
                
            End If
308         intLoop = 0
310         Do Until intLoop > 100
312             intLoop = intLoop + 1
314             If gobjFSO.FileExists(strFile) Then
316                 WriteLog "LoadImageData", "第" & intLoop & "次删除原始文件" & strFile, ""
318                 Call gobjFSO.DeleteFile(strFile)
                Else
320                 If strFile <> "" Then WriteLog "LoadImageData", "原始文件" & strFile & "已删除!", ""
                    Exit Do
                End If
            Loop
'322         intLoop = 0
'324         Do Until intLoop > 100
'326             intLoop = intLoop + 1
'328             If gobjFSO.FileExists(strLocalFile) Then
'330                 WriteLog "LoadImageData", "第" & intLoop & "次删除FTP下载的原始文件" & strLocalFile, ""
'332                 Call gobjFSO.DeleteFile(strLocalFile)
'                Else
'334                 If strLocalFile <> "" Then WriteLog "LoadImageData", "FTP下载的原始文件" & strLocalFile & "已删除!", ""
'                    Exit Do
'                End If
'            Loop
336         strTmp = "": strImageData = ""
338         rsTmp.MoveNext
        Loop
        Exit Function
errHandle:
340     WriteLog "LoadImagedata", CStr(Erl()) & "行 ", Err.Description
End Function

Public Function LoadImageData(ByVal strPath As String, ByVal lngID As Long, Optional ByVal intSaveType As Integer, Optional ByVal intVer As Integer, Optional ByVal strFileName As String) As Boolean
        '从数据库读取一个图形数据，绘制后保存到指定的路径下。
        '入参：
        '   strPath 路径
        '   lngID   检验图像结果的ID
        '   intSaveType :只存的图片类型，0-cht(默认) 1-jpg,2-png
        '   intVer      :版本
        '--如果有的话, 删除原来的临时图形文件
        
        Dim rsTmp As New ADODB.Recordset, rsImage As New ADODB.Recordset
        Dim rsItem As New ADODB.Recordset
        Dim strImageType As String
        Dim strImageData As String
        Dim DrawIndex As Integer
        Dim intLoop As Integer
        Dim lngStart As Long
        Dim strTmp As String
        Dim strSQL  As String
    
        Dim blnPic As Boolean '是否图片格式
        Dim lngFileNum As Long, lngCount As Long, lngBound As Long
        Dim aryChunk() As Byte, strFile As String
        Dim intLayOut As Integer
        Dim objPic As New frmChartPic
        Dim killFile As String
    
        Dim blnFtp As Boolean       'FTP是否可用
        Static strFtpPara As String       '保存FTP参数
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
    
        'FTP连接检查，有效则可以按FTP方式取图片
114     blnFtp = False
116     If strFtpPara = "" Then
118         If intVer = 0 Then
120             strFtpPara = GetPara("FTP设置", glngSys, 1208, "")
            Else
122             strFtpPara = GetPara("FTP设置", glngSys, 2500, "")
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
146             strSQL = "select 标本id,图像类型,图像位置 from 检验图像结果 where id = [1]"
             Else
148             strSQL = "select 标本id,图像类型,图像位置 from 检验报告图像 where id = [1]"
             End If
150          Set rsTmp = OpenSQLRecord(strSQL, "zlLISDev.LoadImageData", lngID)
        Else
152          If intVer = 0 Then
154             strSQL = "select 标本id,图像类型 from 检验图像结果 where id = [1] "
             Else
156             strSQL = "select 标本id,图像类型 from 检验报告图像 where id = [1] "
             End If
158          Set rsTmp = OpenSQLRecord(strSQL, "zlLISDev.LoadImageData", lngID)
             
        End If
160     If rsTmp.EOF = True Then Exit Function

    
        
162     Do Until rsTmp.EOF
164         strImageType = Trim("" & rsTmp("图像类型"))
166         strFtpPath = ""
168         If blnFtp Then strFtpPath = Trim("" & rsTmp!图像位置)
170         If InStr(strFtpPath, ";") <= 0 Or Not blnFtp Then
                '- 图像存在数据库中，按原来的方式处理
                If intVer = 0 Then
172                 gstrSql = "select Zl_FUN_Get检验图像(" & rsTmp("标本id") & ",'" & Nvl(rsTmp("图像类型")) & "',0) from dual "
                Else
                    gstrSql = "select Zl_FUN_Get报告图像(" & rsTmp("标本id") & ",'" & Nvl(rsTmp("图像类型")) & "',0) from dual "
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
                    '-------保存为图片文件
232                 Do While strTmp <> ""
234                     intLoop = intLoop + 1
                        If intVer = 0 Then
236                         gstrSql = "select Zl_FUN_Get检验图像(" & rsTmp("标本id") & ",'" & Nvl(rsTmp("图像类型")) & "'," & intLoop & ") from dual "
                        Else
                            gstrSql = "select Zl_FUN_Get报告图像(" & rsTmp("标本id") & ",'" & Nvl(rsTmp("图像类型")) & "'," & intLoop & ") from dual "
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
                            '图形数据
254                         strImageData = strImageData & Replace(Replace(Trim(strTmp), vbCr, ""), vbLf, "")
                        End If
                    Loop
                
256                 If blnPic Then
258                     strImageData = intLayOut & ";" & strFile
260                     Close lngFileNum
                    End If
                End If
            Else
                '图像存在FTP中，从FTP中取数据
                '图像位置的数据格式为：图像格式;FTP文件路径
            
262             intLayOut = Val(Split(strFtpPath, ";")(0))
264             strFtpPath = Trim(Split(strFtpPath, ";")(1))
266             strImageData = ""
268             If intLayOut >= 100 And intLayOut <= 227 Then
                    ' 图片文件，直接下载到本地
270                 strLocalFile = strPath & "\zlLisPic" & Split(strFtpPath, "/")(UBound(Split(strFtpPath, "/")))
272                 If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
274                 strDownOk = DownFile(strFtpUser, strFtpPass, strFtpIP, strFtpPath, strLocalFile)
276                 If strDownOk = "" Then
278                     strImageData = intLayOut & ";" & strLocalFile
                    End If
                Else
                    ' 图形数据，需要从下载的文本文件中读取数据
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
                '画图并产生图形文件
306             LoadImageData = DrawImgAndSaveFile(strImageType, strImageData, strPath & "\" & strFileName, intSaveType)
            End If
308         intLoop = 0
310         Do Until intLoop > 100
312             intLoop = intLoop + 1
314             If gobjFSO.FileExists(strFile) Then
316                 WriteLog "LoadImageData", "第" & intLoop & "次删除原始文件" & strFile, ""
318                 Call gobjFSO.DeleteFile(strFile)
                Else
320                 If strFile <> "" Then WriteLog "LoadImageData", "原始文件" & strFile & "已删除!", ""
                    Exit Do
                End If
            Loop
322         intLoop = 0
324         Do Until intLoop > 100
326             intLoop = intLoop + 1
328             If gobjFSO.FileExists(strLocalFile) Then
330                 WriteLog "LoadImageData", "第" & intLoop & "次删除FTP下载的原始文件" & strLocalFile, ""
332                 Call gobjFSO.DeleteFile(strLocalFile)
                Else
334                 If strLocalFile <> "" Then WriteLog "LoadImageData", "FTP下载的原始文件" & strLocalFile & "已删除!", ""
                    Exit Do
                End If
            Loop
336         strTmp = "": strImageData = ""
338         rsTmp.MoveNext
        Loop
        Exit Function
errHandle:
340     WriteLog "LoadImagedata", CStr(Erl()) & "行 ", Err.Description
End Function


Public Function CheckGif(ByVal strFile As String) As Boolean
    '检查GIF文件数据是否完整
    'GIF开头，00 3B结束
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
        '判断是否自己画的图片保存的。使用控件保存图片后，
        '保存gif格式图片可能只是后缀保存为gif格式，实际保存的图片还是bmp图片。
        If UCase(Chr(arrTitle(0)) & Chr(arrTitle(1))) = "BM" Then
            CheckGif = True
        End If
    End If
    Exit Function
hErr:
122     WriteLog "CheckGif", CStr(Erl()) & "行 ", Err.Description
    
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'               因为使用绑定变量,对带"'"的字符参数,不需要使用"''"形式。
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strlog As String, varValue As Variant
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        If lngRight = 0 Then Exit Do
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '替换为"?"参数
    strlog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            strlog = Replace(strlog, "[" & i & "]", varValue)
        Case "String" '字符
            strlog = Replace(strlog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strlog = Replace(strlog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '清除原有参数:不然不能重复执行
'    cmdData.CommandText = "" '不为空有时清除参数出错
'    Do While cmdData.Parameters.Count > 0
'        cmdData.Parameters.Delete 0
'    Loop
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '数组
            '这种方式可用于一些IN子句或Union语句
            '表示同一个参数的多个值,参数号不可与其它数组的参数号交叉,且要保证数组的值个数够用
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strlog = Replace(strlog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '字符
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strlog = Replace(strlog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '日期
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strlog = Replace(strlog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '该参数在数组中用到第几个值了
        End Select
    Next

    '执行返回记录集
    'If cmdData.ActiveConnection Is Nothing Then
    
     Set cmdData.ActiveConnection = gcnOracle '这句比较慢(这句执行1000次约0.5x秒)
    'End If
    cmdData.CommandText = strSQL
    
'    Call gobjComLib.SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Set OpenSQLRecord.ActiveConnection = Nothing
'    Call gobjComLib.SQLTest
End Function

Public Sub ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'功能：执行过程语句,并自动对过程参数进行绑定变量处理
'参数：strSQL=过程语句,可能带参数,形如"过程名(参数1,参数2,...)"。
'说明：以下几种情况过程参数不使用绑定变量,仍用老的调用方法：
'  1.参数部份是表达式,这时程序无法处理绑定变量类型和值,如"过程名(参数1,100.12*0.15,...)"
'  2.中间没有传入明确的可选参数,这时程序无法处理绑定变量类型和值,如"过程名(参数1, , ,参数3,...)"
'  3.因为该过程是自动处理,不是一定使用绑定变量,对带"'"的字符参数,仍要使用"''"形式。
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    
    If Right(Trim(strSQL), 1) = ")" Then
        '清除原有参数:不然不能重复执行
'        cmdData.CommandText = "" '不为空有时清除参数出错
'        Do While cmdData.Parameters.Count > 0
'            cmdData.Parameters.Delete 0
'        Loop
        
        '执行的过程名
        strTemp = Trim(strSQL)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        '执行过程参数
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For i = 1 To Len(strTemp)
            '是否在字符串内，以及表达式的括号内
            If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                strPar = Trim(strPar)
                With cmdData
                    If IsNumeric(strPar) Then '数字
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarNumeric, adParamInput, 30, strPar)
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '字符串
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        
                        'Oracle连接符运算:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '双"''"的绑定变量处理
                        If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                        
                        '电子病历处理LOB时，如果用绑定变量转换为RAW时超过2000个字符要用adLongVarChar
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
                        If intMax <= 2000 Then
                            intMax = IIf(intMax <= 200, 200, 2000)
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarChar, adParamInput, intMax, strPar)
                        Else
                            If intMax < 4000 Then intMax = 4000
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adLongVarChar, adParamInput, intMax, strPar)
                        End If
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '日期
                        strPar = Split(strPar, "(")(1)
                        strPar = Trim(Split(strPar, ",")(0))
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        If strPar = "" Then
                            'NULL值当成数字处理可兼容其他类型
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(strPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adDBTimeStamp, adParamInput, , CDate(strPar))
                        End If
                    ElseIf UCase(strPar) = "SYSDATE" Then '日期
                        If datCur = CDate(0) Then datCur = Currentdate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then 'NULL值当成字符处理可兼容其他类型
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarChar, adParamInput, 200, Null)
                    ElseIf strPar = "" Then '可选参数当成NULL处理可能改变了缺省值:因此可选参数不能写在中间
                        GoTo NoneVarLine
                    Else '可能是其他复杂的表达式，无法处理
                        GoTo NoneVarLine
                    End If
                End With
                
                strPar = ""
            Else
                strPar = strPar & Mid(strTemp, i, 1)
            End If
        Next
        
        '程序员调用过程时书写错误
        If blnStr Or intBra <> 0 Then
            Err.Raise -2147483645, , "调用 Oracle 过程""" & strProc & """时，引号或括号书写不匹配。原始语句如下：" & vbCrLf & vbCrLf & strSQL
            Exit Sub
        End If
        
        '补充?号
        strTemp = ""
        For i = 1 To cmdData.Parameters.count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        
        
        '执行过程
        'If cmdData.ActiveConnection Is Nothing Then
            Set cmdData.ActiveConnection = gcnOracle '这句比较慢
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
    
    '说明：为了兼容新连接方式
    '1.新连接用adCmdStoredProc方式在8i下面有问题
    '2.新连接如果不使用{},则即使过程没有参数也要加()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    
    gcnOracle.Execute strSQL, , adCmdText
'    Call gobjComLib.SQLTest
End Sub
Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional inttype As Integer) As String
'功能：读取指定的参数值
'参数：varPara=参数号或参数名，以数字或字符类型传入区分
'      lngSys=使用该参数的系统编号，如100
'      lngModual=使用该参数的模块号，如1230
'      strDefault=当数据库中没有该参数时使用的缺省值(注意不是为空时)
'      blnNotCache=是否不从缓存中读取
'      arrControl=控件数组，如Array(Me.Text1, Me.CheckBox1)，用于函数内部自动处理对应控件的显示颜色，是否禁止设置。
'      blnSetup=调用模块是否有参数设置权限
'      intType=返回参数，返回参数类型
'返回：参数值，字符串形式
    Dim strSQL As String, i As Integer
    Dim blnNew As Boolean, blnEnabled As Boolean
    Dim strDBUser As String
    
    
    strDBUser = GetUserDB()
    On Error GoTo errH
    
    inttype = 0
    
    '第一次加载参数缓存
    If grsParas Is Nothing Then
        blnNew = True
    ElseIf grsParas.State = 0 Then
        blnNew = True
    End If
    If blnNew Then
        strSQL = "Select ID,Nvl(系统,0) as 系统,Nvl(模块,0) as 模块,Nvl(私有,0) as 私有,Nvl(本机,0) as 本机,Nvl(授权,0) as 授权,参数号,参数名," & _
            " Nvl(参数值,缺省值) as 参数值,[1] as 用户名,[2] as 机器名 From zlParameters"
        Set grsParas = New ADODB.Recordset
        Set grsParas = OpenSQLRecord(strSQL, "GetPara", strDBUser, gstrComputerName)
        
        strSQL = _
            " Select 参数ID,Nvl(用户名,'NullUser') as 用户名,Nvl(机器名,'NullMachine') as 机器名,参数值 From zlUserParas Where 用户名=[1]" & _
            " Union" & _
            " Select 参数ID,Nvl(用户名,'NullUser') as 用户名,Nvl(机器名,'NullMachine') as 机器名,参数值 From zlUserParas Where 机器名=[2]"
        Set grsUserParas = New ADODB.Recordset
        Set grsUserParas = OpenSQLRecord(strSQL, "GetPara", strDBUser, gstrComputerName)
    End If
    
    '使用缓存
    If TypeName(varPara) = "String" Then
        grsParas.Filter = "参数名='" & CStr(varPara) & "' And 模块=" & lngModual & " And 系统=" & lngSys
    Else
        grsParas.Filter = "参数号=" & Val(varPara) & " And 模块=" & lngModual & " And 系统=" & lngSys
    End If
    If Not grsParas.EOF Then
        '获取参数值
        If grsParas!私有 = 1 Or grsParas!本机 = 1 Then
            grsUserParas.Filter = "参数ID=" & grsParas!id & _
                IIf(grsParas!私有 = 1, " And 用户名='" & grsParas!用户名 & "'", " And 用户名='NullUser'") & _
                IIf(grsParas!本机 = 1, " And 机器名='" & grsParas!机器名 & "'", " And 机器名='NullMachine'")
            If Not grsUserParas.EOF Then
                GetPara = Nvl(grsUserParas!参数值, strDefault)
            Else
                GetPara = Nvl(grsParas!参数值, strDefault)
            End If
        Else
            GetPara = Nvl(grsParas!参数值, strDefault)
        End If
        
        '返回参数类型：1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
        If grsParas!系统 <> 0 And grsParas!模块 = 0 And grsParas!私有 = 0 And grsParas!本机 = 0 Then
            inttype = 1
        ElseIf grsParas!模块 = 0 And grsParas!私有 = 1 And grsParas!本机 = 0 Then
            inttype = 2
        ElseIf grsParas!系统 <> 0 And grsParas!模块 <> 0 And grsParas!私有 = 0 And grsParas!本机 = 0 Then
            inttype = 3
        ElseIf grsParas!系统 <> 0 And grsParas!模块 <> 0 And grsParas!私有 = 1 And grsParas!本机 = 0 Then
            inttype = 4
        ElseIf grsParas!系统 <> 0 And grsParas!模块 <> 0 And grsParas!私有 = 0 And grsParas!本机 = 1 Then
            inttype = IIf(grsParas!授权 = 1, 15, 5)
        ElseIf grsParas!系统 <> 0 And grsParas!模块 <> 0 And grsParas!私有 = 1 And grsParas!本机 = 1 Then
            inttype = 6
        End If
        
        '处理对应的控件颜色，可控状态
        If IsArray(arrControl) And (inttype = 3 Or (inttype Mod 10) = 5) Then
            blnEnabled = Not ((inttype = 3 Or (inttype Mod 10) = 5 And grsParas!授权 = 1) And Not blnSetup)
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
    '功能：提取服务器上当前日期
    '参数：
    '返回：由于Oracle日期格式的问题，所以
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
    '--  功能:根据调试标志,写日志到当前目录
    '------------------------------------------------------

    '以下变量用于记录调用接口的入参
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As textStream
    Dim objFileSystem As New FileSystemObject

    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    strFileName = App.Path & "\LisDev_" & Format(date, "yyyyMMdd") & ".LOG"

    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (strDate & ":" & StrInput)
    objStream.Close
    Set objStream = Nothing
End Sub
