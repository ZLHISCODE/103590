VERSION 5.00
Begin VB.Form frmPcs 
   Caption         =   "子进程-不可见"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4800
   Icon            =   "frmPcs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   4800
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label lblPro 
      Caption         =   "当前进程:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblSta 
      AutoSize        =   -1  'True
      Caption         =   $"frmPcs.frx":6852
      Height          =   720
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   630
   End
End
Attribute VB_Name = "frmPcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSave As Boolean
Private mclsFtp As New clsFtp   'FTP类

Private Sub Form_Load()
    Dim strCmd As String
    Dim strServerName As String, strUserName As String
    Dim strUserPwd As String, strPro As String
    Dim intUserPosition As Integer, intPwdPosition As Integer
    Dim intServerPosition As Integer
    Dim strFtp As String

    On Error GoTo errH
    '主界面界面将登录信息传入,格式:路径 用户名 密码
    '"Path\zlLisPic2FtpSub.exe zlUserName=" & gstrUserName & "zlPassword=HIS" & "zlServer=" & gstrServer
    strCmd = Command
    If strCmd <> "" Then
        intUserPosition = InStr(1, strCmd, "zlUserName=") + Len("zlUserName=")
        intPwdPosition = InStr(1, strCmd, "zlPassword=") + Len("zlPassword=")
        intServerPosition = InStr(1, strCmd, "zlServer=") + Len("zlServer=")
        
        strUserName = Mid(Left(strCmd, InStr(1, strCmd, "zlPassword=") - 1), intUserPosition)
        strUserPwd = Mid(Left(strCmd, InStr(1, strCmd, "zlServer=") - 1), intPwdPosition)
        strServerName = Mid(strCmd, intServerPosition)
        
        '数据库连接成功,就执行操作,否则写入错误
        If OraDataOpen(strServerName, strUserName, strUserPwd) Then
            '相关操作从注册表中读取,格式: 转出类型(1-FTP 2-保存本地);数据来源(1-旧版LIS 2-新版LIS);进程号;开始时间;结束时间;临时路径;FTP路径
            strPro = GetSetting("LIS图片转出", "转出进度", "进程设置")
        
            '直接上传至FTP模式下,获取FTP的用户\密码\IP\路径
            If Split(strPro, ";")(0) = 1 Then
                strFtp = GetSetting("LIS图片转出", "转出设置", "FTP路径")
                mclsFtp.FuncFtpConnect Split(strFtp, ";")(2), Split(strFtp, ";")(0), Split(strFtp, ";")(1)
                mclsFtp.FuncChangeDir Split(strFtp, ";")(3)
            End If
            
            ImgUpload Split(strPro, ";")(0), Split(strPro, ";")(1), Split(strPro, ";")(2), Split(strPro, ";")(3), Split(strPro, ";")(4), Split(strPro, ";")(5), Split(strPro, ";")(6), Split(strPro, ";")(7)
                

        Else
            SaveSetting "LIS图片转出", "转出进度", "转出错误", "数据库连接时发生错误。"
            WriteErrLog "数据库连接时发生错误。"
        End If
    End If
    
    Unload Me   '执行完操作,退出
    Exit Sub
    
errH:
    SaveSetting "LIS图片转出", "转出进度", "转出错误", "初始化进程时发生错误。"
    WriteErrLog Err.Description
    Unload frmPcs
End Sub


Private Sub ImgUpload(ByVal intType As Integer, ByVal intSource, ByVal intProc As Integer, ByVal intProcNum As Integer, ByVal strStart As String, ByVal strEnd As String, ByVal strTmpPath As String, ByVal strFtpPath As String)
    '功能:以进程号作为循环每天的图片的Step,将图片上传到FTP服务器或者转存到本地
    '参数说明: intType 转出类型 1-同步 2-异步  ; intSource 1-旧版LIS 2-新版LIS ;intProc 进程号 ;intProcNum 进程数; strStart-转出开始时间 ;strEnd -转出结束时间; strTmpPath-临时目录; strFtpPath -FTP目录
    Dim DateMin As Date, DateMax As Date, iDays As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsExp As ADODB.Recordset
    Dim blnDo As Boolean, strDes As String, strFile As String
    Dim strSrc As String, iDay As Integer, str图像位置 As String
    Dim DateS As Date, DateE As Date
    Dim j As Integer, i As Long
    
    On Error GoTo errH
    '获取记录的开始时间和结束时间
    DateMin = CDate(strStart)
    DateMax = CDate(strEnd)

    iDays = DateDiff("d", DateMin, DateMax)
    strSrc = strTmpPath & "\"  '本地目录
    
    If intSource = 1 Then
        strSQL = "Select id From 检验图像结果_EXP_TEMP"
        Set rsExp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetExp")
    Else
        strSQL = "Select id From 检验报告图像_EXP_TEMP"
        Set rsExp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetExp")
    End If
    
    '循环每天的记录,获取图片
    For iDay = intProc - 1 To iDays Step intProcNum
        
        DateS = Format(DateAdd("d", -iDay, DateMax), "yyyy-MM-dd 00:00:00")
        DateE = Format(DateAdd("d", -iDay, DateMax), "yyyy-MM-dd 23:59:59")
        
        If GetSetting("LIS图片转出", "转出进度", "转出错误") <> "" Then '如果有其他进程发生错误,就终止
            WriteErrLog "有进程意外结束,转出终止"
            Exit Sub
        End If
        
        If CheckProcExist("zllispic2ftp.exe") = 0 Then   '主进程被终止
            If CheckProcExist("zlSvrStudio.exe") = 0 Then
                WriteErrLog "主进程意外结束,转出终止"
                Exit Sub
            End If
        End If
        
        If intSource = 1 Then
            strSQL = "Select /*+ rule */ b.ID,b.标本id,a.仪器ID,b.图像类型,b.图像点 " & vbNewLine & _
                            " From 检验标本记录 A, 检验图像结果 B " & vbNewLine & _
                            " Where a.核收时间 Between [1]  And  [2] And a.审核人 Is Not Null And a.Id = b.标本id And b.图像位置 Is Null and b.图像点 Is Not Null"
        Else
            strSQL = "Select /*+ rule */ b.ID,b.标本id,a.仪器ID,b.图像类型,b.图像点 " & vbNewLine & _
                            " From 检验报告记录 A, 检验报告图像 B " & vbNewLine & _
                            " Where a.核收时间 Between [1]  And  [2] And a.审核人 Is Not Null And a.Id = b.标本id And b.图像位置 Is Null and b.图像点 Is Not Null"
        End If
        
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetImg", CDate(DateS), CDate(DateE))
        
        With rsTmp
            If .RecordCount > 0 Then
                Do While Not .EOF
                    '判断当前记录是否保存在导出表中
                    blnDo = True
                    If rsExp.RecordCount > 0 Then
                        rsExp.Filter = "id=" & !id
                        blnDo = rsExp.RecordCount = 0
                        rsExp.Filter = 0
                    End If
                    
                    If blnDo Then
                        strFile = !标本ID & "_" & !图像类型     '图片名称
                        str图像位置 = "110;" & "/" & strFtpPath & "/Dev_" & !仪器ID & "/" & Format(DateS, "yyyyMM") & "/" & strFile & ".JPG"   '数据库中保存的名称
                        
                        If intType = 1 Then  '选择的模式是直接上传到FTP
                            strDes = "Dev_" & !仪器ID & "/" & Format(DateS, "yyyyMM")  'FTP虚拟路径
                            If ImgSaveAsJpg(!图像类型, !图像点, strTmpPath, strFile) Then  '图片转存本地
                                '图片上传
                                '路径创建  : Dev_仪器ID/时间/
                                '每次创建都需要保证当前路径是根路径
                                If i > 0 Then
                                    mclsFtp.FuncFtpCommand strFtpPath, "cdup"
                                    mclsFtp.FuncFtpCommand strFtpPath, "cdup"
                                End If
                                mclsFtp.FuncFtpCommand strFtpPath, "mkd " & "Dev_" & !仪器ID   '这里要先创建首级目录
                                mclsFtp.FuncFtpCommand strFtpPath, "mkd " & strDes
                                
                                If mclsFtp.FuncUploadFile(strDes, strSrc & strFile & ".JPG", strFile & ".JPG") <> 0 Then
                                    '如果上传失败,就删除该图片,下次重新上传
                                    If gobjFile.FileExists(strSrc & strFile & ".JPG") Then
                                        Kill strSrc & strFile & ".JPG"
                                    End If
                                    WriteErrLog "上传图片时发生错误,转存终止,错误图像ID为:" & !id
                                    Exit Sub
                                Else
                                    If intSource = 1 Then
                                        Call ExecuteProcedure("Zl_检验图像结果_Temp_Insert(" & !id & "," & !标本ID & ",'" & !图像类型 & "','" & str图像位置 & "')", Me.Caption)
                                    Else
                                        Call ExecuteProcedure("Zl_检验报告图像_Temp_Insert(" & !id & "," & !标本ID & ",'" & !图像类型 & "','" & str图像位置 & "')", Me.Caption)
                                    End If
                                End If
                                If gobjFile.FileExists(strSrc & strFile & ".JPG") Then
                                    Kill strSrc & strFile & ".JPG"
                                End If
                            Else
                                WriteErrLog "保存图片时发生错误,转存终止,错误图像ID为:" & !id
                                Exit Sub
                            End If
                            i = i + 1
                        Else
                            '直接建立文件夹保存在本地
                            If Not gobjFile.FolderExists(strTmpPath & "\" & "Dev_" & !仪器ID) Then
                                gobjFile.CreateFolder strTmpPath & "\" & "Dev_" & !仪器ID
                            End If
                            If Not gobjFile.FolderExists(strTmpPath & "\" & "Dev_" & !仪器ID & "\" & Format(DateS, "yyyyMM")) Then
                                gobjFile.CreateFolder strTmpPath & "\" & "Dev_" & !仪器ID & "\" & Format(DateS, "yyyyMM")
                            End If
                            strDes = strTmpPath & "\Dev_" & !仪器ID & "\" & Format(DateS, "yyyyMM")
                            If Not ImgSaveAsJpg(!图像类型, !图像点, strDes, strFile) Then
                                WriteErrLog "保存图片时发生错误,转存终止,错误图像ID为:" & !id
                                Exit Sub
                            Else
                                If intSource = 1 Then
                                    Call ExecuteProcedure("Zl_检验图像结果_Temp_Insert(" & !id & "," & !标本ID & ",'" & !图像类型 & "','" & str图像位置 & "')", Me.Caption)
                                Else
                                    Call ExecuteProcedure("Zl_检验报告图像_Temp_Insert(" & !id & "," & !标本ID & ",'" & !图像类型 & "','" & str图像位置 & "')", Me.Caption)
                                End If
                            End If
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End With
        j = j + 1
        SaveSetting "LIS图片转出", "转出进度", "进程" & intProc, j '保存处理进度
    Next
    SaveSetting "LIS图片转出", "转出进度", "进程" & intProc, j & ";" '处理完成,在进度后面加上一个 ";" 标识当前线程转出完成
    Exit Sub
errH:
    SaveSetting "LIS图片转出", "转出进度", "转出错误", "转存图片时发生错误。"
    WriteErrLog Err.Description & " 错误图片:" & strFile
    Unload frmPcs
End Sub


Private Function ImgSaveAsJpg(ByVal strType As String, ByVal strImgStrem As String, ByVal strPath As String, strFileName As String) As Boolean
    '功能:根据传入的图像点转换成图片保存本地,成功返回True
    '参数:strImgStrem-图像点流  strFileName-图片名称    strPath-位置
    Dim strFile As String
    Dim aryChunk() As Byte, intLayOut As Integer
    Dim lngFileNum As Long, lngBound As Long
    Dim frmObj As frmLisChart
    
    On Error GoTo errH
            
    '创建目录
    If Not gobjFile.FolderExists(strPath) Then
        gobjFile.CreateFolder strPath
    End If
    '判断图片类型
    If Val(Mid(strImgStrem, 1, 3)) >= 100 And Val(Mid(strImgStrem, 1, 3)) <= 227 And Mid(strImgStrem, 4, 1) = ";" Then
        '数据库中保存的是图片数据
        If Mid(strImgStrem, 1, 3) >= 100 And Mid(strImgStrem, 1, 3) <= 107 Then
            strFile = strPath & "\" & strFileName & "_Tmp.bmp"
        ElseIf Mid(strImgStrem, 1, 3) >= 110 And Mid(strImgStrem, 1, 3) <= 117 Then
            strFile = strPath & "\" & strFileName & "_Tmp.jpg"
        ElseIf Mid(strImgStrem, 1, 3) >= 120 And Mid(strImgStrem, 1, 3) <= 127 Then
            strFile = strPath & "\" & strFileName & "_Tmp.gif"
        ElseIf Mid(strImgStrem, 1, 3) >= 200 And Mid(strImgStrem, 1, 3) <= 227 Then
            If gobjFile.FolderExists(strPath & "\ZLLIS_ZIP") = False Then
                gobjFile.CreateFolder strPath & "\ZLLIS_ZIP"
            End If
            If gobjFile.FolderExists(strPath & "\ZLLIS_ZIP\" & strFileName) = False Then
                gobjFile.CreateFolder strPath & "\ZLLIS_ZIP\" & strFileName
            End If
            strFile = strPath & "\ZLLIS_ZIP\" & strFileName & "\ZLISPIC.ZIP"
        End If
    
        '如果数据库中保存的是图片数据,就直接保存本地,再进行绘图
        intLayOut = Val(Mid(strImgStrem, 1, 3))
        strImgStrem = Mid(Replace(Replace(Trim(strImgStrem), vbCr, ""), vbLf, ""), 5)
        If gobjFile.FileExists(strFile) Then
            Kill strFile '如果文件存在,就删除
        End If
    
        '创建文件,转换后保存本地
        lngFileNum = FreeFile
        Open strFile For Binary As lngFileNum
        ReDim aryChunk(Len(strImgStrem) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strImgStrem, lngBound * 2 + 1, 2))
        Next
        Put lngFileNum, , aryChunk()
        Close lngFileNum
        
        strImgStrem = intLayOut & ";" & strFile
    Else
        strImgStrem = Replace(Replace(Trim(strImgStrem), vbCr, ""), vbLf, "")
    End If
    
    '利用chart2D绘图 转换成JPG
    Set frmObj = New frmLisChart
    If frmObj.DrawImg(strType, strImgStrem, strPath & "\" & strFileName & ".JPG", 1) Then
        ImgSaveAsJpg = True
    End If
    Unload frmObj
    
    Exit Function
errH:
    SaveSetting "LIS图片转出", "转出进度", "转出错误", "图片保存本地发生错误。"
    WriteErrLog Err.Description & " 错误图片:" & strFileName
    Unload Me
End Function
