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

Private mblnIsLoadReportImage As Boolean '是根据报告图象字段加载的报告图
Private mslQueryLevel As tQueryLevel      '图像显示级别
Private mtyFileLoadType As FileLoadType

'文件加载方式
Property Get ImgLoadType() As FileLoadType
    ImgLoadType = mtyFileLoadType
End Property

Property Let ImgLoadType(value As FileLoadType)
    mtyFileLoadType = value
End Property


'图像显示级别
Property Get QueryLevel() As tQueryLevel
    QueryLevel = mslQueryLevel
End Property

Property Let QueryLevel(value As tQueryLevel)
    mslQueryLevel = value
End Property

Public Function DownloadImages(ByRef arrImages() As String, strQueryValue As String, ByVal lngStartIndex As Long, ByVal lngEndIndex As Long, _
                blnOnlyLoadReportImage As Boolean, blnMoved As Boolean, _
                Optional blnQueryTmpRecord As Boolean) As Boolean
'lngAdviceId：查询条件值
'blnOnlyLoadReportImage：为True时加载 报告图像 字段中的报告图,反之加载所有报告图
'lngStartIndex：第一张图像索引
'lngEndIndex：最后一张图像索引
'blnMoved：图像是否被转存
'slQueryLevel：图像下载级别
'blnQueryTmpRecord：是否临时图像
    Dim rsData As Recordset

    If strQueryValue = "0" Then Exit Function
    If QueryLevel = slAdvice And Not IsNumeric(Val(strQueryValue)) Then Exit Function

    If blnOnlyLoadReportImage Then
        '根据 影像检查记录.报告图像 字段中的值下载，如果为空， 则下载所有报告图像
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
        DownloadImages = LoadViewImageToFaceWithNormal(rsData, arrImages)     '使用原始模式加载
    Else
        DownloadImages = LoadViewImageToFaceWithService(rsData, arrImages)   '使用ZLPacsServerCenter服务,后台加载
    End If

End Function

Private Function GetImageRptData(ByVal lngOrderID As Long, ByVal lngStartIndex As Long, ByVal lngEndIndex As Long, blnMoved As Boolean) As ADODB.Recordset
'根据报告图象 字段获取相关图像
    Dim strSQL As String

    strSQL = "Select rownum As 顺序号,a.医嘱id,a.姓名,a.性别,a.年龄, rownum As 图像号,Replace(Trim(D.Column_Value),'.jpg','') as 图像UID, A.检查UID, " & _
            "'' As 序列UID, 0 as 动态图,'' as 编码名称,'' as 采集时间, '' as 录制长度, '' as 报告图," & _
            "B.FTP用户名 As User1,B.FTP密码 As Pwd1,B.IP地址 As Host1,'/'||B.Ftp目录||'/' As Root1, " & _
            "B.共享目录 as 共享目录1,B.共享目录用户名 as 共享目录用户名1,B.共享目录密码 as 共享目录密码1, " & _
            "Decode(A.接收日期,Null,'',to_Char(A.接收日期,'YYYYMMDD')||'/') ||A.检查UID||'/'||Replace(Trim(D.Column_Value),'.jpg','') As URL,B.设备号 as 设备号1, B.设备名 as 设备名1, " & _
            "C.FTP用户名 As User2,C.FTP密码 As Pwd2,C.IP地址 As Host2,'/'||C.Ftp目录||'/' As Root2, " & _
            "C.共享目录 as 共享目录2,C.共享目录用户名 as 共享目录用户名2,C.共享目录密码 as 共享目录密码2,C.设备号 as 设备号2, C.设备名 as 设备名2 " & _
            "From 影像检查记录 A, 影像设备目录 B, 影像设备目录 C, Table(Cast(f_Str2list(A.报告图象,';') As zlTools.t_Strlist)) D " & _
            "Where A.位置一 = B.设备号(+) And A.位置二 = C.设备号(+) And A.医嘱id = [1]"

    If blnMoved = True Then strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")


    strSQL = "select * from (" & strSQL & " order by 序列UID, 图像号) where 顺序号>=" & lngStartIndex & " and 顺序号<=" & lngEndIndex

    Set GetImageRptData = zlDatabase.OpenSQLRecord(strSQL, "提取报告图像", lngOrderID)
End Function

Private Function GetImageViewData(ByVal slQueryLevel As tQueryLevel, ByVal strSearchValue As String, _
    ByVal lngStartIndex As Long, ByVal lngEndIndex As Long, ByVal blnTmpRecord As Boolean, blnMoved As Boolean, blnOnlyLoadReportImage As Boolean) As ADODB.Recordset
'获取预览图像数据

    Dim strSQL As String

    strSQL = "Select rownum as 顺序号,[2] 医嘱id,c.姓名,c.性别,c.年龄, A.图像号,D.FTP用户名 As User1,D.FTP密码 As Pwd1," & _
            "D.IP地址 As Host1,'/'||D.Ftp目录||'/' As Root1,D.共享目录 as 共享目录1,D.共享目录用户名 as 共享目录用户名1,D.共享目录密码 as 共享目录密码1," & _
            "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/') " & _
            "||C.检查UID||'/'||A.图像UID As URL,d.设备号 as 设备号1, D.设备名 As 设备名1," & _
            "E.FTP用户名 As User2,E.FTP密码 As Pwd2,A.报告图," & _
            "E.IP地址 As Host2,'/'||E.Ftp目录||'/' As Root2,E.共享目录 as 共享目录2,E.共享目录用户名 as 共享目录用户名2,E.共享目录密码 as 共享目录密码2," & _
            "E.设备号 as 设备号2, E.设备名 As 设备名2, A.图像UID,C.检查UID,B.序列UID,A.动态图,A.编码名称,A.采集时间, A.录制长度 " & _
            "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
            "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+)" & IIf(blnOnlyLoadReportImage, " And nvl(A.动态图,0) = 0 ", "")

    If blnTmpRecord Then
        strSQL = Replace(strSQL, "影像检查", "影像临时")
    Else
        If blnMoved Then
            strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
            strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
            strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
        End If
    End If

    Select Case slQueryLevel
        Case slAdvice
            strSQL = "select * from (" & strSQL & " and C.医嘱ID=[1])"
        Case slStudy
            strSQL = "select * from (" & strSQL & " and C.检查UID=[1])"
        Case slSeries
            strSQL = "select * from (" & strSQL & " and B.序列UID=[1])"
        Case slImage
            strSQL = "select * from (" & strSQL & " and A.图像UID=[1])"
    End Select

    strSQL = "select * from (" & strSQL & " order by 序列UID, 图像号) where 顺序号>=" & lngStartIndex & " and 顺序号<=" & lngEndIndex

    Set GetImageViewData = zlDatabase.OpenSQLRecord(strSQL, "查询图像信息", strSearchValue, IIf(blnTmpRecord, "-1", strSearchValue))
End Function


Private Function LoadViewImageToFaceWithNormal(rsCurImageData As ADODB.Recordset, ByRef arrImages() As String) As Boolean
'加载预览图像到界面
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

    '创建本地图像缓存目录
    strCachePath = GetCacheDir
    MkLocalDir strCachePath & objFile.GetParentFolderName(NVL(rsCurImageData("URL")))
    
    ReDim arrImages(0)
BUGEX "LoadViewImageToFaceWithNormal 3"
    Do While Not rsCurImageData.EOF
        '循环加载图像到DicomViewer中
        strImgInstanceUid = Trim(NVL(rsCurImageData!图像UID))

        If InStr(strCurInstanceUids, strImgInstanceUid) <= 0 And strImgInstanceUid <> "" Then

            blnIsAddImage = True

            '设置音视频的显示文件，如果为音视频文件时，该过程将不从服务器中直接下载数据文件
            If NVL(rsCurImageData!动态图, imgTag) = VIDEOTAG Then
                strTmpFile = GetResourceDir & "Avi.bmp"
            ElseIf NVL(rsCurImageData!动态图, imgTag) = AUDIOTAG Then
                strTmpFile = GetResourceDir & "wav.bmp"
            Else
                strTmpFile = strCachePath & NVL(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")
            End If

            strTmpFile = Replace(Trim(strTmpFile), "/", "\")

            If Dir(strTmpFile) = vbNullString Then
                '本地缓存图像不存在，则读取FTP图像
                '建立FTP连接
                If NVL(rsCurImageData("设备号1")) <> vbNullString And Inet1.hConnection = 0 Then
                    If Inet1.FuncFtpConnect(NVL(rsCurImageData("Host1")), NVL(rsCurImageData("User1")), NVL(rsCurImageData("Pwd1"))) = 0 Then
                        If NVL(rsCurImageData("设备号2")) <> vbNullString Then
                            If Inet2.FuncFtpConnect(NVL(rsCurImageData("Host2")), NVL(rsCurImageData("User2")), NVL(rsCurImageData("Pwd2"))) = 0 Then
                                MsgBox "FTP不能正常连接，请检查网络设置。", vbOKOnly, CON_STR_HINT_TITLE
                                Exit Function
                            End If
                        Else
                            MsgBox "FTP不能正常连接，请检查网络设置。", vbOKOnly, CON_STR_HINT_TITLE
                            Exit Function
                        End If
                    End If
                End If

                If Inet1.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsCurImageData("Root1")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")) <> 0 Then
                    '从设备号1提取图像失败，则从设备号2提取图像
                    If NVL(rsCurImageData("设备号2")) <> vbNullString Then
                        If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect NVL(rsCurImageData("Host2")), NVL(rsCurImageData("User2")), NVL(rsCurImageData("Pwd2"))
                        Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsCurImageData("Root2")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", ""))
                    End If
                End If
            End If
            
            
            
            strItem = strTmpFile & "|" & NVL(rsCurImageData!动态图, imgTag) & "|" & NVL(rsCurImageData!报告图) & "|" & _
                NVL(rsCurImageData("编码名称"), "") & "|" & NVL(rsCurImageData("采集时间")) & "|" & strCachePath & NVL(rsCurImageData("URL")) & ".avi" & "|" & _
                Val(NVL(rsCurImageData("录制长度"), "0")) & "|" & NVL(rsCurImageData("图像UID")) & "|" & NVL(rsCurImageData("序列UID")) & "|" & NVL(rsCurImageData("检查UID"))
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
'通过ZLPacsServerCenter服务加载预览图像到界面
'加载预览图像到界面
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
        
    '创建本地图像缓存目录
    MkLocalDir GetResourceDir
    strCachePath = GetCacheDir
    MkLocalDir strCachePath & objFile.GetParentFolderName(NVL(rsCurImageData("URL")))
    
    ReDim arrImages(0)
    Do While Not rsCurImageData.EOF
        '循环加载图像到DicomViewer中
        strImgInstanceUid = Trim(NVL(rsCurImageData!图像UID))
        
        If InStr(strCurInstanceUids, strImgInstanceUid) <= 0 And strImgInstanceUid <> "" Then
            blnIsAddImage = True
            
            '设置音视频的显示文件，如果为音视频文件时，该过程将不从服务器中直接下载数据文件
            If NVL(rsCurImageData!动态图, imgTag) = VIDEOTAG Then
                strTmpFile = GetResourceDir & "Avi.bmp"
            ElseIf NVL(rsCurImageData!动态图, imgTag) = AUDIOTAG Then
                strTmpFile = GetResourceDir & "wav.bmp"
            Else
                strTmpFile = strCachePath & NVL(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")
            End If
            
            strTmpFile = Trim(strTmpFile)
            
            blnIsSendOk = True
            
            If Dir(strTmpFile) = vbNullString Then
                '本地缓存图像不存在，将文件数据发送至服务，则使用服务后台下载
                With fileMsg
                    fileMsg.strAdviceId = Val(NVL(rsCurImageData("医嘱ID")))
                    fileMsg.strName = NVL(rsCurImageData("姓名"))
                    fileMsg.strSex = NVL(rsCurImageData("性别"))
                    fileMsg.strAge = NVL(rsCurImageData("年龄"))
                    
                    fileMsg.ftpInfo.strDeviceId = NVL(rsCurImageData("设备号1"))
                    fileMsg.ftpInfo.strFtpDir = NVL(rsCurImageData("Root1"))
                    fileMsg.ftpInfo.strFTPIP = NVL(rsCurImageData("Host1"))
                    fileMsg.ftpInfo.strFTPPwd = NVL(rsCurImageData("Pwd1"))
                    fileMsg.ftpInfo.strFTPUser = NVL(rsCurImageData("User1"))
                    fileMsg.ftpInfo.strSDDir = NVL(rsCurImageData("共享目录1"))
                    fileMsg.ftpInfo.strSDPswd = NVL(rsCurImageData("共享目录密码1"))
                    fileMsg.ftpInfo.strSDUser = NVL(rsCurImageData("共享目录用户名1"))
                    
                    fileMsg.bakFtpInfo.strDeviceId = NVL(rsCurImageData("设备号2"))
                    fileMsg.bakFtpInfo.strFtpDir = NVL(rsCurImageData("Root2"))
                    fileMsg.bakFtpInfo.strFTPIP = NVL(rsCurImageData("Host2"))
                    fileMsg.bakFtpInfo.strFTPPwd = NVL(rsCurImageData("Pwd2"))
                    fileMsg.bakFtpInfo.strFTPUser = NVL(rsCurImageData("User2"))
                    fileMsg.bakFtpInfo.strSDDir = NVL(rsCurImageData("共享目录2"))
                    fileMsg.bakFtpInfo.strSDPswd = NVL(rsCurImageData("共享目录密码2"))
                    fileMsg.bakFtpInfo.strSDUser = NVL(rsCurImageData("共享目录用户名2"))
                    
                    fileMsg.strLocalDir = strTmpFile
                    fileMsg.strFileName = NVL(rsCurImageData("图像UID")) & IIf(mblnIsLoadReportImage, ".jpg", "")
                    fileMsg.strSubDir = NVL(rsCurImageData("URL"))
                    fileMsg.strMediaType = NVL(rsCurImageData!动态图, imgTag)
                End With
                
                If Not SendDataToservice("缩略图", LoadCommand.COMMAND_RPTIMG_DOWNLOAD, "图像下载", fileMsg) Then
                    blnIsSendOk = False
                End If
            End If
            
            strItem = strTmpFile & "|" & NVL(rsCurImageData!动态图, imgTag) & "|" & NVL(rsCurImageData!报告图) & "|" & _
                NVL(rsCurImageData("编码名称"), "") & "|" & NVL(rsCurImageData("采集时间")) & "|" & strCachePath & NVL(rsCurImageData("URL")) & ".avi" & "|" & _
                Val(NVL(rsCurImageData("录制长度"), "0")) & "|" & NVL(rsCurImageData("图像UID")) & "|" & NVL(rsCurImageData("序列UID")) & "|" & NVL(rsCurImageData("检查UID"))
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
'获取图像总数
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    strSQL = "Select Count(B.Column_Value) 返回值 From 影像检查记录 A, Table(Cast(f_Str2list(Replace(A.报告图象,';',',')) As zlTools.t_Strlist)) B Where 医嘱ID = [1]"
    '如果查询临时记录，则需要将查询表替换为临时存储数据的表
    If blnTmpRecord Then
        strSQL = Replace(strSQL, "影像检查", "影像临时")
    Else
        If blnMoved Then
            strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
            strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        End If
    End If
        
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询图像数量", strSearchValue)
    If rsData.RecordCount > 0 Then lngRecordCount = NVL(rsData!返回值)
    
    If lngRecordCount <= 0 Then
        Select Case slQueryLevel
            Case slAdvice
                strSQL = "select count(1)  as 返回值 from 影像检查图象 a, 影像检查序列 b, 影像检查记录 c where a.序列UID=b.序列UID and b.检查UID=c.检查UID and nvl(a.动态图,0)=0 and c.医嘱ID=[1]"
            Case slStudy
                strSQL = "select count(1)  as 返回值 from 影像检查图象 a, 影像检查序列 b where a.序列UID=b.序列UID and nvl(a.动态图,0)=0 and b.检查UID=[1]"
            Case slSeries
                strSQL = "select count(1)  as 返回值 from 影像检查图象  where nvl(动态图,0)=0 and 序列UID=[1]"
            Case slImage
                strSQL = "select count(1)  as 返回值 from 影像检查图象  where nvl(动态图,0)=0 and 图像UID=[1]"
        End Select
        
        '如果查询临时记录，则需要将查询表替换为临时存储数据的表
        If blnTmpRecord Then
            strSQL = Replace(strSQL, "影像检查", "影像临时")
        Else
            If blnMoved Then
                strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
                strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
            End If
        End If
    
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询图像数量", strSearchValue)
        
        If rsData.RecordCount > 0 Then
            lngRecordCount = NVL(rsData!返回值)
        Else
            lngRecordCount = 0
        End If
    End If
    
    GetRpRecordCount = lngRecordCount
End Function

Public Function GetRecordCount(ByVal slQueryLevel As tQueryLevel, ByVal strSearchValue As String, ByVal blnTmpRecord As Boolean, ByVal blnMoved As Boolean) As Long
'获取图像总数
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    Select Case slQueryLevel
        Case slAdvice
            strSQL = "select count(1)  as 返回值 from 影像检查图象 a, 影像检查序列 b, 影像检查记录 c where a.序列UID=b.序列UID and b.检查UID=c.检查UID and c.医嘱ID=[1]"
        Case slStudy
            strSQL = "select count(1)  as 返回值 from 影像检查图象 a, 影像检查序列 b where a.序列UID=b.序列UID and b.检查UID=[1]"
        Case slSeries
            strSQL = "select count(1)  as 返回值 from 影像检查图象  where  序列UID=[1]"
        Case slImage
            strSQL = "select count(1)  as 返回值 from 影像检查图象  where  图像UID=[1]"
    End Select
    
    '如果查询临时记录，则需要将查询表替换为临时存储数据的表
    If blnTmpRecord Then
        strSQL = Replace(strSQL, "影像检查", "影像临时")
    Else
        If blnMoved Then
            strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
            strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        End If
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询图像数量", strSearchValue)
    
    If rsData.RecordCount > 0 Then
        lngRecordCount = NVL(rsData!返回值)
    Else
        lngRecordCount = 0
    End If
    
    GetRecordCount = lngRecordCount
End Function

Public Function GetImageIdex(ByVal slQueryLevel As tQueryLevel, ByVal strSearchValue As String, ByVal blnTmpRecord As Boolean, ByVal blnMoved As Boolean, strUid As String) As Long
'根据图像UID获取索引
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    Select Case slQueryLevel
        Case slAdvice
            strSQL = "select rownum as 顺序号,a.图像UID,a.序列UID,a.图像号 from 影像检查图象 a, 影像检查序列 b, 影像检查记录 c where a.序列UID=b.序列UID and b.检查UID=c.检查UID and c.医嘱ID=[1]"
        Case slStudy
            strSQL = "select rownum as 顺序号,a.图像UID,a.序列UID,a.图像号 from 影像检查图象 a, 影像检查序列 b where a.序列UID=b.序列UID and b.检查UID=[1]"
        Case slSeries
            strSQL = "select rownum as 顺序号,a.图像UID,a.序列UID,a.图像号 from 影像检查图象  where  序列UID=[1]"
        Case slImage
            strSQL = "select rownum as 顺序号,a.图像UID,a.序列UID,a.图像号 from 影像检查图象  where  图像UID=[1]"
    End Select
    strSQL = "select * from (" & strSQL & " order by 序列UID, 图像号) where 图像UID = [2]"
    '如果查询临时记录，则需要将查询表替换为临时存储数据的表
    If blnTmpRecord Then
        strSQL = Replace(strSQL, "影像检查", "影像临时")
    Else
        If blnMoved Then
            strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
            strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        End If
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询图像数量", strSearchValue, strUid)
    
    If rsData.RecordCount > 0 Then
        lngRecordCount = NVL(rsData!顺序号)
    Else
        lngRecordCount = 0
    End If
    
    GetImageIdex = lngRecordCount
End Function

