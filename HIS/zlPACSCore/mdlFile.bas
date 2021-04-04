Attribute VB_Name = "mdlFile"
Option Explicit
'--------------------------------------------------------
'功  能：跟文件打开，创建目录等相关的函数
'编制人：曾超，胡涛，黄捷
'编制日期：2004.6.12
'-------------------------------------------------------

Public Type DlgFileInfo
    iCount As Long
    sPath As String
    sFile() As String
End Type
'打开文件自定义类型数据
Public Type OpenFileArray
    FilePath As String
    Filename() As String
End Type

Public Sub SaveImages(objImages As DicomImages, ByVal SaveMode As Integer)
'------------------------------------------------
'功能：将图像信息保存到存储服务器中
'参数：
    'SaveMode：0-只存Dicom图像
    '          1-只存报告图像
    '          2-都保存
'返回：
'------------------------------------------------
    Dim Inet As New clsFtp
    Dim strTempPath As String, lngBuffSize As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strDirURL As String, strIp As String, strUser As String, strPwd As String
    Dim RptImageName As String
    Dim strSeriesUID As String
    Dim strStudyUID As String
    Dim img As DicomImage
    Dim dgGlobal As New DicomGlobal
    Dim strReportImages As String
    
    On Error GoTo DBError
    
    '保存临时文件
    strTempPath = Space(255)
    lngBuffSize = GetTempPath(Len(strTempPath), strTempPath)
    strTempPath = Mid(strTempPath, 1, InStrRev(strTempPath, "\"))
    dgGlobal.RegString("UIDRoot") = "1"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '循环保存每一个图像
    For Each img In objImages
        '获取图像的报告图名称
        RptImageName = dgGlobal.NewUID & ".jpg"
        
        '保存临时文件
        If SaveMode <> 1 Then       '存Dicom图像
            img.WriteFile strTempPath & img.InstanceUID, True
        End If
        
        If SaveMode <> 0 Then       '存报告图像
            img.FileExport strTempPath & RptImageName, "JPG"
        End If
        
        '提取FTP保存路径
        If strSeriesUID = "" Or strStudyUID = "" Or strSeriesUID <> img.SeriesUID Then
            '提取图像在数据库中对应的检查UID
            strSeriesUID = img.SeriesUID
            strSQL = "select 检查UID FROM 影像检查序列 where 序列UID =[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询检查UID", CStr(img.SeriesUID))
            If rsTemp.RecordCount = 0 Then
                strStudyUID = PstrCheckUID  '用默认值
            Else
                strStudyUID = rsTemp!检查UID
            End If
        End If
        Call funGetStorageDevice(strStudyUID, strDirURL, strIp, strUser, strPwd)
        Inet.FuncFtpConnect strIp, strUser, strPwd
        
        If SaveMode <> 1 Then
            Inet.FuncUploadFile strDirURL, strTempPath & img.InstanceUID, img.InstanceUID
            Kill strTempPath & img.InstanceUID
        End If
        
        '保存报告图
        If SaveMode <> 0 Then
            '检查报告图数量是否超长
            strSQL = "Select 报告图象 From 影像检查记录 Where 检查UID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询报告图象", strStudyUID)
            If rsTemp.RecordCount > 0 Then
                strReportImages = Nvl(rsTemp("报告图象"))
                If Len(strReportImages & " ;" & RptImageName) >= 4000 Then
                    MsgBox "报告图像数量超过上限，请先删除部分报告图后，再继续保存报告图。", vbInformation, gstrSysName
                Else
                    Inet.FuncUploadFile strDirURL, strTempPath & RptImageName, RptImageName
                    Kill strTempPath & RptImageName
                    
                    strSQL = "ZL_影像检查报告_ADD('" & strStudyUID & "','" & RptImageName & "')"
                    zlDatabase.ExecuteProcedure strSQL, "保存报告图像"
                End If
            End If
        End If
        
        Inet.FuncFtpDisConnect
    Next
    Exit Sub
DBError:
    Inet.FuncFtpDisConnect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function funGetStorageDevice(strStudyUID As String, ByRef strDirURL As String, ByRef strIp As String, _
        ByRef strUser As String, ByRef strPwd As String) As Boolean
'------------------------------------------------
'功能：从数据库中读取制定存储设备ID的FTP访问参数
'参数： strSaveDeviceID －－存储设备ID
'       strDirURL－－[OUT] FTP目录
'       strIp －－[OUT] IP地址
'       strUser －－ [OUT]用户名
'       strPwd －－[OUT]用户名
'返回：True－－获取成功，False－－获取失败
'-----------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    '检查存储设备是否存在
    strSQL = "Select b.接收日期, '/'||Decode(c.Ftp目录,Null,'',c.Ftp目录||'/') As 目录1,c.FTP用户名 As 用户名1,c.FTP密码 As 密码1,c.IP地址 As IP地址1," & _
             " '/'||Decode(d.Ftp目录,Null,'',d.Ftp目录||'/') As 目录2,d.FTP用户名 As 用户名2,d.FTP密码 As 密码2,d.IP地址 As IP地址2 " & _
             " from 影像检查记录 b,影像设备目录  c ,影像设备目录 d " & _
             " where (B.位置一 = C.设备号 And b.位置二=d.设备号(+) )  And b.检查UID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strStudyUID)
     '没有存储设备时退出
    If rsTemp.EOF = True Then
        MsgBox "没有找到存储设备,该图像可能是外部图像。", vbInformation, App.ProductName
        funGetStorageDevice = False
        Exit Function
    End If
    strDirURL = Nvl(rsTemp("目录1"))
    strIp = Nvl(rsTemp("IP地址1"))
    strUser = Nvl(rsTemp("用户名1"))
    strPwd = Nvl(rsTemp("密码1"))
    If strIp = "" Or strUser = "" Then  '位置一可能没有图像，读取位置二的图像
        strDirURL = Nvl(rsTemp("目录2"))
        strIp = Nvl(rsTemp("IP地址2"))
        strUser = Nvl(rsTemp("用户名2"))
        strPwd = Nvl(rsTemp("密码2"))
    End If
    strDirURL = strDirURL & Format(Nvl(rsTemp("接收日期")), "YYYYMMDD") & "/" & strStudyUID & "/"
    funGetStorageDevice = True
End Function

Function funIsLabelMouse(f As frmViewer, Button As Integer, Shift As Integer) As Boolean
'------------------------------------------------
'功能：判断标注鼠标标志是否按下
'参数：f--鼠标单击的窗体；Button--鼠标的左右键编号；Shift--鼠标的shift状态。
'返回：True-成功；False-失败。
'------------------------------------------------
    funIsLabelMouse = False
    If Button_miFrameSelectImage And Button = cMouseUsage("201").lngMouseKey And Shift = cMouseUsage("201").lngShift Then        '框选图象，使用矩形标注
        '框选图像，LabelStyle=矩形
        funIsLabelMouse = True
        intSelectLabelStyle = 2
    ElseIf Button_miLabelRectangle And Button = cMouseUsage("2").lngMouseKey And Shift = cMouseUsage("2").lngShift Then
        '矩形标注，LabelStyle=矩形
        funIsLabelMouse = True
        intSelectLabelStyle = 2
    ElseIf Button_miLabelLine And Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift Then
        '直线标注，LabelStyle=直线
        funIsLabelMouse = True
        intSelectLabelStyle = 3
    ElseIf Button_miLabelVasMeasure And Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift Then
        '血管狭窄测量，LabelStyle=直线
        funIsLabelMouse = True
        intSelectLabelStyle = 3
    ElseIf Button_miLabelCadiothoracicRatio And Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift Then
        '心胸比测量，LabelStyle=直线
        funIsLabelMouse = True
        intSelectLabelStyle = 3
    ElseIf Button_miLabelEllipse And Button = cMouseUsage("3").lngMouseKey And Shift = cMouseUsage("3").lngShift Then
        '椭圆标注，LabelStyle=椭圆
        funIsLabelMouse = True
        intSelectLabelStyle = 1
    ElseIf Button_miLabelArrowhead And Button = cMouseUsage("4").lngMouseKey And Shift = cMouseUsage("4").lngShift Then
        '箭头标注，LabelStyle=箭头
        funIsLabelMouse = True
        intSelectLabelStyle = 10
    ElseIf Button_miLabelPolygon And Button = cMouseUsage("5").lngMouseKey And Shift = cMouseUsage("5").lngShift Then
        '多边形标注，LabelStyle=多边形
        funIsLabelMouse = True
        intSelectLabelStyle = 5
    ElseIf Button_miLabelPolyLine And Button = cMouseUsage("6").lngMouseKey And Shift = cMouseUsage("6").lngShift Then
        '多边线标注，LabelStyle=多边线
        funIsLabelMouse = True
        intSelectLabelStyle = 4
    ElseIf Button_miLabelAngle And Button = cMouseUsage("7").lngMouseKey And Shift = cMouseUsage("7").lngShift Then
        '角度标注，LabelStyle=直线
        funIsLabelMouse = True
        intSelectLabelStyle = 3
    ElseIf Button_miAutoWidthLevel And Button = cMouseUsage("105").lngMouseKey And Shift = cMouseUsage("105").lngShift Then
        '自动窗宽窗位的矩形框，LabelStyle=矩形
        funIsLabelMouse = True
        intSelectLabelStyle = 2
    End If
End Function

Public Function funGetFileList(f As frmViewer) As OpenFileArray
'------------------------------------------------
'功能：打开读取文件对话框,传出全路径名文件数组
'参数：f--窗体。
'返回：全路径文件名数组
'2009用
'------------------------------------------------
    Dim DlgInfo As DlgFileInfo
    Dim i As Integer
    On Error GoTo errHandle
    '选择文件
    With f.Common
        
        .CancelError = False
        .MaxFileSize = 32767 '被打开的文件名尺寸设置为最大，即32K
        .Flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
        .DialogTitle = "选择文件"
        .Filter = "DICOM文件（*.dcm）(*.img)|*.dcm;*.img|图像文件 (*.BMP)(*.JPG)|*.BMP;*.JPG|所有文件（*.*）|*.*"
        .ShowOpen
        If .Filename <> "" Then
            DlgInfo = GetDlgSelectFileInfo(.Filename)
        End If
        .Filename = ""      '在打开了*.pif文件后须将Filename属性置空，
                            '否则当选取多个*.pif文件后，当前路径会改变
    End With
    
    If DlgInfo.iCount <= 0 Then
        ReDim funGetFileList.Filename(0)
        funGetFileList.FilePath = ""
        Exit Function
    End If
    
    ReDim funGetFileList.Filename(DlgInfo.iCount)
    funGetFileList.FilePath = DlgInfo.sPath
    For i = 1 To DlgInfo.iCount
        funGetFileList.Filename(i) = DlgInfo.sFile(i)
    Next i
    Exit Function
errHandle:
    ReDim funGetFileList.Filename(0)
    funGetFileList.FilePath = ""
    MsgBox "打开图像错误，请重试。可能是一次性打开的图像数量过多。", vbExclamation, gstrSysName
End Function

Public Function GetDlgSelectFileInfo(strFileName As String) As DlgFileInfo
'------------------------------------------------
'功能：将文件名转化为全路径数组
'参数：strFileName--文件名，通过打开文件控件来获得。
'返回：全路径数组
'编制人：曾超
'------------------------------------------------
    Dim sPath, tmpStr As String
    Dim sFile() As String
    Dim iCount, i As Integer
    On Error GoTo errHandle
    sPath = CurDir()  '获得当前的路径，因为在CommonDialog中改变路径时会改变当前的Path
    tmpStr = Right$(strFileName, Len(strFileName) - Len(sPath)) '将文件名分离出来
    
    If left$(tmpStr, 1) = Chr$(0) Then
        '选择了多个文件(表现为第一个字符为空格)
        For i = 1 To Len(tmpStr)
            If Mid$(tmpStr, i, 1) = Chr$(0) Then
                iCount = iCount + 1
                ReDim Preserve sFile(iCount)
            Else
                sFile(iCount) = sFile(iCount) & Mid$(tmpStr, i, 1)
            End If
        Next i
    Else
        '只选择了一个文件(注意：根目录下的文件名除去路径后没有"\"）
        iCount = 1
        ReDim Preserve sFile(iCount)
        If left$(tmpStr, 1) = "\" Then tmpStr = Right$(tmpStr, Len(tmpStr) - 1)
        sFile(iCount) = tmpStr
    End If
    
    GetDlgSelectFileInfo.iCount = iCount
    ReDim GetDlgSelectFileInfo.sFile(iCount)
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    GetDlgSelectFileInfo.sPath = sPath
    For i = 1 To iCount
        GetDlgSelectFileInfo.sFile(i) = sFile(i)
    Next i
    Exit Function
errHandle:
    MsgBox "GetDlgSelectFileInfo函数执行错误！", vbExclamation, gstrSysName
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'功能：创建本地目录
'参数： strDir－－本地目录
'返回：无
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function funcCeateAViewer(intSeriesIndex As Integer, thisForm As frmViewer) As Integer
'------------------------------------------------
'功能：在窗体中装载一个Viewer和滚动条，初始化该Viewer所对应的MSF参数,从正本ZLSeriesInfos中复制信息到ZLShowSeriesInfos
'      将图像显示到新增的序列中，并显示图像中原来保存的标注。
'      当图像序列为空时，则仅仅初始化该序列的MSF参数，装载Viewer和滚动条。
'参数：intSeriesIndex--需要装载的图像所在序列的索引,如果为0，则不装载任何图像
'      thisForm--显示图像的窗体
'返回：创建成功的Viewer的Index
'时间：2009-7
'------------------------------------------------
    Dim i As Integer
    Dim oneSeriesInfo As clsSeriesInfo
    Dim oneImageInfo As clsImageInfo
    Dim intViewerIndex As Integer
    Dim intCurrentIndex As Integer
    Dim intImagesCount As Integer
    
    'intSeriesIndex=0表示这个Viewer中的图像暂时还不知道放什么图像好
    If intSeriesIndex = 0 Then
        If ZLSeriesInfos.Count = 0 Then Exit Function
        intCurrentIndex = 1
    Else
        intCurrentIndex = intSeriesIndex
    End If
    If ZLSeriesInfos.Count < intCurrentIndex Then Exit Function
    
    funcCeateAViewer = 0
    On Error GoTo err
    
    '初始化MSFViewer中这个序列的内容
    With thisForm.MSFViewer
        .Rows = .Rows + 1
        intViewerIndex = .Rows - 1
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        .TextMatrix(intViewerIndex, 0) = ZLSeriesInfos(intCurrentIndex).lngSource  '图像来源
        .TextMatrix(intViewerIndex, 1) = True                                     '是否有图
        .TextMatrix(intViewerIndex, 2) = ZLSeriesInfos(intCurrentIndex).StudyUID   '检查UID
        .TextMatrix(intViewerIndex, 3) = 1        '当前选择的图像号
        .TextMatrix(intViewerIndex, 4) = 1        '当前选择的图像处于第几帧
        .TextMatrix(intViewerIndex, 5) = 0        '该序列横向显示图像数目(供序列内单图和多图显示切换用)
        .TextMatrix(intViewerIndex, 6) = 0        '该序列纵向显示图像数目(供序列内单图和多图显示切换用)
        .TextMatrix(intViewerIndex, 7) = 1        '该序列当前显示第一个图像序号(供序列内单图和多图显示切换用)
        .TextMatrix(intViewerIndex, 8) = 1        '该序列当前显示选择图像序号(供序列内单图和多图显示切换用)
        .TextMatrix(intViewerIndex, 9) = True     '该序列内图像是否自动同步
        .TextMatrix(intViewerIndex, 15) = 0       '记录当前序列是否被选择，用于自动和手工序列同步
    End With
    
    '装载Viewer、滚动条和图像
    With thisForm
        '装载Viewer和滚动条
        load .Viewer(intViewerIndex)
        load .VScro(intViewerIndex)
        .Viewer(intViewerIndex).UseScrollBars = False
        .Viewer(intViewerIndex).Visible = False
        .Viewer(intViewerIndex).Tag = intCurrentIndex
        .Viewer(intViewerIndex).CellSpacing = lngCellSpacing
        .Viewer(intViewerIndex).BackColour = lngViewerBackColor
        .VScro(intViewerIndex).Visible = False
         
        '装载ZLShowSeriesInfos结构
        If ZLShowSeriesInfos.Count = intViewerIndex - 1 Then
        
            Set oneSeriesInfo = funGetNewSeriesInfo
            Call funCopySeriesInfo(ZLSeriesInfos(intCurrentIndex), oneSeriesInfo)
            
            ZLShowSeriesInfos.Add oneSeriesInfo
            '装载ZLShowSeriesInfos结构中的图像,如果intSeriesIndex =0 则只装载第一个图
            If intSeriesIndex = 0 Then
                intImagesCount = 1
            Else
                intImagesCount = ZLSeriesInfos(intCurrentIndex).ImageInfos.Count
            End If
            
            For i = 1 To intImagesCount
                Set oneImageInfo = funGetNewImageInfo
                Call funCopyImageInfo(ZLSeriesInfos(intCurrentIndex).ImageInfos(i), oneImageInfo)
                
                ZLShowSeriesInfos(intViewerIndex).ImageInfos.Add oneImageInfo
            Next i
        End If
        '设定图像布局
        Call subSetImageLayout(.Viewer(intViewerIndex), ZLSeriesInfos(intCurrentIndex).strModality, ZLSeriesInfos(intCurrentIndex).ImageInfos.Count)
        '图像排序,正本图像从数据库中读取，是按照图像号排序的，这里要重新设置显示序列的排序方法
        Call subSortImages(thisForm, intViewerIndex, funGetImageSort(ZLSeriesInfos(intCurrentIndex).strModality))
        '装载图像
        Call subShowALLImage(thisForm, .Viewer(intViewerIndex), 1, False)
        
    End With
    funcCeateAViewer = intViewerIndex
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub subSetImageLayout(thisViewer As DicomViewer, strModality As String, intImageCount As Integer)
'------------------------------------------------
'功能：根据影像类别和图像数量来排布图像布局
'参数： thisViewer--进行图像布局重排的序列
'       strModality--进行图像布局重排的影像类别
'       intImageCount--图像总数量
'返回：无，直接重排指定序列的图像布局
'引用的外部参数：G_INT_MAX_IMG_COL；G_INT_MAX_IMG_ROW
'时间：2009-7
'------------------------------------------------
    Dim i As Integer
    Dim blnSetImageLayout As Boolean
    Dim intRows As Integer
    Dim intCols As Integer
    
    For i = 1 To UBound(aPresetLayout)
        If UCase(aPresetLayout(i).strModality) = UCase(strModality) Then
            If aPresetLayout(i).bImageAutoFormat Then
                '是自动布局，则根据图像总数量和最大布局数量设置图像布局
                '但是此时thisViewer还没有摆放到窗口中，它的宽度和高度其实是没有实际意义的。
                ResizeRegion intImageCount, thisViewer.width, thisViewer.height, intRows, intCols, G_INT_MAX_IMG_ROW, G_INT_MAX_IMG_COL
                thisViewer.MultiColumns = intCols
                thisViewer.MultiRows = intRows
            Else
                thisViewer.MultiColumns = aPresetLayout(i).lngImageColumns
                thisViewer.MultiRows = aPresetLayout(i).lngImageRows
            End If
            blnSetImageLayout = True
        End If
    Next i
    
    '如果没有设置用户定义的图像行数和列数，则用默认值1*1
    If blnSetImageLayout = False Then
        thisViewer.MultiColumns = 1
        thisViewer.MultiRows = 1
    End If
End Sub

Public Function funGetImageSort(strModality As String) As Long
'------------------------------------------------
'功能：根据影像类别查找图像排序方式
'参数:
'       strModality--进行图像排序的影像类别
'返回：图像排序方式
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
    
    funGetImageSort = 0
    
    For i = 1 To UBound(aPresetLayout)
        If UCase(aPresetLayout(i).strModality) = UCase(strModality) Then
            funGetImageSort = aPresetLayout(i).lngImageSort - 1
        End If
    Next i
    Exit Function
err:
    '不处理
End Function

Public Sub subShowALLImage(thisForm As frmViewer, thisViewer As DicomViewer, intImageIndex As Integer, blnFast As Boolean)
'------------------------------------------------
'功能： 显示所有可见的图像，intImageIndex指定左上角的图像，
'       根据图像的行*列布局，确定后续的图像是否显示，
'       如果该图像已经存在Viewer中则直接显示，
'       如果图像不在Viewer中，则把图像加入Viewer,对于同时显示的图像也一起处理
'       这个过程用于创建序列后、修改布局、直接拖动滚动条显示某个图像
'参数： thisForm -- 观片窗体
'       thisViewer--进行图像布局重排的序列
'       intImageIndex--图像所在的图像索引
'       blnFast     -- 快速显示图像，则不作subDispframe，subDisplayPatientInfo和thisViewer.Refresh
'返回：无，直接把图像加入并显示出来
'时间：2009-7
'------------------------------------------------
    Dim iFoundImageIndex As Integer
    Dim strSaveDir As String
    Dim cFTP As clsFtp
    Dim iCurrImageIndex As Integer
    Dim intImagesCount As Integer
    Dim blnExit As Boolean
    Dim intViewerIndex As Integer
    Dim intAddImageCount As Integer
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo err
    
    '首先判断当前图像的布局，总共需要显示多少图像，然后再循环装载这些图像
    intViewerIndex = thisViewer.Index
    intImagesCount = ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
    
    iCurrImageIndex = intImageIndex
    If iCurrImageIndex > intImagesCount - thisViewer.MultiColumns * thisViewer.MultiRows + 1 Then
        iCurrImageIndex = intImagesCount - thisViewer.MultiColumns * thisViewer.MultiRows + 1
    End If
    If iCurrImageIndex <= 0 Then iCurrImageIndex = 1
    
    '查找图像的显示位置
    '因为ZLShowSeriesInfos中所有图像都是有序的，因此新增加的图像所在Viwer中的位置应该就是iCurrImageIndex
    '但是可能前面的图像没有显示过而没有加载，因此新增加的图象在Viewer中的位置，从intImageIndex往前找
    'Viwer中每个图像的Tag是这个图像的ImageIndex，因此判断图像的Tag >iCurrImageIndex 。
    iFoundImageIndex = 0
    For i = IIf(thisViewer.Images.Count > iCurrImageIndex, iCurrImageIndex, thisViewer.Images.Count) To 1 Step -1
        If thisViewer.Images(i).Tag = iCurrImageIndex Then
            iFoundImageIndex = i
            Exit For
        ElseIf thisViewer.Images(i).Tag < iCurrImageIndex Then
            iFoundImageIndex = i + 1
            Exit For
        End If
    Next i
    If iFoundImageIndex = 0 Then iFoundImageIndex = 1   '新增加的图像的位置是 iFoundImageIndex
    If iFoundImageIndex > iCurrImageIndex Then iFoundImageIndex = iCurrImageIndex
    
    '连接FTP
'    Set cFTP = New clsFtp
'    cFTP.FuncFtpConnect ZLSeriesInfos(intSeriesIndex).strHostIP, ZLSeriesInfos(intSeriesIndex).strFTPUser, _
'            ZLSeriesInfos(intSeriesIndex).strFTPPasw
'    cFTP.FuncChangeDir ZLSeriesInfos(intSeriesIndex).strFTPDir & Replace(ZLSeriesInfos(intSeriesIndex).strSaveDir, "\", "/")
        

    For i = 1 To thisViewer.MultiRows
        For j = 1 To thisViewer.MultiColumns

'           首先判断图像是否已经加载，如果已经加载，则找到这个图像并显示出来，如果没有加载，则加载该图像
            If ZLShowSeriesInfos(intViewerIndex).ImageInfos(iCurrImageIndex).blnDisplayed = False Then
                '没有加载，则找到位置并加载图像
                intAddImageCount = funcAddAImage(thisViewer, iCurrImageIndex, iFoundImageIndex, cFTP)
                If intAddImageCount > 0 Then
                    '因为多帧图像在装载后，会改变ZLShowSeriesInfos中图像数量的多少，因此需要重新赋值
                    intImagesCount = ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
                    '重新设置滚动条的最大图像数量
                    thisForm.blnVscroInvoked = True
                    thisForm.VScro(intViewerIndex).Max = intImagesCount
                    thisForm.blnVscroInvoked = False
                End If
            End If

'           已经加载,而且是当前布局的第一个图，则找到这个图像并显示，当前布局的其他图像不需要处理
            If i = 1 And j = 1 And iFoundImageIndex <= thisViewer.Images.Count Then
                thisViewer.CurrentIndex = iFoundImageIndex
            End If

            iFoundImageIndex = iFoundImageIndex + 1
            iCurrImageIndex = iCurrImageIndex + 1
            '如果图像的索引大于图像的总数量，则退出循环
            If iFoundImageIndex > intImagesCount Then
                blnExit = True
                Exit For
            End If
        Next j
        '如果内层循环已经退出，则外层循环也一起退出
        If blnExit = True Then
            Exit For
        End If
    Next i
    
    '显示或者隐藏图像中的病人信息
    Call subDisplayPatientInfo(thisViewer)
        
    '如果是快速显示，则不处理以下功能
    If blnFast = False Then
        '图像显示完后，添加Viewer中的标注：图象框、右下角的选择标记等
        Call subDispframe(thisForm, thisViewer)
    
        thisViewer.Refresh
    End If
    
    If Not cFTP Is Nothing Then cFTP.FuncFtpDisConnect
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    If Not cFTP Is Nothing Then cFTP.FuncFtpDisConnect
    Call SaveErrLog
End Sub

Public Function funcAddAImageA(thisViewer As DicomViewer, ByVal intImageIndex As Integer)
'------------------------------------------------
'功能： 把指定序列和图像索引的图像添加到Viewer中，并自动查找图像位置，把图像移动到适合的位置
'       加入图像的时候，查找图像的方式是：首先查看本地缓存---再提取共享目录---最后通过FTP下载
'参数： thisViewer--进行图像布局重排的序列
'       intImageIndex--图像所在的图像索引
'返回：添加的图像数量
'时间：2009-7
'------------------------------------------------
    Dim intViewerIndex As Integer
    Dim iFoundImageIndex As Integer
    Dim iCurrImageIndex As Integer
    Dim cFTP As clsFtp
    Dim i As Integer
    
    On Error GoTo err
    
    intViewerIndex = thisViewer.Index
    iCurrImageIndex = intImageIndex
    
    '查找图像的显示位置
    '因为ZLShowSeriesInfos中所有图像都是有序的，因此新增加的图像所在Viwer中的位置应该就是iCurrImageIndex
    '但是可能前面的图像没有显示过而没有加载，因此新增加的图象在Viewer中的位置，从intImageIndex往前找
    'Viwer中每个图像的Tag是这个图像的ImageIndex，因此判断图像的Tag >iCurrImageIndex 。
    iFoundImageIndex = 0
    For i = IIf(thisViewer.Images.Count > iCurrImageIndex, iCurrImageIndex, thisViewer.Images.Count) To 1 Step -1
        If thisViewer.Images(i).Tag = iCurrImageIndex Then
            iFoundImageIndex = i
            Exit For
        ElseIf thisViewer.Images(i).Tag < iCurrImageIndex Then
            iFoundImageIndex = i + 1
            Exit For
        End If
    Next i
    If iFoundImageIndex = 0 Then iFoundImageIndex = 1   '新增加的图像的位置是 iFoundImageIndex
    If iFoundImageIndex > iCurrImageIndex Then iFoundImageIndex = iCurrImageIndex

    '检查图像是否已经被显示了,则退出程序
    If ZLShowSeriesInfos(intViewerIndex).ImageInfos(iCurrImageIndex).blnDisplayed = True Then
        funcAddAImageA = 0
        Exit Function
    End If
    
    funcAddAImageA = funcAddAImage(thisViewer, iCurrImageIndex, iFoundImageIndex, cFTP)
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function funcAddAImage(thisViewer As DicomViewer, ByVal intImageIndex As Integer, ByVal intCurrentIndex As Integer, cFTP As clsFtp) As Integer
'------------------------------------------------
'功能： 把指定序列和图像索引的图像添加到Viewer中，并移动到适合的位置
'       加入图像的时候，查找图像的方式是：首先查看本地缓存---再提取共享目录---最后通过FTP下载
'参数： thisViewer--进行图像布局重排的序列
'       intImageIndex--图像所在的图像索引
'       intCurrentIndex---图像需要摆放的位置
'       cFTP ---FTP连接，应该已经连接好，并且设置好目录的了
'返回：添加的图像数量
'时间：2009-7
'------------------------------------------------
    Dim NewImg As DicomImage
    Dim img As DicomImage
    Dim intViewerIndex As Integer
    Dim i As Integer
    Dim OldImageInfos As Collection
    Dim OneImageInfos As clsImageInfo
    Dim intImageCount As Integer
    
    On Error GoTo err
    
    intViewerIndex = thisViewer.Index
    Set NewImg = funLoadAImage(intViewerIndex, intImageIndex, 1)
    If NewImg Is Nothing Then
        MsgBox "加载文件出错，请检查本机缓存或者FTP连接。", vbOKOnly, "下载图像提示"
        Exit Function
    End If
    
    '设置图像已经显示过的标记
    ZLShowSeriesInfos(intViewerIndex).ImageInfos(intImageIndex).blnDisplayed = True
    
    '对多帧和单帧图像进行处理
    If NewImg.FrameCount > 1 Then  '多帧图像
        '补充填写ZLShowSeriesInfos结构
        Set OldImageInfos = New Collection
        
        For intImageCount = 1 To ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
            OldImageInfos.Add ZLShowSeriesInfos(intViewerIndex).ImageInfos(intImageCount)
        Next intImageCount
        
        '把多帧图像的图像信息添加到集合中
        Set OneImageInfos = ZLShowSeriesInfos(intViewerIndex).ImageInfos(intImageIndex)
        
        Set ZLShowSeriesInfos(intViewerIndex).ImageInfos = Nothing
        
        For i = 1 To intImageIndex
            ZLShowSeriesInfos(intViewerIndex).ImageInfos.Add OldImageInfos(i)
        Next i
        
        For i = 2 To NewImg.FrameCount
            ZLShowSeriesInfos(intViewerIndex).ImageInfos.Add OneImageInfos
        Next i
        '把后面的图像信息补充回集合中
        For i = intImageIndex + 1 To OldImageInfos.Count
            ZLShowSeriesInfos(intViewerIndex).ImageInfos.Add OldImageInfos(i)
        Next i
        '清空OldImageInfos
        Set OldImageInfos = Nothing
    End If
    
    '加载图像
    For i = 1 To NewImg.FrameCount
        thisViewer.Images.Add NewImg
        Set img = thisViewer.Images(thisViewer.Images.Count)
        img.Tag = intImageIndex
        img.Frame = i
    
        Call subInitAImage(img, intViewerIndex, thisViewer)
        
        '把图像移动到合适的位置
        If intCurrentIndex <> thisViewer.Images.Count And thisViewer.Images.Count <> 0 Then
            thisViewer.Images.Move thisViewer.Images.Count, intCurrentIndex
        End If
        
        intImageIndex = intImageIndex + 1
        intCurrentIndex = intCurrentIndex + 1
    Next i
    
    funcAddAImage = NewImg.FrameCount
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub subInitAImage(img As DicomImage, intViewerIndex As Integer, thisViewer As DicomViewer)
'------------------------------------------------
'功能： 初始化图像,包括对图像的窗口等进行同步
'参数： img--需要初始化的图像
'       intViewerIndex--图像所在的Viewer的索引,0表示不使用这个索引
'       thisViewer -- 图像即将要加入thisViewer中
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim strRefUID As String
    
    If img Is Nothing Then Exit Sub
    
    On Error GoTo err
    '处理宁波公司的MR图象：定位线不能显示问题,修改规则为，将(0020,0052) : Frame of Reference UID修改成后缀为1
    If Not IsNull(img.Attributes(&H8, &H60).Value) And Not IsNull(img.Attributes(&H8, &H70).Value) _
        And Not IsNull(img.Attributes(&H20, &H52).Value) Then
        
        If img.Attributes(&H8, &H60).Value = "MR" And img.Attributes(&H8, &H70).Value = "NingBo XGY Magnetism Co.,LTD." Then
            strRefUID = img.Attributes(&H20, &H52).Value
            strRefUID = left(strRefUID, InStrRev(strRefUID, ".")) & "1"
            img.Attributes.Add &H20, &H52, strRefUID
        End If
    End If
    
    '处理东软的DR图像，设置VOILUT=0，打开图像后才能正常显示
    If Not IsNull(img.Attributes(&H8, &H1090).Value) And Not IsNull(img.Attributes(&H8, &H60).Value) Then
        If UCase(img.Attributes(&H8, &H60).Value) = "DX" And _
            (UCase(img.Attributes(&H8, &H1090).Value) = UCase("NavigationSight") Or UCase(img.Attributes(&H8, &H1090).Value) = UCase("NeuVision DR III")) Then
            img.VOILUT = 0
        End If
    End If
    
    '取消自动剪影,因为DicomObjects控件本身对处理剪影有BUG，存在（0028，6100）时，会自动对图像进行剪影，
    '导致晋煤的DSA图像不能正常显示
    '虽然设置图像的mask=0 ,可以取消剪影，但是每次图像被添加到新的Dicomimages之后，自动又将mask设置成1了，
    '这样在程序中无法很好的控制，因此直接去掉（0028，6100）这个属性。
    If Not IsNull(img.Attributes(&H28, &H6100).Value) Then
        img.Attributes.Remove &H28, &H6100
    End If
    
    '设置图像的放大模式
    img.MagnificationMode = intMagnificationMode
    
    '设置图像的增强幅度
    img.UnsharpLength = 0
    
    '取消显示缓存，可以加快显示速度，并且减少对内存的消耗
    '有两类图像设置CacheDisplay为False会出错，需要特殊处理。飞利浦Intera MR
    If Not IsNull(img.Attributes(&H8, &H60).Value) Then
        If UCase(img.Attributes(&H8, &H60).Value) = "PR" Or UCase(img.Attributes(&H8, &H60).Value) = "KO" Or UCase(img.Attributes(&H8, &H60).Value) = "SR" Then
            '类型为PR的，不做任何处理，否则会出错
        ElseIf UCase(img.Attributes(&H8, &H60).Value) = "MR" And Not IsNull(img.Attributes(&H8, &H16).Value) Then
            If left(img.Attributes(&H8, &H16).Value, Len(img.Attributes(&H8, &H16).Value) - 1) = "1.3.46.670589.11.0.0.12." Or _
                img.Attributes(&H8, &H16).Value = "1.2.840.10008.5.1.4.1.1.66" Then
                  '类型为MR的，经测试得知，如果Sop Class UID ="1.3.46.670589.11.0.0.12.2"或"1.3.46.670589.11.0.0.12.4" ，则也不做任何处理，否则会出错
                  '还可能有其他的SOP ClassUID,因此判断前缀“1.3.46.670589.11.0.0.12.xxx”
            Else
                img.CacheDisplay = False
            End If
        Else
            img.CacheDisplay = False
        End If
    Else
        img.CacheDisplay = False
    End If
    
    
    '初始化图像中的系统标注
    subInitImageLabels intViewerIndex, 1, img, True, True, True     '初始化图像标注信息信息:系统标注；体位信息；标尺；四角信息
    
    '初始化图像的遮盖图
    If left(img.InstanceUID, 24) = "2.16.840.1.113669.632.3." Then
        subDrawImgShutter img, True
    Else
        subDrawImgShutter img
    End If
    
    
    '显示保存在图像中的标注
    subReadLabelFromImg img           ''''读出图像的标注信息
    
    If img.Attributes(&H6000, &H10).Exists = True Then
        img.OverlayVisible(0) = Button_miShowOverlay
    End If
    
    '设置预设的窗宽窗位
    If intViewerIndex > 0 And intViewerIndex <= ZLShowSeriesInfos.Count Then
        If ZLShowSeriesInfos(intViewerIndex).lngWinWidth <> 0 And ZLShowSeriesInfos(intViewerIndex).lngWinLevel <> 0 Then
            img.width = ZLShowSeriesInfos(intViewerIndex).lngWinWidth
            img.Level = ZLShowSeriesInfos(intViewerIndex).lngWinLevel
            img.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & img.width & "-L:" & img.Level
            '使用默认的窗宽窗位，需要设置VOILUT=0才能有效
            img.VOILUT = 0
        End If
    End If
    
    '处理一个Overlay的显示,Overlay的文字一般是白色的，因此最好把图像底色设置成1
    If Not IsNull(img.Attributes(&H6000, &H15).Value) Then
        If img.Attributes(&H6000, &H15).Value = 1 Then
            If img.Level = 0 Then img.Level = 1
            img.OverlayVisible(0) = True
            img.OverlayColour(0) = lngLabelColor
        End If
    End If
    
    '处理图像内容同步
    If Button_miImageInPhase = True And Not thisViewer Is Nothing Then
        If thisViewer.Images.Count > 0 Then
            Call subImageInPhase(img, thisViewer.Images(1), IMG_SYN_All)
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subShowForm(thisForm As frmViewer)
'------------------------------------------------
'功能： 根据ZLSeriesInfos中图像的内容显示窗口，创建并显示Viewer，滚动条，显示Viewer中的图像
'       把需要显示的图像加载进入VIEWER,把Viewer显示到窗口中，显示Viewer相关的滚动条
'       加入图像的时候，查找图像的方式是：首先查看本地缓存---再提取共享目录---最后通过FTP下载
'参数： 无
'返回：无，直接把Viewer和图像加入并显示出来
'时间：2009-7
'------------------------------------------------
    Dim strModality As String
    Dim intSeriesCount As Integer
    Dim blnLoadOver As Boolean
    Dim intCurrentSeries As Integer
    Dim intCurrentViewer  As Integer
    Dim intViewerIndex As Integer
    Dim intSeriesIndex As Integer
    Dim oneImageInfo As clsImageInfo
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    '通过“观片”方式打开和通过“对比”方式打开，对原有Viewer的处理是不一样的。
    
    If ZLSeriesInfos.Count = 0 Then Exit Sub
    
    On Error GoTo err
    
    intSeriesCount = ZLSeriesInfos.Count
    intCurrentSeries = 0
    intViewerIndex = 0
    
    '通过“观片”方式打开时Viewer的数量为1 ，通过“对比”方式打开时，Viewer的数量大于1
    '首先判断是“观片”方式打开还是“对比”方式打开
    If thisForm.Viewer.Count = 1 Then   '“观片”方式打开观片站
        '通过第一个图的影像类别，确定序列布局，摆放分隔条
        strModality = ZLSeriesInfos(1).strModality
        '根据第一个图像的影像类别，获得图像的序列布局
        Call subSetSeriesLayout(thisForm, strModality, intSeriesCount)
        '摆放分隔条
        Call subShowSpliter(thisForm)
    End If
    
    '根据序列布局，创建Viewer，滚动条，并装载图像。
    For i = 1 To thisForm.intCountY
        For j = 1 To thisForm.intCountX
            '布局的数量大于序列的总数量，则退出循环
            If (i - 1) * thisForm.intCountX + j > intSeriesCount Then
                blnLoadOver = True
                Exit For
            End If
            intCurrentSeries = intCurrentSeries + 1
            intViewerIndex = intViewerIndex + 1
            '判断这个Viwer是否存在，如果已经存在，则重新装载这个Viwer中的图像
            If intViewerIndex >= thisForm.Viewer.Count Then  ''Viewer不存在，则创建Viewer
                '创建并摆放一个Viewer
                intCurrentViewer = funcCeateAViewer(intCurrentSeries, thisForm)
                
                '摆放这个Viewer并设置滚动条
                Call subPlaceAViewer(thisForm, intCurrentViewer, i, j)
            End If
            
            
        Next j
        If blnLoadOver = True Then
            Exit For
        End If
    Next i
    
    '设置默认被选择的序列和图像
    If thisForm.Viewer.Count > 1 Then
        If thisForm.Viewer(1).Images.Count > 0 Then
            Set thisForm.SelectedImage = thisForm.Viewer(1).Images(1)
            thisForm.SelectedImageIndex = 1
        End If
        thisForm.intSelectedSerial = 1
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subSetSeriesLayout(thisForm As frmViewer, strModality As String, intSeriesCount As Integer)
'------------------------------------------------
'功能：根据影像类别和序列数量来排布序列布局
'参数： thisForm--进行序列布局重排的窗体
'       strModality--进行序列布局重排的影像类别
'       intSeriesCount--序列总数量
'返回：无，直接重排指定序列的布局
'时间：2009-7
'------------------------------------------------
    Dim i As Integer
    Dim intRows As Integer
    Dim intCols As Integer
    Dim blnSetSeriesLayout As Boolean
    
    For i = 1 To UBound(aPresetLayout)
        If UCase(aPresetLayout(i).strModality) = UCase(strModality) Then
            If aPresetLayout(i).bSeriesAutoFormat = True Then
                '是自动布局，根据序列的总数量和最大布局数量，设置图像布局
                ResizeRegion intSeriesCount, thisForm.width, thisForm.height, intRows, intCols, intMaxAreaY, intMaxAreaX
                thisForm.intCountX = intCols
                thisForm.intCountY = intRows
            Else
                thisForm.intCountX = aPresetLayout(i).lngSeriesColumns
                thisForm.intCountY = aPresetLayout(i).lngSeriesRows
            End If
            blnSetSeriesLayout = True
        End If
    Next i
    
    If blnSetSeriesLayout = False Then
        '设置默认的序列布局
        thisForm.intCountX = 2
        thisForm.intCountY = 2
    End If
End Sub

Public Sub subPlaceAViewer(thisForm As frmViewer, intViewerIndex As Integer, intRow As Integer, intCol As Integer)
'------------------------------------------------
'功能：把Viewer放在指定位置，并显示滚动条和图像选择框
'参数： thisForm--观片窗体
'       intViewerIndex--需要放置的Viewer和滚动条的Index
'       intRow -- 摆放Viewer 的行
'       intCol -- 摆放Viewer 的列
'返回：无，摆放Viewer并且摆放滚动条
'时间：2009-7
'------------------------------------------------
    Dim lngTop As Long
    Dim lngLeft As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngOldWidth As Long
    Dim lngOldHeight As Long
    
    If intRow = 0 Or intCol = 0 Then Exit Sub
    On Error GoTo err
    
    '根据intRow和intCol 计算出来Viewer需要摆放的位置
    With thisForm
        If intCol = 1 Then  ''''计算当前viewer的横向位置
            lngLeft = 0
            lngWidth = .PicX(intCol).left
        Else
            lngLeft = .PicX(intCol - 1).left + intSpaceSize
            If intCol = intMaxAreaX Then
                lngWidth = .picViewer.ScaleWidth - .PicX(intCol - 1).left - intSpaceSize
            Else
                If .PicX(intCol).left - .PicX(intCol - 1).left - intSpaceSize < 0 Then
                    lngWidth = 0
                Else
                    lngWidth = .PicX(intCol).left - .PicX(intCol - 1).left - intSpaceSize
                End If
            End If
        End If
        '''''''''''''''''''''''''''''''''''''''''''计算当前viewer的纵向位置
        If intRow = 1 Then
            lngTop = 0
            lngHeight = .PicY(intRow).top
        Else
            lngTop = .PicY(intRow - 1).top + intSpaceSize
            If intRow = intMaxAreaY Then
                lngHeight = .picViewer.ScaleHeight - .PicY(intRow - 1).top - intSpaceSize
            Else
                If .PicY(intRow).top - .PicY(intRow - 1).top - intSpaceSize < 0 Then
                    lngHeight = 0
                Else
                    lngHeight = .PicY(intRow).top - .PicY(intRow - 1).top - intSpaceSize
                End If
            End If
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If lngHeight < 0 Then lngHeight = 0
    End With
    
    '摆放并显示Viewer
    lngOldWidth = thisForm.Viewer(intViewerIndex).width / thisForm.Viewer(intViewerIndex).MultiColumns
    lngOldHeight = thisForm.Viewer(intViewerIndex).height / thisForm.Viewer(intViewerIndex).MultiRows
    thisForm.Viewer(intViewerIndex).Move lngLeft, lngTop, Abs(lngWidth), Abs(lngHeight)
    thisForm.Viewer(intViewerIndex).Visible = (lngWidth <> 0)
    ZLShowSeriesInfos(intViewerIndex).intRow = intRow
    ZLShowSeriesInfos(intViewerIndex).intCol = intCol
    
    '如果图像不是StretchToFit，则调整Viewer中图像的位置
    If thisForm.Viewer(intViewerIndex).Images.Count > 0 Then
        Call subScaleViewer(thisForm.Viewer(intViewerIndex), thisForm.Viewer(intViewerIndex).Images(1), lngOldWidth, lngOldHeight)
    End If
    
    '判断滚动条是否需要显示，如果需要则显示滚动条，设置滚动条的最大值，最小值，LarghChange等
    subDisplayScrollBar intViewerIndex, thisForm, True
    
    '如果是选择所有序列的状态，将目前序列设置为当前序列，供subDispframe使用
    If thisForm.isSelectAllSerial Then thisForm.intSelectedSerial = intViewerIndex
    
    '图像显示完后，添加Viewer中的标注：图象框、右下角的选择标记等
    Call subDispframe(thisForm, thisForm.Viewer(intViewerIndex))
    
    '自动根据图像大小，判断是否显示病人四角信息,显示或者隐藏图像中的病人信息
    Call subDisplayPatientInfo(thisForm.Viewer(intViewerIndex))
Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subShowSpliter(thisForm As frmViewer)
'------------------------------------------------
'功能：根据序列布局，重新显示分隔条
'参数： thisForm--观片窗体
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim intRows As Integer
    Dim intCols As Integer
    Dim lngAreaWidth As Long
    Dim lngAreaHeight As Long
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    On Error GoTo err
    '如果是观察模式，则显示1*1的布局
    If Button_miLookOrBrowse = True Then '观察模式
        intRows = 1
        intCols = 1
    Else    '如果是浏览模式，则按照intCountX和intCountY的定义显示序列
        intRows = thisForm.intCountY
        intCols = thisForm.intCountX
    End If
    
    '计算并且摆放分隔条
    With thisForm
        If intCols = intMaxAreaX Then   ''计算横向可用宽度，程序中已经确保了intCols不会大于intMaxAreaX
            lngAreaWidth = .picViewer.ScaleWidth - intSpaceSize * (intMaxAreaX - 1)
        Else
            lngAreaWidth = .picViewer.ScaleWidth - intSpaceSize * intCols
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If intRows = intMaxAreaY Then   ''计算横向可用宽度
            lngAreaHeight = .picViewer.ScaleHeight - intSpaceSize * (intMaxAreaY - 1)
        Else
            lngAreaHeight = .picViewer.ScaleHeight - intSpaceSize * intRows
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To intMaxAreaX - 1  ''横向所有的分隔线归位
            .PicX(i).left = .picViewer.ScaleWidth - intSpaceSize
            .PicX(i).Tag = ""
            .PicX(i).height = .picViewer.ScaleHeight
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To intMaxAreaY - 1 ''纵向向所有的分隔线归位
            .PicY(i).top = .picViewer.ScaleHeight - intSpaceSize
            .PicY(i).Tag = ""
            .PicY(i).width = .picViewer.ScaleWidth
        Next
        
        '调整拖动条的宽度和高度
        .PicXX.height = .picViewer.ScaleHeight
        .PicYY.width = .picViewer.ScaleWidth
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To intCols - 1  ''横向需要显示的线位置计算
            .PicX(i).left = lngAreaWidth / intCols * i + intSpaceSize * (i - 1)
            .PicX(i).Tag = .PicX(i).left
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To intRows - 1  ''纵向需要显示的线位置计算
            .PicY(i).top = lngAreaHeight / intRows * i + intSpaceSize * (i - 1)
            .PicY(i).Tag = .PicY(i).top
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To intMaxAreaX - 1  ''排布双向分隔点的位置
            For j = 1 To intMaxAreaY - 1
                k = (j - 1) * (intMaxAreaX - 1) + i
                .PicXY(k).top = .PicY(j).top
                .PicXY(k).left = .PicX(i).left
            Next
        Next
    End With
Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetImageAttribute(objAttr As DicomAttributes, ByVal AttrName As String) As String
'-----------------------------------------------------------------------------
'功能:提取DICOM属性集中的指定属性值,根据VM判断值的维度，使用“\”把各个维度连接成一个串
'-----------------------------------------------------------------------------
    Dim AttrTag() As String
    Dim i As Integer
    
    GetImageAttribute = ""
    AttrTag = Split(AttrName, ":")
    If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Exists Then
        If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).VM = 1 Then
            GetImageAttribute = Nvl(objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).ValueByIndex(1))
        Else
            For i = 1 To objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).VM
                GetImageAttribute = GetImageAttribute & "\" & objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).ValueByIndex(i)
            Next i
        End If
    End If
End Function

Public Sub subDisplayScrollBar(intViewerIndex As Integer, thisForm As frmViewer, blnResizeViewer As Boolean)
'------------------------------------------------
'功能：根据图像的数量，判断显示或者隐藏滚动条
'参数： intViewerIndex--滚动条的索引
'       thisForm --  观片窗体
'       blnResizeViewer ---TrueViewer的宽高重新调整过了。
'返回：无 ，直接显示或隐藏滚动条。
'时间：2009-7
'------------------------------------------------
    Dim thisViewer As DicomViewer
    Dim thisVScro As VScrollBar
    Dim lngImageCount As Long
    
    On Error GoTo err
    '找到Viewer所对应的序列索引
    Set thisViewer = thisForm.Viewer(intViewerIndex)
    Set thisVScro = thisForm.VScro(intViewerIndex)

    If thisViewer.Images.Count = 0 Then Exit Sub
    
    lngImageCount = ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
    
    '判断Viewer所对应的序列中图像的数量是否大于显示数，大于则显示滚动条
    If lngImageCount > thisViewer.MultiColumns * thisViewer.MultiRows Then  '显示滚动条
        '判断滚动条的显示状态是否发生变化,摆放滚动条，并显示滚动条
        If blnResizeViewer = True Or thisVScro.Visible = False Then    '调整Viewer的宽度，将占用Viewer的部分空间
            thisVScro.Move thisViewer.left + thisViewer.width - thisVScro.width, thisViewer.top, thisVScro.width, thisViewer.height
            thisViewer.width = Abs(thisViewer.width - thisVScro.width)
        Else
            thisVScro.Move thisViewer.left + thisViewer.width, thisViewer.top, thisVScro.width, thisViewer.height
        End If
        thisVScro.Visible = thisViewer.Visible
        thisVScro.ZOrder
        thisVScro.Refresh
        '设置滚动条的最大，最小值
        thisVScro.Min = 1
        thisVScro.Max = lngImageCount - thisViewer.MultiColumns * thisViewer.MultiRows + 1
        If thisVScro.Max < 1 Then thisVScro.Max = 1
        thisVScro.LargeChange = thisViewer.MultiColumns * thisViewer.MultiRows
        If thisViewer.CurrentIndex > thisVScro.Max Then
            thisVScro.Value = thisVScro.Max
            thisViewer.CurrentIndex = thisVScro.Max
        Else
            thisVScro.Value = thisViewer.CurrentImage.Tag
        End If
    Else    '图像少于可显示的数量，则隐藏滚动条
        If blnResizeViewer = False And thisVScro.Visible = True Then    '调整Viewer的宽度
            thisViewer.width = thisViewer.width + thisVScro.width
        End If
        thisVScro.Visible = False
        '只有当前没有选定图像时才设置当前图像，否则会对MPR等操作造成影响
        If thisForm.SelectedImage Is Nothing Then
            thisForm.SelectedImageIndex = thisViewer.CurrentIndex
            Set thisForm.SelectedImage = thisViewer.CurrentImage
            thisForm.intSelectedSerial = thisViewer.Index
            thisForm.MSFViewer.TextMatrix(thisViewer.Index, 3) = thisForm.SelectedImageIndex
        End If
    End If
    
    '处理弹出的窗宽窗位菜单
    Call subSetWidthLevelF(thisForm.SelectedImage, thisForm)
    '处理弹出的图像滤镜菜单
    Call subSetFilterF(thisForm.SelectedImage, thisForm)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subShowMiniImages(thisForm As frmViewer)
'------------------------------------------------
'功能：显示或关闭序列缩略图
'参数： intViewerIndex--滚动条的索引
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim i As Integer
    Dim imgs As New DicomImages
    Dim img As DicomImage
    
    On Error GoTo err
    
    '根据工具栏变量判断是否显示缩略图
    If Button_miShowMiniSeries = True Then      '显示缩略图
        For i = 1 To ZLSeriesInfos.Count
            '加载一个图象
            Set img = funLoadAImage(i, 1, 0)
            If Not img Is Nothing Then
                imgs.Add img
            End If
        Next i
        If blnDockMiniImage = True Then
            frmMiniSeries.ShowMe imgs, thisForm, thisForm.dkpMain
        Else
            frmMiniSeries.ShowMe imgs, thisForm
        End If
    Else        '隐藏缩略图
        If blnDockMiniImage = True Then
            frmMiniSeries.CloseMe thisForm.dkpMain
        Else
            frmMiniSeries.CloseMe
        End If
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function funLoadAImage(intSeriesIndex As Integer, intImageIndex As Integer, intLoadType As Integer) As DicomImage
'------------------------------------------------
'功能： 把指定序列和图像索引的图像添加到funLoadAImage中
'       加入图像的时候，查找图像的方式是：首先查看本地缓存---再提取共享目录---最后通过FTP下载
'参数： intViewerIndex--图像所在的序列的索引，具体从哪里读取序列，跟intLoadType相关
'       intImageIndex--图像所在的图像的索引
'       intLoadType -- 装载模式。0--从ZLSeriesInfos装载，1 -- 从ZLShowSeriesInfos装载
'返回：加载图象成功，则返回图像，否则返回NOTHING
'时间：2009-7
'------------------------------------------------
    Dim thisImage As DicomImage
    Dim adata As DicomDataSet
    Dim attr As DicomAttribute
    Dim strInstanceUID As String
    Dim strSeriesUID As String
    Dim intPRISeriesIndex As Integer
    Dim intPRIImageIndex As Integer
    Dim i As Integer
    Dim dssSub1 As DicomDataSets
    Dim dssSub2 As DicomDataSets
    Dim dssSub7060 As DicomDataSets
    Dim dssSub701 As DicomDataSets
    Dim dssSub283110 As DicomDataSets
    Dim dssSub705A As DicomDataSets
    Dim dsSub70601 As DicomDataSet
    Dim dssub70602 As DicomDataSet
    
    Dim OriginImage As DicomImage
    
    
    On Error GoTo err
    
    Set thisImage = funLoadOneImage(intSeriesIndex, intImageIndex, intLoadType)
    Set funLoadAImage = thisImage
    If funLoadAImage Is Nothing Then
        Debug.Print "图像读取不到"
        Exit Function
    End If
    
    '增加PR图像信息
    If UCase(Nvl(thisImage.Attributes(&H8, &H60).Value, "OT")) = "PR" Then
        '查找PR图像对应的原始图像
        '提取PR图像中的序列UID和图像UID
        If thisImage.Attributes(&H8, &H1115).Exists = True Then
            Set dssSub1 = thisImage.Attributes(&H8, &H1115).Value
            If dssSub1(1).Attributes(&H8, &H1140).Exists Then
                Set dssSub2 = dssSub1(1).Attributes(&H8, &H1140).Value
                    If dssSub2(1).Attributes(&H8, &H1155).Exists = True Then
                        strInstanceUID = dssSub2(1).Attributes(&H8, &H1155).Value
                    End If
            End If
            If dssSub1(1).Attributes(&H20, &HE).Exists = True Then
                strSeriesUID = dssSub1(1).Attributes(&H20, &HE).Value
            End If
        End If

        '如果序列UID或者图像UID为空，则退出
        If strSeriesUID = "" Or strInstanceUID = "" Then
            Exit Function
        End If

        '查找PR对应的原始图
        For i = 1 To ZLSeriesInfos.Count
            If ZLSeriesInfos(i).SeriesUID = strSeriesUID Then
                intPRISeriesIndex = i
                Exit For
            End If
        Next i
        If intPRISeriesIndex > ZLSeriesInfos.Count Or intPRISeriesIndex <= 0 Then
            Exit Function
        End If

        For i = 1 To ZLSeriesInfos(intPRISeriesIndex).ImageInfos.Count
            If ZLSeriesInfos(intPRISeriesIndex).ImageInfos(i).InstanceUID = strInstanceUID Then
                intPRIImageIndex = i
                Exit For
            End If
        Next i
        If intPRIImageIndex > ZLSeriesInfos(intPRISeriesIndex).ImageInfos.Count Then
            Exit Function
        End If

        '提取原始图像的信息
        '加载原始图像
        Set OriginImage = funLoadOneImage(intPRISeriesIndex, intPRIImageIndex, 0)

        '提取PR图像的信息
        Set adata = New DicomDataSet
        
       '正常读取PR所有信息
        On Error Resume Next
        For Each attr In thisImage.Attributes
            adata.Attributes.Add attr.Group, attr.Element, attr.Value
        Next
        
        '安科公司产生的PR图像，没有明确的（70,2）名称，而且最后（70,60）中也没有内容，
        '导致标注显示不出来，因此需要特殊处理
'       '第一层Text
        Set dssSub1 = adata.Attributes(&H70, &H1).Value
        If IsNull(dssSub1(1).Attributes(&H70, &H2).Value) Then
            dssSub1(1).Attributes.Add &H70, &H2, "LAYER1"
        End If

        '增加 Graphic Layer Sequence
        If IsNull(adata.Attributes(&H70, &H60).Value) Then
            Set dssSub7060 = New DicomDataSets
            Set dsSub70601 = New DicomDataSet
            dsSub70601.Attributes.Add &H70, &H2, "LAYER1"
            dsSub70601.Attributes.Add &H70, &H62, 1
            dsSub70601.Attributes.Add &H70, &H68, "layer1"
            dssSub7060.Add dsSub70601
            adata.Attributes.Add &H70, &H60, dssSub7060
        End If
        
        Set OriginImage.PresentationState = adata

        '修改一些图像中的必要信息，其中图像的InstanceUID，SeriesUID，StudyUID都不能修改，否则PR图像显示不正常

        OriginImage.Name = thisImage.Name
        OriginImage.AccessionNumber = thisImage.AccessionNumber
        OriginImage.PatientID = thisImage.PatientID
        OriginImage.SeriesDescription = thisImage.SeriesDescription
        OriginImage.StudyDescription = thisImage.StudyDescription
        If thisImage.Attributes(&H20, &H10).Exists Then 'study id
            OriginImage.Attributes.Add &H20, &H10, thisImage.Attributes(&H20, &H10).Value
        End If
        If thisImage.Attributes(&H20, &H11).Exists Then 'series number
            OriginImage.Attributes.Add &H20, &H11, thisImage.Attributes(&H20, &H11).Value
        End If
        If thisImage.Attributes(&H20, &H13).Exists Then 'image number
            OriginImage.Attributes.Add &H20, &H13, thisImage.Attributes(&H20, &H13).Value
        End If
        If thisImage.Attributes(&H8, &H60).Exists Then  'modality
            OriginImage.Attributes.Add &H8, &H60, thisImage.Attributes(&H8, &H60).Value
        End If
        If thisImage.Attributes(&H8, &H20).Exists Then  'study date
            OriginImage.Attributes.Add &H8, &H20, thisImage.Attributes(&H8, &H20).Value
        End If
        If thisImage.Attributes(&H8, &H30).Exists Then  'study time
            OriginImage.Attributes.Add &H8, &H30, thisImage.Attributes(&H8, &H30).Value
        End If
        If thisImage.Attributes(&H28, &H3110).Exists Then   '窗宽窗位
            Set dssSub283110 = thisImage.Attributes(&H28, &H3110).Value
            If dssSub283110(1).Attributes(&H28, &H1050).Exists Then
                If Not IsNull(dssSub283110(1).Attributes(&H28, &H1050).Value) Then
                    OriginImage.Attributes.Add &H28, &H1050, dssSub283110(1).Attributes(&H28, &H1050).Value
                    OriginImage.Level = dssSub283110(1).Attributes(&H28, &H1050).ValueByIndex(1)
                End If
            End If
            If dssSub283110(1).Attributes(&H28, &H1051).Exists Then
                If Not IsNull(dssSub283110(1).Attributes(&H28, &H1051).Value) Then
                    OriginImage.Attributes.Add &H28, &H1051, dssSub283110(1).Attributes(&H28, &H1051).Value
                    OriginImage.width = dssSub283110(1).Attributes(&H28, &H1051).ValueByIndex(1)
                End If
            End If
        End If

        Set funLoadAImage = OriginImage

    End If
     
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function funLoadOneImage(intSeriesIndex As Integer, intImageIndex As Integer, intLoadType As Integer) As DicomImage
'------------------------------------------------
'功能： 把指定序列和图像索引的图像添加到funLoadOneImage中
'       加入图像的时候，查找图像的方式是：首先查看本地缓存---再提取共享目录---最后通过FTP下载
'参数： intViewerIndex--图像所在的序列的索引，具体从哪里读取序列，跟intLoadType相关
'       intImageIndex--图像所在的图像的索引
'       intLoadType -- 装载模式。0--从ZLSeriesInfos装载，1 -- 从ZLShowSeriesInfos装载
'返回：加载图象成功，则返回图像，否则返回NOTHING
'时间：2009-7
'------------------------------------------------
    Dim imgs As New DicomImages
    Dim strSaveDir As String
    Dim lngSource As Long
    Dim strShareDir As String
    Dim strHostIP As String
    Dim strImageName As String
    Dim strLocalImage As String
    Dim strFTPUser As String
    Dim strFTPPaswd As String
    Dim strFTPDir As String
    Dim cFTP As New clsFtp
    Dim lngResult As Long
    Dim lngLocalFileSize As Long
    Dim lngFTPFileSize As Long
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim StrMessage As String
    
    On Error GoTo err
    Set funLoadOneImage = Nothing
    
    If intLoadType = 0 Then         '0--从ZLSeriesInfos装载
        '先判断图像索引是否正确
        If ZLSeriesInfos.Count < intSeriesIndex Then Exit Function
        If ZLSeriesInfos(intSeriesIndex).ImageInfos.Count < intImageIndex Then Exit Function
        
        strSaveDir = ZLSeriesInfos(intSeriesIndex).strSaveDir
        strImageName = ZLSeriesInfos(intSeriesIndex).ImageInfos(intImageIndex).ImageName
        lngSource = ZLSeriesInfos(intSeriesIndex).lngSource
        strShareDir = ZLSeriesInfos(intSeriesIndex).strShareDir
        strHostIP = ZLSeriesInfos(intSeriesIndex).strHostIP
        strFTPUser = ZLSeriesInfos(intSeriesIndex).strFTPUser
        strFTPPaswd = ZLSeriesInfos(intSeriesIndex).strFTPPasw
        strFTPDir = ZLSeriesInfos(intSeriesIndex).strFTPDir
    ElseIf intLoadType = 1 Then     '1 -- 从ZLShowSeriesInfos装载
        '先判断图像索引是否正确
        If ZLShowSeriesInfos.Count < intSeriesIndex Then Exit Function
        If ZLShowSeriesInfos(intSeriesIndex).ImageInfos.Count < intImageIndex Then Exit Function
        
        strSaveDir = ZLShowSeriesInfos(intSeriesIndex).strSaveDir
        strImageName = ZLShowSeriesInfos(intSeriesIndex).ImageInfos(intImageIndex).ImageName
        lngSource = ZLShowSeriesInfos(intSeriesIndex).lngSource
        strShareDir = ZLShowSeriesInfos(intSeriesIndex).strShareDir
        strHostIP = ZLShowSeriesInfos(intSeriesIndex).strHostIP
        strFTPUser = ZLShowSeriesInfos(intSeriesIndex).strFTPUser
        strFTPPaswd = ZLShowSeriesInfos(intSeriesIndex).strFTPPasw
        strFTPDir = ZLShowSeriesInfos(intSeriesIndex).strFTPDir
    End If
    
    If lngSource = 0 Then
        strLocalImage = PstrBufferImagePath & strSaveDir & "\" & strImageName
        Call MkLocalDir(PstrBufferImagePath & strSaveDir)
    Else
        strLocalImage = strSaveDir & "\" & strImageName
    End If
    
    If Dir(strLocalImage) <> vbNullString Then
        '从本机缓存目录中读取文件
        Set funLoadOneImage = ReadImage(strLocalImage, True)
        If funLoadOneImage Is Nothing Then
            Debug.Print "打开错误，可能正被占用"
            If FileIsOccupied(strLocalImage) = True Then
                Debug.Print Now & " 文件正被占用"
                '延时
                TimeDelay 2000
                Debug.Print Now
                
                Set funLoadOneImage = ReadImage(strLocalImage, True)
                If funLoadOneImage Is Nothing Then
                    '延时
                    TimeDelay 2000
                    Debug.Print Now
                    
                    Set funLoadOneImage = ReadImage(strLocalImage, True)
                    If funLoadOneImage Is Nothing Then
                        '延时
                        TimeDelay 2000
                        Debug.Print Now
                        
                        Set funLoadOneImage = ReadImage(strLocalImage, True)
                        If funLoadOneImage Is Nothing Then
                            Debug.Print "goto errdown"
                            GoTo errDown
                        End If
                    End If
                End If
                Debug.Print "延时读取成功"
            Else
                Debug.Print "打开出错，重新下载 goto errdown"
                GoTo errDown
            End If
        End If
    Else
errDown:
        If strShareDir <> "" Then
            '通过共享目录读取文件
            Set funLoadOneImage = imgs.ReadFile("\\" & strHostIP & "\" & strShareDir & "\" & strSaveDir & "\" & strImageName)
        Else
            '通过FTP下载文件
            '连接FTP
            cFTP.FuncFtpConnect strHostIP, strFTPUser, strFTPPaswd
ReDownFile:
            '下载文件
            lngResult = cFTP.FuncDownloadFile(strFTPDir & Replace(strSaveDir, "\", "/"), strLocalImage, strImageName)
            '下载成功后，对比本地文件和FTP文件大小是否一致
            If lngResult = 0 And gblnCompareSize Then
                lngLocalFileSize = objFileSystem.GetFile(strLocalImage).Size
                lngFTPFileSize = cFTP.FuncFtpGetFileSize(strFTPDir & Replace(strSaveDir, "\", "/"), strImageName)
                
                If lngLocalFileSize < lngFTPFileSize Then
                    StrMessage = "下载后的文件大小【" & lngLocalFileSize & "】与FTP中的文件大小【" & lngFTPFileSize & "】不一致，" & vbCrLf & _
                                 "本地文件：" & strLocalImage & vbCrLf & _
                                 "FTP文件：" & strFTPDir & Replace(strSaveDir, "\", "/") & strImageName & vbCrLf & _
                                 "是否需要重新下载？"
                    If MsgBox(StrMessage, vbQuestion + vbYesNo, "提示") = vbYes Then
                        GoTo ReDownFile
                    End If
                End If
            End If
            cFTP.FuncFtpDisConnect
            
            Debug.Print "lngResult = " & lngResult
            
            If lngResult = 0 Then  '当前图像下载成功
                Set funLoadOneImage = imgs.ReadFile(strLocalImage)
            End If
        End If
    End If
     
    cFTP.FuncFtpDisConnect
    Exit Function
err:
    cFTP.FuncFtpDisConnect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub subMoveViewers(thisForm As frmViewer, intRow As Integer, intCol As Integer)
'------------------------------------------------
'功能： 批量调整Viwer的位置，调整intRow下方，intCol右方的所有Viewer的位置
'参数： thisForm--观片站窗体
'       intRow--Viwer所在的行
'       intCol -- Viewer所在的列
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim i As Integer
    
    For i = 1 To thisForm.Viewer.Count - 1
        If ZLShowSeriesInfos(i).ImageInfos.Count <> 0 Then
            If ZLShowSeriesInfos(i).intRow >= intRow And ZLShowSeriesInfos(i).intCol >= intCol Then
                Call subPlaceAViewer(thisForm, i, ZLShowSeriesInfos(i).intRow, ZLShowSeriesInfos(i).intCol)
            End If
        End If
    Next i
End Sub


Public Sub subOpenFiles(thisForm As frmViewer)
'------------------------------------------------
'功能： 弹出打开文件窗体，打开图像文件
'参数： thisForm--观片站窗体
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim arrFileList As OpenFileArray
    
    On Error GoTo err
    
    arrFileList = funGetFileList(thisForm)
    '如果有内容，则打开文件串
    If arrFileList.FilePath <> "" Then
        Call subOpenFileList(thisForm, arrFileList)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subOpenFileList(thisForm As frmViewer, arrFileList As OpenFileArray)
'------------------------------------------------
'功能： 打开文件列表中的图像
'参数： thisForm -- 观片站窗体
'参数： arrFileList -- 要打开的文件列表
'返回：无
'------------------------------------------------
    Dim img As New DicomImage
    Dim iFileIndex As Integer
    Dim intSeriesIndex As Integer
    Dim intImageIndex As Integer
    Dim blnNewSeries As Boolean
    Dim blnNewImage As Boolean
    Dim iPrivImage As Integer
    Dim lngImageNo As Long
    Dim oneSeriesInfo As clsSeriesInfo
    Dim oneImageInfo As clsImageInfo
    Dim i As Integer
    Dim j As Integer

    On Error GoTo err
    
    If arrFileList.FilePath = "" Then Exit Sub
    
    
    '把打开的文件串加载到        ZLSeriesInfos 结构中
    For iFileIndex = 1 To UBound(arrFileList.Filename)
        Set img = ReadImage(arrFileList.FilePath & arrFileList.Filename(iFileIndex))
        
        If Not IsNull(img.Attributes(&H20, &H13).Value) Then
            lngImageNo = Val(img.Attributes(&H20, &H13).Value)
        Else
            lngImageNo = 0
        End If
        '根据序列UID，查找序列是否存在
        blnNewSeries = True
        For intSeriesIndex = 1 To ZLSeriesInfos.Count
            If ZLSeriesInfos(intSeriesIndex).MultiFrame = 1 Then
                '不做任何操作，多帧序列不再增加图像，但是对于多帧图需要判断是否跟当前图像是同一个图
                If img.FrameCount > 1 And ZLSeriesInfos(intSeriesIndex).SeriesUID = img.SeriesUID _
                    And ZLSeriesInfos(intSeriesIndex).ImageInfos(1).InstanceUID = img.InstanceUID Then
                    '同一个图，标记使用同一个序列，后面处理图像的时候会判断不加载图像
                    blnNewSeries = False
                    Exit For
                End If
            Else
                '单帧序列才考虑把图像添加到序列中
                If ZLSeriesInfos(intSeriesIndex).SeriesUID = img.SeriesUID And img.FrameCount = 1 Then
                    blnNewSeries = False
                    Exit For
                End If
            End If
        Next intSeriesIndex
        
        If blnNewSeries = True Then '创建新序列
            '创建新序列
            Set oneSeriesInfo = funGetNewSeriesInfo
            oneSeriesInfo.lngSource = 1     '直接通过打开方式加载的序列
            oneSeriesInfo.SeriesNo = GetImageAttribute(img.Attributes, ATTR_序列号)
            oneSeriesInfo.SeriesUID = img.SeriesUID
            oneSeriesInfo.strModality = GetImageAttribute(img.Attributes, ATTR_影像类别)
            oneSeriesInfo.strSaveDir = arrFileList.FilePath
            oneSeriesInfo.StudyUID = img.StudyUID
            oneSeriesInfo.MultiFrame = IIf(img.FrameCount = 1, 0, 1)
            
            '读取预设窗宽窗位
            For i = 1 To UBound(aPresetWinWL, 2)
                If UCase(aPresetWinWL(3, i).strModality) = UCase(oneSeriesInfo.strModality) Then
                    For j = 3 To 12
                        If aPresetWinWL(j, i).bInUse And aPresetWinWL(j, i).intDefault = 1 Then
                            oneSeriesInfo.lngWinWidth = aPresetWinWL(j, i).lngWinWidth
                            oneSeriesInfo.lngWinLevel = aPresetWinWL(j, i).lngWinLevel
                            Exit For
                        End If
                    Next j
                    Exit For
                End If
            Next i
            
            ZLSeriesInfos.Add oneSeriesInfo, CStr(ZLSeriesInfos.Count + 1)
            
            '填写图像位置
            iPrivImage = 1
            blnNewImage = True
        Else    '使用原有序列
            '根据图像UID，判断图像是否存在，如果图像存在，则同时查找图像位置
            blnNewImage = True
            iPrivImage = 0
            For intImageIndex = 1 To ZLSeriesInfos(intSeriesIndex).ImageInfos.Count
                If ZLSeriesInfos(intSeriesIndex).ImageInfos(intImageIndex).ImageNo < lngImageNo Then
                    iPrivImage = intImageIndex
                End If
                
                If img.InstanceUID = ZLSeriesInfos(intSeriesIndex).ImageInfos(intImageIndex).InstanceUID Then
                    blnNewImage = False
                    Exit For
                End If
            Next intImageIndex
        End If
        
        '添加图像
        If blnNewImage = True Then
            '创建图像
            Set oneImageInfo = funGetNewImageInfo
            oneImageInfo.AcquisitionTime = Format(GetImageAttribute(img.Attributes, ATTR_采集日期) & " " & GetImageAttribute(img.Attributes, ATTR_采集时间), "yyyy-MM-dd HH:MM:SS")
            oneImageInfo.Columns = img.sizeX
            oneImageInfo.FrameOfReferenceUID = GetImageAttribute(img.Attributes, ATTR_参考帧UID)
            oneImageInfo.ImageName = arrFileList.Filename(iFileIndex)
            oneImageInfo.ImageNo = GetImageAttribute(img.Attributes, ATTR_图像号)
            oneImageInfo.ImageOrientationPatient = GetImageAttribute(img.Attributes, ATTR_图像方向病人)
            oneImageInfo.ImagePositionPatient = GetImageAttribute(img.Attributes, ATTR_图像位置病人)
            oneImageInfo.ImageTime = Format(GetImageAttribute(img.Attributes, ATTR_图像日期) & " " & GetImageAttribute(img.Attributes, ATTR_图像时间), "yyyy-MM-dd HH:MM:SS")
            oneImageInfo.InstanceUID = img.InstanceUID
            oneImageInfo.PixelSpacing = GetImageAttribute(img.Attributes, ATTR_像素距离)
            oneImageInfo.Rows = img.sizeY
            oneImageInfo.SliceLocation = GetImageAttribute(img.Attributes, ATTR_切片位置)
            oneImageInfo.SliceThickness = GetImageAttribute(img.Attributes, ATTR_层厚)
            
            If iPrivImage = 0 Then iPrivImage = 1   '在第一个图的前面追加
            If ZLSeriesInfos(intSeriesIndex).ImageInfos.Count = 0 Then
                ZLSeriesInfos(intSeriesIndex).ImageInfos.Add oneImageInfo
            Else
                ZLSeriesInfos(intSeriesIndex).ImageInfos.Add oneImageInfo, , , iPrivImage
            End If
        End If
    Next iFileIndex
    
    '把需要显示的图像加载进入VIEWER,把Viewer显示到窗口中，显示Viewer相关的滚动条
    Call subShowForm(thisForm)
    '重新显示缩略图
    Call subShowMiniImages(thisForm)
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subLookOrBrowsSwitch(thisForm As frmViewer)
'------------------------------------------------
'功能： 切换浏览和观察模式,重新摆放分隔条，重新摆放窗体
'       进入观察模式时，只显示当前被选中的Viewer，隐藏其他的Viewer，Visible=False。
'       进入浏览模式时，循环当前已经存在的Viewer，把他们的Visible设为True，并重新摆放这些Viewer。
'       切换浏览和观察模式，并不删除Viewer和改变ZLShowSeriesInfos结构，只是显示和隐藏Viewer而已。
'       只有在切换布局的时候，才会改变ZLShowSeriesInfos结构。
'参数： thisForm--观片站窗体
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim blnExit As Boolean
    Dim intViwerIndex As Integer
    
    If thisForm.intSelectedSerial < 1 Then Exit Sub
    If thisForm.intSelectedSerial >= thisForm.Viewer.Count Then Exit Sub
    On Error GoTo err
    
    Button_miLookOrBrowse = Not Button_miLookOrBrowse
    thisForm.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_OneBrowse, , True).Checked = Button_miLookOrBrowse
    thisForm.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_OneBrowse, , True).Checked = Button_miLookOrBrowse
    thisForm.ComToolBar.RecalcLayout
            
    '摆放分隔条
    Call subShowSpliter(thisForm)
    
    If Button_miLookOrBrowse = True Then    '观察模式，只摆放一个Viewer
        '隐藏其他的Viewer
        For i = 1 To thisForm.Viewer.Count - 1
            If i = thisForm.intSelectedSerial Then
                thisForm.Viewer(i).Visible = True
            Else
                thisForm.Viewer(i).Visible = False
                thisForm.VScro(i).Visible = False
            End If
        Next i
        '摆放这个Viewer并设置滚动条
        Call subPlaceAViewer(thisForm, thisForm.intSelectedSerial, 1, 1)
    Else    '浏览模式，摆放所有Viewer
        '循环所有序列布局，摆放所有的Viewer
        intViwerIndex = 0
        blnExit = False
        For i = 1 To thisForm.intCountY
            For j = 1 To thisForm.intCountX
                intViwerIndex = intViwerIndex + 1
                If intViwerIndex >= thisForm.Viewer.Count Then
                    blnExit = True
                    Exit For
                End If
                '摆放Viewer，并显示Viewer
                thisForm.Viewer(intViwerIndex).Visible = True
                Call subPlaceAViewer(thisForm, intViwerIndex, i, j)
            Next j
            If blnExit = True Then
                Exit For
            End If
        Next i
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub subDisplayPatientInfo(thisViewer As DicomViewer)
'------------------------------------------------
'功能： 显示或关闭指定Viewer的病人信息，并且根据图像大小决定是否自动隐藏病人信息
'参数： thisViewer--需要显示或着关闭病人信息的Viewer
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim img As DicomImage
    Dim i As Integer
    Dim j As Integer
    Dim blnShowLabel As Boolean
    Dim intImageIndex As Integer
    
    On Error GoTo err
    If thisViewer.Images.Count < 1 Then Exit Sub
    
    'Button_miDispPatientInfo    ---属性显示
    'blnpatientInfoScaleFontSize ---病人信息文字大小是否随着图像一起缩放
    
    If ((Not blnpatientInfoScaleFontSize) And _
            (thisViewer.width / thisViewer.MultiColumns / Screen.TwipsPerPixelX < lngPatientInfoInvisibleSize Or _
             thisViewer.height / thisViewer.MultiRows / Screen.TwipsPerPixelY < lngPatientInfoInvisibleSize)) Then
        blnShowLabel = False
    Else
        blnShowLabel = True
    End If
    
    If blnShowLabel = True Then blnShowLabel = Button_miDispPatientInfo
    
    intImageIndex = thisViewer.CurrentIndex
    '如果病人信息不随图像缩放，且图像小于指定大小，则不显示病人信息
    For i = 1 To thisViewer.MultiColumns
        For j = 1 To thisViewer.MultiRows
            Set img = thisViewer.Images(intImageIndex)
            Call subInitImageLabels(thisViewer.Index, 1, img, blnShowLabel)
            intImageIndex = intImageIndex + 1
            If intImageIndex > thisViewer.Images.Count Then
                Exit Sub
            End If
        Next j
    Next i
        
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subSortImages(thisForm As frmViewer, intViewerIndex As Integer, iSortType As Integer)
'------------------------------------------------
'功能： 对观片窗体中的图像进行排序，参加排序的Viewr的索引为intViewerIndex
'参数： thisForm--进行排序的窗体
'       intViewerIndex -- 进行排序的Viewer的索引
'       iSortType -- 排序方式，0--图像号；1--床位正序；2--床位逆序；3--采集时间；4--图像时间。
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim SortShowImageInfos As Collection
    Dim thisViewer As DicomViewer
    Dim intImagesCount As Integer
    Dim OneImageInfos As clsImageInfo
    Dim tmpListItem As ListItem
    Dim iOldIndex As Integer
    Dim SortImages As New DicomImages
    Dim i As Integer
    Dim j As Integer
    Dim k As String
    
    On Error GoTo err
    
    If ZLShowSeriesInfos(intViewerIndex).intSortType = iSortType Then Exit Sub
    
    '首先对ZLShowSeriesInfos中的图像进行排序
    ZLShowSeriesInfos(intViewerIndex).intSortType = iSortType
    
    Set thisViewer = thisForm.Viewer(intViewerIndex)
    intImagesCount = ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
    If intImagesCount = 0 Then Exit Sub
    
    '如果ZLShowSeriesInfos中没有排序的关键字信息，则从图像中读取这些信息
    If ZLShowSeriesInfos(intViewerIndex).ImageInfos(1).SliceLocation = "" Then
        Call ReadSortInfoFromImage(thisViewer)
    End If
    
    '创建一个SortShowImageInfos，用来保存ZLShowSeriesInfos.ImageInfos的副本
    Set SortShowImageInfos = ZLShowSeriesInfos(intViewerIndex).ImageInfos
    
    '清空ZLShowSeriesInfos.ImageInfos中的内容
    Set ZLShowSeriesInfos(intViewerIndex).ImageInfos = Nothing
    
    '用ListView进行排序
    '把排序关键字填写进入ListView
    thisForm.lvwSort.ListItems.Clear
    For i = 1 To intImagesCount
        If iSortType = 0 Then           '按照图像号排序
            Set tmpListItem = thisForm.lvwSort.ListItems.Add(, , SortShowImageInfos(i).ImageNo)
            tmpListItem.Text = String(6 - Len(tmpListItem.Text), "0") & tmpListItem.Text
        ElseIf iSortType = 1 Or iSortType = 2 Then  '按切片位置正序或者逆序
            Set tmpListItem = thisForm.lvwSort.ListItems.Add(, , SortShowImageInfos(i).SliceLocation)
            k = Val(tmpListItem.Text) * 100 + 100000 '保证正数和负数切片位置都能够统一进行排序
            k = Format(k, "#0")
            If Len(k) > 8 Then
                k = left(k, 8)
            End If
            tmpListItem.Text = String(8 - Len(k), "0") & k
        ElseIf iSortType = 3 Then       '按照采集时间
            Set tmpListItem = thisForm.lvwSort.ListItems.Add(, , SortShowImageInfos(i).AcquisitionTime)
        Else                            '按照图像时间
            Set tmpListItem = thisForm.lvwSort.ListItems.Add(, , SortShowImageInfos(i).ImageTime)
        End If
        tmpListItem.SubItems(1) = i
    Next i
    
    '对ListView的文本进行排序
    thisForm.lvwSort.SortKey = 0
    thisForm.lvwSort.Sorted = True
    If iSortType = 2 Then   '切片位置逆序
        thisForm.lvwSort.SortOrder = lvwDescending
    Else                    '其他方法用正序
        thisForm.lvwSort.SortOrder = lvwAscending
    End If
    
    '排序完成后，根据新旧索引，把图像信息从SortShowImageinfos复制回ZLShowSeriesInfos.ImageInfos
    For i = 1 To thisForm.lvwSort.ListItems.Count
        iOldIndex = Val(thisForm.lvwSort.ListItems(i).SubItems(1))
        ZLShowSeriesInfos(intViewerIndex).ImageInfos.Add SortShowImageInfos(iOldIndex)
    Next i
    
    '然后重新调整Viewer中图像的位置和Tag
    For i = 1 To thisViewer.Images.Count
        SortImages.Add thisViewer.Images(i)
    Next i
    thisViewer.Images.Clear
    For i = 1 To intImagesCount
        '只处理已经显示了的图像
        If ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).blnDisplayed = True Then
            For j = 1 To SortImages.Count
                If ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).InstanceUID = SortImages(j).InstanceUID Then
                    thisViewer.Images.Add SortImages(j)
                    thisViewer.Images(thisViewer.Images.Count).Tag = i
                    SortImages.Remove (j)
                    Exit For
                End If
            Next j
        End If
    Next i
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub ReadSortInfoFromImage(thisViewer As DicomViewer)
'------------------------------------------------
'功能： 从图像中读取排序信息进入ZLSeriesInfos结构
'参数： intSeriesIndex--图像所在序列的索引
'       intViewerIndex -- 进行排序的Viewer的索引
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim img As DicomImage
    Dim intViewerIndex As Integer
    Dim intImageCount As Integer
    Dim i As Integer
    Dim j  As Integer
    
    On Error GoTo err
    intViewerIndex = thisViewer.Index
    If intViewerIndex = 0 Then Exit Sub
    '循环序列中所有图像
    intImageCount = ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
    For i = 1 To intImageCount
        '如果图像已经显示，则从Viewer的图像中提取信息
        If ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).blnDisplayed = True Then
            For j = IIf(thisViewer.Images.Count >= i, i, thisViewer.Images.Count) To 1 Step -1
                If ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).InstanceUID = thisViewer.Images(j).InstanceUID Then
                    Set img = thisViewer.Images(j)
                    Exit For
                End If
            Next j
        Else    '如果图像不存在，下载图像
            Set img = funLoadAImage(intViewerIndex, i, 1)
        End If
        
        '把图像中的信息添加到ZLSeriesInfos结构中
        If Not img Is Nothing Then
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).AcquisitionTime = Format(GetImageAttribute(img.Attributes, ATTR_采集日期) & " " & GetImageAttribute(img.Attributes, ATTR_采集时间), "yyyy-MM-dd HH:MM:SS")
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).Columns = img.sizeX
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).FrameOfReferenceUID = GetImageAttribute(img.Attributes, ATTR_参考帧UID)
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).ImageOrientationPatient = GetImageAttribute(img.Attributes, ATTR_图像方向病人)
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).ImagePositionPatient = GetImageAttribute(img.Attributes, ATTR_图像位置病人)
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).ImageTime = Format(GetImageAttribute(img.Attributes, ATTR_图像日期) & " " & GetImageAttribute(img.Attributes, ATTR_图像时间), "yyyy-MM-dd HH:MM:SS")
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).PixelSpacing = GetImageAttribute(img.Attributes, ATTR_像素距离)
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).Rows = img.sizeY
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).SliceLocation = GetImageAttribute(img.Attributes, ATTR_切片位置)
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).SliceThickness = GetImageAttribute(img.Attributes, ATTR_层厚)
        End If
    Next i
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subResizeSeries(thisForm As frmViewer)
'------------------------------------------------
'功能： 改变窗口大小后，重新调整所有Viewer的位置
'参数： thisForm --- 观片窗体
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim intViewerIndex As Integer
    Dim blnExit As Boolean
    
    On Error GoTo err
    
    '根据序列布局，摆放分隔条
    Call subShowSpliter(thisForm)
    
    '循环所有的Viewer，重新摆放位置
    If Button_miLookOrBrowse = True Then       '观察模式，只摆放一个Viewer
        If thisForm.intSelectedSerial > 1 And thisForm.intSelectedSerial < thisForm.Viewer.Count Then
            '摆放这个Viewer并设置滚动条
            Call subPlaceAViewer(thisForm, thisForm.intSelectedSerial, 1, 1)
        End If
    Else
        intViewerIndex = 0
        For i = 1 To thisForm.intCountY
            For j = 1 To thisForm.intCountX
                intViewerIndex = intViewerIndex + 1
                If intViewerIndex >= thisForm.Viewer.Count Then
                    blnExit = True
                    Exit For
                End If
                '摆放这个Viewer
                Call subPlaceAViewer(thisForm, intViewerIndex, i, j)
            Next j
            If blnExit = True Then Exit For
        Next i
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub subDisplayReferLine(thisViewer As DicomViewer, thisForm As frmViewer, blnCurrentOnly As Boolean)
'------------------------------------------------
'功能： 根据菜单选项，显示三种类型的定位线
'参数： thisViewer ---当前选中的Viewer，将显示该Viewer向其他Viewer投影的定位线
'       thisForm --- 观片窗体
'       blnCurrentOnly --- 只刷新当前定位线
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim strLineTag As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim blnExit As Boolean
    Dim intCurrentImageIndex As Integer
    Dim imgSource As New DicomImage
    Dim imgDest As DicomImage
    Dim viewerDest As DicomViewer
    Dim intDestViewerIndex As Integer
    
    If thisViewer.Images.Count = 0 Then Exit Sub
    
    On Error GoTo err
    '设置定位线TAG的前导符
    If blnCurrentOnly = True Then
        strLineTag = "RLC"
    Else
        strLineTag = "RL"
    End If
    
    For Each viewerDest In thisForm.Viewer
        If viewerDest.Index <> 0 And viewerDest.Images.Count > 0 Then
            intCurrentImageIndex = viewerDest.CurrentIndex
            
            blnExit = False
            For i = 1 To viewerDest.MultiRows
                For j = 1 To viewerDest.MultiColumns
                    '删除thisViewer中可见图像中，旧的定位线
                    subDeleteAppointLabel viewerDest.Images(intCurrentImageIndex), strLineTag
                    '设置下一个图像
                    intCurrentImageIndex = intCurrentImageIndex + 1
                    If intCurrentImageIndex > viewerDest.Images.Count Then
                        blnExit = True
                        Exit For
                    End If
                Next j
                If blnExit = True Then Exit For
            Next i
        End If
        '刷新Viewer
        viewerDest.Refresh
    Next
    
    If Button_miAllReferLine = False And Button_miFLReferLine = False And Button_miCurrentReferLine = False Then Exit Sub
    
    '对于全部定位线和首尾定位线，需要对图像根据图像号来进行排序
    If Button_miAllReferLine Or Button_miFLReferLine Then
        '检查图像的排序方式，按照图像号来排序
        Call subSortImages(thisForm, thisViewer.Index, 0)
    End If
    
    For Each viewerDest In thisForm.Viewer
        intDestViewerIndex = viewerDest.Index
        If intDestViewerIndex <> 0 And intDestViewerIndex <> thisViewer.Index And viewerDest.Images.Count > 0 Then
            intCurrentImageIndex = viewerDest.CurrentIndex
            
            blnExit = False
            For i = 1 To viewerDest.MultiRows
                For j = 1 To viewerDest.MultiColumns
                    '在目标图 imgDest中画定位线
                    Set imgDest = viewerDest.Images(intCurrentImageIndex)
                    
                    '画定位线,不同方式的定位线分别处理
                    If Button_miAllReferLine = True Then    '显示所有定位线
                        For k = 1 To ZLShowSeriesInfos(thisViewer.Index).ImageInfos.Count
                            '填写虚拟图像的内容，然后计算定位线
                            Call subWriteRefLineImage(imgSource, k, thisViewer)
                            Call subDrawRefLine(imgSource, imgDest, True, "RLL", True)
                        Next k
                    End If
                    
                    If Button_miFLReferLine = True Then     '显示首尾定位线
                        '填写虚拟图像的内容，然后计算定位线
                        Call subWriteRefLineImage(imgSource, 1, thisViewer)
                        Call subDrawRefLine(imgSource, imgDest, False, "RLL", True)
                        If ZLShowSeriesInfos(thisViewer.Index).ImageInfos.Count > 1 Then
                            Call subWriteRefLineImage(imgSource, ZLShowSeriesInfos(thisViewer.Index).ImageInfos.Count, thisViewer)
                            Call subDrawRefLine(imgSource, imgDest, False, "RLL", True)
                        End If
                    End If
                    
                    If Button_miCurrentReferLine = True Then    '显示当前定位线
                        '如果当前图像是首尾图像，并且已经显示了首尾定位线，则不处理
                        If Not (Button_miFLReferLine And (thisForm.SelectedImage.Tag = 1 Or thisForm.SelectedImage.Tag = ZLShowSeriesInfos(thisViewer.Index).ImageInfos.Count)) Then
                            Call subWriteRefLineImage(imgSource, thisForm.SelectedImage.Tag, thisViewer)
                            Call subDrawRefLine(imgSource, imgDest, False, "RLC", True)
                        End If
                    End If
                    
                    '设置下一个图像
                    intCurrentImageIndex = intCurrentImageIndex + 1
                    If intCurrentImageIndex > viewerDest.Images.Count Then
                        blnExit = True
                        Exit For
                    End If
                Next j
                If blnExit = True Then Exit For
            Next i
        End If
        '刷新Viewer
        viewerDest.Refresh
    Next
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subWriteRefLineImage(img As DicomImage, intImageIndex As Integer, thisViewer As DicomViewer)
'------------------------------------------------
'功能： 把定位线信息填写到img图像中
'参数： img ---需要填写定位线信息的图像
'       intViewerIndex --- 图像所在Viewer的索引
'       intImageIndex --- 图像所在的图像索引
'       thisViewer --- 图像所在的Viewer
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim oneImageInfo As clsImageInfo
    Dim tmpValues() As String
    Dim intViewerIndex As Integer
    
    On Error GoTo err
    intViewerIndex = thisViewer.Index
    
    '先判断ZLShowSeriesInfos是否有内容，如果没有，则从图像中读取内容
    If ZLShowSeriesInfos(intViewerIndex).ImageInfos(intImageIndex).SliceLocation = "" Then
        Call ReadSortInfoFromImage(thisViewer)
    End If
    '开始填写信息
     
    Set oneImageInfo = ZLShowSeriesInfos(intViewerIndex).ImageInfos(intImageIndex)
    
    img.Attributes.Add &H20, &H13, oneImageInfo.ImageNo
    img.Attributes.Add &H28, &H10, oneImageInfo.Rows
    img.Attributes.Add &H28, &H11, oneImageInfo.Columns
    img.Attributes.Add &H20, &H1041, oneImageInfo.SliceLocation
    img.Attributes.Add &H20, &H52, oneImageInfo.FrameOfReferenceUID
    tmpValues = Split(oneImageInfo.PixelSpacing, "\")
    img.Attributes.Add &H28, &H30, tmpValues
    tmpValues = Split(oneImageInfo.ImagePositionPatient, "\")
    img.Attributes.Add &H20, &H32, tmpValues
    tmpValues = Split(oneImageInfo.ImageOrientationPatient, "\")
    img.Attributes.Add &H20, &H37, tmpValues
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function funGetNewSeriesInfo() As clsSeriesInfo
'------------------------------------------------
'功能： 创建一个新的序列信息，并且赋予初始值
'参数： 无
'返回： 无
'时间：2009-7
'------------------------------------------------
    Dim oneSeriesInfo As New clsSeriesInfo
    oneSeriesInfo.blnImageSyn = True
    oneSeriesInfo.FilterLength = 0
    oneSeriesInfo.FlipState = doFlipNormal
    oneSeriesInfo.intSortType = 0   '记录当前序列的排序方式：0--图像号；1--床位正序；2--床位逆序；3--采集时间；4--图像时间，仅在ZLShowSeriesInfos中使用。
    oneSeriesInfo.lngSource = 0     '图像来源，0-从PACS图像服务器下载；1-直接打开文件；2---混合；3-重新生成的序列，类似矢冠状位重建、图像拼接、伪彩生成的图像
    oneSeriesInfo.lngWinLevel = 0   '0 表示没有预设的窗宽窗位
    oneSeriesInfo.lngWinWidth = 0   '0 表示没有预设的窗宽窗位
    oneSeriesInfo.RotateState = doRotateNormal
    oneSeriesInfo.ScrollX = 0
    oneSeriesInfo.ScrollY = 0
    oneSeriesInfo.StretchToFit = True
    oneSeriesInfo.UnsharpEnhancement = 0    '边缘增强强度。仅在ZLShowSeriesInfos中使用。
    oneSeriesInfo.UnsharpLength = 0         '边缘增强幅度。仅在ZLShowSeriesInfos中使用。
    oneSeriesInfo.Zoom = 1
    oneSeriesInfo.MultiFrame = 0            '默认是单帧图像
    oneSeriesInfo.Selected = False          '默认没有被选择
    oneSeriesInfo.intCol = 0
    oneSeriesInfo.intRow = 0
    Set funGetNewSeriesInfo = oneSeriesInfo
End Function

Public Function funGetNewImageInfo() As clsImageInfo
'------------------------------------------------
'功能： 创建一个新的图像信息，并且赋予初始值
'参数： 无
'返回： 无
'时间：2009-7
'------------------------------------------------
    Dim oneImageInfo As New clsImageInfo
    oneImageInfo.blnDisplayed = False
    oneImageInfo.blnSelected = False
    oneImageInfo.int3DLabelIndex = 0        '表示没有三维鼠标定位线
    oneImageInfo.blnPrinted = False         '表示没有被打印
    Set funGetNewImageInfo = oneImageInfo
End Function

Public Function funCopySeriesInfo(sourceSeriesInfo As clsSeriesInfo, destSeriesInfo As clsSeriesInfo) As Boolean
'------------------------------------------------
'功能： 复制序列信息,只复制序列的信息，不复制的内容包括：包含的图像信息，序列的位置intCol和intRow
'参数： sourceSeriesInfo --- 源序列
'       destSeriesInfo ---- 目标序列
'返回： 是否成功
'时间：2009-7
'------------------------------------------------
    destSeriesInfo.blnImageSyn = sourceSeriesInfo.blnImageSyn
    destSeriesInfo.FilterLength = sourceSeriesInfo.FilterLength
    destSeriesInfo.FlipState = sourceSeriesInfo.FlipState
    destSeriesInfo.intSortType = sourceSeriesInfo.intSortType
    destSeriesInfo.lngSource = sourceSeriesInfo.lngSource
    destSeriesInfo.lngWinLevel = sourceSeriesInfo.lngWinLevel
    destSeriesInfo.lngWinWidth = sourceSeriesInfo.lngWinWidth
    destSeriesInfo.RotateState = sourceSeriesInfo.RotateState
    destSeriesInfo.ScrollX = sourceSeriesInfo.ScrollX
    destSeriesInfo.ScrollY = sourceSeriesInfo.ScrollY
    destSeriesInfo.SeriesNo = sourceSeriesInfo.SeriesNo
    destSeriesInfo.SeriesUID = sourceSeriesInfo.SeriesUID
    destSeriesInfo.StretchToFit = sourceSeriesInfo.StretchToFit
    destSeriesInfo.strFTPDir = sourceSeriesInfo.strFTPDir
    destSeriesInfo.strFTPPasw = sourceSeriesInfo.strFTPPasw
    destSeriesInfo.strFTPUser = sourceSeriesInfo.strFTPUser
    destSeriesInfo.strHostIP = sourceSeriesInfo.strHostIP
    destSeriesInfo.strModality = sourceSeriesInfo.strModality
    destSeriesInfo.strSaveDir = sourceSeriesInfo.strSaveDir
    destSeriesInfo.strShareDir = sourceSeriesInfo.strShareDir
    destSeriesInfo.strShareDirPasw = sourceSeriesInfo.strShareDirPasw
    destSeriesInfo.strShareDirUser = sourceSeriesInfo.strShareDirUser
    destSeriesInfo.StudyUID = sourceSeriesInfo.StudyUID
    destSeriesInfo.UnsharpEnhancement = sourceSeriesInfo.UnsharpEnhancement
    destSeriesInfo.UnsharpLength = sourceSeriesInfo.UnsharpLength
    destSeriesInfo.Zoom = sourceSeriesInfo.Zoom
    destSeriesInfo.MultiFrame = sourceSeriesInfo.MultiFrame
    destSeriesInfo.Selected = sourceSeriesInfo.Selected
    destSeriesInfo.strCName = sourceSeriesInfo.strCName
    destSeriesInfo.strEName = sourceSeriesInfo.strEName
    destSeriesInfo.strSex = sourceSeriesInfo.strSex
    destSeriesInfo.strAge = sourceSeriesInfo.strAge
    destSeriesInfo.strStudyID = sourceSeriesInfo.strStudyID
    destSeriesInfo.strOrderID = sourceSeriesInfo.strOrderID
    funCopySeriesInfo = True
End Function

Public Function funCopyImageInfo(sourceImageInfo As clsImageInfo, destImageInfo As clsImageInfo) As Boolean
'------------------------------------------------
'功能： 复制图像信息
'参数： sourceImageInfo --- 源图像
'       destImageInfo ---- 目标图像
'返回： 是否成功
'时间：2009-7
'------------------------------------------------
    destImageInfo.AcquisitionTime = sourceImageInfo.AcquisitionTime
    destImageInfo.blnDisplayed = sourceImageInfo.blnDisplayed
    destImageInfo.blnSelected = sourceImageInfo.blnSelected
    destImageInfo.Columns = sourceImageInfo.Columns
    destImageInfo.FrameOfReferenceUID = sourceImageInfo.FrameOfReferenceUID
    destImageInfo.ImageName = sourceImageInfo.ImageName
    destImageInfo.ImageNo = sourceImageInfo.ImageNo
    destImageInfo.ImageOrientationPatient = sourceImageInfo.ImageOrientationPatient
    destImageInfo.ImagePositionPatient = sourceImageInfo.ImagePositionPatient
    destImageInfo.ImageTime = sourceImageInfo.ImageTime
    destImageInfo.InstanceUID = sourceImageInfo.InstanceUID
    destImageInfo.PixelSpacing = sourceImageInfo.PixelSpacing
    destImageInfo.Rows = sourceImageInfo.Rows
    destImageInfo.SliceLocation = sourceImageInfo.SliceLocation
    destImageInfo.SliceThickness = sourceImageInfo.SliceThickness
    destImageInfo.int3DLabelIndex = sourceImageInfo.int3DLabelIndex
    destImageInfo.blnPrinted = sourceImageInfo.blnPrinted
    funCopyImageInfo = True
End Function

Public Sub subOpenCurrentImage(thisForm As frmViewer, img As DicomImage)
'------------------------------------------------
'功能： 把当前的图像打开，处理ZLSeriesInfos等结构
'参数： thisForm--打开图像的窗体
'       img --- 需要打开的图像
'返回： 无
'时间：2009-7
'------------------------------------------------
    Dim i As Integer
    Dim blnNewSeries As Boolean
    Dim iCurrentSeries As Integer   '当前序列的索引
    Dim iSiblingSeries As Integer   '兄弟序列的索引
    Dim oneSeriesInfo As clsSeriesInfo
    Dim iCurrentImage As Integer
    Dim iPrivImage As Integer       '当前图像的前一个图像
    Dim oneImageInfo As clsImageInfo
    
    On Error GoTo err
    '检查该该图像的序列是否已经存在
    blnNewSeries = True
    For i = 1 To ZLSeriesInfos.Count
        If ZLSeriesInfos(i).SeriesUID = img.SeriesUID Then
            blnNewSeries = False
            iCurrentSeries = i
        End If
        If ZLSeriesInfos(i).StudyUID = img.StudyUID Then
            iSiblingSeries = i
        End If
    Next i
    If iSiblingSeries = 0 Then Exit Sub     '如果找不到兄弟序列则不打开图像
    
    '如果不存在，则增加该序列，如果存在则直接在该序列中查找图像是否存在
    '序列不排序，图像需按照图像号排序
    If blnNewSeries = True Then
        Set oneSeriesInfo = funGetNewSeriesInfo
        oneSeriesInfo.StudyUID = img.StudyUID
        oneSeriesInfo.SeriesUID = img.SeriesUID
        oneSeriesInfo.SeriesNo = 1
        oneSeriesInfo.strModality = ZLSeriesInfos(iSiblingSeries).strModality
        oneSeriesInfo.MultiFrame = 0
        oneSeriesInfo.strHostIP = ZLSeriesInfos(iSiblingSeries).strHostIP
        oneSeriesInfo.strFTPDir = ZLSeriesInfos(iSiblingSeries).strFTPDir
        oneSeriesInfo.strFTPPasw = ZLSeriesInfos(iSiblingSeries).strFTPPasw
        oneSeriesInfo.strFTPUser = ZLSeriesInfos(iSiblingSeries).strFTPUser
        oneSeriesInfo.strShareDir = ZLSeriesInfos(iSiblingSeries).strShareDir
        oneSeriesInfo.strShareDirUser = ZLSeriesInfos(iSiblingSeries).strShareDirUser
        oneSeriesInfo.strShareDirPasw = ZLSeriesInfos(iSiblingSeries).strShareDirPasw
        oneSeriesInfo.strSaveDir = ZLSeriesInfos(iSiblingSeries).strSaveDir
        ZLSeriesInfos.Add oneSeriesInfo, CStr(ZLSeriesInfos.Count + 1)
        iCurrentSeries = ZLSeriesInfos.Count
    End If
    
    '查找图像是否存在
    iCurrentImage = 0
    iPrivImage = 0
    For i = 1 To ZLSeriesInfos(iCurrentSeries).ImageInfos.Count
        If ZLSeriesInfos(iCurrentSeries).ImageInfos(i).InstanceUID = img.InstanceUID Then
            iCurrentImage = i
            Exit For
        End If
    Next i
    If iCurrentImage <> 0 Then Exit Sub         '图像已经存在，则不打开
    '打开图像
    Set oneImageInfo = funGetNewImageInfo
    oneImageInfo.InstanceUID = img.InstanceUID
    oneImageInfo.ImageNo = 1
    oneImageInfo.ImageName = img.InstanceUID
    oneImageInfo.AcquisitionTime = Format(GetImageAttribute(img.Attributes, ATTR_采集日期) & " " & GetImageAttribute(img.Attributes, ATTR_采集时间), "yyyy-MM-dd HH:MM:SS")
    oneImageInfo.ImageTime = Format(GetImageAttribute(img.Attributes, ATTR_图像日期) & " " & GetImageAttribute(img.Attributes, ATTR_图像时间), "yyyy-MM-dd HH:MM:SS")
    oneImageInfo.SliceThickness = GetImageAttribute(img.Attributes, ATTR_层厚)
    oneImageInfo.SliceLocation = GetImageAttribute(img.Attributes, ATTR_切片位置)
    oneImageInfo.ImageOrientationPatient = GetImageAttribute(img.Attributes, ATTR_图像方向病人)
    oneImageInfo.ImagePositionPatient = GetImageAttribute(img.Attributes, ATTR_图像位置病人)
    oneImageInfo.Rows = img.sizeY
    oneImageInfo.Columns = img.sizeX
    oneImageInfo.PixelSpacing = GetImageAttribute(img.Attributes, ATTR_像素距离)
    '在第一个图像前面追加
    ZLSeriesInfos(iCurrentSeries).ImageInfos.Add oneImageInfo
    
    '显示这个序列,用新序列的图像代替viewer(index)中的图像
    If thisForm.Viewer.Count >= 2 Then
        Call thisForm.funcSwapSeries(thisForm.Viewer.Count - 1, iCurrentSeries)
    End If
        
    '重新显示缩略图
    subShowMiniImages thisForm
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function funShowTempImages(thisForm As frmViewer, imgs As DicomImages, iViewerIndex As Integer) As Integer
'------------------------------------------------
'功能： 临时打开并显示imgs里面的图像，只把图像加载到ZLShowSeriesInfos结构中
'参数： thisForm -- 显示图像的窗体
'       imgs -- 需要临时显示的图像集合
'       iViewerIndex -- 显示图像的Viewer的Index,如果为0， 则表示需要自动查找可用的Viewer
'返回： 添加临时图像的Viewer的Index
'时间：2009-7
'------------------------------------------------
    Dim i As Integer
    Dim oneImageInfo As clsImageInfo
    Dim img As DicomImage
    
    On Error GoTo err
    
    If imgs.Count <= 0 Then Exit Function
    
    '如果iViewerIndex为0 ，则查找或者创建一个可用的Viewer
    If iViewerIndex = 0 Or iViewerIndex >= thisForm.Viewer.Count Then
        iViewerIndex = funcGetAUsableViewer(thisForm)
    End If
    
    '根据第一个图象的信息，修改ZLShowSeriesInfos结构
    Set ZLShowSeriesInfos(iViewerIndex).ImageInfos = Nothing
    ZLShowSeriesInfos(iViewerIndex).lngSource = 1
    ZLShowSeriesInfos(iViewerIndex).SeriesNo = GetImageAttribute(imgs(1).Attributes, ATTR_序列号)
    ZLShowSeriesInfos(iViewerIndex).SeriesUID = imgs(1).SeriesUID
    ZLShowSeriesInfos(iViewerIndex).strModality = GetImageAttribute(imgs(1).Attributes, ATTR_影像类别)
    ZLShowSeriesInfos(iViewerIndex).strSaveDir = ""
    ZLShowSeriesInfos(iViewerIndex).StudyUID = imgs(1).StudyUID
    ZLShowSeriesInfos(iViewerIndex).MultiFrame = 0
    ZLShowSeriesInfos(iViewerIndex).Selected = False
    
    '循环把每一个图象添加到图像结构中
    For i = imgs.Count To 1 Step -1
        Set oneImageInfo = funGetNewImageInfo
        oneImageInfo.AcquisitionTime = Format(GetImageAttribute(imgs(i).Attributes, ATTR_采集日期) & " " & GetImageAttribute(imgs(i).Attributes, ATTR_采集时间), "yyyy-MM-dd HH:MM:SS")
        oneImageInfo.Columns = imgs(i).sizeX
        oneImageInfo.FrameOfReferenceUID = GetImageAttribute(imgs(i).Attributes, ATTR_参考帧UID)
        oneImageInfo.ImageName = imgs(i).InstanceUID
        oneImageInfo.ImageNo = GetImageAttribute(imgs(i).Attributes, ATTR_图像号)
        oneImageInfo.ImageOrientationPatient = GetImageAttribute(imgs(i).Attributes, ATTR_图像方向病人)
        oneImageInfo.ImagePositionPatient = GetImageAttribute(imgs(i).Attributes, ATTR_图像位置病人)
        oneImageInfo.ImageTime = Format(GetImageAttribute(imgs(i).Attributes, ATTR_图像日期) & " " & GetImageAttribute(imgs(i).Attributes, ATTR_图像时间), "yyyy-MM-dd HH:MM:SS")
        oneImageInfo.InstanceUID = imgs(i).InstanceUID
        oneImageInfo.PixelSpacing = GetImageAttribute(imgs(i).Attributes, ATTR_像素距离)
        oneImageInfo.Rows = imgs(i).sizeY
        oneImageInfo.SliceLocation = GetImageAttribute(imgs(i).Attributes, ATTR_切片位置)
        oneImageInfo.SliceThickness = GetImageAttribute(imgs(i).Attributes, ATTR_层厚)
        oneImageInfo.blnDisplayed = True
        '在第一个图的前面追加
        ZLShowSeriesInfos(iViewerIndex).ImageInfos.Add oneImageInfo
    Next i
    
    '打开并显示图像
    thisForm.Viewer(iViewerIndex).Images.Clear
    For i = 1 To imgs.Count
        thisForm.Viewer(iViewerIndex).Images.Add imgs(i)
        
        Set img = thisForm.Viewer(iViewerIndex).Images(thisForm.Viewer(iViewerIndex).Images.Count)
        img.Tag = i
        If img.Labels.Count = 0 Then
            Call subInitAImage(img, iViewerIndex, thisForm.Viewer(iViewerIndex))
        End If
    Next i
    
    '主动刷新Viewer，否则Viewer在鼠标移动MPR重建线的时候不会自己刷新，显得很滞后
    thisForm.Viewer(iViewerIndex).Refresh
    
    '判断滚动条是否需要显示，如果需要则显示滚动条，设置滚动条的最大值，最小值，LarghChange等
    subDisplayScrollBar iViewerIndex, thisForm, False
    
    '如果是选择所有序列的状态，将目前序列设置为当前序列，供subDispframe使用
    If thisForm.isSelectAllSerial Then thisForm.intSelectedSerial = iViewerIndex
    
    '图像显示完后，添加Viewer中的标注：图象框、右下角的选择标记等
    Call subDispframe(thisForm, thisForm.Viewer(iViewerIndex))
    
    '自动根据图像大小，判断是否显示病人四角信息,显示或者隐藏图像中的病人信息
    Call subDisplayPatientInfo(thisForm.Viewer(iViewerIndex))
    
    funShowTempImages = iViewerIndex
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function funcGetAUsableViewer(thisForm As frmViewer) As Integer
'------------------------------------------------
'功能： 从当前窗口中查找一个可以添加图象的Viewer
'       如果当前窗口有空的地方，优先往空白处添加Viewer，并且创建ZLShowSeriesInfos结构。
'       如果没有空白处，则使用当前打开的最后一个Viewer
'参数： thisForm -- 显示图像的窗体
'返回： 找到的Viewer的Index，0表示没有找到
'时间：2009-7
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim intViewerIndex As Integer
    
    On Error GoTo err
    '根据图像的布局来查找空白的Viewer
    '如果在现存的Viewer中有空白Viewer，直接使用
    '否则检查Viewer的数量是否少于界面可以显示的Viewer的总合，如果小于，则创建一个新的Viewer
    '再否则就直接使用当前的最后一个Viewer
    
    '先检查当前Viewer中是否有现存的Viewer
    For i = 1 To thisForm.Viewer.Count - 1
        If thisForm.Viewer(i).Images.Count = 0 Then
            funcGetAUsableViewer = i
            Exit Function
        End If
    Next i
    
    '再检查Viewer的数量是否少于界面可显示的Viewer的总合
    If thisForm.Viewer.Count - 1 < thisForm.intCountX * thisForm.intCountY Then
        '有空白的Viewer，查找空白Viewer的位置，并创建一个Viewer
        For i = 1 To thisForm.intCountY
            For j = 1 To thisForm.intCountX
                intViewerIndex = 0
                For k = 1 To ZLShowSeriesInfos.Count
                    If ZLShowSeriesInfos(k).intRow = i And ZLShowSeriesInfos(k).intCol = j Then
                        intViewerIndex = k
                    End If
                Next k
                If intViewerIndex = 0 Then Exit For
            Next j
            If intViewerIndex = 0 Then Exit For
        Next i
        If intViewerIndex = 0 Then
            '创建一个Viewer
            intViewerIndex = funcCeateAViewer(1, thisForm)
            '摆放一个Viewer
            Call subPlaceAViewer(thisForm, intViewerIndex, i, j)
            funcGetAUsableViewer = intViewerIndex
        End If
    Else
        '直接使用最后一个Viewer
        funcGetAUsableViewer = thisForm.Viewer.Count - 1
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function funMPR(thisForm As frmViewer, Optional blnSilent As Boolean = False) As Boolean
'------------------------------------------------
'功能： 对当前窗体中被选中的序列做矢冠状位重建，或者取消矢冠状位重建，不直接调用
'       thisForm.blnInMPR 说明窗体中是否有图像正在进行重建的过程中
'参数： thisForm -- 显示图像的窗体
'       blnSilent -- 静默结束MRP，不提示
'返回： True--成功，False---取消退出
'时间：2009-7
'------------------------------------------------
    Dim thisViewer As DicomViewer
    Dim intMPRViewerIndex1 As Integer
    Dim intMPRViewerIndex2 As Integer
    Dim intMPRViewerIndex3 As Integer
    Dim resImage1 As New DicomImage
    Dim resImage2 As New DicomImage
    Dim i As Integer
    Dim dGlabal As New DicomGlobal
    Dim lngResult As Long
    Dim blnSortForward As Boolean       'True-按照床位正向排序；False-按照床位逆向排序
    
    On Error GoTo err
    
    '进行矢冠状位重建，但是却没有选中任何图像，则退出
    If thisForm.blnInMPR = False And thisForm.SelectedImage Is Nothing Then
        MsgBox "请先选择图像，然后再进行MPR。", vbOKOnly, "温馨提示"
        Exit Function
    End If
    
    Set thisViewer = thisForm.Viewer(thisForm.intSelectedSerial)
    '先处理取消重建，再处理重建
    If thisForm.blnInMPR = True Then    '取消重建
        
        If blnSilent = True Then       '静默，不提示是否结束MPR
            lngResult = vbNo
        Else
            '提示是否保存重建结果图
            lngResult = MsgBox("是否保存MPR重建的结果图？", vbYesNoCancel, gstrSysName)
            If lngResult = vbCancel Then
                funMPR = False
                Exit Function
            End If
        
        End If
        
        '先清除MPR重建的标记，后面有些操作依赖于这个标记
        thisForm.blnInMPR = False
        
        If lngResult = vbYes Then      '保存重建的结果图
            Set resImage1 = thisForm.Viewer(ZLMPRCube(2).intViewerIndex).Images(1)
            Set resImage2 = thisForm.Viewer(ZLMPRCube(3).intViewerIndex).Images(1)
            resImage1.SeriesUID = ZLMPRSeriesUID
            resImage2.SeriesUID = ZLMPRSeriesUID
            '保存结果图
            Call subSaveImage(resImage1, thisViewer.Images(1).SeriesUID)
            Call subSaveImage(resImage2, thisViewer.Images(1).SeriesUID)
            '把图像追加到观片站中
            Call subOpenCurrentImage(thisForm, resImage1)
            Call subOpenCurrentImage(thisForm, resImage2)
        End If
        
        '恢复原来被替换的Viewer中的图像内容，删除为了MPR而新添加的Viewer
        '处理MPR序列，被替换过了，才需要替换回来
        If ZLMPRCube(1).blnIsMPR = False Then
            Call subMPRReFillImagesToViewer(1, thisForm)
        Else
            '去除图像上面的矢冠状位重建标记
            For i = G_INT_SYS_LABEL_MPRV To G_INT_SYS_LABEL_MPR_POINT_O
                thisForm.Viewer(ZLMPRCube(1).intViewerIndex).Images(1).Labels(i).Visible = False
            Next i
            subMPRLinenPhase thisForm.Viewer(ZLMPRCube(1).intViewerIndex), thisForm.Viewer(ZLMPRCube(1).intViewerIndex).Images(1)
        End If
        
        '处理第二个结果图
        Call subMPRReFillImagesToViewer(3, thisForm)
        
        '处理第一个结果图
        Call subMPRReFillImagesToViewer(2, thisForm)
        
        '如果布局有改变，则调整窗体的序列布局
        If thisForm.intOldCountX <> thisForm.intCountX Or thisForm.intOldCountY <> thisForm.intCountY Then
            '恢复原来的图像布局
            thisForm.intCountX = thisForm.intOldCountX
            thisForm.intCountY = thisForm.intOldCountY
            Call subChangeSeriesLayout(thisForm)
        End If
        
        '清空缓存的三维数组
        ReDim aPixels(0)
    Else        '开始重建
        ZLMPRSeriesUID = dGlabal.NewUID
        '先把序列中的所有图像都加载到Viewer中
        Call funAddAllImages(thisViewer)
        
        '判断是否满足矢冠状位重建的条件
        If LeagelToACRebuild(thisViewer.Images) = 1 Then
            thisForm.blnInMPR = False   '设置重建状态为退出重建
            Exit Function
        End If
            
        '如果出于浏览观察模式，则退出该模式
        If Button_miLookOrBrowse = True Then
            Call subLookOrBrowsSwitch(thisForm)
        End If
        
        '记录原来的序列布局,对于小于2*2的序列布局，矢冠状位重建时需要将布局修改成2*2，退出重建后，要恢复原来的布局
        '对于大于2*2的序列布局，进行矢冠状位重建的时候，不需要调整序列的布局
        thisForm.intOldCountX = thisForm.intCountX
        thisForm.intOldCountY = thisForm.intCountY
        '如果序列布局小于2*2，则调整为2*2
        If thisForm.intCountX < 2 Then thisForm.intCountX = 2
        If thisForm.intCountY < 2 Then thisForm.intCountY = 2
        '如果布局有改变，则调整窗体的序列布局
        If thisForm.intOldCountX <> thisForm.intCountX Or thisForm.intOldCountY <> thisForm.intCountY Then
            Call subChangeSeriesLayout(thisForm)
        End If
        
        '摆放左上角的Viewer
        '如果左上角有图像，则保存这组图像，如果没有图像，则在这里创建Viewer
        intMPRViewerIndex1 = funcReplaceViewer(1, thisForm, thisViewer)
        '给把thisViewer中的图像添加到MPRViewer1中
        If thisViewer.Index <> intMPRViewerIndex1 Then
            Call funShowTempImages(thisForm, thisViewer.Images, intMPRViewerIndex1)
            Set thisViewer = thisForm.Viewer(intMPRViewerIndex1)
        End If
        
        
        '开始重建
        '对图像做重建的初始化
        '对Viewer中的全部图像，初始化矢冠状位重建的控制线和控制点
        Call subInitMPRLine(thisViewer)
        '矢冠状位重建初始化，填写层厚、总高度和像素数组
        Call funcPlaneRestructInit(thisViewer, thisForm)
        
        '创建第一个重建结果图
        intMPRViewerIndex2 = funcReplaceViewer(2, thisForm, Nothing)
        '根据传入的控制线，对MPRViewer内的图像进行重建，并将返回重建的结果图显示到ShowViewerIndex指定的Viewer中
        If funGetMPRImageAndShow(thisViewer.Images(1).Labels(G_INT_SYS_LABEL_MPRV), thisForm, thisViewer, _
                                    1, intMPRViewerIndex2, ToltalHeight, 1, True, True) = False Then
            '重建出错，退出MPR重建
            thisForm.blnInMPR = True
            Call funMPR(thisForm, True)
            funMPR = False
            Exit Function
        End If
                
        '创建第二个重建结果图
        intMPRViewerIndex3 = funcReplaceViewer(3, thisForm, Nothing)
        '根据传入的控制线，对MPRViewer内的图像进行重建，并将返回重建的结果图显示到ShowViewerIndex指定的Viewer中
        If funGetMPRImageAndShow(thisViewer.Images(1).Labels(G_INT_SYS_LABEL_MPRH), thisForm, thisViewer, _
                                    1, intMPRViewerIndex3, ToltalHeight, 2, True, True) = False Then
            '重建出错，退出MPR重建
            thisForm.blnInMPR = True
            Call funMPR(thisForm, True)
            funMPR = False
            Exit Function
        End If
                
        thisForm.intSelectedSerial = thisViewer.Index
         '设置MPR标记
         thisForm.blnInMPR = True
    End If
    
    funMPR = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function funMPRslope(imageViewer As DicomViewer, axialViewer As DicomViewer, _
    CoronalViewer As DicomViewer, SagittalViewer As DicomViewer, parForm As frmViewer) As Boolean
'------------------------------------------------
'功能： 对当前窗体中被选中的序列做矢冠状位斜面重建
'参数： imageViewer -- 图像所在的Viewer，在主窗体中
'       axialViewer -- 轴位图所在的Viewer，在斜面重建窗体中
'       CoronalViewer -- 重建结果图冠状位所在的Viewer，在斜面重建窗体中
'       SagittalViewer -- 重建结果图矢状位所在的Viewer，在斜面重建窗体中
'       parForm -- 图像所在的Form
'返回： True--成功，False---取消退出
'------------------------------------------------
    
    On Error GoTo err
        
    '开始重建
    '对图像做重建的初始化
    
    funMPRslope = False
    
    '矢冠状位重建初始化，填写层厚、总高度和像素数组
    Call funcPlaneRestructInit(imageViewer, parForm)
        
    '根据传入的控制线，对MPRViewer内的图像进行重建，并将返回重建的结果图显示到ShowViewerIndex指定的Viewer中
    If funGetCandSImageAndShow(axialViewer.Images(1).Labels(G_INT_SYS_LABEL_MPRH), imageViewer, _
                                    axialViewer, CoronalViewer, ToltalHeight, 1, True, True) = False Then
        '重建出错，退出MPR重建
        Exit Function
    End If
    
    '根据传入的控制线，对MPRViewer内的图像进行重建，并将返回重建的结果图显示到ShowViewerIndex指定的Viewer中
    If funGetCandSImageAndShow(axialViewer.Images(1).Labels(G_INT_SYS_LABEL_MPRV), imageViewer, _
                                    axialViewer, SagittalViewer, ToltalHeight, 2, True, True) = False Then
        '重建出错，退出MPR重建
        Exit Function
    End If
        
    funMPRslope = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function funAddAllImages(thisViewer As DicomViewer) As Boolean
'------------------------------------------------
'功能： 把整个序列的图像都加载到thisViewer中
'参数： thisViewer -- 需要加载的Viewer
'返回： 无
'时间：2009-7
'------------------------------------------------
    Dim i As Integer
    Dim intViewerIndex As Integer
    
    If thisViewer Is Nothing Then Exit Function
    
    On Error GoTo err
    intViewerIndex = thisViewer.Index
    For i = 1 To ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
        If ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).blnDisplayed = False Then
            Call funcAddAImageA(thisViewer, i)
        End If
    Next i
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub subChangeSeriesLayout(thisForm As frmViewer)
'------------------------------------------------
'功能：根据窗体参数intCountX和intCountY来切换序列布局
'参数： thisForm --- 观片窗口
'返回：无 ，直接切换序列布局
'时间：2009-7
'------------------------------------------------
    '切换序列布局的时候，只把已经加载进入Viewer中的图像显示出来，不显示其他图像
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim blnLoadOver As Boolean
    Dim iCurrentViewerIndex As Integer
    Dim intSeriesIndex As Integer
    Dim vTemp As DicomViewer
    Dim blnFound As Boolean
    
    On Error GoTo err
    '先对比原有序列布局跟当前序列布局中Viewer的数量，如果Viwer增加了，则装载新的Viewer，
    '新装载的Viewer中，依次按照缩略图的顺序，显示所有剩余的序列
    '如果Viewer减少了，则卸载多余的Viewer
    
    '根据序列布局，摆放分隔条
    Call subShowSpliter(thisForm)
    
    iCurrentViewerIndex = 0
    For i = 1 To thisForm.intCountY
        For j = 1 To thisForm.intCountX
            iCurrentViewerIndex = iCurrentViewerIndex + 1
            
            If iCurrentViewerIndex >= thisForm.Viewer.Count Then
                '创建一个Viewer
                '查询新创建的Viewer准备装载哪个序列
                intSeriesIndex = 0
                For k = 1 To ZLSeriesInfos.Count
                    For Each vTemp In thisForm.Viewer
                        If vTemp.Tag = k Then
                            blnFound = True
                            Exit For
                        Else
                            blnFound = False
                        End If
                    Next
                    If blnFound = False Then
                        intSeriesIndex = k
                        Exit For
                    End If
                Next k
                If intSeriesIndex = 0 Then
                    '说明所有序列都已经装在进来了，可以退出装载序列的循环了
                    blnLoadOver = True
                    Exit For
                End If
                iCurrentViewerIndex = funcCeateAViewer(intSeriesIndex, thisForm)
            End If
            '摆放这个Viewer,Viewer中有图像，才需要摆放，否则是一个空的Viewer，不需要摆放
            If thisForm.Viewer(iCurrentViewerIndex).Images.Count <> 0 Then
                Call subPlaceAViewer(thisForm, iCurrentViewerIndex, i, j)
            End If
        Next j
        If blnLoadOver = True Then
            Exit For
        End If
    Next i

    '卸载多余的Viewer
    If iCurrentViewerIndex < thisForm.Viewer.Count Then
        While thisForm.Viewer.Count > 1 And thisForm.Viewer.Count - 1 > iCurrentViewerIndex
            Call subUnloadLastViewer(thisForm)
        Wend
        If thisForm.intSelectedSerial >= thisForm.Viewer.Count Then thisForm.intSelectedSerial = 0
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function funcReplaceViewer(intType As Integer, thisForm As frmViewer, thisViewer As DicomViewer) As Integer
'------------------------------------------------
'功能： 在指定位置创建Viewer
'       如果指定位置有图像，则保存这组图像，如果没有图像，则在这里创建Viewer
'       替换Viewer有两种情况：
'           1、用MPR序列替换空Viewer或者有图像的Viewer。
'               1) 如果是空Viewer，直接创建一个Viewer,从ZLSeriesInfos中读取图像。
'               2) 如果是有图像的Viewer,且Viewer中的图像不是MPR序列，则把原有序列保存到ZLMPRImages(1)中，在此Viewer中显示MPR序列，从ZLSeriesInfos中读取图像。
'               3) 如果是有图像的Viewer，且VIewer中的图像是MPR序列，则做标记，且装载序列中的全部图像。
'           2、用MPR结果序列替换空Viewer或者有图像的Viewer。
'               1) 如果是空Viewer，直接创建一个Viewer，生成结果图后添加到Viewer中。
'               2) 如果是有图像的Viewer，则把原有序列保存到ZLMPRImages(2,3)中，在此Viewer中显示MPR结果图。
'       存储MPR之前图像的结构包含以下内容：
'           1、ZLShowSeriesInfos    --- 原有的ZLShowSeriesInfos结构
'           2、Images               --- 原来Viewer中已经加载的图像，方便恢复图象中的标注、调窗、缩放等信息
'           3、blnIsMPR             --- 是否当前做MPR的序列，如果是，恢复的时候，不需要替换该序列的内容。
'参数： intType     --- 操作类型 1--MPR序列位置（1，1）；2--MPR结果序列竖线，位置（1，2）；3--MPR结果序列横线，位置（2，1）
'       thisForm    --- 显示图像的窗体
'       thisViewer  --- 进行MPR操作的序列
'返回：成功MPRViewer的Index，失败=0
'时间：2009-7
'------------------------------------------------
    Dim i As Integer
    Dim blnNewViewer As Boolean
    Dim intViewerIndex As Integer
    Dim intRow As Integer, intCol As Integer
    Dim oneImageInfo As clsImageInfo
    Dim oneSeriesInfo As clsSeriesInfo
    Dim MPRViewer As DicomViewer    '摆放好的Viewer
    
    On Error GoTo err
    
    funcReplaceViewer = 0
    If intType < 1 Or intType > 3 Then Exit Function
    If thisForm Is Nothing Then Exit Function
    
    '如果是第一行，第一列，则需要确保thisViewer存在
    If intType = 1 And thisViewer Is Nothing Then Exit Function
    
    '初始化ZLMPRCube
    ZLMPRCube(intType).blnIsMPR = False
    ZLMPRCube(intType).Images.Clear
    Set ZLMPRCube(intType).ZLShowSeriesInfos = Nothing
    ZLMPRCube(intType).intViewerIndex = 0
    
    blnNewViewer = True
    intRow = 1
    intCol = 1
    If intType = 2 Then
        intCol = 2
    ElseIf intType = 3 Then
        intRow = 2
    End If
    '先判断指定位置是否有Viewer
    For i = 1 To ZLShowSeriesInfos.Count
        If ZLShowSeriesInfos(i).intRow = intRow And ZLShowSeriesInfos(i).intCol = intCol And ZLShowSeriesInfos(i).ImageInfos.Count <> 0 Then
            blnNewViewer = False
            intViewerIndex = i
            Exit For
        End If
    Next i
    
    '如果指定位置有Viewer,则需要保存这个Viewer中的图像
    If blnNewViewer = False Then
        '是MPR序列，且正好左上角的Viewer就是MPR序列，则不需要保存Viewer的内容
        If Not thisViewer Is Nothing Then
            If intType = 1 And thisViewer.Index = intViewerIndex Then
                ZLMPRCube(intType).blnIsMPR = True
                Set MPRViewer = thisViewer
            End If
        End If
        
        If ZLMPRCube(intType).blnIsMPR = False Then '需要保存原图和ZLShowSeriesInfos结构
            '复制图像
            Set MPRViewer = thisForm.Viewer(intViewerIndex)
            For i = 1 To MPRViewer.Images.Count
                ZLMPRCube(intType).Images.Add MPRViewer.Images(i)
            Next i
            
            '复制ZLShowSeriesInfos结构
            Set oneSeriesInfo = funGetNewSeriesInfo
            Call funCopySeriesInfo(ZLShowSeriesInfos(intViewerIndex), oneSeriesInfo)
            Set ZLMPRCube(intType).ZLShowSeriesInfos = oneSeriesInfo
            Set ZLMPRCube(intType).ZLShowSeriesInfos.ImageInfos = Nothing
            
            '复制ImageInfos的信息
            For i = 1 To ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
                Set oneImageInfo = funGetNewImageInfo
                Call funCopyImageInfo(ZLShowSeriesInfos(intViewerIndex).ImageInfos(i), oneImageInfo)
                ZLMPRCube(intType).ZLShowSeriesInfos.ImageInfos.Add oneImageInfo
            Next i
        End If
    Else    '指定位置没有Viewer，则创建一个Viewer，并且把这个Viewer摆放到指定位置
        '创建一个Viewer
        intViewerIndex = funcCeateAViewer(1, thisForm)
        '摆放一个Viewer
        Call subPlaceAViewer(thisForm, intViewerIndex, intRow, intCol)
        Set MPRViewer = thisForm.Viewer(intViewerIndex)
    End If
    ZLMPRCube(intType).intViewerIndex = intViewerIndex
    funcReplaceViewer = MPRViewer.Index
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function funGetMPRResultImage(la As DicomLabel, thisViewer As DicomViewer, intToltalHeight As Integer, _
    intType As Integer) As DicomImage
'------------------------------------------------
'功能： 根据传入的控制线，对thisViewer内的图像进行重建，并将返回重建的结果图
'参数： la          --- 进行重建的控制线
'       thisViewer  --- 进行重建的图像所在的Viewer
'       intToltalHeight  --- 进行重建图像相叠加的总体高度
'       intType     --- 图像类型：1--竖线，2--横线，直接填写到图像号中
'返回：重建结果图，如果失败，返回nothing
'时间：2009-7
'------------------------------------------------
    Dim LineLong() As POINTAPI      '保存MPR控制线中每个点的坐标的数组
    Dim iPointsCount As Long        '记录MPR控制线中点的总数量
    Dim iImagesCount As Long        '记录Viewer中图像的总数量
    Dim lines() As Integer          '保存图像灰度值的二维数组
    Dim NewLines() As Integer       '保存重建后图像灰度值的二维数组
    Dim i As Long, j As Long
    Dim resImage As DicomImage      '结果图像
    Dim v As Variant
    
    Set funGetMPRResultImage = Nothing
    
    On Error GoTo err
    
    If thisViewer.Images.Count <= 0 Then Exit Function
    
    '获取标注线中每个点的坐标位置数组
    Call subGetArray(la, thisViewer.Images(1), LineLong)
    iPointsCount = UBound(LineLong)
    iImagesCount = thisViewer.Images.Count
    
    '重新定义原图图像灰度值二维数组
    ReDim lines(iPointsCount, iImagesCount) As Integer
    '重新定义重建后灰度值二维数组
    ReDim NewLines(iPointsCount, intToltalHeight) As Integer
    
    '根据MPR控制线所在点的数组，获取图像点的灰度值
    If SafeArrayGetDim(aPixels) = 0 Then
        'MPR的缓存三维数组维度=0，说明超出内存许可，将直接使用图像数据做重建，图像越多，重建越慢
        For i = 1 To iImagesCount
            v = thisViewer.Images(i).Pixels
            For j = 1 To iPointsCount
                lines(j, i) = v(LineLong(j).x, LineLong(j).y, 1)
            Next j
        Next i
    Else
        '使用三维数组做MPR重建，每次重建速度在1秒内
        For i = 1 To iImagesCount
            For j = 1 To iPointsCount
                lines(j, i) = aPixels(LineLong(j).x, LineLong(j).y, i)
            Next j
        Next i
    End If
    
    '根据层厚将采集回来的直线插值成一个连续图像，同时做平滑处理
    Call subACRebuild(lines, NewLines)
    '生成新图像
    Set resImage = thisViewer.Images(1).SubImage(0, 0, thisViewer.Images(1).sizeX, thisViewer.Images(1).sizeY, 1, 1)
    
    '删掉一些无用的位置属性
    resImage.Attributes.Remove &H18, &H50
    resImage.Attributes.Remove &H18, &H1110
    resImage.Attributes.Remove &H18, &H1111
    resImage.Attributes.Remove &H18, &H1120     'Tilt
    resImage.Attributes.Remove &H18, &H1140     'Rotation Direction
    resImage.Attributes.Remove &H18, &H5100     'Patient Position
    resImage.Attributes.Remove &H20, &H32       'Image Position(Patient)
    resImage.Attributes.Remove &H20, &H37       'Image Orientation (Patient)
    resImage.Attributes.Remove &H20, &H1041     'Slice Location
    
    '设置结果图的属性
    resImage.Attributes.Add &H28, &H10, intToltalHeight
    resImage.Attributes.Add &H28, &H11, iPointsCount
    If intType = 1 Then
        resImage.Attributes.Add &H20, &H11, LineLong(1).y
    Else
        resImage.Attributes.Add &H20, &H11, LineLong(1).x
    End If
    resImage.Attributes.Add &H20, &H13, intType
    resImage.Pixels = NewLines
    resImage.width = thisViewer.Images(1).width
    resImage.Level = thisViewer.Images(1).Level
    
    '返回结果图
    Set funGetMPRResultImage = resImage
    
    Exit Function
err:
    If err.Number = 61706 Or err.Number = -2147417848 Then
        err.Description = "内存不足，无法进行MPR重建，请重启计算机或者增加内存后重试。"
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set funGetMPRResultImage = Nothing
End Function

Public Function funGetMPRImageAndShow(la As DicomLabel, thisForm As frmViewer, MPRViewer As DicomViewer, _
    MPRImageIndex As Integer, ShowViewerIndex As Integer, intToltalHeight As Integer, intType As Integer, _
    blnFirst As Boolean, blnChangeLa As Boolean) As Boolean
'------------------------------------------------
'功能： 根据传入的控制线，对MPRViewer内的图像进行重建，并将返回重建的结果图显示到ShowViewerIndex指定的Viewer中
'参数： la          --- 进行重建的控制线
'       thisForm    --- 显示图像的窗体
'       MPRViewer   --- 重建图像所在的Viewer
'       MPRImageIndex --- 重建图像在Viewer中的Index
'       ShowViewerIndex --- 重建结果图所显示的Viewer的Index
'       intToltalHeight --- 总体高度
'       intType     --- 图像类型：1--竖线，2--横线，直接填写到图像号中
'       blnFirst    --- 是否第一次调用，如果是第一次调用，则不记录原来的窗宽床位等图像状态
'       blnChangeLa --- 是否改变MPR控制线，如果只改变图像没动控制线，则不需要重新生成MPR图，只重画对应线即可
'返回：重建结果图，如果失败，False
'时间：2009-7
'------------------------------------------------
    
    Dim resImage As DicomImage
    Dim resImages As New DicomImages
    Dim imgOld As DicomImage
    Dim imgNew As DicomImage
    Dim thisImage As DicomImage
    Dim dblZoom As Double
    Dim lngScrollX As Long
    Dim lngScrollY As Long
    Dim lngWWidth As Long
    Dim lngWLevel As Long
    Dim blnStretchToFit As Boolean
    Dim dblScale As Double
    
    On Error GoTo err
    
    If ShowViewerIndex >= thisForm.Viewer.Count Then
        funGetMPRImageAndShow = False
        Exit Function
    End If
    
    If blnChangeLa = True Then   '移动了MPR控制线，要产生新的结果图
        '根据传入的控制线，对一个Viewer内的图像进行重建，并返回重建结果图
        Set resImage = funGetMPRResultImage(la, MPRViewer, intToltalHeight, intType)
    Else
        Set resImage = thisForm.Viewer(ShowViewerIndex).Images(1)
    End If
    
    '显示结果图
    If resImage Is Nothing Then
        funGetMPRImageAndShow = False
        Exit Function
    Else
    
        Set imgOld = Nothing
        If blnChangeLa = True Then  '改变了MPR控制线，产生新的结果图
            resImages.Clear
            resImages.Add resImage
            '记录原来的图像状态
            If thisForm.Viewer(ShowViewerIndex).Images.Count > 0 And blnFirst = False Then
                Set imgOld = thisForm.Viewer(ShowViewerIndex).Images(1)
                blnStretchToFit = imgOld.StretchToFit
                dblZoom = imgOld.ActualZoom
                lngScrollX = imgOld.ActualScrollX
                lngScrollY = imgOld.ActualScrollY
                lngWWidth = imgOld.width
                lngWLevel = imgOld.Level
            End If
            Call funShowTempImages(thisForm, resImages, ShowViewerIndex)
        End If
        
        If thisForm.Viewer(ShowViewerIndex).Images.Count > 0 Then
            Set imgNew = thisForm.Viewer(ShowViewerIndex).Images(1)
            imgNew.Refresh False
        End If
        
        '恢复原来图像的状态
        If Not imgOld Is Nothing And Not imgNew Is Nothing Then
            imgNew.StretchToFit = blnStretchToFit
            imgNew.Zoom = dblZoom
            imgNew.ScrollX = lngScrollX
            imgNew.ScrollY = lngScrollY
            imgNew.width = lngWWidth
            imgNew.Level = lngWLevel
        End If
        
        '根据条件确定是否显示MPR辅助线
        If blnShowMPRLine = True And Not imgNew Is Nothing Then
            '画轴位和矢冠状位投影线
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).top = MPRImageIndex / MPRViewer.Images.Count * imgNew.sizeY
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).left = 0
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).width = imgNew.sizeX
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).height = 0
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).Visible = True
            
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).top = 0
            
            Set thisImage = MPRViewer.Images(MPRImageIndex)
            '按照生成控制线的规则，来确定中心点在投影线中的位置
            If Abs(la.width) > Abs(la.height) Then  '看成横线
                If la.width < 0 Then
                    dblScale = 1 - Abs((thisImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left - la.left) / la.width)
                ElseIf la.width > 0 Then
                    dblScale = Abs((thisImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left - la.left) / la.width)
                End If
            Else    '看成竖线
                If la.height < 0 Then
                    dblScale = 1 - Abs((thisImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top - la.top) / la.height)
                ElseIf la.height > 0 Then
                    dblScale = Abs((thisImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top - la.top) / la.height)
                End If
            End If

            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).left = dblScale * imgNew.sizeX
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).height = imgNew.sizeY
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).width = 0
            
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).Visible = True
        End If
        
        funGetMPRImageAndShow = True
    End If
        
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub subMPRReFillImagesToViewer(intType As Integer, thisForm As frmViewer)
'------------------------------------------------
'功能： 把MPR中被临时保存下来的图像恢复到原来的Viewer中
'参数： intType     --- MPR的序列类型，1--MPR序列；2--MPR竖线结果序列；3 -- MPR横线结果序列
'       thisForm    --- 显示图像的窗口
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim intViewerIndex As Integer
    Dim oneImageInfo As clsImageInfo
    Dim i As Integer
    
    If intType < 1 Or intType > 3 Then Exit Sub
    '没有图像则要删除这个Viewer
    
    On Error GoTo err
    
    intViewerIndex = ZLMPRCube(intType).intViewerIndex
    
    If thisForm.Viewer.Count <= intViewerIndex Then Exit Sub
    If intViewerIndex = 0 Then Exit Sub
    
    thisForm.Viewer(intViewerIndex).Images.Clear
    
    '如果图像数量为0，表示这个位置原本是没有Viewer的，需要卸载这个Viewer
    If ZLMPRCube(intType).Images.Count = 0 Then
        Call subUnloadViewer(intViewerIndex, thisForm)
    Else
        '用添加临时图像的方式来恢复图像
        
        Call funShowTempImages(thisForm, ZLMPRCube(intType).Images, intViewerIndex)
        '复制ZLShowSeriesInfos信息
        Call funCopySeriesInfo(ZLMPRCube(intType).ZLShowSeriesInfos, ZLShowSeriesInfos(intViewerIndex))
        Set ZLShowSeriesInfos(intViewerIndex).ImageInfos = Nothing
        
        '复制ImageInfos的信息
        For i = 1 To ZLMPRCube(intType).ZLShowSeriesInfos.ImageInfos.Count
            Set oneImageInfo = funGetNewImageInfo
            Call funCopyImageInfo(ZLMPRCube(intType).ZLShowSeriesInfos.ImageInfos(i), oneImageInfo)
            ZLShowSeriesInfos(intViewerIndex).ImageInfos.Add oneImageInfo
        Next i
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subUnloadViewer(ByVal intViwerIndex As Integer, thisForm As frmViewer)
'------------------------------------------------
'功能： 关闭序列，如果Viewer是最后一个的话，则卸载这个Viewer及其相关的内容
'       如果Viewer是中间的某个Viewer，则清除Viewer及其相关内容中的图像，但是不卸载
'参数： intViwerIndex    --- 要卸载的Viewer的索引
'       thisForm    --- 显示图像的窗口
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim oneSeriesInfo As clsSeriesInfo
    
    On Error GoTo err
    
    If intViwerIndex > thisForm.Viewer.Count Then Exit Sub
    
    '初始化通用的公共参数
    thisForm.intSelectedSerial = 0
    thisForm.oldSelectedImageIndex = 0
    thisForm.oldSelectedSerial = 0
    Set thisForm.SelectedImage = Nothing
    thisForm.SelectedImageIndex = 0
    thisForm.txtText.Visible = False
    Set thisForm.SelectedLabel = Nothing
    
    '判断这个Viewer是否最后一个，如果是最后一个Viewer，则同时删除对应的ZLShowSeriesInfos结构
    '如果不是最后一个Viewer，则只清空ZLShowSeriesInfos结构中的图像和信息
    If thisForm.Viewer.Count - 1 = intViwerIndex Then
        '最后一个Viewer，卸载Viewer，及其相关部分
        Call subUnloadLastViewer(thisForm)
    Else
        thisForm.MSFViewer.TextMatrix(intViwerIndex, 1) = False
        Set oneSeriesInfo = funGetNewSeriesInfo
        Call funCopySeriesInfo(oneSeriesInfo, ZLShowSeriesInfos(intViwerIndex))
        Set ZLShowSeriesInfos(intViwerIndex).ImageInfos = Nothing
        '清空Viewer中的图像
        thisForm.Viewer(intViwerIndex).Images.Clear
        thisForm.VScro(intViwerIndex).Visible = False
        'Viewer的位置和可见性需要改变
        thisForm.Viewer(intViwerIndex).Visible = False  '就不会触发Viewer的事件了
        thisForm.Viewer(intViwerIndex).Tag = 0
        thisForm.Viewer(intViwerIndex).left = 1
        thisForm.Viewer(intViwerIndex).top = 1
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subSeriesInPhase(intViewerIndex As Integer, thisForm As frmViewer, img As DicomImage, intType As Integer)
'------------------------------------------------
'功能： 被选中的多个序列，保证其中的图像内容同步
'参数： intViewerIndex  --- 需要同步的序列索引
'       thisForm    --- 显示图像的窗口
'       img         --- 进行同步的标准图像
'       intType     --- 同步类型，宏定义
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim vTemp As DicomViewer
    Dim i As Integer
    
    '如果不是在图像内容同步的状态，则退出
    If Button_miImageInPhase = False Then Exit Sub
    
    On Error GoTo err
    
    For Each vTemp In thisForm.Viewer
        If vTemp.Visible = True Then
            If (ZLShowSeriesInfos(intViewerIndex).Selected = True And ZLShowSeriesInfos(vTemp.Index).Selected = True) Or vTemp.Index = intViewerIndex Then
                For i = 1 To vTemp.Images.Count
                    Call subImageInPhase(vTemp.Images(i), img, intType)
                Next i
                vTemp.Refresh
            End If
        End If
    Next
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subImageInPhase(img As DicomImage, SampleImg As DicomImage, intType As Integer)
'------------------------------------------------
'功能： 两个图像之间状态同步
'参数： Img         --- 需要进行同步的图像
'       SampleImg   --- 进行同步的标准图像
'       intType     --- 同步类型，宏定义
'返回：无
'时间：2009-7
'------------------------------------------------
    
    On Error GoTo err
    
    Select Case intType
    Case IMG_SYN_All                '全部同步
        img.width = SampleImg.width
        img.Level = SampleImg.Level
        img.StretchToFit = SampleImg.StretchToFit
        img.ScrollX = SampleImg.ScrollX
        img.ScrollY = SampleImg.ScrollY
        img.Zoom = SampleImg.Zoom
        If img.Labels.Count >= G_INT_SYS_LABEL_WWWL Then
            img.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & img.width & "-L:" & img.Level
        End If
        img.FlipState = SampleImg.FlipState
        img.RotateState = SampleImg.RotateState
        img.FilterLength = SampleImg.FilterLength
        img.UnsharpEnhancement = SampleImg.UnsharpEnhancement
        img.UnsharpLength = SampleImg.UnsharpLength
        If img.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
            If img.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then  '更新标尺单位
                Call UpdateRuler(img, True)
            End If
        End If
    Case IMG_SYN_WINDOW             '调窗同步
        img.width = SampleImg.width
        img.Level = SampleImg.Level
        If img.Labels.Count >= G_INT_SYS_LABEL_WWWL Then
            img.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & img.width & "-L:" & img.Level
        End If
    Case IMG_SYN_ZOOMPAN            '缩放、漫游同步
        img.StretchToFit = SampleImg.StretchToFit
        img.Zoom = SampleImg.Zoom
        img.ScrollX = SampleImg.ScrollX
        img.ScrollY = SampleImg.ScrollY
        If img.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
            If img.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then  '更新标尺单位
                Call UpdateRuler(img, True)
            End If
        End If
    Case IMG_SYN_ROTATE             '旋转同步
        img.RotateState = SampleImg.RotateState
    Case IMG_SYN_FLIP               '镜像同步
        img.FlipState = SampleImg.FlipState
        img.RotateState = SampleImg.RotateState
    Case IMG_SYN_FILTER             '滤镜同步
        img.FilterLength = SampleImg.FilterLength
        img.UnsharpEnhancement = SampleImg.UnsharpEnhancement
        img.UnsharpLength = SampleImg.UnsharpLength
    End Select
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subUnloadLastViewer(thisForm As frmViewer)
'------------------------------------------------
'功能： 卸载窗体中的最后一个Viewer，及其相关的滚动条，ZLShowSeriesInfos，MSFViewer
'参数： thisForm    --- 显示图像的窗体
'返回：无
'时间：2009-7
'------------------------------------------------
    On Error GoTo err
    
    Call Unload(thisForm.Viewer(thisForm.Viewer.Count - 1))
    Call Unload(thisForm.VScro(thisForm.VScro.Count - 1))
    '清理ZLShowSeriesInfos
    ZLShowSeriesInfos.Remove ZLShowSeriesInfos.Count
    '清理MSF结构
    thisForm.MSFViewer.Rows = thisForm.MSFViewer.Rows - 1
        
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subScaleViewer(thisViewer As DicomViewer, img As DicomImage, lngOldWidth As Long, lngOldHeight As Long)
'------------------------------------------------
'功能： 当Viewer的宽度和高度改变后，对一个Viewer中StretchToFit=False的图像位置和缩放进行修正
'参数： thisViewer  --- 存放图像的Viewer
'       img         --- 需要修正的图像
'       lngOldWidth --- Viewer原来的宽度
'       lngOldHeight--- Viewer原来的高度
'返回：无
'时间：2009-7
'------------------------------------------------
    Dim i As Integer
    
    If thisViewer.Images.Count = 0 Then Exit Sub
    
    On Error GoTo err
    
    '对其中的一个图像进行位置修正
    If img.StretchToFit = False Then
        Call subScaleImage(img, thisViewer, lngOldWidth, lngOldHeight)
        
        '根据这个修正结果，对Viewer中的所有图像做修正
        For i = 1 To thisViewer.Images.Count
            thisViewer.Images(i).StretchToFit = False
            thisViewer.Images(i).Zoom = img.ActualZoom
            thisViewer.Images(i).ScrollX = img.ActualScrollX
            thisViewer.Images(i).ScrollY = img.ActualScrollY
        Next i
        '把修正的结果记录到ZLShowSeriesInfos结构中
        ZLShowSeriesInfos(thisViewer.Index).ScrollX = img.ActualScrollX
        ZLShowSeriesInfos(thisViewer.Index).ScrollY = img.ActualScrollY
        ZLShowSeriesInfos(thisViewer.Index).Zoom = img.ActualZoom
    End If
     
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function FunLogIn(str类型 As String) As String
'功能：对程序进行注册，如果注册成功，则返回注册时间
'参数： str类型 ---'在注册码中使用的类型名称
'返回值：注册成功注册日期；注册失败返回空

    Dim intNUM As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP地址 As String         '需要注册的IP地址
    
    On Error GoTo err
        
    strIP地址 = OS.IP
    
    
    '从注册码中提取授权的数量，-1--无限制；0--禁止；X（X>0）--按照数量控制
    
    If str类型 = LOGIN_TYPE_医技观片站 Then
        intNUM = gint医技观片站数量
    ElseIf str类型 = LOGIN_TYPE_胶片打印机 Then
        intNUM = gint胶片打印机
    Else
        intNUM = 0
    End If
    
    
    'intNUM >0 ,则调用过程注册程序
    If intNUM > 0 Then  '按数量限制
        strSQL = "Zl_影像操作记录_Update('" & strIP地址 & "','" & str类型 & "'," & intNUM & ")"
        zlDatabase.ExecuteProcedure strSQL, "注册" & str类型
        '检查注册是否成功
        strSQL = "Select 启动时间,IP地址 from 影像操作记录 where  类型=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取启动时间", str类型)
        If rsTemp.RecordCount <= intNUM Then
            rsTemp.Filter = "IP地址='" & strIP地址 & "'"
            If rsTemp.RecordCount = 1 Then  '注册成功
                FunLogIn = rsTemp!启动时间
                Exit Function
            End If
        End If
    ElseIf intNUM = -1 Then '无限制
        FunLogIn = Now
        Exit Function
    Else    '=0，或者其他值，禁止，不做任何处理，后面有提示
        
    End If
    '注册失败，可能是两个原因：
    '1、注册的数量超过了许可的数量，无法注册IP地址
    '2、直接通过SQL向表中添加了IP地址，导致表中的记录总数量超过了许可的数量
    Call MsgBox("打开的" & str类型 & "超过您购买的总数量（" & intNUM & "）。请向软件供应商联系。", vbOKOnly, gstrSysName)
    FunLogIn = ""
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function FunLogOut(str类型 As String, str启动时间 As String) As Boolean
'功能：退出程序的时候，检查程序是否合法注册过，避免有人通过触发器等手段定时删除“影像操作记录”表中的记录。
'参数： str类型 ---'在注册码中使用的类型名称
'       str启动时间 --- 注册工作站时返回的时间
'返回值：合法注册True；非法启动的False
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP地址 As String         '需要注册的IP地址
    Dim intNUM As Integer
    
    On Error GoTo err
    strIP地址 = OS.IP
    
    '启动时间为空，表示注册失败，没有正常启动，因此退出的时候不再检测数据库
    If str启动时间 = "" Then
        FunLogOut = True
        Exit Function
    End If
    
    '从注册码中提取授权的数量，-1--无限制；0--禁止；X（X>0）--按照数量控制
    If str类型 = LOGIN_TYPE_医技观片站 Then
        intNUM = gint医技观片站数量
    ElseIf str类型 = LOGIN_TYPE_胶片打印机 Then
        intNUM = gint胶片打印机
    Else
        intNUM = 0
    End If
    
    If intNUM > 0 Then '按照数量控制
        strSQL = "Select 启动时间 from 影像操作记录 where IP地址=[1] and 类型=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取启动时间", strIP地址, str类型)
        If rsTemp.EOF = False Then
            FunLogOut = True
        Else
            '对比启动时间和数据库的时间，如果不是同一天，说明是前一天开启程序后注册信息被删除了，
            '这种情况认为是合法注册
            strSQL = "Select sysdate from dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取数据库时间")
            If Format(rsTemp!sysdate, "yyyy-mm-dd") <> Format(str启动时间, "yyyy-mm-dd") Then
                FunLogOut = True
            Else
                FunLogOut = False
            End If
        End If
    ElseIf intNUM = -1 Then '无限制
        FunLogOut = True
    Else    '=0，或者其他值，禁止
        FunLogOut = False
    End If
    
    If FunLogOut = False Then
        Call MsgBox("打开的" & str类型 & "超过您购买的总数量（" & intNUM & "）。请向软件供应商联系。", vbOKOnly, gstrSysName)
    End If
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function getLicenseCount(strLicenseName As String) As Integer
'读取授权的数量,等授权程序修改后，再修改这个过程
'参数： strLicenseName --- 授权名称
    
    Dim strLiceseCount As String
    
    On Error GoTo err
    
    strLiceseCount = zl9ComLib.zlRegInfo(strLicenseName)
    If strLiceseCount = "" Then '无限制
        getLicenseCount = -1
    ElseIf Val(strLiceseCount) > 0 Then '按照数量限制
        getLicenseCount = Val(strLiceseCount)
    Else '禁止
        getLicenseCount = 0
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadImage(strPath As String, Optional blnSilent As Boolean = False) As DicomImage
'功能：读取一个文件，返回DICOM图像。
'参数： strPath -- 文件路径
'       blnSilent -- 不提示
'返回：DICOM图像
    Dim imgs As New DicomImages
    Dim img As New DicomImage
    
    On Error Resume Next
    err.Clear
    imgs.Clear
    Set img = imgs.ReadFile(strPath)
    If err <> 0 Then        '读取失败，说明不是DICOM文件
        err.Clear
        img.FileImport strPath, ""
        If err <> 0 Then    '导入失败，说明文件不是BMP、JPG、AVI格式的。
            If blnSilent = False Then
                MsgBox "文件" & strPath & "不能打开！", vbInformation, gstrSysName
            End If
            Debug.Print "文件" & strPath & "不能打开！"
            Set ReadImage = Nothing
            Exit Function
        End If
    End If
    Set ReadImage = img
End Function

Private Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function

Private Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function

Public Sub PrintFilmBeep(intTrack As Integer)
'------------------------------------------------
'功能： 胶片打印时的提示声音
'参数： intTrack --- 音轨代码 1-添加图像；2-打印胶片
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    If blnPrintFilmBeep Then
        If intTrack = 1 Then
            Call Beep(BEEP_Do0, 100)
            Call Beep(BEEP_Re, 100)
            Call Beep(BEEP_Mi, 100)
        Else
            Call Beep(BEEP_Do0, 150)
            Call Beep(BEEP_Mi, 150)
            Call Beep(BEEP_Sol, 150)
            Call Beep(BEEP_Do1, 150)
        End If
    End If
    
    Exit Sub
err:
    '出错后不处理
End Sub

Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'------------------------------------------------
'功能：当指定目录的大小达到一定百分比时，清空该目录
'参数： strCacheFolder--需要检查是否清空的目录
'返回：无
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        zl9ComLib.zlCommFun.ShowFlash "正清空图像缓冲目录，请等待！", frmMain
        objCurFolder.Delete True
        zl9ComLib.zlCommFun.StopFlash
    End If
End Sub

Private Function FileIsOccupied(ByVal FilePath As String) As Boolean
'------------------------------------------------
'功能：判断文件是否正在被其他进程占用，通过独占方式打开来判断
'参数： FilePath--需要打开的文件
'返回：True--被占用；False--不被占用
'------------------------------------------------
    Dim fFile     As Integer
    
    fFile = FreeFile
    
    On Error GoTo ErrOpen
    Open FilePath For Binary Lock Read Write As fFile
    Close fFile
    Exit Function
ErrOpen:
    FileIsOccupied = True
End Function

Private Sub TimeDelay(lngTimeDelay As Long)
'------------------------------------------------
'功能：延时
'参数： lngTimeDelay--需要延时的时间长度
'返回：无
'------------------------------------------------
    Dim Savetime As Double
    
    On Error GoTo err
    Savetime = timeGetTime '记下开始时的时间
    While timeGetTime < Savetime + lngTimeDelay '循环等待
'    DoEvents '转让控制权，以便让操作系统处理其它的事件。
    Wend
    Exit Sub
err:
    
End Sub

Public Function funGetCandSImageAndShow(la As DicomLabel, imageViewer As DicomViewer, _
    axialViewer As DicomViewer, resultViewer As DicomViewer, intToltalHeight As Integer, _
    intType As Integer, blnFirst As Boolean, blnChangeLa As Boolean) As Boolean
'------------------------------------------------
'功能： 根据控制线la，对imageViewer内的图像进行冠状位或矢状位重建，结果图显示到ResultViewer中
'参数： la          --- 进行重建的控制线
'       imageViewer   --- 重建图像所在的Viewer，在父窗体中
'       axialViewer -- 轴位图像所在的Viewer，在斜面重建窗体中
'       resultViewer -- 重建结果图所在的Viewer，在斜面重建窗体中
'       intToltalHeight --- 总体高度
'       intType     --- 图像类型：1--竖线，2--横线，直接填写到图像号中
'       blnFirst    --- 是否第一次调用，如果是第一次调用，则不记录原来的窗宽床位等图像状态
'       blnChangeLa --- 是否改变MPR控制线，如果只改变图像没动控制线，则不需要重新生成MPR图，只重画对应线即可
'返回：重建结果图，如果失败，False
'------------------------------------------------
    
    Dim resImage As DicomImage
    Dim resImages As New DicomImages
    Dim imgOld As DicomImage
    Dim imgNew As DicomImage
    Dim dblZoom As Double
    Dim lngScrollX As Long
    Dim lngScrollY As Long
    Dim lngWWidth As Long
    Dim lngWLevel As Long
    Dim blnStretchToFit As Boolean
    Dim img As DicomImage
    
    On Error GoTo err
    
    '获取重建结果图，如果移动了控制线，就产生新图像，否则还使用原来的图像
    If blnChangeLa = True Then   '移动了MPR控制线，要产生新的结果图
        '根据传入的控制线，对一个Viewer内的图像进行重建，并返回重建结果图
        Set resImage = funGetMPRResultImage(la, imageViewer, intToltalHeight, intType)
    Else
        Set resImage = resultViewer.Images(1)
    End If
    
    '显示结果图
    If resImage Is Nothing Then
        funGetCandSImageAndShow = False
        Exit Function
    Else
        Set imgOld = Nothing
        If blnChangeLa = True Then  '改变了MPR控制线，产生新的结果图
            resImages.Clear
            resImages.Add resImage
            '记录原来的图像状态，新的重建结果图继续使用这些状态，第一次重建，不需要记录
            If resultViewer.Images.Count > 0 And blnFirst = False Then
                Set imgOld = resultViewer.Images(1)
                blnStretchToFit = imgOld.StretchToFit
                dblZoom = imgOld.ActualZoom
                lngScrollX = imgOld.ActualScrollX
                lngScrollY = imgOld.ActualScrollY
                lngWWidth = imgOld.width
                lngWLevel = imgOld.Level
            End If
            '将图像添加到结果Viewer中
            resultViewer.Images.Clear
            resultViewer.Images.Add resImage
            
            Set img = resultViewer.Images(1)
            img.Tag = 1
            If img.Labels.Count = 0 Then
                Call subInitAImage(img, 0, resultViewer)
            End If
            
        End If
        
        If resultViewer.Images.Count > 0 Then
            Set imgNew = resultViewer.Images(1)
            imgNew.Refresh False
        End If
        
        '恢复原来图像的状态
        If Not imgOld Is Nothing And Not imgNew Is Nothing Then
            imgNew.StretchToFit = blnStretchToFit
            imgNew.Zoom = dblZoom
            imgNew.ScrollX = lngScrollX
            imgNew.ScrollY = lngScrollY
            imgNew.width = lngWWidth
            imgNew.Level = lngWLevel
        End If
            
        '画重建结果图的控制线
        If Not imgOld Is Nothing Then
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).left = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).left
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).top = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).top
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).width = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).width
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).height = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).height
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).Visible = True
            
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).left = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).left
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).top = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).top
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).width = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).width
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).height = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).height
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).Visible = True
            imgNew.Refresh (False)
        Else
            Call subMPRSlopeDrawResultControlLabels(la, imgNew, imageViewer, axialViewer)
        End If
        
        funGetCandSImageAndShow = True
    End If
        
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub subMPRSlopeDrawResultControlLabels(la As DicomLabel, imgNew As DicomImage, imageViewer As DicomViewer, _
        axialViewer As DicomViewer)
'------------------------------------------------
'功能： 画斜面重建结果图的控制线
'参数： la -- 进行重建的轴位控制线
'       imgNew -- 重建结果图
'       imageViewer -- 原图所在的Viewer，在主窗口中
'       axialViewer -- 轴位图像所在的Viewer，在斜面重建窗口中
'返回:无
'------------------------------------------------
    Dim dblScale As Double
    
    On Error GoTo err
    
    '根据条件确定是否显示MPR辅助线
    If blnShowMPRLine = True And Not imgNew Is Nothing Then
        '画轴位和矢冠状位投影线
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).top = axialViewer.Images(1).Tag / imageViewer.Images.Count * imgNew.sizeY
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).left = 0
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).width = imgNew.sizeX
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).height = 0
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).Visible = True
        
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).top = 0
        
        '按照生成控制线的规则，来确定中心点在投影线中的位置
        If Abs(la.width) > Abs(la.height) Then  '看成横线
            If la.width < 0 Then
                dblScale = 1 - Abs((axialViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_POINT_O).left - la.left) / la.width)
            ElseIf la.width > 0 Then
                dblScale = Abs((axialViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_POINT_O).left - la.left) / la.width)
            End If
        Else    '看成竖线
            If la.height < 0 Then
                dblScale = 1 - Abs((axialViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_POINT_O).top - la.top) / la.height)
            ElseIf la.height > 0 Then
                dblScale = Abs((axialViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_POINT_O).top - la.top) / la.height)
            End If
        End If

        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).left = dblScale * imgNew.sizeX
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).height = imgNew.sizeY
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).width = 0
        
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).Visible = True
        Call imgNew.Refresh(False)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub



