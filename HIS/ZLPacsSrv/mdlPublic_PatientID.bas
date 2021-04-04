Attribute VB_Name = "mdlPublic"
Option Explicit
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Public Const ATTR_检查日期 As String = "Study Date"
Public Const ATTR_检查时间 As String = "Study Time"
Public Const ATTR_序列日期 As String = "Series Date"
Public Const ATTR_序列时间 As String = "Series Time"
Public Const ATTR_影像类别 As String = "Modality"
Public Const ATTR_设备商 As String = "Manufacturer"
Public Const ATTR_检查设备 As String = "Manufacturer's Model Name"

Public gcnAccess As New ADODB.Connection, strBeginDate As String

Public gstrSQL As String

Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Public Sub GetUserInfo()
'功能:得到用户的信息

    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String
    
    rsTemp.CursorLocation = adUseClient
    On Error GoTo errHand
    
    With rsTemp
        strSQL = "select P.*,D.编码 as 部门编码,D.名称 as 部门名称,M.部门ID" & _
                " from 上机人员表 U,人员表 P,部门表 D,部门人员 M " & _
                " Where U.人员id = P.id And P.ID=M.人员ID and  M.缺省=1 and M.部门id = D.id and U.用户名=user"
        .Open strSQL, gcnOracle, adOpenKeyset
                
        If .RecordCount <> 0 Then
            glngUserId = .Fields("ID").Value                '当前用户id
            gstrUserCode = .Fields("编号").Value            '当前用户编码
            gstrUserName = .Fields("姓名").Value            '当前用户姓名
            gstrUserAbbr = IIf(IsNull(.Fields("简码").Value), "", .Fields("简码").Value)          '当前用户简码
            glngDeptId = .Fields("部门id").Value            '当前用户部门id
            gstrDeptCode = .Fields("部门编码").Value        '当前用户
            gstrDeptName = .Fields("部门名称").Value        '当前用户
        Else
            glngUserId = 0
            gstrUserCode = ""
            gstrUserName = ""
            gstrUserAbbr = ""
            glngDeptId = 0
            gstrDeptCode = ""
            gstrDeptName = ""
        End If
        .Close
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0
End Sub

Public Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strFormCaption As String)
'功能：打开记录。同时保存SQL语句
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strFormCaption, gstrSQL)
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Public Sub ExecuteProcedure(ByVal strFormCaption As String)
'功能：执行过程式的SQL语句
    Call SQLTest(App.ProductName, strFormCaption, gstrSQL)
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub SaveImages(Images As DicomImages, ByVal MainDeviceID As String, ByVal BufferDir As String)
'功能：保存图像
    Dim curImage As DicomImage
    Dim i As Integer, iCount As Integer  '保存的图像数
    Dim intSQL As Integer, rsTmp As New ADODB.Recordset
    
    Dim blnAddTmp As Boolean, blnTmp As Boolean
    Dim strAge As String, strBirth As String
    Dim strDirURL As String, strHost As String
    Dim dtReceived As String, dtCurrent As String
    
    Dim ImageType As String, CheckNo As Long, CheckDev As String
    Dim PatientName As String, EnglishName As String, Sex As String, Age As Integer
    Dim CheckUID As String, SeriesUID As String
    
    On Error GoTo DBError
    gcnOracle.BeginTrans
    If gcnAccess.State <> adStateClosed Then gcnAccess.BeginTrans
    
    dtCurrent = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    gstrSQL = "Select 'ftp://'||Decode(用户名,Null,'',用户名||Decode(密码,Null,'',':'||密码))" & _
        "||'@'||IP地址 As Host,'/'||Decode(Ftp目录,Null,'',Ftp目录||'/') As URL " & _
        "From 影像设备目录 " & _
        "Where 设备号='" & MainDeviceID & "'"
    If rsTmp.State <> adStateClosed Then rsTmp.Close
    OpenRecordset rsTmp, "PACS图像保存"
    If rsTmp.EOF Then
        Err.Raise vbObjectError + 1, "PACS图像保存", "设备号设置错误！"
    End If
    strHost = rsTmp("Host"): strDirURL = rsTmp("URL")
    
    iCount = 0
    For Each curImage In Images
        gstrSQL = "Select 图像UID From 影像检查图象 Where 图像UID='" & _
            curImage.InstanceUID & "' Union All Select 图像UID From 影像临时图象 Where 图像UID='" & _
            curImage.InstanceUID & "'"
        If rsTmp.State <> adStateClosed Then rsTmp.Close
        OpenRecordset rsTmp, "PACS图像保存"
        '新图像
        If rsTmp.EOF Then
            gstrSQL = "Select 检查UID From 影像检查记录 Where 检查UID='" & curImage.StudyUID & "'" & _
                " Union All Select 检查UID From 影像临时记录 Where 检查UID='" & curImage.StudyUID & "'"
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            OpenRecordset rsTmp, "PACS图像保存"
            '按病人ID或英文名查找
            If rsTmp.EOF Then
                blnAddTmp = True
                If IsNumeric(curImage.PatientID) Then
                    gstrSQL = "Select Distinct A.医嘱ID,A.发送号 From 影像检查记录 A,病人医嘱发送 B,病人医嘱记录 C" & _
                        " Where A.医嘱ID=B.医嘱ID And A.发送号=B.发送号 And B.医嘱ID=C.ID" & _
                        " And C.病人ID=" & curImage.PatientID & _
                        " And B.执行状态=3 And B.执行过程=2"
                    If rsTmp.State <> adStateClosed Then rsTmp.Close
                    OpenRecordset rsTmp, "PACS图像保存"
                    '与HIS填写的检查记录对应
                    If rsTmp.RecordCount = 1 Then
                        '填入检查UID
                        gstrSQL = "ZL_影像检查记录_SET(" & rsTmp(0) & "," & rsTmp(1) & ",'" & _
                            curImage.StudyUID & "','" & GetImageAttribute(curImage.Attributes, ATTR_检查设备) & "'," & _
                            "to_Date('" & dtCurrent & "','YYYY-MM-DD HH24:MI:SS'),'" & MainDeviceID & "')"
                        ExecuteProcedure "PACS图像保存"
                        blnAddTmp = False
                    End If
                End If
                '插入临时检查记录
                If blnAddTmp Then
                    If IsDate(curImage.DateOfBirthAsDate) Then
                        strAge = CStr(Year(Date) - Year(curImage.DateOfBirthAsDate))
                        strBirth = Format(curImage.DateOfBirthAsDate, "YYYY-MM-DD")
                    Else
                        strAge = "": strBirth = ""
                    End If
                    gstrSQL = "ZL_影像临时检查_INSERT('" & GetImageAttribute(curImage.Attributes, ATTR_影像类别) & "',Null,'" & _
                        curImage.Name & "','" & curImage.Name & "','" & _
                        curImage.Sex & "','" & strAge & "'," & _
                        IIf(Len(strBirth) = 0, "Null", "to_Date('" & strBirth & "','YYYY-MM-DD')") & ",Null,Null,'" & _
                        GetImageAttribute(curImage.Attributes, ATTR_检查设备) & "','" & curImage.StudyUID & "'," & _
                        "to_Date('" & dtCurrent & "','YYYY-MM-DD HH24:MI:SS'),'" & MainDeviceID & "')"
                    ExecuteProcedure "PACS图像保存"
                End If
            End If
            
            gstrSQL = "Select 0 As 临时,接收日期,影像类别,Nvl(检查号,0) As 检查号," & _
                "检查设备,姓名,英文名,性别,Nvl(年龄,'-1') As 年龄,检查UID From 影像检查记录 Where 检查UID='" & curImage.StudyUID & "'" & _
                " Union All Select 1 As 临时,接收日期,影像类别,Nvl(检查号,0) As 检查号," & _
                "检查设备,姓名,英文名,性别,Nvl(年龄,'-1') As 年龄,检查UID From 影像临时记录 Where 检查UID='" & curImage.StudyUID & "'"
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            OpenRecordset rsTmp, "PACS图像保存"
            blnTmp = IIf(rsTmp(0) = 1, True, False) '序列和图像是否放入临时记录中
            dtReceived = Format(rsTmp(1), "yyyyMMdd")
            
            ImageType = Nvl(rsTmp(2)): CheckNo = rsTmp(3): CheckDev = Nvl(rsTmp(4))
            PatientName = Nvl(rsTmp(5)): EnglishName = Nvl(rsTmp(6)): Sex = Nvl(rsTmp(7)): Age = Val(rsTmp(8))
            CheckUID = Nvl(rsTmp(9))
            
            gstrSQL = "Select 序列UID From " & IIf(blnTmp, "影像临时序列", "影像检查序列") & _
                " Where 序列UID='" & curImage.SeriesUID & "'"
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            OpenRecordset rsTmp, "PACS图像保存"
            '插入新的检查序列
            If rsTmp.EOF Then
                gstrSQL = "ZL_影像序列_INSERT('" & curImage.StudyUID & "','" & curImage.SeriesUID & "','" & _
                    curImage.SeriesDescription & "'," & _
                    IIf(blnTmp, 1, 0) & ")"
                ExecuteProcedure "PACS图像保存"
            End If
            
            '插入新的图像
            gstrSQL = "ZL_影像图象_INSERT('" & curImage.InstanceUID & "','" & curImage.SeriesUID & "','" & _
                curImage.SeriesDescription & "'," & _
                IIf(blnTmp, 1, 0) & ")"
            ExecuteProcedure "PACS图像保存"
            
            '保存本地日志
            WriteRecord ImageType, CheckNo, CheckDev, PatientName, EnglishName, Sex, Age, CheckUID, curImage.SeriesUID, blnTmp
            
            '保存图像到缓存目录
            curImage.WriteFile BufferDir & curImage.InstanceUID, True
            WriteToURL BufferDir & curImage.InstanceUID, strHost, strDirURL & _
                dtReceived & "/" & curImage.StudyUID & "/" & curImage.InstanceUID
            Kill BufferDir & curImage.InstanceUID
        Else
            WriteLog 3, vbObjectError + 1, "影像：" & curImage.InstanceUID & "已存在！"
        End If
        iCount = iCount + 1
    Next
    
    If gcnAccess.State <> adStateClosed Then gcnAccess.CommitTrans
    gcnOracle.CommitTrans
    
    For i = 1 To iCount
        Images.Remove 1
    Next
    Exit Sub
DBError:
    If gcnAccess.State <> adStateClosed Then gcnAccess.RollbackTrans
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "检查图像保存"
End Sub

Public Sub WriteToURL(ByVal SrcFileName As String, ByVal DestAddress As String, ByVal DestFileName As String)
'功能：将本地文件保存到远程网络上
    Dim iNet As Object
    
    Set iNet = CreateObject("InetCtls.inet.1")
    iNet.AccessType = 0: iNet.URL = DestAddress
    
    MkDir_Remote DestAddress, DestFileName
    iNet.Execute , "Put " & SrcFileName & " " & DestFileName
    Do While iNet.StillExecuting
        DoEvents
    Loop
End Sub

Public Sub MkDir_Remote(ByVal DestAddress As String, ByVal DestFileName As String)
    Dim iNet As Object, objFile As New Scripting.FileSystemObject, strPath As String
    Dim aNestPath() As Variant, i As Integer
    
    aNestPath = Array()
    
    Set iNet = CreateObject("InetCtls.inet.1")
    iNet.AccessType = 0: iNet.URL = DestAddress
    
    strPath = objFile.GetParentFolderName(DestFileName)
    Do While Len(strPath) > 0
        ReDim Preserve aNestPath(UBound(aNestPath) + 1)
        aNestPath(UBound(aNestPath)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    For i = UBound(aNestPath) To 0 Step -1
        iNet.Execute , "MkDir " & aNestPath(i)
        Do While iNet.StillExecuting
            DoEvents
        Loop
    Next
End Sub

Public Function GetImageAttribute(objAttr As DicomAttributes, ByVal AttrName As String) As Variant
    Dim curAttr As DicomAttribute
    
    GetImageAttribute = ""
    For Each curAttr In objAttr
        If UCase(curAttr.Description) = UCase(AttrName) Then
            If curAttr.Exists Then GetImageAttribute = curAttr.Value
            Exit For
        End If
    Next
End Function

Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer, _
    Optional ByVal MaxRows As Integer = 0, Optional ByVal MaxCols As Integer = 0)
'功能：计算DicomViewer的行列数
    Dim iCols As Integer, iRows As Integer
    
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    If MaxRows > 0 And iRows > MaxRows Then
        iRows = MaxRows
        iCols = CInt(ImageCount / iRows)
        If iRows * iCols < ImageCount Then iCols = iCols + 1
    End If
    If MaxCols > 0 And iCols > MaxCols Then
        iCols = MaxCols
        iRows = CInt(ImageCount / iCols)
        If iRows * iCols < ImageCount Then iRows = iRows + 1
    End If
    If MaxRows > 0 And iRows > MaxRows Then iRows = MaxRows
    
    Rows = iRows: Cols = iCols
End Sub

Public Function ImageExist(Images As DicomImages, SeekImage As DicomImage) As Boolean
    Dim curImage As DicomImage
    
    ImageExist = False
    For Each curImage In Images
        If curImage.InstanceUID = SeekImage.InstanceUID Then ImageExist = True: Exit For
    Next
End Function

Private Sub WriteRecord(ByVal ImageType As String, ByVal CheckNo As Long, ByVal CheckDev As String, _
    ByVal PatientName As String, ByVal EnglishName As String, ByVal Sex As String, Age As Integer, _
    ByVal CheckUID As String, ByVal SeriesUID As String, ByVal ifTmp As Boolean)
    
    Dim rsTmp As ADODB.Recordset, strSQL As String
    If gcnAccess.State = adStateClosed Then Exit Sub
    
    strSQL = "Select * from 影像接收序列 Where 序列UID='" & SeriesUID & "' And 接收时间>cDate('" & _
        strBeginDate & "')"
    Set rsTmp = gcnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        strSQL = "Insert Into 影像接收序列(影像类别,检查号,检查设备,姓名,英文名,性别,年龄,影像数,序列UID,检查UID,对应检查,接收时间)" & _
            " Values('" & ImageType & "'," & IIf(CheckNo = 0, "Null", CheckNo) & ",'" & CheckDev & "','" & _
            PatientName & "','" & EnglishName & "','" & Sex & "'," & IIf(Age = -1, "Null", Age) & ",1,'" & _
            SeriesUID & "','" & CheckUID & "'," & CStr(Not ifTmp) & ",cDate('" & _
            Date & " " & Time() & "'))"
    Else
        strSQL = "Update 影像接收序列 Set 影像数=影像数+1 Where 序列UID='" & SeriesUID & "' And 接收时间>cDate('" & _
        strBeginDate & "')"
    End If
    gcnAccess.Execute strSQL
End Sub

Public Sub WriteLog(ByVal ErrorType As Integer, ErrorNum As Long, ErrorDesc As String)
    Dim strSQL As String
    If gcnAccess.State = adStateClosed Then Exit Sub
    
    strSQL = "Insert Into 错误日志(产生时间,错误类型,错误号,错误信息) " & _
        "Values(cDate('" & Date & " " & Time() & "')," & ErrorType & "," & ErrorNum & ",'" & Replace(ErrorDesc, "'", "''") & "')"
    gcnAccess.Execute strSQL
End Sub
'显示保存目录
Public Function BrowPath(lWindowHwnd As Long, Optional ByVal sTitle As String = "") As String
    Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    With udtBI
        '设置浏览窗口
        .hWndOwner = lWindowHwnd
        '返回选中的目录
        .ulFlags = BIF_RETURNONLYFSDIRS
        If sTitle = "" Then
            .lpszTitle = "请选定开始搜索的文件夹："
        Else
            .lpszTitle = sTitle
        End If
    End With
    '调出浏览窗口
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        '获取路径
        SHGetPathFromIDList lpIDList, sPath
        '释放内存
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    BrowPath = sPath
End Function

