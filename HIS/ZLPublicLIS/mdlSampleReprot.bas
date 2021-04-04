Attribute VB_Name = "mdlSampleReprot"
Option Explicit

Public gstrSysName As String                        '系统名称
Public gstrProductName As String                    'OEM产品名称
Public gstrUnitName As String                       '用户单位名称
Public gcnOracle As New ADODB.Connection                 '公共数据库连接

Public UserInfo As TYPE_USER_INFO

'用户信息
Public Type TYPE_USER_INFO
    ID As Long
    编号 As String
    姓名 As String '人员姓名
    简码 As String
    DeptID As Long '部门ID
    DeptNo As String '部门编号
    DeptName As String '部门名称
    DBUser As String '数据库用户
End Type

Public glngSys As Long                              '系统号
Public glngModule As Long                           '模块号
Public gobjLISInsideComm As Object
Public gobjComLib As Object
Public gobjFile As New FileSystemObject

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.编号 = rsTmp!编号
            UserInfo.简码 = Nvl(rsTmp!简码)
            UserInfo.姓名 = Nvl(rsTmp!姓名)
            UserInfo.DeptID = Nvl(rsTmp!部门ID, 0)
            UserInfo.DeptNo = rsTmp!部门码 & ""
            UserInfo.DeptName = rsTmp!部门名 & ""
            UserInfo.DBUser = rsTmp!用户名 & ""
            GetUserInfo = True
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub InitObjLis()
'判断如果新版LIS部件为空就初始化
    Dim strErr As String
    If gobjLISInsideComm Is Nothing Then
        On Error Resume Next
        Set gobjLISInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLISInsideComm Is Nothing Then
            If gobjLISInsideComm.InitComponentsHIS(glngSys, glngModule, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS部件初始化错误：" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLISInsideComm = Nothing
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub

Public Function writeErrLog(ByVal strObject As String, ByVal strModule As String, ByVal strErrTxt As String, Optional ByVal blnMsg As Boolean) As Boolean
    '错误日志
    'strObject      出错的工程名
    'strModule      出错的模块
    'strTxt         日志内容
    
    Dim objFSO As New FileSystemObject
    Dim objFolder As Folder
    Dim objFiles As Files
    Dim objFile As File
    Dim strFilename As String
    Dim dFileCreatDate As Date
    Dim objTxt As TextStream
    Dim strFolder As String         '错误日志目录
    Dim strFilePath As String       '错误日志路径
    Dim dTimeNow As Date            '当前系统时间
    Dim strDrive As String          '磁盘名称
    Dim DrDrive As Drive            '磁盘驱动器
    
    On Error GoTo errhand
    
'    strFolder = App.Path & "\LisErrLog"
'    strFilePath = App.Path & "\LisErrLog\LisErrLog.log"
'    dTimeNow = Now 'ComCurrDate
'
'
'    '检查目录是否存在，不存在则创建目录
'    If Not objFSO.FolderExists(strFolder) Then
'        objFSO.CreateFolder (strFolder)
'    End If
'
'    '检查文件是否存在，不存在则创建
'     If objFSO.FileExists(strFilePath) = False Then
'        '删除创建超过30天的备份文件
'        Set objFolder = objFSO.GetFolder(strFolder)
'        Set objFiles = objFolder.Files
'        For Each objFile In objFiles
'            strFilename = objFile.Name
'            If IsDate(Format(Mid(strFilename, 4, 4) & "-" & Mid(strFilename, 8, 2) & "-" & Mid(strFilename, 10, 2), "yyyy-mm-dd")) Then
'                dFileCreatDate = CDate(Format(Mid(strFilename, 4, 4) & "-" & Mid(strFilename, 8, 2) & "-" & Mid(strFilename, 10, 2), "yyyy-mm-dd"))
'                If DateDiff("d", dFileCreatDate, dTimeNow) > 30 Then
'                    objFSO.DeleteFile (objFile.Path)
'                End If
'            End If
'        Next
'
'        '创建日志文件之前先检查是否有足够的磁盘空间
'        strDrive = objFSO.GetDriveName(objFSO.GetAbsolutePathName(strFolder)) '获取日志文件所在的磁盘名称
'        Set DrDrive = objFSO.GetDrive(strDrive) '获取磁盘对象
'        If DrDrive.IsReady Then '检查该磁盘是否可用
'            If DrDrive.FreeSpace < 52428800 * 2 Then    '如果磁盘剩余空间小于100M，则禁止创建日志文件
'                MsgBox Mid(strDrive, 1, 1) & "盘空间不足，无法保存日志", vbInformation, "提示"
'                Exit Function
'            End If
'        End If
'        objFSO.CreateTextFile (strFilePath) '创建日志文件
'     Else
'        '如果文件已经存在则检查文件大小，文件太大影响读取效率，这里限定为50M，大于50M则生成备份文件
'        Set objFile = objFSO.GetFile(strFilePath)
'        If objFile.Size > 52428800 Then '判断文件是否大于50M
'            Name strFilePath As strFolder & "\LIS" & Format(dTimeNow, "yyyymmddhhmm") & ".bak"   '修改文件名
'            '当文件被重命名之后，需要重新检查并创建文件
'            Call writeErrLog(strObject, strModule, strErrTxt)
'        End If
'    End If
'
'    '写入错误日志
'    Set objTxt = objFSO.OpenTextFile(strFilePath, ForAppending, False)
'    objTxt.WriteLine "====================" & strObject & ":" & strModule & " " & dTimeNow & "==================" '标记是哪个部件那个模块出的错
'    objTxt.WriteLine strErrTxt
'    objTxt.Close
'    Set objTxt = Nothing
    
    If blnMsg Then MsgBox "抱歉,您正在使用的功能出现异常,请及时联系软件提供商", vbInformation, "提示"
    
    '调用公共方法记录日志
    Call zl9ComLib.LogWrite("错误日志", strModule, Mid(strErrTxt, 4, InStr(strErrTxt, ")") - 4), strErrTxt)
    
    writeErrLog = True
    Exit Function
errhand:
    writeErrLog = False
    MsgBox "写入错误日志出错,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, vbInformation, "提示"
    Err.Clear
End Function


Public Function ComSetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
    '设置参数
    ComSetPara = zlDatabase.SetPara(varPara, strValue, lngSys, lngModual, blnSetup)
End Function

Public Function ComGetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional intType As Integer) As String
    '取参数
    ComGetPara = zlDatabase.GetPara(varPara, lngSys, lngModual, strDefault, arrControl, blnSetup, intType)
End Function
