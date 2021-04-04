VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
'模块名称：clsPDF
'模块说明：输出为指定的PDF文件
'模块内容：
'
'######################################################################################################################
'获取指定系统目录
Private Declare Function GetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Private Const CSIDL_ADMINTOOLS As Long = &H30                '（用户）\开始菜单\程序\系统管理工具
Private Const CSIDL_ALTSTARTUP As Long = &H1D                '未本地化的启动
Private Const CSIDL_APPDATA As Long = &H1A                   '（用户）\应用程序的数据
Private Const CSIDL_BITBUCKET As Long = &HA                  '（桌面）\回收站
Private Const CSIDL_CONTROLS As Long = &H3                   '我的电脑\控制面板
Private Const CSIDL_COOKIES As Long = &H21
Private Const CSIDL_DESKTOP As Long = &H0                    '桌面
Private Const CSIDL_DESKTOPDIRECTORY As Long = &H10          '（用户）\桌面
Private Const CSIDL_FAVORITES As Long = &H6                  '（用户）\个性化设置
Private Const CSIDL_FONTS As Long = &H14                     'windows\字体
Private Const CSIDL_HISTORY As Long = &H22
Private Const CSIDL_INTERNET As Long = &H1                   'IE（桌面上的图标
Private Const CSIDL_INTERNET_CACHE As Long = &H20            '因特网缓存文件夹
Private Const CSIDL_LOCAL_APPDATA     As Long = &H1C            '（用户）\本地设置\应用程序数据
Private Const CSIDL_DRIVES As Long = &H11                    '我的电脑
Private Const CSIDL_MYPICTURES As Long = &H27                'C:\Program Files\My Pictures
Private Const CSIDL_NETHOOD As Long = &H13                   '（用户）\网上邻居中的元素
Private Const CSIDL_NETWORK As Long = &H12                   '网上邻居
Private Const CSIDL_PRINTERS As Long = &H4                   '我的电脑\打印机
Private Const CSIDL_PRINTHOOD As Long = &H1B                 '（用户）\打印机连接
Private Const CSIDL_PERSONAL As Long = &H5                   '我的文档
Private Const CSIDL_PROGRAM_FILES As Long = &H26             'C:\Program Files
Private Const CSIDL_PROGRAM_FILESX86 As Long = &H2A          'x86 apps (Alpha)的程序文件目录
Private Const CSIDL_PROGRAMS As Long = &H2                   '开始菜单\程序
Private Const CSIDL_PROGRAM_FILES_COMMON As Long = &H2B      'Program Files\Common
Private Const CSIDL_PROGRAM_FILES_COMMONX86 As Long = &H2C 'RISC上的x86 \Program Files\Common
Private Const CSIDL_RECENT As Long = &H8                     '(用户)\最近记录目录
Private Const CSIDL_SENDTO As Long = &H9                     '(用户)\发送到目录
Private Const CSIDL_STARTMENU As Long = &HB                  '(用户)\开始菜单
Private Const CSIDL_STARTUP As Long = &H7                    '开始菜单\程序\启动
Private Const CSIDL_SYSTEM As Long = &H25                    'system文件夹
Private Const CSIDL_SYSTEMX86 As Long = &H29                 'x86 apps (Alpha)的system文件夹
Private Const CSIDL_TEMPLATES As Long = &H15
Private Const CSIDL_PROFILE As Long = &H28                   '用户概貌文件夹
Private Const CSIDL_WINDOWS As Long = &H24                   'Windows目录或SYSROOT()
Private Const CSIDL_COMMON_ADMINTOOLS As Long = &H2F         '（所有用户）\开始菜单\程序\系统管理工具
Private Const CSIDL_COMMON_ALTSTARTUP As Long = &H1E         '未本地化的通用启动
Private Const CSIDL_COMMON_APPDATA As Long = &H23            '（所有用户）\应用程序数据
Private Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19 '（所有用户）\桌面
Private Const CSIDL_COMMON_DOCUMENTS As Long = &H2E          '（所有用户）\文档
Private Const CSIDL_COMMON_FAVORITES As Long = &H1F          '(所有用户)\设置
Private Const CSIDL_COMMON_PROGRAMS As Long = &H17           '(所有用户)\程序
Private Const CSIDL_COMMON_STARTMENU As Long = &H16          '(所有用户)\开始菜单
Private Const CSIDL_COMMON_STARTUP As Long = &H18            '(所有用户)\启动
Private Const CSIDL_COMMON_TEMPLATES As Long = &H2D          '(所有用户)\临时
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'进程获取
Private Type MODULEENTRY32
    dwSize                                      As Long
    th32ModuleID                                As Long
    th32ProcessID                               As Long
    GlblcntUsage                                As Long
    ProccntUsage                                As Long
    modBaseAddr                                 As Byte
    modBaseSize                                 As Long
    hModule                                     As Long
    szModule                                    As String * 256
    szExePath                                   As String * 1024
End Type

Private Type PROCESSENTRY32
      lSize                                     As Long
      lUsage                                    As Long
      lProcessId                                As Long
      lDefaultHeapId                            As Long
      lModuleId                                 As Long
      lThreads                                  As Long
      lParentProcessId                          As Long
      lPriClassBase                             As Long
      lFlags                                    As Long
      sExeFile                                  As String * 1024
End Type

Private Const TH32CS_SNAPPROCESS                As Long = &H2
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Private mblnReset       As Boolean
Private marrReset()     As Byte
Private mstrError       As String
Private mblnAllow       As Boolean
Private mblnFoxit       As Boolean
Private mblnFoxit7      As Boolean
Private mstrFileName    As String
Private mobjFso         As New FileSystemObject
Private mstrFoxitCach   As String
Private mstrFoxitCachCur    As String
Private mblnInit        As Boolean
Private marrFoxitSetEx(4)  As String

Public Property Get PrinterDeviceName() As String
    If mblnAllow Then
        If mblnFoxit Then
            PrinterDeviceName = "Foxit Reader PDF Printer"
        Else
            PrinterDeviceName = "TinyPDF"
        End If
    End If
End Property
'******************************************************************************************************************
'功能：初始化，并返回是否可以正常输出PDF
'返回：返回True表示能正常输出PDF文件，False表示不能正常输出PDF文件
'******************************************************************************************************************
Public Function PDFInitialize() As Boolean
    Dim strPDFFile      As String
    Dim strPath         As String * 255
    Dim lngPIDsplwow64  As Long
    Dim objFile         As File
    On Error GoTo Errhand
    mblnAllow = False
    mblnInit = True
    '检测虚拟打印机
    '------------------------------------------------------------------------------------------------------------------
    If CheckPDFDevice = False Then mblnInit = False: Exit Function
    '打印设置
    '------------------------------------------------------------------------------------------------------------------
    Call GetTempPath(255, strPath)
    strPDFFile = Trim(Left(strPath, InStr(strPath, Chr(0)) - 1)) & "Demo.pdf"
    If Dir(strPDFFile) <> "" Then
        Kill strPDFFile
        DoEvents
    End If
    
    If PDFFile(strPDFFile, False, False, False, "", "") = False Then mblnInit = False: Exit Function
    '模拟输出
    '------------------------------------------------------------------------------------------------------------------
    If OutputDemo() = False Then mblnInit = False: Exit Function
    '当没有SPLWOW64.EXE进程，打印之后就可能存在，此时该进程配置文件会从缺省文件配置复制
    If gobjComLib.OS.Is64bit And mblnFoxit7 And mblnFoxit And marrFoxitSetEx(4) = "" Then
        '64位某些情况下使用该进程的配置文件
        lngPIDsplwow64 = GetProcessID("SPLWOW64.EXE")
        If lngPIDsplwow64 <> 0 And mobjFso.FolderExists(mobjFso.GetParentFolderName(marrFoxitSetEx(0))) Then
            For Each objFile In mobjFso.GetFolder(mobjFso.GetParentFolderName(marrFoxitSetEx(0))).Files
                If UCase(objFile.Name) Like "*_" & lngPIDsplwow64 & "__FOXITTEMP.XML" Then
                    marrFoxitSetEx(4) = objFile.Path
                    Exit For
                End If
            Next
        End If
    End If
    
    PDFInitialize = PDFFileSuccess
    mblnAllow = True
    mblnInit = False
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    mblnInit = False
    mstrError = Err.Description
End Function
'******************************************************************************************************************
'功能：设置要生成的PDF文件及相关属性，对于Foxit虚拟打印，必须指定文件名，并且只支持指定文件名，不支持其他参数。
'参数：strPDFFile=输出文件名，包含文件路径和文件扩展名
'                 文件路径必须存在，且会自动覆盖同名文件
'                 如果未指定，则弹出文件保存对话框
'      blnCopyable=输出的PDF文件是否允许复制内容
'      blnEditable=输出的PDF文件是否允许编辑内容
'      blnPrintable=输出的PDF文件是否允许打印输出
'      strPassword=是否要求输入密码
'      strAttachs=要加到PDF中的附件文件名(包含路径),多个以"|"分隔
'返回：
'注意：该函数需要在Printer的任何打印设置之前调用(包括API方式设置)
'******************************************************************************************************************
Public Function PDFFile(ByVal strPDFFile As String, _
                        Optional ByVal blnCopyable As Boolean = False, _
                        Optional ByVal blnEditable As Boolean = False, _
                        Optional ByVal blnPrintable As Boolean = False, _
                        Optional ByVal strPassWord As String = "", _
                        Optional ByVal strAttachs As String = "") As Boolean

    Dim strDeviceName   As String, intLoop As Long
    On Error GoTo Errhand
    If mblnAllow = False And Not mblnInit Then Exit Function
    '检测虚拟打印机，因为可能会切换
    If mblnFoxit Then
        strDeviceName = "Foxit Reader PDF Printer"
    Else
        strDeviceName = "TinyPDF"
    End If
    If Printer.DeviceName <> strDeviceName Then
        For intLoop = 0 To Printers.count - 1
            If Printers(intLoop).DeviceName = strDeviceName Then Set Printer = Printers(intLoop):  Exit For
        Next
        If intLoop >= Printers.count Then
            mstrError = "没有检测到安装了" & strDeviceName & "虚拟打印机，不能输出PDF！"
            Exit Function
        End If
    End If
    If strPDFFile = "" Then
        mstrError = "未指定PDF文件名称，不能输出PDF！"
        Exit Function
    End If
    mstrFileName = strPDFFile
    If mblnFoxit Then
        PDFFile = SetFoxitPrinter(mstrFoxitCachCur)
    Else
        PDFFile = ModifyRegist(strPDFFile, blnCopyable, blnEditable, blnPrintable, strPassWord, strAttachs)
    End If
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    mstrError = Err.Description
End Function

'******************************************************************************************************************
'功能:PDF输出文件是否成功。foxit必须调用该方法，才能生成文件。
'返回：True,文件生成成功，False-文件生成失败
'说明：该函数在打印输出完成后调用
'******************************************************************************************************************
Public Function PDFFileSuccess() As Boolean
    Dim strFileName As String
    Dim objFile     As File
    If mblnAllow = False And Not mblnInit Then Exit Function
    If mblnFoxit Then
        For Each objFile In mobjFso.GetFolder(mstrFoxitCachCur).Files
            strFileName = objFile.Path
            Exit For
        Next
        If strFileName <> "" Then
            Call mobjFso.CopyFile(strFileName, mstrFileName, True)
            PDFFileSuccess = mobjFso.FileExists(mstrFileName)
            Call ClearFolder(mstrFoxitCachCur)
        End If
    Else
        PDFFileSuccess = mobjFso.FileExists(mstrFileName)
    End If
End Function

Public Function GetLastError() As String
    GetLastError = mstrError
End Function

'******************************************************************************************************************
'功能：检测PDF虚拟打印机
'返回：True-存在虚拟打印，False-不存在虚拟打印机
'******************************************************************************************************************
Private Function CheckPDFDevice() As Boolean
    Dim intLoop         As Integer
    Dim strDeviceName   As String
    Dim lngRetrun       As Long
    Dim strPath         As String, strTemp          As String * 100
    Dim objFile         As File, objFolder          As Folder
    Dim lngProcessID    As Long, strProcessName     As String
    Dim lngPos          As Long
    Dim strShortSvrName As String, strProFile       As String
    Dim arrTmp()        As Byte
    Dim lngPIDsplwow64  As Long
    
    mblnFoxit = False
    '检测是否存在TinyPDF(32位系统，先检查Foxit Reader PDF Printer，再检查TinyPDF) Foxit Reader PDF Printer (64位系统)打印机
    strDeviceName = "Foxit Reader PDF Printer"
    If Printer.DeviceName <> strDeviceName Then
        For intLoop = 0 To Printers.count - 1
            If Printers(intLoop).DeviceName = strDeviceName Then Set Printer = Printers(intLoop):  Exit For
        Next
        If intLoop >= Printers.count Then
            If Not gobjComLib.OS.Is64bit Then
                strDeviceName = "TinyPDF"
                For intLoop = 0 To Printers.count - 1
                    If Printers(intLoop).DeviceName = strDeviceName Then Set Printer = Printers(intLoop):  Exit For
                Next
                If intLoop >= Printers.count Then
                    mstrError = "没有检测到安装了TinyPDF或Foxit Reader PDF Printer虚拟打印机，不能输出PDF！"
                    Exit Function
                End If
            Else
                mstrError = "没有检测到安装了" & strDeviceName & "虚拟打印机，不能输出PDF！"
                Exit Function
            End If
        End If
    End If
    If strDeviceName = "Foxit Reader PDF Printer" Then
        If Not mobjFso.FolderExists(mstrFoxitCachCur) Or marrFoxitSetEx(0) = "" Then
            mblnFoxit7 = False
            '获取打印机配置文件目录
            strTemp = String(50, " ")
            lngRetrun = GetSpecialFolderPath(0, strTemp, CSIDL_APPDATA, False)
            strPath = gobjComLib.zlStr.TruncZero(strTemp)
            strPath = strPath & "\Foxit Software\Foxit PDF Creator"
            If mobjFso.FolderExists(strPath) = False Then
                mstrError = "未找到Foxit Reader PDF打印机配置目录，文件将不能输出的指定的目录!" & vbCrLf & "配置目录:" & strPath
                Exit Function
            End If
            
            For intLoop = 0 To 4
                marrFoxitSetEx(intLoop) = ""
            Next
            If mobjFso.FolderExists(strPath & "\Foxit Reader PDF Printer") Then
                'Excel使用的缺省配置文件
                'Foxit 7.0
                mblnFoxit7 = True
                marrFoxitSetEx(0) = strPath & "\Foxit Reader PDF Printer\FoxitPrinterProfile.xml"
                'VB勾选设置为缺省配置时的配置文件
                If gobjComLib.OS.GetRegValue("HKEY_CURRENT_USER\Printers\DevModePerUser", "Foxit Reader PDF Printer", arrTmp) Then
                    strProFile = GetFoxitProfile(arrTmp)
                    If mobjFso.FolderExists(mobjFso.GetParentFolderName(strProFile)) Then
                        marrFoxitSetEx(0) = mobjFso.GetParentFolderName(strProFile) & "\FoxitPrinterProfile.xml"
                        marrFoxitSetEx(1) = strProFile
                    ElseIf gobjComLib.OS.GetRegValue("HKEY_CURRENT_USER\Printers\DevModes2", "Foxit Reader PDF Printer", arrTmp) Then
                        strProFile = GetFoxitProfile(arrTmp)
                        If mobjFso.FolderExists(mobjFso.GetParentFolderName(strProFile)) Then
                            marrFoxitSetEx(0) = mobjFso.GetParentFolderName(strProFile) & "\FoxitPrinterProfile.xml"
                            marrFoxitSetEx(1) = strProFile
                        End If
                    End If
                End If
                '获取splwow64.exe的路径
                If gobjComLib.OS.Is64bit Then
                    '64位某些情况下使用该进程的配置文件
                    lngPIDsplwow64 = GetProcessID("SPLWOW64.EXE")
                    If lngPIDsplwow64 <> 0 Then
                        For Each objFile In mobjFso.GetFolder(strPath & "\Foxit Reader PDF Printer").Files
                            If UCase(objFile.Name) Like "*_" & lngPIDsplwow64 & "__FOXITTEMP.XML" Then
                                marrFoxitSetEx(4) = objFile.Path
                                Exit For
                            End If
                        Next
                    End If
                End If
                'VB程序未勾选使用的配置文件
                If gobjComLib.OS.GetRegValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Printers\Foxit Reader PDF Printer\DsSpooler", "shortServerName", strShortSvrName) Then
                    If gobjComLib.OS.GetRegValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Printers\Foxit Reader PDF Printer", "Default DevMode", arrTmp) Then
                        strProFile = GetFoxitProfile(arrTmp)
                        If mobjFso.FolderExists(mobjFso.GetParentFolderName(strProFile)) Then
                            '当没有配置文件时的缺省配置文件
                            marrFoxitSetEx(2) = mobjFso.GetParentFolderName(strProFile) & "\FoxitPrinterProfile.xml"
                            marrFoxitSetEx(3) = strProFile
                        End If
                    End If
                End If
            Else
                marrFoxitSetEx(0) = strPath & "\FoxitReaderPrinterProfile.xml"
            End If
            
            '获取路径
            If gobjComLib.OS.IsDesinMode Then
                mstrFoxitCach = "C:\APPSOFT\FoxitPrinter"
            Else
                mstrFoxitCach = Mid(App.Path & "\", 1, InStr(4, App.Path & "\", "\")) & "FoxitPrinter"
            End If
            If Not mobjFso.FolderExists(mobjFso.GetParentFolderName(mstrFoxitCach)) Then
                Call mobjFso.CreateFolder(mobjFso.GetParentFolderName(mstrFoxitCach))
            End If
            If Not mobjFso.FolderExists(mstrFoxitCach) Then
                Call mobjFso.CreateFolder(mstrFoxitCach)
            End If
        
            '判断进程是否存在,清理垃圾文件夹
            For Each objFolder In mobjFso.GetFolder(mstrFoxitCach).SubFolders
                lngPos = InStr(objFolder.Name, "_")
                If lngPos > 0 Then
                    lngProcessID = Val(Mid(objFolder.Name, 1, lngPos - 1))
                    strProcessName = Mid(objFolder.Name, lngPos + 1)
                    If Not FindProcess(strProcessName, lngProcessID) Then
                        Call ClearFolder(objFolder.Path, True)
                    End If
                Else
                    Call ClearFolder(objFolder.Path, True)
                End If
            Next
            mstrFoxitCachCur = mstrFoxitCach & "\" & GetCurProcessInfo
            '删除同名文件，否则无法创建文件夹
            If mobjFso.FileExists(mstrFoxitCachCur) Then
                Call FileNormal(mstrFoxitCachCur)
                Call mobjFso.DeleteFile(mstrFoxitCachCur, True)
            End If
            If mobjFso.FolderExists(mstrFoxitCachCur) Then
                Call ClearFolder(mstrFoxitCachCur)
                If mobjFso.GetFolder(mstrFoxitCachCur).Files.count <> 0 Then
                    mstrError = "无法清空Foxit Reader PDF打印机缓存目录，文件将不能输出的指定的目录!" & vbCrLf & "缓存目录:" & mstrFoxitCachCur
                    Exit Function
                End If
            Else
                Call mobjFso.CreateFolder(mstrFoxitCachCur)
            End If
        End If
        Call SetFoxitPrinter(mstrFoxitCachCur)
        mblnFoxit = True
    End If
    CheckPDFDevice = True
End Function

'******************************************************************************************************************
'功能：指定TinyPDF打印机输出参数
'参数：strPDFFile=输出文件名，包含文件路径和文件扩展名
'                 文件路径必须存在，且会自动覆盖同名文件
'                 如果未指定，则弹出文件保存对话框
'      blnCopyable=输出的PDF文件是否允许复制内容
'      blnEditable=输出的PDF文件是否允许编辑内容
'      blnPrintable=输出的PDF文件是否允许打印输出
'      strPassword=是否要求输入密码
'      strAttachs=要加到PDF中的附件文件名(包含路径),多个以"|"分隔
'注意：该函数需要在Printer的任何打印设置之前调用(包括API方式设置)
'******************************************************************************************************************
Private Function ModifyRegist(ByVal strPDFFile As String, Optional ByVal blnCopyable As Boolean, Optional ByVal blnEditable As Boolean, Optional ByVal blnPrintable As Boolean, Optional ByVal strPassWord As String, Optional ByVal strAttachs As String) As Boolean
    Dim arrData() As Byte
    Dim intSect As Integer, intAdr As Integer
    Dim intTag As Integer, strFile As String
    Dim i As Integer, j As Integer
    Dim strWord As String
    Dim strRegister(92) As String
    Dim aryRegister As Variant
    Dim intLoop As Integer
    
    '读取设置
    Call gobjComLib.OS.GetRegValue("HKEY_CURRENT_USER\Printers\DevModePerUser", "TinyPDF", arrData)
    
    On Error Resume Next
    Err = 0
    i = UBound(arrData)
    If Err <> 0 Then i = -1
    On Error GoTo Errhand
    
    If i = -1 Then
        '空注册表

        strRegister(0) = "84,0,105,0,110,0,121,0,80,0,68,0,70,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(1) = "0,0,0,0,0,0,0,0,0,0,0,0,0,1,4,0,4,220,0,236,16,19,78,1,0,1,0,9,0,0,0,0,0,100,0,1,0,15,0,88,2,2,0,1,0,0,0,3,0,0"
        strRegister(2) = "0,65,117,116,111,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(3) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(4) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,63,0,0,0,1,0,0,0,3,0,0,0,44,1,0,0,194,1,0,0,2,80,0,0,3,0,0"
        strRegister(5) = "0,44,1,0,0,194,1,0,0,2,80,0,0,3,0,0,0,176,4,0,0,8,7,0,0,2,0,0,0,0,0,0,0,1,0,0,0,1,0,0,0,100,0,0,0,2,0,0,0,6"
        strRegister(6) = "0,0,0,1,3,0,0,26,1,0,0,44,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(7) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1"
        strRegister(8) = "0,0,0,0,0,0,0,70,1,0,0,72,1,0,0,0,0,0,0,74,1,0,0,76,1,0,0,78,1,0,0,80,1,0,0,82,1,0,0,84,1,0,0,86,1,0,0,88,1,0"
        strRegister(9) = "0,90,1,0,0,0,0,0,0,0,0,65,0,114,0,105,0,97,0,108,0,0,0,65,0,114,0,105,0,97,0,108,0,32,0,78,0,97,0,114,0,114,0,111,0,119,0,0,0,65"
        strRegister(10) = "0,114,0,105,0,97,0,108,0,32,0,85,0,110,0,105,0,99,0,111,0,100,0,101,0,32,0,77,0,83,0,0,0,67,0,101,0,110,0,116,0,117,0,114,0,121,0,32,0,71"
        strRegister(11) = "0,111,0,116,0,104,0,105,0,99,0,0,0,67,0,111,0,117,0,114,0,105,0,101,0,114,0,32,0,78,0,101,0,119,0,0,0,71,0,101,0,111,0,114,0,103,0,105,0,97"
        strRegister(12) = "0,0,0,73,0,109,0,112,0,97,0,99,0,116,0,0,0,76,0,117,0,99,0,105,0,100,0,97,0,32,0,67,0,111,0,110,0,115,0,111,0,108,0,101,0,0,0,84,0,97"
        strRegister(13) = "0,104,0,111,0,109,0,97,0,0,0,84,0,105,0,109,0,101,0,115,0,32,0,78,0,101,0,119,0,32,0,82,0,111,0,109,0,97,0,110,0,0,0,84,0,114,0,101,0,98"
        strRegister(14) = "0,117,0,99,0,104,0,101,0,116,0,32,0,77,0,83,0,0,0,86,0,101,0,114,0,100,0,97,0,110,0,97,0,0,0,0,0,115,82,71,66,32,73,69,67,54,49,57,54,54"
        strRegister(15) = "45,50,46,49,0,85,46,83,46,32,87,101,98,32,67,111,97,116,101,100,32,40,83,87,79,80,41,32,118,50,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(16) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(17) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(18) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(19) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(20) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(21) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(22) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(23) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(24) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(25) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(26) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(27) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(28) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(29) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(30) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(31) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(32) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(33) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(34) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(35) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(36) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(37) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(38) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(39) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(40) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(41) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(42) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(43) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(44) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(45) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(46) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(47) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(48) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(49) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(50) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(51) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(52) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(53) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(54) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(55) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(56) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(57) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(58) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(59) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(60) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(61) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(62) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(63) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(64) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(65) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(66) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(67) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(68) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(69) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(70) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(71) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(72) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(73) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(74) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(75) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(76) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(77) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(78) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(79) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(80) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(81) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(82) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(83) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(84) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(85) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(86) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(87) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(88) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(89) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(90) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
        strRegister(91) = "0"
        
        For i = 0 To 91
            aryRegister = Split(strRegister(i), ",")
            For j = 0 To UBound(aryRegister)
                ReDim Preserve arrData(intLoop)
                arrData(intLoop) = Val(aryRegister(j))
                intLoop = intLoop + 1
            Next
        Next
    End If
    
    If Not mblnReset Then
        marrReset = arrData
        mblnReset = True
    End If

    '更改设置
    arrData(Val("&H00E0")) = &H0 '页边距
    arrData(Val("&H00E1")) = &H0 '页边距
    arrData(Val("&H00E4")) = &H0 '不自动打开
    arrData(Val("&H011C")) = &H1 '嵌入所有字体
    arrData(Val("&H0130")) = &H0 'RGB颜色(sRGB不正常)
    
    If strPassWord <> "" Then
        arrData(Val("&H013C")) = &H1 '有用户密码
        For i = 1 To Len(strPassWord)
            arrData(Val("&H0140") + i - 1) = Asc(Mid(strPassWord, i, 1))
        Next
        arrData(Val("&H0140") + i - 1) = &H0
    Else
        arrData(Val("&H013C")) = &H0 '无用户密码
        arrData(Val("&H0140")) = &H0
    End If
    
    arrData(Val("&H0164")) = &H1 '进行安全控制
    arrData(Val("&H0168")) = &H0  '安全控制密码为空
    If blnPrintable Then
        arrData(Val("&H0189")) = &H2  '允许高分辨率打印
    Else
        arrData(Val("&H0189")) = &H0  '不允许打印
    End If
    If blnEditable Then
        arrData(Val("&H018A")) = &H4  '除提取页面之外的任何内容
    Else
        arrData(Val("&H018A")) = &H0  '不允许更改
    End If
    If blnCopyable Then
        arrData(Val("&H018C")) = &H1  '允许复制内容
    Else
        arrData(Val("&H018C")) = &H0  '不允许复制
    End If
    arrData(Val("&H0190")) = &H1  '不允许复制时，允许屏幕阅读器设备视觉受损地访问文本
    If strPDFFile <> "" Then
        arrData(Val("&H0194")) = &H2  '指定文件输出(包含路径)
    Else
        arrData(Val("&H0194")) = &H0  '提示输出
    End If
    arrData(Val("&H01A0")) = &H1  '直接覆盖文件
    
    '数据段：输出文件、附件文件
    arrData(Val("&H01C8")) = &H0
    arrData(Val("&H01C8") + 1) = &H0
    intAdr = Val("&H01CA")
    intSect = 1 '数据段序号
    intTag = 1 '1-数据内容,2-单项结束
    Do While intAdr <= 4552
        If intSect = 1 Or intSect = 2 Then '嵌入/不嵌入字体段
            If arrData(intAdr) = 0 And arrData(intAdr + 1) = 0 Then
                If intTag = 1 Then
                    intTag = 2
                ElseIf intTag = 2 Then
                    intTag = 1
                    intSect = intSect + 1
                End If
            Else
                intTag = 1
            End If
            intAdr = intAdr + 2
        ElseIf intSect = 3 Then '中间间隔段
            If arrData(intAdr) = 0 Then
                intAdr = intAdr + 1
            Else
                intSect = intSect + 1
            End If
        ElseIf intSect = 4 Or intSect = 5 Then 'RGB/CMYK配置文件段
            If arrData(intAdr) = 0 Then
                intSect = intSect + 1
            End If
            intAdr = intAdr + 1
        ElseIf intSect = 6 Then '输出目录段
            strWord = Hex(intAdr - Val("&H01C8"))
            strWord = String(4 - Len(strWord), "0") & strWord
            arrData(Val("&H0198")) = Val("&H" & Right(strWord, 2)) '低位字节
            arrData(Val("&H0198") + 1) = Val("&H" & Left(strWord, 2)) '高位字节
            
            arrData(intAdr) = 0
            arrData(intAdr + 1) = 0
            intAdr = intAdr + 2
            intSect = intSect + 1
        ElseIf intSect = 7 Then '输出文件段
            strWord = Hex(intAdr - Val("&H01C8"))
            strWord = String(4 - Len(strWord), "0") & strWord
            arrData(Val("&H019C")) = Val("&H" & Right(strWord, 2)) '低位字节
            arrData(Val("&H019C") + 1) = Val("&H" & Left(strWord, 2)) '高位字节
            
            If strPDFFile = "" Then
                arrData(intAdr) = 0
                arrData(intAdr + 1) = 0
                intAdr = intAdr + 2
            Else
                For i = 1 To Len(strPDFFile)
                    strWord = Hex(AscW(Mid(strPDFFile, i, 1)))
                    If Len(strWord) = 2 Then
                        strWord = "00" & strWord
                    End If
                    
                    arrData(intAdr + i * 2 - 2) = Val("&H" & Right(strWord, 2)) '低位Unicode
                    arrData(intAdr + i * 2 - 1) = Val("&H" & Left(strWord, 2)) '高位Unicode
                Next
                intAdr = intAdr + Len(strPDFFile) * 2
                
                arrData(intAdr) = 0
                arrData(intAdr + 1) = 0
                intAdr = intAdr + 2
            End If
            intSect = intSect + 1
        ElseIf intSect = 8 Then '中间间隔段
            strWord = Hex(intAdr - Val("&H01C8"))
            strWord = String(4 - Len(strWord), "0") & strWord
            arrData(Val("&H01A4")) = Val("&H" & Right(strWord, 2)) '低位字节
            arrData(Val("&H01A4") + 1) = Val("&H" & Left(strWord, 2)) '高位字节
            
            For i = 1 To 16
                arrData(intAdr + i - 1) = 0
            Next
            intAdr = intAdr + 16
            intSect = intSect + 1
        ElseIf intSect = 9 Then '附件文件
            '目前发现设置加载附件会导致生成的PDF打开出错
            If strAttachs = "" Then
                arrData(intAdr) = 0
                arrData(intAdr + 1) = 0
                intAdr = intAdr + 2
            Else
                For i = 0 To UBound(Split(strAttachs, "|"))
                    strFile = Split(strAttachs, "|")(i)
                    For j = 1 To Len(strFile)
                        strWord = Hex(AscW(Mid(strFile, j, 1)))
                        If Len(strWord) = 2 Then
                            strWord = "00" & strWord
                        End If
                        
                        arrData(intAdr + j * 2 - 2) = Val("&H" & Right(strWord, 2)) '低位Unicode
                        arrData(intAdr + j * 2 - 1) = Val("&H" & Left(strWord, 2)) '高位Unicode
                    Next
                    intAdr = intAdr + Len(strFile) * 2
                    
                    arrData(intAdr) = 0
                    arrData(intAdr + 1) = 0
                    intAdr = intAdr + 2
                Next
            End If
            '结束退出
            Exit Do
        End If
    Loop
    
    For i = Val("&H01A8") To Val("&H01C4") Step 4
        strWord = Hex(arrData(Val("&H01A4")) + arrData(Val("&H01A4") + 1) * 256 + (i - Val("&H01A8")) / 2 + 2)
        strWord = String(4 - Len(strWord), "0") & strWord
        
        arrData(i) = Val("&H" & Right(strWord, 2)) '低位字节
        arrData(i + 1) = Val("&H" & Left(strWord, 2)) '高位字节
        arrData(i + 2) = 0
        arrData(i + 3) = 0
    Next
    
    '保存设置
    Call gobjComLib.OS.SetRegValue("HKEY_CURRENT_USER\Printers\DevModePerUser", "TinyPDF", arrData)
    Call gobjComLib.OS.SetRegValue("HKEY_CURRENT_USER\Printers\DevModes2", "TinyPDF", arrData)
    ModifyRegist = True
    Exit Function
Errhand:
    mstrError = Err.Description
End Function

Private Function SetFoxitPrinter(ByVal strFilePath As String) As Boolean
'功能：设置Foxit虚拟打印机目录
    Dim objStream As TextStream
    Dim i           As Long
    Dim blnDo       As Boolean
    Dim blnOK       As Boolean
    On Error GoTo errH
    blnOK = False
    For i = 0 To 4
        blnDo = False
        If marrFoxitSetEx(i) <> "" Then
            If mobjFso.FileExists(marrFoxitSetEx(i)) Then
                If SetFoxitFolder(strFilePath, marrFoxitSetEx(i)) Then
                    blnDo = True
                End If
            End If
            If Not blnDo Then
                '创建一个文件
                Set objStream = mobjFso.CreateTextFile(marrFoxitSetEx(i), True)
                If Not mblnFoxit7 Then
                    objStream.WriteLine "<FXCreatorData><General Folder="""" Overwrite=""1"" UseDefFileName=""0"" OpenFile=""0"" ImageCompress=""0"" IgonareBK=""0"" PDFA1B=""0"" PDFVersion=""17"" Gray=""0"" BlackAndWhite=""0"" DPI=""600""/>"
                    objStream.WriteLine "</FXCreatorData>"
                Else
                    objStream.WriteLine "<FXCreatorData><General DefaultFolder="""" Overwrite=""1"" UseDefFileName=""0"" OpenFile=""0"" TemplateName=""Standard"" IgonareBK=""0"" PDFVersion=""17"" ColorSpace=""2"" DPI=""600"" DeleteLogFile=""0""/>"
                    objStream.WriteLine "<Layout UOM=""0"" PaperSize=""9"" PaperWidth=""2100"" PaperLength=""2970"" Orientation=""1"" FormName=""A4""/>"
                    objStream.WriteLine "<DocumentInfo AddDocInfo=""0"" DocTitle="""" DocSubject="""" DocAuthor="""" DocKeyWords="""" DocCreator=""""/>"
                    objStream.WriteLine "</FXCreatorData>"
                End If
                objStream.Close
                Set objStream = Nothing
                If Not SetFoxitFolder(strFilePath, marrFoxitSetEx(i)) Then
                    mstrError = "打印机配置文件格式不是有效的xml文件，请检查." & vbCrLf & "文件路径:" & marrFoxitSetEx(i)
                    marrFoxitSetEx(i) = ""
                Else
                    blnDo = True
                End If
            End If
            If blnDo Then
                blnOK = True
                mstrError = ""
            ElseIf blnOK Then
                mstrError = ""
            End If
        End If
    Next
    SetFoxitPrinter = blnOK
    Exit Function
errH:
    mstrError = "(" & Err.Number & ")" & Err.Description
    Err.Clear
End Function

Private Function SetFoxitFolder(ByVal strFilePath As String, ByVal strFoxitSet As String) As Boolean
    Dim objXML As Object
    Dim objNode As Object
    
    On Error GoTo errH
    Set objXML = CreateObject("MSXML2.DOMDocument")
    If objXML.Load(strFoxitSet) = True Then
        Set objNode = objXML.selectSingleNode("FXCreatorData").selectSingleNode("General")
        If mblnFoxit7 Then
            If objNode.Attributes.getNamedItem("DefaultFolder").Text <> strFilePath Then
                objNode.Attributes.getNamedItem("DefaultFolder").Text = strFilePath
            End If
        Else
            If objNode.Attributes.getNamedItem("Folder").Text <> strFilePath Then
                objNode.Attributes.getNamedItem("Folder").Text = strFilePath
            End If
        End If
        If objNode.Attributes.getNamedItem("Overwrite").Text <> "1" Then
            objNode.Attributes.getNamedItem("Overwrite").Text = "1"
        End If
        If objNode.Attributes.getNamedItem("UseDefFileName").Text <> "1" Then
            objNode.Attributes.getNamedItem("UseDefFileName").Text = "1"
        End If
        If objNode.Attributes.getNamedItem("OpenFile").Text <> "0" Then
            objNode.Attributes.getNamedItem("OpenFile").Text = "0"
        End If
        Call objXML.Save(strFoxitSet)
        SetFoxitFolder = True
    End If
    Set objXML = Nothing
    Exit Function
errH:
    Err.Clear
End Function

'******************************************************************************************************************
'功能：输出测试Demo
'******************************************************************************************************************
Private Function OutputDemo() As Boolean
    On Error GoTo Errhand
    Call OutPut(Printer)
    OutputDemo = True
    Exit Function
Errhand:
    mstrError = Err.Description
    Printer.EndDoc
End Function

Private Sub OutPut(objOut As Object)
    '------
    objOut.Font.Name = "黑体"
    objOut.Font.Size = 18
    objOut.ForeColor = vbRed
    objOut.CurrentY = 300
    objOut.CurrentX = (objOut.ScaleWidth - objOut.TextWidth("PDF文件生成测试示例")) / 2
    objOut.Print "PDF文件生成测试示例"
    
    '------
    objOut.DrawWidth = 2 '线宽在打印机上面区别不是很明显
    objOut.Line (100, 800)-(objOut.ScaleWidth - 100, 800), vbBlue
    
    '------
    objOut.Font.Name = "宋体"
    objOut.Font.Size = 12
    objOut.ForeColor = vbBlack
    objOut.CurrentX = 300
    objOut.CurrentY = 1000 + 100
    objOut.Print "恭喜！"
    
    objOut.CurrentX = 300
    objOut.Print "如果您可以读取这个信息，则说明在本机上可以生成PDF文件。"
    objOut.EndDoc
End Sub

'******************************************************************************************************************
'功能：重置TinyPDF打印机输出参数设置
'说明：该函数在打印输出完成后调用
'******************************************************************************************************************
Private Sub ResetPDF()
    If mblnReset Then
        Call gobjComLib.OS.SetRegValue("HKEY_CURRENT_USER\Printers\DevModePerUser", "TinyPDF", marrReset)
        Call gobjComLib.OS.SetRegValue("HKEY_CURRENT_USER\Printers\DevModes2", "TinyPDF", marrReset)
        Erase marrReset
        mblnReset = False
    Else
        Call gobjComLib.OS.DeleteRegValue("HKEY_CURRENT_USER\Printers\DevModePerUser", "TinyPDF")
        Call gobjComLib.OS.DeleteRegValue("HKEY_CURRENT_USER\Printers\DevModes2", "TinyPDF")
    End If
End Sub

'******************************************************************************************************************
'功能： 清除文件的只读标志
'说明:
'******************************************************************************************************************
Private Function FileNormal(ByVal strSource As String) As Boolean

    On Error Resume Next
    If mobjFso.FileExists(strSource) Then
        If FileSystem.GetAttr(strSource) <> vbNormal Then
            FileSystem.SetAttr strSource, vbNormal
        End If
    End If
    
    FileNormal = Err.Number = 0
    If Err.Number <> 0 Then Err.Clear
End Function

'******************************************************************************************************************
'功能： 判断指定名称和进程ID的进程是否存在
'说明:
'******************************************************************************************************************
Private Function FindProcess(ByVal strProcessName As String, Optional ByVal lngProcID As Long) As Boolean
    Dim uProcess As PROCESSENTRY32
    Dim lngSnapShot As Long, lngRet As Long
    Dim strFindName As String, lngPos As Long
    Dim lngPid As Long
    On Error GoTo errH
    FindProcess = False
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot <> 0 Then
        uProcess.lSize = 1060
        If (Process32First(lngSnapShot, uProcess)) Then
            Do
                lngPos = InStr(1, uProcess.sExeFile, Chr(0))
                strFindName = UCase(Left(uProcess.sExeFile, lngPos - 1))
                If strFindName = strProcessName Then
                    lngPid = uProcess.lProcessId
                    If lngProcID = lngPid Then
                        FindProcess = True
                        Exit Do
                    End If
                End If
            Loop Until (Process32Next(lngSnapShot, uProcess) < 1)
        End If
        lngRet = CloseHandle(lngSnapShot)
    End If
    Exit Function
errH:
    Err.Clear
End Function
'******************************************************************************************************************
'功能： 获取当前进程的进程名称和进程ID
'返回：进程ID_进程EXE名称
'说明:
'******************************************************************************************************************
Public Function GetCurProcessInfo() As String
    Dim lngCurProcID        As Long
    Dim uProcess            As PROCESSENTRY32
    Dim lngSnapShot         As Long, lngRet         As Long
    Dim strFindName         As String, lngPos       As Long
    Dim lngPid              As Long
    
    
    lngCurProcID = GetCurrentProcessId
    
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot <> 0 Then
        uProcess.lSize = 1060
        If (Process32First(lngSnapShot, uProcess)) Then
            Do
                lngPid = uProcess.lProcessId
                If lngCurProcID = lngPid Then
                    lngPos = InStr(1, uProcess.sExeFile, Chr(0))
                    strFindName = UCase(Left(uProcess.sExeFile, lngPos - 1))
                    GetCurProcessInfo = lngCurProcID & "_" & strFindName
                    Exit Do
                End If
            Loop Until (Process32Next(lngSnapShot, uProcess) < 1)
        End If
        lngRet = CloseHandle(lngSnapShot)
    End If
End Function

'******************************************************************************************************************
'功能： 获取当前进程的进程名称和进程ID
'返回：进程ID_进程EXE名称
'说明:
'******************************************************************************************************************
Private Function GetProcessID(ByVal strExeName As String) As Long
    Dim uProcess            As PROCESSENTRY32
    Dim lngSnapShot         As Long, lngRet         As Long
    Dim strFindName         As String, lngPos       As Long
    Dim lngPid              As Long
    
    
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot <> 0 Then
        uProcess.lSize = 1060
        If (Process32First(lngSnapShot, uProcess)) Then
            Do
                lngPos = InStr(1, uProcess.sExeFile, Chr(0))
                strFindName = UCase(Left(uProcess.sExeFile, lngPos - 1))
                If strFindName = strExeName Then
                    GetProcessID = uProcess.lProcessId
                    Exit Do
                End If
            Loop Until (Process32Next(lngSnapShot, uProcess) < 1)
        End If
        lngRet = CloseHandle(lngSnapShot)
    End If
End Function

'******************************************************************************************************************
'功能：清理指定文件夹
'参数：strFolder-清理的文件夹，blnDelFolder-是否删除该文件夹
'说明:
'******************************************************************************************************************
Private Sub ClearFolder(ByVal strFolder As String, Optional ByVal blnDelFolder As Boolean)
    Dim objFile          As File
    On Error Resume Next
    If mobjFso.FolderExists(strFolder) Then
        For Each objFile In mobjFso.GetFolder(strFolder).Files
            Call FileNormal(objFile.Path)
            Call mobjFso.DeleteFile(objFile.Path, True)
        Next
        If blnDelFolder Then Call mobjFso.DeleteFolder(strFolder, True)
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Class_Terminate()
    Call ClearFolder(mstrFoxitCachCur, True)
    Set mobjFso = Nothing
End Sub

Private Function GetFoxitProfile(ByRef arrData As Variant) As String
    Dim i       As Long
    Dim strTmp  As String
    Dim arrTmp()    As Byte
    Dim lngLoop     As Long
    On Error GoTo errH
    Err.Clear
    i = UBound(arrData)
    For i = &HDC - 1 To UBound(arrData)
        If arrData(i) <> 0 Then
            ReDim Preserve arrTmp(lngLoop)
            arrTmp(lngLoop) = Val(arrData(i))
            lngLoop = lngLoop + 1
        End If
    Next
    strTmp = StrConv(arrTmp(), vbUnicode)
    strTmp = Replace(strTmp, "\\\", "\")
    GetFoxitProfile = strTmp
    Exit Function
errH:
    Err.Clear
End Function

