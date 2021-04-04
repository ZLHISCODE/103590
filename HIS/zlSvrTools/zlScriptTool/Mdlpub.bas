Attribute VB_Name = "Mdlpub"
Option Explicit

Public Const HWND_TOPMOST = -1
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000

'---------------------------------------------------------------
'- 注册表安全属性类型...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_NT_WORKSTATION = 1
Private Const VER_NT_DOMAIN_CONTROLLER = 2
Private Const VER_NT_SERVER = 3
Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
    wServicePackMajor As Integer 'win2000 only
    wServicePackMinor As Integer 'win2000 only
    wSuiteMask As Integer 'win2000 only
    wProductType As Byte 'win2000 only
    wReserved As Byte
End Type
Private Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long

Private Const PRODUCT_UNLICENSED As Long = &HABCDABCD
Private Const PRODUCT_BUSINESS As Long = &H6
Private Const PRODUCT_BUSINESS_N As Long = &H10
Private Const PRODUCT_CLUSTER_SERVER As Long = &H12
Private Const PRODUCT_DATACENTER_SERVER As Long = &H8
Private Const PRODUCT_DATACENTER_SERVER_CORE As Long = &HC
Private Const PRODUCT_ENTERPRISE As Long = &H4
Private Const PRODUCT_ENTERPRISE_N As Long = &H1B
Private Const PRODUCT_ENTERPRISE_SERVER As Long = &HA
Private Const PRODUCT_ENTERPRISE_SERVER_CORE As Long = &HE
Private Const PRODUCT_ENTERPRISE_SERVER_IA64 As Long = &HF
Private Const PRODUCT_HOME_BASIC As Long = &H2
Private Const PRODUCT_HOME_BASIC_N As Long = &H5
Private Const PRODUCT_HOME_PREMIUM As Long = &H3
Private Const PRODUCT_HOME_PREMIUM_N As Long = &H1A
Private Const PRODUCT_HOME_SERVER As Long = &H13
Private Const PRODUCT_SERVER_FOR_SMALLBUSINESS As Long = &H18
Private Const PRODUCT_SMALLBUSINESS_SERVER As Long = &H9
Private Const PRODUCT_SMALLBUSINESS_SERVER_PREMIUM As Long = &H19
Private Const PRODUCT_STANDARD_SERVER As Long = &H7
Private Const PRODUCT_STANDARD_SERVER_CORE As Long = &HD
Private Const PRODUCT_STARTER As Long = &H8
Private Const PRODUCT_STORAGE_ENTERPRISE_SERVER As Long = &H17
Private Const PRODUCT_STORAGE_EXPRESS_SERVER As Long = &H14
Private Const PRODUCT_STORAGE_STANDARD_SERVER As Long = &H15
Private Const PRODUCT_STORAGE_WORKGROUP_SERVER As Long = &H16
Private Const PRODUCT_UNDEFINED As Long = &H0
Private Const PRODUCT_ULTIMATE As Long = &H1
Private Const PRODUCT_ULTIMATE_N As Long = &H1C
Private Const PRODUCT_WEB_SERVER As Long = &H11

Public Declare Function GetTempPath Lib "KERNEL32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Declare Function Htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long
Public Declare Function GetWindowsDirectory Lib "KERNEL32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

''''''''局域网登录
Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type
 
Public Const NO_ERROR = 0
Public Const CONNECT_UPDATE_PROFILE = &H1
' The following includes all the constants defined for NETRESOURCE,
' not just the ones used in this example.
Public Const RESOURCETYPE_DISK = &H1
Public Const RESOURCETYPE_PRINT = &H2
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCE_CONNECTED = &H1
Public Const RESOURCE_REMEMBERED = &H3
Public Const RESOURCE_GLOBALNET = &H2
Public Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEDISPLAYTYPE_SERVER = &H2
Public Const RESOURCEDISPLAYTYPE_SHARE = &H3
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const RESOURCEUSAGE_CONTAINER = &H2
' Error Constants:
Public Const ERROR_ACCESS_DENIED1 = 5&
Public Const ERROR_ALREADY_ASSIGNED = 85&
Public Const ERROR_BAD_DEV_TYPE = 66&
Public Const ERROR_BAD_DEVICE = 1200&
Public Const ERROR_BAD_NET_NAME = 67&
Public Const ERROR_BAD_PROFILE = 1206&
Public Const ERROR_BAD_PROVIDER = 1204&
Public Const ERROR_BUSY = 170&
Public Const ERROR_CANCELLED = 1223&
Public Const ERROR_CANNOT_OPEN_PROFILE = 1205&
Public Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&
Public Const ERROR_EXTENDED_ERROR = 1208&
Public Const ERROR_INVALID_PASSWORD = 86&
Public Const ERROR_NO_NET_OR_BAD_PATH = 1203&

Declare Function WNetAddConnection2 Lib "mpr.dll" Alias _
     "WNetAddConnection2A" (lpNetResource As NETRESOURCE, _
     ByVal lpPassword As String, ByVal lpUserName As String, _
     ByVal dwFlags As Long) As Long
      
''''''''局域网登录

'---------------------------------------------------------------
Public gcnOracle As ADODB.Connection     '数据库连接
Public Const gstrSysName = "中联软件"              '系统名称
Public gstrUserName As String               '用户名
Public gstr姓名 As String                   '用户姓名
Public gstr编号 As String                   '用户编号
Public gstrPassword As String               '用户口令
Public gstrServer As String                 '服务器名
Public gstrSql    As String                 '通用的SQL语句变量
Public rsTemp As New ADODB.Recordset        '公用临时recordset对象
Public gstrProductName As String

Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&

Public objFso As New FileSystemObject   '文件操作对象
Public blnExit As Boolean
Public gmdlOSType As Integer
Public gobjWait As Object

Public Sub main()
    Dim objLogin As Object
    
    On Error Resume Next
    If objLogin Is Nothing Then
        Set objLogin = CreateObject("ZLLogin.clsLogin")
    End If
    If objLogin Is Nothing Then
        MsgBox "创建ZLLogin部件对象失败,请检查文件是否存在并且正确注册。"
        Exit Sub
    Else
        Set gcnOracle = objLogin.Login(1, CStr(Command()))
        If gcnOracle Is Nothing Then
            Exit Sub
        ElseIf gcnOracle.State <> adStateOpen Then
            Exit Sub
        End If
    End If
    Set gobjWait = frmScriptEdit
    Load gobjWait
    
    frmMain.Show
End Sub


Public Sub SelAll(objTxt As Control)
'功能：对文本框的的文本选中
    If TypeName(objTxt) = "TextBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Or InStr(strInput, """") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function


Public Function GetFileLineCount(ByVal txtStream As TextStream) As Long
    Do Until txtStream.AtEndOfStream
        txtStream.ReadLine
    Loop
    
    GetFileLineCount = txtStream.Line
End Function


Public Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "", Optional gcnConnect As ADODB.Connection)
'功能：打开记录集
    If rsTemp.State = adStateOpen Then rsTemp.Close
'    rsTemp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle, adOpenStatic, adLockReadOnly
    If gcnConnect Is Nothing Then Set gcnConnect = gcnOracle
    rsTemp.CursorType = 1
    rsTemp.Open (IIf(strSQL = "", gstrSql, strSQL)), gcnConnect
    '  Call WriteLogFile(IIf(strSQL = "", gstrSql, strSQL))
    End Sub


Public Sub ExecuteProcedure(ByVal strCaption As String)
'功能：执行SQL语句
    gcnOracle.Execute gstrSql, , adCmdStoredProc
   ' Call WriteLogFile(gstrSql)
End Sub

Public Sub WriteLogFile(ByVal strLogContext As String)
'功能：将程序运行日志写入日志文件
'参数：strlogConText 日志内容
  Dim filename As String
  Dim logFile As TextStream
  
  filename = App.Path & "\" & Format(Now(), "yyyy-mm-dd") & "_" & "药品处理.log"
  '1、判断日志文件是否已经存在
  If objFso.FileExists(filename) Then
     Set logFile = objFso.OpenTextFile(filename, ForAppending)
  Else
     Set logFile = objFso.CreateTextFile(filename, True)
  End If
  
  '2、写入日志
  logFile.WriteLine (Format(Now(), "yyyy-mm-dd HH:MM:SS") & " ： ")
  logFile.WriteLine (strLogContext)
  logFile.Close
End Sub

Public Function ExcuteSql(ByVal Recordsettest As ADODB.Recordset, ByVal strSQL As String) As Boolean
'功能：16
'参数：16
    On Error GoTo ExcuteSqlError
    'Set Recordsettest = Nothing
    If Recordsettest.State = adStateOpen Then
       Recordsettest.Close
    End If
    'Recordset.CursorType = adOpenKeyset
    'Recordset.LockType = adLockOptimistic
    Call Recordsettest.Open(strSQL, gcnOracle, adOpenKeyset, adLockOptimistic, -1)
    ExcuteSql = True
    Exit Function
ExcuteSqlError:
    MsgBox "错误代码:" & err.Number & vbCrLf & _
            "错误描述:" & err.Description, vbCritical + vbOKOnly, "错误"
    ExcuteSql = False
End Function

Public Function ConnectUserPassword(ByVal RemoteName As String, ByVal LoginUser As String, ByVal LoginPassWord As String) As Boolean
     Dim NetR As NETRESOURCE
     Dim ErrInfo As Long
     Dim MyPass As String, MyUser As String
      
     NetR.dwScope = RESOURCE_GLOBALNET
     NetR.dwType = RESOURCETYPE_DISK
     NetR.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
     NetR.dwUsage = RESOURCEUSAGE_CONNECTABLE
     NetR.lpLocalName = "" '指定为本机映射盘符 如X:
     NetR.lpRemoteName = RemoteName '"\\192.168.0.3\常用软件"
     'NetR.lpComment = "Optional Comment"
     'NetR.lpProvider = ' Leave this undefined
     
     ErrInfo = WNetAddConnection2(NetR, LoginPassWord, LoginUser, CONNECT_UPDATE_PROFILE) '"snp1108$", "zlsoft\zq",
     If ErrInfo = NO_ERROR Then
        ConnectUserPassword = True
     Else
        ConnectUserPassword = False
     End If
End Function

Public Function GetWindowsVersion() As String
' 变量声明
Dim retOSVersionInf As OSVERSIONINFOEX
Dim retLng As Long

    '结构尺寸
    retOSVersionInf.dwOSVersionInfoSize = Len(retOSVersionInf)
    
    '获取 Windows 版本
    retLng = GetVersionEx(retOSVersionInf)
    
    If retLng = 0 Then
        GetWindowsVersion = "未知"
        Exit Function
    End If

    With retOSVersionInf
        If .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
            gmdlOSType = 0
        ElseIf .dwPlatformId = VER_PLATFORM_WIN32_NT And .dwMajorVersion = 5 Then
            gmdlOSType = 0
        ElseIf .dwPlatformId = VER_PLATFORM_WIN32_NT And .dwMajorVersion = 6 Then
            gmdlOSType = 1
        ElseIf .dwMajorVersion > 6 Then
            gmdlOSType = 1
        End If
        
        GetWindowsVersion = GetWindowsVersion & " [Version: " & .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber & "]"
    End With
End Function

Public Function IsPrivs(ByVal strPrivs As String, ByVal strPriv As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function


Public Function SearchPrintData(ByVal objVsf As Object, ByRef objPrintVsf As Object, Optional strNotPrintCol As String = "") As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strFormat As String
    Dim lngNotPrintCols As Long
    Dim lngPrintCol As Long
    
    If objPrintVsf.Cols = 0 Then Exit Function
    If strNotPrintCol <> "" Then
        lngNotPrintCols = UBound(Split(strNotPrintCol, ",")) + 1
        strNotPrintCol = "," & strNotPrintCol & ","
    End If
    
    objPrintVsf.Rows = objVsf.Rows
    objPrintVsf.FixedRows = objVsf.FixedRows
    
    objPrintVsf.Cols = 0
    lngPrintCol = -1
    For lngCol = 0 To objVsf.Cols - 1
        
        If objVsf.ColHidden(lngCol) = False And objVsf.TextMatrix(0, lngCol) <> "" Then
            
            If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                
                lngPrintCol = lngPrintCol + 1
                
                objPrintVsf.Cols = lngPrintCol + 1
                
                objPrintVsf.ColWidth(lngPrintCol) = objVsf.ColWidth(lngCol)
                objPrintVsf.ColAlignmentFixed(lngPrintCol) = objVsf.ColAlignment(lngCol)
                If objVsf.ColDataType(lngCol) = flexDTBoolean Then
                    objPrintVsf.ColAlignment(lngPrintCol) = 4
                Else
                    objPrintVsf.ColAlignment(lngPrintCol) = objVsf.ColAlignment(lngCol)
                End If
            End If
        End If
    Next
    
    If objPrintVsf.Cols = 0 Then Exit Function
    
    For lngRow = 0 To objVsf.Rows - 1

        objPrintVsf.RowHeight(lngRow) = IIf(objVsf.RowHeight(lngRow) < objVsf.RowHeightMin, objVsf.RowHeightMin, objVsf.RowHeight(lngRow))
        lngPrintCol = -1
        For lngCol = 0 To objVsf.Cols - 1
            
            If objVsf.ColHidden(lngCol) = False And objVsf.TextMatrix(0, lngCol) <> "" Then
                If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                
                    lngPrintCol = lngPrintCol + 1
                    
                    If objVsf.ColDataType(lngCol) = flexDTBoolean And lngRow >= objVsf.FixedRows Then
                        objPrintVsf.TextMatrix(lngRow, lngPrintCol) = IIf(Abs(Val(objVsf.TextMatrix(lngRow, lngCol))) = 1, "√", "")
                    Else
                        strFormat = objVsf.ColFormat(lngCol)
                        If strFormat = "" Then
                            objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Trim(objVsf.TextMatrix(lngRow, lngCol))
                        Else
                            objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Format(objVsf.TextMatrix(lngRow, lngCol), strFormat)
                        End If
                    End If
                End If
            End If
        Next
        Call SetMsfForeColor(objPrintVsf, lngRow, Val(objVsf.Cell(flexcpForeColor, lngRow, 1)))
    Next
    SearchPrintData = True
End Function

Public Sub SetMsfForeColor(ByRef msf As Object, ByVal lngRow As Long, ByVal lngColor As Long)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intCol As Integer

    With msf

        .Row = lngRow
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellForeColor = lngColor
        Next

    End With
End Sub
