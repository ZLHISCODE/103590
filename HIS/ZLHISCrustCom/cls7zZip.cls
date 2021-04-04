VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls7zZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mclsZip     As Object       'ZL7z的压缩类
Private mclsUnZip   As Object       'ZL7z的解压类
Private mblnZL7z    As Boolean      '是否能正常使用Zl7z

Private mobjFSO     As New FileSystemObject
Private mobjShell   As New clsShell
Private mstr7zExe   As String       '7z.exe的路径

Property Get Path7z() As String
    Path7z = mstr7zExe
End Property

Property Let Path7z(ByVal str7zExe As String)
    mstr7zExe = str7zExe
End Property

Public Function Init7zZip(Optional ByVal str7zPath As String) As Boolean
'功能：进行7z相关功能的初始化
    Dim strSystem As String
    
    If str7zPath <> "" And Dir(str7zPath) <> "" And (mstr7zExe = "" Or UCase(mstr7zExe) <> UCase(str7zPath)) Then
        mstr7zExe = str7zPath
    Else
        strSystem = mobjFSO.GetSpecialFolder(SystemFolder)
        If Is64bit Then '64系统下32位程序应该放在C:\windows\SysWOW64
            strSystem = mobjFSO.GetParentFolderName(strSystem) & "\SysWOW64"
        End If
        If Dir(strSystem & "\7z.exe") <> "" Then
            mstr7zExe = strSystem & "\7z.exe"
        End If
    End If
    If Not mobjFSO.FileExists(mstr7zExe) Then
        mstr7zExe = ""
    End If
    
    On Error Resume Next
    Set mclsZip = CreateObject("zl7z.Zip")
    If Err.Number <> 0 Then
       Err.Clear
    End If
    Set mclsUnZip = CreateObject("zl7z.UnZip")
    If Err.Number <> 0 Then
       Err.Clear
    End If
    mblnZL7z = Not (mclsZip Is Nothing Or mclsUnZip Is Nothing)
    If mblnZL7z Then
        mblnZL7z = Val(GetSetting("ZLSOFT", "公共模块\自动升级", "7ZEXE调试", "0")) = 0
    End If
    Init7zZip = mstr7zExe <> "" Or mblnZL7z
End Function

Public Function ZipFile(ByVal strFile As String, ByVal strZipName As String, Optional ByVal intRate As Integer = 5, Optional ByRef strErr As String) As Boolean
'功能：进行压缩
'参数：
'       strZipName=保存的文件路径与名称
'       strFile=原始文件
'       intRate=压缩等级，压缩使用。
    If mblnZL7z Then
        Call mclsZip.ZipSingle(strFile, gobjFSO.GetFileName(strFile), strZipName, intRate)
        ZipFile = gobjFSO.FileExists(strZipName)
    Else
        strErr = ""
        ZipFile = DeCompression(strZipName, strFile, intRate, True, strErr)
        ZipFile = ZipFile And strErr = "" And gobjFSO.FileExists(strZipName)
    End If
End Function

Public Function UnZipFile(ByVal strZipName As String, ByVal strFile As String, Optional ByVal blnDeleteZipFile As Boolean = True, Optional ByRef strErr As String, Optional blnFolder As Boolean) As Boolean
'功能：进行压缩
'参数：
'       strZipName=压缩文件名
'       strFile=解压后的名称
    If mblnZL7z Then
        Err.Clear: strErr = ""
        On Error Resume Next
        mclsUnZip.UnZipSingle strZipName, gobjFSO.GetParentFolderName(strFile)
        If Err.Number <> 0 Then
            UnZipFile = False
            strErr = "ZL7Z(压缩文件格式损坏):" & Err.Description
            Err.Clear
        Else
            If blnFolder Then
                UnZipFile = gobjFSO.FolderExists(strFile)
            Else
                UnZipFile = gobjFSO.FileExists(strFile)
            End If
        End If
    Else
        strErr = ""
        UnZipFile = DeCompression(strFile, strZipName, , , strErr, blnFolder)
        If blnFolder Then
            UnZipFile = UnZipFile And strErr = "" And gobjFSO.FolderExists(strFile)
        Else
            UnZipFile = UnZipFile And strErr = "" And gobjFSO.FileExists(strFile)
        End If
    End If
    If blnDeleteZipFile Then
        On Error Resume Next
        If FileSystem.GetAttr(strZipName) <> vbNormal Then
             Call FileSystem.SetAttr(strZipName, vbNormal)
        End If
        '删除原始文件
        Call gobjFSO.DeleteFile(strZipName, True)
        If Err.Number <> 0 Then Err.Clear
    End If
End Function

'===========================================================================
'                2、私有方法
'===========================================================================
Private Function DeCompression(ByVal strDesFile As String, ByVal strSourceFile As String, Optional ByVal intRate As Integer, Optional ByVal blnCompression As Boolean, Optional ByRef strErr As String, Optional blnFolder As Boolean) As Boolean
'功能：进行压缩解压(当前仅支持单文件）
'参数：
'       strDesFile=保存的文件路径与名称
'       strSourceFile=原始文件
'       intRate=压缩等级，压缩使用。
'                   压缩等级 压缩算法 字典大小 快速字节 匹配器 过滤器 描述
'                   0           Copy    无压缩
'                   1           LZMA    64KB     32       HC4   BCJ   最快压缩
'                   3           LZMA    1MB      32       HC4   BCJ   快速压缩
'                   5           LZMA    16MB     32       BT4   BCJ   正常压缩(默认等级）
'                   7           LZMA    32MB     64       BT4   BCJ   最大压缩
'                   9           LZMA    64MB     64       BT4   BCJ2  极限压缩
'       blnCompression=True-压缩，False-解压
'返回：是否成功
'说明：解压缩文件到本地,并删除压缩原始文件
    Dim strCommand As String, strReturn As String
    '获取不了7Z文件路径，则直接退出
    If mstr7zExe = "" Then
        strErr = "7Z.EXE解压程序不存在"
        Exit Function
    End If
    If Not gobjFSO.FileExists(strSourceFile) Then
        strErr = "源文件" & strSourceFile & "不存在"
        Exit Function
    End If
    If gobjFSO.FileExists(strDesFile) Then
        On Error Resume Next
        '删除存在的目的文件
        If FileSystem.GetAttr(strDesFile) <> vbNormal Then
             Call FileSystem.SetAttr(strDesFile, vbNormal)
        End If
        Call gobjFSO.DeleteFile(strDesFile, True)
        If Err.Number <> 0 Then Err.Clear
    End If
    On Error GoTo ErrH
    If blnCompression Then
        '-m 固定传输字符 x=设置等级 mt开启或关闭多线程压缩模式
        strCommand = """" & mstr7zExe & """  a -y """ & strDesFile & """ """ & strSourceFile & """ -mx=" & intRate & " -mmt"
    Else
        '-o 固定传输字符
        If blnFolder Then
            strCommand = """" & mstr7zExe & """  x -y """ & strSourceFile & """ -o""" & strDesFile & """"
        Else
            strCommand = """" & mstr7zExe & """  e -y """ & strSourceFile & """ -o""" & gobjFSO.GetParentFolderName(strDesFile) & """"
        End If
    End If
    If mobjShell.Run(strCommand, strReturn, strErr, 5000) Then
        If strErr = "" And strReturn <> "" Then strErr = strReturn
        If blnFolder Then
            DeCompression = gobjFSO.FolderExists(strDesFile)
        Else
            DeCompression = gobjFSO.FileExists(strDesFile)
        End If
    Else
        DeCompression = False
    End If
    Exit Function
ErrH:
    If strErr = "" Then strErr = Err.Description
    If 0 = 1 Then
        Resume
    End If
End Function

Private Sub Class_Terminate()
    Set mobjFSO = Nothing
    Set mobjShell = Nothing
End Sub
