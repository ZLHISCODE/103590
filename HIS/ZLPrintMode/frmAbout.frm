VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于..."
   ClientHeight    =   4500
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6990
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105.98
   ScaleMode       =   0  'User
   ScaleWidth      =   6554.132
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   5190
      TabIndex        =   0
      Top             =   3765
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "系统信息(&S)..."
      Height          =   345
      Left            =   5190
      TabIndex        =   1
      Top             =   4065
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "中联软件"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   480
      Index           =   1
      Left            =   195
      TabIndex        =   6
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Image imgLogo 
      Height          =   720
      Left            =   780
      Picture         =   "frmAbout.frx":030A
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lblSysName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "打印支持系统"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   2670
      TabIndex        =   5
      Top             =   330
      Width           =   2985
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "For Windows NT/Windows 9X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2925
      TabIndex        =   4
      Top             =   1245
      Width           =   3570
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 2.01"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5085
      TabIndex        =   3
      Top             =   1635
      Width           =   1425
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -84.388
      X2              =   6747.287
      Y1              =   2360.545
      Y2              =   2360.545
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   -70.323
      X2              =   6747.287
      Y1              =   2370.898
      Y2              =   2370.898
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "警告：本软件受软件保护法和软件使用许可证保护。未经中联公司许可，任何人不得复制、销售及解密此软件，否则将承担全部法律责任。"
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   345
      TabIndex        =   2
      Top             =   3615
      Width           =   4635
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "中联软件"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   210
      TabIndex        =   7
      Top             =   1215
      Width           =   1800
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 注册键安全选项...
Const KEY_CREaTE_Sub_KEY = &H4
Const KEY_CREaTE_LINK = &H20
Const KEY_aLL_aCCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREaTE_Sub_KEY + KEY_ENUMERATE_Sub_KEYS + _
                       KEY_NOTIFY + KEY_CREaTE_LINK + READ_CONTROL
                     
' 注册键 ROOT 类型...
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode 以 Null 结尾的字符串
Const REG_DWORD = 4                      ' 32-位数字

Const gREGKEYSYSINFOLOC = "SOFTWaRE\Microsoft\Shared Tools Location"
Const gREGVaLSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWaRE\Microsoft\Shared Tools\MSINFO"
Const gREGVaLSYSINFO = "PaTH"

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public btname As Object, yy As Boolean
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySounda" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function messagebeep Lib "user32" (ByVal wtype As Integer) As Integer


'---------------------------------------------------------------------------------------------------
'StartSysInfo       运行得出系统信息的程序
'GetKeyValue        从注册表指定位置读出数据
'
'
'
'-------------------------------------------------------------------------------------------------------


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Public Sub StartSysInfo()
'功能：运行得出系统信息的程序
'参数：无
'返回：无

    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' 试图从注册表得到系统信息程序路径\名称...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVaLSYSINFO, SysInfoPath) Then
    ' 试图从注册表得到系统信息程序路径...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVaLSYSINFOLOC, SysInfoPath) Then
        ' 验证已知 32 位文件版本的存在
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' 错误 - 文件未找到...
        Else
            GoTo SysInfoErr
        End If
    ' 错误 - 注册项未找到...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "此时系统信息无效", vbOKOnly, gstrSysName
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
'功能：从注册表指定位置读出数据
'参数：KeyRoot          根键的值
'      KeyName          要访问的键名
'      SubKeyRef        要访问的数据名
'      KeyVal           得到的值
'返回：成功返回True ,失败返回False

    Dim i As Long                                           ' 循环指针
    Dim rc As Long                                          ' 返回代码
    Dim hKey As Long                                        ' 打开的注册键的句柄
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' 注册键的数据类型
    Dim tmpVal As String                                    ' 注册键的临时存储区
    Dim KeyValSize As Long                                  ' 注册键变量的大小
    '------------------------------------------------------------
    ' 在根键 {HKEY_LOCaL_MaCHINE...} 下打开注册键
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_aLL_aCCESS, hKey)  ' 打开注册键
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 句柄错误...
    
    tmpVal = String$(1024, 0)                             ' 分配变量空间
    KeyValSize = 1024                                       ' 标记变量大小
    
    '------------------------------------------------------------
    ' 检索注册键值...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, ByVal tmpVal, KeyValSize)    ' 获得/创建键值
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 句柄错误
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 添加以 Null 结尾的字符串...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null 找到，从字符串提取
    Else                                                    ' WinNT 不需要以 Null 结束字符串...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null 未找到， 仅提取字符串
    End If
    '------------------------------------------------------------
    ' 为了转换而决定键值类型..
    '------------------------------------------------------------
    Select Case KeyValType                                  ' 搜索数据类型...
    Case REG_SZ                                             ' 字符串型注册键数据类型
        KeyVal = tmpVal                                     ' 复制字符串值
    Case REG_DWORD                                          ' 双字型注册键数据类型
        For i = Len(tmpVal) To 1 Step -1                    ' 转换每一位
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 一个字符一个字符地建立值
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' 转换双字型为字符串型
    End Select
    
    GetKeyValue = True                                      ' 返回成功
    rc = RegCloseKey(hKey)                                  ' 关闭注册键
    Exit Function                                           ' 退出
    
GetKeyError:      ' 发生错误后清除...
    KeyVal = ""                                             ' 设置返回值为空字符串
    GetKeyValue = False                                     ' 返回失败
    rc = RegCloseKey(hKey)                                  ' 关闭注册键
End Function

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    lblCompanyProduct(0).Caption = gstrSysName
    lblCompanyProduct(1).Caption = gstrSysName
    lblSysName.Caption = Mid(gstrSysName, 1, Len(gstrSysName) - 2) & "打印支持系统"
End Sub

