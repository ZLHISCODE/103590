VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于..."
   ClientHeight    =   4395
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6300
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3033.507
   ScaleMode       =   0  'User
   ScaleWidth      =   5907.157
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "s"
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   4590
      TabIndex        =   0
      Top             =   3420
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "系统信息(&S)..."
      Height          =   345
      Left            =   4605
      TabIndex        =   1
      Top             =   3810
      Width           =   1485
   End
   Begin VB.Label lbl管理工具 
      AutoSize        =   -1  'True
      Caption         =   "管理工具数据版本：#"
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   225
      TabIndex        =   13
      Top             =   2820
      Width           =   1710
   End
   Begin VB.Label lblGrant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   1320
      TabIndex        =   12
      Top             =   1305
      Width           =   90
   End
   Begin VB.Label lbl开发商 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   180
      Left            =   1320
      TabIndex        =   11
      Top             =   2370
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "产品开发商："
      Height          =   180
      Left            =   225
      TabIndex        =   10
      Top             =   2370
      Width           =   1080
   End
   Begin VB.Label lbl技术支持商 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   180
      Left            =   1320
      TabIndex        =   9
      Top             =   1837
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "技术支持商："
      Height          =   180
      Left            =   225
      TabIndex        =   8
      Top             =   1837
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   0
      X2              =   6817.61
      Y1              =   828.261
      Y2              =   828.261
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   0
      X2              =   6831.674
      Y1              =   817.908
      Y2              =   817.908
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "使用权授予："
      Height          =   180
      Left            =   225
      TabIndex        =   7
      Top             =   1305
      Width           =   1080
   End
   Begin VB.Image imgLogo 
      Height          =   720
      Left            =   285
      Picture         =   "frmAbout.frx":0E42
      Stretch         =   -1  'True
      Top             =   165
      Width           =   720
   End
   Begin VB.Label lblSysName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "管理工具"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000087&
      Height          =   405
      Left            =   1410
      TabIndex        =   6
      Top             =   105
      Width           =   1740
   End
   Begin VB.Label lblPlatform 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "For Windows/Oracle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000087&
      Height          =   330
      Left            =   2790
      TabIndex        =   5
      Top             =   825
      Width           =   2595
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "ZLBaseCode Version 2.01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000087&
      Height          =   330
      Left            =   2040
      TabIndex        =   4
      Top             =   525
      Width           =   3405
   End
   Begin VB.Label lblCopyRight 
      AutoSize        =   -1  'True
      Caption         =   "版权所有(C) 中联信息产业公司"
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   3555
      TabIndex        =   3
      Top             =   2820
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -84.388
      X2              =   6747.287
      Y1              =   2267.365
      Y2              =   2267.365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   -70.323
      X2              =   6747.287
      Y1              =   2277.719
      Y2              =   2277.719
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":1D0C
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   210
      TabIndex        =   2
      Top             =   3525
      Width           =   4305
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintMouse As Integer
              
Private Const gREGKEYSYSINFOLOC = "SOFTWaRE\Microsoft\Shared Tools Location"
Private Const gREGVaLSYSINFOLOC = "MSINFO"
Private Const gREGKEYSYSINFO = "SOFTWaRE\Microsoft\Shared Tools\MSINFO"
Private Const gREGVaLSYSINFO = "PaTH"

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub StartSysInfo()
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
    MsgBox "此时系统信息无效" & vbNewLine & err.Description, vbOKOnly, gstrSysName
End Sub

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
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
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 打开注册键
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 句柄错误...
    
    tmpVal = String$(1024, 0)                             ' 分配变量空间
    KeyValSize = 1024                                       ' 标记变量大小
    
    '------------------------------------------------------------
    ' 检索注册键值...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' 获得/创建键值
                        
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

Public Sub ShowAbout()
'功能： 显示关于窗体

    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, objItem As ListItem
    Dim strKind As String, strCode As String
    Dim strSerial As String, strSQL As String
        
    On Error GoTo errH
    
    strKind = gobjRegister.zlRegInfo("授权性质")
    If strKind = "2" Then
        strKind = "(试用)"
    ElseIf strKind = "3" Then
        strKind = "(测试)"
    Else
        strKind = ""
    End If
    With frmAbout
        .lblSysName.Caption = App.Title & strKind
        .lblVersion.Caption = App.ProductName & " Version " & App.Major & "." & App.Minor & "." & App.Revision
        .lblGrant.Caption = Replace(gobjRegister.zlRegInfo("单位名称", , -1), ";", vbCrLf)
        
        .lbl技术支持商.Caption = gobjRegister.zlRegInfo("技术支持商", , -1)
        Call ApplyOEM_Picture(imgLogo, "Picture")
        
        If Trim$(.lbl技术支持商.Caption) = "" Then
            .Label1.Visible = False
            .lbl技术支持商.Visible = False
            .lblCopyRight.Visible = False
        Else
            .Label1.Visible = True
            .lbl技术支持商.Visible = True
            .lblCopyRight.Visible = True
        End If
        
        strCode = gobjRegister.zlRegInfo("产品开发商", , -1)
        If Trim(strCode) = "" Then
            .Label3.Visible = False
            .lbl开发商.Visible = False
        Else
            .Label3.Visible = True
            .lbl开发商.Visible = True
            .lbl开发商.Caption = ""
            For i = 0 To UBound(Split(strCode, ";"))
                .lbl开发商.Caption = .lbl开发商.Caption & Split(strCode, ";")(i) & vbCrLf
            Next
        End If
        
        '显示管理工具本身的版本号
        strCode = gobjRegister.zlRegInfo("版本号")
        If strCode = "" Then
            lbl管理工具.Visible = False
        Else
            lbl管理工具.Caption = "管理工具数据版本：" & strCode
        End If
        .Refresh
    End With
    frmAbout.Show 1, frmMDIMain
   
errH:
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    With Me.lblCopyRight
        If X >= .Left And X <= .Left + .Width And y >= .Top And y <= .Top + .Height Then
            mintMouse = mintMouse + 1
            If mintMouse = 9 Then .Visible = True
        Else
            mintMouse = 0
            .Visible = False
        End If
    End With
End Sub
