VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����..."
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
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   5190
      TabIndex        =   0
      Top             =   3765
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "ϵͳ��Ϣ(&S)..."
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
      Caption         =   "�������"
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
      Caption         =   "��ӡ֧��ϵͳ"
      BeginProperty Font 
         Name            =   "����_GB2312"
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
      Caption         =   "���棺���������������������ʹ�����֤������δ��������˾��ɣ��κ��˲��ø��ơ����ۼ����ܴ���������򽫳е�ȫ���������Ρ�"
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
      Caption         =   "�������"
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

' ע�����ȫѡ��...
Const KEY_CREaTE_Sub_KEY = &H4
Const KEY_CREaTE_LINK = &H20
Const KEY_aLL_aCCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREaTE_Sub_KEY + KEY_ENUMERATE_Sub_KEYS + _
                       KEY_NOTIFY + KEY_CREaTE_LINK + READ_CONTROL
                     
' ע��� ROOT ����...
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode �� Null ��β���ַ���
Const REG_DWORD = 4                      ' 32-λ����

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
'StartSysInfo       ���еó�ϵͳ��Ϣ�ĳ���
'GetKeyValue        ��ע���ָ��λ�ö�������
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
'���ܣ����еó�ϵͳ��Ϣ�ĳ���
'��������
'���أ���

    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' ��ͼ��ע���õ�ϵͳ��Ϣ����·��\����...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVaLSYSINFO, SysInfoPath) Then
    ' ��ͼ��ע���õ�ϵͳ��Ϣ����·��...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVaLSYSINFOLOC, SysInfoPath) Then
        ' ��֤��֪ 32 λ�ļ��汾�Ĵ���
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' ���� - �ļ�δ�ҵ�...
        Else
            GoTo SysInfoErr
        End If
    ' ���� - ע����δ�ҵ�...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "��ʱϵͳ��Ϣ��Ч", vbOKOnly, gstrSysName
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
'���ܣ���ע���ָ��λ�ö�������
'������KeyRoot          ������ֵ
'      KeyName          Ҫ���ʵļ���
'      SubKeyRef        Ҫ���ʵ�������
'      KeyVal           �õ���ֵ
'���أ��ɹ�����True ,ʧ�ܷ���False

    Dim i As Long                                           ' ѭ��ָ��
    Dim rc As Long                                          ' ���ش���
    Dim hKey As Long                                        ' �򿪵�ע����ľ��
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' ע�������������
    Dim tmpVal As String                                    ' ע�������ʱ�洢��
    Dim KeyValSize As Long                                  ' ע��������Ĵ�С
    '------------------------------------------------------------
    ' �ڸ��� {HKEY_LOCaL_MaCHINE...} �´�ע���
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_aLL_aCCESS, hKey)  ' ��ע���
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������...
    
    tmpVal = String$(1024, 0)                             ' ��������ռ�
    KeyValSize = 1024                                       ' ��Ǳ�����С
    
    '------------------------------------------------------------
    ' ����ע���ֵ...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, ByVal tmpVal, KeyValSize)    ' ���/������ֵ
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 ����� Null ��β���ַ���...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null �ҵ������ַ�����ȡ
    Else                                                    ' WinNT ����Ҫ�� Null �����ַ���...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null δ�ҵ��� ����ȡ�ַ���
    End If
    '------------------------------------------------------------
    ' Ϊ��ת����������ֵ����..
    '------------------------------------------------------------
    Select Case KeyValType                                  ' ������������...
    Case REG_SZ                                             ' �ַ�����ע�����������
        KeyVal = tmpVal                                     ' �����ַ���ֵ
    Case REG_DWORD                                          ' ˫����ע�����������
        For i = Len(tmpVal) To 1 Step -1                    ' ת��ÿһλ
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' һ���ַ�һ���ַ��ؽ���ֵ
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' ת��˫����Ϊ�ַ�����
    End Select
    
    GetKeyValue = True                                      ' ���سɹ�
    rc = RegCloseKey(hKey)                                  ' �ر�ע���
    Exit Function                                           ' �˳�
    
GetKeyError:      ' ������������...
    KeyVal = ""                                             ' ���÷���ֵΪ���ַ���
    GetKeyValue = False                                     ' ����ʧ��
    rc = RegCloseKey(hKey)                                  ' �ر�ע���
End Function

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    lblCompanyProduct(0).Caption = gstrSysName
    lblCompanyProduct(1).Caption = gstrSysName
    lblSysName.Caption = Mid(gstrSysName, 1, Len(gstrSysName) - 2) & "��ӡ֧��ϵͳ"
End Sub

