VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����..."
   ClientHeight    =   5310
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6300
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Tag             =   "s"
   Begin MSComctlLib.ListView lvw 
      Height          =   1695
      Left            =   135
      TabIndex        =   2
      Tag             =   "s"
      Top             =   1755
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2990
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "˵��"
         Object.Width           =   6114
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4575
      TabIndex        =   0
      Top             =   4680
      Width           =   1485
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "ϵͳ��Ϣ(&S)..."
      Height          =   350
      Left            =   4575
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2610
      Top             =   2550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":0E42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   5235
      TabIndex        =   15
      Top             =   630
      Width           =   90
   End
   Begin VB.Label lblGrant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   1500
      TabIndex        =   14
      Top             =   1305
      Width           =   90
   End
   Begin VB.Label lbl������ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   180
      Left            =   1320
      TabIndex        =   13
      Top             =   3780
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ʒ�����̣�"
      Height          =   180
      Left            =   225
      TabIndex        =   12
      Top             =   3780
      Width           =   1080
   End
   Begin VB.Label lbl����֧���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   180
      Left            =   1320
      TabIndex        =   11
      Top             =   3525
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����֧���̣�"
      Height          =   180
      Left            =   225
      TabIndex        =   10
      Top             =   3525
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   0
      X2              =   7271
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   0
      X2              =   7286
      Y1              =   1185
      Y2              =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ʹ��Ȩ�����裺"
      Height          =   180
      Left            =   225
      TabIndex        =   9
      Top             =   1305
      Width           =   1260
   End
   Begin VB.Image imgLogo 
      Height          =   720
      Left            =   285
      Picture         =   "frmAbout.frx":0F24
      Stretch         =   -1  'True
      Top             =   165
      Width           =   720
   End
   Begin VB.Label lblFunc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������¹��ܣ�"
      Height          =   180
      Left            =   225
      TabIndex        =   8
      Top             =   1545
      Width           =   1260
   End
   Begin VB.Label lblSysName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������ϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   1410
      TabIndex        =   7
      Top             =   480
      Width           =   1890
   End
   Begin VB.Label lblPlatform 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "����ҽԺ��Ϣϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   1125
      TabIndex        =   6
      Top             =   90
      Width           =   2520
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   1860
      TabIndex        =   5
      Top             =   810
      Width           =   3405
   End
   Begin VB.Label lblCopyRight 
      AutoSize        =   -1  'True
      Caption         =   "��Ȩ����(C) ������Ϣ��ҵ��˾"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   3555
      TabIndex        =   4
      Top             =   1530
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -90
      X2              =   7196
      Y1              =   4335
      Y2              =   4335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   -75
      X2              =   7196
      Y1              =   4350
      Y2              =   4350
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":1DEE
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   210
      TabIndex        =   3
      Top             =   4575
      Width           =   4305
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ע�����ȫѡ��...
Const REaD_CONTROL = &H20000
Const KEY_QUERY_VaLUE = &H1
Const KEY_SET_VaLUE = &H2
Const KEY_CREaTE_Sub_KEY = &H4
Const KEY_ENUMERaTE_Sub_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREaTE_LINK = &H20
Const KEY_aLL_aCCESS = KEY_QUERY_VaLUE + KEY_SET_VaLUE + _
                       KEY_CREaTE_Sub_KEY + KEY_ENUMERaTE_Sub_KEYS + _
                       KEY_NOTIFY + KEY_CREaTE_LINK + REaD_CONTROL
                     
' ע��� ROOT ����...
Const HKEY_LOCaL_MaCHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode �� Null ��β���ַ���
Const REG_DWORD = 4                      ' 32-λ����

Const gREGKEYSYSINFOLOC = "SOFTWaRE\Microsoft\Shared Tools Location"
Const gREGVaLSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWaRE\Microsoft\Shared Tools\MSINFO"
Const gREGVaLSYSINFO = "PaTH"

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal uloptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public btname As Object, yy As Boolean
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySounda" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function messagebeep Lib "User32.dll" (ByVal wtype As Integer) As Integer
Dim intMouse As Integer

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' ��ͼ��ע���õ�ϵͳ��Ϣ����·��\����...
    If GetKeyValue(HKEY_LOCaL_MaCHINE, gREGKEYSYSINFO, gREGVaLSYSINFO, SysInfoPath) Then
    ' ��ͼ��ע���õ�ϵͳ��Ϣ����·��...
    ElseIf GetKeyValue(HKEY_LOCaL_MaCHINE, gREGKEYSYSINFOLOC, gREGVaLSYSINFOLOC, SysInfoPath) Then
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
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_aLL_aCCESS, hKey) ' ��ע���
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������...
    
    tmpVal = String$(1024, 0)                             ' ��������ռ�
    KeyValSize = 1024                                       ' ��Ǳ�����С
    
    '------------------------------------------------------------
    ' ����ע���ֵ...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' ���/������ֵ
                        
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

Private Sub Form_Activate()
    lvw.SetFocus
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With Me.lblCopyRight
        If X >= .Left And X <= .Left + .Width And Y >= .Top And Y <= .Top + .Height Then
            intMouse = intMouse + 1
            If intMouse = 9 Then .Visible = True
        Else
            intMouse = 0
            .Visible = False
        End If
    End With
End Sub

Private Sub lvw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not lvw.HitTest(X, Y) Is Nothing Then
        lvw.ToolTipText = lvw.HitTest(X, Y).SubItems(2)
    Else
        lvw.ToolTipText = ""
    End If
End Sub
