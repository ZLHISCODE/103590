VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ں�������"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "ȷ��(&O)"
      Height          =   330
      Left            =   4560
      TabIndex        =   0
      Top             =   3210
      Width           =   1350
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "ϵͳ��Ϣ(&S)"
      Height          =   330
      Left            =   4560
      TabIndex        =   1
      Top             =   3465
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":014A
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   240
      TabIndex        =   12
      Top             =   3090
      Width           =   4230
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   -60
      X2              =   7515
      Y1              =   2910
      Y2              =   2910
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   15
      X2              =   7605
      Y1              =   2895
      Y2              =   2895
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2265
      TabIndex        =   11
      Top             =   555
      Width           =   1410
   End
   Begin VB.Label lblPlatform 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "For Windows/Oracle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      Top             =   870
      Width           =   2310
   End
   Begin VB.Label lblSysName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   1620
      TabIndex        =   9
      Top             =   105
      Width           =   2250
   End
   Begin VB.Label lbl�����̱��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ʒ�����̣�"
      Height          =   180
      Left            =   1680
      TabIndex        =   8
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����֧���̣�"
      Height          =   180
      Left            =   1665
      TabIndex        =   7
      Top             =   1620
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ʹ��Ȩ���裺"
      Height          =   180
      Left            =   1665
      TabIndex        =   6
      Top             =   1335
      Width           =   1080
   End
   Begin VB.Label lblCopyRight 
      AutoSize        =   -1  'True
      Caption         =   "��Ȩ����(C) 2000-2001 ������Ϣ��ҵ���޹�˾"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   1665
      TabIndex        =   5
      Top             =   2685
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.Label lblGrant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   2775
      TabIndex        =   4
      Top             =   1335
      Width           =   90
   End
   Begin VB.Label lbl������ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   180
      Left            =   2790
      TabIndex        =   3
      Top             =   1920
      Width           =   90
   End
   Begin VB.Label lbl����֧���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   180
      Left            =   2775
      TabIndex        =   2
      Top             =   1620
      Width           =   90
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   1560
      X2              =   6109
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   1560
      X2              =   6109
      Y1              =   1215
      Y2              =   1215
   End
   Begin VB.Image imgLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   2865
      Left            =   45
      Stretch         =   -1  'True
      Top             =   15
      Width           =   1440
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ע�����ȫѡ��...
Const KEY_ALL_ACCESS = &H2003F
' ע���������...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode �ս�β�ַ���
Const REG_DWORD = 4                      ' 32λ��
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

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
        ' ��ע�����ϵͳ��Ϣ����·��\����...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' ����ע�����ϵͳ��Ϣ����·��...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
            ' ��֤��֪�� 32 λ�ļ��汾�Ĵ���
            If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            ' ���� - �ļ��Ҳ���...
            Else
                GoTo SysInfoErr
            End If
        ' ���� - ע������Ҳ���...
        Else
            GoTo SysInfoErr
        End If
        Call Shell(SysInfoPath, vbNormalFocus)
        Exit Sub
SysInfoErr:
        Resume
        MsgBox "��ʱϵͳ��Ϣ������", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' ѭ��������
        Dim rc As Long                                          ' ���ش���
        Dim hKey As Long                                        ' �򿪵�ע�������
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' ע������������
        Dim tmpVal As String                                    ' ��ʱ�洢һ��ע����ֵ
        Dim KeyValSize As Long                                  ' ע����������С
        '------------------------------------------------------------
        ' �ڼ���{HKEY_LOCAL_MACHINE...}֮�´�ע���
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ��ע����
        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ������...
        tmpVal = String$(1024, 0)                             ' ��������ռ�
        KeyValSize = 1024                                       ' ��Ǳ�����С
        '------------------------------------------------------------
        ' ����ע����ֵ...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' ���/������ֵ
        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ������
        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        '------------------------------------------------------------
        ' ����ת���ļ�ֵ����...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' ������������...
        Case REG_SZ                                             ' �ַ���ע������������
                KeyVal = tmpVal                                     ' �����ַ���ֵ
        Case REG_DWORD                                          ' ˫����ע������������
                For i = Len(tmpVal) To 1 Step -1                    ' ת��ÿһҳ
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' һ���ַ�һ���ַ�������ֵ
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' ת��˫����Ϊ�ַ���
        End Select
        GetKeyValue = True                                      ' ���سɹ�
        rc = RegCloseKey(hKey)                                  ' �ر�ע����
        Exit Function                                           ' �˳�
GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                             ' �践��ֵΪ���ַ���
        GetKeyValue = False                                     ' ����ʧ��
        rc = RegCloseKey(hKey)                                  ' �ر�ע����
End Function

Private Sub Form_Load()
    Dim strKind As String, strCode As String
    Dim strSerial As String, strSQL As String
    Dim i As Integer
        
    If Not gcnOracle Is Nothing Then
        strKind = Decode(zlRegInfo("��Ȩ����"), "2", "(����)", "3", "(����)", "")
    Else
        strKind = GetSetting("ZLSOFT", "ע����Ϣ", "KIND", "")
        strKind = IIf(strKind = "" Or strKind = "��ʽ", "", "(" & strKind & ")")
    End If
    lblSysName.Caption = App.Title & strKind
    lblVersion.Caption = App.ProductName & " Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblVersion.Caption = Replace(lblVersion.Caption, "zl9", "")
    lblVersion.Caption = Replace(lblVersion.Caption, "ZL9", "")
    lblVersion.Caption = Replace(lblVersion.Caption, "zl", "")
    lblVersion.Caption = Replace(lblVersion.Caption, "ZL", "")
    
    If Not gcnOracle Is Nothing Then
        lblGrant.Caption = zlRegInfo("��λ����", , -1)
        lbl����֧����.Caption = zlRegInfo("����֧����", , -1)
        
        strCode = zlRegInfo("��Ʒ������", , -1)
        lbl�����̱���.Visible = strCode <> ""
        lbl������.Visible = strCode <> ""
        lbl������.Caption = ""
        For i = 0 To UBound(Split(strCode, ";"))
            lbl������.Caption = lbl������.Caption & Split(strCode, ";")(i) & vbCrLf
        Next
    Else
        lblGrant.Caption = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����")
        lbl����֧����.Caption = GetSetting("ZLSOFT", "ע����Ϣ", "����֧����")
        strCode = GetSetting("ZLSOFT", "ע����Ϣ", "������")
        
        lbl�����̱���.Visible = strCode <> ""
        lbl������.Visible = strCode <> ""
        lbl������.Caption = ""
        For i = 0 To UBound(Split(strCode, ";"))
            lbl������.Caption = lbl������.Caption & Split(strCode, ";")(i) & vbCrLf
        Next
    End If
    
    Set imgLogo.Picture = LoadCustomPicture("Function")
End Sub
