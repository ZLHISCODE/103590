VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����..."
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
   StartUpPosition =   1  '����������
   Tag             =   "s"
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   4590
      TabIndex        =   0
      Top             =   3420
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "ϵͳ��Ϣ(&S)..."
      Height          =   345
      Left            =   4605
      TabIndex        =   1
      Top             =   3810
      Width           =   1485
   End
   Begin VB.Label lbl������ 
      AutoSize        =   -1  'True
      Caption         =   "���������ݰ汾��#"
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
   Begin VB.Label lbl������ 
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
      Caption         =   "��Ʒ�����̣�"
      Height          =   180
      Left            =   225
      TabIndex        =   10
      Top             =   2370
      Width           =   1080
   End
   Begin VB.Label lbl����֧���� 
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
      Caption         =   "����֧���̣�"
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
      Caption         =   "ʹ��Ȩ���裺"
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
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��Ȩ����(C) ������Ϣ��ҵ��˾"
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
    MsgBox "��ʱϵͳ��Ϣ��Ч" & vbNewLine & err.Description, vbOKOnly, gstrSysName
End Sub

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
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
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ��ע���
    
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

Public Sub ShowAbout()
'���ܣ� ��ʾ���ڴ���

    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, objItem As ListItem
    Dim strKind As String, strCode As String
    Dim strSerial As String, strSQL As String
        
    On Error GoTo errH
    
    strKind = gobjRegister.zlRegInfo("��Ȩ����")
    If strKind = "2" Then
        strKind = "(����)"
    ElseIf strKind = "3" Then
        strKind = "(����)"
    Else
        strKind = ""
    End If
    With frmAbout
        .lblSysName.Caption = App.Title & strKind
        .lblVersion.Caption = App.ProductName & " Version " & App.Major & "." & App.Minor & "." & App.Revision
        .lblGrant.Caption = Replace(gobjRegister.zlRegInfo("��λ����", , -1), ";", vbCrLf)
        
        .lbl����֧����.Caption = gobjRegister.zlRegInfo("����֧����", , -1)
        Call ApplyOEM_Picture(imgLogo, "Picture")
        
        If Trim$(.lbl����֧����.Caption) = "" Then
            .Label1.Visible = False
            .lbl����֧����.Visible = False
            .lblCopyRight.Visible = False
        Else
            .Label1.Visible = True
            .lbl����֧����.Visible = True
            .lblCopyRight.Visible = True
        End If
        
        strCode = gobjRegister.zlRegInfo("��Ʒ������", , -1)
        If Trim(strCode) = "" Then
            .Label3.Visible = False
            .lbl������.Visible = False
        Else
            .Label3.Visible = True
            .lbl������.Visible = True
            .lbl������.Caption = ""
            For i = 0 To UBound(Split(strCode, ";"))
                .lbl������.Caption = .lbl������.Caption & Split(strCode, ";")(i) & vbCrLf
            Next
        End If
        
        '��ʾ�����߱���İ汾��
        strCode = gobjRegister.zlRegInfo("�汾��")
        If strCode = "" Then
            lbl������.Visible = False
        Else
            lbl������.Caption = "���������ݰ汾��" & strCode
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
