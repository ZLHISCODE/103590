VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFtpMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "zlPacsFtpTools(����FTP���Թ���)"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9780
   Icon            =   "frmFtpMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   9780
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdClear 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   64
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   27
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdDosMod 
      Caption         =   "DosMod"
      Height          =   375
      Left            =   6000
      TabIndex        =   51
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtCount 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "0"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "��¼������־"
      Height          =   375
      Left            =   5040
      TabIndex        =   49
      Top             =   2040
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkRoot 
      Caption         =   "ǿ��ʹ�ø�·��(����LINUX)"
      Height          =   375
      Left            =   6240
      TabIndex        =   48
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   35
      Top             =   2400
      Width           =   9495
      Begin VB.CheckBox chkSize 
         Caption         =   "1K"
         Height          =   180
         Index           =   0
         Left            =   2040
         TabIndex        =   40
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkSize 
         Caption         =   "512K"
         Height          =   180
         Index           =   1
         Left            =   3600
         TabIndex        =   39
         Top             =   240
         Width           =   930
      End
      Begin VB.CheckBox chkSize 
         Caption         =   "1M"
         Height          =   180
         Index           =   2
         Left            =   5250
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkSize 
         Caption         =   "5M"
         Height          =   180
         Index           =   3
         Left            =   6720
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkSize 
         Caption         =   "10M"
         Height          =   180
         Index           =   4
         Left            =   8160
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblTranLSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   4
         Left            =   8400
         TabIndex        =   63
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lblTranLSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   3
         Left            =   6960
         TabIndex        =   62
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lblTranLSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   2
         Left            =   5400
         TabIndex        =   61
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lblTranLSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   1
         Left            =   3840
         TabIndex        =   60
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lblTranLSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   0
         Left            =   2280
         TabIndex        =   59
         Top             =   555
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ƽ����ʱ(��¼)��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   58
         Top             =   555
         Width           =   1575
      End
      Begin VB.Label lblTranDSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   4
         Left            =   8400
         TabIndex        =   57
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label lblTranDSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   3
         Left            =   6960
         TabIndex        =   56
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label lblTranDSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   2
         Left            =   5400
         TabIndex        =   55
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label lblTranDSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   1
         Left            =   3840
         TabIndex        =   54
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ƽ����ʱ(����)��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   53
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblTranDSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   0
         Left            =   2280
         TabIndex        =   52
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "��֤�ļ���С��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "ƽ����ʱ(�ϴ�)��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   46
         Top             =   885
         Width           =   1575
      End
      Begin VB.Label lblTranSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   0
         Left            =   2280
         TabIndex        =   45
         Top             =   885
         Width           =   360
      End
      Begin VB.Label lblTranSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   1
         Left            =   3840
         TabIndex        =   44
         Top             =   885
         Width           =   360
      End
      Begin VB.Label lblTranSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   2
         Left            =   5400
         TabIndex        =   43
         Top             =   885
         Width           =   360
      End
      Begin VB.Label lblTranSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   3
         Left            =   6960
         TabIndex        =   42
         Top             =   885
         Width           =   360
      End
      Begin VB.Label lblTranSpeed 
         AutoSize        =   -1  'True
         Caption         =   "0 ms"
         Height          =   180
         Index           =   4
         Left            =   8400
         TabIndex        =   41
         Top             =   885
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "��֤"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   31
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdTracert 
      Caption         =   "Tracert"
      Height          =   375
      Left            =   5160
      TabIndex        =   30
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdFtp 
      Caption         =   "Ftp"
      Height          =   375
      Left            =   4440
      TabIndex        =   29
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "Ping"
      Height          =   375
      Left            =   3720
      TabIndex        =   28
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdDownTest 
      Caption         =   "��"
      Height          =   375
      Left            =   1500
      TabIndex        =   26
      Top             =   6045
      Width           =   255
   End
   Begin VB.CommandButton cmdUpTest 
      Caption         =   "��"
      Height          =   375
      Left            =   1500
      TabIndex        =   25
      Top             =   5565
      Width           =   255
   End
   Begin RichTextLib.RichTextBox rtbLog 
      CausesValidation=   0   'False
      Height          =   4815
      Left            =   1800
      TabIndex        =   24
      Top             =   4080
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8493
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmFtpMain.frx":43B2
   End
   Begin VB.CheckBox chkGetSizeTest 
      Caption         =   "��С��ȡ����"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   7530
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkDelFileTest 
      Caption         =   "�ļ�ɾ������"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   8025
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkDelDirTest 
      Caption         =   "Ŀ¼ɾ������"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   8520
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkGetListTest 
      Caption         =   "�б��ȡ����"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   7035
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkMoveTest 
      Caption         =   "�ļ��ƶ�����"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   6540
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkDownTest 
      Caption         =   "�ļ����ز���"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   6045
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkUpTest 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      MaskColor       =   &H8000000F&
      TabIndex        =   17
      Top             =   5565
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox chkChangeTest 
      Caption         =   "Ŀ¼�л�����"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   5070
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkMdkTest 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4575
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox chkLoginTest 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox chkConEvery 
      Caption         =   "ÿ�β��Զ�������"
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox chkPassive 
      Caption         =   "���ñ�������"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.CheckBox chkForcedRead 
      Caption         =   "����ǿ�ƶ�ȡ"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.CheckBox chkContinue 
      Caption         =   "��������"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtTestTimes 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Text            =   "1"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtFtpVirtual 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Text            =   "/"
      Top             =   1080
      Width           =   6615
   End
   Begin VB.TextBox txtFtpPassWord 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5400
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtFtpUser 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox txtFtpAdress 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label lblUpTest 
      AutoSize        =   -1  'True
      Caption         =   "�ļ��ϴ�����"
      Height          =   180
      Left            =   405
      TabIndex        =   34
      Top             =   5655
      Width           =   1080
   End
   Begin VB.Label lblMdkTest 
      AutoSize        =   -1  'True
      Caption         =   "Ŀ¼��������"
      Height          =   180
      Left            =   405
      TabIndex        =   33
      Top             =   4665
      Width           =   1080
   End
   Begin VB.Label lblLoginTest 
      AutoSize        =   -1  'True
      Caption         =   "�û���¼����"
      Height          =   180
      Left            =   405
      TabIndex        =   32
      Top             =   4170
      Width           =   1080
   End
   Begin VB.Label lblTestTimes 
      AutoSize        =   -1  'True
      Caption         =   "���Դ�����"
      Height          =   180
      Left            =   390
      TabIndex        =   8
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label lblFtpVirtual 
      AutoSize        =   -1  'True
      Caption         =   "FTP����Ŀ¼��"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1170
   End
   Begin VB.Label lblFtpPassWord 
      AutoSize        =   -1  'True
      Caption         =   "FTP���룺"
      Height          =   180
      Left            =   4560
      TabIndex        =   2
      Top             =   720
      Width           =   810
   End
   Begin VB.Label lblFtpUser 
      AutoSize        =   -1  'True
      Caption         =   "FTP�û�����"
      Height          =   180
      Left            =   300
      TabIndex        =   1
      Top             =   720
      Width           =   990
   End
   Begin VB.Label lblFtpAdress 
      AutoSize        =   -1  'True
      Caption         =   "FTP��ַ��"
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "frmFtpMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mftpNet As New clsFtp
Private WithEvents mfrmUpLoad As frmUpLoad
Attribute mfrmUpLoad.VB_VarHelpID = -1
Private WithEvents mfrmDownLoad As frmDownLoad
Attribute mfrmDownLoad.VB_VarHelpID = -1
Private marrErrInfo() As String
Private mblnEnd As Boolean
Private mstrFootPath As String
Private mlngPassive As Long
Private Const M_STR_FILENAME = "ZLFtpTest"

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private mobjCurDos As clsDos


Private Function LoginTest(strAdress As String, strUser As String, strPassWord As String, _
    ByRef lngTime As Long) As Boolean
'�û���¼����
    Dim lngResult As Long
    Dim strInfor As String
    Dim lngStart As Long
    
    If mftpNet.hConnection > 0 Then
        mftpNet.FuncFtpDisConnect
    End If
    
    lngStart = GetTickCount
    lngResult = mftpNet.FuncFtpConnect(strAdress, strUser, strPassWord, IIf(chkRoot.Value <> 0, True, False))
    
    lngTime = GetTickCount - lngStart
    
    If lngResult <> 0 Then
        mstrFootPath = mftpNet.GetFtpRootPath
        
        If Len(Trim(mstrFootPath)) > 1 Then
            mstrFootPath = Mid(mstrFootPath, 2)
        End If
        strInfor = "���û���¼���ԡ�--�û���¼���Գɹ���FTP��Ŀ¼Ϊ��""" & mstrFootPath & """"
    Else
        strInfor = "���û���¼���ԡ�--FTP����ʧ�ܣ�����FTP��ַ���û����������Ƿ���ȷ��FTP�����Ƿ������������Ƿ���ȷ��"
    End If
    
    PrintLog strInfor, IIf(lngResult <> 0, True, False)
    
    LoginTest = IIf(lngResult <> 0, True, False)
End Function

Private Function MkDirTest(strVirtualPath As String, strDir As String) As Boolean
'����Ŀ¼����
    Dim lngResult As Long
    Dim strInfo As String
    
    '�ж�����Ŀ¼�Ƿ����
    If GetFTPString(strVirtualPath) <> "/" Then
        If Not IsDirExsisted(strVirtualPath) Then
            lngResult = mftpNet.FuncFtpMkDir("", strVirtualPath)
            If lngResult <> 0 Then
                strInfo = "������Ŀ¼���ԡ�--����Ŀ¼ʧ�ܡ�����ԭ��Ŀ¼���Ʋ��Ϸ����û�Ȩ�޲��㡣"
                PrintLog strInfo, IIf(lngResult = 0, True, False)
                MkDirTest = IIf(lngResult = 0, True, False)
                Exit Function
            End If
        End If
    End If
    
    '��������Ĳ���Ŀ¼���ڣ�Ӧ��ɾ��
    If IsDirExsisted(strVirtualPath & "/" & strDir) Then
        If Len(mftpNet.FuncDirFiles(strVirtualPath & "/" & strDir)) > 0 Then
            DeleteFiles strVirtualPath & "/" & strDir
        End If
        mftpNet.FuncFtpDelDir strVirtualPath, strDir
    End If
    
    lngResult = mftpNet.FuncFtpMkDir(strVirtualPath, strDir)
    
    If lngResult = 0 Then
        strInfo = "������Ŀ¼���ԡ�--����Ŀ¼��" & GetFTPString(strVirtualPath & "/" & strDir) & "��������ɡ�"
    Else
        strInfo = "������Ŀ¼���ԡ�--FTPĿ¼��" & GetFTPString(strVirtualPath) & "���´�������Ŀ¼��" & GetFTPString(strDir) & "��ʱʧ�ܡ�����ԭ���Ѵ�����ͬĿ¼��FTP����Ŀ¼���Ʋ��Ϸ����û�Ȩ�޲��㡣"
    End If
    
    PrintLog strInfo, IIf(lngResult = 0, True, False)
    MkDirTest = IIf(lngResult = 0, True, False)
End Function

Private Function IsDirExsisted(strDir As String) As Boolean

    IsDirExsisted = False
    
    mftpNet.FuncChangeDir ""
    If mftpNet.FuncChangeDir(strDir) = 0 Then
        IsDirExsisted = True
    End If
End Function

Private Function ChangeTest(strDir As String) As Boolean
'Ŀ¼�л�����
    Dim lngResult As Long
    Dim strInfo As String
    
    mftpNet.FuncChangeDir ""
    lngResult = mftpNet.FuncChangeDir(strDir)
        
    If lngResult = 0 Then
        strInfo = "��Ŀ¼�л����ԡ�--�л�FTP����Ŀ¼����" & GetFTPString(mftpNet.GetFtpCWD) & "����ɡ�"
    Else
        strInfo = "��Ŀ¼�л����ԡ�--�л�FTP����Ŀ¼����" & GetFTPString(strDir) & "��ʱʧ�ܡ�����ԭ���л���Ŀ¼�����ڻ��û�Ȩ�޲��㡣"
        
    End If
    
    PrintLog strInfo, IIf(lngResult = 0, True, False)
    ChangeTest = IIf(lngResult = 0, True, False)
End Function

Private Function UpTest(strVirtualPath As String, strLocalFileName As String, _
    strRemoteFileName As String, ByRef lngTime As Long) As Boolean
'�ļ��ϴ�����
    Dim lngResult As Long
    Dim strInfo As String
    Dim lngStart As Long
    
    lngStart = GetTickCount
    
    lngResult = mftpNet.FuncUploadFile(strVirtualPath, strLocalFileName, strRemoteFileName)
    
    lngTime = GetTickCount - lngStart
    
    If lngResult = 0 Then
        strInfo = "���ļ��ϴ����ԡ�--���ز����ļ���" & strLocalFileName & "���ɹ��ϴ���FTPĿ¼��" & GetFTPString(mftpNet.GetFtpCWD) & "���¡�"
    Else
        strInfo = "���ļ��ϴ����ԡ�--���ز����ļ���" & strLocalFileName & "���ϴ���FTPĿ¼��" & GetFTPString(strVirtualPath) & "��ʱʧ�ܡ�����ԭ�򣺱��ز����ļ������ڡ��ϴ���FTPĿ¼�����ڡ�FTP�������õ��³�ʱ���ļ���������ϴ���С���û�Ȩ�޲��㡣"
    End If
    
    PrintLog strInfo, IIf(lngResult = 0, True, False)
    UpTest = IIf(lngResult = 0, True, False)
End Function


Private Function DownTest(strVirtualPath As String, strLocalFileName As String, _
    strRemoteFileName As String, ByRef lngTime As Long) As Boolean
'�ļ����ز���
    Dim lngResult As Long
    Dim strInfo As String
    Dim lngStart As Long
    
    lngStart = GetTickCount
    
    lngResult = mftpNet.FuncDownloadFile(strVirtualPath, strLocalFileName, strRemoteFileName, IIf(chkForcedRead.Value = 1, True, False))
    
    lngTime = GetTickCount - lngStart
    
    If lngResult = 0 Then
        strInfo = "���ļ����ز��ԡ�--FTP�����ļ���" & GetFTPString(mftpNet.GetFtpCWD & "/" & strRemoteFileName) & "���ɹ����ص����ء�" & strLocalFileName & "����"
    Else
        strInfo = "���ļ����ز��ԡ�--FTP�����ļ���" & GetFTPString(strVirtualPath & "/" & strRemoteFileName) & "�����ص����ء�" & strLocalFileName & "��ʱʧ�ܡ�����ԭ��FTP�����ļ������ڡ�����·�������ڡ�FTP�������õ��³�ʱ���û�Ȩ�޲��㡣"
    End If
    
    PrintLog strInfo, IIf(lngResult = 0, True, False)
    DownTest = IIf(lngResult = 0, True, False)
End Function

Private Function MoveTest(ByVal strSourceFile As String, ByVal strNewFile As String, ByVal strFileName As String) As Boolean
'�ļ��ƶ�����
    Dim lngResult As Long
    Dim strInfo As String
    
    lngResult = mftpNet.FuncReNameFile(strSourceFile & "/" & strFileName, strNewFile & "/" & strFileName)
    
    'Ϊ�������Ĳ��ԣ��ƶ��ļ�֮���軹ԭ
    If lngResult = 0 Then
        lngResult = mftpNet.FuncReNameFile(strNewFile & "/" & strFileName, strSourceFile & "/" & strFileName)
        
        If lngResult <> 0 Then
            strInfo = "���ļ��ƶ����ԡ�--FTP�����ļ���" & GetFTPString(mftpNet.GetFtpCWD & strFileName) & "���Ƶ�FTPĿ¼��" & GetFTPString(strSourceFile) & "��ʱʧ�ܡ�����ԭ��FTP�����ļ������ڡ�FTPĿ¼�����ڻ��û�Ȩ�޲��㡣"
        Else
            strInfo = "���ļ��ƶ����ԡ�--FTP�����ļ���" & GetFTPString(mftpNet.GetFtpCWD & strFileName) & "���Ƶ�FTPĿ¼��" & GetFTPString(strNewFile) & "���ɹ�,���ѻָ����ļ��ƶ�ǰ״̬��"
        End If
    Else
        strInfo = "���ļ��ƶ����ԡ�--FTP�����ļ���" & GetFTPString(strSourceFile & strFileName) & "���Ƶ�FTPĿ¼��" & GetFTPString(strNewFile) & "��ʱʧ�ܡ�����ԭ��FTP�����ļ������ڡ�FTPĿ¼�����ڻ��û�Ȩ�޲��㡣"
    End If
    
    PrintLog strInfo, IIf(lngResult = 0, True, False)
    MoveTest = IIf(lngResult = 0, True, False)
End Function

Private Function GetListTest(ByVal strVirtualPath As String) As Boolean
'��ȡ�б����
    Dim strResult As String
    Dim strInfo As String
    Dim arrFile() As String
    Dim strFile As String
    Dim i As Long
    
    strResult = mftpNet.FuncDirFiles(strVirtualPath)
    
    If Len(strResult) > 0 Then
        arrFile = Split(strResult, "|")
        For i = 0 To UBound(arrFile)
            strFile = strFile & "��" & arrFile(i) & "��"
        Next
        strInfo = "����ȡ�б���ԡ�--�ɹ���ȡ��FTPĿ¼��" & GetFTPString(mftpNet.GetFtpCWD) & "���µ��ļ���" & strFile & "��"
    Else
        strInfo = "����ȡ�б���ԡ�--��ȡFTPĿ¼��" & GetFTPString(mftpNet.GetFtpCWD) & "���µ��ļ�ʱʧ�ܡ�����ԭ��FTPĿ¼�����ڡ���Ŀ¼��û���ļ����û�Ȩ�޲��㡣"
    End If
    
    PrintLog strInfo, IIf(Len(strResult) > 0, True, False)
    GetListTest = IIf(Len(strResult) > 0, True, False)
End Function

Private Function GetSizeTest(ByVal strVirtualPath As String, ByVal strFile As String) As Boolean
'��С��ȡ����
    Dim lngResult As Long
    Dim strInfo As String
    
    lngResult = mftpNet.FuncFtpGetFileSize(strVirtualPath, strFile)
    
    If lngResult > 0 Then
        strInfo = "����С��ȡ���ԡ�--�ɹ���ȡ��FTPĿ¼��" & GetFTPString(mftpNet.GetFtpCWD) & "���µġ�" & strFile & "���ļ��Ĵ�СΪ��" & lngResult & "���ֽڡ�"
    Else
        strInfo = "����С��ȡ���ԡ�--��ȡ��FTPĿ¼��" & GetFTPString(strVirtualPath) & "���µġ�" & strFile & "���ļ��Ĵ�Сʱʧ�ܡ�����ԭ��FTP�����ļ������ڡ��ļ���СΪ0���û�Ȩ�޲��㡣"
    End If
    
    PrintLog strInfo, IIf(lngResult > 0, True, False)
    GetSizeTest = IIf(lngResult > 0, True, False)
End Function

Private Function DelFileTest(strVirtualPath As String, strFileName As String) As Boolean
'�ļ�ɾ������
    Dim lngResult As Long
    Dim strInfo As String
    
    lngResult = mftpNet.FuncDelFile(strVirtualPath, strFileName)
        
    If lngResult = 0 Then
        strInfo = "���ļ�ɾ�����ԡ�--�ɹ�ɾ��FTPĿ¼��" & GetFTPString(mftpNet.GetFtpCWD) & "���µ��ļ���" & strFileName & "����"
    Else
        strInfo = "���ļ�ɾ�����ԡ�--ɾ��FTPĿ¼��" & GetFTPString(strVirtualPath) & "���µ��ļ���" & strFileName & "��ʱʧ�ܡ�����ԭ��FTP�����ļ������ڻ��û�Ȩ�޲���"
    End If
    
    PrintLog strInfo, IIf(lngResult = 0, True, False)
    DelFileTest = IIf(lngResult = 0, True, False)
End Function

Private Function DelDirTest(strVirtualPath As String, strDir As String) As Boolean
'Ŀ¼ɾ������
    Dim lngResult As Long
    Dim strInfo As String
    
    If Len(mftpNet.FuncDirFiles(strVirtualPath & "/" & strDir)) > 0 Then
        DeleteFiles strVirtualPath & "/" & strDir
    End If
    
    lngResult = mftpNet.FuncFtpDelDir(strVirtualPath, strDir)
    
    If lngResult = 0 Then
        strInfo = "��Ŀ¼ɾ�����ԡ�--�ɹ�ɾ��FTPĿ¼��" & GetFTPString(strVirtualPath) & "���µ�Ŀ¼��" & GetFTPString(strDir) & "����"
    Else
        strInfo = "��Ŀ¼ɾ�����ԡ�--ɾ��FTPĿ¼��" & GetFTPString(strVirtualPath) & "���µ�Ŀ¼��" & GetFTPString(strDir) & "��ʱʧ�ܡ�����ԭ��FTPĿ¼�����ڡ���Ŀ¼�´����ļ����û�Ȩ�޲��㡣"
    End If
    
    PrintLog strInfo, IIf(lngResult = 0, True, False)
    DelDirTest = IIf(lngResult = 0, True, False)
End Function

Private Sub chkPassive_Click()
    On Error GoTo errHandle
    
    Call SaveSetting("ZLSOFT", "����ģ��\Ftp", "���ñ�������", IIf(chkPassive.Value, 1, 0))
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub ActiveCmd(objShell As WshShell, Optional ByVal lngSleepTime As Long = 20)
    Call objShell.AppActivate("cmd.exe", True)
    'Sleep lngSleepTime
End Sub

Private Sub SendKeysEx(objShell As WshShell, strCmd As String, Optional ByVal lngSleepTime As Long = 30)
On Error Resume Next
    Call ActiveCmd(objShell, lngSleepTime)
    Call objShell.SendKeys(strCmd, True)
Err.Clear
End Sub

Public Function GetLoginCfg(ByVal strDir As String, ByVal strUser As String, ByVal strPwd As String) As String
    Dim lngFileNum As Long
    Dim strFilePath As String
    Dim objFSO As Object
    Dim objLogFile As Object
    Dim strInfo As String
    Dim strLine As String
    
On Error GoTo errHandle

    GetLoginCfg = ""
    
    strFilePath = FormatPath(App.Path & IIf(strDir = "", "", "\" & strDir & "\") & "ftpcfg.dat")

    DFile strFilePath
    
    If Len(Dir(strFilePath)) = 0 Then
        strInfo = strUser & vbCrLf & strPwd
        lngFileNum = FreeFile
        
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        If Len(Dir(strFilePath)) = 0 Then
            objFSO.CreateTextFile strFilePath, True
        End If
        
        Set objLogFile = objFSO.GetFile(strFilePath)
        
        Open Trim(strFilePath) For Output As #lngFileNum
        Print #lngFileNum, strInfo
        Close #lngFileNum
    End If

    GetLoginCfg = strFilePath
Exit Function
errHandle:

End Function

Private Function CheckState(ByVal strResult As String) As Boolean
    Dim strLCase As String
    
    strLCase = LCase(strResult)
    
    CheckState = False
    
    If InStr(strResult, "530") > 0 Or InStr(strLCase, "not logged in") > 0 Then
        MsgBox "��¼ʧ�ܣ����ܼ������ԡ�"
        Exit Function
    End If
    
    'Connection closed
    If InStr(strResult, "426") > 0 Or InStr(strLCase, "connection closed") > 0 Then
        MsgBox "���ӱ��رգ�����Զ�˷������Ƿ�����"
        Exit Function
    End If
    
    If InStr(strResult, "451") > 0 Or InStr(strLCase, "error") > 0 Then
        MsgBox "���������쳣��"
        Exit Function
    End If
    
    
    CheckState = True
End Function

Private Sub CmdTest(ByVal strIp As String, ByVal strVPath As String, _
    ByVal strLoginFile As String, _
    ByVal strLocalDir As String, _
    ByVal strLocalFile As String, _
    ByVal strFtpDir As String, _
    ByVal strFtpFile As String)
On Error GoTo errHandle
    
    Dim strResult As String
    
    rtbLog.Text = ""

    
    Call PrintNullLine(1)
    Call PrintLog("ִ������(FTP��¼)��ftp " & strIp)
    Call mobjCurDos.DosInput("ftp -s:" & strLoginFile & " " & strIp)
    
    strResult = mobjCurDos.DosOutPutEx(, 200)
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd
    
    Call PrintNullLine(1)
    Call PrintLog("ִ�����trace")
    Call mobjCurDos.DosInput("trace")
    strResult = mobjCurDos.DosOutPutEx()
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd

    Call PrintNullLine(1)
    Call PrintLog("ִ�����debug")
    Call mobjCurDos.DosInput("debug")
    strResult = mobjCurDos.DosOutPutEx()
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd
    
    '�л�����·��
    Call PrintNullLine(1)
    Call PrintLog("ִ�����cd /")
    Call mobjCurDos.DosInput("cd /")
    strResult = mobjCurDos.DosOutPutEx()
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd

    If strVPath <> "" Then
        '��������Ŀ¼
        Call PrintLog("ִ������(����Ŀ¼)��cd " & strVPath, , False)
        Call mobjCurDos.DosInput("cd " & strVPath)
        strResult = mobjCurDos.DosOutPutEx()
        Call PrintLog(strResult, , False)
        
        If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd
    End If

     
    'Ŀ¼����
    Call PrintNullLine(1)
    Call PrintLog("ִ������(Ŀ¼����)��mkdir " & strFtpDir)
    Call mobjCurDos.DosInput("mkdir " & strFtpDir)
    strResult = mobjCurDos.DosOutPutEx()
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd

    'Ŀ¼�л�����
    Call PrintNullLine(1)
    Call PrintLog("ִ������(Ŀ¼�л�)��cd " & strFtpDir)
    Call mobjCurDos.DosInput("cd " & strFtpDir)
    strResult = mobjCurDos.DosOutPutEx()
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd
    
    '�ļ��ϴ����� strLocalDir & "\" &
    Call PrintNullLine(1)
    Call PrintLog("ִ������(�ļ��ϴ�)��put " & strLocalFile & " " & strFtpFile)
    Call mobjCurDos.DosInput("put " & strLocalFile & " " & strFtpFile)
    strResult = mobjCurDos.DosOutPutEx(, 500)
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd
    
    '�ļ����ز���
    Call PrintNullLine(1)
    Call PrintLog("ִ������(�ļ�����)��get " & strFtpFile & " " & strLocalDir & "\" & strFtpFile & "_Down")
    Call mobjCurDos.DosInput("get " & strFtpFile & " " & strLocalDir & "\" & strFtpFile & "_Down")
    strResult = mobjCurDos.DosOutPutEx(, 500)
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd
    
    '�ļ��ƶ�����
    Call PrintNullLine(1)
    Call PrintLog("ִ������(�ļ��ƶ�)��rename " & strFtpFile & " /" & strFtpFile & "_rename")
    Call mobjCurDos.DosInput("rename " & strFtpFile & " /" & strFtpFile & "_rename")
    strResult = mobjCurDos.DosOutPutEx()
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd
    
    Call mobjCurDos.DosInput("rename /" & strFtpFile & "_rename" & " " & strFtpFile)
    strResult = mobjCurDos.DosOutPutEx()
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd
    
    '�ļ��б����
    Call PrintNullLine(1)
    Call PrintLog("ִ������(�ļ�ö��)��dir")
    Call mobjCurDos.DosInput("dir")
    strResult = mobjCurDos.DosOutPutEx(, 200)
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd
    
    '�ļ���С����
    Call PrintNullLine(1)
    Call PrintLog("ִ������(�ļ���С)��ls " & strFtpFile)
    Call mobjCurDos.DosInput("ls " & strFtpFile)
    strResult = mobjCurDos.DosOutPutEx()
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd
  
    '�ļ�ɾ������
    Call PrintNullLine(1)
    Call PrintLog("ִ������(�ļ�ɾ��)��Delete " & strFtpFile)
    Call mobjCurDos.DosInput("Delete " & strFtpFile)
    strResult = mobjCurDos.DosOutPutEx()
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd
    
    '�ָ������Ŀ¼
    Call PrintNullLine(1)
    Call PrintLog("ִ�����cd /", , False)
    Call mobjCurDos.DosInput("cd /")
    strResult = mobjCurDos.DosOutPutEx()
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd

    If strVPath <> "" Then
        '��������Ŀ¼
        Call PrintLog("ִ������(����Ŀ¼)��cd " & strVPath, , False)
        Call mobjCurDos.DosInput("cd " & strVPath)
        strResult = mobjCurDos.DosOutPutEx()
        Call PrintLog(strResult, , False)
        
        If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd
    End If
    
    'Ŀ¼ɾ������
    Call PrintNullLine(1)
    Call PrintLog("ִ������(Ŀ¼ɾ��)��RmDir " & strFtpDir)
    Call mobjCurDos.DosInput("RmDir " & strFtpDir)
    strResult = mobjCurDos.DosOutPutEx()
    Call PrintLog(strResult, , False)
    
    If CheckState(strResult) = False Or mblnEnd Then GoTo abortCmd
    
    
abortCmd:
    Call PrintNullLine(1)
    Call PrintLog("ִ�����close")
    Call mobjCurDos.DosInput("close")
    Call PrintLog(mobjCurDos.DosOutPutEx(), , False)
     
    
    Call PrintLog("ִ�����quit")
    Call mobjCurDos.DosInput("quit")
    Call PrintLog(mobjCurDos.DosOutPutEx(), , False)
     
    
'    Call mobjCurDos.DosInput("exit")
'    Call PrintLog(mobjCurDos.DosOutPutEx(), , False)
    
    Call PrintLog("ִ�����taskkill")
    Call mobjCurDos.DosInput("taskkill /F /IM ftp.exe")
    Call PrintLog(mobjCurDos.DosOutPutEx(), , False)
     
    
Exit Sub
errHandle:
    MsgBox Err.Description
End Sub


Private Sub cmdClear_Click()
On Error GoTo errHandle
    
    Call ClearLocalFile

    Call ClearFtpFile
    
    MsgBox "FTP����Ŀ¼���ļ��������."
Exit Sub
errHandle:
    MsgBox Err.Description
End Sub

Private Sub ClearLocalFile()
On Error Resume Next
    Kill FormatPath(App.Path & "\DosMod\*.*")
    Kill FormatPath(App.Path & "\Download\*.*")
    Kill FormatPath(App.Path & "\Upload\*.*")
    
    Err.Clear
End Sub

Private Sub ClearFtpFile()
    Dim strDirs As String
    Dim aryDir() As String
    Dim strVirtualPath As String
    Dim i As Long
    
    Call mftpNet.FuncFtpDisConnect
    
    Call mftpNet.FuncFtpConnect(Trim(txtFtpAdress.Text), Trim(txtFtpUser.Text), Trim(txtFtpPassWord.Text), IIf(chkRoot.Value <> 0, True, False))
    
    
    'Call mftpNet.FuncChangeDir(GetFTPString(txtFtpVirtual.Text))
    
    strVirtualPath = GetFTPString(txtFtpVirtual.Text)
    strDirs = mftpNet.FuncDirFiles(strVirtualPath, "ZLFtpTest_*")
    aryDir = Split(strDirs & "|", "|")
    
    For i = 0 To UBound(aryDir) - 1
        If aryDir(i) <> "" Then
            DeleteFiles strVirtualPath & "/" & aryDir(i)
            mftpNet.FuncFtpDelDir strVirtualPath, aryDir(i)
        End If
    Next
    
    Call mftpNet.FuncFtpDisConnect
    
End Sub

Private Sub cmdDosMod_Click()
On Error GoTo errHandle
    Dim strRandomDir As String
    Dim strRandomFile As String
    Dim strLocalDir As String
    Dim strLocalFile As String
    Dim strVPath As String
    Dim strFtpLoginFile As String
    Dim lngTestTimes As Long
    Dim i As Long
    Dim j As Long

    '���û���
    If Not mobjCurDos Is Nothing Then
        MsgBox "����ִ�У����Ժ������"
        Exit Sub
    End If
    
    mblnEnd = False
    
    lngTestTimes = Val(txtTestTimes.Text)
    strVPath = ""
    If Trim(txtFtpVirtual.Text) <> "" And Trim(txtFtpVirtual.Text) <> "/" Then
        strVPath = txtFtpVirtual.Text
    End If
    
    strFtpLoginFile = GetLoginCfg("DosMod", txtFtpUser.Text, txtFtpPassWord.Text)
    strRandomDir = "DosModDir_" & Format(Now, "mmddhhmmss_") & GetTickCount
    strRandomFile = "DosModFile_" & Format(Now, "mmddhhmmss_") & GetTickCount
    
    strLocalFile = strRandomFile
    strLocalDir = FormatPath(App.Path & "\DosMod\")
    If Dir(strLocalDir, vbDirectory) = "" Then
        Call MkDir(strLocalDir)
    End If
    
    '���������ļ�
    For i = 0 To 4
        If chkSize(i).Value <> 0 Then
            strLocalFile = GetTestFile(i, "DosMod", strLocalFile)
            
            Exit For
        End If
    Next
    
    Set mobjCurDos = New clsDos
    
    Call CmdTest(txtFtpAdress.Text, txtFtpVirtual.Text, _
        strFtpLoginFile, strLocalDir, strLocalFile, strRandomDir, strRandomFile)
    
    Set mobjCurDos = Nothing
    
    DFile strLocalFile
    DFile strLocalFile & "_Down"
    DFile strFtpLoginFile
    
Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub DFile(ByVal strFile As String)
On Error Resume Next
    Kill strFile
Err.Clear
End Sub

Private Sub cmdDownTest_Click()
    On Error GoTo errHandle
    
    Set mfrmDownLoad = New frmDownLoad
    
    mfrmDownLoad.Show 1, Me
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdEnd_Click()
On Error GoTo errHandle
    
    mblnEnd = True
    cmdVerify.Enabled = True
    
    If Not mobjCurDos Is Nothing Then
        Call mobjCurDos.Abort
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdFtp_Click()
    On Error GoTo errHandle

    
    Call Shell("cmd /k ftp " & Trim(txtFtpAdress.Text), vbNormalFocus)
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdPing_Click()
    On Error GoTo errHandle
    
    Shell "cmd /k ping " & Trim(txtFtpAdress.Text), vbNormalFocus
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdTracert_Click()
    On Error GoTo errHandle
    
    Shell "cmd /k tracert  " & Trim(txtFtpAdress.Text), vbNormalFocus
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdUpTest_Click()
    On Error GoTo errHandle
    
    Set mfrmUpLoad = New frmUpLoad
    
    mfrmUpLoad.Show 1, Me
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdVerify_Click()
    On Error GoTo errHandle

    Call Test
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Public Sub Test()
    Dim strAdress As String
    Dim strUser As String
    Dim strPassWord As String
    Dim strVirtual As String
    
    Dim strTestDir As String
    Dim strFileDir As String
    Dim strFileName As String
    Dim arrFileName() As String
    
    Dim i As Long
    Dim j As Long
    
    Dim blnError As Boolean
    Dim blnDelDir As Boolean
    Dim blnResult As Boolean
    
    Dim lngTestTimes As Long
    
    Dim lngUpSpeed As Long
    Dim lngDownSpeed As Long
    Dim lngLoginSpeed As Long
    Dim lngTime As Long
    
    cmdVerify.Enabled = False
'    cmdEnd.Enabled = True
    
    If Not VerifyCheck Then
        cmdVerify.Enabled = True
'        cmdEnd.Enabled = False
        Exit Sub
    End If
    
    mblnEnd = False
    
    strAdress = Trim(txtFtpAdress.Text)
    strUser = Trim(txtFtpUser.Text)
    strPassWord = Trim(txtFtpPassWord.Text)
    lngTestTimes = Val(txtTestTimes.Text)
    lngLoginSpeed = 0
    
    strVirtual = GetFTPString(txtFtpVirtual.Text)
    
    Call ClearFace
    
    blnResult = True
    cmdVerify.Caption = "��֤��..."
    '��¼����
    If Not chkConEvery.Value = 1 Then
        chkLoginTest.ForeColor = &H80000012
        If LoginTest(strAdress, strUser, strPassWord, lngTime) Then
            PrintTag 1
            lblLoginTest.ForeColor = &H8000&
        Else
            lblLoginTest.ForeColor = vbRed
            cmdVerify.Enabled = True
'            cmdEnd.Enabled = False
            
            cmdVerify.Caption = "��֤"
            
            blnResult = False
            Call PrintResult(blnResult)
            Exit Sub
        End If
        
        lngLoginSpeed = lngLoginSpeed + lngTime
    End If
    
    
    For i = 0 To 4
        If mblnEnd Then Exit For
            
        If chkSize(i).Value = 1 Then
            lngUpSpeed = 0
            lngDownSpeed = 0
            
            If chkConEvery.Value <> 0 Then lngLoginSpeed = 0
            
            Call PrintNullLine(3)
            PrintTag 2
            PrintLog "��" & GetSizeInfo(i) & "���ļ���ʼ����..."
            PrintTag 2
            
            strFileDir = GetTestFile(i)
            
                    
            For j = 1 To lngTestTimes
                            
                txtCount.Text = j
                strTestDir = M_STR_FILENAME & "_" & GetTickCount
                
                If mblnEnd Then Exit For
                
                If lngTestTimes > 1 Then
                    PrintTag 4
                    PrintLog GetSizeInfo(i) & "�ļ���" & j & "�β��Կ�ʼ..."
                    PrintTag 4
                End If
                
                Call InitChkColor

                blnDelDir = False
                blnError = False
                     
                DoEvents
                'ÿ�β��Զ����½
                If mblnEnd Then Exit For
                If chkConEvery.Value = 1 Then
                    chkLoginTest.ForeColor = &H80000012
                    If LoginTest(strAdress, strUser, strPassWord, lngTime) Then
                        lblLoginTest.ForeColor = &H8000&
                    Else
                        blnError = True
                        lblLoginTest.ForeColor = vbRed
                        chkSize(i).ForeColor = vbRed
                        cmdVerify.Enabled = True
'                        cmdEnd.Enabled = False
                        cmdVerify.Caption = "��֤"
                        
                        blnResult = False
                        Call PrintResult(blnResult)
                        Exit Sub
                    End If
                    
                    lngLoginSpeed = lngLoginSpeed + lngTime
                    
                End If
                    
                DoEvents
                'Ŀ¼��������
                If mblnEnd Then Exit For
                If chkMdkTest.Value = 1 Then
                    Call PrintNullLine(1)
                    PrintLog "Ŀ¼����������..."
                    If MkDirTest(strVirtual, strTestDir) Then
                        lblMdkTest.ForeColor = &H8000&
                    Else
                        lblMdkTest.ForeColor = vbRed
                        blnError = True
                        blnResult = False
                        
                        chkSize(i).ForeColor = vbRed
                        If chkContinue.Value = 0 Then Exit For
                    End If
                End If
                    
                DoEvents
                'Ŀ¼�л�����
                If mblnEnd Then Exit For
                If chkChangeTest.Value = 1 Then
                    Call PrintNullLine(1)
                    PrintLog "Ŀ¼�л�������..."
                    If ChangeTest(strVirtual & "/" & strTestDir) Then
                        chkChangeTest.ForeColor = &H8000&
                    Else
                        chkChangeTest.ForeColor = vbRed
                        chkSize(i).ForeColor = vbRed
                        blnError = True
                        blnResult = False
                        If chkContinue.Value = 0 Then Exit For
                    End If
                End If
                
                
                arrFileName = Split(strFileDir, "\")
                strFileName = arrFileName(UBound(arrFileName))
    
                DoEvents
                '�ļ��ϴ�����
                If mblnEnd Then Exit For
                If chkUpTest.Value = 1 Then
                    Call PrintNullLine(1)
                    PrintLog "�ļ��ϴ�������..."
                    If UpTest(strVirtual & "/" & strTestDir, strFileDir, strFileName, lngTime) Then
                        lblUpTest.ForeColor = &H8000&
                    Else
                        lblUpTest.ForeColor = vbRed
                        chkSize(i).ForeColor = vbRed
                        blnError = True
                        blnResult = False
                        If chkContinue.Value = 0 Then Exit For
                    End If
                    
                    lngUpSpeed = lngUpSpeed + lngTime

                End If
                    
                DoEvents
                '�ļ����ز���
                If mblnEnd Then Exit For
                If chkDownTest.Value = 1 Then
                    Call PrintNullLine(1)
                    PrintLog "�ļ����ز�����..."
                 
                    blnResult = DownTest(strVirtual & "/" & strTestDir, FormatPath(App.Path & "\Download\" & M_STR_FILENAME & "_" & GetTickCount & ".txt"), strFileName, lngTime)

                    If blnResult Then
                        chkDownTest.ForeColor = &H8000&
                    Else
                        chkDownTest.ForeColor = vbRed
                        chkSize(i).ForeColor = vbRed
                        blnError = True
                        If chkContinue.Value = 0 Then Exit For
                    End If
                    
                    lngDownSpeed = lngDownSpeed + lngTime
                End If
                    
                DoEvents
                '�ļ��ƶ�����
                If mblnEnd Then Exit For
                If chkMoveTest.Value = 1 Then
                    Call PrintNullLine(1)
                    PrintLog "�ļ��ƶ�������..."
                    If MoveTest(strVirtual & "/" & strTestDir, strVirtual, strFileName) Then
                        chkMoveTest.ForeColor = &H8000&
                    Else
                        chkMoveTest.ForeColor = vbRed
                        chkSize(i).ForeColor = vbRed
                        blnError = True
                        blnResult = False
                        If chkContinue.Value = 0 Then Exit For
                    End If

                End If
                    
                DoEvents
                '�б��ȡ����
                If mblnEnd Then Exit For
                If chkGetListTest.Value = 1 Then
                    Call PrintNullLine(1)
                    PrintLog "�б��ȡ������..."
                    If GetListTest(strVirtual & "/" & strTestDir) Then
                        chkGetListTest.ForeColor = &H8000&
                    Else
                        chkGetListTest.ForeColor = vbRed
                        chkSize(i).ForeColor = vbRed
                        blnError = True
                        blnResult = False
                        If chkContinue.Value = 0 Then Exit For
                    End If

                End If
                    
                DoEvents
                '��ȡ��С����
                If mblnEnd Then Exit For
                If chkGetSizeTest.Value = 1 Then
                    Call PrintNullLine(1)
                    PrintLog "��С��ȡ������..."
                    If GetSizeTest(strVirtual & "/" & strTestDir, strFileName) Then
                        chkGetSizeTest.ForeColor = &H8000&
                    Else
                        chkGetSizeTest.ForeColor = vbRed
                        chkSize(i).ForeColor = vbRed
                        blnError = True
                        blnResult = False
                        If chkContinue.Value = 0 Then Exit For
                    End If

                End If
    
                DoEvents
                '�ļ�ɾ������
                If mblnEnd Then Exit For
                If chkDelFileTest.Value = 1 Then
                    Call PrintNullLine(1)
                    PrintLog "�ļ�ɾ��������..."
                    If DelFileTest(strVirtual & "/" & strTestDir, strFileName) Then
                        chkDelFileTest.ForeColor = &H8000&
                    Else
                        chkDelFileTest.ForeColor = vbRed
                        chkSize(i).ForeColor = vbRed
                        blnError = True
                        blnResult = False
                        If chkContinue.Value = 0 Then Exit For
                    End If

                End If
    
                DoEvents
                'ɾ��Ŀ¼����
                If mblnEnd Then Exit For
                If chkDelDirTest.Value = 1 Then
                    Call PrintNullLine(1)
                    PrintLog "Ŀ¼ɾ��������..."
                    If DelDirTest(strVirtual, strTestDir) Then
                        chkDelDirTest.ForeColor = &H8000&
                        blnDelDir = True
                    Else
                        chkDelDirTest.ForeColor = vbRed
                        chkSize(i).ForeColor = vbRed
                        blnError = True
                        blnResult = False
                        If chkContinue.Value = 0 Then Exit For
                    End If
                End If
    
                DoEvents
                If mblnEnd Then Exit For
                If Not blnError And chkSize(i).ForeColor <> vbRed Then
                    chkSize(i).ForeColor = &H8000&
                End If
        
                If mblnEnd Then Exit For
                
                If lngTestTimes > 1 Then
                    PrintTag 4
                    PrintLog GetSizeInfo(i) & "�ļ���" & j & "�β�����ɡ�"
                    PrintTag 4
                End If
                
                lblTranSpeed(i).Caption = Format(lngUpSpeed / j, "0.00") & " ms"
                lblTranDSpeed(i).Caption = Format(lngDownSpeed / j, "0.00") & " ms"
                
                If chkConEvery.Value = 1 Then
                    lblTranLSpeed(i).Caption = Format(lngLoginSpeed / j, "0.00") & " ms"
                Else
                    lblTranLSpeed(i).Caption = Format(lngLoginSpeed, "0.00") & " ms"
                End If
                
                If chkConEvery.Value = 1 Then
                    mftpNet.FuncFtpDisConnect
                End If
              
            Next
            
            If mblnEnd Then Exit For
            If chkContinue.Value = 0 And blnError Then Exit For
            
            PrintTag 2
            PrintLog "��" & GetSizeInfo(i) & "���ļ�������ɡ�"
            PrintTag 2
      
        
            Call ChangeLocation
        End If
    Next
    
    mftpNet.FuncFtpDisConnect
    
    Call PrintResult(blnResult)
    
    cmdVerify.Enabled = True
'    cmdEnd.Enabled = False
    
    cmdVerify.Caption = "��֤"
End Sub

Private Function VerifyCheck() As Boolean
    Dim i As Long
    
    VerifyCheck = False
    
    If Len(Trim(txtFtpAdress.Text)) = 0 Then
        MsgBox "����������Ե�FTP��ַ��", vbInformation, Me.Caption
        txtFtpAdress.SetFocus
        Exit Function
    End If
    
    If Len(Trim(txtFtpUser.Text)) = 0 Then
        MsgBox "������FTP�û�����", vbInformation, Me.Caption
        txtFtpUser.SetFocus
        Exit Function
    End If

    If Len(Trim(txtFtpPassWord.Text)) = 0 Then
        MsgBox "���������롣", vbInformation, Me.Caption
        txtFtpPassWord.SetFocus
        Exit Function
    End If
    
    If Val(txtTestTimes.Text) <= 0 Then
        MsgBox "��������ȷ�Ĳ��Դ�����", vbInformation, Me.Caption
        txtTestTimes.SetFocus
        Exit Function
    End If
    
    i = 0
    While chkSize(i) <> 1
        If i = 4 Then
            MsgBox "��ѡ����Ҫ���Ե��ļ����͡�", vbInformation, Me.Caption
            Exit Function
        End If
        i = i + 1
    Wend
    
    VerifyCheck = True
End Function

Private Function GetFTPString(strDir As String) As String
'��FTP·��ǰ��"/"���д���
        
    If Mid(Trim(strDir), 1, 1) = "/" Then
        GetFTPString = Trim(strDir)
    Else
        GetFTPString = "/" & Trim(strDir)
    End If
    
    GetFTPString = Replace(GetFTPString, "\", "/")
    GetFTPString = Replace(GetFTPString, "///", "/")
    GetFTPString = Replace(GetFTPString, "//", "/")
End Function

Private Sub ClearFace()
    Dim i As Long

    For i = 0 To 4
        lblTranSpeed(i).Caption = "0 ms"
    Next
    Call ChangeLocation
    
    Call InitChkColor
    Call InitLblColor

    rtbLog.Text = ""
    rtbLog.SelStart = 0
End Sub

Private Sub InitChkColor()
    If chkConEvery.Value = 1 Then
        lblLoginTest.ForeColor = &H80000012
    End If
    lblMdkTest.ForeColor = &H80000012
    chkChangeTest.ForeColor = &H80000012
    lblUpTest.ForeColor = &H80000012
    chkDownTest.ForeColor = &H80000012
    chkMoveTest.ForeColor = &H80000012
    chkGetListTest.ForeColor = &H80000012
    chkGetSizeTest.ForeColor = &H80000012
    chkDelFileTest.ForeColor = &H80000012
    chkDelDirTest.ForeColor = &H80000012
End Sub

Private Sub InitLblColor()
    Dim i As Long

    For i = 0 To 4
        chkSize(i).ForeColor = &H80000012
    Next
End Sub

Private Function GetSizeInfo(lngSize As Long) As String
    
    Select Case lngSize
        Case 0
            GetSizeInfo = "1K"
        Case 1
            GetSizeInfo = "512K"
        Case 2
            GetSizeInfo = "1M"
        Case 3
            GetSizeInfo = "5M"
        Case 4
            GetSizeInfo = "10M"
    End Select
    
End Function

'Private Sub PrintTestLog(strItem As String, blnResult As Boolean, Optional strInformation As String)
'    Dim strCurDate As String
'
'    strCurDate = Now() & " " & (Timer() * 1000) Mod 1000
'
''    If blnResult Then
''        rtbLog.SelText = "��" & strCurDate & ">>>" & strItem & "���Գɹ���" & vbCrLf
''    Else
''        rtbLog.SelText = "��" & strCurDate & ">>>" & strItem & "����ʧ�ܣ�" & Error(Err.LastDllError) & vbCrLf
''
''        ChangeColor Len(rtbLog.Text) - Len("��" & strCurDate & ">>>" & strItem & "����ʧ�ܣ�") - 2, Len("��" & strCurDate & ">>>" & strItem & "����ʧ�ܣ�")
''    End If
'
'    rtbLog.SelStart = Len(rtbLog.Text)
'
'    If Len(strInformation) > 0 Then
'        rtbLog.SelText = "��" & strCurDate & ">>>" & strInformation & vbCrLf
'        If Not blnResult Then
'            ChangeColor Len(rtbLog.Text) - Len("��" & strCurDate & ">>>" & strInformation) - 2, Len("��" & strCurDate & ">>>" & strInformation)
'        End If
'    End If
'
'    rtbLog.SelStart = Len(rtbLog.Text)
'
'End Sub


Private Sub PrintLog(ByVal strInfo As String, Optional ByVal blnResult As Boolean = True, _
    Optional ByVal blnHasPrefix As Boolean = True)
    Dim strCurDate As String
    Dim strLogText As String
    
    If chkLog.Value = 0 Then Exit Sub
    
    strCurDate = Now() & " " & (Timer() * 1000) Mod 1000
    strLogText = IIf(blnHasPrefix, "��", "") & strCurDate & ">>>" & strInfo
    
    rtbLog.Text = rtbLog.Text & strLogText & vbCrLf
    
    If Not blnResult Then
        ChangeColor Len(rtbLog.Text) - Len(strLogText) - 2, Len(strLogText)
    End If
    
    rtbLog.SelStart = Len(rtbLog.Text)
    
    LogFile strCurDate & ":" & strInfo
End Sub

Private Sub PrintResult(blnResult As Boolean)
    
    If chkLog.Value <> 0 Then Call PrintNullLine(2)
    
    If blnResult Then
'        rtbLog.SelStart = Len(rtbLog.Text)
'        rtbLog.SelText = "���Գɹ���" & vbCrLf
        rtbLog.Text = rtbLog.Text & "���Գɹ��� " & vbCrLf
        
        ChangeFontSize Len(rtbLog.Text) - Len("���Գɹ��� " & vbCrLf), Len("���Գɹ��� ")
        rtbLog.SelStart = Len(rtbLog.Text)
    Else
'        rtbLog.SelStart = Len(rtbLog.Text)
'        rtbLog.SelText = "����ʧ�ܣ�" & vbCrLf
        rtbLog.Text = rtbLog.Text & "����ʧ�ܣ� " & vbCrLf
        ChangeFontSize Len(rtbLog.Text) - Len("����ʧ�ܣ� " & vbCrLf), Len("����ʧ�ܣ� ")
        ChangeColor Len(rtbLog.Text) - Len("����ʧ�ܣ� "), Len("����ʧ�ܣ� ")
        rtbLog.SelStart = Len(rtbLog.Text)
    End If
End Sub

Private Sub PrintTag(lngStyle As Long)
    If chkLog.Value = 0 Then Exit Sub
    
'    rtbLog.SelStart = Len(rtbLog.Text)
    Select Case lngStyle
        Case 1
'            rtbLog.SelText = "----------------------------------------------" & vbCrLf
            rtbLog.Text = rtbLog.Text & "----------------------------------------------" & vbCrLf
        Case 2
'            rtbLog.SelText = "**********************************************" & vbCrLf
            rtbLog.Text = rtbLog.Text & "**********************************************" & vbCrLf
        Case 3
'            rtbLog.SelText = "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%" & vbCrLf
            rtbLog.Text = rtbLog.Text & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%" & vbCrLf
        Case 4
'            rtbLog.SelText = "==============================================" & vbCrLf
            rtbLog.Text = rtbLog.Text & "==============================================" & vbCrLf
    End Select
    
    rtbLog.SelStart = Len(rtbLog.Text)
End Sub

Private Sub ChangeColor(lngStart As Long, lngLen As Long)
    
    rtbLog.SelStart = lngStart
    rtbLog.SelLength = lngLen
    rtbLog.SelColor = vbRed
End Sub

Private Sub ChangeFontSize(lngStart As Long, lngLen As Long)
    
    rtbLog.SelStart = lngStart
    rtbLog.SelLength = lngLen
    rtbLog.SelFontSize = 12
    rtbLog.SelBold = True
End Sub

Private Sub PrintNullLine(lngRows As Long)
    Dim i As Long
    
    If chkLog.Value = 0 Then Exit Sub
    
    For i = 1 To lngRows
'        rtbLog.SelText = vbCrLf
        rtbLog.Text = rtbLog.Text & vbCrLf
        rtbLog.SelStart = Len(rtbLog.Text)
    Next
End Sub

Private Function GetTestFile(ByVal lngSize As Long, Optional ByVal strDir As String = "", _
    Optional ByVal strFileName As String = "") As String
'���ɲ����ļ�
    
    Dim lngFileNum As Long
    Dim strFilePath As String
    Dim objFSO As Object
    Dim objLogFile As Object
    Dim strInfo As String
    Dim strLine As String
    
    If strFileName = "" Then
        strFilePath = FormatPath(App.Path & "\Upload" & "\" & M_STR_FILENAME & lngSize & ".txt")
    Else
        strFilePath = FormatPath(App.Path & IIf(strDir = "", "", "\" & strDir & "\") & strFileName)
    End If

    If Len(Dir(strFilePath)) = 0 Then
        strInfo = 1
        lngFileNum = FreeFile
        
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        If Len(Dir(strFilePath)) = 0 Then
            objFSO.CreateTextFile strFilePath, True
        End If
        Set objLogFile = objFSO.GetFile(strFilePath)
        
        Select Case lngSize
            Case 0  '1K
                strInfo = GetString(strInfo, 10)
            Case 1  '512K
                strInfo = GetString(strInfo, 19)
            Case 2  '1M
                strInfo = GetString(strInfo, 20)
            Case 3  '5M
                strInfo = GetString(strInfo, 20)
                strInfo = CopyStr(strInfo, 4)
            Case 4  '10M
                strInfo = GetString(strInfo, 20)
                
                strInfo = CopyStr(strInfo, 9)
        End Select
        
        Open Trim(strFilePath) For Output As #lngFileNum
        Print #lngFileNum, strInfo
        Close #lngFileNum
    End If

    GetTestFile = strFilePath
End Function

Private Function GetString(strInfo As String, lngTimes As Long) As String
    Dim i As Long
    
    For i = 1 To lngTimes
        strInfo = CopyStr(strInfo, 1)
    Next
    
    GetString = strInfo
End Function

Private Function CopyStr(strNeedCopy As String, lngTimes As Long) As String
    Dim i As Long
    
    CopyStr = strNeedCopy
    
    For i = 1 To lngTimes
        CopyStr = CopyStr & strNeedCopy
    Next
End Function




Private Sub Form_Load()
    Dim strCmdLine As String
    
    On Error GoTo errHandle
    
    strCmdLine = Command()
    
    If Len(strCmdLine) = 0 Then
        Call InitLocalPars
    Else
        Call InitInterFace(strCmdLine)
    End If
    
    mlngPassive = Val(GetSetting("ZLSOFT", "����ģ��\Ftp", "���ñ�������", 0))
    
    Call CreateLocalDir
    
    Call HookDefend(txtFtpPassWord.hwnd)
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub CreateLocalDir()
On Error GoTo errHandle
    If Dir(FormatPath(App.Path & "\DosMod"), vbDirectory) <> "DosMod" Then MkDir FormatPath(App.Path & "\DosMod")
    If Dir(FormatPath(App.Path & "\Download"), vbDirectory) <> "Download" Then MkDir FormatPath(App.Path & "\Download")
    If Dir(FormatPath(App.Path & "\Upload"), vbDirectory) <> "Upload" Then MkDir FormatPath(App.Path & "\Upload")
Exit Sub
errHandle:
    MsgBox Err.Description
End Sub

Private Sub InitInterFace(strCmdLine As String)
    Dim arrPara() As String

    arrPara = Split(strCmdLine, "||")
    
    If UBound(arrPara) >= 2 Then
        txtFtpUser.Text = arrPara(0)
        txtFtpPassWord.Text = arrPara(1)
        txtFtpAdress.Text = arrPara(2)
        txtFtpVirtual.Text = arrPara(3)
        txtTestTimes.Text = 1
        gblnTest = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
    If Not mobjCurDos Is Nothing Then
        Cancel = True
        MsgBox "��ȴ�ִ�н���."
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Call ChangeLocation
End Sub

Private Sub ChangeLocation()
    Dim i As Long
    
    For i = 0 To 4
        lblTranSpeed(i).Left = chkSize(i).Left + (chkSize(i).Width - lblTranSpeed(i).Width) / 2
    Next
    
End Sub

Private Sub InitLocalPars()
    txtFtpUser.Text = GetSetting("ZLSOFT", "����ģ��\" & App.EXEName, "USER")
    txtFtpAdress.Text = GetSetting("ZLSOFT", "����ģ��\" & App.EXEName, "Adress")
    txtTestTimes.Text = GetSetting("ZLSOFT", "����ģ��\" & App.EXEName, "TestTimes", 1)
    txtFtpVirtual.Text = GetSetting("ZLSOFT", "����ģ��\" & App.EXEName, "FtpVirtual")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    Call SaveSetting("ZLSOFT", "����ģ��\Ftp", "���ñ�������", mlngPassive)
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.EXEName, "User", Trim(txtFtpUser.Text)
    SaveSetting "ZLSOFT", "����ģ��\" & App.EXEName, "Adress", Trim(txtFtpAdress.Text)
    SaveSetting "ZLSOFT", "����ģ��\" & App.EXEName, "TestTimes", Trim(txtTestTimes.Text)
    SaveSetting "ZLSOFT", "����ģ��\" & App.EXEName, "FtpVirtual", Trim(txtFtpVirtual.Text)
     
    Set mfrmDownLoad = Nothing
    Set mfrmUpLoad = Nothing
End Sub

Private Sub mfrmDownLoad_DoDownLoad(ByVal strLocal As String, ByVal strFile As String)
    Dim i As Long, j As Long
    Dim arrFileName() As String
    Dim strFileName As String
    Dim lngResult As Long
    Dim strVirtual As String
    Dim lngTime As Long
    Dim lngTestCount As Long
    Dim strLocalFile As String
     
     If Len(Trim(strFile)) <= 0 Then MsgBox "�����ļ�������Ϊ�ա�"
     
    Call ClearFace
 
    lngTestCount = Val(txtTestTimes.Text)
    
    If (chkConEvery.Value = 0) Or (mftpNet.hConnection = 0) Then
        If Not LoginTest(Trim(txtFtpAdress), Trim(txtFtpUser), Trim(txtFtpPassWord), lngTime) Then
            MsgBox "FTP����ʧ�ܣ������¼��Ϣ��", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    
    cmdVerify.Caption = "��֤��..."
    
    For i = 0 To lngTestCount - 1
    
        txtCount.Text = i
        
        If chkConEvery.Value <> 0 Then
            If Not LoginTest(Trim(txtFtpAdress), Trim(txtFtpUser), Trim(txtFtpPassWord), lngTime) Then
                MsgBox "FTP����ʧ�ܣ������¼��Ϣ��", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        

        strVirtual = ""
        arrFileName = Split(strFile, "/")
        strFileName = arrFileName(UBound(arrFileName))
        
        For j = 0 To UBound(arrFileName) - 1
            If arrFileName(j) <> "" Then
                strVirtual = strVirtual & "/" & arrFileName(j)
            End If
        Next
        
        
        
        Call PrintNullLine(1)
        PrintTag 3
        PrintLog "��ʼ�ļ����ز���..."
        PrintTag 3
        
        strLocalFile = FormatPath(strLocal & "\" & strFileName & "_") & GetTickCount
        
        lngResult = mftpNet.FuncDownloadFile(strVirtual, strLocalFile, strFileName, IIf(chkForcedRead.Value = 1, True, False))
        If lngResult <> 0 Then
            PrintLog "FTP�ļ���" & GetFTPString(strVirtual & "/" & strFileName) & "�����ز���ʧ�ܣ�����ԭ�����ص��ļ������ڻ򱾵�·��������", False
        Else
            PrintLog "FTP�ļ���" & GetFTPString(strVirtual & "/" & strFileName) & "���ɹ����ص����ء�" & strLocal & "����", True
        End If
        
        Call PrintNullLine(2)
            
        If chkConEvery.Value <> 0 Then
            Call mftpNet.FuncFtpDisConnect
        End If
        
        DoEvents

    Next
    
    Call mftpNet.FuncFtpDisConnect
    
    PrintTag 3
    PrintLog "�ļ����ز�����ɡ�"
    PrintTag 3
    
    cmdVerify.Caption = "��֤"
            

End Sub

Private Sub mfrmUpLoad_DoUpLoad(ByVal strFtpRoad As String, arrFiles() As String)
    Dim i As Long
    Dim arrFileName() As String
    Dim strFileName As String
    Dim lngResult As Long
    Dim lngTime As Long
    
    Call ClearFace
    If mftpNet.hConnection = 0 Then
        If Not LoginTest(Trim(txtFtpAdress), Trim(txtFtpUser), Trim(txtFtpPassWord), lngTime) Then
            MsgBox "FTP����ʧ�ܣ������¼��Ϣ��", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    
    cmdVerify.Caption = "��֤��..."
    Call PrintNullLine(1)
    PrintTag 3
    PrintLog "��ʼ�ļ��ϴ�����..."
    PrintTag 3
    
    For i = 0 To UBound(arrFiles)
        If Len(Trim(arrFiles(i))) > 0 Then
            arrFileName = Split(arrFiles(i), "\")
            strFileName = arrFileName(UBound(arrFileName))
            
            mftpNet.FuncChangeDir ""
            
            If mftpNet.FuncChangeDir(strFtpRoad) = 0 Or GetFTPString(strFtpRoad) = "/" Then
                lngResult = mftpNet.FuncUploadFile(strFtpRoad, arrFiles(i), strFileName)
                If lngResult <> 0 Then
                    PrintLog "�����ļ���" & arrFiles(i) & "���ϴ�����ʧ�ܣ�����ԭ�򣺱����ļ������ڻ�FTPĿ¼������", False
                Else
                    PrintLog "�����ļ���" & arrFiles(i) & "���ɹ��ϴ���FTPĿ¼��" & GetFTPString(strFtpRoad) & "���¡�", True
                End If
            Else
                PrintLog "�����ļ���" & arrFiles(i) & "���ϴ�����ʧ�ܣ�ԭ��FTPĿ¼��" & strFtpRoad & "��������", False
            End If

        End If
    Next
    
    PrintTag 3
    PrintLog "�ļ��ϴ�������ɡ�"
    PrintTag 3
    cmdVerify.Caption = "��֤"
End Sub


Private Sub txtTestTimes_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandle
    
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And Chr(KeyAscii) <> vbBack Then
        KeyAscii = 0
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

'Private Sub txtTimeOut_KeyPress(KeyAscii As Integer)
'    On Error GoTo errHandle
'
'    If InStr("0123456789.", Chr(KeyAscii)) = 0 And Chr(KeyAscii) <> vbBack Then
'        KeyAscii = 0
'    End If
'
'    If KeyAscii = 46 And InStr(txtTimeOut.Text, ".") > 0 Then
'        KeyAscii = 0
'    End If
'
'    Exit Sub
'errHandle:
'    MsgBox Err.Description, vbCritical, Me.Caption
'End Sub

Private Sub DeleteFiles(strVirtual As String)
    Dim strResult As String
    Dim arrFile() As String
    
    strResult = mftpNet.FuncDirFiles(strVirtual)
    
    If Len(strResult) = 0 Then Exit Sub
    
    
    arrFile = Split(strResult, "|")
    
    'For i = 0 To UBound(arrFile)
        mftpNet.FuncDelFiles strVirtual, arrFile
    'Next
End Sub

Private Sub LogFile(ByVal strInfo As String)
    Dim lngFileNum As Long
    Dim FilePath As String
    Dim objFSO As Object
    Dim objLogFile As Object
    
    FilePath = FormatPath(App.Path & "\" & "FtpToolTest" & ".log")

    lngFileNum = FreeFile
 
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Len(Dir(FilePath)) = 0 Then
        objFSO.CreateTextFile FilePath, True
    End If
    Set objLogFile = objFSO.GetFile(FilePath)
    If objLogFile = Empty Then
        Open FilePath For Output As #lngFileNum
    Else
        If objLogFile.Size > 2097152 Then
            objLogFile.Copy FormatPath(App.Path & "\FtpToolTest_" & Format(Now(), "yyyymmdd_hhmmss") & ".log")
            
            Open FilePath For Output As #lngFileNum
        Else
            Open FilePath For Append As #lngFileNum
        End If
    End If
 
    Print #lngFileNum, strInfo
    Close #lngFileNum
 
End Sub

