VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLisPic2Ftp 
   BackColor       =   &H00FFFFFF&
   Caption         =   "����ͼƬ����ת��"
   ClientHeight    =   10110
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   15255
   Icon            =   "frmLisPic2Ftp.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmLisPic2Ftp.frx":6852
   ScaleHeight     =   10110
   ScaleWidth      =   15255
   StartUpPosition =   1  '����������
   Begin VB.Frame frmType 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������Դ"
      Height          =   645
      Left            =   840
      TabIndex        =   43
      Top             =   5280
      Width           =   10800
      Begin VB.OptionButton optNew 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�°�LIS"
         Height          =   255
         Left            =   1680
         TabIndex        =   45
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton optOld 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�ϰ�LIS"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label lblBanner 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   3000
         TabIndex        =   46
         Top             =   330
         Width           =   90
      End
   End
   Begin VB.Timer Timer 
      Interval        =   3000
      Left            =   12960
      Top             =   1680
   End
   Begin VB.CommandButton cmdMulti 
      Caption         =   "ת��ͼƬ(&O)"
      Height          =   350
      Left            =   8880
      TabIndex        =   15
      ToolTipText     =   "�����ݿ��е�ͼ������ת��ΪͼƬ���浽���ػ�FTP������"
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Frame frmFtpUp 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "����ģʽ"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   840
      TabIndex        =   37
      Top             =   3600
      Width           =   10815
      Begin VB.TextBox txtProc 
         Height          =   300
         Left            =   840
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "2"
         Top             =   1140
         Width           =   375
      End
      Begin VB.OptionButton optAuto 
         BackColor       =   &H00FFFFFF&
         Caption         =   " ʵʱ�ϴ�"
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton optManu 
         BackColor       =   &H00FFFFFF&
         Caption         =   " �첽�ϴ�"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label lblFTP 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ʹ�ö���̼ӿ�ͼƬת����ٶ�,Ϊ�˷�ֹ�����������Դ����,��������Ϊ10��"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   2
         Left            =   1560
         TabIndex        =   42
         Top             =   1200
         Width           =   6480
      End
      Begin VB.Label lblPorc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label lblFTP 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "����ͼƬת�浽���غ�ʹ���������������ϴ�����Ҫ�������ؿռ䣬�����ʱ��."
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   1560
         TabIndex        =   39
         Top             =   750
         Width           =   6570
      End
      Begin VB.Label lblFTP 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÿ��ͼƬת���������ϴ���FTP�����豾�ؿռ�С�������ʱ�ϳ���"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   38
         Top             =   330
         Width           =   5310
      End
   End
   Begin VB.CommandButton cmdCommit 
      Caption         =   "������Ϣ(&U)"
      Height          =   350
      Left            =   10320
      TabIndex        =   16
      ToolTipText     =   $"frmLisPic2Ftp.frx":E88C
      Top             =   7320
      Width           =   1335
   End
   Begin VB.FileListBox fileList 
      Height          =   450
      Left            =   10560
      Pattern         =   "*.jpg;*.png"
      TabIndex        =   36
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame frmDownd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����Χ"
      Height          =   1245
      Left            =   840
      TabIndex        =   23
      Top             =   6000
      Width           =   10800
      Begin VB.OptionButton OptAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ȫ������"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   420
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optPart 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��������"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   1095
      End
      Begin VB.PictureBox pctTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         ScaleHeight     =   375
         ScaleWidth      =   5415
         TabIndex        =   33
         Top             =   720
         Visible         =   0   'False
         Width           =   5415
         Begin MSComCtl2.DTPicker dtpStart 
            Height          =   345
            Left            =   840
            TabIndex        =   13
            Top             =   15
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   223608835
            CurrentDate     =   43077.4366203704
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   345
            Left            =   3240
            TabIndex        =   14
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   223608835
            CurrentDate     =   43077.4366782407
         End
         Begin VB.Label lblDown 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ʼʱ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   0
            TabIndex        =   35
            Top             =   70
            Width           =   720
         End
         Begin VB.Label lblDown 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "����ʱ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   2400
            TabIndex        =   34
            Top             =   75
            Width           =   720
         End
      End
   End
   Begin VB.Frame fraFTP 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FTP����"
      Height          =   1290
      Left            =   840
      TabIndex        =   0
      Top             =   2160
      Width           =   10800
      Begin VB.CommandButton cmdFile 
         Caption         =   "��"
         Height          =   300
         Left            =   7800
         TabIndex        =   6
         Top             =   780
         Width           =   300
      End
      Begin VB.TextBox txtTmpPath 
         Height          =   300
         Left            =   4320
         TabIndex        =   5
         ToolTipText     =   "ͼ��������ת��ʱ������ͼƬ��������ʱ·����,�������㹻�ı���ռ�"
         Top             =   780
         Width           =   3495
      End
      Begin VB.CommandButton cmdFtp 
         Caption         =   "���Ӳ���"
         Height          =   350
         Left            =   8280
         TabIndex        =   7
         Top             =   755
         Width           =   1215
      End
      Begin VB.TextBox txtFTPPath 
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   780
         Width           =   1500
      End
      Begin VB.TextBox txtFTPIP 
         Height          =   300
         Left            =   1440
         TabIndex        =   1
         Top             =   315
         Width           =   1500
      End
      Begin VB.TextBox txtFTPPWD 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6480
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   315
         Width           =   1620
      End
      Begin VB.TextBox txtFTPUser 
         Height          =   300
         Left            =   4320
         TabIndex        =   2
         Top             =   315
         Width           =   1500
      End
      Begin VB.Label lblDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "������ʱ·��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3195
         TabIndex        =   40
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lblFTP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FTP�ļ�·��"
         Height          =   180
         Index           =   6
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lblFTP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ַ"
         Height          =   180
         Index           =   5
         Left            =   990
         TabIndex        =   19
         Top             =   375
         Width           =   360
      End
      Begin VB.Label lblFTP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   4
         Left            =   6075
         TabIndex        =   18
         Top             =   375
         Width           =   360
      End
      Begin VB.Label lblFTP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�û�"
         Height          =   180
         Index           =   3
         Left            =   3915
         TabIndex        =   17
         Top             =   375
         Width           =   360
      End
   End
   Begin VB.PictureBox pctResult 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   720
      ScaleHeight     =   975
      ScaleWidth      =   10935
      TabIndex        =   28
      Top             =   9240
      Visible         =   0   'False
      Width           =   10935
      Begin VB.Label lblReult 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "���ƺ�ʱ:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblReult 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ȡ����-9000;�ϴ�����-8500;��ȡʧ��-1000;�ϴ�ʧ��-500"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   4860
      End
      Begin VB.Label lblReult 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ת�����:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   810
      End
   End
   Begin VB.PictureBox pctProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   720
      ScaleHeight     =   1215
      ScaleWidth      =   10935
      TabIndex        =   24
      Top             =   7920
      Visible         =   0   'False
      Width           =   10935
      Begin MSComctlLib.ProgressBar pgsBar 
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   10000
      End
      Begin VB.Label lblProgress 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "�Ѿ�ת��4000��"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label lblProgress 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "����10000�����ݴ�ת��"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   1890
      End
   End
   Begin VB.Label lblState 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "���ڲ�ѯ��ת������..."
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   840
      TabIndex        =   32
      Top             =   7320
      Width           =   1890
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmLisPic2Ftp.frx":E922
      Stretch         =   -1  'True
      Top             =   648
      Width           =   480
   End
   Begin VB.Label lblTip 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLisPic2Ftp.frx":F19D
      Height          =   1440
      Left            =   840
      TabIndex        =   22
      Top             =   720
      Width           =   10815
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ͼƬ����ת��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   21
      Top             =   120
      Width           =   1920
   End
End
Attribute VB_Name = "frmLisPic2Ftp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrTmpPath As String    '����ͼƬ��ʱ·��
Private mstrFtpPath As String    '�ϴ�ͼƬFTP·��
Private mblnFtpConncted  As Boolean      '��ʶFTP�Ƿ��Ѿ�����
Private mlngImgUp As Long '����·���´��ϴ�ͼƬ����
Private mlngImgDown As Long '����·���´�����ͼƬ����
Private mintCpu As Integer  'CPU����ֵ
Private mclsFtp As New clsFtp   'FTP��
Private mblnUpload As Boolean   '�Ƿ�����ת��
Private mdblTime As Double 'ת��ʱ��
Private mintLisBanner  As Integer 'LIS�汾:0=û�а�װLis 1=ֻ�оɰ�LIS 2=ֻ���°�LIS 3=���߾���

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function
Private Sub cmdCommit_Click()
    '����ת����Ϻ�,�޸�Դ����
    Dim strSQL As String, strMsg As String
    Dim lngOldNums As Long, lngNewNums As Long, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If CheckTblExist("����ͼ����_Exp_Temp") Then
        strSQL = "Select Count(1) ���� From ����ͼ����_Exp_Temp"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetNums")
        lngOldNums = rsTmp!����
    End If
    If CheckTblExist("���鱨��ͼ��_Exp_Temp") Then
        strSQL = "Select Count(1) ���� From ���鱨��ͼ��_Exp_Temp"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetNums")
        lngNewNums = rsTmp!����
    End If
    
    If lngOldNums + lngNewNums = 0 Then
        MsgBox "��ʱ���ݱ�û������,����ִ��ת��ͼƬ���ܡ�", , "��ʾ"
        Exit Sub
    End If
    
    strMsg = "�������������ɾ������������ݵķ�ʽ������ԭ���ͼƬ·����Ϣ,ͬʱ���LOB�ֶε�ͼ�����ݺ���ʱ���ݱ�" & vbNewLine & _
                    "��ǰ���޸������ݿ⹲��" & lngOldNums + lngNewNums & "�����ݣ������ϴ�������ͼƬ��FTP�����ȷ�Ϻ�ִ�С���ȷ��Ҫ������"
    If MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, "ȷ��") = vbNo Then
        Exit Sub
    End If
    
    pctProgress.Visible = False
    pctResult.Visible = False
    If lngOldNums > 0 Then UpdatePic 1
    If lngNewNums > 0 Then UpdatePic 2
    MousePointer = vbDefault
    lblState.Caption = "���ݸ��³ɹ�,�޸�����" & lngOldNums + lngNewNums & "����"
    Exit Sub
errH:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Sub

Private Sub cmdFile_Click()
    Dim strTmp As String, blnTmp As Boolean
    
    strTmp = OpenFolder(Me, "��ѡ����ʱ·��", mstrTmpPath)
    If strTmp = "" Then
        Exit Sub
    End If
    
    '�������ѡ�����ʱ·����ƥ��,�ᵼ����һ���ϴ�ʧ�ܵ��ļ�����,���ת�����ݶ�ʧ
    blnTmp = True
    If mstrTmpPath <> strTmp And mstrTmpPath <> "" Then
        blnTmp = MsgBox("��ǰѡ�����ʱ·������һ�ε���ʱ·����ͬ���޷������ϴ��ϴ�δ������FTP��ͼƬ���Ƿ������" & vbNewLine & "ע��δ�ϴ���ͼƬ�����ֶ��ϴ���" _
        , vbYesNo + vbQuestion + vbDefaultButton2, "ȷ��") = vbYes
    End If
    
    If blnTmp Then
        mstrTmpPath = strTmp
        If Right(mstrTmpPath, 1) = "\" Then
            mstrTmpPath = Mid(mstrTmpPath, 1, Len(mstrTmpPath) - 1)
        End If
        Call SaveSetting("LISͼƬת��", "ת������", "��ʱ·��", mstrTmpPath)
        txtTmpPath.Text = strTmp
        fileList.Path = strTmp
        
        CheckImg
    End If
End Sub

Private Sub cmdFtp_Click()
    If TestFtp() Then lblState.Caption = "FTP������֤ͨ��"
End Sub


Private Sub cmdMulti_Click()
    Dim i As Integer, strSQL As String, rsTmp As ADODB.Recordset
    Dim rsExp As ADODB.Recordset, strMsg As String
    Dim lngSize As Long, lngTmp As Long, strPath As String
    Dim intProcNum As Integer, lngDays As Long
    
    SetCmdEnable False
    pctProgress.Visible = False
    pctResult.Visible = False
    
    '���ȼ��FTP�����Ƿ�ͨ��
    If Not TestFtp Then
        MousePointer = vbDefault
        SetCmdEnable True
        Exit Sub
    End If
    
    '����ӳ����Ƿ����
    If Right(App.Path, 1) = "\" Then
        strPath = Mid(App.Path, 1, Len(App.Path) - 1)
    Else
        strPath = App.Path
    End If
    If Not gobjFile.FileExists(strPath & "\zlLisPic2FtpSub.exe") Then
        MsgBox "Ŀ¼" & strPath & "�²�����ִ���ļ�:" & "zlLisPic2FtpSub.exe,�޷�����������", , gstrSysName
        SetCmdEnable True
        Exit Sub
    End If
    
    CreateTable IIf(optOld.Value, 1, 2)
    If optOld.Value = True Then
    '��ת��ͼƬ����10000ʱ,�ͽ�����ʾ,�Ƿ����ύһ���ٽ���ת������
        strSQL = "Select Count(1) ���� From ����ͼ����_EXP_TEMP"
    Else
        strSQL = "Select Count(1) ���� From ���鱨��ͼ��_EXP_TEMP"
    End If
    Set rsExp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetExp")

    If rsExp!���� > 100000 Then
        strMsg = "��ǰ�Ѿ�ת������ʱͼƬ����" & rsExp!���� & "�����Ƿ��ȸ�����Ϣ�����ݿ�?" & vbNewLine & _
                        "ע:��ʱ�������ݹ���,�ᵼ��ת��ͼƬ����������Ǹ�����Ϣ��������������ͼƬת��"
        If MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, "ȷ��") = vbYes Then
            lblState.Caption = "�����ύ��ת�����������ݿ���..."
            MousePointer = vbArrowHourglass
            UpdatePic IIf(optOld.Value, 1, 2)
            lblState.Caption = ""
            MousePointer = vbDefault
            SetCmdEnable True
            Exit Sub
        End If
    End If
    
    lblState.Caption = "���ת���������..."
    MousePointer = vbArrowHourglass
    lblState.Refresh
    mlngImgDown = GetDownNum(IIf(optOld.Value, 1, 2))
    
    If mlngImgDown = 0 Then
        MsgBox "���ݿ�������ͼƬ���ݶ��Ѿ�ת��", , "��ʾ"
        lblState.Caption = ""
        MousePointer = vbDefault
        SetCmdEnable True
        Exit Sub
    Else
        If optManu.Value Then   '�ֶ��ϴ�,��Ҫ��ʾռ�ñ��ؿռ�
            If optPart.Value Then
                strMsg = "��������ת��������ת��ͼƬ����" & mlngImgDown & "���������ȷ���Ƿ������"
            Else
                lngSize = GetLobSize(IIf(optOld.Value, 1, 2))
                strMsg = "��������ת��������ת��ͼƬ����" & mlngImgDown & "����Ԥ��ռ����ʱ�ռ�" & lngSize & "M�������ȷ���Ƿ������"
            End If
        Else
            strMsg = "��������ת��������ת��ͼƬ����" & mlngImgDown & "���������ȷ���Ƿ������"
        End If
        
        If MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, "ȷ��") = vbNo Then
            lblState.Caption = ""
            MousePointer = vbDefault
            SetCmdEnable True
            Exit Sub
        End If
    End If
    

    
    '����ˢ��
    lblState.Caption = "���ڿ���ת������..."
    lblState.Refresh
    lngDays = DateDiff("d", CDate(dtpStart.Tag), CDate(dtpEnd.Tag)) + 1
    pgsBar.Max = lngDays
    pgsBar.Value = 0
    lblProgress(0).Caption = Format(CDate(dtpStart.Tag), "yyyy/mm/dd") & "��" & Format(CDate(dtpEnd.Tag), "yyyy/mm/dd") & "����" & lngDays & "������ݴ�ת��"
    lblProgress(1).Caption = "�Ѿ�ת��0������ݡ�"

    pctProgress.Visible = True
    pctResult.Visible = False
    MousePointer = vbArrowHourglass
    Me.Refresh
    
    '��������
    mblnUpload = True
    mdblTime = 0
    SaveSetting "LISͼƬת��", "ת������", "ת������", ""   '��ʼ����ʱ �ȰѴ������
    intProcNum = IIf(Val(txtProc.Text) = 0, 1, txtProc.Text)
    For i = 1 To intProcNum
        SaveSetting "LISͼƬת��", "ת������", "����" & i, 0
        
        '֪ͨ���̿���ת��
        '�����ʽ: ת������(1-FTP 2-���汾��);������Դ(1-�ɰ�LIS 2-�°�LIS);���̺�;��ʼʱ��;����ʱ��;��ʱ·��;FTP·��
        If optAuto.Value = True Then
            SaveSetting "LISͼƬת��", "ת������", "��������", "1;" & IIf(optOld.Value, 1, 2) & ";" & i & ";" & intProcNum & ";" & Format(dtpStart.Tag, "yyyy/mm/dd") & ";" & Format(dtpEnd.Tag, "yyyy/mm/dd") & ";" & mstrTmpPath & ";" & mstrFtpPath 'ͬ���ϴ�
        Else
            SaveSetting "LISͼƬת��", "ת������", "��������", "2;" & IIf(optOld.Value, 1, 2) & ";" & i & ";" & intProcNum & ";" & Format(dtpStart.Tag, "yyyy/mm/dd") & ";" & Format(dtpEnd.Tag, "yyyy/mm/dd") & ";" & mstrTmpPath & ";" & mstrFtpPath    '�첽�ϴ�
        End If
        Shell """" & strPath & "\zlLisPic2FtpSub.exe"" ""zlUserName=" & gstrUserName & "zlPassword=" & gstrPassword & "zlServer=" & gstrServer & " ", vbMaximizedFocus
    Next
    Call Timer_Timer
    

End Sub

Private Sub Form_load()

    On Error GoTo errH
    Call LoadFtpPara
    mstrTmpPath = GetSetting("LISͼƬת��", "ת������", "��ʱ·��")
    txtTmpPath.Text = mstrTmpPath
    dtpStart.Value = Now: dtpEnd.Value = Now
    
    Call CheckImg
    
    mintCpu = GetCpuAdv
    mintLisBanner = CheckLisSys
    If mintLisBanner = 1 Or mintLisBanner = 3 Then
        SetDtpPicker 1
    ElseIf mintLisBanner = 2 Then
        SetDtpPicker 2
    End If
    
    Exit Sub
errH:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsFtp = Nothing
End Sub

Private Sub optNew_Click()
    Call SetDtpPicker(2)
End Sub

Private Sub optOld_Click()
    Call SetDtpPicker(1)
End Sub

Private Sub optPart_Click()
    pctTime.Visible = True
End Sub

Private Sub optAll_Click()
    pctTime.Visible = False
End Sub

Private Sub LoadFtpPara()
    '����:��ȡ�û���FTP��������
    Dim strFtpParas As String
    
    On Error GoTo errH
    strFtpParas = gclsBase.GetPara("FTP����", 100, 1208, 1)
    
    If strFtpParas = "" Then
        strFtpParas = GetSetting("LISͼƬת��", "ת������", "FTP·��")
    End If
        
    If strFtpParas <> "" Then
        txtFTPUser.Text = Split(strFtpParas, ";")(0)
        txtFTPPWD.Text = Split(strFtpParas, ";")(1)
        txtFTPIP.Text = Split(strFtpParas, ";")(2)
        txtFTPPath.Text = Split(strFtpParas, ";")(3)
    End If
    Exit Sub
errH:
    MsgBox "FTP���ö�ȡʧ��", vbExclamation, gstrSysName
End Sub

Private Function TestFtp() As Boolean
    '����:����ftp������֤�����Ƿ�ͨ��
    Dim strUser As String, strPwd As String
    Dim strIp As String, strPath As String
    
    On Error GoTo errH
    strUser = Trim(txtFTPUser.Text): strPwd = Trim(txtFTPPWD.Text)
    strIp = Trim(txtFTPIP.Text): strPath = Trim(txtFTPPath.Text)
    mstrFtpPath = strPath
    
    SaveSetting "LISͼƬת��", "ת������", "FTP·��", strUser & ";" & strPwd & ";" & strIp & ";" & strPath
    
    If mblnFtpConncted Then mclsFtp.FuncFtpDisConnect  '����Ѿ�������FTP,��ô�ͶϿ�֮ǰ��,��ֹռ��FTP�Ự��
    mblnFtpConncted = False
    '�������Ӳ���
    If mclsFtp.FuncFtpConnect(strIp, strUser, strPwd) = 0 Then
        MsgBox "����FTP������ʧ�ܣ������������ַ���ʺš�"
        Exit Function
    End If
    
    '����һ���ļ������ϴ�����
    If Not gobjFile.FolderExists(mstrTmpPath) Then
        MsgBox "��ʱ·��������,����������", , "��ʾ"
        Exit Function
    End If
    If Not gobjFile.FileExists(mstrTmpPath & "\tmp") Then
        gobjFile.CreateTextFile mstrTmpPath & "\tmp", True
    End If
    
    If mclsFtp.FuncUploadFile(strPath, mstrTmpPath & "\tmp", "tmp") <> 0 Then
        MsgBox "FTP·������,δ��ͨ������"
        Exit Function
    End If
    

    'ɾ����ʱ�ļ�
    Kill mstrTmpPath & "\tmp"
    mblnFtpConncted = True
    TestFtp = True
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Function

Private Function CheckImg() As Boolean
    '����:�����ʱĿ¼���Ƿ���ͼƬδ�ϴ�
    
    If gobjFile.FolderExists(mstrTmpPath) Then
        With fileList
            .Path = mstrTmpPath
            .Refresh
            mlngImgUp = .ListCount
        End With
    Else
        mlngImgUp = 0
    End If
    
    If mlngImgUp > 0 Then
        lblState.Caption = "��ǰ��ʱ·���¹���" & mlngImgUp & "��ͼƬδ�ϴ�,��ʹ��ͼƬת�湦�ܼ����ϴ�"
        CheckImg = True
    Else
        lblState.Caption = ""
        CheckImg = False
    End If
End Function

Private Sub Timer_Timer()
    Dim blnDone As Boolean, strTmp As String
    Dim lngDone As Long, intExit As Integer
    Dim i As Integer, intProcNum As Integer, intActive As Integer
    
    If Not mblnUpload Then Exit Sub

    If mdblTime = 0 Then mdblTime = GetTickCount
    
    blnDone = True
    lngDone = 0: intExit = 0
    intProcNum = IIf(txtProc.Text = 0, 1, txtProc.Text)
    
    'ѭ��ÿ������
    For i = 1 To intProcNum
        strTmp = GetSetting("LISͼƬת��", "ת������", "����" & i, 0)   'ת����Ϻ�,�������޸�Ϊ ����;
        If IsNumeric(strTmp) Then
            lngDone = lngDone + strTmp  '�Ѿ�ת��������
        Else
            lngDone = lngDone + Val(Mid(strTmp, 1, Len(strTmp) - 1))
            intExit = intExit + 1   '�Ѿ����ת�����˳��Ľ�������
        End If
        
        If lngDone < pgsBar.Max Then
            blnDone = False
        Else
            blnDone = True
        End If
    Next
    
    '�жϽ����Ƿ���Ϊ����ԭ����ֹ
    intActive = CheckProcExist("zllispic2ftpsub.exe")
    If intActive <> intProcNum Then
        If intProcNum <> intActive + intExit And lngDone <> pgsBar.Max Then    '�˳�����+��Ծ����<>�������� ˵���н���������ֹ
            SaveSetting "LISͼƬת��", "ת������", "ת������", "���̱�������ֹ"
            blnDone = True
        End If
    End If
    '��������
    If Not blnDone Then
        lblState.Caption = "���ڽ���ͼƬת��..."
        lblProgress(1).Caption = "�Ѿ�ת��" & lngDone & "������ݡ�"
        lblProgress(1).Refresh
        
        pgsBar.Value = IIf(lngDone > pgsBar.Max, pgsBar.Max, lngDone)
        pgsBar.Refresh
    Else
        '�������
        SetCmdEnable True
        MousePointer = vbDefault
        mblnUpload = False
        pctResult.Visible = True
        pctResult.Refresh
        lblProgress(1).Caption = "�Ѿ�ת��" & lngDone & "������ݡ�"
        If GetSetting("LISͼƬת��", "ת������", "ת������") <> "" Then
            lblReult(1).Caption = "ת���з�������,������Ϣ�Ѿ���������ǰĿ¼�µ�Lis2FtpErrLog��־�ļ���"
            lblState.Caption = "ת���з�������,��������ԡ�"
        Else
            lblReult(1).Caption = "��ת��ͼƬ:" & mlngImgDown & "��"
            lblState.Caption = "ͼƬת��ɹ�,��ʹ���ύ���ݹ���,���޸ĺ�Ľ���ı������ݿ�"
            pgsBar.Value = pgsBar.Max   '��֤�����������
        End If
        lblReult(2).Caption = "���ƺ�ʱ:" & Format((GetTickCount - mdblTime) / 1000, "0.00") & "S"
    End If
End Sub

Private Sub txtFTPIP_Change()
    If mblnFtpConncted Then mclsFtp.FuncFtpDisConnect '����Ѿ�������FTP,��ô�ͶϿ�֮ǰ��,��ֹռ��FTP�Ự��
    mblnFtpConncted = False
End Sub

Private Sub txtFTPIP_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub txtFTPPath_Change()
    If mblnFtpConncted Then mclsFtp.FuncFtpDisConnect '����Ѿ�������FTP,��ô�ͶϿ�֮ǰ��,��ֹռ��FTP�Ự��
    mblnFtpConncted = False
End Sub

Private Sub txtFTPPath_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub txtFTPPWD_Change()
    If mblnFtpConncted Then mclsFtp.FuncFtpDisConnect '����Ѿ�������FTP,��ô�ͶϿ�֮ǰ��,��ֹռ��FTP�Ự��
    mblnFtpConncted = False
End Sub

Private Sub txtFTPPWD_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub txtFTPUser_Change()
    If mblnFtpConncted Then mclsFtp.FuncFtpDisConnect '����Ѿ�������FTP,��ô�ͶϿ�֮ǰ��,��ֹռ��FTP�Ự��
    mblnFtpConncted = False
End Sub

Private Function GetDownNum(ByVal intType As Integer) As Long
    '����:��ȡ��Ҫ�����ݿ��н���ת����ͼƬ����
    '����: intType 1=�ɰ�LIS 2=�°�LIS
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngTmp As Long
    
    '1.����ȫ�����ݻ��ǲ�������
    If intType = 1 Then
        strSQL = "Select " & IIf(mintCpu = 0, "", "/*+ parallel(a," & mintCpu & ") parallel(b," & mintCpu & ")*/ ") & vbNewLine & _
                        " Count(1) ����, Max(a.����ʱ��) ����ʱ��, Min(a.����ʱ��) ��ʼʱ��" & vbNewLine & _
                        "From ����걾��¼ A, ����ͼ���� B" & vbNewLine & _
                        "Where a.����� Is Not Null And a.Id = b.�걾id And b.ͼ��λ�� Is Null And b.ͼ��� Is Not Null"
    Else
        strSQL = "Select " & IIf(mintCpu = 0, "", "/*+ parallel(a," & mintCpu & ") parallel(b," & mintCpu & ")*/ ") & vbNewLine & _
                " Count(1) ����, Max(a.����ʱ��) ����ʱ��, Min(a.����ʱ��) ��ʼʱ��" & vbNewLine & _
                "From ���鱨���¼ A, ���鱨��ͼ�� B" & vbNewLine & _
                "Where a.����� Is Not Null And a.Id = b.�걾id And b.ͼ��λ�� Is Null And b.ͼ��� Is Not Null"
    End If
    strSQL = strSQL & IIf(optPart = True, " And a.����ʱ�� Between [1] And [2]", "")
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetDownNum", CDate(Format(dtpStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59")))
    lngTmp = rsTmp!����
    
    If lngTmp = 0 Then Exit Function
    
    dtpStart.Tag = rsTmp!��ʼʱ��
    dtpEnd.Tag = rsTmp!����ʱ��
    
    '2.�Ӽ���ͼ����_EXP_TEMP�м�����ת��������
    If Not optPart.Value Then '�����ѡ��ȫ������
        If intType = 1 Then
            strSQL = "Select count(1) ���� From ����ͼ����_EXP_TEMP"
        Else
            strSQL = "Select count(1) ���� From ���鱨��ͼ��_EXP_TEMP"
        End If
    Else
        If intType = 1 Then
            strSQL = "Select Count(1) ����" & vbNewLine & _
                            "From ����걾��¼ A, ����ͼ����_EXP_TEMP B" & vbNewLine & _
                            "Where a.����� Is Not Null And a.Id = b.�걾id" & vbNewLine & _
                            "And a.����ʱ�� Between [1] And [2]"
        Else
            strSQL = "Select Count(1) ����" & vbNewLine & _
                            "From ���鱨���¼ A, ���鱨��ͼ��_EXP_TEMP B" & vbNewLine & _
                            "Where a.����� Is Not Null And a.Id = b.�걾id" & vbNewLine & _
                            "And a.����ʱ�� Between [1] And [2]"
        End If
    End If
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetDownNum", CDate(Format(dtpStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59")))
    lngTmp = lngTmp - rsTmp!����
    
    GetDownNum = lngTmp
End Function

Private Sub txtFTPUser_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub txtProc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If InStr(1, "1234567890", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtProc_LostFocus()
    txtProc.Text = IIf(Val(txtProc.Text) = 0, 1, Val(txtProc.Text))
End Sub

Private Sub txtTmpPath_Change()
    mstrTmpPath = Trim(txtTmpPath.Text)
    
    If Right(mstrTmpPath, 1) = "\" Then
        mstrTmpPath = Mid(mstrTmpPath, 1, Len(mstrTmpPath) - 1)
    End If
End Sub

Private Function GetLobSize(ByVal intType As Integer) As Long
    '����:����ռ�ÿռ�
    '����: intType 1=�ɰ�LIS 2=�°�LIS
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If intType = 1 Then
        strSQL = "Select a.Segment_Name, Round(a.Bytes / 1024 / 1024)  As Lobsize" & vbNewLine & _
                        "From User_Segments A, User_Lobs B" & vbNewLine & _
                        "Where b.Table_Name  = '����ͼ����' And a.Segment_Name = b.Segment_Name"
    Else
        strSQL = "Select a.Segment_Name, Round(a.Bytes / 1024 / 1024)  As Lobsize" & vbNewLine & _
                "From User_Segments A, User_Lobs B" & vbNewLine & _
                "Where b.Table_Name  = '���鱨��ͼ��' And a.Segment_Name = b.Segment_Name"
    End If
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetLobSize")
    If rsTmp.RecordCount = 0 Then Exit Function
    
    GetLobSize = rsTmp!Lobsize & ""
    
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Function

Private Function UpdatePic(ByVal intType As Integer) As Boolean
    '����:�������������ݿ�
    '����: intType 1=�ɰ�LIS 2=�°�LIS
    Dim strSQL As String
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    
    If intType = 1 Then
        strSQL = "Delete ����ͼ���� Where ID In (Select ID From ����ͼ����_Exp_Temp)"
        gcnOracle.Execute strSQL
        
        strSQL = "Insert Into /*+ append */ ����ͼ����" & vbNewLine & _
                        "  (ID, �걾id, ͼ������, ͼ���, ͼ��λ��, ��ת��)" & vbNewLine & _
                        "  Select ID, �걾id, ͼ������, Null, ͼ��λ��, Null From ����ͼ����_Exp_Temp"
        gcnOracle.Execute strSQL
    Else
        strSQL = "Delete ���鱨��ͼ�� Where ID In (Select ID From ���鱨��ͼ��_Exp_Temp)"
        gcnOracle.Execute strSQL
        
        strSQL = "Insert Into /*+ append */ ���鱨��ͼ��" & vbNewLine & _
                        "  (ID, �걾id, ͼ������, ͼ���, ͼ��λ��)" & vbNewLine & _
                        "  Select ID, �걾id, ͼ������, Null, ͼ��λ�� From ���鱨��ͼ��_Exp_Temp"
        gcnOracle.Execute strSQL
    End If
    
    gcnOracle.CommitTrans
    
    'ɾ����ʱ��͹���
    If intType = 1 Then
        strSQL = "Drop Procedure Zl_����ͼ����_Temp_Insert"
        gcnOracle.Execute strSQL
        strSQL = "Drop Table ����ͼ����_EXP_TEMP"
        gcnOracle.Execute strSQL
    Else
        strSQL = "Drop Procedure Zl_���鱨��ͼ��_Temp_Insert"
        gcnOracle.Execute strSQL
        strSQL = "Drop Table ���鱨��ͼ��_EXP_TEMP"
        gcnOracle.Execute strSQL
    End If
        
    Exit Function
errH:
    If InStr(1, UCase(Err.Description), "ORA") Then
        gcnOracle.RollbackTrans
    End If
    MsgBox Err.Description, vbExclamation, gstrSysName
End Function

Private Sub txtTmpPath_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Function GetCpuAdv() As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim intAdvise As Integer
    
    On Error GoTo errH
    strSQL = "Select Nvl(Max(Value),0) CPU From V$parameter Where Name = 'cpu_count'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ����CUP��")
    
    If rsTmp!cpu <= 4 Then
        intAdvise = 1
    ElseIf rsTmp!cpu <= 8 Then
        intAdvise = 4
    ElseIf rsTmp!cpu <= 12 Then
        intAdvise = 8
    Else
        intAdvise = 12
    End If
    
    GetCpuAdv = intAdvise
    Exit Function
errH:
    GetCpuAdv = 0
End Function

Private Function CheckLisSys() As Integer
    '����:��鵱ǰLISϵͳ�İ汾
    '���� 0=û�а�װLis 1=ֻ�оɰ�LIS 2=ֻ���°�LIS 3=���߾���
    Dim blnOld As Boolean, blnNew As Boolean
    
    blnOld = CheckTblExist("����ͼ����")
    blnNew = CheckTblExist("���鱨��ͼ��")
    
    If blnOld And blnNew Then
        CheckLisSys = 3    '����
        lblBanner.Caption = ""
    ElseIf blnOld And Not blnNew Then
        lblBanner.Caption = "��ǰֻ��װ�˾ɰ�LISϵͳ���޷�ѡ��������Դ��"
        optNew.Enabled = False
        CheckLisSys = 1    'ֻ�оɰ�
    ElseIf Not blnOld And blnNew Then
        lblBanner.Caption = "��ǰֻ��װ���°�LISϵͳ���޷�ѡ��������Դ��"
        optOld.Enabled = False
        CheckLisSys = 2    'ֻ���°�
    ElseIf Not blnOld And Not blnNew Then
        lblBanner.Caption = "��ǰû�а�װLISϵͳ���޷�����ת��������"
        SetCmdEnable False
        CheckLisSys = 0    'û��LISϵͳ
    End If
End Function

Private Sub SetDtpPicker(ByVal intType As Integer)
    '����:����dtpPicker��ֵ
    '����:intType 1=�ϰ�LIS 2=�°�LIS
    Dim strSQL As String, rsTmp As ADODB.Recordset
      
    On Error GoTo errH
    '����ת��ͼƬ�����ʼʱ��ͽ���ʱ�����ʱ��ؼ���
    If intType = 1 Then
        strSQL = "Select a.����ʱ��" & vbNewLine & _
                        "From ����걾��¼ A, ����ͼ���� B" & vbNewLine & _
                        "Where a.Id = b.�걾id And (b.Id = (Select Max(ID) As ID From ����ͼ����) Or b.Id = (Select Min(ID) As ID From ����ͼ����))" & vbNewLine & _
                        "Order By a.����ʱ��"
    ElseIf intType = 2 Then
        strSQL = "Select a.����ʱ��" & vbNewLine & _
                        "From ���鱨���¼ A, ���鱨��ͼ�� B" & vbNewLine & _
                        "Where a.Id = b.�걾id And (b.Id = (Select Max(ID) As ID From ���鱨��ͼ��) Or b.Id = (Select Min(ID) As ID From ���鱨��ͼ��))" & vbNewLine & _
                        "Order By a.����ʱ��"
    Else
        Exit Sub
    End If
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "1")
    If rsTmp.RecordCount = 0 Then Exit Sub
    dtpStart.Value = CDate(Trim(rsTmp!����ʱ��)) - 1
    rsTmp.MoveLast
    dtpEnd.Value = CDate(Trim(rsTmp!����ʱ��)) + 1
    
    Exit Sub
errH:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Sub


Private Sub CreateTable(ByVal intType As Integer)
    '����:���ݴ�������ʹ�����ͬ����ʱ������
    '����:  intType 1=�ɰ�Lis���� 2=�°�Lis����
    Dim strSQL As String
    
    If intType = 1 Then
        'û����ʱת����,�ʹ���
        If Not CheckTblExist("����ͼ����_EXP_TEMP") Then
            strSQL = "Create Table ����ͼ����_EXP_TEMP As Select id,�걾id,ͼ������,ͼ��λ�� From ����ͼ���� Where 1=0"
            gcnOracle.Execute strSQL
            strSQL = "Create Or Replace Procedure Zl_����ͼ����_Temp_Insert" & vbNewLine & _
                            "(" & vbNewLine & _
                            "  Id_In       In ����ͼ����_Exp_Temp.Id%Type," & vbNewLine & _
                            "  �걾id_In   In ����ͼ����_Exp_Temp.�걾id%Type," & vbNewLine & _
                            "  ͼ������_In In ����ͼ����_Exp_Temp.ͼ������%Type," & vbNewLine & _
                            "  ͼ��λ��_In In ����ͼ����_Exp_Temp.ͼ��λ��%Type" & vbNewLine & _
                            ") Is" & vbNewLine & _
                            "Begin" & vbNewLine & _
                            "  Insert Into ����ͼ����_Exp_Temp Values (Id_In, �걾id_In, ͼ������_In, ͼ��λ��_In);" & vbNewLine & _
                            "Exception" & vbNewLine & _
                            "  When Others Then" & vbNewLine & _
                            "    zl_ErrorCenter(SQLCode, SQLErrM);" & vbNewLine & _
                            "End Zl_����ͼ����_Temp_Insert;"
            gcnOracle.Execute strSQL
        End If
    Else
        If Not CheckTblExist("���鱨��ͼ��_EXP_TEMP") Then
            strSQL = "Create Table ���鱨��ͼ��_EXP_TEMP As Select id,�걾id,ͼ������,ͼ��λ�� From ���鱨��ͼ�� Where 1=0"
            gcnOracle.Execute strSQL
            strSQL = "Create Or Replace Procedure Zl_���鱨��ͼ��_Temp_Insert" & vbNewLine & _
                            "(" & vbNewLine & _
                            "  Id_In       In ���鱨��ͼ��_Exp_Temp.Id%Type," & vbNewLine & _
                            "  �걾id_In   In ���鱨��ͼ��_Exp_Temp.�걾id%Type," & vbNewLine & _
                            "  ͼ������_In In ���鱨��ͼ��_Exp_Temp.ͼ������%Type," & vbNewLine & _
                            "  ͼ��λ��_In In ���鱨��ͼ��_Exp_Temp.ͼ��λ��%Type" & vbNewLine & _
                            ") Is" & vbNewLine & _
                            "Begin" & vbNewLine & _
                            "  Insert Into ���鱨��ͼ��_Exp_Temp Values (Id_In, �걾id_In, ͼ������_In, ͼ��λ��_In);" & vbNewLine & _
                            "Exception" & vbNewLine & _
                            "  When Others Then" & vbNewLine & _
                            "    zl_ErrorCenter(SQLCode, SQLErrM);" & vbNewLine & _
                            "End Zl_���鱨��ͼ��_Temp_Insert;"
            gcnOracle.Execute strSQL
        End If
    End If
    
End Sub

Private Sub SetCmdEnable(ByVal blnEnable As Boolean)
    cmdFtp.Enabled = blnEnable: cmdFile.Enabled = blnEnable
    cmdMulti.Enabled = blnEnable: cmdCommit.Enabled = blnEnable
End Sub
