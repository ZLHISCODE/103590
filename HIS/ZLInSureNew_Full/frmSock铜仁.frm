VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmSockͭ�� 
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5445
   ControlBox      =   0   'False
   Icon            =   "frmSockͭ��.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   5445
   StartUpPosition =   1  '����������
   Begin VB.Timer timUnload 
      Interval        =   1000
      Left            =   630
      Top             =   180
   End
   Begin MSWinsockLib.Winsock sckCenter 
      Left            =   210
      Top             =   180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lbl˵�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ҽ�����Ľ������ݽ�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   300
      TabIndex        =   0
      Top             =   780
      Width           =   4500
   End
End
Attribute VB_Name = "frmSockͭ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mint���� As Integer

Private mstr���ĵ�ַ As String
Private mbln����   As Boolean
Private mint����   As Integer
Private mstr������ As String   'д���Ĳ���̫�࣬�������Ƚ��鷳�����Ըɴ����ֱ�Ӵ���������ķ�ʽ

Public Function CommIC(ByVal ���ĵ�ַ As String, ByVal ���� As Boolean, ByVal ���� As Integer, ByVal ������ As String) As Boolean
'���ܣ������Ľ������ݽ�����ģ��IC������
'����������   ����ʱΪTrue,д����ΪFalse
'      ����   0 ��ʾ����;1 ��ʾסԺ

    '###############ȷ��֮ǰ����ʹ���κοؼ����������Load�¼�
    mbln���� = ����
    mint���� = False
    mblnOK = False
    mint���� = 0
    mstr���ĵ�ַ = ���ĵ�ַ
    mstr������ = ������
    '###############ȷ��֮ǰ����ʹ���κοؼ����������Load�¼�
    
    frmSockͭ��.Show vbModal
    CommIC = mblnOK
End Function

Private Sub Form_Load()
    On Error Resume Next
    sckCenter.RemoteHost = mstr���ĵ�ַ
    sckCenter.RemotePort = 1800
    Me.Visible = False
    sckCenter.Connect
End Sub

Private Sub sckCenter_Connect()
    Dim str�������� As String
    Dim strInput As String
    
    On Error Resume Next
    '���ɷ�������
    If mbln���� = True Then
        '��ѯ��Ϣ
        str�������� = "000"
    Else
        '����
        str�������� = "100"
    End If
    
    strInput = "*|" & mint���� & "|" & str�������� & "|" & LenB(StrConv(mstr������, vbFromUnicode)) & "|" & mstr������ & "|*"
    '�����Ѿ����������Է�������
    sckCenter.SendData strInput
    If Err <> 0 Then
        '���ִ��󣬿϶�������
        MsgBox "���ݷ���ʧ�ܡ�", vbInformation, gstrSysName
        Unload Me
    End If
End Sub

Private Sub sckCenter_DataArrival(ByVal bytesTotal As Long)
    Dim str������ As String, str�������� As String
    Dim strData As String, var���� As Variant
    
    On Error Resume Next
    '������������
    timUnload.Enabled = False '�������ó�ʱ��
    sckCenter.GetData strData, vbString
    If Err <> 0 Then
        '���ִ��󣬿϶�������
        MsgBox "���ݽ���ʧ�ܡ�", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    var���� = Split(strData, "|")
    If UBound(var����) < 6 Then
        MsgBox "���ݽ��ո�ʽ����", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    If var����(0) <> "*" Or var����(UBound(var����)) <> "*" Then
        MsgBox "���ݽ��ո�ʽ����", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    If var����(1) <> "10" Then
        If var����(1) = "11" Then
            If mint���� <> 1 Then
                MsgBox "��������סԺ�����ܼ�����", vbInformation, gstrSysName
                Unload Me
                Exit Sub
            End If
        Else
            MsgBox "���ݽ��ո�ʽ����", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
    
    If mbln���� = True Then
        '���Ĵ���|ְ������|����|�Ա�|��������|��λ����|ְ�����|���ش���|�Ƿ���Ա|�Ƿ�μӲ���|�ʻ��ۼ�ע��|�ʻ��ۼ�֧��|ͳ��֧�������ۻ�|ͳ��֧������ۻ�|��ЧסԺ����
        With gICͭ��Temp
            .CenterCode = var����(5)
            .Cardno = var����(6)          ' ����
            .IDCardno = Split(mstr������, "|")(0)       ' ���֤�� ���Ȳ����#0
            .MediAccountNo = var����(6)  ' ҽ����
            .Name = var����(7)           ' ����
            .Sex = var����(8)            ' �Ա� 1-��  0-Ů
            .Birthday = var����(9)       ' �������� YYYYMMDD
            .UnitCode = var����(10)       ' ���˵�λ����
            .ClassCode = var����(11)      ' ְ����ݣ�0x����ְ1x������, 05��11Ϊһ���Խɷ�
            .DomainCode = var����(12)     ' ְ������ 0-���� 1-��פ��� 2-��ذ���
            .MediYear = Year(zldatabase.Currentdate())       ' ҽ�����
            .InNo = 0           ' װǮ�ڴ�
            .OutSerialNo = 0    ' ֧��˳���
            .InPerAcc = var����(15)       ' �����ʻ��ۼ�ע����
            .OutPerAcc = var����(16)      ' �����ʻ��ۼ�֧�����
            .PlanPaidFee = var����(17)    ' ͳ�����֧�������ۼƣ�����+���䣩
            .PlanPaidAmt = var����(18)    ' ͳ�����֧������ۼƣ�����+���䣩
            .ChronicPaidFee = 0 ' ���Բ�֧�������ۼ�
            .ChronicPaidAmt = 0 ' ���Բ�֧������ۼ�
            .InHosPaidAmt = 0   ' סԺ�����ʻ�֧�����
            .ClinicPaidAmt = 0  ' ��������ʻ�֧�����
            .Password = Split(mstr������, "|")(1)      ' ��������
            .InHosTimes = var����(19)     ' ������ЧסԺ����
            .IsOffical = var����(13)      ' ����Ա 0-������-��
            .IsAttend = 0       ' ҽ���չ˶��� 0-��1-��
            .InpatientFlag = 0  ' סԺ��־ 0-��סԺ 1-סԺ
            .QuotaPaidAmt = 0   ' ���Բ������֧�����
            .ChronicSillPaidAmt = 0  ' ���Բ��𸶽���֧�����
        End With
    End If
    mblnOK = True
    Unload Me
End Sub

Private Sub sckCenter_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "���ӳ��ִ���" & Description, vbInformation, gstrSysName
    Unload Me
End Sub

Private Sub timUnload_Timer()
    mint���� = mint���� + 1
    If mint���� > 30 Then
        If mbln���� = True Then
            'ֻ�ж����ſ��Գ�ʱ�˳���д��ֻ�ܲ�ͣ�صȴ�
            If MsgBox("���ӳ�ʱ���Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Unload Me
                Exit Sub
            End If
        End If
        mint���� = 0
    End If
End Sub
