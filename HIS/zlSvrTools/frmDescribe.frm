VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDescribe 
   BackColor       =   &H80000005&
   Caption         =   "װж����"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmDescribe.frx":0000
   ScaleHeight     =   4275
   ScaleWidth      =   5445
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ils32 
      Left            =   270
      Top             =   2970
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescribe.frx":04F9
            Key             =   "K01"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescribe.frx":114B
            Key             =   "K02"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescribe.frx":1A25
            Key             =   "K03"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescribe.frx":22FF
            Key             =   "K04"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDescribe.frx":2BD9
            Key             =   "K05"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "װж����"
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
      Left            =   200
      TabIndex        =   1
      Top             =   100
      Width           =   960
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Height          =   3735
      Left            =   930
      TabIndex        =   0
      Top             =   645
      Width           =   4140
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmDescribe.frx":34B3
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmDescribe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstr��� As String    '������

Private Sub Form_Load()
    Select Case mstr���
        Case "01"
            lblTitle.Caption = "װж����"
            lblMain.Caption = "��װ���ж����Ӧ��ϵͳ�����ݷ�������" & _
                vbCrLf & vbCrLf & "������Ҫ������������һЩ����Ȩ�ޣ���Select on sys.v_$session��Select on sys.v_$parameter��Select on sys.dba_role_privs��Execute on sys.dbms_sql�������ֻ�о�����ЩȨ�޼���GRANTѡ���DBA�û���������װж����" & _
                vbCrLf & vbCrLf & "�������ݿ�ϵͳ��������ƶ�ϵͳ������Ч���нϴ��Ӱ�죬�����ϵͳ��װ֮ǰ���������������ļ�������������ϵͳ�ľ��������ȷ�滮��" & _
                vbCrLf & vbCrLf & "����װж�����漰�����ļ�����ɾ������뱣֤��û�������û�ʹ�õ�����½��С�"
        Case "02"
            lblTitle.Caption = "���ݹ���"
            lblMain.Caption = "��ָ��Ӧ��ϵͳ�����ݴ洢�������߼�������ָ���������װ���ı�������װ�صȲ�����" & _
                vbCrLf & vbCrLf & "��������������ݹ���Ĳ�������Ҫ���Ľϴ��ϵͳ��Դ��������ϲ�������������ϵͳ��Ϊ���е�ʱ����У��Խ��Ͷ�����������������Ӱ�죻" & _
                vbCrLf & vbCrLf & "�������ݹ������(���ݵ��롢�������)������ϵͳ���ݵĳ��������������ڴ�֮ǰ��֤�Ѿ����ڰ�ȫ�����ݱ��ݡ�" & _
                vbCrLf & vbCrLf & "���ڵ����ļ�(Export)�����ݵ���(Import)������װ��(Load)����ʹ�����ݿ������ܣ��뱣֤���ݿ������������ȷִ�У�" & _
                vbCrLf & vbCrLf & "�������ݵ�����װ��δ���汾�ĺϷ��ԣ���ȷ����ʵ��������ϣ��Ա�֤�����Ĳ�����"
        Case "03"
            lblTitle.Caption = "���й���"
            lblMain.Caption = "���ϵͳ����״̬�ļ�ء���ʷ��־�Ĳ鿴��δ������״̬�����á�" & _
                vbCrLf & vbCrLf & "��̨��ҵ���������ݿ���ҵ����ʵ�ֵ��Զ�����ʩ������Ҫʹ�ú�̨��ҵ����������ݿ�����init����(��JOB_QUEUE_PROCESSES��JOB_QUEUE_INTERVAL)��" & _
                vbCrLf & vbCrLf & "ͬ����Ҳ������ϵͳ�Ͽ���ʱִ�к�̨��ҵ���Լ��ٺ������������Դ������" & _
                vbCrLf & vbCrLf & "������־��������־�ļ�¼����Ҫռ��һ�������ݿռ����Դ����ע�⾭��������ʷ��־���ݣ����ϵͳ�Ѿ��ϳ�ʱ���ȶ����У�����ͨ��ѡ��رն���־�ļ�¼��"
        Case "04"
            lblTitle.Caption = "Ȩ�޹���"
            lblMain.Caption = "����ϵͳ�Ĺ��ܽ�ɫ�������������û�Ȩ�޲�ָ���û���ݣ���������˵���" & _
                vbCrLf & vbCrLf & "ϵͳ��ɫ�����ݿ�ϵͳinit����max_enabled_roles�����ƣ������Ҫ���ֽ����϶�Ľ�ɫ�����޸ĸ����ݲ����������������ݿ������Ч��" & _
                vbCrLf & vbCrLf & "Ϊ�˸��õؿ���Ȩ�ޣ���ϵͳ���û�Ҳ�������ݿ�ϵͳ���û��������뾭����ת�����벻Ҫ��ͼʹ�ñ�ϵͳ�������û�ֱ��ʹ�����ݿ�Ĺ������ӽ������ݿ⣻" & _
                vbCrLf & vbCrLf & "�Լ������Ľ�ɫȱʡ�Լ�Ҳ���иý�ɫ����һ���㲻��Ҫ�ý�ɫ��ȡ�����㽫��Ҳ�������ý�ɫ�Ĵ��ڣ�������DBA�û���"
        Case "05"
            lblTitle.Caption = "ר���"
            lblMain.Caption = "ʹ�ñ����߿����ϵͳ����Ʊ�ݸ�ʽ��������ݵĶ����޸ġ�" & _
                vbCrLf & vbCrLf & "�ù��߲����������Ĳ��ԣ����ص�ͼԪ���Ʒ�ʽ��ͼ��Ԫ�ص�ѡ��棩����ȷ����Ʊ���뱨�������������ص���Ʊ�ݵ�ֽ������(��С������)�������ʽ�����塢��ɫ�����У�����������Ԥ����ӡ��" & _
                vbCrLf & vbCrLf & vbCrLf & vbCrLf & "ʹ�ú���������ɸ�ϵͳ���ݴ��ݺ����Ĺ������������ı���������򵼵Ķ��塢�޸������á�" & _
                vbCrLf & vbCrLf & "���ݴ��ݺ����Ǳ������Ӧ��ϵͳ���໥��ѡ�������ݵ���Ҫ��ʽ��ʹ����Ӧ�ó�Ϊһ�����������壻�϶��Ӧ���ڲ��������Զ�ƾ֤���ɱ�Ч�����ͱ��������ȡ��Ӧ��ϵͳ�ķ������ݡ�"
    End Select
    Me.Caption = lblTitle.Caption
    imgMain.Picture = ils32.ListImages("K" & mstr���).Picture
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    
    With lblMain
        .Top = imgMain.Top
        .Height = ScaleHeight - .Top * 2
        .Left = imgMain.Left * 2 + imgMain.Width
        .Width = ScaleWidth - .Left - imgMain.Left
    End With
End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub
