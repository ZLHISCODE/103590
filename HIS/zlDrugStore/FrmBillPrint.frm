VERSION 5.00
Begin VB.Form FrmBillPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ʽѡ��"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "FrmBillPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Cmd�����Excel 
      Caption         =   "�����&Excel"
      Height          =   350
      Left            =   2970
      TabIndex        =   7
      Top             =   1710
      Width           =   1425
   End
   Begin VB.CommandButton CmdԤ�� 
      Caption         =   "Ԥ��(&R)"
      Height          =   350
      Left            =   1410
      TabIndex        =   6
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton Cmd��ӡ 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   210
      TabIndex        =   5
      Top             =   1710
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Enabled         =   0   'False
      Height          =   1395
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   4275
      Begin VB.TextBox Txt���ݺ� 
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         Top             =   810
         Width           =   1935
      End
      Begin VB.TextBox Txt�������� 
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   390
         Width           =   1935
      End
      Begin VB.Label Lbl���ݺ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   750
         TabIndex        =   3
         Top             =   870
         Width           =   810
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   570
         TabIndex        =   1
         Top             =   450
         Width           =   990
      End
   End
End
Attribute VB_Name = "FrmBillPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngϵͳ�� As String
Private strƱ�ݺ� As String
Private str�������� As String
Private int�������� As Integer
Private str���ݺ� As String
Private lng��¼״̬ As Long
Private int��λϵ�� As Integer
Private mint����ģʽ As Integer
Private Sub Cmd��ӡ_Click()
    Call BillPrint(2)
End Sub

Private Sub Cmd�����Excel_Click()
    Call BillPrint(3)
End Sub

Private Sub CmdԤ��_Click()
    Call BillPrint(1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Public Function ShowME(ByVal frmParent As Object, ByVal ϵͳ�� As Long, ByVal Ʊ�ݺ� As String, _
                       ByVal ��¼״̬ As Long, ByVal ��λϵ�� As Integer, ByVal �������� As Integer, ByVal �������� As String, ByVal ���ݺ� As String, Optional ByVal int����ģʽ As Integer = 0)
    lngϵͳ�� = ϵͳ��
    strƱ�ݺ� = Ʊ�ݺ�
    lng��¼״̬ = ��¼״̬
    int��λϵ�� = ��λϵ��
    str�������� = ��������
    int�������� = ��������
    str���ݺ� = ���ݺ�
    mint����ģʽ = int����ģʽ
    Me.Show 1, frmParent
End Function

Private Sub BillPrint(ByVal intPrintMode As Integer)
    Select Case int��������
'    Case 1300           'ҩƷ�⹺������
'        ReportOpen gcnOracle, lngϵͳ��, strƱ�ݺ�, Me, "���ݱ��=" & str���ݺ�, "��¼״̬=" & lng��¼״̬, "��λϵ��=" & int��λϵ��, intPrintMode
'    Case 1301           'ҩƷ����������
'        ReportOpen gcnOracle, lngϵͳ��, strƱ�ݺ�, Me, "���ݱ��=" & str���ݺ�, "��¼״̬=" & lng��¼״̬, "��λϵ��=" & int��λϵ��, intPrintMode
'    Case 1302           'ҩƷ����������
'        ReportOpen gcnOracle, lngϵͳ��, strƱ�ݺ�, Me, "���ݱ��=" & str���ݺ�, "��¼״̬=" & lng��¼״̬, "��λϵ��=" & int��λϵ��, intPrintMode
'    Case 1303           '����۵�������
'        ReportOpen gcnOracle, lngϵͳ��, strƱ�ݺ�, Me, "���ݱ��=" & str���ݺ�, "��¼״̬=" & lng��¼״̬, "��λϵ��=" & int��λϵ��, intPrintMode
'    Case 1304           'ҩƷ�ƿ����
'        ReportOpen gcnOracle, lngϵͳ��, strƱ�ݺ�, Me, "���ݱ��=" & str���ݺ�, "��¼״̬=" & lng��¼״̬, "��λϵ��=" & int��λϵ��, intPrintMode
'    Case 1305           'ҩƷ���ù���
'        ReportOpen gcnOracle, lngϵͳ��, strƱ�ݺ�, Me, "���ݱ��=" & str���ݺ�, "��¼״̬=" & lng��¼״̬, "��λϵ��=" & int��λϵ��, intPrintMode
'    Case 1306           'ҩƷ�����������
'        ReportOpen gcnOracle, lngϵͳ��, strƱ�ݺ�, Me, "���ݱ��=" & str���ݺ�, "��¼״̬=" & lng��¼״̬, "��λϵ��=" & int��λϵ��, intPrintMode
'    Case 1307           'ҩƷ�̵����
'        ReportOpen gcnOracle, lngϵͳ��, strƱ�ݺ�, Me, "���ݱ��=" & str���ݺ�, "��¼״̬=" & lng��¼״̬, "��λϵ��=" & int��λϵ��, intPrintMode
    Case 1300
        ReportOpen gcnOracle, lngϵͳ��, strƱ�ݺ�, Me, "���ݱ��=" & str���ݺ�, "��¼״̬=" & lng��¼״̬, "��λϵ��=" & int��λϵ��, str��������, intPrintMode
    Case 1301, 1302, 1303, 1304, 1305, 1306, 1307, 1344
        ReportOpen gcnOracle, lngϵͳ��, strƱ�ݺ�, Me, "���ݱ��=" & str���ݺ�, "��¼״̬=" & lng��¼״̬, "��λϵ��=" & int��λϵ��, intPrintMode
    Case 1320           'ҩƷ�������
        ReportOpen gcnOracle, lngϵͳ��, strƱ�ݺ�, Me, "���ݱ��=" & str���ݺ�, "��¼״̬=" & lng��¼״̬, intPrintMode
    Case 1330           'ҩƷ�ƻ�����
        ReportOpen gcnOracle, lngϵͳ��, strƱ�ݺ�, Me, "���ݱ��=" & str���ݺ�, IIf(mint����ģʽ = 0, "ReportFormat=1", IIf(mint����ģʽ = 1, "ReportFormat=2", "ReportFormat=3")), intPrintMode
    End Select
    If intPrintMode <> 1 Then Unload Me
End Sub

Private Sub Form_Load()
    Me.Txt�������� = str��������
    Me.Txt���ݺ� = str���ݺ�
End Sub
