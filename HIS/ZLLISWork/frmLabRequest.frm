VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.10#0"; "zlIDKind.ocx"
Begin VB.Form frmLabRequest 
   BackColor       =   &H00FDD6C6&
   BorderStyle     =   0  'None
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txt����1 
      Height          =   300
      IMEMode         =   2  'OFF
      Left            =   2790
      MaxLength       =   5
      TabIndex        =   38
      Top             =   630
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   1620
      ScaleHeight     =   330
      ScaleWidth      =   1770
      TabIndex        =   36
      Top             =   6450
      Width           =   1770
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   2
         Left            =   45
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483643
         CustomFormat    =   "yy-MM-dd HH:mm:ss"
         Format          =   60424195
         CurrentDate     =   38222
      End
   End
   Begin VB.ComboBox cbo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   540
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   6450
      Width           =   1110
   End
   Begin VB.ComboBox cboҽ�� 
      Height          =   300
      Left            =   2145
      TabIndex        =   11
      Top             =   4995
      Width           =   1155
   End
   Begin VB.ComboBox cbo�������� 
      Height          =   300
      ItemData        =   "frmLabRequest.frx":0000
      Left            =   540
      List            =   "frmLabRequest.frx":0002
      TabIndex        =   10
      Top             =   4995
      Width           =   1590
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "��"
      Height          =   1110
      Left            =   3030
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "ѡ����Ŀ(*)"
      Top             =   2010
      Width           =   300
   End
   Begin VB.CommandButton cmdExt 
      Caption         =   "��"
      Height          =   255
      Left            =   3015
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "ѡ�����걾"
      Top             =   5385
      Width           =   255
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Index           =   0
      Left            =   930
      TabIndex        =   15
      Top             =   5700
      Width           =   2370
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Index           =   1
      Left            =   540
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6060
      Width           =   1110
   End
   Begin VB.ComboBox cbo�Ա� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      ItemData        =   "frmLabRequest.frx":0004
      Left            =   540
      List            =   "frmLabRequest.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   630
      Width           =   795
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1785
      MaxLength       =   3
      TabIndex        =   2
      Top             =   630
      Width           =   345
   End
   Begin VB.TextBox txtPatientDept 
      Enabled         =   0   'False
      Height          =   300
      Left            =   975
      MaxLength       =   24
      TabIndex        =   6
      Top             =   1365
      Width           =   1785
   End
   Begin VB.TextBox txtID 
      Enabled         =   0   'False
      Height          =   300
      Left            =   735
      Locked          =   -1  'True
      MaxLength       =   18
      TabIndex        =   4
      Top             =   990
      Width           =   1455
   End
   Begin VB.TextBox txtBed 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2730
      MaxLength       =   10
      TabIndex        =   5
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox txtҽ������ 
      Height          =   1080
      IMEMode         =   2  'OFF
      Left            =   135
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2010
      Width           =   2865
   End
   Begin VB.ComboBox cboAge 
      Height          =   300
      IMEMode         =   3  'DISABLE
      ItemData        =   "frmLabRequest.frx":0008
      Left            =   2145
      List            =   "frmLabRequest.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   630
      Width           =   630
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&P"
      Height          =   465
      Left            =   2985
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   90
      Width           =   300
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   1
      Left            =   930
      TabIndex        =   8
      Top             =   3165
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   60424195
      CurrentDate     =   38222
   End
   Begin zl9LisWork.VsfGrid vsf2 
      Height          =   1125
      Left            =   75
      TabIndex        =   9
      Top             =   3795
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   1984
   End
   Begin VB.TextBox txt���� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   810
      MaxLength       =   64
      TabIndex        =   0
      ToolTipText     =   "��������ͷΪ����ID��������סԺ�š���*������š���.���Һŵ��š���/���շѵ��ݺ�"
      Top             =   90
      Width           =   2475
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   5355
      Width           =   2370
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   1620
      ScaleHeight     =   330
      ScaleWidth      =   1770
      TabIndex        =   33
      Top             =   6060
      Width           =   1770
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   45
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483643
         CustomFormat    =   "yy-MM-dd HH:mm:ss"
         Format          =   60424195
         CurrentDate     =   38222
      End
   End
   Begin zlIDKind.IDKind IDKind 
      Height          =   420
      Left            =   135
      TabIndex        =   37
      Top             =   105
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   741
      IDKindStr       =   "��|����|0;ҽ|ҽ����|1;��|���֤��|2;IC|IC����|3;��|�����|4;��|���￨|5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl���δͨ�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   180
      TabIndex        =   40
      Top             =   6870
      Width           =   165
   End
   Begin VB.Label lblRegister 
      BackColor       =   &H00FDD6C6&
      Caption         =   "�����жϵǼ�"
      Height          =   225
      Left            =   2220
      TabIndex        =   39
      Top             =   3540
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   35
      Top             =   6510
      Width           =   360
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   135
      TabIndex        =   34
      Top             =   5055
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(���걾�ֽ�):"
      Height          =   180
      Left            =   135
      TabIndex        =   32
      Top             =   3570
      Width           =   1530
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�걾����"
      Height          =   180
      Left            =   135
      TabIndex        =   31
      Top             =   5430
      Width           =   720
   End
   Begin VB.Label lblҽ������ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������Ŀ"
      Height          =   180
      Left            =   135
      TabIndex        =   30
      Top             =   1785
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�걾��̬"
      Height          =   225
      Index           =   5
      Left            =   135
      TabIndex        =   29
      Top             =   5775
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��"
      Height          =   180
      Index           =   6
      Left            =   135
      TabIndex        =   28
      Top             =   3240
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   27
      Top             =   6120
      Width           =   360
   End
   Begin VB.Label lblCash 
      BackColor       =   &H00FDD6C6&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2970
      TabIndex        =   26
      Top             =   1380
      Width           =   300
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   2325
      TabIndex        =   25
      Top             =   1050
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�"
      Height          =   180
      Left            =   135
      TabIndex        =   24
      Top             =   690
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   1380
      TabIndex        =   23
      Top             =   690
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���˿���"
      Height          =   180
      Left            =   135
      TabIndex        =   22
      Top             =   1425
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʶ��"
      Height          =   180
      Left            =   135
      TabIndex        =   21
      Top             =   1050
      Width           =   540
   End
End
Attribute VB_Name = "frmLabRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintEditMode As Integer, ItemDeptID As Long, mlngDefaultDevice As Long
'------------��ʱδ��------------
Private mlngSampleID As Long, mintSampleType As Integer
'-------------------------------
Private PatientType As Integer, mlng����ID As Long, mstrNO As String '�����շѵ��ݺ�
Private mlngDefaultItemID  As Long  '�ϼ�Ĭ�ϵ�������ĿID
Private mstrAuditer As String
Private iInputType As Integer
Private mblnEmerge As Boolean       '�Ƿ�ʹ�ü���걾
Private mblnPrice As Boolean        '�Ƿ��շѵ��ź���

'����������ǰ����״̬�����һֱ�Ը�״̬���Բ�����ǰ����
'0�����￨
'1������ID
'2��סԺ��
'3�������
'4���Һŵ�
'5���շѵ��ݺ�
'6������
Private rsRelativeAdvice As ADODB.Recordset '�Ǽǵ����ҽ��
Private mstrExtData  As String '�Ǽǵ�������Ŀ��Ϣ
Private mlngCapID As Long '�ɼ���ĿID
Private mstrKeys As String '��ǰ���յ�����ҽ��ID
Private mlngReqDept As Long, mstrReqDoctor As String  'Ĭ�ϵĵǼǿ��Һ�ҽ��
Private mstrPrivs As String   'Ȩ��

Private mblnBarCode As Boolean
Private mblnSaveAdvice As Boolean '�Ƿ���Ҫ����ҽ���������޸���Ժ���˱걾��Ϣ

Private mbln΢������Ŀ As Boolean
Private mlngNoneHomeKey() As Long     '���ҵ���Ҫ�����ǵı걾ID
Private mlngSourceKey() As Long       '����ʱ��Ҫ�����ǵı걾ID
Private mRsSex As New ADODB.Recordset '�Ա�ļ�¼��
Private mstrNONumber As String        '��¼��ǰ¼�������һ���걾��
Private mblnCheckIn As Boolean        '�Ǽ��ǿ��Բ�������Ŀ
Public mMakeNoRule As String          '�걾������ɹ���
Private mintItemRule As Integer       '�Ƿ���Ŀ�ۼӵķ�ʽ�����ɱ걾��
Private mblnCard As Boolean           '�Ƿ�ˢ��
Private mstrMachines As String        '���Բ���������ID
Private mSendReport As Integer        '��˺��Ƿ��Զ����ͱ��� 0=���� 1=������
Private mstr������ As String          '������
Private mbln���۵�ģʽ As Boolean     '�Ƿ�ʹ�û��۵�ģʽ
Private mblnEdit As Boolean           '�Ƿ����ڱ༭
Private mstr�������� As String        'ֻ�ܰ������������������������м�ʹ��","�ָ�)
Private mbln���� As Boolean           '��ǰ���յı걾�Ƿ��Ǽ���걾
Private mblnLoadLastAdvice As Boolean '�Ƿ�Ĭ���ϴεǼ���Ŀ��Ϊ����Ǽǵ�Ĭ����Ŀ
Private mblnShowPwd As Boolean                                          '�Ƿ���ʾ����

'ָ��������
Private Enum ItemCol
    ID = 0
    ���ID
    ���
    ��־
    ����ο�
    ������ĿID
    �������
    ����
    ѡ��
End Enum
Private Enum mCol
    ID = 0
    ����
    ����ҽ��
    ִ��״̬
    �������
    �걾����
    �걾��
    ����
    �Ա�
    ����
    ������Ŀ
    ��ʶ��
    ����
    �������
    ҽ��id
    ����id
    ת��
    ����ID
    �걾ʱ��
    ����ʱ��
    ΢����걾
    �շѵ�
    �Һŵ�
    ������
    �����
    ��������
    Ӥ��
    ���˿���
    ���ͺ�
    ������
    ��ҳID
    ��������ID
    ������
    ��������
    ���䵥λ
    ����
    ������
    �걾��̬
    ������
    ����ʱ��
    ����걾
    NO
    ������
    ����ʱ��
    ���ʱ��
    ����id
    ��������
    ��λ
    ִ�п���ID
    �걾���
    ҽ������
    �걾����
    �������
    ��������
    ����
    ����״̬
    ���淢��
    ���˿���ID
    ������
    ����ʱ��
    ��λ
    ������
    ���δͨ��
    ������Դ
    �����
    סԺ��
End Enum

'������յ��Զ������¼�
Public Event ZlAutoSave(ByVal lngSampleID As Long)

'-------------------------------------------- 2007-10-26 ����һ��֧ͨ��
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private Enum IDKinds
    C0���� = 0
    C1ҽ���� = 1
    C2���֤�� = 2
    C3IC���� = 3
    C4����� = 4
    C5���￨ = 5
End Enum
Private mobjSquareCard As Object                                        'ȡ������

Private Sub Txt����Exec()
    Dim strInput As String
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strField As String
    Dim strBarCode As String
    Dim rsDept As ADODB.Recordset, strSQL As String
    Dim intSelect As Integer
    Dim intPatientSource As Integer                     '������Դ
    Dim rsPaInfo As New ADODB.Recordset                 '��ȡ�걾��¼�е�"��ʶ��"
    Dim blnGetPaInfo As Boolean
    Dim strAge As String
    Dim aAge() As String
    Dim strҽ��ID As String
    Dim rs As New ADODB.Recordset
    
    Dim intMainID   As Integer                          '��ҳid
    Dim strGetSql As String
    Dim rsTest As Recordset
    
    On Error GoTo errH
    If Len(Trim(txt����)) = 0 Or Me.txt����.Enabled = True Or Me.txt���� = Me.txt����.Tag Then Exit Sub

    If txt���� <> txt����.Tag Then mlng����ID = 0
    If mlng����ID > 0 Then
        strSQL = "select ������Դ from ����걾��¼ where id = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngSampleID)
        If rsTmp.RecordCount > 0 Then
            intPatientSource = Nvl(rsTmp("������Դ"), PatientType)
        Else
            intPatientSource = PatientType
        End If

        If txt���� = txt����.Tag And intPatientSource <> 3 Then Exit Sub
    Else
        If txt���� = txt����.Tag Then Exit Sub
    End If

    mblnSaveAdvice = True
'    Cancel = Not StrIsValid(txt����.Text, txt����.MaxLength)
    If StrIsValid(txt����.Text, txt����.MaxLength) = False Then Exit Sub


    '��ʼ������Ϣ
    '2007-10-26 ����һ��֧ͨ��
    If IDKind.Tag = "ҽ����" Or IDKind.IDKind = IDKinds.C1ҽ���� Then
        Set rsTmp = GetPatient("��" & txt����)
        IDKind.Tag = ""
    ElseIf IDKind.Tag = "���֤" Or IDKind.IDKind = IDKinds.C2���֤�� Then
        Set rsTmp = GetPatient("֤" & txt����)
        IDKind.Tag = ""
    Else
        Set rsTmp = GetPatient(txt����)
    End If
    If rsTmp.RecordCount > 0 Then
        If iInputType = 2 Then
            If intMainID = 0 Then
                intMainID = Val(rsTmp("��ҳID") & "")
                mlng����ID = Val(rsTmp("����ID") & "")
                If intMainID <> 0 Then
                    strGetSql = "Select ��ҳid, ��Ժ���� From ������ҳ Where ����id = [1] and ��ҳid=[2] "
                    Set rsTest = zlDatabase.OpenSQLRecord(strGetSql, Me.Caption, mlng����ID, intMainID)
                    If Nvl(rsTest("��Ժ����")) <> "" Then
                        If MsgBox("�ò����ѳ�Ժ���Ƿ����ִ�У�", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                            '�����ʾ
                            mlng����ID = 0
                            mstrKeys = ""
                            Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
                            Me.txtID = "": Me.txtBed = ""
                            Me.txt����.Text = "": 'Cancel = True
                            Me.txt����.Enabled = True:             Me.txt����.SetFocus
                            Me.txtҽ������ = ""
                            Me.txt���� = ""
                            Me.txtҽ������.Tag = "": Me.txt����.Tag = "": mlngCapID = 0
                            Me.txt����.Text = "": Me.txt����1.Text = ""
                            vsf2.Rows = 1
                            vsf2.Rows = 2
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If

    If rsTmp.EOF = True And IDKind.IDKind = IDKinds.C0���� Then
        Set rsTmp = GetPatientInfo(txt����)
        blnGetPaInfo = True
    End If

    strBarCode = txt����
    If rsTmp.RecordCount <= 0 Then
        mlng����ID = 0
        '�Ǽ��²���
        mstrKeys = ""
'        Me.txt���� = "": Me.cboAge.ListIndex = 0
        Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
        Me.txtID = "": Me.txtBed = ""
        '���������Ժ�ڲ��ˣ����������
        If InStr("+-*./", Left(Me.txt����.Text, 1)) > 0 Or mblnBarCode Then
            Me.txt����.Text = "": 'Cancel = True
            Me.txt����.Enabled = True:             Me.txt����.SetFocus
            Exit Sub
        End If
        If mblnBarCode = True Then
            MsgBox "û���ҵ����룬��鿴�����Ƿ���ȡ���󶨣�", vbInformation, Me.Caption
            Me.txt����.Text = "": 'Cancel = True
            Me.txt����.Enabled = True:             Me.txt����.SetFocus
            Exit Sub
        End If
        If IDKind.IDKind = IDKinds.C0���� Then
            PatientType = 1
            '����Ǽǵ�Ĭ�Ͽ��ҡ�ҽ��
            If mlngReqDept > 0 Then
                cbo��������.ListIndex = FindComboItem(cbo��������, mlngReqDept)
                Me.cboҽ��.Text = mstrReqDoctor
            End If
            SetPatientInfoWrite False
            Me.txt����.Tag = Me.txt����.Text
        Else
            Select Case IDKind.IDKind
                Case IDKinds.C1ҽ����
                    MsgBox "û���ҵ�ҽ����Ϊ<" & Me.txt����.Text & ">�Ĳ��ˣ�", vbInformation, Me.Caption
                Case IDKinds.C2���֤��
                    MsgBox "û���ҵ����֤Ϊ<" & Me.txt����.Text & ">�Ĳ��ˣ�", vbInformation, Me.Caption
                Case IDKinds.C3IC����
                    MsgBox "û���ҵ�IC����Ϊ<" & Me.txt����.Text & ">�Ĳ��ˣ�", vbInformation, Me.Caption
                Case IDKinds.C4�����
                    MsgBox "û���ҵ������Ϊ<" & Me.txt����.Text & ">�Ĳ��ˣ�", vbInformation, Me.Caption
                Case IDKinds.C5���￨
                    MsgBox "û���ҵ����￨��Ϊ<" & Me.txt����.Text & ">�Ĳ��ˣ�", vbInformation, Me.Caption
            End Select
            Me.txt����.Text = "": 'Cancel = True
            Me.txt����.Enabled = True:             Me.txt����.SetFocus
            Exit Sub
        End If
    Else
        On Error Resume Next
        Me.txt����.Text = Nvl(rsTmp("����"))
'        Me.txt���� = IIf(IsNull(rsTmp("����")), "", Val(rsTmp("����"))): If Me.txt���� = "0" Then Me.txt���� = ""
'        Me.cboAge.Text = IIf(IsNull(rsTmp("����")), "��", Replace(rsTmp("����"), Val(rsTmp("����")), ""))
        If Trim(Nvl(rsTmp("����"))) <> "" And Trim(Nvl(rsTmp("����1"))) <> "" Then
            If rsTmp("����") <> rsTmp("����1") Then
'                MsgBox "����ͳ������ڼ�������䲻����" & _
                        vbCrLf & "�������ڼ�������Ϊ:" & rsTmp("����1") & _
                        vbCrLf & "��ǰ����Ϊ:" & rsTmp("����")
                Me.txt����.ForeColor = vbRed
            End If
        End If
        Me.txt����.Text = "": Me.txt����1.Text = ""
        '��ʹ���Զ����������,ʹ���´�ҽ��������
        'strAge = IIf(Trim(Nvl(rsTmp("����1"))) = "", Nvl(rsTmp("����")), Nvl(rsTmp("����1")))
        strAge = Nvl(rsTmp("����"))
        
        strAge = Replace(strAge, "Сʱ", "ʱ")
        strAge = Replace(strAge, "����", "��")
        
        If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "��", ""), "��", ""), "��", ""), "ʱ", ""), "��", "")) <> "" Then
            If InStr(strAge, "����") > 0 Or InStr(strAge, "Ӥ��") > 0 Then
                Me.txt����.Text = ""
                Me.cboAge.Text = Trim(strAge)
            Else
                strAge = Replace(Replace(Replace(Replace(Replace(strAge, "��", "��;"), "��", "��;"), "��", "��;"), "ʱ", "ʱ;"), "��", "��;")
                'strAge = Replace(strAge, "����", "Ӥ��")
                aAge = Split(strAge, ";")
                If UBound(aAge) = 1 Then
                    Me.txt����.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
                Else
                    Me.txt����.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
                    Me.txt����1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "��", "����"), "ʱ", "Сʱ")
                End If
            End If
        Else
            If Val(strAge) <> 0 Then
                Me.txt����.Text = Val(strAge)
            End If
            Me.cboAge.ListIndex = 0
        End If
        Me.txt����.Tag = Me.txt����.Text
'        Me.txt���� = IIf(IsNull(rsTmp("����1")), "", IIf(IsNumeric(rsTmp("����1")), Val(rsTmp("����1")), Mid(rsTmp("����1"), 1, Len(rsTmp("����1")) - 1)))
'        If Me.txt���� = "0" Then Me.txt���� = ""
'        Me.cboAge.Text = IIf(IsNull(rsTmp("����")), "��", Right(rsTmp("����1"), 1))
        If cboAge.ListIndex = -1 Then cboAge.ListIndex = 0
        Me.cbo�Ա� = Nvl(rsTmp("�Ա�")) ' CombIndex(cbo�Ա�, Nvl(rsTmp("�Ա�")))
        mlng����ID = Nvl(rsTmp("����ID"), 0): PatientType = Nvl(rsTmp("PatientType"), 1)

        '����Ĭ�Ͽ������ҡ�ҽ��
        cbo��������.ListIndex = FindComboItem(cbo��������, Nvl(rsTmp("���˿���"), 0))
'        DoEvents
        gintSelectFocus = 2
        strField = ""
        strField = rsTmp.Fields("ҽ��").Name
        If strField = "ҽ��" Then
            Me.cboҽ��.Text = Nvl(rsTmp("ҽ��"))
            For i = 0 To Me.cboҽ��.ListCount - 1
                If Me.cboҽ��.List(i) Like Nvl(rsTmp("ҽ��")) Then
                    Me.cboҽ��.ListIndex = i
                    Exit For
                End If
            Next
        End If
        '��ʾ���˿���
        If IsNumeric(rsTmp("���˿���")) = True Then
            strSQL = "Select ���� From ���ű� Where ID=[1]"
            Set rsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(Nvl(rsTmp("���˿���"), 0)))
            If rsDept.EOF Then
                Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
            Else
                Me.txtPatientDept.Text = rsDept("����"): Me.txtPatientDept.Tag = Nvl(rsTmp("���˿���"), 0)
            End If
        Else
            Me.txtPatientDept = Nvl(rsTmp("���˿���"))
        End If

        Me.txtID = Nvl(rsTmp("סԺ��")): If Len(Me.txtID) = 0 Then Me.txtID = Nvl(rsTmp("�����"))

        Me.txtBed = Nvl(rsTmp("��ǰ����"))

        '������걾��¼����ʱ������ʾ����걾��¼�����Ϣ
        If Trim(Me.txtID) = "" Or Trim(Me.txtBed) = "" Or Trim(Me.txtPatientDept) = "" Then
            If blnGetPaInfo = True Then
                Me.txtID = Nvl(rsTmp("��ʶ��"), Me.txtID)
                Me.txtBed = Nvl(rsTmp("����"), Me.txtBed)
                Me.txtPatientDept = Nvl(rsTmp("���˿���"), Me.txtPatientDept)
            End If
        End If

        '����Ǽǵ�Ĭ�Ͽ��ҡ�ҽ��
        If Me.cbo��������.ListIndex = -1 And mlngReqDept > 0 Then
            cbo��������.ListIndex = FindComboItem(cbo��������, mlngReqDept)
            Me.cboҽ��.Text = mstrReqDoctor
        End If
    End If
    '����ʱѡ���������
    If mlng����ID > 0 And Not mintEditMode = 1 And (intPatientSource <> 3 Or mblnBarCode = True) Then
        intSelect = OpenSelect(strBarCode, True)
'        DoEvents
        gintSelectFocus = 2
        Select Case intSelect
            Case 0
                'û��ƥ�����Ŀ
                mstrKeys = ""
                If mlng����ID = 0 Or mblnBarCode Then
'                    mintFocusItem = FocusItem.����
'                    MsgBox "�����ѱ����գ�", vbInformation, Me.Caption
                    mlng����ID = 0
                    txt����.Text = ""
                    If Me.txt����.Enabled = True Then
                        txt����.SetFocus
                    End If
'                    Cancel = True
                Else
                    '����Ǽ�
                    SetAdviceEnable True
                    If mlngDefaultItemID > 0 Then
                        strSQL = "select �걾��λ from ������ĿĿ¼ where id = [1] "
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngDefaultItemID)
                        AdviceSet������� 3, mlngDefaultItemID & ";" & rsTmp("�걾��λ")
                        mstrExtData = mlngDefaultItemID & ";" & rsTmp("�걾��λ")
                        '��ȡ�ɼ���ʽ
                        Set rsTmp = SelectCap(Split(Split(mstrExtData, ";")(0), ",")(0))
                        If rsTmp Is Nothing Then
                            MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        mlngCapID = rsTmp("ID")
                    End If
                    If mblnLoadLastAdvice = True And mstrExtData <> "" Then
                        AdviceSet������� 3, mstrExtData
                    End If
                    If rsRelativeAdvice Is Nothing Then
                        Me.txtҽ������ = ""
                        Me.txt���� = ""
                        Me.txtҽ������.Tag = "": Me.txt����.Tag = "": mlngCapID = 0
                    Else
                        txtҽ������.Text = Get�����������(2, "")
                        txtҽ������.Text = txtҽ������.Text & "(" & Split(mstrExtData, ";")(1) & ")"
                        txtҽ������.Tag = txtҽ������.Text
                        Me.txt���� = Split(mstrExtData, ";")(1)
                        If mintEditMode = 0 Then
                            Call LoadDefaultData
                            Call SelectDefault
'                            mintFocusItem = FocusItem.�걾��

'                            DoEvents
                            vsf2.Col = 2
                            vsf2.ShowCell vsf2.Row, vsf2.Col
                            vsf2.SetFocus
                            gintSelectFocus = 2
                        Else
                            '��ҽ��ID
                            Call SelectDefault
                        End If
                    End If
                End If
            Case 1
                'ѡȡ��һ����Ŀ
                SetAdviceEnable False   '������Ǽ�
                '���������ͱ걾��
                If mintEditMode = 0 Then
                    Call LoadDefaultData
                    Call SelectDefault
'                    mintFocusItem = FocusItem.�걾��

                    vsf2.Col = 2
                    vsf2.ShowCell vsf2.Row, vsf2.Col
                    vsf2.SetFocus
                Else
                    '��ҽ��ID
                    Call SelectDefault
                    vsf2.Col = 2
                    vsf2.ShowCell vsf2.Row, vsf2.Col
                    vsf2.SetFocus
                End If
                '�����Զ�����
                If mblnBarCode Then Me.cbo�Ա�.SetFocus
            Case 2
                'ȡ���˱���ѡ��
'                mintFocusItem = FocusItem.����

                mlng����ID = 0
                mstrKeys = ""
                txt����.Text = ""
                If Me.txt����.Enabled = True Then
                    txt����.SetFocus
                End If
'                Cancel = True
            Case 3
                Me.txt����.Enabled = True
                txt����.SetFocus
        End Select
    Else
        SetAdviceEnable True
        If mlngDefaultItemID > 0 Then
            strSQL = "select �걾��λ from ������ĿĿ¼ where id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngDefaultItemID)
            AdviceSet������� 3, mlngDefaultItemID & ";" & rsTmp("�걾��λ")
            mstrExtData = mlngDefaultItemID & ";" & rsTmp("�걾��λ")
            '��ȡ�ɼ���ʽ
            Set rsTmp = SelectCap(Split(Split(mstrExtData, ";")(0), ",")(0))
            If rsTmp Is Nothing Then
                MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName
                Exit Sub
            End If
            mlngCapID = rsTmp("ID")
        End If

        If mblnLoadLastAdvice = True And mstrExtData <> "" Then
            AdviceSet������� 3, mstrExtData
        End If

        If rsRelativeAdvice Is Nothing Then
            Me.txtҽ������ = ""
            Me.txt���� = ""
            Me.txtҽ������.Tag = "": Me.txt����.Tag = "": mlngCapID = 0
        Else
            txtҽ������.Text = Get�����������(2, "")
            txtҽ������.Text = txtҽ������.Text & "(" & Split(mstrExtData, ";")(1) & ")"
            txtҽ������.Tag = txtҽ������.Text
            Me.txt���� = Split(mstrExtData, ";")(1)
            If mintEditMode <= 1 Then
                Call LoadDefaultData
                Call SelectDefault

                vsf2.Col = 2
                vsf2.ShowCell vsf2.Row, vsf2.Col
                vsf2.SetFocus
            Else
                '��ҽ��ID
                Call SelectDefault
            End If
        End If
    End If

    If mlng����ID > 0 Then
        txt����.Tag = txt����.Text
    End If

    If mblnCheckIn = True And Me.txtҽ������.Tag = "" And mintEditMode <> 3 Then
        If mlngDefaultDevice > 0 And mintEditMode <> 3 Then
            gstrSql = "select ���� from �������� where id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDefaultDevice)
            vsf2.TextMatrix(1, 1) = Nvl(rsTmp("����"))
            vsf2.RowData(1) = mlngDefaultDevice
        Else
            vsf2.TextMatrix(1, 1) = "[�ֹ�]"
            vsf2.RowData(1) = mlngDefaultDevice
        End If
        'ȡ�걾��
        If vsf2.TextMatrix(1, 5) = "-1" Then
            '����
            vsf2.TextMatrix(1, 2) = TransSampleNO_PH(Val(CalcNextCode(Val(vsf2.RowData(1)), 1, 1)), vsf2.RowData(1))
        Else
            vsf2.TextMatrix(1, 2) = TransSampleNO_PH(Val(CalcNextCode(Val(vsf2.RowData(1)), 1, 0)), vsf2.RowData(1))
        End If
    End If

    '--------------------------------------------------------------------------------------------------------------------------------
    '��ִ��������Զ���˵ķ���ʱ���Բ��˷��ý��м��ʱ�����
    gstrSql = " select /*+ rule */ id from ����ҽ����¼ where id in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) " & _
              " Union All " & _
              " select /*+ rule */ id from ����ҽ����¼ where ���id in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) "
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKeys)
    Do While Not rs.EOF
        strҽ��ID = strҽ��ID & "," & rs("id")
        rs.MoveNext
    Loop
    strҽ��ID = Mid(strҽ��ID, 2)
    If Chk���۷���(Me, strҽ��ID, 0) = False And Trim(strҽ��ID) <> "" Then
        Exit Sub
    End If
    '----------------------------------------------------------------------------------------------------------------------------------

    Exit Sub
errH:
    Call InitEdit
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function zlRefresh(ByVal Row As ReportRow) As Boolean
'��ʾ�걾������Ϣ
'lngSampleID���걾��¼ID
    Dim rs As New ADODB.Recordset
    Dim mstrSql As String
    Dim strTmp As String
    Dim strAge As String
    Dim aAge() As String

    On Error GoTo ErrHand

    mblnEdit = False
    ClearItem

    Me.cbo��������.ListIndex = -1: Me.cboҽ��.ListIndex = -1
    Me.cbo(0).ListIndex = -1: Me.cbo(1).ListIndex = -1

    On Error Resume Next
        Me.txt���� = Row.Record(mCol.����).Value
        strTmp = "����='" & CStr(Row.Record(mCol.�Ա�).Value) & "'"
        mRsSex.filter = strTmp
        If mRsSex.EOF = False Then
            Me.cbo�Ա�.Text = mRsSex!���� & "-" & mRsSex!����
        End If

'        Me.cbo�Ա�.Text = Row.Record(mCol.�Ա�).Value
'        Me.txt���� = IIf(IsNull(rs("����")), "", Val(rs("����"))): If Me.txt���� = "0" Then Me.txt���� = ""
'        Me.txt���� = IIf(IsNull(rs("����")), "", IIf(IsNumeric(rs("����")), Val(rs("����")), Mid(rs("����"), 1, Len(rs("����")) - 1))): If Me.txt���� = "0" Then Me.txt���� = ""
        Me.txt����.Text = "": Me.txt����1.Text = ""
        strAge = Row.Record(mCol.����).Caption
        
        strAge = Replace(strAge, "Сʱ", "ʱ")
        strAge = Replace(strAge, "����", "��")
        
        If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "��", ""), "��", ""), "��", ""), "ʱ", ""), "��", "")) <> "" Then
            If InStr(strAge, "����") > 0 Or InStr(strAge, "Ӥ��") > 0 Then
                Me.txt����.Text = ""
                Me.cboAge.Text = Trim(strAge)
            Else
                strAge = Replace(Replace(Replace(Replace(Replace(strAge, "��", "��;"), "��", "��;"), "��", "��;"), "ʱ", "ʱ;"), "��", "��;")
                aAge = Split(strAge, ";")
                If UBound(aAge) = 1 Then
                    Me.txt����.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
                Else
                    Me.txt����.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
                    Me.txt����1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "��", "����"), "ʱ", "Сʱ")
                End If
            End If
        Else
            Me.txt����.Text = Val(strAge)
            Me.cboAge.ListIndex = 0
        End If

        Me.txt���� = Row.Record(mCol.��������).Value
        Me.cboAge = Row.Record(mCol.���䵥λ).Value

        'Me.cboAge.Text = IIf(IsNull(rs("����")), "��", Right(rs("����"), 1))
        If cboAge.ListIndex = -1 Then cboAge.ListIndex = 0
        Me.txtPatientDept = Row.Record(mCol.���˿���).Value

        Me.txtID = Row.Record(mCol.��ʶ��).Value
        If Me.txtID.Text = "" Then Me.txtID.Text = Row.Record(mCol.NO).Value
        '������ʱ�Ƿ�����޸ı�ʶ�źͿ���
        If Row.Record(mCol.��������).Value = 1 And Row.Record(mCol.������Դ).Value = 3 And Row.Record(mCol.סԺ��).Value = "" _
                And Row.Record(mCol.�����).Value = "" Then
            Me.txtID.Tag = "���޸�"
        End If

        Me.txtBed = Row.Record(mCol.����).Value

        Me.cbo��������.Text = Row.Record(mCol.�������).Value

        Me.cboҽ��.Text = Row.Record(mCol.������).Value
        Me.txt���� = Row.Record(mCol.����걾).Value

        Me.DTP(1).Value = Row.Record(mCol.�걾ʱ��).Value
        Me.cbo(1).Text = Row.Record(mCol.������).Value
        Me.DTP(0).Value = Row.Record(mCol.����ʱ��).Value

        mstr������ = Row.Record(mCol.������).Value

        'û�в�����ʱ����ʾ����ʱ��
        If Trim(Me.cbo(1).Text) = "" Then
            Me.cbo(1).Visible = False
            Me.DTP(0).Visible = False
            lbl(0).Visible = False
            Me.Picture1(0).Visible = False
        Else
            Me.cbo(1).Visible = True
            Me.DTP(0).Visible = True
            lbl(0).Visible = True
            Me.Picture1(0).Visible = True
        End If

        Me.cbo(0).Text = Row.Record(mCol.�걾��̬).Value

        If Row.Record(mCol.������).Value = "" Then
            Me.cbo(2).Visible = False
            Me.DTP(2).Visible = False
            lbl(1).Visible = False
            Me.Picture1(1).Visible = False
        Else
            Me.cbo(2).Visible = True
            Me.DTP(2).Visible = True
            lbl(1).Visible = True
            Me.Picture1(1).Visible = True
            Me.cbo(2).Text = Row.Record(mCol.������).Value
            Me.DTP(2).Value = Row.Record(mCol.����ʱ��).Value
        End If

        With vsf2
            .Rows = 2
            .RowData(1) = IIf(Row.Record(mCol.����id).Value = "", -1, Row.Record(mCol.����id).Value)
            .TextMatrix(1, 1) = IIf(Row.Record(mCol.������).Value = "", "�ֹ�", Row.Record(mCol.������).Value)
            .TextMatrix(1, 2) = Row.Record(mCol.�걾��).Caption
            .TextMatrix(1, 4) = Val(Row.Record(mCol.ҽ��id).Value)
            .TextMatrix(1, 5) = IIf(Row.Record(mCol.�걾���).Value = 1, -1, 0) '   IIf(rs("�걾���") = 0, 0, -1)
            '---- �����Ƿ����ּ����־����ʾ������
            If mblnEmerge Then
                .Body.ColWidth(5) = 250
            Else
                .Body.ColWidth(5) = 0
            End If
        End With
        Me.txtҽ������ = Row.Record(mCol.������Ŀ).Value
        Me.txtҽ������.Tag = Row.Record(mCol.������Ŀ).Value

        If lbl(1).Visible = False Then
            Me.lbl���δͨ��.Top = lbl(1).Top
        End If

        If lbl(0).Visible = False Then
            Me.lbl���δͨ��.Top = lbl(0).Top
        End If
        Me.lbl���δͨ��.Caption = Trim(Row.Record(mCol.���δͨ��).Value)
        Me.lbl���δͨ��.Visible = (Me.lbl���δͨ��.Caption <> "")
        Me.lbl���δͨ��.Top = Me.lbl(1).Top + Me.lbl(1).Height + 200

        If Row.Record(mCol.�������).Value = "Ժ��" Then
            lblCash.Caption = ""
        Else

            Select Case ShowCharge(Row.Record(mCol.ID).Value)
                Case -1     'δ���շѵ���
                    lblCash.Caption = ""
                Case 0      '���۵�
                    lblCash.Caption = "��"
                Case 1, 2     '������շ�
                    lblCash.Caption = "��"
            End Select
        End If
        Me.lblRegister.Caption = Nvl(Row.Record(mCol.��������).Value, 0)
        mbln΢������Ŀ = IIf(Val(Row.Record(mCol.΢����걾).Value) = 1, True, False)
    SetPatientInfoWrite True
    zlRefresh = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ZlEditStart(ByVal intEditType As Integer, ByVal lngDeptID As Long, ByVal lngDeviceID As Long, _
    Optional ByVal lngSampleID As Long = 0, _
    Optional ByVal lngAdviceID As Long = 0, Optional ByVal intSampleType As Integer = 0, _
    Optional ByVal strAuditer As String = "", Optional ByVal lngDefaultItemID As Long, _
    Optional ByVal lngPatientID As Long) As Boolean
'�༭�걾������Ϣ
'intEditType����0�����ա�1���Ǽǡ�2�����º��ա�3����������
'lngDeptID����ǰ������
'lngDeviceID��Ĭ�ϼ�������
'lngSampleID����ѡ����ǰ�޸ĵı걾ID
'lngAdviceID����ѡ����ǰҪ���յļ�������ҽ��ID
'intSampleType����ѡ���걾���0����ͨ��1������
'lngDefaultItemID �� ��ѡ��Ĭ����ĿID
    mintEditMode = intEditType: ItemDeptID = lngDeptID
    mlngDefaultDevice = lngDeviceID: mlngSampleID = lngSampleID
    mstrKeys = IIf(mintEditMode = 1, "", lngAdviceID): mintSampleType = intSampleType
    mstrAuditer = strAuditer
    mlngDefaultItemID = lngDefaultItemID

    If Val(mstrKeys) = 0 Then mstrKeys = ""

    mblnSaveAdvice = False
    If mintEditMode = 0 Or mintEditMode = 1 Then
        Me.lblRegister.Caption = 0
    End If

    If InitEdit = False Then
        ZlEditStart = False
        Exit Function
    End If

'    SetActiveWindow Me.Hwnd
    Me.txt����.Enabled = True
    Me.txt����.SetFocus

    ZlEditStart = True
    mblnEdit = True

    '���Ӵ������б�ѡ��ʱ����
    If lngPatientID > 0 Then
        Me.txt����.Enabled = False
        Me.txt����.Text = "-" & lngPatientID
'        Call txt����_Validate(False)
        Call Txt����Exec
        Me.txt����.Enabled = True
    End If
    gintSelectFocus = 2
End Function

Public Function ZlSave(Optional ByVal intEditState As Integer) As Long
'���浱ǰ�걾�༭��Ϣ
'intEditMode����ǰ�༭ģʽ��1�����������༭��0�����������༭
    If ValidData = False Then Exit Function
    If SaveData(intEditState) = False Then Exit Function

    On Error Resume Next

    '����ؼ�����
    Call ResetVsf(vsf2)

    Me.txt���� = "": Me.cbo�Ա�.ListIndex = -1: Me.txt���� = "": mstrKeys = "": Me.cboAge.ListIndex = 0: mstrNO = ""
    mblnEdit = False
    txt����.SetFocus
    ZlSave = mlngSampleID

End Function

Public Function ZlRefuse() As Boolean
'�ܾ���ǰ�걾
'intEditMode����ǰ�༭ģʽ��1���ܾ�������༭��0���ܾ�������༭
    ZlRefuse = True
    mblnEdit = False
End Function

Public Function ZlCancel() As Boolean
'ȡ���༭
    Me.Enabled = False

    mintEditMode = -1
    mstrNO = ""
    ClearItem
    mblnEdit = False
    ZlCancel = True
End Function

Private Function InitEdit() As Boolean
    Dim strSQL As String, rs As New ADODB.Recordset, i As Long

    On Error GoTo ErrHand
    mblnBarCode = False
'    mblnSaveAdvice = True

    PatientType = 1: mlng����ID = 0: Me.txt����.ForeColor = vbBlack
    iInputType = -1: mstrNO = ""
    Set rsRelativeAdvice = Nothing

    strSQL = "SELECT ����,0 AS ID FROM ����걾��̬"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(0), rs)

    '��ʼ�༭��Ŀ
    InitDepts
    '���ݱ༭״̬������沼��
    cmdOpen.Visible = Not (mintEditMode = 1)
    Me.cbo(1).Enabled = mblnBarCode
    SetAdviceEnable (mintEditMode = 1 Or Me.lblRegister = 0)

    If mintEditMode > 1 Then
        vsf2.Body.Editable = flexEDKbdMouse  'flexEDKbdMouse   'flexEDNone
        DTP(1).Enabled = IIf(mbln΢������Ŀ = True, True, False)
        If Len(Trim(Me.txt����)) = 0 Then
            DTP(0).Value = Format(zlDatabase.Currentdate, DTP(0).CustomFormat)
            If Format(DTP(0).Value, "yyyy-mm-dd") = Format(DTP(1).Value, "yyyy-mm-dd") Then DTP(1).Value = DTP(0).Value
        End If
    Else
        vsf2.Body.Editable = flexEDKbdMouse
        DTP(1).Value = Format(zlDatabase.Currentdate, DTP(1).CustomFormat)
        DTP(1).Enabled = True
        DTP(0).Value = DTP(1).Value
    End If

    Me.Enabled = True

    '��ʼ�걾��������Ŀ������
    If mintEditMode > 1 Then
        InitSampleInfo mlngSampleID
        If mlng����ID = 0 Then
            ClearItem
        End If
    Else
        ClearItem
    End If

    InitEdit = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub SetAdviceEnable(ByVal blnEnable As Boolean)

    If InStr(1, mstrPrivs, "ֱ������") = 0 Then blnEnable = False: cmdSel.Enabled = False
    Me.txtҽ������.Enabled = blnEnable: Me.cmdSel.Enabled = blnEnable
    If Me.txtID.Tag <> "" Then
        Me.txtBed.Enabled = blnEnable
        Me.txtBed.Locked = Not blnEnable
        Me.txtPatientDept.Enabled = blnEnable
        Me.txtPatientDept.Locked = Not blnEnable
        Me.txtID.Enabled = blnEnable
        Me.txtID.Locked = Not blnEnable
    End If
'    If mbln΢������Ŀ = False Then
'        Me.txt����.Enabled = blnEnable
'        Me.cmdExt.Enabled = blnEnable
'    Else
        Me.txt����.Enabled = True
        Me.cmdExt.Enabled = True
'    End If

'    Me.cbo��������.Enabled = blnEnable: Me.cboҽ��.Enabled = blnEnable
End Sub

Private Sub AutoSave()
'�Զ����浱ǰ�걾�༭��Ϣ�����뷽ʽ��
    If ValidData = False Then Exit Sub
    If SaveData = False Then Exit Sub

    '����ؼ�����
    Call ResetVsf(vsf2)

    Me.txt���� = "": Me.cbo�Ա�.ListIndex = -1: Me.txt���� = "": Me.txt����1 = "": mstrKeys = "":   Me.cboAge.ListIndex = 0
    txt����.SetFocus

    RaiseEvent ZlAutoSave(mlngSampleID)
End Sub

Private Sub ClearItem()
    Me.txt���� = "": Me.txt����.Tag = "": Me.cbo�Ա�.ListIndex = -1: Me.txt���� = "": Me.txt����1 = "":  Me.cboAge.ListIndex = 0
    Me.txtPatientDept = "": Me.txtID = "": Me.txtID.Tag = "": Me.txtBed = "": Me.txtPatientDept.Tag = 0
    Me.txtҽ������ = "": Me.txt���� = "": Me.txt����.ForeColor = vbBlack
    Me.txtҽ������.Tag = "": Me.txt����.Tag = ""
'    Me.lblCash.Font.Strikethrough = True
    Me.lblCash.Caption = ""
    SetPatientInfoWrite True
    If mintEditMode <= 1 Then ResetVsf vsf2
End Sub

Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strOldText As String

    On Error GoTo errH
    strOldText = Me.cbo��������.Text
    Me.cbo��������.Clear

    strSQL = _
        " Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where B.����ID = A.ID " & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
        " And (B.�������� IN('�ٴ�','���','����'))" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    For i = 1 To rsTmp.RecordCount
        cbo��������.AddItem rsTmp!����
        cbo��������.ItemData(cbo��������.NewIndex) = rsTmp!ID
        If strOldText = rsTmp!���� Then
            cbo��������.ListIndex = cbo��������.NewIndex
        End If
        rsTmp.MoveNext
    Next

    On Error Resume Next
'    Me.cbo��������.Text = strOldText
    If cbo��������.ListCount > 0 And Me.cbo��������.ListIndex = -1 Then
        cbo��������.ListIndex = 0
    End If

    InitDepts = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cboAge_Click()
    If Me.cboAge.Text = "����" Or Me.cboAge.Text = "Ӥ��" Then
        Me.txt����.Text = ""
    End If
End Sub

Private Sub cboAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SetFocusNextIndex Me.cboAge.TabIndex ' zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo��������_GotFocus()
    Call zlControl.TxtSelAll(cbo��������)
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cbo��������_Validate(False)
        SetFocusNextIndex Me.cbo��������.TabIndex  ' zlCommFun.PressKey vbKeyTab
        gintSelectFocus = 2
    End If
End Sub

Private Sub cbo��������_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean

    If cbo��������.ListIndex <> -1 Then mlngReqDept = Me.cbo��������.ItemData(Me.cbo��������.ListIndex): Exit Sub '��ѡ��
    If cbo��������.Text = "" Then '������
        Exit Sub
    End If

    strInput = UCase(NeedName(cbo��������.Text))
    'ȫԺ�ٴ�����
    strSQL = _
        " Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where B.����ID = A.ID " & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
        " And (B.�������� IN('�ٴ�','���'))" & _
        " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
        " Order by A.����"

    On Error GoTo errH
    vRect = GetControlRect(cboҽ��.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��������", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo��������.Height, blnCancel, False, True, UCase(strInput) & "%", UCase(strInput) & "%")
    If Not rsTmp Is Nothing Then
        If Not zlControl.CboLocate(cbo��������, rsTmp!����) Then
            cbo��������.Text = ""
        End If
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    If Me.cbo��������.ListIndex > -1 Then mlngReqDept = Me.cbo��������.ItemData(Me.cbo��������.ListIndex)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdExt_Click()
    Dim tmpExtData As String
    Dim lngKey As Long
    Dim vRect As RECT, blnCancel As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSampleType As String

    On Error Resume Next
    If mstrExtData = "" Then
        gstrSql = "select ������ĿID from ����ҽ����¼ where ���id in (" & mstrKeys & ")"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
        If rsTmp.EOF = True Then Exit Sub   'û��ʱ�˳�
        Do While Not rsTmp.EOF
            tmpExtData = tmpExtData & "," & Nvl(rsTmp("������ĿID"))
            rsTmp.MoveNext
        Loop
        lngKey = Val(Mid(tmpExtData, 2))
    Else
        lngKey = Val(Split(mstrExtData, ";")(0))
    End If

    If lngKey = 0 Then
        gstrSql = "Select ���� as ID,���� From ���Ƽ���걾 order by ���� "
    Else
        gstrSql = "   Select Distinct b.���� as ID,B.����  " & _
                "   From ������ĿĿ¼ A,���Ƽ���걾 B,������Ŀ�ο� C,���鱨����Ŀ D" & _
                "   Where A.ID=D.������ĿID(+) And D.������ĿID=C.��ĿID(+)" & _
                        "       And (C.�걾���� Is Null Or C.�걾����=B.����) And A.ID In (" & lngKey & ") order by b.���� "

    End If
    vRect = GetControlRect(Me.txt����.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "�걾����", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, Me.txt����.Height, blnCancel, False, True)

    If Not rsTmp Is Nothing Then
        strSampleType = Nvl(rsTmp!����)
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ��걾����", vbInformation, gstrSysName
        End If
    End If

    If Trim(strSampleType) <> "" Then
        Me.txtҽ������ = Replace(Me.txtҽ������, "(" & Me.txt���� & ")", "(" & strSampleType & ")")
        Me.txt���� = strSampleType
    End If
    Me.txt����.SetFocus
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdOpen_Click()
    Dim intSelect As Integer

    If mstrKeys = "" Then Exit Sub

    intSelect = OpenSelect("", False): gintSelectFocus = 2 'DoEvents
    Select Case intSelect
        Case 1
            'ѡȡ��һ����Ŀ
            SetAdviceEnable False   '������Ǽ�
            '���������ͱ걾��
            If mintEditMode = 0 Then
                Call LoadDefaultData
                Call SelectDefault

                vsf2.Col = 2
                vsf2.ShowCell vsf2.Row, vsf2.Col
                vsf2.SetFocus
            Else
                '��ҽ��ID
                Call SelectDefault
                Me.cbo�Ա�.SetFocus
            End If
    End Select
End Sub

Private Sub cmdSel_Click()
    '������Ŀ
    Dim rsTmp As New ADODB.Recordset
    If mstrExtData = "" And mintEditMode = 3 And Me.txtҽ������.Enabled = True Then
        gstrSql = "select distinct ������ĿID from ������ͨ��� where ����걾id = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngSampleID)
        Do While Not rsTmp.EOF
            mstrExtData = mstrExtData & "," & Nvl(rsTmp("������Ŀid"))
            rsTmp.MoveNext
        Loop
        If mstrExtData <> "" Then
            mstrExtData = Mid(mstrExtData, 2) & ";" & txt����
        End If
    End If

    If AdviceInput Then
'        DoEvents
        mblnSaveAdvice = True
        gintSelectFocus = 2
        '��ʾ��ȱʡ���õ�ֵ
        txtҽ������.Tag = txtҽ������.Text
        txt����.Tag = txt����.Text

        '��������������걾�ţ��������б걾�Ĳ������غ˻������룩��ֻ���¸�ҽ��ID
        If mintEditMode <= 1 Then
            Call LoadDefaultData
            Call SelectDefault

            With vsf2
                If .Rows > 1 Then
                    .Row = 1
                End If
                .Col = 2
                .ShowCell vsf2.Row, vsf2.Col
                .SetFocus
            End With
        Else
            '��ҽ��ID
            Call SelectDefault
        End If

        Me.txtҽ������.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    Else
'        DoEvents
        gintSelectFocus = 2
        '�ָ�ԭֵ
        txtҽ������.Text = txtҽ������.Tag
        txt����.Text = txt����.Tag
        zlControl.TxtSelAll txtҽ������

        txtҽ������.SetFocus
    End If
    gintSelectFocus = 2
End Sub
'
'Private Sub Form_Activate()
'    On Error Resume Next
'    Select Case mintFocusItem
'        Case FocusItem.�걾��
'            vsf2.SetFocus
'        Case FocusItem.��������
'            Me.cbo��������.SetFocus
'        Case FocusItem.����
'            Me.txt����.SetFocus
'        Case FocusItem.ҽ������
'            Me.txtҽ������.SetFocus
'        Case FocusItem.ҽ��
'            Me.cboҽ��.SetFocus
'    End Select
'    mintFocusItem = 0
'End Sub

Private Sub dtp_Change(Index As Integer)
    If Index = 1 Then
        If Abs(DateDiff("d", DTP(Index).Value, zlDatabase.Currentdate)) > 30 Then
            MsgBox "��ѡ��ļ���ʱ��͵�ǰʱ��������30�죬��ע���Ƿ���ȷ��", vbQuestion, Me.Caption
        End If
    End If
End Sub

Public Sub IdKindChange()
    If Me.ActiveControl Is txt���� Then
       IDKind.IDKind = IIf(IDKind.IDKind = IDKinds.C5���￨, 0, IDKind.IDKind + 1)
    End If
End Sub

Private Sub Form_Load()
    Dim blnEmerge As Boolean
    Dim rs As New ADODB.Recordset, i As Long

    '���ò���
    Call SetPara

    mstrPrivs = gstrPrivs

    mintEditMode = -1
    With vsf2
        .Cols = 0
        .NewColumn "", 0, 4
        .NewColumn "��������", 1600, 1, , 0
        .NewColumn "�걾����", 1200, 1, , 1, 15
        .NewColumn "", 0, 1
        .NewColumn "", 0, 1
        .NewColumn "��", IIf(mblnEmerge, 250, 0), 1, , IIf(mblnEmerge, 1, 0), , flexDTBoolean
        .NoDouble = True
        .FixedCols = 0

        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = &HFDD6C6

        .Body.Appearance = flex3DLight
    End With
    '�Ա�
    Set rs = Nothing
    Set rs = GetDictData("�Ա�")
    Set mRsSex = rs
    cbo�Ա�.Clear
    If Not rs Is Nothing Then
        For i = 1 To rs.RecordCount
            cbo�Ա�.AddItem rs!���� & "-" & rs!����
            If rs!ȱʡ = 1 Then
                cbo�Ա�.ItemData(cbo�Ա�.NewIndex) = 1
                cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
            End If
            rs.MoveNext
        Next
    End If

    gstrSql = "Select Distinct D.ID" & vbNewLine & _
            " From ����С���Ա A, ����С�� B, ����С������ C, �������� D" & vbNewLine & _
            " Where A.С��id = B.ID And B.ID = C.С��id��and ��Աid = [1] And C.����id = D.ID And C.���� = 1"

    Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, UserInfo.ID)
    Do Until rs.EOF
        mstrMachines = mstrMachines & ";" & rs("ID")
        rs.MoveNext
    Loop
    If mstrMachines <> "" Then mstrMachines = mstrMachines & ";"
'    mintFocusItem = FocusItem.����
    '-- 2007-10-26 ����һ��֧ͨ��
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hWnd)

    If mobjSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
            MsgBox "IDKind��ʼ��ʧ��!", vbInformation, gstrSysName
        Else
            IDKind.IDKindStr = mobjSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
        End If
    End If

    IDKind.IDKind = 0 'IC��

    If mobjLisInsideComm Is Nothing Then
        Dim strErr As String
        Set mobjLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not mobjLisInsideComm Is Nothing Then
            '��ʼ��LIS�ӿڲ���
            If mobjLisInsideComm.InitComponentsHIS(glngSys, glngModul, gcnOracle, strErr) = False Then
                If strErr <> "" Then
                    MsgBox "��ʼ��LIS�ӿ�ʧ�ܣ�" & vbCrLf & strErr
                End If
                Set mobjLisInsideComm = Nothing
            End If
        End If
    End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Function GetDictData(strDict As String) As ADODB.Recordset
'���ܣ���ָ�����ֵ��ж�ȡ����
'������strDict=�ֵ��Ӧ�ı���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH

    strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strDict & " Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    If Not rsTmp.EOF Then Set GetDictData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If mblnBarCode = True And KeyAscii <> vbKeyReturn Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then
        If mblnBarCode Then
            AutoSave
        Else
            SetFocusNextIndex Me.cbo�Ա�.TabIndex
'            zlCommFun.PressKey vbKeyTab
        End If
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    zlcommfun.OpenIme False
    '2007-10-26 ����һ��֧ͨ��
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    Set mobjSquareCard = Nothing
    mstrMachines = ""
End Sub

Private Sub IDKind_Click()
    Dim lng�����ID As Long, strOutCardNO As String, strExpand As String, strOutPatiInforXML As String
' 2007-10-26 ����һ��֧ͨ��
    Dim blnCancle As Boolean
    IDKind.Tag = ""
    If Not txt����.Locked And txt����.Text = "" And txt����.Tag = "" Then
        If IDKind.IDKind = IDKinds.C3IC���� Then
            If mobjICCard Is Nothing Then
                Set mobjICCard = CreateObject("zlICCard.clsICCard")
                Set mobjICCard.gcnOracle = gcnOracle
            End If
            If Not mobjICCard Is Nothing Then
                txt����.Text = mobjICCard.Read_Card()
                If txt����.Text <> "" Then
                    IDKind.Tag = "ҽ����"

'                    Call txt����_Validate(blnCancle)
                    Call Txt����Exec
                    Me.txt����.SetFocus
                    gintSelectFocus = 2
                End If
            End If
        End If
    End If
    lng�����ID = Val(IDKind.GetKindItem("�����ID"))
    If lng�����ID = 0 Then Exit Sub

    If mobjSquareCard.zlReadCard(Me, glngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txt����.Text = strOutCardNO
    If txt����.Text <> "" Then Call txt����_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer)
    mblnShowPwd = Trim(IDKind.GetKindItem(7)) <> ""
    Me.txt���� = ""
    If mblnShowPwd = True Then
        Me.txt����.PasswordChar = "*"
    Else
        Me.txt����.PasswordChar = ""
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
' 2007-10-26 ����һ��֧ͨ��
    Dim lngPreIDKind As Long
    Dim blnCancle As Boolean
    IDKind.Tag = ""
    If Not txt����.Locked And txt����.Text = "" And txt����.Tag = "" And Me.ActiveControl Is txt���� Then
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKinds.C2���֤��
        txt����.Text = strID
        IDKind.Tag = "���֤"
'        Call txt����_Validate(blnCancle)
        Call Txt����Exec
        IDKind.IDKind = lngPreIDKind
        gintSelectFocus = 2
    End If
End Sub

Private Sub txtBed_GotFocus()
    zlControl.TxtSelAll txtBed
End Sub

Private Sub txtBed_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
'        zlCommFun.PressKey vbKeyTab
        SetFocusNextIndex Me.txtBed.TabIndex
    End If
End Sub

Private Sub txtID_GotFocus()
    zlControl.TxtSelAll txtID
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
'        zlCommFun.PressKey vbKeyTab
        SetFocusNextIndex Me.txtID.TabIndex
    Else
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
    End If
End Sub

Private Sub txtPatientDept_GotFocus()
    zlControl.TxtSelAll txtPatientDept
End Sub

Private Sub txtPatientDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(Me.txtҽ������)) > 0 Then
            vsf2.Col = 2
            vsf2.ShowCell vsf2.Row, vsf2.Col
            vsf2.SetFocus
        Else
'            zlCommFun.PressKey vbKeyTab
            SetFocusNextIndex Me.txtPatientDept.TabIndex
        End If
        Exit Sub
    End If
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SetFocusNextIndex Me.txt����.TabIndex ' zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(Me.txtҽ������)) > 0 And Me.txtID.Enabled = False Then
            With vsf2
                If .Rows > 1 Then
                    .Row = 1
                End If
                .Col = 2
                .ShowCell vsf2.Row, vsf2.Col
                .SetFocus
            End With
        Else
'            zlCommFun.PressKey vbKeyTab
            SetFocusNextIndex Me.txt����.TabIndex
        End If
        Exit Sub
    Else
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789.*")
    End If
End Sub

Private Sub cbo��������_Click()
    If cbo��������.ListIndex > -1 Then InitDoctors cbo��������.ItemData(cbo��������.ListIndex)
End Sub

Private Sub cboҽ��_GotFocus()
    Call zlControl.TxtSelAll(cboҽ��)
End Sub

Private Sub cboҽ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
'        If mbln΢������Ŀ = True Then
'            Call cboҽ��_Validate(False)
'            dtp(0).SetFocus
'        Else
'            zlCommFun.PressKey vbKeyTab
            Call cboҽ��_Validate(False)
            SetFocusNextIndex Me.cboҽ��.TabIndex + 2
            gintSelectFocus = 2
'        End If
    End If
End Sub

Private Sub cboҽ��_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim lngDept As Long

    If cboҽ��.ListIndex <> -1 Then mstrReqDoctor = Me.cboҽ��.Text: Exit Sub '��ѡ��
    If cboҽ��.Text = "" Then '������
        Exit Sub
    End If

    lngDept = cbo��������.ItemData(cbo��������.ListIndex)

    strInput = UCase(NeedName(cboҽ��.Text))
    'ȫԺҽ��
    strSQL = "Select Distinct ����ID From ��������˵�� Where ������� IN(1,2,3)"
    strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
        " And B.����ID IN(" & strSQL & ")" & IIf(lngDept > 0, " and b.����ID=[3] ", "") & _
        " And (Upper(A.���) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
        " And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
        " Order by A.����"


    On Error GoTo errH
    vRect = GetControlRect(cboҽ��.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҽ��", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cboҽ��.Height, blnCancel, False, True, strInput & "%", strInput & "%", lngDept)
    If Not rsTmp Is Nothing Then
        cboҽ��.Text = rsTmp!����
'        Me.dtp(0).SetFocus
'        SetFocusNextIndex Me.cboҽ��.TabIndex


    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ��ҽ����", vbInformation, gstrSysName
        End If
        Cancel = True: gintSelectFocus = 2: Exit Sub
    End If
    If Len(Trim(Me.cboҽ��.Text)) > 0 Then mstrReqDoctor = Me.cboҽ��.Text
    gintSelectFocus = 2
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngloop As Long

    If KeyAscii = vbKeyReturn Then
        If mblnBarCode And Index = 1 Then AutoSave: Exit Sub

        For lngloop = 0 To cbo(Index).ListCount - 1
            If InStr(cbo(Index).List(lngloop), "-") > 0 Then
                If Mid(cbo(Index).List(lngloop), 1, InStr(cbo(Index).List(lngloop), "-") - 1) = cbo(Index).Text Then
                    cbo(Index).Text = cbo(Index).List(lngloop)
                    Exit For
                End If
            End If
        Next
'        zlCommFun.PressKey vbKeyTab
        SetFocusNextIndex Me.cbo(Index).TabIndex
    End If
End Sub

Private Sub cbo_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            Cancel = Not StrIsValid(cbo(Index).Text, 50)
        Case 1, 2
            Cancel = Not StrIsValid(cbo(Index).Text, 50)
    End Select
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
'        zlCommFun.PressKey vbKeyTab
        SetFocusNextIndex Me.DTP(Index).TabIndex
    End If
End Sub

Private Sub InitDoctors(ByVal lng����ID As Long)
'���ܣ���ȡ��ǰ���������а�����������Ա
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strOldDoctor As String

    strOldDoctor = Me.cboҽ��.Text
    Me.cboҽ��.Clear

    '����ҽ����ʿ
    strSQL = _
        "Select Distinct A.ID,B.����ID,A.���,A.����,Upper(A.����) as ����," & _
        " C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
        " And C.��Ա���� IN('ҽ��') And B.����ID=[1] " & _
        " And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) "

    strSQL = strSQL & " Order by ����,��Ա���� Desc"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)

    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboҽ��.AddItem rsTmp!����
            cboҽ��.ItemData(cboҽ��.ListCount - 1) = rsTmp!����ID
            If rsTmp!���� = strOldDoctor Then
                cboҽ��.ListIndex = cboҽ��.NewIndex
            End If

            If rsTmp!ID = UserInfo.ID And cboҽ��.ListIndex = -1 Then cboҽ��.ListIndex = cboҽ��.NewIndex
            rsTmp.MoveNext
        Next

        If cboҽ��.ListCount = 1 And cboҽ��.ListIndex = -1 Then cboҽ��.ListIndex = 0
    End If
End Sub

Private Sub txt����_GotFocus()
    txt����.SelStart = 0
    txt����.SelLength = Len(txt����.Text)
    If IDKind.IDKind = IDKinds.C0���� Then
        If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    ElseIf IDKind.IDKind = IDKinds.C3IC���� Then
        If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    End If
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) <> 0 Then
        zlCommFun.OpenIme True
    End If
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim blnCard As Boolean
'    '���￨
'    blnCard = False
'    If IDKind.IDKind = IDKinds.C5���￨ Then
'        If KeyCode <> 8 And Len(txt����.Text) = gbytCardNOLen - 1 And txt����.SelLength <> Len(txt����.Text) Then
'            txt����.Text = txt����.Text & UCase(Chr(KeyCode))
'            blnCard = True
'            KeyCode = 0
'        End If
'    End If
'
'
''    SetActiveWindow Me.Hwnd
'    If KeyCode <> vbKeyReturn And blnCard = False Then
'        KeyCode = Asc(UCase(Chr(KeyCode)))
'
'    Else
'        KeyCode = 0
''        zlCommFun.PressKey vbKeyTab
'
'        Me.txt����.Enabled = False
''        Call txt����_Validate(False)
'        Call Txt����Exec
'        Debug.Print Me.txt����
'        If Me.txt����.Text <> "" Then
'            SetFocusNextIndex txt����.TabIndex
'        Else
'            Me.txt����.Enabled = True
'            If mstr�������� <> "" And frmLabMain.mlngMachineID > 0 Then Me.txt����.SetFocus
''            Me.txt����.SetFocus
'            If IDKind.IDKind = IDKinds.C5���￨ Then
'                Me.txt����.Text = ""
'            End If
'        End If
'        gintSelectFocus = 2
'        Me.txt����.Enabled = True
'
'        mblnCard = False
'    End If
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean

    If CheckIsInclude(UCase(Chr(KeyAscii)), "'����;��:��?��|,����""") = True Then KeyAscii = 0

    blnCard = False
    If IDKind.IDKind = IDKinds.C5���￨ Then
        gbytCardNOLen = Val(IDKind.GetKindItem("���ų���", IDKinds.C5���￨))
        If KeyAscii <> 8 And Len(txt����.Text) = gbytCardNOLen - 1 And txt����.SelLength <> Len(txt����.Text) Then
            If KeyAscii <> 13 Then
                txt����.Text = txt����.Text & UCase(Chr(KeyAscii))
            End If
            blnCard = True
            KeyAscii = 0
        End If
    End If

    If IDKind.IDKind = IDKinds.C5���￨ Then
'        mblnCard = zlCommFun.InputIsCard(txt����, KeyAscii, True)
    End If

    If KeyAscii = vbKeyReturn Or blnCard = True Then

        KeyAscii = 0
'        zlCommFun.PressKey vbKeyTab

        Me.txt����.Enabled = False
'        Call txt����_Validate(False)
        Call Txt����Exec

        If Me.txt����.Text <> "" Then
            SetFocusNextIndex txt����.TabIndex
        Else
            Me.txt����.Enabled = True
            If mstr�������� <> "" And frmLabMain.mlngMachineID > 0 Then Me.txt����.SetFocus
'            Me.txt����.SetFocus
            If IDKind.IDKind = IDKinds.C5���￨ Then
                Me.txt����.Text = ""
            End If
        End If
'        Debug.Print Me.txt����
        gintSelectFocus = 2
        Me.txt����.Enabled = True

        mblnCard = False
    End If

End Sub

Private Sub txt����_LostFocus()
    txt����.SelStart = 0
    txt����.SelLength = Len(txt����.Text)
    If txt����.Text = "" And Not txt����.Locked Then
        If IDKind.IDKind = IDKinds.C0���� Then
            If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (True)
        ElseIf IDKind.IDKind = IDKinds.C3IC���� Then
            If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (True)
        End If
    End If

End Sub

Private Sub txt����_Validate(Cancel As Boolean)
'    Dim strInput As String
'    Dim rsTmp As New ADODB.Recordset, i As Integer
'    Dim strField As String
'    Dim strBarCode As String
'    Dim rsDept As ADODB.Recordset, strsql As String
'    Dim intSelect As Integer
'    Dim intPatientSource As Integer                     '������Դ
'    Dim rsPaInfo As New ADODB.Recordset                 '��ȡ�걾��¼�е�"��ʶ��"
'    Dim blnGetPaInfo As Boolean
'
'    On Error GoTo errH
'    If Len(Trim(txt����)) = 0 Or Me.txt����.Enabled = True Or Me.txt���� = Me.txt����.Tag Then Exit Sub
'
'    If txt���� <> txt����.Tag Then mlng����id = 0
'    If mlng����id > 0 Then
'        strsql = "select ������Դ from ����걾��¼ where id = [1] "
'        Set rsTmp = zlDatabase.OpenSQLRecord(strsql, gstrSysName, mlngSampleID)
'        If rsTmp.RecordCount > 0 Then
'            intPatientSource = Nvl(rsTmp("������Դ"), PatientType)
'        Else
'            intPatientSource = PatientType
'        End If
'
'        If txt���� = txt����.Tag And intPatientSource <> 3 Then Exit Sub
'    Else
'        If txt���� = txt����.Tag Then Exit Sub
'    End If
'
'    mblnSaveAdvice = True
'    Cancel = Not StrIsValid(txt����.Text, txt����.MaxLength)
'
'    '��ʼ������Ϣ
'    '2007-10-26 ����һ��֧ͨ��
'    If IDKind.Tag = "ҽ����" Then
'        Set rsTmp = GetPatient("��" & txt����)
'        IDKind.Tag = ""
'    ElseIf IDKind.Tag = "���֤" Then
'        Set rsTmp = GetPatient("֤" & txt����)
'        IDKind.Tag = ""
'    Else
'        Set rsTmp = GetPatient(txt����)
'    End If
'
'    If rsTmp.EOF = True Then
'        Set rsTmp = GetPatientInfo(txt����)
'        blnGetPaInfo = True
'    End If
'
'    strBarCode = txt����
'    If rsTmp.RecordCount <= 0 Then
'        mlng����id = 0
'        '�Ǽ��²���
'        mstrKeys = ""
'        Me.txt���� = "": Me.cboAge.ListIndex = 0
'        Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
'        Me.txtID = "": Me.txtBed = ""
'        '���������Ժ�ڲ��ˣ����������
'        If InStr("+-*./", Left(Me.txt����.Text, 1)) > 0 Or mblnBarCode Then
'            Me.txt����.Text = "": Cancel = True
'            Exit Sub
'        End If
'        If mblnBarCode = True Then
'            MsgBox "û���ҵ����룬��鿴�����Ƿ���ȡ���󶨣�", vbInformation, Me.Caption
'            Me.txt����.Text = "": Cancel = True
'            Exit Sub
'        End If
'        PatientType = 1
'        '����Ǽǵ�Ĭ�Ͽ��ҡ�ҽ��
'        If mlngReqDept > 0 Then
'            cbo��������.ListIndex = FindComboItem(cbo��������, mlngReqDept)
'            Me.cboҽ��.Text = mstrReqDoctor
'        End If
'        SetPatientInfoWrite False
'    Else
'        On Error Resume Next
'        Me.txt����.Text = Nvl(rsTmp("����"))
''        Me.txt���� = IIf(IsNull(rsTmp("����")), "", Val(rsTmp("����"))): If Me.txt���� = "0" Then Me.txt���� = ""
''        Me.cboAge.Text = IIf(IsNull(rsTmp("����")), "��", Replace(rsTmp("����"), Val(rsTmp("����")), ""))
'        Me.txt���� = IIf(IsNull(rsTmp("����")), "", IIf(IsNumeric(rsTmp("����")), Val(rsTmp("����")), Mid(rsTmp("����"), 1, Len(rsTmp("����")) - 1))): If Me.txt���� = "0" Then Me.txt���� = ""
'        Me.cboAge.Text = IIf(IsNull(rsTmp("����")), "��", Right(rsTmp("����"), 1))
'        If cboAge.ListIndex = -1 Then cboAge.ListIndex = 0
'        Me.cbo�Ա� = Nvl(rsTmp("�Ա�")) ' CombIndex(cbo�Ա�, Nvl(rsTmp("�Ա�")))
'        mlng����id = Nvl(rsTmp("����ID"), 0): PatientType = Nvl(rsTmp("PatientType"), 1)
'
'        '����Ĭ�Ͽ������ҡ�ҽ��
'        cbo��������.ListIndex = FindComboItem(cbo��������, Nvl(rsTmp("���˿���"), 0))
''        DoEvents
'        gintSelectFocus = 2
'        strField = ""
'        strField = rsTmp.Fields("ҽ��").Name
'        If strField = "ҽ��" Then
'            Me.cboҽ��.Text = Nvl(rsTmp("ҽ��"))
'            For i = 0 To Me.cboҽ��.ListCount - 1
'                If Me.cboҽ��.List(i) Like Nvl(rsTmp("ҽ��")) Then
'                    Me.cboҽ��.ListIndex = i
'                    Exit For
'                End If
'            Next
'        End If
'        '��ʾ���˿���
'        If IsNumeric(rsTmp("���˿���")) = True Then
'            strsql = "Select ���� From ���ű� Where ID=[1]"
'            Set rsDept = zlDatabase.OpenSQLRecord(strsql, Me.Caption, CLng(Nvl(rsTmp("���˿���"), 0)))
'            If rsDept.EOF Then
'                Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
'            Else
'                Me.txtPatientDept.Text = rsDept("����"): Me.txtPatientDept.Tag = Nvl(rsTmp("���˿���"), 0)
'            End If
'        Else
'            Me.txtPatientDept = Nvl(rsTmp("���˿���"))
'        End If
'
'        Me.txtID = Nvl(rsTmp("סԺ��")): If Len(Me.txtID) = 0 Then Me.txtID = Nvl(rsTmp("�����"))
'
'        Me.txtBed = Nvl(rsTmp("��ǰ����"))
'
'        '������걾��¼����ʱ������ʾ����걾��¼�����Ϣ
'        If Trim(Me.txtID) = "" Or Trim(Me.txtBed) = "" Or Trim(Me.txtPatientDept) = "" Then
'            If blnGetPaInfo = True Then
'                Me.txtID = Nvl(rsTmp("��ʶ��"), Me.txtID)
'                Me.txtBed = Nvl(rsTmp("����"), Me.txtBed)
'                Me.txtPatientDept = Nvl(rsTmp("���˿���"), Me.txtPatientDept)
'            End If
'        End If
'
'        '����Ǽǵ�Ĭ�Ͽ��ҡ�ҽ��
'        If Me.cbo��������.ListIndex = -1 And mlngReqDept > 0 Then
'            cbo��������.ListIndex = FindComboItem(cbo��������, mlngReqDept)
'            Me.cboҽ��.Text = mstrReqDoctor
'        End If
'    End If
'    '����ʱѡ���������
'    If mlng����id > 0 And Not mintEditMode = 1 And (intPatientSource <> 3 Or mblnBarCode = True) Then
'        intSelect = OpenSelect(strBarCode, True)
''        DoEvents
'        gintSelectFocus = 2
'        Select Case intSelect
'            Case 0
'                'û��ƥ�����Ŀ
'                mstrKeys = ""
'                If mlng����id = 0 Or mblnBarCode Then
''                    mintFocusItem = FocusItem.����
'                    MsgBox "�����ѱ����գ�", vbInformation, Me.Caption
'                    mlng����id = 0
'                    txt����.Text = ""
'                    If Me.txt����.Enabled = True Then
'                        txt����.SetFocus
'                    End If
'                    Cancel = True
'                Else
'                    '����Ǽ�
'                    SetAdviceEnable True
'                    If mlngDefaultItemID > 0 Then
'                        strsql = "select �걾��λ from ������ĿĿ¼ where id = [1] "
'                        Set rsTmp = zlDatabase.OpenSQLRecord(strsql, gstrSysName, mlngDefaultItemID)
'                        AdviceSet������� 3, mlngDefaultItemID & ";" & rsTmp("�걾��λ")
'                        mstrExtData = mlngDefaultItemID & ";" & rsTmp("�걾��λ")
'                        '��ȡ�ɼ���ʽ
'                        Set rsTmp = SelectCap(Split(Split(mstrExtData, ";")(0), ",")(0))
'                        If rsTmp Is Nothing Then
'                            MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName
'                            Exit Sub
'                        End If
'                        mlngCapID = rsTmp("ID")
'                    End If
'                    If rsRelativeAdvice Is Nothing Then
'                        Me.txtҽ������ = ""
'                        Me.txt���� = ""
'                        Me.txtҽ������.Tag = "": Me.txt����.Tag = "": mlngCapID = 0
'                    Else
'                        txtҽ������.Text = Get�����������(2, "")
'                        txtҽ������.Text = txtҽ������.Text & "(" & Split(mstrExtData, ";")(1) & ")"
'                        Me.txt���� = Split(mstrExtData, ";")(1)
'                        If mintEditMode = 0 Then
'                            Call LoadDefaultData
'                            Call SelectDefault
''                            mintFocusItem = FocusItem.�걾��
'
''                            DoEvents
'                            vsf2.Col = 2
'                            vsf2.ShowCell vsf2.Row, vsf2.Col
'                            vsf2.SetFocus
'                            gintSelectFocus = 2
'                        Else
'                            '��ҽ��ID
'                            Call SelectDefault
'                        End If
'                    End If
'                End If
'            Case 1
'                'ѡȡ��һ����Ŀ
'                SetAdviceEnable False   '������Ǽ�
'                '���������ͱ걾��
'                If mintEditMode = 0 Then
'                    Call LoadDefaultData
'                    Call SelectDefault
''                    mintFocusItem = FocusItem.�걾��
'
'                    vsf2.Col = 2
'                    vsf2.ShowCell vsf2.Row, vsf2.Col
'                    vsf2.SetFocus
'                Else
'                    '��ҽ��ID
'                    Call SelectDefault
'                    vsf2.Col = 2
'                    vsf2.ShowCell vsf2.Row, vsf2.Col
'                    vsf2.SetFocus
'                End If
'                '�����Զ�����
'                If mblnBarCode Then Me.cbo�Ա�.SetFocus
'            Case 2
'                'ȡ���˱���ѡ��
''                mintFocusItem = FocusItem.����
'
'                mlng����id = 0
'                mstrKeys = ""
'                txt����.Text = ""
'                If Me.txt����.Enabled = True Then
'                    txt����.SetFocus
'                End If
'                Cancel = True
'        End Select
'    Else
'        SetAdviceEnable True
'        If mlngDefaultItemID > 0 Then
'            strsql = "select �걾��λ from ������ĿĿ¼ where id = [1] "
'            Set rsTmp = zlDatabase.OpenSQLRecord(strsql, gstrSysName, mlngDefaultItemID)
'            AdviceSet������� 3, mlngDefaultItemID & ";" & rsTmp("�걾��λ")
'            mstrExtData = mlngDefaultItemID & ";" & rsTmp("�걾��λ")
'            '��ȡ�ɼ���ʽ
'            Set rsTmp = SelectCap(Split(Split(mstrExtData, ";")(0), ",")(0))
'            If rsTmp Is Nothing Then
'                MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName
'                Exit Sub
'            End If
'            mlngCapID = rsTmp("ID")
'        End If
'        If rsRelativeAdvice Is Nothing Then
'            Me.txtҽ������ = ""
'            Me.txt���� = ""
'            Me.txtҽ������.Tag = "": Me.txt����.Tag = "": mlngCapID = 0
'        Else
'            txtҽ������.Text = Get�����������(2, "")
'            txtҽ������.Text = txtҽ������.Text & "(" & Split(mstrExtData, ";")(1) & ")"
'            Me.txt���� = Split(mstrExtData, ";")(1)
'            If mintEditMode <= 1 Then
'                Call LoadDefaultData
'                Call SelectDefault
'
'                vsf2.Col = 2
'                vsf2.ShowCell vsf2.Row, vsf2.Col
'                vsf2.SetFocus
'            Else
'                '��ҽ��ID
'                Call SelectDefault
'            End If
'        End If
'    End If
'
'    If mlng����id > 0 Then
'        txt����.Tag = txt����.Text
'    End If
'    Exit Sub
'errH:
'    Call InitEdit
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'    Call SaveErrLog
End Sub
Private Function GetPatient(strCode As String) As ADODB.Recordset
'���ܣ���ȡ������Ϣ������ʾ�ò��˴��ڵ�ҽ��ʱ��
    Dim strSQL As String, i As Long
    Dim strNO As String, str���� As String, lng����ID As Long
    Dim strSeek As String
    Dim objPoint As POINTAPI, lng�Һ�Ч�� As Long
    Dim strTmp As String, rsTmp As New Recordset, str�Һŵ� As String
    Dim strС�� As String, str���� As String
    Dim strSQLbak As String
    Dim lng�����ID As Long


    On Error GoTo errH

    strС�� = frmLabMain.mstrMachineGroup
    If strС�� <> "����С��" Then
        strС�� = Mid(strС��, 1, InStr(strС��, "-") - 1)
    End If
    str���� = frmLabMain.mlngMachineID
    mstr�������� = ""

'    strsql = "Select ����ֵ From ϵͳ������ Where ������=21"
'    Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption)
'    If rsTmp.RecordCount > 0 Then
'        lng�Һ�Ч�� = Val(0 + rsTmp.Fields("����ֵ"))
'    End If
    lng�Һ�Ч�� = Val(zlDatabase.GetPara(21, glngSys))

    If lng�Һ�Ч�� = 0 Then lng�Һ�Ч�� = 7 'δ����Ϊ���2��

    If IsNumeric(strCode) And Len(strCode) >= 12 And InStr("*-+./", Mid(strCode, 1, 1)) = 0 Then
        'Ԥ�����뵥������
        mblnBarCode = True
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,B.��ҳID,B.���˿���id As ���˿���,B.����ҽ�� As ҽ��," & _
            "a.����,decode(d.����,null,a.�Ա�,d.���� || '-' || a.�Ա�) as �Ա�,a.����,a.����id,a.סԺ��,a.�����,a.��ǰ����,Zl_Age_Calc(A.����ID) as ����1  " & _
            " From ������Ϣ A,����ҽ����¼ B,����ҽ������ C , �Ա� d Where A.����ID=B.����ID+0 And B.ID=C.ҽ��ID+0 and a.�Ա� = d.����(+) " & _
            " And C.��������=[1] order by b.����ʱ�� desc  "
        Set GetPatient = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCode)
'        If GetPatient.EOF = True Then
'            MsgBox "û���ҵ������Ϊ<" & strCode & ">������!", vbInformation, Me.Caption
'        End If
'        Exit Function
        If GetPatient.EOF = False Then
            Exit Function
        End If
    End If
    mblnBarCode = False
    mblnPrice = False

    If strС�� = "����С��" Then
        strSQL = "Select distinct ����ID,�������� From ����С�� A, ����С������ B, ����С���Ա C Where A.ID = B.С��id And A.ID = C.С��id  and c.��Աid = [1] and �������� =1" & _
        IIf(str���� = 0, "", " and  b.����id = [3] ")
    Else
        strSQL = "Select distinct ����ID,�������� From ����С�� A, ����С������ B, ����С���Ա C Where A.ID = B.С��id And A.ID = C.С��id  and A.���� = [2] and �������� = 1 " & _
        IIf(str���� = 0, "", " and  b.����id = [3] ")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID, strС��, Val(str����))
    Set GetPatient = rsTmp
    If str���� > 0 And rsTmp.RecordCount = 1 Then
        mstr�������� = rsTmp("����ID")
        Me.txt����.Text = "": Me.txt����.Tag = ""
        MsgBox "�����ʹ���������룡�����������ڹ���Ա��ϵ��", vbInformation, Me.Caption
        Exit Function
    End If

    '�����¼ѡ������������ʱ��¼ֻ�ܰ��������������
    Do While Not rsTmp.EOF
        mstr�������� = mstr�������� & "," & rsTmp("����ID")
        GetPatient.MoveNext
    Loop

    strSeek = strCode
    '�жϵ�ǰ����ģʽ
    If IsNumeric(strCode) And IsNumeric(Left(strCode, 1)) And Val(IDKind.GetKindItem("�����ID")) = 0 Then    'ˢ��
        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("ȫ��"), strCode, False, lng����ID) = False Then lng����ID = 0
        If lng����ID = 0 Then
            iInputType = 0
            strSeek = strCode
        Else
            iInputType = 1
            strSeek = lng����ID
        End If
    ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '����ID
        iInputType = 1
        strSeek = Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then 'סԺ��
        iInputType = 2
        strSeek = Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '�����
        iInputType = 3
        strSeek = Mid(strCode, 2)
    ElseIf Left(strCode, 1) = "G" Or Left(strCode, 1) = "." Then '�Һŵ�
        iInputType = 4
        strSeek = Mid(strCode, 2)
    ElseIf Left(strCode, 1) = "/" Then '�շѵ��ݺ�
        iInputType = 5
        strSeek = Mid(strCode, 2)
        mblnPrice = True
    ElseIf mblnCard Or IDKind.IDKind = IDKinds.C5���￨ Then
         iInputType = 7
        strSeek = UCase(strCode)
    ElseIf Not IsNumeric(Mid(strCode, 2)) And Val(IDKind.GetKindItem("�����ID")) = 0 Then '��������
        iInputType = 6
        strSeek = Replace(strCode, "(Ӥ��)", "")
    ElseIf strCode Like "��*" Then  'ҽ����
        strCode = Replace(UCase(strCode), "��", "")
        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("ȫ��"), strCode, False, lng����ID) = False Then lng����ID = 0
        If lng����ID = 0 Then
            iInputType = 8
            strSeek = Replace(UCase(strCode), "��", "")
        Else
            iInputType = 1
            strSeek = lng����ID
        End If

    ElseIf strCode Like "֤*" Then  '���֤
        strCode = Replace(UCase(strCode), "֤", "")
        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("ȫ��"), strCode, False, lng����ID) = False Then lng����ID = 0
        If lng����ID = 0 Then
            iInputType = 9
            strSeek = Replace(UCase(strCode), "֤", "")
        Else
            iInputType = 1
            strSeek = lng����ID
        End If

    Else
        If Val(IDKind.GetKindItem("�����ID")) <> 0 Then
            lng�����ID = Val(IDKind.GetKindItem("�����ID"))
            If mobjSquareCard.zlGetPatiID(lng�����ID, strCode, False, lng����ID) = False Then lng����ID = 0
            If lng����ID = 0 Then lng����ID = 0
        Else
            If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("ȫ��"), strCode, False, lng����ID) = False Then lng����ID = 0
        End If
        iInputType = 1
        strSeek = lng����ID
    End If
    mblnCard = False
    If iInputType = 0 Then 'ˢ��
        strSQL = _
            "Select Distinct Decode(A.��ǰ����id, Null, 1, 2) As Patienttype, A.��ҳid," & vbNewLine & _
            "                Decode(A.��ǰ����id, Null, Nvl(B.ִ�в���id, 0), A.��ǰ����id) As ���˿���, B.ִ���� As ҽ��, A.����," & vbNewLine & _
            "                Decode(C.����, Null, A.�Ա�, C.���� || '-' || A.�Ա�) As �Ա�, A.����, A.����id, A.סԺ��, A.�����, A.��ǰ����,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
            "From (Select " & gConst_������Ϣ_���� & " From ������Ϣ a Where ���￨�� = [1]" & vbNewLine & _
            "       Union " & vbNewLine & _
            "       Select " & gConst_������Ϣ_���� & " From ������Ϣ a Where ����� = [2]" & vbNewLine & _
            "       Union " & vbNewLine & _
            "       Select " & gConst_������Ϣ_���� & " From ������Ϣ a Where סԺ�� = [2]) A, ���˹Һż�¼ B, �Ա� C" & vbNewLine & _
            "Where A.����id = B.����id(+) and (b.����id is null or (b.��¼״̬ =1 and b.��¼���� =1)) And A.����� = B.�����(+) And 0 + B.�Ǽ�ʱ��(+) > Sysdate - [5] And A.�Ա� = C.����(+)"

'        strsql = "Select Distinct Decode(a.��ǰ����id, Null, 1, 2) As Patienttype, Nvl(a.סԺ����, 0) As ��ҳid," & vbNewLine & _
'                "                               Decode(a.��ǰ����id, Null, Nvl(b.ִ�в���id, 0), a.��ǰ����id) As ���˿���, b.ִ���� As ҽ��, a.����," & vbNewLine & _
'                "                               Decode(c.����, Null, a.�Ա�, c.���� || '-' || a.�Ա�) As �Ա�, a.����, a.����id, a.סԺ��, a.�����," & vbNewLine & _
'                "                               a.��ǰ����" & vbNewLine & _
'                "From ������Ϣ a, (Select ִ�в���id, ִ����, ����id, ����� From ���˹Һż�¼ Where �Ǽ�ʱ�� > Sysdate - [5]) b, �Ա� c" & vbNewLine & _
'                "Where (a.���￨�� = [1] Or a.����� =[2] Or a.סԺ��=[2]) And a.����id = b.����id(+) And a.����� = b.�����(+) And a.�Ա� = c.����(+)"


'            " And (A.��ǰ����id IS NOT NULL Or NVL(B.ִ��״̬,1) IN (0,2))"
    ElseIf iInputType = 1 Then '����ID
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,A.��ҳID,Nvl(A.��ǰ����id,0) As ���˿���," & _
            "a.����,decode(b.����,null,a.�Ա�,b.���� || '-' || a.�Ա�) as �Ա�,a.����,a.����id,a.סԺ��,a.�����,a.��ǰ����,'' as ҽ��,Zl_Age_Calc(A.����ID) as ����1  " & _
            " From ������Ϣ A , �Ա� B Where A.����ID=[2] And A.�Ա� = B.����(+) "
    ElseIf iInputType = 2 Then 'סԺ��
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,A.��ҳID,Decode(A.��ǰ����id,Null,Nvl(B.��Ժ����ID,0),A.��ǰ����id) As ���˿���,B.סԺҽʦ As ҽ��," & _
            "a.����,decode(c.����,null,a.�Ա�,c.���� || '-' || a.�Ա�) as �Ա�,a.����,a.����id,a.סԺ��,a.�����,a.��ǰ����,Zl_Age_Calc(A.����ID) as ����1  " & _
            " From ������Ϣ A,������ҳ B,�Ա� C Where A.סԺ��=[2] And A.��ҳID=B.��ҳID And A.����ID=B.����ID And a.�Ա� = C.����(+) " ' And A.��ǰ����id IS NOT NULL And B.��Ժ���� Is NULL"
    ElseIf iInputType = 3 Then '�����
        strSQL = "Select Distinct Decode(A.��ǰ����id,Null,1,2) As PatientType,A.��ҳID,Decode(A.��ǰ����id,Null,Nvl(B.ִ�в���ID,0),A.��ǰ����id) As ���˿���,B.ִ���� As ҽ��," & _
            "a.����,decode(c.����,null,a.�Ա�,c.���� || '-' || a.�Ա�) as �Ա�,a.����,a.����id,a.סԺ��,a.�����,a.��ǰ����,Zl_Age_Calc(A.����ID) as ����1  " & _
            " From ������Ϣ A,(Select NO,ִ�в���ID,ִ����,����ID,�����,��¼����,��¼״̬ From ���˹Һż�¼ Where �Ǽ�ʱ��>sysdate-[5]) B,�Ա� C Where A.�����=[2] And A.����ID=B.����ID(+) and (b.����ID is null or(b.��¼״̬ =1 and b.��¼���� =1)) And A.�����=B.�����(+) And a.�Ա� = C.����(+) "
'            " And (A.��ǰ����id IS NOT NULL Or NVL(B.ִ��״̬,1) IN (0,2))"
    ElseIf iInputType = 4 Then '�Һŵ�
        strNO = GetFullNO(strSeek, 12)
'        strsql = "Select Decode(B.��ҳid, Null, 1, 2) As Patienttype, Nvl(B.��ҳid, 0) As ��ҳid, Nvl(B.ִ�в���id, 0) As ���˿���," & vbNewLine & _
                "       B.ִ���� As ҽ��, A.����, Decode(C.����, Null, A.�Ա�, C.���� || '-' || A.�Ա�) As �Ա�, A.����, A.����id," & vbNewLine & _
                "       A.סԺ��, A.�����, A.��ǰ����,Zl_Age_Calc(A.����ID) as ����1  " & vbNewLine & _
                "From ������Ϣ A, סԺ���ü�¼ B, �Ա� C" & vbNewLine & _
                "Where B.��¼���� = 4 And B.����id = A.����id And A.�Ա� = C.����(+) And B.��¼״̬ In (1, 3) And B.��� = 1 And B.NO = [3] "
        strSQL = "Select 1 As Patienttype, 0 As ��ҳid, Nvl(B.ִ�в���id, 0) As ���˿���," & vbNewLine & _
                "       B.ִ���� As ҽ��, A.����, Decode(C.����, Null, A.�Ա�, C.���� || '-' || A.�Ա�) As �Ա�, A.����, A.����id," & vbNewLine & _
                "       A.סԺ��, A.�����, A.��ǰ����,Zl_Age_Calc(A.����ID) as ����1  " & vbNewLine & _
                "From ������Ϣ A, ������ü�¼ B, �Ա� C" & vbNewLine & _
                "Where B.��¼���� = 4 And B.����id = A.����id And A.�Ա� = C.����(+) And B.��¼״̬ In (1, 3) And B.��� = 1 And B.NO = [3] "
'        strSQLbak = strsql
'        strSQLbak = Replace$(strSQLbak, "סԺ���ü�¼", "������ü�¼")
'        strSQLbak = Replace$(strSQLbak, "Decode(B.��ҳid, Null, 1, 2) As Patienttype", "1 As Patienttype")
'        strSQLbak = Replace$(strSQLbak, "Nvl(B.��ҳid, 0)", "0")
'        strsql = strsql & " union all " & strSQLbak

    ElseIf iInputType = 5 Then '�շѵ��ݺ�
        strNO = GetFullNO(strSeek, 13): mstrNO = strNO

        strSQL = "Select 1 As Patienttype, 0 As ��ҳid," & vbNewLine & _
                "       Nvl(A.��ǰ����id, B.��������id) As ���˿���, B.������ As ҽ��, B.����," & vbNewLine & _
                "       Decode(C.����, Null, B.�Ա�, C.���� || '-' || B.�Ա�) As �Ա�, B.����, A.����id, A.��λ�绰, A.������λ," & vbNewLine & _
                "       A.��λ�ʱ�, A.��ͥ��ַ, A.��ͥ�绰, A.��ͥ��ַ�ʱ�, A.�����, A.���֤��, A.�ѱ�, A.ҽ�Ƹ��ʽ, A.����, A.����״��," & vbNewLine & _
                "       A.����, A.ְҵ,decode(a.����ID,Null,b.����,Zl_Age_Calc(A.����ID)) as ����1 " & vbNewLine & _
                "From ������Ϣ A, ������ü�¼ B, �Ա� C" & vbNewLine & _
                "Where B.����id = A.����id(+) And B.�Ա� = C.����(+) And Mod(B.��¼����,10) = 1 And B.��¼״̬ In (1, 3) And B.��� = 1 And" & vbNewLine & _
                "      B.NO = [3] " & vbNewLine

'        strSQLbak = strsql
'        strSQLbak = Replace$(strSQLbak, "סԺ���ü�¼", "������ü�¼")
'        strSQLbak = Replace$(strSQLbak, "Decode(B.��ҳid, Null, 1, 2) As Patienttype", "1 As Patienttype")
'        strSQLbak = Replace$(strSQLbak, "Nvl(B.��ҳid, 0)", "0")
'        strsql = strsql & " union all " & strSQLbak & " Order By ����id "
    ElseIf iInputType = 7 Then '����ĸ�ľ��￨

        strSQL = "Select Distinct Decode(a.��ǰ����id, Null, 1, 2) As Patienttype, A.��ҳid," & vbNewLine & _
                "                               Decode(a.��ǰ����id, Null, Nvl(b.ִ�в���id, 0), a.��ǰ����id) As ���˿���, b.ִ���� As ҽ��, a.����," & vbNewLine & _
                "                               Decode(c.����, Null, a.�Ա�, c.���� || '-' || a.�Ա�) As �Ա�, a.����, a.����id, a.סԺ��, a.�����," & vbNewLine & _
                "                               a.��ǰ����,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
                "From ������Ϣ a, (Select ִ�в���id, ִ����, ����id, �����,��¼״̬,��¼���� From ���˹Һż�¼ Where �Ǽ�ʱ�� > Sysdate - [5]) b, �Ա� c" & vbNewLine & _
                "Where a.���￨�� = [1]  And a.����id = b.����id(+) and (b.����ID is null or (b.��¼״̬=1 and b.��¼���� =1)) And a.����� = b.�����(+) And a.�Ա� = c.����(+)"

    ElseIf iInputType = 8 Then 'ҽ����
        strSQL = "Select Distinct Decode(a.��ǰ����id, Null, 1, 2) As Patienttype, A.��ҳid," & vbNewLine & _
                "                               Decode(a.��ǰ����id, Null, Nvl(b.ִ�в���id, 0), a.��ǰ����id) As ���˿���, b.ִ���� As ҽ��, a.����," & vbNewLine & _
                "                               Decode(c.����, Null, a.�Ա�, c.���� || '-' || a.�Ա�) As �Ա�, a.����, a.����id, a.סԺ��, a.�����," & vbNewLine & _
                "                               a.��ǰ����,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
                "From ������Ϣ a, (Select ִ�в���id, ִ����, ����id, �����,��¼״̬,��¼���� From ���˹Һż�¼ Where �Ǽ�ʱ�� > Sysdate - [5]) b, �Ա� c" & vbNewLine & _
                "Where (a.ҽ���� = [1] or a.IC����= [1]) And a.����id = b.����id(+) and (b.����ID is null or (b.��¼����=1 and b.��¼״̬=1)) And a.����� = b.�����(+) And a.�Ա� = c.����(+)"
    ElseIf iInputType = 9 Then '���֤
        strSQL = "Select Distinct Decode(a.��ǰ����id, Null, 1, 2) As Patienttype, A.��ҳid," & vbNewLine & _
                "                               Decode(a.��ǰ����id, Null, Nvl(b.ִ�в���id, 0), a.��ǰ����id) As ���˿���, b.ִ���� As ҽ��, a.����," & vbNewLine & _
                "                               Decode(c.����, Null, a.�Ա�, c.���� || '-' || a.�Ա�) As �Ա�, a.����, a.����id, a.סԺ��, a.�����," & vbNewLine & _
                "                               a.��ǰ����,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
                "From ������Ϣ a, (Select ִ�в���id, ִ����, ����id, �����,��¼״̬,��¼���� From ���˹Һż�¼ Where �Ǽ�ʱ�� > Sysdate - [5]) b, �Ա� c" & vbNewLine & _
                "Where a.���֤�� = [1]  And a.����id = b.����id(+) And a.����� = b.�����(+) and (b.����ID is null or (b.��¼״̬ =1 and b.��¼���� =1)) And a.�Ա� = c.����(+)"

    Else '��������
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,A.��ҳID,Nvl(A.��ǰ����id,Nvl(C.�������ID,0)) As ���˿���," & _
            "a.����,decode(b.����,null,a.�Ա�,b.���� || '-' || a.�Ա�) as �Ա�,a.����,a.����id,a.סԺ��,a.�����,a.��ǰ����,'' as ҽ��,Zl_Age_Calc(A.����ID) as ����1  " & _
            " From ������Ϣ A , �Ա� B,����걾��¼ C Where A.����id=[4] And a.�Ա� = b.����(+) And a.����id = c.����id(+)"
    End If

    Set GetPatient = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSeek, Val(strSeek), strNO, mlng����ID, lng�Һ�Ч��)
    GetPatient.filter = ""
    If GetPatient.RecordCount > 1 Then
        If iInputType = 3 Or iInputType = 0 Or iInputType >= 7 Then
            '����Ų�ѯʱ��һ�����˹��˶���ţ�����Ҫѡ��.

            If iInputType = 0 Then
                strSQL = _
                    "Select Distinct Decode(A.��ǰ����id, Null, 1, 2) As Patienttype, A.��ҳid," & vbNewLine & _
                    "                Decode(A.��ǰ����id, Null, Nvl(B.ִ�в���id, 0), A.��ǰ����id) As ���˿���, B.ִ���� As ҽ��, A.����," & vbNewLine & _
                    "                Decode(C.����, Null, A.�Ա�, C.���� || '-' || A.�Ա�) As �Ա�, A.����, A.����id, A.סԺ��, A.�����, A.��ǰ����,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
                    "From (Select " & gConst_������Ϣ_���� & " From ������Ϣ a Where ���￨�� = [1]" & vbNewLine & _
                    "       Union " & vbNewLine & _
                    "       Select " & gConst_������Ϣ_���� & " From ������Ϣ a Where ����� = [2]" & vbNewLine & _
                    "       Union " & vbNewLine & _
                    "       Select " & gConst_������Ϣ_���� & " From ������Ϣ a Where סԺ�� = [2]) A, ���˹Һż�¼ B, �Ա� C" & vbNewLine & _
                    "Where A.����id = B.����id(+) And A.����� = B.�����(+) and (b.����� is null or (b.��¼״̬ =1 and b.��¼����=1)) And 0 + B.�Ǽ�ʱ��(+) > Sysdate - [5] And  A.�Ա� = C.����(+) And B.NO = [6]"

'                strSQL = "Select Distinct Decode(a.��ǰ����id, Null, 1, 2) As Patienttype, Nvl(a.סԺ����, 0) As ��ҳid," & vbNewLine & _
'                        "                               Decode(a.��ǰ����id, Null, Nvl(b.ִ�в���id, 0), a.��ǰ����id) As ���˿���, b.ִ���� As ҽ��, a.����," & vbNewLine & _
'                        "                               Decode(c.����, Null, a.�Ա�, c.���� || '-' || a.�Ա�) As �Ա�, a.����, a.����id, a.סԺ��, a.�����," & vbNewLine & _
'                        "                               a.��ǰ����" & vbNewLine & _
'                        "From ������Ϣ a, (Select NO,ִ�в���id, ִ����, ����id, ����� From ���˹Һż�¼ Where �Ǽ�ʱ�� > Sysdate - [5]) b, �Ա� c" & vbNewLine & _
'                        "Where (a.���￨�� = [1] Or a.����� =[2] Or a.סԺ��=[2]) And a.����id = b.����id(+) And a.����� = b.�����(+) And a.�Ա� = c.����(+) And B.NO=[6]"

            ElseIf iInputType = 3 Then
                strSQL = "Select Distinct Decode(a.��ǰ����id, Null, 1, 2) As Patienttype, A.��ҳid," & vbNewLine & _
                        "                               Decode(a.��ǰ����id, Null, Nvl(b.ִ�в���id, 0), a.��ǰ����id) As ���˿���, b.ִ���� As ҽ��, a.����," & vbNewLine & _
                        "                               Decode(c.����, Null, a.�Ա�, c.���� || '-' || a.�Ա�) As �Ա�, a.����, a.����id, a.סԺ��, a.�����," & vbNewLine & _
                        "                               a.��ǰ����,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
                        "From ������Ϣ a, (Select NO,ִ�в���id, ִ����, ����id, �����,��¼״̬,��¼���� From ���˹Һż�¼ Where �Ǽ�ʱ�� > Sysdate - [5]) b, �Ա� c" & vbNewLine & _
                        "Where a.����� = [2] And a.����id = b.����id(+) And a.����� = b.�����(+) and (b.����ID is null or(b.��¼״̬=1 and ��¼����=1)) And a.�Ա� = c.����(+) And B.NO=[6]"
            ElseIf iInputType = 7 Then
                strSQL = "Select Distinct Decode(a.��ǰ����id, Null, 1, 2) As Patienttype, A.��ҳid," & vbNewLine & _
                        "                               Decode(a.��ǰ����id, Null, Nvl(b.ִ�в���id, 0), a.��ǰ����id) As ���˿���, b.ִ���� As ҽ��, a.����," & vbNewLine & _
                        "                               Decode(c.����, Null, a.�Ա�, c.���� || '-' || a.�Ա�) As �Ա�, a.����, a.����id, a.סԺ��, a.�����," & vbNewLine & _
                        "                               a.��ǰ����,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
                        "From ������Ϣ a, (Select NO,ִ�в���id, ִ����, ����id, �����,��¼״̬,��¼���� From ���˹Һż�¼ Where �Ǽ�ʱ�� > Sysdate - [5]) b, �Ա� c" & vbNewLine & _
                        "Where a.���￨�� = [1] And a.����id = b.����id(+) And a.����� = b.�����(+) and (b.����ID is null or (b.��¼״̬=1 and b.��¼���� =1)) And a.�Ա� = c.����(+) And B.NO=[6]"
            ElseIf iInputType = 8 Then

                strSQL = "Select Distinct Decode(a.��ǰ����id, Null, 1, 2) As Patienttype, A.��ҳid," & vbNewLine & _
                        "                               Decode(a.��ǰ����id, Null, Nvl(b.ִ�в���id, 0), a.��ǰ����id) As ���˿���, b.ִ���� As ҽ��, a.����," & vbNewLine & _
                        "                               Decode(c.����, Null, a.�Ա�, c.���� || '-' || a.�Ա�) As �Ա�, a.����, a.����id, a.סԺ��, a.�����," & vbNewLine & _
                        "                               a.��ǰ����,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
                        "From ������Ϣ a, (Select NO,ִ�в���id, ִ����, ����id, �����,��¼״̬,��¼���� From ���˹Һż�¼ Where �Ǽ�ʱ�� > Sysdate - [5]) b, �Ա� c" & vbNewLine & _
                        "Where (a.ҽ���� = [1] or a.IC����= [1]) And a.����id = b.����id(+) And a.����� = b.�����(+) and (b.����ID is null or (b.��¼״̬=1 and b.��¼����=1)) And a.�Ա� = c.����(+) And B.NO=[6]"
            ElseIf iInputType = 9 Then

                strSQL = "Select Distinct Decode(a.��ǰ����id, Null, 1, 2) As Patienttype, A.��ҳid," & vbNewLine & _
                        "                               Decode(a.��ǰ����id, Null, Nvl(b.ִ�в���id, 0), a.��ǰ����id) As ���˿���, b.ִ���� As ҽ��, a.����," & vbNewLine & _
                        "                               Decode(c.����, Null, a.�Ա�, c.���� || '-' || a.�Ա�) As �Ա�, a.����, a.����id, a.סԺ��, a.�����," & vbNewLine & _
                        "                               a.��ǰ����,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
                        "From ������Ϣ a, (Select NO,ִ�в���id, ִ����, ����id, �����,��¼״̬,��¼���� From ���˹Һż�¼ Where �Ǽ�ʱ�� > Sysdate - [5]) b, �Ա� c" & vbNewLine & _
                        "Where  a.���֤�� = [1] And a.����id = b.����id(+) And a.����� = b.�����(+) and (b.����ID is null or (b.��¼״̬=1 and b.��¼����=1)) And a.�Ա� = c.����(+) And B.NO=[6]"

            End If
            '---- ����ѡ����
            If iInputType = 0 Then
                strTmp = _
                    "Select Distinct Rownum As ID, B.NO As �Һŵ���, B.ִ���� As ҽ��, A.�����, A.���￨��, A.סԺ��, A.����," & vbNewLine & _
                    "                Decode(A.��ǰ����id, Null, D.����, E.����) As ���˿���, Decode(C.����, Null, A.�Ա�, C.���� || '-' || A.�Ա�) As �Ա�, A.����," & vbNewLine & _
                    "                To_Char(B.�Ǽ�ʱ��, 'yyyy-MM-dd HH24:MI:SS') As �Һ�ʱ��, A.����id,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
                    "From (Select " & gConst_������Ϣ_���� & " From ������Ϣ a Where ���￨�� = [1]" & vbNewLine & _
                    "       Union " & vbNewLine & _
                    "       Select " & gConst_������Ϣ_���� & " From ������Ϣ a Where ����� = [2]" & vbNewLine & _
                    "       Union " & vbNewLine & _
                    "       Select " & gConst_������Ϣ_���� & " From ������Ϣ a Where סԺ�� = [2]) A, ���˹Һż�¼ B, �Ա� C, ���ű� D, ���ű� E" & vbNewLine & _
                    "Where A.����id = B.����id(+) and (b.����id is null or (b.��¼״̬=1 and b.��¼����=1)) And A.����� = B.�����(+) And 0 + B.�Ǽ�ʱ��(+) > Sysdate - [3] And A.�Ա� = C.����(+) And" & vbNewLine & _
                    "      Nvl(B.ִ�в���id, 0) = D.ID(+) And A.��ǰ����id = E.ID(+)" & vbNewLine & _
                    "Order By B.NO Desc"

'                strTmp = "Select Distinct rownum As ID, B.NO As �Һŵ���, B.ִ���� As ҽ��, A.�����, A.���￨��, A.סԺ��, A.����," & vbNewLine & _
'                        "                Decode(A.��ǰ����id, Null, D.����, E.����) As ���˿���, Decode(C.����, Null, A.�Ա�, C.���� || '-' || A.�Ա�) As �Ա�, A.����," & vbNewLine & _
'                        "                To_Char(B.�Ǽ�ʱ��, 'yyyy-MM-dd HH24:MI:SS') As �Һ�ʱ��, A.����id" & vbNewLine & _
'                        "From ������Ϣ A, (Select �Ǽ�ʱ��, NO, ִ�в���id, ִ����, ����id, ����� From ���˹Һż�¼ Where �Ǽ�ʱ�� > Sysdate - [3]) B, �Ա� C, ���ű� D, ���ű� E" & vbNewLine & _
'                        "Where (A.���￨�� = [1] Or A.����� = [2] Or סԺ��=[2]) And A.����id = B.����id(+) And A.����� = B.�����(+) And A.�Ա� = C.����(+) And" & vbNewLine & _
'                        "      Nvl(B.ִ�в���id, 0) = D.ID(+) And A.��ǰ����id = E.ID(+)" & vbNewLine & _
'                        "Order By B.NO Desc"

            ElseIf iInputType = 3 Then
                strTmp = "Select Distinct rownum As ID, B.NO As �Һŵ���, B.ִ���� As ҽ��, A.�����, A.���￨��, A.סԺ��, A.����," & vbNewLine & _
                        "                Decode(A.��ǰ����id, Null, D.����, E.����) As ���˿���, Decode(C.����, Null, A.�Ա�, C.���� || '-' || A.�Ա�) As �Ա�, A.����," & vbNewLine & _
                        "                To_Char(B.�Ǽ�ʱ��, 'yyyy-MM-dd HH24:MI:SS') As �Һ�ʱ��, A.����id,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
                        "From ������Ϣ A, (Select �Ǽ�ʱ��, NO, ִ�в���id, ִ����, ����id, �����,��¼״̬,��¼���� From ���˹Һż�¼ Where �Ǽ�ʱ�� > Sysdate - [3]) B, �Ա� C, ���ű� D, ���ű� E" & vbNewLine & _
                        "Where A.����� = [2] And A.����id = B.����id(+) And A.����� = B.�����(+) and (b.����ID is null or (b.��¼״̬=1 and b.��¼���� =1)) And A.�Ա� = C.����(+) And" & vbNewLine & _
                        "      Nvl(B.ִ�в���id, 0) = D.ID(+) And A.��ǰ����id = E.ID(+)" & vbNewLine & _
                        "Order By B.NO Desc"

            ElseIf iInputType = 7 Then
                strTmp = "Select Distinct rownum As ID, B.NO As �Һŵ���, B.ִ���� As ҽ��, A.�����, A.���￨��, A.סԺ��, A.����," & vbNewLine & _
                        "                Decode(A.��ǰ����id, Null, D.����, E.����) As ���˿���, Decode(C.����, Null, A.�Ա�, C.���� || '-' || A.�Ա�) As �Ա�, A.����," & vbNewLine & _
                        "                To_Char(B.�Ǽ�ʱ��, 'yyyy-MM-dd HH24:MI:SS') As �Һ�ʱ��, A.����id,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
                        "From ������Ϣ A, (Select �Ǽ�ʱ��, NO, ִ�в���id, ִ����, ����id, �����,��¼״̬,��¼���� From ���˹Һż�¼ Where �Ǽ�ʱ�� > Sysdate - [3]) B, �Ա� C, ���ű� D, ���ű� E" & vbNewLine & _
                        "Where A.���￨�� = [1] And A.����id = B.����id(+) And A.����� = B.�����(+) and (b.����ID is null or (b.��¼״̬=1 and b.��¼����=1)) And A.�Ա� = C.����(+) And" & vbNewLine & _
                        "      Nvl(B.ִ�в���id, 0) = D.ID(+) And A.��ǰ����id = E.ID(+)" & vbNewLine & _
                        "Order By B.NO Desc"
            ElseIf iInputType = 8 Then
                strTmp = "Select Distinct rownum As ID, B.NO As �Һŵ���, B.ִ���� As ҽ��, A.�����, A.���￨��, A.סԺ��, A.����," & vbNewLine & _
                        "                Decode(A.��ǰ����id, Null, D.����, E.����) As ���˿���, Decode(C.����, Null, A.�Ա�, C.���� || '-' || A.�Ա�) As �Ա�, A.����," & vbNewLine & _
                        "                To_Char(B.�Ǽ�ʱ��, 'yyyy-MM-dd HH24:MI:SS') As �Һ�ʱ��, A.����id,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
                        "From ������Ϣ A, (Select �Ǽ�ʱ��, NO, ִ�в���id, ִ����, ����id, �����,��¼״̬,��¼���� From ���˹Һż�¼ Where �Ǽ�ʱ�� > Sysdate - [3]) B, �Ա� C, ���ű� D, ���ű� E" & vbNewLine & _
                        "Where (A.ҽ���� = [1] or a.IC����= [1]) And A.����id = B.����id(+) And A.����� = B.�����(+) and (b.����ID is null or (b.��¼״̬=1 and b.��¼����=1)) And A.�Ա� = C.����(+) And" & vbNewLine & _
                        "      Nvl(B.ִ�в���id, 0) = D.ID(+) And A.��ǰ����id = E.ID(+)" & vbNewLine & _
                        "Order By B.NO Desc"
            ElseIf iInputType = 9 Then
                strTmp = "Select Distinct rownum As ID, B.NO As �Һŵ���, B.ִ���� As ҽ��, A.�����, A.���￨��, A.סԺ��, A.����," & vbNewLine & _
                        "                Decode(A.��ǰ����id, Null, D.����, E.����) As ���˿���, Decode(C.����, Null, A.�Ա�, C.���� || '-' || A.�Ա�) As �Ա�, A.����," & vbNewLine & _
                        "                To_Char(B.�Ǽ�ʱ��, 'yyyy-MM-dd HH24:MI:SS') As �Һ�ʱ��, A.����id,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
                        "From ������Ϣ A, (Select �Ǽ�ʱ��, NO, ִ�в���id, ִ����, ����id, �����,��¼״̬,��¼���� From ���˹Һż�¼ Where �Ǽ�ʱ�� > Sysdate - [3]) B, �Ա� C, ���ű� D, ���ű� E" & vbNewLine & _
                        "Where A.���֤�� = [1] And A.����id = B.����id(+) And A.����� = B.�����(+) and (b.����ID is null or (b.��¼״̬=1 and b.��¼����=1)) And A.�Ա� = C.����(+) And" & vbNewLine & _
                        "      Nvl(B.ִ�в���id, 0) = D.ID(+) And A.��ǰ����id = E.ID(+)" & vbNewLine & _
                        "Order By B.NO Desc"
            End If
            mblnEdit = False
            Call ClientToScreen(txt����.hWnd, objPoint)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strTmp, 0, "����ѡ��", True, "", "", True, True, True, objPoint.X * 15, objPoint.Y * 15, Me.txt����.Height, _
                                                    False, True, False, strSeek, Val(strSeek), lng�Һ�Ч��)
            mblnEdit = True
            GetPatient.filter = "����ID=0"
            If Not rsTmp Is Nothing Then
                If rsTmp.State = adStateOpen Then
                    If rsTmp.RecordCount = 1 Then
                        str�Һŵ� = "" & Nvl(rsTmp.Fields("�Һŵ���"), 0)
                        If str�Һŵ� = "0" Then
                            strSQL = Replace(strSQL, " And B.NO = [6]", " And A.����ID = [6]")
                            str�Һŵ� = rsTmp.Fields("����ID")
                            Set GetPatient = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSeek, Val(strSeek), strNO, mlng����ID, lng�Һ�Ч��, Val(str�Һŵ�))
                        Else
                            Set GetPatient = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSeek, Val(strSeek), strNO, mlng����ID, lng�Һ�Ч��, str�Һŵ�)
                        End If

                        If GetPatient.RecordCount > 1 Then
                            MsgBox "�鵽���˶�����ˣ������ǰ׺��־���в���!"
                            GetPatient.filter = "����ID=0"
                        End If
                    End If

                End If
            End If
        Else
            MsgBox "�鵽���˶�����ˣ������ǰ׺��־���в���!"
            GetPatient.filter = "����ID=0"
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function OpenSelect(ByVal strText As String, Optional ByVal blnWhere As Boolean = False) As Byte
    '--------------------------------------------------------------------------------------------------------
    '����:���б�ṹ��������鵥
    '����:strText       ���˹ؼ���(����ΪסԺ�ţ�����ţ���λ�Ż������)
    '����:0             ȡ������
    '     1             �ɹ�����
    '     2             ������
    '     3             ���ֺ����в���������ʱ��1.����������ͬ�ı걾.2.Ӥ��ҽ����ĸ��ҽ����һ��
    '--------------------------------------------------------------------------------------------------------
    Dim strInput As String, i As Integer
    Dim rs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim strLvw As String, strSQL As String
    Dim objPoint As POINTAPI
    Dim strLastKeys As String
    Dim strStart As String, strEnd As String
    Dim strField As String
    Dim blnNoRange As Boolean, blnCheck As Boolean
    Dim mstrSql As String
    Dim lngCurrDevice As Long
    Dim blMachineFind As Boolean
    Dim strTmp As String
    Dim int��ҳID As Integer
    Dim intCount As Integer
    Dim strFilter As String
    Dim strDate As String
    Dim lng����ID As Long
    Dim intִ��״̬ As Integer
    Dim blnδ�շ���ʾ As Boolean
    Dim strִ�п������� As String
    Dim lngִ�п���ID As Long
    Dim strSQLbak As String
    Dim strAge As String, aAge As Variant
    Dim intPatProperty As Integer, rsPatProperty As New ADODB.Recordset   '�������۲���

    On Error GoTo ErrHand

    OpenSelect = 2

    If blnCheck Then '��ʾ�շ�
        strLvw = "����,900,0,1;�շ�,600,0,1;����ʱ��,1300,0,0;������Ŀ,1800,0,0;�������,1800,0,0;������,810,0,0"
    Else
        strLvw = "����,900,0,1;����ʱ��,1300,0,0;������Ŀ,1800,0,0;�������,1800,0,0;������,810,0,0"
    End If

    blnδ�շ���ʾ = InStr(mstrPrivs, "δ�շѺ���") > 0

    blnNoRange = Val(zlDatabase.GetPara("���պ���ʱ��", 100, 1208, 1))
    blnCheck = Val(zlDatabase.GetPara("������ʾ�շ�", 100, 1208, 1))
    blMachineFind = Val(zlDatabase.GetPara("��������Ŀ����", 100, 1208, 1))


    strStart = GetDateTime(Split(zlDatabase.GetPara("�����շ�Χ", 100, 1208, "��  ��") & ";", ";")(0), 1)
    strEnd = GetDateTime(Split(zlDatabase.GetPara("�����շ�Χ", 100, 1208, "��  ��") & ";", ";")(0), 2)

    If strStart = "�Զ���" Then
        strStart = Format(Split(zlDatabase.GetPara("�����շ�Χ", 100, 1208, "��  ��") & ";", ";")(1), "yyyy-mm-dd 00:00:00")
        strEnd = Format(Split(zlDatabase.GetPara("�����շ�Χ", 100, 1208, "��  ��") & ";", ";")(2), "yyyy-mm-dd 23:59:59")
    Else
        If strStart = "" Then strStart = GetDateTime("��  ��", 1)
        If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
    End If

    lngCurrDevice = 0
    If vsf2.Rows > 1 Then
        lngCurrDevice = Val(vsf2.RowData(1))
    End If

    If mlngDefaultDevice = -1 Then
        If blnCheck Then '��ʾ�շ�
            If mblnBarCode Then
                mstrSql = "SELECT Decode(SUM(Decode(F.��������,[2],1,0)),0,0,1) AS ѡ��"
            Else
                mstrSql = "SELECT Decode(Nvl(z.ҽ�����,0), 0, 1, Decode(Sum(Nvl(z.����,0)), 0, 0, 1)) As ѡ��"
            End If
            mstrSql = mstrSql & ",A.���ID AS ID,  " & _
                              "C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)') As ����," & _
                              "C.�����," & _
                              "C.סԺ��," & _
                              "D.���� AS �������," & _
                              "A.����ҽ�� AS ������," & _
                              "f.������,f.����ʱ��, " & _
                              "'Item' AS ͼ��,Decode(Nvl(z.ҽ�����,0), 0, '��', Decode(Sum(Nvl(z.����,0)), 0, '��', '��')) As �շ�,NVL(A.������־,0) AS ����,Y.��������,MAX(Decode(H.��Ŀ���,2,2,1)) As ��Ŀ���,MAX(F.������) AS ������,MAX(F.����ʱ��) AS ����ʱ�� " & _
                         "FROM ����ҽ����¼ A," & _
                         "������Ϣ C,���ű� D,����ҽ������ F,���鱨����Ŀ G,������Ŀ H,������ĿĿ¼ Y,סԺ���ü�¼ Z " & _
                        "WHERE A.������� = 'C' " & _
                              "AND A.����ID=C.����ID " & _
                              "AND A.��������ID=D.ID " & _
                              "AND A.���id IS NOT NULL " & _
                              "AND A.ҽ��״̬=8 AND A.ID=F.ҽ��id " & _
                              "AND A.������Ŀid=G.������Ŀid AND G.ϸ��ID Is Null " & _
                              "AND G.������ĿID=H.������ĿID " & _
                              "AND A.������ĿID=Y.ID " & _
                              IIf(mbln���۵�ģʽ = True, " and a.������Դ = 2 and c.��Ժʱ�� is null ", " ")

            mstrSql = mstrSql & " " & _
                              "AND F.ִ��״̬ =0 AND A.����ID=[1] AND A.ִ�п���id+0=[3] " & _
                              IIf(blnNoRange, " ", " AND A.����ʱ�� BETWEEN [5] and [6] ") & _
                              "AND F.NO=Z.NO(+) AND F.��¼����=mod(Z.��¼����(+),10) AND F.ҽ��id=Z.ҽ�����(+)+0 " & _
                              IIf(mblnBarCode, " And F.��������=[2]  ", "") & _
                              IIf(int��ҳID = 0, " ", " And a.��ҳID = [9] ") & _
                              " And a.������Դ = 2 " & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              "GROUP BY A.���ID,a.id,C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)'),C.�����,C.סԺ��,D.����,A.����ҽ��,'Item',NVL(A.������־,0),Y.��������,f.������,f.����ʱ�� ,z.ҽ����� "

            mstrSql = mstrSql & " Union all "

            If mblnBarCode Then
                mstrSql = mstrSql & "SELECT Decode(SUM(Decode(F.��������,[2],1,0)),0,0,1) AS ѡ��"
            Else
                mstrSql = mstrSql & "SELECT Decode(Nvl(z.ҽ�����,0), 0, 1, Decode(Sum(Nvl(z.����,0)), 0, 0, 1)) As ѡ��"
            End If

            mstrSql = mstrSql & ",A.���ID AS ID, " & _
                              "C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)') As ����," & _
                              "C.�����," & _
                              "C.סԺ��," & _
                              "D.���� AS �������," & _
                              "A.����ҽ�� AS ������," & _
                              "f.������,f.����ʱ��, " & _
                              "'Item' AS ͼ��,Decode(Nvl(z.ҽ�����,0), 0, '��', Decode(Sum(Nvl(z.����,0)), 0, '��', '��')) As �շ�,NVL(A.������־,0) AS ����,Y.��������,MAX(Decode(H.��Ŀ���,2,2,1)) As ��Ŀ���,MAX(F.������) AS ������,MAX(F.����ʱ��) AS ����ʱ�� " & _
                         "FROM ����ҽ����¼ A," & _
                         "������Ϣ C,���ű� D,����ҽ������ F,���鱨����Ŀ G,������Ŀ H,������ĿĿ¼ Y,סԺ���ü�¼ Z,����걾��¼ J " & _
                        "WHERE A.������� = 'C' " & _
                              "AND A.����ID=C.����ID " & _
                              "AND A.��������ID=D.ID " & _
                              "AND A.���id IS NOT NULL " & _
                              "AND A.ҽ��״̬=8 AND A.ID=F.ҽ��id " & _
                              "AND A.������Ŀid=G.������Ŀid AND G.ϸ��ID Is Null " & _
                              "AND G.������ĿID=H.������ĿID " & _
                              "AND A.������ĿID=Y.ID " & _
                              IIf(mbln���۵�ģʽ = True, " and a.������Դ = 2 and c.��Ժʱ�� is null ", " ")

            mstrSql = mstrSql & " " & _
                              "AND A.����ID=[1] AND A.ִ�п���id+0=[3] " & _
                              IIf(blnNoRange, " ", " AND A.����ʱ�� BETWEEN [5] and [6] ") & _
                              "AND F.NO=Z.NO(+) AND F.��¼����=mod(Z.��¼����(+),10) AND F.ҽ��id=Z.ҽ�����(+)+0 And a.���id = j.ҽ��ID(+) And j.id = [8] " & _
                              IIf(mblnBarCode, " And F.��������=[2]   ", "") & _
                              IIf(int��ҳID = 0, " ", " And a.��ҳID = [9] ") & _
                              " And a.������Դ = 2 " & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              "GROUP BY A.���ID,a.id,C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)'),C.�����,C.סԺ��,D.����,A.����ҽ��,'Item',NVL(A.������־,0),Y.��������,f.������,f.����ʱ�� ,z.ҽ����� "
        Else
            If mblnBarCode Then
                mstrSql = mstrSql & "SELECT Distinct Decode 1 AS ѡ��"
            Else
                mstrSql = mstrSql & "SELECT Distinct 1 AS ѡ��"
            End If
            mstrSql = mstrSql & ",A.���ID AS ID, " & _
                              "C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)') As ����," & _
                              "C.�����," & _
                              "C.סԺ��," & _
                              "D.���� AS �������," & _
                              "A.����ҽ�� AS ������," & _
                              "f.������,f.����ʱ��, " & _
                              "'Item' AS ͼ��,Decode(Nvl(z.ҽ�����,0), 0, '��', Decode(Sum(Nvl(z.����,0)), 0, '��', '��')) As �շ�,NVL(A.������־,0) AS ����,Y.��������,Decode(H.��Ŀ���,2,2,1) As ��Ŀ���,F.������,F.����ʱ�� As ����ʱ�� " & _
                         "FROM ����ҽ����¼ A," & _
                         "������Ϣ C,���ű� D,����ҽ������ F,���鱨����Ŀ G,������Ŀ H,������ĿĿ¼ Y,סԺ���ü�¼ Z " & _
                        "WHERE A.������� = 'C' " & _
                              "AND A.����ID=C.����ID " & _
                              "AND A.��������ID=D.ID " & _
                              "AND A.���id IS NOT NULL " & _
                              "AND A.ҽ��״̬=8 AND A.ID=F.ҽ��id " & _
                              "AND A.������Ŀid=G.������Ŀid AND G.ϸ��ID Is Null " & _
                              "AND G.������ĿID=H.������ĿID " & _
                              "AND A.������ĿID=Y.ID " & _
                              "And a.������Դ = 2 " & _
                              IIf(mbln���۵�ģʽ = True, " and a.������Դ = 2 and c.��Ժʱ�� is null ", " ")

            mstrSql = mstrSql & " " & _
                              "AND F.ִ��״̬ = 0 AND A.����ID=[1] AND A.ִ�п���id+0=[3] " & _
                              "AND F.NO=Z.NO(+) AND F.��¼����=mod(Z.��¼����(+),10) AND F.ҽ��id=Z.ҽ�����(+)+0 " & _
                              IIf(mblnBarCode, " And F.��������=[2]   ", "") & _
                              IIf(int��ҳID = 0, " ", " And a.��ҳID = [9] ") & _
                              IIf(blnNoRange, " ", " AND A.����ʱ�� BETWEEN [5] and [6] ") & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              " GROUP BY A.���ID,a.id,C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)'),C.�����,C.סԺ��,D.����,A.����ҽ��,'Item',NVL(A.������־,0),Y.��������,f.������,f.����ʱ��,Decode(H.��Ŀ���,2,2,1),F.������,F.����ʱ�� ,z.ҽ����� "

            mstrSql = mstrSql & " Union all "

            If mblnBarCode Then
                mstrSql = mstrSql & "SELECT Distinct 1 AS ѡ��"
            Else
                mstrSql = mstrSql & "SELECT Distinct 1 AS ѡ��"
            End If

            mstrSql = mstrSql & ",A.���ID AS ID, " & _
                              "C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)') As ����," & _
                              "C.�����," & _
                              "C.סԺ��," & _
                              "D.���� AS �������," & _
                              "A.����ҽ�� AS ������," & _
                              "f.������,f.����ʱ��, " & _
                              "'Item' AS ͼ��,Decode(Nvl(z.ҽ�����,0), 0, '��', Decode(Sum(Nvl(z.����,0)), 0, '��', '��')) As �շ�,NVL(A.������־,0) AS ����,Y.��������,Decode(H.��Ŀ���,2,2,1) As ��Ŀ���,F.������,F.����ʱ�� As ����ʱ�� " & _
                         "FROM ����ҽ����¼ A," & _
                         "������Ϣ C,���ű� D,����ҽ������ F,���鱨����Ŀ G,������Ŀ H,������ĿĿ¼ Y,����걾��¼ j,סԺ���ü�¼ Z " & _
                        "WHERE A.������� = 'C' " & _
                              "AND A.����ID=C.����ID " & _
                              "AND A.��������ID=D.ID " & _
                              "AND A.���id IS NOT NULL " & _
                              "AND A.ҽ��״̬=8 AND A.ID=F.ҽ��id " & _
                              "AND A.������Ŀid=G.������Ŀid AND G.ϸ��ID Is Null " & _
                              "AND G.������ĿID=H.������ĿID " & _
                              "AND A.������ĿID=Y.ID " & _
                              "And a.������Դ = 2 " & _
                              IIf(mbln���۵�ģʽ = True, " and a.������Դ = 2 and c.��Ժʱ�� is null ", " ")

            mstrSql = mstrSql & " " & _
                              "AND A.����ID=[1] AND A.ִ�п���id+0=[3] And a.���id = j.ҽ��id(+) and j.id = [8] " & _
                              "AND F.NO=Z.NO(+) AND F.��¼����=mod(Z.��¼����(+),10) AND F.ҽ��id=Z.ҽ�����(+)+0 " & _
                              IIf(mblnBarCode, " And F.��������=[2]   ", "") & _
                              IIf(int��ҳID = 0, " ", " And a.��ҳID = [9] ") & _
                              IIf(blnNoRange, " ", " AND A.����ʱ�� BETWEEN [5] and [6] ") & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              " GROUP BY A.���ID,a.id,C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)'),C.�����,C.סԺ��,D.����,A.����ҽ��,'Item',NVL(A.������־,0),Y.��������,f.������,f.����ʱ��,Decode(H.��Ŀ���,2,2,1),F.������,F.����ʱ�� ,z.ҽ����� "
        End If
    Else
        If blnCheck Then '��ʾ�շ�
            If mblnBarCode Then
                mstrSql = mstrSql & "SELECT Decode(SUM(Decode(F.��������,[2],1,0)),0,0,1) AS ѡ��"
            Else
                mstrSql = mstrSql & "SELECT Decode(Nvl(z.ҽ�����,0), 0, 1, Decode(Sum(Nvl(z.����,0)), 0, 0, 1)) As ѡ��"
            End If
            mstrSql = mstrSql & ",A.���ID AS ID, " & _
                              "C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)') As ����," & _
                              "C.�����," & _
                              "C.סԺ��," & _
                              "D.���� AS �������," & _
                              "A.����ҽ�� AS ������," & _
                              "f.������,f.����ʱ��, " & _
                              "'Item' AS ͼ��,Decode(Nvl(z.ҽ�����,0), 0, '��', Decode(Sum(Nvl(z.����,0)), 0, '��', '��')) As �շ�,NVL(A.������־,0) AS ����,H.��������,MAX(Decode(I.��Ŀ���,2,2,1)) As ��Ŀ���,MAX(F.������) AS ������,MAX(F.����ʱ��) AS ����ʱ�� " & _
                         "FROM ����ҽ����¼ A," & _
                         "������Ϣ C,���ű� D,����ҽ������ F,���鱨����Ŀ G,������ĿĿ¼ H,������Ŀ I,����������Ŀ Y,סԺ���ü�¼ Z " & _
                        "WHERE A.������� = 'C' " & _
                              "AND A.����ID=C.����ID " & _
                              "AND A.��������ID=D.ID " & _
                              "AND A.���id IS NOT NULL " & _
                              "AND A.ҽ��״̬=8 AND A.ID=F.ҽ��id " & _
                              "AND A.������Ŀid=G.������Ŀid AND G.ϸ��ID Is Null " & _
                              IIf(blMachineFind, "AND G.������Ŀid=Y.��Ŀid ", "AND G.������Ŀid=Y.��Ŀid(+) ") & _
                              "AND G.������ĿID=I.������ĿID " & _
                              "AND A.������ĿID=H.ID " & _
                              IIf(mbln���۵�ģʽ = True, " and a.������Դ = 2 and c.��Ժʱ�� is null ", " ") & _
                              IIf(mlngDefaultDevice = 0 And lngCurrDevice = 0, "", "AND (Y.����ID+0=[7] Or Y.����ID Is Null)") & _
                              "AND F.ִ��״̬ = 0 AND A.����ID=[1] AND A.ִ�п���id+0=[3] " & _
                              IIf(mblnBarCode, " And F.��������=[2]   ", "")

            mstrSql = mstrSql & " " & _
                              IIf(int��ҳID = 0, " ", " And a.��ҳID = [9] ") & _
                              IIf(blnNoRange, " ", " AND A.����ʱ�� BETWEEN [5] and [6] ") & _
                              "AND F.NO=Z.NO(+) AND F.��¼����=mod(Z.��¼����(+),10) AND F.ҽ��id=Z.ҽ�����(+)+0 " & _
                              "And a.������Դ = 2 " & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              "GROUP BY A.���ID,a.id,C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)'),C.�����,C.סԺ��,D.����,A.����ҽ��,'Item',NVL(A.������־,0),H.��������,f.������,f.����ʱ�� ,z.ҽ����� "

            mstrSql = mstrSql & " Union all "

            If mblnBarCode Then
                mstrSql = mstrSql & "SELECT Decode(SUM(Decode(F.��������,[2],1,0)),0,0,1) AS ѡ��"
            Else
                mstrSql = mstrSql & "SELECT Decode(Nvl(z.ҽ�����,0), 0, 1, Decode(Sum(Nvl(z.����,0)), 0, 0, 1)) As ѡ��"
            End If

            mstrSql = mstrSql & ",A.���ID AS ID, " & _
                              "C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)') As ����," & _
                              "C.�����," & _
                              "C.סԺ��," & _
                              "D.���� AS �������," & _
                              "A.����ҽ�� AS ������," & _
                              "f.������,f.����ʱ��, " & _
                              "'Item' AS ͼ��,Decode(Nvl(z.ҽ�����,0), Null, '��', Decode(Sum(Nvl(z.����,0)), 0, '��', '��')) As �շ�,NVL(A.������־,0) AS ����,H.��������,MAX(Decode(I.��Ŀ���,2,2,1)) As ��Ŀ���,MAX(F.������) AS ������,MAX(F.����ʱ��) AS ����ʱ�� " & _
                         "FROM ����ҽ����¼ A," & _
                         "������Ϣ C,���ű� D,����ҽ������ F,���鱨����Ŀ G,������ĿĿ¼ H,������Ŀ I,����������Ŀ Y,סԺ���ü�¼ Z,����걾��¼ j,������Ŀ�ֲ� k " & _
                        "WHERE A.������� = 'C' " & _
                              "AND A.����ID=C.����ID " & _
                              "AND A.��������ID=D.ID " & _
                              "AND A.���id IS NOT NULL " & _
                              "AND A.ҽ��״̬=8 AND A.ID=F.ҽ��id " & _
                              "AND A.������Ŀid=G.������Ŀid AND G.ϸ��ID Is Null " & _
                              IIf(blMachineFind, "AND G.������Ŀid=Y.��Ŀid ", "AND G.������Ŀid=Y.��Ŀid(+) ") & _
                              "AND G.������ĿID=I.������ĿID " & _
                              "AND A.������ĿID=H.ID " & _
                              IIf(mbln���۵�ģʽ = True, " and a.������Դ = 2 and c.��Ժʱ�� is null ", " ") & _
                              IIf(mlngDefaultDevice = 0 And lngCurrDevice = 0, "", "AND (Y.����ID+0=[7] Or Y.����ID Is Null)") & _
                              "AND A.����ID=[1] AND A.ִ�п���id+0=[3] " & _
                              IIf(mblnBarCode, " And F.��������=[2]   ", "")

            mstrSql = mstrSql & " " & _
                              IIf(int��ҳID = 0, " ", " And a.��ҳID = [9] ") & _
                              IIf(blnNoRange, " ", " AND A.����ʱ�� BETWEEN [5] and [6] ") & _
                              "AND F.NO=Z.NO(+) AND F.��¼����=mod(Z.��¼����(+),10) AND F.ҽ��id=Z.ҽ�����(+)+0 and a.���id = k.ҽ��id(+) and j.id =k.�걾ID and j.id = [8] " & _
                              "And a.������Դ = 2 " & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              "GROUP BY A.���ID,a.id,C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)'),C.�����,C.סԺ��,D.����,A.����ҽ��,'Item',NVL(A.������־,0),H.��������,f.������,f.����ʱ�� ,z.ҽ����� "
        Else
            If mblnBarCode Then
                mstrSql = mstrSql & "SELECT Distinct 1 AS ѡ��"
            Else
                mstrSql = mstrSql & "SELECT Distinct 1 AS ѡ��"
            End If
            mstrSql = mstrSql & ",A.���ID AS ID, " & _
                              "C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)') As ����," & _
                              "C.�����," & _
                              "C.סԺ��," & _
                              "D.���� AS �������," & _
                              "A.����ҽ�� AS ������," & _
                              "f.������,f.����ʱ��, " & _
                              "'Item' AS ͼ��,Decode(Nvl(z.ҽ�����,0), 0, '��', Decode(Sum(Nvl(z.����,0)), 0, '��', '��')) As �շ�,NVL(A.������־,0) AS ����,H.��������,Decode(I.��Ŀ���,2,2,1) As ��Ŀ���,F.������,F.����ʱ�� As ����ʱ�� " & _
                         "FROM ����ҽ����¼ A," & _
                         "������Ϣ C,���ű� D,����ҽ������ F,���鱨����Ŀ G,������ĿĿ¼ H,������Ŀ I,����������Ŀ Y,סԺ���ü�¼ Z " & _
                        "WHERE A.������� = 'C' " & _
                              "AND A.����ID=C.����ID " & _
                              "AND A.��������ID=D.ID " & _
                              "AND A.���id IS NOT NULL " & _
                              "AND A.ҽ��״̬=8 AND A.ID=F.ҽ��id " & _
                              "AND A.������Ŀid=G.������Ŀid AND G.ϸ��ID Is Null " & _
                              IIf(blMachineFind, "AND G.������Ŀid=Y.��Ŀid ", "AND G.������Ŀid=Y.��Ŀid(+) ") & _
                              "AND G.������ĿID=I.������ĿID " & _
                              "AND A.������ĿID=H.ID " & _
                              "And a.������Դ = 2 " & _
                              IIf(mbln���۵�ģʽ = True, " and a.������Դ = 2 and c.��Ժʱ�� is null ", " ")

            mstrSql = mstrSql & " " & _
                              IIf(mlngDefaultDevice = 0 And lngCurrDevice = 0, "", "AND (Y.����ID+0=[7] Or Y.����ID Is Null)") & _
                              "AND F.ִ��״̬ = 0 AND A.����ID=[1] AND A.ִ�п���id+0=[3] " & _
                              "AND F.NO=Z.NO(+) AND F.��¼����=mod(Z.��¼����(+),10) AND F.ҽ��id=Z.ҽ�����(+)+0 " & _
                              IIf(mblnBarCode, " And F.��������=[2]   ", "") & _
                              IIf(int��ҳID = 0, " ", " And a.��ҳID = [9] ") & _
                              IIf(blnNoRange, " ", " AND A.����ʱ�� BETWEEN [5] and [6] ") & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              " GROUP BY A.���ID,a.id,C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)'),C.�����,C.סԺ��,D.����,A.����ҽ��,'Item',NVL(A.������־,0),H.��������,f.������,f.����ʱ��,Decode(I.��Ŀ���,2,2,1),F.������,F.����ʱ�� ,z.ҽ����� "

            mstrSql = mstrSql & " Union all "

            If mblnBarCode Then
                mstrSql = mstrSql & "SELECT Distinct 1 AS ѡ��"
            Else
                mstrSql = mstrSql & "SELECT Distinct 1 AS ѡ��"
            End If

            mstrSql = mstrSql & ",A.���ID AS ID, " & _
                              "C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)') As ����," & _
                              "C.�����," & _
                              "C.סԺ��," & _
                              "D.���� AS �������," & _
                              "A.����ҽ�� AS ������," & _
                              "f.������,f.����ʱ��, " & _
                              "'Item' AS ͼ��,Decode(Nvl(z.ҽ�����,0), 0, '��', Decode(Sum(Nvl(z.����,0)), 0, '��', '��')) As �շ�,NVL(A.������־,0) AS ����,H.��������,Decode(I.��Ŀ���,2,2,1) As ��Ŀ���,F.������,F.����ʱ�� As ����ʱ�� " & _
                         "FROM ����ҽ����¼ A," & _
                         "������Ϣ C,���ű� D,����ҽ������ F,���鱨����Ŀ G,������ĿĿ¼ H,������Ŀ I,����������Ŀ Y,����걾��¼ j,סԺ���ü�¼ Z " & _
                        "WHERE A.������� = 'C' " & _
                              "AND A.����ID=C.����ID " & _
                              "AND A.��������ID=D.ID " & _
                              "AND A.���id IS NOT NULL " & _
                              "AND A.ҽ��״̬=8 AND A.ID=F.ҽ��id " & _
                              "AND A.������Ŀid=G.������Ŀid AND G.ϸ��ID Is Null " & _
                              IIf(blMachineFind, "AND G.������Ŀid=Y.��Ŀid ", "AND G.������Ŀid=Y.��Ŀid(+) ") & _
                              "AND G.������ĿID=I.������ĿID " & _
                              "AND A.������ĿID=H.ID " & _
                              "And a.������Դ = 2 " & _
                              IIf(mbln���۵�ģʽ = True, " and a.������Դ = 2 and c.��Ժʱ�� is null ", " ")

            mstrSql = mstrSql & " " & _
                              IIf(mlngDefaultDevice = 0 And lngCurrDevice = 0, "", "AND (Y.����ID+0=[7] Or Y.����ID Is Null)") & _
                              "AND A.����ID=[1] AND A.ִ�п���id+0=[3] and a.���id = j.ҽ��id(+) and j.id = [8] " & _
                              "AND F.NO=Z.NO(+) AND F.��¼����=mod(Z.��¼����(+),10) AND F.ҽ��id=Z.ҽ�����(+)+0 " & _
                              IIf(mblnBarCode, " And F.��������=[2]   ", "") & _
                              IIf(int��ҳID = 0, " ", " And a.��ҳID = [9] ") & _
                              IIf(blnNoRange, " ", " AND A.����ʱ�� BETWEEN [5] and [6] ") & _
                              IIf(mblnPrice = True, " And z.No = [10] ", " ") & _
                              " GROUP BY A.���ID,a.id,C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)'),C.�����,C.סԺ��,D.����,A.����ҽ��,'Item',NVL(A.������־,0),H.��������,f.������,f.����ʱ��,Decode(I.��Ŀ���,2,2,1),F.������,F.����ʱ�� ,z.ҽ����� "
        End If
    End If
    mstrSql = "Select distinct A.*,TO_CHAR(B.����ʱ��,'YY-MM-DD HH24:MI') AS ����ʱ��," & _
        "B.ҽ������ AS ������Ŀ,B.��������ID,B.����ҽ��,nvl(b.��ҳID,0) as ��ҳID,a.�շ�, " & _
        " B.������Դ,C.���� as ���˿��� " & _
        " From (" & mstrSql & ") A,����ҽ����¼ B,���ű� C " & _
        " Where A.ID=B.ID And B.���˿���ID = C.id  "

    strSQLbak = mstrSql
    strSQLbak = Replace(strSQLbak, "סԺ���ü�¼", "������ü�¼")
    strSQLbak = Replace(strSQLbak, " and a.������Դ = 2 and c.��Ժʱ�� is null ", " And a.������Դ <> 2  and nvl(����״̬,0) <> 1 ")
    strSQLbak = Replace(strSQLbak, " And a.������Դ = 2 ", " And a.������Դ <> 2 and nvl(����״̬,0) <> 1 ")
    mstrSql = mstrSql & " Union ALL " & strSQLbak

    mstrSql = mstrSql & " order by ��ҳID Desc,����ʱ�� "

    rs.CursorLocation = adUseClient
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, _
        Me.Caption, mlng����ID, strText, ItemDeptID, "", _
            CDate(Format(strStart, "yyyy-MM-dd hh:mm:ss")), CDate(Format(strEnd, "yyyy-MM-dd hh:mm:ss")), _
            IIf(mlngDefaultDevice > 0, mlngDefaultDevice, lngCurrDevice), mlngSampleID, int��ҳID, mstrNO)
    If rs.RecordCount > 0 Then
        If Nvl(rs("������Դ"), 0) = 2 Then
            strSQL = "select nvl(max(��ҳID),0) as ��ҳID,�������� from ������ҳ where ����id = [1]  group by ��ҳid ,�������� order by  ��ҳid  desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
            intPatProperty = Val(rsTmp("��������"))
        End If
    End If
    If intPatProperty = 1 Then
        If blnδ�շ���ʾ = False Then
            rs.MoveFirst
            Do Until rs.EOF
                If Nvl(rs("������Դ"), 0) = 2 And Trim(Nvl(rs("�շ�"))) = "��" Then
                    mstrSql = Replace(mstrSql, "סԺ���ü�¼", "������ü�¼")
                    Set rsPatProperty = zlDatabase.OpenSQLRecord(mstrSql, _
                    Me.Caption, mlng����ID, strText, ItemDeptID, "", _
                        CDate(Format(strStart, "yyyy-MM-dd hh:mm:ss")), CDate(Format(strEnd, "yyyy-MM-dd hh:mm:ss")), _
                        IIf(mlngDefaultDevice > 0, mlngDefaultDevice, lngCurrDevice), mlngSampleID, int��ҳID, mstrNO)
                    If rsPatProperty.RecordCount > 0 Then
                        Set rs = rsPatProperty
                        Exit Do
                    End If
                End If
                rs.MoveNext
            Loop
        End If
    End If
    If rs.BOF Then
        If mblnBarCode = True Then
            'ɨ�������Ŀ�����������ʾ
            gstrSql = "select ִ��״̬,a.����ID,a.ִ�п���ID,c.���� as ִ�п������� from ����ҽ����¼ a,����ҽ������ b,���ű� c " & vbNewLine & _
                      " where a.id = b.ҽ��id and a.���id is not null and a.ִ�п���ID = c.id and b.�������� = [1]   "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strText)
            If rsTmp.EOF = True Then
                MsgBox "û���ҵ�����<" & strText & ">!"
            Else
                strִ�п������� = rsTmp("ִ�п�������")
                lngִ�п���ID = rsTmp("ִ�п���ID")
                lng����ID = rsTmp("����ID")
                intִ��״̬ = rsTmp("ִ��״̬")
                gstrSql = "select ����ʱ��,������,���ʱ��,�����,�걾���,����ID from ����걾��¼ where ����id = [1] and ��������  = [2] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng����ID, strText)
                If rsTmp.EOF = False Then
                    If intִ��״̬ = 1 Then
                        MsgBox "����<" & strText & ">�ѱ�" & rsTmp("�����") & "��" & rsTmp("���ʱ��") & "���,�걾��" & _
                            TransSampleNO_PH(rsTmp("�걾���"), Nvl(rsTmp("����ID"), -1)) & "."
                    ElseIf intִ��״̬ = 2 Then
                        MsgBox "����<" & strText & ">�Ѿ���!"
                    ElseIf intִ��״̬ = 3 Then
                        MsgBox "����<" & strText & ">�ѱ�" & rsTmp("������") & "��" & rsTmp("����ʱ��") & "����,�걾��" & _
                            TransSampleNO_PH(rsTmp("�걾���"), Nvl(rsTmp("����ID"), -1)) & "."
                    End If
                Else

                    If intִ��״̬ = 1 Then
                        MsgBox "����<" & strText & ">�����!"
                    ElseIf intִ��״̬ = 2 Then
                        MsgBox "����<" & strText & ">�Ѿ���!"
                    ElseIf intִ��״̬ = 3 Then
                        MsgBox "����<" & strText & ">�Ѻ���!"
                    Else
                        If ItemDeptID <> lngִ�п���ID Then
                            MsgBox "����<" & strText & ">" & "��ִ�п�����<" & strִ�п������� & ">���ǵ�ǰѡ��Ŀ��Ҳ��ܺ��գ�"
                        End If
                    End If
                End If
            End If
        End If
        OpenSelect = 0
        Exit Function
    End If

    'סԺ����ֻ������סԺ��ҽ��
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        If Nvl(rs("������Դ"), 0) = 2 Then
            int��ҳID = Val(Nvl(rsTmp("��ҳID")))
        End If
    End If
    strTmp = ""
    Do Until rs.EOF
        If Nvl(rs("������Դ")) = 2 Then
            If int��ҳID = Val(Nvl(rs("��ҳID"))) Then
                strTmp = strTmp & " or ID=" & Val(Nvl(rs("ID")))
            End If
        Else
            strTmp = strTmp & " or ID=" & Val(Nvl(rs("ID")))
        End If
        rs.MoveNext
    Loop
    If strTmp <> "" Then
        rs.filter = "ID=-1" & strTmp
    End If

    If (rs.RecordCount = 1 And blnWhere) Or mblnBarCode Then GoTo Over
    
    If rs.RecordCount > 0 Then
        '����ʾδ�շѵı걾
        If blnδ�շ���ʾ = False Then
            strFilter = "�շ� <> '��'"
            rs.filter = strFilter
        End If
    End If
    
    Call ClientToScreen(txt����.hWnd, objPoint)
    If frmSelectMuli.ShowSelectSP(Me, rs, strLvw, objPoint.X * 15 - 30, objPoint.Y * 15 + txt����.Height - 30, _
        8000, 5600, Me.Name & "\�����ձ걾ѡ��", "����±��й�ѡ��һ�κ��յı걾") Then
        GoTo Over
    End If
    Exit Function

Over:
    If rs.EOF Then Exit Function

    '�Լ����־�����ж�

    rs.MoveFirst
    Do Until rs.EOF
        If rs("����") = 1 Then
            mbln���� = True
            Exit Do
        End If
        rs.MoveNext
    Loop

    '��û���շѵĲ��˽����ж�
    If blnδ�շ���ʾ = False Then
        rs.MoveFirst
        Do Until rs.EOF

            'סԺ
            If Nvl(rs("������Դ"), 0) = 2 And Trim(Nvl(rs("�շ�"))) = "��" Then
                MsgBox "������Ŀ<" & Nvl(rs("������Ŀ")) & ">��δ�շ���Ŀ���˷���Ŀ���ܺ���", vbInformation, "������ʾ"
                OpenSelect = 3
                Exit Function
            End If

            '����
            If Nvl(rs("������Դ"), 0) = 1 And Trim(Nvl(rs("�շ�"))) = "��" Then
                MsgBox "������Ŀ<" & Nvl(rs("������Ŀ")) & ">��δ�շ���Ŀ���˷���Ŀ���ܺ���", vbInformation, "������ʾ"
                OpenSelect = 3
                Exit Function
            End If

            '��첡��ֻ�ж��˷�
            If Nvl(rs("������Դ"), 0) = 4 And Trim(Nvl(rs("�շ�"))) = "��" Then
                MsgBox "������Ŀ<" & Nvl(rs("������Ŀ")) & ">��δ�շ���Ŀ���˷���Ŀ���ܺ���", vbInformation, "������ʾ"
                OpenSelect = 3
                Exit Function
            End If
            rs.MoveNext
        Loop
    End If

    '�ж��Ƿ񳬹��ͼ�ʱ����
    rs.MoveFirst
    Do Until rs.EOF
        If (IsDate(Nvl(rs("����ʱ��"))) = True And Nvl(rs("������")) <> "") Then
            gstrSql = "Select Min(�ͼ�ʱ��) As �ͼ�ʱ��" & vbNewLine & _
                        "From ����ҽ����¼ A, ������Ŀѡ�� B" & vbNewLine & _
                        "Where A.������Ŀid = B.������Ŀid And A.���id = [1] And Nvl(�ͼ�ʱ��, 0) > 0"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(rs("ID")))
            If rsTmp.EOF = False Then
                If Val(Nvl(rsTmp("�ͼ�ʱ��"))) > 0 Then
                    strDate = zlDatabase.Currentdate
                    If DateDiff("n", Nvl(rs("����ʱ��")), strDate) > Val(Nvl(rsTmp("�ͼ�ʱ��"))) And Val(Nvl(rsTmp("�ͼ�ʱ��"))) > 0 Then
                        strTmp = DateDiff("n", Nvl(rs("����ʱ��")), strDate)
                        If MsgBox("������Ŀ�������ͼ�ʱ�ޣ�" & strTmp & "����),�Ƿ�����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            OpenSelect = 3
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        rs.MoveNext
    Loop

    '��Ŀ��ͬ��������ʾ
    rs.MoveFirst
    strTmp = ""
    Do While Not rs.EOF
        If Val(rs("ѡ��")) = 1 Then
            If InStr("," & strTmp & ",", "," & Trim(rs("������Ŀ")) & ",") > 0 Then
                If MsgBox("ѡ���˶����ͬ����Ŀ���Ƿ�������գ� ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    OpenSelect = 3
                    Exit Function
                Else
                    Exit Do
                End If
            Else
                strTmp = strTmp & "," & Trim(rs("������Ŀ"))
            End If
        End If
        rs.MoveNext
    Loop

    '������Դ��һ��ʱ����һ�����
    rs.MoveFirst
    strTmp = ""
    Do Until rs.EOF
        If strTmp = "" Then strTmp = Trim(Nvl(rs("������Դ")))
        If Trim(strTmp) <> Trim(Nvl(rs("������Դ"))) Then
            MsgBox "��ͬ����Դ����ҽ����һ�����", vbInformation, "������ʾ"
            OpenSelect = 3
            Exit Function
        End If
        rs.MoveNext
    Loop

    'ĸ�׵�ҽ������Ů��ҽ��������һ�����
    rs.MoveFirst
    strTmp = ""
    Do Until rs.EOF
        If strTmp = "" Then strTmp = Nvl(rs("����"))
        If Trim(strTmp) <> Trim(Nvl(rs("����"))) Then
            MsgBox "ĸ�׵�ҽ�����ܺ���Ů��ҽ����һ����գ�", vbInformation, "������ʾ"
            OpenSelect = 3
            Exit Function
        End If
        rs.MoveNext
    Loop

    '��ͬ�ı걾������һ�����
'    rs.MoveFirst
'    strTmp = ""
'    Do Until rs.EOF
'        If strTmp = "" Then strTmp = Nvl(rs("�걾����"))
'        If Trim(strTmp) <> Trim(Nvl(rs("�걾����"))) Then
'            MsgBox "��ͬ�ı걾���Ͳ�����һ����գ�", vbInformation, "������ʾ"
'            OpenSelect = 3
'            Exit Function
'        End If
'        rs.MoveNext
'    Loop

    rs.MoveFirst
    If InStr(Nvl(rs("����")), "(Ӥ��)") > 0 Then
        gstrSql = "Select B.����id, B.��ҳid, B.���, B.Ӥ������, B.Ӥ���Ա�" & vbNewLine & _
                    "From ����ҽ����¼ A, ������������¼ B" & vbNewLine & _
                    "Where A.����id = B.����id And A.��ҳid = B.��ҳid And A.Ӥ�� = B.��� And A.���id = [1] And Rownum = 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Nvl(rs("ID"), 0)))
        If rsTmp.EOF = False Then
            Me.txt���� = ""
            Me.cboAge = "Ӥ��"
            txt����.Text = Nvl(rsTmp("Ӥ������"))
            strTmp = "����='" & CStr(Nvl(rsTmp("Ӥ���Ա�"))) & "'"
            mRsSex.filter = strTmp
            If mRsSex.EOF = False Then
                Me.cbo�Ա�.Text = mRsSex!���� & "-" & mRsSex!����
            End If
        End If
    Else
        On Error Resume Next
        gstrSql = "select  ���� from ����ҽ����¼ where ���id=[1] And Rownum = 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Nvl(rs("ID"), 0)))
        
        strAge = Nvl(rsTmp("����"))
        
        strAge = Replace(strAge, "Сʱ", "ʱ")
        strAge = Replace(strAge, "����", "��")
        
        If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "��", ""), "��", ""), "��", ""), "ʱ", ""), "��", "")) <> "" Then
            If InStr(strAge, "����") > 0 Or InStr(strAge, "Ӥ��") > 0 Then
                Me.txt����.Text = ""
                Me.cboAge.Text = Trim(strAge)
            Else
                strAge = Replace(Replace(Replace(Replace(Replace(strAge, "��", "��;"), "��", "��;"), "��", "��;"), "ʱ", "ʱ;"), "��", "��;")
                'strAge = Replace(strAge, "����", "Ӥ��")
                aAge = Split(strAge, ";")
                If UBound(aAge) = 1 Then
                    Me.txt����.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
                Else
                    Me.txt����.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
                    Me.txt����1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "��", "����"), "ʱ", "Сʱ")
                End If
            End If
        Else
            If Val(strAge) <> 0 Then
                Me.txt����.Text = Val(strAge)
            End If
            Me.cboAge.ListIndex = 0
        End If
    End If
    On Error GoTo ErrHand
    If InStr(txt����, "(Ӥ��)") > 0 Then Me.txt���� = ""

    DTP(0).Value = Format(zlCommFun.Nvl(rs("����ʱ��"), zlDatabase.Currentdate), "YYYY-MM-DD HH:MM:SS")
    mbln΢������Ŀ = (zlCommFun.Nvl(rs("��Ŀ���"), 1) = 2)
    If Nvl(rs("������"), "") = "" Then
        cbo(2).Visible = False
        DTP(2).Visible = False
        lbl(1).Visible = False
    Else
        cbo(2).Visible = True
        DTP(2).Visible = True
        lbl(1).Visible = True
        cbo(2).Text = zlCommFun.Nvl(rs("������"))
        'zlControl.CboLocate cbo(2), zlCommFun.Nvl(rs("������"))
        DTP(2).Value = Format(zlCommFun.Nvl(rs("����ʱ��"), zlDatabase.Currentdate), "YYYY-MM-DD HH:MM:SS")
    End If
    txtPatientDept = Nvl(rs("���˿���"))
    Me.cbo��������.ListIndex = FindComboItem(Me.cbo��������, Nvl(rs("��������ID")))
    cbo(1).Text = zlCommFun.Nvl(rs("������"))
    'û�в�����ʱ����ʾ����ʱ��
    If Trim(Me.cbo(1).Text) = "" Then
        Me.cbo(1).Visible = False
        Me.DTP(0).Visible = False
        lbl(0).Visible = False
    Else
        Me.cbo(1).Visible = True
        Me.DTP(0).Visible = True
        lbl(0).Visible = True
    End If
    Me.txtPatientDept.Text = Nvl(rs("���˿���"))
    Select Case Nvl(Nvl(rs("������Դ")))
        Case 1, 3, 4
            txtID.Text = Nvl(rs("�����"))
            txtBed.Text = ""
        Case 2
            txtID.Text = Nvl(rs("סԺ��"))
    End Select
'    zlControl.CboLocate cbo(1), zlCommFun.Nvl(rs("������"))
    On Error Resume Next
    strField = ""
    strField = rs.Fields("����ҽ��").Name
    If strField = "����ҽ��" Then
        Me.cboҽ��.Text = Nvl(rs("����ҽ��"))
        For i = 0 To Me.cboҽ��.ListCount - 1
            If Me.cboҽ��.List(i) Like Nvl(rs("����ҽ��")) Then
                Me.cboҽ��.ListIndex = i
                Exit For
            End If
        Next
    End If

    gstrSql = "select �걾��λ AS �걾���� from ����ҽ����¼ where ���id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(rs("ID")))
    If rsTmp.EOF = False Then
        Me.txt����.Text = Nvl(rsTmp("�걾����"))
    End If

    On Error GoTo ErrHand

    If rs.RecordCount = 1 Then
        mstrKeys = zlCommFun.Nvl(rs("ID").Value)

        Me.txtҽ������ = rs("������Ŀ")
        Me.txtҽ������.Tag = rs("������Ŀ")

        If blnCheck Then
'            Me.lblCash.Font.Strikethrough = Not (rs("�շ�") = "��")
            Me.lblCash.Caption = IIf((rs("�շ�") = "��"), "��", "")
        Else
'            Me.lblCash.Font.Strikethrough = True
            Me.lblCash.Caption = ""
        End If
    Else
        strLastKeys = mstrKeys: mstrKeys = "": Me.txtҽ������ = ""
        Do While Not rs.EOF
            mstrKeys = mstrKeys & "," & zlCommFun.Nvl(rs("ID").Value)
'            If InStr("," & txtҽ������ & ",", "," & zlCommFun.Nvl(rs("������Ŀ").Value & ",")) <= 0 Then
                txtҽ������ = txtҽ������ & "," & zlCommFun.Nvl(rs("������Ŀ").Value)
'            End If

            If blnCheck Then
'                Me.lblCash.Font.Strikethrough = Not (rs("�շ�") = "��")
                Me.lblCash.Caption = IIf((rs("�շ�") = "��"), "��", "")
            Else
'                Me.lblCash.Font.Strikethrough = True
                Me.lblCash.Caption = ""
            End If

            rs.MoveNext
        Loop
        If mstrKeys = "" Then
            mstrKeys = strLastKeys
        Else
            mstrKeys = Mid(mstrKeys, 2)
            txtҽ������ = Mid(txtҽ������, 2)
            txtҽ������.Tag = Mid(txtҽ������, 2)
        End If
    End If

    OpenSelect = 1

    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CalcNextCode(ByVal lngKey As Long, ByVal intRow As Integer, ByVal iType As Integer) As String
    '--------------------------------------------------------------------------------------------------------
    '����:����ָ�������ڵ����ڵ���һ��ȱʡ�걾��
    '����:lngKey                ��������ID
    '     iType                 �걾���0=��ͨ��1=����
    '����:ȱʡ�걾����
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strToday As String
    Dim strTmp As String
    Dim lng���� As Long
    Dim strLabNo As String, strLabQCNo As String '����걾���ʿر걾
    Dim mstrSql As String, mlngLoop As Long
    Dim strStartDate As String
    Dim strEndDate As String
    Dim lngDefaultItemID As Long
    Dim strItem As String
    Dim rsTmp As New ADODB.Recordset

    'ʱ��,����,�걾��
    On Error GoTo ErrHand

    strToday = Format(DTP(1).Value, "YYYY-MM-DD")
    strStartDate = GetDateTime(mMakeNoRule, 1, DTP(1).Value)
    strEndDate = GetDateTime(mMakeNoRule, 2, DTP(1).Value)

'    lngDefaultItemID = mlngDefaultItemID

    If mintItemRule = 1 Then
        If mstrKeys <> "" Then
            gstrSql = "select /*+ rule */ ������ĿID from ����ҽ����¼ where ���ID in  (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKeys)
            If rsTmp.EOF = False Then
                lngDefaultItemID = rsTmp("������ĿID")
            End If
        Else
            If mstrExtData <> "" Then
                lngDefaultItemID = Val(mstrExtData)
            End If
        End If

    End If

    On Error GoTo point1

    mstrSql = "SELECT NVL(MAX(TO_NUMBER(�걾���)),0) AS ������ FROM ����걾��¼ a,����������Ŀ b " & _
                "WHERE ����ʱ�� BETWEEN [2] and [3] And a.id = b.�걾id(+) And nvl(a.�Ƿ��ʿ�Ʒ,0) = 0 " & _
                    IIf(lngKey = -1, " AND ����id IS NULL " & _
                        IIf(lngDefaultItemID > 0, " And b.������Ŀid = [4] ", ""), "AND ����id= [1] ") & " And ҽ��ID Is Not Null" & _
                    IIf(mblnEmerge, IIf(iType = 1, " And �걾���=1", " And Nvl(�걾���,0)<>1"), "")
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, lngKey, CDate(strStartDate), _
                           CDate(strEndDate), lngDefaultItemID)

    If Not rs.EOF Then strLabNo = zlCommFun.Nvl(rs("������"))

    On Error GoTo ErrHand
    GoTo point2

point1:
    On Error GoTo ErrHand

    mstrSql = "SELECT NVL(MAX(�걾���),'') AS ������ FROM ����걾��¼ a,����������Ŀ b " & _
                "WHERE ����ʱ�� BETWEEN [2] and [3] And a.id = b.�걾id(+) And nvl(a.�Ƿ��ʿ�Ʒ,0) = 0  " & _
                    IIf(lngKey = -1, " AND ����id IS NULL " & _
                    IIf(lngDefaultItemID > 0, " And b.������Ŀid = [4] ", ""), "AND ����id= [1] ") & " And ҽ��ID Is Not Null" & _
                    IIf(mblnEmerge, IIf(iType = 1, " And �걾���=1", " And Nvl(�걾���,0)<>1"), "")
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, lngKey, CDate(strStartDate), _
                            CDate(strEndDate), lngDefaultItemID)

    If Not rs.EOF Then strLabNo = zlCommFun.Nvl(rs("������"))

point2:
    On Error GoTo point3

    mstrSql = "SELECT NVL(MAX(TO_NUMBER(�걾���)),0) AS ������ FROM ����걾��¼ a,����������Ŀ b " & _
                "WHERE ����ʱ�� BETWEEN [2] and [3] And a.id = b.�걾ID(+) And nvl(a.�Ƿ��ʿ�Ʒ,0) = 0 " & _
                    IIf(lngKey = -1, " AND ����id IS NULL " & _
                    IIf(lngDefaultItemID > 0, " And b.������Ŀid = [4] ", ""), "AND ����id= [1] ") & _
                    IIf(mblnEmerge, IIf(iType = 1, " And �걾���=1", " And Nvl(�걾���,0)<>1"), "")
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, lngKey, CDate(strStartDate), _
                            CDate(strEndDate), lngDefaultItemID)

    If Not rs.EOF Then strLabQCNo = zlCommFun.Nvl(rs("������"))

    On Error GoTo ErrHand
    GoTo point4

point3:
    On Error GoTo ErrHand

    mstrSql = "SELECT NVL(MAX(�걾���),'') AS ������ FROM ����걾��¼ a,����������Ŀ b " & _
                "WHERE ����ʱ�� BETWEEN [2] and [3] And a.id = b.�걾ID(+) And nvl(a.�Ƿ��ʿ�Ʒ,0) =  0 " & _
                    IIf(lngKey = -1, " AND ����id IS NULL " & _
                    IIf(lngDefaultItemID > 0, " And b.������Ŀid = [4] ", ""), "AND ����id=[1] ") & _
                    IIf(mblnEmerge, IIf(iType = 1, " And �걾���=1", " And Nvl(�걾���,0)<>1"), "")
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, lngKey, CDate(strStartDate), _
                            CDate(strEndDate), lngDefaultItemID)

    If Not rs.EOF Then strLabQCNo = zlCommFun.Nvl(rs("������"))

point4:
    If Val(strLabNo) >= Val(strLabQCNo) Then
        CalcNextCode = strLabNo
    Else
        CalcNextCode = strLabQCNo
    End If
'    If Val(strLabQCNo) > Val(strLabNo) + 100 Then CalcNextCode = strLabNo

    For mlngLoop = 1 To vsf2.Rows - 1
        If mlngLoop <> intRow Then
            If Val(vsf2.RowData(mlngLoop)) = lngKey Then
                If Val(CalcNextCode) < Val(vsf2.TextMatrix(mlngLoop, 2)) Then
                    CalcNextCode = Val(vsf2.TextMatrix(mlngLoop, 2))
                End If
            End If
        End If
    Next

    If Val(CalcNextCode) <= 0 Then
        CalcNextCode = "1"
        Exit Function
    End If

    CalcNextCode = Val(CalcNextCode) + 1
    Exit Function

ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub LoadDefaultData()
    '--------------------------------------------------------------------------------------------------------
    '����:�������������ͱ걾��
    '--------------------------------------------------------------------------------------------------------
    Dim lngloop As Long
    Dim strNO As String
    Dim lngDefaultRec As Long
    Dim strConnectDevIDs As String, lngTmpNO As Long
    Dim strSubQry As String, mstrSql As String, mRs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim lngCurrDevice As Long '��Ŀ��ȱʡ������������δѡ��ǰ����ʱ��Ч
    Dim blnCurrDevice As Boolean '�Ƿ�ѡ����Ĭ������
    Dim intNow As Integer
    Dim lngItem As Long


    '��ȡ�������ӵļ�������
    strConnectDevIDs = GetConnectDevs

    On Error GoTo ErrHand
    '��ȡ��Ӧ�ļ��������б�

    If mstrKeys <> "" Then
        If mbln΢������Ŀ = False Then
            mstrSql = "SELECT ID,����,ȱʡ����,MIN(���ID) As ҽ��ID FROM " & _
                        "(SELECT DISTINCT NVL(E.ID,-1) AS ID,NVL(E.����,'[�ֹ�]') AS ����,NVL(D.ȱʡ����,-1) AS ȱʡ����,A.���ID " & _
                            "FROM ����ҽ����¼ A, ���鱨����Ŀ B, ����������Ŀ D, �������� E " & _
                            "Where A.������ĿID+0 = B.������ĿID(+) " & _
                            "AND B.������ĿID = D.��Ŀid(+) AND D.����id = E.ID(+) " & _
                            "AND A.����ID=[1] AND Instr(','||[2]||',',','||A.���ID||',')>0 " & _
                            "ORDER BY NVL(E.ID,-1)  DESC) " & _
                        "GROUP BY ȱʡ����,ID,����"
        Else
            mstrSql = "SELECT ID,����,ȱʡ����,MIN(���ID) As ҽ��ID FROM " & _
                        "(SELECT DISTINCT NVL(E.ID," & mlngDefaultDevice & ") AS ID, " & _
                        "NVL(E.����," & " (select ���� from �������� where id = [3]) " & " ) AS ����, " & _
                        "NVL(1,-1) AS ȱʡ����,A.���ID " & _
                            "FROM ����ҽ����¼ A, ���鱨����Ŀ B, ����ϸ������ D, �������� E " & _
                            "Where A.������ĿID+0 = B.������ĿID(+) " & _
                            "AND B.ϸ��ID = D.ϸ��id(+) AND D.����id = E.ID(+) " & _
                            "AND A.����ID=[1] AND Instr(','||[2]||',',','||A.���ID||',')>0 " & _
                            "ORDER BY NVL(E.ID," & mlngDefaultDevice & ")  DESC) " & _
                        "GROUP BY ȱʡ����,ID,����"
        End If
        Set mRs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, mlng����ID, mstrKeys, mlngDefaultDevice, mstrMachines)
    Else
        strSubQry = ""
        rsRelativeAdvice.MoveFirst
        Do While Not rsRelativeAdvice.EOF
            strSubQry = strSubQry & " Union All " & "Select " & rsRelativeAdvice("ID") & " As ID From Dual"

            rsRelativeAdvice.MoveNext
        Loop
        If Len(strSubQry) > 0 Then strSubQry = Mid(strSubQry, 12)
        rsRelativeAdvice.MoveFirst

        gstrSql = "Select A.������Ŀid" & vbNewLine & _
                    "From ���鱨����Ŀ A, ������Ŀ B, (" & strSubQry & ") S" & vbNewLine & _
                    "Where S.ID = A.������Ŀid And A.������Ŀid = B.������Ŀid And B.��Ŀ��� = 2"

        zlDatabase.OpenRecordset rsTmp, gstrSql, Me.Caption
        mbln΢������Ŀ = Not rsTmp.EOF

        If mbln΢������Ŀ = False Then
            gstrSql = "Select A.������Ŀid" & vbNewLine & _
                    "From ���鱨����Ŀ A, ������Ŀ B, (" & strSubQry & ") S" & vbNewLine & _
                    "Where S.ID = A.������Ŀid And A.������Ŀid = B.������Ŀid And B.��Ŀ��� <> 2 "
            zlDatabase.OpenRecordset rsTmp, gstrSql, Me.Caption
            If rsTmp.EOF = False Then
                mstrSql = "SELECT DISTINCT NVL(E.ID,-1) AS ID,NVL(E.����,'[�ֹ�]') AS ����,NVL(D.ȱʡ����,-1) AS ȱʡ���� " & _
                                "FROM ���鱨����Ŀ B, ����������Ŀ D, �������� E,(" & strSubQry & ") S " & _
                                "Where S.ID=B.������ĿID(+) " & _
                                "AND B.������ĿID = D.��Ŀid(+) AND D.����id = E.ID(+) " & _
                                "ORDER BY NVL(D.ȱʡ����,-1)  DESC"
            Else
                mstrSql = "Select id,����,1 as ȱʡ���� from �������� where id = [1] "
            End If
        Else
            mstrSql = "SELECT DISTINCT NVL(E.ID," & mlngDefaultDevice & ") AS ID, " & _
                            "NVL(E.����," & " (select ���� from �������� where id = [1]) " & ") AS ����,-1 AS ȱʡ���� " & _
                            "FROM ���鱨����Ŀ B, ����ϸ������ D, �������� E,(" & strSubQry & ") S " & _
                            "Where S.ID=B.������ĿID(+) " & _
                            "AND B.ϸ��ID = D.ϸ��id(+) AND D.����id = E.ID(+)  "
        End If

        Set mRs = zlDatabase.OpenSQLRecord(mstrSql, gstrSysName, mlngDefaultDevice)
        'û������ʱ�̶�д�͵�ǰ�����ϵ�����ID
        If mRs.RecordCount <= 1 And mlngDefaultDevice > 0 Then
            If mRs("id") = -1 Then
                mstrSql = "select id , ���� , 0 as ȱʡ���� from �������� a where id = [1] "
                Set mRs = zlDatabase.OpenSQLRecord(mstrSql, gstrSysName, mlngDefaultDevice)
            End If
        End If
    End If
    '���������δѡ��ǰ������������Ŀ��Ĭ������Ϊ׼
    If mlngDefaultDevice = 0 Then
        If mstrKeys <> "" Then
            mstrSql = "Select B.סԺ����ID,B.��������ID,B.סԺ�����ֽ�,B.���������ֽ� " & _
                "From ����ҽ����¼ A,������Ŀѡ�� B Where A.������ĿID=B.������ĿID " & _
                "AND A.����ID=[1] AND Instr(','||[2]||',',','||A.���ID||',')>0"
            Set rsTmp = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, mlng����ID, mstrKeys)
        Else
            mstrSql = "SELECT B.סԺ����ID,B.��������ID,B.סԺ�����ֽ�,B.���������ֽ� " & _
                "FROM ������Ŀѡ�� B,(" & strSubQry & ") S " & _
                "Where S.ID=B.������ĿID"
            Set rsTmp = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption)
        End If
        If rsTmp.EOF Then
            lngCurrDevice = 0
        Else
            lngCurrDevice = IIf(PatientType = 2, Nvl(rsTmp("סԺ����ID"), 0), Nvl(rsTmp("��������ID"), 0))
        End If

    End If

    If mRs.BOF = False Then
        ResetVsf vsf2
        '���һ���������б걾��������N���ռ�¼
        vsf2.Rows = mRs.RecordCount + 1

        For lngloop = 1 To vsf2.Rows - 1
            If Val(vsf2.RowData(lngloop)) = 0 Then

                '���������Ƿ��Ѿ�ʹ��,����ʹ��,��ȡһ������,��û����һ��,��ȡ���һ��
                lngDefaultRec = -1: mRs.MoveFirst
                blnCurrDevice = False
                Do While Not mRs.EOF
                    If CheckHave(zlCommFun.Nvl(mRs("ID"), 0)) = False Then
                        If zlCommFun.Nvl(mRs("ID"), 0) = mlngDefaultDevice Then
                            'ȡ��������ָ���ļ�������
                            lngDefaultRec = mRs.AbsolutePosition
                            Exit Do '���ټ�������
                        Else
                            If zlCommFun.Nvl(mRs("ID"), 0) = lngCurrDevice Then
                                lngDefaultRec = mRs.AbsolutePosition
                                blnCurrDevice = True
                            Else
                                If InStr(";" & strConnectDevIDs & ";", ";" & zlCommFun.Nvl(mRs("ID"), 0) & ";") > 0 Then
                                    'Ĭ��ȡ�������ӵļ�������
                                    If Not blnCurrDevice Then lngDefaultRec = mRs.AbsolutePosition
                                Else
                                    If lngDefaultRec = -1 Then lngDefaultRec = mRs.AbsolutePosition '�Ƚ���ǰ����ѡ��
                                End If
                            End If
                        End If
                    End If
                    mRs.MoveNext
                Loop
                If lngDefaultRec = -1 Then
                    mRs.MoveLast
                Else
                    mRs.AbsolutePosition = lngDefaultRec
                End If

                If mblnBarCode = False And mstr�������� <> "" Then
                    If InStr("," & mstr�������� & ",", "," & mRs("ID") & ",") > 0 Then
                        MsgBox "����<" & mRs("����") & ">����ʹ����������!", vbInformation, Me.Caption
                        Exit Sub
                    End If
                End If

                vsf2.TextMatrix(lngloop, 1) = zlCommFun.Nvl(mRs("����"))
                vsf2.RowData(lngloop) = zlCommFun.Nvl(mRs("ID"), 0)

                If mstrKeys <> "" Then
                    vsf2.TextMatrix(lngloop, 4) = zlCommFun.Nvl(mRs("ҽ��ID"), 0)
                    vsf2.TextMatrix(lngloop, 5) = IIf(mbln���� And mblnEmerge = True, "-1", "0")
                End If

                intNow = Val(zlDatabase.GetPara("���ϴ�����ı걾���ۼ�", 100, 1208, 0))

                If intNow = 1 And mstrNONumber <> "" Then
                    vsf2.TextMatrix(lngloop, 2) = TransSampleNO_PH(Val(mstrNONumber) + 1, vsf2.RowData(lngloop))
                Else
                    'ȡ�걾��
                    If vsf2.TextMatrix(lngloop, 5) = "-1" Then
                        '����
                        vsf2.TextMatrix(lngloop, 2) = TransSampleNO_PH(Val(CalcNextCode(Val(vsf2.RowData(lngloop)), lngloop, 1)), vsf2.RowData(lngloop))
                    Else
                        vsf2.TextMatrix(lngloop, 2) = TransSampleNO_PH(Val(CalcNextCode(Val(vsf2.RowData(lngloop)), lngloop, 0)), vsf2.RowData(lngloop))
                    End If
                End If
            End If
        Next
        vsf2.EditMode(1) = 1
        vsf2.ComboList(1) = "..."
    End If
    mbln���� = False
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SelectDefault()
    '--------------------------------------------------------------------------------------------------------
    '����:��ָ����뵽������
    '--------------------------------------------------------------------------------------------------------
    Dim iRow As Integer, iCurrRow As Integer
    Dim blnChkItem As Boolean '�Ƿ��м���ָ��
    Dim lngItemRow As Long
    Dim aItems() As Variant
    Dim astrKey() As String
    Dim intLoop As Integer
    Dim blnCheck As Boolean
    Dim lngloop As Long


    aItems = ReadData
    If UBound(aItems) = -1 Then Exit Sub
    If vsf2.Rows >= 2 And vsf2.RowData(1) = "" Then Exit Sub
    iCurrRow = vsf2.Row
    For iRow = 1 To vsf2.Rows - 1
        vsf2.Row = iRow
        blnChkItem = SelectValidItem(aItems, vsf2.RowData(iRow))

        '����ñ걾û��ָ�꣬��ɾ����΢����ĵ�һ���������⣩
        If Not blnChkItem And vsf2.Rows > 2 And Not (mbln΢������Ŀ And iRow = 1) And mintEditMode <= 1 Then
            vsf2.RemoveItem iRow
            iRow = iRow - 1
        End If
        If iRow = vsf2.Rows - 1 Then Exit For
    Next iRow

    astrKey = Split(mstrKeys, ",")
    For intLoop = 0 To UBound(astrKey)
        blnCheck = False
        For iRow = 1 To vsf2.Rows - 1
            If InStr(vsf2.TextMatrix(iRow, 3), Chr(1) & astrKey(intLoop) & Chr(1)) > 0 Then
                blnCheck = True
                Exit For
            End If
        Next
        If blnCheck = False Then
            For lngloop = 0 To UBound(aItems, 2)
                If aItems(1, lngloop) = astrKey(intLoop) Then
                    vsf2.TextMatrix(1, 3) = vsf2.TextMatrix(1, 3) & "|" & aItems(0, lngloop) & Chr(1) & aItems(1, lngloop) & _
                         Chr(1) & aItems(2, lngloop) & Chr(1) & aItems(3, lngloop) & Chr(1) & aItems(4, lngloop) & Chr(1) & aItems(5, lngloop) & _
                         Chr(1) & aItems(6, lngloop)
                    Exit For
                End If
            Next

        End If
    Next

'    For iRow = 1 To vsf2.Rows - 1
'        If Not mbln΢������Ŀ And Trim(vsf2.TextMatrix(iRow, 3)) = "" Then
'            vsf2.RemoveItem iRow
'            iRow = iRow - 1
'        End If
'        If iRow = vsf2.Rows - 1 Then Exit For
'    Next iRow

    vsf2.Row = iCurrRow
End Sub

Private Function ReadData() As Variant()
    '--------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ��Ŀ����ָ�굽����
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset, rsMicro As New ADODB.Recordset
    Dim strField As String, i As Long
    Dim strSubQry As String, mstrSql As String
    Dim blnOnlyMachine As Boolean   'ֻ���յ�ǰĬ�ϵ�������Ŀ
    Dim strWhere As String
    Dim strItems As String

    ReadData = Array()
    On Error GoTo ErrHand
    If mstrKeys = "" And rsRelativeAdvice Is Nothing Then
        Exit Function
    End If

    blnOnlyMachine = zlDatabase.GetPara("ֻ���յ�ǰ������Ŀ", 100, 1208, 0)

    If blnOnlyMachine = True And mlngDefaultDevice > 0 Then
        strWhere = " And B.������Ŀid = E.��Ŀid and  E.����ID = [3] "
    Else
        strWhere = " And B.������Ŀid = E.��Ŀid(+) "
    End If

    '��ȡ����ļ�����Ŀ(����ָ��)

    If mstrKeys <> "" Then
        If Not mbln΢������Ŀ Then
            If GetApplicationFormShowType = True Then
                strWhere = " and m.�����Ŀ <> 1 "
            End If
'            mstrSQL = "select ID,���ID,���,��־,����ο�,������ĿID,RowNum as �������,����,ѡ�� From " & _
                "(SELECT ID,���id,���,��־,����ο�,������ĿID,�������,����," & _
                "' '||����1||' '||����2||' '||����3||' '||����4||' '||����5||' '||����6||' '||����7||' '||����8||' '||����9||' ' AS ����,0 As ѡ�� " & _
                "FROM " & _
                "   ( " & _
                "    Select a.ID,a.���id,a.���,a.��־,a.����ο�,a.������Ŀid,�������,����, " & _
                "          Max(decode(mod(rownum,9),0,a.����id,'')) as ����1, " & _
                "          Max(decode(mod(rownum,9),1,a.����id,'')) as ����2, " & _
                "          Max(decode(mod(rownum,9),2,a.����id,'')) as ����3, " & _
                "          Max(decode(mod(rownum,9),3,a.����id,'')) as ����4, " & _
                "          Max(decode(mod(rownum,9),4,a.����id,'')) as ����5, " & _
                "          Max(decode(mod(rownum,9),5,a.����id,'')) as ����6, " & _
                "          Max(decode(mod(rownum,9),6,a.����id,'')) as ����7, " & _
                "          Max(decode(mod(rownum,9),7,a.����id,'')) as ����8, " & _
                "          Max(decode(mod(rownum,9),8,a.����id,'')) as ����9 "

'             mstrSQL = mstrSQL & "    From (" & _
                "           SELECT C.ID,A.���id,Decode(D.�������,3,Nvl(D.Ĭ��ֵ,'-'),2,D.Ĭ��ֵ,'') As ���,'' As ��־," & _
                "           Trim(REPLACE(REPLACE(' '||zlGetReference(C.ID,A.�걾��λ,DECODE(F.�Ա�,'��',1,'Ů',2,0),F.��������," & mlngDefaultDevice & "),' .','0.'),'��.','��0.')) AS ����ο�,B.������ĿID,B.�������," & _
                "           e.����ID,decode(D.�������,NULL,M.����,D.�������) as ���� " & _
                "           FROM ����ҽ����¼ A,���鱨����Ŀ B,����������Ŀ C,������Ŀ D,����������Ŀ E,������Ϣ F,������ĿĿ¼ M " & _
                "           WHERE A.���id>0 " & _
                "               AND A.������ĿID+0=B.������ĿID AND B.ϸ��ID Is Null " & _
                "               AND B.������ĿID=C.ID AND A.����ID=F.����ID And B.������ĿID = M.ID " & _
                "               AND D.������ĿID=C.ID AND B.������ĿID=E.��ĿID(+) AND A.����ID=[1] AND Instr(','||[2]||',',','||A.���ID||',')>0 Order by C.ID ) a " & _
                "   Group by  a.ID,a.���id,a.���,a.��־,a.����ο�,a.������Ŀid,�������,���� ) " & _
                "Order by ����,�������) "
            mstrSql = "select ID, ���id, ���, ��־, ����ο�, ������Ŀid, rownum as  �������, ����id, 0 As ѡ��" & vbNewLine & _
            "from" & vbNewLine & _
            "(Select ID, ���id, ���, ��־, ����ο�, ������Ŀid,  �������, ����id, 0 As ѡ��" & vbNewLine & _
            "From (Select C.ID, A.���id, Nvl(D.Ĭ��ֵ, '') As ���, '' As ��־," & vbNewLine & _
            "              Trim(Replace(Replace(' ' || Zlgetreference(C.ID, A.�걾��λ, Decode(F.�Ա�, '��', 1, 'Ů', 2, 0), F.��������), ' .'," & vbNewLine & _
            "                                    '0.'), '��.', '��0.')) As ����ο�, B.������Ŀid, B.�������, E.����id," & vbNewLine & _
            "              lpad(Decode(D.�������, Null, M.����, D.�������),10,'0') As ���� " & vbNewLine & _
            "       From ����ҽ����¼ A, ���鱨����Ŀ B, ����������Ŀ C, ������Ŀ D, ����������Ŀ E, ������Ϣ F, ������ĿĿ¼ M" & vbNewLine & _
            "       Where A.���id > 0 And A.������Ŀid + 0 = B.������Ŀid And B.ϸ��id Is Null And B.������Ŀid = C.ID And A.����id = F.����id And" & vbNewLine & _
            "             B.������Ŀid = M.ID And D.������Ŀid = C.ID  And A.����id = [1] And" & vbNewLine & _
            "             Instr(',' ||[2]|| ',', ',' || A.���id || ',') > 0" & vbNewLine & _
            "             " & strWhere & vbNewLine & _
            "       Order By C.ID)" & vbNewLine & _
            "order by ����,�������)"


        Else
'            mstrSQL = "SELECT ID,���id,���,��־,����ο�,������ĿID,rownum as �������, " & _
                "' '||����1||' '||����2||' '||����3||' '||����4||' '||����5||' '||����6||' '||����7||' '||����8||' '||����9||' ' AS ����,0 As ѡ�� " & _
                "FROM " & _
                "(SELECT D.ID,A.���id,'' As ���,'' As ��־,'' AS ����ο�,B.������ĿID,B.�������," & _
                "          Max(decode(mod(rownum,9),0,e.����id,'')) as ����1, " & _
                "          Max(decode(mod(rownum,9),1,e.����id,'')) as ����2, " & _
                "          Max(decode(mod(rownum,9),2,e.����id,'')) as ����3, " & _
                "          Max(decode(mod(rownum,9),3,e.����id,'')) as ����4, " & _
                "          Max(decode(mod(rownum,9),4,e.����id,'')) as ����5, " & _
                "          Max(decode(mod(rownum,9),5,e.����id,'')) as ����6, " & _
                "          Max(decode(mod(rownum,9),6,e.����id,'')) as ����7, " & _
                "          Max(decode(mod(rownum,9),7,e.����id,'')) as ����8, " & _
                "          Max(decode(mod(rownum,9),8,e.����id,'')) as ����9 " & _
                " FROM ����ҽ����¼ A,���鱨����Ŀ B,����ϸ�� D,����ϸ������ E,������Ϣ F " & _
                " WHERE A.���id>0 " & _
                    "AND A.������ĿID+0=B.������ĿID " & _
                    "AND B.ϸ��ID=D.ID AND A.����ID=F.����ID " & _
                    "AND B.ϸ��ID=E.ϸ��ID(+) AND A.����ID=[1] AND Instr(','||[2]||',',','||A.���ID||',')>0" & _
                " GROUP BY D.ID,A.���id,'','','',B.������ĿID,B.������� Order By B.������ĿID,B.������� Desc)"
            mstrSql = "Select ID, ���id, ���, ��־, ����ο�, ������Ŀid, Rownum As �������, ����id, 0 As ѡ��" & vbNewLine & _
                "From (Select D.ID, A.���id, D.Ĭ�Ͻ�� As ���, '' As ��־, '' As ����ο�, B.������Ŀid, B.�������, ����id" & vbNewLine & _
                "       From ����ҽ����¼ A, ���鱨����Ŀ B, ����ϸ�� D, ����ϸ������ E, ������Ϣ F" & vbNewLine & _
                "       Where A.���id > 0 And A.������Ŀid + 0 = B.������Ŀid And B.ϸ��id = D.ID And A.����id = F.����id And B.ϸ��id = E.ϸ��id(+) And" & vbNewLine & _
                "             A.����id = [1] And Instr(',' ||[2]|| ',', ',' || A.���id || ',') > 0" & vbNewLine & _
                "       Order By B.������Ŀid, B.������� )"

        End If
        Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, mlng����ID, mstrKeys, mlngDefaultDevice, strItems)
        If rs.BOF = False Then
            vsf2.Tag = rs.RecordCount

            ReadData = rs.GetRows
        End If
    Else
        If Not rsRelativeAdvice Is Nothing Then
            strSubQry = ""
            rsRelativeAdvice.MoveFirst
            Do While Not rsRelativeAdvice.EOF
                strSubQry = strSubQry & " Union All " & "Select " & rsRelativeAdvice("ID") & " As ID From Dual"
                If strItems = "" Then
                    strItems = Val(rsRelativeAdvice("ID") & "")
                Else
                    strItems = strItems & "," & Val(rsRelativeAdvice("ID") & "")
                End If
                rsRelativeAdvice.MoveNext
            Loop
            If Len(strSubQry) > 0 Then strSubQry = Mid(strSubQry, 12)
            rsRelativeAdvice.MoveFirst

            '��ȡ����ļ�����Ŀ(����ָ��)
            If Not mbln΢������Ŀ Then
'                mstrSQL = "select ID,���ID,���,��־,����ο�,������ĿID,RowNum as �������,����,ѡ�� From " & _
                    "(SELECT ID,���ID,���,��־,����ο�,������ĿID,�������,����," & _
                    "' '||����1||' '||����2||' '||����3||' '||����4||' '||����5||' '||����6||' '||����7||' '||����8||' '||����9||' ' AS ����,0 As ѡ�� " & _
                    "FROM " & _
                    "  (" & _
                    "   Select a.id , a.���id ,a.���,a.��־ , a.����ο�,a.������Ŀid,a.�������,����," & _
                    "          Max(decode(mod(rownum,9),0,a.����id,'')) as ����1, " & _
                "          Max(decode(mod(rownum,9),1,a.����id,'')) as ����2, " & _
                "          Max(decode(mod(rownum,9),2,a.����id,'')) as ����3, " & _
                "          Max(decode(mod(rownum,9),3,a.����id,'')) as ����4, " & _
                "          Max(decode(mod(rownum,9),4,a.����id,'')) as ����5, " & _
                "          Max(decode(mod(rownum,9),5,a.����id,'')) as ����6, " & _
                "          Max(decode(mod(rownum,9),6,a.����id,'')) as ����7, " & _
                "          Max(decode(mod(rownum,9),7,a.����id,'')) as ����8, " & _
                "          Max(decode(mod(rownum,9),8,a.����id,'')) as ����9 " & _
                    "   From ( " & _
                    "       SELECT C.ID,0 As ���ID,Decode(D.�������,3,Nvl(D.Ĭ��ֵ,'-'),2,D.Ĭ��ֵ,'') As ���,'' As ��־," & _
                    "       Trim(REPLACE(REPLACE(' '||zlGetReference(C.ID,'" & txt����.Text & "'," & Decode(cbo�Ա�.Text, "��", 1, "Ů", 2, 0) & ",NULL," & mlngDefaultDevice & "),' .','0.'),'��.','��0.')) AS ����ο�,B.������ĿID,B.�������," & _
                    "       e.����id,decode(D.�������,NULL,M.����,D.�������) as ����  " & _
                    "       FROM ���鱨����Ŀ B,����������Ŀ C,������Ŀ D,����������Ŀ E,(" & strSubQry & ") S , ������ĿĿ¼ M " & _
                    "       WHERE B.������ĿID=S.ID AND B.ϸ��ID Is Null " & _
                    "           AND B.������ĿID=C.ID And B.������ĿId = M.ID " & _
                    "           AND D.������ĿID=C.ID AND B.������ĿID=E.��ĿID(+) order by c.id  ) a " & _
                    "   Group by a.id , a.���id ,a.���,a.��־ , a.����ο�,a.������Ŀid,a.�������,a.����) " & _
                    "Order by ����,�������) "
                mstrSql = "Select ID, ���id, ���, ��־, ����ο�, ������Ŀid, Rownum As �������, ����id, 0 As ѡ��" & vbNewLine & _
                    "From (Select ID, ���id, ���, ��־, ����ο�, ������Ŀid, �������, ����id, 0 As ѡ��" & vbNewLine & _
                    "       From (Select C.ID, 0 As ���id, Nvl(D.Ĭ��ֵ, '') As ���, '' As ��־," & vbNewLine & _
                    "                     Trim(Replace(Replace(' ' || Zlgetreference(C.ID, '����Ѫ', 0, Null), ' .', '0.'), '��.', '��0.')) As ����ο�," & vbNewLine & _
                    "                     B.������Ŀid, B.�������, E.����id, lpad(Decode(D.�������, Null, M.����, D.�������),10,'0') As ���� " & vbNewLine & _
                    "              From ���鱨����Ŀ B, ����������Ŀ C, ������Ŀ D, ����������Ŀ E , ������ĿĿ¼ M" & vbNewLine & _
                    "              Where B.������Ŀid in (Select * From Table(Cast(f_Num2list([4]) As zlTools.t_Numlist))) And B.ϸ��id Is Null And B.������Ŀid = C.ID And B.������Ŀid = M.ID And D.������Ŀid = C.ID " & vbNewLine & _
                    "              " & strWhere & vbNewLine & _
                    "              Order By C.ID)" & vbNewLine & _
                    "       Order By ����, �������)"

            Else
'                mstrSQL = "SELECT ID,���ID,���,��־,����ο�,������ĿID,rownum as �������," & _
                    "' '||����1||' '||����2||' '||����3||' '||����4||' '||����5||' '||����6||' '||����7||' '||����8||' '||����9||' ' AS ����,0 As ѡ�� " & _
                    "FROM " & _
                    "(SELECT D.ID,0 As ���ID,'' As ���,'' As ��־,'' AS ����ο�,B.������ĿID,B.�������," & _
                    "          Max(decode(mod(rownum,9),0,e.����id,'')) as ����1, " & _
                "          Max(decode(mod(rownum,9),1,e.����id,'')) as ����2, " & _
                "          Max(decode(mod(rownum,9),2,e.����id,'')) as ����3, " & _
                "          Max(decode(mod(rownum,9),3,e.����id,'')) as ����4, " & _
                "          Max(decode(mod(rownum,9),4,e.����id,'')) as ����5, " & _
                "          Max(decode(mod(rownum,9),5,e.����id,'')) as ����6, " & _
                "          Max(decode(mod(rownum,9),6,e.����id,'')) as ����7, " & _
                "          Max(decode(mod(rownum,9),7,e.����id,'')) as ����8, " & _
                "          Max(decode(mod(rownum,9),8,e.����id,'')) as ����9 " & _
                    " FROM ���鱨����Ŀ B,����ϸ�� D,����ϸ������ E,(" & strSubQry & ") S " & _
                    " WHERE B.������ĿID=S.ID " & _
                        "AND B.ϸ��ID=D.ID " & _
                        "AND B.ϸ��ID=E.ϸ��ID(+)" & _
                    " GROUP BY D.ID,0,'','','',B.������ĿID,B.������� Order By B.������ĿID,B.������� Desc)"
                mstrSql = "Select ID, ���id, ���, ��־, ����ο�, ������Ŀid, Rownum As �������, ����id, 0 As ѡ��" & vbNewLine & _
                "From (Select D.ID, 0 As ���id, D.Ĭ�Ͻ�� As ���, '' As ��־, '' As ����ο�, B.������Ŀid, B.�������, E.����id" & vbNewLine & _
                "       From ���鱨����Ŀ B, ����ϸ�� D, ����ϸ������ E " & vbNewLine & _
                "       Where B.������Ŀid in (Select * From Table(Cast(f_Num2list([4]) As zlTools.t_Numlist))) And B.ϸ��id = D.ID And B.ϸ��id = E.ϸ��id(+)" & vbNewLine & _
                "       Order By B.������Ŀid, B.������� )"


            End If
            Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, mlng����ID, mstrKeys, mlngDefaultDevice, strItems)
'            Call OpenRecord(rs, mstrSQL, Me.Caption)
            If rs.BOF = False Then
                vsf2.Tag = rs.RecordCount

                ReadData = rs.GetRows
            End If
        End If
    End If

    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SelectValidItem(aItems() As Variant, ByVal lngDeviceID As Long) As Boolean
'����ǰ����������ָ��ȫ���ӵ��б���
    Dim mlngLoop As Long, lngItemRow  As Long

    SelectValidItem = False
    vsf2.TextMatrix(vsf2.Row, 3) = ""
'    For mlngLoop = UBound(aItems, 2) To 0 Step -1              '��֪����ǰΪʲôҪ������ѭ��
    For mlngLoop = 0 To UBound(aItems, 2)
        If Val(aItems(ItemCol.ID, mlngLoop)) > 0 And Val(aItems(ItemCol.ѡ��, mlngLoop)) = 0 Then
            If (InStr(IIf(Trim(aItems(ItemCol.����, mlngLoop)) = "", "-1", aItems(ItemCol.����, mlngLoop)), lngDeviceID) > 0) Or _
                (lngDeviceID = -1 And mlngDefaultDevice = -1) Or Trim(Nvl(aItems(ItemCol.����, mlngLoop))) = "" Then
                '��дҽ������Ҫ�������б걾�����
                If vsf2.TextMatrix(vsf2.Row, 4) = "" Or Val(vsf2.TextMatrix(vsf2.Row, 4)) = 0 Then
                    vsf2.TextMatrix(vsf2.Row, 4) = aItems(ItemCol.���ID, mlngLoop)
                End If
                aItems(ItemCol.ѡ��, mlngLoop) = 1
                If InStr("|" & vsf2.TextMatrix(vsf2.Row, 3), "|" & aItems(0, mlngLoop) & Chr(1)) = 0 Then
                    SelectValidItem = True

                    vsf2.TextMatrix(vsf2.Row, 3) = vsf2.TextMatrix(vsf2.Row, 3) & "|" & aItems(0, mlngLoop) & Chr(1) & aItems(1, mlngLoop) & _
                         Chr(1) & aItems(2, mlngLoop) & Chr(1) & aItems(3, mlngLoop) & Chr(1) & aItems(4, mlngLoop) & Chr(1) & aItems(5, mlngLoop) & _
                         Chr(1) & aItems(6, mlngLoop)
                    For lngItemRow = 0 To UBound(aItems, 2)
                        If Val(aItems(ItemCol.ID, lngItemRow)) = Val(aItems(ItemCol.ID, mlngLoop)) Then
                            aItems(ItemCol.ѡ��, lngItemRow) = 1
                        End If
                    Next
                End If
            End If
        End If
    Next mlngLoop
    If Len(vsf2.TextMatrix(vsf2.Row, 3)) > 0 Then vsf2.TextMatrix(vsf2.Row, 3) = Mid(vsf2.TextMatrix(vsf2.Row, 3), 2)
End Function

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Private Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date

    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(strNO, "0000000")
        Exit Function
    End If
    GetFullNO = strNO

    strSQL = "Select ��Ź���,Sysdate as ���� From ������Ʊ� Where ��Ŀ���=" & intNum
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!��Ź���, 0)
        curDate = rsTmp!����
    End If

    If intType = 1 Then
        '���ձ��
        strSQL = Format(CDate("1992-" & Format(rsTmp!����, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '������
        GetFullNO = PreFixNO & Format(strNO, "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����:�����Ƿ��Ѿ�ʹ�ù����Ƿ��в���Ȩ��
    '����:
    '����:
    '--------------------------------------------------------------------------------------------------------
    Dim mlngLoop As Long

    '�ж��Ƿ���Ȩ�޲���
    If InStr(mstrMachines, ";" & lngKey & ";") = 0 Then
        CheckHave = True
        Exit Function
    End If

    '�ж��Ƿ���ʹ��
    For mlngLoop = 1 To vsf2.Rows - 1
        If vsf2.RowData(mlngLoop) = lngKey Then
            CheckHave = True
            Exit Function
        End If
    Next

End Function

Private Function ValidData() As Boolean
    '--------------------------------------------------------------------------------------------------------
    '���ܣ�
    '--------------------------------------------------------------------------------------------------------
    Dim varTmp As Variant
    Dim strTmp As String
    Dim strError As String, mstrSql As String
    Dim lngloop As Long
    Dim lngCount As Long
    Dim rs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim i As Integer, iType As Integer
    Dim strStartDate As String
    Dim strEndDate As String
    Dim strҽ��ID As String

    On Error GoTo errH

    Call txt����_LostFocus

    ValidData = False

    If Me.txt����.Text <> Me.txt����.Tag Then
        Me.txt����.Enabled = False
'        Call txt����_Validate(False)
        Call Txt����Exec
        If Me.txt����.Text <> "" Then
            SetFocusNextIndex txt����.TabIndex
        Else
            Me.txt����.Enabled = True
            Me.txt����.SetFocus
        End If
        gintSelectFocus = 2
        Me.txt����.Enabled = True
    End If



    If Len(Trim(Me.txt����)) = 0 Then
'        mintFocusItem = FocusItem.����
        If Me.txt����.Enabled = True Then
            Me.txt����.SetFocus
        End If
        Exit Function
    End If

    If Len(Trim(cbo��������.Text)) = 0 Then
        If Me.cbo��������.Enabled = True Then
            cbo��������.SetFocus
        End If
        Exit Function
    End If

    '������ʱ�䲻�ܴ��ں���ʱ��
    If mintEditMode <> 3 Then
        If Trim(Me.cbo(1).Text) <> "" Then
            If CDate(Me.DTP(0).Value) > CDate(Me.DTP(1).Value) Then
                MsgBox "����ʱ����ں���ʱ�䣬�������ʱ�䣡", vbInformation, Me.Caption
                If Me.DTP(1).Enabled = True And Me.DTP(1).Visible = True Then
                    Me.DTP(1).SetFocus
                End If
                Exit Function
            End If
        End If
    End If

    If mstrKeys = "" And rsRelativeAdvice Is Nothing And mblnCheckIn = False Then
'        mintFocusItem = FocusItem.ҽ������
        On Error Resume Next
        Me.txtҽ������.SetFocus
        Exit Function
    End If

    '1.���ÿһ���걾ָ���ļ��������Ƿ���ȷ
    For i = 1 To vsf2.Rows - 1
        If Trim(vsf2.TextMatrix(i, 2)) = "" Then
            MsgBox "��" & i & "���걾û�б걾�ţ�", vbInformation, gstrSysName: gintSelectFocus = 2 'DoEvents
            vsf2.Row = i
            vsf2.Col = 2
            vsf2.SetFocus
            vsf2.ShowCell vsf2.Row, vsf2.Col
            Exit Function
        End If
    Next
    If vsf2.TextMatrix(vsf2.Row, vsf2.Col) <> vsf2.EditText And vsf2.EditText <> "" Then
        vsf2.TextMatrix(vsf2.Row, vsf2.Col) = vsf2.EditText
    End If
    ReDim mlngNoneHomeKey(vsf2.Rows - 1)
    ReDim mlngSourceKey(vsf2.Rows - 1)
    For i = 1 To vsf2.Rows - 1
        iType = IIf(vsf2.TextMatrix(i, 5) = "-1", 1, 0)
        '����Ƿ���Ч
        If Val(vsf2.RowData(i)) > 0 Then
            mstrSql = "SELECT ID,�걾���,Nvl(�Ƿ��ʿ�Ʒ,0) as �Ƿ��ʿ�Ʒ,���� FROM ����걾��¼ WHERE   ����id= [1] " & _
                " AND ����ʱ�� Between [2] AND [3] AND �걾���=[4]" & _
                IIf(mblnEmerge, IIf(iType = 1, " And �걾���=1", " And Nvl(�걾���,0)<>1"), "")
        Else
            mstrSql = "SELECT ID,�걾���,Nvl(�Ƿ��ʿ�Ʒ,0) as �Ƿ��ʿ�Ʒ,���� FROM ����걾��¼ WHERE    ����id Is Null " & _
                " AND ����ʱ�� Between [2] AND [3] AND �걾���=[4]" & _
                IIf(mblnEmerge, IIf(iType = 1, " And �걾���=1", " And Nvl(�걾���,0)<>1"), "")
        End If

        strStartDate = GetDateTime(mMakeNoRule, 1, DTP(1).Value)
        strEndDate = GetDateTime(mMakeNoRule, 2, DTP(1).Value)

        Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, Val(vsf2.RowData(i)), _
            CDate(Format(strStartDate, "yyyy-MM-dd 00:00:00")), _
            CDate(Format(strEndDate, "yyyy-MM-dd 23:59:59")), TransSampleNO(Trim(vsf2.TextMatrix(i, 2))))

        If rs.BOF = False Then
            If rs("�Ƿ��ʿ�Ʒ") = 1 Then
                '���ʿ�Ʒ���Ѻ��յı걾���ܱ�����
                MsgBox "�����õı걾�����ʿ�Ʒ���������趨�걾�ţ�", vbInformation, Me.Caption
                vsf2.Row = i
                vsf2.Col = 2
                vsf2.SetFocus
                vsf2.ShowCell vsf2.Row, vsf2.Col
                gintSelectFocus = 2
                Exit Function
            End If

            '���ա��Ǽǡ������Ƿ񸴸������걾
            If mintEditMode = 3 Then
                mlngNoneHomeKey(i) = mlngSampleID
                rs.filter = "ID<>" & mlngSampleID
                If rs.RecordCount > 0 Then
                    If Trim(Nvl(rs("����"))) <> "" Then
                        '���ʿ�Ʒ���Ѻ��յı걾���ܱ�����
                        MsgBox "�����õı걾���ѱ����գ��������趨�걾�ţ�", vbInformation, Me.Caption
                        vsf2.Row = i
                        vsf2.Col = 2
                        vsf2.SetFocus
                        vsf2.ShowCell vsf2.Row, vsf2.Col
                        gintSelectFocus = 2
                        Exit Function
                    End If
                    mlngSourceKey(i) = rs("ID")
                End If
            Else
                If Trim(Nvl(rs("����"))) <> "" Then
                    MsgBox "�����õı걾���ѱ����գ��������趨�걾�ţ�", vbInformation, Me.Caption
                    vsf2.Row = i
                    vsf2.Col = 2
                    vsf2.SetFocus
                    vsf2.ShowCell vsf2.Row, vsf2.Col
                    gintSelectFocus = 2
                    Exit Function
                End If
                If MsgBox("�����õı걾���Ѿ����ڣ�����Ҫ�񸴸�?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    ValidData = True
                    mlngNoneHomeKey(i) = rs("ID").Value
                    gintSelectFocus = 2
                Else
                    vsf2.Row = i
                    vsf2.Col = 2
                    vsf2.SetFocus
                    vsf2.ShowCell vsf2.Row, vsf2.Col
                    gintSelectFocus = 2
                    Exit Function
                End If
            End If
        Else
            '����д��һ���µı걾��
            mlngNoneHomeKey(i) = mlngSampleID
        End If
    Next

    '--------------------------------------------------------------------------------------------------------------------------------
    '��ִ��������Զ���˵ķ���ʱ���Բ��˷��ý��м��ʱ�����
    gstrSql = " select /*+ rule */ id from ����ҽ����¼ where id in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) " & _
              " Union All " & _
              " select /*+ rule */ id from ����ҽ����¼ where ���id in (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) "
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKeys)
    Do While Not rs.EOF
        strҽ��ID = strҽ��ID & "," & rs("id")
        rs.MoveNext
    Loop
    strҽ��ID = Mid(strҽ��ID, 2)
    If Chk���۷���(Me, strҽ��ID, 0) = False And Trim(strҽ��ID) <> "" Then
        Exit Function
    End If
    '----------------------------------------------------------------------------------------------------------------------------------
    ValidData = True

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckMuliQuest(ByVal lng����ID As Long, ByVal lng����id As Long, ByVal strNO As String, ByRef lngKey As Long, ByVal iType As Integer, ByRef blnOther As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������     iType                 �걾���0=��ͨ��1=����
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim strStartDate  As String
    Dim strEndDate As String

    On Error GoTo ErrHand

    If lng����id > 0 Then
        strSQL = "SELECT A.ID,B.����ID,C.���� FROM ����걾��¼ A,����ҽ����¼ B,������Ϣ C WHERE A.ҽ��id=B.id AND B.����ID=C.����ID" & _
        " AND A.����id=[2] AND A.����ʱ�� Between [3] And [4] AND A.�걾���= [5] " & _
        IIf(mblnEmerge, IIf(iType = 1, " And A.�걾���=1", " And Nvl(A.�걾���,0)<>1"), "")
    Else
        strSQL = "SELECT A.ID,B.����ID,C.���� FROM ����걾��¼ A,����ҽ����¼ B,������Ϣ C WHERE A.ҽ��id=B.id AND B.����ID=C.����ID" & _
        " AND A.����id IS NULL AND A.����ʱ�� Between [3] And [4] AND A.�걾���= [5] " & _
        IIf(mblnEmerge, IIf(iType = 1, " And A.�걾���=1", " And Nvl(A.�걾���,0)<>1"), "")
    End If

    strStartDate = GetDateTime(mMakeNoRule, 1, DTP(1).Value)
    strEndDate = GetDateTime(mMakeNoRule, 2, DTP(1).Value)

    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng����id, _
        CDate(Format(strStartDate, "yyyy-MM-dd 00:00:00")), _
        CDate(Format(strEndDate, "yyyy-MM-dd 23:59:59")), strNO)

    If rs.BOF = False Then
        If mintEditMode <= 1 Then
            Call MsgBox("�����õı걾���Ѿ����ڣ��������趨�걾��!", vbInformation, gstrSysName): gintSelectFocus = 2 'DoEvents
'            mintFocusItem = FocusItem.�걾��

            vsf2.Col = 2
            vsf2.SetFocus
            vsf2.ShowCell vsf2.Row, vsf2.Col
            Exit Function
        End If
        If Not IsNull(rs("����ID")) Then
            lngKey = zlCommFun.Nvl(rs("ID"), 0)
            blnOther = True
        End If
    End If

    CheckMuliQuest = True

    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Private Function SaveData(Optional ByVal intEditState As Integer) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '���ܣ�
    '--------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim varTmp As Variant
    Dim lngloop As Long
    Dim strSQL() As String
    Dim blnMuliQuest As Boolean
    Dim lngMuliQuestKey As Long
    Dim mlngKey As Long 'ҽ��ID
    Dim lngKey As Long '�걾ID
    Dim lngResultID As Long '���ID������΢����
    Dim lngResultLoop As Long
    Dim i As Integer, varAdviceIDs As Variant 'ָ���Ӧ������ҽ��ID
    Dim strItemRecords As String
    Dim AdviceIDs() As Long, SampleIDs() As Long
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim blnAutoPrint As Boolean
    Dim strTmpNO As String '�걾��
    Dim blnOther As Boolean '�Ƿ��������߱걾
    Dim strTmp As String, rsTmp As ADODB.Recordset
    Dim mlngLoop As Long, blnAuditing As Boolean
    Dim blnNewAdvice As Boolean         '�Ƿ��½�ҽ��
    Dim blnEmergency As Integer         '�Ƿ�ʹ�ü��� 0=��ʹ�ü��� 1=ʹ�ü���
    Dim strStartDate  As String
    Dim strEndDate As String
    Dim strItems As String              '������ĿID,���ʱʹ��","�ָ�
    Dim blnNewpatinet As Boolean        '�Ƿ������ɵ��²���
    ReDim strSQL(1 To 1)
    ReDim AdviceIDs(0)
    ReDim SampleIDs(0)


    blnAuditing = zlDatabase.GetPara("�����ֱ�����", 100, 1208, True, 0)
    blnEmergency = Val(zlDatabase.GetPara("����걾", 100, 1208, 0))

    On Error GoTo ErrHand


    '���������⣬�����ɲ�����Ϣ
    blnNewpatinet = CreatePatient
    If mlng����ID = 0 Then
        MsgBox "��������ʧ�ܣ�������", vbInformation, Me.Caption
        Exit Function
    End If

    '�Ǽǣ�����ҽ��
    If mstrKeys = "" And mblnSaveAdvice Then
        If Not ValidAdvice Then
            SaveData = False
            Exit Function
        End If
        mlngKey = SaveAdviceData(blnNewpatinet)
        If mlngKey = -1 Then Exit Function
        blnNewAdvice = True
    Else
        blnNewAdvice = False
    End If


    blnAutoPrint = zlDatabase.GetPara("��˴�ӡ", 100, 1208, True, 0)

    strStartDate = GetDateTime(mMakeNoRule, 1, DTP(1).Value)
    strEndDate = GetDateTime(mMakeNoRule, 2, DTP(1).Value)

    For mlngLoop = 1 To vsf2.Rows - 1


        '======================================================================================================================
        '���ɱ걾ID
        If mlngNoneHomeKey(mlngLoop) = 0 Then
            lngKey = zlDatabase.GetNextId("����걾��¼")
        Else
            lngKey = mlngNoneHomeKey(mlngLoop)
        End If
        '======================================================================================================================

        '======================================================================================================================
        '����Զ����ɵ��²��ˣ����û�һ����ʾ��
        gstrSql = "select distinct a.ID,a.����ID,a.����,c.����,a.�걾��� from ����걾��¼ a,������Ŀ�ֲ� b,�������� c where a.id = b.�걾id and a.����id = c.id and  a.id = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKey)


        If rsTmp.EOF = False Then
            If Val(Nvl(rsTmp("����ID"))) <> mlng����ID And Val(Nvl(rsTmp("����ID"))) <> 0 Then
                'ѭ����������걾��Ϣ������ʾ
                Do While Not rsTmp.EOF
                    strTmp = strTmp & "������" & rsTmp("����") & "    �걾��:" & rsTmp("�걾���") & vbCrLf
                    rsTmp.MoveNext
                Loop
                rsTmp.MoveFirst
                If MsgBox("�����µǼ����²��ˣ���ǰ�Ĳ���<" & Nvl(rsTmp("����")) & ">�ĺ�����Ŀ���Զ��ع�!" & vbCrLf & strTmp & _
                    "�Ƿ����?", vbYesNo + vbDefaultButton2) = vbNo Then
                    Exit Function
                End If
            End If
        End If
        '=======================================================================================================================

        '=======================================================================================================================
        '����ǰ�Ȼع���ǰҽ�������б걾������ʱ��ʾ�������Ѻ��յı걾����������������Ϊ�˷����û����մ�ʱ�ٽ��к��յĹ���
        If intEditState <> 4 Then   '����ʱ�����
            gstrSql = "Select Distinct ҽ��ID From (Select ҽ��ID From ������Ŀ�ֲ� Where �걾id = [1] " & _
                    "Union All Select ҽ��ID From ����걾��¼ Where ID = [1])"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKey)
    
            Do While Not rsTmp.EOF
                If Not IsNull(rsTmp(0)) Then
                    strSQL(ReDimArray(strSQL)) = "ZL_����걾��¼_תΪ����(" & rsTmp(0) & ")"
                    'zlDatabase.ExecuteProcedure "ZL_����걾��¼_תΪ����(" & rsTmp(0) & ")", gstrSysName
                End If
                rsTmp.MoveNext
            Loop
        End If
        '========================================================================================================================


        '===========================================================================================================================================================
        '���¼���걾��Ϣ�ͼ�����ͨ�����Ϣ
        If Val(vsf2.TextMatrix(mlngLoop, 4)) <> 0 Then
            mlngKey = Val(vsf2.TextMatrix(mlngLoop, 4))
            gstrSql = "Select �걾��λ From ����ҽ����¼ Where Id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
'            If rsTmp.EOF = False Then txt����.Text = rsTmp("�걾��λ") & ""
        Else
            If mlngKey <= 0 Then mlngKey = Val(vsf2.TextMatrix(mlngLoop, 4))  '���յ�Ĭ��ҽ��ID
        End If

        ReDim Preserve AdviceIDs(UBound(AdviceIDs) + 1)
        ReDim Preserve SampleIDs(UBound(SampleIDs) + 1)
        AdviceIDs(UBound(AdviceIDs)) = mlngKey
        SampleIDs(UBound(SampleIDs)) = lngKey
        If vsf2.Row = mlngLoop And vsf2.Col = 2 And Me.vsf2.EditText <> "" Then
            vsf2.TextMatrix(mlngLoop, 2) = Me.vsf2.EditText
        End If
        strTmpNO = TransSampleNO(vsf2.TextMatrix(mlngLoop, 2))
        mstrNONumber = strTmpNO
        
        strSQL(ReDimArray(strSQL)) = "ZL_����걾��¼_�걾����(" & lngKey & "," & _
                                                                mlngKey & ",'" & IIf(Trim(mstrKeys) = "", mlngKey, mstrKeys & "," & mlngKey) & "'," & _
                                                                mlngSourceKey(mlngLoop) & ",'" & _
                                                                strTmpNO & "'," & _
                                                                IIf(cbo(1).Text <> "", "TO_DATE('" & Format(DTP(0).Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'", "Null,'") & _
                                                                IIf(InStr(cbo(1).Text, "-") > 0, zlCommFun.GetNeedName(cbo(1).Text), cbo(1).Text) & "'," & _
                                                                IIf(Val(vsf2.RowData(mlngLoop)) = -1, 0, Val(vsf2.RowData(mlngLoop))) & "," & _
                                                                "TO_DATE('" & Format(DTP(1).Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & _
                                                                IIf(InStr(cbo(0).Text, "-") > 0, zlCommFun.GetNeedName(cbo(0).Text), cbo(0).Text) & "','" & _
                                                                UserInfo.���� & "'," & _
                                                                "TO_DATE('" & Format(DTP(1).Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')," & IIf(mbln΢������Ŀ, 1, "Null") & "," & _
                                                                IIf(mblnEmerge = True, IIf(vsf2.TextMatrix(mlngLoop, 5) = "-1" And mblnEmerge = True, 1, 0), 0) & ",NULL,'" & _
                                                                txt����.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & _
                                                                txt���� & Me.cboAge.Text & Me.txt����1 & "','" & mstrNO & "','" & _
                                                                txt����.Text & "'," & Me.cbo��������.ItemData(Me.cbo��������.ListIndex) & ",'" & Me.cboҽ��.Text & "'," & _
                                                                IIf(Trim(txtID.Text) = "", "NULL", IIf(IsNumeric(txtID.Text), txtID.Text, "NULL")) & "," & _
                                                                IIf(Trim(txtBed.Text) = "", "NULL,", "'" & Me.txtBed & "',") & _
                                                                IIf(Trim(txtPatientDept.Text) = "", "NULL,'", "'" & txtPatientDept.Text & "','") & _
                                                                Me.txtҽ������ & "'," & CInt(IIf(blnNewAdvice, 1, 0)) & _
                                                                "," & mlng����ID & "," & ItemDeptID & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        
        'ע�⣺�����΢���������Ŀ�������ʱ����д������ͨ�����¼
        If vsf2.TextMatrix(mlngLoop, 3) = "" And mbln΢������Ŀ = False And mintEditMode > 1 Then
            vsf2.TextMatrix(mlngLoop, 3) = GetSampleData(mlngSampleID)
        End If
        varTmp = Split(vsf2.TextMatrix(mlngLoop, 3), "|")
        strItemRecords = ""
        strItems = ""
        For lngloop = 0 To UBound(varTmp)
            If mstrKeys <> "" Then
                '������������
                varAdviceIDs = Split(Split(varTmp(lngloop), Chr(1))(ItemCol.���ID), ";")
                For i = 0 To UBound(varAdviceIDs)
                    mlngKey = Val(varAdviceIDs(i)) 'ָ���Ӧ��ҽ��ID
                    If mlngKey > 0 Then
                        strItemRecords = strItemRecords & "|" & mlngKey & "^" & Val(Split(varTmp(lngloop), Chr(1))(ItemCol.ID)) & "^" & _
                            Split(varTmp(lngloop), Chr(1))(ItemCol.���) & "^" & _
                            IIf(Len(Trim(Split(varTmp(lngloop), Chr(1))(ItemCol.��־))) = 0, 0, Decode(Right(Split(varTmp(lngloop), Chr(1))(ItemCol.��־), 2), "ƫ��", 3, "ƫ��", 2, "����", 4, 1)) & "^" & Split(varTmp(lngloop), Chr(1))(ItemCol.����ο�) & _
                            "^" & Split(varTmp(lngloop), Chr(1))(ItemCol.������ĿID) & "^" & Split(varTmp(lngloop), Chr(1))(ItemCol.�������)
                        '��¼������Ŀ
                        If InStr(strItems & ",", "," & Split(varTmp(lngloop), Chr(1))(ItemCol.������ĿID) & ",") <= 0 Then
                            strItems = strItems & "," & Split(varTmp(lngloop), Chr(1))(ItemCol.������ĿID)
                        End If
                    End If
                Next i
            Else
                If Val(Split(varTmp(lngloop), Chr(1))(ItemCol.���ID)) > 0 Then
                    mlngKey = Val(Split(varTmp(lngloop), Chr(1))(ItemCol.���ID))
                End If
                strItemRecords = strItemRecords & "|" & mlngKey & "^" & Val(Split(varTmp(lngloop), Chr(1))(ItemCol.ID)) & "^" & _
                    Split(varTmp(lngloop), Chr(1))(ItemCol.���) & "^" & _
                    IIf(Len(Trim(Split(varTmp(lngloop), Chr(1))(ItemCol.��־))) = 0, 0, Decode(Right(Split(varTmp(lngloop), Chr(1))(ItemCol.��־), 2), "ƫ��", 3, "ƫ��", 2, "����", 4, 1)) & "^" & Split(varTmp(lngloop), Chr(1))(ItemCol.����ο�) & _
                    "^" & Split(varTmp(lngloop), Chr(1))(ItemCol.������ĿID) & "^" & Split(varTmp(lngloop), Chr(1))(ItemCol.�������)
                '��¼������Ŀ
                If InStr(strItems & ",", "," & Split(varTmp(lngloop), Chr(1))(ItemCol.������ĿID) & ",") <= 0 Then
                    strItems = strItems & "," & Split(varTmp(lngloop), Chr(1))(ItemCol.������ĿID)
                End If
            End If
        Next lngloop
        If Len(strItemRecords) > 0 Then
            strItemRecords = Mid(strItemRecords, 2)
            strSQL(ReDimArray(strSQL)) = "Zl_������ͨ���_Write(" & lngKey & "," & _
                IIf(Val(vsf2.RowData(mlngLoop)) = -1, 0, Val(vsf2.RowData(mlngLoop))) & ",'" & _
                strItemRecords & "',0," & IIf(mbln΢������Ŀ, 1, 0) & ")"

            If mbln΢������Ŀ = False Then
                'ɾ����ǰ������Ŀ��û�еĿ���Ŀ
                strSQL(ReDimArray(strSQL)) = "Zl_������ͨ���_DeleteItem(" & lngKey & ",'" & Mid(strItems, 2) & "'," & IIf(mbln΢������Ŀ, 1, 0) & ")"
            End If
        Else
            '�޸Ĳο�ֵ�ͱ�־
            strSQL(ReDimArray(strSQL)) = "Zl_������ͨ���_Write(" & lngKey & "," & _
                IIf(Val(vsf2.RowData(mlngLoop)) = -1, 0, Val(vsf2.RowData(mlngLoop))) & ",'',0," & IIf(mbln΢������Ŀ, 1, 0) & ",'" & IIf(Trim(mstrKeys) = "", mlngKey, mstrKeys & "," & mlngKey) & "')"
        End If
        strSQL(ReDimArray(strSQL)) = "Zl_���¼�����_Cale(" & lngKey & ")"
        '===========================================================================================================================================================================

    Next

    If mlngSampleID = 0 Then mlngSampleID = lngKey
    gcnOracle.BeginTrans
    blnTran = True
    '����ִ��SQL
    For mlngLoop = 1 To UBound(strSQL)
        If strSQL(mlngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(mlngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    If mstrKeys <> "" Then
        ModifyApplyToLIS mstrKeys, 1
    End If
    '����ǩ��
    If Signature(lngKey, gstrDBUser, "����") = False Then
        Exit Function
    End If



    '���ò�����ϢΪ����д��
    SetPatientInfoWrite True

    '�������Զ����
    If blnAuditing And mintEditMode > 1 And Len(mstrAuditer) > 0 Then
        '���յǼǵĲ����������
        For mlngLoop = 1 To vsf2.Rows - 1
            '������˹����ж�
            If VerifyAuditingRule(lngKey) = 1 Then
                If MsgBox("�鵥�н��������ʾֵ!�Ƿ�����?", _
                    vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
            If mSendReport = 1 And mstr������ = "" Then
                '����
                gstrSql = "Zl_����걾��¼_���󱨸�(" & lngKey & ",1,'" & UserInfo.���� & "')"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
            Else
                Call zlDatabase.ExecuteProcedure("ZL_����걾��¼_�������(" & SampleIDs(mlngLoop) & ",'" & mstrAuditer & "','" & UserInfo.��� & _
                                                "','" & UserInfo.���� & "')", Me.Caption)
                If blnAutoPrint Then
                    If GetReportCode(AdviceIDs(mlngLoop), 0, strReportCode, strReportParaNo, bytReportParaMode) Then
                        Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & AdviceIDs(mlngLoop), "�걾ID=" & lngKey, "����ID=" & mlng����ID, 2)
                    End If
                End If
            End If
        Next
    End If
    SaveData = True

    mblnSaveAdvice = False

    '�����ϴν���Ƿ񳬱꣨�����ڲ�ͨ�����������Ƿ��飬����Ӱ�����ܣ�
    Call chkLastRual(lngKey)

    Exit Function
ErrHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If

End Function

Private Sub txtҽ������_GotFocus()
    Call zlControl.TxtSelAll(txtҽ������)
    Me.txtҽ������.IMEMode = 2
End Sub

Private Sub txtҽ������_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH

    If KeyAscii = vbKeyReturn Then
        mblnSaveAdvice = True
        KeyAscii = 0
        If txtҽ������.Text = txtҽ������.Tag Then
'            zlCommFun.PressKey vbKeyTab
            SetFocusNextIndex Me.txtҽ������.TabIndex
            gintSelectFocus = 2
            Exit Sub
        End If

        With txtҽ������
            Set rsTmp = SelectDiagItem()
        End With

        If rsTmp Is Nothing Then 'ȡ����������
            '�ָ�ԭֵ
            txtҽ������.Text = txtҽ������.Tag
            zlControl.TxtSelAll txtҽ������
            txtҽ������.SetFocus: gintSelectFocus = 2: Exit Sub
        End If
        '����Ŀ��¼��

        '����ѡ����Ŀ����ȱʡҽ����Ϣ
        If AdviceInput(rsTmp) Then
'            DoEvents

            gintSelectFocus = 2
            '��ʾ��ȱʡ���õ�ֵ
            txtҽ������.Tag = txtҽ������.Text
            txt����.Tag = txt����.Text

            '��������������걾�ţ��������б걾�Ĳ������غ˻������룩��ֻ���¸�ҽ��ID
            If mintEditMode <= 1 Then
                Call LoadDefaultData
                Call SelectDefault

                With vsf2
                    If .Rows > 1 Then
                        .Row = 1
                    End If
                    .Col = 2
                    .ShowCell vsf2.Row, vsf2.Col
                    .SetFocus
                End With
            Else
                '��ҽ��ID
                Call SelectDefault

                Me.cbo��������.SetFocus
            End If
        Else
'            DoEvents
            gintSelectFocus = 2
            '�ָ�ԭֵ
            txtҽ������.Text = txtҽ������.Tag
            txt����.Text = txt����.Tag
            zlControl.TxtSelAll txtҽ������

            txtҽ������.SetFocus: gintSelectFocus = 2: Exit Sub
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtҽ������_Validate(Cancel As Boolean)
    '�ָ���Ϊ�ĸı�
    If txtҽ������.Text <> txtҽ������.Tag Then
        txtҽ������.Text = txtҽ������.Tag
    End If
End Sub

Private Function SelectDiagItem() As ADODB.Recordset
'ѡ�������Ŀ
    Dim strSQL As String
    Dim objPoint As POINTAPI

    strSQL = "Select Distinct A.ID,A.����,A.����,nvl(A.���㵥λ,'��') As ���㵥λ,nvl(A.�걾��λ,' ') As �걾��λ," + _
        "Decode(A.���,'H',Decode(A.��������,'1','����ȼ�','������')," + _
        "'E',Decode(A.��������,'1','��������','2','��ҩ;��','3','��ҩ�巨',4,'��ҩ�÷�','����')," + _
        "'Z',Decode(A.��������,'1','����','2','סԺ','3','ת��','4','����','5','��Ժ','6','תԺ','����'),A.��������) As ��Ŀ����,A.��� As ���ID,A.ID As ������ĿID,nvl(ִ��Ƶ��,0) As ִ��Ƶ��ID,nvl(���㷽ʽ,0) As ���㷽ʽID,nvl(ִ�а���,0) As ִ�а���ID,nvl(�Ƽ�����,0) As �Ƽ�����ID,nvl(ִ�п���,0) As ִ�п���ID "
    strSQL = strSQL + "From ������ĿĿ¼ A,������Ŀ���� C,����ִ�п��� D Where A.ID=C.������ĿID And A.ID=D.������ĿID And A.���='C' And D.ִ�п���ID=" & ItemDeptID
    strSQL = strSQL + " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
        "And A.������� IN(" & PatientType & ",3,4)  And Nvl(A.�����Ա�,0) IN (" + _
        IIf(Me.cbo�Ա�.Text Like "*��*", "1,0)", "2,0)") + _
        " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
        " And (upper(A.����) Like '" + gstrMatch + UCase(txtҽ������) + "%' Or Upper(A.����) Like '" + txtҽ������ + "%' Or Upper(C.����) Like '" + UCase(txtҽ������) + "%')"

    Call ClientToScreen(txtҽ������.hWnd, objPoint)
    Set SelectDiagItem = zlDatabase.ShowSelect(Me, strSQL, 0, "ѡ��������Ŀ", True, Me.txtҽ������.Text, "", True, True, True, objPoint.X * 15, objPoint.Y * 15, Me.txtҽ������.Height, False, True)
End Function

Private Function AdviceInput(Optional rsInput As ADODB.Recordset = Nothing) As Boolean
'���ܣ����������������Ŀ(���������)����ȱʡ��ҽ������
'������rsInput=�����ѡ�񷵻صļ�¼��
'���أ�����¼���Ƿ���Ч
    Dim rsTmp As ADODB.Recordset
    Dim strHelpText As String
    Dim strSQL As String
    Dim strExtData As String
    Dim blnOk As Boolean
    Dim t_Pati As TYPE_PatiInfoEx

    On Error GoTo errH

    '��Ŀ�����������뼰����Ϸ��Լ��
    '---------------------------------------------------------------------------------------------------------------
    If Not rsInput Is Nothing Then txtҽ������.Text = rsInput!����    '��ʱ��ʾ

    '��Ҫ����������ݵ�һЩ��Ŀ
    '---------------------------------------------------------------------------------------------------------------
    '������Ŀѡ�����걾
    strHelpText = "������Ŀ"

    If Not rsInput Is Nothing Then
        strExtData = rsInput!������ĿID & ";" & rsInput!�걾��λ    '��������Ŀ
    Else
        If Trim(Me.txtҽ������.Text) = "" Then mstrExtData = ""
        strExtData = mstrExtData    '��������Ŀ
    End If

    With t_Pati
        .str�Ա� = NeedName(cbo�Ա�.Text)
    End With
    On Error Resume Next
    '�ӿڸ��죺bytUseType ��ǰû�������ڴ�Ϊ0
    blnOk = frmAdviceEditEx.ShowMe(Me, Me.vsf2.hWnd, t_Pati, 2, 4, 0, 1, PatientType, , , , 0, strExtData, , , , , True)
    On Error GoTo errH

    If Not blnOk Then Exit Function
    If strExtData = "" Or Mid(strExtData, 1, 1) = ";" Then Exit Function

    '��ȡ�ɼ���ʽ
    Set rsTmp = SelectCap(Split(Split(strExtData, ";")(0), ",")(0))
    If rsTmp Is Nothing Then
        MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName
        Exit Function
    End If
    mlngCapID = rsTmp("ID")

    strSQL = "Select C.��Ŀ��� From ������ĿĿ¼ A,���鱨����Ŀ B,������Ŀ C " & _
        "Where A.ID=B.������ĿID And B.������ĿID=C.������ĿID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Split(Split(strExtData, ";")(0), ",")(0))
'    If rsTmp.EOF Then
'        mbln΢������Ŀ = False
'    Else
'        mbln΢������Ŀ = IIf(Nvl(rsTmp("��Ŀ���"), 0) = 2, True, False)
'    End If

    mstrExtData = strExtData
    If Not rsInput Is Nothing Then Me.txt���� = Trim(rsInput("�걾��λ"))

    Call AdviceSet�������(3, mstrExtData)
    txtҽ������.Text = Get�����������(2, "")
    txtҽ������.Text = txtҽ������.Text & "(" & Split(mstrExtData, ";")(1) & ")"
    Me.txt���� = Split(mstrExtData, ";")(1)

    '����ҽ��
    On Error Resume Next
    If Me.cboҽ��.Text = "" Then Me.cboҽ��.ListIndex = 0

    AdviceInput = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitSampleInfo(ByVal lngSampleID As Long)
'���ܣ����ݱ걾ID����ʼ������Ŀ���������ҽ������Ϣ
'������rsInput=�����ѡ�񷵻صļ�¼��
'���أ�����¼���Ƿ���Ч
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer, strTmp As String

    On Error GoTo errH

    strSQL = "Select ҽ��ID,�������ID,������,���鱸ע,����ID,������Դ,�걾��̬,������,����ʱ��,������,����ʱ�� From ����걾��¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngSampleID)
    If rsTmp.EOF Then Exit Sub
    mlng����ID = Nvl(rsTmp("����ID"), 0): If mlng����ID > 0 Then Me.txt����.Tag = Me.txt����
'    If Not IsNull(rsTmp("ҽ��ID")) Then ' And Nvl(rsTmp("������Դ"), 3) <> 3 Then
'        mblnSaveAdvice = False: SetAdviceEnable False
'    End If

    On Error Resume Next
'    If IsNull(rsTmp("ҽ��ID")) And Not IsNull(rsTmp("�������ID")) Then
    If Not IsNull(rsTmp("�������ID")) Then
        Me.cbo��������.ListIndex = FindComboItem(Me.cbo��������, Nvl(rsTmp("�������ID"), 0))
        Me.cboҽ��.Text = Nvl(rsTmp("������"))
    End If

    Me.cbo(0).Text = Nvl(rsTmp("�걾��̬"))

    Me.cbo(1).Text = Nvl(rsTmp("������"))
    DTP(0).Value = Format(zlCommFun.Nvl(rsTmp("����ʱ��"), zlDatabase.Currentdate), "YYYY-MM-DD HH:MM:SS")
    If Nvl(rsTmp("������")) = "" Then
        Me.cbo(2).Visible = False
        Me.DTP(2).Visible = False
        lbl(1).Visible = False
    Else
        Me.cbo(2).Visible = True
        Me.DTP(2).Visible = True
        lbl(1).Visible = True
        Me.cbo(2).Text = Nvl(rsTmp("������"))
        Me.DTP(2).Value = Nvl(rsTmp("����ʱ��"))
    End If
    On Error GoTo errH

    If IsNull(rsTmp("ҽ��ID")) Then
        '�����걾
        strSQL = "Select ������ĿID From ����������Ŀ Where �걾ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngSampleID)
    Else
        '��ҽ����
'        strSQL = "Select Distinct b.������Ŀid From ������Ŀ�ֲ� a , ����ҽ����¼ b  " & _
'                 " Where b.���id = a.ҽ��id(+) And a.�걾id = [1] "
'        Set rsTmp =zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngSampleID)
        strSQL = "Select ������ĿID From ����ҽ����¼ Where ���ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(rsTmp("ҽ��ID")))
    End If
    If rsTmp.EOF Then Exit Sub

    i = 0: mstrExtData = ""

    Do While Not rsTmp.EOF
        i = i + 1
        mstrExtData = mstrExtData & "," & Nvl(rsTmp("������ĿID"), 0)
'        If i = 3 Then Exit Do '�����ʾ3����Ŀ

        rsTmp.MoveNext
    Loop


    If Len(mstrExtData) > 0 Then
        mstrExtData = Mid(mstrExtData, 2)
    Else
        Exit Sub
    End If

    If mlngSampleID > 0 Then
        strSQL = "select �걾���� from ����걾��¼ where id = [1] and �걾���� is not null "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngSampleID)
    End If
    If rsTmp.EOF = True Then
        strSQL = "Select �걾����,Sum(1) From (" & _
            "   Select A.ID,C.�걾����" & _
            "   From ������ĿĿ¼ A,������Ŀ�ο� C,���鱨����Ŀ D" & _
            "   Where A.ID=D.������ĿID And D.������ĿID=C.��ĿID" & _
            "   And A.ID In (" & mstrExtData & ")" & _
            " ) Group By �걾���� Order By Sum(1) Desc "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If

    If rsTmp.EOF Then
        mstrExtData = mstrExtData & ";Ѫ��"
    Else
        mstrExtData = mstrExtData & ";" & rsTmp("�걾����")
    End If

    '��ȡ�ɼ���ʽ
    Set rsTmp = SelectCap(Split(Split(mstrExtData, ";")(0), ",")(0))
    If rsTmp Is Nothing Then
        MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName
        Exit Sub
    End If
    mlngCapID = rsTmp("ID")

'    strsql = "Select C.��Ŀ��� From ������ĿĿ¼ A,���鱨����Ŀ B,������Ŀ C " & _
'        "Where A.ID=B.������ĿID And B.������ĿID=C.������ĿID And A.ID=[1]"
'    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, Split(Split(mstrExtData, ";")(0), ",")(0))
'    If rsTmp.EOF Then
'        mbln΢������Ŀ = False
'    Else
'        mbln΢������Ŀ = IIf(Nvl(rsTmp("��Ŀ���"), 0) = 2, True, False)
'    End If

    Call AdviceSet�������(3, mstrExtData)
'    txtҽ������.Text = Get�����������(2, "")
'    txtҽ������.Text = txtҽ������.Text & "(" & Split(mstrExtData, ";")(1) & ")"
    Me.txt���� = Split(mstrExtData, ";")(1)

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SelectCap(Optional ByVal lngItemID As Long = 0) As ADODB.Recordset
'��ȡ�ɼ���ʽ
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim tmpRect As RECT

    On Error GoTo DBError

    strSQL = "Select Distinct A.ID,A.����,A.���� " + _
        "From ������ĿĿ¼ A,�����÷����� D Where A.ID=D.�÷�ID" + _
        " And A.���='E' And A.��������='6'" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
        " And A.������� IN(" & PatientType & ",3) And Nvl(A.�����Ա�,0) IN (" + _
        IIf(Me.cbo�Ա�.Text Like "*��*", "1,0)", "2,0)") + _
        " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
        " And D.��ĿID=" & lngItemID
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.EOF Then
        strSQL = "Select Distinct A.ID,A.����,A.���� " + _
            "From ������ĿĿ¼ A Where " + _
            " A.���='E' And A.��������='6'" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
            " And A.������� IN(" & PatientType & ",3) And Nvl(A.�����Ա�,0) IN (" + _
            IIf(Me.cbo�Ա�.Text Like "*��*", "1,0)", "2,0)") + _
            " And Nvl(A.ִ��Ƶ��,0) IN(0,1)"
    End If
    If rsTmp.State = adStateOpen Then rsTmp.Close: Set rsTmp = New ADODB.Recordset
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then Set SelectCap = rsTmp

    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceSet�������(ByVal int���� As Integer, ByVal strDataIDs As String)
'���ܣ�1.��������ָ����������Ŀ�Ĳ�λ��,�����������������Ŀ���޸Ĳ�λ
'      2.��������ָ��������Ŀ�ĸ���������������Ŀ��,����������������Ŀ��������Ŀ�ĸ���������������Ŀ
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
'      strDataIDs=���:������鲿λ��Ϣ,����:��������������������Ŀ��Ϣ,���п���û�и�������������
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant

    On Error GoTo errH

    '���������Ŀ
    strDataIDs = Mid(strDataIDs, 1, InStr(strDataIDs, ";") - 1)

    If strDataIDs <> "" Then
        If Not rsRelativeAdvice Is Nothing Then
            rsRelativeAdvice.Close
        Else
            Set rsRelativeAdvice = New ADODB.Recordset
        End If
        strSQL = "Select ID,����,����,nvl(�걾��λ,' ') As �걾��λ," + _
        "���,nvl(�Ƽ�����,0) As �Ƽ�����,nvl(ִ�п���,0) As ִ�п���,�������� From ������ĿĿ¼ Where ID IN(" & strDataIDs & ")"
        Set rsRelativeAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        If Not rsRelativeAdvice Is Nothing Then rsRelativeAdvice.Close: Set rsRelativeAdvice = Nothing
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get�����������(ByVal int���� As Integer, ByVal txtMainAdvice As String) As String
'���ܣ��������ɼ���������ݵ�ҽ������
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
    Dim lngBegin As Long, i As Long
    Dim str���� As String, strTmp As String
    Dim strDate As String

    If rsRelativeAdvice Is Nothing Or int���� = 1 Then Get����������� = txtMainAdvice: Exit Function

    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("����"))) > 0 Then
            strTmp = strTmp & "," & rsRelativeAdvice("����")
        End If

        rsRelativeAdvice.MoveNext
    Loop

    If strTmp <> "" Then
        Get����������� = IIf(Len(Trim(txtMainAdvice)) = 0, "", txtMainAdvice & " �� ") & Mid(strTmp, 2)
    Else
        Get����������� = txtMainAdvice
    End If
End Function

Private Sub vsf2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strPh As String, strMsg As String
    Dim mlngLoop As Long, lngItemRow  As Long

    Select Case Col
        Case 2
            If vsf2.RowData(Row) = -1 Then
                '�ֹ��걾��
                If gblnManualPH Then
                    strPh = ValidPH(vsf2.TextMatrix(Row, Col), strMsg)
                    If Len(strMsg) > 0 Then
                        MsgBox strMsg, vbOKOnly + vbInformation, gstrSysName
                        vsf2.TextMatrix(Row, Col) = ""
                    Else
                        vsf2.TextMatrix(Row, Col) = strPh
                    End If
                End If
            End If
        Case 5
            If Val(vsf2.RowData(Row)) = 0 Then Exit Sub

            If vsf2.TextMatrix(Row, Col) = "-1" Then
            '����
                vsf2.TextMatrix(Row, 2) = TransSampleNO_PH(Val(CalcNextCode(Val(vsf2.RowData(Row)), Row, 1)), vsf2.RowData(Row))
            Else
                vsf2.TextMatrix(Row, 2) = TransSampleNO_PH(Val(CalcNextCode(Val(vsf2.RowData(Row)), Row, 0)), vsf2.RowData(Row))
            End If
    End Select
End Sub

Private Sub vsf2_BeforeDeleteCell(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf2_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf2_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
'    zlCommFun.PressKey vbKeyTab
    SetFocusNextIndex Me.vsf2.TabIndex
    Cancel = True
End Sub

Private Sub vsf2_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    On Error GoTo errH
    If mblnBarCode And KeyAscii = vbKeyReturn Then
        KeyAscii = 0: Me.cbo��������.SetFocus: gintSelectFocus = 2
    Else
        If KeyAscii = vbKeyReturn Then
            If Row + 1 = vsf2.Rows Then
                KeyAscii = 0: Me.cbo��������.SetFocus: gintSelectFocus = 2
            Else
                vsf2.Row = Row + 1
                vsf2.Col = 1
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    On Error GoTo errH
    Select Case Col
        Case 2
            If KeyAscii = 13 Then
'                KeyAscii = 0: Me.DTP(0).SetFocus: gintSelectFocus = 2: Exit Sub
'                KeyAscii = 0: cbo��������.SetFocus: gintSelectFocus = 2: Exit Sub
                If Row + 1 = vsf2.Rows Then
                    KeyAscii = 0: Me.cbo��������.SetFocus: gintSelectFocus = 2
                Else
                    vsf2.Row = Row + 1
                    vsf2.Col = 2
                End If
            End If
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If vsf2.RowData(vsf2.Row) <> -1 Then
                KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
            Else
                '�ֹ��걾��
                If gblnManualPH Then
                    KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789-")
                Else
                    KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
                End If
            End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
'���ҽ�����ݵĺϷ���
Private Function ValidAdvice() As Boolean
    ValidAdvice = True

    On Error Resume Next
    If txt����.Text = "" Then
        ValidAdvice = False
        MsgBox "�����벡�˵�������", vbInformation, gstrSysName: gintSelectFocus = 2 'DoEvents
'        mintFocusItem = FocusItem.����
        txt����.SetFocus: Exit Function
    End If

    If Len(Trim(Me.txtҽ������)) = 0 And mblnCheckIn = False Then
        ValidAdvice = False
        MsgBox "��������������Ŀ��", vbInformation, gstrSysName: gintSelectFocus = 2 'DoEvents
'        mintFocusItem = FocusItem.ҽ������
        Me.txtҽ������.SetFocus: Exit Function
    End If
    If Me.cbo��������.ListIndex = -1 And mblnCheckIn = False Then
        ValidAdvice = False
        MsgBox "��ָ���������ң�", vbInformation, gstrSysName: gintSelectFocus = 2 'DoEvents
'        mintFocusItem = FocusItem.��������
        Me.cbo��������.SetFocus: Exit Function
    End If
    If Len(Trim(Me.cboҽ��.Text)) = 0 And mblnCheckIn = False Then
        ValidAdvice = False
        MsgBox "��ָ������ҽ����", vbInformation, gstrSysName: gintSelectFocus = 2 'DoEvents
'        mintFocusItem = FocusItem.ҽ��
        Me.cboҽ��.SetFocus: Exit Function
    End If
End Function

Private Function SaveAdviceData(blnNewPatient As Boolean) As Long
    '����                   blnNewPatient �Ƿ����²���
    Dim strSQL As String, strDate As String, strNO As String
    Dim lngAdviceID As Long, lngTmpID As Long, lngSendNO As Long
    Dim lngMaxSeq As Long, iSendSeq As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim lng��������ID As Long, lng����ID As Long, strDoctor As String, i As Integer
    Dim strִ�п���ID As String, strִ�п���ID1 As String, lngDept As Long
    Dim rsCard As ADODB.Recordset
    Dim tmpstr��� As String, tmplngClinicID As Long, tmpint�Ƽ����� As Integer, tmpintִ������ As Integer
    Dim rsDept As ADODB.Recordset
    Dim lngPatientHomePage As Long
    Dim blnNewpatinet As Boolean    '�²���
    Dim blnPatientType As Boolean   '�Ǳ�ʶΪ��������

    On Error GoTo ErrHand
    blnPatientType = zlDatabase.GetPara("���еǼǲ��˱�ʶΪ����", 100, 1208, 0)


    '���没����Ϣ
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    blnNewpatinet = blnNewPatient
    '����ҽ��������
    lngAdviceID = zlDatabase.GetNextId("����ҽ����¼")
    '�õ����ҽ�����
    gstrSql = "select max(���) as ��� from ����ҽ����¼ where ����id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.ControlBox, mlng����ID)
    If rsTmp.EOF = False Then
        lngMaxSeq = Val(Nvl(rsTmp("���"), 0))
    Else
        lngMaxSeq = 0
    End If

    lng��������ID = Me.cbo��������.ItemData(Me.cbo��������.ListIndex)
    strDoctor = NeedName(Me.cboҽ��.Text)
    If ItemDeptID = 0 Then
        MsgBox "��û��ѡ��һ��ִ�п��Ҳ��ܽ��б��棬������������ѡ��һ�������ٽ��б��棡", vbInformation, Me.Caption
        SaveAdviceData = -1
        Exit Function
    Else
        strִ�п���ID = ItemDeptID
    End If

    iSendSeq = 1
    '������Ŀ���ɼ���ʽ��Ϊ��ҽ��
    tmplngClinicID = mlngCapID
    'ȡ�ɼ���ʽ��ִ�в���
    strִ�п���ID1 = "NULL"

    lngSendNO = zlDatabase.GetNextNo(10)
    strNO = zlDatabase.GetNextNo(IIf(PatientType = 2, 14, 13))

    gstrSql = "select nvl(max(��ҳID),0) as ��ҳID from ������ҳ where ����ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, mlng����ID)
    lngPatientHomePage = rsTmp("��ҳID")

    '�����סԺ���˵Ǽ��Ƿ��ʶΪ��������
    If blnPatientType = False Then
        If blnNewpatinet = True Then
            PatientType = 3
        End If
    Else
        PatientType = 3
    End If

    If mblnCheckIn = True And rsRelativeAdvice Is Nothing Then Exit Function

    '�������ҽ��
    If Not rsRelativeAdvice Is Nothing Then
        lngMaxSeq = lngMaxSeq + 1
        rsRelativeAdvice.MoveFirst
        Do While Not rsRelativeAdvice.EOF
            lngTmpID = zlDatabase.GetNextId("����ҽ����¼")
            With rsRelativeAdvice
                strSQL = "ZL_����ҽ����¼_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                    lngMaxSeq & "," & PatientType & "," & mlng����ID & "," & IIf(lngPatientHomePage = 0, "NULL", lngPatientHomePage) & "," & _
                    "0,1," & _
                    "1,'" & .Fields("���") & "'," & _
                    .Fields("ID") & ",NULL,NULL,NULL,1," & _
                    "'" & Replace(.Fields("����"), "'", "''") & "',''," & _
                    "'" & Me.txt���� & "','һ����',NULL,NULL,'',NULL," & _
                    .Fields("�Ƽ�����") & "," & _
                    strִ�п���ID & "," & _
                    .Fields("ִ�п���") & ",0," & strDate & ",NULL," & _
                    IIf(Val(Me.txtPatientDept.Tag) = 0, lng��������ID, Val(Me.txtPatientDept.Tag)) & "," & lng��������ID & ",'" & strDoctor & "'," & _
                    "Sysdate,''," & lngAdviceID & ")"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption

                iSendSeq = iSendSeq + 1
                strSQL = "ZL_����ҽ������_Insert(" & _
                    lngTmpID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
                    iSendSeq & ",1,NULL,NULL," & _
                    "Sysdate+1/(24*3600)," & _
                    "0," & strִ�п���ID & ",0,0)"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                .MoveNext
            End With
        Loop
    End If
    '��������Ĳɼ���ʽ�ŵ����
    lngMaxSeq = lngMaxSeq + 1
    strSQL = "ZL_����ҽ����¼_Insert(" & lngAdviceID & ",NULL," & _
        lngMaxSeq & "," & PatientType & "," & mlng����ID & "," & IIf(lngPatientHomePage = 0, "NULL", lngPatientHomePage) & "," & _
        "0,1," & _
        "1,'E'," & mlngCapID & ",NULL,NULL,NULL,1," & _
        "'" & Replace(Me.txtҽ������, "'", "''") & "',''," & _
        "'" & Me.txt���� & "','һ����',NULL,NULL,'',NULL,2," & _
        strִ�п���ID & ",3,0," & strDate & ",NULL," & _
        IIf(Val(Me.txtPatientDept.Tag) = 0, lng��������ID, Val(Me.txtPatientDept.Tag)) & "," & lng��������ID & ",'" & strDoctor & "'," & _
        "Sysdate,''," & lngAdviceID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    iSendSeq = iSendSeq + 1
    '������ҽ��
    strSQL = "ZL_����ҽ������_Insert(" & _
        lngAdviceID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
        iSendSeq & ",1,NULL,NULL," & _
        "Sysdate+1/(24*3600)," & _
        "0," & strִ�п���ID & ",0,1)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    SaveAdviceData = lngAdviceID

    Exit Function
ErrHand:

    Err.Raise Err.Number, "�걾����"
End Function

Private Sub SetPatientInfoWrite(blnTrue As Boolean)
    '����         ���ò�����Ϣ�Ƿ����д��
    '����         �Ƿ��д��
    Dim blnModifyInfo As Boolean                        '�Ƿ����޸Ĳ�����Ϣ

    blnModifyInfo = zlDatabase.GetPara("�Ǽ�ʱ��ֱ�����벡����Ϣ", 100, 1208, 0)
    If blnModifyInfo = 0 Then blnTrue = True

    Me.txtID.Locked = blnTrue
    Me.txtID.Enabled = Not blnTrue
    Me.txtBed.Locked = blnTrue
    Me.txtBed.Enabled = Not blnTrue
    Me.txtPatientDept.Locked = blnTrue
    Me.txtPatientDept.Enabled = Not blnTrue

End Sub
Private Function GetPatientInfo(lngID As String) As ADODB.Recordset

    gstrSql = "Select 1 As Patienttype, 0 As ��ҳid, A.���˿���, A.����, Decode(B.����, Null, A.�Ա�, B.���� || '-' || A.�Ա�) As �Ա�," & vbNewLine & _
                "       A.����, A.����id, C.סԺ��, C.�����, A.���� as ��ǰ����,a.��ʶ��,a.������ as ҽ��,Zl_Age_Calc(A.����ID) as ����1 " & vbNewLine & _
                " From ����걾��¼ A, �Ա� B, ������Ϣ C" & vbNewLine & _
                " Where A.�Ա� = B.����(+) And A.����id = C.����id and " & IIf(IsNumeric(lngID) = False, " 1 = 2 and ��ʶ�� = [1] ", " ��ʶ�� = [1] ")

    Set GetPatientInfo = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lngID))
End Function
Public Function zlRefresh_bak(ByVal lngSampleID As Long) As Boolean
'��ʾ�걾������Ϣ
'lngSampleID���걾��¼ID
    Dim rs As New ADODB.Recordset
    Dim mstrSql As String

    On Error GoTo ErrHand


    mstrSql = "Select A.��ҳid, A.����id, A.����, L.���� || '-' || A.�Ա� As �Ա�, A.����," & vbNewLine & _
                "       Decode(A.������Դ, 3, To_Char(A.NO), 1, To_Char(A.�����), 2, To_Char(A.סԺ��), 4, To_Char(A.�����)) As ���˺�," & vbNewLine & _
                "       Decode(A.סԺ��, Null, Null, A.����) As ����, A.�걾����, A.����ʱ��," & vbNewLine & _
                "       Decode(A.������Դ, 1, '����', 2, 'סԺ', 3, '����', 4, '���') As ������Դ," & vbNewLine & _
                "       Decode(A.����id, Null," & vbNewLine & _
                "               To_Char(Trunc(A.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(A.�걾���, 10000), '0000')," & vbNewLine & _
                "               A.�걾���) As �걾���, A.������ As ����ҽ��, A.����ʱ�� As ����ʱ��," & vbNewLine & _
                "       Nvl(A.���˿���, 'δ֪') As ���˿���, Nvl(B.����, 'δ֪') As ��������, Nvl(C.����, '�ֹ�') As ��������," & vbNewLine & _
                "       A.������Ŀ As ҽ������, A.������, A.����ʱ��, A.�걾��̬, A.����id, Nvl(A.�걾���, 0) As �걾���," & vbNewLine & _
                "       Nvl(A.�걾���, 0) As �걾���, A.����ʱ��, A.��ʶ��, A.���� As ����1" & vbNewLine & _
                "From ����걾��¼ A, ���ű� B, �������� C, �Ա� L" & vbNewLine & _
                "Where A.�Ա� = L.����(+) And A.�������id = B.ID(+) And A.����id = C.ID(+) and a.id = [1] "

    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, lngSampleID)

    On Error Resume Next
    If rs.EOF Then
        ClearItem

        Me.cbo��������.ListIndex = -1: Me.cboҽ��.ListIndex = -1
        Me.cbo(0).ListIndex = -1: Me.cbo(1).ListIndex = -1
    Else
        Me.txt���� = Nvl(rs("����"))
        Me.cbo�Ա�.Text = Nvl(rs("�Ա�"))
'        Me.txt���� = IIf(IsNull(rs("����")), "", Val(rs("����"))): If Me.txt���� = "0" Then Me.txt���� = ""
'        Me.txt���� = IIf(IsNull(rs("����")), "", IIf(IsNumeric(rs("����")), Val(rs("����")), Mid(rs("����"), 1, Len(rs("����")) - 1))): If Me.txt���� = "0" Then Me.txt���� = ""

        If IsNull(rs("����")) Then
            Me.txt���� = ""
        Else
            Me.txt���� = Val(rs("����"))
            If Me.txt���� = 0 Then Me.txt���� = ""
        End If

        If IsNull(rs("����")) = True Then
            Me.cboAge.Text = "��"
        Else
            If Val(rs("����")) = 0 Then
                Me.cboAge.Text = rs("����")
            Else
                Me.cboAge.Text = Mid(rs("����"), Len(CStr(Val(rs("����")))) + 1)
            End If
        End If
        'Me.cboAge.Text = IIf(IsNull(rs("����")), "��", Right(rs("����"), 1))
        If cboAge.ListIndex = -1 Then cboAge.ListIndex = 0
        Me.txtPatientDept = Nvl(rs("���˿���"))
        Me.txtID = Nvl(rs("���˺�"), Nvl(rs("��ʶ��")))
        Me.txtBed = Nvl(rs("����"), Nvl(rs("����1")))

        Me.cbo��������.Text = Nvl(rs("��������"))
        Me.cboҽ��.Text = Nvl(rs("����ҽ��"))
        Me.txt���� = Nvl(rs("�걾����"))

        Me.DTP(1).Value = rs("����ʱ��")
        Me.cbo(1).Text = Nvl(rs("������"))
        Me.DTP(0).Value = rs("����ʱ��")

        Me.cbo(0).Text = Nvl(rs("�걾��̬"))

        With vsf2
            .Rows = 2
            .RowData(1) = Nvl(rs("����ID"), -1)
            .TextMatrix(1, 1) = Nvl(rs("��������"))
            .TextMatrix(1, 2) = Nvl(rs("�걾���"))
            .TextMatrix(1, 5) = IIf(rs("�걾���") = 0, 0, -1)
        End With
        Me.txtҽ������ = ""
        Do While Not rs.EOF
            Me.txtҽ������ = Me.txtҽ������ & "," & Nvl(rs("ҽ������"))

            rs.MoveNext
        Loop
        rs.MoveFirst
        If Len(Me.txtҽ������) > 0 Then Me.txtҽ������ = Mid(Me.txtҽ������, 2)
        If Nvl(rs!������Դ, 0) = 3 Then
            lblCash.Caption = ""
        Else
            lblCash.Caption = IIf(CheckChargeState(lngSampleID, False, False), "��", "")
        End If
    End If

    SetPatientInfoWrite True
    zlRefresh_bak = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetFocusNextIndex(TabIndex As Integer)
    '���ܣ�         ģ�ⰴ��Tab����������һ���ؼ�
    '����:          TabIndex ��ǰ�ؼ���TabIndex
    Dim objThis As Object
    Dim intLoop As Integer
    On Error Resume Next
    For intLoop = TabIndex + 1 To Me.Count - 1
        For Each objThis In Me.Controls
            If objThis.TabIndex = intLoop Then
                If TypeName(objThis) = "VsfGrid" Then
                    objThis.Col = 2
                    objThis.ShowCell vsf2.Row, vsf2.Col
                    objThis.SetFocus
                    gintSelectFocus = 2
                    Exit Sub
                Else
                    If objThis.Enabled = True And objThis.Visible = True Then
                        objThis.SetFocus
                        gintSelectFocus = 2
                        Exit Sub
                    End If
                End If

            End If
        Next
    Next
End Sub

Private Function ShowCharge(ByVal lngKey As Long) As Integer
    '����: ����ҽ����ʾ�շ�״̬
    '����: ҽ��ID
    '����: -1=û���շѵ� 0=���۵� 1=���շ�
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim strReplace As String

    strSQL = "select ������Դ from ����걾��¼ where id = [1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork", lngKey)
    If rs.EOF = True Then Exit Function

    If rs("������Դ") <> 2 Then
        strReplace = "������ü�¼"
    End If

    ShowCharge = -1
    strSQL = _
        "select NVL(A.��¼״̬,-1) As ��¼״̬ " & _
              "from סԺ���ü�¼ A, " & _
              "( " & _
                   "select No,��¼���� from ����ҽ������ where ҽ��id IN (Select ID From ����ҽ����¼ A,(Select ҽ��id From ����걾��¼ Where ID= [1] Union Select ҽ��id From ������Ŀ�ֲ� Where �걾id= [1]) B where B.ҽ��id =A.���id and A.������� = 'C'  ) " & _
                   "Union " & _
                   "select No,��¼���� from ����ҽ������ where ҽ��id IN (Select ID From ����ҽ����¼ A,(Select ҽ��id From ����걾��¼ Where ID= [1] Union Select ҽ��id From ������Ŀ�ֲ� Where �걾id= [1]) B where B.ҽ��id =A.���id and A.������� = 'C'  ) " & _
              ") B " & _
            "Where A.NO = B.NO and mod(a.��¼����,10) = b.��¼���� Order By NVL(A.��¼״̬,-1)"



    If strReplace <> "" Then
        strSQL = Replace$(strSQL, "סԺ���ü�¼", strReplace)
    End If
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork", lngKey)

    If rs.BOF Then Exit Function

    ShowCharge = rs("��¼״̬").Value
End Function

Private Function CreatePatient() As Boolean
    '���ܽ���������Ϣ
     '���没����Ϣ
    Dim strDate As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strCostType As String
    Dim i As Long
    
    Dim strAge As String
    Dim strInfo As String
    Dim lngTmp As Long
    On Error GoTo errH

    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    If mlng����ID <> 0 Then
        strSQL = "select ����ID from ������Ϣ where ����id = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        If rsTmp.EOF = True Then
            mlng����ID = 0: PatientType = 1
        End If
    End If
    If PatientType = 1 And mlng����ID <= 0 Then '���ﲡ��
        If mlng����ID > 0 Then '���еĲ���
'            strsql = _
                "zl_�ҺŲ��˲���_INSERT(3," & mlng����ID & ",Null," & _
                "'',''," & _
                "'" & txt����.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & Replace(Replace(Me.cboAge.Text, "����", "��"), "Ӥ��", "��") & txt����1.Text & "'," & _
                "'�Է�','�Է�'," & _
                "'','',''," & _
                "'','','',0,'','','','',''," & strDate & ",NULL)"
'            strsql = "Zl_���鲡�˲���_Insert(3," & mlng����ID & ",'" & txt����.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & _
                        txt����.Text & Replace(Replace(Me.cboAge.Text, "����", "��"), "Ӥ��", "��") & txt����1.Text & "')"
        Else '�²���
            If txt����.Locked = False Then
                strAge = txt����.Text
                If IsNumeric(strAge) Then strAge = strAge & cboAge.Text & txt����1.Text
                strInfo = CheckAge(strAge)
                If InStr(1, strInfo, "|") > 0 Then
                    lngTmp = Val(Split(strInfo, "|")(0)) '1��ֹ,0��ʾ
                    strInfo = Split(strInfo, "|")(1)
                    If lngTmp = 1 Then
                        MsgBox strInfo, vbInformation, gstrSysName
                        If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Exit Function
                    End If
                End If
            End If
            mlng����ID = zlDatabase.GetNextNo(1)
'            strsql = _
                "zl_�ҺŲ��˲���_INSERT(1," & mlng����ID & ",Null," & _
                "'',''," & _
                "'" & txt����.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & Replace(Replace(Me.cboAge.Text, "����", "��"), "Ӥ��", "��") & txt����1.Text & "'," & _
                "'�Է�','�Է�'," & _
                "'','',''," & _
                "'','','',0,'','','','',''," & strDate & ",NULL)"
            strSQL = "select ����,ȱʡ��־ from �ѱ� order by ����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork")
            Do While Not rsTmp.EOF
                i = i + 1
                If i = 1 Then
                    strCostType = rsTmp("����")
                End If
                If rsTmp("ȱʡ��־") = 1 Then
                    strCostType = rsTmp("����")
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            If strCostType = "" Then strCostType = "�Է�"
            strSQL = "Zl_���鲡�˲���_Insert(1," & mlng����ID & ",'" & txt����.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & _
                        txt����.Text & Replace(Replace(Me.cboAge.Text, "����", "��"), "Ӥ��", "��") & txt����1.Text & "','" & strCostType & "')"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If

        CreatePatient = True
    End If
    Exit Function
errH:
    mlng����ID = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub chkLastRual(lngKey As Long)
    '����   ����ϴν���Ƿ񳬱�
    '����   lnkey = �걾id
    Dim blnChk  As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim rsChk As New ADODB.Recordset

    blnChk = zlDatabase.GetPara("����ʱ��ʾ�ϴγ�����", 100, 1208, False)

    If blnChk = False Then Exit Sub

    On Error GoTo errH

    gstrSql = "select b.������Ŀid from ����걾��¼ a , ������ͨ��� b where a.id = b.����걾id and a.id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKey)

    Do Until rsTmp.EOF
        If Nvl(rsTmp("������ĿID")) > 0 Then
            '���ָ����
            gstrSql = "Select ID, ������Ŀid, ��������, ��������, ������,��д" & vbNewLine & _
                        "From (Select ID, ������Ŀid, ��������, ��������, ������,��д" & vbNewLine & _
                        "       From (Select A.Id, B.������Ŀid,f.��ʾ���� as ��������,f.��ʾ���� as  ��������, B.������,��д" & vbNewLine & _
                        "              From ����걾��¼ A, ������ͨ��� B," & vbNewLine & _
                        "                   (Select A.����id, B.������Ŀid" & vbNewLine & _
                        ",Zl_To_Number(Zl_Get_Reference(1, b.������Ŀid, a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������,a.����id, a.����)) as �ο�ID " & vbNewLine & _
                        "                     From ����걾��¼ A, ������ͨ��� B" & vbNewLine & _
                        "                     Where A.Id = B.����걾id And A.Id = [1]) C, ������Ŀ D,������Ŀ�ο� F " & vbNewLine & _
                        "              Where A.Id = B.����걾id And A.����id = C.����id And B.������Ŀid = C.������Ŀid And A.����ʱ�� Between Sysdate - 1 And Sysdate And" & vbNewLine & _
                        "                    A.Id < [1] And B.������Ŀid = [2] And B.������Ŀid = D.������Ŀid And c.�ο�id=F.ID(+)" & vbNewLine & _
                        "              Order By ID Desc)" & vbNewLine & _
                        "       Where Rownum = 1)" & vbNewLine & _
                        "Where ������ <= �������� Or ������ >= ��������"

            Set rsChk = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKey, CLng(Val(Nvl(rsTmp("������ĿID"), 0))))
            If rsChk.EOF = False Then
                MsgBox "��ǰ���˵���һ���걾�н��<" & rsChk("��д") & ">���꣡��ע�⣡", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        rsTmp.MoveNext
    Loop

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function CheckIsInclude(strSource As String, strTarge As String) As Boolean
    '���strSource�е�ÿһ���ַ��Ƿ���strTarge��
    Dim i As Long
    CheckIsInclude = False

    Select Case strTarge
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "ʱ��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+-_)(*&^%$#@!`~"
    Case "����ʱ��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+_)(*&^%$#@!`~"
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "������"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "��С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "�ɴ�ӡ�ַ�"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/."":;|\=+-_)(*&^%$#@!`~0123456789"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    CheckIsInclude = True
End Function
Public Sub SetPara()
    '������仯�˲���
    mMakeNoRule = zlDatabase.GetPara("�걾������ɹ���", 100, 1208, "��  ��")
    mblnLoadLastAdvice = zlDatabase.GetPara("�Ǽ�ʱ������һ��������Ŀ", 100, 1208, False)
    mblnCheckIn = Val(zlDatabase.GetPara("�Ǽ�ʱ����Ҫ������Ŀ", 100, 1208, 0))
    mintItemRule = Val(zlDatabase.GetPara("�ֹ���Ŀ����Ŀ�ۼӱ걾��", 100, 1208, 0))
    mSendReport = zlDatabase.GetPara("ʹ�ö����������", 100, 1208, 0)
    mblnEmerge = Val(zlDatabase.GetPara("����걾", 100, 1208, 0))
    mbln���۵�ģʽ = InStr(GetSysParVal(80, ""), "C") > 0
End Sub
Private Function GetSampleData(lngKey As Long) As String
    Dim rsTmp As New ADODB.Recordset
    'ȡ�걾����������ִ�
    gstrSql = "Select Distinct ������Ŀid, Decode(�ֲ�ҽ��, Null, �걾ҽ��, �ֲ�ҽ��) ҽ��id, ������, �����־, ����ο�, ������Ŀid, �������" & vbNewLine & _
                "From (Select B.������Ŀid, Null �걾ҽ��, C.ҽ��id �ֲ�ҽ��, B.������, B.�����־, B.����ο�, B.������Ŀid, B.�������" & vbNewLine & _
                "       From ����걾��¼ A, ������ͨ��� B, ������Ŀ�ֲ� C" & vbNewLine & _
                "       Where A.Id = B.����걾id And A.Id = C.�걾id And B.������Ŀid = C.��Ŀid And A.Id = [1]" & vbNewLine & _
                "       Minus" & vbNewLine & _
                "       Select B.������Ŀid, A.ҽ��id �걾ҽ��, Null �ֲ�ҽ��, B.������, B.�����־, B.����ο�, B.������Ŀid, B.�������" & vbNewLine & _
                "       From ����걾��¼ A, ������ͨ��� B" & vbNewLine & _
                "       Where A.Id = B.����걾id And A.Id = [1]) order by ������� "


    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKey)
    Do Until rsTmp.EOF
        GetSampleData = GetSampleData & "|" & rsTmp("������Ŀid") & Chr(1) & rsTmp("ҽ��id") & Chr(1) & _
                        rsTmp("������") & Chr(1) & rsTmp("�����־") & Chr(1) & rsTmp("����ο�") & Chr(1) & _
                        rsTmp("������Ŀid") & Chr(1) & rsTmp("�������")

        rsTmp.MoveNext
    Loop
    If GetSampleData <> "" Then
        GetSampleData = Mid$(GetSampleData, 2)
    End If
End Function

Private Function GetApplicationFormShowType() As Boolean
    If Not mobjLisInsideComm Is Nothing Then
        GetApplicationFormShowType = mobjLisInsideComm.GetApplicationFormShowType()
    Else
        GetApplicationFormShowType = False
    End If
End Function
