VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Begin VB.Form frmSendCardAndDeposit 
   BorderStyle     =   0  'None
   Caption         =   "Ԥ��������"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15030
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   15030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.TabStrip tbDeposit 
      Height          =   405
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   714
      Style           =   2
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   882
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����Ԥ��(&M)"
            Key             =   "K1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "סԺԤ��(&Z)"
            Key             =   "K2"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fra�ſ� 
      Caption         =   "��������Ϣ��"
      ForeColor       =   &H00C00000&
      Height          =   1305
      Left            =   45
      TabIndex        =   30
      Top             =   1500
      Width           =   14970
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1170
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   33
         TabStop         =   0   'False
         Tag             =   "����"
         Top             =   840
         Width           =   1485
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����"
         Height          =   360
         Left            =   3600
         TabIndex        =   25
         Top             =   840
         Width           =   788
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00EBFFFF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1170
         PasswordChar    =   "*"
         TabIndex        =   18
         Tag             =   "����"
         Top             =   405
         Width           =   2625
      End
      Begin VB.TextBox txtPass 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5520
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   20
         Tag             =   "����"
         Top             =   405
         Width           =   1750
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   360
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   840
         Width           =   1750
      End
      Begin VB.TextBox txtAudi 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7860
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   22
         Tag             =   "��֤"
         Top             =   405
         Width           =   1750
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   360
         Left            =   11625
         TabIndex        =   24
         Top             =   405
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   -2147483633
         CalendarTitleBackColor=   16744576
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   53608451
         CurrentDate     =   43424
      End
      Begin VB.CheckBox chkEndTime 
         Caption         =   "��ֹʹ��ʱ��"
         Height          =   240
         Left            =   9855
         TabIndex        =   23
         Top             =   465
         Width           =   1755
      End
      Begin MSComctlLib.TabStrip tbSendCard 
         Height          =   315
         Left            =   75
         TabIndex        =   17
         Top             =   0
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         Style           =   2
         TabFixedHeight  =   526
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         TabMinWidth     =   882
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "�����շ�(&1)"
               Key             =   "CardFee"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "�󶨿���(&2)"
               Key             =   "CardBind"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   540
         TabIndex        =   35
         Top             =   900
         Width           =   480
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   510
         TabIndex        =   34
         Top             =   450
         Width           =   510
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   4995
         TabIndex        =   19
         Top             =   465
         Width           =   480
      End
      Begin VB.Label lbl��֤ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��֤"
         Height          =   240
         Left            =   7335
         TabIndex        =   21
         Top             =   465
         Width           =   480
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3480
         TabIndex        =   31
         Top             =   15
         Width           =   120
      End
      Begin VB.Label lbl���㷽ʽ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���㷽ʽ"
         Height          =   240
         Left            =   4515
         TabIndex        =   26
         Top             =   900
         Width           =   960
      End
   End
   Begin VB.Frame fraԤ�� 
      Caption         =   "��סԺԤ����Ϣ��"
      ForeColor       =   &H00C00000&
      Height          =   1200
      Left            =   30
      TabIndex        =   28
      Top             =   105
      Width           =   14955
      Begin VB.TextBox txt������ 
         Height          =   360
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   14
         Top             =   735
         Width           =   2805
      End
      Begin VB.TextBox txtFact 
         Height          =   360
         Left            =   1215
         MaxLength       =   50
         TabIndex        =   3
         Top             =   345
         Width           =   1470
      End
      Begin VB.TextBox txt������� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9480
         MaxLength       =   30
         TabIndex        =   9
         Top             =   345
         Width           =   2445
      End
      Begin VB.ComboBox cboԤ������ 
         Height          =   360
         Left            =   6345
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   345
         Width           =   1770
      End
      Begin VB.TextBox txtԤ���� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EBFFFF&
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   3525
         MaxLength       =   12
         TabIndex        =   5
         Top             =   345
         Width           =   1335
      End
      Begin VB.CheckBox chk��λ�ɿ� 
         Caption         =   "��λ�ɿ�"
         Height          =   360
         Left            =   13050
         TabIndex        =   10
         Top             =   345
         Width           =   1320
      End
      Begin VB.TextBox txt�ɿλ 
         Height          =   360
         Left            =   1215
         MaxLength       =   50
         TabIndex        =   12
         Top             =   735
         Width           =   2745
      End
      Begin VB.TextBox txt�ʺ� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9480
         MaxLength       =   50
         TabIndex        =   16
         Top             =   735
         Width           =   4800
      End
      Begin VB.Label lblAccno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ʺ�"
         Height          =   240
         Left            =   8880
         TabIndex        =   15
         Top             =   795
         Width           =   480
      End
      Begin VB.Label lblBank 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Left            =   4440
         TabIndex        =   13
         Top             =   795
         Width           =   720
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ɿλ"
         Height          =   240
         Left            =   210
         TabIndex        =   11
         Top             =   795
         Width           =   960
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʵ��Ʊ��"
         Height          =   240
         Left            =   210
         TabIndex        =   2
         Top             =   405
         Width           =   960
      End
      Begin VB.Label lblStyle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ɿʽ"
         Height          =   240
         Left            =   5310
         TabIndex        =   6
         Top             =   405
         Width           =   960
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   240
         Left            =   8400
         TabIndex        =   8
         Top             =   405
         Width           =   960
      End
      Begin VB.Label lblDepositMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2955
         TabIndex        =   4
         Top             =   405
         Width           =   480
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   825
         TabIndex        =   29
         Top             =   1605
         Width           =   480
      End
      Begin VB.Label lblYBMoney 
         AutoSize        =   -1  'True
         Caption         =   "�����ʻ����:"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3585
         TabIndex        =   1
         Top             =   15
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin zlIDKind.ucQRCodePayButton btQRCodeTemp 
      Height          =   315
      Left            =   14640
      TabIndex        =   32
      Top             =   1305
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   556
   End
End
Attribute VB_Name = "frmSendCardAndDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*********************************************************************************************************************************************
'������Ԥ����������
'�ӿ�:
'    1.zlInit:��ʼ���ӿ�
'    2.zlRecalcCardFee-�¼��㿨���ã��ѱ�ҽ�ƿ��䶯ʱ����Ҫ����
'    3.zlSaveDataBeforCheckIsValid-������ݵĺϷ���:�ڱ���ǰ��Ҫ����
'    4.zlSaveData-ִ�����ݱ������
'    5.zlSaveDataAfter-���ݱ���ɹ���ŵ���
'�����¼�:
'    1.RequestRefreshPatiInf-�������¸���XML��ʽ��������ݣ�ˢ�²�����Ϣ
'    2.InputOver-��������¼�(��ʾ���һ��������ɣ��Ա�����ת����һ����������
'��������
'    1.zlGetSendCard-��ȡ��ǰ�ķ���������
'    2.zlSetCardNo-���������¸�ֵ
'    3.zlSetUnitInfo-���ýɿλ��Ϣ(������λ���˺ŵ�������ɺ���Ҫ����)
'    4.zlSetInsureInfo:����ҽ����Ϣ
'    5. zlClearControlInfo-��ǰ��ǰ�����������Ϣ
'    6. zlSetFocus����ƶ�
'��������
'    1.RealName-���õ�ǰ�����Ƿ������ʵ����֤(ʵ����֤����Ҫ��ֵ)
'    2.GetWidth-������
'    3.GetWidth-����߶�
'����:�ɹ�����true,���򷵻�False
'����:���˺�
'����:2019-11-25 14:32:57
'*********************************************************************************************************************************************
'-------------------------------------------------------------------------------------------------
'�ӿڱ���
Private mint����״̬ As Integer '0-����;1-�쳣����;2-�쳣����
Private WithEvents mbtQRCodePay As ucQRCodePayButton
Attribute mbtQRCodePay.VB_VarHelpID = -1
Private mfrmMain As Object
Private mlngModule As Long
Private mbln����Ԥ�� As Boolean, mblnסԺԤ�� As Boolean
Private mlngCardTypeID As Long '�����Ŀ����ID
Attribute mlngCardTypeID.VB_VarHelpID = -1
Private mblnView As Boolean
Private mblnAllowSendCard As Boolean, mblnAllowBoundCard As Boolean
Private mblnCancel As Boolean '�Ƿ�����
Private mlng�쳣ID As Long '�쳣����
Private mbytӦ�ó���   As Byte    '1-ҽ�ƿ�����;2-������Ϣ�Ǽ�;3-������Ժ �Ǽ�;4-ԤԼ�ҺŽ���
Private mblnShowDepositAndSendCard As Boolean '���ܴ治����Ԥ�����������ԣ���Ӧ����ʾ�ڽ����ϣ���Ҫ����Ͻ������ʾ
'------------------------------------------------------------------------------------------------
'�ڲ�����
Private mblnNotClick As Boolean
Private mobjOneCardComLib As clsOneCardComLib
Private mblnInited As Boolean '�Ƿ��ʼ���ɹ���,ֻ�г�ʼ���ɹ��ģ�������༭

Private mobjPubPatient As clsInterFacePatient   '���˹��������ӿ�
Private mobjService As clsService '�����򲿼���ҩƷ�����ļ��ٴ���������
Private mobjExseSvr As clsExpenceSvr
Private mobjPati As clsPatientInfo
Private mobjThirdSwap As clsThirdSwapCard   '�������׵���ؽӿ�
Private mblnICCard As Boolean
Private mdblRQCodeMoney As Double 'ɨ�븶֧�����
Private mbln��ͬ���� As Boolean 'Ԥ���Ϳ���Ϊͬһ�ֽ��㷽ʽ

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private mstr��λ�ʺ�  As String
Private mstr�ɿλ   As String
Private mstr��λ������ As String
Private mstrQRcode As String '��ǰɨ�븶�Ķ�ά��

'-------------------------------------------------------------------------------------------------
'��������
Private mbln���� As Boolean '���￨�����Լ��˷�ʽ��ȡ
Private mbytRegValidDays As Integer '�Һ���Ч��������Ҫ����ʵ������ϵ����ȱʡ��Чʱ��
Private mblnNewPatiMustSendCard As Boolean  '����ͬʱ���뷢��

'-------------------------------------------------------------------------------------------------
'�������
Private Type Ty_CardProperty
       objSendCard As Card  '��������
       lng����ID As Long
       lng�������� As Long
       bln��� As Boolean
       blnOneCard As Boolean '  '�Ƿ�������һ��ͨ�ӿ�,��ģʽ�£�Ʊ���ϸ����Ʊ�ŷ�Χ��ķ�����󶨿����շ�
       rs���� As ADODB.Recordset
       dblӦ�ս�� As Double
       dblʵ�ս�� As Double
End Type
Private mCurSendCard As Ty_CardProperty
Private mblnSendCardLocked As Boolean '�Ƿ���������
Private mrs���� As ADODB.Recordset
Private mintPriceGradeStartType As Integer   '���ü۸�ȼ�����:'   0-δ����,1-ֻ������վ��,2-ֻ������ҽ�Ƹ��ʽ,3-վ���ҽ�ƿʽ��������
Private mstrPriceGrade As String, mstrPrePriceGrade As String

Private mobjShowTotalMoneyControl As Object '��ʾɨ�븶�ܶ�Ŀؼ�
Private mobjCardFeePayCards As Cards  '����֧����ʽ
Private mstrQRCodeTypeIds_CardFee As String '���Ѷ�ά��ɨ�븶
Private mobjCardFeeItems As clsBalanceItems '���ѽ�����Ϣ
Private mblnBoundCarded As Boolean '�Ƿ��Ѿ��󶨿�
Private mrsCardFee As ADODB.Recordset  '���Ѽ�¼��

'-------------------------------------------------------------------------------------------------
'Ԥ��Ʊ�ݼ���ӡ���
Private mobjDepositFact As clsFactProperty 'Ԥ����Ʊ����
Private mblnDepositStrictly As Boolean 'Ԥ���Ƿ��ϸ����
Private mbytԤ��Ʊ�ݳ��� As Byte   'Ԥ��Ʊ�ݳ���
Private mblnDepositPrint As Boolean '�Ƿ��ӡ
Private mobjDepositPayCards As Cards  'Ԥ��֧����ʽ
Private mbytPrepayType As Byte '�ϴ�Ԥ������: 0-����סԺ;1-����;2-סԺ
Private mblnAllowInsureAccDeposit As Boolean  '�Ƿ�����ҽ�����˽�Ԥ��
Private mstrQRCodeTypeIds_Deposit As String 'Ԥ�����ά��ɨ�븶
Private mblnDepositLocked As Boolean 'Ԥ���������
Private mobjDepositItems As clsBalanceItems  'Ԥ����ǰ֧����Ϣ

'-------------------------------------------------------------------------------------------------
'ҽ����ر���
Private mcurYBMoney As Currency  'ҽ�������˻����
Private mintInsure As Integer  'ҽ������
Private mstrҽ���� As String    'ҽ����
Private mstr���� As String   'ҽ������

'-------------------------------------------------------------------------------------------------
'����������
Private mobjKeyboard As Object
'-------------------------------------------------------------------------------------------------
'�����¼�
Public Event RequestRefreshPatiInf(ByVal strCardNo As String, ByVal strPatiInfoXML As String)
Public Event InputOver()    '�������
Public Event ExcuteQRCodePayment() 'ִ��ɨ�븶
Public Event Activate() '�Ӵ��弤��
Public Event ExcuteReadQRCode() 'ɨ�����
Public Event ControlGotFocus(objControl As Object)
'Public Event ControlLostFocus(objControl As Object)

'-------------------------------------------------------------------------------------------------
'���Ա���
Private mblnRealName As Boolean '�Ƿ�ʵ����֤
 
Public Sub zlSetFocus()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��궨λ
    '���
    '����:���˺�
    '����:2020-01-13 17:53:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Me.Enabled And Me.Visible Then Me.SetFocus
    If fraԤ��.Visible Then
        If txtԤ����.Visible And txtԤ����.Enabled Then txtԤ����.SetFocus
    ElseIf fra�ſ�.Visible Then
        If txt����.Enabled And txt����.Visible Then
            txt����.SetFocus
        End If
    End If
End Sub

Public Function zlInit(ByVal frmMain As Object, ByVal lngModule As Long, ByVal bln����Ԥ�� As Boolean, ByVal blnסԺԤ�� As Boolean, _
    ByVal lngCardTypeID As Long, blnAllowSendCard As Boolean, ByVal blnAllowBoundCard As Boolean, ByVal blnAllowInsureAccDeposit As Boolean, _
    Optional btQRCodePay As Object, Optional objShowTotalMoneyControl As Object, Optional blnView As Boolean = False, _
    Optional objOneCardComLib As Object, Optional ByVal blnCancel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���ӿ�
    '���:frmMain-���õ�������
    '     lngModule-ģ���
    '     btQRCodePay-ɨ�븶��ť
    '     objShowTotalMoneyControl-��ʾ���ܶ�ؼ�:lable��Text
    '     bln����Ԥ��-�Ƿ������Ԥ��
    '     blnסԺԤ��-�Ƿ��סԺԤ��
    '     lngSendCardTypeID-��ǰ�������ID:����0ʱ���������blnAllowSendCard��blnAllowBoundCard-��Ч
    '     blnAllowSendCard-������
    '     blnAllowBoundCard-����󶨿�
    '     objOneCardComLib-һ��ͨ��������,nothingʱ�������´���һ��
    '     blnView-�Ƿ�鿴
    '     strPrivs-��ǰ����ģ��Ȩ��
    '     blnAllowInsureAccDeposit-�Ƿ�����ҽ���˻���Ԥ��
    '     blnCancel-��ǰ�Ƿ����ϲ���
    '����:
    '����:��ʼ���ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-23 14:18:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubPatient As clsInterFacePatient
    
    On Error GoTo errHandle
    mblnInited = False
    Set mfrmMain = frmMain: mlngModule = lngModule: mbln����Ԥ�� = bln����Ԥ��: mblnסԺԤ�� = blnסԺԤ��
    mblnAllowSendCard = blnAllowSendCard: mblnAllowBoundCard = blnAllowBoundCard
    mblnAllowInsureAccDeposit = blnAllowInsureAccDeposit: mblnCancel = blnCancel
    mlngCardTypeID = lngCardTypeID
    If objOneCardComLib Is Nothing Then
        Set mobjOneCardComLib = New clsOneCardComLib
        If mobjOneCardComLib.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle) = False Then Exit Function
    Else
        Set mobjOneCardComLib = objOneCardComLib
    End If
    Set mobjExseSvr = New clsExpenceSvr
    Call mobjExseSvr.zlInitCommon(glngSys, mlngModule, gcnOracle, gstrDBUser)
    
    Set mobjService = New clsService
    Call mobjService.zlInitCommon(glngSys, mlngModule, gcnOracle, gstrDBUser)
    Set mbtQRCodePay = btQRCodePay
    Call CreateObjectKeyboard
    
    Set mobjThirdSwap = New clsThirdSwapCard '��ʼ����������
    Call mobjThirdSwap.zlInitCompents(Me, mlngModule, mobjOneCardComLib)
    
    If GetPublicPatient(objPubPatient) = False Then Exit Function
    
    Set mobjShowTotalMoneyControl = objShowTotalMoneyControl
    mblnShowDepositAndSendCard = False
    mblnView = blnView
    
   
    If blnCancel Then zlInit = True: mblnInited = True: Exit Function
    
    zlInit = InitFace
     mblnInited = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlRecalcCardFee(ByVal objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼��㿨����Ϣ
    '���:
    '
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-23 17:44:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs���� As ADODB.Recordset, blnReReadCardFee As Boolean
    
    Set mobjPati = objPati
    If fra�ſ�.Visible Then zlRecalcCardFee = True: Exit Function
    
    ' mintPriceGradeStartType As Integer   '���ü۸�ȼ�����:'   0-δ����,1-ֻ������վ��,2-ֻ������ҽ�Ƹ��ʽ,3-վ���ҽ�ƿʽ��������
    If mintPriceGradeStartType >= 2 Then
       Call GetPriceGrade(gstrNodeNo, 0, 0, mobjPati.ҽ�Ƹ��ʽ, , , mstrPriceGrade)
        '��ȡ�۸�ȼ�
        If mstrPriceGrade <> mstrPrePriceGrade Then
            'Ҫ���»�ȡ�۸�ȼ�
            Set mrs���� = Nothing: blnReReadCardFee = True
                mstrPrePriceGrade = mstrPriceGrade
        End If
    End If
    Set rs���� = GetCardFee(blnReReadCardFee, mstrPriceGrade)
    Call InitCardFee '���ؿ�������
    
    Call ReLoadCardFee  '���¼���
End Function
Public Function zlSetCardNo(ByVal strCardNo As String, objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿ��Ÿ������ı���
    '���:objPati-������Ϣ��
    '
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 18:53:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    txt����.Text = strCardNo

    Call zlRecalcCardFee(objPati)
    zlSetCardNo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub zlSetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(9(С��))��1-��(ȱʡ��12(С��)��
    '����:���˺�
    '����:2014-04-09 11:46:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytFontSize As Byte
    Dim objControl As Control
    On Error GoTo errHandle
    
    bytFontSize = IIf(bytSize = 0, 9, 12)
    Me.Font.Size = bytFontSize
    
    For Each objControl In Me.Controls
        If UCase(TypeName(objControl)) <> UCase("ucQRCodePayButton") Then
            objControl.Font.Size = bytFontSize
            'Debug.Print TypeName(objControl)
            If UCase(TypeName(objControl)) = UCase("TextBox") Or UCase(TypeName(objControl)) = UCase("DTPicker") Then
                 objControl.Height = IIf(bytSize = 0, 300, 360)
            End If
        End If
    Next
    Me.Refresh
    
    fra�ſ�.Left = fraԤ��.Left
    fraԤ��.Top = Me.ScaleTop + 105
    If mbln����Ԥ�� Or mblnסԺԤ�� Then
        fra�ſ�.Top = fraԤ��.Top + fraԤ��.Height + 50
        Me.Height = Me.ScaleTop + fraԤ��.Height + IIf(mlngCardTypeID <> 0, fra�ſ�.Height, 0) + 250
    Else
        fra�ſ�.Top = fraԤ��.Top
        Me.Height = Me.ScaleTop + fra�ſ�.Height + 200
    End If
    
    'λ�õ���
 
    txt�ʺ�.Width = 4800 * (bytFontSize / 12)
    txtFact.Width = 1470 * (bytFontSize / 12)
    txtԤ����.Width = 1335 * (bytFontSize / 12)
    txt�������.Width = 2445 * (bytFontSize / 12)
    txt������.Width = 2805 * (bytFontSize / 12)
    'txt�ɿλ.Width = 2745 * (bytFontSize / 12)
    txt����.Width = 2625 * (bytFontSize / 12)
    txtPass.Width = 1750 * (bytFontSize / 12)
    txtAudi.Width = 1750 * (bytFontSize / 12)
    chkEndTime.Width = IIf(bytSize = 0, 1415, 1755)
  
    dtpDate.Width = IIf(bytSize = 0, 2100, 2625)

  
    lblDepositMoney.Left = txtFact.Left + txtFact.Width + 100
    txtԤ����.Left = lblDepositMoney.Left + lblDepositMoney.Width + 20

    txtFact.Left = lblFact.Left + lblFact.Width + 20
    txt�ɿλ.Left = txtFact.Left
    txt����.Left = txtFact.Left
    txt����.Left = txtFact.Left

    txt�ʺ�.Left = txt�������.Left
    
    txtԤ����.Top = txtFact.Top
    cboԤ������.Top = txtFact.Top
    txt�������.Top = txtFact.Top
    chk��λ�ɿ�.Top = txt�������.Top + (txt�������.Height - chk��λ�ɿ�.Height) \ 2
    lblCode.Top = txt�������.Top + (txt�������.Height - lblCode.Height) \ 2
    lblStyle.Top = lblCode.Top
    lblDepositMoney.Top = lblCode.Top
    lblFact.Top = lblCode.Top
    
    
    txt�ɿλ.Top = txtFact.Top + txtFact.Height + 50
    txt������.Top = txt�ɿλ.Top
    txt�ʺ�.Top = txt�ɿλ.Top
    
    lblAccno.Top = txt�ʺ�.Top + (txt�ʺ�.Height - lblAccno.Height) \ 2
    lblBank.Top = lblAccno.Top
    lblUnit.Top = lblAccno.Top
    
    'lblFact.Left = txtFact.Left - lblFact.Width - 20
    lblUnit.Left = lblFact.Left  'txt�ɿλ.Left - lblUnit.Width - 20
    
    lblDepositMoney.Left = txtԤ����.Left - lblDepositMoney.Width - 10
    lblStyle.Left = cboԤ������.Left - lblStyle.Width - 20
    lblAccno.Left = txt�ʺ�.Left - lblAccno.Width - 20
    lblCode.Left = txt�������.Left - lblCode.Width - 20
    
    txt������.Left = lblBank.Left + lblBank.Width + 20
    cboԤ������.Left = txt������.Left + txt������.Width - cboԤ������.Width
    lblStyle.Left = cboԤ������.Left - lblStyle.Width - 20
    
    txtPass.Top = txt����.Top
    txtAudi.Top = txt����.Top
    dtpDate.Top = txt����.Top
    
    
    txt����.Top = txt����.Top + txt����.Height + 50
    lbl���.Top = txt����.Top + (txt����.Height - lblAccno.Height) \ 2
    lbl���.Left = lbl����.Left
       
       
    cbo��������.Top = txt����.Top
    cbo��������.Left = txtPass.Left
    
    lbl���㷽ʽ.Top = cbo��������.Top + (cbo��������.Height - lbl���㷽ʽ.Height) \ 2
    chk����.Top = cbo��������.Top + (cbo��������.Height - chk����.Height) \ 2
    
   ' lbl����.Left = txt����.Left - lbl����.Width - 20
    lbl����.Left = txtPass.Left - lbl����.Width - 20
    lbl��֤.Left = txtAudi.Left - lbl��֤.Width - 20
    
    lbl����.Top = txt����.Top + (txt����.Height - lbl����.Height) \ 2
    lbl����.Top = lbl����.Top
    lbl��֤.Top = lbl����.Top
    Call Form_Resize
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

 
Public Sub zlSetUnitInfo(ByVal str��λ�ʺ� As String, ByVal str�ɿλ As String, ByVal str��λ������ As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ�λ�˺�
    '����:���˺�
    '����:2019-11-26 13:37:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstr��λ�ʺ� = str��λ�ʺ�: mstr�ɿλ = str�ɿλ: mstr��λ������ = str��λ������
End Sub
Public Sub zlSetInsueInfo(ByVal int���� As Integer, ByVal cur�˻���� As Currency, ByVal strҽ���� As String, ByVal str���� As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ����Ϣ
    '���:int����
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-27 20:07:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPay As Card
    mintInsure = int����: mcurYBMoney = cur�˻����: mstrҽ���� = strҽ����: mstr���� = str����
    lblYBMoney.Caption = "�����ʻ���" & Format(mcurYBMoney, "0.00")
    lblYBMoney.Visible = True And int���� <> 0
    Set objPay = GetDepositPayCard
    
    mblnNotClick = True
    Call LoadԤ�����㷽ʽ
    Call SetLoaclePayModefromCard(objPay, True)
    mblnNotClick = False
End Sub

Public Sub zlClearControlInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ؼ���Ϣ
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-26 13:53:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mobjPati = Nothing
     mstr��λ�ʺ� = "": mstr�ɿλ = "": mstr��λ������ = ""
     txtAudi.Text = "": txtPass.Text = ""
     txt����.Text = "": txt����.Text = ""
     txtFact.Text = "": txtԤ����.Text = "": txt�������.Text = "": txt�ɿλ.Text = ""
     txt������.Text = "": txt�ʺ�.Text = ""
     lblYBMoney.Caption = "�����ʻ����:"
     chk����.value = IIf(mbln���� = True, 1, 0)
     mintInsure = 0: mcurYBMoney = 0: mstrҽ���� = "": mstr���� = ""
     lblYBMoney.Visible = False
    If cboԤ������.ListCount > 0 Then cboԤ������.ListIndex = Val(cboԤ������.Tag)
    If cbo��������.ListCount > 0 Then cbo��������.ListIndex = Val(cbo��������.Tag)
    Set mobjDepositItems = Nothing
    Set mobjCardFeeItems = Nothing
    Call RefreshFactNo
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub


Public Function zlGetSendCard(ByRef objSendCard_Out As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ�ķ�������
    '���:
    '����:objSendCard_Out-���ص�ǰ�����Ķ���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 15:13:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mCurSendCard.objSendCard Is Nothing Then Exit Function
    Set objSendCard_Out = mCurSendCard.objSendCard
    zlGetSendCard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub RefreshFactNo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ͬ���ķ�Ʊ
    '����:���˺�
    '����:2011-07-19 17:47:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If mobjDepositFact Is Nothing Then Set mobjDepositFact = New clsFactProperty
    
    If mobjDepositFact.��ӡ��ʽ = 0 Then txtFact.Text = "": Exit Sub
    If mblnDepositStrictly = False Then
        '��ɢ��ȡ��һ������
        txtFact.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("��ǰԤ��Ʊ�ݺ�", glngSys, mlngModule, "")))
        Exit Sub
    End If
    '�ϸ�:     ȡ��һ������
    mobjDepositFact.����ID = mobjExseSvr.CheckUsedBill(2, IIf(mobjDepositFact.����ID > 0, mobjDepositFact.����ID, mobjDepositFact.LastUseID), , Val(Mid(tbDeposit.SelectedItem.Key, 2)))
    If mobjDepositFact.����ID <= 0 Then
        Select Case mobjDepositFact.����ID
            Case 0 '����ʧ��
'            Case -1
'                MsgBox "��û�����û��õ�Ԥ��Ʊ��,�Ǽǲ�����Ϣʱ����ͬʱ��Ԥ���" & _
'                    "��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
'            Case -2
'                MsgBox "���صĹ���Ʊ���Ѿ�����,�Ǽǲ�����Ϣʱ����ͬʱ��Ԥ���" & _
'                    "��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
        End Select
        txtFact.Text = ""
    Else
        txtFact.Text = mobjExseSvr.GetNextBill(mobjDepositFact.����ID)
    End If
End Sub

Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '���:
    '����:���˺�
    '����:2019-11-25 14:57:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    
    On Error GoTo errHandle
    
    'Ʊ�ݺ��볤�ȡ����￨�ų���
    strValue = zlDatabase.GetPara(20, glngSys, , "||||")
    mbytԤ��Ʊ�ݳ��� = Val(Split(strValue, "|")(1))
    

    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    mblnDepositStrictly = Mid(strValue, 2, 1) = "1" 'Ԥ���ϸ����

       
    strValue = zlDatabase.GetPara(21, glngSys, , "01") & "1"
    mbytRegValidDays = Val(Left(strValue, 1))
    If mbytRegValidDays < Val(Mid(strValue, 2, 1)) Then mbytRegValidDays = Val(Mid(strValue, 2, 1))
    
    mbytPrepayType = Val(zlDatabase.GetPara("�ϴ�Ԥ������", glngSys, mlngModule, "0"))
    
    mbln���� = zlDatabase.GetPara("���Ѽ���", glngSys, mlngModule) = "1"
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub
 

Private Function InitFace() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��
    '���:
    '
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-23 14:21:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnԺ�ⷢ�� As Boolean, blnBoundCard As Boolean
    Dim blnAllowSendCard As Boolean, blnAllowBoundCard As Boolean
    Dim objPubPatient As clsInterFacePatient
    Dim strҽ�Ƹ��ʽ As String, str���� As String
    Dim objSendCard As Card
    Dim varData As Variant, i As Long, varTemp As Variant
    
    
    On Error GoTo errHandle
    lblYBMoney.Visible = False
    
    Call InitPara '��ʼ������ֵ
    Call zlClearControlInfo
    
    If Load֧����ʽ(mbln����Ԥ�� Or mblnסԺԤ��, mlngCardTypeID > 0) = False Then Exit Function
    
    If Not mobjPati Is Nothing Then strҽ�Ƹ��ʽ = mobjPati.ҽ�Ƹ��ʽ
    
    mblnSendCardLocked = False
    
    fraԤ��.Tag = ""
    fraԤ��.Visible = mbln����Ԥ�� Or mblnסԺԤ�� Or mblnShowDepositAndSendCard
    fraԤ��.Tag = IIf(mbln����Ԥ�� Or mblnסԺԤ��, "", "1")
    tbDeposit.Visible = mbln����Ԥ�� And mblnסԺԤ��
    mblnNotClick = True
    If mbln����Ԥ�� And Not mblnסԺԤ�� Then
        'ֻ������Ԥ��
        fraԤ��.Caption = "������Ԥ����Ϣ��"
        tbDeposit.Tabs(1).Selected = True
    ElseIf mblnסԺԤ�� And Not mbln����Ԥ�� Then
        'ֻ��סԺԤ��
        fraԤ��.Caption = "��סԺԤ����Ϣ��"
         tbDeposit.Tabs(2).Selected = True
    Else
        '���߶���
        fraԤ��.Caption = "�����ＰסԺԤ����"
         tbDeposit.Tabs(1).Selected = True
    End If
    
    If mbln����Ԥ�� Or mblnסԺԤ�� Then
        
        With tbDeposit
            mblnNotClick = True
            .Tabs.Clear
            If mbln����Ԥ�� Then .Tabs.Add(, "K1", "����Ԥ��(&M)").Selected = IIf(mbytPrepayType = 1, True, False)
            If mblnסԺԤ�� Then .Tabs.Add(, "K2", "סԺԤ��(&Z)").Selected = IIf(mbytPrepayType = 2, True, False)
            If .Tabs.Count > 0 And .SelectedItem Is Nothing Then
               .Tabs(0).Selected = True
            End If
             
             mblnNotClick = False
            'If Not .SelectedItem Is Nothing Then Call tbDeposit_Click
            
             fraԤ��.Visible = .Tabs.Count <> 0
            If .Tabs.Count <> 0 And tbDeposit.SelectedItem Is Nothing Then Call RefreshFactNo
         End With
         
    End If
    mblnNotClick = False
    mintPriceGradeStartType = GetPriceGradeStartType()
    If mintPriceGradeStartType <> 0 Then
         Call GetPriceGrade(gstrNodeNo, 0, 0, strҽ�Ƹ��ʽ, , , mstrPriceGrade)   '��ȡ�۸�ȼ�
    End If
    
    Call SetCardEditEnabled(1, True)
    Call SetDepositEditEnabled(1)
    fra�ſ�.Visible = mlngCardTypeID <> 0
    tbSendCard.Visible = mlngCardTypeID <> 0
    
    If GetPublicPatient(objPubPatient) = False Then Exit Function
    
    Call SetDepositEditEnabled '����Ԥ
   
    fra�ſ�.Tag = ""
    lbl������.Visible = False
    If mlngCardTypeID <> 0 Then
        chk����.value = IIf(mbln����, 1, 0)
        chk����.Tag = IIf(mbln����, 1, 0)
        '�������󶨿�����
        If mobjOneCardComLib.zlGetCard(mlngCardTypeID, False, objSendCard) = False Then
            fra�ſ�.Visible = False: tbSendCard.Visible = False
            Exit Function
        End If
        If objSendCard Is Nothing Then
           fra�ſ�.Visible = False: tbSendCard.Visible = False
            Exit Function
        End If
        
        Set mCurSendCard.objSendCard = objSendCard
        mCurSendCard.lng�������� = 0
        
        str���� = zlDatabase.GetPara("����ҽ�ƿ�����", glngSys, mlngModule, "0")
        varData = Split(str����, "|")
        For i = 0 To UBound(varData)
             varTemp = Split(varData(i), ",")
             If Val(varTemp(0)) <> 0 Then
                If ExistShareBill(Val(varTemp(0)), 5) Then
                    If Val(varTemp(1)) = objSendCard.�ӿ���� Then
                        mCurSendCard.lng�������� = Val(varTemp(0)): Exit For
                    End If
                End If
             End If
        Next
        lbl������.Visible = True
        lbl������.Caption = "��" & objSendCard.���� & "��"
        
        txt����.PasswordChar = IIf(objSendCard.�������Ĺ��� <> "", "*", "")
        txt����.MaxLength = objSendCard.���ų���
        
        '��Чʱ�䴦��
        chkEndTime.value = 0
        If objSendCard.ȱʡ��Чʱ�� <> "" Then
            chkEndTime.value = vbChecked
            dtpDate = Format(objSendCard.zlGetDefaultDate, "yyyy-MM-dd 23:59:59")
        ElseIf objPubPatient.blnRealName Then
             dtpDate = Format(DateAdd("D", mbytRegValidDays, zlDatabase.Currentdate), "yyyy-MM-dd 23:59:59")
        Else
            dtpDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
            dtpDate.Enabled = False
        End If
        mblnNotClick = True
        
        'ҽ�ƿ�������ؿ���
        '1.�ȴ����Ƿ�������
        blnAllowSendCard = mblnAllowSendCard And objSendCard.�Ƿ񷢿�
        If blnAllowSendCard = False Then '������
            '�Ƴ�����ҳ
            Call RemoveSendCardTabFromKey("CardFee")
        ElseIf Not CheckIsExstSendCardTabFromKey("CardFee") Then
            '�����ڷ���ҳǩ����Ҫ����
            Call RemoveSendCardTabFromKey   '�Ƴ�����ѡ���������
            tbSendCard.Tabs.Add , "CardFee", "�շѷ���(&1)"
            tbSendCard.Width = GetSendCardTabsWidth
        End If
        
        '2.����󶨿�s
        
        blnAllowBoundCard = mblnAllowBoundCard And (objSendCard.���ƿ� = False Or objSendCard.�����ظ�ʹ��)   '����󶨿�
        If Not blnAllowBoundCard Then
            '�����ڰ󶨿���
            Call RemoveSendCardTabFromKey("CardBind")
        ElseIf Not CheckIsExstSendCardTabFromKey("CardBind") Then
            tbSendCard.Tabs.Add , "CardBind", "�󶨿���(&2)"
        End If
        tbSendCard.Width = GetSendCardTabsWidth
        
        'ȱʡ��λ
        Select Case zlDatabase.GetPara("����ģʽ", glngSys, mlngModule, "CardFee")
        Case "CardFee"
              mblnNotClick = True
              If CheckIsExstSendCardTabFromKey("CardFee") Then tbSendCard.Tabs("CardFee").Selected = True
              mblnNotClick = False
        Case "CardBind"
              mblnNotClick = True
              If CheckIsExstSendCardTabFromKey("CardBind") Then tbSendCard.Tabs("CardBind").Selected = True
              mblnNotClick = False
        End Select
        
        If tbSendCard.SelectedItem Is Nothing Then
            If tbSendCard.Tabs.Count > 0 Then
                mblnNotClick = True
                tbSendCard.Tabs(1).Selected = True
                mblnNotClick = False
            End If
        End If
        mblnNotClick = False
        tbSendCard.Width = GetSendCardTabsWidth '����ȱʡ�Ŀ��
        
        Call InitCardFee '���ؿ�������
        If objSendCard.�Ƿ��ϸ���� Then
            
            mCurSendCard.lng����ID = mobjExseSvr.CheckUsedBill(5, IIf(mCurSendCard.lng����ID > 0, mCurSendCard.lng����ID, mCurSendCard.lng��������), , objSendCard.�ӿ����)
            If mCurSendCard.lng����ID <= 0 Then
                Select Case mCurSendCard.lng����ID
                    Case 0 '����ʧ��
                    Case -1
                        'MsgBox "��û�����û��õľ��￨,���ܷ��ţ�" & vbCrLf & _
                        "�����ڱ������ù������λ�����һ���¿�! ", vbExclamation, gstrSysName
                    Case -2
                        ' MsgBox "���ع��õľ��￨������,���ܷ��ţ�" & vbCrLf & _
                        "���������ñ��ع��ÿ����λ�����һ���¿���", vbExclamation, gstrSysName
                    End Select
            End If
        End If
        '��ʼ��������Ϣֵ
        Call SetSendCardCtrolVisibled
    ElseIf mblnShowDepositAndSendCard Then
        fra�ſ�.Visible = True: fra�ſ�.Tag = "1"
        tbSendCard.Visible = False
        Call SetCardEditEnabled(0)
    End If
    
    If Not tbSendCard.SelectedItem Is Nothing Then tbSendCard_Click
    If Not tbDeposit.SelectedItem Is Nothing Then tbDeposit_Click
    InitFace = True
    Exit Function
errHandle:
    mblnNotClick = False
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitCardFee()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ֵ
    '����:���˺�
    '����:2019-11-25 10:32:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs���� As ADODB.Recordset
    Dim str�ѱ� As String, dblMoney As Double
    
    On Error GoTo errHandle
    Set rs���� = GetCardFee()
    
    If rs���� Is Nothing Then
        txt����.Text = "": txt����.Tag = ""
        Exit Sub
    End If
    If rs����.RecordCount = 0 Then
        txt����.Text = "": txt����.Tag = ""
        Exit Sub
    End If
    With rs����
        str�ѱ� = ""
        If Not mobjPati Is Nothing Then str�ѱ� = mobjPati.�ѱ�
        txt����.Text = Format(IIf(Nvl(!�Ƿ���, 0) = 1, Val(Nvl(!ȱʡ�۸�)), Val(Nvl(!�ּ�))), "0.00")
        If Nvl(!�Ƿ���, 0) <> 1 And Nvl(!���ηѱ�, 0) <> 1 Then
            If mobjExseSvr.zl_ExseSvr_Actualmoney(str�ѱ�, !�շ�ϸĿID, !������ĿID, Val(txt����.Text), dblMoney) Then
                txt����.Text = Format(dblMoney, "0.00")
            End If
        End If
        txt����.Tag = txt����.Text  '���ֲ���
        txt����.Locked = Nvl(!�Ƿ���, 0) <> 1
        txt����.TabStop = Nvl(!�Ƿ���, 0) = 1
    End With
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetSendCardTabsWidth() As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ҳ����ܿ��
    '����:�����ܿ��
    '����:���˺�
    '����:2019-11-23 16:12:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, lngWidth As Long
    
    For i = 1 To tbSendCard.Tabs.Count
        
        lngWidth = lngWidth + tbSendCard.Tabs(i).Width + Me.TextWidth("��")
    Next
    GetSendCardTabsWidth = lngWidth
End Function

Private Function CheckIsExstSendCardTabFromKey(ByVal strKey As String, Optional ByRef indIndex_Out As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��Ƿ����ָ��Keyֵ��ҳ��
    '���:
    '����:indIndex_Out-����ʱ�����ظ�tab������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-23 16:04:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    indIndex_Out = tbSendCard.Tabs(strKey).Index
    If Err <> 0 Then
        Err = 0: On Error GoTo 0: CheckIsExstSendCardTabFromKey = False
        Exit Function
    End If
    CheckIsExstSendCardTabFromKey = True
End Function


Private Function RemoveSendCardTabFromKey(Optional ByVal strKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƴ�ָ��Keyֵ��ҳ��
    '���:strKey=""ʱ����ʾ�Ƴ����п�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-23 16:04:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Err = 0: On Error Resume Next
    If strKey = "" Then
        tbSendCard.Tabs.Clear
    Else
        Call tbSendCard.Tabs.Remove(strKey)
        If Err <> 0 Then
            Err = 0: On Error GoTo 0
            
        End If
    End If
    RemoveSendCardTabFromKey = True
End Function
 
Private Sub btQRCodeTemp_GotFocus()
    RaiseEvent ControlGotFocus(btQRCodeTemp)
End Sub

 

Private Sub cbo��������_GotFocus()
    RaiseEvent ControlGotFocus(cbo��������)
End Sub

Private Sub cboԤ������_Click()
    Dim objPayCard As Card
    If mblnNotClick = True Then Exit Sub
    If Not cboԤ������.Enabled Then Exit Sub
    
    Set objPayCard = GetDepositPayCard
    If objPayCard Is Nothing Then Exit Sub
    
    Call SetDepositEditEnabled
    
    If txt�ɿλ.Text <> "" And txt�ɿλ.Enabled = True Then
        chk��λ�ɿ�.value = 1
    Else
        chk��λ�ɿ�.value = 0
    End If
    Call Local���㷽ʽ(objPayCard.�ӿ����, False, IIf(cboԤ������.ItemData(cboԤ������.ListIndex) <> 5, cboԤ������.Text, ""))
    'Call chk��λ�ɿ�_Click
End Sub

Private Sub cboԤ������_GotFocus()
    RaiseEvent ControlGotFocus(cboԤ������)
End Sub

Private Sub chkEndTime_GotFocus()
    RaiseEvent ControlGotFocus(chkEndTime)
End Sub

Private Sub chk��λ�ɿ�_Click()
    If chk��λ�ɿ�.value = 1 And cboԤ������.Enabled Then
        txt�ɿλ.Enabled = True
        txt�ɿλ.BackColor = &H80000005
    Else
        txt�ɿλ.Text = ""
        txt�ɿλ.Enabled = False
        txt�ɿλ.BackColor = Me.BackColor
    End If
End Sub
Private Sub cbo��������_Click()
    Dim objPayCard As Card
     
    If mblnNotClick = True Then Exit Sub
    
'    Set objPayCard = GetCardFeePayCard
'    If objPayCard Is Nothing Then Exit Sub
'    Call Local���㷽ʽ(objPayCard.�ӿ����, True)
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo��������.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cbo��������.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then
        cbo��������.ListIndex = lngIdx
    End If
End Sub

Private Sub chk��λ�ɿ�_GotFocus()
    
    RaiseEvent ControlGotFocus(chk��λ�ɿ�)
End Sub

Private Sub chk����_GotFocus()
    RaiseEvent ControlGotFocus(chk����)
End Sub

Private Sub dtpDate_GotFocus()
    RaiseEvent ControlGotFocus(dtpDate)
End Sub

Private Sub Form_Activate()
    RaiseEvent Activate
End Sub
 
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF6 'ɨ�븶���
            RaiseEvent ExcuteReadQRCode
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        If Me.ActiveControl Is chkEndTime Then
            If Not tbSendCard.SelectedItem Is Nothing Then
                If tbSendCard.SelectedItem.Key = "BoundCard" Then
                    If chkEndTime.value <> 1 Then
                        RaiseEvent InputOver    '�������
                        Exit Sub
                    End If
                End If
            End If
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        If Me.ActiveControl Is chk���� Then
            If cbo��������.Enabled And cbo��������.Visible Then cbo��������.SetFocus
            If chk����.value = Checked And Visible Then
                RaiseEvent InputOver    '�������
            End If
            Exit Sub
        End If
        
        
        If Me.ActiveControl Is dtpDate Then
            If Not tbSendCard.SelectedItem Is Nothing Then
                If tbSendCard.SelectedItem.Key = "BoundCard" Then
                    RaiseEvent InputOver    '�������
                    Exit Sub
                End If
            End If
        End If
        
        If Not (Me.ActiveControl Is txtԤ���� Or Me.ActiveControl Is txtPass Or Me.ActiveControl Is txtAudi Or Me.ActiveControl Is txt����) Then
            zlCommFun.PressKey vbKeyTab
        End If
        Exit Sub
     End If
     If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call HookDefend(txtPass.hWnd)
    Call HookDefend(txtAudi.hWnd)
End Sub
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    fra�ſ�.Width = Me.ScaleWidth - fra�ſ�.Left * 2
    fraԤ��.Left = fra�ſ�.Left
    fraԤ��.Width = fra�ſ�.Width
    
    txt�ʺ�.Left = fraԤ��.Left + fraԤ��.Width - txt�ʺ�.Width - fraԤ��.Left * 2 - 50
    lblAccno.Left = txt�ʺ�.Left - lblAccno.Width - 20
    txt�������.Left = txt�ʺ�.Left
    lblCode.Left = txt�������.Left - lblCode.Width
    chk��λ�ɿ�.Left = txt�ʺ�.Left + txt�ʺ�.Width - chk��λ�ɿ�.Width
    
    dtpDate.Left = fra�ſ�.Left + fra�ſ�.Width - dtpDate.Width - fra�ſ�.Left * 2 - 50
    chkEndTime.Top = dtpDate.Top + (dtpDate.Height - chkEndTime.Height) \ 2
    chkEndTime.Left = dtpDate.Left - chkEndTime.Width
    
    txtAudi.Left = chkEndTime.Left - txtAudi.Width - 200
    lbl��֤.Left = txtAudi.Left - lbl��֤.Width - 20
    
    txtPass.Left = lbl��֤.Left - txtPass.Width - 50
    lbl����.Left = txtPass.Left - lbl����.Width - 20
    
    cbo��������.Left = txtPass.Left
    lbl���㷽ʽ.Left = cbo��������.Left - lbl���㷽ʽ.Width - 20
    
    chk����.Left = lbl���㷽ʽ.Left - chk����.Width - 50
    
     lbl������.Left = fra�ſ�.Left + fra�ſ�.Width - lbl������.Width - 200
End Sub

Private Sub tbDeposit_Click()
    If mblnNotClick Then Exit Sub
    If tbDeposit.SelectedItem Is Nothing Then Exit Sub
    
    Set mobjDepositFact = mobjExseSvr.zl_GetInvoicePreperty(mlngModule, 2, Mid(tbDeposit.SelectedItem.Key, 2))
    mobjDepositFact.����ID = 0
    Call RefreshFactNo
    If txtԤ����.Enabled And txtԤ����.Visible Then txtԤ����.SetFocus
End Sub

Private Sub tbSendCard_Click()
    If mblnNotClick Then Exit Sub
    Call SetSendCardCtrolVisibled   '�����ؼ�λ�ü�visible����
End Sub

Private Sub chkEndTime_Click()
    dtpDate.Enabled = chkEndTime.value
End Sub
Private Sub chk����_Click()
    
    cbo��������.Enabled = chk����.value <> Checked
    Call CalcRQCodePayTotal '�����ܶ�

End Sub

Private Sub SetSendCardCtrolVisibled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������ؼ���Visibled���ԣ���������Ӧ�Ŀؼ�λ��
    '����:���˺�
    '����:2019-11-23 15:29:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnSendCard As Boolean
    Dim lngTop As Long
    
'    If Not fra�ſ�.Visible Then Exit Sub
    If tbSendCard.SelectedItem Is Nothing Then
        blnSendCard = False
    Else
        blnSendCard = tbSendCard.SelectedItem.Key = "CardFee"
    End If
    lbl���.Visible = blnSendCard
    txt����.Visible = blnSendCard
    chk����.Visible = blnSendCard
    cbo��������.Visible = blnSendCard
    lbl���㷽ʽ.Visible = blnSendCard
    '������Ӧλ��
    
    If blnSendCard Then
        lngTop = tbSendCard.Height + 45
    Else
        lngTop = (fra�ſ�.Height - txt����.Height + tbSendCard.Height \ 2) \ 2
    End If
    
    txt����.Top = lngTop: lbl����.Top = txt����.Top + (txt����.Height - lbl����.Height) \ 2
    txtPass.Top = lngTop: lbl����.Top = txtPass.Top + (txtPass.Height - lbl����.Height) \ 2
    txtAudi.Top = lngTop: lbl��֤.Top = txtAudi.Top + (txtAudi.Height - lbl��֤.Height) \ 2
    dtpDate.Top = lngTop: chkEndTime.Top = dtpDate.Top + (dtpDate.Height - chkEndTime.Height) \ 2
    
    txt����.Top = txt����.Top + txt����.Height + 80: lbl���.Top = txt����.Top + (txt����.Height - lbl���.Height) \ 2
    cbo��������.Top = txt����.Top
    chk����.Top = txt����.Top + (txt����.Height - chk����.Height) \ 2
    lbl���㷽ʽ.Top = cbo��������.Top + (cbo��������.Height - lbl���㷽ʽ.Height) \ 2
End Sub


Private Sub SetCardEditEnabled(Optional bytEnabledType As Byte, Optional blnInit As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���þ��￨�༭����
    '���:bytEnabledType-���õ�����:0-������;1-���ý�����Ϣ;2-����������Ϣ
    '       blnInit-�Ƿ�ʱ��ʼ������
    '����:���˺�
    '����:2019-12-02 11:37:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean
    Dim blnʵ����֤ As Boolean
    
    If Not mobjPati Is Nothing Then
        If mobjPati.ʵ����֤ Then blnʵ����֤ = mobjPati.ʵ����֤
    End If
    
    Select Case bytEnabledType
    Case 2  '2-����������Ϣ
        blnEdit = False
        txt����.Enabled = blnEdit
        cbo��������.Enabled = blnEdit
        txt����.Enabled = blnEdit
        chk����.Enabled = blnEdit
        chkEndTime.Enabled = blnEdit
        dtpDate.Enabled = blnEdit
        lbl����.Enabled = blnEdit
        tbSendCard.Enabled = blnEdit
    Case Else   '0-��������,1-���ý�����Ϣ
        blnEdit = mlngCardTypeID <> 0 And mint����״̬ <> 2
        txt����.Enabled = blnEdit
        blnEdit = Trim(txt����.Text) <> "" And mlngCardTypeID <> 0 And mint����״̬ <> 2
        
        cbo��������.Enabled = chk����.value = 0 And blnEdit And mint����״̬ <> 1
        txt����.Enabled = blnEdit
        dtpDate.Enabled = chkEndTime.value = 1 And (blnʵ����֤ Or mobjPubPatient.blnRealName = False)
        chkEndTime.Enabled = blnEdit And (blnʵ����֤ Or mobjPubPatient.blnRealName = False)
        chk����.Enabled = blnEdit
        lbl����.Enabled = blnEdit
        
        If bytEnabledType = 1 Then
            lbl���㷽ʽ.Enabled = False
            cbo��������.Enabled = False
            txt����.Enabled = False
            chk����.Enabled = False
            tbSendCard.Enabled = blnInit
        End If
    End Select
    txtPass.Enabled = blnEdit: txtAudi.Enabled = blnEdit
    
    '������ɫ
    txtPass.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
    txtAudi.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
    txt����.BackColor = IIf(mlngCardTypeID <> 0, &HEBFFFF, &H8000000F)
    txt����.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
    cbo��������.BackColor = IIf(blnEdit, &H80000005, &H8000000F)
    
End Sub
Private Sub SetDepositEditEnabled(Optional bytEnabledType As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ���ı༭����
    '���:bytEnabledType-���õ�����:0-������;1-���ý�����Ϣ��Ԥ����Ϣ;2-����������Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-12-02 11:13:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean
    Dim objPayCard As Card, int���� As Integer
    Set objPayCard = GetDepositPayCard
    
    blnEdit = (Not objPayCard Is Nothing)
    If Not objPayCard Is Nothing Then
        int���� = objPayCard.��������
    End If
    
    blnEdit = blnEdit And fraԤ��.Tag = "" And mint����״̬ <> 2    '��Ԥ����Ϣ
    
    Select Case bytEnabledType
    Case 2  '2-����������Ϣ
        blnEdit = False: txtFact.Enabled = False
    
    Case Else   '0-��������,1-���ý�����Ϣ��Ԥ����Ϣ
         txtFact.Enabled = blnEdit
         If bytEnabledType = 1 Then
            blnEdit = False
         End If
    End Select
    
    txtԤ����.Enabled = blnEdit
    cboԤ������.Enabled = blnEdit And mint����״̬ <> 1
    txt�������.Enabled = blnEdit
    tbDeposit.Enabled = blnEdit
    
    If blnEdit Then blnEdit = int���� <> 3
    If blnEdit Then blnEdit = chk��λ�ɿ�.value = 1
    chk��λ�ɿ�.Enabled = blnEdit: txt������.Enabled = blnEdit
    txt�ʺ�.Enabled = blnEdit
    txt�ɿλ.Enabled = blnEdit And chk��λ�ɿ�.value = 1
    
    '������ɫ
    txtԤ����.BackColor = IIf(txtԤ����.Enabled, &H80000005, &H8000000F)
    txtFact.BackColor = IIf(txtFact.Enabled, &H80000005, &H8000000F)
    cboԤ������.BackColor = IIf(cboԤ������.Enabled, &H80000005, &H8000000F)
    txt�������.BackColor = IIf(txt�������.Enabled, &H80000005, &H8000000F)
    txt������.BackColor = IIf(txt������.Enabled, &H80000005, &H8000000F)
    txt�ʺ�.BackColor = IIf(txt�ʺ�.Enabled, &H80000005, &H8000000F)
    txt�ɿλ.BackColor = IIf(txt�ɿλ.Enabled, &H80000005, &H8000000F)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If fra�ſ�.Visible Then
        zlDatabase.SetPara "����ģʽ", tbSendCard.SelectedItem.Key, glngSys, mlngModule
    End If
    
    Set mCurSendCard.objSendCard = Nothing
    
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    
    Set mbtQRCodePay = Nothing
    Set mobjOneCardComLib = Nothing
    Set mobjPubPatient = Nothing
    Set mobjService = Nothing
    Set mobjExseSvr = Nothing
    Set mobjPati = Nothing
    Set mobjThirdSwap = Nothing
    Set mfrmMain = Nothing
    Set mobjCommEvents = Nothing
    Set mrs���� = Nothing
    Set mobjCardFeePayCards = Nothing
    Set mobjDepositFact = Nothing
    Set mobjDepositPayCards = Nothing
    Set mobjShowTotalMoneyControl = Nothing
    Set mobjCardFeeItems = Nothing
    Set mrsCardFee = Nothing
    Set mobjDepositItems = Nothing
    Set mobjKeyboard = Nothing
    mCurSendCard.lng����ID = 0
    mblnInited = False: mblnNotClick = False: mblnSendCardLocked = False
    mblnDepositLocked = False
    mblnBoundCarded = False: mblnShowDepositAndSendCard = False
    mint����״̬ = 0
End Sub


Public Function GetCardFee(Optional blnReReadCardFee As Boolean = False, Optional ByVal strPriceGrade As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '���:blnRereadCardFee-���¶�ȡ������
    '     strPriceGrade-�۸�ȼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-23 17:31:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objSendCard As Card
    
    On Error GoTo errHandle
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Set GetCardFee = Nothing: Exit Function
    If objSendCard.�ض���Ŀ = "" Then Set GetCardFee = Nothing: Exit Function
    If Not mrs���� Is Nothing Then
        If mrs����.State = 1 Then Set GetCardFee = mrs����: Exit Function
    End If
    
    Set mrs���� = zlGetSpecialItemFee(objSendCard.�ض���Ŀ, strPriceGrade)
    Set GetCardFee = mrs����
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetPublicPatient(ByRef objPubPati_Out As clsInterFacePatient) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����zlPublicPatient����
    '���:
    '����:objPubPati-���ز��˹�������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 10:11:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjPubPatient Is Nothing Then Set objPubPati_Out = mobjPubPatient: GetPublicPatient = True: Exit Function
    On Error GoTo errHandle
    
    Set mobjPubPatient = New clsInterFacePatient
    If mobjPubPatient.Init(Me, glngSys, glngModul, gcnOracle, gstrDBUser) = False Then Exit Function
    Set objPubPati_Out = mobjPubPatient
    GetPublicPatient = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub lbl����_Click()
    Dim strExpand As String, strOutCardNO As String, strPatiInfoXML As String
    Dim objSendCard As Card
    If mblnSendCardLocked Then Exit Sub
    

    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Exit Sub
    
    
    If objSendCard.���� = "���￨" And objSendCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        
        If Not mobjICCard Is Nothing Then
            txt����.Text = mobjICCard.Read_Card()
            If txt����.Text <> "" Then
                mblnICCard = True
                Call CheckFreeCard(txt����.Text)
            End If
        End If
        Exit Sub
    End If
    If (objSendCard.�Ƿ�Ӵ�ʽ���� = False And objSendCard.�Ƿ�ǽӴ�ʽ���� = False) Or objSendCard.�ӿ���� <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strPatiInfoXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strPatiInfoXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\

    If mobjOneCardComLib.zlReadCard(Me, mlngModule, objSendCard.�ӿ����, False, strExpand, strOutCardNO, strPatiInfoXML) = False Then Exit Sub
    txt����.Text = strOutCardNO
    If txt����.Text <> "" Then
        '�����:56599
        If strPatiInfoXML <> "" Then RaiseEvent RequestRefreshPatiInf(strOutCardNO, strPatiInfoXML)
        Call CheckFreeCard(txt����.Text)
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
    Else
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
    End If
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    txt����.Text = strCardNo
    If txt����.Text <> "" Then
        '�����:56599
        If strXmlCardInfor <> "" Then RaiseEvent RequestRefreshPatiInf(strCardNo, strXmlCardInfor)
        Call CheckFreeCard(txt����.Text)
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
    Else
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
    End If
End Sub
 
Private Sub CheckFreeCard(ByVal strCardNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��һ��ͨģʽ�µĿ��ţ��ϸ����Ʊ��ʱ������Ƿ���Ʊ�����÷�Χ�ڣ���Χ֮��Ŀ����շ�
    '���:strCardNo-����
    '����:2019-11-25 12:01:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs���� As ADODB.Recordset
    Dim str�ѱ� As String, dblMoney As Double
    Dim objSendCard As Card
    
    If txt����.Visible = False Then Exit Sub
    
    Set rs���� = GetCardFee()
    If Not mobjPati Is Nothing Then str�ѱ� = mobjPati.�ѱ�
    
    If Not rs���� Is Nothing And Val(txt����.Text) = 0 Then  '�Ȼָ�
        txt����.Text = Format(IIf(rs����!�Ƿ��� = 1, rs����!ȱʡ�۸�, rs����!�ּ�), "0.00")
        txt����.Tag = txt����.Text
    End If
    
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Exit Sub
    
    If mCurSendCard.blnOneCard And objSendCard.�Ƿ��ϸ���� Then
        mCurSendCard.lng����ID = mobjExseSvr.CheckUsedBill(5, IIf(mCurSendCard.lng����ID > 0, mCurSendCard.lng����ID, mCurSendCard.lng��������), strCardNo)
        If mCurSendCard.lng����ID <= 0 Then txt����.Text = "0.00": txt����.Tag = txt����.Text
    End If

    If Not rs���� Is Nothing And Val(txt����.Text) <> 0 Then
        If rs����!�Ƿ��� = 0 Then
            If mobjExseSvr.zl_ExseSvr_Actualmoney(str�ѱ�, rs����!�շ�ϸĿID, rs����!������ĿID, rs����!�ּ�, dblMoney) Then
                txt����.Text = Format(dblMoney, "0.00")
                txt����.Tag = txt����.Text
           End If
        End If
    End If
End Sub


Private Function GetDepositBalanceItems(ByVal dtCurdate As Date, ByRef objBalanceItems_Out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰԤ��������Ϣ
    '���:dtCurDate-��ǰʱ��
    '����:objBalanceItems_Out-Ԥ���Ľ������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 20:16:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, strDepositNo As String
    Dim objCurItem As clsBalanceItem
    Dim lngԤ��ID As Long
    On Error GoTo errHandle
    
    
    Set objBalanceItems_Out = New clsBalanceItems
    If fraԤ��.Visible = False Or StrToNum(txtԤ����.Text) = 0 Then GetDepositBalanceItems = True: Exit Function
    
    
    Set objCard = GetDepositPayCard()
    If objCard Is Nothing Then Exit Function
    Set objCurItem = New clsBalanceItem
    
    If mobjExseSvr.zl_ExseSvr_GetNextNo(11, strDepositNo) = False Then Exit Function   'Ԥ��No
    If mobjExseSvr.zl_ExseSvr_GetNextID("����Ԥ����¼", lngԤ��ID) = False Then Exit Function
     
    With objCurItem
        Set .objCard = objCard
        .�����ID = IIf(objCard.�ӿ���� < 0, 0, objCard.�ӿ����)
        .���ѿ� = objCard.���ѿ�
        .���㷽ʽ = objCard.���㷽ʽ
        .������� = Trim(txt�������.Text)
        .������ = StrToNum(txtԤ����.Text)
        .�������� = objCard.��������
        If .�����ID > 0 Then
           .�������� = IIf(.���ѿ�, 5, 3)
        ElseIf objCard.�������� = 3 Then
             .�������� = 2
        Else
           .�������� = 0  '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        End If
        .����ʱ�� = dtCurdate
        .����ժҪ = ""
        .���ݺ� = strDepositNo
        .Ԥ��ID = lngԤ��ID
        .�Ƿ�Ԥ�� = True
    End With
    objBalanceItems_Out.AddItem objCurItem
    objBalanceItems_Out.������ = objCurItem.������
    objBalanceItems_Out.���ݺ� = objCurItem.���ݺ�
    objBalanceItems_Out.���� = objCurItem.��������
    objBalanceItems_Out.����ʱ�� = Format(dtCurdate, "yyyy-mm-dd HH:MM:SS")
    
    GetDepositBalanceItems = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetCardFeeBalanceItems(ByVal dtCurdate As Date, ByRef objBalanceItems_Out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ���ѽ�����Ϣ
    '���:dtCurDate-��ǰʱ��
    '����:objBalanceItems_Out-���ѵĽ������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 20:16:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, strCardFeeNo As String
    Dim objCurItem As clsBalanceItem, lng����ID As Long
    On Error GoTo errHandle
    
    Set objBalanceItems_Out = New clsBalanceItems
    If fra�ſ�.Visible = False Or tbSendCard.SelectedItem.Key <> "CardFee" Or txt����.Text = "" Then GetCardFeeBalanceItems = True: Exit Function
    
    
    If chk����.value = 1 Then
        If mobjExseSvr.zl_ExseSvr_GetNextNo(16, strCardFeeNo) = False Then Exit Function     'ҽ�ƿ����ݺ�
        objBalanceItems_Out.������ = StrToNum(txt����.Text)
        objBalanceItems_Out.���ݺ� = strCardFeeNo
        objBalanceItems_Out.���� = gEM_���ʵ�
        GetCardFeeBalanceItems = True
        Exit Function
    End If
    
    Set objCard = GetCardFeePayCard()
    If objCard Is Nothing Then Exit Function
    Set objCurItem = New clsBalanceItem
    
    If mobjExseSvr.zl_ExseSvr_GetNextNo(16, strCardFeeNo) = False Then Exit Function    'ҽ�ƿ����ݺ�
    If mobjExseSvr.zl_ExseSvr_GetNextID("���˽��ʼ�¼", lng����ID) = False Then Exit Function
    
    With objCurItem
        Set .objCard = objCard
        .�����ID = IIf(objCard.�ӿ���� < 0, 0, objCard.�ӿ����)
        .���ѿ� = objCard.���ѿ�
        .���㷽ʽ = objCard.���㷽ʽ
        .������� = ""
        .������ = StrToNum(txt����.Text)
        .�������� = objCard.��������
        If .�����ID > 0 Then
           .�������� = IIf(.���ѿ�, 5, 3)
        Else
           .�������� = 0  '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        End If
        .����ʱ�� = dtCurdate
        .����ժҪ = ""
        .����ID = lng����ID
        .���ݺ� = strCardFeeNo
        .�Ƿ�Ԥ�� = False
        
    End With
    objBalanceItems_Out.AddItem objCurItem
    objBalanceItems_Out.������ = objCurItem.������
    objBalanceItems_Out.���ݺ� = objCurItem.���ݺ�
    objBalanceItems_Out.���� = objCurItem.��������
    objBalanceItems_Out.����ʱ�� = Format(dtCurdate, "yyyy-mm-dd HH:MM:SS")
    
    GetCardFeeBalanceItems = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CheckDepsoitAndCardFeePayIsSame(ByVal objDepositItems As clsBalanceItems, ByVal objCardFeeItems As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ҽ�ƿ���Ԥ�����Ƿ�ͬһ��֧����ʽ
    '���:objDepositItems-��ǰ��Ԥ������
    '     objCardFeeItems-��ǰ�Ŀ��ѽ���
    '����:
    '����:ͬһ�ַ���true,���򷵻�False
    '����:���˺�
    '����:2019-11-26 16:02:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If objDepositItems Is Nothing Or objCardFeeItems Is Nothing Then Exit Function
    If objDepositItems.Count <> objCardFeeItems.Count Then Exit Function
    If objDepositItems.Count = 0 Or objCardFeeItems.Count = 0 Then Exit Function
    
    If objDepositItems(1).�����ID = objCardFeeItems(1).�����ID And objDepositItems(1).���ѿ� = objCardFeeItems(1).���ѿ� Then
        CheckDepsoitAndCardFeePayIsSame = True
    End If
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetErrDataToColl(ByVal objPati As clsPatientInfo, ByVal lngҵ��ID As Long, _
    ByVal objCurItems As clsBalanceItems, ByVal intͬ����־ As Integer, ByRef lng�쳣id_Out As Long, ByVal dtCurdate As Date, _
    ByRef cllErrData_out As Collection, Optional ByVal strԤ������ As String, Optional ByVal dblԤ����� As Double, _
    Optional ByVal str���ѵ��� As String, Optional dbl���� As Double, Optional blnCancel As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�쳣���շ���������
    '���:int����״̬-:0-������¼,1-����״̬�����½���˵����2-ɾ���쳣����
    '     objCurItem-��ǰ������Ϣ
    '     lng�쳣id-�쳣ID
    '     dtCurDate-��ǰ����
    '     intͬ����־:   0-������¼;-1-δ��������;1-δ���ýӿ�;2-�ӿڵ��óɹ�,4-ҽ�ƿ���Ϣ���³ɹ�;
    '     blnCancel-�Ƿ�����
    '����:lng�쳣id_Out-�쳣ID
    '       cllErrData_Out-���ش�����Ϣ��(��ʽΪArray(����������,������ֵ)
    '          �������Ŀ���ư���: �쳣ID,��������,���ϱ�־,ҵ��id,�Ƿ�����,����id,��ҳid,����,�Ա�,����,�����,סԺ��,Ԥ������,Ԥ�����,ҽ�ƿ�����,����,�������id,�����������,��������,ͬ��״̬,������Ϣ)
    '          ���н�����ϢΪJson������ʽ����
    '           {"card_no":"00002","cardtype_id":23,"swapno":"J2223432","swapmoney":324,"otherswap_list":[{"swap_name":"POSM","swap_note":"A001"},{}]})
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-26 16:16:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int�������� As Integer, i As Long
    Dim objCurItem As clsBalanceItem
    
    On Error GoTo errHandle
    lng�쳣id_Out = zlDatabase.GetNextId("���˽����쳣��¼")
    '1-ҽ�ƿ�����;2-������Ϣ�Ǽ�;3-������Ժ�Ǽ�;4-ԤԼ�ҺŽ���
    int�������� = IIf(mlngModule = 1101, 2, 3)
    Set cllErrData_out = New Collection
    cllErrData_out.Add Array("�쳣ID", lng�쳣id_Out)
    cllErrData_out.Add Array("��������", int��������)
    cllErrData_out.Add Array("���ϱ�־", IIf(blnCancel, 1, 0))
    cllErrData_out.Add Array("ҵ��ID", lngҵ��ID)
    cllErrData_out.Add Array("�Ƿ�����", 0)
    cllErrData_out.Add Array("����ID", objPati.����ID)
    cllErrData_out.Add Array("��ҳID", objPati.��ҳID)
    
    cllErrData_out.Add Array("����", objPati.����)
    cllErrData_out.Add Array("�Ա�", objPati.�Ա�)
    cllErrData_out.Add Array("����", objPati.����)
    cllErrData_out.Add Array("�����", objPati.�����)
    cllErrData_out.Add Array("סԺ��", objPati.סԺ��)
    
    cllErrData_out.Add Array("Ԥ������", strԤ������)
    cllErrData_out.Add Array("Ԥ�����", dblԤ�����)
    cllErrData_out.Add Array("ҽ�ƿ�����", str���ѵ���)
    cllErrData_out.Add Array("����", dbl����)
    cllErrData_out.Add Array("����Ա����", UserInfo.����)
    cllErrData_out.Add Array("����Ա���", UserInfo.���)
    cllErrData_out.Add Array("�Ǽ�ʱ��", Format(dtCurdate, "yyyy-mm-dd HH:MM:SS"))
    
    
    
    If str���ѵ��� <> "" Then
        cllErrData_out.Add Array("�������ID", mCurSendCard.objSendCard.�ӿ����)
        cllErrData_out.Add Array("�����������", mCurSendCard.objSendCard.����)
        cllErrData_out.Add Array("��������", txt����.Text)
    End If
    cllErrData_out.Add Array("ͬ��״̬", intͬ����־)
    Dim strJson As String
    
    strJson = ""
    If Not objCurItems Is Nothing Then
        objCurItems.�쳣ID = lng�쳣id_Out
        objCurItems.ҵ��ID = lngҵ��ID
        If objCurItems.Count <> 0 Then
            For i = 1 To objCurItems.Count
                objCurItems(i).�쳣ID = lng�쳣id_Out
            Next
            Set objCurItem = objCurItems(1)
            strJson = strJson & "" & GetJsonNodeString("card_no", objCurItem.����, Json_Text)
            strJson = strJson & "," & GetJsonNodeString("cardtype_id", objCurItem.�����ID, Json_num)
            strJson = "{" & strJson & "}"
        End If
    End If
    cllErrData_out.Add Array("������Ϣ", strJson)
    zlGetErrDataToColl = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function GetDelErrDataToColl(ByVal lngҵ��ID As Long, lng�쳣ID As Long, ByRef cllErrData_out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�쳣���շ���������
    '���: lng�쳣ID-�쳣ID
    '     lng�쳣id-�쳣ID
    '����:lng�쳣id_Out-�쳣ID
    '       cllErrData_Out-���ش�����Ϣ��(��ʽΪArray(����������,������ֵ)
    '          �������Ŀ���ư���: �쳣ID,��������,���ϱ�־,ҵ��id,�Ƿ�����,����id,��ҳid,����,�Ա�,����,�����,סԺ��,Ԥ������,Ԥ�����,ҽ�ƿ�����,����,�������id,�����������,��������,ͬ��״̬,������Ϣ)
    '          ���н�����ϢΪJson������ʽ����
    '           {"card_no":"00002","cardtype_id":23,"swapno":"J2223432","swapmoney":324,"otherswap_list":[{"swap_name":"POSM","swap_note":"A001"},{}]})
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-26 16:16:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int�������� As Integer, intͬ����־ As Integer
    Dim objCurItem As clsBalanceItem
    On Error GoTo errHandle
   
    '1-ҽ�ƿ�����;2-������Ϣ�Ǽ�;3-������Ժ �Ǽ�;4-ԤԼ�ҺŽ���
    int�������� = IIf(mlngModule = 1101, 2, 3)
    Set cllErrData_out = New Collection
    cllErrData_out.Add Array("�쳣ID", lng�쳣ID)
    cllErrData_out.Add Array("��������", int��������)
    cllErrData_out.Add Array("ҵ��ID", lngҵ��ID)
    GetDelErrDataToColl = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
 Private Function GetUpdateErrDataSyncTagToColl(lng�쳣ID As Long, ByVal intͬ����־ As Integer, ByRef cllErrData_out As Collection, Optional ByVal cllSendCard As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�쳣���շ���������
    '���: lng�쳣ID-�쳣ID
    '     lng�쳣id-�쳣ID
    '����:lng�쳣id_Out-�쳣ID
    '       cllErrData_Out-���ش�����Ϣ��(��ʽΪArray(����������,������ֵ)
    '          �������Ŀ���ư���: �쳣ID,��������,���ϱ�־,ҵ��id,�Ƿ�����,����id,��ҳid,����,�Ա�,����,�����,סԺ��,Ԥ������,Ԥ�����,ҽ�ƿ�����,����,�������id,�����������,��������,ͬ��״̬,������Ϣ)
    '          ���н�����ϢΪJson������ʽ����
    '           {"card_no":"00002","cardtype_id":23,"swapno":"J2223432","swapmoney":324,"otherswap_list":[{"swap_name":"POSM","swap_note":"A001"},{}]})
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-26 16:16:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int�������� As Integer
    Dim objCurItem As clsBalanceItem
    Dim varData As Variant
    Dim i As Long
    On Error GoTo errHandle
    Set cllErrData_out = New Collection
    cllErrData_out.Add Array("�쳣ID", lng�쳣ID)
    cllErrData_out.Add Array("ͬ��״̬", intͬ����־)
    If Not cllSendCard Is Nothing Then
        '��Ҫ���¿�����Ϣ
        For i = 1 To cllSendCard.Count
            varData = cllSendCard(i)
            Select Case varData(0)
            Case "ҽ�ƿ���"
                cllErrData_out.Add Array("��������", varData(1))
            Case "�����ID"
                cllErrData_out.Add Array("�������ID", varData(1))
            End Select
        Next
        
    End If
    GetUpdateErrDataSyncTagToColl = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function UpdateCardFeeBalanceInfor(ByVal int����״̬ As Integer, ByVal objPati As clsPatientInfo, _
    ByVal cllSendCardInfo As Collection, ByVal objCardFeeItems As clsBalanceItems, ByVal objDepositItems As clsBalanceItems, _
    ByVal cllExpendInfo As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¿�����ؽ�����Ϣ
    '���:int����״̬:0-��ɽ���;1-�ӿڵ���ǰ����;2-�ӿڵ��ú�����
    '     objCardFeeItems-��ǰ���ѽ���֧����Ϣ
    '     objDepositItems-��ǰԤ��֧����ʽ
    '     cllSendCardInfo-������Ϣ (�����ID,�䶯����,����,ԭ����,IC����,����,��������,��ֹʹ��ʱ��,����,������,ժҪ,��������,����ID),��ʽ:array(����,ֵ),"_����"
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-14 11:49:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllUpdateFeeData As Collection, cllTemp As Collection
    Dim objCurItem As clsBalanceItem, blnTrans As Boolean
    Dim strDepositNo As String, strCardFeeNo As String, strErrMsg As String
    Dim varTemp As Variant, lngԤ��ID As Long
    Dim cllPro As Collection, strSql As String, int�쳣״̬ As Integer
    Dim cllErrData As Collection, dtCurdate As Date
    
    On Error GoTo errHandle
    
    Set cllUpdateFeeData = New Collection
    Set cllTemp = New Collection
    
    If objCardFeeItems.���� <> gEM_���ʵ� Then
        
        Set objCurItem = objCardFeeItems(1)
        If Not objDepositItems Is Nothing Then
            If objDepositItems.Count <> 0 Then
                strDepositNo = objDepositItems.���ݺ�
                lngԤ��ID = objDepositItems(1).Ԥ��ID
            End If
        End If
    End If
    If Not objDepositItems Is Nothing Then
        strDepositNo = objDepositItems.���ݺ�
    End If
    
    strCardFeeNo = objCardFeeItems.���ݺ�
    cllTemp.Add Array("Ԥ������", strDepositNo), "_" & "Ԥ������"
    cllTemp.Add Array("Ԥ��ID", lngԤ��ID), "_" & "Ԥ��ID"
    cllTemp.Add Array("�շѵ���", strCardFeeNo), "_" & "�շѵ���"
    If Not objCurItem Is Nothing Then
        cllTemp.Add Array("����ID", IIf(objCurItem.����ID <> 0, objCurItem.����ID, objCurItem.����ID)), "_" & "����ID"
    End If
    cllTemp.Add Array("����ID", objPati.����ID), "_" & "����ID"
    cllTemp.Add Array("����Ա���", UserInfo.���), "_" & "����Ա���"
    cllTemp.Add Array("����Ա����", UserInfo.����), "_" & "����Ա����"
     
    If Not objCurItem Is Nothing Then
        If Val(objCurItem.����ʱ��) = 0 Then
            dtCurdate = zlDatabase.Currentdate
            cllTemp.Add Array("�տ�ʱ��", Format(dtCurdate, "yyyy-mm-dd HH:MM:SS")), "_" & "�տ�ʱ��"
        Else
            cllTemp.Add Array("�տ�ʱ��", Format(objCurItem.����ʱ��, "yyyy-mm-dd HH:MM:SS")), "_" & "�տ�ʱ��"
        End If
    End If
    cllUpdateFeeData.Add cllTemp, "_billinfo"
    
    If Not objCurItem Is Nothing Then
         '������Ϣ
        Set cllTemp = New Collection
        cllTemp.Add Array("���㷽ʽ", objCurItem.���㷽ʽ), "_" & "���㷽ʽ"
        cllTemp.Add Array("�������", objCurItem.�������), "_" & "�������"
        cllTemp.Add Array("�����ID", IIf(objCurItem.���ѿ�, 0, objCurItem.�����ID)), "_" & "�����ID"
        cllTemp.Add Array("���㿨���", IIf(objCurItem.���ѿ�, objCurItem.�����ID, 0)), "_" & "���㿨���"
        cllTemp.Add Array("����", objCurItem.����), "_" & "����"
        cllTemp.Add Array("������ˮ��", objCurItem.������ˮ��), "_" & "������ˮ��"
        cllTemp.Add Array("����˵��", objCurItem.����˵��), "_" & "����˵��"
        cllTemp.Add Array("ժҪ", objCurItem.����ժҪ), "_" & "ժҪ"
        cllTemp.Add Array("������λ", ""), "_" & "������λ"
        
        If Not cllExpendInfo Is Nothing Then
            cllTemp.Add Array("������Ϣ��", cllExpendInfo), "_" & "������Ϣ��"
        End If
        cllUpdateFeeData.Add cllTemp, "_balanceinfo"
    End If
    ' cllUpdateDate-�޸ĵĽ�������
    '         |--billinfo-������Ϣ,"_billinfo"
    '              |-Ԥ������,Ԥ��ID,�շѵ���,����ID,����Ա���,����Ա����,�տ�ʱ��)
    '         |--balanceinfo-������Ϣ,"_balanceinfo"
    '                |--(���㷽ʽ,�������,�����id,���㿨���,����,������ˮ��,����˵��,ժҪ,������λ)
    '                |--������Ϣ��,
    '                |-----������Ϣ:��������,��������
    
    'ͬ��״̬����������=2,3ʱ��0��NULL������¼;-1-δ��������;1-δ���ýӿ�;2-�ӿڵ��óɹ�,3-���ý��������ɹ�;4-ҽ�ƿ���Ϣ�����ɹ�"
    If int����״̬ = 0 Then
         int�쳣״̬ = 2
        If Not GetDelErrDataToColl(objCardFeeItems.ҵ��ID, objCardFeeItems.�쳣ID, cllErrData) Then Exit Function
    ElseIf int����״̬ = 1 Then
        If Not GetUpdateErrDataSyncTagToColl(objCardFeeItems.�쳣ID, 1, cllErrData) Then Exit Function
        int�쳣״̬ = 1
    Else
        If Not GetUpdateErrDataSyncTagToColl(objCardFeeItems.�쳣ID, 3, cllErrData) Then Exit Function
        int�쳣״̬ = 1
        '0-������¼,1-����״̬�����½���˵����2-ɾ���쳣����
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    If Zl_���˽����쳣��¼_Modify(int�쳣״̬, cllErrData) = False Then
        gcnOracle.RollbackTrans: blnTrans = False
        Exit Function
    End If
    
    If mobjExseSvr.Zl_Exsesvr_UpdCardFeeBlncInfo(int����״̬, cllSendCardInfo, cllUpdateFeeData, False, strErrMsg) = False Then
        gcnOracle.RollbackTrans: blnTrans = False
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    UpdateCardFeeBalanceInfor = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

'Public Function Excute_DepositSaveOver(ByVal objPati As clsPatientInfo, ByVal objBalanceItems As clsBalanceItems, _
'    ByVal cllExpendInfo As Collection) As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:���Ԥ������
'    '���:objBalanceItems-��ǰ֧����
'    '����:�ɹ�����true,���򷵻�False
'    '����:���˺�
'    '����:2019-11-14 11:49:36
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cllUpdateFeeData As Collection, cllTemp As Collection
'    Dim objCurItem As clsBalanceItem, blnTrans As Boolean
'    Dim strDepositNo As String, strCardFeeNo As String, strErrMsg As String
'    Dim varTemp As Variant, lng�䶯ID As Long
'    Dim cllPro As Collection, strSQL As String
'    On Error GoTo errHandle
'
'    Set cllUpdateFeeData = New Collection
'    Set cllTemp = New Collection
'
'
'    Set objCurItem = objBalanceItems(1)
'
'    strCardFeeNo = ""
'
'    strDepositNo = objCurItem.���ݺ�
'
'    cllTemp.Add Array("Ԥ������", strDepositNo), "_" & "Ԥ������"
'    cllTemp.Add Array("Ԥ��ID", objCurItem.Ԥ��ID), "_" & "Ԥ��ID"
'    cllTemp.Add Array("����ID", objPati.����ID), "_" & "����ID"
'    cllTemp.Add Array("����Ա���", UserInfo.���), "_" & "����Ա���"
'    cllTemp.Add Array("����Ա����", UserInfo.����), "_" & "����Ա����"
'    cllTemp.Add Array("�տ�ʱ��", Format(objCurItem.����ʱ��, "yyyy-mm-dd HH:MM:SS")), "_" & "�տ�ʱ��"
'    cllUpdateFeeData.Add cllTemp, "_billinfo"
'
'     '������Ϣ
'    Set cllTemp = New Collection
'    Set objCurItem = objBalanceItems(1)
'    cllTemp.Add Array("���㷽ʽ", objCurItem.���㷽ʽ), "_" & "���㷽ʽ"
'    cllTemp.Add Array("�������", objCurItem.�������), "_" & "�������"
'    cllTemp.Add Array("�����ID", IIf(objCurItem.���ѿ�, 0, objCurItem.�����ID)), "_" & "�����ID"
'    cllTemp.Add Array("���㿨���", IIf(objCurItem.���ѿ�, objCurItem.�����ID, 0)), "_" & "���㿨���"
'    cllTemp.Add Array("����", objCurItem.����), "_" & "����"
'    cllTemp.Add Array("������ˮ��", objCurItem.������ˮ��), "_" & "������ˮ��"
'    cllTemp.Add Array("����˵��", objCurItem.����˵��), "_" & "����˵��"
'    cllTemp.Add Array("ժҪ", objCurItem.����ժҪ), "_" & "ժҪ"
'    cllTemp.Add Array("������λ", ""), "_" & "������λ"
'
'    If Not cllExpendInfo Is Nothing Then
'        cllTemp.Add Array("������Ϣ��", cllExpendInfo), "_" & "������Ϣ��"
'    End If
'    cllUpdateFeeData.Add cllTemp, "_balanceinfo"
'    ' cllUpdateDate-�޸ĵĽ�������
'    '         |--billinfo-������Ϣ,"_billinfo"
'    '              |-Ԥ������,Ԥ��ID,�շѵ���,����ID,����Ա���,����Ա����,�տ�ʱ��)
'    '         |--balanceinfo-������Ϣ,"_balanceinfo"
'    '                |--(���㷽ʽ,�������,�����id,���㿨���,����,������ˮ��,����˵��,ժҪ,������λ)
'    '                |--������Ϣ��,
'    '                |-----������Ϣ:��������,��������
'
'    blnTrans = True
'    Set cllTemp = New Collection
'    cllTemp.Add Array("�쳣ID", objBalanceItems.�쳣ID), "_�쳣ID"
'
'    gcnOracle.BeginTrans: blnTrans = True
'    If Zl_���˽����쳣��¼_Modify(2, cllTemp) = False Then
'        gcnOracle.RollbackTrans: blnTrans = True
'        Exit Function
'    End If
'
'    If mobjExseSvr.Zl_Exsesvr_Upddepositblncinfo(0, cllUpdateFeeData, False, strErrMsg) = False Then
'       gcnOracle.RollbackTrans: blnTrans = False
'       MsgBox strErrMsg, vbInformation, gstrSysName
'       Exit Function
'    End If
'    gcnOracle.CommitTrans: blnTrans = False
'
'    Excute_DepositSaveOver = True
'    Exit Function
'errHandle:
'    If blnTrans Then gcnOracle.RollbackTrans
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function


Public Function Excute_Cancel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���쳣���ϲ���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-12-02 15:16:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objDepositItems As clsBalanceItems, objCardFeeItems As clsBalanceItems, objTempItems As clsBalanceItems
    Dim bln��ͬ As Boolean, objCard As Card, strErrMsg As String, intSwapStatu As Integer, cllErrData As Collection, cllAddErrData As Collection
    Dim blnTrans As Boolean, dtCurdate As Date, cllDelFeeData As Collection
    Dim lng�쳣ID As Long, cllSendCardInfo As Collection, lng����ID As Long, lngԤ��ID As Long
    
    On Error GoTo errHandle
    
    dtCurdate = zlDatabase.Currentdate
    
    Set objDepositItems = New clsBalanceItems
    If Not mobjDepositItems Is Nothing Then
        If mobjDepositItems.Count <> 0 Then Set objDepositItems = mobjDepositItems.Clone
    End If
    
    Set objCardFeeItems = New clsBalanceItems
    If Not mobjCardFeeItems Is Nothing Then
       Set objCardFeeItems = mobjCardFeeItems.Clone
       
        If GetSaveSendCardInfotoCollect(mobjPati, dtCurdate, cllSendCardInfo) = False Then Exit Function
    End If
    
    bln��ͬ = CheckDepsoitAndCardFeePayIsSame(objDepositItems, objCardFeeItems)
    
    '������Ԥ��
    If objDepositItems.Count <> 0 And Not bln��ͬ And mobjCardFeeItems Is Nothing Then
        Set objCard = objDepositItems(1).objCard
        If objDepositItems.ͬ��״̬ >= 2 Then
            MsgBox objCard.���� & "�Ѿ�������ɣ����ܽ������ϲ���,��������쳣���ա�����!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        If objDepositItems.ͬ��״̬ = 1 And objDepositItems.���� = gEM_һ��ͨ Then
        
            Set mobjThirdSwap.objPayCards = mobjDepositPayCards
            If mobjThirdSwap.zlThird_IsSwapIsSucces(objDepositItems, intSwapStatu, strErrMsg) = False Then
                '����ʧ��
                'intSwapStatu_Out-�ӿڷ���Falseʱ���˲�����Ч:����״̬: 0-���׵���ʧ��;1-�������ڴ�����
                If intSwapStatu = 1 Then
                    MsgBox "ԭ" & objCard.���� & " �������ڽ����У� ���������ϲ���!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            Else
                MsgBox "ԭ" & objCard.���� & " �����Ѿ��ɹ��� ���������ϲ���!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
     
        End If
        
        If Not GetDelErrDataToColl(objDepositItems.ҵ��ID, objDepositItems.�쳣ID, cllErrData) Then Exit Function
        gcnOracle.BeginTrans: blnTrans = True
        If Zl_���˽����쳣��¼_Modify(2, cllErrData) = False Then
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        If mobjExseSvr.Zl_Exsesvr_DelDepositErrorRec(0, objDepositItems.���ݺ�, False) = False Then
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        Excute_Cancel = True: Exit Function
    End If
    
    If Not mobjCardFeeItems Is Nothing Then
         If mobjCardFeeItems.Count = 0 Then
            If mobjCardFeeItems.���� <> gEM_���ʵ� Then Exit Function 'δ�ҵ��쳣����
        End If
    
        If mobjCardFeeItems.ͬ��״̬ = 1 And mobjCardFeeItems.���� = gEM_һ��ͨ Then
        
            Set mobjThirdSwap.objPayCards = mobjCardFeePayCards
            Set objCard = mobjCardFeeItems(1).objCard
            
            If mobjThirdSwap.zlThird_IsSwapIsSucces(mobjCardFeeItems, intSwapStatu, strErrMsg) = False Then
                '����ʧ��
                'intSwapStatu_Out-�ӿڷ���Falseʱ���˲�����Ч:����״̬: 0-���׵���ʧ��;1-�������ڴ�����
                If intSwapStatu = 1 Then
                    MsgBox "ԭ" & objCard.���� & " �������ڽ����У� ���������ϲ���!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            Else
                MsgBox "ԭ" & objCard.���� & " �����Ѿ��ɹ��� ���������ϲ���!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
        
        If mobjCardFeeItems.ͬ��״̬ = -1 Then
            'δ���ɷ��û�Ԥ����ֱ��ɾ��
            If Not GetDelErrDataToColl(objCardFeeItems.ҵ��ID, objCardFeeItems.�쳣ID, cllErrData) Then Exit Function
            gcnOracle.BeginTrans: blnTrans = True
            If Zl_���˽����쳣��¼_Modify(2, cllErrData) = False Then
                gcnOracle.RollbackTrans: blnTrans = False: Exit Function
            End If
            
            'ɾ���䶯��¼
            If mobjService.zl_PatiSvr_DelCardChangeInfo(mobjPati.����ID, objCardFeeItems.ҵ��ID, Val(cllSendCardInfo("_�����ID")(1)), CStr(cllSendCardInfo("_ҽ�ƿ���")(1))) = False Then
              gcnOracle.RollbackTrans: blnTrans = False: Exit Function
            End If
            gcnOracle.CommitTrans: blnTrans = False
            Excute_Cancel = True
            Exit Function
        End If
        
        
        
        If Not GetDelErrDataToColl(objCardFeeItems.ҵ��ID, objCardFeeItems.�쳣ID, cllErrData) Then Exit Function
        If zlGetErrDataToColl(mobjPati, objCardFeeItems.ҵ��ID, objCardFeeItems, 1, lng�쳣ID, dtCurdate, cllAddErrData, "", 0, objCardFeeItems.���ݺ�, objCardFeeItems.������, True) = False Then Exit Function
        '      cllDelFeeData-�˷�����
        '        |-(���ѵ���,Ԥ������,�Ƿ��˿���,�Ƿ��˲�����,����Ա����,����Ա���,�˷�ʱ��,������Ϣ) array(����,ֵ) ,"_����)
        '        |-������Ϣ:(�˿���,���㷽ʽ,�������,�����id,���㿨���,֧������,������ˮ��,����˵��,������λ,��������ID) Key="_������Ϣ"
        
        Set cllDelFeeData = New Collection
        cllDelFeeData.Add Array("���ѵ���", objCardFeeItems.���ݺ�)
        If mobjDepositItems Is Nothing Then
            cllDelFeeData.Add Array("Ԥ������", "")
        Else
            cllDelFeeData.Add Array("Ԥ������", mobjDepositItems.���ݺ�)
        End If
        cllDelFeeData.Add Array("�Ƿ��˲�����", 1)
        cllDelFeeData.Add Array("�Ƿ��˿���", 1)
        cllDelFeeData.Add Array("����Ա����", UserInfo.����)
        cllDelFeeData.Add Array("����Ա���", UserInfo.���)
        cllDelFeeData.Add Array("�˷�ʱ��", Format(dtCurdate, "yyyy-mm-dd HH:MM:SS"))
        gcnOracle.BeginTrans: blnTrans = True
        '1.��ɾ��ԭ�쳣
         If Zl_���˽����쳣��¼_Modify(2, cllErrData) = False Then
           gcnOracle.RollbackTrans: blnTrans = False: Exit Function
         End If
        '2.���������쳣
        If Zl_���˽����쳣��¼_Modify(0, cllAddErrData) = False Then
           gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        
        '3.ɾ�����ü�Ԥ��
        If mobjExseSvr.Zl_Exsesvr_DelCardfeeInfo(2, cllDelFeeData, lng����ID, lngԤ��ID) = False Then
           gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        '4-ɾ��ҽ�ƿ��䶯��¼
        
        If Not GetDelErrDataToColl(objCardFeeItems.ҵ��ID, lng�쳣ID, cllErrData) Then Exit Function
        gcnOracle.BeginTrans: blnTrans = True
        If Zl_���˽����쳣��¼_Modify(2, cllErrData) = False Then
             gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        
        'ɾ���䶯��¼
        If mobjService.zl_PatiSvr_DelCardChangeInfo(mobjPati.����ID, objCardFeeItems.ҵ��ID, Val(cllSendCardInfo("_�����ID")(1)), CStr(cllSendCardInfo("_ҽ�ƿ���")(1))) = False Then
           gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        Excute_Cancel = True
        Exit Function
      
    End If
    Excute_Cancel = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlSaveData(ByVal blnNewPati As Boolean, ByVal objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���:objPati-������Ϣ��
    '     blnNewPati-�Ƿ��²���
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 13:18:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objDepositPrint As Boolean '�Ƿ�Ԥ����ӡ
    Dim objDepositItems As clsBalanceItems, objCardFeeItems As clsBalanceItems, objTempItems As clsBalanceItems
    Dim dtCurdate As Date, cllSaveSendCardInfo As Collection, cllDepositAndCardFee As Collection, cllSendCardInfo As Collection
    Dim cllErrData As Collection, cllPro As Collection
    Dim lng�䶯id As Long, lng�쳣ID As Long, lngԤ��ID As Long, i As Long, lng����ID As Long
    Dim objCurItem As clsBalanceItem, objItems As clsBalanceItems
    Dim blnTrans As Boolean, dbl�ʻ���� As Double, intͬ��״̬ As Integer, int�쳣����״̬ As Integer
    Dim rsMoney As ADODB.Recordset
    Dim rsExpend As ADODB.Recordset, cllExpend As Collection
    Dim int״̬ As Integer, blnSaveed As Boolean
    Dim int�䶯����   As Integer
    Dim strDepositNo As String, intSwapStatu As Integer, strErrMsg As String
    
    
    On Error GoTo errHandle
    
    If Trim(txtԤ����.Text) = "" And Trim(txt����.Text) = "" Then zlSaveData = True: Exit Function
    If mint����״̬ = 2 Then
        '�쳣���ϲ���
        zlSaveData = Excute_Cancel
        Exit Function
    End If
    
    dtCurdate = zlDatabase.Currentdate
    If fra�ſ�.Visible And Not mblnBoundCarded And mlngCardTypeID <> 0 And Trim(txt����.Text) <> "" Then
         int�䶯���� = GetCurCard_Statu
         If GetSaveSendCardInfotoCollect(objPati, dtCurdate, cllSendCardInfo) = False Then Exit Function
         If int�䶯���� = 11 Then '�󶨿�������������ֱ���˳�
             If mobjService.zlPatisvr_SaveMedcCard(cllSendCardInfo, , True) = False Then Exit Function
             mblnBoundCarded = True
             If fraԤ��.Visible = False Or StrToNum(txtԤ����.Text) = 0 Then zlSaveData = True: Exit Function
         End If
    End If
   
    '��ȡԤ�����ݼ����ѵ��ݽ�����Ϣ
    '������֯
    If mobjDepositItems Is Nothing Then
        If GetDepositBalanceItems(dtCurdate, objDepositItems) = False Then Exit Function
    ElseIf mobjDepositItems.Count = 0 Or mobjDepositItems.�Ƿ񱣴� = False Then
        If GetDepositBalanceItems(dtCurdate, objDepositItems) = False Then Exit Function
    Else
        Set objDepositItems = mobjDepositItems.Clone
    End If
    
    If mobjCardFeeItems Is Nothing Then
        If GetCardFeeBalanceItems(dtCurdate, objCardFeeItems) = False Then Exit Function
    ElseIf mint����״̬ = 1 Then
        Set objCardFeeItems = mobjCardFeeItems.Clone
    ElseIf mobjCardFeeItems.�Ƿ񱣴� = False Then
        If GetCardFeeBalanceItems(dtCurdate, objCardFeeItems) = False Then Exit Function
    Else
        Set objCardFeeItems = mobjCardFeeItems.Clone
    End If
    
    
    mbln��ͬ���� = CheckDepsoitAndCardFeePayIsSame(objDepositItems, objCardFeeItems)
    If GetSaveSendCardInfotoCollect(objPati, dtCurdate, cllSaveSendCardInfo) = False Then Exit Function      '������������
    
    If mbln��ͬ���� = False Or int�䶯���� = 11 Then
        '����ͬ���㣬Ӧ�÷ֲ����н���
        '��һ��:�ȴ���Ԥ����
        Set mobjThirdSwap.objPayCards = mobjDepositPayCards
        If objDepositItems.Count <> 0 And objDepositItems.������� = False Then 'δ�������ʱ����Ҫ�����쳣����
            If objDepositItems.���� = gEM_ҽ�� And mintInsure = 0 Then
                MsgBox "��ǰ���˷�ҽ�����ˣ�������ʹ��" & objDepositItems(1).���㷽ʽ & "���н���.", vbInformation
                Exit Function
            End If
            '1.����Ԥ���쳣
            Set objCurItem = objDepositItems(1)
            If objDepositItems.���� = gEM_һ��ͨ And objDepositItems.ͬ��״̬ < 2 Then '�Ѿ����ýӿڵģ����ٵ���
                '��Ҫ�ȵ��ü��
                intSwapStatu = 0
                If objDepositItems.ͬ��״̬ = 1 Then
                    If mobjThirdSwap.zlThird_IsSwapIsSucces(objDepositItems, intSwapStatu, strErrMsg) = False Then
                        '����ʧ��
                        'intSwapStatu_Out-�ӿڷ���Falseʱ���˲�����Ч:����״̬: 0-���׵���ʧ��;1-�������ڴ�����
                        If intSwapStatu = 1 Then
                            MsgBox "ԭ" & mobjDepositItems(1).objCard.���� & " �������ڽ����У� ����!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                            Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                            mblnDepositLocked = True: Call SetDepositEditEnabled(1)    '�������㷽ʽ
                            Exit Function
                        End If
                    Else
                        intSwapStatu = 1
                    End If
                End If
                If intSwapStatu = 0 Then    'ֻ�н���ʧ��ʱ��������ˢ��
                    If mobjThirdSwap.zlThird_Payment_IsValid(objPati, objCurItem, objItems, dbl�ʻ����) = False Then Exit Function
                    Call objItems.CloneItemsPropertyByItems(objDepositItems)
                    Set objDepositItems = objItems
                End If
                
            ElseIf objDepositItems.���� = gEM_���ѿ� And objDepositItems.ͬ��״̬ < 2 Then
                If GetClassMoney(rsMoney) = False Then Exit Function
                If mobjThirdSwap.zlSquare_Payment_IsValid(objPati, objCurItem, objItems, dbl�ʻ����, , , , rsMoney) = False Then Exit Function
                If objItems.Count > 1 Then
                    MsgBox objCurItem.objCard.���� & "����ͬʱˢ���ſ�����ֻˢһ�ſ����н���!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                Call objItems.CloneItemsPropertyByItems(objDepositItems)
                 Set objDepositItems = objItems
            End If
            
            ' 0��NULL������¼;-1-δ��������;1-δ���ýӿ�;2-�ӿڵ��óɹ�,3-���ý��������ɹ�;4-ҽ�ƿ���Ϣ�����ɹ�
            If objDepositItems.ͬ��״̬ <> 2 Then
                
                'int����-0-������;1-��Ԥ��,2-���Ѽ�Ԥ��
                If GetAddDepositAndCardFeeDataToCollect(1, objPati, Nothing, objDepositItems, dtCurdate, cllDepositAndCardFee) = False Then Exit Function
                lngԤ��ID = objDepositItems(1).Ԥ��ID
                If objDepositItems.ͬ��״̬ = 1 Then   '�Ѿ����ýӿڵģ�ֱ��ɾ��
                    If GetUpdateErrDataSyncTagToColl(objDepositItems.�쳣ID, 1, cllErrData) = False Then Exit Function
                    int�쳣����״̬ = 1
                Else
                    If zlGetErrDataToColl(objPati, lngԤ��ID, objDepositItems, 1, objDepositItems.�쳣ID, dtCurdate, cllErrData, objCurItem.���ݺ�, objCurItem.������) = False Then Exit Function
                    int�쳣����״̬ = IIf(objDepositItems.�Ƿ񱣴�, 1, 0)
                    '0-������¼,1-����״̬�����½���˵����2-ɾ���쳣����
                End If
                
                int״̬ = IIf(objDepositItems.���� = gEM_һ��ͨ Or objDepositItems.���� = gEM_ҽ��, 1, 0)
            
                '------------------------------------------------------------------------------------------------------
                '2.��ʼ���ݱ����
                If objDepositItems.�Ƿ񱣴� = False Then
                    
                    gcnOracle.BeginTrans: blnTrans = True
                    If objDepositItems.���� = gEM_һ��ͨ Or objDepositItems.���� = gEM_ҽ�� Then 'ֻ����������ҽ�������漰�쳣
                        If Zl_���˽����쳣��¼_Modify(int�쳣����״̬, cllErrData) = False Then
                            gcnOracle.RollbackTrans: blnTrans = False
                            Exit Function
                        End If
                    End If
                    
                     '  2.1 ����Ԥ������:����״̬:0-������Ԥ���� ;1-����Ϊδ��Ч��Ԥ����
                    If mobjExseSvr.zl_ExseSvr_AddDepositInfo(int״̬, cllDepositAndCardFee, lngԤ��ID) = False Then
                         gcnOracle.RollbackTrans: Exit Function
                    End If
                
                    gcnOracle.CommitTrans: blnTrans = False: mblnDepositLocked = True
                    
                    Call SetDepositEditEnabled(1)   '����������Ϣ
                    '------------------------------------------------------------------------------------------------------
                    objDepositItems.�Ƿ񱣴� = True: objDepositItems.ͬ��״̬ = 1
                    For i = 1 To objDepositItems.Count
                        objDepositItems(i).�Ƿ񱣴� = True
                    Next
               End If
            End If
         
            Set mobjDepositItems = objDepositItems
            '3.�������������Ҫ����
            If objDepositItems.���� = gEM_һ��ͨ Then
                'һ��ͨ�ۿ�
                If objDepositItems.ͬ��״̬ <> 2 Then
                    Set objCurItem = objDepositItems(1)
                    If mobjThirdSwap.zlThird_Payment(objCurItem.objCard, objPati, cllPro, objDepositItems, objItems, rsExpend, blnSaveed) = False Then
                        If blnSaveed Then
                            Call objItems.CloneItemsPropertyByItems(objDepositItems)
                            If Not objItems Is Nothing Then
                                If objItems.Count > 0 Then Set mobjDepositItems = objItems
                            End If
                        End If
                        Exit Function
                    End If
                    If objItems.Count > 1 Then
                        MsgBox "Ԥ�����֧�ֶ��ֽ��㷽ʽ������", vbInformation + vbOKOnly, Me.Caption
                        Exit Function
                    End If
                                    
                                    
                    Call objItems.CloneItemsPropertyByItems(objDepositItems)
                    Set objDepositItems = objItems
                    Set mobjDepositItems = objDepositItems
                    mobjDepositItems.ͬ��״̬ = 2 '�ӿ��Ѿ��������
                    '��ɽ���
                    Call mobjThirdSwap.zlGetThreeSwapExpendToCollByRecords(rsExpend, cllExpend)
                Else
                    Set cllExpend = Nothing
                End If
                
                If Not mblnDepositLocked Then
                    mblnDepositLocked = True: Call SetDepositEditEnabled(1) '����������Ϣ
                End If
                    
                '����Ԥ��������Ϣ
                If UpdateDepositBlncInfo(0, objPati, objDepositItems, cllExpend) = False Then Exit Function
                objDepositItems.������� = True
                
                If Not mblnDepositLocked Then
                    mblnDepositLocked = True: Call SetDepositEditEnabled(1) '�����������Ϣ
                End If
                
            ElseIf objDepositItems.���� = gEM_ҽ�� Then
                'ҽ������
                If objDepositItems.ͬ��״̬ <> 2 Or mint����״̬ = 1 Then
                    
                    '����ͬ����־ 'ͬ��״̬����������=2,3ʱ��0��NULL������¼;-1-δ��������;1-δ���ýӿ�;2-�ӿڵ��óɹ�,3-���ý��������ɹ�;4-ҽ�ƿ���Ϣ�����ɹ�"
                    If Not GetUpdateErrDataSyncTagToColl(objDepositItems.�쳣ID, 2, cllErrData) Then Exit Function
                    gcnOracle.BeginTrans: blnTrans = True
                    'int����״̬-����״̬:0-������¼,1-����״̬�����½���˵����2-ɾ���쳣����
                    If Zl_���˽����쳣��¼_Modify(1, cllErrData) = False Then
                        gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                    End If
                
                    If Not gclsInsure.TransferSwap(objDepositItems(1).Ԥ��ID, objDepositItems.������, mintInsure) Then
                        gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                    End If
                    gcnOracle.CommitTrans
                    mobjDepositItems.ͬ��״̬ = 2 '�ӿ��Ѿ��������
                     
                    If UpdateDepositBlncInfo(0, objPati, objDepositItems, cllExpend) = False Then Exit Function
                    objDepositItems.������� = True
                    If Not mblnDepositLocked Then
                        mblnDepositLocked = True: Call SetDepositEditEnabled(1) '����������Ϣ
                    End If
                
                Else
                     Set cllExpend = Nothing
                End If
            Else
                '�����޴���
                 objDepositItems.������� = True
            End If
            For i = 1 To objDepositItems.Count
                objDepositItems(i).�Ƿ���� = True
                objDepositItems(i).�Ƿ�����༭ = False
                objDepositItems(i).�Ƿ�����ɾ�� = False
                objDepositItems(i).�Ƿ��������� = False
            Next
            Set mobjDepositItems = objDepositItems
        End If
        
        If int�䶯���� = 11 Then zlSaveData = True: Exit Function
        
        
        '�ڶ���:�ٴ���������
        If fra�ſ�.Visible And objCardFeeItems.������� = False And Trim(txt����.Text) <> "" Then
            Set mobjThirdSwap.objPayCards = mobjCardFeePayCards
            int״̬ = GetCurCard_Statu
            If int״̬ = 11 Then zlSaveData = True: Exit Function
            
            If GetSaveSendCardInfotoCollect(objPati, dtCurdate, cllSendCardInfo) = False Then Exit Function
            If objCardFeeItems.�Ƿ񱣴� = False Or objCardFeeItems.ҵ��ID = 0 Then
                If mobjService.zlPatiSvr_GetNextID("����ҽ�ƿ��䶯", lng�䶯id) = False Then Exit Function
                objCardFeeItems.ҵ��ID = lng�䶯id
            Else
                lng�䶯id = objCardFeeItems.ҵ��ID
                lng�쳣ID = objCardFeeItems.�쳣ID
            End If
            
            ' 0��NULL������¼;-1-δ��������;1-δ���ýӿ�;2-�ӿڵ��óɹ�,3-���ý��������ɹ�;4-ҽ�ƿ���Ϣ�����ɹ�
            'int����-0-������;1-��Ԥ��,2-���Ѽ�Ԥ��
            If objCardFeeItems.���� = gEM_һ��ͨ And objCardFeeItems.ͬ��״̬ < 2 Then
                '��Ҫ�ȵ��ü��
                 Set objCurItem = objCardFeeItems(1)
                 
                intSwapStatu = 0
                If objCardFeeItems.ͬ��״̬ = 1 Then
                    If mobjThirdSwap.zlThird_IsSwapIsSucces(objCardFeeItems, intSwapStatu, strErrMsg) = False Then
                        '����ʧ��
                        'intSwapStatu_Out-�ӿڷ���Falseʱ���˲�����Ч:����״̬: 0-���׵���ʧ��;1-�������ڴ�����
                        If intSwapStatu = 1 Then
                            MsgBox "ԭ" & objCardFeeItems(1).objCard.���� & " �������ڽ����У� ����!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                            Call SetLoaclePayModefromCard(objCardFeeItems(1).objCard, False, True)
                            mblnDepositLocked = True: Call SetDepositEditEnabled(1)    '�������㷽ʽ
                            Exit Function
                        End If
                    Else
                        intSwapStatu = 1
                    End If
                End If
                If intSwapStatu = 0 Then    'ֻ�н���ʧ��ʱ��������ˢ��
                    If mobjThirdSwap.zlThird_Payment_IsValid(objPati, objCurItem, objItems, dbl�ʻ����) = False Then Exit Function
                    Call objItems.CloneItemsPropertyByItems(objCardFeeItems)
                    Set objCardFeeItems = objItems
                End If
                
            ElseIf objCardFeeItems.���� = gEM_���ѿ� And objCardFeeItems.ͬ��״̬ < 2 Then
                If GetClassMoney(rsMoney) = False Then Exit Function
                If mobjThirdSwap.zlSquare_Payment_IsValid(objPati, objCurItem, objItems, dbl�ʻ����, , , , rsMoney) = False Then Exit Function
                
                Call objItems.CloneItemsPropertyByItems(objCardFeeItems)
                Set objCardFeeItems = objItems
            Else
                '����
            End If
            
            ' 0��NULL������¼;-1-δ��������;1-δ���ýӿ�;2-�ӿڵ��óɹ�,3-���ý��������ɹ�;4-ҽ�ƿ���Ϣ�����ɹ�
            If objCardFeeItems.ͬ��״̬ = 0 And objCardFeeItems.�Ƿ񱣴� = False Then 'δ�����䶯��¼��¼
                  
                If objCardFeeItems.ͬ��״̬ = 1 Then   '�Ѿ����ýӿڵģ�ֱ��ɾ��
                    If GetUpdateErrDataSyncTagToColl(objCardFeeItems.�쳣ID, 1, cllErrData) = False Then Exit Function
                    int�쳣����״̬ = 1
                Else
                    If zlGetErrDataToColl(objPati, lng�䶯id, objCardFeeItems, -1, lng�쳣ID, dtCurdate, cllErrData, "", 0, objCardFeeItems.���ݺ�, objCardFeeItems.������) = False Then Exit Function
                    int�쳣����״̬ = IIf(objCardFeeItems.�Ƿ񱣴�, 1, 0)
                    '0-������¼,1-����״̬�����½���˵����2-ɾ���쳣����
                End If
                                
                                
                '------------------------------------------------------------------------------------------------------
                '���ݱ���
                '1.�����쳣���ݼ��䶯��¼
                gcnOracle.BeginTrans: blnTrans = True
                If Zl_���˽����쳣��¼_Modify(int�쳣����״̬, cllErrData) = False Then
                    gcnOracle.RollbackTrans: blnTrans = False
                    Exit Function
                End If
                ' int����״̬:0-������¼;1-�����쳣����;2-ֻ�����䶯��¼
                If mobjService.zlPatisvr_SaveMedcCard(cllSendCardInfo, , , 2, lng�䶯id) = False Then
                   gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                End If
                gcnOracle.CommitTrans: blnTrans = False
                objCardFeeItems.�Ƿ񱣴� = True
                objCardFeeItems.�쳣ID = lng�쳣ID
                objCardFeeItems.ҵ��ID = lng�䶯id
                objCardFeeItems.ͬ��״̬ = -1
                For i = 1 To objCardFeeItems.Count
                    objCardFeeItems(i).�Ƿ񱣴� = True
                    objCardFeeItems(i).�쳣ID = lng�쳣ID
                Next
                Set mobjCardFeeItems = objCardFeeItems
                '------------------------------------------------------------------------------------------------------
            Else
                lng�䶯id = objCardFeeItems.ҵ��ID
                lng�쳣ID = objCardFeeItems.�쳣ID
            End If
            
            '2.���ӿ��ѷ�������
            '����״̬:0-������Ԥ����򿨷ѽɿ�;1-����Ϊδ��Ч��Ԥ������쳣�Ŀ���;2-����Ϊ���ʵ�;3-����Ϊ���۵�
            If objCardFeeItems.ͬ��״̬ = -1 Then
     
                If GetAddDepositAndCardFeeDataToCollect(0, objPati, objCardFeeItems, Nothing, dtCurdate, cllDepositAndCardFee) = False Then Exit Function
                           
                'ͬ��״̬����������=2,3ʱ��0��NULL������¼;-1-δ��������;1-δ���ýӿ�;2-�ӿڵ��óɹ�,3-���ý��������ɹ�;4-ҽ�ƿ���Ϣ�����ɹ�"
                If GetUpdateErrDataSyncTagToColl(lng�쳣ID, IIf(objCardFeeItems.���� = gEM_���ʵ�, 3, 1), cllErrData) = False Then Exit Function
                gcnOracle.BeginTrans
                
                If Zl_���˽����쳣��¼_Modify(1, cllErrData) = False Then
                    gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                End If
                
                int״̬ = IIf(objCardFeeItems.���� = gEM_���ʵ�, 2, 1)
                If mobjExseSvr.Zl_Exsesvr_AddCardFeeInfo(int״̬, cllDepositAndCardFee, lng����ID, lngԤ��ID, True) = False Then
                    '��Ҫɾ���䶯��¼���쳣��¼
                    If GetDelErrDataToColl(lng�䶯id, lng�쳣ID, cllErrData) = False Then
                        gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                        Exit Function
                    End If
                    If Zl_���˽����쳣��¼_Modify(2, cllErrData) = False Then
                          gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                    End If
                    
                    'ɾ���䶯��¼
                    If mobjService.zl_PatiSvr_DelCardChangeInfo(objPati.����ID, lng�䶯id, CLng(cllSendCardInfo("_�����ID")(1)), cllSendCardInfo("_ҽ�ƿ���")(1), True) = False Then
                       gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                    End If
                    gcnOracle.CommitTrans: blnTrans = False: Exit Function
                    Exit Function
                End If
                gcnOracle.CommitTrans: blnTrans = False
                objCardFeeItems.ͬ��״̬ = IIf(objCardFeeItems.���� = gEM_���ʵ�, 3, 1)
                For i = 1 To objCardFeeItems.Count
                    objCardFeeItems(i).����ID = lng����ID
                Next
                Set mobjCardFeeItems = objCardFeeItems
            End If
            
            '------------------------------------------------------------------------------------------------------
            '3.һ���ܵ���ؽ�������
            If objCardFeeItems.���� = gEM_һ��ͨ Then
                'һ��ͨ�ۿ�
                If objCardFeeItems.ͬ��״̬ < 2 Then
                    If mobjThirdSwap.zlThird_Payment(objCurItem.objCard, objPati, cllPro, objCardFeeItems, objItems, rsExpend, blnSaveed) = False Then
                        If blnSaveed Then
                            Call objItems.CloneItemsPropertyByItems(objCardFeeItems)
                            If Not objItems Is Nothing Then
                                If objItems.Count > 0 Then
                                    Set objCardFeeItems = objItems
                                    Set mobjCardFeeItems = objCardFeeItems
                                End If
                            End If
                        End If
                        Exit Function
                    End If
                    If objItems.Count > 1 Then
                        MsgBox "һ��ͨ���ѣ���֧�ֶ��ֽ��㷽ʽ������", vbInformation + vbOKOnly, Me.Caption
                        Exit Function
                    End If
                    
                    Call objItems.CloneItemsPropertyByItems(objCardFeeItems)
                    Set objCardFeeItems = objItems
                            
                    objCardFeeItems.ͬ��״̬ = 2
                    Set mobjCardFeeItems = objCardFeeItems
                Else
                    Set mobjCardFeeItems = objCardFeeItems
                End If
                Call mobjThirdSwap.zlGetThreeSwapExpendToCollByRecords(rsExpend, cllExpend)
                'int����״̬:0-��ɽ���;1-�ӿڵ���ǰ����;2-�ӿڵ��ú�����
                If objCardFeeItems.ͬ��״̬ < 3 Then
                    If UpdateCardFeeBalanceInfor(2, objPati, cllSendCardInfo, objCardFeeItems, Nothing, cllExpend) = False Then Exit Function
                    objCardFeeItems.ͬ��״̬ = 3 '���ý�������
                    Set mobjCardFeeItems = objCardFeeItems
                End If
                If Not mblnSendCardLocked Then
                    mblnSendCardLocked = True:  Call SetCardEditEnabled(1)  '����������Ϣ
                End If
            ElseIf objCardFeeItems.���� = gEM_ҽ�� Then    '������ҽ��
                 'ҽ������
            Else
                '�����޴���
               
            End If
            
            If objCardFeeItems.ͬ��״̬ <= 3 Then
                '4.ҽ�ƿ�����
                'ͬ��״̬����������=2,3ʱ��0��NULL������¼;-1-δ��������;1-δ���ýӿ�;2-�ӿڵ��óɹ�,3-���ý��������ɹ�;4-ҽ�ƿ���Ϣ�����ɹ�"
                If Not GetUpdateErrDataSyncTagToColl(lng�쳣ID, 4, cllErrData, cllSendCardInfo) Then Exit Function
                gcnOracle.BeginTrans: blnTrans = True
                'int����״̬-����״̬:0-������¼,1-����״̬�����½���˵����2-ɾ���쳣����
                If Zl_���˽����쳣��¼_Modify(IIf(objCardFeeItems.���� = gEM_���ʵ�, 2, 1), cllErrData) = False Then
                    gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                End If
                If mobjService.zl_PatiSvr_ConfirmCardChange(objPati.����ID, lng�䶯id, False, cllSendCardInfo) = False Then
                    gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                End If
                gcnOracle.CommitTrans: blnTrans = False:
                objCardFeeItems.ͬ��״̬ = 4
                If objCardFeeItems.���� = gEM_���ʵ� Then
                    objCardFeeItems.������� = True
                End If
                Set mobjCardFeeItems = objCardFeeItems
                
                If Not mblnSendCardLocked Then
                    mblnSendCardLocked = True:     Call SetCardEditEnabled(1)  '����������Ϣ
                End If
            End If
            
            '5.����ȷ��
            'int����״̬:0-��ɽ���;1-�ӿڵ���ǰ����;2-�ӿڵ��ú�����
            If objCardFeeItems.���� <> gEM_���ʵ� Then
                If UpdateCardFeeBalanceInfor(0, objPati, cllSendCardInfo, objCardFeeItems, Nothing, Nothing) = False Then Exit Function
            End If
            If Not mblnSendCardLocked Then
                mblnSendCardLocked = True:  Call SetCardEditEnabled(1)  '����������Ϣ
            End If
            mobjCardFeeItems.������� = True
        End If
        zlSaveData = True
        Exit Function
    End If
    
    '�������Ѽ�Ԥ��ͬ�ֽ��㷽ʽ��ȡ
    If objCardFeeItems.�Ƿ񱣴� = False Or objCardFeeItems.ҵ��ID = 0 Then
        If mobjService.zlPatiSvr_GetNextID("����ҽ�ƿ��䶯", lng�䶯id) = False Then Exit Function
        objCardFeeItems.ҵ��ID = lng�䶯id
        If Not objDepositItems Is Nothing Then
            objDepositItems.ҵ��ID = lng�䶯id
        End If
    Else
        lng�䶯id = objCardFeeItems.ҵ��ID
        lng�쳣ID = objCardFeeItems.�쳣ID
    End If
    Set mobjThirdSwap.objPayCards = mobjCardFeePayCards
    ' 0��NULL������¼;-1-δ��������;1-δ���ýӿ�;2-�ӿڵ��óɹ�,3-���ý��������ɹ�;4-ҽ�ƿ���Ϣ�����ɹ�
    'int����-0-������;1-��Ԥ��,2-���Ѽ�Ԥ��
    If objCardFeeItems.���� = gEM_һ��ͨ And objCardFeeItems.ͬ��״̬ < 2 Then
        '��Ҫ�ȵ��ü��
        intSwapStatu = 0
        If objCardFeeItems.ͬ��״̬ = 1 Then
            Set objItems = objCardFeeItems.Clone
            objItems.������ = objItems.������ + mobjDepositItems.������
            objItems(1).������ = RoundEx(objItems(1).������ + mobjDepositItems.������, 6)
            
            If mobjThirdSwap.zlThird_IsSwapIsSucces(objItems, intSwapStatu, strErrMsg, mobjDepositItems(1).Ԥ��ID) = False Then
                '����ʧ��
                'intSwapStatu_Out-�ӿڷ���Falseʱ���˲�����Ч:����״̬: 0-���׵���ʧ��;1-�������ڴ�����
                If intSwapStatu = 1 Then
                    MsgBox "ԭ" & objCardFeeItems(1).objCard.���� & " �������ڽ����У� ����!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    Call SetLoaclePayModefromCard(objCardFeeItems(1).objCard, True, True)
                    Call SetLoaclePayModefromCard(objCardFeeItems(1).objCard, False, True)
                    mblnSendCardLocked = True: mblnDepositLocked = True
                    Call SetCardEditEnabled(1): Call SetDepositEditEnabled(1)   '�������㷽ʽ
                    
                    Exit Function
                End If
            Else
                intSwapStatu = 1
            End If
        End If
        
        If intSwapStatu = 0 Then    'ֻ�н���ʧ��ʱ��������ˢ��
            Set objCurItem = objCardFeeItems(1).Clone
            objCurItem.������ = objCardFeeItems.������ + objDepositItems.������ '֧������ܶ�
            If mobjThirdSwap.zlThird_Payment_IsValid(objPati, objCurItem, objItems, dbl�ʻ����) = False Then Exit Function
            If objItems.������ <> RoundEx(objCardFeeItems.������ + objDepositItems.������, 5) Then
                MsgBox objCurItem.objCard.���� & "���ص���Ч����뱾��Ҫ����Ľ�һ�£���������Ϊ������ɣ���˲�!" & vbCrLf & _
                    "  ���ؽ��:" & Format(RoundEx(objItems.������, 5), "####0.00;-####0.00;0.00;0.00") & vbCrLf & _
                    "  ���ν���:" & Format(RoundEx(objCardFeeItems.������ + objDepositItems.������, 5), "####0.00;-####0.00;0.00;0.00"), vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            For i = 1 To objCardFeeItems.Count
                  Set objCardFeeItems(i).objCard = objItems(1).objCard
                  objCardFeeItems(i).���㷽ʽ = objItems(1).���㷽ʽ
                  objCardFeeItems(i).������� = objItems(1).�������
                  objCardFeeItems(i).�������� = objItems(1).��������
                  objCardFeeItems(i).���� = objItems(1).����
                  objCardFeeItems(i).�����ID = objItems(1).�����ID
                  objCardFeeItems(i).���� = objItems(1).����
            Next
            If Not objDepositItems Is Nothing Then
                For i = 1 To objDepositItems.Count
                      Set objDepositItems(i).objCard = objItems(1).objCard
                      objDepositItems(i).���㷽ʽ = objItems(1).���㷽ʽ
                      objDepositItems(i).������� = objItems(1).�������
                      objDepositItems(i).�������� = objItems(1).��������
                      objDepositItems(i).���� = objItems(1).����
                      objDepositItems(i).�����ID = objItems(1).�����ID
                      objDepositItems(i).���� = objItems(1).����
                Next
            End If
        End If
    ElseIf objCardFeeItems.���� = gEM_���ѿ� And objCardFeeItems.ͬ��״̬ < 2 Then
        If GetClassMoney(rsMoney) = False Then Exit Function
        
        Set objCurItem = objCardFeeItems(1).Clone
        objCurItem.������ = objCardFeeItems.������ + objDepositItems.������ '֧������ܶ�
        If mobjThirdSwap.zlSquare_Payment_IsValid(objPati, objCurItem, objItems, dbl�ʻ����, , , , rsMoney) = False Then Exit Function
        
        If objItems.������ <> RoundEx(objCardFeeItems.������ + objDepositItems.������, 5) Then
            MsgBox objCurItem.objCard.���� & "���ص���Ч����뱾��Ҫ����Ľ�һ�£���������Ϊ������ɣ���˲�!" & vbCrLf & _
                "  ���ؽ��:" & Format(RoundEx(objItems.������, 5), "####0.00;-####0.00;0.00;0.00") & vbCrLf & _
                "  ���ν���:" & Format(RoundEx(objCardFeeItems.������ + objDepositItems.������, 5), "####0.00;-####0.00;0.00;0.00"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        If objItems.Count > 1 Then
             MsgBox objCurItem.objCard.���� & "����ͬʱˢ���ſ�����ֻˢһ�ſ����н���!", vbInformation + vbOKOnly, gstrSysName
             Exit Function
        End If
        
        For i = 1 To objCardFeeItems.Count
              Set objCardFeeItems(i).objCard = objItems(1).objCard
              objCardFeeItems(i).���㷽ʽ = objItems(1).���㷽ʽ
              objCardFeeItems(i).������� = objItems(1).�������
              objCardFeeItems(i).�������� = objItems(1).��������
              objCardFeeItems(i).���� = objItems(1).����
              objCardFeeItems(i).�����ID = objItems(1).�����ID
              objCardFeeItems(i).���� = objItems(1).����
        Next
        If Not objDepositItems Is Nothing Then
            For i = 1 To objDepositItems.Count
                  Set objDepositItems(i).objCard = objItems(1).objCard
                  objDepositItems(i).���㷽ʽ = objItems(1).���㷽ʽ
                  objDepositItems(i).������� = objItems(1).�������
                  objDepositItems(i).�������� = objItems(1).��������
                  objDepositItems(i).���� = objItems(1).����
                  objDepositItems(i).�����ID = objItems(1).�����ID
                  objDepositItems(i).���� = objItems(1).����
            Next
        End If
    Else
        '��������
    End If
        
    If objCardFeeItems.ͬ��״̬ = 0 And objCardFeeItems.�Ƿ񱣴� = False Then   'δ��������
                    
        If objCardFeeItems.ͬ��״̬ = 1 Then   '�Ѿ����ýӿڵģ�ֱ��ɾ��
            If GetUpdateErrDataSyncTagToColl(objCardFeeItems.�쳣ID, 1, cllErrData) = False Then Exit Function
            int�쳣����״̬ = 1
        Else
            If zlGetErrDataToColl(objPati, lng�䶯id, objCardFeeItems, -1, lng�쳣ID, dtCurdate, cllErrData, objDepositItems.���ݺ�, objDepositItems.������, objCardFeeItems.���ݺ�, objCardFeeItems.������) = False Then Exit Function
            int�쳣����״̬ = IIf(objCardFeeItems.�Ƿ񱣴�, 1, 0)
            '0-������¼,1-����״̬�����½���˵����2-ɾ���쳣����
        End If
            
        '------------------------------------------------------------------------------------------------------
        '���ݱ���
        '1.�����쳣���ݼ��䶯��¼
        gcnOracle.BeginTrans: blnTrans = True
        If Zl_���˽����쳣��¼_Modify(int�쳣����״̬, cllErrData) = False Then
            gcnOracle.RollbackTrans: blnTrans = False
            Exit Function
        End If
        
        If mobjService.zlPatisvr_SaveMedcCard(cllSendCardInfo, , , 2, lng�䶯id) = False Then
           gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        
        gcnOracle.CommitTrans: blnTrans = False
        objCardFeeItems.�Ƿ񱣴� = True
        objCardFeeItems.�쳣ID = lng�쳣ID
        objCardFeeItems.ҵ��ID = lng�䶯id
        objCardFeeItems.ͬ��״̬ = -1
        For i = 1 To objCardFeeItems.Count
            objCardFeeItems(i).�Ƿ񱣴� = True
            objCardFeeItems(i).�쳣ID = lng�쳣ID
        Next
        objDepositItems.�Ƿ񱣴� = True
        objDepositItems.�쳣ID = lng�쳣ID
        objDepositItems.ҵ��ID = lng�䶯id
        objDepositItems.ͬ��״̬ = -1
        For i = 1 To objDepositItems.Count
            objDepositItems(i).�Ƿ񱣴� = True
            objDepositItems(i).�쳣ID = lng�쳣ID
        Next
        
        Set mobjCardFeeItems = objCardFeeItems
        Set mobjDepositItems = objDepositItems
        '------------------------------------------------------------------------------------------------------
    Else
        lng�䶯id = objCardFeeItems.ҵ��ID
        lng�쳣ID = objCardFeeItems.�쳣ID
    End If
    
    '2.���ӿ��ѷ�������
    '����״̬:0-������Ԥ����򿨷ѽɿ�;1-����Ϊδ��Ч��Ԥ������쳣�Ŀ���;2-����Ϊ���ʵ�;3-����Ϊ���۵�
    If objCardFeeItems.ͬ��״̬ = -1 Then
        'int����-0-������;1-��Ԥ��,2-���Ѽ�Ԥ��
        If GetAddDepositAndCardFeeDataToCollect(2, objPati, objCardFeeItems, objDepositItems, dtCurdate, cllDepositAndCardFee) = False Then Exit Function
        lngԤ��ID = objDepositItems(1).Ԥ��ID
        If GetUpdateErrDataSyncTagToColl(lng�쳣ID, 1, cllErrData) = False Then Exit Function
        gcnOracle.BeginTrans
        blnTrans = True:
        If Zl_���˽����쳣��¼_Modify(1, cllErrData) = False Then
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        
        int״̬ = IIf(objCardFeeItems.���� = gEM_���ʵ�, 2, 1)
        If mobjExseSvr.Zl_Exsesvr_AddCardFeeInfo(int״̬, cllDepositAndCardFee, lng����ID, lngԤ��ID, True) = False Then
            '��Ҫɾ���䶯��¼���쳣��¼
            If GetDelErrDataToColl(lng�䶯id, lng�쳣ID, cllErrData) = False Then
                gcnOracle.RollbackTrans: blnTrans = False: Exit Function
                Exit Function
            End If
            If Zl_���˽����쳣��¼_Modify(2, cllErrData) = False Then
                  gcnOracle.RollbackTrans: blnTrans = False: Exit Function
            End If
            
            'ɾ���䶯��¼
            If mobjService.zl_PatiSvr_DelCardChangeInfo(objPati.����ID, lng�䶯id, CLng(cllSendCardInfo("_�����ID")(1)), cllSendCardInfo("_ҽ�ƿ���")(1), True) = False Then
               gcnOracle.RollbackTrans: blnTrans = False: Exit Function
            End If
            gcnOracle.CommitTrans: blnTrans = False: Exit Function
            Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        
        objCardFeeItems.ͬ��״̬ = 1
        objDepositItems.ͬ��״̬ = 1
        For i = 1 To objCardFeeItems.Count
            objCardFeeItems(i).����ID = lng����ID
        Next
        For i = 1 To objDepositItems.Count
            objDepositItems(i).����ID = lng����ID
            objDepositItems(i).Ԥ��ID = lngԤ��ID
        Next
        Set mobjCardFeeItems = objCardFeeItems
        Set mobjDepositItems = objDepositItems
    End If
            
    '------------------------------------------------------------------------------------------------------
    '3.һ��ͨ����ؽ�������
    If objCardFeeItems.���� = gEM_һ��ͨ Then
        'һ��ͨ�ۿ�
        If objCardFeeItems.ͬ��״̬ < 2 Then
            Set objItems = objCardFeeItems.Clone
            If Not objDepositItems Is Nothing Then
                objItems(1).������ = objItems(1).������ + objDepositItems.������
                objItems.������ = objItems.������ + objDepositItems.������
                strDepositNo = objDepositItems.���ݺ�
            End If
            Set objCurItem = objCardFeeItems(1)
            
             If mobjThirdSwap.zlThird_Payment(objCurItem.objCard, objPati, cllPro, objItems, objTempItems, rsExpend, blnSaveed, strDepositNo) = False Then
                If objTempItems Is Nothing Then
                    MsgBox "���������ӿ�֧��ʧ�ܣ�����!", vbInformation, gstrSysName
                    Exit Function
                End If
                If objTempItems.Count = 0 Then
                    MsgBox "���������ӿ�֧��ʧ�ܣ�����!", vbInformation, gstrSysName
                    Exit Function
                End If
                If objTempItems.Count > 1 Then
                    MsgBox "���Ѽ�Ԥ�����ݲ�֧�ֶ��ֽ��㷽ʽ������!", vbInformation, gstrSysName
                    Exit Function
                End If
                Set objItems = objTempItems.Clone
                
                Call objItems.CloneItemsPropertyByItems(objCardFeeItems)
                
                objItems.������ = objCardFeeItems.������
                objItems(1).������ = objCardFeeItems(1).������
                Set objCardFeeItems = objItems
                
                Set objItems = objItems.Clone
                
                objItems.������ = objDepositItems.������
                objItems(1).������ = objDepositItems(1).������
                Set objDepositItems = objItems
                Set mobjCardFeeItems = objCardFeeItems
                Set mobjDepositItems = objDepositItems
                Exit Function
            End If
            
            If RoundEx(objItems.������, 2) <> RoundEx(objTempItems.������, 2) Then
                MsgBox "��ǰ֧���ܶ��뱾��֧�����ܶһ�£�����!", vbInformation, gstrSysName
                Exit Function
            End If
            If objTempItems.Count > 1 Then
                MsgBox "һ��ͨ���ѻ�Ԥ������֧�ֶ��ֽ��㷽ʽ������", vbInformation + vbOKOnly, Me.Caption
                Exit Function
            End If
            Set objItems = objTempItems.Clone
            Call objItems.CloneItemsPropertyByItems(objCardFeeItems)
            
            objItems.������ = objCardFeeItems.������
            objItems(1).������ = objCardFeeItems(1).������
            Set objCardFeeItems = objItems
            
            Set objItems = objItems.Clone
            
            Call objItems.CloneItemsPropertyByItems(objDepositItems)
            
            objItems.������ = objDepositItems.������
            objItems(1).������ = objDepositItems(1).������
            objItems(1).���ݺ� = objDepositItems(1).���ݺ�
            objItems(1).Ԥ��ID = objDepositItems(1).Ԥ��ID
            objItems(1).�Ƿ�Ԥ�� = objDepositItems(1).�Ƿ�Ԥ��
            
            
            Set objDepositItems = objItems
            'ͬ��״̬����������=2,3ʱ��0��NULL������¼;-1-δ��������;1-δ���ýӿ�;2-�ӿڵ��óɹ�,3-���ý��������ɹ�;4-ҽ�ƿ���Ϣ�����ɹ�"
            objCardFeeItems.ͬ��״̬ = 2
            objDepositItems.ͬ��״̬ = 2
            Set mobjCardFeeItems = objCardFeeItems
            Set mobjDepositItems = objDepositItems
            If Not mblnSendCardLocked Then
                mblnSendCardLocked = True: mblnDepositLocked = True
                Call SetCardEditEnabled(1)  '����������Ϣ
                Call SetDepositEditEnabled(1) '����������Ϣ
            End If
            Call mobjThirdSwap.zlGetThreeSwapExpendToCollByRecords(rsExpend, cllExpend)
        Else
            Set mobjCardFeeItems = objCardFeeItems
            Set mobjDepositItems = objDepositItems
            Set cllExpend = Nothing
            If Not mobjCardFeeItems Is Nothing Then Set cllExpend = objCardFeeItems.objTag
            If Not mobjCardFeeItems Is Nothing And cllExpend Is Nothing Then Set cllExpend = objCardFeeItems.objTag
            
        End If
        
        
        'int����״̬:0-��ɽ���;1-�ӿڵ���ǰ����;2-�ӿڵ��ú�����
        If objCardFeeItems.ͬ��״̬ <= 3 Then
            If UpdateCardFeeBalanceInfor(2, objPati, cllSendCardInfo, objCardFeeItems, objDepositItems, cllExpend) = False Then Exit Function
            mobjCardFeeItems.ͬ��״̬ = 3 '���ý�������
            mobjDepositItems.ͬ��״̬ = 3 '���ý�������
        End If
        If Not mblnSendCardLocked Then
            mblnSendCardLocked = True: mblnDepositLocked = True
            Call SetCardEditEnabled(1)  '����������Ϣ
            Call SetDepositEditEnabled(1)  '����������Ϣ
        End If
        
    ElseIf objDepositItems.���� = gEM_ҽ�� Then
         'ҽ������
    Else
        '�����޴���
          
    End If
    
    If objCardFeeItems.ͬ��״̬ <= 3 Then
        '4.ҽ�ƿ�����
        'ͬ��״̬����������=2,3ʱ��0��NULL������¼;-1-δ��������;1-δ���ýӿ�;2-�ӿڵ��óɹ�,3-���ý��������ɹ�;4-ҽ�ƿ���Ϣ�����ɹ�"
        If Not GetUpdateErrDataSyncTagToColl(lng�쳣ID, 4, cllErrData) Then Exit Function
        gcnOracle.BeginTrans: blnTrans = True
        'int����״̬-����״̬:0-������¼,1-����״̬�����½���˵����2-ɾ���쳣����
        If Zl_���˽����쳣��¼_Modify(1, cllErrData) = False Then
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        If mobjService.zl_PatiSvr_ConfirmCardChange(objPati.����ID, lng�䶯id, False, cllSendCardInfo) = False Then
            gcnOracle.RollbackTrans: blnTrans = False: Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        If Not mblnSendCardLocked Then
            mblnSendCardLocked = True: mblnDepositLocked = True
            Call SetCardEditEnabled: Call SetDepositEditEnabled(1) '����������Ϣ
        End If
    End If
    '5.���Ѽ�Ԥ��ȷ��
    'int����״̬:0-��ɽ���;1-�ӿڵ���ǰ����;2-�ӿڵ��ú�����
    If UpdateCardFeeBalanceInfor(0, objPati, cllSendCardInfo, objCardFeeItems, objDepositItems, Nothing) = False Then Exit Function
    Set mobjCardFeeItems = objCardFeeItems
    Set mobjDepositItems = objDepositItems
    mobjCardFeeItems.������� = True
    mobjDepositItems.������� = True
    
    zlSaveData = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetDepositPayCard(Optional ByVal intIndex As Integer = -1) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ֧���Ŀ�����
    '���:intIndex-��ǰ֧��������:-1��ʾֻѡ��ǰѡ���֧�������
    '����:���ؿ�����
    '����:���˺�
    '����:2019-11-06 19:21:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If intIndex = -1 Then
        If cboԤ������.ListIndex < 0 Then Exit Function
        intIndex = cboԤ������.ListIndex
    End If
    Set GetDepositPayCard = mobjDepositPayCards(intIndex + 1)
    Exit Function
errHandle:
    Set GetDepositPayCard = Nothing
End Function
Private Function GetCardFeePayCard(Optional ByVal intIndex As Integer = -1) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ֧���Ŀ�����
    '���:intIndex-��ǰ֧��������:-1��ʾֻѡ��ǰѡ���֧�������
    '����:���ؿ�����
    '����:���˺�
    '����:2019-11-06 19:21:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If intIndex = -1 Then
        If cbo��������.ListIndex < 0 Then Exit Function
        intIndex = cbo��������.ListIndex
    End If
    Set GetCardFeePayCard = mobjCardFeePayCards(intIndex + 1)
    Exit Function
errHandle:
    Set GetCardFeePayCard = Nothing
End Function

Public Function zlSaveDataAfter() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ����ִ��
    '���:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 15:08:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objSendCard As Card
            
    On Error GoTo errHandle
     
     If Not mobjDepositItems Is Nothing And mblnDepositPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & mobjDepositItems.���ݺ�, "�տ�ʱ��=" & Format(mobjDepositItems(1).����ʱ��, "yyyy-mm-dd HH:MM:SS"), _
                                "����ID=" & mobjPati.����ID, IIf(mobjDepositFact.��ӡ��ʽ = 0, "", "ReportFormat=" & mobjDepositFact.��ӡ��ʽ), 2)
            
            If mobjDepositFact.�ϸ���� = False Then
                zlDatabase.SetPara "��ǰԤ��Ʊ�ݺ�", txtFact.Text, glngSys, mlngModule
            End If
     End If
     
    If mbln��ͬ���� And Trim(txtFact.Text) <> "" And Not mobjDepositItems Is Nothing Then
        Call mobjExseSvr.Zl_Exsesvr_Updatedepositinvinf(mobjDepositItems.���ݺ�, mobjDepositFact.����ID, txtFact.Text, UserInfo.����)
    End If
    mblnSendCardLocked = False
    mbln��ͬ���� = False
    Call SetDepositEditEnabled
    Call SetCardEditEnabled
    Call RefreshFactNo
    '���￨���ü��
    Set objSendCard = mCurSendCard.objSendCard
    If Not objSendCard Is Nothing Then
        If objSendCard.�Ƿ��ϸ���� Then
            mCurSendCard.lng����ID = mobjExseSvr.CheckUsedBill(5, IIf(mCurSendCard.lng����ID > 0, mCurSendCard.lng����ID, mCurSendCard.lng��������), , objSendCard.�ӿ����)
            If mCurSendCard.lng����ID <= 0 Then
                Select Case mCurSendCard.lng����ID
                    Case 0 '����ʧ��
                    Case -1
                        If txt����.Text <> "" Then MsgBox "����û�����ü����õ�" & objSendCard.���� & "��,�����ٷ��ţ�" & vbCrLf & _
                            "�����ڱ������ù������λ�����һ���¿���", vbExclamation, gstrSysName
                    Case -2
                        If txt����.Text <> "" Then MsgBox "���ع��õ�" & objSendCard.���� & "��������,�㲻���ٷ��ţ�" & vbCrLf & _
                            "���������ñ��ع��ÿ����λ�����һ���¿���", vbExclamation, gstrSysName
                End Select
            End If
        End If
        'д������
        If fra�ſ�.Visible And objSendCard.�Ƿ�д�� Then Call WriteCard(mobjPati.����ID, objSendCard)
    End If
    Set mobjDepositItems = Nothing
    Set mobjCardFeeItems = Nothing
    zlSaveDataAfter = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function WriteCard(lng����ID As Long, objSendCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:д��
    '���:lng����ID - ����ID
    '����:����
    '����:56599
    '����:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    On Error GoTo ErrHandl:
    If mobjOneCardComLib Is Nothing Then Exit Function
    WriteCard = mobjOneCardComLib.zlBandCardArfter(Me, mlngModule, objSendCard.�ӿ����, lng����ID, strExpend)
    Exit Function
ErrHandl:
    WriteCard = False
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

 
Public Function zlSaveDataBeforCheckIsValid(ByVal blnNewPati As Boolean, ByVal objPati As clsPatientInfo, _
    Optional ByVal bln�Ƿ��Զ�ʶ������֤ As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵĺϷ���
    '���:objPati-������Ϣ��
    '     blnNewPati-�Ƿ��²���
    '     bln�Ƿ��Զ�ʶ������֤-�Ƿ��Զ��ϱ�����֤��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 13:18:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If Trim(txtԤ����.Text) = "" And Trim(txt����.Text) = "" Then zlSaveDataBeforCheckIsValid = True: Exit Function
    If objPati Is Nothing Then
        MsgBox "����ȷ��������Ϣ������!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    If mint����״̬ = 2 Then zlSaveDataBeforCheckIsValid = True: Exit Function '0-����;1-�쳣����;2-�쳣����
    If Not mobjDepositItems Is Nothing Then
        If mobjDepositItems.ͬ��״̬ >= 4 Then zlSaveDataBeforCheckIsValid = True: Exit Function
    
    End If
    
    If Not mobjCardFeeItems Is Nothing Then
        If mobjCardFeeItems.ͬ��״̬ >= 4 Then zlSaveDataBeforCheckIsValid = True: Exit Function
    
    End If
    
    If CheckSendAndBoudCardIsValid(blnNewPati, objPati, bln�Ƿ��Զ�ʶ������֤) = False Then Exit Function
    If CheckDepositIsValid(objPati, mblnDepositPrint) = False Then Exit Function
    
    
    'Ԥ������������ؼ��
    Dim bln��ͬ As Boolean, objCard As Card, objItems As clsBalanceItems, strErrMsg As String, intSwapStatu As Integer
    bln��ͬ = CheckDepsoitAndCardFeePayIsSame(mobjDepositItems, mobjCardFeeItems)
    If bln��ͬ Then
        'һ�����ģ���Ҫ�ж�
        Set objCard = GetCardFeePayCard
        If Not (mobjDepositItems(1).objCard.�ӿ���� = objCard.�ӿ���� And objCard.���ѿ� = mobjDepositItems(1).���ѿ�) And mobjDepositItems.���� = gEM_һ��ͨ Then
            'һ��ͨ���㣬��Ҫ��齻��
            Set objItems = mobjCardFeeItems.Clone
            objItems.������ = objItems.������ + mobjDepositItems.������
            objItems(1).������ = RoundEx(objItems(1).������ + mobjDepositItems.������, 6)
             Set mobjThirdSwap.objPayCards = mobjCardFeePayCards
           
            If mobjThirdSwap.zlThird_IsSwapIsSucces(objItems, intSwapStatu, strErrMsg, mobjDepositItems(1).Ԥ��ID) = False Then
                '����ʧ��
                'intSwapStatu_Out-�ӿڷ���Falseʱ���˲�����Ч:����״̬: 0-���׵���ʧ��;1-�������ڴ�����
                If intSwapStatu = 1 Then
                    MsgBox "ԭ" & mobjDepositItems(1).objCard.���� & " �������ڽ����У����������֧����ʽ,����!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                    Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, False, True)
                    mblnSendCardLocked = True: mblnDepositLocked = True
                    Call SetCardEditEnabled(1): Call SetDepositEditEnabled(1)   '�������㷽ʽ
                    Exit Function
                End If
            Else
                MsgBox "ԭ" & mobjDepositItems(1).objCard.���� & " �����Ѿ��ɹ������������֧����ʽ,����!", vbInformation + vbOKOnly, gstrSysName
                '���׳ɹ�
                Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, False, True)
                mblnSendCardLocked = True: mblnDepositLocked = True
                Call SetCardEditEnabled(1): Call SetDepositEditEnabled(1)   '�������㷽ʽ
                Exit Function
            End If
        ElseIf mobjDepositItems(1).���ѿ� And Not (mobjDepositItems(1).objCard.�ӿ���� = objCard.�ӿ���� And objCard.���ѿ� = mobjDepositItems(1).���ѿ�) Then
            'ԭΪ���ѿ���������˿��ˣ�����ֻ��ԭ����
            MsgBox "ԭ" & mobjDepositItems(1).objCard.���� & " �Ѿ��ۿ�ɹ������������֧����ʽ,����!", vbInformation + vbOKOnly, gstrSysName
            mblnSendCardLocked = True: mblnDepositLocked = True
            Call SetCardEditEnabled(1): Call SetDepositEditEnabled(1)   '�������㷽ʽ
            Exit Function
        End If
        zlSaveDataBeforCheckIsValid = True: Exit Function
    End If
    '����ͬ�ļ��
    'Ԥ�����
    If Not mobjDepositItems Is Nothing Then
    
        If mobjDepositItems.Count <> 0 Then
            Set mobjThirdSwap.objPayCards = mobjDepositPayCards
            Set objCard = GetDepositPayCard
            If Not (mobjDepositItems(1).objCard.�ӿ���� = objCard.�ӿ���� And objCard.���ѿ� = mobjDepositItems(1).���ѿ�) And mobjDepositItems.���� = gEM_һ��ͨ Then
                If mobjThirdSwap.zlThird_IsSwapIsSucces(mobjDepositItems, intSwapStatu, strErrMsg) = False Then
                    '����ʧ��
                    'intSwapStatu_Out-�ӿڷ���Falseʱ���˲�����Ч:����״̬: 0-���׵���ʧ��;1-�������ڴ�����
                    If intSwapStatu = 1 Then
                        MsgBox "ԭ" & mobjDepositItems(1).objCard.���� & " �������ڽ����У����������֧����ʽ,����!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                        Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                        mblnDepositLocked = True: Call SetDepositEditEnabled(1)  '�������㷽ʽ
                        Exit Function
                    End If
                Else
                    '���׳ɹ�
                     MsgBox "ԭ" & mobjDepositItems(1).objCard.���� & " �����Ѿ��ɹ������������֧����ʽ,����!", vbInformation + vbOKOnly, gstrSysName
                    Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                    mblnDepositLocked = True: Call SetDepositEditEnabled(1)  '�������㷽ʽ
                    Exit Function
                End If
            ElseIf mobjDepositItems(1).���ѿ� And Not (mobjDepositItems(1).objCard.�ӿ���� = objCard.�ӿ���� And objCard.���ѿ� = mobjDepositItems(1).���ѿ�) Then
                'ԭΪ���ѿ���������˿��ˣ�����ֻ��ԭ����
                MsgBox "ԭ" & mobjDepositItems(1).objCard.���� & " �Ѿ��ۿ�ɹ������������֧����ʽ,����!", vbInformation + vbOKOnly, gstrSysName
                Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                mblnDepositLocked = True: Call SetDepositEditEnabled(1)  '�������㷽ʽ
                Exit Function
            End If
        End If
    End If
    
    If Not mobjCardFeeItems Is Nothing Then
        If mobjCardFeeItems.Count <> 0 Then
            Set objCard = GetCardFeePayCard
            Set mobjThirdSwap.objPayCards = mobjCardFeePayCards
           
            If Not (mobjCardFeeItems(1).objCard.�ӿ���� = objCard.�ӿ���� And objCard.���ѿ� = mobjCardFeeItems(1).���ѿ�) And mobjCardFeeItems.���� = gEM_һ��ͨ Then
                If mobjThirdSwap.zlThird_IsSwapIsSucces(mobjDepositItems, intSwapStatu, strErrMsg) = False Then
                    '����ʧ��
                    'intSwapStatu_Out-�ӿڷ���Falseʱ���˲�����Ч:����״̬: 0-���׵���ʧ��;1-�������ڴ�����
                    If intSwapStatu = 1 Then
                        MsgBox "ԭ" & mobjCardFeeItems(1).objCard.���� & " �������ڽ����У����������֧����ʽ,����!" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                        Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
                        mblnSendCardLocked = True: Call SetCardEditEnabled(1)  '�������㷽ʽ
                        Exit Function
                    End If
                Else
                    '���׳ɹ�
                     MsgBox "ԭ" & mobjCardFeeItems(1).objCard.���� & " �����Ѿ��ɹ������������֧����ʽ,����!", vbInformation + vbOKOnly, gstrSysName
                    Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
                    mblnSendCardLocked = True: Call SetCardEditEnabled(1)  '�������㷽ʽ
                    Exit Function
                End If
            ElseIf mobjCardFeeItems(1).���ѿ� And Not (mobjCardFeeItems(1).objCard.�ӿ���� = objCard.�ӿ���� And objCard.���ѿ� = mobjCardFeeItems(1).���ѿ�) Then
                'ԭΪ���ѿ���������˿��ˣ�����ֻ��ԭ����
                MsgBox "ԭ" & mobjCardFeeItems(1).objCard.���� & " �Ѿ��ۿ�ɹ������������֧����ʽ,����!", vbInformation + vbOKOnly, gstrSysName
                Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
                mblnSendCardLocked = True: Call SetCardEditEnabled(1)  '�������㷽ʽ
                Exit Function
            End If
        End If
    End If
    zlSaveDataBeforCheckIsValid = True
    
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function




Private Function CheckInputItemIsValid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������ĺϷ���
    '���
    '����:����Ϸ�����true
    '����:���˺�
    '����:2019-11-26 11:55:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objSendCard As Card
    On Error GoTo errHandle
     
    If Not CheckLen(txt�ɿλ, 50, "�ɿλ") Then Exit Function
    If Not CheckLen(txtPass, 10, "����") Then Exit Function
    
    If Not CheckLen(txt������, 50, "������") Then Exit Function
    If Not CheckLen(txt�ʺ�, 50, "�ʺ�") Then Exit Function
    If Not CheckLen(txt�������, 30, "�������") Then Exit Function
        
        
    Set objSendCard = mCurSendCard.objSendCard
    If Not objSendCard Is Nothing Then
        If Not CheckLen(txt����, CInt(objSendCard.���ų���), "����") Then Exit Function
    End If
    CheckInputItemIsValid = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckSendAndBoudCardIsValid(ByVal blnNewPati As Boolean, ByVal objPati As clsPatientInfo, _
    Optional ByVal bln�Ƿ��Զ�ʶ������֤ As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҽ�ƿ��������󶨿����ݵĺϷ���
    '     blnNewPati-�Ƿ��²���
    '     bln�Ƿ��Զ�ʶ������֤-�Ƿ��Զ��ϱ�����֤��
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-09-27 10:21:41
    '����:25302
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCard As String, strICCard As String
    Dim objSendCard As Card, rs���� As ADODB.Recordset
    Dim dtCurrDate As Date, dbl���� As Double
    On Error GoTo errHandle
        
    
    strCard = UCase(txt����.Text)
    dbl���� = Val(txt����.Text)
    strICCard = IIf(mblnICCard, strCard, "")
    
    '-----------------------------------------------------------------------------------------------------------------
    '1.���￨�ļ��
 
    If Not fra�ſ�.Visible Then CheckSendAndBoudCardIsValid = True: Exit Function
    If mblnBoundCarded Then CheckSendAndBoudCardIsValid = True: Exit Function '�Ѿ��󶨿��򷢿��ģ��Ͳ���飬ֱ���˳�
    
    If mlngCardTypeID = 0 Then CheckSendAndBoudCardIsValid = True: Exit Function
    
    
    Set rs���� = GetCardFee()
    Set objSendCard = mCurSendCard.objSendCard
    
    Select Case tbSendCard.SelectedItem.Key
    Case "CardFee"
        If mobjPati Is Nothing Then Set mobjPati = New clsPatientInfo
        
        If (mobjPati.�ѱ� <> objPati.�ѱ� Or mobjPati.ҽ�Ƹ��ʽ <> objPati.ҽ�Ƹ��ʽ) And fra�ſ�.Visible Then
            If tbSendCard.SelectedItem Is Nothing Then Exit Function
            
            If MsgBox("�ѱ�ҽ�Ƹ��ʽ�����˸ı�,�Ƿ���Ҫ���¼��㿨��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
            Call zlRecalcCardFee(objPati)
        End If
        Set mobjPati = objPati
    
        If Trim(txt����.Text) <> "" And Not rs���� Is Nothing Then
            If dbl���� = 0 Then
                MsgBox objSendCard.���� & "δ���뿨�ѣ����飡", vbExclamation, gstrSysName
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus:  Exit Function
            End If
            If rs����!�Ƿ��� = 1 Then
                If rs����!�ּ� <> 0 And Abs(CCur(txt����.Text)) > Abs(rs����!�ּ�) Then
                    MsgBox objSendCard.���� & "��������ֵ���ܴ�������޼ۣ�" & Format(Abs(rs����!�ּ�), "0.00"), vbExclamation, gstrSysName
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus:  Exit Function
                End If
                If rs����!ԭ�� <> 0 And Abs(CCur(txt����.Text)) < Abs(rs����!ԭ��) Then
                    MsgBox objSendCard.���� & "��������ֵ����С������޼ۣ�" & Format(Abs(rs����!ԭ��), "0.00"), vbExclamation, gstrSysName
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Exit Function
                End If
            End If
        End If
        If cbo��������.Visible And txt����.Text <> "" And cbo��������.Enabled And cbo��������.ListIndex = -1 Then
            MsgBox "��ȷ��" & objSendCard.���� & "�Ľɿ���㷽ʽ��", vbExclamation, gstrSysName
            If cbo��������.Enabled And cbo��������.Visible Then cbo��������.SetFocus: Exit Function
        End If
        
        '�������ʵļ��
        If Check��������(objPati.����ID, objSendCard) = False Then Exit Function
        
    Case Else
         Set mobjPati = objPati
    End Select
    
    If bln�Ƿ��Զ�ʶ������֤ = False And InStr(",�������֤,���֤,", "," & objSendCard.���� & ",") > 0 And txt����.Text <> "" Then
            
            MsgBox "�����ֻ֤�����Զ�ʶ��ķ�ʽ���У��������ֶ��������֤���а�!", vbOKOnly + vbInformation, gstrSysName
            txt����.Text = "": txtPass.Text = "": txtAudi.Text = ""
            If txt����.Enabled And txt����.Visible Then txt����.SetFocus
            Exit Function
    End If
    
    
    If txtPass.Text <> txtAudi.Text And txt����.Text <> "" Then
        MsgBox "������������벻һ�£����������룡", vbInformation, gstrSysName
        txtPass.Text = "": txtAudi.Text = ""
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus: Exit Function
    End If
    
    If blnNewPati Then  '�²���
        If Trim(txt����.Text) = "" And txt����.Visible And mblnNewPatiMustSendCard Then
            MsgBox "��ˢ��������" & objSendCard.���� & "���ţ�", vbExclamation, gstrSysName
            If txt����.Enabled And txt����.Enabled Then txt����.SetFocus
            Exit Function
        End If
    End If
    
    
     
    If txt����.Text <> "" Then
        If mobjPubPatient.blnRealName And mobjPati.ʵ����֤ = False And chkEndTime.value = 0 Then
            If MsgBox("δʵ����֤�Ĳ���ֻ�ܷ�����ʱ�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                zlControl.ControlSetFocus chkEndTime
                Exit Function
            End If
            chkEndTime.value = 1
        End If
        
        dtCurrDate = zlDatabase.Currentdate
        If Format(CStr(dtpDate.value), "YYYY-MM-DD HH:MM:SS") < dtCurrDate And chkEndTime.value = vbChecked Then
            MsgBox "��ѡ����ڵ�ǰʱ�����ֹʹ��ʱ�䣡", vbInformation, gstrSysName
            If dtpDate.Enabled And dtpDate.Visible Then dtpDate.SetFocus
            Exit Function
        End If
        
        
        If objSendCard.�Ƿ��ϸ���� Then
            '����ǰ�����￨�Ƿ��У��Ƿ��ڷ�Χ��
           mCurSendCard.lng����ID = mobjExseSvr.CheckUsedBill(5, IIf(mCurSendCard.lng����ID > 0, mCurSendCard.lng����ID, mCurSendCard.lng��������), txt����.Text, objSendCard.�ӿ����)

           If mCurSendCard.lng����ID <= 0 And Not mCurSendCard.blnOneCard Then
               Select Case mCurSendCard.lng����ID
                   Case 0 '����ʧ��
                   Case -1
                           If txt����.Text <> "" Then MsgBox "����û�����ü����õ�" & objSendCard.���� & ",���ܷ��ţ�" & vbCrLf & _
                               "�����ڱ������ù������λ�����һ���¿�! ", vbExclamation, gstrSysName
                   Case -2
                           If txt����.Text <> "" Then MsgBox "���ع��õ�" & objSendCard.���� & "������,���ܷ��ţ�" & vbCrLf & _
                               "���������ñ��ع��ÿ����λ�����һ���¿���", vbExclamation, gstrSysName
                   Case -3
                       MsgBox "���ſ��Ų�����Ч��Χ��,�����Ƿ���ȷˢ����", vbExclamation, gstrSysName
                       If txt����.Enabled And txt����.Enabled Then txt����.SetFocus
               End Select
               Exit Function
           End If
        End If
        
                
        If objSendCard.���ų��� <> zlCommFun.ActualLen(Trim(txt����)) And Not objSendCard.�Ƿ��ϸ���� Then
            '104238:���ϴ���2017/2/15����鿨���Ƿ����㷢����������
            Select Case objSendCard.��������
                Case 0
                    MsgBox "����Ŀ���С��" & objSendCard.���� & "�趨�Ŀ��ų��ȣ����������룡", vbExclamation, gstrSysName
                    If txt����.Visible And txt����.Enabled Then txt����.SetFocus
                    Exit Function
                Case 2
                    If MsgBox("����Ŀ���С��" & objSendCard.���� & "�趨�Ŀ��ų��ȣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        If txt����.Visible And txt����.Enabled Then txt����.SetFocus
                        Exit Function
                    End If
            End Select
        End If
    End If
    
    '������
    If txtPass.Visible Then
        Select Case objSendCard.���볤������
        Case 0
        Case 1
            If Len(txtPass.Text) <> objSendCard.���볤�� Then
                MsgBox "ע��:" & vbCrLf & "�����������" & objSendCard.���볤�� & "λ", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Function
             End If
        Case Else
            If Len(txtPass.Text) < Abs(objSendCard.���볤������) Then
                MsgBox "ע��:" & vbCrLf & "�����������" & Abs(objSendCard.���볤������) & "λ����.", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Function
             End If
        End Select
    End If
                              
    If Len(Trim(txtPass.Text)) <= 0 And Len(Trim(txt����.Text)) > 0 Then 'û����������
        If zl_Get����Ĭ�Ϸ������� = False Then Exit Function
    End If
    
    Dim cllCons As Collection
    Set cllCons = New Collection
    '                   �������Ŀ���ư���:����״̬,����ID,�����ID,����,�¿���
    '                    ����״̬:1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ),7-��ֹʱ�����
    cllCons.Add Array("����״̬", GetCurCard_Statu)
    cllCons.Add Array("����ID", objPati.����ID)
    cllCons.Add Array("�����ID", objSendCard.�ӿ����)
    cllCons.Add Array("����", txt����.Text)
    If mobjService.ZlPatiSvr_ChkCardChangeValid(cllCons) = False Then Exit Function
    
    CheckSendAndBoudCardIsValid = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Check��������(lng����ID As Long, ByVal objSendCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ����Ƿ����Ʋ��˵ķ�������
    '���:lng����ID - ����ID;lng�����ID  - ҽ�ƿ������ID
    '����:����
    '����:57326
    '����:2013-01-30 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, lngPatiID As Long, blnExisted As Boolean
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl:
    
    If Trim(txt����.Text) = "" Or mlngCardTypeID = 0 Then Check�������� = True: Exit Function
    
    If mobjService.ZlPatisvr_CheckCardExist(lng����ID, objSendCard.�ӿ����, "", lngPatiID, blnExisted) = False Then Exit Function
    If Not blnExisted Then Check�������� = True: Exit Function
      
    Select Case objSendCard.��������
    Case 0 '������
        Check�������� = True
    Case 1 'ͬһ������ֻ����һ�ſ�
        MsgBox "�ò����Ѿ�����" & objSendCard.���� & ",�����ڽ��з�������!", vbInformation + vbOKOnly
        Check�������� = False
    Case 2 'ͬһ�������������ſ�,����Ҫ����
       Check�������� = MsgBox("�ò����Ѿ�����" & objSendCard.���� & ",�Ƿ�Ҫ���з�������?", vbQuestion + vbYesNo) = vbYes
    End Select
    Exit Function
ErrHandl:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDepositIsValid(ByVal objPati As clsPatientInfo, Optional blnPrint_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ���ĺϷ���
    '���:objPati-������Ϣ����
    '
    '����:blnPrint_Out-�Ƿ��ӡԤ���վ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 14:12:02
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    
    If fraԤ��.Visible = False Then CheckDepositIsValid = True: Exit Function
    
    If RoundEx(StrToNum(txtԤ����.Text), 4) = 0 Then CheckDepositIsValid = True: Exit Function
 
    If cboԤ������.ListIndex = -1 Then
        MsgBox "��ȷ������Ԥ������㷽ʽ��", vbInformation, gstrSysName
        If cboԤ������.Enabled And cboԤ������.Visible Then cboԤ������.SetFocus
        Exit Function
    End If
    
    If cboԤ������.ItemData(cboԤ������.ListIndex) = 3 Then
        If mintInsure = 0 Then
            MsgBox "��ǰ���˲���ҽ�����ˣ�������ʹ��" & cboԤ������.Text & "����Ԥ����ɿ�.", vbInformation
            Exit Function
        End If
        If mstrҽ���� = "" Then
            MsgBox "��ǰ���˲���ȷ��ҽ���ţ�������ʹ��" & cboԤ������.Text & "����Ԥ����ɿ�.", vbInformation
            Exit Function
        End If
        
        If CCur(StrToNum(txtԤ����.Text)) > mcurYBMoney Then
            MsgBox "ҽ�������ʻ�ת����ܴ������:" & Format(mcurYBMoney, "0.00"), vbInformation, gstrSysName
            If txtԤ����.Enabled And txtԤ����.Visible Then txtԤ����.SetFocus: Exit Function
        End If
  
    End If
            
    blnPrint_Out = True
    Select Case mobjDepositFact.��ӡ��ʽ
    Case "0" '����ӡԤ����Ʊ
        blnPrint_Out = False
    Case "1" '�Զ���ӡ
        blnPrint_Out = True
    Case "2" '��ӡ����
        blnPrint_Out = MsgBox("�Ƿ��ӡԤ����Ʊ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End Select
    
    If blnPrint_Out Then
        If mblnDepositStrictly Then '�ϸ����
            If Trim(txtFact.Text) = "" Then
                MsgBox "��������һ����Ч��Ԥ��Ʊ�ݺ��룡", vbInformation, gstrSysName
                If txtFact.Enabled And txtFact.Visible Then txtFact.SetFocus
                Exit Function
            End If
            
            mobjDepositFact.����ID = mobjExseSvr.CheckUsedBill(2, IIf(mobjDepositFact.����ID > 0, mobjDepositFact.����ID, mobjDepositFact.LastUseID), txtFact.Text, Val(Mid(tbDeposit.SelectedItem.Key, 2)))
            If mobjDepositFact.����ID <= 0 Then
                Select Case mobjDepositFact.����ID
                    Case 0 '����ʧ��
                    Case -1
                        MsgBox "��û�����ú͹��õ�Ԥ��Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    Case -2
                        MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    Case -3
                        MsgBox "Ʊ�ݺ��벻�ڵ�ǰ��Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
                        If txtFact.Enabled And txtFact.Visible Then txtFact.SetFocus
                End Select
                Exit Function
            End If
        Else
            '���ϸ����
            If Len(txtFact.Text) <> mbytԤ��Ʊ�ݳ��� And txtFact.Text <> "" Then
                MsgBox "Ԥ��Ʊ�ݺ��볤��Ӧ��Ϊ " & mbytԤ��Ʊ�ݳ��� & " λ��", vbInformation, gstrSysName
                If txtFact.Enabled And txtFact.Visible Then txtFact.SetFocus
            End If
        End If
    
    End If
     
    CheckDepositIsValid = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub txtAudi_GotFocus()
    zlControl.TxtSelAll txtAudi
    OpenPassKeyboard txtAudi, True
    RaiseEvent ControlGotFocus(txtAudi)
End Sub
Private Sub txtAudi_KeyPress(KeyAscii As Integer)
    Dim objSendCard As Card
    
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Exit Sub
    
    If KeyAscii <> 13 Then
        If objSendCard.������� = 1 Then
            Call zlControl.TxtCheckKeyPress(txtAudi, KeyAscii, m����ʽ)
        End If
    End If
    
    If KeyAscii = 13 Then
        If txtPass.Text <> txtAudi.Text Then
            MsgBox "������������벻һ�£����������룡", vbInformation, gstrSysName
            Call zlControl.TxtSelAll(txtAudi)
            If txtAudi.Enabled And txtAudi.Visible Then txtAudi.SetFocus
        Else
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub
Private Sub txtAudi_LostFocus()
    Call ClosePassKeyboard(txtAudi)
End Sub

Private Sub txtAudi_Validate(Cancel As Boolean)
    Dim objSendCard As Card
    
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Exit Sub
    
    Select Case objSendCard.���볤������
        Case 0
        Case 1
            If Len(txtAudi.Text) <> objSendCard.���볤�� Then
                MsgBox "ע��:" & vbCrLf & "ȷ�������������" & objSendCard.���볤�� & "λ", vbOKOnly + vbInformation
                If txtAudi.Enabled Then txtAudi.SetFocus
                Cancel = True
                Exit Sub
             End If
        Case Else
            If Len(txtAudi.Text) < Abs(objSendCard.���볤������) Then
                MsgBox "ע��:" & vbCrLf & "ȷ�����������" & Abs(objSendCard.���볤������) & "λ����.", vbOKOnly + vbInformation
                If txtAudi.Enabled Then txtAudi.SetFocus
                Cancel = True
                Exit Sub
             End If
        End Select
End Sub





Private Sub txtԤ����_GotFocus()
    If IsNumeric(txtԤ����.Text) Then
        txtԤ����.Text = StrToNum(txtԤ����.Text)
    Else
        txtԤ����.Text = ""
    End If
    txtԤ����.SelStart = 0: txtԤ����.SelLength = Len(txtԤ����.Text)
    RaiseEvent ControlGotFocus(txtԤ����)
End Sub
Private Sub txtԤ����_Validate(Cancel As Boolean)
    Call CalcRQCodePayTotal
End Sub
Private Sub txtԤ����_LostFocus()
    
    If IsNumeric(txtԤ����.Text) Then
        txtԤ����.Text = Format(StrToNum(txtԤ����.Text), "##,##0.00;-##,##0.00; ;")
    Else
        txtԤ����.Text = ""
    End If
    If txtԤ����.MaxLength > 12 Then txtԤ����.MaxLength = 12
End Sub

Private Sub txtԤ����_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If KeyAscii <> 13 Then
        If InStr(txtԤ����.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        '65965:������,2013-09-24,����Ԥ����ʾǧλλ��ʽ
        If (txtԤ����.Text <> "" And txtԤ����.SelLength <> Len(Format(StrToNum(txtԤ����.Text), "##,##0.00;-##,##0.00; ;"))) And _
            (Len(Format(StrToNum(txtԤ����.Text), "##,##0.00;-##,##0.00; ;")) >= txtԤ����.MaxLength) And _
            InStr(Chr(8), Chr(KeyAscii)) = 0 Then
            If txtԤ����.SelLength > 0 And txtԤ����.SelLength <= txtԤ����.MaxLength Then
            Else
                KeyAscii = 0
            End If
        End If
        Exit Sub
    End If
    
    If IsNumeric(txtԤ����.Text) Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If

    '����ȡԤ����,ֱ������
    txtԤ����.Text = ""
    If fra�ſ�.Visible Then
       If txt����.Enabled And txt����.Visible Then txt����.SetFocus
       Exit Sub
    End If
    
    RaiseEvent InputOver '�������
End Sub



Private Sub txtFact_GotFocus()
    zlControl.TxtSelAll txtFact
    RaiseEvent ControlGotFocus(txtFact)
End Sub

Private Sub txtFact_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
    ElseIf Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or InStr("0123456789" & Chr(8), Chr(KeyAscii)) > 0) Then
        KeyAscii = 0
    ElseIf Len(txtFact.Text) = txtFact.MaxLength And KeyAscii <> 8 And txtFact.SelLength <> Len(txtFact) Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub


Private Sub txt�ʺ�_GotFocus()
    If StrToNum(txtԤ����.Text) <> 0 And txt�ʺ�.Text = "" Then txt�ʺ�.Text = mstr��λ�ʺ�
    zlControl.TxtSelAll txt�ʺ�
End Sub

Private Sub txt�ʺ�_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt�ɿλ, KeyAscii
End Sub

Private Sub txt�ʺ�_LostFocus()
    Call zlCommFun.OpenIme
End Sub


Private Sub txt�ɿλ_GotFocus()
    If StrToNum(txtԤ����.Text) <> 0 And txt�ɿλ.Text = "" Then txt�ɿλ.Text = mstr�ɿλ
    zlControl.TxtSelAll txt�ɿλ
    Call zlCommFun.OpenIme(True)
    
    RaiseEvent ControlGotFocus(txt�ɿλ)
End Sub

Private Sub txt�ɿλ_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt�ɿλ, KeyAscii
End Sub

Private Sub txt�ɿλ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt�������_GotFocus()
    zlControl.TxtSelAll txt�������
    RaiseEvent ControlGotFocus(txt�������)
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt�������, KeyAscii
End Sub

  
Private Sub txt������_GotFocus()
    If IsNumeric(txtԤ����.Text) And txt������.Text = "" Then
        txt������.Text = mstr��λ������
    End If
    zlControl.TxtSelAll txt������
    Call zlCommFun.OpenIme(True)
    RaiseEvent ControlGotFocus(txt������)
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt������, KeyAscii
End Sub

Private Sub txt������_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    Dim objSendCard As Card
    
    
    If KeyAscii <> 13 Then
        Set objSendCard = mCurSendCard.objSendCard
        If Not objSendCard Is Nothing Then
            If objSendCard.������� = 1 Then
                Call zlControl.TxtCheckKeyPress(txtPass, KeyAscii, m����ʽ)
            End If
        End If
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPass.Text = "" And txtAudi.Text = "" Then
            If Not txt����.Locked And txt����.TabStop And txt����.Enabled Then
                    txt����.SetFocus
            ElseIf chk����.Visible And chk����.Enabled Then
                chk����.SetFocus
            ElseIf Me.cbo��������.Enabled And cbo��������.Visible Then
                cbo��������.SetFocus
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
           Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
    OpenPassKeyboard txtPass, False
    RaiseEvent ControlGotFocus(txtPass)
End Sub


Private Sub txtPass_LostFocus()
    ClosePassKeyboard txtPass
End Sub
Private Sub txtPass_Validate(Cancel As Boolean)
    Dim objSendCard As Card
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Exit Sub
    Select Case objSendCard.���볤������
    Case 0
    Case 1
        If Len(txtPass.Text) <> objSendCard.���볤�� Then
            MsgBox "ע��:" & vbCrLf & "�����������" & objSendCard.���볤�� & "λ", vbOKOnly + vbInformation
            If txtPass.Enabled Then txtPass.SetFocus
            Exit Sub
         End If
    Case Else
        If Len(txtPass.Text) < Abs(objSendCard.���볤������) Then
            MsgBox "ע��:" & vbCrLf & "�����������" & Abs(objSendCard.���볤������) & "λ����.", vbOKOnly + vbInformation
            If txtPass.Enabled Then txtPass.SetFocus
            Exit Sub
         End If
    End Select
End Sub

Private Function zl_Get����Ĭ�Ϸ�������() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ĭ�Ϸ�������
    '����:�Ƿ������������
    '����:����
    '����:2012-07-06 15:53:14
    '�����:51072
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objSendCard As Card, strID As String
    Dim msgResult As VbMsgBoxResult
 
    Set objSendCard = mCurSendCard.objSendCard
    
    If objSendCard Is Nothing Then Exit Function
    
    If objSendCard.���볤�� = 0 Then  '������
        '������
    ElseIf objSendCard.�Ƿ�ȱʡ���� = 1 Then   'ȱʡ���֤��Nλ
    
        strID = IIf(mobjPati.���֤�� <> "", Trim(mobjPati.���֤��), Trim(mobjPati.��ϵ�����֤��))
        If Len(strID) > 0 Then    '���������֤����ϵ�����֤��
            txtPass.Text = Right(strID, objSendCard.���볤��)
            zl_Get����Ĭ�Ϸ������� = True: Exit Function
        End If
    Else
        zl_Get����Ĭ�Ϸ������� = True: Exit Function
    End If
    
    Select Case objSendCard.������������
        Case 0 '������
            zl_Get����Ĭ�Ϸ������� = True
            Exit Function
        Case 1 'δ��������
            msgResult = MsgBox("δ�������뽫��Ӱ���ʻ���ʹ�ð�ȫ,�Ƿ������", vbQuestion + vbYesNo, gstrSysName)
            zl_Get����Ĭ�Ϸ������� = IIf(msgResult = vbYes, True, False)
            Exit Function
        Case 2 'Ϊ�����ֹ
            MsgBox "δ���뿨����,���ܽ��з�����", vbExclamation, gstrSysName
            zl_Get����Ĭ�Ϸ������� = False
            Exit Function
    End Select
        

End Function


Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    RaiseEvent ControlGotFocus(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rs���� As ADODB.Recordset
    Dim objSendCard As Card
    
    If txt����.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Set rs���� = GetCardFee
        If Not rs���� Is Nothing Then
            Set objSendCard = mCurSendCard.objSendCard
            If rs����!�Ƿ��� = 1 Then
                If rs����!�ּ� <> 0 And Abs(CCur(txt����.Text)) > Abs(rs����!�ּ�) Then
                    MsgBox objSendCard.���� & "��������ֵ���ܴ�������޼ۣ�" & Format(Abs(rs����!�ּ�), "0.00"), vbExclamation, gstrSysName
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Call zlControl.TxtSelAll(txt����): Exit Sub
                End If
                If rs����!ԭ�� <> 0 And Abs(CCur(txt����.Text)) < Abs(rs����!ԭ��) Then
                    MsgBox objSendCard.���� & "��������ֵ����С������޼ۣ�" & Format(Abs(rs����!ԭ��), "0.00"), vbExclamation, gstrSysName
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Call zlControl.TxtSelAll(txt����): Exit Sub
                End If
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr(txt����.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0:  Exit Sub
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0:  Exit Sub
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    If chk����.value = 0 Then Call CalcRQCodePayTotal
End Sub


Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    Call SetBrushCardObject(True)
    RaiseEvent ControlGotFocus(txt����)
End Sub

Private Sub txt����_Change()
    Call SetCardEditEnabled(IIf(mblnSendCardLocked, 1, 0))
    Call CalcRQCodePayTotal '����ɨ�븶�ܶ�
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim objSendCard As Card
    
    'mbln�Ƿ�ɨ�����֤ = False
    
    Set objSendCard = mCurSendCard.objSendCard
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If InStr(":��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> 13 Then
        '118070:���ϴ�,2018/4/24,�豸ֻ���س�����Ҫ����������һλ
        If objSendCard Is Nothing Then Exit Sub
        
        If txt����.SelLength = objSendCard.���ų��� Then txt����.Text = ""
        If Len(txt����.Text) = objSendCard.���ų��� - IIf(objSendCard.�豸�Ƿ����ûس�, 0, 1) And KeyAscii <> 8 Then
            txt����.Text = txt����.Text & IIf(objSendCard.�豸�Ƿ����ûس�, "", Chr(KeyAscii))
            
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        End If
        
    ElseIf txt����.Text = "" Then
        KeyAscii = 0: RaiseEvent InputOver
    Else
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
     
End Sub

Private Sub txt����_LostFocus()
    Call SetBrushCardObject(False)
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim lngPatientID As Long, int�䶯���� As Integer
    Dim blnCardBind As Boolean  '���Ƿ���а�
    Dim objSendCard As Card
    
    Set objSendCard = mCurSendCard.objSendCard
    
    txt����.Text = Trim(txt����.Text)
    Call ReLoadCardFee
    Call CheckFreeCard(txt����.Text)

    If objSendCard.���ų��� = Len(Trim(txt����.Text)) Then
        
        If mobjOneCardComLib.objOneCardObject.zlGetPatiIDFromCardNo(objSendCard.�ӿ����, Trim(txt����.Text), lngPatientID, False, False) = False Then Exit Sub
         
        If objSendCard.���ƿ� And objSendCard.�����ظ�ʹ�� And lngPatientID > 0 Then
        
           Call mobjService.zlPatiSvr_GetCardLastChange(lngPatientID, objSendCard.�ӿ����, txt����.Text, int�䶯����)
            If int�䶯���� = 11 Then
                '����ǰ�
                If MsgBox("����Ϊ��" & txt����.Text & "����{" & objSendCard.���� & "}�Ŀ��Ѿ��벡�˱�ʶΪ��" & lngPatientID & "���Ľ����˰󶨣�" & vbCrLf & "�Ƿ�ȡ���ÿ��İ�?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                    Cancel = True
                    txt����.Text = ""
                    Exit Sub
                End If
                If BlandCancel(objSendCard.�ӿ����, Trim(txt����.Text), lngPatientID) Then Exit Sub
            End If
        End If

        MsgBox "�ÿ����Ѿ�����,���ܰ󶨸ÿ���.", vbInformation, gstrSysName
        Cancel = True
        txt����.Text = ""
        Exit Sub
   End If
    
End Sub

Private Function BlandCancel(ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡ���󶨿�
    '���:intType:0-��ǰ����;1-��ǰ���;2-��ǰ��������
    '����:ȡ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-29 11:18:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtCurdate As Date
    Dim cllSaveCard As Collection
    On Error GoTo errHandle

    dtCurdate = zlDatabase.Currentdate
    
    ' ��� :cllCard-�ڵ����:��������,����ID,�����ID,ԭ����,ҽ�ƿ���,��ά��,�䶯ԭ��,����,IC����,��ʧ��ʽ,��ֹʹ��ʱ��,���ݺ�,����,����ʱ��,����Ա����,����Ա���
    '               ���еĲ�������:1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ),7-��ֹʱ�����
    '               ÿ���ʽ:array("����",ֵ )
    Set cllSaveCard = New Collection
    cllSaveCard.Add Array("��������", 14)
    cllSaveCard.Add Array("����ID", lng����ID)
    cllSaveCard.Add Array("�����ID", lngCardTypeID)
    cllSaveCard.Add Array("ҽ�ƿ���", strCardNo)
    cllSaveCard.Add Array("�䶯ԭ��", "���ظ��Զ�ȡ��ԭ������Ϣ")
    cllSaveCard.Add Array("����ʱ��", Format(dtCurdate, "yyyy-mm-dd HH:MM:SS"))
    cllSaveCard.Add Array("����Ա���", UserInfo.���)
    cllSaveCard.Add Array("����Ա����", UserInfo.����)
    If mobjService.zlPatisvr_SaveMedcCard(cllSaveCard) = False Then Exit Function
    BlandCancel = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub ReLoadCardFee()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼��ؿ�����
    '����:���˺�
    '����:2019-11-25 15:52:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, lng�շ�ϸĿid As Long
    Dim strSql As String, str���� As String
    Dim rsTmp As ADODB.Recordset, rs���� As ADODB.Recordset
    Dim objSendCard As Card, dblMoney As Double
    Dim objPati As clsPatientInfo
    On Error GoTo errHandle
    Set rs���� = GetCardFee
    
    If rs���� Is Nothing Or Trim(txt����.Text) = "" Then Exit Sub
    If mobjPati Is Nothing Or rs����.RecordCount = 0 Then Exit Sub
    
    
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Exit Sub
    
    If objSendCard.�ӿ���� = 0 Then Exit Sub
    
    lng����ID = mobjPati.����ID
    str���� = mobjPati.����
     
    rs����.MoveFirst
    
    strSql = "Select Zl1_Ex_CardFee([1],[2],[3],[4],[5],[6],[7],[8],[9]) as �շ�ϸĿID From Dual "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "����", mlngModule, objSendCard.�ӿ����, Trim(txt����.Text), lng����ID, _
                mobjPati.����, mobjPati.�Ա�, mobjPati.����, mobjPati.���֤��, Val(Nvl(rs����!�շ�ϸĿID)))
    If rsTmp.EOF Then Exit Sub
    
    lng�շ�ϸĿid = Val(Nvl(rsTmp!�շ�ϸĿID))
    Set rsTmp = zlGetSpecialItemFee(objSendCard.�ض���Ŀ, mstrPriceGrade, lng�շ�ϸĿid)
    If Not rsTmp Is Nothing Then Set rs���� = rsTmp
    
    With rs����
        txt����.Text = Format(IIf(Val(Nvl(!�Ƿ���)) = 1, Val(Nvl(!ȱʡ�۸�)), Val(Nvl(!�ּ�))), "0.00")
        txt����.Tag = txt����.Text  '���ֲ���
        txt����.Locked = Not (Val(Nvl(!�Ƿ���)) = 1)
        txt����.TabStop = (Val(Nvl(!�Ƿ���)) = 1)
        If rs����!�Ƿ��� = 0 And Val(txt����.Text) <> 0 Then
            If mobjExseSvr.zl_ExseSvr_Actualmoney(mobjPati.�ѱ�, rs����!�շ�ϸĿID, rs����!������ĿID, rs����!�ּ�, dblMoney) = False Then Exit Sub
            txt����.Text = Format(dblMoney, "0.00")
        End If
    End With
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������봴��
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional blnȷ������ As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, blnȷ������) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function
 
Private Sub CalcRQCodePayTotal(Optional bln�쳣 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ɨ�븶���
    '����:2019-11-25 15:35:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    mdblRQCodeMoney = 0
    
    If mobjShowTotalMoneyControl Is Nothing Then Exit Sub
    
    If chk����.value = 0 And (txt����.Visible Or bln�쳣) And StrToNum(txt����.Text) <> 0 And Trim(txt����.Text) <> "" Then
        mdblRQCodeMoney = StrToNum(txtԤ����.Text) + StrToNum(txt����.Text)
    Else
        mdblRQCodeMoney = StrToNum(txtԤ����.Text)
    End If
        
    If UCase(TypeName(mobjShowTotalMoneyControl)) = UCase("TextBox") Then
        mobjShowTotalMoneyControl.Text = Format(mdblRQCodeMoney, "0.00")
    Else
        mobjShowTotalMoneyControl.Caption = "ɨ��ϼƣ�" & Format(mdblRQCodeMoney, "0.00")
    End If
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function SetBrushCardObject(ByVal blnComm As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ���ӿ�
    '����: true-�ɹ���false-ʧ��
    '����:���ϴ�
    '����:2016/6/20 13:54:56
    '����:97634
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    Dim objSendCard As Card
    Err = 0: On Error Resume Next
    SetBrushCardObject = True
    
    If txt����.Locked Then Exit Function
    If mobjOneCardComLib Is Nothing Then Exit Function
    
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Exit Function
    
    
    If objSendCard.�ӿ���� <= 0 Or Not (objSendCard.�Ƿ�ɨ�� Or objSendCard.�Ƿ�ˢ��) Then Exit Function
    
    If mobjOneCardComLib.zlSetBrushCardObject(objSendCard.�ӿ����, IIf(blnComm, txt����, Nothing), strExpend) Then
        If mobjCommEvents Is Nothing Then Set mobjCommEvents = New clsCommEvents
        Call mobjOneCardComLib.zlInitEvents(Me.hWnd, mobjCommEvents)
    End If
End Function



Private Function GetSaveSendCardInfotoCollect(ByVal objPati As clsPatientInfo, ByVal dtCurdate As Date, ByRef cllCardInfo_out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ҽ�ƿ����ݼ�
    '���:
    '����:cllCardInfo_Out-���ؿ����ݼ�,��ʽ:array(����,ֵ)
    '         |-��������,����ID,�����ID,ԭ����,ҽ�ƿ���,��ά��,�䶯ԭ��,����,IC����,��ʧ��ʽ,��ֹʹ��ʱ��,���ݺ�,����,����ʱ��,����Ա����,����Ա���,��������,����id
    '         ��������:1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ),7-��ֹʱ�����
    '     cllCardFeeInfo_Out-��������Ϣ:
    '
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 18:58:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnInRange   As Boolean, strCardNo As String, strICCard As String
    Dim objSendCard As Card, byt�䶯���� As Byte, strEndDate As String
    Dim str�䶯ԭ�� As String
    
    On Error GoTo errHandle
    
    Set objSendCard = mCurSendCard.objSendCard
    If mlngCardTypeID = 0 Then GetSaveSendCardInfotoCollect = True: Exit Function
    If objSendCard Is Nothing Then Exit Function

    Set cllCardInfo_out = New Collection
    byt�䶯���� = GetCurCard_Statu
    strEndDate = ""
    If chkEndTime.value = vbChecked Then
        strEndDate = Format(dtpDate.value, "yyyy-mm-dd HH:MM:SS")
    End If
    
    str�䶯ԭ�� = Decode(mlngModule, 1101, "������Ϣ�ǼǷ���", "������Ժ�ǼǷ���")
    
    strCardNo = UCase(txt����.Text): strICCard = IIf(mblnICCard, strCardNo, "")
    If strCardNo = "" Then GetSaveSendCardInfotoCollect = True: Exit Function
    
    '1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ),7-��ֹʱ�����
    cllCardInfo_out.Add Array("��������", byt�䶯����), "_��������"
    cllCardInfo_out.Add Array("����ID", objPati.����ID), "_����ID"
    cllCardInfo_out.Add Array("�����ID", objSendCard.�ӿ����), "_�����ID"
    cllCardInfo_out.Add Array("ԭ����", ""), "_ԭ����"
    cllCardInfo_out.Add Array("ҽ�ƿ���", strCardNo), "_ҽ�ƿ���"
    cllCardInfo_out.Add Array("����ID", mCurSendCard.lng����ID), "_����ID"
    
    cllCardInfo_out.Add Array("��ά��", ""), "_��ά��"
    cllCardInfo_out.Add Array("�䶯ԭ��", str�䶯ԭ��), "_�䶯ԭ��"
    cllCardInfo_out.Add Array("����", zlCommFun.zlStringEncode(Trim(txtPass.Text))), "_����"
    cllCardInfo_out.Add Array("IC����", strICCard), "_IC����"
    cllCardInfo_out.Add Array("��ʧ��ʽ", ""), "_��ʧ��ʽ"
    cllCardInfo_out.Add Array("��ֹʹ��ʱ��", strEndDate), "_��ֹʹ��ʱ��"
    cllCardInfo_out.Add Array("���ݺ�", ""), "_���ݺ�"
    cllCardInfo_out.Add Array("����", StrToNum(txt����.Text)), "_����"
    cllCardInfo_out.Add Array("����ʱ��", Format(dtCurdate, "yyyy-mm-dd HH:MM:SS")), "_����ʱ��"
    cllCardInfo_out.Add Array("����Ա����", UserInfo.����), "_����Ա����"
    cllCardInfo_out.Add Array("����Ա���", UserInfo.���), "_����Ա���"
    cllCardInfo_out.Add Array("��������", IIf(mCurSendCard.objSendCard.�����ظ�ʹ��, 1, 0)), "_��������"
    If mCurSendCard.lng����ID > 0 Then cllCardInfo_out.Add Array("����ID", mCurSendCard.lng����ID), "_����ID"
    GetSaveSendCardInfotoCollect = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetDepositSaveDataToCollect(ByVal objPati As clsPatientInfo, ByVal objDepositItems As clsBalanceItems, _
    ByRef cllDeposit As Collection, Optional ByVal strCardFeeNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����Ԥ�����ݼ�
    '���:objPati-������Ϣ����
    '     strCardFeeNo-���ѵ��ݺţ�ͬʱ�ɿ�ʱ������(Ԥ�����ݺ�,��Ʊ��,Ԥ�����,����ID,��ҳid,����,�Ա�,����,�����,סԺ��,���ʽ���,���ʽ����,�ɿ����id,�ɿ���,�ɿλ,��λ������,ժҪ,����id)
    '     strDepositNo_Out-Ԥ�����ݺ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-10 19:50:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney  As Double
    Dim intԤ������ As Integer, int��ҳid As Long

    On Error GoTo errHandle
        
    dblMoney = StrToNum(txtԤ����.Text)
    Set cllDeposit = New Collection
    
    If fraԤ��.Visible = False Or dblMoney = 0 Then GetDepositSaveDataToCollect = True: Exit Function
    
    If objDepositItems Is Nothing Then Exit Function
    
    intԤ������ = Val(Mid(tbDeposit.SelectedItem.Key, 2))
    int��ҳid = 0
    If intԤ������ = 2 Then int��ҳid = objPati.��ҳID
    
    'depositinfo:(Ԥ�����ݺ�,��Ʊ��,Ԥ�����,����ID,��ҳid,����,�Ա�,����,�����,סԺ��,���ʽ���,���ʽ����,�ɿ����id,�ɿ���,�ɿλ,��λ������,ժҪ,����id)
    If objDepositItems.���ݺ� = "" Then Exit Function
    cllDeposit.Add Array("Ԥ��ID", objDepositItems(1).Ԥ��ID), "_Ԥ��ID"
    cllDeposit.Add Array("Ԥ�����ݺ�", objDepositItems.���ݺ�), "_Ԥ�����ݺ�"
    cllDeposit.Add Array("��Ʊ��", IIf(mblnDepositPrint, txtFact.Text, "")), "_��Ʊ��"
    cllDeposit.Add Array("Ԥ�����", Val(Mid(tbDeposit.SelectedItem.Key, 2))), "_Ԥ�����"
    cllDeposit.Add Array("����ID", objPati.����ID), "_����ID"
    cllDeposit.Add Array("��ҳID", int��ҳid), "_��ҳID"
    cllDeposit.Add Array("����", objPati.����), "_����"
    cllDeposit.Add Array("�Ա�", objPati.�Ա�), "_�Ա�"
    cllDeposit.Add Array("����", objPati.����), "_����"
    cllDeposit.Add Array("�����", objPati.�����), "_�����"
    cllDeposit.Add Array("סԺ��", objPati.סԺ��), "_סԺ��"
    cllDeposit.Add Array("���ʽ���", objPati.ҽ�Ƹ��ʽ����), "_���ʽ���"
    cllDeposit.Add Array("���ʽ����", objPati.ҽ�Ƹ��ʽ), "_���ʽ����"
    cllDeposit.Add Array("�ɿ����ID", Val(txt�ɿλ.Tag)), "_�ɿ����ID"
    cllDeposit.Add Array("�ɿ���", dblMoney), "_�ɿ���"
    cllDeposit.Add Array("�ɿλ", txt�ɿλ.Text), "_�ɿλ"
    cllDeposit.Add Array("��λ������", txt������.Text), "_��λ������"
    cllDeposit.Add Array("�������˺�", txt�ʺ�.Text), "_�������˺�"
    cllDeposit.Add Array("ժҪ", IIf(strCardFeeNo = "", "", "ҽ�ƿ�:" & strCardFeeNo)), "_ժҪ"
    cllDeposit.Add Array("����ID", mobjDepositFact.����ID), "_����ID"
    GetDepositSaveDataToCollect = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetCardFeeBalanceSaveDataToColl(ByVal objPati As clsPatientInfo, ByVal objCurBalanceItem As clsBalanceItem, ByRef cllBalanceData_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������Ϣ
    '���:objCurBalanceItem-��ǰ������Ϣ
    '
    '����:cllBalanceData_Out-���������Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-11 09:26:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    ' balanceinfo:(���㷽ʽ,�������,�����id,���㿨���,֧������,������ˮ��,����˵��,������λ,����,ҽ����,ҽ������,���ѿ�ID) Key="_balanceinfo"
    Set cllBalanceData_Out = New Collection
    
    cllBalanceData_Out.Add Array("���㷽ʽ", objCurBalanceItem.���㷽ʽ), "_" & "���㷽ʽ"
    cllBalanceData_Out.Add Array("�������", objCurBalanceItem.�������), "_" & "�������"
    cllBalanceData_Out.Add Array("�����ID", IIf(Not objCurBalanceItem.���ѿ�, objCurBalanceItem.�����ID, "")), "_" & "�����ID"
    cllBalanceData_Out.Add Array("���㿨���", IIf(objCurBalanceItem.���ѿ�, objCurBalanceItem.�����ID, "")), "_" & "���㿨���"
    cllBalanceData_Out.Add Array("֧������", objCurBalanceItem.����), "_" & "֧������"
    cllBalanceData_Out.Add Array("������ˮ��", objCurBalanceItem.������ˮ��), "_" & "������ˮ��"
    cllBalanceData_Out.Add Array("����˵��", objCurBalanceItem.����˵��), "_" & "����˵��"
    cllBalanceData_Out.Add Array("������λ", ""), "_" & "������λ"
    cllBalanceData_Out.Add Array("���ѿ�ID", objCurBalanceItem.���ѿ�ID), "_" & "���ѿ�ID"
    
    If objCurBalanceItem.�������� = 3 Then
        cllBalanceData_Out.Add Array("����", mintInsure), "_" & "����"
        cllBalanceData_Out.Add Array("ҽ����", mstrҽ����), "_" & "ҽ����"
        cllBalanceData_Out.Add Array("����", mstr����), "_" & "����"
    End If
    GetCardFeeBalanceSaveDataToColl = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
  
Private Function GetAddDepositAndCardFeeDataToCollect(ByVal int���� As Integer, ByVal objPati As clsPatientInfo, _
    ByVal objCardFeeItems As clsBalanceItems, ByVal objDepositItems As clsBalanceItems, _
     ByVal dtCurdate As Date, ByRef cllData_out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���һ��ͨ�Ƿ���ȷ
    '���:objPati-������Ϣ��
    '     int����-0-������;1-��Ԥ��,2-���Ѽ�Ԥ��
    '     objCardFeeItems-��ǰ���ѽ�����Ϣ
    '     objDepositItems-��ǰԤ��������Ϣ
    '����:
    '     cllData_Out: �����ݶ���
    '          |--billinfo:(����ϼ�,����Ա���,����Ա����,�Ǽ�ʱ��),Key="_billinfo"
    '          |--patinfo:(����ID,��ҳID,��������,�Ա�,����,�����,סԺ��,���ʽ���,�ѱ�,����),Key="_patinfo"
    '          |--cardinfo:������Ϣ(����,�����ID,������ʽ(0-����,1-����,2-����),��������,����id),key="_cardinfo"
    '          |--cardfeelists:key="_cardfeelists"
    '               |---cardfeelist:(���ѵ��ݺ�,���,�۸񸸺�,��������,�շ����,�շ�ϸĿid,������Ŀid,��׼����,�վݷ�Ŀ,Ӧ�ս��,ʵ�ս��,���˿���id,��������id,���˲���id,
    '                                 ִ�в���id,�Ӱ��־,�Ƿ�����,���ձ���,������Ŀ��,ͳ����,ժҪ,��������,���������ID,������ʽ(0-����,1-����,2-����)) ,Key="_" & ���
    
    '          |--balanceinfo:(���㷽ʽ,�������,�����id,���㿨���,֧������,������ˮ��,����˵��,������λ,����,ҽ����,ҽ������,���ѿ�ID) Key="_balanceinfo"
    '          |--depositinfo:(Ԥ�����ݺ�,��Ʊ��,Ԥ�����,��ҳid,�ɿ����id,�ɿ���,�ɿλ,��λ������,ժҪ,����id),Key="_depositinfo",��Ԥ��ʱ��������
    '          ���ϣ���ʽΪ:,��ʽ��array(����,ֵ)
    '          int����״̬=2-����Ϊ���ʵ�;3-����Ϊ���۵� �ģ�����"balanceinfo"��"depositinfo"�ڵ�
    
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-11 10:37:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllDeposit As Collection, cllCardFee As Collection
    Dim cllBalanceInfo As Collection, strCardFeeNo As String
    Dim cllTemp As Collection
    
    On Error GoTo errHandle
    
    Set cllData_out = New Collection
    
     
    '��ȡ��������Ϣ��
    Set cllCardFee = New Collection
    If int���� = 0 Or int���� = 2 Then  '���Ѵ���
        If GetCardFeeSaveDataToCollect(objPati, objCardFeeItems, cllCardFee) = False Then Exit Function
        If Not objCardFeeItems Is Nothing Then strCardFeeNo = objCardFeeItems.���ݺ�
    End If
    
    '��ȡԤ����Ϣ��
    If int���� = 1 Or int���� = 2 Then
        If GetDepositSaveDataToCollect(objPati, objDepositItems, cllDeposit, strCardFeeNo) = False Then Exit Function
        If objDepositItems Is Nothing Then Set objDepositItems = New clsBalanceItems
    Else
        Set cllDeposit = New Collection
        Set objDepositItems = New clsBalanceItems
    End If
    
    Set cllBalanceInfo = New Collection
    If cllCardFee.Count <> 0 And chk����.value = 0 Or cllDeposit.Count <> 0 Then
        '������������
        If objDepositItems.Count <> 0 Then
            If GetCardFeeBalanceSaveDataToColl(objPati, objDepositItems(1), cllBalanceInfo) = False Then Exit Function
        Else
            If GetCardFeeBalanceSaveDataToColl(objPati, objCardFeeItems(1), cllBalanceInfo) = False Then Exit Function
        End If
    End If
   
   
    '1.����������Ϣ����
    Set cllTemp = New Collection
    If int���� <> 1 Then
        If objDepositItems Is Nothing Then
            cllTemp.Add Array("����ϼ�", RoundEx(objCardFeeItems.������, 5)), "_" & "����ϼ�"
        Else
            cllTemp.Add Array("����ϼ�", RoundEx(objCardFeeItems.������ + objDepositItems.������, 5)), "_" & "����ϼ�"
        End If
    End If
    cllTemp.Add Array("����Ա���", UserInfo.���), "_" & "����Ա���"
    cllTemp.Add Array("����Ա����", UserInfo.����), "_" & "����Ա����"
    If Not mobjCardFeeItems Is Nothing Then
        
        If mobjCardFeeItems.Count <> 0 Then
            cllTemp.Add Array("����ID", mobjCardFeeItems(1).����ID), "_" & "����ID"
        End If
    End If
    cllTemp.Add Array("�Ǽ�ʱ��", Format(dtCurdate, "yyyy-mm-dd HH:MM:SS")), "_" & "�Ǽ�ʱ��"
    cllData_out.Add cllTemp, "_billinfo"
    
    
    '2.����������Ϣ
    If int���� <> 1 Then
        Set cllTemp = New Collection
        cllTemp.Add Array("����ID", objPati.����ID), "_" & "����ID"
        cllTemp.Add Array("��ҳID", objPati.��ҳID), "_" & "��ҳID"
        cllTemp.Add Array("��������", objPati.����), "_" & "��������"
        cllTemp.Add Array("�Ա�", objPati.�Ա�), "_" & "�Ա�"
        cllTemp.Add Array("����", objPati.����), "_" & "����"
        cllTemp.Add Array("�����", objPati.�����), "_" & "�����"
        cllTemp.Add Array("סԺ��", objPati.סԺ��), "_" & "סԺ��"
        cllTemp.Add Array("���ʽ���", objPati.ҽ�Ƹ��ʽ����), "_" & "���ʽ���"
        cllTemp.Add Array("���ʽ����", objPati.ҽ�Ƹ��ʽ), "_" & "���ʽ����"
        cllTemp.Add Array("�ѱ�", objPati.�ѱ�), "_" & "�ѱ�"
        cllTemp.Add Array("����", 0), "_" & "����"
        cllData_out.Add cllTemp, "_patinfo"
        '3.����������Ϣ
        If cllCardFee.Count <> 0 Then
            '����,�����ID,������ʽ(0-����,1-����,2-����),��������,����id
            '2.������Ϣ
            Set cllTemp = New Collection
            cllTemp.Add Array("����", Trim(txt����.Text)), "_" & "����"
            cllTemp.Add Array("�����ID", mCurSendCard.objSendCard.�ӿ����), "_" & "�����ID"
            cllTemp.Add Array("������ʽ", 0), "_" & "������ʽ"  '0-����,1-����,2-����
            cllTemp.Add Array("��������", IIf(mCurSendCard.objSendCard.�����ظ�ʹ��, 1, 0)), "_" & "��������"
            cllTemp.Add Array("����ID", mCurSendCard.lng����ID), "_" & "����ID"
            cllData_out.Add cllTemp, "_cardinfo"
            
            '����
            cllData_out.Add cllCardFee, "_cardfeelists"
        End If
    
    End If
    
    '4.������Ϣ
    If cllBalanceInfo.Count <> 0 Then
        cllData_out.Add cllBalanceInfo, "_balanceinfo"    '������Ϣ
        If Not cllDeposit Is Nothing Then
            If cllDeposit.Count <> 0 Then
                cllData_out.Add cllDeposit, "_depositinfo"  '������Ϣ
            End If
        End If
    End If
    GetAddDepositAndCardFeeDataToCollect = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetCurCard_Statu() As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����״̬
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-25 21:22:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objSendCard As Card, blnInRange As Boolean
    If fra�ſ�.Visible = False Then Exit Function
    
    Set objSendCard = mCurSendCard.objSendCard
    If objSendCard Is Nothing Then Set objSendCard = New Card
    
    blnInRange = True
    If mCurSendCard.blnOneCard And objSendCard.�Ƿ��ϸ���� Then blnInRange = mCurSendCard.lng����ID > 0
    If blnInRange And tbSendCard.SelectedItem.Key = "CardFee" Then
       GetCurCard_Statu = 1
    Else
        GetCurCard_Statu = 11
    End If
End Function
Private Function GetCardFeeSaveDataToCollect(ByVal objPati As clsPatientInfo, ByVal objCardFeeItems As clsBalanceItems, ByRef cllCardFee_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���濨�����ݼ�
    '���:objCardFeeItems-��ǰ������Ϣ
    '����:cllCardFee_Out-��ǰ��������
    '        |-Row:(���ѵ��ݺ�,���,�۸񸸺�,��������,�շ����,�շ�ϸĿid,������Ŀid,��׼����,�վݷ�Ŀ,Ӧ�ս��,ʵ�ս��,���˿���id,��������id,���˲���id,
    '                                 ִ�в���id,�Ƿ�����,���ձ���,������Ŀ��,ͳ����,ժҪ),Key="_" & ���
    '
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-10 19:50:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney  As Double, dblӦ�� As Double, dblʵ�� As Double, lngִ�в���ID As Long
    Dim cllRow As Collection, int��� As Integer
    Dim rs���� As ADODB.Recordset, rs������ As ADODB.Recordset
    
    On Error GoTo errHandle
    
    Set cllCardFee_Out = New Collection
    
    If fra�ſ�.Visible = False Or tbSendCard.SelectedItem.Key <> "CardFee" Then GetCardFeeSaveDataToCollect = True: Exit Function
    
    If objCardFeeItems Is Nothing Then Exit Function
    
    Set rs���� = GetCardFee
    If rs���� Is Nothing Then Exit Function
    If objCardFeeItems.���ݺ� = "" Then Exit Function
    '          |--cardfeelists:key="_cardfeelists"
    '               |---cardfeelist:(���ѵ��ݺ�,���,�۸񸸺�,��������,�շ����,�շ�ϸĿid,������Ŀid,��׼����,�վݷ�Ŀ,Ӧ�ս��,ʵ�ս��,���˿���id,��������id,���˲���id,
    '                                 ִ�в���id,�Ƿ�����,���ձ���,������Ŀ��,ͳ����,ժҪ) ,Key="_" & ���

   
    dblӦ�� = IIf(mCurSendCard.bln��� = False, mCurSendCard.dblӦ�ս��, StrToNum(txt����.Text))
    dblʵ�� = StrToNum(txt����.Text)
     
    int��� = 1
    
    '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
     lngִ�в���ID = zlGetCardFeeExcuteDeptID(Val(Nvl(rs����!�շ�ϸĿID)), Val(Nvl(rs����!���ұ�־)), UserInfo.����ID)
 
    Set cllRow = New Collection
    cllRow.Add Array("���ѵ��ݺ�", objCardFeeItems.���ݺ�), "_" & "���ѵ��ݺ�"
    cllRow.Add Array("���", int���), "_" & "���"
    cllRow.Add Array("�۸񸸺�", 0), "_" & "�۸񸸺�"
    cllRow.Add Array("��������", 0), "_" & "��������"
    cllRow.Add Array("�շ����", Nvl(rs����!�շ����)), "_" & "�շ����"
    cllRow.Add Array("�շ�ϸĿID", Nvl(rs����!�շ�ϸĿID)), "_" & "�շ�ϸĿID"
    cllRow.Add Array("������ĿID", Nvl(rs����!������ĿID)), "_" & "������ĿID"
    cllRow.Add Array("��׼����", dblӦ��), "_" & "��׼����"
    cllRow.Add Array("�վݷ�Ŀ", Nvl(rs����!�վݷ�Ŀ)), "_" & "�վݷ�Ŀ"
    
    cllRow.Add Array("Ӧ�ս��", dblӦ��), "_" & "Ӧ�ս��"
    cllRow.Add Array("ʵ�ս��", dblʵ��), "_" & "ʵ�ս��"
    cllRow.Add Array("���˿���id", UserInfo.����ID), "_" & "���˿���ID"
    cllRow.Add Array("��������id", UserInfo.����ID), "_" & "��������ID"
    cllRow.Add Array("���˲���id", UserInfo.����ID), "_" & "���˲���ID"
    cllRow.Add Array("ִ�в���id", lngִ�в���ID), "_" & "ִ�в���ID"
    cllRow.Add Array("�Ƿ�����", 0), "_" & "�Ƿ�����"
    cllRow.Add Array("���ձ���", ""), "_" & "���ձ���"
    cllRow.Add Array("������Ŀ��", 0), "_������Ŀ��" & ""
    cllRow.Add Array("ͳ����", 0), "_" & "ͳ����"
    cllRow.Add Array("ժҪ", ""), "_" & "ժҪ"
    cllRow.Add Array("�Ӱ��־", IIf(OverTime(), 1, 0)), "_" & "�Ӱ��־"
     
    cllCardFee_Out.Add cllRow, "_" & int���
    GetCardFeeSaveDataToCollect = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetClassMoney(ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ,��ʼ��֧�����(�շ����,ʵ�ս��)
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        '58322
        If .State = adStateOpen Then .Close
        .Fields.Append "�շ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
        .ActiveConnection = Nothing
        If StrToNum(txtԤ����.Text) <> 0 Then
            .AddNew
            !�շ���� = "Ԥ��"
            !��� = StrToNum(txtԤ����.Text)
            .Update
        End If
        
        If mCurSendCard.objSendCard.�ӿ���� <> 0 And cbo��������.Enabled And cbo��������.Visible Then
            .AddNew
            If Not mCurSendCard.rs���� Is Nothing Then !�շ���� = mCurSendCard.rs����!�շ����
            !��� = StrToNum(txt����.Text)
            .Update
        End If
    End With
    GetClassMoney = True
End Function


Private Sub RestorePayStyle()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ָ����ϴ�ѡ���֧����ʽ
    '˵��:lbl�ϼ�.Tag��¼�����ϴ�ѡ���֧����ʽ
    '       cboԤ������.Tag��¼����Ԥ�����ȱʡ֧����ʽ
    '       cbo���㷽ʽ.Tag��¼���ǿ��ѵ�ȱʡ֧����ʽ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intDeposit As Integer, intCardFee As Integer
    Dim varTemp As Variant
    
    On Error GoTo errHandle

    If mobjShowTotalMoneyControl.Tag = "" Then Exit Sub
    varTemp = Split(mobjShowTotalMoneyControl.Tag & "|", "|")
    intDeposit = varTemp(0): intCardFee = varTemp(1)
    mobjShowTotalMoneyControl.Tag = ""
    
    '�ָ�Ԥ������㷽ʽ
        
    If cboԤ������.Visible And cboԤ������.Enabled Then
        If intDeposit > cboԤ������.ListCount - 1 Then
            cboԤ������.ListIndex = Val(cboԤ������.Tag)
        Else
            cboԤ������.ListIndex = intDeposit
        End If
    End If
    '�ָ����ѽ��㷽ʽ
    If cbo��������.Visible And cbo��������.Enabled And chk����.value = 0 Then
        If intCardFee > cbo��������.ListCount - 1 Then
            cbo��������.ListIndex = Val(cbo��������.Tag)
        Else
            cbo��������.ListIndex = intCardFee
        End If
    End If
    
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub mbtQRCodePay_zlErrShow(ByVal strErrMsg As String, ByVal lngErrNum As Long)
    Call RestorePayStyle '�ָ��ϴ�ѡ���֧����ʽ
    If strErrMsg = "" Then Exit Sub
    MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
End Sub

Private Sub mbtQRCodePay_zlGetPayMoney(dblMoney As Double, strExpend As String, blnCancel As Boolean)
    Dim dblDeposit As Double, dblCardFee As Double
    
    Err = 0: On Error GoTo errHandle:
    
    mobjShowTotalMoneyControl.Tag = cboԤ������.ListIndex & "|" & cbo��������.ListIndex  '��¼��ǰ֧����ʽ��Index
    '��λ��ָ�������
    If mbtQRCodePay.Tag = "" Then
        MsgBox "δ�ҵ���Ч��ɨ�븶���,����!", vbInformation + vbOKOnly, gstrSysName
        blnCancel = True
        Exit Sub
    End If

    If fraԤ��.Visible = False And fra�ſ�.Visible = False Then
        MsgBox "û����Ҫɨ�븶�ķ���,����Ҫ����ɨ�븶��!", vbInformation + vbOKOnly, gstrSysName
        blnCancel = True: Exit Sub
    End If
    
    If fraԤ��.Visible And StrToNum(txtԤ����.Text) > 0 Then
        dblDeposit = StrToNum(txtԤ����.Text)
    End If
    
    If fra�ſ�.Visible And chk����.value = 0 And Val(txt����.Text) > 0 And txt����.Enabled Then
        dblCardFee = StrToNum(txt����.Text)
    End If
    
    '��ȡɨ�븶���
    dblMoney = dblDeposit + dblCardFee
    
     If dblMoney < 0 Then
        MsgBox "ɨ��֧�����Ϊ����������!", vbInformation + vbOKOnly, gstrSysName
        blnCancel = True
        Exit Sub
    End If
    
    If dblMoney = 0 Then
        MsgBox "û����Ҫɨ�븶�ķ���,����Ҫ����ɨ�븶��!", vbInformation + vbOKOnly, gstrSysName
        blnCancel = True
        zlControl.ControlSetFocus txtԤ����
        Exit Sub
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    blnCancel = True
End Sub

Private Sub mbtQRCodePay_zlQRCodePayment(ByVal lngCardTypeID As Long, ByVal strPayMentQRCode As String, ByVal strExpendXML As String, blnCancel As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ɨ�븶��
    '���:lngCardTypeID-�����ID
    '       strPayMentQRCode-��ά�븶������
    '       strExpendXML-����
    '����:strExpendXML-����
    '        blnCancel-true��ʾȡ������ɨ�븶,False-��ʾ����ɨ�븶�ɹ�
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle

    If lngCardTypeID = 0 Or blnCancel Then
        blnCancel = True
        Call RestorePayStyle '�ָ��ϴ�ѡ���֧����ʽ
        Exit Sub
    End If

    blnCancel = False
    If LocatePayStyle(lngCardTypeID) = False Then   '��λ��ɨ�븶��ָ�������
        blnCancel = True
        MsgBox "������Чʶ��ǰɨ�븶����𣬿��ܱ�����֧�ָ�����ɨ�븶���������Ա��ϵ��", vbInformation + vbOKOnly, gstrSysName
        Call RestorePayStyle '�ָ��ϴ�ѡ���֧����ʽ
        Exit Sub
    End If
    mstrQRcode = strPayMentQRCode
    RaiseEvent ExcuteQRCodePayment
    mstrQRcode = ""
    Call RestorePayStyle  '�ָ��ϴ�ѡ���֧����ʽ
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    blnCancel = True
    Call RestorePayStyle '�ָ��ϴ�ѡ���֧����ʽ
End Sub

Private Function LocatePayStyle(ByVal lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɨ�븶ʱ,���ݿ����ID,��λ��ָ����֧�������
    '���:lngCardTypeID-ɨ��Ŀ����ID
    '����:True-��λ��ָ����֧�����ɹ���False-��λ��ָ����֧�����ʧ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnFindDeposit As Boolean, blnFindCardFee As Boolean, i As Integer
    Dim objCard As Card
    If lngCardTypeID = 0 Then Exit Function
    
    With cboԤ������
        If .Visible And .Enabled Then
            For i = 1 To mobjDepositPayCards.Count
                Set objCard = mobjDepositPayCards(i)
                If objCard.�ӿ���� = lngCardTypeID Then
                    If .ListCount >= i Then .ListIndex = i - 1: blnFindDeposit = True: Exit For
                End If
            Next
        Else
            blnFindDeposit = True
        End If
    End With
    
    With cbo��������
        If .Visible And .Enabled And chk����.value = 0 Then
            For i = 1 To mobjCardFeePayCards.Count
                Set objCard = mobjCardFeePayCards(i)
                If objCard.�ӿ���� = lngCardTypeID Then
                
                    If .ListCount >= i Then .ListIndex = i - 1: blnFindCardFee = True: Exit For
                End If
            Next
        Else
            blnFindCardFee = True
        End If
    End With
    LocatePayStyle = blnFindDeposit And blnFindCardFee
End Function


Private Sub Local���㷽ʽ(ByVal lng�����ID As Long, Optional blnԤ�� As Boolean = True, Optional ByVal str���㷽ʽ As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��λ���㷽ʽ
    '����:���˺�
    '����:2011-07-26 15:32:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCards As Cards, cboPay As ComboBox
    Dim i As Long, objCard As Card
    If mblnNotClick Then Exit Sub
    
    If blnԤ�� Then
       Set objCards = mobjDepositPayCards
        Set cboPay = cboԤ������
    Else
       Set objCards = mobjCardFeePayCards
        Set cboPay = cbo��������
    End If
    
    If objCards Is Nothing Then Exit Sub
    
    With cboPay
        mblnNotClick = True
        For i = 0 To .ListCount - 1
            Set objCard = objCards(i + 1)
            
            ''��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
            If lng�����ID > 0 Then
                If objCard.�ӿ���� = lng�����ID Then
                    .ListIndex = i: Exit For
                End If
            Else
                If objCard.���㷽ʽ = str���㷽ʽ Then
                    .ListIndex = i: Exit For
                End If
            End If
        Next
        mblnNotClick = False
    End With
End Sub
Public Property Get RealName() As Boolean
       RealName = mblnRealName
End Property

Public Property Let RealName(ByVal vNewValue As Boolean)
    mblnRealName = vNewValue
    
    If Not mobjPubPatient.blnRealName Then
        Exit Property
    End If
    If Not mobjPati Is Nothing Then mobjPati.ʵ����֤ = mblnRealName
    'δ����ʵ��֤��,ֻ������ʱ��
    chkEndTime.value = IIf(mblnRealName, 0, 1)
    chkEndTime.Enabled = Trim(txt����.Text) <> "" And mblnRealName
End Property

Public Property Get GetWidth() As Long
       GetWidth = Me.Width
End Property

Public Property Get GetHeight() As Long
       GetHeight = Me.Height
End Property

Public Sub LoadԤ�����㷽ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ����֧����ʽ
    '����:���˺�
    '����:2019-11-26 14:27:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, str���� As String
    Dim i As Long
    
    str���� = "1,2,8" & IIf(mblnAllowInsureAccDeposit, ",3", "")
    If mobjThirdSwap.zlGetBalanceModeCards(mobjDepositPayCards, , , , mstrQRCodeTypeIds_Deposit, "Ԥ����", str����) = False Then Set mobjDepositPayCards = New Cards
    With cboԤ������
        .Clear
        mblnNotClick = True
        For i = 1 To mobjDepositPayCards.Count
            Set objCard = mobjDepositPayCards(i)
            .AddItem objCard.����
            .ItemData(.NewIndex) = objCard.��������
            If objCard.ȱʡ��־ = 1 Then .ListIndex = i
        Next
        If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
        .Enabled = .ListCount > 0
        mblnNotClick = False
    End With
    lblStyle.Tag = mstrQRCodeTypeIds_Deposit
End Sub
Public Sub Load���ѽ��㷽ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؿ���֧����ʽ
    '����:���˺�
    '����:2019-11-26 14:27:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim i As Long, str���� As String
    
    str���� = "1,2,8"
    If mobjThirdSwap.zlGetBalanceModeCards(mobjCardFeePayCards, , , , mstrQRCodeTypeIds_CardFee, "���￨", str����) = False Then Set mobjDepositPayCards = New Cards
    With cbo��������
        .Clear
        mblnNotClick = True
        For i = 1 To mobjCardFeePayCards.Count
            Set objCard = mobjCardFeePayCards(i)
            .AddItem objCard.����
            .ItemData(.NewIndex) = objCard.��������
            If objCard.ȱʡ��־ = 1 Then .ListIndex = i
        Next
        If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
        .Enabled = .ListCount > 0
        mblnNotClick = False
    End With
    lbl���㷽ʽ.Tag = mstrQRCodeTypeIds_CardFee
End Sub

Private Function Load֧����ʽ(Optional ByVal bln��ʾԤ�� As Boolean = True, Optional ByVal bln��ʾ���� As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����֧����ʽ
    '���:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-26 15:04:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strQRCodeTypeIDs As String, varDeposit As Variant, varCardFee As Variant, varTemp As Variant
    Dim strQRCardTypeIDs As String, strErrMsg As String
    Dim i As Long, j As Long
    Call LoadԤ�����㷽ʽ
    Call Load���ѽ��㷽ʽ
    
    If cboԤ������.ListCount = 0 Then
        MsgBox "Ԥ������û�п��õĽ��㷽ʽ,���ȵ����㷽ʽ���������á�", vbExclamation, gstrSysName
        Exit Function
    End If
    If Not mbtQRCodePay Is Nothing Then
        
        If mstrQRCodeTypeIds_Deposit <> mstrQRCodeTypeIds_CardFee Then
            varDeposit = Split(mstrQRCodeTypeIds_Deposit & ",", ",")
            varCardFee = Split(mstrQRCodeTypeIds_CardFee & ",", ",")
            For i = 0 To UBound(varDeposit)
                For j = 0 To UBound(varCardFee)
                    If varCardFee(j) = varDeposit(i) Then
                        strQRCodeTypeIDs = strQRCodeTypeIDs & "," & varDeposit(i)
                        Exit For
                    End If
                Next
            Next
            If strQRCodeTypeIDs <> "" Then strQRCodeTypeIDs = Mid(strQRCodeTypeIDs, 2)
        Else
            strQRCodeTypeIDs = mstrQRCodeTypeIds_Deposit
        End If
        If strQRCodeTypeIDs <> "" Then mbtQRCodePay.Tag = strQRCodeTypeIDs
            
        '��ʼ��ɨ��ؼ�
        If bln��ʾԤ�� And bln��ʾ���� Then
            strQRCardTypeIDs = mbtQRCodePay.Tag
        ElseIf bln��ʾԤ�� And Not bln��ʾ���� Then
            strQRCardTypeIDs = lblStyle.Tag
        ElseIf Not bln��ʾԤ�� And bln��ʾ���� Then
            strQRCardTypeIDs = lbl���㷽ʽ.Tag
        End If
        
        If mbtQRCodePay.zlInit(Me, strQRCardTypeIDs, glngSys, mlngModule, gcnOracle, gstrDBUser, strErrMsg) = False Then strQRCardTypeIDs = ""
        mbtQRCodePay.Tag = strQRCardTypeIDs
        mbtQRCodePay.Visible = strQRCardTypeIDs <> "" Or mblnShowDepositAndSendCard
        mbtQRCodePay.Enabled = strQRCardTypeIDs <> ""
        mobjShowTotalMoneyControl.Visible = strQRCardTypeIDs <> "" Or mblnShowDepositAndSendCard
        
    End If
    Load֧����ʽ = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function UpdateDepositBlncInfo(ByVal int����״̬ As Integer, ByVal objPati As clsPatientInfo, _
    ByVal objDepositItems As clsBalanceItems, ByVal cllExpendInfo As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¿�����ؽ�����Ϣ
    '���:int����״̬:0-��ɽ���;1-�ӿڵ���ǰ����;2-�ӿڵ��ú�����
    '     objDepositItems-��ǰԤ��֧����ʽ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-14 11:49:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllUpdateFeeData As Collection, cllTemp As Collection
    Dim objCurItem As clsBalanceItem, blnTrans As Boolean
    Dim strDepositNo As String, strCardFeeNo As String, strErrMsg As String
    Dim varTemp As Variant, lng�䶯id As Long, lngԤ��ID As Long
    Dim cllPro As Collection, strSql As String, int�쳣״̬ As Integer
    Dim cllErrData As Collection
    
    On Error GoTo errHandle
    
    Set cllUpdateFeeData = New Collection
    Set cllTemp = New Collection
     
    If objDepositItems Is Nothing Then Exit Function
    If objDepositItems.Count = 0 Then Exit Function
    
    
    Set objCurItem = objDepositItems(1)
    strDepositNo = objDepositItems.���ݺ�
    lngԤ��ID = objCurItem.Ԥ��ID
    
    
    cllTemp.Add Array("Ԥ������", strDepositNo), "_" & "Ԥ������"
    cllTemp.Add Array("Ԥ��ID", lngԤ��ID), "_" & "Ԥ��ID"
    cllTemp.Add Array("��Ʊ��", txtFact.Text), "_" & "��Ʊ��"
    cllTemp.Add Array("����ID", mobjDepositFact.����ID), "_" & "����ID"
    cllTemp.Add Array("����ID", objPati.����ID), "_" & "����ID"
    cllTemp.Add Array("����Ա���", UserInfo.���), "_" & "����Ա���"
    cllTemp.Add Array("����Ա����", UserInfo.����), "_" & "����Ա����"
    cllTemp.Add Array("�տ�ʱ��", Format(objCurItem.����ʱ��, "yyyy-mm-dd HH:MM:SS")), "_" & "�տ�ʱ��"
    cllUpdateFeeData.Add cllTemp, "_billinfo"
    
     '������Ϣ
    Set cllTemp = New Collection
    cllTemp.Add Array("���㷽ʽ", objCurItem.���㷽ʽ), "_" & "���㷽ʽ"
    cllTemp.Add Array("�������", objCurItem.�������), "_" & "�������"
    cllTemp.Add Array("�����ID", IIf(objCurItem.���ѿ�, 0, objCurItem.�����ID)), "_" & "�����ID"
    cllTemp.Add Array("���㿨���", IIf(objCurItem.���ѿ�, objCurItem.�����ID, 0)), "_" & "���㿨���"
    cllTemp.Add Array("����", objCurItem.����), "_" & "����"
    cllTemp.Add Array("������ˮ��", objCurItem.������ˮ��), "_" & "������ˮ��"
    cllTemp.Add Array("����˵��", objCurItem.����˵��), "_" & "����˵��"
    cllTemp.Add Array("ժҪ", objCurItem.����ժҪ), "_" & "ժҪ"
    cllTemp.Add Array("������λ", ""), "_" & "������λ"
    
    If Not cllExpendInfo Is Nothing Then
        cllTemp.Add Array("������Ϣ��", cllExpendInfo), "_" & "������Ϣ��"
    End If
    cllUpdateFeeData.Add cllTemp, "_balanceinfo"

    '   cllUpdateDate-�޸ĵĽ�������
    '         |--billinfo-������Ϣ,"_billinfo"
    '              |-Ԥ������,Ԥ��ID,����Ա���,����Ա����,�տ�ʱ��,����ID,��Ʊ�ţ�����ID)
    '         |--balanceinfo-������Ϣ,"_balanceinfo"
    '                |--(���㷽ʽ,�������,�����id,���㿨���,����,������ˮ��,����˵��,ժҪ,������λ)
    '                |--������Ϣ��,
    '                |-----������Ϣ:��������,��������
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    
    'ͬ��״̬����������=2,3ʱ��0��NULL������¼;-1-δ��������;1-δ���ýӿ�;2-�ӿڵ��óɹ�,3-���ý��������ɹ�;4-ҽ�ƿ���Ϣ�����ɹ�"
    If int����״̬ = 0 Then
         int�쳣״̬ = 2
        If Not GetDelErrDataToColl(objDepositItems.ҵ��ID, objDepositItems.�쳣ID, cllErrData) Then Exit Function
    ElseIf int����״̬ = 1 Then
        If Not GetUpdateErrDataSyncTagToColl(objDepositItems.�쳣ID, 1, cllErrData) Then Exit Function
        int�쳣״̬ = 1
    Else
        If Not GetUpdateErrDataSyncTagToColl(objDepositItems.�쳣ID, 3, cllErrData) Then Exit Function
        int�쳣״̬ = 1
        '0-������¼,1-����״̬�����½���˵����2-ɾ���쳣����
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    If Zl_���˽����쳣��¼_Modify(int�쳣״̬, cllErrData) = False Then
        gcnOracle.RollbackTrans: blnTrans = False
    End If
    
    If mobjExseSvr.Zl_Exsesvr_Upddepositblncinfo(int����״̬, cllUpdateFeeData, False, strErrMsg) = False Then
        gcnOracle.RollbackTrans: blnTrans = False
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    UpdateDepositBlncInfo = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SetLoaclePayModefromCard(ByVal objCard As Card, ByVal blnԤ�� As Boolean, Optional blnAppend As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݿ���������ȱʡ��֧����ʽ
    '���:objCard-��ǰ������
    '     blnAppend-δ�ҵ����Զ�����
    '���� :��λ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2019-11-13 10:10:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objTemp As Card, blnFind As Boolean
    Dim objCombox As ComboBox, objPayCards As Cards
    If objCard Is Nothing Then Exit Function
    
    mblnNotClick = True
    If blnԤ�� Then
        Set objCombox = cboԤ������: Set objPayCards = mobjDepositPayCards
    Else
        Set objCombox = cbo��������: Set objPayCards = mobjCardFeePayCards
    End If
    
    blnFind = False
    For i = 0 To objCombox.ListCount - 1
        If blnԤ�� Then
            Set objTemp = GetDepositPayCard(i)
        Else
            Set objTemp = GetCardFeePayCard(i)
        End If
        
        If objTemp Is Nothing Then Exit Function
        If objTemp.�ӿ���� = objCard.�ӿ���� And objTemp.���㷽ʽ = objCard.���㷽ʽ Then
            blnFind = True: objCombox.ListIndex = i: Exit For
        End If
    Next
    If Not blnFind And blnAppend Then
        'δ�ҵ�
        objCombox.AddItem objCard.���㷽ʽ
        objCombox.ItemData(objCombox.NewIndex) = objCombox.ListCount + 1
        objCombox.ListIndex = objCombox.NewIndex
        objPayCards.Add objCard, "K" & objCombox.ListCount + 1
        blnFind = True
    End If
    mblnNotClick = False
    SetLoaclePayModefromCard = blnFind
End Function

Private Function ReadDepositBalanceDataFromDepositNo(ByVal strNO As String, lng�쳣ID As Long, ByVal lngҵ��ID As Long, intͬ��״̬ As Integer, _
    ByRef objBalanceItems_Out As clsBalanceItems, Optional bln���� As Boolean, Optional ByVal str�쳣������Ϣ As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡԤ����������
    '���:str�쳣������Ϣ-��ǰ�쳣������Ϣ
    '����:objBalanceItems_out-Ԥ����������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-28 10:49:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCurItem As clsBalanceItem
    Dim objPayCard As Card, cllDeposit As Collection, int���� As Integer
    Dim cllSwapinfo As Collection, cllExpends As Collection
    Dim i As Long
    
    On Error GoTo errHandle
    
    Set objBalanceItems_Out = New clsBalanceItems
    objBalanceItems_Out.���ݺ� = strNO
    
    '����Ԥ�����㷽ʽ
    If mobjExseSvr.zl_ExseSvr_GetDepositInfo(strNO, IIf(bln����, 2, 3), cllDeposit, True, "") = False Then Exit Function
    
    Set objCurItem = New clsBalanceItem
   '����:cll_Deposit_Out-����Ԥ��Ʊ�����ݼ�,key="_"+����
    '       |-����ID,��ҳid,Ԥ��ID,Ԥ�����ݺ�,��Ʊ��,Ԥ����� ,�ɿ����id,�ɿ���,�ɿλ ,��λ������,�������˺�,ժҪ,����Ա����,����Ա���,�տ�ʱ�� ,���㷽ʽ,�������,
    '       | �����id,���㿨��� ,���ѿ�ID,֧������,������ˮ��,����˵��,������λ,����״̬ ,��������ID,����,ҽ����,ҽ������
    If Val(cllDeposit("_�����ID")) <> 0 Then
        If mobjOneCardComLib.zlGetCard(cllDeposit("_�����ID"), False, objPayCard) = False Then Exit Function
        int���� = 3
    ElseIf Val(cllDeposit("_���㿨���")) <> 0 Then
         If mobjOneCardComLib.zlGetCard(cllDeposit("_���㿨���"), True, objPayCard) = False Then Exit Function
         int���� = 5
    Else
        Set objPayCard = zlGetCardFromBalanceName(cllDeposit("_���㷽ʽ"))   '��ͨ�Ľ��㷽ʽ
        int���� = 0
        If objPayCard.�������� = 3 Then int���� = 2
    End If
    objBalanceItems_Out.���� = int����
    With objCurItem
        Set .objCard = objPayCard
        .���㷽ʽ = cllDeposit("_���㷽ʽ")
        .������� = cllDeposit("_�������")
        .�������� = 1 '' 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ�
        .��������ID = Val(cllDeposit("_��������ID"))
        .���ݺ� = Trim(cllDeposit("_Ԥ�����ݺ�"))
        .Ԥ��ID = Val(cllDeposit("_Ԥ��ID"))
        .�쳣ID = lng�쳣ID
        .������ = Val(cllDeposit("_�ɿ���"))
        .�������� = int���� ''0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        .����ʱ�� = CDate(cllDeposit("_�տ�ʱ��"))
        .�������� = objPayCard.��������
        .����ժҪ = cllDeposit("_ժҪ")
        .�����ID = Val(cllDeposit("_�����ID"))
        If objPayCard.���ѿ� Then
            .�����ID = Val(cllDeposit("_���㿨���"))
            .���ѿ� = True
            .���ѿ�ID = Val(cllDeposit("_���ѿ�ID"))
        End If
        .���� = cllDeposit("_֧������")
        .������ˮ�� = Trim(cllDeposit("_������ˮ��"))
        .����˵�� = Trim(cllDeposit("_����˵��"))
        .У�Ա�־ = Trim(cllDeposit("_����״̬"))
        .���� = cllDeposit("_ҽ������")
        .�Ƿ�Ԥ�� = True
        .�Ƿ񱣴� = True
        .�Ƿ���� = (.У�Ա�־ = 2 Or .У�Ա�־ = 0)
        .�Ƿ�����༭ = Not .�Ƿ����
        .�Ƿ�����ɾ�� = .�Ƿ�����༭
        .�Ƿ��������� = .�Ƿ�����༭
    End With
  
    tbDeposit.Visible = False
    If Val(cllDeposit("_Ԥ�����")) <= 1 Then
        mbln����Ԥ�� = True: mblnסԺԤ�� = False
        fraԤ��.Caption = "������Ԥ����Ϣ��"
    Else
        mblnסԺԤ�� = True: mbln����Ԥ�� = False
        fraԤ��.Caption = "��סԺԤ����Ϣ��"
    End If
    
    mintInsure = Val(cllDeposit("_����"))
    mstrҽ���� = Trim(cllDeposit("_ҽ����"))
    mstr���� = cllDeposit("_ҽ������")
    
    objBalanceItems_Out.AddItem objCurItem
    objBalanceItems_Out.������ = objCurItem.������
    objBalanceItems_Out.�Ƿ񱣴� = True
    objBalanceItems_Out.������� = IIf(objCurItem.У�Ա�־ = 0, True, False)
    objBalanceItems_Out.ͬ��״̬ = intͬ��״̬
    objBalanceItems_Out.ҵ��ID = lngҵ��ID
    objBalanceItems_Out.�쳣ID = lng�쳣ID
    objBalanceItems_Out.�Ƿ񱣴� = True
    If str�쳣������Ϣ <> "" And intͬ��״̬ <= 2 Then
         If GetErrSwapInfoByJsonString(str�쳣������Ϣ, cllSwapinfo, cllExpends) Then
            '�쳣����
             'cllSwapinfo(����,�����ID,������ˮ��,����˵��,���׽��,��ά��,֧����ʽ,����ժҪ)
            For i = 1 To objBalanceItems_Out.Count
               If cllSwapinfo("_֧����ʽ")(1) <> "" Then objBalanceItems_Out(i).���㷽ʽ = cllSwapinfo("_֧����ʽ")(1)
               objBalanceItems_Out(i).�����ID = cllSwapinfo("_�����ID")(1)
               objBalanceItems_Out(i).������ˮ�� = cllSwapinfo("_������ˮ��")(1)
               objBalanceItems_Out(i).����˵�� = cllSwapinfo("_����˵��")(1)
               objBalanceItems_Out(i).���� = cllSwapinfo("_����")(1)
               objBalanceItems_Out(i).����ժҪ = cllSwapinfo("_����ժҪ")(1)
               objBalanceItems_Out(i).QRCode = cllSwapinfo("_��ά��")(1)
            Next
         End If
    End If
    
    '��ʼ������
    txtԤ����.Text = Format(objBalanceItems_Out.������, "0.00")
    
    mblnNotClick = True
    Call SetLoaclePayModefromCard(objCurItem.objCard, True, True): mblnNotClick = False
    
    txt�������.Text = objCurItem.�������
    If mintInsure = 0 Then
        mblnNotClick = True
        txt�ɿλ.Text = Trim(cllDeposit("_�ɿλ"))
        txt������.Text = Trim(cllDeposit("_��λ������"))
        txt�ʺ�.Text = Trim(cllDeposit("_�������˺�"))
        chk��λ�ɿ�.value = IIf(txt�ɿλ.Text <> "", 1, 0)
        mblnNotClick = False
    Else
        chk��λ�ɿ�.value = 0
    End If
    Call RefreshFactNo      'ˢ�·�Ʊ��
    ReadDepositBalanceDataFromDepositNo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardFeeDataFromColl(ByVal strNO As String, ByVal cllCardFee As Collection, _
    ByRef rsCardFee_Out As Recordset, Optional ByRef objBalanceItems_Out As clsBalanceItems, Optional ByRef dblMoney_Out As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ��񷵻صļ���,����¼����ʽ������Ϣ
    '���:cllCardFee-��ǰ����
    '
    '����:rsCardFee_Out-���صĿ����ü���
    '     objBalanceItems_out-������Ϣ�б���Ҫ�ǿ��ܴ��ڼ��ʣ���Ҫ��objBalanceItems_out
    '     dblMoney_Out:ʵ�ս��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection
    Dim i As Long, bln���� As Boolean
    
    On Error GoTo errHandle
    
    If objBalanceItems_Out Is Nothing Then Set objBalanceItems_Out = New clsBalanceItems
    dblMoney_Out = 0
    Set rsCardFee_Out = New ADODB.Recordset
    With rsCardFee_Out
        If .State = adStateOpen Then .Close
        
        .Fields.Append "���ݺ�", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "����id", adBigInt, , adFldIsNullable
        .Fields.Append "���", adBigInt, , adFldIsNullable
        .Fields.Append "����id", adBigInt, , adFldIsNullable
        .Fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "�Ա�", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�ѱ�", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�շ���Ŀid", adBigInt, , adFldIsNullable
        .Fields.Append "������Ŀid", adBigInt, , adFldIsNullable
        .Fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "Ӧ�ս��", adDouble, , adFldIsNullable
        .Fields.Append "ʵ�ս��", adDouble, , adFldIsNullable
        
        .Fields.Append "������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����Ա���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����Ա����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "�Ǽ�ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��¼״̬", adBigInt, , adFldIsNullable
        
        .Fields.Append "�Ƿ�����", adBigInt, , adFldIsNullable
        .Fields.Append "��Ʊ��", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "�Ƿ����", adBigInt, , adFldIsNullable
        .Fields.Append "����״̬", adBigInt, , adFldIsNullable
        .Fields.Append "�����ID", adBigInt, , adFldIsNullable
        .Fields.Append "����", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "�Ƿ�Һŷ���", adBigInt, , adFldIsNullable
        
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    If cllCardFee Is Nothing Then Exit Function
    '    fee_id  N   1   ����id
    '    fee_num N   1   ���
    '    pati_id N   1   ����id
    '    pati_name   C   1   ����
    '    pati_sex    C   1   �Ա�
    '    pati_age    C   1   ����
    '    fee_category    C   1   �ѱ�
    '    item_id N   1   �շ���Ŀid
    '    income_item_id  N   1   ������Ŀid
    '    quantity    N   1   ����
    '    fee_amrcvb  N   1   Ӧ�ս��
    '    fee_ampaid  N   1   ʵ�ս��
    '    placer  C   1   ������
    '    operator_code   C   1   ����Ա���
    '    operator_name   C   1   ����Ա����
    '    create_time D   1   �Ǽ�ʱ��
    '    happen_time D   1   ����ʱ��
    '    rec_status  N   1   ��¼״̬
    '    mrbkfee_sign N   1   �Ƿ�����:1-�ǲ�����;0-���ǲ�����
    '    invoice_no  N   1   ��Ʊ��
    '    kpbooks_sign N   1   ���ʱ�־:1-�Ǽ���;0-����
    '    fee_status   N   1   ����״̬:1-�쳣״̬;0-��������
    '    cardtype_id N   1   �����ID
    '    card_no C   1   ����
    '    sendcard_reg    N   1   �Ƿ�ҹҺ�ͬ������:1-�ǹҺ�ͬʱ����;0-�ǹҺ�ͬʱ����

    For i = 1 To cllCardFee.Count
        Set cllTemp = cllCardFee(i)
        
        If Not bln���� Then bln���� = Val(Nvl(cllTemp("_kpbooks_sign"))) = 1
        With rsCardFee_Out
            .AddNew
            !���ݺ� = strNO
            !����id = Val(Nvl(cllTemp("_fee_id")))
            !��� = Val(Nvl(cllTemp("_fee_num")))
            !����ID = Val(Nvl(cllTemp("_pati_id")))
            !���� = Nvl(cllTemp("_pati_name"))
            !�Ա� = Nvl(cllTemp("_pati_sex"))
            !���� = Nvl(cllTemp("_pati_age"))
            !�ѱ� = Nvl(cllTemp("_fee_category"))
            !�շ���ĿID = Val(Nvl(cllTemp("_item_id")))
            !������ĿID = Val(Nvl(cllTemp("_income_item_id")))
            !���� = Val(Nvl(cllTemp("_quantity")))
            !Ӧ�ս�� = Val(Nvl(cllTemp("_fee_amrcvb")))
            !ʵ�ս�� = Val(Nvl(cllTemp("_fee_ampaid")))
            !������ = Nvl(cllTemp("_placer"))
            !����Ա��� = Nvl(cllTemp("_operator_code"))
            !����Ա���� = Nvl(cllTemp("_operator_name"))
            !�Ǽ�ʱ�� = Nvl(cllTemp("_create_time"))
            !����ʱ�� = Nvl(cllTemp("_happen_time"))
            !��¼״̬ = Val(Nvl(cllTemp("_rec_status")))
            
            !�Ƿ����� = Val(Nvl(cllTemp("_mrbkfee_sign")))
            !��Ʊ�� = Nvl(cllTemp("_invoice_no"))
            !�Ƿ���� = Val(Nvl(cllTemp("_kpbooks_sign")))
            !����״̬ = Val(Nvl(cllTemp("_fee_status")))
            !�����ID = Val(Nvl(cllTemp("_cardtype_id")))
            !���� = Nvl(cllTemp("_card_no"))
            !�Ƿ�Һŷ��� = Val(Nvl(cllTemp("_sendcard_reg")))
            .Update
            dblMoney_Out = RoundEx(dblMoney_Out + Val(Nvl(rsCardFee_Out!ʵ�ս��)), 5)
        End With
    Next
    If bln���� Then
        objBalanceItems_Out.���� = gEM_���ʵ�
    End If
    objBalanceItems_Out.������ = dblMoney_Out
    objBalanceItems_Out.���ݺ� = strNO
    objBalanceItems_Out.�Ƿ񱣴� = True
    Set cllTemp = Nothing
    zlGetCardFeeDataFromColl = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ReadCardFeeBalanceDataFromNo(ByVal strNO As String, lng�쳣ID As Long, ByVal lngҵ��ID As Long, intͬ��״̬ As Integer, _
    ByRef objBalanceItems_Out As clsBalanceItems, Optional bln���� As Boolean, Optional ByVal str�쳣������Ϣ As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ѽ�������
    '���:
    '����:objBalanceItems_out-���ѽ�������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-28 10:49:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCurItem As clsBalanceItem, cllCardFee As Collection, cllPriceBill As Collection, cllBalance As Collection
    Dim objPayCard As Card, int���� As Integer, dblMoney As Double
    Dim cllSwapinfo As Collection, cllExpends As Collection
    Dim i As Long
    On Error GoTo errHandle
    
    Set objBalanceItems_Out = New clsBalanceItems
    objBalanceItems_Out.���ݺ� = strNO
    
    '����Ԥ�����㷽ʽ
    '��ѯ���ͣ�0-��ȡ��������:1-��ȡ���ϵ���;2-ʣ����õ���
    If mobjExseSvr.zl_ExseSvr_GetCardFeeInfoByNo(strNO, IIf(bln����, 1, 0), cllCardFee, cllPriceBill, cllBalance, Nothing, , , False) = False Then Exit Function
    If zlGetCardFeeDataFromColl(strNO, cllCardFee, mrsCardFee, objBalanceItems_Out, dblMoney) = False Then Exit Function
    If zlGetBalanceItemsFromCardFeeColl(strNO, cllBalance, lng�쳣ID, objBalanceItems_Out, IIf(bln����, True, False)) = False Then Exit Function
    
    objBalanceItems_Out.ҵ��ID = lngҵ��ID
    objBalanceItems_Out.�쳣ID = lng�쳣ID
    objBalanceItems_Out.ͬ��״̬ = intͬ��״̬
    
    txt����.Text = objBalanceItems_Out.������
    If mrsCardFee.RecordCount <> 0 Then
        txt����.Text = Nvl(mrsCardFee!����)
    End If

    If objBalanceItems_Out.���� = gEM_���ʵ� Then
        '���ʵ�
        mblnSendCardLocked = True
        chk����.value = 1
        Call SetCardEditEnabled
        ReadCardFeeBalanceDataFromNo = True
        Exit Function
    End If
    
    If str�쳣������Ϣ <> "" And intͬ��״̬ <= 2 Then
         If GetErrSwapInfoByJsonString(str�쳣������Ϣ, cllSwapinfo, cllExpends) Then
            '�쳣����
             'cllSwapinfo(����,�����ID,������ˮ��,����˵��,���׽��,��ά��,֧����ʽ,����ժҪ)
            For i = 1 To objBalanceItems_Out.Count
               If cllSwapinfo("_֧����ʽ")(1) <> "" Then objBalanceItems_Out(i).���㷽ʽ = cllSwapinfo("_֧����ʽ")(1)
               objBalanceItems_Out(i).�����ID = cllSwapinfo("_�����ID")(1)
               objBalanceItems_Out(i).������ˮ�� = cllSwapinfo("_������ˮ��")(1)
               objBalanceItems_Out(i).����˵�� = cllSwapinfo("_����˵��")(1)
               objBalanceItems_Out(i).���� = cllSwapinfo("_����")(1)
               objBalanceItems_Out(i).����ժҪ = cllSwapinfo("_����ժҪ")(1)
               objBalanceItems_Out(i).QRCode = cllSwapinfo("_��ά��")(1)
            Next
            Set objBalanceItems_Out.objTag = cllExpends
         End If
    End If
        
        
    If intͬ��״̬ <> -1 Then
        Set objCurItem = objBalanceItems_Out(1)
        mblnNotClick = True
        Call SetLoaclePayModefromCard(objCurItem.objCard, False, True): mblnNotClick = False
    End If
    
    ReadCardFeeBalanceDataFromNo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceItemsFromCardFeeColl(ByVal strNO As String, ByVal cllCardFeeBalance As Collection, ByVal lng�쳣ID As Long, _
    ByRef objBalanceItems_Out As clsBalanceItems, _
    Optional ByVal bln�鿴���� As Boolean, Optional blnDelFee As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ��񷵻صļ���,����¼����ʽ���ؽ�����Ϣ
    '���:cllCardFeeBalance-��ǰ����
    '     strNo-���õ��ݺ�
    '     bln�鿴����-��ǰ���ĵ������ϵ���
    '     blnDelFee-��ǰΪ�˷Ѳ���
    '����:objBalanceItems_Out-���صĿ��ѽ�����Ϣ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection, i As Long
    Dim objItem As clsBalanceItem, objCard As Card
    Dim dbl���� As Double
    On Error GoTo errHandle
    
    If objBalanceItems_Out Is Nothing Then Set objBalanceItems_Out = New clsBalanceItems
    
    If cllCardFeeBalance Is Nothing Then Exit Function
    If objBalanceItems_Out.���� = gEM_���ʵ� Then zlGetBalanceItemsFromCardFeeColl = True: Exit Function
    
    
    
    '    blnc_mode   C   1   ���㷽ʽ����
    '    balance_id  N   1   ����ID
    '    blnc_money  N   1   ���ʽ��
    '    pay_cardno  N   1   ֧������
    '    pay_swapno  C   1   ������ˮ��
    '    pay_swapmemo    C   1   ����˵��
    '    relation_id N   1   ��������id
    '    cardtype_id N   1   �����id
    '    consume_card    N   1   �Ƿ����ѿ�:1-��;0-����
    '    blnc_nature N   1   ��������:1-�ֽ���㷽ʽ,2-������ҽ������ , 8-���㿨���� ,9-����
    '    blnc_statu  N   1   ����״̬:1-δ���ýӿ�;2-�ӿڵ��óɹ�,����δ�շ����,0-��������
    '    consume_card_id N   1   ���ѿ�id
    '    blnc_no C   1   �������
    '    blnc_memo   C   1   ժҪ
    
    objBalanceItems_Out.������ = 0
    For i = 1 To cllCardFeeBalance.Count
        Set cllTemp = cllCardFeeBalance(i)
        Set objItem = New clsBalanceItem
        Set objCard = GetCardFromCardType(Val(Nvl(cllTemp("_cardtype_id"))), Val(Nvl(cllTemp("_consume_card"))) = 1, Nvl(cllTemp("_blnc_mode")))
        If Val(Nvl(cllTemp("_blnc_nature"))) = 9 Then
            dbl���� = RoundEx(dbl���� + Val(Nvl(cllTemp("_blnc_money"))), 6)
        Else
            With objItem
                Set .objCard = objCard
                .���ݺ� = strNO
                .�������� = 5   ' 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ�
                .���㷽ʽ = Nvl(cllTemp("_blnc_mode"))
                .������ = Val(Nvl(cllTemp("_blnc_money")))
                .��������ID = Val(Nvl(cllTemp("_relation_id")))
                .������ˮ�� = Nvl(cllTemp("_pay_swapno"))
                .����˵�� = Nvl(cllTemp("_pay_swapmemo"))
                .������� = Nvl(cllTemp("_blnc_no"))
                .�������� = Val(Nvl(cllTemp("_blnc_nature")))
                .����ժҪ = Nvl(cllTemp("_blnc_memo"))
                .���� = Nvl(cllTemp("_pay_cardno"))
                
                .�����ID = Val(Nvl(cllTemp("_cardtype_id")))
                .���ѿ�ID = Val(Nvl(cllTemp("_consume_card_id")))
                .���ѿ� = Val(Nvl(cllTemp("_consume_card"))) = 1
                .�Ƿ����� = objCard.�������Ĺ��� <> ""
                .ԭʼ��� = .������
                .δ�˽�� = .������
                .У�Ա�־ = Val(Nvl(cllTemp("_blnc_statu")))
                .�Ƿ���� = .У�Ա�־ = 2 Or .У�Ա�־ = 0
                .�Ƿ�����༭ = Not .�Ƿ����
                .�Ƿ�����ɾ�� = .�Ƿ�����༭
                .�Ƿ��������� = .�Ƿ�����༭
                .���� = ""
                .�ʻ���� = 0
                If .�����ID = 0 Then   '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                      .�������� = 0
                ElseIf .�����ID <> 0 And .���ѿ� = False Then
                      .�������� = 3
                ElseIf .�����ID <> 0 And .���ѿ� Then
                      .�������� = 5
                Else
                     .�������� = 0
                End If
                .�Ƿ��˿� = blnDelFee
                If bln�鿴���� Then
                    .����ID = Val(Nvl(cllTemp("_balance_id")))
                    .����ID = Val(Nvl(cllTemp("_original_id"))) 'ԭ����ID
                   
                Else
                    .����ID = Val(Nvl(cllTemp("_balance_id")))
                    .����ID = Val(Nvl(cllTemp("_original_id"))) 'ԭ����ID
                End If
                .�쳣ID = lng�쳣ID
            
                .�Ƿ�Ԥ�� = False
            End With
            objBalanceItems_Out.AddItem objItem
            objBalanceItems_Out.���ݺ� = objItem.���ݺ�
            objBalanceItems_Out.������ = RoundEx(objBalanceItems_Out.������ + objItem.������, 6)
            
            If objItem.�����ID <> 0 Then
                objBalanceItems_Out.���� = IIf(objItem.���ѿ�, gEM_���ѿ�, gEM_һ��ͨ)
            Else
                objBalanceItems_Out.���� = gEM_��ͨ����
            End If
        End If
    Next
    objBalanceItems_Out.���� = dbl����
    objBalanceItems_Out.δ�˽�� = objBalanceItems_Out.������
    objBalanceItems_Out.ԭʼ��� = objBalanceItems_Out.������ '�ݶ�Ϊδ�˲���
    objBalanceItems_Out.�쳣ID = lng�쳣ID
    zlGetBalanceItemsFromCardFeeColl = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

 
Private Function GetBalanceItemsFromSwapInfo(ByVal str������Ϣ As String, ByVal blnԤ�� As Boolean, ByVal dblMoney As Double, _
    ByRef objBalanceItems_Out As clsBalanceItems, Optional ByVal blnDefault As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ�����Ϣ,��ȡ��ǰ�Ľ�����Ϣ��
    '���:dblMoney-��ǰ���
    '     str������Ϣ-������Ϣ,���˽����쳣��¼.������Ϣ
    '     blnDefault-��str=������Ϣʱ���Ƿ�ȱʡ,true-ȱʡ;false-��ȱʡ
    '����:objBalanceItems_Out-��ǰ������Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-02-17 15:59:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCurItem As clsBalanceItem
    Dim objCard As Card, cllSwapInfor As Collection, cllExpend As Collection
    Dim bln���ѿ� As Long, lng�����ID As Long, str���㷽ʽ As String
    On Error GoTo errHandle
    
    Set objBalanceItems_Out = New clsBalanceItems
    Set objCurItem = New clsBalanceItem
    If str������Ϣ <> "" Then
        ' GetErrSwapInfoByJsonString
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '����:�����쳣������Ϣ��Json������ȡ�쳣��Ϣ
        '���:
        '����:cllSwapInfo_out-���صĽ�����Ϣ:����,�����ID,������ˮ��,����˵��,���׽��,��ά��,֧����ʽ,����ժҪ
        '     cllExpend_out
        '          |-cllExpend:-��������,��������
        '           ��ʽ:array(����,ֵ),"_����"
        '---------------------------------------------------------------------------------------------------------------------------------------------
        If GetErrSwapInfoByJsonString(str������Ϣ, cllSwapInfor, cllExpend) Then
            bln���ѿ� = Val(cllSwapInfor("_�Ƿ����ѿ�")(1)) = 1
            lng�����ID = Val(cllSwapInfor("_�����ID")(1))
            str���㷽ʽ = Trim(cllSwapInfor("_֧����ʽ")(1))
            
            If lng�����ID <> 0 Then
                If mobjOneCardComLib.zlGetCard(lng�����ID, bln���ѿ�, objCard) = False Then Set objCard = Nothing
            ElseIf str���㷽ʽ <> "" Then
               Set objCard = zlGetCardFromBalanceName(str���㷽ʽ)
            End If
            
            If Not objCard Is Nothing Then
                With objCurItem
                    Set .objCard = objCard
                    .�����ID = IIf(objCard.�ӿ���� < 0, 0, objCard.�ӿ����)
                    .���ѿ� = objCard.���ѿ�
                    .���㷽ʽ = IIf(str���㷽ʽ = "", objCard.���㷽ʽ, str���㷽ʽ)
                    .������ = dblMoney
                    .�������� = objCard.��������
                    .���� = Trim(cllSwapInfor("_����")(1))
                    .������ˮ�� = Trim(cllSwapInfor("_������ˮ��")(1))
                    .����˵�� = Trim(cllSwapInfor("_����˵��")(1))
                    .QRCode = Trim(cllSwapInfor("_��ά��")(1))
                    .����ժҪ = Trim(cllSwapInfor("_����ժҪ")(1))
                    If .�����ID > 0 Then
                       .�������� = IIf(.���ѿ�, 5, 3)
                    ElseIf objCard.�������� = 3 Then
                         .�������� = 2
                    Else
                       .�������� = 0  '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                    End If
                End With
                objBalanceItems_Out.AddItem objCurItem
                objBalanceItems_Out.������ = objCurItem.������
                objBalanceItems_Out.���� = objCurItem.��������
                Set objBalanceItems_Out.objTag = cllExpend
                GetBalanceItemsFromSwapInfo = True
                Exit Function
            End If
        End If
    End If

    If Not blnDefault Then Exit Function
    
    If blnԤ�� Then
         Set objCard = GetDepositPayCard()
    Else
         Set objCard = GetCardFeePayCard()
    End If
    If objCard Is Nothing Then Exit Function
    
    With objCurItem
        Set .objCard = objCard
        .�����ID = IIf(objCard.�ӿ���� < 0, 0, objCard.�ӿ����)
        .���ѿ� = objCard.���ѿ�
        .���㷽ʽ = IIf(str���㷽ʽ = "", objCard.���㷽ʽ, str���㷽ʽ)
        .������ = dblMoney
        .�������� = objCard.��������
        .���� = ""
        .������ˮ�� = ""
        .����˵�� = ""
        .QRCode = ""
        .����ժҪ = ""
        If .�����ID > 0 Then
           .�������� = IIf(.���ѿ�, 5, 3)
        ElseIf objCard.�������� = 3 Then
             .�������� = 2
        Else
           .�������� = 0  '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        End If
    End With
    objBalanceItems_Out.AddItem objCurItem
    objBalanceItems_Out.������ = objCurItem.������
    objBalanceItems_Out.���� = objCurItem.��������
    GetBalanceItemsFromSwapInfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlReadCardAndDepositErrData(ByVal int����״̬ As Integer, Optional ByVal lng�쳣ID As Long, Optional ByRef objPatiInfo_out As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���Ѽ�Ԥ����
    '���:int����״̬ -0-����;1-�쳣����;2-�쳣����
    '     lng�쳣ID-�쳣id
    '����:objPatiInfo_out-���صĲ�����Ϣ����(�쳣���ݵĲ�����Ϣ)
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-28 10:20:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim cllDeposit As Collection, rsPatiInfo As ADODB.Recordset
    Dim bln��ͬ As Boolean, objItems As clsBalanceItems
    Dim strErrMsg As String, intSwapStatu As Integer
    Dim lng�����ID As Long, str������Ϣ As String
    Dim str�������� As String
    Dim i As Long
    
    On Error GoTo errHandle
    
    mlng�쳣ID = lng�쳣ID: mint����״̬ = int����״̬
    Call zlClearControlInfo '���������Ϣ
    mblnShowDepositAndSendCard = True
  
    If int����״̬ = 0 Then
        zlReadCardAndDepositErrData = True: Exit Function
    End If
    strSql = "" & _
    "Select ID, ��������, �Ƿ�����, ҵ��id, �Ƿ�����, ����id, ��ҳid, Ԥ������, ҽ�ƿ�����, �����id, ��������,Ԥ�����,���ѽ��, ͬ��״̬, ������Ϣ, �Ǽ�ʱ��, ����Ա���� " & _
    "     From ���˽����쳣��¼ " & _
    "     Where ID =[1] "
    
    'int����:1-ҽ�ƿ�����;2-������Ϣ�Ǽ�;3-������Ժ �Ǽ�;4-ԤԼ�ҺŽ���
    Set rsTemp = zlDatabase.OpenSQLRecordLob(strSql, Me.Caption, lng�쳣ID)
    If rsTemp.EOF Then
        MsgBox "��ȡ�쳣����ʧ�ܣ������򲢷�ԭ���������ջ����ϣ�����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    mbytӦ�ó��� = Val(Nvl(rsTemp!��������))
    mlngCardTypeID = Val(Nvl(rsTemp!�����ID))
    str������Ϣ = Nvl(rsTemp!������Ϣ)
    str�������� = Nvl(rsTemp!��������)
    
    If InitFace = False Then Exit Function
  
    
    If mbytӦ�ó��� = 3 Then
        ',��ʽ����:һ����:����id:��ҳID,��;һ�֣�����id,��
        '��ѯ����:0-������Ϣ;1-������Ϣ����չ;2-��ȡ��ҳ
        If mobjService.ZlCissvr_GetPatiPageInfo(1, Val(Nvl(rsTemp!����ID)) & ":" & Val(Nvl(rsTemp!��ҳID)), rsPatiInfo, False) = False Then Exit Function
        If rsPatiInfo.RecordCount = 0 Then
            MsgBox "δ��ȡ��������Ϣ������!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        Set mobjPati = New clsPatientInfo
        With rsPatiInfo
            mobjPati.����ID = Val(Nvl(!����ID))
            mobjPati.��ҳID = Val(Nvl(!��ҳID))
            mobjPati.סԺ�� = Trim(Nvl(!סԺ��))
            mobjPati.���� = Trim(Nvl(!����))
            mobjPati.�Ա� = Trim(Nvl(!�Ա�))
            mobjPati.���� = Trim(Nvl(!����))
            mobjPati.�ѱ� = Trim(Nvl(!�ѱ�))
            mobjPati.�������� = Val(Nvl(!��������))
            mobjPati.��˱�־ = Val(Nvl(!��˱�־))
            mobjPati.סԺ״̬ = Val(Nvl(!סԺ״̬))
            mobjPati.��Ժ���� = Trim(Nvl(!��Ժʱ��))
            mobjPati.��Ժ���� = Trim(Nvl(!��Ժʱ��))
            mobjPati.סԺҽʦ = Trim(Nvl(!סԺҽʦ))
            mobjPati.ҽ�Ƹ��ʽ = Trim(Nvl(!ҽ�Ƹ��ʽ����))
            mobjPati.ҽ�Ƹ��ʽ���� = Trim(Nvl(!ҽ�Ƹ��ʽ����))
            mobjPati.��ǰ����id = Val(Nvl(!��ǰ����id))
            mobjPati.��ǰ����id = Val(Nvl(!��ǰ����id))
            mobjPati.ҽ���� = Trim(Nvl(!ҽ����))
            mobjPati.���� = Trim(Nvl(!����))
            mobjPati.���� = Trim(Nvl(!��ǰ����))
            mobjPati.�������� = Trim(Nvl(!��������))
            mobjPati.ѧ�� = Trim(Nvl(!ѧ��))
            mobjPati.ְҵ = Trim(Nvl(!ְҵ))
            mobjPati.���� = Trim(Nvl(!����))
            mobjPati.����״�� = Trim(Nvl(!����״��))
            mobjPati.��Ŀ���� = Trim(Nvl(!��Ŀ����))
            mobjPati.���˱�ע = Trim(Nvl(!���˱�ע))
        End With
    Else
        If mobjOneCardComLib.zlGetPatiInforFromPatiID(Val(Nvl(rsTemp!����ID)), mobjPati) = False Then Exit Function
    End If
     
    Set objPatiInfo_out = mobjPati  '���ز�����Ϣ
    If Nvl(rsTemp!Ԥ������) <> "" Then  '����Ԥ��������Ϣ
         If Val(Nvl(rsTemp!ͬ��״̬)) = -1 Then
             'δ��������
            If GetBalanceItemsFromSwapInfo(str������Ϣ, True, Val(Nvl(rsTemp!Ԥ�����, 0)), mobjDepositItems, True) = False Then Exit Function
            
            For i = 1 To mobjDepositItems.Count
                 mobjDepositItems(i).�������� = 1
                 mobjDepositItems(i).���ݺ� = Nvl(rsTemp!Ԥ������)
                 mobjDepositItems(i).Ԥ��ID = 0
                 mobjDepositItems(i).�쳣ID = lng�쳣ID
                 
                 mobjDepositItems(i).����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
                 
                 mobjDepositItems(i).У�Ա�־ = 1
                 mobjDepositItems(i).�Ƿ�Ԥ�� = True
                 mobjDepositItems(i).�Ƿ񱣴� = True
                 mobjDepositItems(i).�Ƿ���� = False
                 mobjDepositItems(i).�Ƿ�����༭ = True
                 mobjDepositItems(i).�Ƿ�����ɾ�� = True
                 mobjDepositItems(i).�Ƿ��������� = True
            Next
            mobjDepositItems.ͬ��״̬ = -1
            mobjDepositItems.δ�˽�� = Val(Nvl(rsTemp!Ԥ�����, 0))
            mobjDepositItems.�쳣ID = lng�쳣ID
            mobjDepositItems.ҵ��ID = Val(Nvl(rsTemp!ҵ��ID))
            mobjDepositItems.�Ƿ񱣴� = True
            mobjDepositItems.���ݺ� = Nvl(rsTemp!Ԥ������)
            txt�������.Text = mobjDepositItems(1).�������
            If mintInsure = 0 Then
                txt�ɿλ.Text = ""
                txt������.Text = ""
                txt�ʺ�.Text = ""
                chk��λ�ɿ�.value = IIf(txt�ɿλ.Text <> "", 1, 0)
            Else
                chk��λ�ɿ�.value = 0
            End If
            txtԤ����.Text = Format(mobjDepositItems.������, "0.00")
           
         Else
            If ReadDepositBalanceDataFromDepositNo(rsTemp!Ԥ������, lng�쳣ID, Val(Nvl(rsTemp!ҵ��ID)), Val(Nvl(rsTemp!ͬ��״̬)), mobjDepositItems, Val(Nvl(rsTemp!�Ƿ�����)) = 1, str������Ϣ) = False Then Exit Function
            If Val(Nvl(rsTemp!ͬ��״̬)) >= 2 Then
               '�ӿڵ��óɹ��ģ�����Ҫ����
               mblnDepositLocked = True: SetDepositEditEnabled (1) '����������Ϣ
            End If
         End If
            
    Else
        tbDeposit.Visible = False
        mblnSendCardLocked = True: SetDepositEditEnabled (2) '����������Ϣ
    End If
    
    tbSendCard.Visible = False
    lbl������.Visible = False
    fra�ſ�.Visible = True
    
    If mlngCardTypeID = 0 Or Nvl(rsTemp!ҽ�ƿ�����) = "" Then
       mblnSendCardLocked = True: SetCardEditEnabled (2) '��ֹ������Ϣ
    Else
        fra�ſ�.Caption = "��" & mCurSendCard.objSendCard.���� & "������"
        If Val(Nvl(rsTemp!ͬ��״̬)) = -1 Then
            If mobjDepositItems Is Nothing Then
                     'δ��������:' 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ�
                    If GetBalanceItemsFromSwapInfo(str������Ϣ, True, Val(Nvl(rsTemp!���ѽ��, 0)), mobjCardFeeItems, True) = False Then Exit Function
                    For i = 1 To mobjCardFeeItems.Count
                         mobjCardFeeItems(i).�������� = 5
                         mobjCardFeeItems(i).���ݺ� = Nvl(rsTemp!ҽ�ƿ�����)
                         mobjCardFeeItems(i).Ԥ��ID = 0
                         mobjCardFeeItems(i).�쳣ID = lng�쳣ID
                         mobjCardFeeItems(i).����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
                         
                         mobjCardFeeItems(i).У�Ա�־ = 1
                         mobjCardFeeItems(i).�Ƿ�Ԥ�� = True
                         mobjCardFeeItems(i).�Ƿ񱣴� = True
                         mobjCardFeeItems(i).�Ƿ���� = False
                         mobjCardFeeItems(i).�Ƿ�����༭ = True
                         mobjCardFeeItems(i).�Ƿ�����ɾ�� = True
                         mobjCardFeeItems(i).�Ƿ��������� = True
                    Next
                
            Else
                    Set mobjCardFeeItems = mobjDepositItems.Clone
                    mobjCardFeeItems.������ = 0
                    For i = 1 To mobjCardFeeItems.Count
                         mobjCardFeeItems(i).�������� = 5
                         mobjCardFeeItems(i).������ = Val(Nvl(rsTemp!���ѽ��, 0))
                         mobjCardFeeItems(i).���ݺ� = Nvl(rsTemp!ҽ�ƿ�����)
                         mobjCardFeeItems(i).Ԥ��ID = 0
                         mobjCardFeeItems(i).�Ƿ��������� = True
                         mobjCardFeeItems.������ = mobjCardFeeItems.������ + Val(Nvl(rsTemp!���ѽ��, 0))
                    Next
            End If
            mobjCardFeeItems.ͬ��״̬ = -1
            mobjCardFeeItems.δ�˽�� = Val(Nvl(rsTemp!���ѽ��, 0))
            mobjCardFeeItems.�쳣ID = lng�쳣ID
            mobjCardFeeItems.ҵ��ID = Val(Nvl(rsTemp!ҵ��ID))
            mobjCardFeeItems.���ݺ� = Nvl(rsTemp!ҽ�ƿ�����)
            If str�������� <> "" Then
                txt����.Text = str��������
            End If
            txt����.Text = Format(mobjCardFeeItems.������, "0.00")
        Else
            If ReadCardFeeBalanceDataFromNo(rsTemp!ҽ�ƿ�����, lng�쳣ID, Val(Nvl(rsTemp!ҵ��ID)), Val(Nvl(rsTemp!ͬ��״̬)), mobjCardFeeItems, Val(Nvl(rsTemp!�Ƿ�����)) = 1, str������Ϣ) = False Then Exit Function
        End If
    End If
    Call CalcRQCodePayTotal(True)  '����ɨ�븶�ܶ�
    If Val(Nvl(rsTemp!ͬ��״̬)) >= 4 Then
        Call SetCardEditEnabled(2): Call SetDepositEditEnabled(2)   '�������㷽ʽ
        
        mbtQRCodePay.Visible = False
        If Not mobjShowTotalMoneyControl Is Nothing Then mobjShowTotalMoneyControl.Visible = False
        
         If Not mobjDepositItems Is Nothing Then
            If mobjDepositItems.Count <> 0 Then
                 Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
            End If
         End If
         
         If Not mobjCardFeeItems Is Nothing Then
            If mobjCardFeeItems.Count <> 0 Then
                Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
            End If
         End If
         Call RefreshFactNo
         
        zlReadCardAndDepositErrData = True
        Exit Function
    ElseIf Val(Nvl(rsTemp!ͬ��״̬)) >= 2 Then
         mblnSendCardLocked = True: mblnDepositLocked = True
         Call SetCardEditEnabled(1): Call SetDepositEditEnabled(1)   '�������㷽ʽ
         mbtQRCodePay.Visible = False
        If Not mobjShowTotalMoneyControl Is Nothing Then mobjShowTotalMoneyControl.Visible = False
        
         If Not mobjCardFeeItems Is Nothing Then
            If mobjCardFeeItems.Count <> 0 Then
                Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
            End If
         End If
        
         If Not mobjDepositItems Is Nothing Then
            If mobjDepositItems.Count <> 0 Then
                Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
            End If
         End If
         Call RefreshFactNo
         zlReadCardAndDepositErrData = True
         Exit Function
    ElseIf Val(Nvl(rsTemp!ͬ��״̬)) = -1 Then
        'δ��������
         If Not mobjCardFeeItems Is Nothing Then
            If mobjCardFeeItems.Count <> 0 Then
                Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
            End If
         End If
        
         If Not mobjDepositItems Is Nothing Then
            mblnNotClick = True
            If mobjDepositItems.Count <> 0 Then
                Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
            End If
            mblnNotClick = False
         End If
         Call RefreshFactNo
         zlReadCardAndDepositErrData = True
         Exit Function
    End If
    
    bln��ͬ = False
    If Nvl(rsTemp!Ԥ������) <> "" And Nvl(rsTemp!ҽ�ƿ�����) <> "" Then
        'ҽ�ƿ���Ԥ��ͬʱ��
        bln��ͬ = CheckDepsoitAndCardFeePayIsSame(mobjDepositItems, mobjCardFeeItems)
    End If
    
    If bln��ͬ Then
        '--��ͬ
        If mobjDepositItems.���� = gEM_һ��ͨ Then
            Call RefreshFactNo
            Set mobjThirdSwap.objPayCards = mobjCardFeePayCards
            Set objItems = mobjCardFeeItems.Clone
            objItems.������ = mobjDepositItems.������
            objItems(1).������ = RoundEx(objItems(1).������ + mobjDepositItems.������, 6)
            
            If mobjThirdSwap.zlThird_IsSwapIsSucces(objItems, intSwapStatu, strErrMsg, mobjDepositItems(1).Ԥ��ID) = False Then
                '����ʧ��
                'intSwapStatu_Out-�ӿڷ���Falseʱ���˲�����Ч:����״̬: 0-���׵���ʧ��;1-�������ڴ�����
                If intSwapStatu = 1 Then
                    mblnSendCardLocked = True: mblnDepositLocked = True
                    Call SetCardEditEnabled(1): Call SetDepositEditEnabled(1)   '�������㷽ʽ
                    Call SetLoaclePayModefromCard(objItems(1).objCard, False, True)
                    Call SetLoaclePayModefromCard(objItems(1).objCard, True, True)
                    
                End If
            Else
                '���׳ɹ�
                mblnSendCardLocked = True: mblnDepositLocked = True
                Call SetCardEditEnabled(1): Call SetDepositEditEnabled(1)   '�������㷽ʽ
                Call SetLoaclePayModefromCard(objItems(1).objCard, False, True)
                Call SetLoaclePayModefromCard(objItems(1).objCard, True, True)
            End If
        End If
    Else
        '1.Ԥ��
        If Not mobjDepositItems Is Nothing Then
            Call RefreshFactNo
            If mobjDepositItems.Count <> 0 Then
                If mobjDepositItems.���� = gEM_һ��ͨ Then
                    Set mobjThirdSwap.objPayCards = mobjDepositPayCards
                           
                    If mobjThirdSwap.zlThird_IsSwapIsSucces(mobjDepositItems, intSwapStatu, strErrMsg) = False Then
                        '����ʧ��
                        'intSwapStatu_Out-�ӿڷ���Falseʱ���˲�����Ч:����״̬: 0-���׵���ʧ��;1-�������ڴ�����
                        If intSwapStatu = 1 Then
                           mblnDepositLocked = True
                           Call SetDepositEditEnabled(1)   '�������㷽ʽ
                           Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                        End If
                    Else
                        '���׳ɹ�
                        mblnDepositLocked = True: Call SetDepositEditEnabled(1)   '�������㷽ʽ
                        Call SetLoaclePayModefromCard(mobjDepositItems(1).objCard, True, True)
                    End If
                End If
            End If
        End If
        '2.����
        If Not mobjCardFeeItems Is Nothing Then
            If mobjCardFeeItems.Count <> 0 Then
                If mobjCardFeeItems.���� = gEM_һ��ͨ And mobjCardFeeItems.ͬ��״̬ <> -1 Then
                    Set mobjThirdSwap.objPayCards = mobjCardFeePayCards
                    If mobjThirdSwap.zlThird_IsSwapIsSucces(mobjCardFeeItems, intSwapStatu, strErrMsg) = False Then
                        '����ʧ��
                        'intSwapStatu_Out-�ӿڷ���Falseʱ���˲�����Ч:����״̬: 0-���׵���ʧ��;1-�������ڴ�����
                        If intSwapStatu = 1 Then
                            mblnSendCardLocked = True: Call SetCardEditEnabled(1)
                            Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
                        End If
                    Else
                        '���׳ɹ�
                        mblnSendCardLocked = True:  Call SetCardEditEnabled(1)
                        Call SetLoaclePayModefromCard(mobjCardFeeItems(1).objCard, False, True)
                    End If
                End If
            End If
        End If
    End If
    
    
    Call Form_Resize
    zlReadCardAndDepositErrData = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetCardFromCardType(ByVal lng�����ID As Long, bln���ѿ� As Boolean, ByVal str���㷽ʽ As String) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݿ����ID��ȡ������
    '���:lng�����ID-�����ID
    '     bln���ѿ�-�Ƿ����ѿ�
    '     str���㷽ʽ-���㷽ʽ
    '����:
    '����:�ɹ�������
    '����:���˺�
    '����:2018-04-02 14:29:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As New Card
    On Error GoTo errHandle
    If lng�����ID <> 0 Then
        'zlGetCard(ByVal lngCardTypeID As Long, ByVal bln���ѿ� As Boolean,  ByRef objCard As Card) As Boolean
        If mobjOneCardComLib.zlGetCard(lng�����ID, bln���ѿ�, objCard) = False Then
            Set objCard = zlGetCardFromBalanceName(str���㷽ʽ)
        End If
    Else
        Set objCard = zlGetCardFromBalanceName(str���㷽ʽ)
    End If
    Set GetCardFromCardType = objCard: Exit Function

    GetCardFromCardType = True
    Exit Function
errHandle:
    Set objCard = zlGetCardFromBalanceName(str���㷽ʽ)
    Set GetCardFromCardType = objCard: Exit Function
End Function


Private Function GetPatiInfoFromXML(ByVal strPatiXML As String, ByRef int��Ϣ����ģʽ_out As Integer, ByRef cllDrugInfos_Out As Collection, _
    ByRef cllImmuneInfos_Out As Collection, ByRef cllPatiExtInfo_out As Collection, ByRef cllWrangeInfo_out As Collection, _
    ByRef cllOtherPersons_Out As Collection, ByRef cllCertInfos_out As Collection, ByRef dictCardInfo_out As Dictionary) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '����:cllDrugInfos_Out-����ҩ����Ϣ:array(����ҩ������,������Ӧ)
    '     cllImmuneInfos_Out-������Ϣ:array(����ʱ��,��������)
    '     cllPatiExtInfo_out-�ӱ���Ϣ,array(��Ϣ��,��Ϣֵ ),"_" & ��Ϣ��
    '     cllWrangeInfo_out-ҽѧ��ʾ��Ϣarray(��ʾ����,��Ϣֵ),"_��ʾ����",��ʾ���ƣ�ҽѧ��ʾ,������ʾ
    '     cllCertInfos_out-֤����Ϣֵ:array(��Ϣ��,��Ϣֵ )
    '     dictCardInfo_out-ҽ�ƿ�����
    '     cllOtherPersons_Out-������ϵ����Ϣ��
    '       |-cllOtherPerson:��ϵ�ˣ�����,��ϵ,�绰,���֤��) array(����,ֵ),"_����"
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2011-09-08 21:52:04
    'Ŀǰδ�ã����Ժ���չ����ɾ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim i As Long, j As Long, lngCount As Long, lngChildCount As Long
    Dim str����ҩ�� As String, str������Ӧ As String
    Dim str�������� As String, str�������� As String
    Dim strABOѪ�� As String
    Dim str��Ϣ�� As String, str��Ϣֵ As String
    Dim xmlChildNodes As IXMLDOMNodeList, xmlChildNode As IXMLDOMNode
    Dim str���� As String, str��ϵ As String, str�绰 As String, str���֤�� As String, str��ַ As String
    Dim objPati As New clsPatientInfo
    Dim cllTemp As Collection
    On Error GoTo errHandle
        
    Set cllDrugInfos_Out = New Collection
    Set cllImmuneInfos_Out = New Collection
    Set cllPatiExtInfo_out = New Collection
    Set cllWrangeInfo_out = New Collection
    Set cllOtherPersons_Out = New Collection
    Set cllCertInfos_out = New Collection
    Set dictCardInfo_out = New Dictionary
    If strPatiXML = "" Then Exit Function
    
    '    ��Ϣ����ģʽ Integer 1 '0-ǿ�Ƹ��£�1-�������˲����£�2-����������Ϣ��ȱ
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strPatiXML, False) = False Then Exit Function
    Call zlXML_GetNodeValue("��Ϣ����ģʽ", , strValue): int��Ϣ����ģʽ_out = Val(strValue)
    '    ��ʶ    ��������    ����    ����    ˵��
    '    ����    Varchar2    20
    Call zlXML_GetNodeValue("����", , strValue):  objPati.���� = strValue
    '    ����    Varchar2    64
    Call zlXML_GetNodeValue("����", , strValue):  objPati.���� = strValue
    '    �Ա�    Varchar2    4
    Call zlXML_GetNodeValue("�Ա�", , strValue):  objPati.�Ա� = strValue
    '    ����    Varchar2    10
    Call zlXML_GetNodeValue("����", , strValue):  objPati.���� = strValue
    '    ��������    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    Call zlXML_GetNodeValue("��������", , strValue):  objPati.�������� = strValue
    '    �����ص�    Varchar2    50
    Call zlXML_GetNodeValue("�����ص�", , strValue):  objPati.������ַ = strValue
    '    ���֤��    VARCHAR2    18
    Call zlXML_GetNodeValue("���֤��", , strValue):  objPati.���֤�� = strValue
    '    ����֤��    Varchar2    20
    Call zlXML_GetNodeValue("����֤��", , strValue):  objPati.����֤�� = strValue
    '    ְҵ    Varchar2    80
    Call zlXML_GetNodeValue("ְҵ", , strValue):  objPati.ְҵ = strValue
    '    ����    Varchar2    20
    Call zlXML_GetNodeValue("����", , strValue):  objPati.���� = strValue
    '    ����    Varchar2    30
    Call zlXML_GetNodeValue("����", , strValue):  objPati.���� = strValue
    '    ѧ��    Varchar2    10
    Call zlXML_GetNodeValue("ѧ��", , strValue):  objPati.ѧ�� = strValue
    '    ����״��    Varchar2    4
    Call zlXML_GetNodeValue("����״��", , strValue):  objPati.����״�� = strValue
    '    ����    Varchar2    30
    Call zlXML_GetNodeValue("����", , strValue):  objPati.���� = strValue
    '    ��ͥ��ַ    Varchar2    50
    Call zlXML_GetNodeValue("��ͥ��ַ", , strValue):  objPati.��ͥ��ַ = strValue
    '    ���ڵ�ַ    Varchar2    50
    Call zlXML_GetNodeValue("���ڵ�ַ", , strValue):  objPati.���ڵ�ַ = strValue
     '    ��ͥ�绰    Varchar2    20
    Call zlXML_GetNodeValue("��ͥ�绰", , strValue):  objPati.��ͥ�绰 = strValue
    '    ��ͥ��ַ�ʱ�    Varchar2    6
    Call zlXML_GetNodeValue("��ͥ��ַ�ʱ�", , strValue):  objPati.��ͥ�ʱ� = strValue
    '    �໤��  Varchar2    64
    Call zlXML_GetNodeValue("�໤��", , strValue):  objPati.�໤�� = strValue
  
    '    ��ϵ������  Varchar2    64
    Call zlXML_GetNodeValue("��ϵ������", , strValue):  objPati.��ϵ�� = strValue
    '    ��ϵ�˹�ϵ  Varchar2    30
    Call zlXML_GetNodeValue("��ϵ�˹�ϵ", , strValue):  objPati.��ϵ�˹�ϵ = strValue
    '    ��ϵ�˵�ַ  Varchar2    50
    Call zlXML_GetNodeValue("��ϵ�˵�ַ", , strValue):  objPati.��ϵ�˵�ַ = strValue
    '    ��ϵ�˵绰  Varchar2    20
    Call zlXML_GetNodeValue("��ϵ�˵绰", , strValue):  objPati.��ϵ�˵绰 = strValue
     '   ������λ    Varchar2    100
    Call zlXML_GetNodeValue("������λ", , strValue):  objPati.������λ = strValue
    '    ��λ�绰    Varchar2    20
    Call zlXML_GetNodeValue("��λ�绰", , strValue):  objPati.������λ�绰 = strValue
   '�ֻ���   Varchar2    20
    Call zlXML_GetNodeValue("�ֻ���", , strValue):  objPati.�ֻ��� = strValue
    '    ��λ�ʱ�    Varchar2    6
    Call zlXML_GetNodeValue("��λ�ʱ�", , strValue):  objPati.������λ�ʱ� = strValue
    '    ��λ������  Varchar2    50
    Call zlXML_GetNodeValue("��λ������", , strValue):  objPati.������λ�������ʻ� = strValue
    '    ��λ�ʺ�    Varchar2    20
    Call zlXML_GetNodeValue("��λ�ʺ�", , strValue):  objPati.������λ�������ʻ� = strValue
    
    Call zlXML_GetRows("ҩ������", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("ҩ������", i, str����ҩ��)
        Call zlXML_GetNodeValue("ҩ�ﷴӦ", i, str������Ӧ)
        cllDrugInfos_Out.Add Array(str����ҩ��, str������Ӧ)
    Next
    
    lngCount = 0
    '���߼�¼
    Call zlXML_GetRows("��������", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("��������", i, str��������)
        Call zlXML_GetNodeValue("����ʱ��", i, str��������)
        cllImmuneInfos_Out.Add Array(str��������, str��������)
    Next
    
    lngCount = 0
    'ABOѪ��
    Call zlXML_GetNodeValue("ABOѪ��", , strABOѪ��): cllPatiExtInfo_out.Add Array("ABOѪ��", strValue), "_ABOѪ��"
    'RH
    Call zlXML_GetNodeValue("RH", , strValue): cllPatiExtInfo_out.Add Array("RH", strValue), "_RH"
    'ҽѧ��ʾ
    strValue = ""
    Set xmlChildNodes = zlXML_GetChildNodes("�ٴ�������Ϣ")
    
    If Not xmlChildNodes Is Nothing Then
        If xmlChildNodes.length > 0 Then
            For i = 0 To xmlChildNodes.length - 1
                Set xmlChildNode = xmlChildNodes(i)
                If xmlChildNode.Text = "1" Then
                    strValue = strValue & ";" & Replace(xmlChildNode.nodeName, "��־", "")
                End If
            Next
        End If
    End If
    If strValue <> "" Then strValue = Mid(strValue, 2)
    cllWrangeInfo_out.Add Array("ҽѧ��ʾ", strValue), "_ҽѧ��ʾ"
    '����ҽѧ��ʾ
    Call zlXML_GetNodeValue("����ҽѧ��ʾ", , strValue): cllWrangeInfo_out.Add Array("������ʾ", strValue), "_������ʾ"
    '��ϵ��Ϣ
    '    ��ϵ�˵�ַ  Varchar2    50
    Call zlXML_GetNodeValue("��ϵ�˵�ַ", , str��ַ): objPati.��ϵ�˵�ַ = str��ַ
  
     '    ��ϵ������  Varchar2    64
    Call zlXML_GetNodeValue("��ϵ������", , str����): objPati.��ϵ�� = str����
    '    ��ϵ�˹�ϵ  Varchar2    30
    Call zlXML_GetNodeValue("��ϵ�˹�ϵ", , str��ϵ): objPati.��ϵ�˹�ϵ = str��ϵ
    '    ��ϵ�˵绰  Varchar2    20
    Call zlXML_GetNodeValue("��ϵ�˵绰", , str�绰): objPati.��ϵ�˵绰 = str�绰
    '    ��ϵ�����֤ Varchar2   20
    Call zlXML_GetNodeValue("��ϵ�����֤��", , str���֤��): objPati.��ϵ�˵绰 = str���֤��
    Call zlXML_GetRows("��ϵ��Ϣ", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("��ϵ��Ϣ", "����", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Set cllTemp = New Collection
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "����", i, j, str����): cllTemp.Add Array("����", str����), "_����"
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "��ϵ", i, j, str��ϵ): cllTemp.Add Array("��ϵ", str��ϵ), "_��ϵ"
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "�绰", i, j, str�绰): cllTemp.Add Array("�绰", str�绰), "_�绰"
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "���֤��", i, j, str���֤��): cllTemp.Add Array("���֤��", str���֤��), "_���֤��"
                cllOtherPersons_Out.Add cllTemp
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0

    '������Ϣ
    '�����������
    Call zlXML_GetNodeValue("�����������", , strValue): cllPatiExtInfo_out.Add Array("�����������", strValue), "_�����������"
    '��ũ��֤��
    Call zlXML_GetNodeValue("��ũ��֤��", , strValue): cllPatiExtInfo_out.Add Array("��ũ��֤��", strValue), "_��ũ��֤��"

    '����֤��
    Call zlXML_GetRows("����֤��", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("����֤��", "��Ϣ��", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("����֤��", "��Ϣ��", i, j, str��Ϣ��)
                Call zlXML_GetChildNodeValue("����֤��", "��Ϣֵ", i, j, str��Ϣֵ)
                cllCertInfos_out.Add Array(str��Ϣ��, str��Ϣֵ)
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    '������Ϣ
    Call zlXML_GetRows("������Ϣ", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("������Ϣ", "��Ϣ��", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("������Ϣ", "��Ϣ��", i, j, str��Ϣ��)
                Call zlXML_GetChildNodeValue("������Ϣ", "��Ϣֵ", i, j, str��Ϣֵ)
                cllPatiExtInfo_out.Add Array(str��Ϣ��, str��Ϣֵ), "_" & str��Ϣֵ
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    'ҽ�ƿ�����
    Call zlXML_GetRows("ҽ�ƿ�����", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("ҽ�ƿ�����", "��Ϣ��", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("ҽ�ƿ�����", "��Ϣ��", i, j, str��Ϣ��)
                Call zlXML_GetChildNodeValue("ҽ�ƿ�����", "��Ϣֵ", i, j, str��Ϣֵ)
                If dictCardInfo_out.Exists(str��Ϣ��) Then
                    dictCardInfo_out.Item(str��Ϣ��) = str��Ϣֵ
                Else
                    dictCardInfo_out.Add str��Ϣ��, str��Ϣֵ
                End If
            Next
        End If
    Next
    GetPatiInfoFromXML = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

