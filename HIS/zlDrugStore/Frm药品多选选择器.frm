VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmҩƷ��ѡѡ���� 
   Caption         =   "ҩƷѡ����"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8055
   Icon            =   "FrmҩƷ��ѡѡ����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   8055
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MsfҩƷ��� 
      Height          =   2235
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   3942
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf���� 
      Height          =   1305
      Left            =   30
      TabIndex        =   1
      Top             =   2310
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   2302
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmҩƷ��ѡѡ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--�������--
Private IntEditState As Integer                 '�༭״̬(1-���;2-����)
Private LngԴ�ⷿID As Long                     'Դ�ⷿID
Private LngĿ�ⷿID As Long                     'Ŀ�ⷿID
Private Lngʹ�ò���ID As Long                   'ʹ�ò���ID
Private lng��Ӧ��ID As Long                     '��Ӧ��ID
Private strInput As String                      '�����ִ�
Private OutObj As Form                          'ʹ�ñ�����Ĵ��壨�����ṩһ��������¼�������Է��أ�
Private BlnSelect As Boolean                    '�Ƿ�����ѡ��

Private BlnStartUp As Boolean                   '�����ɹ�
Private BlnFirstStart As Boolean                '��һ������
Private RecUnit As New ADODB.Recordset          '��λ
Private StrUnitString As String                 'SQL�ִ�
Private IntStockCheck As Integer                '�����
Private StrFindStyle As String                  'ƥ�䷽ʽ
Private bln�̵㵥 As Boolean                    '�̵㵥�ݱ�־
Private bln������ As Boolean                    '�Ƿ����ӿ����ι�����
Private blnCheck As Boolean                     '�Ƿ����س���ԭ��(�����λ�ʱ��ҩƷ)
Private blnPrice As Boolean                     '�Ƿ�����ʱ�ۻ�����ҩƷ�����
Private mstrCaption As String

'������ʹ�ü�¼��
Private RecData As New ADODB.Recordset          'ҩƷ��;����
Private RecPhysic As New ADODB.Recordset        'ҩƷ��Ƭ
Private RecStock As New ADODB.Recordset         'ҩƷ���

'���ؼ�¼��
Private RecReturn As ADODB.Recordset            '���ؼ�¼��(ҩƷ��Ϣ������,ҩƷĿ¼������,ҩƷ���������)
Private int�ⷿ As Integer                      '1-ҩ��;2-ҩ��;3-�Ƽ���
Private int���� As Integer                      '0-������;1-ҩ�����;2-ҩ������;3-ҩ��ҩ������
Private blnʱ�� As Boolean                      'ʱ��
Private blnStock As Boolean
Private StrCardSortBy As String                 'ҩƷ��Ƭ������
Private StrPhysicSortBy As String               'ҩƷ���������
Private LngCardRow As Long
Private LngPhysicRow As Long
Private LngLastSelectҩƷID As Long             '�ϴ�ѡ���ҩƷID�������Ƿ�ˢ�£�
Private mbln��ҩ�ⷿ As Boolean
Private mblnNoStock As Boolean                  '���ز������Ƿ������̵�û�����ô洢�ⷿ��ҩƷ
Private int���÷�ʽ As Integer                  '0-��ⷿ��ҩ;1-�����������ҩ
Private mbln����ͣ��ҩƷ As Boolean

'����get���ÿ��󣬷��صĿ���������ʵ��������ʵ�ʽ�ʵ�ʲ��
Private mdbl�������� As Double
Private mdblʵ������ As Double
Private mdblʵ�ʽ�� As Double
Private mdblʵ�ʲ�� As Double
Private mdbl������� As Double

'--����--
Private Const StrFormat As String = "'999999999990.99999'"

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrNumberFormat As String
Private mstrMoneyFormat As String

Private mintUnit As Integer             '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintNumberDigit As Integer      '����С��λ��
Private mintMoneyDigit As Integer       '���С��λ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

Private Type WinLocate
    Left As Double
    Top As Double
End Type
Private WindowPosition As WinLocate           '����λ��

'�����б�
Private Const mconIntCol���� As Integer = 17
Private Const mconIntColRID As Integer = 0
Private Const mconIntCol�ⷿ As Integer = 1
Private Const mconIntCol���� As Integer = 2
Private Const mconIntCol������� As Integer = 3
Private Const mconIntCol���� As Integer = 4
Private Const mconIntCol�������� As Integer = 5
Private Const mconIntColʧЧ�� As Integer = 6
Private Const mconIntCol���� As Integer = 7
Private Const mconintCol�ɱ��� As Integer = 8
Private Const mconIntCol�ۼ� As Integer = 9
Private Const mconIntCol�������� As Integer = 10
Private Const mconintCol������� As Integer = 11
Private Const mconIntCol����� As Integer = 12
Private Const mconIntCol����� As Integer = 13
Private Const mconIntCol�ϴι�Ӧ��ID As Integer = 14
Private Const mconIntColʵ������ As Integer = 15
Private Const mconIntCol��׼�ĺ� As Integer = 16

'����б�
Private Const mconIntColSpec���� As Integer = 37
Private Const mconIntColSpec���� As Integer = 0
Private Const mconIntColSpecҩ������ As Integer = 1
Private Const mconIntColSpecͨ������ As Integer = 2
Private Const mconIntColSpecҩƷ��Դ As Integer = 3
Private Const mconIntColSpecҩ��ID As Integer = 4
Private Const mconIntColSpec��;����ID As Integer = 5
Private Const mconIntColSpec������λ As Integer = 6
Private Const mconIntColSpecҩƷ���� As Integer = 7
Private Const mconIntColSpec��Ʒ�� As Integer = 8
Private Const mconIntColSpec��� As Integer = 9
Private Const mconIntColSpec���� As Integer = 10
Private Const mconIntColSpecҩ��ID As Integer = 11
Private Const mconIntColSpecҩƷID As Integer = 12
Private Const mconIntColSpec�ϴβɹ��� As Integer = 13
Private Const mconIntColSpec�ۼ� As Integer = 14
Private Const mconIntColSpec�ۼ۵�λ As Integer = 15
Private Const mconIntColSpec�ۼ۰�װ As Integer = 16
Private Const mconIntColSpec���ﵥλ As Integer = 17
Private Const mconIntColSpec�����װ As Integer = 18
Private Const mconIntColSpecסԺ��λ As Integer = 19
Private Const mconIntColSpecסԺ��װ As Integer = 20
Private Const mconIntColSpecҩ�ⵥλ As Integer = 21
Private Const mconIntColSpecҩ���װ As Integer = 22
Private Const mconIntColSpec�������� As Integer = 23
Private Const mconIntColSpec������� As Integer = 24
Private Const mconIntColSpec����� As Integer = 25
Private Const mconIntColSpec����� As Integer = 26
Private Const mconIntColSpec��Ч�� As Integer = 27
Private Const mconIntColSpecҩ����� As Integer = 28
Private Const mconIntColSpecҩ������ As Integer = 29
Private Const mconIntColSpecʱ�� As Integer = 30
Private Const mconIntColSpecָ�������� As Integer = 31
Private Const mconIntColSpecָ������� As Integer = 32
Private Const mconIntColSpec�ⷿ��λ As Integer = 33
Private Const mconIntColSpec��׼�ĺ� As Integer = 34
Private Const mconIntColSpecʵ������ As Integer = 35
Private Const mconIntColSpec�������� As Integer = 36
Public Property Get In_�༭״̬() As Integer
    In_�༭״̬ = IntEditState
End Property

Public Property Let In_�༭״̬(ByVal vNewValue As Integer)
    IntEditState = vNewValue
End Property

Public Property Get In_Դ�ⷿ() As Long
    In_Դ�ⷿ = LngԴ�ⷿID
End Property

Public Property Let In_Դ�ⷿ(ByVal vNewValue As Long)
    LngԴ�ⷿID = vNewValue
End Property

Public Property Get In_�ִ�() As String
    In_�ִ� = strInput
End Property

Public Property Let In_�ִ�(ByVal vNewValue As String)
    strInput = vNewValue
End Property

Public Property Get In_Ŀ�ⷿ() As Long
    In_Ŀ�ⷿ = LngĿ�ⷿID
End Property

Public Property Let In_Ŀ�ⷿ(ByVal vNewValue As Long)
    LngĿ�ⷿID = vNewValue
End Property

Public Property Get In_����() As Long
    In_���� = Lngʹ�ò���ID
End Property

Public Property Let In_����(ByVal vNewValue As Long)
    Lngʹ�ò���ID = vNewValue
End Property

Public Property Let In_MainFrm(ByVal vNewValue As Form)
    Set OutObj = vNewValue
End Property

Private Sub SetFormat(Optional ByVal IntMain As Integer = 1, Optional ByVal BlnSetHeader As Boolean = False)
    Dim intCol As Integer
    
    '���ø��б�ؼ��ĸ�ʽ
    Select Case IntMain
    Case 1
        With MsfҩƷ���
            
            If BlnSetHeader Then
                .Cols = IIf(int���÷�ʽ = 0, mconIntColSpec���� - 1, mconIntColSpec����)
                '��Ƭ
                .TextMatrix(0, mconIntColSpec����) = "����"
                .TextMatrix(0, mconIntColSpecҩ������) = "ҩ������"
                .TextMatrix(0, mconIntColSpecͨ������) = "ͨ������"
                .TextMatrix(0, mconIntColSpecҩƷ��Դ) = "ҩƷ��Դ"
                .TextMatrix(0, mconIntColSpecҩ��ID) = "ҩ��ID"
                .TextMatrix(0, mconIntColSpec��;����ID) = "��;����ID"
                .TextMatrix(0, mconIntColSpec������λ) = "������λ"
                
                '���
                .TextMatrix(0, mconIntColSpecҩƷ����) = "ҩƷ����"
                .TextMatrix(0, mconIntColSpec��Ʒ��) = "��Ʒ��"
                .TextMatrix(0, mconIntColSpec���) = "���"
                .TextMatrix(0, mconIntColSpec����) = "����"
                .TextMatrix(0, mconIntColSpecҩ��ID) = "ҩ��ID"
                .TextMatrix(0, mconIntColSpecҩƷID) = "ҩƷID"
                .TextMatrix(0, mconIntColSpec�ϴβɹ���) = "�ϴβɹ���"
                .TextMatrix(0, mconIntColSpec�ۼ�) = "�ۼ�"
                .TextMatrix(0, mconIntColSpec�ۼ۵�λ) = "�ۼ۵�λ"
                .TextMatrix(0, mconIntColSpec�ۼ۰�װ) = "�ۼ۰�װ"
                .TextMatrix(0, mconIntColSpec���ﵥλ) = "���ﵥλ"
                .TextMatrix(0, mconIntColSpec�����װ) = "�����װ"
                .TextMatrix(0, mconIntColSpecסԺ��λ) = "סԺ��λ"
                .TextMatrix(0, mconIntColSpecסԺ��װ) = "סԺ��װ"
                .TextMatrix(0, mconIntColSpecҩ�ⵥλ) = "ҩ�ⵥλ"
                .TextMatrix(0, mconIntColSpecҩ���װ) = "ҩ���װ"
                .TextMatrix(0, mconIntColSpec��������) = "��������"
                .TextMatrix(0, mconIntColSpec�������) = "�������"
                .TextMatrix(0, mconIntColSpec�����) = "�����"
                .TextMatrix(0, mconIntColSpec�����) = "�����"
                .TextMatrix(0, mconIntColSpec��Ч��) = "��Ч��"
                .TextMatrix(0, mconIntColSpecҩ�����) = "ҩ�����"
                .TextMatrix(0, mconIntColSpecҩ������) = "ҩ������"
                .TextMatrix(0, mconIntColSpecʱ��) = "ʱ��"
                .TextMatrix(0, mconIntColSpecָ��������) = "ָ��������"
                .TextMatrix(0, mconIntColSpecָ�������) = "ָ�������"
                .TextMatrix(0, mconIntColSpec�ⷿ��λ) = "�ⷿ��λ"
                .TextMatrix(0, mconIntColSpec��׼�ĺ�) = "��׼�ĺ�"
                .TextMatrix(0, mconIntColSpecʵ������) = "ʵ������"
                If int���÷�ʽ = 1 Then
                    .TextMatrix(0, mconIntColSpec��������) = "��������"
                End If
            End If
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            .ColAlignment(mconIntColSpec�ϴβɹ���) = 7
            .ColAlignment(mconIntColSpec�ۼ�) = 7
            .ColAlignment(mconIntColSpec�ۼ۰�װ) = 7
            .ColAlignment(mconIntColSpec�����װ) = 7
            .ColAlignment(mconIntColSpecסԺ��װ) = 7
            .ColAlignment(mconIntColSpecҩ���װ) = 7
            .ColAlignment(mconIntColSpec��������) = 7
            .ColAlignment(mconIntColSpec�������) = 7
            .ColAlignment(mconIntColSpec�����) = 7
            .ColAlignment(mconIntColSpec�����) = 7
            .ColAlignment(mconIntColSpec��Ч��) = 7
            .ColAlignment(mconIntColSpecʱ��) = 7
            .ColAlignment(mconIntColSpecָ��������) = 7
            .ColAlignment(mconIntColSpecָ�������) = 7
            .ColAlignment(mconIntColSpecʵ������) = 7
            If int���÷�ʽ = 1 Then
                .ColAlignment(mconIntColSpec��������) = 7
            End If
            
            If BlnStartUp = False Then
                .ColWidth(mconIntColSpec����) = 500
                .ColWidth(mconIntColSpecҩ������) = 0
                .ColWidth(mconIntColSpecͨ������) = 0
                .ColWidth(mconIntColSpecҩ��ID) = 0
                .ColWidth(mconIntColSpec��;����ID) = 0
                .ColWidth(mconIntColSpec������λ) = 0
                '���
                .ColWidth(mconIntColSpecҩƷ����) = 1000
                .ColWidth(mconIntColSpec��Ʒ��) = 1800
                .ColWidth(mconIntColSpec���) = 1000
                .ColWidth(mconIntColSpec����) = 1200
                .ColWidth(mconIntColSpecҩ��ID) = 0
                .ColWidth(mconIntColSpecҩƷID) = 0
                .ColWidth(mconIntColSpec�ۼ�) = 1200
                .ColWidth(mconIntColSpec��������) = 1200
                .ColWidth(mconIntColSpec�������) = 0
                .ColWidth(mconIntColSpec�����) = 0
                .ColWidth(mconIntColSpec�����) = 0
                .ColWidth(mconIntColSpec��Ч��) = 900
                .ColWidth(mconIntColSpecҩ�����) = 900
                .ColWidth(mconIntColSpecҩ������) = 900
                .ColWidth(mconIntColSpecʱ��) = 900
                .ColWidth(mconIntColSpecָ��������) = 0
                .ColWidth(mconIntColSpecָ�������) = 0
                .ColWidth(mconIntColSpec�ⷿ��λ) = 1500
                .ColWidth(mconIntColSpec��׼�ĺ�) = 1000
                .ColWidth(mconIntColSpecʵ������) = 0
                If int���÷�ʽ = 1 Then
                    .ColWidth(mconIntColSpec��������) = 1000
                End If
                .Row = 1
                Call RestoreFlexState(MsfҩƷ���, App.ProductName & Me.Name)
                
                .ColWidth(mconIntColSpec�ۼ۵�λ) = IIf(mintUnit = mconint�ۼ۵�λ, 900, 0)
                .ColWidth(mconIntColSpec�ۼ۰�װ) = IIf(mintUnit = mconint�ۼ۵�λ, 900, 0)
                .ColWidth(mconIntColSpec���ﵥλ) = IIf(mintUnit = mconint���ﵥλ, 900, 0)
                .ColWidth(mconIntColSpec�����װ) = IIf(mintUnit = mconint���ﵥλ, 900, 0)
                .ColWidth(mconIntColSpecסԺ��λ) = IIf(mintUnit = mconintסԺ��λ, 900, 0)
                .ColWidth(mconIntColSpecסԺ��װ) = IIf(mintUnit = mconintסԺ��λ, 900, 0)
                .ColWidth(mconIntColSpecҩ�ⵥλ) = IIf(mintUnit = mconintҩ�ⵥλ, 900, 0)
                .ColWidth(mconIntColSpecҩ���װ) = IIf(mintUnit = mconintҩ�ⵥλ, 900, 0)
                .ColWidth(mconIntColSpec�ϴβɹ���) = IIf(mstrCaption = "ҩƷ�⹺������", 1200, 0)
            End If
        End With
    Case 0
        With Msf����
            If BlnSetHeader Then
                .Cols = mconIntCol����
                .TextMatrix(0, mconIntColRID) = "RID"
                .TextMatrix(0, mconIntCol�ⷿ) = "�ⷿ"
                .TextMatrix(0, mconIntCol����) = "����"
                .TextMatrix(0, mconIntCol�������) = "�������"
                .TextMatrix(0, mconIntCol����) = "����"
                .TextMatrix(0, mconIntCol��������) = "��������"
                .TextMatrix(0, mconIntColʧЧ��) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")
                .TextMatrix(0, mconIntCol����) = "����"
                .TextMatrix(0, mconintCol�ɱ���) = "�ɱ���"
                .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
                .TextMatrix(0, mconIntCol��������) = "��������"
                .TextMatrix(0, mconintCol�������) = "�������"
                .TextMatrix(0, mconIntCol�����) = "�����"
                .TextMatrix(0, mconIntCol�����) = "�����"
                .TextMatrix(0, mconIntCol�ϴι�Ӧ��ID) = "�ϴι�Ӧ��ID"
                .TextMatrix(0, mconIntColʵ������) = "ʵ������"
                .TextMatrix(0, mconIntCol��׼�ĺ�) = "��׼�ĺ�"
            End If
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            .ColWidth(mconIntColRID) = 0
            .ColAlignment(mconIntCol����) = 7
            .ColAlignment(mconintCol�ɱ���) = 7
            .ColAlignment(mconIntCol�ۼ�) = 7
            .ColAlignment(mconIntCol��������) = 7
            .ColAlignment(mconintCol�������) = 7
            .ColAlignment(mconIntCol�����) = 7
            .ColAlignment(mconIntCol�����) = 7
            If BlnStartUp = False Then
                .ColWidth(mconIntColRID) = 0
                .ColWidth(mconIntCol�ⷿ) = 1200
                .ColWidth(mconIntCol����) = 0
                .ColWidth(mconIntCol����) = 1000
                .ColWidth(mconIntCol��������) = 1000
                .ColWidth(mconIntColʧЧ��) = 1000
                .ColWidth(mconIntCol����) = 1200
                .ColWidth(mconintCol�ɱ���) = 1200
                .ColWidth(mconIntCol�ۼ�) = 1200
                .ColWidth(mconIntCol��������) = 1200
                .ColWidth(mconintCol�������) = 1200
                .ColWidth(mconIntCol�����) = 1200
                .ColWidth(mconIntCol�����) = 1200
                .ColWidth(mconIntCol�ϴι�Ӧ��ID) = 0
                .ColWidth(mconIntColʵ������) = 0
                .ColWidth(mconIntCol��׼�ĺ�) = 1000
                .Row = 1
                
                Call RestoreFlexState(Msf����, App.ProductName & Me.Name)
                .ColWidth(mconIntCol�������) = IIf(mstrCaption = "ҩƷ�ƿ����" Or mstrCaption = "ҩƷ�������", 1000, 0)
            End If
        End With
    End Select
End Sub

Private Sub OnCancel()
    Unload Me
    Exit Sub
End Sub

Private Sub OnSelect()
    Dim blnValid As Boolean
    On Error Resume Next
    
    If In_�༭״̬ = 2 Then If CheckData = False Then Exit Sub
    '�������������������Ƿ�һ��
    If In_�༭״̬ = 2 Then
        blnValid = ���������(LngԴ�ⷿID, LngLastSelectҩƷID)
    Else
        blnValid = ���������(LngĿ�ⷿID, LngLastSelectҩƷID)
    End If
    If Not blnValid Then
        MsgBox "���ָ�ҩƷ�ڵ�ǰ�ⷿ�еĿ���¼���ڴ��󣨿����ǻ����������ô������鵱ǰ�ⷿ�Ĳ������ʼ���ҩƷ�ķ������ԣ���", vbInformation, gstrSysName
        Exit Sub
    End If
    '��װ��¼��
    If CombinateRec = False Then Exit Sub
    
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    BlnStartUp = False
    BlnFirstStart = False
    
    'ȡ�ۼ۵�λ
    StrFindStyle = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
    StrUnitString = ""
    IntStockCheck = 0
    LngLastSelectҩƷID = 0
    Msf����.Visible = (In_�༭״̬ = 2)
    
    '��ʼ����¼��
    InitRec
    
    If strInput = "" Then Exit Sub
    strInput = UCase(strInput)
    If OutObj Is Nothing Then
        MsgBox "��ָ�������壡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��λ
    With WindowPosition
        Me.Left = .Left
        Me.Top = .Top
    End With
    
    '��ȡ��ǰ�����Ʋ���
    gstrSQL = "Select Nvl(��鷽ʽ,0) ����� From ҩƷ������ Where �ⷿID=[1]"
    Set RecUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, LngԴ�ⷿID)
    
    With RecUnit
        If Not .EOF Then
            IntStockCheck = !�����
        End If
    End With
    
    '���Դ�ⷿ�Ƿ�Ϊҩ��
    If LngԴ�ⷿID <> 0 Then
        int�ⷿ = 3

        gstrSQL = "select ����ID from ��������˵�� where (�������� like '%ҩ��' Or �������� like '%�Ƽ���') And ����id=[1]"
        Set RecUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, LngԴ�ⷿID)
       
        If RecUnit.EOF Then
            RecUnit.Close
            
            gstrSQL = "select ����ID from ��������˵�� where �������� like '%ҩ��' And ����id=[1]"
            Set RecUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, LngԴ�ⷿID)
            
            If Not RecUnit.EOF Then int�ⷿ = 1
        Else
            int�ⷿ = 2
        End If
    End If
    
    mstrCaption = GetText(GetParentWindow(OutObj.hWnd))
    '������ƿⰴָ����λ��ʾ
    If mstrCaption = "ҩƷ�������" Then
        Call GetDrugDigit(Lngʹ�ò���ID, mstrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        Lngʹ�ò���ID = 0
    ElseIf mstrCaption = "ҩƷ�ƿ����" Then
        Call GetDrugDigit(Lngʹ�ò���ID, mstrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        Lngʹ�ò���ID = 0
    Else
        Call GetDrugDigit(IIf(LngԴ�ⷿID = 0, LngĿ�ⷿID, LngԴ�ⷿID), mstrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    End If
    
    mstrCostFormat = "'999999999990." & String(mintCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintMoneyDigit, "0") & "'"

    Select Case mintUnit
        Case mconint���ﵥλ
            StrUnitString = "/nvl(�����װ,1)"
        Case mconintסԺ��λ
            StrUnitString = "/nvl(סԺ��װ,1)"
        Case mconintҩ�ⵥλ
            StrUnitString = "/nvl(ҩ���װ,1)"
    End Select
    
    BlnStartUp = RefreshData
    On Error Resume Next
    If RecPhysic.RecordCount = 1 Then
        If Not (((int���� = 3 And int�ⷿ <> 3) Or (int���� = 1 And int�ⷿ = 1) Or (int���� = 2 And int�ⷿ = 2)) And blnPrice) Or IntEditState = 1 Then OnSelect
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    BlnFirstStart = True
    
    If Me.Height < 4185 Then Me.Height = 4185
    If Me.Width < 8295 Then Me.Width = 8295
    
    With MsfҩƷ���
        .Width = Me.ScaleWidth
        .Height = IIf(Msf����.Visible = False, Me.ScaleHeight - .Top, Msf����.Top - 45 - .Top)
        If .RowIsVisible(.Row) = False Then .TopRow = .TopRow + IIf(.Row > .TopRow, 1, -1)
    End With
    
    With Msf����
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - MsfҩƷ���.Top - MsfҩƷ���.Height - 45
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Call SaveFlexState(MsfҩƷ���, App.ProductName & Me.Name)
    Call SaveFlexState(Msf����, App.ProductName & Me.Name)
End Sub

Private Sub Msf����_Click()
    Dim StrHeader As String
    Dim intCol As Integer
    'ʵ��������
    On Error Resume Next
    With Msf����
        If .MouseRow <> 0 Then Exit Sub
        If RecStock.EOF Then Exit Sub
        
        StrHeader = .TextMatrix(0, .MouseCol)
        Set .DataSource = Nothing
        If Mid(StrPhysicSortBy, 2) = StrHeader Then
            StrPhysicSortBy = IIf(Mid(StrPhysicSortBy, 1, 1) = "A", "D", "A") & .TextMatrix(0, .MouseCol)
            RecStock.Sort = .TextMatrix(0, .MouseCol) & IIf(Mid(StrPhysicSortBy, 1, 1) = "A", " Desc", " Asc")
        Else
            StrPhysicSortBy = "A" & .TextMatrix(0, .MouseCol)
            RecStock.Sort = .TextMatrix(0, .MouseCol) & " Asc"
        End If
        Set .DataSource = RecStock
        
        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        
        Call SetFormat(0, False)
    End With
End Sub

Private Sub Msf����_DblClick()
    With RecStock
        If .RecordCount <> 0 Then .MoveFirst
        If .EOF Then Exit Sub
        If .RecordCount = 0 Then Exit Sub
    End With
    
    If BlnSelect Then OnSelect
End Sub

Private Sub Msf����_EnterCell()
    Dim intCol As Integer, LngSelectRow As Long
    Dim RecGetPrice As New ADODB.Recordset
    On Error Resume Next
    
    With Msf����
        .Redraw = False
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If LngPhysicRow <> 0 Then
            .Row = IIf(LngPhysicRow > .Rows - 1, 0, LngPhysicRow) '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        LngPhysicRow = LngSelectRow
        .Row = LngPhysicRow     '���õ�ǰѡ����
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        
        .Redraw = True
    End With
End Sub

Private Sub Msf����_GotFocus()
    With Msf����
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Msf����_DblClick
End Sub

Private Sub Msf����_LostFocus()
    With Msf����
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub MsfҩƷ���_Click()
    Dim StrHeader As String
    Dim intCol As Integer
    
    'ʵ��������
    On Error Resume Next
    With MsfҩƷ���
        If .MouseRow <> 0 Then Exit Sub
        If RecPhysic.EOF Then Exit Sub
        
        StrHeader = .TextMatrix(0, .MouseCol)
        Set .DataSource = Nothing
        If Mid(StrCardSortBy, 2) = StrHeader Then
            StrCardSortBy = IIf(Mid(StrCardSortBy, 1, 1) = "A", "D", "A") & .TextMatrix(0, .MouseCol)
            RecPhysic.Sort = .TextMatrix(0, .MouseCol) & IIf(Mid(StrCardSortBy, 1, 1) = "A", " Desc", " Asc")
        Else
            StrCardSortBy = "A" & .TextMatrix(0, .MouseCol)
            RecPhysic.Sort = .TextMatrix(0, .MouseCol) & " Asc"
        End If
        Set .DataSource = RecPhysic
        
        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        
        Call SetFormat(1, False)
    End With
End Sub

Private Sub MsfҩƷ���_DblClick()
    If RecPhysic.EOF Then Exit Sub
    If RecPhysic.RecordCount = 0 Then Exit Sub
    
    If BlnSelect Then
        OnSelect
    Else
        MsgBox "��ҩƷû�п�棬���ܼ���������", vbInformation, gstrSysName
    End If
End Sub

Private Sub MsfҩƷ���_EnterCell()
    Dim Lng�շ�ϸĿID As Long, intCol As Integer, LngSelectRow As Long
    Dim StrTmp As String, RecGetPrice As New ADODB.Recordset
    Dim strSqlЧ�� As String
    Dim str�ۼ� As String
    
    With MsfҩƷ���
        .Redraw = False
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If LngCardRow <> 0 Then
            .Row = IIf(LngCardRow > .Rows - 1, 0, LngCardRow) '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        LngCardRow = LngSelectRow
        .Row = LngCardRow       '���õ�ǰѡ����
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        
        .Redraw = True
        
        '����ù��ҩƷ�ļ۸�ִ��ʱ�仹δִ��,�򴥷�
        Lng�շ�ϸĿID = Val(.TextMatrix(.Row, mconIntColSpecҩƷID))
        If Lng�շ�ϸĿID = 0 Then
            If Msf����.Visible Then
                Msf����.Clear
                Msf����.Rows = 2
                Call SetFormat(0, True)
                Msf����_EnterCell
            Else
                Call SetFormat(0, True)
            End If
            
            Exit Sub
        End If
        
        If LngLastSelectҩƷID = Lng�շ�ϸĿID Then Exit Sub
        LngLastSelectҩƷID = Lng�շ�ϸĿID
        
        '����ѵ�ִ�����ڶ��۸�δִ�У�ִ�м������
        gstrSQL = " Select ID From �շѼ�Ŀ Where �շ�ϸĿID=[1] And �䶯ԭ��=0"
        Set RecGetPrice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Lng�շ�ϸĿID)
                     
        With RecGetPrice
            If Not .EOF Then
                If Not IsNull(!Id) Then
                    Lng�շ�ϸĿID = !Id
                    gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & Lng�շ�ϸĿID & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-�����۸������¼")
                End If
            End If
        End With
    End With
    
    If In_�༭״̬ = 2 Then
        Msf����.Visible = False
        '������ҩƷ��������е�ҩƷ���ο����Ϣ
        blnʱ�� = (MsfҩƷ���.TextMatrix(MsfҩƷ���.Row, mconIntColSpecʱ��) = "��")
        int���� = 0
        str�ۼ� = MsfҩƷ���.TextMatrix(MsfҩƷ���.Row, mconIntColSpec�ۼ�)
        If MsfҩƷ���.TextMatrix(MsfҩƷ���.Row, mconIntColSpecҩ�����) = "��" Or MsfҩƷ���.TextMatrix(MsfҩƷ���.Row, mconIntColSpecҩ������) = "��" Then
            If MsfҩƷ���.TextMatrix(MsfҩƷ���.Row, mconIntColSpecҩ�����) = "��" And MsfҩƷ���.TextMatrix(MsfҩƷ���.Row, mconIntColSpecҩ������) = "��" Then
                int���� = 3
            ElseIf MsfҩƷ���.TextMatrix(MsfҩƷ���.Row, mconIntColSpecҩ�����) = "��" Then
                int���� = 1
            Else
                int���� = 2
            End If
        End If
        If Not ((int���� = 3 And int�ⷿ <> 3) Or (int���� = 1 And int�ⷿ = 1) Or (int���� = 2 And int�ⷿ = 2)) Then         '�����ҩƷ������
            Msf����.Visible = False
            Form_Resize
        Else
            If Msf����.Visible = False Then Msf����.Visible = True
        End If
        Form_Resize
        
        With RecStock
            If .State = 1 Then .Close
            gstrSQL = ""
            If bln������ Then
                strSqlЧ�� = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")
                gstrSQL = "Select 1 RID,����,0 ����,'' As �������,'��������ҩƷ' ����,NULL ��������,sysdate " & strSqlЧ�� & ",'' ����,'' �ɱ���,''�ۼ�," & _
                          "'' ��������,'' �������,'' �����,'' �����,0 �ϴι�Ӧ��ID,'' ʵ������,'' ��׼�ĺ� " & _
                          " From ���ű�" & _
                          " Where ID=[1] " & _
                          " Union "
            End If
            
            strSqlЧ�� = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "K.Ч��-1 As ��Ч����", "K.Ч�� As ʧЧ��")
            gstrSQL = gstrSQL & " Select 2 RID,P.���� �ⷿ,K.����,TO_CHAR(S.�������, 'YYYY-MM-DD') As �������,K.�ϴ����� ����,To_Char(K.�ϴ���������,'YYYY-MM-DD') AS ��������," & strSqlЧ�� & ",K.�ϴβ��� AS ����,"
            If blnStock Then
                Select Case mintUnit
                Case mconint�ۼ۵�λ
                    StrTmp = " To_Char(K.�ϴβɹ���," & mstrNumberFormat & ") �ɱ���, " & _
                             IIf(blnʱ�� = True, " Decode(Sign(K.ʵ������),1,To_Char(K.ʵ�ʽ��/K.ʵ������," & mstrNumberFormat & "),'" & str�ۼ� & "') �ۼ�, ", " '" & str�ۼ� & "' �ۼ�, ") & _
                             " To_Char(K.��������," & mstrNumberFormat & ") ��������," & _
                             " To_Char(K.ʵ������," & mstrNumberFormat & ") �������,"
                Case mconint���ﵥλ
                    StrTmp = " To_Char(K.�ϴβɹ���*nvl(D.�����װ,1)," & mstrNumberFormat & ") �ɱ���, " & _
                             IIf(blnʱ�� = True, " Decode(Sign(K.ʵ������),1,To_Char(K.ʵ�ʽ��/K.ʵ������*nvl(D.�����װ,1)," & mstrNumberFormat & "),'" & str�ۼ� & "') �ۼ�, ", " '" & str�ۼ� & "' �ۼ�, ") & _
                             " To_Char(K.��������" & StrUnitString & "," & mstrNumberFormat & ") ��������," & _
                             " To_Char(K.ʵ������" & StrUnitString & "," & mstrNumberFormat & ") �������,"
                Case mconintסԺ��λ
                    StrTmp = " To_Char(K.�ϴβɹ���*nvl(D.סԺ��װ,1)," & mstrNumberFormat & ") �ɱ���, " & _
                             IIf(blnʱ�� = True, " Decode(Sign(K.ʵ������),1,To_Char(K.ʵ�ʽ��/K.ʵ������*nvl(D.סԺ��װ,1)," & mstrNumberFormat & "),'" & str�ۼ� & "') �ۼ�, ", " '" & str�ۼ� & "' �ۼ�, ") & _
                             " To_Char(K.��������" & StrUnitString & "," & mstrNumberFormat & ") ��������," & _
                             " To_Char(K.ʵ������" & StrUnitString & "," & mstrNumberFormat & ") �������,"
                Case mconintҩ�ⵥλ
                    StrTmp = " To_Char(K.�ϴβɹ���*nvl(D.ҩ���װ,1)," & mstrNumberFormat & ") �ɱ���, " & _
                             IIf(blnʱ�� = True, " Decode(Sign(K.ʵ������),1,To_Char(K.ʵ�ʽ��/K.ʵ������*nvl(D.ҩ���װ,1)," & mstrNumberFormat & "),'" & str�ۼ� & "') �ۼ�, ", " '" & str�ۼ� & "' �ۼ�, ") & _
                             " To_Char(K.��������" & StrUnitString & "," & mstrNumberFormat & ") ��������," & _
                             " To_Char(K.ʵ������" & StrUnitString & "," & mstrNumberFormat & ") �������,"
                End Select
            Else
                StrTmp = "'' ��������,'' �������,"
            End If
            
            gstrSQL = gstrSQL & StrTmp & IIf(blnStock, " To_Char(K.ʵ�ʽ��," & mstrMoneyFormat & ") �����,", "'' �����,") & _
                     IIf(blnStock, " To_Char(K.ʵ�ʲ��," & mstrMoneyFormat & ") �����", "'' �����") & _
                     " ,NVL(K.�ϴι�Ӧ��id,0) �ϴι�Ӧ��id,To_Char(K.ʵ������," & mstrNumberFormat & ") AS ʵ������,K.��׼�ĺ�  " & _
                     " From ���ű� P,ҩƷ��� D,ҩƷ��� K,ҩƷ�շ���¼ S" & _
                     " Where K.�ⷿID = P.ID And D.ҩƷID = K.ҩƷID And K.�ⷿID=[2] " & _
                     " And K.ҩƷID=[3] And K.����=1 And Decode(Nvl(K.����,0),0,-999,K.����)=S.Id(+) "
            If bln�̵㵥 Then
                gstrSQL = gstrSQL & " And (K.ʵ������<>0 Or K.ʵ�ʽ��<>0 Or K.ʵ�ʲ��<>0)"
            ElseIf glngModul <> 1303 Then   '����ǿ���۵���ģ�飬��������˿������Ϊ0��ҩƷ��¼
                gstrSQL = gstrSQL & " And K.ʵ������<>0 "
            End If

            If gtype_UserSysParms.P150_ҩƷ���������㷨 = 0 Then
                gstrSQL = gstrSQL & " Order by RID,����"
            Else
                gstrSQL = gstrSQL & " Order by RID," & IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��") & ",����"
            End If
        End With
        
        Set RecStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, LngĿ�ⷿID, IIf(LngԴ�ⷿID = 0, LngĿ�ⷿID, LngԴ�ⷿID), LngLastSelectҩƷID)
        
        Dim BlnState As Boolean
        With Msf����
            If Not RecStock.EOF Then
                Set .DataSource = RecStock
                .ColWidth(mconIntColRID) = 0
            Else
                .Clear
                .Rows = 2
            End If
            
            Call SetFormat(0, RecStock.EOF)
            If bln������ And RecStock.RecordCount <> 0 Then .Row = IIf(.Rows > 2, 2, 1)
            BlnState = ((int���� = 3 And int�ⷿ <> 3) Or (int���� = 1 And int�ⷿ = 1) Or (int���� = 2 And int�ⷿ = 2)) And Not RecStock.EOF        '�����ҩƷ������
            .Visible = BlnState
            Msf����_EnterCell
        End With
        Form_Resize
    End If
    
    '���ð�ť״̬
    With RecPhysic
        If .RecordCount <> 0 Then .MoveFirst
        .Find "ҩƷID=" & Val(MsfҩƷ���.TextMatrix(MsfҩƷ���.Row, mconIntColSpecҩƷID))
        If .EOF Then
            MsgBox "�����ڲ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If In_�༭״̬ = 2 And ((int���� = 3 And int�ⷿ <> 3) Or (int���� = 1 And int�ⷿ = 1) Or (int���� = 2 And int�ⷿ = 2)) And blnPrice Then
            BlnSelect = BlnState
        Else
            BlnSelect = True
        End If
    End With
End Sub

Private Sub MsfҩƷ���_GotFocus()
    With MsfҩƷ���
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub MsfҩƷ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then MsfҩƷ���_DblClick
End Sub

Private Sub MsfҩƷ���_LostFocus()
    With MsfҩƷ���
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Function RefreshData() As Boolean
    Dim StrTmp As String, StrGroupBy As String
    Dim str��λת���� As String
    Dim str��ʾ���� As String
    Dim strFindString As String
    
    On Error GoTo ErrHand
    RefreshData = False
    
    str��ʾ���� = IIf(int���÷�ʽ = 1, ",To_Char(S.�������� ," & mstrNumberFormat & ") ��������", "")
    
    '������ҩƷ��;���ࡢ����ָ�����͵Ĺ��ҩƷ
    Select Case mintUnit
        Case mconint�ۼ۵�λ
            str��λת���� = "*1"
        Case mconint���ﵥλ
            str��λת���� = "*D.�����װ"
        Case mconintסԺ��λ
            str��λת���� = "*D.סԺ��װ"
        Case mconintҩ�ⵥλ
            str��λת���� = "*D.ҩ���װ"
    End Select
    
    strFindString = " And (A.���� Like [1] OR B.���� Like [1] OR B.���� LIKE [1])"
    
    If IsNumeric(strInput) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
        If Mid(gtype_UserSysParms.P44_����ƥ��, 1, 1) = "1" Then strFindString = " And (A.���� Like [1] Or B.���� Like [1] And B.����=3)"
    ElseIf zlCommFun.IsCharAlpha(strInput) Then         '01,11.����ȫ����ĸʱֻƥ�����
        If Mid(gtype_UserSysParms.P44_����ƥ��, 2, 1) = "1" Then strFindString = " And B.���� Like [1] "
    ElseIf zlCommFun.IsCharChinese(strInput) Then
        strFindString = " And B.���� Like [1] "
    End If
    
    With RecPhysic
        If .State = 1 Then .Close
        
        '--��Ϊ����������ƻ���룬���˷�ʽ����ָ�����ҩƷ--
        gstrSQL = " Select D.����,D.ҩ������,D.ͨ������,D.ҩƷ��Դ,D.ҩ��ID,D.��;����ID,D.������λ,D.ҩƷ����,D.��Ʒ��,D.���," & IIf(IntEditState = 1, "D.����", "Nvl(D.����,S.����)") & " AS ����," & _
                " D.ҩ��ID,D.ҩƷID,trim(to_char(D.��ʼ�ɱ���" & str��λת���� & "," & mstrCostFormat & ")) As �ϴβɹ���,trim(to_char(P.�ۼ�" & str��λת���� & "," & mstrPriceFormat & ")) As �ۼ�," & _
                " D.�ۼ۵�λ,D.����ϵ��,D.���ﵥλ,D.�����װ,D.סԺ��λ,D.סԺ��װ,D.ҩ�ⵥλ,D.ҩ���װ," & _
                IIf(blnStock, " To_Char(S.�������� " & StrUnitString & " ," & mstrNumberFormat & ") ��������,To_Char(S.������� " & StrUnitString & "," & mstrNumberFormat & ") �������,S.�����,S.�����,", "'' ��������,'' �������,'' �����,'' �����,") & _
                " D.���Ч�� ��Ч��,D.ҩ�����,D.ҩ������,D.ʱ��,D.��ʼ�ɱ���,D.ָ��������,D.ָ�������,E.�ⷿ��λ,D.��׼�ĺ�,To_Char(S.������� ," & mstrNumberFormat & ") ʵ������" & str��ʾ���� & _
                " From"
        gstrSQL = gstrSQL & " (SELECT DISTINCT J.���� ����,C.���� ҩ������,C.���� AS ͨ������,0 AS ҩ��ID,M.����ID AS ��;����ID,M.���㵥λ AS ������λ,C.���� AS ҩƷ����," & _
                " " & IIf(mblnTradeName, "NVL(A.����,C.����)", "C.����") & " ��Ʒ��,C.���,C.����,D.ҩƷ��Դ, D.��׼�ĺ�, D.ҩ��ID,D.ҩƷID, C.���㵥λ AS �ۼ۵�λ," & _
                " To_Char(D.����ϵ��," & StrFormat & " ) ����ϵ��,nvl(To_Char(D.���Ч��,'9999990'),0) ���Ч��," & _
                " DECODE(D.ҩ�����,1,'��','��') ҩ�����,DECODE(D.ҩ������,1,'��','��') ҩ������,DECODE(C.�Ƿ���,1,'��','��') ʱ��," & _
                " D.���ﵥλ,To_Char(D.�����װ," & StrFormat & " ) �����װ,D.סԺ��λ," & _
                " To_Char(D.סԺ��װ," & StrFormat & " ) סԺ��װ,D.ҩ�ⵥλ,To_Char(D.ҩ���װ," & StrFormat & " ) ҩ���װ," & _
                " To_Char(D.ָ��������," & mstrCostFormat & ") ָ��������,nvl(D.�ɱ���,0) ��ʼ�ɱ���,To_Char(D.ָ�������," & StrFormat & " ) ָ�������" & _
                " FROM (" & _
                "     Select A.* From �շ���ĿĿ¼ A,�շ���Ŀ���� B" & _
                "     Where A.ID=B.�շ�ϸĿID AND A.��� IN ('5','6','7') And (A.վ�� = '" & gstrNodeNo & "' Or A.վ�� is Null)" & strFindString & _
                " ) C,ҩƷ��� D,�շ���Ŀ���� A,ҩƷ���� J,ҩƷ���� T,������ĿĿ¼ M," & _
                "             (Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID" & IIf(LngԴ�ⷿID <> 0, "=[2] ", IIf(LngĿ�ⷿID <> 0, "=[3] ", " Is Not NULL")) & " Group By ִ�п���ID,�շ�ϸĿID) K," & _
                "             (Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID" & IIf(LngĿ�ⷿID <> 0, "=[3] ", IIf(LngԴ�ⷿID <> 0, "=[2] ", " Is Not NULL")) & " Group By ִ�п���ID,�շ�ϸĿID) I " & _
                " WHERE C.ID=D.ҩƷID AND D.ҩ��ID=T.ҩ��ID AND T.ҩ��ID=M.ID AND M.��� IN ('5','6','7')" & _
                " AND D.ҩƷID=K.�շ�ϸĿID" & IIf(mblnNoStock, "(+)", "") & " " & _
                " And D.ҩƷID=I.�շ�ϸĿID" & IIf(mblnNoStock, "(+)", "") & " " & _
                " AND D.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3 " & _
                " AND T.ҩƷ����=J.����(+)"
                'IIf(Lngʹ�ò���ID <> 0, " And K.��������ID=I.��������ID And K.��������ID=" & Lngʹ�ò���ID, "")
'       ���Ŀ��ⷿ����ȷ����������������ã������Ƽ��ң�������ҩƷ����
'       ���Ŀ��ⷿ����ҩ�����ҩ����������ҩ���Խ��룻
'       ���Ŀ��ⷿ�ǳ�ҩ����ҩ�������г�ҩ���Խ��룻
'       ���Ŀ��ⷿ����ҩ�����ҩ�������в�ҩ���Խ��룻
'
'       ���Ŀ��ⷿ����ȷ����������������ã�����ҩ�⡢�Ƽ��ң������Ʒ������
'       ���Ŀ��ⷿ�Ƿ��������ﲡ�ˣ���������ҩ���Խ��룻
'       ���Ŀ��ⷿ�Ƿ�����סԺ���ˣ���סԺ��ҩ���Խ��룻
        gstrSQL = gstrSQL & "" & _
            " and ([3] is null" & _
                " or exists(select 1 from ��������˵�� where ��������='�Ƽ���' and ����id=[3])" & _
                " or C.���=(select distinct '5' from ��������˵�� where �������� like '��ҩ%' and ����id=[3])" & _
                " or C.���=(select distinct '6' from ��������˵�� where �������� like '��ҩ%' and ����id=[3])" & _
                " or C.���=(select distinct '7' from ��������˵�� where �������� like '��ҩ%' and ����id=[3]) Or [3]=0)" & _
            " and ([3] is null" & _
                " or exists(select 1 from ��������˵�� where �������� like '%ҩ��' and ����id=[3])" & _
                " or exists(select 1 from ��������˵�� where ��������='�Ƽ���' and ����id=[3])" & _
                " or decode(C.�������,1,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[3] and ������� in(1,3))" & _
                " or decode(C.�������,2,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[3] and ������� in(2,3)) Or [3]=0)"
        
        'ֻ����δͣ�õĹ��ҩƷ����Ҫ���ݴ��������������ʱֻ���̵�ʱ�ò����ſ���ΪTrue��
        If mbln����ͣ��ҩƷ = True Then
            gstrSQL = gstrSQL & ") D,"
        Else
            gstrSQL = gstrSQL & " And (C.����ʱ�� Is Null Or To_char(C.����ʱ��,'yyyy-MM-dd')='3000-01-01')) D,"
        End If
        
        '��ȡ����ҩƷ�ĵ�ǰ�ۼ�
        gstrSQL = gstrSQL & " (Select �շ�ϸĿid,To_Char(�ּ�," & mstrPriceFormat & ") �ۼ� From �շѼ�Ŀ Where (Sysdate Between ִ������ And ��ֹ���� or Sysdate>=ִ������ And ��ֹ���� Is Null)) P,"
        '��ȡ����ҩƷ�ĵ�ǰ�ۼ�
        If blnStock Then
            If int���÷�ʽ = 1 Then
                gstrSQL = gstrSQL & " (Select a.ҩƷid,Max(�ϴβ���) AS ����,To_Char(Sum(a.��������),'99999999999990.99999') ��������," & _
                        " To_Char(Sum(a.ʵ������),'99999999999990.99999') �������," & _
                        " To_Char(Sum(a.ʵ�ʽ��),'99999999999990.99999') �����," & _
                        " To_Char(Sum(a.ʵ�ʲ��),'99999999999990.99999') �����," & _
                        " To_Char(Sum(b.ʵ������),'99999999999990.99') ��������" & _
                        " From ҩƷ��� a ,ҩƷ���� b Where a.����=1 and a.ҩƷid=b.ҩƷid And a.�ⷿid =b.�ⷿid and b.����id=[4] and b.�ڼ�=[5] "
            Else
                gstrSQL = gstrSQL & " (Select a.ҩƷid,Max(�ϴβ���) AS ����,To_Char(Sum(a.��������),'99999999999990.99999') ��������," & _
                        " To_Char(Sum(a.ʵ������),'99999999999990.99999') �������," & _
                        " To_Char(Sum(a.ʵ�ʽ��),'99999999999990.99999') �����," & _
                        " To_Char(Sum(a.ʵ�ʲ��),'99999999999990.99999') �����" & _
                        " From ҩƷ��� a Where a.����=1 "
            End If
        Else
            gstrSQL = gstrSQL & " (Select ҩƷid,' ' ����, '' ��������," & _
                    " '' �������,'' �����,'' �����" & _
                    " From ҩƷ��� a Where ����=1 "
        End If
'        If lng��Ӧ��ID <> 0 Then gstrSQL = gstrSQL & " And (�ϴι�Ӧ��ID Is NULL Or �ϴι�Ӧ��ID=" & lng��Ӧ��ID & ")"
        If LngԴ�ⷿID <> 0 Or LngĿ�ⷿID <> 0 Then
            gstrSQL = gstrSQL & " And a.�ⷿID=" & IIf(LngԴ�ⷿID = 0, "[3]", "[2]") & " Group By a.ҩƷid) S"
        Else
            gstrSQL = gstrSQL & " Group By a.ҩƷid) S"
        End If
        gstrSQL = gstrSQL & ",(Select ҩƷID,�ⷿID,�ⷿ��λ From ҩƷ�����޶� " & _
                  " Where �ⷿID=" & IIf(IntEditState = 2, "[2]", "[3]") & ") E"
        
        '������
        gstrSQL = gstrSQL & " Where D.ҩƷID=P.�շ�ϸĿID And D.ҩƷID=S.ҩƷID "
        '��ϵͳ������ҩƷ�������顱Ϊ�����ֹ���ǳ��ⵥʱ��������Ϊ���ҩƷ
        If Not (IntStockCheck = 2 And In_�༭״̬ = 2) Or bln�̵㵥 Or Not blnCheck Then gstrSQL = gstrSQL & "(+) "
        'If In_�༭״̬ = 2 Then gstrSQL = gstrSQL & " And S.��������<>0"
        gstrSQL = gstrSQL & " And D.ҩƷID=E.ҩƷID(+) Order By D.ҩƷ����"
    End With
    
    Set RecPhysic = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, StrFindStyle & strInput & "%", LngԴ�ⷿID, LngĿ�ⷿID, Lngʹ�ò���ID, Format(zlDatabase.Currentdate(), "yyyy"))
    
    With RecPhysic
        If .EOF Then
            MsgBox "û���ҵ�ƥ���ҩƷ�������ǿ�治���ԭ������ģ������������룡", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    With MsfҩƷ���
        If Not RecPhysic.EOF Then
            Set .DataSource = RecPhysic
        Else
            .Clear
            .Rows = 2
        End If
        Call SetFormat(1, RecPhysic.EOF)
        BlnSelect = (RecPhysic.EOF <> True)
    End With
    
    Call MsfҩƷ���_EnterCell
    RefreshData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function InitRec()
        '������:����
        '��������:2000-11-02
        '��ʼ����¼��
        
        Set RecReturn = New ADODB.Recordset
        With RecReturn
            If .State = 1 Then .Close
            .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "ҩ������", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "ҩƷ��Դ", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "ͨ������", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "ҩ��ID", adDouble, 18, adFldIsNullable
            .Fields.Append "��;����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "������λ", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "ҩƷ����", adLongVarChar, 10, adFldIsNullable
            .Fields.Append "��Ʒ��", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "���", adLongVarChar, 30, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "ҩ��ID", adDouble, 18, adFldIsNullable
            .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
            .Fields.Append "�ۼ�", adDouble, 18, adFldIsNullable
            .Fields.Append "�ۼ۵�λ", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "����ϵ��", adDouble, 11, adFldIsNullable
            .Fields.Append "���Ч��", adDouble, 5, adFldIsNullable
            .Fields.Append "���ﵥλ", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "�����װ", adDouble, 11, adFldIsNullable
            .Fields.Append "סԺ��λ", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "סԺ��װ", adDouble, 11, adFldIsNullable
            .Fields.Append "ҩ�ⵥλ", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "ҩ���װ", adDouble, 11, adFldIsNullable
            .Fields.Append "ҩ�����", adDouble, 2, adFldIsNullable
            .Fields.Append "ҩ������", adDouble, 2, adFldIsNullable
            .Fields.Append "ʱ��", adDouble, 2, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "��������", adDate, , adFldIsNullable
            .Fields.Append "Ч��", adDate, , adFldIsNullable
            .Fields.Append "��������", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "ʵ������", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "ʵ�ʽ��", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "ʵ�ʲ��", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "ָ��������", adDouble, 11, adFldIsNullable
            .Fields.Append "ָ�������", adDouble, 11, adFldIsNullable
            .Fields.Append "�ϴι�Ӧ��ID", adDouble, 18, adFldIsNullable
            .Fields.Append "�������", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "��׼�ĺ�", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "�ɱ���", adDouble, 11, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
End Function

Private Function CombinateRec() As Boolean
    '��װ��¼��
    '��λ��¼��
    Dim blnEof As Boolean                   '�Ƿ���ڿ�����μ�¼
    Dim dblPrice As Double
    Dim rsTemp As New ADODB.Recordset
    Dim strMsg As String
    
    CombinateRec = False
    With RecPhysic
        If .RecordCount <> 0 Then .MoveFirst
        .Find "ҩƷID=" & Val(MsfҩƷ���.TextMatrix(MsfҩƷ���.Row, mconIntColSpecҩƷID))
        If .EOF Then
            MsgBox "�����ڲ�����", vbInformation, gstrSysName
            Exit Function
        End If
        If ((int���� = 3 And int�ⷿ <> 3) Or (int���� = 1 And int�ⷿ = 1) Or (int���� = 2 And int�ⷿ = 2)) And In_�༭״̬ = 2 Then
            With RecStock
                If .RecordCount <> 0 Then .MoveFirst
                .Find "����=" & Val(Msf����.TextMatrix(Msf����.Row, mconIntCol����))
                If .EOF Then
                    blnEof = True
                    If blnPrice Then
                        MsgBox "�����ڲ�����", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End With
        End If
    End With
    
'    '��ȡ��ҩƷ�����۵�λ�۸�
    gstrSQL = "Select �ּ�, B.ָ��������, B.ָ�����ۼ� " & _
            " From �շѼ�Ŀ A, ҩƷ��� B " & _
            " Where A.�շ�ϸĿid = B.ҩƷid And A.�շ�ϸĿID=[1] And Sysdate Between A.ִ������ And Nvl(A.��ֹ����,Sysdate) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ��ҩƷ�����۵�λ�۸�]", CLng(RecPhysic!ҩƷID))
    
    dblPrice = 0
    If Not rsTemp.EOF Then
        dblPrice = Nvl(rsTemp!�ּ�, 0)
    End If
    
    '���ָ�������ۣ�ָ�����ۼۣ�Ϊ0ʱ������Ը�ҩƷ����
    strMsg = ""
    If Not rsTemp.EOF Then
        If rsTemp!ָ�������� = 0 And rsTemp!ָ�����ۼ� = 0 Then
            strMsg = "[" & RecPhysic!ҩ������ & RecPhysic!ͨ������ & "]�ɹ��޼ۺ�ָ���ۼ�Ϊ0���������ü۸�"
        ElseIf rsTemp!ָ�������� = 0 Then
            strMsg = "[" & RecPhysic!ҩ������ & RecPhysic!ͨ������ & "]�ɹ��޼�Ϊ0���������ü۸�"
        ElseIf rsTemp!ָ�����ۼ� = 0 Then
            strMsg = "[" & RecPhysic!ҩ������ & RecPhysic!ͨ������ & "]ָ���ۼ�Ϊ0���������ü۸�"
        End If
    End If
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, gstrSysName
        CombinateRec = False
        Exit Function
    End If
    
'    '����Ƕ���ҩƷ�����ּ۱������0����������Ը�ҩƷ����
'    If IIf(RecPhysic!ʱ�� = "��", 1, 0) = 0 And dblPrice = 0 Then
'        MsgBox "[" & RecPhysic!ҩ������ & RecPhysic!ͨ������ & "]�Ƕ���ҩƷ�������������ۼۡ�", vbInformation, gstrSysName
'        CombinateRec = False
'        Exit Function
'    End If
    
    'װ����д���¼��������������ʹ��
    With RecReturn
        If .EOF Then .AddNew
        !���� = RecPhysic!����
        !ҩ������ = RecPhysic!ҩ������
        !ҩƷ��Դ = RecPhysic!ҩƷ��Դ
        !ͨ������ = RecPhysic!ͨ������
        !ҩ��ID = RecPhysic!ҩ��ID
        !��;����ID = RecPhysic!��;����ID
        !������λ = RecPhysic!������λ
        !ҩƷ���� = RecPhysic!ҩƷ����
        !��Ʒ�� = RecPhysic!��Ʒ��
        !��� = RecPhysic!���
        !���� = RecPhysic!����
'        !���� = RecStock!����
'        !Ч�� = RecStock!ʧЧ��
        !ҩ��ID = RecPhysic!ҩ��ID
        !ҩƷID = RecPhysic!ҩƷID
        !�ۼ� = dblPrice
        !�ۼ۵�λ = RecPhysic!�ۼ۵�λ
        !����ϵ�� = RecPhysic!����ϵ��
        !���Ч�� = RecPhysic!��Ч��
        !���ﵥλ = RecPhysic!���ﵥλ
        !�����װ = RecPhysic!�����װ
        !סԺ��λ = RecPhysic!סԺ��λ
        !סԺ��װ = RecPhysic!סԺ��װ
        !ҩ�ⵥλ = RecPhysic!ҩ�ⵥλ
        !ҩ���װ = RecPhysic!ҩ���װ
        !ҩ����� = IIf(RecPhysic!ҩ����� = "��", 1, 0)
        !ҩ������ = IIf(RecPhysic!ҩ������ = "��", 1, 0)
        !ʱ�� = IIf(RecPhysic!ʱ�� = "��", 1, 0)
        !�ϴι�Ӧ��ID = 0
        !��׼�ĺ� = IIf(IsNull(RecPhysic!��׼�ĺ�), "", RecPhysic!��׼�ĺ�)
        If In_�༭״̬ = 2 And ((int���� = 3 And int�ⷿ <> 3) Or (int���� = 1 And int�ⷿ = 1) Or (int���� = 2 And int�ⷿ = 2)) Then
            If Msf����.TextMatrix(Msf����.Row, mconIntCol����) = "��������ҩƷ" Then
                !���� = -1
            Else
                If Not blnEof Then
                    !���� = Val(RecStock!����)
                    !���� = RecStock!����
                    !�������� = RecStock!��������
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 0 Then
                        !Ч�� = RecStock!ʧЧ��
                    Else
                        !Ч�� = RecStock!��Ч����
                    End If
                
                    !���� = Nvl(RecStock!����)
                    !�ϴι�Ӧ��ID = Nvl(RecStock!�ϴι�Ӧ��ID, 0)
                    !�������� = IIf(IsNull(RecStock!��������), 0, RecStock!��������)
                    !ʵ������ = IIf(IsNull(RecStock!�������), 0, RecStock!�������)
                    !ʵ�ʽ�� = IIf(IsNull(RecStock!�����), 0, RecStock!�����)
                    !ʵ�ʲ�� = IIf(IsNull(RecStock!�����), 0, RecStock!�����)
                    !������� = IIf(IsNull(RecStock!ʵ������), 0, RecStock!ʵ������)
                    !��׼�ĺ� = IIf(IsNull(RecStock!��׼�ĺ�), "", RecStock!��׼�ĺ�)
                    If Not blnStock Then Call Get���ÿ��(!ҩƷID, !����)
                End If
            End If
        Else
            !�������� = IIf(IsNull(RecPhysic!��������), 0, RecPhysic!��������)
            !ʵ������ = IIf(IsNull(RecPhysic!�������), 0, RecPhysic!�������)
            !ʵ�ʽ�� = IIf(IsNull(RecPhysic!�����), 0, RecPhysic!�����)
            !ʵ�ʲ�� = IIf(IsNull(RecPhysic!�����), 0, RecPhysic!�����)
            !������� = IIf(IsNull(RecPhysic!ʵ������), 0, RecPhysic!ʵ������)
            '��ȡ������ҩƷ��������Ч����Ϣ
            gstrSQL = "Select �ϴ�����,Ч��,�ϴι�Ӧ��id,�ϴ��������� AS ��������,��׼�ĺ� From ҩƷ��� " & _
                    " Where �ⷿID=[1] And ҩƷID=[2] And ����=1 "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ������ҩƷ��������Ч����Ϣ]", LngԴ�ⷿID, CLng(RecPhysic!ҩƷID))
            
            If rsTemp.RecordCount <> 0 Then
                !���� = Nvl(rsTemp!�ϴ�����)
                !�ϴι�Ӧ��ID = Nvl(rsTemp!�ϴι�Ӧ��ID, 0)
                If Not IsNull(rsTemp!��������) Then
                    !�������� = rsTemp!��������
                End If
                If Not IsNull(rsTemp!Ч��) Then
                    !Ч�� = rsTemp!Ч��
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And Nvl(!Ч��) <> "" Then
                        '����Ϊ��Ч��
                        !Ч�� = Format(DateAdd("D", -1, !Ч��), "yyyy-mm-dd")
                    End If
                End If
                !��׼�ĺ� = IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
            End If
            
            If Not blnStock Then Call Get���ÿ��(!ҩƷID, 0)
        End If
        
        '�������ʾ�Է��ⷿ�Ŀ�棬��������ȡ������
        If Not blnStock Then
            !�������� = mdbl��������
            !ʵ������ = mdblʵ������
            !ʵ�ʽ�� = mdblʵ�ʽ��
            !ʵ�ʲ�� = mdblʵ�ʲ��
            !������� = mdbl�������
        End If
        
        !ָ�������� = RecPhysic!ָ��������
        !ָ������� = RecPhysic!ָ�������
        !�ɱ��� = IIf(Val(RecPhysic!��ʼ�ɱ���) = 0, Val(RecPhysic!ָ��������), RecPhysic!��ʼ�ɱ���)
        
        .Update
    End With
    
    CombinateRec = True
End Function

Private Function CheckData() As Boolean
    Dim DblCurStock As Double       '��ǰ�����
    '����Ƿ�����ѡ��
    CheckData = False
    
    If BlnSelect = False Then Exit Function
    
    'lng��Ӧ��ID��Ϊ�㣬��ʾ�˻����޿��ʱ��׼����
    If lng��Ӧ��ID <> 0 Then
        If Msf����.Visible Then
            If Val(Msf����.TextMatrix(Msf����.Row, mconIntCol�ϴι�Ӧ��ID)) <> 0 And lng��Ӧ��ID <> Val(Msf����.TextMatrix(Msf����.Row, mconIntCol�ϴι�Ӧ��ID)) Then
                MsgBox "��ѡ����˻��̲��Ǹ�ҩƷ�Ĺ�Ӧ�̣����ܼ���������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    If Msf����.Visible Then
        If blnStock Then
            DblCurStock = Val(Msf����.TextMatrix(Msf����.Row, mconIntCol��������))
        Else
            DblCurStock = Get���ÿ��(Val(MsfҩƷ���.TextMatrix(MsfҩƷ���.Row, mconIntColSpecҩƷID)), Val(Msf����.TextMatrix(Msf����.Row, mconIntCol����)))
        End If
    Else
        If Not RecPhysic.EOF Then
            If blnStock Then
                DblCurStock = Val(MsfҩƷ���.TextMatrix(MsfҩƷ���.Row, mconIntColSpec��������))
            Else
                DblCurStock = Get���ÿ��(Val(MsfҩƷ���.TextMatrix(MsfҩƷ���.Row, mconIntColSpecҩƷID)))
            End If
        End If
    End If
    
    If DblCurStock > 0 Then
        CheckData = True
        Exit Function
    End If
    
    '���Դ�ⷿ��Ŀ�ⷿΪ�գ��������ҩƷĿ¼�Լ��ڽ��г������ã����ж�
    If (LngԴ�ⷿID = 0 And LngĿ�ⷿID = 0) Then
        CheckData = True
        Exit Function
    End If
    
    '������̵㵥����ҩƷѡ�����������жϣ�ֱ���˳�
    If bln�̵㵥 Then
        CheckData = True
        Exit Function
    End If
    
    '�����ҩƷ����۵����������жϣ�ֱ���˳�
    If glngModul = 1303 Then
        CheckData = True
        Exit Function
    End If
    
    If Msf����.Visible Or blnʱ�� Then
        If (DblCurStock > 0) Or Not blnPrice Or Msf����.TextMatrix(Msf����.Row, mconIntCol����) = "��������ҩƷ" Then CheckData = True: Exit Function
        MsgBox "��" & IIf(blnʱ��, "ʱ��", "����") & "ҩƷ�Ѿ�û�п�棬���ܼ���������", vbInformation, gstrSysName
        Exit Function
    Else
        If blnCheck = False Then
           CheckData = True
           Exit Function
        End If
    End If
    
    Select Case IntStockCheck
    Case 1
        If MsgBox("��ҩƷ�Ѿ�û�п�棬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Case 2
        MsgBox "��ҩƷ�Ѿ�û�п�棬���ܼ���������", vbInformation, gstrSysName
        Exit Function
    End Select
    CheckData = True
End Function

Public Function ShowME(ByVal FrmMain As Form, ByVal �༭ģʽ As Integer, Optional ByVal Դ�ⷿ As Long = 0, _
                    Optional ByVal Ŀ�ⷿ As Long = 0, Optional ByVal ʹ�ò��� As Long = 0, Optional ByVal ��ѯ�� As String = "", _
                    Optional ByVal WinLeft As Double = 0, Optional ByVal WinTop As Double = 0, Optional ByVal Bln����� As Boolean = True, _
                    Optional ByVal bln������λ�ʱ�� As Boolean = True, Optional ByVal bln�̵㵥�� As Boolean = False, Optional ByVal bln���ӿ����� As Boolean = False, _
                    Optional ByVal bln��ʾ��� As Boolean = True, Optional ByVal lng��Ӧ�� As Long = 0, Optional ByVal bln���޴洢�ⷿҩƷ As Boolean = False, _
                    Optional ByVal ���÷�ʽ As Integer = 0, Optional ByVal bln����ͣ��ҩƷ As Boolean = False) As ADODB.Recordset
    On Error Resume Next
    'bln�����:��������ҩƷ��ʱ��ҩƷ���治׼����ԭ�򣬿�ǿ������not (���� or ʱ��) ҩƷ����
    'bln������λ�ʱ��:�������������ҩƷ��ʱ��ҩƷ����
    'lng��Ӧ��ID:��Ϊ���ʾ�˻�
    
    With Me
        WindowPosition.Left = WinLeft
        WindowPosition.Top = WinTop
        .In_�༭״̬ = �༭ģʽ
        .In_Դ�ⷿ = Դ�ⷿ
        .In_Ŀ�ⷿ = Ŀ�ⷿ
        .In_���� = ʹ�ò���
        .In_�ִ� = UCase(Trim(��ѯ��))
        .In_MainFrm = FrmMain
        bln�̵㵥 = bln�̵㵥��
        bln������ = bln���ӿ�����
        blnCheck = Bln�����
        blnPrice = bln������λ�ʱ��
        blnStock = bln��ʾ���
        lng��Ӧ��ID = lng��Ӧ��
        mblnNoStock = bln���޴洢�ⷿҩƷ
        int���÷�ʽ = ���÷�ʽ
        mbln����ͣ��ҩƷ = bln����ͣ��ҩƷ
        .Show 1, FrmMain
    End With
    Set ShowME = RecReturn.Clone
End Function

Public Function Get���ÿ��(ByVal lngҩƷID As Long, Optional ByVal lng���� As Long = 0) As Single
    Dim rsStock As New ADODB.Recordset
    
    gstrSQL = " Select Sum(A.��������" & StrUnitString & ") ��������,Sum(A.ʵ������" & StrUnitString & ") ʵ������,sum(A.ʵ�ʽ��) ʵ�ʽ��,sum(A.ʵ�ʲ��) ʵ�ʲ��,Sum(A.ʵ������) ������� " & _
              " From ҩƷ��� A,ҩƷ��� B " & _
              " Where A.ҩƷID=B.ҩƷID And A.����=1 And A.ҩƷID=[1] " & IIf(lng���� = 0, "", " And Nvl(A.����,0)=[2] ")
    If LngԴ�ⷿID <> 0 Or LngĿ�ⷿID <> 0 Then
        gstrSQL = gstrSQL & " And A.�ⷿID=[3] "
    End If
    gstrSQL = gstrSQL & " Group By A.ҩƷid"
    
    Set rsStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ���ÿ��]", lngҩƷID, lng����, IIf(LngԴ�ⷿID = 0, LngĿ�ⷿID, LngԴ�ⷿID))
    
    mdbl�������� = 0
    mdblʵ�ʲ�� = 0
    mdblʵ�ʽ�� = 0
    mdblʵ������ = 0
    mdbl������� = 0
    If Not rsStock.EOF Then
        mdbl�������� = IIf(IsNull(rsStock!��������), 0, rsStock!��������)
        mdblʵ�ʲ�� = IIf(IsNull(rsStock!ʵ�ʲ��), 0, rsStock!ʵ�ʲ��)
        mdblʵ�ʽ�� = IIf(IsNull(rsStock!ʵ�ʽ��), 0, rsStock!ʵ�ʽ��)
        mdblʵ������ = IIf(IsNull(rsStock!ʵ������), 0, rsStock!ʵ������)
        mdbl������� = IIf(IsNull(rsStock!�������), 0, rsStock!�������)
    End If
    Get���ÿ�� = mdbl��������
End Function

