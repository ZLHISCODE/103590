VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPubDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'------------------------
'���ӿڲ��������Ϣ
Public Type TYPE_PASSPati
    lng����ID As Long
    lng��ҳID As Long
    lng�Һ�ID As Long
    dbl��ʶ�� As Double  '���ﴫ�����,סԺ ��סԺ��
    str���� As String
    str�Ա� As String
    str����  As String
    intӤ�� As Integer
    str�Һŵ� As String '����
    Dat�������� As Date
End Type

'�����Ϣ
Public Type TYPE_PASSDiag
    objDiag As Object
    lng�������� As Long
    lng��ϱ��� As Long
    lng������� As Long
End Type

'ӳ��vsAdvice�к�,δӳ����д�ֵΪ-1
Public Type TYPE_PASSCOL
    intCOLID As Integer
    intCOL���ID As Integer
    intCOL������� As Integer
    intCOL������ĿID As Integer
    intCOL�շ�ϸĿID As Integer
    intCOLҽ������ As Integer
    intCOL��Ч As Integer       'סԺ
    intCOL���� As Integer
    intCOL������λ As Integer
    intCOL�÷� As Integer
    intCOL���� As Integer
    intCOLӤ�� As Integer
    intCOL����ʱ�� As Integer
    intCOL����ҽ�� As Integer
    intCOL��ʼʱ�� As Integer
    intCOL��������ID As Integer
    intCOL��ֹʱ�� As Integer
    intCOLƵ�� As Integer
    intCOLƵ�ʴ��� As Integer
    intCOLƵ�ʼ�� As Integer
    intCOL�����λ As Integer
    intCOL��ʾ As Integer
    intCOL��� As Integer
    intCOL״̬ As Integer
    intCOLEDIT As Integer
    intCOL����ID As Integer
    intCOL��ҳID As Integer
    intCOLѡ�� As Integer       'gbytUseType =1ʱ����
    
    intCOLִ������ As Integer   '
    intCOLҩƷ���� As Integer   '������ҩ����
    'ΪҩƷģ���ṩ
    intCOL���� As Integer
    intCOLNO As Integer
    intCOLסԺ�� As Integer
    intCOL����� As Integer
    intCOL����� As Integer
End Type

'���˽ӿڵ�ģ�����
Public Type TYPE_PASS_MainPara
    lngModel As Long                 '���û���
    udtCOL As TYPE_PASSCOL          'ҽ����ӳ��
    objAdvice As Object             'ҽ���б����� vsAdvice
    objCommandBar As Object         '�˵�����
    objCmdAlley As Object           'PASS����ʷ��ť
    bytUseType As Byte              'bytUseType =1 �ٴ�·������ҽ��ʱ����סԺҽ���༭����
    frmMain As Object               '������
    '�����Ϣ  ���ﴫ��
    udtDiag As TYPE_PASSDiag        '������϶���
End Type

'----------------------------
'����ģ����
Public Enum Enu_Model
    m����༭ = 0
    MסԺ�༭ = 1
    MסԺҽ���嵥 = 2
    M��ʿУ�� = 3
    M����ҽ���嵥 = 4
    MסԺ��ҳ = 5
    M������ҩ = 1341        '1341    ҩƷ������ҩ
    M���ŷ�ҩ = 1342        '1342    ҩƷ���ŷ�ҩ
    MPIVA���� = 1345        '1345    PIVA����
End Enum
