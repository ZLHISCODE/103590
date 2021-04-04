VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstrTime As String  '���˽��ʴ���(��ʼ="",����Ϊ"1,2,3...")
Private mdtBeginDate As Date       '���˽��ʵĿ�ʼʱ��,��ʼΪ'1900-01-01'
Private mdtEndDate As Date         '���˽��ʵĽ���ʱ��,��ʼΪ'3000-01-01'
Private mstrDeptIDs As String      '���˽��ʿ���ID��(��ʼ="",����Ϊ"0,1,2,3...",0��ʾ��������IDΪ��)
Private mstrClass As String       '��������=""-���з���(��δ����),"'����','����',..."
Private mstrChargeType As String      '�շ���� '34260
Private mstrBaby As String      '�Ƿ������Ӥ������(0-���з���,1-���˷���,2������-��mbytbaby-1��Ӥ������)
Private mstrItem As String      'Ҫ����վݷ�Ŀ
Private mbytKind As Byte       '0-����ͨ����,1-��������,2-��ͨ���ú�������
Private mblnCurBalanceOwnerFee As Boolean      '��ǰ�Ƿ����ڽᡰ�Էѷ��á�
Private mstrDiag As String
Private mlng����ID As Long
'��������ֵ�ľֲ�����
Private mvarblnNormalFee As Boolean '�ֲ�����
Private mvarblnHealthCheckFee As Boolean '�ֲ�����
Public Property Let blnHealthCheckFee(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.blnHealthCheckFee = 5
    mvarblnHealthCheckFee = vData
End Property


Public Property Get blnHealthCheckFee() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.blnHealthCheckFee
    blnHealthCheckFee = mvarblnHealthCheckFee
End Property



Public Property Let blnNormalFee(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.blnNormalFee = 5
    mvarblnNormalFee = vData
End Property


Public Property Get blnNormalFee() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.blnNormalFee
    blnNormalFee = mvarblnNormalFee
End Property




Public Property Get lng����ID() As Long
    lng����ID = mlng����ID
End Property

Public Property Let lng����ID(ByVal vNewValue As Long)
    mlng����ID = vNewValue
End Property

Public Property Get strTime() As String
    strTime = mstrTime
End Property

Public Property Let strTime(ByVal vNewValue As String)
    mstrTime = vNewValue
End Property

Public Property Get strDiag() As String
    strDiag = mstrDiag
End Property

Public Property Let strDiag(ByVal vNewValue As String)
    mstrDiag = vNewValue
End Property

Public Property Get strDeptIDs() As String
    strDeptIDs = mstrDeptIDs
End Property

Public Property Let strDeptIDs(ByVal vNewValue As String)
    mstrDeptIDs = vNewValue
End Property

Public Property Get strItem() As String
    strItem = mstrItem
End Property

Public Property Let strItem(ByVal vNewValue As String)
    mstrItem = vNewValue
End Property

Public Property Get strClass() As String
    strClass = mstrClass
End Property

Public Property Let strClass(ByVal vNewValue As String)
    mstrClass = vNewValue
End Property

Public Property Get strChargeType() As String
    strChargeType = mstrChargeType
End Property

Public Property Let strChargeType(ByVal vNewValue As String)
    mstrChargeType = vNewValue
End Property

Public Property Get strBaby() As String
    strBaby = mstrBaby
End Property

Public Property Let strBaby(ByVal vNewValue As String)
    mstrBaby = vNewValue
End Property

Public Property Get bytKind() As Byte
    bytKind = mbytKind
End Property

Public Property Let bytKind(ByVal vNewValue As Byte)
    mbytKind = vNewValue
End Property

Public Property Get blnCurBalanceOwnerFee() As Boolean
    blnCurBalanceOwnerFee = mblnCurBalanceOwnerFee
End Property

Public Property Let blnCurBalanceOwnerFee(ByVal vNewValue As Boolean)
    mblnCurBalanceOwnerFee = vNewValue
End Property

Public Property Get dtBeginDate() As Date
    dtBeginDate = mdtBeginDate
End Property

Public Property Let dtBeginDate(ByVal vNewValue As Date)
    mdtBeginDate = vNewValue
End Property

Public Property Get dtEndDate() As Date
    dtEndDate = mdtEndDate
End Property

Public Property Let dtEndDate(ByVal vNewValue As Date)
    mdtEndDate = vNewValue
End Property