VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngID As Long
Private mobjConn As clsConnect
Private mlngDeptID As Long
Private mblnEnabled As Boolean
Private mdatTime As Date
Private mstrType As String
Private mstrBill As String
Private mbytObject As Byte
Private mbytDispenseFunc As Byte
Private mbytDispensingFunc As Byte

'�豸�����Ӷ���
Public Property Get Link() As clsConnect
    Link = mobjConn
End Property
Public Property Let Link(ByVal objConn As clsConnect)
    mobjConn = objConn
End Property

'�豸ʹ�ò���ID
Public Property Get DeptID() As Long
    DeptID = mlngDeptID
End Property
Public Property Let DeptID(ByVal lngDeptID As Long)
    mlngDeptID = lngDeptID
End Property

'�豸����ID�����豸ID��ͬ
Public Property Get ID() As Long
    ID = mlngID
End Property
Public Property Let ID(ByVal lngID As Long)
    mlngID = lngID
End Property

'�豸����
Public Property Get Enabled() As Boolean
    Enabled = mblnEnabled
End Property
Public Property Let Enabled(ByVal blnEnabled As Boolean)
    mblnEnabled = blnEnabled
End Property

'�豸���á�������Чʱ��
'Public Property Get EnabledTime() As Date
'    EnabledTime = mdatTime
''    Dim rsTmp As ADODB.Recordset
''
''    EnabledTime = CDate("2000-1-1 00:00:00")
''
''    On Error GoTo errHandle
''    gstrSQL = "select a.ID, a.��Чʱ�� " & _
''              "from ҩ��ע���豸 a, ҩ���豸���� b " & _
''              "where a.ID = b.�豸ID and a.����ID = [1] and a.������ = [2] and b.������ = 1 "
''    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ���豸", DeptID, Link.Name)
''    Do While rsTmp.EOF = False
''
''        rsTmp.MoveNext
''    Loop
''    rsTmp.Close
''
''    Exit Property
''
''errHandle:
''    If gobjComLib.ErrCenter = 1 Then Resume
'End Property
'Public Property Let EnabledTime(ByVal datTime As Date)
'    EnabledTime = mdatTime
'End Property

'�豸֧�ֵ�ҩƷ����
Public Property Get DrugType() As String
    'TypeName�ж�����
    DrugType = mstrType
End Property
Public Property Let DrugType(ByVal strType As String)
    mstrType = strType
End Property

'�豸֧�ֵķ������
Public Property Get ServiceObject() As Byte
    ServiceObject = mbytObject
End Property
Public Property Let ServiceObject(ByVal bytObject As Byte)
    mbytObject = bytObject
End Property

'�豸֧�ֵĵ���
Public Property Get Bill() As String
    Bill = mstrBill
End Property
Public Property Let Bill(ByVal strBill As String)
    mstrBill = strBill
End Property

'��ҩ��ӦHISҵ��
Public Property Get DispenseFunc() As Byte
    DispenseFunc = mbytDispenseFunc
End Property
Public Property Let DispenseFunc(ByVal bytDispenseFunc As Byte)
    mbytDispenseFunc = bytDispenseFunc
End Property

'��ҩ��ӦHISҵ��
Public Property Get DispensingFunc() As Byte
    DispensingFunc = mbytDispensingFunc
End Property
Public Property Let DispensingFunc(ByVal bytDispensingFunc As Byte)
    mbytDispensingFunc = bytDispensingFunc
End Property