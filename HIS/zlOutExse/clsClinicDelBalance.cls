VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCliniDelBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'��������ֵ�ľֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar�������� As String '�ֲ�����
Private mvarCurDelNos As String '�ֲ�����
Private mvarAllNos As String '�ֲ�����
Private mvarPrintNOs As String '�ֲ�����
Private mvar�˷�ʱ�� As Date   '�ֲ�����
Private mvar�����˷� As Boolean   '�ֲ�����
Private mvarShareUserID As Long '�ֲ�����
Private mvarPatiUseType As String '�ֲ�����
Private mvar�ѱ� As String '�ֲ�����
Private mvar�Ա� As String '�ֲ�����
Private mvar�˷ѽ��� As String '�ֲ�����
Private mvar��ǰ��Ʊ�� As String '�ֲ�����
Private mvar���շ�Ʊ As String '�ֲ�����
Private mvarintInsure As Integer '�ֲ�����
Private mvarҽ�������� As Double '�ֲ�����
Private mvar�˷Ѻϼ� As Double '�ֲ�����
Private mvarҽ������Ʊ�� As Boolean   '�ֲ�����
Private mvar�ɿ� As Double '�ֲ�����
Private mvar�Ҳ� As Double '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar����ID As Long  '�ֲ�����
Private mvarԭ����ID As Long '�ֲ�����
Private mvar������� As Long '�ֲ�����
Private mvarrsBalance As ADODB.Recordset   '�ֲ�����
Private mvarrs���㷽ʽ As ADODB.Recordset   '�ֲ�����
Private mvarȱʡ���㷽ʽ As String '�ֲ�����
Private mvarSaveBilled As Boolean   '�ֲ�����
Private mvarԭ���� As Boolean   '�ֲ�����
Private mvar����IDs As String '�ֲ�����
Private mvarblnOnePatiPrint As Boolean   '�ֲ�����
Private mvarstrOnePatiPrintNos As String     '�ֲ�����
Public Property Let strOnePatiPrintNos(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.strOnePatiPrintNos = 5
    mvarstrOnePatiPrintNos = vData
End Property


Public Property Get strOnePatiPrintNos() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.strOnePatiPrintNos
    strOnePatiPrintNos = mvarstrOnePatiPrintNos
End Property

Public Property Let blnOnePatiPrint(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.blnOnePatiPrint = 5
    mvarblnOnePatiPrint = vData
End Property


Public Property Get blnOnePatiPrint() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.blnOnePatiPrint
    blnOnePatiPrint = mvarblnOnePatiPrint
End Property

Public Property Let ����IDs(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����IDs = 5
    mvar����IDs = vData
End Property


Public Property Get ����IDs() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����IDs
    ����IDs = mvar����IDs
End Property


Public Property Let ԭ����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ԭ���� = 5
    mvarԭ���� = vData
End Property


Public Property Get ԭ����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ԭ����
    ԭ���� = mvarԭ����
End Property

Public Property Let SaveBilled(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SaveBilled = 5
    mvarSaveBilled = vData
End Property


Public Property Get SaveBilled() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SaveBilled
    SaveBilled = mvarSaveBilled
End Property
Public Property Let AllNos(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.AllNos = 5
    mvarAllNos = vData
End Property
Public Property Get AllNos() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.AllNos
    AllNos = mvarAllNos
End Property
Public Property Let CurDelNos(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CurDelNos = 5
    mvarCurDelNos = vData
End Property


Public Property Get CurDelNos() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CurDelNos
    CurDelNos = mvarCurDelNos
End Property
 
Public Property Let ���շ�Ʊ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���շ�Ʊ = 5
    mvar���շ�Ʊ = vData
End Property


Public Property Get ���շ�Ʊ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���շ�Ʊ
    ���շ�Ʊ = mvar���շ�Ʊ
End Property
Public Property Let �˷�ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�˷�ʱ�� = 5
    mvar�˷�ʱ�� = vData
End Property


Public Property Get �˷�ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�˷�ʱ��
    �˷�ʱ�� = mvar�˷�ʱ��
End Property


Public Property Let �Ҳ�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ҳ� = 5
    mvar�Ҳ� = vData
End Property


Public Property Get �Ҳ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ҳ�
    �Ҳ� = mvar�Ҳ�
End Property



Public Property Let �ɿ�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ɿ� = 5
    mvar�ɿ� = vData
End Property


Public Property Get �ɿ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ɿ�
    �ɿ� = mvar�ɿ�
End Property
Public Property Let ҽ��������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ�������� = 5
    mvarҽ�������� = vData
End Property


Public Property Get ҽ��������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ҽ��������
    ҽ�������� = mvarҽ��������
End Property


Public Property Let �˷Ѻϼ�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�˷Ѻϼ� = 5
    mvar�˷Ѻϼ� = vData
End Property


Public Property Get �˷Ѻϼ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�˷Ѻϼ�
    �˷Ѻϼ� = mvar�˷Ѻϼ�
End Property


Public Property Let intInsure(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.intInsure = 5
    mvarintInsure = vData
End Property


Public Property Get intInsure() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.intInsure
    intInsure = mvarintInsure
End Property

Public Property Let �Ա�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ա� = 5
    mvar�Ա� = vData
End Property


Public Property Get �Ա�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ա�
    �Ա� = mvar�Ա�
End Property



Public Property Let �ѱ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ѱ� = 5
    mvar�ѱ� = vData
End Property


Public Property Get �ѱ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ѱ�
    �ѱ� = mvar�ѱ�
End Property
Public Property Let �˷ѽ���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�˷ѽ��� = 5
    mvar�˷ѽ��� = vData
End Property
Public Property Get �˷ѽ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�˷ѽ���
    �˷ѽ��� = mvar�˷ѽ���
End Property
 
Public Property Let PrintNOs(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PrintNOs = 5
    mvarPrintNOs = vData
End Property
Public Property Get PrintNOs() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PrintNOs
    PrintNOs = mvarPrintNOs
End Property
 
 
Public Property Let ��ǰ��Ʊ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ǰ��Ʊ�� = 5
    mvar��ǰ��Ʊ�� = vData
End Property
Public Property Get ��ǰ��Ʊ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ǰ��Ʊ��
    ��ǰ��Ʊ�� = mvar��ǰ��Ʊ��
End Property
 
Public Property Let PatiUseType(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PatiUseType = 5
    mvarPatiUseType = vData
End Property


Public Property Get PatiUseType() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PatiUseType
    PatiUseType = mvarPatiUseType
End Property



Public Property Let ShareUserID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ShareUserID = 5
    mvarShareUserID = vData
End Property


Public Property Get ShareUserID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ShareUserID
    ShareUserID = mvarShareUserID
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property
Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property



Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property
 

Public Property Let ��������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�������� = vData
End Property


Public Property Get ��������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    �������� = mvar��������
End Property

Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar����ID = vData
End Property


Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ID
    ����ID = mvar����ID
End Property

Public Property Get ����ID() As Long
    ����ID = mvar����ID
End Property

Public Property Let ����ID(ByVal vData As Long)
    mvar����ID = vData
End Property


Public Property Get �������() As Long
    ������� = mvar�������
End Property

Public Property Let �������(ByVal vData As Long)
    mvar������� = vData
End Property
Public Property Let ԭ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ԭ����ID = 5
    mvarԭ����ID = vData
End Property

Public Property Get ԭ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ԭ����ID
    ԭ����ID = mvarԭ����ID
End Property


Public Property Let ҽ������Ʊ��(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ������Ʊ�� = 5
    mvarҽ������Ʊ�� = vData
End Property

Public Property Get ҽ������Ʊ��() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ҽ������Ʊ��
    ҽ������Ʊ�� = mvarҽ������Ʊ��
End Property
Public Property Get ����ID() As Long
   ����ID = mvar����ID
End Property

Public Property Let ����ID(ByVal vData As Long)
   mvar����ID = vData
End Property

 

Public Property Get rsBalance() As ADODB.Recordset
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.rsBalance
   Set rsBalance = mvarrsBalance
End Property

Public Property Set rsBalance(ByVal vData As ADODB.Recordset)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.rs���㷽ʽ = 5
    Set mvarrsBalance = vData
End Property


Public Property Set rs���㷽ʽ(ByVal vData As ADODB.Recordset)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.rs���㷽ʽ = 5
    Set mvarrs���㷽ʽ = vData
End Property

Public Property Get rs���㷽ʽ() As ADODB.Recordset
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.rs���㷽ʽ
   Set rs���㷽ʽ = mvarrs���㷽ʽ
End Property

Public Property Let ȱʡ���㷽ʽ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ȱʡ���㷽ʽ = 5
      mvarȱʡ���㷽ʽ = vData
End Property

Public Property Get ȱʡ���㷽ʽ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ȱʡ���㷽ʽ
    ȱʡ���㷽ʽ = mvarȱʡ���㷽ʽ
End Property
Public Property Let �����˷�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�����˷� = 5
      mvar�����˷� = vData
End Property

Public Property Get �����˷�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�����˷�
    �����˷� = mvar�����˷�
End Property

 