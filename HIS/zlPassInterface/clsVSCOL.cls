VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVSCOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'��������ֵ�ľֲ�����
Private mintCOLID As Integer '�ֲ�����
Private mintCOL���ID As Integer '�ֲ�����
Private mintCOL������� As Integer '�ֲ�����
Private mIntCOL������ĿID As Integer '�ֲ�����
Private mintCOL�շ�ϸĿID As Integer '�ֲ�����
Private mintCOLҽ������ As Integer '�ֲ�����
Private mintCOL��Ч As Integer '�ֲ�����
Private mintCOL���� As Integer '�ֲ�����
Private mintCOL������λ As Integer '�ֲ�����
Private mintCOL�÷� As Integer

Private mintCOL���� As Integer
Private mintCOLӤ�� As Integer
Private mintCOL����ʱ�� As Integer
Private mintCOL����ҽ�� As Integer
Private mintCOL��ʼʱ�� As Integer
Private mintCOL��������ID As Integer
Private mintCOL��ֹʱ�� As Integer
Private mintCOLƵ�� As Integer
Private mintCOLƵ�ʴ��� As Integer
Private mintCOLƵ�ʼ�� As Integer

Private mintCOL�����λ As Integer
Private mintCOL��ʾ As Integer
Private mintCOL��� As Integer
Private mintCOL״̬ As Integer
Private mintCOLEDIT As Integer
Private mintCOL����ID As Integer
Private mintCOL��ҳID As Integer
Private mintCOLѡ�� As Integer       'gbytUseType =1ʱ����

Private mintCOLִ������ As Integer   '
Private mintCOLҩƷ���� As Integer   '������ҩ����
Private mintCOL���� As Integer
Private mintCol������λ As Integer
Private mintColҽ������ As Integer
Private mintCol��ҩĿ�� As Integer
Private mintCol�������� As Integer
Private mintCol����ҩƷ˵�� As Integer
Private mintCol��ҩ����  As Integer
Private mintCol��־  As Integer
Private mintCol������  As Integer
Private mintColִ�з���  As Integer
Private mintColִ�п���ID  As Integer

'ΪҩƷģ���ṩ
Private mintCOL���� As Integer
Private mintCOLNO As Integer
Private mintCOLסԺ�� As Integer
Private mintCOL����� As Integer
Private mintCOL����� As Integer

Public Property Let intCOL������λ(ByVal intData As Integer)
    mintCOL������λ = intData
End Property

Public Property Get intCOL������λ() As Integer
    intCOL������λ = mintCOL������λ
End Property

Public Property Let intCOL����(ByVal intData As Integer)
    mintCOL���� = intData
End Property

Public Property Get intCOL����() As Integer
    intCOL���� = mintCOL����
End Property

Public Property Let intCOL��Ч(ByVal intData As Integer)
    mintCOL��Ч = intData
End Property

Public Property Get intCOL��Ч() As Integer
    intCOL��Ч = mintCOL��Ч
End Property

Public Property Let intCOLҽ������(ByVal intData As Integer)
    mintCOLҽ������ = intData
End Property

Public Property Get intCOLҽ������() As Integer
    intCOLҽ������ = mintCOLҽ������
End Property

Public Property Let intCOL�շ�ϸĿID(ByVal intData As Integer)
    mintCOL�շ�ϸĿID = intData
End Property

Public Property Get intCOL�շ�ϸĿID() As Integer
    intCOL�շ�ϸĿID = mintCOL�շ�ϸĿID
End Property

Public Property Let intCOL������ĿID(ByVal intData As Integer)
    mIntCOL������ĿID = intData
End Property

Public Property Get intCOL������ĿID() As Integer
    intCOL������ĿID = mIntCOL������ĿID
End Property

Public Property Let intCOL�������(ByVal intData As Integer)
    mintCOL������� = intData
End Property

Public Property Get intCOL�������() As Integer
    intCOL������� = mintCOL�������
End Property

Public Property Let intCOL���ID(ByVal intData As Integer)
    mintCOL���ID = intData
End Property

Public Property Get intCOL���ID() As Integer
    intCOL���ID = mintCOL���ID
End Property

Public Property Let intCOLID(ByVal intData As Integer)
    mintCOLID = intData
End Property

Public Property Get intCOLID() As Integer
    intCOLID = mintCOLID
End Property

Public Property Let intCOL�÷�(ByVal intData As Integer)
    mintCOL�÷� = intData
End Property


Public Property Get intCOL�÷�() As Integer
    intCOL�÷� = mintCOL�÷�
End Property

Public Property Let intCOL����(ByVal intData As Integer)
    mintCOL���� = intData
End Property

Public Property Get intCOL����() As Integer
    intCOL���� = mintCOL����
End Property

Public Property Let intCOLӤ��(ByVal intData As Integer)
    mintCOLӤ�� = intData
End Property

Public Property Get intCOLӤ��() As Integer
    intCOLӤ�� = mintCOLӤ��
End Property

Public Property Let intCOL����ʱ��(ByVal intData As Integer)
    mintCOL����ʱ�� = intData
End Property

Public Property Get intCOL����ʱ��() As Integer
    intCOL����ʱ�� = mintCOL����ʱ��
End Property

Public Property Let intCOL����ҽ��(ByVal intData As Integer)
    mintCOL����ҽ�� = intData
End Property

Public Property Get intCOL����ҽ��() As Integer
    intCOL����ҽ�� = mintCOL����ҽ��
End Property

Public Property Let intCOL��ʼʱ��(ByVal intData As Integer)
    mintCOL��ʼʱ�� = intData
End Property

Public Property Get intCOL��ʼʱ��() As Integer
    intCOL��ʼʱ�� = mintCOL��ʼʱ��
End Property

Public Property Let intCOL��������ID(ByVal intData As Integer)
    mintCOL��������ID = intData
End Property

Public Property Get intCOL��������ID() As Integer
    intCOL��������ID = mintCOL��������ID
End Property

Public Property Let intCOL��ֹʱ��(ByVal intData As Integer)
    mintCOL��ֹʱ�� = intData
End Property

Public Property Get intCOL��ֹʱ��() As Integer
    intCOL��ֹʱ�� = mintCOL��ֹʱ��
End Property

Public Property Let intCOLƵ��(ByVal intData As Integer)
    mintCOLƵ�� = intData
End Property

Public Property Get intCOLƵ��() As Integer
    intCOLƵ�� = mintCOLƵ��
End Property

Public Property Let intCOLƵ�ʴ���(ByVal intData As Integer)
    mintCOLƵ�ʴ��� = intData
End Property

Public Property Get intCOLƵ�ʴ���() As Integer
    intCOLƵ�ʴ��� = mintCOLƵ�ʴ���
End Property

Public Property Let intCOLƵ�ʼ��(ByVal intData As Integer)
    mintCOLƵ�ʼ�� = intData
End Property

Public Property Get intCOLƵ�ʼ��() As Integer
    intCOLƵ�ʼ�� = mintCOLƵ�ʼ��
End Property

Public Property Let intCOL�����λ(ByVal intData As Integer)
    mintCOL�����λ = intData
End Property

Public Property Get intCOL�����λ() As Integer
    intCOL�����λ = mintCOL�����λ
End Property

Public Property Let intCOL��ʾ(ByVal intData As Integer)
    mintCOL��ʾ = intData
End Property

Public Property Get intCOL��ʾ() As Integer
    intCOL��ʾ = mintCOL��ʾ
End Property

Public Property Let intCOL���(ByVal intData As Integer)
    mintCOL��� = intData
End Property

Public Property Get intCOL���() As Integer
    intCOL��� = mintCOL���
End Property

Public Property Let intCOL״̬(ByVal intData As Integer)
    mintCOL״̬ = intData
End Property

Public Property Get intCOL״̬() As Integer
    intCOL״̬ = mintCOL״̬
End Property

Public Property Let intCOLEDIT(ByVal intData As Integer)
    mintCOLEDIT = intData
End Property

Public Property Get intCOLEDIT() As Integer
    intCOLEDIT = mintCOLEDIT
End Property

Public Property Let intCOL����ID(ByVal intData As Integer)
    mintCOL����ID = intData
End Property

Public Property Get intCOL����ID() As Integer
    intCOL����ID = mintCOL����ID
End Property

Public Property Let intCOL��ҳID(ByVal intData As Integer)
    mintCOL��ҳID = intData
End Property

Public Property Get intCOL��ҳID() As Integer
    intCOL��ҳID = mintCOL��ҳID
End Property

Public Property Let intCOLѡ��(ByVal intData As Integer)
    mintCOLѡ�� = intData
End Property

Public Property Get intCOLѡ��() As Integer
    intCOLѡ�� = mintCOLѡ��
End Property

Public Property Let intCOLִ������(ByVal intData As Integer)
    mintCOLִ������ = intData
End Property

Public Property Get intCOLִ������() As Integer
    intCOLִ������ = mintCOLִ������
End Property

'������ҩ����
Public Property Let intCOLҩƷ����(ByVal intData As Integer)
    mintCOLҩƷ���� = intData
End Property

Public Property Get intCOLҩƷ����() As Integer
    intCOLҩƷ���� = mintCOLҩƷ����
End Property
'ΪҩƷģ���ṩ
Public Property Let intCOL����(ByVal intData As Integer)
    mintCOL���� = intData
End Property

Public Property Get intCOL����() As Integer
    intCOL���� = mintCOL����
End Property

Public Property Let intCOLNO(ByVal intData As Integer)
    mintCOLNO = intData
End Property

Public Property Get intCOLNO() As Integer
    intCOLNO = mintCOLNO
End Property

Public Property Let intCOLסԺ��(ByVal intData As Integer)
    mintCOLסԺ�� = intData
End Property

Public Property Get intCOLסԺ��() As Integer
    intCOLסԺ�� = mintCOLסԺ��
End Property

Public Property Let intCOL�����(ByVal intData As Integer)
    mintCOL����� = intData
End Property

Public Property Get intCOL�����() As Integer
    intCOL����� = mintCOL�����
End Property

Public Property Let intCOL�����(ByVal intData As Integer)
    mintCOL����� = intData
End Property

Public Property Get intCOL�����() As Integer
    intCOL����� = mintCOL�����
End Property

Public Property Let intCOL����(ByVal intData As Integer)
    mintCOL���� = intData
End Property

Public Property Get intCOL����() As Integer
    intCOL���� = mintCOL����
End Property

Public Property Let intcol������λ(ByVal intData As Integer)
    mintCol������λ = intData
End Property

Public Property Get intcol������λ() As Integer
    intcol������λ = mintCol������λ
End Property

Public Property Let intcolҽ������(ByVal intData As Integer)
    mintColҽ������ = intData
End Property

Public Property Get intcolҽ������() As Integer
    intcolҽ������ = mintColҽ������
End Property

Public Property Let intcol��ҩĿ��(ByVal intData As Integer)
    mintCol��ҩĿ�� = intData
End Property

Public Property Get intcol��ҩĿ��() As Integer
    intcol��ҩĿ�� = mintCol��ҩĿ��
End Property

Public Property Let intcol��ҩ����(ByVal intData As Integer)
    mintCol��ҩ���� = intData
End Property

Public Property Get intcol��ҩ����() As Integer
    intcol��ҩ���� = mintCol��ҩ����
End Property

'��������
Public Property Let intCol��������(ByVal intData As Integer)
    mintCol�������� = intData
End Property

Public Property Get intCol��������() As Integer
    intCol�������� = mintCol��������
End Property

Public Property Let intCol����ҩƷ˵��(ByVal intData As Integer)
    mintCol����ҩƷ˵�� = intData
End Property

Public Property Get intCol����ҩƷ˵��() As Integer
    intCol����ҩƷ˵�� = mintCol����ҩƷ˵��
End Property

Public Property Let intCol��־(ByVal intData As Integer)
    mintCol��־ = intData
End Property

Public Property Get intCol��־() As Integer
    intCol��־ = mintCol��־
End Property

Public Property Let intCol������(ByVal intData As Integer)
    mintCol������ = intData
End Property

Public Property Get intCol������() As Integer
    intCol������ = mintCol������
End Property

Public Property Let intColִ�з���(ByVal intData As Integer)
    mintColִ�з��� = intData
End Property

Public Property Get intColִ�з���() As Integer
    intColִ�з��� = mintColִ�з���
End Property

Public Property Let intColִ�п���ID(ByVal intData As Integer)
    mintColִ�п���ID = intData
End Property

Public Property Get intColִ�п���ID() As Integer
    intColִ�п���ID = mintColִ�п���ID
End Property