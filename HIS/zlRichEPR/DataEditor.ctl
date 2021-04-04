VERSION 5.00
Begin VB.UserControl DataEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   Picture         =   "DataEditor.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "DataEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#####################################################################################
'##     ����Ҫ�ر༭��
'#####################################################################################

Option Explicit

'���롢��������Ӣ���������͡����ȡ�С������λ���Ա�����ֵ�������򡢳�ʼֵ����ֵ����
Public Enum DataTypeEnum
    dte�ı� = 0
    dte���� = 1
    dte���� = 2
    dte��ѡ = 3
    dte��ѡ = 4
    dteָ�� = 5
End Enum

'0-���ޣ�1-�У�2-Ů����ʾ����Ŀ�ʺϵĻ����Ա�
Public Enum SexLimitEnum
    sle���� = 0
    sle�� = 1
    sleŮ = 2
End Enum

Private mvar������ID As Long
Private mvar���� As String
Private mvar������ As String
Private mvarӢ���� As String
Private mvar�滻�� As Long
Private mvar���� As DataTypeEnum
Private mvar���� As Long
Private mvarС�� As Long
Private mvar��λ As String
Private mvar�ٴ����� As String
Private mvar��ʾ�� As Long
Private mvar�Ա��� As SexLimitEnum
Private mvar��ֵ�� As String
Private mvar������ As String
Private mvar��ʼֵ As String
Private mvar��ֵ���� As String
Private mvarID As Long
Private mvar����ID As Long
Private mfrmDataEditor As New frmDataEditor
Private mvarWidth As Long
Private mvarHeight As Long
Public lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lID As Long, bBeteenKeys As Boolean, bNeeded As Boolean


Public Property Let Width(ByVal vData As Long)
    mvarWidth = vData
    PropertyChanged "Width"
End Property

Public Property Get Width() As Long
    Width = mvarWidth
End Property

Public Property Let Height(ByVal vData As Long)
    mvarHeight = vData
    PropertyChanged "Height"
End Property

Public Property Get Height() As Long
    Height = mvarHeight
End Property

Public Property Let ����ID(ByVal vData As Long)
    mvar����ID = vData
    PropertyChanged "����ID"
End Property

Public Property Get ����ID() As Long
    ����ID = mvar����ID
End Property

Public Property Let ID(ByVal vData As Long)
    mvarID = vData
    PropertyChanged "ID"
End Property

Public Property Get ID() As Long
    ID = mvarID
End Property

Public Property Let ��ֵ����(ByVal vData As String)
    mvar��ֵ���� = vData
    PropertyChanged "��ֵ����"
End Property

Public Property Get ��ֵ����() As String
    ��ֵ���� = mvar��ֵ����
End Property

Public Property Let ��ʼֵ(ByVal vData As String)
    mvar��ʼֵ = vData
    PropertyChanged "��ʼֵ"
End Property

Public Property Get ��ʼֵ() As String
    ��ʼֵ = mvar��ʼֵ
End Property

Public Property Let ������(ByVal vData As String)
    mvar������ = vData
    PropertyChanged "������"
End Property

Public Property Get ������() As String
    ������ = mvar������
End Property

Public Property Let ��ֵ��(ByVal vData As String)
    mvar��ֵ�� = vData
    PropertyChanged "��ֵ��"
End Property

Public Property Get ��ֵ��() As String
    ��ֵ�� = mvar��ֵ��
End Property

Public Property Let �Ա���(ByVal vData As SexLimitEnum)
    mvar�Ա��� = vData
    PropertyChanged "�Ա���"
End Property

Public Property Get �Ա���() As SexLimitEnum
    �Ա��� = mvar�Ա���
End Property

Public Property Let ��ʾ��(ByVal vData As Long)
    mvar��ʾ�� = vData
    PropertyChanged "��ʾ��"
End Property

Public Property Get ��ʾ��() As Long
    ��ʾ�� = mvar��ʾ��
End Property

Public Property Let �ٴ�����(ByVal vData As String)
    mvar�ٴ����� = vData
    PropertyChanged "�ٴ�����"
End Property

Public Property Get �ٴ�����() As String
    �ٴ����� = mvar�ٴ�����
End Property

Public Property Let ��λ(ByVal vData As String)
    mvar��λ = vData
    PropertyChanged "��λ"
End Property

Public Property Get ��λ() As String
    ��λ = mvar��λ
End Property

Public Property Let С��(ByVal vData As Long)
    mvarС�� = vData
    PropertyChanged "С��"
End Property

Public Property Get С��() As Long
    С�� = mvarС��
End Property

Public Property Let ����(ByVal vData As Long)
    mvar���� = vData
    PropertyChanged "����"
End Property

Public Property Get ����() As Long
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As DataTypeEnum)
    mvar���� = vData
    PropertyChanged "����"
End Property

Public Property Get ����() As DataTypeEnum
    ���� = mvar����
End Property

Public Property Let �滻��(ByVal vData As Long)
    mvar�滻�� = vData
    PropertyChanged "�滻��"
End Property

Public Property Get �滻��() As Long
    �滻�� = mvar�滻��
End Property

Public Property Let Ӣ����(ByVal vData As String)
    mvarӢ���� = vData
    PropertyChanged "Ӣ����"
End Property

Public Property Get Ӣ����() As String
    Ӣ���� = mvarӢ����
End Property

Public Property Let ������(ByVal vData As String)
    mvar������ = vData
    PropertyChanged "������"
End Property

Public Property Get ������() As String
    ������ = mvar������
End Property

Public Property Let ����(ByVal vData As String)
    mvar���� = vData
    PropertyChanged "����"
End Property

Public Property Get ����() As String
    ���� = mvar����
End Property

Public Property Let ������ID(ByVal vData As Long)
    mvar������ID = vData
    PropertyChanged "������ID"
End Property

Public Property Get ������ID() As Long
    ������ID = mvar������ID
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ������ID = PropBag.ReadProperty("������ID", 0)
    ���� = PropBag.ReadProperty("����", "")
    ������ = PropBag.ReadProperty("������", "")
    Ӣ���� = PropBag.ReadProperty("Ӣ����", "")
    �滻�� = PropBag.ReadProperty("�滻��", 0)
    ���� = PropBag.ReadProperty("����", 0)
    ���� = PropBag.ReadProperty("����", 0)
    С�� = PropBag.ReadProperty("С��", 0)
    ��λ = PropBag.ReadProperty("��λ", "")
    �ٴ����� = PropBag.ReadProperty("�ٴ�����", "")
    ��ʾ�� = PropBag.ReadProperty("��ʾ��", 0)
    �Ա��� = PropBag.ReadProperty("�Ա���", 0)
    ��ֵ�� = PropBag.ReadProperty("��ֵ��", "")
    ������ = PropBag.ReadProperty("������", "")
    ��ʼֵ = PropBag.ReadProperty("��ʼֵ", "")
    ��ֵ���� = PropBag.ReadProperty("��ֵ����", "")
    ID = PropBag.ReadProperty("ID", 0)
    ����ID = PropBag.ReadProperty("����ID", 0)
End Sub

Private Sub UserControl_Resize()
    Width = 500
    Height = 480
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "������ID", ������ID, 0
    PropBag.WriteProperty "����", ����, ""
    PropBag.WriteProperty "������", ������, ""
    PropBag.WriteProperty "Ӣ����", Ӣ����, ""
    PropBag.WriteProperty "�滻��", �滻��, 0
    PropBag.WriteProperty "����", ����, 0
    PropBag.WriteProperty "����", ����, 0
    PropBag.WriteProperty "С��", С��, 0
    PropBag.WriteProperty "��λ", ��λ, ""
    PropBag.WriteProperty "�ٴ�����", �ٴ�����, ""
    PropBag.WriteProperty "��ʾ��", ��ʾ��, 0
    PropBag.WriteProperty "�Ա���", �Ա���, 0
    PropBag.WriteProperty "��ֵ��", ��ֵ��, ""
    PropBag.WriteProperty "������", ������, ""
    PropBag.WriteProperty "��ʼֵ", ��ʼֵ, ""
    PropBag.WriteProperty "��ֵ����", ��ֵ����, ""
    PropBag.WriteProperty "ID", ID, 0
    PropBag.WriteProperty "����ID", ����ID, 0
    
    PropertyChanged "������ID"
    PropertyChanged "����"
    PropertyChanged "������"
    PropertyChanged "Ӣ����"
    PropertyChanged "�滻��"
    PropertyChanged "����"
    PropertyChanged "����"
    PropertyChanged "С��"
    PropertyChanged "��λ"
    PropertyChanged "�ٴ�����"
    PropertyChanged "��ʾ��"
    PropertyChanged "�Ա���"
    PropertyChanged "��ֵ��"
    PropertyChanged "������"
    PropertyChanged "��ʼֵ"
    PropertyChanged "��ֵ����"
    PropertyChanged "ID"
    PropertyChanged "����ID"
End Sub

Public Sub ShowEditor(x As Long, y As Long, lWidth As Long, lHeight As Long, eType As DataTypeEnum)
    With mfrmDataEditor
        .lKSS = lKSS
        .lKSE = lKSE
        .lKES = lKES
        .lKEE = lKEE
        .lID = lID
        .bNeeded = bNeeded
        .ShowDataEditor x, y, lWidth, lHeight, UserControl.Parent, eType, mvar��ֵ��, mvar��ʼֵ, mvar������, _
        mvarӢ����, mvar����, mvarС��, mvar��λ, mvar�Ա���, mvar������, mvar��ֵ����
    End With
End Sub






























