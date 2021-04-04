VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZHPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String
Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    '��ʼ����������
    If gobjComLib Is Nothing Then
        Set gobjComLib = GetObject("", "zl9Comlib.clsComlib")
        Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
        Set gobjControl = GetObject("", "zl9Comlib.clsControl")
        Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    End If
End Sub
Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set gobjComLib = Nothing
    Set gobjCommFun = Nothing
    Set gobjControl = Nothing
    Set gobjDatabase = Nothing
    Err = 0
End Sub
'*************************************************************************************************************************************
'��HIS����ؽӿ�
'��������ֵ�ľֲ�����
'�ӿ���
Public Function zlInitCompoent(ByVal intCallType As Integer, cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:zlInitCompoent (��ʼ���ӿڲ���)
    '���:intCallType Byte    HIS���� 0-  ������õ���  1-  סԺ���ʵ���
    '     cnOracle    ADODB.Connection    HIS���� �������Ӷ���
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlInitCompoent = True
    Exit Function
End Function
Public Function zlCardDevSet(ByVal frmMain As Object, Optional lngCardCode As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:zlCardDevSet(�豸�������ýӿ�):��Ҫ��������ص��豸����ز���
    '���:frmMain Form    HIS���� ���õ�������
    '����:
    '����:true:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:18:38
    'HIS����˵��.
    '    1.  �������շѹ����Ĳ�������'�豸����' zlCardDevSet
    '    2.  ��סԺ���ʽ���Ĳ�������'�豸����' zlCardDevSet
    '---------------------------------------------------------------------------------------------------------------------------------------------
    frmSetDrvParaSet.ShowMe frmMain, lngCardCode
End Function
 
Public Function zlBrushCardSquare(ByVal intCallType As Integer, frmMain As Object, ByVal lng�ӿڱ�� As Long, ByVal rsFeeList As ADODB.Recordset, ByVal dbl������� As Double, ByRef rsSquare As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: zlBrushCardSquare (ˢ������ӿ�)
    '���:frmMain:HIS���� ���õ�������
    '     intCallType : HIS���� 0-  ������õ��� 1-  סԺ���ʵ���
    '     rsFeeList: HIS���� ���������൥��,�����е��ݵ���ϸ,�����סԺ���� , ���Ǳ��ν��ʵ�������ϸ
    '     dbl������� :  HIS���� ��ʾˢ�����ܳ����˽��
    '
    '����:rsSquare : �ӿڷ���    ���ؼ�¼��:�ӿڴ���սṹ(�ӿڷ�����ص�����) , �ṹ����:
    '                �ӿڱ�� , ���ѿ�ID, ���㷽ʽ, ������, ���ſ�����, ������ˮ��, ����ʱ��, ��ע
    '     rsSquare˵��:��Ҫ�ǽ��ͬһ����,ˢ���ſ����ѵ����.,�������ˢ���ſ� , ����ӿ����Ѿ�ˢ���Ŀ���Ϣ
    '����:true:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:18:38
    '˵��:
    '    1.  �������շѽ���ʱ,HIS�ڵ�"���㿨"ʱ,���ñ��ӿ�
    '    2.  ��סԺ���ʽ���ʱ,HIS�ڵ�"���㿨"ʱ,���ñ��ӿ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
        zlBrushCardSquare = True
End Function

Public Function zlSquareFee(ByVal intCallType As Integer, frmMain As Object, ByVal str����ID_IN As String, ByVal rsSquare As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: zlSquareFee (����ӿ�)
    '���:frmMain:HIS���� ���õ�������
    '     intCallType : HIS���� 0-  ������õ��� 1-  סԺ���ʵ���
    '     str����ID_IN: HIS���� ���ν��ʵĽ���ID��
    '     rsSquare :  ����Ӧˢ���Ľ���
    '����:
    '����:true:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:18:38
    '˵��:
    '    1. ��"�����շ�"�����"ȷ��"ʱ,���ñ��ӿ�
    '    2. ��"סԺ����"�����"ȷ��"ʱ,���ñ��ӿ�
    'ע:
    '  �˽ӿ���������HIS������ , ��˲����ڴ˽ӿڴ������û������Ĳ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'ֻ���־
    zlSquareFee = True
End Function
Public Function zlCheckDelSquareValied(frmMain As Object, ByVal intCallType As Integer, ByVal str����ID_IN As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ִ���˷�ʱ,�����صĽӿڲ����Ƿ�����
    '���:
    '����:
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-31 16:39:47
    '˵��;
    '     ���˷�ʱ����Ҫ������صļ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    zlCheckDelSquareValied = True
End Function
Public Function zlDelSquareFee(ByVal intCallType As Integer, frmMain As Object, ByVal str����ID_IN As String) As Boolean
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: zlSquareFee (����ӿ�)
    '���:frmMain:HIS���� ���õ�������
    '     intCallType : HIS���� 0-  ������õ��� 1-  סԺ���ʵ���
    '     str����ID_IN: HIS���� ���ν��ʵĽ���ID��
    '����:
    '����:true:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:18:38
    '˵��:
    '    1. "�����շѹ���"��"סԺ���ʹ���"������ʱ,���ô˽ӿ�
    'ע:
    '  �˽ӿ���������HIS������ , ��˲����ڴ˽ӿڴ������û������Ĳ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlDelSquareFee = True
End Function
Public Function zlSimulationBrushCard(ByVal frmMain As Object, ByVal Dblˢ����� As Double, _
    ByRef lng�ӿڱ�� As Long, ByRef strBlanceInfor As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ѡ��ָ��������
    '��Σ�frmMain HIS���� ���õ�������
    '         Dblˢ����� HIS���� ����Ԥ�������еĽ��
    '         Lng�ӿڱ��          HIS������
    '���Σ�Lng�ӿڱ�� ����    �Ժ��ֽ��㿨����
    '         strBlanceInfor  ����    ��||�ָ�: �ӿڱ��||���ѿ�ID(�ɴ�'')||���㷽ʽ||������||����||������ˮ��||����ʱ��(yyyy-mm-dd hh24:mi:ss)||��ע
    '          strBlanceInfor-���뷵��ֵ,����ᱨ��.
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-06-18 11:33:22
    '˵������Ԥ�������������Ԥ��ʱ�����ȷ����ťʱ����(����ǰ����)
    '------------------------------------------------------------------------------------------------------------------------
    strBlanceInfor = lng�ӿڱ�� & "|| ||POS����||100||432143||222||" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "||��ע"
    zlSimulationBrushCard = True
End Function

Public Function zlBrushCardToDeposit(ByVal lngԤ��ID As Long, ByVal lng�ӿڱ�� As Long, ByRef strBlanceInfor As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ����Ԥ������
    '��Σ� lngԤ��ID-Ԥ��ID
    '           lng�ӿڱ��-���㿨���
    '���Σ�strBlanceInfor-����ˢ����Ϣ:
    '         ��||�ָ�: �ӿڱ��||���ѿ�ID(�ɴ�'')||���㷽ʽ||������||����||������ˮ��||����ʱ��(yyyy-mm-dd hh24:mi:ss)||��ע
    '���أ��ɹ�����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-18 11:33:22
    '˵������Ԥ�������������Ԥ��ʱ�����ȷ����ťʱ����(�����е���)
    '          ����һ��Ҫ������ȷ,�������ֳ������
    '------------------------------------------------------------------------------------------------------------------------
    zlBrushCardToDeposit = True
End Function
Public Function zlCheckDepositDelValied(ByVal lngԤ��ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����˷ѽ��׽ӿ�
    '��Σ� lngԤ��ID-Ԥ��ID
    '���أ��ɹ�����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-18 11:33:22
    '˵����1. ��Ԥ��������е��˿�,��ȷ��ʱ(����֮ǰ����)�����øý���
     '------------------------------------------------------------------------------------------------------------------------
    zlCheckDepositDelValied = True
End Function
Public Function zlDepositDel(ByVal lngԤ��ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���Ԥ������
    '��Σ� lngԤ��ID-Ԥ��ID
    '���أ��ɹ�����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-18 11:33:22
    '˵����1. ��Ԥ��������е��˿�, ��ȷ��ʱ(�����е���)�����øý���
     '------------------------------------------------------------------------------------------------------------------------
    zlDepositDel = True
End Function





