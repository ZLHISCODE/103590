VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCashPayEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event zlSendCash(dblTotal As Double, dblMony As Double, blnCancel As Boolean, strErrMsg As String)
Public Event zlSendCard(strCardNO As String, blnCancel As Boolean, strErrMsg As String)
'Public Event zlRetrunMoneyComplete(dblȱ������ As Double)

Public Sub zlRaiseSendCash(dblTotal As Double, dblMony As Double, blnCancel As Boolean, strErrMsg As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ͷ���ֽ��¼�
    '���:dblTotal-Ͷ���ܶ�
    '       dblMoney-��ǰͶ�Ŷ�
    '����:blnCancel-true,����ʧ��;false-�����ɹ�
    '       strErrMsg-���صĴ�����Ϣ
    '����:���˺�
    '����:2012-04-11 14:29:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    RaiseEvent zlSendCash(dblTotal, dblMony, blnCancel, strErrMsg)
End Sub
Public Sub zlRaiseSendCard(strCardNO As String, blnCancel As Boolean, strErrMsg As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������¼�
    '���:strCardNo-��ǰ�����Ŀ���
    '����:blnCancel-true,����ʧ��;false-�����ɹ�
    '       strErrMsg-���صĴ�����Ϣ
    '����:���˺�
    '����:2012-04-11 14:29:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    RaiseEvent zlSendCard(strCardNO, blnCancel, strErrMsg)
End Sub

'Public Sub zlRaiseRetrunMoneyComplete(dblȱ������ As Double)
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:��ȡȱ������
'    '���:bytȱ������
'    '����:����
'    '����:2012-11-19 14:29:02
'    '�����:56055
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    RaiseEvent zlRetrunMoneyComplete(dblȱ������)
'End Sub