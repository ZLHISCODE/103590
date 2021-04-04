VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mfrmCom As frmCom
Public Function OpenPassKeyoardInput(ByVal frmMain As Object, _
    ByVal objPassCtl As Object, Optional blnAffirmPass As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����뽨������
    '���:frmMain-���õ�������
    '       objPassCtl-���������ؼ�
    '       blnAffirmPass-False:����������;true:������ȷ������
    '����:
    '����:�򿪳ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 23:30:54
    '--------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmCom Is Nothing Then Exit Function
    If Not gblnStartKeyboard Then OpenPassKeyoardInput = True: Exit Function
    On Error GoTo errHandle
    OpenPassKeyoardInput = mfrmCom.OpenPassKeyoardInput(frmMain, objPassCtl, blnAffirmPass)
    Exit Function
errHandle:
End Function

Public Function ColsePassKeyoardInput(ByVal frmMain As Object, ByVal objPassCtl As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ر����뽨������
    '���:frmMain-���õ�������
    '       objPassCtl-���������ؼ�
    '����:
    '����:�رճɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 23:30:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmCom Is Nothing Then Exit Function
    If Not gblnStartKeyboard Then ColsePassKeyoardInput = True: Exit Function
    On Error GoTo errHandle
    ColsePassKeyoardInput = mfrmCom.ColsePassKeyoardInput(frmMain, objPassCtl)
    Exit Function
errHandle:
End Function
Public Function zlCardDevSet(ByVal frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:zlCardDevSet(�豸�������ýӿ�):��Ҫ��������ص��豸����ز���
    '���:frmMain Form    HIS���� ���õ�������
    '����:
    '����:true:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:18:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    frmDeviceSet.Show 1, frmMain
End Function

Private Sub Class_Initialize()
    Err = 0: On Error Resume Next
    gblnStartKeyboard = Val(GetSetting("ZLSOFT", "����ģ��\zlKeyboard", "����", "0")) = 1
    Set mfrmCom = New frmCom
    Load mfrmCom
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mfrmCom Is Nothing Then Unload mfrmCom
    Set mfrmCom = Nothing
End Sub