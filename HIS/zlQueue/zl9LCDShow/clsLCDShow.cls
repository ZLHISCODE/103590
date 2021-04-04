VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLCDShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mfrmLCDShow As frmLCDShow
Dim mfrmDrugShow As frmDrugShow

Public Function zlShow(cnOracle As ADODB.Connection, str��������() As String, _
    Optional str���� As String = "", Optional strҽ�� As String = "", _
    Optional strExcludeData As String = "", Optional intViewDataType As Integer = 0, _
    Optional blnComeBackFirst As Boolean = True) As String
    
    
    '�ڵ�����Ļ����ʾ�Ŷ����
    '����ֵΪ�Ŷ�����ַ���
    
    On Error GoTo err
    
    '�ж��Ƿ���Ҫ������ʾ����
    If mfrmLCDShow Is Nothing Then
        '������ʾ����
        Set mfrmLCDShow = New frmLCDShow
        
        Call mfrmLCDShow.SetFaceFont
        Call mfrmLCDShow.SetFacePostion
        mfrmLCDShow.zlShowMe cnOracle, str��������, str����, strҽ��, strExcludeData, intViewDataType, blnComeBackFirst
    Else
        Call mfrmLCDShow.zlSetPara(str��������, str����, strҽ��, strExcludeData, blnComeBackFirst)
    End If
    
    
    
    '�������ص��Ŷ�����ַ���
    zlShow = "" 'mfrmLCDShow.lblShowText.Caption
    
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlDrugShow(ByVal lngStoreID As Long, ByVal strWin As String, ByVal bln��ҩ As Boolean, ByVal bln��ҩȷ�� As Boolean, Optional ByVal strName As String) As String
    'ҩ���Ŷӽк���ʾ���
    '����ֵΪ��ǰ��������
    Dim intMode As Integer '����ģʽ��0-�����ڣ�1-�ര��
    Dim strTemp As String
    Dim strReg As String
    
    On Error GoTo err
    
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    
    intMode = Val(GetSetting("ZLSOFT", "����ģ��\ҩ���Ŷӽк�\Һ������", "����ģʽ", "0"))
    
    If intMode < 0 Or intMode > 1 Then
        intMode = 0
    End If
    
    If intMode = 0 Then
        strTemp = strWin
    End If
    
    If strName = "" Then Set mfrmDrugShow = Nothing
    If mfrmDrugShow Is Nothing Then
        '������ʾ����
        Set mfrmDrugShow = New frmDrugShow
        
        '���´�������
        
        '��ʾ����
        mfrmDrugShow.ShowMe lngStoreID, strTemp, bln��ҩ, bln��ҩȷ��
    Else
        '�����Ѵ���ʱ�����µ�ǰ������
        mfrmDrugShow.ChangeCall strWin, strName
    End If
        
    
    '���ص�ǰ�ĺ�������
    zlDrugShow = ""
    
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlClose() As Boolean
    '�ر�Һ�����ӵ���ʾ����
    
    On Error GoTo err
    
    If mfrmLCDShow Is Nothing Then Exit Function
    
    Unload mfrmLCDShow
    Set mfrmLCDShow = Nothing
    
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlDrugShowClose() As Boolean
    '�ر�Һ�����ӵ���ʾ����
    
    On Error GoTo err
    
    If mfrmDrugShow Is Nothing Then Exit Function
    
    Unload mfrmDrugShow
    Set mfrmDrugShow = Nothing
    
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function zlSetup(frmParent As Object) As Boolean
    '����Һ��������ز���
    
    '�������ã�ʹ��ģʽ��������ʾ
    zlSetup = frmLCDSetup.zlShowMe(frmParent)
    
    If mfrmLCDShow Is Nothing Then Exit Function
    
    Call mfrmLCDShow.SetFaceFont
    Call mfrmLCDShow.SetFacePostion
    
End Function

Public Function zlDrugSetup(frmParent As Object, ByVal strSendWindows As String) As Boolean
    '����ҩ���Ŷӽк���ʾ����豸����ز���
    
    '�������ã�ʹ��ģʽ��������ʾ
    zlDrugSetup = frmDrugSetup.ShowMe(strSendWindows, frmParent)
    
    If mfrmDrugShow Is Nothing Then Exit Function
    
    '���´�������
    Call mfrmDrugShow.SetFacePostion
    Call mfrmDrugShow.SetFont
    
End Function