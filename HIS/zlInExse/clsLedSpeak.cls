VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLedSpeak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'LED����
Private mblnLED As Boolean        '����ʱ�Ƿ�����LED�豸����
Private mblnLedWelcome As Boolean '�Ƿ������겡�˺���ʾ��ӭ��Ϣ
Private mlngModul As Long
Private mlngSys As Long
Private mobjCom As MSComm
Private mfrmMain As Object
Private mdbl�ɿ� As Double, mdbl���� As String, mdbl�����ܶ� As Double  '��Ҫ��¼��Ҫ���ŵ����ݣ�����ֱ��ʹ��zlSpeak_Money


Public Property Let blnLED(ByVal vData As Boolean)
    mblnLED = vData
End Property
Public Property Get blnLED() As Boolean
    blnLED = mblnLED
End Property

Public Property Let �ɿ�(ByVal vData As Double)
    mdbl�ɿ� = vData
End Property
Public Property Get �ɿ�() As Double
    �ɿ� = mdbl�ɿ�
End Property
Public Property Let ����(ByVal vData As Double)
    mdbl���� = vData
End Property
Public Property Get ����() As Double
    ���� = mdbl����
End Property
Public Property Let �����ܶ�(ByVal vData As Double)
    mdbl�����ܶ� = vData
End Property
Public Property Get �����ܶ�() As Double
    �����ܶ� = mdbl�����ܶ�
End Property

Public Sub zlInit(ByVal mfrmMain As Object, ByVal lngModul As Long, ByVal lngSys As Long, ByVal objCom As Object, Optional blnNotStartLed As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '���:mfrmMain-���õ�������
    '     lngModul-ģ���
    '     objCom-MsComm�ؼ�
    '     lngSys-ϵͳ��
    '     blnNotStartLed-true-��ʾ������Led,False-��ʾ���������������Ƿ�����һ��ͨ
    '����:���˺�
    '����:2018-09-26 14:24:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mlngModul = lngModul: mlngSys = lngSys: Set mobjCom = objCom
    mblnLED = False: mblnLedWelcome = False
    If Not blnNotStartLed Then  '��ʾ���������������Ƿ�����һ��ͨ
        If lngModul = 1133 Or lngModul = 1134 Or lngModul = 1135 Or lngModul = 1137 Then
            mblnLED = Val(GetSetting("ZLSOFT", "����ȫ��", "ʹ��", 0)) <> 0
        End If
        If lngModul = 1133 Or lngModul = 1134 Or lngModul = 1135 Or lngModul = 1137 Then
            mblnLedWelcome = zlDatabase.GetPara("LED��ʾ��ӭ��Ϣ", mlngSys, mlngModul, "1") = "1"
        End If
    End If
    Call zlInitLed(objCom)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlInitLed(ByVal objCom As Object, Optional blnReSetCom As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��Led
    '���:objCom-MsComm�ؼ�
    '     blnReSetCom-�Ƿ���������Com�˿�
    '����:���˺�
    '����:2018-09-26 14:55:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mblnLED Then Exit Sub
    
    If blnReSetCom Then zl9LedVoice.Reset objCom
    zl9LedVoice.Init UserInfo.��� & "��Ϊ������", mlngModul, gcnOracle
End Sub
Public Sub zlLedShowWelcome(Optional objPati As clsPatientInfo, Optional blnReSetCom As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ӭ��Ϣ
    '���:blnReSetCom-�Ƿ���������Com�˿�
    '����:���˺�
    '����:2015-01-08 10:54:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If mblnLED = False Then Exit Sub
    
    If mblnLedWelcome Then
        Call zlInitLed(mobjCom, blnReSetCom)
        zl9LedVoice.Speak "#1"
    End If
    Call zlDisplayPatient(objPati)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub zlSpeak_�븶��(ByVal dblMoney As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������㸶��
    '���:dblMoney-���ŵ�ǰ����Ҫ֧���Ľ��
    '����:���˺�
    '����:2018-09-26 16:15:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnLED = False Then Exit Sub
    zl9LedVoice.Speak "#21 " & Format(dblMoney, "0.00")  '��������XXXXԪ
End Sub

Public Sub zlDispCharge(ByVal dbl�ܶ� As Double, ByVal dbl�ɿ� As Double, ByVal dbl���� As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾӦ��,ʵ��,����,������
    '����:���˺�
    '����:2018-09-26 16:36:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
   If mblnLED = False Then Exit Sub
    '��ʾӦ��,ʵ��,����,������
    zl9LedVoice.DispCharge Format(dbl�ܶ�, "0.00"), Format(dbl�ɿ�, "0.00"), Format(dbl����, "0.00")
    
End Sub
Public Sub zlSpeak_PayMoney(ByVal dbl�����ܶ� As Double, ByVal dbl�ɿ� As Double, ByVal dbl���� As Double, Optional blnPaySpeakMoney As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������㸶��(�˹��̻�ı�ɿ���㡢�����ܶ������ֵ )
    '���:dbl�ɿ�-����֧���ɿ���
    '     dbl����-����
    '     blnPaySpeakMoney-true:����Ԥ�ա����㼰�������;False-������
    '����:���˺�
    '����:2018-09-26 16:15:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnLED = False Then Exit Sub
    '��ʾӦ��,ʵ��,����,������
    mdbl�ɿ� = dbl�ɿ�: mdbl�����ܶ� = dbl�����ܶ�: mdbl���� = dbl����
    Call zlDispCharge(dbl�����ܶ�, dbl�ɿ�, dbl����)
    
    If Not blnPaySpeakMoney Then Exit Sub
    '����Ԥ�ա����㼰�������
    Call zlSpeak_Money
End Sub
Public Sub zlSpeak_Money()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ�ա����㼰������壨�����ȶԽɿ�������Ը�ֵ��
    '����:���˺�
    '����:2018-09-26 16:43:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnLED = False Then Exit Sub
    
    zl9LedVoice.Speak "#22 " & Format(mdbl�ɿ�, "0.00")  'Ԥ��XXXXԪ
    zl9LedVoice.Speak "#23 " & Format(mdbl����, "0.00")   '����XXXXԪ
    zl9LedVoice.Speak "#3"   '#3  --�뵱�����, лл!
End Sub
Public Sub zlSpeak_����(ByVal dbl���� As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ�ա����㼰������壨�����ȶԽɿ�������Ը�ֵ��
    '����:���˺�
    '����:2018-09-26 16:43:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnLED = False Then Exit Sub
    zl9LedVoice.Speak "#23 " & Format(dbl����, "0.00")  '����XXXXԪ
End Sub



Public Sub zlDisplayPatient(objPati As clsPatientInfo)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������Ϣ
    '���:objPati-������Ϣ��
    '����:���˺�
    '����:2018-09-26 14:50:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnLED = False Or objPati Is Nothing Then Exit Sub
    zl9LedVoice.DisplayPatient objPati.���� & " " & objPati.�Ա� & " " & objPati.����, objPati.����ID
End Sub
Public Sub zlClearDisplayPatient()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ʾ������Ϣ
    '���:objPati-������Ϣ��
    '����:���˺�
    '����:2018-09-26 14:50:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnLED = False Then Exit Sub
    zl9LedVoice.DisplayPatient ""
End Sub

Public Sub zlClearDisplayBank(Optional blnSpace As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ʾ��Ϣ
    '���:blnSpace �Ƿ���ո�
    '
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-09-26 17:17:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mblnLED Then Exit Sub
    zl9LedVoice.DisplayBank IIf(blnSpace, " ", "")
End Sub

Public Sub zlDisplayBank(ByVal cllBanks As Collection, Optional blnIsWait4 As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Led��Ϣ��ʾ��������������ʾ�����ַ���
    '���:cllBanks-��ʾ�ĸ�����,��ʽΪ:array(��Ŀ,���)
    '     blnIsWait4-�Ƿ�ȴ�4��
    '����:���˺�
    '����:2015-01-07 13:55:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtTmpDate As Date, varBank(1 To 20) As String
    Dim i As Long
    If Not mblnLED Or cllBanks Is Nothing Then Exit Sub
    
    On Error GoTo errHandle
    
    For i = 1 To cllBanks.Count
        If i <= 20 Then
            varBank(i) = cllBanks(i)(0) & cllBanks(i)(1) '��ʽ��:"�ʻ����" & Format(mobjInsureInfor.�ʻ����, "0.00")
        End If
    Next
    
    Select Case cllBanks.Count
    Case 1
        zl9LedVoice.DisplayBank varBank(1)
    Case 2
        zl9LedVoice.DisplayBank varBank(1), varBank(2)
    Case 3
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3)
    Case 4
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4)
    Case 5
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5)
    Case 6
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6)
    Case 7
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7)
    Case 8
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8)
    Case 9
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9)
    Case 10
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10)
    Case 11
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11)
    Case 12
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12)
    Case 13
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13)
    Case 14
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13), varBank(14)
    Case 15
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13), varBank(14), varBank(15)
    Case 16
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13), varBank(14), varBank(15), varBank(16)
    Case 17
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13), varBank(14), varBank(15), varBank(16), varBank(17)
    Case 18
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13), varBank(14), varBank(15), varBank(16), varBank(17), varBank(18)
    Case 19
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13), varBank(14), varBank(15), varBank(16), varBank(17), varBank(18), varBank(19)
    Case Else   '���20��
        zl9LedVoice.DisplayBank varBank(1), varBank(2), varBank(3), varBank(4), varBank(5), varBank(6), varBank(7), varBank(8), varBank(9), varBank(10), varBank(11), varBank(12), varBank(13), varBank(14), varBank(15), varBank(16), varBank(17), varBank(18), varBank(19), varBank(20)
    End Select
    If blnIsWait4 Then
        '�ӳ�ʱ��
        dtTmpDate = Time
        Do While Time < DateAdd("s", 4, dtTmpDate)  '�����10.34.0��ֻ��ҽ���ż�4�룬��10.35�����ܷ���ҲҪ��4�룬����û��ȴ�ʱ���������˻��ǻָ�ֻ��ҽ��������4�루ΪʲôҪ����,��ǰ����)
        Loop
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub



Private Sub Class_Initialize()
    Err = 0: On Error Resume Next

End Sub
Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    
    If mblnLED Then
        zl9LedVoice.DisplayPatient ""
        zl9LedVoice.Reset mobjCom
    End If
    Set mobjCom = Nothing
    Set mfrmMain = Nothing
End Sub