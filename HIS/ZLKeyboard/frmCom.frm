VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Begin VB.Form frmCom 
   Caption         =   "frmCom"
   ClientHeight    =   855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   3090
   Begin MSCommLib.MSComm msCommKeyBoard 
      Left            =   210
      Top             =   165
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnInitCom As Boolean   '�Ƿ��ʼ��Com�ӿڳɹ�
Private WithEvents mTxtPass As TextBox
Attribute mTxtPass.VB_VarHelpID = -1

Private Sub Form_Load()
     Call InitComProperty
     mblnInitCom = InitComm
End Sub

Private Sub msCommKeyBoard_OnComm()
    Dim strKeyChar As String
    '��������
    If mTxtPass Is Nothing Then Exit Sub
    strKeyChar = msCommKeyBoard.Input
    If strKeyChar = "" Then Exit Sub
    '����ַ�
    If Asc(strKeyChar) = 8 Then mTxtPass.Text = "": Exit Sub
    If Asc(strKeyChar) >= Asc("0") And Asc(strKeyChar) <= Asc("9") Then
        mTxtPass.Text = mTxtPass.Text & strKeyChar
        mTxtPass.SelStart = Len(mTxtPass.Text)
    End If
    If Asc(strKeyChar) = 13 Then PressKey vbKeyReturn
End Sub
Private Sub mTxtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Then
        Exit Sub
    End If
    KeyAscii = 0
End Sub
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
    On Error GoTo errHandle
    If mblnInitCom = False Then Exit Function
    With msCommKeyBoard
        If .PortOpen = False Then .PortOpen = True
        If blnAffirmPass Then
            Call HexSend("81H")  '���,���ٴ���������
        Else
            Call HexSend("82H") '���,����������
        End If
    End With
    Set mTxtPass = objPassCtl
    OpenPassKeyoardInput = True
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
    '����:2011-07-28 16:07:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If mblnInitCom = False Then Exit Function
    With msCommKeyBoard
        If .PortOpen = True Then .PortOpen = False
    End With
    Set mTxtPass = Nothing
    ColsePassKeyoardInput = True
    Exit Function
errHandle:
End Function
Private Function InitComm() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�򿪶˿�
    '����:���˺�
    '����:2011-07-28 14:35:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSet As String
    On Error GoTo errHandle
     
     strSet = g_Com_Property.lng������
     strSet = strSet & "," & Switch(g_Com_Property.str��ż����λ = "��", "n", g_Com_Property.str��ż����λ = "��", "o", g_Com_Property.str��ż����λ = "ż", "e", g_Com_Property.str��ż����λ = "�ո�", " ", True, "n")
     strSet = strSet & "," & g_Com_Property.int����λ
     strSet = strSet & "," & g_Com_Property.intֹͣλ
     With msCommKeyBoard
        .CommPort = g_Com_Property.int�˿ں�
        .Settings = strSet
        .InputLen = 1      '���ؽ��ջ������еȴ����ַ���,�����������ʱ��Ч
        .RThreshold = 1 '���յ�1���ֽ����ݾ���������OnComm()�¼�
     End With
    InitComm = True
    Exit Function
errHandle:
End Function
Private Sub HexSend(ByVal strSend As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʮ�������������
    '����:���˺�
    '����:2011-07-28 15:25:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intOutPutLen As Integer    '�������ݵĳ���
    Dim strOutdata As String          '���������ݴ�
    Dim bytSendArr() As Byte       '��������
    Dim strTempSave As String '�����ݴ�
    Dim intCount As Integer
    Dim i As Integer
    Err = 0: On Error Resume Next
    strOutdata = UCase(Replace(strSend, " ", ""))         'ȥ���ո񣬱�ɴ�д��ĸ
    intOutPutLen = Len(strOutdata)            '���ݵĳ���
    For i = 0 To intOutPutLen
        strTempSave = Mid(strOutdata, i + 1, 1)          'ȡһλ����
        If (Asc(strTempSave) >= Asc("0") And Asc(strTempSave) <= Asc("9")) _
        Or (Asc(strTempSave) >= 65 And Asc(strTempSave) <= 70) Then
            intCount = intCount + 1
        Else
            Exit For
        End If
    Next
    If intCount Mod 2 <> 0 Then            '�ж�ʮ�����������Ƿ�Ϊ˫��
        intCount = intCount - 1           '����˫�����ȥ1
    End If
    strOutdata = Left(strOutdata, intCount)       'ȡ����Ч��ʮ����������
    ReDim bytSendArr(intCount / 2 - 1)        '���¶������鳤��
    For i = 0 To intCount / 2 - 1
        bytSendArr(i) = Val("&H" + Mid(strOutdata, i * 2 + 1, 2)) 'ȡ������ת����ʮ�����Ʋ���ŵ�������
    Next
     msCommKeyBoard.Output = bytSendArr          '��������
End Sub


