VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBrushSequareCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Private mobjCard  As Object
Private mstrInterFace As String '����ʱѡ��Ľӿ�����
Private mlng�ӿ���� As Long
Private mobjCurCard As clsCard
Public Event zlBrushCarding(ByVal strCardNo As String)

Public Function zlInitInterFacel(ByVal lng�ӿ���� As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����صĽӿڲ���
    '����:���˺�
    '����:2009-12-15 23:18:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As clsCard
    Dim i As Long
    mlng�ӿ���� = lng�ӿ����:
    
    Err = 0: On Error Resume Next
    '����豸�Ƿ�����
    Set mobjCard = gobjStartCards("K" & lng�ӿ����)(0)
    If Err <> 0 Then
            Err = 0: On Error Resume Next
            Set objCard = gObjXFCards.Item("K" & lng�ӿ����)
            If Err <> 0 Then
                ShowMsgbox "����:" & mlng�ӿ���� & "δ�ҵ�!"
                Call WritLog("clsBrushSequareCard.zlInitInterFacel", "", "����:" & mlng�ӿ���� & "δ�ҵ�!")
                Exit Function
            End If
            If Err <> 0 Then
                ShowMsgbox "����:" & mlng�ӿ���� & "-" & objCard.���� & "û����������!" & vbCrLf & "��ϸ����ϢΪ:" & Err.Description
                Call WritLog("clsBrushSequareCard.zlInitInterFacel", "", "����:" & mlng�ӿ���� & "-" & objCard.���� & "û����������!" & vbCrLf & "��ϸ����ϢΪ:" & Err.Description)
            End If
          Exit Function
    End If
    '��ʽ;array(��������,���ƿ�,�ӿڱ��)
    Err = 0: On Error Resume Next
    Set objCard = gObjXFCards.Item("K" & lng�ӿ����)
    If objCard.�Ƿ��Զ���ȡ = 1 Then gbln�Զ���ȡ = True
    Set mobjCurCard = objCard
    If Err <> 0 Then
        ShowMsgbox "����:" & mlng�ӿ���� & "δ�ҵ�!"
        Call WritLog("clsBrushSequareCard.zlInitInterFacel", "", "����:" & mlng�ӿ���� & "δ�ҵ�!")
        Exit Function
    End If
    
    If Val(gobjStartCards("K" & lng�ӿ����)(1)) = 1 Then
        Call frmCardBrush.Init(Me)
        frmCardBrush.tmrMain.Interval = Val(GetSetting("ZLSOFT", "����ȫ��\SquareCard\" & lng�ӿ����, "�Զ���ȡ���", 300))
        With gTy_TestBug
            Set .objSquareCard = objCard
        End With
    End If
    zlInitInterFacel = True
End Function

'------------------------------------------------------------------
'��ص�ˢ������
Friend Sub zlBrushCarding(ByVal strCardNo As String)
    'ˢ��ʱ
    RaiseEvent zlBrushCarding(strCardNo)
End Sub
Public Sub zlSetAutoBrush(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Զ�ˢ��
    '����:���˺�
    '����:2009-12-15 17:38:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gbln�Զ���ȡ = True Then
        frmCardBrush.tmrMain.Enabled = blnEnabled
    End If
End Sub

'------------------------------------------------------------------
'��صĽӿڲ���
'��������ֵ�ľֲ�����
'�ӿ���
Public Function zlInitComponents(ByVal intCallType As Integer, cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '���:intCallType Byte    HIS���� 0-  ������õ���  1-  סԺ���ʵ���
    '     cnOracle    ADODB.Connection    HIS���� �������Ӷ���
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.���ѿ�ʱ.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set gcnOracle = cnOracle
    zlInitComponents = True
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
    zlCardDevSet = mobjCard.zlCardDevSet(frmMain, lngCardCode)
    
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
   zlBrushCardSquare = mobjCard.zlBrushCardSquare(intCallType, Me, lng�ӿڱ��, rsFeeList, dbl�������, rsSquare)
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
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-06-18 11:33:22
    '˵������Ԥ�������������Ԥ��ʱ�����ȷ����ťʱ����(����ǰ����)
    '------------------------------------------------------------------------------------------------------------------------
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
End Function
Public Function zlReadCard(frmMain As Object, ByRef strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ӿ�:�ݲ����ǲ�����ؽӿ�
    '���:frmMain-���õ�������
    '����:strCardNo-���صĿ���
    '����:��������    True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:44:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjCard Is Nothing Then Exit Function
    If mobjCurCard Is Nothing Then Exit Function
    zlReadCard = mobjCard.zlReadCard(Me, strCardNo)
End Function

Public Function zlCloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    '����ģ�壬��ʵ�־��幦��
End Function
Private Sub Class_Terminate()
    Call zlCloseCard
    Err = 0: On Error Resume Next
    If Not mobjCard Is Nothing Then Set mobjCard = Nothing
    
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub

'----------------------
'�������ڲ�����
'----------------------
Private Sub Class_Initialize()
    On Error GoTo errHandle
    Set mobjCard = Nothing
    mstrInterFace = ""
    gDebug = GetSetting("ZLSOFT", "����ģ��\zlSquareCard", "����", 0) = 1
    glngInstanceCount = glngInstanceCount + 1
    Exit Sub
errHandle:
End Sub

Public Property Get zlobjCard() As Object
    Set zlobjCard = mobjCard
End Property

Public Property Let zlobjCard(ByVal vNewValue As Object)
    Set mobjCard = vNewValue
End Property

Public Property Get zlSetSquareNo() As Long
    zlSetSquareNo = mlng�ӿ����
End Property

 