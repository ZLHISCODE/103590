VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPassMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsVSCOL"
Option Explicit

'��������ֵ�ľֲ�����
Private mVSCOL As clsVSCOL  '�ֲ�����
Private mlngModel As PASS_MODEL '�ֲ�����
Private mvsAdvice As Object  '�ݴ�ҽ������
Private mobjCmdBar As Object         '�˵�����
Private mbytUseType As Byte              'bytUseType =1 �ٴ�·������ҽ��ʱ����סԺҽ���༭����
Private mfrmMain As Object               '������
Private mDiags As clsDiags         '������϶��� ���ﴫ��
Private mPassPati As clsPatient      '������Ϣ
Private mint���� As Integer        ' ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)

Public Property Set PassPati(ByVal clsPati As clsPatient)
    Set mPassPati = clsPati
End Property

Public Property Get PassPati() As clsPatient
    Set PassPati = mPassPati
End Property

Public Property Set frmMain(ByVal vData As Object)
    Set mfrmMain = vData
End Property

Public Property Get frmMain() As Object
    Set frmMain = mfrmMain
End Property

Public Property Let bytUseType(ByVal bytData As Byte)
    mbytUseType = bytData
End Property

Public Property Get bytUseType() As Byte
     bytUseType = mbytUseType
End Property

Public Property Set Diags(ByVal clsData As clsDiags)
    Set mDiags = clsData
End Property

Public Property Get Diags() As clsDiags
    Set Diags = mDiags
End Property

Public Property Let lngModel(ByVal PMData As PASS_MODEL)
    mlngModel = PMData
End Property

Public Property Get lngModel() As PASS_MODEL
    lngModel = mlngModel
End Property

Public Property Set VSCOL(ByRef clsVSData As clsVSCOL)
    Set mVSCOL = clsVSData
End Property

Public Property Get VSCOL() As clsVSCOL
    Set VSCOL = mVSCOL
End Property

Public Property Set vsAdvice(ByRef vsData As Object)
    Set mvsAdvice = vsData
End Property

Public Property Get vsAdvice() As Object
    Set vsAdvice = mvsAdvice
End Property

Public Property Set objCmdBar(ByRef objData As Object)
    Set mobjCmdBar = objData
End Property

Public Property Get objCmdBar() As Object
    Set objCmdBar = mobjCmdBar
End Property

Public Property Let int����(ByVal intData As Integer)
    mint���� = intData
End Property

Public Property Get int����() As Integer
     int���� = mint����
End Property

Public Function GetVSCOL(Optional ByVal intCOLID As Integer = -1, Optional ByVal intCOL���ID As Integer = -1, Optional ByVal intCOL������� As Integer = -1, _
    Optional ByVal intCOL������ĿID As Integer = -1, Optional ByVal intCOL�շ�ϸĿID As Integer = -1, Optional ByVal intCOLҽ������ As Integer = -1, _
    Optional ByVal intCOL��Ч As Integer = -1, Optional ByVal intCOL���� As Integer = -1, Optional ByVal int������λ As Integer = -1, _
    Optional ByVal intCOL�÷� As Integer = -1, Optional ByVal intCOL���� As Integer = -1, Optional ByVal intCOLӤ�� As Integer = -1, _
    Optional ByVal intCOL����ʱ�� As Integer = -1, Optional ByVal intCOL����ҽ�� As Integer = -1, Optional ByVal intCOL��ʼʱ�� As Integer = -1, _
    Optional ByVal intCOL��������ID As Integer = -1, Optional ByVal intCOL��ֹʱ�� As Integer = -1, Optional ByVal intCOLƵ�� As Integer = -1, _
    Optional ByVal intCOLƵ�ʴ��� As Integer = -1, Optional ByVal intCOLƵ�ʼ�� As Integer = -1, Optional ByVal intCOL�����λ As Integer = -1, _
    Optional ByVal intCOL��ʾ As Integer = -1, Optional ByVal intCOL��� As Integer = -1, Optional ByVal intCOL״̬ As Integer = -1, _
    Optional ByVal intCOLEDIT As Integer = -1, Optional ByVal intCOL����ID As Integer = -1, Optional ByVal intCOL��ҳID As Integer = -1, _
    Optional ByVal intCOLѡ�� As Integer = -1, Optional ByVal intCOLִ������ As Integer = -1, Optional ByVal intCOLҩƷ���� As Integer = -1, _
    Optional ByVal intCOL���� As Integer = -1, Optional ByVal intCOLNO As Integer = -1, Optional ByVal intCOLסԺ�� As Integer = -1, _
    Optional ByVal intCOL����� As Integer = -1, Optional ByVal intCOL����� As Integer = -1, Optional ByVal intCOL���� As Integer = -1, _
    Optional ByVal intcol������λ As Integer = -1, Optional ByVal intcolҽ������ As Integer = -1, Optional ByVal intcol��ҩĿ�� As Integer = -1, _
    Optional ByVal intCol�������� As Integer, Optional ByVal intCol����ҩƷ˵�� As Integer, Optional ByVal intcol��ҩ���� As Integer, _
    Optional ByVal intCol��־ As Integer, Optional ByVal intCol������ As Integer, Optional ByVal intColִ�з��� As Integer, _
    Optional ByVal intColִ�п���ID As Integer) As clsVSCOL
    
    Dim objVsCOL As clsVSCOL
    
    Set objVsCOL = New clsVSCOL
    
    With objVsCOL
        .intCOLID = intCOLID
        .intCOL���ID = intCOL���ID
        .intCOL������� = intCOL�������
        .intCOL������ĿID = intCOL������ĿID
        .intCOL�շ�ϸĿID = intCOL�շ�ϸĿID
        .intCOLҽ������ = intCOLҽ������
        .intCOL��Ч = intCOL��Ч
        .intCOL���� = intCOL����
        .intCOL������λ = int������λ
        .intCOL�÷� = intCOL�÷�
        
        .intCOL���� = intCOL����
        .intCOLӤ�� = intCOLӤ��
        .intCOL����ʱ�� = intCOL����ʱ��
        .intCOL����ҽ�� = intCOL����ҽ��
        .intCOL��ʼʱ�� = intCOL��ʼʱ��
        .intCOL��������ID = intCOL��������ID
        .intCOL��ֹʱ�� = intCOL��ֹʱ��
        .intCOLƵ�� = intCOLƵ��
        .intCOLƵ�ʴ��� = intCOLƵ�ʴ���
        .intCOLƵ�ʼ�� = intCOLƵ�ʼ��
        
        .intCOL�����λ = intCOL�����λ
        .intCOL��ʾ = intCOL��ʾ
        .intCOL��� = intCOL���
        .intCOL״̬ = intCOL״̬
        .intCOLEDIT = intCOLEDIT
        .intCOL����ID = intCOL����ID
        .intCOL��ҳID = intCOL��ҳID
        .intCOLѡ�� = intCOLѡ��       'gbytUseType =1ʱ����
        
        .intCOLִ������ = intCOLִ������   '
        .intCOLҩƷ���� = intCOLҩƷ����   '������ҩ����
        'ΪҩƷģ���ṩ
        .intCOL���� = intCOL����
        .intCOLNO = intCOLNO
        .intCOLסԺ�� = intCOLסԺ��
        .intCOL����� = intCOL�����
        .intCOL����� = intCOL�����
        .intCOL���� = intCOL����
        .intcol������λ = intcol������λ
        .intcolҽ������ = intcolҽ������
        .intcol��ҩĿ�� = intcol��ҩĿ��
        .intcol��ҩ���� = intcol��ҩ����
        .intCol�������� = intCol��������
        .intCol����ҩƷ˵�� = intCol����ҩƷ˵��
        .intCol��־ = intCol��־
        .intCol������ = intCol������
        .intColִ�з��� = intColִ�з���
        .intColִ�п���ID = intColִ�п���ID
    End With
    Set GetVSCOL = objVsCOL
End Function

Public Function GetPatient(Optional lng����ID As Long = 0, Optional ByVal lng��ҳID As Long = 0) As clsPatient
    Dim objPati As clsPatient
    Set objPati = New clsPatient
    objPati.lng����ID = lng����ID
    objPati.lng��ҳID = lng��ҳID
    
    Set GetPatient = objPati
    Set objPati = Nothing
End Function

Public Function GetDiags() As clsDiags
    Set GetDiags = New clsDiags
End Function

Private Sub Class_Terminate()
    Set mVSCOL = Nothing
    Set mvsAdvice = Nothing
    Set mobjCmdBar = Nothing
    Set mfrmMain = Nothing
    Set mDiags = Nothing
End Sub