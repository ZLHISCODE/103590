VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsCard"
Option Explicit

Private mObjCard As clsCardDev  '���ӿڶ���
Public gcnOracle As ADODB.Connection '����ҽ����һ��ͨ�ӿ���Ҫʹ��.
Private mstrInterFace As String '����ʱѡ��Ľӿ�����

Private mfrmTimer As frmTimer
Private mlngHwnd As Long

Public Event ShowICCardInfo(ByVal strCardNO As String)

Friend Function GetParent() As Long
    GetParent = mlngHwnd
End Function

Public Sub SetParent(ByVal lngHwnd As Long)
    mlngHwnd = lngHwnd
End Sub

Public Function GetSpare(Optional ByRef strInterFace As String) As Currency
    If Not mObjCard Is Nothing Then
        GetSpare = mObjCard.GetSpare
        strInterFace = mstrInterFace
    End If
End Function

Public Function ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean
'ע��,���ñ��ӿ�֮ǰӦ�ȵ������ӿ�,��ȷ��������һ���ӿڣ��ж���ӿ�ʱ��
    If Not mObjCard Is Nothing Then
        ReturnSwap = mObjCard.ReturnSwap(strCardNO, lngHOSCode, strTransFlow, dblMoney)
    End If
End Function

Public Function PaymentSwap(ByVal dblMoney As Double, ByRef dblBalance As Double, ByRef intCardType As Integer, _
    ByVal lngHOSCode As Long, ByRef strCardNO As String, ByRef strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean
    If Not mObjCard Is Nothing Then
        PaymentSwap = mObjCard.PaymentSwap(dblMoney, dblBalance, intCardType, lngHOSCode, strCardNO, strTransFlow, strBalanceID, strPatiID)
    End If
End Function

Public Function UploadSwap(ByVal lng����ID As Long, ByRef strOut As String) As Boolean
    '����ģ�壬��������������ϴ��Ĺ�����Ŀǰ��Ӧ��������һ��ͨ����ɲ�����Ϣ������Ϣ���ϴ�
    On Error GoTo errHand
    
    If mObjCard Is Nothing Then
        Set mObjCard = CreateObject("zlICCard.clsIC_NBYKT")
        If mObjCard Is Nothing Then
            gLastErr = "�����ӿڶ���δ������"
            Exit Function
        End If
        Set gcnConnect = gcnOracle
        If Not mObjCard.initCard Then Exit Function
    End If
    UploadSwap = mObjCard.UploadSwap(lng����ID, strOut)
    Exit Function
errHand:
    gLastErr = err.Description
End Function

Friend Sub ShowICCardInfo(ByVal strCardNO As String)
    RaiseEvent ShowICCardInfo(strCardNO)
End Sub

Public Function Read_Card(Optional ByVal FrmMain As Object) As String
    '����-��������
    '���:������
    '����:����,��Ϊ�մ�,��ʾ����ʧ��
    Dim blnCancel As Boolean
    Dim strClsCard As String '���ӿ�����,
    Dim objCard As clsCard, int���� As Integer, i As Integer, intReturn As Integer
    Dim objInsure As Object, intType As Integer
    
    On Error GoTo errHandle
    
    Set gcnConnect = gcnOracle
    '��ʾ�������õĿ�
    For Each objCard In Cards
        If objCard.���� = True Then
            i = i + 1
            int���� = objCard.����
            strClsCard = objCard.�ӿڳ�����
            mstrInterFace = objCard.����
        End If
    Next
    If i = 0 Then
        mfrmTimer.tmrMain.Enabled = False
        MsgBox "��ǰû�������κζ����ӿڣ�", vbInformation, "IC���ӿ�"
        Exit Function
    End If
    If i <> 1 Then
        '�ж�����ÿ�,Ҫ�û�ѡ��
        If Cards.Count > 0 Then
            intType = frmCardSelect.SelectCard(gCol, i, FrmMain)
            
            For Each objCard In Cards
                If objCard.���� = intType Then
                    int���� = objCard.����
                    strClsCard = objCard.�ӿڳ�����
                    mstrInterFace = objCard.����
                    gLastErr = ""
                    Exit For
                Else
                    int���� = 0
                    strClsCard = ""
                End If
                
            Next
            If strClsCard = "" Then
                gLastErr = "�û�ȡ��ѡ����޿��ö�������"
            End If
        Else
            gLastErr = "�޿��ö�������"
            GoTo errSow
        End If
    End If
    
    '-----------------------------
    If strClsCard <> "" Then
        If int���� <> 0 Then
            If Not gcnOracle Is Nothing Then
                If gcnOracle.State = 1 Then
                    Set objInsure = CreateObject("zl9Insure.clsInsure")
                    Call objInsure.InitOracle(gcnOracle)
                    Read_Card = objInsure.readcard(int����, intReturn)
                    If intReturn <> 1 Then
                        gLastErr = Read_Card
                        Read_Card = ""
                    End If
                Else
                    Read_Card = ""
                    gLastErr = "���ݿ�����δ�򿪣����ܽ���ҽ��������"
                End If
            Else
                Read_Card = ""
                gLastErr = "δ��ʼ�����ݿ����ӣ����ܽ���ҽ��������"
            End If
        Else
            If i <> 1 Then
                Set mObjCard = CreateObject(strClsCard)
            End If
            If mObjCard Is Nothing Then
                Read_Card = ""
                gLastErr = "�����ӿڶ���δ������"
                GoTo errSow
            End If
            If mObjCard.initCard Then
                Read_Card = mObjCard.readcard("", blnCancel)
            End If
            mObjCard.CloseCard
        End If
    End If
    
errSow:
    '����������ؿ���ȡ��Ϊ��,˵������ʧ��
    If gbln�Զ���ȡ = False Then
        If Read_Card = "" And blnCancel = False Then
            If gLastErr = "" Then
                gLastErr = "����ʱ����δ֪����"
            End If
            Call WritLog("ICCard.read_card", "", gLastErr)
            MsgBox gLastErr, vbExclamation, "IC���ӿ�"
        End If
    End If
    Exit Function
errHandle:
    Read_Card = ""
    If Not mObjCard Is Nothing Then mObjCard.CloseCard
    gLastErr = "����ʱ��������" & vbNewLine & err.Description
    Call WritLog("ICCard.read_card", "", gLastErr)
    MsgBox gLastErr, vbExclamation, "IC���ӿ�"
End Function

Public Function Set_Card(Optional ByVal FrmMain As Object) As Boolean
    '����-�������в�������
    Dim objCard As clsCard
    On Error GoTo errHandle
    If Cards.Count > 0 Then
        'ǰ̨����ʱ����ʾ���п��õĶ����������û�ѡ��Ȼ��������á�
        Call frmCardSelect.SelectCard(gCol, -1, FrmMain)
        For Each objCard In Cards
            objCard.���� = GetSetting("ZLSOFT", "����ģ��\zlICCard", objCard.����, 0) = 1
        Next
        gDebug = GetSetting("ZLSOFT", "����ģ��\zlICCard", "����", 0) = 1
        Set_Card = True
    Else
        gLastErr = "�޿��õĶ�������"
    End If
    Exit Function
errHandle:
    gLastErr = "���ö�������������" & vbNewLine & err.Description
    Call WritLog("ICCard.Set_Card", "", gLastErr)
    MsgBox gLastErr, vbExclamation, "IC���ӿ�"
End Function

Public Property Get ErrorMessage() As String
    '��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
    'Syntax: Debug.Print X.ErrorMessage
    ErrorMessage = gLastErr
End Property



Public Sub SetEnabled(ByVal blnEnabled As Boolean)
    If gbln�Զ���ȡ = True Then
        If mfrmTimer Is Nothing Then Exit Sub
        mfrmTimer.tmrMain.Enabled = blnEnabled
    End If
End Sub

'----------------------
'�������ڲ�����
'----------------------

Private Sub Class_Initialize()
    Dim intCount�Զ���ȡ As Integer
    Dim intCount���� As Integer
    Dim i As Integer
    Dim strClsCard As String
    
    '������ clsInterfaceCard ��ʱ������ mclsCard ����
    On Error GoTo errHandle
    Set gCol = New Collection
    Set gcnOracle = Nothing
    Set mObjCard = Nothing
    mstrInterFace = ""
    
    '--------------
    '��ʱ�ֹ�����һ�����Կ���Ҫ��Ϊ�������ļ��ж�ȡ��
    
    Call initCards
    
    For i = 1 To Cards.Count
        If Item(i).�Ƿ��Զ���ȡ = 1 And Item(i).���� = True Then    '�жϵ�ǰ���Ƿ����ò����Զ���ȡ
            Set mfrmTimer = New frmTimer
            Load mfrmTimer
            mfrmTimer.tmrMain.Interval = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard", "�Զ���ȡ���" & Item(i).����, 300))
            intCount�Զ���ȡ = intCount�Զ���ȡ + 1
            intCount���� = intCount���� + 1
            strClsCard = Item(i).�ӿڳ����� '��ȡ�ӿڳ�����
        ElseIf Item(i).���� = True Then
            intCount���� = intCount���� + 1
            strClsCard = Item(i).�ӿڳ����� '��ȡ�ӿڳ�����
        End If
    Next
    gbln�Զ���ȡ = intCount�Զ���ȡ = 1 And intCount���� = 1
    
    If intCount���� = 1 Then   '�Զ�����
        Set mObjCard = CreateObject(strClsCard)
    End If
    If gbln�Զ���ȡ = True Then Call mfrmTimer.Init(Me)
    
    gDebug = GetSetting("ZLSOFT", "����ģ��\zlICCard", "����", 0) = 1
    
    '��ʼ����������
    If gobjComLib Is Nothing Then
        Set gobjComLib = GetObject("", "zl9Comlib.clsComlib")
        Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
        Set gobjControl = GetObject("", "zl9Comlib.clsControl")
        Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    End If
    
    
    Exit Sub
errHandle:
    gLastErr = "��ʼ����ʧ��!" & vbNewLine & err.Description
    Call WritLog("ICCard.Initialize", "", gLastErr)
    MsgBox gLastErr, vbExclamation, "IC���ӿ�"
End Sub
