VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChargeBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mfrmChargeBillTotal As frmChargeBillTotal
Public Enum TotalType
    EM_�շ�Ա���� = 1
    EM_С���տ� = 2
    EM_С������ = 3
    EM_�����տ� = 4
End Enum

Private Sub initChargeBillTotalWin()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����տƱ�ݻ���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-03 18:22:50
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmChargeBillTotal Is Nothing Then
        Set mfrmChargeBillTotal = New frmChargeBillTotal
        Load mfrmChargeBillTotal
    End If
End Sub

Public Function InitOracle(cnMain As ADODB.Connection, ByVal strDbUser As String, lngSys As Long) As Boolean
    '-------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '���: strDBUser-���ݿ�������
    '     lngSys-ϵͳ��
    '     cnMain-���ݿ�����
    ' ����:��ʼ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-08-13 10:31:00
    '˵��:
    '-------------------------------------------------------------------------------------------------
    Set gcnOracle = cnMain
    If Not CheckValid Then Exit Function
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    gstr��λ���� = GetUnitName()
    
    glngSys = lngSys
    gstr��λ���� = GetUnitName()
    gstrDBUser = strDbUser
    InitOracle = True
End Function

Public Function ChargeBillHaveData() As Boolean
     '-------------------------------------------------------------------------------------------------
    '����:��ǰ�����Ƿ��������
    '����:��������,����true,���򷵻�False
    '����:������
    '����:2017-4-7
    '˵��:
    '-------------------------------------------------------------------------------------------------
    On Error GoTo errH
    If mfrmChargeBillTotal.mrsList Is Nothing And mfrmChargeBillTotal.mrsListBill Is Nothing Then
        ChargeBillHaveData = False
        Exit Function
    End If
    If Not mfrmChargeBillTotal.mrsList Is Nothing Then
        If mfrmChargeBillTotal.mrsList.State = 1 Then
            If mfrmChargeBillTotal.mrsList.RecordCount <> 0 Then
                ChargeBillHaveData = True
                Exit Function
            End If
        End If
    End If
    If Not mfrmChargeBillTotal.mrsListBill Is Nothing Then
        If mfrmChargeBillTotal.mrsListBill.State = 1 Then
            If mfrmChargeBillTotal.mrsListBill.RecordCount <> 0 Then
                ChargeBillHaveData = True
                Exit Function
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub ChargeRollingListShow(ByVal frmMain As Object, _
            ByVal bytType As TotalType, ByVal strChargeRollingID As String, ByVal lngModule As Long, _
            ByVal strPrivs As String, Optional ByVal dtStartDate As Date, _
            Optional ByVal dtEndDate As Date, Optional blnDel As Boolean)
    '-------------------------------------------------------------------------------------------------
    '����:�շ�Ա������ϸ��ʾ�ӿ�
    '���:frmMain-���õ�������
    '����bytType:1-�շ�Ա���ʣ�2-С���տ�;3-С������;4-�����տ
    '       lngChargeRollingID -�շ�Ա������ID
    '      dtStartDate-��ѡ����,��ʼ����ʱ��,lngChargeRollIngID=0ʱ�����봫��
    '      dtEndDate-��ѡ��������������ʱ��,lngChargeRollIngID=0ʱ�����봫��
    '      blnDel-�Ƿ����ϼ�¼
    '����:���˺�
    '����:2013-08-13 10:31:00
    '˵��:
    '-------------------------------------------------------------------------------------------------
    Dim frmNew As frmChargeBillList
    Set frmNew = New frmChargeBillList
    Load frmNew
   Call frmNew.ShowMe(frmMain, lngModule, strPrivs, bytType, strChargeRollingID, dtStartDate, dtEndDate, blnDel)
   If Not frmNew Is Nothing Then Unload frmNew
   Set frmNew = Nothing
 End Sub
 
Public Function ChargeRollingCurtain(ByVal frmMain As Object) As Boolean
        '-------------------------------------------------------------------------------------------------
        '����:�շ�Ա���˽ӿ�
        '���:frmMain-���õ�������
        '        strOperatorName-�շ�Ա
        '����:�շ����ʳɹ�һ������,����true,���򷵻�False
        '����:���˺�
        '����:2013-08-13 10:31:00
        '˵��:
        '-------------------------------------------------------------------------------------------------
        Dim strTitle As String, frmTemp As Form
        glngModul = 1506:
        gstrPrivs = ";" & GetPrivFunc(glngSys, 1506) & ";"
        If zlStr.IsHavePrivs(gstrPrivs, "����") = False Then
            MsgBox "��û�в����շ����ʹ�����Ȩ��,����ϵͳ����Ա��ϵ", vbExclamation + vbOKOnly, gstrSysName
            Exit Function
        End If
        strTitle = "�շ����ʹ���"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Function
            End If
        Next
        Dim frmNew As frmChargeRollingCurtainMgr
        Set frmNew = New frmChargeRollingCurtainMgr
'        Load frmNew
        ChargeRollingCurtain = frmNew.zlShowChargeRollingCourtain(frmMain, glngModul, gstrPrivs)
'        If Not frmNew Is Nothing Then Unload frmNew
'        Set frmNew = Nothing
End Function

Public Sub CheckCash(ByVal frmMain As Object, Optional ByVal dblMoney As Double)
    '-------------------------------------------------------------------------------------------------
    '����:����ֽ�
    '���:frmMain-���õ�������
    '       dblMoney -��ǰ����Ľ��.
    '����:���˺�
    '����:2013-08-13 10:31:00
    '˵��:
    '-------------------------------------------------------------------------------------------------
    frmMoneyEnum.ShowMe frmMain, dblMoney
End Sub

Public Function GetHandIn() As Double
    '-------------------------------------------------------------------------------------------------
    '����:��ȡ�����ֽ���
    '����:������
    '����:2014-12-04 09:34:33
    '˵��:
    '-------------------------------------------------------------------------------------------------
    GetHandIn = mfrmChargeBillTotal.GetCashMoney
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
  '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ9����)��1-���(12��);>1: Ϊָ�����ֺ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-03 18:05:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call initChargeBillTotalWin
    Call mfrmChargeBillTotal.SetFontSize(bytSize)
End Sub
Public Function GetChargeAndBillTotalForm() As Object
    '------------------------------------------------------------------------
    '����:��ȡ"�տƱ�ݻ��ܴ���"����
    '����:��ȡ�������
    '����:���˺�
    '����:2013-08-31 13:26:15
    '˵��:���������ʼ��ʱֻ����һ�λ�ȡ,��������󶨸ö���
    '-------------------------------------------------------------------------
    Call initChargeBillTotalWin
    Set GetChargeAndBillTotalForm = mfrmChargeBillTotal
 End Function
 Public Sub ClearChargeAndBillTotalForm()
    '------------------------------------------------------------------------
    '����:���"�տƱ�ݻ��ܴ���"�����������Ϣ
    '����:���˺�
    '����:2013-08-31 13:26:15
    '-------------------------------------------------------------------------
    Call initChargeBillTotalWin
    Call mfrmChargeBillTotal.ClearData
 End Sub
 
 Public Function LoadChargeAndBillTotalData(ByVal frmMain As Object, _
      ByVal lngModule As Long, ByVal strPrivs As String, _
      ByVal bytType As TotalType, ByVal lngChargeRollingID As Long, _
      Optional ByVal dtStartDate As Date, Optional ByVal dtEndDate As Date, _
      Optional blnOlnyView As Boolean = True, _
      Optional ByVal blnDel As Boolean = False, _
      Optional strPersonName As String = "", _
      Optional strRollingType As String) As Boolean
    '-------------------------------------------------------------------------------------------------
    '����:�շ�Ա���˽ӿ�
    '���:frmMain-���õ�������
    '    lngModule-ģ���
    '    strPrivs-Ȩ�޴�
    '����bytType:1-�շ�Ա���ʣ�2-С���տ�;3-С������;
    '            4-�����տ�(����շ�Ա��������տ�)������տ��ѯ;
    '            5-�����տ�(����Է��շ�Ա�տ�)��
    '    lngChargeRollingID -�շ�Ա������ID
    '    dtStartDate-��ѡ����,��ʼ����ʱ��,lngChargeRollIngID=0ʱ�����봫��
    '    dtEndDate-��ѡ��������������ʱ��,lngChargeRollIngID=0ʱ�����봫��
    '    blnOlnyView-���ܲ鿴(���ܱ���Ʊ�ݺ���)
    '    blnDel-�Ƿ����ϼ�¼
    '    strPersonName-ָ�����շ�Ա,��Ϊ"-"ʱ,��ʾֻ�������
    '    strRollingType-�������(0-�������(��ȫ������),1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨)
    '����:���ݼ��سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-08-13 10:31:00
    '-------------------------------------------------------------------------------------------------
    Call initChargeBillTotalWin
    LoadChargeAndBillTotalData = mfrmChargeBillTotal.LoadChargeAndBillTotalData(frmMain, lngModule, strPrivs, bytType, lngChargeRollingID, dtStartDate, dtEndDate, blnOlnyView, blnDel, strPersonName, strRollingType)
 End Function
 Public Sub zlPrint(ByVal bytMode As Byte, _
    Optional strDeptName As String = "", Optional strMemo As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����б���Ϣ
    '���:bytMode=1-��ӡ,2-Ԥ��,3-�����Excel
    '       strDeptName-�տ������(�շ�Ա����ʱת��)
    '       strMemo-��ע(�շ�Ա����ʱת��)
    '����:���˺�
    '����:2013-09-13 10:23:30
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmChargeBillTotal.zlPrint(bytMode, strDeptName, strMemo)
End Sub
Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mfrmChargeBillTotal Is Nothing Then Unload mfrmChargeBillTotal
    Set mfrmChargeBillTotal = Nothing
End Sub


