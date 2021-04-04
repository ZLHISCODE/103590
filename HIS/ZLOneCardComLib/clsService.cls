VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:���е��÷���ӿ�
'һ��������ӿ�
'       1.zlPatiSvr_GetPatiInfo-����������ȡ������ϸ��Ϣ����ӿ�
'       2.zl_PatiSvr_GetPatiInfsByRange-��������Χ������ȡ������Ϣ����ӿ�
'       3.zlPatiSvr_GetCardTypes-��ȡҽ�ƿ������Ϣ����ӿ�
'       4.zlPatiSvr_GetPatiID-����������ȡ����IDֵ����ӿ�
'       5.zlPatiSvr_GetPatiExtendInfo-��ȡ������Ϣ�ӱ���Ϣ����ӿ�
'       6.ZlPatisvr_GetPatiCardInfo-��ȡ���˵���Ч������ӿ�
'       7.zlPatiSvr_GetInsureByPatiID-��ȡ���˵�ҽ����Ϣ����ӿ�
'����������ӿ�
'  1.zlExseSvr_GetPatiSurplusInfo-��ȡ���˷��������Ϣ
'  2.zlExseSvr_CardSquareTypes-��ȡ���ѿ�������ӿ�
'�����ٴ���ӿ�
'  1.zlCisSvr_GetPatPageInfByRange-����һ����Χ��������ȡ���˲�����Ϣ
'  2.zlCisSvr_GetPatiID-���ݴ��Ż�סԺ�ţ���ȡ����ID
'  3.zlCIsSvr_GetPatiPageInfo-���ݲ���id����ҳid,��ȡ������Ϣ
'����:���˺�
'����:2019-11-02 15:49:29
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Function zlPatiSvr_GetPatiInfsByRange(ByVal intQueryStatus As Integer, ByVal cllFilter As Variant, _
    ByRef cllPatiInfos_out As Collection, Optional ByVal str����Ids As String, Optional ByRef str����IDs As String, _
    Optional ByVal blnExpendInfo As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ��
    '���:intQueryStatus-��ѯ����(0-������;1-��Ժ ;2-���Ｐ��Ժ)
    '     cllFilter-��������
    '     str����Ids-����ID
    '     rsPatiPage-��ҳ��Ϣ
    '     str����IDs-��ǰ����Ids
    '����:cllPatiInfos_out-���ص����ݼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-30 21:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTmp As Collection
    Dim cllData As Collection, cllSubData As Collection, cllContacts As Collection
    Dim i As Long
    
    Set cllPatiInfos_out = New Collection
    If zl_PatiSvr_GetPatiInfsByRange(intQueryStatus, cllFilter, cllData, str����Ids, str����IDs, blnExpendInfo) = False Then Exit Function
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then zlPatiSvr_GetPatiInfsByRange = True: Exit Function
    For i = 1 To cllData.count
        Set cllSubData = cllData(i): Set cllTmp = New Collection
        cllTmp.Add cllSubData("_pati_id"), "����ID"
        
        cllTmp.Add NVL(cllSubData("_outpatient_num")), "�����"
        cllTmp.Add NVL(cllSubData("_fee_category")), "�ѱ�"
        cllTmp.Add NVL(cllSubData("_pati_name")), "����"
        cllTmp.Add NVL(cllSubData("_pati_sex")), "�Ա�"
        cllTmp.Add NVL(cllSubData("_pati_age")), "����"
        cllTmp.Add NVL(cllSubData("_pati_birthdate")), "��������"
        cllTmp.Add NVL(cllSubData("_pati_idcard")), "����֤��"
        cllTmp.Add NVL(cllSubData("_pati_birthplace")), "�����ص�"
        cllTmp.Add NVL(cllSubData("_pat_home_phno")), "��ͥ�绰"
        cllTmp.Add NVL(cllSubData("_pat_home_addr")), "��ͥ��ַ"
        cllTmp.Add NVL(cllSubData("_phone_number")), "�ֻ���"
        cllTmp.Add NVL(cllSubData("_emp_name")), "������λ����"
        cllTmp.Add NVL(cllSubData("_pati_identity")), "����"
        cllTmp.Add NVL(cllSubData("_ocpt_name")), "ְҵ"
        cllTmp.Add NVL(cllSubData("_pati_education")), "ѧ��"
        cllTmp.Add NVL(cllSubData("_vcard_no")), "���￨��"
        cllTmp.Add NVL(cllSubData("_create_time")), "�Ǽ�ʱ��"
        cllTmp.Add NVL(cllSubData("_inpatient_num")), "סԺ��"
        cllTmp.Add NVL(cllSubData("_pati_bed")), "����"
        cllTmp.Add NVL(cllSubData("_insurance_type")), "����"

        If blnExpendInfo Then
            If CollectionExitsValue(cllData, "_contacts") Then
                Set cllContacts = cllData("_contacts")
                cllTmp.Add NVL(cllContacts("_name")), "��ϵ������"
                cllTmp.Add NVL(cllContacts("_phone")), "��ϵ�˵绰"
            Else
                cllTmp.Add "", "��ϵ������"
                cllTmp.Add "", "��ϵ�˵绰"
            End If
            cllTmp.Add NVL(cllSubData("_mdlpay_mode_name")), "ҽ�Ƹ��ʽ����"
            cllTmp.Add NVL(cllSubData("_visit_time")), "����ʱ��"
            cllTmp.Add NVL(cllSubData("_insurance_num")), "ҽ����"
        End If
        cllPatiInfos_out.Add cllTmp, "_" & cllTmp("����ID")
    Next
    
    zlPatiSvr_GetPatiInfsByRange = True
End Function
Public Function zlPatiSvr_GetCardTypes(ByRef cllCardTypes_out As Variant) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ�ƿ�����
    '���:
    '����:cllCardTypes_out-���صĿ�����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 16:53:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlPatiSvr_GetCardTypes = zl_PatiSvr_GetCardTypes(cllCardTypes_out)
End Function

Public Function zlPatiSvr_GetPatiID(ByVal cllFindCons As Collection, ByVal cllOtherFindCons As Collection, _
    ByRef cllPatiDatas_Out As Collection, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg As String, _
    Optional ByVal bln���ʹ��ʱ�� As Boolean = True, Optional ByVal bln���ͣ�û��ʧ As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ID��Ϣ
    '       cllFindCons-��������(array(�ӵ�����,�ӵ�ֵ))
    '                �ӵ����ư���:�����ID,����,��ά��,������,������)
    '       cllOtherFindCons-������������:array(��ѯ������,��ѯ������)
    '                   ��ѯ������:��:�����,���￨�ţ�����֤�ŵ�
    '       blnNotShowErrMsg-����ʾ�������ʾ��Ϣ
    '      bln���ʹ��ʱ��-������������Ч
    '      bln���ͣ�û��ʧ-������������Ч
    '����:strErrMsg-���صĴ�����Ϣ
    '        lng����ID-���صĲ���ID
    '        cllPatiDatas_Out-���ز�����Ϣ����
    '����:���ҳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-03-19 09:36:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
   zlPatiSvr_GetPatiID = zl_PatiSvr_GetPatiID(cllFindCons, cllOtherFindCons, cllPatiDatas_Out, blnNotShowErrMsg, strErrMsg, bln���ʹ��ʱ��, bln���ͣ�û��ʧ)
End Function

Public Function zlPatiSvr_GetPatiIDByComminuty(ByVal int������� As Integer, ByVal str������ As String, _
                ByRef lng����ID As Long, Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ���������Ų��Ҳ���ID
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/15 17:50
    '---------------------------------------------------------------------------------------
    Dim cllFindCons As Collection, cllPati As Collection
    Set cllFindCons = New Collection
    cllFindCons.Add Array("�������", int�������)
    cllFindCons.Add Array("������", str������)
    zlPatiSvr_GetPatiIDByComminuty = zl_PatiSvr_GetPatiID(cllFindCons, Nothing, cllPati, blnNotShowErrMsg, strErrMsg)
    If cllPati Is Nothing Then Exit Function
    If cllPati.count = 0 Then Exit Function
    lng����ID = cllPati("_pati_id")
End Function

Public Function zlCisSvr_PatiIsInhospital(ByVal lng����ID As Long, ByRef blnInhospital As Boolean, _
                Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��鲡���Ƿ���Ժ����
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/18 14:35
    '---------------------------------------------------------------------------------------
    zlCisSvr_PatiIsInhospital = Zl_CisSvr_PatiIsInhospital(lng����ID, blnInhospital, blnNotShowErrMsg, strErrMsg)
End Function

Public Function zlPatiSvr_GetPatiInfo(ByVal lng����ID As Long, _
    ByVal cllOtherFindCons As Collection, ByRef cllPatiDatas_Out As Collection, _
    Optional ByVal int��ѯ���� As Integer = 0, _
    Optional ByVal bln�������� As Boolean, _
    Optional ByVal bln��������ҩ�� As Boolean, _
    Optional ByVal bln����������Ϣ As Boolean, _
    Optional ByVal bln��������Ϣ As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������ϸ��Ϣ����ӿ�
    '���:cllOtherFindCons-������������(array(��ѯ����,��ѯֵ)
    '             ��ѯ����:����IDS,����,�Ա�,�������ڵ�,��query_cons_list[]�б��е���������
    '      int��ѯ����-0-����;1-����+��ϵ��;2-����
    '����:cllPatiDatas_Out-���ز�����Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 18:02:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlPatiSvr_GetPatiInfo = zl_PatiSvr_GetPatiInfo(lng����ID, cllOtherFindCons, cllPatiDatas_Out, int��ѯ����, bln��������, bln��������ҩ��, bln����������Ϣ, bln��������Ϣ, blnNotShowErrMsg, strErrMsg)
End Function

Public Function zlPatiSvr_GetPatiExtendInfo(ByVal lng����ID As Long, ByVal str��Ϣ���� As String, ByRef cllPatiData_Out As Collection, Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ�ӱ���Ϣ����ӿ�
    '���:str��Ϣ����-����ö��ŷ���,�磺ҽѧ��ʾ,��ϵ��2,��ϵ��3��
    '
    '����:cllPatiData_Out-���ز��˴ӱ���Ϣ���ݼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 20:10:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlPatiSvr_GetPatiExtendInfo = zl_PatiSvr_GetPatiExtendInfo(lng����ID, str��Ϣ����, cllPatiData_Out, blnNotShowErrMsg, strErrMsg)
End Function
Public Function ZlPatisvr_GetPatiCardInfo(ByVal strCardTypeIDs As String, ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ӹ���������м���ָ�����˳�����Ч���Ŀ����
    '���: strCardTypeIDs ��������𣬶���ö��ŷָ�
    '����:���ز��˳�����Ч���Ŀ���𣬶���ö��ŷָ�
    '����:���˺�
    '����:2018-12-03 15:43:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    ZlPatisvr_GetPatiCardInfo = Zl_Patisvr_GetPatiCardInfo(strCardTypeIDs, lng����ID)
End Function
Public Function zlPatiSvr_GetInsureByPatiID(lng����ID As Long, Optional ByRef int����_Out As Integer, Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ҽ�������Ƿ����δ�����
    '���:lng����ID
    '     blnNotShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:int����_Out-����
    '     strErrMsg_out-���صĴ�����Ϣֵ
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-12-05 16:40:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlPatiSvr_GetInsureByPatiID = zl_PatiSvr_GetInsureByPatiID(lng����ID, int����_Out, blnNotShowErrMsg, strErrMsg_Out)
End Function

Public Function zlPatiSvr_CheckOutNoIsExist(ByVal lng����ID As Long, ByVal str����� As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ���������Ƿ�����ʹ��
    ' ��� : str�����-������������
    ' ���� : blnUsedByOther:T:������ʹ��
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    zlPatiSvr_CheckOutNoIsExist = zl_PatiSvr_CheckOutNoIsExist(lng����ID, str�����, blnUsedByOther, blnNotShowErrMsg, strErrMsg_Out)
End Function

Public Function zlPatiSvr_PhoneNumberExist(ByVal lng����ID As Long, ByVal str�ֻ��� As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ����ֻ����Ƿ�����ʹ��
    ' ��� :
    ' ���� : blnUsedByOther:T:������ʹ��
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    zlPatiSvr_PhoneNumberExist = zl_PatiSvr_PhoneNumberExist(lng����ID, str�ֻ���, blnUsedByOther, blnNotShowErrMsg, strErrMsg_Out)
End Function

Public Function zlPatiSvr_CheckInsNoIsExist(ByVal strҽ���� As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ���ҽ�����Ƿ�����ʹ��
    ' ��� :
    ' ���� : blnUsedByOther:T:������ʹ��
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    zlPatiSvr_CheckInsNoIsExist = zl_PatiSvr_CheckInsNoIsExist(strҽ����, blnUsedByOther, blnNotShowErrMsg, strErrMsg_Out)
End Function


Public Function zlPatiSvr_GetPatiIdsByRange(ByVal strCondition As String, ByRef strPatiIds As String, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg_Out As String, _
    Optional ByVal blnFindByFilter As Boolean, Optional ByVal cllFilter As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������ֵ��ȡ���������Ĳ���ID
    '���:
    '   strCondition=�����Ǿ��￨�š�����֤�š�IC���š������
    '   blnFindByFilter=True:����������(cllFilter)��ȡ;False:��strCondition��ȡ
    '   cllFilter=��������:Array(Key,Value),Key:��ͬ��λID
    '����:
    '����:ִ�гɹ�����True��ʧ�ܷ���False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlPatiSvr_GetPatiIdsByRange = zl_PatiSvr_GetPatiIdsByRange(strCondition, strPatiIds, _
        blnNotShowErrMsg, strErrMsg_Out, blnFindByFilter, cllFilter)
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'�ٴ���ط���

Public Function zlCisSvr_GetPatPageInfByRange(ByVal intQueryStatus As Integer, ByVal cllFilter As Collection, Optional ByVal str����Ids As String, Optional ByRef str����IDs As String, _
    Optional ByRef cllPatiPages_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ����Χ��������ȡ���˲�����Ϣ
    '���:intQueryStatus-��ѯ����(0-��Ժ����;1-��Ժ����;2-��Ժ���Ժ )
    '     cllFilter-��������
    '     str����Ids-����ö���:����ID����ID:��ҳID
    '     rsPatiPage-��ҳ��Ϣ
    '     str����IDs-��ǰ����Ids
    '����:rsPatiPageInfo_Out-���صĲ�����Ϣ��
    '     strPatiIds_Out-���ص�ǰ���漰�Ĳ���IDs
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-30 21:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlCisSvr_GetPatPageInfByRange = zl_CisSvr_GetPatPageInfByRange(intQueryStatus, cllFilter, str����Ids, str����IDs, cllPatiPages_Out)
    
End Function
Public Function zlCisSvr_GetPatiID(ByVal lng����ID As Long, ByVal str���� As String, ByVal strסԺ�� As String, Optional ByRef lng��ҳID_out As Long) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݴ��ż�סԺ�Ż�ȡ����ID
    '      strסԺ��-סԺ��
    '����:lng��ҳID-���ص�ǰ���ŵ���ҳID
    '����:�ɹ����ز���ID,���򷵻�False
    '����:���˺�
    '����:2012-09-19 15:50:18
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection
    
    Set cllFilter = New Collection
    If Val(strסԺ��) <> 0 Then
        cllFilter.Add Array("סԺ��", strסԺ��)
    Else
        cllFilter.Add Array("����ID", lng����ID)
        cllFilter.Add Array("����", str����)
    End If
    zlCisSvr_GetPatiID = zl_CisSvr_GetPatiID(cllFilter, lng��ҳID_out)
End Function

Public Function zlCIsSvr_GetPatiPageInfo(ByVal int��ѯ���� As Integer, ByVal str������ҳIDs As String, ByRef cllPatiPage_Out As Variant, _
    Optional ByRef bln��ȡ���סԺ As Boolean, Optional bln��Ӥ����Ϣ As Boolean, Optional ByRef bln��ת����Ϣ As Boolean, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg_Out As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID����ҳid��Ϣ����ȡ������ҳ��Ϣ
    '���:int��ѯ����-0-ֻ��ȡ������Ϣ;1-��ȡ������Ϣ+��չ��Ϣ;2-����ȡȡ��ҳID�ֶ�
    '     str������ҳIDs-���ָ�ʽ:
    '           1.����id1:��ҳid1,����id2:��ҳid2...
    '           2.����id1,����id2,...����idn
    '      bln��ȡ���סԺ:����ȡ�������һ�εĲ���,(str������ҳIDs�ڶ��ָ�ʽ��Ч)
    '      bln��Ӥ����Ϣ:�Ƿ����Ӥ����Ϣ
    '      bln��ת����Ϣ:�Ƿ��ת����Ϣ
    '����:cllPatiPageInfo_Out-���صĲ�����Ϣ��
    '     strErrMsg_Out-���صĴ�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2012-09-19 15:50:18
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    zlCIsSvr_GetPatiPageInfo = zl_CIsSvr_GetPatiPageInfo(int��ѯ����, str������ҳIDs, cllPatiPage_Out, bln��ȡ���סԺ, bln��Ӥ����Ϣ, bln��ת����Ϣ, blnNotShowErrMsg, strErrMsg_Out)
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'����������ط���
Public Function zlExseSvr_GetPatiSurplusInfo(ByVal str����Ids As String, ByRef cllSurplusData_Out As Collection, _
    Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByVal strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���˷��������Ϣ
    '���:str����Ids-����ID,����ö��ŷ���
    '     blnNotShowErrMsg-�Ƿ���ʾ������Ϣ��,true-����ʾ;false-��ʾ
    '����:cllSurplusData_Out-���ز�����Ϣ��
    '     strErrMsg_out-����ʾʱ�����ش�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlExseSvr_GetPatiSurplusInfo = zl_ExseSvr_GetPatiSurplusInfo(str����Ids, cllSurplusData_Out, blnNotShowErrMsg, strErrMsg_Out)
End Function
Public Function zlExseSvr_GetConsumerCardType(ByRef cllTypesData_out As Collection, Optional ByVal blnOnlyStart As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ѿ������Ϣ����ӿ�
    '���:blnOnlyStart-ֻ��ȡ���õĿ����
    '����:cllTypesData_out-���ؿ������Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:47:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlExseSvr_GetConsumerCardType = zl_ExseSvr_GetConsumerCardType(cllTypesData_out, blnOnlyStart)
End Function

Public Function zlPatiSvr_GetPatiAddrssInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
                ByVal str��ַ��� As String, ByRef cllAddrList As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ���˽ṹ����ַ��Ϣ
    ' ��� : str��ַ���:��ѯ�ĵ�ַ���1-�����أ�2-����,3-��סַ,4-���ڵ�ַ,5-��ϵ�˵�ַ��6-��λ��ַ��Ϊ0ʱ��ʾ��ѯ�������͵ĵ�ַ��Ϣ
    '        ����ö��ŷָ������磺"3,4"
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim cllData As Collection, cllSubData As Collection, cllTmp As Collection
    Dim i As Integer
    
    On Error GoTo Errhand
    Set cllAddrList = New Collection
    If zl_PatiSvr_GetPatiAddrssInfo(lng����ID, lng��ҳID, str��ַ���, cllData) = False Then Exit Function
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then zlPatiSvr_GetPatiAddrssInfo = True: Exit Function
    
    For i = 1 To cllData.count
        Set cllSubData = cllData(i): Set cllTmp = New Collection
            cllTmp.Add NVL(cllSubData("_pat_addr_type")), "��ַ���"
            cllTmp.Add NVL(cllSubData("_pat_addr_state")), "ʡ"
            cllTmp.Add NVL(cllSubData("_pat_addr_city")), "��"
            cllTmp.Add NVL(cllSubData("_pat_addr_county")), "��"
            cllTmp.Add NVL(cllSubData("_pat_addr_township")), "��"
            cllTmp.Add NVL(cllSubData("_pat_addr_other")), "����"
            cllTmp.Add NVL(cllSubData("_pat_region_code")), "��������"
        cllAddrList.Add cllTmp, "_" & cllTmp("��ַ���")
    Next
    
    zlPatiSvr_GetPatiAddrssInfo = True
    Exit Function
Errhand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function ZlCissvr_GetPatiBaseInfoByPatiId(ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long = -1, _
                Optional ByRef cllPatiList As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ���˽ṹ����ַ��Ϣ
    ' ��� : lng����ID��lng��ҳID
    ' ���� : cllPati(Collect):��Ա()
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim cllData As Collection, cllSubData As Collection, cllTmp As Collection
    Dim i As Integer
    
    On Error GoTo Errhand
    Set cllPatiList = New Collection
    If Zl_Cissvr_GetPatiBaseInfo(lng����ID, lng��ҳID, cllData) = False Then Exit Function
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then ZlCissvr_GetPatiBaseInfoByPatiId = True: Exit Function
    
    For i = 1 To cllData.count
        Set cllSubData = cllData(i): Set cllTmp = New Collection
            cllTmp.Add Val(NVL(cllSubData("_pati_id"))), "����ID"
            cllTmp.Add NVL(cllSubData("_page_id")), "��ҳID"
            cllTmp.Add NVL(cllSubData("_pati_name")), "����"
            cllTmp.Add NVL(cllSubData("_inpatient_num")), "סԺ��"
            cllTmp.Add NVL(cllSubData("_registration_time")), "��Ժʱ��"
            cllTmp.Add NVL(cllSubData("_adtd_time")), "��Ժʱ��"
            
        cllPatiList.Add cllTmp, "_" & cllTmp("����ID") & "_" & cllTmp("��ҳID")
    Next
    
    ZlCissvr_GetPatiBaseInfoByPatiId = True
    Exit Function
Errhand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
