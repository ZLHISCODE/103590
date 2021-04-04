VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'##############################################################################
'#���ܣ����ù���չʾ��һ������Ĵ���
'�������򣺺�����������Ϊͨ���׶���Ӣ�ļ���
'�ϼ��ࣺclsDatabase
'##############################################################################

Public Function VerifyPassWord(frmParent As Object, ByVal strPass As String, Optional ByVal strName As String, Optional ByVal strSex As String, Optional ByVal strOld As String, Optional blnPassEncode As Boolean = True) As Boolean
'���ܣ������������֤
'������frmParent=��ʾ�ĸ�����
'      strPass=��ȷ������
'      strName,strSex,strOld=��ѡ�����������������Ա����䣬��������ʱ����ʾ�������
'      blnPassEncode-strPass�Ƿ���ļ��ܴ�
'���أ�True=������֤ͨ��,False=ȡ�����룬������3��������������
    VerifyPassWord = gobjComLib.zlCommFun.VerifyPassWord(frmParent, strPass, strName, strSex, strOld, blnPassEncode)
End Function

Public Sub StopFlash()
    '���ܣ� ֹͣ���رյȴ��Ķ�̬����
    '������
    '���أ�
    Call gobjComLib.zlCommFun.StopFlash
End Sub

Public Sub ShowFlash(Optional strNote As String, Optional frmParent As Object)
    '���ܣ� ��ʾ�ȴ��Ķ�̬����
    '������
    '   strNote:��ʾ��Ϣ
    '   frmParent�����ڴ���ĸ�����
    '���أ�
    Call gobjComLib.zlCommFun.ShowFlash(strNote, frmParent)
End Sub

Public Function ShowMsgBox(ByVal strCaption As String, ByVal strInfo As String, ByVal strCmds As String, frmParent As Object, Optional vStyle As VbMsgBoxStyle = vbQuestion, Optional ByVal strDateCaption As String, Optional ByRef DateInput As Date, Optional ByVal strDateFormat As String, Optional ByVal strSelectCaption As String, Optional ByRef strSelectInput As String, Optional ByVal strTextCaption As String, Optional ByVal lngTextLength As Long, Optional ByRef strTextInput As String, Optional ByVal strSort As String = "1,2,3", Optional ByVal blnSelectMust As Boolean) As String
'������strCaption=��Ϣ�������
'      strInfo=������ʾ����,����"^"��ʾ����,">"��ʾ������
'      strCmds=��ť����,��"����(&R),!����(&A),?ȡ��(&C)"
'              ����Ҫ��������ť,"!"��ʾȱʡ��ť,"?"��ʾȡ����ť
'              ÿ����ť�������֧��4������
'      vStyle=vbInformation,vbQuestion,vbExclamation,vbCritical
'      strDateCaption=��������ڵı��⣬���<>""����ʾ���ڿؼ������û��������ڣ���DateInput���ء�
'      strDateFormat=ʱ���ʽ ��ʽ""yyyy-MM-dd hh:mm:ss"����HHΪ��д��24Сʱ��"
'      strSelectCaption=ѡ��ı���:ѡ��1|1(1��Ϊȱʡ),ѡ��2|0|1(ѡ��ѡ��ʱ��������д��1�������ڣ�2�����ı���0��������)������
'      strSelectInput=ѡ��ѡ��ķ���ֵ(����ѡ�������)
'      strTextCaption=�ı������
'      lngTextLength=�ı������¼�볤��
'      strTextInput=�ı���ķ���ֵ
'      strSort=���ڡ�ѡ��ı������������=1��ѡ��=2���ı�=3��Ĭ������"1,2,3"
'      blnSelectMust=����е�ѡ������ѡ��һ����������ʾ��
'���أ���ť����,��"��ť2"(������()��&),������رջ�ȡ���򷵻�""
    ShowMsgBox = gobjComLib.zlCommFun.ShowMsgBox(strCaption, strInfo, strCmds, frmParent, vStyle, strDateCaption, DateInput, strDateFormat, strSelectCaption, strSelectInput, strTextCaption, lngTextLength, strTextInput, strSort, blnSelectMust)
End Function

Public Sub ShowTipInfo(ByVal lnghwnd As Long, ByVal strInfo As String, Optional blnMultiRow As Boolean, Optional blnOutline As Boolean, Optional lngMaxWidth As Long)
'���ܣ���ʾ����������ʾ
'������lngHwnd=��ʾ����ԵĿؼ����,������Ϊ0ʱ������ʾ
'      strInfo=��ʾ��Ϣ,������Ϊ��ʱ������ʾ
'      blnMultiRow=��һ���ļ�������ʾ������Ϣ��ÿ�а�vbcrlf�ָ�
'      blnOutline=�Ƿ�ÿ���ı����ַ�|ǰ��������Ϊ��ٵ���һ����ʾ
'      lngMaxWidth=���ڵ���󴰶ȣ�ȱʡΪ0��ʾ�����״̬�Ĵ���������Ϊ׼
    Call gobjComLib.zlCommFun.ShowTipInfo(lnghwnd, strInfo, blnMultiRow, blnOutline, lngMaxWidth)
End Sub

Public Sub DeviceSetup(frmParent As Object, ByVal lngSys As Long, ByVal lngModule As Long)
'����:���ݲ�ͬ��ϵͳ��ģ��,���ò�ͬ���豸����
    Call gobjComLib.zlCommFun.DeviceSetup(frmParent, lngSys, lngModule)
End Sub

Public Function PatiIdentify(frmParent As Object, ByVal lngSys As Long, _
    ByVal lng����ID As Long, ByVal cur��� As Currency, _
    Optional lngModul As Long = 0, _
    Optional bytOperationType As Byte = 0, _
    Optional lngDefaultCardTypeID As Long = 0, _
    Optional blnCheckPassWord As Boolean = True, _
    Optional blnFamilyMoney As Boolean, _
    Optional strFamilyPatiIDs As String = "", _
    Optional blnˢ����֤ As Boolean = True, _
    Optional bln�����벻�鿨 As Boolean) As Boolean
    '����:��ָ�����˽���ˢ��������֤
    '���:lngSys=���ó����ϵͳ���
    '       lng����ID=Ҫ��֤���ݵĲ���ID
    '       cur���=�������ѽ��
    '       lngModul-ģ���
    '       bytOperationType-ҵ������(0-��������ʾ;1-ֻ��ʾ�������;2-ֻ��ʾסԺ���)
    '       lngDefaultCardTypeID-ȱʡ��ˢ�����ID
    '       blnCheckPassWord-�Ƿ���֤����(true��ʾ��Ҫ��֤����,False��ʾ����֤����,ֻ��֤��)
    '        47216
    '       blnFamilyMoney-�Ƿ��ȡ����Ԥ�����
    '       strFamilyPatiIDs-���˼����Ĳ���ID��79868
    '       blnˢ����֤-�Ƿ����ˢ����֤����Ҫ���ڲ�ˢ����֤ʱ��ȡ����IDs
    '       bln�����벻�鿨-���˵�����ҽ�ƿ���û����������ʱ�Ƿ��鿨����ΪTrueʱ��ֻҪ��һ�ſ����������붼Ҫ�����鿨,112418
    '����:��֤�ɹ��ܷ���True��ʧ�ܻ�ȡ������False
    'ע�⣺1.ֻ�����ڱ�׼��
    '      2.���øú����ĵط�Ҫ����"������Ϣ,�������,����ҽ�ƿ���Ϣ"�Ĳ�ѯȨ��
    PatiIdentify = gobjComLib.zlDatabase.PatiIdentify(frmParent, lngSys, lng����ID, cur���, lngModul, bytOperationType, _
                                                    lngDefaultCardTypeID, blnCheckPassWord, blnFamilyMoney, strFamilyPatiIDs, _
                                                    blnˢ����֤, bln�����벻�鿨)
End Function

Public Function UserIdentifyByUser(frmParent As Object, ByVal strNote As String, ByVal lngSys As Long, ByVal lngProgId As Long, ByVal strFunc As String, Optional cnNew As ADODB.Connection, Optional ByVal blnDefaultPreUser As Boolean, Optional ByRef strDBUser As String) As String
'������strNote=��ʾ��Ϣ(���)
'      lngProgID=�������
'      strFunc=��Ȩ����
'      cnNew=Ҫ���ص�����,��Ҫ����ʱ,���봫���Nothing�Ķ���,������Ҫ�ɵ��ó���ر����ӣ�����ǵ�ǰ��¼�û�,����Nothing
'      blnDefaultPreUser-ȱʡ��ʾ�ϴε�¼��
'���أ��ɹ�������Ա����
'      strDBUser=�����������������ݿ��û�
    UserIdentifyByUser = gobjComLib.zlDatabase.UserIdentifyByUser(frmParent, strNote, lngSys, lngProgId, strFunc, cnNew, blnDefaultPreUser, strDBUser)
End Function

Public Function UserIdentify(frmParent As Object, ByVal strNote As String, ByVal lngSys As Long, ByVal lngProgId As Long, ByVal strUser As String, Optional cnNew As ADODB.Connection, Optional ByVal blnDBUser As Boolean, Optional ByRef strDBUser As String, Optional cnNewOLEDB As ADODB.Connection) As String
'������strNote=��ʾ��Ϣ(���)
'      lngProgID=�������
'      strUser=ָ�����û�
'      cnNew=Ҫ���ص�����,��Ҫ����ʱ,���봫���Nothing�Ķ���,������Ҫ�ɵ��ó���ر����ӣ�����ǵ�ǰ��¼�û�,����Nothing
'      blnDBUser=�����ݿ��û�ֱ����֤��¼�������ظ��û����������ӣ���ʱ����lngProgId��strFunc������ֵ��Ч
'      cnNewOLEDB=��Ҫ��ȡ��OLEDB���ӣ���CNNEW����ͬһ�û����������Ӳ�ͬ�����ò�������Nothingʱ���ŷ��� OLEDB���ӣ����򲻷���
'���أ��ɹ�������Ա����
'      strDBUser=�����������������ݿ��û�
    UserIdentify = gobjComLib.zlDatabase.UserIdentify(frmParent, strNote, lngSys, lngProgId, strUser, cnNew, blnDBUser, strDBUser, cnNewOLEDB)
End Function

Public Function ShowPatiColorTip(frmParent As Object) As Boolean
'����:��frmParent�������½���ʾһ���壬����Ϊ���ֲ������͵���ɫ˵��
'˵����������С��������͡���ѯȨ��
    ShowPatiColorTip = gobjComLib.zlDatabase.ShowPatiColorTip(frmParent)
End Function

Public Function ShowRecSelect(ByVal frmMain As Object, ByVal objControl As Object, ByVal rsBindings As ADODB.Recordset, Optional ByRef rsOutSel As ADODB.Recordset, Optional ByVal blnShowHead As Boolean = False, _
                                                    Optional ByVal strHideCols As String = "", Optional ByVal lngSys As Long, _
                                                   Optional ByVal lngModule As Long, Optional ByVal str������ As String = "", Optional ByVal strColAlignment As String = "") As Boolean
    '����:ѡ�������
    '���:frmMain-���õ�������
    '     lngSys-ϵͳ��
    '     lngModule-ģ���
    '     objControl-�ؼ�����(Ŀǰֻ֧:textBox,Combox,VSFlexGrid,BILLEDIT)
    '     rsBindings-�󶨵ļ�¼��(����Ϊ��,��Ҫ�ֶ�,ID,......)(ID���ٱ��룩
    '     blnShowHead-�Ƿ���ʾ����ͷ
    '     str����-���Ի�����Ĳ�����
    '     strColAlignment-��Ҫ�������õĸ����������Լ������ģʽ
    '         ��ʽΪ������1|0,����2|1,����3|2,...����0��1��2...8��ʾ�Ķ��뷽ʽ�ֱ�Ϊ�����ϡ����С����¡����ϡ����С����¡����ϡ����С�����
    '����:rsOutSel-ѡ���ļ�¼��
    '����:ѡ�з���True, ���򷵻�False(���԰�Esc���з���)
    Dim frmNew As New frmListSel
    ShowRecSelect = frmNew.ShowSelect(frmMain, objControl, rsBindings, rsOutSel, blnShowHead, strHideCols, lngSys, lngModule, str������, strColAlignment)
End Function

Public Function ShowSQLSelectEx(frmParent As Object, ByVal objControl As Object, ByVal strSQL As String, bytStyle As Byte, _
                                                ByVal strTitle As String, ByVal blnĩ�� As Boolean, ByVal strSeek As String, ByVal strNote As String, _
                                                ByVal blnShowSub As Boolean, ByVal blnShowRoot As Boolean, ByVal blnNoneWin As Boolean, ByRef Cancel As Boolean, _
                                                ByVal blnMultiOne As Boolean, ByVal blnSearch As Boolean, ByVal blnMulti As Boolean, _
                                                ByVal strOtherInfo As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'���ܣ��๦��ѡ������ʹ��ADO.Command��,����ʹ��[x]����
'           (�����ǰ��ShowSelect��ShowSQLMultiSelect��ShowSQLSelect��
'           �����������󣬾Ϳ����Զ���λ
'������
'     frmParent=��ʾ�ĸ�����
'     objControl=���ý��������
'     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
'     bytStyle=ѡ�������
'       Ϊ0ʱ:�б����:ID,��
'       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
'       Ϊ2ʱ:˫�����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
'             ˫���������������ں�Check��β���ֶΣ�����ֶ���Ϊ�Ƿ�ѡ��ֵ�洢�ֶΡ�=1Ϊ��ѡ��0=����ѡ��
'             ˫����������������*���ƣ�*���룬*����ģ�����ʾ���ϽǵĲ�ѯ���ܣ��Թ���ѯ��Ŀ��
'                    �����б�������ƥ�䣬ƥ��ɹ���λ���÷���ĸ���Ŀ�ϣ���F3֧�ֲ�����һ����
'     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
'     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
'     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
'             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
'             bytStyle=1ʱ,�����Ǳ��������
'     strNote=ѡ������˵������
'     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
'     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
'     blnNoneWin=�����ɷǴ�����
'     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
'     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
'     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
'     blnMulti=�Ƿ�������ѡ
'     strOtherInfo=��ʽΪ����Ŀ����1=����1#��Ŀ2=����2#......
'                ��ǰ��Ŀ�У�bytSize=0,1;�����С(0-С����,1-������;С����Ϊ9����,������Ϊ12����),Ĭ��С����
'                            ColSet=�п�����|����1,����1,0;����2,����2,1;.....|������ʾ|������ ���п��Ⱥ����һ��������ʾ���еĶ��뷽ʽ,0��1��2�ֱ��ʾ����롢�Ҷ�����м����
'                            NotShowNon=0,1;0-Ĭ�ϴ�������ʾû������ķ��࣬1-����ʾû������ķ���;bytStyle=2������
'                            HeadCap=SQL����1,�б�չʾ����1;SQL����2,�б�չʾ����2������Ŀ�����ֹ�ָ��SQL�����б���չʾ���ƣ�һ�����ڱ��������У����ǲ��ı��е�Key
'                            MultiCheckReturn=0,1����ѡʱֻ���ع�ѡ�У����ڶ�ѡ��ȷ��Ĭ�Ϸ��ص�ǰ���������Ӹò������ƣ��ÿ������ú󣬲�֧��Ĭ���еķ��أ������Ծ�֧��˫�����Զ����ء�
'                            HideNullCols=0,1;�Ƿ�����SQl�е�null as д������
'     arrInput=��Ӧ�ĸ���SQL����ֵ,��˳����,����Ϊ��ȷ����
'���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
'˵����
'     1.ID���ϼ�ID����Ϊ�ַ�������
'     2.ĩ�����ֶβ�Ҫ����ֵ
'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б��ȡ�
    Dim frmNew As New frmPubSel
    Dim arrPar() As Variant
    arrPar = arrInput
    Set ShowSQLSelectEx = frmNew.ShowSelectV2(frmParent, objControl, strSQL, bytStyle, strTitle, blnĩ��, strSeek, strNote, _
                                                 blnShowSub, blnShowRoot, blnNoneWin, Cancel, blnMultiOne, blnSearch, blnMulti, _
                                                 strOtherInfo, arrPar)
End Function

Public Function ShowILLSelect(frmParent As Object, ByVal str��� As String, ByVal lng���˿���ID As Long, Optional ByVal str�Ա� As String, Optional ByVal blnMultiSel As Boolean, Optional ByVal blnICD10 As Boolean = True, Optional ByVal strSel As String, Optional ByVal lngSys As Long = 100) As ADODB.Recordset
'���ܣ�ICD10��������/�������ѡ����
'������str���=��������������,����"D,M",Ϊ��ʱ��ʾ����
'         D:ICD-10��������
'         Y:�����ж����ⲿԭ��
'         M:������̬ѧ����
'         S:ICD-9-CM3��������
'         B:��ҽ��������
'         Z:��ҽ��������
'         1:�������-��ҽ
'         2:�������-��ҽ
'      str�Ա�=�����Ա�,���ڴ����������Ա�����
'      blnMultiSel=�Ƿ��������и�ѡ
'      blnICD10=�Ƿ���ʾ��������ѡ����(str���=D,Y,M,S,B,Z)��������ʾ�������ѡ����(str���=1,2)
'      strSel:�Ѿ�ѡ��ı��룬���Զ��ŷָ
'      lngSys=���õ�ϵͳ��ϵͳ��

'���أ�����ѡ�񼲲���Ŀ�ļ�¼����
'      �ֶΣ���ĿID,����,���,����,����ID,��������,����,˵��,����,����id,����,��Ч����,����,�Ƿ���,����ID,���ID
'      ��׼����ã�
'               ICD-10��������ѡ��������ĿID,����,���,����,����,˵��,���ID(�����ֶ�Ϊ�ջ򲻶�ȡ)
'               �������ѡ��������ĿID,����,����,˵��,����,����ID(�����ֶ�Ϊ�ջ򲻶�ȡ��
'      ����ϵͳ���ã���ĿID,����,���,����,����ID,��������,����,˵��,����,��Ч����,����,�Ƿ���(�����ֶ�Ϊ�ջ򲻶�ȡ)
'      ���ȡ��ѡ�񣬷���ΪNothing
'ע�⣺��׼����������øú���ʱ���������Ȩ��
'               zl_�����������_Insert,Zl_�����������_Delete
'               ��������Ŀ¼,�����������,�����������,�����������
'               Zl_������Ͽ���_Insert,Zl_������Ͽ���_Delete
'               �������Ŀ¼,������Ϸ���,�����������
'               ���ű�,������Ա,��������˵��,�ϻ���Ա��
'               ������Ͽ��ҡ�������϶���
'       ����ϵͳ���ã�zl_�����������_Insert,Zl_�����������_Delete
'               ��������Ŀ¼,�����������,�����������,�����������
    Set ShowILLSelect = gobjComLib.zlDatabase.ShowILLSelect(frmParent, str���, lng���˿���ID, str�Ա�, blnMultiSel, blnICD10, strSel, lngSys)
End Function
