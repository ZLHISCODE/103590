VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOwnerCardObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mvarObjectEvts As Object
Private mcnOracle As ADODB.Connection
Private mobjDataBase As clsDataBase

Public Function zlCardDevSet(ByVal frmMain As Object, ByVal lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�豸�������ýӿ�
    '���:frmMain-���õ�������
    '        lngCardTypeID-�����ID
    '����:��������    True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2011-05-23 11:51:55
    '˵��:
    '    1.  ������/�Һ��շѹ����Ĳ�������'�豸����' zlCardDevSet
    '    2.  ������ҽ��վ�����ﻤʿվ������ҽ��վ�ȵĲ�������'�豸����' zlCardDevSet
    '    3.  ������Ҫ����������ӿڵ�,ͬʱ�ִ�������豸���õ�,��Ҫ���øýӿ�.
     '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
  '  Call frmSetDrvParaSet.ShowMe(frmMain, CStr(lngCardTypeID), 1)
    zlCardDevSet = True
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlBrushCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal strPatiName As String, ByVal strSex As String, _
    ByVal strOld As String, ByVal dbl��� As Double, _
    Optional ByRef strCardNo As String, _
    Optional ByRef strPassWord As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ��֧�����,����
    '���:rsClassMoney:�շ����,���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-18 14:18:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlBrushCard = True
 End Function

Public Function zlInitComponents(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    ByVal cnOracle As ADODB.Connection, _
    Optional blnDeviceSet As Boolean = False, _
    Optional strExpand As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '���: frmMain-���õ�������
    '        lngModule-HIS����ģ���
    '       lngSys-�����ϵͳ��
    '       strDBUser-���ݿ��û���
    '       cnOracle -HIS/��������
    '       blnDeviceSet-�豸���õ��ó�ʼ��
    '       strExpand-��չ��Ϣ(����,��Ϊ��)
    '����:��������True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2011-05-23 11:49:39
    '˵��:����ǵ������ӿ�������ȡ����,�����������ĵĲ����д�������.
    '    1.��������ӿڴ���ǩ��������������ڴ˽ӿ��н���ǩ��.
    '    2.����������ӿڶ�ȡ����,��Ҫ�ȳ�ʼ��.
    '������:
    '    1.  ���������շѡ��Һ�ʱ���ñ��ӿ�
    '    2.  ����ҽ��վ����ʿվ��ҽ��վ��ʱ���ñ��ӿ�
    '    3.  �������ӿ���Ҫ������ȡ����ʱ,�����ȳ�ʼ��.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mcnOracle = cnOracle
    If zlGetOneDataBase(cnOracle, mobjDataBase) = False Then Exit Function
    zlInitComponents = True
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlReadCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal blnOlnyCardNO As Boolean, _
    ByVal strExpand As String, _
    ByRef strOutCardNO As String, _
    ByRef strOutPatiInforXML As String, _
    Optional ByRef strPhotoFile As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ӿ�
    '���:frmMain-���õĸ�����
    '       lngModule-���õ�ģ���
    '       strExpand-��չ����,������
    '       blnOlnyCardNO-������ȡ����
    '����:strOutCardNO-���صĿ���
    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '       strPhotoFile-���ӿ��ж�ȡȡ��Ƭ��,����HIS����Ƭ�ļ�(������·��)
    '����:��������    True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2011-05-23 11:56:41
    '˵��:
    '   ��. �����а󶨿�ʱ,��Ҫ����������Ϣ����ʱ����ȡ�ýӿ�
    '   ��. ��������Ҫˢ�����о���ĵط�������Ҫ���øýӿ�
    'strOutPatiInforXML����˵��
    '    ��ʶ    ��������    ����    ����    ˵��
    '    ����    Varchar2    20
    '    ����    Varchar2    100
    '    �Ա�    Varchar2    4
    '    ����    Varchar2    10
    '    ��������    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    '    �����ص�    Varchar2    50
    '    ����֤��    VARCHAR2    18
    '    ����֤��    Varchar2    20
    '    ְҵ    Varchar2    80
    '    ����    Varchar2    20
    '    ����    Varchar2    30
    '    ѧ��    Varchar2    10
    '    ����״��    Varchar2    4
    '    ����    Varchar2    30
    '    ��ͥ��ַ    Varchar2    50
    '    ��ͥ�绰    Varchar2    20
    '    ��ͥ��ַ�ʱ�    Varchar2    6
    '    �໤��  Varchar2    64
    '    ��ϵ������  Varchar2    64
    '    ��ϵ�˹�ϵ  Varchar2    30
    '    ��ϵ�˵�ַ  Varchar2    50
    '    ��ϵ�˵绰  Varchar2    20
    '    ������λ    Varchar2    100
    '    ��λ�绰    Varchar2    20
    '    ��λ�ʱ�    Varchar2    6
    '    ��λ������  Varchar2    50
    '    ��λ�ʺ�    Varchar2    20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Errhand:
  ' ����IC����
    strOutCardNO = ""
    If mobjICCard Is Nothing Then
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    If Not mobjICCard Is Nothing Then
        strOutCardNO = mobjICCard.Read_Card()
    End If
    zlReadCard = strOutCardNO <> ""
    Exit Function
Errhand:
    If mobjDataBase.ErrCenter() = 1 Then Resume
End Function
Public Function zlMakeCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal stInputPatiInforXml As String, _
    ByRef strPhotoFile As String, _
    ByRef strOutPatiInforXML As String, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ƿ��ӿ�
    '���:
    '    frmMain Object  In  ���õ�������
    '    lngModule   Long    In  ���õ�ģ���
    '    lngCardTypeID   Long    In  �����ID
    '    stInputPatiInforXml String(XML) In  HIS�����Ѿ������Ĳ�����Ϣ
    '    strPhotoFile    String  In  ��ʾ��Ƭ�ļ�(������·��)
    '    strExpend   String  In   XML,����,���Ժ���չ
    '����:
    '    strOutPatiInforXml  String(XML) Out ���ýӿں�,����HISϵͳ�Ѿ��ƿ���Ĳ�����Ϣ,�Ա㽨�����˵���
    '    strExpend   String  Out  XML,����,���Ժ���չ
    '    strPhotoFile    String  In  ��ʾ��Ƭ�ļ�(������·��)
    '���� :True:���óɹ�,False:����ʧ��
    '������:
    '    Ŀǰֻ��ҽ�ƿ����Ź����е��ô˽ӿ�(��Ҫ�ǿ����������Ϣ�϶�,�������ڲ���Ӧ)����صĵ�������ҽ�ƿ����Ź������ƿ�
    '  strInputPatiInforXML��strOutPatiInforXML��ʽ
    '    ��ʶ    ��������    ����    ����    ˵��
    '    ����    Varchar2    20
    '    ����    Varchar2    100
    '    �Ա�    Varchar2    4
    '    ����    Varchar2    10
    '    ��������    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    '    �����ص�    Varchar2    50
    '    ����֤��    VARCHAR2    18
    '    ����֤��    Varchar2    20
    '    ְҵ    Varchar2    80
    '    ����    Varchar2    20
    '    ����    Varchar2    30
    '    ѧ��    Varchar2    10
    '    ����״��    Varchar2    4
    '    ����    Varchar2    30
    '    ��ͥ��ַ    Varchar2    50
    '    ��ͥ�绰    Varchar2    20
    '    �����ʱ�    Varchar2    6
    '    �໤��  Varchar2    64
    '    ��ϵ������  Varchar2    64
    '    ��ϵ�˹�ϵ  Varchar2    30
    '    ��ϵ�˵�ַ  Varchar2    50
    '    ��ϵ�˵绰  Varchar2    20
    '    ������λ    Varchar2    100
    '    ��λ�绰    Varchar2    20
    '    ��λ�ʱ�    Varchar2    6
    '    ��λ������  Varchar2    50
    '    ��λ�ʺ�    Varchar2    20
    '    ����ID  Number  18      ����Ϊ��
    '    ABOѪ�� Varchar2    10      ��Ϣ��='ABO'
    '    ��Ϣֵ=(A��;B��;O��;AB��;����)
    '    RH  Varchar2    10      ��Ϣ��='RH'
    '    ��Ϣֵ=(��;��;����;δ��)
    '    ������־    Number  2       ��Ϣ��='ҽѧ��ʾ'
    '    ��Ϣֵ=' ����;���ಡ; ����Ѫ�ܲ�;��'
    '    ���ಡ��־  Number  2
    '    ����Ѫ�ܲ���־  Number  2
    '    ��ﲡ��־  Number  2
    '    ��Ѫ���ұ�־    Number  2
    '    ���򲡱�־  Number  2
    '    ����۱�־  Number  2
    '    ͸����־    Number  2
    '    ������ֲ��־    Number  2
    '    ����ȱʧ��־    Number  2
    '    ��װж��֫��־  Number  2
    '    ����������־  Number  2
    '    ����ҽѧ��ʾ    Varchar2    100
    '    ��ϵ��Ϣ
    '        ����    Varchar2    30
    '        ��ϵ    Varchar2    30
    '        �绰    Varchar2    30
    '        ����֤��    Varchar2    30
    '    �����������    Varchar2    18
    '    ��ũ��֤��  Varchar2    18
    '    ����֤��
    '        ��Ϣ��  Varchar2    15
    '        ��Ϣֵ  Varchar2    100
    '    ������Ϣ ���Է�������ӵ㲻���ڵ�����
    '        ��Ϣ��  Varchar2    20      ��: ������λ, Ѫѹ��
    '        ��Ϣֵ  Varchar2    50      ����:��Ϣ��=Ѫѹ;��Ϣֵ='XXmmHg'
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlMakeCard = False
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlBandCardArfter(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal lng����ID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�󶨿��ӿ�
    '    frmMain Object  In  ���õ�������
    '    lngModule   Long    In  ���õ�ģ���
    '    lngCardTypeID   Long    In  �����ID
    '    lng����ID   Long    IN   ����ID
    '    strExpend   String  In/Out  XML,����,���Ժ���չ
    ' ��������    True:���óɹ�,False:����ʧ��
    '������:
    '    1.  ҽ�ƿ����Ź����а󶨿������
    '    2.  ������Ժ�Ǽ�;���˹ҺŹ���;������Ϣ�Ǽ��н��п���ʱ����
    '����˵��:
    '    ��صĵ������̼�������ģ��İ󶨿�����ϸ˵��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '56615
    zlBandCardArfter = True
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, _
    ByVal strCardNo As String, strExpand As String, dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�ʻ����
    '���:frmMain-���õ�������
    '        lngModule-ģ���
    '        strCardNo-����
    '        strExpand-Ԥ����Ϊ��,�Ժ���չ
    '����:dblMoney-�����ʻ����
    '����:��������    True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2011-05-26 16:29:48
    '˵��:
    '       ��������Ҫ�ۿ�ĵط�����Ҫ����ʻ�����Ƿ����,�ʻ�������ʱ�������ۿ�.
    '       ���ĳЩ�������ӿڲ��������ӿڣ����Թ̶�����һ���Ľ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlGetAccountMoney = zlԤ�����(strCardTypeID, strCardNo, dblMoney)
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlMzInforWriteToCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal lng����ID As Long, _
    ByVal lngBalanceID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:д������Ϣ�ӿ�
    '    frmMain Object  In  ���õ�������
    '    lngModule   Long    In  ���õ�ģ���
    '    lngCardTypeID   Long    In  ����д�����ID:
    '           1)����ˢ�������ID
    '           2)������ʱ,��Ҫѡ��ĳ�������ID
    '    lng����ID   Long    In  ����ID
    '    lngBalanceID    Long    In  �������(ĳ�ν�������)
    '    strExpend   String  In/Out  XML,����,���Ժ���չ
    ' ��������    True:���óɹ�,False:����ʧ��
    '����ʱ��:
    '         ҽ�ƿ����.�Ƿ�д��=1�ŵ���
    '������:
    '    1.  Ŀǰ�շѹ�����ɺ�,���øýӿ�
    '    2.  �˷���ɺ�,���øýӿ�
    '˵��:
    '        ���������Ϣ��д�� , ��Ҫ���������ƵĽӿ��ڲ�д��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '56615
    zlMzInforWriteToCard = True
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlZyInforWriteToCard(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, _
    ByVal lng����ID As Long, _
    ByVal lngBalanceID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:дסԺ��Ϣ�ӿ�
    '    frmMain Object  In  ���õ�������
    '    lngModule   Long    In  ���õ�ģ���
    '    lngCardTypeID   Long    In  ����д�����ID:
    '           1)����ˢ�������ID
    '           2)������ʱ,��Ҫѡ��ĳ�������ID
    '    lng����ID   Long    In  ����ID
    '    lngBalanceID    Long    In  ����ID(���Բ�����)
    '    strExpend   String  In/Out  XML,����,���Ժ���չ
    ' ��������    True:���óɹ�,False:����ʧ��
    '����ʱ��:
    '        ҽ�ƿ����.�Ƿ�д��=1�ŵ���
    '������:
    '    1.  ��Ժʱ
    '    2.  ����ʱ
    '    3.  ��������ʱ
    '˵��:
    '       סԺ�����Ϣ��д�� , ��Ҫ���������ƵĽӿ��ڲ�д��
    '---------------------------------------------------------------------------------------------------------------------------------------------
      '56615
    zlZyInforWriteToCard = True
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlԤ�����(ByVal strCardType As String, ByVal strCardNo As String, ByRef dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݿ���,��ȡԤ�����
    '���:strCardNo-����,lngCardTypeID-�����ID
    '����:dblMoney-��������Ԥ�����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-26 17:05:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, rsTemp As ADODB.Recordset, strSql As String
    Err = 0: On Error GoTo errHandle
    If IsNumeric(strCardType) Then
        strSql = "Select ����ID From ����ҽ�ƿ���Ϣ A,ҽ�ƿ���� B Where A.�����ID=B.ID And B.�Ƿ�����=1 And B.ID=[2] And A.����=[1] "
    Else
        strSql = "Select ����ID From ����ҽ�ƿ���Ϣ A,ҽ�ƿ���� B Where A.�����ID=B.ID And B.�Ƿ�����=1 And B.����=[3] And A.����=[1] "
    End If
    Set rsTemp = mobjDataBase.OpenSQLRecord(strSql, "ͨ��ҽ�ƿ���ȡ������Ϣ", strCardNo, Val(strCardType), strCardType)
    If rsTemp.EOF Then Exit Function
    lng����ID = Val(NVL(rsTemp!����ID))
    strSql = "Select Sum(nvl(Ԥ�����,0)) as Ԥ����� From ������� Where ����=1 and nvl(����,0) in (0,1) And ����ID =[1]"
    Set rsTemp = mobjDataBase.OpenSQLRecord(strSql, "ͨ��ҽ�ƿ���ȡ�������", lng����ID)
    dblMoney = Val(NVL(rsTemp!Ԥ�����))
    zlԤ����� = True
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlThreeInterFaceMgrTool(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lng�����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ӿڹ�������
    '���:frmMain-���õ�������
    '       lngModule-���õ�ģ���
    '       �����IID
    '����:���˺�
    '����:2012-02-13 10:17:26
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
     zlThreeInterFaceMgrTool = True
 End Function


  
Public Function zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNos As String, _
    Optional ByVal strXMLExpend As String) As Boolean
      '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ��ۿ�׼��
    '���:frmMain-���õ�������
    '       lngModule-���õ�ģ���
    '       lngCardTypeID-�����ID
    '       strCardNo-����
    '       dblMoney-֧�����
    '����:�ۿ�Ϸ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2011-05-26 16:42:43
    '˵��:
    '   �ڵ��ÿۿ�ǰ�����ڴ���Oracle�������⣬ �����ٵ��ÿۿ��ǰ�� _
    '   �Ƚ������ݵĺϷ��Լ��,�Ա�������������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblPreMoney As Double
    On Error GoTo errHandle
    zlPaymentCheck = True
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlPaymentMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    ByVal strPrepayNos As String, ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
    ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ��ۿ��
    '���:frmMain-���õ�������
    '        lngModule-����ģ���
    '       strCardNo-����
    '       dblMoney-֧�����
    '       strSwapGlideNO
    '       strSwapExtendInfor -(������ʱ���������������������)
    '����:strSwapGlideNO-������ˮ��
    '       strSwapMemo-����˵��
    '        strSwapExtendInfor -���׵���չ��Ϣ
    '            ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '            ÿ����Ŀ�в��ܰ���|�ַ�
    '����:�ۿ�ɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2011-05-26 17:13:48
    '˵��:
    '   ��������Ҫ�ۿ�ĵط����øýӿ�,Ŀǰ�滮��:�շ��ң��Һ���;������ѯ��;ҽ������վ��ҩ���ȡ�
    '   һ����˵���ɹ��ۿ�󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
    '   �ڿۿ�ɹ��󣬷��ؽ�����ˮ�ź���ر�ע˵���������������������Ϣ�����Է��ڽ���˵�����Ա��˷�.
    '   ������Ԥ������,ֱ�ӷ���true,�ٿ۴����н��д���
    '---------------------------------------------------------------------------------------------------------------------------------------------
   On Error GoTo errHandle
    zlPaymentMoney = True
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zlReturncheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, ByVal dblMoney As Double, _
    ByVal strSwapNo As String, ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ����˽���ǰ�ļ��
    '���:frmMain-���õ�������
    '       lngModule-���õ�ģ���
    '       lngCardTypeID-�����ID
    '       strCardNo-����
    '       strBalanceIDs   String  In  ����֧�����漰�Ľ���ID ��ʽ:�շ�����|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    '                                   �շ�����: 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
    '       dblMoney-�˿���
    '       strSwapNo-������ˮ��(�˿�ʱ���)
    '       strSwapMemo-����˵��(�˿�ʱ����)
    '       strXMLExpend    XML IN  ��ѡ����(��չ��).��δ����
    '����:�˿�Ϸ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2011-05-26 17:24:55
    '˵��:
    '    �ڵ��ÿۿ�ǰ�����ڴ���Oracle�������⣬��ˣ��ٵ��û��˽���ǰ���Ƚ������ݵĺϷ��Լ��,
    '    �Ա�������������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlReturncheck = True
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
    ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ��ۿ���˽���
    '���:frmMain-���õ�������
    '       lngModule-���õ�ģ���
    '       lngCardTypeID-�����ID:ҽ�ƿ����.ID
    '       strCardNo-����
    '       strBalanceIDs-����֧�����漰�Ľ���ID(����ԭ����ID):
    '                           ��ʽ:�շ�����(|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
    '       dblMoney-�˿���
    '       strSwapNo-������ˮ��(�ۿ�ʱ�Ľ�����ˮ��)
    '       strSwapMemo-����˵��(�ۿ�ʱ�Ľ���˵��)
    '       strSwapExtendInfor-�����˷ѵĳ���ID��
    '                           ��ʽ:�շ�����1|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
    '����: strSwapNo-������ˮ��(�˿����ˮ��)
    '         strSwapMemo-����˵��(�˿��˵��)
    '       strSwapExtendInfor-���׵���չ��Ϣ
    '           ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n ÿ����Ŀ�в��ܰ���|�ַ�
    '����:��������    True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2011-05-26 17:31:32
    '˵��:
    '       ��������Ҫ���Ѿ����ѵĵط������˿������£����øýӿ�,Ŀǰ�滮��:�շ��ң��Һŵȡ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    zlReturnMoney = True
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Function

'---------------------------------------------------------------------------
'����ɨ�貿������
Public Property Get zlInitEvents() As Object
    '��ʼ���¼�(�����¼�)
    Set zlInitEvents = mvarObjectEvts
End Property
Public Property Set zlInitEvents(ByVal vNewValue As Object)
    '��ʼ���¼�(�����¼�)
    Set mvarObjectEvts = vNewValue
End Property
Public Sub SetEnable(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ƿ��ڶ���״̬
    '���:blnEnabled-False��ʾ���豸�����ڶ���״̬,true��ʾ���ڶ���״̬
    '����:���˺�
    '����:2012-03-24 15:12:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
   ' mfrmTimer.SetEnable blnEnabled
End Sub
 
Private Sub mfrmTimer_ShowCardNO(ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
  '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ɨ������Ľ��
    '���:strCardNo-ɨ������Ŀ���
    '����:���˺�
    '����:2012-03-24 15:21:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���������¼�����,��Ҫ�ǵ��ô�������
    Call mvarObjectEvts.zlShowCardInfor(0, strCardNo, strXmlCardInfor, strExpended, blnCancel)
End Sub
Public Function zlHealthArchiveIsSHow(frmMain As Object, _
    ByVal lngModule As Long, _
    ByRef strFunName As String, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���񽡵�����
    '    frmMain Object  In  ���õ�������
    '    lngModule   Long    In  ���õ�ģ���
    '    strFunName  String  Out ���ع�����, ����ҽ������վ�Ĳ˵�����ʾ������(�˵�λ�÷���"����"�˵���)
    '    strExpend   String  In/Out  XML,����,���Ժ���չ
    '        Boolean ��������    True:��ʾӦ�ü��뵵����ʾ�˵�,False:�����뵵���˵�
    '������:
    '    1.  ����(סԺ)ҽ������վ����,ҽ��վ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlHealthArchiveIsSHow = False
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlHealthArchivesShow(frmMain As Object, _
    ByVal lngModule As Long, ByVal lng����ID As Long, _
    Optional ByRef strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���񽡵�����
    '    frmMain Object  In  ���õ�������
    '    lngModule   Long    In  ���õ�ģ���
    '    lng����ID-����ID
    '    strExpend   String  In/Out  XML,����,���Ժ���չ
    ' ��������    ҽ������վ���������ж�,�������ƵĽӿڲ����п��Է���ֵ
    ' ������:����(סԺ)ҽ������վ�����
    '˵��:     �ýӿڷ�����:
    '    1.  һ������zl9CardSquare�����иýӿڲ�����ֵ,��ҽ��վֱ�ӵ���,�ýӿڵ����������ƵĽӿڽ���,���򰴱�׼�Ľ���������������ʾ(����ͼ)
    '    2.  �������������ƵĽӿڲ���:�ýӿ���������ֵΪTRUEʱ,��ʾ���������Լ����ƵĽ���չʾ,���򽫰�HIS�̶��Ľ���������Ϣ��ʾ(����ͼ)
    '������:
    '    1.  ����(סԺ)ҽ������վ����,ҽ��վ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlHealthArchivesShow = False
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlReleaseComponent() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ͷ���Դ�ӿ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-06-24 09:44:54
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlReleaseComponent = True
End Function

Public Function ZlGetParaConfig(ByVal frmMain As Object, _
    ByVal intPara As Integer, _
    Optional strErrMsg As String, Optional strExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�ӿڲ���
    '���: frmMain-���õ�������
    '       intPara: ��������ֵ
    '                1-ˢ����֧����ͬһҳ��:true-��ģʽ��False-��ģʽ
    '       strExpend-��չ�������������ִ�Ϊ��
    '����:strErrMsg-���صĴ�����Ϣ
    '       strExpend-��չ�������������ִ�Ϊ��
    '����:��������True:���óɹ�,False:����ʧ��
   '����:2013-06-15 20:22:51
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ZlGetParaConfig = False
End Function

Public Function ZlGetCardNo(ByRef strCardNo As String, Optional strExpend As String) As Boolean
    '��ȡˢ������
    '����:strCardNo-ˢ������
    '     strExpend-��չ�������������ִ�Ϊ��
    ZlGetCardNo = True
End Function

Private Sub Class_Initialize()
    glngInstanceCount = glngInstanceCount + 1
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mobjICCard Is Nothing Then Set mobjICCard = Nothing
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
    If mobjDataBase Is Nothing Then Set mobjDataBase = Nothing
    
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub

