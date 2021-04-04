Attribute VB_Name = "mdlPubComLib"
Option Explicit


Private mblnInit As Boolean                                        '���������Ƿ��ѳ�ʼ��

Public gcnOracle As New ADODB.Connection     '�������ݿ�����
Public gdtStart     As Long                  '����ʱ�䣬�����ж�������Ļ�ĵȴ�ʱ��

Public zl9ComLib As Object
Public zlDatabase As Object
Public zlCommFun As Object
Public zlControl As Object

Public Type TYPE_SYS_INFO   '-----------Ӧ�ó�����Ϣ �� ע����Ϣ
    AppName As String       'ϵͳ���� (��Ʒ���+����������������ҽҵ���)
    ShortName As String     '��Ʒ����
    AppTitle As String      'ϵͳ���⣬��Ʒȫ��
    
    Version As String       'ϵͳ�汾
    AviPath As String       'AVI�ļ�·��
    
    UnitName  As String     '�û���λ����
    Supporter As String     '����֧����
    Develop As String       '������
    SupporterWEB As String  '֧����WEB����
    SupporterMail As String '֧�����ʼ�
    SupporterURL As String  '֧������ַ
    ProductLine  As String  '��Ʒϵ�У�[��׼��],[��ͻ���]
    
    SysNo       As Long     'ϵͳ���
    ModlNo      As Long     'ģ���
End Type

Public Type TYPE_SYS_PARAMETER    'ϵͳ����
    Privs        As String  'ģ��Ȩ��
    
    MachineCount As Integer '��������
    blnEmerge    As Boolean '�Ƿ����ּ���
    BuffDir      As String   '���ػ����¼���Ļ���Ŀ¼
    InvaidWord   As String   '��ȥ���ķǳ��ַ�
    intCA        As Integer  'CA���ı��
    strMatch     As String   '����ƥ��
    
    LogLevel     As LOGTYPE  '��־��¼�ȼ� 3-����4-���� 6-��ʾ 7-����
    strDevList  As String     '���������������б�
    
    ftpSetup    As String    'FTP����
    
End Type

Public Type TYPE_USER_INFO
    ID As Long          '��ԱID
    DeptID As Long      '��Ա��Ӧ�Ĳ���ID
    DeptName As String  '��Ա��Ӧ�Ĳ�������
    No As String        '��Ա���
    Name As String      '��Ա����
    Code As String      '��Ա����
    DBUser As String    '��Ա��Ӧ�����ݿ��û���
End Type

Public UserInfo As TYPE_USER_INFO
Public gSysInfo As TYPE_SYS_INFO
Public gSysParameter As TYPE_SYS_PARAMETER

'���� ��������ComLib��һЩ������������
Public Function ComOpenSQL(ByVal strSQL As String, ByVal strTitle As String, _
    ParamArray arrInput() As Variant) As ADODB.Recordset
    '���ܣ�ͨ��ComLib����򿪴�����SQL�ļ�¼��
    
    Dim lngCount As Long
    Dim var(30) As Variant
    
    If Not mblnInit Then Exit Function
    lngCount = UBound(arrInput)
    If lngCount > 30 Then
        Err.Raise -2147483645, , "��֧�ֳ���30��������SQL��"
        Exit Function
    End If
    For lngCount = LBound(arrInput) To UBound(arrInput)
        var(lngCount) = arrInput(lngCount)
    Next
    Set ComOpenSQL = zlDatabase.OpenSQLRecord(strSQL, strTitle, var(0), var(1), var(2), var(3), var(4), var(5), var(6), var(7), var(8), var(9), var(10), var(11), var(12), var(13), var(14), var(15), var(16), var(17), var(18), var(19), var(20), var(21), var(22), var(23), var(24), var(25), var(26), var(27), var(28), var(29))

End Function

Public Function ComExecuteProc(strSQL As String, ByVal strFormCaption As String) As String
    '���ܣ�ִ�й������,���Զ��Թ��̲������а󶨱�������
    '���أ��޴��󷵻ؿմ������򷵻ش�����ʾ
    If Not mblnInit Then Exit Function
    Call zlDatabase.ExecuteProcedure(strSQL, strFormCaption)
    Exit Function
End Function

Public Function ComInitComLib(ByRef strErr As String) As Boolean
    '��ʼ����������,�ڳ�������ʱ����
    Dim strSQL As String
    On Error GoTo errH
    ComInitComLib = False
    If mblnInit Then
        ComInitComLib = True
        Exit Function
    End If
    
    Set zl9ComLib = CreateObject("zl9ComLib.clsComLib")
    zl9ComLib.InitCommon gcnOracle

    Set zlDatabase = zl9ComLib.zlDatabase
    Set zlCommFun = zl9ComLib.zlCommFun
    Set zlControl = zl9ComLib.zlControl
    
'    If zl9ComLib.RegCheck = False Then
'        strErr = "ע����Ϣ��֤δͨ����"
'        Exit Function
'    End If
    
    '�����������Ч��Ϊ�ջ�Ϊ"-"�������˳�
    gSysInfo.ShortName = zl9ComLib.zlRegInfo("��Ʒ����")
    gSysInfo.UnitName = zl9ComLib.zlRegInfo("��λ����", , -1)
    gSysInfo.AppName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gSysInfo.Version = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gSysInfo.AviPath = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gSysInfo.AppTitle = zl9ComLib.zlRegInfo("��Ʒ����")

     
    gSysInfo.SysNo = 100   'ϵͳ��
    gSysInfo.ModlNo = 1208  'ģ���
    
'    strSQL = "Zl_Createsynonyms(" & gSysInfo.SysNo & ")"
'    zlDatabase.ExecuteProcedure strSQL, "����ͬ���"

    ComInitComLib = True
    mblnInit = True
    Exit Function
errH:
    strErr = Err.Number & " " & Err.Description
 
End Function

Public Function ComGetPrivs(ByVal lngSys As Long, ByVal lngModul As Long) As String
    '��ȡģ��Ȩ��
   ComGetPrivs = zl9ComLib.GetPrivFunc(lngSys, lngModul)
End Function


Public Function ComGetSysParameter(strErr As String) As Boolean
    '��ȡϵͳ����
    On Error GoTo errH
    ComGetSysParameter = False
    gSysParameter.InvaidWord = "`#@$%&|\{}[]?;""'"
    
    gSysParameter.ftpSetup = zlDatabase.GetPara("FTP����", gSysInfo.SysNo, gSysInfo.ModlNo, "")
    ComGetSysParameter = True
    Exit Function
errH:
    strErr = Err.Number & " " & Err.Description
'    Call SaveLog("GetSysPra", LOG_ERR, Err.Number, strErr)
End Function

Public Function ComSetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
    '���ò���
    ComSetPara = zlDatabase.SetPara(varPara, strValue, lngSys, lngModual, blnSetup)
End Function

Public Function ComGetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional intType As Integer) As String
    'ȡ����
    ComGetPara = zlDatabase.GetPara(varPara, lngSys, lngModual, strDefault, arrControl, blnSetup, intType)
End Function
Public Function ComGetUserInfo(ByRef strErr As String) As Boolean
    '���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    
    zl9ComLib.SetDbUser UserInfo.DBUser
    Set rsTmp = zlDatabase.GetUserInfo

    If Not rsTmp.EOF Then
        UserInfo.ID = Val("" & rsTmp!ID)
        UserInfo.No = Trim("" & rsTmp!���)
        UserInfo.DeptID = Val("" & rsTmp!����ID)
        UserInfo.DeptName = Trim("" & rsTmp!������)
        UserInfo.Code = Trim("" & rsTmp!����)
        UserInfo.Name = Trim("" & rsTmp!����)
    
        ComGetUserInfo = True
 
    End If
            
Exit Function
errH:
    strErr = Err.Number & " " & Err.Description
'    Call SaveLog("ComGetUserInfo", LOG_ERR, Err.Number, strErr)
End Function

Public Function ComGetNextID(ByVal strTableName As String) As Long
    'ȡ������Ӧ������
    ComGetNextID = zlDatabase.GetNextId(strTableName)
End Function

Public Function ComOEMPicture(objPic As Object, strAttribute As String, Optional strProductName As String)
    'ȡOEMͼƬ
    On Error GoTo errH
    Call zl9ComLib.ApplyOEM_Picture(objPic, strAttribute, strProductName)
    Exit Function
errH:
'    Call SaveLog("OEMPicture", LOG_ERR, Err.Number, Err.Description)
End Function

Public Function ComGetLike(ByVal strTable As String, ByVal strField As String, ByVal strInput As String) As String
    '��ȡLike����
    ComGetLike = zlCommFun.GetLike(strTable, strField, strInput)
End Function

Public Function ComPressKey(bytKey As Byte)
    'ִ��PressKey����
     Call zlCommFun.PressKey(bytKey)
End Function

Public Function ComGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '��ȡ���뺯��
   ComGetSymbol = zlCommFun.zlGetSymbol(strInput, bytIsWB)
End Function

Public Function ComIncStr(ByVal strVal As String) As String
    ComIncStr = zlCommFun.IncStr(strVal)
End Function

Public Function ComErrCenter() As Byte
    '����������
    ComErrCenter = zl9ComLib.ErrCenter
End Function

Public Function ComCurrDate() As Date
    'ȡ��������ǰ����ʱ��
    ComCurrDate = zlDatabase.Currentdate
End Function







