VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPubObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'ģ�鹦��:��ZLLIS���õ����������������м��� ����  zlreport,
'-------------------1��zl9report ��صĳ���----------------------------------------------
'-------------------2��zl9register��صĳ���---------------------------------------------
'-------------------3��zl9LisComLib��صĳ���--------------------------------------------
'...����ͬ������
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
Option Explicit

'-------------------1��zl9report ��صĳ���----------------------------------------------
'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:������ӡ��ӡ���ù��ܣ��ṩ���ⲿ���ã��ṩ���ⲿ����
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function ReportPrintSet(ByVal cnMain As ADODB.Connection, ByVal lngSys _
    As Long, ByVal varReport As Variant, Optional frmParent As Object) As Boolean
1         On Error GoTo ReportPrintSet_Error

2            ReportPrintSet = FunReportPrintSet(cnMain, lngSys, varReport, frmParent)


3         Exit Function
ReportPrintSet_Error:
4         Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "ִ��(ReportPrintSet)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
5         Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:���������У��򿪱������ܣ��ṩ���ⲿ����
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function ReportOpen(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, _
    ByVal varReport As Variant, frmParent As Object, ParamArray arrPar() As Variant) As Boolean
1         On Error GoTo ReportOpen_Error

          Dim lngCount As Long
          Dim var(30) As Variant
          
2         lngCount = UBound(arrPar)
3         If lngCount > 30 Then
4             Err.Raise -2147483645, , "��֧�ֳ���30�������ı�����"
5             Exit Function
6         End If
7         For lngCount = LBound(arrPar) To UBound(arrPar) - 1
8             var(lngCount) = arrPar(lngCount)
9         Next
10        If UBound(arrPar) > 0 Then
11            var(29) = arrPar(UBound(arrPar))
12        End If
13        If initReport = True Then
14            ReportOpen = FunReportOpen(cnMain, lngSys, varReport, frmParent, var(0), var(1), var(2), var(3), var(4), var(5), var(6), var(7), var(8), var(9), var(10), var(11), var(12), var(13), var(14), var(15), var(16), var(17), var(18), var(19), var(20), var(21), var(22), var(23), var(24), var(25), var(26), var(27), var(28), var(29))
15        End If

16        Exit Function
ReportOpen_Error:
17        Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "ִ��(ReportOpen)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
18        Err.Clear
End Function

'��    ��:������
'����ʱ��:2018/9/27
'��    ��:�����Զ��屨�����ߵĴ�ӡ������Ϣ���ṩ���ⲿ����
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function SetReportPrintSet( _
    ByVal cnOracle As ADODB.Connection, ByVal lngSysNo As Long, ByVal strReportCode As String, ByVal strKey As String, _
    ByVal strValue As String, Optional ByVal bytType As Byte = 1, Optional ByVal intFormat As Integer = 0) As Boolean

1         On Error GoTo PubSetReportPrintSet_Error

2         SetReportPrintSet = FunSetReportPrintSet(cnOracle, lngSysNo, strReportCode, strKey, strValue, bytType, intFormat)


3         Exit Function
PubSetReportPrintSet_Error:
4         Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "ִ��(PubSetReportPrintSet)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
5         Err.Clear

End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/10/10
'��    ��:���Զ��屨�������ݿ����ӽ��г�ʼ��
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Sub InitOracle(ByVal cnMain As ADODB.Connection)

    If initReport = True Then
          zlReport.InitOracle cnMain
    End If

End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/10/10
'���ܣ����ݱ����ı�ż����Ʒ��غϷ�������
'˵�������øýӿ�ǰһ��ע���ȵ���InitOracle
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function GenReportPass(ByVal strCode As String, ByVal strName As String) As String

    If initReport = True Then
         GenReportPass = zlReport.GenReportPass(strCode, strName)
    End If
End Function


'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/10/10
'��    ��:ɾ��һ�ű���
'��    ��:nMain=���ݿ�����,strNum=�������
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function ReportDel(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal strNum As String) As Boolean

    If initReport = True Then
         ReportDel = zlReport.ReportDel(cnMain, lngSys, strNum)
    End If

End Function

'---------------------------------------------------------------------------------------
'��    ��:��С��
'����ʱ��:2018/10/31
'��    ��:�ر������Ӵ���
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    If initReport = True Then
         CloseWindows = zlReport.CloseWindows
    End If
End Function

'---------------------------------------------------------------------------------------
'��    ��:��С��
'����ʱ��:2018/10/31
'��    ��:��ʼ��ӡ
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Sub ReportTaskBegin()
    '���ܣ���ʼ��ӡ�����ڶ౨���ڵ�һ�����д�ӡ
    '˵������Ҫ���ȵ��øú���������ReportOpen���ҵ���ReportOpenʱһ��Ҫ�Դ�ӡ��ʽ
    If initReport = True Then
         zlReport.ReportTaskBegin
    End If
End Sub

'---------------------------------------------------------------------------------------
'��    ��:��С��
'����ʱ��:2018/10/31
'��    ��:������ӡ
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Sub ReportTaskEnd()
    '���ܣ�������ӡ�����ڶ౨���ڵ�һ�����д�ӡ
    '˵����������ReportOpen֮�����ŵ��øú���
    If initReport = True Then
         zlReport.ReportTaskEnd
    End If
End Sub

'-------------------zl9report ��صĳ���END----------------------------------------------
'-------------------2��zl9register��صĳ���---------------------------------------------
'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:��ָ�������ݿ�,����register����
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function GetConnection(ByVal strServer As String, ByVal strUserName As String, ByVal strPassWord As String, ByVal blnTransPassword As Boolean, _
     Optional ByVal bytProvider As Byte = 1, Optional ByRef strError As String = "���뷵�ش�����Ϣ", Optional ByVal blnSaveAccount As Boolean = True) As ADODB.Connection
'���ܣ� ��ָ�������ݿ⣬��������ʵ������ADO���Ӷ���(�����10.35.10��ǰ�����룬���µ�ת�������������),��������������û��������뵽����gstrServer��gstrUserName��gstrPassword
'������ strServer       :�������������߿���ֱ��ָ��IP:Port/SID
'       strUserName     :�û���
'       strPassword     :����
'       blnTransPassword:�Ƿ��������ת��
'       bytProvider     :�����ݿ����ӵ����ַ�ʽ,0-msODBC��ʽ,1-OraOLEDB��ʽ
'       strError        :����ʧ�ܺ����ָ���˴˲������򷵻ش�����Ϣ��δָ��ʱֱ�ӵ�����ʾ��Ϣ��
'       blnSaveAccount  :�����û��������롢����������ȫ�ֱ�����һ�㣬���ڵ�¼����ʱ���棬���ӿ�ReGetConnection��GetUserName��GetServerName��GetPassword��LoginValidateʹ�ã�
'���أ� ���ݿ�򿪳ɹ������Ӷ����״̬���Է���adStateOpen(1),ʧ���򷵻�AdStateClosed(0)

    On Error GoTo GetConnection_Error
    
    Set GetConnection = FunGetConnection(strServer, strUserName, strPassWord, blnTransPassword, bytProvider, strError, blnSaveAccount)

    Exit Function
GetConnection_Error:
    Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "ִ��(GetConnection)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
    Err.Clear
End Function


'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:���ص�¼����̨ʱ�����Ӷ���ʹ��rsgister����
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function ReGetConnection(ByVal bytProvider As Byte, ByRef strError As String, Optional ByRef cnThis As ADODB.Connection) As ADODB.Connection
'���ܣ����ص�¼����̨ʱ��ͬһ�Ự�����Ӷ��󣬻������»�ȡһ��ָ����ʽ��OLEDB��MSODBC���򿪵��»Ự�����Ӷ���
'������bytProvider  :9-��¼����̨ʱ��ͬһ�Ự�����Ӷ��󣬷���Ϊһ���»Ự�����Ӷ���0-msODBC��ʽ,1-OraOLEDB��ʽ,
'      cnThis       :����ò���ʱ�����ݴ򿪸����Ӷ���ʱ������ʺ���Ϣ������һ���»Ự�����Ӷ��󣬲�����ò���ʱ�����õ�¼����̨ʱ���ʺ���Ϣ����һ���»Ự�����Ӷ���
'      strError     :���ش�����ʧ�ܺ�Ĵ�����Ϣ,�������ؼ��֡����뷵�ش�����Ϣ�����򲻵���������ʾ
'���أ� ���ݿ�򿪳ɹ������Ӷ����״̬���Է���adStateOpen(1),ʧ���򷵻�AdStateClosed(0)
    
    On Error GoTo ReGetConnection_Error
    
    Set ReGetConnection = FunReGetConnection(bytProvider, strError, cnThis)

    Exit Function
ReGetConnection_Error:
    Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "ִ��(ReGetConnection)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
    Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:���ݷ����������û�����������֤�û���¼��ʹ��regitser����
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function LoginValidate(ByVal strServer As String, ByVal strUserName As String, ByRef strPassWord As String, ByRef strError As String, _
    Optional lngInstance As Long) As Boolean
'���ܣ����ݷ����������û�����������֤�û���¼�������10.35.10��ǰ�����룬���Զ����µ�ת������������룩
'������strServer    :�������������߿���ֱ��ָ��IP:Port/SID,��������ֵ����ȡ��¼ϵͳ(����GetConnection����ʱ)ʹ�õķ�������
'      strUserName  :�û���
'      strPassword  :����ת���������(ָ���ĳ���ʹ���ŷ���ת����ģ�δָ�����򷵻ش�����ʾ��Ϣ)
'      strError     :��֤ʧ��ʱ���ش�����Ϣ
'      lngInstance  :��ǰӦ�ó���ʵ���ľ�������磺app.hInstance�������Ҫ����ת��������룬��ǰû�д����������������̶�ʱ����Ҫ���룩
'���أ���֤��¼�Ƿ�ɹ�
    On Error GoTo PubLoginValidate_Error

    LoginValidate = FunLoginValidate(strServer, strUserName, strPassWord, strError, lngInstance)

    Exit Function
PubLoginValidate_Error:
    Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "ִ��(PubLoginValidate)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
    Err.Clear
End Function

'--------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'���ܣ����ָ���Ĳ�Ʒ���л�ע����Ȩ��Ϣ
'������ strItem-ָ������Ȩ��Ŀ
'       blnTemp-�Ƿ��δ�������ʱע����Ϣ��֤
'       intBits-����ͬʱ�ж�����Ϣ�ĵ�λ���ơ���Ʒ�����̵�ָ����õڼ�����Ϣ,0-N,Ϊ-1ʱ��ʾ����";"����Ķ��
'       cnOracle:�ô������������ѯ
'���أ���ȷʱ����ָ������Ϣ�����󷵻�""
'--------------------------------------------------
Public Function zlRegInfo(strItem As String, Optional blnTemp As Boolean, Optional intBits As Integer, Optional ByVal cnOracle As ADODB.Connection) As String
1         On Error GoTo zlRegInfo_Error

2         zlRegInfo = FunzlRegInfo(strItem, blnTemp, intBits, cnOracle)
3         Exit Function

zlRegInfo_Error:
4         Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "ִ��(zlRegInfo)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
5         Err.Clear
End Function


'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/10/31
'��    ��:��ʼ��ע�Ჿ���ڵ����Ӷ���(һ�����ڵ�¼����ã�������ȡһ�����Ӷ���ʱ���س�ʼ��)
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Sub zlRegInit(cnMain As ADODB.Connection)
    If initRegister = True Then
        zlRegister.zlRegInit cnMain
    End If
End Sub

'--------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'���ܣ���֤ϵͳע����Ȩ����ȷ��
'������blnTemp-�Ƿ��δ�������ʱע����Ϣ��֤
'���أ���ȷ����"";���󷵻ش�����Ϣ
'--------------------------------------------------
Public Function zlRegCheck(Optional ByVal blnTemp As Boolean, Optional ByVal cnOracle As ADODB.Connection, Optional ByVal blnInit As Boolean) As String
'���ܣ���֤ϵͳע����Ȩ����ȷ�ԣ����ҶԵ�ǰ�Ự������֤������¼ʱ������ã�
'������blnTemp  :�Ƿ��δ�������ʱע����Ϣ��֤��������ע���뵼�빦�ܣ�
'      cnOracle :���ݴ�������ӽ��лỰ��֤�������Բ�����ʼ��zlRegInit�����ӽ��лỰ��֤
'      blnInit  :�Ƿ񽫴��������cnOracle�������в�����ʼ��zlRegInit
'���أ���ȷ����"";���󷵻ش�����Ϣ
    
   zlRegCheck = FunzlRegCheck(blnTemp, cnOracle, blnInit)
End Function

'--------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'���ܣ������Ȩ������Ϣ
'���أ���2�Ĺ���ĩλ�η����ع�������
'--------------------------------------------------
Public Function ZlRegTool(Optional blnTemp As Boolean, Optional ByVal cnOracle As ADODB.Connection) As Long
1         On Error GoTo ZlRegTool_Error

2         If initRegister = True Then
3             On Error GoTo agin
4             ZlRegTool = zlRegister.ZlRegTool(blnTemp, cnOracle)
5             Exit Function
agin:
6             Err.Clear: On Error GoTo ZlRegTool_Error
7             ZlRegTool = zlRegister.ZlRegTool(blnTemp)
8         End If


9         Exit Function
ZlRegTool_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "ִ��(ZlRegTool)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
11        Err.Clear
End Function

'--------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'���ܣ����ص�ǰ�û����е�ָ������Ĺ��ܴ�
'������lngSys     ����ǹ̶�ģ�飬��Ϊ0
'      lngProgId  �������
'���أ��ֺż���Ĺ��ܴ�,Ϊ�ձ�ʾû��Ȩ��
'--------------------------------------------------
Public Function GetPrivFunc(lngSys As Long, lngProgId As Long, Optional ByVal cnOracle As ADODB.Connection) As String
1         On Error GoTo GetPrivFunc_Error

2         If initRegister = True Then
3             On Error GoTo agin
4             GetPrivFunc = zlRegister.zlRegFunc(lngSys, lngProgId, cnOracle)
5             Exit Function
agin:
6             Err.Clear: On Error GoTo GetPrivFunc_Error
7             GetPrivFunc = zlRegister.zlRegFunc(lngSys, lngProgId)
8         End If


9         Exit Function
GetPrivFunc_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "ִ��(GetPrivFunc)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
11        Err.Clear
End Function

Public Function GetUserName(Optional ByVal blnUcase As Boolean = True) As String
'���ܣ���ȡ��ǰ��¼���û���
'������blnUcase     :�Ƿ�ת��Ϊ��д��ĸ
       If initRegister = True Then
        GetUserName = zlRegister.GetUserName(blnUcase)
    End If
    
End Function


Public Function GetServerName(Optional ByVal blnUcase As Boolean = True) As String
'���ܣ���ȡ��ǰ��¼�ķ�������
'������blnUcase     :�Ƿ�ת��Ϊ��д��ĸ
       If initRegister = True Then
        GetServerName = zlRegister.GetServerName(blnUcase)
    End If
    
End Function

Public Function GetPassword(Optional lngInstance As Long) As String
'���ܣ���ȡ��ǰ��¼(����GetConnection֮��)�û����ܺ������
'������lngInstance-��ǰӦ�ó���ʵ���ľ�������磺app.hInstance��û�д����������������̶�ʱ����Ҫ���룩
'˵����������ָ���ĳ���ʹ��ڷ��ʱ��ӿ�(δ����������£�������ʾ���֣����磺δ��Ȩ�ĳ���δ��Ȩ�Ĵ���)
       If initRegister = True Then
        GetPassword = zlRegister.GetPassword(lngInstance)
    End If
    
End Function

'-------------------zl9register��صĳ���END---------------------------------------------
