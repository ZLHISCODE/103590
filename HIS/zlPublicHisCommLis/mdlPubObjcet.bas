Attribute VB_Name = "mdlPubObjcet"
'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'ģ�鹦��:��ZLLIS���õ����������������м��� ����  zlreport,
'-------------------1��zl9report ��صĳ���----------------------------------------------
'-------------------2��zl9register��صĳ���---------------------------------------------
'-------------------3��zl9LisComLib��صĳ���--------------------------------------------
'...����ͬ������
'---------------------------------------------------------------------------------------
Option Explicit

Public zlReport As Object                                           '������
Public zlRegister  As Object                                        'ע�Ჿ��zlRegister
Public gobjSample As Object                                         'LIS���������д�������ص�����


Public gobjEmr As Object                                            '��������
Public gobjEmrInterface As Object                                   '�°���Ӳ�������
Public gobjPublicLIS As Object                                      'LIS�����ӿڲ�������

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:���� ��ӡ��ӡ���ù���
'---------------------------------------------------------------------------------------
Public Function FunReportPrintSetHis(ByVal cnMain As ADODB.Connection, ByVal lngSys _
    As Long, ByVal varReport As Variant, Optional frmParent As Object) As Boolean
1         On Error GoTo ReportPrintSet_Error

2         If initReport = True Then
3            FunReportPrintSetHis = zlReport.ReportPrintSet(cnMain, lngSys, varReport, _
                 frmParent)
4         End If


5         Exit Function
ReportPrintSet_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "ִ��(ReportPrintSet)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
7         Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:�������У��򿪱�����
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function FunReportOpenHis(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, _
    ByVal varReport As Variant, frmParent As Object, ParamArray arrPar() As Variant) As Boolean
1         On Error GoTo ReportOpen_Error

          Dim lngCount As Long
          Dim var(30) As Variant
2          initReport
          
3         lngCount = UBound(arrPar)
4         If lngCount > 30 Then
5             Err.Raise -2147483645, , "��֧�ֳ���30�������ı���"
6             Exit Function
7         End If
8         For lngCount = LBound(arrPar) To UBound(arrPar) - 1
9             var(lngCount) = arrPar(lngCount)
10        Next
11        If UBound(arrPar) > 0 Then
12            var(29) = arrPar(UBound(arrPar))
13        End If
          
14        FunReportOpenHis = zlReport.ReportOpen(cnMain, lngSys, varReport, frmParent, var(0), var(1), var(2), var(3), var(4), var(5), var(6), var(7), var(8), var(9), var(10), var(11), var(12), var(13), var(14), var(15), var(16), var(17), var(18), var(19), var(20), var(21), var(22), var(23), var(24), var(25), var(26), var(27), var(28), var(29))


15        Exit Function
ReportOpen_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "ִ��(ReportOpen)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
17        Err.Clear
End Function


'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:�����Զ��屨���ߵĴ�ӡ������Ϣ
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function FunSetReportPrintSetHis( _
    ByVal cnOracle As ADODB.Connection, ByVal lngSysNo As Long, ByVal strReportCode As String, ByVal strKey As String, _
    ByVal strValue As String, Optional ByVal bytType As Byte = 1, Optional ByVal intFormat As Integer = 0) As Boolean

    If initReport = True Then
        FunSetReportPrintSetHis = zlReport.SetReportPrintSet(cnOracle, lngSysNo, strReportCode, strKey, strValue, bytType, intFormat)
    End If

End Function


'-------------------1��zl9report ��صĳ���----------------------------------------------
'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:��ʼ�� zl9report����
'---------------------------------------------------------------------------------------
Public Function initReport() As Boolean

1         On Error GoTo initReport_Error

2         If zlReport Is Nothing Then
3             Set zlReport = CreateObject("zl9Report.clsReport")
4         End If
5         initReport = True
6         Exit Function
initReport_Error:
7         Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "ִ��(initReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
8         Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:���� ��ӡ��ӡ���ù���
'---------------------------------------------------------------------------------------
Public Function FunReportPrintSet(ByVal cnMain As ADODB.Connection, ByVal lngSys _
    As Long, ByVal varReport As Variant, Optional frmParent As Object) As Boolean
1         On Error GoTo ReportPrintSet_Error

2         If initReport = True Then
3            FunReportPrintSet = zlReport.ReportPrintSet(cnMain, lngSys, varReport, _
                 frmParent)
4         End If


5         Exit Function
ReportPrintSet_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "ִ��(ReportPrintSet)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
7         Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:�������У��򿪱�����
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function FunReportOpen(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, _
    ByVal varReport As Variant, frmParent As Object, ParamArray arrPar() As Variant) As Boolean
1         On Error GoTo ReportOpen_Error

          Dim lngCount As Long
          Dim var(30) As Variant
2          initReport
          
3         lngCount = UBound(arrPar)
4         If lngCount > 30 Then
5             Err.Raise -2147483645, , "��֧�ֳ���30�������ı���"
6             Exit Function
7         End If
8         For lngCount = LBound(arrPar) To UBound(arrPar) - 1
9             var(lngCount) = arrPar(lngCount)
10        Next
11        var(29) = arrPar(UBound(arrPar))
          
12        FunReportOpen = zlReport.ReportOpen(cnMain, lngSys, varReport, frmParent, var(0), var(1), var(2), var(3), var(4), var(5), var(6), var(7), var(8), var(9), var(10), var(11), var(12), var(13), var(14), var(15), var(16), var(17), var(18), var(19), var(20), var(21), var(22), var(23), var(24), var(25), var(26), var(27), var(28), var(29))


13        Exit Function
ReportOpen_Error:
14        Call WriteErrLog("zlPublicHisCommLis", "clsPubObject", "ִ��(ReportOpen)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
15        Err.Clear
End Function


'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:�����Զ��屨���ߵĴ�ӡ������Ϣ
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function FunSetReportPrintSet( _
    ByVal cnOracle As ADODB.Connection, ByVal lngSysNo As Long, ByVal strReportCode As String, ByVal strKey As String, _
    ByVal strValue As String, Optional ByVal bytType As Byte = 1, Optional ByVal intFormat As Integer = 0) As Boolean

    If initReport = True Then
        FunSetReportPrintSet = zlReport.SetReportPrintSet(cnOracle, lngSysNo, strReportCode, strKey, strValue, bytType, intFormat)
    End If

End Function
'-------------------1��zl9report��صĳ���END----------------------------------------------
'-------------------2��zl9register��صĳ���-----------------------------------------------
'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:��ʼ��zlRegister����
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function initRegister() As Boolean
1         On Error GoTo initRegister_Error

2         If zlRegister Is Nothing Then
3             Set zlRegister = CreateObject("zlRegister.clsRegister")
4         End If
5         initRegister = True
6         Exit Function
initRegister_Error:
7         Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "ִ��(initRegister)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & "û��zlRegister������" & " �����У�" & Erl, True)
End Function


'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:��ָ�������ݿ�,����register����
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function FunGetConnection(ByVal strServer As String, ByVal strUserName As String, ByVal strPassWord As String, ByVal blnTransPassword As Boolean, _
                                 Optional ByVal bytProvider As Byte = 0, Optional ByRef strError As String = "���뷵�ش�����Ϣ", Optional ByVal blnSaveAccount As Boolean = True) As ADODB.Connection
      '���ܣ� ��ָ�������ݿ⣬��������ʵ������ADO���Ӷ���(�����10.35.10��ǰ�����룬���µ�ת�������������),��������������û��������뵽����gstrServer��gstrUserName��gstrPassword
      '������ strServer       :�������������߿���ֱ��ָ��IP:Port/SID
      '       strUserName     :�û���
      '       strPassword     :����
      '       blnTransPassword:�Ƿ��������ת��
      '       bytProvider     :�����ݿ����ӵ����ַ�ʽ,0-msODBC��ʽ,1-OraOLEDB��ʽ
      '       strError        :����ʧ�ܺ����ָ���˴˲������򷵻ش�����Ϣ��δָ��ʱֱ�ӵ�����ʾ��Ϣ��
      '       blnSaveAccount  :�����û��������롢����������ȫ�ֱ�����һ�㣬���ڵ�¼����ʱ���棬���ӿ�ReGetConnection��GetUserName��GetServerName��GetPassword��LoginValidateʹ�ã�
      '���أ� ���ݿ�򿪳ɹ������Ӷ����״̬���Է���adStateOpen(1),ʧ���򷵻�AdStateClosed(0)

1         On Error GoTo GetConnection_Error
2         If initRegister = True Then
3             If zlRegister.GetUserName = "" And blnSaveAccount = False Then
4                 blnSaveAccount = True
5             End If
6             On Error GoTo agin
7             Set FunGetConnection = zlRegister.GetConnection(strServer, strUserName, strPassWord, blnTransPassword, , strError, blnSaveAccount)
8             Exit Function
agin:
9             Err.Clear: On Error GoTo GetConnection_Error
10            Set FunGetConnection = zlRegister.GetConnection(strServer, strUserName, strPassWord, blnTransPassword, , strError)
11        End If
12        Exit Function
GetConnection_Error:
13        Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "ִ��(GetConnection)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
14        Err.Clear
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
Public Function FunReGetConnection(ByVal bytProvider As Byte, ByRef strError As String, Optional ByRef cnThis As ADODB.Connection) As ADODB.Connection
      '���ܣ����ص�¼����̨ʱ�����Ӷ��󣬻��߸���֮ǰ�򿪵����ݿ����Ӷ������»�ȡһ��OLEDB��MSODBC��ʽ�򿪵����Ӷ���
      '������bytProvider  :�����ݿ����ӵ����ַ�ʽ,0-msODBC��ʽ,1-OraOLEDB��ʽ,9-��¼����̨ʱ�����Ӷ���(��ͬ�ĻỰ)
      '      strError     :���ش�����ʧ�ܺ�Ĵ�����Ϣ
      '     cnThis       :����ò���ʱ�����ݴ򿪸����Ӷ���ʱ������ʺ���Ϣ������һ���»Ự�����Ӷ��󣬲�����ò���ʱ�����õ�¼����̨ʱ���ʺ���Ϣ����һ���»Ự�����Ӷ���
      '���أ� ���ݿ�򿪳ɹ������Ӷ����״̬���Է���adStateOpen(1),ʧ���򷵻�AdStateClosed(0)

1         On Error GoTo ReGetConnection_Error
2         If initRegister = True Then
3             On Error GoTo agin
4             Set FunReGetConnection = zlRegister.ReGetConnection(bytProvider, strError, cnThis)
5             Exit Function
agin:
6             Err.Clear: On Error GoTo ReGetConnection_Error
7             Set FunReGetConnection = zlRegister.ReGetConnection(bytProvider, strError)
8         End If
9         Exit Function
ReGetConnection_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "ִ��(ReGetConnection)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
11        Err.Clear
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
Public Function FunLoginValidate(ByVal strServer As String, ByVal strUserName As String, ByRef strPassWord As String, ByRef strError As String, _
                                 Optional lngInstance As Long) As Boolean
      '���ܣ����ݷ����������û�����������֤�û���¼�������10.35.10��ǰ�����룬���Զ����µ�ת������������룩
      '������strServer    :�������������߿���ֱ��ָ��IP:Port/SID,��������ֵ����ȡ��¼ϵͳ(����GetConnection����ʱ)ʹ�õķ�������
      '      strUserName  :�û���
      '      strPassword  :����ת���������(ָ���ĳ���ʹ���ŷ���ת����ģ�δָ�����򷵻ش�����ʾ��Ϣ)
      '      strError     :��֤ʧ��ʱ���ش�����Ϣ
      '      lngInstance  :��ǰӦ�ó���ʵ���ľ�������磺app.hInstance�������Ҫ����ת��������룬��ǰû�д����������������̶�ʱ����Ҫ���룩
      '���أ���֤��¼�Ƿ�ɹ�
1         On Error GoTo FunLoginValidate_Error
2         If initRegister = True Then
3             On Error GoTo agin
4             FunLoginValidate = zlRegister.LoginValidate(strServer, strUserName, strPassWord, strError, lngInstance)
5             Exit Function
agin:
6             Err.Clear: On Error GoTo FunLoginValidate_Error
7             FunLoginValidate = zlRegister.LoginValidate(strServer, strUserName, strPassWord, strError)
8         End If
9         Exit Function
FunLoginValidate_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "ִ��(FunLoginValidate)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
11        Err.Clear
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
Public Function FunzlRegInfo(strItem As String, Optional blnTemp As Boolean, Optional intBits As Integer, Optional ByVal cnOracle As ADODB.Connection) As String
1         On Error GoTo zlRegInfo_Error


2         If initRegister = True Then
3             On Error GoTo agin
4             FunzlRegInfo = zlRegister.zlRegInfo(strItem, blnTemp, intBits, cnOracle)
5             Exit Function
agin:
6             Err.Clear: On Error GoTo zlRegInfo_Error
7             FunzlRegInfo = zlRegister.zlRegInfo(strItem, blnTemp, intBits)
8         End If
9         Exit Function
zlRegInfo_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "ִ��(FunzlRegInfo)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
11        Err.Clear
End Function

'--------------------------------------------------
'���ܣ����ָ���Ĳ�Ʒ���л�ע����Ȩ��Ϣ
'������ strItem-ָ������Ȩ��Ŀ
'       blnTemp-�Ƿ��δ�������ʱע����Ϣ��֤
'       intBits-����ͬʱ�ж�����Ϣ�ĵ�λ���ơ���Ʒ�����̵�ָ����õڼ�����Ϣ,0-N,Ϊ-1ʱ��ʾ����";"����Ķ��
'���أ���ȷʱ����ָ������Ϣ�����󷵻�""
'--------------------------------------------------
Public Function HisZlRegInfo(strItem As String, Optional blnTemp As Boolean, Optional intBits As Integer) As String
    Static srsInfo As New ADODB.Recordset
    Static sblnTemp As Boolean
    Dim strInfo As String, aryInfo() As String
    Dim strSQL As String
    
    On Error GoTo Errhand
    If blnTemp Or sblnTemp <> blnTemp Or (srsInfo.State <> adStateOpen) Then
        sblnTemp = blnTemp
        strSQL = "Select Item,Text From Table(Cast(zltools.f_Reg_Info([1]) As zlTools.t_Reg_Rowset))"
        Set srsInfo = OpenSQLRecord(Sel_Lis_DB, strSQL, "zlRegInfo", IIf(blnTemp, 1, 0))
    End If
    
    srsInfo.Filter = "Item='" & strItem & "'"
    If srsInfo.RecordCount <> 1 Then HisZlRegInfo = "": Exit Function
    strInfo = "" & srsInfo!Text
    If (strItem = "��λ����" Or strItem = "��Ʒ������" Or strItem = "����֧����") And intBits <> -1 Then
        aryInfo = Split(strInfo, ";")
        If intBits > UBound(aryInfo) Then
            strInfo = ""
        Else
            strInfo = aryInfo(intBits)
        End If
    End If
    HisZlRegInfo = strInfo
    Exit Function
Errhand:
    HisZlRegInfo = ""
End Function

'--------------------------------------------------
'���ܣ������Ȩ������Ϣ
'���أ���2�Ĺ���ĩλ�η����ع������
'--------------------------------------------------
Public Function HisZlRegTool(Optional blnTemp As Boolean) As Long
    Dim rsTool As ADODB.Recordset
    Dim strSQL As String, lngRetu As Long

    On Error GoTo Errhand
    strSQL = "Select Prog From Table(Cast(zltools.f_Reg_Tool([1]) As zlTools.t_Reg_Rowset))"
    Set rsTool = OpenSQLRecord(Sel_Lis_DB, strSQL, "zlRegTool", IIf(blnTemp, 1, 0))
    lngRetu = 0
    Do While Not rsTool.EOF
        lngRetu = lngRetu + 2 ^ ((Val("" & rsTool.Fields(0).value) Mod 10) - 1)
        rsTool.MoveNext
    Loop
    HisZlRegTool = lngRetu
    Exit Function
Errhand:
    HisZlRegTool = 0
End Function

'--------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'���أ���ȷ����"";���󷵻ش�����Ϣ
'--------------------------------------------------
Public Function FunzlRegCheck(Optional ByVal blnTemp As Boolean, Optional ByVal cnOracle As ADODB.Connection, Optional ByVal blnInit As Boolean) As String
      '���ܣ���֤ϵͳע����Ȩ����ȷ�ԣ����ҶԵ�ǰ�Ự������֤������¼ʱ������ã�
      '������blnTemp  :�Ƿ��δ�������ʱע����Ϣ��֤��������ע���뵼�빦�ܣ�
      '      cnOracle :���ݴ�������ӽ��лỰ��֤�������Բ�����ʼ��zlRegInit�����ӽ��лỰ��֤
      '      blnInit  :�Ƿ񽫴��������cnOracle�������в�����ʼ��zlRegInit
1         On Error GoTo FunzlRegCheck_Error

2         If initRegister = True Then
3             On Error GoTo agin
4             FunzlRegCheck = zlRegister.zlRegCheck(blnTemp, cnOracle, blnInit)
5             Exit Function
agin:
6             Err.Clear: On Error GoTo FunzlRegCheck_Error
7             FunzlRegCheck = zlRegister.zlRegCheck(blnTemp)
8         End If

9         Exit Function
FunzlRegCheck_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "mdlPubObjcet", "ִ��(FunzlRegCheck)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
11        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/29
'��    ��:ͨ��ѡ�������Ӽ�¼����ѡ�����ݲ����ַ����ķ�ʽ����
'�����ϵ�ѡ�����ӿڡ��µ�ѡ���������Ҫʹ�ã���ʹ��SeletItemFromRsnew����
'��    ��:
'           objfrm              ����ѡ�����ĸ�������
'           rsTmpIn             ѡ������������Դ
'           strFind             Ĭ�Ϲ�������
'           lngID               Ĭ�Ϲ���ID�������¼���а���ID�ֶεĻ�
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function SeletItemFromRsOld(objFrm As Object, ByVal rsTmpIn As Recordset, ByVal strFind As String) As String
    SeletItemFromRsOld = frmPubDicSelOld.ShowMe(objFrm, rsTmpIn, strFind)
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/8/16
'��    ��:��ʾ����ѡ����
'��    ��:
'           objfrm      ������Դ
'           rsTmp       ��Ҫչʾ��������Դ
'           strFilter   ��Ҫ���˵�����
'           lngID       Ĭ�Ϲ���ID�������¼���а���ID�ֶεĻ�
'           intShowCol  ��Ҫչʾ���������ݣ��ӵ�0�п�ʼ����������
'           strHiddenID ��Ҫ���ص���,���IDʹ��","�ָ�
'           blnShowCheckBox     �Ƿ���ʾ��ѡ������ʾ��ѡ�����ʾ���Զ�ѡ

'��    ��:
'��    ��:  ѡ������ݣ�ÿ��֮��ʹ�á�;���ָ�
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function SeletItemFromRs(objFrm As Object, ByVal rsTmp As ADODB.Recordset, Optional ByVal strFilter As String, _
                               Optional ByVal lngID As String, Optional ByVal intShowCol As Integer = 3, _
                               Optional ByVal strHiddenID As String, Optional ByVal blnShowCheckBox As Boolean) As String
    '�ı�������������ʱ��ʹ�����뷨���һ�����Ĵ������δ���keypress�¼������η����ڸ��¼��У���ᱻ�������ûᱨ���Ӹ��жϣ���������Ѿ���ʾ��������ʾ
    If frmPubDicSel.Visible = False Then
        SeletItemFromRs = frmPubDicSel.ShowMe(objFrm, rsTmp, strFilter, lngID, intShowCol, strHiddenID, blnShowCheckBox)
    End If
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/29
'��    ��:����vsf�ؼ���ͷ����ʾ˳��,����ʾ������,���ô˹���ʱ,��������һ��������������Щ����
'��    ��:
'           VSFlexGrid                     ��������VSF
'           X                              ���������X����
'           Y                              ���������Y����
'           strPara                        ������
'           lngSysNo                       ϵͳ��
'           lngModlNo                      ģ���
'           [strHiddenCols]                �̶���Զ������ʾ����,����ID,,��Щ
'           [strShwoCols]                  �̶���Զ����ʾ����
'��    ��:
'��    ��:
'           ��������֮�����ͷ˳�� , ����Ĳ���Ҳ�������ʽ
'           ��ʽ:�е�keyֵ1,���,�Ƿ���ʾ(1=��ʾ,0=����ʾ);�е�keyֵ2,���,�Ƿ���ʾ(1=��ʾ,0=����ʾ),,,,,,,,
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function SetVsfColHiden(objFrm As Object, objVSF As Object, ByVal X As Long, ByVal Y As Long, _
                    ByVal strPara As String, ByVal lngSysNo As Long, ByVal lngModlNo As Long, _
                    Optional ByVal strHiddenCols As String, Optional ByVal strShwoCols As String) As String
                    
        SetVsfColHiden = frmPubColShow.ShowMe(objFrm, objVSF, X, Y, strPara, lngSysNo, lngModlNo, strHiddenCols, strShwoCols)
        '�ڹ��������б�����һ�Σ��ڽӿڲ����У���Ҫ�ٴα��档
        Call ComSetPara(Sel_Lis_DB, strPara, SetVsfColHiden, lngSysNo, lngModlNo)

End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/29
'��    ��:  ����ָ���ַ����ļ���
'           ����ָ���ַ������ɼ��룬���������������͵ļ���
'           0��ƴ����ȡÿ�ֵ�����ĸ���ɼ���
'           1����ʣ�ȡÿ�ֵ�����ĸ���ɼ���
'           2����ʣ�����ʹ��򹹳ɼ���
'           �ڴ���Ĳ�����δ���֡����ţ��Ͱ��û���ϵͳѡ�������õķ�ʽ���ɼ��룻
'           ����Ͱ��ڡ����ź������ָ���ķ�ʽǿ�����ɼ��룬���1��ʾ���������ĸ����
'��    ��:
'           strAsk      ������ɼ�����ַ���
'��    ��:
'��    ��:  ����
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function SpellCode(ByVal strAsk As String) As String
    SpellCode = gobjHisComLib.zlCommFun.SpellCode(strAsk)
End Function

Public Sub PressKey(bytKey As Byte)
    gobjHisComLib.zlCommFun.PressKey (bytKey)
End Sub
