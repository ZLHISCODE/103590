VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const gstrSysName As String = "��������"

Public mobjRegister As Object               'ע����Ȩ����zlRegister
Private mobjEMR  As Object
Private mobjRichEMR As Object
Private mobjPublicPACS As Object

Private mstrPrintDocIDs As String
Private mstrUser As String
Private mstrTipInfo As String

Private WithEvents mclsDockAduits   As zlRichEPR.clsDockAduits
Attribute mclsDockAduits.VB_VarHelpID = -1
'��ҳ��Ϣ;ҽ����¼;סԺ����;������¼;��������;���Ʊ���;����֤��;֪���ļ�;�ٴ�·��

Private Function UpdateRS(ByVal rsMed As ADODB.Recordset, ByVal strXML As String) As ADODB.Recordset
'����:�Զ����¼��,������װ�׶���Ϣ
'����:
'strXML��ʽ����
'<items>
'    <item>
'        <id></id>
'        <file_path></file_path>
'    </item>
'</items>
    Dim xmlDoc As New DOMDocument
    Dim xNode As IXMLDOMNode
    Dim xNodeList As IXMLDOMNodeList
    On Error GoTo errH
 

    '��ȡ������Ӧ���ݣ�XML��ʽ��
    xmlDoc.loadXML (strXML)
    Set xNodeList = xmlDoc.selectNodes(".//item")
    For Each xNode In xNodeList
        rsMed.Filter = "ID='" & xNode.selectSingleNode(".//id").Text & "'"
        If Not rsMed.EOF Then
            On Error Resume Next
            rsMed!�ļ�·�� = xNode.selectSingleNode(".//file_path").Text
            On Error GoTo 0
            rsMed!��� = 1
        End If
    Next
    rsMed.Filter = "���=1"
    Set UpdateRS = rsMed
    Exit Function
errH:
  MsgBox Err.Description & vbCrLf & "MakeRS" & "�� " & Erl(), vbExclamation + vbOKOnly, gstrSysName
    
End Function

Private Function InitEMR(ByVal strUser As String, ByVal strPwd As String, ByRef strError As String) As Boolean
    Call UpdateEmrInterface(strUser, strPwd)
    If Not mobjEMR Is Nothing Then
        If Not mobjEMR.IsInited Or mobjEMR.IsOffline Then
            Set mobjEMR = Nothing
        Else
            On Error Resume Next
            Set mobjRichEMR = CreateObject("zlRichEMR.clsDockContent")
            If Err <> 0 Then
                strError = "�°没���������ʧ�ܣ�����ϵ����Ա����Ƿ���ȷ��װ!"
                Set mobjEMR = Nothing
            End If
            Err.Clear
            If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.Init(mobjEMR, gcnOracle, glngSys, 0)
        End If
    End If
    InitEMR = (Not mobjEMR Is Nothing)
End Function

Public Function InitPrint(ByVal strServer As String, ByVal strUser As String, ByVal strPass As String, ByVal strDBPwd As String, Optional ByRef strError As String) As Boolean
'����:��ʼ�����Ӷ��󡢹��ö���
'����:
'strServer-������
'strUser-�û���
'strPass-��������(����̨¼��)
'strDBPwd-���ݿ�����(�°没����Ҫ)
'���ز���:
'strError-������ʾ��Ϣ
    If mobjRegister Is Nothing Then
        Set mobjRegister = GetObject("", "zlRegister.clsRegister")
        Err.Clear
        If mobjRegister Is Nothing Then
            Set mobjRegister = CreateObject("zlRegister.clsRegister")
            Err.Clear
            If mobjRegister Is Nothing Then
                MsgBox "����zlRegister��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    On Error GoTo errHand
    glngSys = 100
    mstrUser = strUser
    If gcnOracle.State = adStateOpen Then gcnOracle.Close
    Set gcnOracle = mobjRegister.GetConnection(strServer, strUser, strPass, True, , strError)

    If strError <> "" Then MsgBox strError, vbInformation
    If gcnOracle.State = adStateClosed Then Exit Function
    Set gobjComlib = CreateObject("zl9ComLib.clsComLib")
    gobjComlib.InitCommon gcnOracle
    Set gobjDatabase = CreateObject("zl9ComLib.clsDatabase")
    Set gobjCommFun = CreateObject("zl9ComLib.clsCommFun")
    Set gobjReport = CreateObject("zl9Report.clsReport")
    gobjReport.InitOracle gcnOracle
    Err.Clear
    Call gobjCommFun.PDFInitialize(strError)
    If Err.Number <> 0 Or strError <> "" Then
        MsgBox "PDF�豸��ʼ��ʧ��:" & strError, vbExclamation, gstrSysName: Exit Function
    End If
    
    If mobjPublicPACS Is Nothing Then
        On Error Resume Next
        Set mobjPublicPACS = CreateObject("zlPublicPACS.clsPublicPACS")
        Err.Clear: On Error GoTo 0
        If Not mobjPublicPACS Is Nothing Then
            Call mobjPublicPACS.InitInterface(gcnOracle, strUser)
        End If
        If mobjPublicPACS Is Nothing Then
            strError = "PACS��������δ�����ɹ���"
        End If
    End If
    Call InitEMR(strUser, strDBPwd, strError)
    InitPrint = True
    Exit Function
errHand:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function PrintDocument(ByVal lngPatiID As Long, ByVal lngVisitID As Long, ByVal strFilePath As String, Optional ByVal strXML As String, Optional ByVal blnMerge As Boolean, Optional ByRef strNoPDF As String) As Boolean
'����:���ָ������ȫ����ָ���ĵ�
'   lngPatiID-����ID
'   lngVisitID-��ҳID
'   strFilePath-PDF�ļ����·��()
'   strXML-�û�ָ��������ļ��嵥,����ʽΪXML���͵��ַ���
'   һ���ĵ���Ӧһ��<Item>�ڵ� ʾ������:
'   <items>
'       <item>
'           <id></id>           --��ͨ�� GetPrintList���id
'           <file_path></file_path>  --�ýڵ�ָ������ļ��ľ���·��(��:D:\TEST\��ҳ.PDF)��û�иýڵ�ʱ��Ĭ�Ϸ�ʽ�����ļ�����
'       </item>
'   </items>
'   blnMerge- T-������ĵ��ϲ�Ϊһ���ĵ�; F-���ϲ��ĵ�
'   strNoPDF-����δ����ɹ���PDF�ĵ���
'
    Dim rsPati As ADODB.Recordset
    Dim lngDeptID As Long
    Dim objKernel As Object
    Dim strDeviceName As String
    Dim strPrint As String
    Dim strFileName As String
    Dim strPath As String
    Dim strType As String
    Dim blnDataMove As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim objRichEPR As Object
    Dim objTendsNew As Object
    Dim objPath As Object
    Dim objDockPath As Object
    Dim objFSO As New Scripting.FileSystemObject        'FSO����
    Dim strReportName As String
    Dim strPDFTK As String
    Dim varParam As Variant
    Dim blnPrint As Boolean
    Dim lngFileCount As Long
    
    On Error GoTo errHand
    mstrPrintDocIDs = ""
    Set rsPati = GetPatiInfo(lngPatiID, lngVisitID)
    If rsPati.EOF Then Exit Function
    lngDeptID = Val(rsPati!��Ժ����ID & ""): blnDataMove = NVL(rsPati!����ת��, 0) = 1
    If Not objFSO.FolderExists(strFilePath) Then
        Call objFSO.CreateFolder(strFilePath)
    End If
    '����Ƿ����TinyPDF(32λϵͳ) Foxit Reader PDF Printer (64λϵͳ)��ӡ��
    strDeviceName = gobjCommFun.PDFPrinterDeviceName()
    
'    '�����ļ�
    Set rsTmp = GetCISStruct(lngPatiID, lngVisitID, "", blnDataMove)
    If strXML <> "" Then
        Set rsTmp = UpdateRS(rsTmp, strXML)
    End If
    strPath = strFilePath & "\" & NVL(rsPati!����) & "_" & lngPatiID & "_" & lngVisitID & "_"
    Do While Not rsTmp.EOF
        varParam = Split(rsTmp!���� & "", ";")
        strFileName = rsTmp!�ļ�·�� & ""
        blnPrint = False
        If rsTmp!ID Like "R11K*" Then
            If strReportName = "" Then strReportName = GetMedRecReportName(lngDeptID)
            If strFileName = "" Then strFileName = strPath & rsTmp!���� & ".PDF"
            mstrTipInfo = "��ʼ�����" & strFileName
            Call SaveSetting("ZLSOFT", "˽��ģ��\" & mstrUser & "\zl9Report\LocalSet\" & strReportName, "Printer", strDeviceName)
            blnPrint = gobjReport.ReportOpen(gcnOracle, glngSys, strReportName, Nothing, "����ID=" & lngPatiID, "��ҳID=" & lngVisitID, "ReportFormat=" & varParam(0), "PDF=" & strFileName, 4)
        ElseIf rsTmp!ID Like "R12K*" Then
            'ҽ���嵥
            If strFileName = "" Then strFileName = strPath & rsTmp!���� & ".PDF"
            mstrTipInfo = "��ʼ�����" & strFileName
            If objKernel Is Nothing Then
                Set objKernel = CreateObject("zlCISKernel.clsCISKernel")
                Call objKernel.InitCISKernel(gcnOracle, Nothing, glngSys, "")
            End If
            blnPrint = objKernel.zlPrintAdvice(Nothing, lngPatiID, lngVisitID, 0, CInt(varParam(0)), strFileName, 4)
        ElseIf rsTmp!ID Like "R2K*" Or rsTmp!ID Like "R4K*" Or rsTmp!ID Like "R5K*" Or rsTmp!ID Like "R6K*" Then            '2-סԺ����;4-��������;5-����֤��;6-֪���ļ�
            If objRichEPR Is Nothing Then
                Set objRichEPR = CreateObject("zlRichEPR.cRichEPR")
                Call objRichEPR.InitRichEPR(gcnOracle, Nothing, glngSys, False)
                Set mclsDockAduits = CreateObject("zlRichEPR.clsDockAduits")
            End If
            If InStr("," & mstrPrintDocIDs, "," & varParam(0) & ",") = 0 Then    '����û���     '
                If rsTmp!�ϼ�id = "R2" Then
                    strType = "סԺ����"
                ElseIf rsTmp!�ϼ�id = "R4" Then
                    strType = "��������"
                ElseIf rsTmp!�ϼ�id = "R5" Then
                    strType = "����֤��"
                ElseIf rsTmp!�ϼ�id = "R6" Then
                    strType = "֪���ļ�"
                End If
                If strFileName = "" Then strFileName = strPath & strType & "_" & varParam(3) & ".PDF"
                mstrTipInfo = "��ʼ�����" & strFileName
                strPrint = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
                Call gobjCommFun.PDFFile(strFileName)
                blnPrint = mclsDockAduits.zlPrintDocument(3, 2, Val(varParam(0)), strPrint, blnDataMove)
            End If
        ElseIf rsTmp!ID Like "R3K*" Then          '������¼
            If varParam(5) = "1" Then   '�ɰ�'
                If objRichEPR Is Nothing Then
                    Set objRichEPR = CreateObject("zlRichEPR.cRichEPR")
                    Call objRichEPR.InitRichEPR(gcnOracle, Nothing, glngSys, False)
                    Set mclsDockAduits = CreateObject("zlRichEPR.clsDockAduits")
                End If
                If Val(varParam(1)) = -1 Then
                    '���µ�
                    If strFileName = "" Then strFileName = strPath & "���µ�_" & varParam(3) & ".PDF"
                    mstrTipInfo = "��ʼ�����" & strFileName
                    strPrint = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
                    Call gobjCommFun.PDFFile(strFileName)
                    Call mclsDockAduits.zlRefreshTendBody(lngPatiID, lngVisitID, Val(varParam(0)), Val(varParam(4)), blnDataMove)
                    blnPrint = mclsDockAduits.zlPrintDocument(1, 2, , strPrint, blnDataMove)
                Else
                    '������¼
                    If strFileName = "" Then strFileName = strPath & "������¼_" & varParam(3) & ".PDF"
                    mstrTipInfo = "��ʼ�����" & strFileName
                    strPrint = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
                    Call gobjCommFun.PDFFile(strFileName)
                    Call mclsDockAduits.zlRefresh(3, Val(varParam(3)), lngPatiID, lngVisitID, Val(varParam(0)), CStr(varParam(2)), , Val(varParam(4)), blnDataMove)
                    blnPrint = mclsDockAduits.zlPrintDocument(2, 2, , strPrint, blnDataMove)
                End If
            Else  '��
                If strFileName = "" Then strFileName = strPath & Decode(Val(varParam(1)), -1, "���µ�", 1, "����ͼ", "������¼") & "_" & Val(varParam(3)) & ".PDF"
                mstrTipInfo = "��ʼ�����" & strFileName
                Call gobjCommFun.PDFFile(strFileName)
                If objTendsNew Is Nothing Then
                    Set objTendsNew = CreateObject("zl9TendFile.clsTendFile")
                    Call objTendsNew.InitTendFile(gcnOracle, glngSys)
                End If
                blnPrint = objTendsNew.zlPrintDocument(lngPatiID, lngVisitID, Val(varParam(4)), Val(varParam(0)), Val(varParam(3)), Decode(Val(varParam(1)), -1, 1, 1, 3, 2), strDeviceName, True)
            End If
        ElseIf rsTmp!ID Like "R7K*" Then                 '���Ʊ���
            'ϵͳ��,FileName,ID
            If objRichEPR Is Nothing Then
                Set objRichEPR = CreateObject("zlRichEPR.cRichEPR")
                Call objRichEPR.InitRichEPR(gcnOracle, Nothing, glngSys, False)
                Set mclsDockAduits = CreateObject("zlRichEPR.clsDockAduits")
            End If
            If strFileName = "" Then strFileName = strPath & "���Ʊ���_" & varParam(3) & "_" & Val(varParam(0)) & ".PDF"
            mstrTipInfo = "��ʼ�����" & strFileName
            Call gobjCommFun.PDFFile(strFileName)
            strPrint = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
            blnPrint = mclsDockAduits.zlPrintDocument(4, 2, Val(varParam(0)), strPrint, blnDataMove)
        ElseIf rsTmp!ID Like "R7P*" Then
            If Not mobjPublicPACS Is Nothing Then
                If strFileName = "" Then strFileName = strPath & "���Ʊ���_" & rsTmp!���� & ".PDF"
                mstrTipInfo = "��ʼ�����" & strFileName
                blnPrint = mobjPublicPACS.PrintReport(varParam(0), strFileName, False)      'TrueԤ��
            End If
        ElseIf rsTmp!ID Like "R8K*" Then
            If strFileName = "" Then strFileName = strPath & rsTmp!���� & ".PDF"
            mstrTipInfo = "��ʼ�����" & strFileName
            If objDockPath Is Nothing Then
                Set objDockPath = CreateObject("zlCISPath.clsDockPath")
                Set objPath = CreateObject("zlCISPath.clsCISPath")
                Call objPath.InitCISPath(gcnOracle, Nothing, glngSys, "")
            End If
            blnPrint = objDockPath.zlFuncPathTableOutPut(4, True, strFileName, lngPatiID, lngVisitID, strDeviceName)
        ElseIf InStr(rsTmp!ID, "R") = 0 And InStr(rsTmp!����, ";EMR;") > 0 Then
            '���ڹ����ĵ�ʱֻ��ӡ���һ��
            If InStr("," & mstrPrintDocIDs, "," & varParam(0) & ",") = 0 And Not mobjRichEMR Is Nothing Then
                If varParam(1) <> "" Then
                    Call mobjRichEMR.zlShowDoc(varParam(0), varParam(1))
                Else
                    Call mobjRichEMR.zlShowDoc(varParam(0), "")
                End If
                If strFileName = "" Then strFileName = strPath & varParam(2) & varParam(0) & ".PDF"
                mstrTipInfo = "��ʼ�����" & strFileName
                
                Call gobjCommFun.PDFFile(strFileName)
                strPrint = IIf(strDeviceName = "TinyPDF", "TinyPDF|" & strFileName, strDeviceName)
                 
                Err.Number = 0: On Error Resume Next
                blnPrint = mobjRichEMR.zlPrintDoc(False, strPrint)
                If Err.Number = 450 Then
                    Err.Number = 0
                    blnPrint = mobjRichEMR.zlPrintDoc(False)
                End If
                Err.Clear: On Error GoTo 0
                mstrPrintDocIDs = mstrPrintDocIDs & varParam(0) & ","  '
            End If
        End If
        Call gobjCommFun.PDFFileSuccess
        
        If Dir(strFileName) <> "" Then
            If blnMerge Then
                strPDFTK = strPDFTK & strFileName & Space(1)
                lngFileCount = lngFileCount + 1
            End If
        Else
            strNoPDF = strNoPDF & strFileName & ";"
        End If
        mstrTipInfo = ""
        rsTmp.MoveNext
    Loop
    
    If strPDFTK <> "" And lngFileCount > 1 Then
        Shell "pdftk.exe " & strPDFTK & " cat output " & strPath & "�����ϲ�.PDF"
    End If
    mstrTipInfo = ""
    PrintDocument = True
    
    Exit Function
errHand:
    mstrTipInfo = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetPrintList(ByVal lngPatiID As Long, ByVal lngVisitID As Long) As String
'����:��ȡָ�����˵Ŀ�����ļ��嵥��
'����ֵ:�ַ��� ��ʽ����
'   <items>
'       <item>
'           <id></id>--�ļ�ID
'           <type></type> --סԺ����\������¼\��������\����֤��\֪���ļ�\���Ʊ���\��ҳ��Ϣ\ҽ����¼\�ٴ�·��
'           <file_name></file_name> --�ĵ�����
'       </item>
'   </items>

    Dim rsPati As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strXML As String
    
    Set rsPati = GetPatiInfo(lngPatiID, lngVisitID)
    If rsPati.EOF Then Exit Function
    Set rsTmp = GetCISStruct(lngPatiID, lngVisitID, "", NVL(rsPati!����ת��, 0) = 1)
    If rsTmp Is Nothing Then Exit Function
    For i = 1 To rsTmp.RecordCount
        strXML = strXML & "<item>"
        strXML = strXML & "<id>" & rsTmp!ID & "</id>"
        strXML = strXML & "<type>" & GetFileType(rsTmp!�ϼ�id) & "</type>"
        strXML = strXML & "<file_name>" & rsTmp!���� & "</file_name>"
        strXML = strXML & "</item>"
        rsTmp.MoveNext
    Next
    If strXML <> "" Then strXML = "<items>" & strXML & "</items>"
    GetPrintList = strXML
End Function

Private Function GetFileType(ByVal strType As String) As String
    ''1-���ﲡ��;2-סԺ����;3-������¼;4-��������;5-����֤��;6-֪���ļ�;7-���Ʊ���,11-��ҳ��Ϣ,12-ҽ����¼,8-�ٴ�·��
    GetFileType = Split(",���ﲡ��,סԺ����,������¼,��������,����֤��,֪���ļ�,���Ʊ���,�ٴ�·��,,,��ҳ��Ϣ,ҽ����¼", ",")(CInt(Replace(strType, "R", "")))
End Function


Private Sub UpdateEmrInterface(ByVal strUser As String, ByVal strPwd As String)
    If Not mobjEMR Is Nothing Then Exit Sub
    On Error Resume Next
    Err.Clear
    Set mobjEMR = CreateObject("zl9EmrInterface.ClsEmrInterface")
    If Err.Number = 0 Then
        Call mobjEMR.CheckUpdate1(strUser, "[DBPASSWORD]" & strPwd, False)
        If Err.Number <> 0 Then
            Err.Clear
            Call mobjEMR.CheckUpdate(strUser, strPwd)
        End If
    Else
        Set mobjEMR = Nothing
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Private Function GetCISStruct(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String, ByVal blnDataMove As Boolean) As ADODB.Recordset
'����:lng��ҳID סԺ����Ϊ��ҳid,���ﲡ��Ϊ�Һ�id
    Dim strSQL As String, strSQL1 As String
    Dim rsTmp As ADODB.Recordset
    Dim rsMedRec As ADODB.Recordset
    Dim strRptIDs As String
    
    On Error GoTo errH
    '1-���ﲡ��;2-סԺ����;3-������¼;4-��������;5-����֤��;6-֪���ļ�;7-���Ʊ���,11-��ҳ��Ϣ,12-ҽ����¼,8-�ٴ�·��;9-סԺ֤;10-��������
    strSQL = " Select 'R11K1' As ID, 'R11' As �ϼ�id, '��ҳ����' As ����, '1' As ����, '1' As ���� From Dual" & _
            " Union All Select 'R11K2' As ID, 'R11' As �ϼ�id, '��ҳ����' As ����, '2' As ����,'2' As ���� From Dual"
    If Val(gobjDatabase.GetPara("������ҳ��׼", glngSys, pסԺҽ��վ, "0")) = 2 Then
        strSQL = strSQL & _
              " Union All Select 'R11K3' As ID, 'R11' As �ϼ�id, '��ҳ��ҳһ' As ����, '3' As ����, '3' As ���� From Dual" & _
              " Union All Select 'R11K4' As ID, 'R11' As �ϼ�id, '��ҳ��ҳ��' As ����, '4' As ����, '4' As ���� From Dual"
    End If
    
    'ҽ������
    strSQL = strSQL & " Union All Select 'R12K1' As ID, 'R12' As �ϼ�id, '��ʱҽ��' As ����, '1' As ����,'1' As ���� From Dual"
    strSQL = strSQL & " Union All Select 'R12K2' As ID, 'R12' As �ϼ�id, '����ҽ��' As ����, '0' As ����,'2' As ���� From Dual"
    
    '�ٴ�·��
    strSQL1 = "Select 1 From �����ٴ�·�� Where ����id = [1] And ��ҳid = [2] And ״̬ <> 0"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL1, "����Ƿ����Ч���ٴ�·������", lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 0 Then
        strSQL = strSQL & _
          " Union All Select 'R8K' As ID, 'R8' As �ϼ�id, '�ٴ�·��' As ����, '' As ����, '1' As ���� From Dual"
    End If
    
    '��������
    'ID=�ϼ�ID+K����ID
    strSQL = strSQL & " Union All Select A.�ϼ�id||'K'||Trim(To_Char(A.ID)) As ID,A.�ϼ�id," & _
        "       Decode(A.ҽ��id,Null,A.����||'('||To_Char(A.����ʱ��, 'YYYY-MM-DD')||')',A.����||'��'||B.ҽ������||'('||To_Char(A.����ʱ��, 'YYYY-MM-DD')||')') As ����," & _
        "       Trim(To_Char(A.ID))||';'||Decode(A.ҽ��id,Null,'0',Trim(To_Char(A.ҽ��id))) || ';'|| A.RISID||';'|| A.����||';'||A.�༭��ʽ ||';'||C.�������� As ����,���� " & _
        " From (Select A.ID, 'R'||A.�������� As �ϼ�id, A.�������� As ����,C.ҽ��id,C.RISID,A.��������,A.�༭��ʽ,A.����ʱ��,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') As ����" & _
        "       From ���Ӳ�����¼ A,����ҽ������ C " & _
        "       Where A.����id = [1] And A.��ҳid = [2] And (A.������Դ=2 And [3]=1 Or Nvl(A.������Դ,0)<>2 And [3]=0)" & _
        "           And C.����id(+)=A.ID And A.�������� In (1, 2, 3, 4, 5, 6, 7)" & _
        "       ) A,����ҽ����¼ B,������ĿĿ¼ C Where A.ҽ��id=B.Id(+) And B.������ĿID = C.ID(+)"
    '��������
    'ID=�ϼ�ID+K�ļ�ID,0,����ID
    '����=����ID;����;��ʼ����ֹ;�ļ�ID
    '��鱾�β�����ʹ�õ����ϰ廹���°�
    strSQL1 = "Select 1 From ���˻�����¼ A Where a.����id = [1] And a.��ҳid = [2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL1, "����Ƿ�����ϰ�����", lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 0 Then
        strSQL = strSQL & " Union All" & _
            " Select 'R3K'||Trim(To_Char(A.ID)) As ID,'R3' As �ϼ�id," & _
            "       A.����||'('||B.����||'��'||To_Char(A.��ʼ, 'YYYY-MM-DD HH24:MI') || '��' ||To_Char(A.��ֹ, 'YYYY-MM-DD HH24:MI') || ')' As ����," & _
            "       Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(����,0)))||';'||To_Char(A.��ʼ, 'YYYY-MM-DD HH24:MI')||'��'||To_Char(A.��ֹ, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID)) ||';'||'0;1' As ����,To_Char(a.��ʼ,'YYYY-MM-DD HH24:MI:SS') As ����" & _
            " From (" & _
            "   Select F.ID, F.���, F.����, R.��ʼ, R.��ֹ, R.����id, ����" & _
            "   From (" & _
            "       Select ID, ���, ����, 3 As ��������, ͨ��, 0 As ����id, ����" & _
            "          From �����ļ��б� Where ���� = 3 And ���� < 0" & _
            "       Union All" & _
            "       Select L.ID, L.���, L.����, F.���� As ��������, L.ͨ��, A.����id, L.����" & _
            "          From ����ҳ���ʽ F, �����ļ��б� L, ����Ӧ�ÿ��� A" & _
            "          Where L.���� = 3 And L.���� = 0 And L.���� = F.���� And L.��� = F.��� And L.ID = A.�ļ�id(+)" & _
            "       ) F,(" & _
            "       Select R.����id, Nvl(Min(R.��������), 3) As ��������, Min(R.����ʱ��) As ��ʼ, Max(R.����ʱ��) As ��ֹ" & _
            "          From ���˻�����¼ R" & _
            "          Where R.������Դ = 2 And R.����id = [1] And Nvl(R.��ҳid, 0) = [2] And Nvl(R.Ӥ��, 0) = 0" & _
            "          Group By R.����id" & _
            "       ) R" & _
            "       Where (F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = R.����id) And F.�������� >= R.��������" & _
            "   ) A, ���ű� B Where A.����id = B.ID "
    Else
        strSQL = strSQL & " Union All" & _
                " Select 'R3K'||Trim(To_Char(A.ID)) As ID,'R3' As �ϼ�id," & vbNewLine & _
                "     A.����||'('||B.����||'��'||To_Char(A.��ʼ, 'YYYY-MM-DD HH24:MI') || '��' ||To_Char(A.��ֹ, 'YYYY-MM-DD HH24:MI') || ')' As ����," & vbNewLine & _
                "      Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(����,0)))||';'||To_Char(A.��ʼ, 'YYYY-MM-DD HH24:MI')||'��'||To_Char(A.��ֹ, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID))||';'||Trim(To_Char(A.Ӥ��))||';2' As ����,To_Char(a.��ʼ,'YYYY-MM-DD HH24:MI:SS') As ����" & vbNewLine & _
                " From (" & vbNewLine & _
                "   Select R.ID, F.���, R.����,R.Ӥ��, R.��ʼ, NVL(R.��ֹ,nvl(R.ʱ��,R.��ʼ)) ��ֹ, R.����id, ����" & vbNewLine & _
                "   From (" & vbNewLine & _
                "       Select L.ID, L.���, L.����, F.���� As ��������, L.ͨ��, L.����" & vbNewLine & _
                "          From ����ҳ���ʽ F, �����ļ��б� L" & vbNewLine & _
                "          Where L.���� = 3 And L.���� = F.���� And L.��� = F.��� And (L.ͨ��=1 OR L.ͨ��=2)" & vbNewLine & _
                "" & vbNewLine & _
                "       ) F,(" & vbNewLine & _
                "       Select R.ID,R.����id,R.�ļ����� ����,R.��ʽID,nvl(R.Ӥ��,0) Ӥ��,Min(R.��ʼʱ��) As ��ʼ, Max(R.����ʱ��) As ��ֹ,MAX(T.����ʱ��) ʱ��" & vbNewLine & _
                "          From ���˻����ļ� R,���˻������� T" & vbNewLine & _
                "          Where R.ID=T.�ļ�ID(+) And R.����id = [1] And Nvl(R.��ҳid, 0) = [2]" & vbNewLine & _
                "          Group By R.ID,R.�ļ�����,R.����id,R.��ʽID,R.Ӥ��" & vbNewLine & _
                "       ) R" & vbNewLine & _
                "       Where F.ID=R.��ʽID" & vbNewLine & _
                "   ) A, ���ű� B Where A.����id = B.ID And DECODE(A.����,-1,0,A.Ӥ��)=A.Ӥ��"
    End If
    
    strSQL = " Select * From (" & strSQL & ") Order By Decode(�ϼ�id,Null,' ',�ϼ�id),����"
    
    If blnDataMove And lng����ID <> 0 Then
        strSQL = Replace(strSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
        strSQL = Replace(strSQL, "���˻�����¼", "H���˻�����¼")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "���˻����ļ�", "H���˻����ļ�")
        strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "GetCISStruct", lng����ID, lng��ҳID, 1)
    Set rsMedRec = gobjDatabase.CopyNewRec(rsTmp, False, "", Array("�ļ�·��", adVarChar, 500, Empty, "���", adInteger, 1, Empty))
    'BUG:138397��������ҽ����Ӧһ�ݱ����ʱ��ȥ���ظ���Ŀ(������IDȥ��)
    strRptIDs = ""
    rsMedRec.Filter = "�ϼ�ID='R7'"
    Do While Not rsMedRec.EOF
        If InStr("," & strRptIDs & ",", "," & Split(NVL(rsMedRec!����), ";")(0) & ",") = 0 Then
            strRptIDs = strRptIDs & "," & Split(NVL(rsMedRec!����), ";")(0)
        Else
            rsMedRec.Delete
            If rsMedRec.EOF Then Exit Do
        End If
        rsMedRec.MoveNext
    Loop
    rsMedRec.Filter = ""
    'EMR
    Set rsTmp = GetEmrCISStruct(lng����ID, lng��ҳID)
    If Not rsTmp Is Nothing Then
        strRptIDs = ""
        Do While Not rsTmp.EOF
            If InStr("," & strRptIDs & ",", "," & rsTmp!ID & ",") = 0 Then
                rsMedRec.AddNew
                rsMedRec!ID = rsTmp!ID
                rsMedRec!�ϼ�id = rsTmp!�ϼ�id
                rsMedRec!���� = rsTmp!����
                rsMedRec!���� = NVL(rsTmp!����) & ";EMR;" & rsTmp!�ϼ�id
                rsMedRec.Update
                strRptIDs = strRptIDs & "," & rsTmp!ID
            End If
            rsTmp.MoveNext
        Loop
    End If
    '�°�PACS
    Set rsTmp = Nothing
    If Not mobjPublicPACS Is Nothing Then
        Set rsTmp = mobjPublicPACS.zlDocGetList(lng����ID, lng��ҳID, str�Һŵ�)
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                rsMedRec.AddNew
                rsMedRec!ID = "R7P" & rsTmp!����ID
                rsMedRec!�ϼ�id = "R7"
                rsMedRec!���� = rsTmp!�ĵ����� & ""
                rsMedRec!���� = rsTmp!����ID
                rsMedRec.Update
                rsTmp.MoveNext
            Loop
        End If
    End If
    rsMedRec.Filter = ""
    Set GetCISStruct = rsMedRec
    Exit Function
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function GetEmrCISStruct(ByVal lngPatiID As Long, ByVal lngPageID As Long) As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strExtendTag As String, strReturn As String, strSQL As String, strSQLNew As String
    
    On Error GoTo errH
    If mobjEMR Is Nothing Then Set GetEmrCISStruct = Nothing: Exit Function
    strExtendTag = GetEMRIn_Tag(lngPatiID, lngPageID)
    If strExtendTag = "" Then Set GetEmrCISStruct = Nothing: Exit Function
    
    '�ϼ�ID��ID�����ƣ�������ͼ��
    strSQL = "Select Decode(e.Kind, '01', 'R1', '02', 'R2', '03', 'R4', '04', 'R5', '05', 'R6', 'R2') �ϼ�id," & vbNewLine & _
            "       Nvl(d.Subdoc_Id, Rawtohex(d.Real_Doc_Id)) As ID, d.Subdoc_Id As ���ĵ�id," & vbNewLine & _
            "       e.Title ||" & vbNewLine & _
            "        Decode(d.Completor, Null, ''," & vbNewLine & _
            "               '�� ' || d.Completor || ' ��' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || 'ǩ����') As ����," & vbNewLine & _
            "       Rawtohex(d.Real_Doc_Id) || Decode(d.Subdoc_Id, Null, ';', ';' || d.Subdoc_Id) || ';' ||Nvl(d.Subdoc_Title, E.Title) As ����, 'object_case' As ͼ��" & vbNewLine & _
            "From (Select Distinct d.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor" & vbNewLine & _
            "       From Bz_Act_Log A, Bz_Act_Log D, Bz_Doc_Tasks C" & vbNewLine & _
            "       Where a.Extend_Tag = :etag And (a.Id = d.Id Or a.Id = d.Basiclog_Id) And d.Id = c.Actlog_Id And" & vbNewLine & _
            "             c.Real_Doc_Id Is Not Null) D, Antetype_List E" & vbNewLine & _
            "Where d.Antetype_Id = e.Id  And e.Title = Decode(e.Type, 3, d.Subdoc_Title, e.Title)" & vbNewLine & _
            "Order By Rawtohex(d.Real_Doc_Id), e.Code, d.Complete_Time"
            
    strSQLNew = "Select Decode(e.Kind, '01', 'R1', '02', 'R2', '03', 'R4', '04', 'R5', '05', 'R6', 'R2') �ϼ�id," & vbNewLine & _
                "       Nvl(d.Subdoc_Id, Rawtohex(d.Real_Doc_Id)) As ID, d.Subdoc_Id As ���ĵ�id," & vbNewLine & _
                "       e.Title ||" & vbNewLine & _
                "        Decode(d.Completor, Null, ''," & vbNewLine & _
                "               '�� ' || d.Completor || ' ��' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || 'ǩ����') As ����," & vbNewLine & _
                "       Rawtohex(d.Real_Doc_Id) || Decode(d.Subdoc_Id, Null, ';', ';' || d.Subdoc_Id) || ';' ||Nvl(d.Subdoc_Title, E.Title) As ����, 'object_case' As ͼ��" & vbNewLine & _
                "From (Select Distinct d.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor, c.Order_No" & vbNewLine & _
                "       From Bz_Act_Log A, Bz_Act_Log D, Bz_Doc_Tasks C" & vbNewLine & _
                "       Where a.Extend_Tag = :etag And (a.Id = d.Id Or a.Id = d.Basiclog_Id) And d.Id = c.Actlog_Id And" & vbNewLine & _
                "             c.Real_Doc_Id Is Not Null And Nvl(c.Intead, 0) = 0) D, Antetype_List E" & vbNewLine & _
                "Where d.Antetype_Id = e.Id " & vbNewLine & _
                "Order By Rawtohex(d.Real_Doc_Id), e.Code, d.Order_No"
    
    Err.Clear
    On Error Resume Next
    strReturn = mobjEMR.OpenSQLRecordset(strSQLNew, strExtendTag & "^16^etag", rsTemp)
    If Err.Number <> 0 Or strReturn <> "" Then
        Err.Clear
        strReturn = mobjEMR.OpenSQLRecordset(strSQL, strExtendTag & "^16^etag", rsTemp)
    End If
    
    If strReturn <> "" Then
        MsgBox strReturn, vbCritical, gstrSysName
        Set GetEmrCISStruct = Nothing: Exit Function
    End If
    
    Set GetEmrCISStruct = rsTemp
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function GetEMRIn_Tag(ByVal lngPatiID As Long, ByVal lngPageID As Long) As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
 
    strSQL = "Select Nvl(a.Id, b.Id) ID" & vbNewLine & _
                "From (Select Max(ID) ID From ���˱䶯��¼ Where ����id = [1] And ��ҳid = [2] And ��ʼԭ�� = 2 And Nvl(���Ӵ�λ, 0) = 0) A," & vbNewLine & _
                "     (Select Max(ID) ID From ���˱䶯��¼ Where ����id = [1] And ��ҳid = [2] And ��ʼԭ�� = 1 And Nvl(���Ӵ�λ, 0) = 0) B"
    
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������ԺID", lngPatiID, lngPageID)
    
    If rsTmp Is Nothing Then Exit Function
    If NVL(rsTmp!ID) = "" Then Exit Function
    GetEMRIn_Tag = "BD_" & rsTmp!ID
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Sub Class_Terminate()
    glngSys = 0
    
    Set mobjRegister = Nothing
    Set gcnOracle = Nothing
    Set gobjReport = Nothing
    Set mobjPublicPACS = Nothing
    Set mobjRichEMR = Nothing
    Set mclsDockAduits = Nothing
    Set mobjEMR = Nothing
End Sub

Private Sub mclsDockAduits_AfterEprPrint(ByVal lngRecordId As Long)
    mstrPrintDocIDs = mstrPrintDocIDs & lngRecordId & ","
End Sub

Private Function GetMedRecReportName(ByVal lngDeptID As Long) As String
'����:��ȡ������ҳ��������
    Dim strReportName As String
    
    Select Case Val(gobjDatabase.GetPara("������ҳ��׼", glngSys, pסԺҽ��վ, "0"))
         Case 0 '��������׼
             If gobjComlib.Sys.DeptHaveProperty(lngDeptID, "��ҽ��") Then
                 strReportName = "ZL1_INSIDE_1261_4"
             Else
                 strReportName = "ZL1_INSIDE_1261_1"
             End If
         Case 1    '�Ĵ�ʡ��׼
             If gobjComlib.Sys.DeptHaveProperty(lngDeptID, "��ҽ��") Then
                 strReportName = "ZL1_INSIDE_1261_6"
             Else
                 strReportName = "ZL1_INSIDE_1261_5"
             End If
         Case 2    '����ʡ��׼
             If gobjComlib.Sys.DeptHaveProperty(lngDeptID, "��ҽ��") Then
                 strReportName = "ZL1_INSIDE_1261_8"
             Else
                 strReportName = "ZL1_INSIDE_1261_7"
             End If
         Case 3    '����ʡ��׼
             If gobjComlib.Sys.DeptHaveProperty(lngDeptID, "��ҽ��") Then
                 strReportName = "ZL1_INSIDE_1261_10"
             Else
                 strReportName = "ZL1_INSIDE_1261_9"
             End If
     End Select
     GetMedRecReportName = strReportName
End Function