VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjDrawReport  As Object
Private mobjPackage     As Object

Public Sub Init(ByVal cnOracle As ADODB.Connection)
'���ܣ��ӿڳ�ʼ��
'������cnOracle=�����
'      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      �������� lngModual=1070����int���ϲ���
    Dim objFso  As FileSystemObject
    Dim rsTmp   As ADODB.Recordset
    
    If Not gblnInited Then
        Set gcnOracle = cnOracle
        Call GetUserInfo
        gblnInited = True
        gstrSysName = "���鱨���ӡ"
        gstrHospital = gobjComLib.zlRegInfo("��λ����")
        InitObjLis
        Set objFso = New FileSystemObject
        gstrFilePath = App.Path & "\LisReportImge"
        
        gstrSignPath = gstrFilePath & "\SignImage"
        
        If Not objFso.FolderExists(gstrFilePath) Then
            objFso.CreateFolder (gstrFilePath)
        End If
        If Not objFso.FolderExists(gstrSignPath) Then
            objFso.CreateFolder (gstrSignPath)
        End If
        Set objFso = Nothing
        gbln��ʾͼƬ = True
        Call DeleteImge
        Set mobjPackage = New clsLisPackage
        '�ж��Ƿ�Ϊ�°�LIS�����ݼ��鱨���¼�����걾��¼�еĺ���ʱ������ж�
        Set rsTmp = mobjPackage.Get_LisVersion()
        If Not ChkRsState(rsTmp) Then
            If CDate(rsTmp!����ʱ��1) > CDate(rsTmp!����ʱ��2) Then
                gblnNewLis = True
                glngSys = 2500
                glngModual = 2500
            Else
                gblnNewLis = False
                glngSys = 100
                glngModual = 1208
            End If
        Else
            MsgBox "�ж�LIS�汾������", vbInformation, gstrSysName
            Exit Sub
        End If
'        gblnNewLis = False
'        glngSys = 100
'        glngModual = 1208
    End If
End Sub

Public Function GetForm() As Object
'���ܣ���������ֵ���ض��󡣹���վ�������ʱ����
On Error GoTo ErrH:
   Set GetForm = New frmListReportView
'frmListReportView.Show vbModal
    Exit Function
ErrH:
    MsgBox err.Description
End Function

Public Sub PrintLisReport(ByRef objForm As Object, ByVal lngPatiID As Long, ByVal str�Һŵ� As String, ByVal lng��ҳID As Long, ByVal bytType As Integer, _
                          bytMode As Byte, ByVal blnDoctor As Boolean, ByVal blnPrintAll As Boolean)
'���ܣ�ˢ���ڲ�����
'������strName ��ǰ��������
'      lngPatiID ����ID,str�Һŵ� �Һŵ������ﲡ�˲��з���Ϊ�գ�lng��ҳID סԺ���˲��з���Ϊ0
'      str�Һŵ� �Һŵ������ﲡ�˴���
'      lng��ҳID ��ҳID��סԺ���˴���
'      bytType 1-ѡ���2-��ť��ӡ
'    Dim bytMode         As Byte        1-Ԥ����2-��ӡ
    Dim strSql As String
    Dim strPage As String
    Dim blnState As Boolean
    Dim rsParameter As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim str�Һ����� As String
    Dim str���뿪ʼʱ�� As String
    Dim str�������ʱ�� As String
    Dim bln�Ƿ��� As Boolean
    Dim lngPageCount As Long

    On Error GoTo ErrH:
    
    gblnPrintAll = blnPrintAll
    gblnDoctor = blnDoctor
    gbtyModel = bytType
    If objForm Is Nothing Then Exit Sub
    If lngPatiID = 0 Then    'Exit Sub
        Call objForm.HideControl(lngPatiID, blnState)
        Exit Sub
    Else
        If gbtyModel = 1 Then objForm.Show

        If Not mobjDrawReport Is Nothing Then Set mobjDrawReport = Nothing    'mobjDrawReport.DestoryDrawReport

        Set mobjDrawReport = New clsDrawLisReport

        If Not (mobjDrawReport Is Nothing) Then
            blnState = mobjDrawReport.InitData("zlPublicLIS.clsLisReportPublic", "����ģ��\" & "zl9PrintMode" & "\Default")
        End If


        '��ȡ��������Ŀ�ʼ�����ʱ��
        Set rsTmp = mobjPackage.Get_RequestTime(lngPatiID, lng��ҳID, str�Һŵ�)

        str�Һ����� = gobjDatabase.GetPara("�Һ���Ч����", 100)
        If Not ChkRsState(rsTmp) Then
            bln�Ƿ��� = (rsTmp!���� = 1)
            If bln�Ƿ��� And lng��ҳID = 0 Then
                str�������ʱ�� = Format(rsTmp!�������ʱ��, "yyyy-MM-dd HH:mm:ss")
                str���뿪ʼʱ�� = Format(CDate(Format(rsTmp!�������ʱ��, "YYYY-MM-DD")) - Mid(str�Һ�����, 2, 1) & " 00:00:00", "yyyy-MM-dd HH:mm:ss")
            ElseIf (Not bln�Ƿ���) And lng��ҳID = 0 Then
                str�������ʱ�� = Format(rsTmp!�������ʱ��, "yyyy-MM-dd HH:mm:ss")
                str���뿪ʼʱ�� = Format(CDate(Format(rsTmp!�������ʱ��, "YYYY-MM-DD")) - Mid(str�Һ�����, 1, 1) & " 00:00:00", "yyyy-MM-dd HH:mm:ss")
            Else
                str���뿪ʼʱ�� = Format(rsTmp!���뿪ʼʱ��, "yyyy-MM-dd HH:mm:ss")
                str�������ʱ�� = Format(rsTmp!�������ʱ��, "yyyy-MM-dd HH:mm:ss")
            End If
        Else
            Exit Sub
        End If

        Set rsParameter = mobjPackage.CreateParameter()

        Call mobjPackage.SetParameter(rsParameter, "��������", "")
        Call mobjPackage.SetParameter(rsParameter, "����ID", lngPatiID)
        Call mobjPackage.SetParameter(rsParameter, "�Һŵ���", str�Һŵ�)
        Call mobjPackage.SetParameter(rsParameter, "��ҳID", lng��ҳID)
        Call mobjPackage.SetParameter(rsParameter, "���뿪ʼʱ��", str���뿪ʼʱ��)
        Call mobjPackage.SetParameter(rsParameter, "�������ʱ��", str�������ʱ��)
        ''        Call mobjPackage.SetParameter(rsParameter, "�걾˳��", mobjPackage.Get_LisSampleSort())
        ''        Call mobjPackage.SetParameter(rsParameter, "���˳��", mobjPackage.Get_LisTypeSort())
        ''        Call mobjPackage.SetParameter(rsParameter, "��Ŀ˳��", mobjPackage.Get_LisNameSort())

        '��������=����'�����С=9'�������=False'����б��=False|����1��������=����'����1�����С=18'����1�������=FALSE'����1����б��=FALSE

        mobjDrawReport.BeginGetData

        '�ڶ�����ҳüҳ��
        '------------------------------------------------------------------------------------------------------------------
        '    If mmobjDrawReport.GetPageHeadData("zlPlugIn.clsElePersonPublic", rsParameter) = False Then GoTo EndHand
        blnState = mobjDrawReport.GetPageFootData("zlPublicLIS.clsLisReportPublic", rsParameter)

        '���Ĳ������������(�������屨���������ѭ����������������������֯)
        '------------------------------------------------------------------------------------------------------------------
        '��ʼ��ȡ����������
        blnState = mobjDrawReport.GetElementData("zlPublicLIS.clsLisResultItem", rsParameter, False, 0)

        Call mobjDrawReport.EndGetData
        If blnState Then
            blnState = mobjDrawReport.ExportReport(objForm, bytMode, "���鱨���ӡ", "", "", strPage)

            'д���ӡ��¼
            If Not grsSample Is Nothing Then
                grsSample.Filter = ""
                If grsSample.RecordCount > 0 Then grsSample.MoveFirst
                Do While Not grsSample.EOF
                    If gblnNewLis Then
                        strSql = "Zl_���鱨���ӡ_Edit(1," & grsSample("�걾ID") & "," & IIf(blnDoctor, 1, 0) & ")"
                        Call gobjPublicHisCommLis.ComExecuteOtherProc(1, strSql, "��ӡ�걾")
                    Else
                        strSql = "Zl_����걾��¼_�걾�ʿ�(" & grsSample("�걾ID") & ",'',1)"
                         Call gobjDatabase.ExecuteProcedure(strSql, gstrSysName)
                    End If
                    grsSample.MoveNext
                Loop
            End If
        End If
        lngPageCount = objForm.GetPageCount()
        If lngPageCount <= 0 Then blnState = False
        Call objForm.HideControl(lngPatiID, blnState)
    End If
    Exit Sub
ErrH:
    MsgBox err.Description, vbCritical, "���鱨���ӡ"
    '    err.Clear
End Sub

Private Sub Class_Initialize()
    If gobjComLib Is Nothing Then
        Set gobjComLib = GetObject("", "zl9Comlib.clsComlib")
        Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
        Set gobjControl = GetObject("", "zl9Comlib.clsControl")
        Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
        Set gobjPrintMode = GetObject("", "zl9PrintMode.zlPrintMethod")
        Set gobjSystem = GetObject("", "zl9Comlib.clsSystem")
    End If
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
End Sub

Private Sub Class_Terminate()
    If Not mobjDrawReport Is Nothing Then mobjDrawReport.DestoryDrawReport
    '�˳�����ʱ��ɾ��ͼƬĿ¼�µ�����ͼƬ
    Call DeleteImge
    Set mobjPackage = Nothing
    gblnInited = False
'    Set gobjComLib = Nothing
End Sub