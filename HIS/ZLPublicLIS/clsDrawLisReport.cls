VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrawLisReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'��������
'######################################################################################################################
Private Type Property
    �������� As String
    �����С As Byte
    ������� As Boolean
    ����б�� As Boolean
    �����»��� As Boolean
    ����ǰ��ɫ As Long
End Type

Private mProperty As Property

Private mfrmMain As Object

Private mobjElement             As Object
Private mstrRegisterPath        As String
Private mclsLisReportLib        As clsLisReportLib
Private mclsLisPackage          As clsLisPackage

Public Event DataFormProgress(ByVal strProgressInfo As String)
Public Event AfterPrinted()

'�ڶ��㣬��ȡԭʼ���ݲ����и�ʽ��
'######################################################################################################################
Public Function InitData(Optional ByVal strElement As String, Optional ByVal strRegisterPath As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mstrRegisterPath = strRegisterPath
    
'    Call gobjComLib.InitCommon(gcnOracle)
    
    '������ʽ���������
    '------------------------------------------------------------------------------------------------------------------
    If strElement <> "" Then
        Set mobjElement = Nothing
        
        err = 0
        On Error Resume Next
        Set mobjElement = CreateObject(strElement)
        On Error GoTo 0
        
        If mobjElement Is Nothing Then
            ShowSimpleMsg "��ָ���ĸ�ʽ����(" & strElement & ")�����ڣ�"
            Exit Function
        End If
    End If
    
    '
    '------------------------------------------------------------------------------------------------------------------
    Set gobjDraw = frmPubResource.picDraw
    
    Set grsData = New ADODB.Recordset
    With grsData
        .Fields.Append "���", adVarChar, 30
        .Fields.Append "ҳ��", adInteger
        .Fields.Append "���", adVarChar, 30
        .Fields.Append "����", adInteger
        .Fields.Append "����", adVarChar, 20
        .Fields.Append "����", adVarChar, 4000
        .Fields.Append "X0", adVarChar, 20
        .Fields.Append "Y0", adVarChar, 20
        .Fields.Append "X1", adVarChar, 20
        .Fields.Append "Y1", adVarChar, 20
        .Fields.Append "B0", adVarChar, 20
        .Fields.Append "R0", adVarChar, 20
        .Fields.Append "����", adVarChar, 20
        .Fields.Append "ǰ��ɫ", adVarChar, 20
        .Fields.Append "����ɫ", adVarChar, 20
        .Fields.Append "��С", adInteger
        .Fields.Append "����", adTinyInt
        .Fields.Append "б��", adTinyInt
        .Fields.Append "�»���", adTinyInt
        .Fields.Append "�������", adTinyInt
        .Fields.Append "�������", adTinyInt
        .Fields.Append "�Զ�����", adTinyInt
        .Fields.Append "��������", adTinyInt
        .Fields.Append "��������", adTinyInt
        .Fields.Append "����", adBigInt
        .Fields.Append "�Զ���Ӧ", adTinyInt
        .Fields.Append "��ת�Ƕ�", adTinyInt
        .Fields.Append "���", adVarChar, 30
        .Open
    End With
                                    
    '------------------------------------------------------------------------------------------------------------------
    Set grsPage = New ADODB.Recordset
    With grsPage
        .Fields.Append "��ҳ", adInteger
        .Fields.Append "ҳ��", adInteger
        .Fields.Append "������ҳ", adInteger
        .Fields.Append "����ҳ��", adInteger
        .Fields.Append "ҳ�����", adTinyInt
        .Fields.Append "��ʾҳü", adTinyInt
        .Fields.Append "��ʾҳ��", adTinyInt
        .Fields.Append "��ʾ����", adVarChar, 50
        .Open
    End With
    
    glngVirtualPages = 0
    
    '------------------------------------------------------------------------------------------------------------------
    Set grsList = New ADODB.Recordset
    With grsList
        .Fields.Append "Ŀ¼����", adTinyInt
        .Fields.Append "Ŀ¼ҳ��", adInteger
        .Fields.Append "Ŀ¼����", adVarChar, 200
        .Fields.Append "Ŀ¼����", adTinyInt
        .Open
    End With
    
    Set grsTempFile = New ADODB.Recordset
    With grsTempFile
        .Fields.Append "�ļ�", adVarChar, 100
        .Open
    End With

    If Not mobjElement Is Nothing Then
        InitData = mobjElement.InitElement(gobjPaper, mstrRegisterPath)
    End If
        
End Function

Public Function PrintSet(frmMain As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    If Not mobjElement Is Nothing Then
        PrintSet = mobjElement.FormatSetup(frmMain, mstrRegisterPath)
    End If
    
End Function

Public Function ExportReport(ByRef frmReport As Object, Optional ByVal bytMode As Byte = 1, Optional ByVal strCaption As String = "", Optional ByVal strFile As String, Optional ByVal strPassWord As String, Optional ByVal strPage As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������bytMode=1��Ԥ��;=2,��ӡ
    '���أ�
    '******************************************************************************************************************
    
    If Not grsData Is Nothing Then
        If grsData.State = adStateOpen Then
            
            If frmReport.InitReport(grsData, grsPage, mstrRegisterPath) Then
                Call frmReport.ExportReport(bytMode, strCaption, strFile, strPassWord, strPage)
            End If
            
        End If
    End If
    
    ExportReport = True
    
End Function

Public Function BeginGetData() As Boolean
    '******************************************************************************************************************
    '���ܣ���ʽ������
    '������
    '���أ�
    '******************************************************************************************************************
    If Not mobjElement Is Nothing Then
        
        gobjRect.Page = 0
        
        BeginGetData = True
    End If
    
End Function

Public Function GetHeadData() As Boolean
    '******************************************************************************************************************
    '���ܣ���ʽ������
    '������
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand

    If Not mobjElement Is Nothing Then

        GetHeadData = mobjElement.FormatHeadData
        
    End If
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetFootData() As Boolean
    '******************************************************************************************************************
    '���ܣ���ʽ������
    '������
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand

    If Not mobjElement Is Nothing Then

        GetFootData = mobjElement.FormatFootData
        
    End If
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function AnalyseProperty(ByVal strProperty As String) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim intPos As Integer
    Dim strTmp As String
    Dim aryProperty As Variant
    
    
    If strProperty = "" Then Exit Function
    aryProperty = Split(strProperty, "'")
    For intLoop = 0 To UBound(aryProperty)
        strTmp = CStr(aryProperty(intLoop))
        intPos = InStr(strTmp, "=")
        If intPos > 0 Then
            Select Case Trim(Left(strTmp, intPos - 1))
            Case "��������"
                mProperty.�������� = Trim(Mid(strTmp, intPos + 1))
            Case "�����С"
                mProperty.�����С = Val(Mid(strTmp, intPos + 1))
            Case "����ǰ��ɫ"
                mProperty.����ǰ��ɫ = Val(Mid(strTmp, intPos + 1))
            Case "�������"
                mProperty.������� = (UCase(Trim(Mid(strTmp, intPos + 1))) = "TRUE")
            Case "����б��"
                mProperty.����б�� = (UCase(Trim(Mid(strTmp, intPos + 1))) = "TRUE")
            Case "�����»���"
                mProperty.�����»��� = (UCase(Trim(Mid(strTmp, intPos + 1))) = "TRUE")
            End Select
        End If
    Next
    
    AnalyseProperty = True
End Function

Public Function GetPageHeadData(ByVal strElement As String, Optional ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ���ʽ������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objElement As Object
    Dim objRect As USERRECT
    Dim objFont As USERFONT
    Dim objSvrFont As USERFONT
    
    On Error GoTo errHand

    If Not mobjElement Is Nothing Then

        objSvrFont = mclsLisReportLib.Font
        objRect = mclsLisReportLib.RECT
        
        objFont.Name = mProperty.��������
        objFont.Size = 9
        objFont.Bold = False
        objFont.Italic = False
        objFont.Underline = False
        objFont.BackColor = USERCOLOR.��ɫ
        objFont.ForeColor = USERCOLOR.��ɫ
        objFont.AutoFit = False
        objFont.HAlignment = 1
        objFont.VAlignment = 2
        objFont.MuliLine = False
        objFont.Rows = 1
        objFont.WordWarp = False

        mProperty.�������� = "����"
        mProperty.������� = False
        mProperty.����б�� = False
        mProperty.�����С = 9
        mProperty.�����»��� = False
        mProperty.����ǰ��ɫ = 0
    
        Call AnalyseProperty(mclsLisPackage.GetParameter(rsParameter, "��������"))

        objFont.Name = mProperty.��������
        objFont.Size = mProperty.�����С
        objFont.Bold = mProperty.�������
        objFont.Italic = mProperty.����б��
        objFont.Underline = mProperty.�����»���
        objFont.ForeColor = mProperty.����ǰ��ɫ
        
        mobjElement.CurrentRect = objRect
        mobjElement.CurrentFont = objFont
        
        Call mobjElement.SetElementParameter(rsParameter)
        GetPageHeadData = mobjElement.GetPageHeadData
    End If
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPageFootData(ByVal strElement As String, Optional ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ���ʽ������
    '������
    '���أ�
    '******************************************************************************************************************
'    Dim objElement As Object
    Dim objRect As USERRECT
    Dim objFont As USERFONT
    Dim objSvrFont As USERFONT
    
    On Error GoTo errHand

'
    If Not mobjElement Is Nothing Then

        objSvrFont = mclsLisReportLib.Font
        objRect = mclsLisReportLib.RECT
        
        objFont.Name = mProperty.��������
        objFont.Size = 9
        objFont.Bold = False
        objFont.Italic = False
        objFont.Underline = False
        objFont.BackColor = USERCOLOR.��ɫ
        objFont.ForeColor = USERCOLOR.��ɫ
        objFont.AutoFit = False
        objFont.HAlignment = 1
        objFont.VAlignment = 2
        objFont.MuliLine = False
        objFont.Rows = 1
        objFont.WordWarp = False

        mProperty.�������� = "����"
        mProperty.������� = False
        mProperty.����б�� = False
        mProperty.�����С = 9
        mProperty.�����»��� = False
        mProperty.����ǰ��ɫ = 0
    
        Call AnalyseProperty(mclsLisPackage.GetParameter(rsParameter, "��������"))

        objFont.Name = mProperty.��������
        objFont.Size = mProperty.�����С
        objFont.Bold = mProperty.�������
        objFont.Italic = mProperty.����б��
        objFont.Underline = mProperty.�����»���
        objFont.ForeColor = mProperty.����ǰ��ɫ
        
        mobjElement.CurrentRect = objRect
        mobjElement.CurrentFont = objFont
        
        Call mobjElement.SetElementParameter(rsParameter)
        GetPageFootData = mobjElement.GetPageFootData()
                
        RaiseEvent DataFormProgress("aa")
        
    End If
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetElementData(ByVal strElement As String, Optional ByVal rsParameter As ADODB.Recordset, Optional ByVal blnNewPage As Boolean, Optional ByVal bytList As Byte = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ���ʽ������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objElement As Object
    Dim objRect As USERRECT
    Dim objFont As USERFONT
    Dim objSvrFont As USERFONT
    
    On Error GoTo errHand

    '������ʽ���������
    '------------------------------------------------------------------------------------------------------------------
    If strElement <> "" Then
        Set objElement = Nothing
        err = 0
        
        On Error Resume Next
        Set objElement = CreateObject(strElement)
        On Error GoTo 0
        
        If objElement Is Nothing Then
            ShowSimpleMsg "��ָ���ĸ�ʽ����(" & strElement & ")�����ڣ�"
            Exit Function
        End If
    End If
    
    If Not objElement Is Nothing Then
        If objElement.InitElement(2, gcnOracle) Then
            
            objSvrFont = mclsLisReportLib.Font
            objRect = mclsLisReportLib.RECT
            
            objFont.Name = mProperty.��������
            objFont.Size = 9
            objFont.Bold = False
            objFont.Italic = False
            objFont.Underline = False
            objFont.BackColor = USERCOLOR.��ɫ
            objFont.ForeColor = USERCOLOR.��ɫ
            objFont.AutoFit = False
            objFont.HAlignment = 1
            objFont.VAlignment = 2
            objFont.MuliLine = False
            objFont.Rows = 1
            objFont.WordWarp = False

            mProperty.�������� = "����"
            mProperty.������� = False
            mProperty.����б�� = False
            mProperty.�����С = 9
            mProperty.�����»��� = False
            mProperty.����ǰ��ɫ = 0
        
            Call AnalyseProperty(mclsLisPackage.GetParameter(rsParameter, "��������"))
            Call AnalyseProperty(mclsLisPackage.GetParameter(rsParameter, "�������"))
            Call AnalyseProperty(mclsLisPackage.GetParameter(rsParameter, "��������"))
    
            objFont.Name = mProperty.��������
            objFont.Size = mProperty.�����С
            objFont.Bold = mProperty.�������
            objFont.Italic = mProperty.����б��
            objFont.Underline = mProperty.�����»���
            objFont.ForeColor = mProperty.����ǰ��ɫ
            
            If blnNewPage Then
                objRect.Page = objRect.Page + 1
                Call InsertPage(objRect.Page, 1, "", True, True)
                objRect.X0 = mclsLisReportLib.Paper.BorderLeft
                objRect.Y0 = mclsLisReportLib.Paper.BorderTop + mclsLisReportLib.Paper.PageHead
                objRect.Y1 = 0
            Else
                objRect.X0 = mclsLisReportLib.Paper.BorderLeft
                objRect.X1 = mclsLisReportLib.Paper.Width - mclsLisReportLib.Paper.BorderRight
            End If
            
            objElement.CurrentRect = objRect
            objElement.CurrentFont = objFont
            
            Call objElement.SetElementParameter(rsParameter)
            GetElementData = objElement.GetElementData()
                
            objRect = objElement.CurrentRect
            
            objRect.Y0 = objRect.Y1 + mclsLisReportLib.GetTwipsY(0.2)
            objRect.Y1 = 0
            mclsLisReportLib.RECT = objRect
            mclsLisReportLib.Font = objSvrFont
        End If
        
        Set objElement = Nothing
        
    End If
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function EndGetData() As Boolean
    '******************************************************************************************************************
    '���ܣ���ʽ������
    '������
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If Not mobjElement Is Nothing Then

        If CalcPagesData(grsPage) = False Then GoTo EndHand
        
        EndGetData = True
    End If
    
    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    
EndHand:
    
End Function

Public Function FormatData() As Boolean
    '******************************************************************************************************************
    '���ܣ���ʽ������
    '������
    '���أ�
    '******************************************************************************************************************
    
    On Error GoTo errHand

    If Not mobjElement Is Nothing Then
        
        gobjCommFun.ShowFlash "������֯���ݣ����Ժ�...."
        
        gobjRect.Page = 0
        If mobjElement.FormatHeadData = False Then GoTo EndHand
        If mobjElement.FormatFootData = False Then GoTo EndHand
        If mobjElement.FormatPageData = False Then GoTo EndHand
        If CalcPagesData(grsPage) = False Then GoTo EndHand
        
        gobjCommFun.StopFlash
        
        FormatData = True
        
    End If
    
    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    gobjCommFun.StopFlash
    
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    
EndHand:
    gobjCommFun.StopFlash
End Function

Private Function CalcPagesData(ByRef rsPage As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ����¼�������ҳ�š�������ҳ����ҳ
    '������
    '���أ�
    '******************************************************************************************************************

    Dim int����ҳ�� As Integer
    Dim int������ҳ As Integer
    Dim int��ҳ As Integer

    On Error GoTo errHand

    '------------------------------------------------------------------------------------------------------------------
    rsPage.Filter = "ҳ�����=1"
    int������ҳ = rsPage.RecordCount

    rsPage.Filter = ""
    int��ҳ = rsPage.RecordCount

    If rsPage.RecordCount > 0 Then
        rsPage.MoveFirst
        Do While Not rsPage.EOF
            rsPage("��ҳ").Value = int��ҳ
            rsPage("������ҳ").Value = int������ҳ

            If Val(rsPage("ҳ�����").Value) = 1 Then
                int����ҳ�� = int����ҳ�� + 1
                rsPage("����ҳ��").Value = int����ҳ��
            End If
            rsPage.MoveNext
        Loop
        rsPage.MoveFirst
    End If
    
    
    'ͬʱ����Ŀ¼ҳ��Ϊ����ҳ��
    
    CalcPagesData = True

    Exit Function

    '��������
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function AdjustListPage(ByVal intFromPage As Integer, ByVal intToPage As Integer) As Boolean
    '******************************************************************************************************************
    '���ܣ�����intFromPage�����е�ҳ�Ƶ���intToPageҳ֮��
    '������
    '���أ�
    '�㷨����ingPage���ҳ�ź���ҳ�ŷֱ����Ŀ¼��ҳ��
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intListPage As Integer

    On Error GoTo errHand
    
    If grsList.RecordCount > 0 And grsPage.RecordCount > 0 And intFromPage > 0 Then  'And intToPage > 0 ����ֱ�ӷ��ڵ�һҳ
        
        '�Ƚ�Ŀ¼ҳ�ű�Ϊ����
        '------------------------------------------------------
        grsPage.Filter = ""
        grsPage.Filter = "ҳ��>=" & intFromPage
        If grsPage.RecordCount = 0 Then Exit Function
        intListPage = grsPage.RecordCount
        grsPage.MoveFirst
        Do While Not grsPage.EOF
            grsPage("ҳ��").Value = 0 - grsPage("ҳ��").Value
            grsPage.MoveNext
        Loop
        
        grsData.Filter = ""
        grsData.Filter = "ҳ��>=" & intFromPage
        If grsData.RecordCount > 0 Then
            grsData.MoveFirst
            Do While Not grsData.EOF
                grsData("ҳ��").Value = 0 - grsData("ҳ��").Value
                grsData.MoveNext
            Loop
        End If
        
        '����ƶ�
        '------------------------------------------------------
        grsPage.Filter = ""
        grsPage.Filter = "ҳ��>" & intToPage
        If grsPage.RecordCount > 0 Then
            grsPage.MoveFirst
            Do While Not grsPage.EOF
                grsPage("ҳ��").Value = grsPage("ҳ��").Value + intListPage
                grsPage.MoveNext
            Loop
        End If
        
        grsData.Filter = ""
        grsData.Filter = "ҳ��>" & intToPage
        If grsData.RecordCount > 0 Then
            grsData.MoveFirst
            Do While Not grsData.EOF
                grsData("ҳ��").Value = grsData("ҳ��").Value + intListPage
                grsData.MoveNext
            Loop
        End If
        
        '��ǰ�ƶ�
        '------------------------------------------------------
        grsPage.Filter = ""
        grsPage.Filter = "ҳ��<0"
        grsPage.Sort = "ҳ�� Desc"
        If grsPage.RecordCount > 0 Then
            grsPage.MoveFirst
            Do While Not grsPage.EOF
                intLoop = intLoop + 1
                intListPage = grsPage("ҳ��").Value
                grsPage("ҳ��").Value = intToPage + intLoop
                grsData.Filter = ""
                grsData.Filter = "ҳ��=" & intListPage
                If grsData.RecordCount > 0 Then
                    grsData.MoveFirst
                    Do While Not grsData.EOF
                        grsData("ҳ��").Value = intToPage + intLoop
                        grsData.MoveNext
                    Loop
                End If
                
                grsPage.MoveNext
            Loop
        End If

        grsPage.Filter = ""
        grsPage.Sort = "ҳ��"
        
    End If
    
    AdjustListPage = True

    Exit Function

    '��������
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function AppendList(ByVal strListName As String, ByVal bytList As Byte, ByVal intPage As Integer) As Boolean
    '******************************************************************************************************************
    '���ܣ����ӵ�Ŀ¼����
    '������
    '���أ�
    '******************************************************************************************************************
    
    AppendList = AppendListData(strListName, bytList, intPage)
    
End Function

Public Property Get GetLastStartPage() As Integer
    '
    GetLastStartPage = gintStartPage
    
End Property

Private Sub Class_Initialize()
    Set grsData = New ADODB.Recordset
    Set grsPage = New ADODB.Recordset
    Set grsList = New ADODB.Recordset
    Set mclsLisReportLib = New clsLisReportLib
    Set mclsLisPackage = New clsLisPackage
    
    ParamInfo.ϵͳ���� = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    ParamInfo.��Ʒ���� = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")

    mProperty.�������� = "����"
    mProperty.������� = False
    mProperty.����б�� = False
    mProperty.�����С = 9
    mProperty.�����»��� = False
    mProperty.����ǰ��ɫ = 0
    
End Sub

Public Sub DestoryDrawReport()
    
    Dim frmThis As Form
    
    On Error Resume Next
    
    Set grsData = Nothing
    Set grsPage = Nothing
    Set grsList = Nothing
    Set mclsLisReportLib = Nothing
    Set mclsLisPackage = Nothing
    Set gcnOracle = Nothing
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    
End Sub

Private Sub Class_Terminate()
    Set grsData = Nothing
    Set grsPage = Nothing
    Set grsList = Nothing
    Set mclsLisReportLib = Nothing
    Set gcnOracle = Nothing
    
    'ɾ����ʱ�ļ�
    If Not (grsTempFile Is Nothing) Then
        If grsTempFile.RecordCount > 0 Then
            grsTempFile.MoveFirst
            Do While Not grsTempFile.EOF
                If grsTempFile("�ļ�").Value <> "" Then
                    If Dir(grsTempFile("�ļ�").Value) <> "" Then
                        On Error Resume Next
                        Kill grsTempFile("�ļ�").Value
                        On Error GoTo 0
                    End If
                End If
                grsTempFile.MoveNext
            Loop
        End If
    End If
    
    Set grsTempFile = Nothing
End Sub

Private Sub mfrmListReportView_AfterPrinted()
    RaiseEvent AfterPrinted
End Sub