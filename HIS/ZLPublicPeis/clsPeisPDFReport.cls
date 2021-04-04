VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPeisPDFReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjPeisReport As Object
Private mblnPrintOver As Boolean
Private mblnPrintDept As Boolean
Private mblnPrintSum As Boolean
Private mblnPrintEmpty As Boolean
Private mblnPrintQuestion As Boolean
Private mstrSQL As String
Private mstrTitle As String

'######################################################################################################################
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'���ܣ����ָ����Ա��PDF��챨�浥
'������
'���أ�
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function Initialize(ByVal lngSys As Long, Optional ByVal lngModual As Long = 2121)
    
    mblnPrintOver = (Val(gobjDatabase.GetPara("��ӡ����", lngSys, lngModual, "1")) >= 1)
    mblnPrintDept = (Val(gobjDatabase.GetPara("��ӡ��Ŀ", lngSys, lngModual, "1")) >= 1)
    mblnPrintSum = (Val(gobjDatabase.GetPara("��ӡ�ܼ�", lngSys, lngModual, "1")) >= 1)
    mblnPrintEmpty = (Val(gobjDatabase.GetPara("��ӡ����", lngSys, lngModual, "1")) >= 1)
    mblnPrintQuestion = (Val(gobjDatabase.GetPara("��ӡ�ʾ�", lngSys, lngModual, "1")) >= 1)
                                                     
    Call SaveSetting("ZLSOFT", "����ģ��\zl9PeisComLib\���˽�����챨��", "��ӡ��", Printer.DeviceName)
    
    
    Initialize = True
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'���ܣ����ָ����Ա��PDF��챨�浥
'������lngTaskKey-����id;lngPersonKey-����id;lngReportID-����Ŀ¼id;strPath-����ļ�����·��
'���أ�
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function ExportPDFReport(ByVal lngTaskKey As Long, ByVal lngPersonKey As Long, ByVal lngReportID As Long, Optional ByVal strPath As String) As Boolean
    Dim rsStyle As New ADODB.Recordset
    Dim objFile As New FileSystemObject
    Dim strFile As String
    Dim varStr As Variant
    Dim lngIndex As Long
    
     If InStrRev(strPath, ".") = 0 Then
        MsgBox "��" & strPath & "���ļ�·������", vbOKOnly + vbExclamation, "��ʾ"
        Exit Function
    End If

    If Trim(strPath) <> "" Then
        varStr = Split(Left(strPath, InStrRev(strPath, "\", Len(strPath)) - 1), "\")
        For lngIndex = LBound(varStr) To UBound(varStr)
            strFile = IIf(Trim(strFile) = "", strFile, strFile & "\") & varStr(lngIndex)
            If objFile.FolderExists(strFile) = False Then
                Call objFile.CreateFolder(strFile)
            End If
        Next
    End If
    
    Set rsStyle = Get_Peistasktypeitem(lngReportID)
    If rsStyle.BOF = False Then
        If mobjPeisReport Is Nothing Then
            Set mobjPeisReport = CreateObject("zl9PeisComLib.clsPeisReport")
            Call mobjPeisReport.InitCommon(gcnOracle)
        End If
        Select Case Val(rsStyle("��������").Value)
        Case 3                  'ר����챨��
            ExportPDFReport = OutputSpecialReport(rsStyle, lngTaskKey, lngPersonKey, strPath)
        Case Else               '��ͨ��챨��
            
            ExportPDFReport = OutputNormalReport(Val(rsStyle("ID").Value), lngTaskKey, lngPersonKey, strPath)
        
        End Select
            
    End If
End Function

'######################################################################################################################
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'���ܣ����ָ����Ա����챨�浥
'������
'���أ�
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Function OutputNormalReport(ByVal lngReportStyleKey As Long, ByVal lngTaskKey As Long, ByVal lngPersonKey As Long, Optional ByVal strFile As String)

    If mobjPeisReport.InitReport(gcnOracle, lngReportStyleKey) Then
        If mobjPeisReport.GetPersonReportData(lngTaskKey, lngPersonKey, Me, mblnPrintOver, 5) Then
        
            If strFile <> "" Then
                
                OutputNormalReport = mobjPeisReport.ExportReport(3, strFile)
            End If
        End If
    End If
            
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'���ܣ����ָ����Ա����챨�浥
'������
'���أ�
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Function OutputSpecialReport(ByVal rsReportStyle As ADODB.Recordset, ByVal lngTaskKey As Long, ByVal lngPatientKey As Long, ByVal strFile As String) As Boolean


    Dim blnNoPrint As Boolean
    Dim strTmp As String
    Dim varTmp As Variant
    Dim intCount As Integer
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim str�����ʽ As String
    Dim rsData As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    Dim strClassName As String
    Dim str������� As String
    Dim lngCount As Long
    Dim strUrlStart As String
    Dim aryProperty As Variant
    Dim intLoop As Integer
    Dim intPos As Integer
    
    On Error GoTo errHand
        
    
    Set rsData = New ADODB.Recordset
    With rsData
        .Fields.Append "����", adTinyInt
        .Fields.Append "ID", adBigInt
        .Fields.Append "��ʽ", adVarChar, 50
        .Fields.Append "����", adTinyInt
        .Fields.Append "�ܼ�", adTinyInt
        .Fields.Append "��Ŀ", adTinyInt
        .Open
    End With
    
    If rsReportStyle.RecordCount > 0 Then
        rsReportStyle.MoveFirst
        
        Set rsCondition = CreateCondition
        
        Call SetCondition(rsCondition, "�������id", gobjComFun.NVL(rsReportStyle("�������id").Value, 0))

        Set rsTmp = GetPeisComponet(2, rsCondition)
        If rsTmp.BOF = False Then
            strClassName = UCase(gobjComFun.NVL(rsTmp("����").Value))
            str������� = Trim(gobjComFun.NVL(rsTmp("�������").Value))
                
                aryProperty = Split(str�������, "'")
                For intLoop = 0 To UBound(aryProperty)
                    strTmp = CStr(aryProperty(intLoop))
                    intPos = InStr(strTmp, "=")
                    If intPos > 0 Then
                        Select Case Trim(Left(strTmp, intPos - 1))
                        Case "������ʽ"
                            str�����ʽ = Trim(Mid(strTmp, intPos + 1))
                        Case "�������"
                            str�����ʽ = Trim(Mid(strTmp, intPos + 1))
                        End Select
                    End If
                Next
                
            Select Case strClassName
            '--------------------------------------------------------------------------------------------------
            Case "ZL21_BILL_2121_4_1"
                rsData.AddNew
                rsData("����").Value = 2
                rsData("ID").Value = 0
                rsData("��ʽ").Value = str�����ʽ
                rsData("����").Value = 1
                rsData("�ܼ�").Value = 0
                rsData("��Ŀ").Value = 0
            '--------------------------------------------------------------------------------------------------
            Case "ZL21_BILL_2121_4_3"
                rsData.AddNew
                rsData("����").Value = 2
                rsData("ID").Value = 0
                rsData("��ʽ").Value = str�����ʽ
                rsData("����").Value = 0
                rsData("�ܼ�").Value = 1
                rsData("��Ŀ").Value = 0
            '--------------------------------------------------------------------------------------------------
            Case "ZL21_BILL_2121_4_2"
                rsData.AddNew
                rsData("����").Value = 2
                rsData("ID").Value = 0
                rsData("��ʽ").Value = str�����ʽ
                rsData("����").Value = 0
                rsData("�ܼ�").Value = 0
                rsData("��Ŀ").Value = 1
            Case Else
                rsData.AddNew
                rsData("����").Value = 1
                rsData("ID").Value = gobjComFun.NVL(rsReportStyle("ID").Value, 0)
                rsData("��ʽ").Value = strClassName
                rsData("����").Value = 1
                rsData("�ܼ�").Value = 0
                rsData("��Ŀ").Value = 0
            End Select
        End If
        
        Call SetCondition(rsCondition, "����Ŀ¼id", Val(rsReportStyle("����Ŀ¼id").Value))
        Set rsTmp = GetPeisReportContent(3, rsCondition)
        If rsTmp.BOF = False Then
            Do While Not rsTmp.EOF
                strClassName = UCase(gobjComFun.NVL(rsTmp("����").Value))
                
                str������� = Trim(gobjComFun.NVL(rsTmp("�������").Value))
                
                aryProperty = Split(str�������, "'")
                For intLoop = 0 To UBound(aryProperty)
                    strTmp = CStr(aryProperty(intLoop))
                    intPos = InStr(strTmp, "=")
                    If intPos > 0 Then
                        Select Case Trim(Left(strTmp, intPos - 1))
                        Case "������ʽ"
                            str�����ʽ = Trim(Mid(strTmp, intPos + 1))
                        Case "�������"
                            str�����ʽ = Trim(Mid(strTmp, intPos + 1))
                        Case "����·��ǰ׺"
                            If str�����ʽ = "��Ŀ" Then
                                strUrlStart = Trim(Mid(strTmp, intPos + 1))
                            End If
                        End Select
                    End If
                Next
                
                Select Case strClassName
                '--------------------------------------------------------------------------------------------------
                Case "ZL21_BILL_2121_4_3"
                
                    rsData.Filter = ""
                    rsData.Filter = "����=2 And ��ʽ='" & str�����ʽ & "'"
                    If rsData.RecordCount = 0 Then
    
                        rsData.AddNew
                        rsData("����").Value = 2
                        rsData("ID").Value = 0
                        rsData("��ʽ").Value = str�����ʽ
                        rsData("����").Value = 0
                        rsData("�ܼ�").Value = 1
                        rsData("��Ŀ").Value = 0
    
                    Else
                        rsData("�ܼ�").Value = 1
                    End If

                '--------------------------------------------------------------------------------------------------
                Case "ZL21_BILL_2121_4_2"
                    rsData.Filter = ""
                    rsData.Filter = "����=2 And ��ʽ='" & str�����ʽ & "'"
                    If rsData.RecordCount = 0 Then
    
                        rsData.AddNew
                        rsData("����").Value = 2
                        rsData("ID").Value = 0
                        rsData("��ʽ").Value = str�����ʽ
                        rsData("����").Value = 0
                        rsData("�ܼ�").Value = 0
                        rsData("��Ŀ").Value = 1
    
                    Else
                        rsData("��Ŀ").Value = 1
                    End If
                '--------------------------------------------------------------------------------------------------
                Case Else
                    rsData.Filter = ""
                    rsData.Filter = "����=1 And ��ʽ='" & strClassName & "'"

                    Select Case str�����ʽ
                    Case "�ʾ�"
                        If rsData.RecordCount = 0 Then
                            rsData.AddNew
                            rsData("����").Value = 1
                            rsData("ID").Value = gobjComFun.NVL(rsReportStyle("ID").Value, 0)
                            rsData("��ʽ").Value = strClassName
                            rsData("����").Value = 0
                            rsData("�ܼ�").Value = 0
                            rsData("��Ŀ").Value = 1
                        Else
                            rsData("��Ŀ").Value = 1
                        End If
                    Case "�ܼ�"
                        If rsData.RecordCount = 0 Then
                            rsData.AddNew
                            rsData("����").Value = 1
                            rsData("ID").Value = gobjComFun.NVL(rsReportStyle("ID").Value, 0)
                            rsData("��ʽ").Value = strClassName
                            rsData("����").Value = 0
                            rsData("�ܼ�").Value = 1
                            rsData("��Ŀ").Value = 0
                        Else
                            rsData("�ܼ�").Value = 1
                        End If
                    Case "��Ŀ"
                        If rsData.RecordCount = 0 Then
                            rsData.AddNew
                            rsData("����").Value = 1
                            rsData("ID").Value = gobjComFun.NVL(rsReportStyle("ID").Value, 0)
                            rsData("��ʽ").Value = strClassName
                            rsData("����").Value = 0
                            rsData("�ܼ�").Value = 0
                            rsData("��Ŀ").Value = 1
                        Else
                            rsData("��Ŀ").Value = 1
                        End If
                    End Select
                    
                End Select
                
                rsTmp.MoveNext
            Loop
        End If
        
        '--------------------------------------------------------------------------------------------------------------
        rsData.Filter = ""
        rsData.Sort = "���� Desc,�ܼ� Desc,��Ŀ Desc"
        If rsData.RecordCount > 0 Then rsData.MoveFirst
        

        lngCount = 0
        'ͬһ�������ڵı����ʽ��ͬ
        Do While Not rsData.EOF
            lngCount = lngCount + 1
            str�����ʽ = gobjComFun.NVL(rsData("��ʽ").Value)
            
            If Val(gobjComFun.NVL(rsData("����").Value)) = 1 Then
                '�̶�����
                If mobjPeisReport.InitReport(gcnOracle, Val(rsData("ID").Value)) Then
                    
                    Call mobjPeisReport.ExportReportSpec(lngTaskKey, lngPatientKey, (Val(rsData("����").Value) = 1 And mblnPrintOver), _
                                                            (Val(rsData("��Ŀ").Value) = 1 And mblnPrintDept), _
                                                            (Val(rsData("�ܼ�").Value) = 1 And mblnPrintSum), _
                                                            mblnPrintEmpty, 5, _
                                                            "", strFile, True, _
                                                            (mblnPrintQuestion And lngCount = 1), _
                                                            strUrlStart)
                End If

            Else
                Call mobjPeisReport.InitReport(gcnOracle, Val(rsData("ID").Value))
                intCount = intCount + 1
                
                '�Զ��屨��
                '------------------------------------------------------------------------------------------------------
                If Val(rsData("����").Value) = 1 And mblnPrintOver Then
                    Set rs = Get_Rptfmtno("ZL21_BILL_2121_4_1", str�����ʽ)


                    If rs.BOF = False Then str�����ʽ = rs("���").Value
                    
                    If Val(str�����ʽ) > 0 Then
                        Call gobjReport.ReportOpen(gcnOracle, glngSys, "ZL21_BILL_2121_4_1", Me, "����id=" & lngTaskKey, "����id=" & lngPatientKey, "REPORTFORMAT=" & Val(str�����ʽ), "PDF=" & strFile, "PrintEmpty=0", 4)
                    End If
                End If
                
                '------------------------------------------------------------------------------------------------------
                If Val(rsData("�ܼ�").Value) = 1 And mblnPrintSum Then
                    Set rs = Get_Rptfmtno("ZL21_BILL_2121_4_3", str�����ʽ)


                    If rs.BOF = False Then str�����ʽ = rs("���").Value
                    
                    If Val(str�����ʽ) > 0 Then
                    
                        Call gobjReport.ReportOpen(gcnOracle, glngSys, "ZL21_BILL_2121_4_3", Me, "����id=" & lngTaskKey, "����id=" & lngPatientKey, "REPORTFORMAT=" & Val(str�����ʽ), "PDF=" & strFile, "PrintEmpty=0", 4)
                    End If
                End If
                
                '------------------------------------------------------------------------------------------------------
                If Val(rsData("��Ŀ").Value) = 1 And mblnPrintDept Then

                    Set rs = Get_Rptfmtno("ZL21_BILL_2121_4_2", str�����ʽ)


                    If rs.BOF = False Then str�����ʽ = rs("���").Value
                    
                    If Val(str�����ʽ) > 0 Then
                        Call gobjReport.ReportOpen(gcnOracle, glngSys, "ZL21_BILL_2121_4_2", Me, "����id=" & lngTaskKey, "����id=" & lngPatientKey, "����=" & IIf(mblnPrintEmpty, 1, 0), "REPORTFORMAT=" & Val(str�����ʽ), "PDF=" & strFile, "PrintEmpty=0", 4)
                    End If
                    


                End If
                                
            End If
            
            rsData.MoveNext
        Loop
    End If
        
    OutputSpecialReport = True


    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function
Private Function Get_Peistasktypeitem(ByVal lngReportID As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ�
    '������
    '˵����
    '******************************************************************************************************************
    Dim rsData As ADODB.Recordset
    On Error GoTo errHand
    
    If lngReportID <> 0 Then
         mstrSQL = _
                "Select Distinct b.Id, b.��������, a.����Ŀ¼id, a.�������, b.�������, b.��������, b.�������id" & vbNewLine & _
                "From ������ͱ��� A, ��챨��Ŀ¼ B" & vbNewLine & _
                "Where a.����Ŀ¼id = b.Id And b.Id =[1]"
        
        Set rsData = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngReportID)
    End If
    
    Set Get_Peistasktypeitem = rsData

    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetPeisComponet(ByVal bytMode As Byte, ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    On Error GoTo errHand
    '�������
    Select Case bytMode
    Case 1
        mstrSQL = _
            "Select '���' As ͼ��,a.ID,Decode(a.�������,0,'����',1,'����',2,'����',3,'ר��') As �������,a.����Ԫ��id,a.�������,a.�������,a.�������,a.�������,a.Ĭ�Ϸ���,a.����˵�� " & vbNewLine & _
            "From ��챨����� a "
        Set GetPeisComponet = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle)
    Case 2
        mstrSQL = _
            "Select '���' As ͼ��,a.ID,Decode(a.�������,0,'����',1,'����',2,'����',3,'ר��') As �������,a.����Ԫ��id,b.����,a.�������,a.�������,a.�������,a.�������,a.Ĭ�Ϸ���,a.����˵�� " & vbNewLine & _
            "From ��챨����� a,��챨��Ԫ�� b Where a.ID=[1] And a.����Ԫ��id=b.ID"
        Set GetPeisComponet = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, GetCondition(rsCondition, "�������id"))
    End Select

    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetPeisReportContent(ByVal bytMode As Byte, ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    On Error GoTo errHand
    '��������
    Select Case bytMode
    Case 1
        mstrSQL = _
            "Select a.ID,Decode(a.��������,1,'�ı�','���') As ͼ��,Decode(a.��������,1,'�ı�','���') As ��������,a.����Ŀ¼id,a.�������,a.�������,a.��������,a.��ֹ���,a.��������,a.�ı�����,a.�ı�����Դ,a.�������id,a.��������,a.������ʾ,a.��ҳ���,a.Ŀ¼���� " & vbNewLine & _
            "From ��챨������ a Where a.����Ŀ¼id=[1] Order By a.������� "
        Set GetPeisReportContent = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, GetCondition(rsCondition, "����Ŀ¼id"))
    Case 2
        mstrSQL = _
            "Select A.ID,Decode(a.��������,1,'�ı�','���') As ͼ��,Decode(a.��������,1,'�ı�','���') As ��������,a.����Ŀ¼id,a.�������,a.�������,b.��������,a.��������,a.��ֹ���,a.��������,a.�ı�����,a.�ı�����Դ,a.�������id,a.��������,a.������ʾ,a.��ҳ���,a.Ŀ¼����,b.�������� " & vbNewLine & _
            "From ��챨������ a,��챨��Ŀ¼ b Where a.ID=[1] And a.����Ŀ¼id=b.ID"
        Set GetPeisReportContent = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, GetCondition(rsCondition, "��������id"))
    Case 3              '���˱����ӡ
        mstrSQL = _
            "Select a.ID,a.����Ŀ¼id,a.�������,a.��ֹ���,a.�������id,b.�������,c.���� " & vbNewLine & _
            "From ��챨������ a,��챨����� b,��챨��Ԫ�� c Where a.����Ŀ¼id=[1] And a.�������id=b.ID And b.����Ԫ��id=c.ID Order By a.������� "
        Set GetPeisReportContent = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, GetCondition(rsCondition, "����Ŀ¼id"))
    End Select

    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function Get_Rptfmtno(ByVal �������_In As String, ByVal ������ʽ_In As String) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ�
    '������
    '˵����
    '******************************************************************************************************************
    On Error GoTo errHand
    mstrSQL = _
        "Select a.��� from zlRPTFMTs a,zlReports b where a.����id=b.id and  b.���=[1] And a.˵��=[2]"
    Set Get_Rptfmtno = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, �������_In, ������ʽ_In)

    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function CreateCondition() As ADODB.Recordset

    Dim rs As New ADODB.Recordset

    With rs
        .Fields.Append "��������", adVarChar, 30
        .Fields.Append "�������", adVarChar, 4000
        .Fields.Append "��������", adVarChar, 30
        .Open
    End With

    Set CreateCondition = rs

End Function

Private Function SetCondition(ByRef rs As ADODB.Recordset, ByVal strConditionName As String, ByVal strConditionValue As String, Optional ByVal strConditionType As String = "�ı�") As Boolean

    rs.Filter = ""
    rs.Filter = "��������='" & strConditionName & "'"
    If rs.RecordCount = 0 Then rs.AddNew
    rs("��������").Value = strConditionName
    rs("�������").Value = strConditionValue
    rs("��������").Value = strConditionType
    SetCondition = True

End Function

Private Function GetCondition(ByRef rs As ADODB.Recordset, ByVal strConditionName As String) As String
    rs.Filter = ""
    rs.Filter = "��������='" & strConditionName & "'"
    If rs.RecordCount > 0 Then
        GetCondition = CStr(rs("�������").Value)
    End If
End Function



Private Sub Class_Terminate()
    Set mobjPeisReport = Nothing
End Sub