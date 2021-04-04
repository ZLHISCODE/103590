Attribute VB_Name = "mdlAboutReport"
Option Explicit

Public Function findThirdReport(ByVal lngSampleID As String, objWeb As WebBrowser)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsTemp As Recordset
    '����LIS����
    Dim strTag As String

    strSQL = "select ҽ��ID,����ID from ����������� where �걾ID=[1] and ҽ��ID is not null"
    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����������", lngSampleID)
    Do While Not rsTmp.EOF
        strSQL = "select b.id as ����ID,b.������,b.������||','||To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI') as �ĵ�����,c.ҽ��ID,b.����,b.��ӡ���� from ����ҽ����¼ a, ҽ���������� b,����ҽ������ c where b.id=c.����id and a.id=c.ҽ��id and c.����id is not null and b.���� in (0,2) and a.id =[1]" & vbCrLf & _
               " union all " & vbCrLf & _
               " select b.id as ����ID,b.������,b.������||','||To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI') as �ĵ�����,c.ҽ��ID,b.����,b.��ӡ���� from ����ҽ����¼ a, ҽ���������� b,����ҽ������ c where b.id=c.����id and a.id=c.ҽ��id and c.����id is not null and b.���� in (0,2) and a.id =[2]"

        Set rsTemp = OpenSQLRecord(Sel_His_DB, strSQL, "��������", Val(rsTmp("ҽ��ID") & ""), Val(rsTmp("����ID") & ""))
        If rsTemp.RecordCount > 0 Then
            strTag = strTag & "<SP>" & rsTemp!����ID & ";" & rsTemp!ҽ��id & ";" & rsTemp!���� & "<sTab>" & rsTemp!������
            Call WebShow(strTag, objWeb)
        End If
    Loop
    If strTag <> "" Then findThirdReport = Mid(strTag, 5)

End Function

Public Sub WebShow(ByVal strKey As String, objWeb As WebBrowser)
'���ܣ�Web�ؼ�չʾ�ļ�
    Dim strURL As String
    If strKey = "" Then
        Call objWeb.Navigate("about:blank")
        objWeb.Visible = False
'        mstrCurFile = ""
    Else
        strURL = GetLisRptFile(strKey)
        If strURL <> "" Then
            objWeb.Navigate strURL
'            mstrCurFile = strURL
        End If
        objWeb.Visible = True
    End If
End Sub

Public Function GetLisRptFile(ByVal strTag As String) As String
'���ܣ���LIS�����ļ��鿴����ȡ��ʱ�ļ�·��
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim strTmp As String
    Dim lng����ID As String
    Dim str������ As String
    Dim lng���� As String
    Dim varTmp As Variant
    Dim strSuffix As String '�ļ���׺��
    
    Screen.MousePointer = 11
    
    varTmp = Split(strTag, ";")
    lng����ID = varTmp(0)
    strTmp = Replace(strTag, varTmp(0) & ";" & varTmp(1) & ";", "")
    varTmp = Split(strTmp, "<sTab>")
    lng���� = varTmp(0)
    If lng���� = 0 Then
        strSuffix = "pdf"
    ElseIf lng���� = 1 Then
        strSuffix = "html"
    Else
        strSuffix = "xps"
    End If
    str������ = varTmp(1)
    
    strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\tmpReport_" & lng����ID & "." & strSuffix
    If Not objFile.FileExists(strFile) Then
        strFile = ReadLob(100, 22, lng����ID, strFile)
        If Not objFile.FileExists(strFile) Then
            MsgBox "�ļ����ݶ�ȡʧ�ܣ�", vbInformation, "������Ϣ":
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    GetLisRptFile = strFile
    Screen.MousePointer = 0
End Function

Public Function CalcVolatility(strCalcA As String, strCalcB As String) As String
    '���������

    On Error Resume Next

    If strCalcA = "" Or strCalcB = "" Then
        CalcVolatility = ""
        Exit Function
    End If
    If Val(strCalcA) = 0 Or Val(strCalcB) = 0 Then
        CalcVolatility = ""
    End If

    '����
    CalcVolatility = (Val(strCalcB) - Val(strCalcA)) / Val(strCalcA) * 100
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/8/27
'��    ��:��̬�����ؼ�
'��    ��:
'           objParent           ����������Ҫ���ĸ������д�������
'           strControlClass     ��Ҫ�����Ŀؼ�
'           strControlClass     �ؼ�����
'           [objPart            ��������Ҫ���Ŀؼ��������ĸ�������,Ĭ��Ϊ�������������]
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function NewControl(objParent As Object, ByVal strControlClass As String, ByVal strName As String, Optional objPart As Object) As Object
          Dim objCrl As Object
          
          '����Э�飬ֻ�ܼ�һ�Σ��ڶ��λ����
          

1         On Error Resume Next
2         Call Licenses.Add(strControlClass)
3         Err.Clear: On Error GoTo NewControl_Error
          '������̬�ؼ�
4         If objPart Is Nothing Then
5             Set objCrl = objParent.Controls.Add(strControlClass, strName)
6         Else
7             Set objCrl = objParent.Controls.Add(strControlClass, strName)
8             Set objCrl.Container = objPart
9             objCrl.Move 0, 0, objPart.Width, objPart.Height
10            objCrl.ZOrder
11            objCrl.Visible = True
12        End If
13        If strControlClass = "zlLisControl.ucLisIDKind" Then
14            If Not objCrl.object.InitControl(objParent, gcnLisOracle, gUserInfo.DBUser) Then
15                Exit Function
16            End If
17        End If
          
18        Set NewControl = objCrl
          


19        Exit Function
NewControl_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "mdlAboutReport", "ִ��(NewControl)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
21        Err.Clear

End Function

Public Function PtintOldReport(objFrm As Object, lngSampleID As Long, Optional lngPaintID As Long, Optional byRunMode As Byte = 2, Optional ByVal intSpecialPrintPage As Integer, Optional strErr As String) As Boolean
  '��ӡ�ϰ汨��
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim strSQL As String
    Dim strChart(0 To 8) As String

    On Error GoTo PtintOldReport_Error

    strSQL = "select ���ͺ�, a.ҽ��id from ����ҽ������ a , ����ҽ����¼ b,����걾��¼  c where b.id = a.ҽ��id and  a.ҽ��id =c.ҽ��id  and c.id = [1]"
    Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�����ӡ", lngSampleID)
    If rsTmp.EOF = False Then
        lng���ͺ� = Val("" & rsTmp("���ͺ�"))
        lngҽ��ID = Val("" & rsTmp("ҽ��id"))
    End If

    If GetReportCode(lngҽ��ID, lng���ͺ�, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
        If byRunMode = 3 Then
            FunReportPrintSetHis gcnHisOracle, 100, strReportCode, objFrm
        Else
            If ReadSampleImage(lngSampleID, strChart, strErr, 10) = False Then
                Exit Function
            End If
            Call FunReportOpenHis(gcnHisOracle, 100, strReportCode, objFrm, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & lngҽ��ID, _
                                "����ID=" & lngPaintID, "�걾ID=" & lngSampleID, _
                                "ͼ��1=" & strChart(0), "ͼ��2=" & strChart(1), "ͼ��3=" & strChart(2), "ͼ��4=" & strChart(3), _
                                "ͼ��5=" & strChart(4), "ͼ��6=" & strChart(5), "ͼ��7=" & strChart(6), "ͼ��8=" & strChart(7), _
                                "ͼ��9=" & strChart(8), "DisabledPrint=1", intSpecialPrintPage, byRunMode)
        End If
    End If
    PtintOldReport = True


    Exit Function
PtintOldReport_Error:
    Call WriteErrLog("zlPublicHisCommLis", "mdlAboutReport", "ִ��(PtintOldReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
    Err.Clear

End Function

Public Function PrintNewReport(objFrm As Object, lngSampleID As Long, Optional byRunMode As Byte = 2, Optional ByVal blnDoctorShow As Boolean, Optional ByVal strPrivs As String, Optional ByVal intSpecialPrintPage As Integer, Optional strErr As String) As Boolean
'����       ��ӡ����
    Dim intCount As Integer
    Dim strNO As String
    Dim intSel As Integer
    Dim strChart(0 To 8) As String
    Dim strSQL As String
    Dim strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim rsReportFormat As ADODB.Recordset
    Dim lngPrintCount As Long

    On Error GoTo PrintNewReport_Error

    strSQL = "select b.id ����id ,b.���� ��������,b.�������,Nvl(a.������Դ,1) ������Դ,a.����ʱ��,a.���Ա���,a.�걾���,a.ҽ��վ��ӡ,�����  from ���鱨���¼ a,����������¼ b where a.����id = b.id and a.id = [1]"
    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����ӡ", lngSampleID)

    If rsTmp.RecordCount = 0 Then Exit Function
    
    
    '�Աȴ�ӡ�����Ͳ���
    If blnDoctorShow Then
        lngPrintCount = Val(ComGetPara(Sel_Lis_DB, "ҽ������վ�����ӡ����", 2500, 2500, 1))
        If lngPrintCount > 0 Then
            If Val(rsTmp("ҽ��վ��ӡ") & "") >= lngPrintCount And Val(rsTmp("������Դ") & "") = 2 Then
                strErr = "������ӡ������ֹ��ӡ"
                PrintNewReport = False
                Exit Function
            End If
        End If

    Else
        If rsTmp("�����") & "" = "" And byRunMode = 2 Then
            If InStr(";" & strPrivs & ";", ";δ��˱����ӡ;") Then
                strErr = "δ��˱��治�ܴ�ӡ"
                Exit Function
            End If
        End If
    End If
    
    strSQL = "select id,����,����,���ﵥ��,סԺ����,��쵥��,Ժ�ⵥ��,�����ʽ,סԺ��ʽ,����ʽ,Ժ���ʽ,��ʽ����," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(���ﵥ��, '00000')) || '-2' ���ﵥ�ݺ�," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(סԺ����, '00000')) || '-2' סԺ���ݺ�," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(��쵥��, '00000')) || '-2' ��쵥�ݺ�," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(Ժ�ⵥ��, '00000')) || '-2' Ժ�ⵥ�ݺ�," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(�����ʽ, '00000')) || '-2' �����ʽ��," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(סԺ��ʽ, '00000')) || '-2' סԺ��ʽ��," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(����ʽ, '00000')) || '-2' ����ʽ��," & vbNewLine & _
           "       'ZLLISBILL' || Trim(To_Char(Ժ���ʽ, '00000')) || '-2' Ժ���ʽ��" & vbNewLine & _
             "from ����������¼ where id = [1] "

    Set rsReportFormat = ComOpenSQL(Sel_Lis_DB, strSQL, "���鼼ʦվ", Val(rsTmp("����ID") & ""))


    rsReportFormat.Filter = "id=" & Val(rsTmp("����ID") & "")
    If Val(rsTmp("�������")) = 1 Then
        If Val(rsTmp("���Ա���") & "") = 1 Then
            '����
            intSel = 0
        Else
            '����
            intSel = 1
        End If
    Else
        intCount = GetSampleValCount(lngSampleID)
        'û�н��ʱ��ʾ
        If intCount = 0 Then
            Exit Function
        End If
        If rsReportFormat.RecordCount > 0 Then
            If Val(rsReportFormat("��ʽ����") & "") > 0 Then
                If intCount > Val(rsReportFormat("��ʽ����") & "") Then
                    intSel = 0
                Else
                    intSel = 1
                End If
            End If
        Else
            intSel = 0
        End If
    End If

    Select Case Val(rsTmp("������Դ"))
    Case 1
        If intSel = 0 Then
            strNO = rsReportFormat("���ﵥ�ݺ�")
        Else
            strNO = rsReportFormat("�����ʽ��")
        End If
    Case 2
        If intSel = 0 Then
            strNO = rsReportFormat("סԺ���ݺ�")
        Else
            strNO = rsReportFormat("סԺ��ʽ��")
        End If
    Case 3
        If intSel = 0 Then
            strNO = rsReportFormat("סԺ���ݺ�")
        Else
            strNO = rsReportFormat("סԺ��ʽ��")
        End If
    Case 4
        If intSel = 0 Then
            strNO = rsReportFormat("Ժ�ⵥ�ݺ�")
        Else
            strNO = rsReportFormat("Ժ���ʽ��")
        End If
    Case Else
        If intSel = 0 Then
            strNO = rsReportFormat("���ﵥ�ݺ�")
        Else
            strNO = rsReportFormat("�����ʽ��")
        End If
    End Select

    If byRunMode = 3 Then
        If strNO <> "" Then
            FunReportPrintSet gcnLisOracle, gSysInfo.SysNo, strNO, objFrm
        End If
    Else
        '��ͼ��
        strTmp = "��ʼ����ͼ��:" & Now & vbCrLf
        If ReadSampleImage(lngSampleID, strChart, strErr, 25) = False Then
            Exit Function
        End If
        strTmp = strTmp & "����ͼ�����:" & Now & vbCrLf

        FunReportOpen gcnLisOracle, gSysInfo.SysNo, strNO, objFrm, "�걾ID=" & lngSampleID, "ͼ��1=" & strChart(0), "ͼ��2=" & strChart(1), "ͼ��3=" & strChart(2), _
                      "ͼ��4=" & strChart(3), "ͼ��5=" & strChart(4), "ͼ��6=" & strChart(5), "ͼ��7=" & strChart(6), "ͼ��8=" & strChart(7), _
                      "ͼ��9=" & strChart(8), "DisabledPrint=1", intSpecialPrintPage, byRunMode
        strTmp = strTmp & "��ӡ���:" & Now & vbCrLf

        '������˹��ı걾��ʶ
        strSQL = "Zl_���鱨���ӡ_Edit(1," & lngSampleID & ",1)"
        Call ComExecuteProc(Sel_Lis_DB, strSQL, "��ӡ�걾")
        strTmp = strTmp & "��ɴ�ӡ:" & Now

        SaveDBLog 18, 6, lngSampleID, "��ӡ", "�����ӡ", 2500, "�ٴ�ʵ���ҹ���"
    End If

    PrintNewReport = True

    '����ˢ�¿��ڸſ��Ѵ�ӡ��ǩ����
    Call SendMessage("RefreshDeptSurvey7")


    Exit Function
PrintNewReport_Error:
    Call WriteErrLog("zlPublicHisCommLis", "mdlAboutReport", "ִ��(PrintNewReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
    Err.Clear

End Function

