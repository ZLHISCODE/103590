Attribute VB_Name = "mdSysConfig"
Option Explicit
'--------------------------------------------------------
'��  �ܣ�ϵͳ��������
'�����ˣ��ƽ�
'�������ڣ�2004.6.12
'���̺����嵥��
'        subGetWWWLToVal()��            ��ȡ��Ԥ�贰��λ�������ݵ�ϵͳ����
'        subGetLayoutToVar������        �����ݿ��л�ȡԤ����Ļ���֣���д��ϵͳ�����С���ȡ��Ԥ����Ļ���֡���
'        subSaveScreenLayout������      ���޸Ĺ�����Ļ���ֱ��浽ϵͳ���������ݿ��У���ϵͳ���������ݱ��浽"Ԥ����Ļ����"���С�
'        subGetMouseUsageToVar������    �����ݿ��ж�ȡ����÷����õ�ֵ��ϵͳ��������ȡ����갴ť���䡱������ݵ�ϵͳ����
'        subGetInfoLabelToVar������     �����ݿ��ȡ��Ϣ��עλ���������ݵ�ϵͳ��������ȡ��ͼ����Ϣ�������ݵ�ϵͳ����
'        subGetDBDicomPrintToVar������  �����ݿ��ȡ��ӡ���Ĳ�������д��ϵͳ�����������棬��ȡ��DICOM��ӡ�����á���
'        subGetInterfaceParaToVar������ �����ݿ��ȡ��Ӱ���������������ݣ������䱣�浽ϵͳ�����С�
'        LoadBarSetup():                ��ȡ���ݿ��ϴα���Ĺ���������
'�޸ļ�¼��
'    2005.06.29     �ƽ�    �����������ƶ���mdlSystemCortrolģ�飬���޸ġ�Ӱ���������������ݿ��
'-------------------------------------------------------

Public Sub subGetFilterToVal()
'------------------------------------------------
'���ܣ���ȡ��Ӱ���˾�ģ�塱�����ݵ�ϵͳ����
'��������
'���أ��ޣ�ֱ���޸�ϵͳ����
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo err
    
    strSQL = "Select Id,Ӱ������,�˾�����,��ǿǿ������,��ǿǿ�ȼ���,��ǿ��������,��ǿ���ȼ���, " _
        & " ƽ������,ƽ������ From Ӱ���˾�ģ�� order by Ӱ������, ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡӰ���˾�")
    
    '��ʼ��Ԥ���˾�ϵͳ����
    ReDim aPresetFilter(Val(rsTemp.RecordCount)) As TPresetFilter
    
    '��ȡԤ����˾�����
    i = 0
    With rsTemp
        While Not .EOF
            aPresetFilter(i).lngID = rsTemp!Id
            aPresetFilter(i).strname = rsTemp!�˾�����
            aPresetFilter(i).strModality = rsTemp!Ӱ������
            aPresetFilter(i).intUnSharpEnhancementUp = Nvl(rsTemp!��ǿǿ������, 0)
            aPresetFilter(i).intUnSharpEnhancementDown = Nvl(rsTemp!��ǿǿ�ȼ���, 0)
            aPresetFilter(i).intUnSharpLengthUp = Nvl(rsTemp!��ǿ��������, 0)
            aPresetFilter(i).intUnSharpLengthDown = Nvl(rsTemp!��ǿ���ȼ���, 0)
            aPresetFilter(i).intFilterLengthUp = Nvl(rsTemp!ƽ������, 0)
            aPresetFilter(i).intFilterLengthDown = Nvl(rsTemp!ƽ������, 0)
            i = i + 1
            .MoveNext
        Wend
    End With
    Exit Sub
err:
    If ErrCenter() = 1 Then
            Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subGetWWWLToVal()
'------------------------------------------------
'���ܣ���ȡ��Ԥ�贰��λ�������ݵ�ϵͳ����
'��������
'���أ��ޣ�ֱ���޸�ϵͳ����
'------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    Dim strModality As String           '���浱ǰ��Ӱ�����
    Dim lngModalityCount As Long
    Dim blnUseDefaultSet As Boolean     '�Ƿ�ʹ��Ĭ������
    
    strModality = ""        '��ʼ����ǰӰ�����
    blnUseDefaultSet = False
    
    '����Ӱ�����͵�����
    If blLocalRun = True Then
        strSQL = "SELECT COUNT(Ӱ������) as iCount FROM (SELECT DISTINCT Ӱ������ FROM Ӱ��Ԥ�贰��λ)"
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT COUNT( Distinct Ӱ������) as iCount FROM Ӱ��Ԥ�贰��λ where ��Աid =[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, glngUserID)
        If rsTemp!iCount = 0 Then
            blnUseDefaultSet = True
            strSQL = "SELECT COUNT( Distinct Ӱ������) as iCount FROM Ӱ��Ԥ�贰��λ where ��Աid =[1] "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, CLng(0))
        End If
    End If
    lngModalityCount = rsTemp!iCount
    
    ''''''''��ʼ��[����λ����]��ϵͳ����''''''''''''''''''''''''
    ReDim aPresetWinWL(3 To 12, lngModalityCount) As TPresetWinWL        ''����Ԥ�贰��λ�����飬
                     ''����Ŀ�ݼ�ֵΪF3--F12����Ӧ��������±�,2Ϊ�Զ�����λ
    
    '�����ݿ����ݱ��浽ϵͳ����������
    '���洰��λ������
    
    On Error GoTo errh
    
    If blLocalRun = True Then
        strSQL = "SELECT ID,Ӱ������,��ݼ�,��������,����Ӣ����,����,��λ,�Ƿ�Ĭ�� FROM Ӱ��Ԥ�贰��λ ORDER BY Ӱ������,��ݼ�"
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT ID,Ӱ������,��ݼ�,��������,����Ӣ����,����,��λ,�Ƿ�Ĭ�� FROM Ӱ��Ԥ�贰��λ " & _
                 " where ��Աid =[1] ORDER BY Ӱ������,��ݼ�"
        If blnUseDefaultSet = True Then
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, CLng(0))
        Else
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, glngUserID)
        End If
    End If
    i = 0
    With rsTemp
        Do While Not .EOF
            If strModality = "" Or strModality <> !Ӱ������ Then
                strModality = !Ӱ������
                i = i + 1
                aPresetWinWL(3, i).strModality = strModality
            End If
            aPresetWinWL(!��ݼ�, i).bInUse = True
            aPresetWinWL(!��ݼ�, i).intDefault = !�Ƿ�Ĭ��
            aPresetWinWL(!��ݼ�, i).strModality = strModality
            aPresetWinWL(!��ݼ�, i).strWinWLCName = !��������
            aPresetWinWL(!��ݼ�, i).strWinWLEName = !����Ӣ����
            aPresetWinWL(!��ݼ�, i).lngWinWidth = !����
            aPresetWinWL(!��ݼ�, i).lngWinLevel = !��λ
            aPresetWinWL(!��ݼ�, i).lngID = !Id
            .MoveNext
        Loop
    End With
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subGetLayoutToVar(lngUserID As Long)
'------------------------------------------------
'���ܣ������ݿ��л�ȡԤ����Ļ���֣���д��ϵͳ�����С���ȡ��Ԥ����Ļ���֡���
'��������
'���أ���
'�ϼ���������̣�
'�¼���������̣�
'���õ��ⲿ������
'�����ˣ��ƽ�
'------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    Dim intModalityCount As Integer
    Dim rsTmp As New ADODB.Recordset
    
    
    '����Ӱ�����͵�����
    If blLocalRun = True Then
        strSQL = "SELECT COUNT(Ӱ������) as iCount FROM (SELECT DISTINCT Ӱ������ FROM Ӱ����Ļ����)"
        Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT COUNT(Ӱ������) as iCount FROM Ӱ����Ļ���� where ��ԱID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngUserID)
    End If
    intModalityCount = rsTmp!iCount
    
    ReDim aModifiedPresetLayout(intModalityCount) As TModifiedPresetLayout
    ReDim aPresetLayout(intModalityCount) As TModifiedPresetLayout         ''����Ԥ����Ļ���ֵ�����
    
    '�����ݿ����ݱ��浽ϵͳ����������
    If blLocalRun = True Then
        strSQL = "SELECT Ӱ������,�Զ����в���,�Զ�ͼ�񲼾�,��������,��������,ͼ������,ͼ������" & _
                ",�Զ�����,��ʾ������Ϣ,ѡ��λ��,ѡ������ͬ��,��ֵģʽ,ͼ������ FROM Ӱ����Ļ���� ORDER BY Ӱ������"
        Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT Ӱ������,�Զ����в���,�Զ�ͼ�񲼾�,��������,��������,ͼ������,ͼ������" & _
                ",�Զ�����,��ʾ������Ϣ,ѡ��λ��,ѡ������ͬ��,��ֵģʽ,ͼ������ FROM Ӱ����Ļ���� Where ��ԱID = [1] ORDER BY Ӱ������"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngUserID)
    End If
    i = 1
    With rsTmp
        While Not .EOF
               '�����ݿ����ݱ��浽ϵͳ����������
               aPresetLayout(i).strModality = !Ӱ������
               aPresetLayout(i).bSeriesAutoFormat = IIf(IsNull(!�Զ����в���), 0, !�Զ����в���)
               aPresetLayout(i).lngSeriesColumns = IIf(IsNull(!��������), 2, !��������)
               aPresetLayout(i).lngSeriesRows = IIf(IsNull(!��������), 1, !��������)
               aPresetLayout(i).bImageAutoFormat = IIf(IsNull(!�Զ�ͼ�񲼾�), 0, !�Զ�ͼ�񲼾�)
               aPresetLayout(i).lngImageColumns = IIf(IsNull(!ͼ������), 2, !ͼ������)
               aPresetLayout(i).lngImageRows = IIf(IsNull(!ͼ������), 1, !ͼ������)
               aPresetLayout(i).bInvert = IIf(IsNull(!�Զ�����), 0, !�Զ�����)
               aPresetLayout(i).bShowPatientInfo = IIf(IsNull(!��ʾ������Ϣ), 0, !��ʾ������Ϣ)
               aPresetLayout(i).bAutoSelectReferenceLine = IIf(IsNull(!ѡ��λ��), 0, !ѡ��λ��)
               aPresetLayout(i).bAutoSelectSeriesSyn = IIf(IsNull(!ѡ������ͬ��), 0, !ѡ������ͬ��)
               aPresetLayout(i).lngInterpolationMode = IIf(IsNull(!��ֵģʽ), 0, !��ֵģʽ)
               aPresetLayout(i).lngImageSort = IIf(IsNull(!ͼ������), 0, !ͼ������)
               i = i + 1
               .MoveNext
        Wend
    End With
End Sub

Public Sub subGetImageShutterToVar(lngUserID As Long)
    '------------------------------------------------
'���ܣ������ݿ��л�ȡԤ��ͼ����������д��ϵͳ�����С���ȡ��ͼ����������
'��������
'���أ���
'�ϼ���������̣�
'�¼���������̣�
'���õ��ⲿ������
'�����ˣ��ƽ�
'------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    Dim intModalityCount As Integer
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errh
    
    If blLocalRun = True Then
        strSQL = "select count(Ӱ������) as iCount from Ӱ��ͼ�������� "
        Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "select count(Ӱ������) as iCount from Ӱ��ͼ�������� where ��ԱID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngUserID)
    End If
    
    intModalityCount = rsTmp!iCount
    
    ReDim aModifiedImageShutter(intModalityCount) As TImageShutter
    ReDim aImageShutter(intModalityCount) As TImageShutter          ''����ͼ������������
    
    If blLocalRun = True Then
        strSQL = "SELECT Ӱ������,��������,Բ��X,Բ��Y,Բ�ΰ뾶,������߽�,�����ұ߽�,�����ϱ߽�" & _
                ",�����±߽�,����ζ���,������ɫ FROM Ӱ��ͼ��������  ORDER BY Ӱ������"
        Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        '�����ݿ����ݱ��浽ϵͳ����������
        strSQL = "SELECT Ӱ������,��������,Բ��X,Բ��Y,Բ�ΰ뾶,������߽�,�����ұ߽�,�����ϱ߽�" & _
                ",�����±߽�,����ζ���,������ɫ FROM Ӱ��ͼ��������  where ��Աid = [1] ORDER BY Ӱ������"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngUserID)
    End If
    i = 1
    With rsTmp
        While Not .EOF
               '�����ݿ����ݱ��浽ϵͳ����������
               aImageShutter(i).strModality = !Ӱ������
               aImageShutter(i).intShutterType = IIf(IsNull(!��������), 0, !��������)
               aImageShutter(i).intCenterX = IIf(IsNull(!Բ��X), 0, !Բ��X)
               aImageShutter(i).intCenterY = IIf(IsNull(!Բ��Y), 0, !Բ��Y)
               aImageShutter(i).intRadius = IIf(IsNull(!Բ�ΰ뾶), 0, !Բ�ΰ뾶)
               aImageShutter(i).intRectLeft = IIf(IsNull(!������߽�), 0, !������߽�)
               aImageShutter(i).intRectRight = IIf(IsNull(!�����ұ߽�), 0, !�����ұ߽�)
               aImageShutter(i).intRectUpper = IIf(IsNull(!�����ϱ߽�), 0, !�����ϱ߽�)
               aImageShutter(i).intRectLower = IIf(IsNull(!�����±߽�), 0, !�����±߽�)
               aImageShutter(i).strVertices = IIf(Not IsNull(!����ζ���), !����ζ���, "")
               aImageShutter(i).lngColor = IIf(IsNull(!������ɫ), 0, !������ɫ)
               i = i + 1
               .MoveNext
        Wend
    End With
    
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subSaveImgShutter()
'------------------------------------------------
'���ܣ����޸Ĺ���ͼ�������������浽ϵͳ���������ݿ��У������������ݱ��浽"ͼ��������"���С�
'��������
'���أ�
'�ϼ���������̣�
'�¼���������̣�
'���õ��ⲿ������
'�����ˣ��ƽ�
'------------------------------------------------

    Dim i As Integer
    Dim strSQL As String
    Dim intModalityCount As Integer
    intModalityCount = UBound(aModifiedImageShutter)
    
    On Error GoTo errh
    
    For i = 1 To intModalityCount
        If aModifiedImageShutter(i).bModified Then
            '�����õĽ�����浽ϵͳ����
            aImageShutter(i).intCenterX = aModifiedImageShutter(i).intCenterX
            aImageShutter(i).intCenterY = aModifiedImageShutter(i).intCenterY
            aImageShutter(i).intRadius = aModifiedImageShutter(i).intRadius
            aImageShutter(i).intRectLeft = aModifiedImageShutter(i).intRectLeft
            aImageShutter(i).intRectLower = aModifiedImageShutter(i).intRectLower
            aImageShutter(i).intRectRight = aModifiedImageShutter(i).intRectRight
            aImageShutter(i).intRectUpper = aModifiedImageShutter(i).intRectUpper
            aImageShutter(i).intShutterType = aModifiedImageShutter(i).intShutterType
            aImageShutter(i).lngColor = aModifiedImageShutter(i).lngColor
            aImageShutter(i).strModality = aModifiedImageShutter(i).strModality
            aImageShutter(i).strVertices = aModifiedImageShutter(i).strVertices
            
            
            If blLocalRun = True Then
                '����ı��˵�ͼ���������õ����ݿ⣬���޸ļ�¼�н��б���
                strSQL = "UPDATE Ӱ��ͼ�������� SET �������� = " & aImageShutter(i).intShutterType _
                         & " , Բ��X = " & aImageShutter(i).intCenterX _
                         & " , Բ��Y = " & aImageShutter(i).intCenterY _
                         & " , Բ�ΰ뾶 = " & aImageShutter(i).intRadius _
                         & " , ������߽� = " & aImageShutter(i).intRectLeft _
                         & " , �����ұ߽� = " & aImageShutter(i).intRectRight _
                         & " , �����ϱ߽� = " & aImageShutter(i).intRectUpper _
                         & " , �����±߽� = " & aImageShutter(i).intRectLower _
                         & " , ����ζ��� = '" & aImageShutter(i).strVertices & "'" _
                         & " , ������ɫ = " & aImageShutter(i).lngColor _
                         & " where Ӱ������ = '" & aImageShutter(i).strModality & "'"
                cnAccess.Execute strSQL, , adCmdText
            Else
                strSQL = "ZL_Ӱ��ͼ��������_UPDATE(" & glngUserID & ",'" & aImageShutter(i).strModality & "','" & _
                aImageShutter(i).intShutterType & "'," & aImageShutter(i).intCenterX & "," & aImageShutter(i).intCenterY & _
                "," & aImageShutter(i).intRadius & "," & aImageShutter(i).intRectLeft & "," & aImageShutter(i).intRectRight & _
                "," & aImageShutter(i).intRectUpper & "," & aImageShutter(i).intRectLower & ",'" & aImageShutter(i).strVertices & _
                "'," & aImageShutter(i).lngColor & ")"
                zlDatabase.ExecuteProcedure strSQL, App.ProductName
            End If
        End If
    Next
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subSaveScreenLayout()
'------------------------------------------------
'���ܣ����޸Ĺ�����Ļ���ֱ��浽ϵͳ���������ݿ��У���ϵͳ���������ݱ��浽"Ԥ����Ļ����"���С�
'��������
'���أ�
'�ϼ���������̣�
'�¼���������̣�
'���õ��ⲿ������
'�����ˣ��ƽ�
'------------------------------------------------

    Dim i As Integer
    Dim strSQL As String
    Dim intModalityCount As Integer
    intModalityCount = UBound(aModifiedPresetLayout)
    
    On Error GoTo errh
    
    For i = 1 To intModalityCount
'        If aModifiedPresetLayout(i).bModified Then
            '�����õĽ�����浽ϵͳ����
            aPresetLayout(i).bSeriesAutoFormat = aModifiedPresetLayout(i).bSeriesAutoFormat
            aPresetLayout(i).lngSeriesColumns = aModifiedPresetLayout(i).lngSeriesColumns
            aPresetLayout(i).lngSeriesRows = aModifiedPresetLayout(i).lngSeriesRows
            aPresetLayout(i).bImageAutoFormat = aModifiedPresetLayout(i).bImageAutoFormat
            aPresetLayout(i).lngImageColumns = aModifiedPresetLayout(i).lngImageColumns
            aPresetLayout(i).lngImageRows = aModifiedPresetLayout(i).lngImageRows
            aPresetLayout(i).bInvert = aModifiedPresetLayout(i).bInvert
            aPresetLayout(i).bShowPatientInfo = aModifiedPresetLayout(i).bShowPatientInfo
            aPresetLayout(i).bAutoSelectReferenceLine = aModifiedPresetLayout(i).bAutoSelectReferenceLine
            aPresetLayout(i).bAutoSelectSeriesSyn = aModifiedPresetLayout(i).bAutoSelectSeriesSyn
            aPresetLayout(i).lngInterpolationMode = aModifiedPresetLayout(i).lngInterpolationMode
            aPresetLayout(i).lngImageSort = aModifiedPresetLayout(i).lngImageSort
            
            If blLocalRun = True Then
                '����ı��˵Ĳ��ֵ����ݿ⣬���޸ļ�¼�н��б���
                strSQL = "UPDATE Ӱ����Ļ���� SET �Զ����в���=" & _
                         IIf(aModifiedPresetLayout(i).bSeriesAutoFormat, 1, 0) & _
                         ",�Զ�ͼ�񲼾� = " & IIf(aModifiedPresetLayout(i).bImageAutoFormat, 1, 0) & _
                         ",�������� = " & aModifiedPresetLayout(i).lngSeriesRows & ",�������� = " & _
                         aModifiedPresetLayout(i).lngSeriesColumns & ",ͼ������ = " & _
                         aModifiedPresetLayout(i).lngImageRows & ",ͼ������ = " & _
                         aModifiedPresetLayout(i).lngImageColumns & ",�Զ�����= " & _
                         aModifiedPresetLayout(i).bInvert & ",��ʾ������Ϣ= " & _
                         aModifiedPresetLayout(i).bShowPatientInfo & ",ѡ��λ�� = " & _
                         aModifiedPresetLayout(i).bAutoSelectReferenceLine & ",ѡ������ͬ��=" & _
                         aModifiedPresetLayout(i).bAutoSelectSeriesSyn & ",��ֵģʽ=" & _
                         aModifiedPresetLayout(i).lngInterpolationMode & ",ͼ������=" & _
                         aModifiedPresetLayout(i).lngImageSort & " WHERE Ӱ������='" & _
                         aModifiedPresetLayout(i).strModality & "'"
                cnAccess.Execute strSQL, , adCmdText
            Else
                strSQL = "ZL_Ӱ����Ļ����_UPDATE(" & glngUserID & ",'" & aModifiedPresetLayout(i).strModality & "'," & _
                IIf(aModifiedPresetLayout(i).bSeriesAutoFormat, 1, 0) & "," & IIf(aModifiedPresetLayout(i).bImageAutoFormat, 1, 0) & _
                "," & aModifiedPresetLayout(i).lngSeriesRows & "," & aModifiedPresetLayout(i).lngSeriesColumns & _
                "," & aModifiedPresetLayout(i).lngImageRows & "," & aModifiedPresetLayout(i).lngImageColumns & "," & _
                CInt(aModifiedPresetLayout(i).bInvert) & "," & CInt(aModifiedPresetLayout(i).bShowPatientInfo) & _
                "," & CInt(aModifiedPresetLayout(i).bAutoSelectReferenceLine) & "," & CInt(aModifiedPresetLayout(i).bAutoSelectSeriesSyn) & _
                "," & aModifiedPresetLayout(i).lngInterpolationMode & "," & aModifiedPresetLayout(i).lngImageSort & ")"
                zlDatabase.ExecuteProcedure strSQL, App.ProductName
                
            End If
'        End If
    Next
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub


Public Sub subGetMouseUsageToVar(lngUserID As Long)
'------------------------------------------------
'���ܣ������ݿ��ж�ȡ����÷����õ�ֵ��ϵͳ��������ȡ����갴ť���䡱������ݵ�ϵͳ����
'��������
'���أ���
'�ϼ���������̣�frmViewer.Form_Load
'�¼���������̣���
'���õ��ⲿ������cMouseUsage
'�����ˣ��ƽ�
'------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim clsOneMouseUsage As clsMouseUsage
    Dim iDrawLabel As Integer
    Dim strField As Variant
    
    On Error GoTo errh
    
    For i = 1 To cMouseUsage.Count
        cMouseUsage.Remove 1
    Next
    
    If blLocalRun = True Then
        strSQL = "select ID,ֱ��,����,��Բ,��ͷ,�����,�����,�Ƕ�,����,����λ,����λ,����,����,�ü�_��ע����," & _
                 " ����Ӧ����,��ά���,����ע from Ӱ����갴ť���� "
        Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "select ��ԱID,ֱ��,����,��Բ,��ͷ,�����,�����,�Ƕ�,����,����λ,����λ,����,����,�ü�_��ע����," & _
                 " ����Ӧ����,��ά���,����ע from Ӱ����갴ť���� where ��Աid = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngUserID)
        If rsTmp.EOF = True Then
            strSQL = "select ��ԱID,ֱ��,����,��Բ,��ͷ,�����,�����,�Ƕ�,����,����λ,����λ,����,����,�ü�_��ע����," & _
                 " ����Ӧ����,��ά���,����ע from Ӱ����갴ť���� where ��Աid = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, CLng(0))
            
        End If
    End If
    
    iDrawLabel = -1
    If rsTmp.EOF = True Then Exit Sub
    For i = 1 To rsTmp.Fields.Count - 1
        '�����ݿ����ݱ��浽ϵͳ����������
        Set clsOneMouseUsage = New clsMouseUsage
        strField = Split(rsTmp(i).Value, ",")
        clsOneMouseUsage.lngFuncNo = strField(0)
        clsOneMouseUsage.lngMouseKey = strField(1)
        clsOneMouseUsage.lngShift = strField(2)
        clsOneMouseUsage.bSelected = strField(3)
        clsOneMouseUsage.strProgramName = strField(4)
        clsOneMouseUsage.ButtomID = strField(5)
        clsOneMouseUsage.strShowName = rsTmp(i).Name
        
        cMouseUsage.Add clsOneMouseUsage, CStr(clsOneMouseUsage.lngFuncNo)
        If clsOneMouseUsage.lngFuncNo = lngDrawLabelFuncNo Then
             iDrawLabel = cMouseUsage.Count
        End If
    Next
    
    '��д��굱ǰѡ��״̬��������ע��ť��
    If cMouseUsage(CStr(lngDrawLabelCurrent)).bSelected Then
        cMouseUsage(iDrawLabel).bSelected = True
    End If
    
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subGetInfoLabelToVar()
'------------------------------------------------
'���ܣ������ݿ��ȡ��Ϣ��עλ���������ݵ�ϵͳ��������ȡ��ͼ����Ϣ�������ݵ�ϵͳ����
'��������
'���أ���
'�ϼ���������̣�frmViewer.Form_Load
'�¼���������̣���
'���õ��ⲿ������aInfoLabelLocate
'�����ˣ��ƽ�
'------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    
    '������Ϣ��ע������
    If blLocalRun = True Then
        strSQL = "SELECT COUNT(id) as iCount FROM Ӱ��ͼ����Ϣ�� WHERE ����=-1"
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT COUNT(id) as iCount FROM Ӱ��ͼ����Ϣ�� WHERE ����=-1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    End If
    lngInfoLabelCount = rsTemp!iCount
    
    ''''''''��ʼ��[��Ϣ��ע����]��ϵͳ����
    ReDim aInfoLabelLocate(lngInfoLabelCount) As TInfoLabelLocate  ''������Ϣ��ע��λ��
    
    On Error GoTo errh
    '�����ݿ��ȡ��Ϣ��עλ��
    If blLocalRun = True Then
        strSQL = "SELECT id,��ʼ��ַ,������ַ,Ӣ�ļ��,���ļ��,��ѡ��,λ��,�������,�ɵ��� FROM Ӱ��ͼ����Ϣ�� WHERE ����=-1"
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT id,��ʼ��ַ,������ַ,Ӣ�ļ��,���ļ��,��ѡ��,λ��,�������,�ɵ��� FROM Ӱ��ͼ����Ϣ�� WHERE ����=-1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    End If
    i = 1
    With rsTemp
        .MoveFirst
        While Not .EOF
            aInfoLabelLocate(i).lngID = !Id
            aInfoLabelLocate(i).bUsed = IIf(IsNull(!��ѡ��), False, IIf(!��ѡ�� = -1, True, False))
            aInfoLabelLocate(i).strGroup = !��ʼ��ַ
            aInfoLabelLocate(i).strElement = !������ַ
            aInfoLabelLocate(i).lngLocation = IIf(IsNull(!λ��) = True, 0, !λ��)
            aInfoLabelLocate(i).lngOrder = IIf(IsNull(!�������), 0, !�������)
            aInfoLabelLocate(i).strCName = IIf(IsNull(!���ļ��), "", !���ļ��)
            aInfoLabelLocate(i).strEName = IIf(IsNull(!Ӣ�ļ��), "", !Ӣ�ļ��)
            aInfoLabelLocate(i).blnIsExport = IIf(IsNull(!�ɵ���), 0, !�ɵ���)
            i = i + 1
            .MoveNext
        Wend
    End With
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subGetDBDicomPrintToVar()
'------------------------------------------------
'���ܣ������ݿ��ȡ��ӡ���Ĳ�������д��ϵͳ�����������棬��ȡ��DICOM��ӡ�����á���
'��������
'���أ���
'�ϼ���������̣�frmViewer.Form_Load
'�¼���������̣���
'���õ��ⲿ������cDICOMPrinter
'�����ˣ��ƽ�
'------------------------------------------------
    '��ʼ������
    Dim i As Integer
    Dim clsOnePrinter As clsDicomPrint
    
    For i = 1 To cDICOMPrinter.Count
        cDICOMPrinter.Remove (1)
    Next
    '�����ݿ��ȡ��Ϣ
    Dim strSQL As String
    
    On Error GoTo errh
    
    cstrPrintAE = GetSetting("ZLSOFT", "����ģ��\zlPacsCore", "����AE", "ZLPACS")
    blnPrintOkEcho = GetSetting("ZLSOFT", "����ģ��\zlPacsCore", "��ӡ�ɹ�����ʾ", "False")
     
    If blLocalRun = True Then
        strSQL = "SELECT ID,��ӡ����,IP��ַ,�˿ں�,AE����,��ӡ��ʽ,���ȼ�,��ӡ����,����,����," & _
                 "��Ƭ���,ѡ��Ƭ��,�ֱ���,�Ŵ�ģʽ,ƽ��ģʽ,����,��С�ܶ�,����ܶ�,�հ��ܶ�," & _
                 "�߿��ܶ�,����,ͼ��λ��,�û�AE���� FROM Ӱ���ӡ������"
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT ID,��ӡ����,IP��ַ,�˿ں�,AE����,��ӡ��ʽ,���ȼ�,��ӡ����,����,����," & _
                 "��Ƭ���,ѡ��Ƭ��,�ֱ���,�Ŵ�ģʽ,ƽ��ģʽ,����,��С�ܶ�,����ܶ�,�հ��ܶ�," & _
                 "�߿��ܶ�,����,ͼ��λ��,�û�AE����,ͼ��߿���,ͼƬ�ֱ��� FROM Ӱ���ӡ������"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    End If
    '�����ݿ�������д��ϵͳ������������
    With rsTemp
        While Not .EOF
            Set clsOnePrinter = New clsDicomPrint
            clsOnePrinter.lngID = !Id
            clsOnePrinter.lngCopies = !��ӡ����
            clsOnePrinter.lngPort = !�˿ں�
            clsOnePrinter.strAETitle = !AE����
            clsOnePrinter.strBorderDensity = IIf(IsNull(!�߿��ܶ�), "", !�߿��ܶ�)
            clsOnePrinter.strEmptyDensity = IIf(IsNull(!�հ��ܶ�), "", !�հ��ܶ�)
            clsOnePrinter.strFilmBox = IIf(IsNull(!ѡ��Ƭ��), "", !ѡ��Ƭ��)
            clsOnePrinter.strFilmSize = IIf(IsNull(!��Ƭ���), "", !��Ƭ���)
            clsOnePrinter.strFormat = IIf(IsNull(!��ӡ��ʽ), "", !��ӡ��ʽ)
            clsOnePrinter.strIPAddress = IIf(IsNull(!IP��ַ), "", !IP��ַ)
            clsOnePrinter.strMagnification = IIf(IsNull(!�Ŵ�ģʽ), "", !�Ŵ�ģʽ)
            clsOnePrinter.strMaxDensity = IIf(IsNull(!����ܶ�), "", !����ܶ�)
            clsOnePrinter.strMedium = IIf(IsNull(!����), "", !����)
            clsOnePrinter.strMinDensity = IIf(IsNull(!��С�ܶ�), "", !��С�ܶ�)
            clsOnePrinter.strname = IIf(IsNull(!��ӡ����), "", !��ӡ����)
            clsOnePrinter.strOrientation = IIf(IsNull(!����), "", !����)
            clsOnePrinter.strPolarity = IIf(IsNull(!����), "", !����)
            clsOnePrinter.strPriority = IIf(IsNull(!���ȼ�), "", !���ȼ�)
            clsOnePrinter.strResolution = IIf(IsNull(!�ֱ���), "", !�ֱ���)
            clsOnePrinter.strSmooth = IIf(IsNull(!ƽ��ģʽ), "", !ƽ��ģʽ)
            clsOnePrinter.strTrim = IIf(IsNull(!����), "", !����)
            clsOnePrinter.lngBitDepth = IIf(IsNull(!ͼ��λ��), 8, !ͼ��λ��)
            clsOnePrinter.strSCUAETitle = IIf(IsNull(!�û�AE����), cstrPrintAE, !�û�AE����)
            clsOnePrinter.lngImageBorderWidth = Val(Nvl(!ͼ��߿���, 1))
            If clsOnePrinter.lngImageBorderWidth < 1 Or clsOnePrinter.lngImageBorderWidth > 99 Then
                clsOnePrinter.lngImageBorderWidth = 1
            End If
            clsOnePrinter.intImageResolution = Val(Nvl(!ͼƬ�ֱ���, 300))
            If clsOnePrinter.intImageResolution < 10 Or clsOnePrinter.intImageResolution > 999 Then
                clsOnePrinter.intImageResolution = 300
            End If
            cDICOMPrinter.Add clsOnePrinter, clsOnePrinter.strname
            .MoveNext
        Wend
    End With
    
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subGetInterfaceParaToVar(Optional blDefaultVal As Long)
'------------------------------------------------
'���ܣ������ݿ��ȡ��Ӱ���������������ݣ������䱣�浽ϵͳ�����С�
'������     blDefaultVal �Ƿ�ȡȱʡֵ =True ȡȱʡֵ =False ȡ����ֵ
'���أ���
'------------------------------------------------
    Dim strDefaultVal As String     ''��ȡ�ֶ���
    Dim StrTmp As String
    
    '��ʼ��ϵͳ����
    lngSelectedImageBorderColor = 1 ''ѡ��ͼ��߿���ɫ
    lngCurrentImageBorderColor = 1  ''ѡ��ͼ��߿���ɫ
    lngCurrentSeriesBorderColor = 1 ''ѡ�����б߿���ɫ
    lngSelectImageForeColour = 1    ''ѡ��ͼ���ʶ���ɫ
    lngPeriodColor = 1              ''ѡ������ɫ
    lngReferenceLineColor = 1       ''��λ����ɫ
    lngViewerBackColor = 1          ''Viewer������ɫ
    lngProgramBackColor = 1         ''���򱳾���ɫ

    lngSelectedImageBorderLineStyle = 0 ''ѡ��ͼ��߿�����
    lngSelectedImageBorderLineWidth = 0 ''ѡ��ͼ��߿��߿��
    lngCurrentImageBorderLineStyle = 0 ''��ǰͼ��߿�����
    lngCurrentImageBorderLineWidth = 0 ''��ǰͼ��߿��߿��
    lngImageIdentifierSize = 0         ''ͼ��ѡ���Ǵ�С
    intPeriodSize = 0                  ''ѡ������С
    lngReferenceLineStyle = 0          ''��λ������
    lngReferenceLineSpacing = 1        ''��λ�߼��

    intSpaceSize = 0                          ''����֮��ļ����ȡ��߶�
    intMaxAreaX = 0                           ''�������ɻ��ֵ�����
    intMaxAreaY = 0                           ''�������ɻ��ֵ�����
    lngCellSpacing = 0                        ''ͼ����
    blnDsipSpilthBorder = False               ''����߿��Ƿ���ʾ
    blnDockMiniImage = False                  ''����ͼͣ���ڲ˵���
    blnShowMiniImageInfo = True               ''����ͼ���Ƿ���ʾͼ����Ϣ
    blnShowMPRLine = True                     ''MPR��ʾ�����ߣ�Ĭ����True
    blnSquareFrame = True                     ''�����ο�ѡ
    blnShowPrintTag = False                   ''�Ƿ���ʾ��Ƭ��ӡ���
    blnPrintFilmBeep = False                  ''��Ƭ��ӡʱ�Ƿ���ʾ������������ӽ�Ƭ����ӡ
    
    '��������ɫ
    lngLabelColor = 1             ''��ע��ʾɫ����ɫ
    lngLabelSelectedColor = 1     ''��עѡ��ɫ����ɫ
    lngRulerLeftColor = 1         ''�����ɫ
    
    lngLabelLineStyleNorm = 0      ''����
    lngLabelLineWidthNorm = 1      ''�߿�
    lngLabelFontSize = 16           ''�����С
    '��ע����
    lngWinWidthLevelLocation = 1    '' ����λλ��
    '�������ֵ���ʾ����
    bROIArea = False      ''��ʾ���
    bROIMean = False      ''��ʾƽ��ֵ
    bROIStandardDeviation = False  ''��ʾ������
    bROILength = False    ''��ʾ�ܳ�
    bROIMax = False       ''��ʾ���ֵ
    bROIMin = False       ''��ʾ��Сֵ
    bROITextChinese = False                    ''����ʹ������
    intTextoOffX = 0                           ''��ע���ֵ�ƫ����
    intTextoOffY = 0                           ''��ע���ֵ�ƫ����
    blnLabelTextScaleFontSize = False          ''��ע���ִ�С�Ƿ�����ͼ��һ������
    '��λ�������
    blnAnatomicMarkersLeft = False     ''�Ƿ���ʾ�����λ���
    blnAnatomicMarkersRight = False      ''�Ƿ���ʾ�ұ���λ���
    blnAnatomicMarkersTop = False       ''�Ƿ���ʾ�ϱ���λ���
    blnAnatomicMarkersBottom = False     ''�Ƿ���ʾ�±���λ���
    blnChinaMark = False                   ''�Ƿ���ú�����ʾ��λ���
    '�������
    blnRulerDsipLeft = False             ''�Ƿ���ʾ��߱��
    blnRulerDsipRight = False           ''�Ƿ���ʾ�ұ߱��
    blnRulerDsipTop = False             ''�Ƿ���ʾ�ϱ߱��
    blnRulerDsipBottom = False          ''�Ƿ���ʾ�±߱��
    intRulerLeft = 0                     ''�����߾�
    intRulerTop = 0                       ''����ϱ߾�
    intRulerWidth = 0                   ''��߿��
    intRulerHeight = 0                   ''��߸߶�
    intRulerLineWidth = 0               ''����߿�
    '����������
    intToolBarIconSize = 32             ''������ͼ���С
    intToolBarPosition = 1              ''������λ��
    blToolBarHide = True                ''��������ʾ
    
    'Ѫ����խ����
    intStandardThreshold = 50           ''����Ѫ�ܲ�������ֵ
    intNarrowThreshold = 50             ''��խѪ�ܲ�������ֵ
    intVasEdgeWidth = 10                   ''Ѫ����խ��������ʾѪ�ܱڶ�ֱ�ߵĿ��
    
    '�������
    lngStackStep = 10                   ''��괩�󲽳�
    lngCruiseStep = 10                  ''������β���
    lngWidthLevelStep = 10              ''����������
    lngZoomStep = 10                    ''������Ų���
    intMouseWheelRoll = 0               ''������
    
    
    '��ȡ���ݿ��ֵ
    Dim strSQL As String
    If blLocalRun = True Then
        strSQL = "select * from Ӱ����������"
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "select * from Ӱ���������� where ��Աid = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, blDefaultVal)
        
        If rsTemp.EOF = True Then
            strSQL = "Select * from Ӱ���������� where ��Աid = 0 "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
        End If
    End If
    
    lngSelectedImageBorderColor = rsTemp("����ͼ��߿���ɫ")                                ''ѡ��ͼ��߿���ɫ
    lngCurrentImageBorderColor = rsTemp("ѡ��ͼ��߿���ɫ")                                 ''ѡ��ͼ��߿���ɫ
    lngCurrentSeriesBorderColor = rsTemp("ѡ�����б߿���ɫ")                                ''��ǰ��δѡ�У����б߿���ɫ
    lngSelectImageForeColour = rsTemp("ͼ������ɫ")                                       ''ѡ��ͼ���ʶ���ɫ
    lngPeriodColor = rsTemp("��עѡ������ɫ")                                             ''ѡ������ɫ
    lngReferenceLineColor = rsTemp("��λ����ɫ")                                            ''��λ����ɫ
    lngViewerBackColor = rsTemp("������ɫ")                                                 ''Viewer������ɫ
    lngProgramBackColor = rsTemp("���򱳾���ɫ")                                            ''���򱳾���ɫ
    lngSelectedImageBorderLineStyle = rsTemp("����ͼ��߿�����")                            ''ѡ��ͼ��߿�����
    lngSelectedImageBorderLineWidth = rsTemp("����ͼ��߿��߿�")                            ''ѡ��ͼ��߿��߿��
    lngCurrentImageBorderLineStyle = rsTemp("ѡ��ͼ��߿�����")                             ''��ǰͼ��߿�����
    lngCurrentImageBorderLineWidth = rsTemp("ѡ��ͼ��߿��߿�")                             ''��ǰͼ��߿��߿�
    lngImageIdentifierSize = rsTemp("ͼ���Ǵ�С")                                         ''ͼ��ѡ���Ǵ�С
    intPeriodSize = rsTemp("��עѡ������С")                                              ''ѡ������С
    lngReferenceLineStyle = rsTemp("��λ������")                                            ''��λ������
    lngReferenceLineSpacing = rsTemp("��λ�߼��")                                          ''��λ�߼��
    intSpaceSize = rsTemp("���м���")                                                     ''����֮��ļ����ȡ��߶�
    intMaxAreaX = rsTemp("�����������")                                                    ''�������ɻ��ֵ�����
    intMaxAreaY = rsTemp("�����������")                                                    ''�������ɻ��ֵ�����
    If intMaxAreaX < 1 Or intMaxAreaX > 8 Then intMaxAreaX = 8
    If intMaxAreaY < 1 Or intMaxAreaY > 8 Then intMaxAreaY = 8
    lngCellSpacing = rsTemp("ͼ����")                                                     ''ͼ����
    blnDsipSpilthBorder = IIf(rsTemp("��ʾ����߿�") = -1, True, False)                     ''����߿��Ƿ���ʾ
    bShowFilmConfig = IIf(rsTemp("ֱ������") = -1, True, False)                             ''�Ƿ�ֱ�����࣬����ʾ��Ƭ���ô���
    intStatusBarFontSize = rsTemp("״̬�������С")                                         ''״̬̬�����С
    blnShowPrintTag = IIf(rsTemp("��ʾ��ӡ���") = -1, True, False)                         ''�Ƿ���ʾ��Ƭ��ӡ���
    '��ȡ�����Ϣ
    lngStackStep = rsTemp("��괩�󲽳�")
    lngCruiseStep = rsTemp("������β���")
    lngWidthLevelStep = rsTemp("����������")
    lngZoomStep = rsTemp("������Ų���")
    intMouseWheelRoll = Nvl(rsTemp("�����ֲ���"), 0)
    '�����ݿ��ȡ������Ϣ��ע����ʾ����
    lngPatientInfoInvisibleSize = rsTemp("������Ϣ��ʾ��Сֵ")
    lngpatientInfoColor = rsTemp("������Ϣ��ɫ")
    blnpatientInfoScaleFontSize = IIf(rsTemp("������Ϣ��ͼ������") = -1, True, False)
    
    StrTmp = rsTemp("������Ϣ����")
    If UBound(Split(StrTmp, "|")) = 3 Then
        '���������Ϣ����������|�ֺ�|����|б�塱
        strPatientInfoFontName = Split(StrTmp, "|")(0)
        lngPatientInfoFontSize = Val(Split(StrTmp, "|")(1))
        blnPatientInfoFontBold = IIf(Split(StrTmp, "|")(2) = 1, True, False)
        blnPatientInfoFontItalic = IIf(Split(StrTmp, "|")(3) = 1, True, False)
    Else
        '��ǰ���ݣ�������Ϣ�����ֶ�ԭ��ֱ�ӱ�����������С
        lngPatientInfoFontSize = Val(StrTmp)
        blnPatientInfoFontBold = False
        blnPatientInfoFontItalic = False
        strPatientInfoFontName = "����"
    End If
    
    lngPatientInfoTitle = rsTemp("������Ϣ��ͷ")
    '��������ɫ
    lngLabelColor = rsTemp("��ע������ɫ")                                                  ''��ע��ʾɫ����ɫ
    lngLabelSelectedColor = rsTemp("��עѡ����ɫ")                                          ''��עѡ��ɫ����ɫ
    lngRulerLeftColor = rsTemp("�����ɫ")                                                  ''�����ɫ
    '��ע����
    lngWinWidthLevelLocation = rsTemp("����λλ��")
    lngLabelLineStyleNorm = rsTemp("��ע��������")
    lngLabelLineWidthNorm = rsTemp("��ע�����߿�")
    lngLabelFontSize = rsTemp("��ע���ִ�С")
    '�������ֵ���ʾ����
    bROIArea = IIf(rsTemp("������ʾ���") = -1, True, False)                                ''��ʾ���
    bROIMean = IIf(rsTemp("������ʾƽ��ֵ") = -1, True, False)                              ''��ʾƽ��ֵ
    bROIStandardDeviation = IIf(rsTemp("������ʾ������") = -1, True, False)                 ''��ʾ������
    bROILength = IIf(rsTemp("������ʾ�ܳ�") = -1, True, False)                              ''��ʾ�ܳ�
    bROIMax = IIf(rsTemp("������ʾ���ֵ") = -1, True, False)                               ''��ʾ���ֵ
    bROIMin = IIf(rsTemp("������ʾ��Сֵ") = -1, True, False)                               ''��ʾ��Сֵ
    bROITextChinese = IIf(rsTemp("������ʾ����") = -1, True, False)                         ''�����Ľ���Ƿ�ʹ������
    intTextoOffX = rsTemp("����X����ƫ��")                                                  ''��ע���ֵ�ƫ����
    intTextoOffY = rsTemp("����Y����ƫ��")                                                  ''��ע���ֵ�ƫ����
    blnLabelTextScaleFontSize = IIf(rsTemp("������ͼ������") = -1, True, False)             ''��ע���ִ�С�Ƿ�����
    '��λ�������
    blnAnatomicMarkersLeft = IIf(Mid(rsTemp("��ʾ��λ���"), 1, 1) = 1, True, False)        ''�Ƿ���ʾ�����λ���
    blnAnatomicMarkersRight = IIf(Mid(rsTemp("��ʾ��λ���"), 3, 1) = 1, True, False)       ''�Ƿ���ʾ�ұ���λ���
    blnAnatomicMarkersTop = IIf(Mid(rsTemp("��ʾ��λ���"), 2, 1) = 1, True, False)         ''�Ƿ���ʾ�ϱ���λ���
    blnAnatomicMarkersBottom = IIf(Mid(rsTemp("��ʾ��λ���"), 4, 1) = 1, True, False)      ''�Ƿ���ʾ�±���λ���
    ''�Ƿ���ú�����ʾ��λ���
    blnChinaMark = IIf(rsTemp("������λ���") = -1, True, False)
    '�������
    blnRulerDsipLeft = IIf(Mid(rsTemp("��ʾ���"), 1, 1) = 1, True, False)                  ''�Ƿ���ʾ��߱��
    blnRulerDsipRight = IIf(Mid(rsTemp("��ʾ���"), 3, 1) = 1, True, False)                 ''�Ƿ���ʾ�ұ߱��
    blnRulerDsipTop = IIf(Mid(rsTemp("��ʾ���"), 2, 1) = 1, True, False)                   ''�Ƿ���ʾ�ϱ߱��
    blnRulerDsipBottom = IIf(Mid(rsTemp("��ʾ���"), 4, 1) = 1, True, False)                ''�Ƿ���ʾ�±߱��
    intRulerLeft = rsTemp("������ұ߾�")                                                   ''�����߾�
    intRulerTop = rsTemp("������±߾�")                                                    ''����ϱ߾�
    intRulerWidth = rsTemp("��߿��")                                                      ''��߿��
    intRulerHeight = rsTemp("��߸߶�")                                                     ''��߸߶�
    intRulerLineWidth = rsTemp("����߿�")                                                  ''����߿�
    '����������
    intToolBarIconSize = rsTemp("������ͼ���С")
    intToolBarPosition = rsTemp("������λ��")
    blToolBarHide = IIf(rsTemp("��������ʾ") = -1, True, False)
    'Ѫ����խ����
    intStandardThreshold = rsTemp("����Ѫ����ֵ")
    intNarrowThreshold = rsTemp("��խѪ����ֵ")
    intVasEdgeWidth = rsTemp("Ѫ�ܱڿ��")
    
    '����ͼ�Ƿ�ͣ���ڲ˵�����
    blnDockMiniImage = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\" & App.EXEName & "\frmSysConfig", "����ͼͣ���ڲ˵���", False)
    blnShowMiniImageInfo = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\" & App.EXEName & "\frmSysConfig", "����ͼ����ʾͼ����Ϣ", True)
    blnSquareFrame = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\" & App.EXEName & "\frmSysConfig", "��ѡ����ͼ", True)
    blnShowMPRLine = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\" & App.EXEName & "\frmSysConfig", "MPR��ʾ������", True)
    blnPrintFilmBeep = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\" & App.EXEName & "\frmSysConfig", "��Ƭ��ӡ��ʾ����", False)
    
     '����FTP�ļ���С�Ա�
    gblnCompareSize = IIf(Val(GetSetting("ZLSOFT", "����ģ��\Ftp", "����FTP�ļ���С�Ա�", 1)) <> 0, True, False)
    Call SaveSetting("ZLSOFT", "����ģ��\Ftp", "����FTP�ļ���С�Ա�", IIf(gblnCompareSize, 1, 0))
End Sub


Public Sub LoadBarSetup(f As frmViewer)
'------------------------------------------------
'���ܣ���ȡ���ݿ��ϴα���Ĺ���������
'������f--���������
'���أ���
'�����ˣ�����
'------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    blfrmRefresh = False
       
    Select Case intToolBarIconSize
        Case 16
            BarterIco f.ImgList16
            CreateMenu f.ComToolBar, 16, 16
            f.ComToolBar.AddImageList f.ImgList16
        Case 24
            BarterIco f.ImgList24
            CreateMenu f.ComToolBar, 24, 24
            f.ComToolBar.AddImageList f.ImgList24
        Case 32
            BarterIco f.ImgList32
            CreateMenu f.ComToolBar, 32, 32
            f.ComToolBar.AddImageList f.ImgList32
    End Select
    
    f.ComToolBar.Item(ToolBar_Main).Position = intToolBarPosition
    
    ArrayToolBar f.ComToolBar, f.top, f.left, f.height, f.width

    For i = 2 To 8
        f.ComToolBar.Item(i).Visible = blToolBarHide
    Next
    f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_ToolBar_Hide, , True).Checked = Not blToolBarHide
    blfrmRefresh = True
End Sub

Public Function CreateUserWWWL(lngUserID As Long) As Boolean
'�Ƿ���Ҫ�����û��Ĵ���λ����
'������ lngUserID --- �û�ID
'����ֵ��True ---�����ɹ���False --- ����ʧ�ܣ�����Ҫ����

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    strSQL = "select count(*) as Count from Ӱ��Ԥ�贰��λ where ��Աid =[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж��Ƿ񴴽��û���������", lngUserID)
    If rsTemp!Count = 0 Then
        strSQL = "Zl_Ӱ��Ԥ�贰��λ_Create(" & lngUserID & ")"
        zlDatabase.ExecuteProcedure strSQL, "�����û���������"
        CreateUserWWWL = True
    End If
    Exit Function
err:
    If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
End Function

Public Sub subSaveInterfaceParaIntoDB()
'------------------------------------------------
'���ܣ�����ǰ��ϵͳ����ֵ�����浽��Ӱ������������
'��������
'���أ���
'------------------------------------------------
    Dim strAnatomicMarkers As String        '������ʱ����λ��ע
    Dim strRulerDsip As String              '������ʱ�ı�߱�ע
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo err
    
    ''������������ʱ�򣬸��ݵ�¼״̬ blLocalRun ���жϣ��Ǳ��浽������MDB���ݿ⣬���Ǳ��浽������ORACLE���ݿ⡣
    
    '�����жϵ�ǰ�û��Ƿ�����������������������ǵ�һ�α����������������Ȳ���һ���û��Լ��Ľ��������¼
    If blLocalRun = True Then
        strSQL = "select * from Ӱ���������� "
        Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
        If rsTmp.EOF = True Then
            '--Ӱ����������
            strSQL = "insert into Ӱ���������� (ID,����ͼ��߿���ɫ,����ͼ��߿�����,����ͼ��߿��߿�,ѡ��ͼ��߿���ɫ,ѡ�����б߿���ɫ,ѡ��ͼ��߿�����,ѡ��ͼ��߿��߿�,ͼ������ɫ,ͼ���Ǵ�С,��עѡ������ɫ,��עѡ������С,��λ����ɫ,��λ������,��λ�߼��,���м���,�����������,�����������,ͼ����,��ʾ����߿�,������ɫ,���򱳾���ɫ,��ע������ɫ,��ע��������,��ע�����߿�,��עѡ����ɫ,��עѡ������,��עѡ���߿�,��ע���ִ�С,������ʾ���,������ʾƽ��ֵ,������ʾ������,������ʾ����,����X����ƫ��,����Y����ƫ��,������ͼ������,��ʾ��λ���,������λ���,��ʾ���,������ұ߾�,������±߾�,��߿��,��߸߶�,����߿�,�����ɫ,����λλ��,��괩�󲽳�,������β���,����������,������Ų���,������Ϣ���±߾�,������Ϣ���ұ߾�,������Ϣ��ɫ,������Ϣ��ʾ��Сֵ,������Ϣ��ͼ������,������Ϣ����,������Ϣ��ͷ,ֱ������,������ͼ���С,������λ��,��������ʾ,״̬�������С,����Ѫ����ֵ,��խѪ����ֵ,Ѫ�ܱڿ��,������ʾ�ܳ�)" & _
                     "VALUES (0,16777215,0,1,16777215,16777088,0,1,16777215,10,16777215,8,16777215,3,7,50,8,8,4,0,986895,131586,16777215,0,4,16777215,0,0,12,1,1,-1,-1,10,8,0,1010,1,1000,36,210,30,600,3,16777215,2,8,10,10,5,4,1,16777215,200,0,10,1,1,24,1,-1,9,51,50,10,-1);"
            cnAccess.Execute strSQL
            strSQL = "select * from Ӱ���������� "
            Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
        End If
    Else
        strSQL = "select * from Ӱ���������� where ��Աid = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡӰ��������", glngUserID)
        If rsTmp.EOF = True Then
            strSQL = "select * from Ӱ���������� where ��Աid = 0 "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡӰ��������")
            If rsTmp.EOF <> True Then
                strSQL = "ZL_Ӱ����������_INSERT(" & glngUserID
                For i = 1 To rsTmp.Fields.Count - 1
                    strSQL = strSQL & "," & rsTmp(i).Value
                Next
                strSQL = strSQL & ")"
                zlDatabase.ExecuteProcedure strSQL, "��ȡӰ��������"
            End If
        End If
    End If
    
    '����������
    If blLocalRun = True Then
        strSQL = "update Ӱ���������� set "
        ''ѡ��ͼ��߿���ɫ
        strSQL = strSQL & "����ͼ��߿���ɫ = '" & lngSelectedImageBorderColor & "',"
        ''ѡ��ͼ��߿�����
        strSQL = strSQL & "����ͼ��߿����� = '" & lngSelectedImageBorderLineStyle & "',"
        ''ѡ��ͼ��߿��߿��
        strSQL = strSQL & "����ͼ��߿��߿� = '" & lngSelectedImageBorderLineWidth & "',"
        ''ѡ��ͼ��߿���ɫ�����ǵ�ǰ��ɫ
        strSQL = strSQL & "ѡ��ͼ��߿���ɫ = '" & lngCurrentImageBorderColor & "',"
        ''��ǰ��δѡ�У����б߿���ɫ
        strSQL = strSQL & "ѡ�����б߿���ɫ = '" & lngCurrentSeriesBorderColor & "',"
        ''��ǰͼ��߿�����
        strSQL = strSQL & "ѡ��ͼ��߿����� = '" & lngCurrentImageBorderLineStyle & "',"
        ''��ǰͼ��߿��߿��
        strSQL = strSQL & "ѡ��ͼ��߿��߿� = '" & lngCurrentImageBorderLineWidth & "',"
        ''ѡ��ͼ���ʶ���ɫ
        strSQL = strSQL & "ͼ������ɫ = '" & lngSelectImageForeColour & "',"
        ''ͼ��ѡ���Ǵ�С
        strSQL = strSQL & "ͼ���Ǵ�С = '" & lngImageIdentifierSize & "',"
        ''ѡ������ɫ
        strSQL = strSQL & "��עѡ������ɫ = '" & lngPeriodColor & "',"
        ''ѡ������С
        strSQL = strSQL & "��עѡ������С = '" & intPeriodSize & "',"
        ''��λ����ɫ
        strSQL = strSQL & "��λ����ɫ = '" & lngReferenceLineColor & "',"
        ''��λ������
        strSQL = strSQL & "��λ������ = '" & lngReferenceLineStyle & "',"
        ''��λ�߼��
        strSQL = strSQL & "��λ�߼�� = '" & lngReferenceLineSpacing & "',"
        ''����֮��ļ����ȡ��߶�
        strSQL = strSQL & "���м��� = '" & intSpaceSize & "',"
        ''�������ɻ��ֵ�����
        strSQL = strSQL & "����������� = '" & intMaxAreaX & "',"
        ''�������ɻ��ֵ�����
        strSQL = strSQL & "����������� = '" & intMaxAreaY & "',"
        ''ͼ����
        strSQL = strSQL & "ͼ���� = '" & lngCellSpacing & "',"
        ''����߿��Ƿ���ʾ
        strSQL = strSQL & "��ʾ����߿� = '" & CInt(blnDsipSpilthBorder) & "',"
        ''Viewer������ɫ
        strSQL = strSQL & "������ɫ = '" & lngViewerBackColor & "',"
        ''���򱳾���ɫ
        strSQL = strSQL & "���򱳾���ɫ = '" & lngProgramBackColor & "',"
        ''��ע��ʾɫ����ɫ
        strSQL = strSQL & "��ע������ɫ = '" & lngLabelColor & "',"
        ''��ע��������
        strSQL = strSQL & "��ע�������� = '" & lngLabelLineStyleNorm & "',"
        ''��ע�����߿�
        strSQL = strSQL & "��ע�����߿� = '" & lngLabelLineWidthNorm & "',"
        ''��עѡ��ɫ����ɫ
        strSQL = strSQL & "��עѡ����ɫ = '" & lngLabelSelectedColor & "',"
        ''��ע���ִ�С
        strSQL = strSQL & "��ע���ִ�С = '" & lngLabelFontSize & "',"
        ''��ʾ���
        strSQL = strSQL & "������ʾ��� = '" & CInt(bROIArea) & "',"
        ''��ʾƽ��ֵ
        strSQL = strSQL & "������ʾƽ��ֵ = '" & CInt(bROIMean) & "',"
        ''��ʾ������
        strSQL = strSQL & "������ʾ������ = '" & CInt(bROIStandardDeviation) & "',"
        ''���������Ϣ�Ƿ�ʹ������
        strSQL = strSQL & "������ʾ���� = '" & CInt(bROITextChinese) & "',"
        ''��ע���ֵ�ƫ����
        strSQL = strSQL & "����X����ƫ�� = '" & intTextoOffX & "',"
        ''��ע���ֵ�ƫ����
        strSQL = strSQL & "����Y����ƫ�� = '" & intTextoOffY & "',"
        ''��ע���ִ�С�Ƿ�����ͼ��һ������
        strSQL = strSQL & "������ͼ������ = '" & CInt(blnLabelTextScaleFontSize) & "',"
        ''��λ��ע
        strAnatomicMarkers = IIf(blnAnatomicMarkersLeft, 1, 0) & IIf(blnAnatomicMarkersTop, 1, 0) _
                             & IIf(blnAnatomicMarkersRight, 1, 0) & IIf(blnAnatomicMarkersBottom, 1, 0)
        strSQL = strSQL & "��ʾ��λ��� = '" & strAnatomicMarkers & "',"
        ''�Ƿ���ú�����ʾ��λ���
        strSQL = strSQL & "������λ��� = '" & CInt(blnChinaMark) & "',"
        ''��ʾ���
        strRulerDsip = IIf(blnRulerDsipLeft, 1, 0) & IIf(blnRulerDsipTop, 1, 0) _
                       & IIf(blnRulerDsipRight, 1, 0) & IIf(blnRulerDsipBottom, 1, 0)
        strSQL = strSQL & "��ʾ��� = '" & strRulerDsip & "',"
        ''�����߾�
        strSQL = strSQL & "������ұ߾� = '" & intRulerLeft & "',"
        ''����ϱ߾�
        strSQL = strSQL & "������±߾� = '" & intRulerTop & "',"
        ''��߿��
        strSQL = strSQL & "��߿�� = '" & intRulerWidth & "',"
        ''��߸߶�
        strSQL = strSQL & "��߸߶� = '" & intRulerHeight & "',"
        ''����߿�
        strSQL = strSQL & "����߿� = '" & intRulerLineWidth & "',"
        ''�����ɫ
        strSQL = strSQL & "�����ɫ = '" & lngRulerLeftColor & "',"
        ''����λλ��
        strSQL = strSQL & "����λλ�� = '" & lngWinWidthLevelLocation & "',"
        ''��괩�󲽳�
        strSQL = strSQL & "��괩�󲽳� = '" & lngStackStep & "',"
        ''������β���
        strSQL = strSQL & "������β��� = '" & lngCruiseStep & "',"
        ''����������
        strSQL = strSQL & "���������� = '" & lngWidthLevelStep & "',"
        ''������Ų���
        strSQL = strSQL & "������Ų��� = '" & lngZoomStep & "',"
        ''�����ֲ���
        strSQL = strSQL & "�����ֲ��� = '" & intMouseWheelRoll & "',"
        ''������Ϣ���±߾�
        strSQL = strSQL & "������Ϣ���±߾� = '0',"
        ''������Ϣ���ұ߾�
        strSQL = strSQL & "������Ϣ���ұ߾� = '0',"
        ''������Ϣ��ɫ
        strSQL = strSQL & "������Ϣ��ɫ = '" & lngpatientInfoColor & "',"
        ''������Ϣ��ʾ��Сֵ
        strSQL = strSQL & "������Ϣ��ʾ��Сֵ = '" & lngPatientInfoInvisibleSize & "',"
        ''������Ϣ��ͼ������
        strSQL = strSQL & "������Ϣ��ͼ������ = '" & CInt(blnpatientInfoScaleFontSize) & "',"
        ''������Ϣ����
        strSQL = strSQL & "������Ϣ���� = '" & lngPatientInfoFontSize & "',"
        ''������Ϣ��ͷ
        strSQL = strSQL & "������Ϣ��ͷ = '" & lngPatientInfoTitle & "',"
        ''�Ƿ�ֱ�����࣬����ʾ��Ƭ���ô���
        strSQL = strSQL & "ֱ������ = '" & CInt(bShowFilmConfig) & "',"
        strSQL = strSQL & "������ͼ���С = '" & intToolBarIconSize & "',"
        strSQL = strSQL & "������λ�� = '" & intToolBarPosition & "',"
        strSQL = strSQL & "��������ʾ = '" & CInt(blToolBarHide) & "',"
        ''״̬�������С
        strSQL = strSQL & "״̬�������С = '" & intStatusBarFontSize & "',"
        strSQL = strSQL & "����Ѫ����ֵ = '" & intStandardThreshold & "',"
        strSQL = strSQL & "��խѪ����ֵ = '" & intNarrowThreshold & "',"
        strSQL = strSQL & "Ѫ�ܱڿ�� = '" & intVasEdgeWidth & "',"
        ''��ʾ�ܳ�
        strSQL = strSQL & "������ʾ�ܳ� = '" & CInt(bROILength) & "' where id = 0"
        cnAccess.Execute strSQL, adCmdText
    Else
        strSQL = "ZL_Ӱ����������_UPDATE('" & glngUserID & "','"
        ''ѡ��ͼ��߿���ɫ
        strSQL = strSQL & lngSelectedImageBorderColor & "','"
        ''ѡ��ͼ��߿�����
        strSQL = strSQL & lngSelectedImageBorderLineStyle & "','"
        ''ѡ��ͼ��߿��߿��
        strSQL = strSQL & lngSelectedImageBorderLineWidth & "','"
        ''ѡ��ͼ��߿���ɫ�����ǵ�ǰ��ɫ
        strSQL = strSQL & lngCurrentImageBorderColor & "','"
        ''��ǰ��δѡ�У����б߿���ɫ
        strSQL = strSQL & lngCurrentSeriesBorderColor & "','"
        ''��ǰͼ��߿�����
        strSQL = strSQL & lngCurrentImageBorderLineStyle & "','"
        ''��ǰͼ��߿��߿��
        strSQL = strSQL & lngCurrentImageBorderLineWidth & "','"
        ''ѡ��ͼ���ʶ���ɫ
        strSQL = strSQL & lngSelectImageForeColour & "','"
        ''ͼ��ѡ���Ǵ�С
        strSQL = strSQL & lngImageIdentifierSize & "','"
        ''ѡ������ɫ
        strSQL = strSQL & lngPeriodColor & "','"
        ''ѡ������С
        strSQL = strSQL & intPeriodSize & "','"
        ''��λ����ɫ
        strSQL = strSQL & lngReferenceLineColor & "','"
        ''��λ������
        strSQL = strSQL & lngReferenceLineStyle & "','"
        ''��λ�߼��
        strSQL = strSQL & lngReferenceLineSpacing & "','"
        ''����֮��ļ����ȡ��߶�
        strSQL = strSQL & intSpaceSize & "','"
        ''�������ɻ��ֵ�����
        strSQL = strSQL & intMaxAreaX & "','"
        ''�������ɻ��ֵ�����
        strSQL = strSQL & intMaxAreaY & "','"
        ''ͼ����
        strSQL = strSQL & lngCellSpacing & "','"
        ''����߿��Ƿ���ʾ
        strSQL = strSQL & CInt(blnDsipSpilthBorder) & "','"
        ''Viewer������ɫ
        strSQL = strSQL & lngViewerBackColor & "','"
        ''���򱳾���ɫ
        strSQL = strSQL & lngProgramBackColor & "','"
        ''��ע��ʾɫ����ɫ
        strSQL = strSQL & lngLabelColor & "','"
        ''��ע��������
        strSQL = strSQL & lngLabelLineStyleNorm & "','"
        ''��ע�����߿�
        strSQL = strSQL & lngLabelLineWidthNorm & "','"
        ''��עѡ��ɫ����ɫ
        strSQL = strSQL & lngLabelSelectedColor & "','"
        ''��ע���ִ�С
        strSQL = strSQL & lngLabelFontSize & "','"
        ''��ʾ���
        strSQL = strSQL & CInt(bROIArea) & "','"
        ''��ʾƽ��ֵ
        strSQL = strSQL & CInt(bROIMean) & "','"
        ''��ʾ������
        strSQL = strSQL & CInt(bROIStandardDeviation) & "','"
        ''���������Ϣ�Ƿ�ʹ������
        strSQL = strSQL & CInt(bROITextChinese) & "','"
        ''��ע���ֵ�ƫ����X
        strSQL = strSQL & intTextoOffX & "','"
        ''��ע���ֵ�ƫ����Y
        strSQL = strSQL & intTextoOffY & "','"
        ''��ע���ִ�С�Ƿ�����ͼ��һ������
        strSQL = strSQL & CInt(blnLabelTextScaleFontSize) & "','"
        ''��λ��ע
        strAnatomicMarkers = IIf(blnAnatomicMarkersLeft, 1, 0) & IIf(blnAnatomicMarkersTop, 1, 0) _
                             & IIf(blnAnatomicMarkersRight, 1, 0) & IIf(blnAnatomicMarkersBottom, 1, 0)
        strSQL = strSQL & strAnatomicMarkers & "','"
        ''�Ƿ���ú�����ʾ��λ���
        strSQL = strSQL & CInt(blnChinaMark) & "','"
        ''��ʾ���
        strRulerDsip = IIf(blnRulerDsipLeft, 1, 0) & IIf(blnRulerDsipTop, 1, 0) _
                       & IIf(blnRulerDsipRight, 1, 0) & IIf(blnRulerDsipBottom, 1, 0)
        strSQL = strSQL & strRulerDsip & "','"
        ''�����߾�
        strSQL = strSQL & intRulerLeft & "','"
        ''����ϱ߾�
        strSQL = strSQL & intRulerTop & "','"
        ''��߿��
        strSQL = strSQL & intRulerWidth & "','"
        ''��߸߶�
        strSQL = strSQL & intRulerHeight & "','"
        ''����߿�
        strSQL = strSQL & intRulerLineWidth & "','"
        ''�����ɫ
        strSQL = strSQL & lngRulerLeftColor & "','"
        ''����λλ��
        strSQL = strSQL & lngWinWidthLevelLocation & "','"
        ''��괩�󲽳�
        strSQL = strSQL & lngStackStep & "','"
        ''������β���
        strSQL = strSQL & lngCruiseStep & "','"
        ''����������
        strSQL = strSQL & lngWidthLevelStep & "','"
        ''������Ų���
        strSQL = strSQL & lngZoomStep & "','"
        '������Ϣ���±߾�
        strSQL = strSQL & "0','"
        '������Ϣ���ұ߾�
        strSQL = strSQL & "0','"
        '������Ϣ��ɫ
        strSQL = strSQL & lngpatientInfoColor & "','"
        '������Ϣ��ʾ��Сֵ
        strSQL = strSQL & lngPatientInfoInvisibleSize & "','"
        '������Ϣ��ͼ������
        strSQL = strSQL & CInt(blnpatientInfoScaleFontSize) & "','"
        '������Ϣ����,������Ϣ��֯��������������|�ֺ�|����|б�塱
        strSQL = strSQL & strPatientInfoFontName & "|" & lngPatientInfoFontSize & "|" & IIf(blnPatientInfoFontBold, 1, 0) & "|" & IIf(blnPatientInfoFontItalic, 1, 0) & "','"
        ''������Ϣ��ͷ
        strSQL = strSQL & lngPatientInfoTitle & "','"
        ''�Ƿ�ֱ�����࣬����ʾ��Ƭ���ô���
        strSQL = strSQL & CInt(bShowFilmConfig) & "','"
        ''������ͼ���С
        strSQL = strSQL & intToolBarIconSize & "','"
        ''������λ��
        strSQL = strSQL & intToolBarPosition & "','"
        ''����������
        strSQL = strSQL & CInt(blToolBarHide) & "','"
        ''״̬�������С
        strSQL = strSQL & intStatusBarFontSize & "','"
        ''Ѫ����խ����������Ѫ����ֵ
        strSQL = strSQL & intStandardThreshold & "','"
        ''Ѫ����խ��������խѪ����ֵ
        strSQL = strSQL & intNarrowThreshold & "','"
        ''Ѫ����խ������Ѫ�ܱڿ��
        strSQL = strSQL & intVasEdgeWidth & "','"
        ''��ʾ�ܳ�
        strSQL = strSQL & CInt(bROILength) & "','"
        ''��ʾ���ֵ
        strSQL = strSQL & CInt(bROIMax) & "','"
        ''��ʾ��Сֵ
        strSQL = strSQL & CInt(bROIMin) & "','"
        ''�����ֲ���
        strSQL = strSQL & CInt(intMouseWheelRoll) & "','"
        ''�Ƿ���ʾ��Ƭ��ӡ���
        strSQL = strSQL & CInt(blnShowPrintTag) & "')"
        
        zlDatabase.ExecuteProcedure strSQL, "����Ӱ��������"
    End If
    Exit Sub
err:
    If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subSaveParameters()
'------------------------------------------------
'���ܣ����������Ĳ���
'��������
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    Call zlDatabase.SetPara("�������϶�����", intMouseWheelDrag, glngSys, 1289)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub subGetParameters()
'------------------------------------------------
'���ܣ���ȡ������Ĳ���
'��������
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    intMouseWheelDrag = Val(zlDatabase.GetPara("�������϶�����", glngSys, 1289, 0))
    If intMouseWheelDrag < 0 Or intMouseWheelDrag > 2 Then intMouseWheelDrag = 0
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub
