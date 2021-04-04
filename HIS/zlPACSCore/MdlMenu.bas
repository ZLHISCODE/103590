Attribute VB_Name = "MdlMenu"
Option Explicit
'--------------------------------------------------------
'��  �ܣ���ģ��Ϊ�˵���ť���õĴ�����
'�����ˣ�����
'�������ڣ�2004.6.12
'���̺����嵥��
'    subFakeColor():            �Դ�����ѡ�е�ͼ����α�ʴ�����ʾһ�����壬���û�ѡ����Ҫ��α�ʷ�����
'    subFunctionWL():           ���ܼ����ô���λ����
'    subFilm():                 ��Ƭ��ӡ
'    subcalibrate():            У׼
'    SubImageUnsharp():         ͼ����ǿ
'    subMnuImageSort():         ����ʽ��������lngToolID����������ͬʱ����˵���ѡ��״̬
'    subMouseRLset():           ����������Ҽ���check״̬��
'    subCurrentCheck():         ��λ�ߴ�������λ�߲˵��ĵ����¼������ƶ�λ��������ذ�ťֻ�ܱ�ѡ��һ��
'    subOutputToPowerPoint():   �����POWERPOINT
'    subDSA():                  DSA���ּ�Ӱ
'    subCutOut():               ������˳��ü�״̬�����ػ���ʾ�ü���ע
'    subDispLabelInfo():        ��ʾ������ͼ����û���ע��Ϣ
'    subManipulation():         ��ƽ�洦��ͼ����ת�����׵ȣ�
'    subSelectAllSerial():      ѡ����������
'    subSelectAllIMage():       ѡ������ͼ��
'    subFullScreen():           �л���Ļ��ȫ��״̬
'�޸ļ�¼��
'    2005.07.08    �ƽ�
'-------------------------------------------------------


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''
''''
''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub subFakeColor(f As frmViewer)
'------------------------------------------------
'���ܣ��Դ�����ѡ�е�ͼ����α�ʴ�����ʾһ�����壬���û�ѡ����Ҫ��α�ʷ�����
'������f--����α�ʴ���Ĵ��塣
'���أ���
'2009��
'------------------------------------------------
    If f.intSelectedSerial < 1 Then Exit Sub
    Dim strSQL As String, rsTemp As Recordset
    Set FrmFakeColor.f = f
    strSQL = "SELECT ��ɫ,���,ϵͳ���� FROM  Ӱ����ɫ�嵥"
    If blLocalRun = True Then
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯα����ɫ")
    End If
    Do While Not rsTemp.EOF
        FrmFakeColor.cobColor.AddItem IIf(rsTemp!ϵͳ���� = 1, "ϵͳ������", "�û�������") & rsTemp!��ɫ
        FrmFakeColor.cobColor.ItemData(FrmFakeColor.cobColor.NewIndex) = rsTemp!���
        rsTemp.MoveNext
    Loop
    rsTemp.MoveFirst
    FrmFakeColor.cobColor.ListIndex = 0
    FrmFakeColor.Show 1, f
End Sub

Public Sub subFunctionFilter(ByVal control As CommandBarControl, f As frmViewer)
'------------------------------------------------
'���ܣ����ܼ������˾�ģ�崦��
'������ Control--�˵��ؼ���
'       f--���塣
'���أ���
'------------------------------------------------
    Dim iRow As Integer
    Dim dblTemp As Double
    
    On Error GoTo err
    If f.SelectedImage Is Nothing Then Exit Sub
    If control Is Nothing Then Exit Sub
    If control.Id >= ID_Active_SieveLens_Model + 1 And control.Id < ID_Active_SieveLens_Model + 40 Then
        iRow = Val(control.Category)
        If iRow >= 0 And iRow < UBound(aPresetFilter) Then
            f.SelectedImage.UnsharpEnhancement = 0
            f.SelectedImage.UnsharpLength = 0
            f.SelectedImage.FilterLength = 0
            
            'ͼ����������ǿֵ
            'ͼ����ǿǿ������
            If aPresetFilter(iRow).intUnSharpEnhancementUp > 0 Then
                Call SubImageFiltering("miUnSharpEnhancementUp", f.SelectedImage, aPresetFilter(iRow).intUnSharpEnhancementUp)
            End If
            'ͼ����ǿǿ�ȼ���
            If aPresetFilter(iRow).intUnSharpEnhancementDown > 0 Then
                Call SubImageFiltering("miUnSharpEnhancementDown", f.SelectedImage, aPresetFilter(iRow).intUnSharpEnhancementDown)
            End If
            
            'ͼ����ǿ��������
            If aPresetFilter(iRow).intUnSharpLengthUp > 0 Then
                Call SubImageFiltering("miUnSharpLengthUp", f.SelectedImage, aPresetFilter(iRow).intUnSharpLengthUp)
            End If
            
            'ͼ����ǿ���ȼ���
            If aPresetFilter(iRow).intUnSharpLengthDown > 0 Then
                Call SubImageFiltering("miUnSharpLengthDown", f.SelectedImage, aPresetFilter(iRow).intUnSharpLengthDown)
            End If
            
            'ƽ������
            If aPresetFilter(iRow).intFilterLengthUp > 0 Then
                Call SubImageFiltering("miFilterLengthUp", f.SelectedImage, aPresetFilter(iRow).intFilterLengthUp)
            End If
            
            'ƽ������
            If aPresetFilter(iRow).intFilterLengthDown > 0 Then
                Call SubImageFiltering("miFilterLengthDown", f.SelectedImage, aPresetFilter(iRow).intFilterLengthDown)
            End If
        End If
    End If
    
    '����������ͼ��ͬ��
    Call subSeriesInPhase(f.intSelectedSerial, f, f.SelectedImage, IMG_SYN_FILTER)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subFunctionWL(ByVal control As CommandBarControl, f As Form)
'------------------------------------------------
'���ܣ����ܼ����ô���λ����
'������ Control--�˵��ؼ���
'       f--���塣
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim iWWidth As Integer
    Dim iWLevel As Integer
    Dim intFormType As Integer  '1������Ƭ���壬2�ǽ�Ƭ��ӡ����,3�ǽ�ƬԤ���󴰿�
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If f.SelectedImage Is Nothing Then Exit Sub
    If control Is Nothing Then Exit Sub
    
    
    If f.Name = "frmFilmView" Then
        intFormType = 3
    ElseIf f.Name = "frmFilm" Then
        intFormType = 2
    Else
        intFormType = 1
    End If
    
    If control.Id = ID_Active_AdjustWindow_HandAdjustWindow_Custom Then
        frmWindowCustom.lngWindow = f.SelectedImage.width
        frmWindowCustom.lngLevel = f.SelectedImage.Level
        frmWindowCustom.Show 1, f
        If frmWindowCustom.bApply Then
            control.Category = frmWindowCustom.lngWindow & "-" & frmWindowCustom.lngLevel
        End If
    End If
    i = InStr(control.Category, "-")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If i <> 0 Then
        f.SelectedImage.width = Val(Mid(control.Category, 1, i - 1))
        f.SelectedImage.Level = Val(Mid(control.Category, i + 1))
        f.SelectedImage.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & f.SelectedImage.width & "-L:" & f.SelectedImage.Level
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If control.Id = ID_Active_AdjustWindow_HandAdjustWindow_ReSet Then
        f.SelectedImage.VOILUT = 1
        '�ж��Ƿ�������Ĭ�ϴ���
        If f.blnDefaultWW2 = False Then
            '��ʾĬ�ϵڶ�������
            If f.SelectedImage.Attributes(&H28, &H1050).VM = 2 And f.SelectedImage.Attributes(&H28, &H1051).VM = 2 Then
                iWWidth = f.SelectedImage.Attributes(&H28, &H1051).ValueByIndex(2)
                iWLevel = f.SelectedImage.Attributes(&H28, &H1050).ValueByIndex(2)
                f.SelectedImage.width = iWWidth
                f.SelectedImage.Level = iWLevel
                f.blnDefaultWW2 = True
            Else
                f.SelectedImage.SetDefaultWindows
            End If
        Else
            f.SelectedImage.SetDefaultWindows
            f.blnDefaultWW2 = False
        End If
        
        
        If f.SelectedImage.Attributes(&H6000, &H15).Value = 1 Then
            If f.SelectedImage.Level = 0 Then
                f.SelectedImage.Level = 1
            End If
        End If
    End If
    '����������ͼ��ͬ��
    If intFormType = 1 Then '����Ƭ����
        Call subSeriesInPhase(f.intSelectedSerial, f, f.SelectedImage, IMG_SYN_WINDOW)
    ElseIf intFormType = 2 Then     '��ƬԤ������
        Call f.subSynchronalImg(False, IMG_SYN_WINDOW)
    End If
End Sub

Public Sub subcalibrate(f As frmViewer)
'------------------------------------------------
'���ܣ�У׼
'������f��������
'���أ���
'2009��
'------------------------------------------------
    Dim va As Variant, l As DicomLabel
    If f.SelectedImage Is Nothing Or f.SelectedLabel Is Nothing Then
        MsgBox "У׼����ѡ��һ��ֱ�߱�ע", vbInformation, gstrSysName
        Exit Sub
    End If
    If f.SelectedLabel.LabelType <> doLabelLine Then
        MsgBox "У׼����ѡ��һ��ֱ�߱�ע", vbInformation, gstrSysName
        Exit Sub
    End If
    va = f.SelectedImage.Attributes(&H28, &H30).Value
    If Not IsNull(va) Then
        If MsgBox("��ͼ���Ѿ���У׼��Ϣ,�Ƿ�����У׼?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    End If
    
    Dim strResult As String
    strResult = 0
    strResult = InputBox("��ֱ��ԭ������Ϊ��" & f.SelectedLabel.ROILength, "ֱ��У׼", f.SelectedLabel.ROILength, _
                    f.left + f.width / 4, f.top + f.height / 4)
    If strResult = "" Then Exit Sub
    If strResult < 0 Then
        MsgBox "У׼����Ӧ�ô���0��������У׼��", vbInformation, gstrSysName
        Exit Sub
    Else
        f.dubCalibrateLength = Val(strResult)
    End If
    If IsNull(va) Then
        ReDim va(1 To 2)
        va(1) = f.dubCalibrateLength / f.SelectedLabel.ROILength
        va(2) = f.dubCalibrateLength / f.SelectedLabel.ROILength
    Else
        va(1) = va(1) * f.dubCalibrateLength / f.SelectedLabel.ROILength
        va(2) = va(2) * f.dubCalibrateLength / f.SelectedLabel.ROILength
    End If
    f.SelectedImage.Attributes.Add &H28, &H30, va
    For Each l In f.SelectedImage.Labels
        If f.SelectedImage.Labels.IndexOf(l) > G_INT_SYS_LABEL_COUNT And l.LabelType = doLabelText And l.Tag = "RIO" Then    '''''�������ʵ����ֱ�עӦ�ø�һ�����ͱ���ʶ��
            l.Text = l.TagObject.ROILength & l.TagObject.ROIDistanceUnits
        End If
    Next
    ''''''����ע���ԭ���Ĳ�����Ϣд�뵽ͼ���һ��LABEL��,�Ա㱣�浽ͼ����
    f.SelectedImage.Refresh False
End Sub

Public Sub SubImageFiltering(strFilterString As String, img As DicomImage, Optional intTimes As Integer = 1)
'------------------------------------------------
'���ܣ���������ͼ����������Ե��ǿ��ƽ������͸�ԭ
'������ strFilterString--��ʾ�Ǿ��������͵��ַ�����
'       img--��Ҫ�����ͼ��
'       intTimes -- ͼ����Ĵ���
'���أ�ֱ�Ӷ�ͼ����д���
'------------------------------------------------
    If img Is Nothing Then Exit Sub
    If intTimes <= 0 Then Exit Sub
    
    Dim dblUnsharpEnhancement As Double
    Dim intUnsharpLength As Integer
    Dim intFilterLength As Integer
    
    dblUnsharpEnhancement = img.UnsharpEnhancement
    intUnsharpLength = img.UnsharpLength
    intFilterLength = img.FilterLength
    
    Select Case strFilterString
    Case "miUnSharpEnhancementUp"      '��Ե��ǿǿ�����ӣ�����0.1
        dblUnsharpEnhancement = dblUnsharpEnhancement + intTimes * 0.1
        If dblUnsharpEnhancement < 30 Then
            img.UnsharpEnhancement = dblUnsharpEnhancement
            If img.UnsharpLength = 0 Then img.UnsharpLength = 1
        End If
    Case "miUnSharpEnhancementDown"         '��Ե��ǿǿ�ȼ��٣�����0.1
        dblUnsharpEnhancement = dblUnsharpEnhancement - intTimes * 0.1
        If dblUnsharpEnhancement >= 0 Then
            img.UnsharpEnhancement = dblUnsharpEnhancement
            If img.UnsharpLength = 0 Then img.UnsharpLength = 1
        Else
            img.UnsharpEnhancement = 0
        End If
    Case "miUnSharpLengthUp"   '��Ե��ǿ�������ӣ�����1
        intUnsharpLength = intUnsharpLength + intTimes
        If intUnsharpLength < 30 Then
            img.UnsharpLength = intUnsharpLength
            If img.UnsharpEnhancement = 0 Then img.UnsharpEnhancement = 0.1
        End If
    Case "miUnSharpLengthDown"   '��Ե��ǿ���ȼ��٣�����1
        intUnsharpLength = intUnsharpLength - intTimes
        If intUnsharpLength >= 0 Then
            img.UnsharpLength = intUnsharpLength
            If img.UnsharpEnhancement = 0 Then img.UnsharpEnhancement = 0.1
        Else
            img.UnsharpLength = 0
        End If
    Case "miFilterLengthUp"       'ƽ�����ӣ�����1
        '�ж�Zoom�Ƿ�1������ǣ����޸�Ϊ0.9999
        If img.ActualZoom = 1 Then
            img.Zoom = 0.9999
        End If
        '�ж�ͼ���ڷŴ����Сģʽ���棬�Ƿ���doFilterMovingAverage��ֻ�����ģʽ�²ſ���ƽ��
        '��Сģʽ�£�Ĭ��ֵ����doFilterMovingAverage�������޸�
        img.MagnificationMode = doFilterMovingAverage
        
        '�ж�FilterLength�Ƿ�0����ǣ�����2/ActualZoom��2��FilterLength֮����е���
        If intFilterLength = 0 Then
            If intTimes = 1 Then
                img.FilterLength = 2 / img.ActualZoom + 1
            ElseIf intTimes > 1 Then
                img.FilterLength = 2 / img.ActualZoom + 1
                img.FilterLength = img.FilterLength + (intTimes - 1)
            End If
        Else    '���������FilterLength��1
            img.FilterLength = intFilterLength + intTimes
        End If
    Case "miFilterLengthDown"    'ƽ�����٣�����1
        '�ж�Zoom�Ƿ�1������ǣ����޸�Ϊ0.9999
        If img.ActualZoom = 1 Then
            img.Zoom = 0.9999
        End If
        '�ж�ͼ���ڷŴ����Сģʽ���棬�Ƿ���doFilterMovingAverage��ֻ�����ģʽ�²ſ���ƽ��
        '��Сģʽ�£�Ĭ��ֵ����doFilterMovingAverage�������޸�
        img.MagnificationMode = doFilterMovingAverage
        
        '�жϵ�ǰFilterLength��1�Ƿ�С�� 2/ActualZoom
        intFilterLength = intFilterLength - intTimes
        img.FilterLength = IIf(intFilterLength < 2 / img.ActualZoom, 0, intFilterLength)
    Case "miRestore"     'ͼ��ԭ
        img.UnsharpEnhancement = 0
        img.UnsharpLength = 0
        img.FilterLength = 0
    End Select
End Sub

Public Sub SubImageUnsharp(strFilterString As String, f As frmViewer)
'------------------------------------------------
'���ܣ�ͼ����������Ե��ǿ��ƽ������͸�ԭ
'������strFilterString--��ʾ�Ǿ��������͵��ַ�����f--ͼ����ǿ�Ĵ���
'���أ���
'2009��
'------------------------------------------------
    If f.SelectedImage Is Nothing Then Exit Sub
    
    Call SubImageFiltering(strFilterString, f.SelectedImage)
    
    f.Viewer(f.intSelectedSerial).Refresh
    ''''''''''''''''''''''''''''''''''''''''��������ͼ��ͬ��'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call subSeriesInPhase(f.intSelectedSerial, f, f.SelectedImage, IMG_SYN_FILTER)
End Sub

Public Sub subSetImageFortF(f As frmViewer)
'------------------------------------------------
'���ܣ����ݵ�ǰѡ�����е�����ʽ����������˵���ѡ��״̬
'������f--����Ĵ���
'���أ���
'------------------------------------------------
    Dim iSortType As Integer    '��¼��ǰ���е�����ʽ��0--ͼ��ţ�1--��λ����2--��λ����3--�ɼ�ʱ�䣻4--ͼ��ʱ�䣬����ZLShowSeriesInfos��ʹ�á�
    Dim lngToolID As Long
    
    On Error GoTo err
    
    If f.intSelectedSerial = 0 Then Exit Sub
    iSortType = ZLShowSeriesInfos(f.intSelectedSerial).intSortType
    
    Select Case iSortType
            Case 0            '����ͼ�������
                lngToolID = ID_View_PhotoSerial_PhotoNumber
            Case 1                 '���մ�λ��������
                lngToolID = ID_View_PhotoSerial_BedASC
            Case 2                '���մ�λ��������
                lngToolID = ID_View_PhotoSerial_BedDESC
            Case 3         '���ղɼ�ʱ������
                lngToolID = ID_View_PhotoSerial_CollectionTime
            Case 4              '����ͼ��ʱ������
                lngToolID = ID_View_PhotoSerial_PhotoTime
        End Select
        
    '����˵��͹�������check״̬
    f.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_PhotoSerial_PhotoNumber, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_PhotoSerial_BedASC, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_PhotoSerial_BedDESC, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_PhotoSerial_CollectionTime, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_PhotoSerial_PhotoTime, , True).Checked = False

    f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_PhotoSerial_PhotoNumber, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_PhotoSerial_BedASC, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_PhotoSerial_BedDESC, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_PhotoSerial_CollectionTime, , True).Checked = False
    f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_PhotoSerial_PhotoTime, , True).Checked = False

    'ѡ������ʽ
    f.ComToolBar.Item(ToolBar_Comm).FindControl(, lngToolID, , True).Checked = True
    f.ComToolBar.Item(ToolBar_Menu).FindControl(, lngToolID, , True).Checked = True
    f.ComToolBar.RecalcLayout
        
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub subMnuImageSort(lngToolID As Long, f As frmViewer)
'------------------------------------------------
'���ܣ�����ʽ��������lngToolID����������ͬʱ����˵���ѡ��״̬
'������Tool--��������Ĺ������ؼ���f--����Ĵ���
'���أ���
'2009��
'------------------------------------------------
    Dim iSortType As Integer        '0--ͼ��ţ�1--��λ����2--��λ����3--�ɼ�ʱ�䣻4--ͼ��ʱ�䡣
    
    '������е�����ʽ
    If lngToolID = ID_View_PhotoSerial_PhotoNumber Or lngToolID = ID_View_PhotoSerial_BedASC _
        Or lngToolID = ID_View_PhotoSerial_BedDESC Or lngToolID = ID_View_PhotoSerial_CollectionTime _
        Or lngToolID = ID_View_PhotoSerial_PhotoTime Then
        
        Select Case lngToolID
            Case ID_View_PhotoSerial_PhotoNumber            '����ͼ�������
                iSortType = 0
            Case ID_View_PhotoSerial_BedASC                 '���մ�λ��������
                iSortType = 1
            Case ID_View_PhotoSerial_BedDESC                '���մ�λ��������
                iSortType = 2
            Case ID_View_PhotoSerial_CollectionTime         '���ղɼ�ʱ������
                iSortType = 3
            Case ID_View_PhotoSerial_PhotoTime              '����ͼ��ʱ������
                iSortType = 4
        End Select
        
        'intSelectedSerial���ڣ����ұ�ѡ�е���������ͼ�񣬲Ž�������
        If f.intSelectedSerial > 0 And f.intSelectedSerial < f.MSFViewer.Rows And f.MSFViewer.TextMatrix(f.intSelectedSerial, 1) = "True" Then
            
            Call subSortImages(f, f.intSelectedSerial, iSortType)
            'ǿ���ù�����ˢ��һ��
            Call subShowALLImage(f, f.Viewer(f.intSelectedSerial), 1, False)
            f.VScro(f.intSelectedSerial).Value = 1
            
            '��������ʽ�����ò˵���ѡ
            Call subSetImageFortF(f)
        End If
    End If
End Sub

Public Sub subMouseRLset(ByVal control As CommandBarControl)
'------------------------------------------------
'���ܣ�����������Ҽ���check״̬��
'������Control--�������ؼ�
'���أ���
'2009��
'------------------------------------------------
    Dim i As Integer, j As Integer
    For i = 1 To cMouseUsage.Count
        If cMouseUsage(i).ButtomID = control.Id Then Exit For
    Next
    If i <= cMouseUsage.Count Then
        For j = 1 To cMouseUsage.Count
            If cMouseUsage(j).strProgramName <> "No" Then
                If cMouseUsage(j).ButtomID <> control.Id And cMouseUsage(i).lngMouseKey = cMouseUsage(j).lngMouseKey And cMouseUsage(i).lngShift = cMouseUsage(j).lngShift Then
                    control.Checked = False
                Else
                    control.Checked = True
                End If
            End If
        Next
    End If
End Sub

Public Sub subCurrentCheck(control As CommandBarControl, f As frmViewer)
'------------------------------------------------
'���ܣ���λ�ߴ�������λ�߲˵��ĵ����¼������ƶ�λ��������ذ�ťֻ�ܱ�ѡ��һ������ť״̬���浽��ʱ������
'������Control--�������Ĳ˵���f--��ʾ��λ�ߵĴ��塣
'���أ���
'2009��
'------------------------------------------------
    If control.Id = ID_Active_PointingLine_ALL Then
        f.ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_FirstLast, , True).Checked = False
        f.ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_Now, , True).Checked = False
    End If
    If control.Id = ID_Active_PointingLine_FirstLast Then
        f.ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_ALL, , True).Checked = False
    End If
    If control.Id = ID_Active_PointingLine_Now Then
        f.ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_ALL, , True).Checked = False
    End If
    f.ComToolBar.Item(ToolBar_Plane).FindControl(, control.Id, , True).Checked = Not control.Checked
    
    Button_miAllReferLine = f.ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_ALL).Checked
    Button_miFLReferLine = f.ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_FirstLast).Checked
    Button_miCurrentReferLine = f.ComToolBar.Item(ToolBar_Plane).FindControl(, ID_Active_PointingLine_Now).Checked
    
    '��������˽�Ƭ��ӡ״̬�����ڽ�Ƭ��ӡ������Ҳ��ʾ��λ��
    If f.blnPrintFilm = True And Not f.mfrmFilm Is Nothing Then
        f.mfrmFilm.subDispReferLineFilm
    End If
    
    '��ʾ��λ��
    Call subDisplayReferLine(f.Viewer(f.intSelectedSerial), f, False)
End Sub

Public Sub subOutputToPowerPoint(f As frmViewer)
'------------------------------------------------
'���ܣ������POWERPOINT
'������f--����
'���أ���
'------------------------------------------------
    Dim v As DicomViewer
    Dim im As DicomImage
    Dim imgs As New DicomImages
    Dim iW, iH As Integer               'ԭʼͼ��Ŀ�͸�
    Dim ix, iy As Integer               '���ڵ��������
    Dim Nw, Nh As Integer               '���ڵĿ�͸�
    Dim intCol, intRow As Integer       '��ǰ����������
    Dim NowImg As Integer               '��ǰ��ͼƬ
    Dim ImgCount As Integer             '����ʾ��ͼ����
    Dim ShowImg As Integer              '����ʾ��ͼ�����ʼλ��
    Dim j As Integer                    'ѭ������
    Dim z As Integer                    '��ʱ����
    Dim x As Integer                    'ѭ������
    Dim i As Integer                    'ѭ������
    Dim JSCount As Integer              '��¼λ��
    Dim TwoBegin                        '�ڶ�����ʼλ��
    Dim PageCount As Integer            'һҳ��������
    Dim ppt As Object                   'PowerPoint����
    Dim blnHaveImage As Boolean
    
    '��PowerPoint
    Set ppt = CreateObject("PowerPoint.Application")
    
    '�ж��Ƿ���ͼ��
    For Each v In f.Viewer
        If v.Index <> 0 And v.Visible Then
            For Each im In v.Images
                If im.Tag <> "" Then
                    blnHaveImage = True
                    Exit For
                End If
            Next
            If blnHaveImage = True Then Exit For
        End If
    Next
    
    If blnHaveImage = False Then
        MsgBox "��ǰû��ѡ���κ�ͼ��,�������!", vbInformation, gstrSysName
        Exit Sub
    End If
    '��ʹ��
    ppt.Visible = True
    ppt.Presentations.Add 1
    ppt.ActiveWindow.view.GotoSlide (ppt.ActivePresentation.Slides.Add(1, 12).SlideIndex)
    
    '��ʹ��λ��Ϊ1
    JSCount = 1
    
    For x = 1 To f.Viewer.Count - 1
        If f.Viewer(x).Index <> 0 And f.Viewer(x).Visible Then
            imgs.Clear
            'д��ͼ��
            For Each im In f.Viewer(x).Images
                If im.Tag <> "" Then imgs.Add im
            Next
            If imgs.Count <> 0 Then
                PageCount = f.Viewer(x).MultiColumns * f.Viewer(x).MultiRows
                j = 1
                For Each im In imgs
                    If j > PageCount Then
                        z = ppt.ActivePresentation.Slides.Count
                        ppt.ActiveWindow.view.GotoSlide (ppt.ActivePresentation.Slides.Add(z + 1, 12).SlideIndex)
                        j = 1
                    End If
                    im.Copy
                    ppt.ActiveWindow.view.Paste
                    j = j + 1
                Next
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                For i = JSCount To ppt.ActivePresentation.Slides.Count
                    ppt.ActiveWindow.view.GotoSlide i
                    If ppt.ActiveWindow.Selection.SlideRange.Shapes.Count > 0 Then
                        With ppt.ActiveWindow.Selection.SlideRange
                            ix = .Shapes(1).left
                            iy = .Shapes(1).top
                            iW = .Shapes(1).width / f.Viewer(x).MultiColumns
                            iH = .Shapes(1).height / f.Viewer(x).MultiRows
                        End With
                        For j = 1 To ppt.ActiveWindow.Selection.SlideRange.Shapes.Count
                            '�õ���ǰͼ��λ��
                            If (j Mod f.Viewer(x).MultiColumns) = 0 Then
                                intRow = j / f.Viewer(x).MultiColumns
                                intCol = f.Viewer(x).MultiColumns
                            Else
                                intRow = Int(j / f.Viewer(x).MultiColumns) + 1
                                intCol = j Mod f.Viewer(x).MultiColumns
                            End If
                            '�ƶ�ͼ��λ��
                            With ppt.ActiveWindow.Selection.SlideRange
                                .Shapes(j).top = iy + (iH * (intRow - 1))
                                .Shapes(j).left = ix + (iW * (intCol - 1))
                                If iH > iW Then
                                    .Shapes(j).width = iW
                                Else
                                    .Shapes(j).height = iH
                                End If
                            End With
                        Next
                        JSCount = JSCount + 1
                    End If
                Next
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '���һ��ʱ��������ҳ
                If x < f.Viewer.Count - 1 Then
                    z = ppt.ActivePresentation.Slides.Count
                    ppt.ActiveWindow.view.GotoSlide (ppt.ActivePresentation.Slides.Add(z + 1, 12).SlideIndex)
                End If
            End If
        End If
    Next
    ppt.ActiveWindow.view.GotoSlide 1
End Sub

Public Sub subDSA(thisForm As frmViewer)
'------------------------------------------------
'���ܣ�DSA���ּ�Ӱ
'������thisForm--�������ּ�Ӱ�Ĵ��塣
'���أ���
'2009��
'------------------------------------------------
    If thisForm.SelectedImage Is Nothing Then Exit Sub             ''''��ǰû��ѡ��ͼ��
    If thisForm.SelectedImage.FrameCount <= 1 Then Exit Sub        ''''��ǰͼ���Ƕ���
    If IsNull(thisForm.SelectedImage.Attributes(&H28, &H4).Value) Then Exit Sub
    If Mid(thisForm.SelectedImage.Attributes(&H28, &H4).Value, 1, 4) <> "MONO" Then Exit Sub
    
    Call FrmDSAConfig.zlShowMe(thisForm.SelectedImage.FrameCount, thisForm.SelectedImage.Frame, thisForm)
End Sub

Public Sub subCutOut(f As frmViewer)
'------------------------------------------------
'���ܣ�������˳��ü�״̬�����ػ���ʾ�ü���ע
'������f--������˳��ü�״̬�Ĵ���
'���أ��ޣ�ֱ�ӿ��Ʋü�״̬��ע����ʾ������
'------------------------------------------------
    Dim i As Integer, im As DicomImage
    If f.SelectedImage Is Nothing Then Exit Sub
    ''''''''''''[�˳��ü�]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not Button_miCutOut Then
        For i = 1 To 5
            f.SelectedImage.Labels(i).Visible = False '
            f.SelectedImage.Labels(i).Tag = f.SelectedImage.Labels(i).left & "_" & f.SelectedImage.Labels(i).top    ''''��¼ԭʼ״̬�����ٴ���ʾ��
            f.SelectedImage.Labels(i).left = G_INT_SYS_LABEL_HIDE_LEFT
            f.SelectedImage.Labels(i).top = G_INT_SYS_LABEL_HIDE_TOP
        Next
        If Not f.SelectedLabel Is Nothing Then
            If f.SelectedImage.Labels.IndexOf(f.SelectedLabel) = 1 Then          ''''�����ǰѡ����ǲü���ע����ȡ����ʾ���
                SubNoDispPeriod f.SelectedImage, f      'Ϊָ��ͼ�����ر�עѡ����
                Set f.SelectedLabel = Nothing
            End If
        End If
    Else
        ''''''''''''[����ü�]'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If f.SelectedImage.Labels(1).Tag <> "" Then         '''''�����ǰ�ü�����ظ���ǰ����ʾλ��
            For i = 1 To 5
                f.SelectedImage.Labels(i).left = Val(left(f.SelectedImage.Labels(i).Tag, InStr(f.SelectedImage.Labels(i).Tag, "_") - 1))
                f.SelectedImage.Labels(i).top = Val(Right(f.SelectedImage.Labels(i).Tag, Len(f.SelectedImage.Labels(i).Tag) - InStr(f.SelectedImage.Labels(i).Tag, "_")))
            Next
        Else
            f.SelectedImage.Labels(1).left = 4
            f.SelectedImage.Labels(1).top = 4
            f.SelectedImage.Labels(1).width = f.SelectedImage.sizex - 8
            f.SelectedImage.Labels(1).height = f.SelectedImage.sizey - 8
        End If
        SubDispPeriod f.SelectedImage.Labels(1), f.SelectedImage, f 'Ϊָ��ͼ���е�ָ����ע����ʾ��עѡ����
        For i = 1 To 5
            f.SelectedImage.Labels(i).Visible = True
        Next
        Set f.SelectedLabel = f.SelectedImage.Labels(1)
    End If
    
    ''''''''''''''�ڲü�״̬�¶Բü���������ͼ��ͬ������'''''''''''''''
    If Button_miImageInPhase = True Then subCutOutInphase f.Viewer(f.intSelectedSerial), f.SelectedImage, f
    f.SelectedImage.Refresh False
    f.ComToolBar.RecalcLayout
End Sub

Public Sub subDispLabelInfo(f As frmViewer)
'------------------------------------------------
'���ܣ���ʾ������ͼ����û���ע��Ϣ
'������ f--��Ҫ��ʾ�������û���ע��Ϣ�Ĵ��壻
'���أ��ޣ�ֱ����ʾ�������û���ע��
'2009��
'------------------------------------------------
    Dim img As DicomImage
    Dim v As DicomViewer
    Dim l As DicomLabel
    Dim i As Integer
    Dim CmdControl As CommandBarControl
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If f.SelectedImage Is Nothing Then Exit Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Button_miDispLabelInfo = Not Button_miDispLabelInfo
    f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_LableShow, , True).Checked = Button_miDispLabelInfo
    f.ComToolBar.RecalcLayout
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each v In f.Viewer
        If v.Index <> 0 Then
            For Each img In v.Images
                If img.Labels.Count > G_INT_SYS_LABEL_COUNT Then
                    For i = G_INT_SYS_LABEL_COUNT + 1 To img.Labels.Count
                        If Button_miDispLabelInfo Then
                            img.Labels(i).Visible = True
                        Else
                            img.Labels(i).Visible = False
                        End If
                    Next i
                    If Not Button_miDispLabelInfo Then
                        For i = 11 To 18                    '���ر�ע���
                            img.Labels(i).Visible = False
                        Next i
                    End If
                End If
            Next
        End If
        v.Refresh
    Next
End Sub

Public Sub subManipulation(strOperation As String, f As frmViewer)
'------------------------------------------------
'���ܣ����ö�ƽ�洦��ͼ����ת�����׵ȣ�
'������ strOperation--��ʾ��ת��ʽ���ַ����� f--���塣
'���أ���
'2009��
'------------------------------------------------
    If f.SelectedImage Is Nothing Then Exit Sub
    
    On Error GoTo err
    
    Call subFlipRotate(f.SelectedImage, strOperation)
    
    ''''''''''''''''''''������ͼ��ͬ��'''''''''''''''''''''''''''
    Select Case strOperation
    Case "FlipHorizontal", "FlipVertical"
        Call subSeriesInPhase(f.intSelectedSerial, f, f.SelectedImage, IMG_SYN_FLIP)
    Case "RotateAnticlockwise", "RotateClockwise"
        Call subSeriesInPhase(f.intSelectedSerial, f, f.SelectedImage, IMG_SYN_ROTATE)
    Case "Invert"
        Call subSeriesInPhase(f.intSelectedSerial, f, f.SelectedImage, IMG_SYN_WINDOW)
    End Select
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub subFlipRotate(img As DicomImage, strOperation As String)
'------------------------------------------------
'���ܣ���ƽ�洦��ͼ����ת�����׵ȣ�
'������img�������д����ͼ��
'      strOperation --��ʾ��ת��ʽ���ַ���
'���أ���
'2009��
'------------------------------------------------
    With img
        Select Case strOperation
        Case "FlipHorizontal"
            .FlipState = .FlipState Xor 1
            If .RotateState = doRotateLeft Or .RotateState = doRotateRight Then
                .RotateState = (.RotateState + 2) Mod 4
            End If
        Case "FlipVertical"
            .FlipState = .FlipState Xor 2
            If .RotateState = doRotateLeft Or .RotateState = doRotateRight Then
                .RotateState = (.RotateState + 2) Mod 4
            End If
        Case "RotateAnticlockwise"
            .RotateState = (.RotateState + 1) And 3
        Case "RotateClockwise"
            .RotateState = (.RotateState + 3) And 3
        Case "Invert"
            If .VOILUT = 1 Then .VOILUT = 0
            .width = -.width
        End Select
    End With
End Sub

Public Sub subSelectAllSerial(f As frmViewer)
'------------------------------------------------
'���ܣ�ѡ����������
'������f--ѡ�����еĴ���
'���أ���
'2009��
'------------------------------------------------
    Dim v As DicomViewer
    f.isSelectAllSerial = Not f.isSelectAllSerial
    
    For Each v In f.Viewer
        If v.Visible = True Then
            ZLShowSeriesInfos(v.Index).Selected = IIf(f.isSelectAllSerial, True, False)
            subDispframe f, v
            v.Refresh
        End If
    Next
End Sub
 
Public Sub subSelectAllIMage(f As frmViewer)
'------------------------------------------------
'���ܣ�ѡ������ͼ��
'������f--ѡ��ͼ��Ĵ���
'���أ���
'2009��
'------------------------------------------------
    Dim v As DicomViewer
    Dim i As Integer
    
    f.isSelectAllImage = Not f.isSelectAllImage
    
    For Each v In f.Viewer
        If v.Visible = True Then
            If v.Images.Count > 0 And (v.Index = f.intSelectedSerial Or ZLShowSeriesInfos(v.Index).Selected = True) Then
                For i = 1 To ZLShowSeriesInfos(v.Index).ImageInfos.Count
                    ZLShowSeriesInfos(v.Index).ImageInfos(i).blnSelected = IIf(f.isSelectAllImage, True, False)
                Next i
                subDispframe f, v
                v.Refresh
            End If
        End If
    Next
End Sub

Public Sub subFullScreen(Frm As frmViewer)
'------------------------------------------------
'���ܣ��л���Ļ��ȫ��״̬
'������Frm--����ȫ���л��Ĵ���
'���أ���
'2009��
'------------------------------------------------
    Dim CmdControl As CommandBar
    Dim ToolBarTop As Long
    Dim ToolBarLeft As Long
    Dim ToolBarHeight As Long
    Dim ToolBarWidth As Long
    Dim i As Integer
    blfrmRefresh = False
    '''''''''''''''''''''''''''''''''''''''[����ȫ��״̬�Ĵ���]''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not Button_miFullScreen Then
        Frm.WindowState = vbMaximized
        Frm.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_FullScreen, , True).Checked = True
        Frm.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_FullScreen, , True).Checked = True
        Frm.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_FullScreen, , True).ToolTipText = "ȡ��ȫ����ʾ"
        Frm.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_FullScreen, , True).ToolTipText = "ȡ��ȫ����ʾ"
        Button_miFullScreen = True
        '''''''''''''''''''''''''''''''''''''''[�رմ���ı�����]'''''''''''''''''''''''''''''''''''''''''''''
        
        ''''���ô����Ƿ���ʾ������
        Call zlcontrol.FormSetCaption(Frm, False)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '���ع������Ͳ˵�
        For i = 1 To 8
            If i <> 7 Then
                Frm.ComToolBar.Item(i).Visible = False
            End If
        Next
        '����״̬��
        blfrmRefresh = True
        Frm.sbStatusBar.Visible = False
        
        Set CmdControl = Frm.ComToolBar.Item(7)
        CmdControl.GetWindowRect ToolBarLeft, ToolBarTop, ToolBarWidth, ToolBarHeight
        
        Frm.ComToolBar.DockToolBar CmdControl, Frm.left, Frm.height + Frm.top - Frm.sbStatusBar.height - (ToolBarHeight - ToolBarTop), xtpBarFloating
        Frm.ComToolBar.RecalcLayout
    Else
        ''''''''''''''''''''''''''''''''''''''[�ظ�������Ļ]''''''''''''''''''''''''''''''''''''''''''''''
        Frm.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_FullScreen, , True).Checked = False
        Frm.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_FullScreen, , True).Checked = False
        Frm.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_FullScreen, , True).ToolTipText = "ȫ����ʾ"
        Frm.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_FullScreen, , True).ToolTipText = "ȫ����ʾ"
        Button_miFullScreen = False
        '''''''''''''''''''''''''''''''''''''''[�رմ���ı�����]'''''''''''''''''''''''''''''''''''''''''''''
        ''''���ô����Ƿ���ʾ������
        Call zlcontrol.FormSetCaption(Frm, True)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''[��������ı�����]''''''''''''''''''''''''''''''''''''''''''''''
        '��ʾ�������Ͳ˵�
        For i = 1 To 8
            Frm.ComToolBar.Item(i).Visible = True
        Next
        '��ʾ״̬��
        Frm.sbStatusBar.Visible = True
        Frm.ComToolBar.Item(7).Position = Frm.ComToolBar.Item(2).Position
        Frm.ComToolBar.RecalcLayout
        '���°�һ����˳��ڷŹ�����λ��
        ArrayToolBar Frm.ComToolBar, Frm.top, Frm.left, Frm.height, Frm.width
        blfrmRefresh = True
    End If
End Sub

Public Function subSaveImage(img As DicomImage, strOldSeriesUID As String) As Boolean
'------------------------------------------------
'���ܣ���ͼ�񱣴浽���ݿ���
'������img �����ͼ��
'���أ�True---��ȷ���棻False ---���ִ���û�б���ͼ��
'------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    
    Dim dtReceived As String
    Dim strStudyUID As String
    Dim blnFirstImage As String     '�Ƿ񱾴μ��ĵ�һ��ͼ��
    Dim lngResult As String         'FTP�������
    Dim NowTime As Date
    Dim strSQL As String
    
    Dim strFTPDir As String
    Dim strFTPIp As String
    Dim strFTPUser As String
    Dim strFTPPassw As String
    Dim Inet As New clsFtp             'FTP��
    
    Dim arrSQL() As Variant         '�����е�SQL�������
    Dim blnInTrans As Boolean       '�Ƿ�����������Ĺ�����
    Dim i As Integer
    
    subSaveImage = False
    
    If img Is Nothing Then
        MsgBox "ƴ�ӵĽ��ͼ�����޷����档", vbOKOnly, "��ʾ��Ϣ"
        Exit Function
    End If
    
    '���ж�ͼ���Ƿ��Ѿ�����
    
    strSQL = "Select ͼ��UID From Ӱ����ͼ�� Where ͼ��UID=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ͼ���Ƿ����", CStr(img.InstanceUID))
    If rsTmp.EOF = False Then
        MsgBox "���ݿ����Ҳ�����ͼ���޷������ⲿֱ�Ӵ򿪵���ʱͼ��", vbOKOnly, "��ʾ��Ϣ"
        Exit Function
    End If
    
    '�ȱ���FTPͼ��
    '��ȡ��������
    strSQL = "select a.��������,a.���UID  from Ӱ�����¼ a,Ӱ�������� b where a.���UID =b.���UID and b.����UID =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���UID", strOldSeriesUID)
    
    If rsTmp.EOF = True Then
        MsgBox "���ݿ����Ҳ��������У��޷������ⲿֱ�Ӵ򿪵���ʱͼ��", vbOKOnly, "��ʾ��Ϣ"
        Exit Function '��ѯ������¼�����˳�����
    End If
    
    NowTime = zlDatabase.Currentdate
    strStudyUID = rsTmp("���UID")
    dtReceived = Format(rsTmp("��������"), "yyyyMMdd")
     
    '����ͼ�񵽻���Ŀ¼
    MkLocalDir PstrBufferImagePath & dtReceived & "/" & strStudyUID & "/"
    '����������ԭͼ��ѹ����ʽ
    img.WriteFile PstrBufferImagePath & dtReceived & "/" & strStudyUID & "/" & img.InstanceUID, True
    
    '����FTP
    Call funGetStorageDevice(strStudyUID, strFTPDir, strFTPIp, strFTPUser, strFTPPassw)
    lngResult = Inet.FuncFtpConnect(strFTPIp, strFTPUser, strFTPPassw)
    
    '����ͼ���ļ�
    If lngResult = 0 Then
        'FTP����ʧ�ܣ���ʾ���󣬲�ɾ������ͼ�е�ͼ��
        MsgBox "FTP����ʧ�ܣ�ͼ���޷����棬�����������á�", vbInformation, gstrSysName
        Exit Function
    Else
        '��FTP�д���Ŀ¼
        Inet.FuncFtpMkDir "/", strFTPDir
        
        '��FTP�ϴ��ļ�
        Inet.FuncUploadFile strFTPDir, _
             PstrBufferImagePath & dtReceived & "/" & strStudyUID & "/" & img.InstanceUID, img.InstanceUID
    End If
    Inet.FuncFtpDisConnect
    
    'ͼ��洢�ɹ��󣬴洢���ݿ���Ϣ
    On Error GoTo DBError
    arrSQL = Array()
    
    strSQL = "Select ����UID From Ӱ��������  Where ����UID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PACSͼ�񱣴�", CStr(img.SeriesUID))
    '�����µļ������
    If rsTmp.EOF Then
        strSQL = "ZL_Ӱ������_INSERT('" & strStudyUID & "','" & img.SeriesUID & "','" & _
            img.SeriesDescription & "',0)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
    End If
    
    '�����µ�ͼ��
    strSQL = "ZL_Ӱ��ͼ��_INSERT('" & img.InstanceUID & "','" & img.SeriesUID & "','" & _
        img.SeriesDescription & "',0," & IIf(GetImageAttribute(img.Attributes, ATTR_ͼ���) = "", 0, GetImageAttribute(img.Attributes, ATTR_ͼ���)) & ","
    If GetImageAttribute(img.Attributes, ATTR_�ɼ�����) <> "" And GetImageAttribute(img.Attributes, ATTR_�ɼ�ʱ��) <> "" Then
        strSQL = strSQL & "to_Date('" & Format(GetImageAttribute(img.Attributes, ATTR_�ɼ�����) & " " & GetImageAttribute(img.Attributes, ATTR_�ɼ�ʱ��), "yyyy-MM-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS'),"
    Else
        strSQL = strSQL & " sysdate,"
    End If
    
    If GetImageAttribute(img.Attributes, ATTR_ͼ������) <> "" And GetImageAttribute(img.Attributes, ATTR_ͼ��ʱ��) <> "" Then
        strSQL = strSQL & "to_Date('" & Format(GetImageAttribute(img.Attributes, ATTR_ͼ������) & " " & GetImageAttribute(img.Attributes, ATTR_ͼ��ʱ��), "yyyy-MM-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS'),'"
    Else
        strSQL = strSQL & " sysdate,'"
    End If
    
        strSQL = strSQL & GetImageAttribute(img.Attributes, ATTR_���) & "','" _
        & GetImageAttribute(img.Attributes, ATTR_ͼ��λ�ò���) & "','" _
        & GetImageAttribute(img.Attributes, ATTR_ͼ������) & "','" _
        & GetImageAttribute(img.Attributes, ATTR_�ο�֡UID) & "','" _
        & GetImageAttribute(img.Attributes, ATTR_��Ƭλ��) & "','" _
        & GetImageAttribute(img.Attributes, ATTR_����) & "','" _
        & GetImageAttribute(img.Attributes, ATTR_����) & "','" _
        & GetImageAttribute(img.Attributes, ATTR_���ؾ���) & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
    
    '��ʼ������
    gcnOracle.BeginTrans
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "����ͼ��")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    
    subSaveImage = True
    
    Exit Function
DBError:
    If blnInTrans = True Then gcnOracle.RollbackTrans
    Inet.FuncFtpDisConnect
    err.Raise err.Number, "���ͼ�񱣴�"
End Function


