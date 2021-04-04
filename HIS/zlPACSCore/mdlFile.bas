Attribute VB_Name = "mdlFile"
Option Explicit
'--------------------------------------------------------
'��  �ܣ����ļ��򿪣�����Ŀ¼����صĺ���
'�����ˣ����������Σ��ƽ�
'�������ڣ�2004.6.12
'-------------------------------------------------------

Public Type DlgFileInfo
    iCount As Long
    sPath As String
    sFile() As String
End Type
'���ļ��Զ�����������
Public Type OpenFileArray
    FilePath As String
    Filename() As String
End Type

Public Sub SaveImages(objImages As DicomImages, ByVal SaveMode As Integer)
'------------------------------------------------
'���ܣ���ͼ����Ϣ���浽�洢��������
'������
    'SaveMode��0-ֻ��Dicomͼ��
    '          1-ֻ�汨��ͼ��
    '          2-������
'���أ�
'------------------------------------------------
    Dim Inet As New clsFtp
    Dim strTempPath As String, lngBuffSize As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strDirURL As String, strIp As String, strUser As String, strPwd As String
    Dim RptImageName As String
    Dim strSeriesUID As String
    Dim strStudyUID As String
    Dim img As DicomImage
    Dim dgGlobal As New DicomGlobal
    Dim strReportImages As String
    
    On Error GoTo DBError
    
    '������ʱ�ļ�
    strTempPath = Space(255)
    lngBuffSize = GetTempPath(Len(strTempPath), strTempPath)
    strTempPath = Mid(strTempPath, 1, InStrRev(strTempPath, "\"))
    dgGlobal.RegString("UIDRoot") = "1"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'ѭ������ÿһ��ͼ��
    For Each img In objImages
        '��ȡͼ��ı���ͼ����
        RptImageName = dgGlobal.NewUID & ".jpg"
        
        '������ʱ�ļ�
        If SaveMode <> 1 Then       '��Dicomͼ��
            img.WriteFile strTempPath & img.InstanceUID, True
        End If
        
        If SaveMode <> 0 Then       '�汨��ͼ��
            img.FileExport strTempPath & RptImageName, "JPG"
        End If
        
        '��ȡFTP����·��
        If strSeriesUID = "" Or strStudyUID = "" Or strSeriesUID <> img.SeriesUID Then
            '��ȡͼ�������ݿ��ж�Ӧ�ļ��UID
            strSeriesUID = img.SeriesUID
            strSQL = "select ���UID FROM Ӱ�������� where ����UID =[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���UID", CStr(img.SeriesUID))
            If rsTemp.RecordCount = 0 Then
                strStudyUID = PstrCheckUID  '��Ĭ��ֵ
            Else
                strStudyUID = rsTemp!���UID
            End If
        End If
        Call funGetStorageDevice(strStudyUID, strDirURL, strIp, strUser, strPwd)
        Inet.FuncFtpConnect strIp, strUser, strPwd
        
        If SaveMode <> 1 Then
            Inet.FuncUploadFile strDirURL, strTempPath & img.InstanceUID, img.InstanceUID
            Kill strTempPath & img.InstanceUID
        End If
        
        '���汨��ͼ
        If SaveMode <> 0 Then
            '��鱨��ͼ�����Ƿ񳬳�
            strSQL = "Select ����ͼ�� From Ӱ�����¼ Where ���UID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ͼ��", strStudyUID)
            If rsTemp.RecordCount > 0 Then
                strReportImages = Nvl(rsTemp("����ͼ��"))
                If Len(strReportImages & " ;" & RptImageName) >= 4000 Then
                    MsgBox "����ͼ�������������ޣ�����ɾ�����ֱ���ͼ���ټ������汨��ͼ��", vbInformation, gstrSysName
                Else
                    Inet.FuncUploadFile strDirURL, strTempPath & RptImageName, RptImageName
                    Kill strTempPath & RptImageName
                    
                    strSQL = "ZL_Ӱ���鱨��_ADD('" & strStudyUID & "','" & RptImageName & "')"
                    zlDatabase.ExecuteProcedure strSQL, "���汨��ͼ��"
                End If
            End If
        End If
        
        Inet.FuncFtpDisConnect
    Next
    Exit Sub
DBError:
    Inet.FuncFtpDisConnect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function funGetStorageDevice(strStudyUID As String, ByRef strDirURL As String, ByRef strIp As String, _
        ByRef strUser As String, ByRef strPwd As String) As Boolean
'------------------------------------------------
'���ܣ������ݿ��ж�ȡ�ƶ��洢�豸ID��FTP���ʲ���
'������ strSaveDeviceID �����洢�豸ID
'       strDirURL����[OUT] FTPĿ¼
'       strIp ����[OUT] IP��ַ
'       strUser ���� [OUT]�û���
'       strPwd ����[OUT]�û���
'���أ�True������ȡ�ɹ���False������ȡʧ��
'-----------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    '���洢�豸�Ƿ����
    strSQL = "Select b.��������, '/'||Decode(c.FtpĿ¼,Null,'',c.FtpĿ¼||'/') As Ŀ¼1,c.FTP�û��� As �û���1,c.FTP���� As ����1,c.IP��ַ As IP��ַ1," & _
             " '/'||Decode(d.FtpĿ¼,Null,'',d.FtpĿ¼||'/') As Ŀ¼2,d.FTP�û��� As �û���2,d.FTP���� As ����2,d.IP��ַ As IP��ַ2 " & _
             " from Ӱ�����¼ b,Ӱ���豸Ŀ¼  c ,Ӱ���豸Ŀ¼ d " & _
             " where (B.λ��һ = C.�豸�� And b.λ�ö�=d.�豸��(+) )  And b.���UID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strStudyUID)
     'û�д洢�豸ʱ�˳�
    If rsTemp.EOF = True Then
        MsgBox "û���ҵ��洢�豸,��ͼ��������ⲿͼ��", vbInformation, App.ProductName
        funGetStorageDevice = False
        Exit Function
    End If
    strDirURL = Nvl(rsTemp("Ŀ¼1"))
    strIp = Nvl(rsTemp("IP��ַ1"))
    strUser = Nvl(rsTemp("�û���1"))
    strPwd = Nvl(rsTemp("����1"))
    If strIp = "" Or strUser = "" Then  'λ��һ����û��ͼ�񣬶�ȡλ�ö���ͼ��
        strDirURL = Nvl(rsTemp("Ŀ¼2"))
        strIp = Nvl(rsTemp("IP��ַ2"))
        strUser = Nvl(rsTemp("�û���2"))
        strPwd = Nvl(rsTemp("����2"))
    End If
    strDirURL = strDirURL & Format(Nvl(rsTemp("��������")), "YYYYMMDD") & "/" & strStudyUID & "/"
    funGetStorageDevice = True
End Function

Function funIsLabelMouse(f As frmViewer, Button As Integer, Shift As Integer) As Boolean
'------------------------------------------------
'���ܣ��жϱ�ע����־�Ƿ���
'������f--��굥���Ĵ��壻Button--�������Ҽ���ţ�Shift--����shift״̬��
'���أ�True-�ɹ���False-ʧ�ܡ�
'------------------------------------------------
    funIsLabelMouse = False
    If Button_miFrameSelectImage And Button = cMouseUsage("201").lngMouseKey And Shift = cMouseUsage("201").lngShift Then        '��ѡͼ��ʹ�þ��α�ע
        '��ѡͼ��LabelStyle=����
        funIsLabelMouse = True
        intSelectLabelStyle = 2
    ElseIf Button_miLabelRectangle And Button = cMouseUsage("2").lngMouseKey And Shift = cMouseUsage("2").lngShift Then
        '���α�ע��LabelStyle=����
        funIsLabelMouse = True
        intSelectLabelStyle = 2
    ElseIf Button_miLabelLine And Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift Then
        'ֱ�߱�ע��LabelStyle=ֱ��
        funIsLabelMouse = True
        intSelectLabelStyle = 3
    ElseIf Button_miLabelVasMeasure And Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift Then
        'Ѫ����խ������LabelStyle=ֱ��
        funIsLabelMouse = True
        intSelectLabelStyle = 3
    ElseIf Button_miLabelCadiothoracicRatio And Button = cMouseUsage("1").lngMouseKey And Shift = cMouseUsage("1").lngShift Then
        '���رȲ�����LabelStyle=ֱ��
        funIsLabelMouse = True
        intSelectLabelStyle = 3
    ElseIf Button_miLabelEllipse And Button = cMouseUsage("3").lngMouseKey And Shift = cMouseUsage("3").lngShift Then
        '��Բ��ע��LabelStyle=��Բ
        funIsLabelMouse = True
        intSelectLabelStyle = 1
    ElseIf Button_miLabelArrowhead And Button = cMouseUsage("4").lngMouseKey And Shift = cMouseUsage("4").lngShift Then
        '��ͷ��ע��LabelStyle=��ͷ
        funIsLabelMouse = True
        intSelectLabelStyle = 10
    ElseIf Button_miLabelPolygon And Button = cMouseUsage("5").lngMouseKey And Shift = cMouseUsage("5").lngShift Then
        '����α�ע��LabelStyle=�����
        funIsLabelMouse = True
        intSelectLabelStyle = 5
    ElseIf Button_miLabelPolyLine And Button = cMouseUsage("6").lngMouseKey And Shift = cMouseUsage("6").lngShift Then
        '����߱�ע��LabelStyle=�����
        funIsLabelMouse = True
        intSelectLabelStyle = 4
    ElseIf Button_miLabelAngle And Button = cMouseUsage("7").lngMouseKey And Shift = cMouseUsage("7").lngShift Then
        '�Ƕȱ�ע��LabelStyle=ֱ��
        funIsLabelMouse = True
        intSelectLabelStyle = 3
    ElseIf Button_miAutoWidthLevel And Button = cMouseUsage("105").lngMouseKey And Shift = cMouseUsage("105").lngShift Then
        '�Զ�����λ�ľ��ο�LabelStyle=����
        funIsLabelMouse = True
        intSelectLabelStyle = 2
    End If
End Function

Public Function funGetFileList(f As frmViewer) As OpenFileArray
'------------------------------------------------
'���ܣ��򿪶�ȡ�ļ��Ի���,����ȫ·�����ļ�����
'������f--���塣
'���أ�ȫ·���ļ�������
'2009��
'------------------------------------------------
    Dim DlgInfo As DlgFileInfo
    Dim i As Integer
    On Error GoTo errHandle
    'ѡ���ļ�
    With f.Common
        
        .CancelError = False
        .MaxFileSize = 32767 '���򿪵��ļ����ߴ�����Ϊ��󣬼�32K
        .Flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
        .DialogTitle = "ѡ���ļ�"
        .Filter = "DICOM�ļ���*.dcm��(*.img)|*.dcm;*.img|ͼ���ļ� (*.BMP)(*.JPG)|*.BMP;*.JPG|�����ļ���*.*��|*.*"
        .ShowOpen
        If .Filename <> "" Then
            DlgInfo = GetDlgSelectFileInfo(.Filename)
        End If
        .Filename = ""      '�ڴ���*.pif�ļ����뽫Filename�����ÿգ�
                            '����ѡȡ���*.pif�ļ��󣬵�ǰ·����ı�
    End With
    
    If DlgInfo.iCount <= 0 Then
        ReDim funGetFileList.Filename(0)
        funGetFileList.FilePath = ""
        Exit Function
    End If
    
    ReDim funGetFileList.Filename(DlgInfo.iCount)
    funGetFileList.FilePath = DlgInfo.sPath
    For i = 1 To DlgInfo.iCount
        funGetFileList.Filename(i) = DlgInfo.sFile(i)
    Next i
    Exit Function
errHandle:
    ReDim funGetFileList.Filename(0)
    funGetFileList.FilePath = ""
    MsgBox "��ͼ����������ԡ�������һ���Դ򿪵�ͼ���������ࡣ", vbExclamation, gstrSysName
End Function

Public Function GetDlgSelectFileInfo(strFileName As String) As DlgFileInfo
'------------------------------------------------
'���ܣ����ļ���ת��Ϊȫ·������
'������strFileName--�ļ�����ͨ�����ļ��ؼ�����á�
'���أ�ȫ·������
'�����ˣ�����
'------------------------------------------------
    Dim sPath, tmpStr As String
    Dim sFile() As String
    Dim iCount, i As Integer
    On Error GoTo errHandle
    sPath = CurDir()  '��õ�ǰ��·������Ϊ��CommonDialog�иı�·��ʱ��ı䵱ǰ��Path
    tmpStr = Right$(strFileName, Len(strFileName) - Len(sPath)) '���ļ����������
    
    If left$(tmpStr, 1) = Chr$(0) Then
        'ѡ���˶���ļ�(����Ϊ��һ���ַ�Ϊ�ո�)
        For i = 1 To Len(tmpStr)
            If Mid$(tmpStr, i, 1) = Chr$(0) Then
                iCount = iCount + 1
                ReDim Preserve sFile(iCount)
            Else
                sFile(iCount) = sFile(iCount) & Mid$(tmpStr, i, 1)
            End If
        Next i
    Else
        'ֻѡ����һ���ļ�(ע�⣺��Ŀ¼�µ��ļ�����ȥ·����û��"\"��
        iCount = 1
        ReDim Preserve sFile(iCount)
        If left$(tmpStr, 1) = "\" Then tmpStr = Right$(tmpStr, Len(tmpStr) - 1)
        sFile(iCount) = tmpStr
    End If
    
    GetDlgSelectFileInfo.iCount = iCount
    ReDim GetDlgSelectFileInfo.sFile(iCount)
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    GetDlgSelectFileInfo.sPath = sPath
    For i = 1 To iCount
        GetDlgSelectFileInfo.sFile(i) = sFile(i)
    Next i
    Exit Function
errHandle:
    MsgBox "GetDlgSelectFileInfo����ִ�д���", vbExclamation, gstrSysName
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'���ܣ���������Ŀ¼
'������ strDir��������Ŀ¼
'���أ���
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function funcCeateAViewer(intSeriesIndex As Integer, thisForm As frmViewer) As Integer
'------------------------------------------------
'���ܣ��ڴ�����װ��һ��Viewer�͹���������ʼ����Viewer����Ӧ��MSF����,������ZLSeriesInfos�и�����Ϣ��ZLShowSeriesInfos
'      ��ͼ����ʾ�������������У�����ʾͼ����ԭ������ı�ע��
'      ��ͼ������Ϊ��ʱ���������ʼ�������е�MSF������װ��Viewer�͹�������
'������intSeriesIndex--��Ҫװ�ص�ͼ���������е�����,���Ϊ0����װ���κ�ͼ��
'      thisForm--��ʾͼ��Ĵ���
'���أ������ɹ���Viewer��Index
'ʱ�䣺2009-7
'------------------------------------------------
    Dim i As Integer
    Dim oneSeriesInfo As clsSeriesInfo
    Dim oneImageInfo As clsImageInfo
    Dim intViewerIndex As Integer
    Dim intCurrentIndex As Integer
    Dim intImagesCount As Integer
    
    'intSeriesIndex=0��ʾ���Viewer�е�ͼ����ʱ����֪����ʲôͼ���
    If intSeriesIndex = 0 Then
        If ZLSeriesInfos.Count = 0 Then Exit Function
        intCurrentIndex = 1
    Else
        intCurrentIndex = intSeriesIndex
    End If
    If ZLSeriesInfos.Count < intCurrentIndex Then Exit Function
    
    funcCeateAViewer = 0
    On Error GoTo err
    
    '��ʼ��MSFViewer��������е�����
    With thisForm.MSFViewer
        .Rows = .Rows + 1
        intViewerIndex = .Rows - 1
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        .TextMatrix(intViewerIndex, 0) = ZLSeriesInfos(intCurrentIndex).lngSource  'ͼ����Դ
        .TextMatrix(intViewerIndex, 1) = True                                     '�Ƿ���ͼ
        .TextMatrix(intViewerIndex, 2) = ZLSeriesInfos(intCurrentIndex).StudyUID   '���UID
        .TextMatrix(intViewerIndex, 3) = 1        '��ǰѡ���ͼ���
        .TextMatrix(intViewerIndex, 4) = 1        '��ǰѡ���ͼ���ڵڼ�֡
        .TextMatrix(intViewerIndex, 5) = 0        '�����к�����ʾͼ����Ŀ(�������ڵ�ͼ�Ͷ�ͼ��ʾ�л���)
        .TextMatrix(intViewerIndex, 6) = 0        '������������ʾͼ����Ŀ(�������ڵ�ͼ�Ͷ�ͼ��ʾ�л���)
        .TextMatrix(intViewerIndex, 7) = 1        '�����е�ǰ��ʾ��һ��ͼ�����(�������ڵ�ͼ�Ͷ�ͼ��ʾ�л���)
        .TextMatrix(intViewerIndex, 8) = 1        '�����е�ǰ��ʾѡ��ͼ�����(�������ڵ�ͼ�Ͷ�ͼ��ʾ�л���)
        .TextMatrix(intViewerIndex, 9) = True     '��������ͼ���Ƿ��Զ�ͬ��
        .TextMatrix(intViewerIndex, 15) = 0       '��¼��ǰ�����Ƿ�ѡ�������Զ����ֹ�����ͬ��
    End With
    
    'װ��Viewer����������ͼ��
    With thisForm
        'װ��Viewer�͹�����
        load .Viewer(intViewerIndex)
        load .VScro(intViewerIndex)
        .Viewer(intViewerIndex).UseScrollBars = False
        .Viewer(intViewerIndex).Visible = False
        .Viewer(intViewerIndex).Tag = intCurrentIndex
        .Viewer(intViewerIndex).CellSpacing = lngCellSpacing
        .Viewer(intViewerIndex).BackColour = lngViewerBackColor
        .VScro(intViewerIndex).Visible = False
         
        'װ��ZLShowSeriesInfos�ṹ
        If ZLShowSeriesInfos.Count = intViewerIndex - 1 Then
        
            Set oneSeriesInfo = funGetNewSeriesInfo
            Call funCopySeriesInfo(ZLSeriesInfos(intCurrentIndex), oneSeriesInfo)
            
            ZLShowSeriesInfos.Add oneSeriesInfo
            'װ��ZLShowSeriesInfos�ṹ�е�ͼ��,���intSeriesIndex =0 ��ֻװ�ص�һ��ͼ
            If intSeriesIndex = 0 Then
                intImagesCount = 1
            Else
                intImagesCount = ZLSeriesInfos(intCurrentIndex).ImageInfos.Count
            End If
            
            For i = 1 To intImagesCount
                Set oneImageInfo = funGetNewImageInfo
                Call funCopyImageInfo(ZLSeriesInfos(intCurrentIndex).ImageInfos(i), oneImageInfo)
                
                ZLShowSeriesInfos(intViewerIndex).ImageInfos.Add oneImageInfo
            Next i
        End If
        '�趨ͼ�񲼾�
        Call subSetImageLayout(.Viewer(intViewerIndex), ZLSeriesInfos(intCurrentIndex).strModality, ZLSeriesInfos(intCurrentIndex).ImageInfos.Count)
        'ͼ������,����ͼ������ݿ��ж�ȡ���ǰ���ͼ�������ģ�����Ҫ����������ʾ���е����򷽷�
        Call subSortImages(thisForm, intViewerIndex, funGetImageSort(ZLSeriesInfos(intCurrentIndex).strModality))
        'װ��ͼ��
        Call subShowALLImage(thisForm, .Viewer(intViewerIndex), 1, False)
        
    End With
    funcCeateAViewer = intViewerIndex
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub subSetImageLayout(thisViewer As DicomViewer, strModality As String, intImageCount As Integer)
'------------------------------------------------
'���ܣ�����Ӱ������ͼ���������Ų�ͼ�񲼾�
'������ thisViewer--����ͼ�񲼾����ŵ�����
'       strModality--����ͼ�񲼾����ŵ�Ӱ�����
'       intImageCount--ͼ��������
'���أ��ޣ�ֱ������ָ�����е�ͼ�񲼾�
'���õ��ⲿ������G_INT_MAX_IMG_COL��G_INT_MAX_IMG_ROW
'ʱ�䣺2009-7
'------------------------------------------------
    Dim i As Integer
    Dim blnSetImageLayout As Boolean
    Dim intRows As Integer
    Dim intCols As Integer
    
    For i = 1 To UBound(aPresetLayout)
        If UCase(aPresetLayout(i).strModality) = UCase(strModality) Then
            If aPresetLayout(i).bImageAutoFormat Then
                '���Զ����֣������ͼ������������󲼾���������ͼ�񲼾�
                '���Ǵ�ʱthisViewer��û�аڷŵ������У����Ŀ�Ⱥ͸߶���ʵ��û��ʵ������ġ�
                ResizeRegion intImageCount, thisViewer.width, thisViewer.height, intRows, intCols, G_INT_MAX_IMG_ROW, G_INT_MAX_IMG_COL
                thisViewer.MultiColumns = intCols
                thisViewer.MultiRows = intRows
            Else
                thisViewer.MultiColumns = aPresetLayout(i).lngImageColumns
                thisViewer.MultiRows = aPresetLayout(i).lngImageRows
            End If
            blnSetImageLayout = True
        End If
    Next i
    
    '���û�������û������ͼ������������������Ĭ��ֵ1*1
    If blnSetImageLayout = False Then
        thisViewer.MultiColumns = 1
        thisViewer.MultiRows = 1
    End If
End Sub

Public Function funGetImageSort(strModality As String) As Long
'------------------------------------------------
'���ܣ�����Ӱ��������ͼ������ʽ
'����:
'       strModality--����ͼ�������Ӱ�����
'���أ�ͼ������ʽ
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
    
    funGetImageSort = 0
    
    For i = 1 To UBound(aPresetLayout)
        If UCase(aPresetLayout(i).strModality) = UCase(strModality) Then
            funGetImageSort = aPresetLayout(i).lngImageSort - 1
        End If
    Next i
    Exit Function
err:
    '������
End Function

Public Sub subShowALLImage(thisForm As frmViewer, thisViewer As DicomViewer, intImageIndex As Integer, blnFast As Boolean)
'------------------------------------------------
'���ܣ� ��ʾ���пɼ���ͼ��intImageIndexָ�����Ͻǵ�ͼ��
'       ����ͼ�����*�в��֣�ȷ��������ͼ���Ƿ���ʾ��
'       �����ͼ���Ѿ�����Viewer����ֱ����ʾ��
'       ���ͼ����Viewer�У����ͼ�����Viewer,����ͬʱ��ʾ��ͼ��Ҳһ����
'       ����������ڴ������к��޸Ĳ��֡�ֱ���϶���������ʾĳ��ͼ��
'������ thisForm -- ��Ƭ����
'       thisViewer--����ͼ�񲼾����ŵ�����
'       intImageIndex--ͼ�����ڵ�ͼ������
'       blnFast     -- ������ʾͼ������subDispframe��subDisplayPatientInfo��thisViewer.Refresh
'���أ��ޣ�ֱ�Ӱ�ͼ����벢��ʾ����
'ʱ�䣺2009-7
'------------------------------------------------
    Dim iFoundImageIndex As Integer
    Dim strSaveDir As String
    Dim cFTP As clsFtp
    Dim iCurrImageIndex As Integer
    Dim intImagesCount As Integer
    Dim blnExit As Boolean
    Dim intViewerIndex As Integer
    Dim intAddImageCount As Integer
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo err
    
    '�����жϵ�ǰͼ��Ĳ��֣��ܹ���Ҫ��ʾ����ͼ��Ȼ����ѭ��װ����Щͼ��
    intViewerIndex = thisViewer.Index
    intImagesCount = ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
    
    iCurrImageIndex = intImageIndex
    If iCurrImageIndex > intImagesCount - thisViewer.MultiColumns * thisViewer.MultiRows + 1 Then
        iCurrImageIndex = intImagesCount - thisViewer.MultiColumns * thisViewer.MultiRows + 1
    End If
    If iCurrImageIndex <= 0 Then iCurrImageIndex = 1
    
    '����ͼ�����ʾλ��
    '��ΪZLShowSeriesInfos������ͼ��������ģ���������ӵ�ͼ������Viwer�е�λ��Ӧ�þ���iCurrImageIndex
    '���ǿ���ǰ���ͼ��û����ʾ����û�м��أ���������ӵ�ͼ����Viewer�е�λ�ã���intImageIndex��ǰ��
    'Viwer��ÿ��ͼ���Tag�����ͼ���ImageIndex������ж�ͼ���Tag >iCurrImageIndex ��
    iFoundImageIndex = 0
    For i = IIf(thisViewer.Images.Count > iCurrImageIndex, iCurrImageIndex, thisViewer.Images.Count) To 1 Step -1
        If thisViewer.Images(i).Tag = iCurrImageIndex Then
            iFoundImageIndex = i
            Exit For
        ElseIf thisViewer.Images(i).Tag < iCurrImageIndex Then
            iFoundImageIndex = i + 1
            Exit For
        End If
    Next i
    If iFoundImageIndex = 0 Then iFoundImageIndex = 1   '�����ӵ�ͼ���λ���� iFoundImageIndex
    If iFoundImageIndex > iCurrImageIndex Then iFoundImageIndex = iCurrImageIndex
    
    '����FTP
'    Set cFTP = New clsFtp
'    cFTP.FuncFtpConnect ZLSeriesInfos(intSeriesIndex).strHostIP, ZLSeriesInfos(intSeriesIndex).strFTPUser, _
'            ZLSeriesInfos(intSeriesIndex).strFTPPasw
'    cFTP.FuncChangeDir ZLSeriesInfos(intSeriesIndex).strFTPDir & Replace(ZLSeriesInfos(intSeriesIndex).strSaveDir, "\", "/")
        

    For i = 1 To thisViewer.MultiRows
        For j = 1 To thisViewer.MultiColumns

'           �����ж�ͼ���Ƿ��Ѿ����أ�����Ѿ����أ����ҵ����ͼ����ʾ���������û�м��أ�����ظ�ͼ��
            If ZLShowSeriesInfos(intViewerIndex).ImageInfos(iCurrImageIndex).blnDisplayed = False Then
                'û�м��أ����ҵ�λ�ò�����ͼ��
                intAddImageCount = funcAddAImage(thisViewer, iCurrImageIndex, iFoundImageIndex, cFTP)
                If intAddImageCount > 0 Then
                    '��Ϊ��֡ͼ����װ�غ󣬻�ı�ZLShowSeriesInfos��ͼ�������Ķ��٣������Ҫ���¸�ֵ
                    intImagesCount = ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
                    '�������ù����������ͼ������
                    thisForm.blnVscroInvoked = True
                    thisForm.VScro(intViewerIndex).Max = intImagesCount
                    thisForm.blnVscroInvoked = False
                End If
            End If

'           �Ѿ�����,�����ǵ�ǰ���ֵĵ�һ��ͼ�����ҵ����ͼ����ʾ����ǰ���ֵ�����ͼ����Ҫ����
            If i = 1 And j = 1 And iFoundImageIndex <= thisViewer.Images.Count Then
                thisViewer.CurrentIndex = iFoundImageIndex
            End If

            iFoundImageIndex = iFoundImageIndex + 1
            iCurrImageIndex = iCurrImageIndex + 1
            '���ͼ�����������ͼ��������������˳�ѭ��
            If iFoundImageIndex > intImagesCount Then
                blnExit = True
                Exit For
            End If
        Next j
        '����ڲ�ѭ���Ѿ��˳��������ѭ��Ҳһ���˳�
        If blnExit = True Then
            Exit For
        End If
    Next i
    
    '��ʾ��������ͼ���еĲ�����Ϣ
    Call subDisplayPatientInfo(thisViewer)
        
    '����ǿ�����ʾ���򲻴������¹���
    If blnFast = False Then
        'ͼ����ʾ������Viewer�еı�ע��ͼ������½ǵ�ѡ���ǵ�
        Call subDispframe(thisForm, thisViewer)
    
        thisViewer.Refresh
    End If
    
    If Not cFTP Is Nothing Then cFTP.FuncFtpDisConnect
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    If Not cFTP Is Nothing Then cFTP.FuncFtpDisConnect
    Call SaveErrLog
End Sub

Public Function funcAddAImageA(thisViewer As DicomViewer, ByVal intImageIndex As Integer)
'------------------------------------------------
'���ܣ� ��ָ�����к�ͼ��������ͼ����ӵ�Viewer�У����Զ�����ͼ��λ�ã���ͼ���ƶ����ʺϵ�λ��
'       ����ͼ���ʱ�򣬲���ͼ��ķ�ʽ�ǣ����Ȳ鿴���ػ���---����ȡ����Ŀ¼---���ͨ��FTP����
'������ thisViewer--����ͼ�񲼾����ŵ�����
'       intImageIndex--ͼ�����ڵ�ͼ������
'���أ���ӵ�ͼ������
'ʱ�䣺2009-7
'------------------------------------------------
    Dim intViewerIndex As Integer
    Dim iFoundImageIndex As Integer
    Dim iCurrImageIndex As Integer
    Dim cFTP As clsFtp
    Dim i As Integer
    
    On Error GoTo err
    
    intViewerIndex = thisViewer.Index
    iCurrImageIndex = intImageIndex
    
    '����ͼ�����ʾλ��
    '��ΪZLShowSeriesInfos������ͼ��������ģ���������ӵ�ͼ������Viwer�е�λ��Ӧ�þ���iCurrImageIndex
    '���ǿ���ǰ���ͼ��û����ʾ����û�м��أ���������ӵ�ͼ����Viewer�е�λ�ã���intImageIndex��ǰ��
    'Viwer��ÿ��ͼ���Tag�����ͼ���ImageIndex������ж�ͼ���Tag >iCurrImageIndex ��
    iFoundImageIndex = 0
    For i = IIf(thisViewer.Images.Count > iCurrImageIndex, iCurrImageIndex, thisViewer.Images.Count) To 1 Step -1
        If thisViewer.Images(i).Tag = iCurrImageIndex Then
            iFoundImageIndex = i
            Exit For
        ElseIf thisViewer.Images(i).Tag < iCurrImageIndex Then
            iFoundImageIndex = i + 1
            Exit For
        End If
    Next i
    If iFoundImageIndex = 0 Then iFoundImageIndex = 1   '�����ӵ�ͼ���λ���� iFoundImageIndex
    If iFoundImageIndex > iCurrImageIndex Then iFoundImageIndex = iCurrImageIndex

    '���ͼ���Ƿ��Ѿ�����ʾ��,���˳�����
    If ZLShowSeriesInfos(intViewerIndex).ImageInfos(iCurrImageIndex).blnDisplayed = True Then
        funcAddAImageA = 0
        Exit Function
    End If
    
    funcAddAImageA = funcAddAImage(thisViewer, iCurrImageIndex, iFoundImageIndex, cFTP)
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function funcAddAImage(thisViewer As DicomViewer, ByVal intImageIndex As Integer, ByVal intCurrentIndex As Integer, cFTP As clsFtp) As Integer
'------------------------------------------------
'���ܣ� ��ָ�����к�ͼ��������ͼ����ӵ�Viewer�У����ƶ����ʺϵ�λ��
'       ����ͼ���ʱ�򣬲���ͼ��ķ�ʽ�ǣ����Ȳ鿴���ػ���---����ȡ����Ŀ¼---���ͨ��FTP����
'������ thisViewer--����ͼ�񲼾����ŵ�����
'       intImageIndex--ͼ�����ڵ�ͼ������
'       intCurrentIndex---ͼ����Ҫ�ڷŵ�λ��
'       cFTP ---FTP���ӣ�Ӧ���Ѿ����Ӻã��������ú�Ŀ¼����
'���أ���ӵ�ͼ������
'ʱ�䣺2009-7
'------------------------------------------------
    Dim NewImg As DicomImage
    Dim img As DicomImage
    Dim intViewerIndex As Integer
    Dim i As Integer
    Dim OldImageInfos As Collection
    Dim OneImageInfos As clsImageInfo
    Dim intImageCount As Integer
    
    On Error GoTo err
    
    intViewerIndex = thisViewer.Index
    Set NewImg = funLoadAImage(intViewerIndex, intImageIndex, 1)
    If NewImg Is Nothing Then
        MsgBox "�����ļ��������鱾���������FTP���ӡ�", vbOKOnly, "����ͼ����ʾ"
        Exit Function
    End If
    
    '����ͼ���Ѿ���ʾ���ı��
    ZLShowSeriesInfos(intViewerIndex).ImageInfos(intImageIndex).blnDisplayed = True
    
    '�Զ�֡�͵�֡ͼ����д���
    If NewImg.FrameCount > 1 Then  '��֡ͼ��
        '������дZLShowSeriesInfos�ṹ
        Set OldImageInfos = New Collection
        
        For intImageCount = 1 To ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
            OldImageInfos.Add ZLShowSeriesInfos(intViewerIndex).ImageInfos(intImageCount)
        Next intImageCount
        
        '�Ѷ�֡ͼ���ͼ����Ϣ��ӵ�������
        Set OneImageInfos = ZLShowSeriesInfos(intViewerIndex).ImageInfos(intImageIndex)
        
        Set ZLShowSeriesInfos(intViewerIndex).ImageInfos = Nothing
        
        For i = 1 To intImageIndex
            ZLShowSeriesInfos(intViewerIndex).ImageInfos.Add OldImageInfos(i)
        Next i
        
        For i = 2 To NewImg.FrameCount
            ZLShowSeriesInfos(intViewerIndex).ImageInfos.Add OneImageInfos
        Next i
        '�Ѻ����ͼ����Ϣ����ؼ�����
        For i = intImageIndex + 1 To OldImageInfos.Count
            ZLShowSeriesInfos(intViewerIndex).ImageInfos.Add OldImageInfos(i)
        Next i
        '���OldImageInfos
        Set OldImageInfos = Nothing
    End If
    
    '����ͼ��
    For i = 1 To NewImg.FrameCount
        thisViewer.Images.Add NewImg
        Set img = thisViewer.Images(thisViewer.Images.Count)
        img.Tag = intImageIndex
        img.Frame = i
    
        Call subInitAImage(img, intViewerIndex, thisViewer)
        
        '��ͼ���ƶ������ʵ�λ��
        If intCurrentIndex <> thisViewer.Images.Count And thisViewer.Images.Count <> 0 Then
            thisViewer.Images.Move thisViewer.Images.Count, intCurrentIndex
        End If
        
        intImageIndex = intImageIndex + 1
        intCurrentIndex = intCurrentIndex + 1
    Next i
    
    funcAddAImage = NewImg.FrameCount
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub subInitAImage(img As DicomImage, intViewerIndex As Integer, thisViewer As DicomViewer)
'------------------------------------------------
'���ܣ� ��ʼ��ͼ��,������ͼ��Ĵ��ڵȽ���ͬ��
'������ img--��Ҫ��ʼ����ͼ��
'       intViewerIndex--ͼ�����ڵ�Viewer������,0��ʾ��ʹ���������
'       thisViewer -- ͼ�񼴽�Ҫ����thisViewer��
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim strRefUID As String
    
    If img Is Nothing Then Exit Sub
    
    On Error GoTo err
    '����������˾��MRͼ�󣺶�λ�߲�����ʾ����,�޸Ĺ���Ϊ����(0020,0052) : Frame of Reference UID�޸ĳɺ�׺Ϊ1
    If Not IsNull(img.Attributes(&H8, &H60).Value) And Not IsNull(img.Attributes(&H8, &H70).Value) _
        And Not IsNull(img.Attributes(&H20, &H52).Value) Then
        
        If img.Attributes(&H8, &H60).Value = "MR" And img.Attributes(&H8, &H70).Value = "NingBo XGY Magnetism Co.,LTD." Then
            strRefUID = img.Attributes(&H20, &H52).Value
            strRefUID = left(strRefUID, InStrRev(strRefUID, ".")) & "1"
            img.Attributes.Add &H20, &H52, strRefUID
        End If
    End If
    
    '�������DRͼ������VOILUT=0����ͼ������������ʾ
    If Not IsNull(img.Attributes(&H8, &H1090).Value) And Not IsNull(img.Attributes(&H8, &H60).Value) Then
        If UCase(img.Attributes(&H8, &H60).Value) = "DX" And _
            (UCase(img.Attributes(&H8, &H1090).Value) = UCase("NavigationSight") Or UCase(img.Attributes(&H8, &H1090).Value) = UCase("NeuVision DR III")) Then
            img.VOILUT = 0
        End If
    End If
    
    'ȡ���Զ���Ӱ,��ΪDicomObjects�ؼ�����Դ����Ӱ��BUG�����ڣ�0028��6100��ʱ�����Զ���ͼ����м�Ӱ��
    '���½�ú��DSAͼ����������ʾ
    '��Ȼ����ͼ���mask=0 ,����ȡ����Ӱ������ÿ��ͼ����ӵ��µ�Dicomimages֮���Զ��ֽ�mask���ó�1�ˣ�
    '�����ڳ������޷��ܺõĿ��ƣ����ֱ��ȥ����0028��6100��������ԡ�
    If Not IsNull(img.Attributes(&H28, &H6100).Value) Then
        img.Attributes.Remove &H28, &H6100
    End If
    
    '����ͼ��ķŴ�ģʽ
    img.MagnificationMode = intMagnificationMode
    
    '����ͼ�����ǿ����
    img.UnsharpLength = 0
    
    'ȡ����ʾ���棬���Լӿ���ʾ�ٶȣ����Ҽ��ٶ��ڴ������
    '������ͼ������CacheDisplayΪFalse�������Ҫ���⴦��������Intera MR
    If Not IsNull(img.Attributes(&H8, &H60).Value) Then
        If UCase(img.Attributes(&H8, &H60).Value) = "PR" Or UCase(img.Attributes(&H8, &H60).Value) = "KO" Or UCase(img.Attributes(&H8, &H60).Value) = "SR" Then
            '����ΪPR�ģ������κδ�����������
        ElseIf UCase(img.Attributes(&H8, &H60).Value) = "MR" And Not IsNull(img.Attributes(&H8, &H16).Value) Then
            If left(img.Attributes(&H8, &H16).Value, Len(img.Attributes(&H8, &H16).Value) - 1) = "1.3.46.670589.11.0.0.12." Or _
                img.Attributes(&H8, &H16).Value = "1.2.840.10008.5.1.4.1.1.66" Then
                  '����ΪMR�ģ������Ե�֪�����Sop Class UID ="1.3.46.670589.11.0.0.12.2"��"1.3.46.670589.11.0.0.12.4" ����Ҳ�����κδ�����������
                  '��������������SOP ClassUID,����ж�ǰ׺��1.3.46.670589.11.0.0.12.xxx��
            Else
                img.CacheDisplay = False
            End If
        Else
            img.CacheDisplay = False
        End If
    Else
        img.CacheDisplay = False
    End If
    
    
    '��ʼ��ͼ���е�ϵͳ��ע
    subInitImageLabels intViewerIndex, 1, img, True, True, True     '��ʼ��ͼ���ע��Ϣ��Ϣ:ϵͳ��ע����λ��Ϣ����ߣ��Ľ���Ϣ
    
    '��ʼ��ͼ����ڸ�ͼ
    If left(img.InstanceUID, 24) = "2.16.840.1.113669.632.3." Then
        subDrawImgShutter img, True
    Else
        subDrawImgShutter img
    End If
    
    
    '��ʾ������ͼ���еı�ע
    subReadLabelFromImg img           ''''����ͼ��ı�ע��Ϣ
    
    If img.Attributes(&H6000, &H10).Exists = True Then
        img.OverlayVisible(0) = Button_miShowOverlay
    End If
    
    '����Ԥ��Ĵ���λ
    If intViewerIndex > 0 And intViewerIndex <= ZLShowSeriesInfos.Count Then
        If ZLShowSeriesInfos(intViewerIndex).lngWinWidth <> 0 And ZLShowSeriesInfos(intViewerIndex).lngWinLevel <> 0 Then
            img.width = ZLShowSeriesInfos(intViewerIndex).lngWinWidth
            img.Level = ZLShowSeriesInfos(intViewerIndex).lngWinLevel
            img.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & img.width & "-L:" & img.Level
            'ʹ��Ĭ�ϵĴ���λ����Ҫ����VOILUT=0������Ч
            img.VOILUT = 0
        End If
    End If
    
    '����һ��Overlay����ʾ,Overlay������һ���ǰ�ɫ�ģ������ð�ͼ���ɫ���ó�1
    If Not IsNull(img.Attributes(&H6000, &H15).Value) Then
        If img.Attributes(&H6000, &H15).Value = 1 Then
            If img.Level = 0 Then img.Level = 1
            img.OverlayVisible(0) = True
            img.OverlayColour(0) = lngLabelColor
        End If
    End If
    
    '����ͼ������ͬ��
    If Button_miImageInPhase = True And Not thisViewer Is Nothing Then
        If thisViewer.Images.Count > 0 Then
            Call subImageInPhase(img, thisViewer.Images(1), IMG_SYN_All)
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subShowForm(thisForm As frmViewer)
'------------------------------------------------
'���ܣ� ����ZLSeriesInfos��ͼ���������ʾ���ڣ���������ʾViewer������������ʾViewer�е�ͼ��
'       ����Ҫ��ʾ��ͼ����ؽ���VIEWER,��Viewer��ʾ�������У���ʾViewer��صĹ�����
'       ����ͼ���ʱ�򣬲���ͼ��ķ�ʽ�ǣ����Ȳ鿴���ػ���---����ȡ����Ŀ¼---���ͨ��FTP����
'������ ��
'���أ��ޣ�ֱ�Ӱ�Viewer��ͼ����벢��ʾ����
'ʱ�䣺2009-7
'------------------------------------------------
    Dim strModality As String
    Dim intSeriesCount As Integer
    Dim blnLoadOver As Boolean
    Dim intCurrentSeries As Integer
    Dim intCurrentViewer  As Integer
    Dim intViewerIndex As Integer
    Dim intSeriesIndex As Integer
    Dim oneImageInfo As clsImageInfo
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    'ͨ������Ƭ����ʽ�򿪺�ͨ�����Աȡ���ʽ�򿪣���ԭ��Viewer�Ĵ����ǲ�һ���ġ�
    
    If ZLSeriesInfos.Count = 0 Then Exit Sub
    
    On Error GoTo err
    
    intSeriesCount = ZLSeriesInfos.Count
    intCurrentSeries = 0
    intViewerIndex = 0
    
    'ͨ������Ƭ����ʽ��ʱViewer������Ϊ1 ��ͨ�����Աȡ���ʽ��ʱ��Viewer����������1
    '�����ж��ǡ���Ƭ����ʽ�򿪻��ǡ��Աȡ���ʽ��
    If thisForm.Viewer.Count = 1 Then   '����Ƭ����ʽ�򿪹�Ƭվ
        'ͨ����һ��ͼ��Ӱ�����ȷ�����в��֣��ڷŷָ���
        strModality = ZLSeriesInfos(1).strModality
        '���ݵ�һ��ͼ���Ӱ����𣬻��ͼ������в���
        Call subSetSeriesLayout(thisForm, strModality, intSeriesCount)
        '�ڷŷָ���
        Call subShowSpliter(thisForm)
    End If
    
    '�������в��֣�����Viewer������������װ��ͼ��
    For i = 1 To thisForm.intCountY
        For j = 1 To thisForm.intCountX
            '���ֵ������������е������������˳�ѭ��
            If (i - 1) * thisForm.intCountX + j > intSeriesCount Then
                blnLoadOver = True
                Exit For
            End If
            intCurrentSeries = intCurrentSeries + 1
            intViewerIndex = intViewerIndex + 1
            '�ж����Viwer�Ƿ���ڣ�����Ѿ����ڣ�������װ�����Viwer�е�ͼ��
            If intViewerIndex >= thisForm.Viewer.Count Then  ''Viewer�����ڣ��򴴽�Viewer
                '�������ڷ�һ��Viewer
                intCurrentViewer = funcCeateAViewer(intCurrentSeries, thisForm)
                
                '�ڷ����Viewer�����ù�����
                Call subPlaceAViewer(thisForm, intCurrentViewer, i, j)
            End If
            
            
        Next j
        If blnLoadOver = True Then
            Exit For
        End If
    Next i
    
    '����Ĭ�ϱ�ѡ������к�ͼ��
    If thisForm.Viewer.Count > 1 Then
        If thisForm.Viewer(1).Images.Count > 0 Then
            Set thisForm.SelectedImage = thisForm.Viewer(1).Images(1)
            thisForm.SelectedImageIndex = 1
        End If
        thisForm.intSelectedSerial = 1
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subSetSeriesLayout(thisForm As frmViewer, strModality As String, intSeriesCount As Integer)
'------------------------------------------------
'���ܣ�����Ӱ�����������������Ų����в���
'������ thisForm--�������в������ŵĴ���
'       strModality--�������в������ŵ�Ӱ�����
'       intSeriesCount--����������
'���أ��ޣ�ֱ������ָ�����еĲ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim i As Integer
    Dim intRows As Integer
    Dim intCols As Integer
    Dim blnSetSeriesLayout As Boolean
    
    For i = 1 To UBound(aPresetLayout)
        If UCase(aPresetLayout(i).strModality) = UCase(strModality) Then
            If aPresetLayout(i).bSeriesAutoFormat = True Then
                '���Զ����֣��������е�����������󲼾�����������ͼ�񲼾�
                ResizeRegion intSeriesCount, thisForm.width, thisForm.height, intRows, intCols, intMaxAreaY, intMaxAreaX
                thisForm.intCountX = intCols
                thisForm.intCountY = intRows
            Else
                thisForm.intCountX = aPresetLayout(i).lngSeriesColumns
                thisForm.intCountY = aPresetLayout(i).lngSeriesRows
            End If
            blnSetSeriesLayout = True
        End If
    Next i
    
    If blnSetSeriesLayout = False Then
        '����Ĭ�ϵ����в���
        thisForm.intCountX = 2
        thisForm.intCountY = 2
    End If
End Sub

Public Sub subPlaceAViewer(thisForm As frmViewer, intViewerIndex As Integer, intRow As Integer, intCol As Integer)
'------------------------------------------------
'���ܣ���Viewer����ָ��λ�ã�����ʾ��������ͼ��ѡ���
'������ thisForm--��Ƭ����
'       intViewerIndex--��Ҫ���õ�Viewer�͹�������Index
'       intRow -- �ڷ�Viewer ����
'       intCol -- �ڷ�Viewer ����
'���أ��ޣ��ڷ�Viewer���ҰڷŹ�����
'ʱ�䣺2009-7
'------------------------------------------------
    Dim lngTop As Long
    Dim lngLeft As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngOldWidth As Long
    Dim lngOldHeight As Long
    
    If intRow = 0 Or intCol = 0 Then Exit Sub
    On Error GoTo err
    
    '����intRow��intCol �������Viewer��Ҫ�ڷŵ�λ��
    With thisForm
        If intCol = 1 Then  ''''���㵱ǰviewer�ĺ���λ��
            lngLeft = 0
            lngWidth = .PicX(intCol).left
        Else
            lngLeft = .PicX(intCol - 1).left + intSpaceSize
            If intCol = intMaxAreaX Then
                lngWidth = .picViewer.ScaleWidth - .PicX(intCol - 1).left - intSpaceSize
            Else
                If .PicX(intCol).left - .PicX(intCol - 1).left - intSpaceSize < 0 Then
                    lngWidth = 0
                Else
                    lngWidth = .PicX(intCol).left - .PicX(intCol - 1).left - intSpaceSize
                End If
            End If
        End If
        '''''''''''''''''''''''''''''''''''''''''''���㵱ǰviewer������λ��
        If intRow = 1 Then
            lngTop = 0
            lngHeight = .PicY(intRow).top
        Else
            lngTop = .PicY(intRow - 1).top + intSpaceSize
            If intRow = intMaxAreaY Then
                lngHeight = .picViewer.ScaleHeight - .PicY(intRow - 1).top - intSpaceSize
            Else
                If .PicY(intRow).top - .PicY(intRow - 1).top - intSpaceSize < 0 Then
                    lngHeight = 0
                Else
                    lngHeight = .PicY(intRow).top - .PicY(intRow - 1).top - intSpaceSize
                End If
            End If
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If lngHeight < 0 Then lngHeight = 0
    End With
    
    '�ڷŲ���ʾViewer
    lngOldWidth = thisForm.Viewer(intViewerIndex).width / thisForm.Viewer(intViewerIndex).MultiColumns
    lngOldHeight = thisForm.Viewer(intViewerIndex).height / thisForm.Viewer(intViewerIndex).MultiRows
    thisForm.Viewer(intViewerIndex).Move lngLeft, lngTop, Abs(lngWidth), Abs(lngHeight)
    thisForm.Viewer(intViewerIndex).Visible = (lngWidth <> 0)
    ZLShowSeriesInfos(intViewerIndex).intRow = intRow
    ZLShowSeriesInfos(intViewerIndex).intCol = intCol
    
    '���ͼ����StretchToFit�������Viewer��ͼ���λ��
    If thisForm.Viewer(intViewerIndex).Images.Count > 0 Then
        Call subScaleViewer(thisForm.Viewer(intViewerIndex), thisForm.Viewer(intViewerIndex).Images(1), lngOldWidth, lngOldHeight)
    End If
    
    '�жϹ������Ƿ���Ҫ��ʾ�������Ҫ����ʾ�����������ù����������ֵ����Сֵ��LarghChange��
    subDisplayScrollBar intViewerIndex, thisForm, True
    
    '�����ѡ���������е�״̬����Ŀǰ��������Ϊ��ǰ���У���subDispframeʹ��
    If thisForm.isSelectAllSerial Then thisForm.intSelectedSerial = intViewerIndex
    
    'ͼ����ʾ������Viewer�еı�ע��ͼ������½ǵ�ѡ���ǵ�
    Call subDispframe(thisForm, thisForm.Viewer(intViewerIndex))
    
    '�Զ�����ͼ���С���ж��Ƿ���ʾ�����Ľ���Ϣ,��ʾ��������ͼ���еĲ�����Ϣ
    Call subDisplayPatientInfo(thisForm.Viewer(intViewerIndex))
Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subShowSpliter(thisForm As frmViewer)
'------------------------------------------------
'���ܣ��������в��֣�������ʾ�ָ���
'������ thisForm--��Ƭ����
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim intRows As Integer
    Dim intCols As Integer
    Dim lngAreaWidth As Long
    Dim lngAreaHeight As Long
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    On Error GoTo err
    '����ǹ۲�ģʽ������ʾ1*1�Ĳ���
    If Button_miLookOrBrowse = True Then '�۲�ģʽ
        intRows = 1
        intCols = 1
    Else    '��������ģʽ������intCountX��intCountY�Ķ�����ʾ����
        intRows = thisForm.intCountY
        intCols = thisForm.intCountX
    End If
    
    '���㲢�Ұڷŷָ���
    With thisForm
        If intCols = intMaxAreaX Then   ''���������ÿ�ȣ��������Ѿ�ȷ����intCols�������intMaxAreaX
            lngAreaWidth = .picViewer.ScaleWidth - intSpaceSize * (intMaxAreaX - 1)
        Else
            lngAreaWidth = .picViewer.ScaleWidth - intSpaceSize * intCols
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If intRows = intMaxAreaY Then   ''���������ÿ��
            lngAreaHeight = .picViewer.ScaleHeight - intSpaceSize * (intMaxAreaY - 1)
        Else
            lngAreaHeight = .picViewer.ScaleHeight - intSpaceSize * intRows
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To intMaxAreaX - 1  ''�������еķָ��߹�λ
            .PicX(i).left = .picViewer.ScaleWidth - intSpaceSize
            .PicX(i).Tag = ""
            .PicX(i).height = .picViewer.ScaleHeight
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To intMaxAreaY - 1 ''���������еķָ��߹�λ
            .PicY(i).top = .picViewer.ScaleHeight - intSpaceSize
            .PicY(i).Tag = ""
            .PicY(i).width = .picViewer.ScaleWidth
        Next
        
        '�����϶����Ŀ�Ⱥ͸߶�
        .PicXX.height = .picViewer.ScaleHeight
        .PicYY.width = .picViewer.ScaleWidth
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To intCols - 1  ''������Ҫ��ʾ����λ�ü���
            .PicX(i).left = lngAreaWidth / intCols * i + intSpaceSize * (i - 1)
            .PicX(i).Tag = .PicX(i).left
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To intRows - 1  ''������Ҫ��ʾ����λ�ü���
            .PicY(i).top = lngAreaHeight / intRows * i + intSpaceSize * (i - 1)
            .PicY(i).Tag = .PicY(i).top
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To intMaxAreaX - 1  ''�Ų�˫��ָ����λ��
            For j = 1 To intMaxAreaY - 1
                k = (j - 1) * (intMaxAreaX - 1) + i
                .PicXY(k).top = .PicY(j).top
                .PicXY(k).left = .PicX(i).left
            Next
        Next
    End With
Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetImageAttribute(objAttr As DicomAttributes, ByVal AttrName As String) As String
'-----------------------------------------------------------------------------
'����:��ȡDICOM���Լ��е�ָ������ֵ,����VM�ж�ֵ��ά�ȣ�ʹ�á�\���Ѹ���ά�����ӳ�һ����
'-----------------------------------------------------------------------------
    Dim AttrTag() As String
    Dim i As Integer
    
    GetImageAttribute = ""
    AttrTag = Split(AttrName, ":")
    If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Exists Then
        If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).VM = 1 Then
            GetImageAttribute = Nvl(objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).ValueByIndex(1))
        Else
            For i = 1 To objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).VM
                GetImageAttribute = GetImageAttribute & "\" & objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).ValueByIndex(i)
            Next i
        End If
    End If
End Function

Public Sub subDisplayScrollBar(intViewerIndex As Integer, thisForm As frmViewer, blnResizeViewer As Boolean)
'------------------------------------------------
'���ܣ�����ͼ����������ж���ʾ�������ع�����
'������ intViewerIndex--������������
'       thisForm --  ��Ƭ����
'       blnResizeViewer ---TrueViewer�Ŀ�����µ������ˡ�
'���أ��� ��ֱ����ʾ�����ع�������
'ʱ�䣺2009-7
'------------------------------------------------
    Dim thisViewer As DicomViewer
    Dim thisVScro As VScrollBar
    Dim lngImageCount As Long
    
    On Error GoTo err
    '�ҵ�Viewer����Ӧ����������
    Set thisViewer = thisForm.Viewer(intViewerIndex)
    Set thisVScro = thisForm.VScro(intViewerIndex)

    If thisViewer.Images.Count = 0 Then Exit Sub
    
    lngImageCount = ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
    
    '�ж�Viewer����Ӧ��������ͼ��������Ƿ������ʾ������������ʾ������
    If lngImageCount > thisViewer.MultiColumns * thisViewer.MultiRows Then  '��ʾ������
        '�жϹ���������ʾ״̬�Ƿ����仯,�ڷŹ�����������ʾ������
        If blnResizeViewer = True Or thisVScro.Visible = False Then    '����Viewer�Ŀ�ȣ���ռ��Viewer�Ĳ��ֿռ�
            thisVScro.Move thisViewer.left + thisViewer.width - thisVScro.width, thisViewer.top, thisVScro.width, thisViewer.height
            thisViewer.width = Abs(thisViewer.width - thisVScro.width)
        Else
            thisVScro.Move thisViewer.left + thisViewer.width, thisViewer.top, thisVScro.width, thisViewer.height
        End If
        thisVScro.Visible = thisViewer.Visible
        thisVScro.ZOrder
        thisVScro.Refresh
        '���ù������������Сֵ
        thisVScro.Min = 1
        thisVScro.Max = lngImageCount - thisViewer.MultiColumns * thisViewer.MultiRows + 1
        If thisVScro.Max < 1 Then thisVScro.Max = 1
        thisVScro.LargeChange = thisViewer.MultiColumns * thisViewer.MultiRows
        If thisViewer.CurrentIndex > thisVScro.Max Then
            thisVScro.Value = thisVScro.Max
            thisViewer.CurrentIndex = thisVScro.Max
        Else
            thisVScro.Value = thisViewer.CurrentImage.Tag
        End If
    Else    'ͼ�����ڿ���ʾ�������������ع�����
        If blnResizeViewer = False And thisVScro.Visible = True Then    '����Viewer�Ŀ��
            thisViewer.width = thisViewer.width + thisVScro.width
        End If
        thisVScro.Visible = False
        'ֻ�е�ǰû��ѡ��ͼ��ʱ�����õ�ǰͼ�񣬷�����MPR�Ȳ������Ӱ��
        If thisForm.SelectedImage Is Nothing Then
            thisForm.SelectedImageIndex = thisViewer.CurrentIndex
            Set thisForm.SelectedImage = thisViewer.CurrentImage
            thisForm.intSelectedSerial = thisViewer.Index
            thisForm.MSFViewer.TextMatrix(thisViewer.Index, 3) = thisForm.SelectedImageIndex
        End If
    End If
    
    '�������Ĵ���λ�˵�
    Call subSetWidthLevelF(thisForm.SelectedImage, thisForm)
    '��������ͼ���˾��˵�
    Call subSetFilterF(thisForm.SelectedImage, thisForm)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subShowMiniImages(thisForm As frmViewer)
'------------------------------------------------
'���ܣ���ʾ��ر���������ͼ
'������ intViewerIndex--������������
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim i As Integer
    Dim imgs As New DicomImages
    Dim img As DicomImage
    
    On Error GoTo err
    
    '���ݹ����������ж��Ƿ���ʾ����ͼ
    If Button_miShowMiniSeries = True Then      '��ʾ����ͼ
        For i = 1 To ZLSeriesInfos.Count
            '����һ��ͼ��
            Set img = funLoadAImage(i, 1, 0)
            If Not img Is Nothing Then
                imgs.Add img
            End If
        Next i
        If blnDockMiniImage = True Then
            frmMiniSeries.ShowMe imgs, thisForm, thisForm.dkpMain
        Else
            frmMiniSeries.ShowMe imgs, thisForm
        End If
    Else        '��������ͼ
        If blnDockMiniImage = True Then
            frmMiniSeries.CloseMe thisForm.dkpMain
        Else
            frmMiniSeries.CloseMe
        End If
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function funLoadAImage(intSeriesIndex As Integer, intImageIndex As Integer, intLoadType As Integer) As DicomImage
'------------------------------------------------
'���ܣ� ��ָ�����к�ͼ��������ͼ����ӵ�funLoadAImage��
'       ����ͼ���ʱ�򣬲���ͼ��ķ�ʽ�ǣ����Ȳ鿴���ػ���---����ȡ����Ŀ¼---���ͨ��FTP����
'������ intViewerIndex--ͼ�����ڵ����е�����������������ȡ���У���intLoadType���
'       intImageIndex--ͼ�����ڵ�ͼ�������
'       intLoadType -- װ��ģʽ��0--��ZLSeriesInfosװ�أ�1 -- ��ZLShowSeriesInfosװ��
'���أ�����ͼ��ɹ����򷵻�ͼ�񣬷��򷵻�NOTHING
'ʱ�䣺2009-7
'------------------------------------------------
    Dim thisImage As DicomImage
    Dim adata As DicomDataSet
    Dim attr As DicomAttribute
    Dim strInstanceUID As String
    Dim strSeriesUID As String
    Dim intPRISeriesIndex As Integer
    Dim intPRIImageIndex As Integer
    Dim i As Integer
    Dim dssSub1 As DicomDataSets
    Dim dssSub2 As DicomDataSets
    Dim dssSub7060 As DicomDataSets
    Dim dssSub701 As DicomDataSets
    Dim dssSub283110 As DicomDataSets
    Dim dssSub705A As DicomDataSets
    Dim dsSub70601 As DicomDataSet
    Dim dssub70602 As DicomDataSet
    
    Dim OriginImage As DicomImage
    
    
    On Error GoTo err
    
    Set thisImage = funLoadOneImage(intSeriesIndex, intImageIndex, intLoadType)
    Set funLoadAImage = thisImage
    If funLoadAImage Is Nothing Then
        Debug.Print "ͼ���ȡ����"
        Exit Function
    End If
    
    '����PRͼ����Ϣ
    If UCase(Nvl(thisImage.Attributes(&H8, &H60).Value, "OT")) = "PR" Then
        '����PRͼ���Ӧ��ԭʼͼ��
        '��ȡPRͼ���е�����UID��ͼ��UID
        If thisImage.Attributes(&H8, &H1115).Exists = True Then
            Set dssSub1 = thisImage.Attributes(&H8, &H1115).Value
            If dssSub1(1).Attributes(&H8, &H1140).Exists Then
                Set dssSub2 = dssSub1(1).Attributes(&H8, &H1140).Value
                    If dssSub2(1).Attributes(&H8, &H1155).Exists = True Then
                        strInstanceUID = dssSub2(1).Attributes(&H8, &H1155).Value
                    End If
            End If
            If dssSub1(1).Attributes(&H20, &HE).Exists = True Then
                strSeriesUID = dssSub1(1).Attributes(&H20, &HE).Value
            End If
        End If

        '�������UID����ͼ��UIDΪ�գ����˳�
        If strSeriesUID = "" Or strInstanceUID = "" Then
            Exit Function
        End If

        '����PR��Ӧ��ԭʼͼ
        For i = 1 To ZLSeriesInfos.Count
            If ZLSeriesInfos(i).SeriesUID = strSeriesUID Then
                intPRISeriesIndex = i
                Exit For
            End If
        Next i
        If intPRISeriesIndex > ZLSeriesInfos.Count Or intPRISeriesIndex <= 0 Then
            Exit Function
        End If

        For i = 1 To ZLSeriesInfos(intPRISeriesIndex).ImageInfos.Count
            If ZLSeriesInfos(intPRISeriesIndex).ImageInfos(i).InstanceUID = strInstanceUID Then
                intPRIImageIndex = i
                Exit For
            End If
        Next i
        If intPRIImageIndex > ZLSeriesInfos(intPRISeriesIndex).ImageInfos.Count Then
            Exit Function
        End If

        '��ȡԭʼͼ�����Ϣ
        '����ԭʼͼ��
        Set OriginImage = funLoadOneImage(intPRISeriesIndex, intPRIImageIndex, 0)

        '��ȡPRͼ�����Ϣ
        Set adata = New DicomDataSet
        
       '������ȡPR������Ϣ
        On Error Resume Next
        For Each attr In thisImage.Attributes
            adata.Attributes.Add attr.Group, attr.Element, attr.Value
        Next
        
        '���ƹ�˾������PRͼ��û����ȷ�ģ�70,2�����ƣ��������70,60����Ҳû�����ݣ�
        '���±�ע��ʾ�������������Ҫ���⴦��
'       '��һ��Text
        Set dssSub1 = adata.Attributes(&H70, &H1).Value
        If IsNull(dssSub1(1).Attributes(&H70, &H2).Value) Then
            dssSub1(1).Attributes.Add &H70, &H2, "LAYER1"
        End If

        '���� Graphic Layer Sequence
        If IsNull(adata.Attributes(&H70, &H60).Value) Then
            Set dssSub7060 = New DicomDataSets
            Set dsSub70601 = New DicomDataSet
            dsSub70601.Attributes.Add &H70, &H2, "LAYER1"
            dsSub70601.Attributes.Add &H70, &H62, 1
            dsSub70601.Attributes.Add &H70, &H68, "layer1"
            dssSub7060.Add dsSub70601
            adata.Attributes.Add &H70, &H60, dssSub7060
        End If
        
        Set OriginImage.PresentationState = adata

        '�޸�һЩͼ���еı�Ҫ��Ϣ������ͼ���InstanceUID��SeriesUID��StudyUID�������޸ģ�����PRͼ����ʾ������

        OriginImage.Name = thisImage.Name
        OriginImage.AccessionNumber = thisImage.AccessionNumber
        OriginImage.PatientID = thisImage.PatientID
        OriginImage.SeriesDescription = thisImage.SeriesDescription
        OriginImage.StudyDescription = thisImage.StudyDescription
        If thisImage.Attributes(&H20, &H10).Exists Then 'study id
            OriginImage.Attributes.Add &H20, &H10, thisImage.Attributes(&H20, &H10).Value
        End If
        If thisImage.Attributes(&H20, &H11).Exists Then 'series number
            OriginImage.Attributes.Add &H20, &H11, thisImage.Attributes(&H20, &H11).Value
        End If
        If thisImage.Attributes(&H20, &H13).Exists Then 'image number
            OriginImage.Attributes.Add &H20, &H13, thisImage.Attributes(&H20, &H13).Value
        End If
        If thisImage.Attributes(&H8, &H60).Exists Then  'modality
            OriginImage.Attributes.Add &H8, &H60, thisImage.Attributes(&H8, &H60).Value
        End If
        If thisImage.Attributes(&H8, &H20).Exists Then  'study date
            OriginImage.Attributes.Add &H8, &H20, thisImage.Attributes(&H8, &H20).Value
        End If
        If thisImage.Attributes(&H8, &H30).Exists Then  'study time
            OriginImage.Attributes.Add &H8, &H30, thisImage.Attributes(&H8, &H30).Value
        End If
        If thisImage.Attributes(&H28, &H3110).Exists Then   '����λ
            Set dssSub283110 = thisImage.Attributes(&H28, &H3110).Value
            If dssSub283110(1).Attributes(&H28, &H1050).Exists Then
                If Not IsNull(dssSub283110(1).Attributes(&H28, &H1050).Value) Then
                    OriginImage.Attributes.Add &H28, &H1050, dssSub283110(1).Attributes(&H28, &H1050).Value
                    OriginImage.Level = dssSub283110(1).Attributes(&H28, &H1050).ValueByIndex(1)
                End If
            End If
            If dssSub283110(1).Attributes(&H28, &H1051).Exists Then
                If Not IsNull(dssSub283110(1).Attributes(&H28, &H1051).Value) Then
                    OriginImage.Attributes.Add &H28, &H1051, dssSub283110(1).Attributes(&H28, &H1051).Value
                    OriginImage.width = dssSub283110(1).Attributes(&H28, &H1051).ValueByIndex(1)
                End If
            End If
        End If

        Set funLoadAImage = OriginImage

    End If
     
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function funLoadOneImage(intSeriesIndex As Integer, intImageIndex As Integer, intLoadType As Integer) As DicomImage
'------------------------------------------------
'���ܣ� ��ָ�����к�ͼ��������ͼ����ӵ�funLoadOneImage��
'       ����ͼ���ʱ�򣬲���ͼ��ķ�ʽ�ǣ����Ȳ鿴���ػ���---����ȡ����Ŀ¼---���ͨ��FTP����
'������ intViewerIndex--ͼ�����ڵ����е�����������������ȡ���У���intLoadType���
'       intImageIndex--ͼ�����ڵ�ͼ�������
'       intLoadType -- װ��ģʽ��0--��ZLSeriesInfosװ�أ�1 -- ��ZLShowSeriesInfosװ��
'���أ�����ͼ��ɹ����򷵻�ͼ�񣬷��򷵻�NOTHING
'ʱ�䣺2009-7
'------------------------------------------------
    Dim imgs As New DicomImages
    Dim strSaveDir As String
    Dim lngSource As Long
    Dim strShareDir As String
    Dim strHostIP As String
    Dim strImageName As String
    Dim strLocalImage As String
    Dim strFTPUser As String
    Dim strFTPPaswd As String
    Dim strFTPDir As String
    Dim cFTP As New clsFtp
    Dim lngResult As Long
    Dim lngLocalFileSize As Long
    Dim lngFTPFileSize As Long
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim StrMessage As String
    
    On Error GoTo err
    Set funLoadOneImage = Nothing
    
    If intLoadType = 0 Then         '0--��ZLSeriesInfosװ��
        '���ж�ͼ�������Ƿ���ȷ
        If ZLSeriesInfos.Count < intSeriesIndex Then Exit Function
        If ZLSeriesInfos(intSeriesIndex).ImageInfos.Count < intImageIndex Then Exit Function
        
        strSaveDir = ZLSeriesInfos(intSeriesIndex).strSaveDir
        strImageName = ZLSeriesInfos(intSeriesIndex).ImageInfos(intImageIndex).ImageName
        lngSource = ZLSeriesInfos(intSeriesIndex).lngSource
        strShareDir = ZLSeriesInfos(intSeriesIndex).strShareDir
        strHostIP = ZLSeriesInfos(intSeriesIndex).strHostIP
        strFTPUser = ZLSeriesInfos(intSeriesIndex).strFTPUser
        strFTPPaswd = ZLSeriesInfos(intSeriesIndex).strFTPPasw
        strFTPDir = ZLSeriesInfos(intSeriesIndex).strFTPDir
    ElseIf intLoadType = 1 Then     '1 -- ��ZLShowSeriesInfosװ��
        '���ж�ͼ�������Ƿ���ȷ
        If ZLShowSeriesInfos.Count < intSeriesIndex Then Exit Function
        If ZLShowSeriesInfos(intSeriesIndex).ImageInfos.Count < intImageIndex Then Exit Function
        
        strSaveDir = ZLShowSeriesInfos(intSeriesIndex).strSaveDir
        strImageName = ZLShowSeriesInfos(intSeriesIndex).ImageInfos(intImageIndex).ImageName
        lngSource = ZLShowSeriesInfos(intSeriesIndex).lngSource
        strShareDir = ZLShowSeriesInfos(intSeriesIndex).strShareDir
        strHostIP = ZLShowSeriesInfos(intSeriesIndex).strHostIP
        strFTPUser = ZLShowSeriesInfos(intSeriesIndex).strFTPUser
        strFTPPaswd = ZLShowSeriesInfos(intSeriesIndex).strFTPPasw
        strFTPDir = ZLShowSeriesInfos(intSeriesIndex).strFTPDir
    End If
    
    If lngSource = 0 Then
        strLocalImage = PstrBufferImagePath & strSaveDir & "\" & strImageName
        Call MkLocalDir(PstrBufferImagePath & strSaveDir)
    Else
        strLocalImage = strSaveDir & "\" & strImageName
    End If
    
    If Dir(strLocalImage) <> vbNullString Then
        '�ӱ�������Ŀ¼�ж�ȡ�ļ�
        Set funLoadOneImage = ReadImage(strLocalImage, True)
        If funLoadOneImage Is Nothing Then
            Debug.Print "�򿪴��󣬿�������ռ��"
            If FileIsOccupied(strLocalImage) = True Then
                Debug.Print Now & " �ļ�����ռ��"
                '��ʱ
                TimeDelay 2000
                Debug.Print Now
                
                Set funLoadOneImage = ReadImage(strLocalImage, True)
                If funLoadOneImage Is Nothing Then
                    '��ʱ
                    TimeDelay 2000
                    Debug.Print Now
                    
                    Set funLoadOneImage = ReadImage(strLocalImage, True)
                    If funLoadOneImage Is Nothing Then
                        '��ʱ
                        TimeDelay 2000
                        Debug.Print Now
                        
                        Set funLoadOneImage = ReadImage(strLocalImage, True)
                        If funLoadOneImage Is Nothing Then
                            Debug.Print "goto errdown"
                            GoTo errDown
                        End If
                    End If
                End If
                Debug.Print "��ʱ��ȡ�ɹ�"
            Else
                Debug.Print "�򿪳����������� goto errdown"
                GoTo errDown
            End If
        End If
    Else
errDown:
        If strShareDir <> "" Then
            'ͨ������Ŀ¼��ȡ�ļ�
            Set funLoadOneImage = imgs.ReadFile("\\" & strHostIP & "\" & strShareDir & "\" & strSaveDir & "\" & strImageName)
        Else
            'ͨ��FTP�����ļ�
            '����FTP
            cFTP.FuncFtpConnect strHostIP, strFTPUser, strFTPPaswd
ReDownFile:
            '�����ļ�
            lngResult = cFTP.FuncDownloadFile(strFTPDir & Replace(strSaveDir, "\", "/"), strLocalImage, strImageName)
            '���سɹ��󣬶Աȱ����ļ���FTP�ļ���С�Ƿ�һ��
            If lngResult = 0 And gblnCompareSize Then
                lngLocalFileSize = objFileSystem.GetFile(strLocalImage).Size
                lngFTPFileSize = cFTP.FuncFtpGetFileSize(strFTPDir & Replace(strSaveDir, "\", "/"), strImageName)
                
                If lngLocalFileSize < lngFTPFileSize Then
                    StrMessage = "���غ���ļ���С��" & lngLocalFileSize & "����FTP�е��ļ���С��" & lngFTPFileSize & "����һ�£�" & vbCrLf & _
                                 "�����ļ���" & strLocalImage & vbCrLf & _
                                 "FTP�ļ���" & strFTPDir & Replace(strSaveDir, "\", "/") & strImageName & vbCrLf & _
                                 "�Ƿ���Ҫ�������أ�"
                    If MsgBox(StrMessage, vbQuestion + vbYesNo, "��ʾ") = vbYes Then
                        GoTo ReDownFile
                    End If
                End If
            End If
            cFTP.FuncFtpDisConnect
            
            Debug.Print "lngResult = " & lngResult
            
            If lngResult = 0 Then  '��ǰͼ�����سɹ�
                Set funLoadOneImage = imgs.ReadFile(strLocalImage)
            End If
        End If
    End If
     
    cFTP.FuncFtpDisConnect
    Exit Function
err:
    cFTP.FuncFtpDisConnect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub subMoveViewers(thisForm As frmViewer, intRow As Integer, intCol As Integer)
'------------------------------------------------
'���ܣ� ��������Viwer��λ�ã�����intRow�·���intCol�ҷ�������Viewer��λ��
'������ thisForm--��Ƭվ����
'       intRow--Viwer���ڵ���
'       intCol -- Viewer���ڵ���
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim i As Integer
    
    For i = 1 To thisForm.Viewer.Count - 1
        If ZLShowSeriesInfos(i).ImageInfos.Count <> 0 Then
            If ZLShowSeriesInfos(i).intRow >= intRow And ZLShowSeriesInfos(i).intCol >= intCol Then
                Call subPlaceAViewer(thisForm, i, ZLShowSeriesInfos(i).intRow, ZLShowSeriesInfos(i).intCol)
            End If
        End If
    Next i
End Sub


Public Sub subOpenFiles(thisForm As frmViewer)
'------------------------------------------------
'���ܣ� �������ļ����壬��ͼ���ļ�
'������ thisForm--��Ƭվ����
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim arrFileList As OpenFileArray
    
    On Error GoTo err
    
    arrFileList = funGetFileList(thisForm)
    '��������ݣ�����ļ���
    If arrFileList.FilePath <> "" Then
        Call subOpenFileList(thisForm, arrFileList)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subOpenFileList(thisForm As frmViewer, arrFileList As OpenFileArray)
'------------------------------------------------
'���ܣ� ���ļ��б��е�ͼ��
'������ thisForm -- ��Ƭվ����
'������ arrFileList -- Ҫ�򿪵��ļ��б�
'���أ���
'------------------------------------------------
    Dim img As New DicomImage
    Dim iFileIndex As Integer
    Dim intSeriesIndex As Integer
    Dim intImageIndex As Integer
    Dim blnNewSeries As Boolean
    Dim blnNewImage As Boolean
    Dim iPrivImage As Integer
    Dim lngImageNo As Long
    Dim oneSeriesInfo As clsSeriesInfo
    Dim oneImageInfo As clsImageInfo
    Dim i As Integer
    Dim j As Integer

    On Error GoTo err
    
    If arrFileList.FilePath = "" Then Exit Sub
    
    
    '�Ѵ򿪵��ļ������ص�        ZLSeriesInfos �ṹ��
    For iFileIndex = 1 To UBound(arrFileList.Filename)
        Set img = ReadImage(arrFileList.FilePath & arrFileList.Filename(iFileIndex))
        
        If Not IsNull(img.Attributes(&H20, &H13).Value) Then
            lngImageNo = Val(img.Attributes(&H20, &H13).Value)
        Else
            lngImageNo = 0
        End If
        '��������UID�����������Ƿ����
        blnNewSeries = True
        For intSeriesIndex = 1 To ZLSeriesInfos.Count
            If ZLSeriesInfos(intSeriesIndex).MultiFrame = 1 Then
                '�����κβ�������֡���в�������ͼ�񣬵��Ƕ��ڶ�֡ͼ��Ҫ�ж��Ƿ����ǰͼ����ͬһ��ͼ
                If img.FrameCount > 1 And ZLSeriesInfos(intSeriesIndex).SeriesUID = img.SeriesUID _
                    And ZLSeriesInfos(intSeriesIndex).ImageInfos(1).InstanceUID = img.InstanceUID Then
                    'ͬһ��ͼ�����ʹ��ͬһ�����У����洦��ͼ���ʱ����жϲ�����ͼ��
                    blnNewSeries = False
                    Exit For
                End If
            Else
                '��֡���вſ��ǰ�ͼ����ӵ�������
                If ZLSeriesInfos(intSeriesIndex).SeriesUID = img.SeriesUID And img.FrameCount = 1 Then
                    blnNewSeries = False
                    Exit For
                End If
            End If
        Next intSeriesIndex
        
        If blnNewSeries = True Then '����������
            '����������
            Set oneSeriesInfo = funGetNewSeriesInfo
            oneSeriesInfo.lngSource = 1     'ֱ��ͨ���򿪷�ʽ���ص�����
            oneSeriesInfo.SeriesNo = GetImageAttribute(img.Attributes, ATTR_���к�)
            oneSeriesInfo.SeriesUID = img.SeriesUID
            oneSeriesInfo.strModality = GetImageAttribute(img.Attributes, ATTR_Ӱ�����)
            oneSeriesInfo.strSaveDir = arrFileList.FilePath
            oneSeriesInfo.StudyUID = img.StudyUID
            oneSeriesInfo.MultiFrame = IIf(img.FrameCount = 1, 0, 1)
            
            '��ȡԤ�贰��λ
            For i = 1 To UBound(aPresetWinWL, 2)
                If UCase(aPresetWinWL(3, i).strModality) = UCase(oneSeriesInfo.strModality) Then
                    For j = 3 To 12
                        If aPresetWinWL(j, i).bInUse And aPresetWinWL(j, i).intDefault = 1 Then
                            oneSeriesInfo.lngWinWidth = aPresetWinWL(j, i).lngWinWidth
                            oneSeriesInfo.lngWinLevel = aPresetWinWL(j, i).lngWinLevel
                            Exit For
                        End If
                    Next j
                    Exit For
                End If
            Next i
            
            ZLSeriesInfos.Add oneSeriesInfo, CStr(ZLSeriesInfos.Count + 1)
            
            '��дͼ��λ��
            iPrivImage = 1
            blnNewImage = True
        Else    'ʹ��ԭ������
            '����ͼ��UID���ж�ͼ���Ƿ���ڣ����ͼ����ڣ���ͬʱ����ͼ��λ��
            blnNewImage = True
            iPrivImage = 0
            For intImageIndex = 1 To ZLSeriesInfos(intSeriesIndex).ImageInfos.Count
                If ZLSeriesInfos(intSeriesIndex).ImageInfos(intImageIndex).ImageNo < lngImageNo Then
                    iPrivImage = intImageIndex
                End If
                
                If img.InstanceUID = ZLSeriesInfos(intSeriesIndex).ImageInfos(intImageIndex).InstanceUID Then
                    blnNewImage = False
                    Exit For
                End If
            Next intImageIndex
        End If
        
        '���ͼ��
        If blnNewImage = True Then
            '����ͼ��
            Set oneImageInfo = funGetNewImageInfo
            oneImageInfo.AcquisitionTime = Format(GetImageAttribute(img.Attributes, ATTR_�ɼ�����) & " " & GetImageAttribute(img.Attributes, ATTR_�ɼ�ʱ��), "yyyy-MM-dd HH:MM:SS")
            oneImageInfo.Columns = img.sizeX
            oneImageInfo.FrameOfReferenceUID = GetImageAttribute(img.Attributes, ATTR_�ο�֡UID)
            oneImageInfo.ImageName = arrFileList.Filename(iFileIndex)
            oneImageInfo.ImageNo = GetImageAttribute(img.Attributes, ATTR_ͼ���)
            oneImageInfo.ImageOrientationPatient = GetImageAttribute(img.Attributes, ATTR_ͼ������)
            oneImageInfo.ImagePositionPatient = GetImageAttribute(img.Attributes, ATTR_ͼ��λ�ò���)
            oneImageInfo.ImageTime = Format(GetImageAttribute(img.Attributes, ATTR_ͼ������) & " " & GetImageAttribute(img.Attributes, ATTR_ͼ��ʱ��), "yyyy-MM-dd HH:MM:SS")
            oneImageInfo.InstanceUID = img.InstanceUID
            oneImageInfo.PixelSpacing = GetImageAttribute(img.Attributes, ATTR_���ؾ���)
            oneImageInfo.Rows = img.sizeY
            oneImageInfo.SliceLocation = GetImageAttribute(img.Attributes, ATTR_��Ƭλ��)
            oneImageInfo.SliceThickness = GetImageAttribute(img.Attributes, ATTR_���)
            
            If iPrivImage = 0 Then iPrivImage = 1   '�ڵ�һ��ͼ��ǰ��׷��
            If ZLSeriesInfos(intSeriesIndex).ImageInfos.Count = 0 Then
                ZLSeriesInfos(intSeriesIndex).ImageInfos.Add oneImageInfo
            Else
                ZLSeriesInfos(intSeriesIndex).ImageInfos.Add oneImageInfo, , , iPrivImage
            End If
        End If
    Next iFileIndex
    
    '����Ҫ��ʾ��ͼ����ؽ���VIEWER,��Viewer��ʾ�������У���ʾViewer��صĹ�����
    Call subShowForm(thisForm)
    '������ʾ����ͼ
    Call subShowMiniImages(thisForm)
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subLookOrBrowsSwitch(thisForm As frmViewer)
'------------------------------------------------
'���ܣ� �л�����͹۲�ģʽ,���°ڷŷָ��������°ڷŴ���
'       ����۲�ģʽʱ��ֻ��ʾ��ǰ��ѡ�е�Viewer������������Viewer��Visible=False��
'       �������ģʽʱ��ѭ����ǰ�Ѿ����ڵ�Viewer�������ǵ�Visible��ΪTrue�������°ڷ���ЩViewer��
'       �л�����͹۲�ģʽ������ɾ��Viewer�͸ı�ZLShowSeriesInfos�ṹ��ֻ����ʾ������Viewer���ѡ�
'       ֻ�����л����ֵ�ʱ�򣬲Ż�ı�ZLShowSeriesInfos�ṹ��
'������ thisForm--��Ƭվ����
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim blnExit As Boolean
    Dim intViwerIndex As Integer
    
    If thisForm.intSelectedSerial < 1 Then Exit Sub
    If thisForm.intSelectedSerial >= thisForm.Viewer.Count Then Exit Sub
    On Error GoTo err
    
    Button_miLookOrBrowse = Not Button_miLookOrBrowse
    thisForm.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_View_OneBrowse, , True).Checked = Button_miLookOrBrowse
    thisForm.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_View_OneBrowse, , True).Checked = Button_miLookOrBrowse
    thisForm.ComToolBar.RecalcLayout
            
    '�ڷŷָ���
    Call subShowSpliter(thisForm)
    
    If Button_miLookOrBrowse = True Then    '�۲�ģʽ��ֻ�ڷ�һ��Viewer
        '����������Viewer
        For i = 1 To thisForm.Viewer.Count - 1
            If i = thisForm.intSelectedSerial Then
                thisForm.Viewer(i).Visible = True
            Else
                thisForm.Viewer(i).Visible = False
                thisForm.VScro(i).Visible = False
            End If
        Next i
        '�ڷ����Viewer�����ù�����
        Call subPlaceAViewer(thisForm, thisForm.intSelectedSerial, 1, 1)
    Else    '���ģʽ���ڷ�����Viewer
        'ѭ���������в��֣��ڷ����е�Viewer
        intViwerIndex = 0
        blnExit = False
        For i = 1 To thisForm.intCountY
            For j = 1 To thisForm.intCountX
                intViwerIndex = intViwerIndex + 1
                If intViwerIndex >= thisForm.Viewer.Count Then
                    blnExit = True
                    Exit For
                End If
                '�ڷ�Viewer������ʾViewer
                thisForm.Viewer(intViwerIndex).Visible = True
                Call subPlaceAViewer(thisForm, intViwerIndex, i, j)
            Next j
            If blnExit = True Then
                Exit For
            End If
        Next i
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub subDisplayPatientInfo(thisViewer As DicomViewer)
'------------------------------------------------
'���ܣ� ��ʾ��ر�ָ��Viewer�Ĳ�����Ϣ�����Ҹ���ͼ���С�����Ƿ��Զ����ز�����Ϣ
'������ thisViewer--��Ҫ��ʾ���Źرղ�����Ϣ��Viewer
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim img As DicomImage
    Dim i As Integer
    Dim j As Integer
    Dim blnShowLabel As Boolean
    Dim intImageIndex As Integer
    
    On Error GoTo err
    If thisViewer.Images.Count < 1 Then Exit Sub
    
    'Button_miDispPatientInfo    ---������ʾ
    'blnpatientInfoScaleFontSize ---������Ϣ���ִ�С�Ƿ�����ͼ��һ������
    
    If ((Not blnpatientInfoScaleFontSize) And _
            (thisViewer.width / thisViewer.MultiColumns / Screen.TwipsPerPixelX < lngPatientInfoInvisibleSize Or _
             thisViewer.height / thisViewer.MultiRows / Screen.TwipsPerPixelY < lngPatientInfoInvisibleSize)) Then
        blnShowLabel = False
    Else
        blnShowLabel = True
    End If
    
    If blnShowLabel = True Then blnShowLabel = Button_miDispPatientInfo
    
    intImageIndex = thisViewer.CurrentIndex
    '���������Ϣ����ͼ�����ţ���ͼ��С��ָ����С������ʾ������Ϣ
    For i = 1 To thisViewer.MultiColumns
        For j = 1 To thisViewer.MultiRows
            Set img = thisViewer.Images(intImageIndex)
            Call subInitImageLabels(thisViewer.Index, 1, img, blnShowLabel)
            intImageIndex = intImageIndex + 1
            If intImageIndex > thisViewer.Images.Count Then
                Exit Sub
            End If
        Next j
    Next i
        
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subSortImages(thisForm As frmViewer, intViewerIndex As Integer, iSortType As Integer)
'------------------------------------------------
'���ܣ� �Թ�Ƭ�����е�ͼ��������򣬲μ������Viewr������ΪintViewerIndex
'������ thisForm--��������Ĵ���
'       intViewerIndex -- ���������Viewer������
'       iSortType -- ����ʽ��0--ͼ��ţ�1--��λ����2--��λ����3--�ɼ�ʱ�䣻4--ͼ��ʱ�䡣
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim SortShowImageInfos As Collection
    Dim thisViewer As DicomViewer
    Dim intImagesCount As Integer
    Dim OneImageInfos As clsImageInfo
    Dim tmpListItem As ListItem
    Dim iOldIndex As Integer
    Dim SortImages As New DicomImages
    Dim i As Integer
    Dim j As Integer
    Dim k As String
    
    On Error GoTo err
    
    If ZLShowSeriesInfos(intViewerIndex).intSortType = iSortType Then Exit Sub
    
    '���ȶ�ZLShowSeriesInfos�е�ͼ���������
    ZLShowSeriesInfos(intViewerIndex).intSortType = iSortType
    
    Set thisViewer = thisForm.Viewer(intViewerIndex)
    intImagesCount = ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
    If intImagesCount = 0 Then Exit Sub
    
    '���ZLShowSeriesInfos��û������Ĺؼ�����Ϣ�����ͼ���ж�ȡ��Щ��Ϣ
    If ZLShowSeriesInfos(intViewerIndex).ImageInfos(1).SliceLocation = "" Then
        Call ReadSortInfoFromImage(thisViewer)
    End If
    
    '����һ��SortShowImageInfos����������ZLShowSeriesInfos.ImageInfos�ĸ���
    Set SortShowImageInfos = ZLShowSeriesInfos(intViewerIndex).ImageInfos
    
    '���ZLShowSeriesInfos.ImageInfos�е�����
    Set ZLShowSeriesInfos(intViewerIndex).ImageInfos = Nothing
    
    '��ListView��������
    '������ؼ�����д����ListView
    thisForm.lvwSort.ListItems.Clear
    For i = 1 To intImagesCount
        If iSortType = 0 Then           '����ͼ�������
            Set tmpListItem = thisForm.lvwSort.ListItems.Add(, , SortShowImageInfos(i).ImageNo)
            tmpListItem.Text = String(6 - Len(tmpListItem.Text), "0") & tmpListItem.Text
        ElseIf iSortType = 1 Or iSortType = 2 Then  '����Ƭλ�������������
            Set tmpListItem = thisForm.lvwSort.ListItems.Add(, , SortShowImageInfos(i).SliceLocation)
            k = Val(tmpListItem.Text) * 100 + 100000 '��֤�����͸�����Ƭλ�ö��ܹ�ͳһ��������
            k = Format(k, "#0")
            If Len(k) > 8 Then
                k = left(k, 8)
            End If
            tmpListItem.Text = String(8 - Len(k), "0") & k
        ElseIf iSortType = 3 Then       '���ղɼ�ʱ��
            Set tmpListItem = thisForm.lvwSort.ListItems.Add(, , SortShowImageInfos(i).AcquisitionTime)
        Else                            '����ͼ��ʱ��
            Set tmpListItem = thisForm.lvwSort.ListItems.Add(, , SortShowImageInfos(i).ImageTime)
        End If
        tmpListItem.SubItems(1) = i
    Next i
    
    '��ListView���ı���������
    thisForm.lvwSort.SortKey = 0
    thisForm.lvwSort.Sorted = True
    If iSortType = 2 Then   '��Ƭλ������
        thisForm.lvwSort.SortOrder = lvwDescending
    Else                    '��������������
        thisForm.lvwSort.SortOrder = lvwAscending
    End If
    
    '������ɺ󣬸����¾���������ͼ����Ϣ��SortShowImageinfos���ƻ�ZLShowSeriesInfos.ImageInfos
    For i = 1 To thisForm.lvwSort.ListItems.Count
        iOldIndex = Val(thisForm.lvwSort.ListItems(i).SubItems(1))
        ZLShowSeriesInfos(intViewerIndex).ImageInfos.Add SortShowImageInfos(iOldIndex)
    Next i
    
    'Ȼ�����µ���Viewer��ͼ���λ�ú�Tag
    For i = 1 To thisViewer.Images.Count
        SortImages.Add thisViewer.Images(i)
    Next i
    thisViewer.Images.Clear
    For i = 1 To intImagesCount
        'ֻ�����Ѿ���ʾ�˵�ͼ��
        If ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).blnDisplayed = True Then
            For j = 1 To SortImages.Count
                If ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).InstanceUID = SortImages(j).InstanceUID Then
                    thisViewer.Images.Add SortImages(j)
                    thisViewer.Images(thisViewer.Images.Count).Tag = i
                    SortImages.Remove (j)
                    Exit For
                End If
            Next j
        End If
    Next i
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub ReadSortInfoFromImage(thisViewer As DicomViewer)
'------------------------------------------------
'���ܣ� ��ͼ���ж�ȡ������Ϣ����ZLSeriesInfos�ṹ
'������ intSeriesIndex--ͼ���������е�����
'       intViewerIndex -- ���������Viewer������
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim img As DicomImage
    Dim intViewerIndex As Integer
    Dim intImageCount As Integer
    Dim i As Integer
    Dim j  As Integer
    
    On Error GoTo err
    intViewerIndex = thisViewer.Index
    If intViewerIndex = 0 Then Exit Sub
    'ѭ������������ͼ��
    intImageCount = ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
    For i = 1 To intImageCount
        '���ͼ���Ѿ���ʾ�����Viewer��ͼ������ȡ��Ϣ
        If ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).blnDisplayed = True Then
            For j = IIf(thisViewer.Images.Count >= i, i, thisViewer.Images.Count) To 1 Step -1
                If ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).InstanceUID = thisViewer.Images(j).InstanceUID Then
                    Set img = thisViewer.Images(j)
                    Exit For
                End If
            Next j
        Else    '���ͼ�񲻴��ڣ�����ͼ��
            Set img = funLoadAImage(intViewerIndex, i, 1)
        End If
        
        '��ͼ���е���Ϣ��ӵ�ZLSeriesInfos�ṹ��
        If Not img Is Nothing Then
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).AcquisitionTime = Format(GetImageAttribute(img.Attributes, ATTR_�ɼ�����) & " " & GetImageAttribute(img.Attributes, ATTR_�ɼ�ʱ��), "yyyy-MM-dd HH:MM:SS")
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).Columns = img.sizeX
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).FrameOfReferenceUID = GetImageAttribute(img.Attributes, ATTR_�ο�֡UID)
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).ImageOrientationPatient = GetImageAttribute(img.Attributes, ATTR_ͼ������)
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).ImagePositionPatient = GetImageAttribute(img.Attributes, ATTR_ͼ��λ�ò���)
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).ImageTime = Format(GetImageAttribute(img.Attributes, ATTR_ͼ������) & " " & GetImageAttribute(img.Attributes, ATTR_ͼ��ʱ��), "yyyy-MM-dd HH:MM:SS")
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).PixelSpacing = GetImageAttribute(img.Attributes, ATTR_���ؾ���)
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).Rows = img.sizeY
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).SliceLocation = GetImageAttribute(img.Attributes, ATTR_��Ƭλ��)
            ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).SliceThickness = GetImageAttribute(img.Attributes, ATTR_���)
        End If
    Next i
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subResizeSeries(thisForm As frmViewer)
'------------------------------------------------
'���ܣ� �ı䴰�ڴ�С�����µ�������Viewer��λ��
'������ thisForm --- ��Ƭ����
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim intViewerIndex As Integer
    Dim blnExit As Boolean
    
    On Error GoTo err
    
    '�������в��֣��ڷŷָ���
    Call subShowSpliter(thisForm)
    
    'ѭ�����е�Viewer�����°ڷ�λ��
    If Button_miLookOrBrowse = True Then       '�۲�ģʽ��ֻ�ڷ�һ��Viewer
        If thisForm.intSelectedSerial > 1 And thisForm.intSelectedSerial < thisForm.Viewer.Count Then
            '�ڷ����Viewer�����ù�����
            Call subPlaceAViewer(thisForm, thisForm.intSelectedSerial, 1, 1)
        End If
    Else
        intViewerIndex = 0
        For i = 1 To thisForm.intCountY
            For j = 1 To thisForm.intCountX
                intViewerIndex = intViewerIndex + 1
                If intViewerIndex >= thisForm.Viewer.Count Then
                    blnExit = True
                    Exit For
                End If
                '�ڷ����Viewer
                Call subPlaceAViewer(thisForm, intViewerIndex, i, j)
            Next j
            If blnExit = True Then Exit For
        Next i
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub subDisplayReferLine(thisViewer As DicomViewer, thisForm As frmViewer, blnCurrentOnly As Boolean)
'------------------------------------------------
'���ܣ� ���ݲ˵�ѡ���ʾ�������͵Ķ�λ��
'������ thisViewer ---��ǰѡ�е�Viewer������ʾ��Viewer������ViewerͶӰ�Ķ�λ��
'       thisForm --- ��Ƭ����
'       blnCurrentOnly --- ֻˢ�µ�ǰ��λ��
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim strLineTag As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim blnExit As Boolean
    Dim intCurrentImageIndex As Integer
    Dim imgSource As New DicomImage
    Dim imgDest As DicomImage
    Dim viewerDest As DicomViewer
    Dim intDestViewerIndex As Integer
    
    If thisViewer.Images.Count = 0 Then Exit Sub
    
    On Error GoTo err
    '���ö�λ��TAG��ǰ����
    If blnCurrentOnly = True Then
        strLineTag = "RLC"
    Else
        strLineTag = "RL"
    End If
    
    For Each viewerDest In thisForm.Viewer
        If viewerDest.Index <> 0 And viewerDest.Images.Count > 0 Then
            intCurrentImageIndex = viewerDest.CurrentIndex
            
            blnExit = False
            For i = 1 To viewerDest.MultiRows
                For j = 1 To viewerDest.MultiColumns
                    'ɾ��thisViewer�пɼ�ͼ���У��ɵĶ�λ��
                    subDeleteAppointLabel viewerDest.Images(intCurrentImageIndex), strLineTag
                    '������һ��ͼ��
                    intCurrentImageIndex = intCurrentImageIndex + 1
                    If intCurrentImageIndex > viewerDest.Images.Count Then
                        blnExit = True
                        Exit For
                    End If
                Next j
                If blnExit = True Then Exit For
            Next i
        End If
        'ˢ��Viewer
        viewerDest.Refresh
    Next
    
    If Button_miAllReferLine = False And Button_miFLReferLine = False And Button_miCurrentReferLine = False Then Exit Sub
    
    '����ȫ����λ�ߺ���β��λ�ߣ���Ҫ��ͼ�����ͼ�������������
    If Button_miAllReferLine Or Button_miFLReferLine Then
        '���ͼ�������ʽ������ͼ���������
        Call subSortImages(thisForm, thisViewer.Index, 0)
    End If
    
    For Each viewerDest In thisForm.Viewer
        intDestViewerIndex = viewerDest.Index
        If intDestViewerIndex <> 0 And intDestViewerIndex <> thisViewer.Index And viewerDest.Images.Count > 0 Then
            intCurrentImageIndex = viewerDest.CurrentIndex
            
            blnExit = False
            For i = 1 To viewerDest.MultiRows
                For j = 1 To viewerDest.MultiColumns
                    '��Ŀ��ͼ imgDest�л���λ��
                    Set imgDest = viewerDest.Images(intCurrentImageIndex)
                    
                    '����λ��,��ͬ��ʽ�Ķ�λ�߷ֱ���
                    If Button_miAllReferLine = True Then    '��ʾ���ж�λ��
                        For k = 1 To ZLShowSeriesInfos(thisViewer.Index).ImageInfos.Count
                            '��д����ͼ������ݣ�Ȼ����㶨λ��
                            Call subWriteRefLineImage(imgSource, k, thisViewer)
                            Call subDrawRefLine(imgSource, imgDest, True, "RLL", True)
                        Next k
                    End If
                    
                    If Button_miFLReferLine = True Then     '��ʾ��β��λ��
                        '��д����ͼ������ݣ�Ȼ����㶨λ��
                        Call subWriteRefLineImage(imgSource, 1, thisViewer)
                        Call subDrawRefLine(imgSource, imgDest, False, "RLL", True)
                        If ZLShowSeriesInfos(thisViewer.Index).ImageInfos.Count > 1 Then
                            Call subWriteRefLineImage(imgSource, ZLShowSeriesInfos(thisViewer.Index).ImageInfos.Count, thisViewer)
                            Call subDrawRefLine(imgSource, imgDest, False, "RLL", True)
                        End If
                    End If
                    
                    If Button_miCurrentReferLine = True Then    '��ʾ��ǰ��λ��
                        '�����ǰͼ������βͼ�񣬲����Ѿ���ʾ����β��λ�ߣ��򲻴���
                        If Not (Button_miFLReferLine And (thisForm.SelectedImage.Tag = 1 Or thisForm.SelectedImage.Tag = ZLShowSeriesInfos(thisViewer.Index).ImageInfos.Count)) Then
                            Call subWriteRefLineImage(imgSource, thisForm.SelectedImage.Tag, thisViewer)
                            Call subDrawRefLine(imgSource, imgDest, False, "RLC", True)
                        End If
                    End If
                    
                    '������һ��ͼ��
                    intCurrentImageIndex = intCurrentImageIndex + 1
                    If intCurrentImageIndex > viewerDest.Images.Count Then
                        blnExit = True
                        Exit For
                    End If
                Next j
                If blnExit = True Then Exit For
            Next i
        End If
        'ˢ��Viewer
        viewerDest.Refresh
    Next
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subWriteRefLineImage(img As DicomImage, intImageIndex As Integer, thisViewer As DicomViewer)
'------------------------------------------------
'���ܣ� �Ѷ�λ����Ϣ��д��imgͼ����
'������ img ---��Ҫ��д��λ����Ϣ��ͼ��
'       intViewerIndex --- ͼ������Viewer������
'       intImageIndex --- ͼ�����ڵ�ͼ������
'       thisViewer --- ͼ�����ڵ�Viewer
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim oneImageInfo As clsImageInfo
    Dim tmpValues() As String
    Dim intViewerIndex As Integer
    
    On Error GoTo err
    intViewerIndex = thisViewer.Index
    
    '���ж�ZLShowSeriesInfos�Ƿ������ݣ����û�У����ͼ���ж�ȡ����
    If ZLShowSeriesInfos(intViewerIndex).ImageInfos(intImageIndex).SliceLocation = "" Then
        Call ReadSortInfoFromImage(thisViewer)
    End If
    '��ʼ��д��Ϣ
     
    Set oneImageInfo = ZLShowSeriesInfos(intViewerIndex).ImageInfos(intImageIndex)
    
    img.Attributes.Add &H20, &H13, oneImageInfo.ImageNo
    img.Attributes.Add &H28, &H10, oneImageInfo.Rows
    img.Attributes.Add &H28, &H11, oneImageInfo.Columns
    img.Attributes.Add &H20, &H1041, oneImageInfo.SliceLocation
    img.Attributes.Add &H20, &H52, oneImageInfo.FrameOfReferenceUID
    tmpValues = Split(oneImageInfo.PixelSpacing, "\")
    img.Attributes.Add &H28, &H30, tmpValues
    tmpValues = Split(oneImageInfo.ImagePositionPatient, "\")
    img.Attributes.Add &H20, &H32, tmpValues
    tmpValues = Split(oneImageInfo.ImageOrientationPatient, "\")
    img.Attributes.Add &H20, &H37, tmpValues
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function funGetNewSeriesInfo() As clsSeriesInfo
'------------------------------------------------
'���ܣ� ����һ���µ�������Ϣ�����Ҹ����ʼֵ
'������ ��
'���أ� ��
'ʱ�䣺2009-7
'------------------------------------------------
    Dim oneSeriesInfo As New clsSeriesInfo
    oneSeriesInfo.blnImageSyn = True
    oneSeriesInfo.FilterLength = 0
    oneSeriesInfo.FlipState = doFlipNormal
    oneSeriesInfo.intSortType = 0   '��¼��ǰ���е�����ʽ��0--ͼ��ţ�1--��λ����2--��λ����3--�ɼ�ʱ�䣻4--ͼ��ʱ�䣬����ZLShowSeriesInfos��ʹ�á�
    oneSeriesInfo.lngSource = 0     'ͼ����Դ��0-��PACSͼ����������أ�1-ֱ�Ӵ��ļ���2---��ϣ�3-�������ɵ����У�����ʸ��״λ�ؽ���ͼ��ƴ�ӡ�α�����ɵ�ͼ��
    oneSeriesInfo.lngWinLevel = 0   '0 ��ʾû��Ԥ��Ĵ���λ
    oneSeriesInfo.lngWinWidth = 0   '0 ��ʾû��Ԥ��Ĵ���λ
    oneSeriesInfo.RotateState = doRotateNormal
    oneSeriesInfo.ScrollX = 0
    oneSeriesInfo.ScrollY = 0
    oneSeriesInfo.StretchToFit = True
    oneSeriesInfo.UnsharpEnhancement = 0    '��Ե��ǿǿ�ȡ�����ZLShowSeriesInfos��ʹ�á�
    oneSeriesInfo.UnsharpLength = 0         '��Ե��ǿ���ȡ�����ZLShowSeriesInfos��ʹ�á�
    oneSeriesInfo.Zoom = 1
    oneSeriesInfo.MultiFrame = 0            'Ĭ���ǵ�֡ͼ��
    oneSeriesInfo.Selected = False          'Ĭ��û�б�ѡ��
    oneSeriesInfo.intCol = 0
    oneSeriesInfo.intRow = 0
    Set funGetNewSeriesInfo = oneSeriesInfo
End Function

Public Function funGetNewImageInfo() As clsImageInfo
'------------------------------------------------
'���ܣ� ����һ���µ�ͼ����Ϣ�����Ҹ����ʼֵ
'������ ��
'���أ� ��
'ʱ�䣺2009-7
'------------------------------------------------
    Dim oneImageInfo As New clsImageInfo
    oneImageInfo.blnDisplayed = False
    oneImageInfo.blnSelected = False
    oneImageInfo.int3DLabelIndex = 0        '��ʾû����ά��궨λ��
    oneImageInfo.blnPrinted = False         '��ʾû�б���ӡ
    Set funGetNewImageInfo = oneImageInfo
End Function

Public Function funCopySeriesInfo(sourceSeriesInfo As clsSeriesInfo, destSeriesInfo As clsSeriesInfo) As Boolean
'------------------------------------------------
'���ܣ� ����������Ϣ,ֻ�������е���Ϣ�������Ƶ����ݰ�����������ͼ����Ϣ�����е�λ��intCol��intRow
'������ sourceSeriesInfo --- Դ����
'       destSeriesInfo ---- Ŀ������
'���أ� �Ƿ�ɹ�
'ʱ�䣺2009-7
'------------------------------------------------
    destSeriesInfo.blnImageSyn = sourceSeriesInfo.blnImageSyn
    destSeriesInfo.FilterLength = sourceSeriesInfo.FilterLength
    destSeriesInfo.FlipState = sourceSeriesInfo.FlipState
    destSeriesInfo.intSortType = sourceSeriesInfo.intSortType
    destSeriesInfo.lngSource = sourceSeriesInfo.lngSource
    destSeriesInfo.lngWinLevel = sourceSeriesInfo.lngWinLevel
    destSeriesInfo.lngWinWidth = sourceSeriesInfo.lngWinWidth
    destSeriesInfo.RotateState = sourceSeriesInfo.RotateState
    destSeriesInfo.ScrollX = sourceSeriesInfo.ScrollX
    destSeriesInfo.ScrollY = sourceSeriesInfo.ScrollY
    destSeriesInfo.SeriesNo = sourceSeriesInfo.SeriesNo
    destSeriesInfo.SeriesUID = sourceSeriesInfo.SeriesUID
    destSeriesInfo.StretchToFit = sourceSeriesInfo.StretchToFit
    destSeriesInfo.strFTPDir = sourceSeriesInfo.strFTPDir
    destSeriesInfo.strFTPPasw = sourceSeriesInfo.strFTPPasw
    destSeriesInfo.strFTPUser = sourceSeriesInfo.strFTPUser
    destSeriesInfo.strHostIP = sourceSeriesInfo.strHostIP
    destSeriesInfo.strModality = sourceSeriesInfo.strModality
    destSeriesInfo.strSaveDir = sourceSeriesInfo.strSaveDir
    destSeriesInfo.strShareDir = sourceSeriesInfo.strShareDir
    destSeriesInfo.strShareDirPasw = sourceSeriesInfo.strShareDirPasw
    destSeriesInfo.strShareDirUser = sourceSeriesInfo.strShareDirUser
    destSeriesInfo.StudyUID = sourceSeriesInfo.StudyUID
    destSeriesInfo.UnsharpEnhancement = sourceSeriesInfo.UnsharpEnhancement
    destSeriesInfo.UnsharpLength = sourceSeriesInfo.UnsharpLength
    destSeriesInfo.Zoom = sourceSeriesInfo.Zoom
    destSeriesInfo.MultiFrame = sourceSeriesInfo.MultiFrame
    destSeriesInfo.Selected = sourceSeriesInfo.Selected
    destSeriesInfo.strCName = sourceSeriesInfo.strCName
    destSeriesInfo.strEName = sourceSeriesInfo.strEName
    destSeriesInfo.strSex = sourceSeriesInfo.strSex
    destSeriesInfo.strAge = sourceSeriesInfo.strAge
    destSeriesInfo.strStudyID = sourceSeriesInfo.strStudyID
    destSeriesInfo.strOrderID = sourceSeriesInfo.strOrderID
    funCopySeriesInfo = True
End Function

Public Function funCopyImageInfo(sourceImageInfo As clsImageInfo, destImageInfo As clsImageInfo) As Boolean
'------------------------------------------------
'���ܣ� ����ͼ����Ϣ
'������ sourceImageInfo --- Դͼ��
'       destImageInfo ---- Ŀ��ͼ��
'���أ� �Ƿ�ɹ�
'ʱ�䣺2009-7
'------------------------------------------------
    destImageInfo.AcquisitionTime = sourceImageInfo.AcquisitionTime
    destImageInfo.blnDisplayed = sourceImageInfo.blnDisplayed
    destImageInfo.blnSelected = sourceImageInfo.blnSelected
    destImageInfo.Columns = sourceImageInfo.Columns
    destImageInfo.FrameOfReferenceUID = sourceImageInfo.FrameOfReferenceUID
    destImageInfo.ImageName = sourceImageInfo.ImageName
    destImageInfo.ImageNo = sourceImageInfo.ImageNo
    destImageInfo.ImageOrientationPatient = sourceImageInfo.ImageOrientationPatient
    destImageInfo.ImagePositionPatient = sourceImageInfo.ImagePositionPatient
    destImageInfo.ImageTime = sourceImageInfo.ImageTime
    destImageInfo.InstanceUID = sourceImageInfo.InstanceUID
    destImageInfo.PixelSpacing = sourceImageInfo.PixelSpacing
    destImageInfo.Rows = sourceImageInfo.Rows
    destImageInfo.SliceLocation = sourceImageInfo.SliceLocation
    destImageInfo.SliceThickness = sourceImageInfo.SliceThickness
    destImageInfo.int3DLabelIndex = sourceImageInfo.int3DLabelIndex
    destImageInfo.blnPrinted = sourceImageInfo.blnPrinted
    funCopyImageInfo = True
End Function

Public Sub subOpenCurrentImage(thisForm As frmViewer, img As DicomImage)
'------------------------------------------------
'���ܣ� �ѵ�ǰ��ͼ��򿪣�����ZLSeriesInfos�Ƚṹ
'������ thisForm--��ͼ��Ĵ���
'       img --- ��Ҫ�򿪵�ͼ��
'���أ� ��
'ʱ�䣺2009-7
'------------------------------------------------
    Dim i As Integer
    Dim blnNewSeries As Boolean
    Dim iCurrentSeries As Integer   '��ǰ���е�����
    Dim iSiblingSeries As Integer   '�ֵ����е�����
    Dim oneSeriesInfo As clsSeriesInfo
    Dim iCurrentImage As Integer
    Dim iPrivImage As Integer       '��ǰͼ���ǰһ��ͼ��
    Dim oneImageInfo As clsImageInfo
    
    On Error GoTo err
    '���ø�ͼ��������Ƿ��Ѿ�����
    blnNewSeries = True
    For i = 1 To ZLSeriesInfos.Count
        If ZLSeriesInfos(i).SeriesUID = img.SeriesUID Then
            blnNewSeries = False
            iCurrentSeries = i
        End If
        If ZLSeriesInfos(i).StudyUID = img.StudyUID Then
            iSiblingSeries = i
        End If
    Next i
    If iSiblingSeries = 0 Then Exit Sub     '����Ҳ����ֵ������򲻴�ͼ��
    
    '��������ڣ������Ӹ����У����������ֱ���ڸ������в���ͼ���Ƿ����
    '���в�����ͼ���谴��ͼ�������
    If blnNewSeries = True Then
        Set oneSeriesInfo = funGetNewSeriesInfo
        oneSeriesInfo.StudyUID = img.StudyUID
        oneSeriesInfo.SeriesUID = img.SeriesUID
        oneSeriesInfo.SeriesNo = 1
        oneSeriesInfo.strModality = ZLSeriesInfos(iSiblingSeries).strModality
        oneSeriesInfo.MultiFrame = 0
        oneSeriesInfo.strHostIP = ZLSeriesInfos(iSiblingSeries).strHostIP
        oneSeriesInfo.strFTPDir = ZLSeriesInfos(iSiblingSeries).strFTPDir
        oneSeriesInfo.strFTPPasw = ZLSeriesInfos(iSiblingSeries).strFTPPasw
        oneSeriesInfo.strFTPUser = ZLSeriesInfos(iSiblingSeries).strFTPUser
        oneSeriesInfo.strShareDir = ZLSeriesInfos(iSiblingSeries).strShareDir
        oneSeriesInfo.strShareDirUser = ZLSeriesInfos(iSiblingSeries).strShareDirUser
        oneSeriesInfo.strShareDirPasw = ZLSeriesInfos(iSiblingSeries).strShareDirPasw
        oneSeriesInfo.strSaveDir = ZLSeriesInfos(iSiblingSeries).strSaveDir
        ZLSeriesInfos.Add oneSeriesInfo, CStr(ZLSeriesInfos.Count + 1)
        iCurrentSeries = ZLSeriesInfos.Count
    End If
    
    '����ͼ���Ƿ����
    iCurrentImage = 0
    iPrivImage = 0
    For i = 1 To ZLSeriesInfos(iCurrentSeries).ImageInfos.Count
        If ZLSeriesInfos(iCurrentSeries).ImageInfos(i).InstanceUID = img.InstanceUID Then
            iCurrentImage = i
            Exit For
        End If
    Next i
    If iCurrentImage <> 0 Then Exit Sub         'ͼ���Ѿ����ڣ��򲻴�
    '��ͼ��
    Set oneImageInfo = funGetNewImageInfo
    oneImageInfo.InstanceUID = img.InstanceUID
    oneImageInfo.ImageNo = 1
    oneImageInfo.ImageName = img.InstanceUID
    oneImageInfo.AcquisitionTime = Format(GetImageAttribute(img.Attributes, ATTR_�ɼ�����) & " " & GetImageAttribute(img.Attributes, ATTR_�ɼ�ʱ��), "yyyy-MM-dd HH:MM:SS")
    oneImageInfo.ImageTime = Format(GetImageAttribute(img.Attributes, ATTR_ͼ������) & " " & GetImageAttribute(img.Attributes, ATTR_ͼ��ʱ��), "yyyy-MM-dd HH:MM:SS")
    oneImageInfo.SliceThickness = GetImageAttribute(img.Attributes, ATTR_���)
    oneImageInfo.SliceLocation = GetImageAttribute(img.Attributes, ATTR_��Ƭλ��)
    oneImageInfo.ImageOrientationPatient = GetImageAttribute(img.Attributes, ATTR_ͼ������)
    oneImageInfo.ImagePositionPatient = GetImageAttribute(img.Attributes, ATTR_ͼ��λ�ò���)
    oneImageInfo.Rows = img.sizeY
    oneImageInfo.Columns = img.sizeX
    oneImageInfo.PixelSpacing = GetImageAttribute(img.Attributes, ATTR_���ؾ���)
    '�ڵ�һ��ͼ��ǰ��׷��
    ZLSeriesInfos(iCurrentSeries).ImageInfos.Add oneImageInfo
    
    '��ʾ�������,�������е�ͼ�����viewer(index)�е�ͼ��
    If thisForm.Viewer.Count >= 2 Then
        Call thisForm.funcSwapSeries(thisForm.Viewer.Count - 1, iCurrentSeries)
    End If
        
    '������ʾ����ͼ
    subShowMiniImages thisForm
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function funShowTempImages(thisForm As frmViewer, imgs As DicomImages, iViewerIndex As Integer) As Integer
'------------------------------------------------
'���ܣ� ��ʱ�򿪲���ʾimgs�����ͼ��ֻ��ͼ����ص�ZLShowSeriesInfos�ṹ��
'������ thisForm -- ��ʾͼ��Ĵ���
'       imgs -- ��Ҫ��ʱ��ʾ��ͼ�񼯺�
'       iViewerIndex -- ��ʾͼ���Viewer��Index,���Ϊ0�� ���ʾ��Ҫ�Զ����ҿ��õ�Viewer
'���أ� �����ʱͼ���Viewer��Index
'ʱ�䣺2009-7
'------------------------------------------------
    Dim i As Integer
    Dim oneImageInfo As clsImageInfo
    Dim img As DicomImage
    
    On Error GoTo err
    
    If imgs.Count <= 0 Then Exit Function
    
    '���iViewerIndexΪ0 ������һ��ߴ���һ�����õ�Viewer
    If iViewerIndex = 0 Or iViewerIndex >= thisForm.Viewer.Count Then
        iViewerIndex = funcGetAUsableViewer(thisForm)
    End If
    
    '���ݵ�һ��ͼ�����Ϣ���޸�ZLShowSeriesInfos�ṹ
    Set ZLShowSeriesInfos(iViewerIndex).ImageInfos = Nothing
    ZLShowSeriesInfos(iViewerIndex).lngSource = 1
    ZLShowSeriesInfos(iViewerIndex).SeriesNo = GetImageAttribute(imgs(1).Attributes, ATTR_���к�)
    ZLShowSeriesInfos(iViewerIndex).SeriesUID = imgs(1).SeriesUID
    ZLShowSeriesInfos(iViewerIndex).strModality = GetImageAttribute(imgs(1).Attributes, ATTR_Ӱ�����)
    ZLShowSeriesInfos(iViewerIndex).strSaveDir = ""
    ZLShowSeriesInfos(iViewerIndex).StudyUID = imgs(1).StudyUID
    ZLShowSeriesInfos(iViewerIndex).MultiFrame = 0
    ZLShowSeriesInfos(iViewerIndex).Selected = False
    
    'ѭ����ÿһ��ͼ����ӵ�ͼ��ṹ��
    For i = imgs.Count To 1 Step -1
        Set oneImageInfo = funGetNewImageInfo
        oneImageInfo.AcquisitionTime = Format(GetImageAttribute(imgs(i).Attributes, ATTR_�ɼ�����) & " " & GetImageAttribute(imgs(i).Attributes, ATTR_�ɼ�ʱ��), "yyyy-MM-dd HH:MM:SS")
        oneImageInfo.Columns = imgs(i).sizeX
        oneImageInfo.FrameOfReferenceUID = GetImageAttribute(imgs(i).Attributes, ATTR_�ο�֡UID)
        oneImageInfo.ImageName = imgs(i).InstanceUID
        oneImageInfo.ImageNo = GetImageAttribute(imgs(i).Attributes, ATTR_ͼ���)
        oneImageInfo.ImageOrientationPatient = GetImageAttribute(imgs(i).Attributes, ATTR_ͼ������)
        oneImageInfo.ImagePositionPatient = GetImageAttribute(imgs(i).Attributes, ATTR_ͼ��λ�ò���)
        oneImageInfo.ImageTime = Format(GetImageAttribute(imgs(i).Attributes, ATTR_ͼ������) & " " & GetImageAttribute(imgs(i).Attributes, ATTR_ͼ��ʱ��), "yyyy-MM-dd HH:MM:SS")
        oneImageInfo.InstanceUID = imgs(i).InstanceUID
        oneImageInfo.PixelSpacing = GetImageAttribute(imgs(i).Attributes, ATTR_���ؾ���)
        oneImageInfo.Rows = imgs(i).sizeY
        oneImageInfo.SliceLocation = GetImageAttribute(imgs(i).Attributes, ATTR_��Ƭλ��)
        oneImageInfo.SliceThickness = GetImageAttribute(imgs(i).Attributes, ATTR_���)
        oneImageInfo.blnDisplayed = True
        '�ڵ�һ��ͼ��ǰ��׷��
        ZLShowSeriesInfos(iViewerIndex).ImageInfos.Add oneImageInfo
    Next i
    
    '�򿪲���ʾͼ��
    thisForm.Viewer(iViewerIndex).Images.Clear
    For i = 1 To imgs.Count
        thisForm.Viewer(iViewerIndex).Images.Add imgs(i)
        
        Set img = thisForm.Viewer(iViewerIndex).Images(thisForm.Viewer(iViewerIndex).Images.Count)
        img.Tag = i
        If img.Labels.Count = 0 Then
            Call subInitAImage(img, iViewerIndex, thisForm.Viewer(iViewerIndex))
        End If
    Next i
    
    '����ˢ��Viewer������Viewer������ƶ�MPR�ؽ��ߵ�ʱ�򲻻��Լ�ˢ�£��Եú��ͺ�
    thisForm.Viewer(iViewerIndex).Refresh
    
    '�жϹ������Ƿ���Ҫ��ʾ�������Ҫ����ʾ�����������ù����������ֵ����Сֵ��LarghChange��
    subDisplayScrollBar iViewerIndex, thisForm, False
    
    '�����ѡ���������е�״̬����Ŀǰ��������Ϊ��ǰ���У���subDispframeʹ��
    If thisForm.isSelectAllSerial Then thisForm.intSelectedSerial = iViewerIndex
    
    'ͼ����ʾ������Viewer�еı�ע��ͼ������½ǵ�ѡ���ǵ�
    Call subDispframe(thisForm, thisForm.Viewer(iViewerIndex))
    
    '�Զ�����ͼ���С���ж��Ƿ���ʾ�����Ľ���Ϣ,��ʾ��������ͼ���еĲ�����Ϣ
    Call subDisplayPatientInfo(thisForm.Viewer(iViewerIndex))
    
    funShowTempImages = iViewerIndex
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function funcGetAUsableViewer(thisForm As frmViewer) As Integer
'------------------------------------------------
'���ܣ� �ӵ�ǰ�����в���һ���������ͼ���Viewer
'       �����ǰ�����пյĵط����������հ״����Viewer�����Ҵ���ZLShowSeriesInfos�ṹ��
'       ���û�пհ״�����ʹ�õ�ǰ�򿪵����һ��Viewer
'������ thisForm -- ��ʾͼ��Ĵ���
'���أ� �ҵ���Viewer��Index��0��ʾû���ҵ�
'ʱ�䣺2009-7
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim intViewerIndex As Integer
    
    On Error GoTo err
    '����ͼ��Ĳ��������ҿհ׵�Viewer
    '������ִ��Viewer���пհ�Viewer��ֱ��ʹ��
    '������Viewer�������Ƿ����ڽ��������ʾ��Viewer���ܺϣ����С�ڣ��򴴽�һ���µ�Viewer
    '�ٷ����ֱ��ʹ�õ�ǰ�����һ��Viewer
    
    '�ȼ�鵱ǰViewer���Ƿ����ִ��Viewer
    For i = 1 To thisForm.Viewer.Count - 1
        If thisForm.Viewer(i).Images.Count = 0 Then
            funcGetAUsableViewer = i
            Exit Function
        End If
    Next i
    
    '�ټ��Viewer�������Ƿ����ڽ������ʾ��Viewer���ܺ�
    If thisForm.Viewer.Count - 1 < thisForm.intCountX * thisForm.intCountY Then
        '�пհ׵�Viewer�����ҿհ�Viewer��λ�ã�������һ��Viewer
        For i = 1 To thisForm.intCountY
            For j = 1 To thisForm.intCountX
                intViewerIndex = 0
                For k = 1 To ZLShowSeriesInfos.Count
                    If ZLShowSeriesInfos(k).intRow = i And ZLShowSeriesInfos(k).intCol = j Then
                        intViewerIndex = k
                    End If
                Next k
                If intViewerIndex = 0 Then Exit For
            Next j
            If intViewerIndex = 0 Then Exit For
        Next i
        If intViewerIndex = 0 Then
            '����һ��Viewer
            intViewerIndex = funcCeateAViewer(1, thisForm)
            '�ڷ�һ��Viewer
            Call subPlaceAViewer(thisForm, intViewerIndex, i, j)
            funcGetAUsableViewer = intViewerIndex
        End If
    Else
        'ֱ��ʹ�����һ��Viewer
        funcGetAUsableViewer = thisForm.Viewer.Count - 1
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function funMPR(thisForm As frmViewer, Optional blnSilent As Boolean = False) As Boolean
'------------------------------------------------
'���ܣ� �Ե�ǰ�����б�ѡ�е�������ʸ��״λ�ؽ�������ȡ��ʸ��״λ�ؽ�����ֱ�ӵ���
'       thisForm.blnInMPR ˵���������Ƿ���ͼ�����ڽ����ؽ��Ĺ�����
'������ thisForm -- ��ʾͼ��Ĵ���
'       blnSilent -- ��Ĭ����MRP������ʾ
'���أ� True--�ɹ���False---ȡ���˳�
'ʱ�䣺2009-7
'------------------------------------------------
    Dim thisViewer As DicomViewer
    Dim intMPRViewerIndex1 As Integer
    Dim intMPRViewerIndex2 As Integer
    Dim intMPRViewerIndex3 As Integer
    Dim resImage1 As New DicomImage
    Dim resImage2 As New DicomImage
    Dim i As Integer
    Dim dGlabal As New DicomGlobal
    Dim lngResult As Long
    Dim blnSortForward As Boolean       'True-���մ�λ��������False-���մ�λ��������
    
    On Error GoTo err
    
    '����ʸ��״λ�ؽ�������ȴû��ѡ���κ�ͼ�����˳�
    If thisForm.blnInMPR = False And thisForm.SelectedImage Is Nothing Then
        MsgBox "����ѡ��ͼ��Ȼ���ٽ���MPR��", vbOKOnly, "��ܰ��ʾ"
        Exit Function
    End If
    
    Set thisViewer = thisForm.Viewer(thisForm.intSelectedSerial)
    '�ȴ���ȡ���ؽ����ٴ����ؽ�
    If thisForm.blnInMPR = True Then    'ȡ���ؽ�
        
        If blnSilent = True Then       '��Ĭ������ʾ�Ƿ����MPR
            lngResult = vbNo
        Else
            '��ʾ�Ƿ񱣴��ؽ����ͼ
            lngResult = MsgBox("�Ƿ񱣴�MPR�ؽ��Ľ��ͼ��", vbYesNoCancel, gstrSysName)
            If lngResult = vbCancel Then
                funMPR = False
                Exit Function
            End If
        
        End If
        
        '�����MPR�ؽ��ı�ǣ�������Щ����������������
        thisForm.blnInMPR = False
        
        If lngResult = vbYes Then      '�����ؽ��Ľ��ͼ
            Set resImage1 = thisForm.Viewer(ZLMPRCube(2).intViewerIndex).Images(1)
            Set resImage2 = thisForm.Viewer(ZLMPRCube(3).intViewerIndex).Images(1)
            resImage1.SeriesUID = ZLMPRSeriesUID
            resImage2.SeriesUID = ZLMPRSeriesUID
            '������ͼ
            Call subSaveImage(resImage1, thisViewer.Images(1).SeriesUID)
            Call subSaveImage(resImage2, thisViewer.Images(1).SeriesUID)
            '��ͼ��׷�ӵ���Ƭվ��
            Call subOpenCurrentImage(thisForm, resImage1)
            Call subOpenCurrentImage(thisForm, resImage2)
        End If
        
        '�ָ�ԭ�����滻��Viewer�е�ͼ�����ݣ�ɾ��Ϊ��MPR������ӵ�Viewer
        '����MPR���У����滻���ˣ�����Ҫ�滻����
        If ZLMPRCube(1).blnIsMPR = False Then
            Call subMPRReFillImagesToViewer(1, thisForm)
        Else
            'ȥ��ͼ�������ʸ��״λ�ؽ����
            For i = G_INT_SYS_LABEL_MPRV To G_INT_SYS_LABEL_MPR_POINT_O
                thisForm.Viewer(ZLMPRCube(1).intViewerIndex).Images(1).Labels(i).Visible = False
            Next i
            subMPRLinenPhase thisForm.Viewer(ZLMPRCube(1).intViewerIndex), thisForm.Viewer(ZLMPRCube(1).intViewerIndex).Images(1)
        End If
        
        '����ڶ������ͼ
        Call subMPRReFillImagesToViewer(3, thisForm)
        
        '�����һ�����ͼ
        Call subMPRReFillImagesToViewer(2, thisForm)
        
        '��������иı䣬�������������в���
        If thisForm.intOldCountX <> thisForm.intCountX Or thisForm.intOldCountY <> thisForm.intCountY Then
            '�ָ�ԭ����ͼ�񲼾�
            thisForm.intCountX = thisForm.intOldCountX
            thisForm.intCountY = thisForm.intOldCountY
            Call subChangeSeriesLayout(thisForm)
        End If
        
        '��ջ������ά����
        ReDim aPixels(0)
    Else        '��ʼ�ؽ�
        ZLMPRSeriesUID = dGlabal.NewUID
        '�Ȱ������е�����ͼ�񶼼��ص�Viewer��
        Call funAddAllImages(thisViewer)
        
        '�ж��Ƿ�����ʸ��״λ�ؽ�������
        If LeagelToACRebuild(thisViewer.Images) = 1 Then
            thisForm.blnInMPR = False   '�����ؽ�״̬Ϊ�˳��ؽ�
            Exit Function
        End If
            
        '�����������۲�ģʽ�����˳���ģʽ
        If Button_miLookOrBrowse = True Then
            Call subLookOrBrowsSwitch(thisForm)
        End If
        
        '��¼ԭ�������в���,����С��2*2�����в��֣�ʸ��״λ�ؽ�ʱ��Ҫ�������޸ĳ�2*2���˳��ؽ���Ҫ�ָ�ԭ���Ĳ���
        '���ڴ���2*2�����в��֣�����ʸ��״λ�ؽ���ʱ�򣬲���Ҫ�������еĲ���
        thisForm.intOldCountX = thisForm.intCountX
        thisForm.intOldCountY = thisForm.intCountY
        '������в���С��2*2�������Ϊ2*2
        If thisForm.intCountX < 2 Then thisForm.intCountX = 2
        If thisForm.intCountY < 2 Then thisForm.intCountY = 2
        '��������иı䣬�������������в���
        If thisForm.intOldCountX <> thisForm.intCountX Or thisForm.intOldCountY <> thisForm.intCountY Then
            Call subChangeSeriesLayout(thisForm)
        End If
        
        '�ڷ����Ͻǵ�Viewer
        '������Ͻ���ͼ���򱣴�����ͼ�����û��ͼ���������ﴴ��Viewer
        intMPRViewerIndex1 = funcReplaceViewer(1, thisForm, thisViewer)
        '����thisViewer�е�ͼ����ӵ�MPRViewer1��
        If thisViewer.Index <> intMPRViewerIndex1 Then
            Call funShowTempImages(thisForm, thisViewer.Images, intMPRViewerIndex1)
            Set thisViewer = thisForm.Viewer(intMPRViewerIndex1)
        End If
        
        
        '��ʼ�ؽ�
        '��ͼ�����ؽ��ĳ�ʼ��
        '��Viewer�е�ȫ��ͼ�񣬳�ʼ��ʸ��״λ�ؽ��Ŀ����ߺͿ��Ƶ�
        Call subInitMPRLine(thisViewer)
        'ʸ��״λ�ؽ���ʼ������д����ܸ߶Ⱥ���������
        Call funcPlaneRestructInit(thisViewer, thisForm)
        
        '������һ���ؽ����ͼ
        intMPRViewerIndex2 = funcReplaceViewer(2, thisForm, Nothing)
        '���ݴ���Ŀ����ߣ���MPRViewer�ڵ�ͼ������ؽ������������ؽ��Ľ��ͼ��ʾ��ShowViewerIndexָ����Viewer��
        If funGetMPRImageAndShow(thisViewer.Images(1).Labels(G_INT_SYS_LABEL_MPRV), thisForm, thisViewer, _
                                    1, intMPRViewerIndex2, ToltalHeight, 1, True, True) = False Then
            '�ؽ������˳�MPR�ؽ�
            thisForm.blnInMPR = True
            Call funMPR(thisForm, True)
            funMPR = False
            Exit Function
        End If
                
        '�����ڶ����ؽ����ͼ
        intMPRViewerIndex3 = funcReplaceViewer(3, thisForm, Nothing)
        '���ݴ���Ŀ����ߣ���MPRViewer�ڵ�ͼ������ؽ������������ؽ��Ľ��ͼ��ʾ��ShowViewerIndexָ����Viewer��
        If funGetMPRImageAndShow(thisViewer.Images(1).Labels(G_INT_SYS_LABEL_MPRH), thisForm, thisViewer, _
                                    1, intMPRViewerIndex3, ToltalHeight, 2, True, True) = False Then
            '�ؽ������˳�MPR�ؽ�
            thisForm.blnInMPR = True
            Call funMPR(thisForm, True)
            funMPR = False
            Exit Function
        End If
                
        thisForm.intSelectedSerial = thisViewer.Index
         '����MPR���
         thisForm.blnInMPR = True
    End If
    
    funMPR = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function funMPRslope(imageViewer As DicomViewer, axialViewer As DicomViewer, _
    CoronalViewer As DicomViewer, SagittalViewer As DicomViewer, parForm As frmViewer) As Boolean
'------------------------------------------------
'���ܣ� �Ե�ǰ�����б�ѡ�е�������ʸ��״λб���ؽ�
'������ imageViewer -- ͼ�����ڵ�Viewer������������
'       axialViewer -- ��λͼ���ڵ�Viewer����б���ؽ�������
'       CoronalViewer -- �ؽ����ͼ��״λ���ڵ�Viewer����б���ؽ�������
'       SagittalViewer -- �ؽ����ͼʸ״λ���ڵ�Viewer����б���ؽ�������
'       parForm -- ͼ�����ڵ�Form
'���أ� True--�ɹ���False---ȡ���˳�
'------------------------------------------------
    
    On Error GoTo err
        
    '��ʼ�ؽ�
    '��ͼ�����ؽ��ĳ�ʼ��
    
    funMPRslope = False
    
    'ʸ��״λ�ؽ���ʼ������д����ܸ߶Ⱥ���������
    Call funcPlaneRestructInit(imageViewer, parForm)
        
    '���ݴ���Ŀ����ߣ���MPRViewer�ڵ�ͼ������ؽ������������ؽ��Ľ��ͼ��ʾ��ShowViewerIndexָ����Viewer��
    If funGetCandSImageAndShow(axialViewer.Images(1).Labels(G_INT_SYS_LABEL_MPRH), imageViewer, _
                                    axialViewer, CoronalViewer, ToltalHeight, 1, True, True) = False Then
        '�ؽ������˳�MPR�ؽ�
        Exit Function
    End If
    
    '���ݴ���Ŀ����ߣ���MPRViewer�ڵ�ͼ������ؽ������������ؽ��Ľ��ͼ��ʾ��ShowViewerIndexָ����Viewer��
    If funGetCandSImageAndShow(axialViewer.Images(1).Labels(G_INT_SYS_LABEL_MPRV), imageViewer, _
                                    axialViewer, SagittalViewer, ToltalHeight, 2, True, True) = False Then
        '�ؽ������˳�MPR�ؽ�
        Exit Function
    End If
        
    funMPRslope = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function funAddAllImages(thisViewer As DicomViewer) As Boolean
'------------------------------------------------
'���ܣ� ���������е�ͼ�񶼼��ص�thisViewer��
'������ thisViewer -- ��Ҫ���ص�Viewer
'���أ� ��
'ʱ�䣺2009-7
'------------------------------------------------
    Dim i As Integer
    Dim intViewerIndex As Integer
    
    If thisViewer Is Nothing Then Exit Function
    
    On Error GoTo err
    intViewerIndex = thisViewer.Index
    For i = 1 To ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
        If ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).blnDisplayed = False Then
            Call funcAddAImageA(thisViewer, i)
        End If
    Next i
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub subChangeSeriesLayout(thisForm As frmViewer)
'------------------------------------------------
'���ܣ����ݴ������intCountX��intCountY���л����в���
'������ thisForm --- ��Ƭ����
'���أ��� ��ֱ���л����в���
'ʱ�䣺2009-7
'------------------------------------------------
    '�л����в��ֵ�ʱ��ֻ���Ѿ����ؽ���Viewer�е�ͼ����ʾ����������ʾ����ͼ��
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim blnLoadOver As Boolean
    Dim iCurrentViewerIndex As Integer
    Dim intSeriesIndex As Integer
    Dim vTemp As DicomViewer
    Dim blnFound As Boolean
    
    On Error GoTo err
    '�ȶԱ�ԭ�����в��ָ���ǰ���в�����Viewer�����������Viwer�����ˣ���װ���µ�Viewer��
    '��װ�ص�Viewer�У����ΰ�������ͼ��˳����ʾ����ʣ�������
    '���Viewer�����ˣ���ж�ض����Viewer
    
    '�������в��֣��ڷŷָ���
    Call subShowSpliter(thisForm)
    
    iCurrentViewerIndex = 0
    For i = 1 To thisForm.intCountY
        For j = 1 To thisForm.intCountX
            iCurrentViewerIndex = iCurrentViewerIndex + 1
            
            If iCurrentViewerIndex >= thisForm.Viewer.Count Then
                '����һ��Viewer
                '��ѯ�´�����Viewer׼��װ���ĸ�����
                intSeriesIndex = 0
                For k = 1 To ZLSeriesInfos.Count
                    For Each vTemp In thisForm.Viewer
                        If vTemp.Tag = k Then
                            blnFound = True
                            Exit For
                        Else
                            blnFound = False
                        End If
                    Next
                    If blnFound = False Then
                        intSeriesIndex = k
                        Exit For
                    End If
                Next k
                If intSeriesIndex = 0 Then
                    '˵���������ж��Ѿ�װ�ڽ����ˣ������˳�װ�����е�ѭ����
                    blnLoadOver = True
                    Exit For
                End If
                iCurrentViewerIndex = funcCeateAViewer(intSeriesIndex, thisForm)
            End If
            '�ڷ����Viewer,Viewer����ͼ�񣬲���Ҫ�ڷţ�������һ���յ�Viewer������Ҫ�ڷ�
            If thisForm.Viewer(iCurrentViewerIndex).Images.Count <> 0 Then
                Call subPlaceAViewer(thisForm, iCurrentViewerIndex, i, j)
            End If
        Next j
        If blnLoadOver = True Then
            Exit For
        End If
    Next i

    'ж�ض����Viewer
    If iCurrentViewerIndex < thisForm.Viewer.Count Then
        While thisForm.Viewer.Count > 1 And thisForm.Viewer.Count - 1 > iCurrentViewerIndex
            Call subUnloadLastViewer(thisForm)
        Wend
        If thisForm.intSelectedSerial >= thisForm.Viewer.Count Then thisForm.intSelectedSerial = 0
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function funcReplaceViewer(intType As Integer, thisForm As frmViewer, thisViewer As DicomViewer) As Integer
'------------------------------------------------
'���ܣ� ��ָ��λ�ô���Viewer
'       ���ָ��λ����ͼ���򱣴�����ͼ�����û��ͼ���������ﴴ��Viewer
'       �滻Viewer�����������
'           1����MPR�����滻��Viewer������ͼ���Viewer��
'               1) ����ǿ�Viewer��ֱ�Ӵ���һ��Viewer,��ZLSeriesInfos�ж�ȡͼ��
'               2) �������ͼ���Viewer,��Viewer�е�ͼ����MPR���У����ԭ�����б��浽ZLMPRImages(1)�У��ڴ�Viewer����ʾMPR���У���ZLSeriesInfos�ж�ȡͼ��
'               3) �������ͼ���Viewer����VIewer�е�ͼ����MPR���У�������ǣ���װ�������е�ȫ��ͼ��
'           2����MPR��������滻��Viewer������ͼ���Viewer��
'               1) ����ǿ�Viewer��ֱ�Ӵ���һ��Viewer�����ɽ��ͼ����ӵ�Viewer�С�
'               2) �������ͼ���Viewer�����ԭ�����б��浽ZLMPRImages(2,3)�У��ڴ�Viewer����ʾMPR���ͼ��
'       �洢MPR֮ǰͼ��Ľṹ�����������ݣ�
'           1��ZLShowSeriesInfos    --- ԭ�е�ZLShowSeriesInfos�ṹ
'           2��Images               --- ԭ��Viewer���Ѿ����ص�ͼ�񣬷���ָ�ͼ���еı�ע�����������ŵ���Ϣ
'           3��blnIsMPR             --- �Ƿ�ǰ��MPR�����У�����ǣ��ָ���ʱ�򣬲���Ҫ�滻�����е����ݡ�
'������ intType     --- �������� 1--MPR����λ�ã�1��1����2--MPR����������ߣ�λ�ã�1��2����3--MPR������к��ߣ�λ�ã�2��1��
'       thisForm    --- ��ʾͼ��Ĵ���
'       thisViewer  --- ����MPR����������
'���أ��ɹ�MPRViewer��Index��ʧ��=0
'ʱ�䣺2009-7
'------------------------------------------------
    Dim i As Integer
    Dim blnNewViewer As Boolean
    Dim intViewerIndex As Integer
    Dim intRow As Integer, intCol As Integer
    Dim oneImageInfo As clsImageInfo
    Dim oneSeriesInfo As clsSeriesInfo
    Dim MPRViewer As DicomViewer    '�ڷźõ�Viewer
    
    On Error GoTo err
    
    funcReplaceViewer = 0
    If intType < 1 Or intType > 3 Then Exit Function
    If thisForm Is Nothing Then Exit Function
    
    '����ǵ�һ�У���һ�У�����Ҫȷ��thisViewer����
    If intType = 1 And thisViewer Is Nothing Then Exit Function
    
    '��ʼ��ZLMPRCube
    ZLMPRCube(intType).blnIsMPR = False
    ZLMPRCube(intType).Images.Clear
    Set ZLMPRCube(intType).ZLShowSeriesInfos = Nothing
    ZLMPRCube(intType).intViewerIndex = 0
    
    blnNewViewer = True
    intRow = 1
    intCol = 1
    If intType = 2 Then
        intCol = 2
    ElseIf intType = 3 Then
        intRow = 2
    End If
    '���ж�ָ��λ���Ƿ���Viewer
    For i = 1 To ZLShowSeriesInfos.Count
        If ZLShowSeriesInfos(i).intRow = intRow And ZLShowSeriesInfos(i).intCol = intCol And ZLShowSeriesInfos(i).ImageInfos.Count <> 0 Then
            blnNewViewer = False
            intViewerIndex = i
            Exit For
        End If
    Next i
    
    '���ָ��λ����Viewer,����Ҫ�������Viewer�е�ͼ��
    If blnNewViewer = False Then
        '��MPR���У����������Ͻǵ�Viewer����MPR���У�����Ҫ����Viewer������
        If Not thisViewer Is Nothing Then
            If intType = 1 And thisViewer.Index = intViewerIndex Then
                ZLMPRCube(intType).blnIsMPR = True
                Set MPRViewer = thisViewer
            End If
        End If
        
        If ZLMPRCube(intType).blnIsMPR = False Then '��Ҫ����ԭͼ��ZLShowSeriesInfos�ṹ
            '����ͼ��
            Set MPRViewer = thisForm.Viewer(intViewerIndex)
            For i = 1 To MPRViewer.Images.Count
                ZLMPRCube(intType).Images.Add MPRViewer.Images(i)
            Next i
            
            '����ZLShowSeriesInfos�ṹ
            Set oneSeriesInfo = funGetNewSeriesInfo
            Call funCopySeriesInfo(ZLShowSeriesInfos(intViewerIndex), oneSeriesInfo)
            Set ZLMPRCube(intType).ZLShowSeriesInfos = oneSeriesInfo
            Set ZLMPRCube(intType).ZLShowSeriesInfos.ImageInfos = Nothing
            
            '����ImageInfos����Ϣ
            For i = 1 To ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
                Set oneImageInfo = funGetNewImageInfo
                Call funCopyImageInfo(ZLShowSeriesInfos(intViewerIndex).ImageInfos(i), oneImageInfo)
                ZLMPRCube(intType).ZLShowSeriesInfos.ImageInfos.Add oneImageInfo
            Next i
        End If
    Else    'ָ��λ��û��Viewer���򴴽�һ��Viewer�����Ұ����Viewer�ڷŵ�ָ��λ��
        '����һ��Viewer
        intViewerIndex = funcCeateAViewer(1, thisForm)
        '�ڷ�һ��Viewer
        Call subPlaceAViewer(thisForm, intViewerIndex, intRow, intCol)
        Set MPRViewer = thisForm.Viewer(intViewerIndex)
    End If
    ZLMPRCube(intType).intViewerIndex = intViewerIndex
    funcReplaceViewer = MPRViewer.Index
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function funGetMPRResultImage(la As DicomLabel, thisViewer As DicomViewer, intToltalHeight As Integer, _
    intType As Integer) As DicomImage
'------------------------------------------------
'���ܣ� ���ݴ���Ŀ����ߣ���thisViewer�ڵ�ͼ������ؽ������������ؽ��Ľ��ͼ
'������ la          --- �����ؽ��Ŀ�����
'       thisViewer  --- �����ؽ���ͼ�����ڵ�Viewer
'       intToltalHeight  --- �����ؽ�ͼ������ӵ�����߶�
'       intType     --- ͼ�����ͣ�1--���ߣ�2--���ߣ�ֱ����д��ͼ�����
'���أ��ؽ����ͼ�����ʧ�ܣ�����nothing
'ʱ�䣺2009-7
'------------------------------------------------
    Dim LineLong() As POINTAPI      '����MPR��������ÿ��������������
    Dim iPointsCount As Long        '��¼MPR�������е��������
    Dim iImagesCount As Long        '��¼Viewer��ͼ���������
    Dim lines() As Integer          '����ͼ��Ҷ�ֵ�Ķ�ά����
    Dim NewLines() As Integer       '�����ؽ���ͼ��Ҷ�ֵ�Ķ�ά����
    Dim i As Long, j As Long
    Dim resImage As DicomImage      '���ͼ��
    Dim v As Variant
    
    Set funGetMPRResultImage = Nothing
    
    On Error GoTo err
    
    If thisViewer.Images.Count <= 0 Then Exit Function
    
    '��ȡ��ע����ÿ���������λ������
    Call subGetArray(la, thisViewer.Images(1), LineLong)
    iPointsCount = UBound(LineLong)
    iImagesCount = thisViewer.Images.Count
    
    '���¶���ԭͼͼ��Ҷ�ֵ��ά����
    ReDim lines(iPointsCount, iImagesCount) As Integer
    '���¶����ؽ���Ҷ�ֵ��ά����
    ReDim NewLines(iPointsCount, intToltalHeight) As Integer
    
    '����MPR���������ڵ�����飬��ȡͼ���ĻҶ�ֵ
    If SafeArrayGetDim(aPixels) = 0 Then
        'MPR�Ļ�����ά����ά��=0��˵�������ڴ���ɣ���ֱ��ʹ��ͼ���������ؽ���ͼ��Խ�࣬�ؽ�Խ��
        For i = 1 To iImagesCount
            v = thisViewer.Images(i).Pixels
            For j = 1 To iPointsCount
                lines(j, i) = v(LineLong(j).x, LineLong(j).y, 1)
            Next j
        Next i
    Else
        'ʹ����ά������MPR�ؽ���ÿ���ؽ��ٶ���1����
        For i = 1 To iImagesCount
            For j = 1 To iPointsCount
                lines(j, i) = aPixels(LineLong(j).x, LineLong(j).y, i)
            Next j
        Next i
    End If
    
    '���ݲ�񽫲ɼ�������ֱ�߲�ֵ��һ������ͼ��ͬʱ��ƽ������
    Call subACRebuild(lines, NewLines)
    '������ͼ��
    Set resImage = thisViewer.Images(1).SubImage(0, 0, thisViewer.Images(1).sizeX, thisViewer.Images(1).sizeY, 1, 1)
    
    'ɾ��һЩ���õ�λ������
    resImage.Attributes.Remove &H18, &H50
    resImage.Attributes.Remove &H18, &H1110
    resImage.Attributes.Remove &H18, &H1111
    resImage.Attributes.Remove &H18, &H1120     'Tilt
    resImage.Attributes.Remove &H18, &H1140     'Rotation Direction
    resImage.Attributes.Remove &H18, &H5100     'Patient Position
    resImage.Attributes.Remove &H20, &H32       'Image Position(Patient)
    resImage.Attributes.Remove &H20, &H37       'Image Orientation (Patient)
    resImage.Attributes.Remove &H20, &H1041     'Slice Location
    
    '���ý��ͼ������
    resImage.Attributes.Add &H28, &H10, intToltalHeight
    resImage.Attributes.Add &H28, &H11, iPointsCount
    If intType = 1 Then
        resImage.Attributes.Add &H20, &H11, LineLong(1).y
    Else
        resImage.Attributes.Add &H20, &H11, LineLong(1).x
    End If
    resImage.Attributes.Add &H20, &H13, intType
    resImage.Pixels = NewLines
    resImage.width = thisViewer.Images(1).width
    resImage.Level = thisViewer.Images(1).Level
    
    '���ؽ��ͼ
    Set funGetMPRResultImage = resImage
    
    Exit Function
err:
    If err.Number = 61706 Or err.Number = -2147417848 Then
        err.Description = "�ڴ治�㣬�޷�����MPR�ؽ�����������������������ڴ�����ԡ�"
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set funGetMPRResultImage = Nothing
End Function

Public Function funGetMPRImageAndShow(la As DicomLabel, thisForm As frmViewer, MPRViewer As DicomViewer, _
    MPRImageIndex As Integer, ShowViewerIndex As Integer, intToltalHeight As Integer, intType As Integer, _
    blnFirst As Boolean, blnChangeLa As Boolean) As Boolean
'------------------------------------------------
'���ܣ� ���ݴ���Ŀ����ߣ���MPRViewer�ڵ�ͼ������ؽ������������ؽ��Ľ��ͼ��ʾ��ShowViewerIndexָ����Viewer��
'������ la          --- �����ؽ��Ŀ�����
'       thisForm    --- ��ʾͼ��Ĵ���
'       MPRViewer   --- �ؽ�ͼ�����ڵ�Viewer
'       MPRImageIndex --- �ؽ�ͼ����Viewer�е�Index
'       ShowViewerIndex --- �ؽ����ͼ����ʾ��Viewer��Index
'       intToltalHeight --- ����߶�
'       intType     --- ͼ�����ͣ�1--���ߣ�2--���ߣ�ֱ����д��ͼ�����
'       blnFirst    --- �Ƿ��һ�ε��ã�����ǵ�һ�ε��ã��򲻼�¼ԭ���Ĵ���λ��ͼ��״̬
'       blnChangeLa --- �Ƿ�ı�MPR�����ߣ����ֻ�ı�ͼ��û�������ߣ�����Ҫ��������MPRͼ��ֻ�ػ���Ӧ�߼���
'���أ��ؽ����ͼ�����ʧ�ܣ�False
'ʱ�䣺2009-7
'------------------------------------------------
    
    Dim resImage As DicomImage
    Dim resImages As New DicomImages
    Dim imgOld As DicomImage
    Dim imgNew As DicomImage
    Dim thisImage As DicomImage
    Dim dblZoom As Double
    Dim lngScrollX As Long
    Dim lngScrollY As Long
    Dim lngWWidth As Long
    Dim lngWLevel As Long
    Dim blnStretchToFit As Boolean
    Dim dblScale As Double
    
    On Error GoTo err
    
    If ShowViewerIndex >= thisForm.Viewer.Count Then
        funGetMPRImageAndShow = False
        Exit Function
    End If
    
    If blnChangeLa = True Then   '�ƶ���MPR�����ߣ�Ҫ�����µĽ��ͼ
        '���ݴ���Ŀ����ߣ���һ��Viewer�ڵ�ͼ������ؽ����������ؽ����ͼ
        Set resImage = funGetMPRResultImage(la, MPRViewer, intToltalHeight, intType)
    Else
        Set resImage = thisForm.Viewer(ShowViewerIndex).Images(1)
    End If
    
    '��ʾ���ͼ
    If resImage Is Nothing Then
        funGetMPRImageAndShow = False
        Exit Function
    Else
    
        Set imgOld = Nothing
        If blnChangeLa = True Then  '�ı���MPR�����ߣ������µĽ��ͼ
            resImages.Clear
            resImages.Add resImage
            '��¼ԭ����ͼ��״̬
            If thisForm.Viewer(ShowViewerIndex).Images.Count > 0 And blnFirst = False Then
                Set imgOld = thisForm.Viewer(ShowViewerIndex).Images(1)
                blnStretchToFit = imgOld.StretchToFit
                dblZoom = imgOld.ActualZoom
                lngScrollX = imgOld.ActualScrollX
                lngScrollY = imgOld.ActualScrollY
                lngWWidth = imgOld.width
                lngWLevel = imgOld.Level
            End If
            Call funShowTempImages(thisForm, resImages, ShowViewerIndex)
        End If
        
        If thisForm.Viewer(ShowViewerIndex).Images.Count > 0 Then
            Set imgNew = thisForm.Viewer(ShowViewerIndex).Images(1)
            imgNew.Refresh False
        End If
        
        '�ָ�ԭ��ͼ���״̬
        If Not imgOld Is Nothing And Not imgNew Is Nothing Then
            imgNew.StretchToFit = blnStretchToFit
            imgNew.Zoom = dblZoom
            imgNew.ScrollX = lngScrollX
            imgNew.ScrollY = lngScrollY
            imgNew.width = lngWWidth
            imgNew.Level = lngWLevel
        End If
        
        '��������ȷ���Ƿ���ʾMPR������
        If blnShowMPRLine = True And Not imgNew Is Nothing Then
            '����λ��ʸ��״λͶӰ��
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).top = MPRImageIndex / MPRViewer.Images.Count * imgNew.sizeY
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).left = 0
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).width = imgNew.sizeX
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).height = 0
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).Visible = True
            
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).top = 0
            
            Set thisImage = MPRViewer.Images(MPRImageIndex)
            '�������ɿ����ߵĹ�����ȷ�����ĵ���ͶӰ���е�λ��
            If Abs(la.width) > Abs(la.height) Then  '���ɺ���
                If la.width < 0 Then
                    dblScale = 1 - Abs((thisImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left - la.left) / la.width)
                ElseIf la.width > 0 Then
                    dblScale = Abs((thisImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left - la.left) / la.width)
                End If
            Else    '��������
                If la.height < 0 Then
                    dblScale = 1 - Abs((thisImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top - la.top) / la.height)
                ElseIf la.height > 0 Then
                    dblScale = Abs((thisImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top - la.top) / la.height)
                End If
            End If

            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).left = dblScale * imgNew.sizeX
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).height = imgNew.sizeY
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).width = 0
            
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).Visible = True
        End If
        
        funGetMPRImageAndShow = True
    End If
        
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub subMPRReFillImagesToViewer(intType As Integer, thisForm As frmViewer)
'------------------------------------------------
'���ܣ� ��MPR�б���ʱ����������ͼ��ָ���ԭ����Viewer��
'������ intType     --- MPR���������ͣ�1--MPR���У�2--MPR���߽�����У�3 -- MPR���߽������
'       thisForm    --- ��ʾͼ��Ĵ���
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim intViewerIndex As Integer
    Dim oneImageInfo As clsImageInfo
    Dim i As Integer
    
    If intType < 1 Or intType > 3 Then Exit Sub
    'û��ͼ����Ҫɾ�����Viewer
    
    On Error GoTo err
    
    intViewerIndex = ZLMPRCube(intType).intViewerIndex
    
    If thisForm.Viewer.Count <= intViewerIndex Then Exit Sub
    If intViewerIndex = 0 Then Exit Sub
    
    thisForm.Viewer(intViewerIndex).Images.Clear
    
    '���ͼ������Ϊ0����ʾ���λ��ԭ����û��Viewer�ģ���Ҫж�����Viewer
    If ZLMPRCube(intType).Images.Count = 0 Then
        Call subUnloadViewer(intViewerIndex, thisForm)
    Else
        '�������ʱͼ��ķ�ʽ���ָ�ͼ��
        
        Call funShowTempImages(thisForm, ZLMPRCube(intType).Images, intViewerIndex)
        '����ZLShowSeriesInfos��Ϣ
        Call funCopySeriesInfo(ZLMPRCube(intType).ZLShowSeriesInfos, ZLShowSeriesInfos(intViewerIndex))
        Set ZLShowSeriesInfos(intViewerIndex).ImageInfos = Nothing
        
        '����ImageInfos����Ϣ
        For i = 1 To ZLMPRCube(intType).ZLShowSeriesInfos.ImageInfos.Count
            Set oneImageInfo = funGetNewImageInfo
            Call funCopyImageInfo(ZLMPRCube(intType).ZLShowSeriesInfos.ImageInfos(i), oneImageInfo)
            ZLShowSeriesInfos(intViewerIndex).ImageInfos.Add oneImageInfo
        Next i
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subUnloadViewer(ByVal intViwerIndex As Integer, thisForm As frmViewer)
'------------------------------------------------
'���ܣ� �ر����У����Viewer�����һ���Ļ�����ж�����Viewer������ص�����
'       ���Viewer���м��ĳ��Viewer�������Viewer������������е�ͼ�񣬵��ǲ�ж��
'������ intViwerIndex    --- Ҫж�ص�Viewer������
'       thisForm    --- ��ʾͼ��Ĵ���
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim oneSeriesInfo As clsSeriesInfo
    
    On Error GoTo err
    
    If intViwerIndex > thisForm.Viewer.Count Then Exit Sub
    
    '��ʼ��ͨ�õĹ�������
    thisForm.intSelectedSerial = 0
    thisForm.oldSelectedImageIndex = 0
    thisForm.oldSelectedSerial = 0
    Set thisForm.SelectedImage = Nothing
    thisForm.SelectedImageIndex = 0
    thisForm.txtText.Visible = False
    Set thisForm.SelectedLabel = Nothing
    
    '�ж����Viewer�Ƿ����һ������������һ��Viewer����ͬʱɾ����Ӧ��ZLShowSeriesInfos�ṹ
    '����������һ��Viewer����ֻ���ZLShowSeriesInfos�ṹ�е�ͼ�����Ϣ
    If thisForm.Viewer.Count - 1 = intViwerIndex Then
        '���һ��Viewer��ж��Viewer��������ز���
        Call subUnloadLastViewer(thisForm)
    Else
        thisForm.MSFViewer.TextMatrix(intViwerIndex, 1) = False
        Set oneSeriesInfo = funGetNewSeriesInfo
        Call funCopySeriesInfo(oneSeriesInfo, ZLShowSeriesInfos(intViwerIndex))
        Set ZLShowSeriesInfos(intViwerIndex).ImageInfos = Nothing
        '���Viewer�е�ͼ��
        thisForm.Viewer(intViwerIndex).Images.Clear
        thisForm.VScro(intViwerIndex).Visible = False
        'Viewer��λ�úͿɼ�����Ҫ�ı�
        thisForm.Viewer(intViwerIndex).Visible = False  '�Ͳ��ᴥ��Viewer���¼���
        thisForm.Viewer(intViwerIndex).Tag = 0
        thisForm.Viewer(intViwerIndex).left = 1
        thisForm.Viewer(intViwerIndex).top = 1
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subSeriesInPhase(intViewerIndex As Integer, thisForm As frmViewer, img As DicomImage, intType As Integer)
'------------------------------------------------
'���ܣ� ��ѡ�еĶ�����У���֤���е�ͼ������ͬ��
'������ intViewerIndex  --- ��Ҫͬ������������
'       thisForm    --- ��ʾͼ��Ĵ���
'       img         --- ����ͬ���ı�׼ͼ��
'       intType     --- ͬ�����ͣ��궨��
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim vTemp As DicomViewer
    Dim i As Integer
    
    '���������ͼ������ͬ����״̬�����˳�
    If Button_miImageInPhase = False Then Exit Sub
    
    On Error GoTo err
    
    For Each vTemp In thisForm.Viewer
        If vTemp.Visible = True Then
            If (ZLShowSeriesInfos(intViewerIndex).Selected = True And ZLShowSeriesInfos(vTemp.Index).Selected = True) Or vTemp.Index = intViewerIndex Then
                For i = 1 To vTemp.Images.Count
                    Call subImageInPhase(vTemp.Images(i), img, intType)
                Next i
                vTemp.Refresh
            End If
        End If
    Next
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subImageInPhase(img As DicomImage, SampleImg As DicomImage, intType As Integer)
'------------------------------------------------
'���ܣ� ����ͼ��֮��״̬ͬ��
'������ Img         --- ��Ҫ����ͬ����ͼ��
'       SampleImg   --- ����ͬ���ı�׼ͼ��
'       intType     --- ͬ�����ͣ��궨��
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    
    On Error GoTo err
    
    Select Case intType
    Case IMG_SYN_All                'ȫ��ͬ��
        img.width = SampleImg.width
        img.Level = SampleImg.Level
        img.StretchToFit = SampleImg.StretchToFit
        img.ScrollX = SampleImg.ScrollX
        img.ScrollY = SampleImg.ScrollY
        img.Zoom = SampleImg.Zoom
        If img.Labels.Count >= G_INT_SYS_LABEL_WWWL Then
            img.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & img.width & "-L:" & img.Level
        End If
        img.FlipState = SampleImg.FlipState
        img.RotateState = SampleImg.RotateState
        img.FilterLength = SampleImg.FilterLength
        img.UnsharpEnhancement = SampleImg.UnsharpEnhancement
        img.UnsharpLength = SampleImg.UnsharpLength
        If img.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
            If img.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then  '���±�ߵ�λ
                Call UpdateRuler(img, True)
            End If
        End If
    Case IMG_SYN_WINDOW             '����ͬ��
        img.width = SampleImg.width
        img.Level = SampleImg.Level
        If img.Labels.Count >= G_INT_SYS_LABEL_WWWL Then
            img.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & img.width & "-L:" & img.Level
        End If
    Case IMG_SYN_ZOOMPAN            '���š�����ͬ��
        img.StretchToFit = SampleImg.StretchToFit
        img.Zoom = SampleImg.Zoom
        img.ScrollX = SampleImg.ScrollX
        img.ScrollY = SampleImg.ScrollY
        If img.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
            If img.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then  '���±�ߵ�λ
                Call UpdateRuler(img, True)
            End If
        End If
    Case IMG_SYN_ROTATE             '��תͬ��
        img.RotateState = SampleImg.RotateState
    Case IMG_SYN_FLIP               '����ͬ��
        img.FlipState = SampleImg.FlipState
        img.RotateState = SampleImg.RotateState
    Case IMG_SYN_FILTER             '�˾�ͬ��
        img.FilterLength = SampleImg.FilterLength
        img.UnsharpEnhancement = SampleImg.UnsharpEnhancement
        img.UnsharpLength = SampleImg.UnsharpLength
    End Select
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subUnloadLastViewer(thisForm As frmViewer)
'------------------------------------------------
'���ܣ� ж�ش����е����һ��Viewer��������صĹ�������ZLShowSeriesInfos��MSFViewer
'������ thisForm    --- ��ʾͼ��Ĵ���
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    On Error GoTo err
    
    Call Unload(thisForm.Viewer(thisForm.Viewer.Count - 1))
    Call Unload(thisForm.VScro(thisForm.VScro.Count - 1))
    '����ZLShowSeriesInfos
    ZLShowSeriesInfos.Remove ZLShowSeriesInfos.Count
    '����MSF�ṹ
    thisForm.MSFViewer.Rows = thisForm.MSFViewer.Rows - 1
        
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subScaleViewer(thisViewer As DicomViewer, img As DicomImage, lngOldWidth As Long, lngOldHeight As Long)
'------------------------------------------------
'���ܣ� ��Viewer�Ŀ�Ⱥ͸߶ȸı�󣬶�һ��Viewer��StretchToFit=False��ͼ��λ�ú����Ž�������
'������ thisViewer  --- ���ͼ���Viewer
'       img         --- ��Ҫ������ͼ��
'       lngOldWidth --- Viewerԭ���Ŀ��
'       lngOldHeight--- Viewerԭ���ĸ߶�
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim i As Integer
    
    If thisViewer.Images.Count = 0 Then Exit Sub
    
    On Error GoTo err
    
    '�����е�һ��ͼ�����λ������
    If img.StretchToFit = False Then
        Call subScaleImage(img, thisViewer, lngOldWidth, lngOldHeight)
        
        '������������������Viewer�е�����ͼ��������
        For i = 1 To thisViewer.Images.Count
            thisViewer.Images(i).StretchToFit = False
            thisViewer.Images(i).Zoom = img.ActualZoom
            thisViewer.Images(i).ScrollX = img.ActualScrollX
            thisViewer.Images(i).ScrollY = img.ActualScrollY
        Next i
        '�������Ľ����¼��ZLShowSeriesInfos�ṹ��
        ZLShowSeriesInfos(thisViewer.Index).ScrollX = img.ActualScrollX
        ZLShowSeriesInfos(thisViewer.Index).ScrollY = img.ActualScrollY
        ZLShowSeriesInfos(thisViewer.Index).Zoom = img.ActualZoom
    End If
     
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function FunLogIn(str���� As String) As String
'���ܣ��Գ������ע�ᣬ���ע��ɹ����򷵻�ע��ʱ��
'������ str���� ---'��ע������ʹ�õ���������
'����ֵ��ע��ɹ�ע�����ڣ�ע��ʧ�ܷ��ؿ�

    Dim intNUM As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP��ַ As String         '��Ҫע���IP��ַ
    
    On Error GoTo err
        
    strIP��ַ = OS.IP
    
    
    '��ע��������ȡ��Ȩ��������-1--�����ƣ�0--��ֹ��X��X>0��--������������
    
    If str���� = LOGIN_TYPE_ҽ����Ƭվ Then
        intNUM = gintҽ����Ƭվ����
    ElseIf str���� = LOGIN_TYPE_��Ƭ��ӡ�� Then
        intNUM = gint��Ƭ��ӡ��
    Else
        intNUM = 0
    End If
    
    
    'intNUM >0 ,����ù���ע�����
    If intNUM > 0 Then  '����������
        strSQL = "Zl_Ӱ�������¼_Update('" & strIP��ַ & "','" & str���� & "'," & intNUM & ")"
        zlDatabase.ExecuteProcedure strSQL, "ע��" & str����
        '���ע���Ƿ�ɹ�
        strSQL = "Select ����ʱ��,IP��ַ from Ӱ�������¼ where  ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ��", str����)
        If rsTemp.RecordCount <= intNUM Then
            rsTemp.Filter = "IP��ַ='" & strIP��ַ & "'"
            If rsTemp.RecordCount = 1 Then  'ע��ɹ�
                FunLogIn = rsTemp!����ʱ��
                Exit Function
            End If
        End If
    ElseIf intNUM = -1 Then '������
        FunLogIn = Now
        Exit Function
    Else    '=0����������ֵ����ֹ�������κδ�����������ʾ
        
    End If
    'ע��ʧ�ܣ�����������ԭ��
    '1��ע���������������ɵ��������޷�ע��IP��ַ
    '2��ֱ��ͨ��SQL����������IP��ַ�����±��еļ�¼��������������ɵ�����
    Call MsgBox("�򿪵�" & str���� & "�������������������" & intNUM & "�������������Ӧ����ϵ��", vbOKOnly, gstrSysName)
    FunLogIn = ""
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function FunLogOut(str���� As String, str����ʱ�� As String) As Boolean
'���ܣ��˳������ʱ�򣬼������Ƿ�Ϸ�ע�������������ͨ�����������ֶζ�ʱɾ����Ӱ�������¼�����еļ�¼��
'������ str���� ---'��ע������ʹ�õ���������
'       str����ʱ�� --- ע�Ṥ��վʱ���ص�ʱ��
'����ֵ���Ϸ�ע��True���Ƿ�������False
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP��ַ As String         '��Ҫע���IP��ַ
    Dim intNUM As Integer
    
    On Error GoTo err
    strIP��ַ = OS.IP
    
    '����ʱ��Ϊ�գ���ʾע��ʧ�ܣ�û����������������˳���ʱ���ټ�����ݿ�
    If str����ʱ�� = "" Then
        FunLogOut = True
        Exit Function
    End If
    
    '��ע��������ȡ��Ȩ��������-1--�����ƣ�0--��ֹ��X��X>0��--������������
    If str���� = LOGIN_TYPE_ҽ����Ƭվ Then
        intNUM = gintҽ����Ƭվ����
    ElseIf str���� = LOGIN_TYPE_��Ƭ��ӡ�� Then
        intNUM = gint��Ƭ��ӡ��
    Else
        intNUM = 0
    End If
    
    If intNUM > 0 Then '������������
        strSQL = "Select ����ʱ�� from Ӱ�������¼ where IP��ַ=[1] and ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ��", strIP��ַ, str����)
        If rsTemp.EOF = False Then
            FunLogOut = True
        Else
            '�Ա�����ʱ������ݿ��ʱ�䣬�������ͬһ�죬˵����ǰһ�쿪�������ע����Ϣ��ɾ���ˣ�
            '���������Ϊ�ǺϷ�ע��
            strSQL = "Select sysdate from dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���ݿ�ʱ��")
            If Format(rsTemp!sysdate, "yyyy-mm-dd") <> Format(str����ʱ��, "yyyy-mm-dd") Then
                FunLogOut = True
            Else
                FunLogOut = False
            End If
        End If
    ElseIf intNUM = -1 Then '������
        FunLogOut = True
    Else    '=0����������ֵ����ֹ
        FunLogOut = False
    End If
    
    If FunLogOut = False Then
        Call MsgBox("�򿪵�" & str���� & "�������������������" & intNUM & "�������������Ӧ����ϵ��", vbOKOnly, gstrSysName)
    End If
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function getLicenseCount(strLicenseName As String) As Integer
'��ȡ��Ȩ������,����Ȩ�����޸ĺ����޸��������
'������ strLicenseName --- ��Ȩ����
    
    Dim strLiceseCount As String
    
    On Error GoTo err
    
    strLiceseCount = zl9ComLib.zlRegInfo(strLicenseName)
    If strLiceseCount = "" Then '������
        getLicenseCount = -1
    ElseIf Val(strLiceseCount) > 0 Then '������������
        getLicenseCount = Val(strLiceseCount)
    Else '��ֹ
        getLicenseCount = 0
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadImage(strPath As String, Optional blnSilent As Boolean = False) As DicomImage
'���ܣ���ȡһ���ļ�������DICOMͼ��
'������ strPath -- �ļ�·��
'       blnSilent -- ����ʾ
'���أ�DICOMͼ��
    Dim imgs As New DicomImages
    Dim img As New DicomImage
    
    On Error Resume Next
    err.Clear
    imgs.Clear
    Set img = imgs.ReadFile(strPath)
    If err <> 0 Then        '��ȡʧ�ܣ�˵������DICOM�ļ�
        err.Clear
        img.FileImport strPath, ""
        If err <> 0 Then    '����ʧ�ܣ�˵���ļ�����BMP��JPG��AVI��ʽ�ġ�
            If blnSilent = False Then
                MsgBox "�ļ�" & strPath & "���ܴ򿪣�", vbInformation, gstrSysName
            End If
            Debug.Print "�ļ�" & strPath & "���ܴ򿪣�"
            Set ReadImage = Nothing
            Exit Function
        End If
    End If
    Set ReadImage = img
End Function

Private Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function

Private Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function

Public Sub PrintFilmBeep(intTrack As Integer)
'------------------------------------------------
'���ܣ� ��Ƭ��ӡʱ����ʾ����
'������ intTrack --- ������� 1-���ͼ��2-��ӡ��Ƭ
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    If blnPrintFilmBeep Then
        If intTrack = 1 Then
            Call Beep(BEEP_Do0, 100)
            Call Beep(BEEP_Re, 100)
            Call Beep(BEEP_Mi, 100)
        Else
            Call Beep(BEEP_Do0, 150)
            Call Beep(BEEP_Mi, 150)
            Call Beep(BEEP_Sol, 150)
            Call Beep(BEEP_Do1, 150)
        End If
    End If
    
    Exit Sub
err:
    '����󲻴���
End Sub

Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'------------------------------------------------
'���ܣ���ָ��Ŀ¼�Ĵ�С�ﵽһ���ٷֱ�ʱ����ո�Ŀ¼
'������ strCacheFolder--��Ҫ����Ƿ���յ�Ŀ¼
'���أ���
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        zl9ComLib.zlCommFun.ShowFlash "�����ͼ�񻺳�Ŀ¼����ȴ���", frmMain
        objCurFolder.Delete True
        zl9ComLib.zlCommFun.StopFlash
    End If
End Sub

Private Function FileIsOccupied(ByVal FilePath As String) As Boolean
'------------------------------------------------
'���ܣ��ж��ļ��Ƿ����ڱ���������ռ�ã�ͨ����ռ��ʽ�����ж�
'������ FilePath--��Ҫ�򿪵��ļ�
'���أ�True--��ռ�ã�False--����ռ��
'------------------------------------------------
    Dim fFile     As Integer
    
    fFile = FreeFile
    
    On Error GoTo ErrOpen
    Open FilePath For Binary Lock Read Write As fFile
    Close fFile
    Exit Function
ErrOpen:
    FileIsOccupied = True
End Function

Private Sub TimeDelay(lngTimeDelay As Long)
'------------------------------------------------
'���ܣ���ʱ
'������ lngTimeDelay--��Ҫ��ʱ��ʱ�䳤��
'���أ���
'------------------------------------------------
    Dim Savetime As Double
    
    On Error GoTo err
    Savetime = timeGetTime '���¿�ʼʱ��ʱ��
    While timeGetTime < Savetime + lngTimeDelay 'ѭ���ȴ�
'    DoEvents 'ת�ÿ���Ȩ���Ա��ò���ϵͳ�����������¼���
    Wend
    Exit Sub
err:
    
End Sub

Public Function funGetCandSImageAndShow(la As DicomLabel, imageViewer As DicomViewer, _
    axialViewer As DicomViewer, resultViewer As DicomViewer, intToltalHeight As Integer, _
    intType As Integer, blnFirst As Boolean, blnChangeLa As Boolean) As Boolean
'------------------------------------------------
'���ܣ� ���ݿ�����la����imageViewer�ڵ�ͼ����й�״λ��ʸ״λ�ؽ������ͼ��ʾ��ResultViewer��
'������ la          --- �����ؽ��Ŀ�����
'       imageViewer   --- �ؽ�ͼ�����ڵ�Viewer���ڸ�������
'       axialViewer -- ��λͼ�����ڵ�Viewer����б���ؽ�������
'       resultViewer -- �ؽ����ͼ���ڵ�Viewer����б���ؽ�������
'       intToltalHeight --- ����߶�
'       intType     --- ͼ�����ͣ�1--���ߣ�2--���ߣ�ֱ����д��ͼ�����
'       blnFirst    --- �Ƿ��һ�ε��ã�����ǵ�һ�ε��ã��򲻼�¼ԭ���Ĵ���λ��ͼ��״̬
'       blnChangeLa --- �Ƿ�ı�MPR�����ߣ����ֻ�ı�ͼ��û�������ߣ�����Ҫ��������MPRͼ��ֻ�ػ���Ӧ�߼���
'���أ��ؽ����ͼ�����ʧ�ܣ�False
'------------------------------------------------
    
    Dim resImage As DicomImage
    Dim resImages As New DicomImages
    Dim imgOld As DicomImage
    Dim imgNew As DicomImage
    Dim dblZoom As Double
    Dim lngScrollX As Long
    Dim lngScrollY As Long
    Dim lngWWidth As Long
    Dim lngWLevel As Long
    Dim blnStretchToFit As Boolean
    Dim img As DicomImage
    
    On Error GoTo err
    
    '��ȡ�ؽ����ͼ������ƶ��˿����ߣ��Ͳ�����ͼ�񣬷���ʹ��ԭ����ͼ��
    If blnChangeLa = True Then   '�ƶ���MPR�����ߣ�Ҫ�����µĽ��ͼ
        '���ݴ���Ŀ����ߣ���һ��Viewer�ڵ�ͼ������ؽ����������ؽ����ͼ
        Set resImage = funGetMPRResultImage(la, imageViewer, intToltalHeight, intType)
    Else
        Set resImage = resultViewer.Images(1)
    End If
    
    '��ʾ���ͼ
    If resImage Is Nothing Then
        funGetCandSImageAndShow = False
        Exit Function
    Else
        Set imgOld = Nothing
        If blnChangeLa = True Then  '�ı���MPR�����ߣ������µĽ��ͼ
            resImages.Clear
            resImages.Add resImage
            '��¼ԭ����ͼ��״̬���µ��ؽ����ͼ����ʹ����Щ״̬����һ���ؽ�������Ҫ��¼
            If resultViewer.Images.Count > 0 And blnFirst = False Then
                Set imgOld = resultViewer.Images(1)
                blnStretchToFit = imgOld.StretchToFit
                dblZoom = imgOld.ActualZoom
                lngScrollX = imgOld.ActualScrollX
                lngScrollY = imgOld.ActualScrollY
                lngWWidth = imgOld.width
                lngWLevel = imgOld.Level
            End If
            '��ͼ����ӵ����Viewer��
            resultViewer.Images.Clear
            resultViewer.Images.Add resImage
            
            Set img = resultViewer.Images(1)
            img.Tag = 1
            If img.Labels.Count = 0 Then
                Call subInitAImage(img, 0, resultViewer)
            End If
            
        End If
        
        If resultViewer.Images.Count > 0 Then
            Set imgNew = resultViewer.Images(1)
            imgNew.Refresh False
        End If
        
        '�ָ�ԭ��ͼ���״̬
        If Not imgOld Is Nothing And Not imgNew Is Nothing Then
            imgNew.StretchToFit = blnStretchToFit
            imgNew.Zoom = dblZoom
            imgNew.ScrollX = lngScrollX
            imgNew.ScrollY = lngScrollY
            imgNew.width = lngWWidth
            imgNew.Level = lngWLevel
        End If
            
        '���ؽ����ͼ�Ŀ�����
        If Not imgOld Is Nothing Then
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).left = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).left
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).top = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).top
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).width = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).width
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).height = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).height
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).Visible = True
            
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).left = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).left
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).top = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).top
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).width = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).width
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).height = imgOld.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).height
            imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).Visible = True
            imgNew.Refresh (False)
        Else
            Call subMPRSlopeDrawResultControlLabels(la, imgNew, imageViewer, axialViewer)
        End If
        
        funGetCandSImageAndShow = True
    End If
        
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub subMPRSlopeDrawResultControlLabels(la As DicomLabel, imgNew As DicomImage, imageViewer As DicomViewer, _
        axialViewer As DicomViewer)
'------------------------------------------------
'���ܣ� ��б���ؽ����ͼ�Ŀ�����
'������ la -- �����ؽ�����λ������
'       imgNew -- �ؽ����ͼ
'       imageViewer -- ԭͼ���ڵ�Viewer������������
'       axialViewer -- ��λͼ�����ڵ�Viewer����б���ؽ�������
'����:��
'------------------------------------------------
    Dim dblScale As Double
    
    On Error GoTo err
    
    '��������ȷ���Ƿ���ʾMPR������
    If blnShowMPRLine = True And Not imgNew Is Nothing Then
        '����λ��ʸ��״λͶӰ��
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).top = axialViewer.Images(1).Tag / imageViewer.Images.Count * imgNew.sizeY
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).left = 0
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).width = imgNew.sizeX
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).height = 0
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_H).Visible = True
        
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).top = 0
        
        '�������ɿ����ߵĹ�����ȷ�����ĵ���ͶӰ���е�λ��
        If Abs(la.width) > Abs(la.height) Then  '���ɺ���
            If la.width < 0 Then
                dblScale = 1 - Abs((axialViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_POINT_O).left - la.left) / la.width)
            ElseIf la.width > 0 Then
                dblScale = Abs((axialViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_POINT_O).left - la.left) / la.width)
            End If
        Else    '��������
            If la.height < 0 Then
                dblScale = 1 - Abs((axialViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_POINT_O).top - la.top) / la.height)
            ElseIf la.height > 0 Then
                dblScale = Abs((axialViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_POINT_O).top - la.top) / la.height)
            End If
        End If

        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).left = dblScale * imgNew.sizeX
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).height = imgNew.sizeY
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).width = 0
        
        imgNew.Labels(G_INT_SYS_LABEL_MPR_RESULT_V).Visible = True
        Call imgNew.Refresh(False)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub



