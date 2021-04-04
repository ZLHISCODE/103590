Attribute VB_Name = "mdlPublic"
Option Explicit
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Public Const ATTR_������� As String = "Study Date"
Public Const ATTR_���ʱ�� As String = "Study Time"
Public Const ATTR_�������� As String = "Series Date"
Public Const ATTR_����ʱ�� As String = "Series Time"
Public Const ATTR_Ӱ����� As String = "Modality"
Public Const ATTR_�豸�� As String = "Manufacturer"
Public Const ATTR_����豸 As String = "Manufacturer's Model Name"

Public gcnAccess As New ADODB.Connection, strBeginDate As String

Public gstrSQL As String

Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Public Sub GetUserInfo()
'����:�õ��û�����Ϣ

    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String
    
    rsTemp.CursorLocation = adUseClient
    On Error GoTo errHand
    
    With rsTemp
        strSQL = "select P.*,D.���� as ���ű���,D.���� as ��������,M.����ID" & _
                " from �ϻ���Ա�� U,��Ա�� P,���ű� D,������Ա M " & _
                " Where U.��Աid = P.id And P.ID=M.��ԱID and  M.ȱʡ=1 and M.����id = D.id and U.�û���=user"
        .Open strSQL, gcnOracle, adOpenKeyset
                
        If .RecordCount <> 0 Then
            glngUserId = .Fields("ID").Value                '��ǰ�û�id
            gstrUserCode = .Fields("���").Value            '��ǰ�û�����
            gstrUserName = .Fields("����").Value            '��ǰ�û�����
            gstrUserAbbr = IIf(IsNull(.Fields("����").Value), "", .Fields("����").Value)          '��ǰ�û�����
            glngDeptId = .Fields("����id").Value            '��ǰ�û�����id
            gstrDeptCode = .Fields("���ű���").Value        '��ǰ�û�
            gstrDeptName = .Fields("��������").Value        '��ǰ�û�
        Else
            glngUserId = 0
            gstrUserCode = ""
            gstrUserName = ""
            gstrUserAbbr = ""
            glngDeptId = 0
            gstrDeptCode = ""
            gstrDeptName = ""
        End If
        .Close
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0
End Sub

Public Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strFormCaption As String)
'���ܣ��򿪼�¼��ͬʱ����SQL���
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strFormCaption, gstrSQL)
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Public Sub ExecuteProcedure(ByVal strFormCaption As String)
'���ܣ�ִ�й���ʽ��SQL���
    Call SQLTest(App.ProductName, strFormCaption, gstrSQL)
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub SaveImages(Images As DicomImages, ByVal MainDeviceID As String, ByVal BufferDir As String)
'���ܣ�����ͼ��
    Dim curImage As DicomImage
    Dim i As Integer, iCount As Integer  '�����ͼ����
    Dim intSQL As Integer, rsTmp As New ADODB.Recordset
    
    Dim blnAddTmp As Boolean, blnTmp As Boolean
    Dim strAge As String, strBirth As String
    Dim strDirURL As String, strHost As String
    Dim dtReceived As String, dtCurrent As String
    
    Dim ImageType As String, CheckNo As Long, CheckDev As String
    Dim PatientName As String, EnglishName As String, Sex As String, Age As Integer
    Dim CheckUID As String, SeriesUID As String
    
    On Error GoTo DBError
    gcnOracle.BeginTrans
    If gcnAccess.State <> adStateClosed Then gcnAccess.BeginTrans
    
    dtCurrent = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    gstrSQL = "Select 'ftp://'||Decode(�û���,Null,'',�û���||Decode(����,Null,'',':'||����))" & _
        "||'@'||IP��ַ As Host,'/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As URL " & _
        "From Ӱ���豸Ŀ¼ " & _
        "Where �豸��='" & MainDeviceID & "'"
    If rsTmp.State <> adStateClosed Then rsTmp.Close
    OpenRecordset rsTmp, "PACSͼ�񱣴�"
    If rsTmp.EOF Then
        Err.Raise vbObjectError + 1, "PACSͼ�񱣴�", "�豸�����ô���"
    End If
    strHost = rsTmp("Host"): strDirURL = rsTmp("URL")
    
    iCount = 0
    For Each curImage In Images
        gstrSQL = "Select ͼ��UID From Ӱ����ͼ�� Where ͼ��UID='" & _
            curImage.InstanceUID & "' Union All Select ͼ��UID From Ӱ����ʱͼ�� Where ͼ��UID='" & _
            curImage.InstanceUID & "'"
        If rsTmp.State <> adStateClosed Then rsTmp.Close
        OpenRecordset rsTmp, "PACSͼ�񱣴�"
        '��ͼ��
        If rsTmp.EOF Then
            gstrSQL = "Select ���UID From Ӱ�����¼ Where ���UID='" & curImage.StudyUID & "'" & _
                " Union All Select ���UID From Ӱ����ʱ��¼ Where ���UID='" & curImage.StudyUID & "'"
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            OpenRecordset rsTmp, "PACSͼ�񱣴�"
            '������ID��Ӣ��������
            If rsTmp.EOF Then
                blnAddTmp = True
                If IsNumeric(curImage.PatientID) Then
                    gstrSQL = "Select Distinct A.ҽ��ID,A.���ͺ� From Ӱ�����¼ A,����ҽ������ B,����ҽ����¼ C" & _
                        " Where A.ҽ��ID=B.ҽ��ID And A.���ͺ�=B.���ͺ� And B.ҽ��ID=C.ID" & _
                        " And C.����ID=" & curImage.PatientID & _
                        " And B.ִ��״̬=3 And B.ִ�й���=2"
                    If rsTmp.State <> adStateClosed Then rsTmp.Close
                    OpenRecordset rsTmp, "PACSͼ�񱣴�"
                    '��HIS��д�ļ���¼��Ӧ
                    If rsTmp.RecordCount = 1 Then
                        '������UID
                        gstrSQL = "ZL_Ӱ�����¼_SET(" & rsTmp(0) & "," & rsTmp(1) & ",'" & _
                            curImage.StudyUID & "','" & GetImageAttribute(curImage.Attributes, ATTR_����豸) & "'," & _
                            "to_Date('" & dtCurrent & "','YYYY-MM-DD HH24:MI:SS'),'" & MainDeviceID & "')"
                        ExecuteProcedure "PACSͼ�񱣴�"
                        blnAddTmp = False
                    End If
                End If
                '������ʱ����¼
                If blnAddTmp Then
                    If IsDate(curImage.DateOfBirthAsDate) Then
                        strAge = CStr(Year(Date) - Year(curImage.DateOfBirthAsDate))
                        strBirth = Format(curImage.DateOfBirthAsDate, "YYYY-MM-DD")
                    Else
                        strAge = "": strBirth = ""
                    End If
                    gstrSQL = "ZL_Ӱ����ʱ���_INSERT('" & GetImageAttribute(curImage.Attributes, ATTR_Ӱ�����) & "',Null,'" & _
                        curImage.Name & "','" & curImage.Name & "','" & _
                        curImage.Sex & "','" & strAge & "'," & _
                        IIf(Len(strBirth) = 0, "Null", "to_Date('" & strBirth & "','YYYY-MM-DD')") & ",Null,Null,'" & _
                        GetImageAttribute(curImage.Attributes, ATTR_����豸) & "','" & curImage.StudyUID & "'," & _
                        "to_Date('" & dtCurrent & "','YYYY-MM-DD HH24:MI:SS'),'" & MainDeviceID & "')"
                    ExecuteProcedure "PACSͼ�񱣴�"
                End If
            End If
            
            gstrSQL = "Select 0 As ��ʱ,��������,Ӱ�����,Nvl(����,0) As ����," & _
                "����豸,����,Ӣ����,�Ա�,Nvl(����,'-1') As ����,���UID From Ӱ�����¼ Where ���UID='" & curImage.StudyUID & "'" & _
                " Union All Select 1 As ��ʱ,��������,Ӱ�����,Nvl(����,0) As ����," & _
                "����豸,����,Ӣ����,�Ա�,Nvl(����,'-1') As ����,���UID From Ӱ����ʱ��¼ Where ���UID='" & curImage.StudyUID & "'"
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            OpenRecordset rsTmp, "PACSͼ�񱣴�"
            blnTmp = IIf(rsTmp(0) = 1, True, False) '���к�ͼ���Ƿ������ʱ��¼��
            dtReceived = Format(rsTmp(1), "yyyyMMdd")
            
            ImageType = Nvl(rsTmp(2)): CheckNo = rsTmp(3): CheckDev = Nvl(rsTmp(4))
            PatientName = Nvl(rsTmp(5)): EnglishName = Nvl(rsTmp(6)): Sex = Nvl(rsTmp(7)): Age = Val(rsTmp(8))
            CheckUID = Nvl(rsTmp(9))
            
            gstrSQL = "Select ����UID From " & IIf(blnTmp, "Ӱ����ʱ����", "Ӱ��������") & _
                " Where ����UID='" & curImage.SeriesUID & "'"
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            OpenRecordset rsTmp, "PACSͼ�񱣴�"
            '�����µļ������
            If rsTmp.EOF Then
                gstrSQL = "ZL_Ӱ������_INSERT('" & curImage.StudyUID & "','" & curImage.SeriesUID & "','" & _
                    curImage.SeriesDescription & "'," & _
                    IIf(blnTmp, 1, 0) & ")"
                ExecuteProcedure "PACSͼ�񱣴�"
            End If
            
            '�����µ�ͼ��
            gstrSQL = "ZL_Ӱ��ͼ��_INSERT('" & curImage.InstanceUID & "','" & curImage.SeriesUID & "','" & _
                curImage.SeriesDescription & "'," & _
                IIf(blnTmp, 1, 0) & ")"
            ExecuteProcedure "PACSͼ�񱣴�"
            
            '���汾����־
            WriteRecord ImageType, CheckNo, CheckDev, PatientName, EnglishName, Sex, Age, CheckUID, curImage.SeriesUID, blnTmp
            
            '����ͼ�񵽻���Ŀ¼
            curImage.WriteFile BufferDir & curImage.InstanceUID, True
            WriteToURL BufferDir & curImage.InstanceUID, strHost, strDirURL & _
                dtReceived & "/" & curImage.StudyUID & "/" & curImage.InstanceUID
            Kill BufferDir & curImage.InstanceUID
        Else
            WriteLog 3, vbObjectError + 1, "Ӱ��" & curImage.InstanceUID & "�Ѵ��ڣ�"
        End If
        iCount = iCount + 1
    Next
    
    If gcnAccess.State <> adStateClosed Then gcnAccess.CommitTrans
    gcnOracle.CommitTrans
    
    For i = 1 To iCount
        Images.Remove 1
    Next
    Exit Sub
DBError:
    If gcnAccess.State <> adStateClosed Then gcnAccess.RollbackTrans
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "���ͼ�񱣴�"
End Sub

Public Sub WriteToURL(ByVal SrcFileName As String, ByVal DestAddress As String, ByVal DestFileName As String)
'���ܣ��������ļ����浽Զ��������
    Dim iNet As Object
    
    Set iNet = CreateObject("InetCtls.inet.1")
    iNet.AccessType = 0: iNet.URL = DestAddress
    
    MkDir_Remote DestAddress, DestFileName
    iNet.Execute , "Put " & SrcFileName & " " & DestFileName
    Do While iNet.StillExecuting
        DoEvents
    Loop
End Sub

Public Sub MkDir_Remote(ByVal DestAddress As String, ByVal DestFileName As String)
    Dim iNet As Object, objFile As New Scripting.FileSystemObject, strPath As String
    Dim aNestPath() As Variant, i As Integer
    
    aNestPath = Array()
    
    Set iNet = CreateObject("InetCtls.inet.1")
    iNet.AccessType = 0: iNet.URL = DestAddress
    
    strPath = objFile.GetParentFolderName(DestFileName)
    Do While Len(strPath) > 0
        ReDim Preserve aNestPath(UBound(aNestPath) + 1)
        aNestPath(UBound(aNestPath)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    For i = UBound(aNestPath) To 0 Step -1
        iNet.Execute , "MkDir " & aNestPath(i)
        Do While iNet.StillExecuting
            DoEvents
        Loop
    Next
End Sub

Public Function GetImageAttribute(objAttr As DicomAttributes, ByVal AttrName As String) As Variant
    Dim curAttr As DicomAttribute
    
    GetImageAttribute = ""
    For Each curAttr In objAttr
        If UCase(curAttr.Description) = UCase(AttrName) Then
            If curAttr.Exists Then GetImageAttribute = curAttr.Value
            Exit For
        End If
    Next
End Function

Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer, _
    Optional ByVal MaxRows As Integer = 0, Optional ByVal MaxCols As Integer = 0)
'���ܣ�����DicomViewer��������
    Dim iCols As Integer, iRows As Integer
    
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    If MaxRows > 0 And iRows > MaxRows Then
        iRows = MaxRows
        iCols = CInt(ImageCount / iRows)
        If iRows * iCols < ImageCount Then iCols = iCols + 1
    End If
    If MaxCols > 0 And iCols > MaxCols Then
        iCols = MaxCols
        iRows = CInt(ImageCount / iCols)
        If iRows * iCols < ImageCount Then iRows = iRows + 1
    End If
    If MaxRows > 0 And iRows > MaxRows Then iRows = MaxRows
    
    Rows = iRows: Cols = iCols
End Sub

Public Function ImageExist(Images As DicomImages, SeekImage As DicomImage) As Boolean
    Dim curImage As DicomImage
    
    ImageExist = False
    For Each curImage In Images
        If curImage.InstanceUID = SeekImage.InstanceUID Then ImageExist = True: Exit For
    Next
End Function

Private Sub WriteRecord(ByVal ImageType As String, ByVal CheckNo As Long, ByVal CheckDev As String, _
    ByVal PatientName As String, ByVal EnglishName As String, ByVal Sex As String, Age As Integer, _
    ByVal CheckUID As String, ByVal SeriesUID As String, ByVal ifTmp As Boolean)
    
    Dim rsTmp As ADODB.Recordset, strSQL As String
    If gcnAccess.State = adStateClosed Then Exit Sub
    
    strSQL = "Select * from Ӱ��������� Where ����UID='" & SeriesUID & "' And ����ʱ��>cDate('" & _
        strBeginDate & "')"
    Set rsTmp = gcnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        strSQL = "Insert Into Ӱ���������(Ӱ�����,����,����豸,����,Ӣ����,�Ա�,����,Ӱ����,����UID,���UID,��Ӧ���,����ʱ��)" & _
            " Values('" & ImageType & "'," & IIf(CheckNo = 0, "Null", CheckNo) & ",'" & CheckDev & "','" & _
            PatientName & "','" & EnglishName & "','" & Sex & "'," & IIf(Age = -1, "Null", Age) & ",1,'" & _
            SeriesUID & "','" & CheckUID & "'," & CStr(Not ifTmp) & ",cDate('" & _
            Date & " " & Time() & "'))"
    Else
        strSQL = "Update Ӱ��������� Set Ӱ����=Ӱ����+1 Where ����UID='" & SeriesUID & "' And ����ʱ��>cDate('" & _
        strBeginDate & "')"
    End If
    gcnAccess.Execute strSQL
End Sub

Public Sub WriteLog(ByVal ErrorType As Integer, ErrorNum As Long, ErrorDesc As String)
    Dim strSQL As String
    If gcnAccess.State = adStateClosed Then Exit Sub
    
    strSQL = "Insert Into ������־(����ʱ��,��������,�����,������Ϣ) " & _
        "Values(cDate('" & Date & " " & Time() & "')," & ErrorType & "," & ErrorNum & ",'" & Replace(ErrorDesc, "'", "''") & "')"
    gcnAccess.Execute strSQL
End Sub
'��ʾ����Ŀ¼
Public Function BrowPath(lWindowHwnd As Long, Optional ByVal sTitle As String = "") As String
    Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    With udtBI
        '�����������
        .hWndOwner = lWindowHwnd
        '����ѡ�е�Ŀ¼
        .ulFlags = BIF_RETURNONLYFSDIRS
        If sTitle = "" Then
            .lpszTitle = "��ѡ����ʼ�������ļ��У�"
        Else
            .lpszTitle = sTitle
        End If
    End With
    '�����������
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        '��ȡ·��
        SHGetPathFromIDList lpIDList, sPath
        '�ͷ��ڴ�
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    BrowPath = sPath
End Function

