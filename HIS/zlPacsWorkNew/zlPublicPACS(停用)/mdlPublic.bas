Attribute VB_Name = "mdlPublic"
Option Explicit

'�ļ�����  1-���ﲡ��;2-סԺ����;3-�����¼;4-������;5-�������;6-֪���ļ�;7-���Ʊ���;8-��������
Public Enum EPRDocTypeEnum
    cpr���ﲡ�� = 1
    cprסԺ���� = 2
    cpr�����¼ = 3
    cpr������ = 4
    cpr������� = 5
    cpr֪���ļ� = 6
    cpr���Ʊ��� = 7             '���Ƶ��ݣ�����
    cpr�������� = 8             '���Ƶ��ݣ�����
End Enum

Public Const ELE_BACKCOLOR = &HD5FEFF               'Ҫ�صı�����ɫ '&HDCDCDC
Public Const ELE_UNDERLINE = cprWave                'Ҫ�ص��»���
Public Const PROTECT_FORECOLOR = &H662200           '�Զ��屣���ı���ǰ��ɫ

Public gobjComLib As Object    'zl9ComLib.clsComLib
Public gcnOracle As ADODB.Connection
Public gstrSysName  As String
Public glngSys As Long

Public gstrSQL As String
Private mclsUnzip As Object

Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long

Public Type NETRESOURCE ' ������Դ
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Public Const RESOURCETYPE_ANY = &H0


Public Sub MkLocalDir(ByVal strDir As String)
'���ܣ���������Ŀ¼
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

Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'���ܣ���ָ��Ŀ¼�Ĵ�С�ﵽһ���ٷֱ�ʱ����ո�Ŀ¼
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        objCurFolder.Delete True
    End If
End Sub

Public Function funcConnectShardDir(strShareRemoteDir As String, strUserName As String, strPassWord As String) As Long
    '����������Դ
    Dim NetR As NETRESOURCE
    Dim lngResult As Long
    
    NetR.dwType = RESOURCETYPE_ANY
    NetR.lpLocalName = vbNullString
    NetR.lpRemoteName = strShareRemoteDir
    NetR.lpProvider = vbNullString
    lngResult = WNetAddConnection2(NetR, strPassWord, strUserName, 0)
    
    If lngResult <> 0 Then
        MsgBox "��������ʧ�ܣ��������������Ƿ���ȷ��", vbInformation, gstrSysName
    End If
    funcConnectShardDir = lngResult
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer)
'-----------------------------------------------------------------------------
'���ܣ����������ͼ��������ͼ������Ŀ�Ⱥ͸߶ȣ�������ѵ�ͼ����������������
'������ ImageCount����ͼ������
'       RegionWidth--ͼ����ʾ����Ŀ��
'       RegionHeight--ͼ����ʾ����ĸ߶�
'       Rows����[����]�������
'       Cols����[����]�������
'���أ������������Rows���������Cols
'-----------------------------------------------------------------------------
    Dim iCols As Integer, iRows As Integer
    Dim iBase As Integer, blnDoLoop As Integer
    Dim lngFreeCount As Long
    
    If RegionHeight = 0 Then RegionHeight = 1
    If RegionWidth = 0 Then RegionWidth = 1
    
    On Error GoTo err
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))

    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    '��ͼ���ʽΪ���µ���ʽʱ����Ҫ�����н�������
    
    '��ʽ1��
    'ͼ1  ͼ2  ͼ3  ͼ4
    'ͼ5  ͼ6  ͼ7  ͼ8
    '��1  ��2  ��3  ��4
    
    '��ʽ2��
    'ͼ1  ͼ2  ͼ3  ͼ4
    'ͼ5  ͼ6  ͼ7  ͼ8
    'ͼ9  ��1  ��2  ��3
    
    lngFreeCount = iRows * iCols - ImageCount
    Do While lngFreeCount >= iCols Or lngFreeCount >= iRows
        If lngFreeCount >= iCols Then
            iRows = iRows - 1
        Else
            iCols = iCols - 1
        End If
        
        lngFreeCount = iRows * iCols - ImageCount
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / iCols > RegionHeight > iRows Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    
    '�ٴ�����������
    lngFreeCount = iRows * iCols - ImageCount
    Do While lngFreeCount >= iCols Or lngFreeCount >= iRows
        If lngFreeCount >= iCols Then
            iRows = iRows - 1
        Else
            iCols = iCols - 1
        End If
        
        lngFreeCount = iRows * iCols - ImageCount
    Loop
    
    Rows = iRows: Cols = iCols
err:
End Sub

Public Function FindNextKey(ByRef edtThis As Object, _
    ByVal lngCurPosition As Long, _
    ByVal strKeyType As String, _
    ByRef lngKey As Long, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTmp As String
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��
    
    sTmp = strKeyType & "S("
    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTmp)
        If i <> 0 Then
            '���Ƿ��ǹؼ���
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                i = i + 1
                GoTo LL1
            End If
            '���ҵ���ʼ�ؼ���
            
            '���ҽ����ؼ���
            j = i + 16
LL2:
            sTmp = strKeyType & "E("
            j = InStr(j, sText, sTmp)
            If j <> 0 Then
                '���Ƿ��ǹؼ���
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '�ҵ������ؼ���
                strKeyType = strKeyType
                lngKSS = i - 1 'ת��Ϊ0��ʼ������λ�á�
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                lngKey = Val(.Range(i + 2, i + 10))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindNextKey = True
            End If
        End If
    End With
End Function

Public Sub ReadRTF(edtThis As Editor, ByVal lngFileID As Long, ByVal blnClearMode As Boolean, ByVal blnMoved As Boolean)
'��ȡRTF�ļ�
On Error GoTo errH
Dim ParaFmt As New cParaFormat, FontFmt As New cFontFormat, i As Long, lngLen As Long
Dim oEles As Object, oTabs As Object, oPics As Object
Dim strZipFile As String, strRtfFile As String, j As Long, rs As New ADODB.Recordset
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
Dim objFSO As New FileSystemObject
Dim strSQL As String

    Set oEles = CreateObject("zlRichEPR.cEPRElements")
    Set oTabs = CreateObject("zlRichEPR.cEPRTables")
    Set oPics = CreateObject("zlRichEPR.cEPRPictures")
    
    strZipFile = zlBlobRead(5, lngFileID, , blnMoved)
    If objFSO.FileExists(strZipFile) Then
        strRtfFile = zlFileUnzip(strZipFile)
        If objFSO.FileExists(strRtfFile) Then
                edtThis.OpenDoc strRtfFile
                objFSO.DeleteFile strRtfFile, True
        End If
        objFSO.DeleteFile strZipFile, True
    End If
    If Trim(edtThis.Text) = "" Then Exit Sub

    '��ȡͼƬ,���,Ҫ��
    strSQL = "Select Level,ID, �ļ�id,��ʼ��, ��ֹ��," & _
                "   ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������id, �������," & vbNewLine & _
                "       ʹ��ʱ��, ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��" & vbNewLine & _
                "From (Select ID, �ļ�id,��ʼ��, ��ֹ��," & _
                "               ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id,�������id," & vbNewLine & _
                "              �������, ʹ��ʱ��, ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��" & vbNewLine & _
                "       From ���Ӳ�������" & vbNewLine & _
                "       Where �ļ�id = [1] And �������>0 and �������<ID)" & vbNewLine & _
                "Start With ��id Is Null" & vbNewLine & _
                "Connect By Prior ID = ��id" & vbNewLine & _
                "Order By �������, �����д�"
    If blnMoved Then strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")
    Set rs = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ԭ��ͼ", lngFileID)
    Do Until rs.EOF
        Select Case rs!��������
            Case 3  '���
                lKey = oTabs.Add(Nvl(rs!������, 0))                  '�ָ�Keyֵ��
                Call oTabs("K" & lKey).FillTableMember(rs, "���Ӳ�������")
            Case 4  'Ҫ��
                lKey = oEles.Add(Nvl(rs!������, 0))
                Call oEles("K" & lKey).FillElementMember(rs, "���Ӳ�������")
            Case 5  'ͼƬ
                lKey = oPics.Add(Nvl(rs("������"), 0))
                Call oPics("K" & lKey).FillPictureMember(rs, "���Ӳ�������")
        End Select
        rs.MoveNext
    Loop
    
    For j = 1 To oPics.Count '��ԭͼƬ
        bFinded = FindKey(edtThis, "P", oPics(j).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            oPics(j).DeleteFromEditor edtThis
            oPics(j).InsertIntoEditor edtThis, -1, True
        End If
    Next
    
    For j = 1 To oTabs.Count '��ԭ���
        bFinded = FindKey(edtThis, "T", oTabs(j).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
                Set ParaFmt = edtThis.Range(lKSE, lKES).Para.GetParaFmt
                Set FontFmt = edtThis.Range(lKSE, lKES).Font.GetFontFmt
                
                If oTabs(j).�Ƿ��� Then
                    edtThis.Range(lKSS, lKEE + 2).Text = ""
                Else
                    edtThis.Range(lKSS, lKEE).Text = ""
                End If
                oTabs(j).InsertIntoEditor edtThis, lKSS, , , True
                
                edtThis.Range(lKSE, lKES).Para.SetParaFmt ParaFmt
                edtThis.Range(lKSE, lKES).Font.SetFontFmt FontFmt
                edtThis.Range(lKSS, lKEE).Font.Protected = True
        End If
    Next
    
    For j = 1 To oEles.Count 'ɾ���յ�Ҫ�أ�չ����Ҫ������ˢ����ȥ���²�����
        If oEles(j).�����ı� = "" Then
            oEles(j).DeleteFromEditor edtThis
        ElseIf oEles(j).������̬ = 1 Then
            oEles(j).Refresh edtThis
        End If
    Next
    

    ' �������ĵ��Ĵ���
    edtThis.SelectAll
    If blnClearMode Then
        edtThis.AuditMode = True
        edtThis.AcceptAuditText    '���ģʽ
    End If
    lngLen = Len(edtThis.Text)
    For i = 0 To lngLen - 1 'ֻ������ɫΪҪ�ر���ɫ��ɫȥ��
        If edtThis.Range(i, i + 1).Font.BackColor = ELE_BACKCOLOR Then
            edtThis.Range(i, i + 1).Font.BackColor = tomAutoColor
        End If
        If edtThis.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR Then
            edtThis.Range(i, i + 1).Font.ForeColor = tomAutoColor
        End If
    Next
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Public Function FindKey(ByRef edtThis As Object, _
        ByRef strKeyType As String, _
        ByRef lngKey As Long, _
        ByRef lngKSS As Long, _
        ByRef lngKSE As Long, _
        ByRef lngKES As Long, _
        ByRef lngKEE As Long, _
        ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTmp As String
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��
    
    sTmp = strKeyType & "S(" & Format(lngKey, "00000000")
    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        i = 1
LL1:
        i = InStr(i, sText, sTmp)
        If i <> 0 Then
            '���Ƿ��ǹؼ���
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                i = i + 1
                GoTo LL1
            End If
            '���ҵ���ʼ�ؼ���
            
            '���ҽ����ؼ���
            j = i + 16
LL2:
            sTmp = strKeyType & "E(" & Format(lngKey, "00000000")
            j = InStr(j, sText, sTmp)
            If j <> 0 Then
                '���Ƿ��ǹؼ���
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '�ҵ������ؼ���
                strKeyType = strKeyType
                lngKSS = i - 1 'ת��Ϊ0��ʼ������λ�á�
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindKey = True
            End If
        End If
    End With
End Function

Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String, Optional ByVal blnMoved As Boolean) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, StrText As String
    Dim rsLob As New ADODB.Recordset
    Dim strSQL As String
    
    err = 0: On Error GoTo errHand
    
    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum
    
    strSQL = "Select Zl_Lob_Read([1],[2],[3],[4]) as Ƭ�� From Dual"
    lngCount = 0
    Do
        Set rsLob = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "zlBlobRead", Action, KeyWord, lngCount, IIf(blnMoved, 1, 0))
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        StrText = rsLob.Fields(0).Value
        
        ReDim aryChunk(Len(StrText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(StrText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

errHand:
    Close lngFileNum
    Kill strFile
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
    zlBlobRead = ""
End Function

Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    Dim objFSO As New FileSystemObject
    
    On Error GoTo errHand
    
    If Not objFSO.FileExists(strZipFile) Then zlFileUnzip = "": Exit Function
    
    strZipPath = objFSO.GetSpecialFolder(2) 'ȡ��ʱĿ¼
    strZipPathTmp = strZipPath & "\" & Format(Now, "yyMMdd") & CStr(100 * Timer)
    If Not objFSO.FolderExists(strZipPathTmp) Then Call objFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp & "\TMP.RTF"
    If objFSO.FileExists(strZipFileTmp) Then objFSO.DeleteFile strZipFileTmp
    
    If mclsUnzip Is Nothing Then Set mclsUnzip = CreateObject("zlRichEPR.cUnzip")

    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    
    If objFSO.FileExists(strZipFileTmp) Then
        
        strZipFileName = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer) & ".RTF"
        If objFSO.FileExists(strZipFileName) Then objFSO.DeleteFile strZipFileName
                
        Call objFSO.CopyFile(strZipFileTmp, strZipFileName)
        
        If objFSO.FileExists(strZipFileTmp) Then objFSO.DeleteFile strZipFileTmp, True
        
        On Error Resume Next
        If objFSO.FolderExists(strZipPathTmp) Then objFSO.DeleteFolder strZipPathTmp, True
        
        zlFileUnzip = strZipFileName
    Else
        zlFileUnzip = ""
    End If
    
    Exit Function
    
errHand:
    Call gobjComLib.SaveErrLog
End Function

Public Function GetFileRange(ByVal lFileId As Long, ByVal lngRecordId As Long, ByVal strCreateTime As String, _
                            ByVal eDocType As EPRDocTypeEnum, ByVal lngPatId As Long, ByVal lngPageId As Long, _
                            Optional ByVal blnMoved As Boolean) As String
    '******************************************************************************************************************
    '���ܣ���ȡ��ǰ����(���ܵ�ǰ����δ����)ǰ���й�����
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTemp As New ADODB.Recordset, strTime As String, dStar As Date, dEnd As Date
    Dim strSQL As String, strIDs As String, blnNewPage As Boolean, n_Num As Integer, n_S As Integer, n_E As Integer

    On Error GoTo errHand
    strTime = Format(strCreateTime, "yyyy-MM-dd HH:mm:ss")
    If strTime = "" Then strTime = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If strTime = "00:00:00" Then strTime = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    blnNewPage = gobjComLib.zlDatabase.GetPara("ת�ƺ�Ҫ����д�Ĺ���������һҳ��ӡ", glngSys, 1251, 1) = 1 '=0��ʾת�ƺ�����������ӡ =1 ��ʾ����һҳ��ӡ

    strSQL = "Select m.Id" & vbNewLine & _
            "From �����ļ��б� L, �����ļ��б� M" & vbNewLine & _
            "Where l.Id = [1] And l.ҳ�� = m.��� And l.ҳ�� = m.ҳ�� And " & vbNewLine & _
            "      m.���� In (" & Decode(eDocType, 4, "4", 1, "1", "2, 5, 6") & ") And l.���� = m.����"
    If blnNewPage And eDocType = 2 Then 'סԺ����ת�ƺ�����һҳ
    strSQL = strSQL & vbNewLine & _
            "Union" & vbNewLine & _
            "Select b.Id" & vbNewLine & _
            "From �����ļ��б� A, �����ļ��б� B, ����ʱ��Ҫ�� C" & vbNewLine & _
            "Where a.Id = [1] And a.ҳ�� = b.ҳ�� And b.���� In (" & Decode(eDocType, 4, "4", 1, "1", "2, 5, 6") & ") And " & vbNewLine & _
            "      a.���� = b.���� And c.�ļ�id = b.Id And c.�¼� = 'ת��' And c.��дʱ�� >= 0"
    End If
    
    If lngRecordId <> 0 Then
        gstrSQL = "Select nvl(���,0) ��� From ���Ӳ�����¼ Where ID=[1]"
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���", lngRecordId)
        n_Num = rsTemp!���
    Else
        n_Num = 9999
    End If
    
    '��ȡҳ���ļ���ǰʱ��֮ǰ���һ����д��¼
    gstrSQL = "Select ����ʱ��,���" & vbNewLine & _
                "From (Select a.����ʱ��,a.���" & vbNewLine & _
                "       From ���Ӳ�����¼ A, (" & strSQL & ") B" & vbNewLine & _
                "       Where a.�ļ�id = b.Id And a.����id = [2] And a.��ҳid = [3] And " & IIf(n_Num = 0, "a.����ʱ�� <= [4]", "((a.����ʱ�� <= [4] and Nvl(a.���,0)=0) Or  a.���<=[5])") & vbNewLine & _
                "       Order By " & IIf(n_Num <> 0, "a.���", "a.����ʱ��") & " Desc)" & vbNewLine & _
                "Where Rownum = 1"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҳ���ļ�֮ǰһ����д", lFileId, lngPatId, lngPageId, CDate(strTime), n_Num)
    If rsTemp.EOF Then
        dStar = CDate("2000-01-01 00:00:00"): n_S = -1 '����֮ǰû��д��ҳ���ļ�
    Else
        dStar = CDate(rsTemp!����ʱ��)
        n_S = IIf(n_Num = 0, 0, Nvl(rsTemp!���, -1))
    End If
    
    '��ȡҳ���ļ���ǰʱ��֮�����һ����д��¼
    gstrSQL = "Select ����ʱ��,���" & vbNewLine & _
                "From (Select a.����ʱ��,a.���" & vbNewLine & _
                "       From ���Ӳ�����¼ A, (" & strSQL & ") B" & vbNewLine & _
                "       Where a.�ļ�id = b.Id And a.����id = [2] And a.��ҳid = [3] And " & IIf(n_Num = 0, "a.����ʱ�� > [4]", "((a.����ʱ�� > [4] and Nvl(a.���,0)=0) Or a.���>[5])") & vbNewLine & _
                "       Order By a.����ʱ��)" & vbNewLine & _
                "Where Rownum = 1"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҳ���ļ�֮��һ����д", lFileId, lngPatId, lngPageId, CDate(strTime), n_Num)
    If rsTemp.EOF Then
        dEnd = CDate("3001-01-01") '����֮��û��д��ҳ���ļ�
        n_E = 9999
    Else
        dEnd = CDate(rsTemp!����ʱ��) - 1 / 24 / 60 / 60 '����֮��д����ȡ����¼��ʱ���һ��,���������˼�¼
        n_E = IIf(n_Num = 0, 0, Nvl(rsTemp!���, 9999) - 1)
    End If
    
    '������ͬ�������Բ�������д��¼
    strSQL = "Select m.Id" & vbNewLine & _
        "From �����ļ��б� L, �����ļ��б� M" & vbNewLine & _
        "Where l.Id = [1] And l.ҳ�� = m.ҳ�� And " & vbNewLine & _
        "      m.���� In (" & Decode(eDocType, 4, "4", 1, "1", "2, 5, 6") & ") And l.���� = m.����"
    gstrSQL = "Select a.ID" & vbNewLine & _
                "   From ���Ӳ�����¼ A, (" & strSQL & ") B" & vbNewLine & _
                "   Where a.�ļ�id = b.Id And a.����id = [2] And a.��ҳid = [3] And " & IIf(n_Num = 0, "a.����ʱ�� Between [4] And [5]", "((a.����ʱ�� Between [4] And [5] and Nvl(a.���,0)=0 ) or a.��� Between [6] And [7])") & vbNewLine & _
                "   Order By a.����ʱ��"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҳ���ļ�֮��һ����д", lFileId, lngPatId, lngPageId, dStar, dEnd, n_S, n_E)
    Do Until rsTemp.EOF
        strIDs = strIDs & "," & rsTemp!Id
        rsTemp.MoveNext
    Loop
    strIDs = Mid(strIDs, 2)
    
    GetFileRange = strIDs
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

