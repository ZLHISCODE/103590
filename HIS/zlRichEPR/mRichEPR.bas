Attribute VB_Name = "mRichEPR"
'#########################################################################
'##ģ �� ����mRichEPR.bas
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��8��11��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ����ȫ�ֱ��������͵Ķ���
'##��    ����
'#########################################################################

Option Explicit

'##########################################################################################
'## ȫ������
'##########################################################################################

Public Type PreDefinedKeyInfo   '�����ؼ���
    KeyStart As String
    KeyEnd As String
End Type

'##########################################################################################
'## ȫ�ֱ���
'##########################################################################################
Public gfrmPublic As frmPublic

Public gblnShowInTaskBar As Boolean         '�Ƿ���ʾ��������������
Public gcnOracle As New ADODB.Connection    '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                  '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrProductName As String            '��Ʒ��ƣ����磺����
Public gstrSysName As String                'ϵͳ���ƣ����磺�������
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public glngModul As Long                    'ģ����
Public glngSys As Long                      'ϵͳ��ţ����磺100

Public gstrDbOwner As String                '��ǰ���ݿ������ߣ���ͬģ����ܲ�һ����
Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����
Public gstrSignName As String               'ǩ������
Public gstrPrivsEpr As String               '�����༭ģ��1070Ȩ��
Public gstrCopyPID As String                '����Դ����ID

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public gstrMatch As String                  '���ݱ��ز�����ƥ��ģʽ��ȷ������ƥ�����
Public gstrSQL As String

Public gobjFSO As New Scripting.FileSystemObject    'FSO����
Public gfrmParent As Object                         'ȫ�ֵĸ�����
Public gobjPacsCore As Object                       '���������������Ǵ����˹�Ƭվ����
Public gobjESign As Object                  '����ǩ���ӿڲ���
Public gobjTendESign As Object           '����ǩ���ӿڲ���(����)
Public gstrESign As String                  '�Ƿ����õ���ǩ��
Public gobjEmr As Object                    '�°���Ӳ���
Public gobjInfection As Object              '��Ⱦ�����濨
Public gobjPlugIn As Object                 '���
Public gobjRegister As Object               'ZLHIS������֤���

Public gKeyWords(1 To 6) As PreDefinedKeyInfo       'Ԥ����ؼ���

Public Const ELE_BACKCOLOR = &HD5FEFF               'Ҫ�صı�����ɫ '&HDCDCDC
Public Const ELE_UNDERLINE = cprwave                'Ҫ�ص��»���
Public Const PROTECT_BGCOLOR = &HE0E0E0             '�Զ��屣���ı��ı���ɫ
Public Const PROTECT_FORECOLOR = &H662200           '�Զ��屣���ı���ǰ��ɫ
Public Const TABLEELE_FORECOLOR = &H100080          '���Ҫ�ص�ǰ��ɫ
Public Const ELE_JUMP_LIMIT = 32                    '�س������Զ�������һҪ�صľ�������

'ˢ������ʱ����Ĳ���״̬
Public Enum TYPE_PATI_State
    ps��Ժ = 0
    psԤ�� = 1
    ps��Ժ = 2
    ps���� = 3          'ҽ��վ:�����ﲡ��(��Ժ)
    ps���� = 4          'ҽ��վ:�ѻ��ﲡ��
    ps���ת�� = 5      'ҽ��վ:���ת�ƻ�ת�����Ĳ���(��Ժ)
    ps��ת�� = 6        'ҽ��վ:��ƴ���ס��ת��������������
End Enum

'##########################################################################################
'## ѹ�����ѹ
'##########################################################################################
Private mclsZip As New cZip
Private mclsUnzip As New cUnzip

'################################################################################################################
'## ���ܣ�  ��ָ����LOB�ֶθ���Ϊ��ʱ�ļ�
'##
'## ������  Action      :�������ͣ����������ǲ����ĸ���
'##         KeyWord     :ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'##         strFile     :�û�ָ����ŵ��ļ�������ָ��ʱ��ȡ��ǰ·�������ļ���
'##
'## ���أ�  ������ݵ��ļ�����ʧ���򷵻��㳤��""
'##
'## ˵����  Actionȡֵ˵����
'##         0-�������ͼ�Σ�1-�����ļ���ʽ��2-�����ļ�ͼ�Σ�3-�������ĸ�ʽ��4-��������ͼ�Σ�5-���Ӳ�����ʽ��6-���Ӳ���ͼ�Σ�
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String, Optional ByVal blnMoved As Boolean) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, StrText As String
    Dim rsLob As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand
    
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
    
    gstrSQL = "Select Zl_Lob_Read([1],[2],[3],[4]) as Ƭ�� From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlBlobRead", Action, KeyWord, lngCount, IIf(blnMoved, 1, 0))
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
    If ErrCenter = 1 Then
        Resume
    End If
    zlBlobRead = ""
End Function


'Writed by zyb 20110907
Public Function zlClobRead(ByVal Action As Long, ByVal KeyWord As String) As String
    'KeyWord:ID,'/ITEM/XH'����'/ITEM/MC'
    Dim lngCount As Long
    Dim StrText As String, strReturn As String
    Dim rsLob As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand
    
    gstrSQL = "Select Zl_Lob_Read([1],[2],[3],[4]) as Ƭ�� From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlClobRead", Action, KeyWord, lngCount, 0)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        
        StrText = rsLob.Fields(0).Value
        strReturn = strReturn & StrText
        lngCount = lngCount + 1
    Loop
    zlClobRead = strReturn
errHand:
End Function

Public Function zlClobSql(ByVal KeyWord As String, ByVal strFileContent As String, ByRef arySql() As String) As Boolean
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, StrText As String
    Dim lngLBound As Long, lngUBound As Long    '�����������С����±�
    
    Err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUBound = UBound(arySql)
    If Err <> 0 Then lngLBound = 0: lngUBound = -1
    Err = 0: On Error GoTo errHand
    
    lngFileSize = Len(strFileContent)
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    
    ReDim Preserve arySql(lngLBound To lngUBound + lngBlocks + 1) As String
    For lngCount = 0 To lngBlocks
        StrText = Mid(strFileContent, conChunkSize * lngCount + 1, conChunkSize)
        arySql(lngUBound + lngCount + 1) = "Zl_Lob_Append(21,'" & KeyWord & "','" & StrText & "'," & IIf(lngCount = 0, 1, 0) & ",1)"
    Next
    zlClobSql = True
    Exit Function

errHand:
    zlClobSql = False
End Function


'################################################################################################################
'## ���ܣ�  ��ָ�����ļ����浽ָ����¼��LOB�ֶ���
'##
'## ������  Action      :�������ͣ����������ǲ����ĸ���
'##         KeyWord     :ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'##         strFile     :�û�ָ����ŵ��ļ�������ָ��ʱ��ȡ��ǰ·�������ļ���
'##
'## ���أ�  �ɹ�����True��ʧ�ܷ���False
'##
'## ˵����  Actionȡֵ˵����
'##         0-�������ͼ�Σ�1-�����ļ���ʽ��2-�����ļ�ͼ�Σ�3-�������ĸ�ʽ��4-��������ͼ�Σ�5-���Ӳ�����ʽ��6-���Ӳ���ͼ�Σ�
'################################################################################################################
Public Function zlBlobSave(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String) As Boolean
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, StrText As String
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    Err = 0: On Error GoTo errHand
    
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        StrText = Join(aryHex, "")
        gstrSQL = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & StrText & "'," & IIf(lngCount = 0, 1, 0) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "zlBlobSave")
    Next
    Close lngFileNum
    zlBlobSave = True
    Exit Function

errHand:
    Close lngFileNum
    zlBlobSave = False
End Function

'################################################################################################################
'## ���ܣ�  ��������ָ�����ļ���ָ�����¼BLOB�ֶε�SQL���
'##
'## ������  Action      :�������ͣ����������ǲ����ĸ���
'##         KeyWord     :ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'##         strFile     :�û�ָ����ŵ��ļ�������ָ��ʱ��ȡ��ǰ·�������ļ���
'##         arySql()    :�ڸ����ݵĻ�������չ���ӱ����SQL��䣻��ָ��ʱ��ȡ��ǰ·�������ļ���
'##
'## ���أ�  �ɹ�����True��ʧ�ܷ���False
'##
'## ˵����  Actionȡֵ˵����
'##         0-�������ͼ�Σ�1-�����ļ���ʽ��2-�����ļ�ͼ�Σ�3-�������ĸ�ʽ��4-��������ͼ�Σ�5-���Ӳ�����ʽ��6-���Ӳ���ͼ�Σ�
'################################################################################################################
Public Function zlBlobSql(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String, ByRef arySql() As String) As Boolean
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, StrText As String
    
    Dim lngLBound As Long, lngUBound As Long    '�����������С����±�
    Err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUBound = UBound(arySql)
    If Err <> 0 Then lngLBound = 0: lngUBound = -1
    Err = 0: On Error GoTo errHand
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    
    ReDim Preserve arySql(lngLBound To lngUBound + lngBlocks + 1) As String
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        StrText = Join(aryHex, "")
        arySql(lngUBound + lngCount + 1) = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & StrText & "'," & IIf(lngCount = 0, 1, 0) & ")"
    Next
    Close lngFileNum
    zlBlobSql = True
    Exit Function

errHand:
    Close lngFileNum
    zlBlobSql = False
End Function

'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then zlFileUnzip = "": Exit Function
    
    strZipPath = gobjFSO.GetSpecialFolder(2) 'ȡ��ʱĿ¼
    strZipPathTmp = strZipPath & "\" & Format(Now, "yyMMdd") & CStr(100 * Timer)
    If Not gobjFSO.FolderExists(strZipPathTmp) Then Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp & "\TMP.RTF"
    If gobjFSO.FileExists(strZipFileTmp) Then gobjFSO.DeleteFile strZipFileTmp
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FileExists(strZipFileTmp) Then
        
        strZipFileName = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer) & ".RTF"
        If gobjFSO.FileExists(strZipFileName) Then gobjFSO.DeleteFile strZipFileName
                
        Call gobjFSO.CopyFile(strZipFileTmp, strZipFileName)
        
        If gobjFSO.FileExists(strZipFileTmp) Then gobjFSO.DeleteFile strZipFileTmp, True
        
        On Error Resume Next
        If gobjFSO.FolderExists(strZipPathTmp) Then gobjFSO.DeleteFolder strZipPathTmp, True
        
        zlFileUnzip = strZipFileName
    Else
        zlFileUnzip = ""
    End If
    
    Exit Function
    
errHand:
    Call SaveErrLog
End Function

'################################################################################################################
'## ���ܣ�  ���ļ�ѹ��Ϊ���ļ��ŵ���ͬĿ¼��
'## ������  strFile     :ԭʼ�ļ�
'## ���أ�  ѹ���ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    On Error GoTo errHand
    If Not gobjFSO.FileExists(strFile) Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = gobjFSO.GetParentFolderName(strFile) & "\ZLZIP" & lngCount & ".ZIP"
        If Not gobjFSO.FileExists(strZipFile) Then Exit Do
        lngCount = lngCount + 1
    Loop

    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'################################################################################################################
'## ���ܣ�  �滻����Ҫ�صĴ���
'##
'## ������  ElementName     :�滻��Ŀ������
'##         sPatientID      :����ID
'##         sPageID         :��ҳID��Һ�id
'##         iPatientType    :0=���1=סԺ
'##         lngҽ��ID       :ҽ��ID
'##
'## ���أ�  �����滻���
'################################################################################################################
Public Function GetReplaceEleValue(ByVal ElementName As String, _
    ByVal sPatientID As String, _
    ByVal sPageID As String, _
    ByVal iPatientType As PatiFromEnum, _
    ByVal lngҽ��id As Long, Optional lngBabyNum As Long) As String

    Dim rsTmp As New ADODB.Recordset
    
    If ElementName = "��λ����" Then
        GetReplaceEleValue = zl9ComLib.zlRegInfo("��λ����")
        Exit Function
    End If
    
    gstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4],[5],[6]) From Dual"
    Err = 0: On Error GoTo DBError
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�滻��", ElementName, CLng(sPatientID), _
        CLng(sPageID), CLng(iPatientType), lngҽ��id, lngBabyNum)
    If rsTmp.EOF Or rsTmp.BOF Then
        GetReplaceEleValue = ""
    Else
        GetReplaceEleValue = Trim(IIf(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value))
    End If
    Exit Function

DBError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

'################################################################################################################
'## ���ܣ�  �ж�ָ���û��Ƿ�������ҽʦ
'##
'## ������  lngUserID       ���û�ID
'##         strUserName     ���û���
'##         lngPatiID       ������ID
'##         lngPatiPageID   ����ҳID
'##
'## ˵����  ���ݡ���Ա���еġ�Ƹ�μ���ְ���ֶ�ȷ��ҽ����������סԺҽʦ������ҽʦ������ҽʦ��
'##         �����˱䶯��¼�е�ҽ�����𣬴Ӷ�ȷ����˼���
'################################################################################################################
Public Function GetUserSignLevel(lngUserID As Long, Optional strUserName As String, _
    Optional lngPatiID As Long, Optional lngPatiPageID As Long) As EPRSignLevelEnum
    Dim rs As New ADODB.Recordset, lngR As Long, lngLevel1 As Long, lngLevel2 As Long
    
    Err = 0: On Error GoTo errHand
    If InStr(gstrPrivsEpr, "ǩ��Ȩ") = 0 Then
        GetUserSignLevel = cprSL_�հ�
        Exit Function
    End If
    
    gstrSQL = "select Ƹ�μ���ְ�� from ��Ա�� p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngUserID)
    If Not rs.EOF Then
        lngR = NVL(rs("Ƹ�μ���ְ��"), 0)
    End If
    Select Case lngR    '1 ����  2 ����  3 �м�  4 ����/ʦ��  5 Ա/ʿ  9 ��Ƹ
    Case 1: lngLevel1 = cprSL_����
    Case 2: lngLevel1 = cprSL_����
    Case 3: lngLevel1 = cprSL_����
    Case Else: lngLevel1 = cprSL_����
    End Select
    rs.Close
    
    If lngPatiID > 0 Then
        gstrSQL = "Select ����ҽʦ, ����ҽʦ, ����ҽʦ " & _
            " From ���˱䶯��¼ " & _
            " Where ����ID = [1] And ��ҳID = [2] And (��ֹʱ�� Is Null Or ��ֹԭ�� = 1) " & _
            "       And ��ʼʱ�� Is Not Null And Nvl(���Ӵ�λ, 0) = 0"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngPatiID, lngPatiPageID)
        If rs.EOF Then
            lngLevel2 = cprSL_����
        Else
            If rs.Fields("����ҽʦ") = IIf(strUserName = "", gstrUserName, strUserName) Then
                lngLevel2 = cprSL_����
            ElseIf rs.Fields("����ҽʦ") = IIf(strUserName = "", gstrUserName, strUserName) Then
                lngLevel2 = cprSL_����
            Else
                lngLevel2 = cprSL_����
            End If
        End If
    End If
    GetUserSignLevel = IIf(lngLevel1 >= lngLevel2, lngLevel1, lngLevel2)
    Exit Function

errHand:
    GetUserSignLevel = cprSL_�հ�
End Function

Public Function GetColorVectorG(ByVal lngVersion As Long) As Long
    '���ݰ汾��ȡRGB��ɫ�е�G��ɫ����ֵ
    Select Case lngVersion
    Case 0
        GetColorVectorG = 0     'δ��ʼ
    Case 1
        GetColorVectorG = 0     '��һ�滹�����޶���
    Case 2
        GetColorVectorG = 10
    Case 3
        GetColorVectorG = 90
    Case 4
        GetColorVectorG = 140
    Case 5
        GetColorVectorG = 145
    Case 6
        GetColorVectorG = 150
    Case 7
        GetColorVectorG = 155
    Case 8
        GetColorVectorG = 160
    Case 9
        GetColorVectorG = 165
    Case 10
        GetColorVectorG = 170
    Case 11
        GetColorVectorG = 175
    Case 12
        GetColorVectorG = 180
    Case 13
        GetColorVectorG = 185
    Case 14
        GetColorVectorG = 190
    Case 15
        GetColorVectorG = 195
    Case 16
        GetColorVectorG = 200
    Case 17
        GetColorVectorG = 205
    End Select
End Function

Public Function GetColorVectorB(ByVal lngVersion As Long) As Long
    '���ݰ汾��ȡRGB��ɫ�е�B��ɫ����ֵ
    Select Case lngVersion
    Case 0
        GetColorVectorB = 0     'δ��ֹ
    Case 1
        GetColorVectorB = 0     '��һ�滹�����޶���
    Case 2
        GetColorVectorB = 10
    Case 3
        GetColorVectorB = 15
    Case 4
        GetColorVectorB = 20
    Case 5
        GetColorVectorB = 25
    Case 6
        GetColorVectorB = 30
    Case 7
        GetColorVectorB = 35
    Case 8
        GetColorVectorB = 40
    Case 9
        GetColorVectorB = 45
    Case 10
        GetColorVectorB = 50
    Case 11
        GetColorVectorB = 55
    Case 12
        GetColorVectorB = 60
    Case 13
        GetColorVectorB = 65
    Case 14
        GetColorVectorB = 70
    Case 15
        GetColorVectorB = 75
    Case 16
        GetColorVectorB = 80
    Case 17
        GetColorVectorB = 85
    End Select
End Function

Public Function Get��ʼ��(ByVal COLOR As OLE_COLOR) As Long
    '��ȡָ����ɫ�Ŀ�ʼ�汾�ţ�Ϊ0��ʾԭʼ�ı���ɫ
    Dim i As Long
    If COLOR = tomAutoColor Or COLOR = vbBlack Then COLOR = vbBlack: Get��ʼ�� = 1: Exit Function
    For i = 1 To 17
        If GetColorVectorG(i) = rgbGreen(COLOR) Then
            Get��ʼ�� = i
            Exit Function
        End If
    Next
    Get��ʼ�� = 1
End Function

Public Function Get��ֹ��(ByVal COLOR As OLE_COLOR) As Long
    '��ȡָ����ɫ����ֹ�汾�ţ�Ϊ0��ʾδ�����������ϴε���ɫֵ��
    Dim i As Long
    If COLOR = tomAutoColor Or COLOR = vbBlack Then COLOR = vbBlack: Get��ֹ�� = 0: Exit Function
    For i = 1 To 17
        If GetColorVectorB(i) = rgbBlue(COLOR) Then
            Get��ֹ�� = i - 1
            Exit Function
        End If
    Next
    Get��ֹ�� = 0
End Function

Public Function GetCharColor(ByVal lng��ʼ�� As Long, ByVal lng��ֹ�� As Long) As OLE_COLOR
    '���ݿ�ʼ�桢��ֹ���ȡ�����ַ���ɫ
    Dim r As Long, g As Long, b As Long
    r = 255
    g = GetColorVectorG(lng��ʼ��)
    b = GetColorVectorB(lng��ֹ��)
    If g = 0 And b = 0 Then
        GetCharColor = vbBlack
    Else
        GetCharColor = RGB(r, g, b)
    End If
End Function


Public Function GetMax(ByVal strTable As String, ByVal strField As String, ByVal intLength As Integer, Optional ByVal strWhere As String) As String
'���ܣ���ȡָ����ı�����������ֵ
'������strTable  ����;
'      strField  �ֶ���;
'      intLength �ֶγ���
'���أ��ɹ����� �¼�������; ���߷��� 0
    Dim rsTemp As New ADODB.Recordset
    Dim varTemp As Variant
    Dim lngLengh As Long
    
    On Error GoTo errHand
    gstrSQL = "SELECT MAX(LPAD(" & strField & "," & intLength & ",' ')) as ""���ֵ"",max(length(" & _
         strField & ")) as ""�ֵ"" FROM " & strTable & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR")
    With rsTemp
        If rsTemp.EOF Then
            GetMax = Format(1, String(intLength, "0"))
            Exit Function
        End If
        varTemp = IIf(IsNull(.Fields("���ֵ").Value), "0", .Fields("���ֵ").Value)
        lngLengh = IIf(IsNull(.Fields("�ֵ").Value), intLength, .Fields("�ֵ").Value)
        If IsNumeric(varTemp) Then
            GetMax = CStr(Val(varTemp) + 1)
            GetMax = Format(GetMax, String(lngLengh, "0"))
        Else
            gstrSQL = "Select ZL_INCSTR([1]) As MAXVALUE From Dual"
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", CStr(varTemp))
            If rsTemp.BOF = False Then
                GetMax = Trim(rsTemp("MAXVALUE").Value)
            End If
        End If
        .Close
    End With
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function CommandBarInit(ByRef cbsMain As CommandBars) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    CommandBarInit = True
    
End Function

Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    DockPannelInit = True
    
End Function

Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, Optional ByVal objPrnVsf As Object, Optional ByVal strPrintTitle As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    Dim bytMode As Byte
        
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet              '��ӡ����
    
        Call zlPrintSet
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel               '��ӡ����,Ԥ������,�����Excel
        
'        If objPrnVsf Is Nothing Then Exit Function
'
'        Call SearchPrintData(objPrnVsf, frmPubResource.msfPrint)
'
'        '���ô�ӡ��������
'        Set objPrint.Body = frmPubResource.msfPrint
'        objPrint.Title.Text = strPrintTitle
'        Set objAppRow = New zlTabAppRow
'        Call objAppRow.Add("")
'        Call objAppRow.Add("��ӡʱ��:" & Now())
'        Call objPrint.BelowAppRows.Add(objAppRow)
'
'        Select Case Control.ID
'        Case conMenu_File_Print
'            bytMode = zlPrintAsk(objPrint)
'            If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
'        Case conMenu_File_Preview
'            zlPrintOrView1Grd objPrint, 2
'        Case conMenu_File_Excel
'            zlPrintOrView1Grd objPrint, 3
'        End Select
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '������
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text      '��ť����
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.STYLE = IIf(objControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size      '��ͼ��
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar         '״̬��
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
    
    Case conMenu_Help_Help              '��������
    
        Call ShowHelp(App.ProductName, frmMain.hWnd, frmMain.Name, Int((glngSys) / 100))
        
    Case conMenu_Help_Web_Home          'Web�ϵ�����
        
        Call zlHomePage(frmMain.hWnd)
        
    Case conMenu_Help_Web_Forum         'Web�ϵ���̳
    
        Call zlWebForum(frmMain.hWnd)
        
    Case conMenu_Help_Web_Mail          '���ͷ���
        
        Call zlMailTo(frmMain.hWnd)
            
    Case conMenu_Help_About             '����
        
        Call ShowAbout(frmMain, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    
    Case conMenu_File_Exit              '�˳�
    
        Unload frmMain
            
    End Select
    
    CommandBarExecutePublic = True
End Function

Public Function CommandBarUpdatePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Select Case Control.ID
    Case conMenu_View_ToolBar_Button            '������
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              'ͼ������
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).STYLE = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '��ͼ��
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '״̬��
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function

Public Function NewCommandBar(objMenu As CommandBarControl, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngId As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal strParameter As String) As CommandBarControl
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpType, lngId, strCaption)
        
        objControl.IconId = IIf(lngIcon = -1, lngId, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        objControl.Parameter = strParameter
        
    End With
    
    Set NewCommandBar = objControl
    
End Function

Public Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngId As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer) As CommandBarControl
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngId, strCaption, intBefore)
        objControl.ID = lngId
        objControl.IconId = IIf(lngIcon = -1, lngId, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.STYLE = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '���ܣ������ַ����ļ���
    '��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
    '���Σ���ȷ�����ַ��������󷵻�"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    
    If bytIsWB Then
        gstrSQL = "Select zlWBcode('" & strInput & "') from dual"
    Else
        gstrSQL = "Select zlSpellcode('" & strInput & "') from dual"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR")
    zlGetSymbol = NVL(rsTmp.Fields(0).Value)
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function

Public Function IsPrivs(ByVal strPrivs As String, ByVal strPriv As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function
Public Function Get�����ļ�ID(ByVal lngRecordId As Long, ByVal lngAdviceID As Long) As String
Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    If lngAdviceID <> 0 Then
        gstrSQL = "Select ����id From ����ҽ������ A Where ҽ��id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ͬҽ��ID�Ļ����¼", lngAdviceID)
    Else
        gstrSQL = "Select b.����id From ����ҽ������ A, ����ҽ������ B Where a.����id = [1] And a.ҽ��id = b.ҽ��id"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ͬҽ��ID�Ļ����¼", lngRecordId)
    End If
    
    Do Until rsTemp.EOF
        Get�����ļ�ID = Get�����ļ�ID & "," & rsTemp!����ID
        rsTemp.MoveNext
    Loop
    If Len(Get�����ļ�ID) > 0 Then
        Get�����ļ�ID = Mid(Get�����ļ�ID, 2)
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function GetFileRange(ByVal lFileId As Long, ByVal lngRecordId As Long, ByVal strCreateTime As String, _
                            ByVal eDocType As EPRDocTypeEnum, ByVal lngPatId As Long, ByVal lngPageId As Long, _
                            Optional ByVal blnMoved As Boolean, Optional ByVal lngAdviceID As Long) As String
    '******************************************************************************************************************
    '���ܣ���ȡ��ǰ����(���ܵ�ǰ����δ����)ǰ���й�����
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTemp As New ADODB.Recordset, strTime As String, dStar As Date, dEnd As Date
    Dim strSQL As String, strIDs As String, blnNewPage As Boolean, n_Num As Integer, n_S As Integer, n_E As Integer

    On Error GoTo errHand
    strIDs = Get�����ļ�ID(lngRecordId, lngAdviceID)
    If strIDs <> "" Then GetFileRange = strIDs: Exit Function
    
    strTime = Format(strCreateTime, "yyyy-MM-dd HH:mm:ss")
    If strTime = "" Then strTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If strTime = "00:00:00" Then strTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    blnNewPage = zlDatabase.GetPara("ת�ƺ�Ҫ����д�Ĺ���������һҳ��ӡ", glngSys, 1251, 1) = 1 '=0��ʾת�ƺ�����������ӡ =1 ��ʾ����һҳ��ӡ

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
        If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���", lngRecordId)
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
    If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҳ���ļ�֮ǰһ����д", lFileId, lngPatId, lngPageId, CDate(strTime), n_Num)
    If rsTemp.EOF Then
        dStar = CDate("2000-01-01 00:00:00"): n_S = -1 '����֮ǰû��д��ҳ���ļ�
    Else
        dStar = CDate(rsTemp!����ʱ��)
        n_S = IIf(n_Num = 0, 0, NVL(rsTemp!���, -1))
    End If
    
    '��ȡҳ���ļ���ǰʱ��֮�����һ����д��¼
    gstrSQL = "Select ����ʱ��,���" & vbNewLine & _
                "From (Select a.����ʱ��,a.���" & vbNewLine & _
                "       From ���Ӳ�����¼ A, (" & strSQL & ") B" & vbNewLine & _
                "       Where a.�ļ�id = b.Id And a.����id = [2] And a.��ҳid = [3] And " & IIf(n_Num = 0, "a.����ʱ�� > [4]", "((a.����ʱ�� > [4] and Nvl(a.���,0)=0) Or a.���>[5])") & vbNewLine & _
                "       Order By a.����ʱ��)" & vbNewLine & _
                "Where Rownum = 1"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҳ���ļ�֮��һ����д", lFileId, lngPatId, lngPageId, CDate(strTime), n_Num)
    If rsTemp.EOF Then
        dEnd = CDate("3001-01-01") '����֮��û��д��ҳ���ļ�
        n_E = 9999
    Else
        dEnd = CDate(rsTemp!����ʱ��) - 1 / 24 / 60 / 60 '����֮��д����ȡ����¼��ʱ���һ��,���������˼�¼
        n_E = IIf(n_Num = 0, 0, NVL(rsTemp!���, 9999) - 1)
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
    If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҳ���ļ�֮��һ����д", lFileId, lngPatId, lngPageId, dStar, dEnd, n_S, n_E)
    Do Until rsTemp.EOF
        strIDs = strIDs & "," & rsTemp!ID
        rsTemp.MoveNext
    Loop
    strIDs = Mid(strIDs, 2)
    
    GetFileRange = strIDs
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function SetPaneRange(dkpMain As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objPan As Pane
    
    On Error Resume Next
    
    Set objPan = dkpMain.FindPane(intPane)
    
    If objPan Is Nothing Then Exit Function
    With objPan
        .MaxTrackSize.SetSize lngMaxW, lngMaxH
        .MinTrackSize.SetSize lngMinW, lngMinH
    End With
    
    SetPaneRange = True
End Function

Public Function CreateTmpFile(Optional ByVal strFileType As String = "tmp", Optional ByVal strName As String, Optional ByVal blnTime As Boolean = True) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strFile As String
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    strFileTemp = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
    
    If blnTime Then
        strFileTemp = strFileTemp & strName & Format(Now, "yyyymmdd") & Format(Timer, "0") & "." & strFileType
    Else
        strFileTemp = strFileTemp & strName & "." & strFileType
    End If
    
    CreateTmpFile = strFileTemp
    
End Function

Public Function RemoveSign(ByRef edtThis As Editor, ByRef objDocument As cEPRDocument) As Boolean
'******************************************************************************************************************
'���ܣ��Ӵ�ӡ/Ԥ���ĵ����Ƴ�ǩ�����ݼ���ǰ׺
'******************************************************************************************************************
Dim intLoop As Integer, strFoot As String, strHead As String
Dim lESS As Long, lESE As Long, lEES As Long, lEEE As Long, blnNeeded As Boolean, blnFinded As Boolean
Dim strAllSign As String, strFSign As String, strSSign As String, strTSign As String
    
    On Error GoTo errHand
    strFoot = edtThis.FootFileText: strHead = edtThis.HeadFileText
    edtThis.ForceEdit = True
    If InStr(strFoot, "{��дǩ��}") > 0 Or InStr(strFoot, "{ҽ��ǩ��}") > 0 Or InStr(strFoot, "{����ǩ��}") > 0 Or InStr(strFoot, "{����ǩ��}") > 0 Or _
        InStr(strHead, "{��дǩ��}") > 0 Or InStr(strHead, "{ҽ��ǩ��}") > 0 Or InStr(strHead, "{����ǩ��}") > 0 Or InStr(strHead, "{����ǩ��}") > 0 Then
        '���Ҳ�����ԭ����ǩ��
        For intLoop = 1 To objDocument.Signs.Count
            blnFinded = False
            blnFinded = FindKey(edtThis, "S", objDocument.Signs(intLoop).Key, lESS, lESE, lEES, lEEE, blnNeeded)
            If blnFinded Then
                Select Case objDocument.Signs(intLoop).ǩ������
                    Case Is <= cprSL_����
                        strFSign = strFSign & " " & edtThis.Range(lESS + 16, lEES).Text
                    Case cprSL_����
                        strSSign = strSSign & " " & edtThis.Range(lESS + 16, lEES).Text
                    Case Is >= cprSL_����
                        strTSign = strTSign & " " & edtThis.Range(lESS + 16, lEES).Text
                End Select
                edtThis.Range(lESE, lEES).Text = ""
            End If
        Next
        
'        For intLoop = 1 To objDocument.Elements.Count
'            If objDocument.Elements(intLoop).�滻�� = 1 Then
'                Select Case objDocument.Elements(intLoop).Ҫ������
'                Case "����ҽʦǩ��"
'                    strFSign = strFSign & " " & objDocument.Elements(intLoop).�����ı�
'                Case "����ҽʦǩ��"
'                    strSSign = strSSign & " " & objDocument.Elements(intLoop).�����ı�
'                Case "����ҽʦǩ��"
'                    strTSign = strTSign & " " & objDocument.Elements(intLoop).�����ı�
'                End Select
'            End If
'        Next
    End If
    edtThis.ForceEdit = False
    strFSign = Mid(strFSign, 2): strSSign = Mid(strSSign, 2): strTSign = Mid(strTSign, 2)
    strAllSign = strFSign & " " & strSSign & " " & strTSign
    objDocument.EPRPatiRecInfo.ҽ��ǩ�� = strFSign
    objDocument.EPRPatiRecInfo.����ǩ�� = strSSign
    objDocument.EPRPatiRecInfo.����ǩ�� = strTSign
    objDocument.EPRPatiRecInfo.��дǩ�� = strAllSign
    RemoveSign = True
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SetHeadFoot(edtThis As Editor, ByVal lngFileID As Long)
'�����ݿ��ȡ��¼ˢ��ҳüҳ��
'��ʽ=PaperKind;PaperOrient;PaperHeight;PaperWidth;MarginLeft;MarginRight;MarginTop;MarginBottom;BackColor;PaperColor;ShowPageNumber;ҳü��ʽ;ҳ�Ÿ�ʽ

Dim strFile As String, lngType As Long, strPage As String, rsTemp As New ADODB.Recordset
    gstrSQL = "Select a.����, a.���, a.��ʽ, a.ҳü, a.ҳ��" & vbNewLine & _
                "From ����ҳ���ʽ A, �����ļ��б� B" & vbNewLine & _
                "Where b.Id = [1] And a.���� = b.���� And b.ҳ�� = a.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҳüҳ��", lngFileID)
    If rsTemp.EOF Then Exit Sub
    If NVL(rsTemp!��ʽ) = "" Then Exit Sub
    
    With edtThis
        .PaperKind = Split(rsTemp!��ʽ, ";")(0)
        .PaperOrient = Split(rsTemp!��ʽ, ";")(1)
        If UBound(Split(rsTemp!��ʽ, ";")) > 10 Then
        .HeadFontFormat = Split(rsTemp!��ʽ, ";")(11)
        .FootFontFormat = Split(rsTemp!��ʽ, ";")(12)
        End If
        .PaperHeight = Split(rsTemp!��ʽ, ";")(2)
        .PaperWidth = Split(rsTemp!��ʽ, ";")(3)
        .MarginLeft = Split(rsTemp!��ʽ, ";")(4)
        .MarginRight = Split(rsTemp!��ʽ, ";")(5)
        .MarginTop = Split(rsTemp!��ʽ, ";")(6)
        .MarginBottom = Split(rsTemp!��ʽ, ";")(7)
    
        strFile = zlBlobRead(7, rsTemp!���� & "-" & rsTemp!���) '��ȡҳüͼƬ
        If gobjFSO.FileExists(strFile) Then
            Set .Picture = LoadPicture(strFile)
            gobjFSO.DeleteFile strFile, True      'ɾ����ʱ�ļ�
        End If
        
        strFile = zlBlobRead(12, rsTemp!���� & "-" & rsTemp!���, App.Path & "\Head.rtf") '��ȡҳü�ļ�
        If gobjFSO.FileExists(strFile) Then
            edtThis.HeadFile = strFile           '��ȡ�ļ�
            gobjFSO.DeleteFile strFile, True      'ɾ����ʱ�ļ�
            If Trim(edtThis.HeadFileText) = "" Then GoTo Headtxt
        Else
Headtxt:
            If NVL(rsTemp!ҳü) <> "" Then
                edtThis.Head = rsTemp!ҳü
                edtThis.HeadTextToFile '�����ֶ���Rtf�ؼ���
            End If
        End If
        
        strFile = zlBlobRead(13, rsTemp!���� & "-" & rsTemp!���, App.Path & "\Foot.rtf") '��ȡҳ���ļ�
        If gobjFSO.FileExists(strFile) Then
            edtThis.FootFile = strFile            '��ȡ�ļ�
            gobjFSO.DeleteFile strFile, True      'ɾ����ʱ�ļ�
            If Trim(edtThis.FootFileText) = "" Then GoTo Foottxt
        Else
Foottxt:
            If NVL(rsTemp!ҳ��) <> "" Then
                edtThis.Foot = rsTemp!ҳ��
                edtThis.FootTextToFile '�����ֶ���Rtf�ؼ���
            End If
        End If
    End With
End Sub
Public Sub ReadRTF(edtThis As Editor, ByVal lngFileID As Long, ByVal blnClearMode As Boolean, ByVal blnMoved As Boolean, Optional ByVal blnClearBColor As Boolean = True)
'��ȡRTF�ļ�
Dim ParaFmt As New cParaFormat, FontFmt As New cFontFormat, i As Long, lngLen As Long
Dim oEles As New cEPRElements, oTabs As New cEPRTables, oPics As New cEPRPictures
Dim strZipFile As String, strRtfFile As String, j As Long, rs As New ADODB.Recordset
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    strZipFile = zlBlobRead(5, lngFileID, , blnMoved)
    If gobjFSO.FileExists(strZipFile) Then
        strRtfFile = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strRtfFile) Then
                edtThis.OpenDoc strRtfFile
                gobjFSO.DeleteFile strRtfFile, True
        End If
        gobjFSO.DeleteFile strZipFile, True
    End If
    If Trim(edtThis.Text) = "" Then Exit Sub

    '��ȡͼƬ,���,Ҫ��
    gstrSQL = "Select Level,ID, �ļ�id,��ʼ��, ��ֹ��," & _
                "   ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������id, �������," & vbNewLine & _
                "       ʹ��ʱ��, ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��" & vbNewLine & _
                "From (Select ID, �ļ�id,��ʼ��, ��ֹ��," & _
                "               ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id,�������id," & vbNewLine & _
                "              �������, ʹ��ʱ��, ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��" & vbNewLine & _
                "       From ���Ӳ�������" & vbNewLine & _
                "       Where �ļ�id = [1] And �������<>ID)" & vbNewLine & _
                "Start With ��id Is Null" & vbNewLine & _
                "Connect By Prior ID = ��id" & vbNewLine & _
                "Order By �������, �����д�"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�������", "H���Ӳ�������")
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ԭ��ͼ", lngFileID)
    Do Until rs.EOF
        Select Case rs!��������
            Case 3  '���
                lKey = oTabs.Add(NVL(rs!������, 0))                  '�ָ�Keyֵ��
                Call oTabs("K" & lKey).FillTableMember(rs, IIf(blnMoved, "H���Ӳ�������", "���Ӳ�������"))
            Case 4  'Ҫ��
                lKey = oEles.Add(NVL(rs!������, 0))
                Call oEles("K" & lKey).FillElementMember(rs, IIf(blnMoved, "H���Ӳ�������", "���Ӳ�������"))
            Case 5  'ͼƬ
                lKey = oPics.Add(NVL(rs("������"), 0))
                Call oPics("K" & lKey).FillPictureMember(rs, IIf(blnMoved, "H���Ӳ�������", "���Ӳ�������"))
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

    ' �������ĵ��Ĵ���
    edtThis.SelectAll
    If blnClearMode Then
        edtThis.AuditMode = True
        edtThis.AcceptAuditText    '���ģʽ
    End If
    
    If blnClearBColor Then
        For j = 1 To oEles.Count 'ɾ���յ�Ҫ�أ�չ����Ҫ������ˢ����ȥ���²�����
            If oEles(j).�����ı� = "" Then
                oEles(j).DeleteFromEditor edtThis
            ElseIf oEles(j).������̬ = 1 Then
                oEles(j).Refresh edtThis
            End If
        Next
        
        lngLen = Len(edtThis.Text)
        For i = 0 To lngLen - 1 'ֻ������ɫΪҪ�ر���ɫ��ɫȥ��
            If edtThis.Range(i, i + 1).Font.BackColor = ELE_BACKCOLOR Then
                edtThis.Range(i, i + 1).Font.BackColor = tomAutoColor
            End If
            If edtThis.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR Then
                edtThis.Range(i, i + 1).Font.ForeColor = tomAutoColor
            End If
        Next
    End If
End Sub
Public Sub BuildRTF(edtThis As Editor, ByVal lngFileID As Long, ByVal blnMoved As Boolean)
'��RTF�ļ�ʱ����ȡ���Ӳ������ݽ�����ʾ���޷��༭
Dim strContent As String, rs As New ADODB.Recordset
    gstrSQL = "Select ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���" & vbNewLine & _
            "From (Select ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���" & vbNewLine & _
            "       From ���Ӳ�������" & vbNewLine & _
            "       Where �ļ�id = [1] And ������� <> ID And ��ֹ�� = 0)" & vbNewLine & _
            "Start With ��id Is Null" & vbNewLine & _
            "Connect By Prior ID = ��id" & vbNewLine & _
            "Order By �������, �����д�"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�������", "H���Ӳ�������")
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", lngFileID)
        
    Do Until rs.EOF
        If rs!�������� <> 1 Then '��ٲ�������ʾ
            strContent = strContent & IIf(rs!�Ƿ��� = 1, vbCrLf, "") & rs!�����ı�
        End If
        rs.MoveNext
    Loop
    
    edtThis.Text = strContent
End Sub
Public Sub ReplacedHeadFootString(ByRef edtThis As Object, ByVal lngRecId As Long, ByVal blnMoved As Boolean)
'���ܣ� ҳü/ҳ���е��滻Ҫ������
'������ objDoc�������������Long,��ô������ǵ��Ӳ�����¼ID
Dim strElements As String, j As Long, aryEle() As String, strEleValue As String
Dim lngStartPos As Long, lngEndPos As Long
Dim strHead As String, strFoot As String
Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select a.��������, a.���ʱ��, a.����id, a.��ҳid, a.������Դ, b.���� ��д����, c.�����ı� As ��дǩ��, d.ҽ��id" & vbNewLine & _
            "From ���Ӳ�����¼ A, ���ű� B, ���Ӳ������� C, ����ҽ������ D" & vbNewLine & _
            "Where a.Id =[1] And a.Id = c.�ļ�id(+) And a.����id = b.Id And c.��������(+) = 8 And c.��ʼ��(+) = 1 And a.Id = d.����id(+)"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�������", "H���Ӳ�������")
    If blnMoved Then gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ϣ", lngRecId)
    
    '��strHead�з������滻Ҫ�أ����ŵ�aryEle������
    '------------------------------------------------------------------------------------------------------------------
    strHead = edtThis.HeadFileText
    lngStartPos = 0
    lngEndPos = 0
    strElements = ""
    For j = 1 To Len(strHead)
        If Mid(strHead, j, 1) = "{" Then lngStartPos = j
        If Mid(strHead, j, 1) = "}" Then lngEndPos = j

        If lngStartPos > 0 And lngEndPos > 0 Then
            If lngEndPos > lngStartPos + 1 Then
                strElements = strElements & ";" & Mid(strHead, lngStartPos + 1, lngEndPos - lngStartPos - 1)
            End If

            lngStartPos = 0
            lngEndPos = 0
        End If
    Next

    '��strFoot�з������滻Ҫ�أ����ŵ�aryEle������
    '------------------------------------------------------------------------------------------------------------------
    strFoot = edtThis.FootFileText
    lngStartPos = 0
    lngEndPos = 0
    For j = 1 To Len(strFoot)
        If Mid(strFoot, j, 1) = "{" Then lngStartPos = j
        If Mid(strFoot, j, 1) = "}" Then lngEndPos = j

        If lngStartPos > 0 And lngEndPos > 0 Then
            If lngEndPos > lngStartPos + 1 Then
                strElements = strElements & ";" & Mid(strFoot, lngStartPos + 1, lngEndPos - lngStartPos - 1)
            End If

            lngStartPos = 0
            lngEndPos = 0
        End If
    Next
    If strElements <> "" Then
        strElements = Mid(strElements, 2)
    Else
        Exit Sub
    End If
    aryEle = Split(strElements, ";")
    
    '����ҽ��ǩ��������ǩ��������ǩ����ֻ����������Ʊ����У���Ҫ��ϵ��Ӳ����������Edit�е�ǩ����Ϣ��ת�Ƶ�ҳüҳ����
    'Ŀǰ�򱾺���ֻ�����ڹ��������ݲ�����
    
    For j = 0 To UBound(aryEle)
        Select Case aryEle(j)
            Case "��������"
                strEleValue = rsTemp!��������
            Case "��д����"
                strEleValue = rsTemp!��д����
            Case "���ʱ��"
                strEleValue = Format(NVL(rsTemp!���ʱ��), "yyyy-MM-dd hh:mm")
            Case "��дǩ��"
                strEleValue = NVL(rsTemp!��дǩ��)
            Case "ҳ��", "��ҳ��", "����", "�ļ���", "·��", "��ӡ����", "��ӡʱ��"
                strEleValue = "" '�����ݹؼ����ɿؼ��ڲ��滻
            Case Else
                strEleValue = GetReplaceEleValue(aryEle(j), rsTemp!����ID, rsTemp!��ҳID, NVL(rsTemp!������Դ, 0), NVL(rsTemp!ҽ��id, 0))
        End Select
        
        If strEleValue <> "" Then 'ȡ��ֵ,�滻ֵ
            Call edtThis.DocHeadReplaceKey("{" & aryEle(j) & "}", strEleValue)
            Call edtThis.DocFootReplaceKey("{" & aryEle(j) & "}", strEleValue)
        End If
    Next
End Sub

'################################################################################################################
'## ���ܣ�  ���ļ�ѹ��Ϊ���ļ��ŵ���ͬĿ¼�в�ɾ��XML�ļ�
'## ������  strFiles     :ԭʼ�ļ�·���ַ�����������ԡ������ָ�����
'## ������  strZipPath   :ѹ������ļ�·��
'## ���أ�  ѹ���ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFilesZip(ByVal strFiles As String, ByVal strZipPath As String) As String
    Dim strZipFile As String, strFile As Variant
    Dim lngFileNum As Long, lngFile As Long, i As Long, j As Long
    Dim aryChunk() As Byte, bytTmp As Byte
    On Error GoTo errHand:
    strFile = Split(strFiles, ",")
        With mclsZip
            .Encrypt = False: .AddComment = False
            .ZipFile = strZipPath
            .StoreFolderNames = False
            .RecurseSubDirs = False
            .ClearFileSpecs
            For i = 0 To UBound(strFile)
              .AddFileSpec strFile(i)
            Next i
            .Zip
            If (.Success) Then
                zlFilesZip = .ZipFile
            Else
                zlFilesZip = ""
            End If
            'ɾ��XML�ļ�
            For i = 0 To UBound(strFile)
                gobjFSO.DeleteFile (strFile(i))
            Next i
        End With
        Exit Function
errHand:
        zlFilesZip = ""
End Function
'################################################################################################################
'## ���ܣ�  ��XMLѹ���ļ���ͬĿ¼�ͷŲ�����ѹXML�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFilesUnZip(ByVal strZipFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String, strZipFileTmp2 As String
    Dim strZipFileName As String, strUnZipFile As File, strUnZipFileName As String
    Dim lngFileNum As Long   ' ����������
    Dim aryChunk() As Byte, lngFile As Long, bytTmp As Byte
    Dim i As Long
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then zlFilesUnZip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = zl9ComLib.OS.TempPath
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp & "\TMP.XML"
    If gobjFSO.FileExists(strZipFileTmp) Then gobjFSO.DeleteFile strZipFileTmp
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    For Each strUnZipFile In gobjFSO.GetFolder(strZipPathTmp).Files
        If InStr(1, strUnZipFile.Name, "�����б�") > 0 Then
            strUnZipFileName = strZipPathTmp & "\" & strUnZipFile.Name
        End If
        If InStr(1, strUnZipFile.Name, ".xml") > 0 And InStr(1, strUnZipFile.Name, "�����б�") < 1 Then
            strUnZipFileName = strZipPathTmp & "\" & strUnZipFile.Name
        End If
        If InStr(1, strUnZipFile.Name, "TMP") > 0 Then
            'ɾ����ѹ���.ZIP�ļ�
            gobjFSO.DeleteFile (strZipPathTmp & "\" & strUnZipFile.Name)
        End If
    Next
    zlFilesUnZip = strUnZipFileName
    Exit Function
errHand:
    Call SaveErrLog
End Function
Public Sub VerifyPatiSign(ByVal frmParent As Object, ByVal lFileId As Long, ByVal blnMoved As Boolean)
    '����:���ݴ���Ĳ����ļ�ID��ȡ����ǩ�������Ϣ��������֤�ӿ�
Dim rsTemp As New ADODB.Recordset, lngSignID As Double
Dim strSource As String, strName As String, strIdentifyNo As String, strOtherParms As String, strSignInfo As String, strPenSignBase64 As String
    On Error GoTo errHand
    gstrSQL = "Select ID,������,��������,�����ı�" & vbNewLine & _
                "From ���Ӳ�������" & vbNewLine & _
                "Where �ļ�id = [1] And �������� = 5"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�������", "H���Ӳ�������")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ǩ����¼", lFileId)
    If rsTemp.EOF Then MsgBox "��ǰ��ѡ����û����Ч������ǩ����¼��", vbInformation, gstrSysName: Exit Sub
    If Split(rsTemp!��������, ";")(0) <> 5 Then MsgBox "��ǰ��������ǩ����������֤��", vbInformation, gstrSysName: Exit Sub
    
    If UBound(Split(rsTemp!�����ı�, "|")) > 2 Then
        strName = Split(rsTemp!�����ı�, "|")(0)
        strIdentifyNo = Split(rsTemp!�����ı�, "|")(1)
        strSignInfo = rsTemp!�����ı�
    End If
    strSource = GetPatiSignSource(lFileId, blnMoved)
    If strSource = "" Then Exit Sub
    
    If gobjESign Is Nothing Then
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        Call gobjESign.Initialize(gcnOracle, glngSys)
    End If
    If gobjESign.EnabledVerifyPatiSign() = False Then
        MsgBox "��ǰ�ӿڲ�֧�ֻ���ǩ����֤��", vbInformation, gstrSysName
    End If
    Call gobjESign.ValidatePenSignature(strSource, strName, strIdentifyNo, strOtherParms, strSignInfo, strPenSignBase64)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetPatiSignSource(ByVal lFileId As Long, ByVal blnMoved As Boolean) As String
'����:�����ļ�ID,��ȡǩ��Դ����,�����ڲ��򿪱༭������½���ǩ����֤
'����:1 ȥ�����һ��ǩ�����ڵ�ǩ��ͼƬ (�����) ͼƬ�������Ϊ " PS(0000000X,0,1) PE(0000000X,0,1) "
'     2 ȥ������S�ؼ��ֵ�ǩ������
'     3 ������ͼƬ,���ؼ����м��"��"�ֻ��ɿո�(��Ϊǩ��ʱ�ǿո��ʾ)
'     4 ��ǩ��Ҫ�ػ�ԭ��ǩ��ʱ��"{����ҽʦǩ��}""����ҽʦǩ��""����ҽʦǩ��"
    Dim lSS As Long, lSE As Long, lES As Long, lEE As Long, bNeeded As Boolean, bFinded As Boolean, lKey As Long, lPos As Long
    Dim strZipFile As String, strRtfFile As String, lSEKey As Long
    Dim rsTemp As New ADODB.Recordset, strSource As String
    
    strZipFile = zlBlobRead(5, lFileId, , blnMoved)
    If gobjFSO.FileExists(strZipFile) Then
        strRtfFile = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strRtfFile) Then
                gfrmPublic.edtBuff.OpenDoc strRtfFile
                gobjFSO.DeleteFile strRtfFile, True
        End If
        gobjFSO.DeleteFile strZipFile, True
    End If
    
    gfrmPublic.edtBuff.Freeze
    gfrmPublic.edtBuff.ForceEdit = True
    
    'ȥ������S�ؼ��ֵ�ǩ������
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "S", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then
            If gfrmPublic.edtBuff.Range(lEE, lEE + 4).Text = "PS(" Then
                gfrmPublic.edtBuff.Range(lEE, lEE + 35).Text = "" 'ǩ�����ڵ�ǩ��ͼƬ,ǩ��ͼƬ����ǩ���ؼ���
            End If
            gfrmPublic.edtBuff.Range(lSS, lEE).Text = ""
        End If
    Loop Until bFinded = False
    
    '�����б��ؼ����м��"��"�ֻ��ɿո�(��Ϊǩ��ʱ�ǿո��ʾ)
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "T", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSE, lES).Text = " ": lPos = lEE + 1
    Loop Until bFinded = False
    
    '������ͼƬ�ؼ����м��"��"�ֻ��ɿո�(��Ϊǩ��ʱ�ǿո��ʾ)
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "P", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSE, lES).Text = " ": lPos = lEE + 1
    Loop Until bFinded = False
    
    '����ǩ��Ҫ��,��Ϊǩ��ʱʹ�õ�Դ����ǩ��Ҫ����"{����ҽʦǩ��}"��ʽ����ǩ���󱻸���Ϊ���������
    gstrSQL = "Select ������,Ҫ������" & vbNewLine & _
            "From ���Ӳ�������" & vbNewLine & _
            "Where �ļ�id = [1] And �������� = 4 And Ҫ������ In ('����ҽʦǩ��', '����ҽʦǩ��', '����ҽʦǩ��')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ��Ҫ������", lFileId)
    Do Until rsTemp.EOF  '����ʹ��ǩ��Ҫ�����ƻ�ԭ
        lPos = 0
        lSEKey = rsTemp!������
        bFinded = FindKey(gfrmPublic.edtBuff, "E", lSEKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then
            If gfrmPublic.edtBuff.Range(lEE, lEE + 4).Text = " PS(" Then 'ʹ����ǩ��ͼƬ
                gfrmPublic.edtBuff.Range(lEE, lEE + 35).Text = ""
            End If
            gfrmPublic.edtBuff.Range(lSE, lES).Text = "{" & rsTemp!Ҫ������ & "}"
        End If
        rsTemp.MoveNext
    Loop
    
    gfrmPublic.edtBuff.ForceEdit = False
    gfrmPublic.edtBuff.UnFreeze
    strSource = gfrmPublic.edtBuff.Text
    strSource = Replace(strSource, Chr(32), "")
    strSource = Replace(strSource, vbCr, "")
    strSource = Replace(strSource, vbLf, "")
    GetPatiSignSource = strSource
End Function

Public Sub VerifySignature(ByVal frmParent As Object, ByVal lFileId As Long, ByVal blnMoved As Boolean)
'����:���ݴ���Ĳ����ļ�ID��ȡǩ����¼ԭʼ��Ϣ��������֤�ӿ�
Dim rsTemp As New ADODB.Recordset, lngSignID As Double, strSource As String
    On Error GoTo errHand
    gstrSQL = "Select ID,������,��������" & vbNewLine & _
                "From (Select ID, ������,�������� From ���Ӳ������� Where �ļ�id = [1] And �������� = 8 Order By ������ Desc)" & vbNewLine & _
                "Where Rownum = 1"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�������", "H���Ӳ�������")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ǩ����¼", lFileId)
    If rsTemp.EOF Then MsgBox "��ǰ��ѡ����û����Ч������ǩ����¼��", vbInformation, gstrSysName: Exit Sub
    If Split(rsTemp!��������, ";")(0) <> 2 Then MsgBox "��ǰ����ǩ����������ǩ����������֤��", vbInformation, gstrSysName: Exit Sub
    
    lngSignID = rsTemp!ID
    Select Case Split(rsTemp!��������, ";")(1)
        Case 1
            strSource = GetSignSourceFromRTF1(lFileId, lngSignID, rsTemp!������, blnMoved)
        Case 2
            strSource = GetSignSourceFromRTF2(lFileId, lngSignID, rsTemp!������, blnMoved)
        Case 3 '��ʾԴ����ɷ�ʽ�����ݿ���ݰ�������ı���
            Dim frmSVerify As New frmEPRSignVerify 'ʹ���µ�Դ����ɷ�ʽ
            Call frmSVerify.ShowMe(frmParent, lFileId)
            Unload frmSVerify: Set frmSVerify = Nothing
    End Select
    
    If strSource = "" Then Exit Sub
    If gobjESign Is Nothing Then
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        Call gobjESign.Initialize(gcnOracle, glngSys)
    End If
    Call gobjESign.VerifySignature(strSource, lngSignID, 2)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Public Function GetSignSourceFromRTF1(ByVal lFileId As Long, ByVal lSignID As Double, ByVal lSignKey As Long, ByVal blnMoved As Boolean) As String
'����:�����ļ�ID,��ȡǩ��Դ����,�����ڲ��򿪱༭������½���ǩ����֤
'����:1 ȥ�����һ��ǩ�����ڵ�ǩ��ͼƬ (�����) ͼƬ�������Ϊ " PS(0000000X,0,1) PE(0000000X,0,1) "
'     2 ȥ������S�ؼ��ֵ�ǩ������
'     3 ������ͼƬ,���ؼ����м��"��"�ֻ��ɿո�(��Ϊǩ��ʱ�ǿո��ʾ)
'     4 ��ǩ��Ҫ�ػ�ԭ��ǩ��ʱ��"{����ҽʦǩ��}""����ҽʦǩ��""����ҽʦǩ��"
    Dim lSS As Long, lSE As Long, lES As Long, lEE As Long, bNeeded As Boolean, bFinded As Boolean, lKey As Long, lPos As Long
    Dim strZipFile As String, strRtfFile As String, lSigns As Long, lSEKey As Long
    Dim rsTemp As New ADODB.Recordset, strSource As String
    
    gstrSQL = "Select a.Id" & vbNewLine & _
                "From ���Ӳ������� A, ���Ӳ������� B" & vbNewLine & _
                "Where a.�ļ�id = [1] And a.�ļ�id = b.�ļ�id And b.Id = [2] And a.��ʼ�� > b.��ʼ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ��������¼", lFileId, lSignID)
    If Not rsTemp.EOF Then MsgBox "��ǰ����ǩ�������޸ģ�������֤�ϴ�ǩ��������޸ģ�", vbInformation, gstrSysName
    
    strZipFile = zlBlobRead(5, lFileId, , blnMoved)
    If gobjFSO.FileExists(strZipFile) Then
        strRtfFile = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strRtfFile) Then
                gfrmPublic.edtBuff.OpenDoc strRtfFile
                gobjFSO.DeleteFile strRtfFile, True
        End If
        gobjFSO.DeleteFile strZipFile, True
    End If
    
    gfrmPublic.edtBuff.Freeze
    gfrmPublic.edtBuff.ForceEdit = True
    '�������һ��ǩ�����ڵ�ǩ��ͼƬ,ǩ��ͼƬ����ǩ���ؼ���
    bFinded = FindKey(gfrmPublic.edtBuff, "S", lSignKey, lSS, lSE, lES, lEE, bNeeded)
    If bFinded Then
        If gfrmPublic.edtBuff.Range(lEE, lEE + 4).Text = " PS(" Then
            gfrmPublic.edtBuff.Range(lEE, lEE + 35).Text = ""
        End If
    End If
    
    'ȥ������S�ؼ��ֵ�ǩ������
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "S", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSS, lEE).Text = "": lSigns = lSigns + 1
    Loop Until bFinded = False
    
    '�����б��ؼ����м��"��"�ֻ��ɿո�(��Ϊǩ��ʱ�ǿո��ʾ,���ǩ��ʱ���״�Ϊ�ո�������Ϊ�ʺ�)
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "T", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSE, lES).Text = IIf(lSigns = 1, " ", "?"): lPos = lEE + 1
    Loop Until bFinded = False
    
    '������ͼƬ�ؼ����м��"��"�ֻ��ɿո�(��Ϊǩ��ʱ�ǿո��ʾ)
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "P", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSE, lES).Text = IIf(lSigns = 1, " ", "?"): lPos = lEE + 1
    Loop Until bFinded = False
    
    '������ǩ��Ҫ��,��Ϊǩ������ʱʹ�õ�Դ����ǩ��Ҫ����"{����ҽʦǩ��}"��ʽ����ǩ���󱻸���Ϊ���������
    gstrSQL = "Select �������� From ���Ӳ������� Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ��Ҫ��", lSignID)
    If Not rsTemp.EOF Then
    If UBound(Split(rsTemp!��������, ";")) > 5 Then '��ʷ�汾����û�е�6��
    lSEKey = Val(Split(rsTemp!��������, ";")(6))
    If lSEKey <> 0 Then 'ǩ������û��ʹ��ǩ��Ҫ��
        gstrSQL = "Select Ҫ������" & vbNewLine & _
                "From ���Ӳ�������" & vbNewLine & _
                "Where �ļ�id = [1] And �������� = 4 And ������ = [2] And Ҫ������ In ('����ҽʦǩ��', '����ҽʦǩ��', '����ҽʦǩ��')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ��Ҫ������", lFileId, lSEKey)
        If Not rsTemp.EOF Then '����ʹ��ǩ��Ҫ�����ƻ�ԭ
            lPos = 0
            bFinded = FindKey(gfrmPublic.edtBuff, "E", lSEKey, lSS, lSE, lES, lEE, bNeeded)
            If bFinded Then
                If gfrmPublic.edtBuff.Range(lEE, lEE + 4).Text = " PS(" Then 'ʹ����ǩ��ͼƬ
                    gfrmPublic.edtBuff.Range(lEE, lEE + 35).Text = ""
                End If
                gfrmPublic.edtBuff.Range(lSE, lES).Text = "{" & rsTemp!Ҫ������ & "}"
            End If
        End If
    End If
    End If
    End If
    
    gfrmPublic.edtBuff.ForceEdit = False
    gfrmPublic.edtBuff.UnFreeze
    strSource = gfrmPublic.edtBuff.Text
    GetSignSourceFromRTF1 = strSource
End Function
Public Function GetSignSourceFromRTF2(ByVal lFileId As Long, ByVal lSignID As Double, ByVal lSignKey As Long, ByVal blnMoved As Boolean) As String
'����:�����ļ�ID,��ȡǩ��Դ����,�����ڲ��򿪱༭������½���ǩ����֤
'����:1 ȥ�����һ��ǩ�����ڵ�ǩ��ͼƬ (�����) ͼƬ�������Ϊ " PS(0000000X,0,1) PE(0000000X,0,1) "
'     2 ȥ������S�ؼ��ֵ�ǩ������
'     3 ������ͼƬ,���ؼ����м��"��"�ֻ��ɿո�(��Ϊǩ��ʱ�ǿո��ʾ)
'     4 ��ǩ��Ҫ�ػ�ԭ��ǩ��ʱ��"{����ҽʦǩ��}""����ҽʦǩ��""����ҽʦǩ��"
    Dim lSS As Long, lSE As Long, lES As Long, lEE As Long, bNeeded As Boolean, bFinded As Boolean, lKey As Long, lPos As Long
    Dim strZipFile As String, strRtfFile As String, lSEKey As Long
    Dim rsTemp As New ADODB.Recordset, strSource As String
        
    gstrSQL = "Select a.Id" & vbNewLine & _
                "From ���Ӳ������� A, ���Ӳ������� B" & vbNewLine & _
                "Where a.�ļ�id = [1] And a.�ļ�id = b.�ļ�id And b.Id = [2] And a.��ʼ�� > b.��ʼ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ��������¼", lFileId, lSignID)
    If Not rsTemp.EOF Then MsgBox "��ǰ����ǩ�������޸ģ�������֤�ϴ�ǩ��������޸ģ�", vbInformation, gstrSysName
    
    strZipFile = zlBlobRead(5, lFileId, , blnMoved)
    If gobjFSO.FileExists(strZipFile) Then
        strRtfFile = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strRtfFile) Then
                gfrmPublic.edtBuff.OpenDoc strRtfFile
                gobjFSO.DeleteFile strRtfFile, True
        End If
        gobjFSO.DeleteFile strZipFile, True
    End If
    
    gfrmPublic.edtBuff.Freeze
    gfrmPublic.edtBuff.ForceEdit = True
    '�������һ��ǩ�����ڵ�ǩ��ͼƬ,ǩ��ͼƬ����ǩ���ؼ���
    bFinded = FindKey(gfrmPublic.edtBuff, "S", lSignKey, lSS, lSE, lES, lEE, bNeeded)
    If bFinded Then
        If gfrmPublic.edtBuff.Range(lEE, lEE + 4).Text = " PS(" Then
            gfrmPublic.edtBuff.Range(lEE, lEE + 35).Text = ""
        End If
    End If
    
    'ȥ������S�ؼ��ֵ�ǩ������
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "S", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSS, lEE).Text = ""
    Loop Until bFinded = False
    
    '�����б��ؼ����м��"��"�ֻ��ɿո�(��Ϊǩ��ʱ�ǿո��ʾ,���ǩ��ʱ���״�Ϊ�ո�������Ϊ�ʺ�)
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "T", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSE, lES).Text = " ": lPos = lEE + 1
    Loop Until bFinded = False
    
    '������ͼƬ�ؼ����м��"��"�ֻ��ɿո�(��Ϊǩ��ʱ�ǿո��ʾ)
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "P", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSE, lES).Text = " ": lPos = lEE + 1
    Loop Until bFinded = False
    
    '������ǩ��Ҫ��,��Ϊǩ������ʱʹ�õ�Դ����ǩ��Ҫ����"{����ҽʦǩ��}"��ʽ����ǩ���󱻸���Ϊ���������
    gstrSQL = "Select �������� From ���Ӳ������� Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ��Ҫ��", lSignID)
    If Not rsTemp.EOF Then
    If UBound(Split(rsTemp!��������, ";")) > 5 Then '��ʷ�汾����û�е�6��
    lSEKey = Val(Split(rsTemp!��������, ";")(6))
    If lSEKey <> 0 Then 'ǩ������û��ʹ��ǩ��Ҫ��
        gstrSQL = "Select Ҫ������" & vbNewLine & _
                "From ���Ӳ�������" & vbNewLine & _
                "Where �ļ�id = [1] And �������� = 4 And ������ = [2] And Ҫ������ In ('����ҽʦǩ��', '����ҽʦǩ��', '����ҽʦǩ��')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ��Ҫ������", lFileId, lSEKey)
        If Not rsTemp.EOF Then '����ʹ��ǩ��Ҫ�����ƻ�ԭ
            lPos = 0
            bFinded = FindKey(gfrmPublic.edtBuff, "E", lSEKey, lSS, lSE, lES, lEE, bNeeded)
            If bFinded Then
                If gfrmPublic.edtBuff.Range(lEE, lEE + 4).Text = " PS(" Then 'ʹ����ǩ��ͼƬ
                    gfrmPublic.edtBuff.Range(lEE, lEE + 35).Text = ""
                End If
                gfrmPublic.edtBuff.Range(lSE, lES).Text = "{" & rsTemp!Ҫ������ & "}"
            End If
        End If
    End If
    End If
    End If
    
    gfrmPublic.edtBuff.ForceEdit = False
    gfrmPublic.edtBuff.UnFreeze
    strSource = gfrmPublic.edtBuff.Text
    strSource = Replace(strSource, Chr(32), "")
    strSource = Replace(strSource, vbCr, "")
    strSource = Replace(strSource, vbLf, "")
    GetSignSourceFromRTF2 = strSource
End Function
Public Function GetSignSourceFromDB(ByVal lFileId As Long, ByVal lSignKey As String)
'���ܣ�ʹ�ñ������ݿ��������ı���������� ǩ��Ҫ�أ�ǩ������,ͼƬ������Ӷ���Ϊ����ǩ��ԭ��
'˵����ID <> ������� ��ʾ�Ǳ����Ӷ��󣬱����Ӷ�����ֹ��=��ǰǩ����Σ��Դ�Ϊ����,������������ֹ��=��ǰǩ�����-1
    Dim rsTemp As New ADODB.Recordset, strSource As String
    gstrSQL = "Select ID, ��id, ��ʼ��, ��ֹ��, ��������,��������, �����ı�, �������, �����д�, Ҫ������" & vbNewLine & _
                "From (Select ID, ��id, ��ʼ��, ��ֹ��, ��������,��������, �����ı�, �������, �����д�, Ҫ������" & vbNewLine & _
                "       From ���Ӳ������� A, (Select ��ʼ�� ��� From ���Ӳ������� Where �ļ�id = [1] And �������� = 8 And ������ = [2]) B" & vbNewLine & _
                "       Where �ļ�id = [1] And Instr(',����ҽʦǩ��,����ҽʦǩ��,����ҽʦǩ��,',',' || Ҫ������ || ',') = 0 And" & vbNewLine & _
                "             (��ʼ�� <= b.��� And ��ֹ�� = 0 Or ��ʼ�� <= b.��� And ��ֹ�� = b.��� And ID <> ������� Or" & vbNewLine & _
                "             ��ʼ�� <= b.��� And ��ֹ�� = b.��� + 1 And ID = �������))" & vbNewLine & _
                "Start With ��id Is Null" & vbNewLine & _
                "Connect By Prior ID = ��id" & vbNewLine & _
                "Order Siblings By Decode(ID, �������, 1, �������), �����д�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡԭ��", lFileId, lSignKey)
    Do Until rsTemp.EOF
        Select Case rsTemp!��������
            Case 1 '������٣���Ϊ��ٲ���ʾ����ٵ�������Ϊ�ı�����棬��SQL��Ҫ��������β�ѯ
            Case 5
                strSource = strSource & rsTemp!ID
            Case 8
                strSource = strSource & Split(rsTemp!�����ı�, ";")(0) & Split(rsTemp!��������, ";")(4)
            Case Else
                strSource = strSource & rsTemp!�����ı�
        End Select
        rsTemp.MoveNext
    Loop
    strSource = Replace(strSource, Chr(32), "")
    strSource = Replace(strSource, vbCr, "") 'ȥ�����еĻس����з�����Ϊ��һ��ʱ�����Ļ���ֻ�лس������ٴ��޸ı����ǩ����󣬱��޸ĳɻس�����
    strSource = Replace(strSource, vbLf, "")
    GetSignSourceFromDB = strSource
End Function

Public Function getPassESign(ByVal lngKind As Long, ByVal lngDeptId As Long) As Long
Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '0-����ҽ���Ͳ�����1-סԺҽ��ҽ���Ͳ�����2-סԺ��ʿҽ����3-ҽ��ҽ���ͱ��棻4-�����¼�ͻ�������5-ҩƷ��ҩ��6-LIS;7-PACS
    
    gstrSQL = "Select Zl_Fun_Getsignpar([1],[2]) as ���� From Dual "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ǩ�����Ʋ���", lngKind, lngDeptId)
    If rsTemp.EOF Then
        getPassESign = 1
    Else
        getPassESign = rsTemp!����
    End If

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function LongIDsTable(ByVal strIDs As String, ByRef idPar() As String, Optional ByVal idParStart As Long = 1, Optional ByVal Alias As String = "B") As String
Dim strSQL As String, lngS As String, N As Integer, strReturn As String, strThis As String
    
    ReDim idPar(10) As String
    strSQL = "Select Column_Value ID From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))"
    N = 0
    Do While True
        If Len(strIDs) <= 4000 Then
            strThis = strIDs
            strIDs = ""
        Else
            strThis = Mid(strIDs, 1, InStrRev(Mid(strIDs, 1, 4000), ",") - 1)
            strIDs = Mid(strIDs, InStrRev(Mid(strIDs, 1, 4000), ",") + 1)
        End If
        
        If N > 9 Then
            strReturn = strReturn & vbNewLine & " Union " & Replace(strSQL, "[1]", "'" & strThis & "'")
        Else
            idPar(N) = strThis
            strReturn = IIf(strReturn = "", "", strReturn & vbNewLine & " Union ") & Replace(strSQL, "[1]", "[" & (N + idParStart) & "]")
        End If
        
        N = N + 1
        If strIDs = "" Then Exit Do
    Loop
    
    LongIDsTable = " (" & strReturn & ") " & Alias & " "
End Function
Public Function GetEPRContentNextId() As Double
'���ܣ���ʹ�õ�����PACS������Ӳ����������������˷ѳ���LONG�����ֵ����ʱ����Ϊ  ������ȡ���Ӳ���������ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    '�����ô������,ԭ��������ʧЧ��û������ʱ,Ӧ�÷��ش���,��Ȼ������,��������!

    
    strSQL = "Select ���Ӳ�������_ID.Nextval From Dual"
    
    Call zl9ComLib.SQLTest(App.ProductName, "mRichEPR", strSQL)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������ȡ���Ӳ���������ID")
    Call zl9ComLib.SQLTest
    GetEPRContentNextId = rsTmp.Fields(0).Value
End Function
