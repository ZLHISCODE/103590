Attribute VB_Name = "mRichEPR"
Option Explicit
Public Const ELE_BACKCOLOR = &HFFEBD7               'Ҫ�صı�����ɫ '&HDCDCDC
Public Const ELE_UNDERLINE = cprWave                'Ҫ�ص��»���
Public Const PROTECT_BGCOLOR = &HE0E0E0             '�Զ��屣���ı��ı���ɫ
Public Const PROTECT_FORECOLOR = &H662200           '�Զ��屣���ı���ǰ��ɫ
Public gobjRegister As Object                       '������֤���
Public gobjESign As Object                          '����ǩ���ӿڲ���
Public Type PreDefinedKeyInfo   '�����ؼ���
    KeyStart As String
    KeyEnd As String
End Type
Public gKeyWords(1 To 6) As PreDefinedKeyInfo       'Ԥ����ؼ���
'##########################################################################################
'## ѹ�����ѹ
'##########################################################################################
Private mclsZip As New cTabZip
Private mclsUnzip As New cTabUnzip

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
    Dim aryChunk() As Byte, strText As String
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
    
    gstrSQL = "Select Zl_Lob_Read(" & Action & ",'" & KeyWord & "'," & "[1]" & IIf(blnMoved, ",1", "") & ") as Ƭ�� From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlBlobRead", lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.fields(0).Value) Then Exit Do
        strText = rsLob.fields(0).Value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
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
    Kill strFile: zlBlobRead = ""
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
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
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
        strText = Join(aryHex, "")
        gstrSQL = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
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
Public Function zlBlobSql(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String, arrSQL As Variant) As Boolean
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    Err = 0: On Error GoTo errHand
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    Err = 0: On Error GoTo errHand
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize '2000�ֽ�����
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0) '����=0��ʾ����
    
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
        strText = Join(aryHex, "")
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
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
Public Function zlFileUnzip(ByVal strZipFile As String, Optional ByVal strExtenName As String = "XML") As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then zlFileUnzip = "": Exit Function 'ԭ�ļ�������ֱ���˳�
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))       '��ȡԭ�ļ�·��
    
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)    'ԭ�ļ�Ŀ¼��������ʱĿ¼
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp & "\TMP." & strExtenName                          'ָ����ʱĿ¼�µĽ�ѹ�ļ�ȫ·��
    If gobjFSO.FileExists(strZipFileTmp) Then gobjFSO.DeleteFile strZipFileTmp      '���ȫ·���ļ�����ɾ��
    
    With mclsUnzip                                                                  '��ѹ
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FileExists(strZipFileTmp) Then                                       '��ѹ����ʱ�ļ�����
        
        strZipFileName = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer) & "." & strExtenName  '����ԭ�ļ�Ŀ¼����ʱ�ļ�
        If gobjFSO.FileExists(strZipFileName) Then gobjFSO.DeleteFile strZipFileName
                
        Call gobjFSO.CopyFile(strZipFileTmp, strZipFileName)                        '����ѹ�ļ�COPY��ԭ�ļ�Ŀ¼��
        
        If gobjFSO.FileExists(strZipFileTmp) Then gobjFSO.DeleteFile strZipFileTmp, True    'ɾ����ѹ�ļ�
        
        zlFileUnzip = strZipFileName
    Else
        zlFileUnzip = ""
    End If
    On Error Resume Next
    If gobjFSO.FolderExists(strZipPathTmp) Then gobjFSO.DeleteFolder strZipPathTmp, True 'ɾ����ѹ�ļ�Ŀ¼
    Err.Clear
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
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
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
    ByVal iPatientType As PatiFrom, _
    ByVal lngҽ��id As Long, Optional lngBabyNum As Long) As String
    
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4],[5],[6]) From Dual"
    Err = 0: On Error GoTo DBError
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�滻��", ElementName, CLng(sPatientID), _
        CLng(sPageID), CLng(iPatientType), lngҽ��id, lngBabyNum)
    If rsTmp.EOF Or rsTmp.BOF Then
        GetReplaceEleValue = ""
    Else
        GetReplaceEleValue = Trim(IIf(IsNull(rsTmp.fields(0).Value), "", rsTmp.fields(0).Value))
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
    Optional lngPatiID As Long, Optional lngPatiPageID As Long) As EPRSignLevel
    Dim rs As New ADODB.Recordset, lngR As Long, lngLevel1 As Long, lngLevel2 As Long
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select g.����" & vbNewLine & _
            "From zlRoleGrant g, Sys.Dba_Role_Privs r, �ϻ���Ա�� p" & vbNewLine & _
            "Where r.Grantee = Upper(p.�û���) And g.��ɫ = r.Granted_Role And g.ϵͳ = [2] And g.��� = [3] And g.���� = [4] And" & vbNewLine & _
            "      p.��Աid = [1]" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select [4] As ���� From �ϻ���Ա�� p Where �û��� = '" & UCase(gstrDbOwner) & "' And p.��Աid = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngUserID, glngSys, 1070, "ǩ��Ȩ")
    If rs.RecordCount <= 0 Then GetUserSignLevel = TabSL_�հ�: Exit Function
    
    gstrSQL = "select Ƹ�μ���ְ�� from ��Ա�� p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngUserID)
    If Not rs.EOF Then
        lngR = Nvl(rs("Ƹ�μ���ְ��"), 0)
    End If
    Select Case lngR    '1 ����  2 ����  3 �м�  4 ����/ʦ��  5 Ա/ʿ  9 ��Ƹ
    Case 1: lngLevel1 = TabSL_����
    Case 2: lngLevel1 = TabSL_����
    Case 3: lngLevel1 = TabSL_����
    Case Else: lngLevel1 = TabSL_����
    End Select
    rs.Close
    
    If lngPatiID > 0 Then
        gstrSQL = "Select ����ҽʦ, ����ҽʦ, ����ҽʦ " & _
            " From ���˱䶯��¼ " & _
            " Where ����ID = [1] And ��ҳID = [2] And (��ֹʱ�� Is Null Or ��ֹԭ�� = 1) " & _
            "       And ��ʼʱ�� Is Not Null And Nvl(���Ӵ�λ, 0) = 0"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngPatiID, lngPatiPageID)
        If rs.EOF Then
            lngLevel2 = TabSL_����
        Else
            If rs.fields("����ҽʦ") = IIf(strUserName = "", UserInfo.����, strUserName) Then
                lngLevel2 = TabSL_����
            ElseIf rs.fields("����ҽʦ") = IIf(strUserName = "", UserInfo.����, strUserName) Then
                lngLevel2 = TabSL_����
            Else
                lngLevel2 = TabSL_����
            End If
        End If
    End If
    GetUserSignLevel = IIf(lngLevel1 >= lngLevel2, lngLevel1, lngLevel2)
    Exit Function

errHand:
    GetUserSignLevel = TabSL_�հ�
End Function
Public Function GetCharColor(ByVal lng��ʼ�� As Long, ByVal lng��ֹ�� As Long) As OLE_COLOR
    '���ݿ�ʼ�桢��ֹ���ȡ�����ַ���ɫ
    Dim R As Long, G As Long, b As Long
    R = 255
    G = GetColorVectorG(lng��ʼ��)
    b = GetColorVectorB(lng��ֹ��)
    If G = 0 And b = 0 Then
        GetCharColor = vbBlack
    Else
        GetCharColor = RGB(R, G, b)
    End If
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
Public Function GetEPRContentNextId() As Double
'���ܣ���ʹ�õ�����PACS������Ӳ����������������˷ѳ���LONG�����ֵ����ʱ����Ϊ  ������ȡ���Ӳ���������ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    '�����ô������,ԭ��������ʧЧ��û������ʱ,Ӧ�÷��ش���,��Ȼ������,��������!

    
    strSQL = "Select ���Ӳ�������_ID.Nextval From Dual"
    
    Call zl9ComLib.SQLTest(App.ProductName, "mRichEPR", strSQL)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������ȡ���Ӳ���������ID")
    Call zl9ComLib.SQLTest
    GetEPRContentNextId = rsTmp.fields(0).Value
End Function