Attribute VB_Name = "mTend"

Option Explicit
Public mclsUnzip As New cUnzip
Public mclsZip As New cZip

Public glngHours As Long
Public gobjBodyEditor As Object
Public gobjPartogram As Object
Public gfrmPublic As Object
Public gobjFSO As New FileSystemObject
Public gobjESign As Object  '����ǩ���ӿڲ���

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
Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������
Public gstrMecState As String                '��ǰ���˲���״̬(���EprIsCommit����ʹ��)

'����
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const GWL_STYLE = (-16)              'Set the window style
Public Const WS_CAPTION = &HC00000
Public Const WS_THICKFRAME = &H40000        '��߿�
Public Const WS_SYSMENU = &H80000           '�ڱ������Ƿ�߱�ϵͳ�˵�
Public Const WS_MINIMIZEBOX = &H20000       '�߱���С����ť
Public Const WS_MAXIMIZEBOX = &H10000       '�߱���󻯰�ť
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'��ȡָ������ı߽���γߴ�
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'��ȡָ�����������
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'�ı䴰��λ�á�Zorder���ߴ��
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'�ı�ָ�����������
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
'�ı�ָ������ĸ�����
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
' ����ָ����Ϣ�����壬�ȴ�������ŷ��أ��� PostMessage() ����������Ϣ���������أ�HWND hWnd Ŀ�괰��ľ����wMsg�����͵���Ϣ��wParam��Ϣ��һ������lParam��Ϣ�ڶ�������
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

''ȥ��TextBox��Ĭ���Ҽ��˵�
'Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
'    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
'    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hwnd, msg, wp, lp)
'End Function

Public Function ReDimArray(ByRef strArray() As String) As Long
    '----------------------------------------------------------------------
    '���ܣ����¶�������
    '----------------------------------------------------------------------
    Dim lngCount As Long
    Dim strTmp As String
    
    On Error GoTo InitHand
    
    strTmp = strArray(1)
    
    lngCount = UBound(strArray) + 1
    
    GoTo OkHand
    
InitHand:
    
    lngCount = 1
    
OkHand:
    
    ReDim Preserve strArray(1 To lngCount)
            
    ReDimArray = lngCount
End Function

Public Function ZVal(ByVal varValue As Variant) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Public Function CreateBodyEditor() As Boolean
    Dim strDLL As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    If gobjBodyEditor Is Nothing Then
        gstrSQL = " Select �²��� From ���²��� Where Nvl(����,0)=1"
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡ���²���")
        If Err <> 0 Then
            strDLL = "zl9TemperatureChart"
        Else
            If rsTemp.RecordCount = 0 Then
                strDLL = "zl9TemperatureChart"
            Else
                strDLL = NVL(rsTemp!�²���, "zl9TemperatureChart")
            End If
        End If
        
        Err = 0
        strDLL = strDLL & ".clsBodyEditor"
        Set gobjBodyEditor = CreateObject(strDLL)
        If Err <> 0 Then
            MsgBox "    �������²���ʧ�ܣ�" & vbCrLf & "    ���򽫴�����׼�����²�����������չ�֣�����ָ�������²����Ƿ���ڻ����𻵣�" & vbCrLf & "    ��ϸ����" & Err.Description, vbInformation, gstrSysName
            
            '�������ָ�������²��������򴴽���׼�����²�������Ϊ���ﲻ����Ļ���������ܴ���ֱ��ʹ�����²����еĶ��󣬴Ӷ����³������
            strDLL = "zl9TemperatureChart.clsBodyEditor"
            Set gobjBodyEditor = CreateObject(strDLL)
        End If
        
        Call gobjBodyEditor.InitBodyEditor(glngSys, gcnOracle)
    End If
    
    CreateBodyEditor = True
    Exit Function
End Function

Public Function CreatePartogram() As Boolean
    Dim strDLL As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    If gobjPartogram Is Nothing Then
        gstrSQL = " Select ���� From ���̲��� Where Nvl(����,0)=1"
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡ���̲���")
        If Err <> 0 Then
            strDLL = "zl9Partogram"
        Else
            If rsTemp.RecordCount = 0 Then
                strDLL = "zl9Partogram"
            Else
                strDLL = NVL(rsTemp!����, "zl9Partogram")
            End If
        End If
        
        Err = 0
        strDLL = strDLL & ".clsPartogram"
        Set gobjPartogram = CreateObject(strDLL)
        If Err <> 0 Then
            MsgBox "    �������̲���ʧ�ܣ�" & vbCrLf & "    ���򽫴�����׼�Ĳ��̲�����������չ�֣�����ָ���Ĳ��̲����Ƿ���ڻ����𻵣�" & vbCrLf & "    ��ϸ����" & Err.Description, vbInformation, gstrSysName
            
            '�������ָ���Ĳ��̲��������򴴽���׼�Ĳ��̲�������Ϊ���ﲻ����Ļ���������ܴ���ֱ��ʹ�ò��̲����еĶ��󣬴Ӷ����³������
            strDLL = "zl9Partogram.clsPartogram"
            Err = 0
            Set gobjPartogram = CreateObject(strDLL)
        End If
        If Err <> 0 Then Err.Clear: Set gobjPartogram = Nothing: Exit Function
        
        Call gobjPartogram.InitPartogram(gcnOracle, glngSys)
    End If
    
    CreatePartogram = True
    Exit Function
End Function

Public Function ArchiveChart(ByVal lngFileID As Long) As Boolean
'���ܣ�����ļ��Ƿ�鵵
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    gstrSQL = "select 1 From ���˻����ļ� where ID=[1] And �鵵�� IS NOT NULL"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���µ��ļ��Ƿ�鵵", lngFileID)
    ArchiveChart = (rsTemp.RecordCount <> 0)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ToStandDate(ByVal strDate As String) As String
    Dim arrData
    Dim strMonth As String, strDay As String
    
    arrData = Split(strDate, "/")
    strMonth = arrData(1)
    strDay = arrData(0)
    If Len(strMonth) = 1 Then strMonth = "0" & strMonth
    If Len(strDay) = 1 Then strDay = "0" & strDay
    ToStandDate = strMonth & "-" & strDay
End Function

Public Sub GetUserInfo()
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo ErrHand
        
    Set rsTemp = zlDatabase.GetUserInfo
    With rsTemp
        If .RecordCount <> 0 Then
            gstrDBUser = .Fields("�û���").Value
            glngUserId = .Fields("ID").Value                '��ǰ�û�id
            gstrUserCode = .Fields("���").Value            '��ǰ�û�����
            gstrUserName = .Fields("����").Value            '��ǰ�û�����
            gstrUserAbbr = NVL(.Fields("����").Value, "")  '��ǰ�û�����
            glngDeptId = .Fields("����id").Value            '��ǰ�û�����id
            gstrDeptCode = .Fields("������").Value        '��ǰ�û�
            gstrDeptName = .Fields("������").Value        '��ǰ�û�
        Else
            gstrDBUser = ""
            glngUserId = 0
            gstrUserCode = ""
            gstrUserName = ""
            gstrUserAbbr = ""
            glngDeptId = 0
            gstrDeptCode = ""
            gstrDeptName = ""
        End If
    End With
    
    gstrSQL = "Select ǩ�� From ��Ա�� Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ������", glngUserId)
    If Not rsTemp.EOF Then
        gstrSignName = NVL(rsTemp!ǩ��, gstrUserName)
    End If
   
   
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0
End Sub

Public Function GetDbOwner(ByVal lngSys As Long) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    GetDbOwner = ""
    Err = 0: On Error GoTo ErrHand
    strSQL = "Select ������ From Zlsystems Where ��� = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetDbOwner", lngSys)
    If rsTemp.RecordCount <> 0 Then GetDbOwner = "" & rsTemp!������
    rsTemp.Close
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "SQL", adVarChar, 300
        .Fields.Append "Trans", adTinyInt                   '1��ʾ��ʼ;2��ʾ����
        .Fields.Append "Custom", adTinyInt
        .Fields.Append "Parameter", adVarChar, 500
        
        .Open
    End With
    
    SQLRecord = True
    
    Exit Function
    
ErrHand:
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    rs.AddNew
    rs("SQL").Value = strSQL
    rs("Trans").Value = intTrans
    rs("Custom").Value = intCustom
    rs("Parameter").Value = strParameter
    SQLRecordAdd = True
    
    Exit Function
    
ErrHand:
End Function

Public Function SQLRecordExecute(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    If rs.RecordCount > 0 Then
        If Len(strTitle) = 0 Then strTitle = gstrSysName
        blnTran = True
        
        If blnHaveTrans Then gcnOracle.BeginTrans
        
        rs.MoveFirst
    
        For intLoop = 1 To rs.RecordCount
            
            If Val(rs("Custom").Value) = 0 Then
                strSQL = CStr(rs("SQL").Value)
                Call zlDatabase.ExecuteProcedure(strSQL, strTitle)
            End If
            
            rs.MoveNext
        Next
    
        If blnHaveTrans Then gcnOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordExecute = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
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
    
    On Error GoTo ErrHand
    
    If Not gobjFSO.FileExists(strZipFile) Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = gobjFSO.GetSpecialFolder(2)
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
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
    
ErrHand:
    Call SaveErrLog
End Function

Public Function GetTmpPath() As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    GetTmpPath = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
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
'## ���ܣ�  �����ݴ�һ��XtremeReportControl�ؼ����Ƶ�VSFlexGrid���Ա���д�ӡ
'################################################################################################################
Public Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '��ȫ����ǿ��չ��,�������ݱ��
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    Dim lngCol As Long, lngRow As Long
    
    On Error GoTo ErrHand:
    For Each rptRow In rptList.Rows
        If rptRow.GroupRow Then rptRow.Expanded = True
    Next
    
    With vfgList
        .Clear
        .Rows = rptList.Records.Count + 1
        .Cols = 0: .Cols = rptList.Columns.Count
        .FixedCols = rptList.GroupsOrder.Count
        
        '�����и���
        .ROW = 0
        lngCol = 0
        For Each rptCol In rptList.GroupsOrder
            .TextMatrix(0, lngCol) = rptCol.Caption
            .ColData(lngCol) = rptCol.ItemIndex
            Select Case rptCol.Alignment
            Case xtpAlignmentLeft: .FixedAlignment(lngCol) = flexAlignLeftCenter
            Case xtpAlignmentCenter: .FixedAlignment(lngCol) = flexAlignCenterCenter
            Case xtpAlignmentRight:  .FixedAlignment(lngCol) = flexAlignRightCenter
            End Select
            .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .FixedAlignment(lngCol)
            .ColWidth(lngCol) = rptCol.Width * 15
            .MergeCol(lngCol) = True
            lngCol = lngCol + 1
        Next
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .TextMatrix(0, lngCol) = rptCol.Caption
                .ColData(lngCol) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .ColAlignment(lngCol) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .ColAlignment(lngCol) = flexAlignCenterCenter
                Case xtpAlignmentRight: .ColAlignment(lngCol) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
                .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .ColAlignment(lngCol)
                If rptCol.Width < 20 Then
                    .ColWidth(lngCol) = 0
                Else
                    .ColWidth(lngCol) = rptCol.Width * 15
                End If
                lngCol = lngCol + 1
            End If
        Next
        vfgList.Cols = lngCol
        
        '�����и���
        lngRow = 0
        For Each rptRow In rptList.Rows
            If rptRow.GroupRow = False Then
                lngRow = lngRow + 1
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(lngRow, lngCol) = rptRow.Record(.ColData(lngCol)).Value
                Next
            End If
        Next
    End With
    zlReportToVSFlexGrid = True
    Exit Function

ErrHand:
    zlReportToVSFlexGrid = False
End Function

'################################################################################################################
'## ���ܣ�  �����ݴ�һ��չʾVSFlexGrid�ؼ����Ƶ���ӡVSFlexGrid���Ա���д�ӡ
'################################################################################################################
Public Function zlDataToPrint(vsfPrint As VSFlexGrid, VsfData As VSFlexGrid) As Boolean
    '-------------------------------------------------
    '��ȫ����ǿ��չ��,�������ݱ��
    
    Dim lngCol As Long, lngRow As Long
    Dim lngPrintCol As Long, lngPrintRow As Long
    On Error GoTo ErrHand:
    
    With vsfPrint
        .Clear
        .MergeCells = flexMergeFixedOnly ' = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .Rows = VsfData.Rows
        .Cols = 0: .Cols = VsfData.Cols
        .FixedCols = VsfData.FixedCols
        .FixedRows = VsfData.FixedRows
        .GridColor = vbBlack
        
        '�����и���
        .ROW = 0
        .Rows = VsfData.Rows
        .Cols = VsfData.Cols
        lngPrintRow = 0
        For lngRow = 0 To VsfData.Rows - 1
            If Not VsfData.RowHidden(lngRow) Then
                lngPrintCol = 0
                For lngCol = 0 To VsfData.Cols - 1
                    If Not VsfData.ColHidden(lngCol) Then
                        .TextMatrix(lngPrintRow, lngPrintCol) = VsfData.TextMatrix(lngRow, lngCol)
                        .ColWidth(lngPrintCol) = VsfData.ColWidth(lngCol)
                         .ColAlignment(lngPrintCol) = VsfData.ColAlignment(lngCol)
                        lngPrintCol = lngPrintCol + 1
                    End If
                Next
                .RowHeight(lngPrintRow) = VsfData.RowHeight(lngRow)
                lngPrintRow = lngPrintRow + 1
            End If
        Next
        
        lngPrintCol = 0
        For lngCol = 0 To .Cols - 1
           If VsfData.ColHidden(lngCol) Then lngPrintCol = lngPrintCol + 1
        Next
        .Cols = VsfData.Cols - lngPrintCol
         lngPrintRow = 0
        For lngRow = 0 To .Rows - 1
           If VsfData.RowHidden(lngRow) Then lngPrintRow = lngPrintRow + 1
        Next
        .Rows = VsfData.Rows - lngPrintRow
        
        lngPrintRow = 0
        For lngRow = 0 To .FixedRows - 1
           If VsfData.RowHidden(lngRow) Then lngPrintRow = lngPrintRow + 1
        Next
        .FixedRows = .FixedRows - lngPrintRow
        If .FixedRows = 0 Then .FixedRows = 1
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
       
        '�ٰ��кϲ�
        For lngRow = 0 To .FixedRows - 1
            .MergeRow(lngRow) = True
        Next
        
        
    End With
    zlDataToPrint = True
    Exit Function

ErrHand:
    zlDataToPrint = False
End Function


Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'���ܣ���ʾ������һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.hWnd, vRect)
    lngStyle = GetWindowLong(objForm.hWnd, GWL_STYLE)
    If blnCaption Then
        lngStyle = lngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then lngStyle = lngStyle Or WS_SYSMENU
        If objForm.MaxButton Then lngStyle = lngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then lngStyle = lngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong objForm.hWnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.hWnd, 0, 0, 0, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Public Function IsAllowInput(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strTime As String, ByVal strCurTime As String) As Boolean
    'ȡ��ָ��������ָ��ʱ��֮��ؼ����ʱ��
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    IsAllowInput = True
    gstrSQL = "" & _
              " SELECT DECODE(��ֹԭ��,1,'��Ժ',3,'ת��',10,'Ԥ��Ժ',15,'ת����',DECODE(��ʼԭ��,10,'��Ժ','δ����')) AS ����,��ֹʱ�� AS ʱ��" & _
              " From ���˱䶯��¼" & _
              " WHERE (��ֹԭ�� IN (1,3,10,15) OR ��ʼԭ��=10) And ����ID=[1] And ��ҳID=[2] And [3] <= ��ֹʱ��" & _
              " ORDER BY ��ֹʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ָ��������ָ��ʱ��֮��ؼ����ʱ��", lng����ID, lng��ҳID, CDate(strTime))
    If rsTemp.RecordCount = 0 Then Exit Function
    
    'ֻȡ��һ�����ϵļ�¼
    strTime = Format(DateAdd("H", glngHours, rsTemp!ʱ��), "yyyy-MM-dd HH:mm")
    strCurTime = Format(strCurTime, "yyyy-MM-dd HH:mm")
    
    If strTime < strCurTime Then IsAllowInput = False
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SQLDIY(strSQL As String)
    If gblnMoved Then
        strSQL = Replace(strSQL, "���˻����ļ�", "H���˻����ļ�")
        strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
        strSQL = Replace(strSQL, "���˻�����ϸ", "H���˻�����ϸ")
        strSQL = Replace(strSQL, "���˻����ӡ", "H���˻����ӡ")
    End If
End Sub

Public Function GetAdviceOutTime(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer) As String
'����:��ȡ���˻�Ӥ����ҽ����Ժʱ��
    Dim rsTemp As New ADODB.Recordset
    Dim strTmp As String, strTime As String
    On Error GoTo ErrHand
    If intӤ�� = 0 Then
        strTmp = ",5,11,"
    Else
        strTmp = ",3,5,11,"
    End If
    gstrSQL = "Select ��ʼִ��ʱ��" & vbNewLine & _
        " From ����ҽ����¼ b, ������ĿĿ¼ c" & vbNewLine & _
        " Where b.������Ŀid + 0 = c.Id And b.ҽ��״̬ = 8 And Nvl(b.Ӥ��, 0) <> 0 And c.��� = 'Z' And instr([4],',' || c.�������� || ',',1)>0 And" & vbNewLine & _
        "      b.����id = [1] And b.��ҳid = [2] And b.Ӥ�� = [3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҽ����Ժʱ��", lng����ID, lng��ҳID, intӤ��, strTmp)
    If rsTemp.RecordCount > 0 Then strTime = Format(rsTemp!��ʼִ��ʱ��, "YYYY-MM-DD HH:mm:ss")
    GetAdviceOutTime = strTime
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function EprIsCommit(ByVal lngPatiID As Long, ByVal lngPageId As Long) As String
'��|�ָ���ʽ����,״̬Ϊ0 ������ 1 �����ֱ���� ����|ɾ��|����

    Dim rsTemp As ADODB.Recordset
    Dim intNew As Integer, intDel As Integer, intMod As Integer
    Dim strState As String
    
    EprIsCommit = "1|1|1"
    strState = "δ���": gstrMecState = strState
    On Error GoTo ErrHand
    gstrSQL = "Select ����״̬ From ������ҳ Where ����id = [1] And ��ҳid = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����״̬", lngPatiID, lngPageId)
    If Not rsTemp.EOF Then
        Select Case NVL(rsTemp!����״̬, 0)
            Case 0
                intNew = 1: intDel = 1: intMod = 1
                strState = "δ���"
            Case 1 '�ȴ����
                intNew = 0: intDel = 0: intMod = 0
                strState = "�ȴ����"
            Case 2 '�ܾ����
                intNew = 1: intDel = 1: intMod = 1
                strState = "�ܾ����"
            Case 3 '�������
                intNew = 0: intDel = 0: intMod = 0
                strState = "�������"
            Case 4 '��鷴��
                intNew = 0: intDel = 0: intMod = 1
                strState = "��鷴��"
            Case 5 '���鵵
                intNew = 0: intDel = 0: intMod = 0
                strState = "���鵵"
            Case 6 '�������
                intNew = 0: intDel = 0: intMod = 1
                strState = "�������"
            Case 13 '���ڳ��
                intNew = 1: intDel = 1: intMod = 1
                strState = "���ڳ��"
            Case 14 '��鷴��
                intNew = 1: intDel = 1: intMod = 1
                strState = "��鷴��"
            Case 16 '�������
                intNew = 1: intDel = 1: intMod = 1
                strState = "�������"
            Case 10 '���մ���
                intNew = 0: intDel = 0: intMod = 0
                strState = "���մ���"
            Case Else
                intNew = 0: intDel = 0: intMod = 0
                strState = "���"
        End Select
    End If
    gstrMecState = strState
    EprIsCommit = CStr(intNew) & "|" & CStr(intDel) & "|" & CStr(intMod)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ISCollectSigned(ByVal lngFileID As Long, ByVal strDate As String, ByVal strTime As String) As Boolean
    Dim blnDetail As Boolean
    Dim str����ʱ�� As String, strStart As String, strEnd As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    str����ʱ�� = Format(strDate, "yyyy-MM-dd")
    
    strStart = str����ʱ�� & " 00:00:00"
    strEnd = Format(DateAdd("d", 2, strStart), "yyyy-MM-dd HH:mm:ss")
    str����ʱ�� = str����ʱ�� & " " & strTime & ":00"
    
    gstrSQL = " Select A.����ʱ��,A.��ʼʱ��,A.����ʱ��,A.�������,B.���,B.��ʼ,B.����,A.ǩ����" & vbNewLine & _
              " From ���˻������� A,�������ʱ�� B" & vbNewLine & _
              " Where B.����(+)=2 And abs(A.�������)=B.���(+) And A.�������<0 And A.ǩ���� Is Not NULL And A.�ļ�ID=[1] And A.����ʱ�� between [2] and [3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ��¼����ʱ�䵱�켰֮��һ��Ļ�������", lngFileID, CDate(strStart), CDate(strEnd))
    With rsTemp
        'ѭ����飬�����˳�
        Do While Not .EOF
            'ƴ��ʼ����ʱ�䴮
            If IsNull(!���) Then
                strEnd = Format(!����ʱ��, "YYYY-MM-DD") & " " & !����ʱ�� & ":59"
                If !����ʱ�� < !��ʼʱ�� Then
                    strStart = Format(DateAdd("d", -1, !����ʱ��), "YYYY-MM-DD") & " " & !��ʼʱ�� & ":00"
                Else
                    strStart = Format(!����ʱ��, "YYYY-MM-DD") & " " & !��ʼʱ�� & ":00"
                End If
            Else
                strEnd = Format(!����ʱ��, "YYYY-MM-DD") & " " & !���� & ":59"
                If !���� < !��ʼ Then
                    strStart = Format(DateAdd("d", -1, !����ʱ��), "YYYY-MM-DD") & " " & !��ʼ & ":00"
                Else
                    strStart = Format(!����ʱ��, "YYYY-MM-DD") & " " & !��ʼ & ":00"
                End If
            End If
            
            If str����ʱ�� >= strStart And str����ʱ�� <= strEnd Then
                ISCollectSigned = True
                Exit Function
            End If
            .MoveNext
        Loop
    End With
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

