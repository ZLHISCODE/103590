Attribute VB_Name = "mdlPubFunc"
Option Explicit
'������������ö��
'====���������=============
'clipClear�����ճ����
'clipCopyFiles:�ļ�����ճ����
'====���̲���==============
'fun_KillProcess:ɱ��ָ�����ƽ���
'====�Զ���洢����==========
'GetBlankProcedure:��ȡ�洢���̲���Ĭ��ֵ
'IsSpaceProcedure:�洢�������Ƿ�ռ��
'====��̬��¼������==========
'CopyNewRec:ֱ�����ɻ��ߴ�ԭʼ��¼������һ�����صľ�̬��¼���������޸ļ�¼�������ݣ�
'RecDelete:ɾ������ָ�������ľ�̬��¼����������
'RecUpdate:��������ָ�������ľ�̬��¼����ĳЩ�ֶ�
'RecDataAppend:��һ����̬��¼�������ݸ��ӵ���һ����̬��¼����
'====������������===========
'ActualLen: ��ȡ�ַ������ֽڳ���
'ActualStr:�ַ�����ȡָ���ֽڳ���
'CancelNetServer:�Ͽ�����������
'Decode:ģ��Oracle��Decode����
'IsNetServer:�����������Ƿ�����
'OpenFolder:ѡ���ļ���
'SetCtrlPosOnLine:����һ��ؼ��Ķ��뷽ʽ�Լ��ؼ����
'CboSetWidth:����cbo�ؼ������б���
'GetControlRect��ȡ�ؼ�����Ļ�е�λ��
'CboSetIndex��Ϊһ��Combo�ؼ�ѡ���б�����ֲ�������Click�¼�
'GetClientPoint����ȡ��ǰָ���Ӧ�ڿؼ��е�λ��

'ѹ����ѹ����
Public Const PROAPPCTION = "7z.exe" 'ִ�г���
Public Const COMPRESSIONRATE = 5 '��׼ѹ��
'''ѹ���ȼ� ѹ���㷨 �ֵ��С �����ֽ� ƥ���� ������ ����
'''0 Copy ��ѹ��
'''1 LZMA 64 KB 32 HC4 BCJ ���ѹ��
'''3 LZMA 1 MB 32 HC4 BCJ ����ѹ��
'''5 LZMA 16 MB 32 BT4 BCJ ����ѹ��
'''7 LZMA 32 MB 64 BT4 BCJ ���ѹ��
'''9 LZMA 64 MB 64 BT4 BCJ2 ����ѹ��

'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'��ȡĳ�����뷨������
Private Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'�ж�ĳ�����뷨�Ƿ��������뷨
Private Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'�л���ָ�������뷨��
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal Flags As Long) As Long
Private Const NORMAL_PRIORITY_CLASS             As Long = &H20&
Private Const STARTF_USESTDHANDLES              As Long = &H100&
Private Const STARTF_USESHOWWINDOW              As Long = &H1
Private Const SW_HIDE                           As Integer = 0 '���ش��ڣ�������һ������
Public Const INFINITE                           As Long = &HFFFF&

Private mrsProgFuncs As ADODB.Recordset   '���ڻ���һ��ģ����ӵ�е���Ȩ����

'1-�䶯����;2-�հ׹���;3-�û�����
Public Enum ProcType
    �䶯���� = 1
    �հ׹��� = 2
    �û����� = 3
End Enum

Public Enum ProcState
    ����� = 0
    ������ = 1
    ������ = 2
    �ѵ��� = 3
    �ޱ仯 = 4
End Enum

'1-�ϴ��Զ�����;2-�ϴα�׼����;3-�����Զ�����;4-���α�׼����
Public Enum ProcTextType
    �ϴ��Զ����� = 1
    �ϴα�׼���� = 2
    �����Զ����� = 3
    ���α�׼���� = 4
End Enum
'�Զ���洢���̹���
Public Enum Color
    ��ɫ = &H80000005
    ��ɫ = &HFF&
    ��ɫ = &HFF0000
    ��ɫ = 0
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    ǳ��ɫ = &HE0E4E7
    ���ɫ = &H8000000C
    ��ɫ = &H8000000F
    ǳ��ɫ = &H80000018
    ����ģ��ɫ = &HC00000
    Ĭ��ǰ��ɫ = &H80000008
    ��ɫ = &HF5F5F5
    ����ɫ = 0
    ͣ��ɫ = 255
End Enum

Public Type AbortInfo
    AbortSys As Long
    AbortFile As String
    AbortLine As Long
    AbortInfo As String
    IsHistory As Boolean
End Type

Private mlngPid As Long '���̲���
Public gHwnd As Long '���̲���
Public gstrSplite As String

Public Enum LogType
    LT_��װ = 0
    LT_������Ǩ = 1
    LT_��ǰ��Ǩ = 2
    LT_��ʷ����Ǩ = 3
    LT_ϵͳ���� = 4 '��Ǩ����ϵͳ���ƣ��Լ���������
    LT_�Զ��� = -1 '�Զ����ļ����Լ��ļ���
    LT_������־ = 999
End Enum
Public gobjRIS As Object '�����ӿڶ���
Public gstrSTOwner As String '��׼��100������
Public gblnRIS As Boolean '���Դ���RIS�ӿ�
Public gblnMustRIS As Boolean

'OpenFolder��ʼ·������
Public gstrAPIPath As String
Public Declare Function GetTickCount Lib "kernel32" () As Long  '��ȡ��ǰʱ��

Public Function clipClear() As Boolean
'��յ�ǰ������
    Call EmptyClipboard
End Function

Public Function clipCopyFiles(File() As String) As Boolean
'       ģ�飺���������
'       ���ܣ����������,�����ļ�Ŀ¼��������
'       ��д��ף��
'       ���ڣ�2011��1��3��
'���ƶ���ļ���������
   On Error Resume Next
   Dim strData As String
   Dim df As DROPFILES
   Dim hGlobal As Long
   Dim lpGlobal As Long
   Dim i As Long
   strData = ""

   
   '������������ִ������
   If OpenClipboard(0&) Then
        '��յ�ǰ������
        Call EmptyClipboard
        
        '�ж��ļ������Ƿ�Ϊ��
        If SafeArrayGetDim(File) = 0 Then Exit Function
        For i = LBound(File) To UBound(File)
            strData = strData & File(i) & vbNullChar
        Next
        
        hGlobal = GlobalAlloc(GHND, Len(df) + LenB(strData))
        
        If hGlobal Then
            lpGlobal = GlobalLock(hGlobal)
         
            df.pFiles = Len(df)
            Call CopyMemory(ByVal lpGlobal, df, Len(df))
            Call CopyMemory(ByVal (lpGlobal + Len(df)), ByVal strData, LenB(strData))
   
            Call GlobalUnlock(hGlobal)
         
            If SetClipboardData(CF_HDROP, hGlobal) Then
                clipCopyFiles = True
            End If

        End If
        
        Call CloseClipboard
    End If
End Function

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
'=���ܣ� ͨ��PIDö�������ľ��,������Ҫ�Ĵ���
    Dim Pid1 As Long
    Dim wText As String * 255
    GetWindowThreadProcessId hwnd, Pid1
    If mlngPid = Pid1 Then
        GetWindowText hwnd, wText, 100
        If InStrRev(wText, "%", -1) > 0 Then
            gHwnd = hwnd
        End If
    End If
    EnumWindowsProc = True
End Function

Private Sub Find_Window(ByVal lngPid As Long)
'       ģ�飺���̾������
'       ���ܣ����̾���������ָ�����̵�Hwnd
'       ��д��ף��
'       ���ڣ�2010��11��24��
    mlngPid = lngPid
    gHwnd = 0
    EnumWindows AddressOf EnumWindowsProc, 0
End Sub

'���ҽ��̵ĺ���
Public Sub fun_KillProcess(ByVal ProcessName As String)
    Dim strData As String
    Dim my As PROCESSENTRY32
    Dim l As Long
    Dim l1 As Long
    Dim mName As String
    Dim i As Integer, Pid As Long
    Dim mProcID As Long
    l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If l Then
        my.dwSize = 1060
        If (Process32First(l, my)) Then
            Do
                i = InStr(1, my.szExeFile, Chr(0))
                mName = LCase(Left(my.szExeFile, i - 1))
                If mName = LCase(ProcessName) Then
                    Pid = my.th32ProcessID
                    mProcID = OpenProcess(1&, -1&, Pid)

                    TerminateProcess mProcID, 0&
                End If
            Loop Until (Process32Next(l, my) < 1)
        End If
        l1 = CloseHandle(l)
    End If
End Sub


Public Function IsSpaceProcedure(ByVal strOwner As String, ByVal strProcName As String) As Boolean
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "Select 1 From zlProcedure Where ����=[1] And ����=2"
    Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSql, "", UCase(strProcName))
    IsSpaceProcedure = (rsData.BOF = False)
    
End Function

Public Function GetBlankProcedure(ByVal strProc As String) As String
    Dim lngCount As Long
'    Dim blnTitleFlag As Boolean
'    Dim strEnd As String
    Dim strSql As String
'    Dim lngInstr As Long
    Dim strArr() As String
    
    Dim strLine As String
    Dim lngPostion As Long
    
    Dim strReturnType As String
    
    strArr = Split(strProc, vbCrLf)
    strSql = ""
    strReturnType = ""
    
    For lngCount = 0 To UBound(strArr)
        
        strLine = Replace(Trim(strArr(lngCount)), Chr(10), "")
        strLine = UCase(Replace(strLine, Chr(13), ""))

        'ȡ��--ע��
        lngPostion = InStr(strLine, "--")
        If lngPostion > 0 Then strLine = Mid(strLine, 1, lngPostion - 1)
        
        lngPostion = InStr(strLine, "RETURN ")
        If lngPostion > 0 Then
            
            If InStr(strLine, " NUMBER") > 0 Then
                strReturnType = "NUMBER"
            ElseIf InStr(strLine, " VARCHAR") > 0 Then
                strReturnType = "VARCHAR"
            ElseIf InStr(strLine, " DATE") > 0 Then
                strReturnType = "DATE"
            End If
        End If
        
        Select Case strLine
        Case "AS", "IS"
            strSql = strSql & strArr(lngCount) & vbCrLf
            Exit For
        Case Else
            If Right(strLine, 3) = " AS" Then
                strSql = strSql & strArr(lngCount) & vbCrLf
                Exit For
            ElseIf Right(strLine, 3) = " IS" Then
                strSql = strSql & strArr(lngCount) & vbCrLf
                Exit For
            Else
                strSql = strSql & strArr(lngCount) & vbCrLf
            End If
        
        End Select
    Next
    strSql = strSql & "Begin" & vbCrLf
    strSql = strSql & " " & vbCrLf
    Select Case strReturnType
    Case "NUMBER"
        strSql = strSql & vbTab & "return 0;" & vbCrLf
    Case "VARCHAR"
        strSql = strSql & vbTab & "return '';" & vbCrLf
    Case "DATE"
        strSql = strSql & vbTab & "return sysdate;" & vbCrLf
    End Select
    strSql = strSql & "End;"
    GetBlankProcedure = strSql
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
'���ܣ�ȡָ���ַ������ֽ���ĳ���
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function ActualStr(ByVal strAsk As String, ByVal lngLen As Long) As String
'���ܣ�ȡָ���ַ������ָ���ֽڳ��ȵ�����
    Dim strTemp As String, i As Long
    
    strTemp = StrConv(LeftB(StrConv(strAsk, vbFromUnicode), lngLen), vbUnicode)
    If InStr(strTemp, Chr(0)) > 0 Then
        strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
    End If
    ActualStr = strTemp
End Function

Public Function HScrollVisible(vsInput As Object) As Boolean
'�ж�ˮƽ�������Ŀɼ���
    Dim i As Long, lpMinPos As Long, lpMaxPos As Long
    
    HScrollVisible = False
    i = GetScrollRange(vsInput.hwnd, SB_HORZ, lpMinPos, lpMaxPos)
    If lpMaxPos <> lpMinPos And Not (lpMaxPos = 100 And lpMinPos = 0) Then HScrollVisible = True
End Function

Public Function VScrollVisible(vsInput As Object) As Boolean
'�жϴ�ֱ�������Ŀɼ���
    Dim i As Long, lpMinPos As Long, lpMaxPos As Long
    
    VScrollVisible = False
    i = GetScrollRange(vsInput.hwnd, SB_VERT, lpMinPos, lpMaxPos)
    If lpMaxPos <> lpMinPos And Not (lpMaxPos = 100 And lpMinPos = 0) Then VScrollVisible = True
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

Public Function OpenFolder(ByVal frmodtvOwner As Form, Optional strTitle As String, Optional ByVal strInitDir As String) As String
'    '----------------------------------------------------------------------------------------------------
'    '����:ѡ���ļ���
'    '����:frmodtvOwner-ѡ���ļ��еĸ�����
'    '       strFolderName-ָ�����ļ���
'    '       strTitle-����
'    '       strInitDir-Ĭ�ϴ�·��
'    '����:strFolderName-����ѡ����ļ���
'    '----------------------------------------------------------------------------------------------------
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    
    gstrAPIPath = strInitDir & Chr(0)
    With tBrowseInfo
        .hwndOwner = frmodtvOwner.hwnd
        .lpszTitle = lstrcat(strTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_STATUSTEXT
        .lpfnCallback = AddressOfFunction(AddressOf OpenDirCallbackProc)
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
       sBuffer = Space(MAX_PATH * 2)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
       OpenFolder = sBuffer
    End If
End Function

Public Function OpenDirCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
 '���ܣ�OpenFolder�ص��������������ô򿪵��ļ��ĳ�ʼ·��
    Dim lpIDList As Long
    Dim ret As Long
    Dim sBuffer As String
  
    On Error Resume Next
    
    Select Case uMsg
        Case BFFM_INITIALIZED
            Call SendMessage(hwnd, BFFM_SETSELECTION, 1, ByVal gstrAPIPath)
        Case BFFM_SELCHANGED
            sBuffer = Space(MAX_PATH * 2)
            ret = SHGetPathFromIDList(lp, sBuffer)
            If ret = 1 Then
                Call SendMessage(hwnd, BFFM_SETSTATUSTEXT, 0, ByVal sBuffer)
            End If
    End Select
    
    OpenDirCallbackProc = 0
End Function

Private Function AddressOfFunction(Address As Long) As Long
'���ܣ�OpenFolder�Ӻ���
    AddressOfFunction = Address
End Function

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean, Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
'������:����
'�޸��ˣ���˶
'�޸����ڣ�2014-1-6
'�޸ĵ㣺���Ӹ��Ƽ�¼���Ĳ����ֶι���
'��������:2000-11-02
'���Ƽ�¼��
'������strFields=��Ҫ���Ƶļ�¼�����ֶε���˳����ֶ�����ɵ��ַ���
'          �磺1 ����1,3 ����2,7 ����3...��ʾ���Ƽ�¼���ĵ�1,3,7..�ֶ���ɼ�¼��������
'              ID ����1,���� ����2,....��ʾ���Ƽ�¼����ID,����...�ֶ���ɼ�¼������
'              ����*Ϊ�µļ�¼��������
'              �������ͻ�����׳���������ͬ�����⣬��ע��
'               *,�ڱ�ʾ����ԭ��¼���������ֶΣ�������Ҫ��ԭ�����ֶ����²�������
'           arrAppFields=׷�ӵ��ֶ���Ϣ������,����,����,Ĭ��ֵ,û��Ĭ��ֵ��Empty,û��ָ�����ȴ�Empty
'      blnOnlyStructure=�Ƿ�ֻ���ƽṹ
'�ڳ����У��������漰���໥���ݼ�¼������ʹ��ADO��Clone���Ʋ����ļ�¼����������һ����¼�������ݷ����仯��ʱ�����и�������������ͬ�ı仯��ͨ��ָ�޸Ļ�ɾ����������������ϣ����Щ��¼���໥�䱣�ֶ���
  
    Dim rsClone As ADODB.Recordset
    Dim rsTarget As ADODB.Recordset
    Dim intFields As Integer, blnALlFileds As Boolean
    Dim arrFieldsName As Variant, strFieldName As String, strFieldNameAlias As String
    Dim arrTmp As Variant, arrFieldsTmp As Variant
    Dim i As Long
    
    If Not rsSource Is Nothing Then
        Set rsClone = rsSource.Clone
        rsClone.Filter = rsSource.Filter
    End If
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        '������¼���ṹ
        If strFields = "" Then
            strFields = "*"
        End If
        arrFieldsTmp = Split(strFields, ",")
        arrFieldsName = Array()
        For intFields = LBound(arrFieldsTmp) To UBound(arrFieldsTmp)
            If Trim(arrFieldsTmp(intFields)) = "*" Then '��ʶ�˴�������ԭ��¼����������
                If Not rsClone Is Nothing Then
                    For i = 0 To rsClone.Fields.Count - 1
                        ReDim Preserve arrFieldsName(UBound(arrFieldsName) + 1)
                        arrFieldsName(UBound(arrFieldsName)) = rsClone.Fields(i).name & ""
                        .Fields.Append rsClone.Fields(i).name, IIf(rsClone.Fields(i).Type = adNumeric, adDouble, rsClone.Fields(i).Type), rsClone.Fields(i).DefinedSize, adFldIsNullable    '0:��ʾ����
                    Next
                End If
            Else
                ReDim Preserve arrFieldsName(UBound(arrFieldsName) + 1)
                '�а�������
                arrTmp = Split(arrFieldsTmp(intFields) & " ", " ")
                strFieldName = Trim(arrTmp(0)): strFieldNameAlias = Trim(arrTmp(1))
                If IsNumeric(strFieldName) Then strFieldName = rsClone.Fields(Val(strFieldName)).name & ""
                '��ȡ�ֶ�ԭ������������
                arrFieldsName(UBound(arrFieldsName)) = strFieldName
                '����ֶ�,�������ڱ������������е�����Ϊ����
                .Fields.Append IIf(strFieldNameAlias = "", strFieldName, strFieldNameAlias), IIf(rsClone.Fields(strFieldName).Type = adNumeric, adDouble, rsClone.Fields(strFieldName).Type), rsClone.Fields(strFieldName).DefinedSize, adFldIsNullable '0:��ʾ����
            End If
        Next
        
        '׷���ֶ����
        If TypeName(arrAppFields) = "Variant()" Then
            For i = LBound(arrAppFields) To UBound(arrAppFields) Step 4
                If arrAppFields(i + 2) = Empty Then
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable, arrAppFields(i + 3)
                    End If
                Else
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable, arrAppFields(i + 3)
                    End If
                End If
            Next
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '��������
        If Not blnOnlyStructure Then
            If rsClone Is Nothing Then Set CopyNewRec = rsTarget: Exit Function
            If rsClone.RecordCount <> 0 Then rsClone.MoveFirst
            Do While Not rsClone.EOF
                .AddNew
                For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                    '�¼�¼�����а�˳����ӣ���˿�������
                    .Fields(intFields).value = rsClone.Fields(arrFieldsName(intFields)).value
                Next
                .Update
                rsClone.MoveNext
            Loop
            If rsClone.RecordCount <> 0 Then .Filter = "": .MoveFirst
        End If
    End With
    
    Set CopyNewRec = rsTarget
End Function

Public Function RecDelete(ByRef rsInput As ADODB.Recordset, Optional ByVal strFilter As String) As Boolean
'���ܣ�ɾ��ָ�������ļ�¼���ļ�¼
'������rsInput=��¼��
'      strFilter=����
'���أ��Ƿ�ɹ�
'      rsInput=����ɾ����ļ�¼��
    rsInput.Filter = strFilter
    If rsInput.RecordCount > 0 Then
        rsInput.MoveFirst
        Do While Not rsInput.EOF
            Call rsInput.Delete
            rsInput.MoveNext
        Loop
        Call rsInput.UpdateBatch
    End If
    RecDelete = True
End Function

Public Function RecUpdate(ByRef rsInput As Recordset, ByVal strFilter As String, ParamArray arrInput() As Variant) As Boolean
'���ܣ�����ָ�������ļ�¼���ļ�¼
'������rsInput=��¼��
'      strFilter=����
'      arrInput=������ֶ����Լ�ֵ����ʽ���ֶ���1,ֵ1, �ֶ���2,ֵ2,....
'���أ��Ƿ�ɹ�
'      rsInput=�������º�ļ�¼��
'˵����arrInput���ֶ�ֵ�����ü�¼���е������ֶ������¸��ֶΣ���ʱ��ʽΪ��!�ֶ��� ������(��ʱ֧��Val)
    Dim strFiledName As String, strFileValue As String, strFun As String, strFindFiled As String
    Dim blnFiled As Boolean, i As Long
    Dim arrTmp As Variant
    
    If rsInput Is Nothing Then Exit Function
    On Error GoTo errH
    With rsInput
        .Filter = strFilter
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            For i = LBound(arrInput) To UBound(arrInput) Step 2
                strFiledName = arrInput(i)
                If arrInput(i + 1) & "" = "" Then
                    rsInput(strFiledName).value = Null
                Else
                    strFun = ""
                    strFindFiled = arrInput(i + 1)
                    If arrInput(i + 1) Like "!?*" Then
                        blnFiled = True
                        On Error Resume Next
                        strFindFiled = Mid(arrInput(i + 1), 2)
                        arrTmp = Split(strFindFiled & " ", " ")
                        strFindFiled = Trim(arrTmp(0))
                        strFun = Trim(arrTmp(1))
                        strFileValue = rsInput(strFindFiled).value & ""
                        If err.Number <> 0 Then err.Clear: blnFiled = False
                        On Error GoTo errH
                    End If
                    If Not blnFiled Then
                        rsInput(strFiledName).value = arrInput(i + 1)
                    Else
                        If strFun = "" Then
                            rsInput(strFiledName).value = rsInput(strFindFiled).value
                        ElseIf strFun = "Val" Then
                            rsInput(strFiledName).value = Val(rsInput(strFindFiled).value & "")
                        ElseIf strFun = "Trim" Then
                            rsInput(strFiledName).value = Trim(rsInput(strFindFiled).value & "")
                            If rsInput(strFiledName).value & "" = "" Then
                                rsInput(strFiledName).value = Null
                            End If
                        Else
                            rsInput(strFiledName).value = rsInput(strFindFiled).value
                        End If
                    End If
                End If
                blnFiled = False
            Next
            .MoveNext
        Loop
        Call rsInput.UpdateBatch
    End With
    RecUpdate = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function RecDataAppend(ByRef rsSource As ADODB.Recordset, ByVal rsAppend As ADODB.Recordset, ParamArray arrInput() As Variant) As Boolean
'���ܣ���ָ����¼����������ӵ���һ����¼����
'������rsSource=Ŀ���¼��
'      rsAppend=���ݼ�¼��
'      arrInput=�ֶζ�Ӧ���򣬸ò�������ʱ��Ĭ������¼���ṹ��ͬ����ʽ��arrInput(0):[��¼��1].�ֶ�1,�ֶ�2...��arrInput(1)��[��¼��2].�ֶ�1,�ֶ�2...
'���أ��Ƿ�ɹ�
'      rsSource=������ݺ�ļ�¼��
    Dim arrSource As Variant, arrAppend As Variant
    Dim i As Long, arrValues() As Variant
    Dim strTmp As String
    
    If rsAppend Is Nothing Then RecDataAppend = True: Exit Function
    If rsAppend.RecordCount = 0 Then RecDataAppend = True: Exit Function
    If rsSource Is Nothing Then Set rsSource = rsAppend: RecDataAppend = True: Exit Function
    On Error GoTo errH
    If LBound(arrInput) = 2 Then
        '�˶δ�����Ҫ������ϸ����
        arrSource = Split(arrInput(LBound(arrInput)), ",")
        arrAppend = Split(arrInput(UBound(arrInput)), ",")
        If UBound(arrSource) <> UBound(arrAppend) Then Exit Function
        ReDim arrValues(UBound(arrAppend)): rsAppend.MoveFirst
        Do While Not rsAppend.EOF
            For i = LBound(arrAppend) To UBound(arrAppend)
                arrValues(i) = rsAppend(arrAppend(i)).value
            Next
            rsSource.AddNew arrSource, arrValues
            Erase arrValues
            rsAppend.MoveNext
        Loop
    ElseIf LBound(arrInput) = 0 Then
        strTmp = ""
        For i = 0 To rsSource.Fields.Count - 1
            strTmp = strTmp & "," & rsSource.Fields(i).name
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        arrSource = Split(strTmp, ",")
        On Error Resume Next
        If rsAppend.RecordCount <> 0 Then rsAppend.MoveFirst
        Do While Not rsAppend.EOF
            rsSource.AddNew
            For i = LBound(arrSource) To UBound(arrSource)
                rsSource.Fields(arrSource(i)).value = rsAppend.Fields(arrSource(i)).value
            Next
            rsSource.Update
            rsAppend.MoveNext
        Loop
        If err.Number <> 0 Then err.Clear
        On Error GoTo errH
    End If
    
    RecDataAppend = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Public Function RecDistinct(ByVal rsSource As ADODB.Recordset, Optional ByVal strDisFieldsName As String, Optional ByVal strFieldsName As String) As ADODB.Recordset
'���ܣ���¼��ȥ�ظ�
'������rsSource=Ҫȥ�ظ��ļ�¼��
'strDisFieldsName=ȥ�ظ����ֶ�,Ϊ�գ���������ֶ�ȥ��
'strFieldsName=���ؽ�����ֶΣ�Ϊ�գ��򷵻�ȥ�ظ����ֶ�
'���أ�������ļ�¼��
    Dim rsReturn As ADODB.Recordset
    Dim arrFilds As Variant, arrValues As Variant
    Dim i As Long, j As Long
    Dim strTmp As String, strOldRow As String

    '��ȡĬ���ֶ���
    If strDisFieldsName = "" Then
        For i = 0 To rsSource.Fields.Count - 1
            strTmp = strTmp & "," & rsSource.Fields(i).name
        Next
        strTmp = Mid(strTmp, 2)
        If strDisFieldsName = "" Then strDisFieldsName = strTmp
    End If
    If strFieldsName = "" Then strFieldsName = strDisFieldsName
    
    Set rsReturn = CopyNewRec(rsSource, , strFieldsName)
    If rsSource.RecordCount = 0 Then Set RecDistinct = rsReturn: Exit Function
    
    rsReturn.Sort = strDisFieldsName '�����Զ�������ƶ�����ͷ
    Do While Not rsReturn.EOF
        strTmp = rsReturn.GetString(, 1, "[ColumnSpliter]", , "[NULLEXP]") '�Զ��ƶ����
        rsReturn.MovePrevious
        If strTmp = strOldRow Then  'ɾ���ظ���
            Call rsReturn.Delete: Call rsReturn.Update
        Else
            strOldRow = strTmp
        End If
        rsReturn.MoveNext
    Loop
    rsReturn.Sort = strDisFieldsName
    Set RecDistinct = rsReturn
End Function

Public Function GetALLPars(Optional ByVal lngSys As Long = -1, Optional ByVal blnDetails As Boolean = True, Optional ByVal blnAddSets As Boolean) As ADODB.Recordset
'��ȡ���в���
'������blnDetails=True ,��ȡ���Ų����������飬����˽�в�����������,false-ֻ��ȡ�����б�
'         lngSys=-1-��ȡ����ϵͳ��<>0��ȡĳһ��ϵͳ,=-9����ȡϵͳ��Ϣ
'         blnAddSets=�Ƿ�����������Ϣ
' ���أ���ȡ�Ĳ�����¼��

    Dim strSql As String
    Dim rsParas As ADODB.Recordset
    If lngSys <> -9 Then
        '���в�����Ϣ
        strSql = "Select 0 ����, Nvl(a.ϵͳ, 0) ϵͳ," & vbNewLine & _
                        "       Nvl(a.ģ��, 0) ģ��, Nvl(a.˽��, 0) ˽��, Nvl(a.����, 0) ����,Nvl(a.����, 0) ����, Nvl(a.����, 0) ����,  Nvl(a.��Ȩ, 0) ��Ȩ, Nvl(a.�̶�, 0) �̶�, a.������, a.������, a.����ֵ, a.ȱʡֵ, a.Ӱ�����˵��, a.����ֵ����, a.����˵��, a.����˵��," & vbNewLine & _
                        "       a.����˵��, 0 ����id, Null �û���, Null ������," & vbNewLine & _
                        "       Null ��ϸ����ֵ" & vbNewLine & _
                        "From zlParameters A" & IIf(lngSys = -1, "", " Where Nvl(a.ϵͳ, 0)=[1]")
        If blnDetails Then
            '���Ų�������
            strSql = strSql & vbNewLine & _
                            "Union All " & vbNewLine & _
                            "Select 1 ����, Nvl(a.ϵͳ, 0) ϵͳ, Nvl(a.ģ��, 0) ģ��, Null  ˽��,Null  ����, Null ����, Null ����,Null  ��Ȩ, Null �̶�, a.������, a.������, Null ����ֵ, Null ȱʡֵ," & vbNewLine & _
                            "       Null Ӱ�����˵��, Null ����ֵ����, Null ����˵��, Null ����˵��,Null ����˵��," & vbNewLine & _
                            "       b.����id ����id, Null �û���, Null ������, b.����ֵ  ��ϸ����ֵ" & vbNewLine & _
                            "From zlParameters A, Zldeptparas B" & vbNewLine & _
                            "Where a.Id = b.����id And Nvl(a.����, 0) = 1" & IIf(lngSys = -1, "", " And Nvl(a.ϵͳ, 0)=[1]")
            '˽�б�����������
            strSql = strSql & vbNewLine & _
                            "Union All " & vbNewLine & _
                            "Select 2 ����, Nvl(a.ϵͳ, 0) ϵͳ, Nvl(a.ģ��, 0) ģ��, Null  ˽��,  Null ����,  Null ����,Null  ����,Null  ��Ȩ, Null  �̶�,a.������, a.������, Null ����ֵ,Null ȱʡֵ," & vbNewLine & _
                            "       Null Ӱ�����˵��, Null ����ֵ����, Null ����˵��, Null ����˵��, Null ����˵��," & vbNewLine & _
                            "       Null ����id, c.�û��� �û���, c.������ ������, c.����ֵ ��ϸ����ֵ" & vbNewLine & _
                            "From zlParameters A, zlUserParas C" & vbNewLine & _
                            "Where a.Id = c.����id And Nvl(a.����, 0) = 0 And (Nvl(a.˽��, 0) = 1 Or Nvl(a.����, 0) = 1)" & IIf(lngSys = -1, "", " And Nvl(a.ϵͳ, 0)=[1]")
            
        End If
    End If
    'ϵͳ��Ϣ��������Ϣ������ͳ����Ϣ
    strSql = IIf(lngSys <> -9, strSql & vbNewLine & _
                    "Union All " & vbNewLine, "") & _
                    "Select -9 ����, Nvl(���, 0) ϵͳ, Null ģ��, Null ˽��, Null ����,  Null ����,  Null  ����, Null ��Ȩ,Null �̶�,Null ������, ���� ������, �汾�� ����ֵ, b.���� || ''   ȱʡֵ," & vbNewLine & _
                    "       Null Ӱ�����˵��, Null  ����ֵ����, Null ����˵��, Null ����˵��, Null ����˵��," & vbNewLine & _
                    "       Null ����id, Null �û���, Null ������, Null ��ϸ����ֵ" & vbNewLine & _
                    "From (Select ���, ����, �汾��" & vbNewLine & _
                    "       From zlSystems" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select 0, '������������', ����" & vbNewLine & _
                    "       From zlRegInfo" & vbNewLine & _
                    "       Where ��Ŀ = '�汾��') A, (Select Nvl(ϵͳ, 0) ϵͳ, Count(1) ���� From zlParameters Group By Nvl(ϵͳ, 0)) B" & vbNewLine & _
                    "Where a.��� = b.ϵͳ" & IIf(lngSys = -1 Or lngSys = -9, "", " And Nvl(a.���, 0)=[1]")
    If blnAddSets And lngSys <> -9 Then
        '������Ϣ
        strSql = strSql & vbNewLine & _
                        "Union All " & vbNewLine & _
                        "Select -99 ����, Null ϵͳ, Null ģ��, " & IIf(blnDetails, 1, 0) & " ˽��, Null ����, Null ����, Null  ����, Null ��Ȩ,Null �̶�,  " & lngSys & " ������, Null ������, To_Char(Sysdate, 'yyyy-mm-dd HH24:mi:ss')  ����ֵ, User ȱʡֵ," & vbNewLine & _
                        "       Null Ӱ�����˵��, Null ����ֵ����, Null ����˵��, Null ����˵��, Null ����˵��," & vbNewLine & _
                        "       Null ����id, Null �û���, Null ������, Null ��ϸ����ֵ" & vbNewLine & _
                        "From Dual"
    End If
    'Ƕ�ײ�����
    strSql = "Select D.ϵͳ||'#'||D.ģ��||'#'||D.������ MainKey,D.*" & vbNewLine & _
                    "From (" & strSql & ") D" & vbNewLine & _
                    "Order By ����, ϵͳ, ģ��, ������"
    '��������ؼ���
    strSql = "Select RowNum SortKey,E.*" & vbNewLine & _
                "From (" & strSql & ") E" & vbNewLine & _
                "Order By ����, ϵͳ, ģ��, ������"
    Set rsParas = gclsBase.OpenSQLRecord(gcnOracle, strSql, "��ȡ���в���", lngSys)
    Set GetALLPars = rsParas
End Function

Public Function GetCompareRec(ByVal rsSouce As ADODB.Recordset, ByVal rsCompare As ADODB.Recordset, ByVal strKeyFields As String, Optional ByVal strComPareFileds As String, Optional ByVal strAddtionFileds As String, Optional arrAppFields As Variant) As ADODB.Recordset
'���ܣ���ȡ��¼���ȽϽ����¼��
'������rsSouce=�Ƚϼ�¼��
'         rsCompare=�Աȼ�¼��
'         strComPareFileds=���жԱȵ��ֶ�,�ֶ�����֮���Զ��ŷָΪ��ֵ��ʾ��rsSouce���ֶ���Ϊ�Ա��ֶ�,"-�ֶδ�"����ʶ���ֶδ����ֶβ�����Ƚ�
'         strAddtionFileds=�����Ƚϼ�¼�����ǲ����бȽ��жϣ���Щ�ֶ�Ϊ�˷���Ƚϼ�¼����ʹ��
'         strKeyFields���ֶ������Զ��ŷָ�����ֶ�,��ʽΪ:Nvl(�ֶ�1,0)_Nvl(�ֶ�2,0)_Nvl(�ֶ�3,0)...
'         arrAppFields=׷�ӵ��ֶ���Ϣ������,����,����,Ĭ��ֵ,û��Ĭ��ֵ��Empty,û��ָ�����ȴ�Empty
    Dim i As Long
    Dim strFileds As String
    Dim varKey As Variant, varCom As Variant, varAddtion As Variant
    Dim rsReturn As ADODB.Recordset, rsSort As ADODB.Recordset
    Dim strTmpKey As String, strPreKey As String
    Dim blnNew As Boolean, strDifCols As String
    Dim strNotCom As String
    Dim cllNumCol As Collection '��ֵ�У������п�ֵ����0
    Dim intState As Integer
    
    On Error GoTo errH
    If strKeyFields = "" Then Exit Function
    Set rsReturn = New ADODB.Recordset
    Set cllNumCol = New Collection
    
    With rsReturn
        .Fields.Append "MainKey", adVarChar, 200, adFldIsNullable
        .Fields.Append "State", adInteger '-1-ɾ����0-���䣬1-������2-����
        .Fields.Append "DifInfo", adVarChar, 2000, adFldIsNullable
        .Fields.Append "Sort", adInteger, Empty, Empty '��������ɾ��������֮�䣬-1ɾ����0-����,1-����,2-����
        varKey = Split(strKeyFields, ",") '�����ֶ�
        varCom = Split(strComPareFileds & "-", "-")
        strComPareFileds = UCase(Trim(varCom(0))) '�Ƚ��ֶ�
        strNotCom = UCase(Trim(varCom(1))) '���Ƚ��ֶ�
        varAddtion = Split(strAddtionFileds, ",") '�����ֶ�
        If strComPareFileds = "" Then
            For i = 0 To rsSouce.Fields.Count - 1
                '�Ƚ��ֶβ���������,���Ƚ��ֶ��븽���ֶ�
                If InStr("," & strKeyFields & ",", "," & rsSouce.Fields(i).name & ",") = 0 And InStr("," & strNotCom & ",", "," & rsSouce.Fields(i).name & ",") = 0 And InStr("," & strAddtionFileds & ",", "," & rsSouce.Fields(i).name & ",") = 0 Then
                    strComPareFileds = strComPareFileds & IIf(strComPareFileds = "", "", ",") & rsSouce.Fields(i).name
                End If
            Next
        Else
            strComPareFileds = "," & strComPareFileds & ","
            If strNotCom <> "" Then
                varCom = Split(strNotCom, ",")
                For i = LBound(varCom) To UBound(varCom)
                    strComPareFileds = Replace(strComPareFileds, "," & varCom(i) & ",", ",")
                Next
            End If
            For i = LBound(varKey) To UBound(varKey)
                strComPareFileds = Replace(strComPareFileds, "," & varKey(i) & ",", ",")
            Next
            For i = LBound(varAddtion) To UBound(varAddtion)
                strComPareFileds = Replace(strComPareFileds, "," & varAddtion(i) & ",", ",")
            Next
            If strComPareFileds = "," Then
                strComPareFileds = ""
            Else
                strComPareFileds = Mid(strComPareFileds, 2, Len(strComPareFileds) - 2)
            End If
        End If
        If strComPareFileds = "" Then Exit Function '�ޱȽ��ֶΣ����ܽ��бȽ�
        varCom = Split(strComPareFileds, ",")
        For i = LBound(varCom) To UBound(varCom)
            If IsType(rsSouce.Fields(varCom(i)).Type, adNumeric) Then
                cllNumCol.Add 1, varCom(i)
            Else
                cllNumCol.Add 0, varCom(i)
            End If
            'ԭʼ�ֶ�
            .Fields.Append varCom(i), IIf(rsSouce.Fields(varCom(i)).Type = adNumeric, adDouble, rsSouce.Fields(varCom(i)).Type), rsSouce.Fields(varCom(i)).DefinedSize, adFldIsNullable
            '���ֶ�
            .Fields.Append varCom(i) & "_New", IIf(rsSouce.Fields(varCom(i)).Type = adNumeric, adDouble, rsSouce.Fields(varCom(i)).Type), rsSouce.Fields(varCom(i)).DefinedSize, adFldIsNullable
        Next
        '����Դ���ֶΣ�����Ϊ����������ӵ���¼���������м�¼���ȶ�
        For i = LBound(varAddtion) To UBound(varAddtion)
            'ԭʼ�ֶ�
            .Fields.Append varAddtion(i), IIf(rsSouce.Fields(varAddtion(i)).Type = adNumeric, adDouble, rsSouce.Fields(varAddtion(i)).Type), rsSouce.Fields(varAddtion(i)).DefinedSize, adFldIsNullable
            '���ֶ�
            .Fields.Append varAddtion(i) & "_New", IIf(rsSouce.Fields(varAddtion(i)).Type = adNumeric, adDouble, rsSouce.Fields(varAddtion(i)).Type), rsSouce.Fields(varAddtion(i)).DefinedSize, adFldIsNullable
        Next
        '׷���ֶ����
        If TypeName(arrAppFields) = "Variant()" Then
            For i = LBound(arrAppFields) To UBound(arrAppFields) Step 4
                If arrAppFields(i + 2) = Empty Then
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable, arrAppFields(i + 3)
                    End If
                Else
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable, arrAppFields(i + 3)
                    End If
                End If
            Next
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '������,���ܻ�  �����ڴ治�㣬�������ö��ֶ�����
'        rsSouce.Sort = strKeyFields
'        rsCompare.Sort = strKeyFields
        Set rsSort = CopyNewRec(Nothing, , , Array("MainKey", adVarChar, 200, Empty, "����", adInteger, 1, 0, "BookMark", adDouble, Empty, Empty))
        '��������
        If rsSouce.RecordCount <> 0 Then rsSouce.MoveFirst
        Do While Not rsSouce.EOF
            strTmpKey = ""
            For i = LBound(varKey) To UBound(varKey)
                strTmpKey = strTmpKey & IIf(strTmpKey = "", "", "#") & Nvl(rsSouce.Fields(varKey(i)).value, 0)
            Next
            rsSort.AddNew Array("MainKey", "����", "BookMark"), Array(strTmpKey, 0, rsSouce.Bookmark)
            rsSouce.MoveNext
        Loop
        If rsSouce.RecordCount <> 0 Then rsSouce.MoveFirst
        '��������
        If rsCompare.RecordCount <> 0 Then rsCompare.MoveFirst
        Do While Not rsCompare.EOF
            strTmpKey = ""
            For i = LBound(varKey) To UBound(varKey)
                strTmpKey = strTmpKey & IIf(strTmpKey = "", "", "#") & Nvl(rsCompare.Fields(varKey(i)).value, 0)
            Next
            rsSort.AddNew Array("MainKey", "����", "BookMark"), Array(strTmpKey, 1, rsCompare.Bookmark)
            rsCompare.MoveNext
        Loop
        If rsCompare.RecordCount <> 0 Then rsCompare.MoveFirst
        rsSort.Sort = "MainKey"
        Do While Not rsSort.EOF
            strTmpKey = rsSort!MainKey
            blnNew = rsSort!���� = 1
            If blnNew Then
                rsCompare.Bookmark = CDbl(rsSort!Bookmark)
            Else
                rsSouce.Bookmark = CDbl(rsSort!Bookmark)
            End If
            If strPreKey <> strTmpKey Then
                .AddNew Array("MainKey", "State", "Sort"), Array(strTmpKey, 0, 0) '�����仯��������һ��
                strPreKey = strTmpKey
            End If
            intState = Val(!State) + IIf(blnNew, 1, -1)
            .Update Array("State", "Sort"), Array(intState, IIf(intState = 1, 2, intState)) '��������������ɾ�����¾��������еĺ����ж��Ǹı�
            On Error Resume Next
            '�Ƚ��������
            For i = LBound(varCom) To UBound(varCom)
                If blnNew Then
                    .Update varCom(i) & "_New", rsCompare.Fields(varCom(i)).value
                Else
                    .Update varCom(i), rsSouce.Fields(varCom(i)).value
                End If
            Next
            '�����������
            For i = LBound(varAddtion) To UBound(varAddtion)
                If blnNew Then
                    .Update varAddtion(i) & "_New", rsCompare.Fields(varAddtion(i)).value
                Else
                    .Update varAddtion(i), rsSouce.Fields(varAddtion(i)).value
                End If
            Next
            If err.Number <> 0 Then err.Clear
            On Error GoTo errH
            rsSort.MoveNext
        Loop
        '�Ƚ�ϸ΢����
        .Filter = "State=0": .Sort = "MainKey"
        Do While Not .EOF
            strDifCols = ""
            For i = LBound(varCom) To UBound(varCom)
                If cllNumCol(varCom(i)) = 1 Then
                    If Val(.Fields(varCom(i) & "_New").value & "") <> Val(.Fields(varCom(i)).value & "") Then  '��ȡ������
                        strDifCols = strDifCols & IIf(strDifCols = "", "", ",") & varCom(i)
                    End If
                Else
                    If .Fields(varCom(i) & "_New").value & "" <> .Fields(varCom(i)).value & "" Then '��ȡ������
                        strDifCols = strDifCols & IIf(strDifCols = "", "", ",") & varCom(i)
                    End If
                End If
            Next
            If strDifCols <> "" Then
                .Update Array("State", "DifInfo", "Sort"), Array(2, strDifCols, 1)
            End If
            .MoveNext
        Loop
        .Filter = ""
    End With
    Set GetCompareRec = rsReturn
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Public Function ZVal(ByVal varValue As Variant) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Public Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'���ܣ��ж�ĳ��ADO�ֶ����������Ƿ���ָ���ֶ�������ͬһ��(������,����,�ַ�,������)
    Dim intA As Integer, intB As Integer
    
    Select Case varBase
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intA = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intA = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intA = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intA = -4
        Case Else
            intA = varBase
    End Select
    Select Case varType
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intB = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intB = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intB = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intB = -4
        Case Else
            intB = varType
    End Select
    IsType = intA = intB
End Function

Public Function SQLAdjust(ByVal varInput As Variant) As String
'���ܣ�������"'"���ŵ��ַ�������ΪOracle����ʶ����ַ�����,�����մ�ת��ΪNull
'˵�����Զ�(����)�����߼�"'"�綨����

    Dim i As Long, strTmp As String, strOneChar As String
    Dim strReturn As String
    Dim lngLine As Long
    
    strReturn = varInput & ""
    If strReturn & "" = "" Then SQLAdjust = "Null": Exit Function
    If InStr(1, strReturn, "'") = 0 And InStr(1, strReturn, Chr(10)) = 0 And InStr(1, strReturn, Chr(13)) = 0 Then SQLAdjust = "'" & strReturn & "'": Exit Function
    
    For i = 1 To Len(strReturn)
        strOneChar = Mid(strReturn, i, 1)
        Select Case strOneChar
            Case "'"
                If i = 1 Then
                    strTmp = "CHR(39)||'"
                ElseIf i = Len(strReturn) Then
                    strTmp = strTmp & "'||CHR(39)"
                Else
                    strTmp = strTmp & "'||CHR(39)||'"
                End If
                lngLine = lngLine + 1 '��ʶ�зǻ����ַ�
            Case Chr(10), Chr(13)
                If i = 1 Then
                    strTmp = "CHR(13)||'"
                ElseIf lngLine = 0 Then '���Ŷ�����У�����һ��
                    If i = Len(strReturn) Then '���һ���ǻ���
                        strTmp = strTmp & "'"
                    End If
                ElseIf i = Len(strReturn) Then
                    strTmp = strTmp & "'||CHR(13)"
                Else
                    strTmp = strTmp & "'||CHR(13)||'"
                End If
                lngLine = 0 '��ʶ�Ѿ��л���
            Case Else
                If i = 1 Then
                    strTmp = "'" & Mid(strReturn, i, 1)
                ElseIf i = Len(strReturn) Then
                    strTmp = strTmp & Mid(strReturn, i, 1) & "'"
                Else
                    strTmp = strTmp & Mid(strReturn, i, 1)
                End If
                lngLine = lngLine + 1 '��ʶ�зǻ����ַ�
        End Select
    Next
    SQLAdjust = strTmp
End Function

Public Sub SetCtrlPosOnLine(ByVal blnvertical As Boolean, ByVal intAligType As Integer, ParamArray arrControls() As Variant)
'����:��ͬһ�еĿؼ�����λ������
'������
'blnvertical  true ,��ֱ�������ÿؼ�λ�ã�false,ˮƽ�������ÿؼ�λ��
'blnvertical=false :intAligType=-1,���˶��룬0-�м���룬1-�׶˶���,blnvertical=true,intAligType=-1,����룬0-ˮƽ���Ķ��룬1-�Ҷ���
'   arrControls��ʽΪ�ؼ�1,���1,�ؼ�2,���2,�ؼ�3,...
    Dim i As Long
    Dim lngPos As Long '��һ���ؼ���ĳһλ��
    Dim dblRate As Double
    If UBound(arrControls) = -1 Then Exit Sub
    If blnvertical Then
        Select Case intAligType
            Case -1
                lngPos = arrControls(0).Left
                dblRate = 0
            Case 0
                lngPos = arrControls(0).Left + 0.5 * arrControls(0).Width
                dblRate = 0.5
            Case 1
                lngPos = arrControls(0).Left + arrControls(0).Width
                dblRate = 1
        End Select
        
        For i = 0 To UBound(arrControls)
            If i > 0 And i Mod 2 = 0 Then
                arrControls(i).Top = arrControls(i - 2).Top + arrControls(i - 2).Height + arrControls(i - 1)
                arrControls(i).Left = lngPos - arrControls(i).Width * dblRate
            End If
        Next
    Else
        Select Case intAligType
            Case -1
                lngPos = arrControls(0).Top
                dblRate = 0
            Case 0
                lngPos = arrControls(0).Top + 0.5 * arrControls(0).Height
                dblRate = 0.5
            Case 1
                lngPos = arrControls(0).Top + arrControls(0).Height
                dblRate = 1
        End Select
        
        For i = 0 To UBound(arrControls)
            If i > 0 And i Mod 2 = 0 Then
                arrControls(i).Left = arrControls(i - 2).Left + arrControls(i - 2).Width + arrControls(i - 1)
                arrControls(i).Top = lngPos - arrControls(i).Height * dblRate
            End If
        Next
    End If
End Sub

Public Sub SetCtrlSameDistance(ByVal blnvertical As Boolean, ByVal intSameType As Integer, ByVal intAligType As Integer, ParamArray arrControls() As Variant)
'����:��һ��ؼ�����Ϊ��ͬ�ļ��
'������
'blnvertical  true ,��ֱ�������ÿؼ�λ�ã�false,ˮƽ�������ÿؼ�λ��
'intAligType=2������
'blnvertical=false :intAligType=-1,���˶��룬0-�м���룬1-�׶˶���,blnvertical=true,intAligType=-1,����룬0-ˮƽ���Ķ��룬1-�Ҷ���
'intSameType=0:�߽�����ͬ��1-���ļ����ͬ,
'arrControls��ʽΪ�ؼ�1,�ؼ�2,�ؼ�3,...
'˵��������λ�����ؼ���Ϊ��׼���Զ������м�ؼ����
    Dim i As Long, lngSart As Long, lngEnd As Long
    Dim lngSum As Long, lngDistance As Long
    Dim lngPos As Long
    Dim dblSameRate As Double, dblAligRate As Double
    
    If UBound(arrControls) < 2 Then Exit Sub '���������ؼ�������
    '��ȡ������
    dblSameRate = IIf(intSameType = 1, 0.5, 1)
    dblAligRate = intAligType / 2
    '������ʼλ��
    If blnvertical Then
        lngSart = arrControls(0).Top + dblSameRate * arrControls(0).Height
        lngEnd = arrControls(UBound(arrControls)).Top + (1 - dblSameRate) * arrControls(UBound(arrControls)).Height
    Else
        lngSart = arrControls(0).Left + dblSameRate * arrControls(0).Width
        lngEnd = arrControls(UBound(arrControls)).Left + (1 - dblSameRate) * arrControls(UBound(arrControls)).Width
    End If
    '��ȡ��Ҫ�޳�����Ч���
    If intSameType = 0 Then '�ؼ���߽�����ͬ
        For i = 1 To UBound(arrControls) - 1
            lngSum = lngSum + IIf(blnvertical, arrControls(i).Height, arrControls(i).Width)
        Next
    Else
        lngSum = 0
    End If
    '��ȡ����λ��
    If intAligType <> 2 Then
        If blnvertical Then
            lngPos = arrControls(0).Left + (dblAligRate + 0.5) * arrControls(0).Width
        Else
            lngPos = arrControls(0).Top + (dblAligRate + 0.5) * arrControls(0).Height
        End If
    End If
    '��ȡƽ�����
    lngDistance = (lngEnd - lngSart - lngSum) / UBound(arrControls)
    '���ÿؼ�λ��
    For i = 1 To UBound(arrControls)
        If blnvertical Then
            arrControls(i).Top = lngSart + lngDistance - (1 - dblSameRate) * arrControls(i).Height
            lngSart = arrControls(i).Top + arrControls(i).Height - (1 - dblSameRate) * arrControls(i).Height
            If intAligType <> 2 Then arrControls(i).Left = lngPos - (dblAligRate + 0.5) * arrControls(0).Width
        Else
            arrControls(i).Left = lngSart + lngDistance - (1 - dblSameRate) * arrControls(i).Width
            lngSart = arrControls(i).Left + arrControls(i).Width - (1 - dblSameRate) * arrControls(i).Width
            If intAligType <> 2 Then arrControls(i).Top = lngPos - (dblAligRate + 0.5) * arrControls(0).Height
        End If
    Next
End Sub

Public Sub SetCtrlEnabled(ByVal blnEnabled As Boolean, ParamArray arrControls() As Variant)
'����:��һ���ؼ���Enabled���Խ�������
'������
'blnEnabled  true ,�ռ���ã�false,�ռ䲻����
'arrControls��ʽΪ�ؼ�1,�ؼ�2,�ؼ�3,...

    Dim i As Long
    
    For i = LBound(arrControls) To UBound(arrControls)
        arrControls(i).Enabled = blnEnabled
    Next
End Sub


Public Function CancelNetServer(ByVal strPath As String) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '����:�Ͽ�����������
    '����:
    '����:���ҳɹ�,����true,���򷵻�False
    '----------------------------------------------------------------------------------------------------------
    err = 0
    On Error Resume Next
    If WNetCancelConnection2(strPath, CONNECT_UPDATE_PROFILE, True) = 0 Then
        CancelNetServer = True
    Else
        CancelNetServer = False
    End If
    err = 0
End Function

Public Function IsNetServer(ByVal strPath As String, ByVal strUser As String, ByVal strPassword As String) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '--����:���������Ƿ�����������
    '--����:strPath -����·��
    '       strUser-�û���
    '       strPassWord -��������
    '����:����˳��,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/06
    '----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
      
    '���˺�:���ܴ���windows��Դ�������Ѿ��з��ʵ���
    '
'    If objFile.FolderExists(strPath) Then
'        IsNetServer = True: Exit Function
'    End If
    
    Dim NetR As NETRESOURCE
    With NetR
        .dwScope = RESOURCE_GLOBALNET
        .dwType = RESOURCETYPE_DISK
        .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
        .dwUsage = RESOURCEUSAGE_CONNECTABLE
        .lpLocalName = "" 'ӳ���������
        .lpRemoteName = strPath  '������·��
    End With
    
    err = 0
    On Error GoTo ErrHand:
    If WNetAddConnection2(NetR, strPassword, strUser, CONNECT_UPDATE_PROFILE) = NO_ERROR Then
       IsNetServer = True
    Else
       IsNetServer = False
    End If
    Exit Function
ErrHand:
       IsNetServer = False
End Function

Public Sub ShowFlash(Optional strInfo As String, Optional sngPer As Single = -1, Optional frmParent As Object, Optional blnPer As Boolean)
'���ܣ���ʾ�����صȴ�����ȴ���(strInfo)
'����:strInfo=�ȴ��������ʾ��Ϣ
'     sngPer=����
    Static blnShow As Boolean
    
    If sngPer > 1 Then sngPer = 1
    
    If strInfo = "" Then
        frmFlash.avi.Close
        Unload frmFlash
        blnShow = False
    Else
        If Not blnShow Then
            On Error Resume Next
            If sngPer = -1 Then
                '��ʾ�ȴ�
                frmFlash.avi.Open GetSetting("ZLSOFT", "ע����Ϣ", "gstrAviPath", "") & "\" & "Findfile.avi"
                If err.Number <> 0 Then
                    err.Clear
                End If
                frmFlash.lbl.Caption = strInfo
                
                If frmParent Is Nothing Then
                    SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                    ShowWindow frmFlash.hwnd, 5
                Else
                    err.Clear
                    frmFlash.Show , frmParent
                    If err.Number <> 0 Then
                        err.Clear
                        SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                        ShowWindow frmFlash.hwnd, 5
                    End If
                End If
                
                frmFlash.avi.Play
                frmFlash.Refresh
            Else
                '��ʾ����
                frmFlash.avi.Visible = False
                frmFlash.picDo.Visible = True
                frmFlash.lbl.Top = frmFlash.lbl.Top - frmFlash.lbl.Height / 2
                frmFlash.lbl.Left = frmFlash.picDo.Left
                frmFlash.lblPer.Top = frmFlash.lbl.Top
                frmFlash.lbl.Caption = strInfo
                frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)
                If blnPer Then
                    If sngPer > 0 Then
                        frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                    Else
                        frmFlash.lblPer.Caption = ""
                    End If
                    frmFlash.lblPer.Visible = True
                End If
                
                If frmParent Is Nothing Then
                    SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                    ShowWindow frmFlash.hwnd, 5
                Else
                    err.Clear
                    frmFlash.Show , frmParent
                    If err.Number <> 0 Then
                        err.Clear
                        SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / 15, (Screen.Height - frmFlash.Height) / 2 / 15, 0, 0, 1
                        ShowWindow frmFlash.hwnd, 5
                    End If
                End If
                
                frmFlash.Refresh
            End If
            blnShow = True
        Else
            frmFlash.lbl.Caption = strInfo
            If sngPer >= 0 Then
                frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)
                If sngPer > 0 Then
                    frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                Else
                    frmFlash.lblPer.Caption = ""
                End If
            End If
            frmFlash.Refresh
        End If
    End If
End Sub

Public Function GetLogPath(ByVal ltLogType As LogType, Optional ByVal strSysCodes As String, Optional ByVal strBakUser As String, Optional ByVal strFolder As String, Optional ByVal strName As String) As String
'���ܣ���ȡ��־Ŀ¼
'ltLogType=��־���ͣ�0-ϵͳ��װ��1-������Ǩ��־��2-��ǰ��Ǩ��־��3-��ʷ�ⵥ����Ǩ��־,4ϵͳ����
'strSysCodes=��Ҫ������ϵͳ��ϵͳ����,���ϵͳ�Զ��ŷָ�
'strBakUser=��ʷ�ⵥ������ʱ������ʷ����
'strFolder=�Զ�������־���ļ���
'strName=�Զ������͵���־��
'���أ���־�ļ����Լ�·��,strFileName��Ҫ���أ���Ҫ������
    Dim strFileName As String
    Dim arrTmp  As Variant, strSys As String
    Dim i As Long
    Dim strTime As String
    
    On Error GoTo errH
    If gblnInIDE Then
        strFolder = GetSetting("ZLSOFT", "����ȫ��", "����·��")
        strFolder = "C:\Appsoft\Log"
    Else
        strFolder = App.Path & "\Log"
    End If
    If Not gobjFile.FolderExists(strFolder) Then
        Call gobjFile.CreateFolder(strFolder)
    End If
    strTime = Format(Now, "YYMMDDHHmm")
    Select Case ltLogType
        Case LT_ϵͳ����
            strFolder = strFolder & "\ϵͳ����"
            strFileName = Mid(strTime, 1, 6) & ".Log"
        Case LT_������־
            strFolder = strFolder & "\��־����"
            strFileName = strTime & ".Log"
        Case LT_��װ, LT_������Ǩ, LT_��ʷ����Ǩ, LT_��ǰ��Ǩ
            strFolder = strFolder & "\��װ��Ǩ"
            arrTmp = Split(strSysCodes, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                strSys = strSys & Format(Val(arrTmp(i)) \ 100, "00")
            Next
            strFileName = Mid(strTime, 1, 6) & "_" & strSys & Decode(ltLogType, LT_��װ, "_Install", LT_������Ǩ, "", LT_��ǰ��Ǩ, "_BEF", LT_��ʷ����Ǩ, "_" & strBakUser) & "_" & Mid(strTime, 7, 4) & ".log"
        Case LT_�Զ���
            strFileName = strName & strTime & ".log"
    End Select
    If Not gobjFile.FolderExists(strFolder) Then
        Call gobjFile.CreateFolder(strFolder)
    End If
    GetLogPath = strFolder & "\" & strFileName
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Public Sub PressKey(bytKey As Byte)
'���ܣ�����̷���һ����,����SendKey
'������bytKey=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Function Identity(ByRef lngCount As Long) As Long
'���ܣ�ģ����������
'������lngCount=��������
    lngCount = lngCount + 1
    Identity = lngCount
End Function

Public Function GetOracleVersion(Optional ByVal blnGetVerNum As Boolean = True, Optional ByVal blnGetBigVer As Boolean) As Variant
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strTmp As String
    Dim arrTmp As Variant
    
    If gstrOracleVer = "" Then
        'CORE    10.2.0.3.0  Production
        strSql = "Select Banner From V$version Where Banner Like  'CORE%'"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, App.Title)
        If rsTmp.RecordCount > 0 Then
            arrTmp = Split(TrimEx(rsTmp!Banner & ""), " ")
            If UBound(arrTmp) = 2 Then
                gstrOracleVer = arrTmp(1)
            End If
        End If
    End If
    
    If gstrOracleVer <> "" Then
        If Not blnGetVerNum Then
            GetOracleVersion = gstrOracleVer
        Else
            If blnGetBigVer Then
                arrTmp = Split(gstrOracleVer, ".")
                GetOracleVersion = Val(arrTmp(0))
            Else
                GetOracleVersion = Val(Replace(Mid(gstrOracleVer, 4), ".", ""))
            End If
        End If
    Else
        GetOracleVersion = IIf(blnGetVerNum, 0, "��ȡʧ��")
    End If
End Function

Public Function ReadFileToString(ByVal strFile As String) As String
    Dim strBuffer As String
    Dim lngHwnd As Long
    Dim lngFileLen As Long

    lngHwnd = FreeFile

    On Error Resume Next
    Open strFile For Binary Shared As lngHwnd
    If err.Number <> 0 Then
        MsgBox "Error " & err.Number & vbCrLf & err.Description & vbCrLf & "Error in ReadFileToString, File='" & strFile & "'", vbCritical
        GoTo Proc_Exit
    End If
    On Error GoTo 0
    
    lngFileLen = LOF(lngHwnd)
    strBuffer = Space(lngFileLen)
    Get lngHwnd, , strBuffer
    
    Close lngHwnd
    
Proc_Exit:
    ReadFileToString = strBuffer
End Function

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Sub CboSetWidth(ByVal hWnd_combo As Long, ByVal lngWidth As Long)
'���ܣ�����Combo�ؼ������б�Ŀ��
'�˴��Ŀ�����������б�Ŀ�ȣ���������TWIPΪ��λ
    Const CB_SETDROPPEDWIDTH As Long = &H160

    SendMessage hWnd_combo, CB_SETDROPPEDWIDTH, lngWidth / Screen.TwipsPerPixelX, 0
End Sub

Public Sub CboSetIndex(ByVal hWnd_combo As Long, ByVal lngIndex As Long)
'���ܣ�����Combo�ؼ���Indexֵ
'Ϊһ��Combo�ؼ�ѡ���б�����ֲ�������Click�¼�
    Const CB_SETCURSEL = &H14E
    
    SendMessage hWnd_combo, CB_SETCURSEL, lngIndex, 0
End Sub

Public Sub WriteTraceLog(Optional ByVal strLog As String)
    If Not gblnTrace Then Exit Sub
    gobjLog.WriteLine strLog
End Sub

Public Function RestoreVsGridWidth(ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, ByVal strKey As String) As Boolean
    '------------------------------------------------------------------------------
    '����:�����ݿ��лָ�����Ŀ�ȵ���Ϣ
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     strKey-����
    '     blnSaveToDataBase-�Ƿ��������ݿ��б������(����������ݿ��б���,��ǿ�Ʊ���Ϊtrue,��������Ƿ�ʹ�ø��Ի������ȷ��)
    '     blnǿ�ƻָ�����-�����Ƿ񽫱���ע���Ĳ���ֵ,����ǿ�ƻָ�
    '����:�ָ��ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, arrTemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    
    strParaValue = Trim((GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & strCaption, strKey)))
    If strParaValue = "" Then Exit Function
    RestoreVsGridWidth = False
    
    'strParaValue:�����ʽ:������,�п�,������|������,�п�,������|...
    err = 0: On Error GoTo ErrHand:
    arrReg = Split(strParaValue, "|")
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            arrTemp = Split(arrReg(intCol) & ",,", ",")
            strColName = arrTemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(arrTemp(1))
                If Val(arrTemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    RestoreVsGridWidth = True
    Exit Function
ErrHand:
End Function

Public Function SaveVsGridWidth(ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, ByVal strKey As String) As Boolean
    '------------------------------------------------------------------------------
    '����:����vsFlex�Ŀ�ȵ�ע���
    '����:vsGrid-��Ӧ������ؼ�
    '     strKey-����
    '����:����ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------
    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '�����ʽ:������,�п�,������|������,�п�,������|...
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & strCaption, strKey, strCol)
    SaveVsGridWidth = True
End Function

Public Function GetClientPoint(ByVal lngHwnd As Long) As POINTAPI
'��ȡ��ǰָ���Ӧ�ڿؼ��е�λ��
    Dim pRet As POINTAPI
    Dim lngReturn As Long
    
    pRet = GetCursorPosition()
    lngReturn = ScreenToClient(lngHwnd, pRet)
    pRet.x = pRet.x * Screen.TwipsPerPixelX
    pRet.y = pRet.y * Screen.TwipsPerPixelY
    GetClientPoint = pRet
End Function

Public Function GetCursorPosition() As POINTAPI
'��ȡ���λ��
    Dim pRet As POINTAPI
    Dim lngReturn As Long
    lngReturn = GetCursorPos(pRet)
    GetCursorPosition = pRet
End Function
Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'���ܣ���ʾ������һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
    Dim vRect As RECT, LngStyle As Long
    
    Call GetWindowRect(objForm.hwnd, vRect)
    LngStyle = GetWindowLong(objForm.hwnd, GWL_STYLE)
    If blnCaption Then
        LngStyle = LngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then LngStyle = LngStyle Or WS_SYSMENU
        If objForm.MaxButton Then LngStyle = LngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then LngStyle = LngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            LngStyle = LngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            LngStyle = LngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong objForm.hwnd, GWL_STYLE, LngStyle
    SetWindowPos objForm.hwnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Public Function CompareFolder(ByVal strPath1 As String, ByVal strPath2 As String, ByVal strReports As String) As Boolean
'���ܣ��Ա����ļ������ļ�������ͬ�ļ����ļ����в���Աȣ��в��������ɱ��档
    Dim strCommand As String
    Dim lngProcess As Long
    Dim lngTemp As Long
    
    err.Clear
    strCommand = GetWinSystemPath & "\wincmp3.exe " & strPath1 & "\" & " " & strPath2 & "\" & " /G:HNISE " & strReports
    lngTemp = Shell(strCommand, vbHide)
    DoEvents
    If err <> 0 Then
        err.Clear
         MsgBox "�ļ��Ƚ�ʧ�ܣ�����" & GetWinSystemPath & "\wincmp3.exe�ļ��Ƿ����", vbExclamation, "�������"
        Exit Function
    End If
    lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
    Do
        Sleep 100
        GetExitCodeProcess lngProcess, lngTemp
    Loop While lngTemp = Still_Active
    CompareFolder = True
    err.Clear
    DoEvents
End Function

Public Function CollectionHave(ByVal Coll As Collection, ByVal strKey As String) As Boolean
    On Error GoTo ErrHand
    
    Dim Item As Variant
    Set Item = Coll.Item(strKey)
    CollectionHave = True
    Set Item = Nothing
    Exit Function
ErrHand:
    '�����ڷ���False
    If err.Number = 5 Then CollectionHave = False
    err.Clear
End Function

Public Function GetRIS() As Boolean
'���������ӿ�
    If Not gblnCreate Then Exit Function
    If Not gobjRIS Is Nothing Then GetRIS = True: Exit Function
    On Error Resume Next
    Set gobjRIS = CreateObject("zl9XWInterface.clsSvrTools")
    err.Clear: On Error GoTo 0
    GetRIS = Not gobjRIS Is Nothing
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As POINTAPI
    vPoint.x = lngX / Screen.TwipsPerPixelX: vPoint.y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.x = vPoint.x * Screen.TwipsPerPixelX: vPoint.y = vPoint.y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function


Public Function OpenIme(Optional blnOpen As Boolean = False, Optional strImeName As String) As Boolean
'����:���������뷨����ر����뷨
'������strImeName-��ָ�������뷨
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String
    
 
    '�û�û�������ã��Ͳ�����
    If blnOpen Then
        If strImeName <> "" Then
            strIme = strImeName
        End If
        If strIme = "" Then Exit Function                  'Ҫ������뷨��������û������
    End If
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))

    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '��Ҫ�����뷨�������ж��Ƿ�ָ�����뷨
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then
                        OpenIme = True
                        Exit Function
                    End If
                End If
            End If
        ElseIf blnOpen = False Then
            '�����������뷨��������Ӧ�˹ر����뷨������
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
        End If
    Loop Until lngCount = 0
    
    If blnOpen = False Then
        '����windows Vistaϵͳ��Ӣ�����뷨��ImmIsIME���Գ���1�����뷨,���,��Ҫ��������.
        '���˺�:2008/09/03
        If ActivateKeyboardLayout(arrIme(0), 0) <> 0 Then OpenIme = True: Exit Function
    End If
End Function


Public Function SetSQLTrace(ByVal strServerName As String, ByVal strUserName As String, ByRef cnOracle As ADODB.Connection) As String
'����:����100046�¼�����SQL Trace����
'����:Trc�ļ���
    Dim strSql As String, strLevel As String, strFile As String
    Dim rsTmp As ADODB.Recordset
    
    strServerName = UCase(strServerName)
    
    If strServerName Like "SQLTRACE*" Then
        On Error Resume Next
        strSql = "alter session set timed_statistics=true"
        cnOracle.ExeCute strSql
        strSql = "alter session set max_dump_file_size='100M'"
        cnOracle.ExeCute strSql
        If err.Number <> 0 Then err.Clear
        
        '������һ�������8.1.7���Ժ��֧��
        strFile = "ZL_" & strUserName
        strSql = "alter session set tracefile_identifier='" & strFile & "'"
        cnOracle.ExeCute strSql
        If err.Number <> 0 Then strFile = "*.trc": err.Clear
        
        strLevel = "12"
        If Replace(strServerName, "SQLTRACE", "") = "4" Then
            strLevel = "4"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "8" Then
            strLevel = "8"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "12" Then
            strLevel = "12"
        End If
        strSql = "alter session set events '10046 trace name context forever ,level " & strLevel & "'"
        cnOracle.ExeCute strSql
        If err.Number = 0 Then
            SetSQLTrace = strFile
            
            If CheckAndAdjustMustTable("ZLREGINFO", , True) Then    '�ȼ��zlreginfo���Ƿ����
                strSql = "Select 1 From zlreginfo Where ��Ŀ='TRACE�ļ�'"
                Set rsTmp = cnOracle.ExeCute(strSql)
                
                If rsTmp.RecordCount > 0 Then
                    strSql = "Update zlreginfo Set ���� ='TRACE�ļ�' Where ��Ŀ='" & strFile & ".trc'"
                Else
                    strSql = "Insert Into zlreginfo (��Ŀ,����) Values ('TRACE�ļ�','" & strFile & ".trc')"
                End If
                cnOracle.ExeCute strSql
            
                If err.Number <> 0 Then
                    MsgBox err.Description
                End If
            End If
        End If
    End If
End Function

Public Function RunCommand(ByVal strCommand As String, Optional ByRef strErr As String, Optional ByVal blnCiper As Boolean, Optional ByVal lngWait As Long = INFINITE) As String
'���ܣ�ִ�������У�����ȡ���������
'�����߼�:���lngWaitΪ0, �����򲻵ȴ�
    Dim piProc          As PROCESS_INFORMATION '������Ϣ
    Dim stStart         As STARTUPINFO '������Ϣ
    Dim saSecAttr       As SECURITY_ATTRIBUTES '��ȫ����
    Dim lnghReadPipe    As Long '��ȡ�ܵ����
    Dim lnghWritePipe   As Long 'д��ܵ����
    Dim lngBytesRead    As Long '�������ݵ��ֽ���
    Dim strBuffer       As String * 256 '��ȡ�ܵ����ַ���buffer
    Dim lngRet          As Long 'API��������ֵ
    Dim lngRetPro       As Long
    Dim strlpOutputs    As String '���������ս��
    
    DoEvents
    On Error Resume Next
    '���ð�ȫ����
    With saSecAttr
        .nLength = LenB(saSecAttr)
        .bInheritHandle = True
        .lpSecurityDescriptor = 0
    End With
    
    '�����ܵ�
    lngRet = CreatePipe(lnghReadPipe, lnghWritePipe, saSecAttr, 0)
    If lngRet = 0 Then
        strErr = "�޷������ܵ���" & GetLastDllErr()
        Exit Function
    End If
    '���ý�������ǰ����Ϣ
    With stStart
        .Cb = LenB(stStart)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE
        .hStdOutput = lnghWritePipe '��������ܵ�
        .hStdError = lnghWritePipe '���ô���ܵ�
    End With
    '��������
    'Command = "c:\windows\system32\ipconfig.exe /all" 'DOS������ipconfig.exeΪ��
    lngRetPro = CreateProcess(vbNullString, strCommand & vbNullChar, saSecAttr, saSecAttr, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, stStart, piProc)
    If lngRetPro = 0 Then
        strErr = "�޷��������̡�" & GetLastDllErr()
        lngRet = CloseHandle(lnghWritePipe)
        lngRet = CloseHandle(lnghReadPipe)
        Exit Function
    Else
        '��Ϊ����д�����ݣ������ȹر�д��ܵ��������������رմ˹ܵ��������޷���ȡ����
        lngRet = CloseHandle(lnghWritePipe)
        WaitForSingleObject piProc.hProcess, lngWait
        Do
            If lngWait <> 0 Then
                lngRet = ReadFile(lnghReadPipe, strBuffer, 256, lngBytesRead, ByVal 0)
            End If
            If lngRet <> 0 Then
                strlpOutputs = strlpOutputs & Left(strBuffer, lngBytesRead)
            Else
                strlpOutputs = strlpOutputs & Left(strBuffer, lngBytesRead)
            End If
            DoEvents
        Loop While (lngRet <> 0) '��ret=0ʱ˵��ReadFileִ��ʧ�ܣ��Ѿ�û�����ݿɶ���
        '��ȡ������ɣ��رո����
        lngRet = CloseHandle(lngRetPro)
        lngRet = CloseHandle(piProc.hProcess)
        lngRet = CloseHandle(piProc.hThread)
        lngRet = CloseHandle(lnghReadPipe)
    End If
    RunCommand = Replace(strlpOutputs, vbNullChar, "")
End Function

Public Function GetProgFuncs(ByVal strProg As String, Optional ByVal blnInitData As Boolean = False) As String
'���ܣ���ȡһ��ģ���Ӧ�Ĺ���Ȩ�޻��ģ��Ȩ���ַ������г�ʼ������
'������
'      strProg��ģ��Ż�ģ��Ȩ���ַ���
'      blnInitData = True����ʱstrProg��ΪȨ���ַ���������ֵΪ��ʼ�����Ȩ���ַ���
'      blnInitData = False����ʱstrProg��Ϊģ��ţ�����ֵΪ��ģ��ӵ�еĹ���Ȩ��
    Dim arrProg() As String
    Dim arrFunc() As String
    Dim strProgNo As String
    Dim i, j As Long
    
    If blnInitData Then
        '��ʼ��Ȩ���ַ���
    
        arrProg = Split(strProg, ",")
        Set mrsProgFuncs = New ADODB.Recordset
        
        '�����ֶ�
        Call mrsProgFuncs.Fields.Append("ģ���", adVarChar, 6)
        Call mrsProgFuncs.Fields.Append("��������", adVarChar, 30)

        '����¼��
        mrsProgFuncs.Open
        With mrsProgFuncs
            For i = 0 To UBound(arrProg)
                strProgNo = Split(arrProg(i), ":")(0)
                arrFunc = Split(Split(arrProg(i) & ":", ":")(1), "|")
                If UBound(arrFunc) >= 0 Then
                    For j = 0 To UBound(arrFunc)
                        .AddNew
                        .Fields("ģ���").value = strProgNo
                        .Fields("��������").value = arrFunc(j)
                    Next
                Else
                    .AddNew
                    .Fields("ģ���").value = strProgNo
                    .Fields("��������").value = ""
                End If
                GetProgFuncs = GetProgFuncs & "," & strProgNo
            Next
        End With
    Else
        '����ģ��ŷ��ع�������
        If Not mrsProgFuncs Is Nothing Then   'mrsProgFuncsΪnothing˵����û�н��г�ʼ������������ǰ��¼�û�ӵ������Ȩ��
            mrsProgFuncs.Filter = "ģ��� = '" & strProg & "'"
            With mrsProgFuncs
                Do While Not .EOF
                    GetProgFuncs = GetProgFuncs & "|" & !��������
                    .MoveNext
                Loop
            End With
        Else
            GetProgFuncs = ""
        End If
    End If
    
    GetProgFuncs = Mid(GetProgFuncs, 2)
End Function
