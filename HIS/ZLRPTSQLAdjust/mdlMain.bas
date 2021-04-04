Attribute VB_Name = "mdlMain"
Option Explicit
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type
Public Type POINTAPI
        x As Long
        y As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 'ǳ����
Public Const BDR_RAISEDINNER = &H4 'ǳ͹��
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '��͹��
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '���
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame������ʽ
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '��Frame������ʽ

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function Htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public gcnOracle As ADODB.Connection     '�������ݿ�����
Public gcnAcc As New ADODB.Connection


Public gstrProductName As String
Public gstrWebSustainer As String
Public gstrWebURL As String
Public gstrWebEmail As String
Public gstrSysName As String                'ϵͳ����
Public gstrUserName As String               '�û���
Public gstrServer As String                 '��������
Public gstrSQL    As String                 'ͨ�õ�SQL������
Public gstrDBUser As String

Public Sub Main()
    Dim objLogin As Object
    'Ϊʵ��XP�������ʾ����ǰ����ִ�иú���
    
    If App.PrevInstance Then
        MsgBox " �Զ����ѷ����Ѿ������� ", vbOKOnly, "�Զ�����"
        Exit Sub
    End If
    On Error Resume Next
    If objLogin Is Nothing Then
        Set objLogin = CreateObject("ZLLogin.clsLogin")
    End If
    If objLogin Is Nothing Then
        MsgBox "����ZLLogin��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ"
        Exit Sub
    Else
        Set gcnOracle = objLogin.Login(2, CStr(Command()), , True)
        If gcnOracle Is Nothing Then
            Exit Sub
        ElseIf gcnOracle.State <> adStateOpen Then
            Exit Sub
        End If
    End If
    
    gstrServer = objLogin.ServerName
    gstrUserName = objLogin.InputUser
    gstrDBUser = objLogin.DBUser
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "")
    frmMDIMain.Show
End Sub

Public Sub SelAll(objTxt As Control)
'���ܣ����ı���ĵ��ı�ѡ��
    If TypeName(objTxt) = "TextBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If InStr(strInput, "'") > 0 Or InStr(strInput, """") > 0 Then
        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Function Currentdate() As Date
    '-------------------------------------------------------------
    '���ܣ���ȡ�������ϵ�ǰ����
    '������
    '���أ�����Oracle���ڸ�ʽ�����⣬����
    '-------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo ErrHand
    With rsTemp
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    Currentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
    
ErrHand:
    Currentdate = Date
    Err = 0
End Function


'��PictureBoxģ���3Dƽ�水ť
'intStyle=0=ƽ��,-1=����,1=͹��,-2=���,2=��͹��
Public Sub RaisEffect(picBox As PictureBox, Optional IntStyle As Integer, Optional strName As String = "")
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .Cls
        .BorderStyle = 0
        
        If IntStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            
            Select Case IntStyle
                Case 1
                    DrawEdge .hDC, PicRect, CLng(BDR_RAISEDINNER), BF_RECT
                Case 2
                    DrawEdge .hDC, PicRect, CLng(EDGE_RAISED), BF_RECT
                Case -1
                    DrawEdge .hDC, PicRect, CLng(BDR_SUNKENOUTER), BF_RECT
                Case -2
                    DrawEdge .hDC, PicRect, CLng(EDGE_SUNKEN), BF_RECT
            End Select
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub



Public Function GetOwnerName(lngSys As Long, cnLink As ADODB.Connection) As String
    Dim rsTmp As New ADODB.Recordset
    
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open "Select ������ From zlSystems Where ���=" & lngSys, cnLink, adOpenKeyset
    If Not rsTmp.EOF Then GetOwnerName = rsTmp!������
End Function

Public Function GetMaxID(strTable As String, cnLink As ADODB.Connection) As Long
    Dim rsTmp As New ADODB.Recordset
    
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open "Select Nvl(Max(ID),0) as ID From " & strTable, cnLink, adOpenKeyset
    If Not rsTmp.EOF Then GetMaxID = rsTmp!id
End Function


Public Function LoadCustomPicture(strID As String) As StdPicture
'����:����Դ�ļ��е�ָ����Դ���ɴ����ļ�
'����:ID=��Դ��,strExt=Ҫ�����ļ�����չ��(��BMP)
'����:�����ļ���
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, "CUSTOM")
    intFile = FreeFile
    
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(Timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    Set LoadCustomPicture = VB.LoadPicture(strR)
    Kill strR
End Function

Public Function GetOEM(ByVal strAsk As String, Optional ByVal blnCorp As Boolean = True) As String
    '-------------------------------------------------------------
    '���ܣ�����ÿ�����ߵ�ASCII��
    '������
    '���أ�
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    'OEMͼƬ���������� ��һ��ָ��˾�ձ꣬��һ���ǲ�Ʒ��ʶ
    strCode = IIf(blnCorp = True, "OEM_", "PIC_")
    For intBit = 1 To Len(strAsk)
        'ȡÿ���ֵ�ASCII��
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function

Public Function OpenDire(odtvOwner As Form, Optional odtvTitle As String) As String
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo
   szTitle = odtvTitle
   With tBrowseInfo
      .hwndOwner = odtvOwner.hwnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS ' + BIF_DONTGOBELOWDOMAIN
   End With
   lpIDList = SHBrowseForFolder(tBrowseInfo)
   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      OpenDire = sBuffer
   End If
End Function

Public Sub ReCompileProcedure(ByVal cnOwner As ADODB.Connection)
    '�Ա��û��������Ѿ�ʧЧ�Ĺ��̽������±���
    Dim rsTemp As New ADODB.Recordset
    Dim lngTime As Long
    
    For lngTime = 1 To 3
        '���������Σ���Ϊ��Щ�������໥���ã�һ�α��벻�ܽ������
        'Ϊ�˿��ٵõ��б������ö���֮������ù�ϵ
        If rsTemp.State = adStateOpen Then rsTemp.Close
        
        gstrSQL = "select OBJECT_NAME from user_objects where object_type='PROCEDURE' and STATUS='INVALID'"
        rsTemp.Open gstrSQL, cnOwner, adOpenStatic, adLockReadOnly
        
        On Error Resume Next
        If rsTemp.RecordCount = 0 Then
            'û�й���ʧЧ��ֱ���˳�
            Exit Sub
        Else
            Do Until rsTemp.EOF
                '�п��ܳ���
                gstrSQL = "alter procedure " & rsTemp("OBJECT_NAME") & " compile"
                cnOwner.Execute gstrSQL
                rsTemp.MoveNext
            Loop
        End If
    Next
End Sub

Public Function CheckSpaceIsUse(ByVal strType As String, ByVal strName As String, ByVal strOwner As String) As Boolean
'���ܣ�����ռ�������ļ��Ƿ��������û�ʹ��
'������strType    ��ռ� �����ļ�
'      strName          ��ռ�������ļ�������
'      strOwner         �����������û�����������
    Dim rsTemp As New ADODB.Recordset
    
    If strType = "��ռ�" Then
        gstrSQL = "select owner from all_tables where tablespace_name='" & UCase(strName) & "' and owner<>'" & UCase(strOwner) & "' AND ROWNUM<2"
    Else
        gstrSQL = "select O.owner  from V$TABLESPACE T,V$DATAFILE F,all_tables O " & _
                  "Where T.TS# = F.TS# And T.name = O.TABLESPACE_NAME " & _
                  "    and F.name='" & UCase(strName) & "' and O.owner<>'" & UCase(strOwner) & "' AND ROWNUM<2 "
    End If
    
    On Error Resume Next
    rsTemp.Open gstrSQL, gcnOracle, , adLockReadOnly
    
    If rsTemp.RecordCount = 0 Then
        'û�������û�ʹ�ã�����ɾ��
        Exit Function
    End If
    '���û�ʹ��
    CheckSpaceIsUse = True
End Function

Public Function GetVerDouble(ByVal varVer As Variant) As Double
'���ܣ����ݰ汾�ַ������������ֻ��İ汾
'������varVer   �汾�ַ�������9.5.0
    Dim varArray As Variant
    
    varVer = IIf(IsNull(varVer), "", varVer)
    varArray = Split(varVer, ".")
    If UBound(varArray) < 2 Then Exit Function
    
    GetVerDouble = Val(varArray(0)) * 10 ^ 8 + Val(varArray(1)) * 10 ^ 4 + Val(varArray(2))
End Function

Public Function GetVerString(ByVal dblVer As Double) As String
'���ܣ��������ֻ��İ汾�������汾�ַ���
'������dblVer   �汾�ַ�������900050000
    
    GetVerString = dblVer \ 10 ^ 8 & "." & (dblVer Mod 10 ^ 8) \ 10 ^ 4 & "." & dblVer Mod 10 ^ 4
End Function

Private Function JudgeManagerVer() As Double
'���ܣ��жϹ����ߵİ汾
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select ��� from zlSvrTools where ���='0502'"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = True Then
        '��������ģ��汾Ϊ9.0.0
        JudgeManagerVer = 9 * 10 ^ 8
        Exit Function
    End If
    rsTemp.Close
    
    gstrSQL = "SELECT CONSTRAINT_NAME FROM All_Constraints C WHERE C.CONSTRAINT_NAME='ZLOPTIONS_PK' AND C.OWNER='ZLTOOLS' AND C.TABLE_NAME='ZLOPTIONS'"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = True Then
        '���������ZLOPTIONS_PKԼ����˵��û��ִ�еڶ��������ű����汾Ϊ9.1.0
        JudgeManagerVer = 9 * 10 ^ 8 + 1 * 10 ^ 4
        Exit Function
    End If
    rsTemp.Close
    
    gstrSQL = "SELECT CONSTRAINT_NAME FROM All_Constraints C WHERE C.CONSTRAINT_NAME='ZLXLSVERIFY_FK_�����' AND C.OWNER='ZLTOOLS' AND C.TABLE_NAME='ZLXLSVERIFY'"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = False Then
        '�������ZLXLSVERIFY_FK_�����Լ����˵��û��ִ�е����������ű����汾Ϊ9.2.0
        JudgeManagerVer = 9 * 10 ^ 8 + 2 * 10 ^ 4
        Exit Function
    End If
    
    JudgeManagerVer = 9 * 10 ^ 8 + 3 * 10 ^ 4
End Function

Public Function LvwSelectColumns(objSet As Object, ByVal strColumn As String, Optional ByVal blnInit As Boolean = False) As Boolean
'����:���б�ؼ����н�������
'����:
'   objSet��Ҫ���õĶ���,Ŀǰֻ֧��ListView���Ժ��ټ���FlexGrid,DataGrid��
'   strColumn���д�����ʽ��"����,�п�,������ֵ,������;����,�п�,������ֵ,������"    ע����֮�����÷ֺ�
'      ���� "����,2000,0,1;����,800,0,0;����,1440,0,0"
'      ��ListView���ԣ�������Ϊ1��ʾ���в���ɾ����������Ϊ0��ʾ���п���ɾ��
'      ��FlexGridView���ԣ������Ի�Ҫ��ʾ�Ƿ����ڹ̶��У��Ա㲻�ܺ������н���˳�����
'   blnInit��True,����ʾѡ�񴰿ڣ�ֱ�ӳ�ʼ��
    Dim varColumns As Variant, varColumn As Variant
    Dim lngCol As Long

    If blnInit Then
        varColumns = Split(strColumn, ";")
        Select Case TypeName(objSet)
            Case "ListView"
                With objSet.ColumnHeaders
                    .Clear
                    For lngCol = LBound(varColumns) To UBound(varColumns)
                        varColumn = Split(varColumns(lngCol), ",")
                        .Add , "_" & varColumn(0), varColumn(0), varColumn(1), varColumn(2)
                    Next
                End With
            Case "MSHFlexGrid"
            Case "DataGrid"
        End Select
    End If
End Function

Public Sub NextLvwPos(lvwObj As Object, ByVal vIndex As Long)
        
    If lvwObj.ListItems.Count > 0 Then
        vIndex = IIf(lvwObj.ListItems.Count > vIndex, vIndex, lvwObj.ListItems.Count)
        lvwObj.ListItems(vIndex).Selected = True
        lvwObj.ListItems(vIndex).EnsureVisible
    End If
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function
Public Sub OpenFolder(ByVal frmodtvOwner As Form, ByRef strFolderName As String, Optional strTitle As String)
    '----------------------------------------------------------------------------------------------------
    '����:ѡ���ļ���
    '����:frmodtvOwner-ѡ���ļ��еĸ�����
    '     strFolderName-ָ�����ļ���
    '     strTitle-����
    '����:strFolderName-����ѡ����ļ���
    '----------------------------------------------------------------------------------------------------
    
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo
   szTitle = strTitle
   With tBrowseInfo
      .hwndOwner = frmodtvOwner.hwnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS ' + BIF_DONTGOBELOWDOMAIN
   End With
   lpIDList = SHBrowseForFolder(tBrowseInfo)
   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      strFolderName = sBuffer
   End If
End Sub

Public Sub OpenAccessRecordset(rsTemp As ADODB.Recordset, strSQL As String, ByVal strFormCaption As String, _
        Optional CursorType As CursorTypeEnum = adOpenStatic, Optional LockType As LockTypeEnum = adLockReadOnly)
    '���ܣ��򿪼�¼��ͬʱ����SQL���
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open strSQL, gcnAcc, CursorType, LockType
End Sub




Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '���ڳ���ʱ,�Զ��ض�
        strTmp = strCode
    End If
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function
Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '��Ҫ�пո������
        strTmp = strCode
    End If
    'ȡ��������ַ�
    Rpad = Replace(strTmp, Chr(0), strChar)
End Function


Public Function AccDataOpen(ByVal strDatabaseName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ��Access���ݿ�
    '������
    '   strDataBaseName�����ݿ�
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim sConnect As String
    Err = 0
    On Error GoTo ErrHand
    Set gcnAcc = New ADODB.Connection
    With gcnAcc
        If .State = adStateOpen Then .Close
        .Provider = "=Microsoft.Jet.OLEDB.4.0"
        sConnect = "Driver={Microsoft Access Driver (*.mdb)};Dbq="
        .Open sConnect & strDatabaseName, strUserName, strUserPwd
    End With
    AccDataOpen = True
    Exit Function
ErrHand:
    MsgBox "���ݿ��ʧ��", vbInformation, ""
    AccDataOpen = False
    Err = 0
End Function


Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡָ���ִ���ֵ,�ִ��п��԰�������
    '--�����:strInfor-ԭ��
    '         lngStart-ֱʼλ��
    '         lngLen-����
    '--������:
    '--��  ��:�Ӵ�
    '-----------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    
    Err = 0
    On Error GoTo ErrHand:

    Substr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    Substr = Replace(Substr, Chr(0), " ")
'    strTmp = Right(Substr, 1)
'    If zlCommFun.ActualLen(strTmp) = 1 Then
'        If asc(strTmp) < 32 Or asc(strTmp) > 126 Then
'            Substr = Left(Substr, Len(Substr) - 1)
'        End If
'    End If
    Exit Function
ErrHand:
    Substr = ""
End Function



